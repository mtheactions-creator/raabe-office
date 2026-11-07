// Kommentare (D-070): Neuer Kommentar, Löschen, Zurück / Weiter, Kommentare anzeigen. A pane at the
// right lists the comments of the current slide with author and date; a comment is anchored at a
// point of the slide (a small marker). Kept in PowerPoint's classic comment parts (p:cmLst per slide,
// p:cmAuthorLst for the presentation), which every PowerPoint version reads; PowerPoint 365's newer
// threaded comments (p188) are not written yet (D-071).

import { TaskPane } from "@raabe/shell";
import { PptxDom } from "@raabe/ooxml";
import type { PresentApp } from "./app";

const { kid, kids, make, parseXml } = PptxDom;
const RT_COMMENTS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments";
const RT_AUTHORS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/commentAuthors";
const CT_COMMENTS = "application/vnd.openxmlformats-officedocument.presentationml.comments+xml";
const CT_AUTHORS = "application/vnd.openxmlformats-officedocument.presentationml.commentAuthors+xml";
const NSP = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"';

export interface Comment { author: string; initials: string; date: string; text: string; x: number; y: number; idx: number; authorId: number }

export class Comments {
  readonly pane: TaskPane;
  constructor(private app: PresentApp, host: HTMLElement, private user: () => string, private userInitials: () => string = () => "") {
    this.pane = new TaskPane(host);
  }

  private get pptx() { return this.app.deck.pptx; }

  private authorsPath(): string | undefined { return this.pptx.related(this.pptx.presPath, RT_AUTHORS); }

  list(slidePath: string): Comment[] {
    const cp = this.pptx.related(slidePath, RT_COMMENTS);
    if (!cp || !this.pptx.has(cp)) return [];
    const ap = this.authorsPath();
    const authors = new Map<number, { name: string; initials: string }>();
    if (ap && this.pptx.has(ap)) for (const a of kids(this.pptx.doc(ap).documentElement, "cmAuthor")) authors.set(Number(a.getAttribute("id")), { name: a.getAttribute("name") ?? "", initials: a.getAttribute("initials") ?? "" });
    return kids(this.pptx.doc(cp).documentElement, "cm").map((c) => {
      const pos = kid(c, "pos");
      const id = Number(c.getAttribute("authorId") ?? 0);
      return { author: authors.get(id)?.name ?? "", initials: authors.get(id)?.initials ?? "", date: c.getAttribute("dt") ?? "", text: kid(c, "text")?.textContent ?? "",
        x: Number(pos?.getAttribute("x") ?? 0), y: Number(pos?.getAttribute("y") ?? 0), idx: Number(c.getAttribute("idx") ?? 0), authorId: id };
    });
  }

  /** Neuer Kommentar: on the current slide, at the selected shape's top right (else the slide's top left). */
  add(text: string): void {
    const slide = this.app.slidePath();
    if (!slide || !text.trim()) return;
    const name = this.user() || "Benutzer";
    const initials = this.userInitials() || name.split(/\s+/).map((w) => w[0] ?? "").join("").slice(0, 3).toUpperCase();
    const v = this.app.editor.selectedViews[0];
    // Positions in the classic format: 1/8 point (PowerPoint's comment units, 576 per inch).
    const x = v ? Math.round(((v.x + v.w) / 12700) * 8) : 80, y = v ? Math.round((v.y / 12700) * 8) : 80;
    let cp = this.pptx.related(slide, RT_COMMENTS);
    const ap = this.authorsPath() ?? "ppt/commentAuthors.xml";
    const planned = cp ?? this.pptx.freeName("ppt/comments/comment#.xml");
    this.app.deck.change("Neuer Kommentar", [slide, planned, ap], () => {
      if (!this.pptx.has(ap)) {
        this.pptx.put(ap, parseXml(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:cmAuthorLst ${NSP}/>`), CT_AUTHORS);
        this.pptx.link(this.pptx.presPath, RT_AUTHORS, ap);
      }
      const adoc = this.pptx.doc(ap);
      let author = kids(adoc.documentElement, "cmAuthor").find((a) => a.getAttribute("name") === name);
      if (!author) {
        const id = Math.max(0, ...kids(adoc.documentElement, "cmAuthor").map((a) => Number(a.getAttribute("id")))) + 1;
        author = make(adoc, "p:cmAuthor", { id, name, initials, lastIdx: 0, clrIdx: id - 1 });
        adoc.documentElement.append(author);
      }
      const idx = Number(author.getAttribute("lastIdx") ?? 0) + 1;
      author.setAttribute("lastIdx", String(idx));
      if (!cp) {
        cp = planned;
        this.pptx.put(cp, parseXml(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:cmLst ${NSP}/>`), CT_COMMENTS);
        this.pptx.link(slide, RT_COMMENTS, cp);
      }
      const cdoc = this.pptx.doc(cp);
      const cm = make(cdoc, "p:cm", { authorId: author.getAttribute("id")!, dt: new Date().toISOString().replace("Z", ""), idx });
      cm.append(make(cdoc, "p:pos", { x, y }), make(cdoc, "p:text", {}, text));
      cdoc.documentElement.append(cm);
    });
    this.refresh();
  }

  remove(i: number): void {
    const slide = this.app.slidePath();
    const cp = slide ? this.pptx.related(slide, RT_COMMENTS) : undefined;
    if (!slide || !cp) return;
    this.app.deck.change("Kommentar löschen", [cp], () => { kids(this.pptx.doc(cp).documentElement, "cm")[i]?.remove(); });
    this.refresh();
  }

  toggle(): void {
    if (this.pane.isOpen) { this.pane.close(); return; }
    this.pane.show({ title: "Kommentare", sets: [{ label: "", tabs: [{ id: "c", label: "Kommentare", icon: "comment", sections: () => [{ title: "Kommentare", open: true, controls: () => [{ type: "custom", render: () => this.render() }] }] }] }] });
  }

  open(): void { if (!this.pane.isOpen) this.toggle(); }

  refresh(): void { if (this.pane.isOpen) this.pane.refresh(); this.app.editor.render(); }

  private render(): HTMLElement {
    const box = document.createElement("div");
    box.className = "comments-pane";
    const slide = this.app.slidePath();
    const list = slide ? this.list(slide) : [];
    const input = document.createElement("textarea");
    input.className = "comment-input";
    input.placeholder = "Kommentar eingeben";
    const post = document.createElement("button");
    post.textContent = "Posten";
    post.onclick = () => { this.add(input.value); input.value = ""; };
    input.onkeydown = (e) => { if (e.key === "Enter" && e.ctrlKey) { e.preventDefault(); post.click(); } e.stopPropagation(); };
    list.forEach((c, i) => {
      const card = document.createElement("div");
      card.className = "comment-card";
      const head = document.createElement("div");
      head.className = "comment-head";
      head.textContent = `${c.author}  ·  ${c.date ? new Date(c.date).toLocaleString("de-DE", { dateStyle: "short", timeStyle: "short" }) : ""}`;
      const del = document.createElement("button");
      del.textContent = "✕";
      del.setAttribute("aria-label", "Kommentar löschen");
      del.onclick = () => this.remove(i);
      head.append(del);
      const t = document.createElement("div");
      t.className = "comment-text";
      t.textContent = c.text;
      card.append(head, t);
      box.append(card);
    });
    if (!list.length) box.append(Object.assign(document.createElement("div"), { className: "muted", textContent: "Keine Kommentare auf dieser Folie." }));
    box.append(input, post);
    return box;
  }

  /** Markers of the current slide's comments (drawn over the slide in Normalansicht). */
  markers(slidePath: string): { x: number; y: number; n: number }[] {
    return this.list(slidePath).map((c, i) => ({ x: (c.x / 8) * 12700, y: (c.y / 8) * 12700, n: i + 1 }));
  }
}
