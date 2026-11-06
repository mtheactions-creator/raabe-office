// Gliederungsansicht and Notizenseite (D-070). The outline lists each slide's title and the text of its
// body placeholders by level; typing changes the slide (one undo step per word run), Eingabe starts a
// new paragraph, a click on a slide's number shows that slide. The Notizenseite shows the slide above
// its notes on a portrait page, the notes editable as in the notes pane.

import { phOf, spTreeOf, PptxDom } from "@raabe/ooxml";
import type { Deck } from "./deck";
import { renderSlide } from "./render/slide";
import { EMU_PX } from "./render/text";
import * as T from "./textops";

export interface ViewHost {
  readonly deck: Deck;
  current(): number;
  setCurrent(i: number): void;
  image(path: string): string | undefined;
}

const TITLE = new Set(["title", "ctrTitle"]);

export class OutlineView {
  readonly el: HTMLElement;
  constructor(private host: ViewHost) {
    this.el = document.createElement("div");
    this.el.className = "outline";
    this.el.hidden = true;
  }

  /** The text shapes of a slide that the outline shows: the title, then body placeholders. */
  private parts(path: string): { el: Element; title: boolean }[] {
    const tree = spTreeOf(this.host.deck.pptx.doc(path));
    const out: { el: Element; title: boolean }[] = [];
    for (const s of PptxDom.kids(tree)) {
      const ph = phOf(s);
      if (!ph || !PptxDom.kid(s, "txBody")) continue;
      if (TITLE.has(ph.type)) out.unshift({ el: s, title: true });
      else if (["body", "obj", "subTitle"].includes(ph.type)) out.push({ el: s, title: false });
    }
    return out;
  }

  render(): void {
    if (this.el.hidden) return;
    const deck = this.host.deck;
    const rows: HTMLElement[] = [];
    deck.slides.forEach((ref, i) => {
      const slide = document.createElement("div");
      slide.className = `ol-slide${i === this.host.current() ? " current" : ""}`;
      const num = document.createElement("span");
      num.className = "ol-num";
      num.textContent = String(i + 1);
      num.onclick = () => this.host.setCurrent(i);
      slide.append(num, Object.assign(document.createElement("span"), { className: "ol-icon" }));
      const parts = this.parts(ref.path);
      if (!parts.some((p) => p.title)) slide.append(this.para(ref.path, null, 0, "", true));
      for (const part of parts) {
        const body = PptxDom.kid(part.el, "txBody")!;
        T.paras(body).forEach((p, pi) => {
          const lvl = Number(PptxDom.kid(p, "pPr")?.getAttribute("lvl") ?? 0);
          slide.append(this.para(ref.path, part.el, pi, T.paraText(p), part.title, part.title ? 0 : lvl + 1));
        });
      }
      rows.push(slide);
    });
    this.el.replaceChildren(...rows);
  }

  private para(path: string, shape: Element | null, pi: number, text: string, title: boolean, level = 0): HTMLElement {
    const d = document.createElement("div");
    d.className = `ol-para${title ? " ol-title" : ""}`;
    d.style.paddingLeft = `${level * 24}px`;
    d.contentEditable = shape ? "true" : "false";
    d.textContent = text;
    d.spellcheck = false;
    if (!shape) return d;
    d.addEventListener("input", () => {
      const deck = this.host.deck;
      deck.change("Eingeben", [path], () => {
        const body = PptxDom.kid(shape, "txBody")!;
        const p = T.paras(body)[pi];
        if (!p) return;
        T.deleteRange(body, { p: pi, o: 0 }, { p: pi, o: T.paraLength(p) });
        if (d.textContent) T.insertText(body, { p: pi, o: 0 }, d.textContent);
      }, `outline:${path}:${pi}`);
    });
    d.addEventListener("keydown", (e) => {
      if (e.key !== "Enter") return;
      e.preventDefault();
      const sel = getSelection();
      const o = sel && d.contains(sel.anchorNode) ? sel.anchorOffset : (d.textContent ?? "").length;
      const deck = this.host.deck;
      deck.change("Eingeben", [path], () => { const body = PptxDom.kid(shape, "txBody")!; T.splitPara(body, { p: pi, o }); });
      this.render();
      const next = this.el.querySelectorAll(".ol-para")[Array.from(this.el.querySelectorAll(".ol-para")).indexOf(d) + 1] as HTMLElement | undefined;
      next?.focus();
    });
    d.addEventListener("focus", () => {
      const i = this.host.deck.slideIndex(path);
      if (i !== this.host.current()) this.host.setCurrent(i);
    });
    return d;
  }
}

export class NotesPageView {
  readonly el: HTMLElement;
  private notes: HTMLTextAreaElement;
  private slideBox: HTMLElement;
  constructor(private host: ViewHost) {
    this.el = document.createElement("div");
    this.el.className = "notes-page-view";
    this.el.hidden = true;
    const page = document.createElement("div");
    page.className = "notes-page";
    this.slideBox = document.createElement("div");
    this.slideBox.className = "notes-page-slide";
    this.notes = document.createElement("textarea");
    this.notes.className = "notes-page-text";
    this.notes.placeholder = "Text bearbeiten";
    this.notes.addEventListener("input", () => {
      const p = this.host.deck.slides[this.host.current()]?.path;
      if (p) this.host.deck.setNotes(p, this.notes.value);
    });
    page.append(this.slideBox, this.notes);
    this.el.append(page);
  }

  render(): void {
    if (this.el.hidden) return;
    const deck = this.host.deck;
    const ref = deck.slides[this.host.current()];
    if (!ref) return;
    const size = deck.pptx.slideSize;
    // The notes page's slide picture: 5,4864 M EMU wide on a 6,858 M EMU wide page (PowerPoint's notes master).
    const pageW = 560, w = pageW * (5486400 / 6858000), s = w / (size.cx / EMU_PX);
    this.slideBox.style.cssText = `width:${w}px;height:${(size.cy / EMU_PX) * s}px`;
    const r = renderSlide(deck.view(ref.path), size, { editing: false, slideNumber: this.host.current() + 1, image: (p) => this.host.image(p), pptx: deck.pptx });
    r.el.style.transform = `scale(${s})`; r.el.style.transformOrigin = "0 0";
    this.slideBox.replaceChildren(r.el);
    const text = deck.notesText(ref.path);
    if (document.activeElement !== this.notes) this.notes.value = text;
  }
}
