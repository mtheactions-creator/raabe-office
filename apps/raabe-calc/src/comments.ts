// Kommentare (Excel 365's threaded comments): a card beside the cell with the thread (author, date,
// text, replies), a box to answer ("Antworten", Posten or Strg+Eingabe), and the "…" menu with
// Thread löschen and Thread auflösen. Neuer Kommentar starts a thread ("Unterhaltung beginnen");
// hovering a cell with the purple corner shows its card. Vorheriger / Nächster Kommentar walk the
// threads in reading order. Each change is one undo step; the file keeps the threads (D-065).

import { cellName, type CommentThread, type ThreadComment } from "@raabe/sheet-engine";
import { h, icon, showMenu } from "@raabe/shell";
import type { CalcApp } from "./app";

const now = () => new Date().toISOString().replace(/Z$/, "").replace(/(\.\d{2})\d*$/, "$1");
const guid = () => {
  const b = crypto.getRandomValues(new Uint8Array(16));
  b[6] = (b[6] & 0x0f) | 0x40;
  b[8] = (b[8] & 0x3f) | 0x80;
  const x = [...b].map((v) => v.toString(16).padStart(2, "0").toUpperCase()).join("");
  return `{${x.slice(0, 8)}-${x.slice(8, 12)}-${x.slice(12, 16)}-${x.slice(16, 20)}-${x.slice(20)}}`;
};
const shownDate = (iso: string) => {
  const d = new Date(iso.endsWith("Z") ? iso : iso + "Z");
  return Number.isNaN(d.getTime()) ? "" : d.toLocaleString("de-DE", { day: "2-digit", month: "2-digit", year: "numeric", hour: "2-digit", minute: "2-digit" });
};

export function commentsFeature(app: CalcApp, userName: () => string) {
  let card: HTMLElement | null = null;
  let pinned = false;

  const threadAt = (row: number, col: number) => app.sheet.threads?.find((t) => t.row === row && t.col === col);

  function close(): void {
    card?.remove();
    card = null;
    pinned = false;
    removeEventListener("pointerdown", outside, true);
  }
  const outside = (e: Event) => { if (card && !card.contains(e.target as Node)) close(); };

  function change(label: string, fn: (list: CommentThread[]) => CommentThread[]): void {
    // A protected sheet keeps its Kommentare unless Objekte bearbeiten is allowed (Excel).
    if (app.isProtected() && !app.protectionAllows("objects")) return;
    const sheet = app.sheet;
    app.change(label, () => {
      app.history.whole(sheet);
      const next = fn((sheet.threads ?? []).map((t) => ({ ...t, replies: [...t.replies] })));
      sheet.threads = next.length ? next : undefined;
    });
  }

  const entry = (c: ThreadComment) => h("div", { class: "cm-entry" },
    h("div", { class: "cm-head" }, h("span", { class: "cm-avatar" }, (c.author || "?").slice(0, 2).toUpperCase()), h("b", null, c.author || "Unbekannt"), h("span", { class: "cm-date" }, shownDate(c.date))),
    h("div", { class: "cm-text" }, c.text));

  /** The card for a cell: the thread if there is one, else a new one to type (pinned keeps it open). */
  function open(row: number, col: number, pin: boolean): void {
    close();
    const t = threadAt(row, col);
    if (!t && !pin) return;
    const b = app.grid.rectOnCanvas({ r0: row, c0: col, r1: row, c1: col });
    if (!b) return;
    const author = userName() || "Benutzer";
    const box = h("textarea", { class: "cm-input", rows: 2, placeholder: t ? "Antworten" : "Unterhaltung beginnen", "aria-label": t ? "Antworten" : "Unterhaltung beginnen" }) as HTMLTextAreaElement;
    const post = () => {
      const text = box.value.trim();
      if (!text) return;
      const c: ThreadComment = { id: guid(), author, date: now(), text };
      if (t) change("Antworten", (list) => list.map((x) => (x.id === t.id ? { ...x, replies: [...x.replies, c] } : x)));
      else change("Neuer Kommentar", (list) => [...list, { ...c, row, col, replies: [] }]);
      open(row, col, true);
    };
    box.addEventListener("keydown", (e) => {
      e.stopPropagation();
      if (e.key === "Enter" && e.ctrlKey) { e.preventDefault(); post(); }
      else if (e.key === "Escape") { e.preventDefault(); close(); app.grid.focus(); }
    });
    const more = h("button", { type: "button", class: "cm-more", "aria-label": "Weitere Threadaktionen", onclick: (e: MouseEvent) => {
      if (!t) return;
      showMenu(app.commands, [
        { label: "Thread löschen", run: () => { close(); change("Thread löschen", (list) => list.filter((x) => x.id !== t.id)); } },
        { label: t.done ? "Thread erneut öffnen" : "Thread auflösen", run: () => { change(t.done ? "Thread erneut öffnen" : "Thread auflösen", (list) => list.map((x) => (x.id === t.id ? { ...x, done: t.done ? undefined : true } : x))); open(row, col, true); } },
      ], (e.currentTarget as HTMLElement).getBoundingClientRect());
    } }, "…");
    card = h("div", { class: `cm-card${t?.done ? " done" : ""}`, role: "dialog", "aria-label": `Kommentar ${cellName(row, col)}` },
      t ? h("div", { class: "cm-bar" }, h("span", null, t.done ? "Aufgelöst" : ""), more) : null,
      ...(t ? [entry(t), ...t.replies.map(entry)] : []),
      h("div", { class: "cm-reply" }, box, h("button", { type: "button", class: "cm-post", "aria-label": "Posten", "data-tip": "Posten (Strg+Eingabe)", onclick: post }, icon("send"))));
    card.style.left = `${b.x + b.w + 10}px`;
    card.style.top = `${Math.max(0, b.y - 4)}px`;
    app.grid.el.append(card);
    pinned = pin;
    if (pin) { setTimeout(() => { box.focus(); addEventListener("pointerdown", outside, true); }); }
  }

  /** Hover: the card of a cell with a thread, while the pointer is on it. */
  addEventListener("pointermove", (e) => {
    // The grid is made after this feature: look it up when the pointer moves.
    if (pinned || !app.grid || app.grid.isEditing) return;
    const canvas = app.grid.el.querySelector(".grid-canvas") as HTMLElement | null;
    if (!canvas || e.target !== canvas) { if (card && !card.contains(e.target as Node)) close(); return; }
    const r = canvas.getBoundingClientRect();
    const at = app.grid.hitCell(e.clientX - r.left, e.clientY - r.top);
    const t = threadAt(at.row, at.col);
    if (t) { if (!card) open(at.row, at.col, false); }
    else if (card) close();
  });

  /** Vorheriger / Nächster Kommentar: the threads in reading order from the active cell. */
  function step(dir: 1 | -1): void {
    const list = [...(app.sheet.threads ?? [])].sort((a, b) => a.row - b.row || a.col - b.col);
    if (!list.length) return;
    const a = app.sheet.view.active;
    const after = (t: CommentThread) => t.row > a.row || (t.row === a.row && t.col > a.col);
    const before = (t: CommentThread) => t.row < a.row || (t.row === a.row && t.col < a.col);
    const t = dir > 0 ? list.find(after) ?? list[0] : [...list].reverse().find(before) ?? list[list.length - 1];
    app.grid.selectCell(t.row, t.col);
    open(t.row, t.col, true);
  }

  // ----- Kommentare anzeigen: the pane with every thread of the sheet --------------------------
  let pane: HTMLElement | null = null;
  function renderPane(): void {
    if (!pane) return;
    const list = [...(app.sheet.threads ?? [])].sort((a, b) => a.row - b.row || a.col - b.col);
    const items = list.length ? list.map((t) => h("button", { type: "button", class: `cm-pane-item${t.done ? " done" : ""}`, onclick: () => { app.grid.selectCell(t.row, t.col); open(t.row, t.col, true); } },
      h("div", { class: "cm-pane-ref" }, cellName(t.row, t.col), t.done ? " · Aufgelöst" : ""), entry(t), t.replies.length ? h("div", { class: "cm-pane-more" }, t.replies.length === 1 ? "1 Antwort" : `${t.replies.length} Antworten`) : null))
      : [h("p", { class: "muted" }, "Auf diesem Blatt gibt es noch keine Kommentare.")];
    pane.replaceChildren(
      h("div", { class: "cm-pane-head" }, h("b", null, "Kommentare"),
        h("button", { type: "button", class: "cm-pane-new", onclick: () => { const a = app.sheet.view.active; open(a.row, a.col, true); } }, icon("comment"), "Neu"),
        h("button", { type: "button", "aria-label": "Schließen", onclick: () => togglePane() }, icon("close"))),
      h("div", { class: "cm-pane-list" }, ...items));
  }
  function togglePane(): void {
    if (pane) { pane.remove(); pane = null; document.body.classList.remove("cm-pane-open"); }
    else {
      pane = h("aside", { class: "cm-pane", role: "complementary", "aria-label": "Kommentare" });
      // From the top of the sheet area (below the ribbon) to the status bar.
      pane.style.top = `${document.querySelector(".calc-main")?.getBoundingClientRect().top ?? 150}px`;
      document.body.append(pane);
      document.body.classList.add("cm-pane-open");
      renderPane();
    }
    requestAnimationFrame(() => app.grid.refreshLayout());
  }

  return {
    togglePane,
    paneOpen: () => !!pane,
    refreshPane: renderPane,
    newComment: () => { const a = app.sheet.view.active; open(a.row, a.col, true); },
    remove: () => { const a = app.sheet.view.active; const t = threadAt(a.row, a.col); if (t) { close(); change("Kommentar löschen", (list) => list.filter((x) => x.id !== t.id)); } },
    has: () => { const a = app.sheet.view.active; return !!threadAt(a.row, a.col); },
    any: () => !!app.sheet.threads?.length,
    next: () => step(1),
    prev: () => step(-1),
    close,
  };
}
