// Notizen as Excel shows them: Neue Notiz (Umschalt+F2) opens a yellow box beside the cell with
// the author's name in bold, the text is typed straight into it, and a click outside (or Esc)
// ends it; the red triangle marks the cell and hovering shows the box again.

import { h } from "@raabe/shell";
import type { CalcApp } from "./app";

export function notesFeature(app: CalcApp, userName: () => string) {
  let box: HTMLElement | null = null;

  function close(save: boolean): void {
    if (!box) return;
    const el = box;
    box = null;
    el.remove();
    removeEventListener("pointerdown", outside, true);
    if (save) el.dispatchEvent(new Event("save"));
    app.grid.focus();
  }

  function outside(e: Event): void {
    if (box && !box.contains(e.target as Node)) close(true);
  }

  /** Opens the note of the active cell for editing (a new one with the author's name first). */
  function edit(): void {
    close(true);
    const { row, col } = app.sheet.view.active;
    const sheet = app.sheet;
    const cell = sheet.get(row, col);
    const author = userName() || "Benutzer";
    const start = cell?.note ?? `${author}:\n`;
    const area = h("textarea", { class: "note-edit", spellcheck: "false" }, start) as HTMLTextAreaElement;
    box = h("div", { class: "note-box" }, area);
    const b = app.grid.rectOnCanvas({ r0: row, c0: col, r1: row, c1: col });
    box.style.left = `${(b?.x ?? 0) + (b?.w ?? 0) + 12}px`;
    box.style.top = `${Math.max(0, (b?.y ?? 0) - 6)}px`;
    box.addEventListener("save", () => {
      const text = area.value;
      if (text === (cell?.note ?? "") || (text.trim() === `${author}:` && !cell?.note)) return;
      app.change("Notiz", () => {
        app.history.cell(sheet, row, col);
        const c = sheet.get(row, col);
        sheet.set(row, col, { ...(c ?? { v: null }), note: text.trim() ? text : undefined });
      });
    });
    area.addEventListener("keydown", (e) => {
      e.stopPropagation();
      if (e.key === "Escape") { e.preventDefault(); close(true); }
    });
    app.grid.el.append(box);
    area.focus();
    area.setSelectionRange(area.value.length, area.value.length);
    setTimeout(() => addEventListener("pointerdown", outside, true));
  }

  return { edit, close: () => close(true) };
}
