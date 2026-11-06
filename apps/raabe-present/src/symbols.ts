// Einfügen > Symbol (D-070): PowerPoint's dialog "Symbol" with Schriftart, Subset and the grid of
// characters; Einfügen types the chosen character at the caret (in a selected shape: at the end of its
// text). Subsets are Unicode blocks with PowerPoint's German names; Zuletzt verwendete Symbole above.

import { h, showDialog } from "@raabe/shell";
import type { PresentApp } from "./app";

const SUBSETS: [string, number, number][] = [
  ["Lateinisch-1, Ergänzung", 0x00a1, 0x00ff], ["Lateinisch erweitert-A", 0x0100, 0x017f], ["Griechisch und Koptisch", 0x0391, 0x03c9],
  ["Allgemeine Interpunktion", 0x2010, 0x2044], ["Währungssymbole", 0x20a0, 0x20bf], ["Buchstabenähnliche Symbole", 0x2100, 0x214f],
  ["Zahlenformen", 0x2150, 0x218b], ["Pfeile", 0x2190, 0x21ff], ["Mathematische Operatoren", 0x2200, 0x22ff], ["Geometrische Formen", 0x25a0, 0x25ff],
  ["Verschiedene Symbole", 0x2600, 0x26ff], ["Dingbats", 0x2700, 0x27bf],
];
const RECENT = "raabe.present.symbols";

function recent(): string[] {
  try { return JSON.parse(localStorage.getItem(RECENT) ?? "[]") as string[]; } catch { return []; }
}

export async function symbolDialog(app: PresentApp): Promise<void> {
  const ed = app.editor;
  const id = ed.text?.id ?? ed.sel[0];
  if (id === undefined) return;
  let chosen = "";
  const grid = h("div", { class: "sym-grid" });
  const code = h("span", { class: "muted" });
  const fill = (from: number, to: number) => {
    grid.replaceChildren();
    for (let c = from; c <= to; c++) {
      const ch = String.fromCodePoint(c);
      const b = h("button", { class: "sym", "data-tip": `U+${c.toString(16).toUpperCase().padStart(4, "0")}`, onclick: () => { chosen = ch; code.textContent = `Zeichencode: ${c.toString(16).toUpperCase().padStart(4, "0")}`; for (const x of Array.from(grid.children)) x.classList.toggle("on", x === b); } }, ch);
      b.ondblclick = () => { chosen = ch; (dialogDone as () => void)(); };
      grid.append(b);
    }
  };
  const sub = h("select", { onchange: () => { const s = SUBSETS[sub.selectedIndex]; fill(s[1], s[2]); } }, ...SUBSETS.map(([n]) => h("option", null, n))) as HTMLSelectElement;
  fill(SUBSETS[0][1], SUBSETS[0][2]);
  const rec = h("div", { class: "sym-recent" }, ...recent().map((ch) => h("button", { class: "sym", onclick: () => { chosen = ch; } }, ch)));
  let dialogDone: (() => void) | null = null;
  const run = showDialog({ title: "Symbol", body: h("div", { class: "sym-dialog" }, h("label", null, "Subset: ", sub), grid, h("div", null, "Zuletzt verwendete Symbole:"), rec, code),
    buttons: [{ label: "Einfügen", value: true, primary: true }, { label: "Abbrechen", value: false }], cancel: false });
  dialogDone = () => (document.querySelector(".dialog .primary") as HTMLButtonElement | null)?.click();
  if (!(await run) || !chosen) return;
  try { localStorage.setItem(RECENT, JSON.stringify([chosen, ...recent().filter((c) => c !== chosen)].slice(0, 20))); } catch { /* not kept */ }
  if (!ed.text) ed.enterText(id);
  ed.typeText(chosen);
}
