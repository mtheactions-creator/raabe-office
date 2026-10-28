// Formelauswertung (Formeln > Formelauswertung): the active cell's formula worked out one part at
// a time as Excel's Formel auswerten does: the part underlined is the next one, Auswerten puts its
// value in its place (shown in italics), at the end the button becomes Neu starten.

import { CellError, ERROR_TEXT_DE, cellName, formulaNumber, nextStep, toFileFormula, type Value } from "@raabe/sheet-engine";
import { h, showDialog } from "@raabe/shell";
import type { CalcApp } from "./app";
import { formulaLocal } from "./engine";

/** A value as it is written in a German formula. */
function literal(v: Value | Value[][]): string {
  if (Array.isArray(v)) return `{${v.map((row) => row.map((x) => literal(x)).join(".")).join(";")}}`;
  if (v === null) return "0";
  if (typeof v === "number") return formulaNumber(v, "local");
  if (typeof v === "boolean") return v ? "WAHR" : "FALSCH";
  if (v instanceof CellError) return ERROR_TEXT_DE[v.code];
  return `"${String(v).replace(/"/g, '""')}"`;
}

export function evaluateFeature(app: CalcApp) {
  async function open(): Promise<void> {
    const sheet = app.sheet, si = app.wb.sheets.indexOf(sheet);
    const { row, col } = sheet.view.active;
    const cell = sheet.get(row, col);
    const start = cell?.f !== undefined ? formulaLocal(cell) : "";
    let text = start;
    let last: { start: number; end: number } | null = null;
    const ref = `${/^[\p{L}_][\p{L}\p{N}_.]*$/u.test(sheet.name) ? sheet.name : `'${sheet.name}'`}!$${cellName(0, col).replace(/\d+$/, "")}$${row + 1}`;
    const box = h("div", { class: "eval-box", "aria-label": "Auswertung" });
    const button = h("button", { type: "button", class: "primary" }, "Auswerten") as HTMLButtonElement;
    const render = () => {
      const step = nextStep(text);
      // The value just put in (italic) inside or beside the next part (underlined): runs of
      // characters with the same marks.
      const parts: (Node | string)[] = [];
      const flag = (i: number) => (step && i >= step.start && i < step.end ? 1 : 0) | (last && i >= last.start && i < last.end ? 2 : 0);
      let i = 0;
      while (i < text.length) {
        const f = flag(i);
        let j = i + 1;
        while (j < text.length && flag(j) === f) j++;
        let node: Node | string = text.slice(i, j);
        if (f & 2) node = h("i", null, node);
        if (f & 1) {
          const prev = parts.at(-1);
          if (prev instanceof HTMLElement && prev.tagName === "U") { prev.append(node); i = j; continue; }
          node = h("u", null, node);
        }
        parts.push(node);
        i = j;
      }
      box.replaceChildren(...parts);
      button.textContent = step ? "Auswerten" : "Neu starten";
    };
    button.onclick = () => {
      const step = nextStep(text);
      if (!step) { text = start; last = null; render(); return; }
      const part = text.slice(step.start, step.end);
      let value: Value | Value[][];
      try {
        value = app.calc.evaluate(si, row, col, toFileFormula("=" + part, app.wb));
      } catch {
        value = new CellError("#VALUE!");
      }
      const lit = literal(value);
      text = text.slice(0, step.start) + lit + text.slice(step.end);
      last = { start: step.start, end: step.start + lit.length };
      render();
    };
    render();
    const body = h("div", { class: "eval-dialog" },
      h("div", { class: "eval-row" }, h("div", { class: "fc-label" }, "Bezug:"), h("div", { class: "fc-label" }, "Auswertung:")),
      h("div", { class: "eval-row" }, h("div", { class: "eval-ref" }, ref), box),
      h("p", { class: "muted" }, start
        ? "Der unterstrichene Teil wird als Nächstes berechnet; der zuletzt berechnete Wert steht kursiv."
        : "Die aktive Zelle enthält keine Formel."),
      h("div", { class: "eval-buttons" }, button,
        h("button", { type: "button", disabled: true }, "Einzelschritt"), h("button", { type: "button", disabled: true }, "Prozedurschritt")));
    if (!start) button.disabled = true;
    setTimeout(() => button.focus());
    await showDialog({ title: "Formel auswerten", body, buttons: [{ label: "Schließen", value: undefined }], cancel: undefined });
  }
  return { open };
}
