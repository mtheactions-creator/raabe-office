// Fehlerüberprüfung (Formeln > Fehlerüberprüfung): the cells with an error indicator, one after
// the other from the active cell on (row by row, as Excel goes), each with what is wrong in our
// own words and Excel's choices: Berechnungsschritte anzeigen (Formelauswertung), Fehler
// ignorieren, In Bearbeitungsleiste bearbeiten, Zurück, Weiter; numbers stored as text offer
// In eine Zahl umwandeln.

import { CellError, cellName, type ErrorCode } from "@raabe/sheet-engine";
import { h, icon, messageBox } from "@raabe/shell";
import type { CalcApp } from "./app";
import { formulaLocal } from "./engine";
import { errorIndicator } from "./grid/render";

/** What an error value means, in our own words. */
const MEANING: Partial<Record<ErrorCode, [string, string]>> = {
  "#DIV/0!": ["Division durch 0", "Die Formel teilt durch null oder durch eine leere Zelle."],
  "#N/A": ["Wert nicht verfügbar", "Ein gesuchter Wert wurde nicht gefunden, oder ein Wert fehlt."],
  "#NAME?": ["Ungültiger Name", "Die Formel enthält einen Namen oder Funktionsnamen, den RaabeCalc nicht kennt (oft ein Tippfehler)."],
  "#NULL!": ["Leere Schnittmenge", "Zwei Bereiche, die sich schneiden sollten, haben keine gemeinsame Zelle."],
  "#NUM!": ["Ungültige Zahl", "Eine Zahl ist zu groß, zu klein oder für die Funktion nicht zulässig."],
  "#REF!": ["Ungültiger Bezug", "Die Formel bezieht sich auf Zellen, die es nicht mehr gibt (zum Beispiel gelöscht)."],
  "#VALUE!": ["Falscher Datentyp", "Ein Wert hat den falschen Typ, zum Beispiel Text, wo eine Zahl erwartet wird."],
  "#SPILL!": ["Überlauf blockiert", "Das Ergebnis braucht mehrere Zellen, aber einige davon sind nicht leer."],
  "#CALC!": ["Berechnungsfehler", "Die Formel kann ihr Ergebnis nicht berechnen, zum Beispiel ein leeres Array."],
};

export function errorCheckFeature(app: CalcApp, steps: () => void) {
  /** Cells with an indicator in reading order. */
  function hits(): { row: number; col: number }[] {
    const out: { row: number; col: number }[] = [];
    for (const [key, cell] of app.sheet.cells) if (errorIndicator(cell)) out.push({ row: Math.floor(key / 16384), col: key % 16384 });
    return out.sort((a, b) => a.row - b.row || a.col - b.col);
  }

  async function open(): Promise<void> {
    const sheet = app.sheet;
    const start = { ...sheet.view.active };
    const list = hits();
    if (!list.length) { await messageBox("RaabeCalc", "Auf diesem Blatt gibt es keine Zellen mit Fehlern."); return; }
    // From the active cell on, then from the top (Excel).
    let i = list.findIndex((p) => p.row > start.row || (p.row === start.row && p.col >= start.col));
    if (i < 0) i = 0;
    const title = h("div", { class: "ec-title" });
    const formula = h("div", { class: "ec-formula" });
    const head = h("div", { class: "ec-head" });
    const text = h("p", { class: "ec-text" });
    const side = h("div", { class: "ec-side" });
    const layer = h("div", { class: "dialog-layer" });
    const close = () => { layer.remove(); removeEventListener("keydown", keys, true); app.grid.focus(); };
    const keys = (e: KeyboardEvent) => { if (e.key === "Escape") { e.preventDefault(); e.stopPropagation(); close(); } };
    const show = () => {
      const p = list[i];
      const cell = sheet.get(p.row, p.col);
      if (!cell) return;
      app.grid.selectCell(p.row, p.col);
      const kind = errorIndicator(cell);
      title.textContent = `Fehler in Zelle ${cellName(p.row, p.col)}`;
      formula.textContent = cell.f !== undefined ? formulaLocal(cell) : String(cell.v ?? "");
      side.replaceChildren();
      const btn = (label: string, fn: () => void) => side.append(h("button", { type: "button", onclick: fn }, label));
      if (kind === "text-number") {
        head.textContent = "Zahl als Text gespeichert";
        text.textContent = "Die Zelle enthält eine Zahl, die als Text gespeichert ist; Formeln rechnen damit nicht wie mit einer Zahl.";
        btn("In eine Zahl umwandeln", () => { app.toNumber(p.row, p.col); list.splice(i, 1); i--; next(); });
      } else {
        const code = cell.v instanceof CellError ? cell.v.code : "#VALUE!";
        const m = MEANING[code] ?? ["Fehler", "Die Formel liefert einen Fehlerwert."];
        head.textContent = `Fehler: ${m[0]}`;
        text.textContent = m[1];
        btn("Berechnungsschritte anzeigen...", () => { close(); steps(); });
      }
      btn("Fehler ignorieren", () => {
        app.change("Fehler ignorieren", () => { app.history.cell(sheet, p.row, p.col); sheet.set(p.row, p.col, { ...cell, ignoreError: true }); });
        list.splice(i, 1);
        if (!list.length) { close(); void messageBox("RaabeCalc", "Die Fehlerüberprüfung für dieses Blatt ist fertig."); return; }
        if (i >= list.length) i = 0;
        show();
      });
      btn("In Bearbeitungsleiste bearbeiten", () => { close(); document.querySelector<HTMLTextAreaElement>(".fx-input")?.focus(); });
    };
    const next = () => {
      if (i + 1 >= list.length) { close(); void messageBox("RaabeCalc", "Die Fehlerüberprüfung für dieses Blatt ist fertig."); return; }
      i++;
      show();
    };
    const back = h("button", { type: "button", onclick: () => { if (i > 0) { i--; show(); } } }, "Zurück");
    const fwd = h("button", { type: "button", class: "primary", onclick: next }, "Weiter");
    const dialog = h("div", { class: "dialog ec-dialog", role: "dialog", "aria-label": "Fehlerüberprüfung" },
      h("div", { class: "dlg-title" }, h("span", { class: "t" }, "Fehlerüberprüfung"), h("button", { "aria-label": "Schließen", onclick: close }, icon("close"))),
      h("div", { class: "dlg-body ec-body" }, h("div", null, title, formula, head, text), side),
      h("div", { class: "dlg-buttons" }, back, fwd));
    layer.append(dialog);
    document.body.append(layer);
    addEventListener("keydown", keys, true);
    show();
    fwd.focus();
  }

  return { open };
}
