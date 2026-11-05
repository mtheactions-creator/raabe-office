// Daten > Was-wäre-wenn-Analyse > Datentabelle (D-066). Excel's Mehrfachoperation: the selected
// range's first row and/or first column hold input values, the formula(s) sit in the other header
// (one-variable tables) or in the top-left corner (two-variable tables). Each result cell is what the
// formula gives with the input cell set to that row's/column's value; the cells show
// {=MEHRFACHOPERATION(Zeileneingabe;Spalteneingabe)} and cannot be changed one by one. The results are
// worked out by setting the input cell, recalculating and putting it back, after every change
// (not under "Automatisch außer bei Datentabellen", then with F9). In the file: <f t="dataTable">
// on the first result cell, as Excel writes it.

import { h, messageBox, showDialog } from "@raabe/shell";
import { cellName, parseRange, type Rect, type Sheet, type Value } from "@raabe/sheet-engine";
import type { CalcApp } from "./app";

export interface DataTable {
  /** The result cells (the selection without its header row and column). */
  ref: Rect;
  /** Werte aus Zeile: the input cell the first row's values go into. */
  rowInput?: { row: number; col: number };
  /** Werte aus Spalte: the input cell the first column's values go into. */
  colInput?: { row: number; col: number };
}

type Pt = { row: number; col: number };
const at = (a: string): Pt | undefined => {
  const r = parseRange(a.replace(/\$/g, ""));
  return r && r.r0 === r.r1 && r.c0 === r.c1 ? { row: r.r0, col: r.c0 } : undefined;
};

/** The sheet's data tables, from the dataTable formulas the file (or RaabeCalc) put on their first cells. */
export function dataTablesOf(sheet: Sheet): DataTable[] {
  const out: DataTable[] = [];
  for (const [, c] of sheet.cells) {
    const f = c.raw?.f;
    if (!f || !/\bt="dataTable"/.test(f)) continue;
    const attr = (n: string) => new RegExp(`\\b${n}="([^"]*)"`).exec(f)?.[1];
    const ref = parseRange(attr("ref") ?? "");
    if (!ref) continue;
    const r1 = attr("r1") ? at(attr("r1")!) : undefined, r2 = attr("r2") ? at(attr("r2")!) : undefined;
    if (attr("dt2D") === "1") out.push({ ref, rowInput: r1, colInput: r2 });
    else if (attr("dtr") === "1") out.push({ ref, rowInput: r1 });
    else out.push({ ref, colInput: r1 });
  }
  return out;
}

/** The formula Excel writes on the first result cell. */
export function dataTableXml(t: DataTable): string {
  const name = (p: Pt) => cellName(p.row, p.col);
  const ref = `${cellName(t.ref.r0, t.ref.c0)}:${cellName(t.ref.r1, t.ref.c1)}`;
  if (t.rowInput && t.colInput) return `<f t="dataTable" ref="${ref}" dt2D="1" dtr="1" r1="${name(t.rowInput)}" r2="${name(t.colInput)}"/>`;
  if (t.rowInput) return `<f t="dataTable" ref="${ref}" dt2D="0" dtr="1" r1="${name(t.rowInput)}"/>`;
  return `<f t="dataTable" ref="${ref}" dt2D="0" dtr="0" r1="${name(t.colInput!)}"/>`;
}

/** What the formula bar shows in each result cell. */
export const dataTableFormula = (t: DataTable) =>
  `MEHRFACHOPERATION(${t.rowInput ? cellName(t.rowInput.row, t.rowInput.col) : ""};${t.colInput ? cellName(t.colInput.row, t.colInput.col) : ""})`;

const inside = (r: Rect, row: number, col: number) => row >= r.r0 && row <= r.r1 && col >= r.c0 && col <= r.c1;

export function whatIfFeature(app: CalcApp) {
  let busy = false;

  /** Works out every data table of the workbook's sheets (after a change, or F9). */
  function refresh(force = false, target?: { wb: import("@raabe/sheet-engine").Workbook; calc: typeof app.calc }): void {
    if (busy) return;
    const wb = target?.wb ?? app.wb;
    const calc = target?.calc ?? app.calc;
    if (wb.calcMode === "autoNoTable" && !force) return;
    busy = true;
    try {
      wb.sheets.forEach((sheet, si) => {
        const tables = dataTablesOf(sheet);
        for (const t of tables) compute(sheet, si, t, calc);
      });
    } finally {
      busy = false;
    }
  }

  function compute(sheet: Sheet, si: number, t: DataTable, calc: typeof app.calc): void {
    const headRow = t.ref.r0 - 1, headCol = t.ref.c0 - 1;
    const inputs = [t.rowInput, t.colInput].filter((p): p is Pt => !!p);
    const saved = inputs.map((p) => sheet.get(p.row, p.col));
    const put = (p: Pt, v: Value) => sheet.set(p.row, p.col, { ...(sheet.get(p.row, p.col) ?? { v: null }), f: undefined, v });
    const results: { row: number; col: number; v: Value }[] = [];
    const two = !!(t.rowInput && t.colInput);
    for (let r = t.ref.r0; r <= t.ref.r1; r++) {
      for (let c = t.ref.c0; c <= t.ref.c1; c++) {
        let formula: Pt;
        if (two) {
          put(t.rowInput!, sheet.value(headRow, c));
          put(t.colInput!, sheet.value(r, headCol));
          formula = { row: headRow, col: headCol };
        } else if (t.colInput) {
          put(t.colInput, sheet.value(r, headCol));
          formula = { row: headRow, col: c };
        } else {
          put(t.rowInput!, sheet.value(headRow, c));
          formula = { row: r, col: headCol };
        }
        calc.changed(si, inputs);
        results.push({ row: r, col: c, v: calc.value(si, formula.row, formula.col) });
      }
    }
    inputs.forEach((p, i) => sheet.set(p.row, p.col, saved[i]));
    calc.changed(si, inputs);
    const shown = dataTableFormula(t);
    const first = sheet.get(t.ref.r0, t.ref.c0);
    const xml = first?.raw?.f ?? dataTableXml(t);
    for (const x of results) {
      const c = sheet.get(x.row, x.col);
      const isFirst = x.row === t.ref.r0 && x.col === t.ref.c0;
      sheet.set(x.row, x.col, { ...(c ?? { v: null }), f: undefined, v: x.v, dt: shown, ...(isFirst ? { raw: { ...c?.raw, f: xml } } : {}) });
    }
    calc.changed(si, results.map(({ row, col }) => ({ row, col })));
  }

  /** Editing inside a data table: Excel refuses it (whole tables can be cleared). */
  function guard(r: Rect): boolean {
    for (const t of dataTablesOf(app.sheet)) {
      const whole = r.r0 <= t.ref.r0 && r.r1 >= t.ref.r1 && r.c0 <= t.ref.c0 && r.c1 >= t.ref.c1;
      const touches = r.r0 <= t.ref.r1 && r.r1 >= t.ref.r0 && r.c0 <= t.ref.c1 && r.c1 >= t.ref.c0;
      if (touches && !whole) { void messageBox("RaabeCalc", "Ein Teil einer Datentabelle lässt sich nicht einzeln ändern."); return false; }
    }
    return true;
  }

  /** Datentabelle: Werte aus Zeile / Werte aus Spalte for the selected range. */
  async function dialog(): Promise<void> {
    const sel = app.grid.selection.at(-1)!;
    if (sel.r1 <= sel.r0 && sel.c1 <= sel.c0) { await messageBox("RaabeCalc", "Markieren Sie zuerst den Bereich der Datentabelle mit den Eingabewerten und der Formel."); return; }
    const row = h("input", { type: "text", style: "width:140px" }) as HTMLInputElement;
    const col = h("input", { type: "text", style: "width:140px" }) as HTMLInputElement;
    const body = h("div", { class: "goalseek" }, h("label", { class: "field" }, "Werte aus Zeile:", row), h("label", { class: "field" }, "Werte aus Spalte:", col));
    setTimeout(() => row.focus());
    const ok = await showDialog({ title: "Datentabelle", body, buttons: [{ label: "OK", value: true, primary: true }, { label: "Abbrechen", value: false }], cancel: false });
    if (!ok) return;
    const ri = row.value.trim() ? at(row.value.trim()) : undefined, ci = col.value.trim() ? at(col.value.trim()) : undefined;
    if ((row.value.trim() && !ri) || (col.value.trim() && !ci) || (!ri && !ci)) { await messageBox("RaabeCalc", "Der Bezug auf die Eingabezelle ist ungültig."); return; }
    for (const p of [ri, ci]) if (p && inside({ r0: sel.r0, c0: sel.c0, r1: sel.r1, c1: sel.c1 }, p.row, p.col)) { await messageBox("RaabeCalc", "Die Eingabezelle darf nicht in der Datentabelle liegen."); return; }
    create(sel, ri, ci);
  }

  /** Makes a data table over the selection `sel` (headers included). */
  function create(sel: Rect, ri?: Pt, ci?: Pt): void {
    const ref = { r0: sel.r0 + 1, c0: sel.c0 + 1, r1: sel.r1, c1: sel.c1 };
    const t: DataTable = { ref, ...(ri ? { rowInput: ri } : {}), ...(ci ? { colInput: ci } : {}) };
    const sheet = app.sheet;
    app.change("Datentabelle", () => {
      app.history.cells(sheet, ref);
      for (let r = ref.r0; r <= ref.r1; r++) for (let c = ref.c0; c <= ref.c1; c++) {
        const x = sheet.get(r, c);
        sheet.set(r, c, { ...(x ?? { v: null }), f: undefined, v: 0, raw: r === ref.r0 && c === ref.c0 ? { ...x?.raw, f: dataTableXml(t) } : x?.raw?.f ? { ...x.raw, f: undefined } : x?.raw });
      }
    });
  }

  return { refresh, guard, dialog, create, tables: () => dataTablesOf(app.sheet) };
}
