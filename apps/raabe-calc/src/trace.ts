// Spur zum Vorgänger / Spur zum Nachfolger / Pfeile entfernen as Excel draws them: blue arrows
// from each referenced cell (a dot at its centre) to the formula, a blue frame around a
// referenced range, a dashed black arrow from a small sheet symbol for references to other
// sheets, red arrows from cells that hold an error. Each further click goes one level deeper;
// any change to the workbook removes the arrows (Excel).

import { CellError, cellKey, type Rect, type Sheet } from "@raabe/sheet-engine";
import type { CalcApp } from "./app";

export interface TraceArrow {
  /** The referenced cell or range on this sheet; null for a reference to another sheet. */
  from: Rect | null;
  to: { row: number; col: number };
  error?: boolean;
}

export function traceFeature(app: CalcApp) {
  const arrows = new Map<Sheet, TraceArrow[]>();
  /** Cells whose precedents (p) or dependents (d) are drawn already. */
  const done = new Map<Sheet, { p: Set<number>; d: Set<number> }>();

  const state = (s: Sheet) => {
    let d = done.get(s);
    if (!d) { d = { p: new Set(), d: new Set() }; done.set(s, d); }
    return d;
  };
  const isError = (sheet: number, r: Rect) => app.calc.value(sheet, r.r0, r.c0) instanceof CellError;

  function add(s: Sheet, a: TraceArrow): void {
    const list = arrows.get(s) ?? [];
    const same = (x: TraceArrow) => x.to.row === a.to.row && x.to.col === a.to.col && JSON.stringify(x.from) === JSON.stringify(a.from);
    if (!list.some(same)) list.push(a);
    arrows.set(s, list);
  }

  /** Spur zum Vorgänger: the active cell's references, or one level further on each click. */
  function precedents(): void {
    const sheet = app.sheet, si = app.wb.sheets.indexOf(sheet), st = state(sheet);
    const { row, col } = sheet.view.active;
    // The next level: formulas among the cells the drawn arrows start from.
    let frontier: { row: number; col: number }[] = [];
    if (!st.p.has(cellKey(row, col))) frontier = [{ row, col }];
    else {
      for (const a of arrows.get(sheet) ?? []) {
        if (!a.from || !st.p.has(cellKey(a.to.row, a.to.col))) continue;
        for (let r = a.from.r0; r <= Math.min(a.from.r1, a.from.r0 + 500); r++)
          for (let c = a.from.c0; c <= Math.min(a.from.c1, a.from.c0 + 100); c++)
            if (sheet.get(r, c)?.f !== undefined && !st.p.has(cellKey(r, c))) frontier.push({ row: r, col: c });
      }
    }
    if (!frontier.length) return;
    for (const at of frontier) {
      st.p.add(cellKey(at.row, at.col));
      for (const d of app.calc.precedents(si, at.row, at.col)) {
        if (d.sheet !== si) add(sheet, { from: null, to: at });
        else add(sheet, { from: d.rect, to: at, error: isError(si, d.rect) });
      }
    }
    app.grid.contentChanged();
  }

  /** Spur zum Nachfolger: arrows from the active cell to the formulas that use it, deeper on each click. */
  function dependents(): void {
    const sheet = app.sheet, si = app.wb.sheets.indexOf(sheet), st = state(sheet);
    const { row, col } = sheet.view.active;
    let frontier: { row: number; col: number }[] = [];
    if (!st.d.has(cellKey(row, col))) frontier = [{ row, col }];
    else {
      for (const a of arrows.get(sheet) ?? []) {
        if (a.from && a.from.r0 === a.from.r1 && a.from.c0 === a.from.c1 && st.d.has(cellKey(a.from.r0, a.from.c0)) && !st.d.has(cellKey(a.to.row, a.to.col)))
          frontier.push(a.to);
      }
    }
    if (!frontier.length) return;
    for (const at of frontier) {
      st.d.add(cellKey(at.row, at.col));
      for (const n of app.calc.dependentsOf(si, at.row, at.col)) {
        if (n.sheet !== si) continue;
        add(sheet, { from: { r0: at.row, c0: at.col, r1: at.row, c1: at.col }, to: { row: n.row, col: n.col }, error: isError(si, { r0: at.row, c0: at.col, r1: at.row, c1: at.col }) });
      }
    }
    app.grid.contentChanged();
  }

  function clear(): void {
    if (!arrows.size) return;
    arrows.clear();
    done.clear();
    app.grid.contentChanged();
  }

  return { precedents, dependents, clear, of: (s: Sheet) => arrows.get(s) ?? [], any: () => arrows.size > 0 };
}
