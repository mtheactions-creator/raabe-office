// The app's one door to the sheet engine: what a cell shows, what typed input becomes, and
// recalculation. Keeping it in one place lets the grid stay independent of the engine's
// internals (formulas, number formats; packages/sheet-engine).

import {
  Calc, CellError, editText as valueEditText, errorText, formatValue, numberEditText, parseInput, toFileFormula, toLocalFormula,
  type Cell, type CellStyle, type Value, type Workbook,
} from "@raabe/sheet-engine";

export { readCsv, writeCsv } from "@raabe/sheet-engine";

export interface Shown {
  text: string;
  color?: string;
  /** Horizontal alignment from the value (General: numbers right, text left …). */
  align: "left" | "right" | "center";
  /** "*x" in the format: x repeated to fill the cell. */
  fill?: string;
  /** "#####": a number or date that does not fit. */
  hashes?: boolean;
}

/** What a cell shows (German), for a width in px when given (Standard shortens numbers to fit). */
export function shown(cell: Cell | undefined, style: CellStyle, widthPx?: number, measure?: (s: string) => number): Shown {
  const v = cell?.v ?? null;
  if (v === null) return { text: "", align: "left" };
  const r = formatValue(v, style.numFmt, { widthPx, measure });
  const align = r.align ?? (typeof v === "number" ? "right" : typeof v === "boolean" || v instanceof CellError ? "center" : "left");
  return { text: r.text, color: r.color, align, fill: r.fill, hashes: r.hashes };
}

/** What a cell shows at a width in device pixels, with a measure in device pixels and Excel's padding. */
export function shownAt(cell: Cell | undefined, style: CellStyle, widthPx: number, measure: (s: string) => number, padPx: number): Shown {
  const v = cell?.v ?? null;
  if (v === null) return { text: "", align: "left" };
  const r = formatValue(v, style.numFmt, { widthPx, measure, padPx });
  const align = r.align ?? (typeof v === "number" ? "right" : typeof v === "boolean" || v instanceof CellError ? "center" : "left");
  return { text: r.text, color: r.color, align, fill: r.fill, hashes: r.hashes };
}

/** The cell's content as the formula bar and the cell editor show it (Excel's rules, measured in numfmt/input). */
export function editText(cell: Cell | undefined, style: CellStyle): string {
  if (!cell) return "";
  if (cell.f !== undefined) return "=" + toLocalFormula(cell.f);
  if (cell.v instanceof CellError) return errorText(cell.v);
  return valueEditText(cell.v, style.numFmt);
}

/** What Excel's COM reports as FormulaLocal (numbers plain, no apostrophe): for the behaviour catalogue. */
export function formulaLocal(cell: Cell | undefined): string {
  if (!cell) return "";
  if (cell.f !== undefined) return "=" + toLocalFormula(cell.f);
  const v = cell.v;
  if (v === null) return "";
  if (typeof v === "number") return numberEditText(v);
  if (typeof v === "boolean") return v ? "WAHR" : "FALSCH";
  if (v instanceof CellError) return errorText(v);
  return v;
}

/** What typed input becomes (Excel's rules for numbers, dates, percent, formulas …). */
export function readInput(text: string, style: CellStyle): { v: Value; f?: string; numFmt?: string } {
  if (style.numFmt === "@") return { v: text };
  const r = parseInput(text, { numFmt: style.numFmt });
  // "+2*3" becomes "=2*3": Excel drops the plus that started the formula.
  if (r.formula !== undefined) return { v: null, f: toFileFormula(text.startsWith("+") && r.formula.startsWith("=+") ? "=" + r.formula.slice(2) : r.formula) };
  return { v: r.value, numFmt: r.numFmt };
}

export type CalcLike = Calc;

export function makeCalc(wb: Workbook): Calc {
  return new Calc(wb);
}
