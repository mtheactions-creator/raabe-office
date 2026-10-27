// Loads every function library and offers single-formula evaluation for the UI.

import "./fn-math";
import "./fn-stats";
import "./fn-logic";
import "./fn-text";
import "./fn-date";
import "./fn-lookup";
import "./fn-array";
import { type Rect } from "../address";
import type { Value, Workbook } from "../model";
import { serialOfDate } from "./dates";
import { Ctx, FUNCS, evalFormula, type Host } from "./eval";
import { ALL_FUNCTIONS_DE, englishName } from "./names";
import { FormulaError, parseFormula } from "./parse";
import { Arr, E, type V } from "./values";

/** A host that reads the values stored in the workbook (no recalculation). */
export class StoredHost implements Host {
  private b = new Map<number, Rect | null>();
  constructor(readonly wb: Workbook, private nowFn: () => number = () => serialOfDate(new Date()), private rnd: () => number = Math.random) {}
  value(sheet: number, row: number, col: number): V {
    return this.wb.sheets[sheet]?.value(row, col) ?? null;
  }
  bounds(sheet: number): Rect | null {
    if (!this.b.has(sheet)) this.b.set(sheet, this.wb.sheets[sheet]?.usedRange() ?? null);
    return this.b.get(sheet)!;
  }
  now(): number {
    return this.nowFn();
  }
  random(): number {
    return this.rnd();
  }
  formula(sheet: number, row: number, col: number): string | undefined {
    return this.wb.sheets[sheet]?.get(row, col)?.f;
  }
  spill(sheet: number, row: number, col: number): Rect | null {
    return this.wb.sheets[sheet]?.get(row, col)?.spill ?? null;
  }
}

/** Rows of an array result (a spilling formula); empty cells show as 0. */
export function arrayRows(a: Arr): Value[][] {
  const out: Value[][] = [];
  for (let r = 0; r < a.rows; r++) out.push(a.d.slice(r * a.cols, (r + 1) * a.cols).map((v) => (v === null ? 0 : v)));
  return out;
}

/**
 * Evaluates one formula (file form, without "=") as if it stood in the given cell, with the
 * values currently stored in the workbook. Returns the value, or the rows of an array result.
 */
export function evaluate(wb: Workbook, sheetIndex: number, row: number, col: number, file: string, host?: Host): Value | Value[][] {
  let node;
  try {
    node = parseFormula(file.startsWith("=") ? file.slice(1) : file);
  } catch (e) {
    if (e instanceof FormulaError) return E.name;
    throw e;
  }
  const r = evalFormula(node, new Ctx(host ?? new StoredHost(wb), sheetIndex, row, col));
  if (r instanceof Arr) return arrayRows(r);
  return r === null ? 0 : r;
}

/** The German functions RaabeCalc computes, with Excel's argument hints, for AutoVervollständigen
 *  (Excel's other names are still translated, but give #NAME?). */
export const FUNCTIONS_DE: { name: string; args: string }[] = ALL_FUNCTIONS_DE.filter((f) => FUNCS.has(englishName(f.name) ?? ""));

/** Whether a function name (English, upper case) is implemented. */
export function implemented(name: string): boolean {
  return FUNCS.has(name.toUpperCase());
}
