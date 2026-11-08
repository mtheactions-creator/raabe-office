// Spezialfilter (Daten > Sortieren und Filtern > Erweitert...; D-090): which records of a list meet the criteria of a
// criteria range, as Excel decides it (recorded: tools/feel-lab/probe_excel_advfilter.py). The criteria range's first
// row names list columns; each further row is one alternative (OR), its filled cells all have to hold (AND). A text
// without an operator matches every entry that begins with it ("Mathe" finds "Mathematik"; ="=Mathe" only "Mathe"),
// operators and wildcards are those of ZÄHLENWENN. A column whose heading is not a list column (or is empty) holds a
// formula criterion, written for the first record and worked out for each one.

import type { Rect } from "./address";
import type { Sheet, Value, Workbook } from "./model";
import { CellError } from "./model";
import { matcher, type Matcher } from "./formula/criteria";
import { shiftFormula } from "./formula/refs";

/** Works out a formula (file form) as if it stood at a cell: what the app's recalculation gives. */
export type FormulaEval = (formula: string, row: number, col: number) => Value | Value[][];

type Test = { col: number; m: Matcher } | { formula: string; at: { row: number; col: number } };

const truthy = (v: Value | Value[][]): boolean => {
  const x = Array.isArray(v) ? v[0]?.[0] ?? null : v;
  if (x instanceof CellError) return false;
  if (typeof x === "boolean") return x;
  if (typeof x === "number") return x !== 0;
  return false;
};

/** The rows (absolute) of the list's records that meet the criteria; with `unique`, a record's repeats are left out. */
export function advancedFilter(wb: Workbook, sheet: Sheet, list: Rect, criteria: Rect, unique: boolean, evaluate?: FormulaEval): number[] {
  const heads: string[] = [];
  for (let c = list.c0; c <= list.c1; c++) heads.push(String(sheet.value(list.r0, c) ?? "").trim().toLowerCase());
  const alternatives: Test[][] = [];
  for (let r = criteria.r0 + 1; r <= criteria.r1; r++) {
    const tests: Test[] = [];
    for (let c = criteria.c0; c <= criteria.c1; c++) {
      const cell = sheet.get(r, c);
      const head = String(sheet.value(criteria.r0, c) ?? "").trim().toLowerCase();
      const col = head ? heads.indexOf(head) : -1;
      if (col < 0) {
        // A formula criterion (its heading is no list column).
        if (cell?.f !== undefined) tests.push({ formula: cell.f, at: { row: r, col: c } });
        continue;
      }
      const v = sheet.value(r, c);
      if (v === null || v === undefined) continue;
      let crit: Value = v;
      // A plain text (no operator, no leading "=") finds every entry that begins with it.
      if (typeof crit === "string" && !/^(<=|>=|<>|<|>|=)/.test(crit)) crit = crit + "*";
      tests.push({ col: list.c0 + col, m: matcher(crit as Exclude<Value, undefined>) });
    }
    alternatives.push(tests);
  }
  const out: number[] = [];
  const seen = new Set<string>();
  for (let row = list.r0 + 1; row <= list.r1; row++) {
    // No criteria rows: every record (Keine Duplikate alone makes a list of the different records).
    const ok = alternatives.length === 0 || alternatives.some((tests) => tests.every((t) => {
      if ("m" in t) return t.m(sheet.value(row, t.col) ?? null);
      if (!evaluate) return false;
      const f = shiftFormula(t.formula, row - (list.r0 + 1), 0);
      return truthy(evaluate(f, t.at.row, t.at.col));
    }));
    if (!ok) continue;
    if (unique) {
      const key: unknown[] = [];
      for (let c = list.c0; c <= list.c1; c++) key.push(sheet.value(row, c) ?? null);
      const k = JSON.stringify(key);
      if (seen.has(k)) continue;
      seen.add(k);
    }
    out.push(row);
  }
  void wb;
  return out;
}
