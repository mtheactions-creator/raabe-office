// Datenüberprüfung: which validation covers a cell, the entries a list allows and whether a
// value passes, as Excel 365 decides it (recorded with probe_excel_validation.py):
// - whole, decimal, date and time take numbers only (no text, no WAHR) and compare them as
//   they are (a date with a time is not "equal" to the date); whole needs an integer;
// - textLength counts the value's text as Excel writes it for editing (46096 for a date);
// - a literal list compares case-sensitively and ignores trailing spaces of the entry, a list
//   from cells compares text case-insensitively; numbers compare as numbers (2,50 = 2,5);
// - custom passes when the formula gives WAHR or a number other than 0;
// - Leere Zellen ignorieren: a criterion that refers to an empty cell lets every entry of the
//   right kind pass; without it the empty cell counts as 0.

import { MAX_COLS, MAX_ROWS, cellKey, inRect, type Rect } from "./address";
import { evaluate, StoredHost } from "./formula/functions";
import { adjustFormula, formulaRefs, shiftFormula } from "./formula/refs";
import { toLocalFormula } from "./formula/convert";
import { numberEditText, parseInput } from "./input";
import { CellError, type Sheet, type Validation, type Value, type Workbook } from "./model";

/** The validation that applies to a cell (Excel uses the first one in the file). */
export function validationAt(sheet: Sheet, row: number, col: number): Validation | undefined {
  return sheet.validations?.find((v) => v.ref.some((r) => inRect(r, row, col)));
}

/** A literal list's entries ("\"Ja,Nein\"" or Excel's x12ac items); null when the list comes from cells or a formula. */
export function listItems(v: Validation): string[] | null {
  if (v.type !== "list") return null;
  if (v.items) return [...v.items];
  const f = v.formula1 ?? "";
  if (!/^".*"$/s.test(f)) return null;
  return f.slice(1, -1).replace(/""/g, '"').split(",");
}

/** formula1 (and items) for a literal list, as Excel writes it: items with a comma or quote need x12ac:list. */
export function literalList(items: string[]): { formula1: string; items?: string[] } {
  const formula1 = `"${items.join(",").replace(/"/g, '""')}"`;
  return items.some((i) => /[,"]/.test(i)) ? { formula1, items: [...items] } : { formula1 };
}

/** A criterion moved to a cell of the validation (relative references follow the cell). */
export function criterionAt(v: Validation, formula: string, row: number, col: number): string {
  const at = v.ref[0];
  return at ? shiftFormula(formula, row - at.r0, col - at.c0) : formula;
}

class EntryHost extends StoredHost {
  constructor(wb: Workbook, private at: { sheet: number; row: number; col: number; value: Value }) {
    super(wb);
  }
  override value(sheet: number, row: number, col: number) {
    const a = this.at;
    return sheet === a.sheet && row === a.row && col === a.col ? a.value : super.value(sheet, row, col);
  }
}

/** A criterion's value at a cell, with the cell holding `value`; arrays come back whole. */
function evalAt(wb: Workbook, sheetIndex: number, v: Validation, formula: string, row: number, col: number, value: Value): Value | Value[][] {
  const f = criterionAt(v, formula, row, col);
  return evaluate(wb, sheetIndex, row, col, f, new EntryHost(wb, { sheet: sheetIndex, row, col, value }));
}

/** Whether a criterion refers to an empty cell (other than the cell being checked). */
function refersToBlank(wb: Workbook, sheetIndex: number, v: Validation, formula: string, row: number, col: number): boolean {
  let refs;
  try {
    refs = formulaRefs(toLocalFormula(criterionAt(v, formula, row, col), wb));
  } catch {
    return false;
  }
  for (const r of refs) {
    if (r.ref.r0 !== r.ref.r1 || r.ref.c0 !== r.ref.c1) continue;
    const s = r.sheet !== undefined ? wb.sheetByName(r.sheet) : wb.sheets[sheetIndex];
    if (!s) continue;
    if (s === wb.sheets[sheetIndex] && r.ref.r0 === row && r.ref.c0 === col) continue;
    const c = s.cells.get(cellKey(r.ref.r0, r.ref.c0));
    if (!c || (c.v === null && c.f === undefined)) return true;
  }
  return false;
}

const first = (x: Value | Value[][]): Value => (Array.isArray(x) ? (x[0]?.[0] ?? null) : x);

/** The values a list allows, in order (literal entries as text; cells' values, empty ones left out). */
export function listValues(wb: Workbook, sheetIndex: number, v: Validation, row: number, col: number): Value[] {
  const items = listItems(v);
  if (items) return items;
  if (v.type !== "list" || !v.formula1) return [];
  let r: Value | Value[][];
  try {
    r = evalAt(wb, sheetIndex, v, v.formula1, row, col, wb.sheets[sheetIndex]?.value(row, col) ?? null);
  } catch {
    return [];
  }
  const flat = Array.isArray(r) ? r.flat() : [r];
  return flat.filter((x) => x !== null && x !== "" && !(x instanceof CellError));
}

/** A literal list entry as Excel compares it with a number (German form, "2,5"). */
function itemNumber(item: string): number | null {
  const p = parseInput(item.trim());
  return typeof p.value === "number" && p.formula === undefined ? p.value : null;
}

/** The text whose length Textlänge counts. */
export function lengthText(value: Value): string {
  if (typeof value === "string") return value;
  if (typeof value === "number") return numberEditText(value);
  if (typeof value === "boolean") return value ? "WAHR" : "FALSCH";
  return "";
}

function compare(op: Validation["operator"], x: number, a: number, b: number): boolean {
  switch (op ?? "between") {
    case "between": return x >= a && x <= b;
    case "notBetween": return x < a || x > b;
    case "equal": return x === a;
    case "notEqual": return x !== a;
    case "greaterThan": return x > a;
    case "lessThan": return x < a;
    case "greaterThanOrEqual": return x >= a;
    case "lessThanOrEqual": return x <= a;
  }
  return true;
}

const usesTwo = (op: Validation["operator"]) => (op ?? "between") === "between" || op === "notBetween";

/**
 * Whether `value` may stand in the cell (row, col) of sheet `sheetIndex` under validation `v`.
 * `value` is what the entry becomes: a typed number, text, or a formula's result. Empty
 * entries always pass (Excel does not check deleting).
 */
export function checkValue(wb: Workbook, sheetIndex: number, v: Validation, row: number, col: number, value: Value): boolean {
  if (value === null || v.type === "any") return true;
  try {
    return check(wb, sheetIndex, v, row, col, value);
  } catch {
    return false;
  }
}

function check(wb: Workbook, sheetIndex: number, v: Validation, row: number, col: number, value: Value): boolean {
  const formulas = [v.formula1, usesTwo(v.operator) && v.type !== "list" && v.type !== "custom" ? v.formula2 : undefined].filter((f): f is string => f !== undefined && f !== "");
  const blankRef = () => !!v.allowBlank && formulas.some((f) => refersToBlank(wb, sheetIndex, v, f, row, col));
  if (v.type === "custom") {
    if (!v.formula1) return true;
    if (blankRef()) return true;
    const r = first(evalAt(wb, sheetIndex, v, v.formula1, row, col, value));
    return r === true || (typeof r === "number" && r !== 0);
  }
  if (v.type === "list") {
    const items = listItems(v);
    if (items) {
      return items.some((item) => {
        if (typeof value === "number") return itemNumber(item) === value;
        if (typeof value === "string") return value === item || value.replace(/ +$/, "") === item;
        return false;
      });
    }
    if (blankRef()) return true;
    const low = (s: string) => s.toLocaleLowerCase("de");
    return listValues(wb, sheetIndex, v, row, col).some((x) =>
      typeof x === "number" ? x === value || (typeof value === "string" && itemNumber(value) === x)
        : typeof x === "string" ? typeof value === "string" && low(x) === low(value) : x === value);
  }
  let x: number;
  if (v.type === "textLength") {
    if (value instanceof CellError) return false;
    x = lengthText(value).length;
  } else {
    if (typeof value !== "number") return false;
    if (v.type === "whole" && !Number.isInteger(value)) return false;
    x = value;
  }
  if (blankRef()) return true;
  const crit = (f: string | undefined): number | null => {
    if (f === undefined || f === "") return null;
    const r = first(evalAt(wb, sheetIndex, v, f, row, col, value));
    if (r === null) return 0;
    return typeof r === "number" ? r : null;
  };
  const a = crit(v.formula1);
  const b = usesTwo(v.operator) ? crit(v.formula2) : 0;
  if (a === null || b === null) return false;
  return compare(v.operator, x, a, b);
}

/**
 * Rows or columns inserted (count > 0) or deleted (count < 0) on `sheet`: its validations' ranges
 * move, grow or shrink, and every sheet's criteria follow (references into `sheet` as formulas do).
 * Record `affectedByShift` for undo before calling.
 */
export function shiftValidations(wb: Workbook, sheet: Sheet, kind: "rows" | "cols", at: number, count: number): void {
  const isRows = kind === "rows";
  const max = isRows ? MAX_ROWS : MAX_COLS;
  for (const s of wb.sheets) {
    if (!s.validations?.length) continue;
    s.validations = s.validations.flatMap((v) => {
      let ref = v.ref;
      if (s === sheet) {
        ref = v.ref.flatMap((r) => {
          const a0 = isRows ? r.r0 : r.c0, a1 = isRows ? r.r1 : r.c1;
          let n0 = a0, n1 = a1;
          if (count > 0) {
            if (a0 >= at) { n0 += count; n1 += count; } else if (a1 >= at) n1 += count;
          } else {
            const d0 = at, d1 = at - count - 1;
            if (a0 >= d0 && a1 <= d1) return [];
            n0 = a0 < d0 ? a0 : a0 > d1 ? a0 + count : d0;
            n1 = a1 > d1 ? a1 + count : Math.min(a1, d0 - 1);
          }
          n1 = Math.min(n1, max - 1);
          if (n0 > n1 || n0 >= max) return [];
          return [isRows ? { ...r, r0: n0, r1: n1 } : { ...r, c0: n0, c1: n1 }];
        });
        if (!ref.length) return [];
      }
      const adjust = (f: string | undefined) => (f === undefined ? f : adjustFormula(f, { kind, sheet: sheet.name, at, count, home: s.name }));
      const out: Validation = { ...v, ref };
      if (v.formula1 !== undefined && !(v.type === "list" && v.formula1.startsWith('"'))) out.formula1 = adjust(v.formula1);
      if (v.formula2 !== undefined) out.formula2 = adjust(v.formula2);
      if (out.rawXml && (ref !== v.ref || out.formula1 !== v.formula1 || out.formula2 !== v.formula2)) delete out.rawXml;
      return [out];
    });
    if (!s.validations.length) s.validations = undefined;
  }
}

/** The sheets `shiftValidations` would change: record them for undo first. */
export function affectedByShift(wb: Workbook, sheet: Sheet): Sheet[] {
  return wb.sheets.filter((s) => s === sheet ? !!s.validations?.length : !!s.validations?.some((v) => [v.formula1, v.formula2].some((f) => f?.includes("!"))));
}

/** Every cell range of the sheet that has a validation (Inhalte auswählen > Datenüberprüfung). */
export function validationRanges(sheet: Sheet): Rect[] {
  return (sheet.validations ?? []).flatMap((v) => v.ref.map((r) => ({ ...r })));
}
