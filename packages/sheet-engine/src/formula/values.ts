// Values while a formula is evaluated, and Excel's rules for turning one kind into another
// (German locale: "," decimal point, WAHR/FALSCH, dates d.m.yyyy).

import type { Rect } from "../address";
import { CellError, err, type ErrorCode } from "../model";
import { MONTHS_DE, MONTHS_DE_SHORT, MAX_SERIAL, serial, ymd } from "./dates";
import { parseInput } from "../input";

/** A cell value inside the engine; null is an empty cell (or a missing argument). */
export type V = number | string | boolean | CellError | null;

/** A reference: one or more areas on one sheet (several after the union operator). */
export class Ref {
  /** `sheetEnd`: the last sheet of a 3D reference (Tabelle1:Tabelle3!A1), which only the
   *  aggregate functions (SUMME, ANZAHL, MITTELWERT …) accept. */
  constructor(readonly sheet: number, readonly areas: Rect[], readonly sheetEnd?: number) {}
  get single(): boolean {
    return this.areas.length === 1 && this.areas[0].r0 === this.areas[0].r1 && this.areas[0].c0 === this.areas[0].c1;
  }
}

/** A two-dimensional array of values (row-major). */
export class Arr {
  constructor(readonly rows: number, readonly cols: number, readonly d: V[]) {}
  get(r: number, c: number): V {
    return this.d[r * this.cols + c];
  }
  static of(rows: V[][]): Arr {
    return new Arr(rows.length, rows[0]?.length ?? 0, rows.flat());
  }
}

export type Any = V | Ref | Arr;

export const isErr = (v: unknown): v is CellError => v instanceof CellError;

export const E = {
  value: err("#VALUE!"), div0: err("#DIV/0!"), na: err("#N/A"), num: err("#NUM!"), ref: err("#REF!"), name: err("#NAME?"),
  null: err("#NULL!"), spill: err("#SPILL!"), calc: err("#CALC!"),
};

export function errOf(code: ErrorCode): CellError {
  return err(code);
}

/** A number result, with Excel's #ZAHL! for overflow and NaN. */
export function num(x: number): number | CellError {
  return Number.isFinite(x) ? (x === 0 ? 0 : x) : E.num;
}

// ----- Text to number --------------------------------------------------------------------

/**
 * A text as Excel reads it when a number is needed: as typed input is read (input.ts,
 * measured against German Excel: "1,5", "1.000,5", "12%", "4 €", "1.1.2024", "12:30" …),
 * after trimming spaces (=" 4 "+1 is 5); null when it is no number.
 */
export function textToNumber(text: string): number | null {
  const t = text.trim();
  if (!t) return null;
  const v = parseInput(t).value;
  return typeof v === "number" ? v : null;
}

const MONTH_RE = new RegExp(`^(${MONTHS_DE.join("|")}|${MONTHS_DE_SHORT.join("|")})\\.?$`, "i");

function monthOf(word: string): number {
  const m = MONTH_RE.exec(word);
  if (!m) return 0;
  const w = m[1].toLowerCase();
  const i = MONTHS_DE.indexOf(w);
  return (i >= 0 ? i : MONTHS_DE_SHORT.indexOf(w)) + 1;
}

function year2(y: string): number {
  const n = Number(y);
  return y.length <= 2 ? (n < 30 ? 2000 + n : 1900 + n) : n;
}

function currentYear(now?: () => number): number {
  return now ? ymd(now()).y : new Date().getFullYear();
}

/** Serial of a German date text ("15.3.2024", "15. März 2024", "2024-03-15"), or null. */
export function dateValue(t: string, now?: () => number): number | null {
  let y: number, m: number, d: number;
  let r: RegExpExecArray | null;
  const s = t.trim();
  if ((r = /^(\d{1,2})[./-](\d{1,2})[./-](\d{1,4})$/.exec(s))) {
    d = +r[1]; m = +r[2]; y = year2(r[3]);
  } else if ((r = /^(\d{1,2})[./](\d{1,2})\.?$/.exec(s))) {
    d = +r[1]; m = +r[2]; y = currentYear(now);
  } else if ((r = /^(\d{4})[-/](\d{1,2})[-/](\d{1,2})$/.exec(s))) {
    y = +r[1]; m = +r[2]; d = +r[3];
  } else if ((r = /^(\d{1,2})[.\- ]?\s*([\p{L}]+\.?)[\-. ]\s*(\d{1,4})$/u.exec(s)) && monthOf(r[2])) {
    d = +r[1]; m = monthOf(r[2]); y = year2(r[3]);
  } else if ((r = /^(\d{1,2})[.\- ]\s*([\p{L}]+\.?)$/u.exec(s)) && monthOf(r[2])) {
    d = +r[1]; m = monthOf(r[2]); y = currentYear(now);
  } else if ((r = /^([\p{L}]+\.?)[\- ]\s*(\d{4}|\d{2})$/u.exec(s)) && monthOf(r[1])) {
    d = 1; m = monthOf(r[1]); y = year2(r[2]);
  } else return null;
  if (m < 1 || m > 12 || d < 1 || y < 1900 || y > 9999) return null;
  if (d > (m === 2 ? (y % 4 === 0 && (y % 100 !== 0 || y % 400 === 0 || y === 1900) ? 29 : 28) : [31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31][m - 1])) return null;
  const v = serial(y, m, d);
  return v > MAX_SERIAL ? null : v;
}

/** Fraction of a day for "12:30", "6:30:15", "12:30:15,5", "1:00 PM"; null when it is no time. */
export function timeValue(t: string): number | null {
  const r = /^(\d{1,5}):(\d{1,5})(?::(\d{1,5})(?:,(\d+))?)?\s*(AM|PM|A|P)?$/i.exec(t.trim());
  if (!r) return null;
  let h = +r[1];
  const mi = +r[2], s = r[3] ? +r[3] : 0;
  const frac = r[4] ? Number("0." + r[4]) : 0;
  if (r[5]) {
    if (h > 12) return null;
    const pm = r[5].toUpperCase().startsWith("P");
    if (h === 12) h = 0;
    if (pm) h += 12;
  }
  if (r[3] !== undefined && (mi > 59 || s > 59)) return null;
  return (h * 3600 + mi * 60 + s + frac) / 86400;
}

// ----- Number to text --------------------------------------------------------------------

/**
 * A number as Excel turns it into text (concatenation, LÄNGE, TEXTKETTE): the General format
 * with up to 15 significant digits (not limited by a column width), "," as decimal point.
 */
export function numberText(v: number): string {
  if (v === 0) return "0";
  const neg = v < 0;
  const a = Math.abs(v);
  const p = a.toPrecision(15);
  let mant: string, exp: number;
  const m = /^(\d)(?:\.(\d+))?e([+-]\d+)$/.exec(p);
  if (m) {
    mant = m[1] + (m[2] ?? "");
    exp = Number(m[3]);
  } else {
    const e = Number(a.toExponential(14).split("e")[1]);
    const digits = a.toExponential(14).split("e")[0].replace(".", "");
    mant = digits;
    exp = e;
  }
  mant = mant.replace(/0+$/, "") || "0";
  let out: string;
  // Written out up to 20 digits before and 18 after the decimal point, else in E notation
  // (recorded: 10^19 in full, 10^20 → 1E+20, 10^-18 in full, 10^-19 → 1E-19).
  if (exp >= 20 || (exp < 0 && -exp - 1 + mant.length > 18)) {
    const md = mant.length > 1 ? mant[0] + "," + mant.slice(1) : mant;
    out = md + "E" + (exp >= 0 ? "+" : "-") + String(Math.abs(exp)).padStart(2, "0");
  } else if (exp >= 0) {
    out = mant.length <= exp + 1 ? mant + "0".repeat(exp + 1 - mant.length) : mant.slice(0, exp + 1) + "," + mant.slice(exp + 1);
  } else {
    out = "0," + "0".repeat(-exp - 1) + mant;
  }
  return (neg ? "-" : "") + out;
}

// ----- Coercion ----------------------------------------------------------------------------

/** Value to number as operators do: empty 0, WAHR 1, numeric text read, other text #WERT!. */
export function toNum(v: V): number | CellError {
  if (typeof v === "number") return v;
  if (v === null) return 0;
  if (typeof v === "boolean") return v ? 1 : 0;
  if (typeof v === "string") {
    const n = textToNumber(v);
    return n === null ? E.value : n;
  }
  return v;
}

/** Value to text: numbers in the General format, WAHR/FALSCH, empty "". */
export function toStr(v: V): string | CellError {
  if (typeof v === "string") return v;
  if (v === null) return "";
  if (typeof v === "number") return numberText(v);
  if (typeof v === "boolean") return v ? "WAHR" : "FALSCH";
  return v;
}

/** Value to a truth value: numbers ≠ 0, "WAHR"/"FALSCH" (any case), empty FALSCH. */
export function toBool(v: V): boolean | CellError {
  if (typeof v === "boolean") return v;
  if (typeof v === "number") return v !== 0;
  if (v === null) return false;
  if (typeof v === "string") {
    const u = v.toUpperCase();
    if (u === "WAHR") return true;
    if (u === "FALSCH") return false;
    return E.value;
  }
  return v;
}

// ----- Comparison --------------------------------------------------------------------------

const COLLATOR = new Intl.Collator("de", { sensitivity: "accent" });

/** Text comparison as Excel's: case-insensitive, by the German collation. */
export function compareText(a: string, b: string): number {
  if (a === b) return 0;
  // Windows' word sort, which Excel uses, passes over hyphens and apostrophes first and
  // expands ß and Æ ("Straße" = "STRASSE"); then the text with more of them comes later.
  const c = COLLATOR.compare(fold(a), fold(b));
  if (c !== 0) return c;
  const h = (a.match(/['-]/g)?.length ?? 0) - (b.match(/['-]/g)?.length ?? 0);
  return Math.sign(h);
}

function fold(s: string): string {
  return s.replace(/['-]/g, "").replace(/ß/g, "ss").replace(/æ/g, "ae").replace(/Æ/g, "AE").replace(/œ/g, "oe").replace(/Œ/g, "OE");
}

function typeRank(v: number | string | boolean): number {
  return typeof v === "number" ? 0 : typeof v === "string" ? 1 : 2;
}

/**
 * Excel's comparison for the = < > operators: numbers before text before logical values;
 * an empty cell counts as 0, "" or FALSCH, whichever the other side is. Errors are the
 * caller's to handle.
 */
export function compare(a: V, b: V): number {
  if (a === null) a = typeof b === "string" ? "" : typeof b === "boolean" ? false : 0;
  if (b === null) b = typeof a === "string" ? "" : typeof a === "boolean" ? false : 0;
  const x = a as number | string | boolean, y = b as number | string | boolean;
  const ra = typeRank(x), rb = typeRank(y);
  if (ra !== rb) return ra < rb ? -1 : 1;
  if (typeof x === "number") return numCompare(x, y as number);
  if (typeof x === "string") return compareText(x, y as string);
  return x === y ? 0 : x ? 1 : -1;
}

/** Numbers equal to 15 significant digits count as equal (=0,1+0,2=0,3 is WAHR in Excel). */
export function numCompare(x: number, y: number): number {
  if (x === y) return 0;
  if (x !== 0 && y !== 0 && Number(x.toPrecision(15)) === Number(y.toPrecision(15))) return 0;
  return x < y ? -1 : 1;
}
