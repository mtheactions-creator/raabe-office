// What German Excel makes of text typed into a cell: numbers with decimal comma and "."
// thousands, percent, euro amounts, scientific notation, fractions, dates, times, WAHR/FALSCH,
// error values, text, and formulas; and which number format it gives the cell. Recorded in Excel
// (tools/feel-lab/probe_excel_input.py); the tests replay it.

import { BUILTIN_FORMATS, builtinId, formatKind } from "./numfmt";
import { dateToSerial, serialToDate } from "./dates";
import { ERROR_TEXT_DE, err, type ErrorCode, type Value } from "./model";

export interface ParsedInput {
  value: Value;
  /** The cell's new number format (file form) when the input changes it. */
  numFmt?: string;
  /** A formula in the German form, starting with "=", as typed (not yet checked). */
  formula?: string;
}

export interface InputOptions {
  date1904?: boolean;
  /** Excel's "Automatische Prozentwerteingabe" (on by default): in a cell formatted as percent,
   * typing 15 gives 15 %, i.e. 0,15. Excel's FormulaLocal does not do it, typing does. */
  autoPercent?: boolean;
  /** Lotus-style formulas (typing): "+1+2", "-A1" become "=+1+2", "=-A1". */
  lotusFormulas?: boolean;
  /** For dates without a year ("1.3", "5. Mai"): the current year comes from here. */
  today?: Date;
}

// Number formats Excel gives typed input (file form).
const F = {
  thousands: "#,##0",
  thousands2: "#,##0.00",
  sci: "0.00E+00",
  pct: "0%",
  pct2: "0.00%",
  cur: BUILTIN_FORMATS[6],
  cur2: BUILTIN_FORMATS[8],
  frac: "# ?/?",
  frac2: "# ??/??",
  date: "m/d/yyyy",
  dayMonth: "d-mmm",
  monthYear: "mmm-yy",
  dayMonthYear: "d-mmm-yy",
  hm: "h:mm",
  hms: "h:mm:ss",
  hmAmPm: "h:mm AM/PM",
  hmsAmPm: "h:mm:ss AM/PM",
  elapsed: "[h]:mm:ss",
  sub: "mm:ss.0",
  dateTime: "m/d/yyyy h:mm",
};

type Category = "number" | "thousands" | "percent" | "currency" | "sci" | "fraction" | "date" | "time" | "general";

interface Recognised { value: number | string; numFmt?: string; cat: Category }

const ERRORS = new Map<string, ErrorCode>(
  (Object.entries(ERROR_TEXT_DE) as [ErrorCode, string][]).filter(([c]) => c !== "#SPILL!" && c !== "#CALC!" && c !== "#GETTING_DATA").map(([c, t]) => [t.toUpperCase(), c]),
);

/** What Excel stores for `text` typed into a cell whose number format is `current.numFmt`. */
export function parseInput(text: string, current?: { numFmt: string }, opts: InputOptions = {}): ParsedInput {
  const fmt = current?.numFmt ?? "General";
  const kind = formatKind(fmt);
  if (text === "") return { value: null };
  if (text.startsWith("'")) return { value: text.slice(1) };
  if (kind === "text") return { value: text };
  if (text.startsWith("=") && text.length > 1) return { value: null, formula: text.replace(/\s+$/, "") };
  if (/^(wahr|falsch)$/i.test(text)) return { value: text.toUpperCase() === "WAHR" };
  const e = ERRORS.get(text.toUpperCase());
  if (e) return { value: err(e) };
  const numeric = kind === "number" || kind === "percent" || kind === "scientific" || kind === "fraction" || kind === "currency";
  const r = recognise(text, numeric, opts);
  if (!r) {
    if ((opts.lotusFormulas ?? true) && /^[+-]./.test(text) && /\S/.test(text.slice(1))) return { value: null, formula: "=" + text.replace(/\s+$/, "") };
    return { value: text };
  }
  let value = r.value;
  if (typeof value === "number" && kind === "percent" && r.cat === "number" && (opts.autoPercent ?? true) && Math.abs(value) >= 1) {
    value = value / 100;
  }
  const out: ParsedInput = { value };
  if (r.numFmt && overrides(fmt, r.cat)) out.numFmt = r.numFmt;
  return out;
}

/** Whether the format Excel recognised in the input replaces the cell's format: always in a
 * General cell; in a cell with one of the formats Excel assigns itself (percent, scientific,
 * fraction, date, time) when the input is of another kind; never in other formats. */
function overrides(fmt: string, cat: Category): boolean {
  const id = builtinId(fmt);
  if (fmt === "General" || id === 0) return true;
  if (cat === "number" || cat === "thousands" || cat === "general") return false;
  const own = id === undefined ? undefined : ownCategory(id);
  return own !== undefined && own !== cat;
}

function ownCategory(id: number): Category | undefined {
  if (id === 9 || id === 10) return "percent";
  if (id === 11) return "sci";
  if (id === 12 || id === 13) return "fraction";
  if ((id >= 14 && id <= 17) || id === 22) return "date";
  if ((id >= 18 && id <= 21) || (id >= 45 && id <= 47)) return "time";
  return undefined;
}

// ---------------------------------------------------------------------------------------------

function recognise(text: string, numericCell: boolean, opts: InputOptions): Recognised | null {
  return number(text) ?? fraction(text) ?? (numericCell ? plainFraction(text) : null) ?? dateTime(text, opts);
}

/** "1.234,5", "-,5", "(5)", "15 %", "5 €", "1,5E-3" and their combinations. */
function number(text: string): Recognised | null {
  let s = text.trim();
  if (!s) return null;
  let neg = false;
  let paren = false;
  if (s.startsWith("(") && s.endsWith(")")) {
    paren = true;
    s = s.slice(1, -1).trim();
  }
  let percent = false, currency = false, prefixCurrency = false;
  // A sign may come first or after a leading "€".
  const sign = () => {
    const m = /^([+-])\s*/.exec(s);
    if (m) {
      if (paren) return false;
      neg = m[1] === "-";
      s = s.slice(m[0].length);
    }
    return true;
  };
  if (!sign()) return null;
  if (s.startsWith("€")) {
    prefixCurrency = true;
    s = s.slice(1).trim();
    if (!neg && !sign()) return null;
  }
  if (s.startsWith("%")) {
    percent = true;
    s = s.slice(1).trim();
  } else if (s.endsWith("%")) {
    percent = true;
    s = s.slice(0, -1).trim();
  } else if (!prefixCurrency && s.endsWith("€")) {
    currency = true;
    s = s.slice(0, -1).trim();
  }
  const m = /^(\d+(?:\.\d{3,})*|)(?:,(\d*))?(?:[eE]([+-]?\d+))?$/.exec(s);
  if (!m || (m[1] === "" && !m[2])) return null;
  const intPart = m[1].replace(/\./g, "");
  const decPart = m[2] ?? "";
  // Each group after a "." needs at least three digits ("12.2026" is 122026 in Excel too).
  const grouped = m[1].includes(".");
  let value = toNumber(intPart, decPart, m[3]);
  if (m[3] !== undefined) {
    // Out of range ("1E400", "1E-400"): Excel keeps the text but still sets the format.
    if (!isFinite(value) || (value === 0 && /[1-9]/.test(intPart + decPart))) return { value: text, numFmt: F.sci, cat: "sci" };
    if (percent) value /= 100;
    return { value: neg || paren ? -value : value, numFmt: F.sci, cat: "sci" };
  }
  if (percent) value /= 100;
  if (neg || paren) value = -value;
  if (value === 0) value = 0;
  const decimals = decPart.length > 0;
  if (percent) return { value, numFmt: decimals ? F.pct2 : F.pct, cat: "percent" };
  // A leading "€" is accepted but leaves the format alone.
  if (prefixCurrency) return { value, cat: "number" };
  if (currency) return { value, numFmt: decimals ? F.cur2 : F.cur, cat: "currency" };
  if (grouped) return { value, numFmt: decimals ? F.thousands2 : F.thousands, cat: "thousands" };
  return { value, cat: "number" };
}

/** Excel keeps 15 significant digits of what is typed and drops (not rounds) the rest. */
function toNumber(int: string, dec: string, exp?: string): number {
  let digits = (int + dec).replace(/^0+/, "");
  const lead = (int + dec).length - digits.length;
  if (digits.length > 15) digits = digits.slice(0, 15) + "0".repeat(digits.length - 15);
  const all = "0".repeat(lead) + digits;
  const s = all.slice(0, int.length) + "." + all.slice(int.length) + (exp !== undefined ? "e" + exp : "");
  return Number(s === "." ? "0" : s);
}

/** "1 1/2", "-0 3/4": a whole number, one space, a fraction. */
function fraction(text: string): Recognised | null {
  const m = /^\s*([+-]?)(\d+) (\d+)\/(\d+)\s*$/.exec(text);
  if (!m) return null;
  const den = Number(m[4]);
  if (den === 0) return null;
  const v = Number(m[2]) + Number(m[3]) / den;
  return { value: m[1] === "-" ? -v : v, numFmt: den >= 10 ? F.frac2 : F.frac, cat: "fraction" };
}

/** "3/4" in a cell formatted as a number is a fraction, elsewhere a date. */
function plainFraction(text: string): Recognised | null {
  const m = /^\s*([+-]?)(\d+)\/(\d+)\s*$/.exec(text);
  if (!m || Number(m[3]) === 0) return null;
  const v = Number(m[2]) / Number(m[3]);
  return { value: m[1] === "-" ? -v : v, numFmt: Number(m[3]) >= 10 ? F.frac2 : F.frac, cat: "fraction" };
}

// ---- dates and times

const MONTH_NAMES: Record<string, number> = {
  jan: 1, januar: 1, feb: 2, februar: 2, "mär": 3, "märz": 3, mrz: 3, apr: 4, april: 4, mai: 5, jun: 6, juni: 6,
  jul: 7, juli: 7, aug: 8, august: 8, sep: 9, sept: 9, september: 9, okt: 10, oktober: 10, nov: 11, november: 11,
  dez: 12, dezember: 12,
};

interface DatePart { serial: number; numFmt: string }
interface TimePart { frac: number; numFmt: string; overflow: boolean; elapsed: boolean }

function dateTime(text: string, opts: InputOptions): Recognised | null {
  if (/^\s/.test(text)) return null;
  const s = text.replace(/\s+$/, "");
  const d = datePart(s, opts);
  if (d) return { value: d.serial, numFmt: d.numFmt, cat: "date" };
  const t = timePart(s);
  if (t) {
    if (t.overflow) return { value: t.frac, cat: "general" };
    return { value: t.frac, numFmt: t.numFmt, cat: "time" };
  }
  // A date and a time, in either order, separated by spaces.
  const parts = s.split(/\s+/);
  for (let k = 1; k < parts.length; k++) {
    const a = parts.slice(0, k).join(" "), b = parts.slice(k).join(" ");
    for (const [ds, ts] of [[a, b], [b, a]]) {
      const dd = datePart(ds, opts);
      const tt = dd && timePart(ts);
      if (dd && tt) {
        const v = dd.serial + tt.frac;
        if (tt.overflow || tt.elapsed) return { value: v, cat: "general" };
        return { value: v, numFmt: F.dateTime, cat: "date" };
      }
    }
  }
  return null;
}

function year(y: string): number | null {
  if (y.length <= 2) {
    const n = Number(y);
    return n < 30 ? 2000 + n : 1900 + n;
  }
  if (y.length === 4) {
    const n = Number(y);
    return n >= 1900 ? n : null;
  }
  return null;
}

function serial(y: number, m: number, d: number, opts: InputOptions): number | null {
  if (m < 1 || m > 12 || d < 1) return null;
  const leap1900 = !opts.date1904 && y === 1900 && m === 2;
  const dim = leap1900 ? 29 : new Date(Date.UTC(y, m, 0)).getUTCDate();
  if (d > dim) return null;
  const s = dateToSerial(y, m, d, opts.date1904);
  return isNaN(s) ? null : s;
}

function thisYear(opts: InputOptions): number {
  return (opts.today ?? new Date()).getFullYear();
}

function datePart(s: string, opts: InputOptions): DatePart | null {
  let m: RegExpExecArray | null;
  // 1.1.2026, 1. 1. 26
  if ((m = /^(\d{1,2})\s*\.\s*(\d{1,2})\s*\.\s*(\d{1,4})$/.exec(s))) return dmy(m[1], m[2], m[3], opts);
  // 1/1/2026, 1-1-2026, 26-1-1
  if ((m = /^(\d{1,2})([/-])(\d{1,2})\2(\d{1,4})$/.exec(s))) return dmy(m[1], m[3], m[4], opts);
  // 2026-01-01, 2026/1/1
  if ((m = /^(\d{4})([/-])(\d{1,2})\2(\d{1,2})$/.exec(s))) {
    const v = serial(Number(m[1]), Number(m[3]), Number(m[4]), opts);
    return v === null || Number(m[1]) < 1900 ? null : { serial: v, numFmt: F.date };
  }
  // 1.3 / 1/3 / 1-3: day and month of this year; otherwise month and two-digit year (12.34).
  if ((m = /^(\d{1,2})([./-])(\d{1,2})$/.exec(s))) {
    const a = Number(m[1]), b = Number(m[3]);
    const dm = serial(thisYear(opts), b, a, opts);
    if (dm !== null) return { serial: dm, numFmt: F.dayMonth };
    if (m[3].length === 2 && a >= 1 && a <= 12) {
      const v = serial(year(m[3])!, a, 1, opts);
      if (v !== null) return { serial: v, numFmt: F.monthYear };
    }
    return null;
  }
  // Month names: "1. Jan", "1-Jan-26", "1. Januar 2026", "Jan 26", "Januar 2026", "Feb-1".
  if ((m = /^(\d{1,2})(?:\.\s*|\s+|-|\/)([A-Za-zÄÖÜäöü]+)(?:(?:\.\s*|\s+|-|\/|\.)(\d{1,4}))?$/.exec(s))) {
    const mon = MONTH_NAMES[m[2].toLowerCase()];
    if (!mon) return null;
    if (m[3] === undefined) {
      const v = serial(thisYear(opts), mon, Number(m[1]), opts);
      return v === null ? null : { serial: v, numFmt: F.dayMonth };
    }
    const y = year(m[3]);
    const v = y === null ? null : serial(y, mon, Number(m[1]), opts);
    return v === null ? null : { serial: v, numFmt: F.dayMonthYear };
  }
  if ((m = /^([A-Za-zÄÖÜäöü]+)\.?(?:\s+|-)(\d{1,4})$/.exec(s))) {
    const mon = MONTH_NAMES[m[1].toLowerCase()];
    const y = year(m[2]);
    if (!mon || y === null) return null;
    const v = serial(y, mon, 1, opts);
    return v === null ? null : { serial: v, numFmt: F.monthYear };
  }
  return null;
}

function dmy(d: string, mo: string, y: string, opts: InputOptions): DatePart | null {
  const yy = year(y);
  if (yy === null) return null;
  const v = serial(yy, Number(mo), Number(d), opts);
  return v === null ? null : { serial: v, numFmt: F.date };
}

/** "12:30", "12:30:15", "12:30:15,5", "30:15,5" (minutes, seconds), "8:00 PM", "8 PM", "25:00". */
function timePart(s: string): TimePart | null {
  const m = /^(\d{1,5})(?::(\d{0,2}))?(?::(\d{1,2}))?(?:,(\d+))?(?:\s+(am|pm|a|p))?$/i.exec(s);
  if (!m) return null;
  const [, a, b, c, f, ap] = m;
  if (b === undefined && !ap) return null;
  if (b === undefined && f !== undefined) return null;
  const frac = f !== undefined ? Number("0." + f) : 0;
  let h: number, mi: number, sec: number;
  let numFmt: string;
  if (f !== undefined && c === undefined) {
    // "a:b,c" is minutes and seconds.
    if (ap) return null;
    h = 0; mi = Number(a); sec = Number(b || "0") + frac;
    numFmt = F.sub;
  } else {
    h = Number(a); mi = Number(b || "0"); sec = Number(c ?? "0") + frac;
    if (f !== undefined) numFmt = F.sub;
    else if (ap) numFmt = c !== undefined ? F.hmsAmPm : F.hmAmPm;
    else numFmt = c !== undefined ? F.hms : F.hm;
  }
  if (ap) {
    if (h > 12 || h === 0) return null;
    const pm = ap[0].toLowerCase() === "p";
    h = (h % 12) + (pm ? 12 : 0);
  }
  if (h >= 10000) return null;
  const elapsed = h >= 24 && f === undefined;
  if (elapsed) numFmt = F.elapsed;
  const overflow = mi >= 60 || sec >= 60;
  const secs = Math.round((h * 3600 + mi * 60 + sec) * 1000) / 1000;
  return { frac: secs / 86400, numFmt, overflow, elapsed };
}

// ---- the text to edit (Bearbeitungsleiste, F2)

/** A number as Excel writes it for editing: 15 significant digits, decimal comma, written out
 * up to 21 characters and in scientific notation beyond (1E+21, 1E-20). Recorded through
 * FormulaLocal, which shows it the same way whatever the cell's format. */
export function numberEditText(v: number): string {
  if (!isFinite(v)) return String(v);
  if (v === 0) return "0";
  const sign = v < 0 ? "-" : "";
  const s = Math.abs(v).toExponential(14);
  const e = Number(s.slice(s.indexOf("e") + 1));
  const digits = (s[0] + s.slice(2, s.indexOf("e"))).replace(/0+$/, "");
  let plain: string;
  if (e >= 0) {
    const int = digits.length > e + 1 ? digits.slice(0, e + 1) : digits.padEnd(e + 1, "0");
    const frac = digits.slice(e + 1);
    plain = int + (frac ? "," + frac : "");
  } else plain = "0," + "0".repeat(-e - 1) + digits;
  if (plain.length <= 21) return sign + plain;
  const mant = digits[0] + (digits.length > 1 ? "," + digits.slice(1) : "");
  return sign + mant + "E" + (e < 0 ? "-" : "+") + String(Math.abs(e)).padStart(2, "0");
}

/** The text the Bearbeitungsleiste shows for a cell's value, so that typing it back gives the
 * same value: dates as TT.MM.JJJJ (with hh:mm:ss when there is a time), times as hh:mm:ss,
 * percentages with "%", text that would be read as something else with a leading apostrophe.
 * The number part is recorded from Excel; the date, time and percent forms follow Excel's
 * formula bar, which COM cannot read from an invisible Excel (not measured). */
export function editText(value: Value, numFmt = "General", date1904 = false): string {
  if (value === null) return "";
  if (typeof value === "boolean") return value ? "WAHR" : "FALSCH";
  if (typeof value === "string") {
    if (formatKind(numFmt) === "text") return value;
    const back = parseInput(value, { numFmt }, { date1904, lotusFormulas: true });
    return back.formula === undefined && back.value === value ? value : "'" + value;
  }
  if (typeof value !== "number") return ERROR_TEXT_DE[value.code];
  const kind = formatKind(numFmt);
  if (kind === "percent") return numberEditText(Number((value * 100).toPrecision(15))) + "%";
  if ((kind === "date" || kind === "datetime" || kind === "time") && (value >= 1 || (kind === "time" && value >= 0)) && value < (date1904 ? 2957004 : 2958466)) {
    const days = Math.floor(value);
    const secs = Math.round((value - days) * 86400);
    const time = [Math.floor(secs / 3600) % 24, Math.floor(secs / 60) % 60, secs % 60].map((n) => String(n).padStart(2, "0")).join(":");
    const day = days + Math.floor(secs / 86400);
    if (kind === "time" && day === 0) return time;
    const d = serialToDate(day, date1904);
    const date = `${String(d.d).padStart(2, "0")}.${String(d.m).padStart(2, "0")}.${d.y}`;
    return secs % 86400 === 0 && kind !== "time" ? date : `${date} ${time}`;
  }
  return numberEditText(value);
}
