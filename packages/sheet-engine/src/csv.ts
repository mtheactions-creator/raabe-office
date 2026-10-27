// CSV as German Excel reads and writes it (recorded with tools/feel-lab/probe_excel_csv.py,
// test/fixtures/excel-csv.json):
// - fields are separated by ";" (the list separator; a file with "," stays one column), a
//   first line "sep=x" chooses another separator; quotes only at the start of a field, ""
//   inside quotes, line breaks inside quotes stay in the cell;
// - each field is read like typed input: decimal comma, "." for thousands, percent, euro
//   amounts, German dates and times, WAHR/FALSCH, German error names, formulas;
// - saving writes the used range (from its first row and column), every row with the same
//   number of fields, each value as the cell shows it (General with up to 11 characters),
//   quotes around fields with ";", '"' or line breaks, CR LF after every row. "CSV UTF-8"
//   starts with a byte order mark, "CSV (Trennzeichen-getrennt)" is Windows-1252 without one.

import { CellError, Sheet, Workbook, err, type Cell, type ErrorCode, type StyleTable, type Value } from "./model";
import { BUILTIN_FORMATS, formatValue } from "./numfmt";
import { dateToSerial, timeToSerial } from "./dates";
import { toFileFormula } from "./formula/api";
import { isValidLocal } from "./formula/parse";

/** The byte order mark Excel puts before "CSV UTF-8" files. */
export const CSV_BOM = String.fromCharCode(0xfeff);

/** A field read like typed input: its value, number format and formula (file form). */
export interface CsvValue { v: Value; numFmt?: string; f?: string }

export interface CsvReadOptions {
  /** Separator; default ";" (or the file's "sep=" line). */
  sep?: string;
  /** Name of the sheet (Excel uses the file name without ".csv"). */
  name?: string;
  /** Reads a field instead of the built-in rules (e.g. the app's typed-input parser). */
  parse?: (text: string) => CsvValue;
  date1904?: boolean;
  /** Today, for dates without a year ("1.5" is 1 May of this year). */
  today?: Date;
}

/** The text of CSV bytes: UTF-8 with a byte order mark or valid UTF-8, else Windows-1252. */
export function decodeCsv(bytes: Uint8Array): string {
  if (bytes[0] === 0xef && bytes[1] === 0xbb && bytes[2] === 0xbf) return new TextDecoder("utf-8").decode(bytes.subarray(3));
  try {
    return new TextDecoder("utf-8", { fatal: true }).decode(bytes);
  } catch {
    return new TextDecoder("windows-1252").decode(bytes);
  }
}

/** The rows and fields of CSV text. */
export function splitCsv(text: string, sep = ";"): string[][] {
  const rows: string[][] = [];
  let row: string[] = [];
  let i = 0;
  const n = text.length;
  let field = "";
  let atStart = true;
  while (i < n) {
    const ch = text[i];
    if (atStart && ch === '"') {
      // Quoted: up to the closing quote; "" is a quote; the rest up to the separator follows as is.
      i++;
      for (;;) {
        const q = text.indexOf('"', i);
        if (q < 0) { field += text.slice(i); i = n; break; }
        field += text.slice(i, q);
        if (text[q + 1] === '"') { field += '"'; i = q + 2; continue; }
        i = q + 1;
        break;
      }
      atStart = false;
      continue;
    }
    atStart = false;
    if (ch === sep) { row.push(field); field = ""; atStart = true; i++; continue; }
    if (ch === "\r" || ch === "\n") {
      row.push(field);
      rows.push(row);
      row = [];
      field = "";
      atStart = true;
      i += ch === "\r" && text[i + 1] === "\n" ? 2 : 1;
      continue;
    }
    field += ch;
    i++;
  }
  if (!atStart || field || row.length) {
    row.push(field);
    rows.push(row);
  }
  return rows;
}

const ERRORS_DE: Record<string, ErrorCode> = {
  "#NULL!": "#NULL!", "#DIV/0!": "#DIV/0!", "#WERT!": "#VALUE!", "#BEZUG!": "#REF!", "#NAME?": "#NAME?", "#ZAHL!": "#NUM!", "#NV": "#N/A",
};

/** A German number: sign or brackets, "." thousands in groups of three, "," decimals, exponent. */
const NUMBER = /^([+-]?)(\(?)(\d{1,3}(?:\.\d{3})+|\d*)(?:,(\d*))?(?:[eE]([+-]?\d+))?(\)?)$/;

function number(t: string): { v: number; grouped: boolean; decimals: boolean; exp: boolean } | null {
  const m = NUMBER.exec(t);
  if (!m || (!m[3] && !m[4]) || (m[2] && !m[6]) || (!m[2] && m[6]) || (m[2] && m[1])) return null;
  const grouped = m[3].includes(".");
  const digits = m[3].replace(/\./g, "") || "0";
  let v = Number(`${digits}.${m[4] ?? ""}${m[5] !== undefined ? `e${m[5]}` : ""}`);
  if (!Number.isFinite(v)) return null;
  // Excel keeps 15 significant digits.
  v = Number(v.toPrecision(15));
  if (m[1] === "-" || m[2]) v = -v;
  return { v: v === 0 ? 0 : v, grouped, decimals: m[4] !== undefined && m[4] !== "", exp: m[5] !== undefined };
}

/** German month abbreviations as Excel reads them ("Okt 26"). */
const MONTHS: Record<string, number> = {
  jan: 1, feb: 2, mrz: 3, "mär": 3, apr: 4, mai: 5, jun: 6, jul: 7, aug: 8, sep: 9, okt: 10, nov: 11, dez: 12,
};

const twoDigitYear = (y: number) => (y < 30 ? 2000 + y : 1900 + y);

/** A field as German Excel reads it from a CSV file. */
export function parseCsvField(raw: string, o: { date1904?: boolean; today?: Date } = {}): CsvValue {
  if (raw === "") return { v: null };
  if (raw.startsWith("=") && raw.length > 1 && isValidLocal(raw)) {
    try {
      return { v: null, f: toFileFormula(raw) };
    } catch {
      return { v: raw };
    }
  }
  const t = raw.trim();
  const d1904 = o.date1904 ?? false;
  if (/^wahr$/i.test(t)) return { v: true };
  if (/^falsch$/i.test(t)) return { v: false };
  const e = ERRORS_DE[t.toUpperCase()];
  if (e) return { v: err(e) };
  let m = /^(.*?)\s?%$/.exec(t);
  if (m) {
    const x = number(m[1]);
    if (x && !x.exp) return { v: Number((x.v / 100).toPrecision(15)), numFmt: x.decimals ? "0.00%" : "0%" };
  }
  m = /^(.*?)\s?€$/.exec(t);
  if (m) {
    const x = number(m[1]);
    if (x && !x.exp) return { v: x.v, numFmt: BUILTIN_FORMATS[x.decimals ? 8 : 6] };
  }
  m = /^€\s?(.*)$/.exec(t);
  if (m) {
    const x = number(m[1]);
    if (x && !x.exp) return { v: x.v, numFmt: x.decimals ? "[$€-2]\\ #,##0.00;[Red]\\-[$€-2]\\ #,##0.00" : "[$€-2]\\ #,##0;[Red]\\-[$€-2]\\ #,##0" };
  }
  const x = number(t);
  if (x) return { v: x.v, numFmt: x.exp ? "0.00E+00" : x.grouped ? (x.decimals ? "#,##0.00" : "#,##0") : undefined };
  // Dates and times.
  const time = (h: string, mi: string, s?: string) => timeToSerial(Number(h), Number(mi), Number(s ?? 0));
  m = /^(\d{1,2})\.(\d{1,2})\.(\d{2}|\d{4})(?:\s+(\d{1,2}):(\d{2})(?::(\d{2}))?)?$/.exec(t) ?? null;
  if (m) {
    const y = m[3].length === 2 ? twoDigitYear(Number(m[3])) : Number(m[3]);
    const d = date(y, Number(m[2]), Number(m[1]), d1904);
    if (d !== null) return m[4] ? { v: d + time(m[4], m[5], m[6]), numFmt: "m/d/yyyy h:mm" } : { v: d, numFmt: "m/d/yyyy" };
  }
  m = /^(\d{4})-(\d{1,2})-(\d{1,2})$/.exec(t);
  if (m) {
    const d = date(Number(m[1]), Number(m[2]), Number(m[3]), d1904);
    if (d !== null) return { v: d, numFmt: "m/d/yyyy" };
  }
  m = /^(\d{1,2})[./](\d{1,2})$/.exec(t);
  if (m) {
    // Day and month of this year ("1.5", "1/2").
    const d = date((o.today ?? new Date()).getFullYear(), Number(m[2]), Number(m[1]), d1904);
    if (d !== null) return { v: d, numFmt: "d-mmm" };
  }
  m = /^([A-Za-zä]{3})[a-zä]*\.?\s+(\d{2}|\d{4})$/.exec(t);
  if (m) {
    const mon = MONTHS[m[1].toLowerCase()];
    if (mon) {
      const y = m[2].length === 2 ? twoDigitYear(Number(m[2])) : Number(m[2]);
      const d = date(y, mon, 1, d1904);
      if (d !== null) return { v: d, numFmt: "mmm-yy" };
    }
  }
  m = /^(\d{1,2}):(\d{2})(?::(\d{2}))?$/.exec(t);
  if (m && Number(m[2]) < 60 && Number(m[3] ?? 0) < 60) return { v: time(m[1], m[2], m[3]), numFmt: m[3] ? "h:mm:ss" : "h:mm" };
  return { v: raw };
}

function date(y: number, m: number, d: number, d1904: boolean): number | null {
  if (m < 1 || m > 12 || d < 1 || d > 31) return null;
  const s = dateToSerial(y, m, d, d1904);
  const back = new Date(Date.UTC(y, m - 1, d));
  return back.getUTCMonth() === m - 1 ? s : null;
}

/** A workbook with one sheet holding the CSV text, read as German Excel does. */
export function readCsv(text: string, opts: string | CsvReadOptions = {}): Workbook {
  const o: CsvReadOptions = typeof opts === "string" ? { name: opts } : opts;
  let body = text.charCodeAt(0) === 0xfeff ? text.slice(1) : text;
  let sep = o.sep ?? ";";
  const m = /^sep=(.)\r?\n/i.exec(body);
  if (m) {
    sep = m[1];
    body = body.slice(m[0].length);
  }
  const wb = Workbook.blank(o.name ?? "Tabelle1");
  wb.date1904 = o.date1904 ?? false;
  const sheet = wb.sheets[0];
  const parse = o.parse ?? ((t: string) => parseCsvField(t, { date1904: wb.date1904, today: o.today }));
  splitCsv(body, sep).forEach((row, r) => {
    row.forEach((field, c) => {
      const p = parse(field);
      if (p.v === null && p.f === undefined) return;
      const cell: Cell = { v: p.v };
      if (p.f !== undefined) cell.f = p.f;
      if (p.numFmt) cell.s = wb.styles.derive(0, (s) => { s.numFmt = p.numFmt!; });
      sheet.set(r, c, cell);
    });
  });
  return wb;
}

export interface CsvWriteOptions {
  /** Separator; default ";". */
  sep?: string;
  /** The workbook's formats, so values are written as the cells show them. */
  styles?: StyleTable;
  date1904?: boolean;
  /** Text of a cell instead of the built-in rules (e.g. the grid's own display). */
  format?: (cell: Cell) => string;
}

/** A sheet as German Excel saves it as CSV (without the byte order mark; see CSV_BOM). */
export function writeCsv(sheet: Sheet, opts: CsvWriteOptions = {}): string {
  const sep = opts.sep ?? ";";
  const used = sheet.usedRange();
  if (!used) return "";
  const text = (c: Cell): string => {
    if (opts.format) return opts.format(c);
    if (c.v === null) return "";
    if (typeof c.v === "string") return c.v;
    const code = opts.styles ? opts.styles.get(c.s).numFmt : "General";
    return formatValue(c.v instanceof CellError ? c.v : c.v, code, { date1904: opts.date1904 }).text;
  };
  const quote = (s: string) => (s.includes(sep) || /["\r\n]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s);
  const lines: string[] = [];
  for (let r = used.r0; r <= used.r1; r++) {
    const fields: string[] = [];
    for (let c = used.c0; c <= used.c1; c++) {
      const cell = sheet.get(r, c);
      fields.push(cell ? quote(text(cell)) : "");
    }
    lines.push(fields.join(sep));
  }
  return lines.map((l) => l + "\r\n").join("");
}
