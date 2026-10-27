// Tokens of a formula in either of Excel's two forms (D-062):
// - "file": English names, "," between arguments, "." as decimal point, array constants
//   {1,2;3,4}; what the model stores (Excel's Range.Formula2 without "=").
// - "local": German names, ";" between arguments, "," as decimal point, array constants
//   {1.2;3.4}; what the user types and sees (Range.Formula2Local).
// Every token keeps its place in the source, so references can be rewritten in place and
// coloured while typing.

import { MAX_COLS, MAX_ROWS, colIndex } from "../address";
import { ERROR_TEXT_DE, type ErrorCode } from "../model";
import { englishName, isEnglishName } from "./names";

export type Dialect = "file" | "local";

export type TokKind =
  | "num" | "str" | "bool" | "err" | "ref" | "func" | "name" | "struct"
  | "op" | "(" | ")" | "sep" | "{" | "}" | "acol" | "arow" | "ws" | "bad";

/** One end of a reference: a cell, or a whole column (row undefined) or row (col undefined). */
export interface RefEnd {
  row?: number;
  col?: number;
  rowAbs: boolean;
  colAbs: boolean;
}

export interface RefTok {
  /** Sheet name without quotes; sheet2 for a 3D reference Tabelle1:Tabelle3!A1. */
  sheet?: string;
  sheet2?: string;
  a: RefEnd;
  /** Second corner of an area. */
  b?: RefEnd;
}

export interface Token {
  k: TokKind;
  /** Start and end offsets in the source text. */
  s: number;
  e: number;
  text: string;
  num?: number;
  str?: string;
  bool?: boolean;
  err?: ErrorCode;
  /** func: the English name when known (upper case), else the text as written. */
  name?: string;
  known?: boolean;
  ref?: RefTok;
  /** Offset where the reference part starts (after the sheet prefix). */
  refStart?: number;
  /** Sheet prefix of a name or error (Tabelle1!Name, Tabelle1!#BEZUG!). */
  sheet?: string;
}

const FILE_ERRORS: ErrorCode[] = ["#NULL!", "#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#NUM!", "#N/A", "#SPILL!", "#CALC!", "#GETTING_DATA"];
const LOCAL_ERRORS = new Map<string, ErrorCode>(FILE_ERRORS.map((c) => [ERROR_TEXT_DE[c].toUpperCase(), c]));
const FILE_ERROR_MAP = new Map<string, ErrorCode>(FILE_ERRORS.map((c) => [c, c]));

const IDENT = /[\p{L}\p{N}_.\\?]/u;
const IDENT_START = /[\p{L}_\\]/u;
const CELL_RE = /(\$?)([A-Za-z]{1,3})(\$?)(\d{1,7})/y;
const COL_RE = /(\$?)([A-Za-z]{1,3})/y;
const ROW_RE = /(\$?)(\d{1,7})/y;
const WS = /[ \n\r\t]/;

function isIdent(c: string | undefined): boolean {
  return c !== undefined && IDENT.test(c);
}

function cellAt(src: string, i: number): { end: RefEnd; next: number } | null {
  CELL_RE.lastIndex = i;
  const m = CELL_RE.exec(src);
  if (!m) return null;
  const col = colIndex(m[2]);
  const row = Number(m[4]) - 1;
  if (col < 0 || col >= MAX_COLS || row < 0 || row >= MAX_ROWS) return null;
  return { end: { row, col, rowAbs: m[3] === "$", colAbs: m[1] === "$" }, next: CELL_RE.lastIndex };
}

function colAt(src: string, i: number): { end: RefEnd; next: number } | null {
  COL_RE.lastIndex = i;
  const m = COL_RE.exec(src);
  if (!m) return null;
  const col = colIndex(m[2]);
  if (col < 0 || col >= MAX_COLS) return null;
  return { end: { col, colAbs: m[1] === "$", rowAbs: false }, next: COL_RE.lastIndex };
}

function rowAt(src: string, i: number): { end: RefEnd; next: number } | null {
  ROW_RE.lastIndex = i;
  const m = ROW_RE.exec(src);
  if (!m) return null;
  const row = Number(m[2]) - 1;
  if (row < 0 || row >= MAX_ROWS) return null;
  return { end: { row, rowAbs: m[1] === "$", colAbs: false }, next: ROW_RE.lastIndex };
}

function skipWs(src: string, i: number): number {
  while (i < src.length && WS.test(src[i])) i++;
  return i;
}

/** A reference body at i (A1, A1:B2, A:B, 1:2); null when there is none. */
function refAt(src: string, i: number): { ref: RefTok; next: number } | null {
  const endOk = (j: number) => !isIdent(src[j]) && src[j] !== "(" && src[j] !== "[" && src[j] !== "!" && src[j] !== "'";
  const cell = cellAt(src, i);
  if (cell && endOk(cell.next)) {
    // An area: the colon may have spaces around it, which Excel drops.
    const j = skipWs(src, cell.next);
    if (src[j] === ":") {
      const k = skipWs(src, j + 1);
      const c2 = cellAt(src, k);
      if (c2 && endOk(c2.next)) return { ref: { a: cell.end, b: c2.end }, next: c2.next };
    }
    return { ref: { a: cell.end }, next: cell.next };
  }
  const col = colAt(src, i);
  if (col && src[col.next] === ":") {
    const c2 = colAt(src, col.next + 1);
    if (c2 && endOk(c2.next) && !/\d/.test(src[c2.next] ?? "")) return { ref: { a: col.end, b: c2.end }, next: c2.next };
  }
  const row = rowAt(src, i);
  if (row && src[row.next] === ":") {
    const r2 = rowAt(src, row.next + 1);
    if (r2 && endOk(r2.next) && !/[.,]/.test(src[r2.next] ?? "")) return { ref: { a: row.end, b: r2.end }, next: r2.next };
  }
  return null;
}

/** A sheet prefix at i ('Daten 2'!, Tabelle1!, Tabelle1:Tabelle3!); null when there is none. */
function sheetAt(src: string, i: number): { sheet: string; sheet2?: string; next: number } | null {
  const one = (j: number): { name: string; next: number } | null => {
    if (src[j] === "'") {
      let name = "";
      let k = j + 1;
      while (k < src.length) {
        if (src[k] === "'") {
          if (src[k + 1] === "'") { name += "'"; k += 2; continue; }
          return { name, next: k + 1 };
        }
        name += src[k++];
      }
      return null;
    }
    let k = j;
    while (k < src.length && (isIdent(src[k]) || src[k] === "-")) k++;
    if (k === j) return null;
    return { name: src.slice(j, k), next: k };
  };
  const a = one(i);
  if (!a) return null;
  if (src[a.next] === "!") {
    // 'Tabelle1:Daten 2'!A1 is a 3D reference (":" cannot be part of a sheet name).
    const colon = a.name.indexOf(":");
    if (colon > 0) return { sheet: a.name.slice(0, colon), sheet2: a.name.slice(colon + 1), next: a.next + 1 };
    return { sheet: a.name, next: a.next + 1 };
  }
  if (src[a.next] === ":") {
    const b = one(a.next + 1);
    if (b && src[b.next] === "!") return { sheet: a.name, sheet2: b.name, next: b.next + 1 };
  }
  return null;
}

function numberAt(src: string, i: number, dec: string): { v: number; next: number } | null {
  let j = i;
  let digits = "";
  while (j < src.length && src[j] >= "0" && src[j] <= "9") digits += src[j++];
  if (src[j] === dec) {
    let frac = "";
    let k = j + 1;
    while (k < src.length && src[k] >= "0" && src[k] <= "9") frac += src[k++];
    if (digits || frac) {
      digits += "." + frac;
      j = k;
    }
  }
  if (!digits || digits === ".") return null;
  digits = cut15(digits);
  if (src[j] === "e" || src[j] === "E") {
    let k = j + 1;
    let exp = "";
    if (src[k] === "+" || src[k] === "-") exp += src[k++];
    let d = "";
    while (k < src.length && src[k] >= "0" && src[k] <= "9") d += src[k++];
    if (d) {
      digits += "e" + exp + d;
      j = k;
    }
  }
  let v = Number(digits);
  if (v !== 0 && Math.abs(v) < 2.2250738585072014e-308) v = 0;
  return { v, next: j };
}

/** Excel reads at most 15 significant digits of a number and drops the rest (does not round). */
function cut15(digits: string): string {
  let sig = 0;
  let out = "";
  let started = false;
  for (const ch of digits) {
    if (ch === ".") { out += ch; continue; }
    if (ch !== "0") started = true;
    if (started) sig++;
    out += sig > 15 ? "0" : ch;
  }
  return out;
}

function errorAt(src: string, i: number, dialect: Dialect): { code: ErrorCode; next: number } | null {
  const map = dialect === "file" ? FILE_ERROR_MAP : LOCAL_ERRORS;
  let best: { code: ErrorCode; next: number } | null = null;
  for (const [text, code] of map) {
    if (src.substr(i, text.length).toUpperCase() === text && (!best || i + text.length > best.next)) best = { code, next: i + text.length };
  }
  return best;
}

/** Strips the prefixes Excel writes in files before newer functions (_xlfn., _xlws.). */
export function stripPrefix(name: string): string {
  let n = name;
  for (;;) {
    const u = n.toUpperCase();
    if (u.startsWith("_XLFN.")) n = n.slice(6);
    else if (u.startsWith("_XLWS.")) n = n.slice(6);
    else return n;
  }
}

/**
 * Splits a formula into tokens. A leading "=" is skipped (its offset is kept). Never throws:
 * text it cannot read becomes a "bad" token.
 */
export function lex(src: string, dialect: Dialect): Token[] {
  const out: Token[] = [];
  const dec = dialect === "file" ? "." : ",";
  const argSep = dialect === "file" ? "," : ";";
  const colSep = dialect === "file" ? "," : ".";
  let brace = 0;
  let i = 0;
  if (src[0] === "=") i = 1;
  const push = (t: Partial<Token> & { k: TokKind }, s: number, e: number) => out.push({ ...t, s, e, text: src.slice(s, e) } as Token);
  // Whether the previous significant token ends an operand (then "#" is a spill operator).
  const lastOperand = () => {
    for (let j = out.length - 1; j >= 0; j--) {
      const t = out[j];
      if (t.k === "ws") return false;
      return t.k === "ref" || t.k === "name" || t.k === ")";
    }
    return false;
  };
  while (i < src.length) {
    const c = src[i];
    const s = i;
    if (WS.test(c)) {
      i = skipWs(src, i);
      push({ k: "ws" }, s, i);
      continue;
    }
    if (c === "\"") {
      let v = "";
      let j = i + 1;
      let closed = false;
      while (j < src.length) {
        if (src[j] === "\"") {
          if (src[j + 1] === "\"") { v += "\""; j += 2; continue; }
          closed = true;
          j++;
          break;
        }
        v += src[j++];
      }
      i = j;
      push(closed ? { k: "str", str: v } : { k: "bad" }, s, i);
      continue;
    }
    if (c === "#") {
      if (lastOperand()) {
        i++;
        push({ k: "op" }, s, i);
        continue;
      }
      const e = errorAt(src, i, dialect);
      if (e) {
        i = e.next;
        push({ k: "err", err: e.code }, s, i);
        continue;
      }
      i++;
      push({ k: "bad" }, s, i);
      continue;
    }
    if (brace > 0 && (c === colSep || c === ";")) {
      // Inside an array constant ";" separates rows in both forms, "," or "." columns.
      i++;
      push({ k: c === ";" ? "arow" : "acol" }, s, i);
      continue;
    }
    if (c === argSep) {
      i++;
      push({ k: "sep" }, s, i);
      continue;
    }
    if (c === "(" || c === ")") {
      i++;
      push({ k: c }, s, i);
      continue;
    }
    if (c === "{" || c === "}") {
      brace += c === "{" ? 1 : -1;
      i++;
      push({ k: c }, s, i);
      continue;
    }
    if (c === "<" || c === ">") {
      i += src[i + 1] === "=" || (c === "<" && src[i + 1] === ">") ? 2 : 1;
      push({ k: "op" }, s, i);
      continue;
    }
    if ("+-*/^&=%:@".includes(c)) {
      i++;
      push({ k: "op" }, s, i);
      continue;
    }
    if (c === "[") {
      // A structured reference without a table name ([@Spalte]); kept as written.
      i = bracketEnd(src, i);
      push({ k: "struct" }, s, i);
      continue;
    }
    // Sheet-qualified reference, name or error.
    const sh = c === "'" || IDENT_START.test(c) || /\d/.test(c) ? sheetAt(src, i) : null;
    if (sh) {
      const r = refAt(src, sh.next);
      if (r) {
        i = r.next;
        push({ k: "ref", ref: { ...r.ref, sheet: sh.sheet, sheet2: sh.sheet2 }, refStart: sh.next }, s, i);
        continue;
      }
      if (src[sh.next] === "#") {
        const e = errorAt(src, sh.next, dialect);
        if (e) {
          i = e.next;
          push({ k: "err", err: e.code, sheet: sh.sheet, refStart: sh.next }, s, i);
          continue;
        }
      }
      let j = sh.next;
      while (j < src.length && isIdent(src[j])) j++;
      if (j > sh.next) {
        i = j;
        push({ k: "name", name: src.slice(sh.next, j), sheet: sh.sheet, refStart: sh.next }, s, i);
        continue;
      }
    }
    if (c === "$" || /[A-Za-z0-9]/.test(c)) {
      const r = refAt(src, i);
      if (r) {
        i = r.next;
        push({ k: "ref", ref: r.ref, refStart: s }, s, i);
        continue;
      }
    }
    if ((c >= "0" && c <= "9") || c === dec) {
      const n = numberAt(src, i, dec);
      if (n) {
        i = n.next;
        push({ k: "num", num: n.v }, s, i);
        continue;
      }
    }
    if (IDENT_START.test(c)) {
      let j = i;
      // In a German array constant "." separates columns ({WAHR.FALSCH}).
      while (j < src.length && isIdent(src[j]) && !(brace > 0 && src[j] === colSep)) j++;
      const word = src.slice(i, j);
      i = j;
      if (src[j] === "[") {
        i = bracketEnd(src, j);
        push({ k: "struct", name: word }, s, i);
        continue;
      }
      if (src[j] === "(") {
        const bare = stripPrefix(word);
        const up = bare.toUpperCase();
        if (dialect === "file") {
          const known = isEnglishName(up);
          push({ k: "func", name: known ? up : bare, known }, s, i);
        } else {
          const en = englishName(up);
          push({ k: "func", name: en ?? word, known: en !== undefined }, s, i);
        }
        continue;
      }
      const up = word.toUpperCase();
      const t = dialect === "file" ? "TRUE" : "WAHR";
      const f = dialect === "file" ? "FALSE" : "FALSCH";
      if (up === t || up === f) {
        push({ k: "bool", bool: up === t }, s, i);
        continue;
      }
      push({ k: "name", name: word }, s, i);
      continue;
    }
    i++;
    push({ k: "bad" }, s, i);
  }
  return out;
}

function bracketEnd(src: string, i: number): number {
  let depth = 0;
  let j = i;
  while (j < src.length) {
    const ch = src[j];
    if (ch === "'" && j + 1 < src.length) { j += 2; continue; }
    if (ch === "[") depth++;
    else if (ch === "]") {
      depth--;
      if (depth === 0) return j + 1;
    }
    j++;
  }
  return j;
}
