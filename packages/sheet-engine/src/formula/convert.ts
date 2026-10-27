// Translation between the German form a user types and the file form the model keeps
// (D-062), token by token so that spacing survives. Excel normalises a formula on entry and
// so do we: numbers are rewritten (1,50 → 1,5; 1E3 → 1000), known function names, booleans
// and errors take Excel's spelling, references are upper case, spaces before an argument
// separator, around ":" and at the end are dropped (all recorded from Excel 365).

import { colName } from "../address";
import { ERROR_TEXT_DE, type Workbook } from "../model";
import { lex, type Dialect, type RefEnd, type RefTok, type Token } from "./lexer";
import { XLFN, germanName } from "./names";

const STRUCT_DE: [string, string][] = [["#Alle", "#All"], ["#Daten", "#Data"], ["#Kopfzeilen", "#Headers"], ["#Ergebnisse", "#Totals"], ["#Diese Zeile", "#This Row"]];

/**
 * "=SUMME(A1:A3;2,5)" → "SUM(A1:A3,2.5)". Accepts the text with or without "=". With the
 * workbook, sheet and defined names take their own spelling as Excel does (=tabelle1!a1 →
 * Tabelle1!A1) and 3D references their sheet order.
 */
export function toFileFormula(local: string, wb?: Workbook): string {
  return convert(local, "local", "file", wb);
}

/** "SUM(A1:A3,2.5)" → "SUMME(A1:A3;2,5)", without "=" (the app adds it). */
export function toLocalFormula(file: string, wb?: Workbook): string {
  return convert(file.startsWith("=") ? file.slice(1) : file, "file", "local", wb);
}

let book: Workbook | undefined;

function convert(src: string, from: Dialect, to: Dialect, wb?: Workbook): string {
  book = wb;
  try {
    return convertTokens(src, from, to);
  } finally {
    book = undefined;
  }
}

function convertTokens(src: string, from: Dialect, to: Dialect): string {
  const toks = lex(src, from);
  let out = "";
  let brace = 0;
  for (let i = 0; i < toks.length; i++) {
    const t = toks[i];
    if (t.k === "{") brace++;
    if (t.k === "}") brace--;
    // Excel drops a plus sign in an array constant ({-1.+2} → {-1,2}).
    if (brace > 0 && t.k === "op" && t.text === "+") continue;
    if (t.k === "ws") {
      // Excel drops spaces before a separator and at the end.
      const n = toks[i + 1];
      if (!n || n.k === "sep") continue;
      out += t.text;
      continue;
    }
    out += tokenText(t, from, to);
  }
  return out;
}

function tokenText(t: Token, from: Dialect, to: Dialect): string {
  switch (t.k) {
    case "num": return formulaNumber(t.num!, to);
    case "str": return "\"" + t.str!.replace(/"/g, "\"\"") + "\"";
    case "bool": return to === "file" ? (t.bool ? "TRUE" : "FALSE") : (t.bool ? "WAHR" : "FALSCH");
    case "err": return (t.sheet !== undefined ? sheetPrefix(sheetName(t.sheet)) : "") + errorName(t.err!, to);
    case "ref": return refText(spelled(t.ref!));
    case "name": return (t.sheet !== undefined ? sheetPrefix(sheetName(t.sheet)) : "") + definedName(t.name!);
    case "func": {
      if (!t.known) return t.name!;
      return to === "file" ? t.name! : germanName(t.name!) ?? t.name!;
    }
    case "sep": return to === "file" ? "," : ";";
    case "acol": return to === "file" ? "," : ".";
    case "arow": return ";";
    case "struct": return structText(t.text, from, to);
    default: return t.text;
  }
}

function errorName(code: keyof typeof ERROR_TEXT_DE, to: Dialect): string {
  return to === "file" ? code : ERROR_TEXT_DE[code];
}

function structText(text: string, from: Dialect, to: Dialect): string {
  if (from === to) return text;
  let s = "";
  for (let i = 0; i < text.length; i++) {
    const c = text[i];
    if (c === "'") {
      s += c + (text[i + 1] ?? "");
      i++;
      continue;
    }
    if (from === "local" && c === ";") s += ",";
    else if (from === "file" && c === ",") s += ";";
    else s += c;
  }
  for (const [de, en] of STRUCT_DE) s = from === "local" ? replaceCI(s, de, en) : replaceCI(s, en, de);
  return s;
}

function replaceCI(s: string, a: string, b: string): string {
  return s.replace(new RegExp(a.replace(/[.*+?^${}()|[\]\\#]/g, "\\$&"), "gi"), b);
}

function sheetName(name: string): string {
  return book?.sheetByName(name)?.name ?? name;
}

function definedName(name: string): string {
  const l = name.toLowerCase();
  return book?.names.find((n) => n.name.toLowerCase() === l)?.name ?? name;
}

/** Sheet names in the workbook's spelling, a 3D reference's sheets in their order. */
function spelled(r: RefTok): RefTok {
  if (!book || r.sheet === undefined) return r;
  let sheet = sheetName(r.sheet), sheet2 = r.sheet2 === undefined ? undefined : sheetName(r.sheet2);
  if (sheet2 !== undefined) {
    const i = book.sheets.findIndex((s) => s.name === sheet), j = book.sheets.findIndex((s) => s.name === sheet2);
    if (i >= 0 && j >= 0 && j < i) [sheet, sheet2] = [sheet2, sheet];
  }
  return { ...r, sheet, sheet2 };
}

/** A formula of the model as the .xlsx file stores it: newer functions get their _xlfn. prefix. */
export function toXlsxFormula(f: string): string {
  let out = "";
  let pos = 0;
  for (const t of lex(f, "file")) {
    if (t.k !== "func" || !t.known) continue;
    const p = XLFN[t.name!];
    if (!p) continue;
    out += f.slice(pos, t.s) + p;
    pos = t.e;
  }
  return out + f.slice(pos);
}

/** A formula from an .xlsx file in the model's form: _xlfn./_xlws. prefixes removed. */
export function fromXlsxFormula(f: string): string {
  let out = "";
  let pos = 0;
  for (const t of lex(f, "file")) {
    if (t.k !== "func" || !/^_xl/i.test(t.text)) continue;
    out += f.slice(pos, t.s) + t.name!;
    pos = t.e;
  }
  return out + f.slice(pos);
}

/** A reference written as Excel writes it (an area's corners top-left to bottom-right). */
export function refText(r: RefTok): string {
  let s = "";
  if (r.sheet !== undefined) s = r.sheet2 !== undefined ? sheetPrefix(r.sheet + ":" + r.sheet2, [r.sheet, r.sheet2]) : sheetPrefix(r.sheet);
  if (!r.b) return s + endText(r.a);
  const a = { ...r.a }, b = { ...r.b };
  if (a.row !== undefined && b.row !== undefined && a.row > b.row) {
    [a.row, b.row] = [b.row, a.row];
    [a.rowAbs, b.rowAbs] = [b.rowAbs, a.rowAbs];
  }
  if (a.col !== undefined && b.col !== undefined && a.col > b.col) {
    [a.col, b.col] = [b.col, a.col];
    [a.colAbs, b.colAbs] = [b.colAbs, a.colAbs];
  }
  return s + endText(a) + ":" + endText(b);
}

export function endText(e: RefEnd): string {
  let s = "";
  if (e.col !== undefined) s += (e.colAbs ? "$" : "") + colName(e.col);
  if (e.row !== undefined) s += (e.rowAbs ? "$" : "") + (e.row + 1);
  return s;
}

/** "Tabelle1!" or "'Daten 2'!": quotes when the name is not a plain word or looks like a reference. */
export function sheetPrefix(name: string, parts?: string[]): string {
  const need = (parts ?? [name]).some(needsQuotes);
  return need ? "'" + name.replace(/'/g, "''") + "'!" : name + "!";
}

function needsQuotes(n: string): boolean {
  if (!/^[\p{L}_\\][\p{L}\p{N}_.]*$/u.test(n)) return true;
  // Names that read as a cell (A1, XFD3) or an R1C1 address need quotes too.
  if (/^[A-Za-z]{1,3}\d+$/.test(n)) return true;
  if (/^[RrZz]\d*([CcSs]\d*)?$/.test(n)) return true;
  return false;
}

/**
 * A number as Excel writes it in a formula: 15 significant digits, no exponent within the
 * range Excel writes out in full (1E20 → 100000000000000000000, 1E-7 → 0.0000001).
 */
export function formulaNumber(v: number, to: Dialect): string {
  const s = plainNumber(v);
  return to === "local" ? s.replace(".", ",") : s;
}

export function plainNumber(v: number): string {
  if (v === 0) return "0";
  const neg = v < 0;
  const p = Math.abs(v).toPrecision(15);
  const m = /^(\d)(?:\.(\d+))?(?:e([+-]\d+))?$/.exec(p);
  let digits: string, exp: number;
  if (m) {
    digits = m[1] + (m[2] ?? "");
    exp = m[3] ? Number(m[3]) : 0;
  } else {
    // toPrecision wrote it without exponent, e.g. "123.456000000000".
    const [ip, fp = ""] = p.split(".");
    const all = (ip + fp).replace(/^0+/, "");
    const lead = ip.replace(/^0+/, "");
    exp = lead.length > 0 ? lead.length - 1 : -(fp.length - fp.replace(/^0+/, "").length) - 1;
    digits = all;
  }
  digits = digits.replace(/0+$/, "") || "0";
  let out: string;
  // In full up to 21 digits before and 18 after the point (1E20 → 100000000000000000000).
  if (exp > 20 || (exp < 0 && -exp - 1 + digits.length > 18)) {
    const mant = digits.length > 1 ? digits[0] + "." + digits.slice(1) : digits;
    out = mant + "E" + (exp >= 0 ? "+" : "-") + String(Math.abs(exp)).padStart(2, "0");
  } else if (exp >= 0) {
    out = digits.length <= exp + 1 ? digits + "0".repeat(exp + 1 - digits.length) : digits.slice(0, exp + 1) + "." + digits.slice(exp + 1);
  } else {
    out = "0." + "0".repeat(-exp - 1) + digits;
  }
  return (neg ? "-" : "") + out;
}
