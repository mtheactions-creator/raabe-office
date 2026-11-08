// Structured references to Excel tables (D-080): Tabelle1[Note], [@Note], Tabelle1[[#Kopfzeilen];[Note]],
// Tabelle1[[Name]:[Note]], Tabelle1[#Alle]. Read in both dialects (English and German keywords, "," and
// ";" between the parts); resolved against the workbook's tables to a rectangle.

import type { Rect } from "../address";
import type { Workbook } from "../model";
import { tableData, type SheetTable } from "../table";

export type StructItem = "all" | "data" | "headers" | "totals" | "thisrow";

export interface StructSpec {
  /** The table's name as written, or null for a reference inside a table without its name ([@Note]). */
  table: string | null;
  items: Set<StructItem>;
  /** The first and last column named (null: all columns). */
  col0: string | null;
  col1: string | null;
}

const KEYWORDS: Record<string, StructItem> = {
  "#all": "all", "#alle": "all",
  "#data": "data", "#daten": "data",
  "#headers": "headers", "#kopfzeilen": "headers",
  "#totals": "totals", "#ergebnisse": "totals",
  "#this row": "thisrow", "#diese zeile": "thisrow",
};

/** A column name or keyword without its escapes ('[ '] '# '' stand for the character). */
function unescape(s: string): string {
  let out = "";
  for (let i = 0; i < s.length; i++) {
    if (s[i] === "'" && i + 1 < s.length) {
      out += s[i + 1];
      i++;
    } else out += s[i];
  }
  return out;
}

/** A column name as a structured reference writes it (escapes for [ ] # '). */
export function escapeColumn(name: string): string {
  return name.replace(/['[\]#]/g, "'$&");
}

/** Whether a column name must stand in its own brackets ([@[Note 1]] rather than [@Note 1]). */
export function needsBrackets(name: string): boolean {
  return /[\s,;:[\]#'@"{}$^&*+=\-<>/\\!%()~`|]/.test(name);
}

/** The bracketed parts of an inner list: "[#Headers],[Note]" → ["#Headers", "Note"] with the separators. */
function parts(inner: string): { text: string; sep: string }[] | null {
  const out: { text: string; sep: string }[] = [];
  let i = 0;
  let sep = "";
  while (i < inner.length) {
    while (inner[i] === " ") i++;
    if (i >= inner.length) break;
    if (inner[i] !== "[") return null;
    let j = i + 1, s = "";
    while (j < inner.length && inner[j] !== "]") {
      if (inner[j] === "'" && j + 1 < inner.length) {
        s += inner[j] + inner[j + 1];
        j += 2;
      } else s += inner[j++];
    }
    if (j >= inner.length) return null;
    out.push({ text: s, sep });
    i = j + 1;
    while (inner[i] === " ") i++;
    if (i >= inner.length) break;
    if (inner[i] === "," || inner[i] === ";" || inner[i] === ":") sep = inner[i++];
    else return null;
  }
  return out;
}

/** The parts of a structured reference's text, or null when it is not one. */
export function parseStruct(text: string): StructSpec | null {
  const open = text.indexOf("[");
  if (open < 0 || !text.endsWith("]")) return null;
  const table = open > 0 ? text.slice(0, open) : null;
  let inner = text.slice(open + 1, -1);
  const spec: StructSpec = { table, items: new Set(), col0: null, col1: null };
  const column = (raw: string): boolean => {
    const k = KEYWORDS[unescape(raw).trim().toLowerCase()];
    if (k) {
      spec.items.add(k);
      return true;
    }
    const name = unescape(raw);
    if (spec.col0 === null) spec.col0 = name;
    else if (spec.col1 === null) spec.col1 = name;
    else return false;
    return true;
  };
  if (inner.startsWith("@")) {
    spec.items.add("thisrow");
    inner = inner.slice(1).trim();
    if (inner === "") return spec;
    if (!inner.startsWith("[")) return column(inner) ? spec : null;
  }
  if (!/(^|[^'])\[/.test(inner)) {
    if (inner === "") return spec;
    return column(inner) ? spec : null;
  }
  const ps = parts(inner);
  if (!ps) return null;
  for (let k = 0; k < ps.length; k++) {
    const p = ps[k];
    if (p.sep === ":" && (spec.col0 === null || KEYWORDS[unescape(p.text).trim().toLowerCase()])) return null;
    if (!column(p.text)) return null;
    if (p.sep !== ":" && spec.col1 !== null) return null;
  }
  return spec;
}

/** The table a structured reference names, with its sheet: by name, or the table holding the cell. */
export function findTable(wb: Workbook, name: string | null, sheet: number, row: number, col: number): { table: SheetTable; sheet: number } | null {
  if (name !== null) {
    const l = name.toLowerCase();
    for (let s = 0; s < wb.sheets.length; s++) {
      const t = wb.sheets[s].tables?.find((x) => x.name.toLowerCase() === l);
      if (t) return { table: t, sheet: s };
    }
    return null;
  }
  const t = wb.sheets[sheet]?.tables?.find((x) => row >= x.ref.r0 && row <= x.ref.r1 && col >= x.ref.c0 && col <= x.ref.c1);
  return t ? { table: t, sheet } : null;
}

export type StructResult = { sheet: number; rect: Rect } | "ref" | "value";

/**
 * The rectangle a structured reference stands for, seen from a cell: "ref" when the table, a column
 * or a part (Kopfzeile, Ergebniszeile) is not there, "value" for [@…] outside the table's data rows.
 */
export function resolveStruct(wb: Workbook, spec: StructSpec, sheet: number, row: number, col: number): StructResult {
  const found = findTable(wb, spec.table, sheet, row, col);
  if (!found) return "ref";
  const t = found.table;
  const colIndex = (name: string): number => {
    const l = name.trim().toLowerCase();
    return t.columns.findIndex((c) => c.name.toLowerCase() === l);
  };
  let c0 = t.ref.c0, c1 = t.ref.c1;
  if (spec.col0 !== null) {
    const a = colIndex(spec.col0);
    const b = spec.col1 !== null ? colIndex(spec.col1) : a;
    if (a < 0 || b < 0) return "ref";
    c0 = t.ref.c0 + Math.min(a, b);
    c1 = t.ref.c0 + Math.max(a, b);
  }
  const data = tableData(t);
  const head = t.header ? t.ref.r0 : -1;
  const tot = t.totals ? t.ref.r1 : -1;
  const it = spec.items;
  let r0: number, r1: number;
  if (it.has("thisrow")) {
    if (found.sheet !== sheet || row < data.r0 || row > data.r1) return "value";
    r0 = r1 = row;
  } else if (it.has("all")) {
    r0 = t.ref.r0;
    r1 = t.ref.r1;
  } else {
    const h = it.has("headers"), d = it.has("data") || it.size === 0, s = it.has("totals");
    if (h && s && !d) return "ref";
    if (h && head < 0) return "ref";
    if (s && tot < 0) return "ref";
    r0 = h ? head : d ? data.r0 : tot;
    r1 = s ? tot : d ? data.r1 : head;
  }
  if (r1 < r0) return "ref";
  return { sheet: found.sheet, rect: { r0, c0, r1, c1 } };
}

/**
 * A structured reference in the file's form: always with the table's name (`table` when the text
 * leaves it out), [@Note] as Tabelle1[[#This Row],[Note]] (recorded from Excel's files).
 */
export function structToFile(text: string, table: string | null): string {
  const spec = parseStruct(text);
  if (!spec) return text;
  const name = spec.table ?? table;
  if (name === null) return text;
  const body = text.slice(text.indexOf("["));
  if (!spec.items.has("thisrow")) return name + body;
  if (spec.items.size > 1) return name + body;
  const cols = spec.col0 === null ? "" : spec.col1 === null ? `,[${escapeColumn(spec.col0)}]` : `,[${escapeColumn(spec.col0)}]:[${escapeColumn(spec.col1)}]`;
  return cols ? `${name}[[#This Row]${cols}]` : `${name}[#This Row]`;
}

/**
 * A structured reference as Excel shows it: [#This Row] as "@" (Tabelle1[[#This Row],[Note]] is
 * Tabelle1[@Note]), and without the table's name when the formula stands in that table.
 */
export function structToShown(text: string, ownTable: string | null): string {
  const spec = parseStruct(text);
  if (!spec) return text;
  const own = spec.table !== null && ownTable !== null && spec.table.toLowerCase() === ownTable.toLowerCase();
  const name = own ? "" : spec.table ?? "";
  const body = text.slice(text.indexOf("["));
  // Excel's file writes the table's name alone as Noten[]; it shows Noten.
  if (body === "[]" && spec.table !== null) return spec.table;
  if (!spec.items.has("thisrow") || spec.items.size !== 1) return name + body;
  const one = (c: string): string => (needsBrackets(c) ? `[${escapeColumn(c)}]` : escapeColumn(c));
  if (spec.col0 === null) return `${name}[@]`;
  if (spec.col1 === null) return `${name}[@${one(spec.col0)}]`;
  return `${name}[@[${escapeColumn(spec.col0)}]:[${escapeColumn(spec.col1)}]]`;
}

/** Whether a formula (file form) names a table or a column that is not there (Excel refuses it on entry). */
export function structProblem(wb: Workbook, texts: string[], sheet: number, row: number, col: number): boolean {
  for (const text of texts) {
    const spec = parseStruct(text);
    if (!spec) return true;
    const found = findTable(wb, spec.table, sheet, row, col);
    if (!found) return true;
    const has = (n: string | null) => n === null || found.table.columns.some((c) => c.name.toLowerCase() === n.trim().toLowerCase());
    if (!has(spec.col0) || !has(spec.col1)) return true;
  }
  return false;
}

/** A structured reference with a column renamed (the table's own references to it). */
export function renameColumn(text: string, from: string, to: string): string {
  const esc = (x: string) => escapeColumn(x).replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  const re = new RegExp(`(\\[|@)${esc(from)}(?=\\])`, "gi");
  const plain = (x: string) => (needsBrackets(x) ? `[${escapeColumn(x)}]` : escapeColumn(x));
  return text.replace(re, (_m, pre: string) => (pre === "@" ? "@" + plain(to) : "[" + escapeColumn(to)));
}
