// Which rows an AutoFilter column lets through: chosen values, Benutzerdefinierter AutoFilter
// (Zahlenfilter/Textfilter: comparisons, * and ? wildcards), Top 10 (items or percent) and
// Über/Unter dem Durchschnitt, with the rules recorded from Excel
// (probe_excel_filter_sub.py, results/2026-10-28-excel-filter-sub.json).

import type { Cell, FilterColumn } from "@raabe/sheet-engine";

/** The limit a Top 10 or average filter needs, from the column's numbers. */
export function prepare(crit: FilterColumn, numbers: number[]): FilterColumn {
  if (crit.top10) {
    const t = crit.top10;
    const sorted = [...numbers].sort((a, b) => (t.top ? b - a : a - b));
    const n = t.percent ? Math.max(1, Math.floor((sorted.length * t.val) / 100)) : t.val;
    const limit = sorted.length ? sorted[Math.min(sorted.length, n) - 1] : undefined;
    return { ...crit, top10: { ...t, filterVal: limit } };
  }
  if (crit.dynamic) {
    const avg = numbers.length ? numbers.reduce((a, b) => a + b, 0) / numbers.length : 0;
    return { ...crit, dynamic: { ...crit.dynamic, val: avg } };
  }
  return crit;
}

/** A filter value with * and ? as Excel reads it (~ escapes them), case-insensitive. */
function wildcard(val: string): RegExp {
  let re = "";
  for (let i = 0; i < val.length; i++) {
    const ch = val[i];
    if (ch === "~" && i + 1 < val.length) re += val[++i].replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
    else if (ch === "*") re += "[\\s\\S]*";
    else if (ch === "?") re += "[\\s\\S]";
    else re += ch.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  }
  return new RegExp(`^${re}$`, "i");
}

function condition(op: string, val: string, cell: Cell | undefined, text: string): boolean {
  const v = cell?.v ?? null;
  const num = val.trim() !== "" && !Number.isNaN(Number(val)) ? Number(val) : null;
  if (typeof v === "number" && num !== null) {
    switch (op) {
      case "equal": return v === num;
      case "notEqual": return v !== num;
      case "greaterThan": return v > num;
      case "greaterThanOrEqual": return v >= num;
      case "lessThan": return v < num;
      case "lessThanOrEqual": return v <= num;
    }
  }
  if (op === "equal" || op === "notEqual") {
    const hit = val === "" ? text === "" : wildcard(val).test(text);
    return op === "equal" ? hit : !hit;
  }
  // Comparing text with text (numbers do not pass a text comparison and the other way round).
  if (typeof v === "number" || num !== null || text === "") return false;
  const c = text.localeCompare(val, "de", { sensitivity: "base" });
  return op === "greaterThan" ? c > 0 : op === "greaterThanOrEqual" ? c >= 0 : op === "lessThan" ? c < 0 : c <= 0;
}

/** Whether the row passes this column's filter (`text` is what the cell shows). */
export function passes(crit: FilterColumn, cell: Cell | undefined, text: string): boolean {
  if (crit.custom) {
    const r = crit.custom.items.map((i) => condition(i.op, i.val, cell, text));
    return crit.custom.and ? r.every(Boolean) : r.some(Boolean);
  }
  const v = cell?.v;
  if (crit.top10) {
    const lim = crit.top10.filterVal;
    return typeof v === "number" && lim !== undefined && (crit.top10.top ? v >= lim : v <= lim);
  }
  if (crit.dynamic) {
    const avg = crit.dynamic.val ?? 0;
    return typeof v === "number" && (crit.dynamic.type === "aboveAverage" ? v > avg : v < avg);
  }
  if (crit.values) return text === "" ? !!crit.blank : crit.values.includes(text);
  return true;
}

/** Benutzerdefinierter AutoFilter's operators (Excel's labels) and how each is written in the file. */
export const CUSTOM_OPS: [string, string][] = [
  ["eq", "ist gleich"], ["ne", "ist nicht gleich"], ["gt", "ist größer als"], ["ge", "ist größer oder gleich"], ["lt", "ist kleiner als"], ["le", "ist kleiner oder gleich"],
  ["begins", "beginnt mit"], ["notBegins", "beginnt nicht mit"], ["ends", "endet mit"], ["notEnds", "endet nicht mit"], ["contains", "enthält"], ["notContains", "enthält nicht"],
];

/** A typed value in the file's form: German numbers with ".", other text as it is. */
function fileValue(v: string): string {
  const t = v.trim();
  return /^[+-]?\d+(,\d+)?$/.test(t) ? t.replace(",", ".") : v;
}

/** Dialog operator and typed value → the file's customFilter. */
export function toCustom(op: string, val: string): { op: string; val: string } {
  switch (op) {
    case "ne": return { op: "notEqual", val: fileValue(val) };
    case "gt": return { op: "greaterThan", val: fileValue(val) };
    case "ge": return { op: "greaterThanOrEqual", val: fileValue(val) };
    case "lt": return { op: "lessThan", val: fileValue(val) };
    case "le": return { op: "lessThanOrEqual", val: fileValue(val) };
    case "begins": return { op: "equal", val: `${val}*` };
    case "notBegins": return { op: "notEqual", val: `${val}*` };
    case "ends": return { op: "equal", val: `*${val}` };
    case "notEnds": return { op: "notEqual", val: `*${val}` };
    case "contains": return { op: "equal", val: `*${val}*` };
    case "notContains": return { op: "notEqual", val: `*${val}*` };
    default: return { op: "equal", val: fileValue(val) };
  }
}

/** The file's customFilter → the dialog's operator and value (as Excel shows it again). */
export function fromCustom(c: { op: string; val: string }): { op: string; val: string } {
  const v = c.val;
  const shownNum = /^[+-]?\d+(\.\d+)?$/.test(v) ? v.replace(".", ",") : v;
  const inner = v.length > 1 && v.startsWith("*") && v.endsWith("*") ? v.slice(1, -1) : null;
  const neg = c.op === "notEqual";
  if (c.op === "equal" || neg) {
    if (inner !== null && !/[*?]/.test(inner)) return { op: neg ? "notContains" : "contains", val: inner };
    if (v.endsWith("*") && !/[*?]/.test(v.slice(0, -1))) return { op: neg ? "notBegins" : "begins", val: v.slice(0, -1) };
    if (v.startsWith("*") && !/[*?]/.test(v.slice(1))) return { op: neg ? "notEnds" : "ends", val: v.slice(1) };
    return { op: neg ? "ne" : "eq", val: shownNum };
  }
  const map: Record<string, string> = { greaterThan: "gt", greaterThanOrEqual: "ge", lessThan: "lt", lessThanOrEqual: "le" };
  return { op: map[c.op] ?? "eq", val: shownNum };
}

type FilterEntry = { label: string; kind: "custom" | "top10" | "above" | "below"; op?: string; op2?: string } | "-";

/** Excel's Zahlenfilter submenu. */
export const NUMBER_FILTERS: FilterEntry[] = [
  { label: "Ist gleich...", kind: "custom", op: "eq" }, { label: "Ist nicht gleich...", kind: "custom", op: "ne" }, "-",
  { label: "Größer als...", kind: "custom", op: "gt" }, { label: "Größer oder gleich...", kind: "custom", op: "ge" },
  { label: "Kleiner als...", kind: "custom", op: "lt" }, { label: "Kleiner oder gleich...", kind: "custom", op: "le" },
  { label: "Zwischen...", kind: "custom", op: "ge", op2: "le" }, "-",
  { label: "Top 10...", kind: "top10" }, { label: "Über dem Durchschnitt", kind: "above" }, { label: "Unter dem Durchschnitt", kind: "below" }, "-",
  { label: "Benutzerdefinierter Filter...", kind: "custom", op: "eq" },
];

/** Excel's Textfilter submenu. */
export const TEXT_FILTERS: FilterEntry[] = [
  { label: "Ist gleich...", kind: "custom", op: "eq" }, { label: "Ist nicht gleich...", kind: "custom", op: "ne" }, "-",
  { label: "Beginnt mit...", kind: "custom", op: "begins" }, { label: "Endet mit...", kind: "custom", op: "ends" }, "-",
  { label: "Enthält...", kind: "custom", op: "contains" }, { label: "Enthält nicht...", kind: "custom", op: "notContains" }, "-",
  { label: "Benutzerdefinierter Filter...", kind: "custom", op: "eq" },
];
