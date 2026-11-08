// Datenschnitte (slicers, D-092): a slicer cache belongs to the workbook and filters one column of an
// Excel table or one field of PivotTables; each slicer on a sheet shows a cache's items as buttons.
// A table slicer keeps its selection in the table's AutoFilter (as Excel's file does: the cache holds no
// items); a PivotTable slicer in the PivotTable's hidden items. The items, their order and whether they
// have data follow Excel's rules (recorded with probe_excel_slicers.py): numbers before text, "(Leer)"
// last, items with data first unless set otherwise, an item has data when a row passing the other
// filters holds it.

import type { PivotTable, PivotValue } from "./pivot";
import type { SheetTable } from "./table";

/** Elementsortierung (Datenschnitteinstellungen). */
export type SlicerSort = "ascending" | "descending" | "dataSourceOrder";
/** Elemente ohne Daten: none = not marked, showItemsWithDataAtTop (default) = marked and last, showItemsWithNoData = marked in place. */
export type SlicerCrossFilter = "none" | "showItemsWithDataAtTop" | "showItemsWithNoData";

export interface SlicerCache {
  /** "Datenschnitt_Fach": the name for formulas (a workbook name with #NV). */
  name: string;
  /** The column or field's name. */
  sourceName: string;
  /** A table's slicer: the table and its column (0-based). */
  table?: { name: string; column: number };
  /** A PivotTable slicer: the PivotTables it filters and the field (index into the source's columns). */
  pivots?: { sheet: string; name: string }[];
  field?: number;
  sort?: SlicerSort;
  /** Beim Sortieren benutzerdefinierte Listen verwenden (on by default). */
  customListSort?: boolean;
  crossFilter?: SlicerCrossFilter;
  /** Elemente ohne Daten ausblenden. */
  hideNoData?: boolean;
  /** The PivotTable cache's id (x14:pivotCacheDefinition pivotCacheId) as read. */
  pivotCacheId?: number;
}

/** A corner of a slicer: a cell and the offset into it in px at 100 % (twoCellAnchor). */
export interface SlicerAnchor { col: number; row: number; x: number; y: number }

export interface Slicer {
  /** The object's name ("Fach", "Fach 1"), unique in the workbook. */
  name: string;
  cache: string;
  /** Beschriftung (the header's text). */
  caption: string;
  /** Kopfzeile anzeigen. */
  showCaption?: boolean;
  /** Schaltflächen: Spalten (1 when unset), Höhe in EMU (251883 when unset, Excel's 0.7 cm). */
  columns?: number;
  rowHeight?: number;
  /** Datenschnitt-Formatvorlage; unset: the workbook's default (SlicerStyleLight1). */
  style?: string;
  /** Excel's column width of the buttons (px at 100 %; unset: the frame's width less the margins). */
  from: SlicerAnchor;
  to: SlicerAnchor;
  /** Size in px at 100 % (kept with the anchor; Excel's default 192 x 275 px = 144 x 206.25 pt). */
  size?: { w: number; h: number };
  /** The drawing anchor as read, and what the model held then (written back verbatim while unchanged). */
  raw?: string;
  rawKey?: string;
}

/** Excel's default button height: 0.7 cm in EMU (rowHeight="251883" in every recorded file). */
export const SLICER_ROW_HEIGHT = 251883;
export const EMU_PER_PT = 12700;
export const BLANK_ITEM = "(Leer)";

export interface SlicerItem { text: string; selected: boolean; hasData: boolean }

/** What a slicer keeps for the file's comparison (everything but its raw XML). */
export function slicerKey(s: Slicer): string {
  const { raw: _r, rawKey: _k, ...rest } = s;
  void _r; void _k;
  return JSON.stringify(rest);
}

/** The next free cache name for a column ("Datenschnitt_Fach", "Datenschnitt_Fach1", …). */
export function nextCacheName(taken: Iterable<string>, source: string): string {
  const used = new Set([...taken].map((n) => n.toLowerCase()));
  const base = `Datenschnitt_${source.replace(/[^\p{L}\p{N}_]/gu, "_")}`;
  if (!used.has(base.toLowerCase())) return base;
  for (let i = 1; ; i++) if (!used.has(`${base}${i}`.toLowerCase())) return `${base}${i}`;
}

/** The next free slicer name ("Fach", "Fach 1", …). */
export function nextSlicerName(taken: Iterable<string>, source: string): string {
  const used = new Set([...taken].map((n) => n.toLowerCase()));
  if (!used.has(source.toLowerCase())) return source;
  for (let i = 1; ; i++) if (!used.has(`${source} ${i}`.toLowerCase())) return `${source} ${i}`;
}

/** Excel's item order: numbers ascending, then text (German collation), "(Leer)" last; descending turns the first two round. */
export function compareItems(a: { text: string; num: number | null }, b: { text: string; num: number | null }, sort: SlicerSort = "ascending"): number {
  const blank = (x: { text: string }) => (x.text === BLANK_ITEM ? 1 : 0);
  if (blank(a) !== blank(b)) return blank(a) - blank(b);
  const rank = (x: { num: number | null }) => (x.num !== null ? 0 : 1);
  let c = rank(a) - rank(b);
  if (!c) c = a.num !== null && b.num !== null ? a.num - b.num : a.text.localeCompare(b.text, "de", { sensitivity: "base" });
  return sort === "descending" ? -c : c;
}

/**
 * The items of a slicer over `rows` (the source's records as shown texts and numbers), with `col` the
 * slicer's column; `selected` its selection (undefined: all) and `others` the other columns' selections
 * (rows they leave out give no data).
 */
export function slicerItems(
  rows: { text: string; num: number | null }[][], col: number, selected: string[] | undefined,
  others: { col: number; values: string[] }[], cache: Pick<SlicerCache, "sort" | "crossFilter" | "hideNoData">,
): SlicerItem[] {
  const seen = new Map<string, { text: string; num: number | null; order: number; data: boolean }>();
  rows.forEach((r, i) => {
    const v = r[col] ?? { text: "", num: null };
    const text = v.text === "" ? BLANK_ITEM : v.text;
    const passes = others.every((o) => { const t = r[o.col]?.text ?? ""; return o.values.includes(t === "" ? BLANK_ITEM : t); });
    const it = seen.get(text.toLowerCase());
    if (!it) seen.set(text.toLowerCase(), { text, num: v.num, order: i, data: passes });
    else if (passes) it.data = true;
  });
  const sort = cache.sort ?? "ascending";
  let list = [...seen.values()];
  if (sort === "dataSourceOrder") list.sort((a, b) => a.order - b.order);
  else list.sort((a, b) => compareItems(a, b, sort));
  const cross = cache.crossFilter ?? "showItemsWithDataAtTop";
  const sel = selected ? new Set(selected.map((s) => (s === "" ? BLANK_ITEM : s).toLowerCase())) : null;
  let items = list.map((x) => ({ text: x.text, selected: !sel || sel.has(x.text.toLowerCase()), hasData: cross === "none" ? true : x.data }));
  if (cache.hideNoData) items = items.filter((x) => x.hasData);
  if (cross === "showItemsWithDataAtTop" || cache.hideNoData) items = [...items.filter((x) => x.hasData), ...items.filter((x) => !x.hasData)];
  return items;
}

/** A table's filter after a slicer's selection: undefined when every item is selected. */
export function tableFilterAfter(t: SheetTable, column: number, selected: string[] | null): SheetTable["filters"] {
  const f = { ...(t.filters ?? {}) };
  if (selected === null) delete f[column];
  else f[column] = selected.map((s) => (s === BLANK_ITEM ? "" : s));
  return Object.keys(f).length ? f : undefined;
}

/** Whether a table row (its shown texts by column) passes the table's filters. */
export function passesTableFilters(filters: SheetTable["filters"], texts: (c: number) => string): boolean {
  if (!filters) return true;
  return Object.entries(filters).every(([c, vals]) => vals.some((v) => v.toLowerCase() === texts(Number(c)).toLowerCase()));
}

const pivotText = (v: PivotValue) => (v === null || v === "" ? BLANK_ITEM : typeof v === "boolean" ? (v ? "WAHR" : "FALSCH") : String(v));

/** A PivotTable slicer's records: the source's rows as texts and numbers. */
export function pivotSlicerRows(data: PivotValue[][]): { text: string; num: number | null }[][] {
  return data.slice(1).filter((r) => r.some((v) => v !== null && v !== "")).map((r) => r.map((v) => ({ text: v === null || v === "" ? "" : pivotText(v), num: typeof v === "number" ? v : null })));
}

/** The PivotTable's hidden items of a field after a slicer's selection (all items selected: none hidden). */
export function pivotHiddenAfter(p: PivotTable, field: number, all: string[], selected: string[] | null): PivotTable["hidden"] {
  const h = { ...(p.hidden ?? {}) };
  const off = selected === null ? [] : all.filter((x) => !selected.some((s) => s.toLowerCase() === x.toLowerCase()));
  if (off.length) h[field] = off;
  else delete h[field];
  return Object.keys(h).length ? h : undefined;
}
