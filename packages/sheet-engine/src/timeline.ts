// Zeitachsen (timelines, D-121): a timeline cache belongs to the workbook and filters one date field of PivotTables;
// each timeline on a sheet shows the field's time as periods (Jahre, Quartale, Monate, Tage) and a selection of
// consecutive periods. As in Excel's file, the filter itself is the PivotTables' date filter (type dateBetween, whole
// days from the first period's first day to the last period's last day); the timeline cache repeats it. Recorded with
// probe_excel_timeline.py (fixture excel-timeline.xlsx): cache names "NativeZeitachse_Datum", "NativeZeitachse_Datum1";
// the bounds run from 1 January of the first year to 1 January after the last; the selection label as Excel writes it
// ("Okt 2025 - Jan 2026", "Q1 - Q2 2026", "4 - 21 Sep 2025", "Alle Zeiträume").

import { dateToSerial, MONTHS_SHORT_DE, serialToDate } from "./dates";
import type { PivotTable } from "./pivot";
import type { SlicerAnchor } from "./slicer";

/** Zeitebene: 0 Jahre, 1 Quartale, 2 Monate, 3 Tage (Excel's level numbers). */
export type TimeLevel = 0 | 1 | 2 | 3;
export const TIME_LEVELS = ["Jahre", "Quartale", "Monate", "Tage"] as const;

export interface TimelineCache {
  /** "NativeZeitachse_Datum": the name for formulas (a workbook name with #NV). */
  name: string;
  /** The date field's name. */
  sourceName: string;
  /** The PivotTables it filters. */
  pivots: { sheet: string; name: string }[];
  /** The PivotTable cache's id (x14 pivotCacheId) as read. */
  pivotCacheId?: number;
}

export interface Timeline {
  name: string;
  cache: string;
  caption: string;
  /** The Zeitebene shown, and the one the selection was made at. */
  level: TimeLevel;
  selectionLevel?: TimeLevel;
  /** The first date shown at the left (serial); unset: the selection's start or the bounds'. */
  scrollPosition?: number;
  showHeader?: boolean;
  showSelectionLabel?: boolean;
  showTimeLevel?: boolean;
  showHorizontalScrollbar?: boolean;
  /** Zeitachsenformatvorlage; unset: TimeSlicerStyleLight1. */
  style?: string;
  from: SlicerAnchor;
  to: SlicerAnchor;
  /** Size in px at 100 % (Excel's default 350 x 144 px = 262.5 x 108 pt). */
  size?: { w: number; h: number };
  raw?: string;
  rawKey?: string;
}

/** Excel's default timeline: 262.5 x 108 pt. */
export const TIMELINE_W = 350, TIMELINE_H = 144;

export function timelineKey(t: Timeline): string {
  const { raw: _r, rawKey: _k, ...rest } = t;
  void _r; void _k;
  return JSON.stringify(rest);
}

/** The next free cache name ("NativeZeitachse_Datum", "NativeZeitachse_Datum1", …). */
export function nextTimelineCacheName(taken: Iterable<string>, source: string): string {
  const used = new Set([...taken].map((n) => n.toLowerCase()));
  const base = `NativeZeitachse_${source.replace(/[^\p{L}\p{N}_]/gu, "_")}`;
  if (!used.has(base.toLowerCase())) return base;
  for (let i = 1; ; i++) if (!used.has(`${base}${i}`.toLowerCase())) return `${base}${i}`;
}

// ----- periods ---------------------------------------------------------------------------------

/** The first day of the period that holds `serial`. */
export function periodStart(serial: number, level: TimeLevel): number {
  const d = serialToDate(Math.floor(serial));
  if (level === 3) return Math.floor(serial);
  if (level === 2) return dateToSerial(d.y, d.m, 1);
  if (level === 1) return dateToSerial(d.y, Math.floor((d.m - 1) / 3) * 3 + 1, 1);
  return dateToSerial(d.y, 1, 1);
}

/** The first day of the period `n` periods after the one that starts at (or holds) `serial`. */
export function addPeriods(serial: number, level: TimeLevel, n: number): number {
  const s = periodStart(serial, level);
  if (level === 3) return s + n;
  const d = serialToDate(s);
  if (level === 2) return dateToSerial(d.y, d.m + n, 1);
  if (level === 1) return dateToSerial(d.y, d.m + 3 * n, 1);
  return dateToSerial(d.y + n, 1, 1);
}

/** The last day of the period that holds `serial`. */
export function periodEnd(serial: number, level: TimeLevel): number {
  return addPeriods(serial, level, 1) - 1;
}

/** How many periods lie from the one holding `a` to the one holding `b` (b's index when a's is 0). */
export function periodsBetween(a: number, b: number, level: TimeLevel): number {
  const x = serialToDate(periodStart(a, level)), y = serialToDate(periodStart(b, level));
  if (level === 3) return periodStart(b, level) - periodStart(a, level);
  if (level === 2) return (y.y - x.y) * 12 + (y.m - x.m);
  if (level === 1) return (y.y - x.y) * 4 + Math.floor((y.m - 1) / 3) - Math.floor((x.m - 1) / 3);
  return y.y - x.y;
}

/** Excel's bounds: 1 January of the first date's year to 1 January after the last date's year (end exclusive). */
export function timelineBounds(dates: number[]): { start: number; end: number } | null {
  const ds = dates.filter((d) => Number.isFinite(d) && d >= 1);
  if (!ds.length) return null;
  const lo = serialToDate(Math.min(...ds)), hi = serialToDate(Math.max(...ds));
  return { start: dateToSerial(lo.y, 1, 1), end: dateToSerial(hi.y + 1, 1, 1) };
}

const MON = MONTHS_SHORT_DE;

/** A period's label in the lower row (MRZ, Q1, 5, 2026) and the upper row's label when a new one starts there. */
export function periodLabels(start: number, level: TimeLevel): { lower: string; upper?: string } {
  const d = serialToDate(start);
  if (level === 0) return { lower: String(d.y) };
  if (level === 1) return { lower: `Q${Math.floor((d.m - 1) / 3) + 1}`, upper: d.m === 1 ? String(d.y) : undefined };
  if (level === 2) return { lower: MON[d.m - 1].toUpperCase(), upper: d.m === 1 ? String(d.y) : undefined };
  return { lower: String(d.d), upper: d.d === 1 ? `${MON[d.m - 1].toUpperCase()} ${d.y}` : undefined };
}

/** The upper row's label for the first visible period (Excel names the year or month the view starts in). */
export function firstUpperLabel(start: number, level: TimeLevel): string | undefined {
  const d = serialToDate(start);
  if (level === 0) return undefined;
  if (level === 3) return `${MON[d.m - 1].toUpperCase()} ${d.y}`;
  return String(d.y);
}

/** The selection label as Excel writes it: the shared trailing parts named once. */
export function selectionLabel(sel: { from: number; to: number } | null, level: TimeLevel): string {
  if (!sel) return "Alle Zeiträume";
  const a = serialToDate(sel.from), b = serialToDate(sel.to);
  const q = (m: number) => `Q${Math.floor((m - 1) / 3) + 1}`;
  if (level === 0) return a.y === b.y ? String(a.y) : `${a.y} - ${b.y}`;
  if (level === 1) {
    if (a.y === b.y) return q(a.m) === q(b.m) ? `${q(a.m)} ${a.y}` : `${q(a.m)} - ${q(b.m)} ${a.y}`;
    return `${q(a.m)} ${a.y} - ${q(b.m)} ${b.y}`;
  }
  if (level === 2) {
    if (a.y === b.y) return a.m === b.m ? `${MON[a.m - 1]} ${a.y}` : `${MON[a.m - 1]} - ${MON[b.m - 1]} ${a.y}`;
    return `${MON[a.m - 1]} ${a.y} - ${MON[b.m - 1]} ${b.y}`;
  }
  if (a.y === b.y && a.m === b.m) return a.d === b.d ? `${a.d} ${MON[a.m - 1]} ${a.y}` : `${a.d} - ${b.d} ${MON[a.m - 1]} ${a.y}`;
  if (a.y === b.y) return `${a.d} ${MON[a.m - 1]} - ${b.d} ${MON[b.m - 1]} ${a.y}`;
  return `${a.d} ${MON[a.m - 1]} ${a.y} - ${b.d} ${MON[b.m - 1]} ${b.y}`;
}

// ----- the filter --------------------------------------------------------------------------------

/** A PivotTable's date filter on a field (whole days, both ends included), or none. */
export function pivotDateFilter(p: PivotTable, field: number): { from: number; to: number } | null {
  const f = p.dateFilters?.find((x) => x.field === field);
  return f ? { from: f.from, to: f.to } : null;
}

/** The PivotTable's date filters after a timeline's selection (null: all periods, the filter goes). */
export function pivotDateFiltersAfter(p: PivotTable, field: number, sel: { from: number; to: number } | null): PivotTable["dateFilters"] {
  const list = (p.dateFilters ?? []).filter((x) => x.field !== field);
  if (sel) list.push({ field, from: Math.floor(sel.from), to: Math.floor(sel.to) });
  return list.length ? list : undefined;
}

/** Whether a value passes a date filter (Excel's useWholeDay: the time of day does not count). */
export function passesDateFilter(v: unknown, f: { from: number; to: number }): boolean {
  if (typeof v !== "number") return false;
  const d = Math.floor(v);
  return d >= f.from && d <= f.to;
}
