// Gliederung (Gruppieren) as Excel shows it: rows or columns with an outline level, a bar beside
// the headers with one lane per level, a bracket along each group and a −/+ box at its summary
// row (the row after the group; the column after it), and level buttons 1, 2, … that show the
// outline down to that level. Collapsing hides the group and marks the summary row "collapsed",
// as the .xlsx file keeps it.

import type { Sheet } from "@raabe/sheet-engine";

export type Axis = "rows" | "cols";
/** Width of one lane of the bar at 100 % (CSS px). */
export const LANE = 16;

type Info = { outline?: number; hidden?: boolean; collapsed?: boolean };
const infos = (sheet: Sheet, axis: Axis) => (axis === "rows" ? sheet.rows : sheet.cols) as Map<number, Info>;

export function maxLevel(sheet: Sheet, axis: Axis): number {
  let m = 0;
  for (const i of infos(sheet, axis).values()) if ((i.outline ?? 0) > m) m = i.outline!;
  return m;
}

/** The bar's width (rows) or height (columns) at a zoom: one lane per level plus one, none without groups. */
export function barSize(sheet: Sheet, axis: Axis, zoom: number): number {
  const m = maxLevel(sheet, axis);
  return m ? Math.round((m + 1) * LANE * zoom + 2 * zoom) : 0;
}

/** The groups at one level: runs of consecutive rows (columns) with at least that level. */
export function groups(sheet: Sheet, axis: Axis, level: number): [number, number][] {
  const idx = [...infos(sheet, axis)].filter(([, i]) => (i.outline ?? 0) >= level).map(([k]) => k).sort((a, b) => a - b);
  const out: [number, number][] = [];
  for (const k of idx) {
    const last = out.at(-1);
    if (last && last[1] === k - 1) last[1] = k;
    else out.push([k, k]);
  }
  return out;
}

const set = (sheet: Sheet, axis: Axis, k: number, change: Partial<Info>) => {
  const m = infos(sheet, axis);
  const next = { ...(m.get(k) ?? {}), ...change };
  for (const key of Object.keys(next) as (keyof Info)[]) if (next[key] === undefined || next[key] === false || next[key] === 0) delete next[key];
  if (Object.keys(next).length) m.set(k, next);
  else m.delete(k);
};

/** −/+ on a group's box: hide the group, or show it again (groups inside it that are collapsed stay so). */
export function toggle(sheet: Sheet, axis: Axis, group: [number, number], level: number): void {
  const summary = group[1] + 1;
  const collapsed = !!infos(sheet, axis).get(summary)?.collapsed;
  if (!collapsed) {
    for (let k = group[0]; k <= group[1]; k++) set(sheet, axis, k, { hidden: true });
    set(sheet, axis, summary, { collapsed: true });
    return;
  }
  set(sheet, axis, summary, { collapsed: undefined });
  for (let k = group[0]; k <= group[1]; k++) set(sheet, axis, k, { hidden: undefined });
  // Inner groups that were collapsed hide their rows again.
  for (let l = level + 1; l <= 7; l++) {
    for (const g of groups(sheet, axis, l)) {
      if (g[0] < group[0] || g[1] > group[1]) continue;
      if (infos(sheet, axis).get(g[1] + 1)?.collapsed) for (let k = g[0]; k <= g[1]; k++) set(sheet, axis, k, { hidden: true });
    }
  }
}

/** Level button n: everything above level n shown, the groups of level n and deeper collapsed. */
export function showLevel(sheet: Sheet, axis: Axis, n: number): void {
  const max = maxLevel(sheet, axis);
  for (const [k, i] of [...infos(sheet, axis)]) {
    if (i.outline) set(sheet, axis, k, { hidden: i.outline >= n ? true : undefined });
  }
  for (let l = 1; l <= max; l++) for (const g of groups(sheet, axis, l)) set(sheet, axis, g[1] + 1, { collapsed: l >= n ? true : undefined });
}

/** Gruppieren (+1) / Gruppierung aufheben (−1) for rows or columns a..b (levels 0 to 7). */
export function regroup(sheet: Sheet, axis: Axis, a: number, b: number, delta: 1 | -1): boolean {
  let changed = false;
  for (let k = a; k <= b; k++) {
    const cur = infos(sheet, axis).get(k)?.outline ?? 0;
    const next = Math.max(0, Math.min(7, cur + delta));
    if (next === cur) continue;
    set(sheet, axis, k, { outline: next || undefined, hidden: next ? infos(sheet, axis).get(k)?.hidden : undefined });
    changed = true;
  }
  return changed;
}
