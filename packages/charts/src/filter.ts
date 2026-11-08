// Diagrammfilter (D-116): series and categories a chart keeps but does not draw. Excel keeps a filtered series in
// the file (c15:filtered…Series) and a filtered category in the series' full range (c15:fullRef); the chart shows
// what is left, each series in the colour it has unfiltered. Drawing works on this view of the chart.

import type { ChartModel, ChartSeries } from "./model";
import { paletteColor } from "./render";

/** Whether anything of the chart is filtered. */
export function isFiltered(m: ChartModel): boolean {
  return m.series.some((s) => s.filtered) || !!m.hiddenCats?.length;
}

/**
 * The chart as drawn: filtered series left out (the others keep their colours), hidden categories taken out of
 * every series' categories and values. `index[i]` is the series' index in the full chart.
 */
export function visibleModel(m: ChartModel, accents?: string[]): { model: ChartModel; index: number[] } {
  if (!isFiltered(m)) return { model: m, index: m.series.map((_, i) => i) };
  const hidden = new Set(m.hiddenCats ?? []);
  const keep = <T>(a: T[] | undefined): T[] | undefined => (a ? a.filter((_, i) => !hidden.has(i)) : a);
  const index: number[] = [];
  const series: ChartSeries[] = [];
  m.series.forEach((s, i) => {
    if (s.filtered) return;
    index.push(i);
    const own = s.color ?? (accents ? paletteColor(m, i, m.series.length, accents) : undefined);
    const v: ChartSeries = { ...s, val: { ...s.val, cache: keep(s.val.cache)! } };
    if (own && !s.color) v.color = own;
    if (s.cat) v.cat = { ...s.cat, cache: keep(s.cat.cache)! };
    if (s.size) v.size = { ...s.size, cache: keep(s.size.cache)! };
    if (s.levels) v.levels = s.levels.map((l) => keep(l)!);
    delete v.filtered;
    series.push(v);
  });
  const { hiddenCats: _h, ...rest } = m;
  void _h;
  return { model: { ...rest, series }, index };
}
