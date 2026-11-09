// Schnelllayout (Diagrammentwurf > Schnelllayout > Layout n): the elements Excel's and Word's layouts set, recorded in
// Excel (QUICK_LAYOUTS in @raabe/charts, D-115). Applied to a chart model in place. RaabeCalc's feature.ts still has
// the same steps inline (it is being changed in another branch tonight); it switches to this after that merge.

import { QUICK_LAYOUTS, isPie, type ChartModel, type DataLabelOptions } from "@raabe/charts";

/** The recorded layouts that fit the chart's kind (Office shows the same eleven for each family). */
export function layoutsFor(m: ChartModel) {
  const fam: Record<string, string> = { column: "column", columnStacked: "column", bar: "bar", barStacked: "bar", line: "line", lineMarkers: "lineMarkers",
    area: "area", pie: "pie", doughnut: "pie", scatter: "scatter", scatterSmooth: "scatter", scatterLines: "scatter" };
  return QUICK_LAYOUTS[fam[m.kind] ?? ""] ?? [];
}

/** How many layouts the gallery offers for the chart. */
export const layoutCount = (m: ChartModel) => layoutsFor(m).filter(Boolean).length;

/** Applies Layout n (1-based) to the chart; false when the chart has no such layout. */
export function applyQuickLayout(c: ChartModel, n: number): boolean {
  const l = layoutsFor(c)[n - 1];
  if (!l) return false;
  c.title.show = l.title;
  c.legend = l.legend as ChartModel["legend"];
  c.dataLabels = l.labels;
  const f = (c.format ??= {});
  if (l.labels) {
    const d: NonNullable<typeof f.dataLabels> = { ...(f.dataLabels ?? {}) };
    delete d.pos; delete d.showCat; delete d.showPercent;
    const defPos = isPie(c.kind) ? "bestFit" : c.kind.endsWith("Stacked") ? "ctr" : c.kind === "column" || c.kind === "bar" ? "outEnd" : "r";
    if (l.labelPos && l.labelPos !== defPos) d.pos = l.labelPos as DataLabelOptions["pos"];
    if (l.showCat) d.showCat = true;
    if (l.showPct) d.showPercent = true;
    if (Object.keys(d).length) f.dataLabels = d; else delete f.dataLabels;
  }
  c.axisTitles = { ...(l.catTitle ? { cat: c.axisTitles.cat ?? "" } : {}), ...(l.valTitle ? { val: c.axisTitles.val ?? "" } : {}) };
  c.gridlines = !!l.majorGrid;
  if (l.minorGrid) f.minorGrid = { ...(f.minorGrid ?? {}), show: true }; else delete f.minorGrid;
  if (l.catGrid) f.catGrid = { ...(f.catGrid ?? {}), show: true }; else delete f.catGrid;
  if (l.dataTable) c.dataTable = { legendKeys: true }; else delete c.dataTable;
  const defGap = c.kind.endsWith("Stacked") ? 150 : c.kind === "bar" ? 182 : 219;
  if (l.gap !== undefined && (c.kind === "column" || c.kind === "bar" || c.kind.endsWith("Stacked"))) c.gapWidth = l.gap === defGap ? undefined : l.gap;
  if (!Object.keys(f).length) delete c.format;
  return true;
}

/** Farben ändern: colour scheme k (Office's 10-26; series with their own colour keep the scheme's). */
export function applyColors(c: ChartModel, k: number): void {
  c.colors = k === 10 ? undefined : k;
  for (const s of c.series) delete s.color;
}
