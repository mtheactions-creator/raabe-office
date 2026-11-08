// Datentabelle (Diagrammelemente > Datentabelle, D-115): below the plot area a table with a row per series (its
// legend key when "Mit Legendensymbolen") and a column per category; the category names head the columns in place
// of the axis labels, as Excel draws it. The plot gives up the table's height; the row names stand left of it.

import { LOOK, seriesColor, type ChartModel } from "./model";
import type { Area, Box, Ctx, Scene } from "./render";

const ROW = 13.5;

/** Draws the chart's axes with a data table under them (`draw` lays out the plot in the area it gets). */
export function withDataTable(c: Ctx, area: Area, draw: (c: Ctx, area: Area) => Pick<Scene, "plot" | "valueScale" | "xScale">,
  helpers: { catText: (i: number) => string; seriesName: (i: number) => string }): Pick<Scene, "plot" | "valueScale" | "xScale"> {
  const m = c.m;
  const n = Math.max(0, ...m.series.map((s) => s.val.cache.length));
  const rows = m.series.length;
  const keys = !!m.dataTable?.legendKeys;
  const size = LOOK.labelSize;
  const nameW = Math.max(0, ...m.series.map((_, i) => c.measure(helpers.seriesName(i), size))) + (keys ? 12 : 0) + 8;
  const tableH = ROW * (rows + 1);
  // The category labels are the table's head: the axis draws none.
  const inner: ChartModel = { ...m, format: { ...(m.format ?? {}), catAxis: { ...(m.format?.catAxis ?? {}), labelPos: "none" } } };
  const sub: Ctx = { ...c, m: inner };
  const a: Area = { ...area, bottom: area.bottom - tableH + 4, left: Math.max(area.left, area.left + nameW - 22) };
  const scene = draw(sub, a);
  c.prims = sub.prims;
  const plot: Box = scene.plot;
  const top = plot.y + plot.h;
  const left = plot.x - nameW;
  const colW = n ? plot.w / n : plot.w;
  const line = { stroke: LOOK.gridline, sw: 0.75 };
  const f = (v: number) => Math.round(v * 100) / 100;
  // Borders: outline, rows and columns (Excel's showHorzBorder, showVertBorder, showOutline).
  let d = `M${f(plot.x)} ${f(top)}V${f(top + tableH)}M${f(plot.x + plot.w)} ${f(top)}V${f(top + tableH)}`;
  for (let r = 1; r <= rows + 1; r++) d += `M${f(r === 1 ? plot.x : left)} ${f(top + r * ROW)}H${f(plot.x + plot.w)}`;
  d += `M${f(left)} ${f(top + ROW)}V${f(top + tableH)}`;
  for (let i = 1; i < n; i++) d += `M${f(plot.x + i * colW)} ${f(top)}V${f(top + tableH)}`;
  c.prims.push({ t: "path", d, fill: "none", ...line, part: "dataTable" });
  for (let i = 0; i < n; i++) {
    c.prims.push({ t: "text", x: plot.x + (i + 0.5) * colW, y: top + ROW - 4, text: helpers.catText(i), size, fill: LOOK.text, anchor: "middle", part: "dataTable" });
  }
  m.series.forEach((s, si) => {
    const y = top + (si + 2) * ROW - 4;
    let x = left + 4;
    if (keys) {
      c.prims.push({ t: "rect", x, y: y - 6, w: 6, h: 6, fill: c.color(si) ?? seriesColor(si, c.accents), part: "dataTable" });
      x += 10;
    }
    c.prims.push({ t: "text", x, y, text: helpers.seriesName(si), size, fill: LOOK.text, anchor: "start", part: "dataTable" });
    s.val.cache.forEach((v, i) => {
      if (v === null) return;
      c.prims.push({ t: "text", x: plot.x + (i + 0.5) * colW, y, text: c.fmt(v, s.val.format), size, fill: LOOK.text, anchor: "middle", part: "dataTable" });
    });
  });
  return scene;
}
