// Netz (radar, with markers, filled) and Kreis aus Kreis / Balken aus Kreis (D-086), drawn as
// Excel's default look: the radar's value gridlines as polygons with spokes to every category, its
// value labels along the first spoke, the categories outside; the "of pie" charts as the first pie
// with the last points gathered in one slice (facing right) and those points again in a second pie
// or a stacked bar, joined by Verbindungslinien.

import { ticks } from "./scale";
import { tightScale } from "./render3d";
import { LOOK } from "./model";
import { fillPaint, linePaint } from "./paint";
import { kindInfo } from "./kinds";
import { catText, drawMarker, f, label, paletteColor, slice, type Area, type Ctx, type Scene } from "./render";

/** Netz: categories clockwise from the top. */
export function drawRadar(c: Ctx, area: Area): Pick<Scene, "plot" | "valueScale"> {
  const { m } = c;
  const n = Math.max(1, ...m.series.map((s) => s.val.cache.length));
  const cats = Array.from({ length: n }, (_, i) => catText(c, m.series[0], i));
  const catW = Math.max(0, ...cats.map((t) => c.measure(t, LOOK.labelSize)));
  const room = Math.max(11, catW * 0.6 + 6);
  const R = Math.max(10, Math.min((area.right - area.left) / 2 - room - 6, (area.bottom - area.top) / 2 - 14));
  const cx = (area.left + area.right) / 2, cy = (area.top + area.bottom) / 2 - 2;
  const vals = m.series.flatMap((s) => s.val.cache.filter((v): v is number => v !== null && isFinite(v)));
  const lo = Math.min(0, ...vals), hi = Math.max(0, ...vals);
  const scale = tightScale(lo, hi, undefined, Math.max(1, Math.min(10, (R * 1.2 - 19.7) / 10)));
  const span = scale.max - scale.min || 1;
  const ang = (i: number) => (i / n) * Math.PI * 2;
  const P = (i: number, v: number) => {
    const r = ((v - scale.min) / span) * R;
    return [cx + r * Math.sin(ang(i)), cy - r * Math.cos(ang(i))] as const;
  };
  const poly = (v: number) => Array.from({ length: n }, (_, i) => P(i, v)).map(([x, y], i) => `${i ? "L" : "M"}${f(x)} ${f(y)}`).join("") + "Z";
  const t = ticks(scale);
  if (m.gridlines) for (const v of t) if (v !== scale.min) c.prims.push({ t: "path", d: poly(v), stroke: LOOK.gridline, sw: 0.75, part: "grid" });
  // Spokes (the category axis lines).
  let spokes = "";
  for (let i = 0; i < n; i++) { const [x, y] = P(i, scale.max); spokes += `M${f(cx)} ${f(cy)}L${f(x)} ${f(y)}`; }
  c.prims.push({ t: "path", d: spokes, stroke: LOOK.gridline, sw: 0.75, part: "axes" });
  t.forEach((v) => {
    const [, y] = P(0, v);
    c.prims.push({ t: "text", x: cx - 3, y: y + 3, text: c.fmt(v, m.series[0]?.val.format), size: LOOK.labelSize, fill: LOOK.text, anchor: "end", part: "axis:val" });
  });
  cats.forEach((text, i) => {
    const a = ang(i);
    const x = cx + (R + 5) * Math.sin(a), y = cy - (R + 5) * Math.cos(a);
    const anchor = Math.abs(Math.sin(a)) < 0.2 ? "middle" : Math.sin(a) > 0 ? "start" : "end";
    c.prims.push({ t: "text", x, y: y + 3.2 - Math.cos(a) * 4, text, size: LOOK.labelSize, fill: LOOK.text, anchor, part: "axis:cat" });
  });
  const filled = m.kind === "radarFilled", markers = !!kindInfo(m.kind).markers;
  m.series.forEach((s, si) => {
    const pts = s.val.cache.map((v, i) => (v === null || !isFinite(v) ? null : P(i, v)));
    const d = pts.filter((p): p is readonly [number, number] => !!p).map(([x, y], i) => `${i ? "L" : "M"}${f(x)} ${f(y)}`).join("") + "Z";
    if (filled) {
      const ln = linePaint(s.line, null);
      c.prims.push({ t: "path", d, fill: fillPaint(s.fill, c.color(si)), stroke: ln?.stroke, sw: ln?.sw, part: `series:${si}` });
    } else {
      const ln = linePaint(s.line, { color: c.color(si), width: 2.25 });
      if (ln) c.prims.push({ t: "path", d, ...ln, cap: "round", part: `series:${si}` });
      pts.forEach((p, i) => { if (p) drawMarker(c, s, si, i, p[0], p[1], markers ? "circle" : "none"); });
    }
    if (m.dataLabels) pts.forEach((p, i) => { if (p) label(c, s, si, i, s.val.cache[i]!, p[0] + 5, p[1] + 3.2, "start"); });
  });
  return { plot: { x: cx - R, y: cy - R, w: 2 * R, h: 2 * R }, valueScale: scale };
}

/** Kreis aus Kreis / Balken aus Kreis. */
export function drawOfPie(c: Ctx, area: Area): Pick<Scene, "plot"> {
  const { m } = c;
  const s = m.series[0];
  const top = (m.title.show ? 37.6 : area.top) + 2, bottom = area.bottom - 10.5;
  const left = area.left + 10.5, right = area.right - 10.5;
  const plot = { x: left, y: top, w: right - left, h: bottom - top };
  if (!s) return { plot };
  const vals = s.val.cache.map((v) => (v !== null && v > 0 ? v : 0));
  const n = vals.length;
  const split = Math.max(0, Math.min(n - 1, m.splitCount ?? Math.min(3, n - 1)));
  const main = vals.slice(0, n - split), second = vals.slice(n - split);
  const other = second.reduce((a, b) => a + b, 0);
  const pieVals = [...main, other];
  const sum = pieVals.reduce((a, b) => a + b, 0);
  if (sum <= 0) return { plot };
  const bar = m.kind === "barOfPie";
  const size = (m.secondSize ?? 75) / 100, gap = (m.gapWidth ?? 100) / 100;
  // Widths (recorded): the pie (D), the gap (pie: half of D x Abstandsbreite, bar: a quarter), the second plot (bar: half its height).
  const g = gap * (bar ? 0.25 : 0.5);
  const D = Math.max(10, Math.min(plot.h, plot.w / (1 + g + (bar ? size * 0.49 : size))));
  const d2 = D * size, w2 = bar ? d2 * 0.49 : d2;
  const total = D + g * D + w2;
  const x0 = plot.x + (plot.w - total) / 2;
  const cx = x0 + D / 2, cy = plot.y + plot.h / 2, r = D / 2;
  // The "Sonstige" slice faces right: it is centred on 90°.
  const otherAngle = (other / sum) * Math.PI * 2;
  let a0 = Math.PI / 2 + otherAngle / 2;
  const ln = linePaint(s.line, { color: "#FFFFFF", width: 1.5 });
  const colorOf = (i: number) => fillPaint(s.fill, paletteColor(m, i, n + 1, c.accents)) ?? "#808080";
  let otherEdges: [number, number] = [0, 0];
  pieVals.forEach((v, k) => {
    if (v <= 0) return;
    const a1 = a0 + (v / sum) * Math.PI * 2;
    const idx = k === main.length ? n : k;
    c.prims.push({ t: "path", d: slice(cx, cy, r, 0, a0, a1), fill: colorOf(idx), stroke: ln?.stroke, sw: ln?.sw, part: `series:0:${idx}` });
    if (k === main.length) otherEdges = [a0, a1];
    if (m.dataLabels) {
      const mid = (a0 + a1) / 2;
      label(c, s, 0, idx, k === main.length ? other : s.val.cache[k] ?? 0, cx + r * 0.7 * Math.sin(mid), cy - r * 0.7 * Math.cos(mid) + 3.2, "middle", v / sum);
    }
    a0 = a1;
  });
  const sx = x0 + D + g * D;
  const line = { stroke: "#A6A6A6", sw: 0.75, part: "serLines" };
  const edge = (a: number) => [cx + r * Math.sin(a), cy - r * Math.cos(a)] as const;
  const [e0, e1] = [edge(otherEdges[0]), edge(otherEdges[1])];
  if (bar) {
    const h = d2, by = cy - h / 2;
    let y = by;
    second.forEach((v, k) => {
      if (other <= 0) return;
      const hh = (v / other) * h;
      c.prims.push({ t: "rect", x: sx, y, w: w2, h: hh, fill: colorOf(main.length + k), stroke: ln?.stroke, sw: ln?.sw, part: `series:0:${main.length + k}` });
      if (m.dataLabels) label(c, s, 0, main.length + k, v, sx + w2 / 2, y + hh / 2 + 3.2, "middle");
      y += hh;
    });
    // Verbindungslinien from the slice's edges to the bar's corners.
    c.prims.push({ t: "path", d: `M${f(e0[0])} ${f(e0[1])}L${f(sx)} ${f(by)}M${f(e1[0])} ${f(e1[1])}L${f(sx)} ${f(by + h)}`, ...line });
  } else {
    const r2 = d2 / 2, cx2 = sx + r2;
    const sum2 = second.reduce((a, b) => a + b, 0);
    // The second pie's points start at the left, as Excel lays them out.
    let b0 = Math.PI * 1.5 - ((second[0] ?? 0) / Math.max(1e-9, sum2)) * Math.PI;
    second.forEach((v, k) => {
      if (v <= 0 || sum2 <= 0) return;
      const b1 = b0 + (v / sum2) * Math.PI * 2;
      c.prims.push({ t: "path", d: slice(cx2, cy, r2, 0, b0, b1), fill: colorOf(main.length + k), stroke: ln?.stroke, sw: ln?.sw, part: `series:0:${main.length + k}` });
      if (m.dataLabels) { const mid = (b0 + b1) / 2; label(c, s, 0, main.length + k, v, cx2 + r2 * 0.7 * Math.sin(mid), cy - r2 * 0.7 * Math.cos(mid) + 3.2, "middle"); }
      b0 = b1;
    });
    // Verbindungslinien: tangents from the slice's edges to the top and bottom of the second pie.
    c.prims.push({ t: "path", d: `M${f(e0[0])} ${f(e0[1])}L${f(cx2)} ${f(cy - r2)}M${f(e1[0])} ${f(e1[1])}L${f(cx2)} ${f(cy + r2)}`, ...line });
  }
  return { plot: { x: x0, y: cy - r, w: total, h: D } };
}
