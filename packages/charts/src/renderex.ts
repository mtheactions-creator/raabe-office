// Excel 2016's chart types (chartex, D-086): Wasserfall, Trichter, Treemap, Sunburst, Kastengrafik
// and Pareto, drawn as Excel 365 draws them with style 201 (recorded with probe_excel_charttypes.py):
// - Wasserfall: Anstieg, Abfall and Summe in accents 1-3, Verbindungslinien between the columns,
//   the values above (below for falls), the legend at the top;
// - Trichter: centred bars, the categories at the left, the values in white inside;
// - Treemap: rectangles laid out "squarified" per top-level category (accent per category), the
//   parent's name at the top left (overlapping), the leaves' names at the bottom left;
// - Sunburst: rings from the inside out (top level inside), labels along the radius;
// - Kastengrafik: box from the first to the third quartile (exclusive method), median line,
//   mean as an x, whiskers to the last values within 1.5 x IQR, outliers as points;
// - Pareto: columns sorted descending, the cumulative share as a line on a 0-100 % axis at the right.

import { autoScale, maxIntervals, ticks, type AxisScale } from "./scale";
import { LOOK, seriesColor } from "./model";
import { fillPaint } from "./paint";
import { formatGeneral } from "./format";
import { countScale, f, gridAndValueLabels, type Area, type Box, type Ctx, type LegendItem, type Scene } from "./render";

const WHITE = "#FFFFFF";

/** The legend entries of Wasserfall (Anstieg, Abfall, Summe) and Treemap (the top-level categories). */
export function chartExLegend(c: Ctx): LegendItem[] | null {
  const { m } = c;
  if (m.kind === "waterfall") return ["Anstieg", "Abfall", "Summe"].map((text, i) => ({ text, color: seriesColor(i, c.accents), kind: "square" as const }));
  if (m.kind === "treemap") return topLevels(m.series[0]?.levels?.[0] ?? (m.series[0]?.cat?.cache ?? []).map((v) => (v === null ? null : String(v)))).map((text, i) => ({ text, color: seriesColor(i, c.accents), kind: "square" as const }));
  return null;
}

function topLevels(lv: (string | null)[]): string[] {
  const out: string[] = [];
  for (const v of lv) if (v !== null && !out.includes(v)) out.push(v);
  return out;
}

export function drawChartEx(c: Ctx, area: Area): Pick<Scene, "plot" | "valueScale"> {
  switch (c.m.kind) {
    case "waterfall": return drawWaterfall(c, area);
    case "funnel": return drawFunnel(c, area);
    case "treemap": return drawTreemap(c, area);
    case "sunburst": return drawSunburst(c, area);
    case "boxWhisker": return drawBox(c, area);
    default: return drawPareto(c, area);
  }
}

/** The frame of the axis charts (as the histogram): value labels at the left, gridlines, category labels below. */
function frame(c: Ctx, area: Area, cats: string[], lo: number, hi: number, rightLabels?: string[]): { plot: Box; scale: AxisScale; y: (v: number) => number } {
  const top = area.top;
  const bottom = area.bottom - 22.75;
  const scale = c.m.kind === "pareto" ? countScale(hi, maxIntervals(bottom - top, false)) : autoScale(lo, hi, maxIntervals(bottom - top, false));
  const labels = ticks(scale).map((v) => c.fmt(v, c.m.series[0]?.val.format));
  const lw = Math.max(0, ...labels.map((l) => c.measure(l, LOOK.labelSize)));
  const rw = rightLabels ? Math.max(0, ...rightLabels.map((l) => c.measure(l, LOOK.labelSize))) + 6.67 : 0;
  const left = area.left + 5.45 + lw + 6.67, right = area.right - 6.75 - rw;
  const plot = { x: left, y: top, w: right - left, h: bottom - top };
  gridAndValueLabels(c, scale, labels, plot, false, 5.45 + lw + area.left);
  const y = (v: number) => bottom - ((v - scale.min) / (scale.max - scale.min || 1)) * plot.h;
  const bw = cats.length ? plot.w / cats.length : 0;
  cats.forEach((t, i) => c.prims.push({ t: "text", x: left + (i + 0.5) * bw, y: bottom + 14.75, text: t, size: LOOK.labelSize, fill: LOOK.text, anchor: "middle", part: "axis:cat" }));
  return { plot, scale, y };
}

const catStrings = (c: Ctx): string[] => (c.m.series[0]?.cat?.cache ?? c.m.series[0]?.val.cache.map((_, i) => String(i + 1)) ?? []).map((v) => (v === null ? "" : typeof v === "number" ? c.fmt(v, c.m.series[0]?.cat?.format) : v));

function drawWaterfall(c: Ctx, area: Area): Pick<Scene, "plot" | "valueScale"> {
  const s = c.m.series[0];
  const vals = s?.val.cache ?? [];
  const totals = new Set(s?.totals ?? []);
  const bars: { from: number; to: number; kind: 0 | 1 | 2; v: number }[] = [];
  let run = 0;
  vals.forEach((v, i) => {
    const x = v ?? 0;
    if (totals.has(i)) { bars.push({ from: 0, to: x, kind: 2, v: x }); run = x; } else { bars.push({ from: run, to: run + x, kind: x >= 0 ? 0 : 1, v: x }); run += x; }
  });
  const ends = bars.flatMap((b) => [b.from, b.to]);
  const { plot, scale, y } = frame(c, area, catStrings(c), Math.min(0, ...ends), Math.max(0, ...ends));
  const step = bars.length ? plot.w / bars.length : 0, bw = step / 1.5;
  let conn = "";
  bars.forEach((b, i) => {
    const x = plot.x + i * step + (step - bw) / 2;
    const y0 = y(Math.max(b.from, b.to)), y1 = y(Math.min(b.from, b.to));
    c.prims.push({ t: "rect", x, y: y0, w: bw, h: Math.max(0.5, y1 - y0), fill: fillPaint(s?.fill, seriesColor(b.kind, c.accents)), part: `series:0:${i}` });
    if (i + 1 < bars.length) conn += `M${f(x + bw)} ${f(y(b.to))}H${f(x + step)}`;
    if (c.m.dataLabels) c.prims.push({ t: "text", x: x + bw / 2, y: b.v >= 0 ? y0 - 3.4 : y1 + 9.6, text: c.fmt(b.v, s?.val.format), size: LOOK.labelSize, fill: LOOK.text, anchor: "middle", part: "labels:0" });
  });
  if (conn) c.prims.push({ t: "path", d: conn, stroke: "#A6A6A6", sw: 0.75, part: "connectors" });
  return { plot, valueScale: scale };
}

function drawFunnel(c: Ctx, area: Area): Pick<Scene, "plot"> {
  const s = c.m.series[0];
  const vals = (s?.val.cache ?? []).map((v) => Math.max(0, v ?? 0));
  const cats = catStrings(c);
  const lw = Math.max(0, ...cats.map((t) => c.measure(t, LOOK.labelSize)));
  const left = area.left + 6.75 + lw + 8, right = area.right - 6.75;
  const plot = { x: left, y: area.top, w: Math.max(10, right - left), h: Math.max(10, area.bottom - 9.4 - area.top) };
  const step = vals.length ? plot.h / vals.length : 0, bh = step / 1.06;
  const max = Math.max(0, ...vals) || 1, cx = plot.x + plot.w / 2;
  vals.forEach((v, i) => {
    const w = (v / max) * plot.w, y = plot.y + i * step + (step - bh) / 2;
    c.prims.push({ t: "rect", x: cx - w / 2, y, w, h: bh, fill: fillPaint(s?.fill, c.color(0)), part: `series:0:${i}` });
    c.prims.push({ t: "text", x: plot.x - 8, y: y + bh / 2 + 3.2, text: cats[i] ?? "", size: LOOK.labelSize, fill: LOOK.text, anchor: "end", part: "axis:cat" });
    if (c.m.dataLabels) c.prims.push({ t: "text", x: cx, y: y + bh / 2 + 3.2, text: c.fmt(s?.val.cache[i] ?? 0, s?.val.format), size: LOOK.labelSize, fill: WHITE, anchor: "middle", part: "labels:0" });
  });
  return { plot };
}

/** The leaves of a Treemap or Sunburst grouped by their path of levels. */
interface Node { name: string; value: number; children: Node[]; top: number; leaf: number }

function hierarchy(c: Ctx): Node {
  const s = c.m.series[0];
  const levels = s?.levels ?? [(s?.cat?.cache ?? []).map((v) => (v === null ? null : String(v)))];
  const vals = s?.val.cache ?? [];
  const root: Node = { name: "", value: 0, children: [], top: -1, leaf: -1 };
  const tops = topLevels(levels[0] ?? []);
  vals.forEach((v, i) => {
    if (v === null || !(v > 0)) return;
    let node = root;
    levels.forEach((lv, d) => {
      const name = lv[i];
      if (name === null || name === undefined) return;
      let ch = node.children.find((x) => x.name === name);
      if (!ch) { ch = { name, value: 0, children: [], top: tops.indexOf(levels[0][i] ?? ""), leaf: d === levels.length - 1 ? i : -1 }; node.children.push(ch); }
      node.value += 0;
      node = ch;
    });
    // Add the value along the path.
    let n2 = root;
    root.value += v;
    for (const lv of levels) { const ch = n2.children.find((x) => x.name === lv[i]); if (!ch) break; ch.value += v; n2 = ch; }
  });
  return root;
}

/** Squarified treemap layout of `nodes` (largest first) in a rectangle. */
function squarify(nodes: Node[], r: Box): { node: Node; box: Box }[] {
  const items = [...nodes].sort((a, b) => b.value - a.value);
  const out: { node: Node; box: Box }[] = [];
  let rect = { ...r };
  let rest = items;
  const total = items.reduce((a, n) => a + n.value, 0) || 1;
  const scale = (r.w * r.h) / total;
  while (rest.length) {
    const short = Math.min(rect.w, rect.h);
    let row: Node[] = [], best = Infinity;
    for (const n of rest) {
      const next = [...row, n];
      const s = next.reduce((a, x) => a + x.value * scale, 0);
      const worst = Math.max(...next.map((x) => Math.max((short * short * x.value * scale) / (s * s), (s * s) / (short * short * x.value * scale))));
      if (worst > best) break;
      best = worst;
      row = next;
    }
    const sum = row.reduce((a, x) => a + x.value * scale, 0);
    const thick = sum / short;
    let off = 0;
    for (const n of row) {
      const len = (n.value * scale) / thick;
      out.push({ node: n, box: rect.w >= rect.h ? { x: rect.x, y: rect.y + off, w: thick, h: len } : { x: rect.x + off, y: rect.y, w: len, h: thick } });
      off += len;
    }
    rect = rect.w >= rect.h ? { x: rect.x + thick, y: rect.y, w: rect.w - thick, h: rect.h } : { x: rect.x, y: rect.y + thick, w: rect.w, h: rect.h - thick };
    rest = rest.slice(row.length);
  }
  return out;
}

function drawTreemap(c: Ctx, area: Area): Pick<Scene, "plot"> {
  const plot = { x: area.left + 1.4, y: area.top, w: area.right - area.left - 2.8, h: area.bottom - 9.4 - area.top };
  const root = hierarchy(c);
  const s = c.m.series[0];
  const paint = (n: Node, b: Box, color: string, depth: number) => {
    if (!n.children.length) {
      c.prims.push({ t: "rect", ...b, fill: color, stroke: WHITE, sw: 1.5, part: `series:0:${n.leaf}` });
      if (c.m.dataLabels && b.w > 12 && b.h > 12) c.prims.push({ t: "text", x: b.x + 3.5, y: b.y + b.h - 4, text: n.name, size: LOOK.labelSize, fill: WHITE, anchor: "start", part: "labels:0" });
      return;
    }
    for (const ch of squarify(n.children, b)) paint(ch.node, ch.box, color, depth + 1);
    if (depth >= 0 && c.m.dataLabels) c.prims.push({ t: "text", x: b.x + 3.5, y: b.y + 11, text: n.name, size: LOOK.labelSize, fill: WHITE, anchor: "start", part: "labels:0" });
  };
  for (const t of squarify(root.children, plot)) paint(t.node, t.box, fillPaint(s?.fill, seriesColor(Math.max(0, t.node.top), c.accents)) ?? WHITE, 0);
  return { plot };
}

function drawSunburst(c: Ctx, area: Area): Pick<Scene, "plot"> {
  const root = hierarchy(c);
  const depth = (n: Node): number => (n.children.length ? 1 + Math.max(...n.children.map(depth)) : 0);
  const rings = Math.max(1, depth(root));
  const h = area.bottom - 9.4 - (area.top - 6);
  const R = Math.max(10, Math.min(h, area.right - area.left) / 2);
  const cx = (area.left + area.right) / 2, cy = area.top - 6 + h / 2;
  const hole = R * 0.27, ringW = (R - hole) / rings;
  const s = c.m.series[0];
  const pt = (r: number, a: number) => `${f(cx + r * Math.sin(a))} ${f(cy - r * Math.cos(a))}`;
  const arc = (r0: number, r1: number, a0: number, a1: number) => {
    const large = a1 - a0 > Math.PI ? 1 : 0;
    if (a1 - a0 >= Math.PI * 2 - 1e-9) return `M${pt(r1, 0)}A${f(r1)} ${f(r1)} 0 1 1 ${pt(r1, Math.PI)}A${f(r1)} ${f(r1)} 0 1 1 ${pt(r1, 0)}ZM${pt(r0, 0)}A${f(r0)} ${f(r0)} 0 1 0 ${pt(r0, Math.PI)}A${f(r0)} ${f(r0)} 0 1 0 ${pt(r0, 0)}Z`;
    return `M${pt(r1, a0)}A${f(r1)} ${f(r1)} 0 ${large} 1 ${pt(r1, a1)}L${pt(r0, a1)}A${f(r0)} ${f(r0)} 0 ${large} 0 ${pt(r0, a0)}Z`;
  };
  const walk = (n: Node, d: number, a0: number, a1: number, color: string | null) => {
    let a = a0;
    for (const ch of n.children) {
      const b = a + ((a1 - a0) * ch.value) / (n.value || 1);
      const col = color ?? fillPaint(s?.fill, seriesColor(Math.max(0, ch.top), c.accents)) ?? WHITE;
      const r0 = hole + d * ringW, r1 = r0 + ringW;
      c.prims.push({ t: "path", d: arc(r0, r1, a, b), fill: col, stroke: WHITE, sw: 1.5, part: `series:0:${ch.leaf}` });
      if (c.m.dataLabels && b - a > 0.08) {
        // Labels along the radius, upright.
        const mid = (a + b) / 2, rm = (r0 + r1) / 2;
        let deg = (mid * 180) / Math.PI - 90;
        if (deg > 90) deg -= 180;
        const x = cx + rm * Math.sin(mid), y = cy - rm * Math.cos(mid);
        c.prims.push({ t: "text", x, y: y + 3.2, text: ch.name, size: LOOK.labelSize, fill: WHITE, anchor: "middle", rotate: Math.abs(deg) > 1 ? deg : undefined, part: "labels:0" });
      }
      walk(ch, d + 1, a, b, col);
      a = b;
    }
  };
  walk(root, 0, 0, Math.PI * 2, null);
  return { plot: { x: cx - R, y: cy - R, w: 2 * R, h: 2 * R } };
}

/** A quantile with Excel's exclusive method (QUARTILE.EXKL): position q x (n + 1). */
export function quantileExc(sorted: number[], q: number): number {
  const n = sorted.length;
  if (!n) return 0;
  const pos = q * (n + 1);
  if (pos <= 1) return sorted[0];
  if (pos >= n) return sorted[n - 1];
  const k = Math.floor(pos);
  return sorted[k - 1] + (pos - k) * (sorted[k] - sorted[k - 1]);
}

function drawBox(c: Ctx, area: Area): Pick<Scene, "plot" | "valueScale"> {
  const { m } = c;
  const all = m.series.flatMap((s) => s.val.cache.filter((v): v is number => v !== null && isFinite(v)));
  const { plot, scale, y } = frame(c, area, ["1"], Math.min(0, ...all), Math.max(0, ...all));
  const k = Math.max(1, m.series.length);
  const cluster = plot.w / 2, slot = cluster / k, bw = slot * 0.8;
  m.series.forEach((s, si) => {
    const v = s.val.cache.filter((x): x is number => x !== null && isFinite(x)).sort((a, b) => a - b);
    if (!v.length) return;
    const q1 = quantileExc(v, 0.25), med = quantileExc(v, 0.5), q3 = quantileExc(v, 0.75), iqr = q3 - q1;
    const inside = v.filter((x) => x >= q1 - 1.5 * iqr && x <= q3 + 1.5 * iqr);
    const lo = Math.min(...inside), hi = Math.max(...inside);
    const mean = v.reduce((a, b) => a + b, 0) / v.length;
    const color = fillPaint(s.fill, c.color(si)) ?? c.color(si);
    const x0 = plot.x + plot.w / 2 - cluster / 2 + si * slot + (slot - bw) / 2, xm = x0 + bw / 2;
    const part = `series:${si}`;
    c.prims.push({ t: "path", d: `M${f(xm)} ${f(y(hi))}V${f(y(q3))}M${f(xm)} ${f(y(q1))}V${f(y(lo))}M${f(xm - bw / 4)} ${f(y(hi))}H${f(xm + bw / 4)}M${f(xm - bw / 4)} ${f(y(lo))}H${f(xm + bw / 4)}`, stroke: "#7F7F7F", sw: 0.75, part });
    c.prims.push({ t: "rect", x: x0, y: y(q3), w: bw, h: Math.max(0.5, y(q1) - y(q3)), fill: color, stroke: color, sw: 0.75, part });
    c.prims.push({ t: "path", d: `M${f(x0)} ${f(y(med))}H${f(x0 + bw)}`, stroke: WHITE, sw: 0.75, part });
    const r = 2.5, my = y(mean);
    c.prims.push({ t: "path", d: `M${f(xm - r)} ${f(my - r)}L${f(xm + r)} ${f(my + r)}M${f(xm + r)} ${f(my - r)}L${f(xm - r)} ${f(my + r)}`, stroke: WHITE, sw: 0.75, part });
    for (const o of v.filter((x) => x < lo || x > hi)) c.prims.push({ t: "circle", cx: xm, cy: y(o), r: 1.75, fill: color, part });
    if (m.dataLabels) c.prims.push({ t: "text", x: x0 + bw + 3, y: y(med) + 3.2, text: formatGeneral(med), size: LOOK.labelSize, fill: LOOK.text, anchor: "start", part: `labels:${si}` });
  });
  return { plot, valueScale: scale };
}

function drawPareto(c: Ctx, area: Area): Pick<Scene, "plot" | "valueScale"> {
  const s = c.m.series[0];
  const cats = catStrings(c);
  const items = (s?.val.cache ?? []).map((v, i) => ({ v: Math.max(0, v ?? 0), cat: cats[i] ?? "" }));
  // Same categories add up (aggregation), then the columns run from the largest down.
  const agg: { v: number; cat: string }[] = [];
  for (const it of items) { const a = agg.find((x) => x.cat === it.cat && it.cat !== ""); if (a) a.v += it.v; else agg.push({ ...it }); }
  agg.sort((a, b) => b.v - a.v);
  const pct = ["0%", "10%", "20%", "30%", "40%", "50%", "60%", "70%", "80%", "90%", "100%"];
  const total = agg.reduce((a, x) => a + x.v, 0) || 1;
  const { plot, scale, y } = frame(c, area, agg.map((x) => x.cat), 0, Math.max(0, ...agg.map((x) => x.v)), pct);
  const step = agg.length ? plot.w / agg.length : 0;
  agg.forEach((a, i) => {
    if (a.v > 0) c.prims.push({ t: "rect", x: plot.x + i * step + 0.25, y: y(a.v), w: step - 0.75, h: plot.y + plot.h - y(a.v), fill: fillPaint(s?.fill, c.color(0)), part: `series:0:${i}` });
  });
  // The percentage axis at the right and the cumulative line.
  const yp = (p: number) => plot.y + plot.h - p * plot.h;
  pct.forEach((t, i) => c.prims.push({ t: "text", x: plot.x + plot.w + 6.67, y: yp(i / 10) + 2.5, text: t, size: LOOK.labelSize, fill: LOOK.text, anchor: "start", part: "axis:val2" }));
  let cum = 0, d = "";
  agg.forEach((a, i) => { cum += a.v; d += `${i ? "L" : "M"}${f(plot.x + (i + 0.5) * step)} ${f(yp(cum / total))}`; });
  if (d) c.prims.push({ t: "path", d, stroke: seriesColor(1, c.accents), sw: 2.25, cap: "round", part: "series:1" });
  return { plot, valueScale: scale };
}
