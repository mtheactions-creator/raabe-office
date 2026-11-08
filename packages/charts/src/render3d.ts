// 3-D charts (D-086): 3D-Säulen and 3D-Balken (clustered, stacked, 100 %, series in depth),
// 3D-Linie, 3D-Fläche (also stacked and 100 %), 3D-Oberfläche (filled and Drahtmodell), Kontur and
// 3D-Kreis. Excel renders these with its own 3-D engine; here the chart's world box (categories
// across, values up, series in depth) is mapped to the page by a linear projection: an oblique one
// for "Rechtwinklige Achsen" (rAngAx, Excel's 3-D columns and bars: the front stays flat, depth goes
// up and to the right by the view's rotY/rotX) and a rotation by rotY and rotX otherwise (lines,
// areas, surfaces). Faces are drawn back to front and shaded per orientation as Excel's default
// light does (sides darker than fronts). Not matched: Excel's perspective (the projection is
// parallel), its exact lighting and the label placement along tilted axes.

import { autoScale, maxIntervals, ticks, type AxisScale } from "./scale";
import { LOOK, lum, seriesColor, type ChartModel } from "./model";
import { fillPaint, fixedScale } from "./paint";
import { kindInfo, view3DOf } from "./kinds";
import {
  LABEL_GAP, PAD, catText, f, label, paletteColor, surfaceBands,
  type Box, type Ctx, type Scene,
} from "./render";

type Area = { left: number; right: number; top: number; bottom: number };
type V3 = [number, number, number];

interface Proj { x(p: V3): number; y(p: V3): number; depth(p: V3): number }

/** A linear projection of the world box [0,1] x [0,hy] x [0,dz] fitted into `plot`. */
function projection(m: ChartModel, plot: Box, hy: number, dz: number): Proj {
  const v = view3DOf(m) ?? { rotX: 15, rotY: 20, rAngAx: true };
  const rx = (v.rotX * Math.PI) / 180, ry = (v.rotY * Math.PI) / 180;
  let raw: (p: V3) => [number, number, number];
  if (v.rAngAx) {
    // Oblique: the front plane as it is, depth up and to the right.
    raw = ([x, y, z]) => [x + z * Math.sin(ry), -y - z * Math.sin(rx), z];
  } else {
    raw = ([x, y, z]) => {
      const X = x * Math.cos(ry) + z * Math.sin(ry), Z = -x * Math.sin(ry) + z * Math.cos(ry);
      return [X, -(y * Math.cos(rx) + Z * Math.sin(rx)), Z * Math.cos(rx) - y * Math.sin(rx)];
    };
  }
  const corners: V3[] = [];
  for (const x of [0, 1]) for (const y of [0, hy]) for (const z of [0, dz]) corners.push([x, y, z]);
  const pr = corners.map(raw);
  const x0 = Math.min(...pr.map((p) => p[0])), x1 = Math.max(...pr.map((p) => p[0]));
  const y0 = Math.min(...pr.map((p) => p[1])), y1 = Math.max(...pr.map((p) => p[1]));
  const sx = plot.w / (x1 - x0 || 1), sy = plot.h / (y1 - y0 || 1);
  return {
    x: (p) => plot.x + (raw(p)[0] - x0) * sx,
    y: (p) => plot.y + (raw(p)[1] - y0) * sy,
    depth: (p) => raw(p)[2],
  };
}

interface Face { pts: V3[]; fill: string; stroke?: string; sw?: number; part?: string; order?: number }

function faceDepth(pr: Proj, fc: Face): number {
  return fc.order ?? fc.pts.reduce((a, p) => a + pr.depth(p), 0) / fc.pts.length;
}

function drawFaces(c: Ctx, pr: Proj, faces: Face[]): void {
  faces.sort((a, b) => faceDepth(pr, b) - faceDepth(pr, a));
  for (const fc of faces) {
    const d = fc.pts.map((p, i) => `${i ? "L" : "M"}${f(pr.x(p))} ${f(pr.y(p))}`).join("") + "Z";
    c.prims.push({ t: "path", d, fill: fc.fill, stroke: fc.stroke, sw: fc.sw, part: fc.part });
  }
}

/** A box's visible faces (front, top or bottom, right and left sides as the view shows them). */
function box(x0: number, x1: number, y0: number, y1: number, z0: number, z1: number, color: string, part: string, rotY: number): Face[] {
  const lo = Math.min(y0, y1), hi = Math.max(y0, y1);
  const p = (x: number, y: number, z: number): V3 => [x, y, z];
  const out: Face[] = [
    { pts: [p(x0, lo, z0), p(x1, lo, z0), p(x1, hi, z0), p(x0, hi, z0)], fill: color, part },
    { pts: [p(x0, hi, z0), p(x1, hi, z0), p(x1, hi, z1), p(x0, hi, z1)], fill: lum(color, 0.82, 0), part },
  ];
  if (rotY >= 0) out.push({ pts: [p(x1, lo, z0), p(x1, lo, z1), p(x1, hi, z1), p(x1, hi, z0)], fill: lum(color, 0.62, 0), part });
  else out.push({ pts: [p(x0, lo, z0), p(x0, lo, z1), p(x0, hi, z1), p(x0, hi, z0)], fill: lum(color, 0.62, 0), part });
  return out;
}

/**
 * Excel's 3-D (and radar) value axes: the unit as for 2-D axes, but the ends on the data rounded
 * to whole units without the 5 % headroom (recorded: 4.5 → 5, 5 → 5).
 */
export function tightScale(lo: number, hi: number, o: Parameters<typeof fixedScale>[2], limit: number): AxisScale {
  // The unit is chosen for the data's own span (no 5 % added before choosing it).
  const s = fixedScale(lo, lo >= 0 && lo < (hi * 5) / 6 ? hi / 1.05 : hi, o, limit);
  const round = (x: number) => Number(x.toFixed(10));
  if (o?.max === undefined && hi > 0) s.max = Math.max(round(Math.ceil(hi / s.major - 1e-9) * s.major), round(s.min + s.major));
  if (o?.min === undefined && lo < 0) s.min = round(Math.floor(lo / s.major + 1e-9) * s.major);
  return s;
}

/** The value range of the chart's series (stacked: the sums, 100 %: 0-1). */
function range(m: ChartModel): [number, number] {
  const g = kindInfo(m.kind).grouping;
  if (g === "percentStacked") return [0, 1];
  let lo = Infinity, hi = -Infinity;
  const n = Math.max(0, ...m.series.map((s) => s.val.cache.length));
  for (let i = 0; i < n; i++) {
    let pos = 0, neg = 0;
    for (const s of m.series) {
      const v = s.val.cache[i];
      if (v === null || v === undefined || !isFinite(v)) continue;
      if (g === "stacked") { if (v >= 0) pos += v; else neg += v; } else { lo = Math.min(lo, v); hi = Math.max(hi, v); }
    }
    if (g === "stacked") { lo = Math.min(lo, neg, pos); hi = Math.max(hi, pos, neg); }
  }
  return lo === Infinity ? [0, 0] : [lo, hi];
}

/** Series values as drawn: stacked on each other (100 %: as shares). */
function stacks(m: ChartModel): { from: number; to: number; v: number | null }[][] {
  const g = kindInfo(m.kind).grouping;
  const n = Math.max(0, ...m.series.map((s) => s.val.cache.length));
  const tot = Array.from({ length: n }, (_, i) => m.series.reduce((a, s) => a + Math.abs(s.val.cache[i] ?? 0), 0));
  const pos = new Array(n).fill(0), neg = new Array(n).fill(0);
  return m.series.map((s) => s.val.cache.map((v0, i) => {
    if (v0 === null || !isFinite(v0)) return { from: 0, to: 0, v: null };
    const v = g === "percentStacked" ? (tot[i] ? v0 / tot[i] : 0) : v0;
    if (g !== "stacked" && g !== "percentStacked") return { from: 0, to: v, v: v0 };
    const acc = v >= 0 ? pos : neg;
    const from = acc[i];
    acc[i] += v;
    return { from, to: acc[i], v: v0 };
  }));
}

/** 3-D columns, bars, lines and areas. */
export function draw3D(c: Ctx, area: Area): Pick<Scene, "plot" | "valueScale"> {
  const { m } = c;
  const info = kindInfo(m.kind);
  const bar = info.dir === "bar";
  const n = Math.max(1, ...m.series.map((s) => s.val.cache.length));
  const ns = Math.max(1, m.series.length);
  const depthRows = info.grouping === "standard";
  const v = view3DOf(m)!;
  const percent = info.grouping === "percentStacked";
  const [lo, hi] = range(m);
  const valFormat = percent ? "0%" : m.series[0]?.val.format;
  const va = m.format?.valAxis;
  const catLabels = Array.from({ length: n }, (_, i) => catText(c, m.series[0], i));
  const serLabels = depthRows ? m.series.map((s, i) => s.name?.text || `Datenreihe${i + 1}`) : [];
  let plot: Box = { x: area.left, y: area.top, w: area.right - area.left, h: area.bottom - area.top };
  let scale: AxisScale = autoScale(lo, hi);
  let vLabels: string[] = [];
  for (let pass = 0; pass < 2; pass++) {
    const len = bar ? plot.w : plot.h;
    scale = percent ? fixedScale(0, 1, { ...(va ?? {}), min: va?.min ?? 0, max: va?.max ?? 1 }, maxIntervals(len, bar)) : tightScale(lo, hi, va, maxIntervals(len * (depthRows ? 0.6 : 1), bar));
    vLabels = ticks(scale).map((t) => c.fmt(t, va?.numFmt ?? valFormat));
    const leftW = Math.max(0, ...(bar ? catLabels : vLabels).map((t) => c.measure(t, LOOK.labelSize)));
    const serW = Math.max(0, ...serLabels.map((t) => c.measure(t, LOOK.labelSize)));
    const left = area.left + PAD + leftW + LABEL_GAP;
    const right = area.right - 10.5 - (serW ? serW + 6 : 0);
    const bottom = area.bottom - 22.6;
    plot = { x: left, y: area.top, w: Math.max(10, right - left), h: Math.max(10, bottom - area.top) };
  }
  // The world: x across the categories (0-1), y the value axis (0-hy), z the depth.
  const catW = 1 / n;
  const dz = ((v.depthPercent ?? 100) / 100) * (depthRows ? ns * catW : catW) * (info.family === "bar" ? 0.67 : info.family === "line" ? 1 : 1.6);
  const hy = bar ? 1 : 0.55;
  const pr = projection(m, plot, hy, dz);
  const span = scale.max - scale.min || 1;
  const val = (t: number) => ((Math.max(scale.min, Math.min(scale.max, t)) - scale.min) / span) * (bar ? 1 : hy);
  const base = val(Math.min(Math.max(0, scale.min), scale.max));
  // World point of (category position, value, depth); bars: categories run up the y axis.
  const W = (cat: number, value: number, z: number): V3 => (bar ? [value, cat * hy, z] : [cat, value, z]);
  // Walls: gridlines on the back wall and the side wall (floor for bars).
  const grid = (a: V3, b: V3, part = "grid") => c.prims.push({ t: "path", d: `M${f(pr.x(a))} ${f(pr.y(a))}L${f(pr.x(b))} ${f(pr.y(b))}`, stroke: LOOK.gridline, sw: 0.75, part });
  ticks(scale).forEach((t, k) => {
    const y = val(t);
    if (m.gridlines) {
      if (bar) { grid([y, 0, 0], [y, 0, dz]); grid([y, 0, dz], [y, hy, dz]); }
      else { grid([0, y, 0], [0, y, dz]); grid([0, y, dz], [1, y, dz]); }
    }
    const at = bar ? W(0, y, 0) : W(0, y, 0);
    if (bar) c.prims.push({ t: "text", x: pr.x(at), y: pr.y(at) + 14, text: vLabels[k], size: LOOK.labelSize, fill: LOOK.text, anchor: "middle", part: "axis:val" });
    else c.prims.push({ t: "text", x: pr.x(at) - LABEL_GAP, y: pr.y(at) + 2.5, text: vLabels[k], size: LOOK.labelSize, fill: LOOK.text, anchor: "end", part: "axis:val" });
  });
  catLabels.forEach((t, i) => {
    const p = W((i + 0.5) * catW, 0, 0);
    if (bar) c.prims.push({ t: "text", x: pr.x([0, (i + 0.5) * catW * hy, 0]) - LABEL_GAP, y: pr.y([0, (i + 0.5) * catW * hy, 0]) + 3.2, text: t, size: LOOK.labelSize, fill: LOOK.text, anchor: "end", part: "axis:cat" });
    else c.prims.push({ t: "text", x: pr.x(p), y: pr.y(p) + 14, text: t, size: LOOK.labelSize, fill: LOOK.text, anchor: "middle", part: "axis:cat" });
  });
  if (depthRows) serLabels.forEach((t, j) => {
    const p: V3 = bar ? [1, 0, (j + 0.5) * (dz / ns)] : [1, 0, (j + 0.5) * (dz / ns)];
    c.prims.push({ t: "text", x: pr.x(p) + 4, y: pr.y(p) + 6, text: t, size: LOOK.labelSize, fill: LOOK.text, anchor: "start", part: "axis:ser" });
  });
  const faces: Face[] = [];
  const st = stacks(m);
  const color = (si: number) => fillPaint(m.series[si].fill, c.color(si)) ?? c.color(si);
  if (info.family === "bar") {
    const gap = (m.gapWidth ?? 150) / 100;
    const k = info.grouping === "clustered" ? ns : 1;
    const bw = catW / (k + gap);
    m.series.forEach((s, si) => st[si].forEach((p, i) => {
      if (p.v === null) return;
      const x0 = i * catW + (bw * gap) / 2 + (info.grouping === "clustered" ? si * bw : 0);
      const zRow = depthRows ? si * (dz / ns) : 0, zW = depthRows ? dz / ns : dz;
      const z0 = zRow + zW * 0.2, z1 = zRow + zW * 0.8;
      const from = info.grouping === "stacked" || info.grouping === "percentStacked" ? val(p.from) : base;
      const to = val(p.to);
      if (bar) {
        const fcs = box(Math.min(from, to), Math.max(from, to), x0 * hy, (x0 + bw) * hy, z0, z1, color(si), `series:${si}:${i}`, v.rotY);
        // Horizontal bars: the "top" face is the end; reorder so the right end is shaded as a side.
        for (const fc of fcs) { fc.order = -(z0 + z1) / 2 - i * 1e-3 - (depthRows ? 0 : si * 1e-4); faces.push(fc); }
      } else {
        for (const fc of box(x0, x0 + bw, from, to, z0, z1, color(si), `series:${si}:${i}`, v.rotY)) {
          fc.order = (depthRows ? zRow : 0) * 10 - i * 1e-3 - si * 1e-4 - (fc.fill === color(si) ? 1e-5 : 0);
          faces.push(fc);
        }
      }
      if (m.dataLabels && p.v !== null) {
        const q = W(x0 + bw / 2, to, z0);
        label(c, s, si, i, p.v, pr.x(q) + (bar ? 4 : 0), pr.y(q) + (bar ? 3.2 : -3.4), bar ? "start" : "middle");
      }
    }));
    // Bars: draw by category and series as Excel stacks them (front faces last per box).
    faces.sort((a, b) => (b.order ?? 0) - (a.order ?? 0));
    for (const fc of faces) {
      const d = fc.pts.map((p, i) => `${i ? "L" : "M"}${f(pr.x(p))} ${f(pr.y(p))}`).join("") + "Z";
      c.prims.push({ t: "path", d, fill: fc.fill, part: fc.part });
    }
  } else {
    // Lines (ribbons) and areas (slabs): one row of depth per series, stacked kinds share one row.
    const rows = depthRows ? ns : 1;
    const rowW = dz / rows;
    const xs = (i: number) => (info.family === "area" ? (n > 1 ? i / (n - 1) : 0.5) : (i + 0.5) * catW);
    const order = depthRows ? [...m.series.keys()].reverse() : [...m.series.keys()];
    for (const si of order) {
      const row = depthRows ? si : 0;
      const z0 = row * rowW + rowW * (info.family === "line" ? 0.25 : 0.15), z1 = row * rowW + rowW * (info.family === "line" ? 0.75 : 0.85);
      const pts = st[si];
      const col = color(si);
      const local: Face[] = [];
      for (let i = 0; i + 1 < pts.length; i++) {
        const a = pts[i], b = pts[i + 1];
        if (a.v === null || b.v === null) continue;
        const ya = val(a.to), yb = val(b.to), x0 = xs(i), x1 = xs(i + 1);
        if (info.family === "line") {
          local.push({ pts: [[x0, ya, z0], [x1, yb, z0], [x1, yb, z1], [x0, ya, z1]], fill: yb >= ya ? lum(col, 0.85, 0) : col, part: `series:${si}` });
        } else {
          const fa = val(a.from), fb = val(b.from);
          local.push({ pts: [[x0, ya, z0], [x1, yb, z0], [x1, yb, z1], [x0, ya, z1]], fill: lum(col, 0.8, 0), part: `series:${si}` });
          local.push({ pts: [[x0, fa, z0], [x1, fb, z0], [x1, yb, z0], [x0, ya, z0]], fill: col, part: `series:${si}`, order: -1e9 + row });
        }
      }
      if (info.family === "area" && pts.length) {
        const last = pts.at(-1)!, xl = xs(pts.length - 1);
        if (last.v !== null) local.push({ pts: [[xl, val(last.from), z0], [xl, val(last.from), z1], [xl, val(last.to), z1], [xl, val(last.to), z0]], fill: lum(col, 0.62, 0), part: `series:${si}` });
      }
      // Each series is drawn as a whole, the farthest first.
      local.sort((a, b) => faceDepth(pr, b) - faceDepth(pr, a));
      for (const fc of local) {
        const d = fc.pts.map((p, i) => `${i ? "L" : "M"}${f(pr.x(p))} ${f(pr.y(p))}`).join("") + "Z";
        c.prims.push({ t: "path", d, fill: fc.fill, stroke: info.family === "area" ? undefined : fc.fill, sw: 0.25, part: fc.part });
      }
      if (m.dataLabels) pts.forEach((p, i) => { if (p.v !== null) { const q: V3 = [xs(i), val(p.to), z0]; label(c, m.series[si], si, i, p.v, pr.x(q), pr.y(q) - 3.4, "middle"); } });
    }
  }
  return { plot, valueScale: scale };
}

/** 3D-Kreis: the slices on a tilted ellipse (rotX: height = width x sin rotX) with their rim below, as Excel draws it. */
export function draw3DPie(c: Ctx, area: Area): Pick<Scene, "plot"> {
  const { m } = c;
  const v = view3DOf(m) ?? { rotX: 30, rotY: 0, rAngAx: false };
  const k = Math.sin((Math.max(10, v.rotX) * Math.PI) / 180);
  const thick = 0.244;
  const top = (m.title.show ? 37.6 : area.top) + 2, bottom = area.bottom - 10.5;
  const left = area.left + 10.5, right = area.right - 10.5;
  const rx = Math.max(10, Math.min((right - left) / 2, (bottom - top) / (2 * k + thick)));
  const ry = rx * k, t = rx * thick;
  const cx = (left + right) / 2, cy = top + ry + Math.max(0, (bottom - top - 2 * ry - t) / 2);
  const s = m.series[0];
  const plot = { x: cx - rx, y: cy - ry, w: 2 * rx, h: 2 * ry + t };
  if (!s) return { plot };
  const vals = s.val.cache.map((x) => (x !== null && x > 0 ? x : 0));
  const sum = vals.reduce((a, b) => a + b, 0);
  if (sum <= 0) return { plot };
  const pt = (a: number, dy = 0) => `${f(cx + rx * Math.sin(a))} ${f(cy - ry * Math.cos(a) + dy)}`;
  let a0 = ((m.firstSliceAngle ?? 0) * Math.PI) / 180;
  const slices = vals.map((x, i) => { const a1 = a0 + (x / sum) * Math.PI * 2; const r = { a0, a1, i }; a0 = a1; return r; });
  const colorOf = (i: number) => fillPaint(s.fill, paletteColor(m, i, vals.length, c.accents)) ?? "#808080";
  // The rim: the part of each slice between 90° and 270° (the front half), darker.
  for (const sl of slices) {
    for (const off of [0, Math.PI * 2]) {
      const b0 = Math.max(sl.a0, Math.PI / 2 + off), b1 = Math.min(sl.a1, (Math.PI * 3) / 2 + off);
      if (b1 <= b0) continue;
      const d = `M${pt(b0)}A${f(rx)} ${f(ry)} 0 0 1 ${pt(b1)}L${pt(b1, t)}A${f(rx)} ${f(ry)} 0 0 0 ${pt(b0, t)}Z`;
      c.prims.push({ t: "path", d, fill: lum(colorOf(sl.i), 0.6, 0), stroke: "#FFFFFF", sw: 0.4, part: `series:0:${sl.i}` });
    }
  }
  for (const sl of slices) {
    if (sl.a1 - sl.a0 <= 0) continue;
    const large = sl.a1 - sl.a0 > Math.PI ? 1 : 0;
    const d = sl.a1 - sl.a0 >= Math.PI * 2 - 1e-9
      ? `M${pt(0)}A${f(rx)} ${f(ry)} 0 1 1 ${pt(Math.PI)}A${f(rx)} ${f(ry)} 0 1 1 ${pt(0)}Z`
      : `M${f(cx)} ${f(cy)}L${pt(sl.a0)}A${f(rx)} ${f(ry)} 0 ${large} 1 ${pt(sl.a1)}Z`;
    c.prims.push({ t: "path", d, fill: colorOf(sl.i), stroke: "#FFFFFF", sw: 1.5, part: `series:0:${sl.i}` });
    if (m.dataLabels) {
      const mid = (sl.a0 + sl.a1) / 2;
      label(c, s, 0, sl.i, s.val.cache[sl.i]!, cx + rx * 0.7 * Math.sin(mid), cy - ry * 0.7 * Math.cos(mid) + 3.2, "middle", vals[sl.i] / sum);
    }
  }
  return { plot };
}

/** The band (0, 1, 2 …) a value falls in, as the surface legend lists them. */
function bandOf(bands: number[], z: number): number {
  for (let i = 1; i < bands.length; i++) if (z <= bands[i] + 1e-9) return i - 1;
  return Math.max(0, bands.length - 2);
}

/** 3D-Oberfläche, Drahtmodell and Kontur: a grid of the series (rows in depth) over the categories, coloured by value band. */
export function drawSurface(c: Ctx, area: Area): Pick<Scene, "plot" | "valueScale"> {
  const { m } = c;
  const info = kindInfo(m.kind);
  const wire = m.kind === "surfaceWireframe" || m.kind === "surfaceContourWireframe";
  const nx = Math.max(1, ...m.series.map((s) => s.val.cache.length)), nz = Math.max(1, m.series.length);
  const bands = surfaceBands(c);
  const scale: AxisScale = { min: bands[0], max: bands.at(-1)!, major: bands[1] - bands[0] };
  const colorOf = (b: number) => seriesColor(b, c.accents);
  const z = (j: number, i: number) => m.series[j]?.val.cache[i] ?? null;
  const catLabels = Array.from({ length: nx }, (_, i) => catText(c, m.series[0], i));
  const serLabels = m.series.map((s, i) => s.name?.text || `Datenreihe${i + 1}`);
  const serW = Math.max(0, ...serLabels.map((t) => c.measure(t, LOOK.labelSize)));
  const vLabels = ticks(scale).map((t) => c.fmt(t, m.series[0]?.val.format));
  const leftW = info.d3 ? Math.max(0, ...vLabels.map((t) => c.measure(t, LOOK.labelSize))) : 0;
  const left = area.left + PAD + leftW + (info.d3 ? LABEL_GAP : 0), right = area.right - 10.5 - serW - 6, bottom = area.bottom - 22.6;
  let plot: Box = { x: left, y: area.top, w: Math.max(10, right - left), h: Math.max(10, bottom - area.top) };
  const sub = 6;
  // Bilinear value inside cell (i, j) at (u, w) in 0-1.
  const at = (i: number, j: number, u: number, w: number): number | null => {
    const a = z(j, i), b = z(j, i + 1), cc = z(j + 1, i), d = z(j + 1, i + 1);
    if (a === null || b === null || cc === null || d === null) return null;
    return a * (1 - u) * (1 - w) + b * u * (1 - w) + cc * (1 - u) * w + d * u * w;
  };
  if (!info.d3) {
    // Kontur: the top view, a square grid; categories below, series at the right.
    const side = Math.min(plot.w, plot.h);
    plot = { x: plot.x + (plot.w - side) / 2, y: plot.y + (plot.h - side) / 2, w: side, h: side };
    const X = (i: number) => plot.x + (nx > 1 ? i / (nx - 1) : 0.5) * plot.w, Y = (j: number) => plot.y + plot.h - (nz > 1 ? j / (nz - 1) : 0.5) * plot.h;
    for (let j = 0; j + 1 < nz; j++) for (let i = 0; i + 1 < nx; i++) {
      for (let a = 0; a < sub; a++) for (let b = 0; b < sub; b++) {
        const val = at(i, j, (a + 0.5) / sub, (b + 0.5) / sub);
        if (val === null) continue;
        const x0 = X(i) + ((X(i + 1) - X(i)) * a) / sub, y0 = Y(j) + ((Y(j + 1) - Y(j)) * (b + 1)) / sub;
        const w = (X(i + 1) - X(i)) / sub, h = (Y(j) - Y(j + 1)) / sub;
        // Light as on the 3-D surface seen from above: the slope darkens.
        const du = (at(i, j, Math.min(1, (a + 1) / sub), (b + 0.5) / sub) ?? val) - (at(i, j, a / sub, (b + 0.5) / sub) ?? val);
        const dw = (at(i, j, (a + 0.5) / sub, Math.min(1, (b + 1) / sub)) ?? val) - (at(i, j, (a + 0.5) / sub, b / sub) ?? val);
        const k = (scale.max - scale.min || 1) / sub;
        const col = lum(colorOf(bandOf(bands, val)), Math.max(0.6, Math.min(1.15, 1 + (-du * 0.5 - dw * 0.8) / k)), 0);
        if (!wire) c.prims.push({ t: "rect", x: x0, y: y0, w: w + 0.3, h: h + 0.3, fill: col, part: "series:0" });
      }
    }
    if (wire) {
      // Drahtmodell: the band boundaries as lines (marching squares) and the grid's edges.
      for (let bi = 1; bi < bands.length - 1; bi++) isolines(c, bands[bi], colorOf(bi - 1), nx, nz, z, X, Y);
      c.prims.push({ t: "rect", x: plot.x, y: plot.y, w: plot.w, h: plot.h, fill: "none", stroke: LOOK.gridline, sw: 0.75, part: "plot" });
    }
    catLabels.forEach((t, i) => c.prims.push({ t: "text", x: X(i), y: plot.y + plot.h + 14, text: t, size: LOOK.labelSize, fill: LOOK.text, anchor: "middle", part: "axis:cat" }));
    serLabels.forEach((t, j) => c.prims.push({ t: "text", x: plot.x + plot.w + 6, y: Y(j) + 3.2, text: t, size: LOOK.labelSize, fill: LOOK.text, anchor: "start", part: "axis:ser" }));
    return { plot, valueScale: scale };
  }
  const hy = 0.55, dz = nz / nx;
  const pr = projection(m, plot, hy, dz);
  const span = scale.max - scale.min || 1;
  const yv = (t: number) => ((t - scale.min) / span) * hy;
  const grid = (a: V3, b: V3) => c.prims.push({ t: "path", d: `M${f(pr.x(a))} ${f(pr.y(a))}L${f(pr.x(b))} ${f(pr.y(b))}`, stroke: LOOK.gridline, sw: 0.75, part: "grid" });
  ticks(scale).forEach((t, k) => {
    const y = yv(t);
    if (m.gridlines) { grid([0, y, 0], [0, y, dz]); grid([0, y, dz], [1, y, dz]); }
    c.prims.push({ t: "text", x: pr.x([0, y, 0]) - LABEL_GAP, y: pr.y([0, y, 0]) + 2.5, text: vLabels[k], size: LOOK.labelSize, fill: LOOK.text, anchor: "end", part: "axis:val" });
  });
  const X = (i: number) => (nx > 1 ? i / (nx - 1) : 0.5), Z = (j: number) => (nz > 1 ? (j / (nz - 1)) * dz : 0);
  catLabels.forEach((t, i) => c.prims.push({ t: "text", x: pr.x([X(i), 0, 0]), y: pr.y([X(i), 0, 0]) + 14, text: t, size: LOOK.labelSize, fill: LOOK.text, anchor: "middle", part: "axis:cat" }));
  serLabels.forEach((t, j) => c.prims.push({ t: "text", x: pr.x([1, 0, Z(j)]) + 5, y: pr.y([1, 0, Z(j)]) + 6, text: t, size: LOOK.labelSize, fill: LOOK.text, anchor: "start", part: "axis:ser" }));
  const faces: Face[] = [];
  const subs = wire ? 1 : 3;
  for (let j = 0; j + 1 < nz; j++) for (let i = 0; i + 1 < nx; i++) {
    for (let a = 0; a < subs; a++) for (let b = 0; b < subs; b++) {
      const u0 = a / subs, u1 = (a + 1) / subs, w0 = b / subs, w1 = (b + 1) / subs;
      const q = [[u0, w0], [u1, w0], [u1, w1], [u0, w1]].map(([u, w]) => [u, w, at(i, j, u, w)] as const);
      if (q.some((p) => p[2] === null)) continue;
      const pts: V3[] = q.map(([u, w, val]) => [X(i) + (X(i + 1) - X(i)) * u, yv(val as number), Z(j) + (Z(j + 1) - Z(j)) * w]);
      const mean = q.reduce((s0, p) => s0 + (p[2] as number), 0) / 4;
      const base = colorOf(bandOf(bands, mean));
      const col = wire ? base : lum(base, shade(pts), 0);
      faces.push(wire ? { pts, fill: "none", stroke: col, sw: 0.75, part: "series:0" } : { pts, fill: col, stroke: col, sw: 0.3, part: "series:0" });
    }
  }
  drawFaces(c, pr, faces);
  return { plot, valueScale: scale };
}

/** Excel's default light on a face: from the upper left front; faces turned away are darker. */
function shade(p: V3[]): number {
  const u = [p[1][0] - p[0][0], p[1][1] - p[0][1], p[1][2] - p[0][2]], w = [p[3][0] - p[0][0], p[3][1] - p[0][1], p[3][2] - p[0][2]];
  let n = [u[1] * w[2] - u[2] * w[1], u[2] * w[0] - u[0] * w[2], u[0] * w[1] - u[1] * w[0]];
  if (n[1] < 0) n = n.map((x) => -x);
  const len = Math.hypot(n[0], n[1], n[2]) || 1;
  const l = [-0.35, 0.8, -0.5], ll = Math.hypot(l[0], l[1], l[2]);
  const d = (n[0] * l[0] + n[1] * l[1] + n[2] * l[2]) / len / ll;
  return 0.62 + 0.45 * Math.max(0, d);
}

/** Marching squares: the line where the surface crosses `level`. */
function isolines(c: Ctx, level: number, color: string, nx: number, nz: number, z: (j: number, i: number) => number | null, X: (i: number) => number, Y: (j: number) => number): void {
  let d = "";
  for (let j = 0; j + 1 < nz; j++) for (let i = 0; i + 1 < nx; i++) {
    const c4 = [[i, j], [i + 1, j], [i + 1, j + 1], [i, j + 1]].map(([a, b]) => ({ x: X(a), y: Y(b), v: z(b, a) }));
    if (c4.some((p) => p.v === null)) continue;
    const cross: [number, number][] = [];
    for (let e = 0; e < 4; e++) {
      const p = c4[e], q = c4[(e + 1) % 4];
      const pv = p.v as number, qv = q.v as number;
      if ((pv - level) * (qv - level) < 0) {
        const t = (level - pv) / (qv - pv);
        cross.push([p.x + (q.x - p.x) * t, p.y + (q.y - p.y) * t]);
      }
    }
    for (let k = 0; k + 1 < cross.length; k += 2) d += `M${f(cross[k][0])} ${f(cross[k][1])}L${f(cross[k + 1][0])} ${f(cross[k + 1][1])}`;
  }
  if (d) c.prims.push({ t: "path", d, stroke: color, sw: 0.75, part: "series:0" });
}

