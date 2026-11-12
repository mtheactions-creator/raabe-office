// The marks of a selected chart and chart element, as Excel draws them (Step 11, D-170; recorded by
// probe_excel_chart_selection.py, results/2026-11-11-excel-chart-selection.json). Office has two kinds of handles:
// - "sizing" handles, grey rings with a white fill, on the parts that can be sized with the mouse (the chart area, the
//   plot area, the legend): eight at the corners and the middles of the sides (the plot area of a pie: corners only);
// - "dots", small blue circles with a light blue fill, on the parts that cannot: the corners of a title's, an axis
//   title's and each data label's box; both ends of every gridline (not the one under the category axis) and of a
//   trendline; the corners of every column or bar of a series (or of its point); a pie's centre and the slices' edges.
//   Axes and markers get smaller dots: at the corners of the axis's box and of a square around each marker.
// The frames are one device pixel wide: dark grey around the chart area and text boxes, a
// lighter grey around the plot area, the legend, an axis, columns and markers; a selected pie gets its slices outlined.
// All sizes are measured at 150 % and kept in CSS px, so they scale with the screen as Office's do.

import type { ChartElement } from "./formatpane";

/** The look in CSS px (Excel at 150 %: 17 px rings 2.6 px wide; 9.5 px dots with 1.3 px rings, 8 px ones on axes and markers with 1.6 px rings). */
export const MARK_LOOK = {
  sizing: { r: 17 / 3, ring: "#616161", ringW: 2.6 / 1.5, fill: "#FFFFFF" },
  dot: { r: 9.5 / 3, ring: "#1683D7", ringW: 1.3 / 1.5, fill: "#D0E7F8" },
  smallDot: { r: 8 / 3, ring: "#1A80D5", ringW: 1.6 / 1.5, fill: "#D0E7F8" },
  frameDark: "#616161",
  frameLight: "#7B7B7B",
  /** A selected pie's slice outlines: two device pixels. */
  outline: "#8A8A8A",
  /** Space between a marker and its square (Excel at 150 %: 2 px). */
  markerPad: 2 / 1.5,
} as const;

type Pt = { x: number; y: number };
export interface MarkRect { x: number; y: number; w: number; h: number }
export interface Marks {
  /** Frames: rectangles drawn with a line of `width` device pixels (1). */
  frames: { r: MarkRect; color: string; width?: number }[];
  /** Outlines: SVG paths in their own user space with the matrix to the marks' CSS px (pie slices). */
  outlines: { d: string; m: [number, number, number, number, number, number]; color: string }[];
  /** Sizing handles' centres. */
  sizing: Pt[];
  /** Dots' centres; small dots (axes, markers). */
  dots: Pt[];
  small: Pt[];
}

const empty = (): Marks => ({ frames: [], outlines: [], sizing: [], dots: [], small: [] });

/** Corners and side middles of a box (Office's eight sizing handles). */
function eight(r: MarkRect, middles = true): Pt[] {
  const xs = [r.x, r.x + r.w / 2, r.x + r.w], ys = [r.y, r.y + r.h / 2, r.y + r.h];
  const out: Pt[] = [];
  for (let j = 0; j < 3; j++) for (let i = 0; i < 3; i++) {
    if (i === 1 && j === 1) continue;
    if (!middles && (i === 1 || j === 1)) continue;
    out.push({ x: xs[i], y: ys[j] });
  }
  return out;
}

function corners(r: MarkRect): Pt[] {
  return [{ x: r.x, y: r.y }, { x: r.x + r.w, y: r.y }, { x: r.x, y: r.y + r.h }, { x: r.x + r.w, y: r.y + r.h }];
}

/** The chart area's marks for a chart of w x h CSS px: its frame and eight sizing handles. */
export function chartMarks(w: number, h: number): Marks {
  const r = { x: 0, y: 0, w, h };
  return { ...empty(), frames: [{ r, color: MARK_LOOK.frameDark }], sizing: eight(r) };
}

/** An element's user-space points in CSS px relative to `origin` (through the screen matrix). */
function toLocal(el: SVGGraphicsElement, origin: DOMRect, pts: Pt[]): Pt[] {
  const m = el.getScreenCTM();
  if (!m) return [];
  return pts.map((p) => {
    const q = new DOMPoint(p.x, p.y).matrixTransform(m);
    return { x: q.x - origin.left, y: q.y - origin.top };
  });
}

function boxOf(els: Element[], origin: DOMRect): MarkRect | null {
  let l = Infinity, t = Infinity, r = -Infinity, b = -Infinity;
  for (const el of els) {
    const rr = el.getBoundingClientRect();
    if (!rr.width && !rr.height) continue;
    l = Math.min(l, rr.left); t = Math.min(t, rr.top); r = Math.max(r, rr.right); b = Math.max(b, rr.bottom);
  }
  return l === Infinity ? null : { x: l - origin.left, y: t - origin.top, w: r - l, h: b - t };
}

function pad(r: MarkRect, px: number, py = px): MarkRect {
  return { x: r.x - px, y: r.y - py, w: r.w + 2 * px, h: r.h + 2 * py };
}

/** The coordinates a path's M, L, H, V and the end points of A commands go through (user space). */
export function pathPoints(d: string): Pt[] {
  const out: Pt[] = [];
  let cur: Pt = { x: 0, y: 0 };
  for (const [, cmd, args] of d.matchAll(/([MLAHVZ])([^MLAHVZ]*)/gi)) {
    const n = (args.match(/-?\d*\.?\d+(?:e-?\d+)?/gi) ?? []).map(Number);
    const c = cmd.toUpperCase();
    if (c === "M" || c === "L") for (let i = 0; i + 1 < n.length; i += 2) { cur = { x: n[i], y: n[i + 1] }; out.push(cur); }
    else if (c === "A") for (let i = 0; i + 6 < n.length; i += 7) { cur = { x: n[i + 5], y: n[i + 6] }; out.push(cur); }
    else if (c === "H") for (const v of n) { cur = { x: v, y: cur.y }; out.push(cur); }
    else if (c === "V") for (const v of n) { cur = { x: cur.x, y: v }; out.push(cur); }
  }
  return out;
}

function uniq(pts: Pt[]): Pt[] {
  const out: Pt[] = [];
  for (const p of pts) if (!out.some((q) => Math.abs(q.x - p.x) < 1 && Math.abs(q.y - p.y) < 1)) out.push(p);
  return out;
}

const near = (a: Pt, b: Pt) => Math.abs(a.x - b.x) < 1.5 && Math.abs(a.y - b.y) < 1.5;

/**
 * The marks of a selected element inside the chart `svg`, in CSS px relative to `origin` (the chart's box on the
 * screen). `selector` picks the element's SVG parts (partsSelector, with the series indices as drawn); `pie` for pie
 * and doughnut charts (the plot area's corners only, the slices' edges).
 */
export function elementMarks(svg: SVGSVGElement, e: ChartElement, selector: string, origin: DOMRect, opts: { pie?: boolean } = {}): Marks {
  const out = empty();
  if (e.kind === "chart") return out;
  if (e.kind === "plot") {
    const box = svg.dataset.plot?.split(",").map(Number);
    const vb = svg.viewBox.baseVal;
    if (!box || box.length !== 4 || !vb?.width) return out;
    const r = svg.getBoundingClientRect();
    const k = r.width / vb.width;
    const rr = { x: r.left - origin.left + box[0] * k, y: r.top - origin.top + box[1] * k, w: box[2] * k, h: box[3] * k };
    out.frames.push({ r: rr, color: MARK_LOOK.frameLight });
    out.sizing.push(...eight(rr, !opts.pie));
    return out;
  }
  const els = selector ? [...svg.querySelectorAll<SVGGraphicsElement>(selector)] : [];
  if (!els.length) return out;
  switch (e.kind) {
    case "legend": {
      const r = boxOf(els, origin);
      if (r) { const p = pad(r, 4, 3); out.frames.push({ r: p, color: MARK_LOOK.frameLight }); out.sizing.push(...eight(p)); }
      return out;
    }
    case "title": case "axisTitle": {
      const r = boxOf(els.filter((x) => x.tagName === "text"), origin) ?? boxOf(els, origin);
      if (r) { const p = pad(r, 4, 2); out.frames.push({ r: p, color: MARK_LOOK.frameDark }); out.dots.push(...corners(p)); }
      return out;
    }
    case "labels": {
      for (const el of els.filter((x) => x.tagName === "text")) {
        const r = boxOf([el], origin);
        if (r) { const p = pad(r, 3, 1); out.frames.push({ r: p, color: MARK_LOOK.frameDark }); out.dots.push(...corners(p)); }
      }
      return out;
    }
    case "majorGrid": case "minorGrid": case "catGrid": case "catMinorGrid": {
      // The gridline under the category axis has no dots (Excel: seven lines' dots for a 0 to 7 axis).
      const axis: Pt[][] = [];
      for (const a of svg.querySelectorAll<SVGGraphicsElement>('path[data-part="axis:cat"],path[data-part="axis:val"]')) {
        const p = pathPoints(a.getAttribute("d") ?? "");
        if (p.length >= 2) axis.push(toLocal(a, origin, [p[0], p[p.length - 1]]));
      }
      for (const el of els) {
        const pts = el.tagName === "path" ? pathPoints(el.getAttribute("d") ?? "") : [];
        if (pts.length < 2) continue;
        const ends = toLocal(el, origin, [pts[0], pts[pts.length - 1]]);
        if (axis.some(([a, b]) => (near(a, ends[0]) && near(b, ends[1])) || (near(a, ends[1]) && near(b, ends[0])))) continue;
        out.dots.push(...ends);
      }
      return out;
    }
    case "trend": {
      for (const el of els.filter((x) => x.tagName === "path")) {
        const pts = pathPoints(el.getAttribute("d") ?? "");
        if (pts.length >= 2) out.dots.push(...toLocal(el, origin, [pts[0], pts[pts.length - 1]]));
      }
      return out;
    }
    case "series": case "point": {
      const own = els.filter((x) => /^series:\d+:\d+$/.test(x.dataset.part ?? ""));
      if (opts.pie && own.length) {
        // A pie: its centre and the slices' edges (a point: its two edges); the whole series also gets its slices outlined.
        const pts: Pt[] = [];
        for (const el of own) {
          const d = el.getAttribute("d") ?? "";
          const p = pathPoints(d);
          if (p.length >= 3) pts.push(...toLocal(el, origin, [p[0], p[1], p[p.length - 1]]));
          const m = el.getScreenCTM();
          if (e.kind === "series" && m) out.outlines.push({ d, m: [m.a, m.b, m.c, m.d, m.e - origin.left, m.f - origin.top], color: MARK_LOOK.outline });
        }
        out.dots.push(...uniq(pts));
        return out;
      }
      // One box per column or bar (dots at its corners), or a square around each marker (small dots).
      const byPoint = new Map<string, SVGGraphicsElement[]>();
      for (const el of own) { const k = el.dataset.part!; byPoint.set(k, [...(byPoint.get(k) ?? []), el]); }
      for (const parts of byPoint.values()) {
        const r = boxOf(parts, origin);
        if (!r) continue;
        if (parts.some((p) => p.tagName === "rect")) {
          out.frames.push({ r, color: MARK_LOOK.frameLight });
          out.dots.push(...corners(r));
        } else {
          const p = pad(r, MARK_LOOK.markerPad);
          out.frames.push({ r: p, color: MARK_LOOK.frameLight });
          out.small.push(...corners(p));
        }
      }
      if (!byPoint.size) {
        // A line or area without markers: its points.
        for (const el of els.filter((x) => x.tagName === "path")) out.dots.push(...uniq(toLocal(el, origin, pathPoints(el.getAttribute("d") ?? ""))));
      }
      return out;
    }
    case "valAxis": case "valAxis2": case "catAxis": {
      const r = boxOf(els, origin);
      if (r) { const p = pad(r, 2); out.frames.push({ r: p, color: MARK_LOOK.frameLight }); out.small.push(...corners(p)); }
      return out;
    }
    default: {
      // Error bars and the rest: the box with dots at its corners.
      const r = boxOf(els, origin);
      if (r) { const p = pad(r, 2); out.frames.push({ r: p, color: MARK_LOOK.frameLight }); out.dots.push(...corners(p)); }
      return out;
    }
  }
}

const NS = "http://www.w3.org/2000/svg";

/**
 * Draws marks into `layer` (an <svg> over the chart, same origin, overflow visible), replacing what it showed.
 * Frames go on whole device pixels: left and top lines on the box's first pixel, right and bottom one pixel beyond
 * its last (as Excel draws the chart area's frame).
 */
export function drawMarks(layer: SVGSVGElement, list: Marks[], dpr = globalThis.devicePixelRatio || 1): void {
  layer.replaceChildren();
  const px = 1 / dpr;
  const el = (tag: string, attrs: Record<string, string | number>) => {
    const n = document.createElementNS(NS, tag);
    for (const [k, v] of Object.entries(attrs)) n.setAttribute(k, String(v));
    layer.append(n);
    return n;
  };
  for (const m of list) {
    for (const o of m.outlines) {
      el("path", { d: o.d, transform: `matrix(${o.m.join(" ")})`, fill: "none", stroke: o.color, "stroke-width": 2 * px, "vector-effect": "non-scaling-stroke" });
    }
    for (const f of m.frames) {
      const w = f.width ?? 1;
      // A line of w device pixels whose first pixel is the box's edge pixel.
      const snap = (v: number) => (Math.round(v * dpr) + w / 2) / dpr;
      const x0 = snap(f.r.x), y0 = snap(f.r.y), x1 = snap(f.r.x + f.r.w), y1 = snap(f.r.y + f.r.h);
      el("path", { d: `M${x0} ${y0}H${x1}V${y1}H${x0}Z`, fill: "none", stroke: f.color, "stroke-width": w * px, "shape-rendering": "crispEdges" });
    }
  }
  const ring = (p: Pt, s: { r: number; ring: string; ringW: number; fill: string }) =>
    el("circle", { cx: p.x, cy: p.y, r: s.r - s.ringW / 2, fill: s.fill, stroke: s.ring, "stroke-width": s.ringW });
  for (const m of list) {
    for (const p of m.sizing) ring(p, MARK_LOOK.sizing);
    for (const p of m.dots) ring(p, MARK_LOOK.dot);
    for (const p of m.small) ring(p, MARK_LOOK.smallDot);
  }
}

/** The overlay <svg> for marks inside a chart's box element (made on first use). */
export function marksLayer(host: HTMLElement): SVGSVGElement {
  let layer = host.querySelector<SVGSVGElement>(":scope > svg.chart-marks");
  if (!layer) {
    layer = document.createElementNS(NS, "svg") as SVGSVGElement;
    layer.setAttribute("class", "chart-marks");
    layer.setAttribute("aria-hidden", "true");
    Object.assign(layer.style, { position: "absolute", left: "0", top: "0", width: "1px", height: "1px", overflow: "visible", pointerEvents: "none", zIndex: "1" });
    host.append(layer);
  }
  return layer;
}
