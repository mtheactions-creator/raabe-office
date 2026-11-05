// Draws a chart as Excel 365 draws its defaults (style 201 of the Office 2013-2022 theme). The
// layout works in points like Excel; positions come from Excel's own SVG export of default
// charts (probe_excel_charts.py, section "layout"): the title's baseline 21.5 pt from the top,
// the plot area's gridlines from 37.5 pt down to 45.5 pt above the bottom with a legend there
// (22.6 pt without), value labels 6.75 pt from the left edge and 8.2 pt before the plot, and
// so on. `chartScene` builds a list of shapes (testable without a DOM); `renderChart` turns it
// into an SVG element sized in CSS px.

import { autoScale, maxIntervals, ticks, type AxisScale } from "./scale";
import { formatGeneral, r2Text, type NumberFormatter } from "./format";
import { fitTrend, histogramBins } from "./stats";
import { fillPaint, fixedScale, linePaint, textLook } from "./paint";
import { schemeColor } from "./colors";
import {
  DEFAULT_AXIS_TITLE, LOOK, OFFICE_ACCENTS, isBar, isPie, isScatter, isStacked, seriesColor, titleText,
  type ChartModel, type ChartSeries, type MarkerSymbol,
} from "./model";

export type Measure = (text: string, sizePt: number) => number;

export interface RenderOptions {
  /** Theme accents 1-6 ("#RRGGBB"); default the Office 2013-2022 theme's. */
  accents?: string[];
  /** Number formatter for labels (value, file-form format code). */
  format?: NumberFormatter;
  /** Text width in pt of `text` at `sizePt`; default: the browser's canvas, else Calibri's metrics. */
  measure?: Measure;
}

export type Prim =
  | { t: "rect"; x: number; y: number; w: number; h: number; fill?: string; stroke?: string; sw?: number; dash?: string; part?: string }
  | { t: "path"; d: string; fill?: string; stroke?: string; sw?: number; dash?: string; cap?: "round" | "butt"; part?: string }
  | { t: "circle"; cx: number; cy: number; r: number; fill: string; stroke?: string; sw?: number; part?: string }
  | { t: "text"; x: number; y: number; text: string; size: number; fill: string; anchor: "start" | "middle" | "end"; rotate?: number; part?: string; bold?: boolean; italic?: boolean };

export interface Box { x: number; y: number; w: number; h: number }

export interface Scene {
  width: number;
  height: number;
  prims: Prim[];
  /** The inner plot area (inside the axes), in pt. */
  plot: Box;
  /** The value axis' scale (scatter: the Y axis). */
  valueScale?: AxisScale;
  /** The X axis' scale of scatter charts. */
  xScale?: AxisScale;
  /** Hit boxes of the chart's elements (title, legend, plot) for the UI. */
  parts: { name: string; box: Box }[];
}

// ----- text measurement -------------------------------------------------------------------------

/** Calibri's advance widths in em (for measuring without a browser). */
const CALIBRI: Record<string, number> = {
  " ": 0.226, "0": 0.507, "1": 0.507, "2": 0.507, "3": 0.507, "4": 0.507, "5": 0.507, "6": 0.507, "7": 0.507, "8": 0.507, "9": 0.507,
  ",": 0.25, ".": 0.252, "-": 0.306, "+": 0.498, "=": 0.498, "(": 0.303, ")": 0.303, "[": 0.307, "]": 0.307, "%": 0.715, "/": 0.386,
  ":": 0.268, "²": 0.35, "€": 0.507, "#": 0.498, "'": 0.221, "!": 0.326, "?": 0.463, "&": 0.682,
  a: 0.479, b: 0.525, c: 0.423, d: 0.525, e: 0.498, f: 0.305, g: 0.471, h: 0.525, i: 0.229, j: 0.239, k: 0.455, l: 0.229, m: 0.799,
  n: 0.525, o: 0.527, p: 0.525, q: 0.525, r: 0.349, s: 0.391, t: 0.335, u: 0.525, v: 0.452, w: 0.715, x: 0.433, y: 0.453, z: 0.395,
  A: 0.579, B: 0.544, C: 0.533, D: 0.615, E: 0.488, F: 0.459, G: 0.631, H: 0.623, I: 0.252, J: 0.319, K: 0.52, L: 0.42, M: 0.855,
  N: 0.646, O: 0.662, P: 0.517, Q: 0.673, R: 0.543, S: 0.459, T: 0.487, U: 0.642, V: 0.567, W: 0.89, X: 0.519, Y: 0.487, Z: 0.468,
  "ä": 0.479, "ö": 0.527, "ü": 0.525, "Ä": 0.579, "Ö": 0.662, "Ü": 0.642, "ß": 0.546,
};

export const calibriMeasure: Measure = (text, size) => {
  let w = 0;
  for (const ch of text) w += CALIBRI[ch] ?? 0.5;
  return w * size;
};

let canvasCtx: CanvasRenderingContext2D | null | undefined;
const widthCache = new Map<string, number>();

/** Measures with the browser's fonts (the same the SVG text uses), else Calibri's metrics. */
export const defaultMeasure: Measure = (text, size) => {
  if (canvasCtx === undefined) {
    canvasCtx = typeof document !== "undefined" ? document.createElement("canvas").getContext("2d") : null;
  }
  if (!canvasCtx) return calibriMeasure(text, size);
  const key = text + "\u0000" + size;
  let w = widthCache.get(key);
  if (w === undefined) {
    canvasCtx.font = `100px ${LOOK.font}`;
    w = (canvasCtx.measureText(text).width / 100) * size;
    if (widthCache.size > 5000) widthCache.clear();
    widthCache.set(key, w);
  }
  return w;
};

// ----- the scene ----------------------------------------------------------------------------------

const PAD = 6.75; // left edge to the first label (Excel)
const LABEL_GAP = 8.2; // value labels to the plot
const LEGEND_SQUARE = 5;

interface Ctx {
  m: ChartModel;
  W: number;
  H: number;
  measure: Measure;
  fmt: (v: number, format?: string) => string;
  accents: string[];
  color: (i: number) => string;
  prims: Prim[];
  parts: { name: string; box: Box }[];
}

/** The chart's shapes for a size in pt. */
export function chartScene(m: ChartModel, W: number, H: number, opts: RenderOptions = {}): Scene {
  const accents = opts.accents ?? OFFICE_ACCENTS;
  const fmt = (v: number, format?: string) => {
    if (opts.format && format && format !== "General") {
      try { return opts.format(v, format); } catch { /* fall back */ }
    }
    return formatGeneral(v);
  };
  const c: Ctx = {
    m, W, H, measure: opts.measure ?? defaultMeasure, fmt, accents,
    color: (i) => m.series[i]?.color ?? paletteColor(m, i, m.series.length, accents),
    prims: [], parts: [],
  };
  // Chart area: white with a light grey border, or its Füllung and Rahmen.
  const ca = m.format?.chart;
  const caLine = linePaint(ca?.line, { color: LOOK.border, width: 0.75 });
  const inset = (caLine?.sw ?? 0) / 2;
  c.prims.push({ t: "rect", x: inset, y: inset, w: W - 2 * inset, h: H - 2 * inset, fill: fillPaint(ca?.fill, "#FFFFFF"), stroke: caLine?.stroke, sw: caLine?.sw, dash: caLine?.dash, part: "chart" });
  const hist = m.kind === "histogram";
  let top = 14.5;
  if (m.title.show) {
    const text = titleText(m);
    const y = hist ? 19 : 21.5;
    const tl = textLook(m.format?.title?.text, { size: LOOK.titleSize, color: LOOK.text });
    const tw = c.measure(text, tl.size);
    const box = { x: W / 2 - tw / 2 - 4, y: y - tl.size - 1, w: tw + 8, h: tl.size + 5 };
    const tf = fillPaint(m.format?.title?.fill, undefined), tln = linePaint(m.format?.title?.line, null);
    if ((tf && tf !== "none") || tln) c.prims.push({ t: "rect", ...box, fill: tf ?? "none", stroke: tln?.stroke, sw: tln?.sw, dash: tln?.dash, part: "title" });
    c.prims.push({ t: "text", x: W / 2, y, text, size: tl.size, fill: tl.fill, bold: tl.bold, italic: tl.italic, anchor: "middle", part: "title" });
    c.parts.push({ name: "title", box });
    // "Titel über dem Diagramm einblenden": the plot keeps the room.
    if (!m.format?.title?.overlay) top = (hist ? 31.75 : 37.5) + (tl.size - LOOK.titleSize);
  }
  // Legend: reserve its room first.
  const area = { left: 0, right: W, top, bottom: H };
  if (m.legend !== "none" && !hist) {
    // "Legende ohne Überlappung des Diagramms anzeigen" off: the plot keeps its room.
    const keep = { ...area };
    layoutLegend(c, area);
    if (m.format?.legend?.overlay) Object.assign(area, keep);
  }
  let scene: Pick<Scene, "plot" | "valueScale" | "xScale">;
  if (isPie(m.kind)) scene = drawPie(c, area);
  else if (hist) scene = drawHistogram(c, area);
  else scene = drawAxes(c, area);
  c.parts.push({ name: "plot", box: scene.plot });
  return { width: W, height: H, prims: c.prims, parts: c.parts, ...scene };
}

// ----- legend ---------------------------------------------------------------------------------------

interface LegendItem { text: string; color: string; kind: "square" | "line" | "marker" | "lineMarker" | "trend" }

const isLineKey = (it: LegendItem) => it.kind === "line" || it.kind === "lineMarker" || it.kind === "trend";

function legendItems(c: Ctx): LegendItem[] {
  const { m } = c;
  if (isPie(m.kind)) {
    const s = m.series[0];
    const n = s ? s.val.cache.length : 0;
    return Array.from({ length: n }, (_, i) => ({ text: catText(c, s, i), color: paletteColor(m, i, n, c.accents), kind: "square" as const }));
  }
  const kind: LegendItem["kind"] = m.kind === "line" ? "line" : m.kind === "lineMarkers" || m.kind === "scatterLines" ? "lineMarker"
    : m.kind === "scatter" ? "marker" : m.kind === "scatterSmooth" ? "line" : "square";
  const items: LegendItem[] = m.series.map((s, i) => ({ text: seriesName(s, i), color: seriesPaint(c, i), kind }));
  // Clustered bars list their series bottom-up, as the bars are drawn (Excel).
  if (m.kind === "bar") items.reverse();
  // Trendlines follow as "Linear (Reihe A)".
  m.series.forEach((s, i) => { if (s.trendline) items.push({ text: trendName(s, i), color: c.color(i), kind: "trend" }); });
  return items;
}

/** A trendline's legend entry: its own name, else "Linear (Reihe A)" and so on (Excel). */
export function trendName(s: ChartSeries, i: number): string {
  const t = s.trendline!;
  if (t.name) return t.name;
  const kind = { linear: "Linear", exp: "Exponentiell", log: "Logarithmisch", poly: "Polynomisch", power: "Potenzfunktion", movingAvg: `${t.period ?? 2} per. Gleitender Durchschnitt` }[t.type];
  return `${kind} (${seriesName(s, i)})`;
}

/** Colour i of n under the chart's Farben ändern scheme (10, Farbig, is the theme's accents in turn). */
export function paletteColor(m: ChartModel, i: number, n: number, accents: string[]): string {
  if (!m.colors || m.colors === 10) return seriesColor(i, accents);
  const theme: Record<string, string> = { dk1: "#000000", lt1: "#FFFFFF" };
  accents.forEach((a, k) => { theme[`accent${k + 1}`] = a; });
  return schemeColor(m.colors, i, n, theme);
}

/** The colour a series shows: its Füllung (bars, areas) or Linie (lines), else its colour. */
function seriesPaint(c: Ctx, i: number): string {
  const s = c.m.series[i];
  const lineKind = c.m.kind === "line" || c.m.kind === "lineMarkers" || isScatter(c.m.kind);
  const own = lineKind ? (s?.line?.type === "solid" ? s.line.color : undefined) : (s?.fill?.type === "solid" ? s.fill.color : undefined);
  return own ?? c.color(i);
}

export function seriesName(s: ChartSeries, i: number): string {
  return s.name?.text || `Datenreihe${i + 1}`;
}

function catText(c: Ctx, s: ChartSeries | undefined, i: number): string {
  const v = s?.cat?.cache[i];
  if (v === undefined || v === null) return s?.cat ? "" : String(i + 1);
  return typeof v === "number" ? c.fmt(v, s?.cat?.format) : v;
}

function itemWidth(c: Ctx, it: LegendItem): number {
  const key = isLineKey(it) ? 19 + 2.2 : LEGEND_SQUARE + 2.1;
  return key + c.measure(it.text, LOOK.labelSize);
}

function drawKey(c: Ctx, it: LegendItem, x: number, base: number): void {
  const mid = base - 2.5;
  if (it.kind === "square") {
    c.prims.push({ t: "rect", x, y: base - LEGEND_SQUARE, w: LEGEND_SQUARE, h: LEGEND_SQUARE, fill: it.color, part: "legend" });
  } else if (it.kind === "marker") {
    c.prims.push({ t: "circle", cx: x + 2.5, cy: mid, r: 2.5, fill: it.color, stroke: it.color, sw: 0.75, part: "legend" });
  } else if (it.kind === "trend") {
    c.prims.push({ t: "path", d: `M${f(x)} ${f(mid)}H${f(x + 19)}`, stroke: it.color, sw: 1.5, dash: "0.1 3", cap: "round", part: "legend" });
  } else {
    c.prims.push({ t: "path", d: `M${f(x)} ${f(mid)}H${f(x + 19)}`, stroke: it.color, sw: it.kind === "line" && c.m.kind !== "scatterSmooth" ? 2.25 : 1.5, cap: "round", part: "legend" });
    if (it.kind === "lineMarker") c.prims.push({ t: "circle", cx: x + 9.5, cy: mid, r: 2.5, fill: it.color, stroke: it.color, sw: 0.75, part: "legend" });
  }
}

function layoutLegend(c: Ctx, area: { left: number; right: number; top: number; bottom: number }): void {
  const items = legendItems(c);
  if (!items.length) return;
  const at = c.prims.length;
  const lf = c.m.format?.legend;
  const lt = textLook(lf?.text, { size: LOOK.labelSize, color: LOOK.text });
  // The legend's own Füllung / Rahmen go under its entries once its box is known.
  const boxUnder = () => {
    const b = c.parts.find((p) => p.name === "legend")?.box;
    const fl = fillPaint(lf?.fill, undefined), ln = linePaint(lf?.line, null);
    if (b && ((fl && fl !== "none") || ln)) c.prims.splice(at, 0, { t: "rect", ...b, fill: fl ?? "none", stroke: ln?.stroke, sw: ln?.sw, dash: ln?.dash, part: "legend" });
  };
  const pos = c.m.legend;
  const textX = (it: LegendItem) => (isLineKey(it) ? 19 + 2.2 : LEGEND_SQUARE + 2.1);
  if (pos === "b" || pos === "t") {
    // One row centred (wrapping to more rows when it is wider than the chart). Space after an
    // entry as measured in Excel: 5.2 pt (line keys 7.5 pt) plus a tenth of its text's width.
    const gapAfter = (it: LegendItem) => (isLineKey(it) ? 7.5 : 5.2) + 0.1 * c.measure(it.text, LOOK.labelSize);
    const rows: LegendItem[][] = [[]];
    let w = 0;
    for (const it of items) {
      const iw = itemWidth(c, it);
      const row = rows.at(-1)!;
      const g = row.length ? gapAfter(row.at(-1)!) : 0;
      if (row.length && w + g + iw > c.W - 20) { rows.push([it]); w = iw; continue; }
      w += g + iw;
      row.push(it);
    }
    const lineH = 12.5;
    const height = rows.length * lineH + 4.4;
    let base = pos === "b" ? c.H - 11.5 - (rows.length - 1) * lineH : area.top + 1.5;
    const y0 = base - 10.9;
    let minX = c.W, maxX = 0;
    for (const row of rows) {
      const rw = row.reduce((s, it, i) => s + itemWidth(c, it) + (i ? gapAfter(row[i - 1]) : 0), 0);
      let x = c.W / 2 - rw / 2;
      minX = Math.min(minX, x);
      maxX = Math.max(maxX, x + rw);
      for (const it of row) {
        drawKey(c, it, x, base);
        c.prims.push({ t: "text", x: x + textX(it), y: base, text: it.text, size: lt.size, fill: lt.fill, bold: lt.bold, italic: lt.italic, anchor: "start", part: "legend" });
        x += itemWidth(c, it) + gapAfter(it);
      }
      base += lineH;
    }
    c.parts.push({ name: "legend", box: { x: minX - 4, y: y0, w: maxX - minX + 8, h: height } });
    boxUnder();
    if (pos === "b") area.bottom = c.H - 22.9 - (rows.length - 1) * lineH;
    else area.top += height + 4;
  } else {
    const lineH = 17;
    const w = Math.max(...items.map((it) => itemWidth(c, it)));
    const x = pos === "r" ? c.W - 10 - w : 10;
    const total = items.length * lineH;
    let base = (area.top + c.H) / 2 - total / 2 + 11;
    const y0 = base - 11;
    for (const it of items) {
      drawKey(c, it, x, base);
      c.prims.push({ t: "text", x: x + textX(it), y: base, text: it.text, size: lt.size, fill: lt.fill, bold: lt.bold, italic: lt.italic, anchor: "start", part: "legend" });
      base += lineH;
    }
    c.parts.push({ name: "legend", box: { x: x - 4, y: y0, w: w + 8, h: total } });
    boxUnder();
    if (pos === "r") area.right = x - 6;
    else area.left = x + w + 6;
  }
}

// ----- pie and doughnut -------------------------------------------------------------------------------

function drawPie(c: Ctx, area: { left: number; right: number; top: number; bottom: number }): Pick<Scene, "plot"> {
  const { m } = c;
  const top = m.title.show ? 37.6 : area.top;
  const bottom = m.legend === "b" ? area.bottom - 10.5 : area.bottom - 10.5;
  const left = area.left + 10.5, right = area.right - 10.5;
  const d = Math.max(10, Math.min(bottom - top, right - left));
  const cx = (left + right) / 2, cy = top + (bottom - top) / 2;
  const r = d / 2;
  const plot = { x: cx - r, y: cy - r, w: d, h: d };
  const series = m.kind === "pie" ? m.series.slice(0, 1) : m.series;
  const hole = m.kind === "doughnut" ? (m.holeSize ?? 75) / 100 : 0;
  const ringW = series.length ? (r * (1 - hole)) / series.length : 0;
  series.forEach((s, si) => {
    const vals = s.val.cache.map((v) => (v !== null && v > 0 ? v : 0));
    const sum = vals.reduce((a, b) => a + b, 0);
    if (sum <= 0) return;
    const ro = r - si * ringW, ri = m.kind === "doughnut" ? ro - ringW : 0;
    let a0 = ((m.firstSliceAngle ?? 0) * Math.PI) / 180;
    const ln = linePaint(s.line, { color: "#FFFFFF", width: 1.5 });
    vals.forEach((v, i) => {
      if (v <= 0) return;
      const a1 = a0 + (v / sum) * Math.PI * 2;
      // Slices take the accents in order (Excel's "Punkte variieren"), unless the series has its own Füllung.
      const color = fillPaint(s.fill, paletteColor(m, i, vals.length, c.accents));
      c.prims.push({ t: "path", d: slice(cx, cy, ro, ri, a0, a1), fill: color, stroke: ln?.stroke, sw: ln?.sw, dash: ln?.dash, part: `series:${si}:${i}` });
      if (m.dataLabels) {
        const mid = (a0 + a1) / 2, lr = ri ? (ro + ri) / 2 : ro * 0.7;
        const tx = cx + lr * Math.sin(mid), ty = cy - lr * Math.cos(mid);
        label(c, s, si, i, s.val.cache[i]!, tx, ty + 3.2, "middle", v / sum);
      }
      a0 = a1;
    });
  });
  return { plot };
}

function slice(cx: number, cy: number, ro: number, ri: number, a0: number, a1: number): string {
  const pt = (r: number, a: number) => `${f(cx + r * Math.sin(a))} ${f(cy - r * Math.cos(a))}`;
  const large = a1 - a0 > Math.PI ? 1 : 0;
  if (a1 - a0 >= Math.PI * 2 - 1e-9) {
    // A full circle: two half arcs.
    const ring = `M${pt(ro, 0)}A${f(ro)} ${f(ro)} 0 1 1 ${pt(ro, Math.PI)}A${f(ro)} ${f(ro)} 0 1 1 ${pt(ro, 0)}Z`;
    return ri ? ring + `M${pt(ri, 0)}A${f(ri)} ${f(ri)} 0 1 0 ${pt(ri, Math.PI)}A${f(ri)} ${f(ri)} 0 1 0 ${pt(ri, 0)}Z` : ring;
  }
  if (ri) return `M${pt(ro, a0)}A${f(ro)} ${f(ro)} 0 ${large} 1 ${pt(ro, a1)}L${pt(ri, a1)}A${f(ri)} ${f(ri)} 0 ${large} 0 ${pt(ri, a0)}Z`;
  return `M${f(cx)} ${f(cy)}L${pt(ro, a0)}A${f(ro)} ${f(ro)} 0 ${large} 1 ${pt(ro, a1)}Z`;
}

// ----- histogram ----------------------------------------------------------------------------------------

/**
 * The count axis of a histogram (Excel 2016's chart engine): the unit as for other value axes,
 * but the top only about 3 % above the largest count (recorded: 29 → 30, 34 → 40, 79 → 90).
 */
export function countScale(maxCount: number, limit: number): AxisScale {
  const s = autoScale(0, maxCount, limit);
  if (maxCount > 0) s.max = Number((Math.ceil((maxCount * 1.032) / s.major - 1e-9) * s.major).toPrecision(12));
  return s;
}

function drawHistogram(c: Ctx, area: { left: number; right: number; top: number; bottom: number }): Pick<Scene, "plot" | "valueScale"> {
  const s = c.m.series[0];
  const bins = histogramBins((s?.val.cache ?? []).filter((v): v is number => v !== null));
  const binLabels = bins.map((b, i) => `${i ? "(" : "["}${formatGeneral(b.from)}, ${formatGeneral(b.to)}]`);
  const binW = Math.max(0, ...binLabels.map((l) => c.measure(l, LOOK.labelSize)));
  const top = area.top, right = area.right - 6.75;
  const maxCount = Math.max(0, ...bins.map((b) => b.count));
  // Bin labels too wide for their bins stand in two rows, the plot 11 pt higher (Excel).
  let stagger = false;
  let bottom = 0, left = 0, scale = autoScale(0, maxCount), labels: string[] = [], lw = 0;
  for (let pass = 0; pass < 2; pass++) {
    bottom = area.bottom - 22.75 - (stagger ? 11 : 0);
    scale = countScale(maxCount, maxIntervals(bottom - top, false));
    labels = ticks(scale).map((v) => formatGeneral(v));
    lw = Math.max(0, ...labels.map((l) => c.measure(l, LOOK.labelSize)));
    left = area.left + 5.45 + lw + 6.67;
    const fits = !bins.length || binW <= (right - left) / bins.length - 4;
    if (fits || stagger) break;
    stagger = true;
  }
  const plot = { x: left, y: top, w: right - left, h: bottom - top };
  const y = (v: number) => bottom - ((v - scale.min) / (scale.max - scale.min)) * plot.h;
  gridAndValueLabels(c, scale, labels, plot, false, 5.45 + lw);
  const bw = bins.length ? plot.w / bins.length : 0;
  bins.forEach((b, i) => {
    const x = left + i * bw;
    if (b.count > 0) c.prims.push({ t: "rect", x: x + 0.25, y: y(b.count), w: bw - 0.75, h: bottom - y(b.count), fill: c.color(0), part: "series:0" });
    const ly = bottom + 14.75 + (stagger && i % 2 === 0 ? 11 : 0);
    c.prims.push({ t: "text", x: x + bw / 2, y: ly, text: binLabels[i], size: LOOK.labelSize, fill: LOOK.text, anchor: "middle", part: "axis:cat" });
    if (c.m.dataLabels && b.count > 0) c.prims.push({ t: "text", x: x + bw / 2, y: y(b.count) - 3.4, text: String(b.count), size: LOOK.labelSize, fill: LOOK.dataLabel, anchor: "middle", part: "labels:0" });
  });
  c.prims.push({ t: "path", d: `M${f(left)} ${f(bottom)}H${f(right)}M${f(left)} ${f(bottom)}V${f(top)}`, stroke: LOOK.gridline, sw: 0.75 });
  return { plot, valueScale: scale };
}

// ----- charts with axes -----------------------------------------------------------------------------------

function numbers(s: ChartSeries): (number | null)[] {
  return s.val.cache;
}

/** The value range the axis must show (stacked charts: the sums per category). */
function valueRange(m: ChartModel): [number, number] {
  let lo = Infinity, hi = -Infinity;
  if (isStacked(m.kind)) {
    const n = Math.max(0, ...m.series.map((s) => s.val.cache.length));
    for (let i = 0; i < n; i++) {
      let pos = 0, neg = 0;
      for (const s of m.series) {
        const v = s.val.cache[i];
        if (v === null || v === undefined) continue;
        if (v >= 0) pos += v; else neg += v;
      }
      lo = Math.min(lo, neg, pos);
      hi = Math.max(hi, pos, neg);
    }
  } else {
    for (const s of m.series) for (const v of numbers(s)) if (v !== null && isFinite(v)) { lo = Math.min(lo, v); hi = Math.max(hi, v); }
  }
  if (lo === Infinity) return [0, 0];
  return [lo, hi];
}

function xRange(m: ChartModel): [number, number] {
  let lo = Infinity, hi = -Infinity;
  for (const s of m.series) {
    s.val.cache.forEach((_, i) => {
      const x = xOf(s, i);
      if (x !== null) { lo = Math.min(lo, x); hi = Math.max(hi, x); }
    });
  }
  return lo === Infinity ? [0, 0] : [lo, hi];
}

/** X of point i of a scatter series: its number, else its position (Excel: text X values count 1, 2, …). */
function xOf(s: ChartSeries, i: number): number | null {
  if (!s.cat || s.cat.cache.some((v) => typeof v === "string")) return i + 1;
  const v = s.cat.cache[i];
  return typeof v === "number" ? v : null;
}

function gridAndValueLabels(c: Ctx, scale: AxisScale, labels: string[], plot: Box, horizontal: boolean, labelRight: number, reverse = false): void {
  const t = ticks(scale);
  const span = scale.max - scale.min || 1;
  const frac = (v: number) => (reverse ? 1 - (v - scale.min) / span : (v - scale.min) / span);
  const fmt = c.m.format;
  const grid = linePaint(fmt?.majorGrid?.line, { color: LOOK.gridline, width: 0.75 });
  const minor = fmt?.minorGrid?.show ? linePaint(fmt.minorGrid.line, { color: "#F2F2F2", width: 0.75 }) : null;
  const vt = textLook(fmt?.valAxis?.text, { size: LOOK.labelSize, color: LOOK.text });
  if (minor) {
    // Hilfsgitternetzlinien: five steps per major unit unless the minor unit is set.
    const unit = fmt?.valAxis?.minorUnit ?? scale.major / 5;
    for (let v = scale.min; v <= scale.max + 1e-9 && unit > 0; v += unit) {
      const p = frac(v);
      c.prims.push(horizontal ? { t: "path", d: `M${f(plot.x + p * plot.w)} ${f(plot.y)}V${f(plot.y + plot.h)}`, ...minor, part: "minorGrid" }
        : { t: "path", d: `M${f(plot.x)} ${f(plot.y + plot.h - p * plot.h)}H${f(plot.x + plot.w)}`, ...minor, part: "minorGrid" });
    }
  }
  t.forEach((v, i) => {
    if (horizontal) {
      const x = plot.x + frac(v) * plot.w;
      if (c.m.gridlines && grid) c.prims.push({ t: "path", d: `M${f(x)} ${f(plot.y)}V${f(plot.y + plot.h)}`, ...grid, part: "grid" });
      if (labels[i]) c.prims.push({ t: "text", x, y: plot.y + plot.h + 14, text: labels[i], size: vt.size, fill: vt.fill, bold: vt.bold, italic: vt.italic, anchor: "middle", part: "axis:val" });
    } else {
      const y = plot.y + plot.h - frac(v) * plot.h;
      if (c.m.gridlines && grid) c.prims.push({ t: "path", d: `M${f(plot.x)} ${f(y)}H${f(plot.x + plot.w)}`, ...grid, part: "grid" });
      if (labels[i]) c.prims.push({ t: "text", x: labelRight, y: y + 2.5, text: labels[i], size: vt.size, fill: vt.fill, bold: vt.bold, italic: vt.italic, anchor: "end", part: "axis:val" });
    }
  });
}

function drawAxes(c: Ctx, area: { left: number; right: number; top: number; bottom: number }): Pick<Scene, "plot" | "valueScale" | "xScale"> {
  const { m } = c;
  const bar = isBar(m.kind), scatter = isScatter(m.kind);
  const nCat = Math.max(0, ...m.series.map((s) => s.val.cache.length));
  const valFormat = m.series[0]?.val.format;
  const [lo, hi] = valueRange(m);
  const [xlo, xhi] = scatter ? xRange(m) : [0, 0];
  const va = m.format?.valAxis, xa = m.format?.catAxis;
  const hideCat = xa?.labelPos === "none" || xa?.visible === false;
  const catLabels = scatter ? [] : Array.from({ length: nCat }, (_, i) => (hideCat ? "" : catText(c, m.series[0], i)));
  const hasValTitle = m.axisTitles.val !== undefined, hasCatTitle = m.axisTitles.cat !== undefined;
  const titleRoom = 15;

  // Pass 1 with a guessed length, pass 2 with the plot's real size (labels change the room).
  let plot: Box = { x: 0, y: area.top, w: c.W, h: c.H };
  let vScale = autoScale(lo, hi), xScale: AxisScale | undefined;
  let vLabels: string[] = [], xLabels: string[] = [];
  let leftLabelsW = 0;
  for (let pass = 0; pass < 3; pass++) {
    const vLen = bar ? plot.w : plot.h;
    vScale = fixedScale(lo, hi, va, maxIntervals(vLen, bar));
    vLabels = va?.labelPos === "none" || va?.visible === false ? ticks(vScale).map(() => "") : ticks(vScale).map((v) => c.fmt(v, va?.numFmt ?? valFormat));
    if (scatter) {
      xScale = fixedScale(xlo, xhi, xa, maxIntervals(plot.w, true));
      xLabels = xa?.labelPos === "none" || xa?.visible === false ? ticks(xScale).map(() => "") : ticks(xScale).map((v) => c.fmt(v, xa?.numFmt ?? m.series[0]?.cat?.format));
    }
    // Labels left of the plot: values (columns, lines, scatter) or categories (bars).
    const leftTexts = bar ? catLabels : vLabels;
    leftLabelsW = Math.max(0, ...leftTexts.map((t) => c.measure(t, LOOK.labelSize)));
    const bottomTexts = bar ? vLabels : scatter ? xLabels : catLabels;
    let left = area.left + PAD + (hasValTitle && !bar ? titleRoom : 0) + (hasCatTitle && bar ? titleRoom : 0) + leftLabelsW + LABEL_GAP;
    let right = area.right - 10.5;
    // Labels centred on the plot's edges (value axes across, area charts) need half their width.
    const edgeLabels = bar || scatter || m.kind === "area";
    if (edgeLabels && bottomTexts.length) {
      right = Math.min(right, area.right - 8.2 - c.measure(bottomTexts.at(-1)!, LOOK.labelSize) / 2);
      left = Math.max(left, area.left + 4 + c.measure(bottomTexts[0], LOOK.labelSize) / 2);
    }
    // Below the plot: its labels (22.6 pt to the legend or the chart's edge) and an axis title.
    const bottomTitle = (hasCatTitle && !bar) || (hasValTitle && bar);
    const bottom = area.bottom - 22.6 - (bottomTitle ? titleRoom + 2 : 0);
    plot = { x: left, y: area.top, w: Math.max(10, right - left), h: Math.max(10, bottom - area.top) };
  }

  // Axis titles.
  if (hasValTitle || hasCatTitle) {
    const vt = m.axisTitles.val, ct = m.axisTitles.cat;
    const leftTitle = bar ? ct : vt, bottomTitle = bar ? vt : ct;
    if (leftTitle !== undefined) {
      const text = leftTitle || DEFAULT_AXIS_TITLE;
      const x = area.left + PAD + 7.5, y = plot.y + plot.h / 2;
      c.prims.push({ t: "text", x, y, text, size: LOOK.axisTitleSize, fill: LOOK.text, anchor: "middle", rotate: -90, part: bar ? "axisTitle:cat" : "axisTitle:val" });
      const tw = c.measure(text, LOOK.axisTitleSize);
      c.parts.push({ name: bar ? "axisTitle:cat" : "axisTitle:val", box: { x: x - 9, y: y - tw / 2 - 3, w: 12, h: tw + 6 } });
    }
    if (bottomTitle !== undefined) {
      const text = bottomTitle || DEFAULT_AXIS_TITLE;
      const y = plot.y + plot.h + 14 + 15;
      c.prims.push({ t: "text", x: plot.x + plot.w / 2, y, text, size: LOOK.axisTitleSize, fill: LOOK.text, anchor: "middle", part: bar ? "axisTitle:val" : "axisTitle:cat" });
      const tw = c.measure(text, LOOK.axisTitleSize);
      c.parts.push({ name: bar ? "axisTitle:val" : "axisTitle:cat", box: { x: plot.x + plot.w / 2 - tw / 2 - 3, y: y - 10, w: tw + 6, h: 13 } });
    }
  }

  // Gridlines and value labels.
  const leftLabelRight = plot.x - LABEL_GAP;
  // Zeichnungsfläche: its own Füllung and Rahmen under everything inside it.
  const pf = fillPaint(m.format?.plot?.fill, undefined), pl = linePaint(m.format?.plot?.line, null);
  if ((pf && pf !== "none") || pl) c.prims.push({ t: "rect", ...plot, fill: pf ?? "none", stroke: pl?.stroke, sw: pl?.sw, dash: pl?.dash, part: "plot" });
  gridAndValueLabels(c, vScale, vLabels, plot, bar, leftLabelRight, !!va?.reverse);
  const vSpan = vScale.max - vScale.min || 1;
  const vFrac = (v: number) => (va?.reverse ? 1 - (v - vScale.min) / vSpan : (v - vScale.min) / vSpan);
  const vPos = (v: number) => (bar ? plot.x + vFrac(v) * plot.w : plot.y + plot.h - vFrac(v) * plot.h);
  const base = vPos(Math.min(Math.max(0, vScale.min), vScale.max));

  if (scatter && xScale) {
    const xs = xScale;
    const xSpan = xs.max - xs.min || 1;
    const xPos = (v: number) => plot.x + ((v - xs.min) / xSpan) * plot.w;
    ticks(xs).forEach((v, i) => {
      if (m.gridlines) c.prims.push({ t: "path", d: `M${f(xPos(v))} ${f(plot.y)}V${f(plot.y + plot.h)}`, stroke: LOOK.gridline, sw: 0.75, part: "grid" });
      c.prims.push({ t: "text", x: xPos(v), y: plot.y + plot.h + 14, text: xLabels[i], size: LOOK.labelSize, fill: LOOK.text, anchor: "middle", part: "axis:cat" });
    });
    // Both axes are drawn where the other crosses zero (or at its minimum).
    const x0 = xPos(Math.min(Math.max(0, xs.min), xs.max));
    c.prims.push({ t: "path", d: `M${f(plot.x)} ${f(base)}H${f(plot.x + plot.w)}M${f(x0)} ${f(plot.y)}V${f(plot.y + plot.h)}`, stroke: LOOK.valueAxisLine, sw: 0.75, part: "axes" });
    clipStart(c, plot);
    m.series.forEach((s, si) => drawScatterSeries(c, s, si, xPos, vPos));
    m.series.forEach((s, si) => { if (s.errorBars) s.val.cache.forEach((v, i) => { const x = xOf(s, i); if (v !== null && x !== null) drawErrorBar(c, s, si, xPos(x), v, vPos, false); }); });
    m.series.forEach((s, si) => { if (s.trendline) drawTrendline(c, s, si, plot, (v) => xPos(v), vPos, (i) => xOf(s, i)); });
    clipEnd(c);
    return { plot, valueScale: vScale, xScale: xs };
  }

  // Category axis: labels and the axis line.
  const n = Math.max(1, nCat);
  const between = m.kind !== "area";
  const catLen = bar ? plot.h : plot.w;
  const step = between ? catLen / n : n > 1 ? catLen / (n - 1) : catLen;
  const catCenter = (i: number) => {
    const j = xa?.reverse ? n - 1 - i : i;
    const p = between ? (j + 0.5) * step : n > 1 ? j * step : catLen / 2;
    return bar ? plot.y + plot.h - p : plot.x + p;
  };
  if (m.format?.catGrid?.show) {
    const cg = linePaint(m.format.catGrid.line, { color: LOOK.gridline, width: 0.75 });
    if (cg) for (let i = 0; i <= n; i++) {
      const p = between ? i * step : Math.min(i, n - 1) * step;
      c.prims.push(bar ? { t: "path", d: `M${f(plot.x)} ${f(plot.y + plot.h - p)}H${f(plot.x + plot.w)}`, ...cg, part: "catGrid" }
        : { t: "path", d: `M${f(plot.x + p)} ${f(plot.y)}V${f(plot.y + plot.h)}`, ...cg, part: "catGrid" });
    }
  }
  const maxCatW = Math.max(0, ...catLabels.map((t) => c.measure(t, LOOK.labelSize)));
  const rotate = !bar && maxCatW > step - 2 && catLabels.length > 1;
  const ct = textLook(xa?.text, { size: LOOK.labelSize, color: LOOK.text });
  catLabels.forEach((t, i) => {
    if (!t) return;
    const look = { size: ct.size, fill: ct.fill, bold: ct.bold, italic: ct.italic };
    if (bar) c.prims.push({ t: "text", x: plot.x - LABEL_GAP, y: catCenter(i) + 3.2, text: t, ...look, anchor: "end", part: "axis:cat" });
    else if (rotate) c.prims.push({ t: "text", x: catCenter(i) + 3, y: plot.y + plot.h + 7, text: t, ...look, anchor: "end", rotate: -45, part: "axis:cat" });
    else c.prims.push({ t: "text", x: catCenter(i), y: plot.y + plot.h + 14, text: t, ...look, anchor: "middle", part: "axis:cat" });
  });
  const axLine = xa?.visible === false ? null : linePaint(xa?.line, { color: LOOK.gridline, width: 0.75 });
  if (axLine) c.prims.push({ t: "path", d: bar ? `M${f(base)} ${f(plot.y)}V${f(plot.y + plot.h)}` : `M${f(plot.x)} ${f(base)}H${f(plot.x + plot.w)}`, ...axLine, part: "axes" });

  clipStart(c, plot);
  if (m.kind === "column" || m.kind === "columnStacked" || bar) drawBars(c, n, step, catCenter, vPos, base);
  else if (m.kind === "area") drawAreas(c, catCenter, vPos, base);
  else drawLines(c, catCenter, vPos);
  m.series.forEach((s, si) => {
    if (!s.errorBars || bar) return;
    // Clustered columns: the bar's middle; other kinds: the category's centre.
    const k = m.kind === "column" ? m.series.length : 1;
    const gap = (m.gapWidth ?? 219) / 100, ov = m.overlap !== undefined ? m.overlap / 100 : -0.27;
    const bw = step / (k - (k - 1) * ov + gap);
    s.val.cache.forEach((v, i) => {
      if (v === null || !isFinite(v)) return;
      const x = m.kind === "column" ? catCenter(i) - step / 2 + (bw * gap) / 2 + si * bw * (1 - ov) + bw / 2 : catCenter(i);
      drawErrorBar(c, s, si, x, v, vPos, false);
    });
  });
  m.series.forEach((s, si) => {
    if (s.trendline) drawTrendline(c, s, si, plot, (x) => (bar ? 0 : catCenter(x - 1)), vPos, (i) => i + 1, bar);
  });
  clipEnd(c);
  return { plot, valueScale: vScale };
}

/** A vertical Fehlerindikator at a point: the amount as the pane sets it, with or without caps. */
function drawErrorBar(c: Ctx, s: ChartSeries, si: number, x: number, v: number, vPos: (v: number) => number, _bar: boolean): void {
  const e = s.errorBars!;
  const vals = s.val.cache.filter((y): y is number => y !== null && isFinite(y));
  const n = vals.length, mean = vals.reduce((a, b) => a + b, 0) / Math.max(1, n);
  const sd = n > 1 ? Math.sqrt(vals.reduce((a, y) => a + (y - mean) ** 2, 0) / (n - 1)) : 0;
  let centre = v, amount: number;
  switch (e.type) {
    case "percentage": amount = Math.abs(v) * (e.value ?? 5) / 100; break;
    // Standard deviation bars stand on the series' mean (Excel).
    case "stdDev": amount = sd * (e.value ?? 1); centre = mean; break;
    case "stdErr": amount = n ? sd / Math.sqrt(n) : 0; break;
    default: amount = e.value ?? 1;
  }
  const hi = e.dir === "minus" ? centre : centre + amount, lo = e.dir === "plus" ? centre : centre - amount;
  const ln = linePaint(e.line, { color: "#595959", width: 0.75 });
  if (!ln) return;
  const y1 = vPos(hi), y0 = vPos(lo), cap = 3;
  let d = `M${f(x)} ${f(y0)}V${f(y1)}`;
  if (!e.noEndCap) {
    if (e.dir !== "minus") d += `M${f(x - cap)} ${f(y1)}H${f(x + cap)}`;
    if (e.dir !== "plus") d += `M${f(x - cap)} ${f(y0)}H${f(x + cap)}`;
  }
  c.prims.push({ t: "path", d, ...ln, part: `errbars:${si}` });
}

function clipStart(c: Ctx, plot: Box): void {
  c.prims.push({ t: "rect", x: plot.x - 4, y: plot.y - 4, w: plot.w + 8, h: plot.h + 8, part: "clip:start" });
}
function clipEnd(c: Ctx): void {
  c.prims.push({ t: "rect", x: 0, y: 0, w: 0, h: 0, part: "clip:end" });
}

function drawBars(c: Ctx, n: number, step: number, catCenter: (i: number) => number, vPos: (v: number) => number, base: number): void {
  const { m } = c;
  const bar = isBar(m.kind), stacked = isStacked(m.kind);
  const gap = (m.gapWidth ?? (m.kind === "column" ? 219 : m.kind === "bar" ? 182 : 150)) / 100;
  const overlap = m.overlap !== undefined ? m.overlap / 100 : m.kind === "column" ? -0.27 : stacked ? 1 : 0;
  const k = stacked ? 1 : m.series.length || 1;
  const bw = step / (k - (k - 1) * overlap + gap);
  const pos = new Array(n).fill(0), neg = new Array(n).fill(0);
  m.series.forEach((s, si) => {
    s.val.cache.forEach((v, i) => {
      if (v === null || !isFinite(v)) return;
      // The group's edge: columns from the left, bars from the bottom (series 1 lowest).
      const start = catCenter(i) + (bar ? 1 : -1) * (step / 2 - (bw * gap) / 2);
      let from: number, to: number;
      if (stacked) {
        const acc = v >= 0 ? pos : neg;
        from = vPos(acc[i]);
        acc[i] += v;
        to = vPos(acc[i]);
      } else {
        from = base;
        to = vPos(v);
      }
      // Offset of series si inside the category group.
      const off = stacked ? 0 : si * bw * (1 - overlap);
      let rect: Box;
      if (bar) {
        const y1 = start - off - bw;
        rect = { x: Math.min(from, to), y: y1, w: Math.abs(to - from), h: bw };
      } else {
        const x1 = start + off;
        rect = { x: x1, y: Math.min(from, to), w: bw, h: Math.abs(to - from) };
      }
      const ln = linePaint(s.line, null);
      const fill = v < 0 && s.invertIfNegative ? "#FFFFFF" : fillPaint(s.fill, c.color(si));
      c.prims.push({ t: "rect", ...rect, fill, stroke: ln?.stroke ?? (v < 0 && s.invertIfNegative ? c.color(si) : undefined), sw: ln?.sw ?? 0.75, dash: ln?.dash, part: `series:${si}:${i}` });
      if (m.dataLabels) {
        // Beschriftungsposition: Außerhalb Ende by default (stacked: Zentriert), or the chosen one.
        const pos = m.format?.dataLabels?.pos ?? (stacked ? "ctr" : "outEnd");
        const neg = v < 0;
        if (bar) {
          const end = neg ? rect.x : rect.x + rect.w, startX = neg ? rect.x + rect.w : rect.x;
          const x = pos === "ctr" ? rect.x + rect.w / 2 : pos === "inEnd" ? end + (neg ? 3 : -3) : pos === "inBase" ? startX + (neg ? -3 : 3) : end + (neg ? -3 : 3);
          const anchor = pos === "ctr" ? "middle" : (pos === "inEnd") !== neg ? "end" : "start";
          label(c, s, si, i, v, x, rect.y + rect.h / 2 + 3.2, anchor as "start" | "middle" | "end");
        } else {
          const top = neg ? rect.y + rect.h : rect.y, bottom = neg ? rect.y : rect.y + rect.h;
          const y = pos === "ctr" ? rect.y + rect.h / 2 + 3.2 : pos === "inEnd" ? (neg ? top - 3.4 : top + 9.6) : pos === "inBase" ? (neg ? bottom + 9.6 : bottom - 3.4) : neg ? top + 9.6 : top - 3.4;
          label(c, s, si, i, v, rect.x + rect.w / 2, y, "middle");
        }
      }
    });
  });
}

function drawAreas(c: Ctx, catCenter: (i: number) => number, vPos: (v: number) => number, base: number): void {
  c.m.series.forEach((s, si) => {
    const pts = s.val.cache.map((v, i) => [catCenter(i), vPos(v ?? 0)] as const);
    if (!pts.length) return;
    const d = `M${f(pts[0][0])} ${f(base)}` + pts.map(([x, y]) => `L${f(x)} ${f(y)}`).join("") + `L${f(pts.at(-1)![0])} ${f(base)}Z`;
    const ln = linePaint(s.line, null);
    c.prims.push({ t: "path", d, fill: fillPaint(s.fill, c.color(si)), stroke: ln?.stroke, sw: ln?.sw, dash: ln?.dash, part: `series:${si}` });
    if (c.m.dataLabels) s.val.cache.forEach((v, i) => {
      if (v !== null) label(c, s, si, i, v, catCenter(i), (pts[i][1] + base) / 2 + 3.2, "middle");
    });
  });
}

function drawLines(c: Ctx, catCenter: (i: number) => number, vPos: (v: number) => number): void {
  const markers = c.m.kind === "lineMarkers";
  c.m.series.forEach((s, si) => {
    const color = c.color(si);
    const ln = linePaint(s.line, { color, width: 2.25 });
    // Runs of points without gaps, drawn straight or smoothed (Linie glätten).
    const runs: [number, number][][] = [[]];
    s.val.cache.forEach((v, i) => {
      if (v === null || !isFinite(v)) { if (runs.at(-1)!.length) runs.push([]); return; }
      runs.at(-1)!.push([catCenter(i), vPos(v)]);
    });
    const d = runs.filter((r) => r.length).map((r) => (s.smooth && r.length > 2 ? smoothPath(r) : r.map(([x, y], i) => `${i ? "L" : "M"}${f(x)} ${f(y)}`).join(""))).join("");
    if (d && ln) c.prims.push({ t: "path", d, ...ln, cap: "round", part: `series:${si}` });
    s.val.cache.forEach((v, i) => {
      if (v === null || !isFinite(v)) return;
      drawMarker(c, s, si, i, catCenter(i), vPos(v), markers ? "circle" : "none");
      if (c.m.dataLabels) label(c, s, si, i, v, catCenter(i) + 5, vPos(v) + 3.2, "start");
    });
  });
}

/** A series' Markierung at a point: its own symbol, size, Füllung and Rahmen, else the chart kind's. */
function drawMarker(c: Ctx, s: ChartSeries, si: number, i: number, x: number, y: number, auto: MarkerSymbol): void {
  const mk = s.marker;
  const sym = !mk?.symbol || mk.symbol === "auto" ? auto : mk.symbol;
  if (sym === "none") return;
  const color = seriesPaint(c, si);
  const r = (mk?.size ?? 5) / 2;
  const fill = fillPaint(mk?.fill, color) ?? color;
  const ln = linePaint(mk?.line, { color, width: 0.75 });
  const st = { stroke: ln?.stroke, sw: ln?.sw, part: `series:${si}:${i}` };
  switch (sym) {
    case "square": c.prims.push({ t: "rect", x: x - r, y: y - r, w: 2 * r, h: 2 * r, fill, ...st }); break;
    case "diamond": c.prims.push({ t: "path", d: `M${f(x)} ${f(y - r)}L${f(x + r)} ${f(y)}L${f(x)} ${f(y + r)}L${f(x - r)} ${f(y)}Z`, fill, ...st }); break;
    case "triangle": c.prims.push({ t: "path", d: `M${f(x)} ${f(y - r)}L${f(x + r)} ${f(y + r)}L${f(x - r)} ${f(y + r)}Z`, fill, ...st }); break;
    case "x": c.prims.push({ t: "path", d: `M${f(x - r)} ${f(y - r)}L${f(x + r)} ${f(y + r)}M${f(x + r)} ${f(y - r)}L${f(x - r)} ${f(y + r)}`, stroke: ln?.stroke ?? color, sw: ln?.sw ?? 0.75, part: st.part }); break;
    case "plus": c.prims.push({ t: "path", d: `M${f(x - r)} ${f(y)}H${f(x + r)}M${f(x)} ${f(y - r)}V${f(y + r)}`, stroke: ln?.stroke ?? color, sw: ln?.sw ?? 0.75, part: st.part }); break;
    case "star": c.prims.push({ t: "path", d: `M${f(x - r)} ${f(y - r)}L${f(x + r)} ${f(y + r)}M${f(x + r)} ${f(y - r)}L${f(x - r)} ${f(y + r)}M${f(x)} ${f(y - r)}V${f(y + r)}`, stroke: ln?.stroke ?? color, sw: ln?.sw ?? 0.75, part: st.part }); break;
    case "dash": c.prims.push({ t: "path", d: `M${f(x - r)} ${f(y)}H${f(x + r)}`, stroke: ln?.stroke ?? color, sw: Math.max(1.5, ln?.sw ?? 0.75), part: st.part }); break;
    case "dot": c.prims.push({ t: "circle", cx: x, cy: y, r: Math.max(1, r / 2), fill, ...st }); break;
    default: c.prims.push({ t: "circle", cx: x, cy: y, r, fill, ...st });
  }
}

/** A data label: value, category, series name, percentage as the Beschriftungsoptionen say, joined by the Trennzeichen. */
function label(c: Ctx, s: ChartSeries, si: number, i: number, v: number, x: number, y: number, anchor: "start" | "middle" | "end", share?: number): void {
  const o = c.m.format?.dataLabels;
  const pie = isPie(c.m.kind);
  const parts: string[] = [];
  if (o?.showSer) parts.push(seriesName(s, si));
  if (o?.showCat) parts.push(catText(c, s, i));
  if (o?.showVal ?? true) parts.push(c.fmt(v, o?.numFmt ?? s.val.format));
  if (o?.showPercent && pie && share !== undefined) parts.push(`${Math.round(share * 100)}%`);
  if (!parts.length) return;
  const look = textLook(o?.text, { size: LOOK.labelSize, color: LOOK.dataLabel });
  const lines = (o?.separator ?? "; ").includes("\n") ? parts : [parts.join(o?.separator ?? "; ")];
  lines.forEach((t, k) => c.prims.push({ t: "text", x, y: y + k * (look.size + 2), text: t, size: look.size, fill: look.fill, bold: look.bold, italic: look.italic, anchor, part: `labels:${si}` }));
}

function drawScatterSeries(c: Ctx, s: ChartSeries, si: number, xPos: (v: number) => number, vPos: (v: number) => number): void {
  const color = c.color(si);
  const pts: [number, number][] = [];
  s.val.cache.forEach((v, i) => {
    const x = xOf(s, i);
    if (v === null || x === null || !isFinite(v)) return;
    pts.push([xPos(x), vPos(v)]);
  });
  const kind = c.m.kind;
  const lineOn = kind === "scatterLines" || kind === "scatterSmooth" || (s.line !== undefined && s.line.type === "solid");
  const ln = lineOn ? linePaint(s.line, { color, width: 1.5 }) : null;
  if (ln && pts.length > 1) {
    const d = kind === "scatterSmooth" || s.smooth ? smoothPath(pts) : pts.map(([x, y], i) => `${i ? "L" : "M"}${f(x)} ${f(y)}`).join("");
    c.prims.push({ t: "path", d, ...ln, cap: "round", part: `series:${si}` });
  }
  pts.forEach(([x, y], i) => drawMarker(c, s, si, i, x, y, kind === "scatterSmooth" ? "none" : "circle"));
  if (c.m.dataLabels) {
    s.val.cache.forEach((v, i) => {
      const x = xOf(s, i);
      if (v === null || x === null) return;
      label(c, s, si, i, v, xPos(x) + 5, vPos(v) + 3.2, "start");
    });
  }
}

/** A Catmull-Rom spline through the points as cubic Béziers (Excel's smoothed lines). */
export function smoothPath(p: [number, number][]): string {
  let d = `M${f(p[0][0])} ${f(p[0][1])}`;
  const t = 1 / 6;
  for (let i = 0; i < p.length - 1; i++) {
    const p0 = p[Math.max(0, i - 1)], p1 = p[i], p2 = p[i + 1], p3 = p[Math.min(p.length - 1, i + 2)];
    const c1 = [p1[0] + (p2[0] - p0[0]) * t, p1[1] + (p2[1] - p0[1]) * t];
    const c2 = [p2[0] - (p3[0] - p1[0]) * t, p2[1] - (p3[1] - p1[1]) * t];
    d += `C${f(c1[0])} ${f(c1[1])} ${f(c2[0])} ${f(c2[1])} ${f(p2[0])} ${f(p2[1])}`;
  }
  return d;
}

/** Linear trendline (dotted, in the series' colour) with Excel's equation and R² label. */
function drawTrendline(c: Ctx, s: ChartSeries, si: number, plot: Box, xPos: (x: number) => number, vPos: (v: number) => number, xOfPoint: (i: number) => number | null, bar = false): void {
  if (bar) return;
  const xs: number[] = [], ys: number[] = [];
  s.val.cache.forEach((v, i) => {
    const x = xOfPoint(i);
    if (v !== null && x !== null && isFinite(v)) { xs.push(x); ys.push(v); }
  });
  const tl = s.trendline!;
  const fit = fitTrend(tl, xs, ys);
  if (!fit) return;
  const x0 = Math.min(...xs) - (tl.backward ?? 0), x1 = Math.max(...xs) + (tl.forward ?? 0);
  // Curves are drawn through enough points to look smooth; a straight line needs two.
  const steps = tl.type === "linear" ? 1 : 60;
  const pts: string[] = [];
  let bx = 0, by = 0;
  for (let k = 0; k <= steps; k++) {
    const x = tl.type === "movingAvg" ? 0 : x0 + ((x1 - x0) * k) / steps;
    const y = fit.at(x);
    if (!isFinite(y)) continue;
    bx = xPos(x); by = vPos(y);
    pts.push(`${pts.length ? "L" : "M"}${f(bx)} ${f(by)}`);
  }
  const path = tl.type === "movingAvg" ? fit.points!.map(([x, y], k) => `${k ? "L" : "M"}${f(xPos(x))} ${f(vPos(y))}`).join("") : pts.join("");
  const tln = linePaint(tl.line, { color: c.color(si), width: 1.5 });
  if (tln) c.prims.push({ t: "path", d: path, stroke: tln.stroke, sw: tln.sw, dash: tl.line?.dash ? tln.dash : "0.1 3", cap: "round", part: `trend:${si}` });
  const lines: string[] = [];
  if (tl.showEq && fit.equation) lines.push(fit.equation);
  if (tl.showR2 && fit.r2 !== undefined) lines.push(r2Text(fit.r2));
  if (!lines.length) return;
  // Excel centres the label 32 pt left of the line's right end, the first line's baseline
  // 2.6 pt above that end (recorded for rising and falling lines).
  const bw = Math.max(...lines.map((l) => c.measure(l, LOOK.labelSize)));
  const cx = Math.max(plot.x + bw / 2 + 2, Math.min(bx - 32, plot.x + plot.w - bw / 2 - 2));
  let first = by - 2.6;
  first = Math.max(plot.y + 9, Math.min(first, plot.y + plot.h - 3 - 11 * (lines.length - 1)));
  lines.forEach((l, i) => c.prims.push({ t: "text", x: cx, y: first + i * 11, text: l, size: LOOK.labelSize, fill: LOOK.text, anchor: "middle", part: `trendLabel:${si}` }));
}

function f(v: number): string {
  return String(Math.round(v * 100) / 100);
}

// ----- SVG ------------------------------------------------------------------------------------------

const SVG_NS = "http://www.w3.org/2000/svg";

/** The chart as an SVG element of `width` x `height` CSS px (drawn in pt, 1 pt = 4/3 px). */
export function renderChart(model: ChartModel, width: number, height: number, opts: RenderOptions = {}): SVGSVGElement {
  const scene = chartScene(model, (width * 72) / 96, (height * 72) / 96, opts);
  const svg = document.createElementNS(SVG_NS, "svg");
  svg.setAttribute("width", String(width));
  svg.setAttribute("height", String(height));
  svg.setAttribute("viewBox", `0 0 ${scene.width} ${scene.height}`);
  svg.setAttribute("font-family", LOOK.font);
  svg.innerHTML = sceneMarkup(scene);
  return svg;
}

/** The scene as SVG markup (inner elements, coordinates in pt). */
export function sceneMarkup(scene: Scene): string {
  const esc = (s: string) => s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
  const out: string[] = [];
  let clipId = 0;
  const part = (p: { part?: string }) => (p.part ? ` data-part="${esc(p.part)}"` : "");
  for (const p of scene.prims) {
    if (p.t === "rect" && p.part === "clip:start") {
      const id = `rc${++clipId}${Math.random().toString(36).slice(2, 7)}`;
      out.push(`<clipPath id="${id}"><rect x="${f(p.x)}" y="${f(p.y)}" width="${f(p.w)}" height="${f(p.h)}"/></clipPath><g clip-path="url(#${id})">`);
      continue;
    }
    if (p.t === "rect" && p.part === "clip:end") { out.push("</g>"); continue; }
    switch (p.t) {
      case "rect":
        out.push(`<rect x="${f(p.x)}" y="${f(p.y)}" width="${f(Math.max(0, p.w))}" height="${f(Math.max(0, p.h))}" fill="${p.fill ?? "none"}"${p.stroke ? ` stroke="${p.stroke}" stroke-width="${p.sw ?? 0.75}"` : ""}${p.dash ? ` stroke-dasharray="${p.dash}"` : ""}${part(p)}/>`);
        break;
      case "path":
        out.push(`<path d="${p.d}" fill="${p.fill ?? "none"}"${p.stroke ? ` stroke="${p.stroke}" stroke-width="${p.sw ?? 0.75}" stroke-linejoin="round"` : ""}${p.cap ? ` stroke-linecap="${p.cap}"` : ""}${p.dash ? ` stroke-dasharray="${p.dash}"` : ""}${part(p)}/>`);
        break;
      case "circle":
        out.push(`<circle cx="${f(p.cx)}" cy="${f(p.cy)}" r="${f(p.r)}" fill="${p.fill}"${p.stroke ? ` stroke="${p.stroke}" stroke-width="${p.sw ?? 0.75}"` : ""}${part(p)}/>`);
        break;
      case "text": {
        const tr = p.rotate ? ` transform="rotate(${p.rotate} ${f(p.x)} ${f(p.y)})"` : "";
        const wt = `${p.bold ? ' font-weight="bold"' : ""}${p.italic ? ' font-style="italic"' : ""}`;
        out.push(`<text x="${f(p.x)}" y="${f(p.y)}" font-size="${p.size}" fill="${p.fill}" text-anchor="${p.anchor}"${wt}${tr}${part(p)}>${esc(p.text)}</text>`);
        break;
      }
    }
  }
  return out.join("");
}
