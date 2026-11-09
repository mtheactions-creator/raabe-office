// Draws a chart as Excel 365 draws its defaults (style 201 of the Office 2013-2022 theme). The
// layout works in points like Excel; positions come from Excel's own SVG export of default
// charts (probe_excel_charts.py, section "layout"): the title's baseline 21.5 pt from the top,
// the plot area's gridlines from 37.5 pt down to 45.5 pt above the bottom with a legend there
// (22.6 pt without), value labels 6.75 pt from the left edge and 8.2 pt before the plot, and
// so on. `chartScene` builds a list of shapes (testable without a DOM); `renderChart` turns it
// into an SVG element sized in CSS px.

import { autoScale, maxIntervals, scaleFrac, ticks, type AxisScale } from "./scale";
import { formatGeneral, r2Text, type NumberFormatter } from "./format";
import { fitTrend, histogramBins } from "./stats";
import { beginDefs, effectFilter, endDefs, fillPaint, setBaseText, fixedScale, linePaint, textAttrs, textLook, unitDivisor, withAlpha, type Stroke } from "./paint";
import { schemeColor } from "./colors";
import {
  DEFAULT_AXIS_TITLE, LOOK, OFFICE_ACCENTS, isBar, isPie, isScatter, lum, seriesColor, titleText,
  type AxisOptions, type ChartModel, type ChartSeries, type MarkerSymbol, type TextFormat, type Effects,
} from "./model";
import { barDefaults, chartGroups, isChartExKind, kindInfo, type ChartGroup } from "./kinds";
import { draw3D, draw3DPie, drawSurface } from "./render3d";
import { drawOfPie, drawRadar } from "./renderspecial";
import { drawChartEx, chartExLegend } from "./renderex";

export type Measure = (text: string, sizePt: number) => number;

export interface RenderOptions {
  /** Theme accents 1-6 ("#RRGGBB"); default the Office 2013-2022 theme's. */
  accents?: string[];
  /** Number formatter for labels (value, file-form format code). */
  format?: NumberFormatter;
  /** Text width in pt of `text` at `sizePt`; default: the browser's canvas, else Calibri's metrics. */
  measure?: Measure;
}

/** Stroke details beyond colour and width (Anschlusstyp, Pfeile, Verbundtyp) and an effect filter. */
export interface PrimExtra { join?: "round" | "bevel" | "miter"; head?: string; tail?: string; compound?: "dbl" | "thickThin" | "thinThick" | "tri"; filter?: string; rx?: number }

export type Prim =
  | ({ t: "rect"; x: number; y: number; w: number; h: number; fill?: string; stroke?: string; sw?: number; dash?: string; cap?: "round" | "butt" | "square"; part?: string } & PrimExtra)
  | ({ t: "path"; d: string; fill?: string; stroke?: string; sw?: number; dash?: string; cap?: "round" | "butt" | "square"; part?: string } & PrimExtra)
  | ({ t: "circle"; cx: number; cy: number; r: number; fill: string; stroke?: string; sw?: number; dash?: string; cap?: "round" | "butt" | "square"; part?: string } & PrimExtra)
  | {
      t: "text"; x: number; y: number; text: string; size: number; fill: string; anchor: "start" | "middle" | "end"; rotate?: number; part?: string; bold?: boolean; italic?: boolean;
      underline?: boolean; strike?: boolean; font?: string; outline?: Stroke; filter?: string;
      /** Spiegelung: a mirrored copy below the baseline. */
      reflection?: { opacity: number; dist: number; blur: number };
    };

export interface Box { x: number; y: number; w: number; h: number }

/** The room left for the plot after title and legend. */
export type Area = { left: number; right: number; top: number; bottom: number };

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
  /** SVG definitions (gradients, patterns, arrowheads, effect filters) the prims refer to. */
  defs?: string[];
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

export const PAD = 6.75; // left edge to the first label (Excel)
export const LABEL_GAP = 8.2; // value labels to the plot
const LEGEND_SQUARE = 5;

export interface Ctx {
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
  const defs = beginDefs();
  setBaseText(m.format?.chart?.text);
  try {
    const scene = buildScene(m, W, H, opts);
    if (defs.list.length) scene.defs = defs.list;
    return scene;
  } finally {
    endDefs();
    setBaseText(undefined);
  }
}

/** A text's rotation in SVG degrees (Textfeld: Benutzerdefinierter Winkel, Textrichtung gedreht). */
export function textRotation(fmt: { text?: TextFormat } | undefined): number | undefined {
  const b = fmt?.text?.box;
  if (!b) return undefined;
  if (b.dir === "vert") return 90;
  if (b.dir === "vert270") return -90;
  return b.rotation ? b.rotation : undefined;
}

/** A stroke's attributes for a primitive (none: no stroke). */
export function strokeOf(s: Stroke | null | undefined): Partial<Pick<Extract<Prim, { t: "path" }>, "stroke" | "sw" | "dash" | "cap" | "join" | "head" | "tail" | "compound">> {
  if (!s) return {};
  const o: ReturnType<typeof strokeOf> = { stroke: s.stroke, sw: s.sw };
  if (s.dash) o.dash = s.dash;
  if (s.cap) o.cap = s.cap;
  if (s.join) o.join = s.join;
  if (s.head) o.head = s.head;
  if (s.tail) o.tail = s.tail;
  if (s.compound) o.compound = s.compound;
  return o;
}

function buildScene(m: ChartModel, W: number, H: number, opts: RenderOptions): Scene {
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
  c.prims.push({ t: "rect", x: inset, y: inset, w: W - 2 * inset, h: H - 2 * inset, fill: fillPaint(ca?.fill, "#FFFFFF"), ...strokeOf(caLine),
    ...(ca?.roundedCorners ? { rx: Math.min(W, H) * 0.05 } : {}), ...(ca?.effects ? { filter: effectFilter(ca.effects) } : {}), part: "chart" });
  const hist = m.kind === "histogram";
  // Excel 2016's chart engine (chartex types) puts the title a little higher.
  const ex = isChartExKind(m.kind);
  let top = 14.5;
  if (m.title.show) {
    const text = titleText(m);
    const y = ex ? 19 : 21.5;
    const tfm = m.format?.title;
    const tl = textLook(tfm?.text, { size: LOOK.titleSize, color: LOOK.text });
    const tw = c.measure(text, tl.size);
    const box = { x: W / 2 - tw / 2 - 4, y: y - tl.size - 1, w: tw + 8, h: tl.size + 5 };
    const tf = fillPaint(tfm?.fill, undefined), tln = linePaint(tfm?.line, null);
    const rot = textRotation(tfm);
    if ((tf && tf !== "none") || tln || tfm?.effects) c.prims.push({ t: "rect", ...box, fill: tf ?? "none", ...strokeOf(tln), filter: effectFilter(tfm?.effects), part: "title" });
    c.prims.push({ t: "text", x: W / 2, y: y - (rot ? tl.size / 3 : 0), text, ...textAttrs(tl), anchor: "middle", ...(rot ? { rotate: rot } : {}), part: "title" });
    c.parts.push({ name: "title", box });
    // "Titel über dem Diagramm einblenden": the plot keeps the room.
    if (!m.format?.title?.overlay) top = (ex ? 31.75 : 37.5) + (tl.size - LOOK.titleSize);
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
  const info = kindInfo(m.kind);
  if (m.kind === "pie3D") scene = draw3DPie(c, area);
  else if (info.family === "ofPie") scene = drawOfPie(c, area);
  else if (isPie(m.kind)) scene = drawPie(c, area);
  else if (hist) scene = drawHistogram(c, area);
  else if (ex) scene = drawChartEx(c, area);
  else if (info.family === "radar") scene = drawRadar(c, area);
  else if (info.family === "surface") scene = drawSurface(c, area);
  else if (info.d3 && info.family !== "bubble") scene = draw3D(c, area);
  else scene = drawAxes(c, area);
  c.parts.push({ name: "plot", box: scene.plot });
  // An axis' own Füllung (and Effekte): a box under its labels, as Excel fills the axis' label area.
  for (const [key, part] of [["valAxis", "axis:val"], ["catAxis", "axis:cat"], ["valAxis2", "axis:val2"]] as const) {
    const o = m.format?.[key];
    const fl = fillPaint(o?.fill, undefined);
    if (!fl || fl === "none") continue;
    const at = c.prims.findIndex((p) => p.part === part && p.t === "text");
    if (at < 0) continue;
    let x0 = Infinity, y0 = Infinity, x1 = -Infinity, y1 = -Infinity;
    for (const p of c.prims) {
      if (p.part !== part || p.t !== "text") continue;
      const w = c.measure(p.text, p.size);
      const l = p.anchor === "start" ? p.x : p.anchor === "end" ? p.x - w : p.x - w / 2;
      x0 = Math.min(x0, l - 1); x1 = Math.max(x1, l + w + 1); y0 = Math.min(y0, p.y - p.size * 0.8); y1 = Math.max(y1, p.y + p.size * 0.25);
    }
    c.prims.splice(at, 0, { t: "rect", x: x0, y: y0, w: x1 - x0, h: y1 - y0, fill: fl, part });
  }
  return { width: W, height: H, prims: c.prims, parts: c.parts, ...scene };
}

// ----- legend ---------------------------------------------------------------------------------------

export interface LegendItem { text: string; color: string; kind: "square" | "line" | "marker" | "lineMarker" | "trend" | "none" | "dot" | "outline" }

const isLineKey = (it: LegendItem) => it.kind === "line" || it.kind === "lineMarker" || it.kind === "trend" || it.kind === "dot" || it.kind === "none";

function legendItems(c: Ctx): LegendItem[] {
  const { m } = c;
  if (isPie(m.kind)) {
    const s = m.series[0];
    const n = s ? s.val.cache.length : 0;
    return Array.from({ length: n }, (_, i) => ({ text: catText(c, s, i), color: paletteColor(m, i, n, c.accents), kind: "square" as const }));
  }
  const own = chartExLegend(c) ?? surfaceLegend(c);
  if (own) return own;
  const items: LegendItem[] = m.series.map((s, i) => ({ text: seriesName(s, i), color: seriesPaint(c, i), kind: legendKey(m, i) }));
  // Clustered bars list their series bottom-up, as the bars are drawn (Excel).
  if (kindInfo(m.kind).dir === "bar" && kindInfo(m.kind).grouping === "clustered") items.reverse();
  // Trendlines follow as "Linear (Reihe A)".
  m.series.forEach((s, i) => { if (s.trendline) items.push({ text: trendName(s, i), color: c.color(i), kind: "trend" }); });
  return items;
}

/** The key a series shows in the legend: its line, marker or a square, as its type draws it. */
function legendKey(m: ChartModel, i: number): LegendItem["kind"] {
  const k = m.series[i]?.kind ?? m.kind, info = kindInfo(k);
  switch (info.family) {
    case "line": return info.d3 ? "square" : info.markers ? "lineMarker" : "line";
    case "radar": return k === "radarFilled" ? "square" : info.markers ? "lineMarker" : "line";
    case "scatter": return info.lines === "none" ? "marker" : info.markers ? "lineMarker" : "line";
    case "bubble": return "marker";
    case "stock": {
      // Volumen columns: a square; the prices have no line, the close of Höchst-Tiefst-Schlusskurs a dot.
      if ((m.kind === "stockVHLC" || m.kind === "stockVOHLC") && i === 0) return "square";
      return i === m.series.length - 1 && (m.kind === "stockHLC" || m.kind === "stockVHLC") ? "dot" : "none";
    }
    default: return "square";
  }
}

/** Oberfläche: one entry per band of the value axis ("0-50", "50-100" …) in the accents. */
export function surfaceBands(c: Ctx): number[] {
  const vals = c.m.series.flatMap((s) => s.val.cache.filter((v): v is number => v !== null && isFinite(v)));
  return ticks(autoScale(Math.min(0, ...vals), Math.max(0, ...vals), 3));
}

function surfaceLegend(c: Ctx): LegendItem[] | null {
  if (kindInfo(c.m.kind).family !== "surface") return null;
  const t = surfaceBands(c);
  const wire = c.m.kind === "surfaceWireframe" || c.m.kind === "surfaceContourWireframe";
  return t.slice(0, -1).map((v, i) => ({ text: `${c.fmt(v)}-${c.fmt(t[i + 1])}`, color: seriesColor(i, c.accents), kind: wire ? "outline" : "square" }));
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
export function seriesPaint(c: Ctx, i: number): string {
  const s = c.m.series[i];
  const si = kindInfo(s?.kind ?? c.m.kind);
  const lineKind = (si.family === "line" && !si.d3) || si.family === "scatter" || (si.family === "radar" && c.m.kind !== "radarFilled");
  const own = lineKind ? (s?.line?.type === "solid" ? s.line.color : undefined) : (s?.fill?.type === "solid" ? s.fill.color : undefined);
  return own ?? c.color(i);
}

export function seriesName(s: ChartSeries, i: number): string {
  return s.name?.text || `Datenreihe${i + 1}`;
}

export function catText(c: Ctx, s: ChartSeries | undefined, i: number): string {
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
  if (it.kind === "none") return;
  if (it.kind === "dot") {
    c.prims.push({ t: "path", d: `M${f(x + 7.5)} ${f(mid)}H${f(x + 11.5)}`, stroke: it.color, sw: 1.5, part: "legend" });
  } else if (it.kind === "outline") {
    c.prims.push({ t: "rect", x: x + 0.4, y: base - LEGEND_SQUARE + 0.4, w: LEGEND_SQUARE - 0.8, h: LEGEND_SQUARE - 0.8, fill: "none", stroke: it.color, sw: 0.75, part: "legend" });
  } else if (it.kind === "square") {
    c.prims.push({ t: "rect", x, y: base - LEGEND_SQUARE, w: LEGEND_SQUARE, h: LEGEND_SQUARE, fill: it.color, part: "legend" });
  } else if (it.kind === "marker") {
    c.prims.push({ t: "circle", cx: x + 2.5, cy: mid, r: 2.5, fill: it.color, stroke: it.color, sw: 0.75, part: "legend" });
  } else if (it.kind === "trend") {
    c.prims.push({ t: "path", d: `M${f(x)} ${f(mid)}H${f(x + 19)}`, stroke: it.color, sw: 1.5, dash: "0.1 3", cap: "round", part: "legend" });
  } else {
    c.prims.push({ t: "path", d: `M${f(x)} ${f(mid)}H${f(x + 19)}`, stroke: it.color, sw: it.kind === "line" && !isScatter(c.m.kind) ? 2.25 : 1.5, cap: "round", part: "legend" });
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
    if (b && ((fl && fl !== "none") || ln || lf?.effects)) c.prims.splice(at, 0, { t: "rect", ...b, fill: fl ?? "none", ...strokeOf(ln), filter: effectFilter(lf?.effects), part: "legend" });
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
        c.prims.push({ t: "text", x: x + textX(it), y: base, text: it.text, ...textAttrs(lt), anchor: "start", part: "legend" });
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
    const x = pos === "r" || pos === "tr" ? c.W - 10 - w : 10;
    const total = items.length * lineH;
    // Oben rechts: the column in the top right corner.
    let base = pos === "tr" ? area.top + 11 : (area.top + c.H) / 2 - total / 2 + 11;
    const y0 = base - 11;
    for (const it of items) {
      drawKey(c, it, x, base);
      c.prims.push({ t: "text", x: x + textX(it), y: base, text: it.text, ...textAttrs(lt), anchor: "start", part: "legend" });
      base += lineH;
    }
    c.parts.push({ name: "legend", box: { x: x - 4, y: y0, w: w + 8, h: total } });
    boxUnder();
    if (pos === "r" || pos === "tr") area.right = x - 6;
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
  const series = m.kind === "pie" ? m.series.slice(0, 1) : m.series;
  // Punktexplosion: the pie shrinks so the slices pulled out still fit (Excel).
  const explode = (s: ChartSeries, i: number) => (pointOf(s, i)?.explosion ?? s.explosion ?? 0) / 100;
  const maxExp = m.kind === "pie" ? Math.max(0, ...series.flatMap((s) => s.val.cache.map((_, i) => explode(s, i)))) : 0;
  const r = d / 2 / (1 + maxExp);
  const plot = { x: cx - r, y: cy - r, w: 2 * r, h: 2 * r };
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
      const pt = pointOf(s, i);
      const color = fillPaint(pt?.fill ?? s.fill, paletteColor(m, i, vals.length, c.accents));
      const mid = (a0 + a1) / 2, ex = explode(s, i) * (ri ? ringW : r);
      const ox = cx + ex * Math.sin(mid), oy = cy - ex * Math.cos(mid);
      const pln = pt?.line ? linePaint(pt.line, { color: "#FFFFFF", width: 1.5 }) : ln;
      const fx = effectFilter(pt?.effects ?? s.effects);
      c.prims.push({ t: "path", d: slice(ox, oy, ro, ri, a0, a1), fill: color, ...strokeOf(pln), ...(fx ? { filter: fx } : {}), part: `series:${si}:${i}` });
      if (m.dataLabels) {
        const lr = ri ? (ro + ri) / 2 : ro * 0.7;
        const tx = ox + lr * Math.sin(mid), ty = oy - lr * Math.cos(mid);
        label(c, s, si, i, s.val.cache[i]!, tx, ty + 3.2, "middle", v / sum);
      }
      a0 = a1;
    });
  });
  return { plot };
}

export function slice(cx: number, cy: number, ro: number, ri: number, a0: number, a1: number): string {
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

export function gridAndValueLabels(c: Ctx, scale: AxisScale, labels: string[], plot: Box, horizontal: boolean, labelRight: number, reverse = false): void {
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

/** The values a group of series puts on its axis: sums per category when stacked, 0-100 % when 100 % stacked. */
function groupRange(m: ChartModel, groups: ChartGroup[]): [number, number] {
  let lo = Infinity, hi = -Infinity;
  for (const g of groups) {
    const grouping = kindInfo(g.kind).grouping;
    const list = g.idx.map((i) => m.series[i]).filter(Boolean);
    if (grouping === "percentStacked") {
      const neg = list.some((s) => s.val.cache.some((v) => v !== null && v < 0));
      lo = Math.min(lo, neg ? -1 : 0);
      hi = Math.max(hi, 1);
    } else if (grouping === "stacked") {
      const n = Math.max(0, ...list.map((s) => s.val.cache.length));
      for (let i = 0; i < n; i++) {
        let pos = 0, neg = 0;
        for (const s of list) {
          const v = s.val.cache[i];
          if (v === null || v === undefined) continue;
          if (v >= 0) pos += v; else neg += v;
        }
        lo = Math.min(lo, neg, pos);
        hi = Math.max(hi, pos, neg);
      }
    } else {
      for (const s of list) for (const v of numbers(s)) if (v !== null && isFinite(v)) { lo = Math.min(lo, v); hi = Math.max(hi, v); }
    }
  }
  if (lo === Infinity) return [0, 0];
  return [lo, hi];
}

/** A value axis' scale: 100 % stacked axes run from 0 to 100 % unless the pane fixes them. */
function axisScale(lo: number, hi: number, o: AxisOptions | undefined, limit: number, percent: boolean): AxisScale {
  if (!percent) return fixedScale(lo, hi, o, limit);
  return fixedScale(lo, hi, { ...(o ?? {}), min: o?.min ?? (lo < 0 ? -1 : 0), max: o?.max ?? 1 }, limit);
}

const Z_ORDER: Record<string, number> = { area: 0, bar: 1, stock: 2, line: 3 };

/** Teilstriche: the length of a major and a minor tick mark in pt beyond the axis line (Excel's PDF export, results/2026-11-09-excel-chart-pane.json). */
export const TICK_MAJOR = 2.75, TICK_MINOR = 2.05;

/** Excel's serial date to a JS date (UTC) and back (1900 system). */
const serialToDate = (v: number) => new Date(Date.UTC(1899, 11, 30) + Math.round(v) * 86400000);
const dateToSerial = (d: Date) => Math.round((d.getTime() - Date.UTC(1899, 11, 30)) / 86400000);

/** One axis line with its tick marks (Teilstriche), in the axis line's paint as Excel draws them. */
function axisLineAndTicks(c: Ctx, o: { horizontal: boolean; at: number; from: number; to: number; line: Stroke | null; majors: number[]; minors: number[];
  major?: AxisOptions["majorTick"]; minor?: AxisOptions["minorTick"]; out: 1 | -1; part: string; effects?: Effects }): void {
  if (!o.line) return;
  const ln = strokeOf(o.line);
  const filter = o.effects ? effectFilter(o.effects) : undefined;
  let d = o.horizontal ? `M${f(o.from)} ${f(o.at)}H${f(o.to)}` : `M${f(o.at)} ${f(o.from)}V${f(o.to)}`;
  const tick = (p: number, kind: AxisOptions["majorTick"], len: number) => {
    if (!kind || kind === "none") return;
    const a = kind === "in" ? 0 : kind === "out" ? len * o.out : len * o.out;
    const b = kind === "in" ? -len * o.out : kind === "out" ? 0 : -len * o.out;
    d += o.horizontal ? `M${f(p)} ${f(o.at + a)}V${f(o.at + b)}` : `M${f(o.at + a)} ${f(p)}H${f(o.at + b)}`;
  };
  for (const p of o.minors) tick(p, o.minor, TICK_MINOR);
  for (const p of o.majors) tick(p, o.major, TICK_MAJOR);
  c.prims.push({ t: "path", d, ...ln, head: undefined, tail: undefined, ...(filter ? { filter } : {}), part: o.part });
  // Pfeile: the arrowheads go on the axis line itself (a path of its own so the ticks do not take them).
  if (o.line.head || o.line.tail) {
    const line = o.horizontal ? `M${f(o.from)} ${f(o.at)}H${f(o.to)}` : `M${f(o.at)} ${f(o.to)}V${f(o.from)}`;
    c.prims.push({ t: "path", d: line, stroke: "none", sw: ln.sw, head: o.line.head, tail: o.line.tail, part: o.part });
  }
}

/** The minor tick values of a scale (the Hilfseinheit, else a fifth of the major unit; log: the multiples in each power). */
function minorTicks(s: AxisScale, unit: number | undefined): number[] {
  const out: number[] = [];
  if (s.log) {
    for (let p = s.min; p < s.max * (1 - 1e-9); p *= s.log) for (let k = 2; k < s.log; k++) out.push(p * k);
    return out;
  }
  const u = unit ?? s.major / 5;
  if (!(u > 0) || (s.max - s.min) / u > 2000) return out;
  for (let v = s.min; v <= s.max + 1e-9; v += u) out.push(Number(v.toPrecision(12)));
  return out;
}

/** Text of an Anzeigeeinheit's label (Excel's German names). */
export function displayUnitText(o: AxisOptions | undefined): string {
  const u = o?.dispUnits;
  const names: Record<string, string> = {
    hundreds: "Hunderte", thousands: "Tausende", tenThousands: "x 10000", hundredThousands: "x 100000", millions: "Millionen",
    tenMillions: "x 10000000", hundredMillions: "x 100000000", billions: "Milliarden", trillions: "Billionen",
  };
  if (u === undefined) return "";
  return typeof u === "number" ? `x ${formatGeneral(u)}` : names[u] ?? "";
}

function drawAxes(c: Ctx, area: { left: number; right: number; top: number; bottom: number }): Pick<Scene, "plot" | "valueScale" | "xScale"> {
  const { m } = c;
  const bar = isBar(m.kind), scatter = isScatter(m.kind);
  const groups = isScatter(m.kind) ? [{ kind: m.kind, idx: m.series.map((_, i) => i), secondary: false }] : chartGroups(m);
  const prim = groups.filter((g) => !g.secondary), sec = groups.filter((g) => g.secondary);
  const percent = prim.some((g) => kindInfo(g.kind).grouping === "percentStacked");
  const percent2 = sec.some((g) => kindInfo(g.kind).grouping === "percentStacked");
  const valFormat = percent ? "0%" : m.series[prim[0]?.idx[0] ?? 0]?.val.format;
  const valFormat2 = percent2 ? "0%" : m.series[sec[0]?.idx[0] ?? 0]?.val.format;
  const [lo, hi] = groupRange(m, prim.length ? prim : groups);
  const [lo2, hi2] = sec.length ? groupRange(m, sec) : [0, 0];
  const [xlo, xhi] = scatter ? xRange(m) : [0, 0];
  const va = m.format?.valAxis, xa = m.format?.catAxis, va2 = m.format?.valAxis2;
  const hideVal = va?.labelPos === "none" || va?.visible === false;
  const hideCat = xa?.labelPos === "none" || xa?.visible === false;
  const allArea = groups.every((g) => kindInfo(g.kind).family === "area");
  // Achsenposition: Zwischen Teilstrichen (columns, lines) or Auf Teilstrichen (areas), or as the pane sets it.
  const between = xa?.onTicks !== undefined ? !xa.onTicks : !allArea;

  // Datumsachse: the categories stand at their dates, one slot per Basis unit.
  const cat0 = m.series[0]?.cat;
  const dateAxis = !scatter && xa?.axisType === "date" && !!cat0 && cat0.cache.some((v) => typeof v === "number");
  const nPoints = Math.max(0, ...m.series.map((s) => s.val.cache.length));
  let slotOf = (i: number) => i;
  let nCat = nPoints;
  let slotLabel: ((k: number) => string) | null = null;
  if (dateAxis) {
    const nums = cat0!.cache.map((v) => (typeof v === "number" ? v : null));
    const valid = nums.filter((v): v is number => v !== null);
    const d0 = Math.min(...valid);
    const unit = xa?.baseTimeUnit ?? "days";
    const months = (v: number) => { const a = serialToDate(d0), b = serialToDate(v); return (b.getUTCFullYear() - a.getUTCFullYear()) * 12 + b.getUTCMonth() - a.getUTCMonth(); };
    const slot = (v: number) => (unit === "days" ? Math.round(v - d0) : unit === "months" ? months(v) : Math.floor(months(v) / 12));
    slotOf = (i) => (nums[i] === null ? i : slot(nums[i]!));
    nCat = Math.max(1, ...valid.map(slot).map((x) => x + 1));
    slotLabel = (k) => {
      const a = serialToDate(d0);
      const d = unit === "days" ? new Date(a.getTime() + k * 86400000) : new Date(Date.UTC(a.getUTCFullYear() + (unit === "years" ? k : 0), a.getUTCMonth() + (unit === "months" ? k : 0), 1));
      return c.fmt(dateToSerial(d), xa?.numFmt ?? cat0!.format ?? "TT.MM.JJJJ");
    };
  }
  // Intervall zwischen Beschriftungen; Datumsachse: a label every Haupteinheit.
  const lblSkip = Math.max(1, Math.round(xa?.tickLblSkip ?? (dateAxis ? xa?.majorUnit ?? 1 : 1)));
  const catLabels = scatter ? [] : Array.from({ length: nCat }, (_, i) => (hideCat || i % lblSkip ? "" : slotLabel ? slotLabel(i) : catText(c, m.series[0], i)));
  const hasValTitle = m.axisTitles.val !== undefined, hasCatTitle = m.axisTitles.cat !== undefined;
  const titleRoom = 15;

  // Which end of the plot each axis' labels stand at (Beschriftungsposition, where the other axis crosses, reversed orders).
  // Value labels go with the value axis, which stands at the first category unless "Vertikale Achse schneidet" moves it.
  const valAnchor = va?.labelPos === "high" ? "last" : va?.labelPos === "low" ? "first"
    : xa?.crosses === "max" || (typeof xa?.crosses === "number" && !scatter && xa.crosses > nCat) ? "last" : "first";
  const valLabelsEnd = (valAnchor === "last") !== !!xa?.reverse;
  const catAnchor = xa?.labelPos === "high" ? "max" : xa?.labelPos === "low" ? "min" : va?.crosses === "max" ? "max" : "min";
  const catLabelsEnd = (catAnchor === "max") !== !!va?.reverse;
  // In the plot's terms: V labels left (start) or right (end), H labels below (start) or above (end).
  const vEnd = bar ? catLabelsEnd : valLabelsEnd, hEnd = bar ? valLabelsEnd : catLabelsEnd;
  const unitLabel = va?.dispUnitsLabel !== false && va?.dispUnits !== undefined && !hideVal ? displayUnitText(va) : "";
  const offset = (xa?.lblOffset ?? 100) / 100;
  const catRot = xa?.text?.box?.rotation ?? (xa?.text?.box?.dir === "vert270" ? -90 : xa?.text?.box?.dir === "vert" ? 90 : undefined);
  const vt0 = textLook(va?.text, { size: LOOK.labelSize, color: LOOK.text });
  const ct0 = textLook(xa?.text, { size: LOOK.labelSize, color: LOOK.text });

  // Pass 1 with a guessed length, pass 2 with the plot's real size (labels change the room).
  let plot: Box = { x: 0, y: area.top, w: c.W, h: c.H };
  let vScale = autoScale(lo, hi), xScale: AxisScale | undefined, vScale2: AxisScale | undefined;
  let vLabels: string[] = [], xLabels: string[] = [], vLabels2: string[] = [];
  const vDiv = unitDivisor(va), xDiv = unitDivisor(xa);
  for (let pass = 0; pass < 3; pass++) {
    const vLen = bar ? plot.w : plot.h;
    vScale = axisScale(lo, hi, va, maxIntervals(vLen, bar), percent);
    vLabels = hideVal ? ticks(vScale).map(() => "") : ticks(vScale).map((v) => c.fmt(v / vDiv, va?.numFmt ?? valFormat));
    if (sec.length) {
      vScale2 = axisScale(lo2, hi2, va2, maxIntervals(vLen, bar), percent2);
      const hide2 = va2?.labelPos === "none" || va2?.visible === false;
      vLabels2 = ticks(vScale2).map((v) => (hide2 ? "" : c.fmt(v / unitDivisor(va2), va2?.numFmt ?? valFormat2)));
    }
    if (scatter) {
      xScale = fixedScale(xlo, xhi, xa, maxIntervals(plot.w, true));
      xLabels = hideCat ? ticks(xScale).map(() => "") : ticks(xScale).map((v) => c.fmt(v / xDiv, xa?.numFmt ?? m.series[0]?.cat?.format));
    }
    const hTexts = bar ? vLabels : scatter ? xLabels : catLabels;
    const vTexts = bar ? catLabels : vLabels;
    const vSize = bar ? ct0.size : vt0.size;
    const vW = Math.max(0, ...vTexts.map((t) => c.measure(t, vSize)));
    const titleLeft = (hasValTitle && !bar) || (hasCatTitle && bar) ? titleRoom : 0;
    const unitRoom = unitLabel && !bar ? 12 : 0;
    let left = area.left + PAD + titleLeft + (vEnd ? 0 : vW + LABEL_GAP + unitRoom);
    if (vEnd) left = Math.max(left, area.left + 10.5 + titleLeft);
    let right = area.right - 10.5;
    if (vEnd) right = area.right - PAD - vW - LABEL_GAP - unitRoom;
    // Sekundärachse: its labels right of the plot (bars: above it).
    if (sec.length && !bar) right -= Math.max(0, ...vLabels2.map((t) => c.measure(t, LOOK.labelSize))) + LABEL_GAP - 4;
    // Labels centred on the plot's edges (value axes across, area charts) need half their width.
    const edgeLabels = bar || scatter || !between;
    if (edgeLabels && hTexts.length) {
      right = Math.min(right, area.right - 8.2 - c.measure(hTexts.at(-1)!, LOOK.labelSize) / 2);
      left = Math.max(left, area.left + 4 + c.measure(hTexts[0], LOOK.labelSize) / 2);
    }
    // Below the plot: its labels (22.6 pt to the legend or the chart's edge) and an axis title.
    const bottomTitle = (hasCatTitle && !bar) || (hasValTitle && bar);
    const hSize = bar ? vt0.size : ct0.size;
    const hRoom = (hTexts.some((t) => t) ? 22.6 + (hSize - LOOK.labelSize) + (bar ? 0 : (offset - 1) * 2.92) : 10.5)
      + (catRot && !bar ? Math.max(0, ...hTexts.map((t) => c.measure(t, hSize))) * Math.abs(Math.sin((catRot * Math.PI) / 180)) : 0)
      + (unitLabel && bar ? 12 : 0);
    const bottom = area.bottom - (hEnd ? 10.5 : hRoom) - (bottomTitle ? titleRoom + 2 : 0);
    const top = area.top + (hEnd ? hRoom - 8.6 : 0) + (sec.length && bar ? 14 : 0);
    plot = { x: left, y: top, w: Math.max(10, right - left), h: Math.max(10, bottom - top) };
  }

  // Axis titles.
  if (hasValTitle || hasCatTitle) {
    const vtt = m.axisTitles.val, ctt = m.axisTitles.cat;
    const leftTitle = bar ? ctt : vtt, bottomTitle = bar ? vtt : ctt;
    const fmtOf = (axis: "cat" | "val") => m.format?.axisTitles?.[axis];
    if (leftTitle !== undefined) {
      const axis = bar ? "cat" : "val";
      const text = leftTitle || DEFAULT_AXIS_TITLE;
      const look = textLook(fmtOf(axis)?.text, { size: LOOK.axisTitleSize, color: LOOK.text });
      const x = area.left + PAD + 7.5, y = plot.y + plot.h / 2;
      const tw = c.measure(text, look.size);
      const box = { x: x - 9, y: y - tw / 2 - 3, w: 12, h: tw + 6 };
      axisTitleBox(c, fmtOf(axis), box, `axisTitle:${axis}`);
      c.prims.push({ t: "text", x, y, text, ...textAttrs(look), anchor: "middle", rotate: textRotation(fmtOf(axis)) ?? -90, part: `axisTitle:${axis}` });
      c.parts.push({ name: `axisTitle:${axis}`, box });
    }
    if (bottomTitle !== undefined) {
      const axis = bar ? "val" : "cat";
      const text = bottomTitle || DEFAULT_AXIS_TITLE;
      const look = textLook(fmtOf(axis)?.text, { size: LOOK.axisTitleSize, color: LOOK.text });
      const y = plot.y + plot.h + 14 + 15 + (hEnd ? -12 : 0);
      const tw = c.measure(text, look.size);
      const box = { x: plot.x + plot.w / 2 - tw / 2 - 3, y: y - 10, w: tw + 6, h: 13 };
      axisTitleBox(c, fmtOf(axis), box, `axisTitle:${axis}`);
      const rot = textRotation(fmtOf(axis));
      c.prims.push({ t: "text", x: plot.x + plot.w / 2, y, text, ...textAttrs(look), anchor: "middle", ...(rot ? { rotate: rot } : {}), part: `axisTitle:${axis}` });
      c.parts.push({ name: `axisTitle:${axis}`, box });
    }
  }

  // Zeichnungsfläche: its own Füllung, Rahmen and Effekte under everything inside it.
  const pfmt = m.format?.plot;
  const pf = fillPaint(pfmt?.fill, undefined), pl = linePaint(pfmt?.line, null);
  if ((pf && pf !== "none") || pl || pfmt?.effects) c.prims.push({ t: "rect", ...plot, fill: pf ?? "none", ...strokeOf(pl), filter: effectFilter(pfmt?.effects), part: "plot" });

  // The value axis (scatter: Y) in the plot's terms.
  const vFrac = (v: number) => { const p = scaleFrac(vScale, v); return va?.reverse ? 1 - p : p; };
  const vPos = (v: number) => (bar ? plot.x + vFrac(v) * plot.w : plot.y + plot.h - vFrac(v) * plot.h);
  const vMinPos = vPos(vScale.min), vMaxPos = vPos(vScale.max);
  // Where the category (scatter: X) axis crosses the value axis: Horizontale Achse schneidet.
  const crossVal = typeof va?.crosses === "number" ? va.crosses : va?.crosses === "max" ? vScale.max : vScale.log ? vScale.min : Math.min(Math.max(0, vScale.min), vScale.max);
  const catAxisAt = vPos(Math.min(Math.max(crossVal, vScale.min), vScale.max));
  const base = catAxisAt;

  // Gridlines (Haupt- and Hilfsgitternetzlinien of the value axis).
  drawGridlines(c, vScale, plot, bar, !!va?.reverse, va?.minorUnit);
  const vTicks = ticks(vScale);

  // Category positions.
  const n = Math.max(1, nCat);
  const catLen = bar ? plot.h : plot.w;
  const step = between ? catLen / n : n > 1 ? catLen / (n - 1) : catLen;
  const along = (p: number) => (bar ? plot.y + plot.h - p : plot.x + p);
  const catAt = (k: number) => { const j = xa?.reverse ? n - 1 - k : k; return along(between ? (j + 0.5) * step : n > 1 ? j * step : catLen / 2); };
  const catCenter = (i: number) => catAt(slotOf(i));
  /** Boundary k (0 … n) between categories (on-tick axes: the category itself). */
  const catEdge = (k: number) => {
    if (!between) return catAt(Math.min(Math.max(0, k), n - 1));
    const j = xa?.reverse ? n - k : k;
    return along(j * step);
  };

  let vPos2 = vPos, base2 = base;
  if (vScale2) {
    const s2 = vScale2;
    const f2 = (v: number) => { const p = scaleFrac(s2, v); return va2?.reverse ? 1 - p : p; };
    vPos2 = (v: number) => (bar ? plot.x + f2(v) * plot.w : plot.y + plot.h - f2(v) * plot.h);
    base2 = vPos2(s2.log ? s2.min : Math.min(Math.max(0, s2.min), s2.max));
    // Kurs with Volumen: the prices' axis has gridlines of its own (Excel).
    const grid2 = m.kind === "stockVHLC" || m.kind === "stockVOHLC";
    const lt2 = textLook(va2?.text, { size: LOOK.labelSize, color: LOOK.text });
    const at2 = bar ? plot.y : plot.x + plot.w;
    ticks(s2).forEach((v, i) => {
      const p = vPos2(v);
      if (grid2 && m.gridlines) c.prims.push(bar ? { t: "path", d: `M${f(p)} ${f(plot.y)}V${f(plot.y + plot.h)}`, stroke: LOOK.gridline, sw: 0.75, part: "grid" } : { t: "path", d: `M${f(plot.x)} ${f(p)}H${f(plot.x + plot.w)}`, stroke: LOOK.gridline, sw: 0.75, part: "grid" });
      if (vLabels2[i]) c.prims.push(bar ? { t: "text", x: p, y: plot.y - 5, text: vLabels2[i], ...textAttrs(lt2), anchor: "middle", part: "axis:val2" }
        : { t: "text", x: plot.x + plot.w + LABEL_GAP, y: p + 2.5, text: vLabels2[i], ...textAttrs(lt2), anchor: "start", part: "axis:val2" });
    });
    if (va2?.visible !== false) {
      axisLineAndTicks(c, { horizontal: bar, at: at2, from: bar ? plot.x : plot.y, to: bar ? plot.x + plot.w : plot.y + plot.h, line: linePaint(va2?.line, null),
        majors: ticks(s2).map(vPos2), minors: minorTicks(s2, va2?.minorUnit).map(vPos2), major: va2?.majorTick, minor: va2?.minorTick, out: bar ? -1 : 1, part: "axis:val2", effects: va2?.effects });
    }
  }

  // Value labels and the value axis: the axis where the category axis puts it (Vertikale Achse schneidet).
  let valAxisAt: number;
  let xs: AxisScale | undefined, xPos: ((v: number) => number) | undefined;
  if (scatter && xScale) {
    const xsc = xScale;
    xs = xsc;
    const xFrac = (v: number) => { const p = scaleFrac(xsc, v); return xa?.reverse ? 1 - p : p; };
    xPos = (v: number) => plot.x + xFrac(v) * plot.w;
    const cx = typeof xa?.crosses === "number" ? xa.crosses : xa?.crosses === "max" ? xsc.max : xsc.log ? xsc.min : Math.min(Math.max(0, xsc.min), xsc.max);
    valAxisAt = xPos(Math.min(Math.max(cx, xsc.min), xsc.max));
  } else {
    const k = xa?.crosses === "max" ? n : typeof xa?.crosses === "number" ? Math.min(n, Math.max(0, Math.round(xa.crosses) - 1)) : 0;
    valAxisAt = catEdge(k);
  }
  // Labels "Achsennah" stand at the axis, "Tief" / "Hoch" at the plot's ends.
  const vEdgeStart = bar ? plot.y + plot.h : plot.x, vEdgeEnd = bar ? plot.y : plot.x + plot.w;
  const valLabelAt = va?.labelPos === "low" ? (xa?.reverse ? vEdgeEnd : vEdgeStart) : va?.labelPos === "high" ? (xa?.reverse ? vEdgeStart : vEdgeEnd) : valAxisAt;
  const valOutward = Math.abs(valLabelAt - vEdgeEnd) < 0.5 ? 1 : -1; // +1: right (bars: up)
  const vt = vt0;
  // Benutzerdefinierter Winkel and Textrichtung of the value labels.
  const vb = va?.text?.box;
  const valRot = vb?.dir === "vert" ? 90 : vb?.dir === "vert270" ? -90 : vb?.rotation;
  const vr = valRot ? { rotate: valRot } : {};
  vTicks.forEach((v, i) => {
    const t = vLabels[i];
    if (!t) return;
    const p = vPos(v);
    if (bar) c.prims.push({ t: "text", x: p, y: valOutward > 0 ? valLabelAt - 5 : valLabelAt + 14 + (vt.size - LOOK.labelSize), text: t, ...textAttrs(vt), anchor: "middle", ...vr, part: "axis:val" });
    else c.prims.push({ t: "text", x: valLabelAt + (valOutward > 0 ? LABEL_GAP : -LABEL_GAP), y: p + 2.5, text: t, ...textAttrs(vt), anchor: valOutward > 0 ? "start" : "end", ...vr, part: "axis:val" });
  });
  if (unitLabel) {
    // Anzeigeeinheitenbeschriftung: beside the labels at the axis' far end.
    const w = Math.max(0, ...vLabels.map((t) => c.measure(t, vt.size)));
    if (bar) c.prims.push({ t: "text", x: plot.x + plot.w, y: plot.y + plot.h + 26, text: unitLabel, ...textAttrs(vt), anchor: "end", part: "axis:val" });
    else c.prims.push({ t: "text", x: valLabelAt + valOutward * (LABEL_GAP + w + 6), y: plot.y + 2, text: unitLabel, ...textAttrs(vt), anchor: "end", rotate: -90, part: "axis:val" });
  }
  if (va?.visible !== false) {
    const auto = scatter ? { color: kindInfo(m.kind).family === "bubble" ? LOOK.gridline : LOOK.valueAxisLine, width: 0.75 } : null;
    axisLineAndTicks(c, { horizontal: bar, at: valAxisAt, from: bar ? plot.x : plot.y, to: bar ? plot.x + plot.w : plot.y + plot.h, line: linePaint(va?.line, auto),
      majors: vTicks.map(vPos), minors: minorTicks(vScale, va?.minorUnit).map(vPos), major: va?.majorTick, minor: va?.minorTick,
      out: Math.abs(valAxisAt - vEdgeEnd) < 0.5 ? (bar ? -1 : 1) : (bar ? 1 : -1), part: "axis:val", effects: va?.effects });
  }
  void vMinPos; void vMaxPos;

  // The category axis' labels (scatter: the X values) at their place: Achsennah at the axis, Tief / Hoch at the ends.
  const hEdgeStart = bar ? plot.x : plot.y + plot.h, hEdgeEnd = bar ? plot.x + plot.w : plot.y;
  const minEdge = va?.reverse ? hEdgeEnd : hEdgeStart, maxEdge = va?.reverse ? hEdgeStart : hEdgeEnd;
  const catLabelAt = xa?.labelPos === "low" ? minEdge : xa?.labelPos === "high" ? maxEdge : catAxisAt;
  const catOutward = Math.abs(catLabelAt - hEdgeEnd) < 0.5 ? 1 : -1; // +1: above (bars: right)
  const ct = ct0;
  if (scatter && xs && xPos) {
    const xp = xPos;
    const xt = ticks(xs);
    xt.forEach((v, i) => {
      if (m.gridlines) c.prims.push({ t: "path", d: `M${f(xp(v))} ${f(plot.y)}V${f(plot.y + plot.h)}`, ...strokeOf(linePaint(m.format?.majorGrid?.line, { color: LOOK.gridline, width: 0.75 })), part: "grid" });
      if (xLabels[i]) c.prims.push({ t: "text", x: xp(v), y: catOutward > 0 ? catLabelAt - 5 : catLabelAt + 14 + (ct.size - LOOK.labelSize), text: xLabels[i], ...textAttrs(ct), anchor: "middle", part: "axis:cat" });
    });
    if (xa?.visible !== false) {
      axisLineAndTicks(c, { horizontal: true, at: catAxisAt, from: plot.x, to: plot.x + plot.w, line: linePaint(xa?.line, { color: kindInfo(m.kind).family === "bubble" ? LOOK.gridline : LOOK.valueAxisLine, width: 0.75 }),
        majors: xt.map(xp), minors: minorTicks(xs, xa?.minorUnit).map(xp), major: xa?.majorTick, minor: xa?.minorTick, out: Math.abs(catAxisAt - plot.y) < 0.5 ? -1 : 1, part: "axis:cat", effects: xa?.effects });
    }
    clipStart(c, plot);
    if (kindInfo(m.kind).family === "bubble") drawBubbles(c, plot, xp, vPos);
    else m.series.forEach((s, si) => drawScatterSeries(c, s, si, xp, vPos));
    m.series.forEach((s, si) => { if (s.errorBars) s.val.cache.forEach((v, i) => { const x = xOf(s, i); if (v !== null && x !== null) drawErrorBar(c, s, si, xp(x), v, vPos, false); }); });
    m.series.forEach((s, si) => { if (s.trendline) drawTrendline(c, s, si, plot, (v) => xp(v), vPos, (i) => xOf(s, i)); });
    clipEnd(c);
    return { plot, valueScale: vScale, xScale: xs };
  }

  // Vertikale Hauptgitternetzlinien of the category axis.
  if (m.format?.catGrid?.show) {
    const cg = linePaint(m.format.catGrid.line, { color: LOOK.gridline, width: 0.75 });
    if (cg) for (let k = 0; k <= (between ? n : n - 1); k++) {
      const p = between ? catEdge(k) : catAt(k);
      c.prims.push(bar ? { t: "path", d: `M${f(plot.x)} ${f(p)}H${f(plot.x + plot.w)}`, ...strokeOf(cg), part: "catGrid" }
        : { t: "path", d: `M${f(p)} ${f(plot.y)}V${f(plot.y + plot.h)}`, ...strokeOf(cg), part: "catGrid" });
    }
  }
  const maxCatW = Math.max(0, ...catLabels.map((t) => c.measure(t, ct.size)));
  const spacing = step * lblSkip;
  const rotate = catRot !== undefined ? catRot !== 0 : !bar && maxCatW > spacing - 2 && catLabels.filter(Boolean).length > 1;
  const rot = catRot ?? -45;
  // Abstand zur Achse: 2.92 pt per 100 (Excel's PDF export: baselines 10.83 / 13.75 / 25.45 pt below the axis at 0 / 100 / 500).
  const gap = 4.33 + 2.92 * offset;
  catLabels.forEach((t, i) => {
    if (!t) return;
    const p = catAt(i);
    if (bar) {
      const lg = LABEL_GAP - 3 + 3 * offset;
      const x = catOutward > 0 ? catLabelAt + lg : catLabelAt - lg;
      c.prims.push({ t: "text", x, y: p + 3.2, text: t, ...textAttrs(ct), anchor: catOutward > 0 ? "start" : "end", ...(catRot ? { rotate: catRot } : {}), part: "axis:cat" });
    } else if (rotate) {
      const y = catOutward > 0 ? catLabelAt - gap : catLabelAt + gap;
      c.prims.push({ t: "text", x: p + 3, y, text: t, ...textAttrs(ct), anchor: (rot < 0) === (catOutward < 0) ? "end" : "start", rotate: rot, part: "axis:cat" });
    } else {
      const y = catOutward > 0 ? catLabelAt - gap - 1.5 : catLabelAt + gap + 6.75 + (ct.size - LOOK.labelSize);
      c.prims.push({ t: "text", x: p, y, text: t, ...textAttrs(ct), anchor: "middle", part: "axis:cat" });
    }
  });
  if (xa?.visible !== false) {
    // Teilstriche of a category axis stand between the categories (Auf Teilstrichen: at them), every Intervall zwischen Teilstrichen.
    const skip = Math.max(1, Math.round(xa?.tickMarkSkip ?? 1));
    const majors: number[] = [];
    for (let k = 0; k <= (between ? n : n - 1); k += skip) majors.push(between ? catEdge(k) : catAt(k));
    if (between && (n % skip) !== 0) majors.push(catEdge(n));
    axisLineAndTicks(c, { horizontal: !bar, at: catAxisAt, from: bar ? plot.y : plot.x, to: bar ? plot.y + plot.h : plot.x + plot.w, line: linePaint(xa?.line, { color: LOOK.gridline, width: 0.75 }),
      majors, minors: [], major: xa?.majorTick, minor: undefined, out: Math.abs(catAxisAt - hEdgeEnd) < 0.5 ? (bar ? 1 : -1) : (bar ? -1 : 1), part: "axis:cat", effects: xa?.effects });
  }

  clipStart(c, plot);
  const ordered = [...groups].sort((a, b) => (Z_ORDER[kindInfo(a.kind).family] ?? 3) - (Z_ORDER[kindInfo(b.kind).family] ?? 3));
  for (const g of ordered) {
    const fam = kindInfo(g.kind).family;
    const vp = g.secondary ? vPos2 : vPos, bs = g.secondary ? base2 : base;
    if (fam === "bar") drawBars(c, n, step, catCenter, vp, bs, g);
    else if (fam === "area") drawAreas(c, catCenter, vp, bs, g);
    else if (fam === "stock") drawStock(c, step, catCenter, vp, g);
    else drawLines(c, catCenter, vp, g);
  }
  m.series.forEach((s, si) => {
    if (!s.errorBars || bar) return;
    // Clustered columns: the bar's middle; other kinds: the category's centre.
    const k = m.kind === "column" ? m.series.length : 1;
    const gp = (m.gapWidth ?? 219) / 100, ov = m.overlap !== undefined ? m.overlap / 100 : -0.27;
    const bw = step / (k - (k - 1) * ov + gp);
    s.val.cache.forEach((v, i) => {
      if (v === null || !isFinite(v)) return;
      const x = m.kind === "column" ? catCenter(i) - step / 2 + (bw * gp) / 2 + si * bw * (1 - ov) + bw / 2 : catCenter(i);
      drawErrorBar(c, s, si, x, v, vPos, false);
    });
  });
  m.series.forEach((s, si) => {
    if (s.trendline) drawTrendline(c, s, si, plot, (x) => (bar ? 0 : catCenter(x - 1)), vPos, (i) => i + 1, bar);
  });
  clipEnd(c);
  return { plot, valueScale: vScale };
}

/** An axis title's own Füllung, Rahmen and Effekte. */
function axisTitleBox(c: Ctx, fmt: { fill?: import("./model").Fill; line?: import("./model").Line; effects?: Effects } | undefined, box: Box, part: string): void {
  const fl = fillPaint(fmt?.fill, undefined), ln = linePaint(fmt?.line, null);
  if ((fl && fl !== "none") || ln || fmt?.effects) c.prims.push({ t: "rect", ...box, fill: fl ?? "none", ...strokeOf(ln), filter: effectFilter(fmt?.effects), part });
}

/** Haupt- and Hilfsgitternetzlinien of a value axis (logarithmic axes: at the powers and their multiples). */
function drawGridlines(c: Ctx, scale: AxisScale, plot: Box, horizontal: boolean, reverse: boolean, minorUnit?: number): void {
  const fmt = c.m.format;
  const frac = (v: number) => { const p = scaleFrac(scale, v); return reverse ? 1 - p : p; };
  const at = (v: number) => (horizontal ? plot.x + frac(v) * plot.w : plot.y + plot.h - frac(v) * plot.h);
  const line = (p: number) => (horizontal ? `M${f(p)} ${f(plot.y)}V${f(plot.y + plot.h)}` : `M${f(plot.x)} ${f(p)}H${f(plot.x + plot.w)}`);
  const minor = fmt?.minorGrid?.show ? linePaint(fmt.minorGrid.line, { color: "#F2F2F2", width: 0.75 }) : null;
  if (minor) {
    const fx = effectFilter(fmt?.minorGrid?.effects);
    for (const v of minorTicks(scale, minorUnit)) c.prims.push({ t: "path", d: line(at(v)), ...strokeOf(minor), ...(fx ? { filter: fx } : {}), part: "minorGrid" });
  }
  const grid = linePaint(fmt?.majorGrid?.line, { color: LOOK.gridline, width: 0.75 });
  if (c.m.gridlines && grid) {
    const fx = effectFilter(fmt?.majorGrid?.effects);
    for (const v of ticks(scale)) c.prims.push({ t: "path", d: line(at(v)), ...strokeOf(grid), ...(fx ? { filter: fx } : {}), part: "grid" });
  }
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

export function clipStart(c: Ctx, plot: Box): void {
  c.prims.push({ t: "rect", x: plot.x - 4, y: plot.y - 4, w: plot.w + 8, h: plot.h + 8, part: "clip:start" });
}
export function clipEnd(c: Ctx): void {
  c.prims.push({ t: "rect", x: 0, y: 0, w: 0, h: 0, part: "clip:end" });
}

/** Kurs: Höchst-Tiefst-Linien between the highest and lowest price, up/down bars from the open to the close, the close as a dot. */
function drawStock(c: Ctx, step: number, catCenter: (i: number) => number, vPos: (v: number) => number, g: ChartGroup): void {
  const { m } = c;
  const list = g.idx.map((i) => m.series[i]).filter(Boolean);
  const n = Math.max(0, ...list.map((s) => s.val.cache.length));
  const open = m.kind === "stockOHLC" || m.kind === "stockVOHLC";
  const bw = step / 2.5;
  for (let i = 0; i < n; i++) {
    const vals = list.map((s) => s.val.cache[i]).filter((v): v is number => v !== null && v !== undefined && isFinite(v));
    if (!vals.length) continue;
    const x = catCenter(i);
    c.prims.push({ t: "path", d: `M${f(x)} ${f(vPos(Math.max(...vals)))}V${f(vPos(Math.min(...vals)))}`, stroke: "#404040", sw: 0.75, part: "hiLo" });
    if (open) {
      const o = list[0].val.cache[i], cl = list.at(-1)!.val.cache[i];
      if (o === null || cl === null || o === undefined || cl === undefined) continue;
      const up = cl >= o;
      const y0 = vPos(Math.max(o, cl)), y1 = vPos(Math.min(o, cl));
      c.prims.push({ t: "rect", x: x - bw / 2, y: y0, w: bw, h: Math.max(0.75, y1 - y0), fill: up ? "#FFFFFF" : "#595959", stroke: "#595959", sw: 0.75, part: up ? "upBars" : "downBars" });
    }
  }
  if (!open) {
    const si = g.idx.at(-1)!, s = m.series[si];
    // The close: Excel's "dot" marker of Kurs charts, a short dash.
    s?.val.cache.forEach((v, i) => {
      if (v === null || !isFinite(v)) return;
      if (s.marker?.symbol && s.marker.symbol !== "dot") { drawMarker(c, s, si, i, catCenter(i), vPos(v), "dot"); return; }
      c.prims.push({ t: "path", d: `M${f(catCenter(i) - 2)} ${f(vPos(v))}H${f(catCenter(i) + 2)}`, stroke: seriesPaint(c, si), sw: 1.5, part: `series:${si}:${i}` });
    });
  }
  if (m.dataLabels) g.idx.forEach((si) => m.series[si].val.cache.forEach((v, i) => { if (v !== null) label(c, m.series[si], si, i, v, catCenter(i) + 5, vPos(v) + 3.2, "start"); }));
}

/** Blase: circles whose area follows the Blasengröße (the largest a quarter of the plot's smaller side), 3-D bubbles shaded. */
function drawBubbles(c: Ctx, plot: Box, xPos: (v: number) => number, vPos: (v: number) => number): void {
  const { m } = c;
  const sizes = m.series.flatMap((s) => (s.size?.cache ?? []).filter((v): v is number => v !== null && isFinite(v)).map(Math.abs));
  const smax = Math.max(0, ...sizes) || 1;
  const dmax = 0.25 * Math.min(plot.w, plot.h) * ((m.bubbleScale ?? 100) / 100);
  const d3 = !!kindInfo(m.kind).d3;
  m.series.forEach((s, si) => {
    const color = seriesPaint(c, si);
    const fill = fillPaint(s.fill, d3 ? color : withAlpha(color, 25)) ?? color;
    s.val.cache.forEach((v, i) => {
      const x = xOf(s, i), sz = s.size?.cache[i];
      if (v === null || x === null || sz === null || sz === undefined || sz <= 0) return;
      const r = (dmax * Math.sqrt(Math.abs(sz) / smax)) / 2;
      const cx = xPos(x), cy = vPos(v);
      c.prims.push({ t: "circle", cx, cy, r, fill, part: `series:${si}:${i}` });
      if (d3) {
        // Excel's 3-D effect: darker at the rim, a light spot up left.
        c.prims.push({ t: "circle", cx, cy, r: r * 0.97, fill: "none", stroke: lum(color, 0.6, 0), sw: r * 0.12, part: `series:${si}:${i}` });
        c.prims.push({ t: "circle", cx: cx - r * 0.3, cy: cy - r * 0.35, r: r * 0.35, fill: withAlpha("#FFFFFF", 55), part: `series:${si}:${i}` });
      }
      if (m.dataLabels) label(c, s, si, i, v, cx, cy + 3.2, "middle");
    });
  });
}

/** The values a 100 % stacked group shows: each point's share of its category's total. */
function shares(m: ChartModel, g: ChartGroup): Map<number, (number | null)[]> {
  const out = new Map<number, (number | null)[]>();
  const n = Math.max(0, ...g.idx.map((i) => m.series[i]?.val.cache.length ?? 0));
  const tot = Array.from({ length: n }, (_, k) => g.idx.reduce((a, i) => a + Math.abs(m.series[i]?.val.cache[k] ?? 0), 0));
  for (const i of g.idx) out.set(i, (m.series[i]?.val.cache ?? []).map((v, k) => (v === null || !tot[k] ? v : v / tot[k])));
  return out;
}

function drawBars(c: Ctx, n: number, step: number, catCenter: (i: number) => number, vPos: (v: number) => number, base: number, g: ChartGroup): void {
  const { m } = c;
  const info = kindInfo(g.kind);
  const bar = info.dir === "bar", stacked = info.grouping === "stacked" || info.grouping === "percentStacked";
  const d = barDefaults(g.kind);
  const gap = (m.gapWidth ?? d.gap) / 100;
  const overlap = m.overlap !== undefined ? m.overlap / 100 : (d.overlap ?? 0) / 100;
  const k = stacked ? 1 : g.idx.length || 1;
  const bw = step / (k - (k - 1) * overlap + gap);
  const pos = new Array(n).fill(0), neg = new Array(n).fill(0);
  const pct = info.grouping === "percentStacked" ? shares(m, g) : null;
  g.idx.forEach((si, gi) => {
    const s = m.series[si];
    const vals = pct?.get(si) ?? s.val.cache;
    vals.forEach((v, i) => {
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
      const off = stacked ? 0 : gi * bw * (1 - overlap);
      let rect: Box;
      if (bar) {
        const y1 = start - off - bw;
        rect = { x: Math.min(from, to), y: y1, w: Math.abs(to - from), h: bw };
      } else {
        const x1 = start + off;
        rect = { x: x1, y: Math.min(from, to), w: bw, h: Math.abs(to - from) };
      }
      // Datenpunkt formatieren: a point's own Füllung, Rahmen and Effekte before the series'.
      const pt = pointOf(s, i);
      const ln = linePaint(pt?.line ?? s.line, null);
      const inv = v < 0 && (pt?.invertIfNegative ?? s.invertIfNegative);
      const fill = inv ? "#FFFFFF" : fillPaint(pt?.fill ?? s.fill, autoColor(c, si, i));
      const fx = effectFilter(pt?.effects ?? s.effects);
      c.prims.push({ t: "rect", ...rect, fill, ...strokeOf(ln), stroke: ln?.stroke ?? (inv ? c.color(si) : undefined), sw: ln?.sw ?? 0.75, ...(fx ? { filter: fx } : {}), part: `series:${si}:${i}` });
      if (m.dataLabels) {
        // Beschriftungsposition: Außerhalb Ende by default (stacked: Zentriert), or the chosen one.
        const lpos = m.format?.dataLabels?.pos ?? (stacked ? "ctr" : "outEnd");
        const neg = v < 0;
        const shown = s.val.cache[i] ?? v;
        if (bar) {
          const end = neg ? rect.x : rect.x + rect.w, startX = neg ? rect.x + rect.w : rect.x;
          const x = lpos === "ctr" ? rect.x + rect.w / 2 : lpos === "inEnd" ? end + (neg ? 3 : -3) : lpos === "inBase" ? startX + (neg ? -3 : 3) : end + (neg ? -3 : 3);
          const anchor = lpos === "ctr" ? "middle" : (lpos === "inEnd") !== neg ? "end" : "start";
          label(c, s, si, i, shown, x, rect.y + rect.h / 2 + 3.2, anchor as "start" | "middle" | "end");
        } else {
          const top = neg ? rect.y + rect.h : rect.y, bottom = neg ? rect.y : rect.y + rect.h;
          const y = lpos === "ctr" ? rect.y + rect.h / 2 + 3.2 : lpos === "inEnd" ? (neg ? top - 3.4 : top + 9.6) : lpos === "inBase" ? (neg ? bottom + 9.6 : bottom - 3.4) : neg ? top + 9.6 : top - 3.4;
          label(c, s, si, i, shown, rect.x + rect.w / 2, y, "middle");
        }
      }
    });
  });
}

/** The values of a stacked group as drawn: each series on top of the ones before (100 %: as shares). */
function stackedValues(m: ChartModel, g: ChartGroup): Map<number, (number | null)[]> {
  const grouping = kindInfo(g.kind).grouping;
  const src = grouping === "percentStacked" ? shares(m, g) : null;
  const out = new Map<number, (number | null)[]>();
  if (grouping !== "stacked" && grouping !== "percentStacked") {
    for (const i of g.idx) out.set(i, m.series[i].val.cache);
    return out;
  }
  const acc: number[] = [];
  for (const i of g.idx) {
    const vals = src?.get(i) ?? m.series[i].val.cache;
    out.set(i, vals.map((v, k) => { acc[k] = (acc[k] ?? 0) + (v ?? 0); return v === null ? null : acc[k]; }));
  }
  return out;
}

function drawAreas(c: Ctx, catCenter: (i: number) => number, vPos: (v: number) => number, base: number, g: ChartGroup): void {
  const stacked = kindInfo(g.kind).grouping === "stacked" || kindInfo(g.kind).grouping === "percentStacked";
  const tops = stackedValues(c.m, g);
  let below: (readonly [number, number])[] | null = null;
  for (const si of g.idx) {
    const s = c.m.series[si];
    const vals = tops.get(si) ?? s.val.cache;
    const pts = vals.map((v, i) => [catCenter(i), vPos(v ?? 0)] as const);
    if (!pts.length) continue;
    const d = stacked && below
      ? `M${f(pts[0][0])} ${f(pts[0][1])}` + pts.slice(1).map(([x, y]) => `L${f(x)} ${f(y)}`).join("") + [...below].reverse().map(([x, y]) => `L${f(x)} ${f(y)}`).join("") + "Z"
      : `M${f(pts[0][0])} ${f(base)}` + pts.map(([x, y]) => `L${f(x)} ${f(y)}`).join("") + `L${f(pts.at(-1)![0])} ${f(base)}Z`;
    const ln = linePaint(s.line, null);
    const fx = effectFilter(s.effects);
    c.prims.push({ t: "path", d, fill: fillPaint(s.fill, c.color(si)), ...strokeOf(ln), ...(fx ? { filter: fx } : {}), part: `series:${si}` });
    const under = below;
    if (c.m.dataLabels) s.val.cache.forEach((v, i) => {
      if (v !== null) label(c, s, si, i, v, catCenter(i), (pts[i][1] + (under ? under[i][1] : base)) / 2 + 3.2, "middle");
    });
    below = pts;
  }
}

function drawLines(c: Ctx, catCenter: (i: number) => number, vPos: (v: number) => number, g: ChartGroup): void {
  const markers = !!kindInfo(g.kind).markers;
  const tops = stackedValues(c.m, g);
  for (const si of g.idx) {
    const s = c.m.series[si];
    const vals = tops.get(si) ?? s.val.cache;
    const color = c.color(si);
    const ln = linePaint(s.line, { color, width: 2.25 });
    // Runs of points without gaps, drawn straight or smoothed (Linie glätten).
    const runs: [number, number][][] = [[]];
    vals.forEach((v, i) => {
      if (v === null || !isFinite(v)) { if (runs.at(-1)!.length) runs.push([]); return; }
      runs.at(-1)!.push([catCenter(i), vPos(v)]);
    });
    const d = runs.filter((r) => r.length).map((r) => (s.smooth && r.length > 2 ? smoothPath(r) : r.map(([x, y], i) => `${i ? "L" : "M"}${f(x)} ${f(y)}`).join(""))).join("");
    const fx = effectFilter(s.effects);
    if (d && ln) c.prims.push({ t: "path", d, ...strokeOf(ln), cap: ln.cap ?? "round", ...(fx ? { filter: fx } : {}), part: `series:${si}` });
    vals.forEach((v, i) => {
      if (v === null || !isFinite(v)) return;
      drawMarker(c, s, si, i, catCenter(i), vPos(v), markers ? "circle" : "none");
      if (c.m.dataLabels) label(c, s, si, i, s.val.cache[i] ?? v, catCenter(i) + 5, vPos(v) + 3.2, "start");
    });
  }
}

/** A series' Markierung at a point: its own symbol, size, Füllung and Rahmen, else the chart kind's. */
export function drawMarker(c: Ctx, s: ChartSeries, si: number, i: number, x: number, y: number, auto: MarkerSymbol): void {
  const own = pointOf(s, i)?.marker;
  const mk = own ? { ...s.marker, ...own } : s.marker;
  const sym = !mk?.symbol || mk.symbol === "auto" ? auto : mk.symbol;
  if (sym === "none") return;
  const color = seriesPaint(c, si);
  const r = (mk?.size ?? 5) / 2;
  const fill = fillPaint(mk?.fill, color) ?? color;
  const ln = linePaint(mk?.line, { color, width: 0.75 });
  const fx = effectFilter(pointOf(s, i)?.effects ?? s.effects);
  const st = { ...strokeOf(ln), ...(fx ? { filter: fx } : {}), part: `series:${si}:${i}` };
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
export function label(c: Ctx, s: ChartSeries, si: number, i: number, v: number, x: number, y: number, anchor: "start" | "middle" | "end", share?: number): void {
  const o = c.m.format?.dataLabels;
  const pie = isPie(c.m.kind);
  const parts: string[] = [];
  // Wert aus Zellen comes first, then Reihenname, Rubrikenname, Wert, Prozentsatz, Blasengröße (Excel's order).
  const cell = o?.fromCells?.cache[i];
  if (cell) parts.push(cell);
  if (o?.showSer) parts.push(seriesName(s, si));
  if (o?.showCat) parts.push(catText(c, s, i));
  if (o?.showVal ?? true) parts.push(c.fmt(v, o?.numFmt ?? s.val.format));
  if (o?.showPercent && pie && share !== undefined) parts.push(`${Math.round(share * 100)}%`);
  const size = s.size?.cache[i];
  if (o?.showBubble && size !== undefined && size !== null) parts.push(c.fmt(size, o?.numFmt ?? s.size?.format));
  if (!parts.length) return;
  const look = textLook(o?.text, { size: LOOK.labelSize, color: LOOK.dataLabel });
  const lines = (o?.separator ?? "; ").includes("\n") ? parts : [parts.join(o?.separator ?? "; ")];
  const rot = textRotation(o);
  // The label's own Füllung, Rahmen and Effekte as a box around its text.
  const fl = fillPaint(o?.fill, undefined), ln = linePaint(o?.line, null);
  if ((fl && fl !== "none") || ln || o?.effects) {
    const w = Math.max(...lines.map((t) => c.measure(t, look.size))) + 6, h = lines.length * (look.size + 2) + 3;
    const x0 = anchor === "start" ? x - 3 : anchor === "end" ? x - w + 3 : x - w / 2;
    c.prims.push({ t: "rect", x: x0, y: y - look.size - 1, w, h, fill: fl ?? "none", ...strokeOf(ln), filter: effectFilter(o?.effects), part: `labels:${si}` });
  }
  lines.forEach((t, k) => c.prims.push({ t: "text", x, y: y + k * (look.size + 2), text: t, ...textAttrs(look), anchor, ...(rot ? { rotate: rot } : {}), part: `labels:${si}` }));
}

/** A data point's own format (Datenpunkt formatieren). */
export function pointOf(s: ChartSeries | undefined, i: number): import("./model").DataPoint | undefined {
  return s?.points?.find((p) => p.idx === i);
}

/** A point's automatic colour: the series', or the accents in turn with Punkte nach Farbe unterscheiden. */
function autoColor(c: Ctx, si: number, i: number): string {
  const m = c.m;
  if (m.varyColors && m.series.length === 1 && !m.series[0].color) return paletteColor(m, i, m.series[0].val.cache.length, c.accents);
  return c.color(si);
}

function drawScatterSeries(c: Ctx, s: ChartSeries, si: number, xPos: (v: number) => number, vPos: (v: number) => number): void {
  const color = c.color(si);
  const pts: [number, number][] = [];
  s.val.cache.forEach((v, i) => {
    const x = xOf(s, i);
    if (v === null || x === null || !isFinite(v)) return;
    pts.push([xPos(x), vPos(v)]);
  });
  const info = kindInfo(c.m.kind);
  const lineOn = info.lines !== "none" || (s.line !== undefined && s.line.type === "solid");
  const ln = lineOn ? linePaint(s.line, { color, width: 1.5 }) : null;
  if (ln && pts.length > 1) {
    const d = info.lines === "smooth" || s.smooth ? smoothPath(pts) : pts.map(([x, y], i) => `${i ? "L" : "M"}${f(x)} ${f(y)}`).join("");
    const fx = effectFilter(s.effects);
    c.prims.push({ t: "path", d, ...strokeOf(ln), cap: ln.cap ?? "round", ...(fx ? { filter: fx } : {}), part: `series:${si}` });
  }
  pts.forEach(([x, y], i) => drawMarker(c, s, si, i, x, y, info.markers ? "circle" : "none"));
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

export function f(v: number): string {
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
  const defs: string[] = [...(scene.defs ?? [])];
  let clipId = 0;
  const uniq = Math.random().toString(36).slice(2, 7);
  const part = (p: { part?: string }) => (p.part ? ` data-part="${esc(p.part)}"` : "");
  const geometry = (p: Exclude<Prim, { t: "text" }>) => {
    switch (p.t) {
      case "rect": return `<rect x="${f(p.x)}" y="${f(p.y)}" width="${f(Math.max(0, p.w))}" height="${f(Math.max(0, p.h))}"${p.rx ? ` rx="${f(p.rx)}"` : ""}`;
      case "path": return `<path d="${p.d}"`;
      case "circle": return `<circle cx="${f(p.cx)}" cy="${f(p.cy)}" r="${f(p.r)}"`;
    }
  };
  const strokeAttrs = (p: Exclude<Prim, { t: "text" }>, sw?: number) => `${p.stroke ? ` stroke="${p.stroke}" stroke-width="${f(sw ?? p.sw ?? 0.75)}" stroke-linejoin="${p.join ?? "round"}"${p.join === "miter" ? ' stroke-miterlimit="8"' : ""}` : ""}${p.cap ? ` stroke-linecap="${p.cap === "square" ? "square" : p.cap}"` : ""}${p.dash ? ` stroke-dasharray="${p.dash}"` : ""}`;
  for (const p of scene.prims) {
    if (p.t === "rect" && p.part === "clip:start") {
      const id = `rc${++clipId}${uniq}`;
      out.push(`<clipPath id="${id}"><rect x="${f(p.x)}" y="${f(p.y)}" width="${f(p.w)}" height="${f(p.h)}"/></clipPath><g clip-path="url(#${id})">`);
      continue;
    }
    if (p.t === "rect" && p.part === "clip:end") { out.push("</g>"); continue; }
    if (p.t === "text") {
      const tr = p.rotate ? ` transform="rotate(${p.rotate} ${f(p.x)} ${f(p.y)})"` : "";
      const deco = [p.underline ? "underline" : "", p.strike ? "line-through" : ""].filter(Boolean).join(" ");
      const wt = `${p.bold ? ' font-weight="bold"' : ""}${p.italic ? ' font-style="italic"' : ""}${deco ? ` text-decoration="${deco}"` : ""}${p.font ? ` font-family="${esc(p.font)}"` : ""}`;
      const o = p.outline;
      const ol = o?.stroke ? ` stroke="${o.stroke}" stroke-width="${f(o.sw ?? 0.75)}" paint-order="stroke"${o.dash ? ` stroke-dasharray="${o.dash}"` : ""}${o.join ? ` stroke-linejoin="${o.join}"` : ""}${o.cap ? ` stroke-linecap="${o.cap}"` : ""}` : "";
      out.push(`<text x="${f(p.x)}" y="${f(p.y)}" font-size="${p.size}" fill="${p.fill}" text-anchor="${p.anchor}"${wt}${ol}${p.filter ? ` filter="${p.filter}"` : ""}${tr}${part(p)}>${esc(p.text)}</text>`);
      if (p.reflection) {
        // Mirrored at the line just below the text's descent, fading with its transparency.
        const axis = p.y + p.size * 0.22 + p.reflection.dist;
        const mirror = `matrix(1 0 0 -1 0 ${f(2 * axis)})${p.rotate ? ` rotate(${p.rotate} ${f(p.x)} ${f(p.y)})` : ""}`;
        out.push(`<text x="${f(p.x)}" y="${f(p.y)}" font-size="${p.size}" fill="${p.fill}" text-anchor="${p.anchor}"${wt} opacity="${p.reflection.opacity * 0.5}" transform="${mirror}"${part(p)}>${esc(p.text)}</text>`);
      }
      continue;
    }
    const filter = p.filter ? ` filter="${p.filter}"` : "";
    const markers = `${p.head ? ` marker-start="${p.head}"` : ""}${p.tail ? ` marker-end="${p.tail}"` : ""}`;
    if (p.compound && p.stroke) {
      // Verbundtyp: the stroke through a mask that cuts its gaps (two or three lines along the path).
      const id = `cm${++clipId}${uniq}`;
      const w = p.sw ?? 0.75;
      const cuts: [string, number][] = p.compound === "tri" ? [["white", w], ["black", w * 0.6], ["white", w * 0.2]]
        : p.compound === "thickThin" || p.compound === "thinThick" ? [["white", w], ["black", w * 0.35]] : [["white", w], ["black", w / 3]];
      const g = geometry(p);
      defs.push(`<mask id="${id}" maskUnits="userSpaceOnUse" x="-10000" y="-10000" width="20000" height="20000">${cuts.map(([c, sw]) => `${g} fill="none" stroke="${c}" stroke-width="${f(sw)}" stroke-linejoin="${p.join ?? "round"}"/>`).join("")}</mask>`);
      out.push(`<g${filter}${part(p)}>${p.fill && p.fill !== "none" ? `${g} fill="${p.fill}"/>` : ""}${g} fill="none"${strokeAttrs(p)} mask="url(#${id})"${markers}/></g>`);
      continue;
    }
    out.push(`${geometry(p)} fill="${p.fill ?? "none"}"${strokeAttrs(p)}${markers}${filter}${part(p)}/>`);
  }
  return (defs.length ? `<defs>${defs.join("")}</defs>` : "") + out.join("");
}
