// The chart model shared by RaabeCalc, RaabePresent and RaabeWrite (Step 7 E, D-061). It follows
// DrawingML's chart part (c:chartSpace) closely so that reading and writing stay simple: series
// refer to cells in the file's form ("Tabelle1!$B$2:$B$6") and carry the values last read
// (Excel's caches), so a chart draws without its workbook (Present, Write).

export type ChartKind =
  | "column" | "columnStacked" | "bar" | "barStacked" | "line" | "lineMarkers" | "area"
  | "pie" | "doughnut" | "scatter" | "scatterSmooth" | "scatterLines" | "histogram";

/** Legend position as in the file (c:legendPos), or no legend. */
export type LegendPos = "b" | "t" | "l" | "r" | "none";

export interface ChartData<T> {
  /** Cell reference in the file's form ("Tabelle1!$A$2:$A$6"); undefined for literal data. */
  ref?: string;
  /** The values last read from the cells (or literal values). */
  cache: T[];
  /** Number format of the first cell (file form), for axis and data labels. */
  format?: string;
}

export interface Trendline {
  type: "linear";
  /** Formel im Diagramm anzeigen. */
  showEq: boolean;
  /** Bestimmtheitsmaß im Diagramm anzeigen. */
  showR2: boolean;
}

export interface ChartSeries {
  name?: { ref?: string; text?: string };
  /** Categories; X values of scatter charts. */
  cat?: ChartData<string | number | null>;
  val: ChartData<number | null>;
  /** Fill/line colour "#RRGGBB"; undefined: the theme's accent for the series' index. */
  color?: string;
  trendline?: Trendline;
}

/** A corner of the chart: a cell and the offset into it in px at 100 % (twoCellAnchor). */
export interface CellAnchor { col: number; row: number; x: number; y: number }

export interface ChartModel {
  kind: ChartKind;
  series: ChartSeries[];
  /** Diagrammtitel; text undefined: automatic (the series name for one series, else "Diagrammtitel"). */
  title: { show: boolean; text?: string };
  /** Achsentitel; an axis title is shown when its text is set. */
  axisTitles: { cat?: string; val?: string };
  legend: LegendPos;
  /** Datenbeschriftungen. */
  dataLabels: boolean;
  /** Primäre horizontale Hauptgitternetzlinien (value gridlines). */
  gridlines: boolean;
  from: CellAnchor;
  to: CellAnchor;
  /** Name of the drawing object ("Diagramm 1"). */
  name?: string;
  /** File parts the model does not cover, kept by the .xlsx layer. */
  raw?: unknown;
}

/** The Office 2013-2022 theme's accents (Excel's default series colours). */
export const OFFICE_ACCENTS = ["#4472C4", "#ED7D31", "#A5A5A5", "#FFC000", "#5B9BD5", "#70AD47"];

/** Excel's default look (Office 2013-2022 chart style 201), in pt. */
export const LOOK = {
  font: "Calibri, Carlito, 'Segoe UI', sans-serif",
  text: "#595959",
  dataLabel: "#404040",
  titleSize: 14,
  labelSize: 9,
  axisTitleSize: 10,
  gridline: "#D9D9D9",
  border: "#D9D9D9",
  /** Axis lines of scatter charts. */
  valueAxisLine: "#BFBFBF",
  /** Default chart size: 360 x 216 pt = 480 x 288 px (12.7 x 7.62 cm). */
  width: 360,
  height: 216,
};

export const DEFAULT_TITLE = "Diagrammtitel";
export const DEFAULT_AXIS_TITLE = "Achsentitel";

/** Series colour i: the accents, then darker and lighter variants as Excel's "farbig" palette. */
export function seriesColor(i: number, accents: string[] = OFFICE_ACCENTS): string {
  const base = accents[i % 6] ?? OFFICE_ACCENTS[i % 6];
  const round = Math.floor(i / 6);
  if (round === 0) return base;
  // Excel: lumMod 60 %, then lumMod 80 % + lumOff 20 %, then lumMod 80 %, lumMod 60 % + lumOff 40 % …
  const steps: [number, number][] = [[0.6, 0], [0.8, 0.2], [0.8, 0], [0.6, 0.4], [0.5, 0], [0.7, 0.3]];
  const [mod, off] = steps[(round - 1) % steps.length];
  return lum(base, mod, off);
}

/** DrawingML's lumMod/lumOff on a colour (in HSL lightness). */
export function lum(hex: string, mod: number, off: number): string {
  const [r, g, b] = [1, 3, 5].map((i) => parseInt(hex.slice(i, i + 2), 16) / 255);
  const max = Math.max(r, g, b), min = Math.min(r, g, b);
  let h = 0, s = 0;
  const l = (max + min) / 2;
  if (max !== min) {
    const d = max - min;
    s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
    h = max === r ? (g - b) / d + (g < b ? 6 : 0) : max === g ? (b - r) / d + 2 : (r - g) / d + 4;
    h /= 6;
  }
  const l2 = Math.min(1, Math.max(0, l * mod + off));
  const hue = (p: number, q: number, t: number) => {
    if (t < 0) t += 1;
    if (t > 1) t -= 1;
    if (t < 1 / 6) return p + (q - p) * 6 * t;
    if (t < 1 / 2) return q;
    if (t < 2 / 3) return p + (q - p) * (2 / 3 - t) * 6;
    return p;
  };
  let rgb: number[];
  if (s === 0) rgb = [l2, l2, l2];
  else {
    const q = l2 < 0.5 ? l2 * (1 + s) : l2 + s - l2 * s, p = 2 * l2 - q;
    rgb = [hue(p, q, h + 1 / 3), hue(p, q, h), hue(p, q, h - 1 / 3)];
  }
  return "#" + rgb.map((v) => Math.round(v * 255).toString(16).padStart(2, "0")).join("").toUpperCase();
}

export const isPie = (k: ChartKind) => k === "pie" || k === "doughnut";
export const isScatter = (k: ChartKind) => k === "scatter" || k === "scatterSmooth" || k === "scatterLines";
export const isBar = (k: ChartKind) => k === "bar" || k === "barStacked";
export const isStacked = (k: ChartKind) => k === "columnStacked" || k === "barStacked";

/** The title as shown: its own text, else the series name for one series (pies: the first), else "Diagrammtitel". */
export function titleText(m: ChartModel): string {
  if (m.title.text !== undefined) return m.title.text;
  if (m.kind === "histogram") return DEFAULT_TITLE;
  const one = isPie(m.kind) ? m.series[0] : m.series.length === 1 ? m.series[0] : undefined;
  return one?.name?.text || DEFAULT_TITLE;
}

/** The legend Excel gives a new chart: pies always (bottom), others with more than one series. */
export function defaultLegend(kind: ChartKind, seriesCount: number): LegendPos {
  if (kind === "histogram") return "none";
  if (isPie(kind)) return "b";
  return seriesCount > 1 ? "b" : "none";
}

/** A new chart of `kind` with Excel's defaults. */
export function newChart(kind: ChartKind, series: ChartSeries[], from: CellAnchor, to: CellAnchor): ChartModel {
  return {
    kind, series, from, to,
    title: { show: true },
    axisTitles: {},
    legend: defaultLegend(kind, series.length),
    dataLabels: false,
    gridlines: true,
  };
}
