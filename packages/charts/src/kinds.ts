// What each chart type is made of (D-086): its family (the DrawingML chart group or the chartex
// layout it is written as), the direction of bars, the grouping, whether it is drawn in 3-D, and
// Excel's xlChartType number (COM's Chart.ChartType). The names are Excel's German ones as
// Einfügen > Diagramme and Diagrammtyp ändern > Alle Diagramme show them (results/2026-11-05-excel-inventory.json).

import type { ChartKind, ChartModel, View3D } from "./model";

export type Family =
  | "bar" | "line" | "area" | "pie" | "doughnut" | "ofPie" | "scatter" | "bubble" | "radar" | "stock" | "surface"
  | "histogram" | "waterfall" | "funnel" | "treemap" | "sunburst" | "boxWhisker" | "pareto";

export type Grouping = "standard" | "clustered" | "stacked" | "percentStacked";

export interface KindInfo {
  family: Family;
  /** Bars: columns ("col") or horizontal bars ("bar"). */
  dir?: "col" | "bar";
  grouping?: Grouping;
  d3?: boolean;
  /** Lines and radar: markers on the points; scatter: markers drawn. */
  markers?: boolean;
  /** Scatter: lines between the points, smoothed or straight. */
  lines?: "none" | "straight" | "smooth";
  /** Excel's xlChartType. */
  xl: number;
  /** Excel's German name. */
  label: string;
}

const K = (family: Family, xl: number, label: string, o: Partial<KindInfo> = {}): KindInfo => ({ family, xl, label, ...o });

export const KINDS: Record<ChartKind, KindInfo> = {
  column: K("bar", 51, "Gruppierte Säulen", { dir: "col", grouping: "clustered" }),
  columnStacked: K("bar", 52, "Gestapelte Säulen", { dir: "col", grouping: "stacked" }),
  columnStacked100: K("bar", 53, "Gestapelte Säulen (100 %)", { dir: "col", grouping: "percentStacked" }),
  column3DClustered: K("bar", 54, "3D-Säulen (gruppiert)", { dir: "col", grouping: "clustered", d3: true }),
  column3DStacked: K("bar", 55, "Gestapelte 3D-Säulen", { dir: "col", grouping: "stacked", d3: true }),
  column3DStacked100: K("bar", 56, "Gestapelte 3D-Säulen (100 %)", { dir: "col", grouping: "percentStacked", d3: true }),
  column3D: K("bar", -4100, "3D-Säulen", { dir: "col", grouping: "standard", d3: true }),
  bar: K("bar", 57, "Gruppierte Balken", { dir: "bar", grouping: "clustered" }),
  barStacked: K("bar", 58, "Gestapelte Balken", { dir: "bar", grouping: "stacked" }),
  barStacked100: K("bar", 59, "Gestapelte Balken (100 %)", { dir: "bar", grouping: "percentStacked" }),
  bar3DClustered: K("bar", 60, "Gruppierte 3D-Balken", { dir: "bar", grouping: "clustered", d3: true }),
  bar3DStacked: K("bar", 61, "Gestapelte 3D-Balken", { dir: "bar", grouping: "stacked", d3: true }),
  bar3DStacked100: K("bar", 62, "Gestapelte 3D-Balken (100 %)", { dir: "bar", grouping: "percentStacked", d3: true }),
  line: K("line", 4, "Linie", { grouping: "standard" }),
  lineStacked: K("line", 63, "Gestapelte Linie", { grouping: "stacked" }),
  lineStacked100: K("line", 64, "Gestapelte Linie (100 %)", { grouping: "percentStacked" }),
  lineMarkers: K("line", 65, "Linie mit Datenpunkten", { grouping: "standard", markers: true }),
  lineMarkersStacked: K("line", 66, "Gestapelte Linie mit Datenpunkten", { grouping: "stacked", markers: true }),
  lineMarkersStacked100: K("line", 67, "Gestapelte Linie mit Datenpunkten (100 %)", { grouping: "percentStacked", markers: true }),
  line3D: K("line", -4101, "3D-Linie", { grouping: "standard", d3: true }),
  area: K("area", 1, "Fläche", { grouping: "standard" }),
  areaStacked: K("area", 76, "Gestapelte Fläche", { grouping: "stacked" }),
  areaStacked100: K("area", 77, "Gestapelte Fläche (100 %)", { grouping: "percentStacked" }),
  area3D: K("area", -4098, "3D-Fläche", { grouping: "standard", d3: true }),
  area3DStacked: K("area", 78, "Gestapelte 3D-Fläche", { grouping: "stacked", d3: true }),
  area3DStacked100: K("area", 79, "Gestapelte 3D-Fläche (100 %)", { grouping: "percentStacked", d3: true }),
  pie: K("pie", 5, "Kreis"),
  pie3D: K("pie", -4102, "3D-Kreis", { d3: true }),
  pieOfPie: K("ofPie", 68, "Kreis aus Kreis"),
  barOfPie: K("ofPie", 71, "Balken aus Kreis"),
  doughnut: K("doughnut", -4120, "Ring"),
  scatter: K("scatter", -4169, "Punkt (XY)", { markers: true, lines: "none" }),
  scatterSmoothMarkers: K("scatter", 72, "Punkte mit interpolierten Linien und Datendarstellungen", { markers: true, lines: "smooth" }),
  scatterSmooth: K("scatter", 73, "Punkte mit interpolierten Linien", { markers: false, lines: "smooth" }),
  scatterLines: K("scatter", 74, "Punkte mit geraden Linien und Datendarstellungen", { markers: true, lines: "straight" }),
  scatterLinesNoMarkers: K("scatter", 75, "Punkte mit geraden Linien", { markers: false, lines: "straight" }),
  bubble: K("bubble", 15, "Blase"),
  bubble3D: K("bubble", 87, "3D-Blase", { d3: true }),
  radar: K("radar", -4151, "Netz"),
  radarMarkers: K("radar", 81, "Netz mit Datenpunkten", { markers: true }),
  radarFilled: K("radar", 82, "Gefülltes Netz"),
  stockHLC: K("stock", 88, "Höchst-Tiefst-Schlusskurs"),
  stockOHLC: K("stock", 89, "Eröffnungs-Höchst-Tiefst-Schlusskurs"),
  stockVHLC: K("stock", 90, "Volumen-Höchst-Tiefst-Schlusskurs"),
  stockVOHLC: K("stock", 91, "Volumen-Eröffnungs-Höchst-Tiefst-Schlusskurs"),
  surface3D: K("surface", 83, "3D-Oberfläche", { d3: true }),
  surfaceWireframe: K("surface", 84, "3D-Oberfläche (Drahtmodell)", { d3: true }),
  surfaceContour: K("surface", 85, "Kontur"),
  surfaceContourWireframe: K("surface", 86, "Kontur (Drahtmodell)"),
  histogram: K("histogram", 118, "Histogramm"),
  pareto: K("pareto", 122, "Pareto"),
  boxWhisker: K("boxWhisker", 121, "Kastengrafik"),
  waterfall: K("waterfall", 119, "Wasserfall"),
  funnel: K("funnel", 123, "Trichter"),
  treemap: K("treemap", 117, "Treemap"),
  sunburst: K("sunburst", 120, "Sunburst"),
};

export const kindInfo = (k: ChartKind): KindInfo => KINDS[k] ?? KINDS.column;

/** Excel 2016's chartex types (xl/charts/chartExN.xml). */
export const isChartExKind = (k: ChartKind): boolean =>
  ["histogram", "pareto", "boxWhisker", "waterfall", "funnel", "treemap", "sunburst"].includes(kindInfo(k).family);

export const isWireframe = (k: ChartKind): boolean => k === "surfaceWireframe" || k === "surfaceContourWireframe";

/** The kind of family and grouping with the given 3-D-ness and direction (for building kinds from files). */
export function kindOf(family: Family, o: { dir?: "col" | "bar"; grouping?: Grouping; d3?: boolean; markers?: boolean } = {}): ChartKind | undefined {
  const g = (x?: Grouping) => (x === "clustered" ? "standard" : x ?? "standard");
  return (Object.keys(KINDS) as ChartKind[]).find((k) => {
    const i = KINDS[k];
    if (i.family !== family || !!i.d3 !== !!o.d3) return false;
    if (family === "bar") return i.dir === (o.dir ?? "col") && (i.grouping === (o.grouping ?? "clustered") || (o.d3 && o.grouping === "standard" && i.grouping === "standard"));
    if (family === "line" || family === "area") return g(i.grouping) === g(o.grouping) && (family === "area" || !!i.markers === !!o.markers || !!o.d3);
    return true;
  });
}

/** The 3-D view Excel gives a new chart of the kind (recorded, results/2026-11-08-excel-charttypes.json). */
export function defaultView3D(k: ChartKind): View3D | undefined {
  const i = kindInfo(k);
  if (!i.d3 && i.family !== "surface") return undefined;
  if (i.family === "pie") return { rotX: 30, rotY: 0, depthPercent: 100, rAngAx: false };
  if (i.family === "surface" && !i.d3) return { rotX: 90, rotY: 0, rAngAx: false, perspective: 0 };
  if (i.family === "bar") return { rotX: 15, rotY: 20, depthPercent: 100, rAngAx: true };
  if (i.family === "line" || i.family === "surface") return { rotX: 15, rotY: 20, rAngAx: false };
  return { rotX: 15, rotY: 20, depthPercent: 100, rAngAx: false };
}

/** The chart's view: its own, else the type's default. */
export const view3DOf = (m: ChartModel): View3D | undefined => m.view3D ?? defaultView3D(m.kind);

/** A Verbund chart (series of different types or on the secondary axis). */
export const isCombo = (m: ChartModel): boolean => m.series.some((s) => (s.kind && s.kind !== m.kind) || s.secondary);

/** The type a series is drawn as. */
export const seriesKind = (m: ChartModel, i: number): ChartKind => m.series[i]?.kind ?? m.kind;

export interface ChartGroup { kind: ChartKind; idx: number[]; secondary: boolean }

/** The chart groups: series of the same type and axis together (Verbund); Kurs with Volumen: the columns, then the prices on the secondary axis. */
export function chartGroups(m: ChartModel): ChartGroup[] {
  if (m.kind === "stockVHLC" || m.kind === "stockVOHLC") {
    return [{ kind: "column", idx: [0], secondary: false }, { kind: m.kind, idx: m.series.map((_, i) => i).slice(1), secondary: true }];
  }
  const out: ChartGroup[] = [];
  m.series.forEach((s, i) => {
    const kind = s.kind ?? m.kind, secondary = !!s.secondary;
    const g = out.find((x) => x.kind === kind && x.secondary === secondary);
    if (g) g.idx.push(i); else out.push({ kind, idx: [i], secondary });
  });
  if (!out.length) out.push({ kind: m.kind, idx: [], secondary: false });
  return out;
}

/** Defaults of Abstandsbreite and Reihenachsenüberlappung per bar type (Excel). */
export function barDefaults(kind: ChartKind): { gap: number; overlap?: number } {
  const i = kindInfo(kind);
  if (i.d3) return { gap: 150 };
  if (i.grouping === "stacked" || i.grouping === "percentStacked") return { gap: 150, overlap: 100 };
  return i.dir === "bar" ? { gap: 182, overlap: 0 } : { gap: 219, overlap: -27 };
}
