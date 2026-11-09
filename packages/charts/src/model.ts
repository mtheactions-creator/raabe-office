// The chart model shared by RaabeCalc, RaabePresent and RaabeWrite (Step 7 E, D-061). It follows
// DrawingML's chart part (c:chartSpace) closely so that reading and writing stay simple: series
// refer to cells in the file's form ("Tabelle1!$B$2:$B$6") and carry the values last read
// (Excel's caches), so a chart draws without its workbook (Present, Write).

import { isChartExKind, kindInfo } from "./kinds";

export type ChartKind =
  | "column" | "columnStacked" | "bar" | "barStacked" | "line" | "lineMarkers" | "area"
  | "pie" | "doughnut" | "scatter" | "scatterSmooth" | "scatterLines" | "histogram"
  // Every other type of Excel 365's Einfügen > Diagramme except the map (D-086); see kinds.ts.
  | "columnStacked100" | "barStacked100" | "lineStacked" | "lineStacked100" | "lineMarkersStacked" | "lineMarkersStacked100"
  | "areaStacked" | "areaStacked100"
  | "column3DClustered" | "column3DStacked" | "column3DStacked100" | "column3D" | "bar3DClustered" | "bar3DStacked" | "bar3DStacked100"
  | "line3D" | "area3D" | "area3DStacked" | "area3DStacked100" | "pie3D" | "pieOfPie" | "barOfPie"
  | "radar" | "radarMarkers" | "radarFilled" | "bubble" | "bubble3D" | "scatterSmoothMarkers" | "scatterLinesNoMarkers"
  | "stockHLC" | "stockOHLC" | "stockVHLC" | "stockVOHLC"
  | "surface3D" | "surfaceWireframe" | "surfaceContour" | "surfaceContourWireframe"
  | "waterfall" | "funnel" | "treemap" | "sunburst" | "boxWhisker" | "pareto";

/** Legend position as in the file (c:legendPos), or no legend. */
export type LegendPos = "b" | "t" | "l" | "r" | "tr" | "none";

export interface ChartData<T> {
  /** Cell reference in the file's form ("Tabelle1!$A$2:$A$6"); undefined for literal data. */
  ref?: string;
  /** The values last read from the cells (or literal values). */
  cache: T[];
  /** Number format of the first cell (file form), for axis and data labels. */
  format?: string;
}

// ----- formatting (Excel's Format panes, D-066) ------------------------------------------------------
// Unset fields keep Excel's automatic look (style 201); what is set is written to the file as
// DrawingML (c:spPr, c:txPr, c:scaling …) and read back.

/** A Farbverlaufsstopp: position 0-100 %, colour, Transparenz %. */
export interface GradientStop { pos: number; color: string; transparency?: number }

/** Farbverlauf: Typ (Linear, Radial, Rechteckig, Pfad), Richtung as an angle (linear, degrees), the stops. */
export interface Gradient {
  kind: "linear" | "radial" | "rect" | "path";
  /** Winkel of a linear gradient in degrees (0: left to right, 90: top to bottom, DrawingML's a:lin ang). */
  angle?: number;
  stops: GradientStop[];
  /** Mit Form drehen. */
  rotWithShape?: boolean;
}

/** Musterfüllung: one of DrawingML's 48 preset patterns with Vordergrund and Hintergrund. */
export interface Pattern { preset: string; fg: string; bg: string }

/** Bild- oder Texturfüllung: a picture (data URL), stretched or tiled (Bild nebeneinander als Textur anordnen). */
export interface PictureFill { src: string; tile?: boolean; transparency?: number; /** Offsets in % of the element (Strecken: links, rechts, oben, unten). */ offsets?: { l: number; t: number; r: number; b: number } }

/** Füllung: Keine Füllung, Einfarbige Füllung (colour, Transparenz in %), Farbverlauf, Bild- oder Texturfüllung, Musterfüllung, Automatisch. */
export interface Fill {
  type: "none" | "solid" | "auto" | "gradient" | "picture" | "pattern";
  color?: string;
  transparency?: number;
  gradient?: Gradient;
  pattern?: Pattern;
  picture?: PictureFill;
}

/** Strichtyp, as DrawingML's preset dashes. */
export type Dash = "solid" | "dot" | "sysDot" | "sysDash" | "sysDashDot" | "sysDashDotDot" | "dash" | "dashDot" | "lgDash" | "lgDashDot" | "lgDashDotDot";

/** Pfeiltyp of a line end (DrawingML's a:headEnd / a:tailEnd), with Pfeilgröße. */
export interface Arrow { type: "triangle" | "stealth" | "diamond" | "oval" | "arrow"; w?: "sm" | "med" | "lg"; len?: "sm" | "med" | "lg" }

/** Rahmen or Linie: Keine Linie, Einfarbige Linie, Farbverlaufslinie, Automatisch; Breite in pt. */
export interface Line {
  type: "none" | "solid" | "auto" | "gradient";
  color?: string;
  transparency?: number;
  width?: number;
  dash?: Dash;
  gradient?: Gradient;
  /** Verbundtyp: Einfach, Doppelt, Dick-dünn, Dünn-dick, Dreifach. */
  compound?: "sng" | "dbl" | "thickThin" | "thinThick" | "tri";
  /** Abschlusstyp: Quadrat, Rund, Flach. */
  cap?: "sq" | "rnd" | "flat";
  /** Anschlusstyp: Rund, Abgeschrägt, Gehrung. */
  join?: "round" | "bevel" | "miter";
  /** Anfangs- / Endpfeiltyp (lines only). */
  head?: Arrow;
  tail?: Arrow;
}

/** Schatten (a:outerShdw / a:innerShdw): Farbe, Transparenz %, Größe %, Weichzeichnen pt, Winkel °, Abstand pt. */
export interface Shadow { inner?: boolean; color?: string; transparency?: number; size?: number; blur?: number; angle?: number; distance?: number }
/** Leuchten: Farbe, Größe pt, Transparenz %. */
export interface Glow { color: string; size: number; transparency?: number }
/** 3D-Format: Abschrägung oben/unten (preset, Breite/Höhe pt), Tiefe, Kontur, Material, Beleuchtung. */
export interface Format3D {
  top?: { type: string; w?: number; h?: number };
  bottom?: { type: string; w?: number; h?: number };
  depth?: { color?: string; size: number };
  contour?: { color?: string; size: number };
  material?: string;
  lighting?: string;
}
/** Spiegelung of text: Transparenz %, Größe % of the text's height, Weichzeichnen pt, Abstand pt. */
export interface Reflection { transparency: number; size: number; blur: number; distance: number }
/** Effekte: Schatten, Leuchten, Weiche Kanten (pt), 3D-Format; Spiegelung (text). */
export interface Effects { shadow?: Shadow; glow?: Glow; softEdges?: number; format3d?: Format3D; reflection?: Reflection }

/** Textfeld of an element: Vertikale Ausrichtung, Textrichtung, Benutzerdefinierter Winkel, Ränder (pt), Textumbruch. */
export interface TextBox {
  vAlign?: "t" | "ctr" | "b";
  dir?: "horz" | "vert" | "vert270" | "wordArtVert";
  /** Benutzerdefinierter Winkel in degrees (-90 … 90). */
  rotation?: number;
  margins?: { l: number; t: number; r: number; b: number };
  wrap?: boolean;
  autofit?: boolean;
  /** The text block centred horizontally in its box (the "zentriert" Vertikale Ausrichtung, the default). */
  anchorCtr?: boolean;
}

/** Text of an element (Start > Schriftart while a chart element is selected; Textoptionen). */
export interface TextFormat {
  size?: number;
  bold?: boolean;
  italic?: boolean;
  /** Einfarbige Füllung's colour (also the colour of Start > Schriftfarbe). */
  color?: string;
  underline?: boolean;
  strike?: boolean;
  font?: string;
  /** Textfüllung: Keine Füllung, Farbverlauf … (solid: `color`). */
  fill?: Fill;
  /** Textkontur. */
  outline?: Line;
  /** Texteffekte. */
  effects?: Effects;
  box?: TextBox;
}

export interface ElementFormat { fill?: Fill; line?: Line; text?: TextFormat; effects?: Effects }

/** Teilstriche: Ohne, Innen, Außen, Kreuz. */
export type TickMark = "none" | "in" | "out" | "cross";

/** Achsenoptionen (value axes: bounds and units; both: ticks, labels, number format). */
export interface AxisOptions extends ElementFormat {
  /** Begrenzungen Minimum / Maximum; undefined: Automatisch. */
  min?: number;
  max?: number;
  /** Einheiten Haupt / Hilfs; undefined: Automatisch. */
  majorUnit?: number;
  minorUnit?: number;
  /** Logarithmische Skalierung with this Basis. */
  logBase?: number;
  /** Werte in umgekehrter Reihenfolge / Kategorien in umgekehrter Reihenfolge. */
  reverse?: boolean;
  /**
   * Where the other axis crosses this one, in this axis' units: the value axis' "Horizontale Achse
   * schneidet" (Automatisch, Achsenwert, Maximaler Achsenwert), the category axis' "Vertikale Achse
   * schneidet" (Automatisch, Bei Kategorienummer, Bei größter Kategorie).
   */
  crosses?: "auto" | "max" | number;
  majorTick?: TickMark;
  minorTick?: TickMark;
  /** Beschriftungsposition: Achsennah, Hoch, Tief, Keine. */
  labelPos?: "nextTo" | "high" | "low" | "none";
  /** Zahlenformat of the labels; undefined: Mit Quelle verknüpfen. */
  numFmt?: string;
  /** The axis is shown (Diagrammelemente > Achsen). */
  visible?: boolean;
  /** Achsentyp of a category axis: Automatisch aus Daten auswählen, Textachse, Datumsachse. */
  axisType?: "auto" | "text" | "date";
  /** Datumsachse: Basis and Haupt-/Hilfseinheit's time unit. */
  baseTimeUnit?: "days" | "months" | "years";
  majorTimeUnit?: "days" | "months" | "years";
  minorTimeUnit?: "days" | "months" | "years";
  /** Anzeigeeinheiten (value axes) and Anzeigeeinheitenbeschriftung im Diagramm anzeigen. */
  dispUnits?: "hundreds" | "thousands" | "tenThousands" | "hundredThousands" | "millions" | "tenMillions" | "hundredMillions" | "billions" | "trillions" | number;
  dispUnitsLabel?: boolean;
  /** Intervall zwischen Teilstrichen / Intervall zwischen Beschriftungen (category axes); undefined: Automatisch. */
  tickMarkSkip?: number;
  tickLblSkip?: number;
  /** Abstand von Achse (category axes, 0-1000, 100 the default). */
  lblOffset?: number;
  /** Achsenposition: Auf Teilstrichen (true) / Zwischen Teilstrichen (category axes). */
  onTicks?: boolean;
}

/** Datenbeschriftungen: what they show and where (Beschriftungsoptionen). */
export interface DataLabelOptions extends ElementFormat {
  showVal?: boolean;
  showCat?: boolean;
  showSer?: boolean;
  showPercent?: boolean;
  showLegendKey?: boolean;
  /** Beschriftungsposition: Zentriert, Innerhalb Ende, Innerhalb Basis, Außerhalb Ende, Links, Rechts, Über, Unter, Optimale Anpassung. */
  pos?: "ctr" | "inEnd" | "inBase" | "outEnd" | "l" | "r" | "t" | "b" | "bestFit";
  /** Trennzeichen between the parts: "; " ", " ". " line break. */
  separator?: string;
  numFmt?: string;
  /** Blasengröße (bubble charts). */
  showBubble?: boolean;
  /** Führungslinien anzeigen (pies; Excel's default on). */
  leaderLines?: boolean;
  /** Wert aus Zellen: the range and its values last read (Excel 2013's c15:datalabelsRange). */
  fromCells?: { ref: string; cache: (string | null)[] };
}

export type MarkerSymbol = "auto" | "none" | "circle" | "square" | "diamond" | "triangle" | "x" | "star" | "dash" | "dot" | "plus";

/** Markierung of line and scatter series. */
export interface MarkerOptions { symbol?: MarkerSymbol; size?: number; fill?: Fill; line?: Line }

export interface Trendline {
  /** Linear, Exponentiell, Logarithmisch, Polynomisch (order 2-6), Potenz, Gleitender Durchschnitt (period). */
  type: "linear" | "exp" | "log" | "poly" | "power" | "movingAvg";
  order?: number;
  period?: number;
  /** Name: undefined: Automatisch ("Linear (Reihe1)"). */
  name?: string;
  /** Prognose vorwärts / rückwärts in category units. */
  forward?: number;
  backward?: number;
  /** Schnittpunkt festlegen. */
  intercept?: number;
  /** Formel im Diagramm anzeigen. */
  showEq: boolean;
  /** Bestimmtheitsmaß im Diagramm anzeigen. */
  showR2: boolean;
  line?: Line;
  effects?: Effects;
}

/** Fehlerbalken (Fehlerindikatoren): direction, end style, amount. */
export interface ErrorBars {
  dir: "both" | "minus" | "plus";
  /** Endformat: Ohne Abschluss (true) / Abschluss. */
  noEndCap?: boolean;
  /** Fehlerbetrag: Fester Wert, Prozentsatz, Standardabweichung(en), Standardfehler. */
  type: "fixedVal" | "percentage" | "stdDev" | "stdErr" | "cust";
  value?: number;
  /** Benutzerdefiniert ("cust"): the amounts above and below each point, from cells (D-125). */
  plus?: ChartData<number | null>;
  minus?: ChartData<number | null>;
  line?: Line;
  effects?: Effects;
}

/** The formatting of the chart's elements that are not series. */
export interface ChartFormat {
  /** Diagrammbereich; roundedCorners: Abgerundete Ecken. */
  chart?: ElementFormat & { roundedCorners?: boolean };
  /** Sekundäre vertikale (Wert) Achse. */
  valAxis2?: AxisOptions;
  plot?: ElementFormat;
  /** Diagrammtitel; overlay: Titel über dem Diagramm einblenden, ohne die Größe zu ändern. */
  title?: ElementFormat & { overlay?: boolean };
  legend?: ElementFormat & { overlay?: boolean };
  catAxis?: AxisOptions;
  valAxis?: AxisOptions;
  /** Primäre Hauptgitternetzlinien of the value axis (their Linie). */
  majorGrid?: ElementFormat;
  /** Primäre Hilfsgitternetzlinien of the value axis, shown when set. */
  minorGrid?: ElementFormat & { show?: boolean };
  /** Primäre vertikale Hauptgitternetzlinien (category axis). */
  catGrid?: ElementFormat & { show?: boolean };
  /** Primäre vertikale Hilfsgitternetzlinien (category axis; Diagrammelemente > Gitternetzlinien, D-115). */
  catMinorGrid?: ElementFormat & { show?: boolean };
  axisTitles?: { cat?: ElementFormat; val?: ElementFormat };
  /** Datenbeschriftungen of all series. */
  dataLabels?: DataLabelOptions;
}

export interface ChartSeries {
  name?: { ref?: string; text?: string };
  /** Categories; X values of scatter charts. */
  cat?: ChartData<string | number | null>;
  val: ChartData<number | null>;
  /** Fill/line colour "#RRGGBB"; undefined: the theme's accent for the series' index. */
  color?: string;
  trendline?: Trendline;
  /** Füllung and Rahmen (columns, areas, pie slices) or Linie (lines, scatter lines) from the Format pane. */
  fill?: Fill;
  line?: Line;
  marker?: MarkerOptions;
  /** Linie glätten. */
  smooth?: boolean;
  /** Invertieren falls negativ. */
  invertIfNegative?: boolean;
  /** Sekundärachse (Datenreihe zeichnen auf). */
  secondary?: boolean;
  errorBars?: ErrorBars;
  /** Verbund (combo charts): the series' own chart type when it differs from the chart's. */
  kind?: ChartKind;
  /** Blasengröße of bubble charts. */
  size?: ChartData<number | null>;
  /** Category levels of Treemap and Sunburst, outermost first (each as long as the values). */
  levels?: (string | null)[][];
  /** Wasserfall: the points shown as Summe (Als Summe festlegen). */
  totals?: number[];
  /** Effekte of the series (Schatten, Leuchten, Weiche Kanten, 3D-Format). */
  effects?: Effects;
  /** Punktexplosion of pie and doughnut series in %. */
  explosion?: number;
  /** Datenpunkt formatieren: points with their own Füllung, Rahmen, Effekte, Explosion or Markierung. */
  points?: DataPoint[];
  /** Hidden by Diagrammfilter (D-116): kept in the chart and the file (c15:filtered…Series) but not drawn. */
  filtered?: boolean;
}

/** A data point with its own format (c:dPt). */
export interface DataPoint { idx: number; fill?: Fill; line?: Line; effects?: Effects; explosion?: number; marker?: MarkerOptions; invertIfNegative?: boolean }

/** Eigenschaften of the chart object (Größe und Eigenschaften). */
export interface ChartObjectProps {
  /** Von Zellposition und -größe abhängig (twoCell, the default), Von Zellposition abhängig, nicht von Größe (oneCell), Von Zellposition und -größe unabhängig (absolute). */
  placement?: "twoCell" | "oneCell" | "absolute";
  /** Objekt drucken (default on). */
  print?: boolean;
  /** Gesperrt (default on). */
  locked?: boolean;
  /** Seitenverhältnis sperren. */
  lockAspect?: boolean;
  /** Alternativtext: Titel and Beschreibung. */
  altTitle?: string;
  altText?: string;
}

/** The 3-D view (Diagrammbereich formatieren > 3D-Drehung) of 3-D charts. */
export interface View3D { rotX: number; rotY: number; depthPercent?: number; rAngAx: boolean; perspective?: number }

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
  /** Formatting from the Format panes; unset: Excel's automatic look. */
  format?: ChartFormat;
  /** Reihenachsenüberlappung and Abstandsbreite (columns and bars), in %. */
  overlap?: number;
  gapWidth?: number;
  /** Winkel des ersten Kreissegments (pies), Innenringgröße (doughnuts, %). */
  firstSliceAngle?: number;
  holeSize?: number;
  /** 3-D view of 3-D charts as read; undefined: the type's default (kinds.ts). */
  view3D?: View3D;
  /** Kreis aus Kreis / Balken aus Kreis: Werte im zweiten Zeichnungsbereich (the last n points) and its size in %. */
  splitCount?: number;
  secondSize?: number;
  /** Blasengröße skalieren (%). */
  bubbleScale?: number;
  /** Diagrammformatvorlage (Excel's style number, 201 the default) and Farben ändern (colour scheme, 10 the default). */
  style?: number;
  colors?: number;
  /** Diagrammfilter > Kategorien (D-116): the categories (0-based, of the full range) not drawn. */
  hiddenCats?: number[];
  /** Ausgeblendete und leere Zellen: leere Zellen als Lücken (default), Null or Datenpunkte mit Linie verbinden. */
  dispBlanksAs?: "gap" | "zero" | "span";
  /** Daten in ausgeblendeten Zeilen und Spalten anzeigen (c:plotVisOnly val="0"). */
  showHidden?: boolean;
  /** #NV als leere Zelle anzeigen (Excel 365's c16r3:dataDisplayOptions16). */
  naAsBlank?: boolean;
  /** Datentabelle below the plot area, with or without the legend keys (D-115). */
  dataTable?: { legendKeys: boolean };
  from: CellAnchor;
  to: CellAnchor;
  /** Name of the drawing object ("Diagramm 1"). */
  name?: string;
  /** Größe und Eigenschaften: Eigenschaften and Alternativtext. */
  props?: ChartObjectProps;
  /** Punkte nach Farbe unterscheiden (single-series charts other than pies). */
  varyColors?: boolean;
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
  // Paint that is not a colour (a gradient or pattern reference) stays as it is.
  if (!/^#[0-9a-f]{6}/i.test(hex)) return hex;
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

/** Kinds whose legend lists the points (Kreis, Ring and their variants). */
export const isPie = (k: ChartKind) => { const f = kindInfo(k).family; return f === "pie" || f === "doughnut" || f === "ofPie"; };
/** XY charts (Punkt and Blase): both axes are value axes. */
export const isScatter = (k: ChartKind) => { const f = kindInfo(k).family; return f === "scatter" || f === "bubble"; };
/** Horizontal bars. */
export const isBar = (k: ChartKind) => kindInfo(k).dir === "bar";
/** Stacked (not 100 %) columns and bars. */
export const isStacked = (k: ChartKind) => kindInfo(k).family === "bar" && kindInfo(k).grouping === "stacked";

/** The title as shown: its own text, else the series name for one series (pies: the first), else "Diagrammtitel". */
export function titleText(m: ChartModel): string {
  if (m.title.text !== undefined) return m.title.text;
  if (isChartExKind(m.kind)) return DEFAULT_TITLE;
  const one = isPie(m.kind) ? m.series[0] : m.series.length === 1 ? m.series[0] : undefined;
  return one?.name?.text || DEFAULT_TITLE;
}

/** The legend Excel gives a new chart: pies always (bottom), others with more than one series (recorded per type). */
export function defaultLegend(kind: ChartKind, seriesCount: number): LegendPos {
  const f = kindInfo(kind).family;
  if (f === "waterfall" || f === "treemap") return "t";
  if (isChartExKind(kind) || f === "bubble") return "none";
  if (f === "radar") return "t";
  if (isPie(kind) || f === "stock" || f === "surface") return "b";
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
