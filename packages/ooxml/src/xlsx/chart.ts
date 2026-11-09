// Charts in .xlsx (Step 7 E, D-061): RaabeCalc's ChartModel to and from DrawingML's chart part
// (c:chartSpace, xl/charts/chartN.xml), Excel 2016's chartex part for histograms
// (cx:chartSpace, xl/charts/chartExN.xml) and the drawing's anchor (xdr:twoCellAnchor in
// xl/drawings/drawingN.xml). The XML is written as Excel 365 writes a new chart (style 201 of the
// Office 2013-2022 theme), so files open in Excel without repair and look the same there.
// A chart read from a file keeps its XML in `raw`; while the modelled parts stay unchanged it is
// written back verbatim (D-024), with Excel recalculating the caches from the cells on opening.

import {
  OFFICE_ACCENTS, barDefaults, chartGroups, defaultView3D, isChartExKind, isCombo, isPie, kindInfo, kindOf, paletteColor, view3DOf,
  type CellAnchor, type ChartGroup, type ChartKind, type ChartModel, type ChartSeries, type LegendPos, type View3D,
} from "@raabe/charts";
import { attrs, child, children, esc, escText, root, text, unesc, type El } from "./xml";
import {
  crossesXml, defRPrAttrs, errBarsXml, labelFlags, lineXml, markerXml, readAxis, readFormat, readShape, readText, scalingXml,
  spPrXml, tickXml, trendlineXml, type Reader,
} from "./chartfmt";
import type { AxisOptions, ChartFormat, DataLabelOptions, ElementFormat, ErrorBars, MarkerSymbol, TextFormat, Trendline } from "@raabe/charts";
import { applyFilterRead, filterXml, unfilterXml, withHiddenCats } from "./chartfilter";

export const CHART_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.drawingml.chart+xml";
export const CHART_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart";
export const CHARTEX_CONTENT_TYPE = "application/vnd.ms-office.chartex+xml";
export const CHARTEX_REL_TYPE = "http://schemas.microsoft.com/office/2014/relationships/chartEx";
export const DRAWING_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.drawing+xml";
export const DRAWING_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing";

const HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n';
const NS_C = "http://schemas.openxmlformats.org/drawingml/2006/chart";
const NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main";
const NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const NS_CX = "http://schemas.microsoft.com/office/drawing/2014/chartex";
/** EMU per CSS px at 96 dpi. */
const EMU_PX = 9525;

/** Histograms, Pareto, Kastengrafik, Wasserfall, Trichter, Treemap and Sunburst are chartex parts (another content and relationship type, see the constants). */
export function isChartEx(m: ChartModel): boolean {
  return isChartExKind(m.kind);
}

// ----- writing ----------------------------------------------------------------------------------

const GREY = '<a:solidFill><a:schemeClr val="tx1"><a:lumMod val="65000"/><a:lumOff val="35000"/></a:schemeClr></a:solidFill>';
const LINE_GREY = '<a:solidFill><a:schemeClr val="tx1"><a:lumMod val="15000"/><a:lumOff val="85000"/></a:schemeClr></a:solidFill>';
const NO_SHAPE = "<c:spPr><a:noFill/><a:ln><a:noFill/></a:ln><a:effectLst/></c:spPr>";

function txPr(size: number, rot = "-60000000", extra = "", t?: TextFormat): string {
  const r = defRPrAttrs(t, size);
  return `<c:txPr><a:bodyPr rot="${rot}" spcFirstLastPara="1" vertOverflow="ellipsis" vert="horz" wrap="square" anchor="ctr" anchorCtr="1"/><a:lstStyle/><a:p><a:pPr><a:defRPr${r.attrs} u="none" strike="noStrike" kern="1200"${extra} baseline="0">${r.fill ?? GREY}<a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:pPr><a:endParaRPr lang="de-DE"/></a:p></c:txPr>`;
}

/** The automatic look of an element without its own shape (no fill, no line). */
const NONE_FILL = "<a:noFill/>", NONE_LINE = "<a:ln><a:noFill/></a:ln>";

function rich(textValue: string, size: number, rot = "0"): string {
  const runs = textValue.split("\n").map((line) => `<a:p><a:pPr><a:defRPr sz="${size}" b="0" i="0" u="none" strike="noStrike" kern="1200" baseline="0">${GREY}<a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:pPr><a:r><a:rPr lang="de-DE"/><a:t>${escText(line)}</a:t></a:r></a:p>`).join("");
  return `<c:tx><c:rich><a:bodyPr rot="${rot}" spcFirstLastPara="1" vertOverflow="ellipsis" vert="horz" wrap="square" anchor="ctr" anchorCtr="1"/><a:lstStyle/>${runs}</c:rich></c:tx>`;
}

/** A colour: the series' own, else the theme accent (with Excel's variants after six series). */
function fill(s: ChartSeries | undefined, i: number): string {
  if (s?.color) return `<a:solidFill><a:srgbClr val="${s.color.replace("#", "").toUpperCase()}"/></a:solidFill>`;
  return `<a:solidFill>${schemeColor(i)}</a:solidFill>`;
}

function schemeColor(i: number): string {
  const accent = `accent${(i % 6) + 1}`;
  const round = Math.floor(i / 6);
  if (!round) return `<a:schemeClr val="${accent}"/>`;
  const steps: [number, number][] = [[60000, 0], [80000, 20000], [80000, 0], [60000, 40000], [50000, 0], [70000, 30000]];
  const [mod, off] = steps[(round - 1) % steps.length];
  return `<a:schemeClr val="${accent}"><a:lumMod val="${mod}"/>${off ? `<a:lumOff val="${off}"/>` : ""}</a:schemeClr>`;
}

/** A scheme colour with an a:alpha (Transparenz). */
function withAlpha(clr: string, alpha: number): string {
  return clr.endsWith("/>") ? clr.slice(0, -2) + `><a:alpha val="${alpha}"/></a:schemeClr>` : clr.replace("</a:schemeClr>", `<a:alpha val="${alpha}"/></a:schemeClr>`);
}

function seriesText(s: ChartSeries): string {
  if (!s.name) return "";
  if (s.name.ref) return `<c:tx><c:strRef><c:f>${esc(s.name.ref)}</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>${escText(s.name.text ?? "")}</c:v></c:pt></c:strCache></c:strRef></c:tx>`;
  return `<c:tx><c:v>${escText(s.name.text ?? "")}</c:v></c:tx>`;
}

function num(v: number): string {
  return String(v);
}

function numData(tag: string, d: { ref?: string; cache: (number | string | null)[]; format?: string }): string {
  const pts = d.cache.map((v, i) => (typeof v === "number" && isFinite(v) ? `<c:pt idx="${i}"><c:v>${num(v)}</c:v></c:pt>` : "")).join("");
  const body = `<c:formatCode>${escText(d.format ?? "General")}</c:formatCode><c:ptCount val="${d.cache.length}"/>${pts}`;
  return d.ref ? `<c:${tag}><c:numRef><c:f>${esc(d.ref)}</c:f><c:numCache>${body}</c:numCache></c:numRef></c:${tag}>` : `<c:${tag}><c:numLit>${body}</c:numLit></c:${tag}>`;
}

function catData(tag: string, d: ChartSeries["cat"]): string {
  if (!d) return "";
  if (d.cache.every((v) => v === null || typeof v === "number") && d.cache.some((v) => v !== null)) return numData(tag, d);
  const pts = d.cache.map((v, i) => (v === null ? "" : `<c:pt idx="${i}"><c:v>${escText(String(v))}</c:v></c:pt>`)).join("");
  const body = `<c:ptCount val="${d.cache.length}"/>${pts}`;
  return d.ref ? `<c:${tag}><c:strRef><c:f>${esc(d.ref)}</c:f><c:strCache>${body}</c:strCache></c:strRef></c:${tag}>` : `<c:${tag}><c:strLit>${body}</c:strLit></c:${tag}>`;
}

function flags(val: boolean, leader = false): string {
  return `<c:showLegendKey val="0"/><c:showVal val="${val ? 1 : 0}"/><c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="0"/><c:showBubbleSize val="0"/>${leader ? '<c:showLeaderLines val="1"/>' : ""}`;
}

/** Datenbeschriftungen of a series with Excel's default position for the chart type, or the pane's options. */
function seriesLabels(kind: ChartKind, o?: DataLabelOptions): string {
  const numFmt = o?.numFmt ? `<c:numFmt formatCode="${esc(o.numFmt)}" sourceLinked="0"/>` : "";
  return `<c:dLbls>${numFmt}${spPrXml(o, NONE_FILL, NONE_LINE)}${txPr(900, "0", "", o?.text)}${labelFlags(o, defLabelPos(kind), isPie(kind))}</c:dLbls>`;
}

/** Excel's default Beschriftungsposition per chart type (3-D columns, areas, radar and surfaces have none). */
export function defLabelPos(kind: ChartKind): string | undefined {
  const i = kindInfo(kind);
  if (i.d3 && i.family !== "pie") return undefined;
  switch (i.family) {
    case "bar": return i.grouping === "clustered" ? "outEnd" : "ctr";
    case "line": case "scatter": case "bubble": case "stock": return "r";
    case "pie": case "ofPie": return "bestFit";
    default: return undefined;
  }
}

function trendline(s: ChartSeries, i: number): string {
  const t = s.trendline;
  if (!t) return "";
  const lbl = t.showEq || t.showR2 ? `<c:trendlineLbl><c:numFmt formatCode="General" sourceLinked="0"/>${NO_SHAPE}${txPr(900, "0")}</c:trendlineLbl>` : "";
  return trendlineXml(t, `<a:ln w="19050" cap="rnd">${fill(s, i)}<a:prstDash val="sysDot"/></a:ln>`, lbl);
}

function errBars(s: ChartSeries, scatter: boolean): string {
  return s.errorBars ? errBarsXml(s.errorBars, scatter, `<a:ln w="9525" cap="flat" cmpd="sng" algn="ctr">${LINE_GREY}<a:round/></a:ln>`, numData) : "";
}

const MARKER_NONE = '<c:marker><c:symbol val="none"/></c:marker>';
function marker(s: ChartSeries, i: number): string {
  return `<c:marker><c:symbol val="circle"/><c:size val="5"/><c:spPr>${fill(s, i)}<a:ln w="9525">${fill(s, i)}</a:ln><a:effectLst/></c:spPr></c:marker>`;
}

function ser(m: ChartModel, s: ChartSeries, i: number): string {
  const head = `<c:ser><c:idx val="${i}"/><c:order val="${i}"/>${seriesText(s)}`;
  const info = kindInfo(m.kind);
  const labels = m.dataLabels && info.family !== "surface" ? seriesLabels(m.kind, m.format?.dataLabels) : "";
  const mk = (auto: string) => markerXml(s.marker, auto, fill(s, i), `<a:ln w="9525">${fill(s, i)}</a:ln>`);
  const family = info.family === "bar" ? "column" : info.family === "line" ? (info.d3 ? "line3D" : "line") : info.family === "area" ? "area"
    : info.family === "pie" || info.family === "doughnut" || info.family === "ofPie" ? "pie" : info.family;
  switch (family) {
    case "column":
      return `${head}${spPrXml(s, fill(s, i), NONE_LINE)}<c:invertIfNegative val="${s.invertIfNegative ? 1 : 0}"/>${labels}${trendline(s, i)}${errBars(s, false)}${catData("cat", s.cat)}${numData("val", s.val)}</c:ser>`;
    case "line":
      return `${head}<c:spPr>${lineXml(s.line, `<a:ln w="28575" cap="rnd">${fill(s, i)}<a:round/></a:ln>`)}<a:effectLst/></c:spPr>${mk(info.markers ? marker(s, i) : MARKER_NONE)}${labels}${trendline(s, i)}${errBars(s, false)}${catData("cat", s.cat)}${numData("val", s.val)}<c:smooth val="${s.smooth ? 1 : 0}"/></c:ser>`;
    case "area":
      return `${head}${spPrXml(s, fill(s, i), NONE_LINE)}${labels}${trendline(s, i)}${errBars(s, false)}${catData("cat", s.cat)}${numData("val", s.val)}</c:ser>`;
    case "line3D":
      return `${head}${spPrXml(s, fill(s, i), NONE_LINE)}${labels}${catData("cat", s.cat)}${numData("val", s.val)}<c:smooth val="0"/></c:ser>`;
    case "pie": {
      // A series with its own Füllung has one colour; otherwise every slice takes the next accent.
      const own = s.fill && s.fill.type !== "auto";
      const ln = lineXml(s.line, '<a:ln w="19050"><a:solidFill><a:schemeClr val="lt1"/></a:solidFill></a:ln>');
      const slice = (k: number) => (m.colors && m.colors !== 10 ? `<a:srgbClr val="${paletteColor(m, k, s.val.cache.length, OFFICE_ACCENTS).slice(1)}"/>` : schemeColor(k));
      // Kreis aus Kreis / Balken aus Kreis: one more point, the "Sonstige" slice of the first plot.
      const n = s.val.cache.length + (info.family === "ofPie" ? 1 : 0);
      const pts = own ? "" : Array.from({ length: n }, (_, k) => k).map((k) => `<c:dPt><c:idx val="${k}"/><c:bubble3D val="0"/><c:spPr><a:solidFill>${slice(k)}</a:solidFill>${ln}<a:effectLst/></c:spPr></c:dPt>`).join("");
      return `${head}${own ? spPrXml(s, "", ln) : ""}${pts}${labels}${catData("cat", s.cat)}${numData("val", s.val)}</c:ser>`;
    }
    case "bubble": {
      // Blase: the accent at 75 % opacity (Excel's style 201), 3-D bubbles shaded.
      const auto = s.color ? `<a:solidFill><a:srgbClr val="${s.color.replace("#", "").toUpperCase()}"><a:alpha val="75000"/></a:srgbClr></a:solidFill>`
        : `<a:solidFill>${withAlpha(schemeColor(i), 75000)}</a:solidFill>`;
      return `${head}${spPrXml(s, auto, NONE_LINE)}<c:invertIfNegative val="0"/>${labels}${catData("xVal", s.cat)}${numData("yVal", s.val)}${s.size ? numData("bubbleSize", s.size) : ""}<c:bubble3D val="${info.d3 ? 1 : 0}"/></c:ser>`;
    }
    case "radar":
      if (m.kind === "radarFilled") return `${head}${spPrXml(s, fill(s, i), NONE_LINE)}${labels}${catData("cat", s.cat)}${numData("val", s.val)}</c:ser>`;
      return `${head}<c:spPr>${lineXml(s.line, `<a:ln w="28575" cap="rnd">${fill(s, i)}<a:round/></a:ln>`)}<a:effectLst/></c:spPr>${mk(info.markers ? marker(s, i) : MARKER_NONE)}${labels}${catData("cat", s.cat)}${numData("val", s.val)}</c:ser>`;
    case "stock": {
      // Kurs: the prices have no line (hi-lo lines and up/down bars show them); the close a dot.
      const close = i === m.series.length - 1 && (m.kind === "stockHLC" || m.kind === "stockVHLC");
      const sym = close ? `<c:marker><c:symbol val="dot"/><c:size val="3"/>${spPrXml(undefined, fill(s, i), `<a:ln w="9525">${fill(s, i)}</a:ln>`)}</c:marker>` : MARKER_NONE;
      return `${head}<c:spPr>${lineXml(s.line, '<a:ln w="19050" cap="rnd"><a:noFill/><a:round/></a:ln>')}<a:effectLst/></c:spPr>${mk(sym)}${labels}${catData("cat", s.cat)}${numData("val", s.val)}<c:smooth val="0"/></c:ser>`;
    }
    case "surface":
      return `${head}${spPrXml(s, fill(s, i), "<a:ln/>")}${catData("cat", s.cat)}${numData("val", s.val)}</c:ser>`;
    default: {
      // Scatter charts: Punkte (no line), with smoothed or straight lines, with or without markers.
      const auto = info.lines === "none" ? '<a:ln w="19050" cap="rnd"><a:noFill/><a:round/></a:ln>' : `<a:ln w="19050" cap="rnd">${fill(s, i)}<a:round/></a:ln>`;
      return `${head}<c:spPr>${lineXml(s.line, auto)}<a:effectLst/></c:spPr>${mk(info.markers ? marker(s, i) : MARKER_NONE)}${labels}${trendline(s, i)}${errBars(s, true)}${catData("xVal", s.cat)}${numData("yVal", s.val)}<c:smooth val="${info.lines === "smooth" || s.smooth ? 1 : 0}"/></c:ser>`;
    }
  }
}

function axisTitle(textValue: string | undefined, vertical: boolean, f?: ElementFormat): string {
  if (textValue === undefined) return "";
  return `<c:title>${rich(textValue || "Achsentitel", 1000, vertical ? "-5400000" : "0")}<c:overlay val="0"/>${spPrXml(f, NONE_FILL, NONE_LINE)}${txPr(1000, vertical ? "-5400000" : "0", "", f?.text)}</c:title>`;
}

function gridXml(tag: "majorGridlines" | "minorGridlines", f: ElementFormat | undefined, autoLine: string): string {
  return `<c:${tag}><c:spPr>${lineXml(f?.line, autoLine)}<a:effectLst/></c:spPr></c:${tag}>`;
}

const GRID_LINE = `<a:ln w="9525" cap="flat" cmpd="sng" algn="ctr">${LINE_GREY}<a:round/></a:ln>`;
const MINOR_GRID_LINE = '<a:ln w="9525" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="tx1"><a:lumMod val="5000"/><a:lumOff val="95000"/></a:schemeClr></a:solidFill><a:round/></a:ln>';

const AX_CAT = 500000001, AX_VAL = 500000002, AX_CAT2 = 500000003, AX_VAL2 = 500000004, AX_SER = 500000005;

function catAx(pos: string, titleText: string | undefined, f?: ChartFormat, id = AX_CAT, cross = AX_VAL, deleted = false): string {
  const o: AxisOptions | undefined = deleted ? { visible: false } : f?.catAxis;
  const grid = f?.catGrid?.show ? gridXml("majorGridlines", f.catGrid, GRID_LINE) : "";
  const numFmt = o?.numFmt ? `<c:numFmt formatCode="${esc(o.numFmt)}" sourceLinked="0"/>` : '<c:numFmt formatCode="General" sourceLinked="1"/>';
  return `<c:catAx><c:axId val="${id}"/>${scalingXml(o)}<c:delete val="${o?.visible === false ? 1 : 0}"/><c:axPos val="${pos}"/>${grid}${axisTitle(titleText, pos === "l", f?.axisTitles?.cat)}${numFmt}${tickXml("majorTickMark", o?.majorTick)}${tickXml("minorTickMark", o?.minorTick)}<c:tickLblPos val="${o?.labelPos ?? "nextTo"}"/><c:spPr>${NONE_FILL}${lineXml(o?.line, GRID_LINE)}<a:effectLst/></c:spPr>${txPr(900, "-60000000", "", o?.text)}<c:crossAx val="${cross}"/>${crossesXml(deleted ? undefined : f?.valAxis)}<c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/><c:noMultiLvlLbl val="0"/></c:catAx>`;
}

/**
 * A value axis; `fmt` null leaves out the number format (Excel: X values that are text). `o` are the
 * axis' own options, `crossing` those of the other axis (where this one crosses it), `f` the chart's format.
 */
function valAx(fmt: string | null, id: number, cross: number, pos: string, titleText: string | undefined, between: "between" | "midCat", line: boolean, grid: boolean,
  o?: AxisOptions, crossing?: AxisOptions, f?: ChartFormat, titleFmt?: ElementFormat, grids = true): string {
  const majorGrid = grid ? gridXml("majorGridlines", grids ? f?.majorGrid : undefined, GRID_LINE) : "";
  const minorGrid = grids && f?.minorGrid?.show ? gridXml("minorGridlines", f.minorGrid, MINOR_GRID_LINE) : "";
  const autoLine = line ? `<a:ln w="9525" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="tx1"><a:lumMod val="25000"/><a:lumOff val="75000"/></a:schemeClr></a:solidFill><a:round/></a:ln>` : NONE_LINE;
  const numFmt = o?.numFmt ? `<c:numFmt formatCode="${esc(o.numFmt)}" sourceLinked="0"/>` : fmt === null ? "" : `<c:numFmt formatCode="${esc(fmt)}" sourceLinked="1"/>`;
  const units = (o?.majorUnit !== undefined ? `<c:majorUnit val="${o.majorUnit}"/>` : "") + (o?.minorUnit !== undefined ? `<c:minorUnit val="${o.minorUnit}"/>` : "");
  return `<c:valAx><c:axId val="${id}"/>${scalingXml(o)}<c:delete val="${o?.visible === false ? 1 : 0}"/><c:axPos val="${pos}"/>${majorGrid}${minorGrid}${axisTitle(titleText, pos === "l", titleFmt)}${numFmt}${tickXml("majorTickMark", o?.majorTick)}${tickXml("minorTickMark", o?.minorTick)}<c:tickLblPos val="${o?.labelPos ?? "nextTo"}"/><c:spPr>${NONE_FILL}${lineXml(o?.line, autoLine)}<a:effectLst/></c:spPr>${txPr(900, "-60000000", "", o?.text)}<c:crossAx val="${cross}"/>${crossesXml(crossing)}<c:crossBetween val="${between}"/>${units}</c:valAx>`;
}

/** Series axis (Tiefenachse) of 3-D charts with series in depth and of surfaces. */
function serAx(): string {
  return `<c:serAx><c:axId val="${AX_SER}"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:spPr>${NONE_FILL}${GRID_LINE}<a:effectLst/></c:spPr>${txPr(900)}<c:crossAx val="${AX_VAL}"/><c:crosses val="autoZero"/></c:serAx>`;
}

/** The kinds written by plotArea's switch as RaabeCalc always wrote them. */
const LEGACY = new Set<ChartKind>(["column", "columnStacked", "bar", "barStacked", "line", "lineMarkers", "area", "pie", "doughnut", "scatter", "scatterSmooth", "scatterLines"]);

const SER_LINES = '<c:serLines><c:spPr><a:ln w="9525" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="tx1"><a:lumMod val="35000"/><a:lumOff val="65000"/></a:schemeClr></a:solidFill><a:round/></a:ln><a:effectLst/></c:spPr></c:serLines>';
const stockLine = (lum: string) => `<a:ln w="9525" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="tx1">${lum}</a:schemeClr></a:solidFill><a:round/></a:ln>`;
const LUM65 = '<a:lumMod val="65000"/><a:lumOff val="35000"/>';
const UP_DOWN = `<c:upDownBars><c:gapWidth val="150"/><c:upBars><c:spPr><a:solidFill><a:schemeClr val="lt1"/></a:solidFill>${stockLine(LUM65)}<a:effectLst/></c:spPr></c:upBars><c:downBars><c:spPr><a:solidFill><a:schemeClr val="dk1">${LUM65}</a:schemeClr></a:solidFill>${stockLine(LUM65)}<a:effectLst/></c:spPr></c:downBars></c:upDownBars>`;

/** One chart group element (c:barChart, c:radarChart …) of a chart's series. */
function groupXml(m: ChartModel, g: ChartGroup): string {
  const sub: ChartModel = { ...m, kind: g.kind };
  const series = g.idx.map((i) => ser(sub, m.series[i], i)).join("");
  const info = kindInfo(g.kind);
  const pieLike = info.family === "pie" || info.family === "ofPie" || info.family === "doughnut";
  const groupLabels = `<c:dLbls>${flags(false, pieLike)}</c:dLbls>`;
  const ids = g.secondary ? [AX_CAT2, AX_VAL2] : [AX_CAT, AX_VAL];
  const depthAx = (standard: boolean) => `<c:axId val="${standard ? AX_SER : 0}"/>`;
  const ax = ids.map((id) => `<c:axId val="${id}"/>`).join("");
  const varied = m.series.some((s) => s.fill && s.fill.type !== "auto") ? 0 : 1;
  switch (info.family) {
    case "bar": {
      const d = barDefaults(g.kind);
      const gap = m.gapWidth ?? d.gap, overlap = info.d3 ? undefined : m.overlap ?? d.overlap;
      const tag = info.d3 ? "c:bar3DChart" : "c:barChart";
      return `<${tag}><c:barDir val="${info.dir}"/><c:grouping val="${info.grouping}"/><c:varyColors val="0"/>${series}${groupLabels}<c:gapWidth val="${gap}"/>${overlap ? `<c:overlap val="${overlap}"/>` : ""}${info.d3 ? '<c:shape val="box"/>' : ""}${ax}${info.d3 ? depthAx(info.grouping === "standard") : ""}</${tag}>`;
    }
    case "line":
      if (info.d3) return `<c:line3DChart><c:grouping val="standard"/><c:varyColors val="0"/>${series}${groupLabels}${ax}${depthAx(true)}</c:line3DChart>`;
      return `<c:lineChart><c:grouping val="${info.grouping}"/><c:varyColors val="0"/>${series}${groupLabels}<c:marker val="1"/>${ax}</c:lineChart>`;
    case "area":
      if (info.d3) return `<c:area3DChart><c:grouping val="${info.grouping}"/><c:varyColors val="0"/>${series}${groupLabels}${ax}${depthAx(info.grouping === "standard")}</c:area3DChart>`;
      return `<c:areaChart><c:grouping val="${info.grouping}"/><c:varyColors val="0"/>${series}${groupLabels}${ax}</c:areaChart>`;
    case "pie":
      if (info.d3) return `<c:pie3DChart><c:varyColors val="${varied}"/>${series}${groupLabels}</c:pie3DChart>`;
      return `<c:pieChart><c:varyColors val="${varied}"/>${series}${groupLabels}<c:firstSliceAng val="${m.firstSliceAngle ?? 0}"/></c:pieChart>`;
    case "doughnut":
      return `<c:doughnutChart><c:varyColors val="${varied}"/>${series}${groupLabels}<c:firstSliceAng val="${m.firstSliceAngle ?? 0}"/><c:holeSize val="${m.holeSize ?? 75}"/></c:doughnutChart>`;
    case "ofPie": {
      const split = m.splitCount !== undefined ? `<c:splitType val="pos"/><c:splitPos val="${m.splitCount}"/>` : "";
      return `<c:ofPieChart><c:ofPieType val="${g.kind === "barOfPie" ? "bar" : "pie"}"/><c:varyColors val="${varied}"/>${series}${groupLabels}<c:gapWidth val="${m.gapWidth ?? 100}"/>${split}<c:secondPieSize val="${m.secondSize ?? 75}"/>${SER_LINES}</c:ofPieChart>`;
    }
    case "bubble":
      return `<c:bubbleChart><c:varyColors val="0"/>${series}${groupLabels}<c:bubbleScale val="${m.bubbleScale ?? 100}"/><c:showNegBubbles val="0"/>${ax}</c:bubbleChart>`;
    case "radar":
      return `<c:radarChart><c:radarStyle val="${g.kind === "radarFilled" ? "filled" : "marker"}"/><c:varyColors val="0"/>${series}${groupLabels}${ax}</c:radarChart>`;
    case "stock": {
      const updown = g.kind === "stockOHLC" || g.kind === "stockVOHLC" ? UP_DOWN : "";
      return `<c:stockChart>${series}${groupLabels}<c:hiLowLines><c:spPr>${stockLine('<a:lumMod val="75000"/><a:lumOff val="25000"/>')}<a:effectLst/></c:spPr></c:hiLowLines>${updown}${ax}</c:stockChart>`;
    }
    case "surface": {
      const tag = info.d3 ? "c:surface3DChart" : "c:surfaceChart";
      return `<${tag}><c:wireframe val="${g.kind === "surfaceWireframe" || g.kind === "surfaceContourWireframe" ? 1 : 0}"/>${series}${ax}${depthAx(true)}</${tag}>`;
    }
    default: {
      const style = info.lines === "smooth" ? "smoothMarker" : "lineMarker";
      return `<c:scatterChart><c:scatterStyle val="${style}"/><c:varyColors val="0"/>${series}${groupLabels}${ax}</c:scatterChart>`;
    }
  }
}

/** The plot area of every kind RaabeCalc did not write before, and of Verbund charts. */
function plotAreaGroups(m: ChartModel): string {
  const groups = chartGroups(m);
  const main = kindInfo(m.kind);
  const t = m.axisTitles, f = m.format, at = f?.axisTitles;
  const first = groups.find((g) => !g.secondary)?.idx[0] ?? 0;
  const vFmt = m.series[first]?.val.format ?? "General";
  const out = groups.map((g) => groupXml(m, g)).join("");
  if (main.family === "pie" || main.family === "ofPie" || main.family === "doughnut") return out;
  if (main.family === "scatter" || main.family === "bubble") {
    const cat = m.series[0]?.cat;
    const xFmt = cat && cat.cache.some((v) => typeof v === "string") ? null : cat?.format ?? "General";
    return out + valAx(xFmt, AX_CAT, AX_VAL, "b", t.cat, "midCat", true, m.gridlines, f?.catAxis, f?.valAxis, f, at?.cat, false)
      + valAx(vFmt, AX_VAL, AX_CAT, "l", t.val, "midCat", true, m.gridlines, f?.valAxis, f?.catAxis, f, at?.val);
  }
  const bar = main.dir === "bar";
  const between = main.family === "area" || main.family === "surface" ? "midCat" : "between";
  const vf = main.grouping === "percentStacked" ? "0%" : vFmt;
  let axes = catAx(bar ? "l" : "b", t.cat, f) + valAx(vf, AX_VAL, AX_CAT, bar ? "b" : "l", t.val, between, false, m.gridlines, f?.valAxis, f?.catAxis, f, at?.val);
  const g2 = groups.find((g) => g.secondary);
  if (g2) {
    const k2 = kindInfo(g2.kind);
    const fmt2 = k2.grouping === "percentStacked" ? "0%" : m.series[g2.idx[0]]?.val.format ?? "General";
    // Sekundärachse: the value axis on the other side (crossing at the maximum), its category axis hidden.
    axes += valAx(fmt2, AX_VAL2, AX_CAT2, bar ? "t" : "r", undefined, between, false, main.family === "stock", undefined, { crosses: "max" }, undefined, undefined, false)
      + catAx(bar ? "l" : "b", undefined, undefined, AX_CAT2, AX_VAL2, true);
  }
  const depth = main.family === "surface" || (!!main.d3 && main.grouping === "standard");
  return out + axes + (depth ? serAx() : "");
}

/** View3D, floor and walls of 3-D charts and surfaces (Excel's defaults: no fill, no line). */
function view3DXml(m: ChartModel): string {
  const v: View3D | undefined = view3DOf(m);
  if (!v) return "";
  const wall = (tag: string) => `<c:${tag}><c:thickness val="0"/><c:spPr><a:noFill/><a:ln><a:noFill/></a:ln><a:effectLst/><a:sp3d/></c:spPr></c:${tag}>`;
  return `<c:view3D><c:rotX val="${v.rotX}"/><c:rotY val="${v.rotY}"/>${v.depthPercent !== undefined ? `<c:depthPercent val="${v.depthPercent}"/>` : ""}<c:rAngAx val="${v.rAngAx ? 1 : 0}"/>${v.perspective !== undefined ? `<c:perspective val="${v.perspective}"/>` : ""}</c:view3D>${wall("floor")}${wall("sideWall")}${wall("backWall")}`;
}

function plotArea(m: ChartModel): string {
  if (!LEGACY.has(m.kind) || isCombo(m)) return plotAreaGroups(m);
  const series = m.series.map((s, i) => ser(m, s, i)).join("");
  const groupLabels = `<c:dLbls>${flags(false, m.kind === "pie" || m.kind === "doughnut")}</c:dLbls>`;
  const ax = `<c:axId val="${AX_CAT}"/><c:axId val="${AX_VAL}"/>`;
  const t = m.axisTitles;
  const vFmt = m.series[0]?.val.format ?? "General";
  const cat = m.series[0]?.cat;
  const xFmt = cat && cat.cache.some((v) => typeof v === "string") ? null : cat?.format ?? "General";
  const f = m.format, at = f?.axisTitles;
  const varied = m.series.some((s) => s.fill && s.fill.type !== "auto") ? 0 : 1;
  switch (m.kind) {
    case "column": case "columnStacked": case "bar": case "barStacked": {
      const bar = m.kind === "bar" || m.kind === "barStacked", stacked = m.kind === "columnStacked" || m.kind === "barStacked";
      const gap = m.gapWidth ?? (stacked ? 150 : bar ? 182 : 219), overlap = m.overlap ?? (stacked ? 100 : bar ? 0 : -27);
      return `<c:barChart><c:barDir val="${bar ? "bar" : "col"}"/><c:grouping val="${stacked ? "stacked" : "clustered"}"/><c:varyColors val="0"/>${series}${groupLabels}<c:gapWidth val="${gap}"/>${overlap ? `<c:overlap val="${overlap}"/>` : ""}${ax}</c:barChart>`
        + catAx(bar ? "l" : "b", t.cat, f) + valAx(vFmt, AX_VAL, AX_CAT, bar ? "b" : "l", t.val, "between", false, m.gridlines, f?.valAxis, f?.catAxis, f, at?.val);
    }
    case "line": case "lineMarkers":
      return `<c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/>${series}${groupLabels}<c:marker val="1"/>${ax}</c:lineChart>`
        + catAx("b", t.cat, f) + valAx(vFmt, AX_VAL, AX_CAT, "l", t.val, "between", false, m.gridlines, f?.valAxis, f?.catAxis, f, at?.val);
    case "area":
      return `<c:areaChart><c:grouping val="standard"/><c:varyColors val="0"/>${series}${groupLabels}${ax}</c:areaChart>`
        + catAx("b", t.cat, f) + valAx(vFmt, AX_VAL, AX_CAT, "l", t.val, "midCat", false, m.gridlines, f?.valAxis, f?.catAxis, f, at?.val);
    case "pie":
      return `<c:pieChart><c:varyColors val="${varied}"/>${series}${groupLabels}<c:firstSliceAng val="${m.firstSliceAngle ?? 0}"/></c:pieChart>`;
    case "doughnut":
      return `<c:doughnutChart><c:varyColors val="${varied}"/>${series}${groupLabels}<c:firstSliceAng val="${m.firstSliceAngle ?? 0}"/><c:holeSize val="${m.holeSize ?? 75}"/></c:doughnutChart>`;
    default: {
      const style = m.kind === "scatterSmooth" ? "smoothMarker" : "lineMarker";
      // Scatter: the X axis is the category axis' options, the Y axis the value axis'.
      return `<c:scatterChart><c:scatterStyle val="${style}"/><c:varyColors val="0"/>${series}${groupLabels}${ax}</c:scatterChart>`
        + valAx(xFmt, AX_CAT, AX_VAL, "b", t.cat, "midCat", true, m.gridlines, f?.catAxis, f?.valAxis, f, at?.cat, false)
        + valAx(vFmt, AX_VAL, AX_CAT, "l", t.val, "midCat", true, m.gridlines, f?.valAxis, f?.catAxis, f, at?.val);
    }
  }
}

/** The modelled parts of a chart (what decides whether a read chart's XML is still current). */
function signature(m: ChartModel): string {
  return JSON.stringify({
    kind: m.kind, title: m.title, axisTitles: m.axisTitles, legend: m.legend, dataLabels: m.dataLabels, gridlines: m.gridlines,
    format: m.format, gapWidth: m.gapWidth, overlap: m.overlap, firstSliceAngle: m.firstSliceAngle, holeSize: m.holeSize, colors: m.colors,
    view3D: m.view3D, splitCount: m.splitCount, secondSize: m.secondSize, bubbleScale: m.bubbleScale,
    hiddenCats: m.hiddenCats, dispBlanksAs: m.dispBlanksAs, showHidden: m.showHidden, naAsBlank: m.naAsBlank, dataTable: m.dataTable,
    series: m.series.map((s) => ({ name: s.name?.ref ?? s.name?.text, cat: s.cat?.ref, val: s.val.ref, color: s.color, trendline: s.trendline,
      fill: s.fill, line: s.line, marker: s.marker, smooth: s.smooth, invertIfNegative: s.invertIfNegative, errorBars: s.errorBars,
      kind: s.kind, secondary: s.secondary, size: s.size?.ref, totals: s.totals, filtered: s.filtered })),
  });
}

interface RawChart { xml: string; sig: string }

function rawOf(m: ChartModel): RawChart | undefined {
  const r = m.raw as Partial<RawChart> | undefined;
  return r && typeof r.xml === "string" && typeof r.sig === "string" ? (r as RawChart) : undefined;
}

/** The chart part (xl/charts/chartN.xml) for a model; histograms: the chartex part. */
export function chartToXml(model: ChartModel): string {
  const raw = rawOf(model);
  if (raw && raw.sig === signature(model)) return raw.xml;
  if (isChartEx(model)) return plainChartXml(model);
  // Diagrammfilter, Datentabelle, Ausgeblendete und leere Zellen in Excel's form (D-116, chartfilter.ts).
  const { model: shown, full } = withHiddenCats(model);
  return filterXml(plainChartXml(shown), model, full);
}

function plainChartXml(model: ChartModel): string {
  // A Farben ändern scheme other than Farbig: its colours are written as the series' own (no colour part is written).
  const m: ChartModel = model.colors && model.colors !== 10 && !isPie(model.kind)
    ? { ...model, series: model.series.map((s, i) => (s.color ? s : { ...s, color: paletteColor(model, i, model.series.length, OFFICE_ACCENTS) })) }
    : model;
  if (isChartEx(m)) return chartExToXml(m);
  const f = m.format;
  const title = m.title.show
    ? `<c:title>${m.title.text !== undefined ? rich(m.title.text, 1400) : ""}<c:overlay val="${f?.title?.overlay ? 1 : 0}"/>${spPrXml(f?.title, NONE_FILL, NONE_LINE)}${txPr(1400, "0", ' spc="0"', f?.title?.text)}</c:title><c:autoTitleDeleted val="0"/>`
    : '<c:autoTitleDeleted val="1"/>';
  const legend = m.legend === "none" ? "" : `<c:legend><c:legendPos val="${m.legend}"/><c:overlay val="${f?.legend?.overlay ? 1 : 0}"/>${spPrXml(f?.legend, NONE_FILL, NONE_LINE)}${txPr(900, "0", "", f?.legend?.text)}</c:legend>`;
  return HEAD + `<c:chartSpace xmlns:c="${NS_C}" xmlns:a="${NS_A}" xmlns:r="${NS_R}"><c:date1904 val="0"/><c:lang val="de-DE"/><c:roundedCorners val="0"/>`
    + `<c:chart>${title}${view3DXml(m)}<c:plotArea><c:layout/>${plotArea(m)}${spPrXml(f?.plot, NONE_FILL, NONE_LINE)}</c:plotArea>${legend}<c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/></c:chart>`
    + spPrXml(f?.chart, '<a:solidFill><a:schemeClr val="bg1"/></a:solidFill>', `<a:ln w="9525" cap="flat" cmpd="sng" algn="ctr">${LINE_GREY}<a:round/></a:ln>`)
    + `<c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr/></a:pPr><a:endParaRPr lang="de-DE"/></a:p></c:txPr>`
    + `<c:printSettings><c:headerFooter/><c:pageMargins b="0.78740157499999996" l="0.7" r="0.7" t="0.78740157499999996" header="0.3" footer="0.3"/><c:pageSetup/></c:printSettings></c:chartSpace>`;
}

/** A histogram as Excel 2016's chartex part (binning automatic, closed on the right). */
export function chartExToXml(m: ChartModel): string {
  if (m.kind !== "histogram") return chartExOtherXml(m);
  const s = m.series[0];
  const vals = s?.val.cache ?? [];
  const pts = vals.map((v, i) => (v === null ? "" : `<cx:pt idx="${i}">${num(v)}</cx:pt>`)).join("");
  const name = s?.name ? `<cx:tx><cx:txData>${s.name.ref ? `<cx:f>${esc(s.name.ref)}</cx:f>` : ""}<cx:v>${escText(s.name.text ?? "")}</cx:v></cx:txData></cx:tx>` : "";
  const title = m.title.show
    ? m.title.text !== undefined
      ? `<cx:title pos="t" align="ctr" overlay="0"><cx:tx><cx:txData><cx:v>${escText(m.title.text)}</cx:v></cx:txData></cx:tx><cx:txPr><a:bodyPr spcFirstLastPara="1" vertOverflow="ellipsis" horzOverflow="overflow" wrap="square" lIns="0" tIns="0" rIns="0" bIns="0" anchor="ctr" anchorCtr="1"/><a:lstStyle/><a:p><a:pPr algn="ctr" rtl="0"><a:defRPr/></a:pPr><a:r><a:rPr lang="de-DE"/><a:t>${escText(m.title.text)}</a:t></a:r></a:p></cx:txPr></cx:title>`
      : '<cx:title pos="t" align="ctr" overlay="0"/>'
    : "";
  const labels = m.dataLabels ? '<cx:dataLabels pos="outEnd"><cx:visibility seriesName="0" categoryName="0" value="1"/></cx:dataLabels>' : "";
  const grid = m.gridlines ? "<cx:majorGridlines/>" : "";
  const ref = s?.val.ref ? `<cx:f>${esc(s.val.ref)}</cx:f>` : "";
  return HEAD + `<cx:chartSpace xmlns:a="${NS_A}" xmlns:r="${NS_R}" xmlns:cx="${NS_CX}"><cx:chartData><cx:data id="0"><cx:numDim type="val">${ref}<cx:lvl ptCount="${vals.length}" formatCode="${esc(s?.val.format ?? "General")}">${pts}</cx:lvl></cx:numDim></cx:data></cx:chartData>`
    + `<cx:chart>${title}<cx:plotArea><cx:plotAreaRegion><cx:series layoutId="clusteredColumn" uniqueId="{00000000-5243-4D4C-8000-000000000001}">${name}${labels}<cx:dataId val="0"/><cx:layoutPr><cx:binning intervalClosed="r"/></cx:layoutPr></cx:series></cx:plotAreaRegion>`
    + `<cx:axis id="0"><cx:catScaling gapWidth="0"/>${axisTitleEx(m.axisTitles.cat)}<cx:tickLabels/></cx:axis><cx:axis id="1"><cx:valScaling/>${axisTitleEx(m.axisTitles.val)}${grid}<cx:tickLabels/></cx:axis></cx:plotArea>`
    + `${m.legend !== "none" ? `<cx:legend pos="${m.legend}" align="ctr" overlay="0"/>` : ""}</cx:chart></cx:chartSpace>`;
}

function axisTitleEx(t: string | undefined): string {
  if (t === undefined) return "";
  return `<cx:title><cx:tx><cx:txData><cx:v>${escText(t || "Achsentitel")}</cx:v></cx:txData></cx:tx></cx:title>`;
}

/** The chartex title as the histogram writes it. */
function titleEx(m: ChartModel): string {
  if (!m.title.show) return "";
  if (m.title.text === undefined) return '<cx:title pos="t" align="ctr" overlay="0"/>';
  return `<cx:title pos="t" align="ctr" overlay="0"><cx:tx><cx:txData><cx:v>${escText(m.title.text)}</cx:v></cx:txData></cx:tx><cx:txPr><a:bodyPr spcFirstLastPara="1" vertOverflow="ellipsis" horzOverflow="overflow" wrap="square" lIns="0" tIns="0" rIns="0" bIns="0" anchor="ctr" anchorCtr="1"/><a:lstStyle/><a:p><a:pPr algn="ctr" rtl="0"><a:defRPr/></a:pPr><a:r><a:rPr lang="de-DE"/><a:t>${escText(m.title.text)}</a:t></a:r></a:p></cx:txPr></cx:title>`;
}

/** Category levels of a chartex data block: cx:lvl elements, the innermost level first. */
function catDimEx(s: ChartSeries): string {
  const levels = s.levels ?? (s.cat ? [s.cat.cache.map((v) => (v === null ? null : String(v)))] : []);
  if (!levels.length) return "";
  const lvls = [...levels].reverse().map((lv) => `<cx:lvl ptCount="${lv.length}">${lv.map((v, i) => (v === null ? "" : `<cx:pt idx="${i}">${escText(String(v))}</cx:pt>`)).join("")}</cx:lvl>`).join("");
  return `<cx:strDim type="cat">${s.cat?.ref ? `<cx:f>${esc(s.cat.ref)}</cx:f>` : ""}${lvls}</cx:strDim>`;
}

function numDimEx(s: ChartSeries, type: "val" | "size"): string {
  const vals = s.val.cache;
  const pts = vals.map((v, i) => (v === null ? "" : `<cx:pt idx="${i}">${num(v)}</cx:pt>`)).join("");
  return `<cx:numDim type="${type}">${s.val.ref ? `<cx:f>${esc(s.val.ref)}</cx:f>` : ""}<cx:lvl ptCount="${vals.length}" formatCode="${esc(s.val.format ?? "General")}">${pts}</cx:lvl></cx:numDim>`;
}

function nameEx(s: ChartSeries | undefined): string {
  return s?.name ? `<cx:tx><cx:txData>${s.name.ref ? `<cx:f>${esc(s.name.ref)}</cx:f>` : ""}<cx:v>${escText(s.name.text ?? "")}</cx:v></cx:txData></cx:tx>` : "";
}

const uid = (i: number, n: number) => `{0000000${i.toString(16).toUpperCase()}-5243-4D4C-8000-00000000000${n}}`;

/** Wasserfall, Trichter, Treemap, Sunburst, Kastengrafik and Pareto as Excel 365 writes them (style 201). */
function chartExOtherXml(m: ChartModel): string {
  const k = m.kind;
  const hier = k === "treemap" || k === "sunburst";
  const list = k === "boxWhisker" ? m.series : m.series.slice(0, 1);
  const data = list.map((s, i) => `<cx:data id="${i}">${k === "boxWhisker" ? "" : catDimEx(s)}${numDimEx(s, hier ? "size" : "val")}</cx:data>`).join("");
  const lbl = (pos: string | null, cat: boolean, value: boolean) => (m.dataLabels
    ? `<cx:dataLabels${pos ? ` pos="${pos}"` : ""}><cx:visibility seriesName="0" categoryName="${cat ? 1 : 0}" value="${value ? 1 : 0}"/></cx:dataLabels>` : "");
  const grid = m.gridlines ? "<cx:majorGridlines/>" : "";
  const valAxis = (id: number) => `<cx:axis id="${id}"><cx:valScaling/>${axisTitleEx(m.axisTitles.val)}${grid}<cx:tickLabels/></cx:axis>`;
  const catAxis = (id: number, gap: string) => `<cx:axis id="${id}"><cx:catScaling gapWidth="${gap}"/>${axisTitleEx(m.axisTitles.cat)}<cx:tickLabels/></cx:axis>`;
  const s0 = list[0];
  let series = "", axes = "";
  switch (k) {
    case "waterfall": {
      const totals = s0?.totals?.length ? `<cx:subtotals>${s0.totals.map((t) => `<cx:idx val="${t}"/>`).join("")}</cx:subtotals>` : "";
      series = `<cx:series layoutId="waterfall" uniqueId="${uid(0, 1)}">${nameEx(s0)}${lbl("outEnd", false, true)}<cx:dataId val="0"/><cx:layoutPr><cx:visibility connectorLines="1"/>${totals}</cx:layoutPr></cx:series>`;
      axes = catAxis(0, "0.5") + valAxis(1);
      break;
    }
    case "funnel":
      series = `<cx:series layoutId="funnel" uniqueId="${uid(0, 2)}">${nameEx(s0)}${lbl(null, false, true)}<cx:dataId val="0"/></cx:series>`;
      axes = `<cx:axis id="1"><cx:catScaling gapWidth="0.06"/><cx:tickLabels/></cx:axis>`;
      break;
    case "treemap":
      series = `<cx:series layoutId="treemap" uniqueId="${uid(0, 3)}">${nameEx(s0)}${lbl("inEnd", true, false)}<cx:dataId val="0"/><cx:layoutPr><cx:parentLabelLayout val="overlapping"/></cx:layoutPr></cx:series>`;
      break;
    case "sunburst":
      series = `<cx:series layoutId="sunburst" uniqueId="${uid(0, 4)}">${nameEx(s0)}${lbl("ctr", true, false)}<cx:dataId val="0"/></cx:series>`;
      break;
    case "boxWhisker":
      series = list.map((s, i) => `<cx:series layoutId="boxWhisker" uniqueId="${uid(i, 5)}">${nameEx(s)}${lbl(null, false, true)}<cx:dataId val="${i}"/><cx:layoutPr><cx:visibility meanLine="0" meanMarker="1" nonoutliers="0" outliers="1"/><cx:statistics quartileMethod="exclusive"/></cx:layoutPr></cx:series>`).join("");
      axes = catAxis(0, "1") + valAxis(1);
      break;
    default: // pareto
      series = `<cx:series layoutId="clusteredColumn" uniqueId="${uid(0, 6)}">${nameEx(s0)}${lbl("outEnd", false, true)}<cx:dataId val="0"/><cx:layoutPr><cx:aggregation/></cx:layoutPr><cx:axisId val="1"/></cx:series><cx:series layoutId="paretoLine" ownerIdx="0" uniqueId="${uid(1, 6)}"><cx:axisId val="2"/></cx:series>`;
      axes = catAxis(0, "0") + valAxis(1) + `<cx:axis id="2"><cx:valScaling max="1" min="0"/><cx:units unit="percentage"/><cx:tickLabels/></cx:axis>`;
  }
  return HEAD + `<cx:chartSpace xmlns:a="${NS_A}" xmlns:r="${NS_R}" xmlns:cx="${NS_CX}"><cx:chartData>${data}</cx:chartData>`
    + `<cx:chart>${titleEx(m)}<cx:plotArea><cx:plotAreaRegion>${series}</cx:plotAreaRegion>${axes}</cx:plotArea>`
    + `${m.legend !== "none" ? `<cx:legend pos="${m.legend}" align="ctr" overlay="0"/>` : ""}</cx:chart></cx:chartSpace>`;
}

// ----- the chart's style and colour parts ------------------------------------------------------------
// Excel 2013 and later write a chart style (chartStyle) and colour style (colorStyle) part next to
// every chart; chartex parts (histograms) do not open without them. Ours describe the default
// look (style 201 and the "farbig" palette) entry by entry as the schema orders them.

export const CHART_STYLE_CONTENT_TYPE = "application/vnd.ms-office.chartstyle+xml";
export const CHART_STYLE_REL_TYPE = "http://schemas.microsoft.com/office/2011/relationships/chartStyle";
export const CHART_COLORS_CONTENT_TYPE = "application/vnd.ms-office.chartcolorstyle+xml";
export const CHART_COLORS_REL_TYPE = "http://schemas.microsoft.com/office/2011/relationships/chartColorStyle";

const NS_CS = "http://schemas.microsoft.com/office/drawing/2012/chartStyle";
const CS_GREY = '<a:schemeClr val="tx1"><a:lumMod val="65000"/><a:lumOff val="35000"/></a:schemeClr>';
const CS_LINE = (w: number, lum = "<a:lumMod val=\"15000\"/><a:lumOff val=\"85000\"/>") => `<a:ln w="${w}" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="tx1">${lum}</a:schemeClr></a:solidFill><a:round/></a:ln>`;
const PH_FILL = '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>';

function csEntry(name: string, o: { grey?: boolean; sz?: number; sp?: string; fillAuto?: boolean; lnAuto?: boolean; mods?: string; body?: boolean } = {}): string {
  const ln = o.lnAuto ? '<cs:lnRef idx="0"><cs:styleClr val="auto"/></cs:lnRef>' : '<cs:lnRef idx="0"/>';
  const fl = o.fillAuto ? '<cs:fillRef idx="0"><cs:styleClr val="auto"/></cs:fillRef>' : '<cs:fillRef idx="0"/>';
  const font = `<cs:fontRef idx="minor">${o.grey ? CS_GREY : '<a:schemeClr val="tx1"/>'}</cs:fontRef>`;
  const body = o.body ? '<cs:bodyPr rot="0" spcFirstLastPara="1" vertOverflow="clip" horzOverflow="clip" vert="horz" wrap="square" lIns="36576" tIns="18288" rIns="36576" bIns="18288" anchor="ctr" anchorCtr="1"><a:spAutoFit/></cs:bodyPr>' : "";
  return `<cs:${name}${o.mods ? ` mods="${o.mods}"` : ""}>${ln}${fl}<cs:effectRef idx="0"/>${font}${o.sp ? `<cs:spPr>${o.sp}</cs:spPr>` : ""}${o.sz ? `<cs:defRPr sz="${o.sz}"/>` : ""}${body}</cs:${name}>`;
}

/** The chart style part (xl/charts/styleN.xml): Excel's default look, style 201. */
export function chartStyleXml(id = 201): string {
  const line = CS_LINE(9525);
  const entries = [
    csEntry("axisTitle", { grey: true, sz: 1000 }),
    csEntry("categoryAxis", { grey: true, sp: line, sz: 900 }),
    csEntry("chartArea", { mods: "allowNoFillOverride allowNoLineOverride", sp: `<a:solidFill><a:schemeClr val="bg1"/></a:solidFill>${line}`, sz: 1000 }),
    csEntry("dataLabel", { grey: true, sz: 900 }),
    csEntry("dataLabelCallout", { grey: true, sp: `<a:solidFill><a:schemeClr val="bg1"/></a:solidFill>${CS_LINE(9525, '<a:lumMod val="25000"/><a:lumOff val="75000"/>')}`, sz: 900, body: true }),
    csEntry("dataPoint", { fillAuto: true, sp: PH_FILL }),
    csEntry("dataPoint3D", { fillAuto: true, sp: PH_FILL }),
    csEntry("dataPointLine", { lnAuto: true, sp: '<a:ln w="28575" cap="rnd"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:round/></a:ln>' }),
    csEntry("dataPointMarker", { fillAuto: true, sp: `${PH_FILL}<a:ln w="9525"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>` }),
    '<cs:dataPointMarkerLayout symbol="circle" size="5"/>',
    csEntry("dataPointWireframe", { lnAuto: true, sp: '<a:ln w="9525" cap="rnd"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:round/></a:ln>' }),
    csEntry("dataTable", { grey: true, sp: `<a:noFill/>${line}`, sz: 900 }),
    csEntry("downBar", { sp: `<a:solidFill><a:schemeClr val="dk1"><a:lumMod val="65000"/><a:lumOff val="35000"/></a:schemeClr></a:solidFill>${CS_LINE(9525, '<a:lumMod val="65000"/><a:lumOff val="35000"/>')}` }),
    csEntry("dropLine", { sp: CS_LINE(9525, '<a:lumMod val="35000"/><a:lumOff val="65000"/>') }),
    csEntry("errorBar", { sp: CS_LINE(9525, '<a:lumMod val="65000"/><a:lumOff val="35000"/>') }),
    csEntry("floor", { sp: `<a:noFill/><a:ln><a:noFill/></a:ln>` }),
    csEntry("gridlineMajor", { sp: line }),
    csEntry("gridlineMinor", { sp: CS_LINE(9525, '<a:lumMod val="5000"/><a:lumOff val="95000"/>') }),
    csEntry("hiLoLine", { sp: CS_LINE(9525, '<a:lumMod val="75000"/><a:lumOff val="25000"/>') }),
    csEntry("leaderLine", { sp: CS_LINE(9525, '<a:lumMod val="35000"/><a:lumOff val="65000"/>') }),
    csEntry("legend", { grey: true, sz: 900 }),
    csEntry("plotArea", { mods: "allowNoFillOverride allowNoLineOverride" }),
    csEntry("plotArea3D", { mods: "allowNoFillOverride allowNoLineOverride" }),
    csEntry("seriesAxis", { grey: true, sp: line, sz: 900 }),
    csEntry("seriesLine", { sp: CS_LINE(9525, '<a:lumMod val="35000"/><a:lumOff val="65000"/>') }),
    csEntry("title", { grey: true, sz: 1400 }),
    csEntry("trendline", { lnAuto: true, sp: '<a:ln w="19050" cap="rnd"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="sysDot"/></a:ln>' }),
    csEntry("trendlineLabel", { grey: true, sz: 900 }),
    csEntry("upBar", { sp: `<a:solidFill><a:schemeClr val="lt1"/></a:solidFill>${CS_LINE(9525, '<a:lumMod val="65000"/><a:lumOff val="35000"/>')}` }),
    csEntry("valueAxis", { grey: true, sz: 900 }),
    csEntry("wall", { sp: `<a:noFill/><a:ln><a:noFill/></a:ln>` }),
  ];
  return HEAD + `<cs:chartStyle xmlns:cs="${NS_CS}" xmlns:a="${NS_A}" id="${id}">${entries.join("")}</cs:chartStyle>`;
}

/** The colour style part (xl/charts/colorsN.xml): the accents in turn, then darker and lighter variants. */
export function chartColorsXml(): string {
  const accents = [1, 2, 3, 4, 5, 6].map((i) => `<a:schemeClr val="accent${i}"/>`).join("");
  const variations = ["", '<a:lumMod val="60000"/>', '<a:lumMod val="80000"/><a:lumOff val="20000"/>', '<a:lumMod val="80000"/>', '<a:lumMod val="60000"/><a:lumOff val="40000"/>', '<a:lumMod val="50000"/>', '<a:lumMod val="70000"/><a:lumOff val="30000"/>', '<a:lumMod val="70000"/>', '<a:lumMod val="50000"/><a:lumOff val="50000"/>']
    .map((v) => (v ? `<cs:variation>${v}</cs:variation>` : "<cs:variation/>")).join("");
  return HEAD + `<cs:colorStyle xmlns:cs="${NS_CS}" xmlns:a="${NS_A}" meth="cycle" id="10">${accents}${variations}</cs:colorStyle>`;
}

/** The chart part's own relationships (chartN.xml.rels) to its style and colour parts. */
export function chartRelsXml(styleTarget: string, colorsTarget: string): string {
  return HEAD + `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId2" Type="${CHART_COLORS_REL_TYPE}" Target="${esc(colorsTarget)}"/><Relationship Id="rId1" Type="${CHART_STYLE_REL_TYPE}" Target="${esc(styleTarget)}"/></Relationships>`;
}

// ----- the drawing's anchor -----------------------------------------------------------------------

export interface AnchorInfo {
  from: CellAnchor;
  to: CellAnchor;
  /** cNvPr id, unique in the drawing (Excel starts at 2). */
  id: number;
  name: string;
  /** Relationship id of the chart part in the drawing's .rels. */
  rId: string;
  /** A chartex part (histogram): wrapped in mc:AlternateContent as Excel writes it. */
  chartEx?: boolean;
}

function corner(tag: string, a: CellAnchor): string {
  return `<xdr:${tag}><xdr:col>${a.col}</xdr:col><xdr:colOff>${Math.round(a.x * EMU_PX)}</xdr:colOff><xdr:row>${a.row}</xdr:row><xdr:rowOff>${Math.round(a.y * EMU_PX)}</xdr:rowOff></xdr:${tag}>`;
}

/** One xdr:twoCellAnchor holding a chart's graphic frame. */
export function anchorXml(a: AnchorInfo): string {
  const nv = `<xdr:nvGraphicFramePr><xdr:cNvPr id="${a.id}" name="${esc(a.name)}"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm>`;
  if (a.chartEx) {
    const frame = `<xdr:graphicFrame macro="">${nv}<a:graphic><a:graphicData uri="${NS_CX}"><cx:chart xmlns:cx="${NS_CX}" xmlns:r="${NS_R}" r:id="${esc(a.rId)}"/></a:graphicData></a:graphic></xdr:graphicFrame>`;
    const fallback = '<xdr:sp macro="" textlink=""><xdr:nvSpPr><xdr:cNvPr id="0" name=""/><xdr:cNvSpPr><a:spLocks noTextEdit="1"/></xdr:cNvSpPr></xdr:nvSpPr><xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="4572000" cy="2743200"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:sp>';
    return `<xdr:twoCellAnchor>${corner("from", a.from)}${corner("to", a.to)}<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:Choice xmlns:cx1="http://schemas.microsoft.com/office/drawing/2015/9/8/chartex" Requires="cx1">${frame}</mc:Choice><mc:Fallback>${fallback}</mc:Fallback></mc:AlternateContent><xdr:clientData/></xdr:twoCellAnchor>`;
  }
  return `<xdr:twoCellAnchor>${corner("from", a.from)}${corner("to", a.to)}<xdr:graphicFrame macro="">${nv}<a:graphic><a:graphicData uri="${NS_C}"><c:chart xmlns:c="${NS_C}" xmlns:r="${NS_R}" r:id="${esc(a.rId)}"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor>`;
}

/** A whole drawing part (xl/drawings/drawingN.xml) of anchors. */
export function drawingXml(anchors: string[]): string {
  return HEAD + `<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="${NS_A}">${anchors.join("")}</xdr:wsDr>`;
}

/** Reads a twoCellAnchor (or oneCellAnchor with ext) holding a chart: corners, name, chart relationship. */
export function anchorFromXml(xml: string): (AnchorInfo & { editAs?: string }) | null {
  const el = root(xml);
  const pt = (tag: string): CellAnchor | null => {
    const c = child(xml, el, `xdr:${tag}`);
    if (!c) return null;
    const v = (n: string) => Number(text(xml, child(xml, c, `xdr:${n}`) ?? c) || 0);
    return { col: v("col"), row: v("row"), x: v("colOff") / EMU_PX, y: v("rowOff") / EMU_PX };
  };
  const from = pt("from"), to = pt("to");
  if (!from || !to) return null;
  const nvp = /<xdr:cNvPr\b[^>]*>/.exec(xml);
  const nv = nvp ? attrsOf(nvp[0]) : {};
  const chart = /<(?:c|cx):chart\b[^>]*>/.exec(xml);
  if (!chart) return null;
  const rid = /r:id="([^"]*)"/.exec(chart[0]);
  return {
    from, to, id: Number(nv.id ?? 2), name: nv.name ?? "Diagramm 1", rId: rid ? rid[1] : "",
    chartEx: chart[0].startsWith("<cx:"), editAs: attrs(xml, el).editAs,
  };
}

function attrsOf(tag: string): Record<string, string> {
  const out: Record<string, string> = {};
  for (const m of tag.matchAll(/([\w:]+)="([^"]*)"/g)) out[m[1]] = unesc(m[2]);
  return out;
}

// ----- reading ----------------------------------------------------------------------------------------

function kidsOf(xml: string, el: El): El[] {
  return children(xml, el.open, el.close);
}

function find(xml: string, el: El | undefined, ...path: string[]): El | undefined {
  let cur = el;
  for (const p of path) {
    if (!cur) return undefined;
    cur = kidsOf(xml, cur).find((k) => k.name === p);
  }
  return cur;
}

function val(xml: string, el: El | undefined): string | undefined {
  return el ? attrs(xml, el).val : undefined;
}

/** All text of a:t runs below an element (rich titles). */
function richText(xml: string, el: El): string {
  const paras: string[] = [];
  const body = xml.slice(el.start, el.end);
  for (const p of body.split(/<a:p>|<a:p\s[^>]*>/).slice(1)) {
    const ts = [...p.matchAll(/<a:t>([^<]*)<\/a:t>|<a:t\s[^>]*>([^<]*)<\/a:t>/g)].map((m) => unesc(m[1] ?? m[2] ?? ""));
    paras.push(ts.join(""));
  }
  return paras.join("\n");
}

function readData(xml: string, el: El | undefined): { ref?: string; cache: (string | number | null)[]; format?: string; numeric: boolean } | undefined {
  if (!el) return undefined;
  const holder = kidsOf(xml, el)[0];
  if (!holder) return undefined;
  const ref = find(xml, holder, "c:f");
  const cache = find(xml, holder, "c:numCache") ?? find(xml, holder, "c:strCache") ?? (holder.name === "c:numLit" || holder.name === "c:strLit" ? holder : undefined)
    ?? find(xml, holder, "c:multiLvlStrCache");
  const numeric = holder.name === "c:numRef" || holder.name === "c:numLit";
  const out: (string | number | null)[] = [];
  let format: string | undefined;
  if (cache) {
    const count = Number(val(xml, find(xml, cache, "c:ptCount")) ?? 0);
    for (let i = 0; i < count; i++) out.push(null);
    const fc = find(xml, cache, "c:formatCode");
    if (fc) format = unesc(text(xml, fc));
    for (const pt of kidsOf(xml, cache).filter((k) => k.name === "c:pt")) {
      const idx = Number(attrs(xml, pt).idx ?? 0);
      const v = find(xml, pt, "c:v");
      const s = v ? unesc(text(xml, v)) : "";
      out[idx] = numeric ? (s === "" || isNaN(Number(s)) ? null : Number(s)) : s;
    }
  }
  return { ref: ref ? unesc(text(xml, ref)) : undefined, cache: out, format, numeric };
}

function readColor(xml: string, spPr: El | undefined): string | undefined {
  if (!spPr) return undefined;
  const body = xml.slice(spPr.start, spPr.end);
  // Fill first (columns, pies), else the line's colour (lines, scatter).
  const m = /^<c:spPr>(?:<a:solidFill><a:srgbClr val="([0-9A-Fa-f]{6})"|<a:noFill\/>)?/.exec(body) ?? null;
  if (m?.[1]) return "#" + m[1].toUpperCase();
  const ln = /<a:ln\b[^>]*><a:solidFill><a:srgbClr val="([0-9A-Fa-f]{6})"/.exec(body);
  return ln ? "#" + ln[1].toUpperCase() : undefined;
}

const GROUPS = ["c:barChart", "c:lineChart", "c:areaChart", "c:pieChart", "c:doughnutChart", "c:scatterChart", "c:bar3DChart", "c:line3DChart", "c:area3DChart", "c:pie3DChart", "c:radarChart", "c:bubbleChart", "c:stockChart", "c:surfaceChart", "c:surface3DChart", "c:ofPieChart"];

/** The chart type of one chart group element (Kurs groups next to a column group: the Volumen types). */
function groupKind(xml: string, g: El, volume: boolean): ChartKind {
  const v = (...path: string[]) => val(xml, find(xml, g, ...path));
  const firstSer = find(xml, g, "c:ser");
  const sym = val(xml, find(xml, firstSer, "c:marker", "c:symbol"));
  const grouping = (v("c:grouping") ?? "standard") as "standard" | "clustered" | "stacked" | "percentStacked";
  switch (g.name) {
    case "c:barChart": case "c:bar3DChart":
      return kindOf("bar", { dir: v("c:barDir") === "bar" ? "bar" : "col", grouping, d3: g.name === "c:bar3DChart" }) ?? "column";
    case "c:lineChart": return kindOf("line", { grouping, markers: sym !== "none" }) ?? "line";
    case "c:line3DChart": return "line3D";
    case "c:areaChart": return kindOf("area", { grouping }) ?? "area";
    case "c:area3DChart": return kindOf("area", { grouping, d3: true }) ?? "area3D";
    case "c:pieChart": return "pie";
    case "c:pie3DChart": return "pie3D";
    case "c:ofPieChart": return v("c:ofPieType") === "bar" ? "barOfPie" : "pieOfPie";
    case "c:doughnutChart": return "doughnut";
    case "c:radarChart": return v("c:radarStyle") === "filled" ? "radarFilled" : sym === "none" ? "radar" : "radarMarkers";
    case "c:bubbleChart": return val(xml, find(xml, firstSer, "c:bubble3D")) === "1" ? "bubble3D" : "bubble";
    case "c:stockChart": {
      const n = kidsOf(xml, g).filter((k) => k.name === "c:ser").length;
      const open = !!find(xml, g, "c:upDownBars") || n >= 4;
      return volume ? (open ? "stockVOHLC" : "stockVHLC") : open ? "stockOHLC" : "stockHLC";
    }
    case "c:surfaceChart": return v("c:wireframe") === "1" ? "surfaceContourWireframe" : "surfaceContour";
    case "c:surface3DChart": return v("c:wireframe") === "1" ? "surfaceWireframe" : "surface3D";
    default: {
      // Scatter charts: Punkte (no line), lines smoothed or straight, with or without markers.
      const smooth = val(xml, find(xml, firstSer, "c:smooth")) === "1" || /smooth/.test(v("c:scatterStyle") ?? "");
      const noLine = firstSer ? /<c:spPr><a:ln\b[^>]*><a:noFill\/>/.test(xml.slice(firstSer.start, firstSer.end)) : true;
      if (noLine) return "scatter";
      if (smooth) return sym === "none" ? "scatterSmooth" : "scatterSmoothMarkers";
      return sym === "none" ? "scatterLinesNoMarkers" : "scatterLines";
    }
  }
}

/**
 * A chart part as a model (without its anchor: pass the one read from the drawing). Chart types
 * the model does not cover come back as the nearest kind, with the XML kept in `raw` so the chart
 * is written back unchanged unless it is edited.
 */
export function chartFromXml(xml: string, anchor?: { from: CellAnchor; to: CellAnchor; name?: string }): ChartModel {
  if (!xml.includes("c15:filtered") && !xml.includes("c15:fullRef") && !/<c:dTable>|<c:dispBlanksAs val="(?:zero|span)"|<c:plotVisOnly val="0"|dispNaAsBlank val="0"|<c:catAx>(?:(?!<\/c:catAx>)[\s\S])*<c:minorGridlines>/.test(xml)) {
    return plainChartFromXml(xml, anchor);
  }
  // Diagrammfilter and the display options (D-116, chartfilter.ts): read in the plain form, the file's XML kept.
  const r = unfilterXml(xml);
  const m = plainChartFromXml(r.xml, anchor);
  applyFilterRead(m, r);
  if (rawOf(m)) m.raw = { xml, sig: signature(m) } satisfies RawChart;
  return m;
}

function plainChartFromXml(xml: string, anchor?: { from: CellAnchor; to: CellAnchor; name?: string }): ChartModel {
  const from = anchor?.from ?? { col: 0, row: 0, x: 0, y: 0 }, to = anchor?.to ?? { col: 8, row: 15, x: 0, y: 0 };
  const rootEl = root(xml);
  if (rootEl.name === "cx:chartSpace") return chartExFromXml(xml, rootEl, from, to, anchor?.name);
  const chart = find(xml, rootEl, "c:chart");
  const plot = find(xml, chart, "c:plotArea");
  const groups = plot ? kidsOf(xml, plot).filter((k) => GROUPS.includes(k.name)) : [];
  const group = groups[0];
  // Kurs with Volumen: a column group and a stock group make one chart type.
  const stockGroup = groups.find((g) => g.name === "c:stockChart");
  const volume = !!stockGroup && groups.some((g) => g.name === "c:barChart");
  const kindsOf = groups.map((g) => groupKind(xml, g, volume));
  let kind: ChartKind = volume ? groupKind(xml, stockGroup!, true) : kindsOf[0] ?? "column";
  if (volume) kindsOf.fill(kind);
  // Value axes on the right (or top): the groups using them are on the Sekundärachse.
  const axesEls = plot ? kidsOf(xml, plot).filter((k) => /^c:(cat|val|date|ser)Ax$/.test(k.name)) : [];
  const secondaryIds = new Set(axesEls.filter((a) => a.name === "c:valAx" && ["r", "t"].includes(val(xml, find(xml, a, "c:axPos")) ?? "")).map((a) => val(xml, find(xml, a, "c:axId"))));
  const isSecondary = (g: El) => !volume && kidsOf(xml, g).some((k) => k.name === "c:axId" && secondaryIds.has(val(xml, k)));
  const scatter = group?.name === "c:scatterChart" || group?.name === "c:bubbleChart";
  const rd: Reader = { xml, find: (el, ...path) => find(xml, el, ...path), val: (el) => val(xml, el) };
  const series: ChartSeries[] = [];
  let dataLabels = !!group && val(xml, find(xml, group, "c:dLbls", "c:showVal")) === "1";
  const all = groups.flatMap((g, gi) => kidsOf(xml, g).filter((k) => k.name === "c:ser").map((s) => ({ s, g, gk: kindsOf[gi] })));
  for (const { s, g: grp, gk } of all) {
    const xy = grp.name === "c:scatterChart" || grp.name === "c:bubbleChart";
    const tx = find(xml, s, "c:tx");
    let name: ChartSeries["name"];
    if (tx) {
      const sr = find(xml, tx, "c:strRef");
      if (sr) {
        const f = find(xml, sr, "c:f"), v = find(xml, sr, "c:strCache", "c:pt", "c:v");
        name = { ref: f ? unesc(text(xml, f)) : undefined, text: v ? unesc(text(xml, v)) : "" };
      } else {
        const v = find(xml, tx, "c:v");
        name = { text: v ? unesc(text(xml, v)) : "" };
      }
    }
    const catEl = find(xml, s, xy ? "c:xVal" : "c:cat");
    const valEl = find(xml, s, xy ? "c:yVal" : "c:val");
    const cat = readData(xml, catEl);
    const v = readData(xml, valEl);
    const out: ChartSeries = { val: { ref: v?.ref, cache: (v?.cache ?? []).map((x) => (typeof x === "number" ? x : null)), format: v?.format } };
    if (name) out.name = name;
    if (cat) out.cat = { ref: cat.ref, cache: cat.cache, format: cat.format };
    const color = readColor(xml, find(xml, s, "c:spPr")) ?? readColor(xml, find(xml, s, "c:marker", "c:spPr"));
    if (color) out.color = color;
    const tl = find(xml, s, "c:trendline");
    const ttype = val(xml, find(xml, tl, "c:trendlineType"));
    if (tl && ttype && ["linear", "exp", "log", "poly", "power", "movingAvg"].includes(ttype)) {
      const t: Trendline = { type: ttype as Trendline["type"], showEq: val(xml, find(xml, tl, "c:dispEq")) === "1", showR2: val(xml, find(xml, tl, "c:dispRSqr")) === "1" };
      const n = (tag: string) => { const v = val(xml, find(xml, tl, tag)); return v === undefined ? undefined : Number(v); };
      if (n("c:order") !== undefined) t.order = n("c:order");
      if (n("c:period") !== undefined) t.period = n("c:period");
      if (n("c:forward")) t.forward = n("c:forward");
      if (n("c:backward")) t.backward = n("c:backward");
      if (n("c:intercept") !== undefined) t.intercept = n("c:intercept");
      const nm = find(xml, tl, "c:name");
      if (nm) t.name = unesc(text(xml, nm));
      const ln = readShape(rd, find(xml, tl, "c:spPr")).line;
      if (ln && ln.type === "solid") t.line = ln;
      out.trendline = t;
    }
    // The Format pane's settings of the series.
    const shape = readShape(rd, find(xml, s, "c:spPr"));
    if (shape.fill?.type === "none" || shape.fill?.transparency) out.fill = shape.fill;
    // Columns, bars and areas have no border by default, Punkte-scatter no line: "none" is automatic there.
    const fam = kindInfo(gk).family;
    const noLineByDefault = gk === "scatter" || fam === "area" || fam === "bar" || fam === "bubble" || fam === "surface" || fam === "stock" || gk === "radarFilled";
    // A line in the series' colour at the kind's own width is the automatic one (the colour is kept as `color`).
    const autoWidth = fam === "line" || fam === "radar" ? 2.25 : 1.5;
    const plain = shape.line?.type === "solid" && !shape.line.dash && !shape.line.transparency && (shape.line.width ?? autoWidth) === autoWidth && shape.line.color === (color ?? shape.line.color);
    if (shape.line && !plain && (shape.line.width !== undefined || shape.line.dash || shape.line.type === "none") && !(noLineByDefault && shape.line.type === "none")) out.line = shape.line;
    const mkEl = find(xml, s, "c:marker");
    if (mkEl) {
      const sym = val(xml, find(xml, mkEl, "c:symbol")) as MarkerSymbol | undefined, size = val(xml, find(xml, mkEl, "c:size"));
      const autoSym = ["line", "scatter", "radar", "stock"].includes(fam) && !kindInfo(gk).markers ? "none" : "circle";
      if ((sym && sym !== autoSym) || (size && size !== "5")) out.marker = { ...(sym ? { symbol: sym } : {}), ...(size ? { size: Number(size) } : {}) };
    }
    if (val(xml, find(xml, s, "c:smooth")) === "1" && kindInfo(gk).lines !== "smooth") out.smooth = true;
    if (gk !== kind) out.kind = gk;
    if (isSecondary(grp)) out.secondary = true;
    const size = readData(xml, find(xml, s, "c:bubbleSize"));
    if (size) out.size = { ref: size.ref, cache: size.cache.map((x) => (typeof x === "number" ? x : null)), format: size.format };
    if (val(xml, find(xml, s, "c:invertIfNegative")) === "1") out.invertIfNegative = true;
    const eb = find(xml, s, "c:errBars");
    if (eb) {
      const e: ErrorBars = { dir: (val(xml, find(xml, eb, "c:errBarType")) ?? "both") as ErrorBars["dir"], type: (val(xml, find(xml, eb, "c:errValType")) ?? "fixedVal") as ErrorBars["type"] };
      if (val(xml, find(xml, eb, "c:noEndCap")) === "1") e.noEndCap = true;
      const v = val(xml, find(xml, eb, "c:val"));
      if (v !== undefined) e.value = Number(v);
      for (const side of ["plus", "minus"] as const) {
        const d = readData(xml, find(xml, eb, `c:${side}`));
        if (d) e[side] = { ref: d.ref, cache: d.cache.map((x) => (typeof x === "number" ? x : null)), format: d.format };
      }
      out.errorBars = e;
    }
    if (val(xml, find(xml, s, "c:dLbls", "c:showVal")) === "1") dataLabels = true;
    series.push(out);
  }
  const titleEl = find(xml, chart, "c:title");
  const autoDeleted = val(xml, find(xml, chart, "c:autoTitleDeleted")) === "1";
  const title: ChartModel["title"] = { show: !!titleEl || (!autoDeleted && series.length === 1) };
  const titleTx = titleEl ? find(xml, titleEl, "c:tx") : undefined;
  if (titleTx) {
    const r = find(xml, titleTx, "c:rich");
    if (r) title.text = richText(xml, r);
    else {
      const v = find(xml, titleTx, "c:strRef", "c:strCache", "c:pt", "c:v");
      if (v) title.text = unesc(text(xml, v));
    }
  }
  const axisTitles: ChartModel["axisTitles"] = {};
  // Pies have no axes: keep the default for when the chart becomes another kind.
  let gridlines = isPie(kind);
  // The primary axes only (Excel writes them first; secondary ones follow, series axes are not modelled).
  const primaryAxes = (): El[] => {
    const out: El[] = [];
    let cats = 0, vals = 0;
    for (const a of axesEls) {
      if (a.name === "c:serAx") continue;
      if (a.name === "c:valAx" && (scatter ? vals < 2 : vals < 1)) { out.push(a); vals++; } else if (a.name !== "c:valAx" && cats < 1 && !scatter) { out.push(a); cats++; }
    }
    return out;
  };
  if (plot) {
    const axes = primaryAxes();
    for (const ax of axes) {
      const pos = val(xml, find(xml, ax, "c:axPos"));
      const isVal = ax.name === "c:valAx" && (!scatter || pos === "l" || pos === "r");
      const t = find(xml, ax, "c:title");
      if (t) {
        const r = find(xml, t, "c:tx", "c:rich");
        axisTitles[isVal ? "val" : "cat"] = r ? richText(xml, r) : "";
      }
      if (isVal && find(xml, ax, "c:majorGridlines")) gridlines = true;
    }
  }
  const legendEl = find(xml, chart, "c:legend");
  const legend = (legendEl ? (val(xml, find(xml, legendEl, "c:legendPos")) ?? "r") : "none") as LegendPos;
  const model: ChartModel = { kind, series, title, axisTitles, legend: ["b", "t", "l", "r"].includes(legend) ? legend : legend === "none" ? "none" : "r", dataLabels, gridlines, from, to };
  if (anchor?.name) model.name = anchor.name;
  // The Format pane's settings of the chart's elements.
  const fmt: ChartFormat = {};
  const put = <K extends keyof ChartFormat>(k: K, v: ChartFormat[K] | undefined) => { if (v && Object.keys(v).length) fmt[k] = v; };
  const space = root(xml);
  const chartSp = readShape(rd, find(xml, space, "c:spPr"));
  put("chart", chartSp.fill || (chartSp.line && chartSp.line.type !== "auto") ? chartSp : undefined);
  // Elements that have no fill and no line by default: "none" is their automatic look.
  const noneIsAuto = <T extends ElementFormat>(x: T | undefined): T | undefined => {
    if (!x) return x;
    if (x.fill?.type === "none") delete x.fill;
    if (x.line?.type === "none") delete x.line;
    return x;
  };
  put("plot", noneIsAuto(readFormat(rd, find(xml, plot, "c:spPr"), undefined, 0)));
  if (titleEl) {
    const tf = readFormat(rd, find(xml, titleEl, "c:spPr"), find(xml, titleEl, "c:txPr"), 1400) ?? {};
    if (val(xml, find(xml, titleEl, "c:overlay")) === "1") (tf as { overlay?: boolean }).overlay = true;
    if (tf.fill?.type === "none") delete tf.fill;
    if (tf.line?.type === "none") delete tf.line;
    put("title", tf);
  }
  if (legendEl) {
    const lf = readFormat(rd, find(xml, legendEl, "c:spPr"), find(xml, legendEl, "c:txPr"), 900) ?? {};
    if (val(xml, find(xml, legendEl, "c:overlay")) === "1") (lf as { overlay?: boolean }).overlay = true;
    if (lf.fill?.type === "none") delete lf.fill;
    if (lf.line?.type === "none") delete lf.line;
    put("legend", lf);
  }
  if (plot) {
    for (const ax of primaryAxes()) {
      const pos = val(xml, find(xml, ax, "c:axPos"));
      const isVal = ax.name === "c:valAx" && (!scatter || pos === "l" || pos === "r");
      const o = readAxis(rd, ax, ax.name === "c:valAx");
      if (o?.line?.type === "none" && isVal && !scatter) delete o.line;
      put(isVal ? "valAxis" : "catAxis", o);
      const mg = find(xml, ax, "c:majorGridlines"), ng = find(xml, ax, "c:minorGridlines");
      if (isVal && mg) { const l = readShape(rd, find(xml, mg, "c:spPr")).line; if (l && l.type !== "auto") put("majorGrid", { line: l }); }
      if (isVal && ng) put("minorGrid", { show: true, ...readShape(rd, find(xml, ng, "c:spPr")) });
      if (!isVal && !scatter && mg) put("catGrid", { show: true, ...readShape(rd, find(xml, mg, "c:spPr")) });
    }
  }
  const dl = group ? find(xml, group, "c:ser", "c:dLbls") ?? find(xml, group, "c:dLbls") : undefined;
  if (dl && dataLabels) {
    const o: DataLabelOptions = {};
    const b = (tag: string) => val(xml, find(xml, dl, tag));
    if (b("c:showVal") === "0") o.showVal = false;
    if (b("c:showCatName") === "1") o.showCat = true;
    if (b("c:showSerName") === "1") o.showSer = true;
    if (b("c:showPercent") === "1") o.showPercent = true;
    if (b("c:showLegendKey") === "1") o.showLegendKey = true;
    const pos = b("c:dLblPos");
    if (pos && pos !== defLabelPos(kind)) o.pos = pos as DataLabelOptions["pos"];
    const sep = find(xml, dl, "c:separator");
    if (sep) o.separator = unesc(text(xml, sep));
    const nf = find(xml, dl, "c:numFmt");
    if (nf && attrs(xml, nf).sourceLinked === "0") o.numFmt = attrs(xml, nf).formatCode;
    const t = readText(rd, find(xml, dl, "c:txPr"), 900);
    if (t) o.text = t;
    const shp = noneIsAuto(readShape(rd, find(xml, dl, "c:spPr")));
    if (shp?.fill) o.fill = shp.fill;
    if (shp?.line) o.line = shp.line;
    put("dataLabels", o);
  }
  if (Object.keys(fmt).length) model.format = fmt;
  if ((group?.name === "c:barChart" || group?.name === "c:bar3DChart") && !volume) {
    const gw = val(xml, find(xml, group, "c:gapWidth")), ov = val(xml, find(xml, group, "c:overlap"));
    const d = barDefaults(kind);
    if (gw !== undefined && Number(gw) !== d.gap) model.gapWidth = Number(gw);
    if (ov !== undefined && Number(ov) !== (d.overlap ?? 0)) model.overlap = Number(ov);
  }
  if (group?.name === "c:ofPieChart") {
    const gw = val(xml, find(xml, group, "c:gapWidth")), sp = val(xml, find(xml, group, "c:secondPieSize"));
    if (gw !== undefined && Number(gw) !== 100) model.gapWidth = Number(gw);
    if (sp !== undefined && Number(sp) !== 75) model.secondSize = Number(sp);
    const pos = val(xml, find(xml, group, "c:splitPos"));
    if (pos !== undefined && val(xml, find(xml, group, "c:splitType")) === "pos") model.splitCount = Number(pos);
  }
  if (group?.name === "c:bubbleChart") {
    const bs = val(xml, find(xml, group, "c:bubbleScale"));
    if (bs !== undefined && Number(bs) !== 100) model.bubbleScale = Number(bs);
  }
  const v3 = find(xml, chart, "c:view3D");
  if (v3) {
    const n = (tag: string) => { const x = val(xml, find(xml, v3, tag)); return x === undefined ? undefined : Number(x); };
    const v: View3D = { rotX: n("c:rotX") ?? 0, rotY: n("c:rotY") ?? 0, rAngAx: n("c:rAngAx") === 1 };
    if (n("c:depthPercent") !== undefined) v.depthPercent = n("c:depthPercent");
    if (n("c:perspective") !== undefined) v.perspective = n("c:perspective");
    const def = defaultView3D(kind);
    if (def && JSON.stringify(def) !== JSON.stringify(v)) model.view3D = v;
  }
  if (group && (kind === "pie" || kind === "doughnut")) {
    const fa = Number(val(xml, find(xml, group, "c:firstSliceAng")) ?? 0), hs = val(xml, find(xml, group, "c:holeSize"));
    if (fa) model.firstSliceAngle = fa;
    if (kind === "doughnut" && hs !== undefined && Number(hs) !== 75) model.holeSize = Number(hs);
  }
  model.raw = { xml, sig: signature(model) } satisfies RawChart;
  return model;
}

function chartExFromXml(xml: string, rootEl: El, from: CellAnchor, to: CellAnchor, name?: string): ChartModel {
  const body = xml.slice(rootEl.start, rootEl.end);
  // The data blocks: category levels (innermost first in the file) and numbers.
  interface Dim { ref?: string; levels: (string | null)[][]; nums: (number | null)[]; format?: string; numRef?: string }
  const dims = new Map<string, Dim>();
  for (const d of body.matchAll(/<cx:data id="(\d+)">([\s\S]*?)<\/cx:data>/g)) {
    const dim: Dim = { levels: [], nums: [] };
    const str = /<cx:strDim type="cat">([\s\S]*?)<\/cx:strDim>/.exec(d[2]);
    if (str) {
      const f = /<cx:f[^>]*>([^<]*)<\/cx:f>/.exec(str[1]);
      if (f) dim.ref = unesc(f[1]);
      for (const lv of str[1].matchAll(/<cx:lvl ptCount="(\d+)"[^>]*>([\s\S]*?)<\/cx:lvl>|<cx:lvl ptCount="(\d+)"[^>]*\/>/g)) {
        const out: (string | null)[] = new Array(Number(lv[1] ?? lv[3])).fill(null);
        for (const pt of (lv[2] ?? "").matchAll(/<cx:pt idx="(\d+)">([^<]*)<\/cx:pt>/g)) out[Number(pt[1])] = unesc(pt[2]);
        dim.levels.unshift(out);
      }
    }
    const nm = /<cx:numDim type="(?:val|size)">([\s\S]*?)<\/cx:numDim>/.exec(d[2]);
    if (nm) {
      const f = /<cx:f[^>]*>([^<]*)<\/cx:f>/.exec(nm[1]);
      if (f) dim.numRef = unesc(f[1]);
      const lvl = /<cx:lvl ptCount="(\d+)"(?: formatCode="([^"]*)")?[^>]*>([\s\S]*?)<\/cx:lvl>/.exec(nm[1]);
      if (lvl) {
        for (let i = 0; i < Number(lvl[1]); i++) dim.nums.push(null);
        for (const m of lvl[3].matchAll(/<cx:pt idx="(\d+)">([^<]*)<\/cx:pt>/g)) dim.nums[Number(m[1])] = Number(m[2]);
        if (lvl[2]) dim.format = unesc(lvl[2]);
      }
    }
    dims.set(d[1], dim);
  }
  const seriesEls = [...body.matchAll(/<cx:series\b([^>]*)>([\s\S]*?)<\/cx:series>/g)].map((m) => ({ attrs: m[1], body: m[2], layout: /layoutId="(\w+)"/.exec(m[1])?.[1] ?? "" }));
  const layouts = seriesEls.map((x) => x.layout);
  if (layouts.includes("regionMap")) throw new Error("Karte (map charts) are not modelled");
  const first = layouts.find((l) => l !== "paretoLine") ?? "clusteredColumn";
  const kind: ChartKind = layouts.includes("paretoLine") ? "pareto"
    : ({ waterfall: "waterfall", funnel: "funnel", treemap: "treemap", sunburst: "sunburst", boxWhisker: "boxWhisker" } as Record<string, ChartKind>)[first] ?? "histogram";
  const series: ChartSeries[] = [];
  for (const se of seriesEls) {
    if (se.layout === "paretoLine") continue;
    const id = /<cx:dataId val="(\d+)"/.exec(se.body)?.[1] ?? "0";
    const dim = dims.get(id) ?? { levels: [], nums: [] };
    const s: ChartSeries = { val: { ref: dim.numRef, cache: dim.nums, format: dim.format } };
    const tx = /^\s*<cx:tx><cx:txData>(?:<cx:f[^>]*>([^<]*)<\/cx:f>)?<cx:v>([^<]*)<\/cx:v>/.exec(se.body);
    if (tx) s.name = { ref: tx[1] ? unesc(tx[1]) : undefined, text: unesc(tx[2]) };
    if (kind !== "histogram" && (dim.ref || dim.levels.length)) {
      s.cat = { ref: dim.ref, cache: dim.levels.at(-1) ?? [] };
      if (dim.levels.length > 1) s.levels = dim.levels;
    }
    const totals = [...(/<cx:subtotals>([\s\S]*?)<\/cx:subtotals>/.exec(se.body)?.[1] ?? "").matchAll(/<cx:idx val="(\d+)"/g)].map((m) => Number(m[1]));
    if (totals.length) s.totals = totals;
    series.push(s);
  }
  const titleEl = /<cx:title\b[^>]*?(\/>|>([\s\S]*?)<\/cx:title>)/.exec(body.slice(body.indexOf("<cx:chart>")));
  const titleText = titleEl?.[2] ? /<cx:v>([^<]*)<\/cx:v>/.exec(titleEl[2]) : null;
  const legendPos = /<cx:legend\b[^>]*pos="(\w)"/.exec(body);
  const model: ChartModel = {
    kind, series,
    title: { show: !!titleEl, ...(titleText ? { text: unesc(titleText[1]) } : {}) },
    axisTitles: {}, legend: (legendPos?.[1] as LegendPos) ?? "none",
    dataLabels: /<cx:dataLabels\b/.test(body), gridlines: /<cx:majorGridlines\b/.test(body), from, to,
  };
  if (name) model.name = name;
  model.raw = { xml, sig: signature(model) } satisfies RawChart;
  return model;
}
