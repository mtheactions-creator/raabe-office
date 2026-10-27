// Charts in .xlsx (Step 7 E, D-061): RaabeCalc's ChartModel to and from DrawingML's chart part
// (c:chartSpace, xl/charts/chartN.xml), Excel 2016's chartex part for histograms
// (cx:chartSpace, xl/charts/chartExN.xml) and the drawing's anchor (xdr:twoCellAnchor in
// xl/drawings/drawingN.xml). The XML is written as Excel 365 writes a new chart (style 201 of the
// Office 2013-2022 theme), so files open in Excel without repair and look the same there.
// A chart read from a file keeps its XML in `raw`; while the modelled parts stay unchanged it is
// written back verbatim (D-024), with Excel recalculating the caches from the cells on opening.

import { type CellAnchor, type ChartKind, type ChartModel, type ChartSeries, type LegendPos } from "@raabe/charts";
import { attrs, child, children, esc, escText, root, text, unesc, type El } from "./xml";

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

/** Histograms are chartex parts (another content and relationship type, see the constants). */
export function isChartEx(m: ChartModel): boolean {
  return m.kind === "histogram";
}

// ----- writing ----------------------------------------------------------------------------------

const GREY = '<a:solidFill><a:schemeClr val="tx1"><a:lumMod val="65000"/><a:lumOff val="35000"/></a:schemeClr></a:solidFill>';
const LINE_GREY = '<a:solidFill><a:schemeClr val="tx1"><a:lumMod val="15000"/><a:lumOff val="85000"/></a:schemeClr></a:solidFill>';
const NO_SHAPE = "<c:spPr><a:noFill/><a:ln><a:noFill/></a:ln><a:effectLst/></c:spPr>";

function txPr(size: number, rot = "-60000000", extra = ""): string {
  return `<c:txPr><a:bodyPr rot="${rot}" spcFirstLastPara="1" vertOverflow="ellipsis" vert="horz" wrap="square" anchor="ctr" anchorCtr="1"/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="${size}" b="0" i="0" u="none" strike="noStrike" kern="1200"${extra} baseline="0">${GREY}<a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:pPr><a:endParaRPr lang="de-DE"/></a:p></c:txPr>`;
}

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

/** Datenbeschriftungen of a series with Excel's default position for the chart type. */
function seriesLabels(kind: ChartKind): string {
  const pos: Partial<Record<ChartKind, string>> = {
    column: "outEnd", bar: "outEnd", columnStacked: "ctr", barStacked: "ctr", line: "r", lineMarkers: "r",
    scatter: "r", scatterSmooth: "r", scatterLines: "r", pie: "bestFit",
  };
  const p = pos[kind];
  return `<c:dLbls>${NO_SHAPE}${txPr(900, "0", "")}${p ? `<c:dLblPos val="${p}"/>` : ""}${flags(true, kind === "pie" || kind === "doughnut")}</c:dLbls>`;
}

function trendline(s: ChartSeries, i: number): string {
  const t = s.trendline;
  if (!t) return "";
  const lbl = t.showEq || t.showR2 ? `<c:trendlineLbl><c:numFmt formatCode="General" sourceLinked="0"/>${NO_SHAPE}${txPr(900, "0")}</c:trendlineLbl>` : "";
  return `<c:trendline><c:spPr><a:ln w="19050" cap="rnd">${fill(s, i)}<a:prstDash val="sysDot"/></a:ln><a:effectLst/></c:spPr><c:trendlineType val="linear"/><c:dispRSqr val="${t.showR2 ? 1 : 0}"/><c:dispEq val="${t.showEq ? 1 : 0}"/>${lbl}</c:trendline>`;
}

const MARKER_NONE = '<c:marker><c:symbol val="none"/></c:marker>';
function marker(s: ChartSeries, i: number): string {
  return `<c:marker><c:symbol val="circle"/><c:size val="5"/><c:spPr>${fill(s, i)}<a:ln w="9525">${fill(s, i)}</a:ln><a:effectLst/></c:spPr></c:marker>`;
}

function ser(m: ChartModel, s: ChartSeries, i: number): string {
  const head = `<c:ser><c:idx val="${i}"/><c:order val="${i}"/>${seriesText(s)}`;
  const labels = m.dataLabels ? seriesLabels(m.kind) : "";
  switch (m.kind) {
    case "column": case "columnStacked": case "bar": case "barStacked":
      return `${head}<c:spPr>${fill(s, i)}<a:ln><a:noFill/></a:ln><a:effectLst/></c:spPr><c:invertIfNegative val="0"/>${labels}${trendline(s, i)}${catData("cat", s.cat)}${numData("val", s.val)}</c:ser>`;
    case "line": case "lineMarkers":
      return `${head}<c:spPr><a:ln w="28575" cap="rnd">${fill(s, i)}<a:round/></a:ln><a:effectLst/></c:spPr>${m.kind === "line" ? MARKER_NONE : marker(s, i)}${labels}${trendline(s, i)}${catData("cat", s.cat)}${numData("val", s.val)}<c:smooth val="0"/></c:ser>`;
    case "area":
      return `${head}<c:spPr>${fill(s, i)}<a:ln><a:noFill/></a:ln><a:effectLst/></c:spPr>${labels}${trendline(s, i)}${catData("cat", s.cat)}${numData("val", s.val)}</c:ser>`;
    case "pie": case "doughnut": {
      const pts = s.val.cache.map((_, k) => `<c:dPt><c:idx val="${k}"/><c:bubble3D val="0"/><c:spPr><a:solidFill>${schemeColor(k)}</a:solidFill><a:ln w="19050"><a:solidFill><a:schemeClr val="lt1"/></a:solidFill></a:ln><a:effectLst/></c:spPr></c:dPt>`).join("");
      return `${head}${pts}${labels}${catData("cat", s.cat)}${numData("val", s.val)}</c:ser>`;
    }
    default: {
      // Scatter charts: Punkte (no line), with smoothed or straight lines.
      const line = m.kind === "scatter" ? '<a:ln w="19050" cap="rnd"><a:noFill/><a:round/></a:ln>' : `<a:ln w="19050" cap="rnd">${fill(s, i)}<a:round/></a:ln>`;
      return `${head}<c:spPr>${line}<a:effectLst/></c:spPr>${m.kind === "scatterSmooth" ? MARKER_NONE : marker(s, i)}${labels}${trendline(s, i)}${catData("xVal", s.cat)}${numData("yVal", s.val)}<c:smooth val="${m.kind === "scatterSmooth" ? 1 : 0}"/></c:ser>`;
    }
  }
}

function axisTitle(textValue: string | undefined, vertical: boolean): string {
  if (textValue === undefined) return "";
  return `<c:title>${rich(textValue || "Achsentitel", 1000, vertical ? "-5400000" : "0")}<c:overlay val="0"/>${NO_SHAPE}${txPr(1000, vertical ? "-5400000" : "0")}</c:title>`;
}

const AX_CAT = 500000001, AX_VAL = 500000002;

function catAx(pos: string, titleText: string | undefined): string {
  return `<c:catAx><c:axId val="${AX_CAT}"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="${pos}"/>${axisTitle(titleText, pos === "l")}<c:numFmt formatCode="General" sourceLinked="1"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:spPr><a:noFill/><a:ln w="9525" cap="flat" cmpd="sng" algn="ctr">${LINE_GREY}<a:round/></a:ln><a:effectLst/></c:spPr>${txPr(900)}<c:crossAx val="${AX_VAL}"/><c:crosses val="autoZero"/><c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/><c:noMultiLvlLbl val="0"/></c:catAx>`;
}

/** A value axis; `fmt` null leaves out the number format (Excel: X values that are text). */
function valAx(fmt: string | null, id: number, cross: number, pos: string, titleText: string | undefined, between: "between" | "midCat", line: boolean, grid: boolean): string {
  const gridXml = grid ? `<c:majorGridlines><c:spPr><a:ln w="9525" cap="flat" cmpd="sng" algn="ctr">${LINE_GREY}<a:round/></a:ln><a:effectLst/></c:spPr></c:majorGridlines>` : "";
  const lineXml = line
    ? `<c:spPr><a:noFill/><a:ln w="9525" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="tx1"><a:lumMod val="25000"/><a:lumOff val="75000"/></a:schemeClr></a:solidFill><a:round/></a:ln><a:effectLst/></c:spPr>`
    : "<c:spPr><a:noFill/><a:ln><a:noFill/></a:ln><a:effectLst/></c:spPr>";
  return `<c:valAx><c:axId val="${id}"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="${pos}"/>${gridXml}${axisTitle(titleText, pos === "l")}${fmt === null ? "" : `<c:numFmt formatCode="${esc(fmt)}" sourceLinked="1"/>`}<c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/>${lineXml}${txPr(900)}<c:crossAx val="${cross}"/><c:crosses val="autoZero"/><c:crossBetween val="${between}"/></c:valAx>`;
}

function plotArea(m: ChartModel): string {
  const series = m.series.map((s, i) => ser(m, s, i)).join("");
  const groupLabels = `<c:dLbls>${flags(false, m.kind === "pie" || m.kind === "doughnut")}</c:dLbls>`;
  const ax = `<c:axId val="${AX_CAT}"/><c:axId val="${AX_VAL}"/>`;
  const t = m.axisTitles;
  const vFmt = m.series[0]?.val.format ?? "General";
  const cat = m.series[0]?.cat;
  const xFmt = cat && cat.cache.some((v) => typeof v === "string") ? null : cat?.format ?? "General";
  switch (m.kind) {
    case "column": case "columnStacked": case "bar": case "barStacked": {
      const bar = m.kind === "bar" || m.kind === "barStacked", stacked = m.kind === "columnStacked" || m.kind === "barStacked";
      const gap = stacked ? 150 : bar ? 182 : 219, overlap = stacked ? 100 : bar ? 0 : -27;
      return `<c:barChart><c:barDir val="${bar ? "bar" : "col"}"/><c:grouping val="${stacked ? "stacked" : "clustered"}"/><c:varyColors val="0"/>${series}${groupLabels}<c:gapWidth val="${gap}"/>${overlap ? `<c:overlap val="${overlap}"/>` : ""}${ax}</c:barChart>`
        + catAx(bar ? "l" : "b", t.cat) + valAx(vFmt, AX_VAL, AX_CAT, bar ? "b" : "l", t.val, "between", false, m.gridlines);
    }
    case "line": case "lineMarkers":
      return `<c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/>${series}${groupLabels}<c:marker val="1"/>${ax}</c:lineChart>`
        + catAx("b", t.cat) + valAx(vFmt, AX_VAL, AX_CAT, "l", t.val, "between", false, m.gridlines);
    case "area":
      return `<c:areaChart><c:grouping val="standard"/><c:varyColors val="0"/>${series}${groupLabels}${ax}</c:areaChart>`
        + catAx("b", t.cat) + valAx(vFmt, AX_VAL, AX_CAT, "l", t.val, "midCat", false, m.gridlines);
    case "pie":
      return `<c:pieChart><c:varyColors val="1"/>${series}${groupLabels}<c:firstSliceAng val="0"/></c:pieChart>`;
    case "doughnut":
      return `<c:doughnutChart><c:varyColors val="1"/>${series}${groupLabels}<c:firstSliceAng val="0"/><c:holeSize val="75"/></c:doughnutChart>`;
    default: {
      const style = m.kind === "scatterSmooth" ? "smoothMarker" : "lineMarker";
      return `<c:scatterChart><c:scatterStyle val="${style}"/><c:varyColors val="0"/>${series}${groupLabels}${ax}</c:scatterChart>`
        + valAx(xFmt, AX_CAT, AX_VAL, "b", t.cat, "midCat", true, m.gridlines) + valAx(vFmt, AX_VAL, AX_CAT, "l", t.val, "midCat", true, m.gridlines);
    }
  }
}

/** The modelled parts of a chart (what decides whether a read chart's XML is still current). */
function signature(m: ChartModel): string {
  return JSON.stringify({
    kind: m.kind, title: m.title, axisTitles: m.axisTitles, legend: m.legend, dataLabels: m.dataLabels, gridlines: m.gridlines,
    series: m.series.map((s) => ({ name: s.name?.ref ?? s.name?.text, cat: s.cat?.ref, val: s.val.ref, color: s.color, trendline: s.trendline })),
  });
}

interface RawChart { xml: string; sig: string }

function rawOf(m: ChartModel): RawChart | undefined {
  const r = m.raw as Partial<RawChart> | undefined;
  return r && typeof r.xml === "string" && typeof r.sig === "string" ? (r as RawChart) : undefined;
}

/** The chart part (xl/charts/chartN.xml) for a model; histograms: the chartex part. */
export function chartToXml(m: ChartModel): string {
  const raw = rawOf(m);
  if (raw && raw.sig === signature(m)) return raw.xml;
  if (isChartEx(m)) return chartExToXml(m);
  const title = m.title.show
    ? `<c:title>${m.title.text !== undefined ? rich(m.title.text, 1400) : ""}<c:overlay val="0"/>${NO_SHAPE}${txPr(1400, "0", ' spc="0"')}</c:title><c:autoTitleDeleted val="0"/>`
    : '<c:autoTitleDeleted val="1"/>';
  const legend = m.legend === "none" ? "" : `<c:legend><c:legendPos val="${m.legend}"/><c:overlay val="0"/>${NO_SHAPE}${txPr(900, "0")}</c:legend>`;
  return HEAD + `<c:chartSpace xmlns:c="${NS_C}" xmlns:a="${NS_A}" xmlns:r="${NS_R}"><c:date1904 val="0"/><c:lang val="de-DE"/><c:roundedCorners val="0"/>`
    + `<c:chart>${title}<c:plotArea><c:layout/>${plotArea(m)}${NO_SHAPE}</c:plotArea>${legend}<c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/></c:chart>`
    + `<c:spPr><a:solidFill><a:schemeClr val="bg1"/></a:solidFill><a:ln w="9525" cap="flat" cmpd="sng" algn="ctr">${LINE_GREY}<a:round/></a:ln><a:effectLst/></c:spPr>`
    + `<c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr/></a:pPr><a:endParaRPr lang="de-DE"/></a:p></c:txPr>`
    + `<c:printSettings><c:headerFooter/><c:pageMargins b="0.78740157499999996" l="0.7" r="0.7" t="0.78740157499999996" header="0.3" footer="0.3"/><c:pageSetup/></c:printSettings></c:chartSpace>`;
}

/** A histogram as Excel 2016's chartex part (binning automatic, closed on the right). */
export function chartExToXml(m: ChartModel): string {
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

/**
 * A chart part as a model (without its anchor: pass the one read from the drawing). Chart types
 * the model does not cover come back as the nearest kind, with the XML kept in `raw` so the chart
 * is written back unchanged unless it is edited.
 */
export function chartFromXml(xml: string, anchor?: { from: CellAnchor; to: CellAnchor; name?: string }): ChartModel {
  const from = anchor?.from ?? { col: 0, row: 0, x: 0, y: 0 }, to = anchor?.to ?? { col: 8, row: 15, x: 0, y: 0 };
  const rootEl = root(xml);
  if (rootEl.name === "cx:chartSpace") return chartExFromXml(xml, rootEl, from, to, anchor?.name);
  const chart = find(xml, rootEl, "c:chart");
  const plot = find(xml, chart, "c:plotArea");
  const group = plot ? kidsOf(xml, plot).find((k) => GROUPS.includes(k.name)) : undefined;
  let kind: ChartKind = "column";
  if (group) {
    const g = group.name;
    if (g === "c:barChart" || g === "c:bar3DChart") {
      const bar = val(xml, find(xml, group, "c:barDir")) === "bar";
      const stacked = /stacked/i.test(val(xml, find(xml, group, "c:grouping")) ?? "");
      kind = bar ? (stacked ? "barStacked" : "bar") : stacked ? "columnStacked" : "column";
    } else if (g === "c:lineChart" || g === "c:line3DChart" || g === "c:radarChart" || g === "c:stockChart") {
      const firstSer = find(xml, group, "c:ser");
      const sym = val(xml, find(xml, firstSer, "c:marker", "c:symbol"));
      kind = sym === "none" ? "line" : "lineMarkers";
    } else if (g === "c:areaChart" || g === "c:area3DChart" || g === "c:surfaceChart" || g === "c:surface3DChart") kind = "area";
    else if (g === "c:pieChart" || g === "c:pie3DChart" || g === "c:ofPieChart") kind = "pie";
    else if (g === "c:doughnutChart") kind = "doughnut";
    else if (g === "c:scatterChart" || g === "c:bubbleChart") {
      const firstSer = find(xml, group, "c:ser");
      const smooth = val(xml, find(xml, firstSer, "c:smooth")) === "1" || /smooth/.test(val(xml, find(xml, group, "c:scatterStyle")) ?? "");
      const noLine = firstSer ? /<c:spPr><a:ln\b[^>]*><a:noFill\/>/.test(xml.slice(firstSer.start, firstSer.end)) : true;
      kind = noLine ? "scatter" : smooth ? "scatterSmooth" : "scatterLines";
    }
  }
  const scatter = group?.name === "c:scatterChart" || group?.name === "c:bubbleChart";
  const series: ChartSeries[] = [];
  let dataLabels = !!group && val(xml, find(xml, group, "c:dLbls", "c:showVal")) === "1";
  for (const s of group ? kidsOf(xml, group).filter((k) => k.name === "c:ser") : []) {
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
    const catEl = find(xml, s, scatter ? "c:xVal" : "c:cat");
    const valEl = find(xml, s, scatter ? "c:yVal" : "c:val");
    const cat = readData(xml, catEl);
    const v = readData(xml, valEl);
    const out: ChartSeries = { val: { ref: v?.ref, cache: (v?.cache ?? []).map((x) => (typeof x === "number" ? x : null)), format: v?.format } };
    if (name) out.name = name;
    if (cat) out.cat = { ref: cat.ref, cache: cat.cache, format: cat.format };
    const color = readColor(xml, find(xml, s, "c:spPr")) ?? readColor(xml, find(xml, s, "c:marker", "c:spPr"));
    if (color) out.color = color;
    const tl = find(xml, s, "c:trendline");
    if (tl && val(xml, find(xml, tl, "c:trendlineType")) === "linear") {
      out.trendline = { type: "linear", showEq: val(xml, find(xml, tl, "c:dispEq")) === "1", showR2: val(xml, find(xml, tl, "c:dispRSqr")) === "1" };
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
  let gridlines = kind === "pie" || kind === "doughnut";
  if (plot) {
    const axes = kidsOf(xml, plot).filter((k) => k.name === "c:catAx" || k.name === "c:valAx" || k.name === "c:dateAx");
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
  model.raw = { xml, sig: signature(model) } satisfies RawChart;
  return model;
}

function chartExFromXml(xml: string, rootEl: El, from: CellAnchor, to: CellAnchor, name?: string): ChartModel {
  const body = xml.slice(rootEl.start, rootEl.end);
  const f = /<cx:numDim type="val">\s*<cx:f[^>]*>([^<]*)<\/cx:f>/.exec(body);
  const pts: (number | null)[] = [];
  const lvl = /<cx:numDim type="val">[\s\S]*?<cx:lvl ptCount="(\d+)"(?: formatCode="([^"]*)")?>([\s\S]*?)<\/cx:lvl>/.exec(body);
  if (lvl) {
    for (let i = 0; i < Number(lvl[1]); i++) pts.push(null);
    for (const m of lvl[3].matchAll(/<cx:pt idx="(\d+)">([^<]*)<\/cx:pt>/g)) pts[Number(m[1])] = Number(m[2]);
  }
  const tx = /<cx:series\b[^>]*>\s*<cx:tx><cx:txData>(?:<cx:f[^>]*>([^<]*)<\/cx:f>)?<cx:v>([^<]*)<\/cx:v>/.exec(body);
  const s: ChartSeries = { val: { ref: f ? unesc(f[1]) : undefined, cache: pts, format: lvl?.[2] ? unesc(lvl[2]) : undefined } };
  if (tx) s.name = { ref: tx[1] ? unesc(tx[1]) : undefined, text: unesc(tx[2]) };
  const titleEl = /<cx:title\b[^>]*?(\/>|>([\s\S]*?)<\/cx:title>)/.exec(body.slice(body.indexOf("<cx:chart>")));
  const titleText = titleEl?.[2] ? /<cx:v>([^<]*)<\/cx:v>/.exec(titleEl[2]) : null;
  const legendPos = /<cx:legend\b[^>]*pos="(\w)"/.exec(body);
  const model: ChartModel = {
    kind: "histogram", series: [s],
    title: { show: !!titleEl, ...(titleText ? { text: unesc(titleText[1]) } : {}) },
    axisTitles: {}, legend: (legendPos?.[1] as LegendPos) ?? "none",
    dataLabels: /<cx:dataLabels\b/.test(body), gridlines: /<cx:majorGridlines\b/.test(body), from, to,
  };
  if (name) model.name = name;
  // Other chartex types (Wasserfall, Treemap …) are not modelled: they draw as a histogram of
  // their values here and are written back unchanged unless edited.
  model.raw = { xml, sig: signature(model) } satisfies RawChart;
  return model;
}
