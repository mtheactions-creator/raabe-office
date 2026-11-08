// What the chart buttons and Datenquelle auswählen change, as Excel 365 writes it (D-115, D-116, D-117;
// recorded by tools/feel-lab/probe_excel_chart_filter.py, fixtures in test/fixtures/excel-chart-filter/):
// - a series hidden by Diagrammfilter moves into the chart group's extLst as c15:filtered…Series/c15:ser, its
//   references in c15:formulaRef instead of c:f;
// - a hidden category: every series' c:f lists the visible cells "(Daten!$A$2:$A$3,Daten!$A$5)", the whole range
//   stays in c15:fullRef, the caches hold the visible points only;
// - Ausgeblendete und leere Zellen: c:plotVisOnly, c:dispBlanksAs and c16r3:dispNaAsBlank;
// - the Datentabelle (c:dTable) and the category axis' Hilfsgitternetz (c:minorGridlines in c:catAx).
// The chart writer and reader stay as they are; this module turns their XML into Excel's form and back.

import { parseRef, refText, type ChartModel } from "@raabe/charts";

const C15_URI = "{02D57815-91ED-43cb-92C2-25804820EDAC}";
const NS_C15 = "http://schemas.microsoft.com/office/drawing/2012/chart";
const NA_URI = "{56B9EC1D-385E-4148-901F-78D8002777C0}";
const NS_C16R3 = "http://schemas.microsoft.com/office/drawing/2017/03/chart";

const GROUP_FILTER: Record<string, string> = {
  barChart: "filteredBarSeries", bar3DChart: "filteredBarSeries", lineChart: "filteredLineSeries", line3DChart: "filteredLineSeries",
  areaChart: "filteredAreaSeries", area3DChart: "filteredAreaSeries", pieChart: "filteredPieSeries", pie3DChart: "filteredPieSeries",
  doughnutChart: "filteredPieSeries", ofPieChart: "filteredPieSeries", scatterChart: "filteredScatterSeries", radarChart: "filteredRadarSeries",
  bubbleChart: "filteredBubbleSeries", surfaceChart: "filteredSurfaceSeries", surface3DChart: "filteredSurfaceSeries", stockChart: "filteredLineSeries",
};

const escAttr = (s: string) => s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
const unescX = (s: string) => s.replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, "&");

/** The visible cells of a one-row or one-column range without the hidden positions, in Excel's form. */
export function partialRef(full: string, hidden: number[]): string | null {
  const r = parseRef(full);
  if (!r) return null;
  const vertical = r.c0 === r.c1;
  const n = vertical ? r.r1 - r.r0 + 1 : r.c1 - r.c0 + 1;
  const skip = new Set(hidden);
  const runs: [number, number][] = [];
  for (let i = 0; i < n; i++) {
    if (skip.has(i)) continue;
    const last = runs[runs.length - 1];
    if (last && last[1] === i - 1) last[1] = i; else runs.push([i, i]);
  }
  if (!runs.length) return null;
  const parts = runs.map(([a, b]) => (vertical ? refText(r.sheet, r.r0 + a, r.c0, r.r0 + b, r.c0) : refText(r.sheet, r.r0, r.c0 + a, r.r0, r.c0 + b)));
  return parts.length === 1 ? parts[0] : `(${parts.join(",")})`;
}

/** Positions of a full range that a partial "(A,B)" form leaves out. */
export function hiddenOf(full: string, partial: string): number[] {
  const r = parseRef(full);
  if (!r) return [];
  const vertical = r.c0 === r.c1;
  const n = vertical ? r.r1 - r.r0 + 1 : r.c1 - r.c0 + 1;
  const shown = new Set<number>();
  for (const p of partial.replace(/^\(|\)$/g, "").split(",")) {
    const q = parseRef(p);
    if (!q) continue;
    for (let row = q.r0; row <= q.r1; row++) for (let col = q.c0; col <= q.c1; col++) shown.add(vertical ? row - r.r0 : col - r.c0);
  }
  return Array.from({ length: n }, (_, i) => i).filter((i) => !shown.has(i));
}

// ----- writing ---------------------------------------------------------------------------------------

/** A model whose hidden categories are taken out of every series (refs as Excel writes them, caches of the visible points). */
export function withHiddenCats(m: ChartModel): { model: ChartModel; full: Map<string, string> } {
  const full = new Map<string, string>();
  const hidden = m.hiddenCats ?? [];
  if (!hidden.length) return { model: m, full };
  const skip = new Set(hidden);
  const keep = <T>(a: T[]) => a.filter((_, i) => !skip.has(i));
  const data = <T extends { ref?: string; cache: unknown[] }>(d: T | undefined): T | undefined => {
    if (!d) return d;
    const out = { ...d, cache: keep(d.cache) } as T;
    if (d.ref) {
      const p = partialRef(d.ref, hidden);
      if (p && p !== d.ref) { full.set(p, d.ref); out.ref = p; }
    }
    return out;
  };
  return { model: { ...m, series: m.series.map((s) => ({ ...s, cat: data(s.cat), val: data(s.val)!, size: data(s.size) })) }, full };
}

/** Excel's form of a chart part written by chartToXml: filtered series, hidden categories, the chart's display options. */
export function filterXml(xml: string, m: ChartModel, full: Map<string, string>): string {
  // Hidden categories: the whole range in c15:fullRef before the partial c:f.
  for (const [part, whole] of full) {
    const f = `<c:f>${escAttr(part)}</c:f>`;
    xml = xml.split(f).join(`<c:extLst><c:ext uri="${C15_URI}" xmlns:c15="${NS_C15}"><c15:fullRef><c15:sqref>${escAttr(whole)}</c15:sqref></c15:fullRef></c:ext></c:extLst>${f}`);
  }
  // Filtered series: out of the group into its extLst.
  const filtered = m.series.map((s, i) => (s.filtered ? i : -1)).filter((i) => i >= 0);
  for (const i of filtered) {
    const start = xml.indexOf(`<c:ser><c:idx val="${i}"/>`);
    if (start < 0) continue;
    const end = xml.indexOf("</c:ser>", start) + "</c:ser>".length;
    let ser = xml.slice(start, end);
    xml = xml.slice(0, start) + xml.slice(end);
    // A hidden category's whole range and the visible cells share one extension in a filtered series (Excel).
    ser = ser.replace(/<c:extLst><c:ext uri="\{02D57815-91ED-43cb-92C2-25804820EDAC\}"[^>]*><c15:fullRef><c15:sqref>([^<]*)<\/c15:sqref><\/c15:fullRef><\/c:ext><\/c:extLst><c:f>([^<]*)<\/c:f>/g,
      (_, whole, part) => `<c:extLst><c:ext uri="${C15_URI}"><c15:fullRef><c15:sqref>${whole}</c15:sqref></c15:fullRef><c15:formulaRef><c15:sqref>${part}</c15:sqref></c15:formulaRef></c:ext></c:extLst>`);
    ser = "<c15:ser>" + ser.slice("<c:ser>".length, -"</c:ser>".length).replace(/<c:f>([^<]*)<\/c:f>/g,
      (_, ref) => `<c:extLst><c:ext uri="${C15_URI}"><c15:formulaRef><c15:sqref>${ref}</c15:sqref></c15:formulaRef></c:ext></c:extLst>`) + "</c15:ser>";
    // The group the series was in: the next closing group tag.
    const close = /<\/c:(\w+Chart)>/g;
    close.lastIndex = start;
    const g = close.exec(xml);
    if (!g) continue;
    const tag = GROUP_FILTER[g[1]] ?? "filteredBarSeries";
    const ext = `<c:ext uri="${C15_URI}" xmlns:c15="${NS_C15}"><c15:${tag}>${ser}</c15:${tag}></c:ext>`;
    const at = g.index;
    // Several filtered series of one group share the extension.
    const before = xml.slice(0, at);
    const openTag = `<c15:${tag}>`;
    const prevExt = before.lastIndexOf(`</c15:${tag}></c:ext></c:extLst>`);
    if (prevExt >= 0 && prevExt + `</c15:${tag}></c:ext></c:extLst>`.length === at) {
      xml = xml.slice(0, prevExt) + ser + xml.slice(prevExt);
    } else {
      xml = before + `<c:extLst>${ext}</c:extLst>` + xml.slice(at);
    }
    void openTag;
  }
  // Datentabelle: before the plot area's own c:spPr.
  if (m.dataTable) {
    const endPlot = xml.indexOf("</c:plotArea>");
    const sp = xml.lastIndexOf("<c:spPr>", endPlot);
    const table = `<c:dTable><c:showHorzBorder val="1"/><c:showVertBorder val="1"/><c:showOutline val="1"/><c:showKeys val="${m.dataTable.legendKeys ? 1 : 0}"/>`
      + `<c:spPr><a:noFill/><a:ln w="9525" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="tx1"><a:lumMod val="15000"/><a:lumOff val="85000"/></a:schemeClr></a:solidFill><a:round/></a:ln><a:effectLst/></c:spPr>`
      + `<c:txPr><a:bodyPr rot="0" spcFirstLastPara="1" vertOverflow="ellipsis" vert="horz" wrap="square" anchor="ctr" anchorCtr="1"/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="900" b="0" i="0" u="none" strike="noStrike" kern="1200" baseline="0"><a:solidFill><a:schemeClr val="tx1"><a:lumMod val="65000"/><a:lumOff val="35000"/></a:schemeClr></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:pPr><a:endParaRPr lang="de-DE"/></a:p></c:txPr></c:dTable>`;
    if (sp > 0 && endPlot > 0) xml = xml.slice(0, sp) + table + xml.slice(sp);
  }
  // Primäres Hilfsgitter of the category axis.
  if (m.format?.catMinorGrid?.show) {
    xml = xml.replace(/(<c:catAx>[\s\S]*?<c:axPos val="\w"\/>(?:<c:majorGridlines>[\s\S]*?<\/c:majorGridlines>)?)/, (a) =>
      a + '<c:minorGridlines><c:spPr><a:ln w="9525" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="tx1"><a:lumMod val="5000"/><a:lumOff val="95000"/></a:schemeClr></a:solidFill><a:round/></a:ln><a:effectLst/></c:spPr></c:minorGridlines>');
  }
  // Ausgeblendete und leere Zellen.
  const opts = `<c:plotVisOnly val="${m.showHidden ? 0 : 1}"/><c:dispBlanksAs val="${m.dispBlanksAs ?? "gap"}"/><c:extLst><c:ext uri="${NA_URI}" xmlns:c16r3="${NS_C16R3}"><c16r3:dataDisplayOptions16><c16r3:dispNaAsBlank val="${m.naAsBlank === false ? 0 : 1}"/></c16r3:dataDisplayOptions16></c:ext></c:extLst>`;
  xml = xml.replace('<c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/>', opts);
  return xml;
}

// ----- reading ---------------------------------------------------------------------------------------

export interface FilterRead {
  /** The XML in the reader's plain form. */
  xml: string;
  /** Positions (document order after sorting by c:order) of the series that were filtered. */
  filtered: number[];
  /** Partial refs of hidden categories, by their whole range. */
  partial: Map<string, string>;
  showHidden?: boolean;
  dispBlanksAs?: "gap" | "zero" | "span";
  naAsBlank?: boolean;
  dataTable?: { legendKeys: boolean };
  catMinorGrid?: boolean;
}

/** Turns Excel's form back into the reader's: filtered series into their group (by c:order), whole ranges into c:f. */
export function unfilterXml(xml: string): FilterRead {
  const out: FilterRead = { xml, filtered: [], partial: new Map() };
  const orig = xml;
  // Whole ranges of hidden categories; in a filtered series both forms share one extension.
  xml = xml.replace(/<c:extLst><c:ext uri="\{02D57815-91ED-43cb-92C2-25804820EDAC\}"[^>]*><c15:fullRef><c15:sqref>([^<]*)<\/c15:sqref><\/c15:fullRef><c15:formulaRef><c15:sqref>([^<]*)<\/c15:sqref><\/c15:formulaRef><\/c:ext><\/c:extLst>/g,
    (_, whole: string, part: string) => { out.partial.set(unescX(whole), unescX(part)); return `<c:f>${whole}</c:f>`; });
  xml = xml.replace(/<c:extLst><c:ext uri="\{02D57815-91ED-43cb-92C2-25804820EDAC\}"[^>]*><c15:fullRef><c15:sqref>([^<]*)<\/c15:sqref><\/c15:fullRef><\/c:ext><\/c:extLst><c:f>([^<]*)<\/c:f>/g,
    (_, whole: string, part: string) => { out.partial.set(unescX(whole), unescX(part)); return `<c:f>${whole}</c:f>`; });
  // Filtered series.
  let pos = 0;
  let serial = 0;
  const groupRe = /<c:(\w+Chart)>/g;
  let res = "";
  let gm: RegExpExecArray | null;
  while ((gm = groupRe.exec(xml))) {
    const name = gm[1];
    if (!GROUP_FILTER[name]) continue;
    const close = xml.indexOf(`</c:${name}>`, gm.index);
    if (close < 0) break;
    let body = xml.slice(gm.index, close);
    const ext = new RegExp(`<c:extLst><c:ext uri="\\{02D57815-91ED-43cb-92C2-25804820EDAC\\}"[^>]*><c15:filtered\\w+Series>([\\s\\S]*?)</c15:filtered\\w+Series></c:ext></c:extLst>$`).exec(body);
    const own = [...body.matchAll(/<c:ser>[\s\S]*?<\/c:ser>/g)].map((x) => ({ xml: x[0], filtered: false }));
    if (ext) {
      body = body.slice(0, ext.index);
      const more = [...ext[1].matchAll(/<c15:ser>([\s\S]*?)<\/c15:ser>/g)].map((x) => ({
        xml: "<c:ser>" + x[1].replace(/<c:extLst><c:ext uri="\{02D57815-91ED-43cb-92C2-25804820EDAC\}"[^>]*><c15:formulaRef><c15:sqref>([^<]*)<\/c15:sqref><\/c15:formulaRef><\/c:ext><\/c:extLst>/g, "<c:f>$1</c:f>") + "</c:ser>",
        filtered: true,
      }));
      const all = [...own, ...more];
      const order = (s: string) => Number(/<c:order val="(\d+)"\/>/.exec(s)?.[1] ?? 0);
      all.sort((a, b) => order(a.xml) - order(b.xml));
      // The series go where the group's own series were (or after its leading settings when all are filtered).
      const firstSer = body.indexOf("<c:ser>");
      const stripped = body.replace(/<c:ser>[\s\S]*?<\/c:ser>/g, "");
      const at = firstSer >= 0 ? firstSer : (() => {
        const v = /<c:varyColors val="\d"\/>/.exec(stripped) ?? /<c:(?:grouping|scatterStyle|radarStyle|wireframe) val="\w+"\/>/.exec(stripped);
        return v ? v.index + v[0].length : stripped.indexOf(">") + 1;
      })();
      body = stripped.slice(0, at) + all.map((x) => x.xml).join("") + stripped.slice(at);
      all.forEach((x, k) => { if (x.filtered) out.filtered.push(serial + k); });
      serial += all.length;
    } else {
      serial += own.length;
    }
    res += xml.slice(pos, gm.index) + body;
    pos = close;
    groupRe.lastIndex = close;
  }
  xml = res + xml.slice(pos);
  out.xml = xml;
  const v = (re: RegExp) => re.exec(orig)?.[1];
  if (v(/<c:plotVisOnly val="(\d)"\/>/) === "0") out.showHidden = true;
  const b = v(/<c:dispBlanksAs val="(\w+)"\/>/);
  if (b === "zero" || b === "span") out.dispBlanksAs = b;
  if (v(/<c16r3:dispNaAsBlank val="(\d)"\/>/) === "0") out.naAsBlank = false;
  const t = /<c:dTable>[\s\S]*?<c:showKeys val="(\d)"\/>/.exec(orig);
  if (t || orig.includes("<c:dTable>")) out.dataTable = { legendKeys: t ? t[1] === "1" : false };
  if (/<c:catAx>(?:(?!<\/c:catAx>)[\s\S])*<c:minorGridlines>/.test(orig)) out.catMinorGrid = true;
  return out;
}

/** Puts what unfilterXml found into the model the reader made. */
export function applyFilterRead(m: ChartModel, r: FilterRead): void {
  for (const i of r.filtered) if (m.series[i]) m.series[i].filtered = true;
  if (r.partial.size) {
    let hidden: number[] | null = null;
    for (const s of m.series) {
      for (const d of [s.cat, s.val, s.size]) {
        if (!d?.ref) continue;
        const part = r.partial.get(d.ref);
        if (!part) continue;
        const hid = hiddenOf(d.ref, part);
        hidden ??= hid;
        // The cache holds the visible points: spread it over the whole range.
        const vis = d.cache.slice();
        const set = new Set(hid);
        const n = vis.length + hid.length;
        const outCache: unknown[] = [];
        let k = 0;
        for (let i = 0; i < n; i++) outCache.push(set.has(i) ? null : vis[k++] ?? null);
        (d as { cache: unknown[] }).cache = outCache;
      }
    }
    if (hidden?.length) m.hiddenCats = hidden;
  }
  if (r.showHidden) m.showHidden = true;
  if (r.dispBlanksAs) m.dispBlanksAs = r.dispBlanksAs;
  if (r.naAsBlank === false) m.naAsBlank = false;
  if (r.dataTable) m.dataTable = r.dataTable;
  if (r.catMinorGrid) { m.format ??= {}; m.format.catMinorGrid = { ...(m.format.catMinorGrid ?? {}), show: true }; }
}
