// Diagrammformatvorlagen (D-118): the gallery of the brush button and of Diagrammentwurf. Each style is what German
// Excel 365 saved after the gallery's entry was applied (chart-styles.gen.ts, gen_chart_styles.py): fills, lines and
// text of every element, which elements show, data labels, gap width. A style is applied as the difference to the
// kind's first style (Excel's default look, which the chart engine draws without settings), colours resolved in the
// workbook's theme. Gradient and pattern fills are drawn with their first colour.

import { kindInfo, isPie, lum, type ChartModel, type DataLabelOptions, type ElementFormat, type Fill, type Line, type TextFormat } from "@raabe/charts";
import { CHART_STYLES } from "./chart-styles.gen";

type ColorRef = { s?: string; rgb?: string; lumMod?: number; lumOff?: number; alpha?: number; shade?: number; tint?: number };
interface FillRef { none?: boolean; c?: ColorRef | null; approx?: string }
interface LineRef { w?: number; none?: boolean; c?: ColorRef | null; dash?: string }
interface TextRef { size?: number; bold?: boolean; c?: ColorRef | null }
interface ElRef { fill?: FillRef; line?: LineRef; text?: TextRef; off?: boolean; hidden?: boolean; on?: boolean; major?: ElRef; minor?: ElRef; pos?: string; symbol?: string | null; size?: number; marker?: ElRef; showCatName?: boolean; showPercent?: boolean }
interface StyleRef { id: number; chart: ElRef; plot: ElRef; title: ElRef; legend: ElRef & { pos: string }; cat: ElRef; val: ElRef; series: ElRef[]; labels?: ElRef; gap?: number; overlap?: number }

let theme: Record<string, string> = { dk1: "#000000", lt1: "#FFFFFF", dk2: "#44546A", lt2: "#E7E6E6",
  accent1: "#4472C4", accent2: "#ED7D31", accent3: "#A5A5A5", accent4: "#FFC000", accent5: "#5B9BD5", accent6: "#70AD47" };

/** The design colours in use (accent1 … as "#RRGGBB"). */
export function styleTheme(): Readonly<Record<string, string>> { return theme; }

/** The workbook's design colours (dk1, lt1, accent1 … as "#RRGGBB"). */
export function setStyleTheme(t: Record<string, unknown>): void {
  const next = { ...theme };
  for (const [k, v] of Object.entries(t)) if (typeof v === "string" && /^#[0-9a-f]{6}$/i.test(v)) next[k] = v.toUpperCase();
  theme = next;
}

const family = (m: ChartModel): string => {
  const f = kindInfo(m.kind).family;
  if (isPie(m.kind)) return "pie";
  if (f === "line" || f === "scatter" || f === "radar") return "line";
  if (f === "bar" && kindInfo(m.kind).dir === "bar") return "bar";
  return "column";
};

/** The styles the gallery offers for the chart's kind. */
export function chartStyles(m: ChartModel): StyleRef[] {
  return (CHART_STYLES[family(m)] ?? CHART_STYLES.column) as StyleRef[];
}

/** The gallery entry the chart has (its style number), if any. */
export function styleIndex(m: ChartModel): number {
  const list = chartStyles(m);
  const i = list.findIndex((s) => s.id === (m.style ?? list[0]?.id));
  return i;
}

function resolve(c: ColorRef | null | undefined, seriesIndex = 0): { color: string; transparency?: number } | null {
  if (!c) return null;
  let hex = c.rgb ?? (c.s === "phClr" ? theme[`accent${(seriesIndex % 6) + 1}`] : theme[c.s ?? "dk1"]) ?? "#000000";
  if (c.shade !== undefined) hex = lum(hex, c.shade, 0);
  if (c.tint !== undefined) hex = lum(hex, c.tint, 1 - c.tint);
  if (c.lumMod !== undefined || c.lumOff !== undefined) hex = lum(hex, c.lumMod ?? 1, c.lumOff ?? 0);
  const out: { color: string; transparency?: number } = { color: hex.toUpperCase() };
  if (c.alpha !== undefined && c.alpha < 1) out.transparency = Math.round((1 - c.alpha) * 100);
  return out;
}

function fillOf(f: FillRef | undefined, si?: number): Fill | undefined {
  if (!f) return undefined;
  if (f.none) return { type: "none" };
  const c = resolve(f.c, si);
  return c ? { type: "solid", ...c } : undefined;
}

function lineOf(l: LineRef | undefined, si?: number): Line | undefined {
  if (!l) return undefined;
  if (l.none) return { type: "none" };
  const c = resolve(l.c, si);
  const out: Line = c ? { type: "solid", ...c } : { type: "auto" };
  if (l.w !== undefined) out.width = l.w;
  if (l.dash) out.dash = l.dash as Line["dash"];
  return out;
}

function textOf(t: TextRef | undefined): TextFormat | undefined {
  if (!t) return undefined;
  const out: TextFormat = {};
  if (t.size !== undefined) out.size = t.size;
  if (t.bold) out.bold = true;
  const c = resolve(t.c);
  if (c) out.color = c.color;
  return out;
}

const same = (a: unknown, b: unknown) => JSON.stringify(a ?? null) === JSON.stringify(b ?? null);

/** An element's format: only what differs from the default style's same element. */
function diff(el: ElRef | undefined, base: ElRef | undefined, si?: number): ElementFormat | undefined {
  const out: ElementFormat = {};
  if (el?.fill && !same(el.fill, base?.fill)) { const f = fillOf(el.fill, si); if (f) out.fill = f; }
  if (el?.line && !same(el.line, base?.line)) { const l = lineOf(el.line, si); if (l) out.line = l; }
  if (el?.text && !same(el.text, base?.text)) { const t = textOf(el.text); if (t) out.text = t; }
  return Object.keys(out).length ? out : undefined;
}

/** Applies gallery entry `i` (0-based) to the chart: the style's look replaces the element formatting. */
export function applyStyle(m: ChartModel, i: number): void {
  const list = chartStyles(m);
  const st = list[i], base = list[0];
  if (!st || !base) return;
  const f: NonNullable<ChartModel["format"]> = {};
  const put = <K extends keyof typeof f>(k: K, v: (typeof f)[K] | undefined) => { if (v && Object.keys(v).length) f[k] = v; };
  put("chart", diff(st.chart, base.chart));
  put("plot", diff(st.plot, base.plot));
  m.title.show = !st.title.off;
  put("title", diff(st.title, base.title));
  m.legend = (st.legend.pos as ChartModel["legend"]) ?? m.legend;
  put("legend", diff(st.legend, base.legend));
  const pie = isPie(m.kind);
  if (!pie) {
    const bar = family(m) === "bar";
    // Excel's bar charts have the category axis vertical; the recorded c:catAx is the category axis in both.
    const cat = { ...(diff(st.cat, base.cat) ?? {}) } as NonNullable<typeof f.catAxis>;
    if (st.cat.hidden) cat.visible = false;
    put("catAxis", cat);
    const val = { ...(diff(st.val, base.val) ?? {}) } as NonNullable<typeof f.valAxis>;
    if (st.val.hidden) val.visible = false;
    put("valAxis", val);
    m.gridlines = !!st.val.major;
    if (st.val.major) put("majorGrid", diff(st.val.major, base.val.major));
    if (st.val.minor) f.minorGrid = { show: true, ...(diff(st.val.minor, base.val.minor ?? {}) ?? {}) };
    if (st.cat.major) f.catGrid = { show: true, ...(diff(st.cat.major, base.cat.major ?? {}) ?? {}) };
    if (st.cat.minor) f.catMinorGrid = { show: true, ...(diff(st.cat.minor, base.cat.minor ?? {}) ?? {}) };
    void bar;
    const kindDefaultGap = base.gap, kindDefaultOverlap = base.overlap;
    m.gapWidth = st.gap !== undefined && st.gap !== kindDefaultGap && kindInfo(m.kind).family === "bar" ? st.gap : undefined;
    m.overlap = st.overlap !== undefined && st.overlap !== kindDefaultOverlap && kindInfo(m.kind).family === "bar" && kindInfo(m.kind).grouping === "clustered" ? st.overlap : undefined;
  }
  m.dataLabels = !!st.labels;
  if (st.labels) {
    const d: DataLabelOptions = { ...(diff(st.labels, base.labels) ?? {}) };
    if (st.labels.pos && st.labels.pos !== base.labels?.pos) d.pos = st.labels.pos as DataLabelOptions["pos"];
    if (st.labels.showCatName) d.showCat = true;
    if (st.labels.showPercent) d.showPercent = true;
    put("dataLabels", d);
  }
  m.series.forEach((s, si) => {
    delete s.fill; delete s.line; delete s.marker; delete s.color;
    const sr = st.series[si % Math.max(1, st.series.length)], br = base.series[si % Math.max(1, base.series.length)];
    if (!sr || pie) {
      // Pies: the slices' border (Formatvorlage n draws them with white or no lines).
      if (pie && sr?.line && !same(sr.line, br?.line)) { const l = lineOf(sr.line, si); if (l) s.line = l; }
      return;
    }
    if (sr.fill && !same(sr.fill, br?.fill)) {
      // The series' own accent stays automatic (Farben ändern keeps working); a changed shade or transparency is set.
      const own = sr.fill.c?.s === `accent${(si % 6) + 1}` && !sr.fill.c.lumMod && !sr.fill.c.lumOff && !sr.fill.c.alpha && !sr.fill.c.shade && !sr.fill.c.tint;
      if (!own) { const fl = fillOf(sr.fill, si); if (fl) s.fill = fl; }
    }
    if (sr.line && !same(sr.line, br?.line)) {
      const own = sr.line.c?.s === `accent${(si % 6) + 1}` && !sr.line.c.lumMod && !sr.line.c.alpha;
      const l = lineOf(sr.line, si);
      if (l) { if (own) { delete l.color; if (l.type === "solid") l.type = "auto"; } s.line = l; }
    }
    if (sr.marker && !same(sr.marker, br?.marker)) {
      const mk = sr.marker;
      s.marker = {};
      if (mk.symbol) s.marker.symbol = mk.symbol as NonNullable<typeof s.marker.symbol>;
      if (mk.size !== undefined) s.marker.size = mk.size;
      const mf = fillOf(mk.fill, si), ml = lineOf(mk.line, si);
      if (mf) s.marker.fill = mf;
      if (ml) s.marker.line = ml;
    }
  });
  if (Object.keys(f).length) m.format = f; else delete m.format;
  if (i === 0) delete m.style; else m.style = st.id;
}
