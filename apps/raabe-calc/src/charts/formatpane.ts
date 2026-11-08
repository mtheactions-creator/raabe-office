// The Format pane of a chart element (Excel: double-click on an element, Strg+1, Format > Auswahl
// formatieren; D-066, D-111). Its option sets, icon tabs and sections follow Excel's pane for the
// element (recorded in results/2026-11-09-excel-chart-pane.json); every control changes the chart at
// once as one undo step. The shared tabs (Füllung und Linie, Effekte) are in panefill.ts, the
// element's own options in paneoptions.ts, text and size in panetext.ts.

import { isBar, isPie, isScatter, kindInfo, seriesColor, seriesName, type ChartModel, type DataPoint, type Effects, type ElementFormat, type Fill, type Line } from "@raabe/charts";
import type { PaneSection, PaneSpec, PaneTab } from "@raabe/shell";
import { effectSections, fillSection, lineSection, type FmtTarget } from "./panefill";
import { axisSections, errorSections, labelSections, pointSections, seriesSections, trendSections } from "./paneoptions";
import { sizeSections, textSections } from "./panetext";

/** A chart element as Excel's Diagrammelemente list names them. */
export type ChartElement =
  | { kind: "chart" } | { kind: "plot" } | { kind: "title" } | { kind: "legend" }
  | { kind: "valAxis" } | { kind: "catAxis" } | { kind: "valAxis2" } | { kind: "majorGrid" } | { kind: "minorGrid" } | { kind: "catGrid" }
  | { kind: "axisTitle"; axis: "cat" | "val" } | { kind: "series"; index: number } | { kind: "point"; index: number; point: number } | { kind: "labels" }
  | { kind: "trend"; index: number } | { kind: "errorBars"; index: number };

/** The element's name in the Diagrammelemente list and the pane's title. */
export function elementName(m: ChartModel, e: ChartElement): string {
  const bar = isBar(m.kind), sc = isScatter(m.kind);
  switch (e.kind) {
    case "chart": return "Diagrammbereich";
    case "plot": return "Zeichnungsfläche";
    case "title": return "Diagrammtitel";
    case "legend": return "Legende";
    case "valAxis": return sc ? "Vertikal (Wert) Achse" : bar ? "Horizontal (Wert) Achse" : "Vertikal (Wert) Achse";
    case "catAxis": return sc ? "Horizontal (Wert) Achse" : bar ? "Vertikal (Kategorie) Achse" : "Horizontal (Kategorie) Achse";
    case "valAxis2": return bar ? "Sekundär Horizontal (Wert) Achse" : "Sekundär Vertikal (Wert) Achse";
    case "majorGrid": return `${sc || !bar ? "Vertikal" : "Horizontal"} (Wert) Achse Hauptgitternetz`;
    case "minorGrid": return `${sc || !bar ? "Vertikal" : "Horizontal"} (Wert) Achse Hilfsgitternetz`;
    case "catGrid": return `${bar ? "Vertikal" : "Horizontal"} (Kategorie) Achse Hauptgitternetz`;
    case "axisTitle": return e.axis === "val" ? "Vertikal (Wert) Achse Titel" : "Horizontal (Kategorie) Achse Titel";
    case "series": return `Reihe "${seriesName(m.series[e.index], e.index)}"`;
    case "point": {
      const s = m.series[e.index];
      const c = s?.cat?.cache[e.point];
      return `Reihe "${seriesName(s, e.index)}" Punkt "${c === null || c === undefined ? e.point + 1 : c}"`;
    }
    case "labels": return "Datenbeschriftungen";
    case "trend": return `Reihe "${seriesName(m.series[e.index], e.index)}" Trendlinie 1`;
    case "errorBars": return `Reihe "${seriesName(m.series[e.index], e.index)}" Y-Fehlerindikatoren`;
  }
}

const PANE_TITLE: Record<ChartElement["kind"], string> = {
  chart: "Diagrammbereich formatieren", plot: "Zeichnungsfläche formatieren", title: "Diagrammtitel formatieren", legend: "Legende formatieren",
  valAxis: "Achse formatieren", catAxis: "Achse formatieren", valAxis2: "Achse formatieren", majorGrid: "Hauptgitternetzlinien formatieren", minorGrid: "Hilfsgitternetzlinien formatieren",
  catGrid: "Hauptgitternetzlinien formatieren", axisTitle: "Achsentitel formatieren", series: "Datenreihen formatieren", point: "Datenpunkt formatieren",
  labels: "Datenbeschriftungen formatieren", trend: "Trendlinie formatieren", errorBars: "Fehlerindikatoren formatieren",
};

const OPTIONS: Record<ChartElement["kind"], string> = {
  chart: "Diagrammoptionen", plot: "Zeichnungsflächenoptionen", title: "Titeloptionen", legend: "Legendenoptionen", valAxis: "Achsenoptionen",
  catAxis: "Achsenoptionen", valAxis2: "Achsenoptionen", majorGrid: "Hauptgitternetzlinienoptionen", minorGrid: "Hilfsgitternetzlinienoptionen", catGrid: "Hauptgitternetzlinienoptionen",
  axisTitle: "Titeloptionen", series: "Datenreihenoptionen", point: "Datenpunktoptionen", labels: "Beschriftungsoptionen", trend: "Trendlinienoptionen", errorBars: "Fehlerindikatoroptionen",
};

/** Every element a chart has, in the order of Excel's Diagrammelemente list. */
export function chartElements(m: ChartModel): ChartElement[] {
  const pie = isPie(m.kind);
  const out: ChartElement[] = [{ kind: "chart" }];
  if (m.title.show) out.push({ kind: "title" });
  if (!pie) {
    out.push({ kind: "catAxis" });
    if (m.axisTitles.cat !== undefined) out.push({ kind: "axisTitle", axis: "cat" });
  }
  if (m.legend !== "none") out.push({ kind: "legend" });
  out.push({ kind: "plot" });
  m.series.forEach((s, i) => {
    out.push({ kind: "series", index: i });
    if (s.trendline) out.push({ kind: "trend", index: i });
    if (s.errorBars) out.push({ kind: "errorBars", index: i });
  });
  if (m.dataLabels) out.push({ kind: "labels" });
  if (!pie) {
    out.push({ kind: "valAxis" });
    if (m.axisTitles.val !== undefined) out.push({ kind: "axisTitle", axis: "val" });
    if (m.series.some((s) => s.secondary)) out.push({ kind: "valAxis2" });
    if (m.gridlines) out.push({ kind: "majorGrid" });
    if (m.format?.minorGrid?.show) out.push({ kind: "minorGrid" });
    if (m.format?.catGrid?.show) out.push({ kind: "catGrid" });
  }
  return out;
}

/**
 * The element a click on a drawn part means (the drawing's data-part marks): a second click on a
 * selected series' point selects that point (Datenpunkt formatieren), as in Excel.
 */
export function elementOfPart(part: string | undefined, current?: ChartElement): ChartElement | null {
  if (!part) return null;
  const [kind, a, b] = part.split(":");
  switch (kind) {
    case "chart": return { kind: "chart" };
    case "plot": return { kind: "plot" };
    case "title": return { kind: "title" };
    case "legend": return { kind: "legend" };
    case "axis": return a === "val" ? { kind: "valAxis" } : a === "val2" ? { kind: "valAxis2" } : { kind: "catAxis" };
    case "axes": return { kind: "catAxis" };
    case "grid": return { kind: "majorGrid" };
    case "minorGrid": return { kind: "minorGrid" };
    case "catGrid": return { kind: "catGrid" };
    case "axisTitle": return { kind: "axisTitle", axis: a === "val" ? "val" : "cat" };
    case "series": {
      const index = Number(a);
      const point = b === undefined ? undefined : Number(b);
      if (point !== undefined && current && (current.kind === "series" || current.kind === "point") && current.index === index) return { kind: "point", index, point };
      return { kind: "series", index };
    }
    case "labels": return { kind: "labels" };
    case "trend": case "trendLabel": return { kind: "trend", index: Number(a) };
    case "errbars": return { kind: "errorBars", index: Number(a) };
    default: return null;
  }
}

/** The SVG parts that make up an element (for its selection frame). */
export function partsSelector(e: ChartElement): string {
  switch (e.kind) {
    case "chart": return "";
    case "valAxis": return '[data-part="axis:val"]';
    case "valAxis2": return '[data-part="axis:val2"]';
    case "catAxis": return '[data-part="axis:cat"],[data-part="axes"]';
    case "majorGrid": return '[data-part="grid"]';
    case "series": return `[data-part^="series:${e.index}"]`;
    case "point": return `[data-part="series:${e.index}:${e.point}"]`;
    case "labels": return '[data-part^="labels:"]';
    case "trend": return `[data-part="trend:${e.index}"]`;
    case "errorBars": return `[data-part="errbars:${e.index}"]`;
    case "axisTitle": return `[data-part="axisTitle:${e.axis}"]`;
    default: return `[data-part="${e.kind}"]`;
  }
}

export interface PaneHost {
  /** Changes the chart as one undo step. */
  edit: (fn: (m: ChartModel) => void) => void;
  /** Selects another element (the pane follows it). */
  select: (e: ChartElement) => void;
  /** The chart's size in cm and setting it (Größe und Eigenschaften). */
  size: () => { w: number; h: number };
  resize: (w: number, h: number) => void;
}

type Fmt = { fill?: Fill; line?: Line; effects?: Effects } & Partial<ElementFormat>;

/** The format object of an element in the model, created on first use. */
export function slot(m: ChartModel, e: ChartElement, create: boolean): Fmt | undefined {
  const f = create ? (m.format ??= {}) : m.format;
  const get = <K extends keyof NonNullable<ChartModel["format"]>>(k: K): Fmt | undefined => {
    if (!f) return undefined;
    return (create ? ((f[k] as object | undefined) ??= {}) : f[k]) as Fmt | undefined;
  };
  switch (e.kind) {
    case "chart": return get("chart");
    case "plot": return get("plot");
    case "title": return get("title");
    case "legend": return get("legend");
    case "valAxis": return get("valAxis");
    case "catAxis": return get("catAxis");
    case "valAxis2": return get("valAxis2");
    case "majorGrid": return get("majorGrid");
    case "minorGrid": return get("minorGrid");
    case "catGrid": return get("catGrid");
    case "labels": return get("dataLabels");
    case "axisTitle": {
      if (!f) return undefined;
      const at = create ? (f.axisTitles ??= {}) : f.axisTitles;
      return at ? (create ? (at[e.axis] ??= {}) : at[e.axis]) : undefined;
    }
    case "series": return m.series[e.index] as Fmt;
    case "point": {
      const s = m.series[e.index];
      if (!s) return undefined;
      let p = s.points?.find((x) => x.idx === e.point);
      if (!p && create) { p = { idx: e.point }; (s.points ??= []).push(p); s.points.sort((a, b) => a.idx - b.idx); }
      return p as Fmt | undefined;
    }
    case "trend": { const t = m.series[e.index]?.trendline; return t ? (t as unknown as Fmt) : undefined; }
    case "errorBars": { const t = m.series[e.index]?.errorBars; return t ? (t as unknown as Fmt) : undefined; }
  }
}

/** Drops a data point's entry when it holds nothing any more. */
function tidyPoints(m: ChartModel): void {
  for (const s of m.series) {
    if (!s.points) continue;
    s.points = s.points.filter((p: DataPoint) => Object.keys(p).some((k) => k !== "idx" && p[k as keyof DataPoint] !== undefined));
    if (!s.points.length) delete s.points;
  }
}

/** The colour an element shows automatically (what a new Einfarbige Füllung / Linie starts with). */
function autoColor(m: ChartModel, e: ChartElement): string {
  if (e.kind === "series" || e.kind === "point" || e.kind === "trend") {
    const s = m.series[e.index];
    return s?.color ?? seriesColor(e.kind === "point" && isPie(m.kind) ? e.point : e.index);
  }
  return "#4472C4";
}

/** The Format pane of an element. */
export function formatPane(m: ChartModel, e: ChartElement, host: PaneHost): PaneSpec {
  const target: FmtTarget = {
    key: JSON.stringify(e),
    get: () => slot(m, e, false),
    change: (fn) => host.edit((c) => { const f = slot(c, e, true); if (f) fn(f); tidyPoints(c); }),
  };
  const series = e.kind === "series" || e.kind === "point" ? m.series[e.index] : undefined;
  const sk = series?.kind ?? m.kind;
  const fam = kindInfo(sk).family;
  const lineKind = (fam === "line" && !kindInfo(sk).d3) || fam === "scatter" || (fam === "radar" && sk !== "radarFilled") || fam === "stock";
  const isLine = ["valAxis", "catAxis", "valAxis2", "majorGrid", "minorGrid", "catGrid", "trend", "errorBars"].includes(e.kind) || ((e.kind === "series" || e.kind === "point") && lineKind);
  const auto = autoColor(m, e);
  const tabs: PaneTab[] = [{ id: "fill", label: "Füllung und Linie", icon: "pane-fill-line", sections: () => fillAndLine(m, e, host, target, isLine, lineKind, auto) }];
  // Effekte: every element but the gridlines' Leuchten-less lines has Schatten, Leuchten, Weiche Kanten; shapes have 3D-Format.
  const shape = !isLine && e.kind !== "labels" && e.kind !== "title" && e.kind !== "legend" && e.kind !== "axisTitle";
  tabs.push({ id: "effects", label: "Effekte", icon: "pane-effects", sections: () => effectSections(target, { autoColor: auto, soft: !isLine || e.kind === "series" || e.kind === "point", format3d: shape || e.kind === "chart" || e.kind === "plot" }) });
  if (e.kind === "chart") tabs.push({ id: "size", label: "Größe und Eigenschaften", icon: "size-position", sections: () => sizeSections(m, host) });
  const own = ownTab(m, e, host);
  if (own) tabs.push(own);
  const sets = [{ label: OPTIONS[e.kind], tabs }];
  if (["chart", "title", "legend", "valAxis", "catAxis", "valAxis2", "labels", "axisTitle"].includes(e.kind)) {
    sets.push({ label: "Textoptionen", tabs: textSections(target, e.kind) });
  }
  return {
    title: PANE_TITLE[e.kind],
    sets,
    choices: chartElements(m).map((x) => ({ label: elementName(m, x), pick: () => host.select(x) })),
  };
}

function fillAndLine(m: ChartModel, e: ChartElement, host: PaneHost, t: FmtTarget, isLine: boolean, lineKind: boolean, auto: string): PaneSection[] {
  const out: PaneSection[] = [];
  const s = e.kind === "series" || e.kind === "point" ? m.series[e.index] : undefined;
  const fam = kindInfo(s?.kind ?? m.kind).family;
  if (!isLine) {
    const extra = [];
    if (e.kind === "series" && fam === "bar") {
      extra.push({ type: "check" as const, label: "Invertieren falls negativ", value: () => !!m.series[e.index]?.invertIfNegative, set: (on: boolean) => host.edit((cm) => { cm.series[e.index].invertIfNegative = on || undefined; }) });
    }
    if (e.kind === "series" && !isPie(m.kind) && m.series.length === 1 && (fam === "bar" || fam === "area")) {
      extra.push({ type: "check" as const, label: "Punkte nach Farbe unterscheiden", value: () => !!m.varyColors, set: (on: boolean) => host.edit((cm) => { cm.varyColors = on || undefined; }) });
    }
    if (e.kind === "series" && isPie(m.kind)) {
      extra.push({ type: "check" as const, label: "Punkte nach Farbe unterscheiden", value: () => !s?.fill || s.fill.type === "auto", set: (on: boolean) => host.edit((cm) => { if (on) delete cm.series[e.index].fill; else cm.series[e.index].fill = { type: "solid", color: auto }; }) });
    }
    out.push(fillSection(t, auto, extra));
  }
  const lineExtra = [];
  if (e.kind === "chart") lineExtra.push({ type: "check" as const, label: "Abgerundete Ecken", value: () => !!m.format?.chart?.roundedCorners, set: (on: boolean) => host.edit((c) => { const ca = ((c.format ??= {}).chart ??= {}); if (on) ca.roundedCorners = true; else delete ca.roundedCorners; }) });
  if (e.kind === "series" && lineKind && fam !== "stock") lineExtra.push({ type: "check" as const, label: "Linie glätten", value: () => !!m.series[e.index]?.smooth, set: (on: boolean) => host.edit((cm) => { cm.series[e.index].smooth = on || undefined; }) });
  const autoWidth = e.kind === "series" || e.kind === "point" ? (fam === "line" || fam === "radar" ? 2.25 : lineKind ? 1.5 : 0.75) : e.kind === "trend" ? 1.5 : 0.75;
  const lineColor = isLine && (e.kind === "series" || e.kind === "point" || e.kind === "trend") ? auto : isLine ? "#BFBFBF" : "#595959";
  out.push(lineSection(t, {
    title: isLine ? "Linie" : "Rahmen", autoColor: lineColor, autoWidth, open: isLine,
    arrows: isLine && !["majorGrid", "minorGrid", "catGrid"].includes(e.kind), extra: lineExtra,
  }));
  if ((e.kind === "series" || e.kind === "point") && lineKind) out.push(...markerSections(m, e, host));
  return out;
}

/** Markierung of line and scatter series (and of one point): Markierungsoptionen, its Füllung and Rahmen. */
function markerSections(m: ChartModel, e: Extract<ChartElement, { kind: "series" } | { kind: "point" }>, host: PaneHost): PaneSection[] {
  const i = e.index;
  const get = () => (e.kind === "point" ? m.series[i]?.points?.find((p) => p.idx === e.point)?.marker ?? m.series[i]?.marker : m.series[i]?.marker);
  const edit = (fn: (mk: NonNullable<DataPoint["marker"]>) => void, clear = false) => host.edit((c) => {
    const s = c.series[i];
    if (e.kind === "point") {
      let p = s.points?.find((x) => x.idx === e.point);
      if (!p) { p = { idx: e.point }; (s.points ??= []).push(p); }
      if (clear) delete p.marker; else fn((p.marker ??= { ...s.marker }));
      tidyPoints(c);
    } else if (clear) delete s.marker;
    else fn((s.marker ??= {}));
  });
  const symbols = [
    { label: "Kreis", value: "circle" }, { label: "Quadrat", value: "square" }, { label: "Raute", value: "diamond" }, { label: "Dreieck", value: "triangle" },
    { label: "X", value: "x" }, { label: "Stern", value: "star" }, { label: "Strich", value: "dash" }, { label: "Punkt", value: "dot" }, { label: "Plus", value: "plus" },
  ];
  const mt: FmtTarget = {
    key: JSON.stringify(e) + "/marker",
    get: () => get(),
    change: (fn) => edit((mk) => fn(mk)),
  };
  const auto = m.series[i]?.color ?? seriesColor(i);
  return [
    { title: "Markierungsoptionen", controls: () => {
      const mk = get();
      const integrated = !!mk?.symbol && mk.symbol !== "none" && mk.symbol !== "auto";
      return [
        { type: "radio", label: "Automatisch", checked: () => !mk?.symbol || mk.symbol === "auto", pick: () => edit(() => undefined, true) },
        { type: "radio", label: "Keine", checked: () => mk?.symbol === "none", pick: () => edit((x) => { x.symbol = "none"; }) },
        { type: "radio", label: "Integriert", checked: () => integrated, pick: () => edit((x) => { x.symbol = "circle"; x.size = x.size ?? 5; }) },
        { type: "select", label: "Typ", options: symbols, value: () => (integrated ? mk!.symbol! : "circle"), set: (v) => edit((x) => { x.symbol = v as typeof x.symbol; }), disabled: () => !integrated },
        { type: "number", label: "Größe", min: 2, max: 72, value: () => mk?.size ?? 5, set: (v) => edit((x) => { x.size = Math.round(v); }), disabled: () => !integrated },
      ];
    } },
    { ...fillSection(mt, auto, [], false), title: "Markierungsfüllung" },
    { ...lineSection(mt, { title: "Markierungsrahmen", autoColor: auto, autoWidth: 0.75 }) },
  ];
}

/** The element's own tab (Achsenoptionen, Datenreihenoptionen …). */
function ownTab(m: ChartModel, e: ChartElement, host: PaneHost): PaneTab | null {
  switch (e.kind) {
    case "legend": return { id: "own", label: "Legendenoptionen", icon: "chart-recommended", sections: () => [{ title: "Legendenoptionen", open: true, controls: () => [
      { type: "heading", label: "Legendenposition" },
      ...([["t", "Oben"], ["b", "Unten"], ["l", "Links"], ["r", "Rechts"], ["tr", "Oben rechts"]] as const).map(([v, label]) =>
        ({ type: "radio" as const, label, checked: () => m.legend === v, pick: () => host.edit((c) => { c.legend = v; }) })),
      { type: "check", label: "Legende ohne Überlappung des Diagramms anzeigen", value: () => !m.format?.legend?.overlay, set: (on) => host.edit((c) => { const l = ((c.format ??= {}).legend ??= {}); if (on) delete l.overlay; else l.overlay = true; }) },
    ] }] };
    case "title": return { id: "own", label: "Titeloptionen", icon: "chart-recommended", sections: () => [{ title: "Position", open: true, controls: () => [
      { type: "check", label: "Titel über dem Diagramm einblenden, ohne die Größe zu ändern", value: () => !!m.format?.title?.overlay, set: (on) => host.edit((c) => { const t = ((c.format ??= {}).title ??= {}); if (on) t.overlay = true; else delete t.overlay; }) },
    ] }] };
    case "valAxis": case "catAxis": case "valAxis2": return { id: "own", label: "Achsenoptionen", icon: "chart-line", sections: () => axisSections(m, e.kind, host) };
    case "series": return { id: "own", label: "Datenreihenoptionen", icon: "chart-column", sections: () => seriesSections(m, e.index, host) };
    case "point": return { id: "own", label: "Datenpunktoptionen", icon: "chart-column", sections: () => pointSections(m, e.index, e.point, host) };
    case "labels": return { id: "own", label: "Beschriftungsoptionen", icon: "number-format", sections: () => labelSections(m, host) };
    case "trend": return { id: "own", label: "Trendlinienoptionen", icon: "chart-line", sections: () => trendSections(m, e.index, host) };
    case "errorBars": return { id: "own", label: "Fehlerindikatoroptionen", icon: "chart-line", sections: () => errorSections(m, e.index, host) };
    default: return null;
  }
}
