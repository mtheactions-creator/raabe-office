// The Format pane of a chart element (Excel: double-click on an element, Strg+1, Format > Auswahl
// formatieren; D-066). Its option sets, icon tabs and sections follow Excel's pane for the element;
// every control changes the chart at once as one undo step. What Excel offers beyond the model
// (gradients, pictures, effects, 3D) is listed in docs/EXCEL-INVENTORY.md as missing.

import {
  isBar, isPie, isScatter, seriesName, type AxisOptions, type ChartModel, type DataLabelOptions, type Dash, type ElementFormat, type ErrorBars,
  type Fill, type Line, type MarkerSymbol, type TextFormat, type Trendline,
} from "@raabe/charts";
import type { PaneControl, PaneSection, PaneSpec, PaneTab } from "@raabe/shell";

/** A chart element as Excel's Diagrammelemente list names them. */
export type ChartElement =
  | { kind: "chart" } | { kind: "plot" } | { kind: "title" } | { kind: "legend" }
  | { kind: "valAxis" } | { kind: "catAxis" } | { kind: "majorGrid" } | { kind: "minorGrid" } | { kind: "catGrid" }
  | { kind: "axisTitle"; axis: "cat" | "val" } | { kind: "series"; index: number } | { kind: "labels" }
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
    case "majorGrid": return `${sc || !bar ? "Vertikal" : "Horizontal"} (Wert) Achse Hauptgitternetz`;
    case "minorGrid": return `${sc || !bar ? "Vertikal" : "Horizontal"} (Wert) Achse Hilfsgitternetz`;
    case "catGrid": return `${bar ? "Vertikal" : "Horizontal"} (Kategorie) Achse Hauptgitternetz`;
    case "axisTitle": return e.axis === "val" ? "Vertikal (Wert) Achse Titel" : "Horizontal (Kategorie) Achse Titel";
    case "series": return `Reihe "${seriesName(m.series[e.index], e.index)}"`;
    case "labels": return "Datenbeschriftungen";
    case "trend": return `Reihe "${seriesName(m.series[e.index], e.index)}" Trendlinie 1`;
    case "errorBars": return `Reihe "${seriesName(m.series[e.index], e.index)}" Y-Fehlerindikatoren`;
  }
}

const PANE_TITLE: Record<ChartElement["kind"], string> = {
  chart: "Diagrammbereich formatieren", plot: "Zeichnungsfläche formatieren", title: "Diagrammtitel formatieren", legend: "Legende formatieren",
  valAxis: "Achse formatieren", catAxis: "Achse formatieren", majorGrid: "Hauptgitternetzlinien formatieren", minorGrid: "Hilfsgitternetzlinien formatieren",
  catGrid: "Hauptgitternetzlinien formatieren", axisTitle: "Achsentitel formatieren", series: "Datenreihen formatieren", labels: "Datenbeschriftungen formatieren",
  trend: "Trendlinie formatieren", errorBars: "Fehlerindikatoren formatieren",
};

const OPTIONS: Record<ChartElement["kind"], string> = {
  chart: "Diagrammoptionen", plot: "Zeichnungsflächenoptionen", title: "Titeloptionen", legend: "Legendenoptionen", valAxis: "Achsenoptionen",
  catAxis: "Achsenoptionen", majorGrid: "Hauptgitternetzlinienoptionen", minorGrid: "Hilfsgitternetzlinienoptionen", catGrid: "Hauptgitternetzlinienoptionen",
  axisTitle: "Titeloptionen", series: "Datenreihenoptionen", labels: "Beschriftungsoptionen", trend: "Trendlinienoptionen", errorBars: "Fehlerindikatoroptionen",
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
    if (m.gridlines) out.push({ kind: "majorGrid" });
    if (m.format?.minorGrid?.show) out.push({ kind: "minorGrid" });
    if (m.format?.catGrid?.show) out.push({ kind: "catGrid" });
  }
  return out;
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

const DASHES: { label: string; value: Dash }[] = [
  { label: "Durchgehend", value: "solid" }, { label: "Rundpunkt", value: "sysDot" }, { label: "Quadratpunkt", value: "sysDash" },
  { label: "Strich", value: "dash" }, { label: "Strich-Punkt", value: "dashDot" }, { label: "Langer Strich", value: "lgDash" },
  { label: "Langer Strich-Punkt", value: "lgDashDot" }, { label: "Langer Strich-Punkt-Punkt", value: "lgDashDotDot" },
];

const TICKS: { label: string; value: string }[] = [{ label: "Keine", value: "none" }, { label: "Innen", value: "in" }, { label: "Außen", value: "out" }, { label: "Kreuz", value: "cross" }];

/** The format object of an element in the model, created on first use. */
function slot(m: ChartModel, e: ChartElement, create: boolean): ElementFormat | undefined {
  const f = create ? (m.format ??= {}) : m.format;
  if (!f) return undefined;
  const get = <K extends keyof NonNullable<ChartModel["format"]>>(k: K) => (create ? ((f[k] as object | undefined) ??= {}) : f[k]) as ElementFormat | undefined;
  switch (e.kind) {
    case "chart": return get("chart");
    case "plot": return get("plot");
    case "title": return get("title");
    case "legend": return get("legend");
    case "valAxis": return get("valAxis");
    case "catAxis": return get("catAxis");
    case "majorGrid": return get("majorGrid");
    case "minorGrid": return get("minorGrid");
    case "catGrid": return get("catGrid");
    case "labels": return get("dataLabels");
    case "axisTitle": {
      const at = create ? (f.axisTitles ??= {}) : f.axisTitles;
      return at ? (create ? (at[e.axis] ??= {}) : at[e.axis]) : undefined;
    }
    case "series": return m.series[e.index] as ElementFormat;
    case "trend": { const t = m.series[e.index]?.trendline; return t ? (t as unknown as ElementFormat) : undefined; }
    case "errorBars": { const t = m.series[e.index]?.errorBars; return t ? (t as unknown as ElementFormat) : undefined; }
  }
}

/** The Format pane of an element. */
export function formatPane(m: ChartModel, e: ChartElement, host: PaneHost): PaneSpec {
  const fmt = () => slot(m, e, false);
  const change = (fn: (f: ElementFormat) => void) => host.edit((c) => { const f = slot(c, e, true); if (f) fn(f); });
  const hasFill = !["valAxis", "catAxis", "majorGrid", "minorGrid", "catGrid", "trend", "errorBars"].includes(e.kind)
    && !(e.kind === "series" && (m.kind === "line" || m.kind === "lineMarkers" || isScatter(m.kind)));
  const lineLabel = hasFill ? "Rahmen" : "Linie";
  const tabs: PaneTab[] = [{ id: "fill", label: "Füllung und Linie", icon: "pane-fill-line", sections: () => fillAndLine(fmt, change, hasFill, lineLabel, m, e, host) }];
  const own = ownTab(m, e, host);
  if (own) tabs.push(own);
  if (e.kind === "chart") tabs.push({ id: "size", label: "Größe und Eigenschaften", icon: "size-position", sections: () => sizeSections(host) });
  const sets = [{ label: OPTIONS[e.kind], tabs }];
  if (["chart", "title", "legend", "valAxis", "catAxis", "labels", "axisTitle"].includes(e.kind)) {
    sets.push({ label: "Textoptionen", tabs: [{ id: "text", label: "Textfüllung und Kontur", icon: "font-color", sections: () => textSections(fmt, change) }] });
  }
  return {
    title: PANE_TITLE[e.kind],
    sets,
    choices: chartElements(m).map((x) => ({ label: elementName(m, x), pick: () => host.select(x) })),
  };
}

function fillAndLine(fmt: () => ElementFormat | undefined, change: (fn: (f: ElementFormat) => void) => void, hasFill: boolean, lineLabel: string,
  m: ChartModel, e: ChartElement, host: PaneHost): PaneSection[] {
  const out: PaneSection[] = [];
  if (hasFill) {
    out.push({
      title: "Füllung", open: true, controls: () => {
        const f = fmt()?.fill;
        const set = (fill: Fill | undefined) => change((x) => { if (fill) x.fill = fill; else delete x.fill; });
        const c: PaneControl[] = [
          { type: "radio", label: "Keine Füllung", checked: () => f?.type === "none", pick: () => set({ type: "none" }) },
          { type: "radio", label: "Einfarbige Füllung", checked: () => f?.type === "solid", pick: () => set({ type: "solid", color: f?.color ?? "#4472C4" }) },
          { type: "radio", label: "Automatisch", checked: () => !f || f.type === "auto", pick: () => set(undefined) },
        ];
        if (e.kind === "series" && (m.kind === "column" || m.kind === "bar" || m.kind.endsWith("Stacked"))) {
          c.push({ type: "check", label: "Invertieren falls negativ", value: () => !!m.series[e.index]?.invertIfNegative, set: (on) => host.edit((cm) => { cm.series[e.index].invertIfNegative = on || undefined; }) });
        }
        if (f?.type === "solid") {
          c.push({ type: "color", label: "Farbe", value: () => f.color ?? null, set: (hex) => set({ ...f, color: hex ?? "#4472C4" }) });
          c.push({ type: "number", label: "Transparenz", unit: " %", min: 0, max: 100, slider: true, value: () => f.transparency ?? 0, set: (v) => set({ ...f, transparency: v || undefined }) });
        }
        return c;
      },
    });
  }
  out.push({
    title: lineLabel, open: !hasFill, controls: () => {
      const l = fmt()?.line;
      const set = (line: Line | undefined) => change((x) => { if (line) x.line = line; else delete x.line; });
      const c: PaneControl[] = [
        { type: "radio", label: "Keine Linie", checked: () => l?.type === "none", pick: () => set({ type: "none" }) },
        { type: "radio", label: "Einfarbige Linie", checked: () => l?.type === "solid", pick: () => set({ type: "solid", color: l?.color ?? "#595959", width: l?.width ?? 0.75 }) },
        { type: "radio", label: "Automatisch", checked: () => !l || l.type === "auto", pick: () => set(undefined) },
      ];
      if (l?.type === "solid") {
        c.push({ type: "color", label: "Farbe", value: () => l.color ?? null, set: (hex) => set({ ...l, color: hex ?? "#595959" }) });
        c.push({ type: "number", label: "Transparenz", unit: " %", min: 0, max: 100, slider: true, value: () => l.transparency ?? 0, set: (v) => set({ ...l, transparency: v || undefined }) });
      }
      if (l?.type !== "none") {
        c.push({ type: "number", label: "Breite", unit: " Pt.", min: 0, max: 1584, step: 0.25, decimals: 2, value: () => l?.width ?? (lineLabel === "Linie" && e.kind === "series" ? 2.25 : 0.75), set: (v) => set({ ...(l ?? { type: "auto" }), width: v }) });
        c.push({ type: "select", label: "Strichtyp", options: DASHES, value: () => l?.dash ?? "solid", set: (v) => set({ ...(l ?? { type: "auto" }), dash: v === "solid" ? undefined : (v as Dash) }) });
      }
      if (e.kind === "series" && (m.kind === "line" || m.kind === "lineMarkers" || isScatter(m.kind))) {
        c.push({ type: "check", label: "Linie glätten", value: () => !!m.series[e.index]?.smooth, set: (on) => host.edit((cm) => { cm.series[e.index].smooth = on || undefined; }) });
      }
      return c;
    },
  });
  if (e.kind === "series" && (m.kind === "line" || m.kind === "lineMarkers" || isScatter(m.kind))) {
    out.push({ title: "Markierung", controls: () => markerControls(m, e.index, host) });
  }
  return out;
}

function markerControls(m: ChartModel, i: number, host: PaneHost): PaneControl[] {
  const mk = m.series[i]?.marker;
  const set = (patch: Partial<NonNullable<typeof mk>> | null) => host.edit((c) => {
    const s = c.series[i];
    if (!patch) delete s.marker;
    else s.marker = { ...s.marker, ...patch };
  });
  const integrated = !!mk?.symbol && mk.symbol !== "none" && mk.symbol !== "auto";
  const symbols: { label: string; value: MarkerSymbol }[] = [
    { label: "Kreis", value: "circle" }, { label: "Quadrat", value: "square" }, { label: "Raute", value: "diamond" }, { label: "Dreieck", value: "triangle" },
    { label: "X", value: "x" }, { label: "Stern", value: "star" }, { label: "Strich", value: "dash" }, { label: "Punkt", value: "dot" }, { label: "Plus", value: "plus" },
  ];
  const c: PaneControl[] = [
    { type: "heading", label: "Markierungsoptionen" },
    { type: "radio", label: "Automatisch", checked: () => !mk?.symbol || mk.symbol === "auto", pick: () => set(null) },
    { type: "radio", label: "Keine", checked: () => mk?.symbol === "none", pick: () => set({ symbol: "none" }) },
    { type: "radio", label: "Integriert", checked: () => integrated, pick: () => set({ symbol: "circle", size: mk?.size ?? 5 }) },
  ];
  if (integrated) {
    c.push({ type: "select", label: "Typ", options: symbols, value: () => mk!.symbol!, set: (v) => set({ symbol: v as MarkerSymbol }) });
    c.push({ type: "number", label: "Größe", min: 2, max: 72, value: () => mk!.size ?? 5, set: (v) => set({ size: Math.round(v) }) });
  }
  return c;
}

function textSections(fmt: () => ElementFormat | undefined, change: (fn: (f: ElementFormat) => void) => void): PaneSection[] {
  const t = () => fmt()?.text;
  const set = (patch: Partial<TextFormat>) => change((x) => {
    const next = { ...x.text, ...patch };
    for (const k of Object.keys(next) as (keyof TextFormat)[]) if (next[k] === undefined) delete next[k];
    if (Object.keys(next).length) x.text = next; else delete x.text;
  });
  return [
    { title: "Textfüllung", open: true, controls: () => [
      { type: "radio", label: "Automatisch", checked: () => !t()?.color, pick: () => set({ color: undefined }) },
      { type: "radio", label: "Einfarbige Füllung", checked: () => !!t()?.color, pick: () => set({ color: t()?.color ?? "#595959" }) },
      ...(t()?.color ? [{ type: "color", label: "Farbe", value: () => t()?.color ?? null, set: (hex: string | null) => set({ color: hex ?? undefined }) } as PaneControl] : []),
    ] },
    { title: "Schrift", open: true, controls: () => [
      { type: "number", label: "Schriftgrad", unit: " Pt.", min: 1, max: 409, step: 1, value: () => t()?.size ?? null, set: (v) => set({ size: v }) },
      { type: "check", label: "Fett", value: () => !!t()?.bold, set: (on) => set({ bold: on || undefined }) },
      { type: "check", label: "Kursiv", value: () => !!t()?.italic, set: (on) => set({ italic: on || undefined }) },
    ] },
  ];
}

function sizeSections(host: PaneHost): PaneSection[] {
  return [{ title: "Größe", open: true, controls: () => [
    { type: "number", label: "Höhe", unit: " cm", min: 0.1, max: 140, step: 0.1, decimals: 2, value: () => host.size().h, set: (v) => host.resize(host.size().w, v) },
    { type: "number", label: "Breite", unit: " cm", min: 0.1, max: 140, step: 0.1, decimals: 2, value: () => host.size().w, set: (v) => host.resize(v, host.size().h) },
  ] }];
}

/** The element's own tab (Achsenoptionen, Datenreihenoptionen …). */
function ownTab(m: ChartModel, e: ChartElement, host: PaneHost): PaneTab | null {
  switch (e.kind) {
    case "legend": return { id: "own", label: "Legendenoptionen", icon: "chart-recommended", sections: () => [{ title: "Legendenoptionen", open: true, controls: () => [
      { type: "heading", label: "Legendenposition" },
      ...([["t", "Oben"], ["b", "Unten"], ["l", "Links"], ["r", "Rechts"]] as const).map(([v, label]) =>
        ({ type: "radio", label, checked: () => m.legend === v, pick: () => host.edit((c) => { c.legend = v; }) }) as PaneControl),
      { type: "check", label: "Legende ohne Überlappung des Diagramms anzeigen", value: () => !m.format?.legend?.overlay, set: (on) => host.edit((c) => { const l = ((c.format ??= {}).legend ??= {}); if (on) delete l.overlay; else l.overlay = true; }) },
    ] }] };
    case "title": return { id: "own", label: "Titeloptionen", icon: "chart-recommended", sections: () => [{ title: "Position", open: true, controls: () => [
      { type: "check", label: "Titel über dem Diagramm einblenden, ohne die Größe zu ändern", value: () => !!m.format?.title?.overlay, set: (on) => host.edit((c) => { const t = ((c.format ??= {}).title ??= {}); if (on) t.overlay = true; else delete t.overlay; }) },
    ] }] };
    case "valAxis": case "catAxis": return { id: "own", label: "Achsenoptionen", icon: "chart-line", sections: () => axisSections(m, e.kind, host) };
    case "series": return { id: "own", label: "Datenreihenoptionen", icon: "chart-column", sections: () => seriesSections(m, e.index, host) };
    case "labels": return { id: "own", label: "Beschriftungsoptionen", icon: "number-format", sections: () => labelSections(m, host) };
    case "trend": return { id: "own", label: "Trendlinienoptionen", icon: "chart-line", sections: () => trendSections(m, e.index, host) };
    case "errorBars": return { id: "own", label: "Fehlerindikatoroptionen", icon: "chart-line", sections: () => errorSections(m, e.index, host) };
    default: return null;
  }
}

function axisSections(m: ChartModel, which: "valAxis" | "catAxis", host: PaneHost): PaneSection[] {
  const o = () => m.format?.[which];
  const set = (patch: Partial<AxisOptions>) => host.edit((c) => {
    const a = ((c.format ??= {})[which] ??= {}) as AxisOptions;
    Object.assign(a, patch);
    for (const k of Object.keys(a) as (keyof AxisOptions)[]) if (a[k] === undefined) delete a[k];
  });
  const isValue = which === "valAxis" || isScatter(m.kind);
  const out: PaneSection[] = [];
  if (isValue) {
    const bound = (label: string, key: "min" | "max" | "majorUnit" | "minorUnit") => ({
      type: "number", label, value: () => o()?.[key] ?? null, set: (v: number) => set({ [key]: v }),
      reset: { label: "Zurücksetzen", isAuto: () => o()?.[key] === undefined, run: () => set({ [key]: undefined }) },
    } as PaneControl);
    const other = which === "valAxis" ? "catAxis" : "valAxis";
    const crosses = () => m.format?.[other]?.crosses;
    const setCross = (v: AxisOptions["crosses"]) => host.edit((c) => { const a = ((c.format ??= {})[other] ??= {}) as AxisOptions; if (v === undefined) delete a.crosses; else a.crosses = v; });
    out.push({ title: "Achsenoptionen", open: true, controls: () => [
      { type: "heading", label: "Begrenzungen" }, bound("Minimum", "min"), bound("Maximum", "max"),
      { type: "heading", label: "Einheiten" }, bound("Haupt", "majorUnit"), bound("Hilfs", "minorUnit"),
      { type: "heading", label: which === "valAxis" ? "Horizontale Achse schneidet" : "Vertikale Achse schneidet" },
      { type: "radio", label: "Automatisch", checked: () => crosses() === undefined || crosses() === "auto", pick: () => setCross(undefined) },
      { type: "radio", label: "Achsenwert", checked: () => typeof crosses() === "number", pick: () => setCross(0) },
      ...(typeof crosses() === "number" ? [{ type: "number", label: "Wert", value: () => crosses() as number, set: (v: number) => setCross(v) } as PaneControl] : []),
      { type: "radio", label: "Maximaler Achsenwert", checked: () => crosses() === "max", pick: () => setCross("max") },
      { type: "check", label: "Logarithmische Skalierung", value: () => !!o()?.logBase, set: (on) => set({ logBase: on ? 10 : undefined }) },
      ...(o()?.logBase ? [{ type: "number", label: "Basis", min: 2, max: 1000, value: () => o()?.logBase ?? 10, set: (v: number) => set({ logBase: v }) } as PaneControl] : []),
      { type: "check", label: "Werte in umgekehrter Reihenfolge", value: () => !!o()?.reverse, set: (on) => set({ reverse: on || undefined }) },
    ] });
  } else {
    out.push({ title: "Achsenoptionen", open: true, controls: () => [
      { type: "heading", label: "Vertikale Achse schneidet" },
      { type: "radio", label: "Automatisch", checked: () => m.format?.valAxis?.crosses !== "max", pick: () => host.edit((c) => { const a = ((c.format ??= {}).valAxis ??= {}); delete a.crosses; }) },
      { type: "radio", label: "Bei größter Kategorie", checked: () => m.format?.valAxis?.crosses === "max", pick: () => host.edit((c) => { ((c.format ??= {}).valAxis ??= {}).crosses = "max"; }) },
      { type: "check", label: "Kategorien in umgekehrter Reihenfolge", value: () => !!o()?.reverse, set: (on) => set({ reverse: on || undefined }) },
    ] });
  }
  out.push({ title: "Teilstriche", controls: () => [
    { type: "select", label: "Hauptteilstrich", options: TICKS, value: () => o()?.majorTick ?? "none", set: (v) => set({ majorTick: v === "none" ? undefined : (v as AxisOptions["majorTick"]) }) },
    { type: "select", label: "Hilfsteilstrich", options: TICKS, value: () => o()?.minorTick ?? "none", set: (v) => set({ minorTick: v === "none" ? undefined : (v as AxisOptions["minorTick"]) }) },
  ] });
  out.push({ title: "Beschriftungen", controls: () => [
    { type: "select", label: "Beschriftungsposition", options: [{ label: "Achsennah", value: "nextTo" }, { label: "Hoch", value: "high" }, { label: "Niedrig", value: "low" }, { label: "Keine", value: "none" }],
      value: () => o()?.labelPos ?? "nextTo", set: (v) => set({ labelPos: v === "nextTo" ? undefined : (v as AxisOptions["labelPos"]) }) },
  ] });
  out.push({ title: "Zahl", controls: () => [
    { type: "text", label: "Formatcode", value: () => o()?.numFmt ?? "Standard", set: (v) => set({ numFmt: v && v !== "Standard" ? v : undefined }) },
    { type: "check", label: "Mit Quelle verknüpfen", value: () => o()?.numFmt === undefined, set: (on) => set({ numFmt: on ? undefined : "General" }) },
  ] });
  return out;
}

function seriesSections(m: ChartModel, i: number, host: PaneHost): PaneSection[] {
  if (isPie(m.kind)) {
    return [{ title: "Datenreihenoptionen", open: true, controls: () => [
      { type: "number", label: "Winkel des ersten Kreissegments", unit: "°", min: 0, max: 360, slider: true, value: () => m.firstSliceAngle ?? 0, set: (v) => host.edit((c) => { c.firstSliceAngle = Math.round(v) || undefined; }) },
      ...(m.kind === "doughnut" ? [{ type: "number", label: "Innenringgröße", unit: " %", min: 10, max: 90, slider: true, value: () => m.holeSize ?? 75, set: (v: number) => host.edit((c) => { c.holeSize = Math.round(v) === 75 ? undefined : Math.round(v); }) } as PaneControl] : []),
    ] }];
  }
  const bars = m.kind === "column" || m.kind === "bar" || m.kind.endsWith("Stacked");
  return [{ title: "Datenreihenoptionen", open: true, controls: () => [
    { type: "heading", label: "Datenreihe zeichnen auf" },
    { type: "radio", label: "Primärachse", checked: () => !m.series[i]?.secondary, pick: () => host.edit((c) => { delete c.series[i].secondary; }) },
    { type: "radio", label: "Sekundärachse", checked: () => !!m.series[i]?.secondary, pick: () => host.edit((c) => { c.series[i].secondary = true; }), disabled: () => true },
    ...(bars ? [
      { type: "number", label: "Reihenachsenüberlappung", unit: " %", min: -100, max: 100, slider: true, value: () => m.overlap ?? (m.kind === "column" ? -27 : m.kind.endsWith("Stacked") ? 100 : 0), set: (v: number) => host.edit((c) => { c.overlap = Math.round(v); }) } as PaneControl,
      { type: "number", label: "Abstandsbreite", unit: " %", min: 0, max: 500, slider: true, value: () => m.gapWidth ?? (m.kind === "column" ? 219 : m.kind === "bar" ? 182 : 150), set: (v: number) => host.edit((c) => { c.gapWidth = Math.round(v); }) } as PaneControl,
    ] : []),
  ] }];
}

function labelSections(m: ChartModel, host: PaneHost): PaneSection[] {
  const o = () => m.format?.dataLabels;
  const set = (patch: Partial<DataLabelOptions>) => host.edit((c) => {
    const d = ((c.format ??= {}).dataLabels ??= {});
    Object.assign(d, patch);
    for (const k of Object.keys(d) as (keyof DataLabelOptions)[]) if (d[k] === undefined) delete d[k];
  });
  const pie = isPie(m.kind), bars = m.kind === "column" || m.kind === "bar" || m.kind.endsWith("Stacked");
  const positions: [DataLabelOptions["pos"], string][] = pie ? [["ctr", "Zentriert"], ["inEnd", "Am inneren Ende"], ["outEnd", "Am äußeren Ende"], ["bestFit", "Optimale Anpassung"]]
    : bars ? [["ctr", "Zentriert"], ["inEnd", "Am inneren Ende"], ["inBase", "Innerhalb der Basis"], ...(m.kind.endsWith("Stacked") ? [] : [["outEnd", "Am äußeren Ende"]] as [DataLabelOptions["pos"], string][])]
    : [["ctr", "Zentriert"], ["l", "Links"], ["r", "Rechts"], ["t", "Oben"], ["b", "Unten"]];
  const defPos = pie ? "bestFit" : bars ? (m.kind.endsWith("Stacked") ? "ctr" : "outEnd") : "r";
  return [
    { title: "Beschriftungsoptionen", open: true, controls: () => [
      { type: "heading", label: "Beschriftung enthält" },
      { type: "check", label: "Reihenname", value: () => !!o()?.showSer, set: (on) => set({ showSer: on || undefined }) },
      { type: "check", label: "Rubrikenname", value: () => !!o()?.showCat, set: (on) => set({ showCat: on || undefined }) },
      { type: "check", label: "Wert", value: () => o()?.showVal !== false, set: (on) => set({ showVal: on ? undefined : false }) },
      ...(pie ? [{ type: "check", label: "Prozentsatz", value: () => !!o()?.showPercent, set: (on: boolean) => set({ showPercent: on || undefined }) } as PaneControl] : []),
      { type: "check", label: "Legendensymbol", value: () => !!o()?.showLegendKey, set: (on) => set({ showLegendKey: on || undefined }) },
      { type: "select", label: "Trennzeichen", options: [{ label: ";", value: "; " }, { label: ",", value: ", " }, { label: ".", value: ". " }, { label: "(Neue Zeile)", value: "\n" }, { label: "(Leerzeichen)", value: " " }],
        value: () => o()?.separator ?? "; ", set: (v) => set({ separator: v === "; " ? undefined : v }) },
      { type: "heading", label: "Beschriftungsposition" },
      ...positions.map(([v, label]) => ({ type: "radio", label, checked: () => (o()?.pos ?? defPos) === v, pick: () => set({ pos: v === defPos ? undefined : v }) }) as PaneControl),
    ] },
    { title: "Zahl", controls: () => [
      { type: "text", label: "Formatcode", value: () => o()?.numFmt ?? "Standard", set: (v) => set({ numFmt: v && v !== "Standard" ? v : undefined }) },
      { type: "check", label: "Mit Quelle verknüpfen", value: () => o()?.numFmt === undefined, set: (on) => set({ numFmt: on ? undefined : "General" }) },
    ] },
  ];
}

function trendSections(m: ChartModel, i: number, host: PaneHost): PaneSection[] {
  const t = () => m.series[i]?.trendline;
  const set = (patch: Partial<Trendline>) => host.edit((c) => {
    const s = c.series[i];
    if (!s.trendline) return;
    Object.assign(s.trendline, patch);
    for (const k of Object.keys(s.trendline) as (keyof Trendline)[]) if (s.trendline[k] === undefined) delete s.trendline[k];
  });
  const type = (v: Trendline["type"], label: string): PaneControl => ({ type: "radio", label, checked: () => t()?.type === v, pick: () => set({ type: v, order: v === "poly" ? t()?.order ?? 2 : undefined, period: v === "movingAvg" ? t()?.period ?? 2 : undefined }) });
  return [{ title: "Trendlinienoptionen", open: true, controls: () => [
    type("exp", "Exponentiell"), type("linear", "Linear"), type("log", "Logarithmisch"), type("poly", "Polynomisch"),
    ...(t()?.type === "poly" ? [{ type: "number", label: "Grad", min: 2, max: 6, value: () => t()?.order ?? 2, set: (v: number) => set({ order: Math.round(v) }) } as PaneControl] : []),
    type("power", "Potenz"), type("movingAvg", "Gleitender Durchschnitt"),
    ...(t()?.type === "movingAvg" ? [{ type: "number", label: "Zeitraum", min: 2, max: 255, value: () => t()?.period ?? 2, set: (v: number) => set({ period: Math.round(v) }) } as PaneControl] : []),
    { type: "heading", label: "Name der Trendlinie" },
    { type: "radio", label: "Automatisch", checked: () => !t()?.name, pick: () => set({ name: undefined }) },
    { type: "radio", label: "Benutzerdefiniert", checked: () => !!t()?.name, pick: () => set({ name: t()?.name ?? "Trendlinie" }) },
    ...(t()?.name ? [{ type: "text", label: "Name", value: () => t()?.name ?? "", set: (v: string) => set({ name: v || undefined }) } as PaneControl] : []),
    { type: "heading", label: "Prognose" },
    { type: "number", label: "Vorwärts", unit: " Perioden", min: 0, value: () => t()?.forward ?? 0, set: (v) => set({ forward: v || undefined }) },
    { type: "number", label: "Rückwärts", unit: " Perioden", min: 0, value: () => t()?.backward ?? 0, set: (v) => set({ backward: v || undefined }) },
    { type: "check", label: "Formel im Diagramm anzeigen", value: () => !!t()?.showEq, set: (on) => set({ showEq: on }) },
    { type: "check", label: "Bestimmtheitsmaß im Diagramm darstellen", value: () => !!t()?.showR2, set: (on) => set({ showR2: on }) },
  ] }];
}

function errorSections(m: ChartModel, i: number, host: PaneHost): PaneSection[] {
  const b = () => m.series[i]?.errorBars;
  const set = (patch: Partial<ErrorBars>) => host.edit((c) => { const s = c.series[i]; if (s.errorBars) Object.assign(s.errorBars, patch); });
  const amount = (v: ErrorBars["type"], label: string, unit?: string, def?: number): PaneControl[] => [
    { type: "radio", label, checked: () => b()?.type === v, pick: () => set({ type: v, value: def }) },
    ...(b()?.type === v && def !== undefined ? [{ type: "number", label: "Wert", unit, min: 0, value: () => b()?.value ?? def, set: (x: number) => set({ value: x }) } as PaneControl] : []),
  ];
  return [{ title: "Vertikaler Fehlerindikator", open: true, controls: () => [
    { type: "heading", label: "Richtung" },
    ...([["both", "Beide"], ["minus", "Minus"], ["plus", "Plus"]] as const).map(([v, label]) => ({ type: "radio", label, checked: () => b()?.dir === v, pick: () => set({ dir: v }) }) as PaneControl),
    { type: "heading", label: "Endformat" },
    { type: "radio", label: "Ohne Abschluss", checked: () => !!b()?.noEndCap, pick: () => set({ noEndCap: true }) },
    { type: "radio", label: "Abschluss", checked: () => !b()?.noEndCap, pick: () => set({ noEndCap: undefined }) },
    { type: "heading", label: "Fehlerbetrag" },
    ...amount("fixedVal", "Fester Wert", undefined, 1), ...amount("percentage", "Prozentsatz", " %", 5), ...amount("stdDev", "Standardabweichung(en)", undefined, 1),
    ...amount("stdErr", "Standardfehler"),
  ] }];
}
