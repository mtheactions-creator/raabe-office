// The element's own options in the chart Format pane (D-066, D-111): Achsenoptionen (value, category,
// date and secondary axes), Datenreihen- and Datenpunktoptionen per chart type, Beschriftungsoptionen,
// Trendlinien- and Fehlerindikatoroptionen. Labels as in Excel 365 (results/2026-11-09-excel-chart-pane.json).

import { isPie, isScatter, kindInfo, barDefaults, type AxisOptions, type ChartModel, type DataLabelOptions, type ErrorBars, type Trendline } from "@raabe/charts";
import type { PaneControl, PaneSection } from "@raabe/shell";
import type { PaneHost } from "./formatpane";

const TICKS = [{ label: "Ohne", value: "none" }, { label: "Innen", value: "in" }, { label: "Außen", value: "out" }, { label: "Kreuz", value: "cross" }];

const DISPLAY_UNITS = [
  { label: "Ohne", value: "" }, { label: "Hunderte", value: "hundreds" }, { label: "Tausende", value: "thousands" }, { label: "10000", value: "tenThousands" },
  { label: "100000", value: "hundredThousands" }, { label: "Millionen", value: "millions" }, { label: "10000000", value: "tenMillions" },
  { label: "100000000", value: "hundredMillions" }, { label: "Milliarden", value: "billions" }, { label: "Billionen", value: "trillions" },
];

/** Zahl > Kategorie: the format code each category writes (file form), with Dezimalstellen where it has them. */
const NUMBER_CATEGORIES: { label: string; value: string; code: (d: number) => string; decimals?: boolean }[] = [
  { label: "Standard", value: "General", code: () => "General" },
  { label: "Zahl", value: "number", code: (d) => (d ? `0.${"0".repeat(d)}` : "0"), decimals: true },
  { label: "Währung", value: "currency", code: (d) => `#,##0${d ? "." + "0".repeat(d) : ""} "€"`, decimals: true },
  { label: "Buchhaltung", value: "accounting", code: (d) => `_-* #,##0${d ? "." + "0".repeat(d) : ""} "€"_-;-* #,##0${d ? "." + "0".repeat(d) : ""} "€"_-;_-* "-"?? "€"_-;_-@_-`, decimals: true },
  { label: "Datum", value: "date", code: () => "dd.mm.yyyy" },
  { label: "Zeit", value: "time", code: () => "hh:mm:ss" },
  { label: "Prozentsatz", value: "percent", code: (d) => `0${d ? "." + "0".repeat(d) : ""}%`, decimals: true },
  { label: "Bruch", value: "fraction", code: () => "# ?/?" },
  { label: "Wissenschaftlich", value: "scientific", code: (d) => `0${d ? "." + "0".repeat(d) : ""}E+00`, decimals: true },
  { label: "Text", value: "text", code: () => "@" },
  { label: "Sonderformat", value: "special", code: () => "00000" },
  { label: "Benutzerdefiniert", value: "custom", code: () => "" },
];

/** The category of a format code and its decimals (for the Zahl section). */
function categoryOf(code: string | undefined): { cat: string; decimals: number } {
  if (!code || code === "General") return { cat: "General", decimals: 0 };
  const dec = /0\.(0+)/.exec(code)?.[1].length ?? 0;
  for (const c of NUMBER_CATEGORIES) if (c.value !== "custom" && c.value !== "General" && c.code(dec) === code) return { cat: c.value, decimals: dec };
  return { cat: "custom", decimals: dec };
}

/** The Zahl section of axes and data labels: Kategorie, Dezimalstellen, Formatcode, Mit Quelle verknüpfen. */
export function numberSection(get: () => string | undefined, set: (code: string | undefined) => void): PaneSection {
  return { title: "Zahl", controls: () => {
    const code = get();
    const { cat, decimals } = categoryOf(code);
    const info = NUMBER_CATEGORIES.find((c) => c.value === cat)!;
    const c: PaneControl[] = [
      { type: "select", label: "Rubrik", options: NUMBER_CATEGORIES.map(({ label, value }) => ({ label, value })), value: () => cat,
        set: (v) => { const k = NUMBER_CATEGORIES.find((x) => x.value === v)!; if (v === "custom") return; set(k.code(k.decimals ? 2 : 0)); } },
    ];
    if (info.decimals) c.push({ type: "number", label: "Dezimalstellen", min: 0, max: 30, value: () => decimals, set: (v) => set(info.code(Math.round(v))) });
    c.push({ type: "text", label: "Formatcode", value: () => code ?? "Standard", set: (v) => set(v && v !== "Standard" ? v : undefined) });
    c.push({ type: "check", label: "Mit Quelle verknüpft", value: () => code === undefined, set: (on) => set(on ? undefined : "General") });
    return c;
  } };
}

export function axisSections(m: ChartModel, which: "valAxis" | "catAxis" | "valAxis2", host: PaneHost): PaneSection[] {
  const o = () => m.format?.[which];
  const set = (patch: Partial<AxisOptions>) => host.edit((c) => {
    const a = ((c.format ??= {})[which] ??= {}) as AxisOptions;
    Object.assign(a, patch);
    for (const k of Object.keys(a) as (keyof AxisOptions)[]) if (a[k] === undefined) delete a[k];
  });
  const scatter = isScatter(m.kind);
  const isValue = which !== "catAxis" || scatter;
  const dateAxis = which === "catAxis" && !scatter && o()?.axisType === "date";
  const out: PaneSection[] = [];
  const bound = (label: string, key: "min" | "max" | "majorUnit" | "minorUnit") => ({
    type: "number", label, value: () => o()?.[key] ?? null, set: (v: number) => set({ [key]: v }),
    reset: { label: o()?.[key] === undefined ? "Automatisch" : "Zurücksetzen", isAuto: () => o()?.[key] === undefined, run: () => set({ [key]: undefined }) },
  } as PaneControl);
  // "… Achse schneidet": the other axis' crossing is this axis' `crosses` (D-111).
  const crosses = () => o()?.crosses;
  const setCross = (v: AxisOptions["crosses"]) => set({ crosses: v === "auto" ? undefined : v });
  if (isValue) {
    const bar = kindInfo(m.kind).dir === "bar";
    const otherName = which === "valAxis2" ? "Horizontale Achse schneidet" : scatter && which === "catAxis" ? "Vertikale Achse schneidet" : bar ? "Vertikale Achse schneidet" : "Horizontale Achse schneidet";
    out.push({ title: "Achsenoptionen", open: true, controls: () => [
      { type: "heading", label: "Grenzen" }, bound("Minimum", "min"), bound("Maximum", "max"),
      { type: "heading", label: "Einheiten" }, bound("Hauptstriche", "majorUnit"), bound("Teilstriche", "minorUnit"),
      { type: "heading", label: otherName },
      { type: "radio", label: "Automatisch", checked: () => crosses() === undefined || crosses() === "auto", pick: () => setCross(undefined) },
      { type: "radio", label: "Achsenwert", checked: () => typeof crosses() === "number", pick: () => setCross(0) },
      { type: "number", label: "Achsenwert", value: () => (typeof crosses() === "number" ? (crosses() as number) : 0), set: (v: number) => setCross(v), disabled: () => typeof crosses() !== "number" },
      { type: "radio", label: "Maximaler Achsenwert", checked: () => crosses() === "max", pick: () => setCross("max") },
      { type: "select", label: "Anzeigeeinheiten", options: DISPLAY_UNITS, value: () => (typeof o()?.dispUnits === "string" ? (o()!.dispUnits as string) : ""),
        set: (v) => set({ dispUnits: (v || undefined) as AxisOptions["dispUnits"], dispUnitsLabel: undefined }) },
      { type: "check", label: "Beschriftung der Anzeigeeinheiten im Diagramm anzeigen", value: () => o()?.dispUnits !== undefined && o()?.dispUnitsLabel !== false,
        set: (on) => set({ dispUnitsLabel: on ? undefined : false }), disabled: () => o()?.dispUnits === undefined },
      { type: "check", label: "Logarithmische Skalierung", value: () => !!o()?.logBase, set: (on) => set({ logBase: on ? 10 : undefined }) },
      { type: "number", label: "Basis", min: 2, max: 1000, value: () => o()?.logBase ?? 10, set: (v: number) => set({ logBase: v }), disabled: () => !o()?.logBase },
      { type: "check", label: "Werte in umgekehrter Reihenfolge", value: () => !!o()?.reverse, set: (on) => set({ reverse: on || undefined }) },
    ] });
  } else {
    out.push({ title: "Achsenoptionen", open: true, controls: () => {
      const c: PaneControl[] = [
        { type: "heading", label: "Achsentyp" },
        { type: "radio", label: "Automatische Auswahl basierend auf Daten", checked: () => !o()?.axisType || o()?.axisType === "auto", pick: () => set({ axisType: undefined }) },
        { type: "radio", label: "Textachse", checked: () => o()?.axisType === "text", pick: () => set({ axisType: "text" }) },
        { type: "radio", label: "Datumsachse", checked: () => o()?.axisType === "date", pick: () => set({ axisType: "date" }) },
      ];
      if (dateAxis) {
        const unit = (label: string, key: "baseTimeUnit" | "majorTimeUnit") => ({ type: "select", label, options: [{ label: "Tage", value: "days" }, { label: "Monate", value: "months" }, { label: "Jahre", value: "years" }],
          value: () => o()?.[key] ?? "days", set: (v: string) => set({ [key]: v as AxisOptions["baseTimeUnit"] }) } as PaneControl);
        c.push({ type: "heading", label: "Grenzen" }, bound("Minimum", "min"), bound("Maximum", "max"), { type: "heading", label: "Einheiten" }, bound("Hauptstriche", "majorUnit"), unit("Haupteinheit", "majorTimeUnit"), unit("Basis", "baseTimeUnit"));
      }
      c.push(
        { type: "heading", label: kindInfo(m.kind).dir === "bar" ? "Horizontale Achse schneidet" : "Vertikale Achse schneidet" },
        { type: "radio", label: "Automatisch", checked: () => crosses() === undefined, pick: () => setCross(undefined) },
        { type: "radio", label: dateAxis ? "Bei Datum" : "Bei Rubriknummer", checked: () => typeof crosses() === "number", pick: () => setCross(1) },
        { type: "number", label: dateAxis ? "Datum" : "Bei Rubriknummer", min: 1, value: () => (typeof crosses() === "number" ? (crosses() as number) : 1), set: (v: number) => setCross(Math.max(1, Math.round(v))), disabled: () => typeof crosses() !== "number" },
        { type: "radio", label: dateAxis ? "Bei maximalem Datum" : "Bei größter Rubrik", checked: () => crosses() === "max", pick: () => setCross("max") },
      );
      if (!dateAxis) c.push(
        { type: "heading", label: "Achsenposition" },
        { type: "radio", label: "Auf Teilstrichen", checked: () => onTicks(m), pick: () => set({ onTicks: defaultOnTicks(m) ? undefined : true }) },
        { type: "radio", label: "Zwischen Teilstrichen", checked: () => !onTicks(m), pick: () => set({ onTicks: defaultOnTicks(m) ? false : undefined }) },
      );
      c.push({ type: "check", label: dateAxis ? "Datumsangaben in umgekehrter Reihenfolge" : "Kategorien in umgekehrter Reihenfolge", value: () => !!o()?.reverse, set: (on) => set({ reverse: on || undefined }) });
      return c;
    } });
  }
  out.push({ title: "Teilstriche", controls: () => [
    ...(which === "catAxis" && !scatter && !dateAxis ? [{ type: "number", label: "Intervall zwischen Teilstrichen", min: 1, max: 31999, value: () => o()?.tickMarkSkip ?? 1, set: (v: number) => set({ tickMarkSkip: Math.round(v) === 1 ? undefined : Math.round(v) }) } as PaneControl] : []),
    { type: "select", label: "Haupttyp", options: TICKS, value: () => o()?.majorTick ?? "none", set: (v) => set({ majorTick: v === "none" ? undefined : (v as AxisOptions["majorTick"]) }) },
    { type: "select", label: "Hilfstyp", options: TICKS, value: () => o()?.minorTick ?? "none", set: (v) => set({ minorTick: v === "none" ? undefined : (v as AxisOptions["minorTick"]) }) },
  ] });
  out.push({ title: "Beschriftungen", controls: () => {
    const c: PaneControl[] = [];
    if (which === "catAxis" && !scatter && !dateAxis) {
      c.push(
        { type: "heading", label: "Intervall zwischen Beschriftungen" },
        { type: "radio", label: "Automatisch", checked: () => o()?.tickLblSkip === undefined, pick: () => set({ tickLblSkip: undefined }) },
        { type: "radio", label: "Intervalleinheit angeben", checked: () => o()?.tickLblSkip !== undefined, pick: () => set({ tickLblSkip: 1 }) },
        { type: "number", label: "Intervalleinheit angeben", min: 1, max: 31999, value: () => o()?.tickLblSkip ?? 1, set: (v) => set({ tickLblSkip: Math.max(1, Math.round(v)) }), disabled: () => o()?.tickLblSkip === undefined },
      );
    }
    if (which === "catAxis" && !scatter) c.push({ type: "number", label: "Abstand zur Achse", min: 0, max: 1000, value: () => o()?.lblOffset ?? 100, set: (v) => set({ lblOffset: Math.round(v) === 100 ? undefined : Math.round(v) }) });
    c.push({ type: "select", label: "Beschriftungsposition", options: [{ label: "Achsennah", value: "nextTo" }, { label: "Hoch", value: "high" }, { label: "Niedrig", value: "low" }, { label: "Ohne", value: "none" }],
      value: () => o()?.labelPos ?? "nextTo", set: (v) => set({ labelPos: v === "nextTo" ? undefined : (v as AxisOptions["labelPos"]) }) });
    return c;
  } });
  out.push(numberSection(() => o()?.numFmt, (code) => set({ numFmt: code })));
  return out;
}

/** Excel's own Achsenposition of a chart type: areas on the tick marks, the others between them. */
function defaultOnTicks(m: ChartModel): boolean {
  const f = kindInfo(m.kind).family;
  return f === "area" || f === "surface";
}
function onTicks(m: ChartModel): boolean {
  return m.format?.catAxis?.onTicks ?? defaultOnTicks(m);
}

/** Datenreihe zeichnen auf (Primärachse / Sekundärachse). */
function axisChoice(m: ChartModel, i: number, host: PaneHost): PaneControl[] {
  return [
    { type: "heading", label: "Datenreihe zeichnen auf" },
    { type: "radio", label: "Primärachse", checked: () => !m.series[i]?.secondary, pick: () => host.edit((c) => { delete c.series[i].secondary; delete c.raw; }) },
    { type: "radio", label: "Sekundärachse", checked: () => !!m.series[i]?.secondary, pick: () => host.edit((c) => { c.series[i].secondary = true; delete c.raw; }),
      disabled: () => m.series.length < 2 || isScatter(m.kind) && false },
  ];
}

export function seriesSections(m: ChartModel, i: number, host: PaneHost): PaneSection[] {
  const fam = kindInfo(m.series[i]?.kind ?? m.kind).family;
  if (isPie(m.kind)) {
    return [{ title: "Reihenoptionen", open: true, controls: () => [
      ...axisChoice(m, i, host),
      { type: "number", label: "Winkel des ersten Segments", unit: "°", min: 0, max: 360, slider: true, value: () => m.firstSliceAngle ?? 0, set: (v) => host.edit((c) => { c.firstSliceAngle = Math.round(v) || undefined; }) },
      { type: "number", label: m.kind === "doughnut" ? "Ringexplosion" : "Kreisexplosion", unit: " %", min: 0, max: 400, slider: true, value: () => m.series[i]?.explosion ?? 0, set: (v) => host.edit((c) => { c.series[i].explosion = Math.round(v) || undefined; }) },
      ...(m.kind === "doughnut" ? [{ type: "number", label: "Innenringgröße", unit: " %", min: 10, max: 90, slider: true, value: () => m.holeSize ?? 75, set: (v: number) => host.edit((c) => { c.holeSize = Math.round(v) === 75 ? undefined : Math.round(v); }) } as PaneControl] : []),
    ] }];
  }
  const d = barDefaults(m.kind);
  const bars = fam === "bar";
  return [{ title: "Reihenoptionen", open: true, controls: () => [
    ...axisChoice(m, i, host),
    ...(bars ? [
      { type: "number", label: "Reihenachsenüberlappung", unit: " %", min: -100, max: 100, slider: true, value: () => m.overlap ?? d.overlap ?? 0, set: (v: number) => host.edit((c) => { c.overlap = Math.round(v); }) } as PaneControl,
      { type: "number", label: "Abstandsbreite", unit: " %", min: 0, max: 500, slider: true, value: () => m.gapWidth ?? d.gap, set: (v: number) => host.edit((c) => { c.gapWidth = Math.round(v); }) } as PaneControl,
    ] : []),
  ] }];
}

export function pointSections(m: ChartModel, i: number, k: number, host: PaneHost): PaneSection[] {
  if (isPie(m.kind)) {
    const p = () => m.series[i]?.points?.find((x) => x.idx === k);
    return [{ title: "Reihenoptionen", open: true, controls: () => [
      ...axisChoice(m, i, host),
      { type: "number", label: "Winkel des ersten Segments", unit: "°", min: 0, max: 360, slider: true, value: () => m.firstSliceAngle ?? 0, set: (v) => host.edit((c) => { c.firstSliceAngle = Math.round(v) || undefined; }) },
      { type: "number", label: "Punktexplosion", unit: " %", min: 0, max: 400, slider: true, value: () => p()?.explosion ?? m.series[i]?.explosion ?? 0,
        set: (v) => host.edit((c) => {
          const s = c.series[i];
          let pt = s.points?.find((x) => x.idx === k);
          if (!pt) { pt = { idx: k }; (s.points ??= []).push(pt); s.points.sort((a, b) => a.idx - b.idx); }
          pt.explosion = Math.round(v) || undefined;
          if (!pt.explosion && Object.keys(pt).filter((x) => x !== "idx" && x !== "explosion").length === 0) s.points = s.points!.filter((x) => x !== pt);
          if (!s.points?.length) delete s.points;
        }) },
    ] }];
  }
  return seriesSections(m, i, host);
}

export function labelSections(m: ChartModel, host: PaneHost): PaneSection[] {
  const o = () => m.format?.dataLabels;
  const set = (patch: Partial<DataLabelOptions>) => host.edit((c) => {
    const d = ((c.format ??= {}).dataLabels ??= {});
    Object.assign(d, patch);
    for (const k of Object.keys(d) as (keyof DataLabelOptions)[]) if (d[k] === undefined) delete d[k];
  });
  const fam = kindInfo(m.kind).family;
  const pie = isPie(m.kind), bars = fam === "bar", stacked = kindInfo(m.kind).grouping === "stacked" || kindInfo(m.kind).grouping === "percentStacked";
  const positions: [DataLabelOptions["pos"], string][] = pie ? [["ctr", "Zentriert"], ["inEnd", "Am Ende innerhalb"], ["outEnd", "Am Ende außerhalb"], ["bestFit", "Größe anpassen"]]
    : bars ? [["ctr", "Zentriert"], ["inEnd", "Am Ende innerhalb"], ["inBase", "Am Anfang innerhalb"], ...(stacked ? [] : [["outEnd", "Am Ende außerhalb"]] as [DataLabelOptions["pos"], string][])]
      : fam === "area" ? [] : [["ctr", "Zentriert"], ["l", "Links"], ["r", "Rechts"], ["t", "Über"], ["b", "Unter"]];
  const defPos = pie ? "bestFit" : bars ? (stacked ? "ctr" : "outEnd") : "r";
  const rangeText = () => o()?.fromCells?.ref ?? "";
  return [
    { title: "Beschriftungsoptionen", open: true, controls: () => [
      { type: "heading", label: "Beschriftung enthält" },
      { type: "check", label: "Wert aus Zellen", value: () => !!o()?.fromCells, set: (on) => set({ fromCells: on ? { ref: m.series[0]?.cat?.ref ?? "", cache: (m.series[0]?.cat?.cache ?? []).map((v) => (v === null ? null : String(v))) } : undefined }) },
      { type: "text", label: "Bereich der Datenbeschriftung", value: rangeText, set: (v) => set({ fromCells: v ? { ref: v, cache: o()?.fromCells?.cache ?? [] } : undefined }), disabled: () => !o()?.fromCells },
      { type: "check", label: "Datenreihenname", value: () => !!o()?.showSer, set: (on) => set({ showSer: on || undefined }) },
      { type: "check", label: isScatter(m.kind) ? "X-Wert" : "Rubrikenname", value: () => !!o()?.showCat, set: (on) => set({ showCat: on || undefined }) },
      { type: "check", label: isScatter(m.kind) ? "Y-Wert" : "Wert", value: () => o()?.showVal !== false, set: (on) => set({ showVal: on ? undefined : false }) },
      ...(pie ? [{ type: "check", label: "Prozentsatz", value: () => !!o()?.showPercent, set: (on: boolean) => set({ showPercent: on || undefined }) } as PaneControl] : []),
      ...(fam === "bubble" ? [{ type: "check", label: "Blasengröße", value: () => !!o()?.showBubble, set: (on: boolean) => set({ showBubble: on || undefined }) } as PaneControl] : []),
      { type: "check", label: "Führungslinien anzeigen", value: () => (pie ? o()?.leaderLines !== false : !!o()?.leaderLines), set: (on) => set({ leaderLines: pie ? (on ? undefined : false) : on || undefined }) },
      { type: "check", label: "Legendensymbol", value: () => !!o()?.showLegendKey, set: (on) => set({ showLegendKey: on || undefined }) },
      { type: "select", label: "Trennzeichen", options: [{ label: ", (Komma)", value: ", " }, { label: "; (Semikolon)", value: "; " }, { label: ". (Punkt)", value: ". " }, { label: "(Neue Zeile)", value: "\n" }, { label: "(Leerzeichen)", value: " " }],
        value: () => o()?.separator ?? "; ", set: (v) => set({ separator: v === "; " ? undefined : v }) },
      { type: "button", label: "Beschriftungstext zurücksetzen", run: () => host.edit((c) => { if (c.format?.dataLabels) { const { fill, line, effects, text } = c.format.dataLabels; c.format.dataLabels = { ...(fill ? { fill } : {}), ...(line ? { line } : {}), ...(effects ? { effects } : {}), ...(text ? { text } : {}) }; } }) },
      ...(positions.length ? [{ type: "heading", label: "Beschriftungsposition" } as PaneControl] : []),
      ...positions.map(([v, label]) => ({ type: "radio", label, checked: () => (o()?.pos ?? defPos) === v, pick: () => set({ pos: v === defPos ? undefined : v }) }) as PaneControl),
    ] },
    numberSection(() => o()?.numFmt, (code) => set({ numFmt: code })),
  ];
}

export function trendSections(m: ChartModel, i: number, host: PaneHost): PaneSection[] {
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
    { type: "number", label: "Grad", min: 2, max: 6, value: () => t()?.order ?? 2, set: (v: number) => set({ order: Math.round(v) }), disabled: () => t()?.type !== "poly" },
    type("power", "Potenz"), type("movingAvg", "Gleitender Durchschnitt"),
    { type: "number", label: "Periode", min: 2, max: 255, value: () => t()?.period ?? 2, set: (v: number) => set({ period: Math.round(v) }), disabled: () => t()?.type !== "movingAvg" },
    { type: "heading", label: "Name der Trendlinie" },
    { type: "radio", label: "Automatisch", checked: () => !t()?.name, pick: () => set({ name: undefined }) },
    { type: "radio", label: "Benutzerdefiniert", checked: () => !!t()?.name, pick: () => set({ name: t()?.name ?? "Trendlinie" }) },
    { type: "text", label: "Benutzerdefiniert", value: () => t()?.name ?? "", set: (v: string) => set({ name: v || undefined }), disabled: () => !t()?.name },
    { type: "heading", label: "Prognose" },
    { type: "number", label: "Vorwärts", unit: " Perioden", min: 0, value: () => t()?.forward ?? 0, set: (v) => set({ forward: v || undefined }) },
    { type: "number", label: "Rückwärts", unit: " Perioden", min: 0, value: () => t()?.backward ?? 0, set: (v) => set({ backward: v || undefined }) },
    { type: "check", label: "Schnittpunkt =", value: () => t()?.intercept !== undefined, set: (on) => set({ intercept: on ? 0 : undefined }), disabled: () => t()?.type === "movingAvg" || t()?.type === "log" || t()?.type === "power" },
    { type: "number", label: "Schnittpunkt =", value: () => t()?.intercept ?? 0, set: (v) => set({ intercept: v }), disabled: () => t()?.intercept === undefined },
    { type: "check", label: "Formel im Diagramm anzeigen", value: () => !!t()?.showEq, set: (on) => set({ showEq: on }) },
    { type: "check", label: "Bestimmtheitsmaß im Diagramm darstellen", value: () => !!t()?.showR2, set: (on) => set({ showR2: on }) },
  ] }];
}

export function errorSections(m: ChartModel, i: number, host: PaneHost): PaneSection[] {
  const b = () => m.series[i]?.errorBars;
  const set = (patch: Partial<ErrorBars>) => host.edit((c) => { const s = c.series[i]; if (s.errorBars) Object.assign(s.errorBars, patch); });
  const amount = (v: ErrorBars["type"], label: string, unit?: string, def?: number): PaneControl[] => [
    { type: "radio", label, checked: () => b()?.type === v, pick: () => set({ type: v, value: def }) },
    ...(def !== undefined ? [{ type: "number", label, unit, min: 0, value: () => (b()?.type === v ? b()?.value ?? def : def), set: (x: number) => set({ value: x }), disabled: () => b()?.type !== v } as PaneControl] : []),
  ];
  return [{ title: kindInfo(m.kind).dir === "bar" ? "Horizontaler Fehlerindikator" : "Vertikaler Fehlerindikator", open: true, controls: () => [
    { type: "heading", label: "Richtung" },
    ...([["both", "Beide"], ["minus", "Minus"], ["plus", "Plus"]] as const).map(([v, label]) => ({ type: "radio", label, checked: () => b()?.dir === v, pick: () => set({ dir: v }) }) as PaneControl),
    { type: "heading", label: "Linienende" },
    { type: "radio", label: "Ohne Abschluss", checked: () => !!b()?.noEndCap, pick: () => set({ noEndCap: true }) },
    { type: "radio", label: "Mit Abschluss", checked: () => !b()?.noEndCap, pick: () => set({ noEndCap: undefined }) },
    { type: "heading", label: "Fehlerbetrag" },
    ...amount("fixedVal", "Fester Wert", undefined, 1), ...amount("percentage", "Prozentsatz", " %", 5), ...amount("stdDev", "Standardabweichung(en)", undefined, 1),
    ...amount("stdErr", "Standardfehler"),
  ] }];
}
