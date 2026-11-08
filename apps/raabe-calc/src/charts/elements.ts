// Diagrammelemente (D-115): what the "+" button beside a chart and Diagrammentwurf > Diagrammelement hinzufügen
// switch on and off, as model changes. Both lists come from the same table: an element's check box in the flyout
// (on/off with Excel's default), its side menu (the arrow at the right end of the entry) and the ribbon's submenu
// with "Keine" first and its "Weitere ...optionen..." entry at the end. Labels as recorded in German Excel 365
// (results/2026-11-09-excel-chart-buttons.json; the ribbon's from the inventory).

import { isBar, isPie, isScatter, kindInfo, type ChartModel, type DataLabelOptions, type ErrorBars, type Trendline } from "@raabe/charts";
import type { ChartElement } from "./formatpane";

/** One choice of a side menu: a check box (several may be on) or one of a set (a radio choice). */
export interface ElementChoice {
  /** The label in the flyout's side menu and in the ribbon's submenu (when they differ: [flyout, ribbon]). */
  label: string | [string, string];
  /** A check box in the flyout (Achsen, Achsentitel, Gitternetzlinien); otherwise one choice of a set. */
  check?: boolean;
  on: (m: ChartModel) => boolean;
  /** Switches the choice on (sets) or, for a check box, to `on`. */
  set: (m: ChartModel, on: boolean) => void;
}

export interface ElementEntry {
  id: "axes" | "axisTitles" | "title" | "labels" | "dataTable" | "errorBars" | "gridlines" | "legend" | "lines" | "trendline" | "updown";
  label: string;
  /** The entry's check box in the flyout. */
  on: (m: ChartModel) => boolean;
  /** Ticking or clearing the check box (Excel's default choice when ticked). */
  toggle: (m: ChartModel, on: boolean) => void;
  choices: ElementChoice[];
  /** The ribbon's "Keine" entry (sets everything off); absent where the ribbon has none (Achsen, Gitternetzlinien). */
  none?: boolean;
  /** The ribbon's "Weitere ...optionen..." label (the flyout says "Weitere Optionen..."). */
  more: string;
  /** The element whose Format pane "Weitere Optionen..." opens. */
  pane: (m: ChartModel) => ChartElement;
  /** Shown but greyed out (the ribbon's Linien and Pos./Neg. Abweichung on a column chart). */
  disabled?: boolean;
  /** Only in the ribbon's menu. */
  ribbonOnly?: boolean;
}

const fmt = (m: ChartModel) => (m.format ??= {});
const tidy = (m: ChartModel) => { if (m.format && !Object.keys(m.format).length) delete m.format; };

/** Charts with category/value axes (not pies, doughnuts, the newer chartEx kinds, surfaces …). */
export const hasAxes = (m: ChartModel) => {
  const f = kindInfo(m.kind).family;
  return !isPie(m.kind) && !["treemap", "sunburst", "waterfall", "funnel", "boxWhisker", "histogram", "pareto"].includes(f) || f === "histogram";
};

const axisVisible = (m: ChartModel, a: "catAxis" | "valAxis") => m.format?.[a]?.visible !== false;
function setAxis(m: ChartModel, a: "catAxis" | "valAxis", on: boolean) {
  const f = fmt(m);
  const o = { ...(f[a] ?? {}) };
  if (on) delete o.visible; else o.visible = false;
  if (Object.keys(o).length) f[a] = o; else delete f[a];
  tidy(m);
}

/** Datenbeschriftungen positions Excel offers for the chart's kind (flyout label, ribbon label, position). */
export function labelPositions(m: ChartModel): [string, string, DataLabelOptions["pos"] | "callout"][] {
  const info = kindInfo(m.kind);
  if (isPie(m.kind)) {
    if (info.family === "doughnut") return [["Datenlegende", "Datenlegende", "callout"]];
    return [["Zentriert", "Mitte", "ctr"], ["Innerhalb Ende", "Am Ende innerhalb", "inEnd"], ["Außerhalb Ende", "Am Ende außerhalb", "outEnd"],
      ["Optimale Anpassung", "Optimale Anpassung", "bestFit"], ["Datenlegende", "Datenlegende", "callout"]];
  }
  if (info.family === "line" || isScatter(m.kind)) {
    return [["Zentriert", "Mitte", "ctr"], ["Links", "Links", "l"], ["Rechts", "Rechts", "r"], ["Über", "Oberhalb", "t"], ["Unter", "Unterhalb", "b"],
      ["Datenlegende", "Datenlegende", "callout"]];
  }
  if (info.family === "area") return [["Datenlegende", "Datenlegende", "callout"]];
  if (info.grouping === "stacked" || info.grouping === "percentStacked") {
    return [["Zentriert", "Mitte", "ctr"], ["Innerhalb Ende", "Am Ende innerhalb", "inEnd"], ["Innerhalb Basis", "Am Anfang innerhalb", "inBase"],
      ["Datenlegende", "Datenlegende", "callout"]];
  }
  return [["Zentriert", "Mitte", "ctr"], ["Innerhalb Ende", "Am Ende innerhalb", "inEnd"], ["Innerhalb Basis", "Am Anfang innerhalb", "inBase"],
    ["Außerhalb Ende", "Am Ende außerhalb", "outEnd"], ["Datenlegende", "Datenlegende", "callout"]];
}

function setLabels(m: ChartModel, pos: DataLabelOptions["pos"] | "callout" | null) {
  if (pos === null) {
    m.dataLabels = false;
    if (m.format?.dataLabels) { delete m.format.dataLabels.pos; delete m.format.dataLabels.showCat; if (!Object.keys(m.format.dataLabels).length) delete m.format.dataLabels; }
    tidy(m);
    return;
  }
  m.dataLabels = true;
  const f = fmt(m);
  const d: DataLabelOptions = { ...(f.dataLabels ?? {}) };
  delete d.pos;
  delete d.showCat;
  // Datenlegende: the category and the value in a callout; drawn as both parts beside the point.
  if (pos === "callout") d.showCat = true;
  else d.pos = pos;
  if (Object.keys(d).length) f.dataLabels = d; else delete f.dataLabels;
  tidy(m);
}

const labelPos = (m: ChartModel): string | null => {
  if (!m.dataLabels) return null;
  const d = m.format?.dataLabels;
  if (d?.showCat && !d.pos) return "callout";
  return d?.pos ?? defaultLabelPos(m);
};

/** Where Excel puts data labels by default for the kind (Datenbeschriftungen ticked). */
export function defaultLabelPos(m: ChartModel): string {
  const info = kindInfo(m.kind);
  if (isPie(m.kind)) return "bestFit";
  if (info.family === "line" || isScatter(m.kind)) return "r";
  if (info.family === "area") return "ctr";
  if (info.grouping === "stacked" || info.grouping === "percentStacked") return "ctr";
  return "outEnd";
}

function setErrorBars(m: ChartModel, type: ErrorBars["type"] | null) {
  for (const s of m.series) {
    if (type === null) delete s.errorBars;
    else s.errorBars = { dir: "both", type, ...(type === "percentage" ? { value: 5 } : type === "stdDev" ? { value: 1 } : {}) };
  }
}
const errType = (m: ChartModel) => m.series.find((s) => s.errorBars)?.errorBars?.type ?? null;

type TrendKind = "linear" | "exp" | "forecast" | "movingAvg";
export function trendOf(kind: TrendKind): Trendline {
  switch (kind) {
    case "linear": return { type: "linear", showEq: false, showR2: false };
    case "exp": return { type: "exp", showEq: false, showR2: false };
    case "forecast": return { type: "linear", forward: 2, showEq: false, showR2: false };
    case "movingAvg": return { type: "movingAvg", period: 2, showEq: false, showR2: false };
  }
}
const trendKind = (t: Trendline | undefined): TrendKind | null => !t ? null : t.type === "linear" ? (t.forward === 2 ? "forecast" : "linear") : t.type === "exp" ? "exp" : t.type === "movingAvg" ? "movingAvg" : null;

/** Series a trendline can go on: Excel asks which one when there are several (Trendlinie hinzufügen). */
export const trendAllowed = (m: ChartModel) => {
  const info = kindInfo(m.kind);
  return !isPie(m.kind) && !info.d3 && (info.grouping === "standard" || info.grouping === "clustered" || isScatter(m.kind)) && !isBar(m.kind)
    && ["bar", "line", "scatter", "area", "bubble"].includes(info.family) && m.series.length > 0;
};

/**
 * The entries of the flyout (and of the ribbon's menu, with `ribbon`) for a chart, in Excel's order.
 * `askSeries` picks the series for a trendline when there are several (Excel's "Trendlinie hinzufügen" dialog);
 * the flyout and the ribbon call it before the change, so the change itself stays one undo step.
 */
export function elementEntries(m: ChartModel, ribbon = false): ElementEntry[] {
  const axes = hasAxes(m);
  const pie = isPie(m.kind);
  const out: ElementEntry[] = [];
  const valAxis = (): ChartElement => ({ kind: "valAxis" });
  if (axes) {
    out.push({
      id: "axes", label: "Achsen", more: "Weitere Achsenoptionen...",
      on: (c) => axisVisible(c, "catAxis") || axisVisible(c, "valAxis"),
      toggle: (c, on) => { setAxis(c, "catAxis", on); setAxis(c, "valAxis", on); },
      choices: [
        { label: "Primär horizontal", check: true, on: (c) => axisVisible(c, isBar(c.kind) ? "valAxis" : "catAxis"), set: (c, on) => setAxis(c, isBar(c.kind) ? "valAxis" : "catAxis", on) },
        { label: "Primär vertikal", check: true, on: (c) => axisVisible(c, isBar(c.kind) ? "catAxis" : "valAxis"), set: (c, on) => setAxis(c, isBar(c.kind) ? "catAxis" : "valAxis", on) },
      ],
      pane: valAxis,
    });
    out.push({
      id: "axisTitles", label: "Achsentitel", more: "Weitere Optionen für Achsentitel...",
      on: (c) => c.axisTitles.cat !== undefined || c.axisTitles.val !== undefined,
      toggle: (c, on) => { c.axisTitles = on ? { cat: c.axisTitles.cat ?? "", val: c.axisTitles.val ?? "" } : {}; },
      choices: [
        { label: "Primär horizontal", check: true, on: (c) => c.axisTitles[isBar(c.kind) ? "val" : "cat"] !== undefined,
          set: (c, on) => { const k = isBar(c.kind) ? "val" : "cat"; if (on) c.axisTitles[k] ??= ""; else delete c.axisTitles[k]; } },
        { label: "Primär vertikal", check: true, on: (c) => c.axisTitles[isBar(c.kind) ? "cat" : "val"] !== undefined,
          set: (c, on) => { const k = isBar(c.kind) ? "cat" : "val"; if (on) c.axisTitles[k] ??= ""; else delete c.axisTitles[k]; } },
      ],
      pane: (c) => ({ kind: "axisTitle", axis: c.axisTitles.cat !== undefined ? "cat" : "val" }),
    });
  }
  const overlay = (c: ChartModel) => !!c.format?.title?.overlay;
  const setTitle = (c: ChartModel, show: boolean, over: boolean) => {
    c.title.show = show;
    const f = fmt(c);
    if (over) f.title = { ...(f.title ?? {}), overlay: true };
    else if (f.title) { delete f.title.overlay; if (!Object.keys(f.title).length) delete f.title; }
    tidy(c);
  };
  out.push({
    id: "title", label: "Diagrammtitel", more: "Weitere Titeloptionen...", none: true,
    on: (c) => c.title.show,
    toggle: (c, on) => setTitle(c, on, false),
    choices: [
      { label: "Über Diagramm", on: (c) => c.title.show && !overlay(c), set: (c) => setTitle(c, true, false) },
      { label: ["Zentriert überlagert", "Mittig überlagert"], on: (c) => c.title.show && overlay(c), set: (c) => setTitle(c, true, true) },
    ],
    pane: () => ({ kind: "title" }),
  });
  out.push({
    id: "labels", label: "Datenbeschriftungen", more: "Weitere Datenbeschriftungsoptionen...", none: true,
    on: (c) => c.dataLabels,
    toggle: (c, on) => setLabels(c, on ? (defaultLabelPos(c) as DataLabelOptions["pos"]) : null),
    choices: labelPositions(m).map(([fl, rb, pos]) => ({
      label: [fl, rb] as [string, string],
      on: (c: ChartModel) => labelPos(c) === pos,
      set: (c: ChartModel) => setLabels(c, pos as DataLabelOptions["pos"] | "callout"),
    })),
    pane: () => ({ kind: "labels" }),
  });
  const canTable = axes && !isScatter(m.kind) && !kindInfo(m.kind).d3 && kindInfo(m.kind).family !== "radar" && kindInfo(m.kind).family !== "histogram";
  if (canTable) {
    out.push({
      id: "dataTable", label: "Datentabelle", more: "Weitere Datentabellenoptionen...", none: true,
      on: (c) => !!c.dataTable,
      toggle: (c, on) => { if (on) c.dataTable = { legendKeys: true }; else delete c.dataTable; },
      choices: [
        { label: "Mit Legendensymbolen", on: (c) => !!c.dataTable?.legendKeys, set: (c) => { c.dataTable = { ...(c.dataTable ?? {}), legendKeys: true }; } },
        { label: "Ohne Legendensymbole", on: (c) => !!c.dataTable && !c.dataTable.legendKeys, set: (c) => { c.dataTable = { ...(c.dataTable ?? {}), legendKeys: false }; } },
      ],
      pane: () => ({ kind: "chart" }),
    });
  }
  const canErr = axes && !kindInfo(m.kind).d3 && kindInfo(m.kind).family !== "radar" && kindInfo(m.kind).family !== "histogram";
  if (canErr) {
    out.push({
      id: "errorBars", label: "Fehlerindikatoren", more: "Weitere Fehlerindikatorenoptionen...", none: true,
      on: (c) => c.series.some((s) => s.errorBars),
      toggle: (c, on) => setErrorBars(c, on ? "stdErr" : null),
      choices: [
        { label: "Standardfehler", on: (c) => errType(c) === "stdErr", set: (c) => setErrorBars(c, "stdErr") },
        { label: "Prozentsatz", on: (c) => errType(c) === "percentage", set: (c) => setErrorBars(c, "percentage") },
        { label: "Standardabweichung", on: (c) => errType(c) === "stdDev", set: (c) => setErrorBars(c, "stdDev") },
      ],
      pane: (c) => ({ kind: "errorBars", index: Math.max(0, c.series.findIndex((s) => s.errorBars)) }),
    });
  }
  if (axes) {
    const grid = (k: "minorGrid" | "catGrid" | "catMinorGrid") => !!m.format?.[k]?.show;
    void grid;
    const setGrid = (c: ChartModel, k: "minorGrid" | "catGrid" | "catMinorGrid", on: boolean) => {
      const f = fmt(c);
      if (on) f[k] = { ...(f[k] ?? {}), show: true };
      else if (f[k]) { delete f[k]!.show; if (!Object.keys(f[k]!).length) delete f[k]; }
      tidy(c);
    };
    const bar = isBar(m.kind);
    // Horizontal and vertical name the direction of the lines on screen: a bar chart's value gridlines are vertical.
    const valMajor: ElementChoice = { label: bar ? "Primäres Hauptgitter vertikal" : "Primäres Hauptgitter horizontal", check: true, on: (c) => c.gridlines, set: (c, on) => { c.gridlines = on; } };
    const catMajor: ElementChoice = { label: bar ? "Primäres Hauptgitter horizontal" : "Primäres Hauptgitter vertikal", check: true, on: (c) => !!c.format?.catGrid?.show, set: (c, on) => setGrid(c, "catGrid", on) };
    const valMinor: ElementChoice = { label: bar ? "Primäres Hilfsgitter vertikal" : "Primäres Hilfsgitter horizontal", check: true, on: (c) => !!c.format?.minorGrid?.show, set: (c, on) => setGrid(c, "minorGrid", on) };
    const catMinor: ElementChoice = { label: bar ? "Primäres Hilfsgitter horizontal" : "Primäres Hilfsgitter vertikal", check: true, on: (c) => !!c.format?.catMinorGrid?.show, set: (c, on) => setGrid(c, "catMinorGrid", on) };
    out.push({
      id: "gridlines", label: "Gitternetzlinien", more: "Weitere Optionen für Gitternetzlinien...",
      on: (c) => c.gridlines || !!c.format?.catGrid?.show || !!c.format?.minorGrid?.show || !!c.format?.catMinorGrid?.show,
      toggle: (c, on) => {
        c.gridlines = on;
        if (!on) for (const k of ["minorGrid", "catGrid", "catMinorGrid"] as const) setGrid(c, k, false);
      },
      choices: bar ? [catMajor, valMajor, catMinor, valMinor] : [valMajor, catMajor, valMinor, catMinor],
      pane: (c) => (c.gridlines ? { kind: "majorGrid" } : c.format?.catGrid?.show ? { kind: "catGrid" } : { kind: "majorGrid" }),
    });
  }
  if (kindInfo(m.kind).family !== "histogram" || ribbon) {
    const pos = (p: "r" | "t" | "l" | "b", label: string): ElementChoice => ({ label, on: (c) => c.legend === p, set: (c) => { c.legend = p; } });
    out.push({
      id: "legend", label: "Legende", more: "Weitere Legendenoptionen...", none: true,
      on: (c) => c.legend !== "none",
      // Ticking Legende puts it where a new chart has it (below), as Excel 365 does.
      toggle: (c, on) => { c.legend = on ? "b" : "none"; },
      choices: [pos("r", "Rechts"), pos("t", "Oben"), pos("l", "Links"), pos("b", "Unten")],
      pane: () => ({ kind: "legend" }),
    });
  }
  if (ribbon && axes && !pie) {
    out.push({ id: "lines", label: "Linien", more: "", on: () => false, toggle: () => {}, choices: [], pane: () => ({ kind: "chart" }), disabled: true, ribbonOnly: true });
  }
  if (trendAllowed(m) || (ribbon && axes)) {
    const set = (c: ChartModel, k: TrendKind | null, only?: number) => {
      c.series.forEach((s, i) => {
        if (only !== undefined && i !== only) return;
        if (k === null) delete s.trendline; else s.trendline = trendOf(k);
      });
    };
    const tk = (c: ChartModel) => trendKind(c.series.find((s) => s.trendline)?.trendline);
    out.push({
      id: "trendline", label: "Trendlinie", more: "Weitere Trendlinienoptionen...", none: true, disabled: !trendAllowed(m),
      on: (c) => c.series.some((s) => s.trendline),
      toggle: (c, on) => set(c, on ? "linear" : null),
      choices: [
        { label: "Linear", on: (c) => tk(c) === "linear", set: (c) => set(c, "linear") },
        { label: "Exponentiell", on: (c) => tk(c) === "exp", set: (c) => set(c, "exp") },
        { label: ["Linear Prognose", "Lineare Prognose"], on: (c) => tk(c) === "forecast", set: (c) => set(c, "forecast") },
        { label: "Gleitender Durchschnitt", on: (c) => tk(c) === "movingAvg", set: (c) => set(c, "movingAvg") },
      ],
      pane: (c) => ({ kind: "trend", index: Math.max(0, c.series.findIndex((s) => s.trendline)) }),
    });
  }
  if (ribbon && axes && !pie) {
    out.push({ id: "updown", label: "Pos./Neg. Abweichung", more: "", on: () => false, toggle: () => {}, choices: [], pane: () => ({ kind: "chart" }), disabled: true, ribbonOnly: true });
  }
  return out;
}

/** The label of a choice in the flyout (0) or the ribbon (1). */
export const choiceLabel = (c: ElementChoice, ribbon: boolean) => (typeof c.label === "string" ? c.label : c.label[ribbon ? 1 : 0]);
