// Diagrammentwurf on slides (D-070): the chart part read into the chart engine's model (shared with
// RaabeCalc, D-062), changed and written back: Diagrammtyp ändern, Schnelllayout (Excel's recorded
// layouts, the same in PowerPoint), Farben ändern, Zeile/Spalte tauschen, Daten bearbeiten (a small
// sheet of the values the chart keeps, as PowerPoint's chart data window shows them).

import { QUICK_LAYOUTS, type ChartKind, type ChartModel } from "@raabe/charts";
import { chartFromXml, chartToXml } from "@raabe/ooxml";
import { h, showDialog, type MenuEntry } from "@raabe/shell";
import type { PresentApp } from "./app";

function chartPath(app: PresentApp): string | undefined {
  const v = app.editor.selectedViews.find((x) => x.frame === "chart");
  return v?.chart;
}

export function hasChart(app: PresentApp): boolean {
  return !!chartPath(app);
}

function edit(app: PresentApp, label: string, fn: (m: ChartModel) => void): void {
  const path = chartPath(app);
  if (!path) return;
  const xml = app.deck.pptx.text(path);
  if (!xml) return;
  const m = chartFromXml(xml);
  fn(m);
  app.editor.edit(label, () => app.deck.pptx.put(path, chartToXml(m)));
  app.deck.markChanged([path]);
}

export const CHART_KINDS: [string, ChartKind][] = [["Gruppierte Säulen", "column"], ["Gestapelte Säulen", "columnStacked"], ["Gruppierte Balken", "bar"], ["Gestapelte Balken", "barStacked"],
  ["Linie", "line"], ["Linie mit Datenpunkten", "lineMarkers"], ["Fläche", "area"], ["Kreis", "pie"], ["Ring", "doughnut"], ["Punkt (X Y)", "scatter"]];

export function changeType(app: PresentApp, kind: ChartKind): void {
  edit(app, "Diagrammtyp ändern", (m) => { m.kind = kind; });
}

export function quickLayout(app: PresentApp, n: number): void {
  edit(app, "Schnelllayout", (c) => {
    const fam: Record<string, string> = { column: "column", columnStacked: "column", bar: "bar", barStacked: "bar", line: "line", lineMarkers: "lineMarkers", area: "area", pie: "pie", doughnut: "pie", scatter: "scatter" };
    const l = (QUICK_LAYOUTS[fam[c.kind] ?? ""] ?? [])[n - 1];
    if (!l) return;
    c.title.show = l.title;
    c.legend = l.legend;
    c.dataLabels = l.labels;
    c.axisTitles = { ...(l.catTitle ? { cat: c.axisTitles.cat ?? "" } : {}), ...(l.valTitle ? { val: c.axisTitles.val ?? "" } : {}) };
    c.gridlines = !!l.majorGrid;
  });
}

export function setColors(app: PresentApp, scheme: number): void {
  edit(app, "Farben ändern", (m) => { m.colors = scheme; });
}

export function switchRowCol(app: PresentApp): void {
  edit(app, "Zeile/Spalte tauschen", (m) => {
    const cats = (m.series[0]?.cat?.cache ?? []).map((c) => String(c ?? ""));
    const names = m.series.map((s, i) => s.name?.text ?? `Datenreihe ${i + 1}`);
    m.series = cats.map((cat, ci) => ({ name: { text: cat }, cat: { cache: names }, val: { cache: m.series.map((s) => s.val.cache[ci] ?? null) } }));
  });
}

/** Daten bearbeiten: the chart's values in a grid (categories down, series across), as the chart keeps them. */
export async function editData(app: PresentApp): Promise<void> {
  const path = chartPath(app);
  const xml = path ? app.deck.pptx.text(path) : null;
  if (!path || !xml) return;
  const m = chartFromXml(xml);
  const cats = (m.series[0]?.cat?.cache ?? []).map((c) => String(c ?? ""));
  const rows = Math.max(cats.length, ...m.series.map((s) => s.val.cache.length)) + 2;
  const cols = m.series.length + 2;
  const inputs: HTMLInputElement[][] = [];
  const table = h("table", { class: "chart-data" });
  for (let r = 0; r < rows; r++) {
    const tr = h("tr");
    inputs.push([]);
    for (let c = 0; c < cols; c++) {
      const v = r === 0 ? (c === 0 ? "" : m.series[c - 1]?.name?.text ?? (c - 1 < m.series.length ? `Datenreihe ${c}` : ""))
        : c === 0 ? cats[r - 1] ?? "" : m.series[c - 1]?.val.cache[r - 1]?.toString().replace(".", ",") ?? "";
      const inp = h("input", { type: "text", value: v }) as HTMLInputElement;
      inputs[r].push(inp);
      tr.append(h("td", null, inp));
    }
    table.append(tr);
  }
  const ok = await showDialog({ title: "Diagramm in Microsoft PowerPoint", body: h("div", { class: "chart-data-wrap" }, table), buttons: [{ label: "OK", value: true, primary: true }, { label: "Abbrechen", value: false }], cancel: false });
  if (!ok) return;
  edit(app, "Daten bearbeiten", (c) => {
    const lastRow = inputs.reduce((n, row, r) => (r > 0 && row.some((x) => x.value.trim()) ? r : n), 0);
    const catsNew = inputs.slice(1, lastRow + 1).map((row) => row[0].value);
    const series = [];
    for (let k = 1; k < cols; k++) {
      const name = inputs[0][k].value.trim();
      const vals = inputs.slice(1, lastRow + 1).map((row) => { const t = row[k].value.trim().replace(",", "."); return t === "" ? null : Number(t); });
      if (!name && vals.every((v) => v === null)) continue;
      series.push({ ...(c.series[k - 1] ?? {}), name: { text: name || `Datenreihe ${k}` }, cat: { cache: catsNew }, val: { cache: vals.map((v) => (v === null || Number.isNaN(v) ? null : v)) } });
    }
    c.series = series as ChartModel["series"];
  });
}

/**
 * Diagrammelement hinzufügen (D-072): PowerPoint's menu (recorded) on the chart model RaabeCalc's engine
 * keeps (title, axis titles, legend, data labels, gridlines, trend lines, error bars); what the model does not
 * cover (axes on and off, data table, label positions, minor gridlines) is greyed, "Weitere ..." opens the pane.
 */
export function chartElementsMenu(app: PresentApp): MenuEntry[] {
  const set = (label: string, fn: (m: ChartModel) => void) => () => edit(app, label, fn);
  const more = (label: string): MenuEntry => ({ label, run: () => void app.commands.run("draw.format") });
  const off = (label: string): MenuEntry => ({ label, run: () => {}, disabled: true });
  return [
    { label: "Achsen", submenu: [off("Primär horizontal"), off("Primär vertikal"), more("Weitere Achsenoptionen...")] },
    { label: "Achsentitel", submenu: [
      { label: "Primär horizontal", run: set("Achsentitel", (m) => { m.axisTitles = { ...m.axisTitles, cat: m.axisTitles.cat === undefined ? "" : undefined }; }) },
      { label: "Primär vertikal", run: set("Achsentitel", (m) => { m.axisTitles = { ...m.axisTitles, val: m.axisTitles.val === undefined ? "" : undefined }; }) },
      more("Weitere Optionen für Achsentitel...")] },
    { label: "Diagrammtitel", submenu: [
      { label: "Ohne", run: set("Diagrammtitel", (m) => { m.title.show = false; }) },
      { label: "Über Diagramm", run: set("Diagrammtitel", (m) => { m.title.show = true; }) },
      { label: "Mittig überlagert", run: set("Diagrammtitel", (m) => { m.title.show = true; }) },
      more("Weitere Titeloptionen...")] },
    { label: "Datenbeschriftungen", submenu: [
      { label: "Keine", run: set("Datenbeschriftungen", (m) => { m.dataLabels = false; }) },
      ...["Mitte", "Am Ende innerhalb", "Am Anfang innerhalb", "Am Ende außerhalb"].map((l) => ({ label: l, run: set("Datenbeschriftungen", (m) => { m.dataLabels = true; }) })),
      off("Datenlegende"), more("Weitere Datenbeschriftungsoptionen...")] },
    { label: "Datentabelle", submenu: [off("Keine"), off("Mit Legendensymbolen"), off("Ohne Legendensymbole"), more("Weitere Datentabellenoptionen...")] },
    { label: "Fehlerindikatoren", submenu: [
      { label: "Keine", run: set("Fehlerindikatoren", (m) => { for (const s of m.series) delete s.errorBars; }) },
      ...([["Standardfehler", "stdErr", undefined], ["Prozentsatz", "percentage", 5], ["Standardabweichung", "stdDev", 1]] as [string, "stdErr" | "percentage" | "stdDev", number | undefined][])
        .map(([l, type, value]) => ({ label: l, run: set("Fehlerindikatoren", (m) => { for (const s of m.series) s.errorBars = { dir: "both", type, ...(value !== undefined ? { value } : {}) }; }) })),
      more("Weitere Fehlerindikatorenoptionen...")] },
    { label: "Gitternetzlinien", submenu: [
      { label: "Primäres Hauptgitter horizontal", run: set("Gitternetzlinien", (m) => { m.gridlines = !m.gridlines; }) },
      off("Primäres Hauptgitter vertikal"), off("Primäres Hilfsgitter horizontal"), off("Primäres Hilfsgitter vertikal"), more("Weitere Optionen für Gitternetzlinien...")] },
    { label: "Legende", submenu: [
      ...([["Keine", "none"], ["Rechts", "r"], ["Oben", "t"], ["Links", "l"], ["Unten", "b"]] as [string, "none" | "r" | "t" | "l" | "b"][]).map(([l, pos]) => ({ label: l, run: set("Legende", (m) => { m.legend = pos; }) })),
      more("Weitere Legendenoptionen...")] },
    { label: "Linien", submenu: [off("Keine"), off("Bezugslinien"), off("Spannweitenlinien")] },
    { label: "Trendlinie", submenu: [
      { label: "Keine", run: set("Trendlinie", (m) => { for (const s of m.series) delete s.trendline; }) },
      ...([["Linear", "linear"], ["Exponentiell", "exp"], ["Gleitender Durchschnitt", "movingAvg"]] as [string, "linear" | "exp" | "movingAvg"][])
        .map(([l, type]) => ({ label: l, run: set("Trendlinie", (m) => { const s = m.series[0]; if (s) s.trendline = { type, showEq: false, showR2: false, ...(type === "movingAvg" ? { period: 2 } : {}) }; }) })),
      more("Weitere Trendlinienoptionen...")] },
    { label: "Pos./Neg. Abweichung", submenu: [off("Keine"), off("Pos./Neg. Abweichung")] },
  ];
}
