// Diagrammentwurf on slides (D-070): the chart part read into the chart engine's model (shared with
// RaabeCalc, D-062), changed and written back: Diagrammtyp ändern, Schnelllayout (Excel's recorded
// layouts, the same in PowerPoint), Farben ändern, Zeile/Spalte tauschen, Daten bearbeiten (a small
// sheet of the values the chart keeps, as PowerPoint's chart data window shows them).

import { QUICK_LAYOUTS, type ChartKind, type ChartModel } from "@raabe/charts";
import { chartFromXml, chartToXml } from "@raabe/ooxml";
import { h, showDialog } from "@raabe/shell";
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
