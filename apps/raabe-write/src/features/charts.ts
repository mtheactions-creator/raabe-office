// Diagramm in RaabeWrite (Step 7b part 2, D-088): Einfügen > Diagramm with the chart types of Word's Diagramm einfügen
// dialog, Word's sample data, and Daten bearbeiten (Word opens Excel for it; RaabeWrite edits the table in a dialog).
// A chart keeps its chart part (drawn by the chart engine shared with RaabeCalc) and an embedded workbook with the
// data, written as Word writes them, so Word can edit the data of a chart made here.

import { h, messageBox, showDialog } from "@raabe/shell";
import { newChart, type ChartKind, type ChartModel } from "@raabe/charts";
import { chartFromXml, chartImage, chartToXml, newWorkbook, writeXlsx } from "@raabe/ooxml";
import type { InlineObject } from "@raabe/text-engine";
import type { AppApi } from "./api";

/** The table behind a chart: categories down, series across (as Word's data sheet). */
export interface ChartTable { categories: string[]; series: { name: string; values: (number | null)[] }[] }

/** Word's sample data for a new chart. */
export const SAMPLE: ChartTable = {
  categories: ["Kategorie 1", "Kategorie 2", "Kategorie 3", "Kategorie 4"],
  series: [{ name: "Datenreihe 1", values: [4.3, 2.5, 3.5, 4.5] }, { name: "Datenreihe 2", values: [2.4, 4.4, 1.8, 2.8] }, { name: "Datenreihe 3", values: [2, 2, 3, 5] }],
};

/** Diagramm einfügen: Word's groups and their first types, with German names. */
const TYPES: { group: string; items: { kind: ChartKind; label: string }[] }[] = [
  { group: "Säule", items: [{ kind: "column", label: "Gruppierte Säulen" }, { kind: "columnStacked", label: "Gestapelte Säulen" }] },
  { group: "Linie", items: [{ kind: "line", label: "Linie" }, { kind: "lineMarkers", label: "Linie mit Datenpunkten" }] },
  { group: "Kreis", items: [{ kind: "pie", label: "Kreis" }, { kind: "doughnut", label: "Ring" }] },
  { group: "Balken", items: [{ kind: "bar", label: "Gruppierte Balken" }, { kind: "barStacked", label: "Gestapelte Balken" }] },
  { group: "Fläche", items: [{ kind: "area", label: "Fläche" }] },
  { group: "Punkt (X Y)", items: [{ kind: "scatter", label: "Punkt (X Y)" }] },
];

const col = (i: number) => String.fromCharCode(65 + i);

/** The chart model for a table (references into the embedded workbook's Tabelle1, as Word writes them). */
export function chartModelOf(kind: ChartKind, t: ChartTable): ChartModel {
  const n = t.categories.length;
  const series = t.series.map((s, i) => ({
    name: { ref: `Tabelle1!$${col(i + 1)}$1`, text: s.name },
    cat: { ref: `Tabelle1!$A$2:$A$${n + 1}`, cache: [...t.categories] },
    val: { ref: `Tabelle1!$${col(i + 1)}$2:$${col(i + 1)}$${n + 1}`, cache: s.values.slice(0, n) },
  }));
  const anchor = { col: 0, row: 0, x: 0, y: 0 };
  return newChart(kind, series, anchor, anchor);
}

/** The embedded workbook with the table in Tabelle1 (A1 empty, names in row 1, categories in column A). */
export function workbookOf(t: ChartTable): Uint8Array {
  const wb = newWorkbook();
  const sh = wb.sheets[0];
  t.series.forEach((s, i) => sh.set(0, i + 1, { v: s.name }));
  t.categories.forEach((c, r) => {
    sh.set(r + 1, 0, { v: c });
    t.series.forEach((s, i) => { const v = s.values[r]; if (v !== null && v !== undefined) sh.set(r + 1, i + 1, { v }); });
  });
  return writeXlsx(wb);
}

/** The table a chart part holds (its series' cached names, categories and values). */
export function tableOf(xml: string): ChartTable {
  const m = chartFromXml(xml);
  const cats = (m.series[0]?.cat?.cache ?? []).map((c) => (c === null || c === undefined ? "" : String(c)));
  const n = Math.max(cats.length, ...m.series.map((s) => s.val.cache.length));
  return {
    categories: Array.from({ length: n }, (_, i) => cats[i] ?? `Kategorie ${i + 1}`),
    series: m.series.map((s, i) => ({ name: s.name?.text ?? `Datenreihe ${i + 1}`, values: Array.from({ length: n }, (_, k) => s.val.cache[k] ?? null) })),
  };
}

export function chartsFeature(api: AppApi) {
  const c = api.commands;
  const ed = () => api.editor();
  const accents = (): string[] | undefined => {
    const colors = (ed() as unknown as { doc?: { styles?: { theme?: { colors?: Record<string, string> } } } })?.doc?.styles?.theme?.colors;
    return colors ? ["accent1", "accent2", "accent3", "accent4", "accent5", "accent6"].map((k) => "#" + (colors[k] ?? "000000").replace("#", "")) : undefined;
  };

  /** A chart object for the text: the chart part, the workbook, the picture shown. Word's default size 432 x 252 pt. */
  function chartObject(kind: ChartKind, t: ChartTable, width = 432, height = 252, keepXml?: string): InlineObject {
    let model = chartModelOf(kind, t);
    if (keepXml) {
      // Daten bearbeiten keeps the chart's own look: only the series' data change.
      const old = chartFromXml(keepXml);
      model = { ...old, series: model.series.map((s, i) => ({ ...(old.series[i] ?? {}), ...s })) };
    }
    const xml = chartToXml(model);
    return { kind: "image", data: undefined, display: "", width, height, image: chartImage(model, width, height, accents()), chart: { xml, workbook: workbookOf(t) } };
  }

  /** Daten bearbeiten: the table in a grid; rows and series can be added and removed. */
  async function editTable(t: ChartTable): Promise<ChartTable | null> {
    let cats = [...t.categories];
    let series = t.series.map((s) => ({ name: s.name, values: [...s.values] }));
    const box = h("div", { class: "chart-data" });
    const draw = () => {
      const head = h("tr", null, h("th"), ...series.map((s, i) => h("th", null, field(s.name, (v) => { series[i].name = v; }))));
      const rows = cats.map((cat, r) => h("tr", null, h("th", null, field(cat, (v) => { cats[r] = v; })),
        ...series.map((s, i) => h("td", null, field(s.values[r] === null || s.values[r] === undefined ? "" : String(s.values[r]).replace(".", ","), (v) => {
          const n = Number(v.replace(/\./g, "").replace(",", "."));
          series[i].values[r] = v.trim() === "" || Number.isNaN(n) ? null : n;
        })))));
      box.replaceChildren(h("table", { class: "chart-data-table" }, head, ...rows),
        h("div", { class: "chart-data-buttons" },
          h("button", { type: "button", class: "btn", onclick: () => { cats = [...cats, `Kategorie ${cats.length + 1}`]; series.forEach((s) => s.values.push(null)); draw(); } }, "Zeile hinzufügen"),
          h("button", { type: "button", class: "btn", disabled: cats.length < 2, onclick: () => { cats = cats.slice(0, -1); series.forEach((s) => s.values.pop()); draw(); } }, "Zeile entfernen"),
          h("button", { type: "button", class: "btn", onclick: () => { series = [...series, { name: `Datenreihe ${series.length + 1}`, values: cats.map(() => null) }]; draw(); } }, "Datenreihe hinzufügen"),
          h("button", { type: "button", class: "btn", disabled: series.length < 2, onclick: () => { series = series.slice(0, -1); draw(); } }, "Datenreihe entfernen")));
    };
    const field = (value: string, set: (v: string) => void) => {
      const input = h("input", { type: "text", value, style: "width:90px" }) as HTMLInputElement;
      input.addEventListener("input", () => set(input.value));
      return input;
    };
    draw();
    const ok = await showDialog({ title: "Diagramm in RaabeWrite – Daten bearbeiten", body: box, buttons: [{ label: "OK", value: true, primary: true }, { label: "Abbrechen", value: false }], cancel: false });
    return ok ? { categories: cats, series } : null;
  }

  /** Diagramm einfügen: groups on the left, types as tiles with a preview of Word's sample data. */
  async function insertDialog(): Promise<void> {
    const e = ed();
    if (!e) return;
    let chosen: ChartKind = "column";
    const tiles: HTMLElement[] = [];
    const preview = (kind: ChartKind) => {
      const img = chartImage(chartModelOf(kind, SAMPLE), 150, 100);
      return h("img", { src: URL.createObjectURL(new Blob([img.bytes as BlobPart], { type: img.type })), width: 200, height: 133, alt: "" });
    };
    const list = h("div", { class: "chart-types" }, ...TYPES.map((g) => h("div", { class: "chart-type-group" }, h("div", { class: "palette-title" }, g.group),
      h("div", { class: "chart-type-tiles" }, ...g.items.map((it) => {
        const b = h("button", { type: "button", class: "chart-type" + (it.kind === chosen ? " on" : ""), "data-tip": it.label, "aria-label": it.label,
          onclick: () => { chosen = it.kind; tiles.forEach((x) => x.classList.toggle("on", x === b)); } }, preview(it.kind), h("span", null, it.label));
        tiles.push(b);
        return b;
      })))));
    const ok = await showDialog({ title: "Diagramm einfügen", body: list, buttons: [{ label: "OK", value: true, primary: true }, { label: "Abbrechen", value: false }], cancel: false });
    if (!ok) return;
    const table = (await editTable(SAMPLE)) ?? SAMPLE;
    e.insertInlineObject(chartObject(chosen, table));
    api.refresh();
  }

  /** Daten bearbeiten for the selected chart. */
  async function editData(): Promise<void> {
    const e = ed();
    const sel = e?.selectedImage();
    if (!e || !sel || !sel.obj.chart) { await messageBox("RaabeWrite", "Markieren Sie zuerst ein Diagramm."); return; }
    const obj = sel.obj;
    const xml = obj.chart!.xml;
    const t = await editTable(tableOf(xml));
    if (!t) return;
    const kind = chartFromXml(xml).kind;
    // Inline or floating (D-131): the chart's part and picture change, its place stays.
    const next = chartObject(kind, t, obj.width, obj.height, xml);
    e.setSelectedChart(next.chart!, next.image!);
    api.refresh();
  }

  const selectedChart = () => { const s = ed()?.selectedImage(); return !!s && !!s.obj.chart; };
  c.implement("insert.chart", { run: () => void insertDialog(), enabled: () => !!ed() });
  c.implement("chart.editData", { run: () => void editData(), enabled: selectedChart });
  /** Without dialogs (tests): a chart of `kind` with `table` at the caret; `edit` changes the selected chart's data. */
  const insertChart = (kind: ChartKind, table: ChartTable = SAMPLE) => { ed()?.insertInlineObject(chartObject(kind, table)); api.refresh(); };
  const setData = (table: ChartTable) => {
    const e = ed(), sel = e?.selectedImage();
    if (!e || !sel || !sel.obj.chart) return false;
    const xml = sel.obj.chart.xml;
    const next = chartObject(chartFromXml(xml).kind, table, sel.obj.width, sel.obj.height, xml);
    e.setSelectedChart(next.chart!, next.image!);
    return true;
  };
  return { insertDialog, editData, selectedChart, insertChart, setData };
}
