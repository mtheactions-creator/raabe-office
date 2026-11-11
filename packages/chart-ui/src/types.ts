// Einfügen > Diagramme: Excel's chart menus and the Diagramm einfügen / Diagrammtyp ändern dialog
// with "Alle Diagramme" (D-086). Groups, names and order as Excel 365 shows them
// (results/2026-11-05-excel-inventory.json); the map (Karte) is a cloud feature and stays disabled.

import { KINDS, newChart, paletteColor, renderChart, schemeColor, type ChartKind, type ChartModel } from "@raabe/charts";
import { h, pictureMenu, showDialog, type MenuEntry } from "@raabe/shell";

/** A Verbund variant of the Kombi menu: the series' types after the first group's. */
export type Combo = "columnLine" | "columnLineSecondary" | "areaColumn" | "custom";

export interface ChartTypeHost {
  insert(kind: ChartKind, tweak?: (m: ChartModel) => void): void;
  /** The selected chart (Diagrammtyp ändern), or null. */
  selected(): ChartModel | null;
  change(fn: (m: ChartModel) => void): void;
  /** The series a new chart of the kind would get from the selection (for the dialog's preview). */
  previewSeries(kind: ChartKind): ChartModel["series"];
}

type Group = { heading: string; kinds: ChartKind[] };

/** The menus of Einfügen > Diagramme, as Excel groups them. */
export const MENUS: Record<"column" | "line" | "pie" | "hierarchy" | "stat" | "scatter" | "other", { groups: Group[]; more: string; category: string }> = {
  column: { more: "Weitere Säulendiagramme...", category: "Säule", groups: [
    { heading: "2D-Säule", kinds: ["column", "columnStacked", "columnStacked100"] },
    { heading: "3D-Säule", kinds: ["column3DClustered", "column3DStacked", "column3DStacked100", "column3D"] },
    { heading: "2D-Balken", kinds: ["bar", "barStacked", "barStacked100"] },
    { heading: "3D-Balken", kinds: ["bar3DClustered", "bar3DStacked", "bar3DStacked100"] },
  ] },
  line: { more: "Weitere Liniendiagramme...", category: "Linie", groups: [
    { heading: "2D-Linie", kinds: ["line", "lineStacked", "lineStacked100", "lineMarkers", "lineMarkersStacked", "lineMarkersStacked100"] },
    { heading: "3D-Linie", kinds: ["line3D"] },
    { heading: "2D-Fläche", kinds: ["area", "areaStacked", "areaStacked100"] },
    { heading: "3D-Fläche", kinds: ["area3D", "area3DStacked", "area3DStacked100"] },
  ] },
  pie: { more: "Weitere Kreisdiagramme...", category: "Kreis", groups: [
    { heading: "2D-Kreis", kinds: ["pie", "pieOfPie", "barOfPie"] },
    { heading: "3D-Kreis", kinds: ["pie3D"] },
    { heading: "Ring", kinds: ["doughnut"] },
  ] },
  hierarchy: { more: "Weitere Hierarchiediagramme...", category: "Treemap", groups: [
    { heading: "Treemap", kinds: ["treemap"] },
    { heading: "Sunburst", kinds: ["sunburst"] },
  ] },
  stat: { more: "Weitere Statistikdiagramme...", category: "Histogramm", groups: [
    { heading: "Histogramm", kinds: ["histogram", "pareto"] },
    { heading: "Kastengrafik", kinds: ["boxWhisker"] },
  ] },
  scatter: { more: "Weitere Punktdiagramme (XY)...", category: "Punkt (XY)", groups: [
    { heading: "Punkt (XY)", kinds: ["scatter", "scatterSmoothMarkers", "scatterSmooth", "scatterLines", "scatterLinesNoMarkers"] },
    { heading: "Blase", kinds: ["bubble", "bubble3D"] },
  ] },
  other: { more: "Weitere Kursdiagramme...", category: "Kurs", groups: [
    { heading: "Wasserfall", kinds: ["waterfall"] },
    { heading: "Trichter", kinds: ["funnel"] },
    { heading: "Kurs", kinds: ["stockHLC", "stockOHLC", "stockVHLC", "stockVOHLC"] },
    { heading: "Oberfläche", kinds: ["surface3D", "surfaceWireframe", "surfaceContour", "surfaceContourWireframe"] },
    { heading: "Netz", kinds: ["radar", "radarMarkers", "radarFilled"] },
  ] },
};

/** Diagrammtyp ändern > Alle Diagramme: the categories at the left and their types in order. */
export const CATEGORIES: { name: string; kinds: ChartKind[]; combo?: boolean; disabled?: boolean }[] = [
  { name: "Säule", kinds: ["column", "columnStacked", "columnStacked100", "column3DClustered", "column3DStacked", "column3DStacked100", "column3D"] },
  { name: "Linie", kinds: ["line", "lineStacked", "lineStacked100", "lineMarkers", "lineMarkersStacked", "lineMarkersStacked100", "line3D"] },
  { name: "Kreis", kinds: ["pie", "pie3D", "pieOfPie", "barOfPie", "doughnut"] },
  { name: "Balken", kinds: ["bar", "barStacked", "barStacked100", "bar3DClustered", "bar3DStacked", "bar3DStacked100"] },
  { name: "Fläche", kinds: ["area", "areaStacked", "areaStacked100", "area3D", "area3DStacked", "area3DStacked100"] },
  { name: "Punkt (XY)", kinds: ["scatter", "scatterSmoothMarkers", "scatterSmooth", "scatterLines", "scatterLinesNoMarkers", "bubble", "bubble3D"] },
  { name: "Karte", kinds: [], disabled: true },
  { name: "Kurs", kinds: ["stockHLC", "stockOHLC", "stockVHLC", "stockVOHLC"] },
  { name: "Oberfläche", kinds: ["surface3D", "surfaceWireframe", "surfaceContour", "surfaceContourWireframe"] },
  { name: "Netz", kinds: ["radar", "radarMarkers", "radarFilled"] },
  { name: "Treemap", kinds: ["treemap"] },
  { name: "Sunburst", kinds: ["sunburst"] },
  { name: "Histogramm", kinds: ["histogram", "pareto"] },
  { name: "Kastengrafik", kinds: ["boxWhisker"] },
  { name: "Wasserfall", kinds: ["waterfall"] },
  { name: "Trichter", kinds: ["funnel"] },
  { name: "Kombi", kinds: [], combo: true },
];

export const COMBOS: [Combo, string][] = [
  ["columnLine", "Gruppierte Säulen/Linien"],
  ["columnLineSecondary", "Gruppierte Säulen/Linien auf der Sekundärachse"],
  ["areaColumn", "Gestapelte Flächen/Gruppierte Säulen"],
  ["custom", "Benutzerdefinierte Kombination"],
];

/** The types a series can take in a custom Verbund chart (Excel's list in the Kombi page). */
const COMBO_KINDS: ChartKind[] = ["column", "columnStacked", "columnStacked100", "bar", "barStacked", "barStacked100", "line", "lineStacked", "lineStacked100",
  "lineMarkers", "lineMarkersStacked", "lineMarkersStacked100", "area", "areaStacked", "areaStacked100", "scatter", "scatterSmoothMarkers", "scatterSmooth",
  "scatterLines", "scatterLinesNoMarkers", "radar", "radarMarkers", "radarFilled"];

/** A Verbund chart: Excel's presets turn the last series into a line (or the others into stacked areas). */
export function applyCombo(m: ChartModel, combo: Combo): void {
  const last = m.series.length - 1;
  if (last < 0) return;
  if (combo === "areaColumn") {
    m.kind = "areaStacked";
    m.series.forEach((s, i) => { delete s.kind; delete s.secondary; if (i === last && last > 0) s.kind = "column"; });
  } else {
    m.kind = "column";
    m.series.forEach((s, i) => { delete s.kind; delete s.secondary; if (i === last && last > 0) { s.kind = "line"; if (combo === "columnLineSecondary") s.secondary = true; } });
  }
  if (!m.series.some((s) => s.kind || s.secondary) && combo !== "custom") return;
  delete m.raw;
}

/** The ribbon's chart menus: each type inserts its chart; "Weitere …" opens the dialog at its category. */
/**
 * A chart type's picture for the Diagramme dropdowns (D-163): the type drawn by our own renderer with sample data,
 * without title, legend, axis labels and gridlines, as Excel's gallery shows a small picture per type.
 */
export function kindPicture(kind: ChartKind, w = 40, hgt = 40, change?: (m: ChartModel) => void): SVGSVGElement {
  const s = (n: string, v: number[]) => ({ name: { text: n }, cat: { cache: ["a", "b", "c", "d"] }, val: { cache: v } });
  const at = { col: 0, row: 0, x: 0, y: 0 };
  let m: ChartModel;
  try {
    m = newChart(kind, [s("1", [3, 5, 4, 6]), s("2", [2, 3, 5, 4]), s("3", [4, 2, 3, 5])], at, at);
    m.title.show = false;
    m.legend = "none";
    m.gridlines = false;
    m.dataLabels = false;
    m.format = { ...m.format, valAxis: { ...m.format?.valAxis, labelPos: "none", majorTick: "none" }, catAxis: { ...m.format?.catAxis, labelPos: "none", majorTick: "none" } };
    change?.(m);
  } catch {
    return document.createElementNS("http://www.w3.org/2000/svg", "svg");
  }
  const big = renderChart(m, w * 2.2, hgt * 2.2);
  big.setAttribute("viewBox", `0 0 ${w * 2.2} ${hgt * 2.2}`);
  big.setAttribute("width", String(w));
  big.setAttribute("height", String(hgt));
  return big;
}

export function chartMenus(host: ChartTypeHost): Record<string, MenuEntry[]> {
  const menu = (key: keyof typeof MENUS): MenuEntry[] => {
    const d = MENUS[key];
    // Excel's Diagramme dropdowns are galleries of pictures, the names are tooltips (D-163).
    return [
      ...d.groups.flatMap((g): MenuEntry[] => [{ heading: g.heading }, { block: () => pictureMenu(g.kinds.map((k) => ({ value: k, label: KINDS[k].label, svg: "", node: () => kindPicture(k) })), 4, [40, 40], undefined, (k) => host.insert(k as ChartKind)) }]),
      "-", { label: d.more, run: () => void chartTypeDialog(host, "insert", d.category) },
    ];
  };
  return {
    chartColumn: menu("column"), chartLine: menu("line"), chartPie: menu("pie"), chartHierarchy: menu("hierarchy"),
    chartStat: menu("stat"), chartScatter: menu("scatter"), chartOther: menu("other"),
    chartCombo: [{ heading: "Kombi" }, { block: () => pictureMenu(COMBOS.slice(0, 3).map(([c, label]) => ({ value: c, label, svg: "", node: () => kindPicture(c === "areaColumn" ? "areaStacked" : "column", 40, 40, (m) => applyCombo(m, c)) })), 4, [40, 40], undefined,
      (c) => host.insert(c === "areaColumn" ? "areaStacked" : "column", (m) => applyCombo(m, c as Combo))) },
      "-", { label: "Benutzerdefiniertes Kombinationsdiagramm erstellen...", run: () => void chartTypeDialog(host, "insert", "Kombi") }],
  };
}

const CSS = `
.ct-dlg { display: flex; flex-direction: column; gap: 8px; width: 760px; }
.ct-tabs { display: flex; gap: 18px; border-bottom: 1px solid #d0d0d0; }
.ct-tabs button { border: 0; background: none; padding: 4px 2px 6px; font: inherit; cursor: default; border-bottom: 2px solid transparent; }
.ct-tabs button.on { border-bottom-color: #217346; font-weight: 600; }
.ct-main { display: flex; gap: 10px; height: 420px; }
.ct-cats { width: 150px; overflow: auto; border-right: 1px solid #e0e0e0; }
.ct-cats div { padding: 5px 8px; cursor: default; }
.ct-cats div.on { background: #d4e8dc; }
.ct-cats div.off { color: #a0a0a0; }
.ct-right { flex: 1; display: flex; flex-direction: column; gap: 8px; min-width: 0; }
.ct-kinds { display: flex; flex-wrap: wrap; gap: 4px; }
.ct-kinds button { font: inherit; font-size: 11px; padding: 3px 6px; border: 1px solid #c8c8c8; background: #fff; cursor: default; }
.ct-kinds button.on { border-color: #217346; background: #e6f2eb; }
.ct-name { font-weight: 600; }
.ct-preview { border: 1px solid #e0e0e0; align-self: flex-start; }
.ct-combo { border-collapse: collapse; }
.ct-combo td, .ct-combo th { padding: 2px 6px; text-align: left; font-weight: normal; }
`;

/**
 * Diagramm einfügen (Weitere …) or Diagrammtyp ändern: Empfohlene Diagramme and Alle Diagramme, the
 * types of the chosen category, a preview of the data in the chosen type, OK.
 */
export async function chartTypeDialog(host: ChartTypeHost, mode: "insert" | "change", category?: string): Promise<void> {
  if (!document.getElementById("raabe-charttypes-css")) document.head.append(h("style", { id: "raabe-charttypes-css" }, CSS));
  const current = mode === "change" ? host.selected() : null;
  if (mode === "change" && !current) return;
  const recommended: ChartKind[] = ["column", "pie", "bar", "funnel"];
  let tab: "rec" | "all" = category || mode === "change" ? "all" : "rec";
  let cat = CATEGORIES.find((c) => c.name === category) ?? CATEGORIES.find((c) => current && (c.kinds.includes(current.kind) || (c.combo && current.series.some((s) => s.kind || s.secondary)))) ?? CATEGORIES[0];
  let kind: ChartKind = current?.kind && cat.kinds.includes(current.kind) ? current.kind : cat.kinds[0] ?? "column";
  let combo: Combo = "columnLine";
  // Custom Verbund: each series' type and axis.
  const seriesOf = () => (current ? current.series : host.previewSeries("column"));
  let custom = seriesOf().map((s, i, all) => ({ name: s.name?.text || `Datenreihe${i + 1}`, kind: (i === all.length - 1 && all.length > 1 ? "line" : "column") as ChartKind, secondary: false }));
  const tabs = h("div", { class: "ct-tabs", role: "tablist" });
  const main = h("div", { class: "ct-main" });
  const model = (): ChartModel => {
    const k = cat.combo ? (combo === "areaColumn" ? "areaStacked" : combo === "custom" ? custom[0]?.kind ?? "column" : "column") : kind;
    const base = current ? structuredClone({ ...current, raw: undefined }) : newChart(k, host.previewSeries(k), { col: 0, row: 0, x: 0, y: 0 }, { col: 8, row: 15, x: 0, y: 0 });
    base.kind = k;
    if (current && !cat.combo) for (const s of base.series) { delete s.kind; delete s.secondary; }
    if (cat.combo && combo !== "custom") applyCombo(base, combo);
    if (cat.combo && combo === "custom") base.series.forEach((s, i) => { s.kind = custom[i]?.kind === k ? undefined : custom[i]?.kind; s.secondary = custom[i]?.secondary || undefined; });
    return base;
  };
  const draw = () => {
    tabs.replaceChildren(...([["rec", "Empfohlene Diagramme"], ["all", "Alle Diagramme"]] as const).map(([t, label]) =>
      h("button", { type: "button", role: "tab", class: tab === t ? "on" : "", "aria-selected": String(tab === t), onclick: () => { tab = t; if (t === "rec") { cat = CATEGORIES[0]; kind = recommended[0]; } draw(); } }, label)));
    const left = h("div", { class: "ct-cats", role: "listbox", "aria-label": "Diagrammtypen" });
    const right = h("div", { class: "ct-right" });
    if (tab === "rec") {
      for (const k of recommended) left.append(h("div", { role: "option", class: kind === k ? "on" : "", onclick: () => { kind = k; cat = CATEGORIES.find((c) => c.kinds.includes(k))!; draw(); } }, KINDS[k].label));
    } else {
      for (const c of CATEGORIES) {
        left.append(h("div", { role: "option", class: c === cat ? "on" : c.disabled ? "off" : "", "aria-disabled": c.disabled ? "true" : undefined,
          onclick: () => { if (c.disabled) return; cat = c; kind = c.kinds[0] ?? "column"; draw(); } }, c.name));
      }
      const tiles = h("div", { class: "ct-kinds" });
      if (cat.combo) for (const [cb, label] of COMBOS) tiles.append(h("button", { type: "button", class: combo === cb ? "on" : "", onclick: () => { combo = cb; draw(); } }, label));
      else for (const k of cat.kinds) tiles.append(h("button", { type: "button", class: kind === k ? "on" : "", onclick: () => { kind = k; draw(); } }, KINDS[k].label));
      right.append(tiles);
    }
    right.append(h("div", { class: "ct-name" }, cat.combo ? COMBOS.find((c) => c[0] === combo)![1] : KINDS[kind].label));
    const m = model();
    try {
      const svg = renderChart(m, 440, 264);
      svg.classList.add("ct-preview");
      right.append(svg);
    } catch { /* a preview that cannot be drawn is left out */ }
    if (tab === "all" && cat.combo && combo === "custom") {
      const table = h("table", { class: "ct-combo" }, h("tr", null, h("th", null, "Datenreihenname"), h("th", null, "Diagrammtyp"), h("th", null, "Sekundärachse")));
      custom.forEach((row) => {
        const sel = h("select", null, ...COMBO_KINDS.map((k) => h("option", { value: k, selected: k === row.kind }, KINDS[k].label))) as HTMLSelectElement;
        sel.addEventListener("change", () => { row.kind = sel.value as ChartKind; draw(); });
        const box = h("input", { type: "checkbox" }) as HTMLInputElement;
        box.checked = row.secondary;
        box.addEventListener("change", () => { row.secondary = box.checked; draw(); });
        table.append(h("tr", null, h("td", null, row.name), h("td", null, sel), h("td", null, box)));
      });
      right.append(table);
    }
    main.replaceChildren(left, right);
  };
  draw();
  const ok = await showDialog({ title: mode === "change" ? "Diagrammtyp ändern" : "Diagramm einfügen", body: h("div", { class: "ct-dlg" }, tabs, main),
    buttons: [{ label: "OK", value: true, primary: true }, { label: "Abbrechen", value: false }], cancel: false });
  if (!ok) return;
  const chosen = model();
  if (mode === "insert") {
    host.insert(chosen.kind, (m) => { m.series.forEach((s, i) => { const c = chosen.series[i]; if (c?.kind) s.kind = c.kind; if (c?.secondary) s.secondary = true; }); });
  } else {
    host.change((m) => {
      m.kind = chosen.kind;
      m.series.forEach((s, i) => { const c = chosen.series[i]; if (c?.kind) s.kind = c.kind; else delete s.kind; if (c?.secondary) s.secondary = true; else delete s.secondary; });
    });
  }
}

/**
 * Farben ändern's picture of a palette (D-164): six squares in its colours, as Excel's and PowerPoint's menus show them
 * (id 10: the theme's accents, 11-26 the colourful and monochrome schemes).
 */
export function paletteSwatches(id: number, accents: string[]): HTMLElement {
  const theme: Record<string, string> = { dk1: "#000000", lt1: "#FFFFFF" };
  accents.forEach((a, k) => { theme[`accent${k + 1}`] = a; });
  const probe = { colors: 10 } as unknown as ChartModel;
  return h("span", { class: "palette-swatches", style: "display:inline-flex;gap:2px" }, ...Array.from({ length: 6 }, (_, i) =>
    h("span", { style: `display:inline-block;width:14px;height:14px;background:${id === 10 ? paletteColor(probe, i, 6, accents) : schemeColor(id, i, 6, theme)}` })));
}
