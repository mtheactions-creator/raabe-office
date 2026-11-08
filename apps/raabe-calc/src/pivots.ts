// PivotTables in RaabeCalc (Einfügen > PivotTable, D-066). The report is computed from its source
// range (sheet-engine's computePivot, equal to Excel's reports in pivot.test.ts) and written into the
// sheet's cells in Excel's default look; it is computed again when its fields change or with
// Aktualisieren (Alt+F5), not when the source changes (as in Excel). The "PivotTable-Felder" pane
// lists the fields and the four areas Filter, Spalten, Zeilen and Werte; fields are added by their
// box or dragged between areas. Cells of a report cannot be typed over (Excel refuses too).

import { computePivot, PIVOT_FUNCS, valueCaption, type PivotFunc, type PivotTable, type PivotValue, type Rect, type Sheet } from "@raabe/sheet-engine";
import { h, messageBox, showDialog, showMenu, type PaneSpec } from "@raabe/shell";
import type { CalcApp } from "./app";
import { cellName, parseRange, rangeName } from "@raabe/sheet-engine";

type Area = "filters" | "cols" | "rows" | "values";

const CSS = `
.pv-pane { display: flex; flex-direction: column; gap: 8px; padding: 0 16px 12px; font: 12px var(--font); }
.pv-fields { border: 1px solid #c8c6c4; max-height: 200px; overflow: auto; padding: 2px 0; }
.pv-fields label { display: flex; align-items: center; gap: 6px; padding: 2px 6px; }
.pv-fields label:hover { background: #f3f2f1; }
.pv-areas { display: grid; grid-template-columns: 1fr 1fr; gap: 8px; }
.pv-area { border: 1px solid #c8c6c4; min-height: 70px; padding: 2px; }
.pv-area-h { font-weight: 600; padding: 2px 4px; color: #444; }
.pv-chip { display: flex; justify-content: space-between; align-items: center; margin: 2px; padding: 2px 6px; background: #f3f2f1; border: 1px solid #e1dfdd; cursor: grab; }
.pv-chip:hover { border-color: #a19f9d; }
.pv-area.over { background: #f0f6ff; }
`;

export function pivotsFeature(app: CalcApp) {
  let paneFor: { sheet: Sheet; p: PivotTable } | null = null;

  function css(): void {
    if (!document.getElementById("raabe-pivot-css")) document.head.append(h("style", { id: "raabe-pivot-css" }, CSS));
  }

  /** The source records (first row: the field names). */
  function data(p: PivotTable): PivotValue[][] {
    const sheet = app.wb.sheets.find((s) => s.name === p.source.sheet);
    if (!sheet) return [];
    const out: PivotValue[][] = [];
    for (let r = p.source.r0; r <= p.source.r1; r++) {
      const row: PivotValue[] = [];
      for (let c = p.source.c0; c <= p.source.c1; c++) {
        const v = sheet.value(r, c);
        row.push(typeof v === "number" || typeof v === "string" || typeof v === "boolean" ? v : null);
      }
      out.push(row);
    }
    return out;
  }

  function fieldNames(p: PivotTable): string[] {
    return (data(p)[0] ?? []).map((v) => (v === null ? "" : String(v)));
  }

  /** Writes the report into the cells (inside an undo step): the old report's cells are emptied first. */
  function write(sheet: Sheet, p: PivotTable): void {
    if (p.area) for (let r = p.area.r0; r <= p.area.r1; r++) for (let c = p.area.c0; c <= p.area.c1; c++) sheet.set(r, c, undefined);
    const out = computePivot(p, data(p));
    if (!out.cells.length) {
      // An empty PivotTable: Excel's placeholder (own words in the box).
      const a = p.at;
      p.area = { r0: a.row, c0: a.col, r1: a.row + 17, c1: a.col + 2 };
      sheet.set(a.row, a.col, { v: p.name, s: style({ bold: true }) });
      sheet.set(a.row + 1, a.col, { v: "Felder im Bereich PivotTable-Felder wählen", s: style({}) });
      return;
    }
    let r0 = Infinity, c0 = Infinity, r1 = -Infinity, c1 = -Infinity;
    const bottom = Math.max(...out.cells.map((x) => x.row));
    for (const cell of out.cells) {
      r0 = Math.min(r0, cell.row); c0 = Math.min(c0, cell.col); r1 = Math.max(r1, cell.row); c1 = Math.max(c1, cell.col);
      const bold = cell.kind === "header" || cell.kind === "corner" || cell.kind === "subtotalLabel" || cell.kind === "subtotal" || cell.kind === "grandLabel" || cell.kind === "grand" || cell.kind === "colLabel";
      const head = cell.row === Math.min(...out.cells.filter((x) => x.kind !== "filter").map((x) => x.row)) + (p.cols.length ? 1 : 0) && cell.kind !== "filter";
      const grand = cell.row === bottom && p.colGrand !== false && p.rows.length > 0;
      sheet.set(cell.row, cell.col, { v: cell.v, s: style({ bold: bold && (cell.kind !== "subtotal" || p.subtotals !== "none"), indent: cell.kind === "rowLabel" && (p.layout ?? "compact") === "compact" ? cell.level ?? 0 : 0,
        headFill: head || grand, top: grand, bottom: head, numFmt: cell.numFmt }) });
    }
    // The cells between the labels (Excel leaves them empty) belong to the report as well.
    p.area = { r0, c0, r1, c1 };
    for (let r = r0; r <= r1; r++) for (let c = c0; c <= c1; c++) {
      if (!sheet.get(r, c) && (r === r1 || r === (out.cells.find((x) => x.kind !== "filter")?.row ?? r0) + (p.cols.length ? 1 : 0))) sheet.set(r, c, { v: null, s: style({ headFill: true, top: r === r1, bottom: r !== r1 }) });
    }
    // Columns as wide as the labels need (Excel fits a new report's columns).
    app.autofitCols([...Array(c1 - c0 + 1).keys()].map((k) => c0 + k));
  }

  /** PivotStyleLight16's look (header and total rows light blue with a blue line). */
  function style(o: { bold?: boolean; indent?: number; headFill?: boolean; top?: boolean; bottom?: boolean; numFmt?: string }): number | undefined {
    if (!o.bold && !o.indent && !o.headFill && !o.numFmt) return undefined;
    return app.wb.styles.derive(undefined, (s) => {
      if (o.bold) s.font = { ...s.font, bold: true };
      if (o.indent) { s.indent = o.indent; s.h = "left"; }
      if (o.headFill) s.fill = "#D9E1F2";
      if (o.top) s.borders = { ...s.borders, top: { style: "thin", color: "#8EA9DB" } };
      if (o.bottom) s.borders = { ...s.borders, bottom: { style: "thin", color: "#8EA9DB" } };
      if (o.numFmt) s.numFmt = o.numFmt;
    });
  }

  function change(label: string, sheet: Sheet, p0: PivotTable, fn: (p: PivotTable) => void): void {
    // After undo the sheet holds copies: work on the one of that name.
    const p = (sheet.pivots ?? []).find((x) => x.name === p0.name) ?? p0;
    app.change(label, () => {
      app.history.whole(sheet);
      fn(p);
      write(sheet, p);
    });
    if (paneFor?.p === p) openPane(sheet, p);
  }

  /** The PivotTable a cell belongs to. */
  function at(row: number, col: number, sheet = app.sheet): PivotTable | null {
    return (sheet.pivots ?? []).find((p) => p.area && row >= p.area.r0 && row <= p.area.r1 && col >= p.area.c0 && col <= p.area.c1) ?? null;
  }

  /** Einfügen > PivotTable: Excel's "PivotTable aus Tabelle oder Bereich". */
  async function create(): Promise<void> {
    if (app.grid.isEditing) app.grid.commitEdit("none");
    const a = app.sheet.view.active;
    const sel = app.grid.selection.at(-1)!;
    const region: Rect = sel.r0 === sel.r1 && sel.c0 === sel.c1 ? app.grid.currentRegion(a.row, a.col) : sel;
    const q = (n: string) => (/^[A-Za-z_][A-Za-z0-9_.]*$/.test(n) ? n : `'${n.replace(/'/g, "''")}'`);
    const range = h("input", { type: "text", style: "width:240px", value: `${q(app.sheet.name)}!${rangeName(region)}` }) as HTMLInputElement;
    const newSheet = h("input", { type: "radio", name: "pvwhere", checked: true }) as HTMLInputElement;
    const here = h("input", { type: "radio", name: "pvwhere" }) as HTMLInputElement;
    const place = h("input", { type: "text", style: "width:200px" }) as HTMLInputElement;
    const body = h("div", { style: "display:flex;flex-direction:column;gap:8px" },
      h("div", null, "Wählen Sie die Daten, die analysiert werden sollen:"), h("label", { class: "field" }, "Tabelle/Bereich:", range),
      h("div", null, "Wählen Sie, wo die PivotTable entstehen soll:"),
      h("label", null, newSheet, " Neues Arbeitsblatt"), h("label", null, here, " Vorhandenes Arbeitsblatt"), h("label", { class: "field" }, "Quelle:", place));
    const ok = await showDialog({ title: "PivotTable aus Tabelle oder Bereich", body, buttons: [{ label: "OK", value: true, primary: true }, { label: "Abbrechen", value: false }], cancel: false });
    if (!ok) return;
    const ref = range.value.replace(/^=/, "");
    const bang = ref.lastIndexOf("!");
    const sheetName = bang > 0 ? ref.slice(0, bang).replace(/^'(.*)'$/, "$1").replace(/''/g, "'") : app.sheet.name;
    const src = parseRange((bang > 0 ? ref.slice(bang + 1) : ref).replace(/\$/g, ""));
    const srcSheet = app.wb.sheets.find((s) => s.name === sheetName);
    if (!src || !srcSheet || src.r1 <= src.r0) { await messageBox("RaabeCalc", "Der Quellbereich ist ungültig. Er braucht eine Zeile mit Feldnamen und mindestens eine Zeile mit Daten."); return; }
    for (let c = src.c0; c <= src.c1; c++) {
      const v = srcSheet.value(src.r0, c);
      if (v === null || v === "") { await messageBox("RaabeCalc", "Jede Spalte des Bereichs braucht in der ersten Zeile einen Feldnamen."); return; }
    }
    const n = app.wb.sheets.reduce((k, s) => k + (s.pivots?.length ?? 0), 0) + 1;
    const p: PivotTable = { name: `PivotTable${n}`, source: { sheet: sheetName, ...src }, at: { row: 2, col: 0 }, rows: [], cols: [], values: [], filters: [] };
    let target: Sheet;
    if (here.checked) {
      const t = parseRange(place.value.replace(/^=|\$/g, "").replace(/^.*!/, ""));
      if (!t) { await messageBox("RaabeCalc", "Der Zielbereich ist ungültig."); return; }
      p.at = { row: t.r0, col: t.c0 };
      target = app.sheet;
    } else {
      app.addSheet(false);
      target = app.sheet;
    }
    app.change("PivotTable einfügen", () => {
      app.history.whole(target);
      target.pivots = [...(target.pivots ?? []), p];
      write(target, p);
    });
    app.grid.selectCell(p.at.row, p.at.col);
    openPane(target, p);
  }

  // ----- PivotTable-Felder ----------------------------------------------------------------------

  function openPane(sheet: Sheet, p: PivotTable): void {
    css();
    paneFor = { sheet, p };
    document.documentElement.style.setProperty("--pane-top", `${document.querySelector(".calc-main")?.getBoundingClientRect().top ?? 150}px`);
    app.taskPane.onClose = () => { paneFor = null; };
    const spec: PaneSpec = { title: "PivotTable-Felder", sets: [{ label: "", tabs: [{ id: "f", label: "Felder", icon: "table", sections: () => [{ title: "Felder", open: true, controls: () => [{ type: "custom", render: () => fieldList(sheet, p) }] }] }] }] };
    app.taskPane.show(spec, false);
  }

  function fieldList(sheet: Sheet, p: PivotTable): HTMLElement {
    const names = fieldNames(p);
    const sample = data(p).slice(1);
    const numeric = (f: number) => sample.some((r) => typeof r[f] === "number") && !sample.some((r) => typeof r[f] === "string" && r[f] !== "");
    const used = (f: number) => p.rows.includes(f) || p.cols.includes(f) || p.filters.some((x) => x.field === f) || p.values.some((v) => v.field === f);
    const search = h("input", { type: "search", placeholder: "Suchen", style: "width:100%" }) as HTMLInputElement;
    const list = h("div", { class: "pv-fields" });
    const fill = () => list.replaceChildren(...names.map((n, f) => {
      if (search.value && !n.toLowerCase().includes(search.value.toLowerCase())) return h("span");
      const box = h("input", { type: "checkbox", checked: used(f) }) as HTMLInputElement;
      box.addEventListener("change", () => change("PivotTable-Feld", sheet, p, (q) => {
        if (box.checked) {
          // Text fields go to Zeilen, numbers to Werte (Excel).
          if (numeric(f)) q.values.push({ field: f, func: "sum" }); else q.rows.push(f);
        } else remove(q, f);
      }));
      const lab = h("label", { draggable: "true" }, box, n);
      lab.addEventListener("dragstart", (e) => e.dataTransfer?.setData("text/plain", JSON.stringify({ f, from: "list" })));
      return lab;
    }));
    search.addEventListener("input", fill);
    fill();
    const area = (key: Area, title: string) => {
      const box = h("div", { class: "pv-area" }, h("div", { class: "pv-area-h" }, title));
      const items: { f: number; label: string; i: number }[] =
        key === "rows" ? p.rows.map((f, i) => ({ f, label: names[f], i })) : key === "cols" ? p.cols.map((f, i) => ({ f, label: names[f], i }))
          : key === "filters" ? p.filters.map((x, i) => ({ f: x.field, label: names[x.field], i })) : p.values.map((v, i) => ({ f: v.field, label: valueCaption(names, v, p.values), i }));
      if (key === "cols" && p.values.length > 1) items.push({ f: -1, label: "Σ Werte", i: items.length });
      for (const it of items) {
        const chip = h("div", { class: "pv-chip", draggable: "true" }, h("span", null, it.label), h("span", null, "˅"));
        chip.addEventListener("dragstart", (e) => e.dataTransfer?.setData("text/plain", JSON.stringify({ f: it.f, from: key, i: it.i })));
        chip.addEventListener("click", () => chipMenu(sheet, p, key, it.i, chip));
        box.append(chip);
      }
      box.addEventListener("dragover", (e) => { e.preventDefault(); box.classList.add("over"); });
      box.addEventListener("dragleave", () => box.classList.remove("over"));
      box.addEventListener("drop", (e) => {
        e.preventDefault();
        box.classList.remove("over");
        const d = JSON.parse(e.dataTransfer?.getData("text/plain") || "null") as { f: number; from: string; i?: number } | null;
        if (!d || d.f < 0) return;
        change("PivotTable-Feld", sheet, p, (q) => moveTo(q, d.f, d.from as Area | "list", key, d.i));
      });
      return box;
    };
    return h("div", { class: "pv-pane" }, h("div", null, "Felder für den Bericht:"), search, list,
      h("div", null, "Felder zwischen den Bereichen ziehen:"),
      h("div", { class: "pv-areas" }, area("filters", "Filter"), area("cols", "Spalten"), area("rows", "Zeilen"), area("values", "Σ Werte")));
  }

  function remove(q: PivotTable, f: number): void {
    q.rows = q.rows.filter((x) => x !== f);
    q.cols = q.cols.filter((x) => x !== f);
    q.filters = q.filters.filter((x) => x.field !== f);
    q.values = q.values.filter((x) => x.field !== f);
  }

  function moveTo(q: PivotTable, f: number, from: Area | "list", to: Area, i?: number): void {
    // From Werte only the value field itself moves; a field in a label area appears once.
    if (from === "values" && i !== undefined) q.values.splice(i, 1);
    else if (from !== "list") { q.rows = q.rows.filter((x) => x !== f); q.cols = q.cols.filter((x) => x !== f); q.filters = q.filters.filter((x) => x.field !== f); }
    if (to === "values") q.values.push({ field: f, func: "sum" });
    else {
      q.rows = q.rows.filter((x) => x !== f); q.cols = q.cols.filter((x) => x !== f); q.filters = q.filters.filter((x) => x.field !== f);
      if (to === "rows") q.rows.push(f); else if (to === "cols") q.cols.push(f); else q.filters.push({ field: f });
    }
  }

  function chipMenu(sheet: Sheet, p: PivotTable, key: Area, i: number, chip: HTMLElement): void {
    const list = key === "values" ? p.values : key === "filters" ? p.filters : p[key];
    const move = (d: number) => change("PivotTable-Feld", sheet, p, (q) => {
      const l = (key === "values" ? q.values : key === "filters" ? q.filters : q[key]) as unknown[];
      const j = Math.max(0, Math.min(l.length - 1, d === Infinity ? l.length - 1 : d === -Infinity ? 0 : i + d));
      const [x] = l.splice(i, 1);
      l.splice(j, 0, x);
    });
    const f = key === "values" ? p.values[i].field : key === "filters" ? p.filters[i].field : p[key][i];
    const to = (a: Area) => () => change("PivotTable-Feld", sheet, p, (q) => moveTo(q, f, key, a, i));
    showMenu(app.commands, [
      { label: "Nach oben", run: () => move(-1), disabled: i === 0 }, { label: "Nach unten", run: () => move(1), disabled: i === list.length - 1 },
      { label: "An den Anfang verschieben", run: () => move(-Infinity), disabled: i === 0 }, { label: "Ans Ende verschieben", run: () => move(Infinity), disabled: i === list.length - 1 }, "-",
      { label: "In Berichtsfilter verschieben", run: to("filters"), disabled: key === "filters" }, { label: "In Zeilenbeschriftung verschieben", run: to("rows"), disabled: key === "rows" },
      { label: "In Spaltenbeschriftung verschieben", run: to("cols"), disabled: key === "cols" }, { label: "In Werte verschieben", run: to("values"), disabled: key === "values" }, "-",
      { label: "Feld entfernen", run: () => change("PivotTable-Feld", sheet, p, (q) => { if (key === "values") q.values.splice(i, 1); else remove(q, f); }) },
      ...(key === "values" ? ["-" as const, { label: "Wertfeldeinstellungen...", run: () => void valueSettings(sheet, p, i) }] : []),
    ], chip.getBoundingClientRect());
  }

  /** Wertfeldeinstellungen: the function and a name of a value field. */
  async function valueSettings(sheet: Sheet, p: PivotTable, i: number): Promise<void> {
    const names = fieldNames(p);
    const v = p.values[i];
    const name = h("input", { type: "text", style: "width:240px", value: valueCaption(names, v, p.values) }) as HTMLInputElement;
    const list = h("select", { size: 8, style: "width:240px" }, ...PIVOT_FUNCS.map((x) => h("option", { value: x.f, selected: x.f === v.func }, x.label))) as HTMLSelectElement;
    const auto = valueCaption(names, v, p.values);
    list.addEventListener("change", () => {
      // The name follows the function while it is Excel's automatic one.
      if (name.value === auto || PIVOT_FUNCS.some((x) => name.value.startsWith(x.label + " von "))) name.value = valueCaption(names, { ...v, func: list.value as PivotFunc, name: undefined }, p.values.map((x) => (x === v ? { ...v, func: list.value as PivotFunc } : x)));
    });
    const ok = await showDialog({ title: "Wertfeldeinstellungen", body: h("div", { style: "display:flex;flex-direction:column;gap:8px" },
      h("div", null, `Quellenname: ${names[v.field]}`), h("label", { class: "field" }, "Benutzerdefinierter Name:", name),
      h("div", null, "Wertfeld zusammenfassen nach"), list),
      buttons: [{ label: "OK", value: true, primary: true }, { label: "Abbrechen", value: false }], cancel: false });
    if (!ok) return;
    change("Wertfeldeinstellungen", sheet, p, (q) => {
      const x = q.values[i];
      x.func = list.value as PivotFunc;
      const autoName = valueCaption(names, { ...x, name: undefined }, q.values);
      x.name = name.value && name.value !== autoName ? name.value : undefined;
    });
  }

  /** The PivotTable of the active cell (for the contextual tabs). */
  const current = (): PivotTable | null => at(app.sheet.view.active.row, app.sheet.view.active.col);

  return {
    create, at, current,
    /** The source records and writing the report (Datenschnitte filter through them, D-092). */
    data, write,
    /** A PivotTable given as a model (tests, and pasting a report's definition). */
    add: (sheet: Sheet, p: PivotTable) => app.change("PivotTable einfügen", () => { app.history.whole(sheet); sheet.pivots = [...(sheet.pivots ?? []), p]; write(sheet, p); }),
    /** Aktualisieren (Alt+F5) and Alle aktualisieren. */
    refresh: (all = false) => {
      const list = all ? app.wb.sheets.flatMap((s) => (s.pivots ?? []).map((p) => ({ s, p }))) : (() => { const p = current(); return p ? [{ s: app.sheet, p }] : []; })();
      if (!list.length) return;
      app.change("Aktualisieren", () => { for (const { s, p } of list) { app.history.whole(s); write(s, p); } });
    },
    fieldList: () => { const p = current(); if (!p) return; if (paneFor?.p === p && app.taskPane.isOpen) app.taskPane.close(); else openPane(app.sheet, p); },
    setLayout: (layout: PivotTable["layout"]) => { const p = current(); if (p) change("Berichtslayout", app.sheet, p, (q) => { q.layout = layout; q.subtotals = layout === "tabular" ? "bottom" : q.subtotals; }); },
    setSubtotals: (v: PivotTable["subtotals"]) => { const p = current(); if (p) change("Teilergebnisse", app.sheet, p, (q) => { q.subtotals = v; }); },
    setGrand: (rows: boolean, cols: boolean) => { const p = current(); if (p) change("Gesamtergebnisse", app.sheet, p, (q) => { q.colGrand = rows; q.rowGrand = cols; }); },
    remove: () => {
      const p = current();
      if (!p) return;
      const sheet = app.sheet;
      app.change("PivotTable löschen", () => {
        app.history.whole(sheet);
        if (p.area) for (let r = p.area.r0; r <= p.area.r1; r++) for (let c = p.area.c0; c <= p.area.c1; c++) sheet.set(r, c, undefined);
        sheet.pivots = (sheet.pivots ?? []).filter((x) => x !== p);
        if (!sheet.pivots.length) sheet.pivots = undefined;
      });
      if (paneFor?.p === p) app.taskPane.close();
    },
    /** A cell inside a report cannot be typed over (the message is our own words). */
    guard: (r: Rect): boolean => {
      const hit = (app.sheet.pivots ?? []).some((p) => p.area && !(r.r1 < p.area.r0 || r.r0 > p.area.r1 || r.c1 < p.area.c0 || r.c0 > p.area.c1));
      if (hit) void messageBox("RaabeCalc", "Zellen einer PivotTable lassen sich nicht direkt ändern. Ändern Sie die Felder im Bereich PivotTable-Felder oder die Quelldaten und aktualisieren Sie dann.");
      return !hit;
    },
    describe: (p: PivotTable) => `${p.name}: ${cellName(p.at.row, p.at.col)}`,
  };
}
