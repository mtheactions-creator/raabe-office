// Daten > Was-wäre-wenn-Analyse > Szenario-Manager (D-084), as German Excel 365 has it (recorded in
// our own invisible Excel: probe_excel_scenarios.py, results/2026-11-08-excel-scenarios.json):
// - a scenario is a name, the Veränderbare Zellen (up to 32, need not touch), their values as
//   typed (German Excel keeps "0,25" in the file), a comment ("Erstellt von <user> am <date>"),
//   and Schutz (Änderungen verhindern, Ausblenden), which counts while the sheet is protected;
// - Anzeigen enters the values as if typed (a formula among the cells is replaced), one undo step;
// - Zusammenfassung makes the Szenariobericht on a new sheet before the active one ("Szenariobericht",
//   "Szenariobericht 2" …): the cells (or their names), the current values and each scenario's values,
//   the Ergebniszellen as each scenario gives them (the other cells as they are), with Excel's look;
// - Zusammenführen copies the scenarios of another sheet; a name that exists gets the date added.
// The dialogs are Excel's (labels from Excel's own dialog resources); explanations are our own words.
// The PivotTable-Szenarioberichts is not built (greyed out).

import { h, messageBox, showDialog } from "@raabe/shell";
import {
  Sheet, cellName, colName, parseRange, type Calc, type CellStyle, type Rect, type Scenario, type Scenarios, type Value, type Workbook,
} from "@raabe/sheet-engine";
import type { CalcApp } from "./app";
import * as A from "./actions";
import { editText, readInput } from "./engine";

type Pt = { row: number; col: number };

/** Excel's limit of changing cells per scenario. */
export const MAX_CELLS = 32;

/** "B1:B2;D4" (German list separator, "," too, $ allowed) as cells, ranges row by row; null when invalid. */
export function parseCells(text: string): Pt[] | null {
  const parts = text.split(/[;,]/).map((t) => t.trim()).filter(Boolean);
  if (!parts.length) return null;
  const out: Pt[] = [];
  const seen = new Set<string>();
  for (const p of parts) {
    const r = parseRange(p.replace(/\$/g, "").replace(/^.*!/, ""));
    if (!r) return null;
    if ((r.r1 - r.r0 + 1) * (r.c1 - r.c0 + 1) > MAX_CELLS) return out.concat(Array(MAX_CELLS + 1).fill({ row: 0, col: 0 }));
    for (let row = r.r0; row <= r.r1; row++) for (let col = r.c0; col <= r.c1; col++) {
      const k = `${row},${col}`;
      if (!seen.has(k)) { seen.add(k); out.push({ row, col }); }
    }
  }
  return out;
}

const abs = (p: Pt) => `$${colName(p.col)}$${p.row + 1}`;

/** The cells as Excel's dialog shows them: blocks as ranges, absolute, ";" between ("$B$1:$B$2;$B$3"). */
export function cellsText(cells: Pt[]): string {
  const out: string[] = [];
  let i = 0;
  while (i < cells.length) {
    const a = cells[i];
    // A run along the row, then as many rows below with the same run as follow in order.
    let w = 1;
    while (i + w < cells.length && cells[i + w].row === a.row && cells[i + w].col === a.col + w) w++;
    let hgt = 1;
    for (;;) {
      const at = i + hgt * w;
      let ok = at + w <= cells.length;
      for (let k = 0; ok && k < w; k++) ok = cells[at + k].row === a.row + hgt && cells[at + k].col === a.col + k;
      if (!ok) break;
      hgt++;
    }
    out.push(w === 1 && hgt === 1 ? abs(a) : `${abs(a)}:${abs({ row: a.row + hgt - 1, col: a.col + w - 1 })}`);
    i += w * hgt;
  }
  return out.join(";");
}

const two = (n: number) => String(n).padStart(2, "0");
export const dateText = (d: Date) => `${two(d.getDate())}.${two(d.getMonth() + 1)}.${d.getFullYear()}`;
/** The comment Szenario hinzufügen starts with (recorded). */
export const createdComment = (user: string, d: Date) => `Erstellt von ${user} am ${dateText(d)}`;
/** Szenario bearbeiten adds a line saying who changed it. */
export const changedComment = (old: string, user: string, d: Date) => `${old ? old + "\n" : ""}Geändert von ${user} am ${dateText(d)}`;

/** Zusammenführen: a name that exists already gets the date (recorded: "Gut 08.11.2026"), then a number. */
export function mergedName(name: string, taken: string[], d: Date): string {
  const has = (n: string) => taken.some((t) => t.toLowerCase() === n.toLowerCase());
  if (!has(name)) return name;
  let n = `${name} ${dateText(d)}`;
  for (let i = 2; has(n); i++) n = `${name} ${dateText(d)} ${i}`;
  return n;
}

/** The new report's sheet name: Szenariobericht, Szenariobericht 2, … (recorded). */
export function reportName(wb: Workbook, base = "Szenariobericht"): string {
  if (!wb.sheetByName(base)) return base;
  for (let i = 2; ; i++) if (!wb.sheetByName(`${base} ${i}`)) return `${base} ${i}`;
}

/** What a changing cell holds now, as Szenariowerte proposes it (a formula's value, not the formula). */
export function currentText(wb: Workbook, sheet: Sheet, p: Pt): string {
  const c = sheet.get(p.row, p.col);
  if (!c || c.v === null) return "";
  return editText({ v: c.v }, wb.styles.get(c.s));
}

/** A name of the workbook that refers to exactly this cell (the report shows it instead of $B$1). */
export function cellLabel(wb: Workbook, sheet: Sheet, p: Pt): string {
  const si = wb.sheets.indexOf(sheet);
  const want = abs(p);
  for (const n of wb.names) {
    if (n.hidden || (n.sheet !== undefined && n.sheet !== si)) continue;
    const m = /^(?:'((?:[^']|'')+)'|([^!]+))!(\$?[A-Z]+\$?\d+)$/.exec(n.ref);
    if (!m) continue;
    const sh = (m[1] ?? m[2]).replace(/''/g, "'");
    if (sh.toLowerCase() === sheet.name.toLowerCase() && m[3] === want) return n.name;
  }
  return want;
}

/** Anzeigen: the scenario's values entered into its cells as if typed (inside an undo step). */
export function applyScenario(ctx: A.Ctx, sheet: Sheet, sc: Scenario): void {
  for (const c of sc.cells) A.enter(ctx, sheet, { r0: c.row, c0: c.col, r1: c.row, c1: c.col }, c.val);
}

/** A value typed into a cell with this style, as the cell would hold it. */
const valueOf = (wb: Workbook, sheet: Sheet, p: Pt, text: string): Value => {
  if (text === "") return null;
  const r = readInput(text, wb.styles.get(sheet.get(p.row, p.col)?.s));
  return r.f !== undefined ? text : r.v;
};

export interface ReportData {
  names: string[];
  comments: string[];
  changing: { label: string; current: Value; fmt: string; values: { v: Value; mine: boolean }[] }[];
  results: { label: string; current: Value; fmt: string; values: Value[] }[];
}

/**
 * The report's content: every scenario's cells in their first order, the current values, each
 * scenario's values (a cell it does not change keeps the current value), and the result cells as
 * each scenario alone gives them.
 */
export function reportData(wb: Workbook, sheet: Sheet, calc: Calc, results: Pt[]): ReportData {
  const si = wb.sheets.indexOf(sheet);
  const list = sheet.scenarios?.list ?? [];
  const cells: Pt[] = [];
  for (const sc of list) for (const c of sc.cells) if (!cells.some((x) => x.row === c.row && x.col === c.col)) cells.push({ row: c.row, col: c.col });
  const fmt = (p: Pt) => wb.styles.get(sheet.get(p.row, p.col)?.s).numFmt;
  const changing = cells.map((p) => ({
    label: cellLabel(wb, sheet, p), current: calc.value(si, p.row, p.col), fmt: fmt(p),
    values: list.map((sc) => {
      const own = sc.cells.find((c) => c.row === p.row && c.col === p.col);
      return own ? { v: valueOf(wb, sheet, p, own.val), mine: true } : { v: calc.value(si, p.row, p.col), mine: false };
    }),
  }));
  const res = results.map((p) => ({ label: cellLabel(wb, sheet, p), current: calc.value(si, p.row, p.col), fmt: fmt(p), values: [] as Value[] }));
  if (results.length) {
    for (const sc of list) {
      const saved = sc.cells.map((c) => sheet.get(c.row, c.col));
      for (const c of sc.cells) {
        const old = sheet.get(c.row, c.col);
        sheet.set(c.row, c.col, { ...(old ?? { v: null }), f: undefined, spill: undefined, array: undefined, v: valueOf(wb, sheet, c, c.val) });
      }
      calc.changed(si, sc.cells);
      res.forEach((r, i) => r.values.push(calc.value(si, results[i].row, results[i].col)));
      sc.cells.forEach((c, i) => sheet.set(c.row, c.col, saved[i]));
      calc.changed(si, sc.cells);
    }
  }
  return { names: list.map((s) => s.name), comments: list.map((s) => s.comment ?? ""), changing, results: res };
}

// The report's look (recorded): purple title rows with white text, grey label cells, navy section labels.
const PURPLE = "#800080", GREY = "#C0C0C0", WHITE = "#FFFFFF", NAVY = "#000080";

/** Our own words for the note below the report (Excel's text is its own, D-058), in three rows as Excel's. */
export const REPORT_NOTE = [
  "Hinweis: Unter „Aktuelle Werte“ stehen die Werte, die die veränderbaren Zellen",
  "beim Erstellen dieses Berichts hatten. Grau hinterlegt sind in jeder Szenariospalte",
  "die Zellen, die das Szenario verändert.",
];

/**
 * The Szenariobericht as a new sheet (not yet in the workbook), laid out as Excel's: column B and C
 * the labels, D the current values, E… one column per scenario; rows: title, names, comments
 * (grouped, hidden), Veränderbare Zellen, Ergebniszellen, the note. Comments are left to `comments`
 * so that the column widths can be fitted without them (they wrap).
 */
export function buildReport(wb: Workbook, data: ReportData, name: string): { sheet: Sheet; lastCol: number; comments: () => void } {
  const rep = new Sheet(name);
  const base = wb.styles.get(0);
  type Look = { bold?: boolean; size?: number; color?: string; fill?: string; h?: "left" | "right"; top?: "thin" | "medium"; bottom?: "thin" | "medium"; wrap?: boolean; v?: "top"; fmt?: string };
  const style = (l: Look): number => {
    const s: CellStyle = structuredClone(base);
    s.font = { ...s.font, ...(l.bold ? { bold: true } : {}), ...(l.size ? { size: l.size } : {}), ...(l.color ? { color: l.color } : {}) };
    delete s.font.theme;
    if (l.fill) { s.fill = l.fill; delete s.fillTheme; delete s.fillXml; }
    s.borders = { ...(l.top ? { top: { style: l.top } } : {}), ...(l.bottom ? { bottom: { style: l.bottom } } : {}) };
    if (l.h) s.h = l.h;
    if (l.wrap) s.wrap = true;
    if (l.v) s.v = l.v;
    if (l.fmt) s.numFmt = l.fmt;
    return wb.styles.id(s);
  };
  const B = 1, C = 2, D = 3, last = D + data.names.length;
  const put = (row: number, col: number, v: Value, l: Look) => rep.set(row, col, { v, s: style(l) });
  /** One row from B to the last column: the labels' look in B and C, the values' in D…. */
  const row = (r: number, label: Look, rest: Look, texts: Record<number, Value> = {}, each?: (c: number) => Look) => {
    for (let c = B; c <= last; c++) put(r, c, texts[c] ?? null, c <= C ? label : { ...rest, ...(each?.(c) ?? {}) });
  };
  for (let c = B; c <= last; c++) put(0, c, null, { bottom: "medium" });
  row(1, { bold: true, size: 12, color: WHITE, fill: PURPLE, h: "left", top: "medium" }, { size: 10, color: WHITE, fill: PURPLE, h: "right", top: "medium" }, { [B]: "Szenariobericht" });
  const names: Record<number, Value> = { [D]: "Aktuelle Werte:" };
  data.names.forEach((n, i) => (names[D + 1 + i] = n));
  row(2, { bold: true, size: 12, color: WHITE, fill: PURPLE, h: "left", bottom: "thin" }, { size: 10, color: WHITE, fill: PURPLE, h: "right", bottom: "thin" }, names);
  const thinBoth: Look = { top: "thin", bottom: "thin" };
  row(3, { bold: true, fill: GREY, h: "left", ...thinBoth }, thinBoth, {}, (c) => (c > D ? { size: 8, wrap: true, v: "top" } : {}));
  row(4, { bold: true, fill: GREY, color: NAVY, h: "left", ...thinBoth }, thinBoth, { [B]: "Veränderbare Zellen:" });
  let r = 5;
  data.changing.forEach((ch, i) => {
    const edge: Look = { ...(i === 0 ? { top: "thin" } : {}), ...(i === data.changing.length - 1 ? { bottom: "thin" } : {}) };
    put(r, B, null, { bold: true, fill: GREY, h: "left", ...edge });
    put(r, C, ch.label, { bold: true, fill: GREY, h: "left", ...edge });
    put(r, D, ch.current, { ...edge, fmt: ch.fmt });
    ch.values.forEach((x, k) => put(r, D + 1 + k, x.v, { ...edge, fmt: ch.fmt, ...(x.mine ? { fill: GREY } : {}) }));
    rep.rows.set(r, { outline: 1 });
    r++;
  });
  if (data.results.length) {
    row(r++, { bold: true, fill: GREY, color: NAVY, h: "left", ...thinBoth }, thinBoth, { [B]: "Ergebniszellen:" });
    data.results.forEach((res, i) => {
      const edge: Look = { ...(i === 0 ? { top: "thin" } : {}), ...(i === data.results.length - 1 ? { bottom: "medium" } : {}) };
      put(r, B, null, { bold: true, fill: GREY, h: "left", ...edge });
      put(r, C, res.label, { bold: true, fill: GREY, h: "left", ...edge });
      put(r, D, res.current, { ...edge, fmt: res.fmt });
      res.values.forEach((v, k) => put(r, D + 1 + k, v, { ...edge, fmt: res.fmt }));
      rep.rows.set(r, { outline: 1 });
      r++;
    });
  } else {
    // Without Ergebniszellen Excel closes the table with an empty row of the section look.
    row(r++, { bold: true, fill: GREY, color: NAVY, h: "left", top: "thin", bottom: "medium" }, { top: "thin", bottom: "medium" });
  }
  for (let c = B; c <= last; c++) put(r, c, c === B ? REPORT_NOTE[0] : null, { top: "medium" });
  put(r + 1, B, REPORT_NOTE[1], {});
  put(r + 2, B, REPORT_NOTE[2], {});
  // The comments row is a group of its own, collapsed (RaabeCalc's outline keeps the summary row below the group).
  rep.rows.set(3, { outline: 1, hidden: true });
  rep.rows.set(4, { collapsed: true });
  for (let c = D; c <= last; c++) rep.cols.set(c, { outline: 1 });
  rep.view.showGrid = false;
  const comments = () => data.comments.forEach((t, i) => { if (t) rep.set(3, D + 1 + i, { ...rep.get(3, D + 1 + i)!, v: t }); });
  return { sheet: rep, lastCol: last, comments };
}

const protectedScenarios = (s: Sheet) => s.protection?.sheet === "1" && s.protection.scenarios === "1";

export function scenarioFeature(app: CalcApp, userName: () => string) {
  const user = () => userName() || "Benutzer";
  const list = (s: Sheet) => s.scenarios?.list ?? [];

  /** Changes the sheet's scenario list as one undo step. */
  function save(label: string, sheet: Sheet, change: (s: Scenarios) => void): void {
    app.change(label, () => {
      app.history.whole(sheet);
      const next: Scenarios = sheet.scenarios ? structuredClone(sheet.scenarios) : { list: [] };
      change(next);
      sheet.scenarios = next.list.length ? next : undefined;
    });
  }

  /** Anzeigen: the scenario's values into its cells (one undo step), remembered as the one shown. */
  function show(sheet: Sheet, index: number): void {
    const sc = list(sheet)[index];
    if (!sc) return;
    app.change("Szenario anzeigen", () => {
      app.history.whole(sheet);
      applyScenario(app.ctx, sheet, sc);
      sheet.scenarios = { ...sheet.scenarios!, current: index, show: index };
    });
  }

  /** Szenario hinzufügen / bearbeiten: name, cells, comment, Schutz. */
  async function editDialog(sc: Scenario | null, init?: { name: string; cells: string; comment: string; locked: boolean; hidden: boolean }) {
    const sel = app.grid.selection.map((r) => (r.r0 === r.r1 && r.c0 === r.c1 ? cellName(r.r0, r.c0) : `${cellName(r.r0, r.c0)}:${cellName(r.r1, r.c1)}`)).join(";");
    const name = h("input", { type: "text", style: "width:100%", value: init?.name ?? sc?.name ?? "" }) as HTMLInputElement;
    const cells = h("input", { type: "text", style: "width:100%", value: init?.cells ?? (sc ? cellsText(sc.cells) : sel) }) as HTMLInputElement;
    const comment = h("textarea", { rows: 4, style: "width:100%;resize:none" }) as HTMLTextAreaElement;
    comment.value = init?.comment ?? (sc ? changedComment(sc.comment ?? "", user(), new Date()) : createdComment(user(), new Date()));
    const locked = h("input", { type: "checkbox", checked: init ? init.locked : sc ? !!sc.locked : true }) as HTMLInputElement;
    const hidden = h("input", { type: "checkbox", checked: init ? init.hidden : !!sc?.hidden }) as HTMLInputElement;
    const body = h("div", { class: "scenario-edit", style: "width:380px" },
      h("label", { class: "fc-label", style: "display:block" }, "Szenarioname:"), name,
      h("label", { class: "fc-label", style: "display:block" }, "Veränderbare Zellen:"), cells,
      h("div", { style: "font-size:11px;color:#555;margin:4px 0" }, "Mehrere Bereiche trennen Sie durch Semikolons."),
      h("label", { class: "fc-label", style: "display:block" }, "Kommentar:"), comment,
      h("fieldset", { style: "margin-top:8px" }, h("legend", null, "Schutz"),
        h("label", { class: "check" }, locked, "Änderungen verhindern"), h("label", { class: "check" }, hidden, "Ausblenden")));
    setTimeout(() => name.focus());
    const ok = await showDialog({ title: sc ? "Szenario bearbeiten" : "Szenario hinzufügen", body, buttons: [{ label: "OK", value: true, primary: true }, { label: "Abbrechen", value: false }], cancel: false });
    if (!ok) return null;
    return { name: name.value.trim(), cells: cells.value.trim(), comment: comment.value, locked: locked.checked, hidden: hidden.checked };
  }

  /** Szenariowerte: one field per changing cell; Hinzufügen saves and starts the next scenario. */
  async function valuesDialog(sheet: Sheet, cells: Pt[], values: string[]): Promise<{ vals: string[]; again: boolean } | null> {
    const inputs = cells.map((_, i) => h("input", { type: "text", style: "width:150px", value: values[i] ?? "" }) as HTMLInputElement);
    const body = h("div", { class: "scenario-values" },
      h("div", { style: "margin-bottom:6px" }, "Werte für die veränderbaren Zellen:"),
      h("div", { style: "max-height:190px;overflow-y:auto;display:grid;grid-template-columns:auto auto 1fr;gap:4px 8px;align-items:center" },
        ...cells.flatMap((p, i) => [h("span", null, `${i + 1}:`), h("span", null, cellLabel(app.wb, sheet, p)), inputs[i]])));
    setTimeout(() => { inputs[0]?.focus(); inputs[0]?.select(); });
    const r = await showDialog<"ok" | "add" | null>({ title: "Szenariowerte", body, buttons: [{ label: "OK", value: "ok", primary: true }, { label: "Abbrechen", value: null }, { label: "Hinzufügen", value: "add" }], cancel: null });
    if (!r) return null;
    return { vals: inputs.map((x) => x.value), again: r === "add" };
  }

  /** Adds (index -1) or replaces a scenario through the two dialogs; true when the user went on with Hinzufügen. */
  async function addOrEdit(sheet: Sheet, index: number): Promise<boolean> {
    const old = index >= 0 ? list(sheet)[index] : null;
    let init: Parameters<typeof editDialog>[1];
    for (;;) {
      const e = await editDialog(old, init);
      if (!e) return false;
      init = e;
      const cells = parseCells(e.cells);
      if (!e.name) { await messageBox("RaabeCalc", "Geben Sie einen Namen für das Szenario ein."); continue; }
      if (list(sheet).some((s, i) => i !== index && s.name.toLowerCase() === e.name.toLowerCase())) { await messageBox("RaabeCalc", "Ein Szenario mit diesem Namen gibt es schon. Wählen Sie einen anderen Namen."); continue; }
      if (!cells) { await messageBox("RaabeCalc", "Der Bezug auf die veränderbaren Zellen ist ungültig."); continue; }
      if (cells.length > MAX_CELLS) { await messageBox("RaabeCalc", `Ein Szenario kann höchstens ${MAX_CELLS} veränderbare Zellen haben.`); continue; }
      const vals = cells.map((p) => old?.cells.find((c) => c.row === p.row && c.col === p.col)?.val ?? currentText(app.wb, sheet, p));
      const v = await valuesDialog(sheet, cells, vals);
      if (!v) return false;
      const sc: Scenario = { name: e.name, cells: cells.map((p, i) => ({ ...p, val: v.vals[i] })), user: old?.user ?? user() };
      if (e.comment) sc.comment = e.comment;
      if (e.locked) sc.locked = true;
      if (e.hidden) sc.hidden = true;
      save(old ? "Szenario bearbeiten" : "Szenario hinzufügen", sheet, (s) => {
        if (old) s.list[index] = sc;
        else { s.list.push(sc); s.current = s.list.length - 1; }
      });
      if (!v.again) return false;
      if (old) return true;
      init = undefined;
    }
  }

  /** Szenarien zusammenführen: the scenarios of another sheet of this workbook. */
  async function mergeDialog(sheet: Sheet): Promise<void> {
    const others = app.wb.sheets.filter((s) => s !== sheet);
    const book = h("select", { style: "width:100%" }, h("option", null, app.wb.docName ?? "Mappe1")) as HTMLSelectElement;
    const sheets = h("select", { size: 6, style: "width:100%" }, ...others.map((s, i) => h("option", { value: String(i) }, s.name))) as HTMLSelectElement;
    const count = h("div", { style: "margin-top:6px" });
    const update = () => { const s = others[sheets.selectedIndex]; count.textContent = s ? `Szenarien auf dem Quellblatt: ${list(s).length}` : ""; };
    sheets.addEventListener("change", update);
    if (others.length) sheets.selectedIndex = 0;
    update();
    const body = h("div", { style: "width:300px" }, h("div", { class: "fc-label" }, "Szenarien zusammenführen aus"),
      h("label", { class: "fc-label", style: "display:block" }, "Mappe:"), book, h("label", { class: "fc-label", style: "display:block" }, "Blatt:"), sheets, count);
    const ok = await showDialog({ title: "Szenarien zusammenführen", body, buttons: [{ label: "OK", value: true, primary: true }, { label: "Abbrechen", value: false }], cancel: false });
    const from = others[sheets.selectedIndex];
    if (!ok || !from || !list(from).length) return;
    save("Szenarien zusammenführen", sheet, (s) => {
      for (const sc of list(from)) s.list.push({ ...structuredClone(sc), name: mergedName(sc.name, s.list.map((x) => x.name), new Date()) });
    });
  }

  /** Zusammenfassung: Berichtstyp and Ergebniszellen; true when a report was made. */
  async function summaryDialog(sheet: Sheet): Promise<boolean> {
    const sel = app.grid.selection.at(-1)!;
    const init = sheet.scenarios?.sqref?.length ? sheet.scenarios.sqref : [sel];
    const kind = h("input", { type: "radio", name: "scen-report", checked: true }) as HTMLInputElement;
    const pivot = h("input", { type: "radio", name: "scen-report", disabled: true }) as HTMLInputElement;
    const cells = h("input", { type: "text", style: "width:100%", value: init.map((r) => (r.r0 === r.r1 && r.c0 === r.c1 ? abs({ row: r.r0, col: r.c0 }) : `${abs({ row: r.r0, col: r.c0 })}:${abs({ row: r.r1, col: r.c1 })}`)).join(";") }) as HTMLInputElement;
    const body = h("div", { style: "width:300px" },
      h("fieldset", null, h("legend", null, "Berichtstyp"), h("label", { class: "check" }, kind, "Szenariobericht"), h("label", { class: "check", title: "Nicht verfügbar" }, pivot, "PivotTable-Szenarioberichts")),
      h("label", { class: "fc-label", style: "display:block" }, "Ergebniszellen:"), cells);
    for (;;) {
      setTimeout(() => { cells.focus(); cells.select(); });
      const ok = await showDialog({ title: "Szenariobericht", body, buttons: [{ label: "OK", value: true, primary: true }, { label: "Abbrechen", value: false }], cancel: false });
      if (!ok) return false;
      const text = cells.value.trim();
      const parts = text ? text.split(/[;,]/).map((t) => parseRange(t.trim().replace(/\$/g, ""))) : [];
      if (parts.some((p) => !p)) { await messageBox("RaabeCalc", "Der Bezug auf die Ergebniszellen ist ungültig."); continue; }
      makeReport(sheet, parts as Rect[]);
      return true;
    }
  }

  /** Makes the Szenariobericht for these result ranges (also from tests). */
  function makeReport(sheet: Sheet, ranges: Rect[]): Sheet {
    const results: Pt[] = [];
    for (const r of ranges) for (let row = r.r0; row <= r.r1; row++) for (let col = r.c0; col <= r.c1; col++) results.push({ row, col });
    const data = reportData(app.wb, sheet, app.calc, results);
    const built = buildReport(app.wb, data, reportName(app.wb));
    app.change("Szenariobericht", () => {
      app.history.whole(sheet);
      sheet.scenarios = { ...sheet.scenarios!, sqref: ranges.length ? ranges : undefined };
      app.history.sheets();
      const at = app.wb.sheets.indexOf(sheet);
      app.wb.sheets.splice(at, 0, built.sheet);
      app.wb.active = at;
      // Column widths as Excel fits them: C to its labels, D… all as wide as the widest (comments left out: they wrap).
      app.autofitCols([2]);
      const cols = Array.from({ length: built.lastCol - 2 }, (_, i) => 3 + i);
      app.autofitCols(cols);
      const w = Math.max(...cols.map((c) => built.sheet.cols.get(c)?.width ?? 0));
      for (const c of cols) built.sheet.cols.set(c, { ...(built.sheet.cols.get(c) ?? {}), width: w, bestFit: true, outline: 1 });
      built.comments();
    });
    return built.sheet;
  }

  /** The Szenario-Manager. */
  async function manager(): Promise<void> {
    const sheet = app.sheet;
    let selected = sheet.scenarios?.current ?? 0;
    for (;;) {
      const prot = protectedScenarios(sheet);
      const visible = list(sheet).map((s, i) => ({ s, i })).filter((x) => !(prot && x.s.hidden));
      if (!visible.some((x) => x.i === selected)) selected = visible[0]?.i ?? -1;
      let pending: "add" | "delete" | "edit" | "merge" | "summary" | "show" | null = null;
      const box = h("select", { size: 8, style: "width:100%;height:150px" }, ...visible.map((x) => h("option", { value: String(x.i), selected: x.i === selected || undefined }, x.s.name))) as HTMLSelectElement;
      const cellsField = h("input", { type: "text", readOnly: true, style: "width:100%" }) as HTMLInputElement;
      const commentField = h("textarea", { readOnly: true, rows: 3, style: "width:100%;resize:none" }) as HTMLTextAreaElement;
      const closeWith = (what: typeof pending) => {
        pending = what;
        (box.closest(".dialog")?.querySelector(".dlg-buttons button:last-child") as HTMLButtonElement | null)?.click();
      };
      const btn = (label: string, what: typeof pending) => h("button", { style: "width:100%;margin-bottom:6px", onclick: () => closeWith(what) }, label) as HTMLButtonElement;
      const add = btn("Hinzufügen...", "add"), del = btn("Löschen", "delete"), edit = btn("Bearbeiten...", "edit"), merge = btn("Zusammenführen...", "merge"), sum = btn("Zusammenfassung...", "summary");
      const empty = h("div", { style: "position:absolute;inset:8px;color:#555;white-space:normal" }, "Es gibt noch keine Szenarien. Mit „Hinzufügen“ legen Sie eines an.");
      const update = () => {
        const i = Number(box.value);
        const sc = visible.length && box.selectedIndex >= 0 ? list(sheet)[i] : undefined;
        selected = sc ? i : -1;
        cellsField.value = sc ? cellsText(sc.cells) : "";
        commentField.value = sc?.comment ?? "";
        const lockedNow = !!sc && prot && !!sc.locked;
        del.disabled = !sc || lockedNow;
        edit.disabled = !sc || lockedNow;
        sum.disabled = !visible.length;
        add.disabled = prot;
        merge.disabled = prot;
        empty.hidden = visible.length > 0;
        const showBtn = box.closest(".dialog")?.querySelector(".dlg-buttons button") as HTMLButtonElement | null;
        if (showBtn) showBtn.disabled = !sc;
      };
      box.addEventListener("change", update);
      box.addEventListener("dblclick", () => { if (selected >= 0) closeWith("show"); });
      const body = h("div", { class: "scenario-manager", style: "width:420px" },
        h("div", { style: "display:grid;grid-template-columns:1fr 130px;gap:10px" },
          h("div", null, h("div", { class: "fc-label" }, "Szenarien:"), h("div", { style: "position:relative" }, box, empty)),
          h("div", { style: "padding-top:26px" }, add, del, edit, merge, sum)),
        h("label", { class: "fc-label", style: "display:block" }, "Veränderbare Zellen:"), cellsField,
        h("label", { class: "fc-label", style: "display:block" }, "Kommentar:"), commentField);
      setTimeout(() => { update(); box.focus(); });
      const r = await showDialog<"show" | null>({ title: "Szenario-Manager", body, buttons: [{ label: "Anzeigen", value: "show", primary: true }, { label: "Schließen", value: null }], cancel: null });
      const what = r === "show" ? (selected >= 0 ? "show" : null) : pending;
      if (!what) return;
      if (what === "show") show(sheet, selected);
      else if (what === "add") { while (await addOrEdit(sheet, -1)); }
      else if (what === "edit") { if (await addOrEdit(sheet, selected)) while (await addOrEdit(sheet, -1)); }
      else if (what === "delete") save("Szenario löschen", sheet, (s) => {
        s.list.splice(selected, 1);
        if (s.current !== undefined && s.current >= s.list.length) s.current = Math.max(0, s.list.length - 1);
        if (s.show !== undefined && s.show >= s.list.length) delete s.show;
      });
      else if (what === "merge") await mergeDialog(sheet);
      else if (what === "summary") { if (await summaryDialog(sheet)) return; }
    }
  }

  return { manager, show, makeReport, save };
}

export type ScenarioFeature = ReturnType<typeof scenarioFeature>;
// Kept for the test API: a scenario added without the dialogs.
export function scenarioOf(name: string, cells: Pt[], vals: string[], opts: Partial<Scenario> = {}): Scenario {
  return { name, cells: cells.map((p, i) => ({ ...p, val: vals[i] })), locked: true, ...opts };
}
