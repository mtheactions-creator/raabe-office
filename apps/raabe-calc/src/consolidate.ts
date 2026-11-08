// Daten > Datentools > Konsolidieren (D-084), as German Excel 365 computes it (recorded in our own
// invisible Excel: probe_excel_consolidate.py, results/2026-11-08-excel-consolidate.json):
// - Excel keeps the references sorted (Vorhandene Verweise) and works through them in that order;
// - by position (no labels) the result is as large as the largest source; with Beschriftung aus
//   Oberster Zeile / Linker Spalte the labels decide, compared without regard to case, spelt as first
//   seen; a label that is new goes before the next label of its own source that is already listed,
//   otherwise to the end;
// - Summe … Varianz over the numbers of a result cell (Anzahl counts text too); a cell nothing
//   counts for stays empty; a result cell takes the number format of its first source cell;
// - the destination: one cell (or one row / one column) grows as needed, a larger range limits the
//   result (labels that do not fit are left out);
// - Verknüpfungen mit Quelldaten: per label one row for each source with links to the source cells
//   (grouped and hidden, with the workbook's name beside them when labels come from the left
//   column), below them the function over them;
// - the settings stay with the destination sheet (<dataConsolidate>) for the dialog's next use.

import { h, messageBox, showDialog } from "@raabe/shell";
import {
  CellError, colName, err, parseRange, type Cell, type ConsolidateFunction, type Consolidation, type Rect, type Sheet, type Value, type Workbook,
} from "@raabe/sheet-engine";
import type { CalcApp } from "./app";
import type { Ctx } from "./actions";

/** The dialog's functions (Excel's list), the file's names and the function a link row's total uses. */
export const FUNCTIONS: { name: string; fn: ConsolidateFunction; formula: string }[] = [
  { name: "Summe", fn: "sum", formula: "SUM" },
  { name: "Anzahl", fn: "count", formula: "COUNTA" },
  { name: "Mittelwert", fn: "average", formula: "AVERAGE" },
  { name: "Maximum", fn: "max", formula: "MAX" },
  { name: "Minimum", fn: "min", formula: "MIN" },
  { name: "Produkt", fn: "product", formula: "PRODUCT" },
  { name: "Anzahl Zahlen", fn: "countNums", formula: "COUNT" },
  { name: "Standardabweichung (Stichprobe)", fn: "stdDev", formula: "STDEV" },
  { name: "Standardabweichung (Grundgesamtheit)", fn: "stdDevp", formula: "STDEVP" },
  { name: "Varianz (Stichprobe)", fn: "var", formula: "VAR" },
  { name: "Varianz (Grundgesamtheit)", fn: "varp", formula: "VARP" },
];

/** A sheet name as a reference writes it ('Mein Blatt'!, Süd!). */
export function quoteSheet(name: string): string {
  return /^[\p{L}_][\p{L}\p{N}_.]*$/u.test(name) && !/^[A-Za-z]{1,3}\d+$/.test(name) ? name : `'${name.replace(/'/g, "''")}'`;
}

const absRect = (r: Rect) => `$${colName(r.c0)}$${r.r0 + 1}${r.r0 === r.r1 && r.c0 === r.c1 ? "" : `:$${colName(r.c1)}$${r.r1 + 1}`}`;

export interface Source { sheet: Sheet; rect: Rect }

/** A source as the dialog lists it ("Nord!$A$1:$C$4"). */
export const sourceText = (s: Source) => `${quoteSheet(s.sheet.name)}!${absRect(s.rect)}`;

/** Excel's order of the references: as the list shows them, sorted. */
export function sortSources(list: Source[]): Source[] {
  return [...list].sort((a, b) => sourceText(a).localeCompare(sourceText(b), "de", { sensitivity: "accent" }));
}

/** A reference typed into Verweis: Blatt!Bereich, a range of the active sheet, or a defined name. */
export function parseSource(text: string, wb: Workbook, active: Sheet): Source | null {
  const t = text.trim().replace(/^=/, "");
  const m = /^(?:'((?:[^']|'')+)'|([^!']+))!(.+)$/.exec(t);
  if (m) {
    const sheet = wb.sheetByName((m[1] ?? m[2]).replace(/''/g, "'"));
    const rect = parseRange(m[3].replace(/\$/g, ""));
    return sheet && rect ? { sheet, rect } : null;
  }
  const rect = parseRange(t.replace(/\$/g, ""));
  if (rect) return { sheet: active, rect };
  const n = wb.names.find((x) => x.name.toLowerCase() === t.toLowerCase() && (x.sheet === undefined || wb.sheets[x.sheet] === active));
  return n ? parseSource(n.ref, wb, active) : null;
}

/** The label a cell gives (compared without case); null for an empty cell. */
export function labelKey(v: Value): string | null {
  if (v === null || v === "") return null;
  if (typeof v === "string") return "s" + v.toLowerCase();
  if (v instanceof CellError) return "e" + v.code;
  return (typeof v === "number" ? "n" : "b") + String(v);
}

export interface Label { key: string; value: Value }

/**
 * Adds a source's labels to the list as Excel does: a new one goes before the next label of the
 * same source that is listed already, else at the end; none beyond `cap`.
 */
export function addLabels(list: Label[], labels: Label[], cap = Infinity): void {
  for (let i = 0; i < labels.length; i++) {
    const l = labels[i];
    if (list.some((x) => x.key === l.key)) continue;
    if (list.length >= cap) continue;
    let at = list.length;
    for (let j = i + 1; j < labels.length; j++) {
      const k = list.findIndex((x) => x.key === labels[j].key);
      if (k >= 0) { at = k; break; }
    }
    list.splice(at, 0, l);
  }
}

/** One source cell that counts for a result cell. */
interface Entry { sheet: Sheet; row: number; col: number; v: Value; fmt: string; src: number }

const counts = (fn: ConsolidateFunction, v: Value) => (fn === "count" ? v !== null && v !== "" : typeof v === "number" || v instanceof CellError);

/** The function over the values that count; null when none does (the result cell stays empty). */
export function aggregate(fn: ConsolidateFunction, values: Value[]): Value {
  const vs = values.filter((v) => counts(fn, v));
  if (!vs.length) return null;
  if (fn === "count") return vs.length;
  const e = vs.find((v) => v instanceof CellError);
  if (e) return e;
  const xs = vs as number[];
  const n = xs.length, sum = xs.reduce((a, b) => a + b, 0), mean = sum / n;
  const ss = xs.reduce((a, b) => a + (b - mean) ** 2, 0);
  const div0 = err("#DIV/0!");
  switch (fn) {
    case "sum": return sum;
    case "average": return mean;
    case "max": return Math.max(...xs);
    case "min": return Math.min(...xs);
    case "product": return xs.reduce((a, b) => a * b, 1);
    case "countNums": return n;
    case "stdDev": return n < 2 ? div0 : Math.sqrt(ss / (n - 1));
    case "stdDevp": return Math.sqrt(ss / n);
    case "var": return n < 2 ? div0 : ss / (n - 1);
    case "varp": return ss / n;
  }
}

export interface Options { fn: ConsolidateFunction; top: boolean; left: boolean; link: boolean; bookName: string }

/** One cell the result writes: a value or a formula (file form), the number format to take. */
export interface Out { row: number; col: number; v: Value; f?: string; fmt?: string }

export interface Plan {
  cells: Out[];
  /** The area the result covers (cleared first). */
  area: Rect;
  /** Verknüpfungen: the link rows (grouped, hidden) and the total rows below them. */
  detailRows: number[];
  totalRows: number[];
  /** The column with the workbook's name (links with labels from the left column). */
  nameCol?: number;
}

/** What Konsolidieren writes for these sources into `dest` (its top-left cell, or the range it may fill). */
export function plan(wb: Workbook, dest: Rect, sources: Source[], o: Options): Plan {
  const srcs = sortSources(sources);
  const one = (a: number, b: number) => a === b;
  const single = one(dest.r0, dest.r1) && one(dest.c0, dest.c1);
  const rowCap = single || one(dest.r0, dest.r1) || o.link ? Infinity : dest.r1 - dest.r0 + 1 - (o.top ? 1 : 0);
  const colCap = single || one(dest.c0, dest.c1) || o.link ? Infinity : dest.c1 - dest.c0 + 1 - (o.left ? 1 : 0);
  const rowLabels: Label[] = [], colLabels: Label[] = [];
  let posRows = 0, posCols = 0;
  // Keys of each source's data rows and columns (null: a row without label, left out).
  const keyed = srcs.map((s) => {
    const r0 = s.rect.r0 + (o.top ? 1 : 0), c0 = s.rect.c0 + (o.left ? 1 : 0);
    const rows: (string | null)[] = [], cols: (string | null)[] = [];
    if (o.left) {
      const ls: Label[] = [];
      for (let r = r0; r <= s.rect.r1; r++) { const v = s.sheet.value(r, s.rect.c0); const k = labelKey(v); rows.push(k); if (k) ls.push({ key: k, value: v }); }
      addLabels(rowLabels, ls, rowCap);
    } else {
      for (let r = r0; r <= s.rect.r1; r++) rows.push(r - r0 < rowCap ? `p${r - r0}` : null);
      posRows = Math.max(posRows, Math.min(rowCap, s.rect.r1 - r0 + 1));
    }
    if (o.top) {
      const ls: Label[] = [];
      for (let c = c0; c <= s.rect.c1; c++) { const v = s.sheet.value(s.rect.r0, c); const k = labelKey(v); cols.push(k); if (k) ls.push({ key: k, value: v }); }
      addLabels(colLabels, ls, colCap);
    } else {
      for (let c = c0; c <= s.rect.c1; c++) cols.push(c - c0 < colCap ? `p${c - c0}` : null);
      posCols = Math.max(posCols, Math.min(colCap, s.rect.c1 - c0 + 1));
    }
    return { s, r0, c0, rows, cols };
  });
  const rowKeys = o.left ? rowLabels.map((l) => l.key) : Array.from({ length: posRows }, (_, i) => `p${i}`);
  const colKeys = o.top ? colLabels.map((l) => l.key) : Array.from({ length: posCols }, (_, i) => `p${i}`);
  // The source cells of each result cell, in the order of the sources and their rows.
  const buckets = new Map<string, Entry[]>();
  keyed.forEach((k, src) => {
    k.rows.forEach((rk, i) => {
      if (rk === null || !rowKeys.includes(rk)) return;
      k.cols.forEach((ck, j) => {
        if (ck === null || !colKeys.includes(ck)) return;
        const row = k.r0 + i, col = k.c0 + j;
        const v = k.s.sheet.value(row, col);
        if (!counts(o.fn, v)) return;
        const key = `${rk}|${ck}`;
        const list = buckets.get(key) ?? [];
        list.push({ sheet: k.s.sheet, row, col, v, fmt: wb.styles.get(k.s.sheet.get(row, col)?.s).numFmt, src });
        buckets.set(key, list);
      });
    });
  });
  const cells: Out[] = [];
  const { r0, c0 } = dest;
  const nameCol = o.link && o.left ? c0 + 1 : undefined;
  const dataC0 = c0 + (o.left ? 1 : 0) + (nameCol !== undefined ? 1 : 0);
  if (o.top) colLabels.forEach((l, j) => cells.push({ row: r0, col: dataC0 + j, v: l.value }));
  const detailRows: number[] = [], totalRows: number[] = [];
  let r = r0 + (o.top ? 1 : 0);
  if (!o.link) {
    rowKeys.forEach((rk, i) => {
      if (o.left) cells.push({ row: r + i, col: c0, v: rowLabels[i].value });
      colKeys.forEach((ck, j) => {
        const list = buckets.get(`${rk}|${ck}`) ?? [];
        const v = aggregate(o.fn, list.map((e) => e.v));
        if (v !== null) cells.push({ row: r + i, col: dataC0 + j, v, fmt: list[0]?.fmt });
      });
    });
    r += rowKeys.length;
  } else {
    const fname = FUNCTIONS.find((f) => f.fn === o.fn)!.formula;
    rowKeys.forEach((rk, i) => {
      const first = r;
      srcs.forEach((_, src) => {
        const per = colKeys.map((ck) => (buckets.get(`${rk}|${ck}`) ?? []).filter((e) => e.src === src));
        const n = Math.max(0, ...per.map((l) => l.length));
        for (let d = 0; d < n; d++) {
          // The workbook's name stands beside the first row of each source (recorded).
          if (nameCol !== undefined && d === 0) cells.push({ row: r, col: nameCol, v: o.bookName });
          per.forEach((l, j) => {
            const e = l[d];
            if (!e) return;
            const fmt = (buckets.get(`${rk}|${colKeys[j]}`) ?? [])[0]?.fmt;
            cells.push({ row: r, col: dataC0 + j, v: e.v, f: `${quoteSheet(e.sheet.name)}!$${colName(e.col)}$${e.row + 1}`, fmt });
          });
          detailRows.push(r);
          r++;
        }
      });
      if (o.left) cells.push({ row: r, col: c0, v: rowLabels[i].value });
      colKeys.forEach((ck, j) => {
        const list = buckets.get(`${rk}|${ck}`) ?? [];
        if (!list.length) return;
        const col = colName(dataC0 + j);
        const range = r - 1 === first ? `${col}${first + 1}` : `${col}${first + 1}:${col}${r}`;
        cells.push({ row: r, col: dataC0 + j, v: aggregate(o.fn, list.map((e) => e.v)), f: `${fname}(${range})`, fmt: list[0].fmt });
      });
      totalRows.push(r);
      r++;
    });
  }
  const lastCol = dataC0 + colKeys.length - 1;
  const area = { r0, c0, r1: Math.max(r0, r - 1), c1: Math.max(c0, lastCol) };
  return { cells, area, detailRows, totalRows, nameCol };
}

/** Writes a plan into the destination sheet (inside an undo step). */
export function applyPlan(ctx: Ctx, sheet: Sheet, p: Plan): void {
  ctx.history.cells(sheet, p.area);
  for (let row = p.area.r0; row <= p.area.r1; row++) for (let col = p.area.c0; col <= p.area.c1; col++) {
    const old = sheet.get(row, col);
    if (old) sheet.set(row, col, old.s !== undefined ? { v: null, s: old.s } : undefined);
  }
  for (const c of p.cells) {
    const old = sheet.get(c.row, c.col);
    let s = old?.s;
    if (c.fmt && c.fmt !== ctx.wb.styles.get(s).numFmt) s = ctx.wb.styles.derive(s, (st) => { st.numFmt = c.fmt!; });
    const cell: Cell = { v: c.v, ...(s !== undefined ? { s } : {}) };
    if (c.f !== undefined) cell.f = c.f;
    sheet.set(c.row, c.col, cell);
  }
  for (const r of p.detailRows) { ctx.history.row(sheet, r); sheet.rows.set(r, { ...(sheet.rows.get(r) ?? {}), outline: 1, hidden: true }); }
  for (const r of p.totalRows) { ctx.history.row(sheet, r); sheet.rows.set(r, { ...(sheet.rows.get(r) ?? {}), collapsed: true }); }
}

/** The settings Excel keeps in the destination sheet. */
export function settingsOf(sources: Source[], o: Omit<Options, "bookName">): Consolidation {
  const c: Consolidation = { refs: sortSources(sources).map((s) => ({ ref: `${colName(s.rect.c0)}${s.rect.r0 + 1}${s.rect.r0 === s.rect.r1 && s.rect.c0 === s.rect.c1 ? "" : `:${colName(s.rect.c1)}${s.rect.r1 + 1}`}`, sheet: s.sheet.name })) };
  if (o.fn !== "sum") c.fn = o.fn;
  if (o.top) c.top = true;
  if (o.left) c.left = true;
  if (o.link) c.link = true;
  return c;
}

/** The references of kept settings as the dialog lists them. */
export function sourcesOf(wb: Workbook, active: Sheet, c: Consolidation | undefined): string[] {
  const out: string[] = [];
  for (const r of c?.refs ?? []) {
    if (r.name) out.push(r.name);
    else if (r.ref) {
      const sheet = r.sheet ? wb.sheetByName(r.sheet) : active;
      const rect = parseRange(r.ref);
      if (sheet && rect) out.push(sourceText({ sheet, rect }));
    }
  }
  return out;
}

const overlaps = (a: Rect, b: Rect) => a.r0 <= b.r1 && b.r0 <= a.r1 && a.c0 <= b.c1 && b.c0 <= a.c1;

export function consolidateFeature(app: CalcApp) {
  /** Konsolidieren into the selection (also from tests, without the dialog); a message when it cannot. */
  async function run(dest: Sheet, destRect: Rect, refs: string[], o: Omit<Options, "bookName">): Promise<boolean> {
    const sources: Source[] = [];
    for (const t of refs) {
      const s = parseSource(t, app.wb, dest);
      if (!s) { await messageBox("RaabeCalc", `Der Bezug „${t}“ ist ungültig.`); return false; }
      sources.push(s);
    }
    if (!sources.length) { await messageBox("RaabeCalc", "Geben Sie mindestens einen Bezug an, der konsolidiert werden soll."); return false; }
    if (o.link && sources.some((s) => s.sheet === dest)) { await messageBox("RaabeCalc", "Verknüpfungen mit Quelldaten gehen nur mit Quellen auf anderen Blättern als dem Zielblatt."); return false; }
    const p = plan(app.wb, destRect, sources, { ...o, bookName: app.wb.docName ?? "Mappe1" });
    if (sources.some((s) => s.sheet === dest && overlaps(s.rect, p.area))) { await messageBox("RaabeCalc", "Quell- und Zielbereich dürfen sich nicht überschneiden."); return false; }
    app.change("Konsolidieren", () => {
      app.history.whole(dest);
      applyPlan(app.ctx, dest, p);
      dest.consolidate = settingsOf(sources, o);
      if (p.nameCol !== undefined) app.autofitCols([p.nameCol]);
    });
    return true;
  }

  /** The Konsolidieren dialog with the settings this sheet keeps. */
  async function dialog(): Promise<void> {
    const sheet = app.sheet;
    const kept = sheet.consolidate;
    const fn = h("select", { style: "width:100%" }, ...FUNCTIONS.map((f) => h("option", { value: f.fn, selected: f.fn === (kept?.fn ?? "sum") || undefined }, f.name))) as HTMLSelectElement;
    const ref = h("input", { type: "text", style: "flex:1" }) as HTMLInputElement;
    let refs = sourcesOf(app.wb, sheet, kept);
    const list = h("select", { size: 6, style: "width:100%" }) as HTMLSelectElement;
    const fill = () => { list.replaceChildren(...refs.map((r) => h("option", null, r))); };
    fill();
    if (refs.length) ref.value = refs[0];
    list.addEventListener("change", () => { if (list.selectedIndex >= 0) ref.value = refs[list.selectedIndex]; });
    const norm = (t: string) => { const s = parseSource(t, app.wb, sheet); return s ? sourceText(s) : null; };
    const add = async () => {
      const t = ref.value.trim();
      if (!t) return;
      const n = norm(t);
      if (!n) { await messageBox("RaabeCalc", "Der Bezug ist ungültig."); return; }
      if (refs.some((r) => r.toLowerCase() === n.toLowerCase())) { await messageBox("RaabeCalc", "Dieser Bezug ist schon in der Liste."); return; }
      refs = [...refs, n].sort((a, b) => a.localeCompare(b, "de", { sensitivity: "accent" }));
      fill();
    };
    const del = () => { const i = list.selectedIndex; if (i < 0) return; refs.splice(i, 1); fill(); ref.value = ""; };
    const top = h("input", { type: "checkbox", checked: !!kept?.top }) as HTMLInputElement;
    const left = h("input", { type: "checkbox", checked: !!kept?.left }) as HTMLInputElement;
    const link = h("input", { type: "checkbox", checked: !!kept?.link }) as HTMLInputElement;
    const body = h("div", { class: "consolidate", style: "width:440px" },
      h("label", { class: "fc-label", style: "display:block" }, "Funktion:"), fn,
      h("label", { class: "fc-label", style: "display:block" }, "Verweis:"),
      h("div", { style: "display:flex;gap:8px" }, ref, h("button", { disabled: true, title: "Nicht verfügbar" }, "Durchsuchen...")),
      h("label", { class: "fc-label", style: "display:block" }, "Vorhandene Verweise:"),
      h("div", { style: "display:flex;gap:8px;align-items:flex-start" }, list,
        h("div", { style: "display:flex;flex-direction:column;gap:6px;min-width:100px" }, h("button", { onclick: () => void add() }, "Hinzufügen"), h("button", { onclick: del }, "Löschen"))),
      h("fieldset", { style: "margin-top:8px" }, h("legend", null, "Beschriftung aus:"),
        h("label", { class: "check" }, top, "Oberster Zeile"), h("label", { class: "check" }, left, "Linker Spalte"),
        h("label", { class: "check" }, link, "Verknüpfungen mit Quelldaten")));
    setTimeout(() => ref.focus());
    for (;;) {
      const ok = await showDialog({ title: "Konsolidieren", body, buttons: [{ label: "OK", value: true, primary: true }, { label: "Schließen", value: false }], cancel: false });
      if (!ok) return;
      // A reference left in Verweis counts as added (Excel).
      const t = ref.value.trim();
      const n = t ? norm(t) : null;
      const all = n && !refs.some((r) => r.toLowerCase() === n.toLowerCase()) ? [...refs, n] : refs;
      const sel = app.grid.selection.at(-1)!;
      if (await run(sheet, sel, all, { fn: fn.value as ConsolidateFunction, top: top.checked, left: left.checked, link: link.checked })) return;
    }
  }

  return { dialog, run };
}
