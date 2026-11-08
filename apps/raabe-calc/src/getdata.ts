// Daten abrufen und transformieren, local only (D-080, D-092): Daten > Daten abrufen with Excel's menu, Aus Datei >
// Aus Arbeitsmappe (a sheet or table of another .xlsx, chosen in a Navigator), Aus Text/CSV and Aus anderen
// Quellen > Aus Tabelle/Bereich load the data as a table on a new sheet named after the source, as Excel's
// Power Query does when it loads a query (TableStyleMedium2, columns fitted). There is no Power Query editor and
// no query is kept: the table holds the values (D-080). Zuletzt verwendete Quellen lists the files loaded before,
// Vorhandene Verbindungen the workbook's connections (from the file's xl/connections.xml) and its tables.
// Cloud, web and database sources are in the menu and greyed out. Labels of the menu as Excel 365 shows them (the
// menu does not open through UI Automation, so its entries were not recorded; the ribbon's labels and keytips were).

import { h, messageBox, showDialog, type MenuEntry } from "@raabe/shell";
import { readXlsx } from "@raabe/ooxml";
import { Sheet, nextTableName, readCsv, tableData, type Rect, type SheetTable, type Workbook } from "@raabe/sheet-engine";
import type { CalcApp } from "./app";

type Run = (id: string, fn: () => void | Promise<void>, enabled?: () => boolean, checked?: () => boolean) => void;

export const GETDATA_COMMANDS = [
  { id: "data.get", label: "Daten abrufen", icon: "document" },
  { id: "data.fromWorkbook", label: "Aus Arbeitsmappe", icon: "document" },
  { id: "data.fromTable", label: "Aus Tabelle/Bereich", icon: "table" },
  { id: "data.recentSources", label: "Zuletzt verwendete Quellen", icon: "document" },
  { id: "data.connections", label: "Vorhandene Verbindungen", icon: "link" },
  { id: "data.fromText", label: "Aus Text/CSV", icon: "document" },
];

const greyed = (label: string): MenuEntry => ({ label, run: () => {}, disabled: true });

/** Daten abrufen: Excel's menu; what needs Power Query's engine, a database or the cloud stays greyed. */
export function getDataMenu(): MenuEntry[] {
  return [
    { label: "Aus Datei", icon: "document", submenu: [{ cmd: "data.fromWorkbook" }, { cmd: "data.fromText" }, greyed("Aus XML"), greyed("Aus JSON"), greyed("Aus PDF"), greyed("Aus Ordner"), greyed("Aus SharePoint-Ordner")] },
    { label: "Aus Datenbank", icon: "table", submenu: [greyed("Aus SQL Server-Datenbank"), greyed("Aus Microsoft Access-Datenbank"), greyed("Aus Analysis Services"), greyed("Aus SQL Server Analysis Services-Datenbank (Import)")] },
    { label: "Aus Azure", icon: "globe", disabled: true, submenu: [] },
    { label: "Aus Power Platform", icon: "globe", disabled: true, submenu: [] },
    { label: "Aus Onlinedienste", icon: "globe", disabled: true, submenu: [] },
    { label: "Aus anderen Quellen", icon: "table", submenu: [{ cmd: "data.fromTable" }, greyed("Aus dem Web"), greyed("Aus Microsoft Query"), greyed("Aus OData-Feed"), greyed("Aus ODBC"), greyed("Aus OLEDB"), greyed("Leere Abfrage")] },
    "-",
    greyed("Abfragen kombinieren"),
    greyed("Power Query-Editor starten..."),
    greyed("Datenquelleneinstellungen..."),
    greyed("Abfrageoptionen"),
  ];
}

const RECENT_KEY = "raabe-calc-recent-sources";
interface Recent { path: string; kind: "xlsx" | "csv"; used: number }
function loadRecent(): Recent[] {
  try { return JSON.parse(localStorage.getItem(RECENT_KEY) ?? "[]") as Recent[]; } catch { return []; }
}
function remember(path: string, kind: Recent["kind"]): void {
  try {
    const list = [{ path, kind, used: Date.now() }, ...loadRecent().filter((r) => r.path.toLowerCase() !== path.toLowerCase())].slice(0, 20);
    localStorage.setItem(RECENT_KEY, JSON.stringify(list));
  } catch { /* no storage */ }
}
const baseName = (p: string) => p.split(/[\\/]/).pop() ?? p;

/** A block of values to load: header texts and rows of values with their number formats. */
interface Block { name: string; head: string[]; rows: { v: string | number | boolean | null; fmt?: string }[][] }

function blockOf(wb: Workbook, sheet: Sheet, r: Rect, name: string): Block {
  const head: string[] = [];
  const rows: Block["rows"] = [];
  for (let c = r.c0; c <= r.c1; c++) { const v = sheet.value(r.r0, c); head.push(v === null || v === "" ? `Column${c - r.c0 + 1}` : String(v)); }
  for (let row = r.r0 + 1; row <= r.r1; row++) {
    const out: Block["rows"][number] = [];
    for (let c = r.c0; c <= r.c1; c++) {
      const cell = sheet.get(row, c);
      const v = cell?.v;
      const fmt = cell?.s !== undefined ? wb.styles.get(cell.s).numFmt : undefined;
      out.push({ v: typeof v === "number" || typeof v === "string" || typeof v === "boolean" ? v : null, fmt: fmt && fmt !== "General" ? fmt : undefined });
    }
    rows.push(out);
  }
  return { name, head, rows };
}

export function getDataFeature(app: CalcApp) {
  /** A name not yet used by `taken` (Excel: "Noten (2)" for sheets, "Noten_2" for tables). */
  function free(base: string, taken: string[], sep: (k: number) => string): string {
    const low = new Set(taken.map((x) => x.toLowerCase()));
    if (!low.has(base.toLowerCase())) return base;
    for (let k = 2; ; k++) if (!low.has((base + sep(k)).toLowerCase())) return base + sep(k);
  }

  /** Loads a block as a table on a new sheet before the active one (as Power Query's Laden does), one undo step. */
  function load(b: Block): void {
    if (app.grid.isEditing) app.grid.commitEdit("none");
    const wb = app.wb;
    const clean = b.name.replace(/[\\/?*[\]:]/g, "_").slice(0, 31) || "Abfrage";
    const sheetName = free(clean, wb.sheets.map((s) => s.name), (k) => ` (${k})`);
    const tables = wb.sheets.flatMap((s) => (s.tables ?? []).map((t) => t.name));
    let tname = b.name.replace(/[^\p{L}\p{N}_.]/gu, "_");
    if (!/^[\p{L}_]/u.test(tname)) tname = "_" + tname;
    if (/^[A-Za-z]{1,3}\d+$/.test(tname)) tname += "_";
    tname = free(tname || nextTableName(tables), tables, (k) => `_${k}`);
    const cols = Math.max(1, b.head.length);
    // Header names unique, as a table needs them.
    const names: string[] = [];
    for (const n of b.head) { let x = n || "Spalte"; let k = 2; while (names.some((y) => y.toLowerCase() === x.toLowerCase())) x = `${n}${k++}`; names.push(x); }
    app.change("Daten laden", () => {
      app.history.sheets();
      const s = new Sheet(sheetName);
      s.defaultColWidth = app.sheet.defaultColWidth;
      names.forEach((n, c) => s.set(0, c, { v: n }));
      b.rows.forEach((row, r) => row.forEach((x, c) => {
        if (x.v === null) return;
        s.set(r + 1, c, { v: x.v, s: x.fmt ? wb.styles.derive(undefined, (st) => { st.numFmt = x.fmt!; }) : undefined });
      }));
      const ref: Rect = { r0: 0, c0: 0, r1: Math.max(1, b.rows.length), c1: cols - 1 };
      const t: SheetTable = { name: tname, ref, header: true, totals: false, columns: names.map((name) => ({ name })), style: "TableStyleMedium2",
        firstColumn: false, lastColumn: false, rowStripes: true, columnStripes: false, filterButton: true };
      s.tables = [t];
      const at = wb.active;
      wb.sheets.splice(at, 0, s);
      wb.active = at;
    });
    app.grid.refreshLayout();
    app.autofitCols([...Array(cols).keys()]);
    app.grid.selectCell(0, 0);
  }

  /** Navigator: the sheets and tables of a workbook; the one chosen is loaded. */
  async function navigator(path: string): Promise<void> {
    let src: Workbook;
    try { src = readXlsx(await app.platform.readFile(path)); } catch (e) {
      await messageBox("RaabeCalc", `Die Datei „${baseName(path)}“ kann nicht gelesen werden.\n\n${e instanceof Error ? e.message : String(e)}`);
      return;
    }
    remember(path, "xlsx");
    const entries: { label: string; kind: "Tabelle" | "Blatt"; block: () => Block | null }[] = [];
    for (const s of src.sheets) for (const t of s.tables ?? []) entries.push({ label: t.name, kind: "Tabelle", block: () => blockOf(src, s, { ...t.ref, r1: tableData(t).r1 }, t.name) });
    for (const s of src.sheets) {
      if (s.hidden) continue;
      entries.push({ label: s.name, kind: "Blatt", block: () => { const u = s.usedRange(); return u ? blockOf(src, s, u, s.name) : null; } });
    }
    if (!entries.length) { await messageBox("RaabeCalc", "Die Arbeitsmappe enthält keine Daten."); return; }
    const list = h("select", { size: "12", style: "width:240px" }) as HTMLSelectElement;
    entries.forEach((e, i) => list.append(h("option", { value: String(i) }, `${e.kind === "Tabelle" ? "▦" : "▤"} ${e.label}`)));
    list.selectedIndex = 0;
    const preview = h("div", { style: "width:420px;height:260px;overflow:auto;border:1px solid #d0d0d0;font-size:12px" });
    const show = () => {
      const b = entries[list.selectedIndex]?.block();
      preview.replaceChildren();
      if (!b) return;
      const tb = h("table", { style: "border-collapse:collapse" });
      tb.append(h("tr", null, ...b.head.map((x) => h("th", { style: "border:1px solid #ddd;padding:2px 6px;background:#f3f3f3;text-align:left" }, x))));
      for (const r of b.rows.slice(0, 50)) tb.append(h("tr", null, ...r.map((x) => h("td", { style: "border:1px solid #eee;padding:2px 6px" }, x.v === null ? "" : String(x.v)))));
      preview.append(tb);
    };
    list.addEventListener("change", show);
    show();
    const body = h("div", { style: "display:flex;flex-direction:column;gap:6px" }, h("div", null, baseName(path)), h("div", { style: "display:flex;gap:10px" }, list, preview));
    const ok = await showDialog({ title: "Navigator", body, buttons: [{ label: "Laden", value: true, primary: true }, { label: "Abbrechen", value: false }], cancel: false });
    if (!ok) return;
    const b = entries[list.selectedIndex]?.block();
    if (b) load(b);
  }

  async function fromWorkbook(): Promise<void> {
    const path = app.platform.openAnyDialog
      ? await app.platform.openAnyDialog("Daten importieren", [{ name: "Excel-Arbeitsmappen", extensions: ["xlsx", "xlsm"] }])
      : await app.platform.openFileDialog("xlsx");
    if (path) await navigator(path);
  }

  /** Aus Text/CSV: the file's data as a table on a new sheet of this workbook (Excel 365's import). */
  async function csv(path: string): Promise<void> {
    let text: string;
    try {
      const bytes = await app.platform.readFile(path);
      const utf8 = new TextDecoder("utf-8", { fatal: true });
      try { text = utf8.decode(bytes); } catch { text = new TextDecoder("windows-1252").decode(bytes); }
      if (text.charCodeAt(0) === 0xfeff) text = text.slice(1);
    } catch (e) {
      await messageBox("RaabeCalc", `Die Datei „${baseName(path)}“ kann nicht gelesen werden.\n\n${e instanceof Error ? e.message : String(e)}`);
      return;
    }
    remember(path, "csv");
    const name = baseName(path).replace(/\.(csv|txt)$/i, "");
    // Power Query finds the separator from the first line (semicolon, comma or tab).
    const first = text.split(/\r?\n/)[0] ?? "";
    const count = (ch: string) => first.split(ch).length - 1;
    const sep = [";", ",", "\t"].reduce((a, b) => (count(b) > count(a) ? b : a), ";");
    const wb = readCsv(text, { name, sep });
    const s = wb.sheets[0];
    const u = s.usedRange();
    if (!u) { await messageBox("RaabeCalc", "Die Datei enthält keine Daten."); return; }
    load(blockOf(wb, s, u, name));
  }
  async function fromText(): Promise<void> {
    const path = app.platform.openAnyDialog
      ? await app.platform.openAnyDialog("Daten importieren", [{ name: "Textdateien", extensions: ["csv", "txt"] }])
      : await app.platform.openFileDialog("xlsx");
    if (path) await csv(path);
  }

  /** Aus Tabelle/Bereich: the table of the active cell (a range becomes a table first, as in Excel) loaded as a copy. */
  async function fromTable(): Promise<void> {
    const a = app.sheet.view.active;
    let t = app.tables.tableAt(a.row, a.col);
    if (!t) {
      const sel = app.grid.selection.at(-1)!;
      const r = sel.r0 === sel.r1 && sel.c0 === sel.c1 ? app.grid.currentRegion(a.row, a.col) : sel;
      const ok = await showDialog({ title: "Tabelle erstellen", body: h("div", null, `Die Daten in ${app.sheet.name}!${r.r0 + 1}:${r.r1 + 1} werden zuerst zu einer Tabelle mit Überschriften.`),
        buttons: [{ label: "OK", value: true, primary: true }, { label: "Abbrechen", value: false }], cancel: false });
      if (!ok) return;
      t = app.tables.create(r, true) ?? undefined;
      if (!t) return;
    }
    const sheet = app.sheet;
    load(blockOf(app.wb, sheet, { ...t.ref, r1: tableData(t).r1 }, t.name));
  }

  async function recent(): Promise<void> {
    const list = loadRecent();
    const sel = h("select", { size: "10", style: "width:460px" }) as HTMLSelectElement;
    list.forEach((r, i) => sel.append(h("option", { value: String(i), title: r.path }, `${baseName(r.path)}   ${r.path}`)));
    if (list.length) sel.selectedIndex = 0;
    const body = h("div", { style: "display:flex;flex-direction:column;gap:6px" }, list.length ? sel : h("div", null, "Es wurden noch keine Quellen verwendet."));
    const ok = await showDialog({ title: "Zuletzt verwendete Quellen", body, buttons: [{ label: "Verbinden", value: true, primary: true }, { label: "Abbrechen", value: false }], cancel: false });
    if (!ok || !list.length) return;
    const r = list[sel.selectedIndex];
    if (r) await (r.kind === "csv" ? csv(r.path) : navigator(r.path));
  }

  /** The workbook's connections as the file has them (names and descriptions). */
  function connections(): { name: string; description: string }[] {
    const raw = app.wb.raw as { files?: Record<string, Uint8Array> } | undefined;
    const f = raw?.files?.["xl/connections.xml"];
    if (!f) return [];
    const xml = new TextDecoder().decode(f);
    const un = (s: string) => s.replace(/&quot;/g, '"').replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&apos;/g, "'").replace(/&amp;/g, "&");
    return [...xml.matchAll(/<connection\b([^>]*)>/g)].map((m) => ({
      name: un(/\bname="([^"]*)"/.exec(m[1])?.[1] ?? ""), description: un(/\bdescription="([^"]*)"/.exec(m[1])?.[1] ?? ""),
    })).filter((c) => c.name);
  }

  async function existing(): Promise<void> {
    const conns = connections();
    const tables = app.wb.sheets.flatMap((s) => (s.tables ?? []).map((t) => ({ s, t })));
    let page = "conn" as "conn" | "tables";
    const list = h("select", { size: "12", style: "width:460px" }) as HTMLSelectElement;
    const tabConn = h("button", { type: "button", class: "tab-btn" }, "Verbindungen");
    const tabTables = h("button", { type: "button", class: "tab-btn" }, "Tabellen");
    const fill = () => {
      list.replaceChildren();
      tabConn.style.fontWeight = page === "conn" ? "700" : "400";
      tabTables.style.fontWeight = page === "tables" ? "700" : "400";
      if (page === "conn") {
        list.append(h("optgroup", { label: "Verbindungen in dieser Arbeitsmappe" }, ...(conns.length ? conns.map((c, i) => h("option", { value: String(i), title: c.description }, c.name)) : [h("option", { disabled: true }, "<Keine Verbindungen gefunden>")])));
      } else {
        list.append(h("optgroup", { label: "Diese Arbeitsmappe" }, ...(tables.length ? tables.map((x, i) => h("option", { value: String(i) }, `${x.s.name}: ${x.t.name}`)) : [h("option", { disabled: true }, "<Keine Tabellen gefunden>")])));
      }
    };
    tabConn.addEventListener("click", () => { page = "conn"; fill(); });
    tabTables.addEventListener("click", () => { page = "tables"; fill(); });
    fill();
    const body = h("div", { style: "display:flex;flex-direction:column;gap:6px" }, h("div", { style: "display:flex;gap:6px" }, tabConn, tabTables), list);
    const ok = await showDialog({ title: "Vorhandene Verbindungen", body, buttons: [{ label: "Öffnen", value: true, primary: true }, { label: "Abbrechen", value: false }], cancel: false });
    if (!ok) return;
    const i = Number(list.value);
    if (page === "tables" && tables[i]) {
      const { s, t } = tables[i];
      load(blockOf(app.wb, s, { ...t.ref, r1: tableData(t).r1 }, t.name));
    } else if (page === "conn" && conns[i]) {
      await messageBox("RaabeCalc", `Die Verbindung „${conns[i].name}“ ist eine Abfrage. RaabeCalc behält sie in der Datei, kann sie aber nicht ausführen; die geladenen Daten stehen in der Arbeitsmappe.`);
    }
  }

  function implement(run: Run): void {
    app.commands.register(...GETDATA_COMMANDS);
    const notEditing = () => !app.grid.isEditing;
    run("data.get", () => {}, notEditing);
    run("data.fromWorkbook", () => void fromWorkbook(), notEditing);
    run("data.fromText", () => void fromText(), notEditing);
    run("data.fromTable", () => void fromTable(), notEditing);
    run("data.recentSources", () => void recent(), notEditing);
    run("data.connections", () => void existing(), notEditing);
  }

  /** The Navigator's choice without the dialog (the harness): the table or sheet `item` of the workbook at `path`. */
  async function loadItem(path: string, item: string): Promise<boolean> {
    const src = readXlsx(await app.platform.readFile(path));
    for (const s of src.sheets) for (const t of s.tables ?? []) if (t.name === item) { load(blockOf(src, s, { ...t.ref, r1: tableData(t).r1 }, t.name)); return true; }
    const s = src.sheets.find((x) => x.name === item);
    const u = s?.usedRange();
    if (!s || !u) return false;
    load(blockOf(src, s, u, s.name));
    return true;
  }

  return { implement, load, navigator, csv, connections, fromTable, loadItem, blockOf: (s: Sheet, r: Rect, name: string) => blockOf(app.wb, s, r, name) };
}
