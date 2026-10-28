// RaabeCalc: wires the shell (title bar, ribbon, backstage) to workbooks. Each open workbook is
// a document tab (D-010) shown in one grid with its sheet tabs, Namenfeld and
// Bearbeitungsleiste. Files are .xlsx (D-004), everything not edited is kept (D-024); CSV
// opens and saves as German Excel does.

import {
  MAX_COLS, MAX_ROWS, Workbook, Sheet, cellName, colName, isDateFormat, FUNCTIONS_DE, RIBBON_FORMATS, localFormat, parseRange, rangeName,
  type CellStyle, type Rect,
} from "@raabe/sheet-engine";
import { readXlsx, writeXlsx, newWorkbook } from "@raabe/ooxml";
import {
  Backstage, Commands, Keytips, Ribbon, TitleBar, baseName, h, icon, installTooltips, messageBox, showDialog, showMenu,
  type BackstageEntry, type MenuEntry, type Platform,
} from "@raabe/shell";
import { COMMANDS } from "./command-list";
import { Grid, type GridHost, type Mode } from "./grid/grid";
import { DEFAULT_COL_CHARS, charsToPx, defaultColPx, pxToPt, pxToWidth, ptToPx, shownChars } from "./grid/metrics";
import { cssFont, fitShown } from "./grid/render";
import { History } from "./ops";
import * as A from "./actions";
import { editText, formulaLocal, makeCalc, readCsv, shown, writeCsv, type CalcLike } from "./engine";
import { ribbonTabs, type RibbonContext } from "./ribbon-def";
import { applyTips } from "./tooltips";
import { CALC_TEMPLATES, templateWorkbook, type CalcTemplate } from "./templates";
import { SheetBar } from "./sheetbar";
import { CalcStatusBar, type Stats } from "./statusbar";
import { parseClipboard, rangeHtml, rangeText, type Copied } from "./clipboard";
import { addRecent, loadRecent, loadSettings, removeRecent, saveSettings, setPinned, type Settings } from "./store";
import { dialogs } from "./dialogs";
import { chartsFeature } from "./charts/feature";
import { printFeature } from "./print";
import { notesFeature } from "./notes";
import { traceFeature } from "./trace";
import { protectionFeature } from "./protection";
import { evaluateFeature } from "./evaluate";
import { errorCheckFeature } from "./errorcheck";
import { isValidEntry, validationFeature } from "./validation";
import { quickFeature, type QuickTile } from "./quick";
import { learn as learnFill, run as runFill } from "./flashfill";
import { regroup } from "./grid/outline";
import { seek } from "./goalseek";
import { spellingFeature } from "./spelling";
import { picturesFeature } from "./pictures";
import { commentsFeature } from "./comments";
import { iconRule } from "./dialogs";
import appIconUrl from "../../../branding/raabe-calc-small.svg?url";

interface Doc {
  id: string;
  name: string;
  path: string | null;
  wb: Workbook;
  history: History;
  calc: CalcLike;
  version: number;
  savedVersion: number;
  recoveryVersion: number;
  lastSaved: number | null;
}

const stripExt = (name: string) => name.replace(/\.(xlsx|csv)$/i, "");
/** A new GUID in Excel's form (Kommentar ids). */
const newId = () => {
  const b = crypto.getRandomValues(new Uint8Array(16));
  b[6] = (b[6] & 0x0f) | 0x40;
  b[8] = (b[8] & 0x3f) | 0x80;
  const x = [...b].map((v) => v.toString(16).padStart(2, "0").toUpperCase()).join("");
  return `{${x.slice(0, 8)}-${x.slice(8, 12)}-${x.slice(12, 16)}-${x.slice(16, 20)}-${x.slice(20)}}`;
};
const samePath = (a: string, b: string) => a.toLowerCase() === b.toLowerCase();

export class CalcApp {
  readonly commands = new Commands();
  private docs: Doc[] = [];
  private active!: Doc;
  private settings: Settings = loadSettings();
  private untitled = 0;
  private titleBar: TitleBar;
  private ribbon: Ribbon;
  private backstage: Backstage;
  readonly grid: Grid;
  private sheetBar: SheetBar;
  private statusBar: CalcStatusBar;
  private copied: Copied | null = null;
  private fontList: string[] = [];
  private lastFontColor = "#FF0000";
  private lastFill = "#FFFF00";
  private painter: { style: number | undefined; keep: boolean } | null = null;
  private measureCtx = document.createElement("canvas").getContext("2d")!;
  private refreshTimer = 0;
  private dlg = dialogs(this);
  readonly charts = chartsFeature(this);
  readonly print = printFeature(this);
  private notes = notesFeature(this, () => this.settings.userName);
  private trace = traceFeature(this);
  private protection = protectionFeature(this);
  private evaluator = evaluateFeature(this);
  private spelling = spellingFeature(this);
  private pictures = picturesFeature(this);
  private comments = commentsFeature(this, () => this.settings.userName);
  private errorCheck = errorCheckFeature(this, () => void this.evaluator.open());
  private quick = quickFeature(this, () => this.quickTiles());
  readonly validation = validationFeature(this, (t, text) => this.gridHost().commit(t, text, {}));
  get calc(): CalcLike { return this.active.calc; }

  constructor(readonly platform: Platform) {
    void platform.listFonts().then((f) => { this.fontList = f; }).catch(() => {});
    this.commands.register(...COMMANDS);
    this.implementCommands();
    applyTips(this.commands);

    this.grid = new Grid(this.gridHost());
    this.grid.setFunctions(FUNCTIONS_DE);
    // Livevorschau: what the list or palette entry under the mouse would do, on the selection.
    this.commands.onHover((t) => {
      if (t === null) this.grid.preview(null);
      else if (typeof t === "function") t();
    });
    // The first workbook exists before the ribbon asks for the active cell's format.
    this.untitled = 1;
    const first = this.makeDoc(newWorkbook(), "Mappe1", null);
    this.docs.push(first);
    this.active = first;
    this.titleBar = new TitleBar(this.commands, platform, {
      appIcon: appIconUrl,
      qat: [
        { cmd: "file.save", key: "1" },
        { cmd: "edit.undo", key: "2" },
        { cmd: "edit.redo", key: "3" },
      ],
      onSelectDoc: (id) => this.activate(this.docs.find((d) => d.id === id)!),
      onCloseDoc: (id) => void this.closeDoc(this.docs.find((d) => d.id === id)!),
      onNewDoc: () => this.newDoc(),
      onSearchDone: () => this.grid.focus(),
    });
    this.ribbon = new Ribbon(this.commands, ribbonTabs(this.ribbonContext()), {
      onFile: () => this.backstage.open("home"),
      onDone: () => this.grid.focus(),
      tabRowRight: [],
    });
    this.backstage = new Backstage(this.commands, this.backstageEntries());
    this.backstage.onClose(() => this.grid.focus());
    this.sheetBar = new SheetBar({
      workbook: () => this.wb,
      activate: (i) => this.activateSheet(i),
      // Arbeitsmappe schützen keeps the sheets as they are.
      add: () => { if (!this.protection.bookLocked()) this.addSheet(true); },
      rename: (i, name) => (this.protection.bookLocked() ? "Die Arbeitsmappe ist geschützt; Blätter können nicht umbenannt werden." : this.renameSheet(i, name)),
      moveSheet: (from, to, copy) => { if (!this.protection.bookLocked()) this.moveSheet(from, to, copy); },
      contextMenu: (e) => this.sheetMenu(e),
    }, this.grid.hscroll);
    this.statusBar = new CalcStatusBar(this.commands, (z) => this.setZoom(z), (v) => shown({ v }, this.wb.styles.get(0)).text);
    const main = h("div", { class: "calc-main" }, this.grid.formulaBar, this.grid.el, this.sheetBar.el);
    document.body.append(this.titleBar.el, this.ribbon.tabRow, this.ribbon.card, main, this.statusBar.el, this.backstage.el);
    document.body.classList.add("calc");
    document.body.classList.toggle("touch", this.settings.touchMode);
    installTooltips();
    new Keytips(
      (el) => el.dataset.keytipScope === "root" && el.getAttribute("role") === "tab" ? this.ribbon.activeTab : null,
      () => this.grid.focus(),
    );
    this.installKeyboard();
    main.addEventListener("wheel", (e) => {
      if (!e.ctrlKey) return;
      e.preventDefault();
      // Strg+mouse wheel zooms in steps of 15 % (recorded in Excel).
      this.setZoom(this.wb.sheet().view.zoom + (e.deltaY < 0 ? 15 : -15));
    }, { passive: false });
    platform.window.onCloseRequested(() => this.confirmQuit());
    setInterval(() => void this.autoRecover(), Math.max(10, this.settings.autoRecoverSeconds) * 1000);
    this.activate(first);
    void this.offerRecovery();
    platform.onOpenFiles((paths) => { for (const p of paths) void this.openPath(p); });
  }

  // ----- documents ------------------------------------------------------------------------------

  get wb(): Workbook { return this.active.wb; }
  get sheet(): Sheet { return this.active.wb.sheet(); }
  get history(): History { return this.active.history; }
  get ctx(): A.Ctx { return { wb: this.wb, history: this.history }; }

  private makeDoc(wb: Workbook, name: string, path: string | null): Doc {
    const history = new History(wb);
    const doc: Doc = { id: crypto.randomUUID(), name, path, wb, history, calc: makeCalc(wb), version: 0, savedVersion: 0, recoveryVersion: 0, lastSaved: null };
    history.onChange = (change) => this.afterChange(doc, change);
    doc.calc.recalc();
    return doc;
  }

  newDoc(template?: CalcTemplate): void {
    this.untitled++;
    const doc = this.makeDoc(template ? templateWorkbook(template) : newWorkbook(), `Mappe${this.untitled}`, null);
    // A new, untouched Mappe1 is replaced by the next opened file (Excel).
    this.docs.push(doc);
    this.activate(doc);
  }

  activate(doc: Doc): void {
    if (this.grid.isEditing) this.grid.commitEdit("none");
    this.active = doc;
    this.grid.setMarquee(null);
    this.copied = this.copied && this.docs.some((d) => d.wb.sheets.includes(this.copied!.sheet)) ? this.copied : null;
    this.grid.refreshLayout();
    this.grid.syncBar();
    this.sheetBar.render();
    this.charts.refresh();
    this.updateTabs();
    this.updateChrome();
    this.grid.focus();
  }

  private updateTabs(): void {
    this.titleBar.setDocs(this.docs.map((d) => ({ id: d.id, name: d.name, dirty: d.version !== d.savedVersion })), this.active.id);
    document.title = `${this.active.name} - RaabeCalc`;
    this.platform.window.setTitle(document.title);
  }

  /** Schnellanalyse's tiles per tab, doing what Excel's do (results/2026-10-28-excel-quickanalysis.json). */
  private quickTiles(): Record<string, QuickTile[]> {
    const rule = (r: Omit<import("@raabe/sheet-engine").CondRule, "priority">) => () => void this.dlg.addRule(r);
    const chart = (k: import("./charts/feature").ChartKind) => () => this.charts.insert(k);
    return {
      Formatierung: [
        { label: "Datenbalken", icon: "cond-format", run: rule({ type: "dataBar", colors: ["#638EC6"] }) },
        { label: "Farbskala", icon: "cond-format", run: rule({ type: "colorScale", colors: ["#F8696B", "#FFEB84", "#63BE7B"] }) },
        { label: "Symbolsatz", icon: "cond-format", run: rule(iconRule("3Arrows")) },
        { label: "Größer als", icon: "cond-format", run: () => { const m = this.dlg.conditionalHighlightMenu()[0]; if (typeof m === "object" && "run" in m) m.run(); } },
        { label: "Obere 10 %", icon: "cond-format", run: rule({ type: "top10", rank: 10, percent: true, format: { fill: "#FFC7CE", color: "#9C0006" } }) },
        { label: "Formatierung löschen", icon: "eraser", run: () => this.dlg.clearConditional(true) },
      ],
      Diagramme: [
        { label: "Gruppierte Säulen", icon: "chart-column", run: chart("column") },
        { label: "Gruppierte Balken", icon: "chart-bar", run: chart("bar") },
        { label: "Linie", icon: "chart-line", run: chart("line") },
        { label: "Kreis", icon: "chart-pie", run: chart("pie") },
        { label: "Punkt (XY)", icon: "chart-scatter", run: chart("scatter") },
        { label: "Weitere Diagramme", icon: "chart-recommended", run: () => void this.commands.run("insert.chartRecommended") },
      ],
      // Below the selection (a row of results) and to its right (a column), Excel's two sets.
      Ergebnisse: [
        ...(["Summe", "Durchschnitt", "Anzahl", "% Gesamt", "Laufende Summe"] as const).map((k) => ({ label: k, icon: k === "% Gesamt" ? "percent" : "autosum", run: (r: Rect) => this.quickTotals(r, k, "below") })),
        ...(["Summe", "Durchschnitt", "Anzahl", "% Gesamt", "Laufende Summe"] as const).map((k) => ({ label: k, icon: "table-insert-column", run: (r: Rect) => this.quickTotals(r, k, "right") })),
      ],
      Tabellen: [
        { label: "Tabelle", icon: "format-table", run: () => void this.dlg.formatAsTable(undefined, true) },
        { label: "Leere PivotTable", icon: "table", run: () => {}, disabled: true },
      ],
      // One sparkline per row of the selection, in the column to its right (Excel).
      Sparklines: (["line", "column", "stacked"] as const).map((t) => ({
        label: { line: "Linie", column: "Säule", stacked: "Gewinn/Verlust" }[t], icon: t === "line" ? "chart-line" : "chart-column",
        run: (r: Rect) => this.addSparklines(t, r, { r0: r.r0, c0: r.c1 + 1, r1: r.r1, c1: r.c1 + 1 }),
      })),
    };
  }

  /**
   * Schnellanalyse > Ergebnisse as Excel writes them (recorded): a text first row is the header
   * and a text first column holds the labels; below the data a bold row with the label (Summe,
   * Mittelwert, Anzahl, % Gesamtergebnis, Laufende Summe) and one formula per column, or to the
   * right a bold column with the label on top and one formula per row. Anzahl counts with
   * ANZAHL2, % Gesamt divides by the sum of all data (0,00%), Laufende Summe adds up from the
   * first data cell.
   */
  private quickTotals(r: Rect, kind: "Summe" | "Durchschnitt" | "Anzahl" | "% Gesamt" | "Laufende Summe", side: "below" | "right"): void {
    const sheet = this.sheet;
    const isText = (row: number, col: number) => typeof sheet.get(row, col)?.v === "string";
    let head = false, labels = false;
    for (let c = r.c0; c <= r.c1; c++) if (isText(r.r0, c) && r.r1 > r.r0) head = true;
    for (let row = r.r0 + (head ? 1 : 0); row <= r.r1; row++) if (isText(row, r.c0) && r.c1 > r.c0) labels = true;
    const d = { r0: r.r0 + (head ? 1 : 0), c0: r.c0 + (labels ? 1 : 0), r1: r.r1, c1: r.c1 };
    if (d.r0 > d.r1 || d.c0 > d.c1) return;
    const label = { Summe: "Summe", Durchschnitt: "Mittelwert", Anzahl: "Anzahl", "% Gesamt": "% Gesamtergebnis", "Laufende Summe": "Laufende Summe" }[kind];
    const fn = { Summe: "SUMME", Durchschnitt: "MITTELWERT", Anzahl: "ANZAHL2", "% Gesamt": "SUMME", "Laufende Summe": "SUMME" }[kind];
    const abs = (row: number, col: number) => `$${colName(col)}$${row + 1}`;
    const ref = (row: number, col: number) => `${colName(col)}${row + 1}`;
    const all = `${abs(d.r0, d.c0)}:${abs(d.r1, d.c1)}`;
    const put = (row: number, col: number, text: string) => {
      A.enter(this.ctx, sheet, { r0: row, c0: col, r1: row, c1: col }, text);
      const c = sheet.get(row, col)!;
      sheet.set(row, col, { ...c, s: this.wb.styles.derive(c.s, (st) => { st.font.bold = true; if (kind === "% Gesamt" && text.startsWith("=")) st.numFmt = "0.00%"; }) });
    };
    this.change("Schnellanalyse", () => {
      if (side === "below") {
        const row = d.r1 + 1;
        this.history.cells(sheet, { r0: row, c0: r.c0, r1: row, c1: r.c1 });
        if (labels) put(row, r.c0, label);
        for (let c = d.c0; c <= d.c1; c++) {
          const part = `${ref(d.r0, c)}:${ref(d.r1, c)}`;
          put(row, c, kind === "% Gesamt" ? `=SUMME(${part})/SUMME(${all})` : kind === "Laufende Summe" ? `=SUMME(${abs(d.r0, d.c0)}:${ref(d.r1, c)})` : `=${fn}(${part})`);
        }
      } else {
        const col = d.c1 + 1;
        this.history.cells(sheet, { r0: r.r0, c0: col, r1: r.r1, c1: col });
        if (head) put(r.r0, col, label);
        for (let row = d.r0; row <= d.r1; row++) {
          const part = `${ref(row, d.c0)}:${ref(row, d.c1)}`;
          put(row, col, kind === "% Gesamt" ? `=SUMME(${part})/SUMME(${all})` : kind === "Laufende Summe" ? `=SUMME(${abs(d.r0, d.c0)}:${ref(row, d.c1)})` : `=${fn}(${part})`);
        }
      }
    });
  }

  /** Ungültige Daten einkreisen, per sheet; a cell loses its circle once it holds a valid entry. */
  private circles = new Map<Sheet, { row: number; col: number }[]>();

  private circleInvalid(): void {
    const sheet = this.sheet, si = this.wb.sheets.indexOf(sheet);
    const out: { row: number; col: number }[] = [];
    for (const v of sheet.validations ?? []) for (const r of v.ref) {
      for (const [key, cell] of sheet.cells) {
        const row = Math.floor(key / 16384), col = key % 16384;
        if (row < r.r0 || row > r.r1 || col < r.c0 || col > r.c1 || cell.v === null || cell.v === "") continue;
        const st = this.wb.styles.get(cell.s);
        if (!isValidEntry(this.wb, si, row, col, editText(cell, st))) out.push({ row, col });
      }
    }
    // Excel circles at most 255 cells.
    this.circles.set(sheet, out.slice(0, 255));
    this.grid.contentChanged();
  }

  /**
   * The menu after a drag with the right mouse button, with Excel's entries (recorded,
   * probe_excel_rightdrag.py): for the fill handle the kinds of filling, for the border where and
   * how the cells go.
   */
  private rightDragMenu(kind: "fill" | "move", source: Rect, target: Rect, at: { x: number; y: number }): void {
    if (this.protection.on()) return;
    const anchor = new DOMRect(at.x, at.y, 0, 0);
    // After a choice the target is selected (Excel); Abbrechen leaves everything as it was.
    const show = (entries: MenuEntry[]) => showMenu(this.commands, entries.map((m) => (typeof m === "object" && "run" in m && m.label !== "Abbrechen"
      ? { ...m, run: () => { m.run(); this.grid.select(target, { row: target.r0, col: target.c0 }); } } : m)), anchor);
    if (kind === "fill") {
      const v = this.sheet.get(source.r0, source.c0);
      const isDate = typeof v?.v === "number" && isDateFormat(this.wb.styles.get(v.s).numFmt);
      const fill = (how: "copy" | "series" | "formats" | "values") => () => this.fillWithOptions(source, target, how);
      const dates = (unit: "days" | "workdays" | "months" | "years") => () => this.change("AutoAusfüllen", () => A.fillDates(this.ctx, this.sheet, source, target, unit));
      const trend = (exp: boolean) => () => this.change("AutoAusfüllen", () => A.fillTrend(this.ctx, this.sheet, source, target, exp));
      const numbers = typeof v?.v === "number";
      show([
        { label: "Zellen kopieren", run: fill("copy") },
        { label: "Datenreihe ausfüllen", run: fill("series") },
        { label: "Nur Formate ausfüllen", run: fill("formats") },
        { label: "Ohne Formatierung ausfüllen", run: fill("values") },
        "-",
        { label: "Tage ausfüllen", run: dates("days"), disabled: !isDate },
        { label: "Arbeitstage ausfüllen", run: dates("workdays"), disabled: !isDate },
        { label: "Monate ausfüllen", run: dates("months"), disabled: !isDate },
        { label: "Jahre ausfüllen", run: dates("years"), disabled: !isDate },
        "-",
        { label: "Linearer Trend", run: trend(false), disabled: !numbers },
        { label: "Exponentieller Trend", run: trend(true), disabled: !numbers },
        { label: "Blitzvorschau", run: () => void this.flashFill() },
        { label: "Reihe...", run: () => void this.dlg.series() },
      ]);
      return;
    }
    const to = { row: target.r0, col: target.c0 };
    const copyAs = (what: "all" | "values" | "formats") => () => this.change("Kopieren", () => { A.copyCells(this.ctx, { sheet: this.sheet, rect: source }, this.sheet, to, what); });
    const shifted = (cut: boolean, rows: boolean) => () => {
      if (cut) { this.gridHost().moveInsert!(source, to, rows); return; }
      this.change("Kopieren", () => {
        A.shiftCells(this.ctx, this.sheet, target, rows ? "down" : "right");
        A.copyCells(this.ctx, { sheet: this.sheet, rect: source }, this.sheet, to);
      });
    };
    show([
      { label: "Hierhin verschieben", run: () => this.gridHost().move(source, to, false) },
      { label: "Hierhin kopieren", run: copyAs("all") },
      { label: "Hierhin nur als Werte kopieren", run: copyAs("values") },
      { label: "Hierhin nur als Format kopieren", run: copyAs("formats") },
      { label: "Verknüpfung hier erstellen", run: () => this.change("Verknüpfung", () => A.linkCells(this.ctx, this.sheet, source, to)) },
      { label: "Link hier erstellen", run: () => this.change("Link", () => A.hyperlinkCells(this.ctx, this.sheet, source, to)) },
      "-",
      { label: "Kopieren und nach unten verschieben", run: shifted(false, true) },
      { label: "Kopieren und nach rechts verschieben", run: shifted(false, false) },
      { label: "Ausschneiden und nach unten verschieben", run: shifted(true, true) },
      { label: "Ausschneiden und nach rechts verschieben", run: shifted(true, false) },
      "-",
      { label: "Abbrechen", run: () => {} },
    ]);
  }

  /**
   * Blitzvorschau (Strg+E): the examples typed in the active cell's column, the texts in the
   * columns to its left as sources (as the cells show them), the rest of the column filled where
   * the sources have text; a first row that does not fit is taken as the header.
   */
  private async flashFill(): Promise<void> {
    if (this.grid.isEditing) this.grid.commitEdit("none");
    const sheet = this.sheet;
    const { row, col } = sheet.view.active;
    const region = this.grid.currentRegion(row, col);
    if (col === 0 || region.c0 >= col) { await messageBox("RaabeCalc", "Blitzvorschau braucht links neben der Spalte Daten, aus denen sich die Einträge ableiten lassen."); return; }
    const text = (r: number, c: number) => { const cell = sheet.get(r, c); return cell?.v === null || cell?.v === undefined ? "" : shown(cell, this.wb.styles.get(cell.s)).text; };
    const srcCols: number[] = [];
    for (let c = region.c0; c < col; c++) srcCols.push(c);
    const rows: number[] = [];
    for (let r = region.r0; r <= region.r1; r++) rows.push(r);
    const sources = (r: number) => srcCols.map((c) => text(r, c));
    const examples = rows.filter((r) => text(r, col) !== "");
    let prog = learnFill(examples.map((r) => ({ sources: sources(r), output: text(r, col) })));
    let body = rows;
    if (!prog && examples[0] === region.r0) {
      prog = learnFill(examples.slice(1).map((r) => ({ sources: sources(r), output: text(r, col) })));
      body = rows.slice(1);
    }
    if (!prog) { await messageBox("RaabeCalc", "Blitzvorschau hat kein Muster gefunden. Geben Sie ein oder zwei weitere Beispiele ein und versuchen Sie es noch einmal."); return; }
    const fills = body.filter((r) => text(r, col) === "" && sources(r).some((s) => s !== ""));
    if (!fills.length) return;
    if (!this.protection.guardRects(fills.map((r) => ({ r0: r, c0: col, r1: r, c1: col })))) return;
    this.change("Blitzvorschau", () => {
      for (const r of fills) {
        const out = runFill(prog!, sources(r));
        if (out !== null) A.enter(this.ctx, sheet, { r0: r, c0: col, r1: r, c1: col }, out);
      }
    });
  }

  /**
   * Gruppieren / Gruppierung aufheben: whole rows or columns directly, otherwise Excel's question
   * whether the rows or the columns of the selection are meant.
   */
  private async group(delta: 1 | -1): Promise<void> {
    const r = this.grid.selection.at(-1)!;
    let axis: "rows" | "cols" | null = r.c0 === 0 && r.c1 === MAX_COLS - 1 ? "rows" : r.r0 === 0 && r.r1 === MAX_ROWS - 1 ? "cols" : null;
    if (!axis) {
      const rows = h("input", { type: "radio", name: "grp", checked: true }) as HTMLInputElement;
      const body = h("div", null, h("label", { class: "check" }, rows, "Zeilen"), h("label", { class: "check" }, h("input", { type: "radio", name: "grp" }), "Spalten"));
      const ok = await showDialog({ title: delta > 0 ? "Gruppieren" : "Gruppierung aufheben", body, buttons: [{ label: "OK", value: true, primary: true }, { label: "Abbrechen", value: false }], cancel: false });
      if (!ok) return;
      axis = rows.checked ? "rows" : "cols";
    }
    const [a, b] = axis === "rows" ? [r.r0, r.r1] : [r.c0, r.c1];
    this.change(delta > 0 ? "Gruppieren" : "Gruppierung aufheben", () => {
      this.history.whole(this.sheet);
      regroup(this.sheet, axis!, a, b, delta);
    });
  }

  /**
   * Zielwertsuche: Zielzelle (a formula), Zielwert, Veränderbare Zelle (a value); the search runs
   * on the workbook itself, and Abbrechen in the result box puts the old value back (Excel).
   */
  private async goalSeek(): Promise<void> {
    const sheet = this.sheet, si = this.wb.sheets.indexOf(sheet);
    const a = sheet.view.active;
    const target = h("input", { type: "text", value: cellName(a.row, a.col), style: "width:120px" }) as HTMLInputElement;
    const goal = h("input", { type: "text", style: "width:120px" }) as HTMLInputElement;
    const changing = h("input", { type: "text", style: "width:120px" }) as HTMLInputElement;
    const body = h("div", { class: "goalseek" },
      h("label", { class: "field" }, "Zielzelle:", target), h("label", { class: "field" }, "Zielwert:", goal), h("label", { class: "field" }, "Veränderbare Zelle:", changing));
    setTimeout(() => goal.focus());
    const ok = await showDialog({ title: "Zielwertsuche", body, buttons: [{ label: "OK", value: true, primary: true }, { label: "Abbrechen", value: false }], cancel: false });
    if (!ok) return;
    const t = parseRange(target.value.replace(/\$/g, "")), c = parseRange(changing.value.replace(/\$/g, ""));
    const g = Number(goal.value.trim().replace(/\./g, "").replace(",", "."));
    if (!t || t.r0 !== t.r1 || t.c0 !== t.c1 || sheet.get(t.r0, t.c0)?.f === undefined) { await messageBox("RaabeCalc", "Die Zielzelle muss eine Formel enthalten."); return; }
    if (!c || c.r0 !== c.r1 || c.c0 !== c.c1 || sheet.get(c.r0, c.c0)?.f !== undefined) { await messageBox("RaabeCalc", "Die veränderbare Zelle muss einen Wert enthalten, keine Formel."); return; }
    if (!goal.value.trim() || !Number.isFinite(g)) { await messageBox("RaabeCalc", "Der Zielwert muss eine Zahl sein."); return; }
    const old = sheet.get(c.r0, c.c0);
    const start = typeof old?.v === "number" ? old.v : 0;
    const calc = this.active.calc;
    const f = (x: number) => {
      sheet.set(c.r0, c.c0, { ...(old ?? { v: null }), v: x });
      calc.changed(si, [{ row: c.r0, col: c.c0 }]);
      const v = calc.value(si, t.r0, t.c0);
      return typeof v === "number" ? v : NaN;
    };
    const r = seek(f, g, start);
    // The search's own changes go; the result is one undo step.
    sheet.set(c.r0, c.c0, old);
    calc.changed(si, [{ row: c.r0, col: c.c0 }]);
    this.change("Zielwertsuche", () => {
      this.history.cell(sheet, c.r0, c.c0);
      sheet.set(c.r0, c.c0, { ...(old ?? { v: null }), v: r.x });
    });
    const shownY = shown(sheet.get(t.r0, t.c0), this.wb.styles.get(sheet.get(t.r0, t.c0)?.s)).text;
    const keep = await showDialog({ title: "Status der Zielwertsuche", body: h("div", null,
      h("p", null, r.found ? `Für Zelle ${cellName(t.r0, t.c0)} wurde ein Wert gefunden, der das Ziel erreicht.` : `Für Zelle ${cellName(t.r0, t.c0)} wurde kein Wert gefunden, der das Ziel erreicht.`),
      h("div", { class: "field" }, "Zielwert:", h("span", null, goal.value)), h("div", { class: "field" }, "Aktueller Wert:", h("span", null, shownY))),
      buttons: [{ label: "OK", value: true, primary: true }, { label: "Abbrechen", value: false }], cancel: false });
    if (!keep) void this.commands.run("edit.undo");
  }

  /** Teilergebnisse: group column, function, columns to total, replace or remove (Excel's dialog). */
  private async subtotalDialog(): Promise<void> {
    const sheet = this.sheet;
    const a = sheet.view.active;
    const region = this.grid.currentRegion(a.row, a.col);
    if (region.r1 <= region.r0) { await messageBox("RaabeCalc", "Markieren Sie eine Zelle in einer Liste mit Überschriften, um Teilergebnisse einzufügen."); return; }
    const names: string[] = [];
    for (let c = region.c0; c <= region.c1; c++) {
      const cell = sheet.get(region.r0, c);
      names.push(cell?.v !== null && cell?.v !== undefined ? shown(cell, this.wb.styles.get(cell.s)).text : `Spalte ${colName(c)}`);
    }
    const by = h("select", null, ...names.map((n, i) => h("option", { value: String(i) }, n))) as HTMLSelectElement;
    const fn = h("select", null, ...A.SUBTOTALS.map((f) => h("option", { value: f.name }, f.name))) as HTMLSelectElement;
    const boxes = names.map((n, i) => { const input = h("input", { type: "checkbox", checked: i === names.length - 1 || undefined }) as HTMLInputElement; return { input, el: h("label", { class: "check" }, input, n) }; });
    const replace = h("input", { type: "checkbox", checked: true }) as HTMLInputElement;
    const breaks = h("input", { type: "checkbox" }) as HTMLInputElement;
    const below = h("input", { type: "checkbox", checked: true, disabled: true }) as HTMLInputElement;
    const body = h("div", { class: "subtotal-dialog" },
      h("div", { class: "fc-label" }, "Gruppieren nach:"), by,
      h("div", { class: "fc-label" }, "Unter Verwendung von:"), fn,
      h("div", { class: "fc-label" }, "Teilergebnis addieren zu:"), h("div", { class: "protect-list" }, ...boxes.map((b) => b.el)),
      h("label", { class: "check" }, replace, "Vorhandene Teilergebnisse ersetzen"),
      h("label", { class: "check" }, breaks, "Seitenumbrüche zwischen Gruppen einfügen"),
      h("label", { class: "check" }, below, "Ergebnisse unterhalb der Daten anzeigen"));
    const choice = await showDialog<"ok" | "remove" | null>({ title: "Teilergebnisse", body, buttons: [{ label: "Alle entfernen", value: "remove", left: true }, { label: "OK", value: "ok", primary: true }, { label: "Abbrechen", value: null }], cancel: null });
    if (!choice) return;
    this.change("Teilergebnis", () => {
      let r = region;
      if (choice === "remove" || replace.checked) r = A.removeSubtotals(this.ctx, sheet, r);
      if (choice === "remove") return;
      const cols = boxes.map((b, i) => (b.input.checked ? region.c0 + i : -1)).filter((c) => c >= 0);
      if (!cols.length) return;
      const done = A.subtotal(this.ctx, sheet, r, region.c0 + Number(by.value), A.SUBTOTALS.find((f) => f.name === fn.value)!, cols);
      if (breaks.checked) {
        // A new page after each group's result row (not after the last one).
        const at: number[] = [];
        for (let row = done.r0 + 1; row < done.r1 - 1; row++) if (sheet.rows.get(row)?.outline === 1) at.push(row + 1);
        sheet.page = { ...(sheet.page ?? {}), rowBreaks: [...new Set([...(sheet.page?.rowBreaks ?? []), ...at])].sort((a, b) => a - b) };
      }
    });
  }

  private breakMemo: { key: string; layout: ReturnType<CalcApp["print"]["layout"]> } | null = null;

  /** Sheets shown in Umbruchvorschau (Excel keeps the view per sheet). */
  private breakView = new Set<Sheet>();

  /** Normal / Umbruchvorschau; Excel shows the preview at 60 % the first time. */
  private setPageBreakView(on: boolean): void {
    if (on === this.breakView.has(this.sheet)) return;
    if (on) {
      this.breakView.add(this.sheet);
      if (this.sheet.view.zoom === 100) this.sheet.view.zoom = 60;
    } else {
      this.breakView.delete(this.sheet);
      if (this.sheet.view.zoom === 60) this.sheet.view.zoom = 100;
    }
    this.grid.refreshLayout();
    this.grid.contentChanged();
    this.scheduleChrome();
  }

  /** Blatt schützen: whether the active sheet is protected, and whether it allows an action ("objects" …). */
  isProtected(): boolean { return this.protection.on(); }
  protectionAllows(what: string): boolean { return this.protection.allows(this.sheet, what); }

  /** Sparklines erstellen: Datenbereich and Positionsbereich (a row or column of cells). */
  private async sparklineDialog(type: "line" | "column" | "stacked"): Promise<void> {
    const sel = this.grid.selection.at(-1)!;
    const data = h("input", { type: "text", value: rangeName(sel), style: "width:180px" }) as HTMLInputElement;
    const place = h("input", { type: "text", value: rangeName({ r0: sel.r0, c0: sel.c1 + 1, r1: sel.r1, c1: sel.c1 + 1 }), style: "width:180px" }) as HTMLInputElement;
    const body = h("div", null, h("label", { class: "field" }, "Datenbereich:", data), h("label", { class: "field" }, "Positionsbereich:", place));
    const ok = await showDialog({ title: "Sparklines erstellen", body, buttons: [{ label: "OK", value: true, primary: true }, { label: "Abbrechen", value: false }], cancel: false });
    if (!ok) return;
    const d = parseRange(data.value.replace(/\$/g, "")), p = parseRange(place.value.replace(/\$/g, ""));
    if (!d || !p) { await messageBox("RaabeCalc", "Der Daten- oder Positionsbereich ist ungültig."); return; }
    if (!this.addSparklines(type, d, p)) await messageBox("RaabeCalc", "Der Positionsbereich muss eine Zeile oder Spalte sein, die zu den Daten passt.");
  }

  /** One sparkline per row (or column) of `data`, placed in the cells of `place`. */
  private addSparklines(type: "line" | "column" | "stacked", data: Rect, place: Rect): boolean {
    const cells: { row: number; col: number }[] = [];
    for (let r = place.r0; r <= place.r1; r++) for (let c = place.c0; c <= place.c1; c++) cells.push({ row: r, col: c });
    const byRow = place.c0 === place.c1 && cells.length === data.r1 - data.r0 + 1;
    const byCol = !byRow && place.r0 === place.r1 && cells.length === data.c1 - data.c0 + 1;
    if (!byRow && !byCol) return false;
    const sheet = this.sheet;
    const q = /^[A-Za-z_][A-Za-z0-9_.]*$/.test(sheet.name) ? sheet.name : `'${sheet.name.replace(/'/g, "''")}'`;
    const items = cells.map((c, i) => ({ ...c, f: `${q}!${rangeName(byRow ? { ...data, r0: data.r0 + i, r1: data.r0 + i } : { ...data, c0: data.c0 + i, c1: data.c0 + i })}` }));
    this.change("Sparklines", () => {
      this.history.whole(sheet);
      // A cell holds one sparkline: new ones replace those in their cells.
      const taken = new Set(items.map((i) => `${i.row},${i.col}`));
      const kept = (sheet.sparklines ?? []).map((g) => ({ ...g, items: g.items.filter((i) => !taken.has(`${i.row},${i.col}`)) })).filter((g) => g.items.length);
      sheet.sparklines = [...kept, { type, color: "#376092", negative: "#D00000", items }];
    });
    return true;
  }

  /** Changes to the active workbook so far (features cache what they read from the cells by it). */
  get version(): number { return this.active.version; }

  /** Zahl als Text gespeichert > In eine Zahl umwandeln for one cell. */
  toNumber(row: number, col: number): void {
    const c = this.sheet.get(row, col);
    if (c && typeof c.v === "string" && c.f === undefined) this.change("In eine Zahl umwandeln", () => A.enter(this.ctx, this.sheet, { r0: row, c0: col, r1: row, c1: col }, String(c.v).trim()));
  }

  private afterChange(doc: Doc, change?: import("./ops").Change): void {
    doc.version++;
    if (doc.wb.calcMode === "manual") {
      // Manuell: no recalculation now; the status bar shows that one is due (Excel's "Berechnen").
      if (doc === this.active) this.statusBar.setFlags([["Berechnen", "Berechnen"]]);
      // Formulas typed now are worked out (Excel computes the entered formula itself); the rest waits.
      const cells = new Map<Sheet, { row: number; col: number }[]>();
      for (const [sh, list] of change?.cells ?? []) cells.set(sh, list.filter((p) => sh.get(p.row, p.col)?.f !== undefined));
      change = { structural: false, cells };
    }
    // Excel removes the trace arrows on any change; a circled cell that is now valid loses its circle.
    this.trace.clear();
    for (const [sheet, list] of this.circles) {
      const si = doc.wb.sheets.indexOf(sheet);
      if (si < 0) continue;
      this.circles.set(sheet, list.filter(({ row, col }) => { const c = sheet.get(row, col); return !!c && c.v !== null && c.v !== "" && !isValidEntry(doc.wb, si, row, col, editText(c, doc.wb.styles.get(c.s))); }));
    }
    queueMicrotask(() => this.validation.afterSelect());
    // Only what depends on the changed cells is calculated again; structural changes (rows,
    // columns, sheets, names) recalculate the whole workbook.
    if (!change || change.structural) doc.calc.recalc();
    else for (const [sheet, cells] of change.cells) {
      const i = doc.wb.sheets.indexOf(sheet);
      if (i >= 0) doc.calc.changed(i, cells);
    }
    if (doc === this.active) {
      this.grid.refreshLayout();
      this.grid.contentChanged();
      this.sheetBar.render();
      this.charts.refresh();
    }
    this.updateTabs();
    this.updateChrome();
  }

  /** Runs a change as one undo step on the active workbook. */
  change(label: string, fn: () => void): void {
    if (this.grid.isEditing) this.grid.commitEdit("none");
    this.history.step(label, fn);
  }

  private activateSheet(i: number): void {
    if (this.grid.isEditing) this.grid.commitEdit("none");
    if (i === this.wb.active) return;
    this.wb.active = i;
    this.grid.refreshLayout();
    this.grid.syncBar();
    this.sheetBar.render();
    this.charts.refresh();
    this.updateChrome();
    this.validation.afterSelect();
  }

  // ----- the grid's host ------------------------------------------------------------------------

  private gridHost(): GridHost {
    return {
      workbook: () => this.wb,
      commit: (targets, text) => {
        if (!this.protection.guardRects([targets])) return;
        this.change("Eingabe", () => {
          A.enter(this.ctx, this.sheet, targets, text);
          this.autoRowHeights(targets);
          this.autoWiden([targets]);
        });
        this.grid.setMarquee(null);
        this.copied = null;
      },
      clearContents: () => { if (this.protection.guard("edit.clearContents")) this.change("Löschen", () => A.clearContents(this.ctx, this.sheet, this.grid.selection)); },
      clearRange: (r) => { if (this.protection.guard("edit.clearContents")) this.change("Löschen", () => A.clearContents(this.ctx, this.sheet, [r])); },
      fill: (source, target, ctrl) => { if (this.protection.guardRects([target])) this.fillWithOptions(source, target, ctrl ? "toggle" : "auto"); },
      move: (source, to, copy) => this.protection.on() ? void this.protection.canEdit(this.sheet.view.active.row, this.sheet.view.active.col) : this.change(copy ? "Kopieren" : "Verschieben", () => {
        if (copy) A.copyCells(this.ctx, { sheet: this.sheet, rect: source }, this.sheet, to);
        else A.moveCells(this.ctx, this.sheet, source, to);
      }),
      rightDrag: (kind, source, target, at) => this.rightDragMenu(kind, source, target, at),
      moveInsert: (source, at, rows) => {
        if (this.protection.on()) return undefined;
        let dst: Rect | undefined;
        this.change("Verschieben", () => { dst = A.moveInsert(this.ctx, this.sheet, source, at, rows); });
        return dst;
      },
      resizeCols: (cols, px) => { if (this.protection.enabled("cells.colWidth")) this.setColWidths(cols, px); },
      resizeRows: (rows, px) => { if (this.protection.enabled("cells.rowHeight")) this.setRowHeights(rows, px); },
      autofitCols: (cols) => { if (this.protection.enabled("cells.autofitCols")) this.autofitCols(cols); },
      autofitRows: (rows) => !this.protection.enabled("cells.autofitRows") ? undefined : this.change("Zeilenhöhe", () => this.autoRowHeights({ r0: Math.min(...rows), c0: 0, r1: Math.max(...rows), c1: MAX_COLS - 1 }, true)),
      selectionChanged: () => {
        this.charts.deselect();
        this.validation.afterSelect();
        if (this.painter) this.applyPainter();
        this.scheduleChrome();
      },
      flagsChanged: (flags) => this.statusBar.setFlags(flags),
      copyMode: (on) => this.statusBar.setMessage(on ? "Ziel wählen, mit Eingabe oder Strg+V einfügen" : null),
      modeChanged: (mode: Mode) => {
        this.statusBar.setMode(mode === "ready" ? "Bereit" : mode === "enter" ? "Eingeben" : mode === "edit" ? "Bearbeiten" : "Zeigen");
        this.commands.refresh();
      },
      contextMenu: (e, kind) => this.contextMenu(e, kind),
      key: (e) => this.shortcut(e),
      filterMenu: (col, anchor) => void this.dlg.filterMenu(col, anchor),
      namesMenu: (anchor) => {
        const names = this.wb.names.filter((n) => !n.name.startsWith("_xlnm") && !n.hidden);
        showMenu(this.commands, names.length ? names.map((n) => ({ label: n.name, run: () => { this.grid.goTo(n.name); this.grid.focus(); } }))
          : [{ label: "(keine Namen)", run: () => {}, disabled: true }], anchor);
      },
      followLink: (target) => {
        if (target.startsWith("#")) {
          const m = /^#(?:'?([^'!]+)'?!)?(.+)$/.exec(target);
          const i = m?.[1] ? this.wb.sheets.findIndex((s) => s.name.toLowerCase() === m[1].toLowerCase()) : this.wb.active;
          if (i >= 0) { this.activateSheet(i); this.grid.goTo(m?.[2] ?? "A1"); }
        } else void this.platform.openExternal?.(target).catch(() => {});
      },
      enterPaste: () => {
        if (!this.copied) return false;
        void this.paste().then(() => { this.copied = null; this.grid.setMarquee(null); this.grid.showOptions(null); });
        return true;
      },
      errorMenu: (kind, anchor) => {
        const { row, col } = this.sheet.view.active;
        const ignore = () => this.change("Fehler ignorieren", () => { this.history.cell(this.sheet, row, col); const c = this.sheet.get(row, col)!; this.sheet.set(row, col, { ...c, ignoreError: true }); });
        const entries: MenuEntry[] = kind === "text-number"
          ? [{ heading: "Zahl als Text gespeichert" }, { label: "In eine Zahl umwandeln", run: () => this.change("In eine Zahl umwandeln", () => {
              for (const r of this.grid.selection) for (let rr = r.r0; rr <= Math.min(r.r1, r.r0 + 100000); rr++) for (let cc = r.c0; cc <= Math.min(r.c1, r.c0 + 200); cc++) {
                const c = this.sheet.get(rr, cc);
                if (c && typeof c.v === "string" && c.f === undefined) A.enter(this.ctx, this.sheet, { r0: rr, c0: cc, r1: rr, c1: cc }, c.v.trim());
              }
            }) }, { label: "Fehler ignorieren", run: ignore }, { label: "In Bearbeitungsleiste bearbeiten", run: () => (document.querySelector<HTMLTextAreaElement>(".fx-input")!).focus() }]
          : [{ heading: "Fehler in der Formel" }, { label: "Fehler ignorieren", run: ignore }, { label: "In Bearbeitungsleiste bearbeiten", run: () => (document.querySelector<HTMLTextAreaElement>(".fx-input")!).focus() }];
        showMenu(this.commands, entries, anchor);
      },
      showFormulas: () => this.showFormulas,
      arrows: () => this.trace.of(this.sheet),
      movePageBreak: (rows, from, to) => this.protection.on() ? undefined : this.change("Seitenumbruch", () => {
        // A moved break is a manual one at its new place (Excel); an automatic one leaves nothing behind.
        this.history.whole(this.sheet);
        const p = { ...(this.sheet.page ?? {}) };
        const list = new Set((rows ? p.rowBreaks : p.colBreaks) ?? []);
        list.delete(from);
        list.add(to);
        const sorted = [...list].sort((a, b) => a - b);
        if (rows) p.rowBreaks = sorted; else p.colBreaks = sorted;
        this.sheet.page = p;
      }),
      pageBreaks: () => {
        if (!this.breakView.has(this.sheet)) return null;
        // Paginating walks every row: kept until the workbook changes.
        const key = `${this.wb.sheets.indexOf(this.sheet)}|${this.version}`;
        if (this.breakMemo?.key !== key) this.breakMemo = { key, layout: this.print.layout(this.sheet) };
        const l = this.breakMemo.layout;
        if (!l.area) return null;
        return { area: l.area, pages: l.pages, manualRows: this.sheet.page?.rowBreaks ?? [], manualCols: this.sheet.page?.colBreaks ?? [] };
      },
      outline: (_axis, fn) => this.change("Gliederung", () => { this.history.whole(this.sheet); fn(); }),
      circles: () => this.circles.get(this.sheet) ?? [],
      canEdit: (row, col) => this.protection.canEdit(row, col),
      validate: (row, col, text) => (this.validation.quickCheck(row, col, text) ? "ok" : this.validation.check(row, col, text)),
      painted: () => { this.pictures.place(); this.validation.place(); this.quick.place(); },
      nextUnlocked: (row, col, back) => this.protection.nextUnlocked(row, col, back),
    };
  }

  /**
   * AutoAusfüllen with Excel's choices afterwards (AutoAusfülloptionen): Zellen kopieren, Datenreihe
   * ausfüllen, Nur Formate ausfüllen, Ohne Formatierung ausfüllen. A choice undoes the fill and
   * fills again.
   */
  private fillWithOptions(source: Rect, target: Rect, how: "auto" | "toggle" | "copy" | "series" | "formats" | "values"): void {
    const single = source.r0 === source.r1 && source.c0 === source.c1;
    const c = this.sheet.get(source.r0, source.c0);
    const numeric = single && typeof c?.v === "number" && c.f === undefined && !isDateStyle(this.wb.styles.get(c.s));
    const opts =
      how === "copy" ? { copy: true } : how === "series" ? { series: true } : how === "formats" ? { formatsOnly: true } : how === "values" ? { valuesOnly: true }
      : numeric ? { series: how === "toggle" } : { copy: how === "toggle" && !single ? true : how === "toggle", series: how !== "toggle" && single && isDateStyle(this.wb.styles.get(c?.s)) };
    this.change("Ausfüllen", () => A.autoFill(this.ctx, this.sheet, source, target, opts));
    this.grid.showOptions(target, "fill", (anchor) => showMenu(this.commands, [
      { label: "Zellen kopieren", run: () => { this.history.undo(); this.fillWithOptions(source, target, "copy"); } },
      { label: "Datenreihe ausfüllen", run: () => { this.history.undo(); this.fillWithOptions(source, target, "series"); } },
      "-",
      { label: "Nur Formate ausfüllen", run: () => { this.history.undo(); this.fillWithOptions(source, target, "formats"); } },
      { label: "Ohne Formatierung ausfüllen", run: () => { this.history.undo(); this.fillWithOptions(source, target, "values"); } },
    ], anchor));
  }

  // ----- sizes ----------------------------------------------------------------------------------

  setColWidths(cols: number[], px: number): void {
    this.change("Spaltenbreite", () => {
      for (const c of cols) {
        this.history.col(this.sheet, c);
        const info = this.sheet.cols.get(c) ?? {};
        if (px <= 0) this.sheet.cols.set(c, { ...info, hidden: true });
        else this.sheet.cols.set(c, { ...info, hidden: false, width: pxToWidth(px), bestFit: false });
      }
    });
  }

  setRowHeights(rows: number[], px: number): void {
    this.change("Zeilenhöhe", () => {
      for (const r of rows) {
        this.history.row(this.sheet, r);
        const info = this.sheet.rows.get(r) ?? {};
        if (px <= 0) this.sheet.rows.set(r, { ...info, hidden: true });
        else this.sheet.rows.set(r, { ...info, hidden: false, height: pxToPt(px), custom: true });
      }
    });
  }

  /** Text width of what a cell shows, in px at 100 % (for AutoAnpassen). */
  private textWidth(row: number, col: number): number {
    const c = this.sheet.get(row, col);
    if (!c || c.v === null) return 0;
    const st = this.wb.styles.get(c.s ?? this.sheet.cols.get(col)?.s);
    this.measureCtx.font = cssFont(st, 1);
    const t = shown(c, st).text;
    return Math.max(...t.split("\n").map((l) => this.measureCtx.measureText(l).width));
  }

  autofitCols(cols: number[]): void {
    const set = new Set(cols);
    const widths = new Map<number, number>();
    for (const [k, c] of this.sheet.cells) {
      const col = k % MAX_COLS;
      if (!set.has(col) || c.v === null) continue;
      const row = Math.floor(k / MAX_COLS);
      if (this.sheet.mergeAt(row, col)) continue;
      widths.set(col, Math.max(widths.get(col) ?? 0, this.textWidth(row, col)));
    }
    this.change("Spaltenbreite", () => {
      for (const c of cols) {
        const w = widths.get(c);
        if (w === undefined) continue;
        this.history.col(this.sheet, c);
        const px = Math.ceil(w) + 7;
        this.sheet.cols.set(c, { ...(this.sheet.cols.get(c) ?? {}), width: pxToWidth(px), hidden: false, bestFit: true });
      }
    });
  }

  /** Rows without their own height grow with larger fonts and wrapped text (Excel). */
  private autoRowHeights(r: Rect, force = false): void {
    const sheet = this.sheet;
    const rows = new Set<number>();
    for (let row = r.r0; row <= Math.min(r.r1, r.r0 + 5000); row++) {
      if (!force && sheet.rows.get(row)?.custom) continue;
      rows.add(row);
    }
    if (r.r1 - r.r0 > 5000) return;
    const need = new Map<number, number>();
    // Small ranges look at their own cells; whole rows look through the sheet's cells.
    const area = (r.r1 - r.r0 + 1) * (Math.min(r.c1, MAX_COLS - 1) - r.c0 + 1);
    const entries: [number, import("@raabe/sheet-engine").Cell][] = [];
    const partial = area <= 5000 && r.c1 - r.c0 < 1000;
    if (partial) {
      for (let row = r.r0; row <= r.r1; row++) for (let col = r.c0; col <= r.c1; col++) { const c = sheet.get(row, col); if (c) entries.push([row * MAX_COLS + col, c]); }
    } else for (const e of sheet.cells) entries.push(e);
    for (const [k, c] of entries) {
      const row = Math.floor(k / MAX_COLS);
      if (!rows.has(row)) continue;
      const col = k % MAX_COLS;
      const st = this.wb.styles.get(c.s);
      if (c.v === null && st.font.size <= 11) continue;
      let lines = 1;
      if (st.wrap && typeof c.v === "string") {
        this.measureCtx.font = cssFont(st, 1);
        const cw = sheet.cols.get(col)?.width;
        const w = (cw !== undefined ? charsToPx(cw) : defaultColPx(sheet.defaultColWidth ?? DEFAULT_COL_CHARS)) - 5;
        let n = 0;
        for (const para of c.v.split("\n")) {
          let line = "";
          n++;
          for (const word of para.split(" ")) {
            const t = line ? line + " " + word : word;
            if (this.measureCtx.measureText(t).width > w && line) { n++; line = word; } else line = t;
          }
        }
        lines = n;
      } else if (typeof c.v === "string" && c.v.includes("\n")) lines = c.v.split("\n").length;
      const lineH = Math.round(st.font.size * 1.3636 * 4) / 4;
      let h = lineH * lines;
      // Turned text needs the height its rotated width takes (Excel grows the row).
      if (st.rotation && typeof c.v !== "object") {
        this.measureCtx.font = cssFont(st, 1);
        const tw = this.measureCtx.measureText(shown(c, st).text).width * 0.75;
        h = st.rotation === 255 ? lineH * String(shown(c, st).text).length : Math.abs(Math.sin(((st.rotation <= 90 ? st.rotation : st.rotation - 90) * Math.PI) / 180)) * tw + lineH * 0.6;
      }
      need.set(row, Math.max(need.get(row) ?? 0, h));
    }
    for (const row of rows) {
      const hgt = need.get(row);
      const info = sheet.rows.get(row);
      const def = sheet.defaultRowHeight;
      const target = hgt && hgt > def ? hgt : undefined;
      if ((info?.height ?? undefined) === target && !force) continue;
      if (!target && !info?.height) continue;
      // Looking at part of a row only, a row may grow but not shrink (another cell may need the height).
      if (partial && (target ?? def) < (info?.height ?? def)) continue;
      this.history.row(sheet, row);
      sheet.rows.set(row, { ...(info ?? {}), height: target, custom: false });
    }
  }

  // ----- formatting -----------------------------------------------------------------------------

  /** The style of the active cell (what the ribbon shows). */
  activeStyle(): CellStyle {
    const { row, col } = this.sheet.view.active;
    return this.wb.styles.get(this.sheet.get(row, col)?.s ?? this.sheet.rows.get(row)?.s ?? this.sheet.cols.get(col)?.s);
  }

  formatSel(label: string, change: (s: CellStyle) => void): void {
    this.change(label, () => A.format(this.ctx, this.sheet, this.grid.selection, change));
    // F4 and Strg+Y (with nothing to redo) repeat the last formatting on the new selection (Excel).
    this.lastRepeat = { label, run: () => this.formatSel(label, change) };
  }

  private lastRepeat: { label: string; run: () => void } | null = null;

  /** Ausrichtung > Text drehen: the angle, and the rows grow for it (Excel). */
  private rotateText(deg: number): void {
    const cur = this.activeStyle().rotation;
    this.formatSel("Ausrichtung", (s) => { s.rotation = cur === deg ? undefined : deg; });
    this.change("Zeilenhöhe", () => this.autoRowHeights(this.grid.selection.at(-1)!));
  }

  private toggleFont(key: "bold" | "italic" | "strike"): void {
    const on = !this.activeStyle().font[key];
    this.formatSel("Format", (s) => { s.font[key] = on || undefined; });
  }

  private setBorders(kind: string): void {
    const sel = this.grid.selection;
    const thin = { style: "thin" };
    this.change("Rahmen", () => {
      for (const r of sel) {
        const sides = (row: number, col: number) => {
          const top = row === r.r0, bottom = row === r.r1, left = col === r.c0, right = col === r.c1;
          return { top, bottom, left, right };
        };
        A.format(this.ctx, this.sheet, [r], () => {});
        this.history.cells(this.sheet, r);
        for (let row = r.r0; row <= r.r1; row++) {
          for (let col = r.c0; col <= r.c1; col++) {
            const c = this.sheet.get(row, col) ?? { v: null };
            const p = sides(row, col);
            const s = this.wb.styles.derive(c.s, (st) => {
              const b = st.borders;
              switch (kind) {
                case "none": st.borders = {}; break;
                case "all": b.top = b.bottom = b.left = b.right = { ...thin }; break;
                case "outside": if (p.top) b.top = { ...thin }; if (p.bottom) b.bottom = { ...thin }; if (p.left) b.left = { ...thin }; if (p.right) b.right = { ...thin }; break;
                case "thickOutside": if (p.top) b.top = { style: "medium" }; if (p.bottom) b.bottom = { style: "medium" }; if (p.left) b.left = { style: "medium" }; if (p.right) b.right = { style: "medium" }; break;
                case "bottom": if (p.bottom) b.bottom = { ...thin }; break;
                case "top": if (p.top) b.top = { ...thin }; break;
                case "left": if (p.left) b.left = { ...thin }; break;
                case "right": if (p.right) b.right = { ...thin }; break;
                case "doubleBottom": if (p.bottom) b.bottom = { style: "double" }; break;
                case "thickBottom": if (p.bottom) b.bottom = { style: "medium" }; break;
                case "topBottom": if (p.top) b.top = { ...thin }; if (p.bottom) b.bottom = { ...thin }; break;
                case "topThickBottom": if (p.top) b.top = { ...thin }; if (p.bottom) b.bottom = { style: "medium" }; break;
                case "topDoubleBottom": if (p.top) b.top = { ...thin }; if (p.bottom) b.bottom = { style: "double" }; break;
              }
            });
            this.sheet.set(row, col, { ...c, s });
          }
        }
      }
    });
    this.lastBorder = kind;
  }
  private lastBorder = "bottom";

  private merge(kind: "center" | "across" | "merge" | "unmerge"): void {
    const sel = this.grid.selection;
    this.change("Verbinden", () => {
      this.history.merges(this.sheet);
      for (const r of sel) {
        this.sheet.merges = this.sheet.merges.filter((m) => !A.overlaps(m, r));
        if (kind === "unmerge") continue;
        const parts = kind === "across" ? Array.from({ length: r.r1 - r.r0 + 1 }, (_, i) => ({ r0: r.r0 + i, r1: r.r0 + i, c0: r.c0, c1: r.c1 })) : [r];
        for (const p of parts) {
          if (p.r0 === p.r1 && p.c0 === p.c1) continue;
          // Excel keeps only the top-left value.
          this.history.cells(this.sheet, p);
          for (let row = p.r0; row <= p.r1; row++) for (let col = p.c0; col <= p.c1; col++) {
            if (row === p.r0 && col === p.c0) continue;
            const c = this.sheet.get(row, col);
            if (c) this.sheet.set(row, col, { v: null, s: c.s });
          }
          this.sheet.merges.push(p);
        }
        if (kind === "center") A.format(this.ctx, this.sheet, [r], (s) => { s.h = "center"; });
      }
    });
    this.grid.select(sel.at(-1)!, this.sheet.view.active);
  }

  private hasMergeInSelection(): boolean {
    return this.grid.selection.some((r) => this.sheet.merges.some((m) => A.overlaps(m, r)));
  }

  /** Dezimalstelle hinzufügen/entfernen: changes the number format as Excel does. */
  private stepDecimals(d: 1 | -1): void {
    const { row, col } = this.sheet.view.active;
    const c = this.sheet.get(row, col);
    const base = this.activeStyle().numFmt;
    let code = base;
    if (code === "General") {
      // From Standard Excel starts with the decimals the active cell shows.
      const t = typeof c?.v === "number" ? shown(c, this.activeStyle()).text : "";
      const m = /,(\d+)/.exec(t);
      const n = m ? m[1].length : 0;
      code = n + d > 0 ? "0." + "0".repeat(n + d) : "0";
      this.formatSel("Zahlenformat", (s) => { s.numFmt = code; });
      return;
    }
    const sections = code.split(";").map((sec) => {
      const m = /(0|#)(\.0*)?(?![^"]*"(?:[^"]*"[^"]*")*[^"]*$)/.exec(sec);
      if (!m) return sec;
      const idx = sec.lastIndexOf(m[0]) >= 0 ? sec.search(/[0#](\.0*)?(?=[^0#.]*$|[%E\s_)"\\*]|$)/) : -1;
      const re = /([0#])(\.(0+))?(?=[^0#.]*$)/;
      void idx;
      return sec.replace(re, (_all, digit: string, _dot: string, zeros: string | undefined) => {
        const n = (zeros?.length ?? 0) + d;
        return n > 0 ? `${digit}.${"0".repeat(n)}` : digit;
      });
    });
    code = sections.join(";");
    this.formatSel("Zahlenformat", (s) => { s.numFmt = code; });
  }

  /** Format übertragen: the copied format goes onto the next selection (double-click: until Esc). */
  private applyPainter(): void {
    const p = this.painter!;
    if (!p.keep) this.painter = null;
    const target = this.grid.selection;
    this.change("Format übertragen", () => {
      for (const r of target) {
        this.history.cells(this.sheet, r);
        for (let row = r.r0; row <= r.r1; row++) for (let col = r.c0; col <= r.c1; col++) {
          const c = this.sheet.get(row, col);
          this.sheet.set(row, col, { ...(c ?? { v: null }), s: p.style });
        }
      }
    });
    document.body.classList.toggle("painting", !!this.painter);
  }

  // ----- clipboard ------------------------------------------------------------------------------

  private async copy(cut: boolean): Promise<void> {
    const sel = this.grid.selection;
    if (sel.length !== 1) {
      await messageBox("RaabeCalc", "Dieser Vorgang kann für Mehrfachmarkierungen nicht ausgeführt werden.");
      return;
    }
    const r = clampToUsed(sel[0], this.sheet);
    const text = rangeText(this.wb, this.sheet, r);
    const html = rangeHtml(this.wb, this.sheet, r);
    this.copied = { sheet: this.sheet, rect: r, cut, text };
    this.grid.setMarquee(r);
    try {
      if (this.platform.clipboardWriteHtml) await this.platform.clipboardWriteHtml(html, text);
      else await this.platform.clipboardWrite(text);
    } catch { /* the internal copy still works */ }
  }

  async paste(what: "all" | "values" | "formats" | "formulas" | "transpose" = "all"): Promise<void> {
    const at = { row: this.grid.selection.at(-1)!.r0, col: this.grid.selection.at(-1)!.c0 };
    let text = "";
    try { text = await this.platform.clipboardRead(); } catch { /* empty */ }
    const c = this.copied;
    const same = c && this.wb.sheets.includes(c.sheet) && (text === "" || text.replace(/\r\n/g, "\n") === c.text.replace(/\r\n/g, "\n"));
    if (c && same) {
      const target = this.grid.selection.at(-1)!;
      const size = what === "transpose" ? { h: c.rect.c1 - c.rect.c0, w: c.rect.r1 - c.rect.r0 } : { h: c.rect.r1 - c.rect.r0, w: c.rect.c1 - c.rect.c0 };
      const written = { r0: target.r0, c0: target.c0, r1: Math.max(target.r1, target.r0 + size.h), c1: Math.max(target.c1, target.c0 + size.w) };
      if (!this.protection.guardRects(c.cut ? [written, c.rect] : [written])) return;
      this.change(c.cut ? "Ausschneiden" : "Einfügen", () => {
        if (c.cut && c.sheet === this.sheet) {
          const dst = A.moveCells(this.ctx, this.sheet, c.rect, at);
          this.grid.select(dst, at);
        } else {
          // A copied range fills a larger selection that is a multiple of it (Excel).
          const h0 = c.rect.r1 - c.rect.r0 + 1, w0 = c.rect.c1 - c.rect.c0 + 1;
          const th = target.r1 - target.r0 + 1, tw = target.c1 - target.c0 + 1;
          const reps = what !== "transpose" && th % h0 === 0 && tw % w0 === 0 ? { r: th / h0, c: tw / w0 } : { r: 1, c: 1 };
          let last: Rect = target;
          for (let i = 0; i < reps.r; i++) for (let j = 0; j < reps.c; j++) {
            last = A.copyCells(this.ctx, { sheet: c.sheet, rect: c.rect }, this.sheet, { row: at.row + i * h0, col: at.col + j * w0 }, what === "transpose" ? "all" : what, what === "transpose");
          }
          if (c.cut) A.clearAll(this.ctx, c.sheet, [c.rect], "all");
          const all = reps.r * reps.c > 1 ? target : last;
          this.grid.select(all, at);
        }
      });
      if (c.cut) { this.copied = null; this.grid.setMarquee(null); }
      else {
        const pasted = this.grid.selection.at(-1)!;
        const again = (w: "all" | "values" | "formats" | "formulas" | "transpose") => () => { this.history.undo(); this.grid.select(target, at); void this.paste(w); };
        this.grid.showOptions(pasted, "paste", (anchor) => showMenu(this.commands, [
          { label: "Einfügen", icons: [
            { icon: "paste", tip: "Einfügen (E)", run: again("all") },
            { icon: "fx", tip: "Formeln (F)", run: again("formulas") },
            { icon: "table-switch", tip: "Transponieren (T)", run: again("transpose") },
          ] },
          { label: "Werte einfügen", icons: [{ icon: "number-value", tip: "Werte (W)", run: again("values") }] },
          { label: "Weitere Einfügeoptionen", icons: [{ icon: "format-painter", tip: "Formatierung (R)", run: again("formats") }] },
        ], anchor));
      }
      return;
    }
    let html = "";
    try { if (this.platform.clipboardReadHtml) html = await this.platform.clipboardReadHtml(); } catch { /* text only */ }
    const rows = parseClipboard(html, text);
    if (!rows.length || (rows.length === 1 && rows[0].length === 1 && rows[0][0].text === "")) return;
    const r = { r0: at.row, c0: at.col, r1: at.row + rows.length - 1, c1: at.col + Math.max(...rows.map((x) => x.length)) - 1 };
    if (!this.protection.guardRects([r])) return;
    this.change("Einfügen", () => {
      rows.forEach((cells, i) => cells.forEach((c, j) => {
        const cell = { r0: at.row + i, c0: at.col + j, r1: at.row + i, c1: at.col + j };
        A.enter(this.ctx, this.sheet, cell, c.text);
        // The look from the other program (Excel's classes, Word's styles): bold, colours, number format.
        const l = c.look;
        if (l && what !== "values") A.format(this.ctx, this.sheet, [cell], (s) => {
          if (l.bold) s.font.bold = true;
          if (l.italic) s.font.italic = true;
          if (l.underline) s.font.underline = "single";
          if (l.color) s.font.color = l.color;
          if (l.fill) s.fill = l.fill;
          if (l.numFmt) s.numFmt = l.numFmt;
        });
      }));
    });
    this.grid.select(r, at);
  }

  // ----- files ----------------------------------------------------------------------------------

  async openDialog(): Promise<void> {
    const path = await this.platform.openFileDialog("xlsx");
    if (path) await this.openPath(path);
  }

  async openPath(path: string): Promise<void> {
    const existing = this.docs.find((d) => d.path && samePath(d.path, path));
    if (existing) { this.activate(existing); this.backstage.close(); return; }
    let wb: Workbook;
    try {
      const bytes = await this.platform.readFile(path);
      if (/\.csv$/i.test(path)) wb = readCsv(new TextDecoder(bytes[0] === 0xef && bytes[1] === 0xbb ? "utf-8" : detectEncoding(bytes)).decode(bytes), baseName(path).replace(/\.csv$/i, ""));
      else wb = readXlsx(bytes);
    } catch (e) {
      removeRecent(path);
      await messageBox("RaabeCalc", `Die Datei „${baseName(path)}“ kann nicht geöffnet werden.\n\n${e instanceof Error ? e.message : String(e)}`);
      return;
    }
    const doc = this.makeDoc(wb, stripExt(baseName(path)), path);
    doc.lastSaved = Date.now();
    // An untouched new workbook makes room for the opened file (Excel).
    const blank = this.docs.length === 1 && this.docs[0].path === null && this.docs[0].version === 0;
    if (blank) this.docs = [];
    this.docs.push(doc);
    addRecent(path, baseName(path));
    this.backstage.close();
    this.activate(doc);
  }

  async save(doc = this.active, as = false): Promise<boolean> {
    if (this.grid.isEditing) this.grid.commitEdit("none");
    let path = doc.path;
    if (as || !path || /\.csv$/i.test(path) && as) {
      path = await this.platform.saveFileDialog(`${doc.name}.xlsx`, "xlsx");
      if (!path) return false;
      if (!/\.(xlsx|csv)$/i.test(path)) path += ".xlsx";
    }
    try {
      const bytes = /\.csv$/i.test(path) ? new TextEncoder().encode("\ufeff" + writeCsv(doc.wb.sheet(), { styles: doc.wb.styles, date1904: doc.wb.date1904 })) : writeXlsx(doc.wb);
      await this.platform.writeFile(path, bytes);
    } catch (e) {
      await messageBox("RaabeCalc", `Die Datei konnte nicht gespeichert werden.\n\n${e instanceof Error ? e.message : String(e)}`);
      return false;
    }
    doc.path = path;
    doc.name = stripExt(baseName(path));
    doc.savedVersion = doc.version;
    doc.lastSaved = Date.now();
    addRecent(path, baseName(path));
    void this.platform.deleteRecovery(doc.id).catch(() => {});
    this.updateTabs();
    return true;
  }

  private async saveCsv(): Promise<void> {
    const path = await this.platform.saveFileDialog(`${this.active.name}.csv`, "csv");
    if (!path) return;
    await this.platform.writeFile(path, new TextEncoder().encode("\ufeff" + writeCsv(this.sheet, { styles: this.wb.styles, date1904: this.wb.date1904 })));
  }

  private async closeDoc(doc: Doc): Promise<void> {
    if (doc.version !== doc.savedVersion) {
      const r = await showDialog({
        title: "RaabeCalc",
        body: `Möchten Sie die Änderungen in „${doc.name}“ speichern?`,
        buttons: [{ label: "Speichern", value: "save", primary: true }, { label: "Nicht speichern", value: "discard" }, { label: "Abbrechen", value: "cancel" }],
        cancel: "cancel",
      });
      if (r === "cancel") return;
      if (r === "save" && !(await this.save(doc))) return;
    }
    void this.platform.deleteRecovery(doc.id).catch(() => {});
    const i = this.docs.indexOf(doc);
    this.docs.splice(i, 1);
    if (!this.docs.length) { await this.platform.window.close(); return; }
    if (doc === this.active) this.activate(this.docs[Math.max(0, i - 1)]);
    else this.updateTabs();
  }

  private async confirmQuit(): Promise<boolean> {
    for (const d of [...this.docs]) {
      if (d.version === d.savedVersion) continue;
      this.activate(d);
      const r = await showDialog({
        title: "RaabeCalc",
        body: `Möchten Sie die Änderungen in „${d.name}“ speichern?`,
        buttons: [{ label: "Speichern", value: "save", primary: true }, { label: "Nicht speichern", value: "discard" }, { label: "Abbrechen", value: "cancel" }],
        cancel: "cancel",
      });
      if (r === "cancel") return false;
      if (r === "save" && !(await this.save(d))) return false;
    }
    for (const d of this.docs) void this.platform.deleteRecovery(d.id).catch(() => {});
    return true;
  }

  private async autoRecover(): Promise<void> {
    for (const d of this.docs) {
      if (d.version === d.recoveryVersion || d.version === d.savedVersion) continue;
      try {
        await this.platform.saveRecovery({ id: d.id, title: d.name, path: d.path, savedAt: Date.now() }, writeXlsx(d.wb));
        d.recoveryVersion = d.version;
      } catch { /* next time */ }
    }
  }

  private async offerRecovery(): Promise<void> {
    const list = await this.platform.listRecovery().catch(() => []);
    for (const e of list) {
      const r = await showDialog({
        title: "Dokumentwiederherstellung",
        body: `RaabeCalc wurde unerwartet beendet. „${e.title}“ (gesichert ${new Date(e.savedAt).toLocaleString("de-DE")}) wiederherstellen?`,
        buttons: [{ label: "Wiederherstellen", value: "yes", primary: true }, { label: "Verwerfen", value: "no" }],
        cancel: "no",
      });
      if (r === "yes") {
        try {
          const wb = readXlsx(await this.platform.readRecovery(e.id));
          const doc = this.makeDoc(wb, e.title, e.path);
          doc.version = 1;
          this.docs.push(doc);
          this.activate(doc);
        } catch { /* broken backup */ }
      }
      await this.platform.deleteRecovery(e.id).catch(() => {});
    }
  }

  // ----- sheets ---------------------------------------------------------------------------------

  addSheet(after: boolean): void {
    this.change("Blatt einfügen", () => {
      this.history.sheets();
      const s = new Sheet(this.nextSheetName());
      s.defaultColWidth = this.sheet.defaultColWidth ?? DEFAULT_COL_CHARS;
      const at = this.wb.active + (after ? 1 : 0);
      this.wb.sheets.splice(at, 0, s);
      this.wb.active = at;
    });
  }

  /** Strg+Umschalt+1/4/5: the format, and the column widens when the number no longer fits (Excel). */
  private numberFormatKey(code: string): void {
    this.change("Zahlenformat", () => {
      A.format(this.ctx, this.sheet, this.grid.selection, (s) => { s.numFmt = code; });
      this.autoWiden(this.grid.selection);
    });
  }

  /**
   * Excel widens a column of standard width when a number or date it shows with a format does
   * not fit (typing 1.1.2026, Währung): to the widest such text plus the cell padding.
   */
  autoWiden(ranges: Rect[]): void {
    const sheet = this.sheet;
    const need = new Map<number, number>();
    for (const r of ranges) {
      if (r.r1 - r.r0 > 2000) continue;
      for (let row = r.r0; row <= r.r1; row++) for (let col = r.c0; col <= Math.min(r.c1, r.c0 + 200); col++) {
        const c = sheet.get(row, col);
        if (!c || typeof c.v !== "number") continue;
        const st = this.wb.styles.get(c.s);
        const info = sheet.cols.get(col);
        const curPx = info?.width !== undefined ? charsToPx(info.width) : defaultColPx(sheet.defaultColWidth ?? DEFAULT_COL_CHARS);
        const full = shown(c, st).text;
        const fitted = fitShown(this.measureCtx, c, st, curPx).text;
        if (fitted === full) continue;
        // Standard numbers widen the column only where the width would turn them into
        // scientific notation or cut a small number; dropping decimals is fine (recorded).
        if (st.numFmt === "General" && !(full.includes("E") || fitted.includes("E") || Math.abs(c.v) < 1)) continue;
        // The width the full text needs, as Excel decides whether it fits.
        let px = curPx;
        while (px < 2000 && fitShown(this.measureCtx, c, st, px).text !== full) px += 1 / devicePixelRatio;
        need.set(col, Math.max(need.get(col) ?? 0, px));
      }
    }
    for (const [col, w] of need) {
      const info = sheet.cols.get(col);
      if (info?.width !== undefined && !info.bestFit) continue;
      const cur = info?.width !== undefined ? charsToPx(info.width) : defaultColPx(sheet.defaultColWidth ?? DEFAULT_COL_CHARS);
      const px = w;
      if (px <= cur) continue;
      this.history.col(sheet, col);
      sheet.cols.set(col, { ...(info ?? {}), width: pxToWidth(px) });
    }
  }

  private nextSheetName(): string {
    for (let i = 1; ; i++) {
      const n = `Tabelle${i}`;
      if (!this.wb.sheetByName(n)) return n;
    }
  }

  renameSheet(i: number, name: string): string | null {
    const bad = /[\\/?*[\]:]/.test(name) || name.length > 31 || name.startsWith("'") || name.endsWith("'");
    if (bad || this.wb.sheets.some((s, j) => j !== i && s.name.toLowerCase() === name.toLowerCase())) {
      void messageBox("RaabeCalc", bad ? "Der eingegebene Blattname ist ungültig." : "Dieser Name wird bereits verwendet. Geben Sie einen anderen Namen ein.");
      return "invalid";
    }
    const old = this.wb.sheets[i].name;
    this.change("Umbenennen", () => {
      this.history.sheets();
      this.wb.sheets[i].name = name;
      // Formulas that name the sheet follow (Excel).
      const quoted = (n: string) => (/^[A-Za-z_][A-Za-z0-9_.]*$/.test(n) ? n : `'${n.replace(/'/g, "''")}'`);
      const re = new RegExp(`(^|[^A-Za-z0-9_.'])(${escapeRe(quoted(old))}|'${escapeRe(old.replace(/'/g, "''"))}')!`, "gi");
      for (const s of this.wb.sheets) for (const [k, c] of s.cells) {
        if (c.f && re.test(c.f)) {
          this.history.whole(s);
          s.cells.set(k, { ...c, f: c.f.replace(re, (_m, pre: string) => `${pre}${quoted(name)}!`) });
        }
        re.lastIndex = 0;
      }
    });
    return null;
  }

  private moveSheet(from: number, to: number, copy: boolean): void {
    if (this.protection.bookLocked()) return;
    this.change(copy ? "Blatt kopieren" : "Blatt verschieben", () => {
      this.history.sheets();
      const s = this.wb.sheets[from];
      if (copy) {
        const c = new Sheet(uniqueCopyName(this.wb, s.name));
        c.cells = new Map([...s.cells].map(([k, v]) => [k, { ...v }]));
        c.cols = new Map([...s.cols].map(([k, v]) => [k, { ...v }]));
        c.rows = new Map([...s.rows].map(([k, v]) => [k, { ...v }]));
        c.merges = s.merges.map((m) => ({ ...m }));
        c.validations = s.validations ? structuredClone(s.validations) : undefined;
        // The copy keeps the sheet's settings (Excel): protection, page setup, filter, rules,
        // Kommentare (new ids) and pictures (written as new ones).
        c.protection = s.protection ? { ...s.protection } : undefined;
        c.page = s.page ? structuredClone(s.page) : undefined;
        c.autoFilter = s.autoFilter ? structuredClone(s.autoFilter) : undefined;
        c.conditional = s.conditional ? structuredClone(s.conditional) : undefined;
        c.threads = s.threads?.map((t) => ({ ...t, id: newId(), personId: undefined, replies: t.replies.map((r) => ({ ...r, id: newId(), personId: undefined })) }));
        c.pictures = s.pictures?.map((p) => ({ ...p, raw: undefined }));
        c.tabColor = s.tabColor;
        c.defaultColWidth = s.defaultColWidth;
        c.defaultRowHeight = s.defaultRowHeight;
        c.view = structuredClone(s.view);
        // Charts come along and show the copy's cells (Excel).
        if (s.charts?.length) {
          const quote = (n: string) => (/^[A-Za-z_][A-Za-z0-9_.]*$/.test(n) ? n : `'${n.replace(/'/g, "''")}'`);
          const from = quote(s.name) + "!", to = quote(c.name) + "!";
          c.charts = structuredClone(s.charts).map((ch) => {
            for (const ser of ch.series) for (const d of [ser.name, ser.cat, ser.val]) if (d?.ref?.startsWith(from)) d.ref = to + d.ref.slice(from.length);
            return ch;
          });
        }
        this.wb.sheets.splice(to, 0, c);
        this.wb.active = to;
      } else {
        this.wb.sheets.splice(from, 1);
        this.wb.sheets.splice(to, 0, s);
        this.wb.active = to;
      }
    });
  }

  private sheetMenu(e: MouseEvent): void {
    const i = this.wb.active;
    const colors = ["#C00000", "#FF0000", "#FFC000", "#FFFF00", "#92D050", "#00B050", "#00B0F0", "#0070C0", "#002060", "#7030A0"];
    const entries: MenuEntry[] = [
      { label: "Einfügen...", run: () => void this.commands.run("cells.insertSheet") },
      { label: "Löschen", run: () => void this.commands.run("cells.deleteSheet"), disabled: this.wb.sheets.filter((s) => !s.hidden).length < 2 },
      { label: "Umbenennen", run: () => this.sheetBar.renameInPlace(null, i) },
      { label: "Verschieben oder kopieren...", run: () => void this.dlg.moveSheet() },
      "-",
      { label: "Code anzeigen", run: () => {}, disabled: true },
      { label: "Blatt schützen...", run: () => {}, disabled: true },
      { label: "Registerfarbe", submenu: [
        { label: "Keine Farbe", disabled: this.protection.bookLocked(), run: () => this.change("Registerfarbe", () => { this.history.whole(this.sheet); this.sheet.tabColor = undefined; }) },
        ...colors.map((c) => ({ label: c, icon: undefined, disabled: this.protection.bookLocked(), run: () => this.change("Registerfarbe", () => { this.history.whole(this.sheet); this.sheet.tabColor = c; }) })),
      ] },
      "-",
      { label: "Ausblenden", run: () => void this.commands.run("cells.hideSheet"), disabled: this.wb.sheets.filter((s) => !s.hidden).length < 2 },
      { label: "Einblenden...", run: () => void this.commands.run("cells.unhideSheet"), disabled: !this.wb.sheets.some((s) => s.hidden) },
      "-",
      { label: "Alle Blätter auswählen", run: () => {}, disabled: true },
    ];
    showMenu(this.commands, entries, new DOMRect(e.clientX, e.clientY, 0, 0));
  }

  // ----- zoom and chrome ------------------------------------------------------------------------

  setZoom(z: number): void {
    this.grid.setZoom(z);
    this.updateChrome();
  }

  private scheduleChrome(): void {
    if (this.refreshTimer) return;
    this.refreshTimer = requestAnimationFrame(() => {
      this.refreshTimer = 0;
      this.updateChrome();
    });
  }

  /** Labels that follow the content, as Excel's do: Fixierung aufheben, sorting by text or by size. */
  private updateLabels(): void {
    this.commands.get("view.freezePanes").label = this.sheet.view.freeze ? "Fixierung aufheben" : "Fenster fixieren";
    const prot = this.protection.on();
    this.commands.get("review.protectSheet").label = prot ? "Blattschutz aufheben..." : "Blatt schützen...";
    this.commands.get("cells.protectSheet").label = prot ? "Blattschutz aufheben..." : "Blatt schützen...";
    const { row, col } = this.sheet.view.active;
    const text = typeof this.sheet.get(row, col)?.v === "string";
    this.commands.get("edit.sortAscMenu").label = text ? "Von A bis Z sortieren" : "Nach Größe sortieren (aufsteigend)";
    this.commands.get("edit.sortDescMenu").label = text ? "Von Z bis A sortieren" : "Nach Größe sortieren (absteigend)";
  }

  /** Inhalte auswählen: Excel's choices of cells to select. */
  private async gotoSpecial(): Promise<void> {
    const opts: [string, string][] = [["notes", "Notizen"], ["constants", "Konstanten"], ["formulas", "Formeln"], ["blanks", "Leerzellen"], ["region", "Aktueller Bereich"],
      ["last", "Letzte Zelle"], ["visible", "Nur sichtbare Zellen"], ["conditional", "Bedingte Formate"]];
    const body = h("div", { class: "goto-special" }, ...opts.map(([v, l], i) => h("label", { class: "check" }, h("input", { type: "radio", name: "gs", value: v, checked: i === 0 || undefined }), l)));
    const ok = await showDialog({ title: "Inhalte auswählen", body, buttons: [{ label: "OK", value: true, primary: true }, { label: "Abbrechen", value: false }], cancel: false });
    if (!ok) return;
    const v = body.querySelector<HTMLInputElement>("input:checked")?.value;
    const a = this.sheet.view.active;
    if (v === "notes") return void this.commands.run("edit.selectNotes");
    if (v === "constants") return void this.commands.run("edit.selectConstants");
    if (v === "formulas") return void this.commands.run("edit.selectFormulas");
    if (v === "conditional") return void this.commands.run("edit.selectConditional");
    if (v === "region") { this.grid.select(this.grid.currentRegion(a.row, a.col), { ...a }); return; }
    if (v === "last") { const u = this.sheet.usedRange(); if (u) this.grid.selectCell(u.r1, u.c1); return; }
    const sel = this.grid.selection.length > 1 || this.grid.selection[0].r0 !== this.grid.selection[0].r1 ? this.grid.selection : [this.sheet.usedRange() ?? this.grid.selection[0]];
    const hits: Rect[] = [];
    for (const r of sel) for (let row = r.r0; row <= Math.min(r.r1, r.r0 + 5000); row++) {
      if (v === "visible" && this.sheet.rows.get(row)?.hidden) continue;
      for (let col = r.c0; col <= Math.min(r.c1, r.c0 + 200); col++) {
        if (v === "visible") { if (!this.sheet.cols.get(col)?.hidden) hits.push({ r0: row, c0: col, r1: row, c1: col }); continue; }
        const c = this.sheet.get(row, col);
        if (!c || c.v === null || c.v === "") hits.push({ r0: row, c0: col, r1: row, c1: col });
      }
    }
    if (!hits.length) { void messageBox("RaabeCalc", "Es wurden keine Zellen gefunden."); return; }
    this.grid.selectMany(hits.slice(0, 5000));
  }

  /** Suchen und Auswählen > Formeln, Notizen, Konstanten …: those cells of the sheet (or the selection) selected. */
  private selectWhere(test: (c: import("@raabe/sheet-engine").Cell, row: number, col: number) => boolean, label: string): void {
    const sel = this.grid.selection;
    const within = sel.length > 1 || sel[0].r0 !== sel[0].r1 || sel[0].c0 !== sel[0].c1 ? sel : null;
    const hits: Rect[] = [];
    for (const [k, c] of this.sheet.cells) {
      const row = Math.floor(k / MAX_COLS), col = k % MAX_COLS;
      if (within && !within.some((r) => row >= r.r0 && row <= r.r1 && col >= r.c0 && col <= r.c1)) continue;
      if (test(c, row, col)) hits.push({ r0: row, c0: col, r1: row, c1: col });
    }
    if (!hits.length) { void messageBox("RaabeCalc", `Es wurden keine Zellen gefunden (${label}).`); return; }
    hits.sort((a, b) => a.r0 - b.r0 || a.c0 - b.c0);
    this.grid.selectMany(hits.slice(0, 5000));
  }

  updateChrome(): void {
    this.updateLabels();
    this.statusBar.update(this.sheet.view.zoom, this.stats());
    this.commands.refresh();
    this.ribbon.refresh?.();
  }

  /** Mittelwert, Anzahl and Summe of the selection (the status bar). */
  private stats(): Stats | null {
    const sel = this.grid.selection;
    if (sel.length === 1 && sel[0].r0 === sel[0].r1 && sel[0].c0 === sel[0].c1) return null;
    const s: Stats = { count: 0, numbers: 0, sum: 0, min: Infinity, max: -Infinity };
    const seen = new Set<number>();
    for (const r of sel) {
      for (const [k, c] of this.sheet.cells) {
        if (seen.has(k)) continue;
        const row = Math.floor(k / MAX_COLS), col = k % MAX_COLS;
        if (row < r.r0 || row > r.r1 || col < r.c0 || col > r.c1) continue;
        if (c.v === null || c.v === "") continue;
        seen.add(k);
        s.count++;
        if (typeof c.v === "number") { s.numbers++; s.sum += c.v; s.min = Math.min(s.min, c.v); s.max = Math.max(s.max, c.v); }
      }
    }
    return s;
  }

  // ----- AutoSumme --------------------------------------------------------------------------------

  /** AutoSumme: SUMME (or another function) over the numbers above or to the left (Excel). */
  autoSum(fn = "SUMME"): void {
    const sel = this.grid.selection.at(-1)!;
    const numeric = (r: number, c: number) => { const x = this.sheet.get(r, c); return typeof x?.v === "number"; };
    const empty = (r: number, c: number) => { const x = this.sheet.get(r, c); return !x || x.v === null || x.v === ""; };
    if (sel.r0 !== sel.r1 || sel.c0 !== sel.c1) {
      // A selected range: a result below each column (or right of each row).
      const targets: { row: number; col: number; ref: Rect }[] = [];
      const lastRowEmpty = Array.from({ length: sel.c1 - sel.c0 + 1 }, (_, j) => empty(sel.r1, sel.c0 + j)).every(Boolean);
      for (let c = sel.c0; c <= sel.c1; c++) {
        if (lastRowEmpty) targets.push({ row: sel.r1, col: c, ref: { r0: sel.r0, c0: c, r1: sel.r1 - 1, c1: c } });
        else targets.push({ row: sel.r1 + 1, col: c, ref: { r0: sel.r0, c0: c, r1: sel.r1, c1: c } });
      }
      this.change("AutoSumme", () => {
        for (const t of targets) A.enter(this.ctx, this.sheet, { r0: t.row, c0: t.col, r1: t.row, c1: t.col }, `=${fn}(${rangeName(t.ref)})`);
      });
      return;
    }
    const { r0: row, c0: col } = sel;
    let ref: Rect | null = null;
    if (row > 0 && numeric(row - 1, col)) {
      let r = row - 1;
      while (r > 0 && !empty(r - 1, col) && (numeric(r - 1, col) || r - 1 > 0)) { if (!numeric(r - 1, col)) break; r--; }
      ref = { r0: r, c0: col, r1: row - 1, c1: col };
    } else if (col > 0 && numeric(row, col - 1)) {
      let c = col - 1;
      while (c > 0 && numeric(row, c - 1)) c--;
      ref = { r0: row, c0: c, r1: row, c1: col - 1 };
    }
    const inner = ref ? rangeName(ref) : "";
    this.grid.startPointing(`=${fn}(${inner})`, `=${fn}(`.length, `=${fn}(`.length + inner.length, ref);
  }

  // ----- commands ---------------------------------------------------------------------------------

  private implementCommands(): void {
    const c = this.commands;
    // Blatt schützen decides first: commands it does not allow are greyed out, changes to locked
    // cells are refused with a message.
    const run = (id: string, fn: () => void | Promise<void>, enabled?: () => boolean, checked?: () => boolean) =>
      c.implement(id, { run: () => (this.protection.guard(id) ? fn() : undefined), enabled: () => this.protection.enabled(id) && !(this.protection.bookLocked() && this.protection.structureCommand(id)) && (enabled ? enabled() : true), checked });
    const notEditing = () => !this.grid?.isEditing;
    const st = () => this.activeStyle();

    run("file.new", () => this.newDoc());
    run("file.open", () => this.backstage.open("open"));
    run("file.save", () => void this.save());
    run("file.saveAs", () => void this.save(this.active, true));
    run("file.saveCsv", () => void this.saveCsv());
    run("file.print", () => this.backstage.open("print"));
    run("file.exportPdf", () => void this.print.exportPdf());
    run("file.close", () => void this.closeDoc(this.active));
    run("file.options", () => void this.dlg.options(this.settings, (s) => { this.settings = s; saveSettings(s); document.body.classList.toggle("touch", s.touchMode); }));

    run("edit.undo", () => {
      if (this.grid.isEditing) this.grid.cancelEdit();
      else {
        const label = this.history.undoLabel;
        this.history.undo();
        // Undoing a paste brings the copied range's marquee back (recorded).
        if (label === "Einfügen" && this.copied && !this.copied.cut && this.copied.sheet === this.sheet) this.grid.setMarquee(this.copied.rect);
      }
      this.grid.syncBar();
    }, () => this.grid?.isEditing || this.history.canUndo);
    run("edit.redo", () => {
      if (this.history.canRedo) this.history.redo();
      else this.lastRepeat?.run();
      this.grid.syncBar();
    }, () => notEditing() && (this.history.canRedo || !!this.lastRepeat));
    run("edit.repeat", () => this.lastRepeat?.run(), () => notEditing() && !!this.lastRepeat);
    run("edit.copy", () => void this.copy(false), notEditing);
    run("edit.cut", () => void this.copy(true), notEditing);
    run("edit.paste", () => void this.paste(), notEditing);
    run("edit.pasteValues", () => void this.paste("values"), notEditing);
    run("edit.pasteFormulas", () => void this.paste("formulas"), notEditing);
    run("edit.pasteFormats", () => void this.paste("formats"), notEditing);
    run("edit.pasteTranspose", () => void this.paste("transpose"), notEditing);
    run("edit.pasteValuesFormats", () => void this.paste("all"), notEditing);
    run("edit.pasteNoBorders", () => void this.paste("all"), notEditing);
    run("edit.pasteKeepWidths", () => void this.paste("all"), notEditing);
    run("edit.pasteSpecial", () => void this.dlg.pasteSpecial(), () => notEditing() && !!this.copied);
    run("edit.formatPainter", () => {
      const s = this.sheet.get(this.sheet.view.active.row, this.sheet.view.active.col)?.s;
      this.painter = this.painter ? null : { style: s, keep: false };
      document.body.classList.toggle("painting", !!this.painter);
    }, notEditing, () => !!this.painter);

    // The Schriftart and Schriftgrad boxes act through their lists; as commands they only focus them.
    run("font.name", () => this.focusRibbonInput("font.name"));
    run("font.size", () => this.focusRibbonInput("font.size"));
    run("font.bold", () => this.toggleFont("bold"), undefined, () => !!st().font.bold);
    run("font.italic", () => this.toggleFont("italic"), undefined, () => !!st().font.italic);
    run("font.strike", () => this.toggleFont("strike"), undefined, () => !!st().font.strike);
    run("font.underline", () => { const on = !st().font.underline; this.formatSel("Unterstreichen", (s) => { s.font.underline = on ? "single" : undefined; }); }, undefined, () => !!st().font.underline);
    run("font.underlineDouble", () => { const on = st().font.underline !== "double"; this.formatSel("Unterstreichen", (s) => { s.font.underline = on ? "double" : undefined; }); }, undefined, () => st().font.underline === "double");
    run("font.grow", () => this.stepFont(1));
    run("font.shrink", () => this.stepFont(-1));
    run("font.color", () => this.formatSel("Schriftfarbe", (s) => { s.font.color = this.lastFontColor || undefined; }));
    run("font.fill", () => this.formatSel("Füllfarbe", (s) => { s.fill = this.lastFill || undefined; s.fillXml = undefined; }));
    run("font.dialog", () => void this.dlg.formatCells("Schrift"));
    run("border.menu", () => this.setBorders(this.lastBorder));
    for (const k of ["bottom", "top", "left", "right", "none", "all", "outside", "thickOutside", "doubleBottom", "thickBottom", "topBottom", "topThickBottom", "topDoubleBottom"]) run(`border.${k}`, () => this.setBorders(k));
    run("border.more", () => void this.dlg.formatCells("Rahmen"));

    const align = (key: "h" | "v", value: string, id: string) => run(id, () => {
      const cur = st()[key];
      this.formatSel("Ausrichtung", (s) => { (s as unknown as Record<string, string | undefined>)[key] = cur === value ? undefined : value; });
    }, undefined, () => st()[key] === value || (key === "v" && value === "bottom" && !st().v));
    align("v", "top", "align.top"); align("v", "center", "align.middle"); align("v", "bottom", "align.bottom");
    align("h", "left", "align.left"); align("h", "center", "align.center"); align("h", "right", "align.right");
    run("align.wrap", () => { const on = !st().wrap; this.formatSel("Textumbruch", (s) => { s.wrap = on || undefined; }); this.change("Zeilenhöhe", () => this.autoRowHeights(this.grid.selection.at(-1)!)); }, undefined, () => !!st().wrap);
    run("align.indentMore", () => this.formatSel("Einzug", (s) => { s.indent = Math.min(250, (s.indent ?? 0) + 1); if (!s.h || s.h === "general" || s.h === "center") s.h = "left"; }));
    run("align.indentLess", () => this.formatSel("Einzug", (s) => { s.indent = Math.max(0, (s.indent ?? 0) - 1) || undefined; }));
    run("align.mergeCenter", () => this.merge(this.hasMergeInSelection() ? "unmerge" : "center"), undefined, () => this.hasMergeInSelection());
    run("align.mergeAcross", () => this.merge("across"));
    run("align.merge", () => this.merge("merge"));
    run("align.unmerge", () => this.merge("unmerge"));
    run("align.dialog", () => void this.dlg.formatCells("Ausrichtung"));
    run("align.orientation", () => {});

    run("number.format", () => {});
    run("number.accounting", () => this.formatSel("Zahlenformat", (s) => { s.numFmt = RIBBON_FORMATS[3].code; }));
    run("number.percent", () => this.formatSel("Zahlenformat", (s) => { s.numFmt = "0%"; }));
    run("number.thousands", () => this.formatSel("Zahlenformat", (s) => { s.numFmt = '_-* #,##0.00\\ _€_-;\\-* #,##0.00\\ _€_-;_-* "-"??\\ _€_-;_-@_-'; }));
    run("number.decimalMore", () => this.stepDecimals(1));
    run("number.decimalLess", () => this.stepDecimals(-1));
    run("number.dialog", () => void this.dlg.formatCells("Zahlen"));
    run("cells.formatDialog", () => void this.dlg.formatCells("Zahlen"));

    run("styles.conditional", () => {});
    run("styles.table", () => {});
    run("styles.cell", () => {});

    const selRect = () => this.grid.selection.at(-1)!;
    const rowsOf = (r: Rect) => ({ at: r.r0, count: r.r1 - r.r0 + 1 });
    const colsOf = (r: Rect) => ({ at: r.c0, count: r.c1 - r.c0 + 1 });
    run("cells.insert", () => {
      const r = selRect();
      if (r.c0 === 0 && r.c1 === MAX_COLS - 1) return void this.commands.run("cells.insertRows");
      if (r.r0 === 0 && r.r1 === MAX_ROWS - 1) return void this.commands.run("cells.insertCols");
      void this.dlg.insertCells();
    });
    run("cells.insertDialog", () => void this.dlg.insertCells());
    run("cells.insertRows", () => { const { at, count } = rowsOf(selRect()); this.change("Zeilen einfügen", () => A.insertDelete(this.ctx, this.sheet, "rows", at, count)); });
    run("cells.insertCols", () => { const { at, count } = colsOf(selRect()); this.change("Spalten einfügen", () => A.insertDelete(this.ctx, this.sheet, "cols", at, count)); });
    run("cells.delete", () => {
      const r = selRect();
      if (r.c0 === 0 && r.c1 === MAX_COLS - 1) return void this.commands.run("cells.deleteRows");
      if (r.r0 === 0 && r.r1 === MAX_ROWS - 1) return void this.commands.run("cells.deleteCols");
      void this.dlg.deleteCells();
    });
    run("cells.deleteRows", () => { const { at, count } = rowsOf(selRect()); this.change("Zeilen löschen", () => A.insertDelete(this.ctx, this.sheet, "rows", at, -count)); });
    run("cells.deleteCols", () => { const { at, count } = colsOf(selRect()); this.change("Spalten löschen", () => A.insertDelete(this.ctx, this.sheet, "cols", at, -count)); });
    // Umschalt+F11 and Einfügen > Blatt einfügen put the new sheet before the active one; the + beside the tabs after it (Excel).
    run("cells.insertSheet", () => this.addSheet(false));
    run("cells.deleteSheet", async () => {
      if (this.wb.sheets.filter((s) => !s.hidden).length < 2) return;
      if (this.sheet.cells.size) {
        const r = await showDialog({ title: "RaabeCalc", body: "Das Blatt wird endgültig gelöscht. Möchten Sie den Vorgang fortsetzen?", buttons: [{ label: "Löschen", value: true, primary: true }, { label: "Abbrechen", value: false }], cancel: false });
        if (!r) return;
      }
      this.change("Blatt löschen", () => {
        this.history.sheets();
        this.wb.sheets.splice(this.wb.active, 1);
        this.wb.active = Math.max(0, this.wb.active - 1);
        while (this.wb.sheet().hidden) this.wb.active = (this.wb.active + 1) % this.wb.sheets.length;
      });
    });
    run("cells.format", () => {});
    run("cells.rowHeight", () => void this.dlg.rowHeight());
    run("cells.colWidth", () => void this.dlg.colWidth());
    run("cells.defaultWidth", () => void this.dlg.defaultWidth());
    run("cells.autofitCols", () => { const r = selRect(); this.autofitCols(Array.from({ length: r.c1 - r.c0 + 1 }, (_, i) => r.c0 + i)); });
    run("cells.autofitRows", () => { const r = selRect(); this.change("Zeilenhöhe", () => this.autoRowHeights(r, true)); });
    run("cells.hideRows", () => { const r = selRect(); this.setRowHidden(r.r0, r.r1, true); });
    run("cells.hideCols", () => { const r = selRect(); this.setColHidden(r.c0, r.c1, true); });
    run("cells.unhideRows", () => { const r = selRect(); this.setRowHidden(r.r0, r.r1, false); });
    run("cells.unhideCols", () => { const r = selRect(); this.setColHidden(r.c0, r.c1, false); });
    run("cells.hideSheet", () => {
      if (this.wb.sheets.filter((s) => !s.hidden).length < 2) return;
      this.change("Ausblenden", () => {
        this.history.sheets();
        this.sheet.hidden = true;
        let i = this.wb.active;
        while (this.wb.sheets[i].hidden) i = (i + 1) % this.wb.sheets.length;
        this.wb.active = i;
      });
    });
    run("cells.unhideSheet", () => void this.dlg.unhideSheet(), () => this.wb.sheets.some((s) => s.hidden));
    run("cells.renameSheet", () => this.sheetBar.renameInPlace(null, this.wb.active));
    run("cells.moveSheet", () => void this.dlg.moveSheet());
    run("cells.tabColor", () => {});
    run("cells.lock", () => { const on = st().locked === false; this.formatSel("Zelle sperren", (s) => { s.locked = on ? undefined : false; }); }, undefined, () => st().locked !== false);
    run("cells.insertMenu", () => void this.commands.run("cells.insert"));
    run("cells.deleteMenu", () => void this.commands.run("cells.delete"));

    run("edit.autosum", () => this.autoSum("SUMME"));
    run("edit.autoAverage", () => this.autoSum("MITTELWERT"));
    run("edit.autoCount", () => this.autoSum("ANZAHL"));
    run("edit.autoMax", () => this.autoSum("MAX"));
    run("edit.autoMin", () => this.autoSum("MIN"));
    const fillDir = (dir: "down" | "right" | "up" | "left") => () => {
      const r = selRect();
      this.change("Ausfüllen", () => {
        if (dir === "down") {
          const src = r.r0 === r.r1 ? { ...r, r0: Math.max(0, r.r0 - 1), r1: Math.max(0, r.r0 - 1) } : { ...r, r1: r.r0 };
          A.autoFill(this.ctx, this.sheet, src, { ...r, r0: src.r0 }, { copy: true });
        } else if (dir === "right") {
          const src = r.c0 === r.c1 ? { ...r, c0: Math.max(0, r.c0 - 1), c1: Math.max(0, r.c0 - 1) } : { ...r, c1: r.c0 };
          A.autoFill(this.ctx, this.sheet, src, { ...r, c0: src.c0 }, { copy: true });
        } else if (dir === "up") A.autoFill(this.ctx, this.sheet, { ...r, r0: r.r1 }, r, { copy: true });
        else A.autoFill(this.ctx, this.sheet, { ...r, c0: r.c1 }, r, { copy: true });
      });
    };
    run("edit.fillDown", fillDir("down"));
    run("edit.fillRight", fillDir("right"));
    run("edit.fillUp", fillDir("up"));
    run("edit.fillLeft", fillDir("left"));
    run("edit.fillSeries", () => void this.dlg.series());
    run("edit.clearAll", () => this.change("Löschen", () => {
      A.clearAll(this.ctx, this.sheet, this.grid.selection, "all");
      // Alles löschen also takes the sparklines out of the cells.
      const sel = this.grid.selection, sheet = this.sheet;
      if (sheet.sparklines?.some((g) => g.items.some((i) => sel.some((r) => i.row >= r.r0 && i.row <= r.r1 && i.col >= r.c0 && i.col <= r.c1)))) {
        this.history.whole(sheet);
        const left = sheet.sparklines.map((g) => ({ ...g, items: g.items.filter((i) => !sel.some((r) => i.row >= r.r0 && i.row <= r.r1 && i.col >= r.c0 && i.col <= r.c1)) })).filter((g) => g.items.length);
        sheet.sparklines = left.length ? left : undefined;
      }
    }));
    run("edit.clearFormats", () => this.change("Löschen", () => A.clearAll(this.ctx, this.sheet, this.grid.selection, "formats")));
    run("edit.clearContents", () => this.change("Löschen", () => A.clearContents(this.ctx, this.sheet, this.grid.selection)));
    run("edit.clearNotes", () => this.change("Löschen", () => A.clearAll(this.ctx, this.sheet, this.grid.selection, "notes")));
    run("edit.clearLinks", () => this.change("Löschen", () => A.clearAll(this.ctx, this.sheet, this.grid.selection, "links")));
    run("edit.find", () => void this.dlg.find("Suchen"));
    run("edit.replace", () => void this.dlg.find("Ersetzen"));
    run("edit.goto", () => void this.dlg.goto());
    run("edit.gotoSpecial", () => void this.gotoSpecial());
    run("edit.selectAll", () => this.grid.select({ r0: 0, c0: 0, r1: MAX_ROWS - 1, c1: MAX_COLS - 1 }, this.sheet.view.active));
    run("edit.dateToday", () => this.typeNow("date"));
    run("edit.timeNow", () => this.typeNow("time"));

    run("edit.sortAscMenu", () => this.dlg.sortQuick(true));
    run("edit.sortDescMenu", () => this.dlg.sortQuick(false));
    run("edit.removeLinks", () => this.change("Links entfernen", () => {
      A.clearAll(this.ctx, this.sheet, this.grid.selection, "links");
      A.format(this.ctx, this.sheet, this.grid.selection, (s) => { if (s.font.color === "#0563C1") { s.font.color = undefined; s.font.underline = undefined; } });
    }));
    run("edit.selectFormulas", () => this.selectWhere((c) => c.f !== undefined, "Formeln"));
    run("edit.selectNotes", () => this.selectWhere((c) => !!c.note, "Notizen"));
    run("edit.selectConstants", () => this.selectWhere((c) => c.f === undefined && c.v !== null && c.v !== "", "Konstanten"));
    run("edit.selectConditional", () => this.selectWhere((_c, row, col) => !!this.sheet.conditional?.some((f) => f.ref.some((r) => row >= r.r0 && row <= r.r1 && col >= r.c0 && col <= r.c1)), "Bedingte Formatierung"));
    run("data.sortAsc", () => this.dlg.sortQuick(true));
    run("data.sortDesc", () => this.dlg.sortQuick(false));
    run("data.sort", () => void this.dlg.sort());
    run("data.filter", () => this.dlg.toggleFilter(), undefined, () => this.dlg.hasFilter());
    run("data.filterClear", () => this.dlg.clearFilter(), () => this.dlg.hasFilter());
    run("data.filterReapply", () => this.dlg.reapplyFilter(), () => this.dlg.hasFilter());
    run("data.textToColumns", () => void this.dlg.textToColumns());
    run("data.removeDuplicates", () => void this.dlg.removeDuplicates());
    run("data.fromText", () => void this.openDialog());

    run("insert.table", () => void this.dlg.formatAsTable());
    run("insert.link", () => void this.dlg.link());
    run("insert.note", () => this.notes.edit());
    run("insert.symbol", () => void this.dlg.symbol());
    run("insert.chartRecommended", () => this.charts.insert("column"));
    run("insert.chartColumn", () => this.charts.insert("column"));
    run("insert.chartLine", () => this.charts.insert("line"));
    run("insert.chartPie", () => this.charts.insert("pie"));
    run("insert.chartScatter", () => this.charts.insert("scatter"));
    run("insert.chartStat", () => this.charts.insert("histogram"));
    run("chart.addElement", () => this.charts.elements(), () => !!this.charts.selected());
    run("chart.changeType", () => {}, () => !!this.charts.selected());

    run("layout.margins", () => {});
    run("layout.orientation", () => {});
    run("layout.portrait", () => this.print.setOrientation("portrait"), undefined, () => this.print.orientation() === "portrait");
    run("layout.landscape", () => this.print.setOrientation("landscape"), undefined, () => this.print.orientation() === "landscape");
    run("layout.size", () => {});
    run("layout.printArea", () => {});
    run("layout.printAreaSet", () => this.print.setPrintArea(selRect()));
    run("layout.printAreaClear", () => this.print.setPrintArea(null));
    run("layout.printTitles", () => void this.print.pageSetupDialog(true));
    run("layout.breaks", () => {});
    run("layout.breakInsert", () => this.print.breaks("insert"));
    run("layout.breakRemove", () => this.print.breaks("remove"), () => { const p = this.sheet.page, a = this.sheet.view.active; return !!(p?.rowBreaks?.includes(a.row) || p?.colBreaks?.includes(a.col)); });
    run("layout.breakReset", () => this.print.breaks("reset"), () => !!(this.sheet.page?.rowBreaks?.length || this.sheet.page?.colBreaks?.length));
    run("insert.headerFooter", () => void this.print.pageSetupDialog("header"));
    run("review.protectSheet", () => void this.protection.protect());
    run("review.protectBook", () => void this.protection.protectBook(), undefined, () => this.protection.bookLocked());
    run("cells.protectSheet", () => void this.protection.protect());
    run("formulas.evaluate", () => void this.evaluator.open(), notEditing);
    run("data.validation", () => void this.validation.openDialog(), notEditing);
    run("edit.flashFill", () => void this.flashFill());
    run("data.group", () => void this.group(1), notEditing);
    run("data.subtotal", () => void this.subtotalDialog(), notEditing);
    run("data.whatIf", () => {});
    run("data.goalSeek", () => void this.goalSeek(), notEditing);
    run("review.spelling", () => void this.spelling.run(), notEditing);
    run("insert.pictures", () => {});
    run("insert.sparkLine", () => void this.sparklineDialog("line"), notEditing);
    run("insert.sparkColumn", () => void this.sparklineDialog("column"), notEditing);
    run("insert.sparkWinLoss", () => void this.sparklineDialog("stacked"), notEditing);
    run("insert.comment", () => this.comments.newComment(), notEditing);
    run("review.newComment", () => this.comments.newComment(), notEditing);
    run("review.deleteComment", () => this.comments.remove(), () => this.comments.has());
    run("review.prevComment", () => this.comments.prev(), () => this.comments.any());
    run("review.nextComment", () => this.comments.next(), () => this.comments.any());
    run("insert.pictureFile", () => void this.pictures.insertFromFile(), notEditing);
    run("data.ungroup", () => void this.group(-1), notEditing);
    run("edit.selectValidation", () => this.validation.select());
    run("data.circleInvalid", () => this.circleInvalid());
    run("data.clearCircles", () => { this.circles.delete(this.sheet); this.grid.contentChanged(); }, () => this.circles.has(this.sheet));
    run("formulas.errorCheck", () => void this.errorCheck.open(), notEditing);
    run("formulas.tracePrecedents", () => this.trace.precedents());
    run("formulas.traceDependents", () => this.trace.dependents());
    run("formulas.removeArrows", () => this.trace.clear(), () => this.trace.any());
    run("layout.pageSetup", () => void this.print.pageSetupDialog());
    run("layout.fitWidth", () => {});
    run("layout.fitHeight", () => {});
    run("layout.scale", () => {});
    run("layout.gridlinesView", () => this.toggleView("showGrid"), undefined, () => this.sheet.view.showGrid);
    run("layout.headingsView", () => this.toggleView("showHeaders"), undefined, () => this.sheet.view.showHeaders);
    run("layout.gridlinesPrint", () => this.print.togglePrintOption("gridLines"), undefined, () => this.print.printOption("gridLines"));
    run("layout.headingsPrint", () => this.print.togglePrintOption("headings"), undefined, () => this.print.printOption("headings"));

    run("formulas.insertFunction", () => void this.dlg.insertFunction());
    for (const id of ["formulas.recent", "formulas.financial", "formulas.logical", "formulas.text", "formulas.dateTime", "formulas.lookup", "formulas.math", "formulas.more"]) run(id, () => {});
    run("formulas.nameManager", () => void this.dlg.nameManager());
    run("formulas.defineName", () => void this.dlg.defineName());
    run("formulas.createFromSelection", () => void this.dlg.defineName());
    run("formulas.showFormulas", () => { this.showFormulas = !this.showFormulas; document.body.classList.toggle("show-formulas", this.showFormulas); this.grid.contentChanged(); }, undefined, () => this.showFormulas);
    run("formulas.calcOptions", () => {});
    // Berechnungsoptionen: Manuell leaves formulas until F9 (the status bar says "Berechnen").
    const setMode = (m: "autoNoTable" | "manual" | undefined) => () => {
      this.wb.calcMode = m;
      if (!m || m === "autoNoTable") { this.active.calc.recalc(); this.grid.contentChanged(); }
      this.active.version++;
      this.scheduleChrome();
    };
    run("formulas.calcAuto", setMode(undefined), undefined, () => !this.wb.calcMode);
    run("formulas.calcPartial", setMode("autoNoTable"), undefined, () => this.wb.calcMode === "autoNoTable");
    run("formulas.calcManual", setMode("manual"), undefined, () => this.wb.calcMode === "manual");
    run("formulas.calcNow", () => { this.active.calc.recalc(); this.statusBar.setFlags([]); this.grid.contentChanged(); });
    run("formulas.calcSheet", () => { this.active.calc.recalc(); this.statusBar.setFlags([]); this.grid.contentChanged(); });

    run("review.statistics", () => void this.dlg.statistics());
    run("review.notes", () => {});
    run("review.newNote", () => this.notes.edit());
    run("review.editNote", () => this.notes.edit(), () => !!this.activeCellNote());
    run("review.deleteNote", () => this.change("Notiz löschen", () => A.clearAll(this.ctx, this.sheet, this.grid.selection, "notes")), () => !!this.activeCellNote());
    run("review.showNotes", () => { this.showNotes = !this.showNotes; this.grid.contentChanged(); }, undefined, () => this.showNotes);

    run("view.normal", () => this.setPageBreakView(false), undefined, () => !this.breakView.has(this.sheet));
    run("view.pageBreak", () => this.setPageBreakView(true), undefined, () => this.breakView.has(this.sheet));
    run("view.gridlines", () => this.toggleView("showGrid"), undefined, () => this.sheet.view.showGrid);
    run("view.headings", () => this.toggleView("showHeaders"), undefined, () => this.sheet.view.showHeaders);
    run("view.formulaBar", () => { document.body.classList.toggle("no-fbar"); }, undefined, () => !document.body.classList.contains("no-fbar"));
    run("view.zoomDialog", () => void this.dlg.zoom());
    run("view.zoom100", () => this.setZoom(100));
    run("view.zoomSelection", () => this.zoomToSelection());
    run("view.freeze", () => {});
    run("view.freezePanes", () => this.freeze(this.sheet.view.active.row, this.sheet.view.active.col));
    run("view.unfreeze", () => this.freeze(0, 0));
    run("view.freezeTop", () => this.freeze(1, 0));
    run("view.freezeFirst", () => this.freeze(0, 1));

    run("sheet.next", () => { let i = this.wb.active + 1; while (i < this.wb.sheets.length && this.wb.sheets[i].hidden) i++; if (i < this.wb.sheets.length) this.activateSheet(i); });
    run("sheet.prev", () => { let i = this.wb.active - 1; while (i >= 0 && this.wb.sheets[i].hidden) i--; if (i >= 0) this.activateSheet(i); });
  }

  private focusRibbonInput(cmd: string): void {
    if (this.ribbon.activeTab !== "start") this.ribbon.select("start");
    requestAnimationFrame(() => {
      this.ribbon.card.querySelector<HTMLInputElement>(`[data-tip^="${this.commands.get(cmd).label}"] input`)?.focus();
    });
  }

  /** Diagrammentwurf appears while a chart is selected, as Excel's contextual tab does. */
  chartSelectionChanged(): void {
    const on = !!this.charts.selected();
    this.ribbon.setContextual(on ? [this.chartTab()] : []);
    if (on) this.ribbon.select("chartDesign");
    this.updateChrome();
  }

  private chartTabDef: import("@raabe/shell").Tab | null = null;
  private chartTab(): import("@raabe/shell").Tab {
    const kinds: [string, import("./charts/feature").ChartKind][] = [["Gruppierte Säulen", "column"], ["Gestapelte Säulen", "columnStacked"], ["Gruppierte Balken", "bar"],
      ["Linie", "line"], ["Linie mit Datenpunkten", "lineMarkers"], ["Fläche", "area"], ["Kreis", "pie"], ["Ring", "doughnut"], ["Punkte (XY)", "scatter"],
      ["Punkte mit geraden Linien", "scatterLines"], ["Histogramm", "histogram"]];
    const b = (cmd: string, key: string): import("@raabe/shell").Control => ({ type: "button", cmd, key, size: "large" });
    return this.chartTabDef ??= {
      id: "chartDesign", label: "Diagrammentwurf", key: "JC", contextual: true, shrink: [],
      groups: [
        { id: "layouts", label: "Diagrammlayouts", key: "ZL", icon: "chart-recommended", controls: [b("chart.addElement", "A"), b("chart.quickLayout", "L")] },
        { id: "styles", label: "Diagrammformatvorlagen", key: "ZS", icon: "styles", controls: [b("chart.colors", "H"), b("chart.styles", "S")] },
        { id: "data", label: "Daten", key: "ZD", icon: "table", controls: [b("chart.switchRowCol", "W"), b("chart.selectData", "E")] },
        { id: "type", label: "Typ", key: "ZT", icon: "chart-column", controls: [{ type: "dropdown", cmd: "chart.changeType", key: "C", size: "large", menu: kinds.map(([label, k]) => ({ label, run: () => this.charts.changeKind(k) })) }] },
        { id: "place", label: "Ort", key: "ZO", icon: "position", controls: [b("chart.move", "V")] },
      ],
    };
  }

  showFormulas = false;
  showNotes = false;

  private activeCellNote(): string | undefined {
    const { row, col } = this.sheet.view.active;
    return this.sheet.get(row, col)?.note;
  }

  private stepFont(d: 1 | -1): void {
    const sizes = [8, 9, 10, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72];
    const cur = this.activeStyle().font.size;
    const next = d > 0 ? sizes.find((s) => s > cur) ?? Math.min(409, cur + 10) : [...sizes].reverse().find((s) => s < cur) ?? Math.max(1, cur - 1);
    this.formatSel("Schriftgrad", (s) => { s.font.size = next; });
    this.change("Zeilenhöhe", () => this.autoRowHeights(this.grid.selection.at(-1)!));
  }

  private toggleView(key: "showGrid" | "showHeaders"): void {
    this.sheet.view[key] = !this.sheet.view[key];
    this.active.version++;
    this.grid.refreshLayout();
    this.updateChrome();
  }

  private freeze(row: number, col: number): void {
    const v = this.sheet.view;
    if (row === 0 && col === 0 || v.freeze) v.freeze = row || col ? { row, col } : undefined;
    else v.freeze = { row, col };
    if (!(row || col)) v.freeze = undefined;
    // Fenster fixieren on A1 freezes the middle of the window (Excel); simplified: top-left visible part.
    if (v.freeze && v.freeze.row === 0 && v.freeze.col === 0) v.freeze = undefined;
    this.active.version++;
    this.grid.refreshLayout();
    this.updateChrome();
  }

  private zoomToSelection(): void {
    const r = this.grid.selection.at(-1)!;
    const el = this.grid.el.getBoundingClientRect();
    const sheet = this.sheet;
    let w = 0, hgt = 0;
    for (let c = r.c0; c <= Math.min(r.c1, r.c0 + 300); c++) { const cw = sheet.cols.get(c)?.width; w += sheet.cols.get(c)?.hidden ? 0 : cw !== undefined ? charsToPx(cw) : defaultColPx(sheet.defaultColWidth ?? DEFAULT_COL_CHARS); }
    for (let rr = r.r0; rr <= Math.min(r.r1, r.r0 + 1000); rr++) hgt += ptToPx(sheet.rows.get(rr)?.height ?? sheet.defaultRowHeight);
    const z = Math.floor(Math.min((el.width - 60) / w, (el.height - 40) / hgt) * 100);
    this.setZoom(Math.max(10, Math.min(400, z)));
    sheet.view.top = r.r0;
    sheet.view.left = r.c0;
    this.grid.refreshLayout();
  }

  private setRowHidden(r0: number, r1: number, hidden: boolean): void {
    this.change(hidden ? "Ausblenden" : "Einblenden", () => {
      for (let r = r0; r <= Math.min(r1, r0 + 100000); r++) {
        const info = this.sheet.rows.get(r);
        if (!hidden && !info?.hidden) continue;
        this.history.row(this.sheet, r);
        this.sheet.rows.set(r, { ...(info ?? {}), hidden: hidden || undefined });
      }
    });
  }

  private setColHidden(c0: number, c1: number, hidden: boolean): void {
    this.change(hidden ? "Ausblenden" : "Einblenden", () => {
      for (let c = c0; c <= c1; c++) {
        const info = this.sheet.cols.get(c);
        if (!hidden && !info?.hidden) continue;
        this.history.col(this.sheet, c);
        const next = { ...(info ?? {}), hidden: hidden || undefined };
        if (!hidden && next.width === 0) next.width = undefined;
        this.sheet.cols.set(c, next);
      }
    });
  }

  /** Strg+. and Strg+Umschalt+. put today's date or the time into the cell being edited or the active cell. */
  private typeNow(kind: "date" | "time"): void {
    const d = new Date();
    const pad = (n: number) => String(n).padStart(2, "0");
    const text = kind === "date" ? `${pad(d.getDate())}.${pad(d.getMonth() + 1)}.${d.getFullYear()}` : `${pad(d.getHours())}:${pad(d.getMinutes())}`;
    if (!this.grid.isEditing) this.grid.startEdit("enter", text);
    else {
      const input = document.querySelector<HTMLTextAreaElement>(".cell-editor.on") ?? document.querySelector<HTMLTextAreaElement>(".fx-input")!;
      const s = input.selectionStart ?? input.value.length;
      input.value = input.value.slice(0, s) + text + input.value.slice(input.selectionEnd ?? s);
      input.dispatchEvent(new Event("input"));
    }
  }

  // ----- keyboard -----------------------------------------------------------------------------------

  private installKeyboard(): void {
    addEventListener("keydown", (e) => {
      if (this.backstage.isOpen) return;
      const t = e.target as HTMLElement;
      const inGrid = t.classList?.contains("cell-editor") || t.classList?.contains("fx-input");
      if (inGrid) return;
      if (t.closest?.(".dialog, .ribbon, .titlebar, .menu, input, textarea, select")) return;
      if (this.shortcut(e)) e.preventDefault();
    });
  }

  /** Excel's shortcuts on the German keyboard (keys by position where Excel uses positions). */
  shortcut(e: KeyboardEvent): boolean {
    // Letters by their character (Z and Y swap places on a German keyboard), digits and
    // punctuation by their position (Strg+Umschalt+4 is "$" on one layout and "4" on the key).
    const ctrl = e.ctrlKey && !e.altKey, shift = e.shiftKey, k = e.key;
    // Umschalt+Alt+Rechts / Links: Gruppieren / Gruppierung aufheben (Excel).
    if (e.altKey && shift && !e.ctrlKey && (k === "ArrowRight" || k === "ArrowLeft")) { void this.commands.run(k === "ArrowRight" ? "data.group" : "data.ungroup"); return true; }
    // Alt+Unten on a cell with a Datenüberprüfung list opens it (before the pick list).
    if (e.altKey && k === "ArrowDown" && this.validation.openList()) return true;
    const code = /^[a-zäöü]$/i.test(k) && /^[a-z]$/i.test(k) ? "Key" + k.toUpperCase() : e.code;
    const run = (id: string) => { void this.commands.run(id); return true; };
    if (this.painter && k === "Escape") { this.painter = null; document.body.classList.remove("painting"); return true; }
    if (k === "F12" && !ctrl) return run("file.saveAs");
    if (k === "F7" && !ctrl && !shift && !e.altKey) return run("review.spelling");
    if (k === "F4" && !e.ctrlKey && !e.altKey && !shift) return run("edit.repeat");
    if (k === "F5" && !ctrl) return run("edit.goto");
    if (k === "F9") return run("formulas.calcNow");
    if (k === "F11" && shift) return run("cells.insertSheet");
    if (k === "F3" && shift && !ctrl) return run("formulas.insertFunction");
    if (k === "F3" && ctrl) return run(shift ? "formulas.createFromSelection" : "formulas.nameManager");
    if (k === "F2" && shift) return run("insert.note");
    if (e.altKey && !e.ctrlKey && (k === "=" || (shift && code === "Digit0"))) return run("edit.autosum");
    if (e.altKey && e.ctrlKey && code === "KeyL") return run("data.filterReapply");
    if (!ctrl) return false;
    if (k === "PageDown") return run("sheet.next");
    if (k === "PageUp") return run("sheet.prev");
    if (shift) {
      switch (code) {
        case "KeyF": return run("font.bold");
        case "KeyK": return run("font.italic");
        case "KeyU": return run("font.underline");
        case "KeyL": return run("data.filter");
        case "KeyA": return run("font.dialog");
        case "KeyP": return run("font.dialog");
        // Recorded in German Excel: Strg+Umschalt+1, 4 and 5 format numbers; 3 and 6 do nothing there.
        case "Digit1": this.numberFormatKey("#,##0.00"); return true;
        case "Digit4": this.numberFormatKey('#,##0.00\\ "€";[Red]\\-#,##0.00\\ "€"'); return true;
        case "Digit5": this.numberFormatKey("0%"); return true;
        case "Digit7": return run("border.outside");
        case "Slash": return run("border.none");
        case "Digit8": case "Digit9": return run("cells.unhideRows");
        case "Digit0": return run("cells.unhideCols");
        case "Period": return run("edit.timeNow");
        case "Equal": case "BracketRight": return run("cells.insert");
      }
      if (k === ":") return run("edit.timeNow");
      return false;
    }
    switch (code) {
      case "KeyS": return run("file.save");
      case "KeyO": return run("file.open");
      case "KeyN": return run("file.new");
      case "KeyP": return run("file.print");
      case "KeyQ": return this.quick.open();
      case "KeyE": return run("edit.flashFill");
      case "KeyW": return run("file.close");
      case "KeyZ": return run("edit.undo");
      case "KeyY": return run("edit.redo");
      case "KeyC": return run("edit.copy");
      case "KeyX": return run("edit.cut");
      case "KeyV": return run("edit.paste");
      case "KeyF": return run("edit.find");
      case "KeyH": return run("edit.replace");
      case "KeyG": return run("edit.goto");
      case "KeyU": return run("edit.fillDown");
      case "KeyR": return run("edit.fillRight");
      case "KeyK": return run("insert.link");
      case "KeyT": case "KeyL": return run("insert.table");
      case "Digit1": return run("cells.formatDialog");
      case "Digit2": return run("font.bold");
      case "Digit3": return run("font.italic");
      case "Digit4": return run("font.underline");
      case "Digit5": return run("font.strike");
      case "Digit9": return run("cells.hideRows");
      case "Period": return run("edit.dateToday");
      case "Backslash": return run("formulas.showFormulas");
      case "BracketRight": case "NumpadAdd": return run("cells.insert");
      case "Slash": case "NumpadSubtract": return run("cells.delete");
    }
    if (k === "+") return run("cells.insert");
    if (k === "-") return run("cells.delete");
    if (k === "#") return run("formulas.showFormulas");
    return false;
  }

  // ----- context menus -------------------------------------------------------------------------------

  private contextMenu(e: MouseEvent, kind: "cell" | "col" | "row"): void {
    const pasteIcons: MenuEntry = { label: "Einfügeoptionen:", icons: [
      { icon: "paste", tip: "Einfügen (E)", run: () => void this.paste() },
      { icon: "number-value", tip: "Werte (W)", run: () => void this.paste("values") },
      { icon: "fx", tip: "Formeln (F)", run: () => void this.paste("formulas") },
      { icon: "table-switch", tip: "Transponieren (T)", run: () => void this.paste("transpose") },
      { icon: "format-painter", tip: "Formatierung (R)", run: () => void this.paste("formats") },
    ] };
    const common: MenuEntry[] = [{ cmd: "edit.cut" }, { cmd: "edit.copy" }, pasteIcons, { cmd: "edit.pasteSpecial" }, "-"];
    let entries: MenuEntry[];
    if (kind === "col") {
      entries = [...common, { label: "Zellen einfügen", run: () => void this.commands.run("cells.insertCols") }, { label: "Zellen löschen", run: () => void this.commands.run("cells.deleteCols") },
        { cmd: "edit.clearContents" }, "-", { cmd: "cells.formatDialog" }, { cmd: "cells.colWidth" }, { label: "Ausblenden", run: () => void this.commands.run("cells.hideCols") }, { label: "Einblenden", run: () => void this.commands.run("cells.unhideCols") }];
    } else if (kind === "row") {
      entries = [...common, { label: "Zellen einfügen", run: () => void this.commands.run("cells.insertRows") }, { label: "Zellen löschen", run: () => void this.commands.run("cells.deleteRows") },
        { cmd: "edit.clearContents" }, "-", { cmd: "cells.formatDialog" }, { cmd: "cells.rowHeight" }, { label: "Ausblenden", run: () => void this.commands.run("cells.hideRows") }, { label: "Einblenden", run: () => void this.commands.run("cells.unhideRows") }];
    } else {
      const note = this.activeCellNote();
      entries = [...common,
        { label: "Zellen einfügen...", run: () => void this.commands.run("cells.insert") }, { label: "Zellen löschen...", run: () => void this.commands.run("cells.delete") }, { cmd: "edit.clearContents" }, "-",
        { label: "Filter", submenu: [{ cmd: "data.filterClear" }, { cmd: "data.filterReapply" }, { label: "Nach dem Wert der ausgewählten Zelle filtern", run: () => this.dlg.filterByActive() }] },
        { label: "Sortieren", submenu: [{ cmd: "data.sortAsc", label: "Nach Größe sortieren (aufsteigend)" }, { cmd: "data.sortDesc", label: "Nach Größe sortieren (absteigend)" }, { cmd: "data.sort", label: "Benutzerdefiniertes Sortieren..." }] },
        "-",
        { label: "Neuer Kommentar", run: () => {}, disabled: true },
        note ? { cmd: "review.editNote" } : { cmd: "review.newNote", label: "Neue Notiz" },
        ...(note ? [{ cmd: "review.deleteNote" } as MenuEntry] : []),
        { cmd: "cells.formatDialog" },
        { label: "Dropdown-Auswahlliste...", run: () => {}, disabled: true },
        { label: "Namen definieren...", run: () => void this.dlg.defineName() },
        { cmd: "insert.link" }];
    }
    showMenu(this.commands, entries, new DOMRect(e.clientX, e.clientY, 0, 0));
  }

  // ----- ribbon context ------------------------------------------------------------------------------

  private ribbonContext(): RibbonContext {
    const fontOptions = () => {
      const list = this.fontList.length ? this.fontList : ["Arial", "Calibri", "Calibri Light", "Cambria", "Times New Roman"];
      return [{ heading: "Designschriftarten", items: ["Calibri Light (Überschriften)", "Calibri (Textkörper)"] }, { heading: "Alle Schriftarten", items: list }];
    };
    const fnMenu = (names: string[]): MenuEntry[] => names.map((n) => ({ label: n, run: () => this.grid.startEdit("enter", `=${n}(`) }));
    const fmtEntry = (label: string, code: string): MenuEntry => ({ label, run: () => this.formatSel("Zahlenformat", (s) => { s.numFmt = code; }) });
    return {
      fontName: () => this.activeStyle().font.name,
      fontSize: () => String(this.activeStyle().font.size).replace(".", ","),
      fontOptions,
      setFont: (name) => this.formatSel("Schriftart", (s) => { s.font.name = name.replace(/ \((Überschriften|Textkörper)\)$/, ""); }),
      setSize: (v) => {
        const n = Number(v.replace(",", "."));
        if (!(n >= 1 && n <= 409)) { void messageBox("RaabeCalc", "Die Zahl muss zwischen 1 und 409 liegen."); return; }
        this.formatSel("Schriftgrad", (s) => { s.font.size = n; });
        this.change("Zeilenhöhe", () => this.autoRowHeights(this.grid.selection.at(-1)!));
      },
      fontColor: () => this.lastFontColor,
      setFontColor: (hex) => { this.lastFontColor = hex; this.formatSel("Schriftfarbe", (s) => { s.font.color = hex || undefined; }); },
      fillColor: () => this.lastFill,
      setFillColor: (hex) => { this.lastFill = hex; this.formatSel("Füllfarbe", (s) => { s.fill = hex || undefined; s.fillXml = undefined; }); },
      numberFormat: () => {
        const code = this.activeStyle().numFmt;
        return RIBBON_FORMATS.find((f) => f.code === code)?.label ?? (code === "General" ? "Standard" : code === "0%" ? "Prozent" : /[dmy]/i.test(code) && !/h/.test(code) ? "Datum" : /0\.0*E/.test(code) ? "Wissenschaft" : "Benutzerdefiniert");
      },
      numberFormats: () => [{ items: RIBBON_FORMATS.map((f) => f.label) }],
      setNumberFormat: (label) => {
        const f = RIBBON_FORMATS.find((x) => x.label === label);
        if (f) this.formatSel("Zahlenformat", (s) => { s.numFmt = f.code; });
      },
      // Livevorschau as in Excel: fonts, sizes and colours show on the selection while the mouse is over them.
      preview: {
        font: (name) => this.grid.preview(name ? (s) => { s.font.name = name.replace(/ \((Überschriften|Textkörper)\)$/, ""); } : null),
        size: (v) => { const n = Number(String(v).replace(",", ".")); this.grid.preview(n > 0 ? (s) => { s.font.size = n; } : null); },
        color: (hex) => this.grid.preview(hex ? (s) => { s.font.color = hex; } : null),
        fill: (hex) => this.grid.preview(hex ? (s) => { s.fill = hex; } : null),
      },
      menus: {
        borders: [
          { heading: "Rahmenlinien" }, { cmd: "border.bottom" }, { cmd: "border.top" }, { cmd: "border.left" }, { cmd: "border.right" }, "-",
          { cmd: "border.none" }, { cmd: "border.all" }, { cmd: "border.outside" }, { cmd: "border.thickOutside" }, "-",
          { cmd: "border.doubleBottom" }, { cmd: "border.thickBottom" }, { cmd: "border.topBottom" }, { cmd: "border.topThickBottom" }, { cmd: "border.topDoubleBottom" },
          { heading: "Rahmenlinien zeichnen" }, { cmd: "border.draw" }, { cmd: "border.drawGrid" }, { cmd: "border.erase" }, { cmd: "border.lineColor" }, { cmd: "border.lineStyle" }, "-",
          { cmd: "border.more" },
        ],
        merge: [{ cmd: "align.mergeCenter" }, { cmd: "align.mergeAcross" }, { cmd: "align.merge" }, { cmd: "align.unmerge" }],
        orientation: [
          { label: "Gegen den Uhrzeigersinn drehen", run: () => this.rotateText(45) },
          { label: "Im Uhrzeigersinn drehen", run: () => this.rotateText(135) },
          { label: "Vertikaler Text", run: () => this.rotateText(255) },
          { label: "Text nach oben drehen", run: () => this.rotateText(90) },
          { label: "Text nach unten drehen", run: () => this.rotateText(180) },
          "-", { label: "Zellenausrichtung formatieren", run: () => void this.dlg.formatCells("Ausrichtung") },
        ],
        accounting: [
          fmtEntry("€ Deutsch (Deutschland)", RIBBON_FORMATS[3].code),
          fmtEntry("$ Englisch (Vereinigte Staaten)", '_-[$$-409]* #,##0.00_ ;_-[$$-409]* \\-#,##0.00\\ ;_-[$$-409]* "-"??_ ;_-@_ '),
          "-", { label: "Weitere Buchhaltungsformate...", run: () => void this.dlg.formatCells("Zahlen") },
        ],
        underline: [{ cmd: "font.underline" }, { cmd: "font.underlineDouble" }],
        conditional: [
          { label: "Regeln zum Hervorheben von Zellen", submenu: () => this.dlg.conditionalHighlightMenu() },
          { label: "Regeln für oberste/unterste Werte", submenu: () => this.dlg.conditionalTopMenu() },
          "-",
          { label: "Datenbalken", submenu: () => this.dlg.conditionalBarMenu() },
          { label: "Farbskalen", submenu: () => this.dlg.conditionalScaleMenu() },
          { label: "Symbolsätze", submenu: () => this.dlg.conditionalIconMenu() },
          "-",
          { label: "Neue Regel...", run: () => {}, disabled: true },
          { label: "Regeln löschen", submenu: [
            { label: "Regeln in ausgewählten Zellen löschen", run: () => this.dlg.clearConditional(true) },
            { label: "Regeln im gesamten Blatt löschen", run: () => this.dlg.clearConditional(false) },
          ] },
          { label: "Regeln verwalten...", run: () => {}, disabled: true },
        ],
        insertCells: [{ cmd: "cells.insertDialog" }, { cmd: "cells.insertRows" }, { cmd: "cells.insertCols" }, "-", { cmd: "cells.insertSheet" }],
        deleteCells: [{ label: "Zellen löschen...", run: () => void this.dlg.deleteCells() }, { cmd: "cells.deleteRows" }, { cmd: "cells.deleteCols" }, "-", { cmd: "cells.deleteSheet" }],
        format: [
          { heading: "Zellengröße" }, { cmd: "cells.rowHeight" }, { cmd: "cells.autofitRows" }, { cmd: "cells.colWidth" }, { cmd: "cells.autofitCols" }, { cmd: "cells.defaultWidth" },
          { heading: "Sichtbarkeit" },
          { label: "Ausblenden und Einblenden", submenu: [{ cmd: "cells.hideRows" }, { cmd: "cells.hideCols" }, { cmd: "cells.hideSheet" }, "-", { cmd: "cells.unhideRows" }, { cmd: "cells.unhideCols" }, { cmd: "cells.unhideSheet" }] },
          { heading: "Blätter anordnen" }, { cmd: "cells.renameSheet" }, { cmd: "cells.moveSheet" },
          { label: "Registerfarbe", submenu: ["#C00000", "#FF0000", "#FFC000", "#FFFF00", "#92D050", "#00B050", "#00B0F0", "#0070C0", "#002060", "#7030A0"].map((col) => ({ label: col, run: () => { if (this.protection.bookLocked()) return; this.change("Registerfarbe", () => { this.history.whole(this.sheet); this.sheet.tabColor = col; }); } })) },
          { heading: "Schutz" }, { cmd: "cells.protectSheet" }, { cmd: "cells.lock" }, { cmd: "cells.formatDialog" },
        ],
        autosum: [{ cmd: "edit.autosum" }, { cmd: "edit.autoAverage" }, { cmd: "edit.autoCount" }, { cmd: "edit.autoMax" }, { cmd: "edit.autoMin" }, "-", { cmd: "formulas.insertFunction", label: "Weitere Funktionen..." }],
        fill: [{ cmd: "edit.fillDown" }, { cmd: "edit.fillRight" }, { cmd: "edit.fillUp" }, { cmd: "edit.fillLeft" }, { cmd: "edit.fillAcross" }, { cmd: "edit.fillSeries" }, { cmd: "edit.justify" }, { cmd: "edit.flashFill" }],
        clear: [{ cmd: "edit.clearAll" }, { cmd: "edit.clearFormats" }, { cmd: "edit.clearContents" }, { cmd: "edit.clearNotes" }, { cmd: "edit.clearLinks" }, "-", { cmd: "edit.removeLinks" }],
        sortFilter: [{ cmd: "edit.sortAscMenu" }, { cmd: "edit.sortDescMenu" }, { label: "Benutzerdefiniertes Sortieren...", run: () => void this.dlg.sort() }, "-", { cmd: "data.filter" }, { cmd: "data.filterClear" }, { cmd: "data.filterReapply" }],
        find: [{ cmd: "edit.find" }, { cmd: "edit.replace" }, { cmd: "edit.goto" }, { cmd: "edit.gotoSpecial" }, "-",
          { cmd: "edit.selectFormulas" }, { cmd: "edit.selectNotes" }, { cmd: "edit.selectConditional" }, { cmd: "edit.selectConstants" }, { cmd: "edit.selectValidation" }, "-",
          { cmd: "edit.selectObjects" }, { cmd: "edit.selectionPane" }],
        paste: [
          { label: "Einfügen", icons: [
            { icon: "paste", tip: "Einfügen (E)", run: () => void this.paste() },
            { icon: "fx", tip: "Formeln (F)", run: () => void this.paste("formulas") },
            { icon: "table-switch", tip: "Transponieren (T)", run: () => void this.paste("transpose") },
          ] },
          { label: "Werte einfügen", icons: [{ icon: "number-value", tip: "Werte (W)", run: () => void this.paste("values") }] },
          { label: "Weitere Einfügeoptionen", icons: [{ icon: "format-painter", tip: "Formatierung (R)", run: () => void this.paste("formats") }] },
          "-", { cmd: "edit.pasteSpecial" },
        ],
        margins: this.print.marginsMenu(),
        orientationPage: [{ cmd: "layout.portrait" }, { cmd: "layout.landscape" }],
        size: this.print.sizeMenu(),
        printArea: [{ cmd: "layout.printAreaSet" }, { cmd: "layout.printAreaClear" }],
        freeze: [{ cmd: "view.freezePanes" }, { cmd: "view.freezeTop" }, { cmd: "view.freezeFirst" }],
        calcOptions: [{ cmd: "formulas.calcAuto" }, { cmd: "formulas.calcPartial" }, { cmd: "formulas.calcManual" }],
        notes: [{ cmd: "review.newNote" }, { cmd: "review.editNote" }, { cmd: "review.deleteNote" }, "-", { cmd: "review.showNotes" }],
        names: [{ cmd: "formulas.defineName" }],
        fnLogical: fnMenu(["FALSCH", "NICHT", "ODER", "UND", "WAHR", "WENN", "WENNFEHLER", "WENNNV", "WENNS", "XODER", "ERSTERWERT"]),
        fnText: fnMenu(["CODE", "ERSETZEN", "FINDEN", "GLÄTTEN", "GROSS", "GROSS2", "IDENTISCH", "KLEIN", "LÄNGE", "LINKS", "RECHTS", "SÄUBERN", "SUCHEN", "TEIL", "TEXT", "TEXTKETTE", "TEXTVERKETTEN", "VERKETTEN", "WECHSELN", "WERT", "WIEDERHOLEN", "ZEICHEN"]),
        fnDate: fnMenu(["ARBEITSTAG", "DATEDIF", "DATUM", "DATWERT", "EDATUM", "HEUTE", "ISOKALENDERWOCHE", "JAHR", "JETZT", "KALENDERWOCHE", "MINUTE", "MONAT", "MONATSENDE", "NETTOARBEITSTAGE", "SEKUNDE", "STUNDE", "TAG", "TAGE", "WOCHENTAG", "ZEIT", "ZEITWERT"]),
        fnLookup: fnMenu(["ADRESSE", "BEREICH.VERSCHIEBEN", "INDEX", "INDIREKT", "SPALTE", "SPALTEN", "SVERWEIS", "VERGLEICH", "WAHL", "WVERWEIS", "XVERWEIS", "ZEILE", "ZEILEN", "FILTER", "SORTIEREN", "EINDEUTIG"]),
        fnMath: fnMenu(["ABRUNDEN", "ABS", "ARCCOS", "ARCSIN", "ARCTAN", "AUFRUNDEN", "BOGENMASS", "COS", "EXP", "FAKULTÄT", "GANZZAHL", "GERADE", "GGT", "GRAD", "KGV", "KOMBINATIONEN", "KÜRZEN", "LN", "LOG", "LOG10", "OBERGRENZE.MATHEMATIK", "PI", "POTENZ", "PRODUKT", "QUADRATESUMME", "REST", "RUNDEN", "SIN", "SUMME", "SUMMENPRODUKT", "SUMMEWENN", "SUMMEWENNS", "TAN", "UNGERADE", "UNTERGRENZE.MATHEMATIK", "VORZEICHEN", "VRUNDEN", "WURZEL", "ZUFALLSBEREICH", "ZUFALLSZAHL"]),
        fnMore: [
          { label: "Statistik", submenu: fnMenu(["ANZAHL", "ANZAHL2", "ANZAHLLEEREZELLEN", "BESTIMMTHEITSMASS", "BINOM.VERT", "GEOMITTEL", "KGRÖSSTE", "KKLEINSTE", "KORREL", "MAX", "MAXWENNS", "MEDIAN", "MIN", "MINWENNS", "MITTELABW", "MITTELWERT", "MITTELWERTWENN", "MITTELWERTWENNS", "MODUS.EINF", "NORM.INV", "NORM.S.VERT", "NORM.VERT", "POISSON.VERT", "QUANTIL.INKL", "QUARTILE.INKL", "RANG.GLEICH", "STABW.N", "STABW.S", "STEIGUNG", "ACHSENABSCHNITT", "VAR.P", "VAR.S", "VARIATIONEN", "ZÄHLENWENN", "ZÄHLENWENNS"]) },
          { label: "Informationen", submenu: fnMenu(["ISTFEHL", "ISTFEHLER", "ISTGERADE", "ISTLEER", "ISTLOG", "ISTNV", "ISTTEXT", "ISTUNGERADE", "ISTZAHL", "NV"]) },
        ],
        fnFinancial: fnMenu(["BW", "KAPZ", "NBW", "RMZ", "ZINS", "ZINSZ", "ZW", "ZZR"]),
        fnRecent: fnMenu(["SUMME", "MITTELWERT", "WENN", "MAX", "MIN", "ANZAHL", "SVERWEIS", "RUNDEN", "ZÄHLENWENN", "SUMMEWENN"]),
        chartColumn: [{ heading: "2D-Säule" }, { label: "Gruppierte Säulen", run: () => this.charts.insert("column") }, { label: "Gestapelte Säulen", run: () => this.charts.insert("columnStacked") },
          { heading: "2D-Balken" }, { label: "Gruppierte Balken", run: () => this.charts.insert("bar") }, { label: "Gestapelte Balken", run: () => this.charts.insert("barStacked") }],
        chartLine: [{ heading: "2D-Linie" }, { label: "Linie", run: () => this.charts.insert("line") }, { label: "Linie mit Datenpunkten", run: () => this.charts.insert("lineMarkers") },
          { heading: "2D-Fläche" }, { label: "Fläche", run: () => this.charts.insert("area") }],
        chartPie: [{ heading: "2D-Kreis" }, { label: "Kreis", run: () => this.charts.insert("pie") }, { heading: "Ring" }, { label: "Ring", run: () => this.charts.insert("doughnut") }],
        chartScatter: [{ heading: "Punkt (XY)" }, { label: "Punkte", run: () => this.charts.insert("scatter") }, { label: "Punkte mit interpolierten Linien", run: () => this.charts.insert("scatterSmooth") },
          { label: "Punkte mit geraden Linien und Datenpunkten", run: () => this.charts.insert("scatterLines") }],
        chartStat: [{ heading: "Histogramm" }, { label: "Histogramm", run: () => this.charts.insert("histogram") }],
        fitWidth: ["Automatisch", "1 Seite", "2 Seiten", "3 Seiten"].map((l, i) => ({ label: l, run: () => this.print.setFit("width", i) })),
        fitHeight: ["Automatisch", "1 Seite", "2 Seiten", "3 Seiten"].map((l, i) => ({ label: l, run: () => this.print.setFit("height", i) })),
      },
      panels: { cellStyles: (close) => this.dlg.cellStylesPanel(close), tableStyles: (close) => this.dlg.tableStylesPanel(close) },
      scale: { value: () => `${this.print.scale()} %`, commit: (v) => this.print.setScale(Number(v.replace(/[^\d]/g, "")) || 100), step: (d) => this.print.setScale(this.print.scale() + d * 5) },
    };
  }

  // ----- backstage --------------------------------------------------------------------------------

  private backstageEntries(): BackstageEntry[] {
    const greeting = () => {
      const hr = new Date().getHours();
      return hr < 11 ? "Guten Morgen" : hr < 18 ? "Guten Tag" : "Guten Abend";
    };
    const thumb = (header: string) => h("div", { class: "thumb calc-thumb" }, ...Array.from({ length: 30 }, (_, i) => h("span", { style: i < 5 && header ? `background:${header}` : "" })));
    const blank = () => h("div", { class: "templates" },
      h("button", { class: "template", onclick: () => { this.backstage.close(); this.newDoc(); } }, thumb(""), h("span", null, "Leere Arbeitsmappe")),
      ...CALC_TEMPLATES.map((t) => h("button", { class: "template", onclick: () => { this.backstage.close(); this.newDoc(t.id); } }, thumb("#F2D7D8"), h("span", null, t.name))));
    return [
      { id: "home", label: "Startseite", icon: "home", page: () => h("div", null, h("h1", null, greeting()), h("h2", null, "Neu"), blank(), h("h2", null, "Zuletzt verwendet"), this.recentTable()) },
      { id: "new", label: "Neu", icon: "new", page: () => h("div", null, h("h1", null, "Neu"), blank()) },
      { id: "open", label: "Öffnen", icon: "open", page: () => h("div", null, h("h1", null, "Öffnen"),
        h("button", { class: "bigbtn", onclick: () => void this.openDialog() }, icon("open", "l"), "Durchsuchen"),
        h("h2", null, "Zuletzt verwendet"), this.recentTable()) },
      { id: "info", label: "Informationen", icon: "info", divider: true, page: () => this.infoPage() },
      { id: "save", label: "Speichern", icon: "save", cmd: "file.save" },
      { id: "saveAs", label: "Speichern unter", icon: "save-as", cmd: "file.saveAs" },
      { id: "print", label: "Drucken", icon: "print", page: () => this.print.page() },
      { id: "export", label: "Exportieren", icon: "export", page: () => h("div", null, h("h1", null, "Exportieren"),
        h("button", { class: "bigbtn", style: "width:200px", onclick: () => void this.print.exportPdf() }, icon("export", "l"), "PDF-Dokument erstellen"),
        h("button", { class: "bigbtn", style: "width:200px", onclick: () => void this.saveCsv() }, icon("save-as", "l"), "Als CSV speichern"),
        h("p", { class: "muted" }, "PDF: zum Abgeben oder Verschicken, sieht auf jedem Gerät gleich aus. CSV: nur die Werte des aktiven Blatts, getrennt durch Semikolons.")) },
      { id: "close", label: "Schließen", icon: "close-doc", cmd: "file.close" },
      { id: "options", label: "Optionen", icon: "settings", cmd: "file.options", bottom: true },
    ];
  }

  private infoPage(): HTMLElement {
    const d = this.active;
    let cells = 0, formulas = 0;
    for (const s of d.wb.sheets) for (const c of s.cells.values()) { if (c.v !== null || c.f) cells++; if (c.f) formulas++; }
    const props = h("table", { class: "props" },
      h("tr", null, h("td", null, "Blätter"), h("td", null, String(d.wb.sheets.length))),
      h("tr", null, h("td", null, "Zellen mit Daten"), h("td", null, cells.toLocaleString("de-DE"))),
      h("tr", null, h("td", null, "Formeln"), h("td", null, formulas.toLocaleString("de-DE"))),
      h("tr", null, h("td", null, "Speicherort"), h("td", null, d.path ?? "Noch nicht gespeichert")),
      h("tr", null, h("td", null, "Zuletzt gespeichert"), h("td", null, d.lastSaved ? new Date(d.lastSaved).toLocaleString("de-DE") : "–")));
    return h("div", null, h("h1", null, "Informationen"), h("h2", { style: "margin-top:0" }, d.name), props);
  }

  private recentTable(): HTMLElement {
    const list = loadRecent();
    if (!list.length) return h("p", { class: "muted" }, "Noch keine zuletzt verwendeten Arbeitsmappen.");
    const sorted = [...list.filter((f) => f.pinned), ...list.filter((f) => !f.pinned)];
    return h("table", { class: "recent" }, ...sorted.map((f) => h("tr", { onclick: () => void this.openPath(f.path) },
      h("td", null, icon("document")), h("td", null, h("div", null, f.name), h("div", { class: "muted" }, f.path)),
      h("td", null, new Date(f.opened).toLocaleDateString("de-DE")),
      h("td", null, h("button", { class: "pinbtn", "aria-label": f.pinned ? "Lösen" : "Anheften", onclick: (e: MouseEvent) => { e.stopPropagation(); setPinned(f.path, !f.pinned); this.backstage.refreshPage(); } }, icon(f.pinned ? "unpin" : "pin"))))));
  }

  /** What a cell shows in its column's width (Standard numbers shorten to fit), as the grid draws it. */
  private shownInColumn(row: number, col: number): string {
    const c = this.sheet.get(row, col);
    const st = this.wb.styles.get(c?.s);
    const cw = this.sheet.cols.get(col)?.width;
    const px = cw !== undefined ? charsToPx(cw) : defaultColPx(this.sheet.defaultColWidth ?? DEFAULT_COL_CHARS);
    return fitShown(this.measureCtx, c, st, px).text;
  }

  /** Test hooks for feel-lab (development and test runs only): the behaviour catalogue (catalogue_calc.py). */
  testApi() {
    const addr = (a: string) => { const r = parseRange(a)!; return r; };
    const excelAddress = (r: Rect) => (r.r0 === 0 && r.r1 === MAX_ROWS - 1 && r.c0 === 0 && r.c1 === MAX_COLS - 1 ? "1:1048576" : rangeName(r));
    const modeName = () => ({ ready: "Bereit", enter: "Eingeben", edit: "Bearbeiten", point: "Zeigen" })[this.grid.mode];
    return {
      app: this,
      grid: this.grid,
      /** A fresh workbook with the case's cells typed in, the start cell selected. */
      reset: (setup: Record<string, string>, start: string, sheets: number) => {
        this.grid.resetKeyModes();
        if (this.grid.isEditing) this.grid.cancelEdit();
        const doc = this.makeDoc(newWorkbook(), this.active.name, null);
        for (let i = 1; i < sheets; i++) doc.wb.sheets.push(new Sheet(`Tabelle${i + 2}`));
        const i = this.docs.indexOf(this.active);
        this.docs[i] = doc;
        this.active = doc;
        this.copied = null;
        this.grid.setMarquee(null);
        doc.history.step("Test", () => { for (const [a, t] of Object.entries(setup)) A.enter(this.ctx, this.sheet, addr(a), t); });
        doc.calc.recalc();
        doc.history.clear();
        this.activate(doc);
        const r = addr(start);
        this.grid.selectCell(r.r0, r.c0);
      },
      /** A template as a new workbook, saved to a given file (check_calc_templates.py). */
      template: async (id: CalcTemplate, path: string) => { this.newDoc(id); await this.platform.writeFile(path, writeXlsx(this.wb)); },
      /** Als PDF exportieren into a given file (check_calc_print.py). */
      pdf: (path: string) => this.print.exportPdf(path),
      save: (path: string) => this.platform.writeFile(path, writeXlsx(this.wb)),
      /** Opens a file as Öffnen does (compare_calc_files.py). */
      open: (path: string) => this.openPath(path),
      /** Where a cell is on the page (CSS px from the viewport's top left). */
      point: (a: string, where: string) => {
        const r = addr(a);
        const b = this.grid.rectOnCanvas(r)!;
        const c = this.grid.el.getBoundingClientRect();
        const x = c.left + b.x, y = c.top + b.y;
        return where === "br" ? { x: x + b.w, y: y + b.h } : where === "right" ? { x: x + b.w, y: y + b.h / 2 } : where === "bottom" ? { x: x + b.w / 2, y: y + b.h } : { x: x + b.w / 2, y: y + b.h / 2 };
      },
      header: (kind: string, which: string) => {
        const c = this.grid.el.getBoundingClientRect();
        if (kind === "col" || kind === "colborder") {
          const r = addr(`${which}1`);
          const b = this.grid.rectOnCanvas(r)!;
          return { x: c.left + b.x + (kind === "col" ? b.w / 2 : b.w), y: c.top + b.y - 10 };
        }
        const b = this.grid.rectOnCanvas(addr(`A${which}`))!;
        return { x: c.left + b.x - 14, y: c.top + b.y + b.h / 2 };
      },
      state: (check: string[]) => {
        const editing = this.grid.isEditing ? (document.querySelector<HTMLTextAreaElement>(".fx-input")!.value) : null;
        const mode = modeName();
        if (this.grid.isEditing) this.grid.cancelEdit();
        const sheet = this.sheet;
        const cells: Record<string, unknown> = {};
        const keys = new Set<number>();
        for (const [k, c] of sheet.cells) if (c.v !== null || c.f) keys.add(k);
        for (const a of check) { const r = addr(a); keys.add(r.r0 * MAX_COLS + r.c0); }
        for (const k of [...keys].sort((a, b) => (a % MAX_COLS) - (b % MAX_COLS) || Math.floor(a / MAX_COLS) - Math.floor(b / MAX_COLS))) {
          const row = Math.floor(k / MAX_COLS), col = k % MAX_COLS;
          const c = sheet.get(row, col);
          const st = this.wb.styles.get(c?.s);
          cells[cellName(row, col)] = {
            formula: formulaLocal(c), text: this.shownInColumn(row, col), format: st.numFmt === "General" ? "Standard" : localFormat(st.numFmt),
            bold: !!st.font.bold, italic: !!st.font.italic, underline: !!st.font.underline, strike: !!st.font.strike, wrap: !!st.wrap,
          };
        }
        const hiddenRows: number[] = [], hiddenCols: number[] = [], widths: Record<string, number> = {};
        for (let r = 0; r < 29; r++) if (sheet.rows.get(r)?.hidden) hiddenRows.push(r + 1);
        for (let c = 0; c < 14; c++) if (sheet.cols.get(c)?.hidden) hiddenCols.push(c + 1);
        for (let c = 0; c < 7; c++) { const w = sheet.cols.get(c)?.width; if (w !== undefined) widths[String(c + 1)] = Math.round(shownChars(w) * 100) / 100; }
        return {
          mode, editing, cells, selection: this.grid.selection.map(excelAddress).join(","), active: cellName(sheet.view.active.row, sheet.view.active.col),
          sheet: sheet.name, sheets: this.wb.sheets.map((s) => s.name), hiddenRows, hiddenCols, widths, filter: !!sheet.autoFilter,
        };
      },
    };
  }
}

function isDateStyle(s: CellStyle): boolean {
  return /[dy]|m{3,}/i.test(s.numFmt.replace(/"[^"]*"|\[[^\]]*\]/g, ""));
}

function clampToUsed(r: Rect, sheet: Sheet): Rect {
  const used = sheet.usedRange();
  if (!used) return { r0: r.r0, c0: r.c0, r1: Math.min(r.r1, r.r0), c1: Math.min(r.c1, r.c0) };
  return { r0: r.r0, c0: r.c0, r1: Math.min(r.r1, Math.max(r.r0, used.r1)), c1: Math.min(r.c1, Math.max(r.c0, used.c1)) };
}

function escapeRe(s: string): string {
  return s.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

function uniqueCopyName(wb: Workbook, name: string): string {
  for (let i = 2; ; i++) {
    const n = `${name} (${i})`;
    if (!wb.sheetByName(n)) return n;
  }
}

function detectEncoding(bytes: Uint8Array): string {
  try {
    new TextDecoder("utf-8", { fatal: true }).decode(bytes);
    return "utf-8";
  } catch {
    return "windows-1252";
  }
}

