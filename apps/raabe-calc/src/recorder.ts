// Makro aufzeichnen (D-119): the user's actions written down as VBA the way Excel 365's macro recorder writes them,
// recorded in our own Excel by hand (tools/feel-lab/probe_excel_recorder.py, results/2026-11-09-excel-recorder.json).
// - A selection is written only when the next action needs it (Excel keeps the last one of a row of moves:
//   Umschalt+Rechts twice gives one Range("A1:C1").Select), and once more when the recording stops.
// - Entries: ActiveCell.FormulaR1C1 = "..." with the value in Excel's US form ("3.5", "2/1/2026", "50%") and formulas
//   in R1C1 relative to the cell; several cells at once (Strg+Eingabe): Selection.FormulaR1C1.
// - Commands as Excel writes them (Selection.Font.Bold = True, With Selection ... End With blocks, the Sort block ...).
// - Relative Verweise verwenden: selections as ActiveCell.Offset(r, c).Range("A1").Select.
// The finished macro goes into the workbook's VBA project ("Diese Arbeitsmappe"): the first recording of a session
// starts a new module (Modul1, Modul2 ...), the next ones are added to it, as in Excel.

import { h, showDialog } from "@raabe/shell";
import { MAX_COLS, MAX_ROWS, cellName, colName, isDateFormat, rangeName, toR1C1, type CellStyle, type Rect, type Sheet } from "@raabe/sheet-engine";
import type { CalcApp } from "./app";
import { modulesOf, setModules } from "./macros";

const NL = "\n";
const I = "    ";

/** A number in VBA's US form, as the recorder writes values ("3.5", 15 digits). */
function usNumber(v: number): string {
  const s = String(Number(v.toPrecision(15)));
  return s.includes("e") ? v.toPrecision(15).replace(/\.?0+e/, "E") : s;
}

function usDate(serial: number): string {
  const d = new Date(Date.UTC(1899, 11, 30) + Math.round(serial * 86400000));
  const day = `${d.getUTCMonth() + 1}/${d.getUTCDate()}/${d.getUTCFullYear()}`;
  const secs = Math.round((serial - Math.floor(serial)) * 86400);
  if (!secs) return day;
  const hh = Math.floor(secs / 3600), mm = Math.floor((secs % 3600) / 60), ss = secs % 60;
  const t = `${hh % 12 || 12}:${String(mm).padStart(2, "0")}:${String(ss).padStart(2, "0")} ${hh < 12 ? "AM" : "PM"}`;
  return Math.floor(serial) ? `${day} ${t}` : t;
}

const q = (s: string) => `"${s.replace(/"/g, '""')}"`;
const bool = (b: unknown) => (b ? "True" : "False");
const wholeRows = (r: Rect) => r.c0 === 0 && r.c1 === MAX_COLS - 1 && !(r.r0 === 0 && r.r1 === MAX_ROWS - 1);
const wholeCols = (r: Rect) => r.r0 === 0 && r.r1 === MAX_ROWS - 1 && !(r.c0 === 0 && r.c1 === MAX_COLS - 1);
const allCells = (r: Rect) => r.r0 === 0 && r.r1 === MAX_ROWS - 1 && r.c0 === 0 && r.c1 === MAX_COLS - 1;

/**
 * Excel's recorder breaks a long statement with " _": after the token that takes the line past 75 characters, when
 * more follows; the next line is indented four more (recorded: the SortFields.Add2 line).
 */
export function wrapLine(line: string): string[] {
  const toks = line.match(/"(?:[^"]|"")*"|[A-Za-z_][\w.]*:=|[A-Za-z_][\w.]*|\d+(?:\.\d+)?|\s+|./g) ?? [line];
  const indent = /^\s*/.exec(line)![0];
  const out: string[] = [];
  let cur = "";
  for (let i = 0; i < toks.length; i++) {
    cur += toks[i];
    const rest = toks.slice(i + 1).join("");
    if (cur.length > 75 && rest.trim() && !/^\s+$/.test(toks[i]) && toks[i + 1] !== undefined) {
      out.push(cur + " _");
      cur = indent + I;
      while (toks[i + 1] !== undefined && /^\s+$/.test(toks[i + 1])) i++;
    }
  }
  out.push(cur);
  return out;
}

interface Recording { name: string; shortcut: string; description: string; lines: string[] }

/** Where the user is, as the recorder last wrote it: the sheet, and per sheet the selection and the active cell. */
interface Place { sheet: Sheet; sel: string; active: { row: number; col: number } }

/** The module this session's recordings go to (per workbook). */
const sessionModule = new WeakMap<object, string>();

export class MacroRecorder {
  private rec: Recording | null = null;
  relative = false;
  private place: Place | null = null;
  private known = new WeakMap<Sheet, string>();
  /** The last entry written (Strg+Eingabe commits every area: one line). */
  private lastCommit = "";

  constructor(private app: CalcApp) {}

  get recording(): boolean { return this.rec !== null; }

  // ----- what the user does -------------------------------------------------------------------------

  private selKey(): Place {
    const sheet = this.app.sheet, sel = this.app.grid.selection, a = sheet.view.active;
    return { sheet, sel: sel.map(rangeName).join(",") + "|" + cellName(a.row, a.col), active: { row: a.row, col: a.col } };
  }

  /** Takes the present place as known without writing it (after Sheets.Add, ActiveSheet.Next.Select …). */
  sync(): void {
    if (!this.rec) return;
    this.place = this.selKey();
    this.known.set(this.place.sheet, this.place.sel);
    this.lastCommit = "";
  }

  /** Writes the selection when it changed since the last line (and the sheet, when another one was chosen). */
  flush(): void {
    if (!this.rec || !this.place) return;
    const now = this.selKey();
    if (now.sheet !== this.place.sheet) {
      this.rec.lines.push(`${I}Sheets(${q(now.sheet.name)}).Select`);
      // The other sheet's own selection is not written: it is the one the sheet had.
      if (!this.known.has(now.sheet)) this.known.set(now.sheet, now.sel);
      const before = this.known.get(now.sheet);
      this.place = { ...now, sel: before ?? now.sel };
      if (before === now.sel) { this.place = now; return; }
    }
    if (now.sel === this.place.sel) return;
    this.rec.lines.push(...this.selectLines(now));
    this.place = now;
    this.known.set(now.sheet, now.sel);
    this.lastCommit = "";
  }

  private selectLines(now: Place): string[] {
    const sel = this.app.grid.selection;
    const a = now.active;
    const first = sel[0];
    const activeIsFirst = sel.length === 1 && a.row === first.r0 && a.col === first.c0;
    const single = sel.length === 1 && first.r0 === first.r1 && first.c0 === first.c1;
    if (this.relative) {
      const from = this.place!.active;
      const off = (r: number, c: number) => `ActiveCell.Offset(${r - from.row}, ${c - from.col})`;
      const rel = (r: Rect) => rangeName({ r0: r.r0 - first.r0, c0: r.c0 - first.c0, r1: r.r1 - first.r0, c1: r.c1 - first.c0 });
      const out: string[] = [];
      if (sel.length === 1 && wholeRows(first)) out.push(`${I}${off(first.r0, from.col)}.Rows(${q(`1:${first.r1 - first.r0 + 1}`)}).EntireRow.Select`);
      else if (sel.length === 1 && wholeCols(first)) out.push(`${I}${off(from.row, first.c0)}.Columns(${q(`A:${colName(first.c1 - first.c0)}`)}).EntireColumn.Select`);
      else if (sel.length === 1 && allCells(first)) out.push(`${I}Cells.Select`);
      else out.push(`${I}${off(first.r0, first.c0)}.Range(${q(sel.map(rel).join(","))}).Select`);
      if (!activeIsFirst && !single && !(wholeRows(first) && a.row === first.r0) && !(wholeCols(first) && a.col === first.c0)) {
        out.push(`${I}ActiveCell.Offset(${a.row - first.r0}, ${a.col - (wholeRows(first) ? from.col : first.c0)}).Range("A1").Activate`);
      }
      return out;
    }
    let line: string;
    if (sel.length === 1 && allCells(first)) line = "Cells.Select";
    else if (sel.length === 1 && wholeRows(first)) line = `Rows(${q(rangeName(first))}).Select`;
    else if (sel.length === 1 && wholeCols(first)) line = `Columns(${q(rangeName(first))}).Select`;
    else line = `Range(${q(sel.map(rangeName).join(","))}).Select`;
    const out = [I + line];
    if (!activeIsFirst && !single) out.push(`${I}Range(${q(cellName(a.row, a.col))}).Activate`);
    return out;
  }

  /** Lines for an action: the selection first. */
  private emit(...lines: string[]): void {
    if (!this.rec) return;
    this.flush();
    for (const l of lines) this.rec.lines.push(...wrapLine(l));
  }

  /** A cell entry (the grid's commit): the value in Excel's form. */
  commit(targets: Rect, sheet: Sheet): void {
    if (!this.rec) return;
    const a = sheet.view.active;
    const many = this.app.grid.selection.length > 1 || targets.r0 !== targets.r1 || targets.c0 !== targets.c1 || targets.r0 !== a.row || targets.c0 !== a.col;
    const line = `${I}${many ? "Selection" : "ActiveCell"}.FormulaR1C1 = ${q(this.valueText(sheet, targets.r0, targets.c0))}`;
    if (many && line === this.lastCommit) return;
    this.emit(line);
    this.lastCommit = line;
  }

  valueText(sheet: Sheet, r: number, c: number): string {
    const cell = sheet.get(r, c);
    if (!cell) return "";
    if (cell.f !== undefined) return "=" + toR1C1(cell.f, r, c);
    const v = cell.v;
    const st = this.app.wb.styles.get(cell.s);
    if (typeof v === "number") {
      if (isDateFormat(st.numFmt)) return usDate(v);
      if (/%/.test(st.numFmt.replace(/"[^"]*"/g, ""))) return usNumber(v * 100) + "%";
      return usNumber(v);
    }
    if (typeof v === "boolean") return v ? "TRUE" : "FALSE";
    if (v === null || v === undefined) return "";
    return typeof v === "string" ? v : String((v as { code?: string }).code ?? v);
  }

  private style(): CellStyle {
    const a = this.app.sheet.view.active;
    const cell = this.app.sheet.get(a.row, a.col);
    return this.app.wb.styles.get(cell?.s);
  }

  /** After a command ran: Excel's line(s) for it, when the recorder knows it. */
  command(id: string): void {
    if (!this.rec) return;
    const s = this.style();
    const sel = "Selection";
    const align = () => this.alignBlock(s);
    const map: Record<string, () => string[] | null> = {
      "font.bold": () => [`${I}${sel}.Font.Bold = ${bool(s.font.bold)}`],
      "font.italic": () => [`${I}${sel}.Font.Italic = ${bool(s.font.italic)}`],
      "font.strike": () => [`${I}${sel}.Font.Strikethrough = ${bool(s.font.strike)}`],
      "font.underline": () => [`${I}${sel}.Font.Underline = ${s.font.underline ? "xlUnderlineStyleSingle" : "xlUnderlineStyleNone"}`],
      "font.underlineDouble": () => [`${I}${sel}.Font.Underline = ${s.font.underline === "double" ? "xlUnderlineStyleDouble" : "xlUnderlineStyleNone"}`],
      "font.grow": () => this.fontBlock(s), "font.shrink": () => this.fontBlock(s),
      "align.left": align, "align.center": align, "align.right": align, "align.top": align, "align.middle": align, "align.bottom": align,
      "align.wrap": align, "align.indentMore": align, "align.indentLess": align,
      "number.percent": () => [`${I}${sel}.Style = "Percent"`],
      "number.thousands": () => [`${I}${sel}.Style = "Comma"`],
      "number.accounting": () => [`${I}${sel}.NumberFormat = ${q(s.numFmt)}`],
      "number.decimalMore": () => [`${I}${sel}.NumberFormat = ${q(s.numFmt)}`], "number.decimalLess": () => [`${I}${sel}.NumberFormat = ${q(s.numFmt)}`],
      "edit.fillDown": () => [`${I}${sel}.FillDown`], "edit.fillRight": () => [`${I}${sel}.FillRight`],
      "edit.fillUp": () => [`${I}${sel}.FillUp`], "edit.fillLeft": () => [`${I}${sel}.FillLeft`],
      "edit.copy": () => [`${I}${sel}.Copy`], "edit.cut": () => [`${I}${sel}.Cut`],
      "edit.paste": () => [`${I}ActiveSheet.Paste`],
      "edit.clearAll": () => [`${I}${sel}.Clear`], "edit.clearFormats": () => [`${I}${sel}.ClearFormats`],
      "edit.clearContents": () => [`${I}${sel}.ClearContents`], "edit.clearNotes": () => [`${I}${sel}.ClearComments`],
      "cells.hideRows": () => [`${I}${sel}.EntireRow.Hidden = True`], "cells.hideCols": () => [`${I}${sel}.EntireColumn.Hidden = True`],
      "cells.unhideRows": () => [`${I}${sel}.EntireRow.Hidden = False`], "cells.unhideCols": () => [`${I}${sel}.EntireColumn.Hidden = False`],
      "cells.autofitCols": () => [`${I}${sel}.Columns.AutoFit`], "cells.autofitRows": () => [`${I}${sel}.Rows.AutoFit`],
      "cells.insertRows": () => [`${I}${sel}.EntireRow.Insert`], "cells.insertCols": () => [`${I}${sel}.EntireColumn.Insert`],
      "cells.deleteRows": () => [`${I}${sel}.EntireRow.Delete`], "cells.deleteCols": () => [`${I}${sel}.EntireColumn.Delete`],
      "align.merge": () => [`${I}${sel}.Merge`], "align.unmerge": () => [`${I}${sel}.UnMerge`],
    };
    const f = map[id];
    if (!f) return;
    const lines = f();
    if (lines) { for (const l of lines) this.rec.lines.push(...wrapLine(l)); }
    this.lastCommit = "";
  }

  /** Called before a command runs: the selection it acts on is written first. */
  before(id: string): void {
    if (!this.rec || id.startsWith("view.") || id === "edit.undo" || id === "edit.redo") return;
    // Recorded: the pending selection is written before ActiveSheet.Next.Select too.
    this.flush();
  }

  /** Commands that change where the user is (and how Excel writes that). */
  after(id: string, shapeBefore: { rows: boolean; cols: boolean }): void {
    if (!this.rec) return;
    switch (id) {
      case "sheet.next": this.rec.lines.push(`${I}ActiveSheet.Next.Select`); this.sync(); return;
      case "sheet.prev": this.rec.lines.push(`${I}ActiveSheet.Previous.Select`); this.sync(); return;
      case "cells.deleteSheet": this.rec.lines.push(`${I}ActiveWindow.SelectedSheets.Delete`); this.sync(); return;
      // Strg++ / Strg+- on whole rows or columns; on cells a dialog asks first (that path is not recorded yet).
      case "cells.insert": case "cells.insertMenu":
        if (shapeBefore.rows || shapeBefore.cols) this.rec.lines.push(`${I}Selection.Insert Shift:=${shapeBefore.cols ? "xlToRight" : "xlDown"}`);
        return;
      case "cells.delete": case "cells.deleteMenu":
        if (shapeBefore.rows || shapeBefore.cols) this.rec.lines.push(`${I}Selection.Delete Shift:=${shapeBefore.cols ? "xlToLeft" : "xlUp"}`);
        return;
    }
    this.command(id);
  }

  /** The shape of the selection before a command (Strg++ inserts rows or columns by it). */
  shape(): { rows: boolean; cols: boolean } {
    const r = this.app.grid.selection[0];
    return { rows: wholeRows(r), cols: wholeCols(r) };
  }

  /** Schriftart, Schriftgrad: Excel's whole font block. */
  font(): void { this.emit(...this.fontBlock(this.style())); }

  private fontBlock(s: CellStyle): string[] {
    const f = s.font;
    const color = f.color && !f.theme ? `${I}${I}.Color = ${bgr(f.color)}` : `${I}${I}.ThemeColor = xlThemeColorLight1`;
    return [`${I}With Selection.Font`, `${I}${I}.Name = ${q(f.name)}`, `${I}${I}.Size = ${usNumber(f.size)}`, `${I}${I}.Strikethrough = ${bool(f.strike)}`,
      `${I}${I}.Superscript = ${bool(f.vertAlign === "superscript")}`, `${I}${I}.Subscript = ${bool(f.vertAlign === "subscript")}`, `${I}${I}.OutlineFont = False`, `${I}${I}.Shadow = False`,
      `${I}${I}.Underline = ${f.underline === "double" ? "xlUnderlineStyleDouble" : f.underline ? "xlUnderlineStyleSingle" : "xlUnderlineStyleNone"}`,
      color, `${I}${I}.TintAndShade = 0`, `${I}${I}.ThemeFont = ${f.scheme === "minor" ? "xlThemeFontMinor" : f.scheme === "major" ? "xlThemeFontMajor" : "xlThemeFontNone"}`, `${I}End With`];
  }

  private alignBlock(s: CellStyle): string[] {
    const hz = ({ left: "xlLeft", center: "xlCenter", right: "xlRight", justify: "xlJustify", fill: "xlFill", centerContinuous: "xlCenterAcrossSelection", distributed: "xlDistributed" } as Record<string, string>)[s.h ?? ""] ?? "xlGeneral";
    const vt = ({ top: "xlTop", center: "xlCenter", justify: "xlJustify", distributed: "xlDistributed" } as Record<string, string>)[s.v ?? ""] ?? "xlBottom";
    return [`${I}With Selection`, `${I}${I}.HorizontalAlignment = ${hz}`, `${I}${I}.VerticalAlignment = ${vt}`, `${I}${I}.WrapText = ${bool(s.wrap)}`,
      `${I}${I}.Orientation = ${s.rotation ?? 0}`, `${I}${I}.AddIndent = False`, `${I}${I}.IndentLevel = ${s.indent ?? 0}`, `${I}${I}.ShrinkToFit = ${bool(s.shrink)}`,
      `${I}${I}.ReadingOrder = xlContext`, `${I}${I}.MergeCells = False`, `${I}End With`];
  }

  /** Schriftfarbe from the palette. */
  fontColor(): void {
    const f = this.style().font;
    this.emit(`${I}With Selection.Font`, f.color ? `${I}${I}.Color = ${bgr(f.color)}` : `${I}${I}.ColorIndex = xlAutomatic`, `${I}${I}.TintAndShade = 0`, `${I}End With`);
  }

  /** Füllfarbe (the button or the palette). */
  fill(): void {
    const s = this.style();
    if (!s.fill) this.emit(`${I}With Selection.Interior`, `${I}${I}.Pattern = xlNone`, `${I}${I}.TintAndShade = 0`, `${I}${I}.PatternTintAndShade = 0`, `${I}End With`);
    else this.emit(`${I}With Selection.Interior`, `${I}${I}.Pattern = xlSolid`, `${I}${I}.PatternColorIndex = xlAutomatic`, `${I}${I}.Color = ${bgr(s.fill)}`,
      `${I}${I}.TintAndShade = 0`, `${I}${I}.PatternTintAndShade = 0`, `${I}End With`);
  }

  numberFormat(code: string): void { this.emit(`${I}Selection.NumberFormat = ${q(code)}`); }

  /** Dragging the fill handle. */
  autoFill(target: Rect): void {
    if (!this.rec) return;
    this.flush();
    this.rec.lines.push(`${I}Selection.AutoFill Destination:=Range(${q(rangeName(target))}), Type:=xlFillDefault`);
  }

  /** Double-click on a column header's border; dragging it. */
  autofitColumns(cols: number[]): void {
    if (!this.rec || !cols.length) return;
    this.flush();
    const r = rangeName({ r0: 0, r1: MAX_ROWS - 1, c0: Math.min(...cols), c1: Math.max(...cols) });
    this.rec.lines.push(`${I}Columns(${q(r)}).EntireColumn.AutoFit`);
  }

  columnWidth(cols: number[], chars: number): void {
    if (!this.rec || !cols.length) return;
    this.flush();
    const r = rangeName({ r0: 0, r1: MAX_ROWS - 1, c0: Math.min(...cols), c1: Math.max(...cols) });
    this.rec.lines.push(`${I}${cols.length > 1 ? "Selection" : `Columns(${q(r)})`}.ColumnWidth = ${usNumber(Math.round(chars * 100) / 100)}`);
  }

  rowHeight(rows: number[], pt: number): void {
    if (!this.rec || !rows.length) return;
    this.flush();
    const r = `${Math.min(...rows) + 1}:${Math.max(...rows) + 1}`;
    this.rec.lines.push(`${I}${rows.length > 1 ? "Selection" : `Rows(${q(r)})`}.RowHeight = ${usNumber(Math.round(pt * 100) / 100)}`);
  }

  /** Esc ends the copy mode: Application.CutCopyMode = False. */
  copyModeEnded(): void {
    if (!this.rec) return;
    this.rec.lines.push(`${I}Application.CutCopyMode = False`);
  }

  /** The sort of Daten > Sortieren (A-Z, Z-A, the dialog): Excel's Sort block on the sheet. */
  sort(sheet: Sheet, body: Rect, keys: { col: number; asc: boolean }[]): void {
    if (!this.rec) return;
    this.flush();
    const ws = `ActiveWorkbook.Worksheets(${q(sheet.name)}).Sort`;
    const lines = [`${I}${ws}.SortFields.Clear`];
    for (const k of keys) {
      lines.push(`${I}${ws}.SortFields.Add2 Key:=Range(${q(rangeName({ r0: body.r0, r1: body.r1, c0: k.col, c1: k.col }))}), SortOn:=xlSortOnValues, Order:=${k.asc ? "xlAscending" : "xlDescending"}, DataOption:=xlSortNormal`);
    }
    lines.push(`${I}With ${ws}`, `${I}${I}.SetRange Range(${q(rangeName(body))})`, `${I}${I}.Header = xlNo`, `${I}${I}.MatchCase = False`,
      `${I}${I}.Orientation = xlTopToBottom`, `${I}${I}.SortMethod = xlPinYin`, `${I}${I}.Apply`, `${I}End With`);
    for (const l of lines) this.rec.lines.push(...wrapLine(l));
  }

  /** A sheet renamed (double-click on its tab, Format > Blatt umbenennen). */
  renamed(oldName: string, newName: string): void {
    if (!this.rec) return;
    this.flush();
    this.rec.lines.push(`${I}Sheets(${q(oldName)}).Select`, `${I}Sheets(${q(oldName)}).Name = ${q(newName)}`);
    this.sync();
  }

  /** Umschalt+F11 / Einfügen > Blatt einfügen (before the active sheet), the "+" beside the tabs (after it). */
  sheetAdded(after: boolean): void {
    if (!this.rec) return;
    this.flush();
    this.rec.lines.push(after ? `${I}Sheets.Add After:=ActiveSheet` : `${I}Sheets.Add`);
    this.sync();
  }

  // ----- starting and stopping -------------------------------------------------------------------------

  /** Names of the workbook's macros (Subs without arguments). */
  private macroNames(): string[] {
    return modulesOf(this.app).flatMap((m) => [...m.code.matchAll(/^\s*(?:Public\s+)?Sub\s+([\p{L}_][\p{L}\p{N}_]*)\s*\(\s*\)/gimu)].map((x) => x[1]));
  }

  defaultName(): string {
    const names = new Set(this.macroNames().map((n) => n.toLowerCase()));
    for (let i = 1; ; i++) if (!names.has(`makro${i}`)) return `Makro${i}`;
  }

  start(name: string, shortcut: string, description: string): void {
    if (this.app.grid.isEditing) this.app.grid.commitEdit("none");
    this.rec = { name, shortcut, description, lines: [] };
    this.known = new WeakMap();
    this.sync();
    this.app.recordingChanged();
  }

  /** Aufzeichnung beenden: the macro into the session's module. */
  stop(): string | null {
    if (!this.rec) return null;
    if (this.app.grid.isEditing) this.app.grid.commitEdit("none");
    this.flush();
    const r = this.rec;
    this.rec = null;
    const head = [`Sub ${r.name}()`];
    if (r.description) head.push(`Attribute ${r.name}.VB_Description = ${q(r.description)}`);
    head.push(`Attribute ${r.name}.VB_ProcData.VB_Invoke_Func = ${q(`${r.shortcut || " "}\\n14`)}`);
    head.push("'", `' ${r.name} Makro`, r.description ? `' ${r.description}` : "'", "'");
    if (r.shortcut) head.push(`' Tastenkombination: Strg+${r.shortcut === r.shortcut.toUpperCase() && /\p{L}/u.test(r.shortcut) ? "Umschalt+" : ""}${r.shortcut}`, "'");
    const code = [...head, ...r.lines, "End Sub"].join(NL);
    const mods = modulesOf(this.app);
    let modName = sessionModule.get(this.app.wb);
    let mod = modName ? mods.find((m) => m.name === modName) : undefined;
    if (!mod) {
      const used = new Set(mods.map((m) => m.name.toLowerCase()));
      let i = 1;
      while (used.has(`modul${i}`)) i++;
      modName = `Modul${i}`;
      mod = { name: modName, kind: "standard", code: "" };
      mods.push(mod);
      sessionModule.set(this.app.wb, modName);
    }
    mod.code = mod.code ? mod.code.replace(/\s*$/, "") + NL + code : code;
    setModules(this.app, mods);
    this.app.recordingChanged();
    return code;
  }

  /** Ansicht > Makros > Makro aufzeichnen... (and the status bar's button): the dialog, then recording. */
  async dialog(): Promise<void> {
    if (this.rec) { this.stop(); return; }
    const name = h("input", { type: "text", value: this.defaultName(), style: "width:100%", "aria-label": "Makroname" }) as HTMLInputElement;
    const key = h("input", { type: "text", maxlength: 1, style: "width:28px", "aria-label": "Tastenkombination" }) as HTMLInputElement;
    const where = h("select", { style: "width:100%", "aria-label": "Makro speichern in" }, h("option", { value: "this" }, "Diese Arbeitsmappe")) as HTMLSelectElement;
    const desc = h("textarea", { rows: 4, style: "width:100%;resize:none", "aria-label": "Beschreibung" }) as HTMLTextAreaElement;
    const body = h("div", { class: "form record-macro", style: "width:330px;display:grid;gap:4px" },
      h("label", null, h("u", null, "M"), "akroname:"), name,
      h("label", null, h("u", null, "T"), "astenkombination:"), h("div", { style: "display:flex;align-items:center;gap:4px" }, "Strg+", key),
      h("label", null, "Makro speichern ", h("u", null, "i"), "n:"), where,
      h("label", null, h("u", null, "B"), "eschreibung:"), desc);
    setTimeout(() => { name.focus(); name.select(); });
    for (;;) {
      const ok = await showDialog({ title: "Makro aufzeichnen", body, buttons: [{ label: "OK", value: true, primary: true }, { label: "Abbrechen", value: false }], cancel: false });
      if (!ok) return;
      const n = name.value.trim();
      const k = key.value.trim();
      if (!validName(n)) {
        await showDialog({ title: "Microsoft Excel", body: h("div", { style: "max-width:380px" }, "Dieser Makroname ist nicht gültig. Er muss mit einem Buchstaben beginnen und darf nur Buchstaben, Ziffern und Unterstriche enthalten."), buttons: [{ label: "OK", value: true, primary: true }], cancel: true });
        continue;
      }
      if (k && !/^\p{L}$/u.test(k)) {
        await showDialog({ title: "Microsoft Excel", body: h("div", { style: "max-width:380px" }, "Als Tastenkombination ist nur ein Buchstabe möglich."), buttons: [{ label: "OK", value: true, primary: true }], cancel: true });
        continue;
      }
      this.start(n, k, desc.value.replace(/\r?\n/g, " ").trim());
      return;
    }
  }
}

function bgr(hex: string): number {
  const n = parseInt(hex.replace("#", "").slice(-6), 16);
  return ((n >> 16) & 255) | (n & 0xff00) | ((n & 255) << 16);
}

/** A VBA name that is not a cell address (A1, R1C1). */
export function validName(n: string): boolean {
  if (!/^[\p{L}][\p{L}\p{N}_]{0,254}$/u.test(n)) return false;
  if (/^[A-Za-z]{1,3}\d+$/.test(n) || /^[RrZz]\d*[CcSs]\d*$/.test(n)) return false;
  return true;
}

/** Macros with a shortcut (Attribute X.VB_ProcData.VB_Invoke_Func = "q\n14"): letter → macro name. */
export function macroShortcuts(code: string): Map<string, string> {
  const out = new Map<string, string>();
  for (const m of code.matchAll(/^Attribute\s+([\p{L}_][\p{L}\p{N}_]*)\.VB_ProcData\.VB_Invoke_Func\s*=\s*"(\S)\\n\d+"/gimu)) out.set(m[2], m[1]);
  return out;
}
