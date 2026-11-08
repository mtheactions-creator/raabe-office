// What Excel's object model for macros (macros.ts, D-091) needs from RaabeCalc, in one place: the workbook, undo steps,
// sheets, sizes, formats, dialogs and saving. Built on the app's own operations, so a macro changes the workbook the way
// the user's commands do (one undo step per run, formulas worked out at once).

import { h, showDialog } from "@raabe/shell";
import {
  Sheet, fileFormat, localFormat, type Rect, type Workbook,
} from "@raabe/sheet-engine";
import { writeXlsx } from "@raabe/ooxml";
import { VbaError } from "@raabe/vba";
import * as A from "./actions";
import type { CalcApp } from "./app";
import { shown } from "./engine";
import { legacyHash } from "./protection";

/** The app's private helpers the facade uses (the app keeps them private for its own UI). */
interface Inner {
  activateSheet(i: number): void;
  moveSheet(from: number, to: number, copy: boolean): void;
  nextSheetName(): string;
  settings: { userName: string };
  statusBar: { setNote(key: string, text: [string, string] | null): void };
  active: { path: string | null; name: string; version: number; savedVersion: number };
  autoRowHeights(r: Rect, force?: boolean): void;
  platform: { writeFile(path: string, data: Uint8Array): Promise<void>; saveFileDialog(name: string, kind: string): Promise<string | null> };
}

export interface MacroApp {
  readonly wb: Workbook;
  readonly sheet: Sheet;
  readonly calc: CalcApp["calc"];
  readonly history: CalcApp["history"];
  readonly grid: CalcApp["grid"];
  readonly userName: string;
  /** DisplayAlerts. */
  alerts: boolean;
  /** Test runs: MsgBox and InputBox answered from `answers`, no security question, errors not shown. */
  testMode: boolean;
  answers: (string | number)[];
  shown: string[];
  changeAsync(label: string, fn: () => Promise<void>): Promise<void>;
  recalc(): void;
  shownText(sheet: Sheet, row: number, col: number): string;
  localFormat(code: string): string;
  fileFormat(local: string): string;
  selectRange(sheet: Sheet, r: Rect): void;
  activateSheet(i: number): void;
  insertSheetAt(at: number): Sheet;
  deleteSheetAt(i: number): void;
  copySheet(i: number, at: number): void;
  moveSheetTo(i: number, at: number): void;
  renameSheet(i: number, name: string): string | null;
  setSheetVisible(sheet: Sheet, state: "visible" | "hidden" | "veryHidden"): void;
  setTabColor(sheet: Sheet, hex: string | null): void;
  columnWidth(sheet: Sheet, col: number): number;
  setColumnWidth(sheet: Sheet, col: number, chars: number): void;
  rowHeight(sheet: Sheet, row: number): number;
  setRowHeight(sheet: Sheet, row: number, pt: number): void;
  autoFitColumns(sheet: Sheet, c0: number, c1: number): void;
  merge(sheet: Sheet, r: Rect, on: boolean): void;
  insertCells(sheet: Sheet, r: Rect, dir: "down" | "right" | null): void;
  deleteCells(sheet: Sheet, r: Rect, dir: "up" | "left" | null): void;
  sortRows(sheet: Sheet, r: Rect, keys: { col: number; desc: boolean }[]): void;
  setNote(sheet: Sheet, row: number, col: number, text: string | null): void;
  protectSheet(sheet: Sheet, password: string): void;
  unprotectSheet(sheet: Sheet, password: string): boolean;
  statusText(text: string | null): void;
  msgBox(prompt: string, options: [string, number][], title: string): Promise<number>;
  inputBox(prompt: string, title: string, def: string): Promise<string | null>;
  vba(): Uint8Array | undefined;
  setVba(bytes: Uint8Array | undefined): void;
  docTitle(): string;
  docPath(): string | null;
  saveTo(path: string | null): Promise<boolean>;
}

export function macroApp(app: CalcApp): MacroApp {
  const inner = app as unknown as Inner;
  const ctx = () => ({ wb: app.wb, history: app.history });
  const m: MacroApp = {
    get wb() { return app.wb; },
    get sheet() { return app.sheet; },
    get calc() { return app.calc; },
    get history() { return app.history; },
    get grid() { return app.grid; },
    get userName() { return inner.settings.userName; },
    alerts: true,
    testMode: false,
    answers: [],
    shown: [],
    async changeAsync(label, fn) {
      if (app.grid.isEditing) app.grid.commitEdit("none");
      await app.history.stepAsync(label, fn);
    },
    recalc: () => app.calc.recalc(),
    shownText: (sheet, row, col) => { const c = sheet.get(row, col); return shown(c, app.wb.styles.get(c?.s)).text; },
    localFormat: (code) => localFormat(code),
    fileFormat: (local) => fileFormat(local),
    selectRange(sheet, r) {
      const i = app.wb.sheets.indexOf(sheet);
      if (i >= 0 && i !== app.wb.active) inner.activateSheet(i);
      app.grid.select(r);
    },
    activateSheet: (i) => inner.activateSheet(i),
    insertSheetAt(at) {
      app.history.sheets();
      const s = new Sheet(inner.nextSheetName());
      s.defaultColWidth = app.sheet.defaultColWidth;
      app.wb.sheets.splice(at, 0, s);
      app.wb.active = at;
      return s;
    },
    deleteSheetAt(i) {
      if (app.wb.sheets.length < 2) return;
      app.history.sheets();
      app.wb.sheets.splice(i, 1);
      app.wb.active = Math.min(app.wb.active, app.wb.sheets.length - 1);
    },
    copySheet: (i, at) => inner.moveSheet(i, at, true),
    moveSheetTo: (i, at) => inner.moveSheet(i, at, false),
    renameSheet: (i, name) => app.renameSheet(i, name),
    setSheetVisible(sheet, state) {
      const sheets = app.wb.sheets, i = sheets.indexOf(sheet);
      // Excel keeps one sheet visible (Laufzeitfehler 1004) and, hiding the active one, shows the next visible
      // sheet to its right, else to its left.
      if (state !== "visible" && !sheet.hidden && sheets.filter((s) => !s.hidden).length < 2) throw new VbaError("Die Visible-Eigenschaft des Worksheet-Objektes kann nicht festgelegt werden.", 1004);
      app.history.sheets();
      sheet.hidden = state === "visible" ? undefined : state === "veryHidden" ? "very" : true;
      if (sheet.hidden && i === app.wb.active) {
        const next = [...sheets.slice(i + 1), ...sheets.slice(0, i).reverse()].find((s) => !s.hidden);
        if (next) app.wb.active = sheets.indexOf(next);
      }
    },
    setTabColor(sheet, hex) { app.history.whole(sheet); sheet.tabColor = hex ?? undefined; },
    columnWidth: (sheet, col) => sheet.cols.get(col)?.width ?? sheet.defaultColWidth ?? 8.43,
    setColumnWidth(sheet, col, chars) { app.history.col(sheet, col); const i = sheet.cols.get(col); sheet.cols.set(col, { ...(i ?? {}), width: chars }); },
    rowHeight: (sheet, row) => sheet.rows.get(row)?.height ?? sheet.defaultRowHeight ?? 15,
    setRowHeight(sheet, row, pt) { app.history.row(sheet, row); const i = sheet.rows.get(row); sheet.rows.set(row, { ...(i ?? {}), height: pt, custom: true }); },
    autoFitColumns(sheet, c0, c1) {
      if (sheet !== app.sheet) return;
      app.autofitCols(Array.from({ length: c1 - c0 + 1 }, (_, k) => c0 + k));
    },
    merge(sheet, r, on) {
      app.history.whole(sheet);
      sheet.merges = sheet.merges.filter((x) => x.r1 < r.r0 || x.r0 > r.r1 || x.c1 < r.c0 || x.c0 > r.c1);
      if (on && (r.r1 > r.r0 || r.c1 > r.c0)) sheet.merges.push({ ...r });
    },
    insertCells(sheet, r, dir) {
      if (dir === null ? r.c0 === 0 && r.c1 >= 16383 : false) A.insertDelete(ctx(), sheet, "rows", r.r0, r.r1 - r.r0 + 1);
      else if (dir === null && r.r0 === 0 && r.r1 >= 1048575) A.insertDelete(ctx(), sheet, "cols", r.c0, r.c1 - r.c0 + 1);
      else A.shiftCells(ctx(), sheet, r, dir === "right" ? "right" : "down");
    },
    deleteCells(sheet, r, dir) {
      if (r.c0 === 0 && r.c1 >= 16383) A.insertDelete(ctx(), sheet, "rows", r.r0, -(r.r1 - r.r0 + 1));
      else if (r.r0 === 0 && r.r1 >= 1048575) A.insertDelete(ctx(), sheet, "cols", r.c0, -(r.c1 - r.c0 + 1));
      else A.shiftCells(ctx(), sheet, r, dir === "left" ? "left" : "up");
    },
    sortRows(sheet, r, keys) {
      app.history.cells(sheet, r);
      const rows: { cells: (import("@raabe/sheet-engine").Cell | undefined)[]; i: number }[] = [];
      for (let row = r.r0; row <= r.r1; row++) rows.push({ i: row, cells: Array.from({ length: r.c1 - r.c0 + 1 }, (_, k) => sheet.get(row, r.c0 + k)) });
      const val = (x: (typeof rows)[number], col: number) => x.cells[col - r.c0]?.v ?? null;
      const rank = (v: unknown) => (typeof v === "number" ? 0 : typeof v === "string" ? 1 : typeof v === "boolean" ? 2 : v === null ? 4 : 3);
      rows.sort((a, b) => {
        for (const k of keys) {
          const x = val(a, k.col), y = val(b, k.col);
          // Empty cells go last in both directions (Excel).
          if (x === null || y === null) { if (x !== y) return x === null ? 1 : -1; continue; }
          let c = rank(x) - rank(y);
          if (!c) c = typeof x === "number" ? (x as number) - (y as number) : String(x).localeCompare(String(y), "de", { sensitivity: "base" });
          if (c) return k.desc ? -c : c;
        }
        return a.i - b.i;
      });
      rows.forEach((x, j) => x.cells.forEach((c, k) => sheet.set(r.r0 + j, r.c0 + k, c)));
    },
    setNote(sheet, row, col, text) {
      app.history.cell(sheet, row, col);
      const c = sheet.get(row, col);
      sheet.set(row, col, { ...(c ?? { v: null }), note: text ?? undefined });
    },
    protectSheet(sheet, password) {
      app.history.whole(sheet);
      sheet.protection = { sheet: "1", objects: "1", scenarios: "1", ...(password ? { password: legacyHash(password) } : {}) };
    },
    unprotectSheet(sheet, password) {
      const p = sheet.protection;
      if (!p) return true;
      if (p.password && p.password !== legacyHash(password)) return false;
      if (p.hashValue && !password) return false;
      app.history.whole(sheet);
      sheet.protection = undefined;
      return true;
    },
    statusText: (text) => inner.statusBar.setNote("macro", text === null ? null : [text, text]),
    async msgBox(prompt, options, title) {
      m.shown.push(prompt);
      if (m.testMode) { const a = m.answers.shift(); return typeof a === "number" ? a : options[0][1]; }
      return showDialog({ title, body: h("div", { style: "white-space:pre-wrap;max-width:420px" }, prompt), buttons: options.map(([label, value], i) => ({ label, value, primary: i === 0 })), cancel: options[options.length - 1][1] });
    },
    async inputBox(prompt, title, def) {
      m.shown.push(prompt);
      if (m.testMode) { const a = m.answers.shift(); return a === undefined ? def : String(a); }
      const input = h("input", { type: "text", value: def, style: "width:320px" }) as HTMLInputElement;
      const ok = await showDialog({ title, body: h("div", { class: "form" }, h("div", { style: "white-space:pre-wrap" }, prompt), input), buttons: [{ label: "OK", value: true, primary: true }, { label: "Abbrechen", value: false }], cancel: false });
      return ok ? input.value : null;
    },
    vba: () => app.wb.vba,
    setVba(bytes) { app.wb.vba = bytes; },
    docTitle: () => inner.active.name,
    docPath: () => inner.active.path,
    async saveTo(path) {
      if (path === null) return app.save();
      await inner.platform.writeFile(path, writeXlsx(app.wb));
      return true;
    },
  };
  return m;
}
