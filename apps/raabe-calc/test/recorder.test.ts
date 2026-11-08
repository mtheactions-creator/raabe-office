// Makro aufzeichnen (D-119): the recorder's text helpers, and the macros Excel's recorder wrote by hand
// (results/2026-11-09-excel-recorder.json) run by our interpreter on a workbook without the app: the same cells as in Excel.

import { readFileSync } from "node:fs";
import { describe, expect, it } from "vitest";
import { Calc, Sheet, Workbook, parseRange, toLocalFormula, type Rect } from "@raabe/sheet-engine";
import * as A from "../src/actions";
import { History } from "../src/ops";
import { shown } from "../src/engine";
import type { MacroApp } from "../src/macro-app";
import { runOn, usValue } from "../src/macros";
import { macroShortcuts, validName, wrapLine } from "../src/recorder";

const EXCEL = JSON.parse(readFileSync("tools/feel-lab/results/2026-11-09-excel-recorder.json", "utf-8"));

describe("the recorder's text", () => {
  it("breaks a long statement as Excel's recorder does", () => {
    const line = '    ActiveWorkbook.Worksheets("Tabelle1").Sort.SortFields.Add2 Key:=Range("A2:A4"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal';
    expect(wrapLine(line)).toEqual([
      '    ActiveWorkbook.Worksheets("Tabelle1").Sort.SortFields.Add2 Key:=Range("A2:A4" _',
      "        ), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal",
    ]);
    expect(wrapLine("    Range(\"A1\").Select")).toEqual(["    Range(\"A1\").Select"]);
  });
  it("reads the shortcut attribute", () => {
    const code = EXCEL.cases.Format.code as string;
    expect([...macroShortcuts(code)]).toEqual([["q", "Makro_Format"]]);
    expect(macroShortcuts(EXCEL.cases.Eingabe.code).size).toBe(0);
  });
  it("names", () => {
    expect(validName("Makro1")).toBe(true);
    expect(validName("Noten_färben")).toBe(true);
    expect(validName("A1")).toBe(false);
    expect(validName("1Makro")).toBe(false);
    expect(validName("Mein Makro")).toBe(false);
  });
  it("values in Excel's US form", () => {
    expect(usValue("3.5")).toEqual({ v: 3.5 });
    expect(usValue("50%")).toEqual({ v: 0.5, fmt: "0%" });
    expect(usValue("2/1/2026")).toEqual({ v: 46054, fmt: "date" });
    expect(usValue("Note")).toBeNull();
  });
});

function book(setup: Record<string, string>, start: string) {
  const wb = Workbook.blank();
  const history = new History(wb);
  const ctx = { wb, history };
  history.step("setup", () => { for (const [a, t] of Object.entries(setup)) A.enter(ctx, wb.sheet(), parseRange(a)!, t); });
  const calc = new Calc(wb);
  calc.recalc();
  const select = (sheet: Sheet, r: Rect, active?: { row: number; col: number }) => {
    const i = wb.sheets.indexOf(sheet);
    if (i >= 0) wb.active = i;
    sheet.view.selection = [r];
    sheet.view.active = active ?? { row: r.r0, col: r.c0 };
  };
  const s0 = parseRange(start)!;
  select(wb.sheet(), s0);
  const m = {
    wb, get sheet() { return wb.sheet(); }, calc, history, userName: "Test", alerts: true, testMode: true, answers: [], shown: [],
    grid: { get selection() { return wb.sheet().view.selection; }, select: (r: Rect, a?: { row: number; col: number }) => select(wb.sheet(), r, a) },
    changeAsync: async (_l: string, fn: () => Promise<void>) => { await history.stepAsync("m", fn); },
    recalc: () => calc.recalc(),
    shownText: (s: Sheet, r: number, c: number) => { const x = s.get(r, c); return shown(x, wb.styles.get(x?.s)).text; },
    selectRange: (s: Sheet, r: Rect) => select(s, r),
    activateSheet: (i: number) => { wb.active = i; },
    insertSheetAt: (at: number) => { const s = new Sheet(`Tabelle${wb.sheets.length + 1}`); wb.sheets.splice(at, 0, s); wb.active = at; return s; },
    renameSheet: (i: number, n: string) => { wb.sheets[i].name = n; return null; },
    insertCells: (s: Sheet, r: Rect, dir: string | null) => {
      if (r.c0 === 0 && r.c1 >= 16383) A.insertDelete(ctx, s, "rows", r.r0, r.r1 - r.r0 + 1);
      else if (r.r0 === 0 && r.r1 >= 1048575) A.insertDelete(ctx, s, "cols", r.c0, r.c1 - r.c0 + 1);
      else A.shiftCells(ctx, s, r, dir === "right" ? "right" : "down");
    },
    deleteCells: (s: Sheet, r: Rect) => {
      if (r.c0 === 0 && r.c1 >= 16383) A.insertDelete(ctx, s, "rows", r.r0, -(r.r1 - r.r0 + 1));
      else if (r.r0 === 0 && r.r1 >= 1048575) A.insertDelete(ctx, s, "cols", r.c0, -(r.c1 - r.c0 + 1));
      else A.shiftCells(ctx, s, r, "up");
    },
    sortRows: (s: Sheet, r: Rect, keys: { col: number; desc: boolean }[]) => {
      const rows: unknown[][] = [];
      for (let row = r.r0; row <= r.r1; row++) rows.push(Array.from({ length: r.c1 - r.c0 + 1 }, (_, k) => s.get(row, r.c0 + k)));
      rows.sort((a, b) => { for (const k of keys) { const x = (a[k.col - r.c0] as { v: unknown } | undefined)?.v, y = (b[k.col - r.c0] as { v: unknown } | undefined)?.v; const c = typeof x === "number" && typeof y === "number" ? x - y : String(x).localeCompare(String(y), "de"); if (c) return k.desc ? -c : c; } return 0; });
      rows.forEach((cells, j) => cells.forEach((c, k) => s.set(r.r0 + j, r.c0 + k, c as never)));
    },
    autoFitColumns: () => undefined,
    merge: (s: Sheet, r: Rect, on: boolean) => { s.merges = s.merges.filter((x) => x.r1 < r.r0 || x.r0 > r.r1 || x.c1 < r.c0 || x.c0 > r.c1); if (on) s.merges.push({ ...r }); },
  } as unknown as MacroApp;
  return { wb, m, calc };
}

describe("Excel's recorded macros in our interpreter", () => {
  const cases = Object.entries(EXCEL.cases as Record<string, { code?: string; state?: { cells: Record<string, { formula: string; text: string; bold: boolean; italic: boolean }>; sheets: string[] } }>)
    .filter(([, c]) => c.code && c.state);
  const setups: Record<string, { setup: Record<string, string>; start: string }> = {
    Eingabe: { setup: {}, start: "A1" },
    Format: { setup: { A1: "Name", B1: "Note", C1: "Fach", A2: "Anna", B2: "0,25", A3: "Ben", B3: "0,5" }, start: "A1" },
    Ausfuellen: { setup: { A1: "Mo", B1: "1", B2: "2" }, start: "A1" },
    Sortieren: { setup: { A1: "Name", B1: "Note", A2: "Cem", B2: "2", A3: "Anna", B3: "1", A4: "Ben", B4: "3" }, start: "A2" },
    Zeilen: { setup: { A1: "1", A2: "2", A3: "3", A4: "4", B1: "b" }, start: "A2" },
    Kopieren: { setup: { A1: "a", A2: "b", A3: "=A1&A2" }, start: "A1" },
    Relativ: { setup: {}, start: "B2" },
  };
  for (const [name, c] of cases) {
    if (!setups[name]) continue;
    it(name, async () => {
      const { wb, m, calc } = book(setups[name].setup, setups[name].start);
      const err = await runOn(m, c.code!, `Makro_${name}`);
      calc.recalc();
      expect(err).toBeNull();
      const sheet = wb.sheet();
      const got: Record<string, { formula: string; text: string; bold: boolean; italic: boolean }> = {};
      for (const [k, cell] of sheet.cells) {
        if (cell.v === null && cell.f === undefined) continue;
        const row = Math.floor(k / 16384), col = k % 16384;
        const a = `${String.fromCharCode(65 + col)}${row + 1}`;
        const st = wb.styles.get(cell.s);
        const formula = cell.f !== undefined ? "=" + toLocalFormula(cell.f, wb) : typeof cell.v === "number" ? String(cell.v).replace(".", ",") : String(cell.v);
        got[a] = { formula, text: shown(cell, st).text, bold: !!st.font.bold, italic: !!st.font.italic };
      }
      const want = Object.fromEntries(Object.entries(c.state!.cells).map(([a, x]) => [a, { formula: x.formula, text: x.text, bold: x.bold, italic: x.italic }]));
      expect(got).toEqual(want);
    });
  }
});
