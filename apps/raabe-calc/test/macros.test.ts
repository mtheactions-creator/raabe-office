// Excel's object model for macros (D-095) on a workbook without the app: the facade's members the cases use, the
// rest unused. The Excel comparison is check_calc_macros.py; these run the same kinds of statements fast.

import { describe, expect, it } from "vitest";
import { Calc, Workbook, parseRange, type Sheet } from "@raabe/sheet-engine";
import * as A from "../src/actions";
import { History } from "../src/ops";
import { shown } from "../src/engine";
import type { MacroApp } from "../src/macro-app";
import { runOn } from "../src/macros";

function book(cells: Record<string, string>) {
  const wb = Workbook.blank();
  const history = new History(wb);
  const ctx = { wb, history };
  const sheet = wb.sheet();
  history.step("setup", () => { for (const [a, t] of Object.entries(cells)) A.enter(ctx, sheet, parseRange(a)!, t); });
  const calc = new Calc(wb);
  calc.recalc();
  const shownMsgs: string[] = [];
  const m = {
    wb, get sheet() { return wb.sheet(); }, calc, history, grid: {}, userName: "Test", alerts: true, testMode: true, answers: [], shown: shownMsgs,
    changeAsync: async (_l: string, fn: () => Promise<void>) => { await history.stepAsync("m", fn); },
    recalc: () => calc.recalc(),
    shownText: (s: Sheet, r: number, c: number) => { const x = s.get(r, c); return shown(x, wb.styles.get(x?.s)).text; },
    deleteCells: (s: Sheet, r: { r0: number; r1: number; c0: number; c1: number }) => A.insertDelete(ctx, s, "rows", r.r0, -(r.r1 - r.r0 + 1)),
  } as unknown as MacroApp;
  const v = (a: string) => { const r = parseRange(a)!; return sheet.get(r.r0, r.c0)?.v ?? null; };
  const run = async (body: string) => { const err = await runOn(m, `Sub T()\n${body}\nEnd Sub`, "T"); calc.recalc(); return err; };
  return { wb, sheet, v, run };
}

describe("Excel's object model", () => {
  it("writes and reads cells", async () => {
    const b = book({ A1: "A", A3: "B", B3: "2" });
    expect(await b.run('Range("C1").Value = 5')).toBeNull();
    expect(b.v("C1")).toBe(5);
    expect(await b.run('Range("C2").Value = Cells(3, 2).Value + 1')).toBeNull();
    expect(b.v("C2")).toBe(3);
    expect(await b.run('Range("C3").Value = Rows.Count')).toBeNull();
    expect(b.v("C3")).toBe(1048576);
  });

  it("finds the last row with End(xlUp) and counts with WorksheetFunction", async () => {
    const b = book({ A1: "A", A3: "B", B3: "2", A6: "C", A7: "D", B9: "9" });
    expect(await b.run('Range("D1").Value = Cells(Rows.Count, 1).End(xlUp).Row')).toBeNull();
    expect(b.v("D1")).toBe(7);
    expect(await b.run('Range("D2").Value = WorksheetFunction.CountA(Rows(3))')).toBeNull();
    expect(b.v("D2")).toBe(2);
  });

  it("deletes empty rows from the bottom", async () => {
    const b = book({ A1: "A", A3: "B", B3: "2", A6: "C", A7: "D", B9: "9" });
    expect(await b.run([
      "Dim i As Long",
      "For i = 9 To 1 Step -1",
      "  If WorksheetFunction.CountA(Rows(i)) = 0 Then Rows(i).Delete",
      "Next i",
    ].join("\n"))).toBeNull();
    expect(["A1", "A2", "A3", "A4", "B5"].map(b.v)).toEqual(["A", "B", "C", "D", 9]);
  });
});
