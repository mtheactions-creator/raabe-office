// LET, LAMBDA and its helpers against Excel (tools/feel-lab/probe_excel_lambda.py, D-080): values and
// spills, the German form, the file's text with _xlfn. and _xlpm., defined names holding a LAMBDA.

import { describe, expect, it } from "vitest";
import { Calc, Sheet, Workbook, bindingProblem, fromXlsxFormula, parseCell, toFileFormula, toLocalFormula, toXlsxFormula } from "../src";
import fixture from "./fixtures/excel-lambda.json";

interface Case { cell: string; typed: string; local?: string; value?: unknown; spill?: string | null; rejected?: boolean }

const norm = (v: unknown): unknown => {
  if (v && typeof v === "object" && "error" in (v as object)) return (v as { error: string }).error;
  if (Array.isArray(v)) return v.map(norm);
  return v;
};

describe("LET and LAMBDA as Excel computes and saves them", () => {
  const wb = new Workbook();
  const sh = new Sheet("Tabelle1");
  wb.sheets.push(sh);
  for (const [a, v] of Object.entries(fixture.data)) {
    const c = parseCell(a)!;
    sh.set(c.row, c.col, { v: v as number });
  }
  for (const [name, f] of fixture.names as [string, string][]) wb.names.push({ name, ref: f.slice(1) });
  const cases = (fixture.cases as Case[]).filter((c) => !c.rejected);
  for (const c of cases) {
    const at = parseCell(c.cell)!;
    sh.set(at.row, at.col, { v: null, f: toFileFormula(c.typed) });
  }
  new Calc(wb).recalc();

  for (const c of cases) {
    it(c.typed, () => {
      const at = parseCell(c.cell)!;
      const cell = sh.get(at.row, at.col)!;
      expect("=" + toLocalFormula(cell.f!)).toBe(c.local);
      const want = norm(c.value);
      if (Array.isArray(want)) {
        const rows = want.length, cols = (want[0] as unknown[]).length;
        const got: unknown[][] = [];
        for (let r = 0; r < rows; r++) {
          const row: unknown[] = [];
          for (let k = 0; k < cols; k++) {
            const v = sh.get(at.row + r, at.col + k)?.v;
            row.push(v && typeof v === "object" ? String(v) : v);
          }
          got.push(row);
        }
        expect(got).toEqual(want);
      } else {
        const v = cell.v;
        expect(v && typeof v === "object" ? String(v) : v).toEqual(want);
      }
      const file = (fixture.file.cells as Record<string, string>)[c.cell];
      expect(toXlsxFormula(cell.f!)).toBe(file);
      expect(fromXlsxFormula(file)).toBe(cell.f);
    });
  }

  it("refuses a name bound twice, as Excel does on entry", () => {
    const rejected = (fixture.cases as Case[]).filter((c) => c.rejected).map((c) => c.typed);
    expect(rejected).toEqual(["=LET(x;1;x;2;x)"]);
    for (const r of rejected) expect(bindingProblem(toFileFormula(r))).toBe(true);
    for (const c of cases) expect(bindingProblem(toFileFormula(c.typed)), c.typed).toBe(false);
  });

  it("writes defined names holding a LAMBDA as Excel does", () => {
    for (const [name, f] of fixture.names as [string, string][]) {
      expect(toXlsxFormula(f.slice(1))).toBe((fixture.file.names as Record<string, string>)[name]);
      expect("=" + toLocalFormula(f.slice(1))).toBe((fixture.names_local as Record<string, string>)[name]);
    }
  });
});
