// Structured references against Excel's own results (tools/feel-lab/probe_excel_structref.py, D-080):
// the value, the form Excel shows and the text the .xlsx file stores.

import { describe, expect, it } from "vitest";
import {
  Calc, Sheet, Workbook, mapStructs, parseCell, structToFile, structToShown, toFileFormula, toLocalFormula, toXlsxFormula, type Value,
} from "../src";
import fixture from "./fixtures/excel-structref.json";

interface Case { cell: string; typed: string; local?: string; formula?: string; value?: unknown; rejected?: boolean }

function book(): Workbook {
  const wb = new Workbook();
  const sh = new Sheet("Tabelle1");
  wb.sheets.push(sh);
  (fixture.data as (string | number)[][]).forEach((r, i) => r.forEach((v, k) => sh.set(i, k, { v })));
  sh.tables = [{
    name: "Noten", ref: { r0: 0, c0: 0, r1: 5, c1: 2 }, header: true, totals: true, style: "TableStyleMedium2",
    columns: [{ name: "Name", totalsLabel: "Ergebnis" }, { name: "Note" }, { name: "Punkte", totalsFunction: "sum" }],
    firstColumn: false, lastColumn: false, rowStripes: true, columnStripes: false, filterButton: true,
  }];
  sh.set(5, 0, { v: "Ergebnis" });
  sh.set(5, 2, { v: null, f: "SUBTOTAL(109,[Punkte])" });
  return wb;
}

const own = (wb: Workbook, row: number, col: number) =>
  wb.sheets[0].tables!.find((t) => row >= t.ref.r0 && row <= t.ref.r1 && col >= t.ref.c0 && col <= t.ref.c1)?.name ?? null;

/** What RaabeCalc stores for a typed formula: the file form with structured references as Excel shows them. */
function entered(wb: Workbook, typed: string, row: number, col: number): string {
  return mapStructs(toFileFormula(typed), (x) => structToShown(x, own(wb, row, col)));
}

function expected(v: unknown): Value | string {
  if (v && typeof v === "object" && "error" in (v as object)) return (v as { error: string }).error;
  return v as Value;
}

describe("structured references as Excel computes, shows and saves them", () => {
  const cases = (fixture.outside as Case[]).filter((c) => !c.rejected);
  const wb = book();
  for (const c of cases) {
    const at = parseCell(c.cell)!;
    wb.sheets[0].set(at.row, at.col, { v: null, f: entered(wb, c.typed, at.row, at.col) });
  }
  new Calc(wb).recalc();

  for (const c of cases) {
    it(c.typed, () => {
      const at = parseCell(c.cell)!;
      const cell = wb.sheets[0].get(at.row, at.col)!;
      const v = cell.v;
      expect(v instanceof Object ? String(v) : v).toEqual(expected(c.value));
      expect("=" + toLocalFormula(cell.f!)).toBe(c.local);
    });
  }

  it("the Ergebniszeile", () => {
    const cell = wb.sheets[0].get(5, 2)!;
    expect(cell.v).toBe(fixture.totals_formula.value);
    expect("=" + toLocalFormula(cell.f!)).toBe(fixture.totals_formula.local);
    expect(toXlsxFormula(mapStructs(cell.f!, (x) => structToFile(x, "Noten")))).toBe(fixture.file.cells.C6);
  });

  it("the file's text (cells right of the table moved when Excel added a column)", () => {
    const file = fixture.file.cells as Record<string, string>;
    for (const c of cases) {
      const at = parseCell(c.cell)!;
      const key = at.row <= 5 ? c.cell.replace("G", "H") : c.cell;
      const f = wb.sheets[0].get(at.row, at.col)!.f!;
      // Excel writes the table's name alone as Noten[].
      const ours = toXlsxFormula(mapStructs(f, (x) => structToFile(x, own(wb, at.row, at.col)), (n) => (n === "Noten" ? "Noten[]" : n)));
      expect(ours, c.typed).toBe(file[key]);
    }
  });

  it("a berechnete Spalte", () => {
    const inside = fixture.inside as Case[];
    expect(structToFile("[@Punkte]", "Noten") + "*2").toBe(fixture.file.calculatedColumnFormula[0]);
    for (const c of inside) expect(c.local).toBe("=[@Punkte]*2");
  });
});
