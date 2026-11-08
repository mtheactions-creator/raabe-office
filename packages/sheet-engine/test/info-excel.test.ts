// ZELLE, INFO, BLATT, BLÄTTER and HYPERLINK against German Excel (tools/feel-lab/probe_excel_info.py, D-080).

import { describe, expect, it } from "vitest";
import { Calc, Sheet, Workbook, fileFormat, parseCell, toFileFormula, toLocalFormula } from "../src";
import { cellFormatCode } from "../src/formula/fn-info";
import fixture from "./fixtures/excel-info.json";

interface Case { cell?: string; typed: string; local?: string; value?: unknown; rejected?: boolean }

/**
 * Where RaabeCalc differs on purpose (D-082): ZELLE without a reference reads the formula's own cell
 * (Excel: the cell changed last); ZELLE("Breite") counts RaabeCalc's default column (Calibri, D-011);
 * INFO("Ursprung") is the sheet's top left cell as RaabeCalc scrolled it.
 */
const KNOWN = new Set(['=ZELLE("Zeile")', '=ZELLE("Breite";A1)', '=ZELLE("Breite";A2)', '=ZELLE("Breite";A11)', '=ZELLE("width";A1)',
  '=ZELLE("width";A2)', '=ZELLE("width";A11)', '=ZELLE("Breite";C1)']);

const norm = (v: unknown): unknown => {
  if (v && typeof v === "object" && "error" in (v as object)) return (v as { error: string }).error;
  if (Array.isArray(v)) return v.map(norm);
  return v;
};

describe("ZELLE, INFO, BLATT, BLÄTTER, HYPERLINK as German Excel", () => {
  const wb = new Workbook();
  for (const n of fixture.sheets as string[]) wb.sheets.push(new Sheet(n));
  wb.docName = "Mappe1";
  const sh = wb.sheets[0];
  for (const [a, v, fmt] of fixture.setup as [string, string | number, string][]) {
    const c = parseCell(a)!;
    const s = fmt === "Standard" ? undefined : wb.styles.derive(undefined, (x) => { x.numFmt = fileFormat(fmt); });
    sh.set(c.row, c.col, { v, s });
  }
  const cases = (fixture.cases as Case[]).filter((c) => !c.rejected);
  for (const c of cases) {
    const at = parseCell(c.cell!)!;
    sh.set(at.row, at.col, { v: null, f: toFileFormula(c.typed) });
  }
  new Calc(wb).recalc();

  for (const c of cases) {
    if (KNOWN.has(c.typed)) continue;
    it(c.typed, () => {
      const at = parseCell(c.cell!)!;
      const cell = sh.get(at.row, at.col)!;
      expect("=" + toLocalFormula(cell.f!)).toBe(c.local);
      const want = norm(c.value);
      if (Array.isArray(want)) {
        const row = (want as unknown[][])[0].map((_, k) => sh.get(at.row, at.col + k)?.v);
        expect([row]).toEqual(want);
      } else {
        const v = cell.v;
        expect(v && typeof v === "object" ? String(v) : v).toEqual(want);
      }
    });
  }

  it("German Excel's ZELLE(\"Format\") codes for 29 number formats", () => {
    for (const f of fixture.formats as { format: string; Format?: string; Farbe?: number; failed?: boolean }[]) {
      if (f.failed) continue;
      const code = f.format === "Standard" ? "General" : fileFormat(f.format);
      expect(cellFormatCode(code), f.format).toBe(f.Format);
    }
  });
});
