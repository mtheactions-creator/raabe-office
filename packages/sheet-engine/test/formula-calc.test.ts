// Recalculation (Calc): dependencies, volatile functions, circular references, spilling,
// speed; and the smaller API pieces (formulaRefs, FUNCTIONS_DE, the .xlsx prefixes).

import { describe, expect, it } from "vitest";
import {
  Calc, FUNCTIONS_DE, Sheet, Workbook, err, evaluate, formulaRefs, fromXlsxFormula, parseCell, toFileFormula, toLocalFormula, toXlsxFormula, type Value,
} from "../src";
import { fixtureWorkbook } from "./formula-excel.test";
import fixture from "./fixtures/excel-formulas.json";

function book(cells: Record<string, string | number>, sheets = ["Tabelle1"]): Workbook {
  const wb = new Workbook();
  for (const s of sheets) wb.sheets.push(new Sheet(s));
  for (const [addr, v] of Object.entries(cells)) set(wb, addr, v);
  return wb;
}

/** Sets a cell: "=…" as a German formula, other text and numbers as values. */
function set(wb: Workbook, addr: string, v: string | number | null): { row: number; col: number; sheet: number } {
  const [sname, a] = addr.includes("!") ? addr.split("!") : [wb.sheets[0].name, addr];
  const sheet = wb.sheets.findIndex((s) => s.name === sname);
  const c = parseCell(a)!;
  const sh = wb.sheets[sheet];
  if (v === null) sh.set(c.row, c.col, undefined);
  else if (typeof v === "string" && v.startsWith("=")) sh.set(c.row, c.col, { v: null, f: toFileFormula(v) });
  else sh.set(c.row, c.col, { v });
  return { ...c, sheet };
}

function val(wb: Workbook, addr: string): Value {
  const [sname, a] = addr.includes("!") ? addr.split("!") : [wb.sheets[0].name, addr];
  const c = parseCell(a)!;
  return wb.sheetByName(sname)!.value(c.row, c.col);
}

describe("Calc", () => {
  it("computes the recorded data block like Excel", () => {
    const wb = fixtureWorkbook();
    // Forget the values Excel computed, then compute them ourselves.
    for (const s of wb.sheets) for (const c of s.cells.values()) if (c.f !== undefined) c.v = null;
    const calc = new Calc(wb);
    calc.recalc();
    for (const s of (fixture as unknown as { setup: { sheet: string; addr: string; value: unknown; formula?: string }[] }).setup) {
      if (!s.formula) continue;
      const want = s.value !== null && typeof s.value === "object" ? err((s.value as { e: "#DIV/0!" }).e) : s.value;
      expect(val(wb, `${s.sheet}!${s.addr}`), s.formula).toEqual(want);
    }
  });

  it("gives the same results as the stored values for every recorded case", () => {
    const wb = fixtureWorkbook();
    const calc = new Calc(wb);
    calc.recalc();
    const cases = (fixture as unknown as { cases: { addr: string; file?: string; refused?: boolean; group: string }[] }).cases;
    let n = 0;
    for (const c of cases) {
      if (c.refused || !c.file || ["ZUFALLSZAHL", "HEUTE"].includes(c.group)) continue;
      const at = parseCell(c.addr)!;
      expect(calc.evaluate(0, at.row, at.col, c.file), c.file).toEqual(evaluate(wb, 0, at.row, at.col, c.file));
      n++;
    }
    expect(n).toBeGreaterThan(2400);
  });

  it("recomputes dependents only and reports changed cells", () => {
    const wb = book({ A1: 1, A2: "=A1*2", A3: "=A2+1", B1: "=SUMME(A1:A3)", C1: 7, C2: "=C1" });
    const calc = new Calc(wb);
    calc.recalc();
    expect([val(wb, "A2"), val(wb, "A3"), val(wb, "B1"), val(wb, "C2")]).toEqual([2, 3, 6, 7]);
    const at = set(wb, "A1", 5);
    const changed = calc.changed(0, [at]).map((c) => `${c.row},${c.col}`).sort();
    expect(changed).toEqual(["0,1", "1,0", "2,0"]);
    expect([val(wb, "A2"), val(wb, "A3"), val(wb, "B1")]).toEqual([10, 11, 26]);
  });

  it("follows new and removed formulas", () => {
    const wb = book({ A1: 2, B1: "=A1+1" });
    const calc = new Calc(wb);
    calc.recalc();
    calc.changed(0, [set(wb, "B1", "=A1*10")]);
    expect(val(wb, "B1")).toBe(20);
    calc.changed(0, [set(wb, "A1", 3)]);
    expect(val(wb, "B1")).toBe(30);
    calc.changed(0, [set(wb, "B1", 4)]);
    calc.changed(0, [set(wb, "A1", 9)]);
    expect(val(wb, "B1")).toBe(4);
  });

  it("works across sheets and through defined names", () => {
    const wb = book({ A1: 1, "Daten!A1": "=Tabelle1!A1*3", B1: "=Daten!A1+Wert", C1: 10 }, ["Tabelle1", "Daten"]);
    wb.names.push({ name: "Wert", ref: "Tabelle1!$C$1" });
    const calc = new Calc(wb);
    calc.recalc();
    expect(val(wb, "B1")).toBe(13);
    calc.changed(0, [set(wb, "A1", 2)]);
    expect(val(wb, "Daten!A1")).toBe(6);
    expect(val(wb, "B1")).toBe(16);
    calc.changed(0, [set(wb, "C1", 0)]);
    expect(val(wb, "B1")).toBe(6);
  });

  it("recomputes volatile functions on every change", () => {
    let r = 0.25;
    const wb = book({ A1: "=ZUFALLSZAHL()", B1: "=A1*2", C1: 1, D1: "=INDIREKT(\"C\"&1)" });
    const calc = new Calc(wb, { random: () => r, now: () => new Date(2024, 0, 15, 12, 0, 0) });
    calc.recalc();
    expect(val(wb, "B1")).toBe(0.5);
    r = 0.5;
    calc.changed(0, [set(wb, "C1", 2)]);
    expect(val(wb, "B1")).toBe(1);
    expect(val(wb, "D1")).toBe(2);
    expect(calc.evaluate(0, 9, 9, "TODAY()")).toBe(45306);
    expect(calc.evaluate(0, 9, 9, "NOW()")).toBe(45306.5);
  });

  it("gives 0 in circular references and reports them", () => {
    const wb = book({ A1: "=B1+1", B1: "=A1", C1: "=A1+10" });
    const calc = new Calc(wb);
    calc.recalc();
    expect([val(wb, "A1"), val(wb, "B1")]).toEqual([0, 0]);
    expect(val(wb, "C1")).toBe(10);
    expect(calc.circular.map((c) => `${c.row},${c.col}`).sort()).toEqual(["0,0", "0,1"]);
    calc.changed(0, [set(wb, "B1", 5)]);
    expect(val(wb, "A1")).toBe(6);
    expect(val(wb, "C1")).toBe(16);
    expect(calc.circular).toEqual([]);
    // A formula referring to itself.
    calc.changed(0, [set(wb, "D1", "=D1+1")]);
    expect(val(wb, "D1")).toBe(0);
    expect(calc.circular.map((c) => `${c.row},${c.col}`)).toEqual(["0,3"]);
  });

  it("spills dynamic arrays, blocks with #ÜBERLAUF! and follows A1#", () => {
    const wb = book({ B1: "=SEQUENZ(3)", C1: "=SUMME(B1#)", D1: "=B3" });
    const calc = new Calc(wb);
    calc.recalc();
    expect([val(wb, "B1"), val(wb, "B2"), val(wb, "B3"), val(wb, "C1"), val(wb, "D1")]).toEqual([1, 2, 3, 6, 3]);
    expect(calc.spillAnchor(0, 2, 1)).toEqual({ row: 0, col: 1 });
    // Typing into the spill range blocks it.
    calc.changed(0, [set(wb, "B2", "x")]);
    expect(val(wb, "B1")).toEqual(err("#SPILL!"));
    expect(val(wb, "B3")).toBe(null);
    expect(val(wb, "D1")).toBe(0);
    // Clearing the cell lets it spill again.
    calc.changed(0, [set(wb, "B2", null)]);
    expect([val(wb, "B1"), val(wb, "B2"), val(wb, "B3"), val(wb, "D1")]).toEqual([1, 2, 3, 3]);
    // A smaller result empties the cells it no longer needs.
    calc.changed(0, [set(wb, "B1", "=SEQUENZ(2)")]);
    expect([val(wb, "B2"), val(wb, "B3"), val(wb, "C1"), val(wb, "D1")]).toEqual([2, null, 3, 0]);
    expect(wb.sheets[0].get(0, 1)!.spill).toEqual({ r0: 0, c0: 1, r1: 1, c1: 1 });
  });

  it("handles a chain of 20,000 formulas in either direction", () => {
    const n = 20000;
    const down = new Workbook();
    down.sheets.push(new Sheet("T"));
    const sh = down.sheets[0];
    sh.set(0, 0, { v: 1 });
    for (let r = 1; r < n; r++) sh.set(r, 0, { v: null, f: `A${r}+1` });
    // The reverse chain needs its last cell first.
    sh.set(n - 1, 1, { v: 1 });
    for (let r = 0; r < n - 1; r++) sh.set(r, 1, { v: null, f: `B${r + 2}+1` });
    const calc = new Calc(down);
    calc.recalc();
    expect(sh.value(n - 1, 0)).toBe(n);
    expect(sh.value(0, 1)).toBe(n);
    sh.set(0, 0, { v: 2 });
    calc.changed(0, [{ row: 0, col: 0 }]);
    expect(sh.value(n - 1, 0)).toBe(n + 1);
  });

  it("recomputes after one edit among 10,000 formulas well under 10 ms", () => {
    const wb = new Workbook();
    wb.sheets.push(new Sheet("T"));
    const sh = wb.sheets[0];
    for (let r = 0; r < 1000; r++) sh.set(r, 0, { v: r });
    // 10,000 formulas: cell references, ranges and a column of running totals.
    for (let r = 0; r < 2500; r++) {
      sh.set(r, 1, { v: null, f: `A${(r % 1000) + 1}*2` });
      sh.set(r, 2, { v: null, f: `SUM(A${(r % 900) + 1}:A${(r % 900) + 100})` });
      sh.set(r, 3, { v: null, f: `IF(B${r + 1}>100,"groß","klein")` });
      sh.set(r, 4, { v: null, f: `ROUND(C${r + 1}/7,2)` });
    }
    const calc = new Calc(wb);
    calc.recalc();
    let best = Infinity;
    for (let i = 0; i < 5; i++) {
      sh.set(950, 0, { v: i });
      const t = performance.now();
      calc.changed(0, [{ row: 950, col: 0 }]);
      best = Math.min(best, performance.now() - t);
    }
    expect(sh.value(950, 1)).toBe(8);
    expect(best).toBeLessThan(10);
  });

  it("fills array formulas (Strg+Umschalt+Eingabe) into their range", () => {
    const wb = book({ A1: 1, A2: 2, A3: 3 });
    wb.sheets[0].set(0, 1, { v: null, f: "A1:A3*10", array: { r0: 0, c0: 1, r1: 3, c1: 1 } });
    const calc = new Calc(wb);
    calc.recalc();
    expect([val(wb, "B1"), val(wb, "B2"), val(wb, "B3"), val(wb, "B4")]).toEqual([10, 20, 30, err("#N/A")]);
  });
});

describe("formula helpers", () => {
  it("finds references with their place in the typed text", () => {
    const f = "=SUMME(A1:B3;'Daten 2'!C4)+$D$5";
    const refs = formulaRefs(f);
    expect(refs.map((r) => f.slice(r.start, r.end))).toEqual(["A1:B3", "'Daten 2'!C4", "$D$5"]);
    expect(refs[0].ref).toEqual({ r0: 0, c0: 0, r1: 2, c1: 1 });
    expect(refs[1].sheet).toBe("Daten 2");
    expect(formulaRefs("=WENN(A1>0;\"B2\";C:C)").map((r) => r.ref)).toEqual([{ r0: 0, c0: 0, r1: 0, c1: 0 }, { r0: 0, c0: 2, r1: 1048575, c1: 2 }]);
  });

  it("lists the German functions with argument hints", () => {
    const sverweis = FUNCTIONS_DE.find((f) => f.name === "SVERWEIS");
    expect(sverweis?.args).toBe("Suchkriterium; Matrix; Spaltenindex; [Bereich_Verweis]");
    expect(FUNCTIONS_DE.some((f) => f.name === "ZÄHLENWENN")).toBe(true);
  });

  it("converts partly typed formulas without throwing", () => {
    expect(toFileFormula("=SUMME(A1;")).toBe("SUM(A1,");
    expect(toLocalFormula("SUM(A1,\"x;y\",1.5)")).toBe("SUMME(A1;\"x;y\";1,5)");
  });

  it("adds and removes the file's _xlfn. prefixes", () => {
    expect(toXlsxFormula("XLOOKUP(A1,B:B,C:C)+SUM(STDEV.S(A1:A3))")).toBe("_xlfn.XLOOKUP(A1,B:B,C:C)+SUM(_xlfn.STDEV.S(A1:A3))");
    expect(toXlsxFormula("FILTER(A1:A3,A1:A3>1)")).toBe("_xlfn._xlws.FILTER(A1:A3,A1:A3>1)");
    expect(fromXlsxFormula("_xlfn._xlws.SORT(_xlfn.UNIQUE(A1:A9))")).toBe("SORT(UNIQUE(A1:A9))");
  });
});
