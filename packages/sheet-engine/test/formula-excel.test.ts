// Replays the formulas recorded from Excel 365 (tools/feel-lab/probe_excel_formulas.py,
// fixtures/excel-formulas.json) in our engine: values, spilled arrays, the German ⇄ file
// form, copying (shiftFormula), inserting and deleting rows and columns (adjustFormula) and
// moving cells. Differences we accept are listed in KNOWN with the reason.

import { describe, expect, it } from "vitest";
import { parseCell, parseRange, Workbook, Sheet, err, type ErrorCode, type Value } from "../src";
import { adjustFormula, evaluate, shiftFormula, toFileFormula, toLocalFormula } from "../src";
import fixture from "./fixtures/excel-formulas.json";

interface Enc { e: string }
type EV = number | string | boolean | null | Enc;
interface Case { group: string; local: string; addr: string; refused?: boolean; file?: string; back?: string; value?: EV; text?: string; spill?: EV[][] }
interface Shift { formula: string; from: string; to: string; result: string }
interface Adjust { kind: "rows" | "cols"; at: number; count: number; formula: string; same: string; otherBefore: string; other: string }
interface Move { from: string; dRow: number; dCol: number; formula: string; same: string; otherBefore?: string; other?: string; inside?: boolean }
interface CellsCase { range: string; dir: "down" | "right" | "up" | "left"; formula: string; other: string }
interface Fixture {
  cells: CellsCase[];
  shift: Shift[];
  adjust: Adjust[];
  move: Move[];
  sheets: string[];
  names: { name: string; ref: string }[];
  setup: { sheet: string; addr: string; value: EV; formula?: string }[];
  cases: Case[];
}

const FIX = fixture as unknown as Fixture;

function dec(v: EV): Value {
  if (v !== null && typeof v === "object") return err(v.e as ErrorCode);
  return v;
}

export function fixtureWorkbook(): Workbook {
  const wb = new Workbook();
  for (const name of FIX.sheets) wb.sheets.push(new Sheet(name));
  for (const s of FIX.setup) {
    const sh = wb.sheetByName(s.sheet)!;
    const c = parseCell(s.addr)!;
    sh.set(c.row, c.col, { v: dec(s.value), ...(s.formula ? { f: s.formula } : {}) });
  }
  wb.names = FIX.names.map((n) => ({ name: n.name, ref: n.ref }));
  return wb;
}

function same(a: Value, b: Value): boolean {
  if (typeof a === "number" && typeof b === "number") {
    if (a === b) return true;
    return Math.abs(a - b) <= 1e-12 * Math.max(Math.abs(a), Math.abs(b));
  }
  if (a !== null && b !== null && typeof a === "object" && typeof b === "object") return a.code === b.code;
  return a === b;
}

function show(v: unknown): string {
  return JSON.stringify(v, (_k, x) => (x && typeof x === "object" && "code" in x ? x.code : x));
}

/** Cases whose result depends on things we do not model, with the reason. */
const KNOWN: Record<string, string> = {
  "=SIN(PI())": "at a zero of the sine Excel's own sine differs from IEEE's in the 4th digit (both about 1,2E-16)",
  "=TAN(PI()/2)": "at the pole of the tangent Excel's own tangent differs in the 4th digit (both about 1,6E+16)",
  "=ZINS(10;-100;800)": "Excel stops its Newton iteration earlier: 1E-11 relative from the exact root we return",
  "=VRUNDEN(6,05;0,1)": "Excel gives 6; every other recorded case fits rounding the quotient at 15 digits, which gives 6,1",
  "=DATEDIF(36556;36556+400;\"yd\")": "31.1.2000 to 6.3.2001: Excel counts 34 days as if 2000 had no 29 February; the other 60 recorded cases fit",
};

export function valueMismatches(): string[] {
  const wb = fixtureWorkbook();
  const out: string[] = [];
  for (const c of FIX.cases) {
    if (c.refused || c.file === undefined) continue;
    if (KNOWN[c.local]) continue;
    const at = parseCell(c.addr)!;
    let got: Value | Value[][];
    try {
      got = evaluate(wb, 0, at.row, at.col, c.file);
    } catch (e) {
      out.push(`${c.group} ${c.local}: threw ${(e as Error).message}`);
      continue;
    }
    if (c.spill) {
      const want = c.spill.map((r) => r.map(dec));
      const ok = Array.isArray(got) && got.length === want.length && got.every((r, i) => r.length === want[i].length && r.every((v, j) => same(v, want[i][j])));
      if (!ok) out.push(`${c.group} ${c.local}: got ${show(got)} want ${show(want)}`);
    } else {
      const want = dec(c.value ?? null);
      if (Array.isArray(got) || !same(got, want)) out.push(`${c.group} ${c.local}: got ${show(got)} want ${show(want)}`);
    }
  }
  return out;
}

export function conversionMismatches(): string[] {
  const out: string[] = [];
  const wb = fixtureWorkbook();
  for (const c of FIX.cases) {
    if (c.refused || c.file === undefined) continue;
    const f = toFileFormula(c.local, wb);
    if (f !== c.file) out.push(`toFile ${c.local}: got ${f} want ${c.file}`);
    const l = toLocalFormula(c.file, wb);
    if ("=" + l !== c.back) out.push(`toLocal ${c.file}: got ${l} want ${c.back!.slice(1)}`);
  }
  return out;
}

describe("formulas recorded from Excel", () => {
  const cases = FIX.cases.filter((c) => !c.refused);
  it(`computes the values of ${cases.length} cases`, () => {
    const bad = valueMismatches();
    if (process.env.FORMULA_REPORT) console.log(bad.join("\n"));
    expect(bad.length, bad.slice(0, 40).join("\n")).toBe(0);
  });
  it("converts between the German and the file form", () => {
    const bad = conversionMismatches();
    if (process.env.FORMULA_REPORT) console.log(bad.join("\n"));
    expect(bad.length, bad.slice(0, 40).join("\n")).toBe(0);
  });
});

// Formulas entered through the legacy Formula property show a deleted range as "@#REF!".
const noAt = (f: string) => f.replace(/@/g, "");

describe("references recorded from Excel", () => {
  it(`copies formulas (${FIX.shift.length} cases)`, () => {
    const bad: string[] = [];
    for (const c of FIX.shift) {
      const a = parseCell(c.from)!, b = parseCell(c.to)!;
      const got = shiftFormula(c.formula, b.row - a.row, b.col - a.col);
      if (got !== c.result) bad.push(`${c.formula} ${c.from}->${c.to}: got ${got} want ${c.result}`);
    }
    expect(bad, bad.join("\n")).toEqual([]);
  });
  it(`inserts and deletes rows and columns (${FIX.adjust.length * 2} cases)`, () => {
    const bad: string[] = [];
    for (const c of FIX.adjust) {
      const at = c.at - 1;
      const same = adjustFormula(c.formula, { kind: c.kind, sheet: "Adj", at, count: c.count, home: "Adj" });
      if (same !== noAt(c.same)) bad.push(`${c.kind} ${c.at} ${c.count} ${c.formula}: got ${same} want ${noAt(c.same)}`);
      const other = adjustFormula(c.otherBefore, { kind: c.kind, sheet: "Adj", at, count: c.count, home: "Andere" });
      if (other !== noAt(c.other)) bad.push(`${c.kind} ${c.at} ${c.count} ${c.otherBefore}: got ${other} want ${noAt(c.other)}`);
    }
    expect(bad, bad.join("\n")).toEqual([]);
  });
  it(`inserts and deletes cells with shifting (${FIX.cells.length} cases)`, () => {
    const bad: string[] = [];
    for (const c of FIX.cells) {
      const got = adjustFormula(c.formula, { kind: "cells", sheet: "Adj", rect: parseRange(c.range)!, dir: c.dir, home: "Andere" });
      if (got !== noAt(c.other)) bad.push(`${c.range} ${c.dir} ${c.formula}: got ${got} want ${noAt(c.other)}`);
    }
    expect(bad, bad.join("\n")).toEqual([]);
  });
  it(`moves cells (${FIX.move.length} cases)`, () => {
    const bad: string[] = [];
    for (const c of FIX.move) {
      const from = parseRange(c.from)!;
      const change = { kind: "move" as const, sheet: "Adj", from, dRow: c.dRow, dCol: c.dCol, home: "Andere" };
      if (c.inside) {
        const got = adjustFormula(c.formula, { ...change, home: "Adj" });
        if (got !== c.same) bad.push(`inside ${c.from} ${c.dRow},${c.dCol}: got ${got} want ${c.same}`);
        continue;
      }
      const got = adjustFormula(c.otherBefore!, change);
      if (got !== noAt(c.other!)) bad.push(`${c.from} ${c.dRow},${c.dCol} ${c.otherBefore}: got ${got} want ${noAt(c.other!)}`);
    }
    expect(bad, bad.join("\n")).toEqual([]);
  });
});
