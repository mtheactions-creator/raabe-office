// Replays the results recorded from Excel 365 for the second round of functions
// (tools/feel-lab/probe_excel_functions2.py, fixtures/excel-functions2.json): statistics and
// distributions, maths, finance, engineering, text and arrays. Values are compared at about
// Excel's precision (1E-9 relative for floating results); differences we accept are listed in
// KNOWN with the reason. Also checks the German names (Formula2Local) and the _xlfn. prefixes.

import { describe, expect, it } from "vitest";
import { parseCell, Workbook, Sheet, err, type ErrorCode, type Value } from "../src";
import { evaluate, toFileFormula, toLocalFormula, XLFN } from "../src";
import fixture from "./fixtures/excel-functions2.json";

interface Enc { e: string }
type EV = number | string | boolean | null | Enc;
interface Case { group: string; file: string; local?: string; refused?: boolean; value?: EV; spill?: EV[][] }
interface Fixture {
  sheet: string;
  at: string;
  setup: { addr: string; value: EV; formula?: string }[];
  names: Record<string, string>;
  xlfn: Record<string, string>;
  cases: Case[];
}

const FIX = fixture as unknown as Fixture;

function dec(v: EV): Value {
  if (v !== null && typeof v === "object") return err(v.e as ErrorCode);
  return v;
}

function workbook(): Workbook {
  const wb = new Workbook();
  const sh = new Sheet(FIX.sheet);
  wb.sheets.push(sh);
  for (const s of FIX.setup) {
    const c = parseCell(s.addr)!;
    sh.set(c.row, c.col, { v: dec(s.value), ...(s.formula ? { f: s.formula } : {}) });
  }
  return wb;
}

/** A German complex number text ("1,5-2i", "-j", "1E-20i") as [re, im]; null when it is none. */
function complexOf(t: string): [number, number] | null {
  const num = "(?:\\d+(?:,\\d*)?(?:E[+-]\\d+)?)";
  const m = new RegExp(`^([+-]?${num})?(?:([+-]?)(${num})?([ij]))?$`).exec(t);
  if (!m || (!m[1] && !m[4])) return null;
  const f = (s: string) => Number(s.replace(",", "."));
  const re = m[1] ? f(m[1]) : 0;
  const im = m[4] ? (m[2] === "-" ? -1 : 1) * (m[3] ? f(m[3]) : 1) : 0;
  return [re, im];
}

function same(a: Value, b: Value): boolean {
  if (typeof a === "number" && typeof b === "number") {
    if (a === b) return true;
    return Math.abs(a - b) <= 1e-9 * Math.max(Math.abs(a), Math.abs(b));
  }
  // Complex results are texts with 15 digits: the last digit may differ (Excel's own sine and
  // cosine), so both parts are compared within 1E-9 of the number's modulus.
  if (typeof a === "string" && typeof b === "string" && a !== b) {
    const x = complexOf(a), y = complexOf(b);
    if (!x || !y || !/[ij]$/.test(a + b)) return false;
    const mod = Math.max(Math.hypot(...x), Math.hypot(...y));
    return Math.abs(x[0] - y[0]) <= 1e-9 * mod && Math.abs(x[1] - y[1]) <= 1e-9 * mod;
  }
  if (a !== null && b !== null && typeof a === "object" && typeof b === "object") return a.code === b.code;
  return a === b;
}

function show(v: unknown): string {
  return JSON.stringify(v, (_k, x) => (x && typeof x === "object" && "code" in x ? x.code : x));
}

/** Cases whose result we deliberately or knowingly compute differently, with the reason. */
const KNOWN: Record<string, string> = {
  "LINEST({1,2,3},{1,1,1},TRUE,TRUE)": "x is constant, so it duplicates the constant term: we drop it and fit the mean exactly; Excel reports an odd split of the sums of squares (1,866/0,134) and 1 degree of freedom",
  "IMABS(\"3.5\")": "German Excel reads 3.5 as the date 3 May of the current year (46145 in 2026); we do the same, but the value changes with the year",
  "IMAGINARY(\"-2.5i\")": "German Excel reads 2.5 as the date 2 May of the current year (-46144 in 2026); we do the same, but the value changes with the year",
  "ACCRINT(DATE(2007,1,10),DATE(2007,7,10),DATE(2009,3,15),0.05,1000,4,1)": "actual/actual over many quarters: Excel divides the last partial quarter (64 days) by 91 although it has 90 days; the other 12 recorded AUFGELZINS cases fit",
  "XIRR(F1:F5,G1:G5)": "Excel's iteration stops 2E-9 short of the root (with the guess 0,5 it ends 2E-10 from it, which we match); we return the root",
  "YIELD(DATE(2008,2,15),DATE(2016,11,15),0,95.04287,100,4)": "Excel's iteration stops 7E-9 (relative) from the root, where the price is already 95,04287 to 1E-9; we return the root",
  "BESSELJ(10,1)": "Excel's J1(10) is 5E-7 away from the true value (0,0434727462); its other recorded Bessel values, including J3(50) built on J1(50), fit the classic polynomial approximations we use",
};

/** Formulas whose German form our converter does not produce yet, with the reason. */
function knownConv(file: string): string | undefined {
  if (/^(GROUPBY|PIVOTBY)\(/.test(file)) {
    return "a function passed by name without brackets (SUMME → SUM) is a name to the lexer, which translates only function calls";
  }
  return undefined;
}

function mismatches(): { bad: string[]; byGroup: Record<string, { ok: number; bad: number; known: number }> } {
  const wb = workbook();
  const at = parseCell(FIX.at)!;
  const bad: string[] = [];
  const byGroup: Record<string, { ok: number; bad: number; known: number }> = {};
  for (const c of FIX.cases) {
    if (c.refused) continue;
    const g = (byGroup[c.group] ??= { ok: 0, bad: 0, known: 0 });
    if (KNOWN[c.file]) {
      g.known++;
      continue;
    }
    let got: Value | Value[][];
    try {
      got = evaluate(wb, 0, at.row, at.col, c.file);
    } catch (e) {
      bad.push(`${c.group} ${c.file}: threw ${(e as Error).message}`);
      g.bad++;
      continue;
    }
    let ok: boolean;
    let want: Value | Value[][];
    if (c.spill) {
      const w = c.spill.map((r) => r.map(dec));
      want = w;
      ok = Array.isArray(got) && got.length === w.length && got.every((r, i) => r.length === w[i].length && r.every((v, j) => same(v, w[i][j])));
    } else {
      want = dec(c.value ?? null);
      ok = !Array.isArray(got) && same(got, want);
    }
    if (ok) g.ok++;
    else {
      g.bad++;
      bad.push(`${c.group} ${c.file}: got ${show(got)} want ${show(want)}`);
    }
  }
  return { bad, byGroup };
}

describe("functions recorded from Excel (second round)", () => {
  const cases = FIX.cases.filter((c) => !c.refused);
  it(`computes the values of ${cases.length} cases`, () => {
    const { bad, byGroup } = mismatches();
    if (process.env.FORMULA_REPORT) {
      console.log(JSON.stringify(byGroup));
      console.log(bad.join("\n"));
    }
    expect(bad.length, bad.slice(0, 60).join("\n")).toBe(0);
  });

  it("converts between the German and the file form", () => {
    const wb = workbook();
    const bad: string[] = [];
    for (const c of cases) {
      if (!c.local || knownConv(c.file)) continue;
      const f = toFileFormula(c.local, wb);
      if (f !== c.file) bad.push(`toFile ${c.local}: got ${f} want ${c.file}`);
      const l = toLocalFormula(c.file, wb);
      if ("=" + l !== c.local) bad.push(`toLocal ${c.file}: got ${l} want ${c.local.slice(1)}`);
    }
    expect(bad.length, bad.slice(0, 40).join("\n")).toBe(0);
  });

  it("writes the _xlfn. prefixes Excel writes", () => {
    const bad = Object.entries(FIX.xlfn).filter(([k, v]) => XLFN[k] !== v).map(([k, v]) => `${k}: got ${XLFN[k]} want ${v}`);
    expect(bad, bad.join("\n")).toEqual([]);
  });
});
