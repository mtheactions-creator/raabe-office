// Spezialfilter against Excel (tools/feel-lab/probe_excel_advfilter.py, D-090): the rows Excel hides in place and the
// records it copies to another place, with and without Keine Duplikate.

import { describe, expect, it } from "vitest";
import { Calc, Sheet, Workbook, advancedFilter, parseInput, toFileFormula } from "../src";
import fixture from "./fixtures/excel-advfilter.json";

interface Case { name: string; action: string; criteria: string[][]; hidden?: number[]; output?: unknown[][]; error?: string }

function setup(crit: string[][]): { wb: Workbook; sh: Sheet; calc: Calc } {
  const wb = new Workbook();
  const sh = new Sheet("Tabelle1");
  wb.sheets.push(sh);
  (fixture.list as unknown[][]).forEach((r, i) => r.forEach((v, k) => { if (v !== "") sh.set(i, k, { v: v as string | number }); }));
  crit.forEach((r, i) => r.forEach((t, k) => {
    if (t === "") return;
    if (t.startsWith("=")) sh.set(i, 7 + k, { v: null, f: toFileFormula(t) });
    else sh.set(i, 7 + k, { v: parseInput(t).value ?? t });
  }));
  const calc = new Calc(wb);
  calc.recalc();
  return { wb, sh, calc };
}

describe("Spezialfilter as Excel filters", () => {
  for (const c of (fixture.cases as Case[]).filter((x) => !x.error)) {
    it(`${c.name} ${c.action}`, () => {
      const { wb, sh, calc } = setup(c.criteria);
      const list = { r0: 0, c0: 0, r1: fixture.list.length - 1, c1: 3 };
      const crit = { r0: 0, c0: 7, r1: c.criteria.length - 1, c1: 7 + c.criteria[0].length - 1 };
      const rows = advancedFilter(wb, sh, list, crit, c.action === "copy_unique", (f, r, col) => calc.evaluate(0, r, col, f));
      if (c.action === "inplace") {
        const hidden = [];
        for (let r = 1; r < fixture.list.length; r++) if (!rows.includes(r)) hidden.push(r + 1);
        expect(hidden).toEqual(c.hidden);
      } else {
        const out = [fixture.list[0], ...rows.map((r) => [0, 1, 2, 3].map((k) => sh.value(r, k) ?? null))];
        expect(out).toEqual((c.output as unknown[][]).map((r) => r.map((v) => (v === undefined ? null : v))));
      }
    });
  }
});
