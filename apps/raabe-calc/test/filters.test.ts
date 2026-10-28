// AutoFilter criteria against Excel (results/2026-10-28-excel-filter-sub.json, recorded with
// probe_excel_filter_sub.py): which rows Excel shows for each criterion, the same here.

import { readFileSync } from "node:fs";
import { describe, expect, it } from "vitest";
import type { Cell, FilterColumn } from "@raabe/sheet-engine";
import { passes, prepare } from "../src/filters";

const data = JSON.parse(readFileSync("tools/feel-lab/results/2026-10-28-excel-filter-sub.json", "utf-8")) as {
  rules: { col: number; c1: string | number; op: number | null; c2: string | null; shown: number[] }[];
};
const TEXTS = ["Deutsch", "Mathe", "Mathematik", "Physik", "Chemie", "Mathe-AG"];
const NUMS = [2, 3, 1, 5, 4, 2.5];
const OPS: Record<string, string> = { ">": "greaterThan", ">=": "greaterThanOrEqual", "<": "lessThan", "<=": "lessThanOrEqual", "<>": "notEqual", "=": "equal" };

/** Excel's criterion text (">=2,5", "Mathe*") as the file's customFilter. */
function custom(c: string): { op: string; val: string } {
  const m = /^(>=|<=|<>|>|<|=)?(.*)$/.exec(c)!;
  return { op: OPS[m[1] ?? "="], val: m[2].replace(",", ".") };
}

function criterion(r: (typeof data.rules)[number]): FilterColumn {
  if (r.op === 3 || r.op === 4 || r.op === 5) return { top10: { top: r.op !== 4, percent: r.op === 5, val: Number(r.c1) } };
  if (r.op === 11) return { dynamic: { type: r.c1 === 33 ? "aboveAverage" : "belowAverage" } };
  const items = [custom(String(r.c1)), ...(r.c2 ? [custom(r.c2)] : [])];
  return { custom: { and: r.op === 1 || undefined, items } };
}

describe("AutoFilter criteria as Excel applies them", () => {
  for (const r of data.rules) {
    it(`column ${r.col}: ${r.c1} ${r.op ?? ""} ${r.c2 ?? ""}`, () => {
      const cells: Cell[] = (r.col === 1 ? TEXTS : NUMS).map((v) => ({ v }));
      const crit = prepare(criterion(r), r.col === 2 ? NUMS : []);
      const shown = cells.map((c, i) => (passes(crit, c, String(c.v).replace(".", ",")) ? i + 2 : 0)).filter(Boolean);
      expect(shown).toEqual(r.shown);
    });
  }
});
