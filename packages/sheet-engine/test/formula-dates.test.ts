// The formula engine's date functions on the edge cases recorded for dates.ts
// (fixtures/excel-input.json, tools/feel-lab/probe_excel_input.py): years 0-1899, month and day
// overflow both ways, 29 Feb 1900, 31.12.9999. 1900 date system only (see report on 1904).

import { readFileSync } from "node:fs";
import { describe, expect, it } from "vitest";
import { CellError, Workbook, evaluate } from "../src";

interface Case { fn: string; args: number[]; value: number | string | null }
const fx = JSON.parse(readFileSync("packages/sheet-engine/test/fixtures/excel-input.json", "utf-8")) as { dateFunctions: Case[] };

describe("date functions against Excel", () => {
  it(`computes ${fx.dateFunctions.length} recorded cases`, () => {
    const wb = Workbook.blank();
    const misses: string[] = [];
    for (const c of fx.dateFunctions) {
      const f = `${c.fn}(${c.args.join(",")})`;
      const v = evaluate(wb, 0, 0, 0, f);
      const got = v instanceof CellError ? "error" : v;
      const want = typeof c.value === "string" ? "error" : c.value;
      if (got !== want) misses.push(`${f}: Excel ${c.value}, ours ${String(got)}`);
    }
    expect(misses, misses.slice(0, 30).join("\n")).toEqual([]);
  });
});
