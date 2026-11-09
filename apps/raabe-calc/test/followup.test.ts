// Lasting cases for what the follow-up of RaabeCalc's by-hand audit built (D-145 to D-148), each against what
// Excel 365 did in an Excel of our own (probe_excel_followup.py, results/2026-11-09-excel-followup.json).

import { describe, expect, it } from "vitest";
import { shiftFormula } from "@raabe/sheet-engine";

describe("sorting moves formulas as Excel does (D-145)", () => {
  // Recorded: each formula sorted from its row to the row Excel put it in.
  const cases: [string, number, string][] = [
    ["SUM(B2:B5)", -4, "SUM(B1:B1)"], // the sum row sorted to the top with its own range
    ["SUM(B2:B3)", -2, "SUM(B1:B1)"],
    ["B2+1", -2, "#REF!+1"], // a single reference off the sheet
    ["SUM(B1:B2)", -2, "SUM(#REF!)"], // both ends off the sheet
    ["SUM($B$2:$B$3)", -2, "SUM($B$2:$B$3)"],
    ["B5*2", 3, "B8*2"],
  ];
  for (const [f, d, want] of cases) it(`${f} moved ${d} rows`, () => expect(shiftFormula(f, d, 0, { clipRanges: true })).toBe(want));
  it("copying still makes #REF! of a range pushed off the sheet", () => expect(shiftFormula("SUM(B2:B5)", -4, 0)).toBe("SUM(#REF!)"));
});
