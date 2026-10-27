// RaabeCalc's templates: they calculate, and they survive saving as .xlsx.

import { describe, expect, it } from "vitest";
import { Calc } from "@raabe/sheet-engine";
import { readXlsx, writeXlsx } from "@raabe/ooxml";
import { CALC_TEMPLATES, templateWorkbook } from "../src/templates";

describe("templates", () => {
  for (const t of CALC_TEMPLATES) {
    it(`${t.name} calculates and round-trips`, () => {
      const wb = templateWorkbook(t.id);
      new Calc(wb).recalc();
      const back = readXlsx(writeXlsx(wb));
      expect(back.sheet().name).toBe(wb.sheet().name);
      expect(back.sheet().charts?.length ?? 0).toBe(wb.sheet().charts?.length ?? 0);
    });
  }

  it("Messwerte fits the sample data", () => {
    const wb = templateWorkbook("messwerte");
    new Calc(wb).recalc();
    expect(wb.sheet().get(3, 4)?.v).toBeCloseTo(2.0086, 3);
    expect(wb.sheet().get(5, 4)?.v).toBeCloseTo(0.999, 3);
  });
});
