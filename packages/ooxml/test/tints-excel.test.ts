// Excel's tints of design colours (tools/feel-lab/probe_excel_tints.py, D-083): 750 cells in five designs.

import { describe, expect, it } from "vitest";
import { tinted } from "../src/xlsx/colors";
import cases from "./fixtures/excel-tints.json";

describe("tints as Excel computes them", () => {
  it("750 theme colour cells", () => {
    const bad = (cases as { base: string; tint: number; rgb: string; design: string; theme: number }[]).filter((c) => tinted(c.base, c.tint) !== c.rgb);
    expect(bad.slice(0, 10).map((c) => `${c.design} ${c.theme} ${c.base} ${c.tint.toFixed(4)}: ${tinted(c.base, c.tint)} vs ${c.rgb}`)).toEqual([]);
  });
});
