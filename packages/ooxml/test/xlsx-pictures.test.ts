import { readFileSync } from "node:fs";
import { describe, expect, it } from "vitest";
import { readXlsx } from "../src";

describe("pictures in a sheet", () => {
  it("are read with their anchor and image (Excel's fixture xlsx-objects.xlsx)", () => {
    const wb = readXlsx(readFileSync("packages/ooxml/test/fixtures/xlsx-objects.xlsx"));
    const pics = wb.sheets.flatMap((s) => s.pictures ?? []);
    expect(pics.length).toBe(1);
    expect(pics[0].src.startsWith("data:image/png;base64,")).toBe(true);
    expect(pics[0].from.col).toBeGreaterThanOrEqual(0);
    expect(pics[0].to ?? pics[0].size).toBeTruthy();
  });
});
