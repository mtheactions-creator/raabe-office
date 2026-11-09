// Datenbalken with Einfarbige Füllung (D-146): Excel 2010's part of the rule in the sheet's extLst, written as
// Excel's gallery writes it (recorded, results/2026-11-09-excel-followup-ui.json), read back, kept and removed.

import { describe, expect, it } from "vitest";
import { strFromU8, unzipSync } from "fflate";
import { parseRange } from "@raabe/sheet-engine";
import { newWorkbook, readXlsx, writeXlsx } from "../src";

const sheetXml = (bytes: Uint8Array) => strFromU8(unzipSync(bytes)["xl/worksheets/sheet1.xml"]);

describe("data bars with Einfarbige Füllung", () => {
  it("writes the x14 part with gradient=0, reads it back as solid, and keeps it on saving again", () => {
    const wb = newWorkbook();
    const s = wb.sheets[0];
    for (let i = 0; i < 5; i++) s.set(i, 1, { v: i + 1 });
    s.conditional = [{ ref: [parseRange("B1:B5")!], rules: [{ type: "dataBar", priority: 1, colors: ["#638EC6"], solid: true }] }];
    const xml = sheetXml(writeXlsx(wb));
    const id = /<x14:id>([^<]+)<\/x14:id>/.exec(xml)?.[1];
    expect(id).toBeTruthy();
    expect(xml).toContain(`<x14:cfRule type="dataBar" id="${id}"><x14:dataBar minLength="0" maxLength="100" gradient="0">`);
    expect(xml).toContain("<xm:sqref>B1:B5</xm:sqref>");
    // The extLst is the worksheet's last child.
    expect(xml.trim().endsWith("</extLst></worksheet>")).toBe(true);
    const back = readXlsx(writeXlsx(wb));
    const rule = back.sheets[0].conditional![0].rules[0];
    expect(rule.solid).toBe(true);
    expect(rule.colors).toEqual(["#638EC6"]);
    // Saved again with a change elsewhere: the rule and its x14 part stay, once.
    back.sheets[0].set(7, 0, { v: "x" });
    const again = sheetXml(writeXlsx(back));
    expect(again.match(/<x14:cfRule /g)?.length).toBe(1);
    expect(readXlsx(writeXlsx(back)).sheets[0].conditional![0].rules[0].solid).toBe(true);
  });

  it("a gradient bar has no x14 part, and removing a solid bar removes its x14 part", () => {
    const wb = newWorkbook();
    const s = wb.sheets[0];
    s.conditional = [{ ref: [parseRange("B1:B5")!], rules: [{ type: "dataBar", priority: 1, colors: ["#638EC6"] }] }];
    expect(sheetXml(writeXlsx(wb))).not.toContain("x14:");
    s.conditional = [{ ref: [parseRange("B1:B5")!], rules: [{ type: "dataBar", priority: 1, colors: ["#FF555A"], solid: true }] }];
    const back = readXlsx(writeXlsx(wb));
    back.sheets[0].conditional = undefined;
    const xml = sheetXml(writeXlsx(back));
    expect(xml).not.toContain("x14:cfRule");
    expect(xml).not.toContain("<extLst>");
  });
});
