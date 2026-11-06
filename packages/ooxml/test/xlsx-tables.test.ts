import { describe, expect, it } from "vitest";
import { Sheet, Workbook, type SheetTable } from "@raabe/sheet-engine";
import { readXlsx, writeXlsx, newWorkbook } from "../src";

describe("Excel tables (D-066)", () => {
  it("writes a table part and reads it back", () => {
    const wb = newWorkbook();
    const sh = wb.sheets[0];
    const vals = [["Name", "Note"], ["Anna", 2], ["Ben", 3]];
    vals.forEach((row, r) => row.forEach((v, c) => sh.set(r, c, { v })));
    const t: SheetTable = { name: "Noten", ref: { r0: 0, c0: 0, r1: 2, c1: 1 }, header: true, totals: false, columns: [{ name: "Name" }, { name: "Note" }],
      style: "TableStyleMedium9", firstColumn: true, lastColumn: false, rowStripes: true, columnStripes: false, filterButton: true };
    sh.tables = [t];
    const back = readXlsx(writeXlsx(wb));
    const bt = back.sheets[0].tables?.[0];
    expect(bt?.name).toBe("Noten");
    expect(bt?.style).toBe("TableStyleMedium9");
    expect(bt?.firstColumn).toBe(true);
    expect(bt?.columns.map((c) => c.name)).toEqual(["Name", "Note"]);
    expect(bt?.ref).toEqual(t.ref);
    // Unchanged it is written back as it was; removed, its part goes.
    const again = readXlsx(writeXlsx(back));
    expect(again.sheets[0].tables?.[0].id).toBe(bt?.id);
    again.sheets[0].tables = undefined;
    const none = readXlsx(writeXlsx(again));
    expect(none.sheets[0].tables).toBeUndefined();
    void Sheet; void Workbook;
  });
});
