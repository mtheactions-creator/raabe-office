// Datenschnitt items as Excel shows them (D-092): every slicer of probe_excel_slicers.py's cases, its items
// in Excel's order with Selected and HasData as COM reports them, computed from the table's rows and the
// other slicers' selections.

import { readFileSync } from "node:fs";
import { describe, expect, it } from "vitest";
import { nextCacheName, nextSlicerName, slicerItems, type SlicerCrossFilter, type SlicerSort } from "../src";

const rec = JSON.parse(readFileSync("packages/ooxml/test/fixtures/excel-slicers.json", "utf-8"));
type Item = { name: string; selected: boolean; hasData: boolean };
type Rec = { sheet: string; slicers: { name: string; sourceName: string; items: Item[]; sortItems: number; crossFilter: number }[] };
// XlSlicerSort as COM reports it (the file of the "settings" case says descending where COM says 3).
const SORT: Record<number, SlicerSort> = { 1: "dataSourceOrder", 2: "ascending", 3: "descending" };
const CROSS: Record<number, SlicerCrossFilter> = { 1: "none", 2: "showItemsWithDataAtTop", 3: "showItemsWithNoData", 4: "showItemsWithDataAtTop" };

function rowsOf(caseName: string): (string | number | null)[][] {
  const data: (string | number | null)[][] = rec.data.map((r: unknown[]) => [...r]);
  if (caseName === "numbers") { data[3][3] = null; data.push(["Finn", "8c", "Mathe", 10]); }
  return data;
}

describe("Datenschnitt items as Excel lists them", () => {
  const cases = Object.entries(rec.cases as Record<string, Rec>).filter(([k]) => k !== "pivot");
  for (const [name, c] of cases) {
    it(name, () => {
      const data = rowsOf(name);
      const head = data[0] as string[];
      const rows = data.slice(1).map((r) => r.map((v) => ({ text: v === null ? "" : String(v), num: typeof v === "number" ? v : null })));
      const selections = c.slicers.map((s) => ({ col: head.indexOf(s.sourceName), values: s.items.filter((i) => i.selected).map((i) => i.name), all: s.items.every((i) => i.selected) }));
      c.slicers.forEach((s, k) => {
        const col = head.indexOf(s.sourceName);
        const others = selections.filter((x, j) => j !== k && !x.all);
        const fileSort = SORT[s.sortItems];
        const hide = s.crossFilter === 4;
        const got = slicerItems(rows, col, selections[k].all ? undefined : selections[k].values, others,
          { sort: fileSort, crossFilter: CROSS[s.crossFilter], hideNoData: hide });
        const want = hide ? s.items.filter((i) => i.hasData) : s.items;
        expect(got.map((i) => [i.text, i.selected, i.hasData]), s.name).toEqual(want.map((i) => [i.name, i.selected, i.hasData]));
      });
    });
  }
  it("names new caches and slicers as Excel does", () => {
    expect(nextCacheName([], "Fach")).toBe("Datenschnitt_Fach");
    expect(nextCacheName(["Datenschnitt_Fach"], "Fach")).toBe("Datenschnitt_Fach1");
    expect(nextCacheName(["Datenschnitt_Fach", "Datenschnitt_Fach1"], "Fach")).toBe("Datenschnitt_Fach2");
    expect(nextSlicerName([], "Fach")).toBe("Fach");
    expect(nextSlicerName(["Fach"], "Fach")).toBe("Fach 1");
    expect(nextSlicerName(["Fach", "Fach 1"], "Fach")).toBe("Fach 2");
  });
});
