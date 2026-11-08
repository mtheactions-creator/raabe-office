// Zeitachse periods, labels and the date filter (D-121), against what Excel showed (probe_excel_timeline.py).

import { describe, expect, it } from "vitest";
import { addPeriods, dateToSerial, periodEnd, periodLabels, periodStart, periodsBetween, selectionLabel, timelineBounds, nextTimelineCacheName, pivotDateFiltersAfter, passesDateFilter, type PivotTable } from "../src";

const d = (y: number, m: number, day: number) => dateToSerial(y, m, day);

describe("periods", () => {
  it("start, end and step at every level", () => {
    expect(periodStart(d(2026, 5, 17), 0)).toBe(d(2026, 1, 1));
    expect(periodStart(d(2026, 5, 17), 1)).toBe(d(2026, 4, 1));
    expect(periodStart(d(2026, 5, 17), 2)).toBe(d(2026, 5, 1));
    expect(periodEnd(d(2026, 5, 17), 1)).toBe(d(2026, 6, 30));
    expect(periodEnd(d(2026, 2, 3), 2)).toBe(d(2026, 2, 28));
    expect(addPeriods(d(2025, 11, 20), 2, 3)).toBe(d(2026, 2, 1));
    expect(addPeriods(d(2025, 11, 20), 1, -1)).toBe(d(2025, 7, 1));
    expect(periodsBetween(d(2025, 8, 1), d(2026, 1, 31), 2)).toBe(5);
    expect(periodsBetween(d(2025, 8, 1), d(2026, 1, 31), 1)).toBe(2);
  });
  it("bounds as Excel saves them: 1 January of the first year to 1 January after the last", () => {
    expect(timelineBounds([d(2025, 8, 18), d(2026, 9, 13)])).toEqual({ start: d(2025, 1, 1), end: d(2027, 1, 1) });
  });
  it("labels as Excel draws them", () => {
    expect(periodLabels(d(2026, 3, 1), 2)).toEqual({ lower: "MRZ" });
    expect(periodLabels(d(2026, 1, 1), 2)).toEqual({ lower: "JAN", upper: "2026" });
    expect(periodLabels(d(2026, 1, 1), 1)).toEqual({ lower: "Q1", upper: "2026" });
    expect(periodLabels(d(2025, 9, 1), 3)).toEqual({ lower: "1", upper: "SEP 2025" });
  });
  it("selection labels as recorded", () => {
    expect(selectionLabel({ from: d(2025, 10, 1), to: d(2026, 1, 31) }, 2)).toBe("Okt 2025 - Jan 2026");
    expect(selectionLabel({ from: d(2026, 1, 1), to: d(2026, 6, 30) }, 1)).toBe("Q1 - Q2 2026");
    expect(selectionLabel({ from: d(2025, 9, 4), to: d(2025, 9, 21) }, 3)).toBe("4 - 21 Sep 2025");
    expect(selectionLabel({ from: d(2026, 1, 1), to: d(2026, 12, 31) }, 0)).toBe("2026");
    expect(selectionLabel(null, 2)).toBe("Alle Zeiträume");
  });
  it("cache names as Excel gives them", () => {
    expect(nextTimelineCacheName([], "Datum")).toBe("NativeZeitachse_Datum");
    expect(nextTimelineCacheName(["NativeZeitachse_Datum"], "Datum")).toBe("NativeZeitachse_Datum1");
  });
  it("the PivotTable's date filter: whole days, both ends included", () => {
    const p = { dateFilters: [{ field: 3, from: 1, to: 2 }] } as unknown as PivotTable;
    expect(pivotDateFiltersAfter(p, 0, { from: d(2026, 1, 1), to: d(2026, 1, 31) })).toEqual([{ field: 3, from: 1, to: 2 }, { field: 0, from: d(2026, 1, 1), to: d(2026, 1, 31) }]);
    expect(pivotDateFiltersAfter({ dateFilters: [{ field: 0, from: 1, to: 2 }] } as unknown as PivotTable, 0, null)).toBeUndefined();
    expect(passesDateFilter(d(2026, 1, 31) + 0.75, { from: d(2026, 1, 1), to: d(2026, 1, 31) })).toBe(true);
    expect(passesDateFilter(d(2026, 2, 1), { from: d(2026, 1, 1), to: d(2026, 1, 31) })).toBe(false);
  });
});
