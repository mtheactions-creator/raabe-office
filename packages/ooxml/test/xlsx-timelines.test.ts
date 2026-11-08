// Zeitachsen in .xlsx (D-121): Excel's own file (probe_excel_timeline.py: a default timeline, filters at the four
// levels, the Anzeigen options off with a style and caption, one timeline on two PivotTables, a cleared filter) read
// as Excel reports it through COM, written back unchanged, and edited (a new filter, a new timeline) in Excel's form.

import { readFileSync } from "node:fs";
import { describe, expect, it } from "vitest";
import { strFromU8, unzipSync } from "fflate";
import { computePivot, dateToSerial, pivotDateFiltersAfter, type Workbook } from "@raabe/sheet-engine";
import { readXlsx, writeXlsx } from "../src";

const FILE = "packages/ooxml/test/fixtures/excel-timeline.xlsx";
const rec = JSON.parse(readFileSync("packages/ooxml/test/fixtures/excel-timeline.json", "utf-8"));
const load = () => readXlsx(new Uint8Array(readFileSync(FILE)));
const write = (wb: Workbook) => writeXlsx(wb, { author: "RaabeCalc-Test", now: new Date("2026-11-09T12:00:00Z") });
const text = (f: Record<string, Uint8Array>, p: string) => (f[p] ? strFromU8(f[p]) : "");
const serial = (iso: string) => dateToSerial(Number(iso.slice(0, 4)), Number(iso.slice(5, 7)), Number(iso.slice(8, 10)));
type Case = { sheet: string; timelines: { name: string; caption: string; cache: string; style: string; state: { startDate: string | null; endDate: string | null }; view: { level: number; showHeader: boolean; showSelectionLabel: boolean; showTimeLevel: boolean; showScrollbar: boolean } }[]; reports: { values: (string | number | null)[][] }[] };

function report(wb: Workbook, sheet: string, k = 0): (string | number | null)[][] {
  const s = wb.sheetByName(sheet)!;
  const p = s.pivots![k];
  const src = wb.sheetByName(p.source.sheet)!;
  const data: (string | number | null)[][] = [];
  for (let r = p.source.r0; r <= p.source.r1; r++) {
    const row: (string | number | null)[] = [];
    for (let c = p.source.c0; c <= p.source.c1; c++) { const v = src.value(r, c); row.push(typeof v === "number" || typeof v === "string" ? v : null); }
    data.push(row);
  }
  const out = computePivot(p, data);
  const rows = new Map<number, (string | number | null)[]>();
  for (const c of out.cells) { const r = rows.get(c.row) ?? []; r[c.col - p.at.col] = c.v as string | number | null; rows.set(c.row, r); }
  return [...rows.keys()].sort((a, b) => a - b).map((r) => rows.get(r)!);
}

describe("Zeitachsen read from Excel's file", () => {
  const wb = load();
  it("has every timeline cache with its field and PivotTables", () => {
    const caches = wb.timelineCaches!;
    expect(caches.length).toBe(8);
    expect(caches.map((c) => c.name)).toContain("NativeZeitachse_Datum");
    const conn = caches.find((c) => c.pivots.length === 2)!;
    expect(conn.pivots).toEqual([{ sheet: "Verbunden", name: "PivotA" }, { sheet: "Verbunden", name: "PivotB" }]);
    expect(caches.every((c) => c.sourceName === "Datum" && (c.pivotCacheId ?? 0) > 0)).toBe(true);
    // The workbook names Excel keeps for them.
    expect(wb.names.filter((n) => n.name.startsWith("NativeZeitachse_")).length).toBe(8);
  });
  it("has every timeline as COM reports it: caption, cache, level, Anzeigen options, style", () => {
    let n = 0;
    for (const [, c] of Object.entries(rec.cases) as [string, Case][]) {
      const s = wb.sheetByName(c.sheet)!;
      for (const want of c.timelines) {
        const got = s.timelines!.find((x) => x.name === want.name)!;
        expect(got, want.name).toBeTruthy();
        expect(got.caption).toBe(want.caption);
        expect(got.cache).toBe(want.cache);
        expect(got.style ?? "TimeSlicerStyleLight1").toBe(want.style);
        expect(got.level).toBe(want.view.level);
        expect(got.showHeader !== false).toBe(want.view.showHeader);
        expect(got.showSelectionLabel !== false).toBe(want.view.showSelectionLabel);
        expect(got.showTimeLevel !== false).toBe(want.view.showTimeLevel);
        expect(got.showHorizontalScrollbar !== false).toBe(want.view.showScrollbar);
        n++;
      }
    }
    // Nine cases: "beforeClear" and "cleared" are one timeline.
    expect(n).toBe(9);
  });
  it("reads each PivotTable's date filter as the timeline's state, and the report shows Excel's values", () => {
    let n = 0;
    for (const [name, c] of Object.entries(rec.cases) as [string, Case][]) {
      if (name === "beforeClear") continue;
      const s = wb.sheetByName(c.sheet)!;
      const want = c.timelines[0].state;
      const df = s.pivots![0].dateFilters?.[0];
      if (want.startDate) expect(df, name).toEqual({ field: 0, from: serial(want.startDate), to: serial(want.endDate!) });
      else expect(df, name).toBeUndefined();
      c.reports.forEach((r, k) => expect(report(wb, c.sheet, k), `${name} ${k}`).toEqual(r.values));
      n++;
    }
    expect(n).toBe(8);
  });
  it("resolves a PivotTable source given by a table's name", () => {
    const p = wb.sheetByName("Filter")!.pivots![0];
    expect(p.sourceTable).toBe("Noten");
    expect(p.source).toEqual({ sheet: "Noten", r0: 0, c0: 0, r1: 24, c1: 2 });
  });
});

describe("Zeitachsen written", () => {
  it("keeps every timeline part byte for byte while nothing changed", () => {
    const before = unzipSync(new Uint8Array(readFileSync(FILE)));
    const after = unzipSync(write(load()));
    for (const p of Object.keys(before).filter((x) => /timeline|drawings\/drawing|pivotTables\/pivotTable|pivotCache\/pivotCacheDefinition/.test(x))) {
      expect(text(after, p), p).toBe(text(before, p));
    }
  });
  it("writes a changed filter in Excel's form: the PivotTable's dateBetween filter, the cache's selection and bounds", () => {
    const wb = load();
    const s = wb.sheetByName("Standard")!;
    const p = s.pivots![0];
    p.dateFilters = pivotDateFiltersAfter(p, 0, { from: dateToSerial(2026, 2, 1), to: dateToSerial(2026, 4, 30) });
    const f = unzipSync(write(wb));
    const tables = Object.keys(f).filter((x) => /^xl\/pivotTables\/pivotTable\d+\.xml$/.test(x)).map((x) => text(f, x));
    const t = tables.find((x) => x.includes('name="PivotStandard"'))!;
    expect(t).toContain('<filter fld="0" type="dateBetween"');
    expect(t).toContain(`<customFilter operator="greaterThanOrEqual" val="${dateToSerial(2026, 2, 1)}"/><customFilter operator="lessThanOrEqual" val="${dateToSerial(2026, 4, 30)}"/>`);
    const caches = Object.keys(f).filter((x) => x.startsWith("xl/timelineCaches/")).map((x) => text(f, x));
    const c = caches.find((x) => x.includes('name="NativeZeitachse_Datum"'))!;
    expect(c).toContain('filterType="dateBetween"><selection startDate="2026-02-01T00:00:00" endDate="2026-04-30T00:00:00"/><bounds startDate="2025-01-01T00:00:00" endDate="2027-01-01T00:00:00"/>');
    // The x14 id the timeline names is the one in the PivotTable's (new) cache.
    const id = /pivotCacheId="(\d+)"/.exec(c)![1];
    const cacheParts = Object.keys(f).filter((x) => /pivotCacheDefinition\d+\.xml$/.test(x)).map((x) => text(f, x));
    expect(cacheParts.some((x) => x.includes(`pivotCacheId="${id}"`) && x.includes('containsDate="1"'))).toBe(true);
    expect(text(f, "xl/workbook.xml")).toContain("<x15:timelineCacheRefs>");
    // Read again: the same filter and timelines.
    const again = readXlsx(write(wb));
    expect(again.sheetByName("Standard")!.pivots![0].dateFilters).toEqual([{ field: 0, from: dateToSerial(2026, 2, 1), to: dateToSerial(2026, 4, 30) }]);
    expect(again.timelineCaches!.length).toBe(8);
    expect(again.sheets.flatMap((x) => x.timelines ?? []).length).toBe(8);
  });
  it("writes a new timeline with its anchor, part, extLst entries and name", () => {
    const wb = load();
    const s = wb.sheetByName("Standard")!;
    wb.timelineCaches!.push({ name: "NativeZeitachse_Datum8", sourceName: "Datum", pivots: [{ sheet: "Standard", name: "PivotStandard" }] });
    wb.names.push({ name: "NativeZeitachse_Datum8", ref: "#N/A" });
    s.timelines!.push({ name: "Datum 7", cache: "NativeZeitachse_Datum8", caption: "Datum", level: 1, showTimeLevel: false, from: { col: 8, row: 2, x: 0, y: 0 }, to: { col: 13, row: 10, x: 0, y: 0 } });
    const f = unzipSync(write(wb));
    const tl = Object.keys(f).filter((x) => x.startsWith("xl/timelines/")).map((x) => text(f, x)).find((x) => x.includes('name="Datum 7"'))!;
    expect(tl).toMatch(/<timeline name="Datum 7" [^>]*cache="NativeZeitachse_Datum8" caption="Datum" showTimeLevel="0" level="1" selectionLevel="1"\/>/);
    const dr = Object.keys(f).filter((x) => x.startsWith("xl/drawings/") && x.endsWith(".xml")).map((x) => text(f, x)).find((x) => x.includes('name="Datum 7"'))!;
    expect(dr).toContain('<tsle:timeslicer xmlns:tsle="http://schemas.microsoft.com/office/drawing/2012/timeslicer" name="Datum 7"/>');
    expect(text(f, "[Content_Types].xml")).toContain("application/vnd.ms-excel.timeline+xml");
    const again = readXlsx(write(wb));
    const t = again.sheetByName("Standard")!.timelines!.find((x) => x.name === "Datum 7")!;
    expect(t).toMatchObject({ level: 1, showTimeLevel: false, from: { col: 8, row: 2 } });
  });
});
