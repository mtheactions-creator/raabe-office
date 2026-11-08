// Datenschnitte in .xlsx (D-092): Excel's own file (probe_excel_slicers.py: table slicers with defaults,
// filtering, the Elemente-ohne-Daten settings, Beschriftung/Spalten/Höhe/Breite/Formatvorlage/Sortierung,
// numbers with an empty cell, a PivotTable slicer) read as Excel reports it through COM, written back
// unchanged, and edited (a new slicer, a changed one, one removed) in Excel's form.

import { readFileSync } from "node:fs";
import { describe, expect, it } from "vitest";
import { strFromU8, unzipSync } from "fflate";
import { type Workbook } from "@raabe/sheet-engine";
import { readXlsx, writeXlsx } from "../src";

const FILE = "packages/ooxml/test/fixtures/excel-slicers.xlsx";
const rec = JSON.parse(readFileSync("packages/ooxml/test/fixtures/excel-slicers.json", "utf-8"));
const load = () => readXlsx(new Uint8Array(readFileSync(FILE)));
const write = (wb: Workbook) => writeXlsx(wb, { author: "RaabeCalc-Test", now: new Date("2026-11-08T12:00:00Z") });
const files = (b: Uint8Array) => unzipSync(b);
const text = (f: Record<string, Uint8Array>, p: string) => (f[p] ? strFromU8(f[p]) : "");
const colOf = (a: string) => [...a.replace(/\d+/g, "")].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0) - 1;
const rowOf = (a: string) => Number(a.replace(/\D+/g, "")) - 1;

describe("Datenschnitte read from Excel's file", () => {
  const wb = load();
  it("has every slicer cache with its source, table column or PivotTable and settings", () => {
    const caches = wb.slicerCaches!;
    expect(caches.length).toBe(18);
    const byName = (n: string) => caches.find((c) => c.name === n)!;
    expect(byName("Datenschnitt_Fach")).toMatchObject({ sourceName: "Fach", table: { name: "Noten", column: 2 } });
    expect(byName("Datenschnitt_Klasse")).toMatchObject({ table: { name: "Noten", column: 1 } });
    expect(byName("Datenschnitt_Note").table).toEqual({ name: "Noten6", column: 3 });
    expect(byName("Datenschnitt_Name1").crossFilter).toBe("none");
    expect(byName("Datenschnitt_Name2").crossFilter).toBeUndefined();
    expect(byName("Datenschnitt_Name3").crossFilter).toBe("showItemsWithNoData");
    expect(byName("Datenschnitt_Name4").hideNoData).toBe(true);
    expect(byName("Datenschnitt_Klasse2").sort).toBe("descending");
    const pivot = byName("Datenschnitt_Fach8");
    expect(pivot.pivots).toEqual([{ sheet: "Pivot", name: "PivotNoten" }]);
    expect(pivot.pivotCacheId).toBeGreaterThan(0);
  });
  it("has every slicer as COM reports it: name, caption, cache, header, columns, button height, style, place", () => {
    let n = 0;
    for (const [, c] of Object.entries(rec.cases) as [string, { sheet: string; slicers: { name: string; caption: string; cache: string; columns: number; rowHeight: number; style: string; displayHeader: boolean; shape: { tl: string; br: string } }[] }][]) {
      const s = wb.sheetByName(c.sheet)!;
      for (const want of c.slicers) {
        const got = s.slicers!.find((x) => x.name === want.name)!;
        expect(got, want.name).toBeTruthy();
        expect(got.caption).toBe(want.caption);
        expect(got.cache).toBe(want.cache);
        expect(got.columns ?? 1).toBe(want.columns);
        // COM gives the button height in pt; the file in EMU (12700 per pt).
        expect(Math.abs((got.rowHeight ?? 251883) / 12700 - want.rowHeight)).toBeLessThan(0.01);
        expect(got.style ?? "SlicerStyleLight1").toBe(want.style);
        expect(got.showCaption !== false).toBe(want.displayHeader);
        // TopLeftCell / BottomRightCell are the cells the corners lie in.
        expect([got.from.col, got.from.row]).toEqual([colOf(want.shape.tl), rowOf(want.shape.tl)]);
        n++;
      }
    }
    expect(n).toBeGreaterThan(15);
  });
  it("keeps a table slicer's selection in the table's AutoFilter, as the rows Excel hid show", () => {
    const t = wb.sheetByName("Filter")!.tables![0];
    expect(t.filters).toEqual({ 2: ["Deutsch", "Englisch"] });
    const z = wb.sheetByName("Zwei")!.tables![0];
    // After Filter löschen on Fach (Excel's last state there) the table has no filter left.
    expect(z.filters).toBeUndefined();
    const num = wb.sheetByName("Zahlen")!.tables![0];
    expect(num.filters![3]).toContain("");
    expect(num.filters![3]).not.toContain("1");
  });
  it("reads no slicer frame as a shape", () => {
    for (const s of wb.sheets) expect(s.shapes ?? []).toEqual([]);
  });
});

describe("Datenschnitte written", () => {
  it("unchanged: every slicer part, the drawings and the extLst entries stay byte for byte", () => {
    const orig = files(new Uint8Array(readFileSync(FILE)));
    const out = files(write(load()));
    for (const p of Object.keys(orig).filter((x) => /slicer|drawings\/drawing/.test(x))) expect(text(out, p), p).toBe(text(orig, p));
    const ext = (x: string) => /<extLst>.*<\/extLst>/s.exec(x)?.[0];
    expect(ext(text(out, "xl/workbook.xml"))).toBe(ext(text(orig, "xl/workbook.xml")));
    expect(ext(text(out, "xl/worksheets/sheet2.xml"))).toBe(ext(text(orig, "xl/worksheets/sheet2.xml")));
  });
  it("after a change: caches, slicers, anchors, names and extLst in Excel's form; read back the same", () => {
    const wb = load();
    const s = wb.sheetByName("Filter")!;
    // Changed: the Fach slicer gets two columns and a style; Name is removed; a Klasse slicer is added.
    const fach = s.slicers!.find((x) => x.name === "Fach 1")!;
    fach.columns = 2;
    fach.style = "SlicerStyleDark2";
    fach.from = { ...fach.from, row: 3 };
    s.slicers = s.slicers!.filter((x) => x.name !== "Name");
    wb.slicerCaches = wb.slicerCaches!.filter((c) => c.name !== "Datenschnitt_Name");
    wb.names = wb.names.filter((n) => n.name !== "Datenschnitt_Name");
    wb.slicerCaches.push({ name: "Datenschnitt_Klasse9", sourceName: "Klasse", table: { name: "Noten2", column: 1 }, crossFilter: "showItemsWithNoData", sort: "descending" });
    wb.names.push({ name: "Datenschnitt_Klasse9", ref: "#N/A" });
    s.slicers.push({ name: "Klasse 9", cache: "Datenschnitt_Klasse9", caption: "Klasse", from: { col: 9, row: 1, x: 0, y: 0 }, to: { col: 12, row: 15, x: 0, y: 0 } });
    const f = files(write(wb));
    const book = text(f, "xl/workbook.xml");
    expect(book).toContain('<ext uri="{46BE6895-7355-4a93-B00E-2C351335B9C9}"');
    expect(book).toContain('<ext uri="{BBE1A952-AA13-448e-AADC-164F8A28A991}"');
    expect(book).not.toContain('"Datenschnitt_Name"');
    const caches = Object.keys(f).filter((p) => p.startsWith("xl/slicerCaches/") && book.includes("slicerCache"));
    const all = caches.map((p) => text(f, p)).join("");
    expect(all).toContain('name="Datenschnitt_Klasse9"');
    expect(all).toMatch(/<x15:tableSlicerCache tableId="2" column="2" sortOrder="descending" crossFilter="showItemsWithNoData"\/>/);
    expect(all).toMatch(/<tabular pivotCacheId="\d+"><items count="3"><i x="1" s="1"\/><i x="2" s="1"\/><i x="0"\/><\/items><\/tabular>/);
    const sheetRels = text(f, "xl/worksheets/_rels/sheet2.xml.rels");
    const sp = /Target="\/?(?:\.\.\/)?(xl\/)?(slicers\/slicer\d+\.xml)"/.exec(sheetRels)!;
    const slicerPart = text(f, `xl/${sp[2]}`);
    expect(slicerPart).toMatch(/<slicer name="Fach 1" xr10:uid="\{[0-9A-F-]+\}" cache="Datenschnitt_Fach1" caption="Fach" columnCount="2" style="SlicerStyleDark2" rowHeight="251883"\/>/);
    expect(slicerPart).toContain('name="Klasse 9"');
    expect(slicerPart).not.toContain('name="Name"');
    const drawing = text(f, "xl/drawings/drawing2.xml");
    expect(drawing).toContain('<sle:slicer xmlns:sle="http://schemas.microsoft.com/office/drawing/2010/slicer" name="Klasse 9"/>');
    expect(drawing).not.toContain('name="Name"');
    expect(drawing).toMatch(/<xdr:from><xdr:col>5<\/xdr:col><xdr:colOff>0<\/xdr:colOff><xdr:row>3<\/xdr:row>/);
    // Read back: the same model.
    const back = readXlsx(write(wb));
    const bs = back.sheetByName("Filter")!;
    expect(bs.slicers!.map((x) => [x.name, x.columns ?? 1, x.style ?? "", x.from.row])).toEqual([["Fach 1", 2, "SlicerStyleDark2", 3], ["Klasse 9", 1, "", 1]]);
    expect(back.slicerCaches!.find((c) => c.name === "Datenschnitt_Klasse9")).toMatchObject({ table: { name: "Noten2", column: 1 }, crossFilter: "showItemsWithNoData", sort: "descending" });
    expect(back.slicerCaches!.length).toBe(18);
    expect(back.sheetByName("Pivot")!.slicers!.length).toBe(1);
  });
  it("a table's chosen values are written as Excel writes them", () => {
    const wb = load();
    const t = wb.sheetByName("Standard")!.tables![0];
    t.filters = { 2: ["Mathe"], 3: ["", "2"] };
    const f = files(write(wb));
    const x = text(f, "xl/tables/table1.xml");
    expect(x).toContain('<filterColumn colId="2"><filters><filter val="Mathe"/></filters></filterColumn><filterColumn colId="3"><filters blank="1"><filter val="2"/></filters></filterColumn>');
  });
});
