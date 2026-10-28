// Edits of Excel-built workbooks, written back: the model must see the edit and keep the
// rest. With ROUNDTRIP_OUT set the edited files are written for xlsx_excel_check.py, which
// opens them in Excel and allows differences only where `expect` says.

import { readFileSync, mkdirSync, writeFileSync } from "node:fs";
import { describe, expect, it } from "vitest";
import { strFromU8, unzipSync } from "fflate";
import { Sheet, cellKey, parseRange, type Workbook } from "@raabe/sheet-engine";
import { newChart, type ChartModel } from "@raabe/charts";
import { newWorkbook, readXlsx, writeXlsx } from "../src";

const fixture = (name: string) => readXlsx(new Uint8Array(readFileSync(`packages/ooxml/test/fixtures/${name}`)));
const cell = (s: Sheet, a: string) => { const r = parseRange(a)!; return s.get(r.r0, r.c0); };
const set = (s: Sheet, a: string, c: Parameters<Sheet["set"]>[2]) => { const r = parseRange(a)!; s.set(r.r0, r.c0, c); };
const sheet = (wb: Workbook, name: string) => wb.sheetByName(name)!;
const charts = (s: Sheet) => (s as unknown as { charts?: ChartModel[] }).charts;

interface Expect {
  /** Paths of Excel's view (xlsx_excel_check.py) that may differ from the original. */
  allow: string[];
  /** Paths that must show these values in Excel. */
  require: [string, unknown][];
}

/** A size Excel shows depending on the screen's dpi. */
function approx(v: number, tol: number) {
  return { approx: v, tol };
}

/** Writes, reads back and hands the file to the Excel check. */
function save(wb: Workbook, name: string, exp: Expect): Workbook {
  const out = writeXlsx(wb, { author: "RaabeCalc-Test", now: new Date("2026-10-27T12:00:00Z") });
  if (process.env.ROUNDTRIP_OUT) {
    mkdirSync(process.env.ROUNDTRIP_OUT, { recursive: true });
    writeFileSync(`${process.env.ROUNDTRIP_OUT}/${name}.edit.xlsx`, out);
    writeFileSync(`${process.env.ROUNDTRIP_OUT}/${name}.edit.expect.json`, JSON.stringify(exp));
  }
  return readXlsx(out);
}

describe("edited workbooks", () => {
  it("values: changed, added, deleted and formatted cells", () => {
    const wb = fixture("xlsx-values.xlsx");
    const s = sheet(wb, "Werte");
    set(s, "A2", { v: 43 });
    const b2 = cell(s, "B2")!;
    b2.s = wb.styles.derive(b2.s, (st) => { st.font.bold = true; });
    s.cells.delete(cellKey(1, 2));
    set(s, "AB2", { v: "neu" });
    const back = save(wb, "xlsx-values", {
      allow: ["/sheets/Werte/cells/A2", "/sheets/Werte/cells/B2/Font", "/sheets/Werte/cells/C2", "/sheets/Werte/cells/AB2", "/sheets/Werte/UsedRange", "/sheets/Werte/cols"],
      require: [["/sheets/Werte/cells/A2/Value2", 43], ["/sheets/Werte/cells/AB2/Value2", "neu"], ["/sheets/Werte/cells/C2/Value2", null]],
    });
    const t = sheet(back, "Werte");
    expect(cell(t, "A2")?.v).toBe(43);
    expect(back.styles.get(cell(t, "B2")?.s).font.bold).toBe(true);
    expect(cell(t, "C2")).toBeUndefined();
    expect(cell(t, "B4")?.v).toBe("Zeile 1\nZeile 2");
  });

  it("styles: a theme colour replaced, a fill added", () => {
    const wb = fixture("xlsx-styles.xlsx");
    const s = sheet(wb, "Formate");
    const c = cell(s, "B14")!;
    c.s = wb.styles.derive(c.s, (st) => { st.font.color = "#FF0000"; });
    const d = cell(s, "B1")!;
    d.s = wb.styles.derive(d.s, (st) => { st.fill = "#92D050"; });
    const back = save(wb, "xlsx-styles", {
      allow: ["/sheets/Formate/cells/B14/Font", "/sheets/Formate/cells/B1/Interior"],
      require: [["/sheets/Formate/cells/B14/Font/8", 0x0000ff], ["/sheets/Formate/cells/B1/Interior/1", 0x50d092]],
    });
    expect(back.styles.get(cell(sheet(back, "Formate"), "B1")?.s)).toMatchObject({ fill: "#92D050", font: { bold: true } });
  });

  it("sheets: added, deleted, unhidden, recoloured, another active", () => {
    const wb = fixture("xlsx-sheets.xlsx");
    const neu = new Sheet("Neu");
    neu.set(0, 0, { v: 5 });
    neu.set(1, 0, { v: 10, f: "A1*2" });
    wb.sheets.push(neu);
    wb.sheets.splice(wb.sheets.findIndex((s) => s.name === "Akzent"), 1);
    sheet(wb, "Versteckt").hidden = undefined;
    sheet(wb, "Rot").tabColor = "#00B050";
    wb.active = 0;
    const back = save(wb, "xlsx-sheets", {
      allow: ["/Sheets", "/ActiveSheet", "/sheets/Neu", "/sheets/Akzent", "/sheets/Rot/Tab", "/sheets/Versteckt/Visible"],
      require: [["/sheets/Neu/cells/A2/Value2", 10], ["/sheets/Rot/Tab", 0x50b000], ["/ActiveSheet", "Erstes"], ["/sheets/Versteckt/Visible", -1]],
    });
    expect(back.sheets.map((s) => s.name)).toEqual(["Erstes", "Versteckt", "Sehr versteckt", "Rot", "Letztes", "Neu"]);
    expect(back.names.find((n) => n.name === "Lokal")?.sheet).toBe(3);
    expect(back.sheets[0].page?.printArea).toEqual({ r0: 0, c0: 0, r1: 3, c1: 1 });
  });

  it("layout: widths, heights, merges, freeze removed", () => {
    const wb = fixture("xlsx-layout.xlsx");
    const s = sheet(wb, "Layout");
    s.cols.set(0, { width: 30 });
    s.rows.set(1, { height: 40, custom: true });
    s.merges.push({ r0: 19, c0: 4, r1: 20, c1: 5 });
    s.view = { ...s.view, freeze: undefined, zoom: 100 };
    const back = save(wb, "xlsx-layout", {
      allow: ["/sheets/Layout/cols/0", "/sheets/Layout/rows/1", "/sheets/Layout/cells/E20", "/sheets/Layout/cells/E21", "/sheets/Layout/cells/F20", "/sheets/Layout/cells/F21", "/sheets/Layout/window"],
      require: [["/sheets/Layout/cols/0/0", approx(30, 0.2)], ["/sheets/Layout/rows/1/0", 40], ["/sheets/Layout/window/FreezePanes", false], ["/sheets/Layout/window/Zoom", 100]],
    });
    expect(sheet(back, "Layout").view.freeze).toBeUndefined();
    expect(sheet(back, "Layout").cols.get(0)?.width).toBeCloseTo(30, 2);
  });

  it("objects: notes, hyperlinks, filter, conditional format, page", () => {
    const wb = fixture("xlsx-objects.xlsx");
    const s = sheet(wb, "Objekte");
    cell(s, "B2")!.note = "Geänderte Notiz";
    delete cell(s, "C4")!.note;
    set(s, "A2", { ...cell(s, "A2")!, note: "Neue Notiz" });
    delete cell(s, "A9")!.link;
    set(s, "A12", { v: "Beispiel", link: "https://example.org/" });
    s.autoFilter = { ...s.autoFilter!, columns: {}, rawXml: undefined };
    for (const r of [...s.rows.keys()]) if (s.rows.get(r)?.hidden) s.rows.set(r, { ...s.rows.get(r)!, hidden: undefined });
    s.conditional![0].rules[0].format = { ...s.conditional![0].rules[0].format, fill: "#C6EFCE" };
    s.page = { ...s.page, orientation: "landscape" };
    const chart = charts(s)![0];
    chart.title = { show: true, text: "Notenschnitt 2026" };
    // The picture moves two columns to the right (its anchor is rewritten, the rest kept).
    const pic = s.pictures![0];
    pic.from = { ...pic.from, col: pic.from.col + 2 };
    if (pic.to) pic.to = { ...pic.to, col: pic.to.col + 2 };
    const back = save(wb, "xlsx-objects", {
      allow: ["/sheets/Objekte/notes", "/sheets/Objekte/links", "/sheets/Objekte/cells/A12", "/sheets/Objekte/cells/A9", "/sheets/Objekte/rows",
        "/sheets/Objekte/page", "/sheets/Objekte/UsedRange", "/sheets/Objekte/cols",
        // Unhidden rows move the chart, picture and note boxes down.
        "/sheets/Objekte/charts", "/sheets/Objekte/shapes"],
      require: [["/sheets/Objekte/page/3", 2], ["/sheets/Objekte/notes/0/0", "A2"], ["/sheets/Objekte/charts/0/7", "Notenschnitt 2026"], ["/sheets/Objekte/charts/0/1", 51]],
    });
    const t = sheet(back, "Objekte");
    expect(cell(t, "B2")?.note).toBe("Geänderte Notiz");
    expect(cell(t, "C4")?.note).toBeUndefined();
    expect(cell(t, "A2")?.note).toBe("Neue Notiz");
    expect(cell(t, "A12")?.link).toBe("https://example.org/");
    expect(cell(t, "A9")?.link).toBeUndefined();
    expect(t.pictures?.[0].from.col).toBe(pic.from.col);
    expect(t.autoFilter?.columns).toEqual({});
    expect(t.conditional?.[0].rules[0].format?.fill).toBe("#C6EFCE");
    expect(t.page?.orientation).toBe("landscape");
    expect(charts(t)?.[0].title.text).toBe("Notenschnitt 2026");
    // The picture, chart style and table parts are untouched; in the drawing only the moved
    // picture's corners differ (the chart was edited, not moved).
    const parts = unzipSync(writeXlsx(wb));
    const orig = unzipSync(new Uint8Array(readFileSync("packages/ooxml/test/fixtures/xlsx-objects.xlsx")));
    for (const p of ["xl/charts/style1.xml", "xl/tables/table1.xml", "xl/media/image1.png"]) expect(strFromU8(parts[p])).toBe(strFromU8(orig[p]));
    const cols = (x: string) => [...x.matchAll(/<xdr:col>(\d+)<\/xdr:col>/g)].map((m) => Number(m[1]));
    const before = strFromU8(orig["xl/drawings/drawing1.xml"]), after = strFromU8(parts["xl/drawings/drawing1.xml"]);
    expect(after.replace(/<xdr:col>\d+<\/xdr:col>/g, "")).toBe(before.replace(/<xdr:col>\d+<\/xdr:col>/g, ""));
    expect(cols(after).filter((c, i) => c !== cols(before)[i]).length).toBe(2);
  });

  it("formulas: a formula without a cached value makes Excel calculate on opening", () => {
    const wb = fixture("xlsx-formulas.xlsx");
    set(sheet(wb, "Formeln"), "A1", { v: null, f: "SUM(Daten!A1:A5)" });
    const back = save(wb, "xlsx-formulas", {
      allow: ["/sheets/Formeln/cells/A"],
      require: [["/sheets/Formeln/cells/A1/Value2", 15], ["/sheets/Formeln/cells/A1/Formula", "=SUM(Daten!A1:A5)"]],
    });
    expect(cell(sheet(back, "Formeln"), "A1")?.f).toBe("SUM(Daten!A1:A5)");
    const xml = strFromU8(unzipSync(writeXlsx(wb))["xl/workbook.xml"]);
    expect(xml).toMatch(/<calcPr [^>]*fullCalcOnLoad="1"/);
  });

  it("a new workbook with everything set from the model", () => {
    const wb = newWorkbook();
    const s = wb.sheets[0];
    s.set(0, 0, { v: "Titel", s: wb.styles.derive(0, (st) => { st.font.bold = true; st.font.size = 14; st.h = "center"; }) });
    s.merges.push({ r0: 0, c0: 0, r1: 0, c1: 2 });
    s.set(1, 0, { v: 1.5, s: wb.styles.derive(0, (st) => { st.numFmt = "0.00"; st.borders = { bottom: { style: "thin" } }; }) });
    s.set(2, 0, { v: 46322, s: wb.styles.derive(0, (st) => { st.numFmt = "m/d/yyyy"; }) });
    s.set(3, 0, { v: 3, f: "A2*2" });
    s.set(4, 0, { v: "Link", link: "https://www.raabeschule.de/", note: "Eine Notiz" });
    s.threads = [{ row: 0, col: 1, id: "{6E1A2B3C-0000-4000-8000-000000000001}", author: "Lehrkraft", date: "2026-10-28T03:40:00.00", text: "Bitte prüfen",
      replies: [{ id: "{6E1A2B3C-0000-4000-8000-000000000002}", author: "Anna", date: "2026-10-28T03:45:00.00", text: "Erledigt" }] }];
    s.cols.set(0, { width: 20 });
    s.rows.set(5, { height: 30, custom: true });
    s.rows.set(2, { outline: 1 });
    s.rows.set(3, { outline: 1 });
    s.view = { ...s.view, freeze: { row: 1, col: 0 }, top: 1 };
    s.autoFilter = { ref: { r0: 1, c0: 0, r1: 4, c1: 0 }, columns: { 0: { color: { fill: "#FFFF00" } } } };
    s.conditional = [{ ref: [{ r0: 1, c0: 0, r1: 3, c1: 0 }], rules: [{ type: "cellIs", priority: 1, operator: "greaterThan", formulas: ["2"], format: { color: "#9C0006", fill: "#FFC7CE" } }] },
      { ref: [{ r0: 1, c0: 3, r1: 3, c1: 3 }], rules: [{ type: "iconSet", priority: 2, iconSet: { set: "3Arrows", cfvo: [{ type: "percent", val: "0" }, { type: "percent", val: "33" }, { type: "percent", val: "67" }] } }] }];
    s.page = { orientation: "landscape", paperSize: 9, margins: { left: 0.7, right: 0.7, top: 0.75, bottom: 0.75, header: 0.3, footer: 0.3 }, printArea: { r0: 0, c0: 0, r1: 4, c1: 2 }, titleRows: [0, 1], titleCols: [0, 0], header: "&CKlasse 7b", footer: "&RSeite &P von &N", rowBreaks: [3] };
    s.tabColor = "#FF0000";
    s.pictures = [{ from: { col: 6, row: 1, x: 0, y: 0 }, to: { col: 8, row: 6, x: 0, y: 0 }, size: { w: 128, h: 100 }, src: "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==", name: "Logo" }];
    const zwei = new Sheet("Zweites");
    zwei.set(0, 0, { v: 1, f: "Tabelle1!A2-0.5" });
    wb.sheets.push(zwei);
    wb.names.push({ name: "Satz", ref: "Tabelle1!$A$2" });
    (s as unknown as { charts: ChartModel[] }).charts = [newChart("column", [{ name: { text: "Werte" }, val: { ref: "Tabelle1!$A$2:$A$4", cache: [1.5, 46322, 3] } }],
      { col: 4, row: 1, x: 0, y: 0 }, { col: 11, row: 15, x: 0, y: 0 })];
    const back = save(wb, "xlsx-new", {
      allow: [],
      require: [
        ["/sheets/Tabelle1/cells/A1/Value2", "Titel"], ["/sheets/Tabelle1/cells/A1/Merge", "A1:C1"], ["/sheets/Tabelle1/cells/A2/NumberFormat", "0.00"],
        ["/sheets/Tabelle1/cells/A3/NumberFormat", "m/d/yyyy"], ["/sheets/Tabelle1/cells/A4/Value2", 3], ["/sheets/Tabelle1/cols/0/0", approx(20, 0.2)],
        ["/sheets/Tabelle1/rows/5/0", 30], ["/sheets/Tabelle1/window/FreezePanes", true], ["/sheets/Tabelle1/autofilter/1", "A2:A5"],
        ["/sheets/Tabelle1/notes/0/1", "Eine Notiz"], ["/sheets/Tabelle1/threads/0/0", "B1"], ["/sheets/Tabelle1/threads/0/1", "Bitte prüfen"], ["/sheets/Tabelle1/threads/0/2", "Lehrkraft"], ["/sheets/Tabelle1/threads/0/3", 1], ["/sheets/Tabelle1/noteCount", 1], ["/sheets/Tabelle1/links/0/1", "https://www.raabeschule.de/"], ["/sheets/Tabelle1/page/3", 2],
        ["/sheets/Tabelle1/page/0", "$A$1:$C$5"], ["/sheets/Tabelle1/page/4", "$1:$2"], ["/sheets/Tabelle1/conditional/1/0", 6], ["/sheets/Tabelle1/shapes/1/0", "Logo"], ["/sheets/Tabelle1/shapes/1/1", 13], ["/sheets/Tabelle1/page/6", [4]], ["/sheets/Tabelle1/rows/2/2", 2], ["/sheets/Tabelle1/rows/3/2", 2], ["/sheets/Tabelle1/conditional/1/1", "D2:D4"], ["/sheets/Tabelle1/page/1", "Klasse 7b"], ["/sheets/Tabelle1/page/2", "Seite &P von &N"], ["/sheets/Tabelle1/page/5", "$A:$A"], ["/sheets/Tabelle1/Tab", 0x0000ff], ["/sheets/Zweites/cells/A1/Value2", 1], ["/sheets/Tabelle1/StandardHeight", approx(15, 0.6)],
        ["/sheets/Tabelle1/charts/0/1", 51], ["/sheets/Tabelle1/charts/0/6/0", "=SERIES(\"Werte\",,Tabelle1!$A$2:$A$4,1)"],
      ],
    });
    expect(back.sheets.map((x) => x.name)).toEqual(["Tabelle1", "Zweites"]);
    expect(back.sheets[0].conditional?.[0].rules[0]).toMatchObject({ type: "cellIs", operator: "greaterThan", formulas: ["2"], format: { color: "#9C0006", fill: "#FFC7CE" } });
    expect(back.names).toEqual([{ name: "Satz", ref: "Tabelle1!$A$2" }]);
    expect(back.sheets[0].autoFilter?.columns[0]).toEqual({ color: { fill: "#FFFF00" } });
    expect(back.sheets[0].threads?.[0]).toMatchObject({ row: 0, col: 1, author: "Lehrkraft", text: "Bitte prüfen", replies: [{ author: "Anna", text: "Erledigt" }] });
    expect(cell(back.sheets[0], "B1")?.note).toBeUndefined();
    expect(back.sheets[0].page).toMatchObject({ titleRows: [0, 1], titleCols: [0, 0], header: "&CKlasse 7b", footer: "&RSeite &P von &N", rowBreaks: [3] });
  });

  it("a copied sheet (Blatt verschieben/kopieren) gets parts of its own", () => {
    const wb = fixture("xlsx-objects.xlsx");
    const src = sheet(wb, "Objekte");
    const copy = Object.assign(new Sheet(""), structuredClone({ ...src }), { name: "Objekte (2)" });
    wb.sheets.push(copy);
    const back = save(wb, "xlsx-objects-kopie", {
      allow: [],
      require: [["/sheets/Objekte (2)/cells/B2/Value2", 2.3], ["/sheets/Objekte (2)/charts/0/1", 51], ["/sheets/Objekte (2)/notes/0/1", "Bitte prüfen"], ["/sheets/Objekte (2)/links/2/1", "https://www.raabeschule.de/"],
        ["/sheets/Objekte/notes/0/1", "Bitte prüfen"], ["/sheets/Objekte/charts/0/1", 51]],
    });
    const t = sheet(back, "Objekte (2)");
    expect(cell(t, "B2")?.note).toBe("Bitte prüfen");
    expect(cell(t, "A9")?.link).toBe("https://www.raabeschule.de/");
    expect(sheet(back, "Objekte").conditional?.length).toBe(t.conditional?.length);
    expect(charts(t)?.length).toBe(1);
  });

  it("a workbook made without the reader (Workbook.blank) is written too", () => {
    const back = readXlsx(writeXlsx(Object.assign(newWorkbook(), { raw: undefined })));
    expect(back.sheets.length).toBe(1);
  });
});
