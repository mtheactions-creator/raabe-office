// Windows, views and ink in .xlsx (D-093), against Excel's own files (probe_excel_windows.py,
// probe_excel_ink.py; results/2026-11-08-excel-windows.json and -excel-ink.json): Teilen as the
// sheetView's pane, Neues Fenster as further workbookViews and sheetViews, Ausblenden, Seitenlayout,
// Benutzerdefinierte Ansichten and Excel's ink (contentPart with InkML) read, written back unchanged and
// written from the model in Excel's form.

import { readFileSync } from "node:fs";
import { describe, expect, it } from "vitest";
import { strFromU8, unzipSync } from "fflate";
import { Sheet, Workbook, type View } from "@raabe/sheet-engine";
import { readXlsx, writeXlsx } from "../src";
import { decodeTrace, encodeTrace, parseInkml } from "../src/xlsx/ink";

const FIX = "packages/ooxml/test/fixtures/";
const load = (f: string) => readXlsx(new Uint8Array(readFileSync(FIX + f)));
const rec = JSON.parse(readFileSync(FIX + "excel-windows.json", "utf-8"));
const inkRec = JSON.parse(readFileSync(FIX + "excel-ink.json", "utf-8"));
const write = (wb: Workbook) => writeXlsx(wb, { author: "RaabeCalc-Test", now: new Date("2026-11-08T12:00:00Z") });
const files = (b: Uint8Array) => unzipSync(b);
const part = (b: Uint8Array, n: string) => strFromU8(files(b)[n]);
const sheetViews = (x: string) => /<sheetViews>[\s\S]*?<\/sheetViews>/.exec(x)?.[0] ?? null;
/** Excel's sheetView without what does not belong to the split (tab selection, the zoom it noted). */
const bare = (x: string) => x.replace(/ tabSelected="1"/, "").replace(/ zoomScaleNormal="\d+"/, "").replace(/ zoomScale="\d+"/, "");

describe("Teilen as Excel stores it", () => {
  const wb = load("excel-split.xlsx");
  it("reads every recorded split: bars in 1/20 pt, the panes' first cells, the active pane", () => {
    rec.splits.forEach((c: { name: string; com: { splitRow: number; splitColumn: number; panes: { scrollRow: number; scrollColumn: number }[]; activePane: number }; xml: { sheetViews: string } }, i: number) => {
      const v = wb.sheets[i].view;
      const pane = /<pane\b([^>]*)\/>/.exec(c.xml.sheetViews)![1];
      const xs = Number(/xSplit="(\d+)"/.exec(pane)?.[1] ?? 0), ys = Number(/ySplit="(\d+)"/.exec(pane)?.[1] ?? 0);
      expect(v.split, c.name).toBeTruthy();
      expect([v.split!.x, v.split!.y], c.name).toEqual([xs, ys]);
      // COM's Panes: 1 = top-left; the last one is the bottom-right (or the second of two).
      const p = c.com.panes;
      expect([v.split!.top + 1, v.split!.left + 1], c.name).toEqual([p[0].scrollRow, p[0].scrollColumn]);
      expect([v.top + 1, v.left + 1], c.name).toEqual([p[p.length - 1].scrollRow, p[p.length - 1].scrollColumn]);
    });
  });
  it("writes the file back unchanged", () => {
    const out = write(load("excel-split.xlsx"));
    rec.splits.forEach((c: { xml: { sheetViews: string } }, i: number) => {
      expect(sheetViews(part(out, `xl/worksheets/sheet${i + 1}.xml`))).toBe(c.xml.sheetViews);
    });
  });
  it("writes each split from the model as Excel writes it", () => {
    rec.splits.forEach((c: { name: string; xml: { sheetViews: string } }, i: number) => {
      const fresh = Workbook.blank();
      const v: View = JSON.parse(JSON.stringify(wb.sheets[i].view));
      fresh.sheets[0].view = v;
      const x = sheetViews(part(write(fresh), "xl/worksheets/sheet1.xml"))!;
      expect(bare(x), c.name).toBe(bare(c.xml.sheetViews));
    });
  });
  it("removes the split when the view has none", () => {
    const w = load("excel-split.xlsx");
    delete w.sheets[2].view.split;
    w.sheets[2].view.top = 0;
    w.sheets[2].view.left = 0;
    expect(sheetViews(part(write(w), "xl/worksheets/sheet3.xml"))).not.toContain("<pane");
  });
});

describe("Neues Fenster, Ausblenden as Excel stores them", () => {
  it("reads two windows with their own sheet, cell, zoom and scroll", () => {
    const wb = load("excel-windows.xlsx");
    expect(wb.windows?.length).toBe(2);
    expect(wb.active).toBe(0);
    expect(wb.windows![1].active).toBe(1);
    const s2 = wb.sheets[1];
    expect(s2.views?.length).toBe(1);
    expect(s2.views![0].zoom).toBe(150);
    expect(s2.views![0].top).toBe(9);
    expect(s2.views![0].active).toEqual({ row: 2, col: 2 });
    expect(wb.sheets[0].view.active).toEqual({ row: 1, col: 1 });
  });
  it("reads a hidden window", () => {
    const wb = load("excel-windows-hidden.xlsx");
    expect(wb.windows!.map((w) => !!w.hidden)).toEqual([false, true]);
  });
  it("writes both windows back unchanged", () => {
    const out = write(load("excel-windows.xlsx"));
    const x = rec.windows.xml_two;
    expect(/<bookViews>[\s\S]*?<\/bookViews>/.exec(part(out, "xl/workbook.xml"))?.[0]).toBe(x.bookViews);
    expect(sheetViews(part(out, "xl/worksheets/sheet1.xml"))).toBe(x.sheets.sheet1.sheetViews);
    expect(sheetViews(part(out, "xl/worksheets/sheet2.xml"))).toBe(x.sheets.sheet2.sheetViews);
  });
  it("writes a new window and Ausblenden from the model in Excel's form", () => {
    const wb = Workbook.blank();
    wb.sheets.push(new Sheet("Tabelle2"));
    wb.windows = [{ active: 0 }, { active: 1, hidden: true }];
    wb.sheets[0].views = [{ ...wb.sheets[0].view }];
    wb.sheets[1].views = [{ ...wb.sheets[1].view, zoom: 150, top: 9, active: { row: 2, col: 2 }, selection: [{ r0: 2, c0: 2, r1: 2, c1: 2 }] }];
    wb.active = 0;
    const out = write(wb);
    const book = /<bookViews>[\s\S]*?<\/bookViews>/.exec(part(out, "xl/workbook.xml"))![0];
    expect(book.match(/<workbookView\b/g)!.length).toBe(2);
    expect(book).toMatch(/<workbookView visibility="hidden"[^>]* activeTab="1"/);
    const s2 = sheetViews(part(out, "xl/worksheets/sheet2.xml"))!;
    expect(s2).toContain('<sheetView tabSelected="1" topLeftCell="A10" zoomScale="150" zoomScaleNormal="150" workbookViewId="1"><selection activeCell="C3" sqref="C3"/></sheetView>');
    // And it reads back as written.
    const back = readXlsx(out);
    expect(back.windows!.map((w) => [w.active, !!w.hidden])).toEqual([[0, false], [1, true]]);
    expect(back.sheets[1].views![0].zoom).toBe(150);
  });
});

describe("Seitenlayout and Benutzerdefinierte Ansichten", () => {
  it("reads the Seitenlayout view with its own zoom and writes it as Excel", () => {
    const wb = load("excel-pagelayout.xlsx");
    const v = wb.sheets[0].view;
    expect([v.layout, v.zoom, v.normalZoom]).toEqual([true, 85, 100]);
    const fresh = Workbook.blank();
    fresh.sheets[0].view = { ...v };
    expect(sheetViews(part(write(fresh), "xl/worksheets/sheet1.xml"))).toBe(rec.pagelayout.xml.sheets.sheet1.sheetViews);
    fresh.sheets[0].view = { ...v, noRuler: true };
    expect(sheetViews(part(write(fresh), "xl/worksheets/sheet1.xml"))).toBe(rec.pagelayout.xml_noruler.sheets.sheet1.sheetViews);
  });
  it("reads Excel's three custom views and writes them back unchanged", () => {
    const wb = load("excel-customviews.xlsx");
    expect(wb.customViews!.map((c) => [c.name, c.print, c.rowCol])).toEqual([["Nur Anzeige", false, false], ["Ohne Zeilen", true, true], ["Alles", true, true]]);
    expect(wb.sheets[0].customViews!.length).toBe(3);
    const out = write(wb);
    expect(/<customWorkbookViews>[\s\S]*?<\/customWorkbookViews>/.exec(part(out, "xl/workbook.xml"))?.[0]).toBe(rec.custom.xml.customWorkbookViews);
    expect(/<customSheetViews>[\s\S]*?<\/customSheetViews>/.exec(part(out, "xl/worksheets/sheet1.xml"))?.[0]).toBe(rec.custom.xml.sheets.sheet1.customSheetViews);
  });
});

describe("Ink as Excel stores it", () => {
  it("decodes InkML's differences as Excel encodes them, and encodes them the same way", () => {
    const saved = inkRec.excel_saved["xl/ink/ink1.xml"] as string;
    const trace = /<inkml:trace\b[^>]*>([\s\S]*?)<\/inkml:trace>/.exec(saved)![1];
    const pts = decodeTrace(trace, 3);
    expect(pts.length).toBe(40);
    expect(pts[0]).toEqual([1000, 1000, 0]);
    expect(encodeTrace(pts)).toBe(trace);
    // Excel scaled our strokes into the box it gave them, keeping the start.
    const sent = parseInkml(inkRec.sent[0].inkml).strokes[0].pts;
    const sx = (pts[39][0] - pts[0][0]) / (sent[39][0] - sent[0][0]);
    pts.forEach((p, i) => expect(Math.abs(p[0] - (1000 + (sent[i][0] - 1000) * sx))).toBeLessThanOrEqual(1));
    // Time stamps are kept.
    expect(pts.map((p) => p[2])).toEqual(sent.map((p) => p[2]));
  });
  it("reads Excel's four ink drawings: places, brushes, strokes", () => {
    const wb = load("excel-ink.xlsx");
    const inks = wb.sheets[0].inks!;
    expect(inks.length).toBe(4);
    expect(wb.sheets[0].pictures ?? []).toEqual([]);
    expect(inks.map((k) => k.strokes.length)).toEqual([1, 1, 1, 2]);
    expect(inks.map((k) => k.strokes[0].brush.color)).toEqual(["#000000", "#FF0000", "#FFFF00", "#0000FF"]);
    expect(inks[2].strokes[0].brush.kind).toBe("highlighter");
    expect(inks.map((k) => k.strokes[0].brush.width)).toEqual([0.35, 0.35, 6, 0.5]);
    // COM: Ink 1 at A2 (moved 30 pt right), Ink 2 B6 … (results/2026-11-08-excel-ink.json "com").
    expect(inks.map((k) => [k.from.col, k.from.row])).toEqual([[0, 1], [1, 5], [0, 9], [2, 1]]);
    for (const k of inks) {
      const xs = k.strokes.flatMap((s) => s.pts.map((p) => p.x)), ys = k.strokes.flatMap((s) => s.pts.map((p) => p.y));
      expect(Math.max(...xs)).toBeCloseTo(k.size.w, 3);
      if (Math.max(...ys) > 0) expect(Math.max(...ys)).toBeCloseTo(k.size.h, 3);
      expect(k.fallback?.startsWith("data:image/png;base64,")).toBe(true);
    }
  });
  it("writes Excel's ink back unchanged and a new stroke in Excel's form", () => {
    const wb = load("excel-ink.xlsx");
    const out = write(wb);
    expect(part(out, "xl/drawings/drawing1.xml")).toBe(inkRec.excel_saved["xl/drawings/drawing1.xml"]);
    const k = wb.sheets[0].inks![0];
    wb.sheets[0].inks!.push({ ...k, raw: undefined, rawKey: undefined, part: undefined, name: undefined, from: { col: 5, row: 5, x: 0, y: 0 }, to: { col: 7, row: 7, x: 0, y: 0 } });
    const out2 = write(wb);
    const d = part(out2, "xl/drawings/drawing1.xml");
    expect(d.match(/<xdr:contentPart\b/g)!.length).toBe(5);
    const anchors = d.match(/<xdr:twoCellAnchor editAs="oneCell">[\s\S]*?<\/xdr:twoCellAnchor>/g)!;
    const last = anchors[anchors.length - 1];
    expect(last).toMatch(/<mc:Choice xmlns:xdr14="[^"]+" Requires="xdr14"><xdr:contentPart xmlns:r="[^"]+" r:id="rId\d+"><xdr14:nvContentPartPr><xdr14:cNvPr id="\d+" name="Freihand \d+"\/>/);
    const f = files(out2);
    const inkParts = Object.keys(f).filter((n) => n.startsWith("xl/ink/"));
    expect(inkParts.length).toBe(5);
    expect(strFromU8(f["[Content_Types].xml"])).toContain('ContentType="application/inkml+xml"');
    const back = readXlsx(out2).sheets[0].inks!;
    expect(back.length).toBe(5);
    back[4].strokes[0].pts.forEach((p, i) => {
      expect(Math.abs(p.x - k.strokes[0].pts[i].x)).toBeLessThan(0.05);
      expect(Math.abs(p.y - k.strokes[0].pts[i].y)).toBeLessThan(0.05);
    });
  });
});
