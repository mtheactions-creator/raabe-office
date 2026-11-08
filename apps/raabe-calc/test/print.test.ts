import { describe, expect, it } from "vitest";
import { cutPages, fitZoom, headerSections } from "../src/print";
import { printRowPx, rowHeights } from "../src/print-metrics";

describe("Kopf- und Fußzeile", () => {
  const f = { page: 2, pages: 5, sheet: "Tabelle1", file: "Noten", now: new Date(2026, 9, 28, 7, 5) };
  it("fills in Excel's codes per section", () => {
    expect(headerSections("&LKlasse 7b&C&A&RSeite &P von &N", f)).toEqual({ L: "Klasse 7b", C: "Tabelle1", R: "Seite 2 von 5" });
    expect(headerSections("&D &T &F", f)).toEqual({ L: "", C: "28.10.2026 07:05 Noten", R: "" });
  });
  it("leaves format codes out and keeps &&", () => {
    expect(headerSections('&C&"Calibri,Fett"&14Titel && mehr&KFF0000!', f)).toEqual({ L: "", C: "Titel & mehr!", R: "" });
  });
});

describe("Excel's printed sizes (D-094, probe_excel_print.py)", () => {
  const prn = (px: number) => Math.round((px * 72) / 96 * 100) / 100;
  it("prints rows at the standard font's printed line height", () => {
    const aptos = rowHeights("Aptos Narrow", 11, 14.5);
    expect(prn(printRowPx(undefined, aptos))).toBe(15.24);
    // Recorded on A4: 8, 10, 15, 30 and 60 pt rows in Aptos Narrow 11 and Calibri 11.
    expect([8, 10, 15, 15.75, 21, 30, 60].map((h) => prn(printRowPx(h, aptos)))).toEqual([8.4, 10.44, 15.72, 16.56, 22.08, 31.44, 63.0]);
    const calibri = rowHeights("Calibri", 11, 14.5);
    expect([8, 15, 30].map((h) => prn(printRowPx(h, calibri)))).toEqual([7.92, 15, 30]);
    // Unknown fonts print as their rows say.
    expect(rowHeights("Comic Sans MS", 11, 15)).toEqual({ printed: 15, screen: 15 });
  });
  it("fits to pages at Excel's whole percent", () => {
    expect(fitZoom({ w: 1000, h: 300 }, { w: 920, h: 600 }, { wide: 1, tall: 1 })).toBe(92);
    expect(fitZoom({ w: 500, h: 300 }, { w: 920, h: 600 }, { wide: 1, tall: 1 })).toBe(100);
    expect(fitZoom({ w: 500, h: 3000 }, { w: 920, h: 600 }, { wide: 1 })).toBe(100);
  });
  it("cuts pages as Excel: whole rows, manual breaks, room for Drucktitel", () => {
    expect(cutPages(0, 9, () => 10, 35)).toEqual([[0, 2], [3, 5], [6, 8], [9, 9]]);
    expect(cutPages(0, 9, () => 10, 35, [5])).toEqual([[0, 2], [3, 4], [5, 7], [8, 9]]);
    expect(cutPages(0, 9, () => 10, 35, undefined, 10, [0, 0])).toEqual([[0, 2], [3, 4], [5, 6], [7, 8], [9, 9]]);
  });
});
