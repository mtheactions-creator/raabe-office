// Teilen (D-093) against Excel's recorded behaviour (probe_excel_windows.py; results/2026-11-08-excel-windows.json
// and the COM runs noted in DECISIONS D-093): where Teilen puts the bars for an active cell, which pane is then
// active, how the panes scroll, Fenster fixieren on a split and Fixierung aufheben; the Seitenlayout view's pages.

import { readFileSync } from "node:fs";
import { describe, expect, it } from "vitest";
import { Sheet, type View } from "@raabe/sheet-engine";
import { readXlsx } from "@raabe/ooxml";
import { Axis } from "../src/grid/metrics";
import { TWIPS_PER_PX, freezeSplit, pxToY, splitPanes, toggleSplit, unfreezeToSplit, yToPx, type Geometry } from "../src/split";
import { GapAxis, layoutPages } from "../src/pagelayout";
import { decodeTrace } from "../../../packages/ooxml/src/xlsx/ink";

/** A window like Excel's in the recording: 24 columns and 48 rows in view (A1:X48). */
function geometry(): Geometry {
  const rows = new Axis(20, 1048576), cols = new Axis(64, 16384);
  rows.set(new Map(), 20);
  cols.set(new Map(), 64);
  return { rows, cols, zoom: 1, headerW: 27, headerH: 20, width: 27 + 24 * 64, height: 20 + 48 * 20 };
}
const view = (cell: [number, number], scroll: [number, number] = [0, 0]): View => ({
  active: { row: cell[0], col: cell[1] }, selection: [{ r0: cell[0], c0: cell[1], r1: cell[0], c1: cell[1] }],
  top: scroll[0], left: scroll[1], zoom: 100, showGrid: true, showHeaders: true,
});
/** Excel's SplitRow/SplitColumn: rows/columns above/left of the bars, counted from the top-left pane's first one. */
function excelSplit(v: View, g: Geometry): { row: number; col: number; panes: [number, number][]; active: string } {
  const s = v.split!;
  const panes = splitPanes(v, g);
  const rowOf = (px: number) => g.rows.at(g.rows.pos(s.top) + (px - g.headerH) / g.zoom + 0.01) - s.top;
  const colOf = (px: number) => g.cols.at(g.cols.pos(s.left) + (px - g.headerW) / g.zoom + 0.01) - s.left;
  return { row: s.y ? rowOf(s.y / TWIPS_PER_PX) : 0, col: s.x ? colOf(s.x / TWIPS_PER_PX) : 0, panes: panes.map((p) => [p.row + 1, p.col + 1]), active: s.pane ?? "topLeft" };
}

describe("Teilen where Excel splits", () => {
  // Window.Split = True with the active cell (and scroll) as recorded through COM.
  const cases: { cell: string; at: [number, number]; scroll?: [number, number]; row: number; col: number; panes: [number, number][]; active: string }[] = [
    { cell: "A1", at: [0, 0], row: 24, col: 12, panes: [[1, 1], [1, 13], [25, 1], [25, 13]], active: "topLeft" },
    { cell: "C1", at: [0, 2], row: 0, col: 2, panes: [[1, 1], [1, 3]], active: "topRight" },
    { cell: "A5", at: [4, 0], row: 4, col: 0, panes: [[1, 1], [5, 1]], active: "bottomLeft" },
    { cell: "C5", at: [4, 2], row: 4, col: 2, panes: [[1, 1], [1, 3], [5, 1], [5, 3]], active: "bottomRight" },
    { cell: "F20 scrolled to C10", at: [19, 5], scroll: [9, 2], row: 10, col: 3, panes: [[10, 3], [10, 6], [20, 3], [20, 6]], active: "bottomRight" },
    { cell: "B1", at: [0, 1], row: 0, col: 1, panes: [[1, 1], [1, 2]], active: "topRight" },
  ];
  for (const c of cases) it(c.cell, () => {
    const g = geometry();
    const v = view(c.at, c.scroll);
    expect(toggleSplit(v, g)).toBe(true);
    const e = excelSplit(v, g);
    expect([e.row, e.col]).toEqual([c.row, c.col]);
    expect(e.panes).toEqual(c.panes);
    expect(e.active).toBe(c.active);
  });

  it("Teilen again removes the split; Teilen while frozen gives a split at the frozen rows and columns", () => {
    const g = geometry();
    const v = view([4, 2]);
    toggleSplit(v, g);
    toggleSplit(v, g);
    expect(v.split).toBeUndefined();
    const f = view([4, 2]);
    f.freeze = { row: 4, col: 2 };
    toggleSplit(f, g);
    expect(f.freeze).toBeUndefined();
    expect(excelSplit(f, g)).toMatchObject({ row: 4, col: 2 });
  });

  it("Fenster fixieren on a split freezes at the bars, Fixierung aufheben gives the split back (recorded)", () => {
    const g = geometry();
    const v = view([4, 2]);
    toggleSplit(v, g);
    freezeSplit(v, g);
    expect(v.freeze).toEqual({ row: 4, col: 2 });
    expect(v.splitFrozen).toBe(true);
    unfreezeToSplit(v, g);
    expect(v.freeze).toBeUndefined();
    expect(excelSplit(v, g)).toMatchObject({ row: 4, col: 2 });
  });

  it("reads Excel's bars at the recorded places: xSplit 4000 = row headers + three 60 pt columns", () => {
    const wb = readXlsx(new Uint8Array(readFileSync("packages/ooxml/test/fixtures/excel-split.xlsx")));
    const rec = JSON.parse(readFileSync("packages/ooxml/test/fixtures/excel-windows.json", "utf-8"));
    rec.splits.forEach((c: { name: string; com: { splitHorizontal: number; splitVertical: number; headers: boolean; zoom: number } }, i: number) => {
      const s = wb.sheets[i].view.split!;
      // The bar lies a header's width (about 20 pt at 100 %) to the right of COM's SplitHorizontal.
      const head = c.com.headers ? 20 * c.com.zoom / 100 : 0;
      if (s.x) expect(Math.abs(s.x / 20 - c.com.splitHorizontal - head), c.name).toBeLessThanOrEqual(2);
      if (s.y) expect(Math.abs(s.y / 20 - c.com.splitVertical - (c.com.headers ? 14.5 * c.com.zoom / 100 : 0)), c.name).toBeLessThanOrEqual(2);
    });
  });
});

describe("the horizontal bar counts the file's rows (recorded ySplit 1740 = 14,5 pt header + 5 rows of 14,5 pt)", () => {
  it("lies between rows 5 and 6 whatever our row height in pixels", () => {
    const g = { ...geometry(), rowPt: () => 14.5, headPt: 14.5 };
    const v = view([0, 0]);
    v.split = { x: 0, y: 1740, top: 0, left: 0 };
    expect(yToPx(v, g, 1740)).toBeCloseTo(g.headerH + 5 * 20, 6);
    expect(pxToY(v, g, g.headerH + 5 * 20)).toBe(1740);
    // At 150 % Excel wrote 2630 (131,5 pt): 21,75 pt header + 5 rows of 21,75 pt is 2610, Excel's pixels round it up.
    const z = { ...g, zoom: 1.5 };
    expect(Math.abs(yToPx(v, z, 2630) - (z.headerH + 5 * 20 * 1.5))).toBeLessThan(3);
  });
});

describe("Seitenlayout pages", () => {
  it("spreads rows and columns over A4 pages with Excel's normal margins", () => {
    const s = new Sheet("T");
    const rows = new Axis(20, 1048576), cols = new Axis(64, 16384);
    rows.set(new Map(), 20);
    cols.set(new Map(), 64);
    const p = layoutPages(s, { rows, cols });
    // A4 printable height (11,69 - 0,75 - 0,75 in) at 96 dpi holds 48 rows of 20 px (978 px); width (8,27 - 1,4 in) 10 columns of 64 px.
    expect(p.rows.starts.slice(0, 3)).toEqual([0, 48, 96]);
    expect(p.cols.starts.slice(0, 2)).toEqual([0, 10]);
    expect(p.rows).toBeInstanceOf(GapAxis);
    // Rows inside a page keep their sizes; between pages the margins and the gap are added.
    expect(p.rows.pos(1) - p.rows.pos(0)).toBe(20);
    expect(p.rows.pos(48) - p.rows.pos(47)).toBeGreaterThan(20 + 0.75 * 96 * 2);
    // Each page is a sheet of paper as tall as A4.
    expect(p.paperY[1] - p.paperY[0]).toBeGreaterThanOrEqual(p.paperH);
    expect(p.rows.at(p.rows.pos(60) + 1)).toBe(60);
  });
});

describe("InkML as Excel writes it", () => {
  it("decodes a trace with second differences and qualifiers per channel", () => {
    expect(decodeTrace("1000 1000 0,'163'255'8,\"0\"0\"0,-1 1 0", 3)).toEqual([[1000, 1000, 0], [1163, 1255, 8], [1326, 1510, 16], [1488, 1766, 24]]);
  });
});
