// The Sparkline tab's operations against Excel (D-085): Gruppieren, Gruppierung aufheben, Löschen and new groups
// give the groups Excel's files hold after the same COM calls (probe_excel_sparklines.py); the drawing puts
// columns, points and the axis where Excel's CopyPicture of the cells has them (measured at 150 %: a cell of
// 336 x 119 device pixels, 1 pt = 2 px).

import { readFileSync } from "node:fs";
import { describe, expect, it } from "vitest";
import type { SparklineGroup } from "@raabe/sheet-engine";
import { readSparklines } from "@raabe/ooxml";
import { clearSparklines, groupSparklines, newGroup, sparkItems, ungroupSparklines } from "../src/sparklines";
import { drawSparkline, pointColor, sparkRanges } from "../src/grid/sparkline-draw";

const FIX = JSON.parse(readFileSync("packages/ooxml/test/fixtures/excel-sparklines.json", "utf-8")) as { theme: Record<string, string>; cases: Record<string, { xml: string }> };
const ext = (xml: string) => `<extLst><ext uri="{05C60535-1F16-4fd2-B633-F4F36F0B64E0}">${xml}</ext></extLst>`;
/** A case's groups with the sheet names left out of the references (each case has its own sheet). */
const read = (name: string): SparklineGroup[] => readSparklines(ext(FIX.cases[name].xml), { theme: FIX.theme, indexed: [] })
  .map((g) => ({ ...g, items: g.items.map((i) => ({ ...i, f: i.f.replace(/^.*!/, "") })) }));
const rect = (r0: number, c0: number, r1 = r0, c1 = c0) => ({ r0, c0, r1, c1 });
const F = 5;

describe("Sparkline tab operations as Excel's files show them", () => {
  it("a new group is Excel's (line, column, Gewinn/Verlust)", () => {
    for (const [name, t] of [["line", "line"], ["column", "column"], ["winloss", "stacked"]] as const) {
      const items = sparkItems("x", rect(0, 0, 2, 4), rect(0, F, 2, F))!.map((i) => ({ ...i, f: i.f.replace(/^.*!/, "") }));
      expect(newGroup(t, items)).toEqual(read(name)[0]);
    }
  });

  it("a row of places takes the data's columns", () => {
    expect(sparkItems("bycolumn", rect(0, 0, 2, 4), rect(4, 0, 4, 4))!.map((i) => ({ ...i, f: i.f.replace(/^.*!/, "") }))).toEqual(read("bycolumn")[0].items);
    expect(sparkItems("x", rect(0, 0, 2, 4), rect(0, F, 1, F))).toBeNull();
  });

  it("Gruppieren takes the active cell's group", () => {
    expect(groupSparklines(read("two-groups"), [rect(0, F, 2, F)], { row: 2, col: F })).toEqual(read("group"));
    expect(groupSparklines(read("two-groups"), [rect(0, F, 2, F)], { row: 0, col: F })).toEqual(read("group-first"));
  });

  it("Gruppierung aufheben makes each selected sparkline a group, new ones first", () => {
    const high = read("line").map((g) => ({ ...g, high: true }));
    expect(ungroupSparklines(high, [rect(0, F, 2, F)])).toEqual(read("ungroup"));
    expect(ungroupSparklines(read("line"), [rect(1, F)])).toEqual(read("ungroup-one"));
  });

  it("Löschen: the selected sparklines or their whole groups", () => {
    expect(clearSparklines(read("line"), [rect(1, F)], false)).toEqual(read("clear-one"));
    expect(clearSparklines(read("two-groups"), [rect(1, F)], true)).toEqual(read("clear-group"));
  });
});

describe("drawing as Excel draws", () => {
  function mock() {
    const rects: number[][] = [], paths: string[] = [];
    let path: number[][] = [];
    const ctx = {
      fillStyle: "", strokeStyle: "", lineWidth: 1, lineJoin: "", lineCap: "",
      save() {}, restore() {}, clip() {}, rect() {}, stroke() { paths.push(`stroke:${path.map((p) => p.join(",")).join(" ")}`); },
      beginPath() { path = []; }, moveTo(x: number, y: number) { path.push([x, y]); }, lineTo(x: number, y: number) { path.push([x, y]); },
      closePath() {}, fill() { paths.push(`fill:${this.fillStyle}:${path.map((p) => p.map((v) => Math.round(v)).join(",")).join(" ")}`); },
      fillRect(x: number, y: number, w: number, h: number) { rects.push([Math.round(x), Math.round(y), Math.round(x + w) - 1, Math.round(y + h) - 1, this.fillStyle as unknown as number]); },
    };
    return { ctx: ctx as unknown as CanvasRenderingContext2D, rects, paths };
  }
  const box = { x: 0, y: 0, w: 336, h: 119 };

  it("columns: 72 % of a slot, from the smallest value up when all are positive (Excel's L4)", () => {
    const g = newGroup("column", []);
    const values = [4, 5, 6, 7, 8, 9, 10];
    const m = mock();
    drawSparkline(m.ctx, g, { values }, sparkRanges(g, [{ values }])[0]!, box, 2, 1);
    // Excel: x 9-42, 56-89, 104-137, 151-184, 198-231, 246-279, 293-326; tops 115, 96, 78, 59, 41, 22, 4; bottoms 115.
    const want = [[9, 42, 115], [56, 89, 96], [104, 137, 78], [151, 184, 59], [198, 231, 41], [246, 279, 22], [293, 326, 4]];
    m.rects.forEach((r, i) => {
      expect(Math.abs(r[0] - want[i][0])).toBeLessThanOrEqual(1);
      expect(Math.abs(r[2] - want[i][1])).toBeLessThanOrEqual(1);
      expect(Math.abs(r[1] - want[i][2])).toBeLessThanOrEqual(1);
      expect(Math.abs(r[3] - 115)).toBeLessThanOrEqual(1);
    });
  });

  it("columns cross zero at zero; Gewinn/Verlust fills the halves (Excel's L1 and N1)", () => {
    const values = [3, 5, -2, 8, 4, 6, 1];
    const m = mock();
    const g = newGroup("column", []);
    drawSparkline(m.ctx, g, { values }, sparkRanges(g, [{ values }])[0]!, box, 2, 1);
    expect(m.rects[2][1]).toBeGreaterThanOrEqual(92); // the negative bar hangs from zero (Excel: 94-115)
    expect(m.rects[2][3]).toBe(115);
    expect(m.rects[3][1]).toBe(4); // the highest bar reaches the top
    expect(Math.abs(m.rects[0][3] - 93)).toBeLessThanOrEqual(1);
    const w = mock();
    const wl = newGroup("stacked", []);
    drawSparkline(w.ctx, wl, { values }, sparkRanges(wl, [{ values }])[0]!, box, 2, 1);
    expect(w.rects[0].slice(0, 4)).toEqual([9, 4, 42, 59]);
    expect(w.rects[2].slice(0, 4)).toEqual([104, 60, 137, 115]);
    expect(w.rects[2][4]).toBe("#D00000"); // Gewinn/Verlust shows negative points from the start
  });

  it("line points at the slots' centres, diamonds of 2 pt + 4/3 of the Stärke (Excel's J4 and Q4)", () => {
    const values = [4, 5, 6, 7, 8, 9, 10];
    const g = { ...newGroup("line", []), markers: true };
    const m = mock();
    drawSparkline(m.ctx, g, { values }, sparkRanges(g, [{ values }])[0]!, box, 2, 1);
    const line = m.paths[0].replace("stroke:", "").split(" ").map((p) => p.split(",").map(Number));
    expect(Math.abs(line[0][0] - 26)).toBeLessThanOrEqual(1);
    expect(Math.abs(line[6][0] - 309.5)).toBeLessThanOrEqual(1.5);
    // J4: markers 6 px across at 0.75 pt; Q4: 20 px at 6 pt.
    const diamond = m.paths[1].split(":")[2].split(" ").map((p) => p.split(",").map(Number));
    expect(diamond[1][0] - diamond[3][0]).toBe(6);
    const q = mock();
    drawSparkline(q.ctx, { ...g, lineWeight: 6 }, { values }, sparkRanges(g, [{ values }])[0]!, box, 2, 1);
    const d2 = q.paths[1].split(":")[2].split(" ").map((p) => p.split(",").map(Number));
    expect(d2[1][0] - d2[3][0]).toBe(20);
  });

  it("point colours: Höchst- and Tiefpunkt before Erster/Letzter, then Negative, then Datenpunkte (Excel's K and M)", () => {
    const g = { ...newGroup("column", []), high: true, low: true, first: true, last: true, negative: true };
    const v = [4, 5, 6, 7, 8, 9, 10];
    expect(pointColor(g, v, 0)).toBe("low");
    expect(pointColor(g, v, 6)).toBe("high");
    const w = [3, 5, -2, 8, 4, 6, 1];
    expect(pointColor(g, w, 0)).toBe("first");
    expect(pointColor(g, w, 2)).toBe("low");
    expect(pointColor(g, w, 6)).toBe("last");
    expect(pointColor({ ...g, low: false }, w, 2)).toBe("negative");
    expect(pointColor(newGroup("column", []), w, 2)).toBeNull(); // negative columns in the series colour unless shown
    expect(pointColor({ ...newGroup("line", []), markers: true }, w, 3)).toBe("markers");
  });

  it("the axis only where zero lies in the range; the same range for the group (Excel's O and P)", () => {
    const g = { ...newGroup("line", []), displayXAxis: true };
    const m = mock();
    drawSparkline(m.ctx, g, { values: [2, 7, 7, 1, 5, 3, 4] }, { lo: 1, hi: 7 }, box, 2, 1);
    expect(m.rects.length).toBe(0);
    const n = mock();
    drawSparkline(n.ctx, g, { values: [3, 5, -2, 8, 4, 6, 1] }, { lo: -2, hi: 8 }, box, 2, 1);
    expect(Math.abs(n.rects[0][1] - 93)).toBeLessThanOrEqual(1);
    const grp = { ...newGroup("column", []), minAxisType: "group" as const, maxAxisType: "group" as const };
    expect(sparkRanges(grp, [{ values: [1, -3] }, { values: [4, 10] }])).toEqual([{ lo: -3, hi: 10 }, { lo: -3, hi: 10 }]);
    expect(sparkRanges({ ...grp, minAxisType: "custom", manualMin: 0 }, [{ values: [1, -3] }])).toEqual([{ lo: 0, hi: 1 }]);
  });
});
