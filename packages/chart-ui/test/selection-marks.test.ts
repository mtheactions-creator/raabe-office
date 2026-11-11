import { describe, expect, it } from "vitest";
import { MARK_LOOK, chartMarks, drawMarks } from "../src/selection";

// The chart's selection marks as Excel draws them (D-170, results/2026-11-11-excel-chart-selection.json).
describe("chart selection marks", () => {
  it("frames the chart area and puts eight sizing handles on its corners and side middles", () => {
    const m = chartMarks(480, 288);
    expect(m.frames).toEqual([{ r: { x: 0, y: 0, w: 480, h: 288 }, color: MARK_LOOK.frameDark }]);
    expect(m.sizing).toHaveLength(8);
    expect(m.sizing).toContainEqual({ x: 240, y: 0 });
    expect(m.sizing).toContainEqual({ x: 480, y: 144 });
    expect(m.sizing).not.toContainEqual({ x: 240, y: 144 });
  });

  it("draws rings of Excel's size: 17 device px at 150 %", () => {
    const layer = document.createElementNS("http://www.w3.org/2000/svg", "svg");
    drawMarks(layer, [chartMarks(100, 50)], 1.5);
    const c = layer.querySelector("circle")!;
    const outer = (Number(c.getAttribute("r")) + Number(c.getAttribute("stroke-width")) / 2) * 2 * 1.5;
    expect(outer).toBeCloseTo(17, 5);
    expect(c.getAttribute("stroke")).toBe("#616161");
    // The frame: one device pixel wide, on whole device pixels.
    const f = layer.querySelector("path")!;
    expect(Number(f.getAttribute("stroke-width"))).toBeCloseTo(1 / 1.5, 5);
  });
});
