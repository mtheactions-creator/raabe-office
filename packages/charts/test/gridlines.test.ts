// Step 12: horizontal and vertical gridlines, major and minor, each on its own (the owner: "there should be separate
// options for horizontal and vertical grid lines"), drawn where Excel draws them and kept in the file as Excel writes them.

import { describe, expect, it } from "vitest";
import { chartScene, newChart, type ChartKind, type ChartModel, type Prim } from "../src/index";
import { chartFromXml, chartToXml } from "../../ooxml/src/xlsx/chart";

const anchor = { col: 0, row: 0, x: 0, y: 0 }, to = { col: 8, row: 15, x: 0, y: 0 };
const cats = ["a", "b", "c", "d"];

function chart(kind: ChartKind): ChartModel {
  const s = { name: { text: "A" }, cat: { cache: kind.startsWith("scatter") ? [1, 2.5, 4, 7] : cats }, val: { cache: [3, 5, 2, 4] } };
  return newChart(kind, [s], anchor, to);
}
const paths = (m: ChartModel, part: string) => chartScene(m, 400, 250).prims.filter((p): p is Extract<Prim, { t: "path" }> => p.t === "path" && p.part === part);
const xOf = (p: Extract<Prim, { t: "path" }>) => Number(/^M([\d.]+)/.exec(p.d)![1]);
function roundTrip(m: ChartModel): ChartModel {
  const back = chartFromXml(chartToXml(m));
  delete back.raw;
  return back;
}

describe("gridlines of an XY (scatter) chart", () => {
  it("shows horizontal and vertical major gridlines in a new chart, as Excel", () => {
    const m = chart("scatter");
    expect(paths(m, "grid").length).toBeGreaterThan(0);
    expect(paths(m, "catGrid").length).toBeGreaterThan(0);
  });

  it("switches the vertical ones off without the horizontal ones, and back", () => {
    const m = chart("scatter");
    m.format = { ...m.format, catGrid: { show: false } };
    expect(paths(m, "catGrid")).toHaveLength(0);
    expect(paths(m, "grid").length).toBeGreaterThan(0);
    m.gridlines = false;
    m.format = { catGrid: { show: true } };
    expect(paths(m, "grid")).toHaveLength(0);
    expect(paths(m, "catGrid").length).toBeGreaterThan(0);
  });

  it("draws the vertical minor gridlines between the major ones", () => {
    const m = chart("scatter");
    m.format = { catGrid: { show: true }, catMinorGrid: { show: true } };
    const majors = paths(m, "catGrid").map(xOf), minors = paths(m, "catMinorGrid").map(xOf);
    expect(minors.length).toBeGreaterThan(majors.length);
    for (const x of minors) expect(majors.some((y) => Math.abs(x - y) < 0.01)).toBe(false);
  });

  it("keeps each of the four and its line in the file", () => {
    const m = chart("scatter");
    m.gridlines = false;
    m.format = { catGrid: { show: true, line: { type: "solid", color: "#C00000", width: 1 } }, catMinorGrid: { show: true }, minorGrid: { show: true } };
    const back = roundTrip(m);
    expect(back.gridlines).toBe(false);
    expect(back.format?.catGrid?.show).toBe(true);
    expect(back.format?.catGrid?.line).toMatchObject({ type: "solid", color: "#C00000" });
    expect(back.format?.catMinorGrid?.show).toBe(true);
    expect(back.format?.minorGrid?.show).toBe(true);
  });
});

describe("the category axis' minor gridlines", () => {
  for (const kind of ["column", "line"] as ChartKind[]) {
    it(`run through the middle of each category in a ${kind} chart (Excel's export)`, () => {
      const m = chart(kind);
      m.format = { catGrid: { show: true }, catMinorGrid: { show: true } };
      const majors = paths(m, "catGrid").map(xOf).sort((a, b) => a - b), minors = paths(m, "catMinorGrid").map(xOf).sort((a, b) => a - b);
      expect(minors).toHaveLength(majors.length - 1);
      minors.forEach((x, i) => expect(Math.abs(x - (majors[i] + majors[i + 1]) / 2)).toBeLessThan(0.01));
    });
  }

  it("keeps its own line in the file, once", () => {
    const m = chart("column");
    m.format = { catMinorGrid: { show: true, line: { type: "solid", color: "#00B050", width: 0.5 } } };
    const xml = chartToXml(m);
    expect(xml.match(/<c:minorGridlines>/g)).toHaveLength(1);
    expect(roundTrip(m).format?.catMinorGrid?.line).toMatchObject({ type: "solid", color: "#00B050" });
  });
});
