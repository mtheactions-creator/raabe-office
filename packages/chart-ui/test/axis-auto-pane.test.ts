// Step 12 (D-200): the axis options show the automatic Grenzen and Einheiten (Office never shows empty boxes there).
import { describe, expect, it } from "vitest";
import { newChart, type ChartModel } from "@raabe/charts";
import { axisSections } from "../src/paneoptions";

const anchor = { col: 0, row: 0, x: 0, y: 0 }, to = { col: 8, row: 15, x: 0, y: 0 };

function model(kind: "column" | "scatter", vals: number[]): ChartModel {
  const s = { name: { text: "Werte" }, cat: { cache: kind === "scatter" ? [1, 2.5, 4, 7] : ["1", "2", "3", "4"] }, val: { cache: vals } };
  const m = newChart(kind, [s], anchor, to);
  m.legend = kind === "scatter" ? "none" : "b";
  m.title = { ...m.title, show: true };
  return m;
}

// 360 x 216 pt, Excel's default chart: 12.7 x 7.62 cm.
const host = (m: ChartModel) => ({ edit: (fn: (x: ChartModel) => void) => fn(m), select: () => {}, size: () => ({ w: 12.7, h: 7.62 }), resize: () => {} });

function fields(m: ChartModel, which: "valAxis" | "catAxis") {
  const sec = axisSections(m, which, host(m)).find((s) => s.title === "Achsenoptionen")!;
  const ctl = (sec.controls as () => { type: string; label: string; value?: () => unknown; reset?: { isAuto: () => boolean } }[])();
  const get = (label: string) => ctl.find((c) => c.type === "number" && c.label === label)!;
  return { get, ctl };
}

describe("automatic axis values in the Format pane", () => {
  it("shows Excel's automatic Minimum, Maximum, Hauptstriche and Teilstriche, marked automatic", () => {
    const { get } = fields(model("column", [12, 47, 120, 83]), "valAxis");
    expect(get("Minimum").value!()).toBe(0);
    expect(get("Maximum").value!()).toBe(140);
    expect(get("Hauptstriche").value!()).toBe(20);
    expect(get("Teilstriche").value!()).toBe(4);
    expect(get("Maximum").reset!.isAuto()).toBe(true);
  });

  it("shows a fixed value as set, and the others still automatic", () => {
    const m = model("column", [12, 47, 120, 83]);
    m.format = { valAxis: { max: 200 } };
    const { get } = fields(m, "valAxis");
    expect(get("Maximum").value!()).toBe(200);
    expect(get("Maximum").reset!.isAuto()).toBe(false);
    expect(get("Minimum").value!()).toBe(0);
  });

  it("fills a scatter chart's X axis too", () => {
    const { get } = fields(model("scatter", [101, 102, 104, 103]), "catAxis");
    expect(get("Minimum").value!()).toBe(0);
    expect(get("Maximum").value!()).toBe(8);
    expect(get("Hauptstriche").value!()).toBe(1);
  });
});
