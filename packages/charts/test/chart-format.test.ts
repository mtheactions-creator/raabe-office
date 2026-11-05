// The Format pane's settings (D-066): written to the chart part and read back unchanged, and drawn.

import { describe, expect, it } from "vitest";
import { chartScene, fitTrend, newChart, schemeColor, type ChartKind, type ChartModel } from "../src/index";
import { fixedScale } from "../src/paint";
import { chartFromXml, chartToXml } from "../../ooxml/src/xlsx/chart";

const series = () => [
  { name: { text: "A" }, cat: { cache: ["x", "y", "z"] }, val: { cache: [1, 3, 2] } },
  { name: { text: "B" }, cat: { cache: ["x", "y", "z"] }, val: { cache: [2, -1, 4] } },
];
const anchor = { col: 0, row: 0, x: 0, y: 0 }, to = { col: 8, row: 15, x: 0, y: 0 };

function roundTrip(m: ChartModel): ChartModel {
  const back = chartFromXml(chartToXml(m));
  delete back.raw;
  return back;
}

describe("chart formats in the file", () => {
  it("keeps element formats, axis options and labels", () => {
    const m = newChart("column", series(), anchor, to);
    m.dataLabels = true;
    m.gapWidth = 50;
    m.overlap = 20;
    m.format = {
      chart: { fill: { type: "solid", color: "#FFF2CC" }, line: { type: "solid", color: "#C00000", width: 2, dash: "dash" } },
      plot: { fill: { type: "solid", color: "#DDEBF7", transparency: 30 } },
      title: { text: { size: 18, bold: true, color: "#1F4E79" }, overlay: true },
      legend: { text: { italic: true } },
      valAxis: { min: -5, max: 10, majorUnit: 2.5, minorUnit: 0.5, reverse: true, numFmt: "0.0", majorTick: "out", labelPos: "low", line: { type: "solid", color: "#000000", width: 1 } },
      catAxis: { reverse: true, labelPos: "high" },
      majorGrid: { line: { type: "solid", color: "#A6A6A6", width: 1.5, dash: "sysDot" } },
      minorGrid: { show: true },
      catGrid: { show: true },
      dataLabels: { showCat: true, showSer: true, pos: "inEnd", separator: ", ", numFmt: "0%" },
    };
    m.series[0].fill = { type: "none" };
    m.series[1].line = { type: "solid", color: "#00B050", width: 1.25 };
    m.series[1].invertIfNegative = true;
    m.series[0].errorBars = { dir: "plus", type: "percentage", value: 10, noEndCap: true };
    m.series[1].trendline = { type: "poly", order: 3, showEq: true, showR2: false, forward: 1, name: "Kurve" };
    const back = roundTrip(m);
    expect(back.format).toEqual(m.format);
    expect(back.gapWidth).toBe(50);
    expect(back.overlap).toBe(20);
    expect(back.series[0].fill).toEqual({ type: "none" });
    expect(back.series[1].line).toEqual(m.series[1].line);
    expect(back.series[1].invertIfNegative).toBe(true);
    expect(back.series[0].errorBars).toEqual(m.series[0].errorBars);
    expect(back.series[1].trendline).toEqual(m.series[1].trendline);
  });

  it("keeps markers, smoothing and pie settings", () => {
    const l = newChart("line", series(), anchor, to);
    l.series[0].marker = { symbol: "diamond", size: 9 };
    l.series[0].smooth = true;
    const lb = roundTrip(l);
    expect(lb.series[0].marker).toEqual({ symbol: "diamond", size: 9 });
    expect(lb.series[0].smooth).toBe(true);
    const d = newChart("doughnut", series().slice(0, 1), anchor, to);
    d.firstSliceAngle = 90;
    d.holeSize = 40;
    const db = roundTrip(d);
    expect([db.firstSliceAngle, db.holeSize]).toEqual([90, 40]);
  });

  it("leaves a default chart without format", () => {
    for (const k of ["column", "line", "pie", "scatter", "area", "bar"] as ChartKind[]) {
      const back = roundTrip(newChart(k, series(), anchor, to));
      expect(back.format).toBeUndefined();
      expect(back.series.every((s) => !s.fill && !s.line && !s.marker)).toBe(true);
    }
  });
});

describe("drawing the formats", () => {
  it("uses fixed bounds and units", () => {
    expect(fixedScale(0, 7, { min: 0, max: 10, majorUnit: 2.5 }, 10)).toEqual({ min: 0, max: 10, major: 2.5 });
    expect(fixedScale(3, 7, { min: 0 }, 10).min).toBe(0);
    const m = newChart("column", series(), anchor, to);
    m.format = { valAxis: { min: -5, max: 10 } };
    expect(chartScene(m, 360, 216).valueScale).toMatchObject({ min: -5, max: 10 });
  });

  it("fits every trendline type", () => {
    const xs = [1, 2, 3, 4, 5], ys = [2.7, 7.4, 20.1, 54.6, 148.4];
    expect(fitTrend({ type: "exp" }, xs, ys)!.equation).toBe("y = 0,9955e1,0012x");
    expect(fitTrend({ type: "poly", order: 2 }, [1, 2, 3, 4], [1, 4, 9, 16])!.equation).toBe("y = x²");
    expect(fitTrend({ type: "movingAvg", period: 2 }, [1, 2, 3], [2, 4, 6])!.points).toEqual([[2, 3], [3, 5]]);
  });
});

describe("Farben ändern", () => {
  // Excel 365's new default theme, in which the schemes were recorded.
  const theme = { accent1: "#156082", accent2: "#E97132", accent3: "#196B24", accent4: "#0F9ED5", accent5: "#A02B93", accent6: "#4EA72E", dk1: "#000000" };
  const measured: Record<string, string[]> = {
    "10/2": ["#156082", "#E97132"], "11/2": ["#156082", "#196B24"], "13/2": ["#4EA72E", "#A02B93"],
    "14/2": ["#115472", "#8599AA"], "14/3": ["#0F4E6A", "#156082", "#A1AFBB"], "14/6": ["#0C445E", "#10516E", "#135B7C", "#5C7D95", "#96A6B4", "#BCC5CD"],
    "20/6": ["#5F5F5F", "#B3B3B3", "#898989", "#212121", "#DADADA", "#AAAAAA"], "21/3": ["#A1AFBB", "#156082", "#0F4E6A"],
  };
  for (const [key, want] of Object.entries(measured)) {
    it(key, () => {
      const [id, n] = key.split("/").map(Number);
      const got = want.map((_, i) => schemeColor(id, i, n, theme));
      // Within two steps of 255 per channel (Excel rounds in its own colour space).
      got.forEach((g, i) => {
        for (const k of [1, 3, 5]) expect(Math.abs(parseInt(g.slice(k, k + 2), 16) - parseInt(want[i].slice(k, k + 2), 16))).toBeLessThanOrEqual(3);
      });
    });
  }
});
