// The chart engine against Excel's own numbers, recorded by tools/feel-lab/probe_excel_charts.py
// (results/2026-10-27-excel-charts.json): automatic axis scaling, series from selections,
// histogram bins, trendline labels and the default layout.

import { readFileSync } from "node:fs";
import { describe, expect, it } from "vitest";
import {
  autoScale, calibriMeasure, chartScene, histogramBins, linearFit, maxIntervals, equationText, r2Text, newChart,
  seriesFromBlock, titleText, defaultLegend, formatGeneral, type ChartKind, type CellValue,
} from "../src/index";

const data = JSON.parse(readFileSync("tools/feel-lab/results/2026-10-27-excel-charts.json", "utf-8"));

const close = (a: number, b: number) => Math.abs(a - b) <= 1e-9 * Math.max(1, Math.abs(b));

interface AxisCase { values: number[]; kind: string; size: number[]; axis: { min: number; max: number; major: number }; inner: { width: number; height: number } }

function scaleFor(values: number[], length: number, horizontal: boolean) {
  return autoScale(Math.min(...values), Math.max(...values), maxIntervals(length, horizontal));
}

describe("automatic value axis (Excel's recorded scales)", () => {
  const cases: AxisCase[] = data.axes.axes;
  const failures: string[] = [];
  for (const c of cases) {
    const horizontal = c.kind === "bar";
    const got = scaleFor(c.values, horizontal ? c.inner.width : c.inner.height, horizontal);
    if (!(close(got.min, c.axis.min) && close(got.max, c.axis.max) && close(got.major, c.axis.major))) {
      failures.push(`${c.kind} ${c.size} [${Math.min(...c.values)}, ${Math.max(...c.values)}]: Excel ${c.axis.min}..${c.axis.max}/${c.axis.major}, ours ${got.min}..${got.max}/${got.major}`);
    }
  }
  it(`matches ${cases.length} recorded axes`, () => {
    expect(failures).toEqual([]);
  });

  it("matches the interval limits recorded for axis lengths", () => {
    const bad: string[] = [];
    for (const c of data.limits as { kind: string; factor: number; m: number; axis: { min: number; max: number; major: number }; inner: { width: number; height: number } }[]) {
      const horizontal = c.kind === "bar";
      const got = scaleFor([0, Number((c.m * c.factor).toPrecision(12))], horizontal ? c.inner.width : c.inner.height, horizontal);
      if (!(close(got.max, c.axis.max) && close(got.major, c.axis.major))) bad.push(`${c.kind} ${c.factor} ${horizontal ? c.inner.width : c.inner.height} m=${c.m}: Excel ${c.axis.max}/${c.axis.major}, ours ${got.max}/${got.major}`);
    }
    // Horizontal axes depend on more than their length in Excel; allow a few misses there.
    expect(bad.filter((b) => b.startsWith("column"))).toEqual([]);
    expect(bad.length).toBeLessThanOrEqual(data.limits.length * 0.03);
  });

  it("scales scatter charts' X and Y axes", () => {
    for (const s of data.axes.scatter as { x: number[]; y: number[]; xAxis: { min: number; max: number; major: number }; yAxis: { min: number; max: number; major: number }; inner: { width: number; height: number } }[]) {
      const y = scaleFor(s.y, s.inner.height, false);
      expect([y.min, y.max, y.major]).toEqual([s.yAxis.min, s.yAxis.max, s.yAxis.major].map((v) => Number(v.toPrecision(12))));
      const x = scaleFor(s.x, s.inner.width, true);
      expect([x.min, x.max]).toEqual([s.xAxis.min, s.xAxis.max].map((v) => Number(v.toPrecision(12))));
    }
  });

  it("scales stacked columns on the sums", () => {
    for (const s of data.axes.stacked as { rows: number[][]; axis: { min: number; max: number; major: number } }[]) {
      const m = newChart("columnStacked", seriesFromBlock(block(s.rows), "columnStacked"), { col: 0, row: 0, x: 0, y: 0 }, { col: 5, row: 15, x: 0, y: 0 });
      const scene = chartScene(m, 360, 216, { measure: calibriMeasure });
      expect([scene.valueScale!.min, scene.valueScale!.max, scene.valueScale!.major]).toEqual([s.axis.min, s.axis.max, s.axis.major]);
    }
  });
});

function block(rows: CellValue[][], sheet = "Tabelle1") {
  return {
    sheet, r0: 0, c0: 0, r1: rows.length - 1, c1: Math.max(...rows.map((r) => r.length)) - 1,
    get: (r: number, c: number) => rows[r]?.[c] ?? null,
  };
}

describe("series from a selection (Excel's rule)", () => {
  for (const c of data.series as { case: string; rows?: CellValue[][]; kind: string; series: string[]; title: string | null; legend: number | null }[]) {
    if (!c.rows) continue;
    it(`${c.case} (${c.kind})`, () => {
      const kind = c.kind as ChartKind;
      const series = seriesFromBlock(block(c.rows!), kind);
      const formulas = series.map((s, i) => `=SERIES(${s.name?.ref ?? ""},${s.cat?.ref ?? ""},${s.val.ref},${i + 1})`);
      expect(formulas).toEqual(c.series);
      const m = newChart(kind, series, { col: 0, row: 0, x: 0, y: 0 }, { col: 1, row: 1, x: 0, y: 0 });
      expect(titleText(m)).toBe(c.title);
      expect(defaultLegend(kind, series.length)).toBe(c.legend === null ? "none" : "b");
    });
  }
});

describe("histogram bins (Excel's automatic bins)", () => {
  for (const h of data.histogram as { values: number[]; bins: string[]; valueLabels: string[] }[]) {
    it(`${h.values.length} values ${Math.min(...h.values)}..${Math.max(...h.values)}`, () => {
      const bins = histogramBins(h.values);
      const labels = bins.map((b, i) => `${i ? "(" : "["}${formatGeneral(b.from)}, ${formatGeneral(b.to)}]`);
      expect(labels).toEqual(h.bins);
      // The count axis of a default-size histogram.
      const m = newChart("histogram", [{ val: { cache: h.values } }], { col: 0, row: 0, x: 0, y: 0 }, { col: 1, row: 1, x: 0, y: 0 });
      const scene = chartScene(m, 360, 216, { measure: calibriMeasure });
      const axis = scene.prims.filter((p) => p.t === "text" && p.part === "axis:val").map((p) => (p as { text: string }).text);
      expect(axis).toEqual(h.valueLabels.filter((l) => /^-?[\d,]+$/.test(l)));
    });
  }
});

describe("linear trendline label (Excel's text)", () => {
  for (const t of data.trend as { x?: number[]; y?: number[]; column?: number[]; label: { text: string }[] }[]) {
    const xs = t.x ?? t.column!.map((_, i) => i + 1), ys = t.y ?? t.column!;
    it(t.label.map((l) => l.text).join(" / "), () => {
      const fit = linearFit(xs, ys)!;
      // Excel's export drops the space after a split-off minus sign ("x -161,92").
      const want = t.label.map((l) => l.text.replace(/x -(\d)/, "x - $1"));
      expect([equationText(fit.slope, fit.intercept), r2Text(fit.r2)]).toEqual(want);
    });
  }
});

describe("default layout (Excel's SVG export of a 360 x 216 pt chart)", () => {
  const lay = data.layout;
  const textsOf = (kind: string) => lay[kind].svg.texts as { text: string; x: number; y: number; size: number }[];

  function sceneOf(kind: ChartKind) {
    const rows = lay[kind].rows as CellValue[][];
    const m = newChart(kind, seriesFromBlock(block(rows), kind), { col: 0, row: 0, x: 0, y: 0 }, { col: 1, row: 1, x: 0, y: 0 });
    return chartScene(m, 360, 216, { measure: calibriMeasure });
  }

  for (const kind of ["column", "columnStacked", "bar", "barStacked", "line", "lineMarkers", "area", "scatter", "pie", "doughnut", "histogram"] as ChartKind[]) {
    it(`${kind}: texts, their baselines and the value axis`, () => {
      const scene = sceneOf(kind);
      const ours = scene.prims.filter((p) => p.t === "text") as { text: string; x: number; y: number; anchor: string; size: number }[];
      const theirs = textsOf(kind);
      // The same texts (Excel's minus sign and order aside).
      expect(ours.map((t) => t.text).sort()).toEqual(theirs.map((t) => t.text).sort());
      for (const t of theirs) {
        const o = ours.find((p) => p.text === t.text && Math.abs(p.y - t.y) < 3);
        expect(o, `${t.text} at baseline ${t.y}`).toBeTruthy();
        // Horizontal position within 3 pt (Excel's x is the text's left edge).
        const w = calibriMeasure(t.text, t.size);
        const left = o!.anchor === "middle" ? o!.x - w / 2 : o!.anchor === "end" ? o!.x - w : o!.x;
        expect(Math.abs(left - t.x), `${t.text}: ours ${left.toFixed(2)}, Excel ${t.x}`).toBeLessThan(3);
      }
    });
  }

  it("the plot area of a column chart lies where Excel's gridlines are", () => {
    const s = sceneOf("column");
    expect(s.plot.y).toBeCloseTo(37.5, 0);
    expect(s.plot.y + s.plot.h).toBeCloseTo(170.5, 0);
    expect(Math.abs(s.plot.x - 19.5)).toBeLessThan(1);
    expect(Math.abs(s.plot.x + s.plot.w - 349.5)).toBeLessThan(1);
  });
});
