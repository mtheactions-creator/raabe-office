// 3-D columns and bars laid out as Excel lays them out (D-125): Excel's own point boxes of the chart fixture
// (probe_excel_3d_geometry.py: each 3-D column's projected box = front face + its depth) give the column's front
// width, the share its depth adds, the category pitch and the value scale; our scene must have the same proportions.

import { describe, expect, it } from "vitest";
import { calibriMeasure, chartScene, newChart, type ChartKind, type ChartModel } from "../src/index";
import geo from "../../../tools/feel-lab/results/2026-11-09-excel-3d-geometry.json";

type Box = { Left: number; Top: number; Width: number; Height: number };
const rec = geo as unknown as Record<string, { points: Box[][]; view: { Rotation: number; Elevation: number } }>;
const V = [[4.3, 2.5, 3.5, 4.5, 3.2], [2.4, 4.4, 1.8, 2.8, 3.6], [2, 2, 3, 5, 4]];
const cats = ["Jan", "Feb", "Mrz", "Apr", "Mai"];

function model(kind: ChartKind): ChartModel {
  const series = V.map((v, i) => ({ name: { text: `Reihe ${"ABC"[i]}` }, cat: { cache: cats }, val: { cache: v } }));
  return newChart(kind, series, { col: 0, row: 0, x: 0, y: 0 }, { col: 7, row: 15, x: 0, y: 0 });
}

/** Bounding boxes of each point's faces in our scene. */
function ours(kind: ChartKind): Box[][] {
  const scene = chartScene(model(kind), 360, 216, { measure: calibriMeasure });
  const out: Box[][] = V.map(() => []);
  for (const p of scene.prims) {
    const m = /^series:(\d+):(\d+)$/.exec((p as { part?: string }).part ?? "");
    if (!m || (p as { t: string }).t !== "path") continue;
    const nums = [...((p as { d: string }).d.matchAll(/(-?[\d.]+) (-?[\d.]+)/g))].map((x) => [Number(x[1]), Number(x[2])]);
    const [si, i] = [Number(m[1]), Number(m[2])];
    const b = out[si][i] ?? { Left: nums[0][0], Top: nums[0][1], Width: 0, Height: 0 };
    const x0 = Math.min(b.Left, ...nums.map((q) => q[0])), y0 = Math.min(b.Top, ...nums.map((q) => q[1]));
    const x1 = Math.max(b.Left + b.Width, ...nums.map((q) => q[0])), y1 = Math.max(b.Top + b.Height, ...nums.map((q) => q[1]));
    out[si][i] = { Left: x0, Top: y0, Width: x1 - x0, Height: y1 - y0 };
  }
  return out;
}

describe("3-D columns and bars in Excel's proportions (D-125)", () => {
  it("3D-Säulen (gruppiert): square columns, the depth adds sin(rotY) across, pitch 4.5 column widths", () => {
    const ex = rec.column3DClustered.points, us = ours("column3DClustered");
    // Excel: series B sits one column width to the right of A, so B.Left - A.Left is the front width.
    const exFront = ex[1][0].Left - ex[0][0].Left, usFront = us[1][0].Left - us[0][0].Left;
    expect((ex[0][0].Width - exFront) / exFront).toBeCloseTo(Math.sin((20 * Math.PI) / 180), 2);
    expect((us[0][0].Width - usFront) / usFront).toBeCloseTo((ex[0][0].Width - exFront) / exFront, 2);
    expect((us[0][1].Left - us[0][0].Left) / usFront).toBeCloseTo((ex[0][1].Left - ex[0][0].Left) / exFront, 2);
    // Front width within 5 % of Excel's (the chart is 360 x 216 pt as Excel's).
    expect(Math.abs(usFront - exFront) / exFront).toBeLessThan(0.05);
  });
  it("3D-Säulen (Reihen in der Tiefe): each series a row of depth one category deep", () => {
    const ex = rec.column3D.points, us = ours("column3D");
    const exRow = ex[1][0].Left - ex[0][0].Left, usRow = us[1][0].Left - us[0][0].Left;
    const exPitch = ex[0][1].Left - ex[0][0].Left, usPitch = us[0][1].Left - us[0][0].Left;
    expect(usRow / usPitch).toBeCloseTo(exRow / exPitch, 2);
    expect(Math.abs(usPitch - exPitch) / exPitch).toBeLessThan(0.05);
  });
  it("3D-Balken: square bars, the depth adds sin(rotX) up", () => {
    const ex = rec.bar3DClustered.points, us = ours("bar3DClustered");
    const exThick = ex[0][0].Top - ex[1][0].Top, usThick = us[0][0].Top - us[1][0].Top;
    expect((us[0][0].Height - usThick) / usThick).toBeCloseTo((ex[0][0].Height - exThick) / exThick, 2);
    expect(Math.abs(usThick - exThick) / exThick).toBeLessThan(0.08);
  });
});
