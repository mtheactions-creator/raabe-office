import { readFileSync, writeFileSync } from "node:fs";
import { describe, expect, it } from "vitest";
import { presetGeometry, presetNames } from "../src/geometry";

describe("preset geometry", () => {
  it("evaluates every preset to finite paths", () => {
    const bad: string[] = [];
    for (const n of presetNames()) {
      for (const [w, h] of [[80, 50], [50, 80], [10, 10]]) {
        const g = presetGeometry(n, w, h);
        if (!g || !g.paths.length || g.paths.some((p) => /NaN|Infinity/.test(p.d)) || [g.text.l, g.text.t, g.text.r, g.text.b].some((v) => !Number.isFinite(v))) bad.push(`${n} ${w}x${h}`);
      }
    }
    expect(bad).toEqual([]);
  });

  it("draws a rounded rectangle with quarter arcs", () => {
    const g = presetGeometry("roundRect", 100, 50)!;
    expect(g.paths[0].d.startsWith("M0 8.334 C")).toBe(true);
    expect(g.text.l).toBeCloseTo(8.334 * 0.29289, 2);
  });

  // check_shapes.py: outlines for the shapes it lists (page coordinates added there).
  it.runIf(!!process.env.SHAPES_IN)("writes outlines for check_shapes.py", () => {
    const spec = JSON.parse(readFileSync(process.env.SHAPES_IN!, "utf-8")) as { prst: string; w: number; h: number; adj: Record<string, number> }[];
    const out = spec.map((s) => {
      const g = presetGeometry(s.prst, s.w, s.h, s.adj);
      return { prst: s.prst, paths: g?.paths ?? [], adjust: g?.adjust ?? {}, text: g?.text };
    });
    writeFileSync(process.env.SHAPES_OUT!, JSON.stringify(out));
  });
});
