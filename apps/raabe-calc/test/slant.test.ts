// Turned text slants its cell (D-125): Excel's PDF of cells with borders, a fill and text turned by -90 to 90
// degrees (probe_excel_rotated_borders.py) gives, for every angle, how far the top of the cell moves sideways
// against the cell's height; slantOf must give the same, and nothing where Excel keeps the rectangle.

import { describe, expect, it } from "vitest";
import recJson from "../../../tools/feel-lab/results/2026-11-09-excel-rotated-borders.json";
import { slantOf } from "../src/grid/render";

type Path = { paint: string; fill: string | null; d: (string | number)[][] };
const rec = recJson as unknown as {
  pages: { paths: Path[] }[]; layout: { case: string; angle: number; edges?: number[] }[];
};

/** Excel's fill parallelograms (filled paths of the cell colour with four corners), in page order (top to bottom). */
function fills(): { shift: number; h: number }[] {
  const out: { y: number; shift: number; h: number }[] = [];
  for (const p of rec.pages[0].paths) {
    if (p.paint !== "fill" || p.fill !== "FFE5CC") continue;
    const pts = p.d.filter((s) => s[0] === "M" || s[0] === "L").map((s) => [s[1] as number, s[2] as number]);
    if (pts.length < 4) continue;
    const ys = pts.map((q) => q[1]);
    const h = Math.max(...ys) - Math.min(...ys);
    if (h < 10) continue; // the thin strips under each cell
    const bottom = pts.filter((q) => q[1] === Math.max(...ys)).map((q) => q[0]);
    const top = pts.filter((q) => q[1] === Math.min(...ys)).map((q) => q[0]);
    out.push({ y: Math.min(...ys), shift: Math.min(...top) - Math.min(...bottom), h });
  }
  // One entry per cell (the PDF paints each fill twice).
  return out.sort((a, b) => a.y - b.y).filter((x, i, all) => i === 0 || Math.abs(x.y - all[i - 1].y) > 0.5 || Math.abs(x.shift - all[i - 1].shift) > 0.5);
}

const style = (rotation: number, sides = true) => ({ rotation, borders: sides ? { left: { style: "thin" }, right: { style: "thin" } } : {} }) as never;
const rot = (deg: number) => (deg >= 0 ? deg : 90 - deg);

describe("turned text slants its cell as Excel's (D-125)", () => {
  it("moves the top by height / tan(angle) for every recorded angle", () => {
    const f = fills();
    const singles = rec.layout.filter((c) => c.case.startsWith("single"));
    expect(f.length).toBeGreaterThanOrEqual(singles.length);
    singles.forEach((c, i) => {
      const want = f[i].shift / f[i].h;
      expect(slantOf(style(rot(c.angle)), 1), c.case).toBeCloseTo(want, 1);
    });
  });
  it("keeps the rectangle without side borders, at 90 degrees and for vertical text", () => {
    expect(slantOf(style(45, false), 40)).toBe(0);
    expect(slantOf(style(90), 40)).toBe(0);
    expect(slantOf(style(180), 40)).toBe(0);
    expect(slantOf(style(255), 40)).toBe(0);
    expect(slantOf(style(0), 40)).toBe(0);
  });
});
