// Zielwertsuche against Excel's Range.GoalSeek (results/2026-10-28-excel-goalseek.json): the same
// outcome (found or not), within Excel's 0,001 of the goal, at the same solution.

import { readFileSync } from "node:fs";
import { describe, expect, it } from "vitest";
import { seek } from "../src/goalseek";

const rec = JSON.parse(readFileSync("tools/feel-lab/results/2026-10-28-excel-goalseek.json", "utf-8")) as { formula: string; goal: number; start: number; found: boolean; x: number }[];
const F: Record<string, (x: number) => number> = {
  "=3*A1+2": (x) => 3 * x + 2, "=A1^2": (x) => x * x, "=A1^2-4*A1+3": (x) => x * x - 4 * x + 3, "=EXP(A1)": Math.exp,
  "=1000*(1+A1)^10": (x) => 1000 * (1 + x) ** 10, "=SIN(A1)": Math.sin, "=1/A1": (x) => (x === 0 ? NaN : 1 / x), "=A1^3-A1": (x) => x ** 3 - x,
};

describe("Zielwertsuche as Excel's", () => {
  for (const c of rec) {
    it(`${c.formula} = ${c.goal} from ${c.start}`, () => {
      const r = seek(F[c.formula], c.goal, c.start);
      expect(r.found).toBe(c.found);
      if (!c.found) return;
      expect(Math.abs(r.y - c.goal)).toBeLessThanOrEqual(0.001);
      // The same solution: within the tolerance's reach of Excel's value.
      expect(Math.abs(r.x - c.x)).toBeLessThan(0.01 * Math.max(1, Math.abs(c.x)));
    });
  }
});
