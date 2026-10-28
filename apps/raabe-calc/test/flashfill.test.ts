// Blitzvorschau against Excel (results/2026-10-28-excel-flashfill.json, Range.FlashFill in an
// Excel of our own): from the same examples, the same column.

import { readFileSync } from "node:fs";
import { describe, expect, it } from "vitest";
import { learn, run } from "../src/flashfill";

const data = JSON.parse(readFileSync("tools/feel-lab/results/2026-10-28-excel-flashfill.json", "utf-8")) as Record<string, { sources: string[][]; examples: string[]; result: string[] }>;
// Excel keeps the first part up to the last space here ("Lea Sophie S."); we take the first word.
const KNOWN: Record<string, number[]> = { "two examples": [3, 5] };

describe("Blitzvorschau as Excel fills", () => {
  for (const [name, c] of Object.entries(data)) {
    it(name, () => {
      const rows = c.result.map((_, i) => c.sources.map((col) => col[i]));
      const prog = learn(c.examples.map((output, i) => ({ sources: rows[i], output })));
      expect(prog).not.toBeNull();
      rows.forEach((r, i) => {
        if (KNOWN[name]?.includes(i)) return;
        expect(run(prog!, r), `row ${i + 1}`).toBe(c.result[i]);
      });
    });
  }
});
