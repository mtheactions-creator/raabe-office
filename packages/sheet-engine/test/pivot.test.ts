// PivotTables against Excel's own reports (tools/feel-lab/probe_excel_pivot.py, D-066).
import { readFileSync, readdirSync } from "node:fs";
import { describe, expect, it } from "vitest";
import { computePivot, type PivotFunc, type PivotTable } from "../src/pivot";

const dir = "tools/feel-lab/results";
const file = readdirSync(dir).filter((f) => f.endsWith("-excel-pivot.json")).sort().at(-1)!;
const rec = JSON.parse(readFileSync(`${dir}/${file}`, "utf-8")) as {
  data: (string | number)[][];
  configs: Record<string, { cfg: { rows?: string[]; cols?: string[]; values: [string, PivotFunc][]; filter?: [string, string]; layout?: number }; top: number; left: number; cells: (string | number | null)[][] }>;
};

describe("PivotTables as Excel reports them", () => {
  const names = rec.data[0] as string[];
  const idx = (n: string) => names.indexOf(n);
  for (const [name, c] of Object.entries(rec.configs)) {
    it(name, () => {
      const filterRows = c.cfg.filter ? 2 : 0;
      const p: PivotTable = {
        name, source: { sheet: "Daten", r0: 0, c0: 0, r1: rec.data.length - 1, c1: names.length - 1 },
        at: { row: 2, col: 0 },
        rows: (c.cfg.rows ?? []).map(idx), cols: (c.cfg.cols ?? []).map(idx),
        values: c.cfg.values.map(([f, func]) => ({ field: idx(f), func })),
        filters: c.cfg.filter ? [{ field: idx(c.cfg.filter[0]), items: [c.cfg.filter[1]] }] : [],
        layout: c.cfg.layout === 1 ? "tabular" : c.cfg.layout === 2 ? "outline" : "compact",
      };
      const out = computePivot(p, rec.data);
      // Our report as a grid from the same corner Excel's TableRange2 starts at.
      const top = 2 - filterRows;
      const grid = c.cells.map((row) => row.map(() => null as string | number | null));
      for (const cell of out.cells) {
        const r = cell.row - top, k = cell.col;
        if (r >= 0 && r < grid.length && k < grid[0].length) grid[r][k] = typeof cell.v === "boolean" ? String(cell.v) : cell.v;
      }
      const norm = (v: unknown) => (typeof v === "number" ? Math.round(v * 1e9) / 1e9 : v);
      expect(grid.map((r) => r.map(norm))).toEqual(c.cells.map((r) => r.map(norm)));
      expect(out.cols).toBe(c.cells[0].length);
    });
  }
});
