// Step 12: the automatic Grenzen and Einheiten are the ones Excel uses (results/2026-11-12-excel-axis-auto.json,
// probe_excel_axis_auto.py: 10 data sets x 9 chart types x 3 sizes, read from Excel through COM). The Format pane shows
// these values while an axis is automatic (D-190 to D-199 range is the designs'; this is D-200).

import { readFileSync } from "node:fs";
import { resolve } from "node:path";
import { describe, expect, it } from "vitest";
import { chartScene, newChart, type ChartKind } from "../src/index";

interface AxisRec { MinimumScale: number | null; MaximumScale: number | null; MajorUnit: number | null; MinorUnit: number | null }
interface Case { data: string; kind: ChartKind; size: [number, number]; val: AxisRec; cat?: AxisRec; plot: number[]; legend: string; title: boolean }
const rec = JSON.parse(readFileSync(resolve(process.cwd(), "tools/feel-lab/results/2026-11-12-excel-axis-auto.json"), "utf8")) as
  { x: number[]; data: Record<string, number[]>; cases: Case[] };

const anchor = { col: 0, row: 0, x: 0, y: 0 }, to = { col: 8, row: 15, x: 0, y: 0 };

function model(c: Case) {
  const vals = rec.data[c.data];
  const scatter = c.kind.startsWith("scatter");
  const s = { name: { text: "Werte" }, cat: { cache: scatter ? rec.x : ["1", "2", "3", "4"] }, val: { cache: vals } };
  // Title and legend as Excel's chart had them (a new category chart from one column: title and a legend at the bottom).
  const m = newChart(c.kind, [s], anchor, to);
  m.legend = c.legend as typeof m.legend;
  m.title = { ...m.title, show: c.title };
  return m;
}

const near = (a: number | undefined, b: number | null) => b === null || (a !== undefined && Math.abs(a - b) <= Math.max(1e-9, Math.abs(b) * 1e-6));

describe("automatic axis values as Excel's", () => {
  const cases = rec.cases.filter((c) => c.val.MaximumScale !== null);
  const results = cases.map((c) => {
    const sc = chartScene(model(c), c.size[0], c.size[1]);
    const v = sc.valueScale;
    const ok = !!v && near(v.min, c.val.MinimumScale) && near(v.max, c.val.MaximumScale) && near(v.major, c.val.MajorUnit) && near(v.major / 5, c.val.MinorUnit);
    let okX = true;
    if (c.kind.startsWith("scatter") && c.cat?.MaximumScale !== null && c.cat) {
      const x = sc.xScale;
      okX = !!x && near(x.min, c.cat.MinimumScale) && near(x.max, c.cat.MaximumScale) && near(x.major, c.cat.MajorUnit);
    }
    return { c, ok: ok && okX, ours: v ? [v.min, v.max, v.major] : null, x: sc.xScale ? [sc.xScale.min, sc.xScale.max, sc.xScale.major] : null };
  });

  it("matches Excel's Minimum, Maximum, Hauptstriche and Teilstriche in every recorded case", () => {
    const bad = results.filter((r) => !r.ok).map((r) => `${r.c.kind} ${r.c.data} ${r.c.size.join("x")}: ours ${JSON.stringify(r.ours)} x ${JSON.stringify(r.x)}, Excel ${r.c.val.MinimumScale}/${r.c.val.MaximumScale}/${r.c.val.MajorUnit}` +
      (r.c.cat?.MaximumScale != null ? ` x ${r.c.cat.MinimumScale}/${r.c.cat.MaximumScale}/${r.c.cat.MajorUnit}` : ""));
    expect(bad, `${bad.length} of ${results.length} differ:\n${bad.join("\n")}`).toEqual([]);
  });
});
