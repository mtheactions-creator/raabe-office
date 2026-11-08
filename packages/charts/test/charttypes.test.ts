// Every chart type of Excel 365 (D-086): Excel's own workbook (tools/feel-lab/probe_excel_charttypes.py,
// one chart per sheet named by its type) is read into the model, each chart's type, series and
// Excel's recorded COM properties are checked, untouched charts write back verbatim, and every type
// goes through write → read with the same model. With RAABE_CHARTTYPES_SVG set to a folder, each
// chart is also drawn there as HTML for the visual comparison (compare_calc_charttypes.py).

import { readFileSync, mkdirSync, writeFileSync } from "node:fs";
import { describe, expect, it } from "vitest";
import { readXlsx } from "../../ooxml/src/xlsx/read";
import { chartFromXml, chartToXml, isChartEx } from "../../ooxml/src/xlsx/chart";
import { formatValue, type Workbook } from "../../sheet-engine/src/index";
import {
  KINDS, calibriMeasure, chartScene, isChartExKind, kindInfo, newChart, parseRef, refreshCaches, sceneMarkup, seriesFromBlock,
  type CellValue, type ChartKind, type ChartModel,
} from "../src/index";

const FIX = "packages/ooxml/test/fixtures";
const recorded = JSON.parse(readFileSync(`${FIX}/excel-charttypes.json`, "utf-8")) as { types: Record<string, { xlType: number; chartType: number; series: { formula: string | null; type: number; axisGroup: number }[] }> };
const wb: Workbook = readXlsx(readFileSync(`${FIX}/excel-charttypes.xlsx`));

function charts(name: string): ChartModel[] {
  const sheet = wb.sheets.find((s) => s.name === name);
  return ((sheet as unknown as { charts?: ChartModel[] })?.charts) ?? [];
}

/** The workbook's cells for a chart (RaabeCalc's readCaches, without hidden rows). */
function refresh(m: ChartModel): void {
  refreshCaches(m, (ref0) => {
    const ref = ref0 && !parseRef(ref0) ? wb.names.find((n) => n.name === ref0)?.ref ?? ref0 : ref0;
    const r = parseRef(ref);
    const sheet = r ? wb.sheets.find((s) => s.name === r.sheet) : undefined;
    if (!r || !sheet) return null;
    const vals: CellValue[] = [];
    let format: string | undefined;
    for (let row = r.r0; row <= r.r1; row++) for (let col = r.c0; col <= r.c1; col++) {
      const cell = sheet.get(row, col);
      if (format === undefined && cell) format = wb.styles.get(cell.s).numFmt;
      vals.push((cell?.v ?? null) as CellValue);
    }
    return { vals, format, cols: r.c1 - r.c0 + 1 };
  }, (v, f) => (typeof v === "number" ? formatValue(v, f ?? "General").text : String(v)));
}

/** The combos are recorded under their own names; their model kind is the first group's. */
const COMBO: Record<string, { kind: ChartKind; kinds: (ChartKind | undefined)[]; secondary: boolean[] }> = {
  comboColumnLine: { kind: "column", kinds: [undefined, undefined, "line"], secondary: [false, false, false] },
  comboColumnLineSecondary: { kind: "column", kinds: [undefined, undefined, "line"], secondary: [false, false, true] },
  comboAreaColumn: { kind: "areaStacked", kinds: [undefined, undefined, "column"], secondary: [false, false, false] },
};

describe("Excel's chart types read", () => {
  for (const [name, rec] of Object.entries(recorded.types)) {
    it(name, () => {
      const [m] = charts(name);
      expect(m, `chart on sheet ${name}`).toBeDefined();
      const combo = COMBO[name];
      expect(m.kind).toBe(combo ? combo.kind : name);
      if (!combo) expect(KINDS[m.kind].xl).toBe(rec.chartType);
      expect(isChartEx(m)).toBe(isChartExKind(m.kind));
      if (combo) {
        expect(m.series.map((s) => s.kind)).toEqual(combo.kinds);
        expect(m.series.map((s) => !!s.secondary)).toEqual(combo.secondary);
      }
      // Series as Excel's COM lists them (chartex charts have no SERIES formulas).
      const formulas = rec.series.map((s) => s.formula).filter((f): f is string => !!f);
      if (formulas.length) {
        const ours = m.series.map((s, i) => `=SERIES(${s.name?.ref ?? ""},${s.cat?.ref ?? ""},${s.val.ref},${i + 1}${s.size?.ref ? `,${s.size.ref}` : ""})`);
        expect(ours).toEqual(formulas);
      } else {
        expect(m.series.length).toBe(name === "boxWhisker" ? 3 : 1);
      }
      // Untouched: written back as Excel wrote it.
      const xml = (m.raw as { xml: string }).xml;
      expect(chartToXml(m)).toBe(xml);
    });
  }

  it("reads the hierarchy, totals and bubble sizes from the cells", () => {
    const [tm] = charts("treemap");
    refresh(tm);
    expect(tm.series[0].levels).toEqual([["Obst", "Obst", "Obst", "Gemüse", "Gemüse", "Brot", "Brot"], ["Äpfel", "Birnen", "Kirschen", "Möhren", "Kohl", "Vollkorn", "Weißbrot"]]);
    expect(tm.series[0].val.cache).toEqual([30, 20, 10, 25, 15, 18, 12]);
    const [wf] = charts("waterfall");
    expect(wf.series[0].totals).toEqual([5]);
    const [bu] = charts("bubble");
    expect(bu.series[0].size?.cache).toEqual([10, 4, 8, 2, 6]);
    expect(charts("bubble3D")[0].kind).toBe("bubble3D");
    expect(charts("pieOfPie")[0].splitCount).toBeUndefined();
  });
});

describe("write → read keeps every new type", () => {
  const rows: CellValue[][] = [[null, "Reihe A", "Reihe B", "Reihe C"], ["Jan", 4.3, 2.4, 2], ["Feb", 2.5, 4.4, 2], ["Mär", 3.5, 1.8, 3], ["Apr", 4.5, 2.8, 5]];
  const block = { sheet: "Daten", r0: 0, c0: 0, r1: 4, c1: 3, get: (r: number, c: number) => rows[r]?.[c] ?? null };
  const strip = (m: ChartModel) => { const { raw: _r, ...rest } = m; void _r; return JSON.parse(JSON.stringify(rest)); };
  for (const kind of Object.keys(KINDS) as ChartKind[]) {
    it(kind, () => {
      const m = newChart(kind, seriesFromBlock(block, kind), { col: 5, row: 1, x: 0, y: 0 }, { col: 12, row: 16, x: 0, y: 0 });
      m.title = { show: true, text: "Titel" };
      for (const s of m.series) {
        s.val.format ??= "General";
        if (s.size) s.size.format ??= "General";
        if (s.cat && !s.cat.cache.some((v) => typeof v === "string")) s.cat.format ??= "General";
      }
      const xml = chartToXml(m);
      const back = chartFromXml(xml, { from: m.from, to: m.to });
      expect(back.kind).toBe(kind);
      if (isChartExKind(kind)) {
        expect(back.series.map((s) => s.val.cache)).toEqual(m.series.map((s) => s.val.cache));
        return;
      }
      const want = strip(m), got = strip(back);
      for (const s of [...want.series, ...got.series]) { delete s.color; delete s.marker; delete s.size?.format; }
      if (kindInfo(kind).family === "bubble") for (const s of want.series) if (s.cat) delete s.cat.format;
      expect(got.series.map((s: { val: unknown }) => s.val)).toEqual(want.series.map((s: { val: unknown }) => s.val));
      expect(got.legend).toBe(want.legend);
      expect(got.title).toEqual(want.title);
    });
  }

  it("a Verbund chart keeps its series' types and the secondary axis", () => {
    const m = newChart("column", seriesFromBlock(block, "column"), { col: 5, row: 1, x: 0, y: 0 }, { col: 12, row: 16, x: 0, y: 0 });
    m.series[2].kind = "line";
    m.series[2].secondary = true;
    const back = chartFromXml(chartToXml(m));
    expect(back.kind).toBe("column");
    expect(back.series.map((s) => [s.kind, !!s.secondary])).toEqual([[undefined, false], [undefined, false], ["line", true]]);
  });
});

describe("drawing", () => {
  const dir = process.env.RAABE_CHARTTYPES_SVG;
  const fmt = (v: number, code?: string) => formatValue(v, code ?? "General").text;
  it("draws every recorded type without errors", () => {
    if (dir) mkdirSync(dir, { recursive: true });
    for (const name of Object.keys(recorded.types)) {
      const [m] = charts(name);
      refresh(m);
      const scene = chartScene(m, 360, 216, { measure: calibriMeasure, format: fmt });
      expect(scene.prims.length).toBeGreaterThan(5);
      if (dir) {
        const svg = `<svg xmlns="http://www.w3.org/2000/svg" width="722" height="434" viewBox="0 0 361 217" font-family="Calibri, Carlito, 'Segoe UI', sans-serif">${sceneMarkup(scene)}</svg>`;
        writeFileSync(`${dir}/${name}.html`, `<!doctype html><html><head><meta charset="utf-8"><style>html,body{margin:0;background:#fff;overflow:hidden}</style></head><body>${svg}</body></html>`);
      }
    }
  });
});
