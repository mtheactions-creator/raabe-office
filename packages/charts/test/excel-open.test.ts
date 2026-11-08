// Workbooks with every chart kind as RaabeCalc writes them (packages/ooxml/src/xlsx/chart.ts), for
// opening in Excel: tools/feel-lab/check_excel_charts.py runs this file with RAABE_CHARTS_XLSX set
// to a folder in TEMP and opens each workbook in an invisible Excel of its own. Without the
// variable nothing is written.

import { mkdirSync, writeFileSync } from "node:fs";
import { describe, expect, it } from "vitest";
import { strToU8, zipSync } from "fflate";
import { KINDS, isChartExKind, newChart, seriesFromBlock, type ChartKind, type ChartModel, type CellValue } from "../src/index";
import {
  CHART_COLORS_CONTENT_TYPE, CHART_CONTENT_TYPE, CHART_REL_TYPE, CHART_STYLE_CONTENT_TYPE, CHARTEX_CONTENT_TYPE, CHARTEX_REL_TYPE,
  DRAWING_CONTENT_TYPE, DRAWING_REL_TYPE, anchorXml, chartColorsXml, chartRelsXml, chartStyleXml, chartToXml, drawingXml, isChartEx,
} from "../../ooxml/src/xlsx/chart";

const KINDS_OLD: ChartKind[] = ["column", "columnStacked", "bar", "barStacked", "line", "lineMarkers", "area", "pie", "doughnut", "scatter", "scatterSmooth", "scatterLines", "histogram"];
const HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n';
const RELS = "http://schemas.openxmlformats.org/package/2006/relationships";
const dir = process.env.RAABE_CHARTS_XLSX;

/** A minimal workbook: one sheet with the data and the charts in a drawing. */
function workbook(rows: CellValue[][], charts: ChartModel[]): Uint8Array {
  const files: Record<string, Uint8Array> = {};
  const types: string[] = [], rels: string[] = [], anchors: string[] = [];
  charts.forEach((m, i) => {
    const n = i + 1, ex = isChartEx(m), part = ex ? `chartEx${n}.xml` : `chart${n}.xml`;
    files[`xl/charts/${part}`] = strToU8(chartToXml(m));
    files[`xl/charts/style${n}.xml`] = strToU8(chartStyleXml());
    files[`xl/charts/colors${n}.xml`] = strToU8(chartColorsXml());
    files[`xl/charts/_rels/${part}.rels`] = strToU8(chartRelsXml(`style${n}.xml`, `colors${n}.xml`));
    types.push(`<Override PartName="/xl/charts/${part}" ContentType="${ex ? CHARTEX_CONTENT_TYPE : CHART_CONTENT_TYPE}"/>`,
      `<Override PartName="/xl/charts/style${n}.xml" ContentType="${CHART_STYLE_CONTENT_TYPE}"/>`,
      `<Override PartName="/xl/charts/colors${n}.xml" ContentType="${CHART_COLORS_CONTENT_TYPE}"/>`);
    rels.push(`<Relationship Id="rId${n}" Type="${ex ? CHARTEX_REL_TYPE : CHART_REL_TYPE}" Target="../charts/${part}"/>`);
    anchors.push(anchorXml({ from: m.from, to: m.to, id: n + 1, name: m.name ?? `Diagramm ${n}`, rId: `rId${n}`, chartEx: ex }));
  });
  const cells = rows.map((r, ri) => `<row r="${ri + 1}">${r.map((v, ci) => {
    const ref = String.fromCharCode(65 + ci) + (ri + 1);
    if (v === null) return "";
    return typeof v === "number" ? `<c r="${ref}"><v>${v}</v></c>` : `<c r="${ref}" t="inlineStr"><is><t>${v}</t></is></c>`;
  }).join("")}</row>`).join("");
  files["xl/drawings/drawing1.xml"] = strToU8(drawingXml(anchors));
  files["xl/drawings/_rels/drawing1.xml.rels"] = strToU8(`${HEAD}<Relationships xmlns="${RELS}">${rels.join("")}</Relationships>`);
  files["[Content_Types].xml"] = strToU8(`${HEAD}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/drawings/drawing1.xml" ContentType="${DRAWING_CONTENT_TYPE}"/>${types.join("")}</Types>`);
  files["_rels/.rels"] = strToU8(`${HEAD}<Relationships xmlns="${RELS}"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`);
  files["xl/workbook.xml"] = strToU8(`${HEAD}<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Tabelle1" sheetId="1" r:id="rId1"/></sheets></workbook>`);
  files["xl/_rels/workbook.xml.rels"] = strToU8(`${HEAD}<Relationships xmlns="${RELS}"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>`);
  files["xl/worksheets/sheet1.xml"] = strToU8(`${HEAD}<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetData>${cells}</sheetData><drawing r:id="rId1"/></worksheet>`);
  files["xl/worksheets/_rels/sheet1.xml.rels"] = strToU8(`${HEAD}<Relationships xmlns="${RELS}"><Relationship Id="rId1" Type="${DRAWING_REL_TYPE}" Target="../drawings/drawing1.xml"/></Relationships>`);
  return zipSync(files);
}

describe.skipIf(!dir)("workbooks for Excel", () => {
  it("writes one workbook of plain charts and one with every element", () => {
    mkdirSync(dir!, { recursive: true });
    const rows: CellValue[][] = [[null, "Reihe A", "Reihe B"], ["Jan", 4.3, 2.4], ["Feb", 2.5, 4.4], ["Mär", 3.5, 1.8], ["Apr", 4.5, 2.8]];
    const block = { sheet: "Tabelle1", r0: 0, c0: 0, r1: 4, c1: 2, get: (r: number, c: number) => rows[r]?.[c] ?? null };
    const expected: Record<string, { kind: ChartKind; xl?: number; series: string[] }[]> = {};
    for (const variant of ["plain", "elements"]) {
      const charts = KINDS_OLD.map((kind, i) => {
        // Three across, 13 rows each: Excel exports charts outside its (invisible) window as empty pictures.
        const col = 4 + (i % 3) * 7, row = 1 + Math.floor(i / 3) * 13;
        const m = newChart(kind, seriesFromBlock(block, kind), { col, row, x: 3, y: 2 }, { col: col + 6, row: row + 12, x: 0, y: 0 });
        m.name = `Diagramm ${i + 1}`;
        if (variant === "elements") {
          m.title = { show: true, text: "Mein Titel" };
          if (kind !== "pie" && kind !== "doughnut" && kind !== "histogram") m.axisTitles = { cat: "Monat", val: "Wert" };
          m.legend = kind === "histogram" ? "none" : "r";
          m.dataLabels = true;
          if (!["pie", "doughnut", "histogram"].includes(kind)) m.series[0].trendline = { type: "linear", showEq: true, showR2: true };
        }
        return m;
      });
      writeFileSync(`${dir}/charts-${variant}.xlsx`, workbook(rows, charts));
      expected[`charts-${variant}.xlsx`] = charts.map((m) => ({
        kind: m.kind,
        series: m.kind === "histogram" ? [] : m.series.map((s, i) => `=SERIES(${s.name?.ref ?? ""},${s.cat?.ref ?? ""},${s.val.ref},${i + 1})`),
      }));
    }
    // Every other type of Excel 365 (D-086), each from the block it needs, and the Verbund presets.
    const typeRows: CellValue[][] = [[null, "Reihe A", "Reihe B", "Reihe C", "Reihe D", "Reihe E"], ["Jan", 43, 24, 20, 30, 1200], ["Feb", 25, 44, 20, 32, 900],
      ["Mär", 35, 18, 30, 34, 1500], ["Apr", 45, 28, 50, 31, 1100], ["Mai", 32, 36, 40, 33, 800]];
    const cols: Partial<Record<ChartKind, [number, number]>> = { stockHLC: [0, 3], stockOHLC: [0, 4], stockVHLC: [0, 4], stockVOHLC: [0, 5], bubble: [1, 3], bubble3D: [1, 3],
      waterfall: [0, 1], funnel: [0, 1], pareto: [0, 1], treemap: [0, 1], sunburst: [0, 1], pieOfPie: [0, 1], barOfPie: [0, 1], pie3D: [0, 1] };
    const fresh = (Object.keys(KINDS) as ChartKind[]).filter((k) => !KINDS_OLD.includes(k));
    const list: { kind: ChartKind; combo?: "line" | "lineSecondary" }[] = [...fresh.map((kind) => ({ kind })), { kind: "column", combo: "line" }, { kind: "column", combo: "lineSecondary" }];
    const typeCharts = list.map(({ kind, combo }, i) => {
      const [c0, c1] = cols[kind] ?? [0, 3];
      const b = { sheet: "Tabelle1", r0: 0, c0, r1: 5, c1, get: (r: number, c: number) => typeRows[r]?.[c] ?? null };
      const col = 7 + (i % 3) * 7, row = 1 + Math.floor(i / 3) * 13;
      const m = newChart(kind, seriesFromBlock(b, kind), { col, row, x: 3, y: 2 }, { col: col + 6, row: row + 12, x: 0, y: 0 });
      m.name = `Diagramm ${i + 1}`;
      if (combo) { m.series[2].kind = "line"; if (combo === "lineSecondary") m.series[2].secondary = true; }
      return m;
    });
    writeFileSync(`${dir}/charts-types.xlsx`, workbook(typeRows, typeCharts));
    expected["charts-types.xlsx"] = typeCharts.map((m) => ({
      kind: m.kind, xl: m.series.some((s) => s.kind) ? -4111 : KINDS[m.kind].xl,
      series: isChartExKind(m.kind) ? [] : m.series.map((s, i) => `=SERIES(${s.name?.ref ?? ""},${s.cat?.ref ?? ""},${s.val.ref},${i + 1}${s.size?.ref ? `,${s.size.ref}` : ""})`),
    }));
    writeFileSync(`${dir}/expected.json`, JSON.stringify(expected, null, 1));
    expect(Object.keys(expected)).toHaveLength(3);
  });
});
