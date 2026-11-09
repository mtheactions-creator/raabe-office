// The chart settings built after D-114/D-118 (D-155): each against what Excel 365 wrote for the same setting
// (results/2026-11-09-excel-chart-followup.json, probe_excel_chart_followup.py).

import { readFileSync } from "node:fs";
import { describe, expect, it } from "vitest";
import { strFromU8, unzipSync } from "fflate";
import { Sheet } from "@raabe/sheet-engine";
import { newChart, type ChartModel, type Fill } from "@raabe/charts";
import { chartFromXml, chartToXml } from "../../../packages/ooxml/src/xlsx/chart";
import { newWorkbook, readXlsx, writeXlsx } from "@raabe/ooxml";
import { GRADIENT_PRESETS, SHADOW_PRESETS, presetGradient } from "../src/charts/panefill";

const EXCEL = JSON.parse(readFileSync("tools/feel-lab/results/2026-11-09-excel-chart-followup.json", "utf-8"));
const anchor = { col: 0, row: 0, x: 0, y: 0 }, to = { col: 8, row: 15, x: 0, y: 0 };
const chart = (kind: Parameters<typeof newChart>[0] = "column"): ChartModel => newChart(kind, [
  { name: { text: "A" }, cat: { cache: ["Jan", "Feb", "Mrz", "Apr"] }, val: { cache: [4, -2, 7, 5] } },
  { name: { text: "B" }, cat: { cache: ["Jan", "Feb", "Mrz", "Apr"] }, val: { cache: [6, 5, 3, 8] } },
], anchor, to);

/** The chart space's own c:spPr. */
const areaSpPr = (xml: string) => { const i = xml.indexOf("</c:chart>"); return /<c:spPr>[\s\S]*?<\/c:spPr>/.exec(xml.slice(i))?.[0] ?? ""; };
/** A gradFill's stops and shading, normalised: lumMod 100000 alone is no change. */
function gradKey(xml: string) {
  const g = /<a:gradFill[\s\S]*?<\/a:gradFill>/.exec(xml)?.[0] ?? "";
  const stops = [...g.matchAll(/<a:gs pos="(\d+)"><a:schemeClr val="(\w+)"(?:\/>|>([\s\S]*?)<\/a:schemeClr>)<\/a:gs>/g)].map(([, pos, sc, mods]) => {
    const mod = /lumMod val="(\d+)"/.exec(mods ?? "")?.[1] ?? "100000", off = /lumOff val="(\d+)"/.exec(mods ?? "")?.[1] ?? "0";
    return `${pos}:${sc}:${mod}:${off}`;
  });
  const shade = /<a:lin [^>]*\/>|<a:path[\s\S]*?<\/a:path>/.exec(g)?.[0] ?? "";
  return { stops, shade };
}

describe("Voreingestellte Farbverläufe (D-155)", () => {
  it("are Excel's 30 entries and write what Excel writes", () => {
    const rec = EXCEL.pane["Voreingestellte Farbverläufe"] as Record<string, string>;
    expect(GRADIENT_PRESETS.map((p) => p.label)).toEqual(Object.keys(rec));
    for (const p of GRADIENT_PRESETS) {
      const m = chart();
      m.format = { chart: { fill: { type: "gradient", gradient: presetGradient(p.value)! } } };
      const ours = gradKey(areaSpPr(chartToXml(m)));
      expect(ours, p.label).toEqual(gradKey(rec[p.label]));
      // Read back: the same gradient (the pane recognises the preset again).
      const back = chartFromXml(chartToXml(m)).format?.chart?.fill as Fill;
      expect(back.gradient, p.label).toEqual(m.format.chart!.fill!.gradient);
    }
  });
});

describe("Helligkeit of a gradient stop (D-155)", () => {
  it("is lumMod/lumOff on the colour, as Excel writes it", () => {
    const m = chart();
    m.format = { chart: { fill: { type: "gradient", gradient: { kind: "linear", angle: 90, stops: [
      { pos: 0, color: "#FF0000" }, { pos: 50, color: "#00FF00", brightness: 40 }, { pos: 70, color: "#00FF00", brightness: -30 }, { pos: 100, color: "#0000FF" }] } } } };
    const ours = /<a:gsLst>[\s\S]*?<\/a:gsLst>/.exec(areaSpPr(chartToXml(m)))![0];
    const excel = /<a:gsLst>[\s\S]*?<\/a:gsLst>/.exec(EXCEL.com["gradient stop Helligkeit"].xml)![0];
    expect(ours).toBe(excel);
    const back = chartFromXml(chartToXml(m)).format?.chart?.fill?.gradient?.stops;
    expect(back?.map((s) => s.brightness ?? 0)).toEqual([0, 40, -30, 0]);
  });
});

describe("Perspektive shadows (D-155)", () => {
  it("write Excel's a:outerShdw and read back as the same preset", () => {
    const rec = EXCEL.pane.Perspektive as Record<string, string>;
    const persp = SHADOW_PRESETS.filter((p) => p.label.startsWith("Perspektive"));
    expect(persp.map((p) => p.label)).toEqual(Object.keys(rec));
    const attrs = (x: string) => Object.fromEntries([...(/<a:outerShdw ([^>]*)>/.exec(x)?.[1] ?? "").matchAll(/(\w+)="([^"]*)"/g)].map(([, k, v]) => [k, v]));
    for (const p of persp) {
      const m = chart();
      m.format = { chart: { effects: { shadow: p.s } } };
      const xml = chartToXml(m);
      const a = attrs(areaSpPr(xml)), e = attrs(rec[p.label]);
      // Excel leaves out dist="0" and algn="b" (the defaults).
      expect({ ...a, dist: a.dist === "0" ? undefined : a.dist, algn: a.algn === "b" ? undefined : a.algn }, p.label).toEqual({ ...e, dist: e.dist, algn: e.algn });
      expect(chartFromXml(xml).format?.chart?.effects?.shadow, p.label).toEqual(p.s);
    }
  });
});

describe("an axis' Keine Füllung (D-155)", () => {
  it("is written with Excel's hiddenFill and comes back as Keine Füllung; Automatisch stays automatic", () => {
    const m = chart();
    m.format = { catAxis: { fill: { type: "none" } } };
    const xml = chartToXml(m);
    const cat = /<c:catAx>[\s\S]*?<\/c:catAx>/.exec(xml)![0];
    expect(cat).toContain(/<a14:hiddenFill[\s\S]*?<\/a14:hiddenFill>/.exec(EXCEL.com["axis Keine Füllung"].catAx)![0]);
    expect(chartFromXml(xml).format?.catAxis?.fill).toEqual({ type: "none" });
    expect(chartFromXml(chartToXml(chart())).format?.catAxis?.fill).toBeUndefined();
  });
});

describe("Keine Füllung / Keine Linie of title, legend and plot area (D-155)", () => {
  it("are written with Excel's hiddenFill / hiddenLine and come back", () => {
    const m = chart();
    m.format = { title: { fill: { type: "none" }, line: { type: "none" } }, legend: { line: { type: "none" } }, plot: { fill: { type: "none" } } };
    const xml = chartToXml(m);
    const title = /<c:title>[\s\S]*?<\/c:title>/.exec(xml)![0];
    expect(/<a:extLst>[\s\S]*?<\/a:extLst>/.exec(title)![0]).toBe(/<a:extLst>[\s\S]*?<\/a:extLst>/.exec(EXCEL.none.title["Keine Linie"])![0]);
    const back = chartFromXml(xml).format;
    expect(back?.title?.fill).toEqual({ type: "none" });
    expect(back?.title?.line).toEqual({ type: "none" });
    expect(back?.legend?.line).toEqual({ type: "none" });
    expect(back?.plot?.fill).toEqual({ type: "none" });
  });
  it("a column series' own border colour does not become its fill colour", () => {
    const m = chart();
    m.series[0].line = { type: "solid", color: "#595959", width: 0.75 };
    const back = chartFromXml(chartToXml(m));
    expect(back.series[0].color).toBeUndefined();
    expect(back.series[0].line?.color).toBe("#595959");
  });
});

describe("3D-Format of an axis (D-155)", () => {
  it("is saved and read back", () => {
    const m = chart();
    m.format = { valAxis: { line: { type: "solid", color: "#FF0000", width: 3 }, effects: { format3d: { top: { type: "circle", w: 6, h: 6 } } } } };
    const back = chartFromXml(chartToXml(m));
    expect(back.format?.valAxis?.effects?.format3d?.top).toEqual({ type: "circle", w: 6, h: 6 });
  });
});

describe("the chart style part (D-155)", () => {
  it("carries the chosen Diagrammformatvorlage as its id, as Excel's does", () => {
    const wb = newWorkbook();
    const s = wb.sheets[0] as Sheet & { charts?: ChartModel[] };
    const m = chart();
    m.style = 205;
    s.charts = [m];
    const files = unzipSync(writeXlsx(wb));
    const style = Object.keys(files).find((k) => /^xl\/charts\/style\d+\.xml$/.test(k));
    expect(style).toBeDefined();
    const head = /<cs:chartStyle\b[^>]*>/.exec(strFromU8(files[style!]))![0];
    expect(/id="(\d+)"/.exec(head)![1]).toBe(/id="(\d+)"/.exec(EXCEL.com["ChartStyle 205"]["cs:chartStyle"])![1]);
    expect(strFromU8(files["[Content_Types].xml"])).toContain("chartstyle+xml");
    // Saved again with another style: the same part, the new id.
    const back = readXlsx(writeXlsx(wb));
    const c2 = (back.sheets[0] as Sheet & { charts?: ChartModel[] }).charts![0];
    c2.style = 207;
    const files2 = unzipSync(writeXlsx(back));
    const parts = Object.keys(files2).filter((k) => /^xl\/charts\/style\d+\.xml$/.test(k));
    expect(parts.length).toBe(1);
    expect(strFromU8(files2[parts[0]])).toContain('id="207"');
  });
});

describe("Benutzerdefiniert error amounts with one value (D-155)", () => {
  it("are written as Excel's numLit", () => {
    const m = chart();
    m.series[0].errorBars = { dir: "plus", type: "cust", plus: { cache: [1, 2, 3, 4] }, minus: { ref: "Tabelle1!$C$2:$C$5", cache: [6, 5, 3, 8] } };
    const xml = chartToXml(m);
    const ours = /<c:errBars>[\s\S]*?<\/c:errBars>/.exec(xml)![0];
    const excel = EXCEL.com["error bars custom array"] as string;
    const part = (x: string, tag: string) => new RegExp(`<c:${tag}>[\\s\\S]*?</c:${tag}>`).exec(x)![0];
    expect(part(ours, "plus")).toBe(part(excel, "plus"));
    expect(part(ours, "minus")).toBe(part(excel, "minus"));
  });
});
