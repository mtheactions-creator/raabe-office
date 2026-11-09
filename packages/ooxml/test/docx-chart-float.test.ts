// Floating charts in RaabeWrite (D-131): a chart can leave the line like a picture; it is written as Word writes a
// floating chart (wp:anchor around the chart's graphic frame, the chart part and its embedded workbook) and read back
// with its place, wrapping and data. Charts read from Word keep their embedded workbook (chart.workbook).

import { describe, expect, it } from "vitest";
import { readFileSync } from "node:fs";
import { join } from "node:path";
import { strFromU8, unzipSync } from "fflate";
import { Doc, makePara, type FloatObject } from "@raabe/text-engine";
import { newChart } from "@raabe/charts";
import { chartToXml, newWorkbook, readDocx, writeDocx, writeXlsx } from "../src";

const part = (bytes: Uint8Array, path: string) => strFromU8(unzipSync(bytes)[path]);

function sampleChart() {
  const series = [1, 2].map((i) => ({
    name: { ref: `Tabelle1!$${String.fromCharCode(65 + i)}$1`, text: `Datenreihe ${i}` },
    cat: { ref: "Tabelle1!$A$2:$A$3", cache: ["Kategorie 1", "Kategorie 2"] },
    val: { ref: `Tabelle1!$${String.fromCharCode(65 + i)}$2:$${String.fromCharCode(65 + i)}$3`, cache: [i, i + 1] },
  }));
  const a = { col: 0, row: 0, x: 0, y: 0 };
  return newChart("column", series, a, a);
}

const WRAPS = [
  ["square", '<wp:wrapSquare wrapText="bothSides"/>'], ["topAndBottom", "<wp:wrapTopAndBottom/>"], ["behind", "<wp:wrapNone/>"],
  ["front", "<wp:wrapNone/>"], ["tight", "<wp:wrapTight "], ["through", "<wp:wrapThrough "],
] as const;

describe("floating charts", () => {
  it.each(WRAPS)("%s: written as wp:anchor with the chart and its workbook, read back at the same place", (wrap, el) => {
    const m = sampleChart();
    m.title.text = "Schwebend";
    const workbook = writeXlsx(newWorkbook());
    const f: FloatObject = {
      at: 0, data: undefined, width: 300, height: 180, x: { rel: "column", offset: 72 }, y: { rel: "paragraph", offset: 36 }, wrap, side: "bothSides",
      dist: [0, 0, 9, 9], chart: { xml: chartToXml(m), workbook },
    };
    const d = new Doc({ paras: [{ ...makePara("Absatz mit Diagramm"), floats: [f] }] });
    const out = writeDocx(d.state(), null);
    const xml = part(out, "word/document.xml");
    expect(xml).toContain("<wp:anchor ");
    expect(xml).toContain(el);
    expect(xml).toContain(`behindDoc="${wrap === "behind" ? 1 : 0}"`);
    expect(xml).toContain('<wp:positionH relativeFrom="column"><wp:posOffset>914400</wp:posOffset></wp:positionH>');
    expect(xml).toContain('<wp:positionV relativeFrom="paragraph"><wp:posOffset>457200</wp:posOffset></wp:positionV>');
    expect(xml).toMatch(/<c:chart [^>]*r:id="rId\d+"\/>/);
    const files = unzipSync(out);
    expect(Object.keys(files).some((k) => k.startsWith("word/charts/chart"))).toBe(true);
    expect(Object.keys(files).some((k) => k.startsWith("word/embeddings/Microsoft_Excel_Worksheet"))).toBe(true);
    const back = readDocx(out).state.paras[0].floats![0];
    expect([back.wrap, back.x.rel, back.x.offset, back.y.rel, back.y.offset, back.width, back.height]).toEqual([wrap, "column", 72, "paragraph", 36, 300, 180]);
    expect(back.chart?.xml).toContain("Schwebend");
    expect(back.chart?.workbook?.length).toBe(workbook.length);
  });

  it("a filtered series' extension stays in the chart; the workbook link goes after it (Word refused the file before)", () => {
    const m = sampleChart();
    m.series[1].filtered = true;
    const f: FloatObject = { at: 0, data: undefined, width: 300, height: 180, x: { rel: "column", offset: 0 }, y: { rel: "paragraph", offset: 0 }, wrap: "square", side: "bothSides",
      dist: [0, 0, 9, 9], chart: { xml: chartToXml(m), workbook: writeXlsx(newWorkbook()) } };
    const out = writeDocx(new Doc({ paras: [{ ...makePara("A"), floats: [f] }] }).state(), null);
    const name = Object.keys(unzipSync(out)).find((k) => /^word\/charts\/chart\d+\.xml$/.test(k))!;
    const xml = part(out, name);
    expect(xml).toContain("c15:filteredBarSeries");
    expect(xml.indexOf("<c:externalData")).toBeGreaterThan(xml.lastIndexOf("</c:chart>"));
  });

  it("an inline chart read from Word keeps its embedded workbook", () => {
    const bytes = new Uint8Array(readFileSync(join(__dirname, "fixtures", "word-charts.docx")));
    const paras = readDocx(bytes).state.paras;
    const charts = paras.flatMap((p) => (p.objects ?? []).filter((o) => o.chart));
    expect(charts.length).toBe(7);
    for (const c of charts) expect(c.chart!.workbook!.length).toBeGreaterThan(1000);
  });
});
