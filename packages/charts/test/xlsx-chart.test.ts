// Charts in .xlsx (packages/ooxml/src/xlsx/chart.ts): Excel's own chart parts read into the model
// (fixtures saved by Excel 365 through tools/feel-lab), written back unchanged while untouched,
// and every chart kind through write → read with the same model.

import { readFileSync } from "node:fs";
import { describe, expect, it } from "vitest";
import { newChart, seriesFromBlock, type ChartKind, type ChartModel, type CellValue } from "../src/index";
import { anchorFromXml, anchorXml, chartColorsXml, chartFromXml, chartStyleXml, chartToXml, isChartEx } from "../../ooxml/src/xlsx/chart";

const fx = (name: string) => readFileSync(`packages/charts/test/fixtures/${name}`, "utf-8");

describe("Excel's chart parts", () => {
  it("reads a clustered column chart", () => {
    const m = chartFromXml(fx("excel-column.xml"));
    expect(m.kind).toBe("column");
    expect(m.title).toEqual({ show: true });
    expect(m.legend).toBe("b");
    expect(m.gridlines).toBe(true);
    expect(m.series.map((s) => [s.name?.ref, s.name?.text, s.cat?.ref, s.val.ref])).toEqual([
      ["Tabelle1!$B$1", "Reihe A", "Tabelle1!$A$2:$A$4", "Tabelle1!$B$2:$B$4"],
      ["Tabelle1!$C$1", "Reihe B", "Tabelle1!$A$2:$A$4", "Tabelle1!$C$2:$C$4"],
    ]);
    expect(m.series[0].cat?.cache).toEqual(["Jan", "Feb", "Mär"]);
    expect(m.series[1].val.cache).toEqual([2.4, 4.4000000000000004, 1.8]);
    expect(m.series[0].color).toBeUndefined();
  });

  it("reads a scatter chart with a linear trendline", () => {
    const m = chartFromXml(fx("excel-scatter-trend.xml"));
    expect(m.kind).toBe("scatter");
    expect(m.series[0].trendline).toEqual({ type: "linear", showEq: true, showR2: true });
    expect(m.series[0].cat?.cache).toEqual([1, 2, 3]);
    expect(m.legend).toBe("none");
    // One series without a title element but autoTitleDeleted 0: Excel shows a title.
    expect(m.title.show).toBe(true);
  });

  it("reads a pie chart and a histogram (chartex)", () => {
    const pie = chartFromXml(fx("excel-pie.xml"));
    expect(pie.kind).toBe("pie");
    expect(pie.series).toHaveLength(1);
    const hist = chartFromXml(fx("excel-histogram.xml"));
    expect(hist.kind).toBe("histogram");
    expect(hist.series[0].val.ref).toBe("_xlchart.v1.1");
    expect(hist.series[0].name).toEqual({ ref: "_xlchart.v1.0", text: "Punkte" });
    expect(hist.gridlines).toBe(true);
  });

  it("writes an untouched chart back as it was read (D-024), a changed one anew", () => {
    const xml = fx("excel-column.xml");
    const m = chartFromXml(xml);
    m.series[0].val.cache = [9, 9, 9];
    expect(chartToXml(m)).toBe(xml);
    m.legend = "r";
    const out = chartToXml(m);
    expect(out).not.toBe(xml);
    expect(chartFromXml(out).legend).toBe("r");
  });

  it("reads Excel's anchor and writes one back", () => {
    const a = anchorFromXml(fx("excel-anchor.xml"))!;
    expect(a.from).toEqual({ col: 5, row: 24, x: 15875 / 9525, y: 142875 / 9525 });
    expect(a.to.col).toBe(11);
    expect(a.rId).toBe("rId1");
    expect(a.name).toBe("Diagramm 1");
    const back = anchorFromXml(anchorXml({ ...a, chartEx: false }))!;
    expect(back.from).toEqual(a.from);
    expect(back.to).toEqual(a.to);
    expect(anchorFromXml(anchorXml({ ...a, chartEx: true }))!.chartEx).toBe(true);
  });
});

describe("write → read keeps the model", () => {
  const rows: CellValue[][] = [[null, "Reihe A", "Reihe B"], ["Jan", 4.3, 2.4], ["Feb", 2.5, 4.4], ["Mär", 3.5, 1.8]];
  const block = { sheet: "Meine Daten", r0: 0, c0: 0, r1: 3, c1: 2, get: (r: number, c: number) => rows[r]?.[c] ?? null };
  const kinds: ChartKind[] = ["column", "columnStacked", "bar", "barStacked", "line", "lineMarkers", "area", "pie", "doughnut", "scatter", "scatterSmooth", "scatterLines", "histogram"];
  const strip = (m: ChartModel) => {
    const { raw: _r, ...rest } = m;
    void _r;
    return JSON.parse(JSON.stringify(rest));
  };
  for (const kind of kinds) {
    it(kind, () => {
      const m = newChart(kind, seriesFromBlock(block, kind), { col: 4, row: 1, x: 3, y: 2 }, { col: 10, row: 16, x: 0, y: 0 });
      m.title = { show: true, text: "Umsatz & <Kosten>" };
      if (kind !== "pie" && kind !== "doughnut" && kind !== "histogram") m.axisTitles = { cat: "Monat", val: "€" };
      m.legend = kind === "histogram" ? "none" : "t";
      m.dataLabels = true;
      if (kind === "column" || kind === "line" || kind.startsWith("scatter")) m.series[0].trendline = { type: "linear", showEq: true, showR2: false };
      m.series[1 % m.series.length].color = "#123456";
      // The file always has a number format for numbers (General when the cells have none).
      for (const ser of m.series) {
        ser.val.format ??= "General";
        if (ser.cat && !ser.cat.cache.some((v) => typeof v === "string")) ser.cat.format ??= "General";
      }
      const xml = chartToXml(m);
      expect(xml.startsWith('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')).toBe(true);
      expect(isChartEx(m)).toBe(kind === "histogram");
      const back = chartFromXml(xml, { from: m.from, to: m.to });
      const want = strip(m);
      // Pie slices take the theme's colours point by point; a series colour is not kept.
      if (kind === "pie" || kind === "doughnut") for (const ser of want.series) delete ser.color;
      // Histograms keep one series; their colour and axis titles are not written.
      if (kind === "histogram") {
        delete want.series[0].color;
        expect(strip(back).series[0].val).toEqual(want.series[0].val);
        expect(back.title).toEqual(m.title);
        return;
      }
      expect(strip(back)).toEqual(want);
    });
  }

  it("writes the style and colour parts Excel needs next to a chart", () => {
    expect(chartStyleXml()).toContain("<cs:dataPointMarkerLayout");
    expect(chartColorsXml()).toContain('meth="cycle"');
  });
});
