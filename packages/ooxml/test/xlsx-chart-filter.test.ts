// Diagrammfilter, Datentabelle and Ausgeblendete und leere Zellen as Excel 365 saves them (D-115, D-116):
// the chart parts Excel wrote (tools/feel-lab/probe_excel_chart_filter.py) read into the model, and the model
// written back in the same form.

import { readFileSync } from "node:fs";
import { describe, expect, it } from "vitest";
import { visibleModel } from "@raabe/charts";
import { chartFromXml, chartToXml } from "../src/index";
import { hiddenOf, partialRef } from "../src/xlsx/chartfilter";

const fx = (name: string) => readFileSync(`packages/ooxml/test/fixtures/excel-chart-filter/${name}.xml`, "utf-8");
/** Written anew (not verbatim): the read XML is dropped. */
const rewrite = (m: ReturnType<typeof chartFromXml>) => chartToXml({ ...m, raw: undefined });

describe("Diagrammfilter in the file", () => {
  it("reads a filtered series where it was, marked filtered", () => {
    const m = chartFromXml(fx("series2-filtered"));
    expect(m.series.map((s) => s.name?.text)).toEqual(["Nord", "Süd", "West"]);
    expect(m.series.map((s) => !!s.filtered)).toEqual([false, true, false]);
    expect(m.series[1].val.ref).toBe("Daten!$C$2:$C$5");
    expect(visibleModel(m).model.series.map((s) => s.name?.text)).toEqual(["Nord", "West"]);
  });

  it("writes a filtered series into c15:filteredBarSeries with formulaRef", () => {
    const m = chartFromXml(fx("series2-filtered"));
    const xml = rewrite(m);
    expect(xml).toContain("<c15:filteredBarSeries><c15:ser><c:idx val=\"1\"/>");
    expect(xml).toContain("<c15:formulaRef><c15:sqref>Daten!$C$2:$C$5</c15:sqref></c15:formulaRef>");
    expect((xml.match(/<c:ser>/g) ?? []).length).toBe(2);
    const back = chartFromXml(xml);
    expect(back.series.map((s) => !!s.filtered)).toEqual([false, true, false]);
  });

  it("keeps Excel's XML verbatim while nothing changes", () => {
    const xml = fx("series2-filtered");
    expect(chartToXml(chartFromXml(xml))).toBe(xml);
  });

  it("reads a hidden category from c15:fullRef and spreads the caches over the whole range", () => {
    const m = chartFromXml(fx("cat3-filtered"));
    expect(m.hiddenCats).toEqual([2]);
    expect(m.series[0].cat?.ref).toBe("Daten!$A$2:$A$5");
    expect(m.series[0].val.cache).toEqual([3, 5, null, 4]);
    expect(m.series[0].cat?.cache).toEqual(["Jan", "Feb", null, "Apr"]);
    const v = visibleModel(m).model;
    expect(v.series[0].cat?.cache).toEqual(["Jan", "Feb", "Apr"]);
  });

  it("writes hidden categories as Excel does: partial c:f, whole range in c15:fullRef, visible points cached", () => {
    const m = chartFromXml(fx("cat3-filtered"));
    m.series[0].val.cache = [3, 5, 2, 4];
    const xml = rewrite(m);
    expect(xml).toContain("<c15:fullRef><c15:sqref>Daten!$B$2:$B$5</c15:sqref></c15:fullRef></c:ext></c:extLst><c:f>(Daten!$B$2:$B$3,Daten!$B$5)</c:f>");
    expect(xml).toMatch(/<c:f>\(Daten!\$B\$2:\$B\$3,Daten!\$B\$5\)<\/c:f><c:numCache><c:formatCode>General<\/c:formatCode><c:ptCount val="3"\/>/);
    expect(chartFromXml(xml).hiddenCats).toEqual([2]);
  });

  it("a series and a category filtered together, line, pie, bar, area and scatter", () => {
    const both = chartFromXml(fx("series-and-cat-filtered"));
    expect(both.series.map((s) => !!s.filtered)).toEqual([true, false, false]);
    expect(both.hiddenCats).toEqual([0]);
    // The filtered series keeps its whole ranges (c15:fullRef and c15:formulaRef in one extension).
    expect(both.series[0].val.ref).toBe("Daten!$B$2:$B$5");
    expect(both.series[0].val.cache).toEqual([null, 5, 2, 4]);
    const xml = rewrite(both);
    expect(xml).toContain("<c15:fullRef><c15:sqref>Daten!$B$2:$B$5</c15:sqref></c15:fullRef><c15:formulaRef><c15:sqref>Daten!$B$3:$B$5</c15:sqref></c15:formulaRef>");
    expect((xml.match(/<c:extLst><c:ext uri="\{02D57815[^>]*><c15:fullRef>/g) ?? []).length).toBe(6);
    const again = chartFromXml(xml);
    expect(again.series.map((s) => !!s.filtered)).toEqual([true, false, false]);
    expect(again.hiddenCats).toEqual([0]);
    for (const [name, filtered] of [["line-series2-filtered", [false, true, false]], ["bar-series3-filtered", [false, false, true]],
      ["area-series1-filtered", [true, false, false]], ["scatter-series1-filtered", [true, false, false]]] as const) {
      const m = chartFromXml(fx(name));
      expect(m.series.map((s) => !!s.filtered), name).toEqual(filtered);
      expect(chartFromXml(rewrite(m)).series.map((s) => !!s.filtered), name).toEqual(filtered);
    }
    expect(chartFromXml(fx("line-cat2-filtered")).hiddenCats).toEqual([1]);
    expect(chartFromXml(fx("pie-cat2-filtered")).hiddenCats).toEqual([1]);
  });

  it("partial references as Excel writes them", () => {
    expect(partialRef("Daten!$A$2:$A$5", [2])).toBe("(Daten!$A$2:$A$3,Daten!$A$5)");
    expect(partialRef("Daten!$A$2:$A$5", [0])).toBe("Daten!$A$3:$A$5");
    expect(partialRef("Daten!$B$1:$E$1", [1, 2])).toBe("(Daten!$B$1,Daten!$E$1)");
    expect(hiddenOf("Daten!$A$2:$A$5", "(Daten!$A$2:$A$3,Daten!$A$5)")).toEqual([2]);
  });
});

describe("Datentabelle, Ausgeblendete und leere Zellen", () => {
  it("reads and writes the Datentabelle with and without legend keys", () => {
    expect(chartFromXml(fx("datatable-keys")).dataTable).toEqual({ legendKeys: true });
    expect(chartFromXml(fx("datatable-nokeys")).dataTable).toEqual({ legendKeys: false });
    const m = chartFromXml(fx("plain"));
    expect(m.dataTable).toBeUndefined();
    m.dataTable = { legendKeys: false };
    const xml = rewrite(m);
    expect(xml).toMatch(/<\/c:valAx><c:dTable><c:showHorzBorder val="1"\/><c:showVertBorder val="1"\/><c:showOutline val="1"\/><c:showKeys val="0"\/>/);
    expect(chartFromXml(xml).dataTable).toEqual({ legendKeys: false });
  });

  it("reads and writes Lücken / Null / verbinden, ausgeblendete Daten and #NV", () => {
    expect(chartFromXml(fx("blanks-zero")).dispBlanksAs).toBe("zero");
    expect(chartFromXml(fx("blanks-span")).dispBlanksAs).toBe("span");
    expect(chartFromXml(fx("hidden-shown")).showHidden).toBe(true);
    const m = chartFromXml(fx("plain"));
    expect([m.dispBlanksAs, m.showHidden, m.naAsBlank]).toEqual([undefined, undefined, undefined]);
    Object.assign(m, { dispBlanksAs: "span", showHidden: true, naAsBlank: false });
    const xml = rewrite(m);
    expect(xml).toContain('<c:plotVisOnly val="0"/><c:dispBlanksAs val="span"/>');
    expect(xml).toContain('<c16r3:dispNaAsBlank val="0"/>');
    const back = chartFromXml(xml);
    expect([back.dispBlanksAs, back.showHidden, back.naAsBlank]).toEqual(["span", true, false]);
  });

  it("writes the category axis' Hilfsgitternetz", () => {
    const m = chartFromXml(fx("plain"));
    m.format = { ...(m.format ?? {}), catMinorGrid: { show: true } };
    const xml = rewrite(m);
    expect(xml).toMatch(/<c:catAx>(?:(?!<\/c:catAx>)[\s\S])*<c:minorGridlines>/);
    expect(chartFromXml(xml).format?.catMinorGrid?.show).toBe(true);
  });
});
