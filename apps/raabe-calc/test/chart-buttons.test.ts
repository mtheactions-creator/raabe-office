// The chart buttons' model side (D-115 - D-118): Diagrammelemente's entries and choices as recorded in Excel
// (results/2026-11-09-excel-chart-buttons.json, inventory 2026-11-05), the Diagrammformatvorlagen from Excel's saved
// charts, the range fields of Datenquelle auswählen and the filtered view the chart engine draws.

import { readFileSync } from "node:fs";
import { describe, expect, it } from "vitest";
import { chartScene, newChart, seriesFromBlock, visibleModel, type ChartModel } from "@raabe/charts";
import { choiceLabel, elementEntries } from "@raabe/chart-ui";
import { applyStyle, chartStyles, styleIndex } from "@raabe/chart-ui";
import { parseName, parseValues } from "../src/charts/selectdata";

const excel = JSON.parse(readFileSync("tools/feel-lab/results/2026-11-09-excel-chart-buttons.json", "utf-8"));
const styles = JSON.parse(readFileSync("tools/feel-lab/results/2026-11-09-excel-chart-filter.json", "utf-8"));

const DATA: (string | number)[][] = [["Monat", "Nord", "Süd", "West"], ["Jan", 3, 4, 2], ["Feb", 5, 2, 3], ["Mrz", 2, 6, 4], ["Apr", 4, 3, 5]];
function chart(kind: ChartModel["kind"] = "column", cols = 4): ChartModel {
  const series = seriesFromBlock({ sheet: "Tabelle1", r0: 0, c0: 0, r1: 4, c1: cols - 1, get: (r, c) => DATA[r][c] }, kind);
  const m = newChart(kind, series, { col: 0, row: 0, x: 0, y: 0 }, { col: 8, row: 15, x: 0, y: 0 });
  for (const s of m.series) {
    s.val.cache = [1, 2, 3, 4].map((r) => DATA[r][Number(/\$([A-Z])\$/.exec(s.val.ref!)![1].charCodeAt(0) - 65)] as number);
    s.cat = { ...s.cat!, cache: ["Jan", "Feb", "Mrz", "Apr"] };
    s.name = { ...s.name, text: DATA[0][Number(/\$([A-Z])\$/.exec(s.val.ref!)![1].charCodeAt(0) - 65)] as string };
  }
  return m;
}

describe("Diagrammelemente", () => {
  it("lists Excel's entries for a column chart, ticked as Excel ticks them", () => {
    const m = chart();
    const rows = elementEntries(m).filter((e) => !e.ribbonOnly && !e.disabled);
    const recorded = [...new Set(excel.flyouts["Element hinzufügen"].main.flatMap((p: { items: { type: string; name: string }[] }) => p.items.filter((i) => i.type === "MenuItem").map((i) => i.name)))];
    expect(rows.map((e) => e.label)).toEqual(recorded);
    expect(rows.filter((e) => e.on(m)).map((e) => e.label)).toEqual(["Achsen", "Diagrammtitel", "Gitternetzlinien", "Legende"]);
  });

  it("the ribbon's Diagrammelement hinzufügen has the inventory's entries", () => {
    expect(elementEntries(chart(), true).map((e) => e.label)).toEqual(["Achsen", "Achsentitel", "Diagrammtitel", "Datenbeschriftungen", "Datentabelle",
      "Fehlerindikatoren", "Gitternetzlinien", "Legende", "Linien", "Trendlinie", "Pos./Neg. Abweichung"]);
    const labels = elementEntries(chart(), true).find((e) => e.id === "labels")!;
    expect(["Keine", ...labels.choices.map((c) => choiceLabel(c, true)), labels.more]).toEqual(["Keine", "Mitte", "Am Ende innerhalb", "Am Anfang innerhalb", "Am Ende außerhalb", "Datenlegende", "Weitere Datenbeschriftungsoptionen..."]);
  });

  it("a pie offers Diagrammtitel, Datenbeschriftungen and Legende", () => {
    expect(elementEntries(chart("pie", 2)).map((e) => e.label)).toEqual(["Diagrammtitel", "Datenbeschriftungen", "Legende"]);
  });

  it("every check box switches its element and back", () => {
    for (const e of elementEntries(chart()).filter((x) => !x.ribbonOnly && !x.disabled)) {
      const m = chart();
      const was = e.on(m);
      e.toggle(m, !was);
      expect(e.on(m), e.label).toBe(!was);
      e.toggle(m, was);
      expect(e.on(m), e.label).toBe(was);
    }
  });

  it("side menu choices set what Excel sets", () => {
    const m = chart();
    const by = (id: string) => elementEntries(m).find((e) => e.id === id)!;
    by("title").choices[1].set(m, true);
    expect([m.title.show, m.format?.title?.overlay]).toEqual([true, true]);
    by("labels").choices[1].set(m, true);
    expect([m.dataLabels, m.format?.dataLabels?.pos]).toEqual([true, "inEnd"]);
    by("dataTable").choices[1].set(m, true);
    expect(m.dataTable).toEqual({ legendKeys: false });
    by("errorBars").choices[2].set(m, true);
    expect(m.series.map((s) => s.errorBars?.type)).toEqual(["stdDev", "stdDev", "stdDev"]);
    by("gridlines").choices[3].set(m, true);
    expect(m.format?.catMinorGrid?.show).toBe(true);
    by("axes").choices[1].set(m, false);
    expect(m.format?.valAxis?.visible).toBe(false);
    by("legend").choices[2].set(m, true);
    expect(m.legend).toBe("l");
  });
});

describe("Diagrammformatvorlagen", () => {
  it("the gallery has as many styles per kind as Excel's", () => {
    expect(chartStyles(chart()).length).toBe(styles.gallery.column.length);
    expect(chartStyles(chart("bar")).length).toBe(styles.gallery.bar.length);
    expect(chartStyles(chart("line")).length).toBe(styles.gallery.line.length);
    expect(chartStyles(chart("pie", 2)).length).toBe(styles.gallery.pie.length);
  });

  it("style 1 is the plain look; style 2 puts labels on, the value axis off, wide gaps, Excel's style number", () => {
    const m = chart();
    applyStyle(m, 0);
    expect(m.format).toBeUndefined();
    expect(m.style).toBeUndefined();
    applyStyle(m, 1);
    expect(m.style).toBe(202);
    expect(m.dataLabels).toBe(true);
    expect(m.format?.valAxis?.visible).toBe(false);
    expect(m.gapWidth).toBe(444);
    expect(styleIndex(m)).toBe(1);
    // Every style of every kind draws.
    for (const kind of ["column", "bar", "line", "pie"] as const) {
      const n = chartStyles(chart(kind, kind === "pie" ? 2 : 4)).length;
      for (let i = 0; i < n; i++) {
        const c = chart(kind, kind === "pie" ? 2 : 4);
        applyStyle(c, i);
        expect(chartScene(c, 360, 216).prims.length, `${kind} ${i + 1}`).toBeGreaterThan(5);
      }
    }
  });
});

describe("Datenquelle auswählen and the filter view", () => {
  it("reads the range fields as Excel accepts them", () => {
    expect(parseValues("=Tabelle1!$B$2:$B$5")).toEqual({ ref: "Tabelle1!$B$2:$B$5" });
    expect(parseValues("={1;2;3}")).toEqual({ values: [1, 2, 3] });
    expect(parseValues("=Quatsch")).toBeNull();
    expect(parseName("=Tabelle1!$B$1")).toEqual({ ref: "Tabelle1!$B$1" });
    expect(parseName('="Nord"')).toEqual({ text: "Nord" });
    expect(parseName("Nord")).toEqual({ text: "Nord" });
  });

  it("filtered series and categories are left out, the others keep their colours", () => {
    const m = chart();
    m.series[1].filtered = true;
    m.hiddenCats = [2];
    const v = visibleModel(m, ["#111111", "#222222", "#333333", "#444444", "#555555", "#666666"]);
    expect(v.index).toEqual([0, 2]);
    expect(v.model.series.map((s) => s.color)).toEqual(["#111111", "#333333"]);
    expect(v.model.series[0].val.cache).toEqual([3, 5, 4]);
    expect(v.model.series[0].cat?.cache).toEqual(["Jan", "Feb", "Apr"]);
  });
});
