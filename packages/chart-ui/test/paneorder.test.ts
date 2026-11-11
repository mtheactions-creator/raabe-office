// The chart Format pane's order against Excel's (Step 11, D-160): for every pane recorded in our own Excel
// (tools/feel-lab/results/2026-11-09-excel-chart-pane.json) the option sets, tabs, sections and controls must stand in
// Excel's order. Controls that only one side has are listed (KNOWN below): a new one fails the test too, so the list
// only shrinks.

import { readFileSync } from "node:fs";
import { describe, expect, it } from "vitest";
import { newChart, type ChartKind, type ChartModel } from "@raabe/charts";
import type { PaneControl, PaneSpec } from "@raabe/shell";
import { formatPane, type ChartElement } from "@raabe/chart-ui";
import { compareChoices, comparePane, excelUnits, oursUnits, type ExcelPane, type PaneDiff } from "./paneorder";

const REC = JSON.parse(readFileSync(`${__dirname}/../../../tools/feel-lab/results/2026-11-09-excel-chart-pane.json`, "utf-8")) as {
  panes: Record<string, ExcelPane>;
  markers: Record<string, { sections: ExcelPane["sets"][number]["tabs"][number]["sections"] }>;
};

const anchor = { col: 0, row: 0, x: 0, y: 0 }, to = { col: 8, row: 15, x: 0, y: 0 };

/** The chart Excel's probe made for each type (probe_excel_chartpane.py: meta "chart setup"). */
function chart(kind: string): ChartModel {
  const ours: Record<string, ChartKind> = { column: "column", bar: "bar", line: "lineMarkers", pie: "pie", scatter: "scatter", area: "area", combo: "column" };
  const series = [
    { name: { text: "A" }, cat: { cache: ["Jan", "Feb", "Mär", "Apr"] }, val: { cache: [4, -2, 7, 5] } },
    { name: { text: "B" }, cat: { cache: ["Jan", "Feb", "Mär", "Apr"] }, val: { cache: [6, 5, 3, 8] } },
  ];
  const m = newChart(ours[kind], series, anchor, to);
  m.dataLabels = true;
  if (kind !== "pie") {
    m.series[0].trendline = { type: "linear", showEq: false, showR2: false };
    m.series[0].errorBars = { dir: "both", type: "fixedVal", value: 1 };
    m.axisTitles.val = "Achsentitel";
    m.gridlines = true;
  }
  if (kind === "combo") { m.series[1].kind = "lineMarkers"; m.series[1].secondary = true; }
  return m;
}

/** Excel's element (the recording's key after the type) as ours. */
const ELEMENT: Record<string, ChartElement | null> = {
  Diagrammbereich: { kind: "chart" }, "Zeichnungsfläche": { kind: "plot" }, Diagrammtitel: { kind: "title" }, Legende: { kind: "legend" },
  "Achse 2 (Wert)": { kind: "valAxis" }, "Achse 1 (Kategorie bzw. X)": { kind: "catAxis" }, Hauptgitternetzlinien: { kind: "majorGrid" },
  Achsentitel: { kind: "axisTitle", axis: "val" }, Datenreihe: { kind: "series", index: 0 }, Datenpunkt: { kind: "point", index: 0, point: 1 },
  Datenbeschriftungen: { kind: "labels" }, Trendlinie: { kind: "trend", index: 0 }, Fehlerindikatoren: { kind: "errorBars", index: 0 },
  "Datenreihe 2": { kind: "series", index: 1 }, "Sekundärachse (Wert)": { kind: "valAxis2" }, "Datenreihe 2 (Linie, Sekundärachse)": { kind: "series", index: 1 },
  // A single data label has no pane of its own in ours (the labels' pane serves it).
  "Datenbeschriftung (eine)": null,
};

function ourPane(key: string): PaneSpec | null {
  const [kind, el] = key.split("/");
  const e = ELEMENT[el];
  if (!e) return null;
  const m = chart(kind);
  return formatPane(m, e, { edit: (fn) => fn(m), select: () => undefined, size: () => ({ w: 12.7, h: 7.62 }), resize: () => undefined });
}

/**
 * Controls only one side has, known and written down (D-160): "<element> | <set > tab > section> | excel-only or
 * ours-only | <label>", the element as Excel's recording names it (every chart type that has it).
 */
const KNOWN = JSON.parse(readFileSync(`${__dirname}/paneorder-known.json`, "utf-8")) as string[];

const onlyKey = (d: PaneDiff) => `${d.pane.split("/")[1]} | ${d.where} | ${d.kind} | ${d.label}`;

describe("the chart Format pane stands in Excel's order", () => {
  const all: PaneDiff[] = [];
  for (const key of Object.keys(REC.panes)) {
    const ours = ourPane(key);
    if (!ours) continue;
    // A pane recorded as the same as another one (same_as) is compared with that one; one Excel failed to open has none.
    let rec: ExcelPane & { same_as?: string } = REC.panes[key];
    while (rec?.same_as) rec = REC.panes[rec.same_as];
    if (!rec?.sets) continue;
    all.push(...comparePane(key, rec, ours, REC.panes, REC.markers));
    all.push(...compareChoices(key, rec, () => ourPane(key)!, REC.panes));
  }
  if (process.env.PANEORDER_DUMP) console.log(JSON.stringify(all.filter((d) => d.kind === "order"), null, 1), JSON.stringify(all.filter((d) => d.kind !== "order").map((d) => onlyKey(d)).filter((x, i, a) => a.indexOf(x) === i).sort(), null, 1));
  it("has no section or control out of Excel's order", () => {
    const order = all.filter((d) => d.kind === "order").map((d) => `${d.pane} | ${d.where} | Excel: ${d.excel.join(", ")} | ours: ${d.ours.join(", ")}`);
    expect(order).toEqual([]);
  });
  it("lists every control only one side has (the list only shrinks)", () => {
    const only = [...new Set(all.filter((d) => d.kind !== "order").map((d) => onlyKey(d)))].sort();
    const known = new Set(KNOWN);
    expect(only.filter((x) => !known.has(x))).toEqual([]);
    // Fixed ones are taken off the list.
    expect(KNOWN.filter((x) => !only.includes(x))).toEqual([]);
  });
  it("reads the recording's units as the pane shows them", () => {
    const ctl = (c: Partial<{ t: string; label: string; picker: string; value: unknown; menu: boolean }>) => ({ t: "", label: "", ...c });
    expect(excelUnits([ctl({ t: "Text", label: "Farbe" }), ctl({ t: "MenuItem", label: "Füllfarbe", picker: "Farbe" }), ctl({ t: "Slider", label: "Transparenz" }), ctl({ t: "Spinner", label: "Transparenz" })]))
      .toEqual(["Farbe", "Transparenz"]);
    const ours: PaneControl[] = [{ type: "row", items: [{ type: "check", label: "a", value: () => true, set: () => undefined }] }, { type: "heading", label: "h" }];
    expect(oursUnits(ours)).toEqual(["a"]);
  });
});
