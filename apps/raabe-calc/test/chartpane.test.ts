// The chart Format pane's spec (D-111, D-112): every element of the main chart types has Excel's tabs and
// sections (results/2026-11-09-excel-chart-pane.json), and its controls change the model as one edit each.

import { describe, expect, it } from "vitest";
import { newChart, type ChartKind, type ChartModel } from "@raabe/charts";
import type { PaneControl, PaneSpec } from "@raabe/shell";
import { chartElements, elementName, elementOfPart, formatPane, type ChartElement } from "../src/charts/formatpane";

const anchor = { col: 0, row: 0, x: 0, y: 0 }, to = { col: 8, row: 15, x: 0, y: 0 };
const series = () => [
  { name: { text: "A" }, cat: { cache: ["Jan", "Feb", "Mrz"] }, val: { cache: [3, -1, 5] } },
  { name: { text: "B" }, cat: { cache: ["Jan", "Feb", "Mrz"] }, val: { cache: [40, 20, 60] } },
];

function chart(kind: ChartKind): ChartModel {
  const m = newChart(kind, series(), anchor, to);
  m.dataLabels = true;
  return m;
}

function host(m: ChartModel, edits: { n: number }) {
  return {
    edit: (fn: (c: ChartModel) => void) => { edits.n++; fn(m); },
    select: () => undefined,
    size: () => ({ w: 12.7, h: 7.62 }),
    resize: () => undefined,
  };
}

/** Every control of a pane (all sets, tabs and sections), with where it stands. */
function controls(spec: PaneSpec): { where: string; c: PaneControl }[] {
  const out: { where: string; c: PaneControl }[] = [];
  for (const s of spec.sets) for (const t of s.tabs) for (const sec of t.sections()) {
    for (const c of sec.controls()) out.push({ where: `${s.label}/${t.label}/${sec.title}`, c });
  }
  return out;
}

const label = (c: PaneControl) => ("label" in c ? c.label : "");

describe("the chart Format pane", () => {
  it("has Excel's tabs for the value axis, with a line that draws", () => {
    const m = chart("column");
    const edits = { n: 0 };
    const spec = formatPane(m, { kind: "valAxis" }, host(m, edits));
    expect(spec.title).toBe("Achse formatieren");
    expect(spec.sets.map((s) => s.label)).toEqual(["Achsenoptionen", "Textoptionen"]);
    expect(spec.sets[0].tabs.map((t) => t.label)).toEqual(["Füllung und Linie", "Effekte", "Größe und Eigenschaften", "Achsenoptionen"]);
    expect(spec.sets[1].tabs.map((t) => t.label)).toEqual(["Textfüllung und -kontur", "Texteffekte", "Textfeld"]);
    const all = controls(spec);
    for (const l of ["Stärke", "Verbundtyp", "Strichtyp", "Abschlusstyp", "Anschlusstyp", "Startpfeiltyp", "Endpfeilgröße", "Haupttyp", "Hilfstyp",
      "Anzeigeeinheiten", "Logarithmische Skalierung", "Beschriftungsposition", "Rubrik", "Mit Quelle verknüpft", "Benutzerdefinierter Winkel"]) {
      expect(all.some((x) => label(x.c) === l), l).toBe(true);
    }
    const solid = all.find((x) => x.where.endsWith("/Linie") && label(x.c) === "Einfarbige Linie")!.c as Extract<PaneControl, { type: "radio" }>;
    solid.pick();
    expect(m.format?.valAxis?.line?.type).toBe("solid");
    expect(edits.n).toBe(1);
  });

  it("offers every chart type's own options", () => {
    const names = (kind: ChartKind, e: ChartElement) => controls(formatPane(chart(kind), e, host(chart(kind), { n: 0 }))).map((x) => label(x.c));
    expect(names("column", { kind: "series", index: 0 })).toEqual(expect.arrayContaining(["Primärachse", "Sekundärachse", "Reihenachsenüberlappung", "Abstandsbreite", "Invertieren, falls negativ"]));
    expect(names("pie", { kind: "series", index: 0 })).toEqual(expect.arrayContaining(["Winkel des ersten Segments", "Kreisexplosion", "Segmentfarbunterscheidung"]));
    expect(names("pie", { kind: "point", index: 0, point: 1 })).toEqual(expect.arrayContaining(["Punktexplosion"]));
    expect(names("lineMarkers", { kind: "series", index: 0 })).toEqual(expect.arrayContaining(["Linie glätten", "Integriert", "Typ"]));
    expect(names("column", { kind: "catAxis" })).toEqual(expect.arrayContaining(["Automatische Auswahl basierend auf Daten", "Bei Rubriknummer", "Auf Teilstrichen", "Intervall zwischen Teilstrichen", "Abstand zur Achse"]));
    expect(names("column", { kind: "labels" })).toEqual(expect.arrayContaining(["Wert aus Zellen", "Datenreihenname", "Rubrikenname", "Führungslinien anzeigen", "Am Ende außerhalb"]));
    expect(names("column", { kind: "chart" })).toEqual(expect.arrayContaining(["Abgerundete Ecken", "Höhe skalieren", "Nur von Zellposition abhängig", "Objekt drucken"]));
    expect(names("column", { kind: "legend" })).toEqual(expect.arrayContaining(["Oben rechts"]));
  });

  it("lets every enabled choice of the Füllung change the element", () => {
    const m = chart("column");
    const spec = formatPane(m, { kind: "plot" }, host(m, { n: 0 }));
    const fill = () => controls(formatPane(m, { kind: "plot" }, host(m, { n: 0 }))).filter((x) => x.where.endsWith("/Füllung"));
    for (const l of ["Keine Füllung", "Einfarbige Füllung", "Farbverlauf", "Musterfüllung"]) {
      (fill().find((x) => label(x.c) === l)!.c as Extract<PaneControl, { type: "radio" }>).pick();
      expect(m.format?.plot?.fill?.type, l).toBe({ "Keine Füllung": "none", "Einfarbige Füllung": "solid", Farbverlauf: "gradient", Musterfüllung: "pattern" }[l]);
    }
    expect(spec.title).toBe("Zeichnungsfläche formatieren");
  });

  it("keeps gradient stops in order when one is added", () => {
    const m = chart("column");
    m.format = { plot: { fill: { type: "gradient", gradient: { kind: "linear", angle: 90, stops: [{ pos: 0, color: "#FFFFFF" }, { pos: 100, color: "#000000" }] } } } };
    const g = controls(formatPane(m, { kind: "plot" }, host(m, { n: 0 }))).find((x) => x.c.type === "gradient")!.c as Extract<PaneControl, { type: "gradient" }>;
    g.add();
    expect(m.format.plot!.fill!.gradient!.stops.map((s) => s.pos)).toEqual([0, 50, 100]);
  });

  it("selects a point with a second click on a selected series", () => {
    expect(elementOfPart("series:1:2", { kind: "chart" })).toEqual({ kind: "series", index: 1 });
    expect(elementOfPart("series:1:2", { kind: "series", index: 1 })).toEqual({ kind: "point", index: 1, point: 2 });
    expect(elementOfPart("axis:val2")).toEqual({ kind: "valAxis2" });
    const m = chart("column");
    m.series[1].secondary = true;
    expect(chartElements(m).map((e) => elementName(m, e))).toContain("Sekundär Vertikal (Wert) Achse");
  });
});
