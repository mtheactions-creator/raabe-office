// Every option of the chart Format pane (D-111): written to the chart part and read back unchanged,
// and visible in the drawing (axis lines, tick marks, logarithmic scales, crossings, display units,
// gradients, patterns, effects, points, explosion, label options, object properties).

import { describe, expect, it } from "vitest";
import { chartScene, newChart, sceneMarkup, type ChartKind, type ChartModel, type Prim } from "../src/index";
import { logScale } from "../src/paint";
import { anchorFromXml, anchorXml, chartFromXml, chartToXml, chartToXmlWithImages } from "../../ooxml/src/xlsx/chart";

const series = () => [
  { name: { text: "A" }, cat: { cache: ["x", "y", "z", "w"] }, val: { cache: [1, 3, 2, 5] } },
  { name: { text: "B" }, cat: { cache: ["x", "y", "z", "w"] }, val: { cache: [2, -1, 4, 3] } },
];
const anchor = { col: 0, row: 0, x: 0, y: 0 }, to = { col: 8, row: 15, x: 0, y: 0 };

function chart(kind: ChartKind = "column", patch: Partial<ChartModel> = {}): ChartModel {
  return { ...newChart(kind, series(), anchor, to), ...patch };
}

function roundTrip(m: ChartModel): ChartModel {
  const back = chartFromXml(chartToXml(m));
  delete back.raw;
  return back;
}

const scene = (m: ChartModel) => chartScene(m, 360, 216);
const parts = (m: ChartModel, part: string) => scene(m).prims.filter((p) => p.part === part);
const paths = (m: ChartModel, part: string) => parts(m, part).filter((p): p is Extract<Prim, { t: "path" }> => p.t === "path");

describe("the value axis and its tick marks", () => {
  it("draws no value axis line by default and one with Einfarbige Linie", () => {
    expect(paths(chart(), "axis:val")).toHaveLength(0);
    const m = chart("column", { format: { valAxis: { line: { type: "solid", color: "#C00000", width: 1.5 } } } });
    const p = paths(m, "axis:val");
    expect(p).toHaveLength(1);
    expect(p[0].stroke).toBe("#C00000");
    expect(p[0].sw).toBe(1.5);
    expect(p[0].d).toMatch(/^M[\d.]+ [\d.]+V[\d.]+$/);
  });

  it("draws tick marks in the axis line's paint (Außen left of the axis, Innen right, Kreuz both)", () => {
    const line = { type: "solid" as const, color: "#000000", width: 0.75 };
    const d = (tick: "in" | "out" | "cross") => paths(chart("column", { format: { valAxis: { line, majorTick: tick } } }), "axis:val")[0].d;
    const segs = (s: string) => [...s.matchAll(/M([\d.]+) ([\d.]+)H([\d.]+)/g)].map((x) => [Number(x[1]), Number(x[3])]);
    const axisX = (s: string) => Number(/^M([\d.]+)/.exec(s)![1]);
    const out = segs(d("out")), inn = segs(d("in")), cross = segs(d("cross"));
    const ax = axisX(d("out"));
    expect(out.length).toBeGreaterThan(3);
    expect(Math.min(...out[0])).toBeLessThan(ax);
    expect(Math.max(...out[0])).toBeCloseTo(ax, 1);
    expect(Math.min(...inn[0])).toBeCloseTo(ax, 1);
    expect(Math.max(...inn[0])).toBeGreaterThan(ax);
    expect(Math.abs(cross[0][1] - cross[0][0])).toBeCloseTo(7.5, 1);
    // No line: no tick marks either (Excel draws them with the axis' line).
    expect(paths(chart("column", { format: { valAxis: { majorTick: "out" } } }), "axis:val")).toHaveLength(0);
  });

  it("draws the category axis' tick marks between the categories", () => {
    const m = chart("column", { format: { catAxis: { majorTick: "out" } } });
    const d = paths(m, "axis:cat")[0].d;
    expect([...d.matchAll(/V/g)].length).toBe(5 + 0); // four categories: five boundaries
    const skip = chart("column", { format: { catAxis: { majorTick: "out", tickMarkSkip: 2 } } });
    expect([...paths(skip, "axis:cat")[0].d.matchAll(/V/g)].length).toBe(3);
  });
});

describe("value axis scales", () => {
  it("takes logarithmic scales in powers of the Basis", () => {
    const s = logScale(3, 4200, { logBase: 10 });
    expect(s).toMatchObject({ min: 1, max: 10000, log: 10 });
    const m = chart("line", { series: [{ val: { cache: [3, 40, 4200] } }], format: { valAxis: { logBase: 10 } } });
    const labels = parts(m, "axis:val").filter((p) => p.t === "text").map((p) => (p as { text: string }).text);
    expect(labels).toEqual(["1", "10", "100", "1000", "10000"]);
  });

  it("divides labels by the Anzeigeeinheit and names it", () => {
    const m = chart("column", { series: [{ val: { cache: [1200, 3400, 5000] } }], format: { valAxis: { dispUnits: "thousands" } } });
    const texts = parts(m, "axis:val").filter((p) => p.t === "text").map((p) => (p as { text: string }).text);
    expect(texts).toContain("Tausende");
    expect(texts).toContain("6");
    const hidden = chart("column", { series: [{ val: { cache: [1200, 3400, 5000] } }], format: { valAxis: { dispUnits: "thousands", dispUnitsLabel: false } } });
    expect(parts(hidden, "axis:val").some((p) => p.t === "text" && p.text === "Tausende")).toBe(false);
  });

  it("puts the category axis where the value axis is crossed (Achsenwert, Maximaler Achsenwert)", () => {
    const y = (m: ChartModel) => Number(/H/.test(paths(m, "axis:cat")[0].d) ? /M[\d.]+ ([\d.]+)H/.exec(paths(m, "axis:cat")[0].d)![1] : 0);
    const base = y(chart()), at3 = y(chart("column", { format: { valAxis: { crosses: 3 } } })), top = y(chart("column", { format: { valAxis: { crosses: "max" } } }));
    expect(at3).toBeLessThan(base);
    expect(top).toBeLessThan(at3);
    // Labels "Niedrig" stay at the bottom while the axis moves.
    const low = chart("column", { format: { valAxis: { crosses: "max" }, catAxis: { labelPos: "low" } } });
    const ly = (parts(low, "axis:cat").find((p) => p.t === "text") as { y: number }).y;
    expect(ly).toBeGreaterThan(base);
  });

  it("moves the value axis to the right at the largest category", () => {
    const s = scene(chart("column", { format: { catAxis: { crosses: "max" } } }));
    const label = s.prims.find((p) => p.part === "axis:val" && p.t === "text") as { x: number; anchor: string };
    expect(label.anchor).toBe("start");
    expect(label.x).toBeGreaterThan(s.plot.x + s.plot.w);
  });

  it("skips category labels by the Intervall and rotates them by the Benutzerdefinierter Winkel", () => {
    const m = chart("column", { format: { catAxis: { tickLblSkip: 2 } } });
    expect(parts(m, "axis:cat").filter((p) => p.t === "text").map((p) => (p as { text: string }).text)).toEqual(["x", "z"]);
    const r = chart("column", { format: { catAxis: { text: { box: { rotation: -30 } } } } });
    expect((parts(r, "axis:cat").find((p) => p.t === "text") as { rotate?: number }).rotate).toBe(-30);
  });
});

describe("fills, lines and effects", () => {
  it("draws gradients, patterns and effects through SVG definitions", () => {
    const m = chart("column", {
      format: {
        chart: { fill: { type: "pattern", pattern: { preset: "dkDnDiag", fg: "#000000", bg: "#FFFFFF" } }, roundedCorners: true, effects: { shadow: { color: "#000000", transparency: 50, blur: 4, distance: 3, angle: 45 } } },
        plot: { fill: { type: "gradient", gradient: { kind: "linear", angle: 90, stops: [{ pos: 0, color: "#FFFFFF" }, { pos: 100, color: "#4472C4", transparency: 20 }] } } },
      },
    });
    m.series[0].effects = { glow: { color: "#FFC000", size: 5, transparency: 40 }, softEdges: 2 };
    const s = scene(m);
    const markup = sceneMarkup(s);
    expect(markup).toContain("<pattern");
    expect(markup).toContain("<linearGradient");
    expect(markup).toContain("feDropShadow".slice(0, 0) + "<filter");
    expect(markup).toMatch(/rx="[\d.]+"/);
    expect(s.prims.find((p) => p.part === "series:0:0")!).toHaveProperty("filter");
  });

  it("draws Verbundtyp, Abschlusstyp, Anschlusstyp and arrowheads", () => {
    const m = chart("line");
    m.series[0].line = { type: "solid", color: "#000000", width: 3, compound: "dbl", cap: "sq", join: "bevel", head: { type: "oval" }, tail: { type: "stealth", w: "lg", len: "lg" } };
    const markup = sceneMarkup(scene(m));
    expect(markup).toContain("<mask");
    expect(markup).toContain('stroke-linecap="square"');
    expect(markup).toContain('stroke-linejoin="bevel"');
    expect(markup).toMatch(/marker-start="url\(#[^)]+\)" marker-end="url\(#[^)]+\)"/);
  });

  it("pulls pie slices out (Punktexplosion) and gives points their own fill", () => {
    const m = chart("pie");
    m.series[0].points = [{ idx: 1, explosion: 25, fill: { type: "solid", color: "#FF0000" } }];
    const s = scene(m);
    const slice = s.prims.find((p) => p.part === "series:0:1") as { fill: string };
    expect(slice.fill).toBe("#FF0000");
    const plain = scene(chart("pie"));
    expect(s.plot.w).toBeLessThan(plain.plot.w);
  });

  it("colours the points of one series in turn (Punkte nach Farbe unterscheiden)", () => {
    const m = chart("column", { series: [series()[0]], varyColors: true });
    const fills = [0, 1, 2].map((i) => (scene(m).prims.find((p) => p.part === `series:0:${i}`) as { fill: string }).fill);
    expect(new Set(fills).size).toBe(3);
  });
});

describe("the pane's options in the file", () => {
  it("keeps axis options (log scale, crossings, units, ticks, intervals, offset, position, types)", () => {
    const m = chart("column", {
      format: {
        valAxis: { logBase: 10, crosses: 2, majorTick: "cross", minorTick: "in", dispUnits: "thousands", line: { type: "solid", color: "#C00000", width: 1.5, tail: { type: "triangle", w: "lg", len: "sm" } }, text: { box: { rotation: -45 } } },
        catAxis: { crosses: "max", tickLblSkip: 2, tickMarkSkip: 3, lblOffset: 300, onTicks: true, axisType: "text", majorTick: "out", effects: { glow: { color: "#FFC000", size: 5, transparency: 60 } } },
      },
    });
    const back = roundTrip(m);
    expect(back.format?.valAxis).toEqual(m.format!.valAxis);
    expect(back.format?.catAxis).toEqual(m.format!.catAxis);
  });

  it("keeps a date axis and display units without label", () => {
    const m = chart("line", { format: { catAxis: { axisType: "date", baseTimeUnit: "months", majorUnit: 2, majorTimeUnit: "months" }, valAxis: { dispUnits: 500, dispUnitsLabel: false } } });
    m.series.forEach((s) => { s.cat = { cache: [45000, 45031, 45061, 45092], format: "mmm yy" }; });
    const back = roundTrip(m);
    expect(back.format?.catAxis).toEqual(m.format!.catAxis);
    expect(back.format?.valAxis).toEqual(m.format!.valAxis);
  });

  it("keeps gradient, pattern fills, line details and effects of every element", () => {
    const grad = { kind: "radial" as const, stops: [{ pos: 0, color: "#FFFFFF" }, { pos: 60, color: "#9DC3E6", transparency: 30 }, { pos: 100, color: "#1F4E79" }] };
    const m = chart("column", {
      format: {
        chart: { fill: { type: "gradient", gradient: { kind: "linear", angle: 45, stops: grad.stops } }, line: { type: "solid", color: "#000000", width: 2, compound: "thickThin", cap: "rnd", join: "miter" }, roundedCorners: true,
          effects: { shadow: { color: "#000000", transparency: 60, blur: 4, distance: 3, angle: 45, size: 102 }, format3d: { top: { type: "circle", w: 6, h: 6 }, material: "metal", lighting: "soft" } } },
        plot: { fill: { type: "pattern", pattern: { preset: "smGrid", fg: "#A5A5A5", bg: "#FFFFFF" } }, effects: { shadow: { inner: true, color: "#000000", transparency: 50, blur: 5, distance: 4, angle: 225 } } },
        title: { fill: { type: "gradient", gradient: grad }, effects: { softEdges: 2.5 }, text: { fill: { type: "gradient", gradient: grad }, outline: { type: "solid", color: "#FF0000", width: 0.5 }, effects: { glow: { color: "#FFFF00", size: 8, transparency: 60 } }, underline: true, box: { rotation: 10, vAlign: "t", margins: { l: 5, t: 2, r: 5, b: 2 } } } },
        legend: { line: { type: "gradient", gradient: grad, width: 1 } },
        majorGrid: { line: { type: "solid", color: "#BFBFBF", width: 0.75, dash: "dash" }, effects: { glow: { color: "#00B0F0", size: 3, transparency: 50 } } },
      },
    });
    const back = roundTrip(m);
    for (const k of ["chart", "plot", "title", "legend", "majorGrid"] as const) expect(back.format?.[k], k).toEqual(m.format![k]);
  });

  it("keeps series effects, explosion, data points, markers and Punkte nach Farbe", () => {
    const pie = chart("pie");
    pie.series[0].explosion = 10;
    pie.series[0].points = [{ idx: 2, explosion: 30, fill: { type: "solid", color: "#70AD47" }, line: { type: "solid", color: "#000000", width: 1 } }];
    const p2 = roundTrip(pie);
    expect(p2.series[0].explosion).toBe(10);
    expect(p2.series[0].points).toEqual(pie.series[0].points);

    const col = chart("column", { series: [series()[0]], varyColors: true });
    col.series[0].effects = { shadow: { color: "#000000", transparency: 60, blur: 4, distance: 3, angle: 45 } };
    col.series[0].points = [{ idx: 1, fill: { type: "pattern", pattern: { preset: "pct50", fg: "#FF0000", bg: "#FFFFFF" } } }];
    const c2 = roundTrip(col);
    expect(c2.varyColors).toBe(true);
    expect(c2.series[0].effects).toEqual(col.series[0].effects);
    expect(c2.series[0].points).toEqual(col.series[0].points);

    const line = chart("lineMarkers");
    line.series[0].marker = { symbol: "diamond", size: 8, fill: { type: "solid", color: "#FFFF00" }, line: { type: "solid", color: "#000000", width: 1 } };
    line.series[0].points = [{ idx: 0, marker: { symbol: "square", size: 10 } }];
    const l2 = roundTrip(line);
    expect(l2.series[0].marker).toEqual(line.series[0].marker);
    expect(l2.series[0].points?.[0].marker).toMatchObject({ symbol: "square", size: 10 });
  });

  it("keeps data label options (Wert aus Zellen, Blasengröße, Führungslinien, effects)", () => {
    const m = chart("column", { dataLabels: true, format: { dataLabels: { fromCells: { ref: "Tabelle1!$D$2:$D$5", cache: ["a", "b", null, "d"] }, leaderLines: true, effects: { glow: { color: "#FFC000", size: 5, transparency: 60 } }, text: { box: { rotation: -90 } } } } });
    const back = roundTrip(m);
    expect(back.format?.dataLabels).toEqual(m.format!.dataLabels);
    const pie = chart("pie", { dataLabels: true, format: { dataLabels: { leaderLines: false, showPercent: true } } });
    expect(roundTrip(pie).format?.dataLabels).toEqual(pie.format!.dataLabels);
  });

  it("keeps the secondary axis' options and the legend at the top right", () => {
    const m = chart("column", { legend: "tr", format: { valAxis2: { min: 0, max: 100, majorTick: "out", line: { type: "solid", color: "#ED7D31", width: 1 } } } });
    m.series[1].secondary = true;
    m.series[1].kind = "lineMarkers";
    const back = roundTrip(m);
    expect(back.legend).toBe("tr");
    expect(back.format?.valAxis2).toEqual(m.format!.valAxis2);
    expect(back.series[1].secondary).toBe(true);
  });

  it("writes a picture fill with a relationship placeholder the drawing replaces", () => {
    const png = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==";
    const m = chart("column", { format: { plot: { fill: { type: "picture", picture: { src: png, tile: true, transparency: 20 } } } } });
    const { xml, images } = chartToXmlWithImages(m);
    expect(images).toEqual([png]);
    expect(xml).toContain('r:embed="@img0"');
    const back = chartFromXml(xml.replace("@img0", "rId9"), { from: anchor, to, image: (id) => (id === "rId9" ? png : undefined) });
    expect(back.format?.plot?.fill).toEqual(m.format!.plot!.fill);
  });

  it("keeps the object's Eigenschaften and Alternativtext in the drawing anchor", () => {
    const props = { placement: "oneCell" as const, print: false, locked: false, lockAspect: true, altTitle: "Umsatz", altText: "Säulen je Monat" };
    const xml = anchorXml({ from: anchor, to, id: 2, name: "Diagramm 1", rId: "rId1", props });
    expect(anchorFromXml(xml)?.props).toEqual(props);
  });
});
