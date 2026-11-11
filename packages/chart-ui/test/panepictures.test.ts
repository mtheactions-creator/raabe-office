// Pictures where Excel shows pictures (D-161, D-163): every dropdown of Excel's chart Format pane that opens a list of
// pictures (recorded: a menu with entries, not a colour picker; results/2026-11-09-excel-chart-pane.json, and the
// marker Typ from results/2026-11-11-excel-picturedrops.json) must be a picture dropdown in ours, in every state of
// its section (each radio choice tried in a fresh pane). Excel's text lists (ComboBox) must not be pictures.

import { readFileSync, writeFileSync } from "node:fs";
import { describe, expect, it } from "vitest";
import { newChart, type ChartKind, type ChartModel } from "@raabe/charts";
import type { PaneControl, PaneSection, PaneSpec } from "@raabe/shell";
import { formatPane, type ChartElement } from "@raabe/chart-ui";
import type { ExcelControl, ExcelPane, ExcelSection } from "./paneorder";

const REC = JSON.parse(readFileSync(`${__dirname}/../../../tools/feel-lab/results/2026-11-09-excel-chart-pane.json`, "utf-8")) as {
  panes: Record<string, ExcelPane & { same_as?: string }>;
  markers: Record<string, { sections: ExcelSection[] }>;
};
const DROPS = JSON.parse(readFileSync(`${__dirname}/../../../tools/feel-lab/results/2026-11-11-excel-picturedrops.json`, "utf-8")) as {
  panes: Record<string, Record<string, Record<string, { types?: string[]; count?: number }>>>;
};

const ELEMENT: Record<string, ChartElement> = {
  Diagrammbereich: { kind: "chart" }, "Zeichnungsfläche": { kind: "plot" }, Diagrammtitel: { kind: "title" }, Legende: { kind: "legend" },
  "Achse 2 (Wert)": { kind: "valAxis" }, "Achse 1 (Kategorie bzw. X)": { kind: "catAxis" }, Hauptgitternetzlinien: { kind: "majorGrid" },
  Achsentitel: { kind: "axisTitle", axis: "val" }, Datenreihe: { kind: "series", index: 0 }, Datenpunkt: { kind: "point", index: 0, point: 1 },
  Datenbeschriftungen: { kind: "labels" }, Trendlinie: { kind: "trend", index: 0 }, Fehlerindikatoren: { kind: "errorBars", index: 0 },
};

function pane(key: string): PaneSpec | null {
  const [kind, el] = key.split("/");
  const e = ELEMENT[el];
  if (!e) return null;
  const ours: Record<string, ChartKind> = { column: "column", bar: "bar", line: "lineMarkers", pie: "pie", scatter: "scatter", area: "area", combo: "column" };
  const s = (n: string, v: number[]) => ({ name: { text: n }, cat: { cache: ["Jan", "Feb", "Mär", "Apr"] }, val: { cache: v } });
  const m: ChartModel = newChart(ours[kind], [s("A", [4, -2, 7, 5]), s("B", [6, 5, 3, 8])], { col: 0, row: 0, x: 0, y: 0 }, { col: 8, row: 15, x: 0, y: 0 });
  m.dataLabels = true;
  if (kind !== "pie") { m.series[0].trendline = { type: "linear", showEq: false, showR2: false }; m.series[0].errorBars = { dir: "both", type: "fixedVal", value: 1 }; m.axisTitles.val = "T"; }
  return formatPane(m, e, { edit: (fn) => fn(m), select: () => undefined, size: () => ({ w: 12.7, h: 7.62 }), resize: () => undefined });
}

const flat = (cs: PaneControl[]): PaneControl[] => cs.flatMap((c) => (c.type === "row" ? flat(c.items) : [c]));
const isPictureMenu = (c: ExcelControl) => c.t === "MenuItem" && !!c.menu && !c.picker && !String(c.value ?? "").startsWith("RGB(");

/** Our control types by label in a section, over its default state and every radio choice. */
function oursByLabel(make: () => PaneSpec | null, set: string, tab: string, title: string): Map<string, Set<string>> {
  const out = new Map<string, Set<string>>();
  const find = (p: PaneSpec | null): PaneSection | undefined => p?.sets.find((s) => s.label === set)?.tabs.find((t) => t.label === tab)?.sections().find((s) => s.title === title);
  const add = (cs: PaneControl[]) => { for (const c of flat(cs)) if ("label" in c) { const s = out.get(c.label) ?? new Set(); s.add(c.type); out.set(c.label, s); } };
  const first = find(make());
  if (!first) return out;
  add(first.controls());
  const radios = flat(first.controls()).filter((c): c is Extract<PaneControl, { type: "radio" }> => c.type === "radio").map((c) => c.label);
  for (const r of radios) {
    const p = make();
    const sec = find(p);
    const radio = sec && flat(sec.controls()).find((c) => c.type === "radio" && c.label === r);
    if (radio?.type !== "radio") continue;
    radio.pick();
    add(find(p)?.controls() ?? []);
  }
  return out;
}

describe("the chart Format pane shows pictures where Excel does", () => {
  const rows: { key: string; where: string; label: string; ours: string }[] = [];
  const textRows: { key: string; where: string; label: string; ours: string }[] = [];
  for (const [key, rec0] of Object.entries(REC.panes)) {
    let rec = rec0;
    while (rec?.same_as) rec = REC.panes[rec.same_as];
    if (!rec?.sets || !pane(key)) continue;
    for (const s of rec.sets) for (const t of s.tabs) for (const sec of t.sections) {
      let src: ExcelSection | undefined = sec;
      for (let d = 0; src?.same_as && d < 5; d++) src = REC.panes[src.same_as]?.sets?.flatMap((x) => x.tabs.flatMap((y) => y.sections)).find((x) => x.title === sec.title);
      const all = [...(src?.controls ?? []), ...Object.values(src?.choices ?? {}).flatMap((c) => (typeof c === "object" ? c.adds ?? [] : []))];
      const pics = new Set(all.filter(isPictureMenu).map((c) => c.label));
      const texts = new Set(all.filter((c) => c.t === "ComboBox").map((c) => c.label));
      if (!pics.size && !texts.size) continue;
      const ours = oursByLabel(() => pane(key), s.label, t.label, sec.title);
      for (const l of pics) rows.push({ key, where: `${s.label} > ${t.label} > ${sec.title}`, label: l, ours: [...(ours.get(l) ?? ["fehlt"])].join("/") });
      for (const l of texts) if (ours.get(l)?.has("picture")) textRows.push({ key, where: `${s.label} > ${t.label} > ${sec.title}`, label: l, ours: "picture" });
    }
  }
  it("has a picture dropdown for every picture list of Excel's pane", () => {
    const wrong = rows.filter((r) => r.ours !== "picture" && r.ours !== "fehlt").map((r) => `${r.key} | ${r.where} | ${r.label}: ${r.ours}`);
    expect(wrong).toEqual([]);
    // Counts for the record (D-163): Excel's picture lists over all recorded panes, and how many ours has.
    expect(rows.length).toBeGreaterThan(100);
    if (process.env.PANEPICTURES_DUMP) writeFileSync(process.env.PANEPICTURES_DUMP, JSON.stringify({ total: rows.length, picture: rows.filter((r) => r.ours === "picture").length, missing: rows.filter((r) => r.ours === "fehlt").map((r) => `${r.key} | ${r.where} | ${r.label}`) }, null, 1));
  });
  it("keeps Excel's text lists as text", () => {
    expect(textRows).toEqual([]);
  });
  it("draws the marker Typ as pictures, as Excel's list (recorded 2026-11-11)", () => {
    const typ = DROPS.panes["Diagramm: Datenreihe (Linie)"]?.Markierung;
    expect(typ).toBeDefined();
    const ours = oursByLabel(() => pane("line/Datenreihe"), "Datenreihenoptionen", "Füllung und Linie", "Markierungsoptionen");
    expect([...(ours.get("Typ") ?? [])]).toEqual(["picture"]);
  });
});
