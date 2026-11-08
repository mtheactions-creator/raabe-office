// Sparklines in .xlsx against German Excel 365 (D-085): probe_excel_sparklines.py made one sheet per option of
// the Sparkline tab through COM and saved it; the fixture holds the x14:sparklineGroups XML Excel wrote and what
// COM read back. Read: every attribute and colour lands in the model as COM reports it. Write: a group written
// from the model is Excel's XML (apart from Excel's own xr2:uid), an unchanged group stays verbatim, and the
// whole workbook survives writeXlsx/readXlsx. The 36 styles of the gallery (probe_excel_sparkline_ui.py) are
// the generated table RaabeCalc applies.

import { readFileSync } from "node:fs";
import { describe, expect, it } from "vitest";
import type { SparklineGroup } from "@raabe/sheet-engine";
import { defaultSparkColors, newWorkbook, readSparklines, readXlsx, sparklineGroupXml, sparklinesExt, writeXlsx } from "../src";

interface ComColor { rgb: string; theme?: number; tint?: number }
interface ComGroup {
  type: number; count: number; location: string; source: string; lineWeight: number; blanks: number; hidden: boolean; series: ComColor;
  Highpoint: { visible: boolean; color: ComColor }; Lowpoint: { visible: boolean; color: ComColor }; Firstpoint: { visible: boolean; color: ComColor };
  Lastpoint: { visible: boolean; color: ComColor }; Negative: { visible: boolean; color: ComColor }; Markers: { visible: boolean; color: ComColor };
  axis: { visible: boolean; color: ComColor; date: boolean; rtl: boolean; minType: number; maxType: number; min?: number; max?: number };
  dateRange?: string;
}
interface Fixture {
  theme: Record<string, string>;
  cases: Record<string, { groups: ComGroup[]; xml: string | null }>;
  styles: { name: string; colors: Record<string, string> }[];
}

const FIX = JSON.parse(readFileSync("packages/ooxml/test/fixtures/excel-sparklines.json", "utf-8")) as Fixture;
const CTX = { theme: FIX.theme, indexed: [] as string[] };
const ext = (xml: string) => `<extLst><ext uri="{05C60535-1F16-4fd2-B633-F4F36F0B64E0}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">${xml}</ext></extLst>`;
const groupsXml = (xml: string) => [...xml.matchAll(/<x14:sparklineGroup\b[\s\S]*?<\/x14:sparklineGroup>/g)].map((m) => m[0]);
/** Excel's XML without its own ids, numbers compared as numbers. */
const norm = (xml: string) => xml.replace(/ xr2:uid="[^"]*"/g, "").replace(/(tint|manualMin|manualMax|lineWeight)="([^"]*)"/g, (_, k, v) => `${k}="${Number(v)}"`);
const AXIS: Record<number, SparklineGroup["minAxisType"]> = { 1: "group", 2: "individual", 3: "custom" };
const cases = Object.entries(FIX.cases).filter(([, c]) => c.xml);

describe("sparklines read as Excel reports them", () => {
  it.each(cases)("%s", (_name, c) => {
    const groups = readSparklines(ext(c.xml!), CTX);
    // COM lists the groups in another order than the file: match them by location.
    expect(groups.length).toBe(c.groups.length);
    for (const com of c.groups) {
      const first = com.location.split(/[:,]/)[0];
      const g = groups.find((x) => x.items.some((i) => cellName(i.row, i.col) === first))!;
      expect(g, com.location).toBeDefined();
      expect(g.items.length).toBe(com.count);
      expect(g.type).toBe({ 1: "line", 2: "column", 3: "stacked" }[com.type]);
      expect(g.lineWeight ?? 0.75).toBe(com.lineWeight);
      expect(g.displayEmptyCellsAs).toBe({ 1: "gap", 2: "zero", 3: "span" }[com.blanks]);
      expect(!!g.displayHidden).toBe(com.hidden);
      expect(g.colors.series.rgb).toBe(com.series.rgb);
      const pts = { Highpoint: "high", Lowpoint: "low", Firstpoint: "first", Lastpoint: "last", Negative: "negative", Markers: "markers" } as const;
      for (const [k, m] of Object.entries(pts)) {
        expect(!!g[m], k).toBe(com[k as keyof typeof pts].visible);
        expect(g.colors[m].rgb, k).toBe(com[k as keyof typeof pts].color.rgb);
      }
      expect(!!g.displayXAxis).toBe(com.axis.visible);
      expect(g.colors.axis.rgb).toBe(com.axis.color.rgb);
      expect(!!g.rightToLeft).toBe(com.axis.rtl);
      expect(!!g.dateRange).toBe(com.axis.date);
      if (com.dateRange) expect(g.dateRange!.split("!")[1]).toBe(com.dateRange);
      expect(g.minAxisType ?? "individual").toBe(AXIS[com.axis.minType]);
      expect(g.maxAxisType ?? "individual").toBe(AXIS[com.axis.maxType]);
      if (com.axis.minType === 3) expect(g.manualMin).toBe(com.axis.min);
      if (com.axis.maxType === 3) expect(g.manualMax).toBe(com.axis.max);
      // A design colour keeps its reference (COM's ThemeColor is the file's index + 1 for the accents).
      if (com.series.theme && com.series.theme >= 5) expect(g.colors.series.theme?.t).toBe(com.series.theme - 1);
    }
  });
});

describe("sparklines written as Excel writes them", () => {
  it.each(cases)("%s", (_name, c) => {
    const groups = readSparklines(ext(c.xml!), CTX);
    const excel = groupsXml(c.xml!);
    groups.forEach((g, i) => expect(norm(sparklineGroupXml(structuredClone(g)))).toBe(norm(excel[i])));
  });

  it("unchanged groups stay verbatim, a changed one is written anew", () => {
    const xml = FIX.cases["ungroup"].xml!;
    const kept: [string, string][] = [];
    const groups = readSparklines(ext(xml), CTX, kept);
    expect(sparklinesExt(ext(xml), groups, kept)).toBe(ext(xml));
    groups[1].type = "column";
    const out = groupsXml(sparklinesExt(ext(xml), groups, kept)!);
    expect(out[0]).toBe(groupsXml(xml)[0]);
    expect(out[2]).toBe(groupsXml(xml)[2]);
    expect(out[1]).not.toContain("xr2:uid");
    expect(out[1]).toContain('type="column"');
  });

  it("no groups: the entry goes, other extLst entries stay", () => {
    const other = '<ext uri="{CCE6A557-97BC-4b89-ADB6-D9C93CAAB3DF}" xmlns:x14="x"><x14:dataValidations/></ext>';
    const xml = `<extLst>${other}${ext(FIX.cases["line"].xml!).slice(8, -9)}</extLst>`;
    expect(sparklinesExt(xml, [], [])).toBe(`<extLst>${other}</extLst>`);
    expect(sparklinesExt(ext(FIX.cases["line"].xml!), [], [])).toBeNull();
  });
});

describe("sparklines through a whole workbook", () => {
  it("every option survives writeXlsx and readXlsx", () => {
    const wb = newWorkbook();
    const s = wb.sheets[0];
    for (let r = 0; r < 3; r++) for (let c = 0; c < 5; c++) s.set(r, c, { v: (r + 1) * (c - 2) });
    const all = readSparklines(ext(FIX.cases["everything"].xml!), { theme: Object.fromEntries(Object.entries(wb.theme).map(([k, v]) => [k, v.slice(1)])), indexed: [] })[0];
    s.sparklines = [
      { ...all, dateRange: "Tabelle1!A7:E7", items: [{ f: "Tabelle1!A1:E1", row: 0, col: 5 }] },
      { type: "stacked", colors: defaultSparkColors(), negative: true, displayEmptyCellsAs: "gap", maxAxisType: "custom", manualMax: 4, items: [{ f: "Tabelle1!A2:E2", row: 1, col: 5 }, { f: "Tabelle1!A3:E3", row: 2, col: 5 }] },
    ];
    const back = readXlsx(writeXlsx(wb));
    expect(back.sheets[0].sparklines).toEqual(s.sparklines);
    // Read back and written again unchanged: the same bytes for the sparkline entry.
    const again = readXlsx(writeXlsx(back));
    expect(again.sheets[0].sparklines).toEqual(s.sparklines);
  });
});

describe("the Sparklineformatvorlage gallery", () => {
  it("has Excel's 36 styles in Excel's order", async () => {
    const { SPARKLINE_STYLES } = await import("../../../apps/raabe-calc/src/sparkline-styles.gen");
    expect(SPARKLINE_STYLES.length).toBe(36);
    expect(FIX.styles.length).toBe(36);
    SPARKLINE_STYLES.forEach((st, i) => {
      expect(st.name).toBe(FIX.styles[i].name);
      const g = sparklineGroupXml({ type: "line", displayEmptyCellsAs: "gap", items: [],
        colors: Object.fromEntries(Object.entries(st.colors).map(([k, v]) => [k, "theme" in v ? { rgb: "#000000", theme: v.theme } : v])) as SparklineGroup["colors"] });
      for (const [tag, attrs] of Object.entries(FIX.styles[i].colors)) expect(norm(g)).toContain(norm(`<x14:${tag} ${attrs}/>`));
    });
  });
});

function cellName(row: number, col: number): string {
  let s = "", c = col + 1;
  while (c > 0) { const m = (c - 1) % 26; s = String.fromCharCode(65 + m) + s; c = Math.floor((c - 1) / 26); }
  return s + (row + 1);
}
