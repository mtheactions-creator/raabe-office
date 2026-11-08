// Konsolidieren against Excel's Range.Consolidate (probe_excel_consolidate.py, fixture
// excel-consolidate.json): every recorded case (by position, by labels, the eleven functions,
// destination ranges, a source on the destination sheet, Verknüpfungen mit Quelldaten) gives the
// same cells (values, formulas as Excel shows them, number formats) and the same outline; the
// order of new labels follows Excel's rule.

import { readFileSync } from "node:fs";
import { describe, expect, it } from "vitest";
import { Calc, Sheet, Workbook, parseRange, cellName, toLocalFormula, type ConsolidateFunction } from "@raabe/sheet-engine";
import * as A from "../src/actions";
import { History } from "../src/ops";
import { addLabels, applyPlan, labelKey, plan, settingsOf, sortSources, type Source } from "../src/consolidate";

type Rec = { formula: string; local: string; value: unknown; text: string; format: string; bold: boolean };
type Case = { name: string; sources: string[]; function: number; top: boolean; left: boolean; links: boolean; dest: string; used?: string; cells?: Record<string, Rec>; rows?: Record<string, [number, boolean]>; error?: string; dataConsolidate: string | null };
const rec = JSON.parse(readFileSync("packages/sheet-engine/test/fixtures/excel-consolidate.json", "utf-8")) as { sources: Record<string, Record<string, string | number>>; cases: Case[]; order: { sources: string[][]; rows: string[] }[] };
const FN: Record<number, ConsolidateFunction> = { [-4157]: "sum", [-4112]: "count", [-4106]: "average", [-4136]: "max", [-4139]: "min", [-4149]: "product", [-4113]: "countNums", [-4155]: "stdDev", [-4156]: "stdDevp", [-4164]: "var", [-4165]: "varp" };

function book() {
  const wb = new Workbook();
  for (const n of ["Gesamt", "Nord", "Süd", "West"]) wb.sheets.push(new Sheet(n));
  const history = new History(wb);
  const ctx = { wb, history };
  history.step("setup", () => {
    for (const [name, data] of Object.entries(rec.sources)) for (const [a, v] of Object.entries(data)) A.enter(ctx, wb.sheetByName(name)!, parseRange(a)!, String(v));
    // The probe's formats (Excel showed "0.00" as "#.000" through COM's locale, the same cell).
    A.format(ctx, wb.sheetByName("Nord")!, [parseRange("B2")!], (s) => { s.numFmt = "0.00"; });
    A.format(ctx, wb.sheetByName("Süd")!, [parseRange("B3")!], (s) => { s.numFmt = "0%"; });
  });
  const calc = new Calc(wb);
  calc.recalc();
  history.onChange = () => calc.recalc();
  return { wb, history, ctx, calc };
}

/** "Nord!R2C2:R4C3" as a source. */
const source = (wb: Workbook, t: string): Source => {
  const [sh, rc] = t.split("!");
  const m = /R(\d+)C(\d+):R(\d+)C(\d+)/.exec(rc)!.slice(1).map(Number);
  return { sheet: wb.sheetByName(sh)!, rect: { r0: m[0] - 1, c0: m[1] - 1, r1: m[2] - 1, c1: m[3] - 1 } };
};

/** The workbook name Excel wrote beside the link rows (its own new workbook's, "MappeN"). */
const bookOf = (c: Case) => Object.values(c.cells ?? {}).map((x) => x.value).find((v): v is string => typeof v === "string" && /^Mappe\d+$/.test(v)) ?? "Mappe1";
const norm = (v: unknown) => (v && typeof v === "object" && "code" in (v as object) ? { error: (v as { code: string }).code } : v);

describe("Konsolidieren as Excel's", () => {
  for (const c of rec.cases) {
    if (!c.cells) continue;
    it(c.name, () => {
      const b = book();
      const dest = b.wb.sheetByName("Gesamt")!;
      if (c.name === "Quelle auf dem Zielblatt") b.history.step("pre", () => { for (const [a, v] of Object.entries({ A10: "1", B10: "2", A11: "3", B11: "4" })) A.enter(b.ctx, dest, parseRange(a)!, v); });
      const p = plan(b.wb, parseRange(c.dest)!, c.sources.map((s) => source(b.wb, s)), { fn: FN[c.function], top: c.top, left: c.left, link: c.links, bookName: bookOf(c) });
      b.history.step("Konsolidieren", () => applyPlan(b.ctx, dest, p));
      b.calc.recalc();
      const used = parseRange(c.used!)!;
      const r1 = Math.max(used.r1, p.area.r1), c1 = Math.max(used.c1, p.area.c1);
      for (let row = 0; row <= r1; row++) for (let col = 0; col <= c1; col++) {
        const a = cellName(row, col);
        const want = c.cells![a];
        const cell = dest.get(row, col);
        const label = `${c.name} ${a}`;
        if (!want || want.formula === "") { expect(cell?.v ?? null, label).toBeNull(); continue; }
        if (want.local.startsWith("=")) expect(cell?.f !== undefined ? "=" + toLocalFormula(cell.f) : cell?.v, label).toBe(want.local);
        else expect(cell?.f, label).toBeUndefined();
        const v = norm(cell?.v ?? null);
        if (typeof want.value === "number") expect(v as number, label).toBeCloseTo(want.value, 9);
        else expect(v, label).toEqual(want.value);
        const fmt = b.wb.styles.get(cell?.s).numFmt;
        expect(fmt === "General" ? "Standard" : fmt === "0.00" ? "#.000" : fmt, label).toBe(want.format);
      }
      const rows: Record<string, [number, boolean]> = {};
      for (const [r, info] of dest.rows) if (info.outline || info.hidden) rows[String(r + 1)] = [(info.outline ?? 0) + 1, !!info.hidden];
      expect(rows, c.name).toEqual(c.rows ?? {});
    });
  }

  it("keeps the settings as Excel writes them into the sheet", () => {
    const b = book();
    for (const c of rec.cases) {
      if (!c.dataConsolidate || !c.cells) continue;
      const set = settingsOf(c.sources.map((s) => source(b.wb, s)), { fn: FN[c.function], top: c.top, left: c.left, link: c.links });
      const refs = [...c.dataConsolidate.matchAll(/<dataRef ref="([^"]*)" sheet="([^"]*)"\/>/g)].map((m) => ({ ref: m[1], sheet: m[2] }));
      expect(set.refs, c.name).toEqual(refs);
      expect(!!set.top, c.name).toBe(c.top);
      expect(set.fn ?? "sum", c.name).toBe(FN[c.function]);
    }
  });

  it("sorts the references as Excel's list", () => {
    const b = book();
    expect(sortSources([source(b.wb, "West!R1C1:R3C4"), source(b.wb, "Nord!R2C2:R4C3")]).map((s) => s.sheet.name)).toEqual(["Nord", "West"]);
  });
});

describe("Where Konsolidieren puts a new label (recorded)", () => {
  for (const o of rec.order) {
    it(JSON.stringify(o.sources), () => {
      const list: { key: string; value: string }[] = [];
      for (const src of o.sources) addLabels(list, src.map((v) => ({ key: labelKey(v)!, value: v })));
      expect(list.map((l) => l.value)).toEqual(o.rows);
    });
  }
});
