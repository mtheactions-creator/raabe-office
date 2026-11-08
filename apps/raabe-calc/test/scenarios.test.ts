// Szenario-Manager against Excel's Worksheet.Scenarios (probe_excel_scenarios.py, fixture
// excel-scenarios.json): Anzeigen gives the cells Excel gives (one after another, a text value, a
// formula replaced), the Szenariobericht has Excel's cells, values, look and outline (with and without
// Ergebniszellen, with names), Zusammenführen names a clash as Excel does.

import { readFileSync } from "node:fs";
import { describe, expect, it } from "vitest";
import { Calc, Workbook, cellName, parseRange, type Scenario } from "@raabe/sheet-engine";
import * as A from "../src/actions";
import { History } from "../src/ops";
import { formulaLocal } from "../src/engine";
import { REPORT_NOTE, applyScenario, buildReport, cellsText, createdComment, mergedName, parseCells, reportData, reportName } from "../src/scenarios";

const rec = JSON.parse(readFileSync("packages/sheet-engine/test/fixtures/excel-scenarios.json", "utf-8"));
// The probe's sheet as typed in German.
const SETUP: Record<string, string> = { A1: "Preis", B1: "10", A2: "Menge", B2: "100", A3: "Kosten", B3: "500", A4: "Gewinn", B4: "=B1*B2-B3", A5: "Steuer", B5: "=B4*0,3", D1: "Rabatt", E1: "0,05", E2: "=E1*2" };
const ERR: Record<string, string> = { "#VALUE!": "#VALUE!" };

function book() {
  const wb = Workbook.blank();
  const history = new History(wb);
  const ctx = { wb, history };
  const sheet = wb.sheet();
  history.step("setup", () => { for (const [a, t] of Object.entries(SETUP)) A.enter(ctx, sheet, parseRange(a)!, t); });
  const calc = new Calc(wb);
  calc.recalc();
  history.onChange = () => calc.recalc();
  return { wb, history, ctx, sheet, calc };
}

type RecScen = { name: string; cells: string; values: (string | number)[]; comment: string; locked: boolean; hidden: boolean };
const scenarios = (list: RecScen[]): Scenario[] => list.map((s) => {
  const cells = parseCells(s.cells.replace(/,/g, ";"))!;
  return { name: s.name, cells: cells.map((p, i) => ({ ...p, val: String(s.values[i]) })), ...(s.comment ? { comment: s.comment.replace("<user>", "Raabe Office") } : {}), ...(s.locked ? { locked: true } : {}), ...(s.hidden ? { hidden: true } : {}) };
});
const norm = (v: unknown) => (v && typeof v === "object" && "code" in (v as object) ? { error: ERR[(v as { code: string }).code] ?? (v as { code: string }).code } : v);

describe("Szenario anzeigen as Excel's", () => {
  it("enters each scenario's values in turn", () => {
    const b = book();
    b.sheet.scenarios = { list: scenarios(rec.added) };
    for (const [name, cells] of Object.entries(rec.shown as Record<string, Record<string, { formula: string; value: unknown }>>)) {
      const sc = b.sheet.scenarios.list.find((s) => s.name === name)!;
      b.history.step("Szenario anzeigen", () => applyScenario(b.ctx, b.sheet, sc));
      for (const [a, want] of Object.entries(cells)) {
        const r = parseRange(a)!;
        const cell = b.sheet.get(r.r0, r.c0);
        const v = norm(cell?.v ?? null);
        if (typeof want.value === "number") expect(v as number, `${name} ${a}`).toBeCloseTo(want.value, 9);
        else expect(v, `${name} ${a}`).toEqual(want.value);
        expect(cell?.f !== undefined, `${name} ${a} formula`).toBe(want.formula.startsWith("="));
      }
    }
    // Undo puts the cells back as they were before the last one.
    b.history.undo();
    expect(b.sheet.value(0, 4)).toBe(0.05);
  });

  it("replaces a formula among the changing cells by the value", () => {
    const b = book();
    b.history.step("x", () => applyScenario(b.ctx, b.sheet, { name: "Formel", cells: [{ row: 1, col: 4, val: "7" }] }));
    expect(formulaLocal(b.sheet.get(1, 4))).toBe(rec.formula_cell.formula);
    expect(b.sheet.value(1, 4)).toBe(rec.formula_cell.value);
  });
});

/** The recorded report's cells against ours: values, look, outline. */
function compareReport(name: string, recReport: { cells: Record<string, { value: unknown; bold: boolean; fill: string | null; fontColor: string; borders: Record<string, { weight: number }>; h: number; wrap: boolean; font: [string, number] }>; rows: Record<string, { level: number; hidden: boolean }>; cols: Record<string, { level: number }> }, results: string[], names: Record<string, string> = {}) {
  const b = book();
  b.sheet.scenarios = { list: scenarios(rec.added) };
  for (const [n, ref] of Object.entries(names)) b.wb.names.push({ name: n, ref });
  const res = results.map((a) => { const r = parseRange(a)!; return { row: r.r0, col: r.c0 }; });
  const data = reportData(b.wb, b.sheet, b.calc, res);
  const built = buildReport(b.wb, data, "Szenariobericht");
  built.comments();
  const rep = built.sheet;
  let noteRow = 0;
  for (const [a, want] of Object.entries(recReport.cells)) {
    const r = parseRange(a)!;
    const cell = rep.get(r.r0, r.c0);
    const st = b.wb.styles.get(cell?.s);
    const label = `${name} ${a}`;
    if (typeof want.value === "string" && want.value.startsWith("[Hinweis")) {
      // Excel's note is its own prose: ours stands in the same rows (own words).
      expect(cell?.v, label).toBe(REPORT_NOTE[noteRow++]);
    } else if (typeof want.value === "number") expect(cell?.v as number, label).toBeCloseTo(want.value, 9);
    else if (typeof want.value === "string" && want.value.startsWith("Erstellt von")) expect(cell?.v, label).toBe(want.value.replace(/von .* am/, "von Raabe Office am"));
    else expect(cell?.v ?? null, label).toEqual(want.value);
    expect(!!st.font.bold, `${label} bold`).toBe(want.bold);
    expect(st.fill ?? null, `${label} fill`).toBe(want.fill);
    expect(st.font.color ?? "#000000", `${label} font colour`).toBe(want.fontColor);
    expect(st.font.size, `${label} size`).toBe(want.font[1]);
    expect(!!st.wrap, `${label} wrap`).toBe(want.wrap);
    const weight = (s?: { style: string }) => (s ? (s.style === "medium" ? -4138 : 2) : undefined);
    expect({ top: weight(st.borders.top), bottom: weight(st.borders.bottom) }, `${label} borders`).toEqual({ top: want.borders.top?.weight, bottom: want.borders.bottom?.weight });
    expect(st.h === "left" ? -4131 : st.h === "right" ? -4152 : 1, `${label} alignment`).toBe(want.h);
  }
  expect(noteRow).toBe(3);
  for (const [r, want] of Object.entries(recReport.rows)) {
    const info = rep.rows.get(Number(r) - 1);
    expect((info?.outline ?? 0) + 1, `${name} row ${r} level`).toBe(want.level);
    expect(!!info?.hidden, `${name} row ${r} hidden`).toBe(want.hidden);
  }
  for (const [c, want] of Object.entries(recReport.cols)) expect((rep.cols.get(Number(c) - 1)?.outline ?? 0) + 1, `${name} col ${c}`).toBe(want.level);
  expect(rep.view.showGrid).toBe(false);
  // Nothing beyond Excel's used range.
  const used = parseRange(recReport.cells && Object.keys(recReport.cells).at(-1)!)!;
  for (const k of rep.cells.keys()) expect(Math.floor(k / 16384), `${name} ${cellName(Math.floor(k / 16384), k % 16384)}`).toBeLessThanOrEqual(used.r1);
}

describe("Szenariobericht as Excel's", () => {
  it("with Ergebniszellen B4:B5", () => compareReport("report", rec.report, ["B4", "B5"]));
  it("with one Ergebniszelle", () => compareReport("report_single", rec.report_single, ["B4"]));
  it("without Ergebniszellen", () => compareReport("report_noresult", rec.report_noresult, []));
  it("with named cells", () => compareReport("report_named", rec.report_named, ["B4", "B5"], { Preis: "Tabelle1!$B$1", Gewinn: "Tabelle1!$B$4" }));

  it("leaves the cells as they were", () => {
    const b = book();
    b.sheet.scenarios = { list: scenarios(rec.added) };
    reportData(b.wb, b.sheet, b.calc, [{ row: 3, col: 1 }]);
    expect(b.calc.value(0, 3, 1)).toBe(500);
    expect(b.sheet.get(0, 1)?.v).toBe(10);
  });

  it("names the sheets as Excel does", () => {
    const b = book();
    expect(reportName(b.wb)).toBe(rec.sheets_after_report[0]);
    b.wb.sheets.unshift(buildReport(b.wb, reportData(b.wb, b.sheet, b.calc, []), "Szenariobericht").sheet);
    expect(reportName(b.wb)).toBe(rec.second_report_name);
  });
});

describe("Szenario-Manager's texts as Excel's", () => {
  it("shows the cells as Excel's dialog", () => {
    expect(cellsText(parseCells("B1:B2")!)).toBe("$B$1:$B$2");
    expect(cellsText(parseCells("$B$1;$B$3")!)).toBe("$B$1;$B$3");
    expect(cellsText(parseCells("A1:B2;D4")!)).toBe("$A$1:$B$2;$D$4");
    expect(parseCells("A1:A40")!.length).toBeGreaterThan(32);
  });

  it("starts the comment and merges names as Excel", () => {
    const d = new Date(2026, 10, 8);
    expect(createdComment("<user>", d)).toBe(rec.added[2].comment);
    const merged = rec.merged.map((s: RecScen) => s.name);
    expect(mergedName("Gut", rec.added.map((s: RecScen) => s.name), d)).toBe(merged[5]);
    expect(mergedName("Neu", rec.added.map((s: RecScen) => s.name), d)).toBe(merged[6]);
  });
});
