// Lasting cases for what the by-hand audit of RaabeCalc found (D-141, D-142, sweep_calc.py): each bug that a
// student met through the ribbon is pinned here at the level of the model.

import { describe, expect, it } from "vitest";
import { Calc, Sheet, Workbook, parseRange, type Rect } from "@raabe/sheet-engine";
import * as A from "../src/actions";
import { History } from "../src/ops";
import { CondEval } from "../src/grid/cond";

const R = (a: string): Rect => parseRange(a)!;

function book(cells: Record<string, string>) {
  const wb = Workbook.blank();
  const history = new History(wb);
  const ctx = { wb, history };
  const sheet = wb.sheet();
  history.step("setup", () => { for (const [a, t] of Object.entries(cells)) A.enter(ctx, sheet, R(a), t); });
  new Calc(wb).recalc();
  return { wb, history, ctx, sheet };
}

describe("colours set through a format are plain colours", () => {
  it("drops the standard font's design reference (Text 1) when a colour is set", () => {
    const b = book({ A1: "x" });
    b.wb.styles.get(undefined);
    b.history.step("Farbe", () => A.format(b.ctx, b.sheet, [R("A1")], (s) => { s.font.theme = { t: 1 }; }));
    b.history.step("Farbe", () => A.format(b.ctx, b.sheet, [R("A1")], (s) => { s.font.color = "#006100"; }));
    const st = b.wb.styles.get(b.sheet.get(0, 0)!.s);
    expect(st.font.color).toBe("#006100");
    expect(st.font.theme).toBeUndefined();
  });

  it("keeps a design reference set together with its colour", () => {
    const b = book({ A1: "x" });
    b.history.step("Farbe", () => A.format(b.ctx, b.sheet, [R("A1")], (s) => { s.font.color = "#4472C4"; s.font.theme = { t: 4 }; }));
    expect(b.wb.styles.get(b.sheet.get(0, 0)!.s).font.theme).toEqual({ t: 4 });
  });

  it("drops the fill's design reference when a plain fill is set", () => {
    const b = book({ A1: "x" });
    b.history.step("Füllung", () => A.format(b.ctx, b.sheet, [R("A1")], (s) => { s.fill = "#4472C4"; s.fillTheme = { t: 4 }; }));
    b.history.step("Füllung", () => A.format(b.ctx, b.sheet, [R("A1")], (s) => { s.fill = "#C6EFCE"; }));
    const st = b.wb.styles.get(b.sheet.get(0, 0)!.s);
    expect(st.fill).toBe("#C6EFCE");
    expect(st.fillTheme).toBeUndefined();
  });
});

describe("undo of steps on the sheet list", () => {
  it("brings back a hidden sheet's visibility and a renamed sheet's name", () => {
    const b = book({ A1: "x" });
    b.wb.sheets.push(new Sheet("Tabelle2"));
    const first = b.wb.sheets[0];
    b.history.step("Ausblenden", () => { b.history.sheets(); first.hidden = true; b.wb.active = 1; });
    b.history.step("Umbenennen", () => { b.history.sheets(); b.wb.sheets[1].name = "Noten"; });
    b.history.undo();
    expect(b.wb.sheets[1].name).toBe("Tabelle2");
    b.history.undo();
    expect(first.hidden).toBeFalsy();
    b.history.redo();
    expect(first.hidden).toBe(true);
    b.history.redo();
    expect(b.wb.sheets[1].name).toBe("Noten");
  });
});

describe("undo of defined names", () => {
  it("takes back a name made by Namen definieren, and brings it back with redo", () => {
    const b = book({ A1: "1" });
    b.history.step("Namen definieren", () => { b.wb.names = [...b.wb.names, { name: "Preis", ref: "Tabelle1!$A$1" }]; });
    expect(b.history.undoLabel).toBe("Namen definieren");
    b.history.undo();
    expect(b.wb.names.some((n) => n.name === "Preis")).toBe(false);
    b.history.redo();
    expect(b.wb.names.some((n) => n.name === "Preis")).toBe(true);
  });
});

describe("conditional formats", () => {
  it("evaluates a rule's formula relative to the format's top-left cell", () => {
    const b = book({ A1: "a", B1: "5", A2: "b", B2: "20", A3: "c", B3: "30" });
    b.sheet.conditional = [{ ref: [R("A1:A3")], rules: [{ type: "expression", priority: 1, formulas: ["$B1>10"], format: { fill: "#FFC7CE" } }] }];
    const ev = new CondEval(b.wb, b.sheet, 0);
    expect(ev.look(0, 0)?.fill).toBeUndefined();
    expect(ev.look(1, 0)?.fill).toBe("#FFC7CE");
    expect(ev.look(2, 0)?.fill).toBe("#FFC7CE");
  });

  it("evaluates a Datum rule (timePeriod) through its formula", () => {
    const b = book({ A1: "=HEUTE()", A2: "=HEUTE()-1", A3: "" });
    b.sheet.conditional = [{ ref: [R("A1:A3")], rules: [{ type: "timePeriod", timePeriod: "today", priority: 1, formulas: ["FLOOR(A1,1)=TODAY()"], format: { fill: "#FFC7CE" } }] }];
    const ev = new CondEval(b.wb, b.sheet, 0);
    expect(ev.look(0, 0)?.fill).toBe("#FFC7CE");
    expect(ev.look(1, 0)?.fill).toBeUndefined();
    expect(ev.look(2, 0)?.fill).toBeUndefined();
  });
});
