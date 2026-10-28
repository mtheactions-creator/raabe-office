// RaabeCalc's cell operations: AutoAusfüllen as recorded in Excel (catalogue_calc.py), inserting
// and deleting rows with formulas following, copying with shifted references, undo and redo.

import { describe, expect, it } from "vitest";
import { Calc, Workbook, parseRange, type Rect } from "@raabe/sheet-engine";
import * as A from "../src/actions";
import { History } from "../src/ops";
import { formulaLocal, shown } from "../src/engine";

function book(cells: Record<string, string>) {
  const wb = Workbook.blank();
  const history = new History(wb);
  const ctx = { wb, history };
  const sheet = wb.sheet();
  history.step("setup", () => { for (const [a, t] of Object.entries(cells)) A.enter(ctx, sheet, parseRange(a)!, t); });
  const calc = new Calc(wb);
  calc.recalc();
  history.onChange = () => calc.recalc();
  const at = (a: string) => { const r = parseRange(a)!; return sheet.get(r.r0, r.c0); };
  const text = (a: string) => shown(at(a), wb.styles.get(at(a)?.s)).text;
  const formula = (a: string) => formulaLocal(at(a));
  return { wb, history, ctx, sheet, text, formula, calc };
}

const R = (a: string): Rect => parseRange(a)!;

describe("AutoAusfüllen as in Excel", () => {
  it("copies one number, counts with Strg", () => {
    const b = book({ A1: "1", B1: "1" });
    b.history.step("f", () => A.autoFill(b.ctx, b.sheet, R("A1"), R("A1:A5")));
    b.history.step("f", () => A.autoFill(b.ctx, b.sheet, R("B1"), R("B1:B5"), { series: true }));
    expect(["A2", "A5"].map(b.text)).toEqual(["1", "1"]);
    expect(["B2", "B5"].map(b.text)).toEqual(["2", "5"]);
  });

  it("continues two numbers, weekdays, months and numbered text", () => {
    const b = book({ A1: "1", A2: "3", B1: "Mo", C1: "Januar", D1: "Woche 1" });
    b.history.step("f", () => {
      A.autoFill(b.ctx, b.sheet, R("A1:A2"), R("A1:A6"));
      A.autoFill(b.ctx, b.sheet, R("B1"), R("B1:B5"));
      A.autoFill(b.ctx, b.sheet, R("C1"), R("C1:C4"));
      A.autoFill(b.ctx, b.sheet, R("D1"), R("D1:D4"));
    });
    expect(["A3", "A6"].map(b.text)).toEqual(["5", "11"]);
    expect(["B2", "B5"].map(b.text)).toEqual(["Di", "Fr"]);
    expect(b.text("C3")).toBe("März");
    expect(b.text("D4")).toBe("Woche 4");
  });

  it("moves relative references of formulas", () => {
    const b = book({ A1: "1", A2: "2", A3: "3", B1: "=A1*2" });
    b.history.step("f", () => A.autoFill(b.ctx, b.sheet, R("B1"), R("B1:B3")));
    expect(b.formula("B3")).toBe("=A3*2");
    expect(b.text("B3")).toBe("6");
  });
});

describe("rows, columns, copying, undo", () => {
  it("inserting and deleting rows moves cells and adjusts formulas, undo restores", () => {
    const b = book({ A1: "1", A2: "2", A3: "3", B1: "=SUMME(A1:A3)", B2: "=A2*10" });
    b.history.step("ins", () => A.insertDelete(b.ctx, b.sheet, "rows", 1, 1));
    expect(b.formula("B1")).toBe("=SUMME(A1:A4)");
    expect(b.formula("B3")).toBe("=A3*10");
    b.history.undo();
    expect(b.formula("B1")).toBe("=SUMME(A1:A3)");
    b.history.step("del", () => A.insertDelete(b.ctx, b.sheet, "rows", 1, -1));
    expect(b.formula("B1")).toBe("=SUMME(A1:A2)");
    expect(b.text("A2")).toBe("3");
    b.history.undo();
    b.history.redo();
    expect(b.text("A2")).toBe("3");
  });

  it("copying shifts references, moving keeps them and updates formulas pointing at the cells", () => {
    const b = book({ A1: "1", A2: "=A1+1", C1: "=A1" });
    b.history.step("copy", () => A.copyCells(b.ctx, { sheet: b.sheet, rect: R("A1:A2") }, b.sheet, { row: 0, col: 3 }));
    expect(b.formula("D2")).toBe("=D1+1");
    b.history.step("move", () => A.moveCells(b.ctx, b.sheet, R("A1"), { row: 4, col: 0 }));
    expect(b.formula("C1")).toBe("=A5");
  });

  it("Alt+Eingabe text turns on Textumbruch, Entf keeps formats", () => {
    const b = book({ A1: "a" + String.fromCharCode(10) + "b" });
    expect(b.wb.styles.get(b.sheet.get(0, 0)?.s).wrap).toBe(true);
    b.history.step("del", () => A.clearContents(b.ctx, b.sheet, [R("A1")]));
    expect(b.sheet.get(0, 0)?.v ?? null).toBe(null);
    expect(b.wb.styles.get(b.sheet.get(0, 0)?.s).wrap).toBe(true);
  });
});

describe("Einfügen and Datenüberprüfung", () => {
  it("takes the copied cells' validation along and replaces the target's; undo restores it", () => {
    const b = book({ A1: "Ja", B1: "x" });
    b.sheet.validations = [
      { ref: [R("A1:A3")], type: "list", formula1: '"Ja,Nein"', allowBlank: true },
      { ref: [R("C5:D6")], type: "custom", formula1: "ISNUMBER(C5)", allowBlank: true },
      { ref: [R("F1")], type: "whole", operator: "greaterThan", formula1: "0" },
    ];
    b.history.step("paste", () => A.copyCells(b.ctx, { sheet: b.sheet, rect: R("A1:A2") }, b.sheet, { row: 0, col: 5 }));
    const f = b.sheet.validations!.filter((v) => v.ref.some((r) => r.c0 === 5));
    expect(f).toHaveLength(1);
    expect(f[0]).toMatchObject({ type: "list", ref: [R("F1:F2")] });
    // A relative criterion moves with the cells.
    b.history.step("paste2", () => A.copyCells(b.ctx, { sheet: b.sheet, rect: R("C5") }, b.sheet, { row: 9, col: 7 }));
    expect(b.sheet.validations!.find((v) => v.ref[0].r0 === 9)?.formula1).toBe("ISNUMBER(H10)");
    b.history.undo();
    b.history.undo();
    expect(b.sheet.validations!.find((v) => v.type === "whole")?.ref).toEqual([R("F1")]);
  });
});

describe("AutoAusfüllen and Datenüberprüfung", () => {
  it("repeats the source's validation over the filled cells", () => {
    const b = book({ A1: "Ja", A2: "1" });
    b.sheet.validations = [{ ref: [R("A1")], type: "list", formula1: '"Ja,Nein"' }, { ref: [R("A2")], type: "whole", operator: "greaterThan", formula1: "0" }];
    b.history.step("f", () => A.autoFill(b.ctx, b.sheet, R("A1:A2"), R("A1:A5")));
    const list = b.sheet.validations!.filter((v) => v.type === "list").flatMap((v) => v.ref);
    const whole = b.sheet.validations!.filter((v) => v.type === "whole").flatMap((v) => v.ref);
    expect(list).toEqual(expect.arrayContaining([R("A1"), R("A3"), R("A5")]));
    expect(whole).toEqual(expect.arrayContaining([R("A2"), R("A4")]));
  });
});
