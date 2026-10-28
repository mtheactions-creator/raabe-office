// RaabeCalc's cell operations: AutoAusfüllen as recorded in Excel (catalogue_calc.py), inserting
// and deleting rows with formulas following, copying with shifted references, undo and redo.

import { readFileSync } from "node:fs";
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

describe("Umschalt+drag inserts the cells (recorded in Excel)", () => {
  const cases: [Record<string, string>, string, { row: number; col: number }, Record<string, string>][] = [
    [{ A1: "1", A2: "2", A3: "3", A4: "4", A5: "5" }, "A1", { row: 3, col: 0 }, { A1: "2", A2: "3", A3: "1", A4: "4", A5: "5" }],
    [{ A1: "1", B1: "2", C1: "3", B2: "x" }, "A1", { row: 1, col: 1 }, { B1: "2", C1: "3", B2: "1", B3: "x" }],
    [{ A1: "1", B1: "2", C1: "3", D1: "4" }, "A1", { row: 1, col: 2 }, { B1: "2", C1: "3", D1: "4", C2: "1" }],
    [{ A1: "1", B1: "2", A2: "3", B2: "4", A3: "5" }, "A1", { row: 2, col: 0 }, { A1: "3", B1: "2", A2: "1", B2: "4", A3: "5" }],
    [{ A1: "1", A2: "2", A3: "3", A4: "4" }, "A4", { row: 1, col: 0 }, { A1: "1", A2: "4", A3: "2", A4: "3" }],
    [{ A1: "1", A2: "2", B1: "3", B2: "4", C1: "5", C2: "6" }, "A1:A2", { row: 1, col: 1 }, { B1: "3", C1: "5", B2: "1", C2: "6", B3: "2", B4: "4" }],
  ];
  for (const [setup, src, at, want] of cases) {
    it(`${src} to row ${at.row + 1}, column ${at.col + 1}`, () => {
      const b = book(setup);
      b.history.step("m", () => A.moveInsert(b.ctx, b.sheet, R(src), at, true));
      const got: Record<string, string> = {};
      for (const [k, c] of b.sheet.cells) if (c.v !== null) got[`${String.fromCharCode(65 + (k % 16384))}${Math.floor(k / 16384) + 1}`] = String(c.v);
      expect(got).toEqual(want);
    });
  }
});

describe("References follow inserted, deleted and insert-moved cells", () => {
  it("Umschalt+drag: formulas point to the moved and the shifted cells", () => {
    const b = book({ A1: "1", A2: "2", A3: "3", B1: "=A1*2", C1: "=A2" });
    b.history.step("m", () => A.moveInsert(b.ctx, b.sheet, R("A1"), { row: 3, col: 0 }, true));
    expect(b.formula("B1")).toBe("=A3*2");
    expect(b.formula("C1")).toBe("=A1");
  });
  it("Zellen einfügen (shift down) moves references", () => {
    const b = book({ A1: "1", A2: "2", B1: "=A2" });
    b.history.step("i", () => A.shiftCells(b.ctx, b.sheet, R("A1"), "down"));
    expect(b.formula("B1")).toBe("=A3");
  });
});

describe("Right-drag fills", () => {
  const recorded = JSON.parse(readFileSync("tools/feel-lab/results/2026-10-28-excel-autofill-types.json", "utf-8")) as { kind: string; source: number[]; values: number[] }[];
  for (const c of recorded) {
    it(`${c.kind} from ${c.source.join(", ")} as Excel`, () => {
      const b = book(Object.fromEntries(c.source.map((v, i) => [`A${i + 1}`, String(v)])));
      b.history.step("f", () => A.fillTrend(b.ctx, b.sheet, R(`A1:A${c.source.length}`), R(`A1:A${c.values.length}`), c.kind === "growth"));
      c.values.forEach((v, i) => expect(b.sheet.get(i, 0)?.v as number).toBeCloseTo(v, 9));
    });
  }
  it("Monate ausfüllen keeps the day or takes the month's last (EDATUM)", () => {
    const b = book({ A1: "31.01.2026" });
    b.history.step("f", () => A.fillDates(b.ctx, b.sheet, R("A1"), R("A1:A4"), "months"));
    expect([1, 2, 3].map((i) => b.text(`A${i + 1}`))).toEqual(["28.02.2026", "31.03.2026", "30.04.2026"]);
  });
  it("Arbeitstage ausfüllen skips weekends", () => {
    const b = book({ A1: "30.01.2026" });
    b.history.step("f", () => A.fillDates(b.ctx, b.sheet, R("A1"), R("A1:A3"), "workdays"));
    expect([b.text("A2"), b.text("A3")]).toEqual(["02.02.2026", "03.02.2026"]);
  });
});

describe("Teilergebnis as Excel inserts it", () => {
  const rec = JSON.parse(readFileSync("tools/feel-lab/results/2026-10-28-excel-subtotal.json", "utf-8")) as Record<string, { cells: string[]; level: number; bold: boolean[] }[]>;
  const FN: Record<string, [string, number[]]> = {
    "Summe Punkte": ["Summe", [3]], "Mittelwert Note und Punkte": ["Mittelwert", [2, 3]], "Anzahl Note": ["Anzahl", [2]], "Max Punkte": ["Maximum", [3]],
    "Min Punkte": ["Minimum", [3]], "Produkt Note": ["Produkt", [2]], "Anzahl Zahlen": ["Anzahl Zahlen", [2]], StdAbw: ["Standardabweichung (Stichprobe)", [3]],
    StdAbwN: ["Standardabweichung (Grundgesamtheit)", [3]], Varianz: ["Varianz (Stichprobe)", [3]], VarianzN: ["Varianz (Grundgesamtheit)", [3]],
  };
  for (const [name, rows] of Object.entries(rec)) {
    it(name, () => {
      const data = [["Klasse", "Name", "Note", "Punkte"], ["7a", "Anna", "2", "11"], ["7a", "Ben", "3", "8"], ["7b", "Cem", "1", "14"], ["7b", "Dora", "4", "6"], ["7b", "Emil", "2", "10"], ["8a", "Finn", "3", "9"]];
      const setup: Record<string, string> = {};
      data.forEach((r, i) => r.forEach((v, j) => { setup[`${"ABCD"[j]}${i + 1}`] = v; }));
      const b = book(setup);
      const [fname, cols] = FN[name];
      b.history.step("t", () => A.subtotal(b.ctx, b.sheet, R("A1:D7"), 0, A.SUBTOTALS.find((f) => f.name === fname)!, cols));
      rows.forEach((row, r) => {
        expect((b.sheet.rows.get(r)?.outline ?? 0) + 1, `level of row ${r + 1}`).toBe(row.level);
        row.cells.forEach((want, c) => {
          const cell = b.sheet.get(r, c);
          const got = cell?.f !== undefined ? formulaLocal(cell) : cell?.v === null || cell?.v === undefined ? "" : String(cell.v);
          expect(got, `${"ABCD"[c]}${r + 1}`).toBe(want);
          expect(!!b.wb.styles.get(cell?.s).font.bold, `bold ${"ABCD"[c]}${r + 1}`).toBe(row.bold[c]);
        });
      });
    });
  }
});

describe("Undo of whole-sheet steps restores the sheet's settings", () => {
  it("page setup, filter, protection and threads come back", () => {
    const b = book({ A1: "1" });
    b.history.step("x", () => {
      b.history.whole(b.sheet);
      b.sheet.page = { orientation: "landscape" };
      b.sheet.autoFilter = { ref: R("A1:A3"), columns: {} };
      b.sheet.protection = { sheet: "1" };
      b.sheet.threads = [{ row: 0, col: 0, id: "{1}", author: "A", date: "", text: "t", replies: [] }];
    });
    b.history.undo();
    expect([b.sheet.page, b.sheet.autoFilter, b.sheet.protection, b.sheet.threads]).toEqual([undefined, undefined, undefined, undefined]);
    b.history.redo();
    expect(b.sheet.page?.orientation).toBe("landscape");
    expect(b.sheet.threads?.[0].text).toBe("t");
  });
});

describe("Review fixes", () => {
  it("Kommentare and pictures move with inserted and deleted rows", () => {
    const b = book({ A1: "1" });
    b.sheet.threads = [{ row: 4, col: 1, id: "{1}", author: "A", date: "", text: "B5", replies: [] }, { row: 1, col: 0, id: "{2}", author: "A", date: "", text: "A2", replies: [] }];
    b.sheet.pictures = [{ from: { col: 2, row: 5, x: 3, y: 4 }, to: { col: 4, row: 8, x: 0, y: 0 }, src: "data:image/png;base64,", name: "P" }];
    b.history.step("i", () => A.insertDelete(b.ctx, b.sheet, "rows", 2, 1));
    expect(b.sheet.threads!.map((t) => [t.text, t.row])).toEqual([["B5", 5], ["A2", 1]]);
    expect([b.sheet.pictures![0].from.row, b.sheet.pictures![0].to!.row]).toEqual([6, 9]);
    b.history.step("d", () => A.insertDelete(b.ctx, b.sheet, "rows", 5, -1));
    expect(b.sheet.threads!.map((t) => t.text)).toEqual(["A2"]);
  });
  it("Tage ausfüllen continues by the step of two dates", () => {
    const b = book({ A1: "01.01.2026", A2: "03.01.2026" });
    b.history.step("f", () => A.fillDates(b.ctx, b.sheet, R("A1:A2"), R("A1:A4"), "days"));
    expect([b.text("A3"), b.text("A4")]).toEqual(["05.01.2026", "07.01.2026"]);
  });
  it("Arbeitstage ausfüllen counts workdays between two dates (Fri, Mon: one)", () => {
    const b = book({ A1: "30.01.2026", A2: "02.02.2026" });
    b.history.step("f", () => A.fillDates(b.ctx, b.sheet, R("A1:A2"), R("A1:A3"), "workdays"));
    expect(b.text("A3")).toBe("03.02.2026");
  });
});
