// Structured references to Excel tables (D-080): parsing, resolving, computing, the file's form.

import { describe, expect, it } from "vitest";
import { Calc, Sheet, Workbook, parseCell, parseStruct, structToFile, structToShown, toFileFormula, toLocalFormula, type Value } from "../src";

function book(): Workbook {
  const wb = new Workbook();
  const sh = new Sheet("Tabelle1");
  wb.sheets.push(sh);
  const rows: (string | number)[][] = [["Name", "Note", "Punkte"], ["Ada", 2, 12], ["Ben", 3, 9], ["Cem", 1, 14]];
  rows.forEach((r, i) => r.forEach((v, k) => sh.set(i, k, { v })));
  sh.tables = [{
    name: "Noten", ref: { r0: 0, c0: 0, r1: 4, c1: 2 }, header: true, totals: true, style: "TableStyleMedium2",
    columns: [{ name: "Name", totalsLabel: "Ergebnis" }, { name: "Note" }, { name: "Punkte", totalsFunction: "sum" }],
    firstColumn: false, lastColumn: false, rowStripes: true, columnStripes: false, filterButton: true,
  }];
  sh.set(4, 0, { v: "Ergebnis" });
  return wb;
}

function calc(wb: Workbook, cells: Record<string, string>): (a: string) => Value {
  for (const [a, f] of Object.entries(cells)) {
    const c = parseCell(a)!;
    wb.sheets[0].set(c.row, c.col, { v: null, f: toFileFormula(f) });
  }
  const calc = new Calc(wb);
  calc.recalc();
  return (a) => {
    const c = parseCell(a)!;
    return wb.sheets[0].get(c.row, c.col)?.v ?? null;
  };
}

describe("structured references", () => {
  it("reads every form", () => {
    expect(parseStruct("Noten[Note]")).toMatchObject({ table: "Noten", col0: "Note", col1: null });
    expect([...parseStruct("[@Note]")!.items]).toEqual(["thisrow"]);
    expect(parseStruct("Noten[[#Kopfzeilen];[Name]:[Note]]")).toMatchObject({ col0: "Name", col1: "Note" });
    expect([...parseStruct("Noten[[#Headers],[#Data],[Note]]")!.items].sort()).toEqual(["data", "headers"]);
    expect(parseStruct("Noten[@[Note 1]]")).toMatchObject({ col0: "Note 1" });
    expect(parseStruct("Noten[Preis '[€']]")).toMatchObject({ col0: "Preis [€]" });
    expect(parseStruct("Noten[[A],[B]]")).toBeNull();
  });

  it("computes as Excel does", () => {
    const v = calc(book(), {
      E1: "=SUMME(Noten[Note])",
      E2: "=ZEILEN(Noten[#Alle])",
      E3: "=ANZAHL2(Noten[[#Kopfzeilen];[Name]])",
      E4: "=SUMME(Noten[[Note]:[Punkte]])",
      E5: "=SUMME(Noten)",
      E6: "=Noten[Fehlt]",
      E7: "=ZEILEN(Noten[#Daten])",
      E8: "=SUMME(Noten[[#Ergebnisse];[Punkte]])",
      F2: "=Noten[@Punkte]*2",
      F1: "=Noten[@Punkte]",
    });
    expect(v("E1")).toBe(6);
    expect(v("E2")).toBe(5);
    expect(v("E3")).toBe(1);
    expect(v("E4")).toBe(41);
    expect(v("E5")).toBe(41);
    expect(String(v("E6"))).toBe("#REF!");
    expect(v("E7")).toBe(3);
    expect(v("E8")).toBe(0);
    expect(v("F2")).toBe(24);
    expect(String(v("F1"))).toBe("#VALUE!");
  });

  it("follows changes of the cells it reads", () => {
    const wb = book();
    wb.sheets[0].set(1, 3, { v: null, f: toFileFormula("=Noten[@Note]*10") });
    const c = new Calc(wb);
    c.recalc();
    expect(wb.sheets[0].get(1, 3)!.v).toBe(20);
    wb.sheets[0].set(1, 1, { v: 5 });
    c.changed(0, [{ row: 1, col: 1 }]);
    expect(wb.sheets[0].get(1, 3)!.v).toBe(50);
  });

  it("is written as Excel writes it and shown as Excel shows it", () => {
    expect(toFileFormula("=SUMME(Noten[[#Kopfzeilen];[Note]])")).toBe("SUM(Noten[[#Headers],[Note]])");
    expect(toLocalFormula("SUM(Noten[[#Headers],[Note]])")).toBe("SUMME(Noten[[#Kopfzeilen];[Note]])");
    expect(structToFile("[@Note]", "Noten")).toBe("Noten[[#This Row],[Note]]");
    expect(structToFile("[@[Note 1]]", "Noten")).toBe("Noten[[#This Row],[Note 1]]");
    expect(structToShown("Noten[[#This Row],[Note]]", "Noten")).toBe("[@Note]");
    expect(structToShown("Noten[[#This Row],[Note 1]]", "Noten")).toBe("[@[Note 1]]");
    expect(structToShown("Noten[[#This Row],[Note]]", null)).toBe("Noten[@Note]");
  });
});
