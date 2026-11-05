// Excel's checks when a formula is typed, as recorded in Excel 365 with alerts on (D-066).
import { describe, expect, it } from "vitest";
import { checkEntry } from "../src/formula/entry";

describe("checkEntry", () => {
  const cases: [string, "ok" | "propose" | "error", string?][] = [
    ["=1+", "propose", "=1"],
    ['="a', "propose", '="a"'],
    ["=(1+2", "propose", "=(1+2)"],
    ["=1)", "propose", "=1"],
    ["=1+*2", "propose", "=1+2"],
    ["=1*/2", "propose", "=1*2"],
    ["=SUMME(1;2))", "propose", "=SUMME(1;2)"],
    ["=2(3)", "propose", "=2*(3)"],
    ["=1+2+", "propose", "=1+2"],
    ["=1-", "propose", "=1"],
    ["=((1)", "propose", "=((1))"],
    ["=1+(2", "propose", "=1+(2)"],
    ["=1,5+", "propose", "=1,5"],
    ["=1^", "propose", "=1"],
    ["=1&", "propose", "=1"],
    ["=1=", "propose", "=1"],
    ["=1 2", "propose", "=12"],
    ["=SUMME(1 2)", "propose", "=SUMME(12)"],
    ["=SUMME(1;2", "ok", "=SUMME(1;2)"],
    ["=SUMME(1;", "ok", "=SUMME(1;)"],
    ["=SUMME(1;2;)", "ok", "=SUMME(1;2;)"],
    ['=WENN(A1>1;"ja";"nein"', "ok", '=WENN(A1>1;"ja";"nein")'],
    ["=1++2", "ok", "=1+2"],
    ["=A1+B", "ok", "=A1+B"],
    ["=1%%", "ok", "=1%%"],
    ["=--1", "ok", "=--1"],
    ["=SUMME", "ok", "=SUMME"],
    ["=SUMME(1;;2)", "ok", "=SUMME(1;;2)"],
    ["=SUMME(A1:)", "error"],
    ["=SUMME(", "error"],
    ["=1;2", "error"],
    ["=SUMME()", "error"],
    ["=A1:", "error"],
    ["=+", "error"],
    ["=-", "error"],
    ["=MITTELWERT(", "error"],
    ["=ABS()", "error"],
    ["=ABS(1;2)", "error"],
    ["=WENN()", "error"],
    ["=WENN(1)", "error"],
    ["=PI(1)", "error"],
    ["=HEUTE(1)", "error"],
  ];
  for (const [typed, kind, text] of cases) {
    it(typed, () => {
      const r = checkEntry(typed);
      expect(r.kind).toBe(kind);
      if (text !== undefined && r.kind !== "error") expect(r.text).toBe(text);
    });
  }
  it("leaves text alone", () => expect(checkEntry("abc")).toEqual({ kind: "ok", text: "abc" }));
});
