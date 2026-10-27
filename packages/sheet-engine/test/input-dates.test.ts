// Date serial arithmetic (dates.ts) against Excel's own DATE, YEAR, …, EDATE, EOMONTH on edge
// cases: years 0-1899, month and day overflow both ways, the fictitious 29 Feb 1900, the last
// day 31.12.9999, the 1904 system (tools/feel-lab/probe_excel_input.py).
import { readFileSync } from "node:fs";
import { describe, expect, it } from "vitest";
import { addMonths, dateToSerial, day, endOfMonth, hour, minute, month, second, serialToDate, weekday, year } from "../src/index";

interface Case { fn: string; args: number[]; value: number | string | null }
const fx = JSON.parse(readFileSync("packages/sheet-engine/test/fixtures/excel-input.json", "utf-8")) as {
  dateFunctions: Case[]; dateFunctions1904: Case[];
};

function ours(c: Case, d1904: boolean): number | string {
  const [a, b, d] = c.args;
  let v: number;
  switch (c.fn) {
    case "DATE": v = dateToSerial(a, b, d, d1904); break;
    case "YEAR": v = year(a, d1904); break;
    case "MONTH": v = month(a, d1904); break;
    case "DAY": v = day(a, d1904); break;
    case "HOUR": v = hour(a, d1904); break;
    case "MINUTE": v = minute(a, d1904); break;
    case "SECOND": v = second(a, d1904); break;
    case "WEEKDAY": v = weekday(a, b, d1904); break;
    case "EDATE": v = addMonths(a, b, d1904); break;
    case "EOMONTH": v = endOfMonth(a, b, d1904); break;
    default: throw new Error(c.fn);
  }
  return isNaN(v) ? "error" : v;
}

describe("date serials against Excel", () => {
  for (const [name, cases, d1904] of [["1900", fx.dateFunctions, false], ["1904", fx.dateFunctions1904, true]] as const) {
    it(`computes Excel's date functions (${name} system)`, () => {
      const misses: string[] = [];
      for (const c of cases) {
        const v = ours(c, d1904);
        if (v !== c.value) misses.push(`${c.fn}(${c.args.join(";")}): Excel ${c.value}, ours ${v}`);
      }
      console.log(`dates ${name}: ${cases.length - misses.length}/${cases.length} match`);
      for (const m of misses.slice(0, 40)) console.log("  " + m);
      expect(misses).toEqual([]);
    });
  }

  it("serial 0, the fictitious 29 Feb 1900 and 1 Mar 1900", () => {
    expect(serialToDate(0)).toEqual({ y: 1900, m: 1, d: 0 });
    expect(serialToDate(60)).toEqual({ y: 1900, m: 2, d: 29 });
    expect(serialToDate(61)).toEqual({ y: 1900, m: 3, d: 1 });
    expect(serialToDate(0, true)).toEqual({ y: 1904, m: 1, d: 1 });
  });
});
