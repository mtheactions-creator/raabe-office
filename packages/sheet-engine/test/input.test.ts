// Replays what German Excel made of typed input (tools/feel-lab/probe_excel_input.py), first in
// empty cells, then in cells formatted beforehand. The probe assigns FormulaLocal, which parses
// like typing except for two documented typing features (automatic percent entry, Lotus-style
// "+1+2" formulas): those are switched off here and tested separately below.
import { readFileSync } from "node:fs";
import { describe, expect, it } from "vitest";
import { CellError, LOCALE_DISPLAY, editText, numberEditText, parseInput, type ErrorCode, type InputOptions } from "../src/index";

interface Rec {
  in: string; rejected?: boolean; value?: unknown; type?: string; formula?: boolean; en?: string; local?: string; text?: string;
  error?: number; f?: string; fl?: string;
}
interface Fixture { date: string; general: Rec[]; formatted: Record<string, Rec[]> }

const fx = JSON.parse(readFileSync("packages/sheet-engine/test/fixtures/excel-input.json", "utf-8")) as Fixture;
const OPTS: InputOptions = { lotusFormulas: false, autoPercent: false, today: new Date(2026, 9, 27) };

const ERRORS: Record<number, ErrorCode> = {
  [-2146826281]: "#DIV/0!", [-2146826246]: "#N/A", [-2146826259]: "#NAME?", [-2146826288]: "#NULL!",
  [-2146826252]: "#NUM!", [-2146826265]: "#REF!", [-2146826273]: "#VALUE!",
};

/** Inputs where we differ from Excel, with the reason. */
const KNOWN = new Map<string, string>([
  ["1 -1/2", "Excel reads it as the date 1.1.2002; we keep it as text"],
  ["١٢", "Arabic-Indic digits: Excel reads 12 with a script format; we keep text"],
]);

/** Excel reports the built-in currency formats in their English form; ours is the file's code. */
const fileCode = (en: string) => unescape(en.startsWith("$") || en.startsWith("_(") ? LOCALE_DISPLAY[en] ?? en : en);
/** Excel's English form leaves out escapes of spaces and minus signs that the file has. */
const unescape = (code: string) => code.replace(/\\([ -])/g, "$1");

function check(r: Rec, current: string): string | null {
  const ours = parseInput(r.in, { numFmt: current }, OPTS);
  if (r.rejected) return ours.formula !== undefined ? null : `Excel rejects it as a formula, ours ${JSON.stringify(ours)}`;
  if (r.formula) return ours.formula !== undefined && ours.value === null ? null : `Excel formula ${r.fl}, ours ${JSON.stringify(ours)}`;
  if (ours.formula !== undefined) return `ours a formula, Excel ${JSON.stringify(r.value)}`;
  const v = ours.value;
  if (r.type === "str") { if (v !== r.value) return `value ${JSON.stringify(v)} vs Excel ${JSON.stringify(r.value)}`; }
  else if (r.type === "bool") { if (v !== r.value) return `value ${JSON.stringify(v)} vs Excel ${r.value}`; }
  else if (r.type === "error") { if (!(v instanceof CellError) || v.code !== ERRORS[r.error!]) return `value ${String(v)} vs Excel error ${r.error}`; }
  else if (typeof v !== "number" || Math.abs(v - (r.value as number)) > 1e-9 * Math.max(1, Math.abs(r.value as number))) {
    return `value ${JSON.stringify(v)} vs Excel ${r.value}`;
  }
  const expected = fileCode(r.en!);
  const got = ours.numFmt ?? current;
  if (fileCode(got) !== expected) return `format ${JSON.stringify(got)} vs Excel ${JSON.stringify(expected)}`;
  return null;
}

describe("typed input against Excel", () => {
  it("parses input in empty (Standard) cells as Excel does", () => {
    const misses: string[] = [];
    let total = 0;
    for (const r of fx.general) {
      if (KNOWN.has(r.in)) continue;
      total++;
      const m = check(r, "General");
      if (m) misses.push(`${JSON.stringify(r.in)}: ${m}`);
    }
    console.log(`input (Standard): ${total - misses.length}/${total} match`);
    for (const m of misses) console.log("  " + m);
    expect(misses).toEqual([]);
  });

  it("keeps or replaces the cell's format as Excel does", () => {
    const misses: string[] = [];
    let total = 0;
    for (const [fmt, recs] of Object.entries(fx.formatted)) {
      for (const r of recs) {
        if (KNOWN.has(r.in)) continue;
        total++;
        const m = check(r, fmt);
        if (m) misses.push(`${fmt} ${JSON.stringify(r.in)}: ${m}`);
      }
    }
    console.log(`input (formatted cells): ${total - misses.length}/${total} match`);
    for (const m of misses) console.log("  " + m);
    expect(misses).toEqual([]);
  });
});

describe("typing features FormulaLocal does not show", () => {
  const today = new Date(2026, 9, 27);
  it("automatic percent entry in percent cells (Excel's documented rule)", () => {
    expect(parseInput("15", { numFmt: "0%" }, { today }).value).toBe(0.15);
    expect(parseInput("0,5", { numFmt: "0%" }, { today }).value).toBe(0.5);
    expect(parseInput("15%", { numFmt: "0%" }, { today }).value).toBe(0.15);
    expect(parseInput("15", { numFmt: "0%" }, { today, autoPercent: false }).value).toBe(15);
  });
  it("Lotus-style formulas starting with + or -", () => {
    expect(parseInput("+1+2")).toEqual({ value: null, formula: "=+1+2" });
    expect(parseInput("-A1")).toEqual({ value: null, formula: "=-A1" });
    expect(parseInput("-5")).toEqual({ value: -5 });
    expect(parseInput("-")).toEqual({ value: "-" });
  });
  it("text cells, apostrophes, formulas", () => {
    expect(parseInput("1,5", { numFmt: "@" })).toEqual({ value: "1,5" });
    expect(parseInput("'1.1.2026")).toEqual({ value: "1.1.2026" });
    expect(parseInput("=SUMME(A1:A3)")).toEqual({ value: null, formula: "=SUMME(A1:A3)" });
    expect(parseInput("")).toEqual({ value: null });
  });
  it("dates in the 1904 system", () => {
    expect(parseInput("1.1.1904", undefined, { date1904: true }).value).toBe(0);
    expect(parseInput("2.1.1904", undefined, { date1904: true })).toEqual({ value: 1, numFmt: "m/d/yyyy" });
  });
});

describe("the text to edit (Bearbeitungsleiste)", () => {
  const ed = JSON.parse(readFileSync("packages/sheet-engine/test/fixtures/excel-input.json", "utf-8")) as {
    editValues: number[]; editTexts: { code: string; edit: string[] }[];
  };
  it("writes numbers as Excel does for editing (15 digits, 1E+21)", () => {
    const misses: string[] = [];
    let total = 0;
    for (const f of ed.editTexts) {
      ed.editValues.forEach((v, i) => {
        total++;
        const ours = numberEditText(v);
        if (ours !== f.edit[i]) misses.push(`${v}: Excel ${f.edit[i]}, ours ${ours}`);
      });
    }
    console.log(`edit texts: ${total - misses.length}/${total} match`);
    expect(misses).toEqual([]);
  });
  it("gives back the value when the edit text is typed again", () => {
    const formats = ["General", "0.00", "0%", "0.00%", "m/d/yyyy", "m/d/yyyy h:mm", "h:mm", "h:mm:ss", "[h]:mm:ss", "# ?/?", "0.00E+00"];
    const values = [0, 1, -1, 0.15, 1234.5, 45000.5, 46023, 46023.520833333336, 0.5208333333333334, 1.0416666666666667, 1 / 3, 1e20, 1e-20];
    for (const f of formats) {
      for (const v of values) {
        const t = editText(v, f);
        const back = parseInput(t, { numFmt: f }, { autoPercent: false });
        expect(typeof back.value, `${f} ${v} "${t}"`).toBe("number");
        expect(Math.abs((back.value as number) - v), `${f} ${v} "${t}"`).toBeLessThan(1e-6 * Math.max(1, Math.abs(v)));
      }
    }
    expect(editText("1,5")).toBe("'1,5");
    expect(editText("abc")).toBe("abc");
    expect(editText("1,5", "@")).toBe("1,5");
    expect(editText(true)).toBe("WAHR");
    expect(editText(46023, "m/d/yyyy")).toBe("01.01.2026");
    expect(editText(0.15, "0%")).toBe("15%");
  });
});
