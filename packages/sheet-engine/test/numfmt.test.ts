// Replays what Excel showed for values in number formats (tools/feel-lab/probe_excel_numfmt.py):
// every recorded case is formatted here and compared with Excel's text. The few cases where we
// differ on purpose or have not matched Excel yet are listed in KNOWN with the reason.
import { readFileSync } from "node:fs";
import { describe, expect, it } from "vitest";
import {
  BUILTIN_FORMATS, FORMAT_CATEGORIES, RIBBON_FORMATS, builtinId, err, fileFormat, formatValue, localFormat,
  type ErrorCode, type FormatOptions, type Value,
} from "../src/index";

interface FormatRec { code: string; invalid?: boolean; en?: string; local?: string; texts?: string[]; other?: string[] }
interface LocalRec { local: string; invalid?: boolean; en?: string; back?: string; text?: string; file?: { id: number; code: string | null } }
interface Fixture {
  measure: { dpi: number; widths: Record<string, number> };
  values: number[];
  other: string[];
  formats: FormatRec[];
  widths: { columns: { width: number; points: number }[]; values: number[]; formats: { code: string; texts: string[][] }[]; strings: string[]; stringTexts: string[][] };
  local: LocalRec[];
  categories: Record<string, LocalRec[]>;
  date1904: { values: number[]; formats: FormatRec[] };
  random?: { values: number[]; formats: FormatRec[] };
  randomWidths?: { columns: { width: number; points: number }[]; texts: string[][] };
  fractionLimits?: { values: number[]; formats: FormatRec[] };
  builtin: { read: Record<string, { en: string; local: string; text: string; neg: string }>; written: Record<string, { id: number; code: string | null }> };
}

const fx = JSON.parse(readFileSync("packages/sheet-engine/test/fixtures/excel-numfmt.json", "utf-8")) as Fixture;
const W = fx.measure.widths;
/** Calibri 11 as GDI measures it at the probe's DPI (each character's advance). */
const measure = (s: string) => [...s].reduce((a, c) => a + (W[c] ?? W["0"]), 0);
const px = (points: number) => Math.round((points * fx.measure.dpi) / 72);
/** The probe's wide columns: 120 characters of Calibri 11. */
const WIDE: FormatOptions = { widthPx: Math.round(120 * W["0"]) + Math.floor((W["0"] * 5) / 7), measure };

const OTHER: Value[] = ["Text", "", true, false, err("#N/A"), err("#DIV/0!"), err("#VALUE!"), err("#NUM!"), "12", "-5"];

/** Cases we do not match, with the reason (kept short; see the report in DECISIONS). */
const KNOWN = new Map<string, string>([
  ["General dd/mm", "Excel's own oddity: General inside a date format shows a fraction of the value"],
]);

/** German forms we do not reproduce: the built-in currency codes set through COM become custom
 * formats in Excel and keep their quotes ('#.##0 "€"'), while the same ids read from a file show
 * "€" unquoted (the form we use); Excel's own rewriting of "0,0" and ";0". */
const KNOWN_LOCAL = new Set([
  ...[5, 6, 7, 8, 41, 42, 43, 44].map((i) => BUILTIN_FORMATS[i]), "0,0", ";0",
]);

interface Miss { where: string; code: string; value: unknown; excel: string; ours: string }

/** Compares one result with Excel's text: "####" by the hashes flag, fills by their ends. */
function same(excel: string, f: { text: string; fill?: string; fillAt?: number; hashes?: boolean }): boolean {
  if (f.hashes) return /^#*$/.test(excel);
  if (/^#+$/.test(excel)) return f.text === excel;
  if (f.fill !== undefined) {
    const a = f.text.slice(0, f.fillAt), b = f.text.slice(f.fillAt);
    if (!excel.startsWith(a) || !excel.endsWith(b) || excel.length < a.length + b.length) return false;
    const mid = excel.slice(a.length, excel.length - b.length);
    return [...mid].every((c) => c === f.fill);
  }
  return excel === f.text;
}

function report(name: string, misses: Miss[], total: number): void {
  const pct = (((total - misses.length) / total) * 100).toFixed(2);
  console.log(`${name}: ${total - misses.length}/${total} match (${pct} %)`);
  for (const m of misses.slice(0, 60)) console.log(`  ${m.where} ${JSON.stringify(m.code)} ${JSON.stringify(m.value)}: Excel ${JSON.stringify(m.excel)} ours ${JSON.stringify(m.ours)}`);
}

describe("number formats against Excel", () => {
  it("formats every recorded value in every recorded format (wide column)", () => {
    const misses: Miss[] = [];
    let total = 0, known = 0;
    for (const f of fx.formats) {
      if (f.invalid || KNOWN.has(f.code)) continue;
      fx.values.forEach((v, i) => {
        const excel = f.texts![i];
        // Known: Excel shows numbers longer than the probe's 120-character column (1,5E+100
        // with thousands separators, 134-140 characters) instead of "####"; not reproduced.
        if (Math.abs(v) === 1.5e100 && excel.length > 120 && !/^#+$/.test(excel)) { known++; return; }
        total++;
        const r = formatValue(v, f.code, WIDE);
        if (!same(excel, r)) misses.push({ where: "value", code: f.code, value: v, excel, ours: r.text });
      });
      OTHER.forEach((v, i) => {
        total++;
        const r = formatValue(v, f.code, WIDE);
        if (!same(f.other![i], r)) misses.push({ where: "other", code: f.code, value: fx.other[i], excel: f.other![i], ours: r.text });
      });
    }
    report("wide", misses, total);
    console.log(`  known differences left out: ${known}`);
    expect(misses.length).toBeLessThanOrEqual(0);
  });

  it("shortens General and shows #### in narrow columns as Excel does", () => {
    const misses: Miss[] = [];
    let total = 0, known = 0;
    const cols = fx.widths.columns;
    for (const f of fx.widths.formats) {
      fx.widths.values.forEach((v, i) => {
        cols.forEach((c, j) => {
          total++;
          const r = formatValue(v, f.code, { widthPx: px(c.points), measure });
          const excel = f.texts[i][j];
          // Known: a negative number in "General;-General" (and in a Text cell, which shows it
          // the same way) in a one-character column: Excel shows "#" for some values and "0"
          // for others (-1, -12,5); we always show "0".
          if (c.width < 2 && v < 0 && excel === "#" && r.text === "0") { known++; return; }
          if (!same(excel, r) || (r.hashes && excel !== r.text)) misses.push({ where: `w=${c.width}`, code: f.code, value: v, excel, ours: r.text });
        });
      });
    }
    fx.widths.strings.forEach((s, i) => {
      // (The first recording left the last format, ";;;", on these cells: no texts to compare.)
      if (fx.widths.stringTexts[i].every((t) => t === "")) return;
      cols.forEach((c, j) => {
        total++;
        // Excel took the digits-only string for a number when the probe assigned it.
        const r = formatValue(/^\d+$/.test(s) ? Number(s) : s, "General", { widthPx: px(c.points), measure });
        if (r.text !== fx.widths.stringTexts[i][j]) misses.push({ where: `w=${c.width}`, code: "General", value: s, excel: fx.widths.stringTexts[i][j], ours: r.text });
      });
    });
    report("widths", misses, total);
    console.log(`  known differences left out: ${known}`);
    expect(misses.length).toBeLessThanOrEqual(0);
  });

  it("formats random numbers of every size as Excel does", () => {
    const misses: Miss[] = [];
    let total = 0;
    const rnd = fx.random!;
    for (const f of rnd.formats) {
      if (f.invalid) continue;
      rnd.values.forEach((v, i) => {
        total++;
        const r = formatValue(v, f.code, WIDE);
        if (!same(f.texts![i], r)) misses.push({ where: "random", code: f.code, value: v, excel: f.texts![i], ours: r.text });
      });
    }
    const cols = fx.randomWidths!.columns;
    rnd.values.forEach((v, i) => {
      cols.forEach((c, j) => {
        total++;
        const excel = fx.randomWidths!.texts[i][j];
        const r = formatValue(v, "General", { widthPx: px(c.points), measure });
        if (excel !== r.text) misses.push({ where: `random w=${c.width}`, code: "General", value: v, excel, ours: r.text });
      });
    });
    report("random", misses, total);
    expect(misses.length).toBeLessThanOrEqual(0);
  });

  it("knows where fractions turn into ####", () => {
    const misses: Miss[] = [];
    let total = 0;
    const fl = fx.fractionLimits!;
    for (const f of fl.formats) {
      fl.values.forEach((v, i) => {
        total++;
        const r = formatValue(v, f.code, WIDE);
        if (!same(f.texts![i], r)) misses.push({ where: "fraction", code: f.code, value: v, excel: f.texts![i], ours: r.text });
      });
    }
    report("fraction limits", misses, total);
    expect(misses.length).toBeLessThanOrEqual(0);
  });

  it("formats dates in the 1904 system", () => {
    const misses: Miss[] = [];
    let total = 0;
    for (const f of fx.date1904.formats) {
      if (f.invalid || KNOWN.has(f.code)) continue;
      fx.date1904.values.forEach((v, i) => {
        total++;
        const r = formatValue(v, f.code, { ...WIDE, date1904: true });
        if (!same(f.texts![i], r)) misses.push({ where: "1904", code: f.code, value: v, excel: f.texts![i], ours: r.text });
      });
    }
    report("1904", misses, total);
    expect(misses.length).toBeLessThanOrEqual(0);
  });

  it("reads the built-in format ids as German Excel does", () => {
    const misses: Miss[] = [];
    let total = 0;
    const wide: FormatOptions = { widthPx: Math.round(60 * W["0"]) + 7, measure };
    for (const [id, r] of Object.entries(fx.builtin.read)) {
      const code = BUILTIN_FORMATS[Number(id)];
      total += 3;
      if (localFormat(code) !== r.local) misses.push({ where: `id ${id} local`, code, value: null, excel: r.local, ours: localFormat(code) });
      const a = formatValue(45000.5, code, wide), b = formatValue(-1234.5, code, wide);
      if (!same(r.text, a)) misses.push({ where: `id ${id}`, code, value: 45000.5, excel: r.text, ours: a.text });
      if (!same(r.neg, b)) misses.push({ where: `id ${id}`, code, value: -1234.5, excel: r.neg, ours: b.text });
    }
    // What Excel writes for each id when it saves.
    for (const [id, w] of Object.entries(fx.builtin.written)) {
      total++;
      const code = BUILTIN_FORMATS[Number(id)];
      const ours = builtinId(code);
      if (ours !== w.id) misses.push({ where: `id ${id} written`, code, value: null, excel: String(w.id), ours: String(ours) });
      if (w.code && w.code !== code && Number(id) < 50) misses.push({ where: `id ${id} written code`, code, value: null, excel: w.code, ours: code });
    }
    report("builtin", misses, total);
    expect(misses.length).toBeLessThanOrEqual(0);
  });
});

describe("German format codes (NumberFormatLocal)", () => {
  it("shows every recorded code in its German form", () => {
    const misses: Miss[] = [];
    let total = 0;
    for (const f of fx.formats) {
      if (f.invalid || KNOWN_LOCAL.has(f.code)) continue;
      total++;
      const ours = localFormat(f.code);
      if (ours !== f.local) misses.push({ where: "local", code: f.code, value: null, excel: f.local!, ours });
    }
    report("localFormat", misses, total);
    expect(misses.length).toBeLessThanOrEqual(0);
  });

  it("turns German codes into the file form Excel writes", () => {
    const misses: Miss[] = [];
    let total = 0;
    const recs = [...fx.local, ...Object.values(fx.categories).flat()];
    for (const r of recs) {
      if (r.invalid || !r.file || r.local === "G/Standard") continue;
      total++;
      const expected = r.file.code ?? BUILTIN_FORMATS[r.file.id];
      const ours = fileFormat(r.local);
      if (ours !== expected && !sameDisplay(ours, expected)) misses.push({ where: "file", code: r.local, value: null, excel: expected, ours });
    }
    report("fileFormat", misses, total);
    expect(misses.length).toBeLessThanOrEqual(0);
  });

  it("offers the categories and ribbon entries with codes Excel accepts", () => {
    const written = new Set(Object.values(fx.categories).flat().map((r) => r.file?.code ?? BUILTIN_FORMATS[r.file?.id ?? 0]));
    for (const c of FORMAT_CATEGORIES) {
      if (c.name === "Benutzerdefiniert" || c.name === "Standard") continue;
      for (const code of c.codes) expect(written.has(code), `${c.name}: ${code}`).toBe(true);
    }
    expect(RIBBON_FORMATS.map((r) => r.label)).toEqual(["Standard", "Zahl", "Währung", "Buchhaltung", "Datum, kurz", "Datum, lang", "Uhrzeit", "Prozent", "Bruch", "Wissenschaft", "Text"]);
    expect(formatValue(45000.5, RIBBON_FORMATS[5].code).text).toBe("Mittwoch, 15. März 2023");
    expect(formatValue(45000.5, RIBBON_FORMATS[6].code).text).toBe("12:00:00");
  });
});

/** Two codes that show the same text for a range of values. */
function sameDisplay(a: string, b: string): boolean {
  const vals: Value[] = [0, 1, -1, 0.5, 1234.5678, -0.25, 45000.5, 0.75, "Text"];
  return vals.every((v) => formatValue(v, a).text === formatValue(v, b).text);
}

describe("number formats: unit cases", () => {
  const f = (v: Value, code: string, opts?: FormatOptions) => formatValue(v, code, opts).text;
  it("errors and booleans", () => {
    expect(f(err("#N/A" as ErrorCode), "0.00")).toBe("#NV");
    expect(f(true, "General")).toBe("WAHR");
    expect(formatValue(false, "0").align).toBe("center");
    expect(f(null, "0.00")).toBe("");
  });
  it("colours and fills", () => {
    expect(formatValue(-1, "0;[Red]-0").color).toBe("#FF0000");
    expect(formatValue(1, "[Color10]0").color).toBe("#008000");
    const r = formatValue(5, "0*-");
    expect([r.text, r.fill, r.fillAt]).toEqual(["5", "-", 1]);
  });
});
