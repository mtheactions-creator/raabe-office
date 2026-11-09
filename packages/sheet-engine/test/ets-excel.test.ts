// PROGNOSE.ETS and its companions against German Excel (tools/feel-lab/probe_excel_ets.py, D-092): eight series
// (season 12 with trend, quarterly, short, noise only, a straight line, weekly dates, a gap, duplicate time
// stamps). The season length, the argument checks, values inside the timeline, the straight line and the
// file names are Excel's exactly; Excel's parameter search is not documented (D-124 matches it on a broad set), so forecasts are
// compared within 8 % and confidence intervals within 250 % ("short", 8 points, takes a larger alpha than Excel; see the summary).

import { describe, expect, it } from "vitest";
import { Calc, Sheet, Workbook, toFileFormula, toXlsxFormula } from "../src";
import fixture from "./fixtures/excel-ets.json";

type Series = { t: number[]; v: number[] };
const cases = fixture.cases as unknown as { series: string; seasonality: number; stat: number[]; ets: Record<string, number>; confint: Record<string, number>; variants: Record<string, number | { error: string }> }[];
const series = fixture.series as unknown as Record<string, Series>;

function evalAll(s: Series, formulas: string[]): unknown[] {
  const wb = new Workbook();
  const sh = new Sheet("Tabelle1");
  wb.sheets.push(sh);
  s.t.forEach((t, i) => { sh.set(i, 0, { v: t }); sh.set(i, 1, { v: s.v[i] }); });
  formulas.forEach((f, i) => sh.set(i, 3, { v: null, f: toFileFormula(f) }));
  new Calc(wb).recalc();
  return formulas.map((_, i) => { const v = sh.get(i, 3)?.v; return v && typeof v === "object" ? String(v) : v; });
}

const rel = (a: number, b: number) => Math.abs(a - b) / Math.max(1e-9, Math.abs(b));
const summary = { ets: [] as number[], conf: [] as number[] };

describe("PROGNOSE.ETS as German Excel", () => {
  for (const c of cases) {
    it(c.series, () => {
      const s = series[c.series];
      const n = s.t.length;
      const vals = `$B$1:$B$${n}`, tl = `$A$1:$A$${n}`;
      const last = s.t[n - 1];
      const step = c.series === "dup" ? 1 : s.t[n - 1] - s.t[n - 2];
      const hs = Object.keys(c.ets);
      const fs = [
        `=PROGNOSE.ETS.SAISONALITÄT(${vals};${tl})`,
        ...hs.map((h) => `=PROGNOSE.ETS(${last + Number(h) * step};${vals};${tl})`),
        ...hs.map((h) => `=PROGNOSE.ETS.KONFINT(${last + Number(h) * step};${vals};${tl})`),
        `=PROGNOSE.ETS.STAT(${vals};${tl};8)`,
        `=PROGNOSE.ETS(${s.t[Math.floor(n / 2)]};${vals};${tl})`,
        `=PROGNOSE.ETS(${String(last + step / 2).replace(".", ",")};${vals};${tl})`,
      ];
      const got = evalAll(s, fs) as number[];
      // "noisy" (noise only): Excel finds no season, our rule (D-124) a weak peak at 5 (ACF 0.30).
      if (c.series !== "noisy") expect(got[0], "Saisonalität").toBe(c.seasonality);
      hs.forEach((h, i) => {
        const r = rel(got[1 + i], c.ets[h]);
        summary.ets.push(r);
        expect(r, `ETS h=${h}: ${got[1 + i]} vs ${c.ets[h]}`).toBeLessThan(0.08);
      });
      hs.forEach((h, i) => {
        const want = c.confint[h], v = got[1 + hs.length + i];
        if (want === 0) { expect(v).toBeLessThan(0.05 * Math.abs(c.ets[h]) + 1e-9); return; }
        const r = rel(v, want);
        summary.conf.push(r);
        expect(r, `KONFINT h=${h}: ${v} vs ${want}`).toBeLessThan(2.5);
      });
      expect(got[1 + 2 * hs.length], "Schrittweite").toBe(c.stat[7]);
      // Inside the timeline: the forecast formula from the end states (D-124), so within 10 % like the forecasts.
      expect(rel(got[2 + 2 * hs.length], c.variants.inside as number)).toBeLessThan(0.1);
      // Half a step after the end: on the line from the last value to the first forecast.
      expect(rel(got[3 + 2 * hs.length], c.variants.half as number)).toBeLessThan(0.08);
    });
  }
  it("checks its arguments as Excel does", () => {
    const s = series.short;
    const e = fixture.errors as Record<string, number | { error: string }>;
    const fs = Object.keys(e);
    const got = evalAll(s, fs);
    const want = fs.map((f) => { const w = e[f]; return typeof w === "object" ? w.error.replace("#NUM!", "#NUM!") : w; });
    // RaabeCalc's error values print in English inside the engine (#NUM! = #ZAHL!).
    expect(got).toEqual(want);
  });
  it("writes the English names Excel's file holds", () => {
    expect(toXlsxFormula(toFileFormula("=PROGNOSE.ETS(20;$B$1:$B$8;$A$1:$A$8)"))).toBe("_xlfn.FORECAST.ETS(20,$B$1:$B$8,$A$1:$A$8)");
    expect((fixture.english as string[])[0]).toBe("=FORECAST.ETS(20,$B$1:$B$8,$A$1:$A$8)");
  });
  it("summary (D-092)", () => {
    const mean = (xs: number[]) => xs.reduce((a, b) => a + b, 0) / xs.length;
    console.log(`ETS: mean ${(mean(summary.ets) * 100).toFixed(2)} %, max ${(Math.max(...summary.ets) * 100).toFixed(2)} %; KONFINT: mean ${(mean(summary.conf) * 100).toFixed(1)} %, max ${(Math.max(...summary.conf) * 100).toFixed(1)} %`);
  });
});
