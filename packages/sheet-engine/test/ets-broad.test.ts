// PROGNOSE.ETS and its companions against German Excel on a broad measurement set (probe_excel_ets_broad.py,
// D-124): 48 series (seasons of length 2 to 24 with trend and noise, exact seasons, trends, random walks, level
// shifts, noise, short series, steps of 5, calendar months, missing points with both kinds of
// Datenvervollständigung, duplicate time stamps with every Aggregation, an unsorted timeline), each with the season
// found and with Saisonalität 0. Compared: the season length, PROGNOSE.ETS 1 to 24 steps ahead, PROGNOSE.ETS.KONFINT,
// PROGNOSE.ETS.STAT 1 to 8 and the values inside the timeline. Excel's parameter search is not documented, so the
// forecasts are compared with tolerances; the summary test prints the mean and worst errors (D-124 has the numbers).

import { describe, expect, it } from "vitest";
import { Calc, Sheet, Workbook, toFileFormula } from "../src";
import fixture from "./fixtures/excel-ets-broad.json";

type Series = { t: number[]; v: number[] };
type Rec = { season: number; stat: (number | string)[]; fit: (number | string)[]; ets: (number | string)[]; conf: (number | string)[] };
const series = fixture.series as unknown as Record<string, Series>;
const cases = fixture.cases as unknown as Record<string, Record<string, Rec>>;

function evalAll(s: Series, formulas: string[]): unknown[] {
  const wb = new Workbook();
  const sh = new Sheet("Tabelle1");
  wb.sheets.push(sh);
  s.t.forEach((t, i) => { sh.set(i, 0, { v: t }); sh.set(i, 1, { v: s.v[i] }); });
  formulas.forEach((f, i) => sh.set(i, 3, { v: null, f: toFileFormula(f) }));
  new Calc(wb).recalc();
  return formulas.map((_, i) => { const v = sh.get(i, 3)?.v; return v && typeof v === "object" ? String(v) : v; });
}

const args: Record<string, [number, number, number]> = { auto: [1, 1, 1], none: [0, 1, 1], zero: [1, 0, 1] };
for (let a = 2; a <= 7; a++) args[`agg${a}`] = [1, 1, a];

/** Relative error, measured against the series' scale where Excel's value is near zero. */
const rel = (a: number, b: number, scale: number) => Math.abs(a - b) / Math.max(Math.abs(b), 0.05 * scale, 1e-9);

export const summary = { ets: [] as number[], conf: [] as number[], fit: [] as number[], season: [0, 0], stat: {} as Record<number, number[]>, worst: { ets: "", conf: "" } };

describe("PROGNOSE.ETS on the broad measurement set (D-124)", () => {
  for (const [name, variants] of Object.entries(cases)) {
    it(name, () => {
      const s = series[name];
      const n = s.t.length;
      const vals = `$B$1:$B$${n}`, tl = `$A$1:$A$${n}`;
      const ts = [...new Set(s.t)].sort((a, b) => a - b);
      const months = name.startsWith("months");
      const step = ts[1] - ts[0];
      const last = ts[ts.length - 1];
      const at = (h: number) => (months ? `EDATUM(${last};${h})` : String(last + h * step));
      const scale = Math.max(...s.v.map(Math.abs));
      for (const [label, rec] of Object.entries(variants)) {
        const [se, fi, ag] = args[label];
        const tail = `;${se};${fi};${ag}`;
        const fs = [
          `=PROGNOSE.ETS.SAISONALITÄT(${vals};${tl};${fi};${ag})`,
          ...Array.from({ length: 8 }, (_, k) => `=PROGNOSE.ETS.STAT(${vals};${tl};${k + 1}${tail})`),
          ...ts.map((t) => `=PROGNOSE.ETS(${t};${vals};${tl}${tail})`),
          ...Array.from({ length: 24 }, (_, h) => `=PROGNOSE.ETS(${at(h + 1)};${vals};${tl}${tail})`),
          ...Array.from({ length: 24 }, (_, h) => `=PROGNOSE.ETS.KONFINT(${at(h + 1)};${vals};${tl};0,95${tail})`),
        ];
        const got = evalAll(s, fs) as number[];
        let k = 0;
        const season = got[k++];
        summary.season[1]++;
        if (season === rec.season) summary.season[0]++;
        const stat = got.slice(k, k + 8); k += 8;
        stat.forEach((v, i) => {
          const w = rec.stat[i];
          if (typeof w !== "number") return;
          (summary.stat[i + 1] ??= []).push(i < 3 ? Math.abs(v - w) : rel(v, w, Math.max(1e-9, Math.abs(w))));
        });
        const fit = got.slice(k, k + ts.length); k += ts.length;
        fit.forEach((v, i) => { const w = rec.fit[i]; if (typeof w === "number") summary.fit.push(rel(v, w, scale)); });
        const ets = got.slice(k, k + 24); k += 24;
        ets.forEach((v, i) => {
          const w = rec.ets[i];
          if (typeof w !== "number") return;
          const r = rel(v, w, scale);
          summary.ets.push(r);
          if (r >= Math.max(...summary.ets)) summary.worst.ets = `${name} ${label} h=${i + 1}: ${v} vs ${w}`;
        });
        if (process.env.RAABE_ETS_VERBOSE) {
          const es = ets.map((v, i) => (typeof rec.ets[i] === "number" ? rel(v, rec.ets[i] as number, scale) : 0));
          console.log(`${name} ${label}: season ${season}/${rec.season} alpha ${stat[0]}/${rec.stat[0]} gamma ${stat[2]}/${rec.stat[2]} ETS mean ${(es.reduce((a, b) => a + b, 0) / 24 * 100).toFixed(2)} % inside max ${(Math.max(...fit.map((v, i) => (typeof rec.fit[i] === "number" ? rel(v, rec.fit[i] as number, scale) : 0))) * 100).toFixed(1)} %`);
        }
        const conf = got.slice(k, k + 24);
        conf.forEach((v, i) => {
          const w = rec.conf[i];
          if (typeof w !== "number" || w === 0) return;
          const r = rel(v, w, scale);
          summary.conf.push(r);
          if (r >= Math.max(...summary.conf)) summary.worst.conf = `${name} ${label} h=${i + 1}: ${v} vs ${w}`;
        });
        expect(typeof ets[0], `${name} ${label}`).toBe("number");
      }
    });
  }
  it("summary (D-124)", () => {
    const mean = (xs: number[]) => xs.reduce((a, b) => a + b, 0) / Math.max(1, xs.length);
    const pct = (x: number) => `${(x * 100).toFixed(2)} %`;
    const line = (xs: number[]) => `mean ${pct(mean(xs))}, max ${pct(Math.max(...xs))}`;
    console.log(`season ${summary.season[0]}/${summary.season[1]}; ETS ${line(summary.ets)} (${summary.worst.ets}); KONFINT ${line(summary.conf)} (${summary.worst.conf}); inside ${line(summary.fit)}`);
    console.log(`STAT: ${Object.entries(summary.stat).map(([k, xs]) => `${k}: ${k <= "3" ? `mean |d| ${mean(xs).toFixed(4)}` : line(xs)}`).join("; ")}`);
    expect(summary.season[0] / summary.season[1]).toBeGreaterThan(0.9);
  });
});
