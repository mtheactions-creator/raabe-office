// Bedingte Formatierung as Excel applies it when drawing: rules in priority order, the first
// matching format wins per property ("Anhalten" stops further rules); data bars and colour
// scales are scaled to the minimum and maximum of their range.

import { cellKey, evaluate, type CondFormat, type CondRule, type Rect, type Sheet, type Value, type Workbook } from "@raabe/sheet-engine";

export interface CondLook {
  fill?: string;
  color?: string;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strike?: boolean;
  /** Datenbalken: the bar's length as a share of the cell and its colour. */
  bar?: { frac: number; color: string };
}

interface Stats { nums: number[]; min: number; max: number; avg: number; counts: Map<string, number> }

const inRef = (refs: Rect[], row: number, col: number) => refs.some((r) => row >= r.r0 && row <= r.r1 && col >= r.c0 && col <= r.c1);

/** Evaluates the conditional formats of one sheet for drawing; build it again when cells change. */
export class CondEval {
  private stats = new Map<CondFormat, Stats>();
  private cache = new Map<number, CondLook | null>();

  constructor(private wb: Workbook, private sheet: Sheet, private sheetIndex: number) {}

  private statsOf(f: CondFormat): Stats {
    let s = this.stats.get(f);
    if (s) return s;
    const nums: number[] = [];
    const counts = new Map<string, number>();
    for (const r of f.ref) {
      for (const [k, c] of this.sheet.cells) {
        const row = Math.floor(k / 16384), col = k % 16384;
        if (row < r.r0 || row > r.r1 || col < r.c0 || col > r.c1 || c.v === null) continue;
        if (typeof c.v === "number") nums.push(c.v);
        const key = String(c.v).toLowerCase();
        counts.set(key, (counts.get(key) ?? 0) + 1);
      }
    }
    const min = nums.length ? Math.min(...nums) : 0, max = nums.length ? Math.max(...nums) : 0;
    const avg = nums.length ? nums.reduce((a, b) => a + b, 0) / nums.length : 0;
    s = { nums: nums.sort((a, b) => a - b), min, max, avg, counts };
    this.stats.set(f, s);
    return s;
  }

  private value(rule: CondRule, i: number, row: number, col: number): Value {
    const f = rule.formulas?.[i];
    if (f === undefined) return null;
    if (/^-?\d+(\.\d+)?(E[+-]?\d+)?$/i.test(f)) return Number(f);
    if (/^"(.*)"$/.test(f)) return f.slice(1, -1).replace(/""/g, '"');
    try {
      const v = evaluate(this.wb, this.sheetIndex, row, col, f);
      return Array.isArray(v) ? (v[0]?.[0] ?? null) : v;
    } catch {
      return null;
    }
  }

  look(row: number, col: number): CondLook | null {
    const list = this.sheet.conditional;
    if (!list?.length) return null;
    const key = cellKey(row, col);
    if (this.cache.has(key)) return this.cache.get(key)!;
    const v = this.sheet.get(row, col)?.v ?? null;
    const rules: { f: CondFormat; r: CondRule }[] = [];
    for (const f of list) if (inRef(f.ref, row, col)) for (const r of f.rules) rules.push({ f, r });
    rules.sort((a, b) => a.r.priority - b.r.priority);
    let out: CondLook | null = null;
    for (const { f, r } of rules) {
      const hit = this.matches(f, r, v, row, col);
      if (!hit) continue;
      out ??= {};
      if (r.type === "dataBar" && typeof v === "number") {
        const s = this.statsOf(f);
        const span = s.max - Math.min(0, s.min);
        out.bar ??= { frac: span ? (v - Math.min(0, s.min)) / span : 0, color: r.colors?.[0] ?? "#638EC6" };
      } else if (r.type === "colorScale" && typeof v === "number") {
        out.fill ??= scaleColor(this.statsOf(f), v, r.colors ?? ["#F8696B", "#FFEB84", "#63BE7B"]);
      } else if (r.format) {
        out.fill ??= r.format.fill;
        out.color ??= r.format.color;
        out.bold ??= r.format.bold;
        out.italic ??= r.format.italic;
        out.underline ??= r.format.underline;
        out.strike ??= r.format.strike;
      }
      if (r.stopIfTrue) break;
    }
    this.cache.set(key, out);
    return out;
  }

  private matches(f: CondFormat, r: CondRule, v: Value, row: number, col: number): boolean {
    const num = typeof v === "number" ? v : null;
    const text = v === null ? "" : String(v).toLowerCase();
    switch (r.type) {
      case "cellIs": {
        if (v === null) return false;
        const a = this.value(r, 0, row, col), b = this.value(r, 1, row, col);
        const cmp = (x: Value, y: Value) => (typeof x === "number" && typeof y === "number" ? x - y : String(x).toLowerCase().localeCompare(String(y).toLowerCase(), "de"));
        switch (r.operator) {
          case "greaterThan": return cmp(v, a) > 0;
          case "lessThan": return cmp(v, a) < 0;
          case "greaterThanOrEqual": return cmp(v, a) >= 0;
          case "lessThanOrEqual": return cmp(v, a) <= 0;
          case "equal": return cmp(v, a) === 0;
          case "notEqual": return cmp(v, a) !== 0;
          case "between": return cmp(v, a) >= 0 && cmp(v, b) <= 0;
          case "notBetween": return cmp(v, a) < 0 || cmp(v, b) > 0;
        }
        return false;
      }
      case "containsText": return !!r.text && text.includes(r.text.toLowerCase());
      case "notContainsText": return !!r.text && !text.includes(r.text.toLowerCase());
      case "beginsWith": return !!r.text && text.startsWith(r.text.toLowerCase());
      case "endsWith": return !!r.text && text.endsWith(r.text.toLowerCase());
      case "containsBlanks": return v === null || String(v).trim() === "";
      case "duplicateValues": return v !== null && (this.statsOf(f).counts.get(text) ?? 0) > 1;
      case "uniqueValues": return v !== null && (this.statsOf(f).counts.get(text) ?? 0) === 1;
      case "top10": {
        if (num === null) return false;
        const s = this.statsOf(f);
        const n = r.percent ? Math.max(1, Math.floor((s.nums.length * (r.rank ?? 10)) / 100)) : (r.rank ?? 10);
        const sorted = r.bottom ? s.nums : [...s.nums].reverse();
        const limit = sorted[Math.min(n, sorted.length) - 1];
        return r.bottom ? num <= limit : num >= limit;
      }
      case "aboveAverage": {
        if (num === null) return false;
        const s = this.statsOf(f);
        return r.above === false ? num < s.avg : num > s.avg;
      }
      case "dataBar":
      case "colorScale": return num !== null;
      case "expression": {
        const x = this.value(r, 0, row, col);
        return x === true || (typeof x === "number" && x !== 0);
      }
    }
    return false;
  }
}

function scaleColor(s: Stats, v: number, colors: string[]): string {
  const t = s.max === s.min ? 0.5 : (v - s.min) / (s.max - s.min);
  if (colors.length === 2) return mix(colors[0], colors[1], t);
  return t < 0.5 ? mix(colors[0], colors[1], t * 2) : mix(colors[1], colors[2], (t - 0.5) * 2);
}

function mix(a: string, b: string, t: number): string {
  const p = (h: string, i: number) => parseInt(h.slice(1 + i * 2, 3 + i * 2), 16);
  const c = [0, 1, 2].map((i) => Math.round(p(a, i) + (p(b, i) - p(a, i)) * t));
  return "#" + c.map((x) => x.toString(16).padStart(2, "0")).join("");
}
