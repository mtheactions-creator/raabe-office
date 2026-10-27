// Statistical functions.

import { Ctx, FUNCS, def, deref } from "./eval";
import { defS, n, numbers, values, type Collect } from "./fn-util";
import { combin } from "./fn-math";
import { Arr, E, Ref, isErr, num, toNum, type Any, type V } from "./values";
import type { CellError } from "../model";

type Agg = (xs: number[]) => number | CellError;

function agg(name: string, f: Agg, opt: Collect = {}): void {
  def(name, 1, 255, (args, ctx) => {
    const xs = numbers(args, ctx, opt);
    if (isErr(xs)) return xs;
    const r = f(xs);
    return typeof r === "number" ? num(r) : r;
  });
}

const sum = (xs: number[]) => {
  let s = 0;
  for (const x of xs) s += x;
  return s;
};
const mean = (xs: number[]) => (xs.length ? sum(xs) / xs.length : E.div0);

/** Sum of squared deviations, computed around the mean in two passes for accuracy. */
function ssd(xs: number[]): number {
  const m = sum(xs) / xs.length;
  let s = 0;
  for (const x of xs) s += (x - m) * (x - m);
  return s;
}
const varS = (xs: number[]) => (xs.length < 2 ? E.div0 : ssd(xs) / (xs.length - 1));
const varP = (xs: number[]) => (xs.length < 1 ? E.div0 : ssd(xs) / xs.length);
const sqrtOr = (v: number | CellError) => (isErr(v) ? v : Math.sqrt(v));

agg("AVERAGE", mean);
agg("AVERAGEA", mean, { a: true, arrText0: true });
agg("MIN", (xs) => (xs.length ? Math.min(...xs) : 0));
agg("MAX", (xs) => (xs.length ? Math.max(...xs) : 0));
agg("MINA", (xs) => (xs.length ? Math.min(...xs) : 0), { a: true });
agg("MAXA", (xs) => (xs.length ? Math.max(...xs) : 0), { a: true });
agg("VAR.S", varS);
agg("VAR", varS);
agg("VAR.P", varP);
agg("VARP", varP);
agg("VARA", varS, { a: true });
agg("STDEV.S", (xs) => sqrtOr(varS(xs)));
agg("STDEV", (xs) => sqrtOr(varS(xs)));
agg("STDEVA", (xs) => sqrtOr(varS(xs)), { a: true });
agg("STDEV.P", (xs) => sqrtOr(varP(xs)));
agg("STDEVP", (xs) => sqrtOr(varP(xs)));
agg("MEDIAN", (xs) => {
  if (!xs.length) return E.num;
  const s = [...xs].sort((a, b) => a - b);
  const h = s.length >> 1;
  return s.length % 2 ? s[h] : (s[h - 1] + s[h]) / 2;
});
function mode(xs: number[]): number | CellError {
  const count = new Map<number, number>();
  let best = 0, bestV = 0;
  for (const x of xs) {
    const c = (count.get(x) ?? 0) + 1;
    count.set(x, c);
    if (c > best) { best = c; bestV = x; }
  }
  // Ties go to the value seen first.
  if (best < 2) return E.na;
  for (const x of xs) if (count.get(x) === best) return x;
  return bestV;
}
// MODUS.EINF skips typed-in text; an empty argument or no value at all gives #WERT!.
for (const name of ["MODE.SNGL", "MODE"]) {
  def(name, 1, 255, (args, ctx, nodes) => {
    if (nodes.some((nd) => nd.t === "miss")) return E.value;
    const xs = numbers(args, ctx, { skipText: true });
    if (isErr(xs)) return xs;
    if (!xs.length && args.every((a) => a instanceof Ref && !hasValue(a, ctx))) return E.value;
    return mode(xs);
  });
}

function hasValue(a: Ref, ctx: Ctx): boolean {
  let any = false;
  ctx.eachCell(a, () => { any = true; });
  return any;
}
agg("AVEDEV", (xs) => {
  if (!xs.length) return E.num;
  const m = sum(xs) / xs.length;
  return sum(xs.map((x) => Math.abs(x - m))) / xs.length;
});
agg("GEOMEAN", (xs) => {
  if (!xs.length || xs.some((x) => x <= 0)) return E.num;
  return Math.exp(sum(xs.map(Math.log)) / xs.length);
});
agg("HARMEAN", (xs) => {
  if (!xs.length) return E.na;
  if (xs.some((x) => x <= 0)) return E.num;
  return xs.length / sum(xs.map((x) => 1 / x));
});
agg("KURT", (xs) => {
  const k = xs.length;
  if (k < 4) return E.div0;
  const m = sum(xs) / k;
  const sd = Math.sqrt(ssd(xs) / (k - 1));
  if (sd === 0) return E.div0;
  const s4 = sum(xs.map((x) => ((x - m) / sd) ** 4));
  return ((k * (k + 1)) / ((k - 1) * (k - 2) * (k - 3))) * s4 - (3 * (k - 1) ** 2) / ((k - 2) * (k - 3));
});
agg("SKEW", (xs) => {
  const k = xs.length;
  if (k < 3) return E.div0;
  const m = sum(xs) / k;
  const sd = Math.sqrt(ssd(xs) / (k - 1));
  if (sd === 0) return E.div0;
  return (k / ((k - 1) * (k - 2))) * sum(xs.map((x) => ((x - m) / sd) ** 3));
});

def("COUNT", 1, 255, (args, ctx) => {
  let c = 0;
  for (const a of args) {
    if (a instanceof Ref) ctx.eachCell(a, (v) => { if (typeof v === "number") c++; });
    else if (a instanceof Arr) { for (const v of a.d) if (typeof v === "number") c++; }
    else if (a === null || typeof a === "number" || typeof a === "boolean") c++;
    else if (typeof a === "string" && !isErr(toNum(a))) c++;
  }
  return c;
});
def("COUNTA", 1, 255, (args, ctx) => {
  let c = 0;
  for (const a of args) {
    if (a instanceof Ref) ctx.eachCell(a, () => { c++; });
    else if (a instanceof Arr) { for (const v of a.d) if (v !== null) c++; }
    else c++;
  }
  return c;
});
def("COUNTBLANK", 1, 1, ([a], ctx) => {
  if (!(a instanceof Ref)) return E.value;
  let c = 0;
  for (let i = 0; i < a.areas.length; i++) for (const v of ctx.refArr(a, i).d) if (v === null || v === "") c++;
  const r = a.areas[0];
  const clipped = ctx.clip(a.sheet, r);
  // Rows beyond the used range are empty too.
  c += (r.r1 - r.r0 + 1) * (r.c1 - r.c0 + 1) - (clipped.r1 - clipped.r0 + 1) * (clipped.c1 - clipped.c0 + 1);
  return c;
});

/** The numbers of a range or array argument (text, logical values and empty cells left out). */
export function nums(a: Any, ctx: Ctx): number[] | CellError {
  return numbers([a instanceof Ref || a instanceof Arr ? a : a], ctx);
}

function kth(name: string, largest: boolean): void {
  defS(name, 2, 2, ([, kk], ctx, raw) => {
    const xs = nums(raw[0], ctx);
    if (isErr(xs)) return xs;
    const k = n(kk);
    if (isErr(k)) return k;
    const kc = Math.ceil(k);
    if (kc < 1 || kc > xs.length) return E.num;
    const s = [...xs].sort((a, b) => (largest ? b - a : a - b));
    return s[kc - 1];
  }, [0]);
}
kth("LARGE", true);
kth("SMALL", false);

function rank(name: string, avg: boolean): void {
  defS(name, 2, 3, ([x0, , ord], ctx, raw) => {
    if (!(raw[1] instanceof Ref)) return E.value;
    const xs = nums(raw[1], ctx);
    if (isErr(xs)) return xs;
    const x = n(x0);
    if (isErr(x)) return x;
    const asc = ord !== undefined && ord !== null && toNum(ord) !== 0;
    let before = 0, same = 0;
    for (const y of xs) {
      if (y === x) same++;
      else if (asc ? y < x : y > x) before++;
    }
    if (!same) return E.na;
    return avg ? before + (same + 1) / 2 : before + 1;
  }, [1]);
}
rank("RANK.EQ", false);
rank("RANK", false);
rank("RANK.AVG", true);

/** Pairs of numbers from two equally sized ranges; pairs with a non-number are skipped. */
function pairs(a: Any, b: Any, ctx: Ctx): [number[], number[]] | CellError {
  const x = values(a, ctx), y = values(b, ctx);
  if (x.length !== y.length) return E.na;
  const xs: number[] = [], ys: number[] = [];
  for (let i = 0; i < x.length; i++) {
    const p = x[i], q = y[i];
    if (isErr(p)) return p;
    if (isErr(q)) return q;
    if (typeof p === "number" && typeof q === "number") {
      xs.push(p);
      ys.push(q);
    }
  }
  return [xs, ys];
}

function lin(ys: number[], xs: number[]) {
  const k = xs.length;
  const mx = sum(xs) / k, my = sum(ys) / k;
  let sxy = 0, sxx = 0, syy = 0;
  for (let i = 0; i < k; i++) {
    sxy += (xs[i] - mx) * (ys[i] - my);
    sxx += (xs[i] - mx) ** 2;
    syy += (ys[i] - my) ** 2;
  }
  return { k, mx, my, sxy, sxx, syy };
}

function twoArr(name: string, f: (ys: number[], xs: number[]) => number | CellError, swap = false): void {
  def(name, 2, 2, ([a, b], ctx) => {
    const p = pairs(a, b, ctx);
    if (isErr(p)) return p;
    const r = swap ? f(p[1], p[0]) : f(p[0], p[1]);
    return typeof r === "number" ? num(r) : r;
  });
}

twoArr("CORREL", (xs, ys) => {
  const l = lin(ys, xs);
  if (l.k === 0 || l.sxx === 0 || l.syy === 0) return E.div0;
  return l.sxy / Math.sqrt(l.sxx * l.syy);
});
twoArr("SLOPE", (ys, xs) => {
  const l = lin(ys, xs);
  if (l.k === 0 || l.sxx === 0) return E.div0;
  return l.sxy / l.sxx;
});
twoArr("INTERCEPT", (ys, xs) => {
  const l = lin(ys, xs);
  if (l.k === 0 || l.sxx === 0) return E.div0;
  return l.my - (l.sxy / l.sxx) * l.mx;
});
twoArr("RSQ", (ys, xs) => {
  const l = lin(ys, xs);
  if (l.k === 0 || l.sxx === 0 || l.syy === 0) return E.div0;
  return (l.sxy * l.sxy) / (l.sxx * l.syy);
});
twoArr("STEYX", (ys, xs) => {
  const l = lin(ys, xs);
  if (l.k < 3 || l.sxx === 0) return E.div0;
  return Math.sqrt((l.syy - (l.sxy * l.sxy) / l.sxx) / (l.k - 2));
});
twoArr("COVARIANCE.P", (xs, ys) => {
  const l = lin(ys, xs);
  return l.k === 0 ? E.div0 : l.sxy / l.k;
});
twoArr("COVARIANCE.S", (xs, ys) => {
  const l = lin(ys, xs);
  return l.k < 2 ? E.div0 : l.sxy / (l.k - 1);
});

function forecast(name: string): void {
  defS(name, 3, 3, ([x0], ctx, raw) => {
    const x = n(x0);
    if (isErr(x)) return x;
    const p = pairs(raw[1], raw[2], ctx);
    if (isErr(p)) return p;
    const l = lin(p[0], p[1]);
    if (l.k === 0 || l.sxx === 0) return E.div0;
    const b = l.sxy / l.sxx;
    return num(l.my - b * l.mx + b * x);
  }, [1, 2]);
}
forecast("FORECAST");
forecast("FORECAST.LINEAR");

function percentile(name: string, mode: "inc" | "exc", quart: boolean): void {
  defS(name, 2, 2, ([, k0], ctx, raw) => {
    const xs = nums(raw[0], ctx);
    if (isErr(xs)) return xs;
    let k = n(k0);
    if (isErr(k)) return k;
    if (quart) {
      k = Math.trunc(k);
      if (mode === "inc" ? k < 0 || k > 4 : k <= 0 || k >= 4) return E.num;
      k /= 4;
    }
    const s = [...xs].sort((a, b) => a - b);
    const m = s.length;
    if (!m) return E.num;
    let h: number;
    if (mode === "inc") {
      if (k < 0 || k > 1) return E.num;
      h = (m - 1) * k;
    } else {
      if (k <= 0 || k >= 1) return E.num;
      h = (m + 1) * k - 1;
      if (h < 0 || h > m - 1) return E.num;
    }
    const lo = Math.floor(h);
    const f = h - lo;
    return lo + 1 < m ? s[lo] + f * (s[lo + 1] - s[lo]) : s[lo];
  }, [0]);
}
percentile("PERCENTILE.INC", "inc", false);
percentile("PERCENTILE", "inc", false);
percentile("PERCENTILE.EXC", "exc", false);
percentile("QUARTILE.INC", "inc", true);
percentile("QUARTILE", "inc", true);
percentile("QUARTILE.EXC", "exc", true);

defS("PERCENTRANK.INC", 2, 3, ([, x0, sig0], ctx, raw) => {
  const xs = nums(raw[0], ctx);
  if (isErr(xs)) return xs;
  const x = n(x0);
  if (isErr(x)) return x;
  const sig = sig0 === undefined || sig0 === null ? 3 : n(sig0);
  if (isErr(sig)) return sig;
  if (sig < 1) return E.num;
  const s = [...xs].sort((a, b) => a - b);
  if (!s.length || x < s[0] || x > s[s.length - 1]) return E.na;
  let r: number;
  const i = s.findIndex((v) => v >= x);
  if (s[i] === x) {
    r = s.indexOf(x) / (s.length - 1);
  } else {
    const lo = i - 1;
    r = (lo + (x - s[lo]) / (s[i] - s[lo])) / (s.length - 1);
  }
  const f = 10 ** Math.trunc(sig);
  return Math.floor(r * f + 1e-9) / f;
}, [0]);

def("FREQUENCY", 2, 2, ([a, b], ctx) => {
  const xs = nums(a, ctx), bins = nums(b, ctx);
  if (isErr(xs)) return xs;
  if (isErr(bins)) return bins;
  const order = bins.map((v, i) => [v, i] as const).sort((p, q) => p[0] - q[0]);
  const out = new Array(bins.length + 1).fill(0);
  for (const x of xs) {
    const k = order.find(([v]) => x <= v);
    out[k ? k[1] : bins.length]++;
  }
  return new Arr(out.length, 1, out);
});

// ----- Distributions -----------------------------------------------------------------------

/** erf with a relative error below 1e-15 (series for small x, continued fraction for large x). */
export function erf(x: number): number {
  if (x < 0) return -erf(-x);
  if (x < 2.5) {
    // Maclaurin series.
    let sumv = x, term = x;
    for (let k = 1; k < 200; k++) {
      term *= (-x * x) / k;
      const add = term / (2 * k + 1);
      sumv += add;
      if (Math.abs(add) < 1e-17 * Math.abs(sumv)) break;
    }
    return (2 / Math.sqrt(Math.PI)) * sumv;
  }
  return 1 - erfc(x);
}

/** Complementary error function (continued fraction for x ≥ 2.5). */
export function erfc(x: number): number {
  if (x < 2.5) return 1 - erf(x);
  // Lentz's algorithm for the continued fraction of erfc.
  let f = x, c = x, d = 0;
  const tiny = 1e-300;
  for (let k = 1; k < 500; k++) {
    const an = k / 2;
    d = x + an * d;
    d = Math.abs(d) < tiny ? tiny : d;
    c = x + an / c;
    c = Math.abs(c) < tiny ? tiny : c;
    d = 1 / d;
    const delta = c * d;
    f *= delta;
    if (Math.abs(delta - 1) < 1e-16) break;
  }
  return Math.exp(-x * x) / (f * Math.sqrt(Math.PI));
}

export function normCdf(z: number): number {
  return z < 0 ? 0.5 * erfc(-z / Math.SQRT2) : 1 - 0.5 * erfc(z / Math.SQRT2);
}

export function normPdf(z: number): number {
  return Math.exp(-0.5 * z * z) / Math.sqrt(2 * Math.PI);
}

/** Inverse of the standard normal distribution (Acklam's start, refined by Newton steps). */
export function normInv(p: number): number {
  if (p > 0.5) return -normInv(1 - p);
  const a = [-3.969683028665376e1, 2.209460984245205e2, -2.759285104469687e2, 1.38357751867269e2, -3.066479806614716e1, 2.506628277459239];
  const b = [-5.447609879822406e1, 1.615858368580409e2, -1.556989798598866e2, 6.680131188771972e1, -1.328068155288572e1];
  const c = [-7.784894002430293e-3, -3.223964580411365e-1, -2.400758277161838, -2.549732539343734, 4.374664141464968, 2.938163982698783];
  const d = [7.784695709041462e-3, 3.224671290700398e-1, 2.445134137142996, 3.754408661907416];
  const pl = 0.02425;
  let x: number;
  if (p < pl) {
    const q = Math.sqrt(-2 * Math.log(p));
    x = (((((c[0] * q + c[1]) * q + c[2]) * q + c[3]) * q + c[4]) * q + c[5]) / ((((d[0] * q + d[1]) * q + d[2]) * q + d[3]) * q + 1);
  } else if (p <= 1 - pl) {
    const q = p - 0.5, r = q * q;
    x = ((((((a[0] * r + a[1]) * r + a[2]) * r + a[3]) * r + a[4]) * r + a[5]) * q) / (((((b[0] * r + b[1]) * r + b[2]) * r + b[3]) * r + b[4]) * r + 1);
  } else {
    const q = Math.sqrt(-2 * Math.log(1 - p));
    x = -(((((c[0] * q + c[1]) * q + c[2]) * q + c[3]) * q + c[4]) * q + c[5]) / ((((d[0] * q + d[1]) * q + d[2]) * q + d[3]) * q + 1);
  }
  for (let i = 0; i < 2; i++) {
    const e = normCdf(x) - p;
    const u = e * Math.sqrt(2 * Math.PI) * Math.exp((x * x) / 2);
    x = x - u / (1 + (x * u) / 2);
  }
  return x;
}

defS("NORM.S.DIST", 2, 2, ([z0, c0]) => {
  const z = n(z0);
  if (isErr(z)) return z;
  const c = n(c0);
  if (isErr(c)) return c;
  return c ? normCdf(z) : normPdf(z);
});
defS("NORM.DIST", 4, 4, ([x0, m0, s0, c0]) => {
  const x = n(x0), m = n(m0), sd = n(s0), c = n(c0);
  for (const v of [x, m, sd, c]) if (isErr(v)) return v;
  if ((sd as number) <= 0) return E.num;
  const z = ((x as number) - (m as number)) / (sd as number);
  return c ? normCdf(z) : normPdf(z) / (sd as number);
});
defS("NORM.S.INV", 1, 1, ([p0]) => {
  const p = n(p0);
  if (isErr(p)) return p;
  if (p <= 0 || p >= 1) return E.num;
  return normInv(p);
});
defS("NORM.INV", 3, 3, ([p0, m0, s0]) => {
  const p = n(p0), m = n(m0), sd = n(s0);
  for (const v of [p, m, sd]) if (isErr(v)) return v;
  if ((p as number) <= 0 || (p as number) >= 1 || (sd as number) <= 0) return E.num;
  return (m as number) + (sd as number) * normInv(p as number);
});
defS("STANDARDIZE", 3, 3, ([x0, m0, s0]) => {
  const x = n(x0), m = n(m0), sd = n(s0);
  for (const v of [x, m, sd]) if (isErr(v)) return v;
  if ((sd as number) <= 0) return E.num;
  return ((x as number) - (m as number)) / (sd as number);
});

function binomPmf(k: number, nn: number, p: number): number {
  if (p === 0) return k === 0 ? 1 : 0;
  if (p === 1) return k === nn ? 1 : 0;
  // In logarithms, so large n does not overflow.
  return Math.exp(lnCombin(nn, k) + k * Math.log(p) + (nn - k) * Math.log(1 - p));
}

function lnCombin(nn: number, k: number): number {
  if (nn < 1000) return Math.log(combin(nn, k));
  return lnGamma(nn + 1) - lnGamma(k + 1) - lnGamma(nn - k + 1);
}

export function lnGamma(x: number): number {
  // Lanczos approximation (g = 7, n = 9).
  const g = 7;
  const cf = [0.99999999999980993, 676.5203681218851, -1259.1392167224028, 771.32342877765313, -176.61502916214059, 12.507343278686905,
    -0.13857109526572012, 9.9843695780195716e-6, 1.5056327351493116e-7];
  if (x < 0.5) return Math.log(Math.PI / Math.abs(Math.sin(Math.PI * x))) - lnGamma(1 - x);
  x -= 1;
  let a = cf[0];
  const t = x + g + 0.5;
  for (let i = 1; i < 9; i++) a += cf[i] / (x + i);
  return 0.5 * Math.log(2 * Math.PI) + (x + 0.5) * Math.log(t) - t + Math.log(a);
}

defS("BINOM.DIST", 4, 4, ([k0, n0, p0, c0]) => {
  const k = n(k0), nn = n(n0), p = n(p0), c = n(c0);
  for (const v of [k, nn, p, c]) if (isErr(v)) return v;
  const kk = Math.trunc(k as number), N = Math.trunc(nn as number), pp = p as number;
  if (kk < 0 || kk > N || pp < 0 || pp > 1) return E.num;
  if (!c) return binomPmf(kk, N, pp);
  let s = 0;
  for (let i = 0; i <= kk; i++) s += binomPmf(i, N, pp);
  return Math.min(1, s);
});
defS("BINOM.INV", 3, 3, ([n0, p0, a0]) => {
  const nn = n(n0), p = n(p0), a = n(a0);
  for (const v of [nn, p, a]) if (isErr(v)) return v;
  const N = Math.trunc(nn as number);
  if (N < 0 || (p as number) < 0 || (p as number) > 1 || (a as number) < 0 || (a as number) > 1) return E.num;
  let s = 0;
  for (let i = 0; i <= N; i++) {
    s += binomPmf(i, N, p as number);
    if (s >= (a as number) - 1e-15) return i;
  }
  return N;
});
defS("POISSON.DIST", 3, 3, ([x0, m0, c0]) => {
  const x = n(x0), m = n(m0), c = n(c0);
  for (const v of [x, m, c]) if (isErr(v)) return v;
  const k = Math.trunc(x as number), mu = m as number;
  if (k < 0 || mu < 0) return E.num;
  const pmf = (i: number) => (mu === 0 ? (i === 0 ? 1 : 0) : Math.exp(i * Math.log(mu) - mu - lnGamma(i + 1)));
  if (!c) return pmf(k);
  let s = 0;
  for (let i = 0; i <= k; i++) s += pmf(i);
  return Math.min(1, s);
});
defS("EXPON.DIST", 3, 3, ([x0, l0, c0]) => {
  const x = n(x0), l = n(l0), c = n(c0);
  for (const v of [x, l, c]) if (isErr(v)) return v;
  if ((x as number) < 0 || (l as number) <= 0) return E.num;
  return c ? 1 - Math.exp(-(l as number) * (x as number)) : (l as number) * Math.exp(-(l as number) * (x as number));
});

// Student's t distribution through the regularised incomplete beta function.
function betacf(a: number, b: number, x: number): number {
  const tiny = 1e-300;
  let c = 1, d = 1 - ((a + b) * x) / (a + 1);
  d = Math.abs(d) < tiny ? tiny : d;
  d = 1 / d;
  let h = d;
  for (let m = 1; m <= 1000; m++) {
    const m2 = 2 * m;
    let aa = (m * (b - m) * x) / ((a + m2 - 1) * (a + m2));
    d = 1 + aa * d; d = Math.abs(d) < tiny ? tiny : d;
    c = 1 + aa / c; c = Math.abs(c) < tiny ? tiny : c;
    d = 1 / d;
    h *= d * c;
    aa = (-(a + m) * (a + b + m) * x) / ((a + m2) * (a + m2 + 1));
    d = 1 + aa * d; d = Math.abs(d) < tiny ? tiny : d;
    c = 1 + aa / c; c = Math.abs(c) < tiny ? tiny : c;
    d = 1 / d;
    const del = d * c;
    h *= del;
    if (Math.abs(del - 1) < 1e-16) break;
  }
  return h;
}

export function ibeta(x: number, a: number, b: number): number {
  if (x <= 0) return 0;
  if (x >= 1) return 1;
  const bt = Math.exp(lnGamma(a + b) - lnGamma(a) - lnGamma(b) + a * Math.log(x) + b * Math.log(1 - x));
  return x < (a + 1) / (a + b + 2) ? (bt * betacf(a, b, x)) / a : 1 - (bt * betacf(b, a, 1 - x)) / b;
}

/** P(T ≤ t) for Student's t with df degrees of freedom. */
export function tCdf(t: number, df: number): number {
  if (t * t < df) {
    // Near 0 through the other tail of the beta function, which keeps the small difference.
    const i = ibeta((t * t) / (df + t * t), 0.5, df / 2);
    return 0.5 + Math.sign(t) * 0.5 * i;
  }
  const x = df / (df + t * t);
  const tail = 0.5 * ibeta(x, df / 2, 0.5);
  return t >= 0 ? 1 - tail : tail;
}

function tPdf(t: number, df: number): number {
  return Math.exp(lnGamma((df + 1) / 2) - lnGamma(df / 2) - 0.5 * Math.log(df * Math.PI) - ((df + 1) / 2) * Math.log(1 + (t * t) / df));
}

function tInv(p: number, df: number): number {
  if (p === 0.5) return 0;
  // Bisection on the CDF, then Newton polishing.
  let lo = -1e3, hi = 1e3;
  for (let i = 0; i < 200; i++) {
    const mid = (lo + hi) / 2;
    if (tCdf(mid, df) < p) lo = mid;
    else hi = mid;
  }
  let x = (lo + hi) / 2;
  for (let i = 0; i < 3; i++) x -= (tCdf(x, df) - p) / tPdf(x, df);
  return x;
}

defS("T.DIST", 3, 3, ([x0, d0, c0]) => {
  const x = n(x0), d = n(d0), c = n(c0);
  for (const v of [x, d, c]) if (isErr(v)) return v;
  const df = Math.trunc(d as number);
  if (df < 1) return E.num;
  return c ? tCdf(x as number, df) : tPdf(x as number, df);
});
defS("T.INV", 2, 2, ([p0, d0]) => {
  const p = n(p0), d = n(d0);
  for (const v of [p, d]) if (isErr(v)) return v;
  const df = Math.trunc(d as number);
  if ((p as number) <= 0 || (p as number) >= 1 || df < 1) return E.num;
  return tInv(p as number, df);
});
defS("T.INV.2T", 2, 2, ([p0, d0]) => {
  const p = n(p0), d = n(d0);
  for (const v of [p, d]) if (isErr(v)) return v;
  const df = Math.trunc(d as number);
  if ((p as number) <= 0 || (p as number) > 1 || df < 1) return E.num;
  return Math.abs(tInv((p as number) / 2, df));
});

def("T.TEST", 4, 4, ([a, b, t0, k0], ctx) => {
  const tails = toNum(deref(t0, ctx) as V), type = toNum(deref(k0, ctx) as V);
  if (isErr(tails)) return tails;
  if (isErr(type)) return type;
  const T = Math.trunc(tails), K = Math.trunc(type);
  if ((T !== 1 && T !== 2) || K < 1 || K > 3) return E.num;
  let t: number, df: number;
  if (K === 1) {
    const p = pairs(a, b, ctx);
    if (isErr(p)) return p;
    const va = values(a, ctx), vb = values(b, ctx);
    if (va.length !== vb.length) return E.na;
    const d = p[0].map((x, i) => x - p[1][i]);
    if (d.length < 2) return E.div0;
    const m = sum(d) / d.length;
    const sd = Math.sqrt(ssd(d) / (d.length - 1));
    if (sd === 0) return E.div0;
    t = m / (sd / Math.sqrt(d.length));
    df = d.length - 1;
  } else {
    const xa = nums(a, ctx), xb = nums(b, ctx);
    if (isErr(xa)) return xa;
    if (isErr(xb)) return xb;
    const n1 = xa.length, n2 = xb.length;
    if (n1 < 2 || n2 < 2) return E.div0;
    const m1 = sum(xa) / n1, m2 = sum(xb) / n2;
    const v1 = ssd(xa) / (n1 - 1), v2 = ssd(xb) / (n2 - 1);
    if (K === 2) {
      df = n1 + n2 - 2;
      const sp = ((n1 - 1) * v1 + (n2 - 1) * v2) / df;
      t = (m1 - m2) / Math.sqrt(sp * (1 / n1 + 1 / n2));
    } else {
      const s1 = v1 / n1, s2 = v2 / n2;
      t = (m1 - m2) / Math.sqrt(s1 + s2);
      df = (s1 + s2) ** 2 / (s1 ** 2 / (n1 - 1) + s2 ** 2 / (n2 - 1));
    }
  }
  const x = df / (df + t * t);
  const one = 0.5 * ibeta(x, df / 2, 0.5);
  return num(T === 1 ? one : 2 * one);
});

// ----- TEILERGEBNIS ------------------------------------------------------------------------

const SUBTOTALS: Record<number, string> = { 1: "AVERAGE", 2: "COUNT", 3: "COUNTA", 4: "MAX", 5: "MIN", 6: "PRODUCT", 7: "STDEV", 8: "STDEVP", 9: "SUM", 10: "VAR", 11: "VARP" };

def("SUBTOTAL", 2, 255, (args, ctx) => {
  const f0 = toNum(deref(args[0], ctx) as V);
  if (isErr(f0)) return f0;
  const code = Math.trunc(f0);
  const name = SUBTOTALS[code > 100 ? code - 100 : code];
  if (!name) return E.value;
  const refs: Any[] = [];
  for (const a of args.slice(1)) {
    if (!(a instanceof Ref)) return E.value;
    // Other TEILERGEBNIS results in the range are left out, and with 101-111 hidden rows too.
    const vals: V[] = [];
    const sheet = ctx.host.wb.sheets[a.sheet];
    for (let i = 0; i < a.areas.length; i++) {
      const ar = ctx.clip(a.sheet, a.areas[i]);
      const arr = ctx.refArr(a, i);
      for (let r = 0; r < arr.rows; r++) {
        if (code > 100 && sheet?.rows.get(ar.r0 + r)?.hidden) continue;
        for (let c = 0; c < arr.cols; c++) {
          const f = ctx.host.formula(a.sheet, ar.r0 + r, ar.c0 + c);
          if (f && /SUBTOTAL\(|AGGREGATE\(/i.test(f)) continue;
          vals.push(arr.get(r, c));
        }
      }
    }
    refs.push(new Arr(vals.length, 1, vals));
  }
  return FUNCS.get(name)!.fn(refs.length ? refs : [new Arr(0, 0, [])], ctx, []);
});
