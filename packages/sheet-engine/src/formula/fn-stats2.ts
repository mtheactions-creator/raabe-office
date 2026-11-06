// Statistical functions, second round: the distributions (beta, chi-squared, F, gamma,
// lognormal, Weibull, hypergeometric, negative binomial …) with their inverses and tests,
// regression (RGP, RKP, TREND, VARIATION), and the compatibility names Excel keeps
// (BETAVERT, CHIVERT, NORMVERT, TVERT …).

import { Ctx, FUNCS, def, deref } from "./eval";
import { defS, n, numbers, values, toArr } from "./fn-util";
import { combin } from "./fn-math";
import { ibeta, lnGamma, normCdf, normInv, normPdf, nums } from "./fn-stats";
import { Arr, E, isErr, num, type Any, type V } from "./values";
import type { CellError } from "../model";

type R = number | CellError;

/** Another name for an existing function (the compatibility versions). */
export function alias(name: string, of: string): void {
  const d = FUNCS.get(of);
  if (d) FUNCS.set(name, d);
}

/** A function of numbers only; missing optional arguments take their default. */
export function numFn(name: string, min: number, max: number, f: (x: number[]) => R, defaults: number[] = []): void {
  defS(name, min, max, (vals) => {
    const xs: number[] = [];
    for (let i = 0; i < max; i++) {
      const v = vals[i];
      if (v === undefined) {
        xs.push(defaults[i] ?? 0);
        continue;
      }
      const x = n(v);
      if (isErr(x)) return x;
      xs.push(x);
    }
    const r = f(xs);
    return typeof r === "number" ? num(r) : r;
  });
}

const sum = (xs: number[]) => xs.reduce((s, x) => s + x, 0);

function ssd(xs: number[]): number {
  const m = sum(xs) / xs.length;
  let s = 0;
  for (const x of xs) s += (x - m) * (x - m);
  return s;
}

// ----- Special functions -------------------------------------------------------------------

/** Regularised lower incomplete gamma function P(a, x). */
export function gammaP(a: number, x: number): number {
  if (x <= 0) return 0;
  if (x < a + 1) return gser(a, x);
  return 1 - gcf(a, x);
}

/** Regularised upper incomplete gamma function Q(a, x) = 1 − P(a, x). */
export function gammaQ(a: number, x: number): number {
  if (x <= 0) return 1;
  if (x < a + 1) return 1 - gser(a, x);
  return gcf(a, x);
}

function gser(a: number, x: number): number {
  let ap = a, del = 1 / a, s = del;
  for (let i = 0; i < 10000; i++) {
    ap++;
    del *= x / ap;
    s += del;
    if (Math.abs(del) < Math.abs(s) * 1e-17) break;
  }
  return s * Math.exp(-x + a * Math.log(x) - lnGamma(a));
}

function gcf(a: number, x: number): number {
  const tiny = 1e-300;
  let b = x + 1 - a, c = 1 / tiny, d = 1 / b, h = d;
  for (let i = 1; i < 10000; i++) {
    const an = -i * (i - a);
    b += 2;
    d = an * d + b;
    if (Math.abs(d) < tiny) d = tiny;
    c = b + an / c;
    if (Math.abs(c) < tiny) c = tiny;
    d = 1 / d;
    const del = d * c;
    h *= del;
    if (Math.abs(del - 1) < 1e-17) break;
  }
  return Math.exp(-x + a * Math.log(x) - lnGamma(a)) * h;
}

/** Γ(x): exact for whole numbers, Lanczos otherwise, with the reflection for x < 0.5. */
export function gamma(x: number): number {
  if (Number.isInteger(x) && x > 0 && x <= 171) {
    let f = 1;
    for (let i = 2; i < x; i++) f *= i;
    return f;
  }
  if (x < 0.5) return Math.PI / (Math.sin(Math.PI * x) * gamma(1 - x));
  return Math.exp(lnGamma(x));
}

const lnBeta = (a: number, b: number) => lnGamma(a) + lnGamma(b) - lnGamma(a + b);

/**
 * The x with cdf(x) = p for an increasing cdf, by bisection down to adjacent doubles (the
 * interval [lo, hi] is widened upwards first while cdf(hi) < p).
 */
export function invert(cdf: (x: number) => number, p: number, lo: number, hi: number): number {
  let guard = 0;
  while (cdf(hi) < p && guard++ < 2000) {
    lo = hi;
    hi *= 2;
  }
  for (let i = 0; i < 2000; i++) {
    const mid = (lo + hi) / 2;
    if (mid <= lo || mid >= hi) break;
    if (cdf(mid) < p) lo = mid;
    else hi = mid;
  }
  return (lo + hi) / 2;
}

/** The t ≥ 0 with P(T > t) = q (q ≤ 0,5), through the beta function's upper tail. */
function tInvTail(q: number, df: number): number {
  if (q >= 0.5) return 0;
  return invert((t) => -tailT(t, df), -q, 0, 1);
}

/** P(T > t) for t ≥ 0. */
function tailT(t: number, df: number): number {
  return 0.5 * ibeta(df / (df + t * t), df / 2, 0.5);
}

// ----- Beta ---------------------------------------------------------------------------------

function betaCdf(z: number, a: number, b: number): number {
  return ibeta(z, a, b);
}

function betaPdf(z: number, a: number, b: number): number {
  return Math.pow(z, a - 1) * Math.pow(1 - z, b - 1) * Math.exp(-lnBeta(a, b));
}

numFn("BETA.DIST", 4, 6, ([x, a, b, c, lo, hi]) => {
  if (a <= 0 || b <= 0 || x < lo || x > hi || lo === hi) return E.num;
  const z = (x - lo) / (hi - lo);
  return c ? betaCdf(z, a, b) : betaPdf(z, a, b) / (hi - lo);
}, [0, 0, 0, 0, 0, 1]);
numFn("BETADIST", 3, 5, ([x, a, b, lo, hi]) => {
  if (a <= 0 || b <= 0 || x < lo || x > hi || lo === hi) return E.num;
  return betaCdf((x - lo) / (hi - lo), a, b);
}, [0, 0, 0, 0, 1]);

function betaInv(p: number, a: number, b: number): number {
  return invert((z) => ibeta(z, a, b), p, 0, 1);
}

for (const name of ["BETA.INV", "BETAINV"]) {
  numFn(name, 3, 5, ([p, a, b, lo, hi]) => {
    if (p <= 0 || p >= 1 || a <= 0 || b <= 0 || lo >= hi) return E.num;
    return lo + (hi - lo) * betaInv(p, a, b);
  }, [0, 0, 0, 0, 1]);
}

// ----- Binomial family ----------------------------------------------------------------------

function binomPmf(k: number, nn: number, p: number): number {
  if (p === 0) return k === 0 ? 1 : 0;
  if (p === 1) return k === nn ? 1 : 0;
  if (nn < 1000) return combin(nn, k) * Math.pow(p, k) * Math.pow(1 - p, nn - k);
  return Math.exp(lnGamma(nn + 1) - lnGamma(k + 1) - lnGamma(nn - k + 1) + k * Math.log(p) + (nn - k) * Math.log(1 - p));
}

defS("BINOM.DIST.RANGE", 3, 4, (vals) => {
  const xs: number[] = [];
  for (const v of vals) {
    if (v === undefined) continue;
    const x = n(v);
    if (isErr(x)) return x;
    xs.push(x);
  }
  const N = Math.trunc(xs[0]), p = xs[1], s = Math.trunc(xs[2]);
  const s2 = xs.length > 3 ? Math.trunc(xs[3]) : s;
  if (N < 0 || p < 0 || p > 1 || s < 0 || s > N || s2 < s || s2 > N) return E.num;
  let r = 0;
  for (let k = s; k <= s2; k++) r += binomPmf(k, N, p);
  return num(r);
});

alias("BINOMDIST", "BINOM.DIST");
alias("CRITBINOM", "BINOM.INV");
alias("EXPONDIST", "EXPON.DIST");
alias("NORMDIST", "NORM.DIST");
alias("NORMINV", "NORM.INV");
alias("NORMSINV", "NORM.S.INV");
alias("POISSON", "POISSON.DIST");
alias("COVAR", "COVARIANCE.P");
alias("PEARSON", "CORREL");
alias("TTEST", "T.TEST");
alias("PERCENTRANK", "PERCENTRANK.INC");
numFn("NORMSDIST", 1, 1, ([z]) => normCdf(z));

numFn("NEGBINOM.DIST", 4, 4, ([f0, s0, p, c]) => {
  const f = Math.trunc(f0), s = Math.trunc(s0);
  if (f < 0 || s < 1 || p <= 0 || p >= 1) return E.num;
  const pmf = (k: number) => combin(k + s - 1, k) * Math.pow(p, s) * Math.pow(1 - p, k);
  if (!c) return pmf(f);
  if (f > 1000) return ibeta(p, s, f + 1);
  let r = 0;
  for (let k = 0; k <= f; k++) r += pmf(k);
  return Math.min(1, r);
});
numFn("NEGBINOMDIST", 3, 3, ([f0, s0, p]) => {
  const f = Math.trunc(f0), s = Math.trunc(s0);
  if (f < 0 || s < 1 || p <= 0 || p >= 1) return E.num;
  return combin(f + s - 1, f) * Math.pow(p, s) * Math.pow(1 - p, f);
});

function hypgeom(x: number[], cum: boolean): R {
  const s = Math.trunc(x[0]), k = Math.trunc(x[1]), K = Math.trunc(x[2]), N = Math.trunc(x[3]);
  if (s < 0 || k <= 0 || k > N || K <= 0 || K > N || N <= 0) return E.num;
  const pmf = (i: number) => (i < Math.max(0, k - N + K) || i > Math.min(k, K) ? 0 : (combin(K, i) * combin(N - K, k - i)) / combin(N, k));
  if (!cum) return pmf(s);
  let r = 0;
  for (let i = 0; i <= s; i++) r += pmf(i);
  return Math.min(1, r);
}
numFn("HYPGEOM.DIST", 5, 5, (x) => hypgeom(x, Boolean(x[4])));
numFn("HYPGEOMDIST", 4, 4, (x) => hypgeom(x, false));

// ----- Chi-squared --------------------------------------------------------------------------

function chiPdf(x: number, k: number): number {
  return Math.exp((k / 2 - 1) * Math.log(x) - x / 2 - (k / 2) * Math.LN2 - lnGamma(k / 2));
}

numFn("CHISQ.DIST", 3, 3, ([x, k0, c]) => {
  const k = Math.trunc(k0);
  if (x < 0 || k < 1 || k > 1e10) return E.num;
  if (c) return gammaP(k / 2, x / 2);
  if (x === 0) return k === 2 ? 0.5 : k === 1 ? E.num : 0;
  return chiPdf(x, k);
});
for (const name of ["CHISQ.DIST.RT", "CHIDIST"]) {
  numFn(name, 2, 2, ([x, k0]) => {
    const k = Math.trunc(k0);
    if (x < 0 || k < 1 || k > 1e10) return E.num;
    return gammaQ(k / 2, x / 2);
  });
}
numFn("CHISQ.INV", 2, 2, ([p, k0]) => {
  const k = Math.trunc(k0);
  if (p < 0 || p >= 1 || k < 1 || k > 1e10) return E.num;
  if (p === 0) return 0;
  return invert((x) => gammaP(k / 2, x / 2), p, 0, Math.max(1, k));
});
for (const name of ["CHISQ.INV.RT", "CHIINV"]) {
  numFn(name, 2, 2, ([p, k0]) => {
    const k = Math.trunc(k0);
    if (p <= 0 || p > 1 || k < 1 || k > 1e10) return E.num;
    if (p === 1) return 0;
    return invert((x) => -gammaQ(k / 2, x / 2), -p, 0, Math.max(1, k));
  });
}

for (const name of ["CHISQ.TEST", "CHITEST"]) {
  def(name, 2, 2, ([a, b], ctx) => {
    const A = toArr(a, ctx), B = toArr(b, ctx);
    if (isErr(A)) return A;
    if (isErr(B)) return B;
    if (A.rows !== B.rows || A.cols !== B.cols) return E.na;
    let chi = 0;
    for (let i = 0; i < A.d.length; i++) {
      const o = A.d[i], e = B.d[i];
      if (isErr(o)) return o;
      if (isErr(e)) return e;
      if (typeof o !== "number" || typeof e !== "number") continue;
      if (e === 0) return E.div0;
      chi += ((o - e) * (o - e)) / e;
    }
    const df = A.rows > 1 && A.cols > 1 ? (A.rows - 1) * (A.cols - 1) : A.rows * A.cols - 1;
    if (df < 1) return E.na;
    return num(gammaQ(df / 2, chi / 2));
  });
}

// ----- F ------------------------------------------------------------------------------------

function fCdf(x: number, d1: number, d2: number): number {
  return ibeta((d1 * x) / (d1 * x + d2), d1 / 2, d2 / 2);
}
function fTail(x: number, d1: number, d2: number): number {
  return ibeta(d2 / (d2 + d1 * x), d2 / 2, d1 / 2);
}

numFn("F.DIST", 4, 4, ([x, a, b, c]) => {
  const d1 = Math.trunc(a), d2 = Math.trunc(b);
  if (x < 0 || d1 < 1 || d2 < 1 || d1 >= 1e10 || d2 >= 1e10) return E.num;
  if (c) return fCdf(x, d1, d2);
  return Math.pow(d1 / d2, d1 / 2) * Math.pow(x, d1 / 2 - 1) * Math.pow(1 + (d1 * x) / d2, -(d1 + d2) / 2) * Math.exp(-lnBeta(d1 / 2, d2 / 2));
});
for (const name of ["F.DIST.RT", "FDIST"]) {
  numFn(name, 3, 3, ([x, a, b]) => {
    const d1 = Math.trunc(a), d2 = Math.trunc(b);
    if (x < 0 || d1 < 1 || d2 < 1 || d1 >= 1e10 || d2 >= 1e10) return E.num;
    return fTail(x, d1, d2);
  });
}
numFn("F.INV", 3, 3, ([p, a, b]) => {
  const d1 = Math.trunc(a), d2 = Math.trunc(b);
  if (p < 0 || p >= 1 || d1 < 1 || d2 < 1 || d1 >= 1e10 || d2 >= 1e10) return E.num;
  if (p === 0) return 0;
  return invert((x) => fCdf(x, d1, d2), p, 0, 1);
});
for (const name of ["F.INV.RT", "FINV"]) {
  numFn(name, 3, 3, ([p, a, b]) => {
    const d1 = Math.trunc(a), d2 = Math.trunc(b);
    if (p <= 0 || p > 1 || d1 < 1 || d2 < 1 || d1 >= 1e10 || d2 >= 1e10) return E.num;
    if (p === 1) return 0;
    return invert((x) => -fTail(x, d1, d2), -p, 0, 1);
  });
}

for (const name of ["F.TEST", "FTEST"]) {
  def(name, 2, 2, ([a, b], ctx) => {
    const xa = nums(a, ctx), xb = nums(b, ctx);
    if (isErr(xa)) return xa;
    if (isErr(xb)) return xb;
    if (xa.length < 2 || xb.length < 2) return E.div0;
    const va = ssd(xa) / (xa.length - 1), vb = ssd(xb) / (xb.length - 1);
    if (va === 0 || vb === 0) return E.div0;
    const f = va / vb, d1 = xa.length - 1, d2 = xb.length - 1;
    const lower = fCdf(f, d1, d2), upper = fTail(f, d1, d2);
    return num(2 * Math.min(lower, upper));
  });
}

// ----- Fisher, gamma, Gauss, Phi ------------------------------------------------------------

numFn("FISHER", 1, 1, ([x]) => (x <= -1 || x >= 1 ? E.num : 0.5 * Math.log((1 + x) / (1 - x))));
numFn("FISHERINV", 1, 1, ([y]) => Math.tanh(y));

numFn("GAMMA", 1, 1, ([x]) => {
  if (x <= 0 && Number.isInteger(x)) return E.num;
  if (x > 171.6243769563027) return E.num;
  return gamma(x);
});
for (const name of ["GAMMALN", "GAMMALN.PRECISE"]) {
  numFn(name, 1, 1, ([x]) => (x <= 0 ? E.num : x === 1 || x === 2 ? 0 : lnGamma(x)));
}
for (const name of ["GAMMA.DIST", "GAMMADIST"]) {
  numFn(name, 4, 4, ([x, a, b, c]) => {
    if (x < 0 || a <= 0 || b <= 0) return E.num;
    if (c) return gammaP(a, x / b);
    if (x === 0) return E.num;
    return Math.exp((a - 1) * Math.log(x) - x / b - a * Math.log(b) - lnGamma(a));
  });
}
for (const name of ["GAMMA.INV", "GAMMAINV"]) {
  numFn(name, 3, 3, ([p, a, b]) => {
    if (p < 0 || p >= 1 || a <= 0 || b <= 0) return E.num;
    if (p === 0) return 0;
    return b * invert((x) => gammaP(a, x), p, 0, Math.max(1, a));
  });
}
numFn("GAUSS", 1, 1, ([z]) => normCdf(z) - 0.5);
numFn("PHI", 1, 1, ([x]) => normPdf(x));

// ----- Lognormal, Weibull --------------------------------------------------------------------

numFn("LOGNORM.DIST", 4, 4, ([x, m, s, c]) => {
  if (x <= 0 || s <= 0) return E.num;
  const z = (Math.log(x) - m) / s;
  return c ? normCdf(z) : normPdf(z) / (x * s);
});
numFn("LOGNORMDIST", 3, 3, ([x, m, s]) => (x <= 0 || s <= 0 ? E.num : normCdf((Math.log(x) - m) / s)));
for (const name of ["LOGNORM.INV", "LOGINV"]) {
  numFn(name, 3, 3, ([p, m, s]) => (p <= 0 || p >= 1 || s <= 0 ? E.num : Math.exp(m + s * normInv(p))));
}
for (const name of ["WEIBULL.DIST", "WEIBULL"]) {
  numFn(name, 4, 4, ([x, a, b, c]) => {
    if (x < 0 || a <= 0 || b <= 0) return E.num;
    const t = Math.pow(x / b, a);
    if (c) return -Math.expm1(-t);
    if (x === 0) return 0;
    return (a / x) * t * Math.exp(-t);
  });
}

// ----- Confidence ----------------------------------------------------------------------------

for (const name of ["CONFIDENCE.NORM", "CONFIDENCE"]) {
  numFn(name, 3, 3, ([a, s, n0]) => {
    const k = Math.trunc(n0);
    if (a <= 0 || a >= 1 || s <= 0 || k < 1) return E.num;
    return (normInv(1 - a / 2) * s) / Math.sqrt(k);
  });
}
numFn("CONFIDENCE.T", 3, 3, ([a, s, n0]) => {
  const k = Math.trunc(n0);
  if (a <= 0 || a >= 1 || s <= 0 || k < 1) return E.num;
  if (k === 1) return E.div0;
  return (tInvTail(a / 2, k - 1) * s) / Math.sqrt(k);
});

// ----- t --------------------------------------------------------------------------------------

numFn("T.DIST.2T", 2, 2, ([x, d]) => {
  const df = Math.trunc(d);
  if (x < 0 || df < 1) return E.num;
  return 2 * tailT(x, df);
});
numFn("T.DIST.RT", 2, 2, ([x, d]) => {
  const df = Math.trunc(d);
  if (df < 1) return E.num;
  return x >= 0 ? tailT(x, df) : 1 - tailT(-x, df);
});
numFn("TDIST", 3, 3, ([x, d, t0]) => {
  const df = Math.trunc(d), tails = Math.trunc(t0);
  if (x < 0 || df < 1 || (tails !== 1 && tails !== 2)) return E.num;
  return tails * tailT(x, df);
});
numFn("TINV", 2, 2, ([p, d]) => {
  const df = Math.trunc(d);
  if (p <= 0 || p > 1 || df < 1) return E.num;
  return tInvTail(p / 2, df);
});

// ----- Z test ---------------------------------------------------------------------------------

for (const name of ["Z.TEST", "ZTEST"]) {
  def(name, 2, 3, ([a, x0, s0], ctx) => {
    const xs = nums(a, ctx);
    if (isErr(xs)) return xs;
    const xv = n(deref(x0, ctx) as V);
    if (isErr(xv)) return xv;
    let sigma: number;
    if (s0 !== undefined && s0 !== null) {
      const s = n(deref(s0, ctx) as V);
      if (isErr(s)) return s;
      sigma = s;
    } else {
      if (xs.length < 2) return E.div0;
      sigma = Math.sqrt(ssd(xs) / (xs.length - 1));
    }
    if (!xs.length) return E.na;
    if (sigma === 0) return E.div0;
    const z = (sum(xs) / xs.length - xv) / (sigma / Math.sqrt(xs.length));
    return num(1 - normCdf(z));
  });
}

// ----- Aggregates -----------------------------------------------------------------------------

function agg(name: string, f: (xs: number[]) => R, a = false): void {
  def(name, 1, 255, (args, ctx) => {
    const xs = numbers(args, ctx, a ? { a: true } : {});
    if (isErr(xs)) return xs;
    const r = f(xs);
    return typeof r === "number" ? num(r) : r;
  });
}

agg("DEVSQ", (xs) => (xs.length ? ssd(xs) : E.num));
agg("STDEVPA", (xs) => (xs.length ? Math.sqrt(ssd(xs) / xs.length) : E.div0), true);
agg("VARPA", (xs) => (xs.length ? ssd(xs) / xs.length : E.div0), true);
agg("SKEW.P", (xs) => {
  const k = xs.length;
  if (k < 1) return E.div0;
  const m = sum(xs) / k;
  const sd = Math.sqrt(ssd(xs) / k);
  if (sd === 0) return E.div0;
  return sum(xs.map((x) => ((x - m) / sd) ** 3)) / k;
});

def("MODE.MULT", 1, 255, (args, ctx) => {
  const xs = numbers(args, ctx, { skipText: true });
  if (isErr(xs)) return xs;
  const count = new Map<number, number>();
  let best = 0;
  for (const x of xs) {
    const c = (count.get(x) ?? 0) + 1;
    count.set(x, c);
    best = Math.max(best, c);
  }
  if (best < 2) return E.na;
  const out: number[] = [];
  for (const x of xs) if (count.get(x) === best && !out.includes(x)) out.push(x);
  return new Arr(out.length, 1, out);
});

defS("TRIMMEAN", 2, 2, ([, p0], ctx, raw) => {
  const xs = nums(raw[0], ctx);
  if (isErr(xs)) return xs;
  const p = n(p0);
  if (isErr(p)) return p;
  if (p < 0 || p >= 1 || !xs.length) return E.num;
  const k = Math.floor((xs.length * p) / 2);
  const s = [...xs].sort((a, b) => a - b).slice(k, xs.length - k);
  return num(sum(s) / s.length);
}, [0]);

defS("PERCENTRANK.EXC", 2, 3, ([, x0, sig0], ctx, raw) => {
  const xs = nums(raw[0], ctx);
  if (isErr(xs)) return xs;
  const x = n(x0);
  if (isErr(x)) return x;
  const sig = sig0 === undefined || sig0 === null ? 3 : n(sig0);
  if (isErr(sig)) return sig;
  if (sig < 1) return E.num;
  const s = [...xs].sort((a, b) => a - b);
  if (!s.length || x < s[0] || x > s[s.length - 1]) return E.na;
  const i = s.findIndex((v) => v >= x);
  const pos = s[i] === x ? i : i - 1 + (x - s[i - 1]) / (s[i] - s[i - 1]);
  const r = (pos + 1) / (s.length + 1);
  const f = 10 ** Math.trunc(sig);
  return Math.floor(r * f + 1e-9) / f;
}, [0]);

numFn("PERMUTATIONA", 2, 2, ([a, b]) => {
  const nn = Math.trunc(a), k = Math.trunc(b);
  if (nn < 0 || k < 0) return E.num;
  return Math.pow(nn, k);
});

def("PROB", 3, 4, ([x0, p0, lo0, hi0], ctx) => {
  const xv = values(x0, ctx), pv = values(p0, ctx);
  if (xv.length !== pv.length) return E.na;
  const lo = n(deref(lo0, ctx) as V);
  if (isErr(lo)) return lo;
  let hi = lo;
  if (hi0 !== undefined && hi0 !== null) {
    const h = n(deref(hi0, ctx) as V);
    if (isErr(h)) return h;
    hi = h;
  }
  let total = 0, r = 0;
  for (let i = 0; i < xv.length; i++) {
    const x = xv[i], p = pv[i];
    if (isErr(x)) return x;
    if (isErr(p)) return p;
    if (typeof p !== "number" || typeof x !== "number") continue;
    if (p < 0 || p > 1) return E.num;
    total += p;
    if (x >= lo && x <= hi) r += p;
  }
  if (Math.abs(total - 1) > 1e-7) return E.num;
  return num(r);
});

// ----- Regression: RGP, RKP, TREND, VARIATION -------------------------------------------------

interface Fit {
  coef: number[]; // b0 (intercept, 0 without constant), then one per variable
  se: number[];
  r2: number;
  sey: number;
  F: number;
  df: number;
  ssreg: number;
  ssresid: number;
  k: number;
}

/** The numbers of an array, or #WERT! when one is no number. */
function numArr(a: Arr): number[] | CellError {
  const out: number[] = [];
  for (const v of a.d) {
    if (isErr(v)) return v;
    if (typeof v !== "number") return E.value;
    out.push(v);
  }
  return out;
}

interface Design { y: number[]; X: number[][]; byRow: boolean; vars: number; }

/** y and the matrix of the known x (one row per observation), as RGP pairs them. */
function design(yA: Arr, xA: Arr | null): Design | CellError {
  const y = numArr(yA);
  if (isErr(y)) return y;
  const nObs = y.length;
  if (!xA) return { y, X: y.map((_, i) => [i + 1]), byRow: yA.rows === 1 && yA.cols > 1, vars: 1 };
  const xs = numArr(xA);
  if (isErr(xs)) return xs;
  const yCol = yA.cols === 1, yRow = yA.rows === 1;
  if (xA.rows * xA.cols === nObs && (xA.rows === 1 || xA.cols === 1)) return { y, X: xs.map((v) => [v]), byRow: yRow && !yCol, vars: 1 };
  if (yCol && xA.rows === nObs) {
    const X = [];
    for (let r = 0; r < nObs; r++) X.push(xs.slice(r * xA.cols, (r + 1) * xA.cols));
    return { y, X, byRow: false, vars: xA.cols };
  }
  if (yRow && xA.cols === nObs) {
    const X = [];
    for (let c = 0; c < nObs; c++) {
      const row = [];
      for (let r = 0; r < xA.rows; r++) row.push(xs[r * xA.cols + c]);
      X.push(row);
    }
    return { y, X, byRow: true, vars: xA.rows };
  }
  return E.ref;
}

/** Least squares by Householder QR; columns that add nothing are dropped (coefficient 0). */
function fit(y: number[], X: number[][], konst: boolean): Fit | CellError {
  const nObs = y.length, k = X[0]?.length ?? 0;
  const cols: number[][] = [];
  if (konst) cols.push(new Array(nObs).fill(1));
  for (let j = 0; j < k; j++) cols.push(X.map((r) => r[j]));
  const p = cols.length;
  // Gram–Schmidt with re-orthogonalisation, dropping dependent columns.
  const Q: number[][] = [], keep: number[] = [];
  const Rm: number[][] = [];
  for (let j = 0; j < p; j++) {
    const v = cols[j].slice();
    const rcol: number[] = [];
    for (let pass = 0; pass < 2; pass++) {
      for (let i = 0; i < Q.length; i++) {
        let d = 0;
        for (let t = 0; t < nObs; t++) d += Q[i][t] * v[t];
        rcol[i] = (rcol[i] ?? 0) + d;
        for (let t = 0; t < nObs; t++) v[t] -= d * Q[i][t];
      }
    }
    const norm = Math.hypot(...v), orig = Math.hypot(...cols[j]);
    if (norm <= 1e-12 * Math.max(orig, 1e-300)) continue;
    Q.push(v.map((x) => x / norm));
    rcol[Q.length - 1] = norm;
    Rm.push(rcol);
    keep.push(j);
  }
  const m = Q.length;
  // Solve R b = Qᵀ y (R upper triangular: Rm[j][i] is R[i][j]).
  const qty = Q.map((q) => q.reduce((s, x, t) => s + x * y[t], 0));
  const b = new Array(m).fill(0);
  for (let i = m - 1; i >= 0; i--) {
    let s = qty[i];
    for (let j = i + 1; j < m; j++) s -= Rm[j][i] * b[j];
    b[i] = s / Rm[i][i];
  }
  // (RᵀR)⁻¹ diagonal through R⁻¹.
  const Rinv: number[][] = Array.from({ length: m }, () => new Array(m).fill(0));
  for (let c = 0; c < m; c++) {
    for (let i = m - 1; i >= 0; i--) {
      let s = i === c ? 1 : 0;
      for (let j = i + 1; j < m; j++) s -= Rm[j][i] * Rinv[j][c];
      Rinv[i][c] = s / Rm[i][i];
    }
  }
  const coefAll = new Array(p).fill(0), varDiag = new Array(p).fill(0);
  keep.forEach((j, i) => {
    coefAll[j] = b[i];
    let s = 0;
    for (let c = 0; c < m; c++) s += Rinv[i][c] * Rinv[i][c];
    varDiag[j] = s;
  });
  let ssresid = 0;
  for (let t = 0; t < nObs; t++) {
    let fitv = 0;
    for (let j = 0; j < p; j++) fitv += coefAll[j] * cols[j][t];
    ssresid += (y[t] - fitv) ** 2;
  }
  const my = sum(y) / nObs;
  const sstot = konst ? y.reduce((s, v) => s + (v - my) ** 2, 0) : y.reduce((s, v) => s + v * v, 0);
  const df = nObs - m;
  const ssreg = sstot - ssresid;
  const s2 = df > 0 ? ssresid / df : NaN;
  const coef = konst ? coefAll : [0, ...coefAll];
  const se = (konst ? varDiag : [NaN, ...varDiag]).map((v) => Math.sqrt(v * s2));
  const kReg = konst ? m - 1 : m;
  return { coef, se, r2: sstot === 0 ? 1 : ssreg / sstot, sey: Math.sqrt(s2), F: (ssreg / kReg) / s2, df, ssreg, ssresid, k };
}

function boolArg(a: Any | undefined, ctx: Ctx, d: boolean): boolean | CellError {
  if (a === undefined || a === null) return d;
  const v = deref(a, ctx);
  const x = n(v instanceof Arr ? v.d[0] : v);
  return isErr(x) ? x : x !== 0;
}

function regression(name: string, log: boolean): void {
  def(name, 1, 4, ([y0, x0, c0, s0], ctx) => {
    const yA = toArr(y0, ctx);
    if (isErr(yA)) return yA;
    const xA = x0 === undefined || x0 === null ? null : toArr(x0, ctx);
    if (isErr(xA)) return xA;
    const konst = boolArg(c0, ctx, true), stats = boolArg(s0, ctx, false);
    if (isErr(konst)) return konst;
    if (isErr(stats)) return stats;
    const d = design(yA, xA);
    if (isErr(d)) return d;
    if (log && d.y.some((v) => v <= 0)) return E.num;
    const f = fit(log ? d.y.map(Math.log) : d.y, d.X, konst);
    if (isErr(f)) return f;
    const k = d.vars;
    const tr = (v: number) => (log ? Math.exp(v) : v);
    const row1: V[] = [];
    for (let j = k; j >= 1; j--) row1.push(num(tr(f.coef[j])) as V);
    row1.push(num(tr(f.coef[0])) as V);
    if (!stats) return new Arr(1, k + 1, row1);
    const na = E.na;
    const fill = (vals: V[]) => [...vals, ...new Array(k + 1 - vals.length).fill(na)];
    const se: V[] = [];
    for (let j = k; j >= 1; j--) se.push(Number.isFinite(f.se[j]) ? f.se[j] : E.num);
    se.push(konst ? (Number.isFinite(f.se[0]) ? f.se[0] : E.num) : na);
    const d0 = (v: number): V => (Number.isFinite(v) ? v : E.num);
    const rows = [row1, se, fill([d0(f.r2), d0(f.sey)]), fill([d0(f.F), f.df]), fill([d0(f.ssreg), d0(f.ssresid)])];
    return new Arr(5, k + 1, rows.flat());
  });
}
regression("LINEST", false);
regression("LOGEST", true);

function predict(name: string, log: boolean): void {
  def(name, 1, 4, ([y0, x0, nx0, c0], ctx) => {
    const yA = toArr(y0, ctx);
    if (isErr(yA)) return yA;
    const xA = x0 === undefined || x0 === null ? null : toArr(x0, ctx);
    if (isErr(xA)) return xA;
    const konst = boolArg(c0, ctx, true);
    if (isErr(konst)) return konst;
    const d = design(yA, xA);
    if (isErr(d)) return d;
    if (log && d.y.some((v) => v <= 0)) return E.num;
    const f = fit(log ? d.y.map(Math.log) : d.y, d.X, konst);
    if (isErr(f)) return f;
    const at = (xs: number[]) => {
      let s = f.coef[0];
      for (let j = 0; j < xs.length; j++) s += f.coef[j + 1] * xs[j];
      return num(log ? Math.exp(s) : s);
    };
    let nxA: Arr;
    if (nx0 === undefined || nx0 === null) nxA = xA ?? new Arr(yA.rows, yA.cols, d.X.map((r) => r[0]));
    else {
      const t = toArr(nx0, ctx);
      if (isErr(t)) return t;
      nxA = t;
    }
    const nx = numArr(nxA);
    if (isErr(nx)) return nx;
    if (d.vars === 1) return new Arr(nxA.rows, nxA.cols, nx.map((v) => at([v])));
    if (!d.byRow) {
      if (nxA.cols !== d.vars) return E.ref;
      const out: V[] = [];
      for (let r = 0; r < nxA.rows; r++) out.push(at(nx.slice(r * nxA.cols, (r + 1) * nxA.cols)));
      return new Arr(nxA.rows, 1, out);
    }
    if (nxA.rows !== d.vars) return E.ref;
    const out: V[] = [];
    for (let c = 0; c < nxA.cols; c++) {
      const xs = [];
      for (let r = 0; r < nxA.rows; r++) xs.push(nx[r * nxA.cols + c]);
      out.push(at(xs));
    }
    return new Arr(1, nxA.cols, out);
  });
}
predict("TREND", false);
predict("GROWTH", true);

