// Financial functions, second round: securities (coupons, KURS, RENDITE, DURATION, odd first and
// last periods, discounted and at-maturity papers, treasury bills), depreciation (GDA, GDA2,
// VDB, LIA, DIA, AMORLINEARK), cash flows (IKV, QIKV, XINTZINSFUSS, XKAPITALWERT) and the
// smaller ones (EFFEKTIV, NOMINAL, KUMZINSZ, KUMKAPITAL, ISPMT, ZW2, PDURATION, ZSATZINVEST …).

import { def, deref } from "./eval";
import { values } from "./fn-util";
import { numFn } from "./fn-stats2";
import { daysInMonth, serial, ymd } from "./dates";
import { Arr, E, isErr, num, type V } from "./values";
import type { CellError } from "../model";

// ----- Day counts -----------------------------------------------------------------------------

const leap = (y: number) => y % 4 === 0 && (y % 100 !== 0 || y % 400 === 0);
const lastDay = (y: number, m: number) => (m === 2 ? (leap(y) ? 29 : 28) : daysInMonth(y, m));
const isLastFeb = (d: { y: number; m: number; d: number }) => d.m === 2 && d.d === lastDay(d.y, 2);

/** Days between two dates by the 30/360 US (NASD) rule, as YEARFRAC's basis 0 counts them. */
function days360us(s: number, e: number): number {
  const A = ymd(s), B = ymd(e);
  let d1 = A.d, d2 = B.d;
  if (isLastFeb(A) && isLastFeb(B)) d2 = 30;
  if (isLastFeb(A)) d1 = 30;
  if (d2 === 31 && d1 >= 30) d2 = 30;
  if (d1 === 31) d1 = 30;
  return (B.y - A.y) * 360 + (B.m - A.m) * 30 + (d2 - d1);
}

/** Days between two dates by the European 30/360 rule (basis 4). */
function days360eu(s: number, e: number): number {
  const A = ymd(s), B = ymd(e);
  return (B.y - A.y) * 360 + (B.m - A.m) * 30 + (Math.min(B.d, 30) - Math.min(A.d, 30));
}

function dayCount(s: number, e: number, basis: number): number {
  if (basis === 0) return days360us(s, e);
  if (basis === 4) return days360eu(s, e);
  return e - s;
}

/** YEARFRAC as Excel computes it (s ≤ e). */
function yearFrac(s: number, e: number, basis: number): number {
  if (s > e) [s, e] = [e, s];
  const A = ymd(s), B = ymd(e);
  switch (basis) {
    case 0: return days360us(s, e) / 360;
    case 1: {
      if (A.y === B.y || (B.y === A.y + 1 && (A.m > B.m || (A.m === B.m && A.d >= B.d)))) {
        let den = 365;
        if (A.y === B.y) den = leap(A.y) ? 366 : 365;
        else {
          const f1 = leap(A.y) ? serial(A.y, 2, 29) : 0, f2 = leap(B.y) ? serial(B.y, 2, 29) : 0;
          if ((f1 && s <= f1) || (f2 && e >= f2)) den = 366;
        }
        return (e - s) / den;
      }
      let days = 0;
      for (let y = A.y; y <= B.y; y++) days += leap(y) ? 366 : 365;
      return (e - s) / (days / (B.y - A.y + 1));
    }
    case 2: return (e - s) / 360;
    case 3: return (e - s) / 365;
    default: return days360eu(s, e) / 360;
  }
}

/** The date `k` months from `d` (k may be negative); month ends stay month ends when `eom`. */
function addMonths(d: number, k: number, eom: boolean): number {
  const t = ymd(d);
  const total = t.y * 12 + (t.m - 1) + k;
  const y = Math.floor(total / 12), m = total - y * 12 + 1;
  const dim = lastDay(y, m);
  return serial(y, m, eom ? dim : Math.min(t.d, dim));
}

const isEom = (d: number) => {
  const t = ymd(d);
  return t.d === lastDay(t.y, t.m);
};

// ----- Coupons --------------------------------------------------------------------------------

interface Coupon { pcd: number; ncd: number; num: number; }

/** The coupon dates around the settlement: steps of 12/f months back from the maturity. */
function coupons(settle: number, mat: number, freq: number): Coupon {
  const eom = isEom(mat), step = 12 / freq;
  let k = 1;
  while (addMonths(mat, -k * step, eom) > settle) k++;
  return { pcd: addMonths(mat, -k * step, eom), ncd: addMonths(mat, -(k - 1) * step, eom), num: k };
}

function couponDays(c: Coupon, basis: number, freq: number): number {
  // Actual/actual: the period from the previous coupon date 12/f months on, without the
  // month-end rule (29.2.2012 to 29.8.2012 is 182 days although the next coupon is on 31.8.).
  if (basis === 1) return addMonths(c.pcd, 12 / freq, false) - c.pcd;
  return (basis === 3 ? 365 : 360) / freq;
}

interface Sec { settle: number; mat: number; freq: number; basis: number; }

function secArgs(s0: number, m0: number, f0: number, b0: number): Sec | CellError {
  const settle = Math.trunc(s0), mat = Math.trunc(m0), freq = Math.trunc(f0), basis = Math.trunc(b0);
  if (settle >= mat || ![1, 2, 4].includes(freq) || basis < 0 || basis > 4 || settle < 0) return E.num;
  return { settle, mat, freq, basis };
}

/**
 * E (days in the coupon period), A (days from its start), N, and the days to the next coupon:
 * ZINSTERMTAGNZ counts them (`DSCdays`, actual days but for the 30/360 bases), the prices take
 * E − A except for actual/actual (`DSC`).
 */
function couponFacts(a: Sec) {
  const c = coupons(a.settle, a.mat, a.freq);
  const E_ = couponDays(c, a.basis, a.freq);
  const A = dayCount(c.pcd, a.settle, a.basis);
  const DSCdays = a.basis === 0 || a.basis === 4 ? E_ - A : c.ncd - a.settle;
  const DSC = a.basis === 1 ? c.ncd - a.settle : E_ - A;
  return { c, E: E_, A, DSC, DSCdays, N: c.num };
}

function coupFn(name: string, f: (a: Sec) => number): void {
  numFn(name, 3, 4, ([s, m, fr, b]) => {
    const a = secArgs(s, m, fr, b);
    return isErr(a) ? a : f(a);
  });
}
coupFn("COUPDAYBS", (a) => couponFacts(a).A);
coupFn("COUPDAYS", (a) => couponFacts(a).E);
coupFn("COUPDAYSNC", (a) => couponFacts(a).DSCdays);
coupFn("COUPNCD", (a) => coupons(a.settle, a.mat, a.freq).ncd);
coupFn("COUPPCD", (a) => coupons(a.settle, a.mat, a.freq).pcd);
coupFn("COUPNUM", (a) => coupons(a.settle, a.mat, a.freq).num);

// ----- KURS, RENDITE, DURATION ----------------------------------------------------------------

function price(a: Sec, rate: number, yld: number, red: number): number {
  const { E: e, A, DSC, N } = couponFacts(a);
  const f = a.freq, cpn = (100 * rate) / f, dsc = DSC / e;
  if (N === 1) return (cpn + red) / (1 + (dsc * yld) / f) - (cpn * A) / e;
  const v = 1 + yld / f;
  let p = red / Math.pow(v, N - 1 + dsc);
  for (let k = 1; k <= N; k++) p += cpn / Math.pow(v, k - 1 + dsc);
  return p - (cpn * A) / e;
}

/** The root of an increasing or decreasing function by Newton steps, falling back to bisection. */
function solve(f: (x: number) => number, guess: number, lo: number, hi: number): number | null {
  let x = guess;
  for (let i = 0; i < 100; i++) {
    const y = f(x);
    if (!Number.isFinite(y)) break;
    if (y === 0) return x;
    const h = Math.max(1e-7, Math.abs(x) * 1e-7);
    const d = (f(x + h) - f(x - h)) / (2 * h);
    if (!d || !Number.isFinite(d)) break;
    const nx = x - y / d;
    if (!Number.isFinite(nx)) break;
    if (Math.abs(nx - x) <= 1e-15 * Math.max(1, Math.abs(nx))) return nx;
    x = nx;
  }
  // Bisection on a sign change.
  let a = lo, b = hi, fa = f(a), fb = f(b);
  if (!(fa * fb <= 0)) return Number.isFinite(x) && Math.abs(f(x)) < 1e-9 ? x : null;
  for (let i = 0; i < 300; i++) {
    const m = (a + b) / 2;
    const fm = f(m);
    if (fm === 0 || b - a < 1e-16 * Math.max(1, Math.abs(m))) return m;
    if (fa * fm < 0) { b = m; fb = fm; } else { a = m; fa = fm; }
  }
  return (a + b) / 2;
}

numFn("PRICE", 6, 7, ([s, m, rate, yld, red, f, b]) => {
  const a = secArgs(s, m, f, b);
  if (isErr(a)) return a;
  if (rate < 0 || yld < 0 || red <= 0) return E.num;
  return price(a, rate, yld, red);
});

numFn("YIELD", 6, 7, ([s, m, rate, pr, red, f, b]) => {
  const a = secArgs(s, m, f, b);
  if (isErr(a)) return a;
  if (rate < 0 || pr <= 0 || red <= 0) return E.num;
  const { E: e, A, DSC, N } = couponFacts(a);
  const fq = a.freq;
  if (N === 1) {
    const cpn = rate / fq, ae = A / e;
    return (((red / 100 + cpn) - (pr / 100 + ae * cpn)) / (pr / 100 + ae * cpn)) * ((fq * e) / DSC);
  }
  const y = solve((x) => price(a, rate, x, red) - pr, 0.1, -0.99 * fq, 10);
  return y === null ? E.num : y;
});

function duration(a: Sec, rate: number, yld: number): number {
  const { E: e, DSC, N } = couponFacts(a);
  const f = a.freq, cpn = (100 * rate) / f, dsc = DSC / e, v = 1 + yld / f;
  let num = 0, den = 0;
  for (let k = 1; k <= N; k++) {
    const t = k - 1 + dsc;
    const cf = cpn + (k === N ? 100 : 0);
    const pv = cf / Math.pow(v, t);
    num += t * pv;
    den += pv;
  }
  return num / den / f;
}

numFn("DURATION", 5, 6, ([s, m, rate, yld, f, b]) => {
  const a = secArgs(s, m, f, b);
  if (isErr(a)) return a;
  if (rate < 0 || yld < 0) return E.num;
  return duration(a, rate, yld);
});
numFn("MDURATION", 5, 6, ([s, m, rate, yld, f, b]) => {
  const a = secArgs(s, m, f, b);
  if (isErr(a)) return a;
  if (rate < 0 || yld < 0) return E.num;
  return duration(a, rate, yld) / (1 + yld / a.freq);
});

// ----- Accrued interest -----------------------------------------------------------------------

numFn("ACCRINT", 6, 8, ([i0, fi0, s0, rate, par, f0, b0]) => {
  const issue = Math.trunc(i0), first = Math.trunc(fi0), settle = Math.trunc(s0), freq = Math.trunc(f0), basis = Math.trunc(b0);
  if (rate <= 0 || par <= 0 || ![1, 2, 4].includes(freq) || basis < 0 || basis > 4 || issue >= settle) return E.num;
  if (basis !== 1) return par * rate * yearFrac(issue, settle, basis);
  // Actual/actual: quasi-coupon periods anchored at the first interest date, each counted
  // against its own length.
  const step = 12 / freq, eom = isEom(first);
  let k = 0;
  while (addMonths(first, k * step, eom) > issue) k--;
  let total = 0;
  for (;; k++) {
    const p0 = addMonths(first, k * step, eom), p1 = addMonths(first, (k + 1) * step, eom);
    if (p0 >= settle) break;
    const from = Math.max(p0, issue), to = Math.min(p1, settle);
    if (to > from) total += (to - from) / (p1 - p0);
  }
  return (par * rate * total) / freq;
}, [0, 0, 0, 0, 0, 0, 0, 1]);

numFn("ACCRINTM", 4, 5, ([i0, s0, rate, par, b0]) => {
  const issue = Math.trunc(i0), settle = Math.trunc(s0), basis = Math.trunc(b0);
  if (rate <= 0 || par <= 0 || basis < 0 || basis > 4 || issue >= settle) return E.num;
  return par * rate * yearFrac(issue, settle, basis);
});

// ----- Discounted papers ----------------------------------------------------------------------

interface Disc { settle: number; mat: number; basis: number; }
function discArgs(s0: number, m0: number, b0: number): Disc | CellError {
  const settle = Math.trunc(s0), mat = Math.trunc(m0), basis = Math.trunc(b0);
  if (settle >= mat || basis < 0 || basis > 4) return E.num;
  return { settle, mat, basis };
}
const yf = (d: Disc) => yearFrac(d.settle, d.mat, d.basis);

numFn("DISC", 4, 5, ([s, m, pr, red, b]) => {
  const d = discArgs(s, m, b);
  if (isErr(d)) return d;
  if (pr <= 0 || red <= 0) return E.num;
  return (red - pr) / red / yf(d);
});
numFn("INTRATE", 4, 5, ([s, m, inv, red, b]) => {
  const d = discArgs(s, m, b);
  if (isErr(d)) return d;
  if (inv <= 0 || red <= 0) return E.num;
  return (red - inv) / inv / yf(d);
});
numFn("RECEIVED", 4, 5, ([s, m, inv, disc, b]) => {
  const d = discArgs(s, m, b);
  if (isErr(d)) return d;
  if (inv <= 0 || disc <= 0) return E.num;
  const den = 1 - disc * yf(d);
  return den <= 0 ? E.num : inv / den;
});
numFn("PRICEDISC", 4, 5, ([s, m, disc, red, b]) => {
  const d = discArgs(s, m, b);
  if (isErr(d)) return d;
  if (disc <= 0 || red <= 0) return E.num;
  return red - disc * red * yf(d);
});
numFn("YIELDDISC", 4, 5, ([s, m, pr, red, b]) => {
  const d = discArgs(s, m, b);
  if (isErr(d)) return d;
  if (pr <= 0 || red <= 0) return E.num;
  return (red - pr) / pr / yf(d);
});

/** Days in the year for the papers paid at maturity; for actual/actual KURSFÄLLIG takes 365
 *  and RENDITEFÄLL the settlement's year (both recorded in a leap year). */
function yearDays(basis: number, settle: number, price: boolean): number {
  if (basis === 1) return price || !leap(ymd(settle).y) ? 365 : 366;
  return basis === 3 ? 365 : 360;
}

numFn("PRICEMAT", 5, 6, ([s, m, i, rate, yld, b]) => {
  const d = discArgs(s, m, b);
  if (isErr(d)) return d;
  const issue = Math.trunc(i);
  if (rate < 0 || yld < 0 || issue >= d.settle) return E.num;
  const B = yearDays(d.basis, d.settle, true);
  const dim = dayCount(issue, d.mat, d.basis) / B, a = dayCount(issue, d.settle, d.basis) / B, dsm = dayCount(d.settle, d.mat, d.basis) / B;
  return (100 + dim * rate * 100) / (1 + dsm * yld) - a * rate * 100;
});
numFn("YIELDMAT", 5, 6, ([s, m, i, rate, pr, b]) => {
  const d = discArgs(s, m, b);
  if (isErr(d)) return d;
  const issue = Math.trunc(i);
  if (rate < 0 || pr <= 0 || issue >= d.settle) return E.num;
  const B = yearDays(d.basis, d.settle, false);
  const dim = dayCount(issue, d.mat, d.basis) / B, a = dayCount(issue, d.settle, d.basis) / B, dsm = dayCount(d.settle, d.mat, d.basis) / B;
  return ((1 + dim * rate - (pr / 100 + a * rate)) / (pr / 100 + a * rate)) / dsm;
});

// ----- Treasury bills -------------------------------------------------------------------------

/** Days to maturity of a treasury bill, which may run one year at most. */
function tbill(s0: number, m0: number): number | CellError {
  const settle = Math.trunc(s0), mat = Math.trunc(m0);
  if (settle >= mat || mat > addMonths(settle, 12, false)) return E.num;
  return mat - settle;
}

numFn("TBILLPRICE", 3, 3, ([s, m, disc]) => {
  const dsm = tbill(s, m);
  if (isErr(dsm)) return dsm;
  if (disc <= 0) return E.num;
  const p = 100 * (1 - (disc * dsm) / 360);
  return p <= 0 ? E.num : p;
});
numFn("TBILLYIELD", 3, 3, ([s, m, pr]) => {
  const dsm = tbill(s, m);
  if (isErr(dsm)) return dsm;
  if (pr <= 0) return E.num;
  return ((100 - pr) / pr) * (360 / dsm);
});
numFn("TBILLEQ", 3, 3, ([s, m, disc]) => {
  const dsm = tbill(s, m);
  if (isErr(dsm)) return dsm;
  if (disc <= 0) return E.num;
  if (dsm <= 182) return (365 * disc) / (360 - disc * dsm);
  const p = 100 * (1 - (disc * dsm) / 360);
  const t = dsm / 365;
  return (-2 * t + 2 * Math.sqrt(t * t - (2 * t - 1) * (1 - 100 / p))) / (2 * t - 1);
});

// ----- Odd first and last periods -------------------------------------------------------------

/** Days of [a, b] by the basis, and the normal length of a quasi-coupon period [p0, p1]. */
function quasiLen(p0: number, p1: number, basis: number, freq: number): number {
  return basis === 1 ? p1 - p0 : (basis === 3 ? 365 : 360) / freq;
}

/** The share of the quasi-coupon period [p0, p1] that [a, b] covers: a whole period counts 1. */
function part(a: number, b: number, p0: number, p1: number, nl: number, basis: number): number {
  return a === p0 && b === p1 ? 1 : dayCount(a, b, basis) / nl;
}

interface OddF { settle: number; mat: number; issue: number; first: number; freq: number; basis: number; }

function oddfArgs(x: number[]): OddF | CellError {
  const [s, m, i, fc] = x.map(Math.trunc);
  const freq = Math.trunc(x[7]), basis = Math.trunc(x[8]);
  if (!(m > fc && fc > s && s > i) || ![1, 2, 4].includes(freq) || basis < 0 || basis > 4) return E.num;
  return { settle: s, mat: m, issue: i, first: fc, freq, basis };
}

function oddfPrice(o: OddF, rate: number, yld: number, red: number): number {
  const f = o.freq, step = 12 / f, cpn = (100 * rate) / f, v = 1 + yld / f;
  const eom = isEom(o.mat);
  // Regular coupons after the first one.
  let N = 0;
  while (addMonths(o.mat, -N * step, eom) > o.first) N++;
  // Quasi-coupon periods back from the first coupon date until the issue.
  const feom = isEom(o.first);
  const q = (k: number) => addMonths(o.first, -k * step, feom);
  let nc = 1;
  while (q(nc) > o.issue) nc++;
  let dcSum = 0, aSum = 0, nq = 0, dsc = 0, nlS = 1;
  for (let k = 1; k <= nc; k++) {
    const p0 = q(k), p1 = q(k - 1);
    const nl = quasiLen(p0, p1, o.basis, f);
    const from = Math.max(p0, o.issue);
    dcSum += part(from, p1, p0, p1, nl, o.basis);
    if (o.settle > from) aSum += part(from, Math.min(o.settle, p1), p0, p1, nl, o.basis);
    if (o.settle >= p0 && o.settle < p1) {
      dsc = o.basis === 0 || o.basis === 4 ? nl - dayCount(p0, o.settle, o.basis) : p1 - o.settle;
      nlS = nl;
      nq = k - 1;
    }
  }
  const t0 = nq + dsc / nlS;
  let p = red / Math.pow(v, N + t0) + (cpn * dcSum) / Math.pow(v, t0);
  for (let k = 1; k <= N; k++) p += cpn / Math.pow(v, k + t0);
  return p - cpn * aSum;
}

numFn("ODDFPRICE", 8, 9, (x) => {
  const o = oddfArgs(x);
  if (isErr(o)) return o;
  const [, , , , rate, yld, red] = x;
  if (rate < 0 || yld < 0 || red <= 0) return E.num;
  return oddfPrice(o, rate, yld, red);
});
numFn("ODDFYIELD", 8, 9, (x) => {
  const o = oddfArgs(x);
  if (isErr(o)) return o;
  const [, , , , rate, pr, red] = x;
  if (rate < 0 || pr <= 0 || red <= 0) return E.num;
  const y = solve((t) => oddfPrice(o, rate, t, red) - pr, 0.1, -0.99 * o.freq, 10);
  return y === null ? E.num : y;
});

interface OddL { settle: number; mat: number; last: number; freq: number; basis: number; }

function oddlArgs(x: number[]): OddL | CellError {
  const [s, m, l] = x.map(Math.trunc);
  const freq = Math.trunc(x[6]), basis = Math.trunc(x[7]);
  if (!(m > s && s > l) || ![1, 2, 4].includes(freq) || basis < 0 || basis > 4) return E.num;
  return { settle: s, mat: m, last: l, freq, basis };
}

/** Σ DC/NL, Σ A/NL and Σ DSC/NL over the quasi-coupon periods from the last coupon to maturity. */
function oddlSums(o: OddL) {
  const step = 12 / o.freq, eom = isEom(o.last);
  let dc = 0, a = 0, dsc = 0;
  for (let k = 0; ; k++) {
    const p0 = addMonths(o.last, k * step, eom), p1 = addMonths(o.last, (k + 1) * step, eom);
    if (p0 >= o.mat) break;
    // Here the actual-day bases measure each quasi period by its real length (recorded for basis 3).
    const nl = o.basis === 0 || o.basis === 4 ? 360 / o.freq : p1 - p0;
    const end = Math.min(p1, o.mat);
    dc += part(p0, end, p0, p1, nl, o.basis);
    if (o.settle > p0) a += part(p0, Math.min(o.settle, end), p0, p1, nl, o.basis);
    if (end > o.settle) dsc += part(Math.max(p0, o.settle), end, p0, p1, nl, o.basis);
  }
  return { dc, a, dsc };
}

numFn("ODDLPRICE", 7, 8, (x) => {
  const o = oddlArgs(x);
  if (isErr(o)) return o;
  const [, , , rate, yld, red] = x;
  if (rate < 0 || yld < 0 || red <= 0) return E.num;
  const f = o.freq, cpn = (100 * rate) / f, s = oddlSums(o);
  return (red + s.dc * cpn) / (1 + (s.dsc * yld) / f) - s.a * cpn;
});
numFn("ODDLYIELD", 7, 8, (x) => {
  const o = oddlArgs(x);
  if (isErr(o)) return o;
  const [, , , rate, pr, red] = x;
  if (rate < 0 || pr <= 0 || red <= 0) return E.num;
  const f = o.freq, cpn = (100 * rate) / f, s = oddlSums(o);
  return ((red + s.dc * cpn - (pr + s.a * cpn)) / (pr + s.a * cpn)) * (f / s.dsc);
});

// ----- Depreciation ---------------------------------------------------------------------------

numFn("SLN", 3, 3, ([c, s, l]) => (l === 0 ? E.div0 : (c - s) / l));
numFn("SYD", 4, 4, ([c, s, l, p]) => {
  if (l <= 0 || p <= 0 || p > l || s < 0) return E.num;
  return ((c - s) * (l - p + 1) * 2) / (l * (l + 1));
});

numFn("DB", 4, 5, ([c, s, l, p0, m0]) => {
  const life = Math.trunc(l), per = Math.trunc(p0), month = Math.trunc(m0);
  if (c < 0 || s < 0 || life <= 0 || per <= 0 || month < 1 || month > 12 || per > life + (month < 12 ? 1 : 0)) return E.num;
  if (c === 0) return 0;
  const rate = Math.round((1 - Math.pow(s / c, 1 / life)) * 1000) / 1000;
  let total = 0, dep = 0;
  for (let i = 1; i <= per; i++) {
    if (i === 1) dep = (c * rate * month) / 12;
    else if (i === life + 1) dep = ((c - total) * rate * (12 - month)) / 12;
    else dep = (c - total) * rate;
    total += dep;
  }
  return dep;
}, [0, 0, 0, 0, 12]);

numFn("DDB", 4, 5, ([c, s, l, p, f]) => {
  if (c < 0 || s < 0 || l <= 0 || p <= 0 || f <= 0 || p > l) return E.num;
  const rate = Math.min(1, f / l);
  const prior = Math.max(0, c * (1 - Math.pow(1 - rate, p - 1)));
  return Math.max(0, Math.min((c - prior) * rate, c - s - prior));
}, [0, 0, 0, 0, 2]);

/** Double-declining balance of one whole period, as VDB uses it. */
function ddbPeriod(c: number, s: number, l: number, p: number, f: number): number {
  let rate = f / l, old: number;
  if (rate >= 1) {
    rate = 1;
    old = p === 1 ? c : 0;
  } else old = c * Math.pow(1 - rate, p - 1);
  const nw = c * Math.pow(1 - rate, p);
  const d = nw < s ? old - s : old - nw;
  return d < 0 ? 0 : d;
}

/** Depreciation from 0 to `period` with the switch to straight-line when that is larger. */
function interVdb(c: number, s: number, l: number, l1: number, period: number, f: number): number {
  let total = 0, rest = c - s, lin = 0, nowLin = false;
  const end = Math.ceil(period);
  for (let i = 1; i <= end; i++) {
    let term: number;
    if (!nowLin) {
      const ddb = ddbPeriod(c, s, l, i, f);
      lin = rest / (l1 - (i - 1));
      if (lin > ddb) {
        term = lin;
        nowLin = true;
      } else {
        term = ddb;
        rest -= ddb;
      }
    } else term = lin;
    if (i === end) term *= period + 1 - end;
    total += term;
  }
  return total;
}

def("VDB", 5, 7, (args, ctx) => {
  const vals: number[] = [];
  for (let i = 0; i < 7; i++) {
    const a = args[i];
    if (a === undefined || a === null) {
      vals.push(i === 5 ? 2 : 0);
      continue;
    }
    const d = deref(a, ctx);
    const v = d instanceof Arr ? d.d[0] : d;
    const x = typeof v === "boolean" ? (v ? 1 : 0) : typeof v === "number" ? v : v === null ? 0 : Number.NaN;
    if (isErr(v)) return v;
    if (Number.isNaN(x)) return E.value;
    vals.push(x);
  }
  const [c, s, l, st, en, f, ns] = vals;
  if (c < 0 || s < 0 || l <= 0 || st < 0 || en < st || en > l || f <= 0) return E.num;
  const iStart = Math.floor(st), iEnd = Math.ceil(en);
  let r = 0;
  if (ns) {
    for (let i = iStart + 1; i <= iEnd; i++) {
      let term = ddbPeriod(c, s, l, i, f);
      if (i === iStart + 1) term *= Math.min(en, iStart + 1) - st;
      else if (i === iEnd) term *= en + 1 - iEnd;
      r += term;
    }
    return num(r);
  }
  // A fractional start shifts the periods: they run from the start on, on the book value there.
  const rest = c - interVdb(c, s, l, l, st, f);
  return num(interVdb(rest, s, l, l - st, en - st, f));
});

numFn("AMORLINC", 6, 7, ([cost, d0, p0, salvage, per0, rate, b0]) => {
  const basis = Math.trunc(b0), per = Math.trunc(per0);
  if (basis < 0 || basis > 4 || basis === 2 || cost < 0 || salvage < 0 || salvage > cost || rate <= 0 || per < 0) return E.num;
  const one = cost * rate;
  const first = yearFrac(Math.trunc(d0), Math.trunc(p0), basis) * rate * cost;
  const full = Math.trunc((cost - salvage - first) / one);
  if (per === 0) return first;
  if (per <= full) return one;
  if (per === full + 1) return cost - salvage - one * full - first;
  return 0;
});

// ----- Interest -------------------------------------------------------------------------------

numFn("EFFECT", 2, 2, ([r, p]) => {
  const k = Math.trunc(p);
  if (r <= 0 || k < 1) return E.num;
  return Math.pow(1 + r / k, k) - 1;
});
numFn("NOMINAL", 2, 2, ([r, p]) => {
  const k = Math.trunc(p);
  if (r <= 0 || k < 1) return E.num;
  return k * (Math.pow(1 + r, 1 / k) - 1);
});
numFn("ISPMT", 4, 4, ([r, per, np, pv]) => (np === 0 ? E.div0 : pv * r * (per / np - 1)));
numFn("PDURATION", 3, 3, ([r, pv, fv]) => (r <= 0 || pv <= 0 || fv <= 0 ? E.num : (Math.log(fv) - Math.log(pv)) / Math.log(1 + r)));
numFn("RRI", 3, 3, ([np, pv, fv]) => (np <= 0 || pv === 0 ? E.num : Math.pow(fv / pv, 1 / np) - 1));

numFn("DOLLARDE", 2, 2, ([x, f0]) => {
  const f = Math.trunc(f0);
  if (f < 0) return E.num;
  if (f === 0) return E.div0;
  const ip = Math.trunc(x);
  return ip + ((x - ip) * Math.pow(10, Math.ceil(Math.log10(f)))) / f;
});
numFn("DOLLARFR", 2, 2, ([x, f0]) => {
  const f = Math.trunc(f0);
  if (f < 0) return E.num;
  if (f === 0) return E.div0;
  const ip = Math.trunc(x);
  return ip + ((x - ip) * f) / Math.pow(10, Math.ceil(Math.log10(f)));
});

function pmt(r: number, np: number, pv: number, t: number): number {
  const g = Math.pow(1 + r, np);
  return (-r * pv * g) / ((1 + r * t) * (g - 1));
}
function fvAt(r: number, np: number, p: number, pv: number, t: number): number {
  const g = Math.pow(1 + r, np);
  return -(pv * g + (p * (1 + r * t) * (g - 1)) / r);
}
function ipmt(r: number, per: number, np: number, pv: number, t: number): number {
  const p = pmt(r, np, pv, t);
  if (t && per === 1) return 0;
  const bal = fvAt(r, t ? per - 2 : per - 1, p, pv, t);
  return t ? (bal - p) * r : bal * r;
}

function cum(name: string, principal: boolean): void {
  numFn(name, 6, 6, ([r, np, pv, s0, e0, t0]) => {
    const s = Math.ceil(s0), e = Math.trunc(e0), t = Math.trunc(t0);
    if (r <= 0 || np <= 0 || pv <= 0 || s < 1 || e < s || e > np || (t !== 0 && t !== 1)) return E.num;
    const p = pmt(r, np, pv, t);
    let sum = 0;
    for (let k = s; k <= e; k++) {
      const ip = ipmt(r, k, np, pv, t);
      sum += principal ? p - ip : ip;
    }
    return sum;
  });
}
cum("CUMIPMT", false);
cum("CUMPRINC", true);

def("FVSCHEDULE", 2, 2, ([p0, s0], ctx) => {
  const pd = deref(p0, ctx);
  const pv = pd instanceof Arr ? pd.d[0] : pd;
  if (isErr(pv)) return pv;
  if (typeof pv !== "number" && pv !== null) return E.value;
  let x = pv ?? 0;
  for (const v of values(s0, ctx)) {
    if (isErr(v)) return v;
    if (v === null) continue;
    if (typeof v !== "number") return E.value;
    x *= 1 + v;
  }
  return num(x);
});

// ----- Cash flows -----------------------------------------------------------------------------

/** The numbers of a range or array (text, logical values and empty cells left out). */
function flows(a: unknown, ctx: Parameters<typeof values>[1]): number[] | CellError {
  const out: number[] = [];
  for (const v of values(a as never, ctx)) {
    if (isErr(v)) return v;
    if (typeof v === "number") out.push(v);
  }
  return out;
}

function optNum(a: unknown, ctx: Parameters<typeof values>[1], d: number): number | CellError {
  if (a === undefined || a === null) return d;
  const v = deref(a as never, ctx);
  const s = (v instanceof Arr ? v.d[0] : v) as V;
  if (s === null) return d;
  if (isErr(s)) return s;
  return typeof s === "number" ? s : typeof s === "boolean" ? (s ? 1 : 0) : E.value;
}

def("IRR", 1, 2, ([a, g0], ctx) => {
  const xs = flows(a, ctx);
  if (isErr(xs)) return xs;
  const g = optNum(g0, ctx, 0.1);
  if (isErr(g)) return g;
  if (!xs.some((x) => x > 0) || !xs.some((x) => x < 0)) return E.num;
  const npv = (r: number) => xs.reduce((s, x, i) => s + x / Math.pow(1 + r, i), 0);
  const r = solve(npv, g, -0.999999999, 1e3);
  return r === null ? E.num : num(r);
});

def("MIRR", 3, 3, ([a, f0, r0], ctx) => {
  const xs = flows(a, ctx);
  if (isErr(xs)) return xs;
  const fr = optNum(f0, ctx, 0), rr = optNum(r0, ctx, 0);
  if (isErr(fr)) return fr;
  if (isErr(rr)) return rr;
  const nn = xs.length;
  let pos = 0, neg = 0;
  xs.forEach((x, i) => {
    if (x > 0) pos += x * Math.pow(1 + rr, nn - 1 - i);
    else neg += x / Math.pow(1 + fr, i);
  });
  if (pos === 0 || neg === 0 || nn < 2) return E.div0;
  return num(Math.pow(pos / -neg, 1 / (nn - 1)) - 1);
});

function datedFlows(v0: unknown, d0: unknown, ctx: Parameters<typeof values>[1]): { xs: number[]; ds: number[] } | CellError {
  const vs = values(v0 as never, ctx), dsv = values(d0 as never, ctx);
  if (vs.length !== dsv.length) return E.num;
  const xs: number[] = [], ds: number[] = [];
  for (let i = 0; i < vs.length; i++) {
    const x = vs[i], d = dsv[i];
    if (isErr(x)) return x;
    if (isErr(d)) return d;
    if (typeof x !== "number" || typeof d !== "number") return E.value;
    xs.push(x);
    ds.push(Math.trunc(d));
  }
  if (!xs.length || ds.some((d) => d < ds[0])) return E.num;
  return { xs, ds };
}

def("XNPV", 3, 3, ([r0, v0, d0], ctx) => {
  const r = optNum(r0, ctx, 0);
  if (isErr(r)) return r;
  const f = datedFlows(v0, d0, ctx);
  if (isErr(f)) return f;
  if (r <= 0) return E.num;
  return num(f.xs.reduce((s, x, i) => s + x / Math.pow(1 + r, (f.ds[i] - f.ds[0]) / 365), 0));
});

def("XIRR", 2, 3, ([v0, d0, g0], ctx) => {
  const f = datedFlows(v0, d0, ctx);
  if (isErr(f)) return f;
  const g = optNum(g0, ctx, 0.1);
  if (isErr(g)) return g;
  if (!f.xs.some((x) => x > 0) || !f.xs.some((x) => x < 0)) return E.num;
  const npv = (r: number) => f.xs.reduce((s, x, i) => s + x / Math.pow(1 + r, (f.ds[i] - f.ds[0]) / 365), 0);
  const r = solve(npv, g, -0.999999999, 1e3);
  return r === null ? E.num : num(r);
});

