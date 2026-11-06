// Engineering functions (number systems, bits, Bessel, complex numbers, UMWANDELN) and the
// further maths functions that work on single values (ARABISCH, RÖMISCH, BASIS, DEZIMAL,
// the reciprocal trigonometric functions, POLYNOMIAL, POTENZREIHE, SUMMEX2MY2 …).

import { Ctx, def, deref } from "./eval";
import { defS, n, numbers, values } from "./fn-util";
import { fact } from "./fn-math";
import { erf, erfc } from "./fn-stats";
import { numFn } from "./fn-stats2";
import { Arr, E, Ref, isErr, num, numberText, textToNumber, toStr, type Any, type V } from "./values";
import type { CellError } from "../model";

// ----- Maths ----------------------------------------------------------------------------------

defS("ARABIC", 1, 1, ([a]) => {
  if (typeof a !== "string") return isErr(a) ? a : a === null ? 0 : E.value;
  let t = a.trim().toUpperCase();
  if (t.length > 255) return E.value;
  let sign = 1;
  if (t.startsWith("-")) {
    sign = -1;
    t = t.slice(1);
  }
  const val: Record<string, number> = { I: 1, V: 5, X: 10, L: 50, C: 100, D: 500, M: 1000 };
  let total = 0, prev = 0;
  for (let i = t.length - 1; i >= 0; i--) {
    const v = val[t[i]];
    if (!v) return E.value;
    if (v < prev) total -= v;
    else {
      total += v;
      prev = v;
    }
  }
  return sign * total;
});

const ROMAN_CHARS = "MDCLXVI", ROMAN_VALUES = [1000, 500, 100, 50, 10, 5, 1];

/** Roman numerals in Excel's five forms (0 classic … 4 simplified). */
function roman(value: number, mode: number): string {
  let out = "";
  let v = value;
  for (let i = 0; i <= 3; i++) {
    let idx = 2 * i;
    const digit = Math.floor(v / ROMAN_VALUES[idx]);
    if (digit % 5 === 4) {
      const idx2 = digit === 4 ? idx - 1 : idx - 2;
      let steps = 0;
      while (steps < mode && idx < 6) {
        steps++;
        if (ROMAN_VALUES[idx2] - ROMAN_VALUES[idx + 1] <= v) idx++;
        else steps = mode;
      }
      out += ROMAN_CHARS[idx] + ROMAN_CHARS[idx2];
      v = v + ROMAN_VALUES[idx] - ROMAN_VALUES[idx2];
    } else {
      if (digit > 4) out += ROMAN_CHARS[idx - 1];
      out += ROMAN_CHARS[idx].repeat(digit % 5);
      v %= ROMAN_VALUES[idx];
    }
  }
  return out;
}

defS("ROMAN", 1, 2, ([a, f0]) => {
  const x = n(a);
  if (isErr(x)) return x;
  let mode = 0;
  if (typeof f0 === "boolean") mode = f0 ? 0 : 4;
  else if (f0 !== undefined && f0 !== null) {
    const m = n(f0);
    if (isErr(m)) return m;
    mode = Math.trunc(m);
  }
  const v = Math.trunc(x);
  if (mode < 0 || mode > 4 || v < 0 || v > 3999) return E.value;
  return roman(v, mode);
});

defS("BASE", 2, 3, ([a, r0, l0]) => {
  const x = n(a), r = n(r0);
  if (isErr(x)) return x;
  if (isErr(r)) return r;
  const radix = Math.trunc(r), v = Math.trunc(x);
  if (v < 0 || x >= 2 ** 53 || radix < 2 || radix > 36) return E.num;
  let len = 0;
  if (l0 !== undefined && l0 !== null) {
    const l = n(l0);
    if (isErr(l)) return l;
    len = Math.trunc(l);
    if (len < 0 || len > 255) return E.num;
  }
  return v.toString(radix).toUpperCase().padStart(len, "0");
});

defS("DECIMAL", 2, 2, ([a, r0]) => {
  const t = toStr(a);
  if (isErr(t)) return t;
  const r = n(r0);
  if (isErr(r)) return r;
  const radix = Math.trunc(r);
  if (radix < 2 || radix > 36 || t.length > 255) return E.num;
  let v = 0;
  for (const ch of t.toUpperCase()) {
    const d = parseInt(ch, 36);
    if (Number.isNaN(d) || d >= radix) return E.num;
    v = v * radix + d;
  }
  return num(v);
});

const TRIG_LIMIT = 2 ** 27;
numFn("ACOT", 1, 1, ([x]) => Math.PI / 2 - Math.atan(x));
numFn("ACOTH", 1, 1, ([x]) => (Math.abs(x) <= 1 ? E.num : 0.5 * Math.log((x + 1) / (x - 1))));
numFn("CSC", 1, 1, ([x]) => (Math.abs(x) >= TRIG_LIMIT ? E.num : x === 0 ? E.div0 : 1 / Math.sin(x)));
numFn("SEC", 1, 1, ([x]) => (Math.abs(x) >= TRIG_LIMIT ? E.num : 1 / Math.cos(x)));
numFn("CSCH", 1, 1, ([x]) => (Math.abs(x) >= TRIG_LIMIT ? E.num : x === 0 ? E.div0 : 1 / Math.sinh(x)));
numFn("SECH", 1, 1, ([x]) => (Math.abs(x) >= TRIG_LIMIT ? E.num : 1 / Math.cosh(x)));
numFn("COTH", 1, 1, ([x]) => (Math.abs(x) >= TRIG_LIMIT ? E.num : x === 0 ? E.div0 : 1 / Math.tanh(x)));

def("MULTINOMIAL", 1, 255, (args, ctx) => {
  const xs = numbers(args, ctx);
  if (isErr(xs)) return xs;
  let total = 0, den = 1;
  for (const x0 of xs) {
    const x = Math.trunc(x0);
    if (x < 0) return E.num;
    total += x;
    den *= fact(x);
  }
  return num(fact(total) / den);
});

def("SERIESSUM", 4, 4, ([x0, n0, m0, c0], ctx) => {
  const sc = (a: Any) => {
    const d = deref(a, ctx);
    return n(d instanceof Arr ? d.d[0] : d);
  };
  const x = sc(x0), nn = sc(n0), m = sc(m0);
  for (const v of [x, nn, m]) if (isErr(v)) return v;
  let s = 0, i = 0;
  for (const c of values(c0, ctx)) {
    if (isErr(c)) return c;
    if (typeof c !== "number") return E.value;
    s += c * Math.pow(x as number, (nn as number) + i * (m as number));
    i++;
  }
  return num(s);
});

/** Pairs of numbers at the same places of two equally sized arguments (others skipped). */
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

function sumPairs(name: string, f: (x: number, y: number) => number): void {
  def(name, 2, 2, ([a, b], ctx) => {
    const p = pairs(a, b, ctx);
    if (isErr(p)) return p;
    let s = 0;
    for (let i = 0; i < p[0].length; i++) s += f(p[0][i], p[1][i]);
    return num(s);
  });
}
sumPairs("SUMX2MY2", (x, y) => x * x - y * y);
sumPairs("SUMX2PY2", (x, y) => x * x + y * y);
sumPairs("SUMXMY2", (x, y) => (x - y) * (x - y));

// ----- Number systems -------------------------------------------------------------------------

/** A number written in base 2, 8 or 16 (up to 10 digits, the 10th the sign) as a value. */
function fromBase(v: V, base: number): number | CellError {
  if (isErr(v)) return v;
  if (typeof v === "boolean") return E.value;
  const t = v === null ? "" : typeof v === "number" ? String(v) : v;
  if (t.length > 10) return E.num;
  let x = 0;
  for (const ch of t.toUpperCase()) {
    const d = parseInt(ch, 16);
    if (Number.isNaN(d) || d >= base) return E.num;
    x = x * base + d;
  }
  if (t.length === 10 && x >= base ** 10 / 2) x -= base ** 10;
  return x;
}

/** A value in base 2, 8 or 16, negative ones as 10-digit two's complement. */
function toBase(x: number, base: number, places: V | undefined): V {
  const limit = base ** 10 / 2;
  x = Math.trunc(x);
  if (x < -limit || x >= limit) return E.num;
  if (x < 0) return (x + base ** 10).toString(base).toUpperCase();
  let s = x.toString(base).toUpperCase();
  if (places !== undefined && places !== null) {
    const p = n(places);
    if (isErr(p)) return p;
    const k = Math.trunc(p);
    if (k < s.length || k > 10 || k <= 0) return E.num;
    s = s.padStart(k, "0");
  }
  return s;
}

const BASES: [string, number][] = [["BIN", 2], ["OCT", 8], ["HEX", 16]];
for (const [from, fb] of BASES) {
  for (const [to, tb] of BASES) {
    if (from === to) continue;
    defS(`${from}2${to}`, 1, 2, ([a, p]) => {
      const x = fromBase(a, fb);
      return isErr(x) ? x : toBase(x, tb, p);
    });
  }
  if (from === "OCT") defS("OCT2DEC", 1, 1, ([a]) => fromBase(a, 8));
}
defS("DEC2OCT", 1, 2, ([a, p]) => {
  const x = n(a);
  return isErr(x) ? x : toBase(x, 8, p);
});

// ----- Bits -----------------------------------------------------------------------------------

const BIT_MAX = 2 ** 48;
function bitArg(v: V): bigint | CellError {
  const x = n(v);
  if (isErr(x)) return x;
  if (x < 0 || x >= BIT_MAX || !Number.isInteger(x)) return E.num;
  return BigInt(x);
}
function bitFn(name: string, f: (a: bigint, b: bigint) => bigint): void {
  defS(name, 2, 2, ([a0, b0]) => {
    const a = bitArg(a0), b = bitArg(b0);
    if (isErr(a)) return a;
    if (isErr(b)) return b;
    return Number(f(a, b));
  });
}
bitFn("BITAND", (a, b) => a & b);
bitFn("BITOR", (a, b) => a | b);
bitFn("BITXOR", (a, b) => a ^ b);
function shift(name: string, dir: number): void {
  defS(name, 2, 2, ([a0, s0]) => {
    const a = bitArg(a0);
    if (isErr(a)) return a;
    const s = n(s0);
    if (isErr(s)) return s;
    const k = Math.trunc(s) * dir;
    if (Math.abs(k) > 53) return E.num;
    const r = k >= 0 ? a << BigInt(k) : a >> BigInt(-k);
    if (r >= BigInt(BIT_MAX)) return E.num;
    return Number(r);
  });
}
shift("BITLSHIFT", 1);
shift("BITRSHIFT", -1);

numFn("DELTA", 1, 2, ([a, b]) => (a === b ? 1 : 0));
numFn("GESTEP", 1, 2, ([a, b]) => (a >= b ? 1 : 0));
defS("ERF", 1, 2, ([a, b]) => {
  const lo = n(a);
  if (isErr(lo)) return lo;
  if (b === undefined || b === null) return erf(lo);
  const hi = n(b);
  if (isErr(hi)) return hi;
  return num(erf(hi) - erf(lo));
});
numFn("ERF.PRECISE", 1, 1, ([x]) => erf(x));
numFn("ERFC", 1, 1, ([x]) => erfc(x));
numFn("ERFC.PRECISE", 1, 1, ([x]) => erfc(x));

// ----- Bessel functions (the classic polynomial approximations Excel's results follow) --------

function j0(x: number): number {
  const ax = Math.abs(x);
  if (ax < 8) {
    const y = x * x;
    const a1 = 57568490574.0 + y * (-13362590354.0 + y * (651619640.7 + y * (-11214424.18 + y * (77392.33017 + y * -184.9052456))));
    const a2 = 57568490411.0 + y * (1029532985.0 + y * (9494680.718 + y * (59272.64853 + y * (267.8532712 + y))));
    return a1 / a2;
  }
  const z = 8 / ax, y = z * z, xx = ax - 0.785398164;
  const a1 = 1 + y * (-0.1098628627e-2 + y * (0.2734510407e-4 + y * (-0.2073370639e-5 + y * 0.2093887211e-6)));
  const a2 = -0.1562499995e-1 + y * (0.1430488765e-3 + y * (-0.6911147651e-5 + y * (0.7621095161e-6 - y * 0.934935152e-7)));
  return Math.sqrt(0.636619772 / ax) * (Math.cos(xx) * a1 - z * Math.sin(xx) * a2);
}

function j1(x: number): number {
  const ax = Math.abs(x);
  if (ax < 8) {
    const y = x * x;
    const a1 = x * (72362614232.0 + y * (-7895059235.0 + y * (242396853.1 + y * (-2972611.439 + y * (15704.4826 + y * -30.16036606)))));
    const a2 = 144725228442.0 + y * (2300535178.0 + y * (18583304.74 + y * (99447.43394 + y * (376.9991397 + y))));
    return a1 / a2;
  }
  const z = 8 / ax, y = z * z, xx = ax - 2.356194491;
  const a1 = 1 + y * (0.183105e-2 + y * (-0.3516396496e-4 + y * (0.2457520174e-5 + y * -0.240337019e-6)));
  const a2 = 0.04687499995 + y * (-0.2002690873e-3 + y * (0.8449199096e-5 + y * (-0.88228987e-6 + y * 0.105787412e-6)));
  const r = Math.sqrt(0.636619772 / ax) * (Math.cos(xx) * a1 - z * Math.sin(xx) * a2);
  return x < 0 ? -r : r;
}

function jn(nn: number, x: number): number {
  if (nn === 0) return j0(x);
  if (nn === 1) return j1(x);
  const ax = Math.abs(x);
  if (ax === 0) return 0;
  let ans: number;
  const tox = 2 / ax;
  if (ax > nn) {
    let bjm = j0(ax), bj = j1(ax);
    for (let j = 1; j < nn; j++) {
      const bjp = j * tox * bj - bjm;
      bjm = bj;
      bj = bjp;
    }
    ans = bj;
  } else {
    const m = 2 * Math.floor((nn + Math.floor(Math.sqrt(40 * nn))) / 2);
    let jsum = false, bjp = 0, sumv = 0, bj = 1;
    ans = 0;
    for (let j = m; j > 0; j--) {
      const bjm = j * tox * bj - bjp;
      bjp = bj;
      bj = bjm;
      if (Math.abs(bj) > 1e10) {
        bj *= 1e-10;
        bjp *= 1e-10;
        ans *= 1e-10;
        sumv *= 1e-10;
      }
      if (jsum) sumv += bj;
      jsum = !jsum;
      if (j === nn) ans = bjp;
    }
    sumv = 2 * sumv - bj;
    ans /= sumv;
  }
  return x < 0 && nn % 2 ? -ans : ans;
}

function y0(x: number): number {
  if (x < 8) {
    const y = x * x;
    const a1 = -2957821389.0 + y * (7062834065.0 + y * (-512359803.6 + y * (10879881.29 + y * (-86327.92757 + y * 228.4622733))));
    const a2 = 40076544269.0 + y * (745249964.8 + y * (7189466.438 + y * (47447.2647 + y * (226.1030244 + y))));
    return a1 / a2 + 0.636619772 * j0(x) * Math.log(x);
  }
  const z = 8 / x, y = z * z, xx = x - 0.785398164;
  const a1 = 1 + y * (-0.1098628627e-2 + y * (0.2734510407e-4 + y * (-0.2073370639e-5 + y * 0.2093887211e-6)));
  const a2 = -0.1562499995e-1 + y * (0.1430488765e-3 + y * (-0.6911147651e-5 + y * (0.7621095161e-6 + y * -0.934945152e-7)));
  return Math.sqrt(0.636619772 / x) * (Math.sin(xx) * a1 + z * Math.cos(xx) * a2);
}

function y1(x: number): number {
  if (x < 8) {
    const y = x * x;
    const a1 = x * (-0.4900604943e13 + y * (0.127527439e13 + y * (-0.5153438139e11 + y * (0.7349264551e9 + y * (-0.4237922726e7 + y * 0.8511937935e4)))));
    const a2 = 0.249958057e14 + y * (0.4244419664e12 + y * (0.3733650367e10 + y * (0.2245904002e8 + y * (0.102042605e6 + y * (0.3549632885e3 + y)))));
    return a1 / a2 + 0.636619772 * (j1(x) * Math.log(x) - 1 / x);
  }
  const z = 8 / x, y = z * z, xx = x - 2.356194491;
  const a1 = 1 + y * (0.183105e-2 + y * (-0.3516396496e-4 + y * (0.2457520174e-5 + y * -0.240337019e-6)));
  const a2 = 0.04687499995 + y * (-0.2002690873e-3 + y * (0.8449199096e-5 + y * (-0.88228987e-6 + y * 0.105787412e-6)));
  return Math.sqrt(0.636619772 / x) * (Math.sin(xx) * a1 + z * Math.cos(xx) * a2);
}

function yn(nn: number, x: number): number {
  if (nn === 0) return y0(x);
  let bym = y0(x), by = y1(x);
  const tox = 2 / x;
  for (let j = 1; j < nn; j++) {
    const byp = j * tox * by - bym;
    bym = by;
    by = byp;
  }
  return by;
}

function i0(x: number): number {
  const ax = Math.abs(x);
  if (ax < 3.75) {
    const y = (x / 3.75) ** 2;
    return 1 + y * (3.5156229 + y * (3.0899424 + y * (1.2067492 + y * (0.2659732 + y * (0.360768e-1 + y * 0.45813e-2)))));
  }
  const y = 3.75 / ax;
  return (Math.exp(ax) / Math.sqrt(ax)) * (0.39894228 + y * (0.1328592e-1 + y * (0.225319e-2 + y * (-0.157565e-2 + y * (0.916281e-2 + y * (-0.2057706e-1 + y * (0.2635537e-1 + y * (-0.1647633e-1 + y * 0.392377e-2))))))));
}

function i1(x: number): number {
  const ax = Math.abs(x);
  let r: number;
  if (ax < 3.75) {
    const y = (x / 3.75) ** 2;
    r = ax * (0.5 + y * (0.87890594 + y * (0.51498869 + y * (0.15084934 + y * (0.2658733e-1 + y * (0.301532e-2 + y * 0.32411e-3))))));
  } else {
    const y = 3.75 / ax;
    let a = 0.2282967e-1 + y * (-0.2895312e-1 + y * (0.1787654e-1 - y * 0.420059e-2));
    a = 0.39894228 + y * (-0.3988024e-1 + y * (-0.362018e-2 + y * (0.163801e-2 + y * (-0.1031555e-1 + y * a))));
    r = a * (Math.exp(ax) / Math.sqrt(ax));
  }
  return x < 0 ? -r : r;
}

function iN(nn: number, x: number): number {
  if (nn === 0) return i0(x);
  if (nn === 1) return i1(x);
  if (x === 0) return 0;
  const tox = 2 / Math.abs(x);
  let bip = 0, ans = 0, bi = 1;
  for (let j = 2 * (nn + Math.floor(Math.sqrt(40 * nn))); j > 0; j--) {
    const bim = bip + j * tox * bi;
    bip = bi;
    bi = bim;
    if (Math.abs(bi) > 1e10) {
      ans *= 1e-10;
      bi *= 1e-10;
      bip *= 1e-10;
    }
    if (j === nn) ans = bip;
  }
  ans *= i0(x) / bi;
  return x < 0 && nn % 2 ? -ans : ans;
}

function k0(x: number): number {
  if (x <= 2) {
    const y = (x * x) / 4;
    return -Math.log(x / 2) * i0(x) + (-0.57721566 + y * (0.4227842 + y * (0.23069756 + y * (0.348859e-1 + y * (0.262698e-2 + y * (0.1075e-3 + y * 0.74e-5))))));
  }
  const y = 2 / x;
  return (Math.exp(-x) / Math.sqrt(x)) * (1.25331414 + y * (-0.7832358e-1 + y * (0.2189568e-1 + y * (-0.1062446e-1 + y * (0.587872e-2 + y * (-0.25154e-2 + y * 0.53208e-3))))));
}

function k1(x: number): number {
  if (x <= 2) {
    const y = (x * x) / 4;
    return Math.log(x / 2) * i1(x) + (1 / x) * (1 + y * (0.15443144 + y * (-0.67278579 + y * (-0.18156897 + y * (-0.1919402e-1 + y * (-0.110404e-2 + y * -0.4686e-4))))));
  }
  const y = 2 / x;
  return (Math.exp(-x) / Math.sqrt(x)) * (1.25331414 + y * (0.23498619 + y * (-0.3655620e-1 + y * (0.1504268e-1 + y * (-0.780353e-2 + y * (0.325614e-2 + y * -0.68245e-3))))));
}

function kN(nn: number, x: number): number {
  if (nn === 0) return k0(x);
  let bkm = k0(x), bk = k1(x);
  const tox = 2 / x;
  for (let j = 1; j < nn; j++) {
    const bkp = bkm + j * tox * bk;
    bkm = bk;
    bk = bkp;
  }
  return bk;
}

function bessel(name: string, f: (nn: number, x: number) => number, positive: boolean): void {
  numFn(name, 2, 2, ([x, n0]) => {
    const nn = Math.trunc(n0);
    if (nn < 0 || (positive && x <= 0)) return E.num;
    return f(nn, x);
  });
}
bessel("BESSELJ", jn, false);
bessel("BESSELI", iN, false);
bessel("BESSELY", yn, true);
bessel("BESSELK", kN, true);

// ----- Complex numbers ------------------------------------------------------------------------

interface Cx { re: number; im: number; suffix: string; }

/** A number part of a complex text, read as German Excel reads numbers. */
function cxNum(t: string): number | null {
  if (t === "" || /\s/.test(t)) return null;
  return textToNumber(t);
}

/** Reads "3+4i", "-2,5j", "i", 5 … ; null when it is no complex number. */
function parseCx(v: V): Cx | CellError {
  if (isErr(v)) return v;
  if (typeof v === "number") return { re: v, im: 0, suffix: "" };
  if (typeof v === "boolean") return E.value;
  const t = v === null ? "0" : v;
  if (t === "") return E.num;
  const last = t[t.length - 1];
  if (last !== "i" && last !== "j") {
    const x = cxNum(t);
    return x === null ? E.num : { re: x, im: 0, suffix: "" };
  }
  const body = t.slice(0, -1);
  let k = -1;
  for (let i = body.length - 1; i > 0; i--) {
    if ((body[i] === "+" || body[i] === "-") && !/[eE]/.test(body[i - 1])) {
      k = i;
      break;
    }
  }
  const reText = k < 0 ? "" : body.slice(0, k), imText = k < 0 ? body : body.slice(k);
  let re = 0;
  if (reText !== "") {
    const x = cxNum(reText);
    if (x === null) return E.num;
    re = x;
  }
  let im: number;
  if (imText === "" || imText === "+") im = 1;
  else if (imText === "-") im = -1;
  else {
    const sign = imText[0] === "-" ? -1 : 1;
    const mag = imText[0] === "+" || imText[0] === "-" ? imText.slice(1) : imText;
    if (mag === "" || mag[0] === "+" || mag[0] === "-") return E.num;
    const x = cxNum(mag);
    if (x === null) return E.num;
    im = sign * x;
  }
  return { re, im, suffix: last };
}

/** A number as a part of a complex text: 15 digits, written out up to 20 digits before the point. */
function cxText(x: number): string {
  if (x !== 0 && Math.abs(x) >= 1e20 && Math.abs(x) < 1e21) {
    return (x < 0 ? "-" : "") + BigInt(Math.round(Math.abs(Number(Math.abs(x).toPrecision(15))))).toString();
  }
  return numberText(x);
}

function formatCx(re: number, im: number, suffix: string): string | CellError {
  if (!Number.isFinite(re) || !Number.isFinite(im)) return E.num;
  re = Number(re.toPrecision(15));
  im = Number(im.toPrecision(15));
  if (re === 0) re = 0;
  if (im === 0) im = 0;
  const s = suffix || "i";
  if (im === 0) return cxText(re);
  const imPart = im === 1 ? s : im === -1 ? "-" + s : cxText(im) + s;
  if (re === 0) return imPart;
  return cxText(re) + (im > 0 ? "+" : "") + imPart;
}

defS("COMPLEX", 2, 3, ([a, b, s0]) => {
  const re = n(a), im = n(b);
  if (isErr(re)) return re;
  if (isErr(im)) return im;
  let s = "i";
  if (s0 !== undefined && s0 !== null) {
    if (isErr(s0)) return s0;
    if (s0 !== "i" && s0 !== "j" && s0 !== "") return E.value;
    s = s0 || "i";
  }
  return formatCx(re, im, s);
});

function cx1(name: string, f: (z: Cx) => V | Cx): void {
  defS(name, 1, 1, ([a]) => {
    const z = parseCx(a);
    if (isErr(z)) return z;
    const r = f(z);
    if (r !== null && typeof r === "object" && "re" in r) return formatCx(r.re, r.im, r.suffix || z.suffix);
    return typeof r === "number" ? num(r) : r;
  });
}

const cx = (re: number, im: number, suffix = ""): Cx => ({ re, im, suffix });
const mul = (a: Cx, b: Cx): Cx => cx(a.re * b.re - a.im * b.im, a.re * b.im + a.im * b.re);
function div(a: Cx, b: Cx): Cx | CellError {
  const d = b.re * b.re + b.im * b.im;
  if (d === 0) return E.num;
  return cx((a.re * b.re + a.im * b.im) / d, (a.im * b.re - a.re * b.im) / d);
}
const cxSin = (z: Cx) => cx(Math.sin(z.re) * Math.cosh(z.im), Math.cos(z.re) * Math.sinh(z.im));
const cxCos = (z: Cx) => cx(Math.cos(z.re) * Math.cosh(z.im), -Math.sin(z.re) * Math.sinh(z.im));
const cxSinh = (z: Cx) => cx(Math.sinh(z.re) * Math.cos(z.im), Math.cosh(z.re) * Math.sin(z.im));
const cxCosh = (z: Cx) => cx(Math.cosh(z.re) * Math.cos(z.im), Math.sinh(z.re) * Math.sin(z.im));
const recip = (z: Cx) => div(cx(1, 0), z);
function lnCx(z: Cx): Cx | CellError {
  if (z.re === 0 && z.im === 0) return E.num;
  return cx(Math.log(Math.hypot(z.re, z.im)), Math.atan2(z.im, z.re));
}

cx1("IMABS", (z) => Math.hypot(z.re, z.im));
cx1("IMAGINARY", (z) => z.im);
cx1("IMREAL", (z) => z.re);
cx1("IMARGUMENT", (z) => (z.re === 0 && z.im === 0 ? E.div0 : Math.atan2(z.im, z.re)));
cx1("IMCONJUGATE", (z) => cx(z.re, -z.im));
cx1("IMCOS", cxCos);
cx1("IMSIN", cxSin);
cx1("IMCOSH", cxCosh);
cx1("IMSINH", cxSinh);
// Tangent, cotangent and secant through the double angle, which gives Excel's last digits.
cx1("IMTAN", (z) => {
  const d = Math.cos(2 * z.re) + Math.cosh(2 * z.im);
  return d === 0 ? E.num : cx(Math.sin(2 * z.re) / d, Math.sinh(2 * z.im) / d);
});
cx1("IMCOT", (z) => {
  const d = Math.cosh(2 * z.im) - Math.cos(2 * z.re);
  return d === 0 ? E.num : cx(Math.sin(2 * z.re) / d, -Math.sinh(2 * z.im) / d);
});
cx1("IMSEC", (z) => {
  const d = Math.cos(2 * z.re) + Math.cosh(2 * z.im);
  return d === 0 ? E.num : cx((2 * Math.cos(z.re) * Math.cosh(z.im)) / d, (2 * Math.sin(z.re) * Math.sinh(z.im)) / d);
});
cx1("IMCSC", (z) => (z.re === 0 && z.im === 0 ? E.num : recip(cxSin(z))));
cx1("IMSECH", (z) => recip(cxCosh(z)));
cx1("IMCSCH", (z) => (z.re === 0 && z.im === 0 ? E.num : recip(cxSinh(z))));
cx1("IMEXP", (z) => cx(Math.exp(z.re) * Math.cos(z.im), Math.exp(z.re) * Math.sin(z.im)));
cx1("IMLN", lnCx);
cx1("IMLOG10", (z) => {
  const l = lnCx(z);
  return isErr(l) ? l : cx(l.re / Math.LN10, l.im / Math.LN10);
});
cx1("IMLOG2", (z) => {
  const l = lnCx(z);
  return isErr(l) ? l : cx(l.re / Math.LN2, l.im / Math.LN2);
});
function polarPow(z: Cx, p: number): Cx | CellError {
  const r = Math.hypot(z.re, z.im);
  if (r === 0) return p > 0 ? cx(0, 0) : E.num;
  const t = Math.atan2(z.im, z.re);
  const rp = Math.pow(r, p);
  return cx(rp * Math.cos(p * t), rp * Math.sin(p * t));
}
cx1("IMSQRT", (z) => polarPow(z, 0.5));
defS("IMPOWER", 2, 2, ([a, p0]) => {
  const z = parseCx(a);
  if (isErr(z)) return z;
  const p = n(p0);
  if (isErr(p)) return p;
  const r = polarPow(z, p);
  return isErr(r) ? r : formatCx(r.re, r.im, z.suffix);
});

/** The common suffix of complex arguments ("" when none has one); null when they mix i and j. */
function suffixOf(zs: Cx[]): string | null {
  let s = "";
  for (const z of zs) {
    if (!z.suffix) continue;
    if (s && z.suffix !== s) return null;
    s = z.suffix;
  }
  return s;
}

defS("IMDIV", 2, 2, ([a, b]) => {
  const x = parseCx(a), y = parseCx(b);
  if (isErr(x)) return x;
  if (isErr(y)) return y;
  const s = suffixOf([x, y]);
  if (s === null) return E.num;
  const r = div(x, y);
  return isErr(r) ? r : formatCx(r.re, r.im, s);
});
defS("IMSUB", 2, 2, ([a, b]) => {
  const x = parseCx(a), y = parseCx(b);
  if (isErr(x)) return x;
  if (isErr(y)) return y;
  const s = suffixOf([x, y]);
  if (s === null) return E.value;
  return formatCx(x.re - y.re, x.im - y.im, s);
});

/** The complex numbers of all arguments (ranges and arrays element by element, empty cells skipped). */
function cxArgs(args: Any[], ctx: Ctx): Cx[] | CellError {
  const out: Cx[] = [];
  for (const a of args) {
    const vals = a instanceof Ref || a instanceof Arr ? values(a, ctx).filter((v) => v !== null) : [deref(a, ctx) as V];
    for (const v of vals) {
      const z = parseCx(v);
      if (isErr(z)) return z;
      out.push(z);
    }
  }
  return out;
}
def("IMSUM", 1, 255, (args, ctx) => {
  const zs = cxArgs(args, ctx);
  if (isErr(zs)) return zs;
  const s = suffixOf(zs);
  if (s === null) return E.value;
  return formatCx(zs.reduce((t, z) => t + z.re, 0), zs.reduce((t, z) => t + z.im, 0), s);
});
def("IMPRODUCT", 1, 255, (args, ctx) => {
  const zs = cxArgs(args, ctx);
  if (isErr(zs)) return zs;
  const s = suffixOf(zs);
  if (s === null) return E.value;
  let p = cx(1, 0);
  for (const z of zs) p = mul(p, z);
  return formatCx(p.re, p.im, s);
});

// ----- UMWANDELN (CONVERT) --------------------------------------------------------------------

/** Unit → [quantity, factor to the quantity's base unit, takes prefixes]. */
const UNITS: Record<string, [string, number, boolean]> = {};
function units(q: string, list: [string[], number, boolean?][]): void {
  for (const [names, f, p] of list) for (const nm of names) UNITS[nm] = [q, f, p ?? false];
}
const IN = 0.0254, FT = 0.3048, MI = 1609.344, NMI = 1852, LY = 9460730472580800, PICA_PT = IN / 72;
const LBF = 4.4482216152605, HP = 550 * FT * LBF;
units("mass", [[["g"], 1, true], [["sg"], (LBF / FT) * 1000], [["lbm"], 453.59237], [["u"], 1.660538782e-24, true], [["ozm"], 28.349523125],
  [["grain"], 0.06479891], [["cwt", "shweight"], 45359.237], [["uk_cwt", "lcwt", "hweight"], 50802.34544], [["stone"], 6350.29318],
  [["ton"], 907184.74], [["uk_ton", "LTON", "brton"], 1016046.9088]]);
units("length", [[["m"], 1, true], [["mi"], MI], [["Nmi"], NMI], [["in"], IN], [["ft"], FT], [["yd"], 0.9144], [["ang"], 1e-10, true],
  [["ell"], 1.143], [["ly"], LY, true], [["parsec", "pc"], 30856775814671900, true], [["Picapt", "Pica"], PICA_PT], [["pica"], IN / 6],
  [["survey_mi"], (5280 * 1200) / 3937]]);
units("time", [[["yr"], 31557600, true], [["day", "d"], 86400, true], [["hr"], 3600, true], [["mn", "min"], 60, true], [["sec", "s"], 1, true]]);
units("pressure", [[["Pa", "p"], 1, true], [["atm", "at"], 101325, true], [["mmHg"], 133.322, true], [["psi"], LBF / (IN * IN)], [["Torr"], 101325 / 760]]);
units("force", [[["N"], 1, true], [["dyn", "dy"], 1e-5, true], [["lbf"], LBF], [["pond"], 0.00980665, true]]);
units("energy", [[["J"], 1, true], [["e"], 1e-7, true], [["c"], 4.184, true], [["cal"], 4.1868, true], [["eV", "ev"], 1.602176487e-19, true],
  [["HPh", "hh"], HP * 3600], [["Wh", "wh"], 3600, true], [["flb"], FT * LBF], [["BTU", "btu"], 1055.05585262]]);
units("power", [[["HP", "h"], HP], [["PS"], 735.49875], [["W", "w"], 1, true]]);
units("magnetism", [[["T"], 1, true], [["ga"], 1e-4, true]]);
units("volume", [[["tsp"], 4.92892159375e-6], [["tspm"], 5e-6], [["tbs"], 1.478676478125e-5], [["oz"], 2.95735295625e-5], [["cup"], 2.365882365e-4],
  [["pt", "us_pt"], 4.73176473e-4], [["uk_pt"], 5.6826125e-4], [["qt"], 9.46352946e-4], [["uk_qt"], 1.1365225e-3], [["gal"], 3.785411784e-3],
  [["uk_gal"], 4.54609e-3], [["l", "L", "lt"], 1e-3, true], [["ang3", "ang^3"], 1e-30, true], [["barrel"], 0.158987294928],
  [["bushel"], 0.03523907016688], [["ft3", "ft^3"], FT ** 3], [["in3", "in^3"], IN ** 3], [["ly3", "ly^3"], LY ** 3], [["m3", "m^3"], 1, true],
  [["mi3", "mi^3"], MI ** 3], [["yd3", "yd^3"], 0.9144 ** 3], [["Nmi3", "Nmi^3"], NMI ** 3], [["Picapt3", "Picapt^3", "Pica3", "Pica^3"], PICA_PT ** 3],
  [["GRT", "regton"], 100 * FT ** 3], [["MTON"], 40 * FT ** 3]]);
units("area", [[["uk_acre"], 4046.8564224], [["us_acre"], 4046.872609874252], [["ang2", "ang^2"], 1e-20, true], [["ar"], 100, true],
  [["ft2", "ft^2"], FT * FT], [["ha"], 10000], [["in2", "in^2"], IN * IN], [["ly2", "ly^2"], LY * LY], [["m2", "m^2"], 1, true], [["Morgen"], 2500],
  [["mi2", "mi^2"], MI * MI], [["Nmi2", "Nmi^2"], NMI * NMI], [["Picapt2", "Pica2", "Pica^2", "Picapt^2"], PICA_PT * PICA_PT], [["yd2", "yd^2"], 0.9144 ** 2]]);
units("information", [[["bit"], 1, true], [["byte"], 8, true]]);
units("speed", [[["admkn"], (6080 * FT) / 3600], [["kn"], NMI / 3600], [["m/h", "m/hr"], 1 / 3600, true], [["m/s", "m/sec"], 1, true], [["mph"], MI / 3600]]);
const TEMPS = new Set(["C", "cel", "F", "fah", "K", "kel", "Rank", "Reau"]);

const PREFIX: Record<string, number> = {
  Y: 1e24, Z: 1e21, E: 1e18, P: 1e15, T: 1e12, G: 1e9, M: 1e6, k: 1e3, h: 1e2, da: 10, e: 10, d: 1e-1, c: 1e-2, m: 1e-3, u: 1e-6,
  n: 1e-9, p: 1e-12, f: 1e-15, a: 1e-18, z: 1e-21, y: 1e-24,
};
const BIN_PREFIX: Record<string, number> = { ki: 2 ** 10, Mi: 2 ** 20, Gi: 2 ** 30, Ti: 2 ** 40, Pi: 2 ** 50, Ei: 2 ** 60, Zi: 2 ** 70, Yi: 2 ** 80 };

function unit(u: string): [string, number] | null {
  if (TEMPS.has(u)) return ["temperature", 1];
  const d = UNITS[u];
  if (d) return [d[0], d[1]];
  for (const [p, f] of Object.entries(BIN_PREFIX)) {
    const b = UNITS[u.slice(p.length)];
    if (u.startsWith(p) && b && b[0] === "information") return [b[0], b[1] * f];
  }
  for (const [p, f] of Object.entries(PREFIX)) {
    if (!u.startsWith(p)) continue;
    const b = UNITS[u.slice(p.length)];
    if (!b || !b[2]) continue;
    // Squared and cubed units take the prefix squared or cubed.
    const pow = b[0] === "area" ? 2 : b[0] === "volume" && u.slice(p.length) !== "l" && u.slice(p.length) !== "L" && u.slice(p.length) !== "lt" ? 3 : 1;
    return [b[0], b[1] * f ** pow];
  }
  return null;
}

function toKelvin(x: number, u: string): number {
  switch (u) {
    case "C": case "cel": return x + 273.15;
    case "F": case "fah": return ((x - 32) * 5) / 9 + 273.15;
    case "Rank": return (x * 5) / 9;
    case "Reau": return x * 1.25 + 273.15;
    default: return x;
  }
}
function fromKelvin(k: number, u: string): number {
  switch (u) {
    case "C": case "cel": return k - 273.15;
    case "F": case "fah": return ((k - 273.15) * 9) / 5 + 32;
    case "Rank": return (k * 9) / 5;
    case "Reau": return (k - 273.15) / 1.25;
    default: return k;
  }
}

defS("CONVERT", 3, 3, ([a, f0, t0]) => {
  const x = n(a);
  if (isErr(x)) return x;
  if (isErr(f0)) return f0;
  if (isErr(t0)) return t0;
  if (typeof f0 !== "string" || typeof t0 !== "string") return E.na;
  const from = unit(f0), to = unit(t0);
  if (!from || !to || from[0] !== to[0]) return E.na;
  if (from[0] === "temperature") return num(fromKelvin(toKelvin(x, f0), t0));
  return num((x * from[1]) / to[1]);
});
