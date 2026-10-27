// Mathematical functions (Mathematik und Trigonometrie).

import { def } from "./eval";
import { defS, n, numbers, roundTo, toArr, r15 } from "./fn-util";
import { Arr, E, isErr, num, type V } from "./values";
import { power, snap } from "./eval";

type Num1 = (x: number) => number | V;

/** A function of one number. */
function math1(name: string, f: Num1): void {
  defS(name, 1, 1, ([a]) => {
    const x = n(a);
    if (isErr(x)) return x;
    const r = f(x);
    return typeof r === "number" ? num(r) : r;
  });
}

/** A function of two numbers (the second optional, default `d`). */
function math2(name: string, min: number, f: (x: number, y: number) => number | V, d = 0): void {
  defS(name, min, 2, ([a, b]) => {
    const x = n(a);
    if (isErr(x)) return x;
    const y = b === undefined ? d : n(b);
    if (isErr(y)) return y;
    const r = f(x, y);
    return typeof r === "number" ? num(r) : r;
  });
}

def("SUM", 1, 255, (args, ctx) => {
  const xs = numbers(args, ctx);
  if (isErr(xs)) return xs;
  let s = 0;
  for (const x of xs) s = snap(s + x, s, x);
  return num(s);
});

def("PRODUCT", 1, 255, (args, ctx) => {
  const xs = numbers(args, ctx);
  if (isErr(xs)) return xs;
  if (xs.length === 0) return 0;
  let p = 1;
  for (const x of xs) p *= x;
  return num(p);
});

def("SUMSQ", 1, 255, (args, ctx) => {
  const xs = numbers(args, ctx);
  if (isErr(xs)) return xs;
  return num(xs.reduce((s, x) => s + x * x, 0));
});

math2("ROUND", 2, (x, d) => roundTo(x, d, "half"));
math2("ROUNDUP", 2, (x, d) => roundTo(x, d, "up"));
math2("ROUNDDOWN", 2, (x, d) => roundTo(x, d, "down"));
math2("TRUNC", 1, (x, d) => roundTo(x, d, "down"));
math1("INT", (x) => Math.floor(r15(x)));
math1("ABS", Math.abs);
math1("SIGN", Math.sign);
math1("SQRT", (x) => (x < 0 ? E.num : Math.sqrt(x)));
math1("SQRTPI", (x) => (x < 0 ? E.num : Math.sqrt(x * Math.PI)));
math1("EXP", Math.exp);
math1("LN", (x) => (x <= 0 ? E.num : Math.log(x)));
math1("LOG10", (x) => (x <= 0 ? E.num : Math.log10(x)));
math2("LOG", 1, (x, base) => {
  if (x <= 0 || base <= 0) return E.num;
  if (base === 1) return E.div0;
  return base === 10 ? Math.log10(x) : Math.log(x) / Math.log(base);
}, 10);
math2("POWER", 2, (x, y) => {
  if (x === 0 && y === 0) return E.num;
  if (x === 0 && y < 0) return E.div0;
  return power(x, y);
});
math2("MOD", 2, (x, y) => {
  if (y === 0) return E.div0;
  const q = x / y;
  // Excel gives up when the quotient's whole part exceeds 2^40 (recorded).
  if (Math.abs(Math.trunc(q)) > 2 ** 40) return E.num;
  // The quotient's whole part and the remainder are exact, as in Excel (=REST(0,5;0,1) is
  // 0,09999999999999998 because 0,1 is a little more than a tenth); the remainder takes the
  // divisor's sign, and one that rounds to the divisor is 0.
  let k = Math.floor(q);
  let r = residual(x, y, k);
  if (y > 0 ? r < 0 : r > 0) r = residual(x, y, --k);
  else if (y > 0 ? r > y : r < y) r = residual(x, y, ++k);
  return r === y ? 0 : r;
});
math2("QUOTIENT", 2, (x, y) => (y === 0 ? E.div0 : Math.trunc(x / y)));

const trigLimit = 2 ** 27;
math1("SIN", (x) => (Math.abs(x) >= trigLimit ? E.num : Math.sin(x)));
math1("COS", (x) => (Math.abs(x) >= trigLimit ? E.num : Math.cos(x)));
math1("TAN", (x) => (Math.abs(x) >= trigLimit ? E.num : Math.tan(x)));
math1("COT", (x) => (x === 0 ? E.div0 : Math.abs(x) >= trigLimit ? E.num : 1 / Math.tan(x)));
math1("ASIN", (x) => (Math.abs(x) > 1 ? E.num : Math.asin(x)));
math1("ACOS", (x) => (Math.abs(x) > 1 ? E.num : Math.acos(x)));
math1("ATAN", Math.atan);
math2("ATAN2", 2, (x, y) => (x === 0 && y === 0 ? E.div0 : Math.atan2(y, x)));
math1("SINH", Math.sinh);
math1("COSH", Math.cosh);
math1("TANH", Math.tanh);
math1("ASINH", Math.asinh);
math1("ACOSH", (x) => (x < 1 ? E.num : Math.acosh(x)));
math1("ATANH", (x) => (Math.abs(x) >= 1 ? E.num : Math.atanh(x)));
math1("RADIANS", (x) => (x * Math.PI) / 180);
math1("DEGREES", (x) => (x * 180) / Math.PI);
def("PI", 0, 0, () => Math.PI);

/** x − y·k with the product taken exactly (Dekker's two-product). */
function residual(x: number, y: number, k: number): number {
  const p = y * k;
  const split = (a: number) => {
    const c = 134217729 * a;
    const hi = c - (c - a);
    return [hi, a - hi];
  };
  const [yh, yl] = split(y), [kh, kl] = split(k);
  const e = ((yh * kh - p) + yh * kl + yl * kh) + yl * kl;
  return (x - p) - e;
}

export function fact(x: number): number {
  let r = 1;
  for (let i = 2; i <= x; i++) r *= i;
  return r;
}

math1("FACT", (x) => {
  x = Math.trunc(x);
  if (x < 0) return E.num;
  if (x > 170) return E.num;
  return fact(x);
});
math1("FACTDOUBLE", (x) => {
  x = Math.trunc(x);
  if (x < -1) return E.num;
  let r = 1;
  for (let i = x; i > 1; i -= 2) r *= i;
  return r;
});

export function combin(nn: number, k: number): number {
  if (k > nn - k) k = nn - k;
  let r = 1;
  for (let i = 1; i <= k; i++) r = (r * (nn - k + i)) / i;
  return Math.round(r) === r || r > 2 ** 53 ? r : Math.round(r);
}

math2("COMBIN", 2, (x, k) => {
  x = Math.trunc(x);
  k = Math.trunc(k);
  if (x < 0 || k < 0 || k > x) return E.num;
  return combin(x, k);
});
math2("COMBINA", 2, (x, k) => {
  x = Math.trunc(x);
  k = Math.trunc(k);
  if (x < 0 || k < 0) return E.num;
  return k === 0 ? 1 : combin(x + k - 1, k);
});
math2("PERMUT", 2, (x, k) => {
  x = Math.trunc(x);
  k = Math.trunc(k);
  if (x < 0 || k < 0 || k > x) return E.num;
  let r = 1;
  for (let i = 0; i < k; i++) r *= x - i;
  return r;
});

function gcd2(a: number, b: number): number {
  while (b) [a, b] = [b, a % b];
  return a;
}

def("GCD", 1, 255, (args, ctx) => {
  const xs = numbers(args, ctx);
  if (isErr(xs)) return xs;
  let g = 0;
  for (const x0 of xs) {
    if (x0 < 0 || x0 > 2 ** 53) return E.num;
    g = gcd2(g, Math.trunc(x0));
  }
  return g;
});
def("LCM", 1, 255, (args, ctx) => {
  const xs = numbers(args, ctx);
  if (isErr(xs)) return xs;
  let l = 1;
  for (const x0 of xs) {
    if (x0 < 0) return E.num;
    const x = Math.trunc(x0);
    if (x === 0) return 0;
    l = (l / gcd2(l, x)) * x;
  }
  return num(l);
});

math1("EVEN", (x) => {
  const a = Math.ceil(Math.abs(x) - 1e-15 * Math.abs(x));
  const e = a % 2 === 0 ? a : a + 1;
  return x < 0 ? -e : e;
});
math1("ODD", (x) => {
  let a = Math.ceil(Math.abs(x) - 1e-15 * Math.abs(x));
  if (a % 2 === 0) a += 1;
  return x < 0 ? -a : a;
});

/** Rounds x to a multiple of step: up (towards +∞) or down. */
function toMultiple(x: number, step: number, up: boolean): number {
  const q = r15(x / step);
  return (up ? Math.ceil(q) : Math.floor(q)) * step;
}

defS("CEILING.MATH", 1, 3, ([a, st, md]) => {
  const x = n(a);
  if (isErr(x)) return x;
  const step0 = st === undefined || st === null ? 1 : n(st);
  if (isErr(step0)) return step0;
  const mode = md === undefined || md === null ? 0 : n(md);
  if (isErr(mode)) return mode;
  const step = Math.abs(step0);
  if (step === 0 || x === 0) return 0;
  // Negative numbers round towards 0, or away from it when Modus is not 0.
  if (x < 0 && mode !== 0) return num(-toMultiple(-x, step, true));
  return num(toMultiple(x, step, true));
});
defS("FLOOR.MATH", 1, 3, ([a, st, md]) => {
  const x = n(a);
  if (isErr(x)) return x;
  const step0 = st === undefined || st === null ? 1 : n(st);
  if (isErr(step0)) return step0;
  const mode = md === undefined || md === null ? 0 : n(md);
  if (isErr(mode)) return mode;
  const step = Math.abs(step0);
  if (step === 0 || x === 0) return 0;
  if (x < 0 && mode !== 0) return num(-toMultiple(-x, step, false));
  return num(toMultiple(x, step, false));
});
math2("CEILING", 2, (x, step) => {
  if (step === 0) return 0;
  if (x > 0 && step < 0) return E.num;
  if (x < 0 && step > 0) return -toMultiple(-x, step, false);
  return x < 0 ? -toMultiple(-x, -step, true) : toMultiple(x, step, true);
});
math2("FLOOR", 2, (x, step) => {
  if (step === 0) return x === 0 ? 0 : E.div0;
  if (x > 0 && step < 0) return E.num;
  if (x < 0 && step > 0) return -toMultiple(-x, step, true);
  return x < 0 ? -toMultiple(-x, -step, false) : toMultiple(x, step, false);
});
math2("MROUND", 2, (x, m) => {
  if (m === 0) return 0;
  if (x * m < 0) return E.num;
  // The quotient rounded at 15 digits, then half away from zero (=VRUNDEN(0,15;0,1) is 0,2).
  const q = r15(x / m);
  return Math.sign(q) * Math.floor(Math.abs(q) + 0.5) * m;
});

def("SUMPRODUCT", 1, 255, (args, ctx) => {
  const arrs: Arr[] = [];
  for (const a of args) {
    const r = toArr(a, ctx);
    if (isErr(r)) return r;
    arrs.push(r);
  }
  const { rows, cols } = arrs[0];
  if (arrs.some((a) => a.rows !== rows || a.cols !== cols)) return E.value;
  let s = 0;
  for (let i = 0; i < rows * cols; i++) {
    let p = 1;
    for (const a of arrs) {
      const v = a.d[i];
      if (isErr(v)) return v;
      p *= typeof v === "number" ? v : 0;
    }
    s += p;
  }
  return num(s);
});

def("RAND", 0, 0, (_a, ctx) => ctx.host.random(), { volatile: true });
defS("RANDBETWEEN", 2, 2, ([a, b], ctx) => {
  const lo = n(a), hi = n(b);
  if (isErr(lo)) return lo;
  if (isErr(hi)) return hi;
  const l = Math.ceil(lo), h = Math.ceil(hi);
  if (l > h) return E.num;
  return l + Math.floor(ctx.host.random() * (h - l + 1));
}, [], { volatile: true });
def("RANDARRAY", 0, 5, (args, ctx) => {
  const get = (i: number, d: number) => {
    const a = args[i];
    if (a === undefined || a === null) return d;
    return n(a as V);
  };
  const rows = get(0, 1), cols = get(1, 1), lo = get(2, 0), hi = get(3, 1), whole = get(4, 0);
  for (const x of [rows, cols, lo, hi, whole]) if (isErr(x)) return x;
  const r = Math.trunc(rows as number), c = Math.trunc(cols as number);
  if (r < 1 || c < 1 || (lo as number) > (hi as number)) return E.value;
  const d: V[] = [];
  for (let i = 0; i < r * c; i++) {
    const u = ctx.host.random();
    d.push(whole ? Math.ceil(lo as number) + Math.floor(u * (Math.floor(hi as number) - Math.ceil(lo as number) + 1)) : (lo as number) + u * ((hi as number) - (lo as number)));
  }
  return new Arr(r, c, d);
}, { volatile: true });

// Number systems.
function toBase(x: number, base: number, places: V | undefined, limit: number): V {
  x = Math.trunc(x);
  if (x < -limit || x >= limit) return E.num;
  let s = (x < 0 ? x + 2 * limit : x).toString(base).toUpperCase();
  if (x < 0) return s;
  if (places !== undefined && places !== null) {
    const p = n(places);
    if (isErr(p)) return p;
    if (p < s.length || p > 10) return E.num;
    s = s.padStart(Math.trunc(p), "0");
  }
  return s;
}

defS("DEC2BIN", 1, 2, ([a, p]) => {
  const x = n(a);
  return isErr(x) ? x : toBase(x, 2, p, 512);
});
defS("DEC2HEX", 1, 2, ([a, p]) => {
  const x = n(a);
  return isErr(x) ? x : toBase(x, 16, p, 2 ** 39);
});
function fromBase(text: V, base: number, bits: number): V {
  const t = typeof text === "number" ? String(text) : typeof text === "string" ? text : null;
  if (t === null) return isErr(text) ? text : E.num;
  if (t.length > 10 || !new RegExp(`^[0-9a-${"fghijklmnopqrstuvwxyz".slice(0, Math.max(0, base - 10))}]*$`, "i").test(base > 10 ? t : t.replace(/[^0-9]/g, (c) => c))) return E.num;
  const v = t === "" ? 0 : parseInt(t, base);
  if (Number.isNaN(v) || t.split("").some((c) => parseInt(c, base) >= base || Number.isNaN(parseInt(c, base)))) return E.num;
  return t.length === 10 && v >= 2 ** (bits - 1) ? v - 2 ** bits : v;
}
defS("BIN2DEC", 1, 1, ([a]) => fromBase(a, 2, 10));
defS("HEX2DEC", 1, 1, ([a]) => fromBase(a, 16, 40));
