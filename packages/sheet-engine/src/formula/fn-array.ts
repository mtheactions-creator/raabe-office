// Dynamic array functions (SORTIEREN, SORTIERENNACH, FILTER, EINDEUTIG, SEQUENZ, MTRANS …)
// and the financial basics (RMZ, ZW, BW, ZZR, ZINS, ZINSZ, KAPZ, NBW).

import { Ctx, def, deref, scalar } from "./eval";
import { defS, n, numbers, toArr } from "./fn-util";
import { Arr, E, compare, isErr, num, toBool, toNum, type Any, type V } from "./values";
import type { CellError } from "../model";

/** Sort order of SORTIEREN: numbers, text, logical values, errors, then empty cells. */
function sortCmp(a: V, b: V): number {
  const rank = (v: V) => (v === null ? 4 : isErr(v) ? 3 : 0);
  const ra = rank(a), rb = rank(b);
  if (ra || rb) return ra - rb;
  return compare(a, b);
}

function optNum(a: Any | undefined, ctx: Ctx, d: number): number | CellError {
  if (a === undefined || a === null) return d;
  const v = scalar(a, ctx);
  if (v === null) return d;
  const x = toNum(v);
  return isErr(x) ? x : Math.trunc(x);
}

function rowsOf(a: Arr): V[][] {
  const out: V[][] = [];
  for (let r = 0; r < a.rows; r++) out.push(a.d.slice(r * a.cols, (r + 1) * a.cols));
  return out;
}

function transpose(a: Arr): Arr {
  const d: V[] = new Array(a.rows * a.cols);
  for (let r = 0; r < a.rows; r++) for (let c = 0; c < a.cols; c++) d[c * a.rows + r] = a.d[r * a.cols + c];
  return new Arr(a.cols, a.rows, d);
}

function zeroEmpty(a: Arr): Arr {
  return new Arr(a.rows, a.cols, a.d.map((v) => (v === null ? 0 : v)));
}

def("SORT", 1, 4, ([a0, i0, o0, b0], ctx) => {
  let a = toArr(a0, ctx);
  if (isErr(a)) return a;
  const idx = optNum(i0, ctx, 1), ord = optNum(o0, ctx, 1);
  if (isErr(idx)) return idx;
  if (isErr(ord)) return ord;
  const byCol = b0 === undefined || b0 === null ? false : toBool(scalar(b0, ctx));
  if (isErr(byCol)) return byCol;
  if (ord !== 1 && ord !== -1) return E.value;
  if (byCol) a = transpose(a);
  if (idx < 1 || idx > a.cols) return E.value;
  const rows = rowsOf(a);
  const sorted = rows.map((r, i) => [r, i] as const).sort((x, y) => sortCmp(x[0][idx - 1], y[0][idx - 1]) * ord || x[1] - y[1]).map((x) => x[0]);
  let out = new Arr(a.rows, a.cols, sorted.flat());
  if (byCol) out = transpose(out);
  return zeroEmpty(out);
});

def("SORTBY", 2, 254, ([a0, ...rest], ctx) => {
  const a = toArr(a0, ctx);
  if (isErr(a)) return a;
  const keys: { k: V[]; ord: number }[] = [];
  let byCol: boolean | null = null;
  for (let i = 0; i < rest.length; i += 2) {
    const k = toArr(rest[i], ctx);
    if (isErr(k)) return k;
    const ord = optNum(rest[i + 1], ctx, 1);
    if (isErr(ord)) return ord;
    if (ord !== 1 && ord !== -1) return E.value;
    const col = k.cols === 1 && k.rows === a.rows;
    const row = k.rows === 1 && k.cols === a.cols && !col;
    if (!col && !row) return E.value;
    if (byCol === null) byCol = row && a.rows !== 1 ? true : row && !col;
    keys.push({ k: k.d, ord });
  }
  const src = byCol ? transpose(a) : a;
  const rows = rowsOf(src);
  const order = rows.map((_, i) => i).sort((x, y) => {
    for (const { k, ord } of keys) {
      const c = sortCmp(k[x], k[y]) * ord;
      if (c) return c;
    }
    return x - y;
  });
  let out = new Arr(src.rows, src.cols, order.flatMap((i) => rows[i]));
  if (byCol) out = transpose(out);
  return zeroEmpty(out);
});

def("FILTER", 2, 3, ([a0, i0, e0], ctx) => {
  const a = toArr(a0, ctx);
  if (isErr(a)) return a;
  const inc = toArr(i0, ctx);
  if (isErr(inc)) return inc;
  const byRow = inc.cols === 1 && inc.rows === a.rows;
  const byCol = inc.rows === 1 && inc.cols === a.cols && !byRow;
  if (!byRow && !byCol) return E.value;
  const keep: number[] = [];
  for (let i = 0; i < inc.d.length; i++) {
    const v = inc.d[i];
    if (isErr(v)) return v;
    const t = v === null ? false : toBool(v);
    if (isErr(t)) return E.value;
    if (t) keep.push(i);
  }
  if (!keep.length) {
    if (e0 === undefined) return E.calc;
    const e = deref(e0, ctx);
    return e === null ? 0 : e;
  }
  if (byRow) {
    const rows = rowsOf(a);
    return zeroEmpty(new Arr(keep.length, a.cols, keep.flatMap((i) => rows[i])));
  }
  const t = rowsOf(transpose(a));
  return zeroEmpty(transpose(new Arr(keep.length, a.rows, keep.flatMap((i) => t[i]))));
});

def("UNIQUE", 1, 3, ([a0, b0, o0], ctx) => {
  const a = toArr(a0, ctx);
  if (isErr(a)) return a;
  const byCol = b0 === undefined || b0 === null ? false : toBool(scalar(b0, ctx));
  const once = o0 === undefined || o0 === null ? false : toBool(scalar(o0, ctx));
  if (isErr(byCol)) return byCol;
  if (isErr(once)) return once;
  const src = byCol ? transpose(a) : a;
  const rows = rowsOf(src);
  const key = (r: V[]) => r.map((v) => (v === null ? "\u0001" : typeof v === "string" ? "s" + v.toLowerCase() : typeof v + String(v))).join("\u0000");
  const count = new Map<string, number>();
  for (const r of rows) count.set(key(r), (count.get(key(r)) ?? 0) + 1);
  const seen = new Set<string>();
  const out: V[][] = [];
  for (const r of rows) {
    const k = key(r);
    if (once ? count.get(k) !== 1 : seen.has(k)) continue;
    seen.add(k);
    out.push(r);
  }
  if (!out.length) return E.calc;
  let res = new Arr(out.length, src.cols, out.flat());
  if (byCol) res = transpose(res);
  return zeroEmpty(res);
});

def("SEQUENCE", 1, 4, ([r0, c0, s0, st0], ctx) => {
  const r = optNum(r0, ctx, 1), c = optNum(c0, ctx, 1);
  if (isErr(r)) return r;
  if (isErr(c)) return c;
  const sv = s0 === undefined || s0 === null ? 1 : toNum(scalar(s0, ctx));
  const stv = st0 === undefined || st0 === null ? 1 : toNum(scalar(st0, ctx));
  if (isErr(sv)) return sv;
  if (isErr(stv)) return stv;
  if (r === 0 || c === 0) return E.calc;
  if (r < 0 || c < 0) return E.value;
  if (r * c > 1048576 * 16384) return E.num;
  const d: V[] = new Array(r * c);
  for (let i = 0; i < r * c; i++) d[i] = sv + i * stv;
  return new Arr(r, c, d);
});

def("TRANSPOSE", 1, 1, ([a0], ctx) => {
  const a = toArr(a0, ctx);
  if (isErr(a)) return a;
  return zeroEmpty(transpose(a));
});

def("MMULT", 2, 2, ([a0, b0], ctx) => {
  const a = toArr(a0, ctx), b = toArr(b0, ctx);
  if (isErr(a)) return a;
  if (isErr(b)) return b;
  if (a.cols !== b.rows) return E.value;
  const d: V[] = [];
  for (let r = 0; r < a.rows; r++) {
    for (let c = 0; c < b.cols; c++) {
      let s = 0;
      for (let k = 0; k < a.cols; k++) {
        const x = a.get(r, k), y = b.get(k, c);
        if (typeof x !== "number" || typeof y !== "number") return E.value;
        s += x * y;
      }
      d.push(s);
    }
  }
  return new Arr(a.rows, b.cols, d);
});

// ----- Finance -----------------------------------------------------------------------------

function fin(name: string, min: number, max: number, f: (x: number[]) => number | CellError, defaults: number[]): void {
  defS(name, min, max, (vals) => {
    const xs: number[] = [];
    for (let i = 0; i < max; i++) {
      const v = vals[i];
      if (v === undefined || v === null) {
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

function pmt(rate: number, nper: number, pv: number, fv: number, type: number): number {
  if (rate === 0) return -(pv + fv) / nper;
  const f = (1 + rate) ** nper;
  return (-rate * (pv * f + fv)) / ((1 + rate * (type ? 1 : 0)) * (f - 1));
}

function fv(rate: number, nper: number, p: number, pv: number, type: number): number {
  if (rate === 0) return -(pv + p * nper);
  const f = (1 + rate) ** nper;
  return -(pv * f + (p * (1 + rate * (type ? 1 : 0)) * (f - 1)) / rate);
}

fin("PMT", 3, 5, ([r, np, pv, f, t]) => (np === 0 ? E.num : pmt(r, np, pv, f, t)), []);
fin("FV", 3, 5, ([r, np, p, pv, t]) => fv(r, np, p, pv, t), []);
fin("PV", 3, 5, ([r, np, p, f, t]) => {
  if (r === 0) return -(f + p * np);
  const g = (1 + r) ** np;
  return -(f + (p * (1 + r * (t ? 1 : 0)) * (g - 1)) / r) / g;
}, []);
fin("NPER", 3, 5, ([r, p, pv, f, t]) => {
  if (r === 0) return p === 0 ? E.num : -(pv + f) / p;
  const k = p * (1 + r * (t ? 1 : 0));
  const x = (k - f * r) / (k + pv * r);
  if (x <= 0) return E.num;
  return Math.log(x) / Math.log(1 + r);
}, []);
fin("RATE", 3, 6, ([np, p, pv, f, t, g]) => {
  // Newton's method on pv·(1+r)^n + p·(1+r·t)·((1+r)^n − 1)/r + fv = 0.
  let r = g;
  for (let i = 0; i < 100; i++) {
    const q = (1 + r) ** np;
    const dq = np * (1 + r) ** (np - 1);
    let y: number, d: number;
    if (Math.abs(r) < 1e-12) {
      y = pv + p * np + f;
      d = pv * np + p * (np * (np - 1)) / 2 + p * t * np;
    } else {
      const a = (q - 1) / r;
      const da = (dq * r - (q - 1)) / (r * r);
      y = pv * q + p * (1 + r * t) * a + f;
      d = pv * dq + p * t * a + p * (1 + r * t) * da;
    }
    const nr = r - y / d;
    if (!Number.isFinite(nr)) return E.num;
    if (Math.abs(nr - r) <= 1e-15 * Math.max(1, Math.abs(nr))) return nr;
    r = nr;
  }
  return E.num;
}, [0, 0, 0, 0, 0, 0.1]);
fin("IPMT", 4, 6, ([r, per, np, pv, f, t]) => {
  if (per < 1 || per > np) return E.num;
  const p = pmt(r, np, pv, f, t);
  if (t && per === 1) return 0;
  const bal = fv(r, t ? per - 2 : per - 1, p, pv, t);
  return t ? (bal - p) * r : bal * r;
}, []);
fin("PPMT", 4, 6, ([r, per, np, pv, f, t]) => {
  if (per < 1 || per > np) return E.num;
  const p = pmt(r, np, pv, f, t);
  let ip: number;
  if (t && per === 1) ip = 0;
  else {
    const bal = fv(r, t ? per - 2 : per - 1, p, pv, t);
    ip = t ? (bal - p) * r : bal * r;
  }
  return p - ip;
}, []);
def("NPV", 2, 255, ([r0, ...rest], ctx) => {
  const r = toNum(scalar(r0, ctx));
  if (isErr(r)) return r;
  const xs = numbers(rest, ctx);
  if (isErr(xs)) return xs;
  let s = 0;
  xs.forEach((x, i) => { s += x / (1 + r) ** (i + 1); });
  return num(s);
});
