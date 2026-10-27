// Lookup and reference functions (SVERWEIS, WVERWEIS, XVERWEIS, VERGLEICH, XVERGLEICH,
// INDEX, VERWEIS) and the criteria functions (ZÄHLENWENN … MINWENNS).

import { Ctx, def, deref, mapN, scalar } from "./eval";
import { toArr } from "./fn-util";
import { matcher, type Matcher } from "./criteria";
import { wildcard } from "./fn-text";
import { Arr, E, Ref, compare, isErr, num, toBool, toNum, type Any, type V } from "./values";
import type { CellError } from "../model";

// ----- Matching ----------------------------------------------------------------------------

function sameType(a: V, b: V): boolean {
  if (typeof a === "number") return typeof b === "number";
  if (typeof a === "string") return typeof b === "string";
  if (typeof a === "boolean") return typeof b === "boolean";
  return false;
}

/** Exact match: same type, text case-insensitive with wildcards when `wild`. */
function exactTest(x: V, wild: boolean): (v: V) => boolean {
  if (typeof x === "string") {
    if (wild && /[*?~]/.test(x)) {
      const re = new RegExp("^" + wildcard(x) + "$", "i");
      return (v) => typeof v === "string" && re.test(v);
    }
    const lx = x.toLowerCase();
    return (v) => typeof v === "string" && (v === x || v.toLowerCase() === lx || compare(v, x) === 0);
  }
  return (v) => v !== null && sameType(v, x) && compare(v, x) === 0;
}

function exactIndex(list: V[], x: V, wild: boolean, reverse = false): number {
  const t = exactTest(x, wild);
  if (reverse) {
    for (let i = list.length - 1; i >= 0; i--) if (t(list[i])) return i;
  } else {
    for (let i = 0; i < list.length; i++) if (t(list[i])) return i;
  }
  return -1;
}

/**
 * Excel's approximate match on sorted data (binary search): the last position whose value is
 * not greater than x (ascending), or not smaller (descending). Values of another type than x
 * are stepped over.
 */
function approxIndex(list: V[], x: V, desc = false): number {
  let lo = 0, hi = list.length - 1, found = -1;
  while (lo <= hi) {
    let mid = (lo + hi) >> 1;
    // Step over values of another type, downwards first.
    let m = mid;
    while (m >= lo && !sameType(list[m], x)) m--;
    if (m < lo) {
      m = mid + 1;
      while (m <= hi && !sameType(list[m], x)) m++;
      if (m > hi) break;
    }
    mid = m;
    const c = compare(list[mid], x);
    if (c === 0) {
      // A hit ends the search; equal values right after it are taken too (the last one wins).
      while (mid + 1 < list.length && sameType(list[mid + 1], x) && compare(list[mid + 1], x) === 0) mid++;
      return mid;
    }
    if (desc ? c > 0 : c < 0) {
      found = mid;
      lo = mid + 1;
    } else {
      hi = mid - 1;
    }
  }
  return found;
}

/** A column (or row) of a lookup table as values. */
function column(a: Arr, i: number, byRow: boolean): V[] {
  const out: V[] = [];
  if (byRow) for (let c = 0; c < a.cols; c++) out.push(a.get(i, c));
  else for (let r = 0; r < a.rows; r++) out.push(a.get(r, i));
  return out;
}

function lookupTable(a: Any, ctx: Ctx): Arr | CellError {
  if (isErr(a)) return a;
  return toArr(a, ctx);
}

function vh(name: string, byRow: boolean): void {
  def(name, 3, 4, ([x0, t0, i0, a0], ctx) => {
    const table = lookupTable(t0, ctx);
    if (isErr(table)) return table;
    const xs = deref(x0, ctx), idx = deref(i0, ctx);
    const ap = a0 === undefined ? true : deref(a0, ctx);
    return mapN([xs, idx, ap as V | Arr], ([x, i, a]) => {
      if (isErr(x)) return x;
      if (isErr(i)) return i;
      const k = toNum(i);
      if (isErr(k)) return k;
      const col = Math.trunc(k);
      if (col < 1) return E.value;
      if (col > (byRow ? table.rows : table.cols)) return E.ref;
      const approx = a === null ? false : toBool(a);
      if (isErr(approx)) return approx;
      const keys = column(table, 0, byRow);
      const pos = approx ? approxIndex(keys, x === null ? 0 : x) : exactIndex(keys, x === null ? 0 : x, true);
      if (pos < 0) return E.na;
      const v = byRow ? table.get(col - 1, pos) : table.get(pos, col - 1);
      return v === null ? 0 : v;
    });
  });
}
vh("VLOOKUP", false);
vh("HLOOKUP", true);

/** The values of a one-row or one-column lookup range; null when it is two-dimensional. */
function vector(a: Arr): V[] | null {
  if (a.rows === 1) return a.d;
  if (a.cols === 1) return a.d;
  return null;
}

def("MATCH", 2, 3, ([x0, l0, t0], ctx) => {
  const list0 = lookupTable(l0, ctx);
  if (isErr(list0)) return list0;
  const list = vector(list0);
  if (!list) return E.na;
  const xs = deref(x0, ctx);
  const ts = t0 === undefined ? 1 : deref(t0, ctx);
  return mapN([xs, ts as V | Arr], ([x, t]) => {
    if (isErr(x)) return x;
    if (isErr(t)) return t;
    const k = t === null ? 0 : toNum(t);
    if (isErr(k)) return k;
    const mode = Math.sign(Math.trunc(k));
    const key = x === null ? 0 : x;
    const pos = mode === 0 ? exactIndex(list, key, true) : mode > 0 ? approxIndex(list, key) : descIndex(list, key);
    return pos < 0 ? E.na : pos + 1;
  });
});

/** VERGLEICH with -1: down the list while values are not smaller than x; the last such one. */
function descIndex(list: V[], x: V): number {
  let found = -1;
  for (let i = 0; i < list.length; i++) {
    const v = list[i];
    if (!sameType(v, x)) continue;
    const c = compare(v, x);
    if (c < 0) break;
    found = i;
    if (c === 0) break;
  }
  return found;
}

/** XVERWEIS/XVERGLEICH: Vergleichsmodus 0 exact, -1 exact or next smaller, 1 exact or next larger, 2 wildcards; Suchmodus 1, -1, 2, -2. */
function xIndex(list: V[], x: V, mode: number, search: number): number {
  if (search === 2 || search === -2) {
    // Binary search on data sorted ascending (2) or descending (-2).
    const desc = search === -2;
    if (mode === 0 || mode === 2) {
      const i = approxIndex(list, x, desc);
      return i >= 0 && compare(list[i], x) === 0 ? i : -1;
    }
    const i = approxIndex(list, x, desc);
    if (i >= 0 && compare(list[i], x) === 0) return i;
    if (mode === -1) return desc ? (i + 1 < list.length ? i + 1 : -1) : i;
    return desc ? i : (i + 1 < list.length ? i + 1 : -1);
  }
  const rev = search === -1;
  if (mode === 0 || mode === 2) return exactIndex(list, x, mode === 2, rev);
  const ex = exactIndex(list, x, false, rev);
  if (ex >= 0) return ex;
  let best = -1;
  const order = rev ? [...list.keys()].reverse() : [...list.keys()];
  for (const i of order) {
    const v = list[i];
    if (!sameType(v, x)) continue;
    const c = compare(v, x);
    if (mode === -1 && c < 0 && (best < 0 || compare(v, list[best]) > 0)) best = i;
    if (mode === 1 && c > 0 && (best < 0 || compare(v, list[best]) < 0)) best = i;
  }
  return best;
}

def("XLOOKUP", 3, 6, ([x0, l0, r0, nf0, m0, s0], ctx, nodes) => {
  const nf = nodes.length > 3 && nodes[3].t !== "miss" ? nf0 : undefined;
  const look = lookupTable(l0, ctx);
  if (isErr(look)) return look;
  const list = vector(look);
  if (!list) return E.value;
  const ret = lookupTable(r0, ctx);
  if (isErr(ret)) return ret;
  const byCol = look.rows === 1 && list.length > 1 ? true : look.cols !== 1;
  if (byCol ? ret.cols !== look.cols : ret.rows !== look.rows) return E.value;
  const mode = m0 === undefined || m0 === null ? 0 : toNum(scalar(m0, ctx));
  const search = s0 === undefined || s0 === null ? 1 : toNum(scalar(s0, ctx));
  if (isErr(mode)) return mode;
  if (isErr(search)) return search;
  if (![0, -1, 1, 2].includes(mode) || ![1, -1, 2, -2].includes(search)) return E.value;
  const xs = deref(x0, ctx);
  const one = (x: V): V | Arr => {
    if (isErr(x)) return x;
    const i = xIndex(list, x === null ? "" : x, mode, search);
    if (i < 0) {
      if (nf === undefined) return E.na;
      const v = deref(nf, ctx);
      return v === null ? 0 : v;
    }
    if (byCol) {
      if (ret.rows === 1) return ret.get(0, i) ?? 0;
      return new Arr(ret.rows, 1, column(ret, i, false).map((v) => v ?? 0));
    }
    if (ret.cols === 1) return ret.get(i, 0) ?? 0;
    return new Arr(1, ret.cols, column(ret, i, true).map((v) => v ?? 0));
  };
  if (xs instanceof Arr) return mapN([xs], ([x]) => {
    const r = one(x);
    return r instanceof Arr ? r.d[0] : r;
  });
  return one(xs);
});

def("XMATCH", 2, 4, ([x0, l0, m0, s0], ctx) => {
  const look = lookupTable(l0, ctx);
  if (isErr(look)) return look;
  const list = vector(look);
  if (!list) return E.value;
  const mode = m0 === undefined || m0 === null ? 0 : toNum(scalar(m0, ctx));
  const search = s0 === undefined || s0 === null ? 1 : toNum(scalar(s0, ctx));
  if (isErr(mode)) return mode;
  if (isErr(search)) return search;
  return mapN([deref(x0, ctx)], ([x]) => {
    if (isErr(x)) return x;
    const i = xIndex(list, x === null ? "" : x, mode, search);
    return i < 0 ? E.na : i + 1;
  });
});

def("LOOKUP", 2, 3, ([x0, l0, r0], ctx) => {
  const look = lookupTable(l0, ctx);
  if (isErr(look)) return look;
  let keys: V[], vals: V[];
  if (r0 === undefined) {
    // Array form: search the first row or column of the longer side, return the last one.
    const byRow = look.cols > look.rows;
    keys = column(look, 0, byRow);
    vals = column(look, byRow ? look.rows - 1 : look.cols - 1, byRow);
  } else {
    const ret = lookupTable(r0, ctx);
    if (isErr(ret)) return ret;
    const k = vector(look), v = vector(ret);
    if (!k || !v) return E.na;
    keys = k;
    vals = v;
  }
  return mapN([deref(x0, ctx)], ([x]) => {
    if (isErr(x)) return x;
    const i = approxIndex(keys, x === null ? 0 : x);
    if (i < 0) return E.na;
    if (i >= vals.length) return E.na;
    return vals[i] ?? 0;
  });
});

def("INDEX", 2, 4, ([a, r0, c0, k0], ctx, nodes) => {
  if (isErr(a)) return a;
  const area0 = k0 === undefined ? 1 : toNum(scalar(k0, ctx));
  if (isErr(area0)) return area0;
  const rv = deref(r0, ctx);
  const cv = c0 === undefined ? undefined : deref(c0, ctx);
  if (rv instanceof Arr || cv instanceof Arr) {
    return mapN([rv, cv ?? null], ([r, c]) => {
      const out = indexOne(a, r, cv === undefined ? undefined : c, area0, ctx, nodes.length);
      const d = deref(out, ctx);
      return d instanceof Arr ? d.d[0] : d === null ? 0 : d;
    });
  }
  return indexOne(a, rv, cv, area0, ctx, nodes.length);
});

function indexOne(a: Any, r: V, c: V | undefined, area: number, _ctx: Ctx, argc: number): Any {
  if (isErr(r)) return r;
  if (c !== undefined && isErr(c)) return c;
  const rr = r === null ? 0 : toNum(r);
  const cc = c === undefined || c === null ? (argc >= 3 ? 0 : undefined) : toNum(c);
  if (isErr(rr)) return rr;
  if (cc !== undefined && isErr(cc)) return cc;
  let row = Math.trunc(rr);
  let col = cc === undefined ? undefined : Math.trunc(cc);
  if (row < 0 || (col !== undefined && col < 0)) return E.value;
  if (a instanceof Ref) {
    const ai = Math.trunc(area) - 1;
    if (ai < 0 || ai >= a.areas.length) return E.ref;
    const ar = a.areas[ai];
    const rows = ar.r1 - ar.r0 + 1, cols = ar.c1 - ar.c0 + 1;
    // With one index on a single row or column, it counts along it.
    if (col === undefined) {
      // One index counts along a single row or column; a two-dimensional area needs both.
      if (rows === 1 && cols !== 1) { col = row; row = 1; }
      else if (cols === 1) col = 1;
      else return E.ref;
    }
    if (row > rows || col > cols) return E.ref;
    const r0 = row === 0 ? ar.r0 : ar.r0 + row - 1, r1 = row === 0 ? ar.r1 : r0;
    const c0 = col === 0 ? ar.c0 : ar.c0 + col - 1, c1 = col === 0 ? ar.c1 : c0;
    return new Ref(a.sheet, [{ r0, c0, r1, c1 }]);
  }
  const arr = a instanceof Arr ? a : new Arr(1, 1, [a]);
  if (col === undefined) {
    if (arr.rows === 1 && arr.cols !== 1) { col = row; row = 1; }
    else if (arr.cols === 1) col = 1;
    else return E.ref;
  }
  if (row > arr.rows || col > arr.cols) return E.ref;
  if (row === 0 && col === 0) return arr;
  if (row === 0) return new Arr(arr.rows, 1, column(arr, col - 1, false));
  if (col === 0) return new Arr(1, arr.cols, column(arr, row - 1, true));
  return arr.get(row - 1, col - 1);
}

// ----- Criteria functions ------------------------------------------------------------------

function rangeArg(a: Any, ctx: Ctx): Arr | CellError {
  if (isErr(a)) return a;
  if (!(a instanceof Ref)) return E.value;
  return ctx.refArr(a);
}

/** The cells of `target` (sized like `shape`, from its top-left) whose criteria all match. */
function whereAll(pairs: { range: Arr; test: Matcher }[]): number[] | CellError {
  const { rows, cols } = pairs[0].range;
  for (const p of pairs) if (p.range.rows !== rows || p.range.cols !== cols) return E.value;
  const out: number[] = [];
  outer: for (let i = 0; i < rows * cols; i++) {
    for (const p of pairs) if (!p.test(p.range.d[i])) continue outer;
    out.push(i);
  }
  return out;
}

/** The sum (or other) range of SUMMEWENN, resized like the criteria range from its corner. */
function resized(a: Any, like: Arr, ctx: Ctx): Arr | CellError {
  if (!(a instanceof Ref)) return isErr(a) ? a : E.value;
  const r = a.areas[0];
  return ctx.refArr(new Ref(a.sheet, [{ r0: r.r0, c0: r.c0, r1: r.r0 + like.rows - 1, c1: r.c0 + like.cols - 1 }]));
}

function critValues(c: Any, ctx: Ctx): V | Arr {
  const d = deref(c, ctx);
  return d;
}

def("COUNTIF", 2, 2, ([r0, c0], ctx) => {
  const range = rangeArg(r0, ctx);
  if (isErr(range)) return range;
  return mapN([critValues(c0, ctx)], ([c]) => {
    const t = matcher(c);
    let k = 0;
    for (const v of range.d) if (t(v)) k++;
    return k;
  });
});

function ifsAgg(name: string, first: number, f: (xs: V[]) => V): void {
  // first = 1: (target, range1, crit1, …); the *IFS layout.
  def(name, 3, 255, (args, ctx) => {
    if ((args.length - first) % 2) return E.value;
    const target = rangeArg(args[0], ctx);
    if (isErr(target)) return target;
    const ranges: Arr[] = [];
    for (let i = first; i < args.length; i += 2) {
      const r = rangeArg(args[i], ctx);
      if (isErr(r)) return r;
      ranges.push(r);
    }
    const crits = [];
    for (let i = first + 1; i < args.length; i += 2) crits.push(critValues(args[i], ctx));
    return mapN(crits, (cs) => {
      const idx = whereAll(ranges.map((range, j) => ({ range, test: matcher(cs[j]) })));
      if (isErr(idx)) return idx;
      if (ranges[0].rows !== target.rows || ranges[0].cols !== target.cols) return E.value;
      return f(idx.map((i) => target.d[i]));
    });
  });
}

const sumOf = (xs: V[]): V => {
  let s = 0;
  for (const v of xs) {
    if (isErr(v)) return v;
    if (typeof v === "number") s += v;
  }
  return num(s);
};
const avgOf = (xs: V[]): V => {
  let s = 0, k = 0;
  for (const v of xs) {
    if (isErr(v)) return v;
    if (typeof v === "number") { s += v; k++; }
  }
  return k ? num(s / k) : E.div0;
};
const maxOf = (xs: V[]): V => {
  let m = -Infinity;
  for (const v of xs) {
    if (isErr(v)) return v;
    if (typeof v === "number" && v > m) m = v;
  }
  return m === -Infinity ? 0 : m;
};
const minOf = (xs: V[]): V => {
  let m = Infinity;
  for (const v of xs) {
    if (isErr(v)) return v;
    if (typeof v === "number" && v < m) m = v;
  }
  return m === Infinity ? 0 : m;
};

ifsAgg("SUMIFS", 1, sumOf);
ifsAgg("AVERAGEIFS", 1, avgOf);
ifsAgg("MAXIFS", 1, maxOf);
ifsAgg("MINIFS", 1, minOf);

def("COUNTIFS", 2, 254, (args, ctx) => {
  if (args.length % 2) return E.value;
  const ranges: Arr[] = [];
  for (let i = 0; i < args.length; i += 2) {
    const r = rangeArg(args[i], ctx);
    if (isErr(r)) return r;
    ranges.push(r);
  }
  const crits = [];
  for (let i = 1; i < args.length; i += 2) crits.push(critValues(args[i], ctx));
  return mapN(crits, (cs) => {
    const idx = whereAll(ranges.map((range, j) => ({ range, test: matcher(cs[j]) })));
    return isErr(idx) ? idx : idx.length;
  });
});

function ifAgg(name: string, f: (xs: V[]) => V): void {
  def(name, 2, 3, ([r0, c0, s0], ctx) => {
    const range = rangeArg(r0, ctx);
    if (isErr(range)) return range;
    const target = s0 === undefined || s0 === null ? range : resized(s0, range, ctx);
    if (isErr(target)) return target;
    return mapN([critValues(c0, ctx)], ([c]) => {
      const t = matcher(c);
      const xs: V[] = [];
      for (let i = 0; i < range.d.length; i++) if (t(range.d[i])) xs.push(target.d[i]);
      return f(xs);
    });
  });
}
ifAgg("SUMIF", sumOf);
ifAgg("AVERAGEIF", avgOf);

