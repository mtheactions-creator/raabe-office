// Array functions, second round: AGGREGAT, the matrix functions (MDET, MINV, MEINHEIT), the
// shaping functions (ERWEITERN, HSTAPELN, VSTAPELN, SPALTENUMBRUCH, ZEILENUMBRUCH, SPALTENWAHL,
// ZEILENWAHL, ÜBERNEHMEN, WEGLASSEN, ZUSPALTE, ZUZEILE, ABSCHNBEREICH) and the grouping
// functions GRUPPIERENNACH and PIVOTMIT.

import { Ctx, FUNCS, def, deref, type FnDef } from "./eval";
import { toArr } from "./fn-util";
import { englishName } from "./names";
import type { Node } from "./parse";
import { Arr, E, Ref, compare, isErr, num, toBool, toNum, type Any, type V } from "./values";
import type { CellError } from "../model";

/** A whole-number argument; `d` when it is missing or empty. */
function intArg(a: Any | undefined, ctx: Ctx, d: number | null): number | null | CellError {
  if (a === undefined || a === null) return d;
  const v = deref(a, ctx);
  const s = v instanceof Arr ? v.d[0] : v;
  if (s === null) return d;
  const x = toNum(s);
  return isErr(x) ? x : Math.trunc(x);
}

function arrArg(a: Any, ctx: Ctx): Arr | CellError {
  if (isErr(a)) return a;
  return toArr(a, ctx);
}

// ----- AGGREGAT -------------------------------------------------------------------------------

const AGG_FNS: Record<number, string> = {
  1: "AVERAGE", 2: "COUNT", 3: "COUNTA", 4: "MAX", 5: "MIN", 6: "PRODUCT", 7: "STDEV.S", 8: "STDEV.P", 9: "SUM", 10: "VAR.S",
  11: "VAR.P", 12: "MEDIAN", 13: "MODE.SNGL", 14: "LARGE", 15: "SMALL", 16: "PERCENTILE.INC", 17: "QUARTILE.INC", 18: "PERCENTILE.EXC",
  19: "QUARTILE.EXC",
};

def("AGGREGATE", 3, 255, (args, ctx) => {
  const f = intArg(args[0], ctx, null), o = intArg(args[1], ctx, null);
  if (isErr(f)) return f;
  if (isErr(o)) return o;
  if (f === null || o === null || !AGG_FNS[f] || o < 0 || o > 7) return E.value;
  const skipErr = [2, 3, 6, 7].includes(o), skipHidden = [1, 3, 5, 7].includes(o), skipNested = o <= 3;
  const collect = (a: Any): V[] => {
    const out: V[] = [];
    if (a instanceof Ref) {
      const sheet = ctx.host.wb.sheets[a.sheet];
      for (let i = 0; i < a.areas.length; i++) {
        const ar = ctx.clip(a.sheet, a.areas[i]);
        const arr = ctx.refArr(a, i);
        for (let r = 0; r < arr.rows; r++) {
          if (skipHidden && sheet?.rows.get(ar.r0 + r)?.hidden) continue;
          for (let c = 0; c < arr.cols; c++) {
            if (skipNested) {
              const fm = ctx.host.formula(a.sheet, ar.r0 + r, ar.c0 + c);
              if (fm && /SUBTOTAL\(|AGGREGATE\(/i.test(fm)) continue;
            }
            out.push(arr.get(r, c));
          }
        }
      }
    } else if (a instanceof Arr) out.push(...a.d);
    else out.push(a);
    return skipErr ? out.filter((v) => !isErr(v)) : out;
  };
  const fn = FUNCS.get(AGG_FNS[f])!;
  if (f >= 14) {
    if (args.length < 4) return E.value;
    const vals = collect(args[2]);
    return fn.fn([new Arr(vals.length, 1, vals), deref(args[3], ctx)], ctx, []);
  }
  const vals: V[] = [];
  for (const a of args.slice(2)) vals.push(...collect(a));
  return fn.fn([new Arr(vals.length, 1, vals)], ctx, []);
});

// ----- Matrices -------------------------------------------------------------------------------

/** A square matrix of numbers, or #WERT!. */
function square(a: Any, ctx: Ctx): number[][] | CellError {
  const m = arrArg(a, ctx);
  if (isErr(m)) return m;
  if (m.rows !== m.cols || !m.rows) return E.value;
  const out: number[][] = [];
  for (let r = 0; r < m.rows; r++) {
    const row: number[] = [];
    for (let c = 0; c < m.cols; c++) {
      const v = m.get(r, c);
      if (isErr(v)) return v;
      if (typeof v !== "number") return E.value;
      row.push(v);
    }
    out.push(row);
  }
  return out;
}

def("MDETERM", 1, 1, ([a], ctx) => {
  const m = square(a, ctx);
  if (isErr(m)) return m;
  const k = m.length;
  let det = 1;
  for (let c = 0; c < k; c++) {
    let p = c;
    for (let r = c + 1; r < k; r++) if (Math.abs(m[r][c]) > Math.abs(m[p][c])) p = r;
    if (m[p][c] === 0) return 0;
    if (p !== c) {
      [m[p], m[c]] = [m[c], m[p]];
      det = -det;
    }
    det *= m[c][c];
    for (let r = c + 1; r < k; r++) {
      const f = m[r][c] / m[c][c];
      for (let j = c; j < k; j++) m[r][j] -= f * m[c][j];
    }
  }
  return num(det);
});

def("MINVERSE", 1, 1, ([a], ctx) => {
  const m = square(a, ctx);
  if (isErr(m)) return m;
  const k = m.length;
  const scale = Math.max(...m.flat().map(Math.abs));
  const inv = m.map((_, i) => m.map((__, j) => (i === j ? 1 : 0)));
  for (let c = 0; c < k; c++) {
    let p = c;
    for (let r = c + 1; r < k; r++) if (Math.abs(m[r][c]) > Math.abs(m[p][c])) p = r;
    if (Math.abs(m[p][c]) <= 1e-15 * scale) return E.num;
    [m[p], m[c]] = [m[c], m[p]];
    [inv[p], inv[c]] = [inv[c], inv[p]];
    const d = m[c][c];
    for (let j = 0; j < k; j++) {
      m[c][j] /= d;
      inv[c][j] /= d;
    }
    for (let r = 0; r < k; r++) {
      if (r === c) continue;
      const f = m[r][c];
      if (f === 0) continue;
      for (let j = 0; j < k; j++) {
        m[r][j] -= f * m[c][j];
        inv[r][j] -= f * inv[c][j];
      }
    }
  }
  return new Arr(k, k, inv.flat().map((v) => num(v) as V));
});

def("MUNIT", 1, 1, ([a], ctx) => {
  const k = intArg(a, ctx, null);
  if (isErr(k)) return k;
  if (k === null || k < 1) return E.value;
  const d: V[] = [];
  for (let r = 0; r < k; r++) for (let c = 0; c < k; c++) d.push(r === c ? 1 : 0);
  return new Arr(k, k, d);
});

// ----- Shaping --------------------------------------------------------------------------------

function sub(a: Arr, r0: number, r1: number, c0: number, c1: number): Arr {
  const d: V[] = [];
  for (let r = r0; r < r1; r++) for (let c = c0; c < c1; c++) d.push(a.get(r, c));
  return new Arr(r1 - r0, c1 - c0, d);
}

def("EXPAND", 2, 4, ([a0, r0, c0, p0], ctx) => {
  const a = arrArg(a0, ctx);
  if (isErr(a)) return a;
  const rows = intArg(r0, ctx, a.rows), cols = intArg(c0, ctx, a.cols);
  if (isErr(rows)) return rows;
  if (isErr(cols)) return cols;
  if (rows === null || cols === null || rows < a.rows || cols < a.cols) return E.value;
  const pad: V = p0 === undefined ? E.na : (deref(p0, ctx) as V);
  const d: V[] = [];
  for (let r = 0; r < rows; r++) for (let c = 0; c < cols; c++) d.push(r < a.rows && c < a.cols ? a.get(r, c) : pad);
  return new Arr(rows, cols, d);
});

function stack(name: string, horizontal: boolean): void {
  def(name, 1, 254, (args, ctx) => {
    const arrs: Arr[] = [];
    for (const x of args) {
      const a = arrArg(x, ctx);
      if (isErr(a)) return a;
      arrs.push(a);
    }
    const rows = horizontal ? Math.max(...arrs.map((a) => a.rows)) : arrs.reduce((s, a) => s + a.rows, 0);
    const cols = horizontal ? arrs.reduce((s, a) => s + a.cols, 0) : Math.max(...arrs.map((a) => a.cols));
    const d: V[] = new Array(rows * cols).fill(E.na);
    let off = 0;
    for (const a of arrs) {
      for (let r = 0; r < a.rows; r++) {
        for (let c = 0; c < a.cols; c++) {
          const rr = horizontal ? r : off + r, cc = horizontal ? off + c : c;
          d[rr * cols + cc] = a.get(r, c);
        }
      }
      off += horizontal ? a.cols : a.rows;
    }
    return new Arr(rows, cols, d);
  });
}
stack("HSTACK", true);
stack("VSTACK", false);

function wrap(name: string, byCol: boolean): void {
  def(name, 2, 3, ([a0, n0, p0], ctx) => {
    const a = arrArg(a0, ctx);
    if (isErr(a)) return a;
    if (a.rows !== 1 && a.cols !== 1) return E.value;
    const k = intArg(n0, ctx, null);
    if (isErr(k)) return k;
    if (k === null || k < 1) return E.num;
    const pad: V = p0 === undefined ? E.na : (deref(p0, ctx) as V);
    const len = a.d.length, inner = Math.min(k, len), outer = Math.ceil(len / k);
    const rows = byCol ? inner : outer, cols = byCol ? outer : inner;
    const d: V[] = new Array(rows * cols).fill(pad);
    for (let i = 0; i < len; i++) {
      const o = Math.floor(i / k), p = i % k;
      if (byCol) d[p * cols + o] = a.d[i];
      else d[o * cols + p] = a.d[i];
    }
    return new Arr(rows, cols, d);
  });
}
wrap("WRAPCOLS", true);
wrap("WRAPROWS", false);

function choose(name: string, cols: boolean): void {
  def(name, 2, 255, ([a0, ...rest], ctx) => {
    const a = arrArg(a0, ctx);
    if (isErr(a)) return a;
    const size = cols ? a.cols : a.rows;
    const idx: number[] = [];
    for (const x of rest) {
      const v = deref(x, ctx);
      for (const e of v instanceof Arr ? v.d : [v]) {
        const k = toNum(e);
        if (isErr(k)) return k;
        const i = Math.trunc(k);
        if (i === 0 || Math.abs(i) > size) return E.value;
        idx.push(i > 0 ? i - 1 : size + i);
      }
    }
    const d: V[] = [];
    if (cols) {
      for (let r = 0; r < a.rows; r++) for (const c of idx) d.push(a.get(r, c));
      return new Arr(a.rows, idx.length, d);
    }
    for (const r of idx) for (let c = 0; c < a.cols; c++) d.push(a.get(r, c));
    return new Arr(idx.length, a.cols, d);
  });
}
choose("CHOOSECOLS", true);
choose("CHOOSEROWS", false);

def("TAKE", 2, 3, ([a0, r0, c0], ctx) => {
  const a = arrArg(a0, ctx);
  if (isErr(a)) return a;
  const r = intArg(r0, ctx, null), c = intArg(c0, ctx, null);
  if (isErr(r)) return r;
  if (isErr(c)) return c;
  if (r === 0 || c === 0) return E.calc;
  const rr = r === null ? [0, a.rows] : r > 0 ? [0, Math.min(r, a.rows)] : [Math.max(0, a.rows + r), a.rows];
  const cc = c === null ? [0, a.cols] : c > 0 ? [0, Math.min(c, a.cols)] : [Math.max(0, a.cols + c), a.cols];
  return sub(a, rr[0], rr[1], cc[0], cc[1]);
});

def("DROP", 2, 3, ([a0, r0, c0], ctx) => {
  const a = arrArg(a0, ctx);
  if (isErr(a)) return a;
  const r = intArg(r0, ctx, 0), c = intArg(c0, ctx, 0);
  if (isErr(r)) return r;
  if (isErr(c)) return c;
  const R = r ?? 0, C = c ?? 0;
  const rr = R >= 0 ? [Math.min(R, a.rows), a.rows] : [0, Math.max(0, a.rows + R)];
  const cc = C >= 0 ? [Math.min(C, a.cols), a.cols] : [0, Math.max(0, a.cols + C)];
  if (rr[1] <= rr[0] || cc[1] <= cc[0]) return E.calc;
  return sub(a, rr[0], rr[1], cc[0], cc[1]);
});

function flatten(name: string, column: boolean): void {
  def(name, 1, 3, ([a0, i0, s0], ctx) => {
    const a = arrArg(a0, ctx);
    if (isErr(a)) return a;
    const ign = intArg(i0, ctx, 0);
    if (isErr(ign)) return ign;
    if (ign === null || ign < 0 || ign > 3) return E.value;
    let byCol = false;
    if (s0 !== undefined && s0 !== null) {
      const v = deref(s0, ctx);
      const b = toBool(v instanceof Arr ? v.d[0] : v);
      if (isErr(b)) return b;
      byCol = b;
    }
    const out: V[] = [];
    const take = (v: V) => {
      if ((ign & 1) && v === null) return;
      if ((ign & 2) && isErr(v)) return;
      out.push(v);
    };
    if (byCol) for (let c = 0; c < a.cols; c++) for (let r = 0; r < a.rows; r++) take(a.get(r, c));
    else for (const v of a.d) take(v);
    if (!out.length) return E.calc;
    return column ? new Arr(out.length, 1, out) : new Arr(1, out.length, out);
  });
}
flatten("TOCOL", true);
flatten("TOROW", false);

def("TRIMRANGE", 1, 3, ([a0, r0, c0], ctx) => {
  if (!(a0 instanceof Ref)) return isErr(a0) ? a0 : E.value;
  if (a0.areas.length !== 1) return E.value;
  const tr = intArg(r0, ctx, 3), tc = intArg(c0, ctx, 3);
  if (isErr(tr)) return tr;
  if (isErr(tc)) return tc;
  if (tr === null || tc === null || tr < 0 || tr > 3 || tc < 0 || tc > 3) return E.value;
  const area = a0.areas[0];
  const arr = ctx.refArr(a0);
  let top = -1, bottom = -1, left = -1, right = -1;
  for (let r = 0; r < arr.rows; r++) {
    for (let c = 0; c < arr.cols; c++) {
      if (arr.get(r, c) === null) continue;
      if (top < 0) top = r;
      bottom = r;
      if (left < 0 || c < left) left = c;
      if (c > right) right = c;
    }
  }
  if (top < 0) return E.ref;
  return new Ref(a0.sheet, [{
    r0: tr & 1 ? area.r0 + top : area.r0, r1: tr & 2 ? area.r0 + bottom : area.r1,
    c0: tc & 1 ? area.c0 + left : area.c0, c1: tc & 2 ? area.c0 + right : area.c1,
  }]);
});

// ----- GRUPPIERENNACH, PIVOTMIT ---------------------------------------------------------------

/** The function given by name (SUMME, MITTELWERT …, without brackets) to the grouping functions. */
function fnArg(node: Node | undefined): FnDef | null {
  if (!node || node.t !== "name") return null;
  const up = node.name.toUpperCase();
  return FUNCS.get(up) ?? FUNCS.get(englishName(up) ?? "") ?? null;
}

function rowsOf(a: Arr): V[][] {
  const out: V[][] = [];
  for (let r = 0; r < a.rows; r++) out.push(a.d.slice(r * a.cols, (r + 1) * a.cols));
  return out;
}

function keyCmp(a: V[], b: V[]): number {
  for (let i = 0; i < a.length; i++) {
    const c = compare(a[i] ?? "", b[i] ?? "");
    if (c) return c;
  }
  return 0;
}

const keyOf = (k: V[]) => JSON.stringify(k.map((v) => (isErr(v) ? { e: v.code } : v)));

function apply(fn: FnDef, vals: V[], ctx: Ctx): V {
  const r = fn.fn([new Arr(vals.length, 1, vals)], ctx, []);
  const v = r instanceof Arr ? r.d[0] : r instanceof Ref ? E.value : r;
  return v === null ? 0 : v;
}

interface Prepared { keys: V[][]; vals: V[][]; head: { keys: V[]; vals: V[] } | null }

/** Rows of the fields and values after the headers and the filter are applied. */
function prepare(fields: Arr, vals: Arr, headers: number | null, filter: Arr | null): Prepared | CellError {
  if (fields.rows !== vals.rows) return E.value;
  let keys = rowsOf(fields), vs = rowsOf(vals);
  let hdr = headers;
  if (hdr === null) hdr = vs.length > 1 && typeof vs[0][0] === "string" && typeof vs[1][0] === "number" ? 3 : 0;
  let head: Prepared["head"] = null;
  let keep = filter ? filter.d.map((v) => {
    const b = toBool(v);
    return !isErr(b) && b;
  }) : null;
  if (hdr === 1 || hdr === 3) {
    if (hdr === 3) head = { keys: keys[0], vals: vs[0] };
    keys = keys.slice(1);
    vs = vs.slice(1);
    keep = keep ? keep.slice(1) : null;
  }
  if (keep) {
    keys = keys.filter((_, i) => keep![i]);
    vs = vs.filter((_, i) => keep![i]);
  }
  return { keys, vals: vs, head };
}

function optArr(a: Any | undefined, ctx: Ctx): Arr | null | CellError {
  if (a === undefined || a === null) return null;
  return arrArg(a, ctx);
}

def("GROUPBY", 3, 8, (args, ctx, nodes) => {
  const fields = arrArg(args[0], ctx), vals = arrArg(args[1], ctx);
  if (isErr(fields)) return fields;
  if (isErr(vals)) return vals;
  const fn = fnArg(nodes[2]);
  if (!fn) return E.value;
  const headers = intArg(args[3], ctx, null), depth = intArg(args[4], ctx, 1), sort = intArg(args[5], ctx, 1);
  for (const x of [headers, depth, sort]) if (isErr(x)) return x;
  const filter = optArr(args[6], ctx);
  if (isErr(filter)) return filter;
  const p = prepare(fields, vals, headers as number | null, filter);
  if (isErr(p)) return p;
  const kr = fields.cols, kv = vals.cols;
  const groups = new Map<string, { key: V[]; rows: number[] }>();
  p.keys.forEach((k, i) => {
    const id = keyOf(k);
    const g = groups.get(id) ?? { key: k, rows: [] };
    g.rows.push(i);
    groups.set(id, g);
  });
  const agg = (rows: number[]) => {
    const out: V[] = [];
    for (let j = 0; j < kv; j++) out.push(apply(fn, rows.map((i) => p.vals[i][j]), ctx));
    return out;
  };
  let lines = [...groups.values()].map((g) => ({ key: g.key, vals: agg(g.rows), rows: g.rows }));
  const s = (sort as number | null) ?? 1;
  const col = Math.abs(s) - 1;
  if (col < 0 || col >= kr + kv) return E.value;
  const dir = s < 0 ? -1 : 1;
  lines.sort((a, b) => {
    const c = col < kr ? compare(a.key[col] ?? "", b.key[col] ?? "") : compare(a.vals[col - kr], b.vals[col - kr]);
    return c ? dir * c : keyCmp(a.key, b.key);
  });
  const d = (depth as number | null) ?? 1;
  const all = p.keys.map((_, i) => i);
  const out: V[][] = [];
  const subtotals = Math.abs(d) >= 2 && kr >= 2;
  const total = (label: string, rows: number[], level: V[] = []) => [...level, label, ...new Array(kr - level.length - 1).fill(""), ...agg(rows)];
  if (subtotals) {
    // Level-1 groups in their sorted order, each with its subtotal.
    const order: V[] = [];
    for (const l of lines) if (!order.some((v) => compare(v, l.key[0]) === 0)) order.push(l.key[0]);
    const block: V[][] = [];
    for (const k1 of order) {
      const members = lines.filter((l) => compare(l.key[0], k1) === 0);
      const rows = members.flatMap((m) => m.rows);
      const sub = [k1, ...new Array(kr - 1).fill(""), ...agg(rows)];
      if (d < 0) block.push(sub);
      for (const m of members) block.push([...m.key, ...m.vals]);
      if (d > 0) block.push(sub);
    }
    lines = [];
    if (d < 0) out.push(total("Gesamtergebnis", all));
    out.push(...block);
    if (d > 0) out.push(total("Gesamtergebnis", all));
  } else {
    if (d < 0) out.push(total("Gesamt", all));
    for (const l of lines) out.push([...l.key, ...l.vals]);
    if (d > 0) out.push(total("Gesamt", all));
  }
  if (p.head) out.unshift([...p.head.keys, ...p.head.vals]);
  if (!out.length) return E.calc;
  return Arr.of(out);
});

def("PIVOTBY", 4, 11, (args, ctx, nodes) => {
  const rf = arrArg(args[0], ctx), cf = arrArg(args[1], ctx), vals = arrArg(args[2], ctx);
  if (isErr(rf)) return rf;
  if (isErr(cf)) return cf;
  if (isErr(vals)) return vals;
  if (rf.rows !== vals.rows || cf.rows !== vals.rows) return E.value;
  const fn = fnArg(nodes[3]);
  if (!fn) return E.value;
  const headers = intArg(args[4], ctx, null), rDepth = intArg(args[5], ctx, 1), rSort = intArg(args[6], ctx, 1);
  const cDepth = intArg(args[7], ctx, 1), cSort = intArg(args[8], ctx, 1);
  for (const x of [headers, rDepth, rSort, cDepth, cSort]) if (isErr(x)) return x;
  const filter = optArr(args[9], ctx);
  if (isErr(filter)) return filter;
  const both = new Arr(rf.rows, rf.cols + cf.cols, rowsOf(rf).flatMap((r, i) => [...r, ...rowsOf(cf)[i]]));
  const p = prepare(both, vals, headers as number | null, filter);
  if (isErr(p)) return p;
  const kr = rf.cols;
  const rowKey = (i: number) => p.keys[i].slice(0, kr), colKey = (i: number) => p.keys[i].slice(kr);
  const distinct = (f: (i: number) => V[], dir: number) => {
    const m = new Map<string, V[]>();
    p.keys.forEach((_, i) => m.set(keyOf(f(i)), f(i)));
    return [...m.values()].sort((a, b) => dir * keyCmp(a, b));
  };
  const rows = distinct(rowKey, (rSort as number) < 0 ? -1 : 1), cols = distinct(colKey, (cSort as number) < 0 ? -1 : 1);
  const all = p.keys.map((_, i) => i);
  const cell = (pick: number[]): V => (pick.length ? apply(fn, pick.map((i) => p.vals[i][0]), ctx) : "");
  const where = (rk: V[] | null, ck: V[] | null) => all.filter((i) => (!rk || keyCmp(rowKey(i), rk) === 0) && (!ck || keyCmp(colKey(i), ck) === 0));
  const withColTotal = (cDepth as number) !== 0, withRowTotal = (rDepth as number) !== 0;
  const out: V[][] = [];
  const hdr: V[] = [...new Array(kr).fill(""), ...cols.map((c) => c[0])];
  if (withColTotal) hdr.push("Gesamt");
  out.push(hdr);
  const body: V[][] = rows.map((rk) => {
    const line: V[] = [...rk, ...cols.map((ck) => cell(where(rk, ck)))];
    if (withColTotal) line.push(cell(where(rk, null)));
    return line;
  });
  const totalLine: V[] = ["Gesamt", ...new Array(kr - 1).fill(""), ...cols.map((ck) => cell(where(null, ck)))];
  if (withColTotal) totalLine.push(cell(all));
  if (withRowTotal && (rDepth as number) < 0) out.push(totalLine);
  out.push(...body);
  if (withRowTotal && (rDepth as number) > 0) out.push(totalLine);
  return Arr.of(out);
});

