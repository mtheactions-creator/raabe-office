// Evaluation of a formula's syntax tree with Excel 365's semantics: dynamic arrays (an
// operator or a scalar function given a range or an array works element by element and the
// result spills), "@" for the old implicit intersection, errors passed on from the left.

import { MAX_COLS, MAX_ROWS, type Rect } from "../address";
import { CellError, type Workbook } from "../model";
import { parseFormula, type Node } from "./parse";
import { Arr, E, Ref, compare, isErr, num, toNum, toStr, type Any, type V } from "./values";

/** What the evaluator needs from its surroundings (a Calc, or the workbook's stored values). */
export interface Host {
  wb: Workbook;
  /** The current value of a cell (computing it first if it is a formula that is out of date). */
  value(sheet: number, row: number, col: number): V;
  /** The rectangle holding every cell of a sheet with content, or null for an empty sheet. */
  bounds(sheet: number): Rect | null;
  /** Serial of the current date and time (fixed during one recalculation). */
  now(): number;
  random(): number;
  /** The formula of a cell (file form) or undefined. */
  formula(sheet: number, row: number, col: number): string | undefined;
  /** The range a dynamic array formula in this cell spills into, or null. */
  spill(sheet: number, row: number, col: number): Rect | null;
}

export type Fn = (args: Any[], ctx: Ctx, nodes: Node[]) => Any;

export interface FnDef {
  fn: Fn;
  min: number;
  max: number;
  /** The function evaluates its own arguments (WENN, WENNFEHLER, WAHL …): args is empty. */
  lazy?: boolean;
  /** Recomputed on every recalculation (JETZT, ZUFALLSZAHL, INDIREKT …). */
  volatile?: boolean;
}

export const FUNCS = new Map<string, FnDef>();

export function def(name: string, min: number, max: number, fn: Fn, opts: { lazy?: boolean; volatile?: boolean } = {}): void {
  FUNCS.set(name, { fn, min, max, ...opts });
}

/** Where a formula is evaluated. */
export class Ctx {
  constructor(readonly host: Host, readonly sheet: number, readonly row: number, readonly col: number) {}

  eval(n: Node): Any {
    return evalNode(n, this);
  }

  /** The values of a reference as an array (huge whole-column areas trimmed to the used rows). */
  refArr(ref: Ref, area = 0): Arr {
    const a = this.clip(ref.sheet, ref.areas[area]);
    const rows = a.r1 - a.r0 + 1, cols = a.c1 - a.c0 + 1;
    const d: V[] = new Array(Math.max(0, rows * cols)).fill(null);
    if (rows <= 0 || cols <= 0) return new Arr(0, 0, []);
    const b = this.host.bounds(ref.sheet);
    if (b) {
      const r0 = Math.max(a.r0, b.r0), r1 = Math.min(a.r1, b.r1), c0 = Math.max(a.c0, b.c0), c1 = Math.min(a.c1, b.c1);
      for (let r = r0; r <= r1; r++) for (let c = c0; c <= c1; c++) d[(r - a.r0) * cols + (c - a.c0)] = this.host.value(ref.sheet, r, c);
    }
    return new Arr(rows, cols, d);
  }

  /** Whole columns or rows reach to the sheet's edge; arrays of them stop after the used range. */
  clip(sheet: number, a: Rect): Rect {
    const cells = (a.r1 - a.r0 + 1) * (a.c1 - a.c0 + 1);
    if (cells <= 1 << 17) return a;
    const b = this.host.bounds(sheet);
    if (!b) return { r0: a.r0, c0: a.c0, r1: a.r0, c1: a.c0 };
    return { r0: a.r0, c0: a.c0, r1: Math.max(a.r0, Math.min(a.r1, b.r1)), c1: Math.max(a.c0, Math.min(a.c1, b.c1)) };
  }

  /** Calls `f` for every non-empty cell of the areas (row by row; sheet by sheet for 3D) with its value. */
  eachCell(ref: Ref, f: (v: V, r: number, c: number) => void): void {
    for (let s = ref.sheet; s <= (ref.sheetEnd ?? ref.sheet); s++) {
      const b = this.host.bounds(s);
      if (!b) continue;
      for (const a of ref.areas) {
        const r0 = Math.max(a.r0, b.r0), r1 = Math.min(a.r1, b.r1), c0 = Math.max(a.c0, b.c0), c1 = Math.min(a.c1, b.c1);
        for (let r = r0; r <= r1; r++) {
          for (let c = c0; c <= c1; c++) {
            const v = this.host.value(s, r, c);
            if (v !== null) f(v, r, c);
          }
        }
      }
    }
  }

  sheetIndex(name: string): number {
    const n = name.toLowerCase();
    return this.host.wb.sheets.findIndex((s) => s.name.toLowerCase() === n);
  }
}

/** The value of a reference with a single cell, or the array of a larger one. */
export function deref(a: Any, ctx: Ctx): V | Arr {
  if (a instanceof Ref) {
    if (a.areas.length > 1 || a.sheetEnd !== undefined) return E.value;
    const r = a.areas[0];
    if (r.r0 === r.r1 && r.c0 === r.c1) return ctx.host.value(a.sheet, r.r0, r.c0);
    return ctx.refArr(a);
  }
  return a;
}

/** A single value: arrays give their first element (callers lift over arrays before this). */
export function scalar(a: Any, ctx: Ctx): V {
  const d = deref(a, ctx);
  if (d instanceof Arr) return d.rows && d.cols ? d.d[0] : E.value;
  return d;
}

/** Applies `f` to every element (or to the value), keeping the shape. */
export function map1(a: V | Arr, f: (v: V) => V): V | Arr {
  if (a instanceof Arr) return new Arr(a.rows, a.cols, a.d.map(f));
  return f(a);
}

/** Excel's broadcasting: a single row or column stretches, a too-short array gives #NV. */
export function map2(a: V | Arr, b: V | Arr, f: (x: V, y: V) => V): V | Arr {
  if (!(a instanceof Arr) && !(b instanceof Arr)) return f(a, b);
  const ra = a instanceof Arr ? a.rows : 1, ca = a instanceof Arr ? a.cols : 1;
  const rb = b instanceof Arr ? b.rows : 1, cb = b instanceof Arr ? b.cols : 1;
  const rows = Math.max(ra, rb), cols = Math.max(ca, cb);
  const d: V[] = new Array(rows * cols);
  for (let r = 0; r < rows; r++) {
    for (let c = 0; c < cols; c++) {
      const x = pick(a, r, c), y = pick(b, r, c);
      d[r * cols + c] = f(x, y);
    }
  }
  return new Arr(rows, cols, d);
}

/** Element (r, c) of a value broadcast to a larger shape. */
export function pick(a: V | Arr, r: number, c: number): V {
  if (!(a instanceof Arr)) return a;
  const rr = a.rows === 1 ? 0 : r, cc = a.cols === 1 ? 0 : c;
  if (rr >= a.rows || cc >= a.cols) return E.na;
  return a.d[rr * a.cols + cc];
}

/** Applies a scalar function over any number of arguments, lifting over arrays. */
export function mapN(args: (V | Arr)[], f: (vals: V[]) => V | Arr): V | Arr {
  let rows = 1, cols = 1, any = false;
  for (const a of args) {
    if (a instanceof Arr) {
      any = true;
      rows = Math.max(rows, a.rows);
      cols = Math.max(cols, a.cols);
    }
  }
  if (!any) return f(args as V[]);
  const d: V[] = new Array(rows * cols);
  for (let r = 0; r < rows; r++) {
    for (let c = 0; c < cols; c++) {
      const vals = args.map((a) => pick(a, r, c));
      const out = f(vals);
      d[r * cols + c] = out instanceof Arr ? out.d[0] ?? null : out;
    }
  }
  return new Arr(rows, cols, d);
}

// ----- Operators ---------------------------------------------------------------------------

/** Excel sets a sum or difference that is zero to 15 digits to exactly 0 when it is the
 *  formula's last operation (=0,1+0,2-0,3 is 0, =(0,1+0,2-0,3) is not). */
export function snap(r: number, a: number, b: number): number {
  if (r !== 0 && Math.abs(r) <= Math.max(Math.abs(a), Math.abs(b)) * 2 ** -49) {
    return 0;
  }
  return r;
}

function arith(op: string, x: V, y: V, top: boolean): V {
  if (isErr(x)) return x;
  const a = toNum(x);
  if (isErr(a)) return a;
  if (isErr(y)) return y;
  const b = toNum(y);
  if (isErr(b)) return b;
  switch (op) {
    case "+": return num(top ? snap(a + b, a, b) : a + b);
    case "-": return num(top ? snap(a - b, a, b) : a - b);
    case "*": return num(a * b);
    case "/": return b === 0 ? E.div0 : num(a / b);
    default: {
      if (a === 0 && b === 0) return E.num;
      if (a === 0 && b < 0) return E.div0;
      return num(power(a, b));
    }
  }
}

/** a^b with Excel's odd roots of negative numbers (=(-8)^(1/3) is -2). */
export function power(a: number, b: number): number {
  if (a < 0 && !Number.isInteger(b)) {
    const inv = 1 / b;
    if (Number.isInteger(Math.round(inv)) && Math.abs(inv - Math.round(inv)) < 1e-10 && Math.round(inv) % 2 !== 0) return -Math.pow(-a, b);
    return NaN;
  }
  return Math.pow(a, b);
}

function cmp(op: string, x: V, y: V): V {
  if (isErr(x)) return x;
  if (isErr(y)) return y;
  const c = compare(x, y);
  switch (op) {
    case "=": return c === 0;
    case "<>": return c !== 0;
    case "<": return c < 0;
    case ">": return c > 0;
    case "<=": return c <= 0;
    default: return c >= 0;
  }
}

function concat(x: V, y: V): V {
  if (isErr(x)) return x;
  if (isErr(y)) return y;
  const a = toStr(x), b = toStr(y);
  if (isErr(a)) return a;
  if (isErr(b)) return b;
  return a + b;
}

function sameSheet(a: Ref, b: Ref): boolean {
  return a.sheet === b.sheet;
}

function bbox(areas: Rect[]): Rect {
  return {
    r0: Math.min(...areas.map((a) => a.r0)), c0: Math.min(...areas.map((a) => a.c0)),
    r1: Math.max(...areas.map((a) => a.r1)), c1: Math.max(...areas.map((a) => a.c1)),
  };
}

function refOp(op: string, a: Any, b: Any): Any {
  if (isErr(a)) return a;
  if (isErr(b)) return b;
  if (!(a instanceof Ref) || !(b instanceof Ref)) return E.value;
  if (!sameSheet(a, b)) return E.value;
  if (op === ":") return new Ref(a.sheet, [bbox([...a.areas, ...b.areas])]);
  if (op === ",") return new Ref(a.sheet, [...a.areas, ...b.areas]);
  const out: Rect[] = [];
  for (const x of a.areas) {
    for (const y of b.areas) {
      const r = { r0: Math.max(x.r0, y.r0), c0: Math.max(x.c0, y.c0), r1: Math.min(x.r1, y.r1), c1: Math.min(x.c1, y.c1) };
      if (r.r0 <= r.r1 && r.c0 <= r.c1) out.push(r);
    }
  }
  return out.length ? new Ref(a.sheet, out) : E.null;
}

/** Excel's implicit intersection: the cell of a column in the formula's row, or of a row in its column. */
export function intersect(a: Any, ctx: Ctx): Any {
  if (a instanceof Arr) return a.rows && a.cols ? a.d[0] : E.value;
  if (!(a instanceof Ref)) return a;
  if (a.areas.length !== 1) return E.value;
  const r = a.areas[0];
  if (r.r0 === r.r1 && r.c0 === r.c1) return a;
  if (r.c0 === r.c1 && ctx.row >= r.r0 && ctx.row <= r.r1 && ctx.sheet === a.sheet) return new Ref(a.sheet, [{ r0: ctx.row, c0: r.c0, r1: ctx.row, c1: r.c0 }]);
  if (r.r0 === r.r1 && ctx.col >= r.c0 && ctx.col <= r.c1 && ctx.sheet === a.sheet) return new Ref(a.sheet, [{ r0: r.r0, c0: ctx.col, r1: r.r0, c1: ctx.col }]);
  if (r.c0 === r.c1 && ctx.row >= r.r0 && ctx.row <= r.r1) return new Ref(a.sheet, [{ r0: ctx.row, c0: r.c0, r1: ctx.row, c1: r.c0 }]);
  if (r.r0 === r.r1 && ctx.col >= r.c0 && ctx.col <= r.c1) return new Ref(a.sheet, [{ r0: r.r0, c0: ctx.col, r1: r.r0, c1: ctx.col }]);
  return E.value;
}

// ----- Names -------------------------------------------------------------------------------

function evalName(name: string, sheetName: string | undefined, ctx: Ctx): Any {
  const wb = ctx.host.wb;
  const lname = name.toLowerCase();
  const home = sheetName !== undefined ? ctx.sheetIndex(sheetName) : ctx.sheet;
  const dn = wb.names.find((n) => n.name.toLowerCase() === lname && n.sheet === home) ?? (sheetName === undefined ? wb.names.find((n) => n.name.toLowerCase() === lname && n.sheet === undefined) : undefined);
  if (!dn) return E.name;
  let node: Node;
  try {
    node = parseFormula(dn.ref.startsWith("=") ? dn.ref.slice(1) : dn.ref);
  } catch {
    return E.name;
  }
  return evalNode(node, ctx);
}

// ----- The tree walk -----------------------------------------------------------------------

export function evalNode(n: Node, ctx: Ctx, top = false): Any {
  switch (n.t) {
    case "num": return n.v;
    case "str": return n.v;
    case "bool": return n.v;
    case "err": return E[codeKey(n.v)] ?? new CellError(n.v);
    case "miss": return null;
    case "struct": return E.ref;
    case "arr": return new Arr(n.rows, n.cols, n.v);
    case "paren": return evalNode(n.a, ctx);
    case "ref": {
      const s = n.sheet === undefined ? ctx.sheet : ctx.sheetIndex(n.sheet);
      if (s < 0) return E.ref;
      const rect = { r0: n.r0, c0: n.c0, r1: n.r1, c1: n.c1 };
      if (n.sheet2 !== undefined) {
        const s2 = ctx.sheetIndex(n.sheet2);
        if (s2 < 0) return E.ref;
        return new Ref(Math.min(s, s2), [rect], Math.max(s, s2));
      }
      return new Ref(s, [rect]);
    }
    case "name": return evalName(n.name, n.sheet, ctx);
    case "fn": return callFn(n.name, n.args, ctx);
    case "un": {
      if (n.op === "@") return intersect(evalNode(n.a, ctx), ctx);
      if (n.op === "#") {
        const a = evalNode(n.a, ctx);
        if (!(a instanceof Ref) || !a.single) return E.ref;
        const r = a.areas[0];
        const sp = ctx.host.spill(a.sheet, r.r0, r.c0);
        return sp ? new Ref(a.sheet, [sp]) : E.ref;
      }
      const a = evalNode(n.a, ctx);
      if (n.op === "+") return a;
      const d = deref(a, ctx);
      if (n.op === "-") return map1(d, (v) => (isErr(v) ? v : neg(toNum(v))));
      return map1(d, (v) => (isErr(v) ? v : pct(toNum(v))));
    }
    case "bin": {
      if (n.op === ":" || n.op === " " || n.op === ",") return refOp(n.op, evalNode(n.a, ctx), evalNode(n.b, ctx));
      const a = deref(evalNode(n.a, ctx), ctx);
      const b = deref(evalNode(n.b, ctx), ctx);
      switch (n.op) {
        case "+": case "-": case "*": case "/": case "^":
          return map2(a, b, (x, y) => arith(n.op, x, y, top));
        case "&":
          return map2(a, b, concat);
        default:
          return map2(a, b, (x, y) => cmp(n.op, x, y));
      }
    }
  }
}

function neg(v: number | CellError): V {
  return isErr(v) ? v : v === 0 ? 0 : -v;
}

function pct(v: number | CellError): V {
  return isErr(v) ? v : v / 100;
}

const CODE_KEYS: Record<string, keyof typeof E> = {
  "#VALUE!": "value", "#DIV/0!": "div0", "#N/A": "na", "#NUM!": "num", "#REF!": "ref", "#NAME?": "name", "#NULL!": "null", "#SPILL!": "spill", "#CALC!": "calc",
};

function codeKey(code: string): keyof typeof E {
  return CODE_KEYS[code];
}

function callFn(name: string, argNodes: Node[], ctx: Ctx): Any {
  const d = FUNCS.get(name.toUpperCase());
  if (!d) return E.name;
  if (argNodes.length < d.min || argNodes.length > d.max) return E.value;
  if (d.lazy) return d.fn([], ctx, argNodes);
  const args = argNodes.map((a) => evalNode(a, ctx));
  return d.fn(args, ctx, argNodes);
}

/** The result of a whole formula: a value, or an array when it spills. */
export function evalFormula(n: Node, ctx: Ctx): V | Arr {
  let r: Any;
  try {
    r = evalNode(n, ctx, n.t === "bin");
  } catch (e) {
    if (e instanceof RangeError) return E.num;
    throw e;
  }
  const d = r instanceof Ref ? (r.areas.length > 1 ? E.value : derefFull(r, ctx)) : r;
  if (d instanceof Arr) {
    if (d.rows === 1 && d.cols === 1) return d.d[0] ?? 0;
    if (d.rows === 0 || d.cols === 0) return E.calc;
    // An array that does not fit between the cell and the sheet's edges cannot spill.
    if (ctx.row + d.rows > MAX_ROWS || ctx.col + d.cols > MAX_COLS) return E.spill;
    return d;
  }
  return d;
}

function derefFull(r: Ref, ctx: Ctx): V | Arr {
  const a = r.areas[0];
  if (a.r0 === a.r1 && a.c0 === a.c1) return ctx.host.value(r.sheet, a.r0, a.c0);
  // A whole column or row cannot spill on a sheet of the same size.
  if (a.r1 - a.r0 + 1 >= MAX_ROWS || a.c1 - a.c0 + 1 >= MAX_COLS) return E.spill;
  return ctx.refArr(r);
}

export { num };
