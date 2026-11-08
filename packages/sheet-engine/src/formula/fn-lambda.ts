// LET, LAMBDA and the functions that call a LAMBDA (D-080): MAP, REDUCE, SCAN, NACHZEILE (BYROW),
// NACHSPALTE (BYCOL), MATRIXERSTELLEN (MAKEARRAY), WURDEAUSGELASSEN (ISOMITTED).
//
// A LAMBDA is a value while a formula is worked out; a cell cannot show one, so it is an error value
// (#CALC!, as Excel shows a formula whose result is a LAMBDA) that carries its parameters, its body
// and the names bound where it was made. Functions that get one where they expect a value pass the
// error on, as Excel does.

import { CellError } from "../model";
import { Ctx, def, deref, setCallValue } from "./eval";
import type { Node } from "./parse";
import { Arr, E, isErr, toNum, type Any, type V } from "./values";

export class Lambda extends CellError {
  constructor(readonly params: string[], readonly body: Node, readonly scope: Map<string, Any> | null) {
    super("#CALC!");
  }
}

/** The mark of a parameter left out in a call (WURDEAUSGELASSEN). */
const OMITTED = "\u0000omitted:";
let depth = 0;

/** Calls a LAMBDA with values (a missing argument counts as left out). */
function invoke(f: Lambda, args: (Any | undefined)[], ctx: Ctx): Any {
  // Every parameter needs its argument; an empty one (f(1;)) counts as left out (recorded).
  if (args.length !== f.params.length) return E.value;
  const scope = new Map(f.scope ?? []);
  f.params.forEach((p, i) => {
    const v = args[i];
    if (v === undefined) {
      scope.set(p, null);
      scope.set(OMITTED + p, true);
    } else {
      scope.set(p, v);
      scope.delete(OMITTED + p);
    }
  });
  if (depth > 1000) return E.num;
  depth++;
  try {
    return ctx.withScope(scope).eval(f.body);
  } finally {
    depth--;
  }
}

setCallValue((f, argNodes, ctx) => {
  if (!(f instanceof Lambda)) return f instanceof CellError ? f : E.value;
  return invoke(f, argNodes.map((n) => (n.t === "miss" ? undefined : ctx.eval(n))), ctx);
});

const nameOf = (n: Node): string | null => (n.t === "name" && n.sheet === undefined ? n.name.toLowerCase() : null);

def("LET", 3, 253, (_a, ctx, nodes) => {
  if (nodes.length % 2 === 0) return E.value;
  let c = ctx;
  for (let i = 0; i + 1 < nodes.length; i += 2) {
    const name = nameOf(nodes[i]);
    if (name === null) return E.value;
    const v = c.eval(nodes[i + 1]);
    const scope = new Map(c.scope ?? []);
    scope.set(name, v);
    c = ctx.withScope(scope);
  }
  return c.eval(nodes[nodes.length - 1]);
}, { lazy: true });

def("LAMBDA", 1, 254, (_a, ctx, nodes) => {
  const params: string[] = [];
  for (const n of nodes.slice(0, -1)) {
    const name = nameOf(n);
    if (name === null || params.includes(name)) return E.value;
    params.push(name);
  }
  return new Lambda(params, nodes[nodes.length - 1], ctx.scope);
}, { lazy: true });

def("ISOMITTED", 1, 1, (_a, ctx, nodes) => {
  const name = nameOf(nodes[0]);
  return name !== null && !!ctx.scope?.get(OMITTED + name);
}, { lazy: true });

/** An argument as an array (a single value as a 1×1 array). */
function arr(a: Any, ctx: Ctx): Arr | CellError {
  const d = deref(a, ctx);
  if (d instanceof Arr) return d;
  if (d instanceof Lambda) return E.value;
  return new Arr(1, 1, [d]);
}

/** A LAMBDA's result for one element: a single value (an array of one), else #CALC!. */
function one(r: Any, ctx: Ctx): V {
  const d = deref(r, ctx);
  if (d instanceof Arr) return d.rows === 1 && d.cols === 1 ? d.d[0] : E.calc;
  return d instanceof Lambda ? E.calc : d;
}

function lambdaArg(a: Any, count: number): Lambda | CellError {
  if (!(a instanceof Lambda)) return isErr(a) ? a : E.value;
  return a.params.length < count ? E.value : a;
}

def("MAP", 2, 254, (args, ctx) => {
  const f = lambdaArg(args[args.length - 1], args.length - 1);
  if (!(f instanceof Lambda)) return f;
  const arrays: Arr[] = [];
  for (const a of args.slice(0, -1)) {
    const x = arr(a, ctx);
    if (x instanceof CellError) return x;
    arrays.push(x);
  }
  const rows = Math.max(...arrays.map((a) => a.rows)), cols = Math.max(...arrays.map((a) => a.cols));
  const d: V[] = [];
  for (let r = 0; r < rows; r++) for (let c = 0; c < cols; c++) {
    if (arrays.some((a) => (a.rows !== 1 && r >= a.rows) || (a.cols !== 1 && c >= a.cols))) { d.push(E.na); continue; }
    d.push(one(invoke(f, arrays.map((a) => a.get(a.rows === 1 ? 0 : r, a.cols === 1 ? 0 : c)), ctx), ctx));
  }
  return new Arr(rows, cols, d);
});

function fold(args: Any[], ctx: Ctx, keep: boolean): Any {
  const [init, array, fn] = args.length === 3 ? args : [null, args[0], args[1]];
  const f = lambdaArg(fn, 2);
  if (!(f instanceof Lambda)) return f;
  const a = arr(array, ctx);
  if (a instanceof CellError) return a;
  let acc: Any = deref(init, ctx);
  if (acc instanceof Arr) return E.value;
  const out: V[] = [];
  for (const v of a.d) {
    acc = invoke(f, [acc, v], ctx);
    if (keep) out.push(one(acc, ctx));
  }
  return keep ? new Arr(a.rows, a.cols, out) : acc;
}

def("REDUCE", 2, 3, (args, ctx) => fold(args, ctx, false));
def("SCAN", 2, 3, (args, ctx) => fold(args, ctx, true));

def("BYROW", 2, 2, (args, ctx) => {
  const f = lambdaArg(args[1], 1);
  if (!(f instanceof Lambda)) return f;
  const a = arr(args[0], ctx);
  if (a instanceof CellError) return a;
  const d: V[] = [];
  for (let r = 0; r < a.rows; r++) d.push(one(invoke(f, [new Arr(1, a.cols, a.d.slice(r * a.cols, (r + 1) * a.cols))], ctx), ctx));
  return new Arr(a.rows, 1, d);
});

def("BYCOL", 2, 2, (args, ctx) => {
  const f = lambdaArg(args[1], 1);
  if (!(f instanceof Lambda)) return f;
  const a = arr(args[0], ctx);
  if (a instanceof CellError) return a;
  const d: V[] = [];
  for (let c = 0; c < a.cols; c++) {
    const col: V[] = [];
    for (let r = 0; r < a.rows; r++) col.push(a.get(r, c));
    d.push(one(invoke(f, [new Arr(a.rows, 1, col)], ctx), ctx));
  }
  return new Arr(1, a.cols, d);
});

def("MAKEARRAY", 3, 3, (args, ctx) => {
  const f = lambdaArg(args[2], 2);
  if (!(f instanceof Lambda)) return f;
  const rows = Math.trunc(toNum(deref(args[0], ctx) as V) as number), cols = Math.trunc(toNum(deref(args[1], ctx) as V) as number);
  if (!Number.isFinite(rows) || !Number.isFinite(cols) || rows < 1 || cols < 1) return E.value;
  if (rows * cols > 1 << 20) return E.num;
  const d: V[] = [];
  for (let r = 1; r <= rows; r++) for (let c = 1; c <= cols; c++) d.push(one(invoke(f, [r, c], ctx), ctx));
  return new Arr(rows, cols, d);
});


