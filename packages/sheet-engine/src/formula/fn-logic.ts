// Logical and information functions (Logik, Information), and the reference functions
// ZEILE, SPALTE, INDIREKT, BEREICH.VERSCHIEBEN, ADRESSE.

import { MAX_COLS, MAX_ROWS, colName } from "../address";
import { CellError } from "../model";
import { Ctx, def, deref, evalNode, map1, mapN, scalar } from "./eval";
import { b, defS, n, values } from "./fn-util";
import { lex } from "./lexer";
import { parseFormula, refRect, type Node } from "./parse";
import { Arr, E, Ref, compare, isErr, toBool, toNum, type Any, type V } from "./values";
import { sheetPrefix } from "./convert";

/** A condition's truth values (lifted over arrays). */
function truth(v: V): boolean | CellError {
  if (typeof v === "string") {
    const u = v.toUpperCase();
    if (u === "WAHR") return true;
    if (u === "FALSCH") return false;
    return E.value;
  }
  return toBool(v);
}

// WENN takes and returns references unchanged, so =SUMME(WENN(WAHR;A1:A3)) works.
def("IF", 1, 3, (_a, ctx, nodes) => {
  const c = deref(evalNode(nodes[0], ctx), ctx);
  const branch = (i: number): Any => (i < nodes.length ? evalNode(nodes[i], ctx) : i === 2 ? false : 0);
  if (!(c instanceof Arr)) {
    if (isErr(c)) return c;
    const t = truth(c);
    if (isErr(t)) return t;
    const r = t ? branch(1) : branch(2);
    return r === null ? 0 : r;
  }
  const yes = deref(branch(1), ctx), no = deref(branch(2), ctx);
  return mapN([c, yes, no], ([cv, y, nv]) => {
    if (isErr(cv)) return cv;
    const t = truth(cv);
    if (isErr(t)) return t;
    const r = t ? y : nv;
    return r === null ? 0 : r;
  });
}, { lazy: true });

def("IFS", 2, 254, (_a, ctx, nodes) => {
  if (nodes.length % 2) return E.value;
  for (let i = 0; i < nodes.length; i += 2) {
    const c = scalar(evalNode(nodes[i], ctx), ctx);
    if (isErr(c)) return c;
    const t = truth(c);
    if (isErr(t)) return t;
    if (t) {
      const r = evalNode(nodes[i + 1], ctx);
      return r === null ? 0 : r;
    }
  }
  return E.na;
}, { lazy: true });

def("IFERROR", 2, 2, (_a, ctx, nodes) => {
  const v = deref(evalNode(nodes[0], ctx), ctx);
  if (v instanceof Arr) {
    const alt = deref(evalNode(nodes[1], ctx), ctx);
    return mapN([v, alt], ([x, y]) => (isErr(x) ? (y === null ? 0 : y) : x));
  }
  if (isErr(v)) {
    const r = evalNode(nodes[1], ctx);
    return r === null ? 0 : r;
  }
  return v === null ? 0 : v;
}, { lazy: true });

def("IFNA", 2, 2, (_a, ctx, nodes) => {
  const v = deref(evalNode(nodes[0], ctx), ctx);
  if (v instanceof Arr) {
    const alt = deref(evalNode(nodes[1], ctx), ctx);
    return mapN([v, alt], ([x, y]) => (x === E.na ? (y === null ? 0 : y) : x));
  }
  if (v === E.na) {
    const r = evalNode(nodes[1], ctx);
    return r === null ? 0 : r;
  }
  return v === null ? 0 : v;
}, { lazy: true });

/** Truth values of UND/ODER/XODER: ranges and arrays skip text and empty cells. */
function truths(args: Any[], ctx: Ctx): boolean[] | CellError {
  const out: boolean[] = [];
  for (const a of args) {
    if (a instanceof Ref || a instanceof Arr) {
      for (const v of values(a, ctx)) {
        if (isErr(v)) return v;
        if (typeof v === "number") out.push(v !== 0);
        else if (typeof v === "boolean") out.push(v);
      }
    } else {
      if (isErr(a)) return a;
      if (a === null) { out.push(false); continue; }
      // Typed-in text other than WAHR/FALSCH is passed over (=UND("x";WAHR) is WAHR).
      const t = truth(a);
      if (isErr(t)) continue;
      out.push(t);
    }
  }
  return out.length ? out : E.value;
}

def("AND", 1, 255, (args, ctx) => {
  const t = truths(args, ctx);
  return isErr(t) ? t : t.every((x) => x);
});
def("OR", 1, 255, (args, ctx) => {
  const t = truths(args, ctx);
  return isErr(t) ? t : t.some((x) => x);
});
def("XOR", 1, 254, (args, ctx) => {
  const t = truths(args, ctx);
  return isErr(t) ? t : t.filter((x) => x).length % 2 === 1;
});
defS("NOT", 1, 1, ([v]) => {
  if (isErr(v)) return v;
  const t = truth(v);
  return isErr(t) ? t : !t;
});
def("TRUE", 0, 0, () => true);
def("FALSE", 0, 0, () => false);

def("SWITCH", 3, 254, (_a, ctx, nodes) => {
  const x = scalar(evalNode(nodes[0], ctx), ctx);
  if (isErr(x)) return x;
  let i = 1;
  for (; i + 1 < nodes.length; i += 2) {
    const v = scalar(evalNode(nodes[i], ctx), ctx);
    if (isErr(v)) return v;
    if (compare(x, v) === 0 && typeof (x ?? 0) === typeof (v ?? 0)) {
      const r = evalNode(nodes[i + 1], ctx);
      return r === null ? 0 : r;
    }
  }
  if (i < nodes.length) {
    const r = evalNode(nodes[i], ctx);
    return r === null ? 0 : r;
  }
  return E.na;
}, { lazy: true });

def("CHOOSE", 2, 255, (_a, ctx, nodes) => {
  const idx = deref(evalNode(nodes[0], ctx), ctx);
  const choose = (v: V): Any => {
    if (isErr(v)) return v;
    const k = toNum(v);
    if (isErr(k)) return k;
    const i = Math.trunc(k);
    if (i < 1 || i >= nodes.length) return E.value;
    const r = evalNode(nodes[i], ctx);
    return r === null ? 0 : r;
  };
  if (idx instanceof Arr) return mapN([idx], ([v]) => {
    const r = deref(choose(v), ctx);
    return r instanceof Arr ? r.d[0] : r;
  });
  return choose(idx);
}, { lazy: true });

// ----- IST… --------------------------------------------------------------------------------

function is(name: string, f: (v: V) => boolean): void {
  defS(name, 1, 1, ([v]) => f(v));
}
is("ISNUMBER", (v) => typeof v === "number");
is("ISTEXT", (v) => typeof v === "string");
is("ISNONTEXT", (v) => typeof v !== "string");
is("ISLOGICAL", (v) => typeof v === "boolean");
is("ISBLANK", (v) => v === null);
is("ISERROR", (v) => isErr(v));
is("ISERR", (v) => isErr(v) && v !== E.na);
is("ISNA", (v) => v === E.na);
defS("ISEVEN", 1, 1, ([v]) => {
  if (v === null) return true;
  if (typeof v === "boolean") return E.value;
  const x = n(v);
  return isErr(x) ? x : Math.trunc(x) % 2 === 0;
});
defS("ISODD", 1, 1, ([v]) => {
  if (v === null) return false;
  if (typeof v === "boolean") return E.value;
  const x = n(v);
  return isErr(x) ? x : Math.abs(Math.trunc(x) % 2) === 1;
});
def("ISREF", 1, 1, ([a]) => a instanceof Ref);
def("ISFORMULA", 1, 1, ([a], ctx) => {
  if (!(a instanceof Ref)) return isErr(a) ? a : E.value;
  const r = a.areas[0];
  return ctx.host.formula(a.sheet, r.r0, r.c0) !== undefined;
});
def("FORMULATEXT", 1, 1, ([a], ctx) => {
  if (!(a instanceof Ref)) return isErr(a) ? a : E.value;
  const r = a.areas[0];
  const f = ctx.host.formula(a.sheet, r.r0, r.c0);
  return f === undefined ? E.na : "=" + f;
});
def("NA", 0, 0, () => E.na);
const ERROR_TYPES = ["#NULL!", "#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#NUM!", "#N/A", "#GETTING_DATA", "", "", "", "", "", "#SPILL!", "#CALC!"];
defS("ERROR.TYPE", 1, 1, ([v]) => (isErr(v) ? ERROR_TYPES.indexOf(v.code) + 1 || E.na : E.na));
def("TYPE", 1, 1, ([a], ctx) => {
  const v = deref(a, ctx);
  if (v instanceof Arr) return 64;
  if (typeof v === "number" || v === null) return 1;
  if (typeof v === "string") return 2;
  if (typeof v === "boolean") return 4;
  return 16;
});
// N and T take the first cell of a range, but work element by element on arrays.
const nOf = (v: V): V => (typeof v === "number" ? v : typeof v === "boolean" ? (v ? 1 : 0) : isErr(v) ? v : 0);
const tOf = (v: V): V => (typeof v === "string" ? v : isErr(v) ? v : "");
def("N", 1, 1, ([a], ctx) => (a instanceof Arr ? map1(a, nOf) : nOf(scalar(a, ctx))));
def("T", 1, 1, ([a], ctx) => (a instanceof Arr ? map1(a, tOf) : tOf(scalar(a, ctx))));

// ----- References --------------------------------------------------------------------------

def("ROW", 0, 1, (args, ctx) => {
  if (!args.length || args[0] === null) return ctx.row + 1;
  const a = args[0];
  if (!(a instanceof Ref)) return isErr(a) ? a : E.value;
  const r = a.areas[0];
  if (r.r0 === r.r1) return r.r0 + 1;
  return new Arr(r.r1 - r.r0 + 1, 1, Array.from({ length: r.r1 - r.r0 + 1 }, (_, i) => r.r0 + i + 1));
});
def("COLUMN", 0, 1, (args, ctx) => {
  if (!args.length || args[0] === null) return ctx.col + 1;
  const a = args[0];
  if (!(a instanceof Ref)) return isErr(a) ? a : E.value;
  const r = a.areas[0];
  if (r.c0 === r.c1) return r.c0 + 1;
  return new Arr(1, r.c1 - r.c0 + 1, Array.from({ length: r.c1 - r.c0 + 1 }, (_, i) => r.c0 + i + 1));
});
def("ROWS", 1, 1, ([a]) => {
  if (a instanceof Ref) return a.areas.length > 1 ? E.ref : a.areas[0].r1 - a.areas[0].r0 + 1;
  if (a instanceof Arr) return a.rows;
  return isErr(a) ? a : 1;
});
def("COLUMNS", 1, 1, ([a]) => {
  if (a instanceof Ref) return a.areas.length > 1 ? E.ref : a.areas[0].c1 - a.areas[0].c0 + 1;
  if (a instanceof Arr) return a.cols;
  return isErr(a) ? a : 1;
});
def("AREAS", 1, 1, ([a]) => (a instanceof Ref ? a.areas.length : isErr(a) ? a : E.value));

/** A reference text in A1 or Z1S1 style ("Tabelle1!A1:B3", "'Daten 2'!A1", "Z2S1") read as INDIREKT does. */
export function textRef(text: string, a1: boolean, ctx: Ctx): Any {
  let t = text.trim();
  if (t.startsWith("=")) t = t.slice(1);
  if (!a1) return r1c1Ref(t, ctx);
  const toks = lex(t, "file").filter((x) => x.k !== "ws");
  if (toks.length === 1 && toks[0].k === "ref") {
    const r = toks[0].ref!;
    const s = r.sheet === undefined ? ctx.sheet : ctx.sheetIndex(r.sheet);
    if (s < 0) return E.ref;
    return new Ref(s, [refRect(r)]);
  }
  if (toks.length === 1 && toks[0].k === "name") {
    let node: Node;
    try {
      node = parseFormula(t);
    } catch {
      return E.ref;
    }
    const v = evalNode(node, ctx);
    return v instanceof Ref ? v : E.ref;
  }
  return E.ref;
}

function r1c1Ref(t: string, ctx: Ctx): Any {
  let sheet = ctx.sheet;
  const bang = t.lastIndexOf("!");
  if (bang >= 0) {
    let name = t.slice(0, bang);
    if (name.startsWith("'")) name = name.slice(1, -1).replace(/''/g, "'");
    sheet = ctx.sheetIndex(name);
    if (sheet < 0) return E.ref;
    t = t.slice(bang + 1);
  }
  const one = (s: string) => {
    const m = /^[ZzRr](\[-?\d+\]|\(-?\d+\)|\d+)?[SsCc](\[-?\d+\]|\(-?\d+\)|\d+)?$/.exec(s);
    if (!m) return null;
    const part = (p: string | undefined, base: number) => (p === undefined ? base : /^[[(]/.test(p) ? base + Number(p.slice(1, -1)) : Number(p) - 1);
    return { row: part(m[1], ctx.row), col: part(m[2], ctx.col) };
  };
  const parts = t.split(":");
  const a = one(parts[0]);
  const bb = parts[1] !== undefined ? one(parts[1]) : a;
  if (!a || !bb || parts.length > 2) return E.ref;
  if (a.row < 0 || a.col < 0 || bb.row >= MAX_ROWS || bb.col >= MAX_COLS) return E.ref;
  return new Ref(sheet, [{ r0: Math.min(a.row, bb.row), c0: Math.min(a.col, bb.col), r1: Math.max(a.row, bb.row), c1: Math.max(a.col, bb.col) }]);
}

def("INDIRECT", 1, 2, ([a, s], ctx) => {
  const t = scalar(a, ctx);
  if (isErr(t)) return t;
  const style = s === undefined ? true : scalar(s, ctx);
  if (isErr(style)) return style;
  const a1 = style === null ? false : toBool(style);
  if (isErr(a1)) return a1;
  if (typeof t !== "string") return E.ref;
  return textRef(t, a1, ctx);
}, { volatile: true });

def("OFFSET", 3, 5, ([a, r0, c0, h0, w0], ctx) => {
  if (!(a instanceof Ref)) return isErr(a) ? a : E.value;
  const get = (x: Any | undefined, d: number): number | CellError => {
    if (x === undefined || x === null) return d;
    const v = scalar(x, ctx);
    if (v === null) return d;
    const k = toNum(v);
    return isErr(k) ? k : Math.trunc(k);
  };
  const base = a.areas[0];
  const dr = get(r0, 0), dc = get(c0, 0), h = get(h0, base.r1 - base.r0 + 1), w = get(w0, base.c1 - base.c0 + 1);
  for (const x of [dr, dc, h, w]) if (isErr(x)) return x;
  const R0 = base.r0 + (dr as number), C0 = base.c0 + (dc as number);
  let hh = h as number, ww = w as number;
  if (hh === 0 || ww === 0) return E.ref;
  // A negative height or width counts upwards / leftwards from the corner.
  const rA = hh > 0 ? R0 : R0 + hh + 1, rB = hh > 0 ? R0 + hh - 1 : R0;
  const cA = ww > 0 ? C0 : C0 + ww + 1, cB = ww > 0 ? C0 + ww - 1 : C0;
  hh = Math.abs(hh);
  ww = Math.abs(ww);
  if (rA < 0 || cA < 0 || rB >= MAX_ROWS || cB >= MAX_COLS) return E.ref;
  return new Ref(a.sheet, [{ r0: rA, c0: cA, r1: rB, c1: cB }]);
}, { volatile: true });

defS("ADDRESS", 2, 5, ([r0, c0, abs0, a10, sh]) => {
  const r = n(r0), c = n(c0);
  if (isErr(r)) return r;
  if (isErr(c)) return c;
  const abs = abs0 === undefined || abs0 === null ? 1 : n(abs0);
  if (isErr(abs)) return abs;
  const a1 = a10 === undefined ? true : b(a10);
  if (isErr(a1)) return a1;
  const row = Math.trunc(r), col = Math.trunc(c), mode = Math.trunc(abs);
  if (row < 1 || col < 1 || row > MAX_ROWS || col > MAX_COLS || mode < 1 || mode > 4) return E.value;
  let s: string;
  if (a1) {
    s = (mode === 1 || mode === 3 ? "$" : "") + colName(col - 1) + (mode === 1 || mode === 2 ? "$" : "") + row;
  } else {
    // German Z1S1 writes relative parts in parentheses: Z(5)S(3).
    s = "Z" + (mode === 1 || mode === 2 ? row : `(${row})`) + "S" + (mode === 1 || mode === 3 ? col : `(${col})`);
  }
  if (sh !== undefined && sh !== null) {
    const name = typeof sh === "string" ? sh : String(sh);
    s = sheetPrefix(name) + s;
  }
  return s;
});

