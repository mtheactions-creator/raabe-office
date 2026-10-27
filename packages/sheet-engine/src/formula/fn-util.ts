// Helpers shared by the function libraries: argument lifting, collecting numbers the way
// Excel's aggregate functions do, and decimal rounding.

import { Ctx, def, deref, mapN, scalar } from "./eval";
import type { Node } from "./parse";
import { Arr, E, Ref, isErr, textToNumber, toBool, toNum, toStr, type Any, type V } from "./values";
import type { CellError } from "../model";

/**
 * A function of single values; ranges and arrays given for its arguments make it work
 * element by element (dynamic arrays). Arguments listed in `raw` are passed unchanged.
 */
export function defS(name: string, min: number, max: number, f: (v: V[], ctx: Ctx, raw: Any[]) => V | Arr, raw: number[] = [], opts: { volatile?: boolean } = {}): void {
  def(name, min, max, (args, ctx) => {
    const lifted: (V | Arr)[] = args.map((a, i) => (raw.includes(i) ? null : deref(a, ctx)));
    return mapN(lifted, (vals) => f(vals, ctx, args));
  }, opts);
}

/** Numeric argument: text read as a number, WAHR 1, empty 0. */
export function n(v: V): number | CellError {
  return toNum(v);
}

export function s(v: V): string | CellError {
  return toStr(v);
}

export function b(v: V): boolean | CellError {
  return toBool(v);
}

export interface Collect {
  /** MITTELWERTA: text in arrays counts as 0 (recorded; the other "A" functions skip it). */
  arrText0?: boolean;
  /** Typed-in text is skipped instead of read as a number (MODUS.EINF). */
  skipText?: boolean;
  /** "A" functions: text in ranges counts 0, WAHR/FALSCH 1/0. */
  a?: boolean;
  /** Errors in ranges are skipped instead of returned (AGGREGAT). */
  skipErrors?: boolean;
}

/**
 * The numbers of an aggregate's arguments as Excel collects them: typed-in values count
 * (text read as a number or #WERT!), in ranges and arrays only numbers count; the first
 * error found is returned.
 */
export function numbers(args: Any[], ctx: Ctx, opt: Collect = {}): number[] | CellError {
  const out: number[] = [];
  for (const a of args) {
    if (a instanceof Ref) {
      let e: CellError | null = null;
      ctx.eachCell(a, (v) => {
        if (e) return;
        if (typeof v === "number") out.push(v);
        else if (isErr(v)) { if (!opt.skipErrors) e = v; }
        else if (opt.a) out.push(typeof v === "boolean" ? (v ? 1 : 0) : 0);
      });
      if (e) return e;
    } else if (a instanceof Arr) {
      for (const v of a.d) {
        if (typeof v === "number") out.push(v);
        else if (isErr(v)) { if (!opt.skipErrors) return v; }
        // In arrays, "A" functions skip text (MITTELWERTA counts it as 0) and count logical values.
        else if (opt.a && typeof v === "boolean") out.push(v ? 1 : 0);
        else if (opt.arrText0 && typeof v === "string") out.push(0);
      }
    } else {
      if (isErr(a)) return a;
      if (opt.skipText && typeof a === "string") continue;
      const x = toNum(a);
      if (isErr(x)) return x;
      out.push(x);
    }
  }
  return out;
}

/** Every value of an argument in order (ranges row by row, including empty cells as null). */
export function values(a: Any, ctx: Ctx): V[] {
  if (a instanceof Ref) {
    const out: V[] = [];
    for (let i = 0; i < a.areas.length; i++) out.push(...ctx.refArr(a, i).d);
    return out;
  }
  if (a instanceof Arr) return a.d;
  return [a];
}

/** An argument as an array (a single value as 1×1). */
export function toArr(a: Any, ctx: Ctx): Arr | CellError {
  if (a instanceof Ref) {
    if (a.areas.length > 1) return E.value;
    return ctx.refArr(a);
  }
  if (a instanceof Arr) return a;
  return new Arr(1, 1, [a]);
}

/** The single value of an argument that must be one (arrays give their first element). */
export function one(a: Any, ctx: Ctx): V {
  return scalar(a, ctx);
}

export function isMissing(nodes: Node[], i: number): boolean {
  return i >= nodes.length || nodes[i].t === "miss";
}

/** Reads a numeric text the way VALUE does; null when it is not one. */
export function parseNum(text: string): number | null {
  return textToNumber(text);
}

// ----- Rounding on decimal digits, as Excel does (2,675 rounds to 2,68) --------------------

/** Rounds to `d` decimals: "half" away from zero, "up" away from zero, "down" towards zero. */
export function roundTo(x: number, d: number, mode: "half" | "up" | "down"): number {
  if (!Number.isFinite(x) || x === 0) return x;
  d = Math.trunc(d);
  const neg = x < 0;
  const p = Math.abs(x).toExponential(14);
  const [mant, e] = p.split("e");
  const digits = mant.replace(".", "");
  const point = Number(e) + 1;
  const k = point + d;
  let r: number;
  if (k >= 15) return x;
  if (k < 0) {
    r = mode === "up" ? Number("1e" + -d) : 0;
  } else {
    const keep = digits.slice(0, k);
    const rest = digits.slice(k);
    let up = false;
    if (mode === "half") up = rest[0] >= "5";
    else if (mode === "up") up = /[1-9]/.test(rest);
    let kept = BigInt(keep || "0");
    if (up) kept += 1n;
    r = Number(kept.toString() + "e" + (point - k));
  }
  return neg ? -r : r;
}

/** The value to 15 significant digits, as Excel stores what it shows. */
export function r15(x: number): number {
  return Number.isFinite(x) && x !== 0 ? Number(x.toPrecision(15)) : x;
}
