// References inside formulas: copying (relative parts move), inserting and deleting rows and
// columns, moving cells, and the positions of references for Excel's coloured ranges while a
// formula is typed. Rules as recorded from Excel 365 (probe_excel_formulas.py).

import { MAX_COLS, MAX_ROWS, type Rect } from "../address";
import { refText, sheetPrefix } from "./convert";
import { lex, type RefEnd, type RefTok, type Token } from "./lexer";
import { refRect } from "./parse";

/** Rewrites every reference token of a file-form formula; `f` returns the new text or null to keep it. */
function rewrite(file: string, f: (t: Token, r: RefTok) => string | null): string {
  const toks = lex(file, "file");
  let out = "";
  let pos = 0;
  for (const t of toks) {
    if (t.k !== "ref") continue;
    const s = f(t, t.ref!);
    if (s === null) continue;
    out += file.slice(pos, t.s) + s;
    pos = t.e;
  }
  return out + file.slice(pos);
}

function refError(r: RefTok): string {
  return (r.sheet !== undefined ? sheetPrefix(r.sheet) : "") + "#REF!";
}

/**
 * The formula copied to a cell dRow rows and dCol columns away (AutoAusfüllen, copy and
 * paste): relative parts move, "$" parts stay, a reference pushed off the sheet becomes #BEZUG!.
 */
export function shiftFormula(file: string, dRow: number, dCol: number): string {
  if (!dRow && !dCol) return file;
  return rewrite(file, (_t, r) => {
    const move = (e: RefEnd): RefEnd | null => {
      const n = { ...e };
      if (n.row !== undefined && !n.rowAbs) n.row += dRow;
      if (n.col !== undefined && !n.colAbs) n.col += dCol;
      if (n.row !== undefined && (n.row < 0 || n.row >= MAX_ROWS)) return null;
      if (n.col !== undefined && (n.col < 0 || n.col >= MAX_COLS)) return null;
      return n;
    };
    const a = move(r.a);
    const b = r.b ? move(r.b) : undefined;
    if (!a || b === null) return refError(r);
    return refText({ ...r, a, b });
  });
}

export interface RowColChange {
  kind: "rows" | "cols";
  /** The sheet where rows or columns are inserted (count > 0) or deleted (count < 0). */
  sheet: string;
  /** First row or column (0-based). */
  at: number;
  count: number;
  /** The sheet the formula is on (its unqualified references); default: `sheet`. */
  home?: string;
}

export interface MoveChange {
  kind: "move";
  sheet: string;
  /** The cells that were cut, and where they went. */
  from: Rect;
  dRow: number;
  dCol: number;
  home?: string;
}

/** Zellen einfügen (down, right) or löschen (up, left) with the neighbours shifting. */
export interface CellsChange {
  kind: "cells";
  sheet: string;
  rect: Rect;
  dir: "down" | "right" | "up" | "left";
  home?: string;
}

export type Change = RowColChange | MoveChange | CellsChange;

function onSheet(r: RefTok, change: Change): boolean {
  if (r.sheet2 !== undefined) return false;
  const s = r.sheet ?? change.home ?? change.sheet;
  return s.toLowerCase() === change.sheet.toLowerCase();
}

/** Formula text after rows/columns were inserted or deleted, or cells were moved, as Excel adjusts it. */
export function adjustFormula(file: string, change: Change): string {
  return rewrite(file, (_t, r) => {
    if (!onSheet(r, change)) return null;
    const out = change.kind === "move" ? moveRef(r, change) : change.kind === "cells" ? cellsRef(r, change) : insertDelete(r, change);
    return out === r ? null : out === null ? refError(r) : refText(out);
  });
}

/** One axis of a reference ("row" or "col") after an insert or delete; null when it is deleted. */
function axis(r: RefTok, key: "row" | "col", at: number, count: number, max: number): RefTok | null {
  const a = r.a[key], b = (r.b ?? r.a)[key];
  if (a === undefined || b === undefined) return r;
  const lo = Math.min(a, b), hi = Math.max(a, b);
  let nlo: number, nhi: number;
  if (count > 0) {
    nlo = lo >= at ? lo + count : lo;
    nhi = hi >= at ? hi + count : hi;
    if (nhi >= max) return null;
  } else {
    const end = at - count - 1;
    if (lo >= at && hi <= end) return null;
    nlo = lo < at ? lo : lo > end ? lo + count : at;
    nhi = hi < at ? hi : hi > end ? hi + count : at - 1;
  }
  if (nlo === lo && nhi === hi) return r;
  const swap = a > b;
  const na: RefEnd = { ...r.a, [key]: swap ? nhi : nlo };
  const nb: RefEnd | undefined = r.b ? { ...r.b, [key]: swap ? nlo : nhi } : undefined;
  return { ...r, a: na, b: nb };
}

function insertDelete(r: RefTok, c: RowColChange): RefTok | null {
  return c.kind === "rows" ? axis(r, "row", c.at, c.count, MAX_ROWS) : axis(r, "col", c.at, c.count, MAX_COLS);
}

/**
 * Zellen einfügen/löschen: a reference inside the band of rows or columns that shifts is
 * adjusted like a row or column insert/delete; a range losing a whole end to deleted cells
 * keeps the rest (shifted when it lies in the band); everything else stays (recorded).
 */
function cellsRef(r: RefTok, c: CellsChange): RefTok | null {
  if (r.a.row === undefined || r.a.col === undefined || (r.b && (r.b.row === undefined || r.b.col === undefined))) return r;
  const rect = refRect(r);
  const R = c.rect;
  const vertical = c.dir === "down" || c.dir === "up";
  const insert = c.dir === "down" || c.dir === "right";
  const size = vertical ? R.r1 - R.r0 + 1 : R.c1 - R.c0 + 1;
  const inBand = vertical ? rect.c0 >= R.c0 && rect.c1 <= R.c1 : rect.r0 >= R.r0 && rect.r1 <= R.r1;
  if (inBand) return vertical ? axis(r, "row", R.r0, insert ? size : -size, MAX_ROWS) : axis(r, "col", R.c0, insert ? size : -size, MAX_COLS);
  if (insert) return r;
  const part = overlap(rect, R);
  if (!part) return r;
  let rest: Rect | null = null;
  if (vertical && part.r0 === rect.r0 && part.r1 === rect.r1) {
    if (part.c0 === rect.c0) rest = { ...rect, c0: part.c1 + 1 };
    else if (part.c1 === rect.c1) rest = { ...rect, c1: part.c0 - 1 };
  } else if (!vertical && part.c0 === rect.c0 && part.c1 === rect.c1) {
    if (part.r0 === rect.r0) rest = { ...rect, r0: part.r1 + 1 };
    else if (part.r1 === rect.r1) rest = { ...rect, r1: part.r0 - 1 };
  }
  if (!rest) return r;
  // What is left moves with the shifted cells when it lies beside the deleted ones.
  if (!vertical && rest.r0 >= R.r0 && rest.r1 <= R.r1 && rest.c0 > R.c1) {
    rest = { ...rest, c0: rest.c0 - size, c1: rest.c1 - size };
  }
  return {
    ...r,
    a: { row: rest.r0, col: rest.c0, rowAbs: r.a.rowAbs, colAbs: r.a.colAbs },
    b: { row: rest.r1, col: rest.c1, rowAbs: (r.b ?? r.a).rowAbs, colAbs: (r.b ?? r.a).colAbs },
  };
}

function within(x: Rect, y: Rect): boolean {
  return x.r0 >= y.r0 && x.r1 <= y.r1 && x.c0 >= y.c0 && x.c1 <= y.c1;
}

function overlap(x: Rect, y: Rect): Rect | null {
  const r = { r0: Math.max(x.r0, y.r0), c0: Math.max(x.c0, y.c0), r1: Math.min(x.r1, y.r1), c1: Math.min(x.c1, y.c1) };
  return r.r0 <= r.r1 && r.c0 <= r.c1 ? r : null;
}

/**
 * Excel's rules for cut and paste: references inside the cut block move with it; a range whose
 * end the block covers across its whole width (moving up or down) or height (moving sideways)
 * stretches to the moved part; references into the destination that the block overwrote
 * become #BEZUG!; all others stay.
 */
function moveRef(r: RefTok, c: MoveChange): RefTok | null {
  const rect = refRect(r);
  const whole = r.a.row === undefined || r.a.col === undefined || (r.b !== undefined && (r.b.row === undefined || r.b.col === undefined));
  if (whole) return r;
  const from = c.from;
  const dest = { r0: from.r0 + c.dRow, c0: from.c0 + c.dCol, r1: from.r1 + c.dRow, c1: from.c1 + c.dCol };
  const shift = (e: RefEnd): RefEnd => ({ ...e, row: e.row! + c.dRow, col: e.col! + c.dCol });
  if (within(rect, from)) return { ...r, a: shift(r.a), b: r.b ? shift(r.b) : undefined };
  if (within(rect, dest)) return null;
  const part = overlap(rect, from);
  if (!part) return r;
  // A slab at one end of the range, as wide (or high) as the range, moving along the range.
  let nr: Rect | null = null;
  if (c.dCol === 0 && part.c0 === rect.c0 && part.c1 === rect.c1 && (part.r0 === rect.r0 || part.r1 === rect.r1)) {
    const atStart = part.r0 === rect.r0;
    const [lo, hi] = slab(atStart, part.r0, part.r1, rect.r0, rect.r1, c.dRow);
    if (lo < 0) return r;
    nr = { r0: lo, r1: hi, c0: rect.c0, c1: rect.c1 };
  } else if (c.dRow === 0 && part.r0 === rect.r0 && part.r1 === rect.r1 && (part.c0 === rect.c0 || part.c1 === rect.c1)) {
    const atStart = part.c0 === rect.c0;
    const [lo, hi] = slab(atStart, part.c0, part.c1, rect.c0, rect.c1, c.dCol);
    if (lo < 0) return r;
    nr = { c0: lo, c1: hi, r0: rect.r0, r1: rect.r1 };
  }
  if (!nr) return r;
  return {
    ...r,
    a: { row: nr.r0, col: nr.c0, rowAbs: r.a.rowAbs, colAbs: r.a.colAbs },
    b: { row: nr.r1, col: nr.c1, rowAbs: (r.b ?? r.a).rowAbs, colAbs: (r.b ?? r.a).colAbs },
  };
}

/**
 * New extent along the move of a range [r0, r1] whose slab [p0, p1] at its start or end moved
 * by d: the moved slab alone when it covers the rest, the span of rest and slab otherwise,
 * [-1, -1] (unchanged) when the slab jumped past the rest.
 */
function slab(atStart: boolean, p0: number, p1: number, r0: number, r1: number, d: number): [number, number] {
  const rest = atStart ? [p1 + 1, r1] : [r0, p0 - 1];
  const n0 = p0 + d, n1 = p1 + d;
  if (n0 <= rest[0] && n1 >= rest[1]) return [n0, n1];
  if (atStart ? n0 > rest[1] : n1 < rest[0]) return [-1, -1];
  return [Math.min(rest[0], n0), Math.max(rest[1], n1)];
}

export interface FormulaRef {
  /** Offsets of the reference in the local text (end exclusive). */
  start: number;
  end: number;
  ref: Rect;
  /** Sheet name when the reference names one. */
  sheet?: string;
}

/**
 * The references of a formula as typed (German form, with or without "="), with their place
 * in the text, for Excel's coloured range borders while editing.
 */
export function formulaRefs(local: string): FormulaRef[] {
  const out: FormulaRef[] = [];
  for (const t of lex(local, "local")) {
    if (t.k !== "ref" || t.ref!.sheet2 !== undefined) continue;
    const r = t.ref!;
    out.push({ start: t.s, end: t.e, ref: refRect(r), ...(r.sheet !== undefined ? { sheet: r.sheet } : {}) });
  }
  return out;
}

