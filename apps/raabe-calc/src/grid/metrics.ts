// Column widths and row heights as Excel computes them (at 96 dpi, in CSS px; the grid scales
// them by the zoom). The model keeps a column width in characters of the default font's widest
// digit (Calibri 11: 7 px), as Excel's Spaltenbreite shows it; the file adds 5 px of padding.
// German Excel's standard width is 10.71 characters = 80 px. Row heights are in points:
// 15 pt = 20 px.

import { MAX_COLS, MAX_ROWS, type Sheet } from "@raabe/sheet-engine";

/** Widest digit of Calibri 11 at 96 dpi (Excel's "maximum digit width"). */
export const MDW = 7;
/** German Excel's standard column width (base width 10, rounded up to 8 px: 80 px). */
export const DEFAULT_COL_CHARS = 75 / 7;

/** Characters (as Excel shows them in Spaltenbreite) → px. */
export function charsToPx(chars: number): number {
  if (chars <= 0) return 0;
  return fileWidthToPx(charsToFileWidth(chars));
}

/** px → the model's width in characters (unrounded, as the .xlsx layer keeps it). */
export function pxToWidth(px: number): number {
  return px <= 0 ? 0 : Math.max(0, (Math.trunc((px / MDW) * 256) / 256) - 5 / MDW);
}

/** Characters → the file's width (8.43 → 9.140625). */
export function charsToFileWidth(chars: number): number {
  return Math.trunc(((chars * MDW + 5) / MDW) * 256) / 256;
}

/** px → characters, rounded to hundredths as Excel shows them. */
export function pxToChars(px: number): number {
  if (px <= 5) return Math.round((px / (MDW + 5)) * 100) / 100;
  return Math.round(((px - 5) / MDW) * 100) / 100;
}

/** The file's width (characters plus padding, in 1/256) → px. */
export function fileWidthToPx(w: number): number {
  return Math.trunc(((256 * w + Math.trunc(128 / MDW)) / 256) * MDW);
}

/** px → the file's width. */
export function pxToFileWidth(px: number): number {
  return Math.trunc((px / MDW) * 256) / 256;
}

export function ptToPx(pt: number): number {
  return Math.round((pt * 96) / 72);
}

export function pxToPt(px: number): number {
  return Math.round(((px * 72) / 96) * 4) / 4;
}

/**
 * Positions of rows or columns: a default size plus a sorted list of exceptions (own sizes,
 * hidden ones), so a million rows cost nothing until they differ.
 */
export class Axis {
  private keys: number[] = [];
  private sizes: number[] = [];
  /** Sum of (size - default) for exceptions before index i in `keys`. */
  private deltas: number[] = [];

  constructor(public def: number, readonly count: number) {}

  set(exceptions: Map<number, number>, def: number): void {
    this.def = def;
    this.keys = [...exceptions.keys()].sort((a, b) => a - b);
    this.sizes = this.keys.map((k) => exceptions.get(k)!);
    this.deltas = [];
    let d = 0;
    for (let i = 0; i < this.keys.length; i++) {
      this.deltas.push(d);
      d += this.sizes[i] - def;
    }
    this.deltas.push(d);
  }

  private find(i: number): number {
    let lo = 0, hi = this.keys.length;
    while (lo < hi) {
      const mid = (lo + hi) >> 1;
      if (this.keys[mid] < i) lo = mid + 1;
      else hi = mid;
    }
    return lo;
  }

  size(i: number): number {
    const k = this.find(i);
    return this.keys[k] === i ? this.sizes[k] : this.def;
  }

  /** Start of row/column i from the start of the sheet. */
  pos(i: number): number {
    const k = this.find(i);
    return i * this.def + this.deltas[k];
  }

  /** The row/column at a position (the last one when past the end). */
  at(p: number): number {
    if (p <= 0) return 0;
    let lo = 0, hi = this.count - 1;
    while (lo < hi) {
      const mid = (lo + hi + 1) >> 1;
      if (this.pos(mid) <= p) lo = mid;
      else hi = mid - 1;
    }
    return lo;
  }

  hidden(i: number): boolean {
    return this.size(i) === 0;
  }

  /** The next visible index after i in direction dir (i itself if none). */
  step(i: number, dir: 1 | -1): number {
    let j = i + dir;
    while (j >= 0 && j < this.count && this.hidden(j)) j += dir;
    return j < 0 || j >= this.count ? i : j;
  }
}

/** Builds the row and column axes of a sheet (CSS px at 100 %). */
export function sheetAxes(sheet: Sheet): { rows: Axis; cols: Axis } {
  const colDef = sheet.defaultColWidth !== undefined ? charsToPx(sheet.defaultColWidth) : charsToPx(DEFAULT_COL_CHARS);
  const cols = new Axis(colDef, MAX_COLS);
  const cx = new Map<number, number>();
  for (const [c, info] of sheet.cols) {
    const w = info.hidden ? 0 : info.width !== undefined ? charsToPx(info.width) : colDef;
    if (w !== colDef) cx.set(c, w);
  }
  cols.set(cx, colDef);
  const rowDef = ptToPx(sheet.defaultRowHeight);
  const rows = new Axis(rowDef, MAX_ROWS);
  const rx = new Map<number, number>();
  for (const [r, info] of sheet.rows) {
    const hgt = info.hidden ? 0 : info.height !== undefined ? ptToPx(info.height) : rowDef;
    if (hgt !== rowDef) rx.set(r, hgt);
  }
  rows.set(rx, rowDef);
  return { rows, cols };
}
