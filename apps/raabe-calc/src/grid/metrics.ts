// Column widths and row heights as Excel computes them. The model keeps a column width in
// characters of the default font's widest digit as the .xlsx layer reads it (the file's width
// minus 5/7, i.e. at 96 dpi). On screen Excel lays widths out with the display's own digit width
// and padding (measured at 150 %: 11 device pixels and 7 px, so a column dragged 40 CSS px wider
// shows 15,73 characters, not 16,43; D-062): the grid does the same, so columns are exactly as
// wide as Excel's at any display scaling. Printing uses the 96-dpi values. Row heights are in
// points: 15 pt = 20 px.

import { MAX_COLS, MAX_ROWS, type Sheet } from "@raabe/sheet-engine";

/** Widest digit of Calibri 11 at 96 dpi (Excel's "maximum digit width"). */
export const MDW = 7;
/** German Excel's standard column width (base width 10, rounded up to 8 px: 80 px at 96 dpi). */
export const DEFAULT_COL_CHARS = 75 / 7;

/** The display's digit width and padding in device pixels (set by the grid from devicePixelRatio). */
const DEV = { dpr: 1, mdw: 7, pad: 5 };

export function setDeviceScale(dpr: number): void {
  if (dpr === DEV.dpr && DEV.mdw) return;
  DEV.dpr = dpr;
  let mdw = Math.round(7 * dpr);
  try {
    const ctx = document.createElement("canvas").getContext("2d");
    if (ctx) {
      ctx.font = `${((11 * 96) / 72) * dpr}px "Calibri", Carlito, sans-serif`;
      const w = Math.max(...[..."0123456789"].map((d) => ctx.measureText(d).width));
      if (w > 0) mdw = Math.round(w);
    }
  } catch { /* no canvas (tests): the 96-dpi value scaled */ }
  DEV.mdw = mdw;
  DEV.pad = Math.floor((5 * mdw) / 7);
}

/** The display's scale and the cell padding in device pixels (for fitting text as Excel does). */
export function device(): { dpr: number; pad: number; mdw: number } {
  return { dpr: DEV.dpr, pad: DEV.pad, mdw: DEV.mdw };
}

/** The file's width (characters plus padding, in 1/256) → device pixels with digit width `mdw`. */
function filePx(w: number, mdw: number): number {
  return Math.trunc(((256 * w + Math.trunc(128 / mdw)) / 256) * mdw);
}

/** A model width (characters) → CSS px on this display. */
export function charsToPx(chars: number): number {
  if (chars <= 0) return 0;
  return filePx(chars + 5 / MDW, DEV.mdw) / DEV.dpr;
}

/** A model width → px at 96 dpi (printing). */
export function charsToPx96(chars: number): number {
  return chars <= 0 ? 0 : filePx(chars + 5 / MDW, MDW);
}

/** The standard width on this display: the base width's digits plus padding, rounded up to 8 px (scaled). */
export function defaultColPx(chars: number): number {
  const base = Math.floor(chars + 1e-6);
  const step = 8 * DEV.dpr;
  return (Math.ceil((base * DEV.mdw + DEV.pad) / step) * step) / DEV.dpr;
}

/** CSS px on this display → the model's width (what a drag of the column border stores). */
export function pxToWidth(px: number): number {
  if (px <= 0) return 0;
  return Math.max(0, Math.trunc(((px * DEV.dpr) / DEV.mdw) * 256) / 256 - 5 / MDW);
}

/** A model width as Excel shows it on this display (Spaltenbreite, the tip while dragging). */
export function shownChars(chars: number): number {
  return Math.max(0, chars + 5 / MDW - DEV.pad / DEV.mdw);
}

/** A width typed into Spaltenbreite (as shown on this display) → the model's width. */
export function shownToChars(shown: number): number {
  return shown <= 0 ? 0 : shown + DEV.pad / DEV.mdw - 5 / MDW;
}

/** CSS px → characters as Excel shows them, rounded to hundredths. */
export function pxToChars(px: number): number {
  return Math.round(shownChars(pxToWidth(px)) * 100) / 100;
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
  const colDef = defaultColPx(sheet.defaultColWidth ?? DEFAULT_COL_CHARS);
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
