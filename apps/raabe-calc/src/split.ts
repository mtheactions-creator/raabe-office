// Ansicht > Fenster > Teilen (D-093): up to four panes with their own scrolling, as Excel 365 splits a window
// (recorded in our own Excel, probe_excel_windows.py). The split bars sit where the file's pane says (xSplit /
// ySplit in 1/20 pt from the window's left/top edge, headers included); the top panes share their rows, the
// left panes their columns. Teilen splits above and left of the active cell, at A1 in the middle of the window
// (recorded: SplitRow 24 / SplitColumn 12 of a window showing A1:X48); the pane holding the active cell becomes
// the active one. The bars are dragged, a double-click removes one, dragging one to the window's edge too.
// Fenster fixieren on a split freezes at the bars and Fixierung aufheben gives the split back (recorded).

import { h } from "@raabe/shell";
import type { Rect, View } from "@raabe/sheet-engine";
import type { Axis } from "./grid/metrics";
import type { Pane } from "./grid/render";

/** The split bar's thickness in CSS px. */
export const BAR = 4;
/** 1/20 pt per CSS px (0.75 pt). */
export const TWIPS_PER_PX = 15;

export interface Geometry {
  rows: Axis; cols: Axis; zoom: number; headerW: number; headerH: number; width: number; height: number;
  /**
   * The rows' heights in pt as the file has them and the column headers' (Excel's: the default row height). The
   * horizontal bar's ySplit counts in these (recorded: 1740 = 14,5 pt header + 5 rows of 14,5 pt), so that a bar
   * between two rows stays between them whatever our pixel heights are.
   */
  rowPt?: (r: number) => number;
  headPt?: number;
}

/** ySplit (1/20 pt at the zoom) to CSS px on the canvas, counting rows from the top pane's first one. */
export function yToPx(v: View, g: Geometry, twips: number): number {
  if (!g.rowPt || !g.headPt || !v.split) return twips / TWIPS_PER_PX;
  let pt = twips / 20 / g.zoom - g.headPt;
  let px = g.headerH;
  if (pt <= 0) return px + (pt * g.zoom * 4) / 3;
  for (let r = v.split.top; r < g.rows.count; r++) {
    const h = g.rowPt(r), size = g.rows.size(r) * g.zoom;
    if (h <= 0) continue;
    if (pt < h - 1e-6) return px + (pt / h) * size;
    pt -= h;
    px += size;
  }
  return px;
}

/** CSS px on the canvas to ySplit: whole rows in the file's pt, the rest in pt of the row it falls into. */
export function pxToY(v: View, g: Geometry, px: number, top = v.split?.top ?? v.top): number {
  if (!g.rowPt || !g.headPt) return Math.round(px * TWIPS_PER_PX);
  let pt = g.headPt;
  let at = g.headerH;
  for (let r = top; r < g.rows.count; r++) {
    const h = g.rowPt(r), size = g.rows.size(r) * g.zoom;
    if (h <= 0 || size <= 0) continue;
    if (px < at + size - 1e-6) { pt += ((px - at) / size) * h; break; }
    pt += h;
    at += size;
  }
  return Math.round(pt * 20 * g.zoom);
}

type PaneName = "topLeft" | "topRight" | "bottomLeft" | "bottomRight";
export interface NamedPane extends Pane { name: PaneName }

/** The bars' places in CSS px on the canvas (0: none). */
export function bars(v: View, g: Geometry): { x: number; y: number } {
  const s = v.split;
  if (!s) return { x: 0, y: 0 };
  const x = s.x ? Math.round(s.x / TWIPS_PER_PX) : 0, y = s.y ? Math.round(yToPx(v, g, s.y)) : 0;
  return { x: x > g.headerW && x < g.width ? x : 0, y: y > g.headerH && y < g.height ? y : 0 };
}

/** The panes of a split view, the bottom-right (main) one last. */
export function splitPanes(v: View, g: Geometry): NamedPane[] {
  const s = v.split!;
  const b = bars(v, g);
  const hw = g.headerW, hh = g.headerH;
  const xs = b.x ? [{ x: hw, w: b.x - hw, col: s.left, left: true }, { x: b.x + BAR, w: g.width - b.x - BAR, col: v.left, left: false }] : [{ x: hw, w: g.width - hw, col: v.left, left: false }];
  const ys = b.y ? [{ y: hh, h: b.y - hh, row: s.top, top: true }, { y: b.y + BAR, h: g.height - b.y - BAR, row: v.top, top: false }] : [{ y: hh, h: g.height - hh, row: v.top, top: false }];
  const out: NamedPane[] = [];
  for (const r of ys) for (const c of xs) {
    const name = `${r.top || (!b.y) ? "top" : "bottom"}${c.left || !b.x ? "Left" : "Right"}` as PaneName;
    // Without a horizontal bar the panes are topLeft/topRight, without a vertical one topLeft/bottomLeft (the file's names).
    const fixed: PaneName = !b.y ? (c.left ? "topLeft" : "topRight") : !b.x ? (r.top ? "topLeft" : "bottomLeft") : name;
    out.push({ name: fixed, row: r.row, col: c.col, x: c.x, y: r.y, w: Math.max(0, c.w), h: Math.max(0, r.h) });
  }
  return out;
}

/** The last row/column a pane shows. */
function lastIn(axis: Axis, first: number, len: number, zoom: number): number {
  return axis.at(axis.pos(first) + len / zoom);
}

/** The pane a range's top-left cell is drawn in: the active pane when it shows it, else the first that does. */
export function paneOf(r: Rect, panes: NamedPane[], g: Geometry, active?: string): NamedPane {
  const shows = (p: NamedPane) => r.r0 >= p.row && r.r0 <= lastIn(g.rows, p.row, p.h, g.zoom) && r.c0 >= p.col && r.c0 <= lastIn(g.cols, p.col, p.w, g.zoom);
  const a = panes.find((p) => p.name === active);
  if (a && shows(a)) return a;
  return panes.find(shows) ?? a ?? panes[panes.length - 1];
}

/** Whether the active pane scrolls the top panes' rows (split.top) and the left panes' columns (split.left). */
export function activeScroll(v: View): { top: boolean; left: boolean } {
  const s = v.split;
  if (!s || v.freeze) return { top: false, left: false };
  const p = s.pane ?? "topLeft";
  return { top: !!s.y && p.startsWith("top"), left: !!s.x && p.endsWith("Left") };
}

/**
 * Teilen at the active cell: the bars above and left of it; at A1 (or the top-left visible cell) in the middle
 * of the window. Returns the split, or null when the cell gives none.
 */
export function splitAt(v: View, g: Geometry): View["split"] | null {
  const { row, col } = v.active;
  let r = row, c = col;
  const atCorner = (row === v.top || row === 0) && (col === v.left || col === 0);
  if (atCorner) {
    // The middle of the window, at the nearest row and column border (recorded: 24 of 48 rows, 12 of 24 columns).
    r = g.rows.at(g.rows.pos(v.top) + (g.height - g.headerH) / 2 / g.zoom + 0.5 * g.rows.size(v.top));
    c = g.cols.at(g.cols.pos(v.left) + (g.width - g.headerW) / 2 / g.zoom + 0.5 * g.cols.size(v.left));
  }
  const y = r > v.top ? g.headerH + (g.rows.pos(r) - g.rows.pos(v.top)) * g.zoom : 0;
  const x = c > v.left ? g.headerW + (g.cols.pos(c) - g.cols.pos(v.left)) * g.zoom : 0;
  if (!x && !y) return null;
  const pane: PaneName = x && y ? (atCorner ? "topLeft" : "bottomRight") : x ? (atCorner ? "topLeft" : "topRight") : atCorner ? "topLeft" : "bottomLeft";
  return { x: Math.round(x * TWIPS_PER_PX), y: y ? pxToY(v, g, y, v.top) : 0, top: v.top, left: v.left, pane };
}

/** After Teilen the bottom/right panes start at the bar's row/column (Excel: A6 below a bar above row 6). */
export function afterSplit(v: View, g: Geometry): void {
  const s = v.split!;
  const b = bars(v, g);
  if (b.y) v.top = g.rows.at(g.rows.pos(s.top) + (b.y - g.headerH) / g.zoom + 0.01);
  if (b.x) v.left = g.cols.at(g.cols.pos(s.left) + (b.x - g.headerW) / g.zoom + 0.01);
}

/** The rows and columns the split's panes freeze (Fenster fixieren on a split, recorded: the bars' cells). */
export function freezeOfSplit(v: View, g: Geometry): { row: number; col: number } {
  const s = v.split!;
  const b = bars(v, g);
  return {
    row: b.y ? g.rows.at(g.rows.pos(s.top) + (b.y - g.headerH) / g.zoom + 0.01) : 0,
    col: b.x ? g.cols.at(g.cols.pos(s.left) + (b.x - g.headerW) / g.zoom + 0.01) : 0,
  };
}

/** The split at the frozen rows and columns (Fixierung aufheben of a split, or Teilen while frozen). */
export function splitOfFreeze(v: View, g: Geometry): View["split"] | null {
  const f = v.freeze;
  if (!f) return null;
  const top = 0, left = 0;
  const y = f.row ? g.headerH + (g.rows.pos(f.row) - g.rows.pos(top)) * g.zoom : 0;
  const x = f.col ? g.headerW + (g.cols.pos(f.col) - g.cols.pos(left)) * g.zoom : 0;
  if (!x && !y) return null;
  return { x: Math.round(x * TWIPS_PER_PX), y: y ? pxToY(v, g, y, top) : 0, top, left, pane: x && y ? "bottomRight" : x ? "topRight" : "bottomLeft" };
}

/**
 * The bars on the sheet (DOM elements over the canvas), their dragging, and the top panes' own scroll bar.
 * `host` gives the grid's geometry and redraws after a change.
 */
export function splitBars(host: {
  el: HTMLElement; vscroll: () => HTMLElement; view(): View; geometry(): Geometry | null; changed(): void; done(): void;
}) {
  const css = `
.split-bar { position: absolute; background: #d4d4d4; z-index: 4; touch-action: none; }
.split-bar.v { top: 0; bottom: 0; width: ${BAR}px; cursor: col-resize; border-left: 1px solid #c2c2c2; border-right: 1px solid #c2c2c2; box-sizing: border-box; }
.split-bar.h { left: 0; height: ${BAR}px; cursor: row-resize; border-top: 1px solid #c2c2c2; border-bottom: 1px solid #c2c2c2; box-sizing: border-box; }
.split-bar.drag { background: #a6a6a6; }
.vscroll.split-top { bottom: auto; }
`;
  const barV = h("div", { class: "split-bar v", "aria-label": "Teilung", hidden: true });
  const barH = h("div", { class: "split-bar h", "aria-label": "Teilung", hidden: true });
  const topScroll = h("div", { class: "vscroll split-top", hidden: true }, h("div"));
  let quiet = false;

  function install(): void {
    if (!document.getElementById("raabe-split-css")) document.head.append(h("style", { id: "raabe-split-css" }, css));
    if (!barV.isConnected) host.el.append(barV, barH, topScroll);
  }

  function drag(e: PointerEvent, axis: "x" | "y"): void {
    if (e.button !== 0) return;
    e.preventDefault();
    e.stopPropagation();
    const el = axis === "x" ? barV : barH;
    el.setPointerCapture(e.pointerId);
    el.classList.add("drag");
    const r0 = host.el.getBoundingClientRect();
    const move = (ev: PointerEvent) => {
      const v = host.view(), g = host.geometry();
      if (!v.split || !g) return;
      const at = axis === "x" ? ev.clientX - r0.left : ev.clientY - r0.top;
      const lo = axis === "x" ? g.headerW : g.headerH, hi = axis === "x" ? g.width : g.height;
      const px = Math.max(lo, Math.min(hi, at));
      v.split[axis] = axis === "y" ? pxToY(v, g, px) : Math.round(px * TWIPS_PER_PX);
      host.changed();
    };
    const up = (ev: PointerEvent) => {
      el.releasePointerCapture(ev.pointerId);
      el.classList.remove("drag");
      el.removeEventListener("pointermove", move);
      el.removeEventListener("pointerup", up);
      const v = host.view(), g = host.geometry();
      if (v.split && g) {
        // Dragged onto the headers or the window's edge: that bar is gone (Excel).
        const px = axis === "y" ? yToPx(v, g, v.split.y) : v.split.x / TWIPS_PER_PX;
        const lo = axis === "x" ? g.headerW : g.headerH, hi = axis === "x" ? g.width : g.height;
        if (px <= lo + 2 || px >= hi - 2) remove(axis);
      }
      host.changed();
      host.done();
    };
    el.addEventListener("pointermove", move);
    el.addEventListener("pointerup", up);
  }

  /** One bar removed; the split goes when it was the last. */
  function remove(axis: "x" | "y"): void {
    const v = host.view();
    const s = v.split;
    if (!s) return;
    if (axis === "y") { v.top = s.y ? v.top : v.top; s.y = 0; s.top = v.top; }
    else { s.x = 0; s.left = v.left; }
    if (!s.x && !s.y) delete v.split;
    else s.pane = s.x ? (s.pane?.endsWith("Left") ? "topLeft" : "topRight") : (s.pane?.startsWith("top") ? "topLeft" : "bottomLeft");
  }

  barV.addEventListener("pointerdown", (e) => drag(e, "x"));
  barH.addEventListener("pointerdown", (e) => drag(e, "y"));
  barV.addEventListener("dblclick", () => { remove("x"); host.changed(); host.done(); });
  barH.addEventListener("dblclick", () => { remove("y"); host.changed(); host.done(); });
  topScroll.addEventListener("scroll", () => {
    if (quiet) return;
    const v = host.view(), g = host.geometry();
    if (!v.split || !g) return;
    v.split.top = g.rows.at(topScroll.scrollTop / g.zoom + 0.5);
    host.changed();
  });

  /** After every paint: bars and the top panes' scroll bar where the split is. */
  function place(): void {
    const v = host.view(), g = host.geometry();
    const b = v.split && !v.freeze && g ? bars(v, g) : { x: 0, y: 0 };
    if (!b.x && !b.y && !barV.isConnected) return;
    install();
    barV.hidden = !b.x;
    barH.hidden = !b.y;
    if (b.x) Object.assign(barV.style, { left: `${b.x}px`, height: `${g!.height}px` });
    if (b.y) Object.assign(barH.style, { top: `${b.y}px`, width: `${g!.width + 17}px` });
    const main = host.vscroll();
    topScroll.hidden = !b.y;
    if (b.y && g) {
      main.style.top = `${b.y + BAR}px`;
      Object.assign(topScroll.style, { height: `${b.y}px` });
      (topScroll.firstElementChild as HTMLElement).style.height = `${g.rows.pos(Math.max(v.split!.top + 200, 1000)) * g.zoom}px`;
      quiet = true;
      topScroll.scrollTop = g.rows.pos(v.split!.top) * g.zoom;
      requestAnimationFrame(() => { quiet = false; });
    } else main.style.top = "";
  }

  return { place, remove };
}

/** A click in another pane makes it the active one; the pane left keeps its selection (as Excel's file does). */
export function activatePane(v: View, name: string): void {
  const s = v.split;
  if (!s || s.pane === name) return;
  const prev = s.pane ?? "topLeft";
  const cell = (r: number, c: number) => `${colLetters(c)}${r + 1}`;
  const sq = v.selection.map((r) => (r.r0 === r.r1 && r.c0 === r.c1 ? cell(r.r0, r.c0) : `${cell(r.r0, r.c0)}:${cell(r.r1, r.c1)}`)).join(" ");
  s.sels = { ...(s.sels ?? {}), [prev]: { ac: cell(v.active.row, v.active.col), sqref: sq || cell(v.active.row, v.active.col) } };
  delete s.sels[name];
  s.pane = name as NonNullable<typeof s.pane>;
}

function colLetters(c: number): string {
  let s = "";
  for (let n = c + 1; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + ((n - 1) % 26)) + s;
  return s;
}

/**
 * Teilen pressed (recorded through Window.Split): a split goes; frozen panes become a split at the same place;
 * otherwise the window splits at the active cell. Returns whether the view changed.
 */
export function toggleSplit(v: View, g: Geometry): boolean {
  if (v.freeze) {
    const s = splitOfFreeze(v, g);
    v.top = Math.max(v.top, v.freeze.row);
    v.left = Math.max(v.left, v.freeze.col);
    delete v.freeze;
    delete v.splitFrozen;
    if (s) v.split = s;
    return true;
  }
  if (v.split) {
    // The panes' scrolling goes; the window shows from the top-left pane's first cell (Excel).
    v.top = v.split.y ? v.split.top : v.top;
    v.left = v.split.x ? v.split.left : v.left;
    delete v.split;
    return true;
  }
  const s = splitAt(v, g);
  if (!s) return false;
  v.split = s;
  afterSplit(v, g);
  return true;
}

/** Fenster fixieren on a split view: frozen at the bars (state frozenSplit); Fixierung aufheben gives it back. */
export function freezeSplit(v: View, g: Geometry): void {
  const f = freezeOfSplit(v, g);
  if (!f.row && !f.col) return;
  v.freeze = f;
  v.splitFrozen = true;
  delete v.split;
}

export function unfreezeToSplit(v: View, g: Geometry): void {
  const s = splitOfFreeze(v, g);
  v.top = Math.max(v.top, v.freeze?.row ?? 0);
  v.left = Math.max(v.left, v.freeze?.col ?? 0);
  delete v.freeze;
  delete v.splitFrozen;
  if (s) v.split = s;
}
