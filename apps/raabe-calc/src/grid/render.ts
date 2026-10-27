// Draws a sheet onto the grid's canvas as Excel does (recorded from Excel 365 at 150 %,
// tools/feel-lab/capture_excel_ui.py): white cells with 1-device-pixel gridlines (#E0E0E0),
// headers on #F0F0F0 with a darker line towards the cells, the selection tinted with a
// 2-device-pixel frame and the fill handle at its corner. Excel's green accent is Raabe red
// (D-018). The canvas has the device's pixel size, so lines stay one device pixel wide.

import { cellKey, inRect, type Cell, type CellStyle, type Rect, type Sheet, type StyleTable, type BorderSide } from "@raabe/sheet-engine";
import { shown, type Shown } from "../engine";
import type { CondEval } from "./cond";
import type { Axis } from "./metrics";

export const COLORS = {
  grid: "#e0e0e0",
  header: "#f0f0f0",
  headerSel: "#e1e1e1",
  headerSep: "#e2e2e2",
  headerSelSep: "#bebebe",
  headerLine: "#ababab",
  headerText: "#444444",
  accent: "#b3141b",
  selTint: "rgba(0, 0, 0, 0.09)",
};

export interface Pane {
  /** Rows and columns drawn (first index; the pane is filled to its size). */
  row: number;
  col: number;
  /** Pane position and size in CSS px on the canvas. */
  x: number;
  y: number;
  w: number;
  h: number;
}

export interface Placed { i: number; at: number; size: number }

export interface DrawState {
  sheet: Sheet;
  styles: StyleTable;
  rows: Axis;
  cols: Axis;
  zoom: number;
  dpr: number;
  width: number;
  height: number;
  headerW: number;
  headerH: number;
  panes: Pane[];
  selection: Rect[];
  active: { row: number; col: number };
  /** The cell being edited: its content is not drawn (the editor shows it). */
  editing?: { row: number; col: number };
  /** Kopieren/Ausschneiden: the marching ants. */
  marquee?: Rect | null;
  marqueePhase?: number;
  /** AutoFill or move preview: a grey frame. */
  ghost?: Rect | null;
  /** Coloured ranges of the formula being typed. */
  refs?: { rect: Rect; color: string }[];
  showGrid: boolean;
  showHeaders: boolean;
  /** Text of cells, cached by key; cleared when cells or sizes change. */
  textCache: Map<number, Shown>;
  /** Bedingte Formatierung of the sheet (null without rules). */
  cond: CondEval | null;
  /** Formulas instead of results (Formeln anzeigen, Strg+#). */
  showFormulas?: boolean;
  /** Livevorschau: the format under the mouse in a list or palette, shown on the selection. */
  preview?: { ranges: Rect[]; change: (s: CellStyle) => void } | null;
}

const PT = 96 / 72;

export function cssFont(style: CellStyle, zoom: number): string {
  const f = style.font;
  // Hoch- and tiefgestellt cells are drawn smaller (Excel: about two thirds of the size).
  const size = f.vertAlign ? f.size * 0.66 : f.size;
  return `${f.italic ? "italic " : ""}${f.bold ? "bold " : ""}${(size * PT * zoom).toFixed(3)}px "${f.name}", Carlito, sans-serif`;
}

/** A fill the grid does not draw as such (patterns, gradients): its look as one colour. */
export function fillColor(style: CellStyle): string | undefined {
  if (style.fill) return style.fill;
  const xml = style.fillXml;
  if (!xml) return undefined;
  const rgb = (tag: string) => new RegExp(`<${tag}[^>]*rgb="(?:FF)?([0-9A-Fa-f]{6})"`).exec(xml)?.[1];
  const grad = [...xml.matchAll(/<color[^>]*rgb="(?:FF)?([0-9A-Fa-f]{6})"/g)].map((m) => m[1]);
  if (/gradientFill/.test(xml) && grad.length) return "#" + mixHex(grad[0], grad[grad.length - 1], 0.5);
  const fg = rgb("fgColor"), bg = rgb("bgColor") ?? "FFFFFF";
  const type = /patternType="([^"]+)"/.exec(xml)?.[1];
  if (!type || type === "none" || !fg) return undefined;
  // Pattern fills: the foreground colour at the pattern's density.
  const density: Record<string, number> = { solid: 1, darkGray: 0.75, mediumGray: 0.5, lightGray: 0.25, gray125: 0.125, gray0625: 0.0625 };
  return "#" + mixHex(bg, fg, density[type] ?? 0.4);
}

function mixHex(a: string, b: string, t: number): string {
  const p = (h: string, i: number) => parseInt(h.slice(i * 2, i * 2 + 2), 16);
  return [0, 1, 2].map((i) => Math.round(p(a, i) + (p(b, i) - p(a, i)) * t).toString(16).padStart(2, "0")).join("");
}

/** The rows or columns of a pane with their positions on the canvas. */
export function placed(axis: Axis, first: number, start: number, length: number, zoom: number, limit = Infinity): Placed[] {
  const out: Placed[] = [];
  let at = start;
  for (let i = first; i < axis.count && i < limit && at < start + length; i++) {
    const size = axis.size(i) * zoom;
    if (size > 0) out.push({ i, at, size });
    at += size;
  }
  return out;
}

export function draw(ctx: CanvasRenderingContext2D, s: DrawState): void {
  const { dpr } = s;
  ctx.setTransform(dpr, 0, 0, dpr, 0, 0);
  ctx.fillStyle = "#ffffff";
  ctx.fillRect(0, 0, s.width, s.height);
  for (const pane of s.panes) drawPane(ctx, s, pane);
  // Fenster fixieren: a thin dark line where the frozen rows and columns end (Excel).
  if (s.panes.length > 1) {
    const px = 1 / dpr;
    const main = s.panes[s.panes.length - 1];
    ctx.fillStyle = "#808080";
    if (main.y > s.headerH + 0.5) ctx.fillRect(s.headerW, snap(main.y, dpr) - px, s.width - s.headerW, px);
    if (main.x > s.headerW + 0.5) ctx.fillRect(snap(main.x, dpr) - px, s.headerH, px, s.height - s.headerH);
  }
  if (s.showHeaders) drawHeaders(ctx, s);
}

/** Snaps a CSS coordinate to the device pixel grid (returns CSS px). */
function snap(v: number, dpr: number): number {
  return Math.round(v * dpr) / dpr;
}

function drawPane(ctx: CanvasRenderingContext2D, s: DrawState, pane: Pane): void {
  const { zoom, dpr, sheet } = s;
  const rows = placed(s.rows, pane.row, pane.y, pane.h, zoom);
  const cols = placed(s.cols, pane.col, pane.x, pane.w, zoom);
  if (!rows.length || !cols.length) return;
  ctx.save();
  ctx.beginPath();
  ctx.rect(pane.x, pane.y, pane.w, pane.h);
  ctx.clip();
  const px = 1 / dpr;
  const r0 = rows[0].i, r1 = rows[rows.length - 1].i;
  const c0 = cols[0].i, c1 = cols[cols.length - 1].i;
  const colAt = new Map(cols.map((c) => [c.i, c]));
  const merges = sheet.merges.filter((m) => m.r1 >= r0 && m.r0 <= r1 && m.c1 >= c0 && m.c0 <= c1);
  const rectOf = (m: Rect) => {
    const x = pane.x + (s.cols.pos(m.c0) - s.cols.pos(pane.col)) * zoom;
    const y = pane.y + (s.rows.pos(m.r0) - s.rows.pos(pane.row)) * zoom;
    const w = (s.cols.pos(m.c1 + 1) - s.cols.pos(m.c0)) * zoom;
    const h = (s.rows.pos(m.r1 + 1) - s.rows.pos(m.r0)) * zoom;
    return { x, y, w, h };
  };

  // Gridlines: the last device pixel of every row and column.
  if (s.showGrid) {
    ctx.fillStyle = COLORS.grid;
    for (const c of cols) ctx.fillRect(snap(c.at + c.size, dpr) - px, pane.y, px, pane.h);
    for (const r of rows) ctx.fillRect(pane.x, snap(r.at + r.size, dpr) - px, pane.w, px);
  }

  // Fills (they cover the gridlines around the cell, as in Excel), merged areas white.
  const previewed = new Map<number | undefined, CellStyle>();
  const styleOf = (cell: Cell | undefined, row: number, col: number): CellStyle => {
    const id = cell?.s ?? sheet.rows.get(row)?.s ?? sheet.cols.get(col)?.s;
    const p = s.preview;
    if (!p || !p.ranges.some((r) => inRect(r, row, col))) return s.styles.get(id);
    let st = previewed.get(id);
    if (!st) { st = structuredClone(s.styles.get(id)); p.change(st); previewed.set(id, st); }
    return st;
  };
  for (const r of rows) {
    for (const c of cols) {
      const cell = sheet.get(r.i, c.i);
      const st = styleOf(cell, r.i, c.i);
      const look = s.cond?.look(r.i, c.i);
      const fill = look?.fill ?? fillColor(st);
      if (fill) {
        ctx.fillStyle = fill;
        ctx.fillRect(snap(c.at, dpr) - px, snap(r.at, dpr) - px, snap(c.size, dpr) + px, snap(r.size, dpr) + px);
      }
      if (look?.bar) {
        // Datenbalken: a gradient bar from the left, 2 px inside the cell (Excel's default).
        const w = Math.max(0, (c.size - 4 * zoom) * Math.min(1, look.bar.frac));
        const g = ctx.createLinearGradient(c.at, 0, c.at + w, 0);
        g.addColorStop(0, look.bar.color);
        g.addColorStop(1, "#ffffff");
        ctx.fillStyle = g;
        ctx.fillRect(c.at + 2 * zoom, r.at + 2 * zoom, w, r.size - 5 * zoom);
        ctx.strokeStyle = look.bar.color;
        ctx.lineWidth = px;
        ctx.strokeRect(c.at + 2 * zoom, r.at + 2 * zoom, w, r.size - 5 * zoom);
      }
    }
  }
  for (const m of merges) {
    const st = styleOf(sheet.get(m.r0, m.c0), m.r0, m.c0);
    const b = rectOf(m);
    ctx.fillStyle = st.fill ?? "#ffffff";
    ctx.fillRect(snap(b.x, dpr), snap(b.y, dpr), snap(b.w, dpr) - px, snap(b.h, dpr) - px);
  }

  // Text, with overflow into empty neighbours as Excel does for unwrapped text.
  const measureCtx = ctx;
  for (const r of rows) {
    const overflowFrom = Math.max(0, c0 - 30), overflowTo = c1 + 30;
    for (let ci = overflowFrom; ci <= overflowTo && ci < s.cols.count; ci++) {
      if (s.editing && s.editing.row === r.i && s.editing.col === ci) continue;
      const key = cellKey(r.i, ci);
      const cell = sheet.cells.get(key);
      if (!cell || cell.v === null || cell.v === "") continue;
      if (merges.some((m) => inRect(m, r.i, ci) && (m.r0 !== r.i || m.c0 !== ci))) continue;
      const merge = merges.find((m) => m.r0 === r.i && m.c0 === ci) ?? sheet.mergeAt(r.i, ci);
      if (merge && (merge.r0 !== r.i || merge.c0 !== ci)) continue;
      const baseSt = styleOf(cell, r.i, ci);
      const look = s.cond?.look(r.i, ci);
      const st = look && (look.bold || look.italic || look.underline || look.strike)
        ? { ...baseSt, font: { ...baseSt.font, bold: look.bold ?? baseSt.font.bold, italic: look.italic ?? baseSt.font.italic, underline: look.underline ? "single" as const : baseSt.font.underline, strike: look.strike ?? baseSt.font.strike } }
        : baseSt;
      const font = cssFont(st, zoom);
      measureCtx.font = font;
      let box: { x: number; y: number; w: number; h: number };
      if (merge) box = rectOf(merge);
      else {
        const c = colAt.get(ci);
        const x = c ? c.at : pane.x + (s.cols.pos(ci) - s.cols.pos(pane.col)) * zoom;
        const w = s.cols.size(ci) * zoom;
        if (w === 0) continue;
        box = { x, y: r.at, w, h: r.size };
      }
      const pad = 2 * zoom + 0.5;
      let text = s.textCache.get(key);
      if (!text) {
        text = s.showFormulas && cell.f !== undefined
          ? { text: "=" + cell.f, align: "left" as const }
          : shown(cell, st, Math.max(0, box.w - 2 * pad) / zoom, (t) => measureCtx.measureText(t).width / zoom);
        s.textCache.set(key, text);
      }
      if (!text.text) continue;
      const h = st.h ?? "general";
      const align = h === "general" ? text.align : h === "centerContinuous" ? "center" : h === "justify" || h === "distributed" ? "left" : h === "fill" ? "left" : h;
      const width = measureCtx.measureText(text.text).width;
      // How far unwrapped text may run: into empty cells beside it, left/right by alignment.
      let left = box.x, right = box.x + box.w;
      const isText = typeof cell.v === "string";
      if (!st.wrap && !merge && isText && width > box.w - 2 * pad) {
        const spill = (dir: 1 | -1, need: number) => {
          let edge = dir > 0 ? right : left;
          let j = ci + dir;
          while (need > 0 && j >= 0 && j < s.cols.count) {
            const n = sheet.get(r.i, j);
            if ((n && n.v !== null && n.v !== "") || sheet.mergeAt(r.i, j)) break;
            const w = s.cols.size(j) * zoom;
            edge += dir * w;
            need -= w;
            j += dir;
          }
          return edge;
        };
        const over = width - (box.w - 2 * pad);
        if (align === "left") right = spill(1, over);
        else if (align === "right") left = spill(-1, over);
        else { left = spill(-1, over / 2); right = spill(1, over / 2); }
      }
      // Clipping to the cell (or the cells it runs into); never over the frozen pane's edge.
      ctx.save();
      ctx.beginPath();
      ctx.rect(left, box.y, right - left, box.h);
      ctx.clip();
      ctx.fillStyle = look?.color ?? text.color ?? st.font.color ?? "#000000";
      ctx.font = font;
      const lines = st.wrap && isText ? wrapLines(measureCtx, text.text, box.w - 2 * pad) : [text.text];
      const m = measureCtx.measureText("Hg");
      const asc = m.fontBoundingBoxAscent, desc = m.fontBoundingBoxDescent;
      const lineH = asc + desc;
      const total = lineH * lines.length;
      const v = st.v ?? "bottom";
      let top = v === "top" ? box.y + 1 * zoom : v === "center" ? box.y + (box.h - total) / 2 : box.y + box.h - total - 1 * zoom;
      for (const line of lines) {
        const lw = measureCtx.measureText(line).width;
        const indent = (st.indent ?? 0) * 3 * 7 * zoom;
        let x = align === "right" ? box.x + box.w - pad - lw - indent : align === "center" ? box.x + (box.w - lw) / 2 : box.x + pad + indent;
        if (!st.wrap && isText && width > box.w - 2 * pad) {
          if (align === "left") x = box.x + pad + indent;
          else if (align === "right") x = box.x + box.w - pad - lw;
          else x = box.x + (box.w - lw) / 2;
        }
        if (text.fill && lines.length === 1) {
          drawFilled(ctx, text, box, pad, zoom);
        } else {
          const shift = st.font.vertAlign === "superscript" ? -st.font.size * PT * zoom * 0.3 : st.font.vertAlign === "subscript" ? st.font.size * PT * zoom * 0.08 : 0;
          ctx.fillText(line, x, top + asc + shift);
          decorate(ctx, st, x, top + asc + shift, lw, zoom, desc);
        }
        top += lineH;
      }
      ctx.restore();
    }
  }

  // Borders (the stronger of two neighbours' sides wins; here: drawn in order, right/bottom last).
  for (const r of rows) {
    for (const c of cols) {
      const cell = sheet.get(r.i, c.i);
      const st = styleOf(cell, r.i, c.i);
      const b = st.borders;
      if (!b.left && !b.right && !b.top && !b.bottom && !b.diagonal) continue;
      const x0 = snap(c.at, dpr) - px, y0 = snap(r.at, dpr) - px;
      const x1 = snap(c.at + c.size, dpr) - px, y1 = snap(r.at + r.size, dpr) - px;
      if (b.top) line(ctx, b.top, x0, y0, x1, y0, dpr, zoom);
      if (b.bottom) line(ctx, b.bottom, x0, y1, x1, y1, dpr, zoom);
      if (b.left) line(ctx, b.left, x0, y0, x0, y1, dpr, zoom);
      if (b.right) line(ctx, b.right, x1, y0, x1, y1, dpr, zoom);
      if (b.diagonal && b.diagonalDown) line(ctx, b.diagonal, x0, y0, x1, y1, dpr, zoom);
      if (b.diagonal && b.diagonalUp) line(ctx, b.diagonal, x0, y1, x1, y0, dpr, zoom);
    }
  }

  // Fehlerüberprüfung: a green triangle in the top left corner of numbers stored as text and of
  // formulas that give an error (Excel's rules that are on by default).
  ctx.fillStyle = "#18a043";
  for (const r of rows) for (const c of cols) {
    const cell = sheet.get(r.i, c.i);
    if (!cell || !errorIndicator(cell)) continue;
    const x = snap(c.at, dpr), y = snap(r.at, dpr);
    const t = 5 * zoom;
    ctx.beginPath();
    ctx.moveTo(x, y);
    ctx.lineTo(x + t, y);
    ctx.lineTo(x, y + t);
    ctx.closePath();
    ctx.fill();
  }

  // Notes: a small red triangle in the top right corner.
  ctx.fillStyle = "#c00000";
  for (const r of rows) for (const c of cols) {
    if (!sheet.get(r.i, c.i)?.note) continue;
    const x = snap(c.at + c.size, dpr) - px, y = snap(r.at, dpr);
    const t = 5 * zoom;
    ctx.beginPath();
    ctx.moveTo(x - t, y);
    ctx.lineTo(x, y);
    ctx.lineTo(x, y + t);
    ctx.closePath();
    ctx.fill();
  }

  // AutoFilter buttons on the filter's first row (Excel: a small box with an arrow, a funnel when filtered).
  const af = sheet.autoFilter;
  if (af && af.ref.r0 >= r0 && af.ref.r0 <= r1) {
    const r = rows.find((x) => x.i === af.ref.r0);
    if (r) for (const c of cols) {
      if (c.i < af.ref.c0 || c.i > af.ref.c1) continue;
      const b = filterButton(c.at, c.size, r.at, r.size, zoom);
      ctx.fillStyle = "#ffffff";
      ctx.fillRect(b.x, b.y, b.w, b.h);
      ctx.strokeStyle = "#a6a6a6";
      ctx.lineWidth = px;
      ctx.strokeRect(b.x + px / 2, b.y + px / 2, b.w - px, b.h - px);
      ctx.fillStyle = "#444444";
      const filtered = !!af.columns[c.i - af.ref.c0];
      const cx = b.x + b.w / 2, cy = b.y + b.h / 2;
      ctx.beginPath();
      if (filtered) {
        ctx.moveTo(cx - 4 * zoom, cy - 3 * zoom); ctx.lineTo(cx + 4 * zoom, cy - 3 * zoom); ctx.lineTo(cx + 1 * zoom, cy + 1 * zoom); ctx.lineTo(cx + 1 * zoom, cy + 4 * zoom); ctx.lineTo(cx - 1 * zoom, cy + 3 * zoom); ctx.lineTo(cx - 1 * zoom, cy + 1 * zoom);
      } else {
        ctx.moveTo(cx - 3.5 * zoom, cy - 1.5 * zoom); ctx.lineTo(cx + 3.5 * zoom, cy - 1.5 * zoom); ctx.lineTo(cx, cy + 2 * zoom);
      }
      ctx.closePath();
      ctx.fill();
    }
  }

  // Coloured references of the formula being typed.
  for (const ref of s.refs ?? []) {
    const b = rectOf(ref.rect);
    ctx.fillStyle = ref.color + "1f";
    ctx.fillRect(b.x, b.y, b.w, b.h);
    frame(ctx, b, ref.color, dpr, 1);
  }

  // Selection: tinted cells except the active one, one frame around the last range.
  const am = expandToMerges(sheet, { r0: s.active.row, c0: s.active.col, r1: s.active.row, c1: s.active.col });
  const a = rectOf(am);
  for (const sel of s.selection) {
    const single = sel.r0 === sel.r1 && sel.c0 === sel.c1;
    if (single && s.selection.length === 1) continue;
    const vis = clipRect(sel, r0, c0, r1, c1);
    if (!vis) continue;
    const b = rectOf(vis);
    ctx.fillStyle = COLORS.selTint;
    const hole = inRect(sel, s.active.row, s.active.col) ? a : null;
    tintExcept(ctx, b, hole);
  }
  const last = s.selection[s.selection.length - 1];
  if (last && s.selection.length === 1) {
    const full = expandToMerges(sheet, last);
    const b = rectOf(full);
    frame(ctx, b, COLORS.accent, dpr, 2);
    // Fill handle: 6 device pixels at the bottom-right corner, with a white edge.
    if (inRect({ r0, c0, r1: r1 + 1, c1: c1 + 1 }, full.r1, full.c1)) {
      const hx = snap(b.x + b.w, dpr), hy = snap(b.y + b.h, dpr);
      const hs = 6 / dpr;
      ctx.fillStyle = "#ffffff";
      ctx.fillRect(hx - hs / 2 - 2 * px - px, hy - hs / 2 - 2 * px - px, hs + 2 * px, hs + 2 * px);
      ctx.fillStyle = COLORS.accent;
      ctx.fillRect(hx - hs / 2 - 2 * px, hy - hs / 2 - 2 * px, hs, hs);
    }
  } else if (s.selection.length > 1) {
    frame(ctx, a, COLORS.accent, dpr, 1);
  }
  if (s.ghost) frame(ctx, rectOf(s.ghost), "#808080", dpr, 2, [3, 2]);
  if (s.marquee) {
    const b = rectOf(s.marquee);
    ctx.save();
    ctx.lineDashOffset = -(s.marqueePhase ?? 0);
    frame(ctx, b, COLORS.accent, dpr, 2, [4 / dpr * 1.5, 4 / dpr * 1.5]);
    ctx.restore();
  }
  ctx.restore();
}

/** Fills b except the rectangle `hole` (the active cell stays white). */
function tintExcept(ctx: CanvasRenderingContext2D, b: { x: number; y: number; w: number; h: number }, hole: { x: number; y: number; w: number; h: number } | null): void {
  if (!hole) { ctx.fillRect(b.x, b.y, b.w, b.h); return; }
  const x1 = b.x + b.w, y1 = b.y + b.h, hx1 = hole.x + hole.w, hy1 = hole.y + hole.h;
  if (hole.y > b.y) ctx.fillRect(b.x, b.y, b.w, Math.min(hole.y, y1) - b.y);
  if (hy1 < y1) ctx.fillRect(b.x, Math.max(hy1, b.y), b.w, y1 - Math.max(hy1, b.y));
  const ty = Math.max(hole.y, b.y), th = Math.min(hy1, y1) - ty;
  if (th > 0) {
    if (hole.x > b.x) ctx.fillRect(b.x, ty, Math.min(hole.x, x1) - b.x, th);
    if (hx1 < x1) ctx.fillRect(Math.max(hx1, b.x), ty, x1 - Math.max(hx1, b.x), th);
  }
}

function clipRect(r: Rect, r0: number, c0: number, r1: number, c1: number): Rect | null {
  const out = { r0: Math.max(r.r0, r0), c0: Math.max(r.c0, c0), r1: Math.min(r.r1, r1), c1: Math.min(r.c1, c1) };
  return out.r0 > out.r1 || out.c0 > out.c1 ? null : out;
}

/** A range grown to cover every merged area it touches. */
export function expandToMerges(sheet: Sheet, r: Rect): Rect {
  const out = { ...r };
  let grown = true;
  while (grown) {
    grown = false;
    for (const m of sheet.merges) {
      if (m.r1 < out.r0 || m.r0 > out.r1 || m.c1 < out.c0 || m.c0 > out.c1) continue;
      if (m.r0 < out.r0 || m.c0 < out.c0 || m.r1 > out.r1 || m.c1 > out.c1) {
        out.r0 = Math.min(out.r0, m.r0); out.c0 = Math.min(out.c0, m.c0);
        out.r1 = Math.max(out.r1, m.r1); out.c1 = Math.max(out.c1, m.c1);
        grown = true;
      }
    }
  }
  return out;
}

/** A frame `width` device pixels wide over the rectangle's edges (Excel: on the gridline and one pixel outside). */
function frame(ctx: CanvasRenderingContext2D, b: { x: number; y: number; w: number; h: number }, color: string, dpr: number, width: number, dash?: number[]): void {
  const px = 1 / dpr;
  const x0 = snap(b.x, dpr) - px * width, y0 = snap(b.y, dpr) - px * width;
  const x1 = snap(b.x + b.w, dpr) - px, y1 = snap(b.y + b.h, dpr) - px;
  const t = px * width;
  if (dash) {
    ctx.strokeStyle = color;
    ctx.lineWidth = t;
    ctx.setLineDash(dash);
    ctx.strokeRect(x0 + t / 2, y0 + t / 2, x1 + t - x0 - t, y1 + t - y0 - t);
    ctx.setLineDash([]);
    return;
  }
  ctx.fillStyle = color;
  ctx.fillRect(x0, y0, x1 + t - x0, t);
  ctx.fillRect(x0, y1 - t + px * width, x1 + t - x0, t);
  ctx.fillRect(x0, y0, t, y1 + t - y0);
  ctx.fillRect(x1, y0, t, y1 + t - y0);
}

const BORDER_WIDTH: Record<string, number> = { hair: 1, thin: 1, dotted: 1, dashed: 1, dashDot: 1, dashDotDot: 1, medium: 2, mediumDashed: 2, mediumDashDot: 2, mediumDashDotDot: 2, slantDashDot: 2, thick: 3, double: 3 };
const BORDER_DASH: Record<string, number[]> = { dotted: [1, 1], hair: [1, 1], dashed: [3, 1], mediumDashed: [6, 2], dashDot: [6, 2, 2, 2], mediumDashDot: [6, 2, 2, 2], dashDotDot: [6, 2, 2, 2, 2, 2], mediumDashDotDot: [6, 2, 2, 2, 2, 2], slantDashDot: [8, 2, 2, 2] };

function line(ctx: CanvasRenderingContext2D, side: BorderSide, x0: number, y0: number, x1: number, y1: number, dpr: number, zoom: number): void {
  const px = 1 / dpr;
  const w = Math.max(1, Math.round((BORDER_WIDTH[side.style] ?? 1) * zoom * Math.max(1, dpr * 0.75))) * px;
  ctx.fillStyle = side.color ?? "#000000";
  const vertical = x0 === x1;
  const dash = BORDER_DASH[side.style];
  if (side.style === "double") {
    const g = px;
    if (vertical) { ctx.fillRect(x0 - g, y0, px, y1 - y0 + px); ctx.fillRect(x0 + g, y0, px, y1 - y0 + px); }
    else { ctx.fillRect(x0, y0 - g, x1 - x0 + px, px); ctx.fillRect(x0, y0 + g, x1 - x0 + px, px); }
    return;
  }
  if (x0 !== x1 && y0 !== y1) {
    ctx.strokeStyle = side.color ?? "#000000";
    ctx.lineWidth = w;
    ctx.beginPath();
    ctx.moveTo(x0, y0);
    ctx.lineTo(x1, y1);
    ctx.stroke();
    return;
  }
  const off = Math.floor((w / px - 1) / 2) * px;
  if (!dash) {
    if (vertical) ctx.fillRect(x0 - off, y0 - off, w, y1 - y0 + w);
    else ctx.fillRect(x0 - off, y0 - off, x1 - x0 + w, w);
    return;
  }
  const unit = px * Math.max(1, Math.round(dpr));
  const len = vertical ? y1 - y0 + px : x1 - x0 + px;
  let pos = 0, k = 0;
  while (pos < len) {
    const seg = dash[k % dash.length] * unit;
    if (k % 2 === 0) {
      const l = Math.min(seg, len - pos);
      if (vertical) ctx.fillRect(x0 - off, y0 + pos, w, l);
      else ctx.fillRect(x0 + pos, y0 - off, l, w);
    }
    pos += seg;
    k++;
  }
}

function decorate(ctx: CanvasRenderingContext2D, st: CellStyle, x: number, base: number, w: number, zoom: number, desc: number): void {
  const t = Math.max(1, Math.round(st.font.size / 11)) * zoom;
  if (st.font.underline) {
    const y = base + Math.min(desc - t, 2 * zoom);
    ctx.fillRect(x, y, w, t * 0.8);
    if (st.font.underline.startsWith("double")) ctx.fillRect(x, y + 2 * t, w, t * 0.8);
  }
  if (st.font.strike) ctx.fillRect(x, base - st.font.size * PT * zoom * 0.28, w, t * 0.8);
}

function drawFilled(ctx: CanvasRenderingContext2D, text: Shown, box: { x: number; y: number; w: number; h: number }, pad: number, zoom: number): void {
  // "*x": the character x repeated where the format puts the star.
  const ch = text.fill ?? " ";
  const cw = ctx.measureText(ch).width || 1;
  const tw = ctx.measureText(text.text).width;
  const room = box.w - 2 * pad - tw;
  const n = Math.max(0, Math.floor(room / cw));
  const m = ctx.measureText("Hg");
  const y = box.y + box.h - m.fontBoundingBoxDescent - 1 * zoom;
  ctx.fillText(ch.repeat(n) + text.text, box.x + box.w - pad - tw - n * cw, y);
}

export function wrapLines(ctx: CanvasRenderingContext2D, text: string, width: number): string[] {
  const out: string[] = [];
  for (const para of text.split("\n")) {
    const words = para.split(/(?<= )/);
    let line = "";
    for (const w of words) {
      if (line && ctx.measureText(line + w.trimEnd()).width > width) {
        out.push(line.trimEnd());
        line = "";
      }
      // A word longer than the cell breaks between characters.
      let word = w;
      while (!line && ctx.measureText(word.trimEnd()).width > width && word.length > 1) {
        let k = word.length - 1;
        while (k > 1 && ctx.measureText(word.slice(0, k)).width > width) k--;
        out.push(word.slice(0, k));
        word = word.slice(k);
      }
      line += word;
    }
    out.push(line.trimEnd());
  }
  return out;
}

function drawHeaders(ctx: CanvasRenderingContext2D, s: DrawState): void {
  const { dpr, zoom, sheet } = s;
  const px = 1 / dpr;
  const sel = s.selection;
  const colSelected = (c: number) => sel.some((r) => c >= r.c0 && c <= r.c1);
  const rowSelected = (r: number) => sel.some((x) => r >= x.r0 && r <= x.r1);
  const fullCol = (c: number) => sel.some((r) => c >= r.c0 && c <= r.c1 && r.r0 === 0 && r.r1 === s.rows.count - 1);
  const fullRow = (r: number) => sel.some((x) => r >= x.r0 && r <= x.r1 && x.c0 === 0 && x.c1 === s.cols.count - 1);
  const font = `${(11 * PT * zoom).toFixed(3)}px "${s.styles.get(0).font.name}", Carlito, sans-serif`;
  ctx.font = font;
  ctx.textBaseline = "middle";
  // Column headers.
  ctx.fillStyle = COLORS.header;
  ctx.fillRect(0, 0, s.width, s.headerH);
  ctx.fillRect(0, 0, s.headerW, s.height);
  const colPanes = uniquePanes(s.panes, "x");
  for (const pane of colPanes) {
    ctx.save();
    ctx.beginPath();
    ctx.rect(pane.x, 0, pane.w, s.headerH);
    ctx.clip();
    for (const c of placed(s.cols, pane.col, pane.x, pane.w, zoom)) {
      const on = colSelected(c.i);
      const x0 = snap(c.at, dpr), x1 = snap(c.at + c.size, dpr);
      if (on) {
        ctx.fillStyle = fullCol(c.i) ? "#c9c9c9" : COLORS.headerSel;
        ctx.fillRect(x0 - px, 0, x1 - x0 + px, s.headerH);
      }
      ctx.fillStyle = on ? COLORS.headerSelSep : COLORS.headerSep;
      ctx.fillRect(x1 - px, 0, px, s.headerH - px);
      ctx.fillStyle = on ? COLORS.accent : COLORS.headerText;
      ctx.textAlign = "center";
      ctx.fillText(colLabel(c.i), (x0 + x1) / 2, s.headerH / 2 + 0.5);
      if (on) {
        ctx.fillStyle = COLORS.accent;
        ctx.fillRect(x0 - px, snap(s.headerH, dpr) - 2 * px, x1 - x0 + px, 2 * px);
      }
    }
    ctx.restore();
  }
  ctx.fillStyle = COLORS.headerLine;
  ctx.fillRect(s.headerW, snap(s.headerH, dpr) - px, s.width, px);
  // Row headers.
  const rowPanes = uniquePanes(s.panes, "y");
  for (const pane of rowPanes) {
    ctx.save();
    ctx.beginPath();
    ctx.rect(0, pane.y, s.headerW, pane.h);
    ctx.clip();
    for (const r of placed(s.rows, pane.row, pane.y, pane.h, zoom)) {
      const on = rowSelected(r.i);
      const y0 = snap(r.at, dpr), y1 = snap(r.at + r.size, dpr);
      if (on) {
        ctx.fillStyle = fullRow(r.i) ? "#c9c9c9" : COLORS.headerSel;
        ctx.fillRect(0, y0 - px, s.headerW, y1 - y0 + px);
      }
      ctx.fillStyle = on ? COLORS.headerSelSep : COLORS.headerSep;
      ctx.fillRect(0, y1 - px, s.headerW - px, px);
      ctx.fillStyle = on ? COLORS.accent : COLORS.headerText;
      ctx.textAlign = "center";
      ctx.fillText(String(r.i + 1), s.headerW / 2, (y0 + y1) / 2 + 0.5);
      if (on) {
        ctx.fillStyle = COLORS.accent;
        ctx.fillRect(snap(s.headerW, dpr) - 2 * px, y0 - px, 2 * px, y1 - y0 + px);
      }
    }
    ctx.restore();
  }
  ctx.fillStyle = COLORS.headerLine;
  ctx.fillRect(snap(s.headerW, dpr) - px, s.headerH, px, s.height);
  // Alles auswählen: the corner with its triangle.
  ctx.fillStyle = COLORS.header;
  ctx.fillRect(0, 0, s.headerW - px, s.headerH - px);
  ctx.fillStyle = "#b4b4b4";
  ctx.beginPath();
  const tx = s.headerW - 4, ty = s.headerH - 4, ts = Math.min(s.headerH, s.headerW) * 0.55;
  ctx.moveTo(tx, ty - ts);
  ctx.lineTo(tx, ty);
  ctx.lineTo(tx - ts, ty);
  ctx.closePath();
  ctx.fill();
  ctx.textBaseline = "alphabetic";
  void sheet;
}

function uniquePanes(panes: Pane[], axis: "x" | "y"): Pane[] {
  const seen = new Set<number>();
  return panes.filter((p) => {
    const k = axis === "x" ? p.x : p.y;
    if (seen.has(k)) return false;
    seen.add(k);
    return true;
  });
}

function colLabel(c: number): string {
  let s = "";
  let n = c + 1;
  while (n > 0) {
    const m = (n - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

/** The AutoFilter button inside a header cell (right edge, square, as tall as the row allows). */
export function filterButton(x: number, w: number, y: number, h: number, zoom: number): { x: number; y: number; w: number; h: number } {
  const size = Math.min(h - 2 * zoom, 17 * zoom);
  return { x: x + w - size - 1 * zoom, y: y + h - size - 1 * zoom, w: size, h: size };
}

/** Excel's default error checking: "Zahl als Text gespeichert" and formulas with an error result. */
export function errorIndicator(cell: Cell): "text-number" | "error" | null {
  if (cell.ignoreError) return null;
  if (cell.f === undefined && typeof cell.v === "string" && /^\s*[+-]?(\d{1,3}(\.\d{3})+|\d+)(,\d+)?\s*$/.test(cell.v)) return "text-number";
  if (cell.f !== undefined && cell.v !== null && typeof cell.v === "object") return "error";
  return null;
}
