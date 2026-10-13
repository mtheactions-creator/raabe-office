// Line breaking and pagination, following Word's rules for the subset
// we support: break after spaces and hyphens, trailing spaces hang into the margin,
// over-long words break anywhere, widow/orphan control and "keep with next".

import { CONTENT_HEIGHT, CONTENT_WIDTH, PT, STYLES, fontCss, lineHeightPx, type ParaStyle, type StyleId } from "./geometry";
import type { Para } from "./doc";

export interface Line {
  /** [start, end) within the paragraph; end includes trailing spaces. */
  start: number;
  end: number;
}

export interface PlacedLine {
  para: number;
  line: number;
  page: number;
  /** Top of the line box, relative to the top of the page's content area. */
  y: number;
}

export interface Layout {
  lines: Line[][];
  placed: PlacedLine[];
  /** Index into `placed` of each paragraph's first line. */
  firstPlaced: number[];
  pageCount: number;
}

// ----- measuring ------------------------------------------------------------------------

const ctx = document.createElement("canvas").getContext("2d")!;
const advanceCache = new Map<StyleId, Map<string, number>>();

export function advance(style: StyleId, ch: string): number {
  let m = advanceCache.get(style);
  if (!m) { m = new Map(); advanceCache.set(style, m); }
  let w = m.get(ch);
  if (w === undefined) {
    ctx.font = fontCss(STYLES[style]);
    ctx.fontKerning = "none";
    w = ctx.measureText(ch).width;
    m.set(ch, w);
  }
  return w;
}

/** Width of text[from, to) in the given style (no kerning, no ligatures, as in Word). */
export function widthOf(style: StyleId, text: string, from: number, to: number): number {
  let w = 0;
  for (let i = from; i < to; i++) w += advance(style, text[i]);
  return w;
}

export function styleOf(id: StyleId): ParaStyle {
  return STYLES[id];
}

// ----- line breaking --------------------------------------------------------------------

const lineCache = new WeakMap<Para, Line[]>();

function isSpace(ch: string): boolean {
  return ch === " " || ch === String.fromCharCode(0xa0);
}

function breakLines(para: Para): Line[] {
  const { text, style } = para;
  const lines: Line[] = [];
  let start = 0;
  let x = 0;
  let lastBreak = -1;
  for (let i = 0; i < text.length; i++) {
    const ch = text[i];
    const w = advance(style, ch);
    if (ch === " ") { x += w; lastBreak = i + 1; continue; }
    if (x + w > CONTENT_WIDTH && i > start) {
      const at = lastBreak > start ? lastBreak : i;
      lines.push({ start, end: at });
      start = at;
      x = widthOf(style, text, start, i);
      lastBreak = -1;
    }
    x += w;
    if (ch === "-" && !isSpace(text[i + 1] ?? " ")) lastBreak = i + 1;
  }
  lines.push({ start, end: text.length });
  return lines;
}

export function linesOf(para: Para): Line[] {
  let l = lineCache.get(para);
  if (!l) { l = breakLines(para); lineCache.set(para, l); }
  return l;
}

// ----- pagination -----------------------------------------------------------------------

export function layout(paras: readonly Para[]): Layout {
  const lines = paras.map(linesOf);
  const placed: PlacedLine[] = [];
  const firstPlaced: number[] = new Array(paras.length);
  let page = 0;
  let y = 0;

  const fits = (count: number, from: number, lh: number) => from + count * lh <= CONTENT_HEIGHT + 0.01;

  for (let i = 0; i < paras.length; i++) {
    const s = styleOf(paras[i].style);
    const lh = lineHeightPx(s);
    const n = lines[i].length;
    const before = y > 0 ? s.spaceBeforePt * PT : 0;

    // How many lines of this paragraph fit on the current page?
    let k = Math.max(0, Math.floor((CONTENT_HEIGHT + 0.01 - y - before) / lh));
    if (k > n) k = n;
    if (n > 1 && k === 1) k = 0;            // orphan: never leave a single first line behind
    if (n >= 3 && n - k === 1) k -= 1;      // widow: never carry a single last line over
    if (s.keepWithNext && k === n && i + 1 < paras.length) {
      // Keep with next: the following paragraph's first lines must fit on the same page.
      const ns = styleOf(paras[i + 1].style);
      const need = Math.min(lines[i + 1].length, 2);
      const yAfter = y + before + n * lh + s.spaceAfterPt * PT + ns.spaceBeforePt * PT;
      if (!fits(need, yAfter, lineHeightPx(ns))) k = 0;
    }
    if (k === 0 && y > 0) { page++; y = 0; }
    else y += before;

    firstPlaced[i] = placed.length;
    for (let j = 0; j < n; j++) {
      const splitHere = j === k && k > 0 && k < n;
      if (y > 0 && (splitHere || !fits(1, y, lh))) { page++; y = 0; }
      placed.push({ para: i, line: j, page, y });
      y += lh;
    }
    y += s.spaceAfterPt * PT;
  }
  return { lines, placed, firstPlaced, pageCount: page + 1 };
}
