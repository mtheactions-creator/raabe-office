// Line breaking and pagination, following rules measured in Word 365 (2026-10-13,
// tools/feel-lab/probe_word_lines.py):
// - a line reaches from the largest ascent to the largest descent of its fonts (see
//   Fonts.lineHeight; checked against Word's PDF output, tools/feel-lab/run.py docx);
//   "multiple" spacing multiplies that and puts the extra space below the text, "at least"
//   puts it above, "exactly" fixes the height with the baseline at 80 %;
// - the space between two paragraphs is the larger of "after" and "before", not the sum;
//   "before" counts at the top of the document and is dropped at the top of later pages;
// - break after spaces, hyphens and soft hyphens; trailing spaces hang into the margin;
//   over-long words break anywhere;
// - default tab stops every 1.25 cm from the margin, plus custom and hanging-indent stops;
// - widow/orphan control, keep with next (chains), keep lines together, page break before.
// All lengths here are CSS px at 100 % zoom (1 pt = 4/3 px).

import { Fonts, PT } from "./fonts";
import { LINE_BREAK, OBJ, PAGE_BREAK, type InlineObject, type Para } from "./model";
import type { ResolvedChar, ResolvedPara, StyleSheet } from "./styles";

export interface PageSetup {
  /** Points. */
  readonly width: number;
  readonly height: number;
  readonly marginTop: number;
  readonly marginBottom: number;
  readonly marginLeft: number;
  readonly marginRight: number;
}

export const A4: PageSetup = { width: 595.3, height: 841.9, marginTop: 70.85, marginBottom: 56.7, marginLeft: 70.85, marginRight: 70.85 };

export interface LayoutContext {
  readonly styles: StyleSheet;
  readonly fonts: Fonts;
  readonly page: PageSetup;
  /** Default tab stop distance in points (Word: 35.4 pt = 1.25 cm). */
  readonly defaultTab: number;
}

export type SegmentKind = "text" | "tab" | "object" | "break";

export interface Segment {
  readonly start: number;
  readonly end: number;
  readonly c: ResolvedChar;
  /** Left edge relative to the line's x0 (px), justification included. */
  readonly x: number;
  readonly width: number;
  readonly kind: SegmentKind;
  readonly obj?: InlineObject;
}

export interface Line {
  readonly start: number;
  /** End offset (exclusive); includes trailing spaces and a terminating break character. */
  readonly end: number;
  /** Left edge of the line's content relative to the left margin (px). */
  readonly x0: number;
  /** Width of the content without trailing spaces (px). */
  readonly width: number;
  readonly height: number;
  /** Distance from the line top to the baseline (px). */
  readonly baseline: number;
  readonly segments: readonly Segment[];
  readonly spaceExtra: number;
  readonly pageBreakAfter: boolean;
  /** x of every boundary start..end relative to x0 (px). */
  readonly charX: Float64Array;
}

export interface ParaLayout {
  readonly rp: ResolvedPara;
  readonly lines: readonly Line[];
  /** Spacing before/after in px. */
  readonly before: number;
  readonly after: number;
}

export interface PlacedLine {
  para: number;
  line: number;
  page: number;
  /** Top of the line box relative to the top of the page's content area (px). */
  y: number;
}

export interface Layout {
  readonly paras: readonly ParaLayout[];
  readonly placed: readonly PlacedLine[];
  /** Index into `placed` of each paragraph's first line. */
  readonly firstPlaced: readonly number[];
  readonly pageCount: number;
  readonly page: PageSetup;
}

const isSpace = (ch: string) => ch === " ";

// ----- one paragraph ----------------------------------------------------------------------

interface CharInfo {
  w: number;
  c: ResolvedChar;
  kind: SegmentKind;
  run: number;
}

export class Layouter {
  private cache = new WeakMap<Para, ParaLayout>();
  private ctxRef: LayoutContext | null = null;

  constructor(private ctx: LayoutContext) {}

  setContext(ctx: LayoutContext): void {
    this.ctx = ctx;
    this.cache = new WeakMap();
  }

  /** Drops cached paragraphs (fonts or styles changed). */
  invalidate(): void {
    this.cache = new WeakMap();
  }

  get context(): LayoutContext {
    return this.ctx;
  }

  contentWidth(): number {
    const p = this.ctx.page;
    return (p.width - p.marginLeft - p.marginRight) * PT;
  }

  contentHeight(): number {
    const p = this.ctx.page;
    return (p.height - p.marginTop - p.marginBottom) * PT;
  }

  paragraph(para: Para): ParaLayout {
    let hit = this.cache.get(para);
    if (!hit || this.ctxRef !== this.ctx) {
      if (this.ctxRef !== this.ctx) { this.cache = new WeakMap(); this.ctxRef = this.ctx; }
      hit = this.layoutPara(para);
      this.cache.set(para, hit);
    }
    return hit;
  }

  private layoutPara(para: Para): ParaLayout {
    const { styles, fonts } = this.ctx;
    const rp = styles.resolvePara(para.props);
    const before = rp.before * PT, after = rp.after * PT;
    if (para.block) {
      const h = Math.max(para.block.height * PT, 10);
      const c = styles.resolveChar(rp.style, para.mark);
      const line: Line = {
        start: 0, end: para.text.length, x0: rp.indLeft * PT, width: this.contentWidth(), height: h, baseline: h,
        segments: [{ start: 0, end: para.text.length, c, x: 0, width: this.contentWidth() - rp.indLeft * PT, kind: "object" }],
        spaceExtra: 0, pageBreakAfter: false, charX: Float64Array.from([0, this.contentWidth() - rp.indLeft * PT]),
      };
      return { rp, lines: [line], before, after };
    }

    // Resolve every character's formatting and width.
    const text = para.text;
    const infos: CharInfo[] = new Array(text.length);
    let objIndex = 0;
    const objs: (InlineObject | undefined)[] = new Array(text.length);
    let pos = 0;
    para.runs.forEach((run, ri) => {
      const c = styles.resolveChar(rp.style, run.props);
      for (let i = pos; i < pos + run.len; i++) {
        const ch = text[i];
        let w = 0;
        let kind: SegmentKind = "text";
        if (ch === "\t") kind = "tab";
        else if (ch === LINE_BREAK || ch === PAGE_BREAK) kind = "break";
        else if (ch === OBJ) {
          kind = "object";
          const o = para.objects?.[objIndex++];
          objs[i] = o;
          if (o?.width !== undefined) w = o.width * PT;
          else if (o) for (const oc of o.display) w += fonts.advance(c, oc);
        } else if (c.hidden || ch === "\u00ad") {
          w = 0;
        } else {
          w = fonts.advance(c, ch);
          if (i > pos) w += fonts.kerning(c, text[i - 1], ch);
        }
        infos[i] = { w, c, kind, run: ri };
      }
      pos += run.len;
    });
    for (const i of infos.keys()) if (infos[i] === undefined) infos[i] = { w: 0, c: styles.resolveChar(rp.style, para.mark), kind: "text", run: -1 };

    const markC = styles.resolveChar(rp.style, para.mark);
    const width = this.contentWidth();
    const lines: Line[] = [];
    let start = 0;
    let first = true;
    const tabStop = (absX: number, firstLine: boolean): { stop: number; align: string } => {
      const leftEdge = rp.indLeft * PT;
      const custom = rp.tabs.filter((t) => t.align !== "clear" && t.align !== "bar").map((t) => ({ stop: t.pos * PT, align: t.align }));
      if (firstLine && rp.indFirst < 0 && absX < leftEdge - 0.01) custom.push({ stop: leftEdge, align: "left" });
      custom.sort((a, b) => a.stop - b.stop);
      const next = custom.find((t) => t.stop > absX + 0.01);
      if (next) return next;
      const step = this.ctx.defaultTab * PT;
      const lastCustom = custom.length ? custom[custom.length - 1].stop : 0;
      let s = Math.max(lastCustom, 0);
      s = Math.floor((Math.max(absX, s) + 0.01) / step + 1) * step;
      return { stop: s, align: "left" };
    };

    while (true) {
      const lineLeft = (rp.indLeft + (first ? rp.indFirst : 0)) * PT;
      const avail = width - lineLeft - rp.indRight * PT;
      let x = 0;
      let lastBreak = -1;
      let end = text.length;
      let forcedEnd = false;
      let pageBreakAfter = false;
      // Blocksatz: Word lets a line overflow by up to a quarter of the width of its spaces and
      // then narrows the spaces (measured: the limit is 25 % for 14 to 30 spaces per line).
      const squeeze = rp.align === "both" ? 0.25 : 0;
      let spaceW = 0;
      for (let i = start; i < text.length; i++) {
        const ch = text[i];
        const info = infos[i];
        if (info.kind === "break") {
          end = i + 1;
          forcedEnd = true;
          pageBreakAfter = ch === PAGE_BREAK;
          break;
        }
        if (info.kind === "tab") {
          const absX = lineLeft + x;
          const t = tabStop(absX, first);
          let w = t.stop - absX;
          if (t.align !== "left") {
            // Right/centre/decimal stops align the text that follows (up to the next tab).
            let seg = 0;
            for (let j = i + 1; j < text.length && infos[j].kind !== "tab" && infos[j].kind !== "break"; j++) {
              if (t.align === "decimal" && (text[j] === "," || text[j] === ".")) break;
              seg += infos[j].w;
            }
            w = t.align === "center" ? w - seg / 2 : w - seg;
          }
          info.w = Math.max(0, w);
          x += info.w;
          lastBreak = i + 1;
          continue;
        }
        if (isSpace(ch)) { x += info.w; spaceW += info.w; lastBreak = i + 1; continue; }
        if (x + info.w > avail + 0.01 + squeeze * spaceW && i > start) {
          end = lastBreak > start ? lastBreak : i;
          break;
        }
        x += info.w;
        if ((ch === "-" || ch === "\u00ad") && i + 1 < text.length && !isSpace(text[i + 1])) lastBreak = i + 1;
      }
      // A soft hyphen at a line end shows a hyphen.
      if (!forcedEnd && end < text.length && text[end - 1] === "\u00ad") infos[end - 1].w = fonts.advance(infos[end - 1].c, "-");
      lines.push(this.buildLine(text, infos, objs, start, end, lineLeft, avail, rp, markC, forcedEnd || end >= text.length, pageBreakAfter));
      first = false;
      if (end >= text.length) {
        // A trailing line break still produces an (empty) last line, as in Word.
        if (forcedEnd && text.length && text[text.length - 1] === LINE_BREAK) {
          lines.push(this.buildLine(text, infos, objs, text.length, text.length, rp.indLeft * PT, width - rp.indLeft * PT - rp.indRight * PT, rp, markC, true, false));
        }
        break;
      }
      start = end;
    }
    return { rp, lines, before, after };
  }

  private buildLine(text: string, infos: CharInfo[], objs: (InlineObject | undefined)[], start: number, end: number,
    lineLeft: number, avail: number, rp: ResolvedPara, markC: ResolvedChar, last: boolean, pageBreakAfter: boolean): Line {
    const { fonts } = this.ctx;
    // Content end without trailing spaces and break characters.
    let contentEnd = end;
    while (contentEnd > start && (isSpace(text[contentEnd - 1]) || infos[contentEnd - 1].kind === "break")) contentEnd--;
    let width = 0;
    for (let i = start; i < contentEnd; i++) width += infos[i].w;

    // Vertical metrics: the largest font on the line (nominal size; super/subscript do not grow lines).
    let asc = 0, desc = 0;
    const consider = (c: ResolvedChar) => {
      const h = fonts.lineHeight(c.font, c.size) * PT;
      const a = fonts.ascent(c.font, c.size) * PT;
      asc = Math.max(asc, a);
      desc = Math.max(desc, h - a);
    };
    let any = false;
    for (let i = start; i < end; i++) {
      const o = objs[i];
      if (infos[i].kind === "object" && o?.height !== undefined) { asc = Math.max(asc, o.height * PT); any = true; continue; }
      if (infos[i].c.hidden) continue;
      consider(infos[i].c);
      any = true;
    }
    if (!any) consider(markC);
    const natural = asc + desc;
    let height = natural, baseline = asc;
    if (rp.lineRule === "auto") height = natural * rp.line;
    // "Genau": Word puts the baseline at 80 % of the line, whatever the font.
    else if (rp.lineRule === "exact") { height = rp.line * PT; baseline = height * 0.8; }
    else { height = Math.max(rp.line * PT, natural); baseline = asc + (height - natural); }

    // Alignment.
    let x0 = lineLeft;
    let spaceExtra = 0;
    if (rp.align === "center") x0 += (avail - width) / 2;
    else if (rp.align === "right") x0 += avail - width;
    else if (rp.align === "both") {
      // Spaces widen to fill the line (not on the last line) or narrow when the line was
      // allowed to overflow (any line).
      let spaces = 0;
      for (let i = start; i < contentEnd; i++) if (isSpace(text[i])) spaces++;
      if (spaces > 0 && (width > avail || (!last && avail > width))) spaceExtra = (avail - width) / spaces;
    }

    // Character positions and segments.
    const charX = new Float64Array(end - start + 1);
    const segments: Segment[] = [];
    let x = 0;
    let segStart = start, segX = 0;
    const flush = (i: number) => {
      if (i > segStart) {
        const info = infos[segStart];
        segments.push({ start: segStart, end: i, c: info.c, x: segX, width: x - segX, kind: info.kind, obj: info.kind === "object" ? objs[segStart] : undefined });
      }
      segStart = i;
      segX = x;
    };
    for (let i = start; i < end; i++) {
      const info = infos[i];
      const prev = i > start ? infos[i - 1] : null;
      if (i > start && (info.run !== prev!.run || info.kind !== "text" || prev!.kind !== "text")) flush(i);
      charX[i - start] = x;
      x += info.w;
      if (spaceExtra && isSpace(text[i]) && i < contentEnd) x += spaceExtra;
    }
    flush(end);
    charX[end - start] = x;
    return { start, end, x0, width: rp.align === "both" && spaceExtra ? avail : width, height, baseline, segments, spaceExtra, pageBreakAfter, charX };
  }

  // ----- pagination -----------------------------------------------------------------------

  layout(paras: readonly Para[]): Layout {
    const pls = paras.map((p) => this.paragraph(p));
    const contentH = this.contentHeight();
    const placed: PlacedLine[] = [];
    const firstPlaced: number[] = new Array(paras.length);
    let page = 0;
    let y = 0;
    let prev: ParaLayout | null = null;

    const lineHeights = (pl: ParaLayout) => pl.lines.map((l) => l.height);
    const sum = (a: number[], from: number, to: number) => { let s = 0; for (let i = from; i < to; i++) s += a[i]; return s; };
    const gapBetween = (a: ParaLayout | null, b: ParaLayout): number => {
      if (!a) return b.before;
      const same = a.rp.style === b.rp.style;
      const after = a.rp.contextual && same ? 0 : a.after;
      const before = b.rp.contextual && same ? 0 : b.before;
      return Math.max(after, before);
    };
    /** Lines of `pl` (from line `from`) that fit when starting at y, honouring widow/orphan control. */
    const fitting = (pl: ParaLayout, from: number, top: number): number => {
      const hs = lineHeights(pl);
      const n = hs.length - from;
      let k = 0, acc = top;
      while (k < n && acc + hs[from + k] <= contentH + 0.01) { acc += hs[from + k]; k++; }
      if (pl.rp.keepLines && from === 0 && k < n) return 0;
      if (pl.rp.widow) {
        if (n > 1 && k === 1 && from === 0) k = 0;
        if (n >= 3 && n - k === 1 && k > 1) k -= 1;
        if (n === 2 && k === 1 && from === 0) k = 0;
      }
      return k;
    };
    /** Height a keep-with-next chain starting at paragraph i needs on one page. */
    const chainHeight = (i: number): number => {
      let h = 0;
      let j = i;
      while (j < pls.length) {
        const pl = pls[j];
        const hs = lineHeights(pl);
        if (j > i) h += gapBetween(pls[j - 1], pl);
        if (pl.rp.keepNext && j + 1 < pls.length) { h += sum(hs, 0, hs.length); j++; continue; }
        const need = pl.rp.keepLines ? hs.length : Math.min(hs.length, pl.rp.widow && hs.length > 1 ? 2 : 1);
        h += sum(hs, 0, need);
        break;
      }
      return h;
    };

    for (let i = 0; i < pls.length; i++) {
      const pl = pls[i];
      const n = pl.lines.length;
      if (pl.rp.pageBreakBefore && y > 0) { page++; y = 0; }
      // "Before" counts at the top of the document but is dropped at the top of every later
      // page, whatever started it (measured: natural, manual and "page break before").
      let gap = y === 0 ? (page === 0 ? pl.before : 0) : gapBetween(prev, pl);

      let k = fitting(pl, 0, y + gap);
      if (k === n && pl.rp.keepNext && i + 1 < pls.length && y > 0) {
        // Keep with next: this paragraph and the start of the following chain share a page.
        const need = sum(lineHeights(pl), 0, n) + gapBetween(pl, pls[i + 1]) + chainHeight(i + 1);
        if (y + gap + need > contentH + 0.01 && chainHeight(i) <= contentH) k = 0;
      }
      if (k === 0 && y > 0) { page++; y = 0; gap = 0; k = fitting(pl, 0, 0); }
      if (k === 0) k = 1; // taller than a page even on its own: split anyway
      y += gap;
      firstPlaced[i] = placed.length;
      let j = 0;
      while (j < n) {
        const take = j === 0 ? k : Math.max(1, fitting(pl, j, y));
        for (let t = 0; t < take && j < n; t++, j++) {
          const line = pl.lines[j];
          placed.push({ para: i, line: j, page, y });
          y += line.height;
          if (line.pageBreakAfter) { page++; y = 0; j++; break; }
        }
        if (j < n && y > 0 && !(pl.lines[j - 1]?.pageBreakAfter)) { page++; y = 0; }
        if (j < n && pl.lines[j - 1]?.pageBreakAfter) continue;
      }
      prev = pl;
    }
    return { paras: pls, placed, firstPlaced, pageCount: page + 1, page: this.ctx.page };
  }
}
