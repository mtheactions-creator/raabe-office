// Line breaking of one paragraph, following rules measured in Word 365 (2026-10-13,
// tools/feel-lab/run.py docx, decisions D-025):
// - a line reaches from the largest ascent to the largest descent of its fonts (see
//   Fonts.lineHeight); "multiple" spacing multiplies that and puts the extra space below the
//   text, "at least" puts it above, "exactly" fixes the height with the baseline at 80 %;
// - break after spaces, hyphens and soft hyphens; trailing spaces hang into the margin;
//   over-long words break anywhere; in "Blocksatz" a line may overflow by a quarter of the
//   width of its spaces, which are then narrowed;
// - default tab stops every 1.25 cm from the frame's left edge, plus custom and
//   hanging-indent stops; a list label is followed by its suffix (tab, space or nothing).
// Placing lines on pages (sections, columns, tables, footnotes …) is in flow.ts.
// All lengths here are CSS px at 100 % zoom (1 pt = 4/3 px).

import { Fonts, PT } from "./fonts";
import { COLUMN_BREAK, LINE_BREAK, OBJ, PAGE_BREAK, type CharProps, type InlineObject, type Para, type ParaProps } from "./model";
import { displayBullet, type ListLabel } from "./numbering";
import type { ResolvedChar, ResolvedPara, StyleSheet } from "./styles";

export { A4, type PageSetup } from "./model";

export interface LayoutContext {
  readonly styles: StyleSheet;
  readonly fonts: Fonts;
  /** Default tab stop distance in points (Word: 35.4 pt = 1.25 cm). */
  readonly defaultTab: number;
  /** Measures an equation given as LaTeX (points); provided by the host page. */
  readonly measureMath?: (latex: string, size: number, display: boolean) => { width: number; height: number; depth: number };
  /** Markup anzeigen: "simple" and "none" hide deleted text, "original" hides inserted text. */
  readonly markup?: "all" | "simple" | "none" | "original";
}

const hiddenVariant = new WeakMap<ResolvedChar, ResolvedChar>();
const plainVariant = new WeakMap<ResolvedChar, ResolvedChar>();
/** Formatting of tracked text as the markup mode shows it (hidden, or without marks). */
function forMarkup(c: ResolvedChar, markup: LayoutContext["markup"]): ResolvedChar {
  if (!c.rev || !markup || markup === "all") return c;
  const hide = ((markup === "none" || markup === "simple") && c.rev.kind === "del") || (markup === "original" && c.rev.kind === "ins");
  const cache = hide ? hiddenVariant : plainVariant;
  let v = cache.get(c);
  if (!v) { v = hide ? { ...c, hidden: true } : { ...c, rev: undefined }; cache.set(c, v); }
  return v;
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
  /** Text shown for an object segment when it differs from the object's own (page numbers). */
  readonly display?: string;
  /** Tab leader ("dot" …) drawn across a tab segment. */
  readonly leader?: string;
}

export interface Line {
  readonly start: number;
  /** End offset (exclusive); includes trailing spaces and a terminating break character. */
  readonly end: number;
  /** Left edge of the line's content relative to the frame's left edge (px). */
  readonly x0: number;
  /** Width of the content without trailing spaces (px). */
  readonly width: number;
  readonly height: number;
  /** Distance from the line top to the baseline (px). */
  readonly baseline: number;
  readonly segments: readonly Segment[];
  readonly spaceExtra: number;
  readonly pageBreakAfter: boolean;
  readonly columnBreakAfter?: boolean;
  /** x of every boundary start..end relative to x0 (px). */
  readonly charX: Float64Array;
  /** The list label, on a paragraph's first line: x relative to the frame. */
  readonly label?: { readonly text: string; readonly x: number; readonly c: ResolvedChar; readonly font?: string };
}

export interface ParaLayout {
  readonly rp: ResolvedPara;
  readonly lines: readonly Line[];
  /** Spacing before/after in px. */
  readonly before: number;
  readonly after: number;
  /** Width of the frame the paragraph was laid out for (px). */
  readonly width: number;
  /** Space taken by a top/bottom border and its distance (px), when the paragraph has one. */
  readonly borderTop: number;
  readonly borderBottom: number;
}

/** What a paragraph is laid out against. */
export interface ParaEnv {
  /** Frame width in px. */
  readonly width: number;
  readonly label?: ListLabel;
  /** Table style formatting for paragraphs in tables. */
  readonly tablePPr?: ParaProps;
  readonly tableRPr?: CharProps;
  /** Values of page fields: PAGE, NUMPAGES, SECTIONPAGES. */
  readonly fields?: Readonly<Record<string, string>>;
  /** Numbers shown for note references, by note story id. */
  readonly notes?: ReadonlyMap<string, string>;
  /**
   * Available horizontal room for a line starting `top` px below the paragraph's first line
   * (pictures with text wrapping). `push` moves the line further down. Uncached when set.
   */
  readonly box?: (top: number, height: number) => { left: number; right: number; push?: number } | null;
}

const isSpace = (ch: string) => ch === " ";

interface CharInfo {
  w: number;
  c: ResolvedChar;
  kind: SegmentKind;
  run: number;
  display?: string;
  /** Equation size (points). */
  math?: { width: number; height: number; depth: number };
}

/** The field type of a field code ("PAGE \* MERGEFORMAT" → "PAGE"). */
export function fieldType(instr: string | undefined): string {
  return (instr ?? "").trim().split(/\s+/)[0]?.toUpperCase() ?? "";
}

export class Layouter {
  private cache = new WeakMap<Para, Map<string, ParaLayout>>();

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

  /** Effective paragraph formatting in its context (table style, list level). */
  resolve(para: Para, env: ParaEnv): ResolvedPara {
    return this.ctx.styles.resolvePara(para.props, env.tablePPr || env.label ? { table: env.tablePPr, level: env.label?.level.pPr } : undefined);
  }

  paragraph(para: Para, env: ParaEnv): ParaLayout {
    if (env.box) return this.layoutPara(para, env);
    const key = `${env.width}|${env.label?.text ?? ""}|${env.label ? oidOf(env.label.level) : 0}|${oidOf(env.tablePPr)}|${oidOf(env.tableRPr)}|` +
      `${env.fields ? Object.values(env.fields).join(",") : ""}|${env.notes && para.objects?.some((o) => o.note) ? para.objects.map((o) => (o.note ? env.notes!.get(o.note) : "")).join(",") : ""}`;
    let byKey = this.cache.get(para);
    let hit = byKey?.get(key);
    if (!hit) {
      hit = this.layoutPara(para, env);
      if (!byKey) { byKey = new Map(); this.cache.set(para, byKey); }
      byKey.set(key, hit);
    }
    return hit;
  }

  private layoutPara(para: Para, env: ParaEnv): ParaLayout {
    const { styles, fonts } = this.ctx;
    const rp = this.resolve(para, env);
    const before = rp.before * PT, after = rp.after * PT;
    const width = env.width;
    const bt = rp.borders?.top && rp.borders.top.style !== "nil" && rp.borders.top.style !== "none" ? (rp.borders.top.size + rp.borders.top.space) * PT : 0;
    const bb = rp.borders?.bottom && rp.borders.bottom.style !== "nil" && rp.borders.bottom.style !== "none" ? (rp.borders.bottom.size + rp.borders.bottom.space) * PT : 0;
    const charOf = (props: CharProps) => forMarkup(styles.resolveChar(rp.style, props, env.tableRPr), this.ctx.markup);
    if (para.block) {
      const h = Math.max(para.block.height * PT, 10);
      const c = charOf(para.mark);
      const line: Line = {
        start: 0, end: para.text.length, x0: rp.indLeft * PT, width, height: h, baseline: h,
        segments: [{ start: 0, end: para.text.length, c, x: 0, width: width - rp.indLeft * PT, kind: "object" }],
        spaceExtra: 0, pageBreakAfter: false, charX: Float64Array.from([0, width - rp.indLeft * PT]),
      };
      return { rp, lines: [line], before, after, width, borderTop: 0, borderBottom: 0 };
    }

    // Resolve every character's formatting and width.
    const text = para.text;
    const infos: CharInfo[] = new Array(text.length);
    let objIndex = 0;
    const objs: (InlineObject | undefined)[] = new Array(text.length);
    // Equations: measured size (points). One alone in its paragraph is a display equation,
    // centred and set in display style, as Word does.
    const mathDims: ({ width: number; height: number; depth: number } | undefined)[] = [];
    const displayMath = text === OBJ && para.objects?.[0]?.kind === "math";
    let pos = 0;
    para.runs.forEach((run, ri) => {
      const c = charOf(run.props);
      for (let i = pos; i < pos + run.len; i++) {
        const ch = text[i];
        let w = 0;
        let kind: SegmentKind = "text";
        let display: string | undefined;
        if (ch === "\t") kind = "tab";
        else if (ch === LINE_BREAK || ch === PAGE_BREAK || ch === COLUMN_BREAK) kind = "break";
        else if (ch === OBJ) {
          kind = "object";
          const o = para.objects?.[objIndex++];
          objs[i] = o;
          if (o?.kind === "field" && env.fields) {
            const f = fieldType(o.instr);
            if (env.fields[f] !== undefined) display = env.fields[f];
          }
          if (o?.note && env.notes?.has(o.note)) display = env.notes.get(o.note);
          if (o?.kind === "math" && o.latex !== undefined && this.ctx.measureMath && o.width === undefined) {
            const m = this.ctx.measureMath(o.latex, c.size, displayMath);
            mathDims[i] = m;
            w = m.width * PT;
          } else if (o?.width !== undefined) w = o.width * PT;
          else if (o) for (const oc of display ?? o.display) w += fonts.advance(c, oc);
        } else if (c.hidden || ch === "\u00ad") {
          w = 0;
        } else {
          w = fonts.advance(c, ch);
          if (i > pos) w += fonts.kerning(c, text[i - 1], ch);
        }
        infos[i] = { w, c, kind, run: ri, display, math: mathDims[i] };
      }
      pos += run.len;
    });
    for (const i of infos.keys()) if (infos[i] === undefined) infos[i] = { w: 0, c: charOf(para.mark), kind: "text", run: -1 };

    const markC = charOf(para.mark);
    const lines: Line[] = [];
    let start = 0;
    let first = true;
    let top = 0;
    const tabStop = (absX: number, firstLine: boolean): { stop: number; align: string; leader?: string } => {
      const leftEdge = rp.indLeft * PT;
      const custom = rp.tabs.filter((t) => t.align !== "clear" && t.align !== "bar").map((t) => ({ stop: t.pos * PT, align: t.align as string, leader: t.leader }));
      if (firstLine && rp.indFirst < 0 && absX < leftEdge - 0.01) custom.push({ stop: leftEdge, align: "left", leader: undefined });
      custom.sort((a, b) => a.stop - b.stop);
      const next = custom.find((t) => t.stop > absX + 0.01);
      if (next) return next;
      const step = this.ctx.defaultTab * PT;
      const lastCustom = custom.length ? custom[custom.length - 1].stop : 0;
      let s = Math.max(lastCustom, 0);
      s = Math.floor((Math.max(absX, s) + 0.01) / step + 1) * step;
      return { stop: s, align: "left" };
    };

    // The list label on the first line: number at the first-line indent, then its suffix.
    let label: Line["label"];
    let labelEnd = 0;
    /** Formatting the label's line height comes from (the level's own font, e.g. Symbol). */
    let labelMetrics: ResolvedChar | undefined;
    if (env.label && env.label.text !== undefined) {
      const lc = charOf(env.label.props);
      const bullet = env.label.level.format === "bullet";
      const shown = bullet ? displayBullet(env.label.text, env.label.props.font ?? lc.font) : { text: env.label.text, font: undefined };
      const c = shown.font && shown.font !== lc.font ? { ...lc, font: shown.font } : lc;
      let lw = 0;
      for (const ch of shown.text) lw += fonts.advance(c, ch);
      const at = (rp.indLeft + rp.indFirst) * PT;
      const lx = env.label.level.align === "right" ? at - lw : env.label.level.align === "center" ? at - lw / 2 : at;
      const end = lx + lw;
      const suffix = env.label.level.suffix;
      labelEnd = suffix === "nothing" ? end : suffix === "space" ? end + fonts.advance(c, " ") : tabStop(end, true).stop;
      label = { text: shown.text, x: lx, c, font: shown.font };
      labelMetrics = lc;
    }

    while (true) {
      let lineLeft = label && first ? labelEnd : (rp.indLeft + (first ? rp.indFirst : 0)) * PT;
      let lineRight = width - rp.indRight * PT;
      let push = 0;
      if (env.box) {
        const est = fonts.lineHeight(markC.font, markC.size) * PT * (rp.lineRule === "auto" ? rp.line : 1);
        const b = env.box(top, est);
        if (b) {
          lineLeft = Math.max(lineLeft, b.left);
          lineRight = Math.min(lineRight, b.right);
          push = b.push ?? 0;
        }
      }
      const avail = Math.max(10, lineRight - lineLeft);
      let x = 0;
      let lastBreak = -1;
      let end = text.length;
      let forcedEnd = false;
      let pageBreakAfter = false;
      let columnBreakAfter = false;
      const squeeze = rp.align === "both" ? 0.25 : 0;
      let spaceW = 0;
      for (let i = start; i < text.length; i++) {
        const ch = text[i];
        const info = infos[i];
        if (info.kind === "break") {
          end = i + 1;
          forcedEnd = true;
          pageBreakAfter = ch === PAGE_BREAK;
          columnBreakAfter = ch === COLUMN_BREAK;
          break;
        }
        if (info.kind === "tab") {
          const absX = lineLeft + x;
          const t = tabStop(absX, first && !label);
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
          (info as CharInfo & { leader?: string }).leader = t.leader;
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
      const line = this.buildLine(text, infos, objs, start, end, lineLeft, avail, rp, markC, forcedEnd || end >= text.length, pageBreakAfter, push,
        first && labelMetrics ? labelMetrics : undefined);
      lines.push(columnBreakAfter || (first && label) ? { ...line, columnBreakAfter: columnBreakAfter || undefined, label: first ? label : undefined } : line);
      top += line.height;
      first = false;
      if (end >= text.length) {
        // A trailing line or column break still produces an (empty) last line, as in Word
        // (after a column break it starts the next column; a page break keeps the mark on its line).
        if (forcedEnd && text.length && (text[text.length - 1] === LINE_BREAK || text[text.length - 1] === COLUMN_BREAK)) {
          lines.push(this.buildLine(text, infos, objs, text.length, text.length, rp.indLeft * PT, width - rp.indLeft * PT - rp.indRight * PT, rp, markC, true, false, 0));
        }
        break;
      }
      start = end;
    }
    return { rp, lines, before, after, width, borderTop: bt, borderBottom: bb };
  }

  private buildLine(text: string, infos: CharInfo[], objs: (InlineObject | undefined)[], start: number, end: number,
    lineLeft: number, avail: number, rp: ResolvedPara, markC: ResolvedChar, last: boolean, pageBreakAfter: boolean, push: number,
    label?: ResolvedChar): Line {
    const { fonts } = this.ctx;
    // Content end without trailing spaces and break characters.
    let contentEnd = end;
    while (contentEnd > start && (isSpace(text[contentEnd - 1]) || infos[contentEnd - 1].kind === "break")) contentEnd--;
    let width = 0;
    for (let i = start; i < contentEnd; i++) width += infos[i].w;

    // Vertical metrics: the largest font on the line (nominal size; super/subscript do not grow
    // lines). Pictures and equations stand on the baseline. Measured: "Mehrfach" adds its extra
    // space for the text only (a 60 pt picture in an 1.08 line adds 8 % of the font's height,
    // not of the picture's), and a list label's font raises the line but does not deepen it.
    let asc = 0, desc = 0, objAsc = 0, objDesc = 0;
    const consider = (c: ResolvedChar) => {
      const h = fonts.lineHeight(c.font, c.size) * PT;
      const a = fonts.ascent(c.font, c.size) * PT;
      asc = Math.max(asc, a);
      desc = Math.max(desc, h - a);
    };
    let anyText = false;
    for (let i = start; i < end; i++) {
      const o = objs[i];
      const md = infos[i].math;
      if (infos[i].kind === "object" && (md || o?.height !== undefined)) {
        const d = (md ? md.depth : o!.depth ?? 0) * PT;
        objAsc = Math.max(objAsc, (md ? md.height : o!.height!) * PT - d);
        objDesc = Math.max(objDesc, d);
        continue;
      }
      if (infos[i].c.hidden) continue;
      consider(infos[i].c);
      anyText = true;
    }
    if (!anyText) consider(markC);
    if (label) asc = Math.max(asc, fonts.ascent(label.font, label.size) * PT);
    const textNatural = asc + desc;
    const natural = Math.max(asc, objAsc) + Math.max(desc, objDesc);
    let height = natural, baseline = Math.max(asc, objAsc);
    if (rp.lineRule === "auto") height = natural + textNatural * (rp.line - 1);
    // "Genau": Word puts the baseline at 80 % of the line, whatever the font.
    else if (rp.lineRule === "exact") { height = rp.line * PT; baseline = height * 0.8; }
    else { height = Math.max(rp.line * PT, natural); baseline = Math.max(asc, objAsc) + (height - natural); }
    height += push;
    baseline += push;

    // Alignment.
    let x0 = lineLeft;
    let spaceExtra = 0;
    if (rp.align === "center" || (text === OBJ && objs[0]?.kind === "math")) x0 += (avail - width) / 2;
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
        const info = infos[segStart] as CharInfo & { leader?: string };
        segments.push({
          start: segStart, end: i, c: info.c, x: segX, width: x - segX, kind: info.kind,
          obj: info.kind === "object" ? objs[segStart] : undefined, display: info.display,
          leader: info.kind === "tab" && info.leader && info.leader !== "none" ? info.leader : undefined,
        });
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
}

let objSeq = 0;
const objIds = new WeakMap<object, number>();
function oidOf(o: object | undefined): number {
  if (!o) return 0;
  let id = objIds.get(o);
  if (id === undefined) { id = ++objSeq; objIds.set(o, id); }
  return id;
}
