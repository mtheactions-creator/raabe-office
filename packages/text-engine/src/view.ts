// DOM rendering. Every formatting segment of a line is one absolutely positioned element
// with real text, so the system renders glyphs like any native text. Its baseline is placed
// exactly with a per-font measurement (Fonts.baselineOffset). Only pages near the viewport
// carry line elements; edits update only lines whose content or position changed.
//
// Zoom scales the page container with a CSS transform; all geometry is computed at 100 %.

import { Fonts, PT } from "./fonts";
import { cmp, type Doc, type Pos } from "./doc";
import { LINE_BREAK, PAGE_BREAK, type Para } from "./model";
import type { Layout, Line, PlacedLine, Segment } from "./layout";
import type { ResolvedChar } from "./styles";

export interface Selection {
  anchor: Pos;
  head: Pos;
  /** Draw a caret at a line end rather than at the start of the next line. */
  headUpstream: boolean;
}

interface PageView {
  el: HTMLDivElement;
  selLayer: HTMLDivElement;
  textLayer: HTMLDivElement;
  lineEls: HTMLDivElement[];
  lineKeys: string[];
  live: boolean;
}

const SELECTION_COLOR = "#c6c6c6";
const GAP = 16;
export const MIN_ZOOM = 0.1;
export const MAX_ZOOM = 5;

/** Word's highlight colours (OOXML names). */
export const HIGHLIGHT: Record<string, string> = {
  yellow: "#ffff00", green: "#00ff00", cyan: "#00ffff", magenta: "#ff00ff", blue: "#0000ff", red: "#ff0000",
  darkBlue: "#000080", darkCyan: "#008080", darkGreen: "#008000", darkMagenta: "#800080", darkRed: "#800000",
  darkYellow: "#808000", darkGray: "#808080", lightGray: "#c0c0c0", black: "#000000",
};

const UNDERLINE: Record<string, string> = {
  single: "underline", words: "underline", double: "underline double", thick: "underline", dotted: "underline dotted",
  dottedHeavy: "underline dotted", dash: "underline dashed", dashedHeavy: "underline dashed", dashLong: "underline dashed",
  dotDash: "underline dashed", dotDotDash: "underline dotted", wave: "underline wavy", wavyDouble: "underline wavy",
};

let lineKeySeq = 0;
const lineIds = new WeakMap<Line, number>();
function lineId(l: Line): number {
  let id = lineIds.get(l);
  if (id === undefined) { id = ++lineKeySeq; lineIds.set(l, id); }
  return id;
}

export class View {
  readonly sizer: HTMLDivElement;
  readonly container: HTMLDivElement;
  readonly caret: HTMLDivElement;
  private pages: PageView[] = [];
  private blinkTimer = 0;
  private layoutRef: Layout | null = null;
  private docRef: Doc | null = null;
  private selRef: Selection | null = null;
  private visibleFrom = 0;
  private visibleTo = -1;
  private zoomFactor = 1;
  private marks = false;
  private pageW = 0;
  private pageH = 0;

  constructor(readonly workspace: HTMLElement, private fonts: Fonts) {
    this.sizer = document.createElement("div");
    this.sizer.style.cssText = "position:relative;margin:0 auto";
    this.container = document.createElement("div");
    this.container.style.cssText = "position:absolute;left:0;top:0;transform-origin:0 0";
    this.sizer.append(this.container);
    workspace.append(this.sizer);
    this.caret = document.createElement("div");
    this.caret.style.cssText = "position:absolute;z-index:3;pointer-events:none;background:#000";
    workspace.addEventListener("scroll", () => this.onScroll(), { passive: true });
    addEventListener("resize", () => this.onScroll());
  }

  private get stride(): number {
    return this.pageH + GAP;
  }

  pageTop(i: number): number {
    return GAP + i * this.stride;
  }

  private margins() {
    const p = this.layoutRef!.page;
    return { top: p.marginTop * PT, left: p.marginLeft * PT };
  }

  get zoom(): number {
    return this.zoomFactor;
  }

  get showMarks(): boolean {
    return this.marks;
  }

  setShowMarks(on: boolean): void {
    this.marks = on;
    for (const pv of this.pages) { pv.lineKeys = []; }
    this.onScroll(true);
  }

  setZoom(z: number): void {
    z = Math.min(MAX_ZOOM, Math.max(MIN_ZOOM, z));
    const ws = this.workspace;
    const anchor = ws.scrollTop / this.zoomFactor;
    this.zoomFactor = z;
    this.applySize();
    ws.scrollTop = anchor * z;
    this.onScroll(true);
  }

  pageWidthZoom(): number {
    return (this.workspace.clientWidth - 2 * GAP - 4) / (this.pageW || 794);
  }

  onePageZoom(): number {
    return Math.min(this.pageWidthZoom(), (this.workspace.clientHeight - 2 * GAP) / (this.pageH || 1123));
  }

  private applySize(): void {
    const z = this.zoomFactor;
    const w = this.pageW + 2 * GAP;
    const h = this.pageTop(this.pages.length);
    this.container.style.transform = z === 1 ? "" : `scale(${z})`;
    this.container.style.width = `${w}px`;
    this.container.style.height = `${h}px`;
    this.sizer.style.width = `${w * z}px`;
    this.sizer.style.height = `${h * z}px`;
  }

  // ----- rendering ----------------------------------------------------------------------

  render(doc: Doc, layout: Layout, sel: Selection): void {
    this.docRef = doc;
    this.layoutRef = layout;
    this.selRef = sel;
    const w = layout.page.width * PT, h = layout.page.height * PT;
    if (w !== this.pageW || h !== this.pageH) {
      this.pageW = w; this.pageH = h;
      for (const pv of this.pages) pv.el.remove();
      this.pages = [];
    }
    this.ensurePageCount(layout.pageCount);
    this.computeVisible();
    const byPage: PlacedLine[][] = [];
    for (const pl of layout.placed) {
      if (pl.page < this.visibleFrom || pl.page > this.visibleTo) continue;
      (byPage[pl.page] ??= []).push(pl);
    }
    for (let i = 0; i < this.pages.length; i++) {
      const pv = this.pages[i];
      if (i >= this.visibleFrom && i <= this.visibleTo) this.renderLines(pv, byPage[i] ?? [], doc, layout);
      else if (pv.live) this.clearPage(pv);
    }
    this.renderSelection();
  }

  private ensurePageCount(n: number): void {
    const before = this.pages.length;
    while (this.pages.length < n) {
      const el = document.createElement("div");
      el.className = "page";
      el.style.cssText = `position:absolute;margin:0;left:${GAP}px;top:${this.pageTop(this.pages.length)}px;width:${this.pageW}px;height:${this.pageH}px`;
      const selLayer = document.createElement("div");
      selLayer.style.cssText = "position:absolute;inset:0;z-index:0";
      const textLayer = document.createElement("div");
      textLayer.style.cssText = "position:absolute;inset:0;z-index:1";
      el.append(selLayer, textLayer);
      this.container.append(el);
      this.pages.push({ el, selLayer, textLayer, lineEls: [], lineKeys: [], live: false });
    }
    while (this.pages.length > n) this.pages.pop()!.el.remove();
    if (before !== n || !this.sizer.style.height) this.applySize();
  }

  private computeVisible(): void {
    const z = this.zoomFactor;
    const top = this.workspace.scrollTop / z;
    const bottom = top + this.workspace.clientHeight / z;
    const n = this.pages.length;
    this.visibleFrom = Math.max(0, Math.floor((top - GAP) / this.stride) - 1);
    this.visibleTo = Math.min(n - 1, Math.floor((bottom - GAP) / this.stride) + 1);
  }

  private onScroll(force = false): void {
    if (!this.layoutRef || !this.docRef || !this.selRef) return;
    const from = this.visibleFrom, to = this.visibleTo;
    this.computeVisible();
    if (force || from !== this.visibleFrom || to !== this.visibleTo) this.render(this.docRef, this.layoutRef, this.selRef);
  }

  private clearPage(pv: PageView): void {
    pv.textLayer.replaceChildren();
    pv.selLayer.replaceChildren();
    pv.lineEls = [];
    pv.lineKeys = [];
    pv.live = false;
  }

  private renderLines(pv: PageView, lines: PlacedLine[], doc: Doc, layout: Layout): void {
    pv.live = true;
    const m = this.margins();
    for (let k = 0; k < lines.length; k++) {
      const pl = lines[k];
      const para = doc.paras[pl.para];
      const line = layout.paras[pl.para].lines[pl.line];
      const isLast = pl.line === layout.paras[pl.para].lines.length - 1;
      const key = `${lineId(line)}|${pl.y}|${this.marks}|${isLast}`;
      if (pv.lineKeys[k] === key) continue;
      let el = pv.lineEls[k];
      if (!el) {
        el = document.createElement("div");
        el.style.cssText = "position:absolute;left:0;top:0;width:0;height:0";
        pv.textLayer.append(el);
        pv.lineEls[k] = el;
      }
      el.replaceChildren(...this.lineContent(para, line, m.left + line.x0, m.top + pl.y, isLast));
      pv.lineKeys[k] = key;
    }
    while (pv.lineEls.length > lines.length) {
      pv.lineEls.pop()!.remove();
      pv.lineKeys.pop();
    }
  }

  private spanStyle(c: ResolvedChar): string {
    const size = Fonts.renderSize(c) * PT;
    const lh = this.fonts.lineHeight(c.font, Fonts.renderSize(c)) * PT;
    let s = `position:absolute;white-space:pre;font:${this.fonts.cssFont(c)};line-height:${lh}px;` +
      `color:${c.color === "auto" ? "#000" : `#${c.color}`};font-kerning:${c.kern && c.size >= c.kern ? "normal" : "none"};font-variant-ligatures:none`;
    if (c.spacing) s += `;letter-spacing:${c.spacing * PT}px`;
    if (c.caps) s += ";text-transform:uppercase";
    if (c.smallCaps) s += ";font-variant-caps:small-caps";
    const deco: string[] = [];
    if (c.underline !== "none" && UNDERLINE[c.underline]) deco.push(UNDERLINE[c.underline]);
    if (c.strike || c.dstrike) deco.push("line-through");
    if (deco.length) {
      s += `;text-decoration:${deco.join(" ")}`;
      if (c.dstrike) s += ";text-decoration-style:double";
      if (c.underline === "thick") s += `;text-decoration-thickness:${Math.max(2, size / 10)}px`;
    }
    if (c.highlight !== "none" && HIGHLIGHT[c.highlight]) s += `;background:${HIGHLIGHT[c.highlight]}`;
    return s;
  }

  /** Vertical position of a run's baseline shift for super/subscript (px, positive = up). */
  private shift(c: ResolvedChar): number {
    // Word (from its PDF output, 8-36 pt): raised by about a third of the size, lowered by
    // about a tenth.
    if (c.vertAlign === "superscript") return c.size * PT * 0.335;
    if (c.vertAlign === "subscript") return -c.size * PT * 0.09;
    return 0;
  }

  private lineContent(para: Para, line: Line, left: number, top: number, isLast: boolean): HTMLElement[] {
    const out: HTMLElement[] = [];
    if (para.block) {
      const b = para.block;
      const box = document.createElement("div");
      box.style.cssText = `position:absolute;left:${left}px;top:${top}px;width:${line.segments[0].width}px;height:${line.height - 4}px;` +
        "box-sizing:border-box;border:1px dashed #9e9e9e;background:#fafafa;color:#616161;font:12px 'Segoe UI',sans-serif;padding:4px 8px;overflow:hidden;white-space:pre-wrap";
      box.textContent = `${b.label}\n${b.lines.join("\n")}`;
      out.push(box);
      return out;
    }
    const baseline = top + line.baseline;
    for (const seg of line.segments) out.push(...this.segment(para, line, seg, left, baseline, top));
    if (this.marks) {
      const c = line.segments.length ? line.segments[line.segments.length - 1].c : null;
      const endX = left + line.charX[line.end - line.start];
      if (isLast && !(para.text.endsWith(LINE_BREAK) && line.end === line.start && false)) {
        const markC = c ?? null;
        if (line.end === para.text.length) out.push(this.markGlyph("¶", endX, baseline, markC, para));
      }
    }
    return out;
  }

  private markGlyph(glyph: string, x: number, baseline: number, c: ResolvedChar | null, para: Para): HTMLElement {
    const el = document.createElement("span");
    const font = c ? this.fonts.cssFont(c) : `14.667px Calibri, sans-serif`;
    const lh = c ? this.fonts.lineHeight(c.font, Fonts.renderSize(c)) * PT : 19;
    const b = c ? this.fonts.baselineOffset(c, lh) : 15;
    el.style.cssText = `position:absolute;white-space:pre;left:${x}px;top:${baseline - b}px;font:${font};line-height:${lh}px;color:#000;pointer-events:none`;
    el.textContent = glyph;
    void para;
    return el;
  }

  private segment(para: Para, line: Line, seg: Segment, left: number, baseline: number, top: number): HTMLElement[] {
    const c = seg.c;
    const x = left + seg.x;
    const lh = this.fonts.lineHeight(c.font, Fonts.renderSize(c)) * PT;
    const b = this.fonts.baselineOffset(c, lh);
    const y = baseline - this.shift(c) - b;
    const out: HTMLElement[] = [];
    if (seg.kind === "tab") {
      if (this.marks) out.push(this.glyphAt("→", c, x + seg.width / 2 - 4, y, lh));
      if (c.highlight !== "none" && HIGHLIGHT[c.highlight]) {
        const hl = document.createElement("div");
        hl.style.cssText = `position:absolute;left:${x}px;top:${y}px;width:${seg.width}px;height:${lh}px;background:${HIGHLIGHT[c.highlight]}`;
        out.push(hl);
      }
      return out;
    }
    if (seg.kind === "break") {
      if (this.marks) {
        const ch = para.text[seg.start];
        if (ch === LINE_BREAK) out.push(this.glyphAt("↵", c, x, y, lh));
        if (ch === PAGE_BREAK) {
          const pb = document.createElement("div");
          pb.style.cssText = `position:absolute;left:${x}px;top:${top + line.height / 2 - 8}px;font:11px 'Segoe UI',sans-serif;color:#5b5b5b;white-space:pre;border-top:1px dotted #5b5b5b;padding-top:1px`;
          pb.textContent = "··········Seitenumbruch··········";
          out.push(pb);
        }
      }
      return out;
    }
    if (seg.kind === "object") {
      const o = seg.obj;
      if (o?.kind === "image" && o.width !== undefined && o.height !== undefined) {
        const box = document.createElement("div");
        box.style.cssText = `position:absolute;left:${x}px;top:${baseline - o.height * PT}px;width:${o.width * PT}px;height:${o.height * PT}px;` +
          "box-sizing:border-box;border:1px dashed #9e9e9e;background:#f3f3f3;color:#757575;font:11px 'Segoe UI',sans-serif;display:flex;align-items:center;justify-content:center";
        box.textContent = "Bild";
        out.push(box);
        return out;
      }
      const el = document.createElement("span");
      el.style.cssText = this.spanStyle(c) + `;left:${x}px;top:${y}px` + (o?.kind === "field" ? ";background:#e8e8e8" : "");
      el.textContent = o?.display ?? "";
      out.push(el);
      return out;
    }
    const el = document.createElement("span");
    let style = this.spanStyle(c) + `;left:${x}px;top:${y}px`;
    if (line.spaceExtra) style += `;word-spacing:${line.spaceExtra}px`;
    el.style.cssText = style;
    let text = para.text.slice(seg.start, seg.end);
    if (c.hidden) text = "";
    el.textContent = text.replace(/\u00ad/g, (_m, i: number) => (seg.start + i === line.end - 1 ? "-" : ""));
    out.push(el);
    if (this.marks) {
      // Middle dots on spaces.
      for (let i = seg.start; i < seg.end; i++) {
        if (para.text[i] !== " ") continue;
        const sx = left + line.charX[i - line.start];
        const w = line.charX[i + 1 - line.start] - line.charX[i - line.start];
        out.push(this.glyphAt("·", c, sx + w / 2 - 2, y, lh));
      }
    }
    return out;
  }

  private glyphAt(glyph: string, c: ResolvedChar, x: number, y: number, lh: number): HTMLElement {
    const el = document.createElement("span");
    el.style.cssText = `position:absolute;white-space:pre;left:${x}px;top:${y}px;font:${this.fonts.cssFont({ ...c, bold: false, italic: false })};line-height:${lh}px;color:#000;pointer-events:none`;
    el.textContent = glyph;
    return el;
  }

  /** Renders every page completely (no virtualisation, no selection) for printing. */
  renderForPrint(target: HTMLElement): void {
    const layout = this.layoutRef, doc = this.docRef;
    if (!layout || !doc) return;
    const m = this.margins();
    const marks = this.marks;
    this.marks = false;
    target.replaceChildren();
    const pages: HTMLDivElement[] = [];
    for (let i = 0; i < layout.pageCount; i++) {
      const el = document.createElement("div");
      el.className = "print-page";
      el.style.cssText = `position:relative;width:${this.pageW}px;height:${this.pageH}px;overflow:hidden;break-after:page`;
      pages.push(el);
      target.append(el);
    }
    for (const pl of layout.placed) {
      const line = layout.paras[pl.para].lines[pl.line];
      const isLast = pl.line === layout.paras[pl.para].lines.length - 1;
      pages[pl.page].append(...this.lineContent(doc.paras[pl.para], line, m.left + line.x0, m.top + pl.y, isLast));
    }
    this.marks = marks;
  }

  renderSelection(): void {
    const layout = this.layoutRef, doc = this.docRef, sel = this.selRef;
    if (!layout || !doc || !sel) return;
    const [from, to] = cmp(sel.anchor, sel.head) <= 0 ? [sel.anchor, sel.head] : [sel.head, sel.anchor];
    const empty = cmp(from, to) === 0;
    for (let i = this.visibleFrom; i <= this.visibleTo; i++) {
      const layer = this.pages[i]?.selLayer;
      if (layer?.firstChild) layer.replaceChildren();
    }
    if (!empty) {
      const m = this.margins();
      for (let p = from.p; p <= to.p; p++) {
        const plo = layout.paras[p];
        const base = layout.firstPlaced[p];
        for (let j = 0; j < plo.lines.length; j++) {
          const pl = layout.placed[base + j];
          if (pl.page < this.visibleFrom || pl.page > this.visibleTo) continue;
          const ln = plo.lines[j];
          const a = p === from.p ? Math.max(from.o, ln.start) : ln.start;
          const b = p === to.p ? Math.min(to.o, ln.end) : ln.end;
          const isLastLine = j === plo.lines.length - 1;
          const includesBreak = isLastLine && p < to.p;
          if (b < a || (a === b && !includesBreak)) continue;
          const x1 = ln.charX[a - ln.start];
          let x2 = ln.charX[b - ln.start];
          if (includesBreak) x2 += 7;
          const r = document.createElement("div");
          r.style.cssText = `position:absolute;background:${SELECTION_COLOR};left:${m.left + ln.x0 + x1}px;top:${m.top + pl.y}px;width:${Math.max(0, x2 - x1)}px;height:${ln.height}px`;
          this.pages[pl.page].selLayer.append(r);
        }
      }
    }
    this.renderCaret(empty);
  }

  private renderCaret(visible: boolean): void {
    if (!this.layoutRef || !this.docRef || !this.selRef) return;
    if (!visible) { this.caret.style.display = "none"; return; }
    const c = this.caretBox(this.selRef.head, this.selRef.headUpstream);
    const pv = this.pages[c.page];
    if (!pv) return;
    if (pv.el !== this.caret.parentElement) pv.el.append(this.caret);
    const px = (devicePixelRatio || 1) * this.zoomFactor;
    this.caret.style.display = "";
    this.caret.style.left = `${Math.round(c.x * px) / px}px`;
    this.caret.style.top = `${c.y}px`;
    this.caret.style.height = `${c.h}px`;
    this.caret.style.width = `${Math.max(1, Math.round(devicePixelRatio || 1)) / px}px`;
    this.restartBlink();
  }

  /**
   * Windows' caret blink (530 ms on, 530 ms off), restarted on every move. A plain timer:
   * restarting a CSS/Web animation on every keystroke delayed the frame by one refresh.
   */
  private restartBlink(): void {
    clearInterval(this.blinkTimer);
    this.caret.style.visibility = "visible";
    let on = true;
    this.blinkTimer = window.setInterval(() => {
      on = !on;
      this.caret.style.visibility = on ? "visible" : "hidden";
    }, 530);
  }

  stopBlink(): void {
    clearInterval(this.blinkTimer);
  }

  // ----- geometry -----------------------------------------------------------------------

  placedFor(pos: Pos, upstream = false): { pl: PlacedLine; index: number } {
    const layout = this.layoutRef!;
    const lines = layout.paras[pos.p].lines;
    let j = 0;
    while (j < lines.length - 1) {
      const end = lines[j].end;
      if (pos.o < end || (pos.o === end && upstream)) break;
      j++;
    }
    const index = layout.firstPlaced[pos.p] + j;
    return { pl: layout.placed[index], index };
  }

  lineOf(pl: PlacedLine): Line {
    return this.layoutRef!.paras[pl.para].lines[pl.line];
  }

  /** Caret box in (unzoomed) page coordinates. */
  caretBox(pos: Pos, upstream = false): { page: number; x: number; y: number; h: number } {
    const doc = this.docRef!;
    const { pl } = this.placedFor(pos, upstream);
    const line = this.lineOf(pl);
    const m = this.margins();
    const para = doc.paras[pos.p];
    const o = Math.max(line.start, Math.min(line.end, pos.o));
    const c = this.docRef!.styles.resolveChar(this.layoutRef!.paras[pos.p].rp.style, o > 0 && para.text.length ? propsBefore(para, o) : para.mark);
    const h = this.fonts.lineHeight(c.font, c.size) * PT;
    const a = this.fonts.ascent(c.font, c.size) * PT;
    return {
      page: pl.page,
      x: m.left + line.x0 + line.charX[o - line.start],
      y: m.top + pl.y + line.baseline - a,
      h,
    };
  }

  clientPoint(pos: Pos, upstream = false): { x: number; y: number; h: number } {
    const c = this.caretBox(pos, upstream);
    const rect = this.sizer.getBoundingClientRect();
    const z = this.zoomFactor;
    return { x: rect.left + (GAP + c.x) * z, y: rect.top + (this.pageTop(c.page) + c.y) * z, h: c.h * z };
  }

  hitLine(clientX: number, clientY: number): { index: number; x: number; inLeftMargin: boolean } {
    const layout = this.layoutRef!;
    const m = this.margins();
    const rect = this.sizer.getBoundingClientRect();
    const z = this.zoomFactor;
    const cy = (clientY - rect.top) / z;
    const page = Math.max(0, Math.min(layout.pageCount - 1, Math.floor((cy - GAP / 2) / this.stride)));
    const y = cy - this.pageTop(page) - m.top;
    const x = (clientX - rect.left) / z - GAP - m.left;
    let lo = 0, hi = layout.placed.length - 1;
    while (lo < hi) {
      const mid = (lo + hi) >> 1;
      if (layout.placed[mid].page < page) lo = mid + 1; else hi = mid;
    }
    // A point between two lines (in paragraph spacing) belongs to the line below, as in Word.
    let best = lo;
    for (let i = lo; i < layout.placed.length && layout.placed[i].page === page; i++) {
      const pl = layout.placed[i];
      best = i;
      if (y < pl.y + this.lineOf(pl).height) break;
    }
    return { index: best, x, inLeftMargin: x < -4 };
  }

  posInLine(index: number, x: number, charUnder = false): { pos: Pos; upstream: boolean } {
    const layout = this.layoutRef!;
    const pl = layout.placed[index];
    const line = this.lineOf(pl);
    const isLast = pl.line === layout.paras[pl.para].lines.length - 1;
    const rx = x - line.x0;
    let o = line.start;
    // Positions inside a break character are not reachable; stop before it.
    const doc = this.docRef!;
    const text = doc.paras[pl.para].text;
    let limit = line.end;
    if (limit > line.start && (text[limit - 1] === LINE_BREAK || text[limit - 1] === PAGE_BREAK)) limit--;
    while (o < limit) {
      const a = line.charX[o - line.start], b = line.charX[o + 1 - line.start];
      if ((charUnder ? b : (a + b) / 2) > rx) break;
      o++;
    }
    if (doc.paras[pl.para].block) return { pos: { p: pl.para, o: 0 }, upstream: false };
    return { pos: { p: pl.para, o }, upstream: !isLast && o === line.end };
  }

  scrollIntoView(pos: Pos, upstream: boolean): void {
    const c = this.caretBox(pos, upstream);
    const z = this.zoomFactor;
    const top = (this.pageTop(c.page) + c.y) * z;
    const h = c.h * z;
    const ws = this.workspace;
    const margin = 40;
    if (top < ws.scrollTop + margin) ws.scrollTop = top - margin;
    else if (top + h > ws.scrollTop + ws.clientHeight - margin) ws.scrollTop = top + h - ws.clientHeight + margin;
  }
}

function propsBefore(p: Para, o: number) {
  let pos = 0;
  for (const r of p.runs) {
    if (o - 1 < pos + r.len) return r.props;
    pos += r.len;
  }
  return p.mark;
}
