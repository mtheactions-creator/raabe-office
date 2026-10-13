// DOM rendering for prototype B. Every visual line is one absolutely positioned element
// containing real text, so the system renders glyphs exactly as in any native text
// (subpixel positioning, ClearType-style smoothing). Only pages near the viewport carry
// line elements; edits update just the lines whose text or position changed.

import { PAGE, PT, fontCss, lineHeightPx } from "../shared/geometry";
import { cmp, type Doc, type Pos } from "./doc";
import { advance, styleOf, widthOf, type Layout, type PlacedLine } from "./layout";

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
const STRIDE = PAGE.height + PAGE.gap;

export class View {
  readonly container: HTMLDivElement;
  readonly caret: HTMLDivElement;
  private pages: PageView[] = [];
  private blinkTimer = 0;
  private layoutRef: Layout | null = null;
  private docRef: Doc | null = null;
  private selRef: Selection | null = null;
  private visibleFrom = 0;
  private visibleTo = -1;

  constructor(readonly workspace: HTMLElement) {
    this.container = document.createElement("div");
    this.container.style.cssText = `position:relative;margin:0 auto;width:${PAGE.width + 2 * PAGE.gap}px`;
    workspace.append(this.container);

    this.caret = document.createElement("div");
    const dpr = devicePixelRatio || 1;
    this.caret.style.cssText = `position:absolute;z-index:3;pointer-events:none;background:#000;` +
      `width:${Math.max(1, Math.round(dpr)) / dpr}px`;
    workspace.addEventListener("scroll", () => this.onScroll(), { passive: true });
    addEventListener("resize", () => this.onScroll());
  }

  pageTop(i: number): number {
    return PAGE.gap + i * STRIDE;
  }

  // ----- rendering ----------------------------------------------------------------------

  render(doc: Doc, layout: Layout, sel: Selection): void {
    this.docRef = doc;
    this.layoutRef = layout;
    this.selRef = sel;
    this.ensurePageCount(layout.pageCount);
    this.computeVisible();
    const byPage = this.linesByPage(layout);
    for (let i = 0; i < this.pages.length; i++) {
      const pv = this.pages[i];
      if (i >= this.visibleFrom && i <= this.visibleTo) this.renderLines(pv, byPage[i] ?? [], doc, layout);
      else if (pv.live) this.clearPage(pv);
    }
    this.renderSelection();
  }

  private linesByPage(layout: Layout): PlacedLine[][] {
    const out: PlacedLine[][] = [];
    for (const pl of layout.placed) {
      if (pl.page < this.visibleFrom || pl.page > this.visibleTo) continue;
      (out[pl.page] ??= []).push(pl);
    }
    return out;
  }

  private ensurePageCount(n: number): void {
    while (this.pages.length < n) {
      const el = document.createElement("div");
      el.className = "page";
      el.style.cssText = `position:absolute;margin:0;left:${PAGE.gap}px;top:${this.pageTop(this.pages.length)}px;` +
        `width:${PAGE.width}px;height:${PAGE.height}px`;
      const selLayer = document.createElement("div");
      selLayer.style.cssText = "position:absolute;inset:0;z-index:0";
      const textLayer = document.createElement("div");
      textLayer.style.cssText = "position:absolute;inset:0;z-index:1";
      el.append(selLayer, textLayer);
      this.container.append(el);
      this.pages.push({ el, selLayer, textLayer, lineEls: [], lineKeys: [], live: false });
    }
    while (this.pages.length > n) this.pages.pop()!.el.remove();
    this.container.style.height = `${this.pageTop(n)}px`;
  }

  private computeVisible(): void {
    const top = this.workspace.scrollTop;
    const bottom = top + this.workspace.clientHeight;
    const n = this.pages.length;
    this.visibleFrom = Math.max(0, Math.floor((top - PAGE.gap) / STRIDE) - 1);
    this.visibleTo = Math.min(n - 1, Math.floor((bottom - PAGE.gap) / STRIDE) + 1);
  }

  private onScroll(): void {
    if (!this.layoutRef || !this.docRef || !this.selRef) return;
    const from = this.visibleFrom, to = this.visibleTo;
    this.computeVisible();
    if (from !== this.visibleFrom || to !== this.visibleTo) this.render(this.docRef, this.layoutRef, this.selRef);
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
    for (let k = 0; k < lines.length; k++) {
      const pl = lines[k];
      const para = doc.paras[pl.para];
      const ln = layout.lines[pl.para][pl.line];
      const text = para.text.slice(ln.start, ln.end);
      const s = styleOf(para.style);
      const single = s.sizePt * PT * s.singleEm;
      // Word puts the extra space of "multiple" line spacing above the text.
      const top = PAGE.marginTop + pl.y + (lineHeightPx(s) - single);
      const key = `${para.style}|${top}|${text}`;
      if (pv.lineKeys[k] === key) continue;
      let el = pv.lineEls[k];
      if (!el) {
        el = document.createElement("div");
        el.style.cssText = `position:absolute;white-space:pre;left:${PAGE.marginLeft}px`;
        pv.textLayer.append(el);
        pv.lineEls[k] = el;
      }
      el.style.top = `${top}px`;
      el.style.font = fontCss(s);
      el.style.lineHeight = `${single}px`;
      el.style.color = s.color;
      el.textContent = text;
      pv.lineKeys[k] = key;
    }
    while (pv.lineEls.length > lines.length) {
      pv.lineEls.pop()!.remove();
      pv.lineKeys.pop();
    }
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
      for (let p = from.p; p <= to.p; p++) {
        const para = doc.paras[p];
        const s = styleOf(para.style);
        const lh = lineHeightPx(s);
        const base = layout.firstPlaced[p];
        const lines = layout.lines[p];
        for (let j = 0; j < lines.length; j++) {
          const pl = layout.placed[base + j];
          if (pl.page < this.visibleFrom || pl.page > this.visibleTo) continue;
          const ln = lines[j];
          const a = p === from.p ? Math.max(from.o, ln.start) : ln.start;
          const b = p === to.p ? Math.min(to.o, ln.end) : ln.end;
          const isLastLine = j === lines.length - 1;
          const includesBreak = isLastLine && p < to.p;
          if (b < a || (a === b && !includesBreak)) continue;
          let x1 = widthOf(para.style, para.text, ln.start, a);
          let x2 = x1 + widthOf(para.style, para.text, a, b);
          if (includesBreak) x2 += advance(para.style, "¶");
          const r = document.createElement("div");
          r.style.cssText = `position:absolute;background:${SELECTION_COLOR};` +
            `left:${PAGE.marginLeft + x1}px;top:${PAGE.marginTop + pl.y}px;width:${x2 - x1}px;height:${lh}px`;
          this.pages[pl.page].selLayer.append(r);
        }
      }
    }
    this.renderCaret(empty);
  }

  private renderCaret(visible: boolean): void {
    const layout = this.layoutRef, doc = this.docRef, sel = this.selRef;
    if (!layout || !doc || !sel) return;
    if (!visible) { this.caret.style.display = "none"; return; }
    const c = this.caretBox(sel.head, sel.headUpstream);
    const pv = this.pages[c.page];
    if (pv.el !== this.caret.parentElement) pv.el.append(this.caret);
    const dpr = devicePixelRatio || 1;
    this.caret.style.display = "";
    this.caret.style.left = `${Math.round(c.x * dpr) / dpr}px`;
    this.caret.style.top = `${c.y}px`;
    this.caret.style.height = `${c.h}px`;
    this.restartBlink();
  }

  /**
   * Windows' caret blink (530 ms on, 530 ms off), restarted on every move. A plain timer:
   * restarting a CSS/Web animation on every keystroke delayed the frame by one refresh
   * (measured: 66 ms instead of 33 ms keystroke-to-screen).
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

  // ----- geometry -----------------------------------------------------------------------

  /** The placed line that displays `pos`. Upstream picks the end of the previous line. */
  placedFor(pos: Pos, upstream = false): { pl: PlacedLine; index: number } {
    const layout = this.layoutRef!;
    const lines = layout.lines[pos.p];
    let j = 0;
    while (j < lines.length - 1) {
      const end = lines[j].end;
      if (pos.o < end || (pos.o === end && upstream)) break;
      j++;
    }
    const index = layout.firstPlaced[pos.p] + j;
    return { pl: layout.placed[index], index };
  }

  /** Caret box in page coordinates. */
  caretBox(pos: Pos, upstream = false): { page: number; x: number; y: number; h: number } {
    const doc = this.docRef!;
    const layout = this.layoutRef!;
    const { pl } = this.placedFor(pos, upstream);
    const para = doc.paras[pos.p];
    const ln = layout.lines[pos.p][pl.line];
    const s = styleOf(para.style);
    const single = s.sizePt * PT * s.singleEm;
    const lh = lineHeightPx(s);
    return {
      page: pl.page,
      x: PAGE.marginLeft + widthOf(para.style, para.text, ln.start, pos.o),
      y: PAGE.marginTop + pl.y + (lh - single),
      h: single,
    };
  }

  /** Client (viewport) coordinates of a position: x, line-box top, line-box height. */
  clientPoint(pos: Pos, upstream = false): { x: number; y: number; h: number } {
    const c = this.caretBox(pos, upstream);
    const rect = this.container.getBoundingClientRect();
    return { x: rect.left + PAGE.gap + c.x, y: rect.top + this.pageTop(c.page) + c.y, h: c.h };
  }

  /** Index into layout.placed of the line nearest to a client point, plus the page x. */
  hitLine(clientX: number, clientY: number): { index: number; x: number; inLeftMargin: boolean } {
    const layout = this.layoutRef!;
    const doc = this.docRef!;
    const rect = this.container.getBoundingClientRect();
    const cy = clientY - rect.top;
    const page = Math.max(0, Math.min(layout.pageCount - 1, Math.floor((cy - PAGE.gap / 2) / STRIDE)));
    const y = cy - this.pageTop(page) - PAGE.marginTop;
    const x = clientX - rect.left - PAGE.gap - PAGE.marginLeft;
    // Binary search for the first line on this page, then scan.
    let lo = 0, hi = layout.placed.length - 1;
    while (lo < hi) {
      const mid = (lo + hi) >> 1;
      if (layout.placed[mid].page < page) lo = mid + 1; else hi = mid;
    }
    // A point between two lines (in paragraph spacing) belongs to the line below, as in Word.
    let best = lo;
    for (let i = lo; i < layout.placed.length && layout.placed[i].page === page; i++) {
      const pl = layout.placed[i];
      const lh = lineHeightPx(styleOf(doc.paras[pl.para].style));
      best = i;
      if (y < pl.y + lh) break;
    }
    return { index: best, x, inLeftMargin: x < -4 };
  }

  /**
   * Text position within a placed line: the boundary nearest to x, or with `charUnder` the
   * character under x (for double-click). Returns upstream=true at a soft line end.
   */
  posInLine(index: number, x: number, charUnder = false): { pos: Pos; upstream: boolean } {
    const layout = this.layoutRef!;
    const doc = this.docRef!;
    const pl = layout.placed[index];
    const para = doc.paras[pl.para];
    const ln = layout.lines[pl.para][pl.line];
    const isLast = pl.line === layout.lines[pl.para].length - 1;
    let acc = 0;
    let o = ln.start;
    while (o < ln.end) {
      const w = advance(para.style, para.text[o]);
      if (acc + (charUnder ? w : w / 2) > x) break;
      acc += w;
      o++;
    }
    // At the end of a wrapped line the position equals the next line's start; keep the
    // caret on this line.
    return { pos: { p: pl.para, o }, upstream: !isLast && o === ln.end };
  }

  scrollIntoView(pos: Pos, upstream: boolean): void {
    const c = this.caretBox(pos, upstream);
    const top = this.pageTop(c.page) + c.y;
    const ws = this.workspace;
    const margin = 40;
    if (top < ws.scrollTop + margin) ws.scrollTop = top - margin;
    else if (top + c.h > ws.scrollTop + ws.clientHeight - margin) ws.scrollTop = top + c.h - ws.clientHeight + margin;
  }
}
