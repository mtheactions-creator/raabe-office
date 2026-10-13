// DOM rendering. Every formatting segment of a line is one absolutely positioned element
// with real text, so the system renders glyphs like any native text. Its baseline is placed
// exactly with a per-font measurement (Fonts.baselineOffset). Only pages near the viewport
// carry elements; edits update only lines whose content or position changed.
//
// Pages may differ in size (orientation per section). Zoom scales the page container with a
// CSS transform; all geometry is computed at 100 %, in px from each page's top left.

import { Fonts, PT } from "./fonts";
import { cmp, type Doc, type Pos } from "./doc";
import { COLUMN_BREAK, LINE_BREAK, PAGE_BREAK, type ImageData, type Para } from "./model";
import type { Line, Segment } from "./layout";
import type { Deco, Frame, Layout, PageLayout, PlacedLine, StoryLayout } from "./flow";
import type { ResolvedChar } from "./styles";

export interface Selection {
  anchor: Pos;
  head: Pos;
  /** Draw a caret at a line end rather than at the start of the next line. */
  headUpstream: boolean;
}

/** A highlighted range of a paragraph (spelling error, search hit, commented text …). */
export interface TextMark {
  readonly start: number;
  readonly end: number;
  readonly kind: "spell" | "grammar" | "find" | "findCurrent" | "comment" | "commentActive";
}

export interface RenderOptions {
  /** Layout key of the story with the caret ("main", "fn:2", "hdr:1@0" …). */
  readonly active: string;
  /** Cells selected as a whole (cellKey), drawn filled. */
  readonly cells?: ReadonlySet<string>;
}

interface PageView {
  el: HTMLDivElement;
  back: HTMLDivElement;
  selLayer: HTMLDivElement;
  textLayer: HTMLDivElement;
  front: HTMLDivElement;
  overlay: HTMLDivElement;
  lineEls: HTMLDivElement[];
  lineKeys: string[];
  decoKey: string;
  overlayKey: string;
  live: boolean;
  w: number;
  h: number;
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

/** Colours Word gives the authors of tracked changes, in order of appearance. */
export const AUTHOR_COLORS = ["#b5082e", "#2e74b5", "#548235", "#7030a0", "#c55a11", "#0e7c7b", "#7f6000", "#bf0098"];

const UNDERLINE: Record<string, string> = {
  single: "underline", words: "underline", double: "underline double", thick: "underline", dotted: "underline dotted",
  dottedHeavy: "underline dotted", dash: "underline dashed", dashedHeavy: "underline dashed", dashLong: "underline dashed",
  dotDash: "underline dashed", dotDotDash: "underline dotted", wave: "underline wavy", wavyDouble: "underline wavy",
};

const BORDER_STYLE: Record<string, string> = {
  single: "solid", thick: "solid", double: "double", dotted: "dotted", dashed: "dashed", dashSmallGap: "dashed",
  dotDash: "dashed", dotDotDash: "dotted", triple: "double", wave: "solid", thinThickSmallGap: "double", thickThinSmallGap: "double",
};

const SQUIGGLE = (color: string) => `url("data:image/svg+xml,${encodeURIComponent(
  `<svg xmlns="http://www.w3.org/2000/svg" width="4" height="3"><path d="M0 2.5 L1 0.5 L2 2.5 L3 0.5 L4 2.5" fill="none" stroke="${color}" stroke-width="0.8"/></svg>`)}")`;

let keySeq = 0;
const ids = new WeakMap<object, number>();
function idOf(o: object | undefined): number {
  if (!o) return 0;
  let id = ids.get(o);
  if (id === undefined) { id = ++keySeq; ids.set(o, id); }
  return id;
}

const imageUrls = new WeakMap<ImageData, string>();
/** A blob URL for an embedded picture (kept while the picture data lives). */
export function imageUrl(img: ImageData): string {
  let url = imageUrls.get(img);
  if (!url) {
    url = URL.createObjectURL(new Blob([img.bytes as BlobPart], { type: img.type }));
    imageUrls.set(img, url);
  }
  return url;
}

/** Doc story id of a layout key ("hdr:1@3" → "hdr:1"). */
export const storyIdOf = (key: string) => (key.includes("@") ? key.slice(0, key.lastIndexOf("@")) : key);

export class View {
  readonly sizer: HTMLDivElement;
  readonly container: HTMLDivElement;
  readonly caret: HTMLDivElement;
  private pages: PageView[] = [];
  private tops: number[] = [];
  private maxW = 0;
  private blinkTimer = 0;
  private layoutRef: Layout | null = null;
  private docRef: Doc | null = null;
  private selRef: Selection | null = null;
  private opts: RenderOptions = { active: "main" };
  private visibleFrom = 0;
  private visibleTo = -1;
  private zoomFactor = 1;
  private marks = false;
  /** Highlighted ranges per paragraph (spelling, search hits …), from the host. */
  marksFor: ((story: string, para: Para) => readonly TextMark[] | undefined) | null = null;
  /** Renders an equation (LaTeX) at a font size in px; from the host. */
  renderMath: ((latex: string, sizePx: number) => HTMLElement) | null = null;
  /** Author colours for tracked changes. */
  private authors = new Map<string, string>();

  constructor(readonly workspace: HTMLElement, private fonts: Fonts) {
    this.sizer = document.createElement("div");
    this.sizer.style.cssText = "position:relative;margin:0 auto";
    this.container = document.createElement("div");
    this.container.style.cssText = "position:absolute;left:0;top:0;transform-origin:0 0";
    this.sizer.append(this.container);
    workspace.append(this.sizer);
    this.caret = document.createElement("div");
    this.caret.style.cssText = "position:absolute;z-index:5;pointer-events:none;background:#000";
    workspace.addEventListener("scroll", () => this.onScroll(), { passive: true });
    addEventListener("resize", () => this.onScroll());
  }

  pageTop(i: number): number {
    return this.tops[i] ?? GAP;
  }

  pageLeft(i: number): number {
    const w = this.pages[i]?.w ?? this.maxW;
    return GAP + (this.maxW - w) / 2;
  }

  get zoom(): number {
    return this.zoomFactor;
  }

  get showMarks(): boolean {
    return this.marks;
  }

  get active(): string {
    return this.opts.active;
  }

  setShowMarks(on: boolean): void {
    this.marks = on;
    for (const pv of this.pages) pv.lineKeys = [];
    this.onScroll(true);
  }

  /** Forces a redraw of every line (marks changed). */
  refresh(): void {
    for (const pv of this.pages) pv.lineKeys = [];
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
    return (this.workspace.clientWidth - 2 * GAP - 4) / (this.maxW || 794);
  }

  onePageZoom(): number {
    const h = this.pages[0]?.h ?? 1123;
    return Math.min(this.pageWidthZoom(), (this.workspace.clientHeight - 2 * GAP) / h);
  }

  private applySize(): void {
    const z = this.zoomFactor;
    const w = this.maxW + 2 * GAP;
    const n = this.pages.length;
    const h = n ? this.pageTop(n - 1) + this.pages[n - 1].h + GAP : 2 * GAP;
    this.container.style.transform = z === 1 ? "" : `scale(${z})`;
    this.container.style.width = `${w}px`;
    this.container.style.height = `${h}px`;
    this.sizer.style.width = `${w * z}px`;
    this.sizer.style.height = `${h * z}px`;
  }

  /** The layout of a story by layout key (main text when unknown). */
  storyLayout(key = this.opts.active): StoryLayout {
    const layout = this.layoutRef!;
    return key === "main" ? layout.main : layout.stories.get(key) ?? layout.main;
  }

  /** Paragraphs of the story with layout key `key`. */
  private storyParas(key: string): readonly Para[] {
    const doc = this.docRef!;
    return key === "main" ? doc.main : doc.story(storyIdOf(key)) ?? [];
  }

  // ----- rendering ----------------------------------------------------------------------

  render(doc: Doc, layout: Layout, sel: Selection, opts: RenderOptions = { active: "main" }): void {
    this.docRef = doc;
    this.layoutRef = layout;
    this.selRef = sel;
    this.opts = opts;
    this.ensurePages(layout.pages);
    this.computeVisible();
    const byPage: PlacedLine[][] = [];
    for (const pl of layout.placed) {
      if (pl.page < this.visibleFrom || pl.page > this.visibleTo) continue;
      (byPage[pl.page] ??= []).push(pl);
    }
    for (let i = 0; i < this.pages.length; i++) {
      const pv = this.pages[i];
      if (i >= this.visibleFrom && i <= this.visibleTo) this.renderPage(pv, layout.pages[i], byPage[i] ?? [], doc, layout);
      else if (pv.live) this.clearPage(pv);
    }
    this.renderSelection();
  }

  private ensurePages(pages: readonly PageLayout[]): void {
    let changed = pages.length !== this.pages.length;
    while (this.pages.length < pages.length) {
      const el = document.createElement("div");
      el.className = "page";
      el.style.cssText = "position:absolute;margin:0;overflow:hidden";
      const layer = (z: number) => {
        const d = document.createElement("div");
        d.style.cssText = `position:absolute;inset:0;z-index:${z}`;
        return d;
      };
      const back = layer(0), selLayer = layer(1), textLayer = layer(2), front = layer(3), overlay = layer(4);
      overlay.style.pointerEvents = "none";
      el.append(back, selLayer, textLayer, front, overlay);
      this.container.append(el);
      this.pages.push({ el, back, selLayer, textLayer, front, overlay, lineEls: [], lineKeys: [], decoKey: "", overlayKey: "", live: false, w: 0, h: 0 });
    }
    while (this.pages.length > pages.length) this.pages.pop()!.el.remove();
    let maxW = 0;
    for (const p of pages) maxW = Math.max(maxW, p.setup.width * PT);
    if (maxW !== this.maxW) { this.maxW = maxW; changed = true; }
    let top = GAP;
    pages.forEach((p, i) => {
      const pv = this.pages[i];
      const w = p.setup.width * PT, h = p.setup.height * PT;
      if (pv.w !== w || pv.h !== h || this.tops[i] !== top) changed = true;
      pv.w = w; pv.h = h;
      this.tops[i] = top;
      top += h + GAP;
    });
    this.tops.length = pages.length;
    if (changed || !this.sizer.style.height) {
      this.pages.forEach((pv, i) => {
        pv.el.style.left = `${this.pageLeft(i)}px`;
        pv.el.style.top = `${this.tops[i]}px`;
        pv.el.style.width = `${pv.w}px`;
        pv.el.style.height = `${pv.h}px`;
      });
      this.applySize();
    }
  }

  private pageAtY(y: number): number {
    const n = this.tops.length;
    let lo = 0, hi = n - 1;
    while (lo < hi) {
      const mid = (lo + hi + 1) >> 1;
      if (this.tops[mid] - GAP / 2 <= y) lo = mid; else hi = mid - 1;
    }
    return Math.max(0, lo);
  }

  private computeVisible(): void {
    const z = this.zoomFactor;
    const top = this.workspace.scrollTop / z;
    const bottom = top + this.workspace.clientHeight / z;
    const n = this.pages.length;
    this.visibleFrom = Math.max(0, this.pageAtY(top) - 1);
    this.visibleTo = Math.min(n - 1, this.pageAtY(bottom) + 1);
  }

  private onScroll(force = false): void {
    if (!this.layoutRef || !this.docRef || !this.selRef) return;
    const from = this.visibleFrom, to = this.visibleTo;
    this.computeVisible();
    if (force || from !== this.visibleFrom || to !== this.visibleTo) this.render(this.docRef, this.layoutRef, this.selRef, this.opts);
  }

  private clearPage(pv: PageView): void {
    pv.textLayer.replaceChildren();
    pv.selLayer.replaceChildren();
    pv.back.replaceChildren();
    pv.front.replaceChildren();
    pv.overlay.replaceChildren();
    pv.lineEls = [];
    pv.lineKeys = [];
    pv.decoKey = "";
    pv.overlayKey = "";
    pv.live = false;
  }

  /** Header/footer text is dimmed while the body is edited, and the body while they are. */
  private dimmed(storyKey: string): boolean {
    const a = this.opts.active;
    const isHf = (k: string) => k.startsWith("hdr:") || k.startsWith("ftr:");
    return isHf(a) ? !isHf(storyKey) : isHf(storyKey);
  }

  private renderPage(pv: PageView, page: PageLayout, mainLines: PlacedLine[], doc: Doc, layout: Layout): void {
    pv.live = true;
    // Lines of every story on this page: header, body, footer, notes.
    const items: { key: string; story: StoryLayout; pl: PlacedLine }[] = [];
    const add = (key: string | undefined) => {
      if (!key) return;
      const s = layout.stories.get(key);
      if (s) for (const pl of s.placed) if (!pl.hide) items.push({ key, story: s, pl });
    };
    add(page.headerKey);
    for (const pl of mainLines) if (!pl.hide) items.push({ key: "main", story: layout.main, pl });
    add(page.footerKey);
    for (const id of page.notes) add(id);
    for (let k = 0; k < items.length; k++) {
      const { key, story, pl } = items[k];
      const paras = key === "main" ? doc.main : this.storyParas(key);
      const para = paras[pl.para];
      const pll = story.paras[pl.para];
      if (!para || !pll) continue;
      const line = pll.lines[pl.line];
      const isLast = pl.line === pll.lines.length - 1;
      const marks = this.marksFor?.(key === "main" ? "main" : storyIdOf(key), para);
      const dim = this.dimmed(key);
      const lineKey = `${idOf(line)}|${pl.x}|${pl.y}|${this.marks}|${isLast}|${dim}|${idOf(marks)}|${idOf(para)}`;
      if (pv.lineKeys[k] === lineKey) continue;
      let el = pv.lineEls[k];
      if (!el) {
        el = document.createElement("div");
        el.style.cssText = "position:absolute;left:0;top:0;width:0;height:0";
        pv.textLayer.append(el);
        pv.lineEls[k] = el;
      }
      el.style.opacity = dim ? "0.45" : "";
      el.replaceChildren(...this.lineContent(para, line, pl.x + line.x0, pl.y, isLast, marks));
      pv.lineKeys[k] = lineKey;
    }
    while (pv.lineEls.length > items.length) {
      pv.lineEls.pop()!.remove();
      pv.lineKeys.pop();
    }
    this.renderDeco(pv, page);
    this.renderOverlay(pv, page, layout);
  }

  private renderDeco(pv: PageView, page: PageLayout): void {
    const key = page.deco.map((d) => {
      if (d.kind === "rect") return `r${d.x},${d.y},${d.w},${d.h},${d.fill},${this.dimmed(d.story)}`;
      if (d.kind === "line") return `l${d.x1},${d.y1},${d.x2},${d.y2},${d.width},${d.color},${d.style},${this.dimmed(d.story)}`;
      return `i${d.x},${d.y},${d.w},${d.h},${idOf(d.image)},${d.behind},${this.dimmed(d.story)}`;
    }).join(";");
    if (key === pv.decoKey) return;
    pv.decoKey = key;
    const back: HTMLElement[] = [], front: HTMLElement[] = [];
    for (const d of page.deco) {
      const el = this.decoElement(d);
      if (this.dimmed(d.story)) el.style.opacity = "0.45";
      (d.kind === "image" && !d.behind ? front : back).push(el);
    }
    pv.back.replaceChildren(...back);
    pv.front.replaceChildren(...front);
  }

  private decoElement(d: Deco): HTMLElement {
    const el = document.createElement("div");
    if (d.kind === "rect") {
      el.style.cssText = `position:absolute;left:${d.x}px;top:${d.y}px;width:${d.w}px;height:${d.h}px;background:${d.fill}`;
    } else if (d.kind === "line") {
      const w = Math.max(d.width, 0.75);
      const style = BORDER_STYLE[d.style] ?? "solid";
      const width = style === "double" ? Math.max(3, w * 3) : w;
      if (Math.abs(d.y1 - d.y2) < 0.01) {
        el.style.cssText = `position:absolute;left:${Math.min(d.x1, d.x2)}px;top:${d.y1 - width / 2}px;width:${Math.abs(d.x2 - d.x1)}px;height:0;border-top:${width}px ${style} ${d.color}`;
      } else {
        el.style.cssText = `position:absolute;left:${d.x1 - width / 2}px;top:${Math.min(d.y1, d.y2)}px;width:0;height:${Math.abs(d.y2 - d.y1)}px;border-left:${width}px ${style} ${d.color}`;
      }
    } else {
      el.style.cssText = `position:absolute;left:${d.x}px;top:${d.y}px;width:${d.w}px;height:${d.h}px`;
      if (d.image) {
        const img = document.createElement("img");
        img.src = imageUrl(d.image);
        img.draggable = false;
        img.style.cssText = "width:100%;height:100%;display:block";
        el.append(img);
      } else {
        el.style.cssText += ";box-sizing:border-box;border:1px dashed #9e9e9e;background:#f3f3f3";
      }
    }
    el.style.pointerEvents = "none";
    return el;
  }

  /** While a header or footer is edited: the dashed boundaries and their labels. */
  private renderOverlay(pv: PageView, page: PageLayout, layout: Layout): void {
    const a = this.opts.active;
    const editing = a.startsWith("hdr:") || a.startsWith("ftr:");
    const sections = new Set(layout.pages.map((p) => p.section)).size;
    const key = editing ? `${page.bodyTop}|${page.bodyBottom}|${page.section}|${sections}|${pv.w}` : "";
    if (key === pv.overlayKey) return;
    pv.overlayKey = key;
    if (!editing) { pv.overlay.replaceChildren(); return; }
    const sect = sections > 1 ? ` -Abschnitt ${page.section + 1}-` : "";
    const first = page.section !== layout.pages[page.index - 1]?.section;
    const sectDef = page.index >= 0 ? layout.pages[page.index] : page;
    void sectDef;
    const label = (text: string, y: number, above: boolean) => {
      const line = document.createElement("div");
      line.style.cssText = `position:absolute;left:0;width:${pv.w}px;top:${y}px;border-top:1px dashed #2b579a`;
      const tag = document.createElement("div");
      tag.style.cssText = `position:absolute;left:${page.setup.marginLeft * PT - 8}px;${above ? `top:${y - 18}px` : `top:${y + 1}px`};` +
        "font:11px 'Segoe UI',sans-serif;color:#fff;background:#2b579a;padding:1px 6px;white-space:pre";
      tag.textContent = text;
      return [line, tag];
    };
    const hdrName = first && page.header === undefined ? "Kopfzeile" : "Kopfzeile";
    pv.overlay.replaceChildren(
      ...label(`${hdrName}${sect}`, page.bodyTop - 2, false),
      ...label(`Fußzeile${sect}`, page.bodyBottom + 2, true),
    );
  }

  private authorColor(author: string): string {
    let c = this.authors.get(author);
    if (!c) { c = AUTHOR_COLORS[this.authors.size % AUTHOR_COLORS.length]; this.authors.set(author, c); }
    return c;
  }

  private spanStyle(c: ResolvedChar): string {
    const size = Fonts.renderSize(c) * PT;
    const lh = this.fonts.lineHeight(c.font, Fonts.renderSize(c)) * PT;
    const rev = c.rev ? this.authorColor(c.rev.author) : null;
    let s = `position:absolute;white-space:pre;font:${this.fonts.cssFont(c)};line-height:${lh}px;` +
      `color:${rev ?? (c.color === "auto" ? "#000" : `#${c.color}`)};font-kerning:${c.kern && c.size >= c.kern ? "normal" : "none"};font-variant-ligatures:none`;
    if (c.spacing) s += `;letter-spacing:${c.spacing * PT}px`;
    if (c.caps) s += ";text-transform:uppercase";
    if (c.smallCaps) s += ";font-variant-caps:small-caps";
    const deco: string[] = [];
    if (c.underline !== "none" && UNDERLINE[c.underline]) deco.push(UNDERLINE[c.underline]);
    else if (c.rev?.kind === "ins") deco.push("underline");
    if (c.strike || c.dstrike || c.rev?.kind === "del") deco.push("line-through");
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

  private lineContent(para: Para, line: Line, left: number, top: number, isLast: boolean, marks?: readonly TextMark[]): HTMLElement[] {
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
    if (marks?.length) out.push(...this.markElements(line, marks, left, top, baseline));
    if (line.label) {
      const c = line.label.c;
      const lh = this.fonts.lineHeight(c.font, Fonts.renderSize(c)) * PT;
      const el = document.createElement("span");
      el.style.cssText = this.spanStyle(c) + `;left:${left - line.x0 + line.label.x}px;top:${baseline - this.fonts.baselineOffset(c, lh)}px;text-decoration:none`;
      el.textContent = line.label.text;
      out.push(el);
    }
    for (const seg of line.segments) out.push(...this.segment(para, line, seg, left, baseline, top));
    if (this.marks && isLast && line.end === para.text.length) {
      const c = line.segments.length ? line.segments[line.segments.length - 1].c : null;
      out.push(this.markGlyph(para.cell && false ? "¤" : "¶", left + line.charX[line.end - line.start], baseline, c));
    }
    return out;
  }

  private markElements(line: Line, marks: readonly TextMark[], left: number, top: number, baseline: number): HTMLElement[] {
    const out: HTMLElement[] = [];
    for (const m of marks) {
      const a = Math.max(m.start, line.start), b = Math.min(m.end, line.end);
      if (b <= a) continue;
      const x1 = left + line.charX[a - line.start], x2 = left + line.charX[b - line.start];
      const el = document.createElement("div");
      if (m.kind === "spell" || m.kind === "grammar") {
        el.style.cssText = `position:absolute;left:${x1}px;top:${baseline + 1.5}px;width:${x2 - x1}px;height:3px;` +
          `background:${SQUIGGLE(m.kind === "spell" ? "#e81123" : "#2b579a")} repeat-x;pointer-events:none;z-index:1`;
      } else {
        const color = m.kind === "find" ? "#ffff00" : m.kind === "findCurrent" ? "#ffc000" : m.kind === "comment" ? "#f8e3c4" : "#f5c58c";
        el.style.cssText = `position:absolute;left:${x1}px;top:${top}px;width:${x2 - x1}px;height:${line.height}px;background:${color};pointer-events:none;z-index:-1`;
      }
      out.push(el);
    }
    return out;
  }

  private markGlyph(glyph: string, x: number, baseline: number, c: ResolvedChar | null): HTMLElement {
    const el = document.createElement("span");
    const font = c ? this.fonts.cssFont({ ...c, bold: false, italic: false }) : `14.667px Calibri, sans-serif`;
    const lh = c ? this.fonts.lineHeight(c.font, Fonts.renderSize(c)) * PT : 19;
    const b = c ? this.fonts.baselineOffset(c, lh) : 15;
    el.style.cssText = `position:absolute;white-space:pre;left:${x}px;top:${baseline - b}px;font:${font};line-height:${lh}px;color:#000;pointer-events:none`;
    el.textContent = glyph;
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
      if (seg.leader && seg.width > 0) out.push(this.leader(seg, c, x, y, lh));
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
        if (ch === PAGE_BREAK || ch === COLUMN_BREAK) {
          const pb = document.createElement("div");
          pb.style.cssText = `position:absolute;left:${x}px;top:${top + line.height / 2 - 8}px;font:11px 'Segoe UI',sans-serif;color:#5b5b5b;white-space:pre;border-top:1px dotted #5b5b5b;padding-top:1px`;
          pb.textContent = ch === PAGE_BREAK ? "··········Seitenumbruch··········" : "··········Spaltenumbruch··········";
          out.push(pb);
        }
      }
      return out;
    }
    if (seg.kind === "object") {
      const o = seg.obj;
      if (o?.kind === "image" && o.width !== undefined && o.height !== undefined) {
        const w = o.width * PT, h = o.height * PT;
        if (o.image) {
          const img = document.createElement("img");
          img.src = imageUrl(o.image);
          img.draggable = false;
          img.style.cssText = `position:absolute;left:${x}px;top:${baseline - h}px;width:${w}px;height:${h}px`;
          out.push(img);
        } else {
          const box = document.createElement("div");
          box.style.cssText = `position:absolute;left:${x}px;top:${baseline - h}px;width:${w}px;height:${h}px;` +
            "box-sizing:border-box;border:1px dashed #9e9e9e;background:#f3f3f3;color:#757575;font:11px 'Segoe UI',sans-serif;display:flex;align-items:center;justify-content:center";
          box.textContent = "Bild";
          out.push(box);
        }
        return out;
      }
      if (o?.kind === "math" && o.latex !== undefined && this.renderMath) {
        const el = this.renderMath(o.latex, c.size * PT);
        const h = (o.height ?? c.size) * PT, d = (o.depth ?? 0) * PT;
        el.style.position = "absolute";
        el.style.left = `${x}px`;
        el.style.top = `${baseline - (h - d)}px`;
        el.style.width = `${seg.width}px`;
        el.style.height = `${h}px`;
        el.style.color = c.color === "auto" ? "#000" : `#${c.color}`;
        out.push(el);
        return out;
      }
      const el = document.createElement("span");
      el.style.cssText = this.spanStyle(c) + `;left:${x}px;top:${y}px` + (o?.href ? ";color:#0563c1;text-decoration:underline" : "");
      el.textContent = seg.display ?? o?.display ?? "";
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

  /** Tab leader: dots (or dashes, a line) across the tab, ending at the stop. */
  private leader(seg: Segment, c: ResolvedChar, x: number, y: number, lh: number): HTMLElement {
    const ch = seg.leader === "hyphen" ? "-" : seg.leader === "underscore" || seg.leader === "heavy" ? "_" : seg.leader === "middleDot" ? "·" : ".";
    const w = this.fonts.advance(c, ch) || 3;
    const n = Math.max(0, Math.floor((seg.width - w / 2) / w));
    const el = document.createElement("span");
    el.style.cssText = this.spanStyle({ ...c, underline: "none", strike: false }) + `;left:${x + seg.width - n * w}px;top:${y}px;line-height:${lh}px`;
    el.textContent = ch.repeat(n);
    return el;
  }

  private glyphAt(glyph: string, c: ResolvedChar, x: number, y: number, lh: number): HTMLElement {
    const el = document.createElement("span");
    el.style.cssText = `position:absolute;white-space:pre;left:${x}px;top:${y}px;font:${this.fonts.cssFont({ ...c, bold: false, italic: false })};line-height:${lh}px;color:#000;pointer-events:none`;
    el.textContent = glyph;
    return el;
  }

  /** Renders every page completely (no virtualisation, no selection) for printing. */
  renderForPrint(target: HTMLElement, pages?: readonly number[]): void {
    const layout = this.layoutRef, doc = this.docRef;
    if (!layout || !doc) return;
    const marks = this.marks, opts = this.opts, marksFor = this.marksFor;
    this.marks = false;
    this.opts = { active: "main" };
    this.marksFor = null;
    target.replaceChildren();
    // Named pages carry each page's size to the printer (sections may change orientation).
    const sizes = new Map<string, string>();
    const style = document.createElement("style");
    target.append(style);
    for (const page of layout.pages) {
      if (pages && !pages.includes(page.index)) continue;
      const w = page.setup.width, h = page.setup.height;
      const name = `p${Math.round(w * 10)}x${Math.round(h * 10)}`;
      sizes.set(name, `@page ${name} { size: ${w}pt ${h}pt; margin: 0 }`);
      const el = document.createElement("div");
      el.className = "print-page";
      el.style.cssText = `position:relative;width:${w * PT}px;height:${h * PT}px;overflow:hidden;break-after:page;page:${name}`;
      const pv: PageView = {
        el, back: document.createElement("div"), selLayer: document.createElement("div"), textLayer: document.createElement("div"),
        front: document.createElement("div"), overlay: document.createElement("div"), lineEls: [], lineKeys: [], decoKey: "", overlayKey: "x",
        live: false, w: w * PT, h: h * PT,
      };
      for (const l of [pv.back, pv.textLayer, pv.front]) { l.style.cssText = "position:absolute;inset:0"; el.append(l); }
      this.renderPage(pv, page, layout.placed.filter((pl) => pl.page === page.index), doc, layout);
      target.append(el);
    }
    style.textContent = [...sizes.values()].join("\n");
    this.marks = marks;
    this.opts = opts;
    this.marksFor = marksFor;
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
    const story = this.storyLayout();
    const cells = this.opts.cells;
    if (cells?.size) {
      for (const fr of layout.frames) {
        if (fr.kind !== "cell" || !fr.cell || !cells.has(fr.cell) || fr.story !== this.opts.active) continue;
        if (fr.page < this.visibleFrom || fr.page > this.visibleTo) continue;
        const r = document.createElement("div");
        r.style.cssText = `position:absolute;background:${SELECTION_COLOR};left:${fr.x}px;top:${fr.y}px;width:${fr.w}px;height:${fr.h}px`;
        this.pages[fr.page].selLayer.append(r);
      }
    } else if (!empty) {
      for (let p = from.p; p <= to.p; p++) {
        const plo = story.paras[p];
        const idx = story.index[p];
        if (!plo || !idx) continue;
        for (let j = 0; j < plo.lines.length; j++) {
          const pl = story.placed[idx[j]];
          if (!pl || pl.hide || pl.page < this.visibleFrom || pl.page > this.visibleTo) continue;
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
          r.style.cssText = `position:absolute;background:${SELECTION_COLOR};left:${pl.x + ln.x0 + x1}px;top:${pl.y}px;width:${Math.max(0, x2 - x1)}px;height:${ln.height}px`;
          this.pages[pl.page].selLayer.append(r);
        }
      }
    }
    this.renderCaret(empty && !cells?.size);
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

  /** The placed line holding a position of the active story. */
  placedFor(pos: Pos, upstream = false): { pl: PlacedLine; index: number } {
    const story = this.storyLayout();
    const pll = story.paras[pos.p];
    const lines = pll?.lines ?? [];
    let j = 0;
    while (j < lines.length - 1) {
      const end = lines[j].end;
      if (pos.o < end || (pos.o === end && upstream)) break;
      j++;
    }
    const index = story.index[pos.p]?.[j] ?? story.firstPlaced[pos.p] ?? 0;
    return { pl: story.placed[index] ?? story.placed[0], index };
  }

  lineOf(pl: PlacedLine, key = this.opts.active): Line {
    return this.storyLayout(key).paras[pl.para].lines[pl.line];
  }

  frameOf(pl: PlacedLine): Frame | undefined {
    return this.layoutRef!.frames[pl.frame];
  }

  /** Caret box in (unzoomed) page coordinates. */
  caretBox(pos: Pos, upstream = false): { page: number; x: number; y: number; h: number } {
    const doc = this.docRef!;
    const { pl } = this.placedFor(pos, upstream);
    if (!pl) return { page: 0, x: 0, y: 0, h: 16 };
    const line = this.lineOf(pl);
    const para = doc.paras[pos.p];
    const o = Math.max(line.start, Math.min(line.end, pos.o));
    const rp = this.storyLayout().paras[pos.p].rp;
    const c = doc.styles.resolveChar(rp.style, o > 0 && para.text.length ? propsBefore(para, o) : para.mark);
    const h = this.fonts.lineHeight(c.font, c.size) * PT;
    const a = this.fonts.ascent(c.font, c.size) * PT;
    return { page: pl.page, x: pl.x + line.x0 + line.charX[o - line.start], y: pl.y + line.baseline - a, h };
  }

  clientPoint(pos: Pos, upstream = false): { x: number; y: number; h: number } {
    const c = this.caretBox(pos, upstream);
    const rect = this.sizer.getBoundingClientRect();
    const z = this.zoomFactor;
    return { x: rect.left + (this.pageLeft(c.page) + c.x) * z, y: rect.top + (this.pageTop(c.page) + c.y) * z, h: c.h * z };
  }

  /** Page under a client point and the point in that page's coordinates. */
  pointAt(clientX: number, clientY: number): { page: number; x: number; y: number } {
    const rect = this.sizer.getBoundingClientRect();
    const z = this.zoomFactor;
    const cy = (clientY - rect.top) / z;
    const page = Math.min(this.pages.length - 1, this.pageAtY(cy));
    return { page, x: (clientX - rect.left) / z - this.pageLeft(page), y: cy - this.pageTop(page) };
  }

  /** What part of a page a client point is in. */
  regionAt(clientX: number, clientY: number): { page: number; kind: "header" | "footer" | "body" | "note"; key?: string } {
    const layout = this.layoutRef!;
    const { page, x, y } = this.pointAt(clientX, clientY);
    const pg = layout.pages[page];
    for (const id of pg.notes) {
      const fr = layout.frames.find((f) => f.page === page && f.story === id);
      if (fr && y >= fr.y && y <= fr.y + fr.h && x >= fr.x - 20 && x <= fr.x + fr.w) return { page, kind: "note", key: id };
    }
    if (y < pg.bodyTop) return { page, kind: "header", key: pg.headerKey };
    if (y > pg.bodyBottom) return { page, kind: "footer", key: pg.footerKey };
    return { page, kind: "body" };
  }

  /**
   * The line of the active story nearest to a client point: inside a table cell the cell's
   * lines, otherwise the lines of the column under the point; between two lines the one
   * below (as in Word).
   */
  hitLine(clientX: number, clientY: number): { index: number; x: number; inLeftMargin: boolean } {
    const layout = this.layoutRef!;
    const story = this.storyLayout();
    const pt = this.pointAt(clientX, clientY);
    // The nearest page that has lines of this story.
    const onPage = (p: number) => story.placed.filter((pl) => pl.page === p && !pl.ghost && !pl.hide);
    let cand = onPage(pt.page);
    let page = pt.page;
    for (let d = 1; !cand.length && d < layout.pageCount + 1; d++) {
      for (const p of [pt.page - d, pt.page + d]) {
        if (p < 0 || p >= layout.pageCount) continue;
        const c = onPage(p);
        if (c.length) { cand = c; page = p; break; }
      }
    }
    if (!cand.length) return { index: 0, x: pt.x, inLeftMargin: false };
    let { x, y } = pt;
    if (page !== pt.page) y = page < pt.page ? Infinity : -Infinity;
    const lineH = (pl: PlacedLine) => story.paras[pl.para].lines[pl.line].height;
    const frames = new Map<number, Frame>();
    for (const pl of cand) frames.set(pl.frame, layout.frames[pl.frame]);
    // A cell (or other frame) that contains the point wins.
    let inside = [...frames.values()].filter((f) => f && f.kind === "cell" && x >= f.x && x <= f.x + f.w && y >= f.y && y <= f.y + f.h);
    if (!inside.length) {
      inside = [...frames.values()].filter((f) => f && x >= f.x && x <= f.x + f.w);
      if (!inside.length) {
        // Left or right of every frame: the horizontally nearest ones.
        let best = Infinity;
        for (const f of frames.values()) if (f) best = Math.min(best, Math.abs(x < f.x ? f.x - x : x - (f.x + f.w)));
        inside = [...frames.values()].filter((f) => f && Math.abs(x < f.x ? f.x - x : x - (f.x + f.w)) <= best + 0.5);
      }
    }
    const ok = new Set(inside.map((f) => f.id));
    const lines = cand.filter((pl) => ok.has(pl.frame)).sort((a, b) => a.y - b.y);
    let hit = lines[lines.length - 1];
    for (const pl of lines) if (y < pl.y + lineH(pl)) { hit = pl; break; }
    const index = story.placed.indexOf(hit);
    const fr = layout.frames[hit.frame];
    const s = layout.pages[hit.page].setup;
    const inLeftMargin = (!fr || fr.kind === "body" || fr.kind === "header" || fr.kind === "footer" || fr.kind === "note") && x < s.marginLeft * PT - 4 && page === pt.page;
    return { index, x, inLeftMargin };
  }

  /** Position in a placed line of the active story at x (page px). */
  posInLine(index: number, x: number, charUnder = false): { pos: Pos; upstream: boolean } {
    const story = this.storyLayout();
    const pl = story.placed[index];
    const line = this.lineOf(pl);
    const isLast = pl.line === story.paras[pl.para].lines.length - 1;
    const rx = x - pl.x - line.x0;
    let o = line.start;
    // Positions inside a break character are not reachable; stop before it.
    const doc = this.docRef!;
    const text = doc.paras[pl.para].text;
    let limit = line.end;
    if (limit > line.start && (text[limit - 1] === LINE_BREAK || text[limit - 1] === PAGE_BREAK || text[limit - 1] === COLUMN_BREAK)) limit--;
    while (o < limit) {
      const a = line.charX[o - line.start], b = line.charX[o + 1 - line.start];
      if ((charUnder ? b : (a + b) / 2) > rx) break;
      o++;
    }
    if (doc.paras[pl.para].block) return { pos: { p: pl.para, o: 0 }, upstream: false };
    return { pos: { p: pl.para, o }, upstream: !isLast && o === line.end };
  }

  /** Image decoration under a client point (floating pictures), for selection. */
  imageAt(clientX: number, clientY: number): Extract<Deco, { kind: "image" }> | null {
    const { page, x, y } = this.pointAt(clientX, clientY);
    const pg = this.layoutRef?.pages[page];
    if (!pg) return null;
    for (let i = pg.deco.length - 1; i >= 0; i--) {
      const d = pg.deco[i];
      if (d.kind === "image" && x >= d.x && x <= d.x + d.w && y >= d.y && y <= d.y + d.h && d.story === "main") return d;
    }
    return null;
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

  /** Scrolls a page to the top of the window (navigation pane, Gehe zu). */
  scrollToPage(page: number): void {
    this.workspace.scrollTop = (this.pageTop(page) - GAP / 2) * this.zoomFactor;
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
