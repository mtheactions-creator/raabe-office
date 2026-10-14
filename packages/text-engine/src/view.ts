// DOM rendering. Every formatting segment of a line is one absolutely positioned element
// with real text, so the system renders glyphs like any native text. Its baseline is placed
// exactly with a per-font measurement (Fonts.baselineOffset). Only pages near the viewport
// carry elements; edits update only lines whose content or position changed.
//
// Pages may differ in size (orientation per section). Zoom scales the page container with a
// CSS transform; all geometry is computed at 100 %, in px from each page's top left.

import { pictureElement } from "./picture";
import { shapeElement, shapeGeometry, shapeTextRect } from "./shape";

/** Where each drawn equation belongs (its paragraph and offset), for the formula editor. */
export const mathElements = new WeakMap<Element, { para: Para; o: number }>();
import { flattenPath } from "./geometry";
import { Fonts, PT } from "./fonts";
import { cmp, type Doc, type DocSettings, type Pos } from "./doc";
import { COLUMN_BREAK, LINE_BREAK, OBJ, PAGE_BREAK, type ImageData, type InlineObject, type Para } from "./model";
import type { Line, Segment } from "./layout";
import type { Deco, Frame, Layout, PageLayout, PlacedLine, StoryLayout } from "./flow";
import { themeColor, type ResolvedChar } from "./styles";
import { textFxCss } from "./textfx";

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
  /** Further selected objects (Umschalt+Klick): their frames (page px). */
  readonly others?: readonly { readonly page: number; readonly x: number; readonly y: number; readonly w: number; readonly h: number }[];
  /** A selected picture: its box gets Word's eight sizing handles. */
  readonly image?: {
    readonly page: number; readonly x: number; readonly y: number; readonly w: number; readonly h: number;
    /** Zuschneiden: the whole picture (page px) and its address, shown dimmed around the frame. */
    readonly crop?: { readonly x: number; readonly y: number; readonly w: number; readonly h: number; readonly url: string };
    /** Pictures get the rotation handle and Word's layout options button. */
    readonly picture?: boolean;
    /** A shape whose text is being edited: its frame is dashed. */
    readonly editing?: boolean;
    /** Shapes: the yellow adjust handles (page px). */
    readonly adjust?: readonly { readonly x: number; readonly y: number }[];
    /** Lines and connectors: their two ends (page px), shown instead of the frame. */
    readonly ends?: readonly [{ readonly x: number; readonly y: number }, { readonly x: number; readonly y: number }];
  };
  /** The table (id) that shows Word's move and resize handles (caret or mouse in it). */
  readonly table?: number;
}

/** Word's rotation handle (a circular arrow) and the layout options button's icon. */
const ROTATE_SVG = '<svg viewBox="0 0 18 18" width="18" height="18" aria-hidden="true"><path d="M13.5 6.2A5.5 5.5 0 1 0 14.5 10" fill="none" stroke="#5f5f5f" stroke-width="1.6" stroke-linecap="round"/><path d="M14.6 2.8v4h-4" fill="none" stroke="#5f5f5f" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round"/></svg>';
const LAYOUT_SVG = '<svg viewBox="0 0 16 16" width="16" height="16" aria-hidden="true"><path d="M1 3h14M1 13h14M1 8h3M12 8h3" stroke="#2b579a" stroke-width="1.2"/><path d="M5 10.5a3 3 0 0 1 6 0z" fill="#2b579a"/></svg>';

/** Word's table move handle: a small box with a four-way arrow. */
const MOVE_HANDLE_SVG = '<svg viewBox="0 0 11 11" width="11" height="11" aria-hidden="true"><path d="M5.5 1.5v8M1.5 5.5h8" stroke="#000" stroke-width="1"/>' +
  '<path d="M5.5 0.8 4 2.6h3zM5.5 10.2 4 8.4h3zM0.8 5.5 2.6 4v3zM10.2 5.5 8.4 4v3z" fill="#000"/></svg>';

export type Handle = "nw" | "n" | "ne" | "e" | "se" | "s" | "sw" | "w" | "rot" | "layout" | `crop-${"nw" | "n" | "ne" | "e" | "se" | "s" | "sw" | "w"}` | `adj-${number}` | "end-0" | "end-1";
const HANDLES: Handle[] = ["nw", "n", "ne", "e", "se", "s", "sw", "w"];

interface PageView {
  el: HTMLDivElement;
  back: HTMLDivElement;
  selLayer: HTMLDivElement;
  /** Selection in a shape's text: above the shapes (the front layer). */
  selTop: HTMLDivElement;
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

export type MarkupMode = "all" | "simple" | "none" | "original";
const MAIN_KEY = "main";

/** Whether a line shows tracked changes (its text, or the paragraph mark on the last line). */
function lineHasRevision(para: Para, line: Line, isLast: boolean): boolean {
  if (isLast && para.mark.rev) return true;
  let pos = 0;
  for (const r of para.runs) {
    const a = pos;
    pos += r.len;
    if (pos <= line.start) continue;
    if (a >= line.end) break;
    if (r.props.rev) return true;
  }
  return false;
}

/** Doc story id of a layout key ("hdr:1@3" → "hdr:1"). */
export const storyIdOf = (key: string) => (key.includes("@") ? key.slice(0, key.lastIndexOf("@")) : key);

export class View {
  readonly sizer: HTMLDivElement;
  readonly container: HTMLDivElement;
  readonly caret: HTMLDivElement;
  private pages: PageView[] = [];
  private tops: number[] = [];
  private lefts: number[] = [];
  private maxW = 0;
  /** Width and height of all pages as placed (px at 100 %). */
  private areaW = 0;
  private areaH = 0;
  /**
   * How pages stand: "pages" (Drucklayout: one below the other, side by side when several fit,
   * as in Word), "flat" (Weblayout: one page without gap or shadow), "row" (Lesemodus: screens
   * side by side, turned horizontally).
   */
  arrangement: "pages" | "flat" | "row" = "pages";
  /** Pages per row as last placed. */
  private perRow = 1;
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
  /** Measures an equation (points), the same function the layout uses; from the host. */
  measureMath: ((o: InlineObject, size: number, display: boolean) => { width: number; height: number; depth: number }) | null = null;
  /** Renders an equation (LaTeX) at a font size in px; from the host. */
  renderMath: ((o: InlineObject, sizePx: number, display: boolean) => HTMLElement) | null = null;
  /** Author colours for tracked changes. */
  private authors = new Map<string, string>();
  /** Markup für Überarbeitung anzeigen; decides the change bars in the margin. */
  markupMode: MarkupMode = "all";
  /** Room right of the pages for comment cards (px at 100 %); 0 without comments. */
  private markupW = 0;
  /** Where the host puts comment cards: right of the pages, scrolled and zoomed with them. */
  readonly markupLayer: HTMLDivElement;
  /** Gitternetzlinien anzeigen: dashed cell edges on screen (on by default, as in Word). */
  showGridlines = true;
  private printing = false;

  constructor(readonly workspace: HTMLElement, private fonts: Fonts) {
    this.sizer = document.createElement("div");
    this.sizer.style.cssText = "position:relative;margin:0 auto";
    this.container = document.createElement("div");
    this.container.style.cssText = "position:absolute;left:0;top:0;transform-origin:0 0";
    this.sizer.append(this.container);
    this.markupLayer = document.createElement("div");
    this.markupLayer.style.cssText = "position:absolute;top:0;display:none";
    this.container.append(this.markupLayer);
    workspace.append(this.sizer);
    this.caret = document.createElement("div");
    this.caret.style.cssText = "position:absolute;z-index:5;pointer-events:none;background:#000";
    workspace.addEventListener("scroll", () => this.onScroll(), { passive: true });
    addEventListener("blur", () => { this.windowActive = false; this.caret.style.display = "none"; this.stopBlink(); });
    addEventListener("focus", () => { this.windowActive = true; if (this.caretWanted) this.renderCaret(true); });
    addEventListener("resize", () => this.onScroll());
    if (typeof ResizeObserver !== "undefined") new ResizeObserver(() => this.resized()).observe(workspace);
  }

  pageTop(i: number): number {
    return this.tops[i] ?? this.gap;
  }

  pageLeft(i: number): number {
    return this.lefts[i] ?? this.gap + (this.maxW - (this.pages[i]?.w ?? this.maxW)) / 2;
  }

  private get gap(): number {
    return this.arrangement === "flat" ? 0 : GAP;
  }

  /** Ansicht > Gitternetzlinien: Word's drawing grid (9 pt from the margins) over the text area. */
  get drawingGrid(): boolean {
    return this.workspace.classList.contains("drawgrid");
  }

  setDrawingGrid(on: boolean): void {
    this.workspace.classList.toggle("drawgrid", on);
  }

  /** Changes how pages stand (Drucklayout, Weblayout, Lesemodus). */
  setArrangement(a: View["arrangement"]): void {
    if (a === this.arrangement) return;
    this.arrangement = a;
    this.workspace.classList.toggle("flatpages", a === "flat");
    this.workspace.classList.toggle("rowpages", a === "row");
    this.placePages(true);
  }

  /** Mehrere Seiten: the zoom at which `n` pages stand side by side in the window. */
  multiplePagesZoom(n = 2): number {
    return (this.workspace.clientWidth - (n + 1) * GAP - 8) / (n * (this.maxW || 794));
  }

  /** Lesemodus: the page shown at the left of the window, and the number of pages. */
  screenInfo(): { page: number; count: number } {
    const x = this.workspace.scrollLeft / this.zoomFactor + this.gap + 4;
    let page = 0;
    for (let i = 0; i < this.lefts.length; i++) if (this.lefts[i] <= x) page = i;
    return { page, count: this.pages.length };
  }

  /** Lesemodus: turns to a page (the window shows it at its left). */
  showScreen(page: number): void {
    const i = Math.max(0, Math.min(this.pages.length - 1, page));
    this.workspace.scrollLeft = (this.pageLeft(i) - this.gap) * this.zoomFactor;
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
    // The page at the top of the window stays there (pages may move into rows).
    const y = ws.scrollTop / this.zoomFactor;
    const page = this.pageAtY(y);
    const within = y - this.pageTop(page);
    this.zoomFactor = z;
    // Pages side by side depend on the zoom.
    if (!this.placePages(false)) this.applySize();
    ws.scrollTop = Math.max(0, (this.pageTop(page) + within) * z);
    this.onScroll(true);
  }

  /**
   * Places the pages (tops and lefts) for the arrangement, zoom and window width. Returns
   * whether anything moved.
   */
  private placePages(force: boolean): boolean {
    const gap = this.gap;
    const n = this.pages.length;
    const z = this.zoomFactor;
    // Several pages per row once they fit (not while comment cards need the room beside them).
    const fit = Math.floor((this.workspace.clientWidth / z - gap) / ((this.maxW || 794) + gap));
    const perRow = this.arrangement === "row" ? Math.max(1, n) : this.arrangement === "flat" || this.markupW ? 1 : Math.max(1, Math.min(n, fit));
    const tops: number[] = [], lefts: number[] = [];
    let top = gap, areaW = 0;
    for (let r = 0; r * perRow < n; r++) {
      const from = r * perRow, to = Math.min(n, from + perRow);
      let rowH = 0, x = gap;
      for (let i = from; i < to; i++) {
        const pv = this.pages[i];
        // One column: pages centred on the widest; in rows each takes its own width.
        lefts[i] = perRow === 1 ? gap + (this.maxW - pv.w) / 2 : x;
        tops[i] = top;
        x += (perRow === 1 ? this.maxW : pv.w) + gap;
        rowH = Math.max(rowH, pv.h);
      }
      areaW = Math.max(areaW, x);
      top += rowH + gap;
    }
    const changed = force || perRow !== this.perRow || tops.some((t, i) => t !== this.tops[i]) || lefts.some((l, i) => l !== this.lefts[i]) || n !== this.tops.length;
    this.perRow = perRow;
    this.tops = tops;
    this.lefts = lefts;
    this.areaW = n ? areaW : this.maxW + 2 * gap;
    this.areaH = n ? top : 2 * gap;
    if (!changed) return false;
    this.pages.forEach((pv, i) => {
      pv.el.style.left = `${lefts[i]}px`;
      pv.el.style.top = `${tops[i]}px`;
      pv.el.style.width = `${pv.w}px`;
      pv.el.style.height = `${pv.h}px`;
    });
    this.applySize();
    return true;
  }

  pageWidthZoom(): number {
    return (this.workspace.clientWidth - 2 * GAP - 4) / (this.maxW || 794);
  }

  /** Called when the window changes size: pages side by side follow. */
  resized(): void {
    if (this.placePages(false)) this.onScroll(true);
  }

  onePageZoom(): number {
    const h = this.pages[0]?.h ?? 1123;
    return Math.min(this.pageWidthZoom(), (this.workspace.clientHeight - 2 * GAP) / h);
  }

  /** Makes room for the markup area (comment cards) right of the pages, as Word does. */
  setMarkupWidth(w: number): void {
    if (w === this.markupW) return;
    this.markupW = w;
    if (!this.placePages(false)) this.applySize();
  }

  get markupWidth(): number {
    return this.markupW;
  }

  private applySize(): void {
    const z = this.zoomFactor;
    const w = Math.max(this.areaW, this.maxW + 2 * this.gap) + this.markupW;
    this.markupLayer.style.display = this.markupW ? "" : "none";
    this.markupLayer.style.left = `${GAP + this.maxW + 12}px`;
    this.markupLayer.style.width = `${Math.max(0, this.markupW - 24)}px`;
    const h = this.areaH || 2 * this.gap;
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
      const back = layer(0), selLayer = layer(1), textLayer = layer(2), front = layer(3), selTop = layer(3), overlay = layer(4);
      overlay.style.pointerEvents = "none";
      selTop.style.pointerEvents = "none";
      el.append(back, selLayer, textLayer, front, selTop, overlay);
      this.container.append(el);
      this.pages.push({ el, back, selLayer, selTop, textLayer, front, overlay, lineEls: [], lineKeys: [], decoKey: "", overlayKey: "", live: false, w: 0, h: 0 });
    }
    while (this.pages.length > pages.length) this.pages.pop()!.el.remove();
    let maxW = 0;
    for (const p of pages) maxW = Math.max(maxW, p.setup.width * PT);
    if (maxW !== this.maxW) { this.maxW = maxW; changed = true; }
    pages.forEach((p, i) => {
      const pv = this.pages[i];
      const w = p.setup.width * PT, h = p.setup.height * PT;
      if (pv.w !== w || pv.h !== h) changed = true;
      pv.w = w; pv.h = h;
      // Gitternetzlinien cover the text area (CSS on .drawgrid).
      const st = pv.el.style;
      const ml = `${p.setup.marginLeft * PT}px`, mt = `${p.setup.marginTop * PT}px`, mr = `${p.setup.marginRight * PT}px`, mb = `${p.setup.marginBottom * PT}px`;
      if (st.getPropertyValue("--ml") !== ml) { st.setProperty("--ml", ml); st.setProperty("--mt", mt); st.setProperty("--mr", mr); st.setProperty("--mb", mb); }
    });
    if (!this.placePages(changed || !this.sizer.style.height) && changed) this.applySize();
  }

  /** The page at a point of the page area (px at 100 %): by row, then by column. */
  private pageAt(x: number, y: number): number {
    const n = this.tops.length;
    if (!n) return 0;
    let lo = 0, hi = n - 1;
    while (lo < hi) {
      const mid = (lo + hi + 1) >> 1;
      if (this.tops[mid] - GAP / 2 <= y) lo = mid; else hi = mid - 1;
    }
    // lo is the last page of its row; the row's pages share its top.
    const top = this.tops[lo];
    let first = lo;
    while (first > 0 && this.tops[first - 1] === top) first--;
    let page = first;
    for (let i = first; i <= lo; i++) if (this.lefts[i] - GAP / 2 <= x) page = i;
    return page;
  }

  private pageAtY(y: number): number {
    return this.pageAt(0, y);
  }

  private computeVisible(): void {
    const z = this.zoomFactor;
    const ws = this.workspace;
    const top = ws.scrollTop / z, bottom = top + ws.clientHeight / z;
    const left = ws.scrollLeft / z, right = left + ws.clientWidth / z;
    const n = this.pages.length;
    let from = -1, to = -1;
    for (let i = 0; i < n; i++) {
      const pv = this.pages[i];
      const t = this.tops[i], l = this.lefts[i] ?? 0;
      if (t + pv.h >= top && t <= bottom && l + pv.w >= left && l <= right) { if (from < 0) from = i; to = i; }
    }
    if (from < 0) { from = Math.max(0, this.pageAtY(top)); to = from; }
    // One page more on each side (a row more when pages stand side by side).
    const extra = this.arrangement === "row" ? 1 : this.perRow;
    this.visibleFrom = Math.max(0, from - extra);
    this.visibleTo = Math.min(n - 1, to + extra);
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
    pv.selTop.replaceChildren();
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
    // Seitenfarbe (not printed, as in Word's default).
    const bg = doc.settings.background;
    const color = bg && !this.printing ? `#${(bg.colorTheme ? themeColor(bg.colorTheme, doc.styles.theme.colors) : undefined) ?? bg.color}` : "";
    if (pv.el.style.background !== color) pv.el.style.background = color;
    // Lines of every story on this page: header, body, footer, notes.
    const items: { key: string; story: StoryLayout; pl: PlacedLine }[] = [];
    const add = (key: string | undefined) => {
      if (!key) return;
      const s = layout.stories.get(key);
      if (s) for (const pl of s.placed) if (!pl.hide && pl.page === page.index) items.push({ key, story: s, pl });
    };
    add(page.headerKey);
    for (const pl of mainLines) if (!pl.hide) items.push({ key: "main", story: layout.main, pl });
    add(page.footerKey);
    // Foot- and endnotes: the stories of the note frames on this page.
    const notes = new Set<string>();
    for (const id of page.frames) { const f = layout.frames[id]; if (f.kind === "note") notes.add(f.story); }
    for (const id of notes) add(id);
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
      const lineKey = `${idOf(line)}|${pl.x}|${pl.y}|${this.marks}|${isLast}|${dim}|${idOf(marks)}|${idOf(para)}|${this.markupMode}`;
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
      // Geänderte Zeilen: a bar in the left margin next to lines with tracked changes.
      if ((this.markupMode === "all" || this.markupMode === "simple") && !this.printing && lineHasRevision(para, line, isLast)) {
        const bar = document.createElement("div");
        bar.style.cssText = `position:absolute;left:${page.setup.marginLeft * PT - 14}px;top:${pl.y}px;width:0;height:${line.height}px;` +
          `border-left:1px solid ${this.markupMode === "simple" ? "#d13438" : "#7f7f7f"};pointer-events:none`;
        el.append(bar);
      }
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
    const layout = this.layoutRef!;
    const grid = this.showGridlines && !this.printing
      ? page.frames.map((id) => layout.frames[id]).filter((f) => f.kind === "cell") : [];
    const key = grid.map((f) => `g${f.x},${f.y},${f.w},${f.h}`).join(";") + "|" + page.deco.map((d) => {
      if (d.kind === "rect") return `r${d.x},${d.y},${d.w},${d.h},${d.fill},${this.dimmed(d.story)}`;
      if (d.kind === "line") return `l${d.x1},${d.y1},${d.x2},${d.y2},${d.width},${d.color},${d.style},${this.dimmed(d.story)}`;
      const text = (d.shape?.story ? this.shapeTextKey(d.shape.story, page.index) : "") +
        (d.group ? `g${idOf(d.group)}:${d.group.members.map((m) => (m.shape?.story ? this.shapeTextKey(m.shape.story, page.index) : "")).join("/")}` : "");
      return `i${d.x},${d.y},${d.w},${d.h},${idOf(d.image)},${d.behind},${this.dimmed(d.story)},${d.z ?? 0},${d.pic ? JSON.stringify(d.pic) : ""},${d.shape ? idOf(d.shape) : ""},${text}`;
    }).join(";") + "|" + (page.lineNumbers ?? []).map((n) => `${n.x},${n.y},${n.text}`).join(";") + "|" + this.watermarkKey();
    if (key === pv.decoKey) return;
    pv.decoKey = key;
    const back: HTMLElement[] = [], front: HTMLElement[] = [];
    // Wasserzeichen: behind everything, centred on the margins.
    const wm = this.docRef?.settings.watermark;
    if (wm) back.push(this.watermarkElement(wm, page));
    // Zeilennummern: right-aligned at their x, on the line's baseline.
    for (const n of page.lineNumbers ?? []) {
      const el = document.createElement("span");
      el.textContent = n.text;
      const lh = this.fonts.lineHeight(n.c.font, Fonts.renderSize(n.c)) * PT;
      el.style.cssText = `position:absolute;white-space:pre;right:${pv.w - n.x}px;top:${n.y - this.fonts.baselineOffset(n.c, lh)}px;font:${this.fonts.cssFont(n.c)};` +
        `line-height:${lh}px;color:${n.c.color === "auto" ? "#000" : `#${n.c.color}`};pointer-events:none`;
      back.push(el);
    }
    // Gitternetzlinien only where a cell edge has no border of its own (as in Word).
    const borders = page.deco.filter((d): d is Extract<Deco, { kind: "line" }> => d.kind === "line");
    const covered = (x1: number, y1: number, x2: number, y2: number) => borders.some((l) => {
      const tol = Math.max(1.5, l.width);
      if (y1 === y2) return Math.abs(l.y1 - l.y2) < 0.5 && Math.abs(l.y1 - y1) <= tol && Math.min(l.x1, l.x2) <= x1 + tol && Math.max(l.x1, l.x2) >= x2 - tol;
      return Math.abs(l.x1 - l.x2) < 0.5 && Math.abs(l.x1 - x1) <= tol && Math.min(l.y1, l.y2) <= y1 + tol && Math.max(l.y1, l.y2) >= y2 - tol;
    });
    for (const f of grid) {
      const edges: [string, boolean][] = [
        ["top", !covered(f.x, f.y, f.x + f.w, f.y)], ["bottom", !covered(f.x, f.y + f.h, f.x + f.w, f.y + f.h)],
        ["left", !covered(f.x, f.y, f.x, f.y + f.h)], ["right", !covered(f.x + f.w, f.y, f.x + f.w, f.y + f.h)],
      ];
      if (!edges.some(([, on]) => on)) continue;
      const el = document.createElement("div");
      el.style.cssText = `position:absolute;left:${f.x}px;top:${f.y}px;width:${f.w}px;height:${f.h}px;box-sizing:border-box;pointer-events:none;` +
        edges.map(([side, on]) => `border-${side}:1px ${on ? "dashed #a6c3e8" : "solid transparent"}`).join(";");
      back.push(el);
    }
    // Floating pictures in their stacking order (relativeHeight).
    const ordered = [...page.deco].sort((x, y) => (x.kind === "image" ? x.z ?? 0 : -1e9) - (y.kind === "image" ? y.z ?? 0 : -1e9));
    for (const d of ordered) {
      const el = this.decoElement(d);
      if (this.dimmed(d.story)) el.style.opacity = "0.45";
      (d.kind === "image" && !d.behind ? front : back).push(el);
    }
    pv.back.replaceChildren(...back);
    pv.front.replaceChildren(...front);
  }

  private watermarkKey(): string {
    const wm = this.docRef?.settings.watermark;
    if (!wm) return "";
    return wm.kind === "text" ? JSON.stringify(wm) : `p${idOf(wm.image)},${wm.width},${wm.height},${wm.washout}`;
  }

  /**
   * A watermark as Word draws it: WordArt text stretched to fill its shape (VML fitshape),
   * turned by 315° when diagonal, half transparent; or a picture, washed out.
   */
  private watermarkElement(wm: NonNullable<DocSettings["watermark"]>, page: PageLayout): HTMLElement {
    const s = page.setup;
    const w = wm.width * PT, h = wm.height * PT;
    const cx = (s.marginLeft + (s.width - s.marginLeft - s.marginRight) / 2) * PT, cy = (s.marginTop + (s.height - s.marginTop - s.marginBottom) / 2) * PT;
    const box = document.createElement("div");
    box.style.cssText = `position:absolute;left:${cx - w / 2}px;top:${cy - h / 2}px;width:${w}px;height:${h}px;pointer-events:none;` +
      (wm.kind === "text" && wm.diagonal ? "transform:rotate(-45deg);" : "");
    if (wm.kind === "picture") {
      const img = document.createElement("img");
      img.src = imageUrl(wm.image);
      img.style.cssText = `width:100%;height:100%;${wm.washout ? "filter:brightness(1.55) contrast(0.32);" : ""}`;
      box.append(img);
      return box;
    }
    // The glyphs' own box fills the shape, as VML's fitshape does.
    const g = (View.measureCtx ??= document.createElement("canvas").getContext("2d")!);
    g.font = `100px "${wm.font}"`;
    const m = g.measureText(wm.text);
    const asc = m.actualBoundingBoxAscent || 72, desc = m.actualBoundingBoxDescent || 0;
    const left = m.actualBoundingBoxLeft || 0, width = (m.actualBoundingBoxRight || m.width) + left;
    const ns = "http://www.w3.org/2000/svg";
    const svg = document.createElementNS(ns, "svg");
    svg.setAttribute("viewBox", `${-left} ${-asc} ${width} ${asc + desc}`);
    svg.setAttribute("preserveAspectRatio", "none");
    svg.setAttribute("width", String(w));
    svg.setAttribute("height", String(h));
    const t = document.createElementNS(ns, "text");
    t.setAttribute("x", "0");
    t.setAttribute("y", "0");
    t.setAttribute("font-family", wm.font);
    t.setAttribute("font-size", "100");
    t.setAttribute("fill", `#${wm.color}`);
    if (wm.semi) t.setAttribute("fill-opacity", "0.5");
    t.textContent = wm.text;
    svg.append(t);
    box.append(svg);
    return box;
  }

  private static measureCtx: CanvasRenderingContext2D | null = null;

  /** What a shape's text looks like on a page (for redrawing only when it changed). */
  private shapeTextKey(story: string, page: number): string {
    const sl = this.layoutRef?.stories.get(story);
    const paras = this.storyParas(story);
    if (!sl) return "";
    return `${sl.placed.filter((pl) => pl.page === page).map((pl) => `${idOf(sl.paras[pl.para]?.lines[pl.line])}@${pl.x},${pl.y}`).join(";")}|${paras.map(idOf).join(",")}|${this.marks}|${this.opts.active === story}`;
  }

  /** A shape with its text (in page coordinates, so the lines sit where the layout put them). */
  /** A group: its shapes (with their text) and pictures inside the group's box. */
  private groupDeco(d: Extract<Deco, { kind: "image" }>): HTMLElement {
    const wrap = document.createElement("div");
    wrap.style.cssText = "position:absolute;left:0;top:0;width:0;height:0;pointer-events:none";
    const page = this.layoutRef!.pages.findIndex((pg) => pg.deco.includes(d));
    for (const m of d.group!.members) {
      const md = { ...d, x: d.x + m.x * PT, y: d.y + m.y * PT, w: m.width * PT, h: m.height * PT, shape: m.shape, pic: m.pic, image: m.image, group: undefined };
      if (m.shape) wrap.append(this.shapeDeco(md, page));
      else if (m.image) {
        const el = document.createElement("div");
        el.style.cssText = `position:absolute;left:${md.x}px;top:${md.y}px;width:${md.w}px;height:${md.h}px`;
        el.append(pictureElement(imageUrl(m.image), md.w, md.h, m.pic));
        wrap.append(el);
      }
    }
    return wrap;
  }

  private shapeDeco(d: Extract<Deco, { kind: "image" }>, pageIndex?: number): HTMLElement {
    const wrap = document.createElement("div");
    wrap.style.cssText = "position:absolute;left:0;top:0;width:0;height:0";
    const shape = shapeElement(d.w, d.h, d.pic, d.shape!);
    shape.style.left = `${d.x}px`;
    shape.style.top = `${d.y}px`;
    const rot = shape.style.transform;
    wrap.append(shape);
    const chain = d.shape!.chain;
    const story = d.shape!.story ?? chain?.story;
    const sl = story ? this.layoutRef?.stories.get(story) : undefined;
    if (story && sl) {
      const paras = this.storyParas(story);
      const text = document.createElement("div");
      text.style.cssText = "position:absolute;left:0;top:0;width:0;height:0";
      // Vertical text is turned (not while it is edited, so the caret matches the lines).
      const vert = d.shape!.body.vert;
      const turn = vert && this.opts.active !== story ? (vert === "vert270" ? "rotate(-90deg)" : "rotate(90deg)") : "";
      const tr = shapeTextRect(d.pic, d.shape!, d.w, d.h);
      const tf = [rot, turn].filter(Boolean).join(" ");
      if (tf) { text.style.transform = tf; text.style.transformOrigin = turn && !rot ? `${d.x + tr.x + tr.w / 2}px ${d.y + tr.y + tr.h / 2}px` : `${d.x + d.w / 2}px ${d.y + d.h / 2}px`; }
      const page = pageIndex ?? this.layoutRef!.pages.findIndex((pg) => pg.deco.includes(d));
      const prevColor = this.autoColor;
      this.autoColor = d.shape!.fontColor ? `#${d.shape!.fontColor}` : "#000";
      for (const pl of sl.placed) {
        if (pl.page !== page || pl.hide) continue;
        // Linked boxes: only the lines that flowed into this box.
        if (chain && (pl.x < d.x - 1 || pl.x > d.x + d.w + 1 || pl.y < d.y - 1 || pl.y > d.y + d.h + 1)) continue;
        const para = paras[pl.para];
        const pll = sl.paras[pl.para];
        const line = pll?.lines[pl.line];
        if (!para || !line) continue;
        const marks = this.marksFor?.(story, para);
        const holder = document.createElement("div");
        holder.style.cssText = "position:absolute;left:0;top:0;width:0;height:0";
        holder.append(...this.lineContent(para, line, pl.x + line.x0, pl.y, pl.line === pll.lines.length - 1, marks));
        text.append(holder);
      }
      this.autoColor = prevColor;
      // Text boxes cut off text below their bottom edge, as in Word.
      if (d.shape!.kind === "textbox" && !chain) text.style.clipPath = `inset(${d.y}px calc(100% - ${d.x + d.w}px) calc(100% - ${d.y + d.h}px) ${d.x}px)`;
      wrap.append(text);
    }
    wrap.style.pointerEvents = "none";
    return wrap;
  }

  private decoElement(d: Deco): HTMLElement {
    if (d.kind === "image" && d.group) return this.groupDeco(d);
    if (d.kind === "image" && d.shape) return this.shapeDeco(d);
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
        el.append(pictureElement(imageUrl(d.image), d.w, d.h, d.pic));
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
    const img = this.opts.image && this.opts.image.page === page.index ? this.opts.image : null;
    const tb = this.opts.table !== undefined ? this.tableBounds(page, this.opts.table) : null;
    const others = (this.opts.others ?? []).filter((o) => o.page === page.index);
    const key = others.map((o) => `o${o.x},${o.y},${o.w},${o.h}`).join("") + (editing ? `${page.bodyTop}|${page.bodyBottom}|${page.section}|${sections}|${pv.w}` : "") +
      (img ? `|img${img.x},${img.y},${img.w},${img.h},${img.crop ? `${img.crop.x},${img.crop.y},${img.crop.w},${img.crop.h}` : ""},${img.picture},${img.editing},${(img.adjust ?? []).map((a) => `${a.x},${a.y}`).join(";")},${(img.ends ?? []).map((a) => `${a.x},${a.y}`).join(";")}` : "") + (tb ? `|tb${tb.x},${tb.y},${tb.w},${tb.h},${tb.first},${tb.last}` : "");
    if (key === pv.overlayKey) return;
    pv.overlayKey = key;
    const handles: HTMLElement[] = [];
    for (const o of others) {
      const fr = document.createElement("div");
      fr.style.cssText = `position:absolute;left:${o.x}px;top:${o.y}px;width:${o.w}px;height:${o.h}px;box-sizing:border-box;border:1px solid #6f8db9;pointer-events:none`;
      for (const [hx, hy] of [[0, 0], [0.5, 0], [1, 0], [0, 0.5], [1, 0.5], [0, 1], [0.5, 1], [1, 1]]) {
        const d = document.createElement("div");
        d.style.cssText = `position:absolute;left:${hx * o.w - 4}px;top:${hy * o.h - 4}px;width:7px;height:7px;border-radius:50%;background:#fff;border:1px solid #6f8db9`;
        fr.append(d);
      }
      handles.push(fr);
    }
    if (img && img.crop) {
      // Zuschneiden: the whole picture dimmed, the frame with Word's black crop marks.
      const ghost = document.createElement("img");
      ghost.src = img.crop.url;
      ghost.style.cssText = `position:absolute;left:${img.crop.x}px;top:${img.crop.y}px;width:${img.crop.w}px;height:${img.crop.h}px;opacity:0.45;max-width:none`;
      handles.push(ghost);
      const outline = document.createElement("div");
      outline.style.cssText = `position:absolute;left:${img.crop.x}px;top:${img.crop.y}px;width:${img.crop.w}px;height:${img.crop.h}px;box-sizing:border-box;border:1px solid #9e9e9e`;
      handles.push(outline);
      const frame = document.createElement("div");
      frame.style.cssText = `position:absolute;left:${img.x}px;top:${img.y}px;width:${img.w}px;height:${img.h}px;box-sizing:border-box;border:1px solid #000`;
      handles.push(frame);
      for (const hd of HANDLES) {
        const { x, y } = this.handlePos(img, hd);
        const el = document.createElement("div");
        const corner = hd.length === 2;
        const len = 12, th = 4;
        if (corner) {
          const sx = hd.includes("w") ? 0 : -len + th, sy = hd.startsWith("n") ? 0 : -len + th;
          el.style.cssText = `position:absolute;left:${x + sx - (hd.includes("w") ? th / 2 : th / 2)}px;top:${y + sy - th / 2}px;width:${len}px;height:${len}px;` +
            `border-${hd.startsWith("n") ? "top" : "bottom"}:${th}px solid #000;border-${hd.includes("w") ? "left" : "right"}:${th}px solid #000;box-sizing:border-box`;
        } else if (hd === "n" || hd === "s") el.style.cssText = `position:absolute;left:${x - 8}px;top:${y - th / 2}px;width:16px;height:${th}px;background:#000`;
        else el.style.cssText = `position:absolute;left:${x - th / 2}px;top:${y - 8}px;width:${th}px;height:16px;background:#000`;
        handles.push(el);
      }
    } else if (img?.ends) {
      // Lines: Word's two round handles at the ends (and yellow ones of connectors).
      for (const p of img.ends) {
        const el = document.createElement("div");
        el.style.cssText = `position:absolute;left:${p.x - 4.5}px;top:${p.y - 4.5}px;width:9px;height:9px;box-sizing:border-box;border:1px solid #5f5f5f;background:#fff;border-radius:50%`;
        handles.push(el);
      }
      for (const a of img.adjust ?? []) {
        const el = document.createElement("div");
        el.style.cssText = `position:absolute;left:${a.x - 4.5}px;top:${a.y - 4.5}px;width:9px;height:9px;box-sizing:border-box;border:1px solid #5f5f5f;background:#ffd200;border-radius:50%`;
        handles.push(el);
      }
    } else if (img) {
      const frame = document.createElement("div");
      frame.style.cssText = `position:absolute;left:${img.x}px;top:${img.y}px;width:${img.w}px;height:${img.h}px;box-sizing:border-box;border:1px ${img.editing ? "dashed" : "solid"} #6f8db9`;
      handles.push(frame);
      // Word 365: eight round handles; pictures also get the rotation handle and the layout button.
      for (const hd of HANDLES) {
        const { x, y } = this.handlePos(img, hd);
        const el = document.createElement("div");
        el.style.cssText = `position:absolute;left:${x - 4.5}px;top:${y - 4.5}px;width:9px;height:9px;box-sizing:border-box;border:1px solid #5f5f5f;background:#fff;border-radius:50%`;
        handles.push(el);
      }
      if (img.picture) {
        const rot = document.createElement("div");
        rot.style.cssText = `position:absolute;left:${img.x + img.w / 2 - 9}px;top:${img.y - 26}px;width:18px;height:18px`;
        rot.innerHTML = ROTATE_SVG;
        handles.push(rot);
        const lb = document.createElement("div");
        lb.style.cssText = `position:absolute;left:${img.x + img.w + 10}px;top:${img.y}px;width:22px;height:22px;box-sizing:border-box;border:1px solid #9e9e9e;background:#fff;display:flex;align-items:center;justify-content:center`;
        lb.innerHTML = LAYOUT_SVG;
        lb.title = "Layoutoptionen";
        handles.push(lb);
      }
      // Shapes: Word's yellow handles that change the shape (corner radius, arrow head …).
      for (const a of img.adjust ?? []) {
        const el = document.createElement("div");
        el.style.cssText = `position:absolute;left:${a.x - 4.5}px;top:${a.y - 4.5}px;width:9px;height:9px;box-sizing:border-box;border:1px solid #5f5f5f;background:#ffd200;border-radius:50%`;
        handles.push(el);
      }
    }
    if (tb && !this.printing) {
      // Word's handles: the move handle above the table's top-left corner on the page it starts,
      // the sizing handle below its bottom-right corner on the page it ends.
      if (tb.first) {
        const mh = document.createElement("div");
        mh.style.cssText = `position:absolute;left:${tb.x - 15}px;top:${tb.y - 15}px;width:13px;height:13px;box-sizing:border-box;` +
          "border:1px solid #a0a0a0;background:#fff;display:flex;align-items:center;justify-content:center";
        mh.innerHTML = MOVE_HANDLE_SVG;
        handles.push(mh);
      }
      if (tb.last) {
        const rh = document.createElement("div");
        rh.style.cssText = `position:absolute;left:${tb.x + tb.w + 2}px;top:${tb.y + tb.h + 2}px;width:7px;height:7px;box-sizing:border-box;border:1px solid #a0a0a0;background:#fff`;
        handles.push(rh);
      }
    }
    if (!editing) { pv.overlay.replaceChildren(...handles); return; }
    const sect = sections > 1 ? ` -Abschnitt ${page.section + 1}-` : "";
    const label = (text: string, y: number, above: boolean) => {
      const line = document.createElement("div");
      line.style.cssText = `position:absolute;left:0;width:${pv.w}px;top:${y}px;border-top:1px dashed #2b579a`;
      const tag = document.createElement("div");
      tag.style.cssText = `position:absolute;left:${page.setup.marginLeft * PT - 8}px;${above ? `top:${y - 18}px` : `top:${y + 1}px`};` +
        "font:11px 'Segoe UI',sans-serif;color:#fff;background:#2b579a;padding:1px 6px;white-space:pre";
      tag.textContent = text;
      return [line, tag];
    };
    pv.overlay.replaceChildren(
      ...label(`Kopfzeile${sect}`, page.bodyTop - 2, false),
      ...label(`Fußzeile${sect}`, page.bodyBottom + 2, true),
      ...handles,
    );
  }

  private handlePos(b: { x: number; y: number; w: number; h: number }, hd: Handle): { x: number; y: number } {
    const x = hd.includes("w") ? b.x : hd.includes("e") ? b.x + b.w : b.x + b.w / 2;
    const y = hd.startsWith("n") ? b.y : hd.startsWith("s") ? b.y + b.h : b.y + b.h / 2;
    return { x, y };
  }

  /** The sizing handle of the selected picture under a client point. */
  /** A table's cells on a page: bounding box, whether the table starts and ends on it. */
  tableBounds(page: PageLayout, tid: number): { x: number; y: number; w: number; h: number; first: boolean; last: boolean } | null {
    const layout = this.layoutRef!;
    let x1 = Infinity, y1 = Infinity, x2 = -Infinity, y2 = -Infinity;
    for (const id of page.frames) {
      const f = layout.frames[id];
      if (f.kind !== "cell" || f.story !== "main" || !f.cell || Number(f.cell.split(":")[0]) !== tid) continue;
      x1 = Math.min(x1, f.x); y1 = Math.min(y1, f.y); x2 = Math.max(x2, f.x + f.w); y2 = Math.max(y2, f.y + f.h);
    }
    if (x1 === Infinity) return null;
    const pagesWith = layout.pages.filter((pg) => pg.frames.some((id) => { const f = layout.frames[id]; return f.kind === "cell" && f.cell && Number(f.cell.split(":")[0]) === tid; }));
    return { x: x1, y: y1, w: x2 - x1, h: y2 - y1, first: pagesWith[0]?.index === page.index, last: pagesWith[pagesWith.length - 1]?.index === page.index };
  }

  /** The table whose cells (or handles) are under a client point. */
  tableAt(clientX: number, clientY: number): number | null {
    const layout = this.layoutRef;
    if (!layout) return null;
    const pt = this.pointAt(clientX, clientY);
    const pg = layout.pages[pt.page];
    if (!pg) return null;
    for (const id of pg.frames) {
      const f = layout.frames[id];
      if (f.kind !== "cell" || f.story !== "main" || !f.cell) continue;
      if (pt.x >= f.x - 18 && pt.x <= f.x + f.w + 12 && pt.y >= f.y - 18 && pt.y <= f.y + f.h + 12) return Number(f.cell.split(":")[0]);
    }
    return null;
  }

  /** Word's table move or sizing handle under a client point. */
  tableHandleAt(clientX: number, clientY: number): { kind: "move" | "size"; table: number } | null {
    const tid = this.opts.table;
    if (tid === undefined || !this.layoutRef) return null;
    const pt = this.pointAt(clientX, clientY);
    const pg = this.layoutRef.pages[pt.page];
    const tb = pg ? this.tableBounds(pg, tid) : null;
    if (!tb) return null;
    if (tb.first && pt.x >= tb.x - 16 && pt.x <= tb.x - 1 && pt.y >= tb.y - 16 && pt.y <= tb.y - 1) return { kind: "move", table: tid };
    if (tb.last && pt.x >= tb.x + tb.w + 1 && pt.x <= tb.x + tb.w + 11 && pt.y >= tb.y + tb.h + 1 && pt.y <= tb.y + tb.h + 11) return { kind: "size", table: tid };
    return null;
  }

  handleAt(clientX: number, clientY: number): Handle | null {
    const img = this.opts.image;
    if (!img) return null;
    const pt = this.pointAt(clientX, clientY);
    if (pt.page !== img.page) return null;
    const tol = (img.crop ? 9 : 6) / this.zoomFactor;
    // The yellow handles sit inside the shape and win over the frame's.
    const adj = img.adjust ?? [];
    for (let i = 0; i < adj.length; i++) if (Math.abs(pt.x - adj[i].x) <= tol && Math.abs(pt.y - adj[i].y) <= tol) return `adj-${i}`;
    if (img.ends) {
      for (let i = 0; i < 2; i++) if (Math.abs(pt.x - img.ends[i].x) <= tol && Math.abs(pt.y - img.ends[i].y) <= tol) return i ? "end-1" : "end-0";
      return null;
    }
    for (const hd of HANDLES) {
      const p = this.handlePos(img, hd);
      if (Math.abs(pt.x - p.x) <= tol && Math.abs(pt.y - p.y) <= tol) return img.crop ? `crop-${hd}` as Handle : hd;
    }
    if (img.picture && !img.crop) {
      if (Math.abs(pt.x - (img.x + img.w / 2)) <= 10 && Math.abs(pt.y - (img.y - 17)) <= 10) return "rot";
      if (pt.x >= img.x + img.w + 10 && pt.x <= img.x + img.w + 32 && pt.y >= img.y && pt.y <= img.y + 22) return "layout";
    }
    return null;
  }

  /** Client rectangle of the layout options button of the selected picture. */
  layoutButtonRect(): DOMRect | null {
    const img = this.opts.image;
    if (!img) return null;
    const left = this.pageLeft(img.page) + img.x + img.w + 10, top = this.pageTop(img.page) + img.y;
    const ws = this.workspace.getBoundingClientRect();
    const z = this.zoomFactor;
    return new DOMRect(ws.left + left * z - this.workspace.scrollLeft, ws.top + top * z - this.workspace.scrollTop, 22 * z, 22 * z);
  }

  /** A column border of a table under a client point (for dragging it): the cell left of it. */
  cellBorderAt(clientX: number, clientY: number): { cell: string; x: number } | null {
    const layout = this.layoutRef;
    if (!layout || this.opts.active !== "main") return null;
    const pt = this.pointAt(clientX, clientY);
    const pg = layout.pages[pt.page];
    if (!pg) return null;
    const tol = 3 / this.zoomFactor;
    for (const id of pg.frames) {
      const f = layout.frames[id];
      if (f.kind !== "cell" || !f.cell || f.story !== "main") continue;
      if (pt.y < f.y || pt.y > f.y + f.h) continue;
      if (Math.abs(pt.x - (f.x + f.w)) <= tol) return { cell: f.cell, x: f.x + f.w };
    }
    return null;
  }

  /** The bottom border of a table row under a client point (row heights are dragged there). */
  rowBorderAt(clientX: number, clientY: number): { cell: string; y: number; h: number } | null {
    const layout = this.layoutRef;
    if (!layout || this.opts.active !== "main") return null;
    const pt = this.pointAt(clientX, clientY);
    const pg = layout.pages[pt.page];
    if (!pg) return null;
    const tol = 3 / this.zoomFactor;
    for (const id of pg.frames) {
      const f = layout.frames[id];
      if (f.kind !== "cell" || !f.cell || f.story !== "main") continue;
      if (pt.x < f.x || pt.x > f.x + f.w) continue;
      if (Math.abs(pt.y - (f.y + f.h)) <= tol) return { cell: f.cell, y: f.y + f.h, h: f.h };
    }
    return null;
  }

  private authorColor(author: string): string {
    let c = this.authors.get(author);
    if (!c) { c = AUTHOR_COLORS[this.authors.size % AUTHOR_COLORS.length]; this.authors.set(author, c); }
    return c;
  }

  /** Colour of "auto" text: black, or the text colour of the shape being drawn. */
  private autoColor = "#000";

  private spanStyle(c: ResolvedChar): string {
    const size = Fonts.renderSize(c) * PT;
    const lh = this.fonts.lineHeight(c.font, Fonts.renderSize(c)) * PT;
    const rev = c.rev ? this.authorColor(c.rev.author) : null;
    let s = `position:absolute;white-space:pre;font:${this.fonts.cssFont(c)};line-height:${lh}px;` +
      `color:${rev ?? (c.color === "auto" ? this.autoColor : `#${c.color}`)};font-kerning:${c.kern && c.size >= c.kern ? "normal" : "none"};font-variant-ligatures:none`;
    if (c.spacing) s += `;letter-spacing:${c.spacing * PT}px`;
    if (c.textFx) s += `;${textFxCss(c.textFx, this.docRef?.styles.theme.colors)}`;
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
          const img = pictureElement(imageUrl(o.image), w, h, o.pic);
          const ex = o.pic?.extent;
          img.style.left = `${x + (ex ? ex.l * PT : 0)}px`;
          img.style.top = `${baseline - h - (ex ? ex.b * PT : 0)}px`;
          out.push(img);
        } else if (o.shape) {
          const el = shapeElement(w, h, o.pic, o.shape);
          el.style.left = `${x}px`;
          el.style.top = `${baseline - h}px`;
          out.push(el);
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
        const display = para.text === OBJ && !o.mathInline;
        const el = this.renderMath(o, c.size * PT, display);
        const m = this.measureMath?.(o, c.size, display);
        // The editor finds an equation's drawing by its paragraph and offset.
        mathElements.set(el, { para, o: seg.start });
        const h = (m?.height ?? o.height ?? c.size) * PT, d = (m?.depth ?? o.depth ?? 0) * PT;
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
      if (o?.font) el.style.fontFamily = `"${o.font}"`;
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
    this.printing = true;
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
        el, back: document.createElement("div"), selLayer: document.createElement("div"), selTop: document.createElement("div"), textLayer: document.createElement("div"),
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
    this.printing = false;
  }

  renderSelection(): void {
    const layout = this.layoutRef, doc = this.docRef, sel = this.selRef;
    if (!layout || !doc || !sel) return;
    const [from, to] = cmp(sel.anchor, sel.head) <= 0 ? [sel.anchor, sel.head] : [sel.head, sel.anchor];
    const empty = cmp(from, to) === 0;
    for (let i = this.visibleFrom; i <= this.visibleTo; i++) {
      for (const layer of [this.pages[i]?.selLayer, this.pages[i]?.selTop]) if (layer?.firstChild) layer.replaceChildren();
    }
    const inShape = this.opts.active.startsWith("tx:");
    // A selected picture shows its frame and handles, not a text highlight (Word).
    if (this.opts.image?.picture && !empty && from.p === to.p && to.o - from.o === 1) return;
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
          if (inShape) r.style.mixBlendMode = "multiply";
          (inShape ? this.pages[pl.page].selTop : this.pages[pl.page].selLayer).append(r);
        }
      }
    }
    this.renderCaret(empty && !cells?.size);
  }

  /** Word shows no caret while its window is not the active one. */
  private windowActive = typeof document === "undefined" || document.hasFocus();
  private caretWanted = false;

  /** The formula editor's caret (page px): drawn instead of the text caret while set. */
  caretOverride: { page: number; x: number; y: number; h: number } | null = null;

  /** A page's element (for the formula editor's selection and box). */
  pageElement(i: number): HTMLElement | null {
    return this.pages[i]?.el ?? null;
  }

  /** The drawn element of the equation at offset `o` of a paragraph, and its page. */
  mathElementFor(para: Para, o: number): { el: Element; page: number } | null {
    for (let i = 0; i < this.pages.length; i++) {
      for (const el of Array.from(this.pages[i].textLayer.querySelectorAll("math"))) {
        const m = mathElements.get(el);
        if (m && m.para === para && m.o === o) return { el, page: i };
      }
    }
    return null;
  }

  /** Redraws the caret (after the formula editor moved its own). */
  refreshCaret(): void {
    this.renderCaret(this.caretWanted || this.caretOverride !== null);
  }

  private renderCaret(visible: boolean): void {
    if (!this.layoutRef || !this.docRef || !this.selRef) return;
    if (this.caretOverride) visible = true;
    this.caretWanted = visible;
    if (!visible || !this.windowActive) { this.caret.style.display = "none"; this.stopBlink(); return; }
    const c = this.caretOverride ?? this.caretBox(this.selRef.head, this.selRef.headUpstream);
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
  /** Top of the main-text line holding `pos`, in container coordinates (comment cards). */
  mainLineTop(pos: Pos): number {
    const { pl } = this.placedFor(pos, false, MAIN_KEY);
    return pl ? this.pageTop(pl.page) + pl.y : 0;
  }

  placedFor(pos: Pos, upstream = false, key = this.opts.active): { pl: PlacedLine; index: number } {
    const story = this.storyLayout(key);
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
  /** The page under a client point, or -1 between and beside the pages. */
  pageIndexAt(clientX: number, clientY: number): number {
    if (!this.pages.length) return -1;
    const { page, x, y } = this.pointAt(clientX, clientY);
    const pv = this.pages[page];
    return pv && x >= 0 && y >= 0 && x <= pv.w && y <= pv.h ? page : -1;
  }

  pointAt(clientX: number, clientY: number): { page: number; x: number; y: number } {
    const rect = this.sizer.getBoundingClientRect();
    const z = this.zoomFactor;
    const cy = (clientY - rect.top) / z, cx = (clientX - rect.left) / z;
    const page = Math.min(this.pages.length - 1, this.pageAt(cx, cy));
    return { page, x: cx - this.pageLeft(page), y: cy - this.pageTop(page) };
  }

  /** What part of a page a client point is in. */
  regionAt(clientX: number, clientY: number): { page: number; kind: "header" | "footer" | "body" | "note"; key?: string } {
    const layout = this.layoutRef!;
    const { page, x, y } = this.pointAt(clientX, clientY);
    const pg = layout.pages[page];
    for (const fid of pg.frames) {
      const fr = layout.frames[fid];
      if (fr.kind === "note" && y >= fr.y && y <= fr.y + fr.h && x >= fr.x - 20 && x <= fr.x + fr.w) return { page, kind: "note", key: fr.story };
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

  private draft: { page: number; el: HTMLElement } | null = null;

  /** A shape being drawn (Einfügen > Formen), shown above everything on its page. */
  setDraft(d: { page: number; el: HTMLElement } | null): void {
    this.draft?.el.remove();
    this.draft = d;
    if (!d) return;
    const pv = this.pages[d.page];
    if (!pv) return;
    d.el.style.zIndex = "6";
    d.el.style.pointerEvents = "none";
    pv.el.append(d.el);
  }

  /** Image decoration under a client point (floating pictures), for selection. */
  imageAt(clientX: number, clientY: number): Extract<Deco, { kind: "image" }> | null {
    const { page, x, y } = this.pointAt(clientX, clientY);
    const pg = this.layoutRef?.pages[page];
    if (!pg) return null;
    // Topmost first (stacking order), lines within a few pixels of their stroke.
    const tol = 4 / this.zoomFactor;
    const hits = pg.deco.filter((d): d is Extract<Deco, { kind: "image" }> => d.kind === "image" && d.story === "main" &&
      x >= d.x - tol && x <= d.x + d.w + tol && y >= d.y - tol && y <= d.y + d.h + tol);
    hits.sort((a, b) => (b.z ?? 0) - (a.z ?? 0));
    for (const d of hits) {
      if (!d.shape) { if (x >= d.x && x <= d.x + d.w && y >= d.y && y <= d.y + d.h) return d; continue; }
      const g = shapeGeometry(d.pic, d.shape, d.w, d.h);
      const open = g?.paths.filter((p) => p.fill === "none" || d.shape!.kind === "connector") ?? [];
      const filled = g?.paths.some((p) => p.fill !== "none") && d.shape.kind !== "connector";
      if (filled && x >= d.x && x <= d.x + d.w && y >= d.y && y <= d.y + d.h) return d;
      // Lines and outlines: near the stroke.
      for (const p of open) {
        const pts = flattenPath(p.d, 8);
        for (let k = 1; k < pts.length; k++) {
          const [ax, ay] = pts[k - 1], [bx, by] = pts[k];
          const L = Math.hypot(bx - ax, by - ay) || 1;
          const t = Math.max(0, Math.min(1, ((x - d.x - ax) * (bx - ax) + (y - d.y - ay) * (by - ay)) / (L * L)));
          if (Math.hypot(x - d.x - (ax + t * (bx - ax)), y - d.y - (ay + t * (by - ay))) <= tol + 2) return d;
        }
      }
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
    // Pages side by side (Lesemodus): the caret's page comes into view.
    const left = (this.pageLeft(c.page) + c.x) * z;
    if (left < ws.scrollLeft || left > ws.scrollLeft + ws.clientWidth - 8) {
      if (this.arrangement === "row") this.showScreen(c.page);
      else ws.scrollLeft = Math.max(0, left - ws.clientWidth / 2);
    }
  }

  /** Scrolls a page to the top of the window (navigation pane, Gehe zu). */
  scrollToPage(page: number): void {
    if (this.arrangement === "row") { this.showScreen(page); return; }
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
