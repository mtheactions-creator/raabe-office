// Normalansicht's slide area (D-070): PowerPoint's two selection levels (a selected frame with a
// solid outline, or text being edited with a dashed one and a caret), the eight sizing handles and
// the rotation handle with their modifier keys, moving with Umschalt (one direction), Strg (copy) and
// Alt (no snapping), smart guides to other shapes and the slide, the Auswahlrahmen, nudging with the
// arrow keys and every key of text editing. Typing goes through a hidden textarea (IME included).

import type { Pptx, ShapeView, SlideView } from "@raabe/ooxml";
import { PptxDom, slideView } from "@raabe/ooxml";
import type { Deck } from "./deck";
import { renderSlide, type Rendered } from "./render/slide";
import { AUTOFIT_STEPS, EMU_PX, renderText, textHeight } from "./render/text";
import { presetAdjust, presetGeometry, shapeTextRect } from "@raabe/text-engine";
import * as T from "./textops";
import * as TC from "./tablecells";
import { AutoCorrect } from "@raabe/text-engine";
import { removeShapeEl } from "./equations";
import { samplePath } from "./animations";
import * as S from "./shapeops";
import type { Pos } from "./textops";

const ZWSP = String.fromCharCode(0x200b);

export type Tool = { kind: "shape"; prst: string; head?: boolean; tail?: boolean; freeform?: string } | { kind: "textbox"; vertical?: boolean }
  | { kind: "ink"; color: string; width: number; alpha?: number } | { kind: "eraser" } | null;

export interface EditorHost {
  readonly deck: Deck;
  slidePath(): string | null;
  slideNumber(): number;
  image(path: string): string | undefined;
  selectionChanged(): void;
  contextMenu(e: MouseEvent, kind: "shape" | "text" | "slide"): void;
  shortcut(e: KeyboardEvent): boolean;
  doubleClick(v: ShapeView): boolean;
  /** Comment markers of a slide (slide coordinates in EMU). */
  commentMarkers?(path: string): { x: number; y: number; n: number }[];
  toolDone(): void;
  /** Motion paths of a shape (PowerPoint notation), drawn while it is selected. */
  motionPaths?(path: string, id: number): string[];
  /** Bild ab / Bild auf, Pos1, Ende with nothing selected: another slide (PowerPoint). */
  goSlide?(key: "PageDown" | "PageUp" | "Home" | "End"): void;
}

interface TextState {
  id: number;
  /** A table cell being edited (row, column). */
  cell?: { r: number; c: number };
  /** A range of cells (drag across cells, Umschalt+Klick, Umschalt+Pfeil past the cell's text). */
  range?: TC.CellRange;
  caret: Pos;
  anchor: Pos;
  /** Formatting set at a caret for the text typed next (Strg+Umschalt+F with nothing selected). */
  pending: Element | null;
  goalX?: number;
}

type Handle = "nw" | "n" | "ne" | "e" | "se" | "s" | "sw" | "w" | "rot";

const HANDLE_POS: Record<Exclude<Handle, "rot">, [number, number]> = { nw: [0, 0], n: [0.5, 0], ne: [1, 0], e: [1, 0.5], se: [1, 1], s: [0.5, 1], sw: [0, 1], w: [0, 0.5] };
const CURSORS: Record<string, string> = { nw: "nwse-resize", se: "nwse-resize", ne: "nesw-resize", sw: "nesw-resize", n: "ns-resize", s: "ns-resize", e: "ew-resize", w: "ew-resize" };
/** Smart guides snap within this many screen pixels (PowerPoint's distance, estimated). */
const SNAP_PX = 5;

export class SlideEditor {
  readonly el: HTMLElement;
  private canvas: HTMLElement;
  private stage: HTMLElement;
  private overlay: HTMLElement;
  readonly input: HTMLTextAreaElement;
  view: SlideView | null = null;
  private rendered: Rendered | null = null;
  /** Selected shapes (ids); several with Umschalt/Strg. */
  sel: number[] = [];
  text: TextState | null = null;
  tool: Tool = null;
  zoom = 100;
  fit = true;
  /** Autofit steps found while editing, by shape id. */
  private fits = new Map<number, { fontScale: number; lnSpcReduction: number }>();
  private blinkTimer = 0;
  /** Zuschneiden: the picture being cropped (its crop handles replace the sizing handles). */
  crop: number | null = null;
  /** Smart guides are on (Ansicht > Raster und Führungslinien). */
  smartGuides = true;
  snapToGrid = false;
  gridSpacing = 72008;
  /** Ansicht > Lineal: rulers above and beside the slide in cm, 0 at the slide's centre (as PowerPoint). */
  showRuler = false;
  /** Drawn over the page and not part of it (the Handzettelmaster's slide places). */
  pageDecor: ((size: { cx: number; cy: number }) => HTMLElement | null) | null = null;
  showGrid = false;
  showGuides = false;
  private composing = false;

  constructor(readonly host: EditorHost) {
    this.el = h("div", "editor");
    this.canvas = h("div", "canvas");
    this.stage = h("div", "stage");
    this.overlay = h("div", "overlay");
    this.canvas.append(this.stage, this.overlay);
    this.el.append(this.canvas);
    this.input = document.createElement("textarea");
    this.input.className = "ime";
    this.input.setAttribute("aria-label", "Folie");
    this.input.spellcheck = false;
    this.el.append(this.input);
    this.el.addEventListener("mousedown", (e) => this.onDown(e));
    // Double clicks are taken from the click count in onDown: the first click redraws the slide, after
    // which the browser's dblclick may not arrive.
    this.el.addEventListener("contextmenu", (e) => this.onContext(e));
    this.el.addEventListener("mousemove", (e) => this.hover(e));
    this.input.addEventListener("keydown", (e) => this.onKey(e));
    this.input.addEventListener("compositionstart", () => { this.composing = true; });
    this.input.addEventListener("compositionend", () => { this.composing = false; this.flushInput(); });
    this.input.addEventListener("input", () => { if (!this.composing) this.flushInput(); });
    this.input.addEventListener("paste", (e) => e.preventDefault());
    new ResizeObserver(() => { if (this.fit) this.fitZoom(); else this.layout(); }).observe(this.el);
  }

  get deck(): Deck { return this.host.deck; }
  get pptx(): Pptx { return this.host.deck.pptx; }
  get path(): string | null { return this.host.slidePath(); }
  get scale(): number { return this.zoom / 100; }

  /** The page being edited: a slide's size, or the notes page's for the Notizenmaster. */
  get pageSize(): { cx: number; cy: number } {
    if (this.path && /(notesMasters|handoutMasters)\//.test(this.path)) {
      const n = PptxDom.at(this.pptx.presentation, "notesSz");
      return { cx: PptxDom.num(n, "cx") ?? 6858000, cy: PptxDom.num(n, "cy") ?? 9144000 };
    }
    return this.pptx.slideSize;
  }

  focus(): void { this.input.focus({ preventScroll: true }); }

  // ----- drawing -----------------------------------------------------------------------------------

  /** Draws the slide again (after any change) with the selection and caret. */
  render(): void {
    const path = this.path;
    if (!path) { this.stage.replaceChildren(); this.overlay.replaceChildren(); this.view = null; this.rendered = null; return; }
    this.view = this.deck.view(path);
    // Drop selected ids that no longer exist (undo, delete).
    const all = new Set(flat(this.view.shapes).map((s) => s.id));
    this.sel = this.sel.filter((id) => all.has(id));
    if (this.text && !all.has(this.text.id)) this.text = null;
    const sel = this.text ? { id: this.text.id, cell: this.text.cell, cells: this.text.range, sel: T.cmpPos(this.text.caret, this.text.anchor) ? { from: T.minPos(this.text.caret, this.text.anchor), to: T.maxPos(this.text.caret, this.text.anchor) } : null } : null;
    this.rendered = renderSlide(this.view, this.pageSize, { editing: true, slideNumber: this.host.slideNumber(), image: (p) => this.host.image(p), selection: sel, fits: this.fits, pptx: this.pptx });
    const decor = this.pageDecor?.(this.pageSize);
    if (decor) this.rendered.el.append(decor);
    if (this.showGrid) this.rendered.el.append(this.gridLayer());
    this.stage.replaceChildren(this.rendered.el);
    this.layout();
    this.drawOverlay();
    this.placeCaret();
  }

  private gridLayer(): HTMLElement {
    const size = this.pageSize;
    const g = h("div", "grid-layer");
    const step = this.gridSpacing / EMU_PX;
    g.style.cssText = `position:absolute;inset:0;pointer-events:none;background-image:radial-gradient(circle, #808080 0.6px, transparent 0.8px);background-size:${step}px ${step}px;background-position:${(size.cx / EMU_PX / 2) % step}px ${(size.cy / EMU_PX / 2) % step}px`;
    return g;
  }

  fitZoom(): void {
    const size = this.pageSize;
    const w = size.cx / EMU_PX, hh = size.cy / EMU_PX;
    const r = this.el.getBoundingClientRect();
    if (!r.width || !r.height) return;
    // PowerPoint leaves a margin around the slide when it fits it to the window.
    const room = this.showRuler ? 48 : 0;
    const z = Math.min((r.width - 48 - room) / w, (r.height - 32 - room) / hh) * 100;
    this.zoom = Math.max(10, Math.floor(z));
    this.layout();
    this.drawOverlay();
  }

  setZoom(z: number): void {
    this.fit = false;
    this.zoom = Math.max(10, Math.min(400, Math.round(z)));
    this.layout();
    this.drawOverlay();
    this.placeCaret();
  }

  private layout(): void {
    const size = this.pageSize;
    const w = (size.cx / EMU_PX) * this.scale, hh = (size.cy / EMU_PX) * this.scale;
    const r = this.el.getBoundingClientRect();
    const room = this.showRuler ? 24 : 0;
    const padX = Math.max(24 + room, (r.width - w) / 2), padY = Math.max(16 + room, (r.height - hh) / 2);
    this.canvas.style.cssText = `position:relative;width:${w + padX * 2}px;height:${hh + padY * 2}px`;
    this.stage.style.cssText = `position:absolute;left:${padX}px;top:${padY}px;transform:scale(${this.scale});transform-origin:0 0;--s:${this.scale}`;
    this.overlay.style.cssText = `position:absolute;left:${padX}px;top:${padY}px;width:${w}px;height:${hh}px;pointer-events:none`;
    this.canvas.querySelectorAll(".ruler").forEach((x) => x.remove());
    if (this.showRuler) this.canvas.append(this.ruler("h", padX, padY, w, size.cx), this.ruler("v", padX, padY, hh, size.cy));
  }

  private ruler(dir: "h" | "v", padX: number, padY: number, len: number, emu: number): HTMLCanvasElement {
    const c = document.createElement("canvas");
    c.className = "ruler";
    const dpr = window.devicePixelRatio || 1, thick = 18;
    const w = dir === "h" ? len : thick, hh = dir === "h" ? thick : len;
    c.width = Math.round(w * dpr); c.height = Math.round(hh * dpr);
    c.style.cssText = `position:absolute;width:${w}px;height:${hh}px;${dir === "h" ? `left:${padX}px;top:${Math.max(0, padY - thick - 4)}px` : `left:${Math.max(0, padX - thick - 4)}px;top:${padY}px`}`;
    const g = c.getContext("2d");
    if (!g) return c;
    g.scale(dpr, dpr);
    g.fillStyle = "#ffffff"; g.fillRect(0, 0, w, hh);
    g.strokeStyle = "#8a8a8a"; g.fillStyle = "#505050"; g.font = "9px Segoe UI, sans-serif"; g.textAlign = "center"; g.textBaseline = "middle";
    const pxPerCm = (360000 / EMU_PX) * this.scale, mid = len / 2, half = emu / 360000 / 2;
    for (let k = -Math.ceil(half * 4); k <= Math.ceil(half * 4); k++) {
      const pos = mid + (k / 4) * pxPerCm;
      if (pos < 0 || pos > len) continue;
      const cm = k / 4, whole = k % 4 === 0, halfMark = k % 2 === 0;
      if (whole && k !== 0) {
        const label = String(Math.abs(cm));
        if (dir === "h") g.fillText(label, pos, thick / 2); else { g.save(); g.translate(thick / 2, pos); g.rotate(-Math.PI / 2); g.fillText(label, 0, 0); g.restore(); }
        continue;
      }
      const t = k === 0 ? thick * 0.6 : halfMark ? thick * 0.35 : thick * 0.18;
      g.beginPath();
      if (dir === "h") { g.moveTo(pos + 0.5, (thick - t) / 2); g.lineTo(pos + 0.5, (thick + t) / 2); } else { g.moveTo((thick - t) / 2, pos + 0.5); g.lineTo((thick + t) / 2, pos + 0.5); }
      g.stroke();
    }
    return c;
  }

  /** The shape view of an id (members of groups included). */
  shapeView(id: number): ShapeView | undefined {
    return this.view ? flat(this.view.shapes).find((s) => s.id === id) : undefined;
  }

  get selectedViews(): ShapeView[] {
    return this.sel.map((id) => this.shapeView(id)).filter((v): v is ShapeView => !!v);
  }

  private frameEl(v: ShapeView, editing: boolean, primary: boolean): HTMLElement {
    const s = this.scale;
    const f = h("div", `frame${editing ? " editing" : ""}${v.kind === "cxnSp" ? " line" : ""}`);
    f.style.cssText = `left:${(v.x / EMU_PX) * s}px;top:${(v.y / EMU_PX) * s}px;width:${(v.w / EMU_PX) * s}px;height:${(v.h / EMU_PX) * s}px;transform:rotate(${v.rot}deg)`;
    f.dataset.id = String(v.id);
    // The outline itself can be grabbed to move the shape (PowerPoint's four-headed arrow).
    for (const side of ["t", "r", "b", "l"]) { const b = h("div", `edge edge-${side}`); b.dataset.edge = side; f.append(b); }
    if (!primary && !editing) return f;
    if (this.crop === v.id && v.pic) {
      // Crop mode: the whole picture faded behind, black crop handles on the kept part (PowerPoint).
      const c = v.pic.crop, s2 = this.scale;
      const fw = v.w / Math.max(0.01, 1 - c.l - c.r), fh = v.h / Math.max(0.01, 1 - c.t - c.b);
      const ghost = h("div", "crop-ghost");
      ghost.style.cssText = `left:${(-c.l * fw / EMU_PX) * s2}px;top:${(-c.t * fh / EMU_PX) * s2}px;width:${(fw / EMU_PX) * s2}px;height:${(fh / EMU_PX) * s2}px`;
      const url = v.pic.path ? this.host.image(v.pic.path) : undefined;
      if (url) ghost.style.backgroundImage = `url("${url}")`;
      f.append(ghost);
      for (const k of Object.keys(HANDLE_POS) as Exclude<Handle, "rot">[]) {
        const hd = this.handle(k, v);
        hd.classList.add("h-crop");
        hd.dataset.handle = `crop-${k}`;
        f.append(hd);
      }
      return f;
    }
    if (v.kind === "cxnSp") {
      for (const k of ["nw", "se"] as const) f.append(this.handle(k, v));
    } else {
      for (const k of Object.keys(HANDLE_POS) as Exclude<Handle, "rot">[]) {
        // Side handles only on sides long enough (measured: none on a 29 pt high text box at 66 %).
        if ((v.w / EMU_PX) * s < 30 && (k === "n" || k === "s")) continue;
        if ((v.h / EMU_PX) * s < 30 && (k === "e" || k === "w")) continue;
        f.append(this.handle(k, v));
      }
      if (v.kind !== "graphicFrame" || v.frame !== "table") f.append(this.handle("rot", v));
      // The yellow handles that change a preset's adjust values (Rechteck: abgerundete Ecken, arrows ...).
      if (v.kind === "sp" && !v.ph && v.geometry !== "custom" && primary) {
        const w = v.w / EMU_PX, hh = v.h / EMU_PX;
        const g = presetGeometry(v.geometry, w, hh, { ...presetAdjust(v.geometry), ...(v.info.adjust ?? {}) });
        g?.handles.forEach((hd, k) => {
          const el = h("div", "handle h-adj");
          el.dataset.handle = `adj${k}`;
          el.dataset.id = String(v.id);
          el.style.left = `${((v.flipH ? w - hd.x : hd.x) / w) * 100}%`;
          el.style.top = `${((v.flipV ? hh - hd.y : hd.y) / hh) * 100}%`;
          f.append(el);
        });
      }
    }
    return f;
  }

  private handle(k: Handle, v: ShapeView): HTMLElement {
    const el = h("div", `handle h-${k}`);
    el.dataset.handle = k;
    el.dataset.id = String(v.id);
    if (k === "rot") el.innerHTML = '<svg viewBox="0 0 16 16" width="16" height="16"><path d="M12.6 5.2A5.5 5.5 0 1 0 13.5 9" fill="none" stroke="currentColor" stroke-width="1.6"/><path d="M13.9 1.8v4.3H9.6" fill="none" stroke="currentColor" stroke-width="1.6"/></svg>';
    else {
      const [px, py] = HANDLE_POS[k];
      el.style.left = `${px * 100}%`;
      el.style.top = `${py * 100}%`;
      // The pointer follows the shape's rotation (PowerPoint turns its sizing pointers too).
      el.style.cursor = rotatedCursor(k, v.rot);
    }
    return el;
  }

  drawOverlay(): void {
    this.overlay.replaceChildren();
    if (!this.view) return;
    const views = this.selectedViews;
    views.forEach((v, i) => this.overlay.append(this.frameEl(v, this.text?.id === v.id, i === 0 || views.length > 1)));
    // Animationspfade of the selected shapes: a dotted line from the shape's centre, green start, red end (as PowerPoint).
    for (const v of views) for (const mp of (this.path && this.host.motionPaths?.(this.path, v.id)) || []) this.overlay.append(this.motionPathEl(v, mp));
    for (const m of this.path && this.host.commentMarkers ? this.host.commentMarkers(this.path) : []) {
      const el = h("div", "comment-marker");
      el.textContent = String(m.n);
      el.style.left = `${(m.x / EMU_PX) * this.scale}px`;
      el.style.top = `${(m.y / EMU_PX) * this.scale}px`;
      this.overlay.append(el);
    }
  }

  private motionPathEl(v: ShapeView, path: string): SVGSVGElement {
    const size = this.pageSize, s = this.scale;
    const w = (size.cx / EMU_PX) * s, hh = (size.cy / EMU_PX) * s;
    const cx = ((v.x + v.w / 2) / EMU_PX) * s, cy = ((v.y + v.h / 2) / EMU_PX) * s;
    const pts = samplePath(path).map(([x, y]) => `${(cx + x * w).toFixed(1)},${(cy + y * hh).toFixed(1)}`);
    const NS = "http://www.w3.org/2000/svg";
    const svg = document.createElementNS(NS, "svg");
    svg.setAttribute("class", "motion-path");
    svg.setAttribute("width", String(w)); svg.setAttribute("height", String(hh));
    const line = document.createElementNS(NS, "polyline");
    line.setAttribute("points", pts.join(" ")); line.setAttribute("fill", "none"); line.setAttribute("stroke", "#7f7f7f"); line.setAttribute("stroke-dasharray", "3 3");
    svg.append(line);
    const end = (p: string, color: string) => { const [x, y] = p.split(",").map(Number); const c = document.createElementNS(NS, "circle"); c.setAttribute("cx", String(x)); c.setAttribute("cy", String(y)); c.setAttribute("r", "4"); c.setAttribute("fill", color); svg.append(c); };
    end(pts[0], "#00b050"); end(pts[pts.length - 1], "#ff0000");
    return svg;
  }

  // ----- caret -----------------------------------------------------------------------------------------

  private textEl(id: number): HTMLElement | null {
    const cell = this.text?.id === id ? this.text.cell : undefined;
    if (cell) return (this.rendered?.shapes.get(id)?.querySelector(`td[data-r="${cell.r}"][data-c="${cell.c}"] .txt`) as HTMLElement | null) ?? null;
    return (this.rendered?.shapes.get(id)?.querySelector(":scope > .txt") as HTMLElement | null) ?? null;
  }

  /** Local position (in the text frame, unscaled px) and height of a text position. */
  caretBox(id: number, pos: Pos): { x: number; y: number; h: number } | null {
    const txt = this.textEl(id);
    if (!txt) return null;
    const para = txt.querySelector(`.para[data-p="${pos.p}"]`) as HTMLElement | null;
    if (!para) return null;
    const spans = Array.from(para.querySelectorAll(":scope > [data-o]")) as HTMLElement[];
    let target: HTMLElement | null = null, local = 0;
    for (const sp of spans) {
      const o = Number(sp.dataset.o);
      const len = sp.classList.contains("end") || sp.classList.contains("prompt") ? 0 : (sp.textContent ?? "").length;
      if (o <= pos.o && pos.o <= o + len) {
        // At a boundary the caret stays with the text before (not after a line break).
        if (target && pos.o === o && !target.classList.contains("br")) break;
        target = sp; local = pos.o - o;
        if (pos.o < o + len) break;
      }
    }
    target ??= spans[spans.length - 1] ?? null;
    if (!target) return null;
    if (target.classList.contains("prompt")) local = 0;
    const marker = document.createElement("span");
    marker.textContent = ZWSP;
    marker.className = "cm";
    let placed = false;
    if (!target.classList.contains("end") && !target.classList.contains("prompt") && !target.classList.contains("br")) {
      const walker = document.createTreeWalker(target, NodeFilter.SHOW_TEXT);
      let n: Node | null, acc = 0;
      while ((n = walker.nextNode())) {
        const t = n as Text;
        if (local <= acc + t.length) {
          const rest = t.splitText(local - acc);
          rest.parentNode!.insertBefore(marker, rest);
          placed = true;
          break;
        }
        acc += t.length;
      }
    }
    if (!placed) {
      if (target.classList.contains("br") && local > 0) target.after(marker);
      else target.prepend(marker);
    }
    let x = 0, y = 0, el: HTMLElement | null = marker;
    while (el && el !== txt) { x += el.offsetLeft; y += el.offsetTop; el = el.offsetParent as HTMLElement | null; }
    const hh = marker.offsetHeight;
    const parent = marker.parentNode as HTMLElement;
    marker.remove();
    parent.normalize();
    return { x, y, h: hh };
  }

  private placeCaret(): void {
    clearInterval(this.blinkTimer);
    this.el.querySelectorAll(".caret").forEach((c) => c.remove());
    const t = this.text;
    if (!t || t.range || T.cmpPos(t.caret, t.anchor) !== 0) { this.placeInput(null); return; }
    const box = this.caretBox(t.id, t.caret);
    const txt = this.textEl(t.id);
    if (!box || !txt) return;
    const c = h("div", "caret");
    c.style.cssText = `left:${box.x}px;top:${box.y}px;height:${box.h}px;width:${Math.max(1, 1 / this.scale)}px`;
    txt.append(c);
    this.blinkTimer = window.setInterval(() => c.classList.toggle("off"), 530);
    this.placeInput(c);
  }

  /** The hidden textarea sits at the caret so the IME window opens there. */
  private placeInput(caret: HTMLElement | null): void {
    const r = (caret ?? this.stage).getBoundingClientRect(), er = this.el.getBoundingClientRect();
    this.input.style.left = `${r.left - er.left + this.el.scrollLeft}px`;
    this.input.style.top = `${r.top - er.top + this.el.scrollTop}px`;
  }

  /** The client point of a local point in a text frame (through all transforms). */
  private localToClient(id: number, x: number, y: number): { x: number; y: number } | null {
    const txt = this.textEl(id);
    if (!txt) return null;
    const probe = h("div", "probe");
    probe.style.cssText = `position:absolute;left:${x}px;top:${y}px;width:0;height:0`;
    txt.append(probe);
    const r = probe.getBoundingClientRect();
    probe.remove();
    return { x: r.left, y: r.top };
  }

  /** The text position at a client point inside shape `id`'s text. */
  posAt(id: number, cx: number, cy: number): Pos | null {
    const txt = this.textEl(id);
    if (!txt) return null;
    const range = document.caretRangeFromPoint?.(cx, cy);
    const paras = this.textBody(id) ? T.paras(this.textBody(id)!) : [];
    if (!range || !txt.contains(range.startContainer)) {
      // Above the text: its start; below or beside: the nearest paragraph's end.
      const r = txt.getBoundingClientRect();
      if (cy < r.top + 2) return { p: 0, o: 0 };
      return T.endPos(this.textBody(id)!);
    }
    let node: Node | null = range.startContainer;
    let off = range.startOffset;
    const spanOf = (n: Node | null) => (n instanceof HTMLElement ? n : n?.parentElement)?.closest("[data-o]") as HTMLElement | null;
    let span = spanOf(node);
    const para = (node instanceof HTMLElement ? node : node?.parentElement)?.closest(".para") as HTMLElement | null;
    if (!para) return T.endPos(this.textBody(id)!);
    const p = Number(para.dataset.p);
    if (!span) {
      // In the paragraph's box but between spans: the child at the offset.
      const kids = Array.from(para.childNodes);
      const k = kids[Math.min(off, kids.length - 1)] as HTMLElement | undefined;
      span = k?.closest?.("[data-o]") as HTMLElement | null;
      node = null; off = 0;
      if (!span) return { p, o: off };
    }
    const base = Number(span.dataset.o);
    if (span.classList.contains("end") || span.classList.contains("prompt") || span.classList.contains("bullet")) return { p, o: span.classList.contains("bullet") ? 0 : base };
    if (span.classList.contains("br")) return { p, o: base + (off > 0 ? 1 : 0) };
    let acc = 0;
    if (node && node.nodeType === Node.TEXT_NODE) {
      const walker = document.createTreeWalker(span, NodeFilter.SHOW_TEXT);
      let n: Node | null;
      while ((n = walker.nextNode()) && n !== node) acc += (n as Text).length;
      acc += off;
    } else acc = off > 0 ? (span.textContent ?? "").length : 0;
    const max = paras[p] ? T.paraLength(paras[p]) : 0;
    return { p, o: Math.min(max, base + acc) };
  }

  // ----- the model ----------------------------------------------------------------------------------------

  private get doc(): Document | null { return this.path ? this.pptx.doc(this.path) : null; }

  shapeEl(id: number): Element | undefined {
    const d = this.doc;
    return d ? S.findShape(S.treeOf(d), id) : undefined;
  }

  textBody(id: number): Element | undefined {
    const el = this.shapeEl(id);
    const cell = this.text?.id === id ? this.text.cell : undefined;
    if (cell && el) return cellBody(el, cell.r, cell.c);
    return el ? S.txBodyOf(el) : undefined;
  }

  /** Changes the current slide in one undo step and draws it again. */
  edit(label: string, fn: () => void, merge?: string): void {
    const path = this.path;
    if (!path) return;
    // Autofit follows the text in the same undo step.
    this.deck.change(label, [path], () => { fn(); if (this.text) this.updateAutofit(this.text.id); }, merge);
    this.render();
    this.host.selectionChanged();
  }

  // ----- selection ------------------------------------------------------------------------------------------

  select(ids: number[]): void {
    this.endText();
    if (this.crop !== null && !ids.includes(this.crop)) this.crop = null;
    this.sel = ids;
    this.render();
    this.host.selectionChanged();
  }

  clear(): void {
    this.endText();
    this.crop = null;
    this.sel = [];
    this.render();
    this.host.selectionChanged();
  }

  /** Text editing of shape `id`, the caret at `pos` (end when missing), or all text selected. */
  enterText(id: number, pos?: Pos | "all", cell?: { r: number; c: number }): void {
    const el = this.shapeEl(id);
    if (!el) return;
    if (cell) {
      const body = cellBody(el, cell.r, cell.c);
      if (!body) return;
      const end = T.endPos(body);
      const caret = pos === "all" || !pos ? end : pos;
      this.sel = [id];
      this.text = { id, cell, caret, anchor: pos === "all" ? { p: 0, o: 0 } : caret, pending: null };
      this.deck.breakMerge();
      this.render();
      this.host.selectionChanged();
      return;
    }
    if (!S.txBodyOf(el)) {
      if (el.localName !== "sp") return;
      this.deck.change("Text", [this.path!], () => { S.ensureTxBody(el); });
    }
    const body = S.txBodyOf(el)!;
    const end = T.endPos(body);
    const caret = pos === "all" || !pos ? end : pos;
    const anchor = pos === "all" ? { p: 0, o: 0 } : caret;
    this.sel = [id];
    this.text = { id, caret, anchor, pending: null };
    this.deck.breakMerge();
    this.render();
    this.host.selectionChanged();
  }

  /** Leaves text editing; the frame stays selected (Esc) unless `deselect`. */
  endText(): void {
    if (!this.text) return;
    this.text = null;
    this.deck.breakMerge();
  }

  get mode(): "none" | "frame" | "text" {
    return this.text ? "text" : this.sel.length ? "frame" : "none";
  }

  // ----- mouse ------------------------------------------------------------------------------------------------

  /** Client point → slide coordinates in EMU. */
  toSlide(cx: number, cy: number): { x: number; y: number } {
    const r = this.overlay.getBoundingClientRect();
    return { x: ((cx - r.left) / this.scale) * EMU_PX, y: ((cy - r.top) / this.scale) * EMU_PX };
  }

  /** The top-level shape (or, inside a selected group, the member) under a client point. */
  /**
   * What a click hits, as PowerPoint decides it (behaviour catalogue D-072): the letters of a text (a run,
   * a bullet, the prompt) start editing; otherwise the topmost shape whose outline is within 4 screen
   * pixels, whose fill is under the pointer, or, without a fill (text boxes, placeholders), whose line of
   * text is. Empty areas of an unfilled shape are passed through. Lines are hit within 4 pixels.
   */
  private hit(e: MouseEvent): { v: ShapeView; text: boolean } | null {
    const t = document.elementFromPoint(e.clientX, e.clientY) as HTMLElement | null;
    if (!t || !this.el.contains(t)) return null;
    const glyph = t.closest(".txt [data-o], .txt .bullet") as HTMLElement | null;
    if (glyph && this.stage.contains(glyph)) {
      const hit = this.glyphHit(glyph);
      if (hit) return hit;
    }
    if (!this.view) return null;
    const p = this.toSlide(e.clientX, e.clientY);
    const tol = (4 / this.scale) * EMU_PX;
    const shapes = this.view.shapes.filter((s) => !s.hidden);
    for (let i = shapes.length - 1; i >= 0; i--) {
      const r = this.hitShape(shapes[i], p, tol, e);
      if (r) return r;
    }
    return null;
  }

  private hitShape(v: ShapeView, p: { x: number; y: number }, tol: number, e: MouseEvent): { v: ShapeView; text: boolean } | null {
    if (v.kind === "grpSp" && v.children) {
      for (let i = v.children.length - 1; i >= 0; i--) {
        const r = this.hitShape(v.children[i], p, tol, e);
        if (r) return this.sel.includes(v.id) || this.text?.id === v.id ? r : { v, text: false };
      }
      return null;
    }
    // The point in the shape's own frame (turned back about its centre).
    const cx = v.x + v.w / 2, cy = v.y + v.h / 2, a = (-v.rot * Math.PI) / 180;
    const dx = p.x - cx, dy = p.y - cy;
    const lx = dx * Math.cos(a) - dy * Math.sin(a), ly = dx * Math.sin(a) + dy * Math.cos(a);
    const hw = v.w / 2, hh = v.h / 2;
    if (v.kind === "cxnSp" || v.geometry === "line" || v.geometry.startsWith("straightConnector")) {
      const x1 = v.flipH ? hw : -hw, y1 = v.flipV ? hh : -hh;
      return segmentDistance(lx, ly, x1, y1, -x1, -y1) <= tol ? { v, text: false } : null;
    }
    if (Math.abs(lx) > hw + tol || Math.abs(ly) > hh + tol) return null;
    if (Math.abs(lx) > hw - tol || Math.abs(ly) > hh - tol) return { v, text: false };
    // Inside: the text being edited takes the click (caret), an empty placeholder starts editing.
    if (this.text?.id === v.id) return { v, text: true };
    if (v.ph && v.text?.empty) return { v, text: true };
    if (this.sel.includes(v.id) || filled(v)) return { v, text: false };
    return this.onTextLine(v.id, e) ? { v, text: false } : null;
  }

  /** Whether the pointer is on one of the shape's lines of text (across the text area's width). */
  private onTextLine(id: number, e: MouseEvent): boolean {
    const txt = this.rendered?.shapes.get(id)?.querySelector(":scope > .txt") as HTMLElement | null;
    if (!txt) return false;
    const area = txt.getBoundingClientRect();
    if (e.clientX < area.left || e.clientX > area.right) return false;
    return Array.from(txt.querySelectorAll(".para")).some((p) => { const r = p.getBoundingClientRect(); return e.clientY >= r.top && e.clientY <= r.bottom; });
  }

  private glyphHit(t: HTMLElement): { v: ShapeView; text: boolean } | null {
    let shp = t.closest(".shp") as HTMLElement | null;
    if (!shp) return null;
    // Climb to the outermost shape unless its group is selected (then its member).
    let top = shp;
    for (let p = shp.parentElement?.closest(".shp") as HTMLElement | null; p; p = p.parentElement?.closest(".shp") as HTMLElement | null) {
      if (this.sel.includes(Number(p.dataset.id)) || (this.text && Number(p.dataset.id) === this.text.id)) break;
      top = p;
    }
    shp = top;
    const v = this.shapeView(Number(shp.dataset.id));
    if (!v) return null;
    // A group's member is hit as its group (unless the group is selected).
    return { v, text: t.closest(".shp") === shp };
  }

  private onDown(e: MouseEvent): void {
    if (e.button === 2) return;
    if (e.button !== 0) return;
    const target = e.target as HTMLElement;
    e.preventDefault();
    this.focus();
    if (this.tool) { this.drawTool(e); return; }
    const handle = target.closest(".handle") as HTMLElement | null;
    if (handle?.dataset.handle?.startsWith("crop-")) { this.dragCrop(Number(handle.dataset.id), handle.dataset.handle.slice(5) as Exclude<Handle, "rot">); return; }
    if (handle?.dataset.handle?.startsWith("adj")) { this.dragAdjust(Number(handle.dataset.id), Number(handle.dataset.handle.slice(3))); return; }
    if (handle) { this.dragHandle(e, handle.dataset.handle as Handle, Number(handle.dataset.id)); return; }
    const edge = target.closest(".edge") as HTMLElement | null;
    if (edge) {
      const id = Number((edge.parentElement as HTMLElement).dataset.id);
      this.endText();
      if (!this.sel.includes(id)) this.sel = [id];
      this.render();
      this.host.selectionChanged();
      this.dragMove(e, false);
      return;
    }
    const hit = this.hit(e);
    if (!hit) {
      if (!e.shiftKey && !e.ctrlKey) this.clear();
      this.marquee(e);
      return;
    }
    const { v } = hit;
    if (e.detail === 2) { this.onDouble(e); return; }
    const td = (e.target as HTMLElement).closest("td[data-r]") as HTMLElement | null;
    if (v.frame === "table" && td && e.shiftKey && !e.ctrlKey && this.text?.id === v.id && this.text.cell) {
      // Umschalt+Klick in another cell: the cells between.
      this.setRange({ r: Number(td.dataset.r), c: Number(td.dataset.c) });
      return;
    }
    if (v.frame === "table" && td && !e.shiftKey && !e.ctrlKey) {
      const cell = { r: Number(td.dataset.r), c: Number(td.dataset.c) };
      const same = this.text?.id === v.id && this.text.cell?.r === cell.r && this.text.cell?.c === cell.c;
      if (!same) { this.endText(); this.enterText(v.id, undefined, cell); }
      const pos = this.posAt(v.id, e.clientX, e.clientY);
      if (pos && this.text) { this.text.range = undefined; this.text.caret = pos; this.text.anchor = pos; this.render(); this.dragCells(v.id, cell); }
      return;
    }
    // Clicks in the text being edited move the caret or select.
    if (this.text && this.text.id === v.id && hit.text) {
      const pos = this.posAt(v.id, e.clientX, e.clientY);
      if (pos) {
        if (e.shiftKey) this.text.caret = pos;
        else { this.text.caret = pos; this.text.anchor = pos; }
        this.text.goalX = undefined;
        this.deck.breakMerge();
        if (e.detail === 3) this.selectParagraph();
        else this.render();
        this.host.selectionChanged();
        this.dragText(e);
      }
      return;
    }
    // On the letters of a shape's text the caret goes there at once and a drag selects text; the frame
    // moves only when dragged beside the text or by its outline (catalogue "drag rect plain").
    const onLetters = hit.text && v.kind !== "grpSp" && !!v.text && !v.text.empty;
    if (onLetters && !e.shiftKey && !e.ctrlKey) {
      const pos = this.posAt(v.id, e.clientX, e.clientY);
      if (this.text?.id !== v.id) this.endText();
      this.enterText(v.id, pos ?? undefined);
      this.dragText(e);
      return;
    }
    const multi = e.shiftKey || e.ctrlKey;
    const wasSelected = this.sel.includes(v.id);
    if (!multi && !wasSelected) { this.endText(); this.sel = [v.id]; this.render(); this.host.selectionChanged(); }
    if (multi && !wasSelected) { this.endText(); this.sel = [...this.sel, v.id]; this.render(); this.host.selectionChanged(); }
    const textCapable = !!v.text || v.kind === "sp";
    const enterText = hit.text && textCapable && (v.ph || v.textBox || !!(v.text && !v.text.empty) || v.kind === "sp");
    const start = { x: e.clientX, y: e.clientY };
    let moved = false;
    const move = (m: MouseEvent) => {
      if (moved) return;
      if (Math.hypot(m.clientX - start.x, m.clientY - start.y) < 4) return;
      moved = true;
      window.removeEventListener("mousemove", move);
      window.removeEventListener("mouseup", up);
      if (onLetters) {
        // Umschalt/Strg+drag on letters selects text too (PowerPoint).
        const pos = this.posAt(v.id, start.x, start.y);
        this.sel = [v.id];
        this.enterText(v.id, pos ?? undefined);
        this.dragText(e);
        return;
      }
      this.endText();
      this.render();
      this.dragMove(e, true, m);
    };
    const up = () => {
      window.removeEventListener("mousemove", move);
      window.removeEventListener("mouseup", up);
      if (moved) return;
      if (multi && wasSelected) { this.sel = this.sel.filter((x) => x !== v.id); this.render(); this.host.selectionChanged(); return; }
      if (!multi && enterText && v.kind !== "grpSp") {
        const pos = v.text && !v.text.empty ? this.posAt(v.id, start.x, start.y) : { p: 0, o: 0 };
        this.enterText(v.id, pos ?? undefined);
      }
    };
    window.addEventListener("mousemove", move);
    window.addEventListener("mouseup", up);
  }

  /** Dragging from a table cell: text inside it, a range of cells once the pointer is over another cell. */
  private dragCells(id: number, start: { r: number; c: number }): void {
    const move = (m: MouseEvent) => {
      const t = this.text;
      if (!t || t.id !== id) return;
      const td = (document.elementFromPoint(m.clientX, m.clientY) as HTMLElement | null)?.closest("td[data-r]") as HTMLElement | null;
      const frame = this.rendered?.shapes.get(id);
      if (td && frame?.contains(td) && (Number(td.dataset.r) !== start.r || Number(td.dataset.c) !== start.c)) {
        const before = JSON.stringify(t.range);
        this.setRange({ r: Number(td.dataset.r), c: Number(td.dataset.c) }, false);
        if (JSON.stringify(t.range) !== before) this.render();
        return;
      }
      if (t.range) { t.range = undefined; this.render(); }
      const pos = this.posAt(id, m.clientX, m.clientY);
      if (pos && T.cmpPos(pos, t.caret)) { t.caret = pos; this.render(); }
    };
    const up = () => { window.removeEventListener("mousemove", move); window.removeEventListener("mouseup", up); this.host.selectionChanged(); };
    window.addEventListener("mousemove", move);
    window.addEventListener("mouseup", up);
  }

  /** The table of the frame being edited, if a cell is. */
  tableOf(): Element | undefined {
    const t = this.text;
    if (!t?.cell) return undefined;
    return this.shapeEl(t.id)?.getElementsByTagNameNS(PptxDom.NS_A, "tbl")[0];
  }

  /** The cell range from the cell being edited to `to` (grown to whole merged cells). */
  setRange(to: { r: number; c: number }, redraw = true): void {
    const t = this.text, tbl = this.tableOf();
    if (!t?.cell || !tbl) return;
    const g = TC.grid(tbl);
    const x = TC.normalize(g, t.cell, to);
    t.range = TC.isMulti(g, x) ? x : undefined;
    if (redraw) { this.render(); this.host.selectionChanged(); }
  }

  /** Tabelle auswählen while a cell is being edited: every cell. */
  selectAllCells(): void {
    const tbl = this.tableOf(), t = this.text;
    if (!tbl || !t?.cell) return;
    const g = TC.grid(tbl);
    t.range = { r0: 0, c0: 0, r1: g.length - 1, c1: (g[0]?.length ?? 1) - 1 };
    this.render();
    this.host.selectionChanged();
  }

  /** Spalte / Zeile auswählen: the cells of the column or row of the cell being edited. */
  selectLine(what: "row" | "col"): void {
    const tbl = this.tableOf(), t = this.text;
    if (!tbl || !t?.cell) return;
    const g = TC.grid(tbl);
    const x = what === "row" ? TC.normalize(g, { r: t.cell.r, c: 0 }, { r: t.cell.r, c: g[0].length - 1 }) : TC.normalize(g, { r: 0, c: t.cell.c }, { r: g.length - 1, c: t.cell.c });
    t.range = TC.isMulti(g, x) ? x : undefined;
    this.render();
    this.host.selectionChanged();
  }

  /**
   * Selecting by dragging: by characters inside the first word, by whole words once the pointer leaves it,
   * with the paragraph mark when it goes past the end of a line (PowerPoint's "ganze Wörter markieren";
   * measured: a drag from inside "Form" to beyond it selects "Form").
   */
  private dragText(_e: MouseEvent): void {
    const id = this.text!.id;
    const a0 = this.text!.anchor;
    const move = (m: MouseEvent) => {
      const t = this.text;
      if (!t || t.id !== id) return;
      const pos = this.posAt(id, m.clientX, m.clientY);
      if (!pos) return;
      const body = this.textBody(id);
      const ps = body ? T.paras(body) : [];
      const w0 = ps[a0.p] ? T.wordAt(ps[a0.p], a0.o) : { from: a0.o, to: a0.o };
      const inFirst = T.cmpPos(pos, a0) === 0 || (pos.p === a0.p && pos.o > w0.from && pos.o < w0.to);
      let anchor = a0, caret = pos;
      if (!inFirst && ps.length) {
        const fwd = T.cmpPos(pos, a0) > 0;
        anchor = { p: a0.p, o: fwd ? w0.from : w0.to };
        const pp = ps[pos.p];
        const len = T.paraLength(pp);
        if (fwd && pos.o >= len && pos.p + 1 < ps.length) caret = { p: pos.p + 1, o: 0 };
        else { const w = T.wordAt(pp, pos.o); caret = { p: pos.p, o: fwd ? Math.max(pos.o, w.to) : Math.min(pos.o, w.from) }; }
      }
      if (T.cmpPos(caret, t.caret) || T.cmpPos(anchor, t.anchor)) { t.caret = caret; t.anchor = anchor; this.render(); }
    };
    const up = () => { window.removeEventListener("mousemove", move); window.removeEventListener("mouseup", up); this.host.selectionChanged(); };
    window.addEventListener("mousemove", move);
    window.addEventListener("mouseup", up);
  }

  private onDouble(e: MouseEvent): void {
    const hit = this.hit(e);
    if (!hit) return;
    const v = hit.v;
    const onLetters = hit.text && !!v.text && !v.text.empty;
    if (!onLetters && this.host.doubleClick(v)) return;
    if (!onLetters && v.text && !v.text.empty) {
      // Beside the letters: the caret goes to the nearest place, nothing is selected (measured).
      const pos = this.posAt(v.id, e.clientX, e.clientY) ?? undefined;
      if (this.text?.id !== v.id) { this.endText(); this.enterText(v.id, pos); } else if (pos && this.text) { this.text.caret = pos; this.text.anchor = pos; this.render(); }
      return;
    }
    if (!v.text && v.kind !== "sp") return;
    const body = this.textBody(v.id);
    const pos = this.posAt(v.id, e.clientX, e.clientY) ?? { p: 0, o: 0 };
    if (!this.text || this.text.id !== v.id) this.enterText(v.id, pos);
    if (!body || !this.text) return;
    const p = T.paras(body)[pos.p];
    if (!p) return;
    // The word without the space after it (measured: "Zeile" in "Erste Zeile hier", 5 characters).
    const w = T.wordAt(p, pos.o);
    this.text.anchor = { p: pos.p, o: w.from };
    this.text.caret = { p: pos.p, o: w.to };
    this.render();
    this.host.selectionChanged();
  }

  private selectParagraph(): void {
    if (!this.text) return;
    const body = this.textBody(this.text.id)!;
    const p = this.text.caret.p;
    const ps = T.paras(body);
    // With its paragraph mark, as PowerPoint selects it (17 characters for "Erste Zeile hier").
    this.text.anchor = { p, o: 0 };
    this.text.caret = p + 1 < ps.length ? { p: p + 1, o: 0 } : { p, o: T.paraLength(ps[p]) };
    this.render();
  }

  private onContext(e: MouseEvent): void {
    e.preventDefault();
    this.focus();
    const hit = this.hit(e);
    if (!hit) {
      if (!this.text) { this.sel = []; this.render(); this.host.selectionChanged(); }
      this.host.contextMenu(e, "slide");
      return;
    }
    if (this.text && this.text.id === hit.v.id) {
      // A right-click outside the selected text puts the caret there first.
      const pos = this.posAt(hit.v.id, e.clientX, e.clientY);
      const t = this.text;
      const inSel = pos && T.cmpPos(T.minPos(t.caret, t.anchor), pos) <= 0 && T.cmpPos(pos, T.maxPos(t.caret, t.anchor)) <= 0;
      if (pos && !inSel) { t.caret = pos; t.anchor = pos; this.render(); }
      this.host.contextMenu(e, "text");
      return;
    }
    if (hit.text && hit.v.text && !hit.v.text.empty) {
      // On the letters: the caret goes there (PowerPoint), then the text menu.
      this.endText();
      this.enterText(hit.v.id, this.posAt(hit.v.id, e.clientX, e.clientY) ?? undefined);
      this.host.contextMenu(e, "text");
      return;
    }
    if (!this.sel.includes(hit.v.id)) { this.endText(); this.sel = [hit.v.id]; this.render(); this.host.selectionChanged(); }
    this.host.contextMenu(e, "shape");
  }

  private hover(e: MouseEvent): void {
    if (e.buttons) return;
    const t = e.target as HTMLElement;
    if (this.tool) { this.el.style.cursor = "crosshair"; return; }
    if (t.closest(".handle") || t.closest(".edge")) { this.el.style.cursor = ""; return; }
    const hit = this.hit(e);
    let cursor = "default";
    if (hit) {
      const textish = hit.text && (hit.v.ph || hit.v.textBox || (this.text?.id === hit.v.id));
      cursor = textish ? "text" : "move";
    }
    this.el.style.cursor = cursor;
  }

  /** Auswahlrahmen: shapes completely inside the rectangle are selected. */
  private marquee(e: MouseEvent): void {
    const start = this.toSlide(e.clientX, e.clientY);
    const box = h("div", "marquee");
    let shown = false;
    const move = (m: MouseEvent) => {
      const p = this.toSlide(m.clientX, m.clientY);
      const x = Math.min(start.x, p.x), y = Math.min(start.y, p.y), w = Math.abs(p.x - start.x), hh = Math.abs(p.y - start.y);
      if (!shown && Math.hypot(m.clientX - e.clientX, m.clientY - e.clientY) < 3) return;
      if (!shown) { this.overlay.append(box); shown = true; }
      const s = this.scale / EMU_PX;
      box.style.cssText = `left:${x * s}px;top:${y * s}px;width:${w * s}px;height:${hh * s}px`;
    };
    const up = (m: MouseEvent) => {
      window.removeEventListener("mousemove", move);
      window.removeEventListener("mouseup", up);
      box.remove();
      if (!shown || !this.view) return;
      const p = this.toSlide(m.clientX, m.clientY);
      const x0 = Math.min(start.x, p.x), y0 = Math.min(start.y, p.y), x1 = Math.max(start.x, p.x), y1 = Math.max(start.y, p.y);
      const inside = this.view.shapes.filter((s) => !s.hidden && s.x >= x0 && s.y >= y0 && s.x + s.w <= x1 && s.y + s.h <= y1).map((s) => s.id);
      const ids = e.shiftKey || e.ctrlKey ? [...new Set([...this.sel, ...inside])] : inside;
      this.select(ids);
    };
    window.addEventListener("mousemove", move);
    window.addEventListener("mouseup", up);
  }

  // ----- moving -------------------------------------------------------------------------------------------------

  /** Guide lines for a moving box: snaps to the slide's and other shapes' edges and centres. */
  private snap(box: { x: number; y: number; w: number; h: number }, skip: Set<number>, noSnap: boolean): { dx: number; dy: number; lines: { x?: number; y?: number; from: number; to: number }[] } {
    const lines: { x?: number; y?: number; from: number; to: number }[] = [];
    if (!this.view || noSnap || !this.smartGuides) return { dx: 0, dy: 0, lines };
    const size = this.pageSize;
    const tol = (SNAP_PX / this.scale) * EMU_PX;
    const others = this.view.shapes.filter((s) => !skip.has(s.id) && !s.hidden);
    const xs: { v: number; from: number; to: number }[] = [{ v: 0, from: 0, to: size.cy }, { v: size.cx / 2, from: 0, to: size.cy }, { v: size.cx, from: 0, to: size.cy }];
    const ys: { v: number; from: number; to: number }[] = [{ v: 0, from: 0, to: size.cx }, { v: size.cy / 2, from: 0, to: size.cx }, { v: size.cy, from: 0, to: size.cx }];
    for (const o of others) {
      for (const v of [o.x, o.x + o.w / 2, o.x + o.w]) xs.push({ v, from: Math.min(o.y, box.y), to: Math.max(o.y + o.h, box.y + box.h) });
      for (const v of [o.y, o.y + o.h / 2, o.y + o.h]) ys.push({ v, from: Math.min(o.x, box.x), to: Math.max(o.x + o.w, box.x + box.w) });
    }
    const best = (cands: { v: number; from: number; to: number }[], mine: number[]) => {
      let hit: { d: number; c: { v: number; from: number; to: number } } | null = null;
      for (const c of cands) for (const m of mine) { const d = c.v - m; if (Math.abs(d) <= tol && (!hit || Math.abs(d) < Math.abs(hit.d))) hit = { d, c }; }
      return hit;
    };
    const bx = best(xs, [box.x, box.x + box.w / 2, box.x + box.w]);
    const by = best(ys, [box.y, box.y + box.h / 2, box.y + box.h]);
    if (bx) lines.push({ x: bx.c.v, from: bx.c.from, to: bx.c.to });
    if (by) lines.push({ y: by.c.v, from: by.c.from, to: by.c.to });
    return { dx: bx?.d ?? 0, dy: by?.d ?? 0, lines };
  }

  private drawGuides(lines: { x?: number; y?: number; from: number; to: number }[]): void {
    this.overlay.querySelectorAll(".guide").forEach((g) => g.remove());
    const s = this.scale / EMU_PX;
    for (const l of lines) {
      const g = h("div", "guide");
      if (l.x !== undefined) g.style.cssText = `left:${l.x * s}px;top:${l.from * s}px;width:0;height:${(l.to - l.from) * s}px`;
      else g.style.cssText = `top:${l.y! * s}px;left:${l.from * s}px;height:0;width:${(l.to - l.from) * s}px`;
      this.overlay.append(g);
    }
  }

  /** Drags the selected shapes; Umschalt keeps one direction, Strg copies, Alt switches snapping off. */
  private dragMove(e: MouseEvent, started: boolean, first?: MouseEvent): void {
    const views = this.selectedViews;
    if (!views.length) return;
    const start = this.toSlide(e.clientX, e.clientY);
    const box = union(views);
    const skip = new Set(views.map((v) => v.id));
    let dx = 0, dy = 0, copy = false;
    const els = views.map((v) => this.rendered?.shapes.get(v.id)).filter((x): x is HTMLElement => !!x);
    const frames = Array.from(this.overlay.querySelectorAll(".frame")) as HTMLElement[];
    const move = (m: MouseEvent) => {
      const p = this.toSlide(m.clientX, m.clientY);
      dx = p.x - start.x; dy = p.y - start.y;
      if (m.shiftKey) { if (Math.abs(dx) > Math.abs(dy)) dy = 0; else dx = 0; }
      const snapped = this.snap({ x: box.x + dx, y: box.y + dy, w: box.w, h: box.h }, skip, m.altKey);
      if (!m.shiftKey || dy === 0) dx += snapped.dx;
      if (!m.shiftKey || dx === 0) dy += snapped.dy;
      if (this.snapToGrid && !m.altKey && !snapped.lines.length) {
        dx = Math.round((box.x + dx) / this.gridSpacing) * this.gridSpacing - box.x;
        dy = Math.round((box.y + dy) / this.gridSpacing) * this.gridSpacing - box.y;
      }
      this.drawGuides(snapped.lines);
      copy = m.ctrlKey;
      const tx = (dx / EMU_PX), ty = (dy / EMU_PX);
      for (const el of els) { el.style.translate = `${tx}px ${ty}px`; el.style.opacity = copy ? "0.6" : ""; }
      for (const f of frames) f.style.translate = `${tx * this.scale}px ${ty * this.scale}px`;
      this.el.style.cursor = copy ? "copy" : "move";
    };
    const up = () => {
      window.removeEventListener("mousemove", move);
      window.removeEventListener("mouseup", up);
      this.drawGuides([]);
      this.el.style.cursor = "";
      if (!dx && !dy) { this.render(); return; }
      if (copy) this.copyMoved(views, dx, dy);
      else this.edit("Verschieben", () => { for (const v of views) this.moveShape(v, dx, dy); });
    };
    window.addEventListener("mousemove", move);
    window.addEventListener("mouseup", up);
    if (started && first) move(first);
  }

  /** Moves one shape (also a group member, through the group's child coordinates). */
  moveShape(v: ShapeView, dx: number, dy: number): void {
    const el = this.shapeEl(v.id);
    if (!el) return;
    const g = this.groupOf(v.id);
    const k = g ? this.groupScale(g) : { sx: 1, sy: 1 };
    const own = this.ownXfrm(el, v);
    S.setXfrm(el, { x: own.x + dx / k.sx, y: own.y + dy / k.sy, w: own.w, h: own.h });
  }

  /** The shape's xfrm in its own coordinates (for placeholders: the inherited one). */
  private ownXfrm(el: Element, v: ShapeView): S.Xfrm {
    const x = el.localName === "graphicFrame" ? PptxDom.kid(el, "xfrm") : PptxDom.at(el, `${el.localName === "grpSp" ? "grpSpPr" : "spPr"}/xfrm`);
    const off = PptxDom.kid(x, "off"), ext = PptxDom.kid(x, "ext");
    if (off && ext) return { x: PptxDom.num(off, "x") ?? 0, y: PptxDom.num(off, "y") ?? 0, w: PptxDom.num(ext, "cx") ?? 0, h: PptxDom.num(ext, "cy") ?? 0, rot: v.rot, flipH: v.flipH, flipV: v.flipV };
    return { x: v.x, y: v.y, w: v.w, h: v.h, rot: v.rot, flipH: v.flipH, flipV: v.flipV };
  }

  private groupOf(id: number): Element | null {
    const el = this.shapeEl(id);
    const p = el?.parentElement;
    return p && p.localName === "grpSp" ? p : null;
  }

  private groupScale(g: Element): { sx: number; sy: number } {
    const x = PptxDom.at(g, "grpSpPr/xfrm");
    const ext = PptxDom.kid(x, "ext"), ce = PptxDom.kid(x, "chExt");
    const sx = (PptxDom.num(ext, "cx") ?? 1) / (PptxDom.num(ce, "cx") || 1), sy = (PptxDom.num(ext, "cy") ?? 1) / (PptxDom.num(ce, "cy") || 1);
    return { sx: sx || 1, sy: sy || 1 };
  }

  private copyMoved(views: ShapeView[], dx: number, dy: number): void {
    const ids: number[] = [];
    this.edit("Kopieren", () => {
      const tree = S.treeOf(this.doc!);
      let next = S.maxId(tree);
      for (const v of views) {
        const el = this.shapeEl(v.id);
        if (!el) continue;
        const copy = el.cloneNode(true) as Element;
        S.renumber(copy, () => ++next);
        tree.append(copy);
        ids.push(S.shapeId(copy));
        const own = this.ownXfrm(el, v);
        S.setXfrm(copy, { x: own.x + dx, y: own.y + dy, w: own.w, h: own.h });
      }
    });
    this.sel = ids;
    this.render();
    this.host.selectionChanged();
  }

  // ----- sizing and turning -------------------------------------------------------------------------------------

  private dragHandle(_e: MouseEvent, k: Handle, id: number): void {
    const v = this.shapeView(id);
    if (!v) return;
    this.endText();
    const cx = v.x + v.w / 2, cy = v.y + v.h / 2;
    const el = this.rendered?.shapes.get(id);
    const frame = this.overlay.querySelector(`.frame[data-id="${id}"]`) as HTMLElement | null;
    const keepRatio = v.kind === "pic" || (v.el.localName === "graphicFrame" && v.frame === "chart" ? false : !!PptxDom.at(v.el, "nvPicPr/cNvPicPr/picLocks")?.getAttribute("noChangeAspect"));
    let result: S.Xfrm | null = null;
    // The handle follows the pointer's movement from where it was grabbed (not the pointer's place: the
    // handle is drawn around the edge; measured, PowerPoint sizes by the full movement).
    const grab = this.toSlide(_e.clientX, _e.clientY);
    const [hfx, hfy] = k === "rot" ? [0.5, 0] : HANDLE_POS[k];
    const ang0 = (v.rot * Math.PI) / 180, hlx = (hfx - 0.5) * v.w, hly = (hfy - 0.5) * v.h;
    const off = k === "rot" ? { x: 0, y: 0 } : { x: cx + hlx * Math.cos(ang0) - hly * Math.sin(ang0) - grab.x, y: cy + hlx * Math.sin(ang0) + hly * Math.cos(ang0) - grab.y };
    // A text box that takes its text's height (spAutoFit) keeps it: n and s change nothing (measured).
    const autoHeight = v.text?.body.autofit === "shape";
    if (autoHeight && (k === "n" || k === "s")) {
      const done = () => { window.removeEventListener("mouseup", done); };
      window.addEventListener("mouseup", done);
      return;
    }
    const move = (m: MouseEvent) => {
      const p0 = this.toSlide(m.clientX, m.clientY);
      const p = { x: p0.x + off.x, y: p0.y + off.y };
      if (k === "rot") {
        let ang = (Math.atan2(p.y - cy, p.x - cx) * 180) / Math.PI + 90;
        if (m.shiftKey) ang = Math.round(ang / 15) * 15;
        ang = ((ang % 360) + 360) % 360;
        result = { x: v.x, y: v.y, w: v.w, h: v.h, rot: ang, flipH: v.flipH, flipV: v.flipV };
      } else result = resize(v, k, p, { shift: m.shiftKey, ctrl: m.ctrlKey, keepRatio: keepRatio && !m.shiftKey ? true : m.shiftKey && (k.length === 2) });
      const r = result;
      // Preview: the shape's element and frame take the new box.
      const s = this.scale;
      if (el) { el.style.left = `${r.x / EMU_PX}px`; el.style.top = `${r.y / EMU_PX}px`; el.style.width = `${r.w / EMU_PX}px`; el.style.height = `${r.h / EMU_PX}px`; el.style.transform = `rotate(${r.rot}deg)${r.flipH ? " scaleX(-1)" : ""}${r.flipV ? " scaleY(-1)" : ""}`; el.style.opacity = "0.75"; }
      if (frame) frame.style.cssText = `left:${(r.x / EMU_PX) * s}px;top:${(r.y / EMU_PX) * s}px;width:${(r.w / EMU_PX) * s}px;height:${(r.h / EMU_PX) * s}px;transform:rotate(${r.rot}deg)`;
    };
    const up = () => {
      window.removeEventListener("mousemove", move);
      window.removeEventListener("mouseup", up);
      if (!result) { this.render(); return; }
      const r: S.Xfrm = result;
      this.edit(k === "rot" ? "Drehen" : "Größe ändern", () => {
        const sh = this.shapeEl(id);
        if (!sh) return;
        const g = this.groupOf(id);
        if (g) {
          const gv = this.groupView(g);
          const sc = this.groupScale(g);
          const ch = S.groupChildFrame(g)!;
          S.setXfrm(sh, { x: ch.chOff[0] + (r.x - gv.x) / sc.sx, y: ch.chOff[1] + (r.y - gv.y) / sc.sy, w: r.w / sc.sx, h: r.h / sc.sy, rot: r.rot - gv.rot, flipH: r.flipH, flipV: r.flipV });
        } else {
          S.setXfrm(sh, r);
          if (sh.localName === "grpSp") this.scaleGroup(sh, v, r);
        }
        // A sized text box no longer grows sideways with its text: it wraps (PowerPoint).
        const bp = S.bodyPrOf(sh);
        if (bp && k !== "rot" && bp.getAttribute("wrap") === "none" && (k === "e" || k === "w" || k.length === 2)) bp.setAttribute("wrap", "square");
        // Its height follows its text again, the top staying where it was dragged.
        if (autoHeight && k !== "rot") this.updateAutofit(id);
      });
    };
    window.addEventListener("mousemove", move);
    window.addEventListener("mouseup", up);
  }

  /**
   * A yellow handle: the adjust value is searched so that the handle lies where the pointer is, through
   * the evaluated geometry (every preset follows its own formulas, as in RaabeWrite and RaabeCalc).
   */
  private dragAdjust(id: number, i: number): void {
    const v = this.shapeView(id);
    if (!v) return;
    this.endText();
    const w = v.w / EMU_PX, hh = v.h / EMU_PX;
    const base = { ...presetAdjust(v.geometry), ...(v.info.adjust ?? {}) };
    const hd = presetGeometry(v.geometry, w, hh, base)?.handles[i];
    if (!hd) return;
    const at = (a: Record<string, number>) => presetGeometry(v.geometry, w, hh, a)?.handles[i];
    const search = (adj: Record<string, number>, key: string, lo: number, hi: number, pos: (a: Record<string, number>) => number, target: number) => {
      let best = adj[key] ?? lo, bestD = Infinity;
      let a = Math.min(lo, hi), c = Math.max(lo, hi);
      for (let round = 0; round < 3; round++) {
        for (let k = 0; k <= 48; k++) {
          const val = a + ((c - a) * k) / 48;
          const d = Math.abs(pos({ ...adj, [key]: val }) - target);
          if (d < bestD) { bestD = d; best = val; }
        }
        const span = (c - a) / 48;
        a = Math.max(Math.min(lo, hi), best - span);
        c = Math.min(Math.max(lo, hi), best + span);
      }
      return Math.round(best);
    };
    let adj = base;
    const move = (ev: MouseEvent) => {
      const p = this.toSlide(ev.clientX, ev.clientY);
      const cx = v.x + v.w / 2, cy = v.y + v.h / 2, rad = (-v.rot * Math.PI) / 180;
      const dx = (p.x - cx) / EMU_PX, dy = (p.y - cy) / EMU_PX;
      let qx = dx * Math.cos(rad) - dy * Math.sin(rad) + w / 2, qy = dx * Math.sin(rad) + dy * Math.cos(rad) + hh / 2;
      if (v.flipH) qx = w - qx;
      if (v.flipV) qy = hh - qy;
      const next: Record<string, number> = { ...base };
      if (hd.kind === "polar") {
        if (hd.gdY) {
          let ang = (Math.atan2(qy - hh / 2, qx - w / 2) * 180) / Math.PI;
          if (ang < 0) ang += 360;
          next[hd.gdY] = Math.max(hd.minY ?? 0, Math.min(hd.maxY ?? 21599999, Math.round(ang * 60000)));
        }
        if (hd.gdX) next[hd.gdX] = search(next, hd.gdX, hd.minX ?? 0, hd.maxX ?? 50000, (a) => { const q = at(a); return q ? Math.hypot(q.x - w / 2, q.y - hh / 2) : 0; }, Math.hypot(qx - w / 2, qy - hh / 2));
      } else {
        const big = (x: number | undefined, fallback: number) => (x === undefined || Math.abs(x) > 1e8 ? fallback : x);
        if (hd.gdX) next[hd.gdX] = search(next, hd.gdX, big(hd.minX, -400000), big(hd.maxX, 400000), (a) => at(a)?.x ?? 0, qx);
        if (hd.gdY) next[hd.gdY] = search(next, hd.gdY, big(hd.minY, -400000), big(hd.maxY, 400000), (a) => at(a)?.y ?? 0, qy);
      }
      adj = next;
      // Preview: the handle follows; the shape is drawn on release.
      const q = at(adj);
      const hEl = this.overlay.querySelector(`.handle[data-id="${id}"][data-handle="adj${i}"]`) as HTMLElement | null;
      if (q && hEl) { hEl.style.left = `${((v.flipH ? w - q.x : q.x) / w) * 100}%`; hEl.style.top = `${((v.flipV ? hh - q.y : q.y) / hh) * 100}%`; }
    };
    const up = () => {
      window.removeEventListener("mousemove", move);
      window.removeEventListener("mouseup", up);
      if (JSON.stringify(adj) === JSON.stringify(base)) { this.render(); return; }
      this.edit("Form anpassen", () => {
        const el = this.shapeEl(id);
        const prst = el && PptxDom.at(el, "spPr/prstGeom");
        if (!prst) return;
        let av = PptxDom.kid(prst, "avLst");
        if (!av) { av = PptxDom.make(el.ownerDocument, "a:avLst"); prst.append(av); }
        while (av.firstChild) av.firstChild.remove();
        // PowerPoint writes the values it was given, also the defaults of the others it moved.
        for (const [k, val] of Object.entries(adj)) if (k in (v.info.adjust ?? {}) || val !== presetAdjust(v.geometry)[k]) av.append(PptxDom.make(el.ownerDocument, "a:gd", { name: k, fmla: `val ${val}` }));
      });
    };
    window.addEventListener("mousemove", move);
    window.addEventListener("mouseup", up);
  }

  /** Starts or ends Zuschneiden for the selected picture. */
  toggleCrop(): void {
    const v = this.selectedViews[0];
    this.crop = this.crop === null && v?.kind === "pic" ? v.id : null;
    this.render();
    this.host.selectionChanged();
  }

  /** A crop handle: the edge moves over the picture, which stays where it is (a:srcRect and the frame change). */
  private dragCrop(id: number, k: Exclude<Handle, "rot">): void {
    const v = this.shapeView(id);
    if (!v?.pic) return;
    const c0 = v.pic.crop;
    const fw = v.w / Math.max(0.01, 1 - c0.l - c0.r), fh = v.h / Math.max(0.01, 1 - c0.t - c0.b);
    const start = { x: v.x, y: v.y, r: v.x + v.w, b: v.y + v.h };
    let box = { ...start };
    const move = (ev: MouseEvent) => {
      const p = this.toSlide(ev.clientX, ev.clientY);
      box = { ...start };
      if (k.includes("w")) box.x = Math.min(p.x, start.r - 1);
      if (k.includes("e")) box.r = Math.max(p.x, start.x + 1);
      if (k.includes("n")) box.y = Math.min(p.y, start.b - 1);
      if (k.includes("s")) box.b = Math.max(p.y, start.y + 1);
      const fr = this.overlay.querySelector(`.frame[data-id="${id}"]`) as HTMLElement | null;
      const s = this.scale / EMU_PX;
      if (fr) { fr.style.left = `${box.x * s}px`; fr.style.top = `${box.y * s}px`; fr.style.width = `${(box.r - box.x) * s}px`; fr.style.height = `${(box.b - box.y) * s}px`; }
    };
    const up = () => {
      window.removeEventListener("mousemove", move);
      window.removeEventListener("mouseup", up);
      const l = c0.l + (box.x - start.x) / fw, t = c0.t + (box.y - start.y) / fh, r = c0.r + (start.r - box.r) / fw, b = c0.b + (start.b - box.b) / fh;
      this.edit("Zuschneiden", () => {
        const el = this.shapeEl(id)!;
        const bf = PptxDom.kid(el, "blipFill")!;
        let src = PptxDom.kid(bf, "srcRect");
        if (!src) { src = PptxDom.make(el.ownerDocument, "a:srcRect"); bf.insertBefore(src, PptxDom.kid(bf, "stretch") ?? PptxDom.kid(bf, "tile") ?? null); }
        const set = (n: string, x: number) => { const val = Math.round(x * 100000); if (val) src!.setAttribute(n, String(val)); else src!.removeAttribute(n); };
        set("l", l); set("t", t); set("r", r); set("b", b);
        S.setXfrm(el, { x: box.x, y: box.y, w: box.r - box.x, h: box.b - box.y });
      });
    };
    window.addEventListener("mousemove", move);
    window.addEventListener("mouseup", up);
  }

  private groupView(g: Element): ShapeView {
    return this.shapeView(S.shapeId(g))!;
  }

  /** A sized group keeps its members' child coordinates (chOff/chExt), so they scale with it. */
  private scaleGroup(_g: Element, _before: ShapeView, _after: S.Xfrm): void {
    // Nothing to do: chOff/chExt stay, the members' coordinates map onto the new box.
  }

  // ----- drawing new shapes ------------------------------------------------------------------------------------

  private drawTool(e: MouseEvent): void {
    const tool = this.tool!;
    if (tool.kind === "ink") { this.drawInk(e, tool); return; }
    if (tool.kind === "eraser") {
      // Radierer: a click (or a drag) over an ink stroke removes it.
      const erase = (m: MouseEvent) => {
        const t = document.elementFromPoint(m.clientX, m.clientY) as HTMLElement | null;
        const id = Number(t?.closest(".shp")?.getAttribute("data-id"));
        const v = id ? this.shapeView(id) : undefined;
        if (v && v.name.startsWith("Freihand")) this.edit("Radierer", () => this.shapeEl(id)?.remove());
      };
      erase(e);
      const up = () => { window.removeEventListener("mousemove", erase); window.removeEventListener("mouseup", up); };
      window.addEventListener("mousemove", erase);
      window.addEventListener("mouseup", up);
      return;
    }
    const start = this.toSlide(e.clientX, e.clientY);
    const ghost = h("div", "ghost");
    this.overlay.append(ghost);
    let end = start, dragged = false, shift = false;
    const move = (m: MouseEvent) => {
      end = this.toSlide(m.clientX, m.clientY);
      shift = m.shiftKey;
      if (Math.hypot(m.clientX - e.clientX, m.clientY - e.clientY) > 3) dragged = true;
      const b = drawBox(start, end, shift, tool.kind === "shape" && S.isLine(tool.prst));
      const s = this.scale / EMU_PX;
      ghost.style.cssText = `left:${b.x * s}px;top:${b.y * s}px;width:${b.w * s}px;height:${b.h * s}px`;
    };
    const up = () => {
      window.removeEventListener("mousemove", move);
      window.removeEventListener("mouseup", up);
      ghost.remove();
      this.tool = null;
      this.host.toolDone();
      const line = tool.kind === "shape" && S.isLine(tool.prst);
      let b = drawBox(start, end, shift, line);
      // A click: PowerPoint's default size (2,54 cm square for shapes; a growing text box).
      if (!dragged) b = { x: start.x, y: start.y, w: 914400, h: tool.kind === "textbox" ? 369332 : 914400, rot: 0, flipH: false, flipV: false };
      let id = 0;
      this.edit(tool.kind === "textbox" ? "Textfeld" : "Form einfügen", () => {
        const tree = S.treeOf(this.doc!);
        id = S.maxId(tree) + 1;
        const el = tool.kind === "textbox"
          ? S.newTextBox(this.doc!, id, b.x, b.y, dragged ? b.w : null, 369332)
          : S.newShape(this.doc!, id, tool.prst, { ...b, rot: 0 });
        if (tool.kind === "shape" && (tool.head || tool.tail)) {
          const sp = PptxDom.kid(el, "spPr")!;
          const ln = PptxDom.make(el.ownerDocument, "a:ln");
          if (tool.head) ln.append(PptxDom.make(el.ownerDocument, "a:headEnd", { type: "triangle" }));
          if (tool.tail) ln.append(PptxDom.make(el.ownerDocument, "a:tailEnd", { type: "triangle" }));
          sp.append(ln);
        }
        if (tool.kind === "textbox" && tool.vertical) S.bodyPrOf(el)!.setAttribute("vert", "eaVert");
        tree.append(el);
      });
      if (tool.kind === "textbox") this.enterText(id, { p: 0, o: 0 });
      else this.select([id]);
    };
    window.addEventListener("mousemove", move);
    window.addEventListener("mouseup", up);
  }

  /** Zeichnen: a stroke of the pen or highlighter becomes a freeform shape (custom geometry, no fill). */
  private drawInk(e: MouseEvent, tool: { color: string; width: number; alpha?: number }): void {
    const pts: { x: number; y: number }[] = [this.toSlide(e.clientX, e.clientY)];
    const svgNs = "http://www.w3.org/2000/svg";
    const svg = document.createElementNS(svgNs, "svg");
    svg.setAttribute("class", "ink-preview");
    const path = document.createElementNS(svgNs, "path");
    path.setAttribute("fill", "none");
    path.setAttribute("stroke", `#${tool.color}`);
    path.setAttribute("stroke-opacity", String(1 - (tool.alpha ?? 0)));
    path.setAttribute("stroke-width", String((tool.width * 12700 / EMU_PX) * this.scale));
    path.setAttribute("stroke-linecap", "round");
    path.setAttribute("stroke-linejoin", "round");
    svg.append(path);
    this.overlay.append(svg);
    const s = this.scale / EMU_PX;
    const draw = () => path.setAttribute("d", pts.map((p, i) => `${i ? "L" : "M"}${(p.x * s).toFixed(1)} ${(p.y * s).toFixed(1)}`).join(" "));
    const move = (m: MouseEvent) => { pts.push(this.toSlide(m.clientX, m.clientY)); draw(); };
    const up = () => {
      window.removeEventListener("mousemove", move);
      window.removeEventListener("mouseup", up);
      svg.remove();
      if (pts.length < 2) pts.push({ x: pts[0].x + 1, y: pts[0].y + 1 });
      const x0 = Math.min(...pts.map((p) => p.x)), y0 = Math.min(...pts.map((p) => p.y));
      const w = Math.max(1, Math.max(...pts.map((p) => p.x)) - x0), hh = Math.max(1, Math.max(...pts.map((p) => p.y)) - y0);
      const r = (v: number) => Math.round(v);
      const body = pts.map((p, i) => `<a:${i ? "lnTo" : "moveTo"}><a:pt x="${r(p.x - x0)}" y="${r(p.y - y0)}"/></a:${i ? "lnTo" : "moveTo"}>`).join("");
      const alpha = tool.alpha ? `<a:alpha val="${Math.round((1 - tool.alpha) * 100000)}"/>` : "";
      this.edit("Freihand", () => {
        const doc = this.pptx.doc(this.path!);
        const tree = S.treeOf(doc);
        const id = S.maxId(tree) + 1;
        tree.append(PptxDom.fragment(doc, `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="Freihand ${id - 1}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="${r(x0)}" y="${r(y0)}"/><a:ext cx="${r(w)}" cy="${r(hh)}"/></a:xfrm>` +
          `<a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="l" t="t" r="r" b="b"/><a:pathLst><a:path w="${r(w)}" h="${r(hh)}">${body}</a:path></a:pathLst></a:custGeom><a:noFill/>` +
          `<a:ln w="${Math.round(tool.width * 12700)}" cap="rnd"><a:solidFill><a:srgbClr val="${tool.color}">${alpha}</a:srgbClr></a:solidFill><a:round/></a:ln></p:spPr></p:sp>`)[0]);
      });
    };
    window.addEventListener("mousemove", move);
    window.addEventListener("mouseup", up);
  }

  // ----- keys ---------------------------------------------------------------------------------------------

  private flushInput(): void {
    const v = this.input.value;
    this.input.value = "";
    if (!v) return;
    this.typeText(v);
  }

  /** Typed characters: into the text being edited; with a frame selected at the end of its text (measured: "Form" + "xy" gives "Formxy"). */
  typeText(v: string): void {
    const text = v.replace(/\r?\n/g, "\n");
    if (!this.text) {
      const id = this.sel[0];
      const el = id !== undefined ? this.shapeEl(id) : undefined;
      if (!el || this.sel.length !== 1 || (el.localName !== "sp")) return;
      if (!S.txBodyOf(el)) this.edit("Eingeben", () => { S.ensureTxBody(el); }, `type:${id}`);
      const end = T.endPos(S.txBodyOf(el)!);
      this.text = { id, caret: end, anchor: end, pending: null };
    }
    const t = this.text!;
    const body = this.textBody(t.id);
    if (!body) return;
    this.edit("Eingeben", () => {
      let pos = T.deleteRange(body, t.anchor, t.caret);
      const parts = text.split("\n");
      parts.forEach((part, i) => {
        if (i > 0) pos = T.splitPara(body, pos);
        if (part) pos = T.insertText(body, pos, part, t.pending);
      });
      t.caret = pos; t.anchor = pos; t.goalX = undefined;
    }, `type:${t.id}`);
    t.pending = null;
    if (text.length === 1 && this.autoCorrectOn) this.autoCorrect(text);
  }

  /** AutoKorrektur (Office's defaults, shared with RaabeWrite: D-042, D-072). */
  autoCorrectOn = true;

  /**
   * After a typed character: the replacement list's symbols ((c), -->, ...) at once, the word just ended
   * (sentence capitals, ZWei GRoßbuchstaben, weekdays, dashes) at a word end, typographic quotes for the
   * text's language. Each correction is its own undo step, so Strg+Z brings back what was typed.
   */
  private autoCorrect(ch: string, typed = true): void {
    const t = this.text;
    const body = t ? this.textBody(t.id) : undefined;
    if (!t || !body || T.cmpPos(t.caret, t.anchor) !== 0) return;
    const para = T.paras(body)[t.caret.p];
    if (!para) return;
    const ptxt = T.paraText(para), at = t.caret.o;
    const lang = para.getElementsByTagNameNS(PptxDom.NS_A, "rPr")[0]?.getAttribute("lang") ?? "de-DE";
    let fixes: AutoCorrect.Fix[] = [];
    if (!typed) fixes = AutoCorrect.wordFixes(ptxt, at, lang, { cellStart: !!t.cell });
    else if (AutoCorrect.INSTANT_TRIGGERS.has(ch)) {
      const f = AutoCorrect.instantFix(ptxt, at);
      // A symbol (the arrows) is its character of the symbol font's private use area, as PowerPoint writes it.
      if (f) fixes = [f.symbol ? { a: f.a, b: f.b, text: String.fromCharCode(f.symbol.char), symbol: f.symbol } : { a: f.a, b: f.b, text: f.text }];
    }
    if (typed && !fixes.length && AutoCorrect.TRIGGERS.has(ch)) {
      fixes = AutoCorrect.wordFixes(ptxt, at - 1, lang, { cellStart: !!t.cell });
      if (ch === "\"" || ch === "'") { const q = AutoCorrect.smartQuote(ptxt, at, lang); if (q) fixes.push(q); }
    }
    if (!fixes.length) return;
    this.deck.breakMerge();
    this.edit("AutoKorrektur", () => {
      let caret = t.caret.o;
      for (const f of [...fixes].sort((x, y) => y.a - x.a)) {
        T.deleteRange(body, { p: t.caret.p, o: f.a }, { p: t.caret.p, o: f.b });
        if (f.text) T.insertText(body, { p: t.caret.p, o: f.a }, f.text);
        const sym = f.symbol;
        if (sym) T.formatRange(body, { p: t.caret.p, o: f.a }, { p: t.caret.p, o: f.a + f.text.length }, (r) => {
          PptxDom.kid(r, "sym")?.remove();
          const el = PptxDom.make(r.ownerDocument, "a:sym", { typeface: sym.font, pitchFamily: 2, charset: 2 });
          const before = Array.from(r.children).find((c) => ["hlinkClick", "hlinkMouseOver", "rtl", "extLst"].includes(c.localName)) ?? null;
          r.insertBefore(el, before);
        });
        if (caret >= f.b) caret += f.text.length - (f.b - f.a);
      }
      t.caret = { p: t.caret.p, o: caret }; t.anchor = t.caret;
    });
    this.deck.breakMerge();
  }

  private onKey(e: KeyboardEvent): void {
    if (this.composing || e.isComposing) return;
    // A key an open menu already took (Esc closing the context menu) does nothing on the slide.
    if (e.defaultPrevented) return;
    // Text shortcuts on an empty shape start editing it (PowerPoint: Strg+E on a selected ellipse).
    const lk = e.key.toLowerCase();
    if (!this.text && e.ctrlKey && !e.altKey && this.sel.length === 1 && ((!e.shiftKey && ["e", "l", "r", "j"].includes(lk)) || (e.shiftKey && lk === "f"))) {
      const v = this.selectedViews[0];
      if (v && v.kind === "sp" && !v.ph && (!v.text || v.text.empty)) this.enterText(v.id, { p: 0, o: 0 });
    }
    if (this.host.shortcut(e)) { e.preventDefault(); return; }
    const handled = this.text ? this.textKey(e) : this.frameKey(e);
    if (handled) e.preventDefault();
  }

  private frameKey(e: KeyboardEvent): boolean {
    const k = e.key, ctrl = e.ctrlKey;
    const views = this.selectedViews;
    if (k === "Escape") {
      if (this.crop !== null) { this.crop = null; this.render(); return true; }
      if (this.tool) { this.tool = null; this.host.toolDone(); return true; }
      if (this.sel.length) { this.clear(); return true; }
      return false;
    }
    if (k === "Tab" && !ctrl) { this.cycle(e.shiftKey ? -1 : 1); return true; }
    if (!views.length) {
      if ((k === "PageDown" || k === "PageUp" || k === "Home" || k === "End") && !ctrl && this.host.goSlide) { this.host.goSlide(k); return true; }
      return false;
    }
    if (k.startsWith("Arrow")) {
      if (e.altKey) {
        // Alt+Pfeil turns by 15 degrees (Strg: 1); Alt+Umschalt+Pfeil does nothing on a frame (measured).
        if (!e.shiftKey) this.rotateBy(k === "ArrowLeft" ? -15 : k === "ArrowRight" ? 15 : 0, ctrl);
        return true;
      }
      if (e.shiftKey) {
        // Umschalt+Pfeil sizes about the centre: 10 % per press, Strg+Umschalt 1 % (measured: 200 pt to 220,
        // three presses 360 to 479.2). Rechts/Oben larger, Links/Unten smaller.
        const f = 1 + (ctrl ? 0.01 : 0.1);
        const sx = k === "ArrowRight" ? f : k === "ArrowLeft" ? 1 / f : 1, sy = k === "ArrowUp" ? f : k === "ArrowDown" ? 1 / f : 1;
        this.edit("Größe ändern", () => { for (const v of views) this.scaleShape(v, sx, sy); }, "nudge-size");
        return true;
      }
      // Arrow keys move by one pixel at the zoom, by the grid spacing when Am Raster ausrichten is on
      // (measured: 0.8 pt at the catalogue's zoom, as Strg+Pfeil).
      const step = ctrl || !this.snapToGrid ? (1 / this.scale) * EMU_PX : this.gridSpacing;
      const dx = k === "ArrowLeft" ? -step : k === "ArrowRight" ? step : 0, dy = k === "ArrowUp" ? -step : k === "ArrowDown" ? step : 0;
      this.edit("Verschieben", () => { for (const v of views) this.moveShape(v, dx, dy); }, "nudge");
      return true;
    }
    if (k === "Delete" && ctrl) return true;
    if (k === "Backspace" && ctrl) {
      // Strg+Rücktaste on a frame: editing at the end of its text (measured).
      if (views.length === 1 && (views[0].text || views[0].kind === "sp")) this.enterText(views[0].id);
      return true;
    }
    if (k === "Delete" || k === "Backspace") { this.deleteSelected(); return true; }
    if (k === "Enter" && e.shiftKey) return true;
    if ((k === "F2" || k === "Enter") && views.length === 1 && (views[0].text || views[0].kind === "sp")) { this.enterText(views[0].id, "all"); return true; }
    return false;
  }

  /** Sizes a shape about its centre (Umschalt+Pfeil); a text box keeps the new height (measured: 29.1 to 32 pt). */
  private scaleShape(v: ShapeView, sx: number, sy: number): void {
    const el = this.shapeEl(v.id);
    if (!el) return;
    const own = this.ownXfrm(el, v);
    const w = Math.round(own.w * sx), hh = Math.round(own.h * sy);
    S.setXfrm(el, { ...own, x: Math.round(own.x - (w - own.w) / 2), y: Math.round(own.y - (hh - own.h) / 2), w, h: hh });
  }

  private rotateBy(deg: number, fine: boolean): void {
    const d = fine ? Math.sign(deg) : deg;
    if (!d) return;
    this.edit("Drehen", () => {
      for (const v of this.selectedViews) {
        const el = this.shapeEl(v.id)!;
        const own = this.ownXfrm(el, v);
        S.setXfrm(el, { ...own, rot: v.rot + d });
      }
    }, "rotate");
  }

  /** Tab / Umschalt+Tab: the next / previous shape in z-order. */
  private cycle(dir: 1 | -1): void {
    if (!this.view) return;
    const list = this.view.shapes.filter((s) => !s.hidden).map((s) => s.id);
    if (!list.length) return;
    const cur = this.sel.length ? list.indexOf(this.sel[this.sel.length - 1]) : -1;
    const next = cur < 0 ? (dir > 0 ? 0 : list.length - 1) : (cur + dir + list.length) % list.length;
    this.select([list[next]]);
  }

  /** Entf on frames: removes them; a placeholder with text loses its text first (PowerPoint). */
  deleteSelected(): void {
    const views = this.selectedViews;
    if (!views.length) return;
    this.edit("Löschen", () => {
      for (const v of views) {
        const el = this.shapeEl(v.id);
        if (!el) continue;
        if (v.ph && v.text && !v.text.empty) {
          // The placeholder stays, empty, and comes to the top of the order (as PowerPoint re-creates it).
          T.replaceAll(S.txBodyOf(el)!, "");
          el.parentElement?.append(el);
          continue;
        }
        removeShapeEl(el);
      }
    });
    this.sel = [];
    this.render();
    this.host.selectionChanged();
  }

  private textKey(e: KeyboardEvent): boolean {
    const t = this.text!;
    const body = this.textBody(t.id);
    if (!body) return false;
    const k = e.key, ctrl = e.ctrlKey && !e.altKey, shift = e.shiftKey;
    const collapse = (pos: Pos) => { t.caret = pos; if (!shift) t.anchor = pos; this.deck.breakMerge(); this.render(); this.host.selectionChanged(); };
    const hasSel = T.cmpPos(t.caret, t.anchor) !== 0;
    const ps = T.paras(body);
    const DIRS: Record<string, [number, number]> = { ArrowUp: [-1, 0], ArrowDown: [1, 0], ArrowLeft: [0, -1], ArrowRight: [0, 1] };
    // Alt+Umschalt+Pfeil: Höher/Tiefer stufen, the caret stays (measured: no selection).
    if (e.altKey && shift && (k === "ArrowRight" || k === "ArrowLeft")) { this.changeLevel(k === "ArrowRight" ? 1 : -1); return true; }
    if (t.range && t.cell) {
      const tbl = this.tableOf()!;
      if (k === "Delete" || k === "Backspace") { const x = t.range; this.edit("Löschen", () => TC.clear(tbl, x)); return true; }
      if (shift && DIRS[k]) {
        // Umschalt+Pfeil grows or shrinks the range from its far corner.
        const x = t.range, far = { r: x.r0 === t.cell.r ? x.r1 : x.r0, c: x.c0 === t.cell.c ? x.c1 : x.c0 };
        const g = TC.grid(tbl), d = DIRS[k];
        this.setRange({ r: Math.max(0, Math.min(g.length - 1, far.r + d[0])), c: Math.max(0, Math.min(g[0].length - 1, far.c + d[1])) });
        return true;
      }
      if (k === "Escape" || DIRS[k] || k.length === 1) { t.range = undefined; this.render(); this.host.selectionChanged(); if (k.length !== 1) return true; }
    }
    if (shift && t.cell && DIRS[k]) {
      // Umschalt+Pfeil past the end (start) of the cell's text takes in the neighbouring cell (PowerPoint).
      const atEnd = T.cmpPos(t.caret, T.endPos(body)) === 0, atStart = t.caret.p === 0 && t.caret.o === 0;
      const tbl = this.tableOf();
      if (tbl && ((atEnd && (k === "ArrowRight" || k === "ArrowDown")) || (atStart && (k === "ArrowLeft" || k === "ArrowUp")))) {
        const g = TC.grid(tbl), d = DIRS[k];
        const to = { r: t.cell.r + d[0], c: t.cell.c + d[1] };
        if (to.r >= 0 && to.r < g.length && to.c >= 0 && to.c < g[0].length) { this.setRange(to); return true; }
      }
    }
    switch (k) {
      case "Escape":
        this.endText();
        this.render();
        this.host.selectionChanged();
        return true;
      case "F2":
        this.endText(); this.render(); this.host.selectionChanged(); return true;
      case "ArrowLeft":
      case "ArrowRight": {
        const dir = k === "ArrowLeft" ? -1 : 1;
        if (!shift && hasSel) { collapse(dir < 0 ? T.minPos(t.caret, t.anchor) : T.maxPos(t.caret, t.anchor)); return true; }
        let pos = t.caret;
        if (ctrl) pos = T.wordJump(body, pos, dir);
        else if (dir < 0) pos = pos.o > 0 ? { p: pos.p, o: pos.o - 1 } : pos.p > 0 ? { p: pos.p - 1, o: T.paraLength(ps[pos.p - 1]) } : pos;
        else pos = pos.o < T.paraLength(ps[pos.p]) ? { p: pos.p, o: pos.o + 1 } : pos.p + 1 < ps.length ? { p: pos.p + 1, o: 0 } : pos;
        t.goalX = undefined;
        collapse(pos);
        return true;
      }
      case "ArrowUp":
      case "ArrowDown": {
        if (ctrl) {
          // Strg+Oben/Unten: start of this / the previous / the next paragraph.
          const pos = k === "ArrowUp" ? (t.caret.o > 0 ? { p: t.caret.p, o: 0 } : { p: Math.max(0, t.caret.p - 1), o: 0 }) : { p: Math.min(ps.length - 1, t.caret.p + 1), o: t.caret.p + 1 < ps.length ? 0 : T.paraLength(ps[ps.length - 1]) };
          collapse(pos);
          return true;
        }
        const pos = this.lineMove(t, k === "ArrowUp" ? -1 : 1);
        collapse(pos);
        return true;
      }
      case "Home":
      case "End": {
        if (ctrl) { collapse(k === "Home" ? { p: 0, o: 0 } : T.endPos(body)); return true; }
        collapse(this.lineEdge(t, k === "Home" ? -1 : 1));
        return true;
      }
      case "Backspace":
      case "Delete": {
        let from = t.anchor, to = t.caret;
        if (!hasSel) {
          if (k === "Backspace") {
            if (ctrl) from = T.wordJump(body, t.caret, -1);
            else if (t.caret.o === 0 && t.caret.p > 0) {
              // Rücktaste at the start of an indented paragraph lowers its level first (PowerPoint)?
              // PowerPoint joins the paragraphs.
              from = { p: t.caret.p - 1, o: T.paraLength(ps[t.caret.p - 1]) };
            } else if (t.caret.o === 0) return true;
            else from = { p: t.caret.p, o: t.caret.o - 1 };
          } else {
            // Strg+Entf inside a word deletes the rest of the word, not the space after it (measured).
            if (ctrl) { const w = T.wordAt(ps[t.caret.p], t.caret.o); to = t.caret.o > w.from && t.caret.o < w.to ? { p: t.caret.p, o: w.to } : T.wordJump(body, t.caret, 1); }
            else if (t.caret.o >= T.paraLength(ps[t.caret.p])) { if (t.caret.p + 1 >= ps.length) return true; to = { p: t.caret.p + 1, o: 0 }; }
            else to = { p: t.caret.p, o: t.caret.o + 1 };
            from = t.caret;
          }
        }
        this.edit("Löschen", () => { const pos = T.deleteRange(body, from, to); t.caret = pos; t.anchor = pos; }, `del:${t.id}`);
        return true;
      }
      case "Enter": {
        if (ctrl) return false;
        // The word Eingabe ends is corrected first (AutoKorrektur).
        if (!hasSel && !t.cell && this.autoCorrectOn) this.autoCorrect(String.fromCharCode(10), false);
        this.edit("Eingeben", () => {
          let pos = T.deleteRange(body, t.anchor, t.caret);
          pos = shift ? T.insertBreak(body, pos) : T.splitPara(body, pos);
          t.caret = pos; t.anchor = pos;
        }, `type:${t.id}`);
        return true;
      }
      case "Tab": {
        if (t.cell) { this.nextCell(shift ? -1 : 1); return true; }
        if (ctrl) { this.typeText("\t"); return true; }
        // At the start of a paragraph (or with several selected) Tab changes the level, else it types a tab.
        const allSel = hasSel && T.minPos(t.caret, t.anchor).p !== T.maxPos(t.caret, t.anchor).p;
        // Umschalt+Tab on the top level has no level to go to: it types a tab (measured).
        const lvl = Number(PptxDom.kid(ps[t.caret.p], "pPr")?.getAttribute("lvl") ?? 0);
        if (shift && !hasSel && !allSel && lvl === 0 && t.caret.o > 0) { this.typeText(String.fromCharCode(9)); return true; }
        if (t.caret.o === 0 && T.cmpPos(t.caret, t.anchor) === 0 || allSel || (shift && !hasSel)) { this.changeLevel(shift ? -1 : 1); return true; }
        this.typeText("\t");
        return true;
      }
      case "a":
      case "A":
        if (ctrl && !shift) { t.anchor = { p: 0, o: 0 }; t.caret = T.endPos(body); this.render(); this.host.selectionChanged(); return true; }
        return false;
    }
    if (e.altKey && shift && (k === "ArrowRight" || k === "ArrowLeft")) { this.changeLevel(k === "ArrowRight" ? 1 : -1); return true; }
    return false;
  }

  /** Tab / Umschalt+Tab in a table: the next (previous) cell, its text selected; Tab in the last cell adds a row. */
  private nextCell(d: 1 | -1): void {
    const t = this.text;
    if (!t?.cell) return;
    const el = this.shapeEl(t.id)!;
    const tbl = el.getElementsByTagNameNS(PptxDom.NS_A, "tbl")[0];
    const rows = PptxDom.kids(tbl, "tr");
    const cols = PptxDom.kids(rows[0], "tc").length;
    let { r, c } = t.cell;
    c += d;
    if (c >= cols) { c = 0; r++; }
    if (c < 0) { c = cols - 1; r--; }
    if (r < 0) return;
    if (r >= rows.length) {
      this.edit("Zeile einfügen", () => {
        const last = rows[rows.length - 1];
        const nr = last.cloneNode(true) as Element;
        for (const tc of PptxDom.kids(nr, "tc")) { const b = PptxDom.kid(tc, "txBody"); if (b) T.replaceAll(b, ""); }
        last.after(nr);
        const ext = PptxDom.kid(PptxDom.kid(el, "xfrm"), "ext");
        if (ext) ext.setAttribute("cy", String((PptxDom.num(ext, "cy") ?? 0) + (PptxDom.num(last, "h") ?? 370840)));
      });
    }
    this.enterText(t.id, "all", { r, c });
  }

  /** Listenebene erhöhen / verringern for the paragraphs of the selection (0…8). */
  changeLevel(d: 1 | -1): void {
    const t = this.text;
    const id = t?.id ?? this.sel[0];
    const body = id !== undefined ? this.textBody(id) : undefined;
    if (!body) return;
    const from = t ? T.minPos(t.caret, t.anchor).p : 0, to = t ? T.maxPos(t.caret, t.anchor).p : T.paras(body).length - 1;
    this.edit(d > 0 ? "Listenebene erhöhen" : "Listenebene verringern", () => {
      T.paraProps(body, from, to, (pPr) => {
        const lvl = Math.max(0, Math.min(8, Number(pPr.getAttribute("lvl") ?? 0) + d));
        if (lvl) pPr.setAttribute("lvl", String(lvl)); else pPr.removeAttribute("lvl");
        // Indents written for the old level go, so the new level's apply.
        pPr.removeAttribute("marL"); pPr.removeAttribute("indent");
      });
    });
  }

  /** The position one line up or down, keeping the column the caret started in. */
  private lineMove(t: TextState, dir: -1 | 1): Pos {
    const box = this.caretBox(t.id, t.caret);
    const body = this.textBody(t.id)!;
    if (!box) return t.caret;
    const x = t.goalX ?? box.x;
    const ty = dir < 0 ? box.y - box.h * 0.5 : box.y + box.h * 1.5;
    const pt = this.localToClient(t.id, x, ty);
    const pos = pt ? this.posAt(t.id, pt.x, pt.y) : null;
    const keep = t.goalX ?? box.x;
    let out = pos ?? t.caret;
    const nb = this.caretBox(t.id, out);
    // On the first (last) line the caret stays where it is (measured in PowerPoint; Word would go to the start).
    if (!pos || (nb && Math.abs(nb.y - box.y) < 1)) out = t.caret;
    void body;
    t.goalX = keep;
    return out;
  }

  private lineEdge(t: TextState, dir: -1 | 1): Pos {
    const box = this.caretBox(t.id, t.caret);
    const txt = this.textEl(t.id);
    if (!box || !txt) return t.caret;
    const pt = this.localToClient(t.id, dir < 0 ? -5000 : 5000, box.y + box.h / 2);
    // caretRangeFromPoint needs a point inside the frame: clamp to its edges.
    const r = txt.getBoundingClientRect();
    if (!pt) return t.caret;
    const cx = Math.max(r.left + 1, Math.min(r.right - 1, pt.x));
    const own = this.localToClient(t.id, box.x, box.y + box.h / 2)!;
    const pos = this.posAt(t.id, cx, own.y);
    return pos ?? t.caret;
  }

  // ----- Autofit -----------------------------------------------------------------------------------------

  /** After a change of text: PowerPoint's Autofit for "Text bei Überlauf verkleinern", and growing text boxes. */
  updateAutofit(id: number): void {
    // A fresh view: inside a change the cached one still shows the text before it.
    const v = slideView(this.pptx, this.path!);
    const sv = flat(v.shapes).find((s) => s.id === id);
    if (!sv?.text) return;
    const el = this.shapeEl(id);
    if (!el) return;
    const body = sv.text.body;
    const w = sv.w / EMU_PX, hh = sv.h / EMU_PX;
    const rect = shapeTextRect(sv.format, { ...sv.info, body: { l: body.l / 12700, t: body.t / 12700, r: body.r / 12700, b: body.b / 12700, anchor: "t", wrap: true, autofit: "none" } }, w, hh);
    const probe = h("div", "measure");
    probe.style.cssText = "position:absolute;left:-10000px;top:0;visibility:hidden";
    document.body.append(probe);
    try {
      if (body.autofit === "norm") {
        let found = AUTOFIT_STEPS[AUTOFIT_STEPS.length - 1];
        for (const st of AUTOFIT_STEPS) {
          probe.replaceChildren(renderText(sv.text, rect.w, rect.h, { prompt: false, fit: st }));
          if (textHeight(probe.firstElementChild as HTMLElement) <= rect.h + 0.5) { found = st; break; }
        }
        if (S.setAutofit(el, found)) this.deck.invalidate(this.path!);
      } else if (body.autofit === "shape") {
        probe.replaceChildren(renderText(sv.text, body.wrap ? rect.w : 100000, 100000, { prompt: false }));
        const col = probe.querySelector(".txt-col") as HTMLElement;
        const needH = col.scrollHeight * EMU_PX + body.t + body.b;
        const own = this.ownXfrm(el, sv);
        const next: Partial<S.Xfrm> = { x: own.x, y: own.y, w: own.w, h: Math.round(needH) };
        if (!body.wrap) {
          const width = Math.ceil(Math.max(...Array.from(col.querySelectorAll(".para")).map((p) => (p as HTMLElement).scrollWidth)) * EMU_PX + body.l + body.r);
          next.w = width;
        }
        if (Math.abs((next.h ?? 0) - own.h) > 1 || Math.abs((next.w ?? 0) - own.w) > 1) { S.setXfrm(el, next); this.deck.invalidate(this.path!); }
      }
    } finally {
      probe.remove();
    }
  }
}

function h(tag: string, cls: string): HTMLElement {
  const e = document.createElement(tag);
  e.className = cls;
  return e;
}

/** The text body of table cell (r, c) of a graphic frame. */
function cellBody(el: Element, r: number, c: number): Element | undefined {
  const tbl = el.getElementsByTagNameNS(PptxDom.NS_A, "tbl")[0];
  const tc = PptxDom.kids(PptxDom.kids(tbl, "tr")[r], "tc")[c];
  if (!tc) return undefined;
  let b = PptxDom.kid(tc, "txBody");
  if (!b) { b = PptxDom.fragment(el.ownerDocument, '<a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="de-DE"/></a:p></a:txBody>')[0]; tc.insertBefore(b, tc.firstChild); }
  return b;
}

export function flat(shapes: ShapeView[]): ShapeView[] {
  return shapes.flatMap((s) => [s, ...(s.children ? flat(s.children) : [])]);
}

export function union(views: ShapeView[]): { x: number; y: number; w: number; h: number } {
  const x0 = Math.min(...views.map((v) => v.x)), y0 = Math.min(...views.map((v) => v.y));
  const x1 = Math.max(...views.map((v) => v.x + v.w)), y1 = Math.max(...views.map((v) => v.y + v.h));
  return { x: x0, y: y0, w: x1 - x0, h: y1 - y0 };
}

function rotatedCursor(k: Exclude<Handle, "rot">, rot: number): string {
  const order = ["n", "ne", "e", "se", "s", "sw", "w", "nw"];
  const i = (order.indexOf(k) + Math.round(rot / 45) + 8) % 8;
  return CURSORS[order[i]];
}

/** The box of a drag from `a` to `b`: Umschalt keeps it square (lines: in 45° steps). */
function drawBox(a: { x: number; y: number }, b: { x: number; y: number }, shift: boolean, line: boolean): S.Xfrm {
  let dx = b.x - a.x, dy = b.y - a.y;
  if (shift) {
    if (line) {
      const ang = Math.round(Math.atan2(dy, dx) / (Math.PI / 4)) * (Math.PI / 4), len = Math.hypot(dx, dy);
      dx = Math.round(Math.cos(ang) * len); dy = Math.round(Math.sin(ang) * len);
    } else { const m = Math.max(Math.abs(dx), Math.abs(dy)); dx = Math.sign(dx || 1) * m; dy = Math.sign(dy || 1) * m; }
  }
  return { x: Math.min(a.x, a.x + dx), y: Math.min(a.y, a.y + dy), w: Math.abs(dx), h: Math.abs(dy), rot: 0, flipH: line && dx < 0, flipV: line && dy < 0 };
}

/** A box sized by handle `k` to point `p` (EMU), in the shape's own rotated frame. */
function resize(v: ShapeView, k: Exclude<Handle, "rot">, p: { x: number; y: number }, o: { shift: boolean; ctrl: boolean; keepRatio: boolean }): S.Xfrm {
  const rad = (v.rot * Math.PI) / 180, cos = Math.cos(rad), sin = Math.sin(rad);
  const cx = v.x + v.w / 2, cy = v.y + v.h / 2;
  // The pointer in the shape's frame (origin at its centre, unrotated).
  const lx = (p.x - cx) * cos + (p.y - cy) * sin, ly = -(p.x - cx) * sin + (p.y - cy) * cos;
  let l = -v.w / 2, r = v.w / 2, t = -v.h / 2, b = v.h / 2;
  if (k.includes("w")) l = lx; if (k.includes("e")) r = lx;
  if (k.includes("n")) t = ly; if (k.includes("s")) b = ly;
  if (o.ctrl) {
    // Strg: from the centre.
    if (k.includes("w")) r = -l; if (k.includes("e")) l = -r;
    if (k.includes("n")) b = -t; if (k.includes("s")) t = -b;
  }
  let w = r - l, hh = b - t;
  if (o.keepRatio && v.w && v.h) {
    const ratio = v.w / v.h;
    if (k.length === 2) {
      if (Math.abs(w) / ratio > Math.abs(hh)) hh = Math.sign(hh || 1) * Math.abs(w) / ratio; else w = Math.sign(w || 1) * Math.abs(hh) * ratio;
      if (k.includes("w")) l = r - w; else r = l + w;
      if (k.includes("n")) t = b - hh; else b = t + hh;
      if (o.ctrl) { l = -w / 2; r = w / 2; t = -hh / 2; b = hh / 2; }
    }
  }
  const flipH = (r < l) !== v.flipH, flipV = (b < t) !== v.flipV;
  const nl = Math.min(l, r), nr = Math.max(l, r), nt = Math.min(t, b), nb = Math.max(t, b);
  const mx = (nl + nr) / 2, my = (nt + nb) / 2;
  const ncx = cx + mx * cos - my * sin, ncy = cy + mx * sin + my * cos;
  const W = nr - nl, H = nb - nt;
  return { x: ncx - W / 2, y: ncy - H / 2, w: W, h: H, rot: v.rot, flipH: v.kind === "pic" ? v.flipH : flipH, flipV: v.kind === "pic" ? v.flipV : flipV };
}

/** Whether a shape has a fill that takes clicks (pictures, tables and charts always do). */
function filled(v: ShapeView): boolean {
  if (v.kind !== "sp") return true;
  const f = v.format?.fill ?? v.info?.styleFill;
  return !!f && f.kind !== "none";
}

/** Distance of a point from a line segment. */
function segmentDistance(px: number, py: number, x1: number, y1: number, x2: number, y2: number): number {
  const dx = x2 - x1, dy = y2 - y1, len = dx * dx + dy * dy;
  const t = len ? Math.max(0, Math.min(1, ((px - x1) * dx + (py - y1) * dy) / len)) : 0;
  return Math.hypot(px - (x1 + t * dx), py - (y1 + t * dy));
}
