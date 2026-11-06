// The slide panel at the left of Normalansicht (D-070): numbered thumbnails, the current slide
// framed in the accent colour, Umschalt/Strg for several, dragging to move them with an insertion
// line, Abschnitte as collapsible headings, keys as PowerPoint (Entf, Strg+D, Eingabe for a new slide,
// arrows and Bild auf/ab, Pos1/Ende), a right-click menu. Thumbnails are the real slide drawn small,
// redrawn only when their slide (or its layout) changed.

import type { Deck } from "./deck";
import { renderSlide } from "./render/slide";
import { EMU_PX } from "./render/text";

export interface PanelHost {
  readonly deck: Deck;
  current(): number;
  setCurrent(i: number): void;
  image(path: string): string | undefined;
  panelMenu(e: MouseEvent, onSlide: boolean): void;
  keys(e: KeyboardEvent): boolean;
  focusSlide(): void;
  /** Folienmaster view: the master and its layouts instead of the slides. */
  masterPaths(): string[] | null;
}

export class SlidePanel {
  readonly el: HTMLElement;
  private list: HTMLElement;
  /** Selected slides (indices); the current one is always among them. */
  selected = new Set<number>([0]);
  private anchor = 0;
  private thumbs = new Map<string, { stamp: number; el: HTMLElement; width: number }>();
  /** Collapsed sections by id. */
  collapsed = new Set<string>();
  width = 200;

  constructor(private host: PanelHost) {
    this.el = document.createElement("div");
    this.el.className = "slide-panel";
    this.el.tabIndex = 0;
    this.el.setAttribute("aria-label", "Folien");
    this.list = document.createElement("div");
    this.list.className = "thumbs";
    this.el.append(this.list);
    this.el.addEventListener("keydown", (e) => { if (this.host.keys(e) || this.key(e)) e.preventDefault(); });
    this.el.addEventListener("contextmenu", (e) => {
      e.preventDefault();
      const item = (e.target as HTMLElement).closest(".thumb-item") as HTMLElement | null;
      if (item) {
        const i = Number(item.dataset.i);
        if (!this.selected.has(i)) this.pick(i, false, false);
      }
      this.el.focus();
      this.host.panelMenu(e, !!item);
    });
    new ResizeObserver(() => { const w = this.list.clientWidth; if (w && Math.abs(w - this.width) > 2) { this.width = w; this.render(); } }).observe(this.list);
  }

  get deck(): Deck { return this.host.deck; }

  /** The parts the panel lists: slides, or in the Folienmaster view the master and its layouts. */
  private paths(): { id: number; path: string }[] {
    const m = this.host.masterPaths();
    return m ? m.map((path, i) => ({ id: -1 - i, path })) : this.deck.slides;
  }

  focus(): void { this.el.focus({ preventScroll: true }); }

  /** Selects slide i (Umschalt: a range from the anchor, Strg: toggles). */
  pick(i: number, shift: boolean, ctrl: boolean): void {
    const n = this.paths().length;
    if (!n) return;
    i = Math.max(0, Math.min(n - 1, i));
    if (shift) {
      this.selected = new Set();
      for (let k = Math.min(this.anchor, i); k <= Math.max(this.anchor, i); k++) this.selected.add(k);
    } else if (ctrl) {
      if (this.selected.has(i) && this.selected.size > 1) this.selected.delete(i); else this.selected.add(i);
      this.anchor = i;
    } else {
      this.selected = new Set([i]);
      this.anchor = i;
    }
    this.host.setCurrent(i);
  }

  /** After slides were added, moved or deleted: keep a valid selection. */
  syncSelection(current: number): void {
    const n = this.deck.slides.length;
    this.selected = new Set([...this.selected].filter((i) => i < n));
    if (!this.selected.has(current) && current >= 0) this.selected = new Set([current]);
    this.anchor = current;
  }

  private thumb(i: number, w: number): HTMLElement {
    const ref = this.paths()[i];
    const stamp = this.deck.partStamp(ref.path);
    const hit = this.thumbs.get(ref.path);
    if (hit && hit.stamp === stamp && hit.width === w) return hit.el;
    const size = this.deck.pptx.slideSize;
    const sw = size.cx / EMU_PX, sh = size.cy / EMU_PX;
    const s = w / sw;
    const box = document.createElement("div");
    box.className = "thumb-slide";
    box.style.cssText = `width:${w}px;height:${sh * s}px`;
    const r = renderSlide(this.deck.view(ref.path), size, { editing: false, slideNumber: i + 1, image: (p) => this.host.image(p), pptx: this.deck.pptx });
    r.el.style.transform = `scale(${s})`;
    r.el.style.transformOrigin = "0 0";
    r.el.style.pointerEvents = "none";
    box.append(r.el);
    this.thumbs.set(ref.path, { stamp, el: box, width: w });
    return box;
  }

  render(): void {
    const cur = this.host.current();
    const slides = this.paths();
    const master = !!this.host.masterPaths();
    const sections = master ? [] : this.deck.sections();
    const sectionOf = new Map<number, { name: string; id: string; first: boolean; count: number }>();
    for (const sec of sections) {
      const idx = sec.slides.map((id) => slides.findIndex((s) => s.id === id)).filter((x) => x >= 0).sort((a, b) => a - b);
      idx.forEach((k, j) => sectionOf.set(k, { name: sec.name, id: sec.id, first: j === 0, count: idx.length }));
    }
    const w0 = Math.max(60, (this.list.clientWidth || this.width) - 44);
    const items: HTMLElement[] = [];
    // Sections without slides still show their heading (PowerPoint), at their place in the list.
    slides.forEach((ref, i) => {
      const sec = sectionOf.get(i);
      if (sec?.first) {
        const head = document.createElement("div");
        head.className = "section-head";
        head.dataset.section = sec.id;
        const open = !this.collapsed.has(sec.id);
        head.innerHTML = `<span class="tri">${open ? "&#9662;" : "&#9656;"}</span>`;
        const name = document.createElement("span");
        name.textContent = open ? sec.name : `${sec.name} (${sec.count})`;
        head.append(name);
        head.onclick = () => { if (open) this.collapsed.add(sec.id); else this.collapsed.delete(sec.id); this.render(); };
        items.push(head);
      }
      if (sec && this.collapsed.has(sec.id)) return;
      const item = document.createElement("div");
      item.className = `thumb-item${i === cur ? " current" : ""}${this.selected.has(i) ? " selected" : ""}${master ? (i === 0 ? " master" : " layout") : ""}`;
      item.dataset.i = String(i);
      const num = document.createElement("div");
      num.className = "thumb-num";
      num.textContent = master ? (i === 0 ? "1" : "") : String(i + 1);
      if (!master && this.deck.pptx.doc(ref.path).documentElement.getAttribute("show") === "0") num.classList.add("hidden-slide");
      const frame = document.createElement("div");
      frame.className = "thumb-frame";
      frame.append(this.thumb(i, master && i > 0 ? Math.round(w0 * 0.8) : w0));
      item.append(num, frame);
      item.addEventListener("mousedown", (e) => { if (master) { e.preventDefault(); this.el.focus(); this.pick(i, false, false); } else this.down(e, i); });
      items.push(item);
    });
    this.list.replaceChildren(...items);
    const curEl = this.list.querySelector(".thumb-item.current") as HTMLElement | null;
    curEl?.scrollIntoView({ block: "nearest" });
    // Drop thumbnails of slides that are gone.
    const live = new Set(slides.map((s) => s.path));
    for (const k of this.thumbs.keys()) if (!live.has(k)) this.thumbs.delete(k);
  }

  private down(e: MouseEvent, i: number): void {
    if (e.button !== 0) return;
    e.preventDefault();
    this.el.focus();
    const already = this.selected.has(i);
    if (!already || e.shiftKey || e.ctrlKey) this.pick(i, e.shiftKey, e.ctrlKey);
    else this.host.setCurrent(i);
    const start = { x: e.clientX, y: e.clientY };
    let dragging = false, target = -1;
    const line = document.createElement("div");
    line.className = "drop-line";
    const move = (m: MouseEvent) => {
      if (!dragging && Math.hypot(m.clientX - start.x, m.clientY - start.y) < 5) return;
      dragging = true;
      const items = Array.from(this.list.querySelectorAll(".thumb-item")) as HTMLElement[];
      target = items.length;
      for (const it of items) {
        const r = it.getBoundingClientRect();
        if (m.clientY < r.top + r.height / 2) { target = Number(it.dataset.i); break; }
      }
      const ref = items.find((it) => Number(it.dataset.i) === target);
      const lr = this.list.getBoundingClientRect();
      const y = ref ? ref.getBoundingClientRect().top - lr.top + this.list.scrollTop - 4 : (items[items.length - 1]?.getBoundingClientRect().bottom ?? lr.top) - lr.top + this.list.scrollTop + 4;
      line.style.top = `${y}px`;
      if (!line.parentElement) this.list.append(line);
    };
    const up = () => {
      window.removeEventListener("mousemove", move);
      window.removeEventListener("mouseup", up);
      line.remove();
      if (!dragging) {
        if (already && !e.shiftKey && !e.ctrlKey) { this.selected = new Set([i]); this.anchor = i; this.render(); }
        return;
      }
      if (target < 0) return;
      const sel = [...this.selected].sort((a, b) => a - b);
      // The target counts among the slides that stay.
      const before = sel.filter((s) => s < target).length;
      const to = target - before;
      const moved = sel.map((_, k) => to + k);
      this.deck.moveSlides(sel, to);
      this.selected = new Set(moved);
      this.anchor = moved[0];
      this.host.setCurrent(moved[sel.indexOf(Math.min(...sel))] ?? moved[0]);
    };
    window.addEventListener("mousemove", move);
    window.addEventListener("mouseup", up);
  }

  private key(e: KeyboardEvent): boolean {
    const cur = this.host.current();
    const n = this.paths().length;
    switch (e.key) {
      case "ArrowUp": case "ArrowLeft": case "PageUp": this.pick(cur - 1, e.shiftKey, false); return true;
      case "ArrowDown": case "ArrowRight": case "PageDown": this.pick(cur + 1, e.shiftKey, false); return true;
      case "Home": this.pick(0, e.shiftKey, false); return true;
      case "End": this.pick(n - 1, e.shiftKey, false); return true;
    }
    return false;
  }
}
