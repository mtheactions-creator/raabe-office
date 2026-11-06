// Foliensortierung (D-070): the slides as a grid of thumbnails with their numbers, sections as
// headings; click, Umschalt and Strg select, dragging moves the selected slides, a double-click opens
// the slide in Normalansicht, Entf / Strg+D / Strg+C / Strg+V as in the slide panel; the zoom slider
// changes the thumbnails' size.

import type { PresentApp } from "./app";
import { renderSlide } from "./render/slide";
import { EMU_PX } from "./render/text";

export class Sorter {
  readonly el: HTMLElement;
  private grid: HTMLElement;
  zoom = 66;

  constructor(private app: PresentApp) {
    this.el = document.createElement("div");
    this.el.className = "sorter";
    this.el.tabIndex = 0;
    this.el.hidden = true;
    this.grid = document.createElement("div");
    this.grid.className = "sorter-grid";
    this.el.append(this.grid);
    this.el.addEventListener("keydown", (e) => {
      if (this.app.keys(e)) { e.preventDefault(); this.render(); return; }
      const cur = this.app.current();
      const perRow = Math.max(1, Math.floor(this.grid.clientWidth / (this.tileWidth() + 40)));
      const go = (i: number) => { e.preventDefault(); this.app.panel.pick(i, e.shiftKey, false); this.render(); };
      if (e.key === "ArrowRight") go(cur + 1); else if (e.key === "ArrowLeft") go(cur - 1);
      else if (e.key === "ArrowDown") go(cur + perRow); else if (e.key === "ArrowUp") go(cur - perRow);
      else if (e.key === "Home") go(0); else if (e.key === "End") go(this.app.deck.slides.length - 1);
    });
  }

  focus(): void { this.el.focus({ preventScroll: true }); }

  private tileWidth(): number { return Math.round(2.6 * this.zoom); }

  setZoom(z: number): void {
    this.zoom = Math.max(20, Math.min(200, Math.round(z)));
    this.render();
  }

  render(): void {
    const deck = this.app.deck;
    const size = deck.pptx.slideSize;
    const tw = this.tileWidth(), s = tw / (size.cx / EMU_PX);
    const items: HTMLElement[] = [];
    const sections = deck.sections();
    const firstOf = new Map<number, string>();
    for (const sec of sections) {
      const idx = sec.slides.map((id) => deck.slides.findIndex((x) => x.id === id)).filter((x) => x >= 0);
      if (idx.length) firstOf.set(Math.min(...idx), sec.name);
    }
    deck.slides.forEach((ref, i) => {
      const name = firstOf.get(i);
      if (name !== undefined) { const hd = document.createElement("div"); hd.className = "sorter-section"; hd.textContent = name; items.push(hd); }
      const it = document.createElement("div");
      it.className = `sorter-item${this.app.panel.selected.has(i) ? " selected" : ""}${i === this.app.current() ? " current" : ""}`;
      it.dataset.i = String(i);
      const box = document.createElement("div");
      box.className = "sorter-thumb";
      box.style.cssText = `width:${tw}px;height:${(size.cy / EMU_PX) * s}px`;
      const r = renderSlide(deck.view(ref.path), size, { editing: false, slideNumber: i + 1, image: (p) => this.app.image(p), pptx: deck.pptx });
      r.el.style.transform = `scale(${s})`; r.el.style.transformOrigin = "0 0"; r.el.style.pointerEvents = "none";
      box.append(r.el);
      const num = document.createElement("div");
      num.className = "sorter-num";
      num.textContent = String(i + 1);
      it.append(box, num);
      it.addEventListener("mousedown", (e) => { if (e.button === 0) { this.app.panel.pick(i, e.shiftKey, e.ctrlKey); this.render(); this.focus(); } });
      it.addEventListener("dblclick", () => { this.app.setCurrent(i); this.app.setView("normal"); });
      it.addEventListener("contextmenu", (e) => { e.preventDefault(); if (!this.app.panel.selected.has(i)) this.app.panel.pick(i, false, false); this.app.panelMenu(e, true); });
      items.push(it);
    });
    this.grid.replaceChildren(...items);
  }
}
