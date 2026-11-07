// Ansicht > Fenster > Alle anordnen and Überlappend (D-074). PowerPoint arranges its document windows on the
// screen; RaabePresent keeps its presentations in one window (D-071's tabs), so the arrangement happens inside
// it: each open presentation (and each Neues Fenster) gets a frame with its name, side by side (Alle anordnen)
// or offset one behind the other (Überlappend). The active one holds the working area (slide panel, slide,
// notes); the others show their current slide and come forward with a click, as a window does. A
// double-click on a frame's title (or Maximieren on it) goes back to one presentation at a time.


import { renderSlide } from "./render/slide";
import { EMU_PX } from "./render/text";

export type Arrangement = "max" | "tile" | "cascade";

export interface WindowDoc { id: string; name: string; current: number; deck: { slides: { path: string }[]; view(p: string): import("@raabe/ooxml").SlideView; pptx: { slideSize: { cx: number; cy: number } } } }

export interface WindowHost {
  docs(): WindowDoc[];
  active(): WindowDoc;
  activate(d: WindowDoc): void;
  /** The working area (slide panel, slide, notes) that the active frame holds. */
  main: HTMLElement;
  image(path: string): string | undefined;
  pptx(d: WindowDoc): import("@raabe/ooxml").Pptx;
  /** After the working area changed size. */
  refit(): void;
}

export class Windows {
  mode: Arrangement = "max";
  private box: HTMLElement | null = null;
  private home: { parent: HTMLElement; next: Node | null } | null = null;

  constructor(private host: WindowHost) {
    addEventListener("resize", () => { if (this.mode !== "max") this.host.refit(); });
  }

  arrange(mode: Arrangement): void {
    this.mode = this.host.docs().length > 1 ? mode : "max";
    this.render();
  }

  /** After the active presentation or the list of presentations changed. */
  render(): void {
    const docs = this.host.docs();
    if (docs.length < 2) this.mode = "max";
    const main = this.host.main;
    if (this.mode === "max") {
      if (this.box && this.home) {
        this.home.parent.insertBefore(main, this.box);
        this.box.remove();
        this.box = null;
        main.classList.remove("in-frame");
        this.host.refit();
      }
      return;
    }
    if (!this.box) {
      this.home = { parent: main.parentElement!, next: main.nextSibling };
      this.box = document.createElement("div");
      this.box.className = "doc-windows";
      this.home.parent.insertBefore(this.box, main);
    }
    this.box.dataset.mode = this.mode;
    this.box.replaceChildren();
    const active = this.host.active();
    const n = docs.length;
    const cols = this.mode === "tile" ? (n <= 3 ? n : Math.ceil(Math.sqrt(n))) : 1;
    this.box.style.gridTemplateColumns = this.mode === "tile" ? `repeat(${cols}, minmax(0, 1fr))` : "";
    // Overlapping: the active frame in front, the others behind in their order.
    const order = this.mode === "cascade" ? [...docs.filter((d) => d !== active), active] : docs;
    order.forEach((d) => {
      const i = docs.indexOf(d);
      const frame = document.createElement("div");
      frame.className = `doc-frame${d === active ? " active" : ""}`;
      if (this.mode === "cascade") frame.style.cssText = `left:${i * 26}px;top:${i * 26}px;right:${(n - 1 - i) * 26}px;bottom:${(n - 1 - i) * 26}px`;
      const title = document.createElement("div");
      title.className = "doc-frame-title";
      title.textContent = d.name;
      title.addEventListener("dblclick", () => { this.host.activate(d); this.arrange("max"); });
      const max = document.createElement("button");
      max.className = "doc-frame-max";
      max.title = "Maximieren";
      max.setAttribute("aria-label", "Maximieren");
      max.textContent = "□";
      max.addEventListener("click", (e) => { e.stopPropagation(); this.host.activate(d); this.arrange("max"); });
      title.append(max);
      const body = document.createElement("div");
      body.className = "doc-frame-body";
      frame.append(title, body);
      if (d === active) {
        main.classList.add("in-frame");
        body.append(main);
      } else {
        body.append(this.preview(d));
        frame.addEventListener("mousedown", () => this.host.activate(d));
      }
      this.box!.append(frame);
    });
    this.host.refit();
  }

  /** An inactive frame: its current slide, fitted. */
  private preview(d: WindowDoc): HTMLElement {
    const wrap = document.createElement("div");
    wrap.className = "doc-frame-preview";
    const ref = d.deck.slides[Math.min(d.current, d.deck.slides.length - 1)];
    if (!ref) return wrap;
    const size = d.deck.pptx.slideSize;
    const r = renderSlide(d.deck.view(ref.path), size, { editing: false, slideNumber: d.current + 1, image: (p) => this.host.image(p), pptx: this.host.pptx(d) });
    const W = size.cx / EMU_PX, H = size.cy / EMU_PX;
    const slot = document.createElement("div");
    slot.className = "doc-frame-slide";
    slot.append(r.el);
    wrap.append(slot);
    // Fitted once it has a size.
    requestAnimationFrame(() => {
      const b = wrap.getBoundingClientRect();
      const s = Math.max(0.05, Math.min((b.width - 24) / W, (b.height - 24) / H));
      slot.style.cssText = `width:${W * s}px;height:${H * s}px`;
      r.el.style.transform = `scale(${s})`;
      r.el.style.transformOrigin = "0 0";
    });
    return wrap;
  }
}
