// Bildschirmpräsentation (D-070): full screen, PowerPoint's keys (N, Eingabe, Bild ab, Rechts, Unten,
// Leertaste and a click go on; P, Bild auf, Links, Oben, Rücktaste go back; a number and Eingabe go
// to that slide; Pos1 / Ende; B or "." black, W or "," white; Esc or "-" end; Strg+P Stift, Strg+L
// Laserpointer, Strg+A arrow, Strg+E Radierer, E erases the ink, Strg+M shows or hides it; H the next
// hidden slide), animations click by click, transitions, the end screen, the right-click menu and the
// Referentenansicht (Alt+F5, or "Referentenansicht anzeigen"). Ink is drawn on top and not kept.

import { showMenu } from "@raabe/shell";
import type { PresentApp } from "./app";
import { renderSlide } from "./render/slide";
import { EMU_PX } from "./render/text";
import { interactiveSequences, samplePath, mainSequence, type Step } from "./animations";
import { readTransition, transitionSound } from "./transitions";
import { showFlag } from "./showsetup";
import { readShowSetup, type ShowSetup } from "./showsetup";
import { mediaOf, mediaType } from "./media";

type Tool = "arrow" | "pen" | "highlighter" | "laser" | "eraser";

/** The Referentenansicht's copy of the current slide (two screens): drawn again, animated with the show. */
interface Mirror { stage: HTMLElement; ink: SVGSVGElement; laser: HTMLElement; shapes: Map<number, HTMLElement> }

/** Two screens (D-074): the audience's window, and the main window showing the Referentenansicht. */
export interface ShowScreens { popup: Window; swap: () => Promise<boolean> }

export class SlideShow {
  readonly done: Promise<void>;
  private finish!: () => void;
  private root: HTMLElement;
  private stageEl: HTMLElement;
  private ink: SVGSVGElement;
  private laser: HTMLElement;
  private blank: HTMLElement;
  private index: number;
  private prevView: import("@raabe/ooxml").SlideView | null = null;
  /** Click steps played on the current slide. */
  private stepIdx = 0;
  private steps: Step[] = [];
  private shapes = new Map<number, HTMLElement>();
  private typed = "";
  private tool: Tool = "arrow";
  private inkVisible = true;
  private atEnd = false;
  private presenter = false;
  private started = Date.now();
  private history: number[] = [];
  /** Folienzoom with Zurück zu Zoom: after slide `to` the show goes back to `from`. */
  private zoomReturn: { from: number; to: number } | null = null;
  private triggers: { trigger: number; steps: Step[]; hidden: Set<number>; at: number }[] = [];
  lastIndex: number;
  private keyHandler = (e: KeyboardEvent) => this.key(e);
  private setup: ShowSetup;
  /** Slides shown (Folien anzeigen: von-bis), 0-based, inclusive. */
  private first = 0;
  private last = 0;
  private timer = 0;

  /** A custom show: only these slides, in this order (Benutzerdefinierte Bildschirmpräsentation). */
  private only: number[] | null = null;

  /** The audience's window and document: the main one, or the show window on the other screen. */
  private win: Window = window;
  private doc: Document = document;
  /** Two screens: the Referentenansicht's root in the main window and its copy of the current slide. */
  private proot: HTMLElement | null = null;
  private mirror: Mirror | null = null;
  private screens: ShowScreens | null = null;
  /** Timer of the Referentenansicht: paused time and when it was paused. */
  private pausedAt = 0;
  /** Folie zoomen: the point (shares of the slide) shown at twice the size, and whether the next click picks it. */
  private zoom: { fx: number; fy: number } | null = null;
  private zoomPick = false;
  private notesSize = 20;
  private clock = 0;

  constructor(private app: PresentApp, from: number, private reading = false, only: number[] | null = null, screens: ShowScreens | null = null) {
    this.only = only && only.length ? only : null;
    this.screens = screens;
    if (screens) { this.win = screens.popup; this.doc = screens.popup.document; this.presenter = true; }
    this.done = new Promise((r) => { this.finish = r; });
    this.setup = readShowSetup(app.deck);
    this.first = this.setup.from !== undefined ? Math.max(0, this.setup.from - 1) : 0;
    this.last = this.setup.to !== undefined ? Math.min(app.deck.slides.length - 1, this.setup.to - 1) : app.deck.slides.length - 1;
    if (from === 0) from = this.first;
    this.index = from;
    this.lastIndex = from;
    this.root = document.createElement("div");
    this.root.className = `show${reading ? " reading" : ""}`;
    this.root.tabIndex = 0;
    this.stageEl = document.createElement("div");
    this.stageEl.className = "show-stage";
    this.ink = document.createElementNS("http://www.w3.org/2000/svg", "svg");
    this.ink.classList.add("show-ink");
    this.laser = document.createElement("div");
    this.laser.className = "show-laser";
    this.laser.hidden = true;
    this.blank = document.createElement("div");
    this.blank.className = "show-blank";
    this.blank.hidden = true;
    this.root.append(this.stageEl, this.ink, this.laser, this.blank);
    this.doc.body.append(this.root);
    addEventListener("keydown", this.keyHandler, true);
    this.bind(this.root, this.ink);
    addEventListener("resize", this.onResize);
    if (screens) {
      // The show window: its own keys and size; closing it (Alt+F4) ends the show.
      this.win.addEventListener("keydown", this.keyHandler, true);
      this.win.addEventListener("resize", this.onResize);
      this.win.addEventListener("pagehide", () => this.end());
      this.proot = document.createElement("div");
      this.proot.className = "show presenter dual";
      this.proot.tabIndex = 0;
      const stage = document.createElement("div");
      stage.className = "show-stage";
      const ink = document.createElementNS("http://www.w3.org/2000/svg", "svg");
      ink.classList.add("show-ink");
      const laser = document.createElement("div");
      laser.className = "show-laser";
      laser.hidden = true;
      this.mirror = { stage, ink, laser, shapes: new Map() };
      this.proot.append(stage, ink, laser);
      document.body.append(this.proot);
      this.bind(this.proot, ink, stage);
      this.proot.focus();
    } else {
      if (!reading) void document.documentElement.requestFullscreen?.().catch(() => {});
      this.root.focus();
    }
    this.clock = window.setInterval(() => this.tick(), 1000);
    // Hidden slides are skipped, also as the first one (PowerPoint starts at the next shown).
    if (this.only) this.index = this.only[0];
    else while (this.index < this.count - 1 && this.hiddenSlide(this.index)) this.index++;
    this.show(this.index, false);
  }

  private onResize = () => this.layout();

  private get count(): number { return this.app.deck.slides.length; }

  private hiddenSlide(i: number): boolean {
    const p = this.app.deck.slides[i]?.path;
    return !!p && this.app.deck.pptx.doc(p).documentElement.getAttribute("show") === "0";
  }

  state() {
    return { slide: this.index + 1, step: this.stepIdx, steps: this.steps.length, end: this.atEnd, blank: this.blank.hidden ? null : this.blank.dataset.kind ?? null, tool: this.tool, presenter: this.presenter, dual: !!this.screens };
  }

  private layout(): void {
    if (this.atEnd) return;
    const size = this.app.deck.pptx.slideSize;
    const w = size.cx / EMU_PX, h = size.cy / EMU_PX;
    const fit = (stage: HTMLElement, ink: SVGSVGElement, area: DOMRect) => {
      const s0 = Math.min(area.width / w, area.height / h);
      let s = s0;
      let left = area.left + (area.width - w * s) / 2, top = area.top + (area.height - h * s) / 2;
      let clip = "";
      if (this.zoom) {
        // Folie zoomen (D-074): twice the size, the chosen point in the middle of the slide's place.
        const box = { l: left, t: top, w: w * s0, h: h * s0 };
        s = s0 * 2;
        left = Math.min(box.l, Math.max(box.l + box.w - w * s, box.l + box.w / 2 - this.zoom.fx * w * s));
        top = Math.min(box.t, Math.max(box.t + box.h - h * s, box.t + box.h / 2 - this.zoom.fy * h * s));
        const x0 = (box.l - left) / s, y0 = (box.t - top) / s, x1 = (box.l + box.w - left) / s, y1 = (box.t + box.h - top) / s;
        clip = `clip-path:inset(${y0}px ${w - x1}px ${h - y1}px ${x0}px);`;
        ink.style.clipPath = "";
      }
      stage.style.cssText = `left:${left}px;top:${top}px;width:${w}px;height:${h}px;transform:scale(${s});transform-origin:0 0;${clip}`;
      ink.setAttribute("viewBox", `0 0 ${w} ${h}`);
      ink.style.cssText = `left:${left}px;top:${top}px;width:${w * s}px;height:${h * s}px;${this.zoom ? `clip-path:inset(${Math.max(0, area.top - top)}px ${Math.max(0, left + w * s - area.right)}px ${Math.max(0, top + h * s - area.bottom)}px ${Math.max(0, area.left - left)}px);` : ""}`;
    };
    if (this.mirror && this.proot) {
      fit(this.stageEl, this.ink, this.root.getBoundingClientRect());
      fit(this.mirror.stage, this.mirror.ink, this.presenter ? this.presenterArea(this.proot) : this.proot.getBoundingClientRect());
      return;
    }
    fit(this.stageEl, this.ink, this.presenter ? this.presenterArea(this.root) : this.root.getBoundingClientRect());
  }

  /** The Referentenansicht's current slide: left, below the timer, above its tools (PowerPoint's layout). */
  private presenterArea(root: HTMLElement): DOMRect {
    const r = root.getBoundingClientRect();
    // The slide sits at the top of its place (its tools follow right below it).
    const sz = this.app.deck.pptx.slideSize, w = r.width * 0.65;
    return new DOMRect(r.left + 16, r.top + r.height * 0.084, w, Math.min(r.height * 0.66, (w * sz.cy) / sz.cx));
  }

  /** Every drawing of a shape: the audience's and the Referentenansicht's copy. */
  private each(id: number, fn: (el: HTMLElement) => void): void {
    const a = this.shapes.get(id);
    if (a) fn(a);
    const b = this.mirror?.shapes.get(id);
    if (b) fn(b);
  }

  /** Mouse input of a surface: the audience's root, or the Referentenansicht's (its slide copy is `stage`). */
  private bind(root: HTMLElement, ink: SVGSVGElement, stage?: HTMLElement): void {
    const onSlide = (e: MouseEvent) => !stage || stage.contains(e.target as Node) || ink.contains(e.target as Node);
    root.addEventListener("mousedown", (e) => { if (onSlide(e)) this.down(e, ink); });
    root.addEventListener("mousemove", (e) => this.move(e, ink));
    root.addEventListener("contextmenu", (e) => { e.preventDefault(); this.menu(e); });
    root.addEventListener("wheel", (e) => { if (!onSlide(e)) return; if (e.deltaY > 0) this.next(); else this.prev(); });
  }

  /** Shows slide i; `back` shows it with all its animations played (going back in PowerPoint). */
  private show(i: number, back: boolean, transition = true): void {
    this.atEnd = false;
    this.history.push(this.index);
    this.index = i;
    this.lastIndex = i;
    const path = this.app.deck.slides[i].path;
    const view = this.app.deck.view(path);
    const r = renderSlide(view, this.app.deck.pptx.slideSize, { editing: false, slideNumber: i + 1, image: (p) => this.app.image(p), pptx: this.app.deck.pptx });
    const root = this.app.deck.pptx.doc(path).documentElement;
    const seq = mainSequence(root);
    this.steps = this.setup.noAnimations ? [] : seq.steps;
    this.shapes = r.shapes;
    if (this.mirror) {
      const m = renderSlide(view, this.app.deck.pptx.slideSize, { editing: false, slideNumber: i + 1, image: (p) => this.app.image(p), pptx: this.app.deck.pptx });
      this.mirror.stage.replaceChildren(m.el);
      this.mirror.shapes = m.shapes;
    }
    this.stepIdx = back ? this.steps.length : 0;
    if (!back && !this.setup.noAnimations) for (const id of seq.hidden) this.each(id, (el) => { el.style.visibility = "hidden"; });
    // Trigger: effects started by a click on a shape; their shapes that appear start hidden.
    this.triggers = this.setup.noAnimations ? [] : interactiveSequences(root).map((x) => ({ ...x, at: 0 }));
    for (const t of this.triggers) for (const id of t.hidden) this.each(id, (el) => { el.style.visibility = "hidden"; });
    const tr = readTransition(root);
    if (transition) this.playTransitionSound(path);
    const old = this.stageEl.firstElementChild as HTMLElement | null;
    this.stageEl.append(r.el);
    // Video and audio play in their box with the browser's controls.
    for (const s of view.shapes) {
      if (s.kind !== "pic") continue;
      const m = mediaOf(this.app, path, s.el);
      const bytes = m ? this.app.deck.pptx.files[m.path] : undefined;
      if (!m || !bytes) continue;
      const el = document.createElement(m.kind);
      el.src = URL.createObjectURL(new Blob([bytes as BlobPart], { type: mediaType(m.path) }));
      el.controls = showFlag(this.app.deck, "media");
      el.style.cssText = "position:absolute;left:0;top:0;width:100%;height:100%;pointer-events:auto;background:#000";
      if (m.kind === "audio") el.style.cssText = "position:absolute;left:0;bottom:-40px;width:300px;pointer-events:auto";
      el.addEventListener("mousedown", (ev) => ev.stopPropagation());
      r.shapes.get(s.id)?.append(el);
      // A narration is heard, not seen (PowerPoint hides its icon in the show: showWhenStopped="0").
      const node = Array.from(root.getElementsByTagNameNS("http://schemas.openxmlformats.org/presentationml/2006/main", "cMediaNode")).find((c) => c.getElementsByTagNameNS("http://schemas.openxmlformats.org/presentationml/2006/main", "spTgt")[0]?.getAttribute("spid") === String(s.id));
      if (node?.getAttribute("showWhenStopped") === "0") { const box = r.shapes.get(s.id); if (box) { box.style.opacity = "0"; box.style.pointerEvents = "none"; } }
    }
    if (old && transition && tr && tr.kind === "morph" && this.prevView) {
      this.morph(old, r, view, tr.dur);
    } else if (old && transition && tr && tr.kind !== "cut" && tr.kind !== "none") {
      r.el.animate(transitionKeyframes(tr.kind, tr.dir, tr.thruBlk), { duration: tr.dur, easing: "ease-in-out" }).finished.then(() => old.remove()).catch(() => old.remove());
    } else old?.remove();
    this.prevView = view;
    this.clearInk();
    this.layout();
    if (!back && seq.auto) this.play(seq.auto);
    else if (back && seq.auto) this.finishStep(seq.auto);
    this.renderPresenter();
    if (this.onSlideChange && this.history.length > 1 && this.history[this.history.length - 1] !== i) this.onSlideChange(i);
    // Nach: the slide goes on by itself after its time (when nothing waits for a click).
    clearTimeout(this.timer);
    if (tr?.advTm !== undefined && !this.reading && this.setup.useTimings) this.timer = window.setTimeout(() => this.next(), tr.advTm + (tr.dur ?? 0));
  }

  /**
   * Morphen (our drawing of PowerPoint's): shapes with the same name on both slides glide from their old
   * place, size and turn to the new one; the others fade out and in.
   */
  private morph(old: HTMLElement, r: { el: HTMLElement; shapes: Map<number, HTMLElement> }, view: import("@raabe/ooxml").SlideView, dur: number): void {
    const before = new Map((this.prevView?.shapes ?? []).map((s) => [s.name, s]));
    const opts = { duration: dur, easing: "ease-in-out" };
    old.animate([{ opacity: 1 }, { opacity: 0 }], opts).finished.then(() => old.remove()).catch(() => old.remove());
    for (const s of view.shapes) {
      const el = r.shapes.get(s.id);
      const b = before.get(s.name);
      if (!el) continue;
      if (!b) { el.animate([{ opacity: 0 }, { opacity: 1 }], opts); continue; }
      const k = 1 / 9525;
      const dx = (b.x + b.w / 2 - (s.x + s.w / 2)) * k, dy = (b.y + b.h / 2 - (s.y + s.h / 2)) * k;
      const sx = s.w ? b.w / s.w : 1, sy = s.h ? b.h / s.h : 1;
      const end = el.style.transform === "none" ? "" : el.style.transform;
      el.animate([{ transform: `translate(${dx}px, ${dy}px) rotate(${b.rot - s.rot}deg) scale(${sx}, ${sy}) ${end}` }, { transform: end || "none" }], opts);
    }
    // The old slide's matched shapes vanish at once (their new selves move in their place).
    const names = new Set(view.shapes.map((s) => s.name));
    for (const el of Array.from(old.querySelectorAll(".shp"))) {
      const id = Number((el as HTMLElement).dataset.id);
      const s = this.prevView?.shapes.find((x) => x.id === id);
      if (s && names.has(s.name)) (el as HTMLElement).style.visibility = "hidden";
    }
  }

  /** The transition's sound (D-072): started with the slide, [Vorherigen Sound anhalten] stops the one playing. */
  private sound: HTMLAudioElement | null = null;
  private playTransitionSound(path: string): void {
    const snd = transitionSound(this.app, path);
    if (!snd) return;
    if (snd.stop || snd.media) { this.sound?.pause(); this.sound = null; }
    const bytes = snd.media ? this.app.deck.pptx.files[snd.media] : undefined;
    if (!bytes) return;
    const a = new Audio(URL.createObjectURL(new Blob([bytes as BlobPart], { type: "audio/wav" })));
    a.loop = !!snd.loop;
    void a.play().catch(() => {});
    this.sound = a;
  }

  private play(step: Step): void {
    for (const fx of step.effects) this.each(fx.spid, (el) => {
      // A media call (Wiedergabe, a narration): the shape's video or audio plays (D-078).
      if (fx.cls === "mediacall") { const m = el.querySelector("video, audio") as HTMLMediaElement | null; if (m && el.ownerDocument === this.doc) window.setTimeout(() => void m.play().catch(() => {}), fx.delay); return; }
      if (fx.path) {
        // An Animationspfad: the shape follows the path (fractions of the slide) and stays at its end.
        const size = this.app.deck.pptx.slideSize, w = size.cx / 9525, hh = size.cy / 9525;
        const kf = samplePath(fx.path).map(([px, py]) => ({ translate: `${px * w}px ${py * hh}px` }));
        el.animate(kf, { duration: Math.max(1, fx.dur), delay: fx.delay, fill: "forwards", easing: "ease-in-out" });
        return;
      }
      const kf = effectKeyframes(fx.cls, fx.effect);
      if (fx.cls === "entr") el.style.visibility = "visible";
      const a = el.animate(kf, { duration: Math.max(1, fx.dur), delay: fx.delay, fill: "forwards", easing: "ease-out" });
      if (fx.cls === "exit") a.finished.then(() => { el.style.visibility = "hidden"; }).catch(() => {});
    });
  }

  next(): void {
    if (!this.blank.hidden) { this.unblank(); return; }
    if (this.atEnd) { this.end(); return; }
    if (this.stepIdx < this.steps.length) { this.play(this.steps[this.stepIdx++]); this.renderPresenter(); return; }
    if (this.zoomReturn && this.index === this.zoomReturn.to) { const back = this.zoomReturn.from; this.zoomReturn = null; this.show(back, false); return; }
    if (this.only) {
      const k = this.only.indexOf(this.index);
      if (k >= 0 && k < this.only.length - 1) { this.show(this.only[k + 1], false); return; }
      if (this.setup.loop) { this.show(this.only[0], false); return; }
      this.endScreen();
      return;
    }
    let i = this.index + 1;
    while (i <= this.last && this.hiddenSlide(i)) i++;
    if (i > this.last) {
      // Wiederholen, bis "Esc" gedrückt wird: from the first slide again.
      if (this.setup.loop) { this.show(this.first, false); return; }
      this.endScreen();
      return;
    }
    this.show(i, false);
  }

  prev(): void {
    if (!this.blank.hidden) { this.unblank(); return; }
    if (this.atEnd) { this.atEnd = false; this.show(this.index, true, false); return; }
    if (this.stepIdx > 0) {
      // Going back undoes the last click's animation (PowerPoint shows the state before it).
      this.stepIdx--;
      const keep = this.stepIdx;
      this.show(this.index, false, false);
      for (let k = 0; k < keep; k++) this.finishStep(this.steps[k]);
      this.stepIdx = keep;
      return;
    }
    if (this.only) {
      const k = this.only.indexOf(this.index);
      if (k > 0) this.show(this.only[k - 1], true, false);
      return;
    }
    let i = this.index - 1;
    while (i > 0 && this.hiddenSlide(i)) i--;
    if (i < 0) return;
    this.show(i, true, false);
  }

  private finishStep(step: Step): void {
    for (const fx of step.effects) this.each(fx.spid, (el) => {
      if (fx.path) { const p = samplePath(fx.path, 1).pop()!; const size = this.app.deck.pptx.slideSize; el.style.translate = `${(p[0] * size.cx) / 9525}px ${(p[1] * size.cy) / 9525}px`; return; }
      el.style.visibility = fx.cls === "exit" ? "hidden" : "visible";
    });
  }

  goto(i: number): void {
    if (i < 0 || i >= this.count) return;
    this.unblank();
    this.show(i, false);
  }

  private endScreen(): void {
    if (this.reading) { this.end(); return; }
    this.atEnd = true;
    const black = document.createElement("div");
    black.className = "show-end";
    black.textContent = "Ende der Bildschirmpräsentation, zum Beenden klicken.";
    this.stageEl.replaceChildren(black);
    const r = this.root.getBoundingClientRect();
    this.stageEl.style.cssText = `left:0;top:0;width:${r.width}px;height:${r.height}px`;
    if (this.mirror) {
      const area = this.presenter && this.proot ? this.presenterArea(this.proot) : this.proot!.getBoundingClientRect();
      this.mirror.stage.replaceChildren(black.cloneNode(true));
      this.mirror.stage.style.cssText = `left:${area.left}px;top:${area.top}px;width:${area.width}px;height:${area.height}px`;
    }
    this.renderPresenter();
  }

  private blankScreen(kind: "black" | "white"): void {
    if (!this.blank.hidden && this.blank.dataset.kind === kind) { this.unblank(); return; }
    this.blank.dataset.kind = kind;
    this.blank.style.background = kind === "black" ? "#000" : "#fff";
    this.blank.hidden = false;
  }

  private unblank(): void { this.blank.hidden = true; }

  /** Aufzeichnen (D-078): told when the show moves to another slide; an overlay (the recording bar) on the presenter's screen. */
  onSlideChange: ((index: number) => void) | null = null;
  addOverlay(el: HTMLElement): void { (this.proot ?? this.root).append(el); }

  private ended = false;
  end(): void {
    if (this.ended) return;
    this.ended = true;
    this.sound?.pause(); this.sound = null;
    removeEventListener("keydown", this.keyHandler, true);
    removeEventListener("resize", this.onResize);
    clearTimeout(this.timer);
    clearInterval(this.clock);
    if (document.fullscreenElement) void document.exitFullscreen().catch(() => {});
    this.root.remove();
    this.proot?.remove();
    if (this.screens) { this.win.removeEventListener("keydown", this.keyHandler, true); try { this.win.close(); } catch { /* closed already */ } }
    this.finish();
  }

  // ----- input -----------------------------------------------------------------------------------------

  private key(e: KeyboardEvent): void {
    e.preventDefault();
    e.stopPropagation();
    const k = e.key, ctrl = e.ctrlKey;
    if (/^\d$/.test(k) && !ctrl) { this.typed += k; return; }
    if (k === "Enter" && this.typed) { const n = Number(this.typed); this.typed = ""; this.goto(n - 1); return; }
    this.typed = "";
    if (ctrl) {
      switch (k.toLowerCase()) {
        case "p": this.setTool("pen"); return;
        case "l": this.setTool("laser"); return;
        case "a": this.setTool("arrow"); return;
        case "e": this.setTool("eraser"); return;
        case "i": this.setTool("highlighter"); return;
        case "m": this.inkVisible = !this.inkVisible; for (const k of this.inks()) k.style.visibility = this.inkVisible ? "" : "hidden"; return;
        case "h": this.root.style.cursor = "none"; return;
        case "s": this.allSlides(); return;
      }
      return;
    }
    if (e.altKey && k === "F5") { this.togglePresenter(); return; }
    switch (k) {
      case "n": case "N": case "Enter": case "PageDown": case "ArrowRight": case "ArrowDown": case " ": this.next(); return;
      case "p": case "P": case "PageUp": case "ArrowLeft": case "ArrowUp": case "Backspace": this.prev(); return;
      case "Home": this.goto(0); return;
      case "End": this.goto(this.count - 1); return;
      case "b": case "B": case ".": this.blankScreen("black"); return;
      case "w": case "W": case ",": this.blankScreen("white"); return;
      case "e": case "E": this.clearInk(); return;
      case "h": case "H": { const i = this.index + 1; if (i < this.count && this.hiddenSlide(i)) this.show(i, false); else this.next(); return; }
      case "Escape": case "-": this.end(); return;
      case "Tab": return;
    }
  }

  private setTool(t: Tool): void {
    this.tool = this.tool === t && t !== "arrow" ? "arrow" : t;
    this.laser.hidden = this.tool !== "laser";
    if (this.mirror) this.mirror.laser.hidden = this.tool !== "laser";
    this.root.dataset.tool = this.tool;
    if (this.proot) this.proot.dataset.tool = this.tool;
    this.renderPresenter();
  }

  /** The ink layers: the audience's and, with two screens, the Referentenansicht's copy. */
  private inks(): SVGSVGElement[] { return this.mirror ? [this.ink, this.mirror.ink] : [this.ink]; }

  private toSlide(e: MouseEvent, ink: SVGSVGElement = this.ink): { x: number; y: number } {
    const r = ink.getBoundingClientRect();
    const size = this.app.deck.pptx.slideSize;
    return { x: ((e.clientX - r.left) / r.width) * (size.cx / EMU_PX), y: ((e.clientY - r.top) / r.height) * (size.cy / EMU_PX) };
  }

  private down(e: MouseEvent, ink: SVGSVGElement = this.ink): void {
    if (e.button !== 0) return;
    const win = (ink.ownerDocument.defaultView ?? window) as Window;
    if (this.tool === "pen" || this.tool === "highlighter") {
      // Drawn on either screen, the stroke appears on both (slide coordinates).
      const p = this.toSlide(e, ink);
      const paths = this.inks().map((k) => {
        const path = document.createElementNS("http://www.w3.org/2000/svg", "path");
        path.setAttribute("fill", "none");
        path.setAttribute("stroke", this.tool === "pen" ? "#E81123" : "rgba(255,240,0,0.5)");
        path.setAttribute("stroke-width", this.tool === "pen" ? "3" : "14");
        path.setAttribute("stroke-linecap", "round");
        path.setAttribute("stroke-linejoin", "round");
        k.append(path);
        return path;
      });
      let d = `M${p.x.toFixed(1)} ${p.y.toFixed(1)}`;
      for (const path of paths) path.setAttribute("d", d);
      const move = (m: MouseEvent) => { const q = this.toSlide(m, ink); d += ` L${q.x.toFixed(1)} ${q.y.toFixed(1)}`; for (const path of paths) path.setAttribute("d", d); };
      const up = () => { win.removeEventListener("mousemove", move); win.removeEventListener("mouseup", up); };
      win.addEventListener("mousemove", move);
      win.addEventListener("mouseup", up);
      return;
    }
    if (this.tool === "eraser") {
      const t = ink.ownerDocument.elementFromPoint(e.clientX, e.clientY);
      if (t && ink.contains(t) && t !== ink) {
        const i = Array.from(ink.children).indexOf(t);
        for (const k of this.inks()) k.children[i]?.remove();
      }
      return;
    }
    if (this.tool === "laser") return;
    // Folie zoomen: the click picks the point; a click on the zoomed slide goes back.
    if (this.zoomPick || this.zoom) {
      if (this.zoom) this.zoom = null;
      else { const r = ink.getBoundingClientRect(); this.zoom = { fx: (e.clientX - r.left) / r.width, fy: (e.clientY - r.top) / r.height }; }
      this.zoomPick = false;
      this.root.style.cursor = "";
      this.layout();
      this.renderPresenter();
      return;
    }
    // A trigger shape plays its next step instead of going on.
    const shp = (e.target as HTMLElement).closest?.(".shp[data-id]") as HTMLElement | null;
    const trig = shp ? this.triggers.find((t) => t.trigger === Number(shp.dataset.id) && t.at < t.steps.length) : undefined;
    if (trig) { this.play(trig.steps[trig.at++]); return; }
    // Folienzoom: to its slide (and back afterwards when Zurück zu Zoom is on).
    const zoom = (e.target as HTMLElement).closest?.("[data-zoom]") as HTMLElement | null;
    if (zoom) {
      // Abschnittszoom: the section's slides, then back.
      const sec = zoom.dataset.zoomSection ? this.app.deck.sections().find((s) => s.id === zoom.dataset.zoomSection) : undefined;
      const ids = sec ? sec.slides : [Number(zoom.dataset.zoom)];
      const idx = ids.map((id) => this.app.deck.slides.findIndex((s) => s.id === id)).filter((i) => i >= 0);
      if (idx.length) { this.zoomReturn = zoom.dataset.zoomBack ? { from: this.index, to: Math.max(...idx) } : null; this.goto(Math.min(...idx)); }
      return;
    }
    // Makro ausführen (D-079): the macro runs with this show as its SlideShowWindow.View.
    const macro = (e.target as HTMLElement).closest?.("[data-macro]") as HTMLElement | null;
    if (macro) {
      const ctl = { position: () => this.index + 1, next: () => this.next(), prev: () => this.prev(), goto: (i: number) => this.goto(i), end: () => this.end(), refresh: () => this.show(this.index, true, false) };
      void import("./macros").then((m) => m.runMacro(this.app, macro.dataset.macro!, ctl));
      return;
    }
    // A shape's action (Einfügen > Aktion) or a link is followed instead of going on (PowerPoint).
    const action = (e.target as HTMLElement).closest?.("[data-action]") as HTMLElement | null;
    if (action) {
      const j = action.dataset.action;
      if (j === "nextslide") this.next(); else if (j === "previousslide") this.prev(); else if (j === "firstslide") this.goto(0);
      else if (j === "lastslide") this.goto(this.count - 1); else if (j === "endshow") this.end();
      else if (j === "lastslideviewed") { const last = this.history[this.history.length - 2]; if (last !== undefined) this.goto(last); }
      return;
    }
    const link = (e.target as HTMLElement).closest?.("[data-link]") as HTMLElement | null;
    if (link) {
      void import("./links").then(({ linkTarget }) => {
        const t = linkTarget(this.app, this.app.deck.slides[this.index].path, link.dataset.link!);
        if (t.slide !== undefined) this.goto(t.slide);
        else if (t.url) void this.app.platform.openExternal?.(t.url);
      });
      return;
    }
    this.next();
  }

  private move(e: MouseEvent, ink: SVGSVGElement = this.ink): void {
    if (this.tool === "arrow") this.root.style.cursor = (e.target as HTMLElement).closest?.("[data-link], [data-action], [data-macro]") ? "pointer" : "";
    if (this.tool !== "laser") return;
    // The laser shows on both screens at the same place of the slide.
    const from = ink.getBoundingClientRect();
    const fx = (e.clientX - from.left) / from.width, fy = (e.clientY - from.top) / from.height;
    const pairs: [HTMLElement, SVGSVGElement][] = [[this.laser, this.ink]];
    if (this.mirror) pairs.push([this.mirror.laser, this.mirror.ink]);
    for (const [laser, k] of pairs) {
      const r = k.getBoundingClientRect();
      laser.style.left = `${r.left + fx * r.width}px`;
      laser.style.top = `${r.top + fy * r.height}px`;
    }
  }

  private clearInk(): void { for (const k of this.inks()) k.replaceChildren(); }

  private menu(e: MouseEvent): void {
    showMenu(this.app.commands, [
      { label: "Weiter", run: () => this.next() }, { label: "Zurück", run: () => this.prev() },
      { label: "Zuletzt angesehen", run: () => { const last = this.history[this.history.length - 2]; if (last !== undefined) this.goto(last); } }, "-",
      { label: "Alle Folien anzeigen", run: () => this.allSlides() }, "-",
      { label: this.presenter ? "Referentenansicht ausblenden" : "Referentenansicht anzeigen", run: () => this.togglePresenter() },
      { label: "Bildschirm", submenu: [{ label: "Schwarzer Bildschirm", run: () => this.blankScreen("black") }, { label: "Weißer Bildschirm", run: () => this.blankScreen("white") }] },
      { label: "Zeigeroptionen", submenu: [{ label: "Laserpointer", checked: this.tool === "laser", run: () => this.setTool("laser") }, { label: "Stift", checked: this.tool === "pen", run: () => this.setTool("pen") },
        { label: "Textmarker", checked: this.tool === "highlighter", run: () => this.setTool("highlighter") }, { label: "Radierer", checked: this.tool === "eraser", run: () => this.setTool("eraser") },
        { label: "Alle Freihandanmerkungen auf der Folie löschen", run: () => this.clearInk() }] }, "-",
      { label: "Präsentation beenden", run: () => this.end() },
    ], this.menuAt(e));
  }

  /** A right-click on the audience's screen opens its menu in the Referentenansicht (one page draws both). */
  private menuAt(e: MouseEvent): DOMRect {
    const own = (e.target as Node).ownerDocument === document;
    if (own || !this.proot) return new DOMRect(e.clientX, e.clientY, 0, 0);
    const r = this.presenterArea(this.proot);
    return new DOMRect(r.left + r.width / 2, r.top + r.height / 2, 0, 0);
  }

  /** Alle Folien anzeigen: thumbnails to pick from. */
  private allSlides(): void {
    const grid = document.createElement("div");
    grid.className = "show-all";
    const size = this.app.deck.pptx.slideSize;
    const tw = 220, s = tw / (size.cx / EMU_PX);
    this.app.deck.slides.forEach((ref, i) => {
      const b = document.createElement("button");
      b.className = "show-all-item";
      const box = document.createElement("div");
      box.style.cssText = `width:${tw}px;height:${(size.cy / EMU_PX) * s}px;overflow:hidden;position:relative`;
      const r = renderSlide(this.app.deck.view(ref.path), size, { editing: false, slideNumber: i + 1, image: (p) => this.app.image(p), pptx: this.app.deck.pptx });
      r.el.style.transform = `scale(${s})`; r.el.style.transformOrigin = "0 0";
      box.append(r.el);
      b.append(box, Object.assign(document.createElement("span"), { textContent: String(i + 1) }));
      b.onclick = (ev) => { ev.stopPropagation(); grid.remove(); this.goto(i); };
      grid.append(b);
    });
    grid.addEventListener("mousedown", (ev) => ev.stopPropagation());
    (this.proot ?? this.root).append(grid);
  }

  private togglePresenter(): void {
    this.presenter = !this.presenter;
    this.root.classList.toggle("presenter", this.presenter && !this.proot);
    this.proot?.classList.toggle("presenter", this.presenter);
    this.layout();
    this.renderPresenter();
  }

  private presenterEl: HTMLElement | null = null;

  /** The Referentenansicht's timer: time shown since the start, without the paused time. */
  private elapsed(): number {
    const now = this.pausedAt || Date.now();
    return Math.max(0, Math.floor((now - this.started) / 1000));
  }

  private tick(): void {
    const t = this.presenterEl?.querySelector(".presenter-time");
    if (t) t.textContent = this.timeText();
    const c = this.presenterEl?.querySelector(".presenter-clock");
    if (c) c.textContent = new Date().toLocaleTimeString("de-DE", { hour: "2-digit", minute: "2-digit" });
  }

  private timeText(): string {
    const s = this.elapsed();
    return `${Math.floor(s / 3600)}:${String(Math.floor(s / 60) % 60).padStart(2, "0")}:${String(s % 60).padStart(2, "0")}`;
  }

  /**
   * Referentenansicht (PowerPoint's layout): a bar with Anzeigeeinstellungen and Bildschirmpräsentation
   * beenden; the timer (anhalten, neu starten) and the clock; the current slide with its tools below
   * (Stift- und Laserpointertools, Alle Folien anzeigen, schwarz schalten, weitere Optionen) and the
   * navigation; the next slide and the notes (Text vergrößern, verkleinern) on the right.
   */
  private renderPresenter(): void {
    this.presenterEl?.remove();
    this.presenterEl = null;
    if (!this.presenter) return;
    const host = this.proot ?? this.root;
    const el = <K extends keyof HTMLElementTagNameMap>(tag: K, cls: string, text = "") => { const e = document.createElement(tag); e.className = cls; if (text) e.textContent = text; return e; };
    const btn = (label: string, glyph: string, run: (ev: MouseEvent) => void, on = false) => {
      const b = el("button", `presenter-btn${on ? " on" : ""}`, glyph);
      b.title = label;
      b.setAttribute("aria-label", label);
      b.addEventListener("mousedown", (ev) => ev.stopPropagation());
      b.addEventListener("click", (ev) => { ev.stopPropagation(); run(ev); host.focus(); });
      return b;
    };
    const at = (ev: MouseEvent) => { const r = (ev.currentTarget as HTMLElement).getBoundingClientRect(); return new DOMRect(r.left, r.bottom, 0, 0); };
    const p = el("div", "presenter-ui");
    // The bar at the top.
    const bar = el("div", "presenter-bar");
    const display = btn("Anzeigeeinstellungen", "Anzeigeeinstellungen \u25BE", (ev) => showMenu(this.app.commands, [
      { label: "Referentenansicht und Bildschirmpräsentation tauschen", disabled: !this.screens, run: () => void this.screens?.swap().then(() => this.layout()) },
      { label: "Bildschirmpräsentation duplizieren", checked: !!this.proot && !this.presenter, disabled: !this.screens, run: () => this.togglePresenter() },
    ], at(ev)));
    display.classList.add("text");
    const endBtn = btn("Bildschirmpräsentation beenden", "Bildschirmpräsentation beenden", () => this.end());
    endBtn.classList.add("text");
    // Taskleiste anzeigen: the main window leaves full screen (maximised), so Windows' taskbar shows.
    const taskbar = btn("Taskleiste anzeigen", "Taskleiste anzeigen", () => void import("./screens").then((m) => m.setMainFullscreen(false)).then(() => this.layout()));
    taskbar.classList.add("text");
    bar.append(taskbar, display, endBtn);
    // Timer and clock.
    const timer = el("div", "presenter-timer");
    timer.append(
      btn(this.pausedAt ? "Timer fortsetzen" : "Timer anhalten", this.pausedAt ? "\u25B6" : "\u275A\u275A", () => {
        if (this.pausedAt) { this.started += Date.now() - this.pausedAt; this.pausedAt = 0; } else this.pausedAt = Date.now();
        this.renderPresenter();
      }),
      el("span", "presenter-time", this.timeText()),
      btn("Timer neu starten", "\u21BB", () => { this.started = Date.now(); if (this.pausedAt) this.pausedAt = this.started; this.tick(); }),
    );
    const clock = el("div", "presenter-clock", new Date().toLocaleTimeString("de-DE", { hour: "2-digit", minute: "2-digit" }));
    // Tools under the current slide.
    const tools = el("div", "presenter-tools");
    tools.append(
      btn("Stift- und Laserpointertools", "\u270E", (ev) => showMenu(this.app.commands, [
        { label: "Laserpointer", checked: this.tool === "laser", run: () => this.setTool("laser") }, { label: "Stift", checked: this.tool === "pen", run: () => this.setTool("pen") },
        { label: "Textmarker", checked: this.tool === "highlighter", run: () => this.setTool("highlighter") }, { label: "Radierer", checked: this.tool === "eraser", run: () => this.setTool("eraser") },
        { label: "Alle Freihandanmerkungen auf der Folie löschen", run: () => this.clearInk() },
      ], at(ev)), this.tool !== "arrow"),
      btn("Alle Folien anzeigen", "\u25A6", () => this.allSlides()),
      btn("Folie zoomen", "\u2315", () => { if (this.zoom) { this.zoom = null; this.layout(); } else { this.zoomPick = !this.zoomPick; (this.proot ?? this.root).style.cursor = this.zoomPick ? "zoom-in" : ""; } this.renderPresenter(); }, this.zoomPick || !!this.zoom),
      btn("Bildschirmpräsentation schwarz schalten oder nicht schwarz schalten", "\u25A0", () => this.blankScreen("black"), !this.blank.hidden),
      btn("Weitere Optionen für Bildschirmpräsentationen", "\u2026", (ev) => showMenu(this.app.commands, [
        { label: "Zuletzt angesehen", run: () => { const last = this.history[this.history.length - 2]; if (last !== undefined) this.goto(last); } }, "-",
        { label: "Weißer Bildschirm", run: () => this.blankScreen("white") },
        { label: this.presenter ? "Referentenansicht ausblenden" : "Referentenansicht anzeigen", run: () => this.togglePresenter() }, "-",
        { label: "Präsentation beenden", run: () => this.end() },
      ], at(ev))),
    );
    tools.style.top = `${this.presenterArea(host).bottom - host.getBoundingClientRect().top + 14}px`;
    const nav = el("div", "presenter-nav");
    const shown = this.only ? this.only.indexOf(this.index) + 1 : this.index + 1;
    const total = this.only ? this.only.length : this.count;
    nav.append(btn("Vorherige Animation oder Folie", "\u25C0", () => this.prev()), el("span", "presenter-pos", this.atEnd ? "Ende" : `Folie ${shown} von ${total}`), btn("Nächste Animation oder Folie", "\u25B6", () => this.next()));
    // The next slide.
    const nextBox = el("div", "presenter-next");
    const size = this.app.deck.pptx.slideSize;
    const hostRect = host.getBoundingClientRect();
    const tw = Math.max(200, hostRect.width * 0.308), s = tw / (size.cx / EMU_PX);
    const ni = this.nextIndex();
    nextBox.append(el("div", "presenter-label", ni !== null ? (this.stepIdx < this.steps.length ? "Nächste Animation" : "Nächste Folie") : "Ende der Bildschirmpräsentation"));
    if (ni !== null) {
      const box = document.createElement("div");
      box.style.cssText = `width:${tw}px;height:${(size.cy / EMU_PX) * s}px;overflow:hidden;position:relative`;
      const r = renderSlide(this.app.deck.view(this.app.deck.slides[ni].path), size, { editing: false, slideNumber: ni + 1, image: (q) => this.app.image(q), pptx: this.app.deck.pptx });
      r.el.style.transform = `scale(${s})`; r.el.style.transformOrigin = "0 0";
      box.append(r.el);
      nextBox.append(box);
    }
    // The notes, with their text size.
    const notes = el("div", "presenter-notes", this.app.deck.notesText(this.app.deck.slides[this.index].path) || "Keine Notizen.");
    notes.style.fontSize = `${this.notesSize}px`;
    notes.style.top = `${hostRect.height * 0.062 + 30 + (ni !== null ? (size.cy / EMU_PX) * s + 16 : 0)}px`;
    const zoom = el("div", "presenter-notes-zoom");
    zoom.append(btn("Text vergrößern", "A\u207A", () => { this.notesSize = Math.min(72, this.notesSize + 4); this.renderPresenter(); }),
      btn("Text verkleinern", "A\u207B", () => { this.notesSize = Math.max(8, this.notesSize - 4); this.renderPresenter(); }));
    p.append(bar, timer, clock, tools, nav, nextBox, notes, zoom);
    host.append(p);
    this.presenterEl = p;
  }

  /** The slide after the current one in this show (hidden slides skipped, a custom show's order), or null. */
  private nextIndex(): number | null {
    if (this.atEnd) return null;
    if (this.stepIdx < this.steps.length) return this.index;
    if (this.only) { const k = this.only.indexOf(this.index); return k >= 0 && k < this.only.length - 1 ? this.only[k + 1] : null; }
    let i = this.index + 1;
    while (i <= this.last && this.hiddenSlide(i)) i++;
    return i <= this.last ? i : null;
  }
}

/** Keyframes for a slide's entrance by transition kind (our drawing of PowerPoint's effects). */
export function transitionKeyframes(kind: string, dir?: string, thruBlk?: boolean): Keyframe[] {
  // Verblassen durch Schwarz: the new slide comes in black over the old one, then lights up.
  if (kind === "fade" && thruBlk) return [{ opacity: 0, filter: "brightness(0)" }, { opacity: 1, filter: "brightness(0)", offset: 0.5 }, { opacity: 1, filter: "brightness(1)" }];
  const from = dir === "r" ? "-100%" : dir === "u" ? "0, 100%" : dir === "d" ? "0, -100%" : "100%";
  switch (kind) {
    case "push": case "cover": return [{ transform: `translate(${from.includes(",") ? from : `${from}, 0`})` }, { transform: "translate(0, 0)" }];
    case "wipe": return [{ clipPath: "inset(0 100% 0 0)" }, { clipPath: "inset(0 0 0 0)" }];
    case "split": return [{ clipPath: "inset(0 50% 0 50%)" }, { clipPath: "inset(0 0 0 0)" }];
    case "circle": return [{ clipPath: "circle(0% at 50% 50%)" }, { clipPath: "circle(75% at 50% 50%)" }];
    case "zoom": return [{ transform: "scale(0.2)", opacity: 0 }, { transform: "scale(1)", opacity: 1 }];
    case "uncover": return [{ opacity: 0 }, { opacity: 1 }];
    default: return [{ opacity: 0 }, { opacity: 1 }];
  }
}

export function effectKeyframes(cls: string, effect: string): Keyframe[] {
  const out = cls === "exit";
  const pair = (a: Keyframe, b: Keyframe) => (out ? [b, a] : [a, b]);
  switch (effect) {
    case "appear": return out ? [{ opacity: 1 }, { opacity: 0 }] : [{ opacity: 1 }, { opacity: 1 }];
    case "fly": return pair({ transform: "translateY(110vh)" }, { transform: "translateY(0)" });
    case "wipe": return pair({ clipPath: "inset(0 0 100% 0)" }, { clipPath: "inset(0 0 0 0)" });
    case "barn": case "split": return pair({ clipPath: "inset(0 50% 0 50%)" }, { clipPath: "inset(0 0 0 0)" });
    case "circle": return pair({ clipPath: "circle(0% at 50% 50%)" }, { clipPath: "circle(75% at 50% 50%)" });
    case "zoom": return pair({ transform: "scale(0)", opacity: 0 }, { transform: "scale(1)", opacity: 1 });
    case "spin": return [{ transform: "rotate(0deg)" }, { transform: "rotate(360deg)" }];
    case "pulse": case "grow": return [{ transform: "scale(1)" }, { transform: "scale(1.1)" }, { transform: "scale(1)" }];
    default: return pair({ opacity: 0 }, { opacity: 1 });
  }
}
