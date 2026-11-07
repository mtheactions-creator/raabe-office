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

  constructor(private app: PresentApp, from: number, private reading = false) {
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
    document.body.append(this.root);
    addEventListener("keydown", this.keyHandler, true);
    this.root.addEventListener("mousedown", (e) => this.down(e));
    this.root.addEventListener("mousemove", (e) => this.move(e));
    this.root.addEventListener("contextmenu", (e) => { e.preventDefault(); this.menu(e); });
    this.root.addEventListener("wheel", (e) => { if (e.deltaY > 0) this.next(); else this.prev(); });
    addEventListener("resize", this.onResize);
    if (!reading) void document.documentElement.requestFullscreen?.().catch(() => {});
    this.root.focus();
    // Hidden slides are skipped, also as the first one (PowerPoint starts at the next shown).
    while (this.index < this.count - 1 && this.hiddenSlide(this.index)) this.index++;
    this.show(this.index, false);
  }

  private onResize = () => this.layout();

  private get count(): number { return this.app.deck.slides.length; }

  private hiddenSlide(i: number): boolean {
    const p = this.app.deck.slides[i]?.path;
    return !!p && this.app.deck.pptx.doc(p).documentElement.getAttribute("show") === "0";
  }

  state() {
    return { slide: this.index + 1, step: this.stepIdx, steps: this.steps.length, end: this.atEnd, blank: this.blank.hidden ? null : this.blank.dataset.kind ?? null, tool: this.tool, presenter: this.presenter };
  }

  private layout(): void {
    const size = this.app.deck.pptx.slideSize;
    const w = size.cx / EMU_PX, h = size.cy / EMU_PX;
    const area = this.presenter ? this.presenterArea() : this.root.getBoundingClientRect();
    const s = Math.min(area.width / w, area.height / h);
    const left = area.left + (area.width - w * s) / 2, top = area.top + (area.height - h * s) / 2;
    this.stageEl.style.cssText = `left:${left}px;top:${top}px;width:${w}px;height:${h}px;transform:scale(${s});transform-origin:0 0`;
    this.ink.setAttribute("viewBox", `0 0 ${w} ${h}`);
    this.ink.style.cssText = `left:${left}px;top:${top}px;width:${w * s}px;height:${h * s}px`;
  }

  private presenterArea(): DOMRect {
    const r = this.root.getBoundingClientRect();
    return new DOMRect(r.left + 16, r.top + 56, r.width * 0.58, r.height * 0.62);
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
    this.stepIdx = back ? this.steps.length : 0;
    if (!back && !this.setup.noAnimations) for (const id of seq.hidden) { const el = r.shapes.get(id); if (el) el.style.visibility = "hidden"; }
    // Trigger: effects started by a click on a shape; their shapes that appear start hidden.
    this.triggers = this.setup.noAnimations ? [] : interactiveSequences(root).map((x) => ({ ...x, at: 0 }));
    for (const t of this.triggers) for (const id of t.hidden) { const el = r.shapes.get(id); if (el) el.style.visibility = "hidden"; }
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
    }
    if (old && transition && tr && tr.kind === "morph" && this.prevView) {
      this.morph(old, r, view, tr.dur);
    } else if (old && transition && tr && tr.kind !== "cut" && tr.kind !== "none") {
      r.el.animate(transitionKeyframes(tr.kind, tr.dir), { duration: tr.dur, easing: "ease-in-out" }).finished.then(() => old.remove()).catch(() => old.remove());
    } else old?.remove();
    this.prevView = view;
    this.clearInk();
    this.layout();
    if (!back && seq.auto) this.play(seq.auto);
    else if (back && seq.auto) this.finishStep(seq.auto);
    this.renderPresenter();
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
    for (const fx of step.effects) {
      const el = this.shapes.get(fx.spid);
      if (!el) continue;
      if (fx.path) {
        // An Animationspfad: the shape follows the path (fractions of the slide) and stays at its end.
        const size = this.app.deck.pptx.slideSize, w = size.cx / 9525, hh = size.cy / 9525;
        const kf = samplePath(fx.path).map(([px, py]) => ({ translate: `${px * w}px ${py * hh}px` }));
        el.animate(kf, { duration: Math.max(1, fx.dur), delay: fx.delay, fill: "forwards", easing: "ease-in-out" });
        continue;
      }
      const kf = effectKeyframes(fx.cls, fx.effect);
      if (fx.cls === "entr") el.style.visibility = "visible";
      const a = el.animate(kf, { duration: Math.max(1, fx.dur), delay: fx.delay, fill: "forwards", easing: "ease-out" });
      if (fx.cls === "exit") a.finished.then(() => { el.style.visibility = "hidden"; }).catch(() => {});
    }
  }

  next(): void {
    if (!this.blank.hidden) { this.unblank(); return; }
    if (this.atEnd) { this.end(); return; }
    if (this.stepIdx < this.steps.length) { this.play(this.steps[this.stepIdx++]); this.renderPresenter(); return; }
    if (this.zoomReturn && this.index === this.zoomReturn.to) { const back = this.zoomReturn.from; this.zoomReturn = null; this.show(back, false); return; }
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
    let i = this.index - 1;
    while (i > 0 && this.hiddenSlide(i)) i--;
    if (i < 0) return;
    this.show(i, true, false);
  }

  private finishStep(step: Step): void {
    for (const fx of step.effects) {
      const el = this.shapes.get(fx.spid);
      if (el && fx.path) { const p = samplePath(fx.path, 1).pop()!; const size = this.app.deck.pptx.slideSize; el.style.translate = `${(p[0] * size.cx) / 9525}px ${(p[1] * size.cy) / 9525}px`; continue; }
      if (el) el.style.visibility = fx.cls === "exit" ? "hidden" : "visible";
    }
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
    this.renderPresenter();
  }

  private blankScreen(kind: "black" | "white"): void {
    if (!this.blank.hidden && this.blank.dataset.kind === kind) { this.unblank(); return; }
    this.blank.dataset.kind = kind;
    this.blank.style.background = kind === "black" ? "#000" : "#fff";
    this.blank.hidden = false;
  }

  private unblank(): void { this.blank.hidden = true; }

  end(): void {
    this.sound?.pause(); this.sound = null;
    removeEventListener("keydown", this.keyHandler, true);
    removeEventListener("resize", this.onResize);
    clearTimeout(this.timer);
    if (document.fullscreenElement) void document.exitFullscreen().catch(() => {});
    this.root.remove();
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
        case "m": this.inkVisible = !this.inkVisible; this.ink.style.visibility = this.inkVisible ? "" : "hidden"; return;
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
    this.root.dataset.tool = this.tool;
  }

  private toSlide(e: MouseEvent): { x: number; y: number } {
    const r = this.ink.getBoundingClientRect();
    const size = this.app.deck.pptx.slideSize;
    return { x: ((e.clientX - r.left) / r.width) * (size.cx / EMU_PX), y: ((e.clientY - r.top) / r.height) * (size.cy / EMU_PX) };
  }

  private down(e: MouseEvent): void {
    if (e.button !== 0) return;
    if (this.tool === "pen" || this.tool === "highlighter") {
      const p = this.toSlide(e);
      const path = document.createElementNS("http://www.w3.org/2000/svg", "path");
      path.setAttribute("fill", "none");
      path.setAttribute("stroke", this.tool === "pen" ? "#E81123" : "rgba(255,240,0,0.5)");
      path.setAttribute("stroke-width", this.tool === "pen" ? "3" : "14");
      path.setAttribute("stroke-linecap", "round");
      path.setAttribute("stroke-linejoin", "round");
      let d = `M${p.x.toFixed(1)} ${p.y.toFixed(1)}`;
      path.setAttribute("d", d);
      this.ink.append(path);
      const move = (m: MouseEvent) => { const q = this.toSlide(m); d += ` L${q.x.toFixed(1)} ${q.y.toFixed(1)}`; path.setAttribute("d", d); };
      const up = () => { removeEventListener("mousemove", move); removeEventListener("mouseup", up); };
      addEventListener("mousemove", move);
      addEventListener("mouseup", up);
      return;
    }
    if (this.tool === "eraser") {
      const t = document.elementFromPoint(e.clientX, e.clientY);
      if (t && this.ink.contains(t) && t !== this.ink) t.remove();
      return;
    }
    if (this.tool === "laser") return;
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

  private move(e: MouseEvent): void {
    if (this.tool === "arrow") this.root.style.cursor = (e.target as HTMLElement).closest?.("[data-link], [data-action]") ? "pointer" : "";
    if (this.tool === "laser") { this.laser.style.left = `${e.clientX}px`; this.laser.style.top = `${e.clientY}px`; }
  }

  private clearInk(): void { this.ink.replaceChildren(); }

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
    ], new DOMRect(e.clientX, e.clientY, 0, 0));
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
    this.root.append(grid);
  }

  private togglePresenter(): void {
    this.presenter = !this.presenter;
    this.root.classList.toggle("presenter", this.presenter);
    this.layout();
    this.renderPresenter();
  }

  private presenterEl: HTMLElement | null = null;

  /** Referentenansicht: the current slide, the next one, the notes and the time. */
  private renderPresenter(): void {
    this.presenterEl?.remove();
    this.presenterEl = null;
    if (!this.presenter) return;
    const p = document.createElement("div");
    p.className = "presenter-ui";
    const elapsed = Math.floor((Date.now() - this.started) / 1000);
    const time = `${String(Math.floor(elapsed / 3600)).padStart(1, "0")}:${String(Math.floor(elapsed / 60) % 60).padStart(2, "0")}:${String(elapsed % 60).padStart(2, "0")}`;
    const top = document.createElement("div");
    top.className = "presenter-top";
    top.textContent = `${time}    ${new Date().toLocaleTimeString("de-DE", { hour: "2-digit", minute: "2-digit" })}`;
    const nextBox = document.createElement("div");
    nextBox.className = "presenter-next";
    const size = this.app.deck.pptx.slideSize;
    const tw = 360, s = tw / (size.cx / EMU_PX);
    const ni = this.index + 1;
    nextBox.append(Object.assign(document.createElement("div"), { className: "presenter-label", textContent: ni < this.count ? "Nächste Folie" : "Ende der Bildschirmpräsentation" }));
    if (ni < this.count) {
      const box = document.createElement("div");
      box.style.cssText = `width:${tw}px;height:${(size.cy / EMU_PX) * s}px;overflow:hidden;position:relative`;
      const r = renderSlide(this.app.deck.view(this.app.deck.slides[ni].path), size, { editing: false, slideNumber: ni + 1, image: (q) => this.app.image(q), pptx: this.app.deck.pptx });
      r.el.style.transform = `scale(${s})`; r.el.style.transformOrigin = "0 0";
      box.append(r.el);
      nextBox.append(box);
    }
    const notes = document.createElement("div");
    notes.className = "presenter-notes";
    notes.textContent = this.app.deck.notesText(this.app.deck.slides[this.index].path) || "Keine Notizen.";
    const nav = document.createElement("div");
    nav.className = "presenter-nav";
    nav.textContent = `Folie ${this.index + 1} von ${this.count}`;
    p.append(top, nextBox, notes, nav);
    this.root.append(p);
    this.presenterEl = p;
  }
}

/** Keyframes for a slide's entrance by transition kind (our drawing of PowerPoint's effects). */
function transitionKeyframes(kind: string, dir?: string): Keyframe[] {
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

function effectKeyframes(cls: string, effect: string): Keyframe[] {
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
