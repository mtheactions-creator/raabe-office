// Editing an equation in place, as PowerPoint does (D-072): a caret inside the drawn equation, typing builds
// structures up as it goes (x^2 Leertaste, a/b, \sqrt …: text-engine's MathEditor, shared with RaabeWrite),
// the arrows walk through slots, Rücktaste and Entf take structures apart. The equation is drawn from its
// tree with every atom marked (seqToMathml's data-a/data-s), and the caret is placed from those marks (the
// geometry is RaabeWrite's, features/math.ts). Esc, Eingabe or a click outside ends editing; the result is
// written back as OMML in one undo step. A new equation that stays empty goes again.

import { MathEditor, seqToMathml, type MathPos, type MathStep, type Seq } from "@raabe/text-engine";
import { ommlToSeq } from "@raabe/ooxml";
import type { PresentApp } from "./app";

// ----- where things are in a drawn equation (RaabeWrite's Geometry) ---------------------------------------------

const pathKey = (path: readonly MathStep[]) => path.map((s) => `${s.i}.${s.slot}`).join("/");
const parsePath = (key: string): MathStep[] => (key ? key.split("/").map((s) => { const [i, slot] = s.split("."); return { i: Number(i), slot }; }) : []);

interface Atom { start: number; count: number; el: Element }

class Geometry {
  readonly atoms = new Map<string, Atom[]>();
  readonly slots = new Map<string, Element[]>();

  constructor(readonly root: Element) {
    for (const el of Array.from(root.querySelectorAll("[data-a]"))) {
      const [path, start, count] = (el.getAttribute("data-a") ?? "||").split("|");
      const list = this.atoms.get(path) ?? [];
      list.push({ start: Number(start), count: Number(count), el });
      this.atoms.set(path, list);
    }
    for (const list of this.atoms.values()) list.sort((a, b) => a.start - b.start);
    for (const el of Array.from(root.querySelectorAll("[data-s]"))) {
      const k = el.getAttribute("data-s") ?? "";
      const list = this.slots.get(k) ?? [];
      list.push(el);
      this.slots.set(k, list);
    }
  }

  private xIn(a: Atom, k: number): number {
    const r = a.el.getBoundingClientRect();
    if (k <= 0) return r.left;
    if (k >= a.count) return r.right;
    const text = a.el.firstChild;
    if (text && text.nodeType === Node.TEXT_NODE) {
      const chars = [...(text.textContent ?? "")];
      if (chars.length === a.count) {
        const range = document.createRange();
        range.setStart(text, chars.slice(0, k).join("").length);
        range.collapse(true);
        const rr = range.getBoundingClientRect();
        if (rr.left || rr.width) return rr.left;
      }
    }
    return r.left + (r.width * k) / a.count;
  }

  private visible(path: string): Atom[] {
    return (this.atoms.get(path) ?? []).filter((a) => a.count > 0);
  }

  private extent(path: string, near?: Atom): { top: number; bottom: number } {
    const els = this.slots.get(path) ?? [this.root];
    const el = near ? els.find((e) => e.contains(near.el)) ?? els[0] : els[0];
    const r = el.getBoundingClientRect();
    return { top: r.top, bottom: r.bottom };
  }

  caret(p: MathPos): { x: number; top: number; bottom: number } {
    const key = pathKey(p.path);
    const list = this.visible(key);
    const hit = list.find((a) => p.o >= a.start && p.o < a.start + a.count);
    if (hit) return { x: this.xIn(hit, p.o - hit.start), ...this.extent(key, hit) };
    const before = [...list].reverse().find((a) => a.start + a.count <= p.o);
    if (before) return { x: before.el.getBoundingClientRect().right, ...this.extent(key, before) };
    if (list.length) return { x: list[0].el.getBoundingClientRect().left, ...this.extent(key, list[0]) };
    const ph = (this.atoms.get(key) ?? [])[0];
    const r = (ph?.el ?? this.slots.get(key)?.[0] ?? this.root).getBoundingClientRect();
    return { x: r.left + 0.5, top: r.top, bottom: r.bottom };
  }

  rect(path: readonly MathStep[], from: number, to: number): DOMRect | null {
    let x1 = Infinity, x2 = -Infinity, y1 = Infinity, y2 = -Infinity;
    for (const a of this.visible(pathKey(path))) {
      const lo = Math.max(from, a.start), hi = Math.min(to, a.start + a.count);
      if (lo >= hi) continue;
      const r = a.el.getBoundingClientRect();
      x1 = Math.min(x1, this.xIn(a, lo - a.start));
      x2 = Math.max(x2, this.xIn(a, hi - a.start));
      y1 = Math.min(y1, r.top);
      y2 = Math.max(y2, r.bottom);
    }
    return x1 === Infinity ? null : new DOMRect(x1, y1, x2 - x1, y2 - y1);
  }

  nearest(path: readonly MathStep[], x: number, length: number): number {
    let o = 0, dist = Infinity;
    for (const a of this.visible(pathKey(path))) for (let k = 0; k <= a.count; k++) { const d = Math.abs(this.xIn(a, k) - x); if (d < dist) { dist = d; o = a.start + k; } }
    return Math.min(o, length);
  }

  pick(x: number, y: number): MathPos {
    let best: { key: string; area: number; depth: number } | null = null;
    for (const [key, els] of this.slots) for (const el of els) {
      const r = el.getBoundingClientRect();
      if (x < r.left - 2 || x > r.right + 2 || y < r.top - 2 || y > r.bottom + 2) continue;
      const depth = key ? key.split("/").length : 0, area = r.width * r.height;
      if (!best || depth > best.depth || (depth === best.depth && area < best.area)) best = { key, area, depth };
    }
    const key = best?.key ?? "";
    let o = 0, dist = Infinity;
    for (const a of this.visible(key)) {
      const r = a.el.getBoundingClientRect();
      const dy = y < r.top ? r.top - y : y > r.bottom ? y - r.bottom : 0;
      for (let k = 0; k <= a.count; k++) { const d = Math.abs(this.xIn(a, k) - x) + dy * 4; if (d < dist) { dist = d; o = a.start + k; } }
    }
    return { path: parsePath(key), o };
  }
}

// ----- the session ------------------------------------------------------------------------------------------

interface Session { id: number; me: MathEditor; insert(seq: Seq): void; finish(commit: boolean): void }
let active: Session | null = null;

/** The equation being edited in place, if any (the Formel tab inserts at its caret). */
export function activeMath(): Session | null { return active; }

/**
 * Starts editing equation shape `id` in place. `write` puts a new tree into the shape (one undo step);
 * `remove` takes away a new equation left empty.
 */
export function startMathEdit(app: PresentApp, id: number, o: { isNew?: boolean; point?: { x: number; y: number }; write: (seq: Seq) => void; remove: () => void; m: Element }): void {
  active?.finish(true);
  const ed = app.editor;
  const shp = ed.el.querySelector(`.shp[data-id="${id}"]`) as HTMLElement | null;
  const span = shp?.querySelector(".math") as HTMLElement | null;
  if (!shp || !span) return;
  const seq = ommlToSeq(o.m) ?? [];
  const before = JSON.stringify(seq);
  const display = o.m.localName === "oMathPara";
  const me = new MathEditor(seq, "end");
  const cs = getComputedStyle(span);
  const overlay = document.createElement("span");
  overlay.className = "math-edit";
  const box = document.createElement("div");
  box.className = "math-edit-box";
  const caretEl = document.createElement("div");
  caretEl.className = "math-edit-caret";
  const selEl = document.createElement("div");
  selEl.className = "math-edit-sel";
  document.body.append(box, overlay, selEl, caretEl);
  span.style.visibility = "hidden";
  let g: Geometry | null = null;
  let column: number | null = null;
  const draw = () => {
    const r = span.getBoundingClientRect();
    overlay.style.cssText = `position:fixed;left:${r.left}px;top:${r.top}px;font-size:${cs.fontSize};font-family:'Cambria Math', serif;color:${cs.color};z-index:60;line-height:${cs.lineHeight};white-space:nowrap`;
    overlay.innerHTML = seqToMathml(me.value(), display, true);
    g = new Geometry(overlay.firstElementChild ?? overlay);
    const ob = overlay.getBoundingClientRect();
    box.style.cssText = `position:fixed;left:${ob.left - 3}px;top:${ob.top - 2}px;width:${ob.width + 6}px;height:${ob.height + 4}px;border:1px solid #b5b5b5;background:#f2f2f2;z-index:59;pointer-events:none`;
    const c = g.caret(me.head);
    caretEl.style.cssText = `position:fixed;left:${c.x}px;top:${c.top}px;width:1px;height:${Math.max(6, c.bottom - c.top)}px;background:#000;z-index:61;pointer-events:none`;
    const same = me.anchor.o === me.head.o && pathKey(me.anchor.path) === pathKey(me.head.path);
    const rg = same ? null : me.range();
    const rect = rg ? g.rect(rg.path, rg.from, rg.to) : null;
    selEl.style.cssText = rect ? `position:fixed;left:${rect.left}px;top:${rect.top}px;width:${rect.width}px;height:${rect.height}px;background:rgba(0,90,200,.25);z-index:61;pointer-events:none` : "display:none";
  };
  const finish = (commit: boolean) => {
    if (active !== session) return;
    active = null;
    window.removeEventListener("keydown", onKey, true);
    window.removeEventListener("mousedown", onDown, true);
    overlay.remove(); box.remove(); caretEl.remove(); selEl.remove();
    span.style.visibility = "";
    const v = me.value();
    if (commit && o.isNew && !v.length) { o.remove(); return; }
    if (commit && JSON.stringify(v) !== before) o.write(v);
    ed.select([id]);
    ed.focus();
  };
  const onKey = (e: KeyboardEvent) => {
    const k = e.key, shift = e.shiftKey, ctrl = e.ctrlKey || e.metaKey;
    let handled = true;
    if (k === "Escape" || (k === "Enter" && !shift)) finish(true);
    else if (k === "ArrowLeft") { ctrl ? me.word(-1, shift) : me.left(shift); column = null; }
    else if (k === "ArrowRight") { ctrl ? me.word(1, shift) : me.right(shift); column = null; }
    else if (k === "ArrowUp" || k === "ArrowDown") {
      if (g && column === null) column = g.caret(me.head).x;
      const x = column ?? 0;
      me.vertical(k === "ArrowUp" ? -1 : 1, (path, len) => g?.nearest(path, x, len) ?? 0);
    } else if (k === "Home") me.home(shift);
    else if (k === "End") me.end(shift);
    else if (k === "Backspace") me.backspace();
    else if (k === "Delete") me.del();
    else if (ctrl && (k === "a" || k === "A")) me.selectAll();
    else if (k === "Enter" && shift) me.newLine();
    else if (k.length === 1 && !ctrl && !e.altKey) { me.type(k); column = null; }
    else if (k === "Tab") { /* stays in the equation */ }
    else handled = false;
    if (!handled) return;
    e.preventDefault();
    e.stopPropagation();
    if (active === session) draw();
  };
  const onDown = (e: MouseEvent) => {
    const r = overlay.getBoundingClientRect();
    if (g && e.clientX >= r.left - 4 && e.clientX <= r.right + 4 && e.clientY >= r.top - 4 && e.clientY <= r.bottom + 4) {
      e.preventDefault(); e.stopPropagation();
      const p = g.pick(e.clientX, e.clientY);
      me.head = p;
      if (!e.shiftKey) me.anchor = p;
      column = null;
      draw();
      return;
    }
    finish(true);
  };
  const session: Session = {
    id, me, finish,
    insert(s: Seq) { me.insertAtoms(JSON.parse(JSON.stringify(s)) as Seq); draw(); },
  };
  active = session;
  window.addEventListener("keydown", onKey, true);
  window.addEventListener("mousedown", onDown, true);
  draw();
  if (o.point && g) { const p = (g as Geometry).pick(o.point.x, o.point.y); me.anchor = me.head = p; draw(); }
}
