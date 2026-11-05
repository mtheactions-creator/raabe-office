// Shapes and text boxes (Einfügen > Formen, Textfeld; D-066). A shape is drawn by dragging on the
// sheet after choosing it (a click gives Excel's 2,54 cm square; Umschalt keeps it square), floats
// over the cells like pictures, and is selected with a click: eight handles size it (Umschalt keeps
// the proportions), the round handle above turns it, dragging moves it, Entf deletes it, the arrow keys
// nudge it. Typing while it is selected writes its text; a double-click edits the text. Strg+1 opens
// "Form formatieren". Its look until it is formatted is the theme's shape style, as in Excel.

import { h, type MenuEntry, type PaneControl, type PaneSection, type PaneSpec } from "@raabe/shell";
import type { SheetShape } from "@raabe/sheet-engine";
import { presetAdjust, presetGeometry, shadeLinear, SHAPE_GALLERY, SHAPE_NAMES, shapeElement, shapeIcon, shapeTextRect,
  type GalleryEntry, type PictureFormat, type ShapeInfo } from "@raabe/text-engine";
import type { CalcApp } from "./app";

type Box = { x: number; y: number; w: number; h: number };
type Handle = "nw" | "n" | "ne" | "e" | "se" | "s" | "sw" | "w" | "rot" | "p0" | "p1" | `adj${number}`;

/** Lines and connectors (no fill, two end handles). */
const isLinePreset = (prst: string) => prst === "line" || /Connector\d$/.test(prst);

/** Excel's size of a shape made with a click: 2,54 cm (96 px at 100 %). */
const CLICK_SIZE = 96;

const CSS = `
.shape-layer { position: absolute; right: 17px; bottom: 0; overflow: hidden; pointer-events: none; z-index: 3; }
.shape-layer .shp { position: absolute; pointer-events: auto; cursor: move; }
.shape-layer .shp > svg { position: absolute; left: 0; top: 0; overflow: visible; }
.shape-layer .shp-svg { position: absolute; left: 0; top: 0; transform-origin: 0 0; pointer-events: none; }
.shape-layer .shp-text { position: absolute; display: flex; flex-direction: column; box-sizing: border-box; overflow: hidden;
  white-space: pre-wrap; word-break: break-word; font-family: Calibri, Carlito, sans-serif; outline: none; cursor: text; line-height: 1.2; }
.shape-layer .shp-h { position: absolute; width: 8px; height: 8px; margin: -5px 0 0 -5px; border: 1px solid #5f5f5f; border-radius: 50%; background: #fff; display: none; }
.shape-layer .shp.on .shp-h { display: block; }
.shape-layer .shp.on { outline: 1px solid #7a7a7a; }
.shape-layer .shp.line.on { outline: none; }
.shape-layer .shp-h.rot { border-radius: 50%; width: 12px; height: 12px; margin: -7px 0 0 -7px; cursor: grab; }
.shape-layer .shp-h.adj { background: #ffd800; border-color: #5f5f5f; cursor: default; }
.shape-draw { position: absolute; inset: 0; z-index: 6; cursor: crosshair; }
.shape-gallery { display: grid; grid-template-columns: repeat(12, 26px); gap: 1px; padding: 2px 8px 6px; color: #444; }
.shape-gallery button { width: 26px; height: 26px; border: 1px solid transparent; background: none; padding: 2px; display: flex; align-items: center; justify-content: center; }
.shape-gallery button:hover { border-color: #c6c6c6; background: #f3f3f3; }
`;

export function shapesFeature(app: CalcApp) {
  const layer = h("div", { class: "shape-layer" });
  let shown: SheetShape[] | undefined;
  const els: HTMLElement[] = [];
  let selected: SheetShape | null = null;
  let live: { s: SheetShape; box: Box; rot?: number } | null = null;
  let editing: { s: SheetShape; el: HTMLElement } | null = null;
  let drawFor: string | null = null;

  const z = () => app.sheet.view.zoom / 100;
  const accent = () => app.wb.theme?.accent1 ?? "#4472C4";

  function css(): void {
    if (!document.getElementById("raabe-shapes-css")) document.head.append(h("style", { id: "raabe-shapes-css" }, CSS));
  }

  // ----- places ---------------------------------------------------------------------------------

  function anchorAt(x: number, y: number): SheetShape["from"] {
    const hit = app.grid.hitCell(x, y);
    const c = app.grid.rectOnCanvas({ r0: hit.row, c0: hit.col, r1: hit.row, c1: hit.col })!;
    return { col: hit.col, row: hit.row, x: Math.max(0, (x - c.x) / z()), y: Math.max(0, (y - c.y) / z()) };
  }

  function boxOf(s: SheetShape): Box | null {
    const a = app.grid.rectOnCanvas({ r0: s.from.row, c0: s.from.col, r1: s.from.row, c1: s.from.col });
    const b = app.grid.rectOnCanvas({ r0: s.to.row, c0: s.to.col, r1: s.to.row, c1: s.to.col });
    if (!a || !b) return null;
    const x = a.x + s.from.x * z(), y = a.y + s.from.y * z();
    return { x, y, w: b.x + s.to.x * z() - x, h: b.y + s.to.y * z() - y };
  }

  function placeAt(s: SheetShape, b: Box): SheetShape {
    const x0 = Math.min(b.x, b.x + b.w), y0 = Math.min(b.y, b.y + b.h), w = Math.abs(b.w), hh = Math.abs(b.h);
    return { ...s, from: anchorAt(x0, y0), to: anchorAt(x0 + w, y0 + hh), size: { w: w / z(), h: hh / z() } };
  }

  function change(label: string, fn: (list: SheetShape[]) => SheetShape[]): void {
    const sheet = app.sheet;
    app.change(label, () => {
      app.history.whole(sheet);
      const next = fn([...(sheet.shapes ?? [])]);
      sheet.shapes = next.length ? next : undefined;
    });
  }

  /** Replaces the selected shape (one undo step). */
  function edit(label: string, fn: (s: SheetShape) => SheetShape): void {
    const s = selected;
    if (!s) return;
    const next = fn(s);
    change(label, (list) => list.map((q) => (q === s ? next : q)));
    selected = next;
    if (app.taskPane.isOpen && paneFor) openPane();
  }

  // ----- drawing --------------------------------------------------------------------------------

  /**
   * The shape as the shared Office renderer takes it: its own fill and line, else Excel's shape
   * style (shapes: Akzent 1 with its outline 50 % darker, 1 pt; lines Akzent 1, 0,5 pt; text boxes
   * white with a grey 0,75 pt outline, as the style references of a new shape resolve).
   */
  function engineOf(s: SheetShape): { fmt: PictureFormat; info: ShapeInfo; text: string } {
    const a = accent().replace("#", "").toUpperCase();
    const lines = app.wb.themeLines;
    const line = isLinePreset(s.geom) || !!s.connector;
    const kind: ShapeInfo["kind"] = line ? "connector" : s.textBox ? "textbox" : "shape";
    const info: ShapeInfo = {
      kind, adjust: s.adjust,
      styleFill: kind === "shape" ? { kind: "solid", color: a } : kind === "textbox" ? { kind: "solid", color: "FFFFFF" } : { kind: "none" },
      styleLine: kind === "shape" ? { color: shadeLinear(a, s.style?.shade ?? 0.15), width: lines[(s.style?.ln ?? 2) - 1] ?? 1 }
        : kind === "textbox" ? { color: shadeLinear("FFFFFF", 0.5), width: 0.75 }
        : { color: s.style?.shade !== undefined ? shadeLinear(a, s.style.shade) : a, width: lines[(s.style?.ln ?? 1) - 1] ?? 0.5 },
      body: { l: 7.2, t: 3.6, r: 7.2, b: 3.6, anchor: s.anchor ?? "t", wrap: true, autofit: "none" },
    };
    const hex = (c?: string) => (c ?? "").replace("#", "").toUpperCase();
    const end = (t?: string) => (t && t !== "none" ? { type: t as "triangle" } : undefined);
    const fmt: PictureFormat = {
      geometry: s.geom,
      ...(s.fill ? { fill: s.fill.type === "none" ? { kind: "none" } : { kind: "solid", color: hex(s.fill.color) || a, transparency: (s.fill.transparency ?? 0) / 100 } } : {}),
      ...(s.line ? { line: s.line.type === "none" ? null : { color: hex(s.line.color), width: s.line.width ?? -1, dash: s.line.dash, head: end(s.line.head), tail: end(s.line.tail) } } : {}),
    };
    return { fmt, info, text: s.font?.color ?? (kind === "shape" ? "#FFFFFF" : "#000000") };
  }

  /** The shape drawn at its size at 100 % (the caller scales it with the zoom). */
  function drawn(s: SheetShape, w: number, hh: number): HTMLElement {
    const { fmt, info } = engineOf(s);
    const el = shapeElement(Math.max(w, 0.01), Math.max(hh, 0.01), fmt, info);
    if (s.flipH || s.flipV) { el.style.transform = `scale(${s.flipH ? -1 : 1}, ${s.flipV ? -1 : 1})`; el.style.transformOrigin = "50% 50%"; }
    return el;
  }

  /** Called after every paint of the grid. */
  function place(): void {
    const list = app.sheet.shapes;
    if (!list?.length) { if (layer.isConnected) layer.remove(); shown = undefined; selected = null; return; }
    css();
    if (!layer.isConnected) {
      const canvas = app.grid.el.querySelector(".grid-canvas");
      app.grid.el.insertBefore(layer, canvas ? canvas.nextSibling : app.grid.el.firstChild);
    }
    if (shown !== list) {
      if (selected && !list.includes(selected)) selected = list.find((q) => q.raw && q.raw === selected!.raw) ?? null;
      layer.replaceChildren();
      els.length = 0;
      for (const s of list) {
        const line = isLinePreset(s.geom) || !!s.connector;
        const adj = line ? [] : (presetGeometry(s.geom, 100, 100, s.adjust ?? {})?.handles ?? []).map((_, k): Handle => `adj${k}`);
        const handles: Handle[] = line ? ["p0", "p1"] : ["nw", "n", "ne", "e", "se", "s", "sw", "w", "rot", ...adj];
        const el = h("div", { class: `shp${line ? " line" : ""}` }, h("div", { class: "shp-svg" }), ...(line ? [] : [h("div", { class: "shp-text" })]),
          ...handles.map((k) => h("span", { class: `shp-h ${k.startsWith("adj") ? "adj" : k}`, "data-h": k })));
        el.addEventListener("pointerdown", (e) => down(e, s, el));
        el.addEventListener("dblclick", (e) => { e.preventDefault(); e.stopPropagation(); if (!line) startText(s, el, true); });
        els.push(el);
        layer.append(el);
      }
      shown = list;
    }
    const v = app.sheet.view;
    const tl = app.grid.rectOnCanvas({ r0: v.top, c0: v.left, r1: v.top, c1: v.left });
    if (!tl) return;
    Object.assign(layer.style, { left: `${tl.x}px`, top: `${tl.y}px` });
    list.forEach((s, i) => {
      const el = els[i];
      const b0 = boxOf(s);
      if (!b0) { el.style.display = "none"; return; }
      const b = live?.s === s ? live.box : b0;
      const rot = live?.s === s && live.rot !== undefined ? live.rot : s.rot ?? 0;
      Object.assign(el.style, { display: "", left: `${b.x - tl.x}px`, top: `${b.y - tl.y}px`, width: `${Math.max(1, b.w)}px`, height: `${Math.max(1, b.h)}px`,
        transform: rot ? `rotate(${rot}deg)` : "" });
      const lw = Math.max(0.01, b.w / z()), lh = Math.max(0.01, b.h / z());
      const key = JSON.stringify([s.geom, s.fill, s.line, s.adjust, s.flipH, s.flipV, Math.round(lw * 4), Math.round(lh * 4), accent()]);
      const sv = el.querySelector<HTMLElement>(".shp-svg")!;
      if (sv.dataset.key !== key) { sv.replaceChildren(drawn(s, lw, lh)); sv.dataset.key = key; }
      sv.style.transform = `scale(${z()})`;
      const t = el.querySelector<HTMLElement>(".shp-text");
      if (t) {
        // The text in the preset's text rectangle, inside Excel's insets (7,2 / 3,6 pt).
        const e = engineOf(s);
        const r = shapeTextRect(e.fmt, e.info, lw, lh);
        Object.assign(t.style, { left: `${r.x * z()}px`, top: `${r.y * z()}px`, width: `${r.w * z()}px`, height: `${r.h * z()}px` });
      }
      if (t && editing?.s !== s) {
        const l = engineOf(s);
        t.textContent = s.text ?? "";
        Object.assign(t.style, {
          color: l.text, fontSize: `${((s.font?.size ?? 11) * 96 / 72) * z()}px`, fontWeight: s.font?.bold ? "bold" : "", fontStyle: s.font?.italic ? "italic" : "",
          textDecoration: s.font?.underline ? "underline" : "",
          justifyContent: { t: "flex-start", ctr: "center", b: "flex-end" }[s.anchor ?? "t"],
          textAlign: { l: "left", ctr: "center", r: "right" }[s.align ?? "l"],
        });
      }
      const pos: Record<string, [number, number]> = { nw: [0, 0], n: [0.5, 0], ne: [1, 0], e: [1, 0.5], se: [1, 1], s: [0.5, 1], sw: [0, 1], w: [0, 0.5],
        p0: [s.flipH ? 1 : 0, s.flipV ? 1 : 0], p1: [s.flipH ? 0 : 1, s.flipV ? 0 : 1] };
      const gh = presetGeometry(s.geom, lw, lh, s.adjust ?? {})?.handles ?? [];
      for (const hd of el.querySelectorAll<HTMLElement>(".shp-h")) {
        const k = hd.dataset.h as Handle;
        if (k === "rot") { hd.style.left = `${b.w / 2}px`; hd.style.top = `-22px`; continue; }
        if (k.startsWith("adj")) {
          const q = gh[Number(k.slice(3))];
          if (q) { hd.style.left = `${(s.flipH ? lw - q.x : q.x) * z()}px`; hd.style.top = `${(s.flipV ? lh - q.y : q.y) * z()}px`; }
          continue;
        }
        hd.style.left = `${pos[k][0] * b.w}px`;
        hd.style.top = `${pos[k][1] * b.h}px`;
      }
      el.classList.toggle("on", s === selected);
    });
  }

  function select(s: SheetShape | null): void {
    if (editing && editing.s !== s) endText();
    selected = s;
    shown = undefined;
    app.shapeSelectionChanged();
    app.grid.contentChanged();
    if (!s && paneFor) { paneFor = false; app.taskPane.close(); }
  }

  // ----- mouse ----------------------------------------------------------------------------------

  function down(e: PointerEvent, s: SheetShape, el: HTMLElement): void {
    if (e.button !== 0) return;
    if (editing?.s === s && (e.target as HTMLElement).closest(".shp-text")) return;
    e.preventDefault();
    e.stopPropagation();
    if (!app.protectionAllows("objects")) return;
    if (app.grid.isEditing) app.grid.commitEdit("none");
    if (selected !== s) select(s);
    const b = boxOf(s);
    if (!b) return;
    const handle = (e.target as HTMLElement).dataset.h as Handle | undefined;
    if (handle?.startsWith("adj")) { dragAdjust(e, s, el, b, Number(handle.slice(3))); return; }
    const sx = e.clientX, sy = e.clientY;
    el.setPointerCapture(e.pointerId);
    const move = (ev: PointerEvent) => {
      let dx = ev.clientX - sx, dy = ev.clientY - sy;
      if (!handle) { live = { s, box: { ...b, x: b.x + dx, y: b.y + dy } }; place(); return; }
      if (handle === "rot") {
        const cx = b.x + b.w / 2, cy = b.y + b.h / 2;
        const r = app.grid.el.getBoundingClientRect();
        let deg = (Math.atan2(ev.clientY - r.top - cy, ev.clientX - r.left - cx) * 180) / Math.PI + 90;
        if (ev.shiftKey) deg = Math.round(deg / 15) * 15;
        live = { s, box: b, rot: ((deg % 360) + 360) % 360 };
        place();
        return;
      }
      if (handle === "p0" || handle === "p1") {
        // Lines: one end follows the pointer, the other stays.
        const ends = endsOf(s, b);
        const moving = handle === "p0" ? 0 : 1;
        const p = { x: ends[moving].x + dx, y: ends[moving].y + dy }, q = ends[1 - moving];
        const flipH = (moving === 0 ? p.x > q.x : q.x > p.x), flipV = (moving === 0 ? p.y > q.y : q.y > p.y);
        live = { s: { ...s, flipH, flipV }, box: { x: Math.min(p.x, q.x), y: Math.min(p.y, q.y), w: Math.abs(p.x - q.x), h: Math.abs(p.y - q.y) } };
        live.s = s;
        (live as { flip?: [boolean, boolean] }).flip = [flipH, flipV];
        place();
        return;
      }
      let { x, y, w, h: hh } = b;
      if (handle.includes("e")) w += dx;
      if (handle.includes("s")) hh += dy;
      if (handle.includes("w")) { x += dx; w -= dx; }
      if (handle.includes("n")) { y += dy; hh -= dy; }
      if (ev.shiftKey && handle.length === 2) {
        // Umschalt keeps the proportions at the corners.
        const k = Math.max(w / b.w, hh / b.h);
        const nw = b.w * k, nh = b.h * k;
        if (handle.includes("w")) x = b.x + b.w - nw;
        if (handle.includes("n")) y = b.y + b.h - nh;
        w = nw; hh = nh;
      }
      dx = 0; dy = 0;
      live = { s, box: { x: w < 0 ? x + w : x, y: hh < 0 ? y + hh : y, w: Math.abs(w), h: Math.abs(hh) } };
      place();
    };
    const up = () => {
      el.removeEventListener("pointermove", move);
      el.removeEventListener("pointerup", up);
      el.removeEventListener("pointercancel", up);
      const l = live as (typeof live & { flip?: [boolean, boolean] }) | null;
      live = null;
      if (!l) { place(); return; }
      if (l.rot !== undefined) { edit("Drehen", (q) => ({ ...q, rot: Math.round(l.rot!) || undefined })); return; }
      const same = Math.abs(l.box.x - b.x) < 1 && Math.abs(l.box.y - b.y) < 1 && Math.abs(l.box.w - b.w) < 1 && Math.abs(l.box.h - b.h) < 1;
      if (same && !l.flip) { place(); return; }
      edit(handle ? "Größe ändern" : "Verschieben", (q) => {
        const moved = placeAt(q, l.box);
        if (l.flip) { moved.flipH = l.flip[0] || undefined; moved.flipV = l.flip[1] || undefined; }
        return moved;
      });
    };
    el.addEventListener("pointermove", move);
    el.addEventListener("pointerup", up);
    el.addEventListener("pointercancel", up);
  }

  /**
   * A yellow handle: the adjust value whose handle lands nearest the pointer, searched in the
   * evaluated geometry so that every preset follows its own formulas (as in RaabeWrite).
   */
  function dragAdjust(e: PointerEvent, s: SheetShape, el: HTMLElement, b: Box, i: number): void {
    const w = b.w / z(), hh = b.h / z();
    const base = { ...presetAdjust(s.geom), ...(s.adjust ?? {}) };
    const hd = presetGeometry(s.geom, w, hh, base)?.handles[i];
    if (!hd) return;
    const at = (a: Record<string, number>) => presetGeometry(s.geom, w, hh, a)?.handles[i];
    const search = (adj: Record<string, number>, key: string, lo: number, hi: number, pos: (a: Record<string, number>) => number, target: number) => {
      let best = adj[key] ?? lo, bestD = Infinity;
      let a = Math.min(lo, hi), c = Math.max(lo, hi);
      for (let round = 0; round < 3; round++) {
        for (let k = 0; k <= 48; k++) {
          const v = a + ((c - a) * k) / 48;
          const d = Math.abs(pos({ ...adj, [key]: v }) - target);
          if (d < bestD) { bestD = d; best = v; }
        }
        const span = (c - a) / 48;
        a = Math.max(Math.min(lo, hi), best - span);
        c = Math.min(Math.max(lo, hi), best + span);
      }
      return Math.round(best);
    };
    let adj = base;
    el.setPointerCapture(e.pointerId);
    const move = (ev: PointerEvent) => {
      // The pointer in the shape's own coordinates (unrotated, unflipped, at 100 %).
      const r = el.getBoundingClientRect();
      const cx = r.left + r.width / 2, cy = r.top + r.height / 2;
      const rad = (-(s.rot ?? 0) * Math.PI) / 180;
      const dx = ev.clientX - cx, dy = ev.clientY - cy;
      let qx = (dx * Math.cos(rad) - dy * Math.sin(rad)) / z() + w / 2, qy = (dx * Math.sin(rad) + dy * Math.cos(rad)) / z() + hh / 2;
      if (s.flipH) qx = w - qx;
      if (s.flipV) qy = hh - qy;
      const next: Record<string, number> = { ...base };
      if (hd.kind === "polar") {
        if (hd.gdY) {
          let ang = (Math.atan2(qy - hh / 2, qx - w / 2) * 180) / Math.PI;
          if (ang < 0) ang += 360;
          next[hd.gdY] = Math.max(hd.minY ?? 0, Math.min(hd.maxY ?? 21599999, Math.round(ang * 60000)));
        }
        if (hd.gdX) next[hd.gdX] = search(next, hd.gdX, hd.minX ?? 0, hd.maxX ?? 50000, (a) => { const p = at(a); return p ? Math.hypot(p.x - w / 2, p.y - hh / 2) : 0; }, Math.hypot(qx - w / 2, qy - hh / 2));
      } else {
        const big = (v: number | undefined, fallback: number) => (v === undefined || Math.abs(v) > 1e8 ? fallback : v);
        if (hd.gdX) next[hd.gdX] = search(next, hd.gdX, big(hd.minX, -400000), big(hd.maxX, 400000), (a) => at(a)?.x ?? 0, qx);
        if (hd.gdY) next[hd.gdY] = search(next, hd.gdY, big(hd.minY, -400000), big(hd.maxY, 400000), (a) => at(a)?.y ?? 0, qy);
      }
      adj = next;
      const sv = el.querySelector<HTMLElement>(".shp-svg")!;
      sv.replaceChildren(drawn({ ...s, adjust: adj }, w, hh));
      sv.dataset.key = "";
      const q = at(adj);
      const hEl = el.querySelector<HTMLElement>(`[data-h="adj${i}"]`);
      if (q && hEl) { hEl.style.left = `${(s.flipH ? w - q.x : q.x) * z()}px`; hEl.style.top = `${(s.flipV ? hh - q.y : q.y) * z()}px`; }
    };
    const up = () => {
      el.removeEventListener("pointermove", move);
      el.removeEventListener("pointerup", up);
      el.removeEventListener("pointercancel", up);
      if (JSON.stringify(adj) !== JSON.stringify(base)) {
        // Only the values that differ from the preset's defaults are kept (as Excel writes avLst).
        const def = presetAdjust(s.geom);
        const own = Object.fromEntries(Object.entries(adj).filter(([k, v]) => def[k] !== v));
        edit("Form anpassen", (q) => ({ ...q, adjust: Object.keys(own).length ? own : undefined }));
      } else place();
    };
    el.addEventListener("pointermove", move);
    el.addEventListener("pointerup", up);
    el.addEventListener("pointercancel", up);
  }

  function endsOf(s: SheetShape, b: Box): { x: number; y: number }[] {
    const a = { x: s.flipH ? b.x + b.w : b.x, y: s.flipV ? b.y + b.h : b.y };
    const c = { x: s.flipH ? b.x : b.x + b.w, y: s.flipV ? b.y : b.y + b.h };
    return [a, c];
  }

  // ----- text -----------------------------------------------------------------------------------

  function startText(s: SheetShape, el: HTMLElement, selectAll: boolean, first?: string): void {
    const t = el.querySelector<HTMLElement>(".shp-text");
    if (!t) return;
    editing = { s, el: t };
    t.contentEditable = "plaintext-only";
    t.focus();
    if (first !== undefined) t.textContent = (s.text ?? "") + first;
    const r = document.createRange();
    r.selectNodeContents(t);
    if (!selectAll || first !== undefined) r.collapse(false);
    const sel = getSelection();
    sel?.removeAllRanges();
    sel?.addRange(r);
    t.addEventListener("keydown", textKey);
  }

  function textKey(e: KeyboardEvent): void {
    e.stopPropagation();
    if (e.key === "Escape") { e.preventDefault(); endText(); }
  }

  function endText(): void {
    const ed = editing;
    if (!ed) return;
    editing = null;
    ed.el.removeEventListener("keydown", textKey);
    ed.el.contentEditable = "false";
    const text = ed.el.innerText.replace(/\n$/, "");
    if (text !== (ed.s.text ?? "")) {
      const s = ed.s;
      selected = s;
      edit("Text eingeben", (q) => ({ ...q, text: text || undefined }));
    } else place();
    app.grid.focus();
  }

  // ----- keys and clicks elsewhere -----------------------------------------------------------------

  addEventListener("pointerdown", (e) => {
    const t = e.target as HTMLElement;
    if (editing && !t.closest?.(".shp")) endText();
    if (selected && !t.closest?.(".shp") && !t.closest?.(".taskpane") && !t.closest?.(".ribbon-card") && !t.closest?.(".menu") && !t.closest?.(".popup")) select(null);
  }, true);
  addEventListener("keydown", (e) => {
    if (!selected || editing || app.grid.isEditing || !app.protectionAllows("objects")) return;
    if ((e.target as HTMLElement).closest?.(".taskpane, input, textarea, select, .dialog-layer")) return;
    const s = selected;
    if (e.key === "Delete" || e.key === "Backspace") {
      e.preventDefault(); e.stopPropagation();
      selected = null;
      change("Löschen", (list) => list.filter((q) => q !== s));
      app.shapeSelectionChanged();
    } else if (e.key === "Escape") { e.preventDefault(); e.stopPropagation(); select(null); app.grid.focus(); }
    else if (e.key.startsWith("Arrow")) {
      e.preventDefault(); e.stopPropagation();
      const b = boxOf(s);
      if (!b) return;
      const d = e.ctrlKey ? 1 : 1 * z();
      const dx = e.key === "ArrowLeft" ? -d : e.key === "ArrowRight" ? d : 0, dy = e.key === "ArrowUp" ? -d : e.key === "ArrowDown" ? d : 0;
      edit("Verschieben", (q) => placeAt(q, { ...b, x: b.x + dx, y: b.y + dy }));
    } else if (e.ctrlKey && (e.key === "1" || e.code === "Digit1")) { e.preventDefault(); e.stopPropagation(); openPane(); }
    else if (e.key.length === 1 && !e.ctrlKey && !e.altKey && !isLinePreset(s.geom) && !s.connector) {
      // Typing while a shape is selected writes its text (Excel).
      e.preventDefault(); e.stopPropagation();
      const i = (app.sheet.shapes ?? []).indexOf(s);
      if (els[i]) startText(s, els[i], false, e.key);
    }
  }, true);

  // ----- inserting ------------------------------------------------------------------------------

  /** Einfügen > Formen > a shape, or Textfeld: the next drag on the sheet draws it. */
  function startDraw(prst: string, textBox = false, entry?: GalleryEntry): void {
    if (app.grid.isEditing) app.grid.commitEdit("none");
    css();
    drawFor = prst;
    if (entry) { const i = recent.findIndex((r) => r.label === entry.label); if (i >= 0) recent.splice(i, 1); recent.unshift(entry); recent.length = Math.min(recent.length, 24); }
    const ends = entry?.tool.pic?.line;
    const over = h("div", { class: "shape-draw" });
    app.grid.el.append(over);
    let start: { x: number; y: number } | null = null;
    const r0 = app.grid.el.getBoundingClientRect();
    over.addEventListener("pointerdown", (e) => { e.preventDefault(); start = { x: e.clientX - r0.left, y: e.clientY - r0.top }; over.setPointerCapture(e.pointerId); });
    over.addEventListener("pointerup", (e) => {
      over.remove();
      if (!start || !drawFor) return;
      let w = e.clientX - r0.left - start.x, hh = e.clientY - r0.top - start.y;
      const line = isLinePreset(drawFor);
      if (Math.abs(w) < 3 && Math.abs(hh) < 3) { w = CLICK_SIZE * z(); hh = line ? 0 : (textBox ? CLICK_SIZE * 0.4 : CLICK_SIZE) * z(); }
      if (e.shiftKey && !line) { const m = Math.max(Math.abs(w), Math.abs(hh)); w = Math.sign(w || 1) * m; hh = Math.sign(hh || 1) * m; }
      // Excel numbers new objects after all of the sheet's drawing objects ("Rechteck 3").
      const n = (app.sheet.shapes?.length ?? 0) + (app.sheet.pictures?.length ?? 0) + (app.sheet.charts?.length ?? 0) + 1;
      const label = SHAPE_NAMES[drawFor] ?? entry?.label ?? "Form";
      let s: SheetShape = { from: { col: 0, row: 0, x: 0, y: 0 }, to: { col: 0, row: 0, x: 0, y: 0 }, geom: drawFor, name: textBox ? `Textfeld ${n}` : `${label} ${n}` };
      if (line) {
        s.connector = true; if (w < 0) s.flipH = true; if (hh < 0) s.flipV = true;
        // Excel's lines take line style 1 of the Office 2013-2022 theme, style 2 of the current one (recorded).
        s.style = { ln: (app.wb.themeLines[0] ?? 0.5) >= 1 ? 2 : 1 };
      }
      if (ends?.head || ends?.tail) s.line = { type: "solid", ...(ends.head ? { head: ends.head.type } : {}), ...(ends.tail ? { tail: ends.tail.type } : {}) };
      if (textBox) s.textBox = true;
      s = placeAt(s, { x: start.x, y: start.y, w, h: hh });
      change(textBox ? "Textfeld" : "Form einfügen", (list) => [...list, s]);
      drawFor = null;
      selected = s;
      shown = undefined;
      app.shapeSelectionChanged();
      if (textBox) queueMicrotask(() => { place(); const i = (app.sheet.shapes ?? []).indexOf(s); if (els[i]) startText(s, els[i], false); });
    });
    over.addEventListener("keydown", (e) => { if (e.key === "Escape") { over.remove(); drawFor = null; } });
  }

  const recent: GalleryEntry[] = [];

  /** The Formen gallery (Einfügen > Formen and Formformat > Formen einfügen): Office's, recent ones first. */
  function gallery(close: () => void): HTMLElement {
    const out = h("div", { class: "menu shape-menu" });
    const section = (title: string, items: GalleryEntry[]) => {
      if (!items.length) return;
      out.append(h("div", { class: "palette-title" }, title));
      const grid = h("div", { class: "shape-gallery" });
      for (const e of items) {
        const b = h("button", { type: "button", "data-tip": e.label, "aria-label": e.label, onclick: () => { close(); startDraw(e.tool.geometry, e.tool.kind === "textbox", e); } });
        b.append(shapeIcon(e, 20));
        grid.append(b);
      }
      out.append(grid);
    };
    section("Zuletzt verwendete Formen", recent);
    // Freihand shapes are drawn point by point (not yet in RaabeCalc).
    for (const g of SHAPE_GALLERY) section(g.title, g.items.filter((e) => !e.tool.freeform));
    return out;
  }

  // ----- Form formatieren -------------------------------------------------------------------------

  let paneFor = false;

  function openPane(): void {
    if (!selected) return;
    paneFor = true;
    document.documentElement.style.setProperty("--pane-top", `${document.querySelector(".calc-main")?.getBoundingClientRect().top ?? 150}px`);
    app.taskPane.onClose = () => { paneFor = false; };
    app.taskPane.show(paneSpec());
  }

  function paneSpec(): PaneSpec {
    const s = () => selected!;
    const line = isLinePreset(s().geom) || !!s().connector;
    const fillSec: PaneSection = { title: "Füllung", open: true, controls: () => {
      const f = s().fill;
      const c: PaneControl[] = [
        { type: "radio", label: "Keine Füllung", checked: () => f?.type === "none", pick: () => edit("Füllung", (q) => ({ ...q, fill: { type: "none" } })) },
        { type: "radio", label: "Einfarbige Füllung", checked: () => f?.type === "solid", pick: () => edit("Füllung", (q) => ({ ...q, fill: { type: "solid", color: f?.color ?? accent() } })) },
        { type: "radio", label: "Automatisch", checked: () => !f, pick: () => edit("Füllung", (q) => ({ ...q, fill: undefined })) },
      ];
      if (f?.type === "solid") {
        c.push({ type: "color", label: "Farbe", value: () => f.color ?? null, set: (hex) => edit("Füllung", (q) => ({ ...q, fill: { ...f, color: hex ?? accent() } })) });
        c.push({ type: "number", label: "Transparenz", unit: " %", min: 0, max: 100, slider: true, value: () => f.transparency ?? 0, set: (v) => edit("Füllung", (q) => ({ ...q, fill: { ...f, transparency: v || undefined } })) });
      }
      return c;
    } };
    const lineSec: PaneSection = { title: "Linie", open: line, controls: () => {
      const l = s().line;
      const set = (patch: Partial<NonNullable<SheetShape["line"]>> | null) => edit("Linie", (q) => ({ ...q, line: patch === null ? undefined : { type: "solid", ...q.line, ...patch } as SheetShape["line"] }));
      const c: PaneControl[] = [
        { type: "radio", label: "Keine Linie", checked: () => l?.type === "none", pick: () => edit("Linie", (q) => ({ ...q, line: { type: "none" } })) },
        { type: "radio", label: "Einfarbige Linie", checked: () => l?.type === "solid", pick: () => set({ type: "solid", color: l?.color ?? "#2F528F" }) },
        { type: "radio", label: "Automatisch", checked: () => !l, pick: () => set(null) },
      ];
      if (l?.type !== "none") {
        c.push({ type: "color", label: "Farbe", value: () => l?.color ?? null, set: (hex) => set({ type: "solid", color: hex ?? undefined }) });
        c.push({ type: "number", label: "Breite", unit: " Pt.", min: 0, max: 1584, step: 0.25, decimals: 2, value: () => l?.width ?? (s().textBox ? 0.75 : app.wb.themeLines[(s().style?.ln ?? (line ? 1 : 2)) - 1] ?? 1), set: (v) => set({ width: v }) });
        c.push({ type: "select", label: "Strichtyp", options: [["solid", "Durchgehend"], ["sysDot", "Rundpunkt"], ["sysDash", "Quadratpunkt"], ["dash", "Strich"], ["dashDot", "Strich-Punkt"], ["lgDash", "Langer Strich"]].map(([value, label]) => ({ value, label })),
          value: () => l?.dash ?? "solid", set: (v) => set({ dash: v === "solid" ? undefined : v }) });
        if (line) {
          const ends = [["none", "Keine"], ["triangle", "Pfeil"], ["stealth", "Offener Pfeil"], ["oval", "Oval"], ["diamond", "Raute"]].map(([value, label]) => ({ value, label }));
          c.push({ type: "select", label: "Anfangspfeiltyp", options: ends, value: () => l?.head ?? "none", set: (v) => set({ head: v === "none" ? undefined : v }) });
          c.push({ type: "select", label: "Endpfeiltyp", options: ends, value: () => l?.tail ?? "none", set: (v) => set({ tail: v === "none" ? undefined : v }) });
        }
      }
      return c;
    } };
    const sizeSec: PaneSection = { title: "Größe", open: true, controls: () => {
      const cm = (px: number) => Math.round((px / 96) * 2.54 * 100) / 100;
      const resize = (w: number, hh: number) => { const b = boxOf(s()); if (b) edit("Größe ändern", (q) => placeAt(q, { x: b.x, y: b.y, w: (w / 2.54) * 96 * z(), h: (hh / 2.54) * 96 * z() })); };
      const size = () => { const b = boxOf(s()); return { w: b ? cm(b.w / z()) : 0, h: b ? cm(b.h / z()) : 0 }; };
      return [
        { type: "number", label: "Höhe", unit: " cm", min: 0, max: 558, step: 0.1, decimals: 2, value: () => size().h, set: (v) => resize(size().w, v) },
        { type: "number", label: "Breite", unit: " cm", min: 0, max: 558, step: 0.1, decimals: 2, value: () => size().w, set: (v) => resize(v, size().h) },
        { type: "number", label: "Drehung", unit: "°", min: -3600, max: 3600, value: () => s().rot ?? 0, set: (v) => edit("Drehen", (q) => ({ ...q, rot: ((Math.round(v) % 360) + 360) % 360 || undefined })) },
      ];
    } };
    const textSec: PaneSection = { title: "Textfeld", open: true, controls: () => [
      { type: "select", label: "Vertikale Ausrichtung", options: [{ value: "t", label: "Oben" }, { value: "ctr", label: "Mitte" }, { value: "b", label: "Unten" }],
        value: () => s().anchor ?? "t", set: (v) => edit("Ausrichtung", (q) => ({ ...q, anchor: v as SheetShape["anchor"] })) },
      { type: "select", label: "Horizontale Ausrichtung", options: [{ value: "l", label: "Links" }, { value: "ctr", label: "Zentriert" }, { value: "r", label: "Rechts" }],
        value: () => s().align ?? "l", set: (v) => edit("Ausrichtung", (q) => ({ ...q, align: v as SheetShape["align"] })) },
    ] };
    const tabs = [{ id: "fill", label: "Füllung und Linie", icon: "pane-fill-line", sections: () => (line ? [lineSec] : [fillSec, lineSec]) },
      { id: "size", label: "Größe und Eigenschaften", icon: "size-position", sections: () => (line ? [sizeSec] : [sizeSec, textSec]) }];
    return { title: line ? "Form formatieren" : s().textBox ? "Form formatieren" : "Form formatieren", sets: [{ label: "Formoptionen", tabs },
      ...(line ? [] : [{ label: "Textoptionen", tabs: [{ id: "text", label: "Textfüllung und Kontur", icon: "font-color", sections: () => [{ title: "Text", open: true, controls: (): PaneControl[] => [
        { type: "color", label: "Farbe", value: () => s().font?.color ?? null, set: (hex) => edit("Schriftfarbe", (q) => ({ ...q, font: { ...q.font, color: hex ?? undefined } })) },
        { type: "number", label: "Schriftgrad", unit: " Pt.", min: 1, max: 409, value: () => s().font?.size ?? 11, set: (v) => edit("Schriftgrad", (q) => ({ ...q, font: { ...q.font, size: v } })) },
        { type: "check", label: "Fett", value: () => !!s().font?.bold, set: (on) => edit("Fett", (q) => ({ ...q, font: { ...q.font, bold: on || undefined } })) },
        { type: "check", label: "Kursiv", value: () => !!s().font?.italic, set: (on) => edit("Kursiv", (q) => ({ ...q, font: { ...q.font, italic: on || undefined } })) },
      ] }] }] }]),
    ] };
  }

  // ----- Formformat commands ------------------------------------------------------------------------

  const arrange = {
    forward: () => { const s = selected; if (!s) return; change("Ebene nach vorne", (list) => { const i = list.indexOf(s); if (i < list.length - 1) [list[i], list[i + 1]] = [list[i + 1], list[i]]; return list; }); shown = undefined; },
    backward: () => { const s = selected; if (!s) return; change("Ebene nach hinten", (list) => { const i = list.indexOf(s); if (i > 0) [list[i], list[i - 1]] = [list[i - 1], list[i]]; return list; }); shown = undefined; },
    rotate: (deg: number) => edit("Drehen", (q) => ({ ...q, rot: (((q.rot ?? 0) + deg) % 360 + 360) % 360 || undefined })),
    flip: (k: "flipH" | "flipV") => edit("Spiegeln", (q) => ({ ...q, [k]: q[k] ? undefined : true })),
  };

  const rotateMenu: MenuEntry[] = [
    { label: "Rechtsdrehung 90°", run: () => arrange.rotate(90) }, { label: "Linksdrehung 90°", run: () => arrange.rotate(-90) },
    { label: "Vertikal kippen", run: () => arrange.flip("flipV") }, { label: "Horizontal kippen", run: () => arrange.flip("flipH") },
  ];

  return {
    place, startDraw, gallery, openPane, arrange, rotateMenu,
    selected: (): SheetShape | null => selected,
    deselect: (): void => { if (selected) select(null); },
    setFill: (hex: string | null) => edit("Fülleffekt", (q) => ({ ...q, fill: hex === null ? { type: "none" } : { type: "solid", color: hex } })),
    setLine: (hex: string | null) => edit("Formkontur", (q) => ({ ...q, line: hex === null ? { type: "none" } : { type: "solid", ...q.line, color: hex } })),
  };
}
