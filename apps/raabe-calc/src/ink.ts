// The Zeichnen tab (D-093): ink drawn with the mouse, pen or finger on the sheet, stored as Excel stores ink
// (contentPart with an InkML part, ooxml/xlsx/ink.ts; recorded with probe_excel_ink.py). The tab's controls,
// groups and keytips are Excel 365's (captures/excel-ui/tab-Zeichnen.txt, inventory 2026-11-05); the pens of
// the toolbox and their colours and widths are Excel's defaults as Office keeps them for Excel
// (InkToolboxPensExcelV2: Stift Schwarz 0,35 mm, Stift Rot #E71224 0,35 mm, Galaxie #AE198D 0,5 mm, Bleistift
// #849398 0,5 mm, Textmarker #FFFC00 6 mm, Füllfederhalter #004F8B 1 mm, Pinselstift #008C3A 1 mm).
// Clicking the chosen pen again opens its menu (Stärke and colours; not recorded, our own lists of widths and
// Office's ink colours). Each stroke is one ink drawing ("Freihand N"). Radierer erases whole strokes,
// Lassoauswahl selects the drawings inside a loop, Auswählen returns to the cells (a click on ink selects it,
// dragging moves it, Entf deletes it). Freihand in Form turns a stroke that looks like a line, triangle,
// rectangle, rhombus or ellipse into that shape (the recognition shared with RaabePresent, D-072).
// Wiedergabe in Freihand draws the sheet's ink again stroke by stroke in its recorded timing.

import { h, showMenu, type Control, type MenuEntry, type Tab } from "@raabe/shell";
import type { InkBrush, SheetInk, SheetShape } from "@raabe/sheet-engine";
import { recognize } from "@raabe/text-engine";
import type { CalcApp } from "./app";

type Kind = "pen" | "pencil" | "highlighter" | "fountain" | "brush" | "galaxy" | "action";
interface Pen { kind: Kind; name: string; color: string; width: number }
type Tool = { t: "select" } | { t: "lasso" } | { t: "eraser" } | { t: "pen"; i: number } | { t: "added"; pen: Pen };

/** Excel's toolbox (Office's settings for Excel, InkToolboxPensExcelV2; thickness in 1/100 mm). */
const DEFAULT_PENS: Pen[] = [
  { kind: "pen", name: "Stift", color: "#000000", width: 0.35 },
  { kind: "pen", name: "Stift", color: "#E71224", width: 0.35 },
  { kind: "galaxy", name: "Stift", color: "#AE198D", width: 0.5 },
  { kind: "pencil", name: "Bleistift", color: "#849398", width: 0.5 },
  { kind: "highlighter", name: "Textmarker", color: "#FFFC00", width: 6 },
  { kind: "fountain", name: "Füllfederhalter", color: "#004F8B", width: 1 },
  { kind: "brush", name: "Pinselstift", color: "#008C3A", width: 1 },
  { kind: "action", name: "Aktionsstift", color: "#000000", width: 0.25 },
];

/** Ink colours: the toolbox's and Office's highlighter colours (InkToolboxPens, FavoritePens); the list itself is ours. */
const COLORS: [string, string][] = [
  ["Schwarz", "#000000"], ["Weiß", "#FFFFFF"], ["Grau", "#849398"], ["Rot", "#E71224"], ["Orange", "#FF8517"], ["Gelb", "#FFFC00"],
  ["Grün", "#008C3A"], ["Hellgrün", "#00F900"], ["Türkis", "#00FDFF"], ["Blau", "#3165BB"], ["Dunkelblau", "#004F8B"], ["Violett", "#AE198D"],
  ["Pink", "#FF40FF"], ["Braun", "#7E5228"],
];
const WIDTHS: Record<string, number[]> = { highlighter: [2, 4, 6, 8, 10], default: [0.25, 0.35, 0.5, 0.7, 1, 2, 3.5] };

/** Excel writes the widths with a point ("0.35 mm"). */
const fmt = (mm: number) => String(mm);
const colorName = (c: string) => COLORS.find((x) => x[1].toUpperCase() === c.toUpperCase())?.[0] ?? c;
/** The tile's label as Excel names its pens ("Stift: Schwarz, 0.35 mm", Excel writes the point). */
function penLabel(p: Pen): string {
  if (p.kind === "action") return "Aktionsstift";
  if (p.kind === "galaxy") return `Stift: Galaxie, ${fmt(p.width)} mm`;
  if (p.kind === "highlighter") return `Textmarker: ${colorName(p.color)}, ${fmt(p.width)} mm, an Text ausrichten: Aus`;
  return `${p.name}: ${colorName(p.color)}, ${fmt(p.width)} mm`;
}

const MM_PX = 96 / 25.4;

const CSS = `
.ink-layer { position: absolute; left: 0; top: 0; right: 17px; bottom: 0; overflow: hidden; pointer-events: none; z-index: 3; }
.ink-layer svg { position: absolute; overflow: visible; }
.ink-layer path { pointer-events: visibleStroke; cursor: move; }
.ink-layer svg.on { outline: 1px dashed #7a7a7a; outline-offset: 2px; }
.ink-draw { position: absolute; left: 0; top: 0; right: 17px; bottom: 0; z-index: 7; touch-action: none; }
.ink-draw.pen { cursor: crosshair; }
.ink-draw.eraser { cursor: cell; }
.ink-draw svg { position: absolute; left: 0; top: 0; width: 100%; height: 100%; overflow: visible; pointer-events: none; }
.pen-tile { width: 40px; height: 52px; display: block; }
`;

export function inkFeature(app: CalcApp) {
  const pens = DEFAULT_PENS.map((p) => ({ ...p }));
  let tool: Tool = { t: "select" };
  let toShape = false;
  let selected: SheetInk[] = [];
  const layer = h("div", { class: "ink-layer" });
  let over: HTMLElement | null = null;
  let shown: SheetInk[] | undefined;
  let hiddenWhileErasing = new Set<SheetInk>();
  let replaying = false;
  const z = () => app.sheet.view.zoom / 100;
  const css = () => { if (!document.getElementById("raabe-ink-css")) document.head.append(h("style", { id: "raabe-ink-css" }, CSS)); };

  const currentPen = (): Pen | null => (tool.t === "pen" ? pens[tool.i] : tool.t === "added" ? tool.pen : null);

  // ----- places ----------------------------------------------------------------------------------

  function anchorAt(x: number, y: number): SheetInk["from"] {
    const hit = app.grid.hitCell(x, y);
    const c = app.grid.rectOnCanvas({ r0: hit.row, c0: hit.col, r1: hit.row, c1: hit.col })!;
    return { col: hit.col, row: hit.row, x: Math.max(0, (x - c.x) / z()), y: Math.max(0, (y - c.y) / z()) };
  }

  function boxOf(k: SheetInk): { x: number; y: number; w: number; h: number } | null {
    const a = app.grid.rectOnCanvas({ r0: k.from.row, c0: k.from.col, r1: k.from.row, c1: k.from.col });
    if (!a) return null;
    return { x: a.x + k.from.x * z(), y: a.y + k.from.y * z(), w: k.size.w * z(), h: k.size.h * z() };
  }

  // ----- drawing ink -------------------------------------------------------------------------------

  /** A stroke as an SVG path in px of its box at 100 %. */
  function pathD(pts: { x: number; y: number }[]): string {
    if (!pts.length) return "";
    if (pts.length === 1) return `M${pts[0].x} ${pts[0].y}l0.01 0`;
    let d = `M${pts[0].x.toFixed(2)} ${pts[0].y.toFixed(2)}`;
    for (let i = 1; i < pts.length - 1; i++) {
      const mx = (pts[i].x + pts[i + 1].x) / 2, my = (pts[i].y + pts[i + 1].y) / 2;
      d += ` Q${pts[i].x.toFixed(2)} ${pts[i].y.toFixed(2)} ${mx.toFixed(2)} ${my.toFixed(2)}`;
    }
    const l = pts[pts.length - 1];
    return d + ` L${l.x.toFixed(2)} ${l.y.toFixed(2)}`;
  }

  function strokeAttrs(b: InkBrush, scale: number): Record<string, string> {
    const hl = b.kind === "highlighter";
    return {
      fill: "none", stroke: b.color, "stroke-width": String(Math.max(0.5, b.width * MM_PX * scale)), "stroke-linecap": hl ? "square" : "round", "stroke-linejoin": "round",
      ...(hl ? { "stroke-opacity": "0.5", style: "mix-blend-mode:multiply" } : b.kind === "pencil" ? { "stroke-opacity": "0.85" } : {}),
    };
  }

  function svgOf(k: SheetInk, scale: number, upTo?: number): SVGSVGElement {
    const svg = document.createElementNS("http://www.w3.org/2000/svg", "svg");
    k.strokes.forEach((s, i) => {
      if (upTo !== undefined && i > upTo) return;
      const p = document.createElementNS("http://www.w3.org/2000/svg", "path");
      p.setAttribute("d", pathD(s.pts.map((q) => ({ x: q.x * scale, y: q.y * scale }))));
      for (const [a, v] of Object.entries(strokeAttrs(s.brush, scale))) p.setAttribute(a, v);
      svg.append(p);
    });
    return svg;
  }

  /** Called after every paint of the grid: the sheet's ink in place. */
  function place(): void {
    const list = app.sheet.inks;
    if (!list?.length) { if (layer.isConnected) layer.remove(); shown = undefined; selected = []; return; }
    css();
    if (!layer.isConnected) {
      const canvas = app.grid.el.querySelector(".grid-canvas");
      app.grid.el.insertBefore(layer, canvas ? canvas.nextSibling : app.grid.el.firstChild);
    }
    if (shown !== list) {
      selected = selected.filter((k) => list.includes(k));
      layer.replaceChildren();
      for (const k of list) {
        const svg = svgOf(k, 1);
        svg.addEventListener("pointerdown", (e) => grab(e, k));
        layer.append(svg);
      }
      shown = list;
    }
    list.forEach((k, i) => {
      const svg = layer.children[i] as SVGSVGElement | undefined;
      if (!svg) return;
      const b = boxOf(k);
      if (!b || replaying || hiddenWhileErasing.has(k)) { svg.style.display = "none"; return; }
      const pad = 12 * z();
      Object.assign(svg.style, { display: "", left: `${b.x}px`, top: `${b.y}px`, width: `${Math.max(1, b.w)}px`, height: `${Math.max(1, b.h)}px` });
      const key = String(z());
      if (svg.dataset.z !== key) {
        svg.replaceChildren(...[...svgOf(k, z()).childNodes]);
        svg.dataset.z = key;
      }
      void pad;
      svg.classList.toggle("on", selected.includes(k));
      // Ink can be picked with Auswählen (the cells' tool), not while a pen draws.
      svg.style.pointerEvents = tool.t === "select" ? "" : "none";
    });
  }

  /** A click on ink: selects it; dragging moves the selected drawings (one undo step). */
  function grab(e: PointerEvent, k: SheetInk): void {
    if (tool.t !== "select" || e.button !== 0) return;
    e.preventDefault();
    e.stopPropagation();
    if (!selected.includes(k)) selected = e.shiftKey || e.ctrlKey ? [...selected, k] : [k];
    app.grid.focus();
    const sx = e.clientX, sy = e.clientY;
    const el = e.currentTarget as SVGSVGElement;
    el.setPointerCapture(e.pointerId);
    let dx = 0, dy = 0;
    const move = (ev: PointerEvent) => {
      dx = ev.clientX - sx; dy = ev.clientY - sy;
      for (const q of selected) {
        const svg = layer.children[app.sheet.inks!.indexOf(q)] as SVGSVGElement | undefined;
        if (svg) svg.style.transform = `translate(${dx}px, ${dy}px)`;
      }
    };
    const up = () => {
      el.removeEventListener("pointermove", move);
      el.removeEventListener("pointerup", up);
      for (const svg of layer.children) (svg as SVGSVGElement).style.transform = "";
      if (Math.hypot(dx, dy) > 2) {
        const moved = new Map(selected.map((q) => {
          const b = boxOf(q)!;
          const from = anchorAt(b.x + dx, b.y + dy), to = anchorAt(b.x + dx + b.w, b.y + dy + b.h);
          return [q, { ...q, from, to }] as const;
        }));
        change("Verschieben", (list) => list.map((q) => moved.get(q) ?? q));
        selected = [...moved.values()];
      }
      place();
      app.updateChrome();
    };
    el.addEventListener("pointermove", move);
    el.addEventListener("pointerup", up);
    place();
    app.updateChrome();
  }

  function change(label: string, fn: (list: SheetInk[]) => SheetInk[]): void {
    const sheet = app.sheet;
    app.change(label, () => {
      app.history.whole(sheet);
      const next = fn([...(sheet.inks ?? [])]);
      sheet.inks = next.length ? next : undefined;
    });
    shown = undefined;
    place();
  }

  /** The picture Excel keeps beside the ink for programs without it (mc:Fallback). */
  function fallbackPng(k: SheetInk): string | undefined {
    const pad = Math.max(...k.strokes.map((s) => s.brush.width * MM_PX / 2), 1);
    const c = document.createElement("canvas");
    c.width = Math.ceil(k.size.w + 2 * pad);
    c.height = Math.ceil(k.size.h + 2 * pad);
    const g = c.getContext("2d");
    if (!g) return undefined;
    for (const s of k.strokes) {
      g.strokeStyle = s.brush.color;
      g.globalAlpha = s.brush.kind === "highlighter" ? 0.5 : 1;
      g.lineWidth = Math.max(0.5, s.brush.width * MM_PX);
      g.lineCap = s.brush.kind === "highlighter" ? "square" : "round";
      g.lineJoin = "round";
      g.stroke(new Path2D(pathD(s.pts.map((p) => ({ x: p.x + pad, y: p.y + pad })))));
    }
    try { return c.toDataURL("image/png"); } catch { return undefined; }
  }

  function brushOf(p: Pen): InkBrush {
    const kind = p.kind === "highlighter" ? "highlighter" : p.kind === "pencil" ? "pencil" : "pen";
    return { color: p.color.toUpperCase(), width: p.width, kind, ...(p.kind === "pencil" ? { extra: { inkEffects: "pencil" } } : p.kind === "galaxy" ? { extra: { inkEffects: "galaxy" } } : {}) };
  }

  /** A finished stroke (points in CSS px of the grid, t in ms) becomes an ink drawing (or a shape). */
  function addStroke(pts: { x: number; y: number; t: number }[], pen: Pen): void {
    if (!pts.length) return;
    const xs = pts.map((p) => p.x), ys = pts.map((p) => p.y);
    const x0 = Math.min(...xs), y0 = Math.min(...ys), w = Math.max(...xs) - x0, hh = Math.max(...ys) - y0;
    if (toShape) {
      const r = recognize(pts.map((p) => ({ x: p.x / z() * 9525, y: p.y / z() * 9525 })));
      if (r) { addShape(r, pen); return; }
    }
    const n = (app.sheet.inks?.length ?? 0) + (app.sheet.shapes?.length ?? 0) + (app.sheet.pictures?.length ?? 0) + 1;
    const ink: SheetInk = {
      from: anchorAt(x0, y0), to: anchorAt(x0 + w, y0 + hh), size: { w: Math.max(w / z(), 0.01), h: Math.max(hh / z(), 0.01) },
      strokes: [{ brush: brushOf(pen), pts: pts.map((p) => ({ x: (p.x - x0) / z(), y: (p.y - y0) / z(), t: Math.round(p.t) })) }],
      name: `Freihand ${n}`, time: new Date().toISOString(),
    };
    ink.fallback = fallbackPng(ink);
    change("Freihand", (list) => [...list, ink]);
  }

  function addShape(r: NonNullable<ReturnType<typeof recognize>>, pen: Pen): void {
    const px = (v: number) => (v / 9525) * z();
    const x = px(r.x), y = px(r.y), w = Math.max(1, px(r.w)), hh = Math.max(1, px(r.h));
    const n = (app.sheet.shapes?.length ?? 0) + 1;
    const s: SheetShape = {
      from: anchorAt(x, y), to: anchorAt(x + w, y + hh), geom: r.prst, size: { w: w / z(), h: hh / z() }, name: `Freihand ${n}`,
      fill: { type: "none" }, line: { type: "solid", color: pen.color, width: Math.max(0.25, pen.width * 72 / 25.4) },
      ...(r.flipV ? { flipV: true } : {}), ...(r.prst === "line" ? { connector: true } : {}),
    };
    const sheet = app.sheet;
    app.change("Freihand in Form", () => {
      app.history.whole(sheet);
      sheet.shapes = [...(sheet.shapes ?? []), s];
    });
  }

  /** The overlay that takes the pointer while a pen, the eraser or the lasso is chosen. */
  function overlay(): void {
    over?.remove();
    over = null;
    if (tool.t === "select") return;
    css();
    over = h("div", { class: `ink-draw ${tool.t === "eraser" ? "eraser" : "pen"}` });
    const svg = document.createElementNS("http://www.w3.org/2000/svg", "svg");
    over.append(svg);
    app.grid.el.append(over);
    over.addEventListener("pointerdown", (e) => down(e, svg));
    over.addEventListener("wheel", (e) => {
      // Scrolling while drawing: hand the wheel to the grid below.
      over!.style.pointerEvents = "none";
      const below = document.elementFromPoint(e.clientX, e.clientY);
      over!.style.pointerEvents = "";
      below?.dispatchEvent(new WheelEvent("wheel", e));
    }, { passive: true });
  }

  function down(e: PointerEvent, svg: SVGSVGElement): void {
    if (e.button !== 0 || !over) return;
    e.preventDefault();
    over.setPointerCapture(e.pointerId);
    const r0 = app.grid.el.getBoundingClientRect();
    const at = (ev: PointerEvent) => ({ x: ev.clientX - r0.left, y: ev.clientY - r0.top });
    const t0 = performance.now();
    const pts = [{ ...at(e), t: 0 }];
    const pen = currentPen();
    const live = document.createElementNS("http://www.w3.org/2000/svg", "path");
    if (pen) for (const [a, v] of Object.entries(strokeAttrs(brushOf(pen), z()))) live.setAttribute(a, v);
    else { live.setAttribute("fill", "none"); live.setAttribute("stroke", "#6f6f6f"); live.setAttribute("stroke-dasharray", "4 3"); }
    svg.replaceChildren(live);
    const erased = new Set<SheetInk>();
    const move = (ev: PointerEvent) => {
      const coalesced = typeof ev.getCoalescedEvents === "function" ? ev.getCoalescedEvents() : [];
      for (const c of coalesced.length ? coalesced : [ev]) {
        const p = at(c);
        const last = pts[pts.length - 1];
        if (Math.hypot(p.x - last.x, p.y - last.y) < 0.75) continue;
        pts.push({ ...p, t: performance.now() - t0 });
        if (tool.t === "eraser") eraseAt(p.x, p.y, erased);
      }
      if (tool.t !== "eraser") live.setAttribute("d", pathD(pts));
    };
    const up = () => {
      over?.removeEventListener("pointermove", move);
      over?.removeEventListener("pointerup", up);
      svg.replaceChildren();
      if (tool.t === "eraser") {
        eraseAt(pts[0].x, pts[0].y, erased);
        hiddenWhileErasing = new Set();
        if (erased.size) change("Radieren", (list) => list.filter((k) => !erased.has(k)));
        else place();
      } else if (tool.t === "lasso") lasso(pts);
      else if (pen) addStroke(pts, pen);
    };
    over.addEventListener("pointermove", move);
    over.addEventListener("pointerup", up);
  }

  /** Radierer (Strichradierer): every drawing whose stroke passes under the pointer goes. */
  function eraseAt(x: number, y: number, erased: Set<SheetInk>): void {
    for (const k of app.sheet.inks ?? []) {
      if (erased.has(k)) continue;
      const b = boxOf(k);
      if (!b) continue;
      const tol = 4 + Math.max(...k.strokes.map((s) => s.brush.width * MM_PX * z() / 2));
      if (x < b.x - tol || x > b.x + b.w + tol || y < b.y - tol || y > b.y + b.h + tol) continue;
      const hit = k.strokes.some((s) => s.pts.some((p, i) => {
        const q = s.pts[i + 1] ?? p;
        const ax = b.x + p.x * z(), ay = b.y + p.y * z(), bx = b.x + q.x * z(), by = b.y + q.y * z();
        const l2 = (bx - ax) ** 2 + (by - ay) ** 2;
        const t = l2 ? Math.max(0, Math.min(1, ((x - ax) * (bx - ax) + (y - ay) * (by - ay)) / l2)) : 0;
        return Math.hypot(x - (ax + t * (bx - ax)), y - (ay + t * (by - ay))) <= tol;
      }));
      if (hit) { erased.add(k); hiddenWhileErasing.add(k); }
    }
    if (erased.size) place();
  }

  /** Lassoauswahl: the drawings whose middle lies inside the loop are selected. */
  function lasso(pts: { x: number; y: number }[]): void {
    const inside = (x: number, y: number) => {
      let c = false;
      for (let i = 0, j = pts.length - 1; i < pts.length; j = i++) {
        const a = pts[i], b = pts[j];
        if ((a.y > y) !== (b.y > y) && x < ((b.x - a.x) * (y - a.y)) / (b.y - a.y) + a.x) c = !c;
      }
      return c;
    };
    selected = (app.sheet.inks ?? []).filter((k) => { const b = boxOf(k); return !!b && inside(b.x + b.w / 2, b.y + b.h / 2); });
    place();
    app.updateChrome();
  }

  function setTool(t: Tool): void {
    tool = t;
    if (t.t !== "select" && t.t !== "lasso") selected = [];
    overlay();
    place();
    app.updateChrome();
  }

  /** The chosen pen's menu: Stärke and Farben (clicking the chosen pen again, as in Office). */
  function penMenu(i: number, anchor: DOMRect): void {
    const p = pens[i];
    const widths = WIDTHS[p.kind === "highlighter" ? "highlighter" : "default"];
    const entries: MenuEntry[] = [
      { label: "Stärke", submenu: () => widths.map((w) => ({ label: `${fmt(w)} mm`, checked: p.width === w, run: () => { p.width = w; app.updateChrome(); } })) },
      "-",
      ...COLORS.map(([name, c]) => ({ label: name, checked: p.color.toUpperCase() === c, run: () => { p.color = c; app.updateChrome(); } })),
    ];
    showMenu(app.commands, entries, anchor);
  }

  // ----- Wiedergabe in Freihand ---------------------------------------------------------------------

  /** Draws the sheet's ink again, stroke after stroke, at the speed it was drawn (gaps between strokes shortened). */
  function replay(): void {
    const list = app.sheet.inks ?? [];
    if (!list.length || replaying) return;
    css();
    replaying = true;
    place();
    const play = h("div", { class: "ink-layer" });
    app.grid.el.append(play);
    const items: { el: SVGPathElement; len: number; dur: number; start: number }[] = [];
    let clock = 0;
    for (const k of list) {
      const b = boxOf(k);
      if (!b) continue;
      const svg = svgOf(k, z());
      Object.assign(svg.style, { left: `${b.x}px`, top: `${b.y}px`, width: `${b.w}px`, height: `${b.h}px` });
      play.append(svg);
      [...svg.querySelectorAll("path")].forEach((path, i) => {
        const s = k.strokes[i];
        const ts = s.pts.map((p) => p.t ?? 0);
        const dur = Math.min(3000, Math.max(120, ts[ts.length - 1] - ts[0] || s.pts.length * 8));
        const len = path.getTotalLength?.() ?? 1;
        path.style.strokeDasharray = `${len} ${len}`;
        path.style.strokeDashoffset = String(len);
        items.push({ el: path, len, dur, start: clock });
        clock += dur + 150;
      });
    }
    const t0 = performance.now();
    const step = () => {
      const t = performance.now() - t0;
      for (const it of items) it.el.style.strokeDashoffset = String(it.len * (1 - Math.max(0, Math.min(1, (t - it.start) / it.dur))));
      if (t < clock) requestAnimationFrame(step);
      else { play.remove(); replaying = false; place(); }
    };
    requestAnimationFrame(step);
  }

  // ----- keys, tab, commands ----------------------------------------------------------------------------

  addEventListener("keydown", (e) => {
    if (e.key === "Escape" && tool.t !== "select") { e.preventDefault(); setTool({ t: "select" }); return; }
    if (!selected.length || app.grid.isEditing) return;
    if ((e.target as HTMLElement).closest?.("input, textarea:not(.cell-editor), .dialog-layer, .taskpane")) return;
    if (e.key === "Delete" || e.key === "Backspace") {
      e.preventDefault();
      e.stopPropagation();
      const gone = new Set(selected);
      selected = [];
      change("Löschen", (list) => list.filter((k) => !gone.has(k)));
    } else if (e.key === "Escape") { selected = []; place(); }
  }, true);

  /** The pen tile: the pen in its colour (our own drawing). */
  function penTile(p: Pen | null, kind: "select" | "lasso" | "eraser" | null): HTMLElement {
    const svg = (inner: string) => { const s = h("span", { class: "pen-tile" }); s.innerHTML = `<svg viewBox="0 0 40 52" width="40" height="52" aria-hidden="true">${inner}</svg>`; return s; };
    if (kind === "select") return svg('<path d="M12 10 L12 38 L19 31 L24 42 L28 40 L23 29 L32 29 Z" fill="#fff" stroke="#333" stroke-width="1.5" stroke-linejoin="round"/>');
    if (kind === "lasso") return svg('<ellipse cx="20" cy="22" rx="13" ry="10" fill="none" stroke="#333" stroke-width="1.5" stroke-dasharray="3 3"/><path d="M14 31 q-2 6 3 9 q4 2 8-1" fill="none" stroke="#333" stroke-width="1.5"/>');
    if (kind === "eraser") return svg('<rect x="13" y="6" width="14" height="34" rx="2" fill="#e8e8e8" stroke="#555"/><rect x="13" y="32" width="14" height="10" rx="2" fill="#d9a0c8" stroke="#555"/>');
    const c = p!.color;
    const body = p!.kind === "highlighter"
      ? `<rect x="12" y="4" width="16" height="30" rx="2" fill="${c}" stroke="#555"/><path d="M14 34 L26 34 L24 46 L16 42 Z" fill="${c}" stroke="#555"/>`
      : p!.kind === "pencil"
        ? `<rect x="14" y="4" width="12" height="30" fill="${c}" stroke="#555"/><path d="M14 34 L26 34 L20 48 Z" fill="#f3d9b1" stroke="#555"/><path d="M18.5 44.5 L21.5 44.5 L20 48 Z" fill="#333"/>`
        : `<rect x="13" y="4" width="14" height="28" rx="3" fill="${p!.kind === "galaxy" ? "#3b2a6b" : c}" stroke="#555"/><path d="M13 32 L27 32 L22 44 L18 44 Z" fill="#fff" stroke="#555"/><path d="M18 44 L22 44 L20 49 Z" fill="${c}"/>`;
    return svg(body);
  }

  const TILE_TOOLS = 3;
  function tab(): Tab {
    const tiles: Control = {
      type: "tiles", cmd: "ink.tools", key: "G", count: TILE_TOOLS + pens.length, visible: TILE_TOOLS + pens.length, size: [40, 52], cls: "ink-tiles",
      render: (i) => (i === 0 ? penTile(null, "select") : i === 1 ? penTile(null, "lasso") : i === 2 ? penTile(null, "eraser") : penTile(pens[i - TILE_TOOLS], null)),
      tip: (i) => (i === 0 ? "Auswählen" : i === 1 ? "Lassoauswahl" : i === 2 ? "Radierer" : penLabel(pens[i - TILE_TOOLS])),
      pick: (i) => {
        if (i === 0) return setTool({ t: "select" });
        if (i === 1) return setTool({ t: "lasso" });
        if (i === 2) return setTool({ t: "eraser" });
        const k = i - TILE_TOOLS;
        if (tool.t === "pen" && tool.i === k) {
          const b = document.querySelectorAll<HTMLElement>(".ink-tiles .tile-item")[i];
          if (b) penMenu(k, b.getBoundingClientRect());
          return;
        }
        setTool({ t: "pen", i: k });
      },
      current: () => (tool.t === "select" ? 0 : tool.t === "lasso" ? 1 : tool.t === "eraser" ? 2 : tool.t === "pen" ? TILE_TOOLS + tool.i : undefined),
      redraw: () => JSON.stringify(pens),
      panel: () => h("div", { style: "padding:8px;max-width:260px" }, "Die Stifte stehen im Menüband; ein zweiter Klick auf den gewählten Stift öffnet Stärke und Farben."),
    };
    const add: Control = {
      type: "dropdown", cmd: "ink.add", key: "A1", size: "large", menu: [
        { label: "Stift", run: () => setTool({ t: "added", pen: { kind: "pen", name: "Stift", color: "#000000", width: 0.5 } }) },
        { label: "Textmarker", run: () => setTool({ t: "added", pen: { kind: "highlighter", name: "Textmarker", color: "#FFFC00", width: 6 } }) },
        { label: "Füllfederhalter", run: () => setTool({ t: "added", pen: { kind: "fountain", name: "Füllfederhalter", color: "#004F8B", width: 1 } }) },
        { label: "Pinselstift", run: () => setTool({ t: "added", pen: { kind: "brush", name: "Pinselstift", color: "#008C3A", width: 1 } }) },
        { label: "Bleistift", run: () => setTool({ t: "added", pen: { kind: "pencil", name: "Bleistift", color: "#849398", width: 0.5 } }) },
        { label: "Textstift", disabled: true, run: () => {} },
      ],
    };
    return {
      id: "draw", label: "Zeichnen", key: "JI",
      shrink: ["replay", "convert", "tools", "undo"],
      groups: [
        { id: "undo", label: "Rückgängig", key: "Z1", icon: "undo", controls: [{ type: "stack", items: [{ type: "button", cmd: "edit.undo", key: "AZ" }, { type: "button", cmd: "edit.redo", key: "W1" }] }] },
        { id: "tools", label: "Zeichentools", key: "ZD", icon: "pen", controls: [tiles, add] },
        { id: "convert", label: "Konvertieren", key: "ZC", icon: "shapes", controls: [{ type: "button", cmd: "ink.toShape", key: "N", size: "large" }, { type: "button", cmd: "ink.toMath", key: "A2", size: "large" }] },
        { id: "replay", label: "Wiedergeben", key: "Z2", icon: "play", controls: [{ type: "button", cmd: "ink.replay", key: "W2", size: "large" }] },
      ],
    };
  }

  type Run = (id: string, fn: () => void | Promise<void>, enabled?: () => boolean, checked?: () => boolean) => void;
  function implement(run: Run): void {
    run("ink.tools", () => {});
    run("ink.add", () => {});
    run("ink.toShape", () => { toShape = !toShape; app.updateChrome(); }, undefined, () => toShape);
    // Freihand in Gleichung needs handwriting recognition (Office's is a Windows service): greyed.
    run("ink.toMath", () => {}, () => false);
    run("ink.replay", () => replay(), () => !!app.sheet.inks?.length && !replaying);
  }

  /** Test hook: a stroke as if drawn (points in CSS px of the grid, t in ms) with pen `i` of the toolbox. */
  function testStroke(i: number, pts: { x: number; y: number; t: number }[]): void {
    addStroke(pts, pens[i]);
  }

  return { place, tab, implement, setTool, replay, testStroke, get tool() { return tool; }, get selected() { return selected; }, pens };
}
