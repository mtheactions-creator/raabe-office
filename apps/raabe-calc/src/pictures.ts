// Pictures of a sheet (the school logo in a template …) shown where the file anchors them: over
// the cells, following scrolling and zoom. Einfügen > Bilder > Dieses Gerät adds one at the active
// cell at its natural size (as Excel). A picture is drawn with its Bildformat as Excel's file has it
// (D-087: crop, Korrekturen, Farbe, Transparenz, frame shape, Bildrahmen, effects, rotation) by the
// shared renderer RaabeWrite uses (text-engine's pictureElement). A click selects it (frame, eight
// handles and the round turning handle; the corners keep the proportions while Seitenverhältnis
// sperren is on, as for Excel's inserted pictures), dragging moves it, Entf deletes it, the arrow
// keys nudge it, Strg+1 opens Grafik formatieren. Zuschneiden shows the whole picture behind the
// frame with black crop handles. Each change is one undo step; the file writes changed pictures
// with Excel's DrawingML and keeps the rest as they were.

import { h, messageBox } from "@raabe/shell";
import type { SheetPicture } from "@raabe/sheet-engine";
import { bakePicture, pictureElement, type PictureFormat } from "@raabe/text-engine";
import { pictureBakedKey as bakedKey } from "@raabe/ooxml";
import { sheetAxes } from "./grid/metrics";
import type { CalcApp } from "./app";

// SVG is left out: Excel stores it with a PNG copy beside it, which RaabeCalc does not make.
const MIME: Record<string, string> = { png: "image/png", jpg: "image/jpeg", jpeg: "image/jpeg", gif: "image/gif", bmp: "image/bmp" };

export type Box = { x: number; y: number; w: number; h: number };
type Handle = "nw" | "n" | "ne" | "e" | "se" | "s" | "sw" | "w" | "rot";
const HANDLES: Handle[] = ["nw", "n", "ne", "e", "se", "s", "sw", "w", "rot"];
const POS: Record<string, [number, number]> = { nw: [0, 0], n: [0.5, 0], ne: [1, 0], e: [1, 0.5], se: [1, 1], s: [0.5, 1], sw: [0, 1], w: [0, 0.5] };

const CSS = `
.picture-layer { position: absolute; right: 17px; bottom: 0; overflow: hidden; pointer-events: none; z-index: 3; }
.picture-layer .pic { position: absolute; pointer-events: auto; cursor: move; outline: none; }
.picture-layer .pic.on { outline: 1px solid #7a7a7a; }
.picture-layer .pic-look { position: absolute; left: 0; top: 0; transform-origin: 0 0; pointer-events: none; }
.picture-layer .pic-ghost { position: absolute; opacity: 0.35; pointer-events: none; }
.picture-layer .pic-ghost img { width: 100%; height: 100%; display: block; }
.picture-layer .pic-h { position: absolute; width: 8px; height: 8px; margin: -5px 0 0 -5px; border: 1px solid #5f5f5f; border-radius: 50%; background: #fff; display: none; }
.picture-layer .pic.on .pic-h { display: block; }
.picture-layer .pic.cropping .pic-h { display: none; }
.picture-layer .pic-h.rot { width: 12px; height: 12px; margin: -7px 0 0 -7px; cursor: grab; }
.picture-layer .pic-h.nw, .picture-layer .pic-h.se { cursor: nwse-resize; }
.picture-layer .pic-h.ne, .picture-layer .pic-h.sw { cursor: nesw-resize; }
.picture-layer .pic-h.n, .picture-layer .pic-h.s { cursor: ns-resize; }
.picture-layer .pic-h.e, .picture-layer .pic-h.w { cursor: ew-resize; }
.picture-layer .pic-c { position: absolute; background: #000; display: none; z-index: 2; }
.picture-layer .pic.cropping .pic-c { display: block; }
.picture-layer .pic.picking { cursor: crosshair; }
`;

/** Excel anchors a picture turned by 45° to 135° (or 225° to 315°) by its turned box: width and height change places. */
const swapped = (rot: number | undefined) => { const r = (((rot ?? 0) % 360) + 360) % 360; return (r >= 45 && r < 135) || (r >= 225 && r < 315); };

export function picturesFeature(app: CalcApp) {
  const layer = h("div", { class: "picture-layer" });
  let shown: SheetPicture[] | undefined;
  const els: HTMLElement[] = [];
  let selected: SheetPicture | null = null;
  let live: { p: SheetPicture; box: Box; rot?: number } | null = null;
  /** Livevorschau of a gallery entry on the selected picture. */
  let preview: { p: SheetPicture; f: PictureFormat } | null = null;
  /** Zuschneiden: the frame and the whole picture behind it (canvas px). */
  let crop: { p: SheetPicture; frame: Box; full: Box } | null = null;
  /** Transparente Farbe bestimmen: the next click on the picture picks the colour. */
  let picking = false;
  /** Am Raster ausrichten: moving snaps to cell edges. */
  let snapGrid = false;
  const naturals = new Map<string, { w: number; h: number }>();
  const pending = new Set<Promise<void>>();
  const listeners: (() => void)[] = [];

  const z = () => app.sheet.view.zoom / 100;
  const fmt = (p: SheetPicture): PictureFormat => (p.format ?? {}) as PictureFormat;

  function css(): void {
    if (!document.getElementById("raabe-pictures-css")) document.head.append(h("style", { id: "raabe-pictures-css" }, CSS));
  }

  /** The picture's own size in px (its pixels at 96 dpi), once known. */
  function natural(p: SheetPicture): { w: number; h: number } | undefined {
    const n = naturals.get(p.src);
    if (n) return n;
    const img = new Image();
    img.onload = () => { naturals.set(p.src, { w: img.naturalWidth, h: img.naturalHeight }); };
    img.src = p.src;
    return undefined;
  }

  async function naturalAsync(p: SheetPicture): Promise<{ w: number; h: number }> {
    const n = naturals.get(p.src);
    if (n) return n;
    return new Promise((resolve) => {
      const img = new Image();
      img.onload = () => { const s = { w: img.naturalWidth, h: img.naturalHeight }; naturals.set(p.src, s); resolve(s); };
      img.onerror = () => resolve({ w: 100, h: 100 });
      img.src = p.src;
    });
  }

  // ----- places ---------------------------------------------------------------------------------

  /** A point on the canvas (CSS px) as a cell and the offset in it at 100 %. */
  function anchorAt(x: number, y: number): { col: number; row: number; x: number; y: number } {
    const hit = app.grid.hitCell(Math.max(0, x), Math.max(0, y));
    const c = app.grid.rectOnCanvas({ r0: hit.row, c0: hit.col, r1: hit.row, c1: hit.col })!;
    let col = hit.col, row = hit.row, ox = (x - c.x) / z(), oy = (y - c.y) / z();
    // Left of or above the first cell shown (a part scrolled out of view, D-127): counted back through the sheet's
    // columns and rows, so the object stays where it was dropped instead of jumping to the first cell shown.
    if (ox < 0 || oy < 0) {
      const ax = sheetAxes(app.sheet);
      if (ox < 0) { const abs = Math.max(0, ax.cols.pos(col) + ox); col = ax.cols.at(abs); ox = abs - ax.cols.pos(col); }
      if (oy < 0) { const abs = Math.max(0, ax.rows.pos(row) + oy); row = ax.rows.at(abs); oy = abs - ax.rows.pos(row); }
    }
    return { col, row, x: Math.max(0, ox), y: Math.max(0, oy) };
  }

  /** The anchored box on the canvas (CSS px). */
  function anchorBox(p: SheetPicture): Box | null {
    const a = app.grid.rectOnCanvas({ r0: p.from.row, c0: p.from.col, r1: p.from.row, c1: p.from.col });
    if (!a) return null;
    const x = a.x + p.from.x * z(), y = a.y + p.from.y * z();
    let w = (p.size?.w ?? 0) * z(), hh = (p.size?.h ?? 0) * z();
    if (p.to) {
      const b = app.grid.rectOnCanvas({ r0: p.to.row, c0: p.to.col, r1: p.to.row, c1: p.to.col });
      if (b) { w = b.x + p.to.x * z() - x; hh = b.y + p.to.y * z() - y; }
    }
    return { x, y, w, h: hh };
  }

  /** Where a picture's frame is on the canvas (unturned; CSS px). */
  function boxOf(p: SheetPicture): Box | null {
    let b = anchorBox(p);
    // A picture in a group: its share of the group's box.
    if (b && p.group) b = { x: b.x + p.group.x * b.w, y: b.y + p.group.y * b.h, w: p.group.w * b.w, h: p.group.h * b.h };
    if (!b || !swapped(fmt(p).rotation)) return b;
    const cx = b.x + b.w / 2, cy = b.y + b.h / 2;
    return { x: cx - b.h / 2, y: cy - b.w / 2, w: b.h, h: b.w };
  }

  /** The picture with its frame at a box on the canvas (both corners, and its size at 100 %). */
  function placeAt(p: SheetPicture, b: Box, rot = fmt(p).rotation): SheetPicture {
    let a = b;
    if (swapped(rot)) { const cx = b.x + b.w / 2, cy = b.y + b.h / 2; a = { x: cx - b.h / 2, y: cy - b.w / 2, w: b.h, h: b.w }; }
    return { ...p, from: anchorAt(a.x, a.y), to: anchorAt(a.x + a.w, a.y + a.h), size: { w: b.w / z(), h: b.h / z() } };
  }

  // ----- changes --------------------------------------------------------------------------------

  function change(label: string, fn: (list: SheetPicture[]) => SheetPicture[]): void {
    const sheet = app.sheet;
    app.change(label, () => {
      app.history.whole(sheet);
      const next = fn([...(sheet.pictures ?? [])]);
      sheet.pictures = next.length ? next : undefined;
    });
  }

  /**
   * Office keeps Korrekturen, Sättigung and Farbton worked into the stored picture: after a change
   * of them the picture is drawn into a copy with them (in the background; the screen shows them
   * through filters meanwhile), which the file then stores.
   */
  function bake(p: SheetPicture): void {
    const key = bakedKey(fmt(p));
    if (key === (p.baked ?? "") || !key) return;
    const job = bakePicture(p.src, fmt(p)).then((url) => {
      if (!url) return;
      p.baked = key;
      p.bakedSrc = url;
    }).finally(() => pending.delete(job));
    pending.add(job);
  }

  /** A changed picture without effects to keep worked in drops the copy made for them. */
  function normalize(p: SheetPicture): SheetPicture {
    if (!bakedKey(fmt(p)) && p.bakedSrc) { const { baked: _b, bakedSrc: _s, ...rest } = p; void _b; void _s; return rest; }
    return p;
  }

  /** Changes the selected picture (one undo step). */
  function edit(label: string, fn: (p: SheetPicture) => SheetPicture): void {
    const p = selected;
    if (!p) return;
    const next = normalize(fn(p));
    change(label, (list) => list.map((q) => (q === p ? next : q)));
    selected = next;
    bake(next);
    notify();
  }

  /** Changes the selected picture's format: `patch` is merged, undefined keys removed. */
  function setFormat(label: string, patch: PictureFormat | ((f: PictureFormat) => PictureFormat)): void {
    edit(label, (p) => {
      const merged = typeof patch === "function" ? patch(fmt(p)) : { ...fmt(p), ...patch };
      const clean = Object.fromEntries(Object.entries(merged).filter(([, v]) => v !== undefined));
      return { ...p, format: Object.keys(clean).length ? clean : undefined };
    });
  }

  function notify(): void {
    shown = undefined;
    app.grid.contentChanged();
    for (const l of listeners) l();
  }

  // ----- drawing --------------------------------------------------------------------------------

  /** The format as drawn: effects already worked into the stored picture are not applied again. */
  function lookOf(p: SheetPicture, f: PictureFormat): { url: string; f: PictureFormat } {
    const key = bakedKey(f);
    const { rotation: _r, ...rest } = f;
    void _r;
    if (p.baked !== undefined && key === p.baked) {
      return { url: p.bakedSrc ?? p.src, f: { ...rest, brightness: undefined, contrast: undefined, saturation: undefined, temperature: undefined, sharpen: undefined } };
    }
    return { url: p.src, f: rest };
  }

  function formatShown(p: SheetPicture): PictureFormat {
    return preview?.p === p ? preview.f : fmt(p);
  }

  /** Called after every paint of the grid. */
  function place(): void {
    const list = app.sheet.pictures;
    if (!list?.length) { if (layer.isConnected) layer.remove(); shown = undefined; if (selected) { selected = null; app.pictureSelectionChanged(); } return; }
    css();
    if (!layer.isConnected) {
      const canvas = app.grid.el.querySelector(".grid-canvas");
      app.grid.el.insertBefore(layer, canvas ? canvas.nextSibling : app.grid.el.firstChild);
    }
    if (shown !== list) {
      if (selected && !list.includes(selected)) {
        const was = selected;
        selected = list.find((p) => p.raw !== undefined && p.raw === was.raw && p.src === was.src) ?? null;
        if (!selected) queueMicrotask(() => app.pictureSelectionChanged());
      }
      layer.replaceChildren();
      els.length = 0;
      for (const p of list) {
        const el = h("div", { class: "pic", tabindex: "-1" }, h("div", { class: "pic-look" }),
          ...HANDLES.map((k) => h("span", { class: `pic-h ${k}`, "data-h": k })),
          ...HANDLES.filter((k) => k !== "rot").map((k) => h("span", { class: `pic-c c-${k}`, "data-c": k })));
        el.addEventListener("pointerdown", (e) => down(e, p, el));
        els.push(el);
        layer.append(el);
      }
      shown = list;
    }
    const v = app.sheet.view;
    const tl = app.grid.rectOnCanvas({ r0: v.top, c0: v.left, r1: v.top, c1: v.left });
    if (!tl) return;
    Object.assign(layer.style, { left: `${tl.x}px`, top: `${tl.y}px` });
    list.forEach((p, i) => {
      const el = els[i];
      const b0 = boxOf(p);
      if (!b0 || p.hidden) { el.style.display = "none"; return; }
      const f = formatShown(p);
      const cropping = crop?.p === p;
      const b = live?.p === p ? live.box : cropping ? crop!.frame : b0;
      const rot = live?.p === p && live.rot !== undefined ? live.rot : f.rotation ?? 0;
      Object.assign(el.style, { display: b.w > 0 && b.h > 0 ? "" : "none", left: `${b.x - tl.x}px`, top: `${b.y - tl.y}px`, width: `${b.w}px`, height: `${b.h}px`,
        transform: rot && !cropping ? `rotate(${rot}deg)` : "" });
      // Drawn at 100 % and scaled, so that lines and effects keep their size in points.
      const lw = Math.max(0.01, b.w / z()), lh = Math.max(0.01, b.h / z());
      let look = lookOf(p, f);
      if (cropping) look = { ...look, f: { ...look.f, crop: cropOf(crop!.frame, crop!.full) } };
      const key = JSON.stringify([look.url.length, look.url.slice(-40), look.f, Math.round(lw * 4), Math.round(lh * 4)]);
      const sv = el.querySelector<HTMLElement>(".pic-look")!;
      if (sv.dataset.key !== key) {
        sv.replaceChildren(pictureElement(look.url, lw, lh, look.f));
        sv.dataset.key = key;
      }
      sv.style.transform = `scale(${z()})`;
      // The whole picture behind the frame while cropping.
      el.querySelector(".pic-ghost")?.remove();
      if (cropping) {
        const full = crop!.full;
        const g = h("div", { class: "pic-ghost", style: `left:${full.x - b.x}px;top:${full.y - b.y}px;width:${full.w}px;height:${full.h}px` }, h("img", { src: look.url, draggable: "false" }));
        el.prepend(g);
      }
      for (const hd of el.querySelectorAll<HTMLElement>(".pic-h")) {
        const k = hd.dataset.h as Handle;
        if (k === "rot") { hd.style.left = `${b.w / 2}px`; hd.style.top = "-22px"; continue; }
        hd.style.left = `${POS[k][0] * b.w}px`;
        hd.style.top = `${POS[k][1] * b.h}px`;
      }
      for (const c of el.querySelectorAll<HTMLElement>(".pic-c")) {
        const k = c.dataset.c!;
        const [px, py] = POS[k];
        const len = 14, th = 4;
        const corner = k.length === 2;
        const wv = corner ? len : px === 0.5 ? len : th, hv = corner ? len : py === 0.5 ? len : th;
        Object.assign(c.style, { left: `${px * b.w - (px === 1 ? wv - 1 : px === 0.5 ? wv / 2 : 1)}px`, top: `${py * b.h - (py === 1 ? hv - 1 : py === 0.5 ? hv / 2 : 1)}px`, width: `${wv}px`, height: `${hv}px`,
          cursor: k === "n" || k === "s" ? "ns-resize" : k === "e" || k === "w" ? "ew-resize" : k === "nw" || k === "se" ? "nwse-resize" : "nesw-resize" });
        // Corners are L-shaped: drawn as two bars by clip-path.
        if (corner) {
          const t = `${th}px`;
          const poly = k === "nw" ? `polygon(0 0,100% 0,100% ${t},${t} ${t},${t} 100%,0 100%)` : k === "ne" ? `polygon(0 0,100% 0,100% 100%,calc(100% - ${t}) 100%,calc(100% - ${t}) ${t},0 ${t})`
            : k === "sw" ? `polygon(0 0,${t} 0,${t} calc(100% - ${t}),100% calc(100% - ${t}),100% 100%,0 100%)` : `polygon(calc(100% - ${t}) 0,100% 0,100% 100%,0 100%,0 calc(100% - ${t}),calc(100% - ${t}) calc(100% - ${t}))`;
          c.style.clipPath = poly;
        }
      }
      el.classList.toggle("on", p === selected);
      el.classList.toggle("cropping", cropping);
      el.classList.toggle("picking", picking && p === selected);
      void natural(p);
    });
  }

  function select(p: SheetPicture | null): void {
    if (crop && crop.p !== p) endCrop(true);
    picking = false;
    preview = null;
    selected = p;
    shown = undefined;
    app.pictureSelectionChanged();
    app.grid.contentChanged();
    if (!p && paneOpen) { paneOpen = false; app.taskPane.close(); }
    else if (p && paneOpen && paneSpec) app.taskPane.show(paneSpec());
  }

  // ----- mouse ----------------------------------------------------------------------------------

  function down(e: PointerEvent, p: SheetPicture, el: HTMLElement): void {
    if (e.button !== 0 || p.group) return;
    e.preventDefault();
    e.stopPropagation();
    // A protected sheet keeps its pictures where they are unless Objekte bearbeiten is allowed.
    if (!app.protectionAllows("objects")) return;
    if (app.grid.isEditing) app.grid.commitEdit("none");
    if (picking && p === selected) { void pickColor(e, p, el); return; }
    if (selected !== p) select(p);
    // Selecting rebuilds the pictures' elements: the drag goes on with the new one (D-096).
    if (shown === undefined) { place(); el = els[(app.sheet.pictures ?? []).indexOf(p)] ?? el; }
    if (crop?.p === p) { cropDrag(e, el); return; }
    const b = boxOf(p);
    if (!b) return;
    const handle = (e.target as HTMLElement).dataset.h as Handle | undefined;
    const f = fmt(p);
    const sx = e.clientX, sy = e.clientY;
    el.setPointerCapture(e.pointerId);
    const move = (ev: PointerEvent) => {
      const dx = ev.clientX - sx, dy = ev.clientY - sy;
      if (!handle) {
        let nb = { ...b, x: b.x + dx, y: b.y + dy };
        if (snapGrid || ev.altKey) nb = snap(nb);
        live = { p, box: nb };
        place();
        return;
      }
      if (handle === "rot") {
        const cx = b.x + b.w / 2, cy = b.y + b.h / 2;
        const r = app.grid.el.getBoundingClientRect();
        let deg = (Math.atan2(ev.clientY - r.top - cy, ev.clientX - r.left - cx) * 180) / Math.PI + 90;
        if (ev.shiftKey) deg = Math.round(deg / 15) * 15;
        live = { p, box: b, rot: ((deg % 360) + 360) % 360 };
        place();
        return;
      }
      // Sizing in the frame's own (turned) directions.
      const rad = ((f.rotation ?? 0) * Math.PI) / 180;
      const ux = dx * Math.cos(rad) + dy * Math.sin(rad), uy = -dx * Math.sin(rad) + dy * Math.cos(rad);
      let { x, y, w, h: hh } = b;
      if (handle.includes("e")) w += ux;
      if (handle.includes("s")) hh += uy;
      if (handle.includes("w")) { x += ux; w -= ux; }
      if (handle.includes("n")) { y += uy; hh -= uy; }
      // The corners keep the proportions while Seitenverhältnis sperren is on (Excel's inserted pictures), else with Umschalt.
      if (handle.length === 2 && (!!p.lockAspect !== ev.shiftKey)) {
        const k = Math.max(w / b.w, hh / b.h);
        const nw = b.w * k, nh = b.h * k;
        if (handle.includes("w")) x = b.x + b.w - nw;
        if (handle.includes("n")) y = b.y + b.h - nh;
        w = nw; hh = nh;
      }
      w = Math.max(4, w); hh = Math.max(4, hh);
      live = { p, box: { x, y, w, h: hh } };
      place();
    };
    const up = () => {
      el.removeEventListener("pointermove", move);
      el.removeEventListener("pointerup", up);
      el.removeEventListener("pointercancel", up);
      const l = live;
      live = null;
      if (!l) { place(); return; }
      if (l.rot !== undefined) { const deg = Math.round(l.rot) % 360; edit("Drehen", (q) => placeAt({ ...q, format: { ...fmt(q), rotation: deg || undefined } }, b, deg)); return; }
      const same = Math.abs(l.box.x - b.x) < 1 && Math.abs(l.box.y - b.y) < 1 && Math.abs(l.box.w - b.w) < 1 && Math.abs(l.box.h - b.h) < 1;
      if (same) { place(); return; }
      edit(handle ? "Größe ändern" : "Verschieben", (q) => placeAt(q, l.box));
    };
    el.addEventListener("pointermove", move);
    el.addEventListener("pointerup", up);
    el.addEventListener("pointercancel", up);
  }

  /** Am Raster ausrichten: the top-left corner to the nearest cell edges. */
  function snap(b: Box): Box {
    const hit = app.grid.hitCell(Math.max(0, b.x), Math.max(0, b.y));
    const c = app.grid.rectOnCanvas({ r0: hit.row, c0: hit.col, r1: hit.row, c1: hit.col });
    if (!c) return b;
    const x = b.x - c.x < c.w / 2 ? c.x : c.x + c.w, y = b.y - c.y < c.h / 2 ? c.y : c.y + c.h;
    return { ...b, x, y };
  }

  // ----- Zuschneiden ----------------------------------------------------------------------------

  const cropOf = (frame: Box, full: Box) => {
    const r = (v: number) => Math.round(v * 100000) / 100000;
    return { l: r((frame.x - full.x) / full.w), t: r((frame.y - full.y) / full.h), r: r((full.x + full.w - frame.x - frame.w) / full.w), b: r((full.y + full.h - frame.y - frame.h) / full.h) };
  };

  function fullOf(frame: Box, c: PictureFormat["crop"]): Box {
    const k = c ?? { l: 0, t: 0, r: 0, b: 0 };
    const w = frame.w / Math.max(0.01, 1 - k.l - k.r), hh = frame.h / Math.max(0.01, 1 - k.t - k.b);
    return { x: frame.x - k.l * w, y: frame.y - k.t * hh, w, h: hh };
  }

  function startCrop(): void {
    const p = selected;
    const b = p && boxOf(p);
    if (!p || !b) return;
    if (crop) { endCrop(true); return; }
    crop = { p, frame: b, full: fullOf(b, fmt(p).crop) };
    notify();
  }

  function endCrop(keep: boolean): void {
    const c = crop;
    if (!c) return;
    crop = null;
    const b = boxOf(c.p);
    const k = cropOf(c.frame, c.full);
    const old = fmt(c.p).crop ?? { l: 0, t: 0, r: 0, b: 0 };
    const changed = !b || Math.abs(c.frame.x - b.x) > 0.5 || Math.abs(c.frame.y - b.y) > 0.5 || Math.abs(c.frame.w - b.w) > 0.5 || Math.abs(c.frame.h - b.h) > 0.5
      || (["l", "t", "r", "b"] as const).some((s) => Math.abs(k[s] - old[s]) > 1e-4);
    if (keep && changed && selected === c.p) {
      const none = !k.l && !k.t && !k.r && !k.b;
      edit("Zuschneiden", (q) => placeAt({ ...q, format: { ...fmt(q), crop: none ? undefined : k } }, c.frame));
    } else notify();
  }

  /** While cropping: a black handle moves an edge of the frame, the frame itself moves the picture behind it. */
  function cropDrag(e: PointerEvent, el: HTMLElement): void {
    const c = crop!;
    const k = (e.target as HTMLElement).dataset.c;
    const sx = e.clientX, sy = e.clientY;
    const frame0 = { ...c.frame }, full0 = { ...c.full };
    el.setPointerCapture(e.pointerId);
    const move = (ev: PointerEvent) => {
      const dx = ev.clientX - sx, dy = ev.clientY - sy;
      if (!k) { c.full = { ...full0, x: full0.x + dx, y: full0.y + dy }; place(); return; }
      let { x, y, w, h: hh } = frame0;
      if (k.includes("e")) w += dx;
      if (k.includes("s")) hh += dy;
      if (k.includes("w")) { x += dx; w -= dx; }
      if (k.includes("n")) { y += dy; hh -= dy; }
      if (w < 8) { if (k.includes("w")) x -= 8 - w; w = 8; }
      if (hh < 8) { if (k.includes("n")) y -= 8 - hh; hh = 8; }
      c.frame = { x, y, w, h: hh };
      place();
    };
    const up = () => {
      el.removeEventListener("pointermove", move);
      el.removeEventListener("pointerup", up);
      el.removeEventListener("pointercancel", up);
    };
    el.addEventListener("pointermove", move);
    el.addEventListener("pointerup", up);
    el.addEventListener("pointercancel", up);
  }

  /** Ausfüllen / Einpassen: the whole picture fills the frame (cropping what is over) or fits in it. */
  async function cropFill(fit: boolean): Promise<void> {
    const p = selected;
    const b = p && boxOf(p);
    if (!p || !b) return;
    const n = await naturalAsync(p);
    const a = n.w / n.h;
    let fw: number, fh: number;
    if ((b.w / b.h > a) !== fit) { fw = b.w; fh = b.w / a; } else { fh = b.h; fw = b.h * a; }
    const full = { x: b.x + (b.w - fw) / 2, y: b.y + (b.h - fh) / 2, w: fw, h: fh };
    const k = cropOf(b, full);
    setFormat(fit ? "Einpassen" : "Ausfüllen", { crop: !k.l && !k.t && !k.r && !k.b ? undefined : k });
  }

  /** Seitenverhältnis: the frame cut to w:h around its middle, the picture staying where it is. */
  function cropToRatio(r: number): void {
    const p = selected;
    const b = p && boxOf(p);
    if (!p || !b) return;
    const full = fullOf(b, fmt(p).crop);
    let w = b.w, hh = b.h;
    if (w / hh > r) w = hh * r; else hh = w / r;
    const frame = { x: b.x + (b.w - w) / 2, y: b.y + (b.h - hh) / 2, w, h: hh };
    edit("Zuschneiden", (q) => placeAt({ ...q, format: { ...fmt(q), crop: cropOf(frame, full) } }, frame));
  }

  // ----- Transparente Farbe bestimmen -----------------------------------------------------------

  function startPick(): void {
    if (!selected) return;
    picking = true;
    notify();
  }

  async function pickColor(e: PointerEvent, p: SheetPicture, el: HTMLElement): Promise<void> {
    picking = false;
    const r = el.getBoundingClientRect();
    const f = fmt(p);
    const n = await naturalAsync(p);
    const c = f.crop ?? { l: 0, t: 0, r: 0, b: 0 };
    let u = (e.clientX - r.left) / r.width, v = (e.clientY - r.top) / r.height;
    if (f.flipH) u = 1 - u;
    if (f.flipV) v = 1 - v;
    const px = Math.floor((c.l + u * (1 - c.l - c.r)) * n.w), py = Math.floor((c.t + v * (1 - c.t - c.b)) * n.h);
    const img = new Image();
    await new Promise<void>((res) => { img.onload = () => res(); img.onerror = () => res(); img.src = p.src; });
    const cv = document.createElement("canvas");
    cv.width = n.w; cv.height = n.h;
    const g = cv.getContext("2d");
    if (!g) return;
    g.drawImage(img, 0, 0);
    const d = g.getImageData(Math.max(0, Math.min(n.w - 1, px)), Math.max(0, Math.min(n.h - 1, py)), 1, 1).data;
    const hex = [d[0], d[1], d[2]].map((x) => x.toString(16).padStart(2, "0")).join("").toUpperCase();
    setFormat("Transparente Farbe bestimmen", { transparentColor: hex });
  }

  // ----- keys -----------------------------------------------------------------------------------

  // A click elsewhere ends the selection (and Zuschneiden).
  addEventListener("pointerdown", (e) => {
    const t = e.target as HTMLElement;
    if (!selected || t.closest?.(".pic") || t.closest?.(".ribbon-card, .tab, .taskpane, .popup, .rpanel, .menu, .dialog-layer, [role=dialog], [role=menu]")) return;
    select(null);
  }, true);
  addEventListener("keydown", (e) => {
    if (!selected || app.grid.isEditing) return;
    const target = e.target as HTMLElement;
    if (target.closest?.(".taskpane, input, textarea, [contenteditable=true]")) return;
    if (crop && (e.key === "Escape" || e.key === "Enter")) { e.preventDefault(); e.stopPropagation(); endCrop(true); return; }
    if (e.key === "Escape") { e.preventDefault(); e.stopPropagation(); if (picking) { picking = false; notify(); } else select(null); return; }
    if (e.ctrlKey && e.key === "1") { e.preventDefault(); e.stopPropagation(); openPane(); return; }
    if (!app.protectionAllows("objects")) return;
    if (e.key === "Delete" || e.key === "Backspace") {
      e.preventDefault();
      e.stopPropagation();
      const p = selected;
      selected = null;
      change("Bild löschen", (list) => list.filter((q) => q !== p));
      app.pictureSelectionChanged();
      return;
    }
    const step = { ArrowLeft: [-1, 0], ArrowRight: [1, 0], ArrowUp: [0, -1], ArrowDown: [0, 1] }[e.key];
    if (step && !e.altKey) {
      e.preventDefault();
      e.stopPropagation();
      const b = boxOf(selected);
      if (b) edit("Verschieben", (q) => placeAt(q, { ...b, x: b.x + step[0], y: b.y + step[1] }));
    }
  }, true);

  // ----- inserting and replacing ----------------------------------------------------------------

  async function readImage(path: string): Promise<{ src: string; w: number; h: number }> {
    const bytes = await app.platform.readFile(path);
    return imageFromBytes(bytes, path.split(".").pop()?.toLowerCase() ?? "");
  }

  async function imageFromBytes(bytes: Uint8Array, ext = ""): Promise<{ src: string; w: number; h: number }> {
    const type = bytes[0] === 0x89 && bytes[1] === 0x50 ? "image/png" : bytes[0] === 0xff && bytes[1] === 0xd8 ? "image/jpeg" : bytes[0] === 0x47 && bytes[1] === 0x49 ? "image/gif"
      : bytes[0] === 0x42 && bytes[1] === 0x4d ? "image/bmp" : MIME[ext];
    if (!type) throw new Error("Dieses Bildformat wird nicht unterstützt (PNG, JPEG, GIF oder BMP).");
    let bin = "";
    for (let i = 0; i < bytes.length; i += 0x8000) bin += String.fromCharCode(...bytes.subarray(i, i + 0x8000));
    const src = `data:${type};base64,${btoa(bin)}`;
    const size = await new Promise<{ w: number; h: number }>((resolve, reject) => {
      const img = new Image();
      img.onload = () => resolve({ w: img.naturalWidth, h: img.naturalHeight });
      img.onerror = () => reject(new Error("Kein lesbares Bild."));
      img.src = src;
    });
    naturals.set(src, size);
    return { src, ...size };
  }

  /** Einfügen > Bilder > Dieses Gerät: the picture at the active cell, at its natural size. */
  async function insertFromFile(): Promise<void> {
    const path = await app.platform.openImageDialog();
    if (!path) return;
    try { insertImage(await readImage(path)); } catch (e) {
      await messageBox("RaabeCalc", `Das Bild konnte nicht eingefügt werden.\n\n${e instanceof Error ? e.message : String(e)}`);
    }
  }

  function insertImage(img: { src: string; w: number; h: number }): SheetPicture {
    const a = app.sheet.view.active;
    const c = app.grid.rectOnCanvas({ r0: a.row, c0: a.col, r1: a.row, c1: a.col });
    // Excel numbers new objects after all of the sheet's drawing objects ("Grafik 3").
    const n = (app.sheet.pictures?.length ?? 0) + (app.sheet.shapes?.length ?? 0) + (app.sheet.charts?.length ?? 0) + 1;
    let p: SheetPicture = { from: { col: a.col, row: a.row, x: 0, y: 0 }, size: { w: img.w, h: img.h }, src: img.src, name: `Grafik ${n}`, lockAspect: true };
    if (c) p = placeAt(p, { x: c.x, y: c.y, w: img.w * z(), h: img.h * z() });
    change("Bild einfügen", (list) => [...list, p]);
    selected = p;
    shown = undefined;
    app.pictureSelectionChanged();
    return p;
  }

  /** Bild ändern: the new picture in the old one's frame (its width kept, height to its proportions), the format kept. */
  async function replaceWith(img: { src: string; w: number; h: number }): Promise<void> {
    const p = selected;
    const b = p && boxOf(p);
    if (!p || !b) return;
    const { crop: _c, ...rest } = fmt(p);
    void _c;
    edit("Bild ändern", (q) => {
      const { baked: _b, bakedSrc: _s, ...plain } = q;
      void _b; void _s;
      return placeAt({ ...plain, src: img.src, format: Object.keys(rest).length ? rest : undefined }, { ...b, h: (b.w * img.h) / img.w });
    });
  }

  async function changeFromFile(): Promise<void> {
    const path = await app.platform.openImageDialog();
    if (!path) return;
    try { await replaceWith(await readImage(path)); } catch (e) {
      await messageBox("RaabeCalc", `Das Bild konnte nicht geändert werden.\n\n${e instanceof Error ? e.message : String(e)}`);
    }
  }

  async function changeFromClipboard(): Promise<void> {
    const bytes = await app.platform.clipboardReadImage?.();
    if (!bytes) { await messageBox("RaabeCalc", "In der Zwischenablage ist kein Bild."); return; }
    try { await replaceWith(await imageFromBytes(bytes)); } catch { /* not an image */ }
  }

  /** Freistellen: the background of the picture made transparent (a new PNG; Excel keeps its marks beside the original). */
  async function removeBg(run: (bytes: Uint8Array, type: string) => Promise<Uint8Array | null>): Promise<void> {
    const p = selected;
    if (!p) return;
    const m = /^data:([^;]+);base64,(.*)$/.exec(p.src);
    if (!m) return;
    const bin = atob(m[2]);
    const bytes = new Uint8Array(bin.length);
    for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
    const out = await run(bytes, m[1]);
    if (!out) return;
    const img = await imageFromBytes(out, "png");
    edit("Freistellen", (q) => { const { baked: _b, bakedSrc: _s, ...plain } = q; void _b; void _s; return { ...plain, src: img.src }; });
  }

  /** Bild zurücksetzen (the format goes, size and crop stay) or Bild und Größe zurücksetzen (also its own size, uncropped). */
  async function reset(all: boolean): Promise<void> {
    const p = selected;
    const b = p && boxOf(p);
    if (!p || !b) return;
    if (!all) {
      const c = fmt(p).crop;
      edit("Bild zurücksetzen", (q) => placeAt({ ...q, format: c ? { crop: c } : undefined }, b, 0));
      return;
    }
    const n = await naturalAsync(p);
    edit("Bild und Größe zurücksetzen", (q) => placeAt({ ...q, format: undefined }, { x: b.x, y: b.y, w: n.w * z(), h: n.h * z() }, 0));
  }

  /** Bilder komprimieren: the pixels the shown size needs at `ppi`, cropped parts dropped if asked. */
  async function compress(ppi: number, dropCropped: boolean, onlyThis: boolean): Promise<void> {
    const sheets = onlyThis ? [app.sheet] : app.wb.sheets;
    const targets = onlyThis ? (selected ? [selected] : []) : sheets.flatMap((s) => s.pictures ?? []);
    const done = new Map<SheetPicture, SheetPicture>();
    for (const p of targets) {
      const n = await naturalAsync(p);
      const f = fmt(p);
      const c = dropCropped && f.crop ? f.crop : { l: 0, t: 0, r: 0, b: 0 };
      const sx = Math.max(0, c.l) * n.w, sy = Math.max(0, c.t) * n.h, sw = n.w * (1 - Math.max(0, c.l) - Math.max(0, c.r)), sh = n.h * (1 - Math.max(0, c.t) - Math.max(0, c.b));
      const shownW = (p.size?.w ?? n.w) / 96, shownH = (p.size?.h ?? n.h) / 96;
      const k = Math.min(1, (shownW * ppi) / sw, (shownH * ppi) / sh);
      if (k > 0.98 && !(dropCropped && f.crop)) continue;
      const img = new Image();
      await new Promise<void>((res) => { img.onload = () => res(); img.onerror = () => res(); img.src = p.src; });
      const cv = document.createElement("canvas");
      cv.width = Math.max(1, Math.round(sw * k)); cv.height = Math.max(1, Math.round(sh * k));
      cv.getContext("2d")?.drawImage(img, sx, sy, sw, sh, 0, 0, cv.width, cv.height);
      const jpeg = p.src.startsWith("data:image/jpeg");
      const src = cv.toDataURL(jpeg ? "image/jpeg" : "image/png", 0.9);
      naturals.set(src, { w: cv.width, h: cv.height });
      // Only the cropped-away parts outside the picture go; padding (negative crop) stays.
      const left = f.crop && dropCropped ? { l: Math.min(0, f.crop.l), t: Math.min(0, f.crop.t), r: Math.min(0, f.crop.r), b: Math.min(0, f.crop.b) } : f.crop;
      const nf = { ...f, crop: left && (left.l || left.t || left.r || left.b) ? left : undefined };
      const { baked: _b, bakedSrc: _s, ...plain } = p;
      void _b; void _s;
      done.set(p, { ...plain, src, format: Object.values(nf).some((v) => v !== undefined) ? nf : undefined });
    }
    if (!done.size) return;
    app.change("Bilder komprimieren", () => {
      for (const s of sheets) {
        if (!s.pictures?.some((p) => done.has(p))) continue;
        app.history.whole(s);
        s.pictures = s.pictures.map((p) => done.get(p) ?? p);
      }
    });
    if (selected && done.has(selected)) selected = done.get(selected)!;
    for (const p of done.values()) bake(p);
    notify();
  }

  // ----- Anordnen -------------------------------------------------------------------------------

  function reorder(label: string, how: "forward" | "backward" | "front" | "back"): void {
    const p = selected;
    if (!p) return;
    change(label, (list) => {
      const i = list.indexOf(p);
      if (i < 0) return list;
      list.splice(i, 1);
      const j = how === "front" ? list.length : how === "back" ? 0 : how === "forward" ? Math.min(list.length, i + 1) : Math.max(0, i - 1);
      list.splice(j, 0, p);
      return list;
    });
    notify();
  }

  // ----- Grafik formatieren ----------------------------------------------------------------------

  let paneOpen = false;
  let paneSpec: (() => import("@raabe/shell").PaneSpec) | null = null;

  /** Grafik formatieren, at a tab ("fill", "effects", "size", "picture") when asked. */
  function openPane(tab?: string): void {
    if (!selected || !paneSpec) return;
    paneOpen = true;
    document.documentElement.style.setProperty("--pane-top", `${document.querySelector(".calc-main")?.getBoundingClientRect().top ?? 150}px`);
    app.taskPane.onClose = () => { paneOpen = false; };
    if (tab) app.taskPane.showAt(paneSpec(), tab);
    else app.taskPane.show(paneSpec());
  }
  listeners.push(() => { if (paneOpen && app.taskPane.isOpen && selected && paneSpec) app.taskPane.show(paneSpec()); });

  return {
    place, insertFromFile,
    insertPicture: (p: SheetPicture) => change("Bild einfügen", (list) => [...list, p]),
    insertImage,
    selected: (): SheetPicture | null => selected,
    select: (p: SheetPicture | null) => select(p),
    deselect: (): void => { if (selected) select(null); },
    format: (): PictureFormat => (selected ? fmt(selected) : {}),
    box: (): Box | null => (selected ? boxOf(selected) : null),
    natural: (p: SheetPicture) => natural(p),
    naturalAsync,
    lookOf: (p: SheetPicture, f: PictureFormat) => lookOf(p, f),
    edit, setFormat, placeAt, boxOf,
    /** Livevorschau: the selected picture drawn with `f` (null ends it). */
    preview: (f: PictureFormat | null) => { preview = f && selected ? { p: selected, f } : null; shown = undefined; app.grid.contentChanged(); },
    startCrop, endCrop, cropping: () => !!crop, cropFill, cropToRatio, startPick,
    changeFromFile, changeFromClipboard, removeBg, reset, compress, reorder,
    snapGrid: () => snapGrid, toggleSnap: () => { snapGrid = !snapGrid; },
    setPane: (spec: () => import("@raabe/shell").PaneSpec) => { paneSpec = spec; },
    openPane, paneOpen: () => paneOpen && app.taskPane.isOpen,
    onChange: (l: () => void) => listeners.push(l),
    /** The pictures worked on in the background (Korrekturen worked into the picture) are done. */
    settle: async () => { while (pending.size) await Promise.all([...pending]); },
    change,
  };
}
