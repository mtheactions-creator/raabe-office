// A slide as DOM (D-070): the background, the layout's and master's design shapes, then the
// slide's shapes in z-order: preset and custom geometry through the text engine's shape renderer
// (shared with RaabeWrite and RaabeCalc), pictures with their crop and frame, tables, charts and
// groups. Sizes are CSS pixels at 96 dpi (1 px = 9525 EMU); zoom scales the whole slide.

import { shapeElement, shapeTextRect, pictureElement, type PictureFormat } from "@raabe/text-engine";
import type { Pptx, ShapeView, SlideView } from "@raabe/ooxml";
import { EMU_PX, renderText, type TextOptions } from "./text";
import { renderTable } from "./table";
import { renderChart } from "./chart";

export interface RenderOptions {
  /** Normalansicht: prompts and outlines of empty placeholders. */
  editing: boolean;
  slideNumber?: number;
  /** Picture bytes by part path → object URL. */
  image(path: string): string | undefined;
  /** Text selection of the shape being edited. */
  selection?: { id: number; sel: TextOptions["selection"]; cell?: { r: number; c: number }; cells?: { r0: number; c0: number; r1: number; c1: number } } | null;
  /** Autofit computed while editing, by shape id. */
  fits?: Map<number, { fontScale: number; lnSpcReduction: number }>;
  pptx: Pptx;
}

export interface Rendered {
  el: HTMLElement;
  /** Element of each shape by id (top level and group members). */
  shapes: Map<number, HTMLElement>;
}

function fillCss(f: SlideView["background"]["fill"]): string {
  if (f.kind === "solid") return `#${f.color}`;
  if (f.kind === "gradient") {
    const stops = f.stops.map((s) => `#${s.color} ${(s.pos * 100).toFixed(1)}%`).join(",");
    return f.path ? `radial-gradient(circle,${stops})` : `linear-gradient(${(f.angle ?? 0) + 90}deg,${stops})`;
  }
  return "#FFFFFF";
}

export function renderSlide(v: SlideView, size: { cx: number; cy: number }, o: RenderOptions): Rendered {
  const w = size.cx / EMU_PX, h = size.cy / EMU_PX;
  const el = document.createElement("div");
  el.className = "slide";
  el.style.cssText = `position:relative;width:${w}px;height:${h}px;overflow:hidden;background:${fillCss(v.background.fill)}`;
  if (v.background.image) {
    const url = o.image(v.background.image);
    if (url) el.style.background = `url("${url}") center / cover no-repeat`;
  }
  const shapes = new Map<number, HTMLElement>();
  for (const s of v.behind) if (!s.hidden) el.append(renderShape(s, { ...o, editing: false, selection: null }, null));
  for (const s of v.shapes) {
    if (s.hidden) continue;
    el.append(renderShape(s, o, shapes));
  }
  return { el, shapes };
}

const transformOf = (s: ShapeView) => {
  const parts: string[] = [];
  if (s.rot) parts.push(`rotate(${s.rot}deg)`);
  if (s.flipH) parts.push("scaleX(-1)");
  if (s.flipV) parts.push("scaleY(-1)");
  return parts.join(" ");
};

export function renderShape(s: ShapeView, o: RenderOptions, map: Map<number, HTMLElement> | null, parent?: ShapeView): HTMLElement {
  const x = s.x / EMU_PX, y = s.y / EMU_PX, w = Math.max(s.w / EMU_PX, 0), h = Math.max(s.h / EMU_PX, 0);
  const box = document.createElement("div");
  box.className = `shp shp-${s.kind}`;
  box.dataset.id = String(s.id);
  const px = parent ? x - parent.x / EMU_PX : x, py = parent ? y - parent.y / EMU_PX : y;
  box.style.cssText = `position:absolute;left:${px}px;top:${py}px;width:${w}px;height:${h}px;transform:${transformOf(s) || "none"};transform-origin:50% 50%`;
  map?.set(s.id, box);
  if (s.kind === "grpSp") {
    // Group members are placed on the slide already (scaled); the group only turns them.
    for (const c of s.children ?? []) if (!c.hidden) box.append(renderShape({ ...c, rot: c.rot - s.rot }, o, map, s));
    return box;
  }
  if (s.kind === "pic") {
    const url = s.pic?.path ? o.image(s.pic.path) : undefined;
    const f: PictureFormat = { ...s.format, crop: s.pic?.crop };
    if (url) box.append(pictureElement(url, w, h, f));
    else box.append(placeholderBox(w, h));
  } else if (s.kind === "graphicFrame") {
    if (s.frame === "table") box.append(renderTable(s, o));
    else if (s.frame === "chart" && s.chart) box.append(renderChart(o.pptx, s.chart, w, h));
    else if (s.frame === "diagram" && s.children) for (const c of s.children) box.append(renderShape({ ...c, rot: c.rot - s.rot }, { ...o, editing: false, selection: null }, null, s));
    else box.append(placeholderBox(w, h));
  } else if (s.kind === "sp" || s.kind === "cxnSp") {
    const url = s.pic?.path ? o.image(s.pic.path) : undefined;
    if (url) box.append(pictureElement(url, w, h, { ...s.format, fill: undefined, crop: s.pic?.crop }));
    else {
      const look = s.ph && !s.format.fill && !s.format.line ? null : shapeElement(w, h, s.format, s.info);
      if (look) box.append(look);
    }
  }
  if (s.text) {
    const editing = o.selection && o.selection.id === s.id;
    // An empty placeholder shows its prompt and a dashed outline in Normalansicht.
    if (o.editing && s.ph && s.text.empty) box.classList.add("ph-empty");
    const r = shapeTextRect(s.format, { ...s.info, body: { l: s.text.body.l / 12700, t: s.text.body.t / 12700, r: s.text.body.r / 12700, b: s.text.body.b / 12700, anchor: "t", wrap: true, autofit: "none" } }, w, h);
    const t = renderText(s.text, r.w, r.h, { prompt: o.editing && !editing, slideNumber: o.slideNumber, selection: editing ? o.selection!.sel : null, fit: o.fits?.get(s.id) });
    t.style.left = `${r.x}px`;
    t.style.top = `${r.y}px`;
    // Text is not mirrored with the outline: Horizontal kippen leaves it readable, Vertikal kippen
    // turns it upside down (PowerPoint), so one flip is undone sideways and two not at all.
    if (s.flipH !== s.flipV) t.style.transform = "scaleX(-1)";
    if (s.text.body.rot) t.style.transform = `${t.style.transform} rotate(${s.text.body.rot}deg)`;
    box.append(t);
  }
  return box;
}

function placeholderBox(w: number, h: number): HTMLElement {
  const d = document.createElement("div");
  d.style.cssText = `position:absolute;left:0;top:0;width:${w}px;height:${h}px;background:#f2f2f2;outline:1px solid #bfbfbf`;
  return d;
}
