// A slide as DOM (D-070): the background, the layout's and master's design shapes, then the
// slide's shapes in z-order: preset and custom geometry through the text engine's shape renderer
// (shared with RaabeWrite and RaabeCalc), pictures with their crop and frame, tables, charts and
// groups. Sizes are CSS pixels at 96 dpi (1 px = 9525 EMU); zoom scales the whole slide.

import { NS_ASK, drawSketch } from "../sketch";
import { cameraTransform, shapeElement, shapeTextRect, pictureElement, type PictureFormat } from "@raabe/text-engine";
import type { Pptx, ShapeView, SlideView } from "@raabe/ooxml";
import { EMU_PX, renderText, type TextOptions } from "./text";
import { renderTable } from "./table";
import { renderChart } from "./chart";
import { applyTextGradients, drawWarp, warpOf } from "./warp";

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
  /** The slide's theme colours (set by renderSlide; text gradients and warps resolve their colours with them). */
  themeColors?: Record<string, string>;
  /** The slide's design for charts (set by renderSlide, D-195): accents, background and text colour. */
  design?: { accents: string[]; surface: { background: string; text: string } };
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

let duoIds = 0;
/**
 * A background picture recoloured as Office's a:duotone (D-108): its brightness mapped linearly from the dark
 * colour to the light one (an SVG filter: luminance, then a table per channel), stretched over the slide.
 */
function duotoneLayer(url: string, d: { dark: string; light: string }, w: number, h: number): Element[] {
  const id = `raabe-duotone-${++duoIds}`;
  const ch = (hex: string, i: number) => (parseInt(hex.slice(i, i + 2), 16) / 255).toFixed(4);
  const svg = document.createElementNS("http://www.w3.org/2000/svg", "svg");
  svg.setAttribute("width", "0"); svg.setAttribute("height", "0");
  svg.style.position = "absolute";
  svg.innerHTML = `<filter id="${id}" color-interpolation-filters="sRGB"><feColorMatrix type="matrix" values="0.2126 0.7152 0.0722 0 0 0.2126 0.7152 0.0722 0 0 0.2126 0.7152 0.0722 0 0 0 0 0 1 0"/>` +
    `<feComponentTransfer><feFuncR type="table" tableValues="${ch(d.dark, 0)} ${ch(d.light, 0)}"/><feFuncG type="table" tableValues="${ch(d.dark, 2)} ${ch(d.light, 2)}"/><feFuncB type="table" tableValues="${ch(d.dark, 4)} ${ch(d.light, 4)}"/></feComponentTransfer></filter>`;
  const layer = document.createElement("div");
  layer.style.cssText = `position:absolute;left:0;top:0;width:${w}px;height:${h}px;background:url("${url}") 0 0 / 100% 100% no-repeat;filter:url(#${id})`;
  return [svg, layer];
}

export function renderSlide(v: SlideView, size: { cx: number; cy: number }, o: RenderOptions): Rendered {
  const w = size.cx / EMU_PX, h = size.cy / EMU_PX;
  const el = document.createElement("div");
  el.className = "slide";
  el.style.cssText = `position:relative;width:${w}px;height:${h}px;overflow:hidden;background:${fillCss(v.background.fill)}`;
  if (v.background.image) {
    const url = o.image(v.background.image);
    if (url && v.background.duotone) el.append(...duotoneLayer(url, v.background.duotone, w, h));
    else if (url) el.style.background = `url("${url}") center / cover no-repeat`;
  }
  const shapes = new Map<number, HTMLElement>();
  const tc = v.theme.colors;
  const bgFill = v.background.fill;
  const bgHex = bgFill.kind === "solid" ? bgFill.color : bgFill.kind === "gradient" && bgFill.stops.length ? bgFill.stops[Math.floor(bgFill.stops.length / 2)].color : tc.bg1 ?? "FFFFFF";
  o = { ...o, themeColors: tc, design: { accents: [1, 2, 3, 4, 5, 6].map((i) => `#${tc[`accent${i}`] ?? "000000"}`), surface: { background: `#${bgHex}`, text: `#${tc.tx1 ?? "000000"}` } } };
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
  // Einfügen > Aktion: the shape's click action, followed in the show.
  const nv = Array.from(s.el.children).find((c) => c.localName.startsWith("nv"));
  const act = nv?.firstElementChild ? Array.from(nv.firstElementChild.children).find((c) => c.localName === "hlinkClick") : undefined;
  if (act && !o.editing) { const a = act.getAttribute("action") ?? ""; const jump = /jump=(\w+)/.exec(a)?.[1]; const macro = /^ppaction:\/\/macro\?name=(.+)$/i.exec(a)?.[1]; if (jump) box.dataset.action = jump; else if (macro) box.dataset.macro = decodeURIComponent(macro); else box.dataset.link = act.getAttributeNS("http://schemas.openxmlformats.org/officeDocument/2006/relationships", "id") ?? ""; }
  const px = parent ? x - parent.x / EMU_PX : x, py = parent ? y - parent.y / EMU_PX : y;
  box.style.cssText = `position:absolute;left:${px}px;top:${py}px;width:${w}px;height:${h}px;transform:${transformOf(s) || "none"};transform-origin:50% 50%`;
  map?.set(s.id, box);
  // 3D-Drehung of a shape (pictures turn inside pictureElement).
  if (s.kind === "sp" || s.kind === "cxnSp") { const cam = cameraTransform(s.format.scene3dXml); if (cam) box.style.transform = `${transformOf(s) || ""} ${cam}`.trim(); }
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
    else if (s.frame === "chart" && s.chart) box.append(renderChart(o.pptx, s.chart, w, h, o.design?.accents, o.design?.surface));
    else if (s.frame === "zoom" && s.zoom?.tiles) {
      // Zusammenfassungszoom: each section's tile at its place in the frame (scaled with the frame).
      const sx = s.w ? w / (s.w / EMU_PX) : 1, sy = s.h ? h / (s.h / EMU_PX) : 1;
      for (const t of s.zoom.tiles) {
        const tile = document.createElement("div");
        const tw = (t.w / EMU_PX) * sx, th = (t.h / EMU_PX) * sy;
        tile.style.cssText = `position:absolute;left:${(t.x / EMU_PX) * sx}px;top:${(t.y / EMU_PX) * sy}px;width:${tw}px;height:${th}px;box-shadow:inset 0 0 0 0.33px #d3d3d3`;
        const url = t.path ? o.image(t.path) : undefined;
        if (url) tile.append(pictureElement(url, tw, th, { ...s.format, crop: undefined }));
        if (!o.editing) { tile.dataset.zoom = "0"; tile.dataset.zoomSection = t.section; tile.dataset.zoomBack = "1"; }
        box.append(tile);
      }
    } else if (s.frame === "zoom") {
      // Folienzoom: its preview picture with a 0.25 pt light grey line; in the show a click goes to the slide.
      const url = s.pic?.path ? o.image(s.pic.path) : undefined;
      if (url) box.append(pictureElement(url, w, h, { ...s.format, crop: s.pic?.crop }));
      else box.append(placeholderBox(w, h));
      const line = document.createElement("div");
      line.style.cssText = "position:absolute;inset:0;box-shadow:inset 0 0 0 0.33px #d3d3d3;pointer-events:none";
      box.append(line);
      if (!o.editing && s.zoom) { box.dataset.zoom = String(s.zoom.sldId); if (s.zoom.section) box.dataset.zoomSection = s.zoom.section; if (s.zoom.back) box.dataset.zoomBack = "1"; }
    } else if (s.frame === "diagram" && s.children) {
      // The whole graphic's background and outline (SmartArt > Format > Fülleffekt, Formkontur; D-154), under its shapes.
      if ((s.diagramBg && s.diagramBg.kind !== "none") || s.diagramLine) {
        const bg = document.createElement("div");
        bg.className = "diagram-bg";
        bg.style.cssText = `position:absolute;inset:0;pointer-events:none;${s.diagramBg && s.diagramBg.kind !== "none" ? `background:${fillCss(s.diagramBg)};` : ""}${s.diagramLine ? `box-shadow:inset 0 0 0 ${(s.diagramLine.width * 96) / 72}px #${s.diagramLine.color};` : ""}`;
        box.append(bg);
      }
      for (const c of s.children) box.append(renderShape({ ...c, rot: c.rot - s.rot }, { ...o, editing: false, selection: null }, null, s));
    }
    else box.append(placeholderBox(w, h));
  } else if (s.kind === "sp" || s.kind === "cxnSp") {
    const url = s.pic?.path ? o.image(s.pic.path) : undefined;
    if (url) {
      const pe = pictureElement(url, w, h, { ...s.format, fill: undefined, crop: s.pic?.crop });
      // Fülleffekt > Struktur (D-078): a tiled picture (a:tile) repeats at its own size, as PowerPoint draws textures.
      const tile = Array.from(s.el.getElementsByTagNameNS("http://schemas.openxmlformats.org/drawingml/2006/main", "tile"))[0];
      const img = pe.querySelector("img");
      if (tile && img) {
        const k = Number(tile.getAttribute("sx") ?? 100000) / 100000;
        const frame = img.parentElement as HTMLElement;
        img.remove();
        frame.style.backgroundImage = `url("${url}")`;
        frame.style.backgroundRepeat = "repeat";
        if (k !== 1) frame.style.backgroundSize = `${k * 100}% auto`;
      }
      box.append(pe);
      if (tile && s.format.line) { const look = shapeElement(w, h, { ...s.format, fill: undefined }, s.info); if (look) box.append(look); }
    } else {
      const look = s.ph && !s.format.fill && !s.format.line ? null : shapeElement(w, h, s.format, s.info);
      if (look) box.append(look);
    }
  }
  if (s.text) {
    const editing = o.selection && o.selection.id === s.id;
    // An empty placeholder shows its prompt and a dashed outline in Normalansicht.
    if (o.editing && s.ph && s.text.empty) box.classList.add("ph-empty");
    const r = shapeTextRect(s.format, { ...s.info, body: { l: s.text.body.l / 12700, t: s.text.body.t / 12700, r: s.text.body.r / 12700, b: s.text.body.b / 12700, anchor: "t", wrap: true, autofit: "none" } }, w, h);
    // Texteffekte > Transformieren (D-076): drawn bent unless the text is being edited (then plain, to type into).
    const warp = !editing && s.kind === "sp" ? warpOf(s.el) : null;
    const t = warp ? document.createElement("div") : renderText(s.text, r.w, r.h, { prompt: o.editing && !editing, slideNumber: o.slideNumber, selection: editing ? o.selection!.sel : null, fit: o.fits?.get(s.id) });
    // PowerPoint spreads warped text over the whole shape, without the insets (measured, pptx-texteffects).
    if (warp) { t.className = "txt txt-warped"; t.style.cssText = `position:absolute;width:${w}px;height:${h}px`; t.append(drawWarp(s.text, warp, w, h, s.el, o.themeColors ?? {})); }
    else applyTextGradients(t, s.el, o.themeColors ?? {});
    t.style.left = `${warp ? 0 : r.x}px`;
    t.style.top = `${warp ? 0 : r.y}px`;
    // Text is not mirrored with the outline: Horizontal kippen leaves it readable, Vertikal kippen
    // turns it upside down (PowerPoint), so one flip is undone sideways and two not at all.
    if (s.flipH !== s.flipV) t.style.transform = "scaleX(-1)";
    if (s.text.body.rot) t.style.transform = `${t.style.transform} rotate(${s.text.body.rot}deg)`;
    box.append(t);
  }
  // Formkontur > Skizziert: the shape drawn with a sketched outline.
  if (s.kind === "sp" || s.kind === "cxnSp" || s.kind === "pic") {
    const t = s.el.getElementsByTagNameNS(NS_ASK, "type")[0]?.firstElementChild?.localName;
    if (t) drawSketch(box, t, Number(s.el.getElementsByTagNameNS(NS_ASK, "lineSketchStyleProps")[0]?.getAttribute("sd") ?? 1));
  }
  return box;
}

function placeholderBox(w: number, h: number): HTMLElement {
  const d = document.createElement("div");
  d.style.cssText = `position:absolute;left:0;top:0;width:${w}px;height:${h}px;background:#f2f2f2;outline:1px solid #bfbfbf`;
  return d;
}
