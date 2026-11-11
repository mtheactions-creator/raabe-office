// Formkontur > Skizziert (D-072): PowerPoint's sketched outlines, written as PowerPoint writes them (recorded by
// probe_powerpoint_sketch.py): an extension in the shape's a:ln with ask:lineSketchStyleProps, a random seed
// (sd), a copy of the shape's geometry and the type (lineSketchCurved, lineSketchFreehand, lineSketchScribble).
// We draw the shape with a displacement of our own, stronger from Gekrümmt to Skizze (PowerPoint's own
// drawing of the sketch is not documented).

import { pictureMenu, type MenuEntry } from "@raabe/shell";
import { PptxDom } from "@raabe/ooxml";
import type { PresentApp } from "./app";

const { kid, kids, make } = PptxDom;
export const NS_ASK = "http://schemas.microsoft.com/office/drawing/2018/sketchyshapes";
const URI = "{C807C97D-BFC1-408E-A445-0C87EB9F89A2}";
export const SKETCH_TYPES: [string, string | null][] = [["Keine", null], ["Gekrümmt", "lineSketchCurved"], ["Freihand", "lineSketchFreehand"], ["Skizze", "lineSketchScribble"]];

/** The sketch type of a shape (its spPr's a:ln), or null. */
export function sketchOf(sp: Element): string | null {
  const t = sp.getElementsByTagNameNS(NS_ASK, "type")[0];
  return t?.firstElementChild?.localName ?? null;
}

/** Sets or removes the sketch of a shape's outline. */
export function setSketch(spPr: Element, type: string | null): void {
  const doc = spPr.ownerDocument;
  let ln = kid(spPr, "ln");
  const old = ln ? kids(kid(ln, "extLst")).find((e) => e.getAttribute("uri") === URI) : undefined;
  old?.remove();
  if (ln && kid(ln, "extLst") && !kid(ln, "extLst")!.children.length) kid(ln, "extLst")!.remove();
  if (!type) return;
  if (!ln) {
    ln = make(doc, "a:ln");
    spPr.insertBefore(ln, kids(spPr).find((c) => ["effectLst", "effectDag", "scene3d", "sp3d", "extLst"].includes(c.localName)) ?? null);
  }
  let ext = kid(ln, "extLst");
  if (!ext) { ext = make(doc, "a:extLst"); ln.append(ext); }
  const e = make(doc, "a:ext", { uri: URI });
  const props = doc.createElementNS(NS_ASK, "ask:lineSketchStyleProps");
  props.setAttribute("sd", String(Math.floor(Math.random() * 2147483647)));
  const geom = kids(spPr).find((c) => c.localName === "prstGeom" || c.localName === "custGeom");
  props.append(geom ? geom.cloneNode(true) : make(doc, "a:prstGeom", { prst: "rect" }));
  const t = doc.createElementNS(NS_ASK, "ask:type");
  t.append(doc.createElementNS(NS_ASK, `ask:${type}`));
  props.append(t);
  e.append(props);
  ext.append(e);
}

/** Formkontur > Skizziert's entries. */
/** The pictures of the Skizziert gallery (D-163): a line drawn straight, curved, by hand or scribbled (our own). */
const SKETCH_PATHS: Record<string, string> = {
  none: "M4 12 H100",
  lineSketchCurved: "M4 12 C 30 8, 50 16, 76 11 S 96 12, 100 12",
  lineSketchFreehand: "M4 12 C 14 9, 20 15, 32 11 S 48 15, 58 10 S 78 15, 88 11 L 100 12",
  lineSketchScribble: "M4 11 C 20 8, 36 15, 52 10 S 84 14, 100 11 M6 13 C 24 15, 40 9, 60 13 S 86 10, 99 13",
};

export function sketchMenu(app: PresentApp): MenuEntry[] {
  const apply = (label: string, type: string | null) => {
    const views = app.editor.selectedViews.filter((v) => v.kind === "sp" || v.kind === "cxnSp" || v.kind === "pic");
    app.editor.edit(type ? `Skizziert: ${label}` : "Skizziert", () => {
      for (const v of views) { const sp = kid(app.editor.shapeEl(v.id), "spPr"); if (sp) setSketch(sp, type); }
    });
  };
  // PowerPoint shows the four kinds as pictures, their names as tooltips.
  const entries = SKETCH_TYPES.map(([label, type]) => ({ value: type ?? "none", label,
    svg: `<svg xmlns="http://www.w3.org/2000/svg" width="104" height="24" aria-hidden="true"><path d="${SKETCH_PATHS[type ?? "none"]}" fill="none" stroke="#404040" stroke-width="1.6" stroke-linecap="round"/></svg>` }));
  return [{ block: () => pictureMenu(entries, 1, [104, 24], undefined, (v) => { const t = SKETCH_TYPES.find(([, x]) => (x ?? "none") === v); if (t) apply(t[0], t[1]); }) }];
}

let filterCount = 0;
/** Draws a sketched outline: an SVG displacement over the shape's drawing, its strength by the type. */
export function drawSketch(box: HTMLElement, type: string, seed: number): void {
  const scale = type === "lineSketchScribble" ? 5 : type === "lineSketchFreehand" ? 3.5 : 2;
  const freq = type === "lineSketchCurved" ? 0.012 : 0.03;
  const id = `sketch-${++filterCount}`;
  const ns = "http://www.w3.org/2000/svg";
  const svg = document.createElementNS(ns, "svg");
  svg.setAttribute("width", "0"); svg.setAttribute("height", "0");
  svg.style.position = "absolute";
  svg.innerHTML = `<filter id="${id}" x="-5%" y="-5%" width="110%" height="110%"><feTurbulence type="fractalNoise" baseFrequency="${freq}" numOctaves="2" seed="${seed % 1000}" result="n"/><feDisplacementMap in="SourceGraphic" in2="n" scale="${scale}" xChannelSelector="R" yChannelSelector="G"/></filter>`;
  box.append(svg);
  // The shape's drawing is the box's first child (its text comes after it and stays straight).
  const el = box.firstElementChild as Element | null;
  if (el && el !== svg && el instanceof HTMLElement) el.style.filter = `url(#${id}) ${el.style.filter}`.trim();
}
