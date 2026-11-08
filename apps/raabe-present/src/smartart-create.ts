// Einfügen > SmartArt (D-077): the dialog "SmartArt-Grafik auswählen" (PowerPoint's categories on the left, the
// layouts as pictures drawn by our engine, the chosen one large on the right), the new diagram with its parts,
// laying a diagram out again after its items changed (Textbereich, Form hinzufügen, Höher/Tiefer stufen, Nach
// oben/unten, Layout ändern) and Start > In SmartArt konvertieren (a text box's paragraphs become the items).

import { h, showDialog } from "@raabe/shell";
import { PptxDom } from "@raabe/ooxml";
import type { PresentApp } from "./app";
import * as S from "./shapeops";
import { SMARTART_LAYOUTS, type CatalogEntry } from "./smartart-catalog.gen";
import { fitSize, layout } from "./smartart-engine";
import { CT, NS_DGM, NS_DSP, RT, colorsXml, dataModelXml, drawingXml, frameXml, layoutXml, quickStyleXml, schemeOfColors, styleOfQuick, type DItem } from "./smartart-parts";

const { kid, kids } = PptxDom;

/** PowerPoint's categories in its dialog's order. */
export const CATEGORIES: [string, string][] = [["", "Alle"], ["list", "Liste"], ["process", "Prozess"], ["cycle", "Zyklus"], ["hierarchy", "Hierarchie"],
  ["relationship", "Beziehung"], ["matrix", "Matrix"], ["pyramid", "Pyramide"], ["picture", "Grafik"]];

export const entryById = (id: string): CatalogEntry => SMARTART_LAYOUTS.find((e) => e.id === id) ?? SMARTART_LAYOUTS[0];

/** The drawing for items in a w × h frame (EMU). */
export function drawItems(entry: CatalogEntry, items: DItem[], w: number, h: number) {
  const shapes = layout(entry.spec, items, w, h);
  const size = fitSize(shapes, items.map((i) => i.text || "[Text]"));
  for (const s of shapes) if (s.item !== undefined) s.size = items[s.item].level > 0 ? Math.max(5, size * 0.8) : size;
  return shapes;
}

/** A picture of a layout for the gallery: our engine's drawing with three short items, as SVG. */
export function thumb(entry: CatalogEntry, w = 96, h = 72): SVGSVGElement {
  const ns = "http://www.w3.org/2000/svg";
  const svg = document.createElementNS(ns, "svg");
  svg.setAttribute("width", String(w)); svg.setAttribute("height", String(h));
  svg.setAttribute("viewBox", `0 0 ${w * 12700} ${h * 12700}`);
  const items = [{ text: "", level: 0 }, { text: "", level: 0 }, { text: "", level: 0 }];
  for (const s of layout(entry.spec, items, w * 12700 * 0.9, h * 12700 * 0.9)) {
    const el = document.createElementNS(ns, s.prst === "ellipse" ? "ellipse" : "rect");
    const x = s.x + w * 12700 * 0.05, y = s.y + h * 12700 * 0.05;
    if (s.prst === "ellipse") { el.setAttribute("cx", String(x + s.w / 2)); el.setAttribute("cy", String(y + s.h / 2)); el.setAttribute("rx", String(s.w / 2)); el.setAttribute("ry", String(s.h / 2)); }
    else { el.setAttribute("x", String(x)); el.setAttribute("y", String(y)); el.setAttribute("width", String(Math.max(s.w, 6000))); el.setAttribute("height", String(Math.max(s.h, 6000))); if (s.prst === "roundRect") el.setAttribute("rx", String(Math.min(s.w, s.h) * 0.15)); }
    if (s.rot) el.setAttribute("transform", `rotate(${s.rot} ${x + s.w / 2} ${y + s.h / 2})`);
    el.setAttribute("fill", s.look === "line" || s.look === "text" ? "none" : s.look === "conn" ? "#9DC3E6" : s.look === "bg" ? "rgba(91,155,213,.45)" : "#4472C4");
    el.setAttribute("stroke", s.look === "line" ? "#4472C4" : "#fff");
    el.setAttribute("stroke-width", "12700");
    svg.append(el);
  }
  return svg;
}

/** SmartArt-Grafik auswählen: the chosen layout's id, or null. */
export async function chooseLayout(current?: string): Promise<CatalogEntry | null> {
  let cat = "";
  let pick = current ? entryById(current) : SMARTART_LAYOUTS[0];
  const grid = h("div", { class: "sa-grid", role: "listbox" });
  const big = h("div", { class: "sa-preview" });
  const showPick = () => { big.replaceChildren(thumb(pick, 240, 180), h("div", { class: "sa-name" }, pick.name)); for (const b of Array.from(grid.children)) b.classList.toggle("on", (b as HTMLElement).dataset.id === pick.id); };
  let accept: (() => void) | null = null;
  const fill = () => {
    grid.replaceChildren(...SMARTART_LAYOUTS.filter((e) => !cat || e.category === cat).map((e) => {
      const b = h("button", { type: "button", class: "sa-item", title: e.name, "data-id": e.id, onclick: () => { pick = e; showPick(); }, ondblclick: () => { pick = e; accept?.(); } }, thumb(e));
      return b;
    }));
    showPick();
  };
  const cats = h("div", { class: "sa-cats" }, ...CATEGORIES.map(([id, label]) => h("button", { type: "button", class: `sa-cat${id === cat ? " on" : ""}`, onclick: (ev: MouseEvent) => {
    cat = id;
    for (const c of Array.from(cats.children)) c.classList.toggle("on", c === ev.currentTarget);
    fill();
  } }, label)));
  fill();
  const body = h("div", { class: "sa-dialog" }, cats, grid, big);
  const result = showDialog({ title: "SmartArt-Grafik auswählen", body, buttons: [{ label: "OK", value: true, primary: true }, { label: "Abbrechen", value: false }], cancel: false });
  accept = () => (body.closest("[role=dialog]")?.querySelector(".btn.primary, button.primary") as HTMLButtonElement | null)?.click();
  return (await result) ? pick : null;
}

/** Writes a new diagram's parts and its frame on the slide; returns the frame's id. */
/** A diagram's look: Farben ändern (colour set), SmartArt-Formatvorlage (quick style), Von rechts nach links. */
export interface DiagramLook { scheme?: string; style?: string; rtl?: boolean }

export function createDiagram(app: PresentApp, entry: CatalogEntry, items: DItem[], box: { x: number; y: number; w: number; h: number }, replace?: Element, scheme = "accent1_2", look: DiagramLook = {}): number {
  const pptx = app.deck.pptx;
  const slide = app.slidePath()!;
  // A diagram laid out again keeps its parts (the same names and relationships); a new one gets the next free.
  const old = replace?.getElementsByTagNameNS(NS_DGM, "relIds")[0];
  const oldData = old ? pptx.target(slide, old.getAttributeNS(PptxDom.NS_R, "dm") ?? "") : undefined;
  const oldExt = oldData && pptx.has(oldData) ? pptx.doc(oldData).getElementsByTagNameNS(NS_DSP, "dataModelExt")[0] : undefined;
  const oldDrawRel = oldExt?.getAttribute("relId") ?? "";
  const target = (rid: string | null) => (rid ? pptx.target(slide, rid) : undefined);
  const n = pptx.freeName("ppt/diagrams/data#.xml").match(/(\d+)\.xml$/)![1];
  const p = {
    data: oldData ?? `ppt/diagrams/data${n}.xml`, layout: target(old?.getAttributeNS(PptxDom.NS_R, "lo") ?? null) ?? `ppt/diagrams/layout${n}.xml`,
    style: target(old?.getAttributeNS(PptxDom.NS_R, "qs") ?? null) ?? `ppt/diagrams/quickStyle${n}.xml`, colors: target(old?.getAttributeNS(PptxDom.NS_R, "cs") ?? null) ?? `ppt/diagrams/colors${n}.xml`,
    drawing: (oldDrawRel && target(oldDrawRel)) || `ppt/diagrams/drawing${n}.xml`,
  };
  let id = 0;
  app.editor.edit(replace ? "SmartArt" : "SmartArt einfügen", () => {
    const rids = { dm: pptx.link(slide, RT.data, p.data), lo: pptx.link(slide, RT.layout, p.layout), qs: pptx.link(slide, RT.style, p.style), cs: pptx.link(slide, RT.colors, p.colors) };
    const drawRel = pptx.link(slide, RT.drawing, p.drawing);
    const { xml, ids } = dataModelXml(items, { layoutId: entry.id, layoutCat: entry.category, drawingRel: drawRel, scheme, style: look.style, rtl: look.rtl });
    pptx.put(p.data, xml, CT.data);
    pptx.put(p.layout, layoutXml(entry.id, entry.category, entry.name), CT.layout);
    pptx.put(p.style, quickStyleXml(look.style), CT.style);
    pptx.put(p.colors, colorsXml(scheme), CT.colors);
    // The theme's colours and format matrix: the drawing resolves the quick style's references as PowerPoint does (D-101).
    const th = app.deck.view(slide).theme;
    const theme = { colors: th.colors, fills: th.matrix.fills, lines: th.matrix.lines, effects: th.matrix.effects };
    pptx.put(p.drawing, drawingXml(drawItems(entry, items, box.w, box.h), ids, items.map((i) => i.text), scheme, { style: look.style, theme, rtl: look.rtl, width: box.w }), CT.drawing);
    const doc = pptx.doc(slide);
    const tree = S.treeOf(doc);
    id = replace ? Number(replace.getElementsByTagNameNS(PptxDom.NS_P, "cNvPr")[0]?.getAttribute("id")) || S.maxId(tree) + 1 : S.maxId(tree) + 1;
    const frame = PptxDom.fragment(doc, frameXml(id, `Diagramm ${id - 1}`, box, rids))[0];
    if (replace) replace.replaceWith(frame); else tree.append(frame);
  });
  app.deck.invalidate(slide);
  app.editor.select([id]);
  return id;
}

/** Einfügen > SmartArt: the dialog, then a diagram with three empty items in the content area. */
export async function insertSmartArt(app: PresentApp): Promise<void> {
  if (!app.slidePath()) return;
  const entry = await chooseLayout();
  if (!entry) return;
  const size = app.deck.pptx.slideSize;
  // PowerPoint's place for a new SmartArt: the content area of the Titel-und-Inhalt layout.
  createDiagram(app, entry, [{ text: "", level: 0 }, { text: "", level: 0 }, { text: "", level: 0 }], { x: 838200, y: 1825625, w: size.cx - 2 * 838200, h: 4351338 });
}

/** The diagram's items (texts and levels) from its data model, and its parts. */
export function diagramOf(app: PresentApp, id: number): { el: Element; data: string; drawing?: string; layoutId: string; scheme: string; style: string; rtl: boolean; items: DItem[] } | null {
  const path = app.slidePath();
  const el = path ? app.editor.shapeEl(id) : undefined;
  const rel = el?.getElementsByTagNameNS(NS_DGM, "relIds")[0];
  const dm = rel?.getAttributeNS(PptxDom.NS_R, "dm");
  const data = path && dm ? app.deck.pptx.target(path, dm) : undefined;
  if (!el || !path || !data || !app.deck.pptx.has(data)) return null;
  const doc = app.deck.pptx.doc(data);
  const pts = Array.from(doc.getElementsByTagNameNS(NS_DGM, "pt"));
  const docPt = pts.find((p) => p.getAttribute("type") === "doc");
  const layoutId = kid(docPt, "prSet")?.getAttribute("loTypeId") ?? SMARTART_LAYOUTS[0].id;
  const scheme = schemeOfColors(kid(docPt, "prSet")?.getAttribute("csTypeId"));
  const style = styleOfQuick(kid(docPt, "prSet")?.getAttribute("qsTypeId"));
  const dir = kid(kid(kid(docPt, "prSet"), "presLayoutVars"), "dir");
  const rtl = dir?.getAttribute("val") === "rev";
  const parent = new Map<string, { src: string; ord: number }>();
  for (const c of Array.from(doc.getElementsByTagNameNS(NS_DGM, "cxn"))) if ((c.getAttribute("type") ?? "parOf") === "parOf") parent.set(c.getAttribute("destId") ?? "", { src: c.getAttribute("srcId") ?? "", ord: Number(c.getAttribute("srcOrd") ?? 0) });
  const docId = docPt?.getAttribute("modelId") ?? "";
  const nodes = pts.filter((p) => !p.getAttribute("type") || p.getAttribute("type") === "node");
  const textOf = (pt: Element) => { const t = kid(pt, "t"); return t ? Array.from(t.getElementsByTagNameNS(PptxDom.NS_A, "t")).map((x) => x.textContent ?? "").join("") : ""; };
  // The items in tree order (children by srcOrd).
  const kidsOf = (id: string) => nodes.filter((n) => parent.get(n.getAttribute("modelId") ?? "")?.src === id).sort((a, b) => parent.get(a.getAttribute("modelId") ?? "")!.ord - parent.get(b.getAttribute("modelId") ?? "")!.ord);
  const items: DItem[] = [];
  const walk = (id: string, level: number) => { for (const n of kidsOf(id)) { items.push({ text: textOf(n), level }); walk(n.getAttribute("modelId") ?? "", level + 1); } };
  walk(docId, 0);
  const ext = doc.getElementsByTagNameNS(NS_DSP, "dataModelExt")[0];
  const drawing = ext ? app.deck.pptx.target(path, ext.getAttribute("relId") ?? "") : undefined;
  return { el, data, drawing, layoutId, scheme, style, rtl, items };
}

/**
 * Lays a diagram out again with new items or another layout: its data model and drawing are written anew (the
 * same frame, the same parts' names through a new diagram), in one undo step.
 */
export function relayout(app: PresentApp, id: number, items: DItem[], layoutId?: string, scheme?: string, look: DiagramLook = {}): void {
  const d = diagramOf(app, id);
  if (!d) return;
  const xf = kid(kid(d.el, "xfrm"), "off"), ext = kid(kid(d.el, "xfrm"), "ext");
  const box = { x: Number(xf?.getAttribute("x") ?? 0), y: Number(xf?.getAttribute("y") ?? 0), w: Number(ext?.getAttribute("cx") ?? 0), h: Number(ext?.getAttribute("cy") ?? 0) };
  createDiagram(app, entryById(layoutId ?? d.layoutId), items.length ? items : [{ text: "", level: 0 }], box, d.el, scheme ?? d.scheme, { style: look.style ?? d.style, rtl: look.rtl ?? d.rtl });
}

/** Farben ändern. */
export function setColors(app: PresentApp, id: number, scheme: string): void {
  const d = diagramOf(app, id);
  if (d) relayout(app, id, d.items, undefined, scheme);
}

/** SmartArt-Formatvorlagen (D-101). */
export function setQuickStyle(app: PresentApp, id: number, style: string): void {
  const d = diagramOf(app, id);
  if (d) relayout(app, id, d.items, undefined, undefined, { style });
}

/** Von rechts nach links: the arrangement mirrored, on and off (D-101). */
export function toggleRtl(app: PresentApp, id: number): void {
  const d = diagramOf(app, id);
  if (d) relayout(app, id, d.items, undefined, undefined, { rtl: !d.rtl });
}

/** SmartArt-Design commands on the selected diagram's items (the item at `at`, from the Textbereich or the end). */
export function editItems(app: PresentApp, id: number, op: "add" | "addBefore" | "promote" | "demote" | "up" | "down", at?: number): void {
  const d = diagramOf(app, id);
  if (!d) return;
  const items = d.items.map((i) => ({ ...i }));
  const i = Math.min(items.length - 1, Math.max(0, at ?? items.length - 1));
  const sub = (k: number) => { let e = k + 1; while (e < items.length && items[e].level > items[k].level) e++; return e; };
  if (op === "add") items.splice(sub(i), 0, { text: "", level: items[i]?.level ?? 0 });
  if (op === "addBefore") items.splice(i, 0, { text: "", level: items[i]?.level ?? 0 });
  if (op === "promote" && items[i] && items[i].level > 0) items[i].level--;
  if (op === "demote" && items[i] && i > 0 && items[i].level <= items[i - 1].level) items[i].level++;
  if (op === "up" || op === "down") {
    const end = sub(i);
    const block = items.slice(i, end);
    if (op === "up") {
      let k = i - 1;
      while (k >= 0 && items[k].level > items[i].level) k--;
      if (k < 0 || items[k].level !== items[i].level) return;
      items.splice(i, block.length); items.splice(k, 0, ...block);
    } else {
      if (end >= items.length || items[end].level !== items[i].level) return;
      const nextEnd = sub(end);
      const next = items.slice(end, nextEnd);
      items.splice(i, block.length + next.length, ...next, ...block);
    }
  }
  relayout(app, id, items);
}

/** Start > In SmartArt konvertieren: the selected text box's paragraphs (with their levels) as a diagram. */
export function convertToSmartArt(app: PresentApp, layoutId: string): void {
  const id = app.editor.sel[0];
  const el = id !== undefined ? app.editor.shapeEl(id) : undefined;
  const view = id !== undefined ? app.deck.view(app.slidePath()!).shapes.find((s) => s.id === id) : undefined;
  if (!el || !view?.text) return;
  const items = view.text.paras.map((p) => ({ text: p.runs.map((r) => r.text).join(""), level: p.lvl })).filter((x) => x.text.trim());
  const box = { x: view.x, y: view.y, w: view.w, h: view.h };
  createDiagram(app, entryById(layoutId), items.length ? items : [{ text: "", level: 0 }], box, el);
}

/** The Layouts gallery of SmartArt-Design / In SmartArt konvertieren: the layouts of a category (or the first 20). */
export function layoutsFor(category?: string, limit = 20): CatalogEntry[] {
  return SMARTART_LAYOUTS.filter((e) => !category || e.category === category).slice(0, limit);
}

export { kids };
