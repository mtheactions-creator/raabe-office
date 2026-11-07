// Shapes on a slide as PowerPoint writes them (D-070): place and size (a:xfrm, also for
// placeholders, which get their own xfrm once moved), new shapes, text boxes, pictures, tables and
// charts with PowerPoint's German names ("Rechteck 3": the id less one), copies, order, groups.

import { PptxDom, alternateShapes, phOf, spTreeOf } from "@raabe/ooxml";
import { SHAPE_NAMES } from "@raabe/text-engine";

const { kid, kids, at, num, make, fragment } = PptxDom;

export interface Xfrm { x: number; y: number; w: number; h: number; rot: number; flipH: boolean; flipV: boolean }

export const SHAPE_TAGS = ["sp", "pic", "graphicFrame", "grpSp", "cxnSp"];

export function treeOf(doc: Document): Element {
  return spTreeOf(doc)!;
}

export function shapeEls(tree: Element): Element[] {
  // Shapes inside mc:AlternateContent (equations) are found where they are drawn from (D-072).
  return kids(tree).flatMap((c) => (SHAPE_TAGS.includes(c.localName) ? [c] : c.localName === "AlternateContent" ? alternateShapes(c) : []));
}

export function cNvPr(el: Element): Element | undefined {
  const nv = kids(el).find((c) => c.localName.startsWith("nv"));
  return kid(nv, "cNvPr");
}

export function shapeId(el: Element): number {
  return num(cNvPr(el), "id") ?? 0;
}

/** The element of shape `id` anywhere in the tree (also inside groups). */
export function findShape(tree: Element, id: number): Element | undefined {
  for (const s of shapeEls(tree)) {
    if (shapeId(s) === id) return s;
    if (s.localName === "grpSp") { const hit = findShape(s, id); if (hit) return hit; }
  }
  return undefined;
}

export function maxId(tree: Element): number {
  let m = 1;
  const walk = (t: Element) => { for (const s of shapeEls(t)) { m = Math.max(m, shapeId(s)); if (s.localName === "grpSp") walk(s); } };
  walk(tree);
  return m;
}

function xfrmEl(el: Element, create: boolean): Element | undefined {
  if (el.localName === "graphicFrame") {
    let x = kid(el, "xfrm");
    if (!x && create) { x = make(el.ownerDocument, "p:xfrm"); el.insertBefore(x, kid(el, "graphic") ?? null); }
    return x;
  }
  const prName = el.localName === "grpSp" ? "grpSpPr" : "spPr";
  let pr = kid(el, prName);
  if (!pr && create) { pr = make(el.ownerDocument, `p:${prName}`); el.insertBefore(pr, kids(el)[1] ?? null); }
  let x = kid(pr, "xfrm");
  if (!x && create && pr) { x = make(el.ownerDocument, "a:xfrm"); pr.insertBefore(x, pr.firstChild); }
  return x;
}

/** Writes place, size, rotation and flips (EMU, degrees) as PowerPoint does: whole EMU, rot in 60000ths. */
export function setXfrm(el: Element, t: Partial<Xfrm>): void {
  const x = xfrmEl(el, true)!;
  const doc = el.ownerDocument;
  let off = kid(x, "off"), ext = kid(x, "ext");
  if (!off) { off = make(doc, "a:off"); x.insertBefore(off, x.firstChild); }
  if (!ext) { ext = make(doc, "a:ext"); off.after(ext); }
  if (t.x !== undefined) off.setAttribute("x", String(Math.round(t.x)));
  if (t.y !== undefined) off.setAttribute("y", String(Math.round(t.y)));
  if (t.w !== undefined) ext.setAttribute("cx", String(Math.max(0, Math.round(t.w))));
  if (t.h !== undefined) ext.setAttribute("cy", String(Math.max(0, Math.round(t.h))));
  if (t.rot !== undefined) {
    const r = Math.round((((t.rot % 360) + 360) % 360) * 60000);
    if (r) x.setAttribute("rot", String(r)); else x.removeAttribute("rot");
  }
  if (t.flipH !== undefined) { if (t.flipH) x.setAttribute("flipH", "1"); else x.removeAttribute("flipH"); }
  if (t.flipV !== undefined) { if (t.flipV) x.setAttribute("flipV", "1"); else x.removeAttribute("flipV"); }
}

/** For groups: moving or sizing the group keeps chOff/chExt, so the members follow (PowerPoint). */
export function groupChildFrame(el: Element): { chOff: [number, number]; chExt: [number, number] } | null {
  const x = xfrmEl(el, false);
  const co = kid(x, "chOff"), ce = kid(x, "chExt");
  return co && ce ? { chOff: [num(co, "x") ?? 0, num(co, "y") ?? 0], chExt: [num(ce, "cx") ?? 0, num(ce, "cy") ?? 0] } : null;
}

// ----- new shapes -------------------------------------------------------------------------------------


const xfrmXml = (x: number, y: number, w: number, h: number) => `<a:xfrm><a:off x="${Math.round(x)}" y="${Math.round(y)}"/><a:ext cx="${Math.round(w)}" cy="${Math.round(h)}"/></a:xfrm>`;

/** PowerPoint 365's style of a new shape: Akzent 1 fill, its 15 % shade as outline, white text. */
export const SHAPE_STYLE = '<p:style><a:lnRef idx="2"><a:schemeClr val="accent1"><a:shade val="15000"/></a:schemeClr></a:lnRef><a:fillRef idx="1"><a:schemeClr val="accent1"/></a:fillRef>' +
  '<a:effectRef idx="0"><a:schemeClr val="accent1"/></a:effectRef><a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef></p:style>';
export const LINE_STYLE = '<p:style><a:lnRef idx="1"><a:schemeClr val="accent1"/></a:lnRef><a:fillRef idx="0"><a:schemeClr val="accent1"/></a:fillRef>' +
  '<a:effectRef idx="0"><a:schemeClr val="accent1"/></a:effectRef><a:fontRef idx="minor"><a:schemeClr val="tx1"/></a:fontRef></p:style>';

export const isLine = (prst: string) => prst === "line" || /Connector\d$/.test(prst);

export function shapeName(prst: string): string {
  return SHAPE_NAMES[prst] ?? (prst === "line" ? "Gerader Verbinder" : "Form");
}

/** A new shape element (p:sp, or p:cxnSp for lines and connectors) owned by `doc`. */
export function newShape(doc: Document, id: number, prst: string, box: Xfrm, custom?: string): Element {
  const name = `${custom ? "Freihandform: Form" : shapeName(prst)} ${id - 1}`;
  const geom = custom ?? `<a:prstGeom prst="${prst}"><a:avLst/></a:prstGeom>`;
  const flip = `${box.flipH ? ' flipH="1"' : ""}${box.flipV ? ' flipV="1"' : ""}${box.rot ? ` rot="${Math.round(box.rot * 60000)}"` : ""}`;
  const xf = `<a:xfrm${flip}><a:off x="${Math.round(box.x)}" y="${Math.round(box.y)}"/><a:ext cx="${Math.round(box.w)}" cy="${Math.round(box.h)}"/></a:xfrm>`;
  if (isLine(prst)) {
    return fragment(doc, `<p:cxnSp><p:nvCxnSpPr><p:cNvPr id="${id}" name="${PptxDom.esc(name)}"/><p:cNvCxnSpPr/><p:nvPr/></p:nvCxnSpPr><p:spPr>${xf}${geom}</p:spPr>${LINE_STYLE}</p:cxnSp>`)[0];
  }
  return fragment(doc, `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${PptxDom.esc(name)}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr>${xf}${geom}</p:spPr>${SHAPE_STYLE}` +
    `<p:txBody><a:bodyPr rtlCol="0" anchor="ctr"/><a:lstStyle/><a:p><a:pPr algn="ctr"/><a:endParaRPr lang="de-DE"/></a:p></p:txBody></p:sp>`)[0];
}

/** Einfügen > Textfeld: a click gives a box that grows with its text (no wrap), a drag one of that width. */
export function newTextBox(doc: Document, id: number, x: number, y: number, w: number | null, h: number): Element {
  const wrap = w === null ? ' wrap="none"' : ' wrap="square"';
  return fragment(doc, `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="Textfeld ${id - 1}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr>${xfrmXml(x, y, w ?? 182880, h)}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr>` +
    `<p:txBody><a:bodyPr${wrap} rtlCol="0"><a:spAutoFit/></a:bodyPr><a:lstStyle/><a:p><a:endParaRPr lang="de-DE" dirty="0"/></a:p></p:txBody></p:sp>`)[0];
}

export function newPicture(doc: Document, id: number, rid: string, box: { x: number; y: number; w: number; h: number }, name = "Grafik"): Element {
  return fragment(doc, `<p:pic><p:nvPicPr><p:cNvPr id="${id}" name="${PptxDom.esc(name)} ${id - 1}"/><p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr/></p:nvPicPr>` +
    `<p:blipFill><a:blip r:embed="${rid}"/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr>${xfrmXml(box.x, box.y, box.w, box.h)}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>`)[0];
}

/** Einfügen > Tabelle: rows x cols in PowerPoint's default style, 10 160 000 EMU wide like PowerPoint's. */
export function newTable(doc: Document, id: number, rows: number, cols: number, box: { x: number; y: number; w: number }, styleId: string): Element {
  const cw = Math.round(box.w / cols), rh = 370840;
  const cell = '<a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="de-DE" dirty="0"/></a:p></a:txBody><a:tcPr/></a:tc>';
  const grid = Array.from({ length: cols }, () => `<a:gridCol w="${cw}"/>`).join("");
  const trs = Array.from({ length: rows }, () => `<a:tr h="${rh}">${cell.repeat(cols)}</a:tr>`).join("");
  return fragment(doc, `<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="${id}" name="Tabelle ${id - 1}"/><p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr><p:nvPr/></p:nvGraphicFramePr>` +
    `<p:xfrm><a:off x="${Math.round(box.x)}" y="${Math.round(box.y)}"/><a:ext cx="${cw * cols}" cy="${rh * rows}"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">` +
    `<a:tbl><a:tblPr firstRow="1" bandRow="1"><a:tableStyleId>${styleId}</a:tableStyleId></a:tblPr><a:tblGrid>${grid}</a:tblGrid>${trs}</a:tbl></a:graphicData></a:graphic></p:graphicFrame>`)[0];
}

export function newChartFrame(doc: Document, id: number, rid: string, box: { x: number; y: number; w: number; h: number }): Element {
  return fragment(doc, `<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="${id}" name="Diagramm ${id - 1}"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>` +
    `<p:xfrm><a:off x="${Math.round(box.x)}" y="${Math.round(box.y)}"/><a:ext cx="${Math.round(box.w)}" cy="${Math.round(box.h)}"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">` +
    `<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" r:id="${rid}"/></a:graphicData></a:graphic></p:graphicFrame>`)[0];
}

/** Gives a copied element new ids (its own and its members') and names numbered by them. */
export function renumber(el: Element, next: () => number): void {
  const walk = (e: Element) => {
    const c = cNvPr(e);
    if (c) {
      const id = next();
      c.setAttribute("id", String(id));
      const name = c.getAttribute("name") ?? "";
      if (/ \d+$/.test(name)) c.setAttribute("name", name.replace(/ \d+$/, ` ${id - 1}`));
      // creationId extensions identify the original; copies drop them as PowerPoint makes new ones.
      for (const ext of Array.from(c.getElementsByTagNameNS("http://schemas.microsoft.com/office/drawing/2014/main", "creationId"))) ext.closest("ext")?.remove();
    }
    if (e.localName === "grpSp") for (const k of shapeEls(e)) walk(k);
  };
  walk(el);
}

/** The text body of a shape (p:txBody), when it can hold text. */
export function txBodyOf(el: Element): Element | undefined {
  return kid(el, "txBody");
}

/** Gives a shape that has none a text body (typing into a picture-less shape). */
export function ensureTxBody(el: Element): Element | undefined {
  if (el.localName !== "sp") return undefined;
  let b = kid(el, "txBody");
  if (!b) {
    b = fragment(el.ownerDocument, '<p:txBody><a:bodyPr rtlCol="0" anchor="ctr"/><a:lstStyle/><a:p><a:pPr algn="ctr"/><a:endParaRPr lang="de-DE"/></a:p></p:txBody>')[0];
    el.append(b);
  }
  return b;
}

export function isPlaceholder(el: Element): boolean {
  return !!phOf(el);
}

/** bodyPr of a shape's text, created when missing. */
export function bodyPrOf(el: Element): Element | undefined {
  const b = kid(el, "txBody");
  if (!b) return undefined;
  let pr = kid(b, "bodyPr");
  if (!pr) { pr = make(el.ownerDocument, "a:bodyPr"); b.insertBefore(pr, b.firstChild); }
  return pr;
}

/** Writes PowerPoint's Autofit result (normAutofit fontScale / lnSpcReduction) into the slide's bodyPr. */
export function setAutofit(el: Element, fit: { fontScale: number; lnSpcReduction: number }): boolean {
  const pr = bodyPrOf(el);
  if (!pr) return false;
  const fs = Math.round(fit.fontScale * 100000), ls = Math.round(fit.lnSpcReduction * 100000);
  let na = kid(pr, "normAutofit");
  const cur = na ? [num(na, "fontScale") ?? 100000, num(na, "lnSpcReduction") ?? 0] : [100000, 0];
  if (cur[0] === fs && cur[1] === ls && na) return false;
  if (!na) {
    if (fs === 100000 && ls === 0) return false;
    na = make(el.ownerDocument, "a:normAutofit");
    pr.insertBefore(na, kids(pr).find((c) => ["scene3d", "sp3d", "flatTx", "extLst"].includes(c.localName)) ?? null);
  }
  if (fs !== 100000) na.setAttribute("fontScale", String(fs)); else na.removeAttribute("fontScale");
  if (ls) na.setAttribute("lnSpcReduction", String(ls)); else na.removeAttribute("lnSpcReduction");
  return true;
}

export { at };
