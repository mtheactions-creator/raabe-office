// Bildformat > Bildformatvorlagen (D-072): PowerPoint's 28 entries, named as recorded. Each sets the
// picture's outline (width, colour, compound line), its shape and its effects (shadow, reflection, soft
// edges) in p:spPr, as PowerPoint's styles do; the bevels and 3-D perspectives of some entries are left
// out (our renderer draws no 3-D), so those get only their frame and shadow.

import { h } from "@raabe/shell";
import { PptxDom } from "@raabe/ooxml";
import type { PresentApp } from "./app";

const { kid, kids, fragment } = PptxDom;

const SHADOW = '<a:outerShdw blurRad="76200" dist="38100" dir="2700000" algn="tl" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="40000"/></a:srgbClr></a:outerShdw>';
const DROP = '<a:outerShdw blurRad="0" dist="76200" dir="2700000" algn="tl" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="50000"/></a:srgbClr></a:outerShdw>';
const CENTER = '<a:outerShdw blurRad="190500" algn="ctr" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="70000"/></a:srgbClr></a:outerShdw>';
const REFLECT = '<a:reflection blurRad="12700" stA="38000" endPos="28000" dist="5000" dir="5400000" sy="-100000" algn="bl" rotWithShape="0"/>';
const SOFT = '<a:softEdge rad="112500"/>';

interface PicStyle { name: string; geom: string; line?: { w: number; color: string; cmpd?: string }; fx: string[] }

const W = "FFFFFF", B = "000000", M = "BFBFBF";
export const PIC_STYLES: PicStyle[] = [
  { name: "Einfacher Rahmen, weiß", geom: "rect", line: { w: 7, color: W }, fx: [SHADOW] },
  { name: "Abgeschrägt matt, weiß", geom: "rect", line: { w: 9, color: W }, fx: [SHADOW] },
  { name: "Metallrahmen", geom: "rect", line: { w: 9, color: M, cmpd: "thickThin" }, fx: [SHADOW] },
  { name: "Schlagschattenrechteck", geom: "rect", fx: [DROP] },
  { name: "Reflektiertes abgerundetes Rechteck", geom: "roundRect", fx: [REFLECT] },
  { name: "Rechteck mit weichen Kanten", geom: "rect", fx: [SOFT] },
  { name: "Doppelrahmen, schwarz", geom: "rect", line: { w: 6, color: B, cmpd: "dbl" }, fx: [] },
  { name: "Dick matt, schwarz", geom: "rect", line: { w: 9, color: B }, fx: [] },
  { name: "Einfacher Rahmen, schwarz", geom: "rect", line: { w: 6, color: B }, fx: [] },
  { name: "Abgeschrägtes Oval, schwarz", geom: "ellipse", line: { w: 3, color: B }, fx: [SHADOW] },
  { name: "Verbundrahmen, schwarz", geom: "rect", line: { w: 6, color: B, cmpd: "thickThin" }, fx: [] },
  { name: "Moderater Rahmen, schwarz", geom: "rect", line: { w: 4.5, color: B }, fx: [SHADOW] },
  { name: "Schattenrechteck zentriert", geom: "rect", fx: [CENTER] },
  { name: "Abgerundete diagonale Ecken, weiß", geom: "round2DiagRect", line: { w: 7, color: W }, fx: [SHADOW] },
  { name: "Diagonal liegende Ecken schneiden, weiß", geom: "snip2DiagRect", line: { w: 7, color: W }, fx: [] },
  { name: "Moderater Rahmen, weiß", geom: "rect", line: { w: 4.5, color: W }, fx: [SHADOW] },
  { name: "Gedreht, weiß", geom: "rect", line: { w: 7, color: W }, fx: [SHADOW] },
  { name: "Perspektivischer Schatten, weiß", geom: "rect", line: { w: 7, color: W }, fx: [DROP] },
  { name: "Perspektive oberhalb, weiß", geom: "rect", line: { w: 7, color: W }, fx: [SHADOW] },
  { name: "Oval mit weichen Kanten", geom: "ellipse", fx: [SOFT] },
  { name: "Abgeschrägtes Rechteck", geom: "rect", fx: [SHADOW] },
  { name: "Abgeschrägte Perspektive", geom: "rect", fx: [SHADOW] },
  { name: "Reflektierte Perspektive rechts", geom: "rect", fx: [REFLECT] },
  { name: "Abgeschrägte Perspektive links, weiß", geom: "rect", line: { w: 7, color: W }, fx: [SHADOW] },
  { name: "Reflektierte Abschrägung, schwarz", geom: "rect", line: { w: 3, color: B }, fx: [REFLECT] },
  { name: "Reflektierte Abschrägung, weiß", geom: "rect", line: { w: 3, color: W }, fx: [REFLECT] },
  { name: "Metallenes abgerundetes Rechteck", geom: "roundRect", line: { w: 9, color: M, cmpd: "thickThin" }, fx: [SHADOW] },
  { name: "Metallenes Oval", geom: "ellipse", line: { w: 9, color: M, cmpd: "thickThin" }, fx: [SHADOW] },
];

export function applyPicStyle(app: PresentApp, s: PicStyle): void {
  const ed = app.editor;
  const views = ed.selectedViews.filter((v) => v.kind === "pic");
  if (!views.length) return;
  ed.edit("Bildformatvorlage", () => {
    for (const v of views) {
      const spPr = kid(ed.shapeEl(v.id), "spPr");
      if (!spPr) continue;
      for (const x of kids(spPr)) if (["prstGeom", "custGeom", "ln", "effectLst", "effectDag", "scene3d", "sp3d"].includes(x.localName)) x.remove();
      const xml = `<a:prstGeom prst="${s.geom}"><a:avLst/></a:prstGeom>` +
        (s.line ? `<a:ln w="${Math.round(s.line.w * 12700)}"${s.line.cmpd ? ` cmpd="${s.line.cmpd}"` : ""}><a:solidFill><a:srgbClr val="${s.line.color}"/></a:solidFill></a:ln>` : "") +
        (s.fx.length ? `<a:effectLst>${s.fx.join("")}</a:effectLst>` : "");
      const xfrm = kid(spPr, "xfrm");
      let after: Element | null = xfrm ?? null;
      for (const el of fragment(spPr.ownerDocument, xml)) {
        // Schema order: xfrm, geometry, fill, ln, effects. A fill of the picture's frame stays after the geometry.
        if (el.localName === "prstGeom") { spPr.insertBefore(el, after ? after.nextSibling : spPr.firstChild); after = el; continue; }
        const fill = kids(spPr).filter((x) => x.localName.endsWith("Fill")).pop();
        const ref = el.localName === "ln" ? (fill ?? after) : (kid(spPr, "ln") ?? fill ?? after);
        spPr.insertBefore(el, ref ? ref.nextSibling : spPr.firstChild);
      }
    }
  });
}

function picTileImg(s: PicStyle): HTMLElement {
  const img = h("div", { class: "ps-img" });
  const radius = s.geom === "ellipse" ? "50%" : s.geom === "roundRect" ? "8px" : s.geom === "round2DiagRect" ? "8px 0" : "0";
  img.style.cssText = `border-radius:${radius};border:${s.line ? `${Math.max(1, s.line.w / 2)}px ${s.line.cmpd === "dbl" ? "double" : "solid"} #${s.line.color}` : "none"};` +
    `box-shadow:${s.fx.includes(DROP) ? "3px 3px 0 rgba(0,0,0,.5)" : s.fx.includes(CENTER) ? "0 0 6px rgba(0,0,0,.7)" : s.fx.includes(SHADOW) ? "2px 2px 3px rgba(0,0,0,.4)" : "none"};` +
    (s.fx.includes(SOFT) ? "filter:blur(1px);" : "") + (s.fx.includes(REFLECT) ? "-webkit-box-reflect:below 1px linear-gradient(transparent 55%, rgba(0,0,0,.35));" : "");
  return img;
}

/** The gallery: a grey picture in each style. */
export function picStylesPanel(app: PresentApp, close: () => void): HTMLElement {
  const wrap = h("div", { class: "menu pic-styles" });
  const grid = h("div", { class: "ps-grid" });
  for (const s of PIC_STYLES) {
    const t = h("button", { class: "ps-tile", "aria-label": s.name, "data-tip": s.name, onclick: () => { close(); applyPicStyle(app, s); } });
    t.append(picTileImg(s));
    grid.append(t);
  }
  wrap.append(grid);
  return wrap;
}

/** Bildformat > Bildformatvorlagen in the ribbon (PowerPoint shows them there). */
export function picStyleTiles(app: PresentApp, visible = 3): import("@raabe/shell").Control {
  return {
    type: "tiles", cmd: "pic.quickStyles", key: "S", count: PIC_STYLES.length, visible, size: [48, 44], cls: "ps-tiles",
    render: (i) => h("div", { class: "ps-rtile" }, picTileImg(PIC_STYLES[i])),
    tip: (i) => PIC_STYLES[i].name,
    pick: (i) => applyPicStyle(app, PIC_STYLES[i]),
    redraw: () => "pic",
    panel: (close) => picStylesPanel(app, close),
  };
}
