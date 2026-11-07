// Einfügen > Zoom > Folienzoom (D-072). Recorded in PowerPoint (probe_powerpoint_zoom.py): a zoom is
// mc:AlternateContent with a p:graphicFrame (pslz:sldZm: the target's sldId and creationId, the preview picture
// in p166:blipFill, a 0.25 pt light grey line) for PowerPoint 2019 and later, and as Fallback a picture with a
// slide-jump link for older readers. New zooms are a quarter of the slide wide, in its middle. The preview
// picture is the target slide drawn by us (our slide's DOM in an SVG foreignObject, onto a canvas).
// Zusammenfassungszoom and Abschnittszoom are not built (greyed).

import { h, showDialog, type MenuEntry } from "@raabe/shell";
import { PptxDom, slideView } from "@raabe/ooxml";
import type { PresentApp } from "./app";
import { renderSlide } from "./render/slide";
import { EMU_PX } from "./render/text";
import * as S from "./shapeops";

const { kid } = PptxDom;
const NS_P14 = "http://schemas.microsoft.com/office/powerpoint/2010/main";
const RT_IMAGE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";
const RT_SLIDE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide";

export function zoomMenu(app: PresentApp): MenuEntry[] {
  return [
    { label: "Zusammenfassungszoom", disabled: true, run: () => {} },
    { label: "Abschnittszoom", disabled: true, run: () => {} },
    { label: "Folienzoom", run: () => void folienzoomDialog(app) },
  ];
}

/** The slide's title text (for the dialog's list). */
function titleOf(app: PresentApp, path: string): string {
  const v = app.deck.view(path);
  const t = v.shapes.find((s) => s.ph && (s.ph.type === "title" || s.ph.type === "ctrTitle"));
  return t?.text ? t.text.paras.map((p) => p.runs.map((r) => r.text).join("")).join(" ").trim() : "";
}

async function folienzoomDialog(app: PresentApp): Promise<void> {
  const boxes = app.deck.slides.map((s, i) => {
    const c = h("input", { type: "checkbox" }) as HTMLInputElement;
    return { c, i, el: h("label", { class: "dlg-check" }, c, ` ${i + 1}. ${titleOf(app, s.path) || `Folie ${i + 1}`}`) };
  });
  const ok = await showDialog({ title: "Folienzoom einfügen", body: h("div", { class: "form", style: "max-height:360px;overflow:auto;min-width:300px" }, h("b", null, "Wählen Sie die Folien aus, die Sie einfügen möchten."), ...boxes.map((b) => b.el)),
    buttons: [{ label: "Einfügen", value: true, primary: true }, { label: "Abbrechen", value: false }], cancel: false });
  if (!ok) return;
  const picked = boxes.filter((b) => b.c.checked).map((b) => b.i);
  if (picked.length) await insertSlideZooms(app, picked);
}

/** Folienzoom for each slide index, on the current slide. */
export async function insertSlideZooms(app: PresentApp, targets: number[]): Promise<void> {
  const path = app.slidePath();
  if (!path) return;
  const pptx = app.deck.pptx, size = pptx.slideSize;
  const w = Math.round(size.cx / 4), hh = Math.round(size.cy / 4);
  const pngs = await Promise.all(targets.map((i) => slidePng(app, app.deck.slides[i].path, 480)));
  const ids: number[] = [];
  app.editor.edit("Folienzoom", () => {
    const doc = pptx.doc(path), tree = S.treeOf(doc);
    targets.forEach((ti, k) => {
      const target = app.deck.slides[ti];
      const png = pngs[k];
      const media = pptx.freeName("ppt/media/image#.png");
      pptx.put(media, png ?? new Uint8Array());
      pptx.setType(media, "image/png", true);
      const img = pptx.link(path, RT_IMAGE, media);
      const sld = pptx.link(path, RT_SLIDE, target.path);
      const id = S.maxId(tree) + 1;
      ids.push(id);
      // Several zooms step down and right from the middle, 12 pt apart.
      const x = Math.round((size.cx - w) / 2) + k * 152400, y = Math.round((size.cy - hh) / 2) + k * 152400;
      tree.append(zoomElement(doc, { id, name: `Folienzoom ${id - 1}`, sldId: target.id, cId: creationId(app, target.path), img, sld, x, y, w, h: hh }));
    });
  });
  app.editor.select(ids);
}

/** The target slide's creationId (p14:creationId), made when it has none (PowerPoint gives every slide one). */
function creationId(app: PresentApp, path: string): number {
  const root = app.deck.pptx.doc(path).documentElement;
  const found = root.getElementsByTagNameNS(NS_P14, "creationId")[0];
  if (found) return Number(found.getAttribute("val"));
  const val = Math.floor(Math.random() * 4294967295);
  app.deck.change("Folienzoom", [path], () => {
    let ext = Array.from(root.children).find((c) => c.localName === "extLst");
    if (!ext) { ext = PptxDom.make(root.ownerDocument, "p:extLst"); root.append(ext); }
    ext.append(PptxDom.fragment(root.ownerDocument, `<p:ext xmlns:p="${PptxDom.NS_P}" uri="{BB962C8B-B14F-4D97-AF65-F5344CB8AC3E}"><p14:creationId xmlns:p14="${NS_P14}" val="${val}"/></p:ext>`)[0]);
  });
  return val;
}

function guid(): string {
  const hex = Array.from(crypto.getRandomValues(new Uint8Array(16)), (b) => b.toString(16).padStart(2, "0")).join("").toUpperCase();
  return `{${hex.slice(0, 8)}-${hex.slice(8, 12)}-${hex.slice(12, 16)}-${hex.slice(16, 20)}-${hex.slice(20)}}`;
}

function zoomElement(doc: Document, z: { id: number; name: string; sldId: number; cId: number; img: string; sld: string; x: number; y: number; w: number; h: number }): Element {
  const name = PptxDom.esc(z.name);
  const ln = `<a:ln w="3175"><a:solidFill><a:prstClr val="ltGray"/></a:solidFill></a:ln>`;
  const xml = `<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:p="${PptxDom.NS_P}" xmlns:a="${PptxDom.NS_A}" xmlns:r="${PptxDom.NS_R}">` +
    `<mc:Choice xmlns:pslz="http://schemas.microsoft.com/office/powerpoint/2016/slidezoom" Requires="pslz"><p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="${z.id}" name="${name}"/><p:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></p:cNvGraphicFramePr><p:nvPr/></p:nvGraphicFramePr>` +
    `<p:xfrm><a:off x="${z.x}" y="${z.y}"/><a:ext cx="${z.w}" cy="${z.h}"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/powerpoint/2016/slidezoom"><pslz:sldZm><pslz:sldZmObj sldId="${z.sldId}" cId="${z.cId}">` +
    `<pslz:zmPr id="${guid()}" returnToParent="0" transitionDur="1000"><p166:blipFill xmlns:p166="http://schemas.microsoft.com/office/powerpoint/2016/6/main"><a:blip r:embed="${z.img}"/><a:stretch><a:fillRect/></a:stretch></p166:blipFill>` +
    `<p166:spPr xmlns:p166="http://schemas.microsoft.com/office/powerpoint/2016/6/main"><a:xfrm><a:off x="0" y="0"/><a:ext cx="${z.w}" cy="${z.h}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom>${ln}</p166:spPr></pslz:zmPr></pslz:sldZmObj></pslz:sldZm></a:graphicData></a:graphic></p:graphicFrame></mc:Choice>` +
    `<mc:Fallback><p:pic><p:nvPicPr><p:cNvPr id="${z.id}" name="${name}"><a:hlinkClick r:id="${z.sld}" action="ppaction://hlinksldjump"/></p:cNvPr><p:cNvPicPr><a:picLocks noGrp="1" noRot="1" noChangeAspect="1" noMove="1" noResize="1" noEditPoints="1" noAdjustHandles="1" noChangeArrowheads="1" noChangeShapeType="1"/></p:cNvPicPr><p:nvPr/></p:nvPicPr>` +
    `<p:blipFill><a:blip r:embed="${z.img}"/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr><a:xfrm><a:off x="${z.x}" y="${z.y}"/><a:ext cx="${z.w}" cy="${z.h}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom>${ln}</p:spPr></p:pic></mc:Fallback></mc:AlternateContent>`;
  return PptxDom.fragment(doc, xml)[0];
}

// ----- a slide as a picture -------------------------------------------------------------------------------

async function dataUrl(url: string): Promise<string> {
  const blob = await (await fetch(url)).blob();
  return await new Promise((res) => { const r = new FileReader(); r.onload = () => res(String(r.result)); r.readAsDataURL(blob); });
}

/** The slide drawn by our renderer as a PNG `width` pixels wide (null when the browser cannot draw it). */
export async function slidePng(app: PresentApp, path: string, width: number): Promise<Uint8Array | null> {
  const pptx = app.deck.pptx, size = pptx.slideSize;
  const W = size.cx / EMU_PX, H = size.cy / EMU_PX;
  const r = renderSlide(slideView(pptx, path), size, { editing: false, image: (p) => app.image(p), pptx });
  let html = new XMLSerializer().serializeToString(r.el);
  // Pictures as data: an SVG image may not load anything from outside.
  const urls = [...new Set(html.match(/blob:[^"')\s&]+/g) ?? [])];
  for (const u of urls) { try { html = html.split(u).join(await dataUrl(u)); } catch { /* left out */ } }
  const css = Array.from(document.styleSheets).flatMap((ss) => { try { return Array.from(ss.cssRules).map((x) => x.cssText); } catch { return []; } }).join("\n");
  const scale = width / W;
  const svg = `<svg xmlns="http://www.w3.org/2000/svg" width="${Math.round(W * scale)}" height="${Math.round(H * scale)}"><foreignObject width="${W}" height="${H}" transform="scale(${scale})">` +
    `<div xmlns="http://www.w3.org/1999/xhtml" style="width:${W}px;height:${H}px;position:relative;overflow:hidden"><style>${css.replace(/</g, "\\3c ")}</style>${html}</div></foreignObject></svg>`;
  const img = new Image();
  const ok = await new Promise<boolean>((res) => { img.onload = () => res(true); img.onerror = () => res(false); img.src = "data:image/svg+xml;charset=utf-8," + encodeURIComponent(svg); });
  if (!ok) return null;
  const c = document.createElement("canvas");
  c.width = Math.round(W * scale); c.height = Math.round(H * scale);
  const g = c.getContext("2d");
  if (!g) return null;
  g.fillStyle = "#fff";
  g.fillRect(0, 0, c.width, c.height);
  try {
    g.drawImage(img, 0, 0);
    const blob = await new Promise<Blob | null>((res) => c.toBlob(res, "image/png"));
    return blob ? new Uint8Array(await blob.arrayBuffer()) : null;
  } catch { return null; }
}

/** After a zoom moved or changed size: its inner picture frame and the Fallback picture follow (PowerPoint keeps them equal). */
export function syncZoom(frame: Element): void {
  const zm = frame.getElementsByTagNameNS("http://schemas.microsoft.com/office/powerpoint/2016/slidezoom", "zmPr")[0];
  if (!zm) return;
  const xf = kid(frame, "xfrm"), off = kid(xf, "off"), ext = kid(xf, "ext");
  const inner = zm.getElementsByTagNameNS(PptxDom.NS_A, "ext")[0];
  if (inner && ext) { inner.setAttribute("cx", ext.getAttribute("cx") ?? "0"); inner.setAttribute("cy", ext.getAttribute("cy") ?? "0"); }
  const fb = frame.parentElement?.nextElementSibling;
  const pxf = fb?.localName === "Fallback" ? fb.getElementsByTagNameNS(PptxDom.NS_A, "xfrm")[0] : undefined;
  if (pxf && off && ext) {
    for (const [from, name] of [[off, "off"], [ext, "ext"]] as const) {
      const to = kid(pxf, name);
      for (const at of Array.from(from.attributes)) to?.setAttribute(at.name, at.value);
    }
  }
}
