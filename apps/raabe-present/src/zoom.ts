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
    { label: "Zusammenfassungszoom", run: () => void zusammenfassungszoomDialog(app) },
    { label: "Abschnittszoom", disabled: !app.deck.sections().length, run: () => void abschnittszoomDialog(app) },
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

async function zusammenfassungszoomDialog(app: PresentApp): Promise<void> {
  const boxes = app.deck.slides.map((s, i) => {
    const c = h("input", { type: "checkbox" }) as HTMLInputElement;
    return { c, i, el: h("label", { class: "dlg-check" }, c, ` ${i + 1}. ${titleOf(app, s.path) || `Folie ${i + 1}`}`) };
  });
  const ok = await showDialog({ title: "Zusammenfassungszoom einfügen", body: h("div", { class: "form", style: "max-height:360px;overflow:auto;min-width:300px" }, h("b", null, "Wählen Sie die Folien aus, die Sie einfügen möchten."), ...boxes.map((b) => b.el)),
    buttons: [{ label: "Einfügen", value: true, primary: true }, { label: "Abbrechen", value: false }], cancel: false });
  if (!ok) return;
  const picked = boxes.filter((b) => b.c.checked).map((b) => b.i);
  if (picked.length) await insertSummaryZoom(app, picked);
}

/**
 * Zusammenfassungszoom, as PowerPoint makes it (recorded): a new "Titel und Inhalt" slide before the first chosen
 * slide; sections "Standardabschnitt" (the slides before), "Zusammenfassungsabschnitt" (the new slide) and one per
 * chosen slide named after its title (up to the next chosen slide); the content placeholder's place holds one
 * frame with a tile per section in a grid (PowerPoint's measured spacing: tiles 3.75 % of their width apart, the
 * grid centred in 91.7 % of the frame).
 */
export async function insertSummaryZoom(app: PresentApp, chosen: number[]): Promise<void> {
  const deck = app.deck, pptx = deck.pptx;
  const picks = [...new Set(chosen)].sort((a, b) => a - b);
  const first = picks[0];
  const pngs = await Promise.all(picks.map((i) => slidePng(app, deck.slides[i].path, 480)));
  const layout = deck.layouts().find((l) => l.name === "Titel und Inhalt")?.path ?? deck.layouts()[1]?.path;
  const path = deck.addSlide(first - 1, layout);
  const ids = deck.slides.map((s) => s.id);
  const sumId = deck.slides[first].id;
  // Sections: the chosen slides moved one on by the new slide.
  const starts = picks.map((i) => i + 1);
  const guidOf = () => guid();
  const secs: { name: string; id: string; slides: number[] }[] = [];
  if (first > 0) secs.push({ name: "Standardabschnitt", id: guidOf(), slides: ids.slice(0, first) });
  secs.push({ name: "Zusammenfassungsabschnitt", id: guidOf(), slides: [sumId] });
  // Slides between the summary and the first chosen one stay with the summary's section.
  secs[secs.length - 1].slides.push(...ids.slice(first + 1, starts[0]));
  starts.forEach((st, k) => {
    const end = k + 1 < starts.length ? starts[k + 1] : ids.length;
    secs.push({ name: titleOf(app, deck.slides[st].path) || `Folie ${st}`, id: guidOf(), slides: ids.slice(st, end) });
  });
  deck.setSections(secs);
  app.setCurrent(first);
  const size = pptx.slideSize;
  app.editor.edit("Zusammenfassungszoom", () => {
    const doc = pptx.doc(path), tree = S.treeOf(doc);
    // The content placeholder's place (from the layout), which the zoom takes.
    const v = deck.view(path);
    const body = v.shapes.find((s) => s.ph && s.ph.type !== "title" && s.ph.type !== "ctrTitle");
    const box = body ? { x: body.x, y: body.y, w: body.w, h: body.h } : { x: size.cx * 0.07, y: size.cy * 0.27, w: size.cx * 0.86, h: size.cy * 0.63 };
    if (body) S.shapeEls(tree).find((e) => Number(e.getElementsByTagNameNS(PptxDom.NS_P, "cNvPr")[0]?.getAttribute("id")) === body.id)?.remove();
    const n = picks.length, cols = Math.ceil(Math.sqrt(n)), rows = Math.ceil(n / cols), gapF = 0.0375, fill = 0.9169;
    const aspect = size.cy / size.cx;
    const tw = Math.min((box.w * fill) / (cols + (cols - 1) * gapF), (box.h * fill) / (rows + (rows - 1) * gapF) / aspect);
    const th = tw * aspect, gap = tw * gapF;
    const gx = (box.w - (cols * tw + (cols - 1) * gap)) / 2, gy = (box.h - (rows * th + (rows - 1) * gap)) / 2;
    const id = S.maxId(tree) + 1;
    const tiles = picks.map((si, k) => {
      const media = pptx.freeName("ppt/media/image#.png");
      pptx.put(media, pngs[k] ?? new Uint8Array());
      pptx.setType(media, "image/png", true);
      const target = deck.slides[si + 1];
      return { img: pptx.link(path, RT_IMAGE, media), sld: pptx.link(path, RT_SLIDE, target.path), section: secs.find((s) => s.slides[0] === target.id)!.id,
        x: Math.round(gx + (k % cols) * (tw + gap)), y: Math.round(gy + Math.floor(k / cols) * (th + gap)), w: Math.round(tw), h: Math.round(th) };
    });
    tree.append(summaryElement(doc, id, `Zusammenfassungszoom ${id - 1}`, { x: Math.round(box.x), y: Math.round(box.y), w: Math.round(box.w), h: Math.round(box.h) }, tiles));
  });
}

function summaryElement(doc: Document, id: number, name: string, f: { x: number; y: number; w: number; h: number }, tiles: { img: string; sld: string; section: string; x: number; y: number; w: number; h: number }[]): Element {
  const ln = `<a:ln w="3175"><a:solidFill><a:prstClr val="ltGray"/></a:solidFill></a:ln>`;
  const P166 = "http://schemas.microsoft.com/office/powerpoint/2016/6/main";
  const objs = tiles.map((t) => `<psuz:summaryZmObj sectionId="${t.section}"><psuz:zmPr id="${guid()}" transitionDur="1000"><p166:blipFill xmlns:p166="${P166}"><a:blip r:embed="${t.img}"/><a:stretch><a:fillRect/></a:stretch></p166:blipFill>` +
    `<p166:spPr xmlns:p166="${P166}"><a:xfrm><a:off x="${t.x}" y="${t.y}"/><a:ext cx="${t.w}" cy="${t.h}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom>${ln}</p166:spPr></psuz:zmPr></psuz:summaryZmObj>`).join("");
  const pics = tiles.map((t, k) => `<p:pic><p:nvPicPr><p:cNvPr id="${id + 1 + k}" name="Grafik ${id + 1 + k}"><a:hlinkClick r:id="${t.sld}" action="ppaction://hlinksldjump"/></p:cNvPr><p:cNvPicPr><a:picLocks noSelect="1" noRot="1" noChangeAspect="1" noMove="1" noResize="1" noEditPoints="1" noAdjustHandles="1" noChangeArrowheads="1" noChangeShapeType="1"/></p:cNvPicPr><p:nvPr/></p:nvPicPr>` +
    `<p:blipFill><a:blip r:embed="${t.img}"/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr><a:xfrm><a:off x="${f.x + t.x}" y="${f.y + t.y}"/><a:ext cx="${t.w}" cy="${t.h}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom>${ln}</p:spPr></p:pic>`).join("");
  const xfrm = `<a:off x="${f.x}" y="${f.y}"/><a:ext cx="${f.w}" cy="${f.h}"/>`;
  const xml = `<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:p="${PptxDom.NS_P}" xmlns:a="${PptxDom.NS_A}" xmlns:r="${PptxDom.NS_R}">` +
    `<mc:Choice xmlns:psuz="http://schemas.microsoft.com/office/powerpoint/2016/summaryzoom" Requires="psuz"><p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="${id}" name="${PptxDom.esc(name)}"/><p:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></p:cNvGraphicFramePr><p:nvPr/></p:nvGraphicFramePr>` +
    `<p:xfrm>${xfrm}</p:xfrm><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/powerpoint/2016/summaryzoom"><psuz:summaryZm>${objs}<psuz:gridLayout/></psuz:summaryZm></a:graphicData></a:graphic></p:graphicFrame></mc:Choice>` +
    `<mc:Fallback><p:grpSp><p:nvGrpSpPr><p:cNvPr id="${id}" name="${PptxDom.esc(name)}"/><p:cNvGrpSpPr><a:grpSpLocks noGrp="1" noUngrp="1" noRot="1" noChangeAspect="1" noMove="1" noResize="1"/></p:cNvGrpSpPr><p:nvPr/></p:nvGrpSpPr>` +
    `<p:grpSpPr><a:xfrm>${xfrm}<a:chOff x="${f.x}" y="${f.y}"/><a:chExt cx="${f.w}" cy="${f.h}"/></a:xfrm></p:grpSpPr>${pics}</p:grpSp></mc:Fallback></mc:AlternateContent>`;
  return PptxDom.fragment(doc, xml)[0];
}

async function abschnittszoomDialog(app: PresentApp): Promise<void> {
  const secs = app.deck.sections();
  const boxes = secs.map((s, i) => {
    const c = h("input", { type: "checkbox" }) as HTMLInputElement;
    return { c, s, el: h("label", { class: "dlg-check" }, c, ` Abschnitt ${i + 1} „${s.name}“`) };
  });
  const ok = await showDialog({ title: "Abschnittszoom einfügen", body: h("div", { class: "form", style: "max-height:360px;overflow:auto;min-width:300px" }, h("b", null, "Wählen Sie die Abschnitte aus, die Sie einfügen möchten."), ...boxes.map((b) => b.el)),
    buttons: [{ label: "Einfügen", value: true, primary: true }, { label: "Abbrechen", value: false }], cancel: false });
  if (!ok) return;
  const picked = boxes.filter((b) => b.c.checked && b.s.slides.length).map((b) => b.s);
  if (picked.length) await insertZooms(app, picked.map((s) => ({ slide: app.deck.slides.findIndex((x) => x.id === s.slides[0]), section: s.id })));
}

/** Folienzoom for each slide index, on the current slide. */
export async function insertSlideZooms(app: PresentApp, targets: number[]): Promise<void> {
  await insertZooms(app, targets.map((slide) => ({ slide })));
}

/** Zooms on the current slide: to a slide, or with `section` to a section (its first slide as the picture). */
export async function insertZooms(app: PresentApp, items: { slide: number; section?: string }[]): Promise<void> {
  const targets = items.map((x) => x.slide).filter((i) => i >= 0);
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
      const section = items[k].section;
      tree.append(zoomElement(doc, { id, name: `${section ? "Abschnittszoom" : "Folienzoom"} ${id - 1}`, sldId: target.id, cId: section ? 0 : creationId(app, target.path), section, img, sld, x, y, w, h: hh }));
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

function zoomElement(doc: Document, z: { id: number; name: string; sldId: number; cId: number; section?: string; img: string; sld: string; x: number; y: number; w: number; h: number }): Element {
  const name = PptxDom.esc(z.name);
  const ln = `<a:ln w="3175"><a:solidFill><a:prstClr val="ltGray"/></a:solidFill></a:ln>`;
  const xml = `<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:p="${PptxDom.NS_P}" xmlns:a="${PptxDom.NS_A}" xmlns:r="${PptxDom.NS_R}">` +
    (z.section ? `<mc:Choice xmlns:psez="http://schemas.microsoft.com/office/powerpoint/2016/sectionzoom" Requires="psez"><p:graphicFrame>` : `<mc:Choice xmlns:pslz="http://schemas.microsoft.com/office/powerpoint/2016/slidezoom" Requires="pslz"><p:graphicFrame>`) +
    `<p:nvGraphicFramePr><p:cNvPr id="${z.id}" name="${name}"/><p:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></p:cNvGraphicFramePr><p:nvPr/></p:nvGraphicFramePr>` +
    `<p:xfrm><a:off x="${z.x}" y="${z.y}"/><a:ext cx="${z.w}" cy="${z.h}"/></p:xfrm>` +
    (z.section ? `<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/powerpoint/2016/sectionzoom"><psez:sectionZm><psez:sectionZmObj sectionId="${z.section}"><psez:zmPr id="${guid()}" transitionDur="1000">`
      : `<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/powerpoint/2016/slidezoom"><pslz:sldZm><pslz:sldZmObj sldId="${z.sldId}" cId="${z.cId}"><pslz:zmPr id="${guid()}" returnToParent="0" transitionDur="1000">`) +
    `<p166:blipFill xmlns:p166="http://schemas.microsoft.com/office/powerpoint/2016/6/main"><a:blip r:embed="${z.img}"/><a:stretch><a:fillRect/></a:stretch></p166:blipFill>` +
    `<p166:spPr xmlns:p166="http://schemas.microsoft.com/office/powerpoint/2016/6/main"><a:xfrm><a:off x="0" y="0"/><a:ext cx="${z.w}" cy="${z.h}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom>${ln}</p166:spPr>` + (z.section ? "</psez:zmPr></psez:sectionZmObj></psez:sectionZm>" : "</pslz:zmPr></pslz:sldZmObj></pslz:sldZm>") + `</a:graphicData></a:graphic></p:graphicFrame></mc:Choice>` +
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
  const sum = frame.getElementsByTagNameNS("http://schemas.microsoft.com/office/powerpoint/2016/summaryzoom", "summaryZm")[0];
  if (sum) {
    // Zusammenfassungszoom: the Fallback group follows the frame (its chOff/chExt stay, so its pictures scale).
    const xf = kid(frame, "xfrm");
    const fb = frame.parentElement?.nextElementSibling;
    const gxf = fb?.localName === "Fallback" ? fb.getElementsByTagNameNS(PptxDom.NS_A, "xfrm")[0] : undefined;
    for (const name of ["off", "ext"]) { const from = kid(xf, name), to = kid(gxf, name); if (from && to) for (const at of Array.from(from.attributes)) to.setAttribute(at.name, at.value); }
    return;
  }
  const zm = frame.getElementsByTagNameNS("http://schemas.microsoft.com/office/powerpoint/2016/slidezoom", "zmPr")[0] ?? frame.getElementsByTagNameNS("http://schemas.microsoft.com/office/powerpoint/2016/sectionzoom", "zmPr")[0];
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
