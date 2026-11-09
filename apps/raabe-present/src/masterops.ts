// Folienmaster: Folienmaster einfügen, Masterlayout... and Platzhalter einfügen, as PowerPoint 365 does them (D-153;
// recorded in results/2026-11-09-powerpoint-followup.json by probe_powerpoint_followup.py).

import { h, showDialog } from "@raabe/shell";
import { CT_PPTX, OFFICE_DESIGN, PptxDom, masterXml, newPptxFiles, phOf } from "@raabe/ooxml";
import type { Deck } from "./deck";
import * as S from "./shapeops";

const { kids, at, make, num, fragment, parseXml, esc } = PptxDom;
const RT = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const td = new TextDecoder();

/** The ids of p:sldMasterId and p:sldLayoutId share one range (from 2147483648); the highest used. */
function maxMasterLayoutId(deck: Deck): number {
  const pptx = deck.pptx;
  let max = 2147483647;
  for (const m of kids(at(pptx.presentation, "sldMasterIdLst"), "sldMasterId")) max = Math.max(max, num(m, "id") ?? 0);
  for (const m of pptx.masters()) for (const l of kids(at(pptx.doc(m).documentElement, "sldLayoutIdLst"), "sldLayoutId")) max = Math.max(max, num(l, "id") ?? 0);
  return max;
}

/** Positions written for 16:9 (the template) scaled to the presentation's slide size. */
function scaleBoxes(doc: Document, sx: number, sy: number): void {
  if (sx === 1 && sy === 1) return;
  for (const off of Array.from(doc.getElementsByTagNameNS(PptxDom.NS_A, "off"))) {
    off.setAttribute("x", String(Math.round((num(off, "x") ?? 0) * sx)));
    off.setAttribute("y", String(Math.round((num(off, "y") ?? 0) * sy)));
  }
  for (const ext of Array.from(doc.getElementsByTagNameNS(PptxDom.NS_A, "ext"))) {
    if (!ext.hasAttribute("cx")) continue;
    ext.setAttribute("cx", String(Math.round((num(ext, "cx") ?? 0) * sx)));
    ext.setAttribute("cy", String(Math.round((num(ext, "cy") ?? 0) * sy)));
  }
}

/**
 * Folienmaster einfügen: a new master after the others with PowerPoint's eleven layouts and the Office design, named
 * "Benutzerdefiniertes Design" and kept (preserve), as PowerPoint inserts it (recorded). Returns the new master's path.
 */
export function insertMaster(deck: Deck): string {
  const pptx = deck.pptx;
  const src = newPptxFiles({ ...OFFICE_DESIGN, name: "Benutzerdefiniertes Design" });
  const size = pptx.slideSize, sx = size.cx / 12192000, sy = size.cy / 6858000;
  const master = pptx.freeName("ppt/slideMasters/slideMaster#.xml");
  const theme = pptx.freeName("ppt/theme/theme#.xml");
  const srcLayouts = Object.keys(src).filter((k) => /^ppt\/slideLayouts\/slideLayout\d+\.xml$/.test(k)).sort((a, b) => Number(a.match(/(\d+)\.xml$/)![1]) - Number(b.match(/(\d+)\.xml$/)![1]));
  const layouts: string[] = [];
  for (let i = 0; i < srcLayouts.length; i++) {
    let p = "";
    for (let k = 1; ; k++) { p = `ppt/slideLayouts/slideLayout${k}.xml`; if (!pptx.has(p) && !layouts.includes(p)) break; }
    layouts.push(p);
  }
  const paths = [pptx.presPath, master, theme, ...layouts];
  deck.change("Folienmaster einfügen", paths, () => {
    pptx.put(theme, parseXml(td.decode(src["ppt/theme/theme1.xml"])), CT_PPTX.theme);
    const mdoc = parseXml(td.decode(src["ppt/slideMasters/slideMaster1.xml"]));
    mdoc.documentElement.setAttribute("preserve", "1");
    scaleBoxes(mdoc, sx, sy);
    const list = at(mdoc.documentElement, "sldLayoutIdLst")!;
    for (const x of kids(list, "sldLayoutId")) x.remove();
    let id = maxMasterLayoutId(deck);
    const masterId = ++id;
    pptx.put(master, mdoc, CT_PPTX.slideMaster);
    srcLayouts.forEach((s, i) => {
      const ldoc = parseXml(td.decode(src[s]));
      scaleBoxes(ldoc, sx, sy);
      pptx.put(layouts[i], ldoc, CT_PPTX.slideLayout);
      pptx.link(layouts[i], `${RT}/slideMaster`, master);
      const rid = pptx.link(master, `${RT}/slideLayout`, layouts[i]);
      const el = make(mdoc, "p:sldLayoutId", { id: ++id });
      el.setAttributeNS(PptxDom.NS_R, "r:id", rid);
      list.append(el);
    });
    pptx.link(master, `${RT}/theme`, theme);
    const pres = pptx.presentation;
    const mlist = at(pres, "sldMasterIdLst")!;
    const rid = pptx.link(pptx.presPath, `${RT}/slideMaster`, master);
    const el = make(pres.ownerDocument, "p:sldMasterId", { id: masterId });
    el.setAttributeNS(PptxDom.NS_R, "r:id", rid);
    mlist.append(el);
    pptx.touch(pptx.presPath);
  });
  return master;
}

// ----- Masterlayout... -----------------------------------------------------------------------------------------------

/** The dialog's check boxes in PowerPoint's order: label, placeholder type, access key. */
export const MASTER_LAYOUT_BOXES: [string, string, string][] = [["Titel", "title", "T"], ["Text", "body", "x"], ["Datum", "dt", "D"], ["Foliennummer", "sldNum", "e"], ["Fußzeile", "ftr", "u"]];

const MASTER_PH_NAMES: Record<string, string> = { title: "Titelplatzhalter", body: "Textplatzhalter", dt: "Datumsplatzhalter", ftr: "Fußzeilenplatzhalter", sldNum: "Foliennummernplatzhalter" };

export function masterHas(deck: Deck, master: string, type: string): boolean {
  return S.shapeEls(S.treeOf(deck.pptx.doc(master))).some((sp) => (phOf(sp)?.type ?? "body") === type && !!phOf(sp));
}

/** Puts a placeholder of the master back at Office's place (scaled to the slide size), after the master's shapes. */
function restoreMasterPh(deck: Deck, master: string, type: string): void {
  const size = deck.pptx.slideSize;
  const tpl = parseXml(masterXml(OFFICE_DESIGN));
  const sp = S.shapeEls(S.treeOf(tpl)).find((x) => (phOf(x)?.type ?? "body") === type);
  if (!sp) return;
  scaleBoxes(tpl, size.cx / 12192000, size.cy / 6858000);
  const doc = deck.pptx.doc(master);
  const tree = S.treeOf(doc);
  const copy = doc.importNode(sp, true) as Element;
  const id = S.maxId(tree) + 1;
  const c = S.cNvPr(copy)!;
  c.setAttribute("id", String(id));
  c.setAttribute("name", `${MASTER_PH_NAMES[type]} ${id - 1}`);
  tree.append(copy);
}

/** Masterlayout...: which of the master's placeholders it has; unticked ones go, ticked missing ones come back. */
export async function masterLayoutDialog(deck: Deck, master: string): Promise<boolean> {
  const boxes = MASTER_LAYOUT_BOXES.map(([label, type, key]) => {
    const input = h("input", { type: "checkbox", accesskey: key.toLowerCase() }) as HTMLInputElement;
    input.checked = masterHas(deck, master, type);
    const k = label.indexOf(key);
    const text = k >= 0 ? [label.slice(0, k), h("u", null, label[k]), label.slice(k + 1)] : [label];
    return { type, input, row: h("label", { class: "ml-box" }, input, " ", ...text) };
  });
  const body = h("div", { class: "form master-layout" }, h("fieldset", null, h("legend", null, "Platzhalter"), ...boxes.map((b) => b.row)));
  const ok = await showDialog({ title: "Masterlayout", body, buttons: [{ label: "OK", value: true, primary: true }, { label: "Abbrechen", value: false }], cancel: false });
  if (!ok) return false;
  const want = boxes.filter((b) => b.input.checked !== masterHas(deck, master, b.type));
  if (!want.length) return false;
  deck.change("Masterlayout", [master], () => {
    for (const b of want) {
      if (b.input.checked) restoreMasterPh(deck, master, b.type);
      else for (const sp of S.shapeEls(S.treeOf(deck.pptx.doc(master)))) if (phOf(sp) && (phOf(sp)?.type ?? "body") === b.type) sp.remove();
    }
  });
  return true;
}

// ----- Platzhalter einfügen --------------------------------------------------------------------------------------------

/** The menu's entries in PowerPoint's order (recorded), their p:ph type ("" for Inhalt) and the shape's name. */
export const PLACEHOLDER_KINDS: { label: string; type: string; name: string; ext?: string }[] = [
  { label: "Inhalt", type: "", name: "Inhaltsplatzhalter" },
  { label: "Text", type: "body", name: "Textplatzhalter" },
  { label: "Bild", type: "pic", name: "Bildplatzhalter" },
  { label: "Onlinegrafik", type: "clipArt", name: "Onlinebild-Platzhalter" },
  { label: "Diagramm", type: "chart", name: "Diagrammplatzhalter" },
  { label: "Tabelle", type: "tbl", name: "Tabellenplatzhalter" },
  { label: "SmartArt", type: "dgm", name: "SmartArt-Platzhalter" },
  { label: "Medien", type: "media", name: "Medienplatzhalter" },
  // Cameo: a media placeholder marked as the camera's (PowerPoint's p232 extension, recorded).
  { label: "Cameo", type: "media", name: "Cameo-Platzhalter", ext: '<p:extLst><p:ext uri="{56F484CC-4922-43CF-B6FB-B326C6A72FC8}"><p232:phTypeExt xmlns:p232="http://schemas.microsoft.com/office/powerpoint/2023/02/main"><p232:type><p232:cameo/></p232:type></p232:phTypeExt></p:ext></p:extLst>' },
];

const LEVELS = ["Mastertextformat bearbeiten", "Zweite Ebene", "Dritte Ebene", "Vierte Ebene", "Fünfte Ebene"];

/**
 * A placeholder drawn on a layout: Inhalt and Text with the master's five levels as their prompt (PowerPoint), the
 * others empty; idx the next free one of the layout (at least 13, after the footers' 10 to 12).
 */
export function newPlaceholder(doc: Document, id: number, kind: { type: string; name: string; ext?: string }, box: { x: number; y: number; w: number; h: number }): Element {
  const tree = S.treeOf(doc);
  let idx = 12;
  for (const sp of S.shapeEls(tree)) { const p = at(sp, "nvSpPr/nvPr/ph") ?? at(sp, "nvPicPr/nvPr/ph") ?? at(sp, "nvGraphicFramePr/nvPr/ph"); if (p) idx = Math.max(idx, num(p, "idx") ?? 0); }
  const text = kind.type === "" || kind.type === "body";
  const paras = text ? LEVELS.map((t, i) => `<a:p><a:pPr lvl="${i}"/><a:r><a:rPr lang="de-DE"/><a:t>${esc(t)}</a:t></a:r></a:p>`).join("") : '<a:p><a:endParaRPr lang="de-DE"/></a:p>';
  const attrs = `${kind.type ? ` type="${kind.type}"` : ""} sz="quarter" idx="${idx + 1}"`;
  const ph = kind.ext ? `<p:ph${attrs}>${kind.ext}</p:ph>` : `<p:ph${attrs}/>`;
  return fragment(doc, `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${esc(kind.name)} ${id - 1}"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr>${ph}</p:nvPr></p:nvSpPr>` +
    `<p:spPr><a:xfrm><a:off x="${Math.round(box.x)}" y="${Math.round(box.y)}"/><a:ext cx="${Math.round(box.w)}" cy="${Math.round(box.h)}"/></a:xfrm></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/>${paras}</p:txBody></p:sp>`)[0];
}

export const isMasterPath = (p: string | undefined): boolean => !!p && /slideMasters\/slideMaster\d+\.xml$/.test(p);
