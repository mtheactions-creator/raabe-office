// SmartArt from PowerPoint's files (D-072): drawn from PowerPoint's own drawing of it, and its texts edited in
// our Textbereich: each point of the data model (dgm:pt) as a line, indented by its level. A change goes into
// the data model and into the drawing's shape for that point (dsp:sp with the same modelId), so PowerPoint
// and RaabePresent show the same. Layout changes, new shapes and SmartArt-Formatvorlagen need PowerPoint's
// layout definitions, which we do not have (creating SmartArt is not built, see the decision).

import { SMARTART_COLORS, SMARTART_STYLES } from "./smartart-styles.gen";
import { h, showDialog, type Control, type Tab } from "@raabe/shell";
import { PptxDom } from "@raabe/ooxml";
import { colorXml } from "./colors";
import * as T from "./textops";
import type { PresentApp } from "./app";
import * as sc from "./smartart-create";

const { kid, kids } = PptxDom;
const NS_DGM = "http://schemas.openxmlformats.org/drawingml/2006/diagram";
const NS_DSP = "http://schemas.microsoft.com/office/drawing/2008/diagram";

interface Parts { data: string; drawing?: string }

function partsOf(app: PresentApp, id: number): Parts | null {
  const path = app.slidePath();
  const el = path ? app.editor.shapeEl(id) : undefined;
  const rel = el?.getElementsByTagNameNS(NS_DGM, "relIds")[0];
  const dm = rel?.getAttributeNS(PptxDom.NS_R, "dm");
  const data = path && dm ? app.deck.pptx.target(path, dm) : undefined;
  if (!path || !data || !app.deck.pptx.has(data)) return null;
  const ext = app.deck.pptx.doc(data).getElementsByTagNameNS("http://schemas.microsoft.com/office/drawing/2008/diagram", "dataModelExt")[0];
  const drawing = ext ? app.deck.pptx.target(path, ext.getAttribute("relId") ?? "") : undefined;
  return { data, drawing: drawing && app.deck.pptx.has(drawing) ? drawing : undefined };
}

/**
 * Format > Fülleffekt / Formkontur with the whole SmartArt selected: PowerPoint writes the graphic's background
 * (dgm:bg, recorded: probe_powerpoint_smartart_format.py) and outline (dgm:whole/a:ln). null: Keine Füllung / Kontur;
 * "AUTO": the graphic's own fill or line taken away.
 */
export function setDiagramWhole(app: PresentApp, id: number, what: "fill" | "line", hex: string | null, width?: number): void {
  const parts = partsOf(app, id);
  const slide = app.slidePath();
  if (!parts || !slide) return;
  const doc = app.deck.pptx.doc(parts.data);
  const root = doc.documentElement;
  const el = (name: string) => {
    let e = PptxDom.kids(root).find((k) => k.localName === name);
    if (!e) { e = doc.createElementNS(NS_DGM, `dgm:${name}`); const before = PptxDom.kids(root).find((k) => k.localName === (name === "bg" ? "whole" : "extLst")); root.insertBefore(e, before ?? null); }
    return e;
  };
  const fill = (v: string | null) => (v === null ? "<a:noFill/>" : `<a:solidFill>${colorXml(v)}</a:solidFill>`);
  app.deck.change(what === "fill" ? "Fülleffekt" : "Formkontur", [parts.data, slide], () => {
    if (what === "fill") {
      const bg = el("bg");
      for (const c of PptxDom.kids(bg)) if (["noFill", "solidFill", "gradFill", "blipFill", "pattFill", "grpFill"].includes(c.localName)) c.remove();
      if (hex !== "AUTO") bg.prepend(PptxDom.fragment(doc, fill(hex))[0]);
    } else {
      const whole = el("whole");
      let ln = PptxDom.kid(whole, "ln");
      if (!ln) { ln = PptxDom.make(doc, "a:ln"); whole.prepend(ln); }
      if (width !== undefined) ln.setAttribute("w", String(Math.round(width * 12700)));
      if (hex !== undefined) {
        for (const c of PptxDom.kids(ln)) if (["noFill", "solidFill", "gradFill", "pattFill"].includes(c.localName)) c.remove();
        if (hex !== "AUTO") ln.prepend(PptxDom.fragment(doc, fill(hex))[0]);
      }
    }
  });
  // The slide's view of the frame reads these parts: drawn anew.
  app.deck.invalidate(slide);
  app.editor.render();
}

/**
 * WordArt-Formate (Textfüllung, Textkontur, Texteffekte ...) with the whole SmartArt selected: every run of its text,
 * in the data model's points and in the drawing PowerPoint shows, one undo step with the diagram's parts. (A new
 * layout writes the drawing anew from the items: the formatting does not travel with it yet.)
 */
export function formatDiagramText(app: PresentApp, ids: number[], label: string, fn: (rPr: Element) => void): void {
  const slide = app.slidePath();
  const all = ids.map((id) => partsOf(app, id)).filter((p): p is Parts => !!p);
  if (!slide || !all.length) return;
  const paths = [slide, ...all.flatMap((p) => [p.data, ...(p.drawing ? [p.drawing] : [])])];
  app.deck.change(label, paths, () => {
    for (const p of all) {
      for (const t of Array.from(app.deck.pptx.doc(p.data).getElementsByTagNameNS(NS_DGM, "t"))) if (t.getElementsByTagNameNS(PptxDom.NS_A, "t").length) T.formatAll(t, fn);
      if (p.drawing) for (const b of Array.from(app.deck.pptx.doc(p.drawing).getElementsByTagNameNS(NS_DSP, "txBody"))) if (b.getElementsByTagNameNS(PptxDom.NS_A, "t").length) T.formatAll(b, fn);
    }
  });
  app.deck.invalidate(slide);
  app.editor.render();
}

/** The text points of the data model with their level (from the parOf connections), in the model's order. */
function textPoints(doc: Document): { pt: Element; level: number; text: string }[] {
  const pts = Array.from(doc.getElementsByTagNameNS(NS_DGM, "pt"));
  const parent = new Map<string, string>();
  for (const c of Array.from(doc.getElementsByTagNameNS(NS_DGM, "cxn"))) {
    const type = c.getAttribute("type") ?? "parOf";
    if (type === "parOf") parent.set(c.getAttribute("destId") ?? "", c.getAttribute("srcId") ?? "");
  }
  const docId = pts.find((p) => p.getAttribute("type") === "doc")?.getAttribute("modelId") ?? "";
  const level = (id: string) => { let n = 0, cur = parent.get(id); while (cur && cur !== docId && n < 10) { n++; cur = parent.get(cur); } return n; };
  return pts.filter((p) => !p.getAttribute("type") || p.getAttribute("type") === "node").map((pt) => {
    const t = kid(pt, "t");
    const text = t ? Array.from(t.getElementsByTagNameNS(PptxDom.NS_A, "p")).map((p) => Array.from(p.getElementsByTagNameNS(PptxDom.NS_A, "t")).map((x) => x.textContent ?? "").join("")).join(" ") : "";
    return { pt, level: level(pt.getAttribute("modelId") ?? ""), text };
  });
}

/** Puts plain text into a text body (dgm:t or dsp:txBody): one paragraph, the first run's look kept. */
function setText(body: Element, text: string): void {
  const doc = body.ownerDocument;
  const ps = kids(body, "p");
  const p = ps[0] ?? body.appendChild(PptxDom.make(doc, "a:p"));
  for (const extra of ps.slice(1)) extra.remove();
  const runs = kids(p, "r");
  const r = runs[0];
  for (const x of runs.slice(1)) x.remove();
  for (const x of kids(p).filter((c) => c.localName === "br" || c.localName === "fld")) x.remove();
  if (!text) { r?.remove(); return; }
  if (r) { const t = kid(r, "t"); if (t) t.textContent = text; else r.append(PptxDom.make(doc, "a:t", {}, text)); return; }
  const nr = PptxDom.make(doc, "a:r", {}, PptxDom.make(doc, "a:rPr", { lang: "de-DE" }), PptxDom.make(doc, "a:t", {}, text));
  p.insertBefore(nr, kid(p, "endParaRPr") ?? null);
}

/**
 * SmartArt-Design > Textbereich (D-077): the items as lines, as PowerPoint's text pane: Eingabe makes a new item,
 * Tab moves an item a level down and Umschalt+Tab up (shown by its indent). When only texts changed, they go
 * into the data model and into the drawing's shapes (PowerPoint's own drawing stays); when items were added,
 * removed or moved between levels, the diagram is laid out again.
 */
export async function textPane(app: PresentApp, id: number): Promise<void> {
  const parts = partsOf(app, id);
  const d = sc.diagramOf(app, id);
  if (!parts || !d) return;
  const pptx = app.deck.pptx;
  const points = textPoints(pptx.doc(parts.data));
  const TAB = String.fromCharCode(9);
  const area = h("textarea", { class: "sa-textpane", rows: 12, spellcheck: "true", "aria-label": "Text hier eingeben" }) as HTMLTextAreaElement;
  area.value = d.items.map((i) => TAB.repeat(i.level) + i.text).join(String.fromCharCode(10));
  area.addEventListener("keydown", (e) => {
    if (e.key !== "Tab") return;
    e.preventDefault();
    const v = area.value, at = area.selectionStart;
    const start = v.lastIndexOf(String.fromCharCode(10), at - 1) + 1;
    if (e.shiftKey) { if (v[start] === TAB) { area.value = v.slice(0, start) + v.slice(start + 1); area.selectionStart = area.selectionEnd = Math.max(start, at - 1); } }
    else { area.value = v.slice(0, start) + TAB + v.slice(start); area.selectionStart = area.selectionEnd = at + 1; }
  });
  const ok = await showDialog({ title: "Text hier eingeben", body: h("div", { class: "form" }, area), buttons: [{ label: "OK", value: true, primary: true }, { label: "Abbrechen", value: false }], cancel: false });
  if (!ok) return;
  const lines = area.value.split(String.fromCharCode(10));
  const items = lines.map((l) => { const lvl = /^\t*/.exec(l)![0].length; return { text: l.slice(lvl).trim(), level: lvl }; }).filter((x, _i, all) => x.text || all.length === 1);
  // Levels never jump by more than one below the item before.
  items.forEach((it, i) => { it.level = i === 0 ? 0 : Math.min(it.level, items[i - 1].level + 1); });
  const sameShape = items.length === d.items.length && items.every((it, i) => it.level === d.items[i].level);
  if (!sameShape) { sc.relayout(app, id, items); return; }
  const changed = points.map((p, i) => ({ ...p, next: items[i]?.text ?? p.text })).filter((x) => x.next !== x.text);
  if (!changed.length) return;
  const path = app.slidePath()!;
  app.deck.change("SmartArt-Text", [path, parts.data, ...(parts.drawing ? [parts.drawing] : [])], () => {
    // The parts are edited through their documents; the data model's points and the drawing's shapes follow each other.
    const drawing = parts.drawing ? pptx.doc(parts.drawing) : null;
    for (const c of changed) {
      let t = kid(c.pt, "t");
      if (!t) { t = PptxDom.fragment(c.pt.ownerDocument, `<dgm:t xmlns:dgm="${NS_DGM}" xmlns:a="${PptxDom.NS_A}"><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="de-DE"/></a:p></dgm:t>`)[0]; c.pt.append(t); }
      setText(t, c.next);
      // The drawing's shape for a node is one of its presentation points (dgm:pt type "pres" whose prSet names
      // the node in presAssocID), or (our diagrams) the node itself; the one that carries text gets it.
      const mid = c.pt.getAttribute("modelId") ?? "";
      const pres = new Set([mid, ...Array.from(c.pt.ownerDocument.getElementsByTagNameNS(NS_DGM, "pt"))
        .filter((p) => kid(p, "prSet")?.getAttribute("presAssocID") === mid).map((p) => p.getAttribute("modelId") ?? "")]);
      const sp = drawing ? Array.from(drawing.getElementsByTagNameNS(NS_DSP, "sp")).find((x) => pres.has(x.getAttribute("modelId") ?? "") && kid(x, "txBody")) : undefined;
      const body = sp ? kid(sp, "txBody") : undefined;
      if (body) setText(body, c.next);
    }
  });
  app.deck.invalidate(path);
  app.editor.render();
}

/** SmartArt-Design: the contextual tab as PowerPoint's (D-077). */
export function smartArtTab(app: PresentApp): Tab {
  const b = (cmd: string, key: string, size: "large" | "small" = "large"): Control => ({ type: "button", cmd, key, size });
  void app;
  return {
    id: "smartartDesign", label: "SmartArt-Design", key: "JS", contextual: true, shrink: ["styles", "layouts"],
    groups: [
      { id: "create", label: "Grafik erstellen", key: "ZE", icon: "slide-text", controls: [
        { type: "split", cmd: "smartart.addShape", key: "O", size: "large", menu: [{ cmd: "smartart.addAfter" }, { cmd: "smartart.addBefore" }, { cmd: "smartart.addAbove" }, { cmd: "smartart.addBelow" }] },
        { type: "stack", items: [b("smartart.addBullet", "A1", "small"), b("smartart.textPane", "T", "small"), b("smartart.rtl", "V", "small")] },
        { type: "stack", items: [b("smartart.promote", "H", "small"), b("smartart.demote", "I", "small")] },
        { type: "stack", items: [b("smartart.moveUp", "M", "small"), b("smartart.moveDown", "A2", "small")] }] },
      { id: "layouts", label: "Layouts", key: "ZL", icon: "shapes", controls: [{ type: "dropdown", cmd: "smartart.layouts", key: "L", size: "large", get menu() { return smartArtLayoutMenu(app); } }] },
      { id: "styles", label: "SmartArt-Formatvorlagen", key: "ZF", icon: "shapes", controls: [{ type: "dropdown", cmd: "smartart.colors", key: "F", size: "large", get menu() { return smartArtColorMenu(app); } },
        { type: "dropdown", cmd: "smartart.styles", key: "S", size: "large", get menu() { return smartArtStyleMenu(app); } }] },
      { id: "reset", label: "Zurücksetzen", key: "ZZ", icon: "slide-text", controls: [b("smartart.reset", "G"), { type: "dropdown", cmd: "smartart.convert", key: "K", size: "large", menu: [{ cmd: "smartart.toText" }, { cmd: "smartart.toShapes" }] }] }],
  };
}

/** The selected diagram's id. */
export const selectedDiagram = (app: PresentApp): number | null => app.editor.selectedViews.find((v) => v.frame === "diagram")?.id ?? null;

/** Layouts: the selected diagram's category's layouts, and Weitere Layouts... (the dialog). */
function smartArtLayoutMenu(app: PresentApp) {
  const id = selectedDiagram(app);
  const d = id !== null ? sc.diagramOf(app, id) : null;
  const cur = d ? sc.entryById(d.layoutId) : undefined;
  return [...sc.layoutsFor(cur?.category).map((e) => ({ label: e.name, checked: e.id === cur?.id, run: () => { if (id !== null && d) sc.relayout(app, id, d.items, e.id); } })), "-" as const,
    { label: "Weitere Layouts...", run: () => void sc.chooseLayout(cur?.id).then((e) => { if (e && id !== null && d) sc.relayout(app, id, d.items, e.id); }) }];
}

/** Farben ändern: PowerPoint's 38 colour sets by name under its headings (recorded, D-101). */
function smartArtColorMenu(app: PresentApp) {
  const id = selectedDiagram(app);
  // the current set is shown by the diagram itself; the menu shows each set's colours
  const out: ({ label: string; run: () => void; preview?: HTMLElement } | { heading: string })[] = [];
  let last = "";
  for (const c of SMARTART_COLORS) {
    const head = c.cat === "mainScheme" ? "Primärdesignfarben" : c.cat === "colorful" ? "Bunt" : `Akzent ${c.cat.replace("accent", "")}`;
    if (head !== last) { out.push({ heading: head }); last = head; }
    const key = c.id.split("/").pop()!;
    out.push({ label: c.name, preview: colorSetSwatches(c), run: () => { if (id !== null) sc.setColors(app, id, key); } });
  }
  return out;
}

/** A SmartArt colour set's picture (D-164): its first four nodes' fill (and outline) in the Office colours. */
function colorSetSwatches(c: (typeof SMARTART_COLORS)[number]): HTMLElement {
  const base: Record<string, string> = { lt1: "#FFFFFF", dk1: "#000000", tx1: "#000000", bg1: "#FFFFFF", dk2: "#44546A", lt2: "#E7E6E6",
    accent1: "#4472C4", accent2: "#ED7D31", accent3: "#A5A5A5", accent4: "#FFC000", accent5: "#5B9BD5", accent6: "#70AD47" };
  const node = (c.labels as Record<string, { fill?: { clr?: string[] }; line?: { clr?: string[] } }>).node1 ?? {};
  const pick = (list: string[] | undefined, i: number) => {
    const raw = list?.length ? list[i % list.length] : "accent1";
    const [name, ...mods] = raw.split(" ");
    let hex = base[name] ?? "#4472C4";
    for (const m of mods) {
      const [k, v] = m.split("=");
      const f = Number(v) / 100000;
      if (k === "alpha" && f === 0) return "transparent";
      const ch = (x: number) => (k === "tint" ? Math.round(x + (255 - x) * (1 - f)) : k === "shade" ? Math.round(x * f) : x);
      if (k === "tint" || k === "shade") hex = "#" + [1, 3, 5].map((j) => ch(parseInt(hex.slice(j, j + 2), 16)).toString(16).padStart(2, "0")).join("");
    }
    return hex;
  };
  return h("span", { style: "display:inline-flex;gap:2px" }, ...[0, 1, 2, 3].map((i) =>
    h("span", { style: `display:inline-block;width:13px;height:13px;background:${pick(node.fill?.clr, i)};outline:1px solid ${pick(node.line?.clr, i)}` })));
}

/** SmartArt-Formatvorlagen: PowerPoint's 14 quick styles by name (recorded, D-101). */
function smartArtStyleMenu(app: PresentApp) {
  const id = selectedDiagram(app);
  const cur = id !== null ? sc.diagramOf(app, id)?.style : undefined;
  return SMARTART_STYLES.map((s) => {
    const key = s.id.split("/").pop()!;
    return { label: s.name, checked: key === cur, run: () => { if (id !== null) sc.setQuickStyle(app, id, key); } };
  });
}

/**
 * SmartArt-Design > Konvertieren: In Formen konvertieren puts PowerPoint's drawing of the SmartArt on the slide
 * as a group of ordinary shapes; In Text konvertieren makes a text box with its points as a bulleted list (by
 * level). Both replace the SmartArt, as in PowerPoint.
 */
export function convertSmartArt(app: PresentApp, id: number, to: "shapes" | "text"): void {
  const parts = partsOf(app, id);
  const el = app.editor.shapeEl(id);
  if (!parts || !el) return;
  const pptx = app.deck.pptx;
  const xf = kid(el, "xfrm");
  const off = kid(xf, "off"), ext = kid(xf, "ext");
  const x = off?.getAttribute("x") ?? "0", y = off?.getAttribute("y") ?? "0", cx = ext?.getAttribute("cx") ?? "0", cy = ext?.getAttribute("cy") ?? "0";
  const doc = el.ownerDocument;
  const P = PptxDom.NS_P, A = PptxDom.NS_A;
  app.editor.edit(to === "shapes" ? "In Formen konvertieren" : "In Text konvertieren", () => {
    let made: Element;
    if (to === "shapes" && parts.drawing) {
      const drawing = pptx.doc(parts.drawing);
      const ser = new XMLSerializer();
      let n = id;
      const shapes = Array.from(drawing.getElementsByTagNameNS(NS_DSP, "sp")).map((sp) => {
        n++;
        const part = (name: string) => { const c = kid(sp, name); return c ? Array.from(c.childNodes).map((x) => ser.serializeToString(x)).join("") : ""; };
        const bodyPr = kid(sp, "txBody") ? `<p:txBody>${part("txBody")}</p:txBody>` : "";
        const style = kid(sp, "style") ? `<p:style>${part("style")}</p:style>` : "";
        return `<p:sp><p:nvSpPr><p:cNvPr id="${n}" name="Form ${n - 1}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr>${part("spPr")}</p:spPr>${style}${bodyPr}</p:sp>`;
      });
      made = PptxDom.fragment(doc, `<p:grpSp xmlns:p="${P}" xmlns:a="${A}"><p:nvGrpSpPr><p:cNvPr id="${id}" name="Gruppieren ${id - 1}"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>` +
        `<p:grpSpPr><a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${cx}" cy="${cy}"/><a:chOff x="0" y="0"/><a:chExt cx="${cx}" cy="${cy}"/></a:xfrm></p:grpSpPr>${shapes.join("")}</p:grpSp>`)[0];
      // New ids must not clash with the slide's other shapes.
      const tree = el.parentElement!;
      let next = Math.max(...Array.from(tree.getElementsByTagNameNS(P, "cNvPr")).map((c) => Number(c.getAttribute("id")) || 0)) + 1;
      for (const c of Array.from(made.getElementsByTagNameNS(P, "cNvPr")).slice(1)) c.setAttribute("id", String(next++));
    } else {
      const pts = textPoints(pptx.doc(parts.data));
      const esc = PptxDom.esc;
      const paras = pts.map((p) => `<a:p><a:pPr marL="${285750 + p.level * 457200}" lvl="${Math.min(8, p.level)}" indent="-285750"><a:buFont typeface="Arial"/><a:buChar char="•"/></a:pPr><a:r><a:rPr lang="de-DE"/><a:t>${esc(p.text)}</a:t></a:r></a:p>`).join("");
      made = PptxDom.fragment(doc, `<p:sp xmlns:p="${P}" xmlns:a="${A}"><p:nvSpPr><p:cNvPr id="${id}" name="Textfeld ${id - 1}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>` +
        `<p:spPr><a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr>` +
        `<p:txBody><a:bodyPr wrap="square" rtlCol="0"><a:spAutoFit/></a:bodyPr><a:lstStyle/>${paras || '<a:p><a:endParaRPr lang="de-DE"/></a:p>'}</p:txBody></p:sp>`)[0];
    }
    el.replaceWith(made);
  });
  app.editor.select([id]);
}
