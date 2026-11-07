// SmartArt from PowerPoint's files (D-072): drawn from PowerPoint's own drawing of it, and its texts edited in
// our Textbereich: each point of the data model (dgm:pt) as a line, indented by its level. A change goes into
// the data model and into the drawing's shape for that point (dsp:sp with the same modelId), so PowerPoint
// and RaabePresent show the same. Layout changes, new shapes and SmartArt-Formatvorlagen need PowerPoint's
// layout definitions, which we do not have (creating SmartArt is not built, see the decision).

import { h, showDialog, type Control, type Tab } from "@raabe/shell";
import { PptxDom } from "@raabe/ooxml";
import type { PresentApp } from "./app";

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

/** SmartArt-Design > Textbereich: the SmartArt's texts, one line per point. */
export async function textPane(app: PresentApp, id: number): Promise<void> {
  const parts = partsOf(app, id);
  if (!parts) return;
  const pptx = app.deck.pptx;
  const points = textPoints(pptx.doc(parts.data));
  if (!points.length) return;
  const inputs = points.map((p) => h("input", { type: "text", value: p.text, style: `width:${320 - p.level * 20}px;margin-left:${p.level * 20}px` }) as HTMLInputElement);
  const ok = await showDialog({ title: "Text hier eingeben", body: h("div", { class: "form", style: "max-height:420px;overflow:auto" }, ...inputs.map((i) => h("div", null, "• ", i))),
    buttons: [{ label: "OK", value: true, primary: true }, { label: "Abbrechen", value: false }], cancel: false });
  if (!ok) return;
  const changed = points.map((p, i) => ({ ...p, next: inputs[i].value })).filter((x) => x.next !== x.text);
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
      // the node in presAssocID); the one that carries text gets it.
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

/** SmartArt-Design: the contextual tab, with what works on PowerPoint's SmartArt here (Textbereich). */
export function smartArtTab(app: PresentApp): Tab {
  const b = (cmd: string, key: string): Control => ({ type: "button", cmd, key, size: "large" });
  void app;
  return {
    id: "smartartDesign", label: "SmartArt-Design", key: "JS", contextual: true, shrink: [],
    groups: [{ id: "create", label: "Grafik erstellen", key: "ZE", icon: "slide-text", controls: [b("smartart.textPane", "T")] }],
  };
}
