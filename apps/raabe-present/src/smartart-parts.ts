// Creating SmartArt (D-077): the parts PowerPoint writes for a diagram, made by us. The data model (dgm:dataModel)
// holds the items: a document point, a point per item, and per item a parOf connection with its parent and its
// transition points (parTrans, sibTrans), as PowerPoint writes them; the layout, quick style and colours parts
// carry PowerPoint's ids (loTypeId, qsTypeId, csTypeId) so that PowerPoint lays the diagram out with its own
// definitions when it opens the file, while their content is ours (no definition of Microsoft's is copied); the
// drawing part (dsp:drawing) is our layout engine's result, which RaabePresent draws and PowerPoint shows until
// it lays the diagram out again.

import { PptxDom } from "@raabe/ooxml";

export const NS_DGM = "http://schemas.openxmlformats.org/drawingml/2006/diagram";
export const NS_DSP = "http://schemas.microsoft.com/office/drawing/2008/diagram";
const NS_A = PptxDom.NS_A;

export const RT = {
  data: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramData",
  layout: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramLayout",
  style: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramQuickStyle",
  colors: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramColors",
  drawing: "http://schemas.microsoft.com/office/2007/relationships/diagramDrawing",
};
export const CT = {
  data: "application/vnd.openxmlformats-officedocument.drawingml.diagramData+xml",
  layout: "application/vnd.openxmlformats-officedocument.drawingml.diagramLayout+xml",
  style: "application/vnd.openxmlformats-officedocument.drawingml.diagramStyle+xml",
  colors: "application/vnd.openxmlformats-officedocument.drawingml.diagramColors+xml",
  drawing: "application/vnd.ms-office.drawingml.diagramDrawing+xml",
};
export const QS_SIMPLE = "urn:microsoft.com/office/officeart/2005/8/quickstyle/simple1";
export const CS_ACCENT1 = "urn:microsoft.com/office/officeart/2005/8/colors/accent1_2";

export interface DItem { text: string; level: number }

export const guid = () => {
  const hex = Array.from(crypto.getRandomValues(new Uint8Array(16)), (b) => b.toString(16).padStart(2, "0")).join("").toUpperCase();
  return `{${hex.slice(0, 8)}-${hex.slice(8, 12)}-${hex.slice(12, 16)}-${hex.slice(16, 20)}-${hex.slice(20)}}`;
};

const esc = (s: string) => PptxDom.esc(s);
const emptyT = `<dgm:t><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="de-DE"/></a:p></dgm:t>`;

/** The data model for items (each with its level, 0 = top): the item points and their parOf connections. */
/** Farben ändern: our colour sets by PowerPoint's ids (csTypeId) and how our drawing paints them. */
export const COLOR_SETS: Record<string, { id: string; cat: string }> = {
  "dk1-outline": { id: "urn:microsoft.com/office/officeart/2005/8/colors/accent0_1", cat: "mainScheme" },
  "dk2-outline": { id: "urn:microsoft.com/office/officeart/2005/8/colors/accent0_2", cat: "mainScheme" },
  dk2: { id: "urn:microsoft.com/office/officeart/2005/8/colors/accent0_3", cat: "mainScheme" },
  colorful: { id: "urn:microsoft.com/office/officeart/2005/8/colors/colorful1", cat: "colorful" },
  ...Object.fromEntries([1, 2, 3, 4, 5, 6].flatMap((k) => [[`accent${k}`, { id: `urn:microsoft.com/office/officeart/2005/8/colors/accent${k}_2`, cat: `accent${k}` }],
    [`accent${k}-outline`, { id: `urn:microsoft.com/office/officeart/2005/8/colors/accent${k}_1`, cat: `accent${k}` }]])),
};
export const schemeOfColors = (csId: string | null | undefined): string => Object.entries(COLOR_SETS).find(([, v]) => v.id === csId)?.[0] ?? "accent1";

export function dataModelXml(items: DItem[], o: { layoutId: string; layoutCat: string; drawingRel: string; scheme?: string }): { xml: string; ids: string[] } {
  const cs = COLOR_SETS[o.scheme ?? "accent1"] ?? COLOR_SETS.accent1;
  const doc = guid();
  const ids = items.map(() => guid());
  const pts: string[] = [`<dgm:pt modelId="${doc}" type="doc"><dgm:prSet loTypeId="${o.layoutId}" loCatId="${o.layoutCat}" qsTypeId="${QS_SIMPLE}" qsCatId="simple" csTypeId="${cs.id}" csCatId="${cs.cat}" phldr="1"/><dgm:spPr/>${emptyT}</dgm:pt>`];
  const cxns: string[] = [];
  const order = new Map<string, number>();
  items.forEach((it, i) => {
    const text = it.text ? `<dgm:t><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="de-DE"/><a:t>${esc(it.text)}</a:t></a:r></a:p></dgm:t>` : `<dgm:t><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="de-DE"/></a:p></dgm:t>`;
    pts.push(`<dgm:pt modelId="${ids[i]}"><dgm:prSet phldrT="[Text]"${it.text ? "" : ' phldr="1"'}/><dgm:spPr/>${text}</dgm:pt>`);
    // The parent: the nearest item before with a lower level, else the document.
    let parent = doc;
    for (let k = i - 1; k >= 0; k--) if (items[k].level < it.level) { parent = ids[k]; break; }
    const cxn = guid(), par = guid(), sib = guid();
    pts.push(`<dgm:pt modelId="${par}" type="parTrans" cxnId="${cxn}"><dgm:prSet/><dgm:spPr/>${emptyT}</dgm:pt>`);
    pts.push(`<dgm:pt modelId="${sib}" type="sibTrans" cxnId="${cxn}"><dgm:prSet/><dgm:spPr/>${emptyT}</dgm:pt>`);
    const ord = order.get(parent) ?? 0;
    order.set(parent, ord + 1);
    cxns.push(`<dgm:cxn modelId="${cxn}" srcId="${parent}" destId="${ids[i]}" srcOrd="${ord}" destOrd="0" parTransId="${par}" sibTransId="${sib}"/>`);
  });
  const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<dgm:dataModel xmlns:dgm="${NS_DGM}" xmlns:a="${NS_A}"><dgm:ptLst>${pts.join("")}</dgm:ptLst><dgm:cxnLst>${cxns.join("")}</dgm:cxnLst><dgm:bg/><dgm:whole/>` +
    `<dgm:extLst><a:ext uri="http://schemas.microsoft.com/office/drawing/2008/diagram"><dsp:dataModelExt xmlns:dsp="${NS_DSP}" relId="${o.drawingRel}" minVer="http://schemas.openxmlformats.org/drawingml/2006/diagram"/></a:ext></dgm:extLst></dgm:dataModel>`;
  return { xml, ids };
}

/** Our layout part: PowerPoint's id for the layout, our own minimal definition (one layout node). */
export function layoutXml(id: string, cat: string, name: string): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<dgm:layoutDef xmlns:dgm="${NS_DGM}" xmlns:a="${NS_A}" uniqueId="${id}"><dgm:title val="${esc(name)}"/><dgm:desc val=""/>` +
    `<dgm:catLst><dgm:cat type="${cat}" pri="1000"/></dgm:catLst><dgm:layoutNode name="diagram"><dgm:alg type="composite"/><dgm:shape xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:blip=""><dgm:adjLst/></dgm:shape><dgm:presOf/><dgm:constrLst/><dgm:ruleLst/></dgm:layoutNode></dgm:layoutDef>`;
}

/** Our quick style part (SmartArt-Formatvorlage "Einfach", PowerPoint's id). */
export function quickStyleXml(): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<dgm:styleDef xmlns:dgm="${NS_DGM}" xmlns:a="${NS_A}" uniqueId="${QS_SIMPLE}"><dgm:title val=""/><dgm:desc val=""/><dgm:catLst><dgm:cat type="simple" pri="10100"/></dgm:catLst>` +
    `<dgm:scene3d><a:camera prst="orthographicFront"/><a:lightRig rig="threePt" dir="t"/></dgm:scene3d><dgm:styleLbl name="node1"><dgm:scene3d><a:camera prst="orthographicFront"/><a:lightRig rig="threePt" dir="t"/></dgm:scene3d><dgm:sp3d/><dgm:txPr/><dgm:style><a:lnRef idx="2"><a:scrgbClr r="0" g="0" b="0"/></a:lnRef><a:fillRef idx="1"><a:scrgbClr r="0" g="0" b="0"/></a:fillRef><a:effectRef idx="0"><a:scrgbClr r="0" g="0" b="0"/></a:effectRef><a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef></dgm:style></dgm:styleLbl></dgm:styleDef>`;
}

/** Our colours part (Farbig - Akzentfarbe 1's id). */
export function colorsXml(scheme = "accent1"): string {
  const cs = COLOR_SETS[scheme] ?? COLOR_SETS.accent1;
  const acc = /accent(\d)/.exec(scheme)?.[1] ?? "1";
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<dgm:colorsDef xmlns:dgm="${NS_DGM}" xmlns:a="${NS_A}" uniqueId="${cs.id}"><dgm:title val=""/><dgm:desc val=""/><dgm:catLst><dgm:cat type="${cs.cat}" pri="11200"/></dgm:catLst>` +
    `<dgm:styleLbl name="node1"><dgm:fillClrLst meth="repeat"><a:schemeClr val="accent${acc}"/></dgm:fillClrLst><dgm:linClrLst meth="repeat"><a:schemeClr val="lt1"/></dgm:linClrLst><dgm:effectClrLst/><dgm:txLinClrLst/><dgm:txFillClrLst meth="repeat"><a:schemeClr val="lt1"/></dgm:txFillClrLst><dgm:txEffectClrLst/></dgm:styleLbl></dgm:colorsDef>`;
}

export interface DrawShape {
  prst: string; x: number; y: number; w: number; h: number; rot?: number; flipH?: boolean; flipV?: boolean;
  adj?: Record<string, number>;
  /** The item it shows (its text), or none (connectors, backgrounds). */
  item?: number;
  /** Look: "node" accent fill with light outline and light text; "conn" a lighter accent (arrows, connectors);
   *  "line" outline only; "text" no fill, dark text; "bg" a pale accent area. */
  look: "node" | "conn" | "line" | "text" | "bg" | "inner";
  /** Font size in points (all of a diagram's texts share one, as PowerPoint sizes them). */
  size?: number;
  anchor?: "t" | "ctr" | "b";
  algn?: "l" | "ctr" | "r";
}

function fills(scheme: string, i: number): Record<DrawShape["look"], string> {
  const acc = scheme === "colorful" ? `accent${(i % 5) + 2}` : scheme.startsWith("dk") ? (scheme.startsWith("dk1") ? "tx1" : "tx2") : `accent${/accent(\d)/.exec(scheme)?.[1] ?? 1}`;
  const outline = scheme.endsWith("-outline");
  return {
    node: outline ? `<a:solidFill><a:schemeClr val="lt1"/></a:solidFill><a:ln w="19050"><a:solidFill><a:schemeClr val="${acc}"/></a:solidFill></a:ln>` : `<a:solidFill><a:schemeClr val="${acc}"/></a:solidFill><a:ln w="19050"><a:solidFill><a:schemeClr val="lt1"/></a:solidFill></a:ln>`,
    conn: `<a:solidFill><a:schemeClr val="${acc}"><a:tint val="60000"/></a:schemeClr></a:solidFill><a:ln><a:noFill/></a:ln>`,
    line: `<a:noFill/><a:ln w="12700"><a:solidFill><a:schemeClr val="${acc}"/></a:solidFill></a:ln>`,
    text: `<a:noFill/><a:ln><a:noFill/></a:ln>`,
    inner: `<a:noFill/><a:ln><a:noFill/></a:ln>`,
    bg: `<a:solidFill><a:schemeClr val="${acc}"><a:tint val="40000"/><a:alpha val="90000"/></a:schemeClr></a:solidFill><a:ln w="12700"><a:solidFill><a:schemeClr val="${acc}"><a:tint val="40000"/></a:schemeClr></a:solidFill></a:ln>`,
  };
}


/** The drawing part: our engine's shapes, each linked to its item point by modelId. */
export function drawingXml(shapes: DrawShape[], ids: string[], texts: string[], scheme = "accent1"): string {
  const outline = scheme.endsWith("-outline");
  const sps = shapes.map((s) => {
    const FILL = fills(scheme, s.item ?? 0);
    const xf = `<a:xfrm${s.rot ? ` rot="${Math.round(s.rot * 60000)}"` : ""}${s.flipH ? ' flipH="1"' : ""}${s.flipV ? ' flipV="1"' : ""}><a:off x="${Math.round(s.x)}" y="${Math.round(s.y)}"/><a:ext cx="${Math.max(0, Math.round(s.w))}" cy="${Math.max(0, Math.round(s.h))}"/></a:xfrm>`;
    const av = s.adj ? Object.entries(s.adj).map(([k, v]) => `<a:gd name="${k}" fmla="val ${Math.round(v)}"/>`).join("") : "";
    const text = s.item !== undefined ? texts[s.item] ?? "" : "";
    const dark = s.look === "text" || s.look === "line" || s.look === "bg" || (outline && s.look === "node");
    const run = text ? `<a:r><a:rPr lang="de-DE" sz="${Math.round((s.size ?? 18) * 100)}" kern="1200">${dark ? `<a:solidFill><a:schemeClr val="tx1"/></a:solidFill>` : ""}</a:rPr><a:t>${esc(text)}</a:t></a:r>` : "";
    const ins = Math.round(Math.min(s.w, s.h) * 0.06);
    return `<dsp:sp modelId="${s.item !== undefined ? ids[s.item] : guid()}"><dsp:nvSpPr><dsp:cNvPr id="0" name=""/><dsp:cNvSpPr/></dsp:nvSpPr><dsp:spPr>${xf}<a:prstGeom prst="${s.prst}"><a:avLst>${av}</a:avLst></a:prstGeom>${FILL[s.look]}<a:effectLst/></dsp:spPr>` +
      `<dsp:style><a:lnRef idx="2"><a:scrgbClr r="0" g="0" b="0"/></a:lnRef><a:fillRef idx="1"><a:scrgbClr r="0" g="0" b="0"/></a:fillRef><a:effectRef idx="0"><a:scrgbClr r="0" g="0" b="0"/></a:effectRef><a:fontRef idx="minor"><a:schemeClr val="${dark ? "tx1" : "lt1"}"/></a:fontRef></dsp:style>` +
      `<dsp:txBody><a:bodyPr spcFirstLastPara="0" vert="horz" wrap="square" lIns="${ins}" tIns="${ins}" rIns="${ins}" bIns="${ins}" numCol="1" spcCol="1270" anchor="${s.anchor ?? "ctr"}" anchorCtr="0"><a:noAutofit/></a:bodyPr><a:lstStyle/>` +
      `<a:p><a:pPr marL="0" lvl="0" indent="0" algn="${s.algn ?? "ctr"}"><a:lnSpc><a:spcPct val="90000"/></a:lnSpc><a:spcBef><a:spcPct val="0"/></a:spcBef><a:spcAft><a:spcPct val="35000"/></a:spcAft><a:buNone/></a:pPr>${run}${run ? "" : '<a:endParaRPr lang="de-DE"/>'}</a:p></dsp:txBody></dsp:sp>`;
  }).join("");
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<dsp:drawing xmlns:dgm="${NS_DGM}" xmlns:dsp="${NS_DSP}" xmlns:a="${NS_A}"><dsp:spTree><dsp:nvGrpSpPr><dsp:cNvPr id="0" name=""/><dsp:cNvGrpSpPr/></dsp:nvGrpSpPr><dsp:grpSpPr/>${sps}</dsp:spTree></dsp:drawing>`;
}

/** The slide's graphic frame for the diagram. */
export function frameXml(id: number, name: string, box: { x: number; y: number; w: number; h: number }, rids: { dm: string; lo: string; qs: string; cs: string }): string {
  return `<p:graphicFrame xmlns:p="${PptxDom.NS_P}" xmlns:a="${NS_A}" xmlns:r="${PptxDom.NS_R}"><p:nvGraphicFramePr><p:cNvPr id="${id}" name="${esc(name)}"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>` +
    `<p:xfrm><a:off x="${Math.round(box.x)}" y="${Math.round(box.y)}"/><a:ext cx="${Math.round(box.w)}" cy="${Math.round(box.h)}"/></p:xfrm>` +
    `<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/diagram"><dgm:relIds xmlns:dgm="${NS_DGM}" r:dm="${rids.dm}" r:lo="${rids.lo}" r:qs="${rids.qs}" r:cs="${rids.cs}"/></a:graphicData></a:graphic></p:graphicFrame>`;
}
