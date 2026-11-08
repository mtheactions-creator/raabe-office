// Creating SmartArt (D-077): the parts PowerPoint writes for a diagram, made by us. The data model (dgm:dataModel)
// holds the items: a document point, a point per item, and per item a parOf connection with its parent and its
// transition points (parTrans, sibTrans), as PowerPoint writes them; the layout, quick style and colours parts
// carry PowerPoint's ids (loTypeId, qsTypeId, csTypeId) so that PowerPoint lays the diagram out with its own
// definitions when it opens the file, while their content is ours (no definition of Microsoft's is copied); the
// drawing part (dsp:drawing) is our layout engine's result, which RaabePresent draws and PowerPoint shows until
// it lays the diagram out again.

import { PptxDom, readColor } from "@raabe/ooxml";
import { SMARTART_COLORS, SMARTART_STYLES, type ClrList, type ColorDef, type StyleDef, type StyleLbl } from "./smartart-styles.gen";

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
/**
 * Farben ändern and SmartArt-Formatvorlagen (D-101): a colour set and a quick style by the last part of
 * PowerPoint's id ("accent1_2", "simple1"), with the parameters recorded from PowerPoint
 * (smartart-styles.gen.ts). Keys of Step 8b ("accent1", "accent1-outline", "colorful" ...) still work.
 */
const LEGACY: Record<string, string> = { "dk1-outline": "accent0_1", "dk2-outline": "accent0_2", dk2: "accent0_3", colorful: "colorful1",
  ...Object.fromEntries([1, 2, 3, 4, 5, 6].flatMap((k) => [[`accent${k}`, `accent${k}_2`], [`accent${k}-outline`, `accent${k}_1`]])) };
const tail = (id: string) => id.split("/").pop() ?? id;
export const colorKey = (k: string | null | undefined): string => (k ? LEGACY[k] ?? k : "accent1_2");
export const colorDef = (k: string | null | undefined): ColorDef => SMARTART_COLORS.find((c) => tail(c.id) === colorKey(k)) ?? SMARTART_COLORS.find((c) => tail(c.id) === "accent1_2")!;
export const styleDef = (k: string | null | undefined): StyleDef => SMARTART_STYLES.find((c) => tail(c.id) === (k || "simple1")) ?? SMARTART_STYLES[0];
/** The colour set's key of a data model's csTypeId (accent1_2 when unknown). */
export const schemeOfColors = (csId: string | null | undefined): string => (csId && SMARTART_COLORS.some((c) => c.id === csId) ? tail(csId) : "accent1_2");
/** The quick style's key of a data model's qsTypeId (simple1 when unknown). */
export const styleOfQuick = (qsId: string | null | undefined): string => (qsId && SMARTART_STYLES.some((c) => c.id === qsId) ? tail(qsId) : "simple1");
/** The style labels our drawing uses, by the look of a drawn shape. */
const ROLE: Record<string, string | null> = { node: "node1", conn: "sibTrans2D1", line: "parChTrans1D1", text: "revTx", bg: "bgShp", inner: null };
const LABELS = ["node1", "sibTrans2D1", "parChTrans1D1", "revTx", "bgShp"];

export function dataModelXml(items: DItem[], o: { layoutId: string; layoutCat: string; drawingRel: string; scheme?: string; style?: string; rtl?: boolean }): { xml: string; ids: string[] } {
  const cs = colorDef(o.scheme), qs = styleDef(o.style);
  const doc = guid();
  const ids = items.map(() => guid());
  const pts: string[] = [`<dgm:pt modelId="${doc}" type="doc"><dgm:prSet loTypeId="${o.layoutId}" loCatId="${o.layoutCat}" qsTypeId="${qs.id}" qsCatId="${qs.cat}" csTypeId="${cs.id}" csCatId="${cs.cat}" phldr="1">${o.rtl ? "<dgm:presLayoutVars><dgm:dir val=\"rev\"/></dgm:presLayoutVars>" : ""}</dgm:prSet><dgm:spPr/>${emptyT}</dgm:pt>`];
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

const attrs = (o: Record<string, string>) => Object.entries(o).map(([k, v]) => ` ${k}="${esc(v)}"`).join("");

/** A recorded label's 3-D shape (sp3d attributes, bevels) as DrawingML. */
function sp3dXml(l: StyleLbl | undefined, tag: string, contour?: string): string {
  const sp = l?.sp3d;
  if (!sp) return `<${tag}/>`;
  const flat = Object.fromEntries(Object.entries(sp).filter(([, v]) => typeof v === "string")) as Record<string, string>;
  const bevel = (k: string) => (typeof sp[k] === "object" ? `<a:${k}${attrs(sp[k] as Record<string, string>)}/>` : "");
  return `<${tag}${attrs(flat)}>${bevel("bevelT")}${bevel("bevelB")}${contour && flat.contourW ? `<a:contourClr>${contour}</a:contourClr>` : ""}</${tag}>`;
}

/** The scene of a label: its camera (or the quick style's) and light. */
function sceneXml(st: StyleDef, l: StyleLbl | undefined, tag: string): string {
  const cam = l?.camera ?? st.camera ?? "orthographicFront";
  const light = l?.light ?? ["threePt", "t"];
  return `<${tag}><a:camera prst="${cam}"/><a:lightRig rig="${light[0]}" dir="${light[1]}"/></${tag}>`;
}

/** Our quick style part: PowerPoint's id and, for the labels our drawing uses, its references and 3-D (D-101). */
export function quickStyleXml(key = "simple1"): string {
  const st = styleDef(key);
  const lbls = LABELS.filter((n) => st.labels[n]).map((n) => {
    const l = st.labels[n];
    const ref = (k: "lnRef" | "fillRef" | "effectRef") => `<a:${k} idx="${l.refs[k] ?? 0}"><a:scrgbClr r="0" g="0" b="0"/></a:${k}>`;
    return `<dgm:styleLbl name="${n}">${sceneXml(st, l, "dgm:scene3d")}${sp3dXml(l, "dgm:sp3d")}<dgm:txPr/><dgm:style>${ref("lnRef")}${ref("fillRef")}${ref("effectRef")}<a:fontRef idx="${l.font ?? "minor"}"/></dgm:style></dgm:styleLbl>`;
  }).join("");
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<dgm:styleDef xmlns:dgm="${NS_DGM}" xmlns:a="${NS_A}" uniqueId="${st.id}"><dgm:title val=""/><dgm:desc val=""/><dgm:catLst><dgm:cat type="${st.cat}" pri="10100"/></dgm:catLst>` +
    `${sceneXml(st, undefined, "dgm:scene3d")}${lbls}</dgm:styleDef>`;
}

/** A recorded colour ("accent1 tint=40000", "FFFFFF") as DrawingML, with more modifiers after its own. */
function clrXml(spec: string, extra = ""): string {
  const [v, ...mods] = spec.split(" ");
  const tag = /^[0-9A-Fa-f]{6}$/.test(v) ? "srgbClr" : "schemeClr";
  const m = mods.map((x) => { const [k, val] = x.split("="); return `<a:${k} val="${val}"/>`; }).join("");
  return `<a:${tag} val="${v}">${m}${extra}</a:${tag}>`;
}

/** Our colours part: PowerPoint's id and, for the labels our drawing uses, its colour lists (D-101). */
export function colorsXml(scheme = "accent1_2"): string {
  const cs = colorDef(scheme);
  const list = (tag: string, l: ClrList | undefined) => (l ? `<dgm:${tag} meth="${l.meth}"${l.hueDir ? ` hueDir="${l.hueDir}"` : ""}>${l.clr.map((c) => clrXml(c)).join("")}</dgm:${tag}>` : `<dgm:${tag}/>`);
  const lbls = LABELS.filter((n) => cs.labels[n]).map((n) => {
    const l = cs.labels[n];
    return `<dgm:styleLbl name="${n}">${list("fillClrLst", l.fill)}${list("linClrLst", l.line)}${list("effectClrLst", l.effect)}${list("txLinClrLst", l.txLine)}${list("txFillClrLst", l.txFill)}${list("txEffectClrLst", l.txEffect)}</dgm:styleLbl>`;
  }).join("");
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<dgm:colorsDef xmlns:dgm="${NS_DGM}" xmlns:a="${NS_A}" uniqueId="${cs.id}"><dgm:title val=""/><dgm:desc val=""/><dgm:catLst><dgm:cat type="${cs.cat}" pri="11200"/></dgm:catLst>${lbls}</dgm:colorsDef>`;
}

function rgbToHsl(hex: string): [number, number, number] {
  const [r, g, b] = [0, 2, 4].map((i) => parseInt(hex.slice(i, i + 2), 16) / 255);
  const max = Math.max(r, g, b), min = Math.min(r, g, b), l = (max + min) / 2;
  if (max === min) return [0, 0, l];
  const d = max - min, s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
  const h = max === r ? (g - b) / d + (g < b ? 6 : 0) : max === g ? (b - r) / d + 2 : (r - g) / d + 4;
  return [h * 60, s, l];
}

/** What a theme gives a colour spec (hex and opacity), through our colour reader. */
function resolveSpec(spec: string, theme: Record<string, string>): { hex: string; alpha: number } {
  const el = new DOMParser().parseFromString(`<a:solidFill xmlns:a="${NS_A}">${clrXml(spec)}</a:solidFill>`, "application/xml").documentElement;
  const c = readColor(el, theme);
  return { hex: c?.hex ?? "000000", alpha: 1 - (c?.transparency ?? 0) };
}

/**
 * The colour of node i of n from a colour list, as PowerPoint writes it into its drawing (measured, D-101):
 * repeat takes the list in order and keeps its last; span goes from the first colour to the last over the nodes
 * and cycle there and back, both written as the first colour with hue, saturation, luminance and alpha offsets.
 */
export function pickColor(l: ClrList | undefined, i: number, n: number, theme: Record<string, string>): string | null {
  if (!l || !l.clr.length) return null;
  if (l.meth === "repeat" || l.clr.length === 1) return clrXml(l.clr[Math.min(i, l.clr.length - 1)]);
  const k = Math.ceil(n / 2);
  // Cycle steps by 2/n there and back (five nodes: 0, 0.4, 0.8, 0.8, 0.4 of the way, measured), span by 1/(n-1).
  const t = l.meth === "cycle" ? ((i < k ? i : 2 * k - 1 - i) * 2) / n : n <= 1 ? 0 : i / (n - 1);
  const a = resolveSpec(l.clr[0], theme), b = resolveSpec(l.clr[l.clr.length - 1], theme);
  const [h1, s1, l1] = rgbToHsl(a.hex), [h2, s2, l2] = rgbToHsl(b.hex);
  let dh = h2 - h1;
  // Without a direction PowerPoint takes the plain difference of the two hues (accent 5 to 6: -202 degrees, measured).
  if (l.hueDir === "ccw") { while (dh > 0) dh -= 360; } else if (l.hueDir === "cw") { while (dh < 0) dh += 360; }
  const off = `<a:hueOff val="${Math.round(dh * 60000 * t)}"/><a:satOff val="${Math.round((s2 - s1) * 100000 * t)}"/><a:lumOff val="${Math.round((l2 - l1) * 100000 * t)}"/><a:alphaOff val="${Math.round((b.alpha - a.alpha) * 100000 * t)}"/>`;
  return clrXml(l.clr[0], off);
}

/** A theme style element (fill, line, effect list) with its placeholder colour (phClr) replaced by a colour. */
function withColor(styleEl: Element | undefined, color: string | null): string {
  if (!styleEl) return "";
  let xml = new XMLSerializer().serializeToString(styleEl).replace(/ xmlns(:\w+)?="[^"]*"/g, "");
  if (!color) return xml;
  const m = /^<a:(\w+) val="([^"]+)">([\s\S]*)<\/a:\1>$/.exec(color);
  if (!m) return xml;
  const [, tag, val, mods] = m;
  xml = xml.replace(/<a:schemeClr val="phClr"\s*\/>/g, `<a:${tag} val="${val}">${mods}</a:${tag}>`)
    .replace(/<a:schemeClr val="phClr">([\s\S]*?)<\/a:schemeClr>/g, (_, inner: string) => `<a:${tag} val="${val}">${mods}${inner}</a:${tag}>`);
  return xml;
}

/** The theme's format matrix our drawing takes fills, lines and effects from (Office's when not given). */
export interface DrawTheme { colors: Record<string, string>; fills: readonly Element[]; lines: readonly Element[]; effects: readonly Element[] }

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


/** Shapes that look the same mirrored (Von rechts nach links moves them without flipping). */
const SYMMETRIC = new Set(["rect", "roundRect", "ellipse", "round2SameRect", "snip2SameRect", "hexagon", "octagon", "diamond", "plus", "donut", "triangle", "trapezoid", "flowChartProcess", "flowChartAlternateProcess", "can", "frame", "plaque", "line", "straightConnector1"]);

/**
 * The drawing part: our engine's shapes, each linked to its item point by modelId. With the theme, each shape
 * looks as PowerPoint draws it (D-101): its look's style label gives the colours (the set's lists, the node's
 * place among the nodes) and, by the quick style's references, the theme's fill, line and effect styles with
 * the placeholder colour replaced, and the label's 3-D scene and shape. Von rechts nach links mirrors the
 * arrangement in the frame (width `o.width`).
 */
export function drawingXml(shapes: DrawShape[], ids: string[], texts: string[], scheme = "accent1_2", o: { style?: string; theme?: DrawTheme; rtl?: boolean; width?: number } = {}): string {
  const legacy = !o.theme;
  const outline = colorKey(scheme).endsWith("_1") || colorKey(scheme).startsWith("accent0_1") || colorKey(scheme) === "accent0_2";
  const cs = colorDef(scheme), st = styleDef(o.style);
  const n = Math.max(1, texts.length);
  const counter = new Map<string, number>();
  if (o.rtl && o.width) {
    shapes = shapes.map((s) => ({ ...s, x: o.width! - s.x - s.w, rot: s.rot ? -s.rot : s.rot, flipH: SYMMETRIC.has(s.prst) ? s.flipH : !s.flipH }));
  }
  const sps = shapes.map((s) => {
    const FILL = fills(legacy ? scheme : "accent1", s.item ?? 0);
    const role = ROLE[s.look];
    const k = role ? (s.item ?? (counter.get(role) ?? 0)) : 0;
    if (role && s.item === undefined) counter.set(role, k + 1);
    const cl = role ? cs.labels[role] : undefined, sl = role ? st.labels[role] : undefined;
    const th = o.theme;
    let look = FILL[s.look];
    let refs = { lnRef: 2, fillRef: 1, effectRef: 0 };
    let txColor: string | null = null;
    if (th && role) {
      const fillC = pickColor(cl?.fill, k, n, th.colors), lineC = pickColor(cl?.line, k, n, th.colors);
      txColor = pickColor(cl?.txFill, k, n, th.colors) ?? `<a:schemeClr val="${role === "node1" || role === "sibTrans2D1" ? "lt1" : "tx1"}"/>`;
      refs = { lnRef: sl?.refs.lnRef ?? 0, fillRef: sl?.refs.fillRef ?? 0, effectRef: sl?.refs.effectRef ?? 0 };
      const fill = refs.fillRef > 0 && fillC ? withColor(th.fills[refs.fillRef - 1], fillC) || `<a:solidFill>${fillC}</a:solidFill>` : "<a:noFill/>";
      const line = refs.lnRef > 0 && lineC ? withColor(th.lines[refs.lnRef - 1], lineC) || `<a:ln><a:solidFill>${lineC}</a:solidFill></a:ln>` : "<a:ln><a:noFill/></a:ln>";
      const eff = refs.effectRef > 0 ? withColor(th.effects[refs.effectRef - 1], fillC) || "<a:effectLst/>" : "<a:effectLst/>";
      const three = sl && (sl.sp3d || sl.light || sl.camera || st.camera !== "orthographicFront") ? sceneXml(st, sl, "a:scene3d") + sp3dXml(sl, "a:sp3d", lineC ?? undefined) : "";
      look = fill + line + eff.replace(/^<a:effectStyle>|<\/a:effectStyle>$/g, "") + three;
    } else if (th) look = FILL.inner + "<a:effectLst/>";
    const xf = `<a:xfrm${s.rot ? ` rot="${Math.round(s.rot * 60000)}"` : ""}${s.flipH ? ' flipH="1"' : ""}${s.flipV ? ' flipV="1"' : ""}><a:off x="${Math.round(s.x)}" y="${Math.round(s.y)}"/><a:ext cx="${Math.max(0, Math.round(s.w))}" cy="${Math.max(0, Math.round(s.h))}"/></a:xfrm>`;
    const av = s.adj ? Object.entries(s.adj).map(([k, v]) => `<a:gd name="${k}" fmla="val ${Math.round(v)}"/>`).join("") : "";
    const text = s.item !== undefined ? texts[s.item] ?? "" : "";
    const dark = s.look === "text" || s.look === "line" || s.look === "bg" || (outline && s.look === "node");
    const color = txColor ? `<a:solidFill>${txColor}</a:solidFill>` : dark ? `<a:solidFill><a:schemeClr val="tx1"/></a:solidFill>` : "";
    const run = text ? `<a:r><a:rPr lang="de-DE" sz="${Math.round((s.size ?? 18) * 100)}" kern="1200">${color}</a:rPr><a:t>${esc(text)}</a:t></a:r>` : "";
    const ins = Math.round(Math.min(s.w, s.h) * 0.06);
    return `<dsp:sp modelId="${s.item !== undefined ? ids[s.item] : guid()}"><dsp:nvSpPr><dsp:cNvPr id="0" name=""/><dsp:cNvSpPr/></dsp:nvSpPr><dsp:spPr>${xf}<a:prstGeom prst="${s.prst}"><a:avLst>${av}</a:avLst></a:prstGeom>${legacy ? FILL[s.look] + "<a:effectLst/>" : look}</dsp:spPr>` +
      `<dsp:style><a:lnRef idx="${refs.lnRef}"><a:scrgbClr r="0" g="0" b="0"/></a:lnRef><a:fillRef idx="${refs.fillRef}"><a:scrgbClr r="0" g="0" b="0"/></a:fillRef><a:effectRef idx="${refs.effectRef}"><a:scrgbClr r="0" g="0" b="0"/></a:effectRef><a:fontRef idx="minor"><a:schemeClr val="${dark ? "tx1" : "lt1"}"/></a:fontRef></dsp:style>` +
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
