// What a slide shows (D-070): every shape with its place, look and text after PowerPoint's
// inheritance. A placeholder takes what it does not set itself from the matching placeholder of
// its layout, and that from the master: position and size, geometry, fill and line, the text
// frame (bodyPr) and the text styles of each level. Text levels start at the master's text styles
// (titleStyle for titles, otherStyle for date, footer and slide number, bodyStyle for the rest);
// shapes that are not placeholders start at the presentation's defaultTextStyle.

import type { Fill, PictureFormat, ShapeInfo } from "@raabe/text-engine";
import { custGeomDef, formatMatrix, readAdjust, resolveStyle, type FormatMatrix } from "../shapes";
import { readColor, readShapeProps } from "../drawingml";
import { NS_A, NS_R, at, kid, kids, num } from "./dom";
import { RT, type Pptx } from "./package";

export interface ThemeInfo {
  /** Scheme colours after the master's colour map: dk1 … folHlink, plus tx1/bg1/tx2/bg2. */
  colors: Record<string, string>;
  major: string;
  minor: string;
  name: string;
  matrix: FormatMatrix;
}

export interface BodyProps {
  /** Insets in EMU. */
  l: number; t: number; r: number; b: number;
  anchor: "t" | "ctr" | "b";
  anchorCtr: boolean;
  wrap: boolean;
  autofit: "none" | "norm" | "shape";
  /** normAutofit: font scale (1 = 100 %) and line spacing reduction (0.1 = 10 %). */
  fontScale: number;
  lnSpcReduction: number;
  vert?: string;
  numCol: number;
  spcCol: number;
  /** Text rotation (bodyPr rot) in degrees. */
  rot: number;
  upright: boolean;
}

export interface Bullet {
  kind: "char" | "auto" | "pic";
  char?: string;
  font?: string;
  /** Hex colour; undefined: the first run's colour. */
  color?: string;
  /** Size relative to the first run's (1 = 100 %), or absolute points. */
  sizePct?: number;
  sizePts?: number;
  autoType?: string;
  startAt?: number;
}

export interface RunStyle {
  size: number;
  bold: boolean;
  italic: boolean;
  underline: string;
  strike: string;
  font: string;
  color: string;
  /** Text colour transparency 0…1. */
  alpha: number;
  highlight?: string;
  baseline: number;
  caps: "none" | "all" | "small";
  /** Zeichenabstand in points. */
  spacing: number;
  lang?: string;
  link?: string;
  /** The symbol font (a:sym) for characters of the private use area (Wingdings arrows from AutoKorrektur). */
  symFont?: string;
  /** Textkontur: the run's outline (a:ln with a solid fill), width in points. */
  outline?: { color: string; width: number };
}

export interface TextRun extends RunStyle {
  kind: "r" | "br" | "fld" | "math";
  /** An equation (a14:m): its m:oMathPara or m:oMath element. */
  math?: Element;
  text: string;
  /** For fields: slidenum, datetime… */
  fieldType?: string;
}

export interface ParaStyle {
  lvl: number;
  algn: "l" | "ctr" | "r" | "just" | "dist";
  /** EMU. */
  marL: number;
  indent: number;
  /** Line spacing: percent (1 = single) or points. */
  lnPct?: number;
  lnPts?: number;
  /** Space before / after in points, or as share of the line. */
  spcBefPts?: number;
  spcBefPct?: number;
  spcAftPts?: number;
  spcAftPct?: number;
  bullet: Bullet | null;
  defTab: number;
  tabs: { pos: number; algn: string }[];
  rtl: boolean;
}

export interface TextPara extends ParaStyle {
  runs: TextRun[];
  /** The style of the paragraph mark (end of paragraph: empty paragraphs' height). */
  end: RunStyle;
}

export interface TextInfo {
  body: BodyProps;
  paras: TextPara[];
  /** Whether the text is empty (a placeholder then shows its prompt in Normalansicht). */
  empty: boolean;
  /** The prompt text of an empty placeholder, one entry per paragraph. */
  prompt?: string[];
}

export interface PicInfo {
  path?: string;
  /** Cropped share per side (srcRect, 0…1). */
  crop: { l: number; t: number; r: number; b: number };
}

export interface ShapeView {
  el: Element;
  kind: "sp" | "pic" | "cxnSp" | "grpSp" | "graphicFrame" | "contentPart" | "other";
  id: number;
  name: string;
  hidden: boolean;
  ph?: { type: string; idx?: number; orient?: string; sz?: string };
  /** Place on the slide in EMU (for group members: on the slide, after the group's scaling). */
  x: number; y: number; w: number; h: number;
  rot: number;
  flipH: boolean;
  flipV: boolean;
  /** Look for the text engine's shape renderer. */
  format: PictureFormat;
  info: ShapeInfo;
  geometry: string;
  text?: TextInfo;
  pic?: PicInfo;
  /** graphicFrame: "table", "chart", "diagram", "ole" … */
  frame?: string;
  /** Chart part path. */
  chart?: string;
  children?: ShapeView[];
  /** Whether the shape is a text box (txBox="1"). */
  textBox: boolean;
  /** The layout or master element it inherits from (placeholders). */
  inherits: Element[];
}

export interface SlideView {
  path: string;
  layout?: string;
  master?: string;
  theme: ThemeInfo;
  background: { fill: Fill; image?: string };
  shapes: ShapeView[];
  /** Layout and master shapes drawn behind the slide's (not placeholders), unless hidden. */
  behind: ShapeView[];
}

const DEFAULT_COLORS: Record<string, string> = {
  dk1: "000000", lt1: "FFFFFF", dk2: "44546A", lt2: "E7E6E6", accent1: "4472C4", accent2: "ED7D31", accent3: "A5A5A5",
  accent4: "FFC000", accent5: "5B9BD5", accent6: "70AD47", hlink: "0563C1", folHlink: "954F72",
};

export function readThemeInfo(pptx: Pptx, themePath: string | undefined, clrMap: Element | undefined): ThemeInfo {
  const doc = themePath && pptx.has(themePath) ? pptx.doc(themePath) : null;
  const colors: Record<string, string> = { ...DEFAULT_COLORS };
  let major = "Calibri Light", minor = "Calibri", name = "Office";
  if (doc) {
    name = doc.documentElement.getAttribute("name") ?? name;
    const scheme = doc.getElementsByTagNameNS(NS_A, "clrScheme")[0];
    for (const c of kids(scheme)) {
      const v = kids(c)[0];
      const hex = v?.localName === "sysClr" ? v.getAttribute("lastClr") : v?.getAttribute("val");
      if (hex && /^[0-9A-Fa-f]{6}$/.test(hex)) colors[c.localName] = hex.toUpperCase();
    }
    const font = (n: string) => doc.getElementsByTagNameNS(NS_A, n)[0]?.getElementsByTagNameNS(NS_A, "latin")[0]?.getAttribute("typeface");
    major = font("majorFont") || major;
    minor = font("minorFont") || minor;
  }
  const map = (k: string, def: string) => clrMap?.getAttribute(k) ?? def;
  const base = { ...colors };
  colors.tx1 = base[map("tx1", "dk1")] ?? base.dk1;
  colors.bg1 = base[map("bg1", "lt1")] ?? base.lt1;
  colors.tx2 = base[map("tx2", "dk2")] ?? base.dk2;
  colors.bg2 = base[map("bg2", "lt2")] ?? base.lt2;
  // readColor maps tx1 → dk1 etc. itself: give it the mapped colours under those names.
  colors.dk1 = colors.tx1; colors.lt1 = colors.bg1; colors.dk2 = colors.tx2; colors.lt2 = colors.bg2;
  return { colors, major, minor, name, matrix: formatMatrix(doc) };
}

// ----- placeholders ---------------------------------------------------------------------------------

export function phOf(sp: Element): ShapeView["ph"] | undefined {
  const ph = at(nvPr(sp), "ph") ?? kid(nvPr(sp), "ph");
  if (!ph) return undefined;
  return { type: ph.getAttribute("type") ?? "obj", idx: num(ph, "idx"), orient: ph.getAttribute("orient") ?? undefined, sz: ph.getAttribute("sz") ?? undefined };
}

function nvPr(sp: Element): Element | undefined {
  const nv = kids(sp).find((c) => c.localName.startsWith("nv"));
  return kid(nv, "nvPr");
}

const TITLE = new Set(["title", "ctrTitle"]);
const FOOT = new Set(["dt", "ftr", "sldNum", "hdr"]);
/** Body-like placeholder types match each other by index. */
const sameFamily = (a: string, b: string) => a === b || (TITLE.has(a) && TITLE.has(b)) || (!TITLE.has(a) && !FOOT.has(a) && !TITLE.has(b) && !FOOT.has(b));

/** The placeholder of a layout (or master) that a slide's (layout's) placeholder inherits from. */
export function matchPlaceholder(tree: Element | undefined, ph: NonNullable<ShapeView["ph"]>, master: boolean): Element | undefined {
  if (!tree) return undefined;
  const cands = allShapes(tree).filter((s) => phOf(s));
  if (master) {
    // The master has one placeholder per kind: title, body, dt, ftr, sldNum.
    const want = TITLE.has(ph.type) ? "title" : FOOT.has(ph.type) ? ph.type : "body";
    return cands.find((s) => phOf(s)!.type === want) ?? (want !== "body" ? undefined : cands.find((s) => phOf(s)!.type === "obj"));
  }
  if (ph.idx !== undefined) {
    const byIdx = cands.find((s) => phOf(s)!.idx === ph.idx && sameFamily(phOf(s)!.type, ph.type));
    if (byIdx) return byIdx;
    const anyIdx = cands.find((s) => phOf(s)!.idx === ph.idx);
    if (anyIdx && !TITLE.has(ph.type)) return anyIdx;
  }
  return cands.find((s) => phOf(s)!.type === ph.type) ?? (TITLE.has(ph.type) ? cands.find((s) => TITLE.has(phOf(s)!.type)) : undefined);
}

const SHAPES = ["sp", "pic", "graphicFrame", "grpSp", "cxnSp"];

/**
 * The shapes of a tree. Shapes inside mc:AlternateContent (equations, newer objects) are drawn from
 * the Fallback PowerPoint writes for older readers (an equation's picture), as long as we do not
 * understand the Choice.
 */
export function allShapes(tree: Element): Element[] {
  const out: Element[] = [];
  for (const c of kids(tree)) {
    if (SHAPES.includes(c.localName)) out.push(c);
    else if (c.localName === "AlternateContent") out.push(...alternateShapes(c));
  }
  return out;
}

const NS_A14 = "http://schemas.microsoft.com/office/drawing/2010/main";

/**
 * The shapes of an mc:AlternateContent that are drawn and edited: the Choice of an equation (a text
 * box with a14:m, which is understood), else the Fallback PowerPoint writes for older readers.
 */
export function alternateShapes(ac: Element): Element[] {
  const choice = kid(ac, "Choice");
  const pick = choice && choice.getAttribute("Requires") === "a14" && choice.getElementsByTagNameNS(NS_A14, "m").length ? choice : kid(ac, "Fallback");
  return kids(pick).filter((x) => SHAPES.includes(x.localName));
}

export function spTreeOf(doc: Document): Element | undefined {
  return at(doc.documentElement, "cSld/spTree");
}

// ----- text styles -----------------------------------------------------------------------------------

/** Paragraph properties of level `lvl` from a list style (a:lstStyle or p:bodyStyle …). */
function levelOf(list: Element | undefined, lvl: number): Element[] {
  if (!list) return [];
  const out: Element[] = [];
  const def = kid(list, "defPPr");
  if (def) out.push(def);
  const l = kid(list, `lvl${lvl + 1}pPr`);
  if (l) out.push(l);
  return out;
}

const BU_TYPE = ["buNone", "buChar", "buAutoNum", "buBlip"];
const BU_CLR = ["buClrTx", "buClr"];
const BU_SZ = ["buSzTx", "buSzPct", "buSzPts"];
const BU_FONT = ["buFontTx", "buFont"];
const FILL = ["noFill", "solidFill", "gradFill", "blipFill", "pattFill", "grpFill"];

interface MergedP {
  attrs: Record<string, string>;
  groups: Record<string, Element>;
  rpr: MergedR;
}

interface MergedR {
  attrs: Record<string, string>;
  groups: Record<string, Element>;
}

const groupOf = (name: string): string =>
  BU_TYPE.includes(name) ? "buType" : BU_CLR.includes(name) ? "buClr" : BU_SZ.includes(name) ? "buSz" : BU_FONT.includes(name) ? "buFont" : FILL.includes(name) ? "fill" : name;

function mergeR(into: MergedR, rpr: Element | undefined): void {
  if (!rpr) return;
  for (const a of Array.from(rpr.attributes)) into.attrs[a.localName] = a.value;
  for (const c of kids(rpr)) into.groups[groupOf(c.localName)] = c;
}

function mergeP(into: MergedP, ppr: Element | undefined): void {
  if (!ppr) return;
  for (const a of Array.from(ppr.attributes)) into.attrs[a.localName] = a.value;
  for (const c of kids(ppr)) {
    if (c.localName === "defRPr") mergeR(into.rpr, c);
    else into.groups[groupOf(c.localName)] = c;
  }
}

export interface TextContext {
  theme: ThemeInfo;
  /** List styles from lowest priority to highest (master text style, master ph, layout ph, own). */
  lists: (Element | undefined)[];
  /** Text colour a shape style gives (fontRef). */
  fontColor?: string;
}

function spacing(el: Element | undefined): { pts?: number; pct?: number } {
  if (!el) return {};
  const pts = kid(el, "spcPts"), pct = kid(el, "spcPct");
  if (pts) return { pts: (num(pts, "val") ?? 0) / 100 };
  if (pct) return { pct: (num(pct, "val") ?? 0) / 100000 };
  return {};
}

function fontName(v: string | undefined, theme: ThemeInfo, fallback: string): string {
  if (!v) return fallback;
  if (v.startsWith("+mj")) return theme.major;
  if (v.startsWith("+mn")) return theme.minor;
  return v;
}

function runStyle(m: MergedR, theme: ThemeInfo, fontColor?: string): RunStyle {
  const a = m.attrs;
  const fill = m.groups.fill;
  let color = fontColor ?? theme.colors.tx1, alpha = 0;
  if (fill?.localName === "solidFill") {
    const c = readColor(fill, theme.colors);
    if (c) { color = c.hex; alpha = c.transparency; }
  } else if (fill?.localName === "gradFill") {
    const gs = fill.getElementsByTagNameNS(NS_A, "gs")[0];
    const c = gs ? readColor(gs, theme.colors) : null;
    if (c) color = c.hex;
  } else if (fill?.localName === "noFill") alpha = 1;
  const hl = m.groups.highlight ? readColor(m.groups.highlight, theme.colors)?.hex : undefined;
  const latin = m.groups.latin?.getAttribute("typeface") ?? undefined;
  const link = m.groups.hlinkClick?.getAttributeNS(NS_R, "id") ?? undefined;
  const ln = m.groups.ln, lnFill = ln ? kid(ln, "solidFill") : undefined;
  const lnColor = lnFill ? readColor(lnFill, theme.colors)?.hex : undefined;
  return {
    size: (Number(a.sz ?? 1800)) / 100,
    bold: a.b === "1" || a.b === "true",
    italic: a.i === "1" || a.i === "true",
    underline: a.u ?? "none",
    strike: a.strike ?? "noStrike",
    font: fontName(latin, theme, theme.minor),
    color: link && !fill ? theme.colors.hlink : color,
    alpha,
    highlight: hl,
    baseline: Number(a.baseline ?? 0) / 100000,
    caps: (a.cap as RunStyle["caps"]) ?? "none",
    spacing: Number(a.spc ?? 0) / 100,
    lang: a.lang,
    link,
    ...(lnColor ? { outline: { color: lnColor, width: (num(ln, "w") ?? 12700) / 12700 } } : {}),
    ...(m.groups.sym?.getAttribute("typeface") ? { symFont: m.groups.sym.getAttribute("typeface")! } : {}),
  };
}

function paraStyle(m: MergedP, lvl: number, theme: ThemeInfo): ParaStyle {
  const a = m.attrs, g = m.groups;
  const ln = spacing(g.lnSpc), bef = spacing(g.spcBef), aft = spacing(g.spcAft);
  let bullet: Bullet | null = null;
  const bt = g.buType;
  if (bt && bt.localName !== "buNone") {
    bullet = { kind: bt.localName === "buChar" ? "char" : bt.localName === "buAutoNum" ? "auto" : "pic" };
    if (bt.localName === "buChar") bullet.char = bt.getAttribute("char") ?? "•";
    if (bt.localName === "buAutoNum") { bullet.autoType = bt.getAttribute("type") ?? "arabicPeriod"; bullet.startAt = num(bt, "startAt") ?? 1; }
    if (g.buFont?.localName === "buFont") bullet.font = fontName(g.buFont.getAttribute("typeface") ?? undefined, theme, theme.minor);
    if (g.buClr?.localName === "buClr") bullet.color = readColor(g.buClr, theme.colors)?.hex;
    if (g.buSz?.localName === "buSzPct") bullet.sizePct = (num(g.buSz, "val") ?? 100000) / 100000;
    if (g.buSz?.localName === "buSzPts") bullet.sizePts = (num(g.buSz, "val") ?? 1800) / 100;
  }
  const tabs = kids(g.tabLst, "tab").map((t) => ({ pos: num(t, "pos") ?? 0, algn: t.getAttribute("algn") ?? "l" }));
  return {
    lvl,
    algn: (a.algn as ParaStyle["algn"]) ?? "l",
    marL: Number(a.marL ?? 0), indent: Number(a.indent ?? 0),
    lnPct: ln.pct ?? (ln.pts === undefined ? 1 : undefined), lnPts: ln.pts,
    spcBefPts: bef.pts, spcBefPct: bef.pct, spcAftPts: aft.pts, spcAftPct: aft.pct,
    bullet, defTab: Number(a.defTabSz ?? 914400), tabs, rtl: a.rtl === "1",
  };
}

/** The merged paragraph properties of level `lvl` through the context's list styles, then `ppr`. */
export function levelProps(ctx: TextContext, lvl: number, ppr?: Element): MergedP {
  const m: MergedP = { attrs: {}, groups: {}, rpr: { attrs: {}, groups: {} } };
  for (const list of ctx.lists) for (const p of levelOf(list, lvl)) mergeP(m, p);
  mergeP(m, ppr);
  return m;
}

/** Reads a text body (p:txBody or a:txBody) with the inherited styles. */
export function readText(txBody: Element | undefined, ctx: TextContext, body: BodyProps): TextInfo {
  const paras: TextPara[] = [];
  let empty = true;
  for (const p of kids(txBody, "p")) {
    const ppr = kid(p, "pPr");
    const lvl = Math.min(8, Math.max(0, num(ppr, "lvl") ?? 0));
    const m = levelProps(ctx, lvl, ppr);
    const ps = paraStyle(m, lvl, ctx.theme);
    const runs: TextRun[] = [];
    for (const r of kids(p)) {
      if (r.localName === "m" && r.namespaceURI === NS_A14) {
        // An equation: its size from the first run inside it.
        const inner = r.getElementsByTagNameNS(NS_A, "rPr")[0];
        const rm: MergedR = { attrs: { ...m.rpr.attrs }, groups: { ...m.rpr.groups } };
        mergeR(rm, inner);
        const st = runStyle(rm, ctx.theme, ctx.fontColor && !rm.groups.fill ? ctx.fontColor : undefined);
        runs.push({ ...st, kind: "math", text: "", math: r.firstElementChild ?? undefined });
        empty = false;
        continue;
      }
      if (r.localName !== "r" && r.localName !== "br" && r.localName !== "fld") continue;
      const rm: MergedR = { attrs: { ...m.rpr.attrs }, groups: { ...m.rpr.groups } };
      mergeR(rm, kid(r, "rPr"));
      const st = runStyle(rm, ctx.theme, ctx.fontColor && !rm.groups.fill ? ctx.fontColor : undefined);
      const text = r.localName === "br" ? "\n" : kid(r, "t")?.textContent ?? "";
      if (text && r.localName !== "br") empty = false;
      if (r.localName === "br") empty = false;
      runs.push({ ...st, kind: r.localName as TextRun["kind"], text, fieldType: r.localName === "fld" ? r.getAttribute("type") ?? undefined : undefined });
    }
    const em: MergedR = { attrs: { ...m.rpr.attrs }, groups: { ...m.rpr.groups } };
    mergeR(em, kid(p, "endParaRPr"));
    paras.push({ ...ps, runs, end: runStyle(em, ctx.theme, ctx.fontColor && !em.groups.fill ? ctx.fontColor : undefined) });
  }
  if (paras.length > 1) empty = false;
  return { body, paras, empty };
}

const BODY_DEFAULT: BodyProps = { l: 91440, t: 45720, r: 91440, b: 45720, anchor: "t", anchorCtr: false, wrap: true, autofit: "none", fontScale: 1, lnSpcReduction: 0, numCol: 1, spcCol: 0, rot: 0, upright: false };

/** bodyPr through the chain (lowest priority first). */
export function readBodyProps(chain: (Element | undefined)[]): BodyProps {
  const b: BodyProps = { ...BODY_DEFAULT };
  for (const el of chain) {
    if (!el) continue;
    const v = (n: string) => num(el, n);
    if (v("lIns") !== undefined) b.l = v("lIns")!;
    if (v("tIns") !== undefined) b.t = v("tIns")!;
    if (v("rIns") !== undefined) b.r = v("rIns")!;
    if (v("bIns") !== undefined) b.b = v("bIns")!;
    const an = el.getAttribute("anchor");
    if (an === "t" || an === "ctr" || an === "b") b.anchor = an;
    if (el.hasAttribute("anchorCtr")) b.anchorCtr = el.getAttribute("anchorCtr") === "1";
    if (el.hasAttribute("wrap")) b.wrap = el.getAttribute("wrap") !== "none";
    if (el.hasAttribute("vert")) b.vert = el.getAttribute("vert") === "horz" ? undefined : el.getAttribute("vert")!;
    if (v("numCol") !== undefined) b.numCol = v("numCol")!;
    if (v("spcCol") !== undefined) b.spcCol = v("spcCol")!;
    if (v("rot") !== undefined) b.rot = v("rot")! / 60000;
    if (el.hasAttribute("upright")) b.upright = el.getAttribute("upright") === "1";
    if (kid(el, "noAutofit")) { b.autofit = "none"; b.fontScale = 1; b.lnSpcReduction = 0; }
    const na = kid(el, "normAutofit");
    if (na) { b.autofit = "norm"; b.fontScale = (num(na, "fontScale") ?? 100000) / 100000; b.lnSpcReduction = (num(na, "lnSpcReduction") ?? 0) / 100000; }
    if (kid(el, "spAutoFit")) { b.autofit = "shape"; b.fontScale = 1; b.lnSpcReduction = 0; }
  }
  return b;
}

// ----- shapes ---------------------------------------------------------------------------------------------

/** German PowerPoint's prompts for empty placeholders in Normalansicht. */
export const PROMPTS: Record<string, string> = {
  title: "Titel durch Klicken hinzufügen",
  ctrTitle: "Titel durch Klicken hinzufügen",
  subTitle: "Untertitel durch Klicken hinzufügen",
  body: "Text durch Klicken hinzufügen",
  obj: "Text durch Klicken hinzufügen",
  pic: "Bild durch Klicken auf Symbol hinzufügen",
  tbl: "Tabelle durch Klicken auf Symbol hinzufügen",
  chart: "Diagramm durch Klicken auf Symbol hinzufügen",
  dgm: "SmartArt-Grafik durch Klicken auf Symbol hinzufügen",
  media: "Medien durch Klicken auf Symbol hinzufügen",
  clipArt: "ClipArt durch Klicken auf Symbol hinzufügen",
};

interface Inherit {
  /** Placeholder elements this one inherits from, nearest first (layout's, then master's). */
  chain: Element[];
  /** The master text style the placeholder starts from. */
  base: Element | undefined;
}

export class SlideResolver {
  readonly theme: ThemeInfo;
  readonly layoutDoc?: Document;
  readonly masterDoc?: Document;
  constructor(readonly pptx: Pptx, readonly path: string) {
    const kind = pathKind(path);
    const layout = kind === "slide" ? pptx.layoutOf(path) : kind === "layout" ? path : undefined;
    const master = kind === "master" ? path : layout ? pptx.masterOf(layout) : undefined;
    this.layoutPath = kind === "slide" ? layout : undefined;
    this.masterPath = kind === "master" ? undefined : master;
    this.layoutDoc = this.layoutPath && pptx.has(this.layoutPath) ? pptx.doc(this.layoutPath) : undefined;
    this.masterDoc = master && pptx.has(master) ? pptx.doc(master) : undefined;
    const own = pptx.doc(path).documentElement;
    let clrMap = kid(this.masterDoc?.documentElement ?? own, "clrMap");
    // A slide or layout may override the colour map (p:clrMapOvr/a:overrideClrMapping).
    for (const d of [this.layoutDoc?.documentElement, own]) {
      const ov = at(d, "clrMapOvr/overrideClrMapping");
      if (ov) clrMap = ov;
    }
    this.theme = readThemeInfo(pptx, master ? pptx.themeOf(master) : undefined, clrMap);
    this.own = own;
  }

  readonly layoutPath?: string;
  readonly masterPath?: string;
  private own: Element;

  private txStyle(name: "titleStyle" | "bodyStyle" | "otherStyle"): Element | undefined {
    return at((this.masterDoc ?? this.own.ownerDocument).documentElement, `txStyles/${name}`);
  }

  private inherit(ph: NonNullable<ShapeView["ph"]>): Inherit {
    const chain: Element[] = [];
    const kind = pathKind(this.path);
    if (kind === "slide") {
      const l = matchPlaceholder(this.layoutDoc ? spTreeOf(this.layoutDoc) : undefined, ph, false);
      if (l) chain.push(l);
      const lph = l ? phOf(l)! : ph;
      const m = matchPlaceholder(this.masterDoc ? spTreeOf(this.masterDoc) : undefined, lph, true);
      if (m) chain.push(m);
    } else if (kind === "layout") {
      const m = matchPlaceholder(this.masterDoc ? spTreeOf(this.masterDoc) : undefined, ph, true);
      if (m) chain.push(m);
    }
    const base = TITLE.has(ph.type) ? this.txStyle("titleStyle") : FOOT.has(ph.type) ? this.txStyle("otherStyle") : this.txStyle("bodyStyle");
    return { chain, base };
  }

  /** Everything the slide draws. */
  view(): SlideView {
    const tree = spTreeOf(this.own.ownerDocument);
    const shapes = tree ? this.shapes(tree, null) : [];
    const behind: ShapeView[] = [];
    const showMaster = this.own.getAttribute("showMasterSp") !== "0";
    if (pathKind(this.path) === "slide" && showMaster) {
      const layoutShow = this.layoutDoc?.documentElement.getAttribute("showMasterSp") !== "0";
      if (this.masterDoc && layoutShow) behind.push(...this.decoShapes(this.masterDoc, this.masterPath!));
      if (this.layoutDoc) behind.push(...this.decoShapes(this.layoutDoc, this.layoutPath!));
    } else if (pathKind(this.path) === "layout" && showMaster && this.masterDoc) behind.push(...this.decoShapes(this.masterDoc, this.masterPath!));
    return { path: this.path, layout: this.layoutPath, master: this.masterPath, theme: this.theme, background: this.background(), shapes, behind };
  }

  /** Non-placeholder shapes of a layout or master (its design), resolved in that part's context. */
  private decoShapes(doc: Document, path: string): ShapeView[] {
    const tree = spTreeOf(doc);
    if (!tree) return [];
    const r = new SlideResolver(this.pptx, path);
    return r.shapes(tree, null).filter((s) => !s.ph);
  }

  background(): { fill: Fill; image?: string } {
    for (const [doc, path] of [[this.own.ownerDocument, this.path], [this.layoutDoc, this.layoutPath], [this.masterDoc, this.masterPath]] as const) {
      const bg = at(doc?.documentElement, "cSld/bg");
      if (!bg) continue;
      const pr = kid(bg, "bgPr");
      if (pr) {
        const f = kids(pr).find((c) => FILL.includes(c.localName));
        if (f?.localName === "blipFill") {
          const rid = at(f, "blip")?.getAttributeNS(NS_R, "embed");
          return { fill: { kind: "solid", color: "FFFFFF" }, image: rid && path ? this.pptx.target(path, rid) : undefined };
        }
        const fill = f ? readFillLocal(f, this.theme) : undefined;
        if (fill) return { fill };
      }
      const ref = kid(bg, "bgRef");
      if (ref) {
        const idx = num(ref, "idx") ?? 1001;
        const color = readColor(ref, this.theme.colors)?.hex ?? "FFFFFF";
        const el = idx >= 1001 ? this.theme.matrix.bgFills[idx - 1001] : this.theme.matrix.fills[idx - 1];
        const fill = el ? readFillLocal(el, { ...this.theme, colors: { ...this.theme.colors, phClr: color } }) : undefined;
        return { fill: fill ?? { kind: "solid", color } };
      }
    }
    return { fill: { kind: "solid", color: this.theme.colors.bg1 ?? "FFFFFF" } };
  }

  /** The shapes of a tree (spTree or grpSp); `group` maps child coordinates onto the slide. */
  shapes(tree: Element, group: GroupMap | null): ShapeView[] {
    const out: ShapeView[] = [];
    for (const el of allShapes(tree)) {
      const v = this.shape(el, group);
      if (v) out.push(v);
    }
    return out;
  }

  shape(el: Element, group: GroupMap | null): ShapeView | null {
    const kind = (["sp", "pic", "cxnSp", "grpSp", "graphicFrame"].includes(el.localName) ? el.localName : "other") as ShapeView["kind"];
    const nv = kids(el).find((c) => c.localName.startsWith("nv"));
    const c = kid(nv, "cNvPr");
    const ph = phOf(el);
    const inh = ph ? this.inherit(ph) : { chain: [], base: undefined };
    const spPr = kid(el, kind === "grpSp" ? "grpSpPr" : "spPr");
    const xfrmOf = (e: Element) => kind === "graphicFrame" && e === el ? kid(e, "xfrm") : kid(kid(e, e.localName === "grpSp" ? "grpSpPr" : "spPr"), "xfrm");
    let xfrm = xfrmOf(el);
    if (!xfrm || !kid(xfrm, "off")) for (const p of inh.chain) { const x = xfrmOf(p); if (x && kid(x, "off")) { xfrm = x; break; } }
    const off = kid(xfrm, "off"), ext = kid(xfrm, "ext");
    let x = num(off, "x") ?? 0, y = num(off, "y") ?? 0, w = num(ext, "cx") ?? 0, h = num(ext, "cy") ?? 0;
    if (group) [x, y, w, h] = group.map(x, y, w, h);
    const rot = (num(xfrm, "rot") ?? 0) / 60000 + (group?.rot ?? 0);
    const flipH = xfrm?.getAttribute("flipH") === "1", flipV = xfrm?.getAttribute("flipV") === "1";
    // Look: own spPr, else the inherited placeholders' (only what the own one lacks).
    const style = kid(el, "style");
    const styleBits = style ? resolveStyle(style, this.theme.colors, this.theme.matrix) : {};
    const props: PictureFormat[] = [];
    for (const p of [...inh.chain].reverse()) { const s = kid(p, "spPr"); if (s) props.push(readShapeProps(s, this.theme.colors)); }
    if (spPr) props.push(readShapeProps(spPr, this.theme.colors, styleBits.styleLine));
    const merged = Object.assign({}, ...props) as Record<string, unknown>;
    delete merged.rotation; delete merged.flipH; delete merged.flipV;
    const format = merged as PictureFormat;
    const prst = [spPr, ...inh.chain.map((p) => kid(p, "spPr"))].map((s) => kid(s, "prstGeom")).find(Boolean);
    const cust = [spPr, ...inh.chain.map((p) => kid(p, "spPr"))].map((s) => kid(s, "custGeom")).find(Boolean);
    const geometry = cust ? "custom" : prst?.getAttribute("prst") ?? (kind === "cxnSp" ? "line" : "rect");
    const textBox = kid(nv, "cNvSpPr")?.getAttribute("txBox") === "1";
    const info: ShapeInfo = {
      kind: kind === "cxnSp" ? "connector" : textBox ? "textbox" : "shape",
      adjust: readAdjust(prst), body: { l: 7.2, t: 3.6, r: 7.2, b: 3.6, anchor: "t", wrap: true, autofit: "none" },
      ...styleBits, custom: cust ? custGeomDef(cust) : undefined,
    };
    const v: ShapeView = {
      el, kind, id: num(c, "id") ?? 0, name: c?.getAttribute("name") ?? "", hidden: c?.getAttribute("hidden") === "1",
      ph, x, y, w, h, rot, flipH, flipV, format: { ...format, geometry: geometry === "rect" ? undefined : geometry } as PictureFormat, info, geometry,
      textBox, inherits: inh.chain,
    };
    if (kind === "sp" || (kind === "pic" && kid(el, "txBody"))) {
      const bodyChain = [...inh.chain].reverse().map((p) => kid(kid(p, "txBody"), "bodyPr"));
      bodyChain.push(kid(kid(el, "txBody"), "bodyPr"));
      const body = readBodyProps(bodyChain);
      const lists: (Element | undefined)[] = [];
      if (ph) {
        lists.push(inh.base);
        for (const p of [...inh.chain].reverse()) lists.push(kid(kid(p, "txBody"), "lstStyle"));
      } else {
        lists.push(at(this.pptx.presentation, "defaultTextStyle"));
      }
      // The shape style's text colour (fontRef: white on a new shape) comes after the inherited text
      // styles and before the shape's own list style and runs.
      if (styleBits.fontColor) lists.push(colorList(styleBits.fontColor));
      lists.push(kid(kid(el, "txBody"), "lstStyle"));
      const text = readText(kid(el, "txBody"), { theme: this.theme, lists }, body);
      if (ph && text.empty) text.prompt = this.prompt(ph, inh.chain[0]);
      v.text = text;
    }
    // A shape filled with a picture (also the picture PowerPoint keeps as an equation's fallback).
    const spBlip = kind === "sp" ? kid(spPr, "blipFill") : undefined;
    if (spBlip) {
      const blip = kid(spBlip, "blip");
      const rid = blip?.getAttributeNS(NS_R, "embed");
      const src = kid(spBlip, "srcRect");
      const pc = (n: string) => (num(src, n) ?? 0) / 100000;
      v.pic = { path: rid ? this.pptx.target(this.path, rid) : undefined, crop: { l: pc("l"), t: pc("t"), r: pc("r"), b: pc("b") } };
    }
    if (kind === "pic") {
      const blip = at(el, "blipFill/blip");
      const rid = blip?.getAttributeNS(NS_R, "embed") ?? blip?.getAttributeNS(NS_R, "link");
      const src = at(el, "blipFill/srcRect");
      const pc = (n: string) => (num(src, n) ?? 0) / 100000;
      v.pic = { path: rid ? this.pptx.target(this.path, rid) : undefined, crop: { l: pc("l"), t: pc("t"), r: pc("r"), b: pc("b") } };
    }
    if (kind === "graphicFrame") {
      const gd = at(el, "graphic/graphicData");
      const uri = gd?.getAttribute("uri") ?? "";
      v.frame = uri.endsWith("/table") ? "table" : uri.endsWith("/chart") ? "chart" : uri.endsWith("/diagram") ? "diagram" : uri.endsWith("/ole") ? "ole" : "other";
      if (v.frame === "diagram") {
        // SmartArt: drawn from the shapes PowerPoint caches in the diagram's drawing part
        // (dsp:drawing, linked from the data model's dataModelExt relId), placed in the frame.
        const relIds = kids(gd).find((k) => k.localName === "relIds");
        const dm = relIds ? this.pptx.target(this.path, relIds.getAttributeNS(NS_R, "dm") ?? "") : undefined;
        const ext = dm && this.pptx.has(dm) ? this.pptx.doc(dm).getElementsByTagNameNS("http://schemas.microsoft.com/office/drawing/2008/diagram", "dataModelExt")[0] : undefined;
        const drawing = ext ? this.pptx.target(this.path, ext.getAttribute("relId") ?? "") : undefined;
        const tree = drawing && this.pptx.has(drawing) ? kids(this.pptx.doc(drawing).documentElement).find((k) => k.localName === "spTree") : undefined;
        if (tree) v.children = this.shapes(tree, new GroupMap(x, y, w, h, 0, 0, w, h, rot));
      }
      if (v.frame === "chart") {
        const ch = kids(gd).find((k) => k.localName === "chart");
        const rid = ch?.getAttributeNS(NS_R, "id");
        v.chart = rid ? this.pptx.target(this.path, rid) : undefined;
      }
    }
    if (kind === "grpSp") {
      const gx = kid(spPr, "xfrm");
      const chOff = kid(gx, "chOff"), chExt = kid(gx, "chExt");
      const map = new GroupMap(x, y, w, h, num(chOff, "x") ?? 0, num(chOff, "y") ?? 0, num(chExt, "cx") ?? w, num(chExt, "cy") ?? h, rot);
      v.children = this.shapes(el, map);
    }
    return v;
  }

  /** The prompt of an empty placeholder: the layout's own text when it has a custom prompt. */
  private prompt(ph: NonNullable<ShapeView["ph"]>, layoutPh: Element | undefined): string[] | undefined {
    const custom = layoutPh && at(nvPr(layoutPh), "ph")?.getAttribute("hasCustomPrompt") === "1";
    if (custom) {
      const texts = kids(kid(layoutPh, "txBody"), "p").map((p) => kids(p, "r").map((r) => kid(r, "t")?.textContent ?? "").join(""));
      if (texts.some(Boolean)) return texts;
    }
    if (FOOT.has(ph.type)) return undefined;
    if (pathKind(this.path) !== "slide") return undefined;
    const t = PROMPTS[ph.type] ?? PROMPTS.body;
    return [t];
  }
}

/** A list style whose every level has this text colour (the shape style's fontRef, a table style's text). */
export function colorList(hex: string, bold?: boolean): Element {
  return new DOMParser().parseFromString(`<a:lstStyle xmlns:a="${NS_A}"><a:defPPr><a:defRPr${bold === undefined ? "" : ` b="${bold ? 1 : 0}"`}><a:solidFill><a:srgbClr val="${hex}"/></a:solidFill></a:defRPr></a:defPPr></a:lstStyle>`, "application/xml").documentElement;
}

export class GroupMap {
  constructor(readonly x: number, readonly y: number, readonly w: number, readonly h: number,
    readonly cx: number, readonly cy: number, readonly cw: number, readonly ch: number, readonly rot: number) {}
  get sx(): number { return this.cw ? this.w / this.cw : 1; }
  get sy(): number { return this.ch ? this.h / this.ch : 1; }
  map(x: number, y: number, w: number, h: number): [number, number, number, number] {
    return [this.x + (x - this.cx) * this.sx, this.y + (y - this.cy) * this.sy, w * this.sx, h * this.sy];
  }
  /** Slide coordinates back into the group's child coordinates. */
  unmap(x: number, y: number, w: number, h: number): [number, number, number, number] {
    return [this.cx + (x - this.x) / this.sx, this.cy + (y - this.y) / this.sy, w / this.sx, h / this.sy];
  }
}

function readFillLocal(f: Element, theme: ThemeInfo): Fill | undefined {
  switch (f.localName) {
    case "noFill": return { kind: "none" };
    case "solidFill": { const c = readColor(f, theme.colors); return c ? { kind: "solid", color: c.hex, transparency: c.transparency || undefined } : undefined; }
    case "gradFill": {
      const stops = kids(kid(f, "gsLst"), "gs").map((g) => {
        const c = readColor(g, theme.colors) ?? { hex: "000000", transparency: 0 };
        return { pos: (num(g, "pos") ?? 0) / 100000, color: c.hex, transparency: c.transparency || undefined };
      });
      const lin = kid(f, "lin");
      return { kind: "gradient", stops, angle: lin ? (num(lin, "ang") ?? 0) / 60000 : undefined, path: kid(f, "path")?.getAttribute("path") ?? undefined };
    }
    case "pattFill": return { kind: "solid", color: readColor(kid(f, "bgClr"), theme.colors)?.hex ?? "FFFFFF" };
    default: return undefined;
  }
}

export function pathKind(path: string): "slide" | "layout" | "master" | "notes" | "notesMaster" | "other" {
  if (/slideLayouts\//.test(path)) return "layout";
  if (/slideMasters\//.test(path)) return "master";
  if (/notesSlides\//.test(path)) return "notes";
  if (/notesMasters\//.test(path)) return "notesMaster";
  if (/slides\//.test(path)) return "slide";
  return "other";
}

/** A slide's view in one call. */
export function slideView(pptx: Pptx, path: string): SlideView {
  return new SlideResolver(pptx, path).view();
}

export { RT };
