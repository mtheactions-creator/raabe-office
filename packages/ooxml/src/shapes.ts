// Shapes and text boxes (wps:wsp): the shape style resolved against the theme's format
// matrix, the text frame (wps:bodyPr), custom geometry, and the XML Word writes for new and
// changed shapes (recorded with tools/feel-lab/probe_word_shapes.py).

import { DEFAULT_SHAPE_BODY, type Fill, type PictureFormat, type ShapeBody, type ShapeInfo } from "@raabe/text-engine";
import { colorXml, readColor, readEffects, readFill, readShapeProps, shapePropsXml } from "./drawingml";

export const NS_WPS = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape";
const NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main";
const EMU = 12700;

const kids = (el: Element) => Array.from(el.children);
const kid = (el: Element | undefined, name: string) => (el ? kids(el).find((c) => c.localName === name) : undefined);
const num = (el: Element | undefined, a: string) => {
  const v = el?.getAttribute(a);
  return v === null || v === undefined || v === "" ? undefined : Number(v);
};

// The format matrix of Word's Office themes (2013–2022 and 2023 share it): fill, line and
// effect styles that a shape style's references (fillRef idx …) point to.
const PH = '<a:schemeClr val="phClr"/>';
const grad = (stops: [number, string][]) => `<a:gradFill rotWithShape="1"><a:gsLst>${stops.map(([pos, mods]) => `<a:gs pos="${pos}"><a:schemeClr val="phClr">${mods}</a:schemeClr></a:gs>`).join("")}</a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill>`;
const m = (lumMod: number, satMod: number, x: string) => `<a:lumMod val="${lumMod}"/><a:satMod val="${satMod}"/>${x}`;
const line = (w: number) => `<a:ln w="${w}" cap="flat" cmpd="sng" algn="ctr"><a:solidFill>${PH}</a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln>`;
export const OFFICE_FMT_SCHEME = `<a:fmtScheme name="Office"><a:fillStyleLst><a:solidFill>${PH}</a:solidFill>` +
  grad([[0, m(110000, 105000, '<a:tint val="67000"/>')], [50000, m(105000, 103000, '<a:tint val="73000"/>')], [100000, m(105000, 109000, '<a:tint val="81000"/>')]]) +
  grad([[0, '<a:satMod val="103000"/><a:lumMod val="102000"/><a:tint val="94000"/>'], [50000, '<a:satMod val="110000"/><a:lumMod val="100000"/><a:shade val="100000"/>'], [100000, '<a:lumMod val="99000"/><a:satMod val="120000"/><a:shade val="78000"/>']]) +
  `</a:fillStyleLst><a:lnStyleLst>${line(6350)}${line(12700)}${line(19050)}</a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle>` +
  '<a:effectStyle><a:effectLst><a:outerShdw blurRad="57150" dist="19050" dir="5400000" algn="ctr" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="63000"/></a:srgbClr></a:outerShdw></a:effectLst></a:effectStyle></a:effectStyleLst>' +
  `<a:bgFillStyleLst><a:solidFill>${PH}</a:solidFill><a:solidFill><a:schemeClr val="phClr"><a:tint val="95000"/><a:satMod val="170000"/></a:schemeClr></a:solidFill>` +
  grad([[0, '<a:tint val="93000"/><a:satMod val="150000"/><a:shade val="98000"/><a:lumMod val="102000"/>'], [50000, '<a:tint val="98000"/><a:satMod val="130000"/><a:shade val="90000"/><a:lumMod val="103000"/>'], [100000, '<a:shade val="63000"/><a:satMod val="120000"/>']]) +
  "</a:bgFillStyleLst></a:fmtScheme>";

export interface FormatMatrix {
  readonly fills: readonly Element[];
  readonly lines: readonly Element[];
  readonly effects: readonly Element[];
  readonly bgFills: readonly Element[];
}

/** The theme's format matrix (a:fmtScheme), or Office's when the document has no theme. */
export function formatMatrix(themeDoc: Document | null): FormatMatrix {
  let scheme = themeDoc?.getElementsByTagNameNS(NS_A, "fmtScheme")[0];
  if (!scheme) scheme = new DOMParser().parseFromString(`<r xmlns:a="${NS_A}">${OFFICE_FMT_SCHEME}</r>`, "application/xml").documentElement.firstElementChild!;
  const list = (name: string) => kids(kid(scheme, name) ?? scheme!);
  return {
    fills: list("fillStyleLst"), lines: list("lnStyleLst"),
    effects: list("effectStyleLst").map((e) => kid(e, "effectLst") ?? e), bgFills: list("bgFillStyleLst"),
  };
}

/** What a shape style (wps:style) gives a shape: fill, outline, shadow and text colour. */
export function resolveStyle(style: Element | undefined, theme: Record<string, string>, matrix: FormatMatrix): Pick<ShapeInfo, "styleFill" | "styleLine" | "styleShadow" | "fontColor"> {
  if (!style) return {};
  const ref = (name: string) => {
    const r = kid(style, name);
    const color = readColor(r, theme);
    return { idx: num(r, "idx") ?? 0, color: color?.hex, el: r };
  };
  const out: { -readonly [K in keyof Pick<ShapeInfo, "styleFill" | "styleLine" | "styleShadow" | "fontColor">]?: ShapeInfo[K] } = {};
  const withPh = (hex: string | undefined) => ({ ...theme, phClr: hex ?? "000000" });
  const ln = ref("lnRef");
  const lnEl = ln.idx > 0 ? matrix.lines[ln.idx - 1] : undefined;
  if (lnEl && ln.color) {
    const c = readColor(kid(lnEl, "solidFill"), withPh(ln.color));
    if (c) out.styleLine = { color: c.hex, width: (num(lnEl, "w") ?? 9525) / EMU };
  }
  const fl = ref("fillRef");
  const flEl = fl.idx >= 1001 ? matrix.bgFills[fl.idx - 1001] : fl.idx > 0 ? matrix.fills[fl.idx - 1] : undefined;
  if (flEl && fl.color) out.styleFill = readFill(flEl, withPh(fl.color));
  const fx = ref("effectRef");
  const fxEl = fx.idx > 0 ? matrix.effects[fx.idx - 1] : undefined;
  if (fxEl) {
    const e = readEffects(fxEl, withPh(fx.color));
    if (e.shadow) out.styleShadow = e.shadow;
  }
  const font = ref("fontRef");
  if (font.color) out.fontColor = font.color;
  return out;
}

/** A text frame's settings (wps:bodyPr). */
export function readBody(body: Element | undefined, textBox: boolean): ShapeBody {
  const d = DEFAULT_SHAPE_BODY;
  if (!body) return { ...d, anchor: textBox ? "t" : d.anchor };
  const ins = (a: string, def: number) => (num(body, a) ?? def * EMU) / EMU;
  const anchor = body.getAttribute("anchor");
  const autofit = kid(body, "spAutoFit") ? "shape" : kid(body, "normAutofit") ? "normal" : "none";
  const vert = body.getAttribute("vert") ?? undefined;
  return {
    l: ins("lIns", d.l), t: ins("tIns", d.t), r: ins("rIns", d.r), b: ins("bIns", d.b),
    anchor: anchor === "ctr" ? "ctr" : anchor === "b" ? "b" : anchor === "t" ? "t" : textBox ? "t" : "t",
    wrap: body.getAttribute("wrap") !== "none", autofit, vert: vert && vert !== "horz" ? vert : undefined,
  };
}

/** Adjust values of a preset geometry (a:avLst gd "val n"). */
export function readAdjust(prstGeom: Element | undefined): Record<string, number> | undefined {
  const av = kid(prstGeom, "avLst");
  if (!av) return undefined;
  const out: Record<string, number> = {};
  for (const gd of kids(av)) {
    const mm = /^val\s+(-?\d+)/.exec(gd.getAttribute("fmla") ?? "");
    const name = gd.getAttribute("name");
    if (mm && name) out[name] = Number(mm[1]);
  }
  return Object.keys(out).length ? out : undefined;
}

/** Custom geometry (a:custGeom) in the engine's definition syntax (geometry.ts). */
export function custGeomDef(cust: Element): string {
  const lines: string[] = [];
  for (const gd of kids(kid(cust, "avLst") ?? cust).filter((x) => x.localName === "gd")) {
    const mm = /^val\s+(-?\S+)/.exec(gd.getAttribute("fmla") ?? "");
    if (mm) lines.push(`av ${gd.getAttribute("name")} ${mm[1]}`);
  }
  for (const gd of kids(kid(cust, "gdLst") ?? cust).filter((x) => x.localName === "gd")) lines.push(`gd ${gd.getAttribute("name")} ${gd.getAttribute("fmla")}`);
  const rect = kid(cust, "rect");
  if (rect) lines.push(`rect ${["l", "t", "r", "b"].map((a) => rect.getAttribute(a) ?? "0").join(" ")}`);
  const pt = (e: Element | undefined) => `${e?.getAttribute("x") ?? 0} ${e?.getAttribute("y") ?? 0}`;
  for (const path of kids(kid(cust, "pathLst") ?? cust).filter((x) => x.localName === "path")) {
    const attrs: string[] = [];
    if (path.getAttribute("w")) attrs.push(`w=${path.getAttribute("w")}`);
    if (path.getAttribute("h")) attrs.push(`h=${path.getAttribute("h")}`);
    if (path.getAttribute("fill")) attrs.push(`fill=${path.getAttribute("fill")}`);
    if (path.getAttribute("stroke") === "0" || path.getAttribute("stroke") === "false") attrs.push("stroke=false");
    const cmds: string[] = [];
    for (const c of kids(path)) {
      const pts = kids(c).filter((x) => x.localName === "pt");
      if (c.localName === "moveTo") cmds.push(`M ${pt(pts[0])}`);
      else if (c.localName === "lnTo") cmds.push(`L ${pt(pts[0])}`);
      else if (c.localName === "cubicBezTo") cmds.push(`C ${pts.map(pt).join(" ")}`);
      else if (c.localName === "quadBezTo") cmds.push(`Q ${pts.map(pt).join(" ")}`);
      else if (c.localName === "arcTo") cmds.push(`A ${c.getAttribute("wR")} ${c.getAttribute("hR")} ${c.getAttribute("stAng")} ${c.getAttribute("swAng")}`);
      else if (c.localName === "close") cmds.push("Z");
    }
    lines.push(`path ${attrs.join(" ")}: ${cmds.join(" ")}`);
  }
  return lines.join("\n");
}

// ----- writing ------------------------------------------------------------------------------

/** The wps:style Word gives a new shape (Word 365, any Office theme). */
export const DEFAULT_SHAPE_STYLE = '<wps:style><a:lnRef idx="2"><a:schemeClr val="accent1"><a:shade val="15000"/></a:schemeClr></a:lnRef><a:fillRef idx="1"><a:schemeClr val="accent1"/></a:fillRef>' +
  '<a:effectRef idx="0"><a:schemeClr val="accent1"/></a:effectRef><a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef></wps:style>';
/** … and a new line or connector. */
export const DEFAULT_LINE_STYLE = '<wps:style><a:lnRef idx="1"><a:schemeClr val="accent1"/></a:lnRef><a:fillRef idx="0"><a:schemeClr val="accent1"/></a:fillRef>' +
  '<a:effectRef idx="0"><a:schemeClr val="accent1"/></a:effectRef><a:fontRef idx="minor"><a:schemeClr val="tx1"/></a:fontRef></wps:style>';

/** wps:bodyPr as Word writes it for a shape or text box. */
export function bodyXml(b: ShapeBody, original?: string): string {
  const e = (pt: number) => Math.round(pt * EMU);
  const set = (xml: string, a: string, v: string) => (new RegExp(`\\s${a}="[^"]*"`).test(xml) ? xml.replace(new RegExp(`(\\s${a}=")[^"]*(")`), `$1${v}$2`) : xml.replace(/^<wps:bodyPr/, `<wps:bodyPr ${a}="${v}"`));
  let open = original ? (/^<wps:bodyPr[^>]*?\/?>/.exec(original)?.[0] ?? "<wps:bodyPr>").replace(/\/>$/, ">") : '<wps:bodyPr rot="0" spcFirstLastPara="0" vertOverflow="overflow" horzOverflow="overflow" vert="horz" wrap="square" lIns="91440" tIns="45720" rIns="91440" bIns="45720" numCol="1" spcCol="0" rtlCol="0" fromWordArt="0" anchor="ctr" anchorCtr="0" forceAA="0" compatLnSpc="1">';
  open = set(open, "lIns", String(e(b.l)));
  open = set(open, "tIns", String(e(b.t)));
  open = set(open, "rIns", String(e(b.r)));
  open = set(open, "bIns", String(e(b.b)));
  open = set(open, "anchor", b.anchor);
  open = set(open, "wrap", b.wrap ? "square" : "none");
  open = set(open, "vert", b.vert ?? "horz");
  let inner = original ? (/^<wps:bodyPr[^>]*?>([\s\S]*)<\/wps:bodyPr>$/.exec(original)?.[1] ?? "") : '<a:prstTxWarp prst="textNoShape"><a:avLst/></a:prstTxWarp>';
  inner = inner.replace(/<a:(noAutofit|spAutoFit|normAutofit)[^>]*\/>|<a:normAutofit[^>]*>[\s\S]*?<\/a:normAutofit>/g, "");
  inner += b.autofit === "shape" ? "<a:spAutoFit/>" : b.autofit === "normal" ? "<a:normAutofit/>" : "<a:noAutofit/>";
  return `${open}${inner}</wps:bodyPr>`;
}

/** a:avLst content for a shape's adjust values. */
export function adjustXml(adj: Readonly<Record<string, number>> | undefined): string {
  return Object.entries(adj ?? {}).map(([k, v]) => `<a:gd name="${k}" fmla="val ${Math.round(v)}"/>`).join("");
}

/** A fill as DrawingML (for shapes whose fill was changed in RaabeWrite). */
export function fillXml(f: Fill): string {
  if (f.kind === "none") return "<a:noFill/>";
  if (f.kind === "solid") return `<a:solidFill>${colorXml(f.color, f.transparency ?? 0)}</a:solidFill>`;
  if (f.kind === "gradient") {
    const gs = f.stops.map((s) => `<a:gs pos="${Math.round(s.pos * 100000)}">${colorXml(s.color, s.transparency ?? 0)}</a:gs>`).join("");
    const how = f.path ? `<a:path path="${f.path}"><a:fillToRect l="50000" t="50000" r="50000" b="50000"/></a:path>` : `<a:lin ang="${Math.round((f.angle ?? 90) * 60000)}" scaled="1"/>`;
    return `<a:gradFill flip="none" rotWithShape="1"><a:gsLst>${gs}</a:gsLst>${how}<a:tileRect/></a:gradFill>`;
  }
  return `<a:solidFill>${colorXml(f.color)}</a:solidFill>`;
}

/** wps:spPr for a shape: shapePropsXml with the adjust values and a changed fill. */
export function shapeSpPr(f: PictureFormat, s: ShapeInfo, cx: number, cy: number, custom?: string): string {
  const fx: PictureFormat = { ...f, geometryAdj: adjustXml(s.adjust) || undefined, fillXml: f.fillXml ?? (f.fill ? fillXml(f.fill) : undefined) };
  let xml = shapePropsXml(fx, cx, cy);
  if (custom) xml = xml.replace(/<a:prstGeom[\s\S]*?<\/a:prstGeom>/, custom);
  return xml;
}

/** Custom geometry in the definition syntax (geometry.ts) as a:custGeom. */
export function custGeomXml(def: string): string {
  const av: string[] = [], gd: string[] = [], paths: string[] = [];
  let rect = '<a:rect l="l" t="t" r="r" b="b"/>';
  for (const raw of def.split("\n")) {
    const line = raw.trim();
    if (!line) continue;
    const [kw, ...rest] = line.split(/\s+/);
    if (kw === "av") av.push(`<a:gd name="${rest[0]}" fmla="val ${rest[1]}"/>`);
    else if (kw === "gd") gd.push(`<a:gd name="${rest[0]}" fmla="${rest.slice(1).join(" ")}"/>`);
    else if (kw === "rect") rect = `<a:rect l="${rest[0]}" t="${rest[1]}" r="${rest[2]}" b="${rest[3]}"/>`;
    else if (kw === "path" || kw === "path:") {
      const colon = line.indexOf(":");
      const attrs = line.slice(4, colon).trim().split(/\s+/).filter(Boolean).map((a) => {
        const [k, v] = a.split("=");
        return k === "stroke" ? ` stroke="${v === "false" ? 0 : 1}"` : ` ${k}="${v}"`;
      }).join("");
      const tokens = line.slice(colon + 1).trim().split(/\s+/).filter(Boolean);
      const cmds: string[][] = [];
      for (const t of tokens) { if (/^[MLAQCZ]$/.test(t)) cmds.push([t]); else cmds[cmds.length - 1]?.push(t); }
      const pt = (x: string, y: string) => `<a:pt x="${x}" y="${y}"/>`;
      const body = cmds.map((c) => {
        const [k, ...v] = c;
        if (k === "M") return `<a:moveTo>${pt(v[0], v[1])}</a:moveTo>`;
        if (k === "L") return `<a:lnTo>${pt(v[0], v[1])}</a:lnTo>`;
        if (k === "C") return `<a:cubicBezTo>${pt(v[0], v[1])}${pt(v[2], v[3])}${pt(v[4], v[5])}</a:cubicBezTo>`;
        if (k === "Q") return `<a:quadBezTo>${pt(v[0], v[1])}${pt(v[2], v[3])}</a:quadBezTo>`;
        if (k === "A") return `<a:arcTo wR="${v[0]}" hR="${v[1]}" stAng="${v[2]}" swAng="${v[3]}"/>`;
        return "<a:close/>";
      }).join("");
      paths.push(`<a:path${attrs}>${body}</a:path>`);
    }
  }
  return `<a:custGeom><a:avLst>${av.join("")}</a:avLst><a:gdLst>${gd.join("")}</a:gdLst><a:ahLst/><a:cxnLst/>${rect}<a:pathLst>${paths.join("")}</a:pathLst></a:custGeom>`;
}

/**
 * A Formenart as Word applies it: the shape style (resolved against the theme) and the fill
 * and outline it sets on the shape itself (the theme styles set none).
 */
export function shapePreset(style: string, spPr: string, theme: Record<string, string>): { shape: Pick<ShapeInfo, "styleFill" | "styleLine" | "styleShadow" | "fontColor" | "styleXml">; pic: PictureFormat } {
  const doc = new DOMParser().parseFromString(`<r xmlns:wps="${NS_WPS}" xmlns:a="${NS_A}">${style}<wps:spPr>${spPr}</wps:spPr></r>`, "application/xml");
  const root = doc.documentElement;
  const styleEl = kids(root).find((x) => x.localName === "style");
  const resolved = resolveStyle(styleEl, theme, formatMatrix(null));
  const sp = kids(root).find((x) => x.localName === "spPr")!;
  const pic = readShapeProps(sp, theme, resolved.styleLine);
  return { shape: { ...resolved, styleXml: style || undefined }, pic };
}
