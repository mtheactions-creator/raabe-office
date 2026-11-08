// The Format pane's settings of a chart element in DrawingML (D-066, D-111): fills (solid, gradient,
// picture, pattern) and lines (c:spPr), effects (a:effectLst, a:scene3d, a:sp3d), text (c:txPr:
// fill, outline, effects, the text box), and the small value elements of axes, series, labels,
// trendlines and error bars. Written in the order the schema wants (Excel refuses parts out of
// order) and read back from the same XML.

import type {
  Arrow, AxisOptions, DataLabelOptions, Dash, Effects, ElementFormat, ErrorBars, Fill, Format3D, Gradient, Line, MarkerOptions, TextBox, TextFormat, Trendline,
} from "@raabe/charts";
import { attrs, children, esc, type El } from "./xml";

const EMU_PT = 12700;
const emu = (pt: number) => Math.round(pt * EMU_PT);
const pt = (e: string | undefined) => (e === undefined ? undefined : Math.round((Number(e) / EMU_PT) * 100) / 100);

/**
 * Pictures of Bild- oder Texturfüllungen met while writing a chart part: their data URLs in order;
 * the part refers to picture k as r:embed="@img{k}", which the drawing writer replaces by the
 * relationship it makes (setImageSink / takeImages).
 */
let images: string[] | null = null;
export function setImageSink(on: boolean): void { images = on ? [] : null; }
export function takeImages(): string[] { const out = images ?? []; images = null; return out; }

function srgb(color: string, transparency?: number): string {
  const hex = color.replace("#", "").slice(0, 6).toUpperCase();
  return transparency ? `<a:srgbClr val="${hex}"><a:alpha val="${Math.round((100 - transparency) * 1000)}"/></a:srgbClr>` : `<a:srgbClr val="${hex}"/>`;
}

function gradXml(g: Gradient): string {
  const stops = [...g.stops].sort((a, b) => a.pos - b.pos).map((s) => `<a:gs pos="${Math.round(s.pos * 1000)}">${srgb(s.color, s.transparency)}</a:gs>`).join("");
  const shade = g.kind === "linear"
    ? `<a:lin ang="${Math.round((((g.angle ?? 0) % 360) + 360) % 360 * 60000)}" scaled="1"/>`
    : `<a:path path="${g.kind === "radial" ? "circle" : g.kind === "rect" ? "rect" : "shape"}"><a:fillToRect l="50000" t="50000" r="50000" b="50000"/></a:path>`;
  return `<a:gradFill flip="none" rotWithShape="${g.rotWithShape === false ? 0 : 1}"><a:gsLst>${stops}</a:gsLst>${shade}<a:tileRect/></a:gradFill>`;
}

/** A fill element, or "" to keep the automatic one (`auto` is what Excel writes then). */
export function fillXml(f: Fill | undefined, auto = ""): string {
  if (!f || f.type === "auto") return auto;
  if (f.type === "none") return "<a:noFill/>";
  if (f.type === "gradient" && f.gradient?.stops.length) return gradXml(f.gradient);
  if (f.type === "pattern" && f.pattern) return `<a:pattFill prst="${f.pattern.preset}"><a:fgClr>${srgb(f.pattern.fg)}</a:fgClr><a:bgClr>${srgb(f.pattern.bg)}</a:bgClr></a:pattFill>`;
  if (f.type === "picture" && f.picture?.src) {
    const p = f.picture;
    let rid = "";
    if (images) { images.push(p.src); rid = `@img${images.length - 1}`; }
    if (!rid) return auto;
    const alpha = p.transparency ? `<a:alphaModFix amt="${Math.round((100 - p.transparency) * 1000)}"/>` : "";
    const o = p.offsets;
    const mode = p.tile ? '<a:tile tx="0" ty="0" sx="100000" sy="100000" flip="none" algn="tl"/>'
      : `<a:stretch><a:fillRect${o ? ` l="${Math.round(o.l * 1000)}" t="${Math.round(o.t * 1000)}" r="${Math.round(o.r * 1000)}" b="${Math.round(o.b * 1000)}"` : ""}/></a:stretch>`;
    return `<a:blipFill dpi="0" rotWithShape="1"><a:blip r:embed="${rid}">${alpha}</a:blip><a:srcRect/>${mode}</a:blipFill>`;
  }
  if (f.type !== "solid") return auto;
  return `<a:solidFill>${srgb(f.color ?? "#FFFFFF", f.transparency)}</a:solidFill>`;
}

function arrowXml(tag: "headEnd" | "tailEnd", a: Arrow | undefined): string {
  if (!a) return "";
  return `<a:${tag} type="${a.type}"${a.w ? ` w="${a.w}"` : ""}${a.len ? ` len="${a.len}"` : ""}/>`;
}

/** An a:ln element; `auto` is Excel's automatic line for the element. */
export function lineXml(l: Line | undefined, auto = ""): string {
  const extras = !!l && (l.compound !== undefined || l.cap !== undefined || l.join !== undefined || !!l.head || !!l.tail);
  if (!l || (l.type === "auto" && !l.width && !l.dash && !extras)) return auto;
  if (l.type === "none") return "<a:ln><a:noFill/></a:ln>";
  const w = l.width !== undefined ? ` w="${emu(l.width)}"` : "";
  let fill = "";
  if (l.type === "solid") fill = `<a:solidFill>${srgb(l.color ?? "#000000", l.transparency)}</a:solidFill>`;
  else if (l.type === "gradient" && l.gradient?.stops.length) fill = gradXml(l.gradient);
  else if (l.type === "auto") {
    // Automatic colour with own width or details: the automatic line's fill.
    fill = /<a:ln\b[^>]*>(<a:solidFill>[\s\S]*?<\/a:solidFill>)/.exec(auto)?.[1] ?? "";
  }
  const dash = l.dash ? `<a:prstDash val="${l.dash}"/>` : "";
  const join = l.join === "bevel" ? "<a:bevel/>" : l.join === "miter" ? '<a:miter lim="800000"/>' : "<a:round/>";
  return `<a:ln${w} cap="${l.cap ?? "flat"}" cmpd="${l.compound ?? "sng"}" algn="ctr">${fill}${dash}${join}${arrowXml("headEnd", l.head)}${arrowXml("tailEnd", l.tail)}</a:ln>`;
}

/** a:effectLst (Schatten, Leuchten, Weiche Kanten) in the schema's order. */
export function effectLstXml(e: Effects | undefined): string {
  if (!e || (!e.shadow && !e.glow && !e.softEdges)) return "<a:effectLst/>";
  let out = "";
  if (e.glow) out += `<a:glow rad="${emu(e.glow.size)}">${srgb(e.glow.color, e.glow.transparency ?? 60)}</a:glow>`;
  const s = e.shadow;
  const shadowColor = s ? srgb(s.color ?? "#000000", s.transparency ?? 60) : "";
  const dir = s ? Math.round((((s.angle ?? 45) % 360) + 360) % 360 * 60000) : 0;
  if (s?.inner) out += `<a:innerShdw blurRad="${emu(s.blur ?? 4)}" dist="${emu(s.distance ?? 3)}" dir="${dir}">${shadowColor}</a:innerShdw>`;
  if (s && !s.inner) {
    const sc = Math.round((s.size ?? 100) * 1000);
    out += `<a:outerShdw blurRad="${emu(s.blur ?? 4)}" dist="${emu(s.distance ?? 3)}" dir="${dir}"${sc !== 100000 ? ` sx="${sc}" sy="${sc}"` : ""} algn="tl" rotWithShape="0">${shadowColor}</a:outerShdw>`;
  }
  if (e.softEdges) out += `<a:softEdge rad="${emu(e.softEdges)}"/>`;
  return `<a:effectLst>${out}</a:effectLst>`;
}

/** a:scene3d and a:sp3d of a 3D-Format ("" without one). */
export function sp3dXml(f: Format3D | undefined): string {
  if (!f || (!f.top && !f.bottom && !f.depth && !f.contour && !f.material && !f.lighting)) return "";
  const bevel = (tag: string, b: Format3D["top"]) => (b ? `<a:${tag}${b.w !== undefined ? ` w="${emu(b.w)}"` : ""}${b.h !== undefined ? ` h="${emu(b.h)}"` : ""} prst="${b.type}"/>` : "");
  const scene = `<a:scene3d><a:camera prst="orthographicFront"/><a:lightRig rig="${f.lighting ?? "threePt"}" dir="t"/></a:scene3d>`;
  const attr = `${f.depth ? ` extrusionH="${emu(f.depth.size)}"` : ""}${f.contour ? ` contourW="${emu(f.contour.size)}"` : ""}${f.material ? ` prstMaterial="${f.material}"` : ""}`;
  const clr = (tag: string, c?: string) => (c ? `<a:${tag}>${srgb(c)}</a:${tag}>` : "");
  return `${scene}<a:sp3d${attr}>${bevel("bevelT", f.top)}${bevel("bevelB", f.bottom)}${clr("extrusionClr", f.depth?.color)}${clr("contourClr", f.contour?.color)}</a:sp3d>`;
}

/** c:spPr from a format; `autoFill` / `autoLine` are Excel's defaults for the element. */
export function spPrXml(fmt: { fill?: Fill; line?: Line; effects?: Effects } | undefined, autoFill: string, autoLine: string): string {
  return `<c:spPr>${fillXml(fmt?.fill, autoFill)}${lineXml(fmt?.line, autoLine)}${effectLstXml(fmt?.effects)}${sp3dXml(fmt?.effects?.format3d)}</c:spPr>`;
}

/** The run properties of a text format inside a:defRPr: attributes, and the children before a:latin (outline, fill, effects). */
export function defRPrAttrs(t: TextFormat | undefined, size: number): { attrs: string; fill: string | null; after?: string; font?: string } {
  const sz = t?.size !== undefined ? Math.round(t.size * 100) : size;
  let fill: string | null = t?.color ? `<a:solidFill>${srgb(t.color)}</a:solidFill>` : null;
  if (t?.fill && t.fill.type !== "auto") fill = fillXml(t.fill.type === "solid" ? { ...t.fill, color: t.fill.color ?? t.color } : t.fill, "") || null;
  const ln = t?.outline && t.outline.type !== "auto" ? lineXml(t.outline, "") : "";
  const fx = t?.effects && (t.effects.shadow || t.effects.glow || t.effects.softEdges) ? effectLstXml(t.effects) : "";
  return {
    attrs: ` sz="${sz}" b="${t?.bold ? 1 : 0}" i="${t?.italic ? 1 : 0}"`,
    fill: ln + (fill ?? "") || null,
    after: fx,
    font: t?.font,
  };
}

/** The a:bodyPr attributes of a text box (rotation, direction, anchor, insets). */
export function bodyPrAttrs(b: TextBox | undefined, rot: string): string {
  const r = b?.rotation !== undefined ? String(Math.round(b.rotation * 60000)) : rot;
  const ins = b?.margins ? ` lIns="${emu(b.margins.l)}" tIns="${emu(b.margins.t)}" rIns="${emu(b.margins.r)}" bIns="${emu(b.margins.b)}"` : "";
  return ` rot="${r}" spcFirstLastPara="1" vertOverflow="ellipsis" vert="${b?.dir ?? "horz"}" wrap="${b?.wrap === false ? "none" : "square"}"${ins} anchor="${b?.vAlign ?? "ctr"}" anchorCtr="1"`;
}

export function tickXml(tag: "majorTickMark" | "minorTickMark", v: AxisOptions["majorTick"]): string {
  return `<c:${tag} val="${v ?? "none"}"/>`;
}

/** c:scaling: Basis, orientation, Maximum, Minimum (schema order). */
export function scalingXml(o: AxisOptions | undefined): string {
  return `<c:scaling>${o?.logBase ? `<c:logBase val="${o.logBase}"/>` : ""}<c:orientation val="${o?.reverse ? "maxMin" : "minMax"}"/>${o?.max !== undefined ? `<c:max val="${o.max}"/>` : ""}${o?.min !== undefined ? `<c:min val="${o.min}"/>` : ""}</c:scaling>`;
}

/** Where this axis crosses the other: from the other axis' options (D-111: `crosses` is where the other axis crosses that one). */
export function crossesXml(o: AxisOptions | undefined): string {
  if (typeof o?.crosses === "number") return `<c:crossesAt val="${o.crosses}"/>`;
  return `<c:crosses val="${o?.crosses === "max" ? "max" : "autoZero"}"/>`;
}

/** c:dispUnits of a value axis ("" without Anzeigeeinheiten). */
export function dispUnitsXml(o: AxisOptions | undefined): string {
  const u = o?.dispUnits;
  if (u === undefined) return "";
  const unit = typeof u === "number" ? `<c:custUnit val="${u}"/>` : `<c:builtInUnit val="${u}"/>`;
  return `<c:dispUnits>${unit}${o?.dispUnitsLabel === false ? "" : '<c:dispUnitsLbl><c:layout/></c:dispUnitsLbl>'}</c:dispUnits>`;
}

/** The flags and position of data labels (c:dLbls without the shape and text parts). */
export function labelFlags(o: DataLabelOptions | undefined, defPos: string | undefined, leader: boolean): string {
  const pos = o?.pos ?? defPos;
  const b = (v: boolean | undefined, d: boolean) => ((v ?? d) ? 1 : 0);
  const lead = leader ? `<c:showLeaderLines val="${o?.leaderLines === false ? 0 : 1}"/>` : o?.leaderLines ? '<c:showLeaderLines val="1"/>' : "";
  return `${pos ? `<c:dLblPos val="${pos}"/>` : ""}<c:showLegendKey val="${b(o?.showLegendKey, false)}"/><c:showVal val="${b(o?.showVal, true)}"/><c:showCatName val="${b(o?.showCat, false)}"/><c:showSerName val="${b(o?.showSer, false)}"/><c:showPercent val="${b(o?.showPercent, false)}"/><c:showBubbleSize val="${b(o?.showBubble, false)}"/>${o?.separator !== undefined ? `<c:separator>${esc(o.separator)}</c:separator>` : ""}${lead}`;
}

/** Wert aus Zellen: the dLbls' extension (c15:showDataLabelsRange) and the series' range (c15:datalabelsRange). */
export function labelsRangeExt(o: DataLabelOptions | undefined): { dLbls: string; ser: string } {
  const r = o?.fromCells;
  if (!r) return { dLbls: "", ser: "" };
  const pts = r.cache.map((v, i) => (v === null ? "" : `<c:pt idx="${i}"><c:v>${esc(v)}</c:v></c:pt>`)).join("");
  return {
    dLbls: '<c:extLst><c:ext uri="{CE6537A1-D6FC-4f65-9D91-7224C49458BB}" xmlns:c15="http://schemas.microsoft.com/office/drawing/2012/chart"><c15:showDataLabelsRange val="1"/><c15:showLeaderLines val="1"/></c:ext></c:extLst>',
    ser: `<c:extLst><c:ext uri="{02D57815-91ED-43cb-92C2-25804820EDAC}" xmlns:c15="http://schemas.microsoft.com/office/drawing/2012/chart"><c15:datalabelsRange><c15:f>${esc(r.ref)}</c15:f><c15:dlblRangeCache><c:ptCount val="${r.cache.length}"/>${pts}</c15:dlblRangeCache></c15:datalabelsRange></c:ext></c:extLst>`,
  };
}

export function markerXml(mk: MarkerOptions | undefined, auto: string, autoFill: string, autoLine: string): string {
  if (!mk || (!mk.symbol && mk.size === undefined && !mk.fill && !mk.line)) return auto;
  const sym = mk.symbol && mk.symbol !== "auto" ? mk.symbol : "circle";
  if (sym === "none") return '<c:marker><c:symbol val="none"/></c:marker>';
  return `<c:marker><c:symbol val="${sym}"/><c:size val="${mk.size ?? 5}"/>${spPrXml(mk, autoFill, autoLine)}</c:marker>`;
}

export function trendlineXml(t: Trendline, autoLine: string, label: string, effects?: Effects): string {
  const type = t.type;
  const extra = (type === "poly" ? `<c:order val="${t.order ?? 2}"/>` : "") + (type === "movingAvg" ? `<c:period val="${t.period ?? 2}"/>` : "");
  const fc = (t.forward ? `<c:forward val="${t.forward}"/>` : "") + (t.backward ? `<c:backward val="${t.backward}"/>` : "") + (t.intercept !== undefined ? `<c:intercept val="${t.intercept}"/>` : "");
  return `<c:trendline>${t.name ? `<c:name>${esc(t.name)}</c:name>` : ""}<c:spPr>${lineXml(t.line, autoLine)}${effectLstXml(effects)}</c:spPr><c:trendlineType val="${type}"/>${extra}${fc}<c:dispRSqr val="${t.showR2 ? 1 : 0}"/><c:dispEq val="${t.showEq ? 1 : 0}"/>${label}</c:trendline>`;
}

export function errBarsXml(e: ErrorBars, scatter: boolean, autoLine: string): string {
  return `<c:errBars>${scatter ? '<c:errDir val="y"/>' : ""}<c:errBarType val="${e.dir}"/><c:errValType val="${e.type}"/><c:noEndCap val="${e.noEndCap ? 1 : 0}"/>${e.type === "fixedVal" || e.type === "percentage" || e.type === "stdDev" ? `<c:val val="${e.value ?? (e.type === "percentage" ? 5 : 1)}"/>` : ""}<c:spPr>${lineXml(e.line, autoLine)}<a:effectLst/></c:spPr></c:errBars>`;
}

// ----- reading ---------------------------------------------------------------------------------------

type Find = (el: El | undefined, ...path: string[]) => El | undefined;

export interface Reader {
  xml: string;
  find: Find;
  val: (el: El | undefined) => string | undefined;
  /** The picture of an r:embed relationship as a data URL (Bild- oder Texturfüllung). */
  image?: (rId: string) => string | undefined;
}

function colorOf(r: Reader, holder: El | undefined): { color?: string; transparency?: number } {
  const c = r.find(holder, "a:srgbClr");
  if (!c) return {};
  const hex = attrs(r.xml, c).val;
  const alpha = r.find(c, "a:alpha");
  const a = alpha ? Number(attrs(r.xml, alpha).val) : undefined;
  return { color: hex ? "#" + hex.toUpperCase() : undefined, ...(a !== undefined && a < 100000 ? { transparency: Math.round(100 - a / 1000) } : {}) };
}

const kidsOf = (r: Reader, el: El) => children(r.xml, el.open, el.close);

function readGradient(r: Reader, g: El): Gradient | undefined {
  const lst = r.find(g, "a:gsLst");
  if (!lst) return undefined;
  const stops = kidsOf(r, lst).filter((k) => k.name === "a:gs").map((gs) => {
    const c = colorOf(r, gs);
    return { pos: Math.round(Number(attrs(r.xml, gs).pos ?? 0) / 100) / 10, color: c.color ?? "#FFFFFF", ...(c.transparency ? { transparency: c.transparency } : {}) };
  });
  if (!stops.length) return undefined;
  const lin = r.find(g, "a:lin"), path = r.find(g, "a:path");
  const out: Gradient = { kind: "linear", stops };
  if (lin) out.angle = Math.round(Number(attrs(r.xml, lin).ang ?? 0) / 60000);
  else if (path) { const p = attrs(r.xml, path).path; out.kind = p === "circle" ? "radial" : p === "rect" ? "rect" : "path"; }
  if (attrs(r.xml, g).rotWithShape === "0") out.rotWithShape = false;
  return out;
}

/** A fill element (noFill, solidFill with an explicit colour, gradFill, pattFill, blipFill), else undefined. */
export function readFill(r: Reader, kids: El[]): Fill | undefined {
  const nf = kids.find((k) => k.name === "a:noFill"), sf = kids.find((k) => k.name === "a:solidFill");
  if (nf) return { type: "none" };
  if (sf && r.find(sf, "a:srgbClr")) return { type: "solid", ...colorOf(r, sf) };
  const gf = kids.find((k) => k.name === "a:gradFill");
  if (gf) { const g = readGradient(r, gf); if (g) return { type: "gradient", gradient: g }; }
  const pf = kids.find((k) => k.name === "a:pattFill");
  if (pf) {
    const fg = colorOf(r, r.find(pf, "a:fgClr")).color ?? "#000000", bg = colorOf(r, r.find(pf, "a:bgClr")).color ?? "#FFFFFF";
    return { type: "pattern", pattern: { preset: attrs(r.xml, pf).prst ?? "pct5", fg, bg } };
  }
  const bf = kids.find((k) => k.name === "a:blipFill");
  if (bf) {
    const blip = r.find(bf, "a:blip");
    const rid = blip ? attrs(r.xml, blip)["r:embed"] : undefined;
    const src = rid ? r.image?.(rid) : undefined;
    if (src) {
      const amt = r.find(blip, "a:alphaModFix");
      const a = amt ? Number(attrs(r.xml, amt).amt ?? 100000) : 100000;
      const pic: NonNullable<Fill["picture"]> = { src };
      if (r.find(bf, "a:tile")) pic.tile = true;
      if (a < 100000) pic.transparency = Math.round(100 - a / 1000);
      const fr = r.find(bf, "a:stretch", "a:fillRect");
      const fa = fr ? attrs(r.xml, fr) : {};
      if (fa.l || fa.t || fa.r || fa.b) pic.offsets = { l: Number(fa.l ?? 0) / 1000, t: Number(fa.t ?? 0) / 1000, r: Number(fa.r ?? 0) / 1000, b: Number(fa.b ?? 0) / 1000 };
      return { type: "picture", picture: pic };
    }
  }
  return undefined;
}

function readArrow(r: Reader, el: El | undefined): Arrow | undefined {
  if (!el) return undefined;
  const a = attrs(r.xml, el);
  if (!a.type || a.type === "none") return undefined;
  const out: Arrow = { type: a.type as Arrow["type"] };
  if (a.w) out.w = a.w as Arrow["w"];
  if (a.len) out.len = a.len as Arrow["len"];
  return out;
}

/** An a:ln as a Line, as far as it differs from an automatic one. */
export function readLine(r: Reader, ln: El): Line | undefined {
  const lk = kidsOf(r, ln);
  const a = attrs(r.xml, ln);
  const dash = r.val(lk.find((k) => k.name === "a:prstDash"));
  if (lk.some((k) => k.name === "a:noFill")) return { type: "none" };
  const lf = lk.find((k) => k.name === "a:solidFill");
  const explicit = lf && r.find(lf, "a:srgbClr");
  const gf = lk.find((k) => k.name === "a:gradFill");
  const grad = gf ? readGradient(r, gf) : undefined;
  const head = readArrow(r, lk.find((k) => k.name === "a:headEnd")), tail = readArrow(r, lk.find((k) => k.name === "a:tailEnd"));
  const join = lk.some((k) => k.name === "a:bevel") ? "bevel" : lk.some((k) => k.name === "a:miter") ? "miter" : undefined;
  const cmpd = a.cmpd && a.cmpd !== "sng" ? (a.cmpd as Line["compound"]) : undefined;
  const cap = a.cap === "rnd" || a.cap === "sq" ? (a.cap as Line["cap"]) : undefined;
  if (!explicit && !grad && !dash && !head && !tail && !join && !cmpd) return undefined;
  const out: Line = { type: grad ? "gradient" : explicit ? "solid" : "auto", ...(explicit ? colorOf(r, lf) : {}) };
  if (grad) out.gradient = grad;
  if (a.w) out.width = pt(a.w);
  if (dash && dash !== "solid") out.dash = dash as Dash;
  if (cmpd) out.compound = cmpd;
  if (cap) out.cap = cap;
  if (join) out.join = join;
  if (head) out.head = head;
  if (tail) out.tail = tail;
  return out;
}

/** Schatten, Leuchten, Weiche Kanten of an a:effectLst and 3D-Format of an a:sp3d (with its a:scene3d). */
export function readEffects(r: Reader, holder: El | undefined): Effects | undefined {
  if (!holder) return undefined;
  const out: Effects = {};
  const lst = r.find(holder, "a:effectLst");
  if (lst) {
    const glow = r.find(lst, "a:glow");
    if (glow) { const c = colorOf(r, glow); out.glow = { color: c.color ?? "#000000", size: pt(attrs(r.xml, glow).rad) ?? 0, ...(c.transparency !== undefined ? { transparency: c.transparency } : { transparency: 0 }) }; }
    const sh = r.find(lst, "a:outerShdw") ?? r.find(lst, "a:innerShdw");
    if (sh) {
      const a = attrs(r.xml, sh), c = colorOf(r, sh);
      out.shadow = { color: c.color ?? "#000000", transparency: c.transparency ?? 0, blur: pt(a.blurRad) ?? 0, distance: pt(a.dist) ?? 0, angle: Math.round(Number(a.dir ?? 0) / 60000) };
      if (sh.name === "a:innerShdw") out.shadow.inner = true;
      if (a.sx && Number(a.sx) !== 100000) out.shadow.size = Number(a.sx) / 1000;
    }
    const se = r.find(lst, "a:softEdge");
    if (se) out.softEdges = pt(attrs(r.xml, se).rad);
  }
  const sp3d = r.find(holder, "a:sp3d");
  if (sp3d) {
    const a = attrs(r.xml, sp3d);
    const f: Format3D = {};
    const bevel = (tag: string) => { const b = r.find(sp3d, tag); if (!b) return undefined; const ba = attrs(r.xml, b); return { type: ba.prst ?? "circle", w: pt(ba.w ?? "76200"), h: pt(ba.h ?? "76200") }; };
    const top = bevel("a:bevelT"), bottom = bevel("a:bevelB");
    if (top) f.top = top;
    if (bottom) f.bottom = bottom;
    if (a.extrusionH) f.depth = { size: pt(a.extrusionH)!, ...colorOf(r, r.find(sp3d, "a:extrusionClr")) };
    if (a.contourW) f.contour = { size: pt(a.contourW)!, ...colorOf(r, r.find(sp3d, "a:contourClr")) };
    if (a.prstMaterial) f.material = a.prstMaterial;
    const rig = r.find(holder, "a:scene3d", "a:lightRig");
    if (rig && attrs(r.xml, rig).rig && attrs(r.xml, rig).rig !== "threePt") f.lighting = attrs(r.xml, rig).rig;
    if (Object.keys(f).length) out.format3d = f;
  }
  return Object.keys(out).length ? out : undefined;
}

/** Fill, line and effects of a c:spPr, as far as they are explicit (scheme colours stay automatic). */
export function readShape(r: Reader, spPr: El | undefined): { fill?: Fill; line?: Line; effects?: Effects } {
  if (!spPr) return {};
  const out: { fill?: Fill; line?: Line; effects?: Effects } = {};
  const kids = kidsOf(r, spPr);
  const fill = readFill(r, kids);
  if (fill) out.fill = fill;
  const ln = kids.find((k) => k.name === "a:ln");
  if (ln) {
    const l = readLine(r, ln);
    if (l) out.line = l;
  }
  const fx = readEffects(r, spPr);
  if (fx) out.effects = fx;
  return out;
}

/** The text box of a c:txPr / c:rich bodyPr (rotation other than the automatic one, direction, anchor, insets). */
function readBox(r: Reader, bodyPr: El | undefined, autoRot: string): TextBox | undefined {
  if (!bodyPr) return undefined;
  const a = attrs(r.xml, bodyPr);
  const out: TextBox = {};
  if (a.rot !== undefined && a.rot !== autoRot && a.rot !== "-60000000") out.rotation = Math.round(Number(a.rot) / 60000);
  if (a.vert && a.vert !== "horz") out.dir = a.vert as TextBox["dir"];
  if (a.anchor && a.anchor !== "ctr") out.vAlign = a.anchor as TextBox["vAlign"];
  if (a.wrap === "none") out.wrap = false;
  if (a.lIns !== undefined && a.tIns !== undefined && a.rIns !== undefined && a.bIns !== undefined && !(a.lIns === "36576" && a.tIns === "18288" && a.rIns === "36576" && a.bIns === "18288")) {
    out.margins = { l: pt(a.lIns)!, t: pt(a.tIns)!, r: pt(a.rIns)!, b: pt(a.bIns)! };
  }
  return Object.keys(out).length ? out : undefined;
}

/** Size (pt), bold, italic, colour, fill, outline, effects and text box of a c:txPr's first a:defRPr. */
export function readText(r: Reader, txPr: El | undefined, defaultSize: number, autoRot = "0"): TextFormat | undefined {
  const box = readBox(r, r.find(txPr, "a:bodyPr"), autoRot);
  const d = r.find(txPr, "a:p", "a:pPr", "a:defRPr");
  const out: TextFormat = {};
  if (d) {
    const a = attrs(r.xml, d);
    if (a.sz && Number(a.sz) !== defaultSize) out.size = Number(a.sz) / 100;
    if (a.b === "1") out.bold = true;
    if (a.i === "1") out.italic = true;
    if (a.u && a.u !== "none") out.underline = true;
    if (a.strike && a.strike !== "noStrike") out.strike = true;
    const kids = kidsOf(r, d);
    const fill = readFill(r, kids);
    if (fill?.type === "solid") out.color = fill.color;
    if (fill && fill.type !== "solid") out.fill = fill;
    if (fill?.type === "solid" && fill.transparency) out.fill = fill;
    const ln = kids.find((k) => k.name === "a:ln");
    if (ln) { const l = readLine(r, ln); if (l) out.outline = l; }
    const fx = readEffects(r, d);
    if (fx) out.effects = fx;
    const latin = r.find(d, "a:latin");
    const face = latin ? attrs(r.xml, latin).typeface : undefined;
    if (face && !face.startsWith("+")) out.font = face;
  }
  if (box) out.box = box;
  return Object.keys(out).length ? out : undefined;
}

export function readFormat(r: Reader, spPr: El | undefined, txPr: El | undefined, size: number, autoRot = "0"): ElementFormat | undefined {
  const s = readShape(r, spPr);
  const t = readText(r, txPr, size, autoRot);
  const out: ElementFormat = { ...s, ...(t ? { text: t } : {}) };
  return Object.keys(out).length ? out : undefined;
}

/**
 * An axis' own options. `crosses` is filled in by the caller from the other axis' element (D-111):
 * the file says where an axis crosses the other one, the pane where the other one crosses it.
 */
export function readAxis(r: Reader, ax: El, isVal: boolean): AxisOptions | undefined {
  const o: AxisOptions = {};
  const sc = r.find(ax, "c:scaling");
  const num = (el: El | undefined) => { const v = r.val(el); return v === undefined ? undefined : Number(v); };
  if (r.val(r.find(sc, "c:orientation")) === "maxMin") o.reverse = true;
  const mx = num(r.find(sc, "c:max")), mn = num(r.find(sc, "c:min")), lb = num(r.find(sc, "c:logBase"));
  if (mx !== undefined) o.max = mx;
  if (mn !== undefined) o.min = mn;
  if (lb !== undefined) o.logBase = lb;
  if (r.val(r.find(ax, "c:delete")) === "1") o.visible = false;
  const nf = r.find(ax, "c:numFmt");
  if (nf && attrs(r.xml, nf).sourceLinked === "0") o.numFmt = attrs(r.xml, nf).formatCode;
  const mt = r.val(r.find(ax, "c:majorTickMark")), nt = r.val(r.find(ax, "c:minorTickMark"));
  if (mt && mt !== "none") o.majorTick = mt as AxisOptions["majorTick"];
  if (nt && nt !== "none") o.minorTick = nt as AxisOptions["minorTick"];
  const lp = r.val(r.find(ax, "c:tickLblPos"));
  if (lp && lp !== "nextTo") o.labelPos = lp as AxisOptions["labelPos"];
  if (isVal) {
    const mu = num(r.find(ax, "c:majorUnit")), nu = num(r.find(ax, "c:minorUnit"));
    if (mu !== undefined) o.majorUnit = mu;
    if (nu !== undefined) o.minorUnit = nu;
    const du = r.find(ax, "c:dispUnits");
    if (du) {
      const b = r.val(r.find(du, "c:builtInUnit")), cu = num(r.find(du, "c:custUnit"));
      if (b) o.dispUnits = b as AxisOptions["dispUnits"];
      else if (cu !== undefined) o.dispUnits = cu;
      if (!r.find(du, "c:dispUnitsLbl")) o.dispUnitsLabel = false;
    }
  } else {
    if (ax.name === "c:dateAx") {
      o.axisType = "date";
      const bu = r.val(r.find(ax, "c:baseTimeUnit")), mtu = r.val(r.find(ax, "c:majorTimeUnit")), ntu = r.val(r.find(ax, "c:minorTimeUnit"));
      if (bu) o.baseTimeUnit = bu as AxisOptions["baseTimeUnit"];
      if (mtu) o.majorTimeUnit = mtu as AxisOptions["majorTimeUnit"];
      if (ntu) o.minorTimeUnit = ntu as AxisOptions["minorTimeUnit"];
      const mu = num(r.find(ax, "c:majorUnit")), nu = num(r.find(ax, "c:minorUnit"));
      if (mu !== undefined) o.majorUnit = mu;
      if (nu !== undefined) o.minorUnit = nu;
    } else if (r.val(r.find(ax, "c:auto")) === "0") o.axisType = "text";
    const skipL = num(r.find(ax, "c:tickLblSkip")), skipM = num(r.find(ax, "c:tickMarkSkip")), off = num(r.find(ax, "c:lblOffset"));
    if (skipL !== undefined) o.tickLblSkip = skipL;
    if (skipM !== undefined) o.tickMarkSkip = skipM;
    if (off !== undefined && off !== 100) o.lblOffset = off;
  }
  const f = readFormat(r, r.find(ax, "c:spPr"), r.find(ax, "c:txPr"), 900, "-60000000");
  if (f?.line) o.line = f.line;
  if (f?.fill && f.fill.type !== "none") o.fill = f.fill;
  if (f?.effects) o.effects = f.effects;
  if (f?.text) o.text = f.text;
  return Object.keys(o).length ? o : undefined;
}

/** Where an axis element says it crosses the other axis (c:crossesAt / c:crosses max). */
export function readCrossing(r: Reader, ax: El): AxisOptions["crosses"] | undefined {
  const at = r.val(r.find(ax, "c:crossesAt"));
  if (at !== undefined) return Number(at);
  return r.val(r.find(ax, "c:crosses")) === "max" ? "max" : undefined;
}
