// Shapes and text boxes of a drawing (xdr:sp, xdr:cxnSp in two-cell anchors; D-066): read into the
// model, written as Excel 365 writes a new shape (preset geometry, the theme's shape style, a text
// body), and written back verbatim while their content is unchanged.

import type { SheetShape } from "@raabe/sheet-engine";
import { custGeomXml } from "../shapes";
import { custGeomFromXml } from "./custgeom";

const EMU_PX = 9525, EMU_PT = 12700;

const num = (re: RegExp, s: string) => { const m = re.exec(s); return m ? Number(m[1]) : undefined; };
const esc = (s: string) => s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
const unesc = (s: string) => s.replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, "&");

function point(ax: string, tag: string): SheetShape["from"] | undefined {
  const m = new RegExp(`<xdr:${tag}>([\\s\\S]*?)</xdr:${tag}>`).exec(ax);
  if (!m) return undefined;
  const v = (n: string) => Number(new RegExp(`<xdr:${n}>([^<]*)</xdr:${n}>`).exec(m[1])?.[1] ?? 0);
  return { col: v("col"), row: v("row"), x: v("colOff") / EMU_PX, y: v("rowOff") / EMU_PX };
}

/** What a shape holds (everything but its XML), to see whether it changed. */
export function shapeKey(s: SheetShape): string {
  const { raw: _r, rawKey: _k, ...rest } = s;
  void _r; void _k;
  return JSON.stringify(rest, Object.keys(rest).sort());
}

function srgbOf(block: string | undefined): { color?: string; transparency?: number } {
  if (!block) return {};
  const m = /<a:srgbClr val="([0-9A-Fa-f]{6})"(?:\s*\/>|>([\s\S]*?)<\/a:srgbClr>)/.exec(block);
  if (!m) return {};
  const alpha = m[2] ? num(/<a:alpha val="(\d+)"/, m[2]) : undefined;
  return { color: "#" + m[1].toUpperCase(), ...(alpha !== undefined && alpha < 100000 ? { transparency: Math.round(100 - alpha / 1000) } : {}) };
}

/** A shape anchor as a model (null: not a shape Excel's model covers: groups, pictures, charts). */
export function shapeFromXml(ax: string): SheetShape | null {
  if (!/^<xdr:twoCellAnchor\b/.test(ax) || !/<xdr:(sp|cxnSp|grpSp)\b/.test(ax)) return null;
  // A Datenschnitt's frame (its fallback is a plain shape): read by slicers.ts (D-092).
  if (/<(sle|tsle):(slicer|timeslicer)\b/.test(ax)) return null;
  const from = point(ax, "from"), to = point(ax, "to");
  if (!from || !to) return null;
  if (/^<xdr:twoCellAnchor\b[^>]*>\s*<xdr:from>[\s\S]*?<\/xdr:to>\s*<xdr:grpSp\b/.test(ax)) return groupFromXml(ax, from, to);
  const connector = /<xdr:cxnSp\b/.test(ax);
  const spPr = /<xdr:spPr>([\s\S]*?)<\/xdr:spPr>/.exec(ax)?.[1] ?? "";
  const s: SheetShape = { from, to, geom: /<a:prstGeom prst="(\w+)"/.exec(spPr)?.[1] ?? "rect" };
  const custom = custGeomFromXml(spPr);
  if (custom !== undefined) s.custom = custom;
  const av = /<a:prstGeom\b[^>]*>\s*<a:avLst>([\s\S]*?)<\/a:avLst>/.exec(spPr)?.[1];
  if (av) {
    const adj = Object.fromEntries([...av.matchAll(/<a:gd name="(\w+)" fmla="val (-?\d+)"\s*\/>/g)].map((m) => [m[1], Number(m[2])]));
    if (Object.keys(adj).length) s.adjust = adj;
  }
  if (connector) s.connector = true;
  const xfrm = /<a:xfrm\b([^>]*)>/.exec(spPr)?.[1] ?? "";
  const rot = num(/\brot="(-?\d+)"/, xfrm);
  if (rot) s.rot = rot / 60000;
  if (/flipH="1"/.test(xfrm)) s.flipH = true;
  if (/flipV="1"/.test(xfrm)) s.flipV = true;
  // Fill and line as far as they are explicit colours; scheme colours keep the theme's look.
  const lnBlock = /<a:ln\b[^>]*>([\s\S]*?)<\/a:ln>|<a:ln\b[^>]*\/>/.exec(spPr);
  const beforeLn = lnBlock ? spPr.slice(0, lnBlock.index) : spPr;
  if (/<a:noFill\/>/.test(beforeLn)) s.fill = { type: "none" };
  else { const c = srgbOf(/<a:solidFill>([\s\S]*?)<\/a:solidFill>/.exec(beforeLn)?.[1]); if (c.color) s.fill = { type: "solid", ...c }; }
  if (lnBlock) {
    const lb = lnBlock[1] ?? "";
    const w = num(/<a:ln\b[^>]*\bw="(\d+)"/, lnBlock[0]);
    if (/<a:noFill\/>/.test(lb)) s.line = { type: "none" };
    else {
      const c = srgbOf(/<a:solidFill>([\s\S]*?)<\/a:solidFill>/.exec(lb)?.[1]);
      const dash = /<a:prstDash val="(\w+)"/.exec(lb)?.[1];
      const head = /<a:headEnd type="(\w+)"/.exec(lb)?.[1], tail = /<a:tailEnd type="(\w+)"/.exec(lb)?.[1];
      if (c.color || w !== undefined || dash || head || tail) {
        s.line = { type: "solid", ...(c.color ? { color: c.color } : {}), ...(w !== undefined ? { width: Math.round((w / EMU_PT) * 100) / 100 } : {}),
          ...(dash && dash !== "solid" ? { dash } : {}), ...(head && head !== "none" ? { head } : {}), ...(tail && tail !== "none" ? { tail } : {}) };
      }
    }
  }
  if (/<xdr:cNvSpPr\b[^>]*txBox="1"/.test(ax)) s.textBox = true;
  const lnRef = /<a:lnRef idx="(\d+)">\s*<a:schemeClr val="accent1"(?:\s*\/>|>\s*(?:<a:shade val="(\d+)"\/>)?)/.exec(/<xdr:style>([\s\S]*?)<\/xdr:style>/.exec(ax)?.[1] ?? "");
  if (lnRef && !s.textBox) s.style = { ln: Number(lnRef[1]), ...(lnRef[2] ? { shade: Number(lnRef[2]) / 100000 } : {}) };
  const name = /<xdr:cNvPr\b[^>]*name="([^"]*)"/.exec(ax)?.[1];
  if (name) s.name = unesc(name);
  if (/<xdr:cNvPr\b[^>]*\bhidden="1"/.test(ax)) s.hidden = true;
  const body = /<xdr:txBody>([\s\S]*?)<\/xdr:txBody>/.exec(ax)?.[1];
  if (body) {
    const paras = body.split(/<a:p>|<a:p\s[^>]*>/).slice(1).map((p) => [...p.matchAll(/<a:t>([^<]*)<\/a:t>|<a:t\s[^>]*>([^<]*)<\/a:t>/g)].map((m) => unesc(m[1] ?? m[2] ?? "")).join(""));
    const text = paras.join("\n");
    if (text.trim()) s.text = text;
    const anchor = /<a:bodyPr\b[^>]*\banchor="(\w+)"/.exec(body)?.[1];
    if (anchor === "t" || anchor === "ctr" || anchor === "b") s.anchor = anchor;
    const algn = /<a:pPr\b[^>]*\balgn="(\w+)"/.exec(body)?.[1];
    if (algn === "l" || algn === "ctr" || algn === "r") s.align = algn;
    const rPr = /<a:rPr\b([^>]*)(?:\/>|>([\s\S]*?)<\/a:rPr>)/.exec(body);
    if (rPr) {
      const font: NonNullable<SheetShape["font"]> = {};
      const sz = num(/\bsz="(\d+)"/, rPr[1]);
      if (sz) font.size = sz / 100;
      if (/\bb="1"/.test(rPr[1])) font.bold = true;
      if (/\bi="1"/.test(rPr[1])) font.italic = true;
      if (/\bu="sng"/.test(rPr[1])) font.underline = true;
      const c = srgbOf(rPr[2]);
      if (c.color) font.color = c.color;
      if (Object.keys(font).length) s.font = font;
    }
  }
  s.raw = ax;
  s.rawKey = shapeKey(s);
  return s;
}

const pointXml = (tag: string, p: SheetShape["from"]) =>
  `<xdr:${tag}><xdr:col>${p.col}</xdr:col><xdr:colOff>${Math.round(p.x * EMU_PX)}</xdr:colOff><xdr:row>${p.row}</xdr:row><xdr:rowOff>${Math.round(p.y * EMU_PX)}</xdr:rowOff></xdr:${tag}>`;

function srgb(color: string, transparency?: number): string {
  const hex = color.replace("#", "").toUpperCase();
  return transparency ? `<a:srgbClr val="${hex}"><a:alpha val="${Math.round((100 - transparency) * 1000)}"/></a:srgbClr>` : `<a:srgbClr val="${hex}"/>`;
}

/** The shape's anchor as Excel writes a new one; `size` is its extent in px at 100 %. */
export function shapeToXml(s: SheetShape, id: number, size: { w: number; h: number }): string {
  if (s.raw && s.rawKey === shapeKey(s)) return s.raw;
  if (s.geom === "group") return groupToXml(s, id, size);
  const name = esc(s.name ?? (s.textBox ? `Textfeld ${id - 1}` : `Form ${id - 1}`));
  const xfrmAttrs = `${s.rot ? ` rot="${Math.round(s.rot * 60000)}"` : ""}${s.flipH ? ' flipH="1"' : ""}${s.flipV ? ' flipV="1"' : ""}`;
  const xfrm = `<a:xfrm${xfrmAttrs}><a:off x="0" y="0"/><a:ext cx="${Math.round(size.w * EMU_PX)}" cy="${Math.round(size.h * EMU_PX)}"/></a:xfrm>${s.custom !== undefined ? custGeomXml(s.custom) : `<a:prstGeom prst="${s.geom}"><a:avLst${s.adjust ? `>${Object.entries(s.adjust).map(([k, v]) => `<a:gd name="${k}" fmla="val ${Math.round(v)}"/>`).join("")}</a:avLst>` : "/>"}</a:prstGeom>`}`;
  const fill = s.fill ? (s.fill.type === "none" ? "<a:noFill/>" : `<a:solidFill>${srgb(s.fill.color ?? "#4472C4", s.fill.transparency)}</a:solidFill>`)
    : s.textBox ? '<a:solidFill><a:schemeClr val="lt1"/></a:solidFill>' : "";
  const ends = (s.line?.head ? `<a:headEnd type="${s.line.head}"/>` : "") + (s.line?.tail ? `<a:tailEnd type="${s.line.tail}"/>` : "");
  const line = s.line
    ? s.line.type === "none" ? "<a:ln><a:noFill/></a:ln>"
      : `<a:ln${s.line.width !== undefined ? ` w="${Math.round(s.line.width * EMU_PT)}"` : ""}>${s.line.color ? `<a:solidFill>${srgb(s.line.color)}</a:solidFill>` : ""}${s.line.dash ? `<a:prstDash val="${s.line.dash}"/>` : ""}${ends}</a:ln>`
    : s.textBox ? '<a:ln w="9525" cmpd="sng"><a:solidFill><a:schemeClr val="lt1"><a:shade val="50000"/></a:schemeClr></a:solidFill></a:ln>' : "";
  if (s.connector) {
    return `<xdr:twoCellAnchor>${pointXml("from", s.from)}${pointXml("to", s.to)}<xdr:cxnSp macro=""><xdr:nvCxnSpPr><xdr:cNvPr id="${id}" name="${name}"${s.hidden ? ' hidden="1"' : ""}/><xdr:cNvCxnSpPr/></xdr:nvCxnSpPr>`
      + `<xdr:spPr>${xfrm}${line}</xdr:spPr><xdr:style><a:lnRef idx="${s.style?.ln ?? 1}"><a:schemeClr val="accent1"/></a:lnRef><a:fillRef idx="0"><a:schemeClr val="accent1"/></a:fillRef><a:effectRef idx="${(s.style?.ln ?? 1) >= 2 ? 1 : 0}"><a:schemeClr val="accent1"/></a:effectRef><a:fontRef idx="minor"><a:schemeClr val="tx1"/></a:fontRef></xdr:style></xdr:cxnSp><xdr:clientData/></xdr:twoCellAnchor>`;
  }
  const f = s.font;
  const rAttrs = ` lang="de-DE"${f?.size ? ` sz="${Math.round(f.size * 100)}"` : ' sz="1100"'}${f?.bold ? ' b="1"' : ""}${f?.italic ? ' i="1"' : ""}${f?.underline ? ' u="sng"' : ""}`;
  const rFill = f?.color ? `<a:solidFill>${srgb(f.color)}</a:solidFill>` : "";
  // Excel 365 starts a new shape's text at the top left, as DrawingML's defaults are.
  const algn = s.align ?? "l";
  const paras = (s.text ?? "").split("\n").map((t) => t
    ? `<a:p><a:pPr algn="${algn}"/><a:r><a:rPr${rAttrs}>${rFill}</a:rPr><a:t>${esc(t)}</a:t></a:r></a:p>`
    : `<a:p><a:pPr algn="${algn}"/><a:endParaRPr${rAttrs}/></a:p>`).join("");
  const style = s.textBox
    ? '<xdr:style><a:lnRef idx="0"><a:scrgbClr r="0" g="0" b="0"/></a:lnRef><a:fillRef idx="0"><a:scrgbClr r="0" g="0" b="0"/></a:fillRef><a:effectRef idx="0"><a:scrgbClr r="0" g="0" b="0"/></a:effectRef><a:fontRef idx="minor"><a:schemeClr val="dk1"/></a:fontRef></xdr:style>'
    : `<xdr:style><a:lnRef idx="${s.style?.ln ?? 2}"><a:schemeClr val="accent1"><a:shade val="${Math.round((s.style?.shade ?? 0.15) * 100000)}"/></a:schemeClr></a:lnRef><a:fillRef idx="1"><a:schemeClr val="accent1"/></a:fillRef><a:effectRef idx="0"><a:schemeClr val="accent1"/></a:effectRef><a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef></xdr:style>`;
  return `<xdr:twoCellAnchor>${pointXml("from", s.from)}${pointXml("to", s.to)}<xdr:sp macro="" textlink=""><xdr:nvSpPr><xdr:cNvPr id="${id}" name="${name}"${s.hidden ? ' hidden="1"' : ""}/><xdr:cNvSpPr${s.textBox ? ' txBox="1"' : ""}/></xdr:nvSpPr>`
    + `<xdr:spPr>${xfrm}${fill}${line}</xdr:spPr>${style}<xdr:txBody><a:bodyPr vertOverflow="clip" horzOverflow="clip" wrap="square" rtlCol="0" anchor="${s.anchor ?? "t"}"/><a:lstStyle/>${paras}</xdr:txBody></xdr:sp><xdr:clientData/></xdr:twoCellAnchor>`;
}

/** The direct child elements (sp, cxnSp, grpSp, pic, graphicFrame) of a group's XML. */
function groupChildren(xml: string): string[] {
  const out: string[] = [];
  const re = /<(\/?)xdr:(sp|cxnSp|grpSp|pic|graphicFrame)\b[^>]*?(\/?)>/g;
  let depth = 0, start = -1;
  for (let m = re.exec(xml); m; m = re.exec(xml)) {
    if (m[3] === "/") { if (depth === 0) out.push(m[0]); continue; }
    if (m[1] !== "/") { if (depth === 0) start = m.index; depth++; }
    else { depth--; if (depth === 0 && start >= 0) { out.push(xml.slice(start, m.index + m[0].length)); start = -1; } }
  }
  return out;
}

/**
 * A group (xdr:grpSp) as a model: its children with their boxes in the group's own coordinates
 * (px at 100 %, chOff/chExt mapped onto the group's extent); null when it holds a picture or a chart.
 */
function groupFromXml(ax: string, from: SheetShape["from"], to: SheetShape["from"]): SheetShape | null {
  const inner = /<xdr:grpSp>([\s\S]*)<\/xdr:grpSp>/.exec(ax)?.[1];
  if (inner === undefined) return null;
  const grpSpPr = /<xdr:grpSpPr>([\s\S]*?)<\/xdr:grpSpPr>/.exec(inner)?.[1] ?? "";
  const n = (tag: string, a: string) => Number(new RegExp(`<a:${tag}\\b[^>]*\\b${a}="(-?\\d+)"`).exec(grpSpPr)?.[1] ?? 0);
  const ext = { w: n("ext", "cx"), h: n("ext", "cy") }, chOff = { x: n("chOff", "x"), y: n("chOff", "y") }, chExt = { w: n("chExt", "cx") || ext.w || 1, h: n("chExt", "cy") || ext.h || 1 };
  const body = inner.replace(/<xdr:nvGrpSpPr>[\s\S]*?<\/xdr:nvGrpSpPr>/, "").replace(/<xdr:grpSpPr>[\s\S]*?<\/xdr:grpSpPr>/, "");
  const children: SheetShape[] = [];
  const kx = ext.w / chExt.w || 1, ky = ext.h / chExt.h || 1;
  for (const c of groupChildren(body)) {
    if (/^<xdr:(pic|graphicFrame)\b/.test(c)) return null;
    const wrapped = `<xdr:twoCellAnchor>${pointXml("from", from)}${pointXml("to", to)}${c}<xdr:clientData/></xdr:twoCellAnchor>`;
    const child = shapeFromXml(wrapped);
    if (!child) return null;
    const x = /<a:xfrm\b[^>]*>\s*<a:off x="(-?\d+)" y="(-?\d+)"\/>\s*<a:ext cx="(\d+)" cy="(\d+)"\/>/.exec(c);
    if (x) child.rel = { x: ((Number(x[1]) - chOff.x) * kx) / EMU_PX, y: ((Number(x[2]) - chOff.y) * ky) / EMU_PX, w: (Number(x[3]) * kx) / EMU_PX, h: (Number(x[4]) * ky) / EMU_PX };
    delete child.raw;
    delete child.rawKey;
    children.push(child);
  }
  const s: SheetShape = { from, to, geom: "group", children, frame: { w: ext.w / EMU_PX, h: ext.h / EMU_PX } };
  const name = /<xdr:nvGrpSpPr>[\s\S]*?<xdr:cNvPr\b[^>]*name="([^"]*)"/.exec(inner)?.[1];
  if (name) s.name = unesc(name);
  const rot = /<xdr:grpSpPr>\s*<a:xfrm\b[^>]*\brot="(-?\d+)"/.exec(inner)?.[1];
  if (rot) s.rot = Number(rot) / 60000;
  s.raw = ax;
  s.rawKey = shapeKey(s);
  return s;
}

/** A group's anchor: its children placed in the group's coordinates (chOff 0, chExt = its frame). */
function groupToXml(s: SheetShape, id: number, size: { w: number; h: number }): string {
  const frame = s.frame ?? size;
  const kids = (s.children ?? []).map((c, k) => {
    const r = c.rel ?? { x: 0, y: 0, w: frame.w, h: frame.h };
    const xml = shapeToXml({ ...c, raw: undefined }, id * 100 + k + 1, { w: r.w, h: r.h });
    const el = /<xdr:(sp|cxnSp|grpSp)\b[\s\S]*<\/xdr:\1>/.exec(xml)?.[0] ?? "";
    return el.replace(/<a:off x="0" y="0"\/>/, `<a:off x="${Math.round(r.x * EMU_PX)}" y="${Math.round(r.y * EMU_PX)}"/>`);
  }).join("");
  const W = Math.round(size.w * EMU_PX), H = Math.round(size.h * EMU_PX), FW = Math.round(frame.w * EMU_PX), FH = Math.round(frame.h * EMU_PX);
  return `<xdr:twoCellAnchor>${pointXml("from", s.from)}${pointXml("to", s.to)}<xdr:grpSp><xdr:nvGrpSpPr><xdr:cNvPr id="${id}" name="${esc(s.name ?? `Gruppieren ${id - 1}`)}"/><xdr:cNvGrpSpPr/></xdr:nvGrpSpPr>`
    + `<xdr:grpSpPr><a:xfrm${s.rot ? ` rot="${Math.round(s.rot * 60000)}"` : ""}><a:off x="0" y="0"/><a:ext cx="${W}" cy="${H}"/><a:chOff x="0" y="0"/><a:chExt cx="${FW}" cy="${FH}"/></a:xfrm></xdr:grpSpPr>`
    + `${kids}</xdr:grpSp><xdr:clientData/></xdr:twoCellAnchor>`;
}
