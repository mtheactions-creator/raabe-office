// xl/styles.xml: Excel's cell formats (cellXfs) become the model's whole CellStyle objects.
// Each style read keeps the index of its xf, so a format the user did not change is written
// back as the same xf, and a changed one reuses the file's font, fill, border and colour
// elements wherever they still match (theme colours stay theme colours).

import { BUILTIN_FORMATS, StyleTable, builtinId, type CondRule, type BorderSide, type Borders, type CellStyle, type Font, type HAlign, type VAlign } from "@raabe/sheet-engine";
import { colorXml, readIndexed, resolveColor, type ColorContext } from "./colors";
export type { ColorContext };
import { attrs, child, children, esc, outer, parseAttrs, place, setAttrs, split, joinSplit, type El, type Split, XML_HEAD } from "./xml";

/** Built-in number formats as German Excel reads them (the number-format engine's table). */
export const BUILTIN_NUMFMTS: Readonly<Record<number, string>> = BUILTIN_FORMATS;

/** Ids Excel writes together with their code (locale-dependent currency and accounting). */
const WITH_CODE = new Set([5, 6, 7, 8, 41, 42, 43, 44]);

/** CT_Stylesheet's order of children. */
const STYLESHEET_ORDER = ["numFmts", "fonts", "fills", "borders", "cellStyleXfs", "cellXfs", "cellStyles", "dxfs", "tableStyles", "colors", "extLst"];

interface Fill { fill?: string; fillXml?: string }

/** What the writer needs of the file's styles.xml. */
export interface StylesRaw {
  xml: string;
  numFmts: [number, string][];
  fonts: string[];
  fills: string[];
  borders: string[];
  xfs: string[];
  /** Formats of conditional formatting (<dxf>). */
  dxfs: string[];
  /** cellStyleXfs index → its cell style name; name → cellStyleXfs index. */
  styleNames: [number, string][];
  ctx: ColorContext;
}

const on = (a: Record<string, string> | undefined): boolean | undefined =>
  a === undefined ? undefined : !(a.val === "0" || a.val === "false");

function readFont(xml: string, el: El, ctx: ColorContext): Font {
  const f: Font = { name: "Calibri", size: 11 };
  for (const c of children(xml, el.open, el.close)) {
    const a = attrs(xml, c);
    switch (c.name) {
      case "name": case "rFont": f.name = a.val ?? f.name; break;
      case "sz": f.size = Number(a.val ?? 11); break;
      case "b": if (on(a)) f.bold = true; break;
      case "i": if (on(a)) f.italic = true; break;
      case "strike": if (on(a)) f.strike = true; break;
      case "u": if ((a.val ?? "single") !== "none") f.underline = (a.val ?? "single") as Font["underline"]; break;
      case "vertAlign": if (a.val === "superscript" || a.val === "subscript") f.vertAlign = a.val; break;
      case "color": { const hex = resolveColor(a, ctx); if (hex) f.color = hex; break; }
      case "scheme": if (a.val && a.val !== "none") f.scheme = a.val; break;
    }
  }
  return f;
}

function readFill(xml: string, el: El, ctx: ColorContext): Fill {
  const pf = child(xml, el, "patternFill");
  if (pf) {
    const type = attrs(xml, pf).patternType ?? "none";
    if (type === "none") return {};
    if (type === "solid") {
      const fg = child(xml, pf, "fgColor");
      return { fill: (fg && resolveColor(attrs(xml, fg), ctx)) ?? "#000000" };
    }
  }
  return { fillXml: outer(xml, el) };
}

const SIDES = ["left", "right", "top", "bottom", "diagonal"] as const;

function readBorder(xml: string, el: El, ctx: ColorContext): Borders {
  const b: Borders = {};
  const a = attrs(xml, el);
  for (const c of children(xml, el.open, el.close)) {
    const name = c.name === "start" ? "left" : c.name === "end" ? "right" : c.name;
    if (!(SIDES as readonly string[]).includes(name)) continue;
    const style = attrs(xml, c).style;
    if (!style || style === "none") continue;
    const col = child(xml, c, "color");
    const side: BorderSide = { style };
    const hex = col ? resolveColor(attrs(xml, col), ctx) : undefined;
    if (hex) side.color = hex;
    b[name as (typeof SIDES)[number]] = side;
  }
  if (a.diagonalUp === "1" || a.diagonalUp === "true") b.diagonalUp = true;
  if (a.diagonalDown === "1" || a.diagonalDown === "true") b.diagonalDown = true;
  return b;
}

/** The parts of styles.xml as element lists. */
function sections(xml: string) {
  const s = split(xml);
  const r = { s, el: {} as Record<string, El> };
  // Re-find the elements for their positions.
  const rootEl = { name: "", start: 0, end: xml.length, open: s.head.length, close: xml.length - s.tail.length };
  for (const c of children(xml, rootEl.open, rootEl.close)) r.el[c.name] = c;
  return r;
}

export interface ReadStyles {
  table: StyleTable;
  raw: StylesRaw;
}

/** The model's style table from styles.xml; style ids equal the file's xf indexes. */
export function readStyles(xml: string, theme: Record<string, string>): ReadStyles {
  const { el } = sections(xml);
  const kids = (name: string) => (el[name] ? children(xml, el[name].open, el[name].close) : []);
  const ctx: ColorContext = { theme, indexed: readIndexed(xml, el.colors) };
  const numFmts = new Map<number, string>();
  for (const n of kids("numFmts")) {
    const a = attrs(xml, n);
    numFmts.set(Number(a.numFmtId), a.formatCode ?? "General");
  }
  const fontEls = kids("fonts"), fillEls = kids("fills"), borderEls = kids("borders");
  const fonts = fontEls.map((f) => readFont(xml, f, ctx));
  const fills = fillEls.map((f) => readFill(xml, f, ctx));
  const borders = borderEls.map((b) => readBorder(xml, b, ctx));
  const styleNames = new Map<number, string>();
  for (const cs of kids("cellStyles")) {
    const a = attrs(xml, cs);
    if (a.xfId !== undefined && !styleNames.has(Number(a.xfId))) styleNames.set(Number(a.xfId), a.name ?? "");
  }
  const xfEls = kids("cellXfs");
  const list = xfEls.map((x, i) => xfStyle(xml, x, i, { numFmts, fonts, fills, borders, styleNames }));
  const table = new StyleTable(list[0] ?? xfStyle(xml, undefined, 0, { numFmts, fonts, fills, borders, styleNames }));
  for (const s of list.slice(1)) table.id(s);
  return {
    table,
    raw: {
      xml, numFmts: [...numFmts], ctx, styleNames: [...styleNames],
      fonts: fontEls.map((e) => outer(xml, e)), fills: fillEls.map((e) => outer(xml, e)), borders: borderEls.map((e) => outer(xml, e)),
      xfs: xfEls.map((e) => outer(xml, e)),
      dxfs: kids("dxfs").map((e) => outer(xml, e)),
    },
  };
}

/** A conditional format's look (CondRule.format) of a <dxf>. */
export type DxfFormat = NonNullable<CondRule["format"]>;

export function readDxf(xml: string, ctx: ColorContext): DxfFormat {
  const out: DxfFormat = {};
  const r = { name: "", start: 0, end: xml.length, open: xml.indexOf(">") + 1, close: xml.lastIndexOf("</") };
  for (const c of children(xml, r.open, r.close)) {
    if (c.name === "font") {
      const f = readFont(xml, c, ctx);
      if (f.color) out.color = f.color;
      if (f.bold) out.bold = true;
      if (f.italic) out.italic = true;
      if (f.underline) out.underline = true;
      if (f.strike) out.strike = true;
    } else if (c.name === "fill") {
      // In a dxf a solid fill's colour is its bgColor.
      const pf = child(xml, c, "patternFill");
      const bg = pf && child(xml, pf, "bgColor");
      const fg = pf && child(xml, pf, "fgColor");
      const hex = (bg && resolveColor(attrs(xml, bg), ctx)) ?? (fg && resolveColor(attrs(xml, fg), ctx));
      if (hex) out.fill = hex;
    } else if (c.name === "border") {
      const side = children(xml, c.open, c.close).map((k) => attrs(xml, k).style).find((st) => st && st !== "none");
      if (side) out.border = side;
    }
  }
  return out;
}

/** A <dxf> for a format, as Excel writes them. */
export function dxfXml(f: DxfFormat): string {
  const font = f.color || f.bold || f.italic || f.underline || f.strike
    ? `<font>${f.bold ? "<b/>" : ""}${f.italic ? "<i/>" : ""}${f.strike ? "<strike/>" : ""}${f.underline ? "<u/>" : ""}${f.color ? colorXml("color", f.color) : ""}</font>` : "";
  const fill = f.fill ? `<fill><patternFill>${colorXml("bgColor", f.fill)}</patternFill></fill>` : "";
  const side = (n: string) => (f.border ? `<${n} style="${esc(f.border)}"><color auto="1"/></${n}>` : "");
  const border = f.border ? `<border>${side("left")}${side("right")}${side("top")}${side("bottom")}</border>` : "";
  return `<dxf>${font}${fill}${border}</dxf>`;
}

interface Parts {
  numFmts: Map<number, string>;
  fonts: Font[];
  fills: Fill[];
  borders: Borders[];
  styleNames: Map<number, string>;
}

function xfStyle(xml: string, x: El | undefined, index: number, p: Parts): CellStyle {
  const a = x ? attrs(xml, x) : {};
  const n = (k: string) => Number(a[k] ?? 0);
  const numFmtId = n("numFmtId");
  const s: CellStyle = {
    font: structuredClone(p.fonts[n("fontId")] ?? { name: "Calibri", size: 11, scheme: "minor" }),
    borders: structuredClone(p.borders[n("borderId")] ?? {}),
    numFmt: p.numFmts.get(numFmtId) ?? BUILTIN_NUMFMTS[numFmtId] ?? "General",
  };
  const fill = p.fills[n("fillId")] ?? {};
  if (fill.fill) s.fill = fill.fill;
  if (fill.fillXml) s.fillXml = fill.fillXml;
  if (x) {
    const al = child(xml, x, "alignment");
    if (al) {
      const b = attrs(xml, al);
      if (b.horizontal && b.horizontal !== "general") s.h = b.horizontal as HAlign;
      if (b.vertical && b.vertical !== "bottom") s.v = b.vertical as VAlign;
      if (b.wrapText === "1" || b.wrapText === "true") s.wrap = true;
      if (b.indent && b.indent !== "0") s.indent = Number(b.indent);
      if (b.textRotation && b.textRotation !== "0") s.rotation = Number(b.textRotation);
      if (b.shrinkToFit === "1" || b.shrinkToFit === "true") s.shrink = true;
    }
    const pr = child(xml, x, "protection");
    if (pr) {
      const b = attrs(xml, pr);
      if (b.locked === "0" || b.locked === "false") s.locked = false;
      if (b.hidden === "1" || b.hidden === "true") s.hidden = true;
    }
  }
  const xfId = n("xfId");
  if (xfId) {
    const name = p.styleNames.get(xfId);
    if (name) s.styleName = name;
  }
  s.xf = index;
  return s;
}

/** Canonical JSON (sorted keys) for comparing formats. */
export function canon(v: unknown): string {
  if (v === null || typeof v !== "object") return JSON.stringify(v) ?? "null";
  if (Array.isArray(v)) return `[${v.map(canon).join(",")}]`;
  const o = v as Record<string, unknown>;
  return `{${Object.keys(o).filter((k) => o[k] !== undefined).sort().map((k) => `${JSON.stringify(k)}:${canon(o[k])}`).join(",")}}`;
}

const FAMILY: Record<string, number> = {
  calibri: 2, "calibri light": 2, arial: 2, verdana: 2, tahoma: 2, "segoe ui": 2, aptos: 2, "aptos narrow": 2, "arial narrow": 2,
  "times new roman": 1, cambria: 1, georgia: 1, garamond: 1, "courier new": 3, consolas: 3, "comic sans ms": 4,
};

/**
 * Builds styles.xml for a workbook: the file's own lists, with formats the model added or
 * changed appended. Gives the xf index for each model style id.
 */
export class StylesWriter {
  private numFmts: Map<number, string>;
  private fonts: string[];
  private fills: string[];
  private borders: string[];
  private xfs: string[];
  private dxfs: string[];
  private readonly orig: { fonts: number; fills: number; borders: number; xfs: number; numFmts: number; dxfs: number };
  private readonly base: ReadStyles;
  private readonly styleXfByName = new Map<string, number>();
  private readonly cache = new Map<number, number>();
  private readonly xfIndex = new Map<string, number>();
  private readonly fontModels: Font[];
  private readonly fillModels: Fill[];
  private readonly borderModels: Borders[];

  constructor(raw: StylesRaw, private table: StyleTable, private themeFonts: { major: string; minor: string }) {
    this.base = readStyles(raw.xml, raw.ctx.theme);
    this.numFmts = new Map(raw.numFmts);
    this.fonts = [...raw.fonts];
    this.fills = [...raw.fills];
    this.borders = [...raw.borders];
    this.xfs = [...raw.xfs];
    this.dxfs = [...raw.dxfs];
    this.orig = { fonts: this.fonts.length, fills: this.fills.length, borders: this.borders.length, xfs: this.xfs.length, numFmts: this.numFmts.size, dxfs: this.dxfs.length };
    for (const [i, name] of raw.styleNames) if (!this.styleXfByName.has(name)) this.styleXfByName.set(name, i);
    const ctx = raw.ctx;
    const { el } = sections(raw.xml);
    const kids = (name: string) => (el[name] ? children(raw.xml, el[name].open, el[name].close) : []);
    this.fontModels = kids("fonts").map((f) => readFont(raw.xml, f, ctx));
    this.fillModels = kids("fills").map((f) => readFill(raw.xml, f, ctx));
    this.borderModels = kids("borders").map((b) => readBorder(raw.xml, b, ctx));
    this.xfs.forEach((x, i) => { if (!this.xfIndex.has(x)) this.xfIndex.set(x, i); });
  }

  /** The dxf index for a conditional format's look: the file's own when it still matches. */
  dxf(f: DxfFormat, old?: number): number {
    if (old !== undefined && old < this.orig.dxfs && canon(readDxf(this.dxfs[old], this.base.raw.ctx)) === canon(f)) return old;
    const xml = dxfXml(f);
    const at = this.dxfs.indexOf(xml);
    if (at >= 0) return at;
    this.dxfs.push(xml);
    return this.dxfs.length - 1;
  }

  /** The xf index for a model style id. */
  xf(id: number | undefined): number {
    const key = id ?? 0;
    let hit = this.cache.get(key);
    if (hit === undefined) {
      hit = this.xfFor(this.table.get(key));
      this.cache.set(key, hit);
    }
    return hit;
  }

  private xfFor(style: CellStyle): number {
    const read = style.xf !== undefined ? this.base.table.list[style.xf] : undefined;
    if (read && read.xf === style.xf && canon(read) === canon(style) && style.xf! < this.orig.xfs) return style.xf!;
    const baseIndex = read ? style.xf! : 0;
    const baseXml = this.xfs[baseIndex] ?? '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>';
    const b = parseAttrs(baseXml.slice(0, baseXml.indexOf(">") + 1));
    const bn = (k: string) => Number(b[k] ?? 0);
    const fontId = canon(this.fontModels[bn("fontId")]) === canon(style.font) ? bn("fontId") : this.font(style.font, this.fontModels[bn("fontId")] ? bn("fontId") : 0);
    const fillId = this.fill(style, bn("fillId"));
    const borderId = canon(this.borderModels[bn("borderId")] ?? {}) === canon(style.borders) ? bn("borderId") : this.border(style.borders, bn("borderId"));
    const numFmtId = this.numFmt(style.numFmt, bn("numFmtId"));
    let xfId = 0;
    if (style.styleName) xfId = this.styleXfByName.get(style.styleName) ?? 0;
    const al = alignmentXml(style, baseXml);
    const pr = style.locked === false || style.hidden ? `<protection${style.locked === false ? ' locked="0"' : ""}${style.hidden ? ' hidden="1"' : ""}/>` : "";
    const keep = Object.entries(b).filter(([k]) => !/^(numFmtId|fontId|fillId|borderId|xfId|apply\w+)$/.test(k)).map(([k, v]) => ` ${k}="${esc(v)}"`).join("");
    const flag = (name: string, set: boolean) => (set ? ` ${name}="1"` : "");
    const xml = `<xf numFmtId="${numFmtId}" fontId="${fontId}" fillId="${fillId}" borderId="${borderId}" xfId="${xfId}"${keep}` +
      flag("applyNumberFormat", numFmtId !== 0) + flag("applyFont", fontId !== 0) + flag("applyFill", fillId !== 0) + flag("applyBorder", borderId !== 0) +
      flag("applyAlignment", al !== "") + flag("applyProtection", pr !== "") + (al || pr ? `>${al}${pr}</xf>` : "/>");
    let i = this.xfIndex.get(xml);
    if (i === undefined) {
      i = this.xfs.length;
      this.xfs.push(xml);
      this.xfIndex.set(xml, i);
    }
    return i;
  }

  private add(list: string[], models: unknown[], xml: string, model: unknown): number {
    const at = list.indexOf(xml);
    if (at >= 0) return at;
    list.push(xml);
    models.push(model);
    return list.length - 1;
  }

  private font(f: Font, baseId: number): number {
    const base = this.fontModels[baseId];
    const baseXml = this.fonts[baseId] ?? "";
    let color = "";
    if (f.color) {
      const old = /<color\b[^>]*\/>/.exec(baseXml)?.[0];
      color = old && base?.color === f.color ? old : colorXml("color", f.color);
    }
    const sameName = base && base.name === f.name;
    const family = sameName ? /<family\b[^>]*\/>/.exec(baseXml)?.[0] ?? "" : FAMILY[f.name.toLowerCase()] ? `<family val="${FAMILY[f.name.toLowerCase()]}"/>` : "";
    const charset = sameName ? /<charset\b[^>]*\/>/.exec(baseXml)?.[0] ?? "" : "";
    // A scheme font follows the theme: only when the name is the theme's font.
    const scheme = f.scheme && f.name === (f.scheme === "major" ? this.themeFonts.major : this.themeFonts.minor) ? `<scheme val="${esc(f.scheme)}"/>` : "";
    const xml = `<font>${f.bold ? "<b/>" : ""}${f.italic ? "<i/>" : ""}${f.strike ? "<strike/>" : ""}` +
      `${f.underline ? (f.underline === "single" ? "<u/>" : `<u val="${f.underline}"/>`) : ""}${f.vertAlign ? `<vertAlign val="${f.vertAlign}"/>` : ""}` +
      `<sz val="${f.size}"/>${color}<name val="${esc(f.name)}"/>${family}${charset}${scheme}</font>`;
    return this.add(this.fonts, this.fontModels, xml, f);
  }

  private fill(s: CellStyle, baseId: number): number {
    const want: Fill = s.fillXml ? { fillXml: s.fillXml } : s.fill ? { fill: s.fill } : {};
    if (canon(this.fillModels[baseId] ?? {}) === canon(want)) return baseId;
    if (!want.fill && !want.fillXml) return 0;
    const same = this.fillModels.findIndex((m) => canon(m) === canon(want));
    if (same >= 0) return same;
    const xml = want.fillXml ?? `<fill><patternFill patternType="solid">${colorXml("fgColor", want.fill)}<bgColor indexed="64"/></patternFill></fill>`;
    return this.add(this.fills, this.fillModels, xml, want);
  }

  private border(bd: Borders, baseId: number): number {
    const base = this.borderModels[baseId] ?? {};
    const baseXml = this.borders[baseId] ?? "";
    const side = (name: (typeof SIDES)[number]) => {
      const s = bd[name];
      if (!s) return `<${name}/>`;
      const old = new RegExp(`<${name}\\b[^>]*>(<color\\b[^>]*/>)</${name}>`).exec(baseXml)?.[1];
      const color = old && base[name]?.color === s.color ? old : s.color ? colorXml("color", s.color) : '<color indexed="64"/>';
      return `<${name} style="${esc(s.style)}">${color}</${name}>`;
    };
    const xml = `<border${bd.diagonalUp ? ' diagonalUp="1"' : ""}${bd.diagonalDown ? ' diagonalDown="1"' : ""}>${SIDES.map(side).join("")}</border>`;
    return this.add(this.borders, this.borderModels, xml, bd);
  }

  private numFmt(code: string, baseId: number): number {
    const baseCode = this.numFmts.get(baseId) ?? BUILTIN_NUMFMTS[baseId];
    if (baseCode === code) return baseId;
    for (const [id, c] of this.numFmts) if (c === code) return id;
    const builtin = builtinId(code);
    if (builtin !== undefined) {
      if (WITH_CODE.has(builtin)) this.numFmts.set(builtin, code);
      return builtin;
    }
    let id = 164;
    for (const k of this.numFmts.keys()) if (k >= id) id = k + 1;
    this.numFmts.set(id, code);
    return id;
  }

  /** styles.xml with the appended formats; the file's own text when nothing was added. */
  xml(): string {
    const o = this.orig;
    if (this.fonts.length === o.fonts && this.fills.length === o.fills && this.borders.length === o.borders && this.xfs.length === o.xfs && this.numFmts.size === o.numFmts && this.dxfs.length === o.dxfs) return this.base.raw.xml;
    const s: Split = split(this.base.raw.xml);
    const list = (name: string, items: string[], extra?: (tag: string) => string) => {
      const old = s.parts.find((p) => p.name === name)?.xml;
      let tag = old ? old.slice(0, old.indexOf(">") + 1) : `<${name}>`;
      if (tag.endsWith("/>")) tag = tag.slice(0, -2).replace(/\s+$/, "") + ">";
      tag = setAttrs(tag, { count: String(items.length) });
      if (extra) tag = extra(tag);
      place(s, name, items.length || name !== "numFmts" ? `${tag}${items.join("")}</${name}>` : null, STYLESHEET_ORDER);
    };
    list("numFmts", [...this.numFmts].map(([id, c]) => `<numFmt numFmtId="${id}" formatCode="${esc(c)}"/>`));
    list("fonts", this.fonts);
    list("fills", this.fills);
    list("borders", this.borders);
    list("cellXfs", this.xfs);
    list("dxfs", this.dxfs);
    return joinSplit(s);
  }
}

function alignmentXml(s: CellStyle, baseXml: string): string {
  const old = /<alignment\b[^>]*\/>/.exec(baseXml)?.[0];
  const keep = old ? Object.entries(parseAttrs(old)).filter(([k]) => !/^(horizontal|vertical|wrapText|indent|textRotation|shrinkToFit)$/.test(k)) : [];
  const parts: string[] = [];
  if (s.h && s.h !== "general") parts.push(`horizontal="${s.h}"`);
  if (s.v && s.v !== "bottom") parts.push(`vertical="${s.v}"`);
  if (s.rotation) parts.push(`textRotation="${s.rotation}"`);
  if (s.wrap) parts.push('wrapText="1"');
  if (s.indent) parts.push(`indent="${s.indent}"`);
  for (const [k, v] of keep) parts.push(`${k}="${esc(v)}"`);
  if (s.shrink) parts.push('shrinkToFit="1"');
  return parts.length ? `<alignment ${parts.join(" ")}/>` : "";
}

/** Excel's styles.xml of a new workbook (Calibri 11, Office 2013-2022 design), as it writes it. */
export const NEW_STYLES_XML = `${XML_HEAD}<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac x16r2 xr" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" xmlns:x16r2="http://schemas.microsoft.com/office/spreadsheetml/2015/02/main" xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision">` +
  `<fonts count="1" x14ac:knownFonts="1"><font><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font></fonts>` +
  `<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>` +
  `<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>` +
  `<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>` +
  `<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>` +
  `<cellStyles count="1"><cellStyle name="Standard" xfId="0" builtinId="0"/></cellStyles><dxfs count="0"/>` +
  `<tableStyles count="0" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16"/></styleSheet>`;
