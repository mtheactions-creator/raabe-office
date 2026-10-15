// DrawingML for pictures and shapes: colours (with the document theme), the Bildformat
// settings of a picture (blip effects, cropping, frame shape, line, effects), read into the
// engine's PictureFormat and written back the way Word writes them (recorded with
// tools/feel-lab/probe_word_picture.py).

import type { Fill, LineEnd, PictureFormat, PictureShadow } from "@raabe/text-engine";

export const NS_A14 = "http://schemas.microsoft.com/office/drawing/2010/main";

/**
 * The colours Word 365 uses for a document without a theme part (its 2023 Office theme).
 * New RaabeWrite documents carry an Office 2013–2022 theme part instead (D-046).
 */
export const DEFAULT_THEME_COLORS: Readonly<Record<string, string>> = {
  dk1: "000000", lt1: "FFFFFF", dk2: "0E2841", lt2: "E8E8E8", accent1: "156082", accent2: "E97132", accent3: "196B24",
  accent4: "0F9ED5", accent5: "A02B93", accent6: "4EA72E", hlink: "467886", folHlink: "96607D",
};

const PRESET: Record<string, string> = {
  black: "000000", white: "FFFFFF", red: "FF0000", green: "008000", blue: "0000FF", yellow: "FFFF00", gray: "808080", grey: "808080",
  darkGray: "A9A9A9", lightGray: "D3D3D3",
};

const kids = (el: Element) => Array.from(el.children);
const kid = (el: Element | undefined, name: string) => (el ? kids(el).find((c) => c.localName === name) : undefined);
const num = (el: Element | undefined, a: string) => {
  const v = el?.getAttribute(a);
  return v === null || v === undefined || v === "" ? undefined : Number(v);
};

/** Theme colours of word/theme/theme1.xml (dk1 … folHlink). */
export function themeColors(doc: Document | null): Record<string, string> {
  const out: Record<string, string> = { ...DEFAULT_THEME_COLORS };
  if (!doc) return out;
  const scheme = doc.getElementsByTagNameNS("http://schemas.openxmlformats.org/drawingml/2006/main", "clrScheme")[0];
  if (!scheme) return out;
  for (const c of kids(scheme)) {
    const v = kids(c)[0];
    const hex = v?.localName === "sysClr" ? v.getAttribute("lastClr") : v?.getAttribute("val");
    if (hex && /^[0-9A-Fa-f]{6}$/.test(hex)) out[c.localName] = hex.toUpperCase();
  }
  return out;
}

function rgbToHsl(r: number, g: number, b: number): [number, number, number] {
  const max = Math.max(r, g, b), min = Math.min(r, g, b);
  const l = (max + min) / 2;
  if (max === min) return [0, 0, l];
  const d = max - min;
  const s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
  const h = max === r ? (g - b) / d + (g < b ? 6 : 0) : max === g ? (b - r) / d + 2 : (r - g) / d + 4;
  return [h / 6, s, l];
}

function hslToRgb(h: number, s: number, l: number): [number, number, number] {
  if (s === 0) return [l, l, l];
  const q = l < 0.5 ? l * (1 + s) : l + s - l * s, p = 2 * l - q;
  const f = (t: number) => {
    t = (t + 1) % 1;
    return t < 1 / 6 ? p + (q - p) * 6 * t : t < 1 / 2 ? q : t < 2 / 3 ? p + (q - p) * (2 / 3 - t) * 6 : p;
  };
  return [f(h + 1 / 3), f(h), f(h - 1 / 3)];
}

const toLinear = (c: number) => (c <= 0.04045 ? c / 12.92 : ((c + 0.055) / 1.055) ** 2.4);
const toSrgb = (c: number) => (c <= 0.0031308 ? c * 12.92 : 1.055 * Math.max(0, c) ** (1 / 2.4) - 0.055);

const toHex = (r: number, g: number, b: number) =>
  [r, g, b].map((x) => Math.round(Math.max(0, Math.min(1, x)) * 255).toString(16).padStart(2, "0")).join("").toUpperCase();

/**
 * A DrawingML colour element's parent (a:solidFill, a:glow …): the colour as hex and its
 * transparency (0 = opaque), with the modifiers Word uses (lumMod, lumOff, tint, shade,
 * satMod, alpha).
 */
export function readColor(parent: Element | undefined, theme: Record<string, string>): { hex: string; transparency: number } | null {
  if (!parent) return null;
  const c = kids(parent).find((x) => ["srgbClr", "schemeClr", "prstClr", "sysClr", "scrgbClr"].includes(x.localName));
  if (!c) return null;
  let hex = "000000";
  const v = c.getAttribute("val") ?? "";
  switch (c.localName) {
    case "srgbClr": hex = v; break;
    case "schemeClr": hex = theme[v === "tx1" ? "dk1" : v === "bg1" ? "lt1" : v === "tx2" ? "dk2" : v === "bg2" ? "lt2" : v] ?? "000000"; break;
    case "prstClr": hex = PRESET[v] ?? "000000"; break;
    case "sysClr": hex = c.getAttribute("lastClr") ?? (v === "window" ? "FFFFFF" : "000000"); break;
    default: hex = "000000";
  }
  let [r, g, b] = [0, 2, 4].map((i) => parseInt(hex.slice(i, i + 2), 16) / 255);
  let alpha = 1;
  for (const m of kids(c)) {
    const val = Number(m.getAttribute("val") ?? 0) / 100000;
    if (m.localName === "alpha") alpha = val;
    else if (m.localName === "lumMod" || m.localName === "lumOff" || m.localName === "satMod") {
      const [h, s, l] = rgbToHsl(r, g, b);
      const next = m.localName === "lumMod" ? [h, s, l * val] : m.localName === "lumOff" ? [h, s, l + val] : [h, Math.min(1, s * val), l];
      [r, g, b] = hslToRgb(next[0], next[1], Math.max(0, Math.min(1, next[2])));
    } else if (m.localName === "tint") [r, g, b] = [r, g, b].map((x) => toSrgb(1 - (1 - toLinear(x)) * val));
    // Shade and tint work in linear light: accent1 4472C4 shaded to 15 % is Word's 172C51.
    else if (m.localName === "shade") [r, g, b] = [r, g, b].map((x) => toSrgb(toLinear(x) * val));
  }
  return { hex: toHex(r, g, b), transparency: 1 - alpha };
}

/** A colour as DrawingML (srgbClr with alpha). */
export function colorXml(hex: string, transparency = 0): string {
  const a = Math.round((1 - transparency) * 100000);
  return `<a:srgbClr val="${hex.toUpperCase()}">${a < 100000 ? `<a:alpha val="${a}"/>` : ""}</a:srgbClr>`;
}

const EMU = 12700;
const pctOf = (v: number | undefined) => (v === undefined ? undefined : v / 100000);

/** The Bildformat settings of a pic:pic element. */
export function readPicture(pic: Element, theme: Record<string, string>): PictureFormat {
  const out: Record<string, unknown> = {};
  const blipFill = kids(pic).find((c) => c.localName === "blipFill");
  const blip = kid(blipFill, "blip");
  if (blip) {
    for (const e of kids(blip)) {
      switch (e.localName) {
        case "alphaModFix": out.transparency = 1 - (pctOf(num(e, "amt")) ?? 1); break;
        case "lum": {
          const b = pctOf(num(e, "bright")) ?? 0, c = pctOf(num(e, "contrast")) ?? 0;
          if (Math.abs(b - 0.7) < 1e-6 && Math.abs(c + 0.7) < 1e-6) out.recolor = { kind: "washout" };
          else if (b || c) out.lum = { bright: b, contrast: c };
          break;
        }
        case "grayscl": out.recolor = { kind: "grayscale" }; break;
        case "clrChange": {
          // Transparente Farbe bestimmen: the colour that becomes see-through.
          const from = kid(e, "clrFrom");
          const hex = from ? readColor(from, theme)?.hex : undefined;
          if (hex) out.transparentColor = hex;
          break;
        }
        case "biLevel": out.recolor = { kind: "bw", threshold: pctOf(num(e, "thresh")) ?? 0.5 }; break;
        case "duotone": {
          const cs = kids(e).map((x) => {
            const holder = e.ownerDocument.createElementNS(e.namespaceURI, "a:wrap");
            holder.append(x.cloneNode(true));
            return readColor(holder, theme)?.hex ?? "000000";
          });
          const sepia = e.innerHTML?.includes("D9C3A5") || new XMLSerializer().serializeToString(e).includes('val="D9C3A5"');
          out.recolor = sepia ? { kind: "sepia" } : { kind: "duotone", dark: cs[0] ?? "000000", light: cs[1] ?? "FFFFFF" };
          break;
        }
        case "extLst":
          for (const fx of Array.from(e.getElementsByTagNameNS(NS_A14, "imgEffect"))) {
            for (const x of kids(fx)) {
              if (x.localName === "brightnessContrast") {
                const b = pctOf(num(x, "bright")), c = pctOf(num(x, "contrast"));
                if (b) out.brightness = b;
                if (c) out.contrast = c;
              } else if (x.localName === "saturation") out.saturation = pctOf(num(x, "sat")) ?? 1;
              else if (x.localName === "colorTemperature") out.temperature = num(x, "colorTemp") ?? 6500;
              else if (x.localName === "sharpenSoften") out.sharpen = pctOf(num(x, "amount")) ?? 0;
            }
          }
          break;
      }
    }
  }
  const src = kid(blipFill, "srcRect");
  if (src) {
    const l = pctOf(num(src, "l")) ?? 0, t = pctOf(num(src, "t")) ?? 0, r = pctOf(num(src, "r")) ?? 0, b = pctOf(num(src, "b")) ?? 0;
    if (l || t || r || b) out.crop = { l, t, r, b };
  }
  const sp = kids(pic).find((c) => c.localName === "spPr");
  if (sp) Object.assign(out, readShapeProps(sp, theme));
  return out as PictureFormat;
}

/** A fill element (a:solidFill, a:gradFill, a:noFill …) as the engine's Fill. */
export function readFill(el: Element | undefined, theme: Record<string, string>): Fill | undefined {
  if (!el) return undefined;
  switch (el.localName) {
    case "noFill": return { kind: "none" };
    case "solidFill": {
      const c = readColor(el, theme);
      return c ? { kind: "solid", color: c.hex, transparency: c.transparency || undefined } : { kind: "solid", color: "000000" };
    }
    case "gradFill": {
      const stops = kids(kid(el, "gsLst") ?? el).filter((g) => g.localName === "gs").map((g) => {
        const c = readColor(g, theme) ?? { hex: "000000", transparency: 0 };
        return { pos: (num(g, "pos") ?? 0) / 100000, color: c.hex, transparency: c.transparency || undefined };
      });
      const lin = kid(el, "lin"), path = kid(el, "path");
      return { kind: "gradient", stops, angle: lin ? (num(lin, "ang") ?? 0) / 60000 : undefined, path: path?.getAttribute("path") ?? undefined };
    }
    case "pattFill": return { kind: "other", color: readColor(kid(el, "fgClr"), theme)?.hex ?? "000000" };
    case "blipFill": return { kind: "other", color: "BFBFBF" };
    case "grpFill": return { kind: "none" };
    default: return undefined;
  }
}

const FILLS = ["solidFill", "gradFill", "pattFill", "noFill", "blipFill", "grpFill"];
const END_TYPES = new Set(["triangle", "stealth", "diamond", "oval", "arrow"]);

/** An outline (a:ln) as the engine's line; null when it has no fill. */
export function readLine(ln: Element, theme: Record<string, string>, style?: { color: string; width: number }): PictureFormat["line"] {
  if (kid(ln, "noFill")) return null;
  const fill = kids(ln).find((x) => FILLS.includes(x.localName));
  const f = readFill(fill, theme);
  // Without a fill of its own the outline takes the shape style's colour (and width).
  const color = f?.kind === "solid" ? { hex: f.color, transparency: f.transparency ?? 0 } : f?.kind === "gradient" ? { hex: f.stops[0]?.color ?? "000000", transparency: 0 }
    : style ? { hex: style.color, transparency: 0 } : null;
  if (!color) return undefined;
  const end = (name: string): LineEnd | undefined => {
    const e = kid(ln, name);
    const type = e?.getAttribute("type");
    if (!e || !type || !END_TYPES.has(type)) return undefined;
    return { type: type as LineEnd["type"], w: (e.getAttribute("w") ?? undefined) as LineEnd["w"], len: (e.getAttribute("len") ?? undefined) as LineEnd["len"] };
  };
  const compound = ln.getAttribute("cmpd") ?? undefined;
  const join = kid(ln, "round") ? "round" : kid(ln, "bevel") ? "bevel" : kid(ln, "miter") ? "miter" : undefined;
  return {
    color: color.hex, width: num(ln, "w") !== undefined ? num(ln, "w")! / EMU : style?.width ?? 0.75, dash: kid(ln, "prstDash")?.getAttribute("val") ?? undefined,
    compound: compound && compound !== "sng" ? compound : undefined, cap: ln.getAttribute("cap") ?? undefined,
    miter: join === "miter" ? true : undefined, join: join && join !== "miter" ? join : undefined,
    transparency: color.transparency || undefined, head: end("headEnd"), tail: end("tailEnd"),
  };
}

/** An effect list (a:effectLst): shadow, reflection, glow, soft edges. */
export function readEffects(fx: Element, theme: Record<string, string>): Partial<PictureFormat> {
  const out: Record<string, unknown> = {};
  for (const e of kids(fx)) {
    if (e.localName === "outerShdw" || e.localName === "innerShdw") {
      const color = readColor(e, theme) ?? { hex: "000000", transparency: 0.6 };
      const shadow: PictureShadow = {
        kind: e.localName === "outerShdw" ? "outer" : "inner", color: color.hex, transparency: color.transparency,
        blur: (num(e, "blurRad") ?? 0) / EMU, dist: (num(e, "dist") ?? 0) / EMU, dir: (num(e, "dir") ?? 0) / 60000,
        sx: pctOf(num(e, "sx")), sy: pctOf(num(e, "sy")),
        kx: num(e, "kx") !== undefined ? num(e, "kx")! / 60000 : undefined, ky: num(e, "ky") !== undefined ? num(e, "ky")! / 60000 : undefined,
        algn: e.getAttribute("algn") ?? undefined,
        rotWithShape: e.hasAttribute("rotWithShape") ? e.getAttribute("rotWithShape") === "1" : undefined,
      };
      out.shadow = shadow;
    } else if (e.localName === "reflection") {
      out.reflection = {
        size: pctOf(num(e, "endPos")) ?? 0.9, dist: (num(e, "dist") ?? 0) / EMU, transparency: 1 - (pctOf(num(e, "stA")) ?? 0.5),
        blur: (num(e, "blurRad") ?? 0) / EMU, endA: pctOf(num(e, "endA")),
      };
    } else if (e.localName === "glow") {
      const color = readColor(e, theme) ?? { hex: "4472C4", transparency: 0.6 };
      out.glow = { color: color.hex, radius: (num(e, "rad") ?? 0) / EMU, transparency: color.transparency };
    } else if (e.localName === "softEdge") out.softEdge = (num(e, "rad") ?? 0) / EMU;
  }
  return out as Partial<PictureFormat>;
}

/** Shape properties (pic:spPr, wps:spPr): transform, geometry, fill, outline, effects, 3D. */
export function readShapeProps(sp: Element, theme: Record<string, string>, styleLine?: { color: string; width: number }): PictureFormat {
  const out: Record<string, unknown> = {};
  const xfrm = kid(sp, "xfrm");
  if (xfrm) {
    const rot = num(xfrm, "rot");
    if (rot) out.rotation = rot / 60000;
    if (xfrm.getAttribute("flipH") === "1") out.flipH = true;
    if (xfrm.getAttribute("flipV") === "1") out.flipV = true;
  }
  const prst = kid(sp, "prstGeom");
  const geom = prst?.getAttribute("prst");
  if (geom && geom !== "rect") out.geometry = geom;
  if (kid(sp, "custGeom")) out.geometry = "custom";
  const av = kid(prst, "avLst");
  if (av && av.children.length) out.geometryAdj = Array.from(av.children).map((x) => new XMLSerializer().serializeToString(x)).join("");
  const ser = (x: Element | undefined) => (x ? new XMLSerializer().serializeToString(x) : undefined);
  const fill = kids(sp).find((x) => FILLS.includes(x.localName));
  if (fill) {
    out.fillXml = ser(fill);
    out.fill = readFill(fill, theme);
  }
  if (kid(sp, "scene3d")) out.scene3dXml = ser(kid(sp, "scene3d"));
  if (kid(sp, "sp3d")) out.sp3dXml = ser(kid(sp, "sp3d"));
  const ln = kid(sp, "ln");
  if (ln) {
    const line = readLine(ln, theme, styleLine);
    if (line !== undefined) out.line = line;
  }
  const fx = kid(sp, "effectLst");
  if (fx) Object.assign(out, readEffects(fx, theme));
  const bevel = kid(kid(sp, "sp3d"), "bevelT");
  if (bevel) out.bevel = { kind: bevel.getAttribute("prst") ?? "circle", w: (num(bevel, "w") ?? 76200) / EMU, h: (num(bevel, "h") ?? 76200) / EMU };
  return out as PictureFormat;
}

/** The space effects take (wp:effectExtent) in points, or undefined when none. */
export function readExtent(el: Element | undefined): PictureFormat["extent"] {
  if (!el) return undefined;
  const v = (a: string) => (num(el, a) ?? 0) / EMU;
  const x = { l: v("l"), t: v("t"), r: v("r"), b: v("b") };
  return x.l || x.t || x.r || x.b ? x : undefined;
}

/**
 * The settings in a fragment of pic:spPr children (a picture style or effect preset, as Word
 * writes it), parsed like a picture's own.
 */
export function pictureFromSpPr(fragment: string, theme: Record<string, string> = DEFAULT_THEME_COLORS): PictureFormat {
  const xml = `<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><pic:spPr>${fragment}</pic:spPr></pic:pic>`;
  const doc = new DOMParser().parseFromString(xml, "application/xml");
  return readPicture(doc.documentElement, theme);
}

const e = (pt: number) => Math.round(pt * EMU);
const p = (x: number) => Math.round(x * 100000);

/** The blip's children for a picture's colour settings (Word's order and elements). */
export function blipEffectsXml(f: PictureFormat): string {
  let out = "";
  if (f.transparentColor) {
    const c = f.transparentColor.toUpperCase();
    out += `<a:clrChange><a:clrFrom><a:srgbClr val="${c}"/></a:clrFrom><a:clrTo><a:srgbClr val="${c}"><a:alpha val="0"/></a:srgbClr></a:clrTo></a:clrChange>`;
  }
  if (f.transparency !== undefined && f.transparency > 0) out += `<a:alphaModFix amt="${p(1 - f.transparency)}"/>`;
  if (f.lum) out += `<a:lum${f.lum.bright ? ` bright="${p(f.lum.bright)}"` : ""}${f.lum.contrast ? ` contrast="${p(f.lum.contrast)}"` : ""}/>`;
  const rc = f.recolor;
  if (rc?.kind === "grayscale") out += "<a:grayscl/>";
  else if (rc?.kind === "sepia") out += '<a:duotone><a:prstClr val="black"/><a:srgbClr val="D9C3A5"><a:tint val="50000"/><a:satMod val="180000"/></a:srgbClr></a:duotone>';
  else if (rc?.kind === "washout") out += '<a:lum bright="70000" contrast="-70000"/>';
  else if (rc?.kind === "bw") out += `<a:biLevel thresh="${p(rc.threshold)}"/>`;
  else if (rc?.kind === "duotone") out += `<a:duotone>${colorXml(rc.dark)}${colorXml(rc.light)}</a:duotone>`;
  return out;
}

/** a14 picture effects (Office 2010+): sharpen, brightness/contrast, saturation, temperature. */
export function imgEffectsXml(f: PictureFormat, layerRid?: string): string {
  const fx: string[] = [];
  if (f.sharpen) fx.push(`<a14:imgEffect><a14:sharpenSoften amount="${p(f.sharpen)}"/></a14:imgEffect>`);
  if (f.saturation !== undefined && f.saturation !== 1) fx.push(`<a14:imgEffect><a14:saturation sat="${p(f.saturation)}"/></a14:imgEffect>`);
  if (f.temperature !== undefined && f.temperature !== 6500) fx.push(`<a14:imgEffect><a14:colorTemperature colorTemp="${Math.round(f.temperature)}"/></a14:imgEffect>`);
  if (f.brightness || f.contrast) fx.push(`<a14:imgEffect><a14:brightnessContrast${f.brightness ? ` bright="${p(f.brightness)}"` : ""}${f.contrast ? ` contrast="${p(f.contrast)}"` : ""}/></a14:imgEffect>`);
  if (!fx.length) return "";
  return `<a:ext uri="{BEBA8EAE-BF5A-486C-A8C5-ECC9F3942E4B}"><a14:imgProps xmlns:a14="${NS_A14}"><a14:imgLayer${layerRid ? ` r:embed="${layerRid}"` : ""}>${fx.join("")}</a14:imgLayer></a14:imgProps></a:ext>`;
}

/** pic:spPr children after a:xfrm: frame shape, line, effects, bevel. */
export function shapePropsXml(f: PictureFormat, cx: number, cy: number): string {
  const rot = f.rotation ? ` rot="${Math.round(f.rotation * 60000)}"` : "";
  const flip = `${f.flipH ? ' flipH="1"' : ""}${f.flipV ? ' flipV="1"' : ""}`;
  let out = `<a:xfrm${rot}${flip}><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="${f.geometry ?? "rect"}"><a:avLst>${f.geometryAdj ?? ""}</a:avLst></a:prstGeom>`;
  if (f.fillXml) out += f.fillXml;
  if (f.line === null) out += "<a:ln><a:noFill/></a:ln>";
  else if (f.line) {
    const l = f.line;
    const end = (tag: string, x: LineEnd | undefined) => (x ? `<a:${tag} type="${x.type}"${x.w ? ` w="${x.w}"` : ""}${x.len ? ` len="${x.len}"` : ""}/>` : "");
    out += `<a:ln${l.width >= 0 ? ` w="${e(l.width)}"` : ""}${l.cap ? ` cap="${l.cap}"` : ""}${l.compound ? ` cmpd="${l.compound}"` : ""}>${l.color ? `<a:solidFill>${colorXml(l.color, l.transparency ?? 0)}</a:solidFill>` : ""}` +
      `${l.dash ? `<a:prstDash val="${l.dash}"/>` : ""}${l.join === "round" ? "<a:round/>" : l.join === "bevel" ? "<a:bevel/>" : l.miter ? '<a:miter lim="800000"/>' : ""}${end("headEnd", l.head)}${end("tailEnd", l.tail)}</a:ln>`;
  }
  const fx: string[] = [];
  if (f.glow) fx.push(`<a:glow rad="${e(f.glow.radius)}">${colorXml(f.glow.color, f.glow.transparency)}</a:glow>`);
  const s = f.shadow;
  const pos = (sh: NonNullable<PictureFormat["shadow"]>) => `${sh.blur ? ` blurRad="${e(sh.blur)}"` : ""}${sh.dist ? ` dist="${e(sh.dist)}"` : ""}${sh.dir ? ` dir="${Math.round(sh.dir * 60000)}"` : ""}`;
  if (s?.kind === "inner") fx.push(`<a:innerShdw${pos(s)}>${colorXml(s.color, s.transparency)}</a:innerShdw>`);
  if (s?.kind === "outer") {
    const k = (v: number | undefined, a: string) => (v !== undefined ? ` ${a}="${Math.round(v * 60000)}"` : "");
    fx.push(`<a:outerShdw${pos(s)}${s.sx !== undefined ? ` sx="${p(s.sx)}"` : ""}${s.sy !== undefined ? ` sy="${p(s.sy)}"` : ""}` +
      `${k(s.kx, "kx")}${k(s.ky, "ky")}${s.algn ? ` algn="${s.algn}"` : ""}${s.rotWithShape !== undefined ? ` rotWithShape="${s.rotWithShape ? 1 : 0}"` : ""}>${colorXml(s.color, s.transparency)}</a:outerShdw>`);
  }
  const r = f.reflection;
  if (r) fx.push(`<a:reflection${r.blur ? ` blurRad="${e(r.blur)}"` : ""} stA="${p(1 - r.transparency)}"${r.endA !== undefined ? ` endA="${p(r.endA)}"` : ""} endPos="${p(r.size)}"${r.dist ? ` dist="${e(r.dist)}"` : ""} dir="5400000" sy="-100000" algn="bl" rotWithShape="0"/>`);
  if (f.softEdge) fx.push(`<a:softEdge rad="${e(f.softEdge)}"/>`);
  if (fx.length) out += `<a:effectLst>${fx.join("")}</a:effectLst>`;
  if (f.scene3dXml) out += f.scene3dXml;
  else if (f.bevel) out += '<a:scene3d><a:camera prst="orthographicFront"/><a:lightRig rig="threePt" dir="t"/></a:scene3d>';
  // 6 pt is DrawingML's default bevel size (Word leaves it out).
  const bv = (a: string, v: number) => (e(v) !== 76200 ? ` ${a}="${e(v)}"` : "");
  const bevelT = f.bevel ? `<a:bevelT${bv("w", f.bevel.w)}${bv("h", f.bevel.h)}${f.bevel.kind !== "circle" ? ` prst="${f.bevel.kind}"` : ""}/>` : "";
  if (f.sp3dXml) out += f.sp3dXml.replace(/<a:bevelT[^>]*\/>/, "").replace(/(<a:sp3d[^>]*?)(\/>|>)/, (_m, open: string, end: string) => end === "/>" ? `${open}>${bevelT}</a:sp3d>` : `${open}>${bevelT}`);
  else if (bevelT) out += `<a:sp3d>${bevelT}</a:sp3d>`;
  return out;
}
