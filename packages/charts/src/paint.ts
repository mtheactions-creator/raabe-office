// The Format pane's fills, lines and effects as SVG paint (D-066, D-111): an element's own formatting
// where it has one, else Excel's automatic look passed in by the renderer. Gradients, patterns,
// pictures, arrowheads, compound lines and effects need SVG definitions; while a scene is built
// they are collected in `defs` (beginDefs / endDefs) and referenced as url(#id).

import { autoScale, type AxisScale } from "./scale";
import type { Arrow, AxisOptions, Dash, Effects, Fill, Gradient, Line, TextFormat } from "./model";

/** The SVG definitions of the scene being built. */
export interface Defs { list: string[]; prefix: string; n: number; cache: Map<string, string> }

let sink: Defs | null = null;

/** Starts collecting definitions for a scene (ids unique per scene so several charts share a page). */
export function beginDefs(): Defs {
  sink = { list: [], prefix: `cd${Math.random().toString(36).slice(2, 7)}`, n: 0, cache: new Map() };
  return sink;
}

export function endDefs(): void { sink = null; }

/** Adds a definition (once per distinct markup) and returns its id. */
function define(make: (id: string) => string, key: string): string | undefined {
  if (!sink) return undefined;
  const hit = sink.cache.get(key);
  if (hit) return hit;
  const id = `${sink.prefix}${++sink.n}`;
  sink.list.push(make(id));
  sink.cache.set(key, id);
  return id;
}

const escAttr = (s: string) => s.replace(/&/g, "&amp;").replace(/"/g, "&quot;").replace(/</g, "&lt;");

/** A colour with Transparenz (0-100 %) as #RRGGBBAA. */
export function withAlpha(color: string, transparency?: number): string {
  if (!transparency) return color;
  const a = Math.round((1 - Math.max(0, Math.min(100, transparency)) / 100) * 255);
  return color.slice(0, 7) + a.toString(16).padStart(2, "0").toUpperCase();
}

const opacity = (t?: number) => (t ? String(Math.round((1 - Math.max(0, Math.min(100, t)) / 100) * 1000) / 1000) : "1");

/** A gradient definition; returns url(#id) (or the first stop's colour outside a scene). */
export function gradientPaint(g: Gradient): string {
  const stops = [...g.stops].sort((a, b) => a.pos - b.pos);
  const first = stops[0]?.color ?? "#FFFFFF";
  const body = stops.map((s) => `<stop offset="${Math.max(0, Math.min(100, s.pos))}%" stop-color="${s.color.slice(0, 7)}" stop-opacity="${opacity(s.transparency)}"/>`).join("");
  const key = JSON.stringify(g);
  const id = define((id) => {
    if (g.kind === "linear") {
      // DrawingML's angle turns the left-to-right direction clockwise.
      const a = g.angle ?? 0;
      return `<linearGradient id="${id}" x1="0" y1="0.5" x2="1" y2="0.5" gradientTransform="rotate(${a} 0.5 0.5)">${body}</linearGradient>`;
    }
    // Radial, Rechteckig and Pfad from the centre outwards (a rectangle drawn as a circle).
    return `<radialGradient id="${id}" cx="0.5" cy="0.5" r="${g.kind === "radial" ? 0.5 : 0.71}">${body}</radialGradient>`;
  }, "g" + key);
  return id ? `url(#${id})` : first;
}

/**
 * DrawingML's preset patterns (a:pattFill prst) as 8 x 8 bitmaps, one byte per row (bit 7 left):
 * the Musterfüllung grid of the Format pane.
 */
export const PATTERNS: Record<string, number[]> = {
  pct5: [0x80, 0, 0, 0, 0x08, 0, 0, 0], pct10: [0x80, 0, 0x08, 0, 0x80, 0, 0x08, 0], pct20: [0x88, 0, 0x22, 0, 0x88, 0, 0x22, 0],
  pct25: [0x88, 0x22, 0x88, 0x22, 0x88, 0x22, 0x88, 0x22], pct30: [0xAA, 0x44, 0xAA, 0x11, 0xAA, 0x44, 0xAA, 0x11],
  pct40: [0xAA, 0x55, 0xAA, 0x15, 0xAA, 0x55, 0xAA, 0x51], pct50: [0xAA, 0x55, 0xAA, 0x55, 0xAA, 0x55, 0xAA, 0x55],
  pct60: [0xEE, 0x55, 0xBB, 0x55, 0xEE, 0x55, 0xBB, 0x55], pct70: [0x55, 0xBB, 0x55, 0xEE, 0x55, 0xBB, 0x55, 0xEE].map((x) => x | 0x22),
  pct75: [0x77, 0xDD, 0x77, 0xDD, 0x77, 0xDD, 0x77, 0xDD], pct80: [0x77, 0xFF, 0xDD, 0xFF, 0x77, 0xFF, 0xDD, 0xFF],
  pct90: [0x7F, 0xFF, 0xF7, 0xFF, 0x7F, 0xFF, 0xF7, 0xFF],
  horz: [0xFF, 0, 0, 0, 0, 0, 0, 0], vert: [0x80, 0x80, 0x80, 0x80, 0x80, 0x80, 0x80, 0x80],
  ltHorz: [0xFF, 0, 0, 0, 0xFF, 0, 0, 0], ltVert: [0x88, 0x88, 0x88, 0x88, 0x88, 0x88, 0x88, 0x88],
  dkHorz: [0xFF, 0xFF, 0, 0, 0xFF, 0xFF, 0, 0], dkVert: [0xCC, 0xCC, 0xCC, 0xCC, 0xCC, 0xCC, 0xCC, 0xCC],
  narHorz: [0xFF, 0, 0xFF, 0, 0xFF, 0, 0xFF, 0], narVert: [0xAA, 0xAA, 0xAA, 0xAA, 0xAA, 0xAA, 0xAA, 0xAA],
  dashHorz: [0xF0, 0, 0, 0, 0x0F, 0, 0, 0], dashVert: [0x80, 0x80, 0x80, 0x80, 0x08, 0x08, 0x08, 0x08],
  cross: [0xFF, 0x80, 0x80, 0x80, 0x80, 0x80, 0x80, 0x80],
  dnDiag: [0x80, 0x40, 0x20, 0x10, 0x08, 0x04, 0x02, 0x01], upDiag: [0x01, 0x02, 0x04, 0x08, 0x10, 0x20, 0x40, 0x80],
  ltDnDiag: [0x88, 0x44, 0x22, 0x11, 0x88, 0x44, 0x22, 0x11], ltUpDiag: [0x11, 0x22, 0x44, 0x88, 0x11, 0x22, 0x44, 0x88],
  dkDnDiag: [0xCC, 0x66, 0x33, 0x99, 0xCC, 0x66, 0x33, 0x99], dkUpDiag: [0x33, 0x66, 0xCC, 0x99, 0x33, 0x66, 0xCC, 0x99],
  wdDnDiag: [0xC1, 0xE0, 0x70, 0x38, 0x1C, 0x0E, 0x07, 0x83], wdUpDiag: [0x83, 0x07, 0x0E, 0x1C, 0x38, 0x70, 0xE0, 0xC1],
  dashDnDiag: [0x88, 0x44, 0x22, 0x11, 0, 0, 0, 0], dashUpDiag: [0x11, 0x22, 0x44, 0x88, 0, 0, 0, 0],
  diagCross: [0x81, 0x42, 0x24, 0x18, 0x18, 0x24, 0x42, 0x81],
  smCheck: [0xCC, 0xCC, 0x33, 0x33, 0xCC, 0xCC, 0x33, 0x33], lgCheck: [0xF0, 0xF0, 0xF0, 0xF0, 0x0F, 0x0F, 0x0F, 0x0F],
  smGrid: [0xFF, 0x88, 0x88, 0x88, 0xFF, 0x88, 0x88, 0x88], lgGrid: [0xFF, 0x80, 0x80, 0x80, 0x80, 0x80, 0x80, 0x80],
  dotGrid: [0xAA, 0, 0x80, 0, 0x80, 0, 0x80, 0],
  smConfetti: [0x80, 0x04, 0x40, 0x02, 0x20, 0x01, 0x10, 0x08], lgConfetti: [0x4C, 0x0C, 0xC0, 0xC8, 0x13, 0x30, 0x32, 0x03],
  horzBrick: [0xFF, 0x80, 0x80, 0x80, 0xFF, 0x08, 0x08, 0x08], diagBrick: [0x01, 0x02, 0x04, 0x08, 0x18, 0x24, 0x42, 0x81],
  solidDmnd: [0x10, 0x38, 0x7C, 0xFE, 0x7C, 0x38, 0x10, 0], openDmnd: [0x80, 0x41, 0x22, 0x14, 0x08, 0x14, 0x22, 0x41],
  dotDmnd: [0x80, 0, 0x22, 0, 0x08, 0, 0x22, 0], plaid: [0xAA, 0x55, 0xAA, 0x55, 0xF0, 0xF0, 0xF0, 0xF0],
  sphere: [0x77, 0x98, 0xF8, 0xF8, 0x77, 0x89, 0x8F, 0x8F], weave: [0x88, 0x54, 0x22, 0x45, 0x88, 0x15, 0x22, 0x51],
  divot: [0, 0x10, 0x08, 0x10, 0, 0x01, 0x80, 0x01], shingle: [0x03, 0x84, 0x48, 0x30, 0x0C, 0x02, 0x01, 0x01],
  wave: [0, 0x18, 0xA4, 0x03, 0, 0x18, 0xA4, 0x03], trellis: [0xFF, 0x66, 0xFF, 0x99, 0xFF, 0x66, 0xFF, 0x99],
  zigZag: [0x81, 0x42, 0x24, 0x18, 0x81, 0x42, 0x24, 0x18],
};

/** The pattern's foreground cells as an SVG path in a 8 x 8 grid of `cell` units. */
export function patternPath(preset: string, cell = 0.75): string {
  const rows = PATTERNS[preset] ?? PATTERNS.pct50;
  let d = "";
  rows.forEach((bits, y) => {
    for (let x = 0; x < 8; x++) if (bits & (0x80 >> x)) d += `M${x * cell} ${y * cell}h${cell}v${cell}h${-cell}z`;
  });
  return d;
}

/** The fill to draw: "none", the element's colour or paint reference, or the automatic one. */
export function fillPaint(fill: Fill | undefined, auto: string | undefined): string | undefined {
  if (!fill || fill.type === "auto") return auto;
  if (fill.type === "none") return "none";
  if (fill.type === "gradient" && fill.gradient?.stops.length) return gradientPaint(fill.gradient);
  if (fill.type === "pattern" && fill.pattern) {
    const p = fill.pattern;
    const id = define((id) => `<pattern id="${id}" patternUnits="userSpaceOnUse" width="6" height="6"><rect width="6" height="6" fill="${p.bg.slice(0, 7)}"/><path d="${patternPath(p.preset)}" fill="${p.fg.slice(0, 7)}"/></pattern>`, "p" + JSON.stringify(p));
    return id ? `url(#${id})` : p.fg;
  }
  if (fill.type === "picture" && fill.picture?.src) {
    const pic = fill.picture;
    const op = opacity(pic.transparency);
    const id = define((id) => pic.tile
      ? `<pattern id="${id}" patternUnits="userSpaceOnUse" width="48" height="48"><image href="${escAttr(pic.src)}" width="48" height="48" opacity="${op}" preserveAspectRatio="none"/></pattern>`
      : `<pattern id="${id}" patternContentUnits="objectBoundingBox" width="1" height="1"><image href="${escAttr(pic.src)}" width="1" height="1" opacity="${op}" preserveAspectRatio="none"/></pattern>`, "i" + pic.src.length + pic.src.slice(-40) + pic.tile + op);
    return id ? `url(#${id})` : auto;
  }
  if (fill.type === "gradient" || fill.type === "picture" || fill.type === "pattern") return auto;
  return withAlpha(fill.color ?? auto ?? "#FFFFFF", fill.transparency);
}

/** A representative solid colour of a fill (legend keys of patterns, 3-D shading). */
export function fillColor(fill: Fill | undefined, auto: string): string {
  if (!fill || fill.type === "auto" || fill.type === "none") return auto;
  if (fill.type === "gradient") return fill.gradient?.stops[0]?.color ?? auto;
  if (fill.type === "pattern") return fill.pattern?.fg ?? auto;
  if (fill.type === "picture") return auto;
  return fill.color ?? auto;
}

export interface Stroke {
  stroke?: string;
  sw?: number;
  dash?: string;
  cap?: "round" | "butt" | "square";
  join?: "round" | "bevel" | "miter";
  /** Arrowheads (marker references). */
  head?: string;
  tail?: string;
  /** Verbundtyp other than Einfach. */
  compound?: "dbl" | "thickThin" | "thinThick" | "tri";
}

/** An arrowhead marker in the line's colour; returns its url. */
function arrowMarker(a: Arrow, color: string, start: boolean): string | undefined {
  const size = { sm: 2, med: 3, lg: 5 };
  const w = size[a.w ?? "med"], l = size[a.len ?? "med"];
  const shape: Record<Arrow["type"], string> = {
    triangle: `M0 0L${l} ${w / 2}L0 ${w}z`,
    stealth: `M0 0L${l} ${w / 2}L0 ${w}L${l * 0.35} ${w / 2}z`,
    diamond: `M0 ${w / 2}L${l / 2} 0L${l} ${w / 2}L${l / 2} ${w}z`,
    oval: `M0 ${w / 2}a${l / 2} ${w / 2} 0 1 0 ${l} 0a${l / 2} ${w / 2} 0 1 0 ${-l} 0z`,
    arrow: `M0 0L${l} ${w / 2}L0 ${w}`,
  };
  const open = a.type === "arrow";
  const id = define((id) => `<marker id="${id}" markerUnits="strokeWidth" markerWidth="${l}" markerHeight="${w}" refX="${a.type === "triangle" || open ? l * 0.9 : l / 2}" refY="${w / 2}" orient="${start ? "auto-start-reverse" : "auto"}" overflow="visible"><path d="${shape[a.type]}" fill="${open ? "none" : color.slice(0, 7)}"${open ? ` stroke="${color.slice(0, 7)}" stroke-width="0.6"` : ""}/></marker>`, `a${JSON.stringify(a)}${color}${start}`);
  return id ? `url(#${id})` : undefined;
}

/** The line to draw (null: no line), with DrawingML's preset dashes scaled to the width as Excel draws them. */
export function linePaint(line: Line | undefined, auto: { color?: string; width: number } | null): Stroke | null {
  if (line?.type === "none") return null;
  let out: Stroke;
  if (!line || line.type === "auto") {
    if (!auto?.color) return null;
    out = { stroke: auto.color, sw: line?.width ?? auto.width };
  } else {
    const sw = line.width ?? auto?.width ?? 0.75;
    const color = line.type === "gradient" && line.gradient?.stops.length ? gradientPaint(line.gradient) : withAlpha(line.color ?? auto?.color ?? "#000000", line.transparency);
    out = { stroke: color, sw };
  }
  const sw = out.sw ?? 0.75;
  if (line?.dash) out.dash = dashArray(line.dash, sw);
  if (line?.cap) out.cap = line.cap === "rnd" ? "round" : line.cap === "sq" ? "square" : "butt";
  if (line?.join) out.join = line.join;
  if (line?.compound && line.compound !== "sng") out.compound = line.compound;
  const solid = out.stroke && !out.stroke.startsWith("url") ? out.stroke : (line?.gradient?.stops[0]?.color ?? "#000000");
  if (line?.head) out.head = arrowMarker(line.head, solid, true);
  if (line?.tail) out.tail = arrowMarker(line.tail, solid, false);
  return out;
}

/** DrawingML's preset dashes (lengths in line widths). */
export function dashArray(d: Dash, w: number): string | undefined {
  const pat: Record<Dash, number[]> = {
    solid: [], dot: [1, 3], sysDot: [1, 1], sysDash: [3, 1], sysDashDot: [3, 1, 1, 1], sysDashDotDot: [3, 1, 1, 1, 1, 1], dash: [4, 3], dashDot: [4, 3, 1, 3],
    lgDash: [8, 3], lgDashDot: [8, 3, 1, 3], lgDashDotDot: [8, 3, 1, 3, 1, 3],
  };
  // A preset the file has that is not in the list draws solid.
  const p = pat[d] ?? [];
  return p.length ? p.map((x) => String(Math.round(x * w * 100) / 100)).join(" ") : undefined;
}

/**
 * An SVG filter for an element's Effekte (Schatten, Leuchten, Weiche Kanten, 3D-Format); undefined
 * when it has none. Sizes are in pt (the scene's unit).
 */
export function effectFilter(e: Effects | undefined): string | undefined {
  if (!e || (!e.shadow && !e.glow && !e.softEdges && !e.format3d)) return undefined;
  const id = define((id) => {
    const parts: string[] = [];
    let src = "SourceGraphic";
    if (e.softEdges) {
      // Weiche Kanten: the edge fades over the given width.
      const r = e.softEdges / 2;
      parts.push(`<feMorphology in="SourceAlpha" operator="erode" radius="${r}" result="se1"/><feGaussianBlur in="se1" stdDeviation="${r}" result="se2"/><feComposite in="SourceGraphic" in2="se2" operator="in" result="soft"/>`);
      src = "soft";
    }
    if (e.format3d?.top) {
      // Abschrägung: a light from the top left over the edge's slope.
      const bw = e.format3d.top.w ?? 6;
      parts.push(`<feGaussianBlur in="SourceAlpha" stdDeviation="${Math.max(0.5, bw / 3)}" result="bv1"/><feSpecularLighting in="bv1" surfaceScale="${Math.max(1, (e.format3d.top.h ?? 6) / 2)}" specularConstant="0.8" specularExponent="18" lighting-color="#FFFFFF" result="bv2"><feDistantLight azimuth="225" elevation="40"/></feSpecularLighting><feComposite in="bv2" in2="SourceAlpha" operator="in" result="bv3"/><feComposite in="${src}" in2="bv3" operator="arithmetic" k1="0" k2="1" k3="0.6" k4="0" result="bevel"/>`);
      src = "bevel";
    }
    const under: string[] = [];
    if (e.glow) {
      const g = e.glow;
      parts.push(`<feMorphology in="SourceAlpha" operator="dilate" radius="${g.size / 2}" result="gl1"/><feGaussianBlur in="gl1" stdDeviation="${g.size / 3}" result="gl2"/><feFlood flood-color="${g.color.slice(0, 7)}" flood-opacity="${opacity(g.transparency ?? 60)}"/><feComposite in2="gl2" operator="in" result="glow"/>`);
      under.push("glow");
    }
    if (e.shadow) {
      const s = e.shadow;
      const a = ((s.angle ?? 45) * Math.PI) / 180, dist = s.distance ?? 3;
      const dx = Math.round(Math.cos(a) * dist * 100) / 100, dy = Math.round(Math.sin(a) * dist * 100) / 100;
      const op = opacity(s.transparency ?? 60);
      if (s.inner) {
        parts.push(`<feFlood flood-color="${(s.color ?? "#000000").slice(0, 7)}" flood-opacity="${op}"/><feComposite in2="SourceAlpha" operator="out" result="is1"/><feGaussianBlur in="is1" stdDeviation="${(s.blur ?? 4) / 2}" result="is2"/><feOffset in="is2" dx="${dx}" dy="${dy}" result="is3"/><feComposite in="is3" in2="SourceAlpha" operator="in" result="ishadow"/>`);
        parts.push(`<feMerge result="withInner"><feMergeNode in="${src}"/><feMergeNode in="ishadow"/></feMerge>`);
        src = "withInner";
      } else {
        const sc = (s.size ?? 100) / 100;
        parts.push(`<feGaussianBlur in="SourceAlpha" stdDeviation="${(s.blur ?? 4) / 2}" result="sh1"/><feOffset in="sh1" dx="${dx}" dy="${dy}" result="sh2"/>${sc !== 1 ? `<feMorphology in="sh2" operator="${sc > 1 ? "dilate" : "erode"}" radius="${Math.abs(sc - 1) * 3}" result="sh2"/>` : ""}<feFlood flood-color="${(s.color ?? "#000000").slice(0, 7)}" flood-opacity="${op}"/><feComposite in2="sh2" operator="in" result="shadow"/>`);
        under.unshift("shadow");
      }
    }
    const merge = `<feMerge>${under.map((u) => `<feMergeNode in="${u}"/>`).join("")}<feMergeNode in="${src}"/></feMerge>`;
    return `<filter id="${id}" x="-50%" y="-50%" width="200%" height="200%" color-interpolation-filters="sRGB">${parts.join("")}${merge}</filter>`;
  }, "f" + JSON.stringify(e));
  return id ? `url(#${id})` : undefined;
}

/** Text paint: own size, colour, weight, fill, outline and effects, else the automatic ones. */
export function textLook(t: TextFormat | undefined, auto: { size: number; color: string }): {
  size: number; fill: string; bold?: boolean; italic?: boolean; underline?: boolean; strike?: boolean; font?: string; outline?: Stroke | null; filter?: string; rotate?: number; reflection?: { opacity: number; dist: number; blur: number };
} {
  let fill = t?.color ?? auto.color;
  if (t?.fill && t.fill.type !== "auto") fill = t.fill.type === "solid" ? withAlpha(t.fill.color ?? fill, t.fill.transparency) : fillPaint(t.fill, fill) ?? fill;
  const out: ReturnType<typeof textLook> = { size: t?.size ?? auto.size, fill, bold: t?.bold, italic: t?.italic };
  if (t?.underline) out.underline = true;
  if (t?.strike) out.strike = true;
  if (t?.font) out.font = t.font;
  if (t?.outline && t.outline.type !== "auto") out.outline = linePaint(t.outline, { color: "#000000", width: 0.75 });
  const fx = effectFilter(t?.effects);
  if (fx) out.filter = fx;
  const rf = t?.effects?.reflection;
  if (rf) out.reflection = { opacity: Math.round((1 - rf.transparency / 100) * 100) / 100, dist: rf.distance, blur: rf.blur };
  return out;
}

/** The text look's attributes for a text primitive (spread into it). */
export function textAttrs(l: ReturnType<typeof textLook>): { size: number; fill: string; bold?: boolean; italic?: boolean; underline?: boolean; strike?: boolean; font?: string; outline?: Stroke; filter?: string; reflection?: { opacity: number; dist: number; blur: number } } {
  const o: ReturnType<typeof textAttrs> = { size: l.size, fill: l.fill };
  if (l.bold) o.bold = true;
  if (l.italic) o.italic = true;
  if (l.underline) o.underline = true;
  if (l.strike) o.strike = true;
  if (l.font) o.font = l.font;
  if (l.outline) o.outline = l.outline;
  if (l.filter) o.filter = l.filter;
  if (l.reflection) o.reflection = l.reflection;
  return o;
}

/** Excel's Anzeigeeinheiten as divisors. */
export const DISPLAY_UNITS: Record<string, number> = {
  hundreds: 100, thousands: 1e3, tenThousands: 1e4, hundredThousands: 1e5, millions: 1e6, tenMillions: 1e7, hundredMillions: 1e8, billions: 1e9, trillions: 1e12,
};

/** The divisor of an axis' Anzeigeeinheiten (1: Ohne). */
export function unitDivisor(o: AxisOptions | undefined): number {
  const u = o?.dispUnits;
  if (u === undefined) return 1;
  return typeof u === "number" ? u || 1 : DISPLAY_UNITS[u] ?? 1;
}

/**
 * A value axis' scale with the pane's fixed Begrenzungen and Einheiten: what is fixed stays, the
 * rest is worked out around it as Excel does (the unit from the range, the bounds on the unit).
 * Logarithmische Skalierung: powers of the Basis around the values (the unit is the Basis).
 */
export function fixedScale(lo: number, hi: number, o: AxisOptions | undefined, limit: number): AxisScale {
  if (o?.logBase && o.logBase > 1) return logScale(lo, hi, o);
  const auto = autoScale(o?.min ?? lo, o?.max ?? hi, limit);
  if (!o || (o.min === undefined && o.max === undefined && o.majorUnit === undefined)) return auto;
  const major = o.majorUnit ?? (o.min !== undefined || o.max !== undefined ? autoScale(o.min ?? Math.min(lo, 0), o.max ?? hi, limit).major : auto.major);
  const round = (v: number) => Number(v.toFixed(10));
  const min = o.min ?? round(Math.floor(Math.min(lo, o.max ?? lo) / major + 1e-9) * major);
  let max = o.max ?? round(Math.ceil(Math.max(hi, min) / major - 1e-9) * major);
  if (max <= min) max = round(min + major);
  return { min, max, major };
}

/** A logarithmic scale: from the power of the Basis at or below the smallest positive value to the one at or above the largest. */
export function logScale(lo: number, hi: number, o: AxisOptions): AxisScale {
  const b = o.logBase!;
  const pos = (v: number) => (v > 0 && isFinite(v) ? v : 1);
  const lg = (v: number) => Math.log(v) / Math.log(b);
  let min = o.min !== undefined && o.min > 0 ? o.min : b ** Math.floor(lg(pos(Math.min(lo, hi) > 0 ? Math.min(lo, hi) : 1)) + 1e-9);
  let max = o.max !== undefined && o.max > 0 ? o.max : b ** Math.ceil(lg(pos(Math.max(lo, hi))) - 1e-9);
  if (max <= min) max = min * b;
  min = Number(min.toPrecision(12)); max = Number(max.toPrecision(12));
  return { min, max, major: o.majorUnit && o.majorUnit > 1 ? o.majorUnit : b, log: b };
}
