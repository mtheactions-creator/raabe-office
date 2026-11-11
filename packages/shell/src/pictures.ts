// The pictures of the picture dropdowns (D-161), drawn by us as SVG: line dashes, compound lines, arrowheads and
// their sizes, marker symbols, shadows, glows, soft edges, bevels, gradients, gradient directions, textures, patterns
// and cell border styles. Each returns a whole <svg> of the given size. None is copied from Office: they show the same
// thing (a dashed line, a shadow to the lower right) in our own drawing.

import type { PictureEntry } from "./picturedrop";

const INK = "#404040";

/** Office's line dropdowns with their labels (Excel, Word and PowerPoint use the same) and our pictures (D-161). */
export const DASH_TYPES: [string, string][] = [
  ["solid", "Uni"], ["sysDot", "Runder Punkt"], ["sysDash", "Viereckiger Punkt"], ["dash", "Strich"], ["dashDot", "Strichpunkt"], ["lgDash", "Langer Strich"],
  ["lgDashDot", "Langer Strichpunkt"], ["lgDashDotDot", "Langer Strichpunkt-Punkt"],
];
export const COMPOUND_TYPES: [string, string][] = [["sng", "Einfach"], ["dbl", "Doppelt"], ["thickThin", "Dick Dünn"], ["thinThick", "Dünn Dick"], ["tri", "Dreifach"]];
export const ARROW_TYPES: [string, string][] = [["none", "Kein Pfeil"], ["triangle", "Pfeil"], ["arrow", "Offener Pfeil"], ["stealth", "Rahmenpfeil"], ["diamond", "Pfeil mit Raute"], ["oval", "Ovaler Pfeil"]];
/** The sizes recorded in Excel's lists (CSS px). */
export const PIC_SIZE = { dash: [104, 20] as [number, number], compound: [72, 24] as [number, number], arrow: [33, 33] as [number, number], tile: [56, 56] as [number, number] };
export const dashEntries = (): PictureEntry[] => DASH_TYPES.map(([value, label]) => ({ value, label, svg: dashPicture(value, ...PIC_SIZE.dash) }));
export const compoundEntries = (): PictureEntry[] => COMPOUND_TYPES.map(([value, label]) => ({ value, label, svg: compoundPicture(value, ...PIC_SIZE.compound) }));
export const arrowEntries = (end: "start" | "end"): PictureEntry[] => ARROW_TYPES.map(([value, label]) => ({ value, label, svg: arrowPicture(value, end, ...PIC_SIZE.arrow) }));
/** Pfeilgröße 1-9 as "w,len" (sm/med/lg): width small to large, each with length short to long (Excel's order). */
export const arrowSizeEntries = (end: "start" | "end"): PictureEntry[] => ["sm", "med", "lg"].flatMap((w, i) => ["sm", "med", "lg"].map((len, j) => ({
  value: `${w},${len}`, label: `Pfeil ${end === "start" ? "L" : "R"} Größe ${i * 3 + j + 1}`, svg: arrowSizePicture(w, len, end, ...PIC_SIZE.arrow),
})));
let uid = 0;
const id = (p: string) => `${p}${++uid}`;
const svg = (w: number, h: number, body: string, extra = "") =>
  `<svg xmlns="http://www.w3.org/2000/svg" width="${w}" height="${h}" viewBox="0 0 ${w} ${h}" aria-hidden="true"${extra}>${body}</svg>`;

/** OOXML preset dashes as multiples of the line width (a:prstDash). */
export const DASH_PATTERN: Record<string, number[] | null> = {
  solid: null, sysDot: [1, 1], sysDash: [3, 1], dash: [4, 3], dashDot: [4, 3, 1, 3], lgDash: [8, 3], lgDashDot: [8, 3, 1, 3],
  lgDashDotDot: [8, 3, 1, 3, 1, 3], dot: [1, 3], sysDashDot: [3, 1, 1, 1], sysDashDotDot: [3, 1, 1, 1, 1, 1],
};

/** A horizontal line in a dash (Strichtyp): round dots for sysDot ("Runder Punkt"), square ones otherwise. */
export function dashPicture(dash: string, w = 96, h = 14, sw = 2): string {
  const p = DASH_PATTERN[dash] ?? null;
  const round = dash === "sysDot";
  const arr = p ? ` stroke-dasharray="${p.map((x, i) => (round && i % 2 === 0 ? 0.01 : round ? x * sw + sw : x * sw)).join(" ")}"` : "";
  return svg(w, h, `<line x1="${round ? sw / 2 + 2 : 2}" y1="${h / 2}" x2="${w - 2}" y2="${h / 2}" stroke="${INK}" stroke-width="${sw}"${arr} stroke-linecap="${round ? "round" : "butt"}"/>`);
}

/** Verbundtyp: one, two or three strokes of the line's total width. */
export function compoundPicture(kind: string, w = 96, h = 14): string {
  const bands: Record<string, [number, number][]> = {
    sng: [[3, 8]], dbl: [[3, 3], [8, 3]], thickThin: [[3, 4.5], [9.5, 1.5]], thinThick: [[3, 1.5], [6.5, 4.5]], tri: [[3, 1.5], [6.25, 1.5], [9.5, 1.5]],
  };
  const top = (h - 8) / 2;
  return svg(w, h, (bands[kind] ?? bands.sng).map(([y, t]) => `<rect x="2" y="${top + y - 3}" width="${w - 4}" height="${t}" fill="${INK}"/>`).join(""));
}

/** An arrowhead at a line's start (left) or end (right): Pfeil, Offener Pfeil, Rahmenpfeil, Raute, Oval, or none. */
export function arrowPicture(type: string, end: "start" | "end", w = 96, h = 14, scale: [number, number] = [1, 1]): string {
  const y = h / 2, len = 9 * scale[1], half = 4.5 * scale[0];
  const flip = (x: number) => (end === "end" ? w - x : x);
  const pts = (list: [number, number][]) => list.map(([x, yy]) => `${flip(x)},${yy}`).join(" ");
  const tip = 2;
  let head = "";
  let lineFrom = 2;
  switch (type) {
    case "triangle": head = `<polygon points="${pts([[tip, y], [tip + len, y - half], [tip + len, y + half]])}" fill="${INK}"/>`; lineFrom = tip + len - 1; break;
    case "arrow": head = `<polyline points="${pts([[tip + len, y - half], [tip, y], [tip + len, y + half]])}" fill="none" stroke="${INK}" stroke-width="1.5" stroke-linejoin="miter"/>`; lineFrom = tip + 1; break;
    case "stealth": head = `<polygon points="${pts([[tip, y], [tip + len, y - half], [tip + len * 0.65, y], [tip + len, y + half]])}" fill="${INK}"/>`; lineFrom = tip + len * 0.65; break;
    case "diamond": head = `<polygon points="${pts([[tip, y], [tip + len / 2, y - half], [tip + len, y], [tip + len / 2, y + half]])}" fill="${INK}"/>`; lineFrom = tip + len / 2; break;
    case "oval": head = `<ellipse cx="${flip(tip + len / 2)}" cy="${y}" rx="${len / 2}" ry="${half}" fill="${INK}"/>`; lineFrom = tip + len / 2; break;
    default: lineFrom = 2;
  }
  return svg(w, h, `<line x1="${flip(lineFrom)}" y1="${y}" x2="${flip(w - 2)}" y2="${y}" stroke="${INK}" stroke-width="1.5"/>${head}`);
}

/** Pfeilgröße: the arrowhead (Pfeil) in a width (sm, med, lg) and a length (sm, med, lg). */
export function arrowSizePicture(wid: string, len: string, end: "start" | "end", w = 40, h = 18): string {
  const f = { sm: 0.65, med: 1, lg: 1.45 } as Record<string, number>;
  return arrowPicture("triangle", end, w, h, [f[wid] ?? 1, f[len] ?? 1]);
}

/** A marker symbol (Markierungstyp). */
export function markerPicture(symbol: string, w = 24, h = 18, fill = "#4472C4"): string {
  const cx = w / 2, cy = h / 2, r = 5;
  const st = `fill="${fill}" stroke="${fill}" stroke-width="1"`;
  const line = `stroke="${fill}" stroke-width="1.6" fill="none"`;
  const body: Record<string, string> = {
    circle: `<circle cx="${cx}" cy="${cy}" r="${r}" ${st}/>`,
    square: `<rect x="${cx - r}" y="${cy - r}" width="${2 * r}" height="${2 * r}" ${st}/>`,
    diamond: `<polygon points="${cx},${cy - r - 1} ${cx + r + 1},${cy} ${cx},${cy + r + 1} ${cx - r - 1},${cy}" ${st}/>`,
    triangle: `<polygon points="${cx},${cy - r} ${cx + r},${cy + r} ${cx - r},${cy + r}" ${st}/>`,
    x: `<path d="M${cx - r} ${cy - r}L${cx + r} ${cy + r}M${cx + r} ${cy - r}L${cx - r} ${cy + r}" ${line}/>`,
    star: `<path d="M${cx - r} ${cy - r}L${cx + r} ${cy + r}M${cx + r} ${cy - r}L${cx - r} ${cy + r}M${cx} ${cy - r}V${cy + r}" ${line}/>`,
    dash: `<path d="M${cx - r} ${cy}H${cx + r}" stroke="${fill}" stroke-width="2.4"/>`,
    dot: `<circle cx="${cx}" cy="${cy}" r="2" ${st}/>`,
    plus: `<path d="M${cx - r} ${cy}H${cx + r}M${cx} ${cy - r}V${cy + r}" ${line}/>`,
  };
  return svg(w, h, body[symbol] ?? "");
}

/** The tile every effect picture is drawn on: a white square with a thin frame. */
const TILE = (x: number, y: number, s: number, extra = "") => `<rect x="${x}" y="${y}" width="${s}" height="${s}" fill="#fff" stroke="#a6a6a6" stroke-width="1"${extra}/>`;

/** Kein / Ohne: an empty tile. */
export function nonePicture(w = 40, h = 40): string {
  const s = Math.min(w, h) - 14;
  return svg(w, h, TILE((w - s) / 2, (h - s) / 2, s));
}

/** A shadow preset: outer (offset by angle), inner (along the inside edge) or perspective (a slanted shadow behind). */
export function shadowPicture(kind: "outer" | "inner" | "persp", angle: number, w = 40, h = 40, centre = false): string {
  const s = Math.min(w, h) - 16, x = (w - s) / 2, y = (h - s) / 2;
  const f = id("sh");
  const blur = `<filter id="${f}" x="-50%" y="-50%" width="200%" height="200%"><feGaussianBlur stdDeviation="1.6"/></filter>`;
  const dx = centre ? 0 : Math.round(Math.cos((angle * Math.PI) / 180) * 3), dy = centre ? 0 : Math.round(Math.sin((angle * Math.PI) / 180) * 3);
  if (kind === "outer") {
    const grow = centre ? 2 : 0;
    return svg(w, h, `<defs>${blur}</defs><rect x="${x + dx - grow}" y="${y + dy - grow}" width="${s + 2 * grow}" height="${s + 2 * grow}" fill="#000" opacity=".45" filter="url(#${f})"/>${TILE(x, y, s)}`);
  }
  if (kind === "inner") {
    const c = id("ic");
    // The frame's stroke, half inside the tile, shifted away from the lit side: the shade lies along the inner edges.
    const ix = centre ? 0 : -dx, iy = centre ? 0 : -dy;
    return svg(w, h, `<defs>${blur}<clipPath id="${c}"><rect x="${x}" y="${y}" width="${s}" height="${s}"/></clipPath></defs>${TILE(x, y, s)}`
      + `<g clip-path="url(#${c})"><rect x="${x}" y="${y}" width="${s}" height="${s}" fill="none" stroke="#000" stroke-width="${centre ? 9 : 6}" opacity=".5" filter="url(#${f})" transform="translate(${ix} ${iy})"/></g>`);
  }
  // Perspective: a flattened, slanted shadow from the bottom edge (unterhalb) or leaning back from a corner.
  const bx = x, by = y + s;
  const poly = angle === 90 ? `${bx - 2},${by + 3} ${bx + s + 2},${by + 3} ${bx + s - 2},${by + 7} ${bx + 2},${by + 7}`
    : angle === 225 || angle === 315 ? `${bx},${by} ${bx + s},${by} ${bx + s + (angle === 315 ? 9 : -9)},${by - s * 0.45} ${bx + (angle === 315 ? 9 : -9)},${by - s * 0.45}`
      : `${bx},${by} ${bx + s},${by} ${bx + s + (angle === 45 ? 9 : -9)},${by + 5} ${bx + (angle === 45 ? 9 : -9)},${by + 5}`;
  return svg(w, h, `<defs>${blur}</defs><polygon points="${poly}" fill="#000" opacity=".4" filter="url(#${f})"/>${TILE(x, y, s)}`);
}

/** A glow of a size (pt) in a colour around the tile. */
export function glowPicture(size: number, color: string, w = 40, h = 40): string {
  const s = Math.min(w, h) - 18, x = (w - s) / 2, y = (h - s) / 2;
  const f = id("gl");
  const r = Math.min(6, 1 + size / 4);
  return svg(w, h, `<defs><filter id="${f}" x="-60%" y="-60%" width="220%" height="220%"><feGaussianBlur stdDeviation="${r / 2}"/></filter></defs>`
    + `<rect x="${x - r / 2}" y="${y - r / 2}" width="${s + r}" height="${s + r}" rx="3" fill="${color}" opacity=".6" filter="url(#${f})"/>${TILE(x, y, s)}`);
}

/** Weiche Kanten of a size (pt): the tile's edges fade (the tile is filled so that the fading shows). */
export function softEdgePicture(size: number, w = 40, h = 40, fill = "#4472C4"): string {
  const s = Math.min(w, h) - 14, x = (w - s) / 2, y = (h - s) / 2;
  if (!size) return svg(w, h, `<rect x="${x}" y="${y}" width="${s}" height="${s}" fill="${fill}"/>`);
  const f = id("se");
  const blur = Math.min(5, 0.4 + size / 5);
  const m = id("sm");
  return svg(w, h, `<defs><filter id="${f}"><feGaussianBlur stdDeviation="${blur}"/></filter><mask id="${m}"><rect x="${x + blur}" y="${y + blur}" width="${s - 2 * blur}" height="${s - 2 * blur}" fill="#fff" filter="url(#${f})"/></mask></defs>`
    + `<rect x="${x}" y="${y}" width="${s}" height="${s}" fill="${fill}" mask="url(#${m})"/>`);
}

/**
 * An Abschrägung (bevel) preset: the tile's four edges lit from the upper left. The kinds differ by the width of the
 * edge, how its light runs (flat, rounded, inset) and whether a groove or a ridge runs along it.
 */
export function bevelPicture(type: string, w = 40, h = 40, fill = "#4472C4"): string {
  const s = Math.min(w, h) - 12, x = (w - s) / 2, y = (h - s) / 2;
  if (type === "none") return svg(w, h, `<rect x="${x}" y="${y}" width="${s}" height="${s}" fill="${fill}"/>`);
  const look: Record<string, { e: number; light: number; dark: number; inset?: boolean; ridge?: boolean; round?: boolean }> = {
    circle: { e: 9, light: 0.6, dark: 0.35, round: true }, relaxedInset: { e: 5, light: 0.35, dark: 0.3, inset: true }, cross: { e: 7, light: 0.5, dark: 0.4, ridge: true },
    coolSlant: { e: 10, light: 0.55, dark: 0.25 }, angle: { e: 6, light: 0.6, dark: 0.45 }, softRound: { e: 8, light: 0.4, dark: 0.3, round: true },
    convex: { e: 11, light: 0.45, dark: 0.45, round: true }, slope: { e: 6, light: 0.3, dark: 0.55 }, divot: { e: 6, light: 0.4, dark: 0.4, inset: true, ridge: true },
    riblet: { e: 4, light: 0.5, dark: 0.5, ridge: true }, hardEdge: { e: 3, light: 0.7, dark: 0.55 }, artDeco: { e: 8, light: 0.5, dark: 0.5, ridge: true, round: true },
  };
  const l = look[type] ?? look.circle;
  const e = l.e;
  const top = l.inset ? `rgba(0,0,0,${l.dark})` : `rgba(255,255,255,${l.light})`;
  const bottom = l.inset ? `rgba(255,255,255,${l.light})` : `rgba(0,0,0,${l.dark})`;
  const g = id("bv");
  const grad = (dir: string, c: string) => `<linearGradient id="${g}${dir}" ${dir === "t" ? 'x2="0" y2="1"' : dir === "b" ? 'x1="0" y1="1" x2="0" y2="0"' : dir === "l" ? 'x2="1" y2="0"' : 'x1="1" x2="0" y2="0"'}><stop offset="0" stop-color="${c}"/><stop offset="1" stop-color="${c}" stop-opacity="${l.round ? 0 : 1}"/></linearGradient>`;
  const X1 = x + s, Y1 = y + s;
  const edges = `<polygon points="${x},${y} ${X1},${y} ${X1 - e},${y + e} ${x + e},${y + e}" fill="url(#${g}t)"/>`
    + `<polygon points="${x},${y} ${x + e},${y + e} ${x + e},${Y1 - e} ${x},${Y1}" fill="url(#${g}l)"/>`
    + `<polygon points="${x},${Y1} ${x + e},${Y1 - e} ${X1 - e},${Y1 - e} ${X1},${Y1}" fill="url(#${g}b)"/>`
    + `<polygon points="${X1},${y} ${X1},${Y1} ${X1 - e},${Y1 - e} ${X1 - e},${y + e}" fill="url(#${g}r)"/>`;
  const ridge = l.ridge ? `<rect x="${x + e / 2}" y="${y + e / 2}" width="${s - e}" height="${s - e}" fill="none" stroke="rgba(255,255,255,.55)" stroke-width="1"/>` : "";
  return svg(w, h, `<defs>${grad("t", top)}${grad("l", top)}${grad("b", bottom)}${grad("r", bottom)}</defs><rect x="${x}" y="${y}" width="${s}" height="${s}" fill="${fill}"/>${edges}${ridge}`);
}

export interface GradientStopPic { pos: number; color: string; opacity?: number }

/** A gradient (Voreingestellte Farbverläufe, Richtung): linear at an angle, or radial / rectangular / path from a focus. */
export function gradientPicture(kind: string, stops: GradientStopPic[], angle = 90, focus?: { x: number; y: number }, w = 40, h = 40): string {
  const g = id("gr");
  const st = [...stops].sort((a, b) => a.pos - b.pos).map((s) => `<stop offset="${s.pos / 100}" stop-color="${s.color}"${s.opacity !== undefined ? ` stop-opacity="${s.opacity}"` : ""}/>`).join("");
  const pad = 4, s = Math.min(w, h) - 2 * pad, x = (w - s) / 2, y = (h - s) / 2;
  let def: string;
  if (kind === "linear") {
    const a = (angle * Math.PI) / 180, cx = 0.5 - Math.cos(a) / 2, cy = 0.5 - Math.sin(a) / 2;
    def = `<linearGradient id="${g}" x1="${cx}" y1="${cy}" x2="${1 - cx}" y2="${1 - cy}">${st}</linearGradient>`;
  } else {
    const fx = focus?.x ?? 0.5, fy = focus?.y ?? 0.5;
    def = `<radialGradient id="${g}" cx="${fx}" cy="${fy}" fx="${fx}" fy="${fy}" r="${kind === "rect" || kind === "path" ? 0.75 : 0.7}">${st}</radialGradient>`;
  }
  return svg(w, h, `<defs>${def}</defs><rect x="${x}" y="${y}" width="${s}" height="${s}" fill="url(#${g})"/>`);
}

/** A picture (a texture tile or any image) filling the tile. */
export function imagePicture(src: string, w = 40, h = 40): string {
  const pad = 3, s = Math.min(w, h) - 2 * pad;
  const p = id("tx");
  return svg(w, h, `<defs><pattern id="${p}" width="16" height="16" patternUnits="userSpaceOnUse"><image href="${src}" width="16" height="16"/></pattern></defs>`
    + `<rect x="${(w - s) / 2}" y="${(h - s) / 2}" width="${s}" height="${s}" fill="url(#${p})"/>`);
}

/** A cell border style (Zellen formatieren > Rahmen > Art): Excel's 13 line styles plus Keine. */
export function borderStylePicture(style: string, w = 46, h = 10): string {
  const y = h / 2;
  const line = (sw: number, dash = "", yy = y) => `<line x1="2" y1="${yy}" x2="${w - 2}" y2="${yy}" stroke="#000" stroke-width="${sw}"${dash ? ` stroke-dasharray="${dash}"` : ""}/>`;
  const map: Record<string, string> = {
    none: "", hair: line(1, "1 1"), dotted: line(1, "1 2"), dashDotDot: line(1, "6 2 2 2 2 2"), dashDot: line(1, "6 2 2 2"), dashed: line(1, "3 2"), thin: line(1),
    mediumDashDotDot: line(2, "6 2 2 2 2 2"), slantDashDot: line(2, "7 1 2 1"), mediumDashDot: line(2, "7 2 2 2"), mediumDashed: line(2, "6 3"), medium: line(2), thick: line(3),
    double: line(1, "", y - 1.5) + line(1, "", y + 1.5),
  };
  return svg(w, h, map[style] ?? "");
}

/** A pattern (Musterfüllung): `path` draws the foreground in an 8 x 8 cell, repeated. */
export function patternPicture(path: string, fg: string, bg: string, w = 20, h = 20): string {
  const p = id("pt");
  return svg(w, h, `<defs><pattern id="${p}" width="6" height="6" patternUnits="userSpaceOnUse" viewBox="0 0 6 6"><rect width="6" height="6" fill="${bg}"/><path d="${path}" fill="${fg}"/></pattern></defs><rect width="${w}" height="${h}" fill="url(#${p})"/>`);
}

/** A reflection (Spiegelung): the tile and its mirror image below, fading, at a size (%) and a distance (pt). */
export function reflectionPicture(size: number, distance: number, w = 40, h = 40): string {
  const s = 16, x = (w - s) / 2, y = 4;
  const g = id("rf");
  const rh = Math.max(2, (s * size) / 100);
  const gap = Math.min(5, distance / 2);
  return svg(w, h, `<defs><linearGradient id="${g}" x2="0" y2="1"><stop offset="0" stop-color="#4472C4" stop-opacity=".5"/><stop offset="1" stop-color="#4472C4" stop-opacity="0"/></linearGradient></defs>`
    + `<rect x="${x}" y="${y}" width="${s}" height="${s}" fill="#4472C4"/><rect x="${x}" y="${y + s + gap}" width="${s}" height="${rh}" fill="url(#${g})"/>`);
}

/** A material (3D-Format > Material): a ball whose shine and edge show the surface. */
export function materialPicture(type: string, w = 40, h = 40): string {
  const look: Record<string, [string, string, number, number]> = {
    matte: ["#9fb3d9", "#4f6fa8", 0.15, 0], warmMatte: ["#b9c6e4", "#4f6fa8", 0.25, 0], plastic: ["#ffffff", "#2f55a4", 0.6, 0.1], metal: ["#ffffff", "#203864", 0.85, 0.25],
    dkEdge: ["#8ea9db", "#0f2350", 0.3, 0.45], softEdge: ["#dde6f5", "#6f8fc8", 0.2, 0], flat: ["#4472C4", "#4472C4", 0, 0], legacyWireframe: ["none", "none", 0, 0],
    powder: ["#c9d5ee", "#5a78b6", 0.1, 0], translucentPowder: ["#e4ebf7", "#8aa5d6", 0.15, 0], clear: ["#ffffff", "#9db4e0", 0.7, 0],
  };
  const [hi, lo, shine, edge] = look[type] ?? look.warmMatte;
  if (type === "legacyWireframe") return svg(w, h, `<circle cx="${w / 2}" cy="${h / 2}" r="13" fill="none" stroke="#4472C4"/><ellipse cx="${w / 2}" cy="${h / 2}" rx="13" ry="5" fill="none" stroke="#4472C4"/><ellipse cx="${w / 2}" cy="${h / 2}" rx="5" ry="13" fill="none" stroke="#4472C4"/>`);
  const g = id("mt");
  return svg(w, h, `<defs><radialGradient id="${g}" cx=".38" cy=".35" r=".7"><stop offset="0" stop-color="${hi}"/><stop offset="${0.15 + shine * 0.2}" stop-color="${hi}" stop-opacity="${0.4 + shine * 0.6}"/><stop offset="1" stop-color="${lo}"/></radialGradient></defs>`
    + `<circle cx="${w / 2}" cy="${h / 2}" r="13" fill="url(#${g})"${edge ? ` stroke="#000" stroke-opacity="${edge}" stroke-width="1.5"` : ""}${type === "clear" ? ' fill-opacity=".55"' : ""}/>`);
}

/** A lighting rig (3D-Format > Beleuchtung): a ball lit from the rig's side, warm, cold or neutral. */
export function lightingPicture(type: string, w = 40, h = 40): string {
  const look: Record<string, [number, number, string, string]> = {
    threePt: [0.35, 0.3, "#ffffff", "#4472C4"], balanced: [0.5, 0.35, "#ffffff", "#5b7fc7"], soft: [0.45, 0.4, "#f2f5fb", "#7895cf"], harsh: [0.3, 0.25, "#ffffff", "#1f3f80"],
    flood: [0.5, 0.2, "#ffffff", "#3a63b0"], contrasting: [0.25, 0.3, "#ffffff", "#14306a"], morning: [0.3, 0.3, "#fff4d6", "#4a6fb5"], sunrise: [0.25, 0.4, "#ffe0b0", "#4a6fb5"],
    sunset: [0.75, 0.4, "#ffc9a0", "#3d5a99"], chilly: [0.35, 0.3, "#e6f3ff", "#3d6fb8"], freezing: [0.4, 0.3, "#d9f0ff", "#2c5fa8"], flat: [0.5, 0.5, "#4472C4", "#4472C4"],
    twoPt: [0.35, 0.35, "#ffffff", "#4f78c2"], glow: [0.5, 0.5, "#ffffff", "#6f93d6"], brightRoom: [0.5, 0.3, "#ffffff", "#8fa9dc"],
  };
  const [cx, cy, hi, lo] = look[type] ?? look.threePt;
  const g = id("lt");
  return svg(w, h, `<defs><radialGradient id="${g}" cx="${cx}" cy="${cy}" r=".75"><stop offset="0" stop-color="${hi}"/><stop offset="1" stop-color="${lo}"/></radialGradient></defs><circle cx="${w / 2}" cy="${h / 2}" r="13" fill="url(#${g})"/>`);
}
