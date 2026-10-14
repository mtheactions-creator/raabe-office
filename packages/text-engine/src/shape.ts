// Shapes as Word draws them: the preset geometry (geometry.ts) filled and outlined with the
// shape's own settings or, where it has none, its style's (resolved from the theme when the
// file was read). Shaded sub-paths (darken, lighten …) use Word's measured factors; arrowheads,
// dashes and effects follow DrawingML. The shape's text is drawn by the view, above this.

import { customGeometry, flattenPath, presetGeometry, type Geometry } from "./geometry";
import type { Fill, LineEnd, PictureFormat, PictureShadow } from "./picture";
import type { ShapeInfo } from "./model";

const SVG = "http://www.w3.org/2000/svg";
let gradientSeq = 0;

const toLinear = (c: number) => (c <= 0.04045 ? c / 12.92 : ((c + 0.055) / 1.055) ** 2.4);
const toSrgb = (c: number) => (c <= 0.0031308 ? c * 12.92 : 1.055 * Math.max(0, c) ** (1 / 2.4) - 0.055);

/** DrawingML's shade (in linear light, as Word: accent1 4472C4 at 15 % is 172C51). */
export function shadeLinear(hex: string, val: number): string {
  return [0, 2, 4].map((i) => Math.round(Math.max(0, Math.min(1, toSrgb(toLinear((parseInt(hex.slice(i, i + 2), 16) || 0) / 255) * val))) * 255).toString(16).padStart(2, "0")).join("").toUpperCase();
}

/** Fill colour of a shaded sub-path, as Word computes it (check_shapes.py, the bevel). */
export function shadeColor(hex: string, mode: string): string {
  const f = mode === "darken" ? (c: number) => c * 0.6 : mode === "darkenLess" ? (c: number) => c * 0.8
    : mode === "lighten" ? (c: number) => c + (255 - c) * 0.4 : mode === "lightenLess" ? (c: number) => c + (255 - c) * 0.2 : (c: number) => c;
  return [0, 2, 4].map((i) => Math.round(f(parseInt(hex.slice(i, i + 2), 16) || 0)).toString(16).padStart(2, "0")).join("").toUpperCase();
}

/** The fill, outline and shadow a shape is drawn with (its own, else its style's). */
export function shapeLook(f: PictureFormat | undefined, s: ShapeInfo): { fill: Fill; line: NonNullable<PictureFormat["line"]> | null; shadow?: PictureShadow } {
  const fill: Fill = f?.fill ?? s.styleFill ?? { kind: "none" };
  const own = f?.line;
  // An outline with colour "" or width -1 takes those from the style (Word's arrow lines).
  const line = own === null ? null : own
    ? { ...own, color: own.color || s.styleLine?.color || "000000", width: own.width >= 0 ? own.width : s.styleLine?.width ?? 0.75 }
    : s.styleLine ? { color: s.styleLine.color, width: s.styleLine.width } : null;
  return { fill: s.kind === "connector" ? { kind: "none" } : fill, line, shadow: f?.shadow ?? s.styleShadow };
}

/** The geometry of a shape for its size in px (null for unknown or custom geometry). */
export function shapeGeometry(f: PictureFormat | undefined, s: ShapeInfo, w: number, h: number): Geometry | null {
  if (s.custom) return customGeometry(s.custom, w, h, s.adjust ?? {});
  return presetGeometry(f?.geometry ?? "rect", w, h, s.adjust ?? {});
}

/** DrawingML's preset dashes as multiples of the line width (dash, gap …). */
const DASHES: Record<string, number[]> = {
  dot: [1, 1], sysDot: [1, 1], dash: [4, 3], sysDash: [3, 1], lgDash: [8, 3], dashDot: [4, 3, 1, 3],
  lgDashDot: [8, 3, 1, 3], lgDashDotDot: [8, 3, 1, 3, 1, 3], sysDashDot: [3, 1, 1, 1], sysDashDotDot: [3, 1, 1, 1, 1, 1],
};

const rgba = (hex: string, transparency = 0) => {
  const [r, g, b] = [0, 2, 4].map((i) => parseInt(hex.slice(i, i + 2), 16) || 0);
  return `rgba(${r},${g},${b},${Math.max(0, Math.min(1, 1 - transparency))})`;
};

function el<K extends keyof SVGElementTagNameMap>(tag: K, attrs: Record<string, string | number>): SVGElementTagNameMap[K] {
  const e = document.createElementNS(SVG, tag);
  for (const [k, v] of Object.entries(attrs)) e.setAttribute(k, String(v));
  return e;
}

/** An SVG paint for a fill in a sub-path's shading; gradients are added to `defs`. */
function paint(fill: Fill, mode: string, defs: SVGDefsElement): string {
  if (fill.kind === "none") return "none";
  if (fill.kind === "solid") return rgba(shadeColor(fill.color, mode), fill.transparency);
  if (fill.kind === "other") return rgba(shadeColor(fill.color, mode));
  const id = `raabe-grad-${++gradientSeq}`;
  const radial = fill.path === "circle" || fill.path === "shape" || fill.path === "rect";
  const g = radial ? el("radialGradient", { id, cx: "50%", cy: "50%", r: "70%" }) : (() => {
    const a = ((fill.angle ?? 90) * Math.PI) / 180;
    const dx = Math.cos(a) / 2, dy = Math.sin(a) / 2;
    return el("linearGradient", { id, x1: 0.5 - dx, y1: 0.5 - dy, x2: 0.5 + dx, y2: 0.5 + dy });
  })();
  const stops = [...fill.stops].sort((a, b) => a.pos - b.pos);
  for (const st of radial ? [...stops].reverse().map((x) => ({ ...x, pos: 1 - x.pos })) : stops) {
    g.append(el("stop", { offset: st.pos, "stop-color": `#${shadeColor(st.color, mode)}`, "stop-opacity": 1 - (st.transparency ?? 0) }));
  }
  defs.append(g);
  return `url(#${id})`;
}

/**
 * The size of an arrowhead in px for a line of width lw px: 2, 3 or 5 (klein, mittel, groß)
 * times the line width, but at least times 2 pt (measured, probe_word_arrows.py).
 */
function endSize(e: LineEnd, lw: number): { w: number; len: number } {
  const k = (v: string | undefined) => (v === "sm" ? 2 : v === "lg" ? 5 : 3);
  const base = Math.max(lw, 2 * (96 / 72));
  return { w: k(e.w) * base, len: k(e.len) * base };
}

/** An arrowhead at point (x, y) pointing in direction (dx, dy). */
function arrowHead(e: LineEnd, x: number, y: number, dx: number, dy: number, lw: number, color: string): SVGElement {
  const { w, len } = endSize(e, lw);
  const n = Math.hypot(dx, dy) || 1;
  const ux = dx / n, uy = dy / n, px = -uy, py = ux;
  const at = (a: number, b: number) => `${x + ux * a + px * b},${y + uy * a + py * b}`;
  if (e.type === "oval") return el("ellipse", { cx: x, cy: y, rx: len / 2, ry: w / 2, fill: color, transform: `rotate(${(Math.atan2(uy, ux) * 180) / Math.PI} ${x} ${y})` });
  if (e.type === "diamond") return el("polygon", { points: [at(len / 2, 0), at(0, w / 2), at(-len / 2, 0), at(0, -w / 2)].join(" "), fill: color });
  if (e.type === "arrow") return el("polyline", { points: [at(-len, w / 2), at(0, 0), at(-len, -w / 2)].join(" "), fill: "none", stroke: color, "stroke-width": lw, "stroke-linejoin": "miter" });
  if (e.type === "stealth") return el("polygon", { points: [at(0, 0), at(-len, w / 2), at(-len * 0.6, 0), at(-len, -w / 2)].join(" "), fill: color });
  return el("polygon", { points: [at(0, 0), at(-len, w / 2), at(-len, -w / 2)].join(" "), fill: color });
}

/** How far a line stops short of an arrowhead's tip (under triangles and stealth arrows). */
function setBack(e: LineEnd | undefined, lw: number): number {
  if (!e) return 0;
  const len = endSize(e, lw).len;
  return e.type === "triangle" ? len * 0.9 : e.type === "stealth" ? len * 0.5 : 0;
}

/** A path with its straight first and last segments shortened (for arrowheads). */
function shortened(d: string, start: number, end: number): string {
  const cmds = [...d.matchAll(/([MLQCZ])([^MLQCZ]*)/g)].map((m) => [m[1], ...m[2].trim().split(/\s+/).filter(Boolean).map(Number)] as [string, ...number[]]);
  const cut = (a: [number, number], b: [number, number], by: number): [number, number] => {
    const len = Math.hypot(b[0] - a[0], b[1] - a[1]);
    if (len <= by || len === 0) return a;
    return [a[0] + ((b[0] - a[0]) * by) / len, a[1] + ((b[1] - a[1]) * by) / len];
  };
  if (start > 0 && cmds.length > 1 && cmds[0][0] === "M" && cmds[1][0] === "L") {
    const [x, y] = cut([cmds[0][1], cmds[0][2]], [cmds[1][1], cmds[1][2]], start);
    cmds[0] = ["M", x, y];
  }
  const n = cmds.length - 1;
  if (end > 0 && n >= 1 && cmds[n][0] === "L") {
    const prev = cmds[n - 1];
    const [x, y] = cut([cmds[n][1], cmds[n][2]], [prev[prev.length - 2] as number, prev[prev.length - 1] as number], end);
    cmds[n] = ["L", x, y];
  }
  return cmds.map(([k, ...v]) => `${k}${v.map((x) => Math.round(x * 1000) / 1000).join(" ")}`).join(" ");
}

let filterSeq = 0;

/**
 * A shape as an element of w × h px (positioned by the caller). Its text is not included.
 */
export function shapeElement(w: number, h: number, f: PictureFormat | undefined, s: ShapeInfo): HTMLElement {
  const outer = document.createElement("div");
  outer.style.cssText = `position:absolute;width:${w}px;height:${h}px`;
  const geom = shapeGeometry(f, s, w, h);
  const { fill, line, shadow } = shapeLook(f, s);
  const svg = el("svg", { width: Math.max(w, 1), height: Math.max(h, 1), viewBox: `0 0 ${Math.max(w, 1)} ${Math.max(h, 1)}` });
  svg.style.cssText = "position:absolute;left:0;top:0;overflow:visible";
  const defs = el("defs", {});
  svg.append(defs);
  const g = el("g", {});
  svg.append(g);
  const lw = line ? line.width * (96 / 72) : 0;
  const stroke = line && lw > 0 ? rgba(line.color, line.transparency) : "none";
  const dash = line?.dash && DASHES[line.dash] ? DASHES[line.dash].map((x) => x * Math.max(lw, 0.5)).join(" ") : "";
  const cap = line?.cap === "rnd" ? "round" : line?.cap === "sq" ? "square" : (line?.dash?.startsWith("sys") ? "butt" : "butt");
  const join = line?.join ?? (line?.miter ? "miter" : "round");
  const openPath = line && (line.head || line.tail) ? geom?.paths.find((p) => p.stroke && !/Z\s*$/.test(p.d)) : undefined;
  for (const p of geom?.paths ?? [{ d: `M0 0 L${w} 0 L${w} ${h} L0 ${h} Z`, fill: "norm" as const, stroke: true }]) {
    const fp = p.fill === "none" ? "none" : paint(fill, p.fill, defs);
    const sp = p.stroke ? stroke : "none";
    if (fp === "none" && sp === "none") continue;
    const d = p === openPath ? shortened(p.d, setBack(line?.head, lw), setBack(line?.tail, lw)) : p.d;
    const attrs: Record<string, string | number> = { d, fill: fp, stroke: sp };
    if (sp !== "none") {
      attrs["stroke-width"] = lw;
      attrs["stroke-linecap"] = cap;
      attrs["stroke-linejoin"] = join;
      if (dash) attrs["stroke-dasharray"] = dash;
      // Doppelt (dbl) and the other compound lines: two thin lines.
      if (line?.compound && line.compound !== "sng") {
        g.append(el("path", { ...attrs, fill: "none", "stroke-width": lw }));
        attrs.stroke = "white";
        attrs["stroke-width"] = lw / 3;
        attrs.fill = "none";
      }
    }
    g.append(el("path", attrs));
  }
  // Arrowheads on the first open, stroked path (lines and connectors).
  if (line && lw > 0 && openPath) {
    {
      const pts = flattenPath(openPath.d, 16);
      if (pts.length >= 2) {
        const c = rgba(line.color, line.transparency);
        if (line.head) { const [a, b] = [pts[0], pts.find((q) => q[0] !== pts[0][0] || q[1] !== pts[0][1]) ?? pts[1]]; g.append(arrowHead(line.head, a[0], a[1], a[0] - b[0], a[1] - b[1], lw, c)); }
        if (line.tail) {
          const a = pts[pts.length - 1];
          const b = [...pts].reverse().find((q) => q[0] !== a[0] || q[1] !== a[1]) ?? pts[pts.length - 2];
          g.append(arrowHead(line.tail, a[0], a[1], a[0] - b[0], a[1] - b[1], lw, c));
        }
      }
    }
  }
  // Weiche Kanten, Abschrägung: SVG filters on the drawn shape.
  const prims: string[] = [];
  if (f?.softEdge && f.softEdge > 0) {
    const r = f.softEdge * (96 / 72);
    prims.push(`<feGaussianBlur in="SourceAlpha" stdDeviation="${r / 2}" result="soft"/><feComponentTransfer in="soft" result="mask"><feFuncA type="linear" slope="2" intercept="-0.6"/></feComponentTransfer><feComposite in="SourceGraphic" in2="mask" operator="in" result="edged"/>`);
  }
  if (f?.bevel || shadow?.kind === "inner") {
    const src = prims.length ? "edged" : "SourceGraphic";
    if (f?.bevel) {
      const bw = Math.max(1, f.bevel.w * (96 / 72));
      prims.push(`<feGaussianBlur in="SourceAlpha" stdDeviation="${bw / 2}" result="bb"/><feSpecularLighting in="bb" surfaceScale="${bw / 2}" specularConstant="0.6" specularExponent="12" lighting-color="#fff" result="spec"><fePointLight x="-5000" y="-8000" z="12000"/></feSpecularLighting><feComposite in="spec" in2="SourceAlpha" operator="in" result="lit"/><feComposite in="${src}" in2="lit" operator="arithmetic" k1="0" k2="1" k3="0.7" k4="0" result="beveled"/>`);
    }
    if (shadow?.kind === "inner") {
      const a = (shadow.dir * Math.PI) / 180, d = shadow.dist * (96 / 72);
      const from = f?.bevel ? "beveled" : src;
      prims.push(`<feComponentTransfer in="SourceAlpha" result="inv"><feFuncA type="table" tableValues="1 0"/></feComponentTransfer><feOffset in="inv" dx="${Math.cos(a) * d}" dy="${Math.sin(a) * d}" result="off"/><feGaussianBlur in="off" stdDeviation="${shadow.blur * (96 / 72) / 2}" result="blur"/><feFlood flood-color="${rgba(shadow.color, shadow.transparency)}"/><feComposite in2="blur" operator="in" result="sh"/><feComposite in="sh" in2="SourceAlpha" operator="in" result="ish"/><feMerge><feMergeNode in="${from}"/><feMergeNode in="ish"/></feMerge>`);
    }
  }
  if (prims.length) {
    const id = `raabe-shape-fx-${++filterSeq}`;
    defs.insertAdjacentHTML("beforeend", `<filter id="${id}" x="-20%" y="-20%" width="140%" height="140%" color-interpolation-filters="sRGB">${prims.join("")}</filter>`);
    g.setAttribute("filter", `url(#${id})`);
  }
  outer.append(svg);
  const effects: string[] = [];
  if (f?.glow && f.glow.radius > 0) {
    const c = rgba(f.glow.color, f.glow.transparency);
    effects.push(`drop-shadow(0 0 ${f.glow.radius * (96 / 72) * 0.5}px ${c})`, `drop-shadow(0 0 ${f.glow.radius * (96 / 72) * 0.5}px ${c})`);
  }
  if (shadow?.kind === "outer") {
    const a = (shadow.dir * Math.PI) / 180, d = shadow.dist * (96 / 72);
    effects.push(`drop-shadow(${Math.cos(a) * d}px ${Math.sin(a) * d}px ${shadow.blur * (96 / 72) * 0.5}px ${rgba(shadow.color, shadow.transparency)})`);
  }
  if (effects.length) svg.style.filter = effects.join(" ");
  if (f?.reflection && f.reflection.size > 0) {
    const start = 1 - f.reflection.size;
    outer.style.setProperty("-webkit-box-reflect", `below ${f.reflection.dist * (96 / 72)}px linear-gradient(transparent ${start * 100}%, rgba(255,255,255,${(1 - f.reflection.transparency) * 0.9}))`);
  }
  const tf: string[] = [];
  if (f?.rotation) tf.push(`rotate(${f.rotation}deg)`);
  if (f?.flipH) tf.push("scaleX(-1)");
  if (f?.flipV) tf.push("scaleY(-1)");
  if (tf.length) outer.style.transform = tf.join(" ");
  return outer;
}

/** The text rectangle of a shape in px (geometry's rectangle minus the insets), shape coordinates. */
export function shapeTextRect(f: PictureFormat | undefined, s: ShapeInfo, w: number, h: number, px = 96 / 72): { x: number; y: number; w: number; h: number } {
  const g = shapeGeometry(f, s, w, h);
  const r = g?.text ?? { l: 0, t: 0, r: w, b: h };
  const x = r.l + s.body.l * px, y = r.t + s.body.t * px;
  return { x, y, w: Math.max(1, r.r - s.body.r * px - x), h: Math.max(0, r.b - s.body.b * px - y) };
}

/** Word's names for new shapes (wp:docPr name before the number), by preset geometry. */
export const SHAPE_NAMES: Record<string, string> = {
  rect: "Rechteck", roundRect: "Rechteck: abgerundete Ecken", ellipse: "Ellipse", triangle: "Gleichschenkliges Dreieck", rtTriangle: "Rechtwinkliges Dreieck",
  line: "Gerader Verbinder", straightConnector1: "Gerade Verbindung mit Pfeil", bentConnector3: "Verbinder: gewinkelt", curvedConnector3: "Verbinder: gekrümmt",
  rightArrow: "Pfeil: nach rechts", leftArrow: "Pfeil: nach links", upArrow: "Pfeil: nach oben", downArrow: "Pfeil: nach unten", star5: "Stern: 5 Zacken",
  wedgeRectCallout: "Sprechblase: rechteckig", wedgeRoundRectCallout: "Sprechblase: rechteckig mit abgerundeten Ecken", wedgeEllipseCallout: "Sprechblase: oval",
  cloudCallout: "Denkblase: wolkenförmig", cloud: "Wolke", heart: "Herz", smileyFace: "Smiley", diamond: "Raute", pentagon: "Fünfeck", hexagon: "Sechseck",
};
