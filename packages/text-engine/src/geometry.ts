// Preset shape geometry (DrawingML prstGeom) as Word draws it: the guide formulas of
// ECMA-376 (presetShapeDefinitions), evaluated for a shape's size and adjust values, give its
// paths, its text rectangle and its adjust handles (the yellow dots). The definitions are in
// geometry-presets.ts in a compact form of the standard's XML; each was checked against
// Word's PDF export (tools/feel-lab/check_shapes.py).
//
// Definition syntax, one statement per line:
//   av <name> <value>                    adjust value and its default (avLst)
//   gd <name> <op> <args...>             guide (gdLst), ops as in the standard ("*/", "+-", "pin" …)
//   hxy <adjX|-> <minX> <maxX> <adjY|-> <minY> <maxY> <x> <y>    adjust handle (ahXY)
//   hpolar <adjR|-> <minR> <maxR> <adjA|-> <minA> <maxA> <x> <y> adjust handle (ahPolar)
//   rect <l> <t> <r> <b>                 text rectangle
//   path [w=..] [h=..] [fill=none|darken|darkenLess|lighten|lightenLess] [stroke=false] : <commands>
// Commands: M x y, L x y, A wR hR stAng swAng (arcTo), Q x1 y1 x y, C x1 y1 x2 y2 x y, Z.

import { PRESET_DEFS } from "./geometry-presets";

export type PathFill = "norm" | "none" | "darken" | "darkenLess" | "lighten" | "lightenLess";

export interface GeomPath {
  /** SVG path data in the shape's own coordinates (points, origin top left). */
  readonly d: string;
  readonly fill: PathFill;
  readonly stroke: boolean;
}

export interface GeomHandle {
  /** Position of the handle in shape coordinates. */
  readonly x: number;
  readonly y: number;
  readonly kind: "xy" | "polar";
  /** The adjust values it moves and their ranges (evaluated). */
  readonly gdX?: string;
  readonly minX?: number;
  readonly maxX?: number;
  readonly gdY?: string;
  readonly minY?: number;
  readonly maxY?: number;
}

export interface Geometry {
  readonly paths: readonly GeomPath[];
  /** Text rectangle in shape coordinates. */
  readonly text: { readonly l: number; readonly t: number; readonly r: number; readonly b: number };
  readonly handles: readonly GeomHandle[];
  /** Adjust values with their defaults (in the standard's units). */
  readonly adjust: Readonly<Record<string, number>>;
}

interface PathDef {
  readonly w?: number;
  readonly h?: number;
  readonly fill: PathFill;
  readonly stroke: boolean;
  readonly cmds: readonly (readonly string[])[];
}

interface Def {
  readonly av: readonly [string, number][];
  readonly gd: readonly [string, string, readonly string[]][];
  readonly handles: readonly { kind: "xy" | "polar"; args: readonly string[] }[];
  readonly rect: readonly string[];
  readonly paths: readonly PathDef[];
}

const ANGLE = 60000; // angles are in 60000ths of a degree
const RAD = Math.PI / 180 / ANGLE;
const NUM = /^-?\d+(\.\d+)?$/;
const parsed = new Map<string, Def>();

function parse(src: string): Def {
  const av: [string, number][] = [];
  const gd: [string, string, string[]][] = [];
  const handles: { kind: "xy" | "polar"; args: string[] }[] = [];
  let rect: string[] = ["l", "t", "r", "b"];
  const paths: PathDef[] = [];
  for (const raw of src.split("\n")) {
    const line = raw.trim();
    if (!line) continue;
    const [kw, ...rest] = line.split(/\s+/);
    if (kw === "av") av.push([rest[0], Number(rest[1])]);
    else if (kw === "gd") gd.push([rest[0], rest[1], rest.slice(2)]);
    else if (kw === "hxy" || kw === "hpolar") handles.push({ kind: kw === "hxy" ? "xy" : "polar", args: rest });
    else if (kw === "rect") rect = rest;
    else if (kw === "path" || kw === "path:") {
      const colon = line.indexOf(":");
      const attrs = line.slice(4, colon).trim().split(/\s+/).filter(Boolean);
      let w: number | undefined, h: number | undefined, fill: PathFill = "norm", stroke = true;
      for (const a of attrs) {
        const [k, v] = a.split("=");
        if (k === "w") w = Number(v);
        else if (k === "h") h = Number(v);
        else if (k === "fill") fill = v as PathFill;
        else if (k === "stroke") stroke = v !== "false";
      }
      const tokens = line.slice(colon + 1).trim().split(/\s+/).filter(Boolean);
      const cmds: string[][] = [];
      for (const t of tokens) {
        if (/^[MLAQCZ]$/.test(t)) cmds.push([t]);
        else cmds[cmds.length - 1]?.push(t);
      }
      paths.push({ w, h, fill, stroke, cmds });
    }
  }
  return { av, gd, handles, rect, paths };
}

function def(prst: string): Def | null {
  let d = parsed.get(prst);
  if (d) return d;
  const src = PRESET_DEFS[prst];
  if (src === undefined) return null;
  d = parse(src);
  parsed.set(prst, d);
  return d;
}

const customDefs = new Map<string, Def>();

/** Custom geometry (a:custGeom, converted to the definition syntax) for a shape of w × h. */
export function customGeometry(src: string, w: number, h: number, adjust: Readonly<Record<string, number>> = {}): Geometry {
  let d = customDefs.get(src);
  if (!d) {
    d = parse(src);
    if (customDefs.size > 500) customDefs.clear();
    customDefs.set(src, d);
  }
  return evaluate(d, "", w, h, adjust);
}

/** The names of all preset geometries that can be drawn. */
export function presetNames(): string[] {
  return Object.keys(PRESET_DEFS);
}

export function hasPreset(prst: string): boolean {
  return prst in PRESET_DEFS;
}

/** Default adjust values of a preset (avLst). */
export function presetAdjust(prst: string): Record<string, number> {
  const d = def(prst);
  return d ? Object.fromEntries(d.av) : {};
}

function builtins(w: number, h: number): Map<string, number> {
  const ss = Math.min(w, h), ls = Math.max(w, h);
  const g = new Map<string, number>([
    ["l", 0], ["t", 0], ["r", w], ["b", h], ["w", w], ["h", h], ["hc", w / 2], ["vc", h / 2], ["ss", ss], ["ls", ls],
    ["cd2", 10800000], ["cd4", 5400000], ["cd8", 2700000], ["3cd4", 16200000], ["3cd8", 8100000], ["5cd8", 13500000], ["7cd8", 18900000],
  ]);
  for (const n of [2, 3, 4, 5, 6, 8, 10, 12, 32]) { g.set(`wd${n}`, w / n); g.set(`hd${n}`, h / n); }
  for (const n of [2, 4, 6, 8, 16, 32]) g.set(`ssd${n}`, ss / n);
  return g;
}

function formula(op: string, a: number, b: number, c: number): number {
  switch (op) {
    case "*/": return c === 0 ? 0 : (a * b) / c;
    case "+-": return a + b - c;
    case "+/": return c === 0 ? 0 : (a + b) / c;
    case "?:": return a > 0 ? b : c;
    case "abs": return Math.abs(a);
    case "at2": return Math.atan2(b, a) / RAD;
    case "cat2": return a * Math.cos(Math.atan2(c, b));
    case "sat2": return a * Math.sin(Math.atan2(c, b));
    case "cos": return a * Math.cos(b * RAD);
    case "sin": return a * Math.sin(b * RAD);
    case "tan": return a * Math.tan(b * RAD);
    case "max": return Math.max(a, b);
    case "min": return Math.min(a, b);
    case "mod": return Math.sqrt(a * a + b * b + c * c);
    case "pin": return b < a ? a : b > c ? c : b;
    case "sqrt": return Math.sqrt(Math.max(0, a));
    case "val": return a;
    default: return 0;
  }
}

/** Evaluates all guides of a preset for a shape of w × h with the given adjust values. */
function guides(d: Def, w: number, h: number, adjust: Readonly<Record<string, number>>): Map<string, number> {
  const g = builtins(w, h);
  for (const [name, value] of d.av) g.set(name, adjust[name] ?? value);
  const val = (t: string) => (NUM.test(t) ? Number(t) : g.get(t) ?? 0);
  for (const [name, op, args] of d.gd) g.set(name, formula(op, val(args[0] ?? "0"), val(args[1] ?? "0"), val(args[2] ?? "0")));
  return g;
}

function fmt(n: number): string {
  return String(Math.round(n * 1000) / 1000);
}

/**
 * The geometry of a preset for a shape of w × h (points) with adjust values (missing ones
 * take the preset's defaults); null for unknown presets.
 */
export function presetGeometry(prst: string, w: number, h: number, adjust: Readonly<Record<string, number>> = {}): Geometry | null {
  const d = def(prst);
  return d ? evaluate(d, prst, w, h, adjust) : null;
}

function evaluate(d: Def, prst: string, w: number, h: number, adjust: Readonly<Record<string, number>>): Geometry {
  let g = guides(d, w, h, adjust);
  // A speech bubble whose tail point lies inside it: Word moves the point out to the edge
  // along the line from the centre, so the tail disappears.
  if (prst.startsWith("wedge") && w > 0 && h > 0) {
    const dx = (g.get("xPos") ?? 0) - w / 2, dy = (g.get("yPos") ?? 0) - h / 2;
    const k = prst === "wedgeEllipseCallout" ? Math.hypot(dx / (w / 2), dy / (h / 2)) : Math.max(Math.abs(dx) / (w / 2), Math.abs(dy) / (h / 2));
    if (k > 0 && k < 1) g = guides(d, w, h, { ...adjust, adj1: (dx / k / w) * 100000, adj2: (dy / k / h) * 100000 });
  }
  const val = (t: string) => (NUM.test(t) ? Number(t) : g.get(t) ?? 0);
  const paths: GeomPath[] = d.paths.map((p) => {
    const sx = p.w ? w / p.w : 1, sy = p.h ? h / p.h : 1;
    const out: string[] = [];
    let cx = 0, cy = 0, startX = 0, startY = 0;
    for (const c of p.cmds) {
      const k = c[0];
      if (k === "M" || k === "L") {
        cx = val(c[1]) * sx; cy = val(c[2]) * sy;
        out.push(`${k}${fmt(cx)} ${fmt(cy)}`);
        if (k === "M") { startX = cx; startY = cy; }
      } else if (k === "Q") {
        const x1 = val(c[1]) * sx, y1 = val(c[2]) * sy;
        cx = val(c[3]) * sx; cy = val(c[4]) * sy;
        out.push(`Q${fmt(x1)} ${fmt(y1)} ${fmt(cx)} ${fmt(cy)}`);
      } else if (k === "C") {
        const x1 = val(c[1]) * sx, y1 = val(c[2]) * sy, x2 = val(c[3]) * sx, y2 = val(c[4]) * sy;
        cx = val(c[5]) * sx; cy = val(c[6]) * sy;
        out.push(`C${fmt(x1)} ${fmt(y1)} ${fmt(x2)} ${fmt(y2)} ${fmt(cx)} ${fmt(cy)}`);
      } else if (k === "A") {
        // Angles are visual in the path's own coordinates, before it is scaled to the shape.
        const [d2, ex, ey] = arcTo(cx, cy, val(c[1]), val(c[2]), val(c[3]), val(c[4]), sx, sy);
        if (d2) out.push(d2);
        cx = ex; cy = ey;
      } else if (k === "Z") {
        out.push("Z");
        cx = startX; cy = startY;
      }
    }
    return { d: out.join(" "), fill: p.fill, stroke: p.stroke };
  });
  const [l, t, r, b] = d.rect.map(val);
  const handles: GeomHandle[] = d.handles.map((hd) => {
    const a = hd.args;
    const opt = (s: string) => (s === "-" ? undefined : s);
    return {
      kind: hd.kind, x: val(a[6]), y: val(a[7]),
      gdX: opt(a[0]), minX: opt(a[0]) ? val(a[1]) : undefined, maxX: opt(a[0]) ? val(a[2]) : undefined,
      gdY: opt(a[3]), minY: opt(a[3]) ? val(a[4]) : undefined, maxY: opt(a[3]) ? val(a[5]) : undefined,
    };
  });
  return { paths, text: { l, t, r, b }, handles, adjust: Object.fromEntries(d.av) };
}

/**
 * DrawingML's arcTo as SVG: an arc of the ellipse wR × hR from the current point, starting at
 * the (visual) angle stAng and sweeping swAng. Returns the SVG commands and the end point.
 */
function arcTo(x0: number, y0: number, wR0: number, hR0: number, stAng: number, swAng: number, sx = 1, sy = 1): [string, number, number] {
  if (wR0 <= 0 || hR0 <= 0) return ["", x0, y0];
  const param = (ang: number) => Math.atan2(wR0 * Math.sin(ang * RAD), hR0 * Math.cos(ang * RAD));
  const wR = wR0 * sx, hR = hR0 * sy;
  const wrap = (x: number) => x - 2 * Math.PI * Math.round(x / (2 * Math.PI));
  const t0 = param(stAng);
  const cx = x0 - wR * Math.cos(t0), cy = y0 - hR * Math.sin(t0);
  // The end angle is visual too: the parametric sweep is the given one, corrected at both ends
  // by the difference between visual and parametric angle (which keeps whole turns).
  const sweep = swAng * RAD + wrap(param(stAng + swAng) - (stAng + swAng) * RAD) - wrap(t0 - stAng * RAD);
  const parts: string[] = [];
  const steps = Math.max(1, Math.ceil(Math.abs(sweep) / (Math.PI / 2) - 1e-9));
  let a = t0;
  let ex = x0, ey = y0;
  // Cubic Bézier segments of at most 90° each (as Word's PDF export draws arcs).
  for (let i = 0; i < steps; i++) {
    const da = sweep / steps;
    const k = (4 / 3) * Math.tan(da / 4);
    const c1x = cx + wR * (Math.cos(a) - k * Math.sin(a)), c1y = cy + hR * (Math.sin(a) + k * Math.cos(a));
    const a2 = a + da;
    ex = cx + wR * Math.cos(a2); ey = cy + hR * Math.sin(a2);
    const c2x = ex + wR * k * Math.sin(a2), c2y = ey - hR * k * Math.cos(a2);
    parts.push(`C${fmt(c1x)} ${fmt(c1y)} ${fmt(c2x)} ${fmt(c2y)} ${fmt(ex)} ${fmt(ey)}`);
    a = a2;
  }
  return [parts.join(" "), ex, ey];
}

/** Points along SVG path data made of M/L/Q/C/Z commands (for hit testing and comparison). */
export function flattenPath(d: string, step = 8): [number, number][] {
  const out: [number, number][] = [];
  const re = /([MLQCZ])([^MLQCZ]*)/g;
  let m: RegExpExecArray | null;
  let cx = 0, cy = 0, sx = 0, sy = 0;
  while ((m = re.exec(d))) {
    const n = m[2].trim().split(/\s+/).filter(Boolean).map(Number);
    if (m[1] === "M") { cx = sx = n[0]; cy = sy = n[1]; out.push([cx, cy]); }
    else if (m[1] === "L") { cx = n[0]; cy = n[1]; out.push([cx, cy]); }
    else if (m[1] === "Z") { cx = sx; cy = sy; out.push([cx, cy]); }
    else {
      const pts = m[1] === "Q" ? [[cx, cy], [n[0], n[1]], [n[2], n[3]]] : [[cx, cy], [n[0], n[1]], [n[2], n[3]], [n[4], n[5]]];
      for (let i = 1; i <= step; i++) {
        const t = i / step;
        let x = 0, y = 0;
        if (pts.length === 3) {
          x = (1 - t) ** 2 * pts[0][0] + 2 * (1 - t) * t * pts[1][0] + t * t * pts[2][0];
          y = (1 - t) ** 2 * pts[0][1] + 2 * (1 - t) * t * pts[1][1] + t * t * pts[2][1];
        } else {
          x = (1 - t) ** 3 * pts[0][0] + 3 * (1 - t) ** 2 * t * pts[1][0] + 3 * (1 - t) * t * t * pts[2][0] + t ** 3 * pts[3][0];
          y = (1 - t) ** 3 * pts[0][1] + 3 * (1 - t) ** 2 * t * pts[1][1] + 3 * (1 - t) * t * t * pts[2][1] + t ** 3 * pts[3][1];
        }
        out.push([x, y]);
      }
      const last = pts[pts.length - 1];
      cx = last[0]; cy = last[1];
    }
  }
  return out;
}
