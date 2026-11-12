// Designs offered in Entwurf > Designs and Datei > Neu (D-070): Office's design as PowerPoint 2013-2022 ships it
// (Calibri Light / Calibri and its colours), plus our own designs (D-078, D-190): the Raabeschule's in Raabe red
// (D-018), general ones and designs for the school's subjects and occasions. Their colours, fonts and effect style
// are shared with RaabeCalc and RaabeWrite (@raabe/shell school-designs.ts); the artwork is drawn here in code, as
// plain DrawingML shapes in the design's scheme colours on the master and on the Titelfolie and
// Abschnittsüberschrift layouts, so files open in PowerPoint unchanged and every Variante recolours the artwork.
// PowerPoint's own artwork, names and looks are not copied (D-078).
//
// Units below are pixels of a 1280 x 720 slide (16:9, 9525 EMU each). The placeholders keep Office's places, so
// artwork stays out of the text areas or is faint where it reaches into them (D-193):
//   master title 88-1192 x 38-177, body 88-1192 x 192-649, footers 667-705;
//   Titelfolie title 160-1120 x 118-369, subtitle x 378-552; Abschnitt title 87-1191 x 179-479, text x 482-640.

import { OFFICE_DESIGN, type DesignSpec, type LayoutDesign } from "@raabe/ooxml";
import { SCHOOL_DESIGNS, designVariants, type SchoolDesign } from "@raabe/shell";
import { THEME_EFFECTS } from "./themeeffects.gen";

// ----- drawing helpers ------------------------------------------------------------------------------------------

const P = 9525;
const e = (v: number) => Math.round(v * P);
const isHex = (c: string) => /^[0-9A-F]{6}$/.test(c);

interface Tone { alpha?: number; lumMod?: number; lumOff?: number }
const color = (c: string, o: Tone = {}) => {
  const tag = isHex(c) ? "srgbClr" : "schemeClr";
  const mods = (o.lumMod !== undefined ? `<a:lumMod val="${Math.round(o.lumMod * 1000)}"/>` : "") + (o.lumOff !== undefined ? `<a:lumOff val="${Math.round(o.lumOff * 1000)}"/>` : "") +
    (o.alpha !== undefined ? `<a:alpha val="${Math.round(o.alpha * 100000)}"/>` : "");
  return `<a:${tag} val="${c}">${mods}</a:${tag}>`;
};
const fill = (c: string, o?: Tone) => `<a:solidFill>${color(c, o)}</a:solidFill>`;
const NOFILL = "<a:noFill/>";
const NOLINE = "<a:ln><a:noFill/></a:ln>";
const stroke = (c: string, pt: number, o?: Tone, dash?: string) =>
  `<a:ln w="${Math.round(pt * 12700)}" cap="rnd"><a:solidFill>${color(c, o)}</a:solidFill>${dash ? `<a:prstDash val="${dash}"/>` : ""}<a:round/></a:ln>`;
const prst = (p: string, adj: Record<string, number> = {}) => `<a:prstGeom prst="${p}"><a:avLst>${Object.entries(adj).map(([k, v]) => `<a:gd name="${k}" fmla="val ${v}"/>`).join("")}</a:avLst></a:prstGeom>`;

/** Custom geometry from SVG-like path data in a w × h box (M, L, C, Q, Z; several paths, open ones with fill none). */
function path(w: number, h: number, ...paths: (string | { d: string; open?: boolean })[]): string {
  const pt = (x: string, y: string) => `<a:pt x="${Math.round(Number(x))}" y="${Math.round(Number(y))}"/>`;
  const one = (d: string) => {
    const tokens = d.trim().split(/[\s,]+/);
    const out: string[] = [];
    for (let i = 0; i < tokens.length;) {
      const k = tokens[i++];
      const n = (c: number) => tokens.slice(i, (i += c));
      if (k === "M") { const [x, y] = n(2); out.push(`<a:moveTo>${pt(x, y)}</a:moveTo>`); }
      else if (k === "L") { const [x, y] = n(2); out.push(`<a:lnTo>${pt(x, y)}</a:lnTo>`); }
      else if (k === "C") { const v = n(6); out.push(`<a:cubicBezTo>${pt(v[0], v[1])}${pt(v[2], v[3])}${pt(v[4], v[5])}</a:cubicBezTo>`); }
      else if (k === "Q") { const v = n(4); out.push(`<a:quadBezTo>${pt(v[0], v[1])}${pt(v[2], v[3])}</a:quadBezTo>`); }
      else if (k === "Z") out.push("<a:close/>");
    }
    return out.join("");
  };
  const body = paths.map((p) => typeof p === "string" ? `<a:path w="${w}" h="${h}">${one(p)}</a:path>` : `<a:path w="${w}" h="${h}"${p.open ? ' fill="none"' : ""}>${one(p.d)}</a:path>`).join("");
  return `<a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="l" t="t" r="r" b="b"/><a:pathLst>${body}</a:pathLst></a:custGeom>`;
}

interface Opt { rot?: number; flipH?: boolean; flipV?: boolean; line?: string }
/** A design shape (userDrawn, no text) at x, y, w, h px. Ids are numbered when the master or layout is written. */
function sp(name: string, x: number, y: number, w: number, h: number, geom: string, fillXml: string, o: Opt = {}): string {
  const attrs = `${o.rot ? ` rot="${Math.round(o.rot * 60000)}"` : ""}${o.flipH ? ' flipH="1"' : ""}${o.flipV ? ' flipV="1"' : ""}`;
  return `<p:sp><p:nvSpPr><p:cNvPr id="1" name="${name}"/><p:cNvSpPr/><p:nvPr userDrawn="1"/></p:nvSpPr><p:spPr><a:xfrm${attrs}><a:off x="${e(x)}" y="${e(y)}"/><a:ext cx="${Math.max(1, e(w))}" cy="${Math.max(1, e(h))}"/></a:xfrm>` +
    `${geom}${fillXml}${o.line ?? NOLINE}</p:spPr><p:txBody><a:bodyPr rtlCol="0" anchor="ctr"/><a:lstStyle/><a:p><a:pPr algn="ctr"/><a:endParaRPr lang="de-DE"/></a:p></p:txBody></p:sp>`;
}
const rect = (name: string, x: number, y: number, w: number, h: number, c: string, t?: Tone, o?: Opt) => sp(name, x, y, w, h, prst("rect"), fill(c, t), o);
const oval = (name: string, x: number, y: number, w: number, h: number, c: string, t?: Tone, o?: Opt) => sp(name, x, y, w, h, prst("ellipse"), fill(c, t), o);
const ring = (name: string, x: number, y: number, w: number, h: number, c: string, pt: number, t?: Tone, o?: Opt) => sp(name, x, y, w, h, prst("ellipse"), NOFILL, { ...o, line: stroke(c, pt, t) });
/** A polygon from absolute slide points. */
function poly(name: string, pts: [number, number][], c: string, t?: Tone, o: Opt = {}): string {
  const xs = pts.map((p) => p[0]), ys = pts.map((p) => p[1]);
  const x0 = Math.min(...xs), y0 = Math.min(...ys), w = Math.max(1, Math.max(...xs) - x0), h = Math.max(1, Math.max(...ys) - y0);
  const d = pts.map((p, i) => `${i ? "L" : "M"} ${((p[0] - x0) / w) * 1000} ${((p[1] - y0) / h) * 1000}`).join(" ") + " Z";
  return sp(name, x0, y0, w, h, path(1000, 1000, d), o.line && !c ? NOFILL : fill(c, t), o);
}
/** An open line through absolute slide points (straight segments). */
function polyline(name: string, pts: [number, number][], c: string, pt: number, t?: Tone, dash?: string): string {
  const xs = pts.map((p) => p[0]), ys = pts.map((p) => p[1]);
  const x0 = Math.min(...xs), y0 = Math.min(...ys), w = Math.max(1, Math.max(...xs) - x0), h = Math.max(1, Math.max(...ys) - y0);
  const d = pts.map((p, i) => `${i ? "L" : "M"} ${((p[0] - x0) / w) * 1000} ${((p[1] - y0) / h) * 1000}`).join(" ");
  return sp(name, x0, y0, w, h, path(1000, 1000, { d, open: true }), NOFILL, { line: stroke(c, pt, t, dash) });
}
/** An open smooth curve through absolute slide points (Catmull-Rom as cubic Béziers). */
function curve(name: string, pts: [number, number][], c: string, pt: number, t?: Tone, dash?: string): string {
  const xs = pts.map((p) => p[0]), ys = pts.map((p) => p[1]);
  const x0 = Math.min(...xs), y0 = Math.min(...ys), w = Math.max(1, Math.max(...xs) - x0), h = Math.max(1, Math.max(...ys) - y0);
  const q = pts.map(([x, y]) => [((x - x0) / w) * 1000, ((y - y0) / h) * 1000]);
  return sp(name, x0, y0, w, h, path(1000, 1000, { d: smooth(q, false), open: true }), NOFILL, { line: stroke(c, pt, t, dash) });
}

/** Catmull-Rom through points as path data (closed: back to the first point). */
function smooth(q: number[][], closed: boolean): string {
  const n = q.length;
  const at = (i: number) => (closed ? q[(i + n) % n] : q[Math.max(0, Math.min(n - 1, i))]);
  let d = `M ${q[0][0]} ${q[0][1]}`;
  for (let i = 0; i < (closed ? n : n - 1); i++) {
    const p0 = at(i - 1), p1 = at(i), p2 = at(i + 1), p3 = at(i + 2);
    d += ` C ${p1[0] + (p2[0] - p0[0]) / 6} ${p1[1] + (p2[1] - p0[1]) / 6} ${p2[0] - (p3[0] - p1[0]) / 6} ${p2[1] - (p3[1] - p1[1]) / 6} ${p2[0]} ${p2[1]}`;
  }
  return closed ? d + " Z" : d;
}

/** A small deterministic random sequence (the same artwork every time). */
function rng(seed: number): () => number {
  let a = seed >>> 0;
  return () => { a = (a + 0x6d2b79f5) >>> 0; let t = a; t = Math.imul(t ^ (t >>> 15), t | 1); t ^= t + Math.imul(t ^ (t >>> 7), t | 61); return ((t ^ (t >>> 14)) >>> 0) / 4294967296; };
}

/** A soft irregular round shape in a 1000 x 1000 box. */
function blobPath(seed: number, wobble = 0.16, n = 8): string {
  const r = rng(seed);
  const q = Array.from({ length: n }, (_, i) => { const a = (i / n) * Math.PI * 2, k = 1 - wobble + r() * wobble * 2; return [500 + Math.cos(a) * 480 * k, 500 + Math.sin(a) * 480 * k]; });
  const xs = q.map((p) => p[0]), ys = q.map((p) => p[1]);
  const x0 = Math.min(...xs), y0 = Math.min(...ys), sx = 1000 / (Math.max(...xs) - x0), sy = 1000 / (Math.max(...ys) - y0);
  return smooth(q.map(([x, y]) => [(x - x0) * sx, (y - y0) * sy]), true);
}
const blob = (name: string, x: number, y: number, w: number, h: number, seed: number, c: string, t?: Tone, o?: Opt) => sp(name, x, y, w, h, path(1000, 1000, blobPath(seed)), fill(c, t), o);
/** Concentric wobbly outlines (contour lines). */
function contours(name: string, cx: number, cy: number, rx: number, ry: number, rings: number, seed: number, c: string): string {
  let s = "";
  for (let i = 0; i < rings; i++) {
    const k = (i + 1) / rings;
    s += sp(`${name} ${i + 1}`, cx - rx * k, cy - ry * k, rx * 2 * k, ry * 2 * k, path(1000, 1000, blobPath(seed + i * 7, 0.12, 9)), NOFILL, { line: stroke(c, 1.25, { alpha: 0.35 + 0.4 * (1 - k) }) });
  }
  return s;
}
/** A smooth wave line from x0 to x1 around y. */
function wave(name: string, x0: number, x1: number, y: number, amp: number, period: number, c: string, pt: number, t?: Tone, phase = 0): string {
  const pts: [number, number][] = [];
  for (let x = x0; x <= x1 + 0.1; x += period / 8) pts.push([x, y + Math.sin(((x - x0) / period) * Math.PI * 2 + phase) * amp]);
  return curve(name, pts, c, pt, t);
}

const gradBg = (a: string, b: string, ang = 90) => `<p:bgPr><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0">${color(a)}</a:gs><a:gs pos="100000">${color(b)}</a:gs></a:gsLst><a:lin ang="${ang * 60000}" scaled="0"/></a:gradFill><a:effectLst/></p:bgPr>`;
const solidBg = (c: string, t?: Tone) => `<p:bgPr>${fill(c, t)}<a:effectLst/></p:bgPr>`;

// ----- artwork of each design -----------------------------------------------------------------------------------

interface Art { master?: string; background?: string; title?: LayoutDesign; section?: LayoutDesign }

const LEAF = "M 0 500 C 220 60 700 -20 1000 500 C 700 1020 220 940 0 500 Z";
const QUOTE = "M 300 0 C 470 0 600 130 600 300 C 600 560 420 820 120 1000 L 60 900 C 230 790 330 650 350 560 C 150 560 0 430 0 280 C 0 120 140 0 300 0 Z";

function confetti(seed: number, boxes: [number, number, number, number][], count: number): string {
  const r = rng(seed);
  let s = "";
  for (let i = 0; i < count; i++) {
    const [bx, by, bw, bh] = boxes[i % boxes.length];
    const x = bx + r() * bw, y = by + r() * bh, k = r(), acc = `accent${1 + Math.floor(r() * 6)}`;
    s += k < 0.45 ? rect(`Konfetti ${i + 1}`, x, y, 14, 6, acc, undefined, { rot: Math.round(r() * 180) }) : k < 0.8 ? oval(`Konfetti ${i + 1}`, x, y, 8, 8, acc) : sp(`Konfetti ${i + 1}`, x, y, 12, 10, prst("triangle"), fill(acc), { rot: Math.round(r() * 360) });
  }
  return s;
}

function mosaic(x: number, y: number, cols: number, rows: number, size: number, seed: number): string {
  const r = rng(seed);
  let s = "";
  for (let i = 0; i < cols; i++) for (let j = 0; j < rows; j++) {
    if (r() < 0.22) continue;
    s += rect(`Mosaik ${i + 1}-${j + 1}`, x + i * (size + 4), y + j * (size + 4), size, size, `accent${1 + Math.floor(r() * 6)}`, { alpha: 0.35 + r() * 0.6 });
  }
  return s;
}

function gridPaper(x: number, y: number, w: number, h: number, step: number, c: string, alpha: number): string {
  let d = "";
  for (let gx = 0; gx <= w + 0.1; gx += step) d += ` M ${(gx / w) * 1000} 0 L ${(gx / w) * 1000} 1000`;
  for (let gy = 0; gy <= h + 0.1; gy += step) d += ` M 0 ${(gy / h) * 1000} L 1000 ${(gy / h) * 1000}`;
  return sp("Kästchen", x, y, w, h, path(1000, 1000, { d, open: true }), NOFILL, { line: stroke(c, 0.5, { alpha }) });
}

function hexagon(name: string, cx: number, cy: number, r: number, c: string, pt: number, inner = false): string {
  return sp(name, cx - r, cy - r * 0.866, r * 2, r * 1.732, prst("hexagon", { adj: 25000, vf: 115470 }), NOFILL, { line: stroke(c, pt) }) +
    (inner ? ring(`${name} Ring`, cx - r * 0.55, cy - r * 0.55, r * 1.1, r * 1.1, c, pt * 0.8) : "");
}

function circuit(seed: number, x0: number, y0: number, w: number, h: number, lines: number, c: string): string {
  const r = rng(seed);
  let s = "";
  for (let i = 0; i < lines; i++) {
    let x = x0 + r() * w, y = y0 + r() * h;
    const pts: [number, number][] = [[x, y]];
    for (let k = 0; k < 3; k++) {
      if (k % 2 === 0) x = Math.max(x0, Math.min(x0 + w, x + (r() - 0.5) * w * 0.8)); else y = Math.max(y0, Math.min(y0 + h, y + (r() - 0.5) * h * 0.8));
      pts.push([x, y]);
    }
    s += polyline(`Leiterbahn ${i + 1}`, pts, c, 1.5, { alpha: 0.55 });
    s += oval(`Lötpunkt ${i + 1}`, pts[0][0] - 4, pts[0][1] - 4, 8, 8, c, { alpha: 0.8 }) + ring(`Anschluss ${i + 1}`, x - 5, y - 5, 10, 10, c, 1.5, { alpha: 0.8 });
  }
  return s;
}

function staff(x0: number, x1: number, y: number, gap: number, c: string, alpha: number, bend = 0): string {
  let s = "";
  for (let i = 0; i < 5; i++) {
    const yy = y + i * gap;
    s += bend ? curve(`Notenlinie ${i + 1}`, [[x0, yy + bend], [(x0 * 2 + x1) / 3, yy - bend * 0.6], [(x0 + x1 * 2) / 3, yy + bend * 0.4], [x1, yy - bend]], c, 1, { alpha }) : rect(`Notenlinie ${i + 1}`, x0, yy, x1 - x0, 1, c, { alpha });
  }
  return s;
}
function note(name: string, x: number, y: number, size: number, c: string, up = true): string {
  return oval(name, x, y, size * 1.35, size, c, undefined, { rot: -20 }) + rect(`${name} Hals`, up ? x + size * 1.2 : x + size * 0.05, up ? y - size * 3 : y + size * 0.5, Math.max(1.5, size * 0.14), size * 3, c);
}

function laurel(cx: number, cy: number, c: string, side: 1 | -1, n = 7): string {
  let s = "";
  for (let i = 0; i < n; i++) {
    const a = (Math.PI * (0.15 + (0.6 * i) / n));
    const x = cx + side * Math.cos(a) * 90, y = cy - Math.sin(a) * 60 + 30;
    s += sp(`Blatt ${side > 0 ? "rechts" : "links"} ${i + 1}`, x - 14, y - 6, 28, 12, path(1000, 1000, LEAF), fill(c), { rot: side > 0 ? -((a * 180) / Math.PI) + 120 : (a * 180) / Math.PI + 60 });
  }
  return s;
}

function columnShape(x: number, y: number, w: number, h: number, c: string): string {
  return rect("Kapitell", x - w * 0.12, y, w * 1.24, h * 0.05, c, { alpha: 0.45 }) + rect("Kapitell 2", x - w * 0.04, y + h * 0.05, w * 1.08, h * 0.03, c, { alpha: 0.35 }) +
    rect("Säule", x, y + h * 0.08, w, h * 0.84, c, { alpha: 0.25 }) +
    [0.25, 0.5, 0.75].map((k, i) => rect(`Kannelur ${i + 1}`, x + w * k - 1, y + h * 0.1, 2, h * 0.8, c, { alpha: 0.35 })).join("") +
    rect("Basis", x - w * 0.12, y + h * 0.92, w * 1.24, h * 0.08, c, { alpha: 0.45 });
}

function meander(x0: number, x1: number, y: number, c: string, alpha: number): string {
  const units = Math.floor((x1 - x0) / 20);
  const w = units * 20, h = 12;
  let d = "";
  for (let i = 0; i < units; i++) {
    const x = i * 20;
    d += ` M ${x} ${h} L ${x} 0 L ${x + 15} 0 L ${x + 15} 8 L ${x + 5} 8 L ${x + 5} 4 L ${x + 10} 4`;
  }
  d += ` M 0 ${h} L ${w} ${h}`;
  return sp("Mäander", x0, y, w, h, path(w, h, { d, open: true }), NOFILL, { line: stroke(c, 1.25, { alpha }) });
}

function rays(cx: number, cy: number, n: number, from: number, to: number, len: number, c: string, alpha: number): string {
  let s = "";
  for (let i = 0; i < n; i++) {
    const a = ((from + ((to - from) * i) / (n - 1)) * Math.PI) / 180, da = 0.035;
    s += poly(`Strahl ${i + 1}`, [[cx, cy], [cx + Math.cos(a - da) * len, cy + Math.sin(a - da) * len], [cx + Math.cos(a + da) * len, cy + Math.sin(a + da) * len]], c, { alpha });
  }
  return s;
}

function mountains(y: number, h: number, c1: string, c2: string, a1: number, a2: number): string {
  return poly("Berge hinten", [[0, y + h], [0, y + h * 0.55], [180, y + h * 0.15], [330, y + h * 0.5], [520, y], [700, y + h * 0.45], [860, y + h * 0.2], [1060, y + h * 0.55], [1180, y + h * 0.3], [1280, y + h * 0.5], [1280, y + h]], c1, { alpha: a1 }) +
    poly("Berge vorne", [[0, y + h], [0, y + h * 0.75], [140, y + h * 0.5], [290, y + h * 0.8], [440, y + h * 0.45], [640, y + h * 0.85], [820, y + h * 0.55], [990, y + h * 0.8], [1150, y + h * 0.6], [1280, y + h * 0.75], [1280, y + h]], c2, { alpha: a2 });
}

function arches(x: number, y: number, n: number, w: number, h: number, gap: number): string {
  let s = "";
  for (let i = 0; i < n; i++) s += sp(`Bogen ${i + 1}`, x + i * (w + gap), y, w, h, prst("round2SameRect", { adj1: 50000, adj2: 0 }), fill(`accent${(i % 6) + 1}`, { alpha: 0.85 }));
  return s;
}

function blocks(seed: number, x: number, y: number, n: number, size: number): string {
  const r = rng(seed);
  let s = "";
  for (let i = 0; i < n; i++) s += sp(`Baustein ${i + 1}`, x + (i % 4) * (size + 6) + (Math.floor(i / 4) % 2) * size * 0.5, y - Math.floor(i / 4) * (size + 6), size, size, prst("roundRect", { adj: 18000 }), fill(`accent${1 + Math.floor(r() * 6)}`, { alpha: 0.9 }), { rot: Math.round((r() - 0.5) * 16) });
  return s;
}

const ART: Record<string, Art> = {
  Raabeschule: {
    master: rect("Raabe-Band", 0, 0, 1280, 24, "accent1") + rect("Raabe-Linie", 88, 657.6, 1104, 1.33, "accent2"),
    title: { deco: rect("Titelstrich", 580, 371, 120, 5, "accent1") },
    section: { deco: rect("Abschnittsbalken", 56, 300, 10, 175, "accent1") },
  },
  "Raabeschule Band": {
    master: rect("Seitenband", 0, 0, 28, 720, "accent1") + rect("Seitenband dunkel", 28, 0, 6, 720, "tx2"),
    title: { hideMaster: true, background: solidBg("bg2"), colors: "dark",
      deco: poly("Schräge", [[1130, 0], [1280, 0], [1280, 720], [1010, 720]], "accent1") + poly("Schräge Licht", [[1100, 0], [1130, 0], [1010, 720], [980, 720]], "tx1", { alpha: 0.18 }) },
    section: { hideMaster: true, background: solidBg("bg2"), colors: "dark", deco: rect("Abschnittsbalken", 50, 300, 10, 330, "accent1") + rect("Seitenband", 0, 0, 28, 720, "accent1") },
  },
  Klar: {
    master: rect("Seitenleiste", 0, 0, 24, 720, "accent1"),
    title: { deco: rect("Fußfläche", 0, 600, 1280, 120, "bg2") + rect("Fußlinie", 0, 600, 1280, 4, "accent2") },
    section: { deco: rect("Abschnittsstrich", 98, 160, 80, 6, "accent2") },
  },
  Tafel: {
    master: sp("Tafelrahmen", 16, 16, 1248, 688, prst("rect"), NOFILL, { line: stroke("tx1", 1.5, { alpha: 0.22 }) }),
    title: { deco: curve("Kreidestrich", [[460, 374], [540, 369], [620, 376], [700, 368], [780, 375], [820, 371]], "accent1", 3) },
    section: { deco: curve("Kreidestrich", [[98, 484], [160, 479], [220, 486], [290, 480]], "accent1", 3) },
  },
  Wellenband: {
    master: sp("Welle hinten", 0, 556.4, 1280, 163.6, path(1000, 1000, "M 0 420 C 250 120 500 820 750 360 C 870 140 950 260 1000 300 L 1000 1000 L 0 1000 Z"), fill("accent2", { alpha: 0.55 })) +
      sp("Welle vorne", 0, 593.2, 1280, 126.8, path(1000, 1000, "M 0 500 C 220 900 480 120 720 520 C 850 740 940 560 1000 430 L 1000 1000 L 0 1000 Z"), fill("accent1")),
    title: { deco: sp("Welle oben", 0, 0, 1280, 92, path(1000, 1000, "M 0 0 L 1000 0 L 1000 520 C 760 900 520 220 280 640 C 160 860 60 700 0 560 Z"), fill("accent2", { alpha: 0.3 })) },
  },
  "Kreisbögen": {
    master: oval("Kreis groß", 976.4, -220.5, 482.9, 482.9, "accent1", { alpha: 0.85 }) + oval("Kreis mittel", 1112.9, 94.5, 252, 252, "accent2", { alpha: 0.8 }) + oval("Kreis klein", -73.5, 577.4, 199.5, 199.5, "accent4", { alpha: 0.7 }),
    title: { deco: ring("Ring", 40, 470, 150, 150, "accent3", 3) },
  },
  Keil: {
    master: sp("Keil dunkel", 0, 0, 80, 720, path(1000, 1000, "M 0 0 L 1000 0 L 420 1000 L 0 1000 Z"), fill("accent1")) +
      sp("Keil hell", 0, 0, 120, 720, path(1000, 1000, "M 600 0 L 760 0 L 300 1000 L 140 1000 Z"), fill("accent2", { alpha: 0.7 })),
    title: { deco: poly("Keil rechts", [[1060, 0], [1280, 0], [1280, 720], [1190, 720]], "accent1") + poly("Keil rechts hell", [[1010, 0], [1050, 0], [1180, 720], [1150, 720]], "accent2", { alpha: 0.7 }) },
  },
  Punkteraster: {
    master: Array.from({ length: 6 }, (_, r) => Array.from({ length: 6 - r }, (_, c) => oval(`Punkt ${r}-${c}`, 1240 - c * 31.5, 24.15 + r * 31.5, 12.6, 12.6, r % 2 ? "accent2" : "accent1")).join("")).join("") + rect("Linie", 88, 657.6, 1104, 1.33, "accent1"),
    title: { deco: Array.from({ length: 4 }, (_, r) => Array.from({ length: 6 - r }, (_, c) => oval(`Feld ${r}-${c}`, 28 + c * 26, 690 - r * 26, 10, 10, "accent2", { alpha: 0.7 })).join("")).join("") },
  },
  "Schräge Streifen": {
    master: Array.from({ length: 5 }, (_, i) => sp(`Streifen ${i + 1}`, 776.9 + i * 54.6, 0, 94.5, 67, path(1000, 1000, "M 300 0 L 620 0 L 320 1000 L 0 1000 Z"), fill(i % 2 ? "accent2" : "accent1", i % 2 ? { alpha: 0.8 } : undefined))).join("") +
      rect("Fußlinie", 0, 708.7, 1280, 11.3, "accent1"),
  },
  Abendblau: {
    background: gradBg("bg2", "bg1"),
    master: sp("Horizont", 0, 651, 1280, 69, path(1000, 1000, "M 0 600 C 300 200 700 200 1000 600 L 1000 1000 L 0 1000 Z"), fill("accent1", { alpha: 0.25 })),
    title: { deco: (() => { const r = rng(7); return Array.from({ length: 16 }, (_, i) => { const s = 3 + r() * 4; return oval(`Stern ${i + 1}`, r() < 0.5 ? r() * 150 : 1130 + r() * 140, 20 + r() * 520, s, s, "tx1", { alpha: 0.45 + r() * 0.4 }); }).join(""); })() },
  },
  Stille: {
    master: rect("Linie", 55, 70, 2, 580, "accent1", { alpha: 0.6 }) + oval("Punkt", 51, 52, 10, 10, "accent1"),
    title: { deco: oval("Kreis", 960, 420, 420, 420, "accent2", { alpha: 0.25 }) + oval("Kreis klein", 1030, 330, 70, 70, "accent1", { alpha: 0.5 }) },
    section: { deco: oval("Kreis", -160, -200, 460, 460, "accent2", { alpha: 0.2 }) },
  },
  Papier: {
    master: rect("Linie", 88, 180, 1104, 1.5, "accent1") + rect("Linie fein", 88, 184, 1104, 0.75, "accent1", { alpha: 0.5 }) +
      poly("Eselsohr", [[1226, 720], [1280, 666], [1280, 720]], "bg2", { lumMod: 85 }) + poly("Eselsohr Schatten", [[1226, 720], [1226, 680], [1280, 666]], "tx1", { alpha: 0.12 }),
    title: { hideMaster: true, deco: rect("Zierlinie links", 470, 373, 150, 1.25, "accent1") + sp("Raute", 632, 366, 16, 16, prst("diamond"), fill("accent1")) + rect("Zierlinie rechts", 660, 373, 150, 1.25, "accent1") },
    section: { hideMaster: true, deco: sp("Anführung", 40, 150, 50, 60, path(600, 1000, QUOTE), fill("accent1", { alpha: 0.35 }), { rot: 180 }) },
  },
  Kontrast: {
    master: rect("Block", 0, 0, 16, 720, "accent1") + rect("Titelstrich", 98, 178, 140, 8, "accent1"),
    title: { hideMaster: true, deco: poly("Block", [[0, 0], [150, 0], [105, 720], [0, 720]], "accent1") + poly("Streifen", [[165, 0], [185, 0], [140, 720], [120, 720]], "accent2") },
    section: { deco: rect("Band", 0, 0, 1280, 150, "accent1") + rect("Streifen", 0, 150, 1280, 8, "accent2") },
  },
  Minimal: {
    master: rect("Akzent", 98, 30, 48, 5, "accent1"),
    title: { hideMaster: true, deco: rect("Akzent", 616, 372, 48, 4, "accent1") },
    section: { deco: rect("Akzent", 98, 480, 48, 4, "accent1") },
  },
  Mitternacht: {
    background: gradBg("bg2", "bg1"),
    master: (() => {
      const pts: [number, number][] = [[1150, 30], [1196, 64], [1238, 38], [1226, 104], [1178, 112]];
      return polyline("Sternbild", [...pts, pts[1]], "tx2", 0.75, { alpha: 0.45 }) + pts.map((p, i) => oval(`Stern ${i + 1}`, p[0] - 3.5, p[1] - 3.5, 7, 7, "accent1")).join("");
    })(),
    title: { deco: oval("Schimmer", -200, 600, 1680, 400, "accent1", { alpha: 0.12 }) + (() => {
      const pts: [number, number][] = [[60, 80], [120, 140], [96, 230], [170, 260], [210, 190]];
      return polyline("Sternbild 2", pts, "tx2", 0.75, { alpha: 0.45 }) + pts.map((p, i) => oval(`Stern links ${i + 1}`, p[0] - 4, p[1] - 4, 8, 8, i % 2 ? "accent2" : "accent1")).join("");
    })() },
  },
  Pastell: {
    master: blob("Fleck oben", 1090, -70, 260, 200, 11, "accent1", { alpha: 0.6 }) + blob("Fleck unten", -70, 610, 250, 170, 23, "accent3", { alpha: 0.6 }) + blob("Fleck klein", 1200, 120, 60, 50, 5, "accent2", { alpha: 0.7 }),
    title: { hideMaster: true, deco: blob("Fleck 1", -120, -90, 420, 320, 31, "accent1", { alpha: 0.55 }) + blob("Fleck 2", 1000, 470, 380, 320, 41, "accent2", { alpha: 0.6 }) + blob("Fleck 3", 1080, 40, 160, 130, 51, "accent3", { alpha: 0.7 }) + blob("Fleck 4", 60, 560, 130, 110, 61, "accent4", { alpha: 0.7 }) },
    section: { deco: blob("Fleck", -110, 420, 330, 360, 71, "accent5", { alpha: 0.55 }) },
  },
  Konfetti: {
    master: confetti(3, [[1080, 6, 190, 26], [10, 690, 260, 24]], 22),
    title: { hideMaster: true, deco: confetti(5, [[20, 10, 1240, 90], [20, 575, 1240, 80], [10, 110, 130, 460], [1135, 110, 135, 460]], 70) },
    section: { deco: confetti(9, [[20, 10, 1240, 150]], 34) },
  },
  Mosaik: {
    master: mosaic(1204, 4, 2, 18, 36, 13),
    title: { hideMaster: true, deco: mosaic(4, 4, 3, 18, 36, 17) + mosaic(1164, 4, 3, 18, 36, 19) },
    section: { deco: mosaic(4, 4, 32, 4, 36, 23) },
  },
  Rahmen: {
    master: sp("Rahmen", 14, 14, 1252, 692, prst("rect"), NOFILL, { line: stroke("tx2", 1) }) + sp("Rahmen innen", 20, 20, 1240, 680, prst("rect"), NOFILL, { line: stroke("accent2", 0.5) }) +
      [[10, 10], [1260, 10], [10, 700], [1260, 700]].map(([x, y], i) => rect(`Ecke ${i + 1}`, x, y, 10, 10, "accent2")).join(""),
    title: { deco: rect("Zierlinie links", 480, 373, 140, 1, "accent2") + sp("Raute", 633, 367, 14, 14, prst("diamond"), fill("accent2")) + rect("Zierlinie rechts", 660, 373, 140, 1, "accent2") },
  },
  Grundformen: {
    master: sp("Dreieck", 1158, 22, 30, 28, prst("triangle"), fill("accent3")) + oval("Kreis", 1194, 14, 40, 40, "accent1") + rect("Quadrat", 1240, 22, 28, 28, "accent2"),
    title: { hideMaster: true, deco: oval("Kreis", 1100, 470, 260, 260, "accent1") + rect("Quadrat", 1150, 110, 130, 130, "accent2", undefined, { rot: 12 }) + sp("Dreieck", 960, 570, 150, 130, prst("triangle"), fill("accent3")) + rect("Linie", 0, 640, 1280, 2, "tx1") },
    section: { deco: oval("Kreis", 1110, -90, 240, 240, "accent1") + rect("Quadrat", 1040, 40, 70, 70, "accent2", undefined, { rot: 12 }) },
  },
  Horizont: {
    master: oval("Sonne", 1180, 640, 50, 50, "accent2", { alpha: 0.8 }) +
      sp("Hügel hinten", 0, 660, 1280, 60, path(1000, 1000, "M 0 700 C 250 0 500 0 700 500 C 820 800 930 300 1000 200 L 1000 1000 L 0 1000 Z"), fill("accent1", { alpha: 0.3 })) +
      sp("Hügel vorne", 0, 684, 1280, 36, path(1000, 1000, "M 0 300 C 200 900 450 0 650 400 C 800 700 900 100 1000 500 L 1000 1000 L 0 1000 Z"), fill("accent5", { alpha: 0.45 })),
    title: { hideMaster: true, background: gradBg("bg2", "bg1"), deco: oval("Sonne", 1100, 500, 120, 120, "accent2", { alpha: 0.85 }) +
      sp("Hügel hinten", 0, 560, 1280, 160, path(1000, 1000, "M 0 700 C 250 0 500 0 700 500 C 820 800 930 300 1000 200 L 1000 1000 L 0 1000 Z"), fill("accent1", { alpha: 0.35 })) +
      sp("Hügel vorne", 0, 620, 1280, 100, path(1000, 1000, "M 0 300 C 200 900 450 0 650 400 C 800 700 900 100 1000 500 L 1000 1000 L 0 1000 Z"), fill("accent5", { alpha: 0.5 })) },
  },
  Eckpunkt: {
    master: poly("Ecke oben", [[0, 0], [70, 0], [0, 70]], "accent1") + poly("Ecke unten", [[1280, 720], [1210, 720], [1280, 650]], "accent2", { alpha: 0.8 }),
    title: { hideMaster: true, deco: poly("Ecke oben", [[0, 0], [260, 0], [0, 260]], "accent1") + poly("Ecke oben dunkel", [[0, 0], [170, 0], [0, 170]], "tx2") +
      poly("Ecke unten", [[1280, 720], [1020, 720], [1280, 460]], "accent1", { alpha: 0.85 }) + poly("Ecke unten hell", [[1280, 720], [1130, 720], [1280, 570]], "accent2") },
    section: { deco: poly("Ecke oben", [[0, 0], [200, 0], [0, 200]], "accent1") },
  },

  // ----- Fächer -----
  Mathematik: {
    master: gridPaper(1208, 0, 72, 720, 12, "accent1", 0.3) + rect("Rand", 1206, 0, 2, 720, "accent1", { alpha: 0.6 }),
    title: { deco: gridPaper(0, 0, 1280, 720, 40, "accent1", 0.12) +
      polyline("x-Achse", [[40, 650], [430, 650]], "tx1", 1.5) + polyline("y-Achse", [[80, 690], [80, 410]], "tx1", 1.5) +
      sp("Pfeil x", 428, 645, 12, 10, prst("triangle"), fill("tx1"), { rot: 90 }) + sp("Pfeil y", 75, 400, 10, 12, prst("triangle"), fill("tx1")) +
      curve("Parabel", [[110, 430], [150, 540], [200, 600], [250, 615], [300, 600], [350, 540], [390, 430]], "accent2", 2.5) +
      wave("Sinus", 880, 1260, 60, 28, 190, "accent1", 2) },
    section: { deco: gridPaper(0, 0, 1280, 150, 30, "accent1", 0.18) },
  },
  Deutsch: {
    master: rect("Heftrand", 68, 0, 2, 720, "accent1", { alpha: 0.5 }) +
      sp("Anführung 1", 1186, 18, 30, 50, path(600, 1000, QUOTE), fill("accent1", { alpha: 0.3 }), { rot: 180 }) + sp("Anführung 2", 1222, 18, 30, 50, path(600, 1000, QUOTE), fill("accent1", { alpha: 0.3 }), { rot: 180 }),
    title: { hideMaster: true, deco: Array.from({ length: 6 }, (_, i) => rect(`Zeile ${i + 1}`, 0, 580 + i * 22, 1280, 1, "accent4", { alpha: 0.3 })).join("") + rect("Heftrand", 90, 0, 2, 720, "accent1", { alpha: 0.5 }) +
      sp("Anführung 1", 30, 120, 45, 75, path(600, 1000, QUOTE), fill("accent1", { alpha: 0.35 }), { rot: 180 }) + sp("Anführung 2", 84, 120, 45, 75, path(600, 1000, QUOTE), fill("accent1", { alpha: 0.35 }), { rot: 180 }) },
    section: { deco: sp("Anführung 1", 92, 60, 45, 75, path(600, 1000, QUOTE), fill("accent1", { alpha: 0.35 }), { rot: 180 }) + sp("Anführung 2", 146, 60, 45, 75, path(600, 1000, QUOTE), fill("accent1", { alpha: 0.35 }), { rot: 180 }) },
  },
  Englisch: {
    master: sp("Sprechblase 1", 1166, 12, 64, 40, prst("wedgeRoundRectCallout", { adj1: -20833, adj2: 62500, adj3: 16667 }), fill("accent1", { alpha: 0.9 })) +
      sp("Sprechblase 2", 1214, 44, 50, 32, prst("wedgeRoundRectCallout", { adj1: 20833, adj2: 62500, adj3: 16667 }), fill("accent2", { alpha: 0.85 })),
    title: { deco: sp("Sprechblase 1", 1010, 14, 200, 96, prst("wedgeRoundRectCallout", { adj1: -25000, adj2: 66000, adj3: 16667 }), fill("accent1", { alpha: 0.9 })) +
      sp("Sprechblase 2", 1130, 60, 130, 70, prst("wedgeRoundRectCallout", { adj1: 25000, adj2: 66000, adj3: 16667 }), fill("accent2", { alpha: 0.85 })) +
      sp("Sprechblase 3", 40, 590, 150, 80, prst("wedgeRoundRectCallout", { adj1: 25000, adj2: -66000, adj3: 16667 }), NOFILL, { line: stroke("accent3", 2.5) }) },
    section: { deco: sp("Sprechblase", 98, 40, 170, 90, prst("wedgeRoundRectCallout", { adj1: -25000, adj2: 66000, adj3: 16667 }), fill("accent1", { alpha: 0.9 })) +
      sp("Sprechblase 2", 230, 80, 110, 60, prst("wedgeRoundRectCallout", { adj1: 25000, adj2: 66000, adj3: 16667 }), fill("accent2", { alpha: 0.85 })) },
  },
  "Französisch": {
    master: rect("Streifen 1", 0, 0, 8, 720, "accent1") + rect("Streifen 2", 8, 0, 8, 720, "accent3") + rect("Streifen 3", 16, 0, 8, 720, "accent2"),
    title: { deco: rect("Streifen 1", 0, 0, 24, 720, "accent1") + rect("Streifen 2", 24, 0, 24, 720, "accent3") + rect("Streifen 3", 48, 0, 24, 720, "accent2") +
      sp("Sprechblase", 1050, 20, 190, 105, prst("wedgeEllipseCallout", { adj1: -30000, adj2: 70000 }), fill("accent1", { alpha: 0.85 })) +
      sp("Sprechblase 2", 1040, 590, 140, 76, prst("wedgeEllipseCallout", { adj1: 30000, adj2: -70000 }), fill("accent2", { alpha: 0.8 })) },
    section: { deco: rect("Band 1", 0, 0, 1280, 10, "accent1") + rect("Band 2", 0, 10, 1280, 10, "accent3") + rect("Band 3", 0, 20, 1280, 10, "accent2") },
  },
  Latein: {
    master: meander(0, 1280, 706, "accent2", 0.8),
    title: { hideMaster: true, deco: columnShape(40, 120, 70, 480, "accent2") + columnShape(1170, 120, 70, 480, "accent2") + meander(0, 1280, 640, "accent2", 0.8) + meander(0, 1280, 20, "accent2", 0.8) },
    section: { deco: meander(0, 1280, 120, "accent2", 0.8) },
  },
  Geschichte: {
    master: rect("Zeitstrahl", 0, 712, 1280, 2, "accent1", { alpha: 0.7 }) + Array.from({ length: 8 }, (_, i) => oval(`Ereignis ${i + 1}`, 80 + i * 160 - 5, 708, 10, 10, i === 5 ? "accent3" : "accent1")).join(""),
    title: { hideMaster: true, deco: rect("Zeitstrahl", 40, 612, 1180, 3, "accent1") + sp("Pfeil", 1214, 603, 22, 22, prst("triangle"), fill("accent1"), { rot: 90 }) +
      Array.from({ length: 11 }, (_, i) => rect(`Strich ${i + 1}`, 80 + i * 105, i % 5 === 0 ? 596 : 604, 2, i % 5 === 0 ? 34 : 18, "accent1")).join("") +
      oval("Ereignis 1", 285 - 9, 604, 18, 18, "accent3") + oval("Ereignis 2", 705 - 9, 604, 18, 18, "accent2") + oval("Ereignis 3", 915 - 9, 604, 18, 18, "accent3") },
    section: { deco: rect("Zeitstrahl", 50, 60, 3, 600, "accent1") + Array.from({ length: 5 }, (_, i) => oval(`Ereignis ${i + 1}`, 44.5, 90 + i * 130, 14, 14, i === 2 ? "accent3" : "accent1")).join("") },
  },
  Erdkunde: {
    master: contours("Höhenlinie", 1200, 40, 130, 100, 4, 3, "accent1"),
    title: { deco: contours("Höhenlinie", 1170, 80, 190, 150, 6, 5, "accent1") + contours("Höhenlinie unten", 90, 650, 230, 160, 6, 17, "accent2") +
      sp("Nordpfeil", 60, 40, 30, 46, path(1000, 1000, "M 500 0 L 1000 1000 L 500 760 L 0 1000 Z"), fill("tx2")) },
    section: { deco: contours("Höhenlinie", 1130, 90, 220, 150, 6, 29, "accent1") },
  },
  "Politik-Wirtschaft": {
    master: [14, 24, 20, 34].map((hh, i) => rect(`Säule ${i + 1}`, 1170 + i * 22, 56 - hh, 16, hh, i === 3 ? "accent2" : "accent1", { alpha: 0.85 })).join("") + polyline("Trend", [[1172, 40], [1194, 30], [1216, 34], [1240, 14]], "accent2", 1.5),
    title: { deco: (() => {
      let s = "", k = 0;
      for (let row = 0; row < 4; row++) {
        const rr = 70 + row * 28, n = 7 + row * 3;
        for (let i = 0; i < n; i++) {
          const a = Math.PI - (Math.PI * i) / (n - 1), share = i / (n - 1);
          s += oval(`Sitz ${++k}`, 1140 + Math.cos(a) * rr - 6, 700 - Math.sin(a) * rr - 6, 12, 12, share < 0.3 ? "accent1" : share < 0.55 ? "accent3" : share < 0.8 ? "accent2" : "accent4");
        }
      }
      return s;
    })() + [40, 70, 55, 95].map((hh, i) => rect(`Säule ${i + 1}`, 40 + i * 40, 180 - hh, 30, hh, i === 3 ? "accent2" : "accent1", { alpha: 0.8 })).join("") },
    section: { deco: [40, 70, 55, 95, 120].map((hh, i) => rect(`Säule ${i + 1}`, 1050 + i * 40, 150 - hh, 30, hh, i === 4 ? "accent2" : "accent1", { alpha: 0.8 })).join("") },
  },
  Biologie: {
    master: sp("Blatt 1", 1180, 18, 76, 34, path(1000, 1000, LEAF), fill("accent1"), { rot: -30 }) + sp("Blatt 2", 1150, 46, 60, 26, path(1000, 1000, LEAF), fill("accent2"), { rot: 20 }),
    title: { deco: sp("Blatt 1", 20, 560, 180, 80, path(1000, 1000, LEAF), fill("accent1"), { rot: -40 }) + sp("Blatt 2", 110, 610, 150, 66, path(1000, 1000, LEAF), fill("accent2"), { rot: -10 }) +
      sp("Blatt 3", -20, 480, 140, 60, path(1000, 1000, LEAF), fill("accent1", { alpha: 0.6 }), { rot: -70 }) +
      oval("Zelle", 1130, 40, 120, 100, "accent2", { alpha: 0.3 }) + oval("Zellkern", 1175, 70, 34, 30, "accent1", { alpha: 0.8 }) + oval("Zelle 2", 1080, 130, 70, 60, "accent3", { alpha: 0.25 }) + oval("Zellkern 2", 1105, 152, 18, 16, "accent3", { alpha: 0.8 }) },
    section: { deco: sp("Blatt 1", 1080, 40, 160, 70, path(1000, 1000, LEAF), fill("accent1"), { rot: -25 }) + sp("Blatt 2", 1010, 100, 120, 52, path(1000, 1000, LEAF), fill("accent2"), { rot: 15 }) },
  },
  Chemie: {
    master: hexagon("Ring 1", 1180, 36, 20, "accent1", 1.5, true) + hexagon("Ring 2", 1214, 56, 20, "accent1", 1.5) + polyline("Bindung", [[1234, 56], [1262, 40]], "accent1", 1.5) + oval("Atom", 1258, 35, 9, 9, "accent2"),
    title: { deco: hexagon("Ring 1", 1100, 100, 40, "accent1", 2, true) + hexagon("Ring 2", 1170, 140, 40, "accent1", 2) + hexagon("Ring 3", 1170, 60, 40, "accent2", 2) +
      polyline("Bindung 1", [[1210, 140], [1250, 160]], "accent1", 2) + oval("Atom 1", 1246, 154, 14, 14, "accent3") + polyline("Bindung 2", [[1060, 100], [1020, 80]], "accent1", 2) + oval("Atom 2", 1010, 72, 14, 14, "accent4") +
      hexagon("Ring unten", 90, 640, 46, "accent2", 2, true) + polyline("Bindung 3", [[136, 640], [180, 610]], "accent2", 2) + oval("Atom 3", 176, 602, 14, 14, "accent1") },
    section: { deco: hexagon("Ring 1", 1100, 90, 40, "accent1", 2, true) + hexagon("Ring 2", 1170, 130, 40, "accent1", 2) + hexagon("Ring 3", 1240, 90, 40, "accent2", 2) },
  },
  Physik: {
    master: wave("Welle", 0, 1280, 712, 5, 80, "accent1", 1.25, { alpha: 0.55 }),
    title: { deco: [0, 60, 120].map((r, i) => ring(`Bahn ${i + 1}`, 1080, 60, 180, 60, "accent1", 2, { alpha: 0.85 }, { rot: r })).join("") + oval("Kern", 1161, 81, 18, 18, "accent2") +
      wave("Welle", 0, 1280, 600, 26, 260, "accent2", 2.5, { alpha: 0.6 }) + wave("Welle 2", 0, 1280, 600, 14, 130, "accent1", 1.5, { alpha: 0.4 }, 1) },
    section: { deco: wave("Welle", 0, 1280, 110, 30, 320, "accent1", 2.5, { alpha: 0.7 }) + wave("Welle 2", 0, 1280, 110, 16, 160, "accent2", 1.5, { alpha: 0.5 }, 1.2) },
  },
  Informatik: {
    master: circuit(3, 1150, 20, 120, 140, 4, "accent1") + circuit(5, 1050, 670, 220, 40, 3, "accent2"),
    title: { deco: circuit(11, 20, 40, 120, 600, 7, "accent1") + circuit(13, 1140, 40, 120, 600, 7, "accent2") + circuit(17, 200, 600, 880, 90, 6, "accent1") },
    section: { deco: circuit(19, 900, 20, 360, 140, 8, "accent1") },
  },
  Musik: {
    master: staff(0, 1280, 696, 4, "accent1", 0.35) + note("Note 1", 1050, 701, 6, "accent1") + note("Note 2", 1110, 695, 6, "accent2"),
    title: { hideMaster: true, deco: staff(0, 1280, 578, 14, "accent1", 0.45, 22) + note("Note 1", 200, 580, 16, "accent1") + note("Note 2", 300, 607, 16, "accent2") + note("Note 3", 1000, 590, 16, "accent1") + note("Note 4", 1080, 566, 16, "accent2", false) +
      oval("Ton", 1140, 60, 70, 70, "accent3", { alpha: 0.35 }) + oval("Ton 2", 1190, 100, 40, 40, "accent2", { alpha: 0.35 }) },
    section: { deco: staff(0, 1280, 90, 12, "accent1", 0.45, 14) + note("Note 1", 1050, 100, 14, "accent1") + note("Note 2", 1120, 116, 14, "accent2") },
  },
  Kunst: {
    master: blob("Farbfleck 1", 1180, -20, 110, 90, 101, "accent1", { alpha: 0.85 }) + blob("Farbfleck 2", 1150, 50, 60, 50, 103, "accent3", { alpha: 0.85 }) +
      sp("Pinselstrich", -20, 690, 300, 30, path(1000, 1000, "M 0 600 C 200 100 500 300 700 200 C 850 120 950 250 1000 300 C 930 600 800 700 600 650 C 400 600 200 900 0 800 Z"), fill("accent2", { alpha: 0.7 })),
    title: { deco: blob("Farbfleck 1", -60, -40, 250, 200, 107, "accent1", { alpha: 0.85 }) + blob("Farbfleck 2", 120, 40, 110, 90, 109, "accent3", { alpha: 0.85 }) +
      blob("Farbfleck 3", 1060, 540, 260, 210, 113, "accent2", { alpha: 0.85 }) + blob("Farbfleck 4", 1000, 600, 90, 80, 127, "accent5", { alpha: 0.8 }) + blob("Farbfleck 5", 1180, 40, 80, 70, 131, "accent4", { alpha: 0.85 }) +
      sp("Pinselstrich", 280, 590, 700, 50, path(1000, 1000, "M 0 600 C 200 100 500 300 700 200 C 850 120 950 250 1000 300 C 930 600 800 700 600 650 C 400 600 200 900 0 800 Z"), fill("accent6", { alpha: 0.6 })) },
    section: { deco: sp("Pinselstrich", 60, 645, 600, 30, path(1000, 1000, "M 0 600 C 200 100 500 300 700 200 C 850 120 950 250 1000 300 C 930 600 800 700 600 650 C 400 600 200 900 0 800 Z"), fill("accent1", { alpha: 0.45 })) +
      blob("Farbfleck", 1100, 30, 140, 110, 137, "accent3", { alpha: 0.85 }) },
  },
  Sport: {
    master: [0, 1, 2, 3].map((i) => sp(`Tempostreifen ${i + 1}`, 1120 + i * 34, 690, 44, 30, prst("parallelogram", { adj: 60000 }), fill(`accent${i + 1}`))).join(""),
    title: { deco: [0, 1, 2, 3, 4].map((i) => sp(`Bahn ${i + 1}`, -260 - i * 40, 460 - i * 40, 520 + i * 80, 520 + i * 80, prst("arc", { adj1: 16200000, adj2: 0 }), NOFILL, { line: stroke(i === 2 ? "accent1" : "tx2", i === 2 ? 3 : 1.5, { alpha: i === 2 ? 1 : 0.35 }) })).join("") +
      [0, 1, 2, 3].map((i) => sp(`Tempostreifen ${i + 1}`, 1020 + i * 60, 20, 90, 60, prst("parallelogram", { adj: 60000 }), fill(`accent${i + 1}`))).join("") },
    section: { deco: [0, 1, 2, 3, 4, 5].map((i) => sp(`Tempostreifen ${i + 1}`, 880 + i * 64, 60, 100, 70, prst("parallelogram", { adj: 60000 }), fill(`accent${(i % 4) + 1}`, { alpha: 0.9 }))).join("") },
  },
  Religion: {
    master: rays(0, 0, 7, 8, 82, 420, "accent2", 0.1) + oval("Licht", -60, -60, 120, 120, "accent2", { alpha: 0.35 }),
    title: { hideMaster: true, deco: rays(640, -60, 9, 30, 150, 900, "accent2", 0.07) + oval("Licht", 560, -140, 160, 160, "accent2", { alpha: 0.4 }) + rect("Linie", 560, 372, 160, 2, "accent2") },
    section: { deco: rays(0, 360, 9, -60, 60, 700, "accent2", 0.07) },
  },
  "Werte und Normen": {
    master: ring("Ring 1", 1170, 22, 44, 44, "accent1", 2) + ring("Ring 2", 1200, 22, 44, 44, "accent2", 2) + ring("Ring 3", 1185, 46, 44, 44, "accent3", 2),
    title: { deco: ring("Ring 1", 30, 420, 170, 170, "accent1", 3) + ring("Ring 2", 140, 420, 170, 170, "accent2", 3) + ring("Ring 3", 85, 515, 170, 170, "accent3", 3) +
      polyline("Waage", [[1100, 140], [1240, 110]], "tx2", 2.5) + sp("Stütze", 1158, 126, 26, 70, prst("triangle"), fill("tx2", { alpha: 0.8 })) },
    section: { deco: ring("Ring 1", 1040, 40, 120, 120, "accent1", 3) + ring("Ring 2", 1120, 40, 120, 120, "accent2", 3) },
  },
  "Darstellendes Spiel": {
    master: rect("Bühnenboden", 0, 706, 1280, 14, "accent1", { alpha: 0.5 }) + poly("Lichtkegel", [[640, -20], [380, 720], [900, 720]], "tx1", { alpha: 0.04 }),
    title: { deco: oval("Spot", 330, 120, 620, 440, "tx1", { alpha: 0.06 }) +
      sp("Vorhang links", 0, 0, 150, 720, path(1000, 1000, "M 0 0 L 1000 0 C 700 300 650 600 900 1000 L 0 1000 Z"), fill("accent2")) +
      sp("Vorhang rechts", 1130, 0, 150, 720, path(1000, 1000, "M 0 0 L 1000 0 C 700 300 650 600 900 1000 L 0 1000 Z"), fill("accent2"), { flipH: true }) +
      [30, 70, 110].map((x, i) => curve(`Falte links ${i + 1}`, [[x, 0], [x - 10, 360], [x + 10, 720]], "tx1", 1, { alpha: 0.2 })).join("") +
      [1170, 1210, 1250].map((x, i) => curve(`Falte rechts ${i + 1}`, [[x, 0], [x + 10, 360], [x - 10, 720]], "tx1", 1, { alpha: 0.2 })).join("") +
      sp("Volant", 0, 0, 1280, 50, path(1000, 1000, "M 0 0 L 1000 0 L 1000 600 Q 950 1000 900 600 Q 850 1000 800 600 Q 750 1000 700 600 Q 650 1000 600 600 Q 550 1000 500 600 Q 450 1000 400 600 Q 350 1000 300 600 Q 250 1000 200 600 Q 150 1000 100 600 Q 50 1000 0 600 Z"), fill("accent2", { lumMod: 75 })) },
    section: { deco: sp("Vorhang", 0, 0, 70, 720, path(1000, 1000, "M 0 0 L 1000 0 C 700 300 650 600 900 1000 L 0 1000 Z"), fill("accent2")) },
  },
  Seminarfach: {
    master: rect("Karte 2", 1182, 16, 80, 50, "bg2", undefined, { rot: 8, line: stroke("accent1", 1) }) + rect("Karte 1", 1170, 22, 80, 50, "bg1", undefined, { rot: -4, line: stroke("accent1", 1) }) +
      rect("Zeile 1", 1180, 38, 60, 1.5, "accent1", { alpha: 0.6 }, { rot: -4 }) + rect("Zeile 2", 1180, 50, 48, 1.5, "accent1", { alpha: 0.6 }, { rot: -4 }) + rect("Fußnotenstrich", 88, 652, 140, 1, "accent1"),
    title: { deco: rect("Karte 3", 1080, 60, 170, 110, "bg2", undefined, { rot: 10, line: stroke("accent1", 1.25) }) + rect("Karte 2", 1060, 70, 170, 110, "bg2", undefined, { rot: 3, line: stroke("accent1", 1.25) }) +
      rect("Karte 1", 1040, 80, 170, 110, "bg1", undefined, { rot: -5, line: stroke("accent1", 1.25) }) +
      [0, 1, 2].map((i) => rect(`Zeile ${i + 1}`, 1060, 112 + i * 22, 120 - i * 20, 2, "accent1", { alpha: 0.6 }, { rot: -5 })).join("") +
      rect("Klammer", 1040, 70, 3, 40, "accent2", undefined, { rot: -5 }) + rect("Fußnotenstrich", 160, 600, 200, 1.25, "accent1") },
    section: { deco: rect("Karte 2", 1080, 40, 150, 96, "bg2", undefined, { rot: 6, line: stroke("accent1", 1.25) }) + rect("Karte 1", 1060, 50, 150, 96, "bg1", undefined, { rot: -4, line: stroke("accent1", 1.25) }) },
  },

  // ----- Anlässe -----
  Projektwoche: {
    master: blocks(3, 1160, 690, 7, 22),
    title: { deco: blocks(5, 1000, 660, 12, 52) + blocks(7, 20, 90, 6, 36) },
    section: { deco: blocks(11, 1010, 120, 8, 44) },
  },
  Klassenfahrt: {
    master: mountains(676, 44, "accent1", "accent3", 0.3, 0.35),
    title: { deco: oval("Sonne", 1080, 60, 110, 110, "accent4", { alpha: 0.85 }) + mountains(560, 160, "accent1", "accent3", 0.35, 0.45) +
      curve("Route", [[90, 610], [220, 572], [360, 596], [490, 548]], "accent2", 2.5, undefined, "dash") +
      sp("Ortsmarke", 478, 504, 30, 30, prst("teardrop"), fill("accent2"), { rot: 135 }) + oval("Ortsmarke Punkt", 487, 513, 12, 12, "bg1") },
    section: { deco: curve("Route", [[98, 120], [300, 70], [520, 110], [740, 60]], "accent2", 2.5, undefined, "dash") + sp("Ortsmarke", 728, 22, 26, 26, prst("teardrop"), fill("accent2"), { rot: 135 }) },
  },
  Elternabend: {
    master: oval("Kreis 1", 1150, -30, 90, 90, "accent1", { alpha: 0.3 }) + oval("Kreis 2", 1200, 10, 80, 80, "accent2", { alpha: 0.3 }) + oval("Kreis 3", 1160, 40, 60, 60, "accent3", { alpha: 0.35 }),
    title: { deco: rect("Fußfläche", 0, 620, 1280, 100, "bg2") + oval("Kreis 1", 1030, -60, 220, 220, "accent1", { alpha: 0.28 }) + oval("Kreis 2", 1140, 40, 180, 180, "accent2", { alpha: 0.28 }) + oval("Kreis 3", 1060, 110, 120, 120, "accent3", { alpha: 0.32 }) },
    section: { deco: oval("Kreis 1", 1080, -40, 160, 160, "accent1", { alpha: 0.28 }) + oval("Kreis 2", 1160, 40, 130, 130, "accent2", { alpha: 0.28 }) },
  },
  Abitur: {
    master: rect("Goldlinie oben", 0, 10, 1280, 2, "accent1") + rect("Goldlinie unten", 0, 708, 1280, 2, "accent1"),
    title: { hideMaster: true, deco: sp("Rahmen", 30, 30, 1220, 660, prst("rect"), NOFILL, { line: stroke("accent1", 1.5) }) + sp("Rahmen innen", 38, 38, 1204, 644, prst("rect"), NOFILL, { line: stroke("accent1", 0.75) }) +
      laurel(600, 600, "accent1", -1) + laurel(680, 600, "accent1", 1) + oval("Mitte", 632, 618, 16, 16, "accent1") },
    section: { deco: rect("Goldbalken", 50, 300, 6, 330, "accent1") },
  },
  "Tag der offenen Tür": {
    master: arches(1140, 676, 4, 26, 44, 8),
    title: { deco: poly("Lichtschein", [[1140, 556], [1250, 556], [1280, 720], [1060, 720]], "accent3", { alpha: 0.4 }) +
      sp("Türrahmen", 1140, 226, 110, 330, prst("rect"), NOFILL, { line: stroke("tx2", 3) }) + poly("Tür", [[1140, 226], [1196, 256], [1196, 586], [1140, 556]], "accent1") + oval("Klinke", 1184, 396, 7, 7, "accent3") +
      arches(30, 580, 5, 40, 110, 12) },
    section: { deco: arches(980, 40, 6, 36, 90, 10) },
  },
};

// ----- the designs ------------------------------------------------------------------------------------------------

/** A design of ours or Office's: the theme part's data, its artwork and where the gallery lists it. */
export type PresentDesign = DesignSpec & { group: string; school?: SchoolDesign };

const effectsXml = (name: string) => (name === "Office" ? undefined : THEME_EFFECTS.find((x) => x.name === name)?.xml);

function build(d: SchoolDesign): PresentDesign {
  const art = ART[d.name] ?? {};
  return {
    name: d.name, colorName: d.name, fontName: d.name, major: d.major, minor: d.minor, colors: d.colors, dark: d.dark, titleColor: d.title,
    effects: effectsXml(d.effects), masterDeco: art.master, background: art.background, layouts: { title: art.title, section: art.section },
    group: d.group, school: d,
  };
}

/** Office's design (PowerPoint 2013-2022) as one to vary too. */
const OFFICE_SCHOOL: SchoolDesign = { name: "Office", group: "Allgemein", major: OFFICE_DESIGN.major, minor: OFFICE_DESIGN.minor, effects: "Office", colors: OFFICE_DESIGN.colors as SchoolDesign["colors"] };

export const DESIGNS: PresentDesign[] = [{ ...OFFICE_DESIGN, group: "Office", school: OFFICE_SCHOOL }, ...SCHOOL_DESIGNS.map(build)];
export const RAABE_DESIGN = DESIGNS.find((d) => d.name === "Raabeschule")!;

/** Headings of the full Designs gallery, in order (D-197). */
export const DESIGN_HEADINGS = ["Office", "Raabeschule", "Allgemein", "Fächer", "Anlässe"];

/** Variante i (0-3) of a design: its colours, colour scheme name and brightness; the artwork stays (D-192). */
export function variantOf(d: PresentDesign, i: number): PresentDesign {
  if (!d.school || i === 0) return d;
  const v = designVariants(d.school)[i];
  return { ...d, colors: v.colors, colorName: v.colorName, dark: v.dark };
}

/** The variants of a design (four). */
export function variantsOf(d: PresentDesign): PresentDesign[] {
  return [0, 1, 2, 3].map((i) => variantOf(d, i));
}
