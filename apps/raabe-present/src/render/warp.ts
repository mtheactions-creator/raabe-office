// Texteffekte > Transformieren drawn (D-076): the text warp's paths (ECMA-376's definitions, textwarps.gen.ts,
// evaluated for the text rectangle and the warp's adjust values) carry the text. With an odd number of paths
// (Folgepfad: textArchUp, textArchDown, textCircle; textButton's three) the text follows each path, its letters
// spread over the path's length; with pairs of paths (the Verzerrung warps) each pair is the top and the bottom
// of a band and the band's lines are stretched between them. Each letter is drawn on its own (SVG text), placed,
// turned and scaled from the two curves at its position: our approximation of PowerPoint, which bends the letter
// outlines themselves. Also the text fill's Farbverlauf (a:gradFill in a run), for warped and plain text.

import { customGeometry, flattenPath } from "@raabe/text-engine";
import { PptxDom, readFill, type TextInfo } from "@raabe/ooxml";
import { TEXT_WARPS } from "../textwarps.gen";
import { fitSize } from "./text";

const PT_PX = 96 / 72;
const SVG = "http://www.w3.org/2000/svg";
type Pt = [number, number];

/** The shape's text warp (a:bodyPr/a:prstTxWarp), unless none or textNoShape. */
export function warpOf(sp: Element): { prst: string; adj: Record<string, number> } | null {
  const bp = sp.getElementsByTagNameNS(PptxDom.NS_A, "bodyPr")[0];
  const w = bp ? Array.from(bp.children).find((c) => c.localName === "prstTxWarp") : undefined;
  const prst = w?.getAttribute("prst");
  if (!prst || prst === "textNoShape" || !TEXT_WARPS[prst]) return null;
  const adj: Record<string, number> = {};
  for (const gd of Array.from(w!.getElementsByTagNameNS(PptxDom.NS_A, "gd"))) {
    const m = /^val\s+(-?\d+)/.exec(gd.getAttribute("fmla") ?? "");
    if (m) adj[gd.getAttribute("name") ?? ""] = Number(m[1]);
  }
  return { prst, adj };
}

/** A polyline measured along its length. */
class Curve {
  private len: number[] = [0];
  constructor(private pts: Pt[]) {
    for (let i = 1; i < pts.length; i++) this.len.push(this.len[i - 1] + Math.hypot(pts[i][0] - pts[i - 1][0], pts[i][1] - pts[i - 1][1]));
  }
  get length(): number { return this.len[this.len.length - 1]; }
  /** The point at a share of the length (0…1). */
  at(u: number): Pt {
    const L = this.length;
    if (!L || this.pts.length < 2) return this.pts[0] ?? [0, 0];
    const d = Math.min(Math.max(u, 0), 1) * L;
    let i = 1;
    while (i < this.len.length - 1 && this.len[i] < d) i++;
    const seg = this.len[i] - this.len[i - 1] || 1;
    const t = (d - this.len[i - 1]) / seg;
    const a = this.pts[i - 1], b = this.pts[i];
    return [a[0] + (b[0] - a[0]) * t, a[1] + (b[1] - a[1]) * t];
  }
  tangent(u: number): number {
    const a = this.at(Math.max(0, u - 0.002)), b = this.at(Math.min(1, u + 0.002));
    return Math.atan2(b[1] - a[1], b[0] - a[0]);
  }
}

const lerp = (a: Pt, b: Pt, t: number): Pt => [a[0] + (b[0] - a[0]) * t, a[1] + (b[1] - a[1]) * t];

interface Line { text: string; font: string; size: number; bold: boolean; italic: boolean; color: string; alpha: number; outline?: { width: number; color: string }; grad?: string }

let ctx: CanvasRenderingContext2D | null = null;
/** The line's ink above and below the baseline (px at its size): what PowerPoint stretches over a band. */
function ink(l: Line): { asc: number; desc: number } {
  ctx ??= document.createElement("canvas").getContext("2d");
  if (!ctx) return { asc: l.size * 0.7, desc: l.size * 0.2 };
  ctx.font = `${l.italic ? "italic " : ""}${l.bold ? "bold " : ""}100px "${l.font}", Calibri, Carlito, sans-serif`;
  const m = ctx.measureText(l.text);
  return { asc: (m.actualBoundingBoxAscent / 100) * l.size, desc: Math.max(0, (m.actualBoundingBoxDescent / 100) * l.size) };
}

function advances(l: Line): number[] {
  ctx ??= document.createElement("canvas").getContext("2d");
  if (!ctx) return Array.from(l.text, () => l.size * 0.5);
  ctx.font = `${l.italic ? "italic " : ""}${l.bold ? "bold " : ""}100px "${l.font}", Calibri, Carlito, sans-serif`;
  return Array.from(l.text, (c) => (ctx!.measureText(c).width / 100) * l.size);
}

/** The text's lines, each with its first run's look (warped text is one look per line, as WordArt mostly is). */
function lines(t: TextInfo, theme: Record<string, string>, sp: Element, gradId: (css: Fill) => string): Line[] {
  const out: Line[] = [];
  const ps = Array.from(sp.getElementsByTagNameNS(PptxDom.NS_A, "p"));
  t.paras.forEach((p, pi) => {
    const first = p.runs.find((r) => r.kind !== "math") ?? p.end;
    const gf = runGradient(ps[pi], theme);
    const base = { font: first.font, size: fitSize(first.size, t.body.autofit === "norm" ? t.body.fontScale : 1) * PT_PX, bold: first.bold, italic: first.italic, color: first.color, alpha: first.alpha, outline: first.outline, grad: gf ? gradId(gf) : undefined };
    const text = p.runs.map((r) => (r.kind === "br" ? "\n" : r.kind === "math" ? "" : r.text)).join("");
    for (const part of text.split("\n")) out.push({ ...base, text: part });
  });
  while (out.length > 1 && !out[out.length - 1].text) out.pop();
  return out;
}

type Fill = Extract<NonNullable<ReturnType<typeof readFill>>, { kind: "gradient" }>;

/** The first run's a:gradFill in a paragraph, if any. */
function runGradient(p: Element | undefined, theme: Record<string, string>): Fill | null {
  const r = p ? Array.from(p.getElementsByTagNameNS(PptxDom.NS_A, "rPr")).find((x) => Array.from(x.children).some((c) => c.localName === "gradFill")) : undefined;
  const g = r ? Array.from(r.children).find((c) => c.localName === "gradFill") : undefined;
  const f = g ? readFill(g, theme) : undefined;
  return f && f.kind === "gradient" ? f : null;
}

let gradCount = 0;

/** The warped text as SVG, w × h (the text rectangle). */
export function drawWarp(t: TextInfo, warp: { prst: string; adj: Record<string, number> }, w: number, h: number, sp: Element, theme: Record<string, string>): SVGSVGElement {
  const svg = document.createElementNS(SVG, "svg");
  svg.setAttribute("width", String(w)); svg.setAttribute("height", String(h));
  svg.style.cssText = `position:absolute;left:0;top:0;overflow:visible`;
  svg.classList.add("txt-warp");
  const defs = document.createElementNS(SVG, "defs");
  svg.append(defs);
  const gradId = (f: Fill) => {
    const id = `tg${++gradCount}`;
    defs.append(gradientEl(id, f, w, h));
    return `url(#${id})`;
  };
  const geom = customGeometry(TEXT_WARPS[warp.prst], w, h, warp.adj);
  // Open paths are read left to right (some of the standard's run right to left); closed ones as they are.
  const curves = geom.paths.map((p) => {
    const pts = flattenPath(p.d, 16);
    const a = pts[0], b = pts[pts.length - 1];
    const closed = !!a && !!b && Math.hypot(a[0] - b[0], a[1] - b[1]) < 1;
    return new Curve(!closed && a && b && b[0] < a[0] - 0.5 ? pts.reverse() : pts);
  });
  const ls = lines(t, theme, sp, gradId).filter((l) => l.text.length);
  if (!ls.length || !curves.length) return svg;
  const glyph = (l: Line, ch: string, x: number, y: number, rot: number, sx: number, sy: number) => {
    const el = document.createElementNS(SVG, "text");
    el.textContent = ch;
    el.setAttribute("text-anchor", "middle");
    el.setAttribute("transform", `translate(${x.toFixed(2)} ${y.toFixed(2)}) rotate(${((rot * 180) / Math.PI).toFixed(2)}) scale(${sx.toFixed(4)} ${sy.toFixed(4)})`);
    el.setAttribute("font-family", `"${l.font}", Calibri, Carlito, sans-serif`);
    el.setAttribute("font-size", l.size.toFixed(2));
    if (l.bold) el.setAttribute("font-weight", "bold");
    if (l.italic) el.setAttribute("font-style", "italic");
    el.setAttribute("fill", l.grad ?? `#${l.color}`);
    if (l.alpha) el.setAttribute("fill-opacity", String(1 - l.alpha));
    if (l.outline) { el.setAttribute("stroke", `#${l.outline.color}`); el.setAttribute("stroke-width", String(Math.max(0.5, l.outline.width * PT_PX))); el.setAttribute("paint-order", "stroke fill"); }
    svg.append(el);
  };
  if (curves.length % 2 === 1) {
    // Folgepfad: the lines shared out over the paths (one path takes them all, joined).
    const groups: Line[][] = curves.map(() => []);
    if (curves.length === 1) groups[0] = [{ ...ls[0], text: ls.map((l) => l.text).join(" ") }];
    else ls.forEach((l, i) => groups[Math.min(curves.length - 1, Math.floor((i * curves.length) / ls.length))].push(l));
    groups.forEach((g, ci) => {
      if (!g.length) return;
      const l = { ...g[0], text: g.map((x) => x.text).join(" ") };
      const adv = advances(l);
      const total = adv.reduce((a, b) => a + b, 0) || 1;
      const c = curves[ci];
      const k = c.length / total;
      let acc = 0;
      Array.from(l.text).forEach((ch, i) => {
        const u = (acc + adv[i] / 2) / total;
        acc += adv[i];
        const [x, y] = c.at(u);
        const rot = c.tangent(u);
        // The baseline lies on the path, the letter stands on it (its descent below).
        glyph(l, ch, x, y, rot, k, 1);
      });
    });
    return svg;
  }
  // Verzerrung: pairs of paths as bands, the lines shared out over the bands.
  const bands = curves.length / 2;
  for (let b = 0; b < bands; b++) {
    const top = curves[b * 2], bot = curves[b * 2 + 1];
    const mine = ls.filter((_, i) => Math.min(bands - 1, Math.floor((i * bands) / ls.length)) === b);
    mine.forEach((l, j) => {
      const v0 = j / mine.length, v1 = (j + 1) / mine.length;
      const adv = advances(l);
      const total = adv.reduce((a, c) => a + c, 0) || 1;
      const k = ink(l);
      const em = Math.max(1, k.asc + k.desc);
      let acc = 0;
      Array.from(l.text).forEach((ch, i) => {
        const u0 = acc / total, u1 = (acc + adv[i]) / total, u = (u0 + u1) / 2;
        acc += adv[i];
        const tp = lerp(top.at(u), bot.at(u), v0), bp = lerp(top.at(u), bot.at(u), v1);
        const height = Math.hypot(bp[0] - tp[0], bp[1] - tp[1]);
        const mid0 = lerp(top.at(u0), bot.at(u0), (v0 + v1) / 2), mid1 = lerp(top.at(u1), bot.at(u1), (v0 + v1) / 2);
        const width = Math.hypot(mid1[0] - mid0[0], mid1[1] - mid0[1]);
        const rot = Math.atan2(bp[1] - tp[1], bp[0] - tp[0]) - Math.PI / 2;
        // Where the band runs backwards (the far side of a ring) the letter comes out mirrored, as PowerPoint's
        // bending of the letter outlines gives it.
        const back = (mid1[0] - mid0[0]) * Math.cos(rot) + (mid1[1] - mid0[1]) * Math.sin(rot) < 0;
        const sy = height / em, sx = (adv[i] ? width / adv[i] : sy) * (back ? -1 : 1);
        // The line's ink from the top point down: its baseline lies the ink's ascent below the top.
        const bx = tp[0] - Math.sin(rot) * k.asc * sy, by = tp[1] + Math.cos(rot) * k.asc * sy;
        glyph(l, ch, bx, by, rot, sx, sy);
      });
    });
  }
  return svg;
}

function gradientEl(id: string, f: Fill, w: number, h: number): SVGElement {
  const g = document.createElementNS(SVG, f.path ? "radialGradient" : "linearGradient");
  g.id = id;
  g.setAttribute("gradientUnits", "userSpaceOnUse");
  if (f.path) { g.setAttribute("cx", String(w / 2)); g.setAttribute("cy", String(h / 2)); g.setAttribute("r", String(Math.hypot(w, h) / 2)); }
  else {
    const a = ((f.angle ?? 0) * Math.PI) / 180, cx = w / 2, cy = h / 2;
    const half = (Math.abs(Math.cos(a)) * w + Math.abs(Math.sin(a)) * h) / 2;
    g.setAttribute("x1", String(cx - Math.cos(a) * half)); g.setAttribute("y1", String(cy - Math.sin(a) * half));
    g.setAttribute("x2", String(cx + Math.cos(a) * half)); g.setAttribute("y2", String(cy + Math.sin(a) * half));
  }
  for (const s of f.stops) {
    const st = document.createElementNS(SVG, "stop");
    st.setAttribute("offset", String(s.pos));
    st.setAttribute("stop-color", `#${s.color}`);
    if (s.transparency) st.setAttribute("stop-opacity", String(1 - s.transparency));
    g.append(st);
  }
  return g;
}

/** Textfüllung > Farbverlauf on plain text: each run with a:gradFill gets its gradient painted into its letters. */
export function applyTextGradients(textEl: HTMLElement, sp: Element, theme: Record<string, string>): void {
  const ps = Array.from(sp.getElementsByTagNameNS(PptxDom.NS_A, "p"));
  ps.forEach((p, pi) => {
    const para = textEl.querySelector(`.para[data-p="${pi}"]`);
    if (!para) return;
    const spans = Array.from(para.querySelectorAll(":scope > .run, :scope > .br"));
    let k = 0;
    for (const c of Array.from(p.children)) {
      if (c.localName !== "r" && c.localName !== "br" && c.localName !== "fld") continue;
      const span = spans[k++] as HTMLElement | undefined;
      if (!span || c.localName === "br") continue;
      const rPr = Array.from(c.children).find((x) => x.localName === "rPr");
      const g = rPr ? Array.from(rPr.children).find((x) => x.localName === "gradFill") : undefined;
      const f = g ? readFill(g, theme) : undefined;
      if (!f || f.kind !== "gradient") continue;
      const stops = f.stops.map((s) => `#${s.color}${s.transparency ? Math.round((1 - s.transparency) * 255).toString(16).padStart(2, "0") : ""} ${(s.pos * 100).toFixed(1)}%`).join(",");
      span.style.backgroundImage = f.path ? `radial-gradient(circle, ${stops})` : `linear-gradient(${(f.angle ?? 0) + 90}deg, ${stops})`;
      span.style.backgroundClip = "text";
      span.style.setProperty("-webkit-background-clip", "text");
      span.style.setProperty("-webkit-text-fill-color", "transparent");
    }
  });
}
