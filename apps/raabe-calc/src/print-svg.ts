// A drawing context that records the grid renderer's canvas calls as SVG (D-094), so printed pages and PDFs
// get the sheet exactly as the screen draws it (conditional formats, data bars, table styles, sparklines,
// rotated text) while text stays text and lines stay vectors in the PDF. Only what the renderer uses is
// implemented: rectangles, paths with lines and arcs, clipping, transforms, text, dashes, linear gradients.

const NS = "http://www.w3.org/2000/svg";

type M = [number, number, number, number, number, number];
const I: M = [1, 0, 0, 1, 0, 0];
const mul = (m: M, n: M): M => [
  m[0] * n[0] + m[2] * n[1], m[1] * n[0] + m[3] * n[1],
  m[0] * n[2] + m[2] * n[3], m[1] * n[2] + m[3] * n[3],
  m[0] * n[4] + m[2] * n[5] + m[4], m[1] * n[4] + m[3] * n[5] + m[5],
];
const f = (v: number) => String(Math.round(v * 1000) / 1000);

interface State {
  m: M;
  fill: string | Gradient;
  stroke: string | Gradient;
  lineWidth: number;
  font: string;
  textAlign: CanvasTextAlign;
  textBaseline: CanvasTextBaseline;
  lineCap: CanvasLineCap;
  lineJoin: CanvasLineJoin;
  dash: number[];
  dashOffset: number;
  alpha: number;
  clip: string | null;
}

class Gradient {
  stops: [number, string][] = [];
  constructor(readonly x0: number, readonly y0: number, readonly x1: number, readonly y1: number, readonly m: M) {}
  addColorStop(at: number, color: string): void { this.stops.push([at, color]); }
}

let measurer: CanvasRenderingContext2D | null = null;
const measure = () => (measurer ??= document.createElement("canvas").getContext("2d")!);

export class SvgContext {
  readonly svg: SVGSVGElement;
  private defs: SVGDefsElement;
  private st: State = { m: I, fill: "#000000", stroke: "#000000", lineWidth: 1, font: "10px sans-serif", textAlign: "start", textBaseline: "alphabetic",
    lineCap: "butt", lineJoin: "miter", dash: [], dashOffset: 0, alpha: 1, clip: null };
  private stack: State[] = [];
  private path = "";
  private ids = 0;
  private readonly prefix: string;
  /** Font sizes rounded to whole device pixels (printing). */
  wholeFontPx = false;

  constructor(width: number, height: number, prefix = "p") {
    this.prefix = prefix;
    this.svg = document.createElementNS(NS, "svg");
    this.svg.setAttribute("width", f(width));
    this.svg.setAttribute("height", f(height));
    this.svg.setAttribute("viewBox", `0 0 ${f(width)} ${f(height)}`);
    this.svg.style.overflow = "hidden";
    this.defs = document.createElementNS(NS, "defs");
    this.svg.append(this.defs);
  }

  // ----- state -----------------------------------------------------------------------------------

  get fillStyle(): string | Gradient { return this.st.fill; }
  set fillStyle(v: string | Gradient) { this.st.fill = v; }
  get strokeStyle(): string | Gradient { return this.st.stroke; }
  set strokeStyle(v: string | Gradient) { this.st.stroke = v; }
  get lineWidth(): number { return this.st.lineWidth; }
  set lineWidth(v: number) { this.st.lineWidth = v; }
  get font(): string { return this.st.font; }
  set font(v: string) { this.st.font = v; }
  get textAlign(): CanvasTextAlign { return this.st.textAlign; }
  set textAlign(v: CanvasTextAlign) { this.st.textAlign = v; }
  get textBaseline(): CanvasTextBaseline { return this.st.textBaseline; }
  set textBaseline(v: CanvasTextBaseline) { this.st.textBaseline = v; }
  get lineCap(): CanvasLineCap { return this.st.lineCap; }
  set lineCap(v: CanvasLineCap) { this.st.lineCap = v; }
  get lineJoin(): CanvasLineJoin { return this.st.lineJoin; }
  set lineJoin(v: CanvasLineJoin) { this.st.lineJoin = v; }
  get lineDashOffset(): number { return this.st.dashOffset; }
  set lineDashOffset(v: number) { this.st.dashOffset = v; }
  get globalAlpha(): number { return this.st.alpha; }
  set globalAlpha(v: number) { this.st.alpha = v; }
  setLineDash(d: number[]): void { this.st.dash = [...d]; }
  getLineDash(): number[] { return [...this.st.dash]; }

  save(): void { this.stack.push({ ...this.st, dash: [...this.st.dash] }); }
  restore(): void { const s = this.stack.pop(); if (s) this.st = s; }

  setTransform(a: number, b: number, c: number, d: number, e: number, g: number): void { this.st.m = [a, b, c, d, e, g]; }
  resetTransform(): void { this.st.m = I; }
  transform(a: number, b: number, c: number, d: number, e: number, g: number): void { this.st.m = mul(this.st.m, [a, b, c, d, e, g]); }
  translate(x: number, y: number): void { this.transform(1, 0, 0, 1, x, y); }
  scale(x: number, y: number): void { this.transform(x, 0, 0, y, 0, 0); }
  rotate(a: number): void { const c = Math.cos(a), s = Math.sin(a); this.transform(c, s, -s, c, 0, 0); }

  private pt(x: number, y: number): string {
    const m = this.st.m;
    return `${f(m[0] * x + m[2] * y + m[4])} ${f(m[1] * x + m[3] * y + m[5])}`;
  }

  /** How much the current transform scales lengths (for line widths). */
  private get k(): number {
    const m = this.st.m;
    return Math.sqrt(Math.abs(m[0] * m[3] - m[1] * m[2])) || 1;
  }

  // ----- paths -----------------------------------------------------------------------------------

  beginPath(): void { this.path = ""; }
  moveTo(x: number, y: number): void { this.path += `M${this.pt(x, y)}`; }
  lineTo(x: number, y: number): void { this.path += (this.path ? "L" : "M") + this.pt(x, y); }
  closePath(): void { this.path += "Z"; }
  rect(x: number, y: number, w: number, h: number): void {
    this.path += `M${this.pt(x, y)}L${this.pt(x + w, y)}L${this.pt(x + w, y + h)}L${this.pt(x, y + h)}Z`;
  }
  arc(x: number, y: number, r: number, a0: number, a1: number, ccw = false): void { this.ellipse(x, y, r, r, 0, a0, a1, ccw); }
  ellipse(x: number, y: number, rx: number, ry: number, rot: number, a0: number, a1: number, ccw = false): void {
    let sweep = a1 - a0;
    if (!ccw && sweep < 0) sweep += Math.PI * 2;
    if (ccw && sweep > 0) sweep -= Math.PI * 2;
    if (Math.abs(a1 - a0) >= Math.PI * 2) sweep = ccw ? -Math.PI * 2 : Math.PI * 2;
    const n = Math.max(8, Math.ceil((Math.abs(sweep) / (Math.PI * 2)) * 72));
    const cr = Math.cos(rot), sr = Math.sin(rot);
    for (let i = 0; i <= n; i++) {
      const a = a0 + (sweep * i) / n;
      const px = rx * Math.cos(a), py = ry * Math.sin(a);
      const X = x + px * cr - py * sr, Y = y + px * sr + py * cr;
      this.path += (i === 0 && this.path ? "L" : i === 0 ? "M" : "L") + this.pt(X, Y);
    }
  }

  // ----- painting --------------------------------------------------------------------------------

  private paint(style: string | Gradient): string {
    if (typeof style === "string") return style;
    const id = `${this.prefix}g${this.ids++}`;
    const g = document.createElementNS(NS, "linearGradient");
    g.setAttribute("id", id);
    g.setAttribute("gradientUnits", "userSpaceOnUse");
    const m = style.m;
    const p0 = [m[0] * style.x0 + m[2] * style.y0 + m[4], m[1] * style.x0 + m[3] * style.y0 + m[5]];
    const p1 = [m[0] * style.x1 + m[2] * style.y1 + m[4], m[1] * style.x1 + m[3] * style.y1 + m[5]];
    g.setAttribute("x1", f(p0[0])); g.setAttribute("y1", f(p0[1])); g.setAttribute("x2", f(p1[0])); g.setAttribute("y2", f(p1[1]));
    for (const [at, color] of style.stops) {
      const s = document.createElementNS(NS, "stop");
      s.setAttribute("offset", String(at));
      s.setAttribute("stop-color", color);
      g.append(s);
    }
    this.defs.append(g);
    return `url(#${id})`;
  }

  private emit(el: SVGElement): void {
    if (this.st.clip) el.setAttribute("clip-path", `url(#${this.st.clip})`);
    if (this.st.alpha < 1) el.setAttribute("opacity", String(this.st.alpha));
    this.svg.append(el);
  }

  private pathEl(d: string): SVGPathElement {
    const p = document.createElementNS(NS, "path");
    p.setAttribute("d", d);
    return p;
  }

  createLinearGradient(x0: number, y0: number, x1: number, y1: number): Gradient { return new Gradient(x0, y0, x1, y1, this.st.m); }

  fill(): void {
    if (!this.path) return;
    const p = this.pathEl(this.path);
    p.setAttribute("fill", this.paint(this.st.fill));
    this.emit(p);
  }

  stroke(): void {
    if (!this.path) return;
    const p = this.pathEl(this.path);
    p.setAttribute("fill", "none");
    this.strokeAttrs(p);
    this.emit(p);
  }

  private strokeAttrs(p: SVGElement): void {
    const k = this.k;
    p.setAttribute("stroke", this.paint(this.st.stroke));
    p.setAttribute("stroke-width", f(this.st.lineWidth * k));
    if (this.st.lineCap !== "butt") p.setAttribute("stroke-linecap", this.st.lineCap);
    if (this.st.lineJoin !== "miter") p.setAttribute("stroke-linejoin", this.st.lineJoin);
    if (this.st.dash.length) {
      p.setAttribute("stroke-dasharray", this.st.dash.map((d) => f(d * k)).join(" "));
      if (this.st.dashOffset) p.setAttribute("stroke-dashoffset", f(this.st.dashOffset * k));
    }
  }

  fillRect(x: number, y: number, w: number, h: number): void {
    if (w === 0 || h === 0) return;
    const p = this.pathEl(`M${this.pt(x, y)}L${this.pt(x + w, y)}L${this.pt(x + w, y + h)}L${this.pt(x, y + h)}Z`);
    p.setAttribute("fill", this.paint(this.st.fill));
    this.emit(p);
  }

  strokeRect(x: number, y: number, w: number, h: number): void {
    const p = this.pathEl(`M${this.pt(x, y)}L${this.pt(x + w, y)}L${this.pt(x + w, y + h)}L${this.pt(x, y + h)}Z`);
    p.setAttribute("fill", "none");
    this.strokeAttrs(p);
    this.emit(p);
  }

  clearRect(x: number, y: number, w: number, h: number): void {
    const keep = this.st.fill;
    this.st.fill = "#ffffff";
    this.fillRect(x, y, w, h);
    this.st.fill = keep;
  }

  clip(): void {
    const id = `${this.prefix}c${this.ids++}`;
    const c = document.createElementNS(NS, "clipPath");
    c.setAttribute("id", id);
    if (this.st.clip) c.setAttribute("clip-path", `url(#${this.st.clip})`);
    c.append(this.pathEl(this.path || "M0 0Z"));
    this.defs.append(c);
    this.st.clip = id;
  }

  // ----- text ------------------------------------------------------------------------------------

  measureText(text: string): TextMetrics {
    const m = measure();
    m.font = this.st.font;
    m.textBaseline = this.st.textBaseline;
    m.textAlign = this.st.textAlign;
    return m.measureText(text);
  }

  fillText(text: string, x: number, y: number): void {
    const m = measure();
    m.font = this.st.font;
    // SVG text sits on its alphabetic baseline: move the canvas's baseline there.
    let dy = 0;
    if (this.st.textBaseline !== "alphabetic") {
      m.textBaseline = "alphabetic";
      const a = m.measureText(text).fontBoundingBoxAscent;
      m.textBaseline = this.st.textBaseline;
      const b = m.measureText(text).fontBoundingBoxAscent;
      dy = a - b;
    }
    const t = document.createElementNS(NS, "text");
    const sz = /([\d.]+)px/.exec(this.st.font);
    const family = this.st.font.replace(/^.*?[\d.]+px\s*/, "");
    t.setAttribute("font-family", family.replace(/"/g, "'"));
    // Printers set type at whole printer pixels (Excel's 11 pt prints as 92 px at 600 dpi, 11,04 pt).
    if (sz) t.setAttribute("font-size", this.wholeFontPx ? f(Math.round(Number(sz[1]) * this.k) / this.k) : sz[1]);
    if (/\bbold\b/.test(this.st.font)) t.setAttribute("font-weight", "bold");
    if (/\bitalic\b/.test(this.st.font)) t.setAttribute("font-style", "italic");
    t.setAttribute("x", f(x));
    t.setAttribute("y", f(y + dy));
    const anchor = this.st.textAlign === "center" ? "middle" : this.st.textAlign === "right" || this.st.textAlign === "end" ? "end" : null;
    if (anchor) t.setAttribute("text-anchor", anchor);
    const mm = this.st.m;
    if (mm !== I && !(mm[0] === 1 && mm[1] === 0 && mm[2] === 0 && mm[3] === 1 && mm[4] === 0 && mm[5] === 0)) t.setAttribute("transform", `matrix(${mm.map(f).join(" ")})`);
    t.setAttribute("fill", this.paint(this.st.fill));
    t.style.whiteSpace = "pre";
    t.textContent = text;
    // A clip path is read in the user space of the element it is on, transform included: a group carries it.
    if (t.hasAttribute("transform") && this.st.clip) {
      const g = document.createElementNS(NS, "g");
      g.append(t);
      this.emit(g);
    } else this.emit(t);
  }

  strokeText(text: string, x: number, y: number): void { this.fillText(text, x, y); }
}

/** The context as the renderer's type (it uses only what SvgContext has). */
export function svgContext(width: number, height: number, prefix?: string, wholeFontPx = false): { ctx: CanvasRenderingContext2D; svg: SVGSVGElement } {
  const c = new SvgContext(width, height, prefix);
  c.wholeFontPx = wholeFontPx;
  return { ctx: c as unknown as CanvasRenderingContext2D, svg: c.svg };
}
