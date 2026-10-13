// Fonts: vertical metrics (from the font files, via the platform), font fallbacks, CSS font
// strings and horizontal measurement. Widths come from the browser's own shaping (canvas),
// so they match what the DOM renders; kerning, when a run asks for it, is applied as
// measured pair adjustments.

import type { ResolvedChar } from "./styles";

export const PT = 4 / 3;

/** Size at which characters are measured (see Fonts.measureFont). */
const MEASURE_PX = 1000;

/** Vertical metrics in font units: the hhea table and the OS/2 Windows values. */
export interface FontMetrics {
  readonly found: boolean;
  readonly unitsPerEm: number;
  readonly ascender: number;
  readonly descender: number;
  readonly lineGap: number;
  readonly winAscent?: number;
  readonly winDescent?: number;
  /** The font has a MATH table (Cambria Math): only the hhea values count. */
  readonly math?: boolean;
}

export type MetricsProvider = (families: string[]) => Promise<Record<string, FontMetrics>>;

// Known values so the first layout is right before the platform answers.
const KNOWN: Record<string, FontMetrics> = {
  "calibri": { found: true, unitsPerEm: 2048, ascender: 1536, descender: -512, lineGap: 452, winAscent: 1950, winDescent: 550 },
  "calibri light": { found: true, unitsPerEm: 2048, ascender: 1536, descender: -512, lineGap: 452, winAscent: 1950, winDescent: 550 },
  "times new roman": { found: true, unitsPerEm: 2048, ascender: 1825, descender: -443, lineGap: 87, winAscent: 1825, winDescent: 443 },
  "arial": { found: true, unitsPerEm: 2048, ascender: 1854, descender: -434, lineGap: 67, winAscent: 1854, winDescent: 434 },
};

/** Metric-compatible open fonts for when a Microsoft font is missing (iPad, Linux). */
const COMPATIBLE: Record<string, string> = {
  "calibri": "Carlito", "calibri light": "Carlito", "cambria": "Caladea", "arial": "Liberation Sans",
  "times new roman": "Liberation Serif", "courier new": "Liberation Mono",
};

export class Fonts {
  private metrics = new Map<string, FontMetrics>(Object.entries(KNOWN));
  private pending = new Set<string>();
  private listeners: (() => void)[] = [];
  private ctx = document.createElement("canvas").getContext("2d")!;
  private advances = new Map<string, Map<string, number>>();
  private pairs = new Map<string, Map<string, number>>();
  private baselines = new Map<string, number>();
  installed: Set<string> | null = null;

  constructor(private provider: MetricsProvider | null = null) {}

  setProvider(p: MetricsProvider): void {
    this.provider = p;
  }

  /** Called when metrics arrive that may change the layout. */
  onChange(l: () => void): void {
    this.listeners.push(l);
  }

  /** Requests metrics for families not known yet (asynchronously). */
  require(families: Iterable<string>): void {
    if (!this.provider) return;
    const missing = [...families].map((f) => f.toLowerCase()).filter((f) => !this.metrics.has(f) && !this.pending.has(f));
    if (!missing.length) return;
    for (const f of missing) this.pending.add(f);
    void this.provider(missing).then((res) => {
      for (const [k, v] of Object.entries(res)) this.metrics.set(k.toLowerCase(), v);
      for (const f of missing) this.pending.delete(f);
      this.baselines.clear();
      for (const l of this.listeners) l();
    }).catch(() => { for (const f of missing) this.pending.delete(f); });
  }

  private canvasMetrics(family: string): FontMetrics {
    this.ctx.font = `100px ${this.cssFamily(family)}`;
    const m = this.ctx.measureText("Hg");
    const asc = m.fontBoundingBoxAscent ?? 90, desc = m.fontBoundingBoxDescent ?? 25;
    return { found: true, unitsPerEm: 100, ascender: asc, descender: -desc, lineGap: 0 };
  }

  metricsFor(family: string): FontMetrics {
    const m = this.metrics.get(family.toLowerCase());
    if (m && m.found) return m;
    const alt = COMPATIBLE[family.toLowerCase()];
    if (alt) {
      const am = this.metrics.get(alt.toLowerCase());
      if (am?.found) return am;
    }
    return this.canvasMetrics(family);
  }

  // Word's rules, measured against Word's PDF output for twelve fonts (docs/DECISIONS.md D-025):
  // the single line height is the larger of the hhea sum (ascender + |descender| + line gap)
  // and the Windows sum (winAscent + winDescent); below the baseline lies winDescent, the rest
  // (including any line gap) lies above it. Math fonts (Cambria Math, whose Windows values
  // are about five times too large) count with their hhea values only.

  /** Word's single line height (points) for a font at a size. */
  lineHeight(family: string, sizePt: number): number {
    const m = this.metricsFor(family);
    const hhea = m.ascender - m.descender + m.lineGap;
    const win = m.math ? 0 : (m.winAscent ?? 0) + (m.winDescent ?? 0);
    return (sizePt * Math.max(hhea, win)) / m.unitsPerEm;
  }

  /** Part of the single line height below the baseline (points). */
  descent(family: string, sizePt: number): number {
    const m = this.metricsFor(family);
    return (sizePt * ((!m.math && m.winDescent) || -m.descender)) / m.unitsPerEm;
  }

  /** Part of the single line height above the baseline (points). */
  ascent(family: string, sizePt: number): number {
    return this.lineHeight(family, sizePt) - this.descent(family, sizePt);
  }

  /** Metric requests still waiting for the platform (the layout changes when they arrive). */
  get pendingCount(): number {
    return this.pending.size;
  }

  cssFamily(family: string): string {
    const alt = COMPATIBLE[family.toLowerCase()];
    return `"${family}"${alt ? `, "${alt}"` : ""}, sans-serif`;
  }

  /**
   * Effective size in points for measuring/rendering. Super- and subscript: 65 % rounded to
   * half points, as Word does (measured from its PDF output for 8-36 pt).
   */
  static renderSize(c: ResolvedChar): number {
    return c.vertAlign === "baseline" ? c.size : Math.round(c.size * 1.3) / 2;
  }

  cssFont(c: ResolvedChar): string {
    return `${c.italic ? "italic " : ""}${c.bold ? "bold " : ""}${Fonts.renderSize(c) * PT}px ${this.cssFamily(c.font)}`;
  }

  /** Key identifying everything that changes glyph widths. */
  static widthKey(c: ResolvedChar): string {
    return `${c.font}|${c.bold}|${c.italic}|${Fonts.renderSize(c)}|${c.caps}|${c.smallCaps}|${c.spacing}`;
  }

  private shaped(c: ResolvedChar, ch: string): string {
    return c.caps || c.smallCaps ? ch.toUpperCase() : ch;
  }

  /**
   * Canvas font for measuring. Measuring happens at a large size and is scaled down: at small
   * sizes the canvas returns hinted, whole-pixel advances for some fonts (Cambria: "i" 4 px
   * instead of 4.45 px), while both Word and the DOM use the design advances.
   */
  private measureFont(c: ResolvedChar): string {
    return `${c.italic ? "italic " : ""}${c.bold ? "bold " : ""}${MEASURE_PX}px ${this.cssFamily(c.font)}`;
  }

  /** Advance width of one character in CSS px (including character spacing). */
  advance(c: ResolvedChar, ch: string): number {
    const key = Fonts.widthKey(c);
    let m = this.advances.get(key);
    if (!m) { m = new Map(); this.advances.set(key, m); }
    let w = m.get(ch);
    if (w === undefined) {
      this.ctx.font = this.measureFont(c);
      this.ctx.fontKerning = "none";
      const s = this.shaped(c, ch);
      w = this.ctx.measureText(s).width * (Fonts.renderSize(c) * PT / MEASURE_PX);
      if (c.smallCaps && ch !== s) w *= 0.8;
      w += c.spacing * PT;
      m.set(ch, w);
    }
    return w;
  }

  /** Kerning adjustment between two characters of the same run (CSS px), when enabled. */
  kerning(c: ResolvedChar, a: string, b: string): number {
    if (!c.kern || c.size < c.kern) return 0;
    const key = Fonts.widthKey(c);
    let m = this.pairs.get(key);
    if (!m) { m = new Map(); this.pairs.set(key, m); }
    const pair = a + b;
    let k = m.get(pair);
    if (k === undefined) {
      this.ctx.font = this.measureFont(c);
      this.ctx.fontKerning = "normal";
      const both = this.ctx.measureText(this.shaped(c, a) + this.shaped(c, b)).width;
      this.ctx.fontKerning = "none";
      const apart = this.ctx.measureText(this.shaped(c, a)).width + this.ctx.measureText(this.shaped(c, b)).width;
      k = (both - apart) * (Fonts.renderSize(c) * PT / MEASURE_PX);
      m.set(pair, k);
    }
    return k;
  }

  /**
   * Distance (CSS px) from the top of an element with `line-height: <lh>px` to its text
   * baseline, measured once per font so the text sits exactly where the layout wants it.
   */
  baselineOffset(c: ResolvedChar, lineHeightPx: number): number {
    const key = `${this.cssFont(c)}|${lineHeightPx}`;
    let b = this.baselines.get(key);
    if (b === undefined) {
      const probe = document.createElement("div");
      probe.style.cssText = `position:absolute;visibility:hidden;left:-9999px;top:0;white-space:pre;font:${this.cssFont(c)};line-height:${lineHeightPx}px`;
      const mark = document.createElement("span");
      mark.style.cssText = "display:inline-block;width:0;height:0;vertical-align:baseline";
      probe.append("x", mark);
      document.body.append(probe);
      // Whole pixels, as the browser places the baseline when painting (a sub-pixel
      // measurement made the drawn text deviate more, measured in the PDF output).
      b = mark.offsetTop;
      probe.remove();
      this.baselines.set(key, b);
    }
    return b;
  }

  /** Clears measurement caches (e.g. after web fonts finished loading). */
  reset(): void {
    this.advances.clear();
    this.pairs.clear();
    this.baselines.clear();
  }
}
