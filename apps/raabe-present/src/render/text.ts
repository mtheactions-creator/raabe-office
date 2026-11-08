// Text frames as PowerPoint draws them (D-070): paragraphs with their level's indents, bullets
// hanging in the indent, line spacing as a share of the font's line height, space before and after,
// anchoring top / middle / bottom, Autofit's font scale and line spacing reduction. Lines are broken
// by the browser at the frame's width; positions are measured against PowerPoint's own pictures
// (check_present_render.py). Every run is a span carrying its paragraph offset (data-o), so the
// editor maps between the caret and the DOM.

import { ommlToSeq } from "@raabe/ooxml";
import { ensureFontFamily, seqToMathml } from "@raabe/text-engine";

/** A font family for CSS, made known to the browser first: Office's cloud fonts (Aptos …) and fonts known only by
 *  their full names are registered by @raabe/text-engine (D-089, D-109). */
const fam = (name: string | undefined): string => { ensureFontFamily(name); return `"${name ?? ""}"`; };
import type { BodyProps, RunStyle, TextInfo, TextPara, TextRun } from "@raabe/ooxml";

export const EMU_PX = 9525;
export const PT_PX = 96 / 72;

const metricsCache = new Map<string, { ascent: number; descent: number }>();
let measureCtx: CanvasRenderingContext2D | null = null;

/** The font's ascent and descent as shares of its size (the browser's metrics for the face). */
export function fontMetrics(font: string): { ascent: number; descent: number } {
  let m = metricsCache.get(font);
  if (m) return m;
  measureCtx ??= document.createElement("canvas").getContext("2d");
  m = { ascent: 0.75, descent: 0.25 };
  if (measureCtx) {
    measureCtx.font = `100px "${font}", Calibri, Carlito, sans-serif`;
    const t = measureCtx.measureText("Hg");
    if (t.fontBoundingBoxAscent) m = { ascent: t.fontBoundingBoxAscent / 100, descent: t.fontBoundingBoxDescent / 100 };
  }
  metricsCache.set(font, m);
  return m;
}

export interface TextOptions {
  /** Show the prompt of an empty placeholder (Normalansicht, not in thumbnails or the show). */
  prompt: boolean;
  /** Text of fields: slide number, date. */
  slideNumber?: number;
  /** Selected text as paragraph offsets (drawn behind the text). */
  selection?: { from: { p: number; o: number }; to: { p: number; o: number } } | null;
  /** Autofit applied (font scale, spacing reduction); the body's own when missing. */
  fit?: { fontScale: number; lnSpcReduction: number };
}

const css = (s: Record<string, string | number | undefined>) => Object.entries(s).filter(([, v]) => v !== undefined && v !== "").map(([k, v]) => `${k}:${v}`).join(";");

function hexA(hex: string, alpha: number): string {
  if (!alpha) return `#${hex}`;
  const n = parseInt(hex, 16);
  return `rgba(${(n >> 16) & 255},${(n >> 8) & 255},${n & 255},${(1 - alpha).toFixed(3)})`;
}

/** A font size under Autofit's scale: PowerPoint rounds it to whole points (28 pt at 55 % is 15 pt, read back). */
export function fitSize(size: number, scale: number): number {
  return scale < 1 ? Math.max(1, Math.round(size * scale)) : size;
}

export function runCss(r: RunStyle, scale: number): string {
  const deco = [r.underline !== "none" || r.link ? "underline" : "", r.strike !== "noStrike" ? "line-through" : ""].filter(Boolean).join(" ");
  const size = fitSize(r.size, scale) * (r.baseline ? 2 / 3 : 1);
  return css({
    "font-family": `${fam(r.font)}, Calibri, Carlito, sans-serif`,
    "font-size": `${(size * PT_PX).toFixed(3)}px`,
    "font-weight": r.bold ? "bold" : undefined,
    "font-style": r.italic ? "italic" : undefined,
    "text-decoration": deco || undefined,
    "text-decoration-style": r.underline === "dbl" ? "double" : r.underline === "dotted" || r.underline === "dotDash" ? "dotted" : r.underline === "dash" ? "dashed" : r.underline === "wavy" ? "wavy" : undefined,
    color: hexA(r.color, r.alpha),
    "-webkit-text-stroke": r.outline ? `${Math.max(0.5, r.outline.width * PT_PX * (scale < 1 ? scale : 1)).toFixed(2)}px #${r.outline.color}` : undefined,
    "paint-order": r.outline ? "stroke fill" : undefined,
    "background-color": r.highlight ? `#${r.highlight}` : undefined,
    "text-transform": r.caps === "all" ? "uppercase" : undefined,
    "font-variant": r.caps === "small" ? "small-caps" : undefined,
    "letter-spacing": r.spacing ? `${(r.spacing * PT_PX).toFixed(2)}px` : undefined,
    "vertical-align": r.baseline > 0 ? `${(r.baseline * r.size * scale * PT_PX).toFixed(2)}px` : r.baseline < 0 ? `${(r.baseline * r.size * scale * PT_PX).toFixed(2)}px` : undefined,
  });
}

const ROMAN = [[1000, "m"], [900, "cm"], [500, "d"], [400, "cd"], [100, "c"], [90, "xc"], [50, "l"], [40, "xl"], [10, "x"], [9, "ix"], [5, "v"], [4, "iv"], [1, "i"]] as const;
const roman = (n: number) => { let s = ""; for (const [v, r] of ROMAN) while (n >= v) { s += r; n -= v; } return s; };
const alpha = (n: number) => { let s = ""; while (n > 0) { n--; s = String.fromCharCode(97 + (n % 26)) + s; n = Math.floor(n / 26); } return s; };

/** The label of an automatic number (buAutoNum type) for item `n`. */
export function autoNumber(type: string, n: number): string {
  const m = /^(arabic|alphaLc|alphaUc|romanLc|romanUc)(Period|ParenR|ParenBoth|Plain|Minus)?/.exec(type);
  const kind = m?.[1] ?? "arabic", suffix = m?.[2] ?? "Period";
  let v = kind === "arabic" ? String(n) : kind.startsWith("alpha") ? alpha(n) : roman(n);
  if (kind.endsWith("Uc")) v = v.toUpperCase();
  return suffix === "Period" ? `${v}.` : suffix === "ParenR" ? `${v})` : suffix === "ParenBoth" ? `(${v})` : suffix === "Minus" ? `${v} -` : v;
}

/** Text of a field run. */
function fieldText(r: TextRun, o: TextOptions): string {
  if (r.fieldType === "slidenum" && o.slideNumber) return String(o.slideNumber);
  if (r.fieldType?.startsWith("datetime")) return new Date().toLocaleDateString("de-DE", { day: "2-digit", month: "2-digit", year: "numeric" });
  return r.text;
}

/** Line height in px of a paragraph's largest run at its spacing. */
function lineHeight(p: TextPara, scale: number, reduction: number): number {
  // PowerPoint's single line is 1,2 times the largest font size of the line, whatever the font
  // (read back from PowerPoint: 28 pt at 90 % gives 30,24 pt lines, 44 pt titles 47,52 pt).
  const size = fitSize(largest(p), scale);
  if (p.lnPts !== undefined) return p.lnPts * PT_PX;
  const pct = Math.max(0.1, (p.lnPct ?? 1) - reduction);
  return size * 1.2 * pct * PT_PX;
}

/** The largest run of a paragraph (its mark when it has no text). */
function largest(p: TextPara): number {
  return p.runs.length ? Math.max(...p.runs.map((r) => r.size)) : p.end.size;
}

function mainRun(p: TextPara): RunStyle {
  return p.runs.reduce<RunStyle>((m, r) => (r.size > m.size ? r : m), p.runs[0] ?? p.end);
}

/**
 * The text frame's content as DOM: one div per paragraph inside a column that is anchored in the
 * frame. `w` and `h` are the text area in px (shape box minus insets).
 */
export function renderText(t: TextInfo, w: number, h: number, o: TextOptions): HTMLElement {
  // (w and h are swapped below for turned text.)
  const body = t.body;
  const fit = o.fit ?? { fontScale: body.autofit === "norm" ? body.fontScale : 1, lnSpcReduction: body.autofit === "norm" ? body.lnSpcReduction : 0 };
  const outer = document.createElement("div");
  outer.className = "txt";
  const anchor = body.anchor === "ctr" ? "center" : body.anchor === "b" ? "flex-end" : "flex-start";
  // Vertical text (Alle Texte um 90° / 270° drehen, and the vertical placeholders): PowerPoint lays the
  // text out across the frame's height and turns the whole frame; Latin text is not stacked.
  const turn = body.vert === "vert" || body.vert === "eaVert" || body.vert === "mongolianVert" ? 90 : body.vert === "vert270" ? -90 : 0;
  if (turn) { const t2 = w; w = h; h = t2; }
  outer.style.cssText = css({ position: "absolute", left: "0", top: "0", width: `${w}px`, height: `${h}px`, display: "flex", "flex-direction": "column", "justify-content": anchor,
    translate: turn ? `${(h - w) / 2}px ${(w - h) / 2}px` : undefined, rotate: turn ? `${turn}deg` : undefined });
  if (turn) outer.dataset.turn = String(turn);
  const col = document.createElement("div");
  col.className = "txt-col";
  col.style.cssText = css({ width: body.wrap ? `${w}px` : "max-content", "align-self": body.anchorCtr ? "center" : body.wrap ? undefined : "flex-start",
    "column-count": body.numCol > 1 ? body.numCol : undefined, "column-gap": body.numCol > 1 ? `${body.spcCol / EMU_PX}px` : undefined,
    "writing-mode": body.vert === "wordArtVert" ? "vertical-rl" : undefined, "text-orientation": body.vert === "wordArtVert" ? "upright" : undefined });
  outer.append(col);
  const showPrompt = t.empty && o.prompt && t.prompt;
  const counters: number[] = [];
  t.paras.forEach((p, pi) => {
    const div = document.createElement("div");
    div.className = "para";
    div.dataset.p = String(pi);
    const lh = lineHeight(p, fit.fontScale, fit.lnSpcReduction);
    // Space before and after stays as set under Autofit (PowerPoint's read-back: 10 pt at 55 %).
    const bef = p.spcBefPts !== undefined ? p.spcBefPts * PT_PX : (p.spcBefPct ?? 0) * lh;
    const aft = p.spcAftPts !== undefined ? p.spcAftPts * PT_PX : (p.spcAftPct ?? 0) * lh;
    const marL = p.marL / EMU_PX, indent = p.indent / EMU_PX;
    div.style.cssText = css({
      position: "relative", "padding-left": `${marL}px`, "text-indent": p.bullet ? undefined : `${indent}px`,
      "text-align": p.algn === "ctr" ? "center" : p.algn === "r" ? "right" : p.algn === "just" || p.algn === "dist" ? "justify" : "left",
      "line-height": `${lh.toFixed(3)}px`, "font-family": `${fam(mainRun(p).font)}, Calibri, Carlito, sans-serif`, "font-size": `${(fitSize(largest(p), fit.fontScale) * PT_PX).toFixed(3)}px`, "margin-top": pi === 0 ? undefined : `${bef.toFixed(2)}px`, "margin-bottom": `${aft.toFixed(2)}px`,
      "white-space": body.wrap ? "pre-wrap" : "pre", "overflow-wrap": body.wrap ? "break-word" : undefined, "tab-size": `${p.defTab / EMU_PX}px`,
      direction: p.rtl ? "rtl" : undefined,
    });
    // Automatic numbers count paragraphs of the same level and type in a row.
    if (p.bullet?.kind === "auto") {
      counters.length = p.lvl + 1;
      counters[p.lvl] = (counters[p.lvl] ?? (p.bullet.startAt ?? 1) - 1) + 1;
    } else counters.length = Math.min(counters.length, p.lvl);
    const promptText = showPrompt ? t.prompt![Math.min(pi, t.prompt!.length - 1)] : null;
    const hasText = p.runs.some((r) => r.text.length);
    if (p.bullet && (hasText || promptText)) {
      const first = p.runs[0] ?? p.end;
      const bu = document.createElement("span");
      bu.className = "bullet";
      bu.contentEditable = "false";
      const size = p.bullet.sizePts ?? first.size * (p.bullet.sizePct ?? 1);
      bu.textContent = p.bullet.kind === "auto" ? autoNumber(p.bullet.autoType ?? "arabicPeriod", counters[p.lvl]) : p.bullet.char ?? "•";
      bu.style.cssText = css({ position: "absolute", left: `${marL + indent}px`, top: "0", "font-family": p.bullet.kind === "char" && p.bullet.font ? fam(p.bullet.font) : `${fam(first.font)}, Calibri, sans-serif`,
        "font-size": `${(size * fit.fontScale * PT_PX).toFixed(3)}px`, color: `#${p.bullet.color ?? first.color}`, "text-indent": "0", "white-space": "pre" });
      div.append(bu);
    }
    if (promptText !== null) {
      const sp = document.createElement("span");
      sp.className = "prompt";
      sp.dataset.o = "0";
      sp.style.cssText = runCss(p.end, fit.fontScale);
      sp.textContent = promptText;
      div.append(sp);
    } else {
      let off = 0;
      for (const r of p.runs) {
        if (r.kind === "math") { div.append(mathSpan(r, fit.fontScale)); continue; }
        const sp = document.createElement("span");
        sp.dataset.o = String(off);
        sp.className = r.kind === "br" ? "br" : "run";
        if (r.link) sp.dataset.link = r.link;
        sp.style.cssText = runCss(r, fit.fontScale);
        // Symbol characters (U+F0xx) in their symbol font (a:sym), as PowerPoint draws them.
        if (r.symFont && Array.from(r.text).some((c) => c.charCodeAt(0) >= 0xf000 && c.charCodeAt(0) <= 0xf0ff)) sp.style.fontFamily = fam(r.symFont);
        const text = r.kind === "br" ? "\n" : r.kind === "fld" ? fieldText(r, o) : r.text;
        if (o.selection) appendSelected(sp, text, pi, off, o.selection);
        else sp.textContent = text;
        div.append(sp);
        off += r.kind === "br" ? 1 : r.text.length;
      }
      // The paragraph mark: keeps an empty paragraph one line high at its own size.
      const end = document.createElement("span");
      end.className = "end";
      end.dataset.o = String(off);
      end.style.cssText = runCss(p.end, fit.fontScale);
      end.textContent = "\u200b";
      div.append(end);
    }
    col.append(div);
  });
  return outer;
}

/** Fills a run's span, marking the selected part. */
function appendSelected(sp: HTMLElement, text: string, pi: number, off: number, sel: NonNullable<TextOptions["selection"]>): void {
  const before = (p: number, o: number, q: { p: number; o: number }) => p < q.p || (p === q.p && o < q.o);
  let i = 0;
  while (i < text.length) {
    const inSel = !before(pi, off + i, sel.from) && before(pi, off + i, sel.to);
    let j = i + 1;
    while (j < text.length && (!before(pi, off + j, sel.from) && before(pi, off + j, sel.to)) === inSel) j++;
    const part = text.slice(i, j);
    if (inSel) { const m = document.createElement("mark"); m.className = "tsel"; m.textContent = part; sp.append(m); }
    else sp.append(document.createTextNode(part));
    i = j;
  }
}

/**
 * PowerPoint's Autofit steps (normAutofit): the font shrinks in steps and line spacing drops by
 * up to 20 % until the text fits. Returns the first step at which the content fits.
 */
export const AUTOFIT_STEPS: { fontScale: number; lnSpcReduction: number }[] = [
  { fontScale: 1, lnSpcReduction: 0 }, { fontScale: 1, lnSpcReduction: 0.1 }, { fontScale: 0.925, lnSpcReduction: 0.1 }, { fontScale: 0.925, lnSpcReduction: 0.2 },
  { fontScale: 0.85, lnSpcReduction: 0.2 }, { fontScale: 0.775, lnSpcReduction: 0.2 }, { fontScale: 0.7, lnSpcReduction: 0.2 }, { fontScale: 0.625, lnSpcReduction: 0.2 },
  { fontScale: 0.55, lnSpcReduction: 0.2 }, { fontScale: 0.475, lnSpcReduction: 0.2 }, { fontScale: 0.4, lnSpcReduction: 0.2 }, { fontScale: 0.325, lnSpcReduction: 0.2 },
  { fontScale: 0.25, lnSpcReduction: 0.2 },
];

export function textHeight(el: HTMLElement): number {
  const col = el.querySelector(".txt-col") as HTMLElement | null;
  return col ? col.scrollHeight : 0;
}

export function bodyInsets(b: BodyProps): { l: number; t: number; r: number; b: number } {
  return { l: b.l / EMU_PX, t: b.t / EMU_PX, r: b.r / EMU_PX, b: b.b / EMU_PX };
}

/** An equation (a14:m) drawn as MathML in Cambria Math; one the converter does not know shows its text. */
function mathSpan(r: TextRun, scale: number): HTMLElement {
  const sp = document.createElement("span");
  sp.className = "math";
  sp.style.cssText = runCss(r, scale) + ";font-family:'Cambria Math', serif";
  const seq = r.math ? ommlToSeq(r.math) : null;
  if (seq) sp.innerHTML = seqToMathml(seq, r.math?.localName === "oMathPara");
  else sp.textContent = r.math?.textContent ?? "";
  return sp;
}
