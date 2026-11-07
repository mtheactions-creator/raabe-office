// Bildformat > Korrekturen (Schärfen/Weichzeichnen), Farbe (Farbsättigung, Farbton) and Künstlerische Effekte
// (D-072). Recorded in PowerPoint (probe_powerpoint_picfx.py): the effects are a list in the picture's
// a:blip/a:extLst (a14:imgProps > a14:imgLayer > a14:imgEffect), the untouched original is kept behind the
// layer's relationship and the picture itself (a:blip r:embed) becomes the rendered result. We do the same, so
// PowerPoint sees its own effects and every other program the rendered picture.
//
// Farbsättigung, Farbton and Schärfen/Weichzeichnen follow PowerPoint's measured results (saturation scales HSL
// saturation as s·k/(1+(k−1)·s); the colour temperature is PowerPoint's per-channel curve fitted on a grey ramp;
// soften is a small Gaussian, sharpen an unsharp mask). The 22 artistic effects are our own drawings of the idea
// behind each name (PowerPoint's filters are not documented); PowerPoint redraws them its way when it edits.

import { PptxDom } from "@raabe/ooxml";
import { closePopup, h, type MenuEntry } from "@raabe/shell";
import type { PresentApp } from "./app";

const { kid, make } = PptxDom;
const NS_A14 = "http://schemas.microsoft.com/office/drawing/2010/main";
const EXT_IMGPROPS = "{BEBA8EAE-BF5A-486C-A8C5-ECC9F3942E4B}";
const RT_IMAGE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";
const RT_HDPHOTO = "http://schemas.microsoft.com/office/2007/relationships/hdphoto";

/** Freistellen's marks (a14:foregroundMark / backgroundMark): lines in 1/1000 % of the picture. */
export interface BgMark { keep: boolean; x1: number; y1: number; x2: number; y2: number }
export interface ImgEffect { name: string; attrs: Record<string, string>; marks?: BgMark[] }

/** The artistic effects in the order of PowerPoint's gallery, with its German names. */
export const ARTISTIC: [string, string][] = [
  ["artisticMarker", "Markierung"], ["artisticPencilGrayscale", "Bleistift: Graustufen"], ["artisticPencilSketch", "Bleistiftskizze"],
  ["artisticLineDrawing", "Strichzeichnung"], ["artisticChalkSketch", "Kreideskizze"], ["artisticPaintStrokes", "Farbstriche"],
  ["artisticPaintBrush", "Pinsel"], ["artisticGlowDiffused", "Leuchteffekt: diffus"], ["artisticBlur", "Weichzeichnen"],
  ["artisticLightScreen", "Heller Bildschirm"], ["artisticWatercolorSponge", "Wasserfarbe: Schwamm"], ["artisticFilmGrain", "Filmkörnung"],
  ["artisticMosiaicBubbles", "Mosaik: Tupfen"], ["artisticGlass", "Glas"], ["artisticCement", "Zement"], ["artisticTexturizer", "Strukturierung"],
  ["artisticCrisscrossEtching", "Kreuzschraffur: Radierung"], ["artisticPastelsSmooth", "Pastellkreide: weich"], ["artisticPlasticWrap", "Plastikfolie"],
  ["artisticCutout", "Silhouette"], ["artisticPhotocopy", "Fotokopie"], ["artisticGlowEdges", "Leuchteffekt: Ränder"],
];

/** PowerPoint's colour temperature as out = a·in + b per channel, measured at these temperatures (K). */
const TEMP_FIT: [number, number, number, number, number, number, number][] = [
  [1500, 0.6579, -7.45, 1.3125, 10.66, 21.5397, 104.82], [2500, 0.7328, -6.87, 1.0942, 3.61, 1.9938, 25.4],
  [3500, 0.8412, -5.11, 1.0385, 1.52, 1.3837, 12.8], [4000, 0.8808, -4.14, 1.0245, 1.05, 1.2596, 9.16],
  [4700, 0.9245, -2.75, 1.0124, 0.61, 1.1482, 5.46], [5300, 0.9546, -1.74, 1.0088, 0.1, 1.083, 3.23],
  [5900, 0.98, -0.86, 1.0058, -0.15, 1.0362, 1.4], [6500, 1, 0, 1, 0, 1, 0],
  [7200, 1.0203, 0.81, 0.9946, 0.34, 0.9669, -1.27], [8000, 1.0399, 1.56, 0.9952, -0.11, 0.937, -2.35],
  [8800, 1.0561, 2.19, 0.996, -0.27, 0.9127, -3.12], [10000, 1.0757, 2.97, 0.9929, -0.11, 0.8851, -4.03],
  [11200, 1.0915, 3.54, 0.9917, -0.09, 0.8637, -4.59], [11500, 1.0954, 3.58, 0.9915, -0.09, 0.8589, -4.69],
];

// ----- the XML ------------------------------------------------------------------------------------------------

function imgLayer(blip: Element): Element | undefined {
  const ext = Array.from(kid(blip, "extLst")?.children ?? []).find((e) => e.getAttribute("uri") === EXT_IMGPROPS);
  return ext?.getElementsByTagNameNS(NS_A14, "imgLayer")[0];
}

/** The effects of a picture's blip, in their order. */
export function effectsOf(blip: Element): ImgEffect[] {
  const layer = imgLayer(blip);
  if (!layer) return [];
  return Array.from(layer.getElementsByTagNameNS(NS_A14, "imgEffect")).flatMap((e) => Array.from(e.children).map((c) => ({
    name: c.localName, attrs: Object.fromEntries(Array.from(c.attributes).map((a) => [a.name, a.value])),
    ...(c.children.length ? { marks: Array.from(c.children).filter((m) => m.localName.endsWith("Mark")).map((m) => ({ keep: m.localName === "foregroundMark",
      x1: Number(m.getAttribute("x1") ?? 0), y1: Number(m.getAttribute("y1") ?? 0), x2: Number(m.getAttribute("x2") ?? 0), y2: Number(m.getAttribute("y2") ?? 0) })) } : {}),
  })));
}

function writeEffects(blip: Element, effects: ImgEffect[], originalRid: string | null): void {
  const doc = blip.ownerDocument;
  let extLst = kid(blip, "extLst");
  const old = Array.from(extLst?.children ?? []).find((e) => e.getAttribute("uri") === EXT_IMGPROPS);
  old?.remove();
  if (!effects.length || !originalRid) { if (extLst && !extLst.children.length) extLst.remove(); return; }
  if (!extLst) { extLst = make(doc, "a:extLst"); blip.append(extLst); }
  const ext = make(doc, "a:ext", { uri: EXT_IMGPROPS });
  const props = doc.createElementNS(NS_A14, "a14:imgProps");
  const layer = doc.createElementNS(NS_A14, "a14:imgLayer");
  layer.setAttributeNS(PptxDom.NS_R, "r:embed", originalRid);
  for (const fx of effects) {
    const wrap = doc.createElementNS(NS_A14, "a14:imgEffect");
    const el = doc.createElementNS(NS_A14, `a14:${fx.name}`);
    for (const [k, v] of Object.entries(fx.attrs)) el.setAttribute(k, v);
    for (const m of fx.marks ?? []) {
      const me = doc.createElementNS(NS_A14, m.keep ? "a14:foregroundMark" : "a14:backgroundMark");
      for (const k of ["x1", "y1", "x2", "y2"] as const) me.setAttribute(k, String(Math.round(m[k])));
      el.append(me);
    }
    wrap.append(el);
    layer.append(wrap);
  }
  props.append(layer);
  ext.append(props);
  // PowerPoint keeps imgProps first in the blip's extLst.
  extLst.insertBefore(ext, extLst.firstElementChild);
}

/** The order effects are drawn in: the artistic effect, then sharpness, temperature and saturation. */
const ORDER = (n: string) => (n === "backgroundRemoval" ? -1 : n.startsWith("artistic") ? 0 : n === "sharpenSoften" ? 1 : n === "colorTemperature" ? 2 : n === "saturation" ? 3 : 4);

// ----- applying -----------------------------------------------------------------------------------------------

type Group = "artistic" | "sharpenSoften" | "colorTemperature" | "saturation" | "backgroundRemoval";

/** Sets (or with null removes) one kind of effect on the selected pictures and redraws them. */
export async function setEffect(app: PresentApp, group: Group, fx: ImgEffect | null, label: string): Promise<void> {
  const path = app.slidePath();
  if (!path) return;
  const pptx = app.deck.pptx;
  const views = app.editor.selectedViews.filter((v) => v.kind === "pic");
  type Job = { blip: Element; effects: ImgEffect[]; originalRid: string | null; png: Uint8Array | null };
  const jobs: Job[] = [];
  for (const v of views) {
    const blip = PptxDom.at(app.editor.shapeEl(v.id)!, "blipFill/blip");
    if (!blip) continue;
    const layer = imgLayer(blip);
    const before = effectsOf(blip);
    const effects = before.filter((e) => (group === "artistic" ? !e.name.startsWith("artistic") : e.name !== group));
    if (fx) effects.push(fx);
    effects.sort((a, b) => ORDER(a.name) - ORDER(b.name));
    // The original: the layer's picture, else the picture itself (no effects yet).
    const originalRid = layer?.getAttributeNS(PptxDom.NS_R, "embed") ?? blip.getAttributeNS(PptxDom.NS_R, "embed");
    const originalPart = originalRid ? pptx.target(path, originalRid) : undefined;
    let source = originalPart ? pptx.files[originalPart] : undefined;
    let draw = effects;
    // PowerPoint keeps its originals as HD Photo (.wdp), which a browser cannot read: draw the change on the
    // picture as it is now (PowerPoint redraws from the original when it opens the file).
    if (originalPart && /\.wdp$/i.test(originalPart)) {
      const cur = blip.getAttributeNS(PptxDom.NS_R, "embed");
      const curPart = cur ? pptx.target(path, cur) : undefined;
      source = curPart ? pptx.files[curPart] : undefined;
      draw = effects.filter((e) => !before.some((b) => b.name === e.name && JSON.stringify(b.attrs) === JSON.stringify(e.attrs)));
    }
    if (!source || !originalRid) continue;
    const png = effects.length ? await render(source, typeOf(originalPart ?? ""), draw) : null;
    jobs.push({ blip, effects, originalRid, png });
  }
  if (!jobs.length) return;
  app.editor.edit(label, () => {
    for (const j of jobs) {
      const layer = imgLayer(j.blip);
      if (!j.effects.length) {
        // No effects left: the original becomes the picture again (unless PowerPoint's .wdp, which stays layered).
        const part = j.originalRid ? pptx.target(path, j.originalRid) : undefined;
        if (layer && part && !/\.wdp$/i.test(part)) {
          j.blip.setAttributeNS(PptxDom.NS_R, "r:embed", pptx.link(path, RT_IMAGE, part));
          writeEffects(j.blip, [], null);
        } else if (layer) writeEffects(j.blip, [], null);
        continue;
      }
      let layerRid = j.originalRid;
      if (!layer) {
        // First effect: the original moves behind the layer's relationship, as PowerPoint keeps it.
        const part = pptx.target(path, j.originalRid!)!;
        layerRid = pptx.link(path, RT_HDPHOTO, part);
      }
      writeEffects(j.blip, j.effects, layerRid);
      if (j.png) {
        const media = pptx.freeName("ppt/media/image#.png");
        pptx.put(media, j.png);
        pptx.setType(media, "image/png", true);
        j.blip.setAttributeNS(PptxDom.NS_R, "r:embed", pptx.link(path, RT_IMAGE, media));
      }
    }
  });
}

function typeOf(part: string): string {
  return /\.jpe?g$/i.test(part) ? "image/jpeg" : /\.gif$/i.test(part) ? "image/gif" : /\.bmp$/i.test(part) ? "image/bmp" : "image/png";
}

async function render(bytes: Uint8Array, type: string, effects: ImgEffect[]): Promise<Uint8Array | null> {
  const url = URL.createObjectURL(new Blob([bytes as BlobPart], { type }));
  try {
    const img = await new Promise<HTMLImageElement | null>((res) => { const i = new Image(); i.onload = () => res(i); i.onerror = () => res(null); i.src = url; });
    if (!img) return null;
    const c = document.createElement("canvas");
    c.width = img.naturalWidth; c.height = img.naturalHeight;
    const g = c.getContext("2d", { willReadFrequently: true });
    if (!g) return null;
    g.drawImage(img, 0, 0);
    let px = g.getImageData(0, 0, c.width, c.height);
    for (const fx of effects) px = applyEffect(px, fx);
    g.putImageData(px, 0, 0);
    const blob = await new Promise<Blob | null>((r) => c.toBlob(r, "image/png"));
    return blob ? new Uint8Array(await blob.arrayBuffer()) : null;
  } finally { URL.revokeObjectURL(url); }
}

// ----- the pixel work (exported for the tests) ----------------------------------------------------------------

export interface Pixels { width: number; height: number; data: Uint8ClampedArray }

const num = (fx: ImgEffect, k: string, def: number) => (fx.attrs[k] !== undefined ? Number(fx.attrs[k]) : def);

export function applyEffect<P extends Pixels>(p: P, fx: ImgEffect): P {
  const d = p.data;
  switch (fx.name) {
    case "saturation": saturate(d, num(fx, "sat", 100000) / 100000); return p;
    case "colorTemperature": temperature(d, num(fx, "colorTemp", 6500)); return p;
    case "backgroundRemoval": {
      const f = (k: string, def: number) => num(fx, k, def) / 100000;
      removeBackground(p, { t: f("t", 10000), b: f("b", 90000), l: f("l", 10000), r: f("r", 90000) }, fx.marks ?? []);
      return p;
    }
    case "sharpenSoften": {
      const a = num(fx, "amount", 0) / 100000;
      if (a < 0) d.set(blur(p, 1.6 * -a));
      else if (a > 0) { const b = blur(p, 0.8); for (let i = 0; i < d.length; i++) if (i % 4 !== 3) d[i] = d[i] + 1.38 * a * (d[i] - b[i]); }
      return p;
    }
    default: artistic(p, fx); return p;
  }
}

/**
 * Hintergrund entfernen (our own method; PowerPoint's is not documented): everything outside the rectangle
 * (PowerPoint's default: 10 % in from each side) is background; from there the background grows into
 * neighbouring pixels of a similar colour that also look like the background's colours. What is left keeps
 * its pixels, the rest becomes transparent (hard edges, as in PowerPoint's result).
 */
export function removeBackground(p: Pixels, rect: { t: number; b: number; l: number; r: number }, marks: BgMark[] = []): void {
  const { width: w, height: h, data: d } = p;
  const n = w * h;
  const bg = new Uint8Array(n);
  // Freistellen's marks: pixels along a "keep" line never become background, along a "remove" line they are.
  const keepMask = new Uint8Array(n), removeMask = new Uint8Array(n);
  const rad = Math.max(1, Math.round(Math.min(w, h) / 100));
  for (const m of marks) {
    const ax = (m.x1 / 100000) * w, ay = (m.y1 / 100000) * h, bx = (m.x2 / 100000) * w, by = (m.y2 / 100000) * h;
    const steps = Math.max(1, Math.ceil(Math.hypot(bx - ax, by - ay)));
    for (let k = 0; k <= steps; k++) {
      const cx = Math.round(ax + ((bx - ax) * k) / steps), cy = Math.round(ay + ((by - ay) * k) / steps);
      for (let yy = cy - rad; yy <= cy + rad; yy++) for (let xx = cx - rad; xx <= cx + rad; xx++) {
        if (xx < 0 || yy < 0 || xx >= w || yy >= h) continue;
        (m.keep ? keepMask : removeMask)[yy * w + xx] = 1;
      }
    }
  }
  const queue = new Int32Array(n);
  let qh = 0, qt = 0;
  const x0 = Math.floor(rect.l * w), x1 = Math.ceil(rect.r * w), y0 = Math.floor(rect.t * h), y1 = Math.ceil(rect.b * h);
  const outside = (x: number, y: number) => x < x0 || x >= x1 || y < y0 || y >= y1;
  const full = x0 <= 0 && y0 <= 0 && x1 >= w && y1 >= h;
  for (let y = 0; y < h; y++) for (let x = 0; x < w; x++) {
    const seed = full ? x === 0 || y === 0 || x === w - 1 || y === h - 1 : outside(x, y);
    const i = y * w + x;
    if ((seed || removeMask[i]) && !keepMask[i]) { bg[i] = 1; queue[qt++] = i; }
  }
  // The background's colours: the seeds' colours in coarse bins, the 16 most frequent.
  const bins = new Map<number, { n: number; r: number; g: number; b: number }>();
  for (let k = 0; k < qt; k++) {
    const i = queue[k] * 4, key = ((d[i] >> 4) << 8) | ((d[i + 1] >> 4) << 4) | (d[i + 2] >> 4);
    const e = bins.get(key) ?? { n: 0, r: 0, g: 0, b: 0 };
    e.n++; e.r += d[i]; e.g += d[i + 1]; e.b += d[i + 2];
    bins.set(key, e);
  }
  const model = [...bins.values()].sort((a, b) => b.n - a.n).slice(0, 16).map((e) => [e.r / e.n, e.g / e.n, e.b / e.n]);
  const nearModel = (i: number) => model.some(([r, g, b]) => Math.abs(d[i] - r) + Math.abs(d[i + 1] - g) + Math.abs(d[i + 2] - b) < 90);
  const near = (i: number, j: number) => Math.abs(d[i] - d[j]) + Math.abs(d[i + 1] - d[j + 1]) + Math.abs(d[i + 2] - d[j + 2]) < 30;
  while (qh < qt) {
    const q = queue[qh++], x = q % w, y = (q - x) / w;
    for (const nb of [x > 0 ? q - 1 : -1, x < w - 1 ? q + 1 : -1, y > 0 ? q - w : -1, y < h - 1 ? q + w : -1]) {
      if (nb < 0 || bg[nb]) continue;
      if (keepMask[nb]) continue;
      if (near(q * 4, nb * 4) && (nearModel(nb * 4) || removeMask[q])) { bg[nb] = 1; queue[qt++] = nb; }
    }
  }
  // The kept areas keep a pixel of their edge (anti-aliased border pixels belong to the object, as in PowerPoint).
  const keep = new Uint8Array(n);
  for (let i = 0; i < n; i++) {
    if (!bg[i]) continue;
    const x = i % w;
    if ((x > 0 && !bg[i - 1]) || (x < w - 1 && !bg[i + 1]) || (i >= w && !bg[i - w]) || (i + w < n && !bg[i + w])) keep[i] = 1;
  }
  for (let i = 0; i < n; i++) if (bg[i] && !keep[i]) d[i * 4 + 3] = 0;
}

function saturate(d: Uint8ClampedArray, k: number): void {
  for (let i = 0; i < d.length; i += 4) {
    const r = d[i], g = d[i + 1], b = d[i + 2];
    const mx = Math.max(r, g, b), mn = Math.min(r, g, b), l = (mx + mn) / 2;
    if (mx === mn) continue;
    const s = l <= 127.5 ? (mx - mn) / (mx + mn) : (mx - mn) / (510 - mx - mn);
    const f = k / (1 + (k - 1) * s);
    d[i] = l + (r - l) * f; d[i + 1] = l + (g - l) * f; d[i + 2] = l + (b - l) * f;
  }
}

/** PowerPoint's measured curve at `t` K (interpolated between the measured temperatures). */
export function temperatureFit(t: number): number[] {
  const k = Math.min(Math.max(t, TEMP_FIT[0][0]), TEMP_FIT[TEMP_FIT.length - 1][0]);
  let i = 0;
  while (i < TEMP_FIT.length - 2 && TEMP_FIT[i + 1][0] < k) i++;
  const a = TEMP_FIT[i], b = TEMP_FIT[i + 1], f = (k - a[0]) / (b[0] - a[0]);
  return a.slice(1).map((v, j) => v + (b[j + 1] - v) * f);
}

function temperature(d: Uint8ClampedArray, t: number): void {
  const [ar, br, ag, bg, ab, bb] = temperatureFit(t);
  for (let i = 0; i < d.length; i += 4) { d[i] = ar * d[i] + br; d[i + 1] = ag * d[i + 1] + bg; d[i + 2] = ab * d[i + 2] + bb; }
}

/** A Gaussian blur (three box passes) of the colour channels; returns new data. */
export function blur(p: Pixels, sigma: number): Uint8ClampedArray {
  const { width: w, height: h } = p;
  const out = new Uint8ClampedArray(p.data);
  if (sigma <= 0.05) return out;
  const r = Math.max(1, Math.round(Math.sqrt((12 * sigma * sigma) / 3 + 1) / 2 - 0.5)) || 1;
  const exact = sigma < 1;
  const tmp = new Float32Array(p.data.length);
  for (let pass = 0; pass < (exact ? 1 : 3); pass++) {
    const rad = exact ? 1 : r;
    // Below one pixel: one 3-tap pass with the Gaussian's weights.
    const wsum = exact ? 1 + 2 * Math.exp(-1 / (2 * sigma * sigma)) : 2 * rad + 1;
    const wn = exact ? Math.exp(-1 / (2 * sigma * sigma)) : 1;
    for (let y = 0; y < h; y++) for (let x = 0; x < w; x++) for (let c = 0; c < 3; c++) {
      let s = 0;
      for (let k = -rad; k <= rad; k++) { const xx = Math.min(w - 1, Math.max(0, x + k)); s += out[(y * w + xx) * 4 + c] * (k === 0 ? 1 : wn); }
      tmp[(y * w + x) * 4 + c] = s / wsum;
    }
    for (let y = 0; y < h; y++) for (let x = 0; x < w; x++) for (let c = 0; c < 3; c++) {
      let s = 0;
      for (let k = -rad; k <= rad; k++) { const yy = Math.min(h - 1, Math.max(0, y + k)); s += tmp[(yy * w + x) * 4 + c] * (k === 0 ? 1 : wn); }
      out[(y * w + x) * 4 + c] = s / wsum;
    }
  }
  return out;
}

function gray(d: ArrayLike<number>, i: number): number { return 0.299 * d[i] + 0.587 * d[i + 1] + 0.114 * d[i + 2]; }

/** Edge strength (Sobel on brightness), 0..~1020. */
function edges(p: Pixels): Float32Array {
  const { width: w, height: h, data: d } = p;
  const L = new Float32Array(w * h);
  for (let i = 0; i < w * h; i++) L[i] = gray(d, i * 4);
  const at = (x: number, y: number) => L[Math.min(h - 1, Math.max(0, y)) * w + Math.min(w - 1, Math.max(0, x))];
  const e = new Float32Array(w * h);
  for (let y = 0; y < h; y++) for (let x = 0; x < w; x++) {
    const gx = at(x + 1, y - 1) + 2 * at(x + 1, y) + at(x + 1, y + 1) - at(x - 1, y - 1) - 2 * at(x - 1, y) - at(x - 1, y + 1);
    const gy = at(x - 1, y + 1) + 2 * at(x, y + 1) + at(x + 1, y + 1) - at(x - 1, y - 1) - 2 * at(x, y - 1) - at(x + 1, y - 1);
    e[y * w + x] = Math.hypot(gx, gy);
  }
  return e;
}

/** A fixed pseudo-random sequence, so an effect looks the same each time it is drawn. */
function noise(seed: number): () => number {
  let a = seed >>> 0;
  return () => { a = (a + 0x6d2b79f5) >>> 0; let t = a; t = Math.imul(t ^ (t >>> 15), t | 1); t ^= t + Math.imul(t ^ (t >>> 7), t | 61); return ((t ^ (t >>> 14)) >>> 0) / 4294967296; };
}

function posterize(d: Uint8ClampedArray, levels: number): void {
  const step = 255 / (levels - 1);
  for (let i = 0; i < d.length; i++) if (i % 4 !== 3) d[i] = Math.round(d[i] / step) * step;
}

/** PowerPoint's default Transparenz of the effects that have one (recorded, in 1/1000 %). */
const DEFAULT_TRANS: Record<string, number> = { artisticCrisscrossEtching: 75000, artisticGlowEdges: 15000, artisticLineDrawing: 25000, artisticPhotocopy: 30000 };

function artistic(p: Pixels, fx: ImgEffect): void {
  const { width: w, height: h, data: d } = p;
  const n = w * h;
  const original = new Uint8ClampedArray(d);
  const rnd = noise(w * 7919 + h);
  const scale = Math.max(1, Math.min(w, h) / 240);
  const each = (fn: (k: number, x: number, y: number) => void) => { for (let y = 0; y < h; y++) for (let x = 0; x < w; x++) fn((y * w + x) * 4, x, y); };
  const mul = (k: number, f: number) => { d[k] *= f; d[k + 1] *= f; d[k + 2] *= f; };
  const add = (k: number, v: number) => { d[k] += v; d[k + 1] += v; d[k + 2] += v; };
  const toward = (k: number, t: number, f: number) => { d[k] += (t - d[k]) * f; d[k + 1] += (t - d[k + 1]) * f; d[k + 2] += (t - d[k + 2]) * f; };
  const setGray = (k: number, v: number) => { d[k] = d[k + 1] = d[k + 2] = v; };
  /** A noise field, blurred (and stretched along one direction) into a texture of values about 0. */
  const field = (sx: number, sy: number, seed = 1): Float32Array => {
    const r = noise(seed * 104729 + w), f = new Float32Array(n);
    for (let i = 0; i < n; i++) f[i] = r() - 0.5;
    const g = new Float32Array(n);
    const rx = Math.max(0, Math.round(sx * scale)), ry = Math.max(0, Math.round(sy * scale));
    for (let y = 0; y < h; y++) for (let x = 0; x < w; x++) { let s = 0, c = 0; for (let k = -rx; k <= rx; k++) { const xx = x + k; if (xx >= 0 && xx < w) { s += f[y * w + xx]; c++; } } g[y * w + x] = s / c; }
    for (let y = 0; y < h; y++) for (let x = 0; x < w; x++) { let s = 0, c = 0; for (let k = -ry; k <= ry; k++) { const yy = y + k; if (yy >= 0 && yy < h) { s += g[yy * w + x]; c++; } } f[y * w + x] = s / c; }
    let m = 0;
    for (let i = 0; i < n; i++) m = Math.max(m, Math.abs(f[i]));
    if (m > 0) for (let i = 0; i < n; i++) f[i] /= 2 * m;
    return f;
  };
  /** Light from the top left on a height field: the relief of a texture. */
  const relief = (f: Float32Array): Float32Array => {
    const e = new Float32Array(n);
    for (let y = 1; y < h; y++) for (let x = 1; x < w; x++) e[y * w + x] = f[y * w + x] - f[(y - 1) * w + x - 1];
    let m = 0;
    for (let i = 0; i < n; i++) m = Math.max(m, Math.abs(e[i]));
    if (m > 0) for (let i = 0; i < n; i++) e[i] /= m;
    return e;
  };
  switch (fx.name) {
    case "artisticBlur": d.set(blur(p, num(fx, "radius", 10) * 0.5 * scale)); break;
    case "artisticMarker": {
      // Nearly the picture, its colours slightly flattened.
      d.set(blur(p, 0.6 * scale));
      posterize(d, 32);
      break;
    }
    case "artisticPaintBrush": {
      // Vertical brush marks: each pixel takes the colour at the start of its short column band.
      const band = Math.max(2, Math.round(num(fx, "brushSize", 2) * 2 * scale));
      const src = blur(p, 0.8 * scale);
      each((k, x, y) => { const j = (y * w + (x - (x % band))) * 4; d[k] = src[j]; d[k + 1] = src[j + 1]; d[k + 2] = src[j + 2]; });
      break;
    }
    case "artisticPaintStrokes": {
      d.set(blur(p, 1.2 * scale));
      each((k) => { if (gray(d, k) < 95) setGray(k, 0); });
      break;
    }
    case "artisticLineDrawing": {
      // The picture with light strokes scratched over it.
      const f = field(2, 0, 3);
      each((k) => { const v = f[k / 4]; if (v > 0.12) toward(k, 255, Math.min(0.8, (v - 0.12) * 4)); });
      break;
    }
    case "artisticCrisscrossEtching": {
      // Hatching in the picture's colours; PowerPoint's default lets 75 % of the original through.
      const sp = Math.max(3, Math.round(3 * scale));
      each((k, x, y) => { const v = gray(d, k); toward(k, 255, 0.4); if (((x + y) % sp === 0 && v < 220) || ((x - y + h * sp) % sp === 0 && v < 140)) mul(k, 0.55); });
      break;
    }
    case "artisticPastelsSmooth": {
      d.set(blur(p, 0.8 * scale));
      each((k) => { const v = (rnd() - 0.5) * 50; add(k, v); });
      break;
    }
    case "artisticPhotocopy": {
      // Washed-out colours, darker along the edges.
      const e = edges(p);
      each((k) => { toward(k, 255, 0.7); add(k, -Math.min(90, e[k / 4] * 0.25)); });
      break;
    }
    case "artisticPencilGrayscale": case "artisticChalkSketch": {
      // Grey with vertical pencil (or chalk) strokes.
      const f = field(0, 6, 5);
      const dark = fx.name === "artisticChalkSketch" ? 0.6 : 1;
      each((k) => { const v = gray(d, k) * dark; setGray(k, v * (1 + f[k / 4] * 0.9)); });
      break;
    }
    case "artisticPencilSketch": {
      // Light parts white, dark parts in a grain of pencil dots.
      each((k) => { const v = gray(d, k); setGray(k, v > 175 ? 255 : rnd() * 175 < v ? 255 : v * 0.6); });
      break;
    }
    case "artisticGlowDiffused": {
      each((k) => { const v = gray(d, k) / 255; toward(k, 255, Math.min(1, v * v * 1.1 + (rnd() - 0.5) * 0.2 * v)); });
      break;
    }
    case "artisticCutout": {
      // A few flat areas: brightness bands of the blurred picture, each in its mean colour.
      d.set(blur(p, 6 * scale));
      const bins = Math.max(2, num(fx, "numberOfShades", 2) + 3);
      const sum = Array.from({ length: bins }, () => [0, 0, 0, 0]);
      const bin = new Uint8Array(n);
      each((k) => { const b = Math.min(bins - 1, Math.floor((gray(d, k) / 256) * bins)); bin[k / 4] = b; sum[b][0] += d[k]; sum[b][1] += d[k + 1]; sum[b][2] += d[k + 2]; sum[b][3]++; });
      each((k) => { const s = sum[bin[k / 4]]; d[k] = s[0] / s[3]; d[k + 1] = s[1] / s[3]; d[k + 2] = s[2] / s[3]; });
      break;
    }
    case "artisticFilmGrain": {
      const amt = num(fx, "grainSize", 40) * 1.4;
      for (let i = 0; i < d.length; i += 4) for (let c = 0; c < 3; c++) d[i + c] += (rnd() - 0.5) * amt;
      break;
    }
    case "artisticMosiaicBubbles": {
      const cell = Math.max(3, Math.round(5 * scale));
      each((k, x, y) => { const rr = Math.hypot((x % cell) - cell / 2, (y % cell) - cell / 2) / (cell / 2); mul(k, rr > 0.95 ? 0.8 : 1); if (rnd() > 0.97) toward(k, 255, 0.6); });
      break;
    }
    case "artisticGlass": {
      const src = new Uint8ClampedArray(d);
      const fx1 = field(3, 3, 7), fy1 = field(3, 3, 11);
      const s = 8 * scale;
      each((k, x, y) => {
        const i = k / 4, sx = Math.min(w - 1, Math.max(0, Math.round(x + fx1[i] * s))), sy = Math.min(h - 1, Math.max(0, Math.round(y + fy1[i] * s))), j = (sy * w + sx) * 4;
        d[k] = src[j]; d[k + 1] = src[j + 1]; d[k + 2] = src[j + 2];
      });
      break;
    }
    case "artisticCement": case "artisticTexturizer": case "artisticLightScreen": {
      // A relief lit from the top left: cement blocks, a fine canvas, or a grid of small cushions.
      let f: Float32Array;
      if (fx.name === "artisticLightScreen") {
        const cell = Math.max(3, Math.round(num(fx, "gridSize", 4) * 2.5 * scale));
        f = new Float32Array(n);
        for (let y = 0; y < h; y++) for (let x = 0; x < w; x++) { const u = (x % cell) / cell - 0.5, v = (y % cell) / cell - 0.5; f[y * w + x] = 0.5 - Math.max(Math.abs(u), Math.abs(v)); }
      } else if (fx.name === "artisticCement") {
        const cell = Math.max(4, Math.round(10 * scale));
        const r = noise(31);
        const off = Array.from({ length: Math.ceil(h / cell) + 1 }, () => Math.floor(r() * cell));
        f = new Float32Array(n);
        for (let y = 0; y < h; y++) for (let x = 0; x < w; x++) { const row = Math.floor(y / cell), u = ((x + off[row]) % cell) / cell, v = (y % cell) / cell; f[y * w + x] = Math.min(u, 1 - u, v, 1 - v) < 0.12 ? -0.5 : 0.2; }
        const fn = field(1, 1, 13);
        for (let i = 0; i < n; i++) f[i] += fn[i] * 0.6;
      } else f = field(1, 1, 17);
      const e = relief(f);
      const k0 = fx.name === "artisticTexturizer" ? 0.35 : fx.name === "artisticLightScreen" ? 0.22 : 0.5;
      each((k) => { const v = e[k / 4] * k0; if (v > 0) toward(k, 255, v); else mul(k, 1 + v); });
      break;
    }
    case "artisticWatercolorSponge": {
      // Patches of colour with a sponge's soft edges.
      const f = field(4, 4, 19);
      d.set(blur(p, 1.5 * scale));
      posterize(d, 6);
      each((k) => add(k, f[k / 4] * 70));
      break;
    }
    case "artisticPlasticWrap": {
      d.set(blur(p, 1.5 * scale));
      const e = edges(p);
      for (let i = 0; i < n; i++) { const k = i * 4, f = Math.min(1, e[i] / 300); for (let c = 0; c < 3; c++) d[k + c] = d[k + c] * (1 - f) + 255 * f; }
      break;
    }
    case "artisticGlowEdges": {
      const e = edges(p);
      for (let i = 0; i < n; i++) { const k = i * 4, f = Math.min(1, e[i] / 200); for (let c = 0; c < 3; c++) d[k + c] = d[k + c] * Math.max(f, 0.08); }
      break;
    }
  }
  // The effect's Transparenz (trans, 0-100000) lets the original show through.
  const trans = num(fx, "trans", DEFAULT_TRANS[fx.name] ?? 0) / 100000;
  if (trans > 0) for (let i = 0; i < d.length; i++) if (i % 4 !== 3) d[i] = d[i] * (1 - trans) + original[i] * trans;
}

// ----- menus: galleries of the selected picture, as PowerPoint shows them ----------------------------------

type Tile = { label: string; group: Group; fx: ImgEffect | null };

/** The selected picture's original, small (for the gallery tiles). */
async function thumbSource(app: PresentApp, w: number, hgt: number): Promise<ImageData | null> {
  const path = app.slidePath();
  const v = app.editor.selectedViews.find((x) => x.kind === "pic");
  const blip = v && path ? PptxDom.at(app.editor.shapeEl(v.id)!, "blipFill/blip") : undefined;
  if (!blip || !path) return null;
  const pptx = app.deck.pptx;
  let rid = imgLayer(blip)?.getAttributeNS(PptxDom.NS_R, "embed") ?? blip.getAttributeNS(PptxDom.NS_R, "embed");
  let part = rid ? pptx.target(path, rid) : undefined;
  if (part && /\.wdp$/i.test(part)) { rid = blip.getAttributeNS(PptxDom.NS_R, "embed"); part = rid ? pptx.target(path, rid) : undefined; }
  const bytes = part ? pptx.files[part] : undefined;
  if (!bytes) return null;
  const url = URL.createObjectURL(new Blob([bytes as BlobPart], { type: typeOf(part!) }));
  try {
    const img = await new Promise<HTMLImageElement | null>((res) => { const i = new Image(); i.onload = () => res(i); i.onerror = () => res(null); i.src = url; });
    if (!img) return null;
    const c = document.createElement("canvas");
    c.width = w; c.height = hgt;
    const g = c.getContext("2d", { willReadFrequently: true })!;
    const sc = Math.max(w / img.naturalWidth, hgt / img.naturalHeight);
    g.drawImage(img, (w - img.naturalWidth * sc) / 2, (hgt - img.naturalHeight * sc) / 2, img.naturalWidth * sc, img.naturalHeight * sc);
    return g.getImageData(0, 0, w, hgt);
  } finally { URL.revokeObjectURL(url); }
}

function tileGrid(app: PresentApp, tiles: Tile[], cols: number): HTMLElement {
  const W = 72, H = 54;
  const grid = h("div", { class: "pfx-grid", style: `grid-template-columns:repeat(${cols}, ${W + 8}px)` });
  const canvases = tiles.map((t) => {
    const c = h("canvas", { width: String(W), height: String(H) }) as HTMLCanvasElement;
    grid.append(h("button", { class: "pfx-tile", role: "menuitem", "aria-label": t.label, "data-tip": t.label, onclick: () => { closePopup(); void setEffect(app, t.group, t.fx, t.label); } }, c));
    return c;
  });
  void thumbSource(app, W, H).then((src) => {
    if (!src) return;
    tiles.forEach((t, i) => {
      const px = new ImageData(new Uint8ClampedArray(src.data), W, H);
      if (t.fx) applyEffect(px, t.fx);
      canvases[i].getContext("2d")?.putImageData(px, 0, 0);
    });
  });
  return grid;
}

/** Schärfen/Weichzeichnen entries of Korrekturen (PowerPoint's labels). */
export function sharpenEntries(app: PresentApp): MenuEntry[] {
  const t = (amount: number, label: string): Tile => ({ label, group: "sharpenSoften", fx: amount ? { name: "sharpenSoften", attrs: { amount: String(amount) } } : null });
  return [{ heading: "Schärfen/Weichzeichnen" }, { block: () => tileGrid(app, [t(-50000, "Weichzeichnen: 50 %"), t(-25000, "Weichzeichnen: 25 %"), t(0, "Schärfen: 0 %"), t(25000, "Schärfen: 25 %"), t(50000, "Schärfen: 50 %")], 5) }];
}

/** Farbsättigung and Farbton entries of Farbe (PowerPoint's labels and values). */
export function colorEntries(app: PresentApp): MenuEntry[] {
  const sat = (pct: number): Tile => ({ label: `Sättigung: ${pct} %`, group: "saturation", fx: pct === 100 ? null : { name: "saturation", attrs: { sat: String(pct * 1000) } } });
  const temp = (k: number): Tile => ({ label: `Temperatur: ${k} K`, group: "colorTemperature", fx: k === 6500 ? null : { name: "colorTemperature", attrs: { colorTemp: String(k) } } });
  return [{ heading: "Farbsättigung" }, { block: () => tileGrid(app, [0, 33, 66, 100, 200, 300, 400].map(sat), 7) },
    { heading: "Farbton" }, { block: () => tileGrid(app, [4700, 5300, 5900, 6500, 7200, 8800, 11200].map(temp), 7) }];
}

/** Künstlerische Effekte (PowerPoint's 23 tiles and its options entry). */
export function artisticEntries(app: PresentApp): MenuEntry[] {
  const tiles: Tile[] = [{ label: "Keine", group: "artistic", fx: null }, ...ARTISTIC.map(([name, label]): Tile => ({ label, group: "artistic", fx: { name, attrs: {} } }))];
  return [{ block: () => tileGrid(app, tiles, 5) }, "-", { label: "Optionen für künstlerische Effekte...", run: () => void app.commands.run("draw.format") }];
}
