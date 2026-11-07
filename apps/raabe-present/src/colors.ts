// Colours as RaabePresent writes them (D-072): a hex colour ("4472C4") becomes a:srgbClr; a design colour
// from the palette ("accent1", or with a variant "accent1,20000,80000" for lumMod and lumOff) becomes
// a:schemeClr with PowerPoint's lumMod/lumOff, so it follows the design when the design changes.

import { PptxDom } from "@raabe/ooxml";

const { make } = PptxDom;
const HEX = /^[0-9A-F]{6}$/i;

/** The colour element for a hex colour or a design reference. */
export function colorEl(doc: Document, spec: string): Element {
  if (HEX.test(spec)) return make(doc, "a:srgbClr", { val: spec.toUpperCase() });
  const [val, mod, off] = spec.toLowerCase().split(",");
  const el = make(doc, "a:schemeClr", { val });
  if (mod) el.append(make(doc, "a:lumMod", { val: mod }));
  if (off) el.append(make(doc, "a:lumOff", { val: off }));
  return el;
}

/** The same as XML text (for the places that build their XML as text). */
export function colorXml(spec: string): string {
  if (HEX.test(spec)) return `<a:srgbClr val="${spec.toUpperCase()}"/>`;
  const [val, mod, off] = spec.toLowerCase().split(",");
  return `<a:schemeClr val="${val}">${mod ? `<a:lumMod val="${mod}"/>` : ""}${off ? `<a:lumOff val="${off}"/>` : ""}</a:schemeClr>`;
}

const SLOT: Record<string, string> = { bg1: "lt1", tx1: "dk1", bg2: "lt2", tx2: "dk2" };

/** The hex colour a reference stands for in a design (for the colour bar under a ribbon button). */
export function specHex(spec: string, colors: Readonly<Record<string, string>>): string {
  if (HEX.test(spec) || spec === "auto" || spec === "none") return spec;
  const [val, mod, off] = spec.toLowerCase().split(",");
  const base = colors[SLOT[val] ?? val] ?? "000000";
  const n = parseInt(base, 16);
  const m = mod ? Number(mod) / 100000 : 1, o = off ? Number(off) / 100000 : 0;
  // lumMod/lumOff act on the HSL lightness.
  let r = ((n >> 16) & 255) / 255, g = ((n >> 8) & 255) / 255, b = (n & 255) / 255;
  const max = Math.max(r, g, b), min = Math.min(r, g, b);
  let h = 0, s = 0, l = (max + min) / 2;
  if (max !== min) {
    const d = max - min;
    s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
    h = max === r ? (g - b) / d + (g < b ? 6 : 0) : max === g ? (b - r) / d + 2 : (r - g) / d + 4;
    h /= 6;
  }
  l = Math.min(1, Math.max(0, l * m + o));
  const f = (p: number, q: number, t: number) => { t = (t + 1) % 1; return t < 1 / 6 ? p + (q - p) * 6 * t : t < 1 / 2 ? q : t < 2 / 3 ? p + (q - p) * (2 / 3 - t) * 6 : p; };
  if (s === 0) r = g = b = l;
  else { const q = l < 0.5 ? l * (1 + s) : l + s - l * s, p = 2 * l - q; r = f(p, q, h + 1 / 3); g = f(p, q, h); b = f(p, q, h - 1 / 3); }
  return [r, g, b].map((x) => Math.round(x * 255).toString(16).padStart(2, "0")).join("").toUpperCase();
}
