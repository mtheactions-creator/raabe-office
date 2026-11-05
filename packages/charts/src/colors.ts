// Diagrammentwurf > Farben ändern: Excel's colour schemes 10-26 as saved in its colour parts
// (tools/feel-lab/probe_excel_chartcolors.py, D-066). "cycle" runs through the colours and varies them
// in each further round; "withinLinear" shades one colour from dark to light across the series.

import { lum } from "./model";

export interface ColorScheme { meth: "cycle" | "withinLinear" | "withinLinearReversed"; colors: string[]; variations: { op: string; val: number }[][] }

export const COLOR_SCHEMES: Record<number, ColorScheme> = {"10": {"meth": "cycle", "colors": ["accent1", "accent2", "accent3", "accent4", "accent5", "accent6"], "variations": [[], [{"op": "lumMod", "val": 0.6}], [{"op": "lumMod", "val": 0.8}, {"op": "lumOff", "val": 0.2}], [{"op": "lumMod", "val": 0.8}], [{"op": "lumMod", "val": 0.6}, {"op": "lumOff", "val": 0.4}], [{"op": "lumMod", "val": 0.5}], [{"op": "lumMod", "val": 0.7}, {"op": "lumOff", "val": 0.3}], [{"op": "lumMod", "val": 0.7}], [{"op": "lumMod", "val": 0.5}, {"op": "lumOff", "val": 0.5}]]}, "11": {"meth": "cycle", "colors": ["accent1", "accent3", "accent5"], "variations": [[], [{"op": "lumMod", "val": 0.6}], [{"op": "lumMod", "val": 0.8}, {"op": "lumOff", "val": 0.2}], [{"op": "lumMod", "val": 0.8}], [{"op": "lumMod", "val": 0.6}, {"op": "lumOff", "val": 0.4}], [{"op": "lumMod", "val": 0.5}], [{"op": "lumMod", "val": 0.7}, {"op": "lumOff", "val": 0.3}], [{"op": "lumMod", "val": 0.7}], [{"op": "lumMod", "val": 0.5}, {"op": "lumOff", "val": 0.5}]]}, "12": {"meth": "cycle", "colors": ["accent2", "accent4", "accent6"], "variations": [[], [{"op": "lumMod", "val": 0.6}], [{"op": "lumMod", "val": 0.8}, {"op": "lumOff", "val": 0.2}], [{"op": "lumMod", "val": 0.8}], [{"op": "lumMod", "val": 0.6}, {"op": "lumOff", "val": 0.4}], [{"op": "lumMod", "val": 0.5}], [{"op": "lumMod", "val": 0.7}, {"op": "lumOff", "val": 0.3}], [{"op": "lumMod", "val": 0.7}], [{"op": "lumMod", "val": 0.5}, {"op": "lumOff", "val": 0.5}]]}, "13": {"meth": "cycle", "colors": ["accent6", "accent5", "accent4"], "variations": [[], [{"op": "lumMod", "val": 0.6}], [{"op": "lumMod", "val": 0.8}, {"op": "lumOff", "val": 0.2}], [{"op": "lumMod", "val": 0.8}], [{"op": "lumMod", "val": 0.6}, {"op": "lumOff", "val": 0.4}], [{"op": "lumMod", "val": 0.5}], [{"op": "lumMod", "val": 0.7}, {"op": "lumOff", "val": 0.3}], [{"op": "lumMod", "val": 0.7}], [{"op": "lumMod", "val": 0.5}, {"op": "lumOff", "val": 0.5}]]}, "14": {"meth": "withinLinear", "colors": ["accent1"], "variations": []}, "15": {"meth": "withinLinear", "colors": ["accent2"], "variations": []}, "16": {"meth": "withinLinear", "colors": ["accent3"], "variations": []}, "17": {"meth": "withinLinear", "colors": ["accent4"], "variations": []}, "18": {"meth": "withinLinear", "colors": ["accent5"], "variations": []}, "19": {"meth": "withinLinear", "colors": ["accent6"], "variations": []}, "20": {"meth": "cycle", "colors": ["dk1"], "variations": [[{"op": "tint", "val": 0.885}], [{"op": "tint", "val": 0.55}], [{"op": "tint", "val": 0.75}], [{"op": "tint", "val": 0.985}], [{"op": "tint", "val": 0.3}], [{"op": "tint", "val": 0.6}], [{"op": "tint", "val": 0.8}]]}, "21": {"meth": "withinLinearReversed", "colors": ["accent1"], "variations": []}, "22": {"meth": "withinLinearReversed", "colors": ["accent2"], "variations": []}, "23": {"meth": "withinLinearReversed", "colors": ["accent3"], "variations": []}, "24": {"meth": "withinLinearReversed", "colors": ["accent4"], "variations": []}, "25": {"meth": "withinLinearReversed", "colors": ["accent5"], "variations": []}, "26": {"meth": "withinLinearReversed", "colors": ["accent6"], "variations": []}};

/** sRGB channel (0-1) to linear light and back (DrawingML's tint and shade work in linear light). */
const toLin = (c: number) => (c <= 0.04045 ? c / 12.92 : ((c + 0.055) / 1.055) ** 2.4);
const toSrgb = (c: number) => (c <= 0.0031308 ? c * 12.92 : 1.055 * c ** (1 / 2.4) - 0.055);

function channels(hex: string): number[] {
  return [1, 3, 5].map((i) => parseInt(hex.slice(i, i + 2), 16) / 255);
}

function hex(ch: number[]): string {
  return "#" + ch.map((v) => Math.round(Math.max(0, Math.min(1, v)) * 255).toString(16).padStart(2, "0")).join("").toUpperCase();
}

/** DrawingML's tint (towards white) and shade (towards black) of a colour, in linear light. */
export function tint(color: string, t: number): string {
  return hex(channels(color).map((c) => toSrgb(toLin(c) * t + (1 - t))));
}
export function shade(color: string, s: number): string {
  return hex(channels(color).map((c) => toSrgb(toLin(c) * s)));
}

function apply(color: string, v: { op: string; val: number }[]): string {
  let mod = 1, off = 0, out = color;
  for (const x of v) {
    if (x.op === "lumMod") mod = x.val;
    else if (x.op === "lumOff") off = x.val;
    else if (x.op === "tint") out = tint(out, x.val);
    else if (x.op === "shade") out = shade(out, x.val);
  }
  return mod !== 1 || off !== 0 ? lum(out, mod, off) : out;
}

/**
 * How far series i of n moves from the colour in a monochrome scheme: negative darker, positive
 * lighter (fractions of linear light), as measured in Excel for 2 to 6 series.
 */
const SPREAD: Record<number, number> = { 1: 0, 2: 0.235, 3: 0.35, 4: 0.416, 5: 0.457, 6: 0.5 };

/** Series i's colour in scheme `id` for n series; `theme` maps accent1-6, dk1, lt1 … to colours. */
export function schemeColor(id: number, i: number, n: number, theme: Record<string, string>): string {
  const s = COLOR_SCHEMES[id] ?? COLOR_SCHEMES[10];
  const base = (name: string) => theme[name] ?? "#000000";
  if (s.meth === "cycle") {
    const c = base(s.colors[i % s.colors.length]);
    const v = s.variations.length ? s.variations[Math.floor(i / s.colors.length) % s.variations.length] : [];
    // Schemes whose variations start with "none" leave the first round alone (Farbig); others vary every series.
    const vi = s.variations[0]?.length ? s.variations[i % s.variations.length] : v;
    return apply(c, vi);
  }
  const c = base(s.colors[0]);
  if (n <= 1) return c;
  const m = SPREAD[Math.min(6, n)] ?? 0.5;
  const k = s.meth === "withinLinearReversed" ? n - 1 - i : i;
  const off = -m + (2 * m * k) / (n - 1);
  if (Math.abs(off) < 1e-9) return c;
  return off < 0 ? shade(c, 1 + off) : tint(c, 1 - off);
}
