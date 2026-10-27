// Colours in SpreadsheetML: rgb (ARGB), theme index with tint, the indexed legacy palette
// and "auto". The model keeps the resolved "#RRGGBB"; the file's own colour element is kept
// for formats that are written back unchanged.

import { attrs, children, type El } from "./xml";

/** Excel's default palette (indexed 0-63); 64 is the system foreground, 65 the background. */
export const INDEXED: readonly string[] = [
  "000000", "FFFFFF", "FF0000", "00FF00", "0000FF", "FFFF00", "FF00FF", "00FFFF",
  "000000", "FFFFFF", "FF0000", "00FF00", "0000FF", "FFFF00", "FF00FF", "00FFFF",
  "800000", "008000", "000080", "808000", "800080", "008080", "C0C0C0", "808080",
  "9999FF", "993366", "FFFFCC", "CCFFFF", "660066", "FF8080", "0066CC", "CCCCFF",
  "000080", "FF00FF", "FFFF00", "00FFFF", "800080", "800000", "008080", "0000FF",
  "00CCFF", "CCFFFF", "CCFFCC", "FFFF99", "99CCFF", "FF99CC", "CC99FF", "FFCC99",
  "3366FF", "33CCCC", "99CC00", "FFCC00", "FF9900", "FF6600", "666699", "969696",
  "003366", "339966", "003300", "333300", "993300", "993366", "333399", "333333",
];

/** Office 2013-2022 theme colours (Calibri design): RaabeCalc's new workbooks (D-011). */
export const OFFICE_2013_THEME: Readonly<Record<string, string>> = {
  dk1: "000000", lt1: "FFFFFF", dk2: "44546A", lt2: "E7E6E6", accent1: "4472C4", accent2: "ED7D31", accent3: "A5A5A5",
  accent4: "FFC000", accent5: "5B9BD5", accent6: "70AD47", hlink: "0563C1", folHlink: "954F72",
};

/** SpreadsheetML's theme index: 0 and 1 (and 2, 3) are swapped against the scheme's order. */
export const THEME_INDEX = ["lt1", "dk1", "lt2", "dk2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6", "hlink", "folHlink"];

export interface Theme {
  /** dk1 … folHlink as "RRGGBB". */
  colors: Record<string, string>;
  major: string;
  minor: string;
}

/** Colours and fonts of a theme part (xl/theme/theme1.xml). */
export function readTheme(xml: string | null): Theme {
  const out: Theme = { colors: { ...OFFICE_2013_THEME }, major: "Calibri Light", minor: "Calibri" };
  if (!xml) return out;
  const scheme = /<a:clrScheme\b[\s\S]*?<\/a:clrScheme>/.exec(xml)?.[0];
  if (scheme) {
    for (const m of scheme.matchAll(/<a:(dk1|lt1|dk2|lt2|accent[1-6]|hlink|folHlink)>\s*<a:(sysClr|srgbClr)\b([^>]*)>/g)) {
      const hex = (m[2] === "sysClr" ? /lastClr="([0-9A-Fa-f]{6})"/ : /val="([0-9A-Fa-f]{6})"/).exec(m[3])?.[1];
      if (hex) out.colors[m[1]] = hex.toUpperCase();
    }
  }
  const font = (kind: string) => new RegExp(`<a:${kind}>\\s*<a:latin typeface="([^"]*)"`).exec(xml)?.[1];
  out.major = font("majorFont") ?? out.major;
  out.minor = font("minorFont") ?? out.minor;
  return out;
}

// Windows' integer HLS (ColorRGBToHLS / ColorHLSToRGB, HLSMAX 240), which Excel uses for tints.
const HLS = 240, RGB = 255;
const idiv = (a: number, b: number) => Math.trunc(a / b);

function rgbToHls(r: number, g: number, b: number): [number, number, number] {
  const max = Math.max(r, g, b), min = Math.min(r, g, b);
  const l = idiv((max + min) * HLS + RGB, 2 * RGB);
  if (max === min) return [idiv(HLS * 2, 3), l, 0];
  const s = l <= HLS / 2 ? idiv((max - min) * HLS + idiv(max + min, 2), max + min)
    : idiv((max - min) * HLS + idiv(2 * RGB - max - min, 2), 2 * RGB - max - min);
  const d = (c: number) => idiv((max - c) * (HLS / 6) + idiv(max - min, 2), max - min);
  let h = r === max ? d(b) - d(g) : g === max ? HLS / 3 + d(r) - d(b) : (2 * HLS) / 3 + d(g) - d(r);
  if (h < 0) h += HLS;
  if (h > HLS) h -= HLS;
  return [h, l, s];
}

function hueToRgb(n1: number, n2: number, h: number): number {
  if (h < 0) h += HLS;
  if (h > HLS) h -= HLS;
  if (h < HLS / 6) return n1 + idiv((n2 - n1) * h + HLS / 12, HLS / 6);
  if (h < HLS / 2) return n2;
  if (h < (HLS * 2) / 3) return n1 + idiv((n2 - n1) * ((HLS * 2) / 3 - h) + HLS / 12, HLS / 6);
  return n1;
}

function hlsToRgb(h: number, l: number, s: number): [number, number, number] {
  if (s === 0) {
    const v = idiv(l * RGB + HLS / 2, HLS);
    return [v, v, v];
  }
  const m2 = l <= HLS / 2 ? idiv(l * (HLS + s) + HLS / 2, HLS) : l + s - idiv(l * s + HLS / 2, HLS);
  const m1 = 2 * l - m2;
  const c = (x: number) => idiv(hueToRgb(m1, m2, x) * RGB + HLS / 2, HLS);
  return [c(h + HLS / 3), c(h), c(h - HLS / 3)];
}

/**
 * A colour lightened (tint > 0) or darkened (tint < 0) as SpreadsheetML defines it: the
 * luminance moves towards white (L·(1−t) + t) or black (L·(1+t)), in Windows' HLS as Excel
 * computes it (checked against Excel's colours, xlsx_excel_check.py).
 */
export function tinted(hex: string, tint: number): string {
  if (!tint) return hex.toUpperCase();
  const [r, g, b] = [0, 2, 4].map((i) => parseInt(hex.slice(i, i + 2), 16));
  const [h, l0, s] = rgbToHls(r, g, b);
  const l = Math.round(tint < 0 ? l0 * (1 + tint) : l0 * (1 - tint) + HLS * tint);
  return hlsToRgb(h, Math.max(0, Math.min(HLS, l)), s).map((x) => Math.max(0, Math.min(255, x)).toString(16).padStart(2, "0")).join("").toUpperCase();
}

export interface ColorContext {
  theme: Record<string, string>;
  /** The workbook's own palette (styles.xml <colors><indexedColors>), else Excel's. */
  indexed: readonly string[];
}

/**
 * "#RRGGBB" of a colour element's attributes (rgb, theme + tint, indexed); undefined for
 * automatic colours (auto="1", indexed 64, no attributes).
 */
export function resolveColor(a: Record<string, string>, ctx: ColorContext): string | undefined {
  let hex: string | undefined;
  if (a.rgb) hex = a.rgb.length === 8 ? a.rgb.slice(2) : a.rgb;
  else if (a.theme !== undefined) hex = ctx.theme[THEME_INDEX[Number(a.theme)]] ?? "000000";
  else if (a.indexed !== undefined) {
    const i = Number(a.indexed);
    if (i === 64 || i === 81) return undefined; // system foreground, comment text
    hex = i === 65 ? "FFFFFF" : ctx.indexed[i] ?? INDEXED[i];
  }
  if (!hex || !/^[0-9A-Fa-f]{6}$/.test(hex)) return undefined;
  const t = a.tint ? Number(a.tint) : 0;
  return "#" + tinted(hex.toUpperCase(), t);
}

/** The colour element (color, fgColor, tabColor …) for a model colour. */
export function colorXml(tag: string, hex: string | undefined): string {
  return hex ? `<${tag} rgb="FF${hex.replace("#", "").toUpperCase()}"/>` : `<${tag} auto="1"/>`;
}

/** The workbook palette from styles.xml (<colors><indexedColors>). */
export function readIndexed(xml: string, colorsEl: El | undefined): readonly string[] {
  if (!colorsEl) return INDEXED;
  const ic = children(xml, colorsEl.open, colorsEl.close).find((c) => c.name === "indexedColors");
  if (!ic) return INDEXED;
  const list = children(xml, ic.open, ic.close).map((c) => (attrs(xml, c).rgb ?? "FF000000").slice(-6).toUpperCase());
  return list.length ? list : INDEXED;
}
