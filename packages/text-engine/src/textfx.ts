// Texteffekte (Word 2010+ run properties in the w14 namespace): glow, shadow, reflection,
// textOutline, textFill, scene3d, props3d and the typography (ligatures, numForm, numSpacing,
// stylisticSets, cntxtAlts). They are kept as Word's XML on the run (CharProps.textFx), so
// files keep them exactly; these helpers read and replace single elements and turn them into
// CSS for the screen (shadows, glow, reflection, outline, gradient fill, OpenType features).

/** Word's order of the w14 elements at the end of w:rPr. */
export const TEXTFX_ORDER = ["glow", "shadow", "reflection", "textOutline", "textFill", "scene3d", "props3d", "ligatures", "numForm", "numSpacing", "stylisticSets", "cntxtAlts"] as const;
export type TextFxPart = (typeof TEXTFX_ORDER)[number];
export const W14_NS = "http://schemas.microsoft.com/office/word/2010/wordml";

/** The w14 elements of a textFx string, by name. */
export function fxParts(fx: string | undefined): Partial<Record<TextFxPart, string>> {
  const out: Partial<Record<TextFxPart, string>> = {};
  if (!fx) return out;
  const re = /<w14:([A-Za-z0-9]+)\b[^>]*?(?:\/>|>[\s\S]*?<\/w14:\1>)/g;
  for (let m = re.exec(fx); m; m = re.exec(fx)) {
    const name = m[1] as TextFxPart;
    if ((TEXTFX_ORDER as readonly string[]).includes(name)) out[name] = m[0];
  }
  return out;
}

/** A textFx string from parts, in Word's order; undefined when empty. */
export function fxJoin(parts: Partial<Record<TextFxPart, string | undefined>>): string | undefined {
  const s = TEXTFX_ORDER.map((k) => parts[k] ?? "").join("");
  return s || undefined;
}

/** Replaces (or removes, with null) one element of a textFx string. */
export function fxSet(fx: string | undefined, part: TextFxPart, xml: string | null): string | undefined {
  const parts = fxParts(fx);
  if (xml) parts[part] = xml; else delete parts[part];
  return fxJoin(parts);
}

const attr = (xml: string, name: string): string | undefined => new RegExp(`w14:${name}="([^"]*)"`).exec(xml)?.[1];
const num = (xml: string, name: string, fallback = 0) => { const v = attr(xml, name); return v === undefined ? fallback : Number(v); };
const EMU_PER_PX = 9525;

const SCHEME: Record<string, string> = { tx1: "dk1", bg1: "lt1", tx2: "dk2", bg2: "lt2", phClr: "accent1" };

function hslOf(hex: string): [number, number, number] {
  const n = parseInt(hex, 16);
  const r = ((n >> 16) & 255) / 255, g = ((n >> 8) & 255) / 255, b = (n & 255) / 255;
  const max = Math.max(r, g, b), min = Math.min(r, g, b), l = (max + min) / 2, d = max - min;
  if (!d) return [0, 0, l];
  const s = d / (1 - Math.abs(2 * l - 1));
  let h = max === r ? ((g - b) / d) % 6 : max === g ? (b - r) / d + 2 : (r - g) / d + 4;
  h = (h * 60 + 360) % 360;
  return [h, s, l];
}

function hexOf(h: number, s: number, l: number): string {
  l = Math.max(0, Math.min(1, l));
  s = Math.max(0, Math.min(1, s));
  const c = (1 - Math.abs(2 * l - 1)) * s, x = c * (1 - Math.abs(((h / 60) % 2) - 1)), m = l - c / 2;
  const [r, g, b] = h < 60 ? [c, x, 0] : h < 120 ? [x, c, 0] : h < 180 ? [0, c, x] : h < 240 ? [0, x, c] : h < 300 ? [x, 0, c] : [c, 0, x];
  return [r, g, b].map((v) => Math.round((v + m) * 255).toString(16).padStart(2, "0")).join("").toUpperCase();
}

/**
 * The first DrawingML colour in `xml` (w14:srgbClr, w14:schemeClr, w14:prstClr) with its
 * modifiers (lumMod, lumOff, satMod, tint, shade, alpha) as CSS rgba; `fallback` without one.
 */
export function fxColor(xml: string, colors: Readonly<Record<string, string>> | undefined, fallback = "#000"): string {
  const m = /<w14:(srgbClr|schemeClr|prstClr)\s+w14:val="([^"]*)"\s*(\/>|>([\s\S]*?)<\/w14:\1>)/.exec(xml);
  if (!m) return fallback;
  const presets: Record<string, string> = { black: "000000", white: "FFFFFF", red: "FF0000", green: "008000", blue: "0000FF", yellow: "FFFF00" };
  let hex = m[1] === "srgbClr" ? m[2] : m[1] === "schemeClr" ? colors?.[SCHEME[m[2]] ?? m[2]] ?? "000000" : presets[m[2]] ?? "000000";
  const mods = m[4] ?? "";
  const mod = (k: string) => { const v = new RegExp(`<w14:${k}\\s+w14:val="(\\d+)"`).exec(mods)?.[1]; return v === undefined ? undefined : Number(v) / 100000; };
  let [h, s, l] = hslOf(hex);
  const lumMod = mod("lumMod"), lumOff = mod("lumOff"), satMod = mod("satMod"), tint = mod("tint"), shade = mod("shade");
  if (lumMod !== undefined || lumOff !== undefined) l = l * (lumMod ?? 1) + (lumOff ?? 0);
  if (satMod !== undefined) s *= satMod;
  if (tint !== undefined) l = l * tint + (1 - tint);
  if (shade !== undefined) l *= shade;
  hex = hexOf(h, s, l);
  const alpha = mod("alpha") ?? 1;
  const n = parseInt(hex, 16);
  return alpha >= 0.999 ? `#${hex}` : `rgba(${(n >> 16) & 255},${(n >> 8) & 255},${n & 255},${Math.round(alpha * 1000) / 1000})`;
}

const cache = new Map<string, string>();

/**
 * CSS for a run's text effects at `scale` px per px of the layout (zoom is applied by the
 * view's transform, so 1). Shadows and glows become text-shadow, the outline a text stroke,
 * a gradient fill a clipped background, the reflection -webkit-box-reflect.
 */
export function textFxCss(fx: string, colors: Readonly<Record<string, string>> | undefined): string {
  const key = `${fx}|${colors ? Object.values(colors).join(",") : ""}`;
  const hit = cache.get(key);
  if (hit !== undefined) return hit;
  const p = fxParts(fx);
  const css: string[] = [];
  const shadows: string[] = [];
  if (p.glow) {
    const r = num(p.glow, "rad") / EMU_PER_PX;
    const c = fxColor(p.glow, colors, "rgba(0,0,0,.5)");
    if (r > 0) shadows.push(`0 0 ${r * 0.5}px ${c}`, `0 0 ${r}px ${c}`, `0 0 ${r * 1.4}px ${c}`);
  }
  if (p.shadow) {
    const dist = num(p.shadow, "dist") / EMU_PER_PX, dir = (num(p.shadow, "dir") / 60000) * (Math.PI / 180), blur = num(p.shadow, "blurRad") / EMU_PER_PX;
    const c = fxColor(p.shadow, colors, "rgba(0,0,0,.6)");
    shadows.push(`${(Math.cos(dir) * dist).toFixed(2)}px ${(Math.sin(dir) * dist).toFixed(2)}px ${blur.toFixed(2)}px ${c}`);
  }
  if (shadows.length) css.push(`text-shadow:${shadows.join(",")}`);
  if (p.textOutline && !/<w14:noFill\/>/.test(p.textOutline)) {
    const w = Math.max(0.5, num(p.textOutline, "w") / EMU_PER_PX);
    css.push(`-webkit-text-stroke:${w.toFixed(2)}px ${fxColor(p.textOutline, colors)}`);
  }
  if (p.textFill) {
    if (/<w14:noFill\/>/.test(p.textFill)) css.push("color:transparent", "-webkit-text-fill-color:transparent");
    else if (/<w14:gradFill>/.test(p.textFill)) {
      const stops = [...p.textFill.matchAll(/<w14:gs\s+w14:pos="(\d+)"\s*>([\s\S]*?)<\/w14:gs>/g)].map((g) => `${fxColor(g[2], colors)} ${Number(g[1]) / 1000}%`);
      const ang = num(p.textFill, "ang", 5400000) / 60000;
      if (stops.length) css.push(`background-image:linear-gradient(${ang + 90}deg,${stops.join(",")})`, "-webkit-background-clip:text", "background-clip:text", "-webkit-text-fill-color:transparent");
    } else css.push(`-webkit-text-fill-color:${fxColor(p.textFill, colors)}`);
  }
  if (p.reflection) {
    const stA = num(p.reflection, "stA", 50000) / 100000, endA = num(p.reflection, "endA", 300) / 100000, endPos = num(p.reflection, "endPos", 35000) / 1000;
    const dist = num(p.reflection, "dist") / EMU_PER_PX;
    css.push(`-webkit-box-reflect:below ${(dist - 4).toFixed(1)}px linear-gradient(transparent ${100 - endPos}%, rgba(0,0,0,${endA}) ${100 - endPos}%, rgba(0,0,0,${stA}))`);
  }
  // Typography: OpenType features.
  const features: string[] = [];
  const lig = p.ligatures ? attr(p.ligatures, "val") : undefined;
  if (lig) {
    const map: Record<string, string> = { none: "none", standard: "common-ligatures", contextual: "contextual", historical: "historical-ligatures", discretional: "discretionary-ligatures",
      standardContextual: "common-ligatures contextual", standardHistorical: "common-ligatures historical-ligatures", contextualHistorical: "contextual historical-ligatures",
      standardDiscretional: "common-ligatures discretionary-ligatures", contextualDiscretional: "contextual discretionary-ligatures", historicalDiscretional: "historical-ligatures discretionary-ligatures",
      standardContextualHistorical: "common-ligatures contextual historical-ligatures", standardContextualDiscretional: "common-ligatures contextual discretionary-ligatures",
      standardHistoricalDiscretional: "common-ligatures historical-ligatures discretionary-ligatures", contextualHistoricalDiscretional: "contextual historical-ligatures discretionary-ligatures",
      all: "common-ligatures contextual historical-ligatures discretionary-ligatures" };
    css.push(`font-variant-ligatures:${map[lig] ?? "normal"}`);
  }
  const numeric: string[] = [];
  const form = p.numForm ? attr(p.numForm, "val") : undefined;
  if (form === "lining") numeric.push("lining-nums");
  if (form === "oldStyle") numeric.push("oldstyle-nums");
  const spacing = p.numSpacing ? attr(p.numSpacing, "val") : undefined;
  if (spacing === "proportional") numeric.push("proportional-nums");
  if (spacing === "tabular") numeric.push("tabular-nums");
  if (numeric.length) css.push(`font-variant-numeric:${numeric.join(" ")}`);
  if (p.stylisticSets) for (const m of p.stylisticSets.matchAll(/w14:id="(\d+)"/g)) features.push(`"ss${m[1].padStart(2, "0")}" 1`);
  if (p.cntxtAlts && attr(p.cntxtAlts, "val") !== "0" && attr(p.cntxtAlts, "val") !== "false") features.push('"calt" 1');
  if (features.length) css.push(`font-feature-settings:${features.join(",")}`);
  const out = css.join(";");
  if (cache.size > 500) cache.clear();
  cache.set(key, out);
  return out;
}

// ----- elements as Word writes them --------------------------------------------------------------

/** w14:textOutline with a solid colour, width (pt) and dash, as Word's Kontur menu writes it. */
export function outlineXml(colorXml: string | null, widthPt: number, dash = "solid"): string {
  const fill = colorXml ? `<w14:solidFill>${colorXml}</w14:solidFill>` : "<w14:noFill/>";
  return `<w14:textOutline w14:w="${Math.round(widthPt * 12700)}" w14:cap="flat" w14:cmpd="sng" w14:algn="ctr">${fill}<w14:prstDash w14:val="${dash}"/><w14:round/></w14:textOutline>`;
}

/** A DrawingML colour element in the w14 namespace: a theme colour (w:themeColor reference) or RGB. */
export function fxColorXml(hex: string, themeRef?: string, alpha?: number): string {
  const a = alpha !== undefined && alpha < 1 ? `<w14:alpha w14:val="${Math.round((1 - alpha) * 100000)}"/>` : "";
  const scheme = themeRef ? /w:themeColor="([^"]*)"/.exec(themeRef)?.[1] : undefined;
  const map: Record<string, string> = { text1: "tx1", background1: "bg1", text2: "tx2", background2: "bg2" };
  if (scheme) {
    const t = /w:themeTint="([0-9A-Fa-f]+)"/.exec(themeRef!)?.[1], s = /w:themeShade="([0-9A-Fa-f]+)"/.exec(themeRef!)?.[1];
    // Word writes tints and shades of theme colours as lumMod/lumOff.
    const mods = t ? `<w14:lumMod w14:val="${Math.round((parseInt(t, 16) / 255) * 100000)}"/><w14:lumOff w14:val="${Math.round((1 - parseInt(t, 16) / 255) * 100000)}"/>`
      : s ? `<w14:lumMod w14:val="${Math.round((parseInt(s, 16) / 255) * 100000)}"/>` : "";
    return `<w14:schemeClr w14:val="${map[scheme] ?? scheme}">${mods}${a}</w14:schemeClr>`;
  }
  return a ? `<w14:srgbClr w14:val="${hex}">${a}</w14:srgbClr>` : `<w14:srgbClr w14:val="${hex}"/>`;
}

/** The parameters of a shadow, reflection or glow element (for the pane). */
export function fxNumbers(xml: string | undefined): Record<string, number> {
  const out: Record<string, number> = {};
  if (!xml) return out;
  const head = /^<w14:[A-Za-z0-9]+\b([^>]*)>/.exec(xml)?.[1] ?? "";
  for (const m of head.matchAll(/w14:([A-Za-z]+)="(-?\d+)"/g)) out[m[1]] = Number(m[2]);
  return out;
}
