// Our own designs (D-078, D-190): the theme part of each (colours, heading and body font, effect style), shared by the
// three apps. RaabePresent draws each with artwork of its own on the master and layouts (apps/raabe-present/src/
// designs.ts); RaabeCalc and RaabeWrite offer the same colours and fonts in their Designs galleries, so a chart or a
// table in a worksheet or a document can use the school design of the slides (D-196).
//
// Every design has four Varianten as PowerPoint shows them (Entwurf > Varianten): the design itself, two with other
// leading colours and one with the opposite brightness (a dark design gets a light one and the other way round). The
// variants are derived from the design's colours here, then made readable: text colours keep WCAG AA contrast
// against the background in every variant (D-192). Colours are RRGGBB without "#".

export type DesignGroup = "Raabeschule" | "Allgemein" | "Fächer" | "Anlässe";
export const DESIGN_GROUPS: readonly DesignGroup[] = ["Raabeschule", "Allgemein", "Fächer", "Anlässe"];

export const COLOR_KEYS = ["dk1", "lt1", "dk2", "lt2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6", "hlink", "folHlink"] as const;
export type ColorKey = (typeof COLOR_KEYS)[number];
export type Colors = Readonly<Record<ColorKey, string>>;

export interface SchoolDesign {
  name: string;
  group: DesignGroup;
  /** Heading and body font: Windows fonts (with Calibri / sans-serif as the renderer's fallback, D-191). */
  major: string;
  minor: string;
  /** The theme's effect style, by the name of one of PowerPoint's effect schemes (Entwurf > Effekte). */
  effects: string;
  colors: Colors;
  /** Light text on the dark colours (PowerPoint's colour map bg1 = dk1). */
  dark?: boolean;
  /** Colour of titles as a scheme name (accent1, tx2 …); tx1 when missing. */
  title?: string;
  /** Colours of Varianten 2-4 in place of the derived ones (merged over the design's colours). */
  variants?: readonly (Partial<Record<ColorKey, string>> & { dark?: boolean })[];
}

export interface DesignVariant {
  /** "Mathematik" for the first, "Mathematik Variante 2" …: the name of the theme's colour scheme. */
  colorName: string;
  colors: Colors;
  dark: boolean;
}

// ----- colours ------------------------------------------------------------------------------------------------

const rgb = (hex: string): [number, number, number] => [0, 2, 4].map((i) => parseInt(hex.slice(i, i + 2), 16)) as [number, number, number];
const toHex = (c: readonly number[]) => c.map((v) => Math.max(0, Math.min(255, Math.round(v))).toString(16).padStart(2, "0")).join("").toUpperCase();

/** WCAG relative luminance of RRGGBB. */
export function luminance(hex: string): number {
  const [r, g, b] = rgb(hex).map((v) => { const s = v / 255; return s <= 0.03928 ? s / 12.92 : ((s + 0.055) / 1.055) ** 2.4; });
  return 0.2126 * r + 0.7152 * g + 0.0722 * b;
}

/** WCAG contrast ratio of two RRGGBB colours (1 … 21). */
export function contrast(a: string, b: string): number {
  const [x, y] = [luminance(a), luminance(b)].sort((p, q) => q - p);
  return (x + 0.05) / (y + 0.05);
}

/** `hex` mixed with `to` by share t (0 = hex, 1 = to). */
export function mix(hex: string, to: string, t: number): string {
  const a = rgb(hex), b = rgb(to);
  return toHex(a.map((v, i) => v + (b[i] - v) * t));
}

/** `hex` made darker (towards black) or lighter (towards white) in small steps until it reaches `ratio` against `bg`. */
export function readable(hex: string, bg: string, ratio: number): string {
  if (contrast(hex, bg) >= ratio) return hex;
  const to = luminance(bg) > 0.18 ? "000000" : "FFFFFF";
  for (let t = 0.04; t <= 1.001; t += 0.04) {
    const c = mix(hex, to, t);
    if (contrast(c, bg) >= ratio) return c;
  }
  return to;
}

// ----- Varianten ----------------------------------------------------------------------------------------------

/** Contrast the apps keep: body text 7 (AAA, for projectors), titles and links 4.5 (AA), accents 3 (shapes, charts). */
export const MIN_CONTRAST = { body: 7, text: 4.5, accent: 3 } as const;

const ACCENTS = ["accent1", "accent2", "accent3", "accent4", "accent5", "accent6"] as const;

/** The background, text and second text colour of a colour set as shown (dark: bg1 = dk1). */
export function surfaces(c: Colors, dark: boolean): { bg: string; text: string; bg2: string; text2: string } {
  return dark ? { bg: c.dk1, text: c.lt1, bg2: c.dk2, text2: c.lt2 } : { bg: c.lt1, text: c.dk1, bg2: c.lt2, text2: c.dk2 };
}

/** The scheme colour a title uses in a colour set (tx1/tx2/bg… mapped as the colour map maps them). */
export function schemeHex(c: Colors, dark: boolean, name: string): string {
  const s = surfaces(c, dark);
  const map: Record<string, string> = { tx1: s.text, tx2: s.text2, bg1: s.bg, bg2: s.bg2 };
  return map[name] ?? (c as Record<string, string>)[name] ?? s.text;
}

/**
 * Colours made readable on their background: titles and links, and the accents too when the variant changed the
 * brightness (a design's own light accents, Pastell's, stay as they are on their own background).
 */
function fixContrast(c: Record<ColorKey, string>, dark: boolean, title: string | undefined, accents: boolean): Record<ColorKey, string> {
  const bg = surfaces(c, dark).bg;
  if (accents) for (const k of ACCENTS) c[k] = readable(c[k], bg, MIN_CONTRAST.accent);
  if (title && (ACCENTS as readonly string[]).includes(title)) c[title as ColorKey] = readable(c[title as ColorKey], bg, MIN_CONTRAST.text);
  c.hlink = readable(c.hlink, bg, MIN_CONTRAST.text);
  c.folHlink = readable(c.folHlink, bg, MIN_CONTRAST.text);
  return c;
}

/** The design's colours with the accents turned by n places (accent2 leads for n = 1). */
function rotated(c: Colors, n: number): Record<ColorKey, string> {
  const out = { ...c } as Record<ColorKey, string>;
  ACCENTS.forEach((k, i) => { out[k] = c[ACCENTS[(i + n) % 6]]; });
  out.hlink = out.accent1;
  return out;
}

/** The design's colours with the opposite brightness. */
function flipped(c: Colors, dark: boolean): Record<ColorKey, string> {
  const out = { ...c } as Record<ColorKey, string>;
  if (!dark) {
    // Light to dark: a deep shade of the design's second dark colour as the background.
    out.dk1 = mix(c.dk2, "000000", 0.62);
    out.dk2 = readable(mix(c.dk2, "000000", 0.35), "FFFFFF", MIN_CONTRAST.body);
    out.lt1 = "FFFFFF";
    out.lt2 = readable(c.lt2, out.dk2, MIN_CONTRAST.text);
  } else {
    // Dark to light: white with the design's dark colours as text.
    out.lt1 = "FFFFFF";
    out.lt2 = luminance(c.lt2) > 0.75 ? c.lt2 : mix(c.lt2, "FFFFFF", 0.7);
    out.dk1 = readable(c.dk1, "FFFFFF", MIN_CONTRAST.body);
    out.dk2 = readable(c.dk2, "FFFFFF", MIN_CONTRAST.body);
  }
  return out;
}

/** The four Varianten of a design (PowerPoint shows four per design). */
export function designVariants(d: SchoolDesign): DesignVariant[] {
  const dark = !!d.dark;
  const own = d.variants ?? [];
  const derived: { colors: Record<ColorKey, string>; dark: boolean }[] = [
    { colors: { ...d.colors }, dark },
    { colors: rotated(d.colors, 1), dark },
    { colors: rotated(d.colors, 3), dark },
    { colors: flipped(d.colors, dark), dark: !dark },
  ];
  return derived.map((v, i) => {
    // An own variant is merged over the design's colours (over the flipped ones when it changes the brightness).
    const o = i > 0 && own[i - 1] && Object.keys(own[i - 1]).length ? own[i - 1] : undefined;
    const vdark = o?.dark ?? v.dark;
    const colors = o ? ({ ...(vdark !== dark ? derived[3].colors : d.colors), ...Object.fromEntries(Object.entries(o).filter(([k]) => k !== "dark")) } as Record<ColorKey, string>) : v.colors;
    return { colorName: i === 0 ? d.name : `${d.name} Variante ${i + 1}`, colors: i === 0 ? colors : fixContrast(colors, vdark, d.title, vdark !== dark), dark: vdark };
  });
}

// ----- the designs --------------------------------------------------------------------------------------------

/** A colour set from its dark 2, light 2 and six accents (dark 1 black, light 1 white, links from the accents). */
function pal(dk2: string, lt2: string, a: readonly string[], more: Partial<Record<ColorKey, string>> = {}): Colors {
  return { dk1: "000000", lt1: "FFFFFF", dk2, lt2, accent1: a[0], accent2: a[1], accent3: a[2], accent4: a[3], accent5: a[4], accent6: a[5], hlink: a[0], folHlink: a[4], ...more };
}

const RAABE = pal("3B3B3B", "F2F2F2", ["B3141B", "595959", "D98C1E", "2E75B6", "7F7F7F", "5E9E3A"], { hlink: "B3141B", folHlink: "7A0E13" });

export const SCHOOL_DESIGNS: readonly SchoolDesign[] = [
  // ----- Raabeschule: the school's red (D-018) ------------------------------------------------------------------
  { name: "Raabeschule", group: "Raabeschule", major: "Calibri Light", minor: "Calibri", effects: "Office", colors: RAABE, title: "accent1",
    variants: [{ accent2: "2E75B6", accent4: "595959" }, { dk2: "2F3B45", accent2: "2F3B45", accent3: "8C9BA5", accent4: "D98C1E" }] },
  { name: "Raabeschule Band", group: "Raabeschule", major: "Segoe UI Semibold", minor: "Segoe UI", effects: "Office",
    colors: pal("7A0E13", "F6F1F1", ["B3141B", "404040", "D98C1E", "2E75B6", "8C8C8C", "5E9E3A"], { hlink: "B3141B", folHlink: "7A0E13" }), title: "tx2",
    variants: [{ dk2: "2F3B45", accent2: "2F3B45" }, { dk2: "1F3B5C", accent2: "2E75B6", accent4: "404040" }] },

  // ----- Allgemein ------------------------------------------------------------------------------------------------
  { name: "Klar", group: "Allgemein", major: "Segoe UI Semibold", minor: "Segoe UI", effects: "Office",
    colors: pal("1F3864", "EEF2F7", ["2F5597", "00A3A3", "F4B183", "8FAADC", "C55A11", "548235"], { hlink: "2F5597", folHlink: "7030A0" }) },
  { name: "Tafel", group: "Allgemein", major: "Calibri Light", minor: "Calibri", effects: "Office", dark: true, title: "accent1",
    colors: { dk1: "1E2B24", lt1: "FFFFFF", dk2: "2C3D33", lt2: "E7E6E6", accent1: "F2C14E", accent2: "6EC6CA", accent3: "F28482", accent4: "A7C957", accent5: "C3B1E1", accent6: "F7F7F7", hlink: "6EC6CA", folHlink: "C3B1E1" } },
  { name: "Wellenband", group: "Allgemein", major: "Segoe UI Light", minor: "Segoe UI", effects: "Office",
    colors: pal("16324F", "EAF2F8", ["1F77B4", "2CA6A4", "F2A541", "5DA9E9", "E4572E", "76B041"], { hlink: "1F77B4", folHlink: "6A4C93" }) },
  { name: "Kreisbögen", group: "Allgemein", major: "Calibri Light", minor: "Calibri", effects: "Office",
    colors: pal("2B2D42", "F1F1F4", ["EF476F", "FFD166", "06D6A0", "118AB2", "073B4C", "8D99AE"], { hlink: "107FA4", folHlink: "C0305A" }) },
  { name: "Keil", group: "Allgemein", major: "Segoe UI Semibold", minor: "Segoe UI", effects: "Office",
    colors: pal("263238", "ECEFF1", ["00897B", "26A69A", "FFB300", "5C6BC0", "EF5350", "8D6E63"], { hlink: "008476", folHlink: "5C6BC0" }) },
  { name: "Punkteraster", group: "Allgemein", major: "Calibri Light", minor: "Calibri", effects: "Office",
    colors: pal("3A3A3A", "F4F1EC", ["B5651D", "6B8E23", "4682B4", "CD853F", "8B4513", "708090"], { hlink: "4078A6", folHlink: "8B4513" }) },
  { name: "Schräge Streifen", group: "Allgemein", major: "Segoe UI Light", minor: "Segoe UI", effects: "Office",
    colors: pal("2E1A47", "F3EFF8", ["7B2CBF", "C77DFF", "FF9E00", "3A86FF", "FB5607", "8AC926"], { hlink: "3171D6", folHlink: "7B2CBF" }) },
  { name: "Abendblau", group: "Allgemein", major: "Calibri Light", minor: "Calibri", effects: "Milchglas", dark: true, title: "accent1",
    colors: { dk1: "0B1D3A", lt1: "FFFFFF", dk2: "142B52", lt2: "E3ECF7", accent1: "4FC3F7", accent2: "FFD54F", accent3: "81C784", accent4: "F06292", accent5: "BA68C8", accent6: "FF8A65", hlink: "4FC3F7", folHlink: "BA68C8" } },
  { name: "Stille", group: "Allgemein", major: "Segoe UI Light", minor: "Segoe UI", effects: "Subtile Körper", title: "tx2",
    colors: pal("2F3E3A", "EEF2EF", ["5B7F6E", "A3B9A8", "C9A66B", "6C8EA4", "8E7A6B", "4F6D7A"], { dk1: "1F2624", hlink: "3E6B57", folHlink: "6B5A4C" }) },
  { name: "Papier", group: "Allgemein", major: "Georgia", minor: "Georgia", effects: "Office", title: "accent1",
    colors: pal("4A3F35", "EFE6D6", ["9C4A2F", "3E6E73", "C08A2D", "6B5B95", "7A8B4F", "8C6D5A"], { dk1: "2B2622", lt1: "FBF7EF", folHlink: "6B5B95" }) },
  { name: "Kontrast", group: "Allgemein", major: "Bahnschrift", minor: "Segoe UI", effects: "Office", dark: true, title: "accent1",
    colors: { dk1: "111111", lt1: "FFFFFF", dk2: "2A2A2A", lt2: "F0F0F0", accent1: "FFD400", accent2: "FF5A36", accent3: "00C2FF", accent4: "7CE36B", accent5: "FF7BD5", accent6: "BDBDBD", hlink: "00C2FF", folHlink: "FF7BD5" },
    variants: [{}, {}, { dark: false, accent1: "1A1A1A", accent2: "E03A12", accent3: "0077B6", accent4: "2E8B2E", accent5: "C2187E", accent6: "707070", hlink: "0077B6", folHlink: "C2187E" }] },
  { name: "Minimal", group: "Allgemein", major: "Segoe UI Light", minor: "Segoe UI", effects: "Subtile Körper",
    colors: pal("222222", "F5F5F5", ["2B6CB0", "E53E3E", "38A169", "D69E2E", "805AD5", "718096"], { dk1: "1A1A1A" }) },
  { name: "Mitternacht", group: "Allgemein", major: "Segoe UI Light", minor: "Segoe UI", effects: "Milchglas", dark: true, title: "accent1",
    colors: { dk1: "101820", lt1: "F5F7FA", dk2: "1C2A38", lt2: "C9D6E3", accent1: "7FD1FF", accent2: "FFB86B", accent3: "A0E7A0", accent4: "FF8FA3", accent5: "C5A3FF", accent6: "6EE7D8", hlink: "7FD1FF", folHlink: "C5A3FF" } },
  { name: "Pastell", group: "Allgemein", major: "Trebuchet MS", minor: "Trebuchet MS", effects: "Subtile Körper", title: "tx2",
    colors: pal("3D3553", "FFF6F0", ["F4A6B7", "A7D8DE", "FFD59E", "B9E4A8", "C8B6E2", "F7C8A0"], { dk1: "2A2438", hlink: "6A4C93", folHlink: "8E5572" }) },
  { name: "Konfetti", group: "Allgemein", major: "Trebuchet MS", minor: "Segoe UI", effects: "Office", title: "tx2",
    colors: pal("25304A", "F7F7FB", ["E63946", "F4A261", "2A9D8F", "457B9D", "E9C46A", "9B5DE5"], { hlink: "1D5C8C", folHlink: "7B3FC4" }) },
  { name: "Mosaik", group: "Allgemein", major: "Bahnschrift", minor: "Segoe UI", effects: "Office", title: "tx2",
    colors: pal("1E3A4C", "EEF4F7", ["1B998B", "2D3047", "FF9B71", "E84855", "F2C14E", "3A6EA5"], { hlink: "1F6F9F", folHlink: "8E3B5F" }) },
  { name: "Rahmen", group: "Allgemein", major: "Constantia", minor: "Constantia", effects: "Office", title: "tx2",
    colors: pal("1F2A44", "F3F1EA", ["1F4E79", "A67C2E", "6B2737", "3F6B4F", "7F7F7F", "4A6FA5"], { dk1: "1A1A1A" }) },
  { name: "Grundformen", group: "Allgemein", major: "Bahnschrift", minor: "Segoe UI", effects: "Office",
    colors: pal("1D1D1B", "F4F1EA", ["D7263D", "1B4F9C", "F2C12E", "2E8B57", "F46036", "5C5C5C"], { hlink: "1B4F9C", folHlink: "7A3E9D" }) },
  { name: "Horizont", group: "Allgemein", major: "Segoe UI Light", minor: "Segoe UI", effects: "Subtile Körper", title: "tx2",
    colors: pal("263A52", "E8F1F8", ["2E86AB", "F18F01", "C73E1D", "3B1F2B", "44BBA4", "7A9CC6"], { hlink: "1F6F94", folHlink: "6A4C93" }) },
  { name: "Eckpunkt", group: "Allgemein", major: "Segoe UI Semibold", minor: "Segoe UI", effects: "Office",
    colors: pal("2B2D42", "EDF2F4", ["D90429", "8D99AE", "2B2D42", "EF8354", "4F5D75", "BFC0C0"], { hlink: "B00020", folHlink: "4F5D75" }) },

  // ----- Fächer -------------------------------------------------------------------------------------------------
  { name: "Mathematik", group: "Fächer", major: "Cambria", minor: "Calibri", effects: "Office", title: "accent1",
    colors: pal("1B2F5B", "EEF3FA", ["1F5AA6", "D1495B", "2A9D8F", "F2A541", "6C4AB6", "5C6B73"]) },
  { name: "Deutsch", group: "Fächer", major: "Constantia", minor: "Calibri", effects: "Office", title: "accent1",
    colors: pal("3A1F1D", "F7F2EE", ["8E1B2C", "2F4858", "C77D2E", "33658A", "86BBD8", "6B8F71"]) },
  { name: "Englisch", group: "Fächer", major: "Segoe UI Semibold", minor: "Segoe UI", effects: "Office", title: "accent1",
    colors: pal("13294B", "EEF2F8", ["1D3F72", "C8102E", "5B8DB8", "E9A23B", "2E8540", "8A8D91"]) },
  { name: "Französisch", group: "Fächer", major: "Candara", minor: "Candara", effects: "Office", title: "accent1",
    colors: pal("1E2A4A", "F3F4F9", ["2B4C9B", "D7263D", "7FA7D9", "E7B85B", "4E8D7C", "9B6A9E"]) },
  { name: "Latein", group: "Fächer", major: "Palatino Linotype", minor: "Palatino Linotype", effects: "Office", title: "accent1",
    colors: pal("3E2C23", "F4EEE4", ["9A3B26", "B08D57", "4F5D2F", "7A6A58", "2F4B6E", "A35D3F"], { dk1: "231A15" }) },
  { name: "Geschichte", group: "Fächer", major: "Georgia", minor: "Georgia", effects: "Office", title: "accent1",
    colors: pal("3B2F2A", "F3ECE2", ["8B5A2B", "4A6670", "A63D40", "C9A227", "6E7F5E", "5C4B6B"], { dk1: "241C18", lt1: "FFFDF8" }) },
  { name: "Erdkunde", group: "Fächer", major: "Segoe UI Semibold", minor: "Segoe UI", effects: "Office", title: "accent1",
    colors: pal("23403A", "EEF4EC", ["2E7D32", "1E88A8", "A1887F", "F2B134", "6A994E", "3D5A80"]) },
  { name: "Politik-Wirtschaft", group: "Fächer", major: "Segoe UI Semibold", minor: "Segoe UI", effects: "Office", title: "accent1",
    colors: pal("1C2B3A", "EEF1F5", ["1F4E79", "E07A1F", "2A9D8F", "B23A48", "6C757D", "8AB17D"]) },
  { name: "Biologie", group: "Fächer", major: "Candara", minor: "Candara", effects: "Subtile Körper", title: "accent1",
    colors: pal("1F3B2C", "EEF6EE", ["2F7038", "8DB255", "2F6690", "E9C46A", "B5651D", "7B4F9E"]) },
  { name: "Chemie", group: "Fächer", major: "Segoe UI Semibold", minor: "Segoe UI", effects: "Office", title: "accent1",
    colors: pal("23233F", "F1F0F8", ["5A3E9B", "1A9E8F", "E76F51", "F4A261", "2A6F97", "9C6644"]) },
  { name: "Physik", group: "Fächer", major: "Segoe UI Light", minor: "Segoe UI", effects: "Office", title: "accent1",
    colors: pal("0F2742", "EDF3F9", ["0B5394", "F18701", "2BA84A", "D62246", "5E60CE", "7D8597"]) },
  { name: "Informatik", group: "Fächer", major: "Consolas", minor: "Segoe UI", effects: "Office", dark: true, title: "accent1",
    colors: { dk1: "0D1B1E", lt1: "F2F7F5", dk2: "15282C", lt2: "CFE3DD", accent1: "3DDC97", accent2: "4CC9F0", accent3: "F9C74F", accent4: "F94144", accent5: "B388EB", accent6: "90BE6D", hlink: "4CC9F0", folHlink: "B388EB" } },
  { name: "Musik", group: "Fächer", major: "Constantia", minor: "Calibri", effects: "Office", title: "accent1",
    colors: pal("2A1E3B", "F5F1F8", ["6A2C70", "B83B5E", "F08A5D", "F9C74F", "3F72AF", "1B998B"]) },
  { name: "Kunst", group: "Fächer", major: "Segoe Print", minor: "Segoe UI", effects: "Office",
    colors: pal("2B2B2B", "FAF6F0", ["E63946", "1D70A2", "F4B400", "2A9D8F", "8E44AD", "F77F00"], { dk1: "1E1E1E", hlink: "1D70A2", folHlink: "8E44AD" }) },
  { name: "Sport", group: "Fächer", major: "Bahnschrift", minor: "Segoe UI", effects: "Office", title: "tx2",
    colors: pal("102A43", "EEF2F6", ["E4572E", "17BEBB", "FFC914", "2E282A", "76B041", "3D5A80"], { hlink: "1F5F99", folHlink: "6A4C93" }) },
  { name: "Religion", group: "Fächer", major: "Constantia", minor: "Corbel", effects: "Subtile Körper", title: "accent1",
    colors: pal("2E2545", "F5F2EA", ["5B3E8E", "C89B3C", "3C6E71", "A44A3F", "7E8BA3", "9C7A97"]) },
  { name: "Werte und Normen", group: "Fächer", major: "Corbel", minor: "Corbel", effects: "Subtile Körper", title: "accent1",
    colors: pal("23394D", "EFF3F4", ["2F6F8F", "D98E04", "7A9E7E", "9C4F6B", "4B4E6D", "8C9A5B"]) },
  { name: "Darstellendes Spiel", group: "Fächer", major: "Constantia", minor: "Segoe UI", effects: "Office", dark: true, title: "accent1",
    colors: { dk1: "1A0F14", lt1: "FFF8F0", dk2: "2B1A21", lt2: "E9D8C8", accent1: "E0A458", accent2: "C0392B", accent3: "F5E6CA", accent4: "8E6C8A", accent5: "4F9DA6", accent6: "D35400", hlink: "E0A458", folHlink: "C9A0C6" } },
  { name: "Seminarfach", group: "Fächer", major: "Cambria", minor: "Calibri", effects: "Office", title: "accent1",
    colors: pal("22313F", "F1F4F6", ["2C5F7C", "B5522B", "5A7D4E", "7C6A9E", "C9A13B", "6B7B8C"]) },

  // ----- Anlässe ------------------------------------------------------------------------------------------------
  { name: "Projektwoche", group: "Anlässe", major: "Segoe UI Semibold", minor: "Segoe UI", effects: "Office", title: "tx2",
    colors: pal("263238", "F4F6F7", ["F25F5C", "FFE066", "247BA0", "70C1B3", "50514F", "9C89B8"], { hlink: "1F6F94", folHlink: "6A4C93" }) },
  { name: "Klassenfahrt", group: "Anlässe", major: "Segoe UI Semibold", minor: "Segoe UI", effects: "Office", title: "tx2",
    colors: pal("1D3557", "F1FAEE", ["457B9D", "E63946", "2A9D8F", "F4A261", "A8DADC", "6D597A"], { hlink: "2F5F80", folHlink: "6D597A" }) },
  { name: "Elternabend", group: "Anlässe", major: "Segoe UI Light", minor: "Segoe UI", effects: "Subtile Körper", title: "tx2",
    colors: pal("2F3E46", "F2F4F3", ["3E6680", "B56576", "7F9C6B", "E0A96D", "6D6875", "355070"]) },
  { name: "Abitur", group: "Anlässe", major: "Palatino Linotype", minor: "Corbel", effects: "Office", dark: true, title: "accent1",
    colors: { dk1: "0F1E3D", lt1: "FFFFFF", dk2: "1B2E57", lt2: "E6E9F0", accent1: "D4AF37", accent2: "8FB3E8", accent3: "F2E3B3", accent4: "C97B63", accent5: "7FB685", accent6: "B9A3E3", hlink: "8FB3E8", folHlink: "B9A3E3" } },
  { name: "Tag der offenen Tür", group: "Anlässe", major: "Segoe UI Semibold", minor: "Segoe UI", effects: "Office", title: "tx2",
    colors: pal("2D2A4A", "FFF8EE", ["E76F51", "2A9D8F", "E9C46A", "264653", "F4A261", "8AB17D"], { hlink: "1F6F66", folHlink: "6A4C93" }) },
];
