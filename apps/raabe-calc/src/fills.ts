// Pattern and gradient fills of cells (Zellen formatieren > Ausfüllen: Musterformat, Musterfarbe, Fülleffekte,
// D-147): their file form as Excel writes it (recorded, results/2026-11-09-excel-followup.json), and how the grid
// and the dialog's sample draw them. A solid fill stays CellStyle.fill; the others are CellStyle.fillXml.

/** Excel's Musterformat entries in the list's order, with the file's pattern names. */
export const PATTERNS: [string, string][] = [
  // Names as Excel's list has them (recorded, probe_excel_patterns_ui.py, results/2026-11-09-excel-patterns-ui.json).
  ["solid", "Uni"], ["darkGray", "75% Grau"], ["mediumGray", "50% Grau"], ["lightGray", "25% Grau"], ["gray125", "12,5 % Grau"], ["gray0625", "6,25 % Grau"],
  ["darkHorizontal", "Horizontal Gestreift"], ["darkVertical", "Vertikal Gestreift"], ["darkDown", "Umgekehrt diagonal Gestreift"], ["darkUp", "Diagonal Gestreift"],
  ["darkGrid", "Diagonal Kariert"], ["darkTrellis", "Dick Diagonal Kariert"], ["lightHorizontal", "Dünn Horizontal Gestreift"], ["lightVertical", "Dünn Vertikal Gestreift"],
  ["lightDown", "Dünn Umgekehrt diagonal Gestreift"], ["lightUp", "Dünn Diagonal Gestreift"], ["lightGrid", "Dünn Horizontal Kariert"], ["lightTrellis", "Dünn Diagonal Kariert"],
];

/** Which pixel of the 8 x 8 tile is in the pattern's colour. */
function on(type: string, x: number, y: number): boolean {
  const m = (a: number, n: number) => ((a % n) + n) % n;
  switch (type) {
    case "solid": return true;
    case "darkGray": return !(m(x, 2) === 0 && m(y, 2) === 0);
    case "mediumGray": return m(x + y, 2) === 0;
    case "lightGray": return m(x, 2) === 0 && m(y, 2) === 0;
    case "gray125": return m(x, 4) === 0 && m(y, 2) === 0 && m(x / 4 + y / 2, 2) === 0;
    case "gray0625": return m(x, 4) === 0 && m(y, 4) === 0;
    case "darkHorizontal": return m(y, 4) < 2;
    case "darkVertical": return m(x, 4) < 2;
    case "darkDown": return m(x - y, 4) < 2;
    case "darkUp": return m(x + y, 4) < 2;
    case "darkGrid": return m((x >> 1) + (y >> 1), 2) === 0;
    case "darkTrellis": return m(x + y, 2) === 0 || m(y, 2) === 0;
    case "lightHorizontal": return m(y, 4) === 0;
    case "lightVertical": return m(x, 4) === 0;
    case "lightDown": return m(x - y, 4) === 0;
    case "lightUp": return m(x + y, 4) === 0;
    case "lightGrid": return m(x, 4) === 0 || m(y, 4) === 0;
    case "lightTrellis": return m(x + y, 4) === 0 || m(x - y, 4) === 0;
    default: return false;
  }
}

export interface PatternFill { kind: "pattern"; type: string; fg: string; bg?: string }
export interface GradientFill { kind: "gradient"; degree?: number; path?: { left: number; right: number; top: number; bottom: number }; stops: { pos: number; color: string }[] }
export type FancyFill = PatternFill | GradientFill;

const hex = (rgb: string | undefined) => (rgb ? "#" + rgb.slice(-6).toUpperCase() : undefined);
const argb = (c: string) => "FF" + c.replace("#", "").toUpperCase();

/** A fill's XML (CellStyle.fillXml) read back into its parts; undefined for what this does not describe. */
export function parseFill(xml: string | undefined): FancyFill | undefined {
  if (!xml) return undefined;
  if (/<gradientFill/.test(xml)) {
    const head = /<gradientFill\b([^>]*)>/.exec(xml)?.[1] ?? "";
    const num = (k: string) => { const m = new RegExp(`\\b${k}="([^"]+)"`).exec(head); return m ? Number(m[1]) : undefined; };
    const stops = [...xml.matchAll(/<stop position="([^"]+)"><color rgb="([0-9A-Fa-f]{8})"\/><\/stop>/g)].map((m) => ({ pos: Number(m[1]), color: hex(m[2])! }));
    if (!stops.length) return undefined;
    if (/type="path"/.test(head)) return { kind: "gradient", path: { left: num("left") ?? 0, right: num("right") ?? 0, top: num("top") ?? 0, bottom: num("bottom") ?? 0 }, stops };
    return { kind: "gradient", degree: num("degree") ?? 0, stops };
  }
  const type = /patternType="([^"]+)"/.exec(xml)?.[1];
  if (!type || type === "none") return undefined;
  const fg = hex(/<fgColor[^>]*rgb="([0-9A-Fa-f]{8})"/.exec(xml)?.[1]) ?? "#000000";
  const bg = hex(/<bgColor[^>]*rgb="([0-9A-Fa-f]{8})"/.exec(xml)?.[1]);
  return { kind: "pattern", type, fg, bg };
}

/** The file's XML for a pattern or gradient, as Excel writes it (fgColor the pattern's colour, bgColor the cell's). */
export function fillXmlOf(f: FancyFill): string {
  if (f.kind === "pattern") {
    return `<fill><patternFill patternType="${f.type}"><fgColor rgb="${argb(f.fg)}"/>${f.bg ? `<bgColor rgb="${argb(f.bg)}"/>` : '<bgColor indexed="64"/>'}</patternFill></fill>`;
  }
  const stops = f.stops.map((s) => `<stop position="${s.pos}"><color rgb="${argb(s.color)}"/></stop>`).join("");
  if (f.path) return `<fill><gradientFill type="path" left="${f.path.left}" right="${f.path.right}" top="${f.path.top}" bottom="${f.path.bottom}">${stops}</gradientFill></fill>`;
  return `<fill><gradientFill${f.degree ? ` degree="${f.degree}"` : ""}>${stops}</gradientFill></fill>`;
}

const tiles = new Map<string, HTMLCanvasElement>();

/** The pattern's 8 x 8 tile (one CSS pixel per pattern pixel, as Excel draws them on screen). */
export function patternTile(type: string, fg: string, bg = "#FFFFFF"): HTMLCanvasElement {
  const key = `${type}|${fg}|${bg}`;
  let c = tiles.get(key);
  if (c) return c;
  c = document.createElement("canvas");
  c.width = c.height = 8;
  const g = c.getContext("2d")!;
  g.fillStyle = bg;
  g.fillRect(0, 0, 8, 8);
  g.fillStyle = fg;
  for (let y = 0; y < 8; y++) for (let x = 0; x < 8; x++) if (on(type, x, y)) g.fillRect(x, y, 1, 1);
  tiles.set(key, c);
  return c;
}

/** What the grid paints a cell with for a pattern or gradient fill in the rectangle (x, y, w, h). */
export function canvasFill(ctx: CanvasRenderingContext2D, f: FancyFill, x: number, y: number, w: number, h: number): string | CanvasPattern | CanvasGradient {
  if (f.kind === "pattern") {
    if (f.type === "solid") return f.fg;
    const p = ctx.createPattern(patternTile(f.type, f.fg, f.bg), "repeat");
    return p ?? f.fg;
  }
  let g: CanvasGradient;
  if (f.path) {
    const cx = x + w * (f.path.left + f.path.right) / 2, cy = y + h * (f.path.top + f.path.bottom) / 2;
    const r = Math.hypot(Math.max(cx - x, x + w - cx), Math.max(cy - y, y + h - cy));
    g = ctx.createRadialGradient(cx, cy, 0, cx, cy, r);
  } else {
    const a = ((f.degree ?? 0) * Math.PI) / 180, dx = Math.cos(a), dy = Math.sin(a);
    const half = (Math.abs(dx) * w + Math.abs(dy) * h) / 2, cx = x + w / 2, cy = y + h / 2;
    g = ctx.createLinearGradient(cx - dx * half, cy - dy * half, cx + dx * half, cy + dy * half);
  }
  for (const s of f.stops) g.addColorStop(Math.max(0, Math.min(1, s.pos)), s.color);
  return g;
}

/** The same fill as a CSS background (the dialog's Beispiel and the variant previews). */
export function cssFill(f: FancyFill | undefined, solid?: string): string {
  if (!f) return solid ?? "#ffffff";
  if (f.kind === "pattern") return f.type === "solid" ? f.fg : `url(${patternTile(f.type, f.fg, f.bg).toDataURL()}) repeat`;
  const stops = f.stops.map((s) => `${s.color} ${Math.round(s.pos * 100)}%`).join(", ");
  if (f.path) return `radial-gradient(circle at ${Math.round((f.path.left + f.path.right) * 50)}% ${Math.round((f.path.top + f.path.bottom) * 50)}%, ${stops})`;
  // CSS angles run clockwise from the top; Excel's degrees from the left edge towards the bottom.
  return `linear-gradient(${(f.degree ?? 0) + 90}deg, ${stops})`;
}

/** Fülleffekte's Schattierungsarten and their variants (Varianten) for two colours. */
export const SHADINGS = ["Horizontal", "Vertikal", "Diagonal oben", "Diagonal unten", "Aus der Ecke", "Aus der Mitte"] as const;
export type Shading = (typeof SHADINGS)[number];

export function shadingFill(kind: Shading, variant: number, c1: string, c2: string): GradientFill {
  const two = (a: string, b: string) => [{ pos: 0, color: a }, { pos: 1, color: b }];
  const three = (a: string, b: string) => [{ pos: 0, color: a }, { pos: 0.5, color: b }, { pos: 1, color: a }];
  const stops = variant === 0 ? two(c1, c2) : variant === 1 ? two(c2, c1) : variant === 2 ? three(c1, c2) : three(c2, c1);
  if (kind === "Aus der Ecke") {
    const corner = [[0, 0], [1, 0], [0, 1], [1, 1]][variant];
    return { kind: "gradient", path: { left: corner[0], right: corner[0], top: corner[1], bottom: corner[1] }, stops: two(c1, c2) };
  }
  if (kind === "Aus der Mitte") return { kind: "gradient", path: { left: 0.5, right: 0.5, top: 0.5, bottom: 0.5 }, stops: variant === 0 ? two(c1, c2) : two(c2, c1) };
  const degree = { Horizontal: 90, Vertikal: 0, "Diagonal oben": 45, "Diagonal unten": 135 }[kind];
  return { kind: "gradient", degree, stops };
}
