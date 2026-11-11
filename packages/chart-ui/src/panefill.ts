// The Format pane's "Füllung und Linie" and "Effekte" tabs, shared by every chart element (D-066,
// D-111): Füllung (Keine, Einfarbig, Farbverlauf, Bild- oder Texturfüllung, Muster, Automatisch),
// Rahmen / Linie (Keine, Einfarbig, Farbverlaufslinie, Automatisch with Breite, Verbundtyp,
// Strichtyp, Abschlusstyp, Anschlusstyp, Pfeile), Schatten, Leuchten, Weiche Kanten and 3D-Format.
// Every control changes the element at once as one undo step.

import { OFFICE_ACCENTS, PATTERNS, lum, patternPath, type Arrow, type Dash, type Effects, type Fill, type Format3D, type Gradient, type Line } from "@raabe/charts";
import { pictures as pic, type PaneControl, type PaneSection, type PictureEntry } from "@raabe/shell";
import { styleTheme } from "./styles";
import { TEXTURES, textureSrc } from "./textures";

/** What a section edits: an element's fill, line and effects (series, axes, the chart area …). */
export interface FmtTarget {
  /** A key for the pane's own state (the chosen gradient stop). */
  key: string;
  get: () => { fill?: Fill; line?: Line; effects?: Effects } | undefined;
  /** Changes the element as one undo step (created on first use). */
  change: (fn: (f: { fill?: Fill; line?: Line; effects?: Effects }) => void) => void;
}

export const DASHES: { label: string; value: Dash }[] = [
  { label: "Uni", value: "solid" }, { label: "Runder Punkt", value: "sysDot" }, { label: "Viereckiger Punkt", value: "sysDash" },
  { label: "Strich", value: "dash" }, { label: "Strichpunkt", value: "dashDot" }, { label: "Langer Strich", value: "lgDash" },
  { label: "Langer Strichpunkt", value: "lgDashDot" }, { label: "Langer Strichpunkt-Punkt", value: "lgDashDotDot" },
];

const COMPOUND = [
  { label: "Einfach", value: "sng" }, { label: "Doppelt", value: "dbl" }, { label: "Dick Dünn", value: "thickThin" }, { label: "Dünn Dick", value: "thinThick" }, { label: "Dreifach", value: "tri" },
];
const CAPS = [{ label: "Rechteckig", value: "sq" }, { label: "Rund", value: "rnd" }, { label: "Flach", value: "flat" }];
const JOINS = [{ label: "Rund", value: "round" }, { label: "Abschrägung", value: "bevel" }, { label: "Winkel", value: "miter" }];
const ARROWS = [
  { label: "Kein Pfeil", value: "none" }, { label: "Pfeil", value: "triangle" }, { label: "Offener Pfeil", value: "arrow" }, { label: "Rahmenpfeil", value: "stealth" },
  { label: "Pfeil mit Raute", value: "diamond" }, { label: "Ovaler Pfeil", value: "oval" },
];
const SIZES = ["sm", "med", "lg"] as const;
/** Pfeilgröße 1-9: width small to large, each with length short to long (Excel's order). */
const arrowSizes = (side: "L" | "R") => SIZES.flatMap((w, i) => SIZES.map((len, j) => ({ label: `Pfeil ${side} Größe ${i * 3 + j + 1}`, value: `${w},${len}` })));

/**
 * The pictures of the pane's picture dropdowns (D-161), in the sizes Excel's lists have (recorded 2026-11-11,
 * results/2026-11-11-excel-picturedrops.json): lines and arrows as wide strips in a list, effects as tiles in a grid.
 */
export const DASH_PIC: [number, number] = [104, 20];
export const COMPOUND_PIC: [number, number] = [72, 24];
export const ARROW_PIC: [number, number] = [33, 33];
/** Schatten, Leuchten, Abschrägung, Material, Beleuchtung; Weiche Kanten; Spiegelung; Farbverläufe and Richtung; Textur. */
export const TILE_PIC: [number, number] = [56, 56];
const SOFT_PIC: [number, number] = [44, 44];
const REFLECT_PIC: [number, number] = [66, 66];
const GRADIENT_PIC: [number, number] = [62, 62];
const TEXTURE_PIC: [number, number] = [88, 88];
const entries = (list: { label: string; value: string }[], draw: (v: string) => string): PictureEntry[] => list.map((o) => ({ ...o, svg: draw(o.value) }));
export const compoundEntries = () => entries(COMPOUND, (v) => pic.compoundPicture(v, COMPOUND_PIC[0], COMPOUND_PIC[1]));
export const dashEntries = () => entries(DASHES, (v) => pic.dashPicture(v, DASH_PIC[0], DASH_PIC[1]));
const arrowEntries = (end: "start" | "end") => entries(ARROWS, (v) => pic.arrowPicture(v, end, ARROW_PIC[0], ARROW_PIC[1]));
const arrowSizeEntries = (side: "L" | "R") => entries(arrowSizes(side), (v) => { const [w, len] = v.split(","); return pic.arrowSizePicture(w, len, side === "L" ? "start" : "end", ARROW_PIC[0], ARROW_PIC[1]); });

/** A picture dropdown control of the pane. */
const picture = (label: string, options: PictureEntry[], columns: number, size: [number, number], value: () => string, set: (v: string) => void, disabled?: () => boolean): PaneControl =>
  ({ type: "picture", label, options, columns, size, value, set, disabled });

/** Excel's gradient directions of a linear Farbverlauf (Richtung) as angles. */
const DIRECTIONS = [
  { label: "Linear diagonal - oben links nach unten rechts", value: "45" }, { label: "Linear unten", value: "90" },
  { label: "Linear diagonal - oben rechts nach unten links", value: "135" }, { label: "Linear rechts", value: "0" },
  { label: "Linear links", value: "180" }, { label: "Linear diagonal - unten links nach oben rechts", value: "315" },
  { label: "Linear oben", value: "270" }, { label: "Linear diagonal - unten rechts nach oben links", value: "225" },
];

/**
 * The Musterfüllung grid: Excel's 48 patterns in its order (8 per row) under its names (recorded 2026-11-09,
 * results/2026-11-09-excel-chart-pane.json); the six other OOXML presets (horz, vert, cross …) stay readable from
 * files but are not offered, as in Excel (D-161).
 */
const PATTERN_ORDER: [string, string][] = [
  ["pct5", "Gepunktet: 5 %"], ["pct10", "Gepunktet: 10 %"], ["pct20", "Gepunktet: 20 %"], ["pct25", "Gepunktet: 25 %"], ["pct30", "Gepunktet: 30 %"],
  ["pct40", "Gepunktet: 40 %"], ["pct50", "Gepunktet: 50 %"], ["pct60", "Gepunktet: 60 %"], ["pct70", "Gepunktet: 70 %"], ["pct75", "Gepunktet: 75 %"],
  ["pct80", "Gepunktet: 80 %"], ["pct90", "Gepunktet: 90 %"], ["ltDnDiag", "Diagonale Streifen: hell nach unten"], ["ltUpDiag", "Diagonale Streifen: hell nach oben"],
  ["dkDnDiag", "Diagonale Streifen: dunkel nach unten"], ["dkUpDiag", "Diagonale Streifen: dunkel nach oben"], ["wdDnDiag", "Diagonale Streifen: weit nach unten"],
  ["wdUpDiag", "Diagonale Streifen: weit nach oben"], ["ltVert", "Vertikale Streifen: hell"], ["ltHorz", "Horizontale Streifen: hell"], ["narVert", "Vertikale Streifen: dünn"],
  ["narHorz", "Horizontale Streifen: dünn"], ["dkVert", "Vertikale Streifen: dunkel"], ["dkHorz", "Horizontale Streifen: dunkel"], ["dashDnDiag", "Streifen: umgekehrte Schrägstriche"],
  ["dashUpDiag", "Streifen: Schrägstriche"], ["dashHorz", "Horizontale Streifen: alternierend horizontale Linien"], ["dashVert", "Vertikale Streifen: alternierend vertikale Linien"],
  ["smConfetti", "Kleines Konfetti"], ["lgConfetti", "Große Konfetti"], ["zigZag", "Zickzack"], ["wave", "Welle"], ["diagBrick", "Diagonale Steine"], ["horzBrick", "Horizontale Steine"],
  ["weave", "Kleine Welle"], ["plaid", "Geflochten"], ["divot", "Krähenfüße"], ["dotGrid", "Gepunktetes Raster"], ["dotDmnd", "Raster aus gepunkteten Rauten"], ["shingle", "Dachplatten"],
  ["trellis", "Gitter"], ["sphere", "Kugeln"], ["smGrid", "Kleines Raster"], ["lgGrid", "Großes Raster"], ["smCheck", "Kleine Schachfelder"], ["lgCheck", "Große Schachfelder"],
  ["openDmnd", "Raster aus konturierten Rauten"], ["solidDmnd", "Raster aus ausgefüllten Rauten"],
];
export const PATTERN_OPTIONS = PATTERN_ORDER.filter(([k]) => k in PATTERNS).map(([k, label]) => ({ value: k, label, path: patternPath(k) }));

/** The chosen gradient stop per element (the pane's own state, not the chart's). */
const chosenStop = new Map<string, number>();

/** Excel's default gradient for a colour (light to darker, top to bottom). */
export function defaultGradient(color: string): Gradient {
  return { kind: "linear", angle: 90, stops: [
    { pos: 0, color: lum(color, 0.05, 0.95) }, { pos: 74, color: lum(color, 0.45, 0.55) }, { pos: 83, color: lum(color, 0.45, 0.55) }, { pos: 100, color: lum(color, 0.3, 0.7) },
  ] };
}

/**
 * Voreingestellte Farbverläufe (D-155): Excel's five gradients in each accent, as Excel 365 writes them
 * (probe_excel_chart_followup.py): stops in the accent with Helligkeit (lumMod/lumOff), linear or from a focus.
 */
const PRESET_GRADIENTS: { name: string; kind: Gradient["kind"]; angle?: number; focus?: Gradient["focus"]; stops: [number, number][] }[] = [
  { name: "Heller Farbverlauf", kind: "linear", angle: 90, stops: [[0, 95], [74, 55], [83, 55], [100, 70]] },
  { name: "Lichtakzent oben", kind: "radial", focus: { l: 50, t: -80, r: 50, b: 180 }, stops: [[0, 100], [35, 100], [100, 0]] },
  { name: "Mittlerer Farbverlauf", kind: "linear", angle: 270, stops: [[0, -33], [48, 3], [100, 40]] },
  { name: "Lichtakzent unten", kind: "radial", focus: { l: 50, t: 130, r: 50, b: -30 }, stops: [[0, 60], [46, 5], [100, -40]] },
  { name: "Radialer Farbverlauf", kind: "radial", focus: { l: 50, t: 50, r: 50, b: 50 }, stops: [[0, -11], [23, -11], [69, -25], [97, -30]] },
];

/** The 30 entries of Voreingestellte Farbverläufe ("Heller Farbverlauf - Akzent 1" …), row by row as Excel's gallery. */
export const GRADIENT_PRESETS = PRESET_GRADIENTS.flatMap((p) => [1, 2, 3, 4, 5, 6].map((k) => ({ label: `${p.name} - Akzent ${k}`, value: `${p.name}|${k}` })));

/** A preset gradient in the workbook's design colours. */
export function presetGradient(value: string): Gradient | undefined {
  const [name, k] = value.split("|");
  const p = PRESET_GRADIENTS.find((x) => x.name === name);
  if (!p) return undefined;
  const theme = `accent${k}`;
  const color = styleTheme()[theme] ?? OFFICE_ACCENTS[Number(k) - 1] ?? "#4472C4";
  const g: Gradient = { kind: p.kind, stops: p.stops.map(([pos, b]) => ({ pos, color, theme, ...(b ? { brightness: b } : {}) })) };
  if (p.angle !== undefined) g.angle = p.angle;
  if (p.focus && (p.focus.l !== 50 || p.focus.t !== 50 || p.focus.r !== 50 || p.focus.b !== 50)) g.focus = { ...p.focus };
  return g;
}

/** The pictures of Voreingestellte Farbverläufe: each preset in its accent, drawn as its stops run. */
function gradientPresetEntries(): PictureEntry[] {
  return GRADIENT_PRESETS.map((o) => {
    const g = presetGradient(o.value)!;
    const stops = g.stops.map((s) => ({ pos: s.pos, color: s.brightness ? lum(s.color, s.brightness > 0 ? 1 - s.brightness / 100 : 1 + s.brightness / 100, s.brightness > 0 ? s.brightness / 100 : 0) : s.color }));
    const focus = g.focus ? { x: (g.focus.l + 100 - g.focus.r) / 200, y: (g.focus.t + 100 - g.focus.b) / 200 } : undefined;
    return { ...o, svg: pic.gradientPicture(g.kind, stops, g.angle ?? 90, focus, ...GRADIENT_PIC) };
  });
}

/** The pictures of the Schatten presets, grouped as Excel's gallery (Außen, Innen, Perspektive). */
function shadowEntries(): PictureEntry[] {
  return SHADOW_PRESETS.map((p) => {
    const s = p.s;
    if (!s) return { label: "Kein Schatten", value: p.value, svg: pic.nonePicture(...TILE_PIC), group: "Kein Schatten", columns: 1 };
    const kind = s.persp ? "persp" : s.inner ? "inner" : "outer";
    const centre = !s.distance && !s.persp;
    return { label: p.label, value: p.value, svg: pic.shadowPicture(kind, s.angle ?? 45, ...TILE_PIC, centre), group: kind === "persp" ? "Perspektive" : kind === "inner" ? "Innen" : "Außen" };
  });
}

/** The preset a gradient is (its value), if it is one. */
function presetOf(g: Gradient): string {
  const key = (x: Gradient) => JSON.stringify([x.kind, x.angle ?? null, x.focus ?? null, x.stops.map((s) => [s.pos, s.theme ?? s.color, s.brightness ?? 0, s.transparency ?? 0])]);
  return GRADIENT_PRESETS.find((p) => { const q = presetGradient(p.value); return !!q && key(q) === key(g); })?.value ?? "";
}

/** The controls of a gradient (Typ, Richtung, Winkel, Farbverlaufstopps, Farbe, Position, Transparenz, Helligkeit, Mit Form drehen). */
function gradientControls(key: string, g: Gradient, set: (g: Gradient) => void): PaneControl[] {
  const sel = Math.min(chosenStop.get(key) ?? 0, g.stops.length - 1);
  const stop = g.stops[sel];
  const withStop = (patch: Partial<Gradient["stops"][number]>) => {
    const changedStop = { ...g.stops[sel], ...patch };
    const stops = g.stops.map((s, i) => (i === sel ? changedStop : s)).sort((a, b) => a.pos - b.pos);
    chosenStop.set(key, stops.indexOf(changedStop));
    set({ ...g, stops });
  };
  const preset = presetOf(g);
  const c: PaneControl[] = [
    picture("Voreingestellte Farbverläufe", gradientPresetEntries(), 6, GRADIENT_PIC, () => preset,
      (v) => { const p = presetGradient(v); if (p) { chosenStop.set(key, 0); set({ ...p, ...(g.rotWithShape === false ? { rotWithShape: false } : {}) }); } }),
    { type: "select", label: "Typ", options: [{ label: "Linear", value: "linear" }, { label: "Radial", value: "radial" }, { label: "Rechteckig", value: "rect" }, { label: "Pfad", value: "path" }],
      value: () => g.kind, set: (v) => set({ ...g, kind: v as Gradient["kind"], angle: v === "linear" ? g.angle ?? 90 : undefined }) },
  ];
  if (g.kind === "linear") {
    // The pictures show the gradient's own colours in each direction, as Excel's do.
    const stops = g.stops.map((s) => ({ pos: s.pos, color: s.color, opacity: s.transparency ? 1 - s.transparency / 100 : undefined }));
    c.push(picture("Richtung", entries(DIRECTIONS, (v) => pic.gradientPicture("linear", stops, Number(v), undefined, ...GRADIENT_PIC)), 5, GRADIENT_PIC, () => String(g.angle ?? 0), (v) => set({ ...g, angle: Number(v) })));
    c.push({ type: "number", label: "Winkel", unit: "°", min: 0, max: 359.9, step: 1, value: () => g.angle ?? 0, set: (v) => set({ ...g, angle: Math.round(v * 10) / 10 }) });
  }
  c.push({
    type: "gradient", label: "Farbverlaufstopps", stops: () => g.stops, selected: () => sel,
    select: (i) => { chosenStop.set(key, i); },
    add: () => {
      // A new stop between the chosen one and the next, in the chosen one's colour.
      const next = g.stops[sel + 1] ?? g.stops[sel - 1];
      const pos = Math.round(((stop?.pos ?? 0) + (next?.pos ?? 100)) / 2);
      // The stops stay in their order along the bar (as the file keeps them).
      const added = { pos, color: stop?.color ?? "#FFFFFF", ...(stop?.transparency ? { transparency: stop.transparency } : {}) };
      const stops = [...g.stops, added].sort((a, b) => a.pos - b.pos);
      chosenStop.set(key, stops.indexOf(added));
      set({ ...g, stops });
    },
    remove: () => {
      if (g.stops.length <= 2) return;
      chosenStop.set(key, Math.max(0, sel - 1));
      set({ ...g, stops: g.stops.filter((_, i) => i !== sel) });
    },
    move: (i, pos) => {
      const moved = { ...g.stops[i], pos };
      const stops = g.stops.map((s, k) => (k === i ? moved : s)).sort((a, b) => a.pos - b.pos);
      chosenStop.set(key, stops.indexOf(moved));
      set({ ...g, stops });
    },
  });
  if (stop) {
    // A colour picked from the palette is a colour of its own (no longer the design colour).
    c.push({ type: "color", label: "Farbe", value: () => stop.color, set: (hex) => withStop({ color: hex ?? stop.color, theme: undefined }) });
    c.push({ type: "number", label: "Position", unit: " %", min: 0, max: 100, slider: true, value: () => stop.pos, set: (v) => withStop({ pos: Math.round(v) }) });
    c.push({ type: "number", label: "Transparenz", unit: " %", min: 0, max: 100, slider: true, value: () => stop.transparency ?? 0, set: (v) => withStop({ transparency: Math.round(v) || undefined }) });
    c.push({ type: "number", label: "Helligkeit", unit: " %", min: -100, max: 100, slider: true, value: () => stop.brightness ?? 0, set: (v) => withStop({ brightness: Math.round(v) || undefined }) });
  }
  c.push({ type: "check", label: "Mit Form drehen", value: () => g.rotWithShape !== false, set: (on) => set({ ...g, rotWithShape: on ? undefined : false }) });
  return c;
}

/** Bild- oder Texturfüllung: Einfügen... (a picture file), Zwischenablage, Transparenz, nebeneinander als Textur. */
function pictureControls(f: Fill, set: (f: Fill) => void): PaneControl[] {
  const p = f.picture;
  const texture = (p?.src && TEXTURES.find((t) => textureSrc(t) === p.src)) || "";
  const choose = () => {
    const input = document.createElement("input");
    input.type = "file";
    input.accept = "image/png,image/jpeg,image/gif,image/bmp";
    input.addEventListener("change", () => {
      const file = input.files?.[0];
      if (!file) return;
      const reader = new FileReader();
      reader.onload = () => set({ type: "picture", picture: { ...p, src: String(reader.result) } });
      reader.readAsDataURL(file);
    });
    input.click();
  };
  const paste = async () => {
    try {
      for (const item of await navigator.clipboard.read()) {
        const type = item.types.find((t) => t.startsWith("image/"));
        if (!type) continue;
        const blob = await item.getType(type);
        const reader = new FileReader();
        reader.onload = () => set({ type: "picture", picture: { ...p, src: String(reader.result) } });
        reader.readAsDataURL(blob);
        return;
      }
    } catch { /* no picture on the clipboard */ }
  };
  const c: PaneControl[] = [
    { type: "heading", label: "Bildquelle" },
    { type: "button", label: "Einfügen...", run: choose },
    { type: "button", label: "Zwischenablage", run: () => { void paste(); } },
    // The chosen texture is recognised by its picture (the file keeps no name, as in Excel).
    picture("Textur", TEXTURES.map((t) => ({ label: t, value: t, svg: pic.imagePicture(textureSrc(t) ?? "", ...TEXTURE_PIC) })), 5, TEXTURE_PIC, () => texture,
      (v) => { const src = textureSrc(v); if (src) set({ type: "picture", picture: { src, tile: true, ...(p?.transparency ? { transparency: p.transparency } : {}) } }); }),
    { type: "number", label: "Transparenz", unit: " %", min: 0, max: 100, slider: true, value: () => p?.transparency ?? 0, set: (v) => p && set({ type: "picture", picture: { ...p, transparency: Math.round(v) || undefined } }), disabled: () => !p },
    { type: "check", label: "Bild nebeneinander als Textur anordnen", value: () => !!p?.tile, set: (on) => p && set({ type: "picture", picture: { ...p, tile: on || undefined } }), disabled: () => !p },
  ];
  if (p && !p.tile) {
    const off = p.offsets ?? { l: 0, t: 0, r: 0, b: 0 };
    const side = (label: string, k: keyof typeof off): PaneControl => ({ type: "number", label, unit: " %", min: -100, max: 100, value: () => off[k], set: (v) => set({ type: "picture", picture: { ...p, offsets: { ...off, [k]: Math.round(v) } } }) });
    c.push(side("Offset links", "l"), side("Offset rechts", "r"), side("Offset oben", "t"), side("Offset unten", "b"));
  }
  return c;
}

/** The Füllung section. `extra` follows the type choice (Invertieren falls negativ, Punkte nach Farbe unterscheiden). */
/** `withAuto` false: no Automatisch (Textfüllung), Einfarbige Füllung in the automatic colour is the default. */
/**
 * `autoNone`: the element's automatic fill is none (data labels): Keine Füllung is shown chosen until another is
 * picked, as in Excel (D-160).
 */
export function fillSection(t: FmtTarget, autoColor: string, extra: PaneControl[] = [], open = true, withAuto = true, autoNone = false): PaneSection {
  return {
    title: "Füllung", open, controls: () => {
      const f = t.get()?.fill;
      const set = (fill: Fill | undefined) => t.change((x) => { if (fill) x.fill = fill; else delete x.fill; });
      const type = f?.type ?? (autoNone ? "none" : withAuto ? "auto" : "solid");
      const radio = (label: string, v: Fill["type"], make: () => Fill | undefined): PaneControl => ({ type: "radio", label, checked: () => type === v, pick: () => set(make()) });
      const color = f?.color ?? autoColor;
      const c: PaneControl[] = [
        radio("Keine Füllung", "none", () => (autoNone ? undefined : { type: "none" })),
        radio("Einfarbige Füllung", "solid", () => ({ type: "solid", color })),
        radio("Farbverlauf", "gradient", () => ({ type: "gradient", gradient: f?.gradient ?? defaultGradient(color) })),
        // Excel fills with its first texture at once (tiled), so the choice is saved even without a picture (D-155).
        radio("Bild- oder Texturfüllung", "picture", () => {
          const src = f?.picture?.src ? undefined : textureSrc(TEXTURES[0]);
          return { type: "picture", picture: f?.picture?.src ? f.picture : src ? { src, tile: true } : f?.picture };
        }),
        radio("Musterfüllung", "pattern", () => ({ type: "pattern", pattern: f?.pattern ?? { preset: "pct5", fg: color, bg: "#FFFFFF" } })),
        ...(withAuto ? [radio("Automatisch", "auto", () => (autoNone ? { type: "auto" } : undefined))] : []),
        ...extra,
      ];
      // Excel shows the colour under Automatisch too; picking one makes the fill Einfarbig.
      if (type === "auto") c.push({ type: "color", label: "Farbe", value: () => null, auto: autoColor, set: (hex) => set({ type: "solid", color: hex ?? autoColor }) });
      if (type === "solid") {
        const sf: Fill = f?.type === "solid" ? f : { type: "solid", color: autoColor };
        c.push({ type: "color", label: "Farbe", value: () => sf.color ?? null, set: (hex) => set({ ...sf, color: hex ?? autoColor }) });
        c.push({ type: "number", label: "Transparenz", unit: " %", min: 0, max: 100, slider: true, value: () => sf.transparency ?? 0, set: (v) => set({ ...sf, transparency: Math.round(v) || undefined }) });
      } else if (f?.type === "gradient" && f.gradient) {
        c.push(...gradientControls(`${t.key}/fill`, f.gradient, (g) => set({ type: "gradient", gradient: g })));
      } else if (f?.type === "picture") {
        c.push(...pictureControls(f, set));
      } else if (f?.type === "pattern" && f.pattern) {
        const p = f.pattern;
        c.push({ type: "pattern", label: "Muster", options: PATTERN_OPTIONS, value: () => p.preset, fg: () => p.fg, bg: () => p.bg, set: (v) => set({ type: "pattern", pattern: { ...p, preset: v } }) });
        c.push({ type: "color", label: "Vordergrund", value: () => p.fg, set: (hex) => set({ type: "pattern", pattern: { ...p, fg: hex ?? p.fg } }) });
        c.push({ type: "color", label: "Hintergrund", value: () => p.bg, set: (hex) => set({ type: "pattern", pattern: { ...p, bg: hex ?? p.bg } }) });
      }
      return c;
    },
  };
}

export interface LineOptions {
  /** "Rahmen" (shapes) or "Linie" (lines, axes, gridlines). */
  title: string;
  autoColor: string;
  autoWidth: number;
  /** Pfeile (Start- / Endpfeiltyp): lines that have ends; "disabled": shown greyed, as Excel shows them on a Rahmen. */
  arrows?: boolean | "disabled";
  /** The element's automatic line is none (data labels, scatter series without lines): Keine Linie is shown chosen. */
  autoNone?: boolean;
  /** Controls after the type choice or at the end (Abgerundete Ecken, Linie glätten). */
  extra?: PaneControl[];
  open?: boolean;
  /** false: no Automatisch (Textkontur); no line is the default. */
  withAuto?: boolean;
}

/** The Rahmen / Linie section. */
export function lineSection(t: FmtTarget, o: LineOptions): PaneSection {
  return {
    title: o.title, open: o.open, controls: () => {
      const l = t.get()?.line;
      const set = (line: Line | undefined) => t.change((x) => { if (line) x.line = line; else delete x.line; });
      const patch = (p: Partial<Line>) => {
        const cur: Line = !l || l.type === "auto" ? { ...l, type: "solid", color: l?.color ?? o.autoColor } : l;
        const next: Line = { ...cur, ...p };
        for (const k of Object.keys(next) as (keyof Line)[]) if (next[k] === undefined) delete next[k];
        set(next);
      };
      const type = l?.type ?? (o.withAuto === false || o.autoNone ? "none" : "auto");
      const color = l?.color ?? o.autoColor;
      const c: PaneControl[] = [
        { type: "radio", label: "Keine Linie", checked: () => type === "none", pick: () => set(o.withAuto === false || o.autoNone ? undefined : { type: "none" }) },
        { type: "radio", label: "Einfarbige Linie", checked: () => type === "solid", pick: () => set({ ...l, type: "solid", color, width: l?.width ?? o.autoWidth }) },
        { type: "radio", label: "Farbverlaufslinie", checked: () => type === "gradient", pick: () => set({ ...l, type: "gradient", gradient: l?.gradient ?? defaultGradient(color), width: l?.width ?? o.autoWidth }) },
        ...(o.withAuto === false ? [] : [{ type: "radio", label: "Automatisch", checked: () => type === "auto", pick: () => set(o.autoNone ? { type: "auto" } : undefined) } as PaneControl]),
      ];
      // Excel shows Farbe and Transparenz under Automatisch too; changing them makes the line Einfarbig.
      if (type === "auto") {
        c.push({ type: "color", label: "Farbe", value: () => null, auto: o.autoColor, set: (hex) => patch({ color: hex ?? o.autoColor }) });
        c.push({ type: "number", label: "Transparenz", unit: " %", min: 0, max: 100, slider: true, value: () => 0, set: (v) => patch({ transparency: Math.round(v) || undefined }) });
      }
      if (type === "solid") {
        c.push({ type: "color", label: "Farbe", value: () => l?.color ?? null, set: (hex) => patch({ color: hex ?? o.autoColor }) });
        c.push({ type: "number", label: "Transparenz", unit: " %", min: 0, max: 100, slider: true, value: () => l?.transparency ?? 0, set: (v) => patch({ transparency: Math.round(v) || undefined }) });
      }
      if (type === "gradient" && l?.gradient) c.push(...gradientControls(`${t.key}/line`, l.gradient, (g) => patch({ gradient: g })));
      if (type !== "none") {
        c.push({ type: "number", label: "Stärke", unit: " Pt.", min: 0, max: 1584, step: 0.25, decimals: 2, value: () => l?.width ?? o.autoWidth, set: (v) => patch({ width: v }) });
        c.push(picture("Verbundtyp", compoundEntries(), 1, COMPOUND_PIC, () => l?.compound ?? "sng", (v) => patch({ compound: v === "sng" ? undefined : (v as Line["compound"]) })));
        c.push(picture("Strichtyp", dashEntries(), 1, DASH_PIC, () => l?.dash ?? "solid", (v) => patch({ dash: v === "solid" ? undefined : (v as Dash) })));
        c.push({ type: "select", label: "Abschlusstyp", options: CAPS, value: () => l?.cap ?? "flat", set: (v) => patch({ cap: v === "flat" ? undefined : (v as Line["cap"]) }) });
        c.push({ type: "select", label: "Anschlusstyp", options: JOINS, value: () => l?.join ?? "round", set: (v) => patch({ join: v === "round" ? undefined : (v as Line["join"]) }) });
        if (o.arrows) {
          const arrow = (which: "head" | "tail", label: string, sizeLabel: string): PaneControl[] => {
            const a = l?.[which];
            return [
              picture(label, arrowEntries(which === "head" ? "start" : "end"), 3, ARROW_PIC, () => a?.type ?? "none",
                (v) => patch({ [which]: v === "none" ? undefined : { ...a, type: v as Arrow["type"] } }), () => o.arrows === "disabled"),
              picture(sizeLabel, arrowSizeEntries(which === "head" ? "L" : "R"), 3, ARROW_PIC, () => `${a?.w ?? "med"},${a?.len ?? "med"}`,
                (v) => { const [w, len] = v.split(",") as [Arrow["w"], Arrow["len"]]; if (a) patch({ [which]: { ...a, w, len } }); }, () => !a || o.arrows === "disabled"),
            ];
          };
          c.push(...arrow("head", "Startpfeiltyp", "Startpfeilgröße"), ...arrow("tail", "Endpfeiltyp", "Endpfeilgröße"));
        }
      }
      // Keine Linie hides the line's own options too (Abgerundete Ecken, Linie glätten), as in Excel.
      if (type !== "none") c.push(...(o.extra ?? []));
      return c;
    },
  };
}

export const SHADOW_PRESETS: { label: string; value: string; s?: Effects["shadow"] }[] = [
  { label: "Ohne", value: "none" },
  ...([["unten rechts", 45], ["unten", 90], ["unten links", 135], ["rechts", 0], ["Mitte", -1], ["links", 180], ["oben rechts", 315], ["oben", 270], ["oben links", 225]] as const).map(([n, a]) =>
    ({ label: `Offset: ${n}`, value: `o-${n}`, s: a < 0 ? { color: "#000000", transparency: 60, size: 102, blur: 5, angle: 0, distance: 0 } : { color: "#000000", transparency: 60, blur: 4, angle: a, distance: 3 } })),
  ...([["oben links", 225], ["oben", 270], ["oben rechts", 315], ["links", 180], ["Mitte", -1], ["rechts", 0], ["unten links", 135], ["unten", 90], ["unten rechts", 45]] as const).map(([n, a]) =>
    ({ label: `Innen: ${n}`, value: `i-${n}`, s: a < 0 ? { inner: true, color: "#000000", transparency: 0, blur: 9, angle: 0, distance: 0 } : { inner: true, color: "#000000", transparency: 50, blur: 5, angle: a, distance: 4 } })),
  // Perspektive (D-155), as Excel 365 writes them (probe_excel_chart_followup.py): scaled, skewed about a corner.
  { label: "Perspektive: oben links", value: "p-oben links", s: { color: "#000000", transparency: 80, blur: 6, angle: 225, distance: 0, persp: { sy: 23, kx: 20, align: "br" } } },
  { label: "Perspektive: oben rechts", value: "p-oben rechts", s: { color: "#000000", transparency: 80, blur: 6, angle: 315, distance: 0, persp: { sy: 23, kx: -20, align: "bl" } } },
  { label: "Perspektive: unterhalb", value: "p-unterhalb", s: { color: "#000000", transparency: 85, blur: 12, angle: 90, distance: 25, persp: { sx: 90, sy: -19, align: "b" } } },
  { label: "Perspektive: unten links", value: "p-unten links", s: { color: "#000000", transparency: 80, blur: 6, angle: 135, distance: 1, persp: { sy: -23, kx: 13.34, align: "br" } } },
  { label: "Perspektive: unten rechts", value: "p-unten rechts", s: { color: "#000000", transparency: 80, blur: 6, angle: 45, distance: 1, persp: { sy: -23, kx: -13.34, align: "bl" } } },
];

/** JSON with sorted keys: a shadow read from the file matches its preset whatever the order of its fields. */
const canon = (v: unknown): string => JSON.stringify(v, (_k, x) => (x && typeof x === "object" && !Array.isArray(x) ? Object.fromEntries(Object.entries(x).sort(([a], [b]) => a.localeCompare(b))) : x));

const BEVELS = [
  { label: "Keine", value: "none" }, { label: "Rund", value: "circle" }, { label: "Leichte Absenkung", value: "relaxedInset" }, { label: "Säule", value: "cross" },
  { label: "Schräge", value: "coolSlant" }, { label: "Starke Abschrägung", value: "angle" }, { label: "Weiche Abrundung", value: "softRound" },
  { label: "Konvex", value: "convex" }, { label: "Silhouette", value: "slope" }, { label: "Ringförmige Absenkung", value: "divot" }, { label: "Säule auf abgerundetem Sockel", value: "riblet" },
  { label: "Gewölbter Rand", value: "hardEdge" }, { label: "Rund konvex", value: "artDeco" },
];
const MATERIALS = [
  { label: "Matt", value: "matte" }, { label: "Warm matt", value: "warmMatte" }, { label: "Plastik", value: "plastic" }, { label: "Metall", value: "metal" },
  { label: "Dunkle Kante", value: "dkEdge" }, { label: "Helle Kante", value: "softEdge" }, { label: "Flach", value: "flat" }, { label: "Drahtmodell", value: "legacyWireframe" },
  { label: "Pulver", value: "powder" }, { label: "Durchsichtiges Pulver", value: "translucentPowder" }, { label: "Transparent", value: "clear" },
];
const LIGHTS = [
  { label: "Drei Lichtpunkte", value: "threePt" }, { label: "Gleichmäßig", value: "balanced" }, { label: "Weich", value: "soft" }, { label: "Hart", value: "harsh" },
  { label: "Flutlicht", value: "flood" }, { label: "Kontrast", value: "contrasting" }, { label: "Früher Morgen", value: "morning" }, { label: "Sonnenaufgang", value: "sunrise" },
  { label: "Sonnenuntergang", value: "sunset" }, { label: "Kühl", value: "chilly" }, { label: "Kalt", value: "freezing" }, { label: "Flach", value: "flat" },
  { label: "Zwei Lichtpunkte", value: "twoPt" }, { label: "Leuchteffekt", value: "glow" }, { label: "Heller Raum", value: "brightRoom" },
];

export interface EffectOptions { shadow?: boolean; glow?: boolean; soft?: boolean; format3d?: boolean; reflection?: boolean; autoColor: string }

/** The Effekte tab's sections. */
export function effectSections(t: FmtTarget, o: EffectOptions): PaneSection[] {
  const fx = () => t.get()?.effects;
  const set = (patch: Partial<Effects>) => t.change((x) => {
    const next: Effects = { ...x.effects, ...patch };
    for (const k of Object.keys(next) as (keyof Effects)[]) if (next[k] === undefined) delete next[k];
    if (Object.keys(next).length) x.effects = next; else delete x.effects;
  });
  const out: PaneSection[] = [];
  if (o.shadow !== false) out.push({ title: "Schatten", controls: () => {
    const s = fx()?.shadow;
    const preset = !s ? "none" : SHADOW_PRESETS.find((p) => p.s && canon(p.s) === canon(s))?.value ?? "custom";
    const patch = (p: Partial<NonNullable<Effects["shadow"]>>) => set({ shadow: { ...(s ?? SHADOW_PRESETS[1].s!), ...p } });
    const c: PaneControl[] = [
      picture("Voreinstellungen", shadowEntries(), 3, TILE_PIC, () => preset, (v) => set({ shadow: v === "none" ? undefined : { ...SHADOW_PRESETS.find((p) => p.value === v)!.s! } })),
      { type: "color", label: "Farbe", value: () => s?.color ?? "#000000", set: (hex) => patch({ color: hex ?? "#000000" }) },
      { type: "number", label: "Transparenz", unit: " %", min: 0, max: 100, slider: true, value: () => s?.transparency ?? 60, set: (v) => patch({ transparency: Math.round(v) }) },
      { type: "number", label: "Größe", unit: " %", min: 1, max: 200, slider: true, value: () => s?.size ?? 100, set: (v) => patch({ size: Math.round(v) === 100 ? undefined : Math.round(v) }), disabled: () => !!s?.inner },
      { type: "number", label: "Weichzeichnen", unit: " Pt.", min: 0, max: 100, slider: true, value: () => s?.blur ?? 4, set: (v) => patch({ blur: v }) },
      { type: "number", label: "Winkel", unit: "°", min: 0, max: 359.9, slider: true, value: () => s?.angle ?? 45, set: (v) => patch({ angle: v }) },
      { type: "number", label: "Abstand", unit: " Pt.", min: 0, max: 200, slider: true, value: () => s?.distance ?? 3, set: (v) => patch({ distance: v }) },
    ];
    return c;
  } });
  if (o.reflection) out.push({ title: "Spiegelung", controls: () => {
    const r = fx()?.reflection;
    // Enge, Halbe and Volle Spiegelung (size 35, 55, 90 %), touching or 4 / 8 Pt. away (Excel's gallery).
    const kinds: [string, number][] = [["Enge Spiegelung", 35], ["Halbe Spiegelung", 55], ["Volle Spiegelung", 90]];
    const gaps: [string, number][] = [["mit Berührung", 0], ["4 Pt. Offset", 4], ["8 Pt. Offset", 8]];
    const presets = gaps.flatMap(([g, d]) => kinds.map(([k, size]) => ({ label: `${k}: ${g}`, value: `${size},${d}` })));
    const current = r ? `${r.size},${r.distance}` : "none";
    const patch = (p: Partial<NonNullable<Effects["reflection"]>>) => set({ reflection: { transparency: 50, size: 55, blur: 0.5, distance: 0, ...r, ...p } });
    return [
      picture("Voreinstellungen", [
        { label: "Ohne", value: "none", svg: pic.nonePicture(...REFLECT_PIC), group: "Keine Spiegelung", columns: 1 },
        ...presets.map((p) => { const [size, distance] = p.value.split(",").map(Number); return { ...p, svg: pic.reflectionPicture(size, distance, ...REFLECT_PIC), group: "Spiegelungsvarianten" }; }),
      ], 3, REFLECT_PIC, () => current, (v) => { if (v === "none") set({ reflection: undefined }); else { const [size, distance] = v.split(",").map(Number); set({ reflection: { transparency: 50, size, blur: 0.5, distance } }); } }),
      { type: "number", label: "Transparenz", unit: " %", min: 0, max: 100, slider: true, value: () => r?.transparency ?? 50, set: (v) => patch({ transparency: Math.round(v) }) },
      { type: "number", label: "Größe", unit: " %", min: 0, max: 100, slider: true, value: () => r?.size ?? 0, set: (v) => patch({ size: Math.round(v) }) },
      { type: "number", label: "Weichzeichnen", unit: " Pt.", min: 0, max: 100, slider: true, value: () => r?.blur ?? 0, set: (v) => patch({ blur: v }) },
      { type: "number", label: "Abstand", unit: " Pt.", min: 0, max: 100, slider: true, value: () => r?.distance ?? 0, set: (v) => patch({ distance: v }) },
    ];
  } });
  if (o.glow !== false) out.push({ title: "Leuchten", controls: () => {
    const g = fx()?.glow;
    // Leuchteffekt in 5, 8, 11 and 18 Pt. in each accent (60 % transparent), as Excel's gallery.
    const presets = [5, 8, 11, 18].flatMap((size) => OFFICE_ACCENTS.map((color, k) => ({ label: `Leuchteffekt: ${size} Pt.; Akzentfarbe ${k + 1}`, value: `${size},${color}` })));
    const current = g ? `${g.size},${g.color}` : "none";
    return [
      picture("Voreinstellungen", [
        { label: "Ohne", value: "none", svg: pic.nonePicture(...TILE_PIC), group: "Kein Leuchteffekt", columns: 1 },
        ...presets.map((p) => { const [size, color] = p.value.split(","); return { ...p, svg: pic.glowPicture(Number(size), color, ...TILE_PIC), group: "Leuchteffektvarianten" }; }),
      ], 6, TILE_PIC, () => current, (v) => { if (v === "none") set({ glow: undefined }); else { const [size, color] = v.split(","); set({ glow: { color, size: Number(size), transparency: 60 } }); } }),
      { type: "color", label: "Farbe", value: () => g?.color ?? o.autoColor, set: (hex) => set({ glow: { size: g?.size ?? 5, transparency: g?.transparency ?? 60, color: hex ?? o.autoColor } }) },
      { type: "number", label: "Größe", unit: " Pt.", min: 0, max: 150, slider: true, value: () => g?.size ?? 0, set: (v) => set({ glow: v ? { color: g?.color ?? o.autoColor, transparency: g?.transparency ?? 60, size: v } : undefined }) },
      { type: "number", label: "Transparenz", unit: " %", min: 0, max: 100, slider: true, value: () => g?.transparency ?? 60, set: (v) => set({ glow: { color: g?.color ?? o.autoColor, size: g?.size ?? 5, transparency: Math.round(v) } }) },
    ];
  } });
  if (o.soft !== false) out.push({ title: "Weiche Kanten", controls: () => {
    const v = fx()?.softEdges ?? 0;
    const presets = [0, 1, 2.5, 5, 10, 25, 50];
    return [
      picture("Voreinstellungen", presets.map((p) => ({ label: p ? `${String(p).replace(".", ",")} Punkt` : "Keine weichen Kanten", value: String(p), svg: pic.softEdgePicture(p, ...SOFT_PIC), ...(p ? { group: "Varianten weicher Kanten" } : { group: "Keine weichen Kanten", columns: 1 }) })),
        3, SOFT_PIC, () => (presets.includes(v) ? String(v) : ""), (x) => set({ softEdges: Number(x) || undefined })),
      { type: "number", label: "Größe", unit: " Pt.", min: 0, max: 100, slider: true, value: () => v, set: (x) => set({ softEdges: x || undefined }) },
    ];
  } });
  if (o.format3d) out.push({ title: "3D-Format", controls: () => {
    const f = fx()?.format3d;
    const set3 = (p: Partial<Format3D>) => {
      const next: Format3D = { ...f, ...p };
      for (const k of Object.keys(next) as (keyof Format3D)[]) if (next[k] === undefined) delete next[k];
      set({ format3d: Object.keys(next).length ? next : undefined });
    };
    const bevel = (which: "top" | "bottom", label: string): PaneControl[] => {
      const b = f?.[which];
      return [
        picture(label, BEVELS.map((x) => ({ ...x, svg: pic.bevelPicture(x.value, ...TILE_PIC), ...(x.value === "none" ? { group: "Keine Abschrägung", columns: 1 } : { group: "Abschrägung" }) })), 4, TILE_PIC,
          () => b?.type ?? "none", (v) => set3({ [which]: v === "none" ? undefined : { w: b?.w ?? 6, h: b?.h ?? 6, type: v } })),
        { type: "number", label: "Breite", unit: " Pt.", min: 0, max: 1584, value: () => b?.w ?? 0, set: (v) => b && set3({ [which]: { ...b, w: v } }), disabled: () => !b },
        { type: "number", label: "Höhe", unit: " Pt.", min: 0, max: 1584, value: () => b?.h ?? 0, set: (v) => b && set3({ [which]: { ...b, h: v } }), disabled: () => !b },
      ];
    };
    return [
      ...bevel("top", "Abschrägung oben"), ...bevel("bottom", "Abschrägung unten"),
      { type: "color", label: "Tiefe", value: () => f?.depth?.color ?? null, set: (hex) => set3({ depth: { size: f?.depth?.size ?? 0, ...(hex ? { color: hex } : {}) } }) },
      { type: "number", label: "Größe", unit: " Pt.", min: 0, max: 1584, value: () => f?.depth?.size ?? 0, set: (v) => set3({ depth: v ? { ...f?.depth, size: v } : undefined }) },
      { type: "color", label: "Kontur", value: () => f?.contour?.color ?? null, set: (hex) => set3({ contour: { size: f?.contour?.size ?? 0, ...(hex ? { color: hex } : {}) } }) },
      { type: "number", label: "Größe", unit: " Pt.", min: 0, max: 1584, value: () => f?.contour?.size ?? 0, set: (v) => set3({ contour: v ? { ...f?.contour, size: v } : undefined }) },
      picture("Material", entries(MATERIALS, (v) => pic.materialPicture(v, ...TILE_PIC)).map((e, i) => ({ ...e, group: i < 4 ? "Standard" : i < 8 ? "Spezialeffekt" : "Durchsichtig" })), 4, TILE_PIC, () => f?.material ?? "warmMatte", (v) => set3({ material: v === "warmMatte" ? undefined : v })),
      picture("Beleuchtung", entries(LIGHTS, (v) => pic.lightingPicture(v, ...TILE_PIC)).map((e, i) => ({ ...e, group: i < 6 ? "Neutral" : i < 9 ? "Warm" : i < 11 ? "Kalt" : "Spezialeffekt" })), 4, TILE_PIC, () => f?.lighting ?? "threePt", (v) => set3({ lighting: v === "threePt" ? undefined : v })),
      { type: "button", label: "Zurücksetzen", run: () => set({ format3d: undefined }) },
    ];
  } });
  return out;
}
