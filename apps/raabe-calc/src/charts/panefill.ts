// The Format pane's "Füllung und Linie" and "Effekte" tabs, shared by every chart element (D-066,
// D-111): Füllung (Keine, Einfarbig, Farbverlauf, Bild- oder Texturfüllung, Muster, Automatisch),
// Rahmen / Linie (Keine, Einfarbig, Farbverlaufslinie, Automatisch with Breite, Verbundtyp,
// Strichtyp, Abschlusstyp, Anschlusstyp, Pfeile), Schatten, Leuchten, Weiche Kanten and 3D-Format.
// Every control changes the element at once as one undo step.

import { PATTERNS, lum, patternPath, type Arrow, type Dash, type Effects, type Fill, type Format3D, type Gradient, type Line } from "@raabe/charts";
import type { PaneControl, PaneSection } from "@raabe/shell";

/** What a section edits: an element's fill, line and effects (series, axes, the chart area …). */
export interface FmtTarget {
  /** A key for the pane's own state (the chosen gradient stop). */
  key: string;
  get: () => { fill?: Fill; line?: Line; effects?: Effects } | undefined;
  /** Changes the element as one undo step (created on first use). */
  change: (fn: (f: { fill?: Fill; line?: Line; effects?: Effects }) => void) => void;
}

export const DASHES: { label: string; value: Dash }[] = [
  { label: "Durchgehend", value: "solid" }, { label: "Rundpunkt", value: "sysDot" }, { label: "Quadratpunkt", value: "sysDash" },
  { label: "Strich", value: "dash" }, { label: "Strich-Punkt", value: "dashDot" }, { label: "Langer Strich", value: "lgDash" },
  { label: "Langer Strich-Punkt", value: "lgDashDot" }, { label: "Langer Strich-Punkt-Punkt", value: "lgDashDotDot" },
];

const COMPOUND = [
  { label: "Einfach", value: "sng" }, { label: "Doppelt", value: "dbl" }, { label: "Dick-Dünn", value: "thickThin" }, { label: "Dünn-Dick", value: "thinThick" }, { label: "Dreifach", value: "tri" },
];
const CAPS = [{ label: "Quadrat", value: "sq" }, { label: "Rund", value: "rnd" }, { label: "Flach", value: "flat" }];
const JOINS = [{ label: "Rund", value: "round" }, { label: "Abgeschrägt", value: "bevel" }, { label: "Gehrung", value: "miter" }];
const ARROWS = [
  { label: "Kein Pfeil", value: "none" }, { label: "Pfeil", value: "triangle" }, { label: "Offener Pfeil", value: "arrow" }, { label: "Spitzer Pfeil", value: "stealth" },
  { label: "Rautenförmiger Pfeil", value: "diamond" }, { label: "Ovaler Pfeil", value: "oval" },
];
const SIZES = ["sm", "med", "lg"] as const;
const SIZE_NAME = { sm: "klein", med: "mittel", lg: "groß" };
const ARROW_SIZES = SIZES.flatMap((w) => SIZES.map((len) => ({ label: `Breite ${SIZE_NAME[w]}, Länge ${SIZE_NAME[len]}`, value: `${w},${len}` })));

/** Excel's gradient directions of a linear Farbverlauf (Richtung) as angles. */
const DIRECTIONS = [
  { label: "Linear diagonal - oben links nach unten rechts", value: "45" }, { label: "Linear nach unten", value: "90" },
  { label: "Linear diagonal - oben rechts nach unten links", value: "135" }, { label: "Linear nach rechts", value: "0" },
  { label: "Linear nach links", value: "180" }, { label: "Linear diagonal - unten links nach oben rechts", value: "315" },
  { label: "Linear nach oben", value: "270" }, { label: "Linear diagonal - unten rechts nach oben links", value: "225" },
];

/** The pattern names of the Musterfüllung grid (our own words; the grid shows the patterns). */
const PATTERN_NAMES: Record<string, string> = {
  pct5: "5 %", pct10: "10 %", pct20: "20 %", pct25: "25 %", pct30: "30 %", pct40: "40 %", pct50: "50 %", pct60: "60 %", pct70: "70 %", pct75: "75 %", pct80: "80 %", pct90: "90 %",
  horz: "Horizontal", vert: "Vertikal", ltHorz: "Horizontal hell", ltVert: "Vertikal hell", dkHorz: "Horizontal dunkel", dkVert: "Vertikal dunkel", narHorz: "Horizontal schmal",
  narVert: "Vertikal schmal", dashHorz: "Horizontal gestrichelt", dashVert: "Vertikal gestrichelt", cross: "Gitter", dnDiag: "Diagonal abwärts", upDiag: "Diagonal aufwärts",
  ltDnDiag: "Diagonal abwärts hell", ltUpDiag: "Diagonal aufwärts hell", dkDnDiag: "Diagonal abwärts dunkel", dkUpDiag: "Diagonal aufwärts dunkel", wdDnDiag: "Diagonal abwärts breit",
  wdUpDiag: "Diagonal aufwärts breit", dashDnDiag: "Diagonal abwärts gestrichelt", dashUpDiag: "Diagonal aufwärts gestrichelt", diagCross: "Diagonales Gitter", smCheck: "Kleines Schachbrett",
  lgCheck: "Großes Schachbrett", smGrid: "Kleines Gitter", lgGrid: "Großes Gitter", dotGrid: "Punktiertes Gitter", smConfetti: "Kleines Konfetti", lgConfetti: "Großes Konfetti",
  horzBrick: "Horizontale Ziegel", diagBrick: "Diagonale Ziegel", solidDmnd: "Volle Raute", openDmnd: "Offene Raute", dotDmnd: "Punktierte Raute", plaid: "Karo", sphere: "Kugel",
  weave: "Geflecht", divot: "Rasen", shingle: "Schindeln", wave: "Welle", trellis: "Spalier", zigZag: "Zickzack",
};
export const PATTERN_OPTIONS = Object.keys(PATTERNS).map((k) => ({ value: k, label: PATTERN_NAMES[k] ?? k, path: patternPath(k) }));

/** The chosen gradient stop per element (the pane's own state, not the chart's). */
const chosenStop = new Map<string, number>();

/** Excel's default gradient for a colour (light to darker, top to bottom). */
export function defaultGradient(color: string): Gradient {
  return { kind: "linear", angle: 90, stops: [
    { pos: 0, color: lum(color, 0.05, 0.95) }, { pos: 74, color: lum(color, 0.45, 0.55) }, { pos: 83, color: lum(color, 0.45, 0.55) }, { pos: 100, color: lum(color, 0.3, 0.7) },
  ] };
}

/** The controls of a gradient (Typ, Richtung, Winkel, Farbverlaufsstopps, Farbe, Position, Transparenz, Mit Form drehen). */
function gradientControls(key: string, g: Gradient, set: (g: Gradient) => void): PaneControl[] {
  const sel = Math.min(chosenStop.get(key) ?? 0, g.stops.length - 1);
  const stop = g.stops[sel];
  const withStop = (patch: Partial<Gradient["stops"][number]>) => set({ ...g, stops: g.stops.map((s, i) => (i === sel ? { ...s, ...patch } : s)) });
  const c: PaneControl[] = [
    { type: "select", label: "Typ", options: [{ label: "Linear", value: "linear" }, { label: "Radial", value: "radial" }, { label: "Rechteckig", value: "rect" }, { label: "Pfad", value: "path" }],
      value: () => g.kind, set: (v) => set({ ...g, kind: v as Gradient["kind"] }) },
  ];
  if (g.kind === "linear") {
    c.push({ type: "select", label: "Richtung", options: DIRECTIONS, value: () => String(g.angle ?? 0), set: (v) => set({ ...g, angle: Number(v) }) });
    c.push({ type: "number", label: "Winkel", unit: "°", min: 0, max: 359.9, step: 1, value: () => g.angle ?? 0, set: (v) => set({ ...g, angle: Math.round(v * 10) / 10 }) });
  }
  c.push({
    type: "gradient", label: "Farbverlaufsstopps", stops: () => g.stops, selected: () => sel,
    select: (i) => { chosenStop.set(key, i); },
    add: () => {
      // A new stop between the chosen one and the next, in the chosen one's colour.
      const next = g.stops[sel + 1] ?? g.stops[sel - 1];
      const pos = Math.round(((stop?.pos ?? 0) + (next?.pos ?? 100)) / 2);
      const stops = [...g.stops, { pos, color: stop?.color ?? "#FFFFFF", ...(stop?.transparency ? { transparency: stop.transparency } : {}) }];
      chosenStop.set(key, stops.length - 1);
      set({ ...g, stops });
    },
    remove: () => {
      if (g.stops.length <= 2) return;
      chosenStop.set(key, Math.max(0, sel - 1));
      set({ ...g, stops: g.stops.filter((_, i) => i !== sel) });
    },
    move: (i, pos) => set({ ...g, stops: g.stops.map((s, k) => (k === i ? { ...s, pos } : s)) }),
  });
  if (stop) {
    c.push({ type: "color", label: "Farbe", value: () => stop.color, set: (hex) => withStop({ color: hex ?? stop.color }) });
    c.push({ type: "number", label: "Position", unit: " %", min: 0, max: 100, slider: true, value: () => stop.pos, set: (v) => withStop({ pos: Math.round(v) }) });
    c.push({ type: "number", label: "Transparenz", unit: " %", min: 0, max: 100, slider: true, value: () => stop.transparency ?? 0, set: (v) => withStop({ transparency: Math.round(v) || undefined }) });
  }
  c.push({ type: "check", label: "Mit Form drehen", value: () => g.rotWithShape !== false, set: (on) => set({ ...g, rotWithShape: on ? undefined : false }) });
  return c;
}

/** Bild- oder Texturfüllung: Einfügen... (a picture file), Zwischenablage, Transparenz, nebeneinander als Textur. */
function pictureControls(f: Fill, set: (f: Fill) => void): PaneControl[] {
  const pic = f.picture;
  const choose = () => {
    const input = document.createElement("input");
    input.type = "file";
    input.accept = "image/png,image/jpeg,image/gif,image/bmp";
    input.addEventListener("change", () => {
      const file = input.files?.[0];
      if (!file) return;
      const reader = new FileReader();
      reader.onload = () => set({ type: "picture", picture: { ...pic, src: String(reader.result) } });
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
        reader.onload = () => set({ type: "picture", picture: { ...pic, src: String(reader.result) } });
        reader.readAsDataURL(blob);
        return;
      }
    } catch { /* no picture on the clipboard */ }
  };
  const c: PaneControl[] = [
    { type: "heading", label: "Bildquelle" },
    { type: "button", label: "Einfügen...", run: choose },
    { type: "button", label: "Zwischenablage", run: () => { void paste(); } },
    { type: "number", label: "Transparenz", unit: " %", min: 0, max: 100, slider: true, value: () => pic?.transparency ?? 0, set: (v) => pic && set({ type: "picture", picture: { ...pic, transparency: Math.round(v) || undefined } }), disabled: () => !pic },
    { type: "check", label: "Bild nebeneinander als Textur anordnen", value: () => !!pic?.tile, set: (on) => pic && set({ type: "picture", picture: { ...pic, tile: on || undefined } }), disabled: () => !pic },
  ];
  if (pic && !pic.tile) {
    const off = pic.offsets ?? { l: 0, t: 0, r: 0, b: 0 };
    const side = (label: string, k: keyof typeof off): PaneControl => ({ type: "number", label, unit: " %", min: -100, max: 100, value: () => off[k], set: (v) => set({ type: "picture", picture: { ...pic, offsets: { ...off, [k]: Math.round(v) } } }) });
    c.push(side("Offset links", "l"), side("Offset rechts", "r"), side("Offset oben", "t"), side("Offset unten", "b"));
  }
  return c;
}

/** The Füllung section. `extra` follows the type choice (Invertieren falls negativ, Punkte nach Farbe unterscheiden). */
export function fillSection(t: FmtTarget, autoColor: string, extra: PaneControl[] = [], open = true): PaneSection {
  return {
    title: "Füllung", open, controls: () => {
      const f = t.get()?.fill;
      const set = (fill: Fill | undefined) => t.change((x) => { if (fill) x.fill = fill; else delete x.fill; });
      const type = f?.type ?? "auto";
      const radio = (label: string, v: Fill["type"], make: () => Fill | undefined): PaneControl => ({ type: "radio", label, checked: () => type === v, pick: () => set(make()) });
      const color = f?.color ?? autoColor;
      const c: PaneControl[] = [
        radio("Keine Füllung", "none", () => ({ type: "none" })),
        radio("Einfarbige Füllung", "solid", () => ({ type: "solid", color })),
        radio("Farbverlauf", "gradient", () => ({ type: "gradient", gradient: f?.gradient ?? defaultGradient(color) })),
        radio("Bild- oder Texturfüllung", "picture", () => ({ type: "picture", picture: f?.picture })),
        radio("Musterfüllung", "pattern", () => ({ type: "pattern", pattern: f?.pattern ?? { preset: "pct5", fg: color, bg: "#FFFFFF" } })),
        radio("Automatisch", "auto", () => undefined),
        ...extra,
      ];
      if (f?.type === "solid") {
        c.push({ type: "color", label: "Farbe", value: () => f.color ?? null, set: (hex) => set({ ...f, color: hex ?? autoColor }) });
        c.push({ type: "number", label: "Transparenz", unit: " %", min: 0, max: 100, slider: true, value: () => f.transparency ?? 0, set: (v) => set({ ...f, transparency: Math.round(v) || undefined }) });
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
  /** Pfeile (Anfangs- / Endpfeiltyp): lines that have ends. */
  arrows?: boolean;
  /** Controls after the type choice or at the end (Abgerundete Ecken, Linie glätten). */
  extra?: PaneControl[];
  open?: boolean;
}

/** The Rahmen / Linie section. */
export function lineSection(t: FmtTarget, o: LineOptions): PaneSection {
  return {
    title: o.title, open: o.open, controls: () => {
      const l = t.get()?.line;
      const set = (line: Line | undefined) => t.change((x) => { if (line) x.line = line; else delete x.line; });
      const patch = (p: Partial<Line>) => {
        const next: Line = { ...(l ?? { type: "auto" }), ...p };
        for (const k of Object.keys(next) as (keyof Line)[]) if (next[k] === undefined) delete next[k];
        set(next);
      };
      const type = l?.type ?? "auto";
      const color = l?.color ?? o.autoColor;
      const c: PaneControl[] = [
        { type: "radio", label: "Keine Linie", checked: () => type === "none", pick: () => set({ type: "none" }) },
        { type: "radio", label: "Einfarbige Linie", checked: () => type === "solid", pick: () => set({ ...l, type: "solid", color, width: l?.width ?? o.autoWidth }) },
        { type: "radio", label: "Farbverlaufslinie", checked: () => type === "gradient", pick: () => set({ ...l, type: "gradient", gradient: l?.gradient ?? defaultGradient(color), width: l?.width ?? o.autoWidth }) },
        { type: "radio", label: "Automatisch", checked: () => type === "auto", pick: () => set(undefined) },
      ];
      if (type === "solid") {
        c.push({ type: "color", label: "Farbe", value: () => l?.color ?? null, set: (hex) => patch({ color: hex ?? o.autoColor }) });
        c.push({ type: "number", label: "Transparenz", unit: " %", min: 0, max: 100, slider: true, value: () => l?.transparency ?? 0, set: (v) => patch({ transparency: Math.round(v) || undefined }) });
      }
      if (type === "gradient" && l?.gradient) c.push(...gradientControls(`${t.key}/line`, l.gradient, (g) => patch({ gradient: g })));
      if (type !== "none") {
        c.push({ type: "number", label: "Breite", unit: " Pt.", min: 0, max: 1584, step: 0.25, decimals: 2, value: () => l?.width ?? o.autoWidth, set: (v) => patch({ width: v }) });
        c.push({ type: "select", label: "Verbundtyp", options: COMPOUND, value: () => l?.compound ?? "sng", set: (v) => patch({ compound: v === "sng" ? undefined : (v as Line["compound"]) }) });
        c.push({ type: "select", label: "Strichtyp", options: DASHES, value: () => l?.dash ?? "solid", set: (v) => patch({ dash: v === "solid" ? undefined : (v as Dash) }) });
        c.push({ type: "select", label: "Abschlusstyp", options: CAPS, value: () => l?.cap ?? "flat", set: (v) => patch({ cap: v === "flat" ? undefined : (v as Line["cap"]) }) });
        c.push({ type: "select", label: "Anschlusstyp", options: JOINS, value: () => l?.join ?? "round", set: (v) => patch({ join: v === "round" ? undefined : (v as Line["join"]) }) });
        if (o.arrows) {
          const arrow = (which: "head" | "tail", label: string, sizeLabel: string): PaneControl[] => {
            const a = l?.[which];
            return [
              { type: "select", label, options: ARROWS, value: () => a?.type ?? "none", set: (v) => patch({ [which]: v === "none" ? undefined : { ...a, type: v as Arrow["type"] } }) },
              { type: "select", label: sizeLabel, options: ARROW_SIZES, value: () => `${a?.w ?? "med"},${a?.len ?? "med"}`, disabled: () => !a,
                set: (v) => { const [w, len] = v.split(",") as [Arrow["w"], Arrow["len"]]; if (a) patch({ [which]: { ...a, w, len } }); } },
            ];
          };
          c.push(...arrow("head", "Anfangspfeiltyp", "Anfangspfeilgröße"), ...arrow("tail", "Endpfeiltyp", "Endpfeilgröße"));
        }
      }
      c.push(...(o.extra ?? []));
      return c;
    },
  };
}

const SHADOWS: { label: string; value: string; s?: Effects["shadow"] }[] = [
  { label: "Kein Schatten", value: "none" },
  { label: "Außen: Versatz: unten rechts", value: "o-br", s: { color: "#000000", transparency: 60, size: 100, blur: 4, angle: 45, distance: 3 } },
  { label: "Außen: Versatz: unten", value: "o-b", s: { color: "#000000", transparency: 60, size: 100, blur: 4, angle: 90, distance: 3 } },
  { label: "Außen: Versatz: unten links", value: "o-bl", s: { color: "#000000", transparency: 60, size: 100, blur: 4, angle: 135, distance: 3 } },
  { label: "Außen: Versatz: rechts", value: "o-r", s: { color: "#000000", transparency: 60, size: 100, blur: 4, angle: 0, distance: 3 } },
  { label: "Außen: Versatz: Mitte", value: "o-c", s: { color: "#000000", transparency: 60, size: 102, blur: 5, angle: 0, distance: 0 } },
  { label: "Außen: Versatz: oben rechts", value: "o-tr", s: { color: "#000000", transparency: 60, size: 100, blur: 4, angle: 315, distance: 3 } },
  { label: "Innen: oben links", value: "i-tl", s: { inner: true, color: "#000000", transparency: 50, blur: 5, angle: 225, distance: 4 } },
  { label: "Innen: Mitte", value: "i-c", s: { inner: true, color: "#000000", transparency: 0, blur: 9, angle: 0, distance: 0 } },
  { label: "Innen: unten rechts", value: "i-br", s: { inner: true, color: "#000000", transparency: 50, blur: 5, angle: 45, distance: 4 } },
];

const BEVELS = [
  { label: "Ohne", value: "none" }, { label: "Kreis", value: "circle" }, { label: "Relief", value: "relaxedInset" }, { label: "Kreuz", value: "cross" },
  { label: "Kühles Schräglicht", value: "coolSlant" }, { label: "Winkel", value: "angle" }, { label: "Weiche Rundung", value: "softRound" },
  { label: "Konvex", value: "convex" }, { label: "Schräg", value: "slope" }, { label: "Absenken", value: "divot" }, { label: "Riblet", value: "riblet" },
  { label: "Harte Kante", value: "hardEdge" }, { label: "Art Deco", value: "artDeco" },
];
const MATERIALS = [
  { label: "Standard: Matt", value: "matte" }, { label: "Standard: Warm matt", value: "warmMatte" }, { label: "Standard: Plastik", value: "plastic" },
  { label: "Standard: Metall", value: "metal" }, { label: "Spezialeffekt: Dunkle Kante", value: "dkEdge" }, { label: "Spezialeffekt: Weiche Kante", value: "softEdge" },
  { label: "Spezialeffekt: Flach", value: "flat" }, { label: "Spezialeffekt: Drahtmodell", value: "legacyWireframe" }, { label: "Durchsichtig: Puder", value: "powder" },
  { label: "Durchsichtig: Durchsichtiges Pulver", value: "translucentPowder" }, { label: "Durchsichtig: Klar", value: "clear" },
];
const LIGHTS = [
  { label: "Neutral: Dreipunkt", value: "threePt" }, { label: "Neutral: Ausgewogen", value: "balanced" }, { label: "Neutral: Weich", value: "soft" },
  { label: "Neutral: Hart", value: "harsh" }, { label: "Neutral: Flutlicht", value: "flood" }, { label: "Neutral: Kontrastreich", value: "contrasting" },
  { label: "Warm: Morgen", value: "morning" }, { label: "Warm: Sonnenaufgang", value: "sunrise" }, { label: "Warm: Sonnenuntergang", value: "sunset" },
  { label: "Kalt: Kühl", value: "chilly" }, { label: "Kalt: Frost", value: "freezing" }, { label: "Speziell: Flach", value: "flat" }, { label: "Speziell: Zwei Punkte", value: "twoPt" },
  { label: "Speziell: Leuchten", value: "glow" }, { label: "Speziell: Hell", value: "brightRoom" },
];

export interface EffectOptions { shadow?: boolean; glow?: boolean; soft?: boolean; format3d?: boolean; autoColor: string }

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
    const preset = !s ? "none" : SHADOWS.find((p) => p.s && JSON.stringify(p.s) === JSON.stringify(s))?.value ?? "custom";
    const patch = (p: Partial<NonNullable<Effects["shadow"]>>) => set({ shadow: { ...(s ?? SHADOWS[1].s!), ...p } });
    const c: PaneControl[] = [
      { type: "select", label: "Voreinstellungen", options: [...SHADOWS.map(({ label, value }) => ({ label, value })), ...(preset === "custom" ? [{ label: "Benutzerdefiniert", value: "custom" }] : [])],
        value: () => preset, set: (v) => set({ shadow: v === "none" ? undefined : { ...SHADOWS.find((p) => p.value === v)!.s! } }) },
      { type: "color", label: "Farbe", value: () => s?.color ?? "#000000", set: (hex) => patch({ color: hex ?? "#000000" }) },
      { type: "number", label: "Transparenz", unit: " %", min: 0, max: 100, slider: true, value: () => s?.transparency ?? 60, set: (v) => patch({ transparency: Math.round(v) }) },
      { type: "number", label: "Größe", unit: " %", min: 1, max: 200, slider: true, value: () => s?.size ?? 100, set: (v) => patch({ size: Math.round(v) }), disabled: () => !!s?.inner },
      { type: "number", label: "Weichzeichnen", unit: " Pt.", min: 0, max: 100, slider: true, value: () => s?.blur ?? 4, set: (v) => patch({ blur: v }) },
      { type: "number", label: "Winkel", unit: "°", min: 0, max: 359.9, slider: true, value: () => s?.angle ?? 45, set: (v) => patch({ angle: v }) },
      { type: "number", label: "Abstand", unit: " Pt.", min: 0, max: 200, slider: true, value: () => s?.distance ?? 3, set: (v) => patch({ distance: v }) },
    ];
    return c;
  } });
  if (o.glow !== false) out.push({ title: "Leuchten", controls: () => {
    const g = fx()?.glow;
    const presets = [0, 5, 8, 11, 18];
    return [
      { type: "select", label: "Voreinstellungen", options: presets.map((p) => ({ label: p ? `Leuchten: ${p} Pt.` : "Kein Leuchten", value: String(p) })),
        value: () => String(g?.size ?? 0), set: (v) => set({ glow: Number(v) ? { color: g?.color ?? o.autoColor, size: Number(v), transparency: g?.transparency ?? 60 } : undefined }) },
      { type: "color", label: "Farbe", value: () => g?.color ?? o.autoColor, set: (hex) => set({ glow: { size: g?.size ?? 5, transparency: g?.transparency ?? 60, color: hex ?? o.autoColor } }) },
      { type: "number", label: "Größe", unit: " Pt.", min: 0, max: 150, slider: true, value: () => g?.size ?? 0, set: (v) => set({ glow: v ? { color: g?.color ?? o.autoColor, transparency: g?.transparency ?? 60, size: v } : undefined }) },
      { type: "number", label: "Transparenz", unit: " %", min: 0, max: 100, slider: true, value: () => g?.transparency ?? 60, set: (v) => set({ glow: { color: g?.color ?? o.autoColor, size: g?.size ?? 5, transparency: Math.round(v) } }) },
    ];
  } });
  if (o.soft !== false) out.push({ title: "Weiche Kanten", controls: () => {
    const v = fx()?.softEdges ?? 0;
    const presets = [0, 1, 2.5, 5, 10, 25, 50];
    return [
      { type: "select", label: "Voreinstellungen", options: presets.map((p) => ({ label: p ? `${String(p).replace(".", ",")} Pt.` : "Keine weichen Kanten", value: String(p) })),
        value: () => String(presets.includes(v) ? v : 0), set: (x) => set({ softEdges: Number(x) || undefined }) },
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
        { type: "select", label, options: BEVELS, value: () => b?.type ?? "none", set: (v) => set3({ [which]: v === "none" ? undefined : { w: b?.w ?? 6, h: b?.h ?? 6, type: v } }) },
        { type: "number", label: "Breite", unit: " Pt.", min: 0, max: 1584, value: () => b?.w ?? 0, set: (v) => b && set3({ [which]: { ...b, w: v } }), disabled: () => !b },
        { type: "number", label: "Höhe", unit: " Pt.", min: 0, max: 1584, value: () => b?.h ?? 0, set: (v) => b && set3({ [which]: { ...b, h: v } }), disabled: () => !b },
      ];
    };
    return [
      ...bevel("top", "Abschrägung oben"), ...bevel("bottom", "Abschrägung unten"),
      { type: "color", label: "Farbe der Tiefe", value: () => f?.depth?.color ?? null, set: (hex) => set3({ depth: { size: f?.depth?.size ?? 0, ...(hex ? { color: hex } : {}) } }) },
      { type: "number", label: "Tiefe", unit: " Pt.", min: 0, max: 1584, value: () => f?.depth?.size ?? 0, set: (v) => set3({ depth: v ? { ...f?.depth, size: v } : undefined }) },
      { type: "color", label: "Konturfarbe", value: () => f?.contour?.color ?? null, set: (hex) => set3({ contour: { size: f?.contour?.size ?? 0, ...(hex ? { color: hex } : {}) } }) },
      { type: "number", label: "Kontur", unit: " Pt.", min: 0, max: 1584, value: () => f?.contour?.size ?? 0, set: (v) => set3({ contour: v ? { ...f?.contour, size: v } : undefined }) },
      { type: "select", label: "Material", options: MATERIALS, value: () => f?.material ?? "warmMatte", set: (v) => set3({ material: v === "warmMatte" ? undefined : v }) },
      { type: "select", label: "Beleuchtung", options: LIGHTS, value: () => f?.lighting ?? "threePt", set: (v) => set3({ lighting: v === "threePt" ? undefined : v }) },
      { type: "button", label: "Zurücksetzen", run: () => set({ format3d: undefined }) },
    ];
  } });
  return out;
}
