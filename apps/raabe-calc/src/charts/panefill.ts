// The Format pane's "Füllung und Linie" and "Effekte" tabs, shared by every chart element (D-066,
// D-111): Füllung (Keine, Einfarbig, Farbverlauf, Bild- oder Texturfüllung, Muster, Automatisch),
// Rahmen / Linie (Keine, Einfarbig, Farbverlaufslinie, Automatisch with Breite, Verbundtyp,
// Strichtyp, Abschlusstyp, Anschlusstyp, Pfeile), Schatten, Leuchten, Weiche Kanten and 3D-Format.
// Every control changes the element at once as one undo step.

import { OFFICE_ACCENTS, PATTERNS, lum, patternPath, type Arrow, type Dash, type Effects, type Fill, type Format3D, type Gradient, type Line } from "@raabe/charts";
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
  const withStop = (patch: Partial<Gradient["stops"][number]>) => {
    const changedStop = { ...g.stops[sel], ...patch };
    const stops = g.stops.map((s, i) => (i === sel ? changedStop : s)).sort((a, b) => a.pos - b.pos);
    chosenStop.set(key, stops.indexOf(changedStop));
    set({ ...g, stops });
  };
  const c: PaneControl[] = [
    { type: "select", label: "Typ", options: [{ label: "Linear", value: "linear" }, { label: "Radial", value: "radial" }, { label: "Rechteckig", value: "rect" }, { label: "Pfad", value: "path" }],
      value: () => g.kind, set: (v) => set({ ...g, kind: v as Gradient["kind"], angle: v === "linear" ? g.angle ?? 90 : undefined }) },
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
/** `withAuto` false: no Automatisch (Textfüllung), Einfarbige Füllung in the automatic colour is the default. */
export function fillSection(t: FmtTarget, autoColor: string, extra: PaneControl[] = [], open = true, withAuto = true): PaneSection {
  return {
    title: "Füllung", open, controls: () => {
      const f = t.get()?.fill;
      const set = (fill: Fill | undefined) => t.change((x) => { if (fill) x.fill = fill; else delete x.fill; });
      const type = f?.type ?? (withAuto ? "auto" : "solid");
      const radio = (label: string, v: Fill["type"], make: () => Fill | undefined): PaneControl => ({ type: "radio", label, checked: () => type === v, pick: () => set(make()) });
      const color = f?.color ?? autoColor;
      const c: PaneControl[] = [
        radio("Keine Füllung", "none", () => ({ type: "none" })),
        radio("Einfarbige Füllung", "solid", () => ({ type: "solid", color })),
        radio("Farbverlauf", "gradient", () => ({ type: "gradient", gradient: f?.gradient ?? defaultGradient(color) })),
        radio("Bild- oder Texturfüllung", "picture", () => ({ type: "picture", picture: f?.picture })),
        radio("Musterfüllung", "pattern", () => ({ type: "pattern", pattern: f?.pattern ?? { preset: "pct5", fg: color, bg: "#FFFFFF" } })),
        ...(withAuto ? [radio("Automatisch", "auto", () => undefined)] : []),
        ...extra,
      ];
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
  /** Pfeile (Anfangs- / Endpfeiltyp): lines that have ends. */
  arrows?: boolean;
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
        const next: Line = { ...(l ?? { type: "auto" }), ...p };
        for (const k of Object.keys(next) as (keyof Line)[]) if (next[k] === undefined) delete next[k];
        set(next);
      };
      const type = l?.type ?? (o.withAuto === false ? "none" : "auto");
      const color = l?.color ?? o.autoColor;
      const c: PaneControl[] = [
        { type: "radio", label: "Keine Linie", checked: () => type === "none", pick: () => set(o.withAuto === false ? undefined : { type: "none" }) },
        { type: "radio", label: "Einfarbige Linie", checked: () => type === "solid", pick: () => set({ ...l, type: "solid", color, width: l?.width ?? o.autoWidth }) },
        { type: "radio", label: "Farbverlaufslinie", checked: () => type === "gradient", pick: () => set({ ...l, type: "gradient", gradient: l?.gradient ?? defaultGradient(color), width: l?.width ?? o.autoWidth }) },
        ...(o.withAuto === false ? [] : [{ type: "radio", label: "Automatisch", checked: () => type === "auto", pick: () => set(undefined) } as PaneControl]),
      ];
      if (type === "solid") {
        c.push({ type: "color", label: "Farbe", value: () => l?.color ?? null, set: (hex) => patch({ color: hex ?? o.autoColor }) });
        c.push({ type: "number", label: "Transparenz", unit: " %", min: 0, max: 100, slider: true, value: () => l?.transparency ?? 0, set: (v) => patch({ transparency: Math.round(v) || undefined }) });
      }
      if (type === "gradient" && l?.gradient) c.push(...gradientControls(`${t.key}/line`, l.gradient, (g) => patch({ gradient: g })));
      if (type !== "none") {
        c.push({ type: "number", label: "Stärke", unit: " Pt.", min: 0, max: 1584, step: 0.25, decimals: 2, value: () => l?.width ?? o.autoWidth, set: (v) => patch({ width: v }) });
        c.push({ type: "select", label: "Verbundtyp", options: COMPOUND, value: () => l?.compound ?? "sng", set: (v) => patch({ compound: v === "sng" ? undefined : (v as Line["compound"]) }) });
        c.push({ type: "select", label: "Strichtyp", options: DASHES, value: () => l?.dash ?? "solid", set: (v) => patch({ dash: v === "solid" ? undefined : (v as Dash) }) });
        c.push({ type: "select", label: "Abschlusstyp", options: CAPS, value: () => l?.cap ?? "flat", set: (v) => patch({ cap: v === "flat" ? undefined : (v as Line["cap"]) }) });
        c.push({ type: "select", label: "Anschlusstyp", options: JOINS, value: () => l?.join ?? "round", set: (v) => patch({ join: v === "round" ? undefined : (v as Line["join"]) }) });
        if (o.arrows) {
          const arrow = (which: "head" | "tail", label: string, sizeLabel: string): PaneControl[] => {
            const a = l?.[which];
            return [
              { type: "select", label, options: ARROWS, value: () => a?.type ?? "none", set: (v) => patch({ [which]: v === "none" ? undefined : { ...a, type: v as Arrow["type"] } }) },
              { type: "select", label: sizeLabel, options: arrowSizes(which === "head" ? "L" : "R"), value: () => `${a?.w ?? "med"},${a?.len ?? "med"}`, disabled: () => !a,
                set: (v) => { const [w, len] = v.split(",") as [Arrow["w"], Arrow["len"]]; if (a) patch({ [which]: { ...a, w, len } }); } },
            ];
          };
          c.push(...arrow("head", "Startpfeiltyp", "Startpfeilgröße"), ...arrow("tail", "Endpfeiltyp", "Endpfeilgröße"));
        }
      }
      c.push(...(o.extra ?? []));
      return c;
    },
  };
}

const SHADOWS: { label: string; value: string; s?: Effects["shadow"] }[] = [
  { label: "Ohne", value: "none" },
  ...([["unten rechts", 45], ["unten", 90], ["unten links", 135], ["rechts", 0], ["Mitte", -1], ["links", 180], ["oben rechts", 315], ["oben", 270], ["oben links", 225]] as const).map(([n, a]) =>
    ({ label: `Offset: ${n}`, value: `o-${n}`, s: a < 0 ? { color: "#000000", transparency: 60, size: 102, blur: 5, angle: 0, distance: 0 } : { color: "#000000", transparency: 60, blur: 4, angle: a, distance: 3 } })),
  ...([["oben links", 225], ["oben", 270], ["oben rechts", 315], ["links", 180], ["Mitte", -1], ["rechts", 0], ["unten links", 135], ["unten", 90], ["unten rechts", 45]] as const).map(([n, a]) =>
    ({ label: `Innen: ${n}`, value: `i-${n}`, s: a < 0 ? { inner: true, color: "#000000", transparency: 0, blur: 9, angle: 0, distance: 0 } : { inner: true, color: "#000000", transparency: 50, blur: 5, angle: a, distance: 4 } })),
];

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
    const preset = !s ? "none" : SHADOWS.find((p) => p.s && JSON.stringify(p.s) === JSON.stringify(s))?.value ?? "custom";
    const patch = (p: Partial<NonNullable<Effects["shadow"]>>) => set({ shadow: { ...(s ?? SHADOWS[1].s!), ...p } });
    const c: PaneControl[] = [
      { type: "select", label: "Voreinstellungen", options: [...SHADOWS.map(({ label, value }) => ({ label, value })), ...(preset === "custom" ? [{ label: "Benutzerdefiniert", value: "custom" }] : [])],
        value: () => preset, set: (v) => set({ shadow: v === "none" ? undefined : { ...SHADOWS.find((p) => p.value === v)!.s! } }) },
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
      { type: "select", label: "Voreinstellungen", options: [{ label: "Ohne", value: "none" }, ...presets, ...(r && !presets.some((p) => p.value === current) ? [{ label: "Benutzerdefiniert", value: current }] : [])],
        value: () => current, set: (v) => { if (v === "none") set({ reflection: undefined }); else { const [size, distance] = v.split(",").map(Number); set({ reflection: { transparency: 50, size, blur: 0.5, distance } }); } } },
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
      { type: "select", label: "Voreinstellungen", options: [{ label: "Ohne", value: "none" }, ...presets, ...(g && !presets.some((p) => p.value === current) ? [{ label: "Benutzerdefiniert", value: current }] : [])],
        value: () => current, set: (v) => { if (v === "none") set({ glow: undefined }); else { const [size, color] = v.split(","); set({ glow: { color, size: Number(size), transparency: 60 } }); } } },
      { type: "color", label: "Farbe", value: () => g?.color ?? o.autoColor, set: (hex) => set({ glow: { size: g?.size ?? 5, transparency: g?.transparency ?? 60, color: hex ?? o.autoColor } }) },
      { type: "number", label: "Größe", unit: " Pt.", min: 0, max: 150, slider: true, value: () => g?.size ?? 0, set: (v) => set({ glow: v ? { color: g?.color ?? o.autoColor, transparency: g?.transparency ?? 60, size: v } : undefined }) },
      { type: "number", label: "Transparenz", unit: " %", min: 0, max: 100, slider: true, value: () => g?.transparency ?? 60, set: (v) => set({ glow: { color: g?.color ?? o.autoColor, size: g?.size ?? 5, transparency: Math.round(v) } }) },
    ];
  } });
  if (o.soft !== false) out.push({ title: "Weiche Kanten", controls: () => {
    const v = fx()?.softEdges ?? 0;
    const presets = [0, 1, 2.5, 5, 10, 25, 50];
    return [
      { type: "select", label: "Voreinstellungen", options: presets.map((p) => ({ label: p ? `${String(p).replace(".", ",")} Punkt` : "Keine", value: String(p) })),
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
      { type: "color", label: "Tiefe", value: () => f?.depth?.color ?? null, set: (hex) => set3({ depth: { size: f?.depth?.size ?? 0, ...(hex ? { color: hex } : {}) } }) },
      { type: "number", label: "Größe", unit: " Pt.", min: 0, max: 1584, value: () => f?.depth?.size ?? 0, set: (v) => set3({ depth: v ? { ...f?.depth, size: v } : undefined }) },
      { type: "color", label: "Kontur", value: () => f?.contour?.color ?? null, set: (hex) => set3({ contour: { size: f?.contour?.size ?? 0, ...(hex ? { color: hex } : {}) } }) },
      { type: "number", label: "Größe", unit: " Pt.", min: 0, max: 1584, value: () => f?.contour?.size ?? 0, set: (v) => set3({ contour: v ? { ...f?.contour, size: v } : undefined }) },
      { type: "select", label: "Material", options: MATERIALS, value: () => f?.material ?? "warmMatte", set: (v) => set3({ material: v === "warmMatte" ? undefined : v }) },
      { type: "select", label: "Beleuchtung", options: LIGHTS, value: () => f?.lighting ?? "threePt", set: (v) => set3({ lighting: v === "threePt" ? undefined : v }) },
      { type: "button", label: "Zurücksetzen", run: () => set({ format3d: undefined }) },
    ];
  } });
  return out;
}
