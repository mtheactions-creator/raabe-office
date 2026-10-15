// The task panes Grafik formatieren and Form formatieren (Step 4c, D-054), laid out as Word 365
// has them (recorded with tools/feel-lab/probe_word_format_pane.py): docked at the right, a row
// of tabs (Füllung und Linie, Effekte, Layout und Eigenschaften and, for pictures, Bild; shapes
// have them under Formoptionen, with Textoptionen beside) and fold-out sections with a slider
// and a value field per setting. Changes apply at once; dragging a slider is one undo step. The
// "…optionen" entries of the ribbon's menus open their section here, as in Word.

import { closePopup, colorGrid, h, icon, showMenu, showPopup } from "@raabe/shell";
import type { PictureFormat, PictureShadow, ShapeInfo } from "@raabe/text-engine";
import { DEFAULT_THEME_COLORS, pictureFromSpPr } from "@raabe/ooxml";
import { CM, parseNum, type AppApi } from "./api";
import type { Dock } from "./dock";
import {
  BRIGHT_CONTRAST, SATURATION, SHARPEN, TEMPERATURE, TRANSPARENCY, galleryPanel, recolorGallery, type GalleryItem, type GallerySection,
} from "./picture-format";
import * as P from "./picture-presets";

/** The sections an "…optionen" entry can open. */
export type FormatSection =
  | "fill" | "line" | "shadow" | "reflection" | "glow" | "softEdge" | "format3d" | "rotation3d" | "artistic"
  | "textbox" | "corrections" | "color" | "transparency" | "crop" | "textFill" | "textEffects" | "textLayout";

type PaneTab = "fillLine" | "effects" | "layout" | "picture" | "textFill" | "textEffects" | "textLayout";

const TAB_OF: Record<FormatSection, PaneTab> = {
  fill: "fillLine", line: "fillLine", shadow: "effects", reflection: "effects", glow: "effects", softEdge: "effects",
  format3d: "effects", rotation3d: "effects", artistic: "effects", textbox: "layout", corrections: "picture", color: "picture",
  transparency: "picture", crop: "picture", textFill: "textFill", textEffects: "textEffects", textLayout: "textLayout",
};

/** Opens the pane (at a section); the feature fills it in. */
export const FORMAT_PANE = { open: (_section?: FormatSection) => {} };

const DASHES: [string, string][] = [["solid", "Durchgezogen"], ["sysDot", "Rundpunkt"], ["sysDash", "Quadratpunkt"], ["dash", "Strich"], ["dashDot", "Strich-Punkt"],
  ["lgDash", "Langer Strich"], ["lgDashDot", "Lang-Strich-Punkt"], ["lgDashDotDot", "Lang-Strich-Punkt-Punkt"]];
const COMPOUNDS: [string, string][] = [["sng", "Einfach"], ["dbl", "Doppelt"], ["thickThin", "Dick-Dünn"], ["thinThick", "Dünn-Dick"], ["tri", "Dreifach"]];
const CAPS: [string, string][] = [["sq", "Quadratisch"], ["rnd", "Rund"], ["flat", "Flach"]];
const JOINS: [string, string][] = [["round", "Rund"], ["bevel", "Abgeschrägt"], ["miter", "Winkel"]];
const ARROWS: [string, string][] = [["none", "Kein Pfeil"], ["triangle", "Pfeil"], ["arrow", "Offener Pfeil"], ["stealth", "Spitzer Pfeil"],
  ["diamond", "Rautenförmiger Pfeil"], ["oval", "Ovaler Pfeil"]];
const ARROW_SIZES: [string, string][] = ["sm", "med", "lg"].flatMap((w) => ["sm", "med", "lg"].map((l): [string, string] =>
  [`${w}/${l}`, `${{ sm: "Schmal", med: "Mittel", lg: "Breit" }[w]}, ${{ sm: "kurz", med: "mittel", lg: "lang" }[l]}`]));

const DEFAULT_SHADOW: PictureShadow = { kind: "outer", color: "000000", transparency: 0.6, blur: 4, dist: 3, dir: 45, sx: 1, sy: 1 };
const DEFAULT_REFLECTION = { size: 0.5, dist: 0, transparency: 0.5, blur: 0.5 };
const DEFAULT_GLOW = { color: "4472C4", radius: 5, transparency: 0.6 };

/** The 3D-Drehung of a scene (degrees): X turns about the upright axis (lon), Y about the level one (lat). */
const rotOf = (xml?: string) => {
  const m = /<a:rot lat="(\d+)" lon="(\d+)" rev="(\d+)"\/>/.exec(xml ?? "");
  return m ? { x: Number(m[2]) / 60000, y: Number(m[1]) / 60000, z: Number(m[3]) / 60000 } : { x: 0, y: 0, z: 0 };
};
const withRot = (xml: string | undefined, r: { x: number; y: number; z: number }) => {
  const deg = (v: number) => Math.round((((v % 360) + 360) % 360) * 60000);
  const rot = `<a:rot lat="${deg(r.y)}" lon="${deg(r.x)}" rev="${deg(r.z)}"/>`;
  if (!xml) return `<a:scene3d><a:camera prst="orthographicFront">${rot}</a:camera><a:lightRig rig="threePt" dir="t"/></a:scene3d>`;
  if (/<a:rot [^>]*\/>/.test(xml)) return xml.replace(/<a:rot [^>]*\/>/, rot);
  return xml.replace(/<a:camera ([^>]*?)\/>/, `<a:camera $1>${rot}</a:camera>`);
};

interface Target {
  readonly kind: "picture" | "shape";
  readonly f: PictureFormat;
  readonly shape?: ShapeInfo;
  readonly w: number;
  readonly h: number;
  readonly float: boolean;
}

/** A number setting: a slider (unless `slider` is false) and a value field with its unit. */
interface NumField {
  label: string;
  get: () => number | null;
  /** `start`: the first change of a slider drag or a typed value (a new undo step). */
  put: (v: number, start: boolean) => void;
  min: number;
  max: number;
  step?: number;
  unit: string;
  slider?: boolean;
  disabled?: boolean;
  decimals?: number;
}

export function formatPaneFeature(api: AppApi, docArea: HTMLElement, dock: Dock, text: { body(t: "fill" | "effects"): HTMLElement }) {
  const c = api.commands;
  const ed = () => api.editor();
  const theme = () => ({ ...DEFAULT_THEME_COLORS, ...(ed()?.doc.styles.theme.colors ?? {}) });
  const pane = h("div", { class: "fxpane fmtpane", hidden: true, role: "complementary" });
  docArea.append(pane);
  dock.add(pane, "right", 320);
  let isOpen = false;
  let pictureTab: PaneTab = "effects";
  let shapeTab: PaneTab = "fillLine";
  const opened = new Set<string>();
  let reveal: string | null = null;
  let builtKey = "";
  let syncs: (() => void)[] = [];

  const target = (): Target | null => {
    const e = ed();
    if (!e) return null;
    const s = e.selectedShape();
    if (s?.obj.shape) return { kind: "shape", f: s.obj.pic ?? {}, shape: s.obj.shape, w: s.obj.width ?? 0, h: s.obj.height ?? 0, float: true };
    const img = e.selectedImage();
    if (img?.obj.image) return { kind: "picture", f: img.obj.pic ?? {}, w: img.obj.width ?? 0, h: img.obj.height ?? 0, float: img.kind === "float" };
    return null;
  };
  const set = (change: PictureFormat, start = true) => {
    const e = ed();
    const t = target();
    if (!e || !t) return;
    if (t.kind === "shape") e.setShapeFormat(change, start);
    else e.setPictureFormat(change, start);
    api.refresh();
  };
  const setBody = (patch: Partial<ShapeInfo["body"]>) => {
    const e = ed();
    if (!e?.selectedShape()) return;
    e.setShape((s) => ({ ...s, body: { ...s.body, ...patch } }));
    api.refresh();
  };

  // ----- fields -------------------------------------------------------------------------------
  const fmtNum = (v: number, n: NumField) => {
    const s = v.toLocaleString("de-DE", { maximumFractionDigits: n.decimals ?? 1 });
    return n.unit === "°" ? `${s}°` : n.unit ? `${s} ${n.unit}` : s;
  };
  const spinArrow = (up: boolean) => h("span", { class: `fp-arrow ${up ? "up" : "down"}` });
  const num = (n: NumField): HTMLElement => {
    const box = h("input", { class: "fp-val", "aria-label": n.label, disabled: !!n.disabled, spellcheck: false }) as HTMLInputElement;
    const range = n.slider === false ? null
      : h("input", { type: "range", min: String(n.min), max: String(n.max), step: String(n.step ?? 1), "aria-label": n.label, disabled: !!n.disabled, tabIndex: -1 }) as HTMLInputElement;
    const clamp = (v: number) => Math.min(n.max, Math.max(n.min, v));
    let dragging = false;
    const show = () => {
      const v = n.get();
      if (document.activeElement !== box) box.value = v === null ? "" : fmtNum(v, n);
      if (range && !dragging) range.value = String(v ?? n.min);
    };
    const commit = () => { const v = parseNum(box.value); if (Number.isFinite(v)) n.put(clamp(v), true); show(); };
    const step = (d: number) => {
      const cur = parseNum(box.value);
      const v = clamp(Math.round(((Number.isFinite(cur) ? cur : n.get() ?? 0) + d * (n.step ?? 1)) * 100) / 100);
      n.put(v, true);
      box.value = fmtNum(v, n);
    };
    box.addEventListener("change", commit);
    box.addEventListener("keydown", (ev) => {
      if (ev.key === "ArrowUp" || ev.key === "ArrowDown") { ev.preventDefault(); step(ev.key === "ArrowUp" ? 1 : -1); }
      else if (ev.key === "Enter") { ev.preventDefault(); commit(); box.select(); }
      else if (ev.key === "Escape") { ev.preventDefault(); box.blur(); show(); api.focus(); }
    });
    box.addEventListener("focus", () => box.select());
    box.addEventListener("blur", show);
    if (range) {
      range.addEventListener("input", () => { const v = Number(range.value); box.value = fmtNum(v, n); n.put(v, !dragging); dragging = true; });
      range.addEventListener("change", () => { dragging = false; });
    }
    const spin = h("span", { class: "fp-spin" },
      h("button", { "aria-label": "Mehr", tabIndex: -1, disabled: !!n.disabled, onclick: () => step(1) }, spinArrow(true)),
      h("button", { "aria-label": "Weniger", tabIndex: -1, disabled: !!n.disabled, onclick: () => step(-1) }, spinArrow(false)));
    syncs.push(show);
    show();
    return h("div", { class: `fp-row${n.disabled ? " off" : ""}` }, h("span", { class: "fp-label" }, n.label), range ?? h("span"), h("span", { class: "fp-box" }, box, spin));
  };
  const choice = (label: string, options: [string, string][], get: () => string, put: (v: string) => void, disabled = false): HTMLElement => {
    const s = h("select", { "aria-label": label, disabled }, ...options.map(([v, l]) => h("option", { value: v }, l))) as HTMLSelectElement;
    const show = () => { if (document.activeElement !== s) s.value = get(); };
    s.onchange = () => put(s.value);
    syncs.push(show);
    show();
    return h("div", { class: `fp-row fp-choice${disabled ? " off" : ""}` }, h("span", { class: "fp-label" }, label), s);
  };
  const check = (label: string, get: () => boolean, put: (v: boolean) => void, disabled = false): HTMLElement => {
    const i = h("input", { type: "checkbox", disabled }) as HTMLInputElement;
    const show = () => { i.checked = get(); };
    i.onchange = () => put(i.checked);
    syncs.push(show);
    show();
    return h("label", { class: `fp-check${disabled ? " off" : ""}` }, i, label);
  };
  const radios = (name: string, options: [string, string, boolean?][], cur: string, put: (v: string) => void): HTMLElement =>
    h("div", { class: "fp-radios", role: "radiogroup" }, ...options.map(([v, l, off]) =>
      h("label", { class: `fp-check${off ? " off" : ""}` }, h("input", { type: "radio", name, checked: v === cur, disabled: !!off, onchange: () => put(v) }), l)));
  const pickColor = (done: (hex: string) => void) => {
    const input = h("input", { type: "color", style: "position:fixed;opacity:0;width:0;height:0" }) as HTMLInputElement;
    document.body.append(input);
    input.addEventListener("change", () => { done(input.value.slice(1).toUpperCase()); input.remove(); });
    input.addEventListener("blur", () => setTimeout(() => input.remove(), 500));
    input.click();
  };
  const colorRow = (label: string, get: () => string | null, put: (hex: string) => void, disabled = false): HTMLElement => {
    const sw = h("span", { class: "fp-swatch" });
    const b = h("button", { class: "fp-drop", "aria-label": label, "data-tip": label, disabled }, sw, icon("chevron-down")) as HTMLButtonElement;
    b.onclick = () => {
      const r = b.getBoundingClientRect();
      const panel = h("div", { class: "rpanel palette" }, ...colorGrid((hex) => { closePopup(); put(hex); }, theme()), h("hr"),
        h("button", { class: "link", onclick: () => { closePopup(); pickColor(put); } }, "Weitere Farben..."));
      showPopup(panel, new DOMRect(r.left, r.bottom, r.width, 0));
    };
    const show = () => { const v = get(); sw.style.background = v ? `#${v}` : "transparent"; };
    syncs.push(show);
    show();
    return h("div", { class: `fp-row fp-choice${disabled ? " off" : ""}` }, h("span", { class: "fp-label" }, label), b);
  };
  /** Voreinstellungen: the ribbon's gallery for this setting, as a dropdown. */
  const presets = (label: string, iconName: string, sections: () => GallerySection[], disabled = false): HTMLElement => {
    const b = h("button", { class: "fp-drop", "aria-label": label, "data-tip": label, disabled }, icon(iconName), icon("chevron-down")) as HTMLButtonElement;
    b.onclick = () => {
      const e = ed();
      if (!e) return;
      const r = b.getBoundingClientRect();
      showPopup(galleryPanel(e, sections(), [], () => closePopup(), (item) => set(item.format)), new DOMRect(r.left, r.bottom, r.width, 0));
    };
    return h("div", { class: `fp-row fp-choice${disabled ? " off" : ""}` }, h("span", { class: "fp-label" }, label), b);
  };
  const resetButton = (run: () => void, enabled: boolean) => h("button", { class: "fp-reset", disabled: !enabled, onclick: run }, "Zurücksetzen");
  const sub = (title: string) => h("div", { class: "fp-sub" }, title);
  const rule = () => h("hr", { class: "fp-rule" });
  const note = (text: string) => h("div", { class: "fp-note" }, text);

  /** A gallery of Word's presets for one effect (the entries of the ribbon's menu). */
  const fxSections = (list: readonly (P.PicturePreset & { readonly id: string })[], keys: (keyof PictureFormat)[], noneTitle: string, columns: number, parts?: [string, number][]): GallerySection[] => {
    const items: GalleryItem[] = list.map((p) => {
      const f = pictureFromSpPr(p.xml, theme()) as Record<string, unknown>;
      const format: Record<string, unknown> = {};
      for (const k of keys) format[k] = f[k];
      if (p.extent && target()?.kind === "picture") format.extent = { l: p.extent[0], t: p.extent[1], r: p.extent[2], b: p.extent[3] };
      return { label: p.label, format: format as PictureFormat };
    });
    const none = items.filter((_, i) => list[i].id === "none");
    const rest = items.filter((_, i) => list[i].id !== "none");
    const out: GallerySection[] = none.length ? [{ title: noneTitle, items: none, columns: 1, tile: [48, 36] }] : [];
    if (parts) {
      let at = 0;
      for (const [title, n] of parts) { out.push({ title, items: rest.slice(at, at + n), columns, tile: [48, 36] }); at += n; }
    } else out.push({ items: rest, columns, tile: [48, 36] });
    return out;
  };

  // ----- sections --------------------------------------------------------------------------------
  const section = (id: string, title: string, build: () => (HTMLElement | null | false)[]): HTMLElement => {
    const open = opened.has(id);
    const head = h("button", {
      class: "fp-sec-head", "aria-expanded": String(open),
      onclick: () => { if (open) opened.delete(id); else opened.add(id); render(); },
    }, icon(open ? "chevron-down" : "chevron-right"), h("span", null, title));
    const el = h("div", { class: "fp-sec", "data-sec": id }, head);
    if (open) el.append(h("div", { class: "fp-sec-body" }, ...build().filter((x): x is HTMLElement => !!x)));
    return el;
  };

  // Füllung: shapes have their own fill (or the style's); a picture has one behind it.
  const fillKind = (t: Target): string => {
    if (t.kind === "picture") return /<a:solidFill>/.test(t.f.fillXml ?? "") ? "solid" : "none";
    const fl = t.f.fill ?? t.shape?.styleFill;
    return !fl ? "none" : fl.kind === "other" ? "picture" : fl.kind;
  };
  const fillColor = (t: Target): string => {
    if (t.kind === "picture") return /<a:srgbClr val="([0-9A-Fa-f]{6})"/.exec(t.f.fillXml ?? "")?.[1]?.toUpperCase() ?? "FFFFFF";
    const fl = t.f.fill ?? t.shape?.styleFill;
    return fl?.kind === "solid" || fl?.kind === "other" ? fl.color : fl?.kind === "gradient" ? fl.stops[0]?.color ?? "4472C4" : "4472C4";
  };
  const fillAlpha = (t: Target): number => {
    if (t.kind === "picture") { const a = /<a:alpha val="(\d+)"/.exec(t.f.fillXml ?? ""); return a ? Number(a[1]) / 100000 : 1; }
    const fl = t.f.fill ?? t.shape?.styleFill;
    return fl?.kind === "solid" ? 1 - (fl.transparency ?? 0) : 1;
  };
  const setSolidFill = (hex: string, alpha: number, start = true) => {
    const t = target();
    if (!t) return;
    if (t.kind === "picture") {
      const a = alpha < 0.999 ? `<a:alpha val="${Math.round(alpha * 100000)}"/>` : "";
      set({ fillXml: `<a:solidFill><a:srgbClr val="${hex}"${a ? `>${a}</a:srgbClr>` : "/>"}</a:solidFill>` }, start);
    } else set({ fill: { kind: "solid", color: hex, transparency: alpha < 0.999 ? 1 - alpha : undefined }, fillXml: undefined }, start);
  };
  const mix = (hex: string, k: number) => [0, 2, 4].map((i) => Math.round(parseInt(hex.slice(i, i + 2), 16) + (255 - parseInt(hex.slice(i, i + 2), 16)) * k).toString(16).padStart(2, "0")).join("").toUpperCase();
  const fillSection = (t: Target) => section("fill", "Füllung", () => {
    const k = fillKind(t);
    const put = (v: string) => {
      const cur = target();
      if (!cur) return;
      if (v === "none") set(cur.kind === "picture" ? { fillXml: undefined } : { fill: { kind: "none" }, fillXml: undefined });
      else if (v === "solid") setSolidFill(fillColor(cur), 1);
      else if (v === "gradient") { const base = fillColor(cur); set({ fill: { kind: "gradient", angle: 90, stops: [{ pos: 0, color: mix(base, 0.55) }, { pos: 1, color: base }] }, fillXml: undefined }); }
    };
    const out: HTMLElement[] = [radios("fp-fill", [["none", "Keine Füllung"], ["solid", "Einfarbige Füllung"], ["gradient", "Farbverlauf", t.kind === "picture"],
      ["picture", "Bild- oder Texturfüllung", true], ["pattern", "Musterfüllung", true]], k, put)];
    if (k === "solid") {
      out.push(rule(),
        colorRow("Farbe", () => { const x = target(); return x ? fillColor(x) : null; }, (hex) => { const x = target(); if (x) setSolidFill(hex, fillAlpha(x)); }),
        num({ label: "Transparenz", unit: "%", min: 0, max: 100, get: () => { const x = target(); return x ? Math.round((1 - fillAlpha(x)) * 100) : null; },
          put: (v, start) => { const x = target(); if (x) setSolidFill(fillColor(x), 1 - v / 100, start); } }));
    } else if (k === "gradient") {
      const grad = () => { const fl = target()?.f.fill ?? target()?.shape?.styleFill; return fl?.kind === "gradient" ? fl : null; };
      const putGrad = (patch: Partial<{ angle: number; path: string; stops: { pos: number; color: string; transparency?: number }[] }>, start = true) => {
        const g = grad();
        if (g) set({ fill: { ...g, ...patch }, fillXml: undefined }, start);
      };
      out.push(rule(),
        choice("Typ", [["linear", "Linear"], ["circle", "Radial"], ["rect", "Rechteckig"], ["shape", "Pfad"]], () => grad()?.path ?? "linear",
          (v) => putGrad(v === "linear" ? { path: undefined, angle: grad()?.angle ?? 90 } : { path: v, angle: undefined })),
        num({ label: "Winkel", unit: "°", min: 0, max: 359.9, step: 1, slider: false, disabled: !!grad()?.path, get: () => grad()?.angle ?? 0, put: (v, start) => putGrad({ angle: v }, start) }),
        sub("Farbverlaufstopps"),
        ...(grad()?.stops ?? []).flatMap((st, i) => [
          colorRow(`Stopp ${i + 1}: Farbe`, () => grad()?.stops[i]?.color ?? null, (hex) => { const g = grad(); if (g) putGrad({ stops: g.stops.map((s, j) => (j === i ? { ...s, color: hex } : s)) }); }),
          num({ label: `Stopp ${i + 1}: Position`, unit: "%", min: 0, max: 100, get: () => Math.round((grad()?.stops[i]?.pos ?? st.pos) * 100),
            put: (v, start) => { const g = grad(); if (g) putGrad({ stops: g.stops.map((s, j) => (j === i ? { ...s, pos: v / 100 } : s)) }, start); } }),
        ]));
    }
    return out;
  });

  // Linie: the picture's frame (Bildrahmen) or the shape's outline.
  type Line = NonNullable<NonNullable<PictureFormat["line"]>>;
  const lineOf = (t: Target): Line | null => {
    if (t.f.line === null) return null;
    const own = t.f.line;
    if (t.kind === "picture") return own ?? null;
    const style = t.shape?.styleLine;
    if (!own) return style ? { color: style.color, width: style.width } : null;
    return { ...own, color: own.color || style?.color || "000000", width: own.width >= 0 ? own.width : style?.width ?? 0.75 };
  };
  const setLine = (patch: Partial<Line> | null, start = true) => {
    const t = target();
    if (!t) return;
    if (patch === null) { set({ line: null }, start); return; }
    const cur = lineOf(t) ?? { color: t.shape?.styleLine?.color ?? "000000", width: 0.75 };
    set({ line: { ...cur, ...patch } }, start);
  };
  const isLineShape = (t: Target) => t.kind === "shape" && (t.shape?.kind === "connector" || (t.f.geometry ?? "rect") === "line" || !!t.shape?.custom);
  const lineSection = (t: Target) => section("line", "Linie", () => {
    const ln = lineOf(t);
    const out: HTMLElement[] = [radios("fp-line", [["none", "Keine Linie"], ["solid", "Einfarbige Linie"], ["gradient", "Farbverlaufslinie", true]], ln ? "solid" : "none",
      (v) => setLine(v === "none" ? null : {}))];
    if (!ln) return out;
    const cur = () => { const x = target(); return x ? lineOf(x) : null; };
    const arrows = isLineShape(t);
    const end = (which: "head" | "tail", label: string) => choice(label, ARROWS, () => cur()?.[which]?.type ?? "none",
      (v) => setLine({ [which]: v === "none" ? undefined : { ...(cur()?.[which] ?? {}), type: v } } as Partial<Line>), !arrows);
    const size = (which: "head" | "tail", label: string) => choice(label, ARROW_SIZES, () => `${cur()?.[which]?.w ?? "med"}/${cur()?.[which]?.len ?? "med"}`,
      (v) => { const [w, len] = v.split("/") as ("sm" | "med" | "lg")[]; const e0 = cur()?.[which]; if (e0) setLine({ [which]: { ...e0, w, len } } as Partial<Line>); },
      !arrows || !ln[which]);
    out.push(rule(),
      colorRow("Farbe", () => cur()?.color ?? null, (hex) => setLine({ color: hex })),
      num({ label: "Transparenz", unit: "%", min: 0, max: 100, get: () => Math.round((cur()?.transparency ?? 0) * 100), put: (v, start) => setLine({ transparency: v / 100 || undefined }, start) }),
      num({ label: "Stärke", unit: "Pt.", min: 0, max: 1584, step: 0.25, slider: false, decimals: 2, get: () => cur()?.width ?? 0, put: (v, start) => setLine({ width: v }, start) }),
      choice("Verbundtyp", COMPOUNDS, () => cur()?.compound ?? "sng", (v) => setLine({ compound: v === "sng" ? undefined : v })),
      choice("Strichtyp", DASHES, () => cur()?.dash ?? "solid", (v) => setLine({ dash: v === "solid" ? undefined : v })),
      choice("Abschlusstyp", CAPS, () => cur()?.cap ?? "flat", (v) => setLine({ cap: v === "flat" ? undefined : v })),
      choice("Anschlusstyp", JOINS, () => cur()?.join ?? (cur()?.miter ? "miter" : "round"), (v) => setLine({ join: v, miter: v === "miter" ? true : undefined })),
      end("head", "Startpfeiltyp"), size("head", "Startpfeilgröße"), end("tail", "Endpfeiltyp"), size("tail", "Endpfeilgröße"));
    return out;
  });

  // Effekte: Schatten, Spiegelung, Leuchten, Weiche Kanten, 3D-Format, 3D-Drehung.
  const shadowSection = () => section("shadow", "Schatten", () => {
    const sh = () => target()?.f.shadow;
    const put = (patch: Partial<PictureShadow>, start: boolean) => set({ shadow: { ...(sh() ?? DEFAULT_SHADOW), ...patch } }, start);
    return [
      presets("Voreinstellungen", "shadow", () => fxSections(P.SHADOWS, ["shadow"], "Kein Schatten", 3, [["Außen", 9], ["Innen", 9], ["Perspektive", 5]])),
      colorRow("Farbe", () => sh()?.color ?? null, (hex) => put({ color: hex }, true)),
      num({ label: "Transparenz", unit: "%", min: 0, max: 100, get: () => (sh() ? Math.round(sh()!.transparency * 100) : null), put: (v, s) => put({ transparency: v / 100 }, s) }),
      num({ label: "Größe", unit: "%", min: 1, max: 200, get: () => (sh() ? Math.round((sh()!.sx ?? 1) * 100) : null), put: (v, s) => put({ sx: v / 100, sy: v / 100 }, s) }),
      num({ label: "Weichzeichnen", unit: "Pt.", min: 0, max: 100, step: 0.5, get: () => sh()?.blur ?? null, put: (v, s) => put({ blur: v }, s) }),
      num({ label: "Winkel", unit: "°", min: 0, max: 359, get: () => (sh() ? Math.round(sh()!.dir) : null), put: (v, s) => put({ dir: v }, s) }),
      num({ label: "Abstand", unit: "Pt.", min: 0, max: 200, step: 0.5, get: () => sh()?.dist ?? null, put: (v, s) => put({ dist: v }, s) }),
    ];
  });
  const reflectionSection = () => section("reflection", "Spiegelung", () => {
    const rf = () => target()?.f.reflection;
    const put = (patch: Partial<NonNullable<PictureFormat["reflection"]>>, start: boolean) => set({ reflection: { ...(rf() ?? DEFAULT_REFLECTION), ...patch } }, start);
    return [
      presets("Voreinstellungen", "reflection", () => fxSections(P.REFLECTIONS, ["reflection"], "Keine Spiegelung", 3)),
      num({ label: "Transparenz", unit: "%", min: 0, max: 100, get: () => (rf() ? Math.round(rf()!.transparency * 100) : null), put: (v, s) => put({ transparency: v / 100 }, s) }),
      num({ label: "Größe", unit: "%", min: 1, max: 100, get: () => (rf() ? Math.round(rf()!.size * 100) : null), put: (v, s) => put({ size: v / 100 }, s) }),
      num({ label: "Weichzeichnen", unit: "Pt.", min: 0, max: 100, step: 0.5, get: () => rf()?.blur ?? null, put: (v, s) => put({ blur: v }, s) }),
      num({ label: "Abstand", unit: "Pt.", min: 0, max: 100, step: 0.5, get: () => rf()?.dist ?? null, put: (v, s) => put({ dist: v }, s) }),
    ];
  });
  const glowSection = () => section("glow", "Leuchten", () => {
    const gl = () => target()?.f.glow;
    const put = (patch: Partial<NonNullable<PictureFormat["glow"]>>, start: boolean) => set({ glow: { ...(gl() ?? DEFAULT_GLOW), ...patch } }, start);
    return [
      presets("Voreinstellungen", "effects", () => fxSections(P.GLOWS, ["glow"], "Kein Leuchteffekt", 6)),
      colorRow("Farbe", () => gl()?.color ?? null, (hex) => put({ color: hex }, true)),
      num({ label: "Größe", unit: "Pt.", min: 0, max: 150, step: 0.5, get: () => gl()?.radius ?? null, put: (v, s) => put({ radius: v }, s) }),
      num({ label: "Transparenz", unit: "%", min: 0, max: 100, get: () => (gl() ? Math.round(gl()!.transparency * 100) : null), put: (v, s) => put({ transparency: v / 100 }, s) }),
    ];
  });
  const softSection = () => section("softEdge", "Weiche Kanten", () => [
    presets("Voreinstellungen", "picture-effects", () => fxSections(P.SOFT_EDGES, ["softEdge"], "Keine weichen Kanten", 3)),
    num({ label: "Größe", unit: "Pt.", min: 0, max: 100, step: 0.5, get: () => target()?.f.softEdge ?? 0, put: (v, s) => set({ softEdge: v || undefined }, s) }),
  ]);
  const format3dSection = (t: Target) => section("format3d", "3D-Format", () => {
    const bv = () => target()?.f.bevel;
    const put = (patch: Partial<NonNullable<PictureFormat["bevel"]>>, start: boolean) => set({ bevel: { ...(bv() ?? { kind: "circle", w: 6, h: 6 }), ...patch } }, start);
    return [
      sub("Abschrägung oben"),
      presets("Voreinstellungen", "picture-effects", () => fxSections(P.BEVELS, ["bevel", "sp3dXml", "scene3dXml"], "Keine Abschrägung", 4)),
      num({ label: "Breite", unit: "Pt.", min: 0, max: 1584, step: 0.5, slider: false, get: () => bv()?.w ?? 0, put: (v, s) => put({ w: v }, s) }),
      num({ label: "Höhe", unit: "Pt.", min: 0, max: 1584, step: 0.5, slider: false, get: () => bv()?.h ?? 0, put: (v, s) => put({ h: v }, s) }),
      note("Abschrägung unten, Tiefe, Kontur, Material und Beleuchtung zeigt RaabeWrite nicht an; Word behält sie bei."),
      resetButton(() => set({ bevel: undefined, sp3dXml: undefined }), !!(t.f.bevel || t.f.sp3dXml)),
    ];
  });
  const rotation3dSection = (t: Target) => section("rotation3d", "3D-Drehung", () => {
    const r = () => rotOf(target()?.f.scene3dXml);
    const put = (patch: Partial<{ x: number; y: number; z: number }>, start: boolean) => set({ scene3dXml: withRot(target()?.f.scene3dXml, { ...r(), ...patch }) }, start);
    const turn = (label: string, k: "x" | "y" | "z") =>
      num({ label, unit: "°", min: 0, max: 359.9, step: 1, slider: false, get: () => Math.round(r()[k] * 10) / 10, put: (v, s) => put({ [k]: v }, s) });
    return [
      presets("Voreinstellungen", "rotate", () => fxSections(P.ROTATIONS_3D, ["scene3dXml"], "Keine Drehung", 4, [["Parallel", 10], ["Perspektive", 11], ["Schräg", 4]])),
      turn("X-Drehung", "x"), turn("Y-Drehung", "y"), turn("Z-Drehung", "z"),
      num({ label: "Perspektive", unit: "°", min: 0, max: 120, slider: false, disabled: true, get: () => null, put: () => {} }),
      check("Text nicht drehen", () => false, () => {}, true),
      resetButton(() => set({ scene3dXml: undefined }), !!t.f.scene3dXml),
    ];
  });
  const artisticSection = () => section("artistic", "Künstlerische Effekte", () => [
    presets("Künstlerische Effekte", "artistic", () => [], true),
    resetButton(() => {}, false),
  ]);

  // Layout und Eigenschaften: Textfeld (only shapes have text; for pictures it is off, as in Word).
  const textboxSection = (t: Target, id = "textbox") => section(id, "Textfeld", () => {
    const off = t.kind !== "shape";
    const body = () => target()?.shape?.body;
    const margin = (label: string, k: "l" | "r" | "t" | "b") =>
      num({ label, unit: "cm", min: 0, max: 50, step: 0.1, slider: false, decimals: 2, disabled: off, get: () => (body() ? body()![k] / CM : null), put: (v) => setBody({ [k]: v * CM }) });
    return [
      choice("Vertikale Ausrichtung", [["t", "Oben"], ["ctr", "Mitte"], ["b", "Unten"]], () => body()?.anchor ?? "t", (v) => setBody({ anchor: v as "t" | "ctr" | "b" }), off),
      choice("Textrichtung", [["horz", "Horizontal"], ["vert", "Gesamten Text um 90° drehen"], ["vert270", "Gesamten Text um 270° drehen"]],
        () => body()?.vert ?? "horz", (v) => setBody({ vert: v === "horz" ? undefined : v }), off),
      check("Text nicht drehen", () => false, () => {}, true),
      check("Größe der Form dem Text anpassen", () => body()?.autofit === "shape", (v) => setBody({ autofit: v ? "shape" : "none" }), off),
      off ? check("Textüberlauf über Form zulassen", () => false, () => {}, true) : null,
      margin("Linker Rand", "l"), margin("Rechter Rand", "r"), margin("Oberer Rand", "t"), margin("Unterer Rand", "b"),
      check("Text in Form umbrechen", () => !!body()?.wrap, (v) => setBody({ wrap: v }), off),
    ];
  });

  // Bild: Bildkorrekturen, Bildfarbe, Bildtransparenz, Zuschnitt.
  const pct = (k: "sharpen" | "brightness" | "contrast") => (): number => Math.round((target()?.f[k] ?? 0) * 100);
  const correctionsSection = (t: Target) => section("corrections", "Bildkorrekturen", () => [
    sub("Schärfen/Weichzeichnen"),
    presets("Voreinstellungen", "corrections", () => [{ items: SHARPEN, columns: 5, tile: [64, 48] }]),
    num({ label: "Schärfe", unit: "%", min: -100, max: 100, get: pct("sharpen"), put: (v, s) => set({ sharpen: v / 100 || undefined }, s) }),
    sub("Helligkeit/Kontrast"),
    presets("Voreinstellungen", "corrections", () => [{ items: BRIGHT_CONTRAST, columns: 5, tile: [64, 48] }]),
    num({ label: "Helligkeit", unit: "%", min: -100, max: 100, get: pct("brightness"), put: (v, s) => set({ brightness: v / 100 || undefined, lum: undefined }, s) }),
    num({ label: "Kontrast", unit: "%", min: -100, max: 100, get: pct("contrast"), put: (v, s) => set({ contrast: v / 100 || undefined, lum: undefined }, s) }),
    resetButton(() => set({ sharpen: undefined, brightness: undefined, contrast: undefined, lum: undefined }), !!(t.f.sharpen || t.f.brightness || t.f.contrast || t.f.lum)),
  ]);
  const colorSection = (t: Target) => section("color", "Bildfarbe", () => [
    sub("Farbsättigung"),
    presets("Voreinstellungen", "picture-color", () => [{ items: SATURATION, columns: 7, tile: [56, 42] }]),
    num({ label: "Sättigung", unit: "%", min: 0, max: 400, get: () => Math.round((target()?.f.saturation ?? 1) * 100), put: (v, s) => set({ saturation: Math.abs(v - 100) < 0.5 ? undefined : v / 100 }, s) }),
    sub("Farbton"),
    presets("Voreinstellungen", "picture-color", () => [{ items: TEMPERATURE, columns: 7, tile: [56, 42] }]),
    num({ label: "Temperatur", unit: "", min: 1500, max: 11500, step: 100, decimals: 0, get: () => target()?.f.temperature ?? 6500, put: (v, s) => set({ temperature: v === 6500 ? undefined : v }, s) }),
    presets("Neu einfärben", "picture", () => { const items = recolorGallery(theme()); return [{ items: items.slice(0, 7), columns: 7, tile: [56, 42] }, { items: items.slice(7), columns: 7, tile: [56, 42] }]; }),
    resetButton(() => set({ saturation: undefined, temperature: undefined, recolor: undefined, lum: undefined, transparentColor: undefined }),
      !!(t.f.saturation !== undefined || t.f.temperature || t.f.recolor || t.f.transparentColor)),
  ]);
  const transparencySection = () => section("transparency", "Bildtransparenz", () => [
    presets("Voreinstellungen", "transparency", () => [{ items: TRANSPARENCY, columns: 7, tile: [56, 42] }]),
    num({ label: "Transparenz", unit: "%", min: 0, max: 100, get: () => Math.round((target()?.f.transparency ?? 0) * 100), put: (v, s) => set({ transparency: v / 100 || undefined }, s) }),
  ]);
  // Zuschnitt: the whole picture (Bildposition: size and the offset of its middle from the
  // frame's middle) and the frame (Zuschnittsposition), in cm.
  const cropSection = () => section("crop", "Zuschnitt", () => {
    const geo = () => {
      const t = target();
      if (!t) return null;
      const c0 = t.f.crop ?? { l: 0, t: 0, r: 0, b: 0 };
      const iw = t.w / Math.max(0.01, 1 - c0.l - c0.r), ih = t.h / Math.max(0.01, 1 - c0.t - c0.b);
      return { w: t.w, h: t.h, iw, ih, ox: ((c0.r - c0.l) * iw) / 2, oy: ((c0.b - c0.t) * ih) / 2 };
    };
    const cropFor = (g: { w: number; h: number; iw: number; ih: number; ox: number; oy: number }) => {
      const l = (g.iw / 2 - g.w / 2 - g.ox) / g.iw, r = (g.iw / 2 - g.w / 2 + g.ox) / g.iw;
      const t = (g.ih / 2 - g.h / 2 - g.oy) / g.ih, b = (g.ih / 2 - g.h / 2 + g.oy) / g.ih;
      const near0 = (x: number) => (Math.abs(x) < 1e-6 ? 0 : x);
      const c = { l: near0(l), t: near0(t), r: near0(r), b: near0(b) };
      return c.l || c.t || c.r || c.b ? c : undefined;
    };
    const putPicture = (patch: Partial<{ iw: number; ih: number; ox: number; oy: number }>, start: boolean) => {
      const g = geo();
      if (g) set({ crop: cropFor({ ...g, ...patch }) }, start);
    };
    const putFrame = (patch: Partial<{ w: number; h: number }>, start: boolean) => {
      const e = ed();
      const g = geo();
      if (!e || !g) return;
      const next = { ...g, ...patch };
      e.setImageSize(next.w, next.h, start);
      e.setPictureFormat({ crop: cropFor(next) }, false);
      api.refresh();
    };
    const cm = (label: string, get: (g: NonNullable<ReturnType<typeof geo>>) => number, put: (v: number, s: boolean) => void, min = 0.01, disabled = false) =>
      num({ label, unit: "cm", min, max: 55.87, step: 0.1, slider: false, decimals: 2, disabled, get: () => { const g = geo(); return g ? get(g) / CM : null; }, put: (v, s) => put(v * CM, s) });
    return [
      sub("Bildposition"),
      cm("Breite", (g) => g.iw, (v, s) => putPicture({ iw: v }, s)),
      cm("Höhe", (g) => g.ih, (v, s) => putPicture({ ih: v }, s)),
      cm("X-Offset", (g) => g.ox, (v, s) => putPicture({ ox: v }, s), -55.87),
      cm("Y-Offset", (g) => g.oy, (v, s) => putPicture({ oy: v }, s), -55.87),
      sub("Zuschnittsposition"),
      cm("Breite", (g) => g.w, (v, s) => putFrame({ w: v }, s)),
      cm("Höhe", (g) => g.h, (v, s) => putFrame({ h: v }, s)),
      num({ label: "Links", unit: "cm", min: 0, max: 0, slider: false, disabled: true, get: () => null, put: () => {} }),
      num({ label: "Oben", unit: "cm", min: 0, max: 0, slider: false, disabled: true, get: () => null, put: () => {} }),
    ];
  });

  // ----- the pane ---------------------------------------------------------------------------------
  const TABS: Record<PaneTab, [string, string]> = {
    fillLine: ["Füllung und Linie", "pane-fill-line"], effects: ["Effekte", "pane-effects"], layout: ["Layout und Eigenschaften", "pane-layout"],
    picture: ["Bild", "pane-picture"], textFill: ["Textfüllung und -kontur", "text-fill"], textEffects: ["Texteffekte", "text-effects"],
    textLayout: ["Layout und Eigenschaften", "pane-layout"],
  };
  const SHAPE_TABS: PaneTab[] = ["fillLine", "effects", "layout"];
  const TEXT_TABS: PaneTab[] = ["textFill", "textEffects", "textLayout"];
  const tabNow = (t: Target): PaneTab => {
    if (t.kind === "picture") return pictureTab === "textFill" || pictureTab === "textEffects" || pictureTab === "textLayout" ? "effects" : pictureTab;
    return shapeTab === "picture" ? "fillLine" : shapeTab;
  };
  const choose = (tab: PaneTab) => {
    const t = target();
    if (t?.kind === "shape") shapeTab = tab;
    else pictureTab = tab;
    render();
  };

  const close = () => { isOpen = false; pane.hidden = true; builtKey = ""; api.refresh(); api.focus(); };
  const head = (title: string) => h("div", { class: "fx-head" }, h("span", null, title),
    h("span", { class: "fp-headbtns" },
      h("button", {
        "aria-label": "Optionen für den Aufgabenbereich", "data-tip": "Optionen für den Aufgabenbereich",
        onclick: (ev: MouseEvent) => showMenu(c, [{ label: "Schließen", run: close }], (ev.currentTarget as HTMLElement).getBoundingClientRect()),
      }, icon("chevron-down")),
      h("button", { "aria-label": "Bereich schließen", "data-tip": "Bereich schließen", onclick: close }, icon("close"))));

  const keyOf = (t: Target | null): string => {
    if (!t) return "none";
    const tab = tabNow(t);
    return JSON.stringify([t.kind, tab, [...opened].sort(), t.float, fillKind(t), !!lineOf(t), isLineShape(t), !!lineOf(t)?.head, !!lineOf(t)?.tail,
      t.f.fill?.kind === "gradient" ? [t.f.fill.path ?? "", t.f.fill.stops.length] : 0,
      !!(t.f.bevel || t.f.sp3dXml), !!t.f.scene3dXml, !!(t.f.sharpen || t.f.brightness || t.f.contrast || t.f.lum),
      !!(t.f.saturation !== undefined || t.f.temperature || t.f.recolor || t.f.transparentColor),
      tab === "textFill" || tab === "textEffects" ? ed()?.textFxAt() ?? "" : ""]);
  };

  function render(): void {
    if (!isOpen) return;
    const t = target();
    const key = keyOf(t);
    if (key === builtKey) { for (const s of syncs) s(); return; }
    builtKey = key;
    syncs = [];
    const oldBody = pane.querySelector<HTMLElement>(".fx-body");
    const scroll = oldBody?.scrollTop ?? 0;
    if (!t) {
      pane.setAttribute("aria-label", "Grafik formatieren");
      pane.replaceChildren(head("Grafik formatieren"), h("div", { class: "fx-body" }, note("Markieren Sie ein Bild oder eine Form, um es hier zu formatieren.")));
      return;
    }
    const title = t.kind === "picture" ? "Grafik formatieren" : "Form formatieren";
    pane.setAttribute("aria-label", title);
    const tab = tabNow(t);
    const parts: HTMLElement[] = [head(title)];
    if (t.kind === "shape") {
      // Formoptionen and Textoptionen, each with its row of tabs.
      const textSide = TEXT_TABS.includes(tab);
      parts.push(h("div", { class: "fp-outer", role: "tablist" },
        h("button", { class: textSide ? "" : "active", role: "tab", "aria-selected": String(!textSide), onclick: () => choose("fillLine") }, "Formoptionen"),
        h("button", { class: textSide ? "active" : "", role: "tab", "aria-selected": String(textSide), onclick: () => choose("textFill") }, "Textoptionen")));
    }
    const tabs = t.kind === "picture" ? (["fillLine", "effects", "layout", "picture"] as PaneTab[]) : TEXT_TABS.includes(tab) ? TEXT_TABS : SHAPE_TABS;
    parts.push(h("div", { class: "fx-tabs", role: "tablist" }, ...tabs.map((id) => h("button", {
      class: id === tab ? "active" : "", role: "tab", "aria-selected": String(id === tab), "aria-label": TABS[id][0], "data-tip": TABS[id][0], onclick: () => choose(id),
    }, icon(TABS[id][1])))));
    let body: HTMLElement;
    if (tab === "textFill" || tab === "textEffects") body = text.body(tab === "textFill" ? "fill" : "effects");
    else {
      body = h("div", { class: "fx-body" });
      if (tab === "fillLine") body.append(fillSection(t), lineSection(t));
      else if (tab === "effects") body.append(shadowSection(), reflectionSection(), glowSection(), softSection(), format3dSection(t), rotation3dSection(t), ...(t.kind === "picture" ? [artisticSection()] : []));
      else if (tab === "layout") body.append(textboxSection(t));
      else if (tab === "textLayout") body.append(textboxSection(t, "textLayout"));
      else body.append(correctionsSection(t), colorSection(t), transparencySection(), cropSection());
    }
    parts.push(body);
    pane.replaceChildren(...parts);
    body.scrollTop = scroll;
    if (reveal) {
      const sec = body.querySelector<HTMLElement>(`[data-sec="${reveal}"]`);
      reveal = null;
      if (sec) body.scrollTop = sec.offsetTop - body.offsetTop;
    }
  }

  const open = (sectionId?: FormatSection) => {
    const t = target();
    if (!t) return;
    if (sectionId) {
      const tab = TAB_OF[sectionId];
      if (t.kind === "shape") shapeTab = tab;
      else pictureTab = tab;
      opened.add(sectionId);
      reveal = sectionId;
    }
    isOpen = true;
    pane.hidden = false;
    builtKey = "";
    render();
    api.refresh();
  };
  FORMAT_PANE.open = open;

  const picked = () => !!target();
  // The group launchers: Bildformatvorlagen opens the pane at Effekte, Formenarten at Füllung und Linie.
  c.implement("picture.formatPane", { run: () => { pictureTab = "effects"; open(); }, enabled: picked });
  c.implement("shape.formatPane", { run: () => { shapeTab = "fillLine"; open(); }, enabled: picked });
  return { open, refresh: render, isOpen: () => isOpen };
}
