// The contextual tab "Formformat" as Word 365 has it (D-046): Formen einfügen (gallery, Form
// bearbeiten, Textfeld erstellen), Formenarten (Word's 77 styles, Fülleffekt, Formkontur,
// Formeffekte), WordArt-Formate, Text (Textrichtung, Text ausrichten), Barrierefreiheit,
// Anordnen and Größe. Layouts and reduction order from capture_ribbon_widths.py Formformat;
// the styles were recorded from Word (shape-presets.ts).

import { closePopup, colorGrid, h, showDialog, showMenu, type Command, type Control, type MenuEntry, type Tab } from "@raabe/shell";
import { shapeElement, type PictureFormat, type ShapeInfo } from "@raabe/text-engine";
import { DEFAULT_THEME_COLORS, pictureFromSpPr, shapePreset } from "@raabe/ooxml";
import { CM, fmtCm, parseNum, type AppApi } from "./api";
import { galleryPanel as pictureGallery, type GalleryItem, type GallerySection } from "./picture-format";
import * as P from "./picture-presets";
import { SHAPE_STYLES } from "./shape-presets";
import { WORDART } from "./text-effects";
import { GROUP_MENU } from "./groups";
import { FORMAT_PANE, type FormatSection } from "./format-pane";

const cmd = (id: string, label: string, icon?: string, extra: Partial<Command> = {}): Command => ({ id, label, icon, ...extra });

export const SHAPE_TAB_COMMANDS: Command[] = [
  cmd("shape.gallery", "Formen", "shapes", { tip: "Formen\nEine vorgefertigte Form einfügen, z. B. einen Kreis, ein Quadrat oder einen Pfeil." }),
  cmd("shape.edit", "Form bearbeiten", "edit-shape", { tip: "Form bearbeiten\nDie Form dieser Zeichnung ändern, sie in eine Freihandform konvertieren oder die Umbruchpunkte bearbeiten, um zu bestimmen, wie Text um die Zeichnung umbrochen wird." }),
  cmd("shape.styles", "Schnellformatvorlagen", "styles", { tip: "Formenarten\nEine visuelle Formatvorlage für die Form oder Linie auswählen." }),
  cmd("shape.fill", "Fülleffekt", "shape-fill", { tip: "Fülleffekt\nDie ausgewählte Form mit einer Volltonfarbe, einem Farbverlauf, einem Bild oder einer Textur füllen." }),
  cmd("shape.outline", "Formkontur", "shape-outline", { tip: "Formkontur\nDie Farbe, die Breite und die Linienart für die Kontur der ausgewählten Form auswählen." }),
  cmd("shape.effects", "Formeffekte", "shape-effects", { tip: "Formeffekte\nDer ausgewählten Form einen visuellen Effekt hinzufügen, z. B. einen Schatten, ein Leuchten, eine Spiegelung oder eine 3D-Drehung." }),
  cmd("shape.formatPane", "Form formatieren", undefined),
  cmd("shape.textDirection", "Textrichtung", "text-direction", { tip: "Textrichtung\nDie Ausrichtung des Texts in vertikal, gestapelt oder gedreht ändern." }),
  cmd("shape.textAlign", "Text ausrichten", "text-align-middle", { tip: "Text ausrichten\nAuswählen, wie Text im Textfeld ausgerichtet wird." }),
  cmd("shape.alignTop", "Oben", "text-align-top"),
  cmd("shape.alignMiddle", "Zentriert", "text-align-middle"),
  cmd("shape.alignBottom", "Unten", "text-align-bottom"),
  cmd("shape.horizontal", "Horizontal", "text-direction"),
  cmd("shape.rotate90", "Gesamten Text um 90° drehen", "text-direction"),
  cmd("shape.rotate270", "Gesamten Text um 270° drehen", "text-direction"),
  cmd("wordart.styles", "Schnellformatvorlagen", "wordart", { tip: "WordArt-Formate\nDem Text einen künstlerischen Effekt geben." }),
  cmd("text.fill", "Textfüllung", "text-fill", { tip: "Textfüllung\nDen Text mit einer Volltonfarbe, einem Farbverlauf, einem Bild oder einer Textur füllen." }),
  cmd("text.outline", "Textkontur", "text-outline", { tip: "Textkontur\nDie Farbe, die Breite und die Linienart für die Kontur des Texts anpassen." }),
  cmd("text.effects", "Texteffekte", "text-effects", { tip: "Texteffekte\nDem Text einen visuellen Effekt hinzufügen, z. B. einen Schatten, ein Leuchten oder eine Spiegelung." }),
  cmd("text.effectsPane", "Texteffekte formatieren", undefined),
  cmd("shape.createLink", "Verknüpfung erstellen", "create-link", { tip: "Verknüpfung erstellen\nDieses Textfeld mit einem anderen verknüpfen, sodass der Text vom einen in das andere fließt." }),
  cmd("shape.height", "Formenhöhe", "shape-height"),
  cmd("shape.width", "Formenbreite", "shape-width"),
];

/** Shapes for the in-ribbon gallery of Formen einfügen: icon, name, drawing. */
export interface RibbonShapes { count: number; icon: (i: number) => Element; label: (i: number) => string; draw: (i: number) => void }

export function shapeTab(api: AppApi, galleryPanel: (close: () => void) => HTMLElement, ribbonShapes: RibbonShapes) {
  const c = api.commands;
  const ed = () => api.editor();
  const shape = () => ed()?.selectedShape() ?? null;
  const picked = () => !!shape();
  const theme = () => ({ ...DEFAULT_THEME_COLORS, ...(ed()?.doc.styles.theme.colors ?? {}) });
  // Commands that act on the shape itself leave its text first (as Esc would).
  const onShape = (fn: () => void) => () => {
    const e = ed();
    if (e?.inShapeText) e.leaveShapeText(true);
    fn();
    api.refresh();
  };
  const run = (id: string, fn: () => void, enabled: () => boolean = picked, checked?: () => boolean) =>
    c.implement(id, { run: () => { fn(); api.focus(); }, enabled, checked });
  const setFormat = (f: PictureFormat | ((f: PictureFormat) => PictureFormat)) => { ed()?.setShapeFormat(f); api.refresh(); };

  // ----- Formen einfügen -------------------------------------------------------------------
  c.implement("shape.gallery", { run: () => {}, enabled: () => !!ed() });
  const formChange = (): MenuEntry[] => [];
  c.implement("shape.edit", { run: () => {}, enabled: picked });
  const editMenu: MenuEntry[] = [
    { label: "Form ändern", icon: "edit-shape", flyout: () => changeShapePanel() },
    { label: "Punkte bearbeiten", disabled: true, run: () => {} },
    "-",
    { label: "Umleiten von Verbindungen", disabled: true, run: () => {} },
  ];
  void formChange;
  // Form ändern: the gallery without lines and text box; the new geometry keeps the look.
  let changeShapePanel: () => HTMLElement = () => h("div");
  const setChangePanel = (fn: () => HTMLElement) => { changeShapePanel = fn; };

  // ----- Formenarten -----------------------------------------------------------------------
  const presetFor = (i: number) => shapePreset(SHAPE_STYLES[i].style, SHAPE_STYLES[i].spPr, theme());
  const applyStyle = (i: number, snapshot = true) => {
    const e = ed();
    const s = shape();
    if (!e || !s) return;
    const p = presetFor(i);
    // Word replaces the style and the shape's own fill and outline.
    e.setShapeFormat((f) => ({ ...f, fill: p.pic.fill, fillXml: p.pic.fillXml, line: p.pic.line }), snapshot);
    e.setShape((sh) => ({ ...sh, styleFill: p.shape.styleFill, styleLine: p.shape.styleLine, styleShadow: p.shape.styleShadow, fontColor: p.shape.fontColor, styleXml: p.shape.styleXml }), false);
  };
  const styleLook = (i: number, base: ShapeInfo, f: PictureFormat): [ShapeInfo, PictureFormat] => {
    const p = presetFor(i);
    return [{ ...base, ...p.shape }, { ...f, fill: p.pic.fill, fillXml: p.pic.fillXml, line: p.pic.line }];
  };
  const stylesPanel = (close: () => void) => {
    const s = shape();
    if (!s) return h("div");
    const tile = (i: number) => {
      const [info, look] = styleLook(i, s.obj.shape!, s.obj.pic ?? {});
      const b = h("button", { class: "tile shape-style-tile", "data-tip": SHAPE_STYLES[i].label, "aria-label": SHAPE_STYLES[i].label, onclick: () => { close(); applyStyle(i); api.refresh(); api.focus(); } });
      const el = shapeElement(64, 34, { ...look, geometry: "roundRect", rotation: undefined, flipH: undefined, flipV: undefined }, { ...info, kind: "shape", adjust: {} });
      el.style.position = "relative";
      el.style.margin = "4px";
      b.append(el);
      const abc = h("span", { class: "abc", style: `color:#${info.fontColor ?? "000000"}` }, "Abc");
      b.append(abc);
      return b;
    };
    return h("div", { class: "rpanel shape-styles" },
      h("div", { class: "title" }, "Designformatvorlagen"), h("div", { class: "tiles", style: "grid-template-columns:repeat(7, auto)" }, ...SHAPE_STYLES.slice(0, 42).map((_, i) => tile(i))),
      h("div", { class: "title" }, "Voreinstellungen"), h("div", { class: "tiles", style: "grid-template-columns:repeat(7, auto)" }, ...SHAPE_STYLES.slice(42).map((_, i) => tile(42 + i))));
  };
  c.implement("shape.styles", { run: () => {}, enabled: picked });
  let hoverBase: { shape: ShapeInfo; pic: PictureFormat | undefined } | null = null;
  const restoreHover = () => {
    const e = ed();
    if (!e || !hoverBase) return;
    const hb = hoverBase;
    hoverBase = null;
    e.setShapeFormat(() => hb.pic ?? {}, false);
    e.setShape(() => hb.shape, false);
  };
  const styleTiles = (visible: number): Control => ({
    type: "tiles", cmd: "shape.styles", key: "K", count: SHAPE_STYLES.length, visible, fill: true, size: [64, 48],
    render: (i) => {
      const s = shape();
      if (!s) return null;
      const base = hoverBase ?? { shape: s.obj.shape!, pic: s.obj.pic };
      const [info, look] = styleLook(i, base.shape, base.pic ?? {});
      const el = shapeElement(46, 30, { ...look, geometry: "rect", rotation: undefined, flipH: undefined, flipV: undefined }, { ...info, kind: "shape", adjust: {} });
      el.style.position = "relative";
      return el;
    },
    tip: (i) => SHAPE_STYLES[i].label,
    pick: (i) => { restoreHover(); applyStyle(i); api.refresh(); api.focus(); },
    hover: (i) => {
      const e = ed();
      const s = shape();
      if (!e || !s) return;
      if (i === null) { restoreHover(); return; }
      hoverBase ??= { shape: s.obj.shape!, pic: s.obj.pic };
      const [info, look] = styleLook(i, hoverBase.shape, hoverBase.pic ?? {});
      e.setShapeFormat(() => look, false);
      e.setShape(() => info, false);
    },
    current: () => undefined,
    redraw: () => { const s = shape(); return s ? `${s.obj.pic?.geometry ?? "rect"}` : ""; },
    panel: stylesPanel,
  });

  // Fülleffekt: theme colours, Keine Füllung, Weitere Füllfarben, Farbverlauf.
  const fillPanel = (close: () => void) => {
    const colors = colorGrid((hex) => { close(); setFormat({ fill: { kind: "solid", color: hex }, fillXml: undefined }); api.focus(); });
    const link = (label: string, fn: () => void, sub = false) => h("button", { class: `link${sub ? " has-sub" : ""}`, onclick: fn }, label);
    return h("div", { class: "rpanel palette" }, ...colors, h("hr"),
      link("Keine Füllung", () => { close(); setFormat({ fill: { kind: "none" }, fillXml: undefined }); api.focus(); }),
      link("Weitere Füllfarben...", () => { close(); void moreColor("Farben", (hex) => setFormat({ fill: { kind: "solid", color: hex }, fillXml: undefined })); }),
      link("Farbverlauf", (ev?: unknown) => { void ev; close(); gradientMenu(); }, true));
  };
  const gradientMenu = () => {
    const s = shape();
    if (!s) return;
    const base = currentFillColor(s.obj.pic, s.obj.shape!);
    const light = (k: number) => mix(base, k, true), dark = (k: number) => mix(base, k, false);
    const items: GalleryItem[] = [
      { label: "Keine Farbverläufe", format: { fill: { kind: "solid", color: base }, fillXml: undefined } },
      ...[0, 45, 90, 135, 180, 225, 270, 315, 360].map((ang, i): GalleryItem => ({
        label: `Helle Variante ${i + 1}`,
        format: { fill: i === 8 ? { kind: "gradient", path: "circle", stops: [{ pos: 0, color: light(0.4) }, { pos: 1, color: base }] } : { kind: "gradient", angle: ang, stops: [{ pos: 0, color: light(0.45) }, { pos: 1, color: base }] }, fillXml: undefined },
      })),
      ...[0, 45, 90, 135, 180, 225, 270, 315, 360].map((ang, i): GalleryItem => ({
        label: `Dunkle Variante ${i + 1}`,
        format: { fill: i === 8 ? { kind: "gradient", path: "circle", stops: [{ pos: 0, color: base }, { pos: 1, color: dark(0.6) }] } : { kind: "gradient", angle: ang, stops: [{ pos: 0, color: base }, { pos: 1, color: dark(0.6) }] }, fillXml: undefined },
      })),
    ];
    const e = ed();
    if (!e) return;
    const sections: GallerySection[] = [
      { items: items.slice(0, 1), columns: 1, tile: [48, 36] },
      { title: "Helle Varianten", items: items.slice(1, 10), columns: 3, tile: [48, 36] },
      { title: "Dunkle Varianten", items: items.slice(10), columns: 3, tile: [48, 36] },
    ];
    const anchor = document.querySelector("[data-cmd='shape.fill']")?.getBoundingClientRect() ?? new DOMRect(200, 120, 0, 0);
    const panel = pictureGallery(e, sections, [], () => closePopup(), (item) => setFormat(item.format));
    import("@raabe/shell").then((m) => m.showPopup(panel, new DOMRect(anchor.left, anchor.bottom, 0, 0)));
  };
  run("shape.fill", onShape(() => { const s = shape(); if (s) setFormat({ fill: { kind: "solid", color: currentFillColor(s.obj.pic, s.obj.shape!) }, fillXml: undefined }); }));

  // Formkontur: colours, Keine Kontur, Weitere Konturfarben, Stärke, Striche, Pfeile.
  const lineOf = () => { const s = shape(); return s ? s.obj.pic?.line ?? (s.obj.shape?.styleLine ? { color: s.obj.shape.styleLine.color, width: s.obj.shape.styleLine.width } : undefined) : undefined; };
  const setLine = (patch: Partial<NonNullable<PictureFormat["line"]>> | null) => {
    if (patch === null) { setFormat({ line: null }); return; }
    const cur = lineOf();
    const base = cur && cur.color !== "" ? cur : { color: shape()?.obj.shape?.styleLine?.color ?? "000000", width: shape()?.obj.shape?.styleLine?.width ?? 0.75, ...(cur ?? {}) };
    setFormat({ line: { ...base, color: base.color || "000000", width: base.width >= 0 ? base.width : 0.75, ...patch } });
  };
  const weights = [0.25, 0.5, 0.75, 1, 1.5, 2.25, 3, 4.5, 6];
  const dashes: [string, string][] = [["solid", "Durchgezogen"], ["sysDot", "Rundpunkt"], ["sysDash", "Quadratpunkt"], ["dash", "Strich"], ["dashDot", "Strich-Punkt"], ["lgDash", "Langer Strich"], ["lgDashDot", "Lang-Strich-Punkt"], ["lgDashDotDot", "Lang-Strich-Punkt-Punkt"]];
  const arrows: [string, Partial<NonNullable<PictureFormat["line"]>>][] = [
    ["Pfeilart 1", { head: undefined, tail: undefined }], ["Pfeilart 2", { head: undefined, tail: { type: "arrow" } }], ["Pfeilart 3", { head: { type: "arrow" }, tail: undefined }],
    ["Pfeilart 4", { head: { type: "arrow" }, tail: { type: "arrow" } }], ["Pfeilart 5", { head: undefined, tail: { type: "triangle" } }], ["Pfeilart 6", { head: { type: "triangle" }, tail: undefined }],
    ["Pfeilart 7", { head: { type: "triangle" }, tail: { type: "triangle" } }], ["Pfeilart 8", { head: { type: "oval" }, tail: { type: "triangle" } }],
    ["Pfeilart 9", { head: { type: "oval" }, tail: { type: "oval" } }], ["Pfeilart 10", { head: { type: "diamond" }, tail: { type: "diamond" } }],
    ["Pfeilart 11", { head: { type: "oval" }, tail: { type: "stealth" } }],
  ];
  const moreLines: MenuEntry = { label: "Weitere Linien...", run: () => FORMAT_PANE.open("line") };
  const outlinePanel = (close: () => void) => {
    const colors = colorGrid((hex) => { close(); setLine({ color: hex }); api.focus(); });
    const sub = (ev: MouseEvent, entries: MenuEntry[]) => {
      const r = (ev.currentTarget as HTMLElement).getBoundingClientRect();
      close();
      showMenu(c, entries, new DOMRect(r.right, r.top, 0, 0));
    };
    const link = (label: string, fn: (ev: MouseEvent) => void, subMenu = false) => h("button", { class: `link${subMenu ? " has-sub" : ""}`, onclick: fn }, label);
    const isLine = () => { const s = shape(); return !!s && (s.obj.shape?.kind === "connector" || (s.obj.pic?.geometry ?? "rect") === "line" || !!s.obj.shape?.custom); };
    return h("div", { class: "rpanel palette" }, ...colors, h("hr"),
      link("Keine Kontur", () => { close(); setLine(null); api.focus(); }),
      link("Weitere Konturfarben...", () => { close(); void moreColor("Farben", (hex) => setLine({ color: hex })); }),
      link("Stärke", (ev) => sub(ev, [...weights.map((w): MenuEntry => ({ label: `${w.toLocaleString("de-DE")} Pt.`, checked: lineOf()?.width === w, run: () => setLine({ width: w }) })), "-", moreLines]), true),
      link("Striche", (ev) => sub(ev, [...dashes.map(([v, l]): MenuEntry => ({ label: l, checked: (lineOf()?.dash ?? "solid") === v, run: () => setLine({ dash: v === "solid" ? undefined : v }) })), "-", moreLines]), true),
      ...(isLine() ? [link("Pfeile", (ev) => sub(ev, [...arrows.map(([l, patch]): MenuEntry => ({ label: l, run: () => setLine(patch) })), "-", { label: "Weitere Pfeile...", run: () => FORMAT_PANE.open("line") }]), true)] : []));
  };
  run("shape.outline", onShape(() => setLine({})));

  // Formeffekte: Word's presets (the same as for pictures).
  const fx = (title: string, noneLabel: string, list: readonly (P.PicturePreset & { readonly id: string })[], keys: (keyof PictureFormat)[], columns: number, sections?: [string, number][], more?: [string, FormatSection]) => (): HTMLElement => {
    const e = ed();
    if (!e) return h("div");
    const items = list.map((p) => {
      const f = pictureFromSpPr(p.xml, theme()) as Record<string, unknown>;
      const format: Record<string, unknown> = {};
      for (const k of keys) format[k] = f[k];
      return { label: p.label, format: format as PictureFormat } as GalleryItem;
    });
    const none = items.find((_, i) => list[i].id === "none");
    const rest = items.filter((_, i) => list[i].id !== "none");
    const parts: GallerySection[] = [];
    if (none) parts.push({ title: noneLabel, items: [none], columns: 1, tile: [48, 36] });
    if (sections) {
      let at = 0;
      for (const [t, n] of sections) { parts.push({ title: t, items: rest.slice(at, at + n), columns, tile: [48, 36] }); at += n; }
    } else parts.push({ title, items: rest, columns, tile: [48, 36] });
    return pictureGallery(e, parts, more ? [{ label: more[0], run: () => FORMAT_PANE.open(more[1]) }] : [], () => closePopup(), (item) => setFormat((f) => ({ ...f, ...item.format })));
  };
  const effectsMenu: MenuEntry[] = [
    { label: "Voreinstellung", icon: "shape-effects", flyout: fx("Voreinstellungen", "Keine Voreinstellungen", P.EFFECT_PRESETS, ["shadow", "glow", "reflection", "softEdge", "bevel", "scene3dXml", "sp3dXml"], 4, undefined, ["3D-Optionen...", "format3d"]) },
    { label: "Schatten", icon: "shape-effects", flyout: fx("Schatten", "Kein Schatten", P.SHADOWS, ["shadow"], 3, [["Außen", 9], ["Innen", 9], ["Perspektive", 5]], ["Schattenoptionen...", "shadow"]) },
    { label: "Spiegelung", icon: "shape-effects", flyout: fx("Spiegelungsvarianten", "Keine Spiegelung", P.REFLECTIONS, ["reflection"], 3, undefined, ["Spiegelungsoptionen...", "reflection"]) },
    { label: "Leuchten", icon: "shape-effects", flyout: fx("Leuchteffektvarianten", "Kein Leuchteffekt", P.GLOWS, ["glow"], 6, undefined, ["Leuchteffektoptionen...", "glow"]) },
    { label: "Weiche Kanten", icon: "shape-effects", flyout: fx("Varianten weicher Kanten", "Keine weichen Kanten", P.SOFT_EDGES, ["softEdge"], 3, undefined, ["Optionen für weiche Kanten...", "softEdge"]) },
    { label: "Abschrägung", icon: "shape-effects", flyout: fx("Abschrägung", "Keine Abschrägung", P.BEVELS, ["bevel", "sp3dXml", "scene3dXml"], 4, undefined, ["3D-Optionen...", "format3d"]) },
    { label: "3D-Drehung", icon: "shape-effects", flyout: fx("3D-Drehung", "Keine Drehung", P.ROTATIONS_3D, ["scene3dXml"], 4, [["Parallel", 10], ["Perspektive", 11], ["Schräg", 4]], ["3D-Drehungsoptionen...", "rotation3d"]) },
  ];
  c.implement("shape.effects", { run: () => {}, enabled: picked });

  // ----- Text ------------------------------------------------------------------------------
  const body = () => shape()?.obj.shape?.body;
  const setBody = (patch: Partial<ShapeInfo["body"]>) => { ed()?.setShape((s) => ({ ...s, body: { ...s.body, ...patch } })); api.refresh(); };
  run("shape.alignTop", () => setBody({ anchor: "t" }), picked, () => body()?.anchor === "t");
  run("shape.alignMiddle", () => setBody({ anchor: "ctr" }), picked, () => body()?.anchor === "ctr");
  run("shape.alignBottom", () => setBody({ anchor: "b" }), picked, () => body()?.anchor === "b");
  run("shape.horizontal", () => setBody({ vert: undefined }), picked, () => !body()?.vert);
  run("shape.rotate90", () => setBody({ vert: "vert" }), picked, () => body()?.vert === "vert");
  run("shape.rotate270", () => setBody({ vert: "vert270" }), picked, () => body()?.vert === "vert270");
  c.implement("shape.textAlign", { run: () => {}, enabled: picked });
  c.implement("shape.textDirection", { run: () => {}, enabled: picked });
  const alignMenu: MenuEntry[] = [{ cmd: "shape.alignTop" }, { cmd: "shape.alignMiddle" }, { cmd: "shape.alignBottom" }];
  const directionMenu: MenuEntry[] = [{ cmd: "shape.horizontal" }, { cmd: "shape.rotate90" }, { cmd: "shape.rotate270" }, "-", { label: "Textrichtungsoptionen...", run: () => FORMAT_PANE.open("textLayout") }];

  // ----- Größe -----------------------------------------------------------------------------
  const size = () => { const s = shape(); return s ? { w: s.obj.width ?? 0, h: s.obj.height ?? 0 } : null; };
  const setSide = (side: "w" | "h", pt: number) => {
    const s = size();
    if (!s || !(pt >= 0)) return;
    const e = ed();
    if (e?.inShapeText) e.leaveShapeText(true);
    ed()?.setImageSize(side === "w" ? pt : s.w, side === "h" ? pt : s.h);
    api.refresh();
  };
  const spinner = (id: string, key: string, side: "w" | "h", label: string): Control => ({
    type: "spinner", cmd: id, key, label, icon: side === "h" ? "shape-height" : "shape-width",
    value: () => { const s = size(); return s ? fmtCm(side === "w" ? s.w : s.h) : ""; },
    commit: (v) => setSide(side, parseNum(v) * CM),
    step: (d) => { const s = size(); if (s) setSide(side, Math.max(0, (side === "w" ? s.w : s.h) + d * 0.1 * CM)); },
  });
  c.implement("shape.height", { run: () => {}, enabled: picked });
  c.implement("shape.width", { run: () => {}, enabled: picked });

  // ----- the tab ---------------------------------------------------------------------------
  const B = (cmdId: string, key: string, size: "large" | "medium" | "small", label?: string): Control => ({ type: "button", cmd: cmdId, key, size, label });
  const D = (cmdId: string, key: string, size: "large" | "medium" | "small", menu: MenuEntry[], label?: string): Control => ({ type: "dropdown", cmd: cmdId, key, size, menu, label });
  const PD = (cmdId: string, key: string, size: "large" | "medium" | "small", popup: (close: () => void) => HTMLElement, label?: string): Control => ({ type: "dropdown", cmd: cmdId, key, size, menu: [], popup, label });
  /** WordArt-Formate menus (text-effects.ts fills them in). */
  const WD = (cmdId: string, key: string, size: "large" | "medium" | "small", get: () => MenuEntry[]): Control => ({ type: "dropdown", cmd: cmdId, key, size, get menu() { return get(); } });
  const SP = (cmdId: string, key: string, size: "medium" | "small", popup: (close: () => void) => HTMLElement): Control => ({ type: "split", cmd: cmdId, key, size, menu: [], popup });
  const stack = (...items: Control[]): Control => ({ type: "stack", items });
  const shapesTiles: Control = {
    type: "tiles", cmd: "shape.gallery", key: "SH", count: ribbonShapes.count, visible: ribbonShapes.count, size: [18, 18], cls: "shape-tiles-ribbon",
    render: (i) => ribbonShapes.icon(i) as HTMLElement, tip: (i) => ribbonShapes.label(i), pick: (i) => ribbonShapes.draw(i),
    current: () => undefined, redraw: () => "1", panel: galleryPanel,
  };
  const insertGroup = (variant: 0 | 1 | 2): Control[] => [
    variant < 2 ? shapesTiles : PD("shape.gallery", "SH", "large", galleryPanel),
    stack(D("shape.edit", "E", variant === 0 ? "medium" : "small", editMenu), B("insert.textBoxDraw", "X", variant === 0 ? "medium" : "small")),
  ];
  const styleTools = (size: "medium" | "small") => stack(SP("shape.fill", "SF", size, fillPanel), SP("shape.outline", "SO", size, outlinePanel), D("shape.effects", "SE", size, effectsMenu));
  const textTools = (size: "large" | "medium") => [D("shape.textDirection", "AD", size, directionMenu), D("shape.textAlign", "AT", size, alignMenu)];
  // Wide: icon and label; narrower: the icon only (Word).
  const sizeControls = (labels: boolean): Control[] => [stack(spinner("shape.height", "H", "h", labels ? "Höhe:" : ""), spinner("shape.width", "W", "w", labels ? "Breite:" : ""))];
  const layers = (size: "large" | "medium"): Control[] => [
    { type: "split", cmd: "picture.bringForward", key: "AF", size, menu: [{ cmd: "picture.bringForward" }, { cmd: "picture.bringToFront" }, { cmd: "picture.bringInFrontOfText" }] },
    { type: "split", cmd: "picture.sendBackward", key: "AE", size, menu: [{ cmd: "picture.sendBackward" }, { cmd: "picture.sendToBack" }, { cmd: "picture.sendBehindText" }] },
    { type: "button", cmd: "layout.selectionPane", key: "AP", size }];
  const arrangeTools = (size: "medium" | "small"): Control => stack(
    { type: "dropdown", cmd: "picture.align", key: "AA", size, menu: [], popup: (close) => alignObjects(close) },
    { type: "dropdown", cmd: "layout.group", key: "GG", size, menu: GROUP_MENU },
    { type: "dropdown", cmd: "picture.rotate", key: "AY", size, menu: [{ cmd: "picture.rotateRight" }, { cmd: "picture.rotateLeft" }, { cmd: "picture.flipV" }, { cmd: "picture.flipH" }, "-", { cmd: "picture.sizeDialog", label: "Weitere Drehungsoptionen..." }] });
  let alignObjects: (close: () => void) => HTMLElement = () => h("div");
  const pos: Control = { type: "dropdown", cmd: "picture.position", key: "PO", size: "large", menu: [], popup: (close) => positionPanel(close) };
  let positionPanel: (close: () => void) => HTMLElement = () => h("div");
  const wrap: Control = { type: "dropdown", cmd: "layout.wrapText", key: "TW", size: "large", menu: [{ cmd: "picture.wrapInline" }, "-", { cmd: "picture.wrapSquare" }, { cmd: "picture.wrapTight" }, { cmd: "picture.wrapThrough" }, { cmd: "picture.wrapTopBottom" }, "-", { cmd: "picture.wrapBehind" }, { cmd: "picture.wrapFront" }, "-", { cmd: "picture.moreLayout" }] };

  // Layouts and reduction order as Word 365 at window widths from 2560 down to 800 px
  // (capture_ribbon_widths.py Formformat).
  const tab: Tab = {
    id: "shapeFormat", label: "Formformat", key: "JD", contextual: true,
    shrink: ["text", "arrange", "styles", "size", "insert", "wordart", "wordart", "arrange", "insert", "size", "styles", "arrange", "text", "wordart"],
    groups: [
      { id: "insert", label: "Formen einfügen", key: "ZI", icon: "shapes", controls: insertGroup(0), smaller: [insertGroup(1), insertGroup(2)] },
      { id: "styles", label: "Formenarten", key: "ZS", icon: "shape-fill", launcher: { cmd: "shape.formatPane", key: "Ü" },
        controls: [styleTiles(7), styleTools("medium")], smaller: [[styleTiles(3), styleTools("medium")], [styleTiles(3), styleTools("small")]] },
      { id: "wordart", label: "WordArt-Formate", key: "ZW", icon: "wordart", launcher: { cmd: "text.effectsPane", key: "Ö" },
        controls: [PD("wordart.styles", "Q", "large", (close) => WORDART.styles(close), "Schnellformatvorlagen"), stack(WD("text.fill", "TI", "medium", () => WORDART.fill()), WD("text.outline", "TO", "medium", () => WORDART.outline()), WD("text.effects", "TX", "medium", () => WORDART.effects()))],
        // Word: icons at 1700 px, the large button at 1500 px, collapsed at 800 px.
        smaller: [[PD("wordart.styles", "Q", "large", (close) => WORDART.styles(close), "Schnellformatvorlagen"), stack(WD("text.fill", "TI", "small", () => WORDART.fill()), WD("text.outline", "TO", "small", () => WORDART.outline()), WD("text.effects", "TX", "small", () => WORDART.effects()))],
          [PD("wordart.styles", "Q", "large", (close) => WORDART.styles(close), "Schnellformatvorlagen"), stack(WD("text.fill", "TI", "small", () => WORDART.fill()), WD("text.outline", "TO", "small", () => WORDART.outline()), WD("text.effects", "TX", "small", () => WORDART.effects()))]] },
      { id: "text", label: "Text", key: "ZT", icon: "text-direction", controls: [...textTools("large"), B("shape.createLink", "C", "large")], smaller: [[stack(...textTools("medium"), B("shape.createLink", "C", "medium"))]] },
      { id: "access", label: "Barrierefreiheit", key: "ZB", icon: "alt-text", controls: [B("picture.altText", "AT", "large", "Alternativtext")] },
      { id: "arrange", label: "Anordnen", key: "ZA", icon: "wrap-text",
        controls: [pos, wrap, ...layers("large"), arrangeTools("medium")],
        smaller: [[pos, wrap, stack(...layers("medium")), arrangeTools("medium")], [pos, wrap, stack(...layers("medium")), arrangeTools("small")]] },
      { id: "size", label: "Größe", key: "ZG", icon: "shape-height", launcher: { cmd: "picture.sizeDialog", key: "ÖG" }, controls: sizeControls(true), smaller: [sizeControls(false)] },
    ],
  };
  return {
    tab,
    /** Panels shared with the Bildformat and Layout tabs (set by the app). */
    layoutCommands: {
      setPositionPanel: (fn: (close: () => void) => HTMLElement) => { positionPanel = fn; },
      setAlignPanel: (fn: (close: () => void) => HTMLElement) => { alignObjects = fn; },
      setChangePanel,
    },
  };
}

/** The fill colour a shape shows now (its own or its style's). */
function currentFillColor(pic: PictureFormat | undefined, s: ShapeInfo): string {
  const f = pic?.fill ?? s.styleFill;
  if (f?.kind === "solid") return f.color;
  if (f?.kind === "gradient") return f.stops[0]?.color ?? "4472C4";
  return "4472C4";
}

function mix(hex: string, k: number, toWhite: boolean): string {
  const v = [0, 2, 4].map((i) => parseInt(hex.slice(i, i + 2), 16));
  return v.map((x) => Math.round(toWhite ? x + (255 - x) * k : x * (1 - k * 0.5)).toString(16).padStart(2, "0")).join("").toUpperCase();
}

async function moreColor(title: string, done: (hex: string) => void): Promise<void> {
  const input = h("input", { type: "color", value: "#4472c4" });
  const ok = await showDialog({ title, body: h("div", { class: "grid2" }, "Farbe:", input), buttons: [{ label: "OK", value: true, primary: true }, { label: "Abbrechen", value: false }], cancel: false });
  if (ok) done(input.value.slice(1).toUpperCase());
}
