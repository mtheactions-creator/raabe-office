// The contextual tab "Bildformat" of German Excel 365 while a picture is selected (D-087), recorded with
// tools/feel-lab/probe_excel_pictures.py (labels and keytips from our own Excel's ribbon through UI
// Automation): Anpassen (In Zelle, Freistellen, Korrekturen, Farbe, Künstlerische Effekte, Transparenz,
// Bilder komprimieren..., Bild ändern, Bild zurücksetzen), Bildformatvorlagen (Schnellformatvorlagen,
// Bildrahmen, Bildeffekte, Bildlayout; launcher Bild...), Barrierefreiheit (Alternativtext), Anordnen
// (Ebene nach vorne / nach hinten, Auswahlbereich..., Ausrichten, Gruppieren, Drehen) and Größe
// (Zuschneiden, Formenhöhe, Formenbreite; launcher Größe und Eigenschaften...). The gallery entries and
// presets are Office's (recorded from Word, shared in text-engine; Excel's 28 picture styles write the
// same DrawingML, picture-styles-excel.test.ts). Also the task panes Grafik formatieren and Alternativtext.

import { closePopup, colorGrid, dashMenu, h, largeOf, pictures, showDialog, weightMenu, type Command, type Control, type MenuEntry, type PaneControl, type PaneSection, type PaneSpec, type Tab } from "@raabe/shell";
import {
  BEVELS, BRIGHT_CONTRAST, EFFECT_PRESETS, GLOWS, PICTURE_STYLES, REFLECTIONS, ROTATIONS_3D, SATURATION, SHADOWS, SHARPEN, SOFT_EDGES, TEMPERATURE, TRANSPARENCY,
  previewElement, recolorGallery, removeBackground, type GalleryItem, type PictureFormat, type PicturePreset,
} from "@raabe/text-engine";
import { DEFAULT_THEME_COLORS, pictureFromSpPr } from "@raabe/ooxml";
import type { CalcApp } from "./app";

const cmd = (id: string, label: string, icon?: string, extra: Partial<Command> = {}): Command => ({ id, label, icon, ...extra });

export const PICTURE_COMMANDS: Command[] = [
  // "In Zelle" (a picture inside a cell, Excel's IMAGE data type) is not built yet.
  cmd("pic.inCell", "In Zelle", "insert-picture", { deferred: true }),
  cmd("pic.removeBackground", "Freistellen", "remove-background"),
  cmd("pic.corrections", "Korrekturen", "corrections"),
  cmd("pic.color", "Farbe", "picture-color"),
  // Künstlerische Effekte are Office's own image filters; pictures that have one show it as stored.
  cmd("pic.artistic", "Künstlerische Effekte", "artistic", { deferred: true }),
  cmd("pic.transparency", "Transparenz", "transparency"),
  cmd("pic.compress", "Bilder komprimieren...", "compress"),
  cmd("pic.change", "Bild ändern", "change-picture"),
  cmd("pic.changeFile", "Dieses Gerät...", "open"),
  cmd("pic.changeClipboard", "Aus Zwischenablage...", "paste"),
  cmd("pic.reset", "Bild zurücksetzen", "reset-picture"),
  cmd("pic.resetAll", "Bild und Größe zurücksetzen", "reset-picture"),
  cmd("pic.styles", "Schnellformatvorlagen", "picture-styles"),
  cmd("pic.border", "Bildrahmen", "picture-border"),
  cmd("pic.effects", "Bildeffekte", "picture-effects"),
  // Bildlayout turns pictures into SmartArt, which RaabeCalc does not have.
  cmd("pic.layout", "Bildlayout", "picture-layout", { deferred: true }),
  cmd("pic.formatPane", "Bild...", "format-picture"),
  cmd("pic.altText", "Alternativtext", "alt-text"),
  cmd("pic.forward", "Ebene nach vorne", "bring-forward"),
  cmd("pic.toFront", "In den Vordergrund", "bring-forward"),
  cmd("pic.backward", "Ebene nach hinten", "send-backward"),
  cmd("pic.toBack", "In den Hintergrund", "send-backward"),
  cmd("pic.selectionPane", "Auswahlbereich...", "selection-pane"),
  cmd("pic.align", "Ausrichten", "align-objects"),
  cmd("pic.group", "Gruppieren", "group-objects"),
  cmd("pic.rotate", "Drehen", "rotate"),
  cmd("pic.rotateRight", "Rechtsdrehung 90 Grad", "rotate"),
  cmd("pic.rotateLeft", "Linksdrehung 90 Grad", "rotate"),
  cmd("pic.flipV", "Vertikal umdrehen", "flip-v"),
  cmd("pic.flipH", "Horizontal spiegeln", "flip-h"),
  cmd("pic.rotateMore", "Weitere Drehungsoptionen...", "rotate"),
  cmd("pic.crop", "Zuschneiden", "crop"),
  cmd("pic.cropFill", "Ausfüllen", "crop-fill"),
  cmd("pic.cropFit", "Einpassen", "crop-fit"),
  cmd("pic.height", "Formenhöhe", "shape-height"),
  cmd("pic.width", "Formenbreite", "shape-width"),
  cmd("pic.sizePane", "Größe und Eigenschaften...", "image-size"),
];

/** Keys of a picture style (frame shape, fill, line, effects, 3D): applying one replaces these. */
const STYLE_KEYS: (keyof PictureFormat)[] = ["geometry", "geometryAdj", "fillXml", "fill", "line", "shadow", "reflection", "glow", "softEdge", "bevel", "scene3dXml", "sp3dXml"];

/** Excel's names of the Korrekturen and Farbe entries where they differ from Word's ("4700 K", "-40 %", "Ausgeblichen"). */
export function excelLabel(label: string): string {
  return label.replace(/−/g, "-").replace(/(\d)\.(\d{3}) K/, "$1$2 K").replace("Schärfen und Weichzeichnen: 0 %", "Schärfen: 0 %").replace(/^Ausgewaschen$/, "Ausgeblichen");
}

const CM = 96 / 2.54;
const fmtCm = (px: number) => `${(Math.round((px / CM) * 100) / 100).toLocaleString("de-DE", { minimumFractionDigits: 2, maximumFractionDigits: 2 })} cm`;
const parseCm = (v: string) => Number(v.replace(/[^\d,.-]/g, "").replace(",", ".")) * CM;

export function pictureTab(app: CalcApp) {
  const c = app.commands;
  c.register(...PICTURE_COMMANDS);
  const pics = app.pictures;
  const sel = () => pics.selected();
  const on = () => !!sel() && app.protectionAllows("objects");
  const run = (id: string, fn: () => void | Promise<void>, enabled: () => boolean = on, checked?: () => boolean) =>
    c.implement(id, { run: () => { void fn(); }, enabled, checked });
  const f = (): PictureFormat => pics.format();
  const theme = (): Record<string, string> => ({
    ...DEFAULT_THEME_COLORS,
    ...Object.fromEntries(Object.entries(app.wb.theme ?? {}).map(([k, v]) => [k, String(v).replace("#", "").toUpperCase()])),
  });
  const apply = (label: string, item: GalleryItem) => pics.setFormat(label, item.format);

  // ----- galleries ------------------------------------------------------------------------------

  interface Section { title?: string; items: GalleryItem[]; columns: number; tile: [number, number] }

  const current = (fm: PictureFormat, item: GalleryItem) =>
    Object.entries(item.format).every(([k, v]) => JSON.stringify((fm as Record<string, unknown>)[k]) === JSON.stringify(v));

  /** A gallery popup: the selected picture with each entry; hovering previews it on the sheet, a click applies it. */
  function galleryPanel(label: string, sections: Section[], links: MenuEntry[], close: () => void): HTMLElement {
    const p = sel();
    const base = f();
    const panel = h("div", { class: "rpanel picture-gallery" });
    if (!p) return panel;
    const look = pics.lookOf(p, {});
    for (const s of sections) {
      if (s.title) panel.append(h("div", { class: "title" }, s.title));
      const grid = h("div", { class: "tiles", style: `display:grid;grid-template-columns:repeat(${s.columns}, auto);gap:2px;padding:2px 6px` });
      for (const item of s.items) {
        const [w, hh] = s.tile;
        const tile = h("button", {
          class: `tile pic-tile${current(base, item) ? " on" : ""}`, "data-tip": item.label, "aria-label": item.label,
          style: `width:${w + 8}px;height:${hh + 8}px;padding:3px`,
          onclick: () => { pics.preview(null); close(); apply(label, item); },
        });
        tile.append(previewElement(look.url, w, hh, { ...base, ...item.format, rotation: undefined }));
        tile.addEventListener("mouseenter", () => pics.preview({ ...base, ...item.format }));
        tile.addEventListener("mouseleave", () => pics.preview(null));
        grid.append(tile);
      }
      panel.append(grid);
    }
    const watch = new MutationObserver(() => { if (!panel.isConnected) { pics.preview(null); watch.disconnect(); } });
    setTimeout(() => watch.observe(document.body, { childList: true, subtree: true }));
    if (links.length) {
      panel.append(h("hr"));
      for (const l of links) {
        if (l === "-" || !("run" in l)) continue;
        panel.append(h("button", { class: "link", disabled: !!l.disabled, onclick: () => { pics.preview(null); closePopup(); l.run(); } }, l.label));
      }
    }
    return panel;
  }

  const relabel = (items: GalleryItem[]) => items.map((i) => ({ ...i, label: excelLabel(i.label) }));
  const tile: [number, number] = [90, 68];

  /** Neu einfärben with Excel's names: the theme colour's name, "Akzentfarbe 1 dunkel" … */
  function recolorItems(): GalleryItem[] {
    const t = theme();
    const items = recolorGallery(t).map((i) => ({ ...i, label: excelLabel(i.label) }));
    const names = colourNames();
    const role = (k: string) => k === "dk2" ? "Textfarbe 2" : k === "lt2" ? "Hintergrundfarbe 2" : `Akzentfarbe ${k.slice(6)}`;
    const keysDark = ["dk2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6"];
    const keysLight = ["lt2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6"];
    for (let i = 0; i < 7; i++) {
      const d = items[7 + i], l = items[14 + i];
      if (d) d.label = `${names[keysDark[i]] ? names[keysDark[i]] + ", " : ""}${role(keysDark[i])} dunkel`;
      if (l) l.label = `${names[keysLight[i]] ? names[keysLight[i]] + ", " : ""}${role(keysLight[i])} hell`;
    }
    return items;
  }

  /** The theme colours' names as the palette shows them ("Dunkelblaugrün, Akzent 1" gives "Dunkelblaugrün"). */
  function colourNames(): Record<string, string> {
    const out: Record<string, string> = {};
    const order = ["lt1", "dk1", "lt2", "dk2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6"];
    const grid = colorGrid(() => {}, theme())[0];
    const labels = Array.from(grid?.querySelectorAll("button") ?? []).slice(0, 10).map((b) => b.getAttribute("aria-label") ?? "");
    order.forEach((k, i) => { const m = /^([^,]+),/.exec(labels[i] ?? ""); if (m) out[k] = m[1]; });
    return out;
  }

  const corrections = (close: () => void) => galleryPanel("Korrekturen", [
    { title: "Schärfen/Weichzeichnen", items: relabel(SHARPEN), columns: 5, tile },
    { title: "Helligkeit/Kontrast", items: relabel(BRIGHT_CONTRAST), columns: 5, tile },
  ], [{ label: "Bildkorrekturoptionen...", run: () => pics.openPane("picture") }], close);
  const color = (close: () => void) => galleryPanel("Farbe", [
    { title: "Farbsättigung", items: relabel(SATURATION), columns: 7, tile },
    { title: "Farbton", items: relabel(TEMPERATURE), columns: 7, tile },
    { title: "Neu einfärben", items: recolorItems(), columns: 7, tile },
  ], [{ label: "Transparente Farbe bestimmen", run: () => pics.startPick() }, { label: "Bildfarboptionen...", run: () => pics.openPane("picture") }], close);
  const transparency = (close: () => void) => galleryPanel("Transparenz", [{ items: relabel(TRANSPARENCY), columns: 7, tile }],
    [{ label: "Optionen für die Bildtransparenz...", run: () => pics.openPane("picture") }], close);

  /** A preset's settings (the given keys only, the others of the picture stay). */
  const fromPreset = (p: PicturePreset, keys: (keyof PictureFormat)[], extra: PictureFormat = {}): PictureFormat => {
    const parsed = pictureFromSpPr(p.xml, theme()) as Record<string, unknown>;
    const out: Record<string, unknown> = {};
    for (const k of keys) out[k] = parsed[k];
    return { ...out, ...extra } as PictureFormat;
  };
  const presetItems = (list: readonly PicturePreset[], keys: (keyof PictureFormat)[], extra?: (i: number) => PictureFormat): GalleryItem[] =>
    list.map((p, i) => ({ label: p.label, format: fromPreset(p, keys, extra?.(i)) }));
  let styleCache: { theme: string; items: GalleryItem[] } | null = null;
  const styleItems = (): GalleryItem[] => {
    const t = JSON.stringify(theme());
    if (styleCache?.theme !== t) styleCache = { theme: t, items: presetItems(PICTURE_STYLES, STYLE_KEYS, (i) => ({ preset: i })) };
    return styleCache.items;
  };

  // ----- Anpassen -------------------------------------------------------------------------------

  run("pic.corrections", () => {});
  run("pic.color", () => {});
  run("pic.transparency", () => {});
  run("pic.removeBackground", () => pics.removeBg(removeBackground));
  run("pic.change", () => {});
  run("pic.changeFile", () => pics.changeFromFile());
  run("pic.changeClipboard", () => pics.changeFromClipboard(), () => on() && !!app.platform.clipboardReadImage);
  run("pic.reset", () => pics.reset(false));
  run("pic.resetAll", () => pics.reset(true));
  run("pic.compress", async () => {
    const only = h("input", { type: "checkbox", checked: true }) as HTMLInputElement;
    const cropped = h("input", { type: "checkbox", checked: true }) as HTMLInputElement;
    const opts: [string, number][] = [["HD (330 ppi): gute Qualität für High-Definition-Anzeigen (HD)", 330], ["Drucken (220 ppi): hervorragende Qualität auf den meisten Druckern und Bildschirmen", 220],
      ["Web (150 ppi): gut für Webseiten und Projektoren", 150], ["E-Mail (96 ppi): Dokumentgröße für die Freigabe minimieren", 96]];
    const radios = opts.map(([l, v], i) => ({ el: h("input", { type: "radio", name: "ppi", checked: i === 1 }) as HTMLInputElement, l, v }));
    const ok = await showDialog({
      title: "Bilder komprimieren",
      body: h("div", { style: "width:430px;display:flex;flex-direction:column;gap:6px" },
        h("b", null, "Komprimierungsoptionen:"),
        h("label", { class: "check" }, only, "Nur für dieses Bild übernehmen"),
        h("label", { class: "check" }, cropped, "Zugeschnittene Bildbereiche löschen"),
        h("b", { style: "margin-top:6px" }, "Auflösung:"),
        ...radios.map((r) => h("label", { class: "check" }, r.el, r.l))),
      buttons: [{ label: "OK", value: true, primary: true }, { label: "Abbrechen", value: false }], cancel: false,
    });
    if (!ok) return;
    await pics.compress(radios.find((r) => r.el.checked)?.v ?? 220, cropped.checked, only.checked);
  });

  // ----- Bildformatvorlagen ---------------------------------------------------------------------

  run("pic.styles", () => {});
  let hoverBase = false;
  const styleTiles: Control = {
    type: "tiles", cmd: "pic.styles", key: "S", count: PICTURE_STYLES.length, visible: 4, fill: true, max: 9, size: [48, 48],
    render: (i) => {
      const p = sel();
      if (!p) return null;
      return previewElement(pics.lookOf(p, {}).url, 44, 36, { ...f(), ...styleItems()[i].format, rotation: undefined });
    },
    tip: (i) => PICTURE_STYLES[i].label,
    pick: (i) => { pics.preview(null); hoverBase = false; apply("Bildformatvorlage", styleItems()[i]); },
    hover: (i) => {
      if (i === null) { if (hoverBase) { pics.preview(null); hoverBase = false; } return; }
      hoverBase = true;
      pics.preview({ ...f(), ...styleItems()[i].format });
    },
    current: () => f().preset,
    redraw: () => { const p = sel(); return p ? `${p.src.length}:${p.src.slice(-24)}` : ""; },
    panel: (close) => galleryPanel("Bildformatvorlage", [{ items: styleItems(), columns: 7, tile: [72, 54] }], [], close),
  };

  // Bildrahmen: colours (the palette), Keine Kontur, Weitere Konturfarben..., Stärke, Skizziert, Striche.
  const lineOf = () => f().line ?? undefined;
  const setLine = (patch: Partial<NonNullable<PictureFormat["line"]>> | null) => {
    if (patch === null) { pics.setFormat("Bildrahmen", { line: null }); return; }
    const cur = lineOf();
    pics.setFormat("Bildrahmen", { line: { color: cur?.color ?? "000000", width: cur?.width ?? 0.75, ...(cur ?? {}), ...patch } });
  };
  run("pic.border", () => {});
  const weights = [0.25, 0.5, 0.75, 1, 1.5, 2.25, 3, 4.5, 6];
  const moreLines: MenuEntry = { label: "Weitere Linien...", run: () => pics.openPane("fill") };
  const borderExtra = (): MenuEntry[] => [
    { label: "Keine Kontur", run: () => setLine(null) },
    { label: "Stärke", submenu: () => [...weightMenu(weights.map((w): [string, number] => [`${w.toLocaleString("de-DE")} Pt.`, w]), (w) => setLine({ width: w })), "-", moreLines] },
    { label: "Skizziert", submenu: [], disabled: true },
    { label: "Striche", submenu: () => [dashMenu(lineOf()?.dash ?? "solid", (v) => setLine({ dash: v === "solid" ? undefined : v })), "-", moreLines] },
  ];
  const border: Control = {
    type: "color", cmd: "pic.border", key: "V", palette: "font", showLabel: true, more: "Weitere Konturfarben...",
    current: () => (lineOf()?.color ?? "000000").replace("#", ""), pick: (hex) => setLine({ color: hex.replace("#", "").toUpperCase() }),
    extra: borderExtra, theme: () => app.designs.paletteTheme?.(),
  };

  // Bildeffekte: Office's presets in submenus.
  run("pic.effects", () => {});
  const fx = (noneLabel: string, list: readonly (PicturePreset & { readonly id: string })[], keys: (keyof PictureFormat)[], columns: number, sections?: [string, number][], more?: [string, string]) => (): HTMLElement => {
    const items = presetItems(list, keys);
    const none = items.find((_, i) => list[i].id === "none");
    const rest = items.filter((_, i) => list[i].id !== "none");
    const parts: Section[] = [];
    if (none) parts.push({ title: noneLabel, items: [none], columns: 1, tile: [48, 36] });
    if (sections) {
      let at = 0;
      for (const [t, n] of sections) { parts.push({ title: t, items: rest.slice(at, at + n), columns, tile: [48, 36] }); at += n; }
    } else parts.push({ items: rest, columns, tile: [48, 36] });
    return galleryPanel("Bildeffekte", parts, more ? [{ label: more[0], run: () => pics.openPane(more[1]) }] : [], () => closePopup());
  };
  const effectsMenu: MenuEntry[] = [
    { label: "Voreinstellung", icon: "picture-effects", flyout: fx("Keine Voreinstellungen", EFFECT_PRESETS, ["shadow", "glow", "reflection", "softEdge", "bevel", "scene3dXml", "sp3dXml"], 4, [["Voreinstellungen", 12]], ["3D-Optionen...", "effects"]) },
    { label: "Schatten", icon: "picture-effects", flyout: fx("Kein Schatten", SHADOWS, ["shadow"], 3, [["Außen", 9], ["Innen", 9], ["Perspektive", 5]], ["Schattenoptionen...", "effects"]) },
    { label: "Spiegelung", icon: "picture-effects", flyout: fx("Keine Spiegelung", REFLECTIONS, ["reflection"], 3, [["Spiegelungsvarianten", 9]], ["Spiegelungsoptionen...", "effects"]) },
    { label: "Leuchteffekt", icon: "picture-effects", flyout: fx("Kein Leuchteffekt", GLOWS, ["glow"], 6, [["Leuchteffektvarianten", 24]], ["Leuchteffektoptionen...", "effects"]) },
    { label: "Weiche Kanten", icon: "picture-effects", flyout: fx("Keine weichen Kanten", SOFT_EDGES, ["softEdge"], 3, [["Varianten weicher Kanten", 6]], ["Optionen für weiche Kanten...", "effects"]) },
    { label: "Abschrägung", icon: "picture-effects", flyout: fx("Keine Abschrägung", BEVELS, ["bevel", "sp3dXml", "scene3dXml"], 4, [["Abschrägung", 12]], ["3D-Optionen...", "effects"]) },
    { label: "3D-Drehung", icon: "picture-effects", flyout: fx("Keine Drehung", ROTATIONS_3D, ["scene3dXml"], 4, [["Parallel", 10], ["Perspektive", 11], ["Schräg", 4]], ["3D-Drehungsoptionen...", "effects"]) },
  ];
  run("pic.formatPane", () => pics.openPane());

  // ----- Barrierefreiheit -----------------------------------------------------------------------

  run("pic.altText", () => altTextPane());
  function altTextPane(): void {
    const p = sel();
    if (!p) return;
    document.documentElement.style.setProperty("--pane-top", `${document.querySelector(".calc-main")?.getBoundingClientRect().top ?? 150}px`);
    const render = (): HTMLElement => {
      const q = sel();
      const area = h("textarea", { rows: "6", style: "width:100%;box-sizing:border-box;font:inherit;resize:vertical" }) as HTMLTextAreaElement;
      area.value = q?.alt ?? "";
      area.addEventListener("change", () => { const v = area.value.trim(); if (sel() && v !== (sel()!.alt ?? "")) pics.edit("Alternativtext", (x) => ({ ...x, alt: v || undefined })); });
      area.addEventListener("keydown", (e) => e.stopPropagation());
      return h("div", { style: "display:flex;flex-direction:column;gap:6px" }, h("label", null, "Beschreibung (für Personen, die das Bild nicht sehen):"), area);
    };
    app.taskPane.show({ title: "Alternativtext", sets: [{ label: "", tabs: [{ id: "alt", label: "Alternativtext", icon: "alt-text", sections: () => [{ title: "", open: true, controls: () => [{ type: "custom", render }] }] }] }] });
  }

  // ----- Anordnen -------------------------------------------------------------------------------

  run("pic.forward", () => pics.reorder("Ebene nach vorne", "forward"));
  run("pic.toFront", () => pics.reorder("In den Vordergrund", "front"));
  run("pic.backward", () => pics.reorder("Ebene nach hinten", "backward"));
  run("pic.toBack", () => pics.reorder("In den Hintergrund", "back"));
  run("pic.selectionPane", () => app.shapes.selectionPane(), () => true);
  // The pictures in the Auswahlbereich, below the shapes (they are drawn under them), frontmost first.
  app.shapes.setPaneExtra({
    rows: () => [...(app.sheet.pictures ?? [])].filter((p) => !p.group).reverse().map((p) => ({
      name: p.name ?? "Grafik", hidden: !!p.hidden, on: p === sel(),
      pick: () => pics.select(p),
      toggle: () => { pics.change(p.hidden ? "Anzeigen" : "Ausblenden", (list) => list.map((q) => (q === p ? { ...q, hidden: q.hidden ? undefined : true } : q))); if (p === sel()) pics.deselect(); },
      rename: (v: string) => pics.change("Umbenennen", (list) => list.map((q) => (q === p ? { ...q, name: v } : q))),
    })),
    all: (hidden) => {
      if (!(app.sheet.pictures ?? []).some((q) => !!q.hidden !== hidden)) return;
      pics.change(hidden ? "Ausblenden" : "Anzeigen", (list) => list.map((q) => ({ ...q, hidden: hidden || undefined })));
      if (hidden) pics.deselect();
    },
  });
  run("pic.align", () => {});
  // Gruppieren needs two objects; RaabeCalc selects one picture at a time (Excel greys it then too).
  c.implement("pic.group", { run: () => {}, enabled: () => false });
  const alignMenu = (): MenuEntry[] => {
    const one = (label: string): MenuEntry => ({ label, disabled: true, run: () => {} });
    return [one("Linksbündig"), one("Horizontal zentrieren"), one("Rechtsbündig"), "-", one("Oben ausrichten"), one("Vertikal zentrieren"), one("Unten ausrichten"), "-",
      one("Horizontal verteilen"), one("Vertikal verteilen"), "-",
      { label: "Am Raster ausrichten", checked: pics.snapGrid(), run: () => pics.toggleSnap() },
      { label: "An Form ausrichten", disabled: true, run: () => {} },
      { label: "Gitternetzlinien anzeigen", checked: app.sheet.view.showGrid, run: () => void c.run("layout.gridlinesView") }];
  };
  const turn = (deg: number) => pics.setFormat("Drehen", (fm) => ({ ...fm, rotation: ((((fm.rotation ?? 0) + deg) % 360) + 360) % 360 || undefined }));
  run("pic.rotate", () => {});
  run("pic.rotateRight", () => rotateKeepingPlace(90));
  run("pic.rotateLeft", () => rotateKeepingPlace(-90));
  run("pic.flipV", () => pics.setFormat("Vertikal umdrehen", (fm) => ({ ...fm, flipV: fm.flipV ? undefined : true })));
  run("pic.flipH", () => pics.setFormat("Horizontal spiegeln", (fm) => ({ ...fm, flipH: fm.flipH ? undefined : true })));
  run("pic.rotateMore", () => pics.openPane("size"));
  /** Turning by 90°: the frame stays where it is (Excel turns around the middle; the anchor follows). */
  function rotateKeepingPlace(deg: number): void {
    const b = pics.box();
    if (!b) { turn(deg); return; }
    pics.edit("Drehen", (q) => {
      const fm = (q.format ?? {}) as PictureFormat;
      const r = ((((fm.rotation ?? 0) + deg) % 360) + 360) % 360;
      return pics.placeAt({ ...q, format: { ...fm, rotation: r || undefined } }, b, r);
    });
  }
  const rotateMenu: MenuEntry[] = [{ cmd: "pic.rotateRight" }, { cmd: "pic.rotateLeft" }, { cmd: "pic.flipV" }, { cmd: "pic.flipH" }, "-", { cmd: "pic.rotateMore" }];

  // ----- Größe ----------------------------------------------------------------------------------

  run("pic.crop", () => pics.startCrop(), on, () => pics.cropping());
  run("pic.cropFill", () => pics.cropFill(false));
  run("pic.cropFit", () => pics.cropFill(true));
  run("pic.sizePane", () => pics.openPane("size"));
  const ratios: [string, number][] = [["1:1", 1], ["2:3", 2 / 3], ["3:4", 3 / 4], ["3:5", 3 / 5], ["4:5", 4 / 5], ["3:2", 3 / 2], ["4:3", 4 / 3], ["5:3", 5 / 3], ["5:4", 5 / 4], ["16:9", 16 / 9], ["16:10", 16 / 10]];
  const SHAPES: [string, string][] = [["rect", "Rechteck"], ["roundRect", "Rechteck: abgerundete Ecken"], ["snip1Rect", "Rechteck: eine Ecke abgeschnitten"],
    ["snip2DiagRect", "Rechteck: diagonal liegende Ecken abgeschnitten"], ["round2DiagRect", "Rechteck: diagonal liegende Ecken abgerundet"], ["ellipse", "Ellipse"],
    ["triangle", "Gleichschenkliges Dreieck"], ["rtTriangle", "Rechtwinkliges Dreieck"], ["parallelogram", "Parallelogramm"], ["trapezoid", "Trapez"], ["diamond", "Raute"],
    ["pentagon", "Fünfeck"], ["hexagon", "Sechseck"], ["octagon", "Achteck"], ["star5", "Stern: 5 Zacken"], ["heart", "Herz"], ["cloud", "Wolke"]];
  const cropMenu: MenuEntry[] = [{ cmd: "pic.crop" },
    { label: "Auf Form zuschneiden", submenu: () => SHAPES.map(([g, l]) => ({ label: l, checked: (f().geometry ?? "rect") === g, run: () => pics.setFormat("Auf Form zuschneiden", { geometry: g === "rect" ? undefined : g, geometryAdj: undefined }) })) },
    { label: "Seitenverhältnis", submenu: () => [{ heading: "Quadrat" }, ...ratios.slice(0, 1).map(([l, r]) => ({ label: l, run: () => pics.cropToRatio(r) })),
      { heading: "Hochformat" }, ...ratios.slice(1, 5).map(([l, r]) => ({ label: l, run: () => pics.cropToRatio(r) })),
      { heading: "Querformat" }, ...ratios.slice(5).map(([l, r]) => ({ label: l, run: () => pics.cropToRatio(r) }))] },
    "-", { cmd: "pic.cropFill" }, { cmd: "pic.cropFit" }];

  /** Höhe / Breite: the other side follows while Seitenverhältnis sperren is on. */
  function setSide(side: "w" | "h", px: number): void {
    const p = sel();
    const b = pics.box();
    if (!p || !b || !(px > 1)) return;
    const z = app.sheet.view.zoom / 100;
    const v = px * z;
    const w = side === "w" ? v : p.lockAspect ? (b.w * v) / b.h : b.w;
    const hh = side === "h" ? v : p.lockAspect ? (b.h * v) / b.w : b.h;
    pics.edit("Größe ändern", (q) => pics.placeAt(q, { x: b.x, y: b.y, w, h: hh }));
  }
  const sizeOf = () => { const b = pics.box(); const z = app.sheet.view.zoom / 100; return b ? { w: b.w / z, h: b.h / z } : null; };
  const spinner = (id: string, label: string, key: string, side: "w" | "h"): Control => ({
    type: "spinner", cmd: id, label, key, icon: side === "h" ? "shape-height" : "shape-width",
    value: () => { const s = sizeOf(); return s ? fmtCm(side === "w" ? s.w : s.h) : ""; },
    commit: (v) => setSide(side, parseCm(v)),
    step: (d) => { const s = sizeOf(); if (s) setSide(side, (side === "w" ? s.w : s.h) + d * 0.1 * CM); },
  });
  run("pic.height", () => {});
  run("pic.width", () => {});

  // ----- Grafik formatieren ---------------------------------------------------------------------

  pics.setPane(() => paneSpec());
  function paneSpec(): PaneSpec {
    const fm = () => f();
    const lineSec: PaneSection = { title: "Linie", open: true, controls: () => {
      const l = fm().line;
      const c2: PaneControl[] = [
        { type: "radio", label: "Keine Linie", checked: () => !l, pick: () => setLine(null) },
        { type: "radio", label: "Einfarbige Linie", checked: () => !!l, pick: () => setLine({}) },
      ];
      if (l) {
        c2.push({ type: "color", label: "Farbe", value: () => `#${l.color}`, set: (hex) => setLine({ color: (hex ?? "#000000").replace("#", "").toUpperCase() }) });
        c2.push({ type: "number", label: "Transparenz", unit: " %", min: 0, max: 100, slider: true, value: () => Math.round((l.transparency ?? 0) * 100), set: (v) => setLine({ transparency: v / 100 || undefined }) });
        c2.push({ type: "number", label: "Breite", unit: " Pt.", min: 0, max: 1584, step: 0.25, decimals: 2, value: () => l.width, set: (v) => setLine({ width: v }) });
        c2.push({ type: "picture", label: "Strichtyp", options: pictures.dashEntries(), columns: 1, size: pictures.PIC_SIZE.dash, value: () => l.dash ?? "solid", set: (v) => setLine({ dash: v === "solid" ? undefined : v }) });
      }
      return c2;
    } };
    const sh = () => fm().shadow;
    const pt = (label: string, value: () => number | null, set: (v: number) => void, max = 100, unit = " Pt."): PaneControl => ({ type: "number", label, unit, min: 0, max, step: 0.5, decimals: 1, value, set });
    const shadowSec: PaneSection = { title: "Schatten", open: !!sh(), controls: () => {
      const s = sh();
      const setS = (patch: Partial<NonNullable<PictureFormat["shadow"]>> | null) => pics.setFormat("Schatten", { shadow: patch === null ? undefined : { kind: "outer", color: "000000", transparency: 0.6, blur: 4, dist: 3, dir: 45, ...(s ?? {}), ...patch } as PictureFormat["shadow"] });
      const c2: PaneControl[] = [{ type: "check", label: "Schatten", value: () => !!s, set: (v) => setS(v ? {} : null) }];
      if (s) {
        c2.push({ type: "color", label: "Farbe", value: () => `#${s.color}`, set: (hex) => setS({ color: (hex ?? "#000000").replace("#", "").toUpperCase() }) });
        c2.push({ type: "number", label: "Transparenz", unit: " %", min: 0, max: 100, slider: true, value: () => Math.round(s.transparency * 100), set: (v) => setS({ transparency: v / 100 }) });
        c2.push(pt("Weichzeichnen", () => s.blur, (v) => setS({ blur: v })));
        c2.push({ type: "number", label: "Winkel", unit: "°", min: 0, max: 359, value: () => Math.round(s.dir), set: (v) => setS({ dir: v }) });
        c2.push(pt("Abstand", () => s.dist, (v) => setS({ dist: v })));
      }
      return c2;
    } };
    const reflSec: PaneSection = { title: "Spiegelung", open: !!fm().reflection, controls: () => {
      const r = fm().reflection;
      const setR = (patch: Partial<NonNullable<PictureFormat["reflection"]>> | null) => pics.setFormat("Spiegelung", { reflection: patch === null ? undefined : { size: 0.55, dist: 0, transparency: 0.5, blur: 0.5, endA: 0.003, ...(r ?? {}), ...patch } });
      const c2: PaneControl[] = [{ type: "check", label: "Spiegelung", value: () => !!r, set: (v) => setR(v ? {} : null) }];
      if (r) {
        c2.push({ type: "number", label: "Transparenz", unit: " %", min: 0, max: 100, slider: true, value: () => Math.round(r.transparency * 100), set: (v) => setR({ transparency: v / 100 }) });
        c2.push({ type: "number", label: "Größe", unit: " %", min: 0, max: 100, slider: true, value: () => Math.round(r.size * 100), set: (v) => setR({ size: v / 100 }) });
        c2.push(pt("Weichzeichnen", () => r.blur, (v) => setR({ blur: v })));
        c2.push(pt("Abstand", () => r.dist, (v) => setR({ dist: v })));
      }
      return c2;
    } };
    const glowSec: PaneSection = { title: "Leuchten", open: !!fm().glow, controls: () => {
      const g = fm().glow;
      const setG = (patch: Partial<NonNullable<PictureFormat["glow"]>> | null) => pics.setFormat("Leuchteffekt", { glow: patch === null ? undefined : { color: (theme().accent1 ?? "4472C4"), radius: 5, transparency: 0.6, ...(g ?? {}), ...patch } });
      const c2: PaneControl[] = [{ type: "check", label: "Leuchten", value: () => !!g, set: (v) => setG(v ? {} : null) }];
      if (g) {
        c2.push({ type: "color", label: "Farbe", value: () => `#${g.color}`, set: (hex) => setG({ color: (hex ?? "#000000").replace("#", "").toUpperCase() }) });
        c2.push(pt("Größe", () => g.radius, (v) => setG({ radius: v }), 150));
        c2.push({ type: "number", label: "Transparenz", unit: " %", min: 0, max: 100, slider: true, value: () => Math.round(g.transparency * 100), set: (v) => setG({ transparency: v / 100 }) });
      }
      return c2;
    } };
    const softSec: PaneSection = { title: "Weiche Kanten", open: !!fm().softEdge, controls: () => [
      pt("Größe", () => fm().softEdge ?? 0, (v) => pics.setFormat("Weiche Kanten", { softEdge: v || undefined })),
    ] };
    const sizeSec: PaneSection = { title: "Größe", open: true, controls: () => {
      const s = sizeOf();
      return [
        { type: "number", label: "Höhe", unit: " cm", min: 0, max: 558, step: 0.1, decimals: 2, value: () => (s ? Math.round((s.h / CM) * 100) / 100 : 0), set: (v) => setSide("h", v * CM) },
        { type: "number", label: "Breite", unit: " cm", min: 0, max: 558, step: 0.1, decimals: 2, value: () => (s ? Math.round((s.w / CM) * 100) / 100 : 0), set: (v) => setSide("w", v * CM) },
        { type: "number", label: "Drehung", unit: "°", min: -3600, max: 3600, value: () => fm().rotation ?? 0, set: (v) => {
          const r = ((Math.round(v) % 360) + 360) % 360;
          const b = pics.box();
          pics.edit("Drehen", (q) => { const nf = { ...((q.format ?? {}) as PictureFormat), rotation: r || undefined }; return b ? pics.placeAt({ ...q, format: nf }, b, r) : { ...q, format: nf }; });
        } },
        { type: "check", label: "Seitenverhältnis sperren", value: () => !!sel()?.lockAspect, set: (v) => pics.edit("Seitenverhältnis sperren", (q) => ({ ...q, lockAspect: v || undefined })) },
        { type: "button", label: "Zurücksetzen", run: () => void pics.reset(true) },
      ];
    } };
    const altSec: PaneSection = { title: "Alternativtext", open: false, controls: () => [
      { type: "text", label: "Titel", value: () => sel()?.title ?? "", set: (v) => pics.edit("Alternativtext", (q) => ({ ...q, title: v.trim() || undefined })) },
      { type: "text", label: "Beschreibung", value: () => sel()?.alt ?? "", set: (v) => pics.edit("Alternativtext", (q) => ({ ...q, alt: v.trim() || undefined })) },
    ] };
    const pctCtl = (label: string, value: () => number, set: (v: number) => void, min = -100, max = 100): PaneControl =>
      ({ type: "number", label, unit: " %", min, max, slider: true, value: () => Math.round(value() * 100), set: (v) => set(v / 100) });
    const corrSec: PaneSection = { title: "Bildkorrekturen", open: true, controls: () => [
      pctCtl("Schärfe", () => fm().sharpen ?? 0, (v) => pics.setFormat("Korrekturen", { sharpen: v || undefined })),
      pctCtl("Helligkeit", () => fm().brightness ?? 0, (v) => pics.setFormat("Korrekturen", { brightness: v || undefined })),
      pctCtl("Kontrast", () => fm().contrast ?? 0, (v) => pics.setFormat("Korrekturen", { contrast: v || undefined })),
      { type: "button", label: "Zurücksetzen", run: () => pics.setFormat("Korrekturen", { sharpen: undefined, brightness: undefined, contrast: undefined }) },
    ] };
    const colorSec: PaneSection = { title: "Bildfarbe", open: true, controls: () => [
      pctCtl("Sättigung", () => fm().saturation ?? 1, (v) => pics.setFormat("Farbe", { saturation: v === 1 ? undefined : v }), 0, 400),
      { type: "number", label: "Temperatur", min: 1500, max: 11500, step: 100, value: () => fm().temperature ?? 6500, set: (v) => pics.setFormat("Farbe", { temperature: v === 6500 ? undefined : v }) },
      { type: "button", label: "Zurücksetzen", run: () => pics.setFormat("Farbe", { saturation: undefined, temperature: undefined, recolor: undefined }) },
    ] };
    const transSec: PaneSection = { title: "Bildtransparenz", open: true, controls: () => [
      pctCtl("Transparenz", () => fm().transparency ?? 0, (v) => pics.setFormat("Transparenz", { transparency: v || undefined }), 0, 100),
    ] };
    const cropSec: PaneSection = { title: "Zuschneiden", open: true, controls: () => {
      const k = fm().crop ?? { l: 0, t: 0, r: 0, b: 0 };
      const setK = (patch: Partial<typeof k>) => {
        const b = pics.box();
        const nk = { ...k, ...patch };
        if (!b) return;
        // The frame shrinks or grows with the crop; the picture keeps its scale.
        const full = { w: b.w / Math.max(0.01, 1 - k.l - k.r), h: b.h / Math.max(0.01, 1 - k.t - k.b) };
        const frame = { x: b.x + (nk.l - k.l) * full.w, y: b.y + (nk.t - k.t) * full.h, w: full.w * (1 - nk.l - nk.r), h: full.h * (1 - nk.t - nk.b) };
        if (frame.w < 2 || frame.h < 2) return;
        pics.edit("Zuschneiden", (q) => pics.placeAt({ ...q, format: { ...((q.format ?? {}) as PictureFormat), crop: nk.l || nk.t || nk.r || nk.b ? nk : undefined } }, frame));
      };
      return [
        { type: "heading", label: "Zuschneideposition" },
        pctCtl("Links", () => k.l, (v) => setK({ l: v }), -100, 99), pctCtl("Oben", () => k.t, (v) => setK({ t: v }), -100, 99),
        pctCtl("Rechts", () => k.r, (v) => setK({ r: v }), -100, 99), pctCtl("Unten", () => k.b, (v) => setK({ b: v }), -100, 99),
      ];
    } };
    const tabs = [
      { id: "fill", label: "Füllung und Linie", icon: "pane-fill-line", sections: () => [lineSec] },
      { id: "effects", label: "Effekte", icon: "pane-effects", sections: () => [shadowSec, reflSec, glowSec, softSec] },
      { id: "size", label: "Größe und Eigenschaften", icon: "size-position", sections: () => [sizeSec, altSec] },
      { id: "picture", label: "Bild", icon: "format-picture", sections: () => [corrSec, colorSec, transSec, cropSec] },
    ];
    return { title: "Grafik formatieren", sets: [{ label: "Bildoptionen", tabs }] };
  }

  // ----- the tab --------------------------------------------------------------------------------

  const arrange: Control[] = [
    { type: "stack", items: [
      { type: "split", cmd: "pic.forward", key: "GE", size: "medium", menu: [{ cmd: "pic.forward" }, { cmd: "pic.toFront" }] },
      { type: "split", cmd: "pic.backward", key: "GH", size: "medium", menu: [{ cmd: "pic.backward" }, { cmd: "pic.toBack" }] },
      { type: "button", cmd: "pic.selectionPane", key: "Q2", size: "medium" }] },
    { type: "stack", items: [
      { type: "dropdown", cmd: "pic.align", key: "DA", size: "medium", get menu() { return alignMenu(); } },
      { type: "dropdown", cmd: "pic.group", key: "GR", size: "medium", menu: [{ cmd: "shape.group" }, { cmd: "shape.ungroup" }, { cmd: "shape.regroup" }] },
      { type: "dropdown", cmd: "pic.rotate", key: "DR", size: "medium", menu: rotateMenu }] }];
  const big = (id: string, key: string, popup: (close: () => void) => HTMLElement): Control => ({ type: "dropdown", cmd: id, key, size: "large", menu: [], popup });
  const tab: Tab = {
    id: "pictureFormat", label: "Bildformat", key: "JP", contextual: true, shrink: [],
    groups: [
      { id: "adjust", label: "Anpassen", key: "ZT", icon: "corrections", sepAfter: [0, 1], controls: [
        { type: "button", cmd: "pic.inCell", key: "Y", size: "large" },
        { type: "button", cmd: "pic.removeBackground", key: "T1", size: "large" },
        big("pic.corrections", "O", corrections), big("pic.color", "F", color),
        { type: "dropdown", cmd: "pic.artistic", key: "Ü", size: "large", menu: [] },
        big("pic.transparency", "T2", transparency),
        { type: "stack", items: [
          { type: "button", cmd: "pic.compress", key: "L", size: "medium" },
          { type: "dropdown", cmd: "pic.change", key: "Q1", size: "medium", menu: [{ cmd: "pic.changeFile" }, { cmd: "pic.changeClipboard" }] },
          { type: "split", cmd: "pic.reset", key: "U", size: "medium", menu: [{ cmd: "pic.reset" }, { cmd: "pic.resetAll" }] }] }] },
      { id: "styles", label: "Bildformatvorlagen", key: "ZS", icon: "picture-styles", launcher: { cmd: "pic.formatPane", key: "P" }, controls: [
        styleTiles,
        { type: "stack", items: [border, { type: "dropdown", cmd: "pic.effects", key: "E", size: "medium", menu: effectsMenu },
          { type: "dropdown", cmd: "pic.layout", key: "A1", size: "medium", menu: [] }] }] },
      { id: "access", label: "Barrierefreiheit", key: "Z1", icon: "alt-text", controls: [{ type: "button", cmd: "pic.altText", key: "AT", size: "large" }] },
      { id: "arrange", label: "Anordnen", key: "ZA", icon: "bring-forward", controls: largeOf(arrange, { "pic.selectionPane": "Auswahlbereich" }), smaller: [arrange] },
      { id: "size", label: "Größe", key: "ZZ", icon: "image-size", launcher: { cmd: "pic.sizePane", key: "ß" }, controls: [
        { type: "split", cmd: "pic.crop", key: "C", size: "large", menu: cropMenu },
        { type: "stack", items: [spinner("pic.height", "Höhe:", "H", "h"), spinner("pic.width", "Breite:", "RR", "w")] }] },
    ],
  };
  return { tab };
}
