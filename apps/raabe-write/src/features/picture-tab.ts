// The contextual tab "Bildformat" as Word 365 has it (D-045): Anpassen (Freistellen,
// Korrekturen, Farbe, Transparenz, Bilder komprimieren, Bild ändern, Bild zurücksetzen),
// Bildformatvorlagen (styles, Bildrahmen, Bildeffekte), Barrierefreiheit (Alternativtext),
// Anordnen (Position, Textumbruch, Ebenen, Ausrichten, Drehen) and Größe (Zuschneiden, Höhe,
// Breite). Gallery entries and presets were recorded from Word (picture-presets.ts).

import { closePopup, colorGrid, h, showDialog, showMenu, type Command, type Control, type MenuEntry, type Tab } from "@raabe/shell";
import { imageInfo, imageUrl, naturalSizePt, previewElement, type PictureFormat } from "@raabe/text-engine";
import { DEFAULT_THEME_COLORS, pictureFromSpPr } from "@raabe/ooxml";
import { CM, fmtCm, parseNum, type AppApi } from "./api";
import { GROUP_MENU } from "./groups";
import {
  BRIGHT_CONTRAST, SATURATION, SHARPEN, TEMPERATURE, TRANSPARENCY, galleryPanel, withItem, type GalleryItem, type GallerySection,
} from "./picture-format";
import * as P from "./picture-presets";

const cmd = (id: string, label: string, icon?: string, extra: Partial<Command> = {}): Command => ({ id, label, icon, ...extra });

export const PICTURE_TAB_COMMANDS: Command[] = [
  cmd("picture.removeBackground", "Freistellen", "remove-background", { tip: "Freistellen\nUnerwünschte Teile des Bilds automatisch entfernen." }),
  cmd("picture.corrections", "Korrekturen", "corrections", { tip: "Korrekturen\nHelligkeit, Kontrast oder Schärfe des Bilds verbessern." }),
  cmd("picture.color", "Farbe", "picture-color", { tip: "Farbe\nDie Farbe ändern, um die Qualität zu verbessern oder den Inhalt anzupassen." }),
  cmd("picture.artistic", "Künstlerische Effekte", "artistic", { deferred: true }),
  cmd("picture.transparency", "Transparenz", "transparency", { tip: "Transparenz\nDas Bild transparenter machen." }),
  cmd("picture.compress", "Bilder komprimieren", "compress", { tip: "Bilder komprimieren\nBilder im Dokument komprimieren, um ihre Größe zu verringern." }),
  cmd("picture.change", "Bild ändern", "change-picture"),
  cmd("picture.reset", "Bild zurücksetzen", "reset-picture"),
  cmd("picture.resetAll", "Bild und Größe zurücksetzen", "reset-picture"),
  cmd("picture.border", "Bildrahmen", "picture-border"),
  cmd("picture.effects", "Bildeffekte", "picture-effects"),
  cmd("picture.layout", "Bildlayout", "picture-layout", { deferred: true }),
  cmd("picture.position", "Position", "position"),
  cmd("picture.bringForward", "Ebene nach vorne", "bring-forward"),
  cmd("picture.bringToFront", "In den Vordergrund", "bring-forward"),
  cmd("picture.bringInFrontOfText", "Vor den Text bringen", "bring-forward"),
  cmd("picture.sendBackward", "Ebene nach hinten", "send-backward"),
  cmd("picture.sendToBack", "In den Hintergrund", "send-backward"),
  cmd("picture.sendBehindText", "Hinter den Text bringen", "send-backward"),
  cmd("picture.align", "Ausrichten", "align-objects"),
  cmd("picture.rotate", "Drehen", "rotate"),
  cmd("picture.rotateRight", "Rechtsdrehung 90°", "rotate"),
  cmd("picture.rotateLeft", "Linksdrehung 90°", "rotate"),
  cmd("picture.flipV", "Vertikal kippen", "flip-v"),
  cmd("picture.flipH", "Horizontal kippen", "flip-h"),
  cmd("picture.cropMenu", "Zuschneiden", "crop"),
  cmd("picture.cropFill", "Ausfüllen", "crop-fill"),
  cmd("picture.cropFit", "Einpassen", "crop-fit"),
  cmd("picture.sizeDialog", "Layout", undefined),
];

/** A picture with an entry's settings, for previews: the entry's keys replace the picture's. */
const fromPreset = (p: P.PicturePreset, theme: Record<string, string>, keys: (keyof PictureFormat)[], extra: PictureFormat = {}): PictureFormat => {
  const f = pictureFromSpPr(p.xml, theme) as Record<string, unknown>;
  const out: Record<string, unknown> = {};
  for (const k of keys) out[k] = f[k];
  if (p.extent) out.extent = { l: p.extent[0], t: p.extent[1], r: p.extent[2], b: p.extent[3] };
  return { ...out, ...extra } as PictureFormat;
};

const STYLE_KEYS: (keyof PictureFormat)[] = ["geometry", "geometryAdj", "fillXml", "line", "shadow", "reflection", "glow", "softEdge", "bevel", "scene3dXml", "sp3dXml"];

export function pictureTab(api: AppApi, insertFromPath: (path: string) => Promise<void>) {
  const c = api.commands;
  const ed = () => api.editor();
  const sel = () => ed()?.selectedImage() ?? null;
  const picked = () => !!sel();
  const isFloat = () => sel()?.kind === "float";
  const theme = () => ({ ...DEFAULT_THEME_COLORS, ...(ed()?.doc.styles.theme.colors ?? {}) });
  const run = (id: string, fn: () => void | Promise<void>, enabled: () => boolean = picked, checked?: () => boolean) =>
    c.implement(id, { run: async () => { await fn(); api.focus(); }, enabled, checked });
  const set = (f: PictureFormat) => { ed()?.setPictureFormat(f); api.refresh(); };
  const apply = (item: GalleryItem) => set(item.format);
  const gallery = (sections: GallerySection[], links: MenuEntry[] = []) => (close: () => void) => {
    const e = ed();
    return e ? galleryPanel(e, sections, links, close, apply) : h("div");
  };

  // ----- Anpassen ----------------------------------------------------------------------
  const presetItems = (list: readonly P.PicturePreset[], keys: (keyof PictureFormat)[], extra?: (i: number) => PictureFormat): GalleryItem[] =>
    list.map((p, i) => ({ label: p.label, format: fromPreset(p, theme(), keys, extra?.(i)) }));
  const recolorItems = (): GalleryItem[] => {
    const t = theme();
    const names = ["Text 2", "Akzent 1", "Akzent 2", "Akzent 3", "Akzent 4", "Akzent 5", "Akzent 6"];
    const keys = ["dk2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6"];
    const lightBg = ["lt2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6"];
    const mix = (hex: string, k: number, toWhite: boolean) => {
      const v = [0, 2, 4].map((i) => parseInt(hex.slice(i, i + 2), 16));
      const out = v.map((x) => Math.round(toWhite ? x + (255 - x) * (1 - k) : x * k));
      return out.map((x) => x.toString(16).padStart(2, "0")).join("").toUpperCase();
    };
    return [
      { label: "Keine Neueinfärbung", format: { recolor: undefined, lum: undefined } },
      { label: "Graustufe", format: { recolor: { kind: "grayscale" } } },
      { label: "Sepia", format: { recolor: { kind: "sepia" } } },
      { label: "Ausgewaschen", format: { recolor: { kind: "washout" } } },
      { label: "Schwarzweiß: 25 %", format: { recolor: { kind: "bw", threshold: 0.25 } } },
      { label: "Schwarzweiß: 50 %", format: { recolor: { kind: "bw", threshold: 0.5 } } },
      { label: "Schwarzweiß: 75 %", format: { recolor: { kind: "bw", threshold: 0.75 } } },
      ...keys.map((k, i) => ({ label: `${names[i]}, dunkle Farbe`, format: { recolor: { kind: "duotone" as const, dark: "000000", light: mix(t[k] ?? "4472C4", 0.45, true) } } })),
      ...lightBg.map((k, i) => ({ label: `${names[i] === "Text 2" ? "Hintergrund 2" : names[i]}, helle Farbe`, format: { recolor: { kind: "duotone" as const, dark: mix(t[k] ?? "4472C4", 0.45, false), light: "FFFFFF" } } })),
    ];
  };
  const tile: [number, number] = [90, 68];
  const corrections = gallery([
    { title: "Schärfen/Weichzeichnen", items: SHARPEN, columns: 5, tile },
    { title: "Helligkeit/Kontrast", items: BRIGHT_CONTRAST, columns: 5, tile },
  ], [{ label: "Bildkorrekturoptionen...", run: () => void c.run("picture.formatPane") }]);
  const color = gallery([
    { title: "Farbsättigung", items: SATURATION, columns: 7, tile },
    { title: "Farbton", items: TEMPERATURE, columns: 7, tile },
    { title: "Neu einfärben", items: recolorItems(), columns: 7, tile },
  ], [{ label: "Bildfarboptionen...", run: () => void c.run("picture.formatPane") }]);
  const transparency = gallery([{ items: TRANSPARENCY, columns: 7, tile }], [{ label: "Optionen für die Bildtransparenz...", run: () => void c.run("picture.formatPane") }]);
  c.implement("picture.corrections", { run: () => {}, enabled: picked });
  c.implement("picture.color", { run: () => {}, enabled: picked });
  c.implement("picture.transparency", { run: () => {}, enabled: picked });

  // Freistellen: the background (what touches the edges in about the edge colour) goes.
  run("picture.removeBackground", async () => {
    const s = sel();
    const img = s?.obj.image;
    if (!s || !img) return;
    const bytes = await removeBackground(img.bytes, img.type);
    if (bytes) ed()?.replaceImageData({ bytes, type: "image/png" });
  });

  // Bild ändern, Bild zurücksetzen.
  run("picture.change", async () => {
    const path = await api.platform.openImageDialog();
    if (!path) return;
    const bytes = await api.platform.readFile(path);
    const info = imageInfo(bytes);
    if (!info) return;
    const size = naturalSizePt(info);
    ed()?.replacePicture({ bytes, type: info.type }, size.w, size.h);
  });
  run("picture.reset", () => ed()?.resetPicture());
  run("picture.resetAll", () => {
    const img = sel()?.obj.image;
    const info = img ? imageInfo(img.bytes) : null;
    ed()?.resetPicture(info ? naturalSizePt(info) : undefined);
  });
  c.implement("picture.changeFromFile", { run: () => c.run("picture.change"), enabled: picked });
  c.implement("picture.changeFromClipboard", {
    run: async () => {
      const clip = await api.clipboard();
      if (!clip.image) return;
      const bytes = new Uint8Array(await clip.image.arrayBuffer());
      const info = imageInfo(bytes);
      if (info) { const sz = naturalSizePt(info); ed()?.replacePicture({ bytes, type: info.type }, sz.w, sz.h); }
    }, enabled: picked,
  });

  // Bilder komprimieren.
  run("picture.compress", async () => {
    const only = h("input", { type: "checkbox", checked: true });
    const cropped = h("input", { type: "checkbox", checked: true });
    const opts: [string, number][] = [["HD (330 ppi): gute Qualität für High-Definition-Anzeigen (HD)", 330], ["Drucken (220 ppi): hervorragende Qualität auf den meisten Druckern und Bildschirmen", 220],
      ["Web (150 ppi): gut für Webseiten und Projektoren", 150], ["E-Mail (96 ppi): Dokumentgröße für die Freigabe minimieren", 96]];
    const radios = opts.map(([l, v], i) => ({ el: h("input", { type: "radio", name: "ppi", checked: i === 1 }), l, v }));
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
    const ppi = radios.find((r) => r.el.checked)?.v ?? 220;
    await ed()?.compressPictures(ppi, cropped.checked, only.checked);
  });

  // ----- Bildformatvorlagen ------------------------------------------------------------
  const styles = gallery([{ items: presetItems(P.PICTURE_STYLES, STYLE_KEYS, (i) => ({ preset: i })), columns: 5, tile: [72, 54] }]);
  c.implement("picture.styles", { run: () => {}, enabled: picked });

  // Bildrahmen: colours, Stärke, Striche.
  const lineOf = () => ed()?.pictureFormat()?.line ?? undefined;
  const setLine = (patch: Partial<NonNullable<PictureFormat["line"]>> | null) => {
    const cur = lineOf();
    if (patch === null) { set({ line: null }); return; }
    set({ line: { color: cur?.color ?? "000000", width: cur?.width ?? 0.75, ...patch } });
  };
  const border = (close: () => void) => {
    const colors = colorGrid((hex) => { close(); setLine({ color: hex }); api.focus(); });
    const link = (label: string, fn: () => void) => h("button", { class: "link", onclick: fn }, label);
    const weights = [0.25, 0.5, 0.75, 1, 1.5, 2.25, 3, 4.5, 6];
    const dashes: [string, string][] = [["solid", "Durchgezogen"], ["sysDot", "Rundpunkt"], ["sysDash", "Quadratpunkt"], ["dash", "Strich"], ["dashDot", "Strich-Punkt"], ["lgDash", "Langer Strich"], ["lgDashDot", "Lang-Strich-Punkt"], ["lgDashDotDot", "Lang-Strich-Punkt-Punkt"]];
    const sub = (ev: MouseEvent, entries: MenuEntry[]) => {
      const r = (ev.currentTarget as HTMLElement).getBoundingClientRect();
      close();
      showMenu(c, entries, new DOMRect(r.right, r.top, 0, 0));
    };
    return h("div", { class: "rpanel palette" }, ...colors,
      h("hr"),
      link("Keine Kontur", () => { close(); setLine(null); api.focus(); }),
      link("Weitere Konturfarben...", () => { close(); void moreColor((hex) => setLine({ color: hex })); }),
      h("button", { class: "link has-sub", onclick: (ev: MouseEvent) => sub(ev, weights.map((w) => ({ label: `${w.toLocaleString("de-DE")} Pt.`, checked: lineOf()?.width === w, run: () => setLine({ width: w }) }))) }, "Stärke"),
      h("button", { class: "link has-sub", onclick: (ev: MouseEvent) => sub(ev, dashes.map(([v, l]) => ({ label: l, checked: (lineOf()?.dash ?? "solid") === v, run: () => setLine({ dash: v === "solid" ? undefined : v }) }))) }, "Striche"));
  };
  const moreColor = async (done: (hex: string) => void) => {
    const input = h("input", { type: "color", value: `#${lineOf()?.color ?? "000000"}` });
    const ok = await showDialog({ title: "Farben", body: h("div", { class: "grid2" }, "Farbe:", input), buttons: [{ label: "OK", value: true, primary: true }, { label: "Abbrechen", value: false }], cancel: false });
    if (ok) done(input.value.slice(1).toUpperCase());
  };
  c.implement("picture.border", { run: () => setLine({}), enabled: picked });

  // Bildeffekte: submenus with Word's presets.
  const fx = (title: string, noneLabel: string, list: readonly (P.PicturePreset & { readonly id: string })[], keys: (keyof PictureFormat)[], columns: number, sections?: [string, number][]) => (): HTMLElement => {
    const e = ed();
    if (!e) return h("div");
    const items = presetItems(list, keys);
    const none = items.find((_, i) => list[i].id === "none");
    const rest = items.filter((_, i) => list[i].id !== "none");
    const parts: GallerySection[] = [];
    if (none) parts.push({ title: noneLabel, items: [none], columns: 1, tile: [48, 36] });
    if (sections) {
      let at = 0;
      for (const [t, n] of sections) { parts.push({ title: t, items: rest.slice(at, at + n), columns, tile: [48, 36] }); at += n; }
    } else parts.push({ title, items: rest, columns, tile: [48, 36] });
    return galleryPanel(e, parts, [], () => closePopup(), apply);
  };
  const effectsMenu: MenuEntry[] = [
    { label: "Voreinstellung", icon: "picture-effects", flyout: fx("Voreinstellungen", "Keine Voreinstellungen", P.EFFECT_PRESETS, ["shadow", "glow", "reflection", "softEdge", "bevel", "scene3dXml", "sp3dXml"], 4) },
    { label: "Schatten", icon: "picture-effects", flyout: fx("Schatten", "Kein Schatten", P.SHADOWS, ["shadow"], 3, [["Außen", 9], ["Innen", 9], ["Perspektive", 5]]) },
    { label: "Spiegelung", icon: "picture-effects", flyout: fx("Spiegelungsvarianten", "Keine Spiegelung", P.REFLECTIONS, ["reflection"], 3) },
    { label: "Leuchteffekt", icon: "picture-effects", flyout: fx("Leuchteffektvarianten", "Kein Leuchteffekt", P.GLOWS, ["glow"], 6) },
    { label: "Weiche Kanten", icon: "picture-effects", flyout: fx("Varianten weicher Kanten", "Keine weichen Kanten", P.SOFT_EDGES, ["softEdge"], 3) },
    { label: "Abschrägung", icon: "picture-effects", flyout: fx("Abschrägung", "Keine Abschrägung", P.BEVELS, ["bevel", "sp3dXml", "scene3dXml"], 4) },
    { label: "3D-Drehung", icon: "picture-effects", flyout: fx("3D-Drehung", "Keine Drehung", P.ROTATIONS_3D, ["scene3dXml"], 4, [["Parallel", 10], ["Perspektive", 11], ["Schräg", 4]]) },
  ];
  c.implement("picture.effects", { run: () => {}, enabled: picked });

  // ----- Barrierefreiheit --------------------------------------------------------------
  run("picture.altText", async () => {
    const e = ed();
    const cur = e?.pictureAlt();
    if (!e || !cur) return;
    const text = h("textarea", { rows: 5, style: "width:100%;box-sizing:border-box", placeholder: "Beschreiben Sie dieses Objekt für Personen mit Sehbehinderungen." });
    text.value = cur.descr ?? "";
    const deco = h("input", { type: "checkbox", checked: !!cur.decorative });
    const ok = await showDialog({
      title: "Alternativtext",
      body: h("div", { style: "width:360px;display:flex;flex-direction:column;gap:8px" },
        h("div", null, "Wie würden Sie dieses Objekt und seinen Kontext jemandem beschreiben, der es nicht sehen kann?"), text,
        h("label", { class: "check" }, deco, "Als dekorativ markieren")),
      buttons: [{ label: "OK", value: true, primary: true }, { label: "Abbrechen", value: false }], cancel: false,
    });
    if (ok) e.setPictureAlt({ descr: text.value.trim() || undefined, decorative: deco.checked || undefined, title: cur.title });
  });

  // ----- Anordnen ----------------------------------------------------------------------
  const positions: [string, "left" | "center" | "right", "top" | "center" | "bottom"][] = [
    ["Oben links", "left", "top"], ["Oben Mitte", "center", "top"], ["Oben rechts", "right", "top"],
    ["Mitte links", "left", "center"], ["Mitte zentriert", "center", "center"], ["Mitte rechts", "right", "center"],
    ["Unten links", "left", "bottom"], ["Unten Mitte", "center", "bottom"], ["Unten rechts", "right", "bottom"],
  ];
  const positionPanel = (close: () => void) => {
    const btn = (label: string, fn: () => void, x?: string, y?: string) => {
      const b = h("button", { class: "tile pos-tile", "data-tip": label, onclick: () => { close(); fn(); api.focus(); } });
      b.append(pagePreview(x, y));
      return b;
    };
    return h("div", { class: "rpanel" },
      h("div", { class: "title" }, "Mit Text in Zeile"), h("div", { class: "tiles" }, btn("Mit Text in Zeile", () => ed()?.setImageWrap("inline"))),
      h("div", { class: "title" }, "Mit Textumbruch"),
      h("div", { class: "tiles" }, ...positions.map(([l, x, y]) => btn(`${l} mit quadratischem Textumbruch`, () => ed()?.positionPicture(x, y), x, y))),
      h("hr"), h("button", { class: "link", onclick: () => { close(); void c.run("picture.moreLayout"); } }, "Weitere Layoutoptionen..."));
  };
  c.implement("picture.position", { run: () => {}, enabled: () => picked() && ed()!.doc.activeStory === "main" });

  const layer = (dir: "forward" | "front" | "frontText" | "backward" | "back" | "behindText") => () => ed()?.pictureLayer(dir);
  run("picture.bringForward", layer("forward"), isFloat);
  run("picture.bringToFront", layer("front"), isFloat);
  run("picture.bringInFrontOfText", layer("frontText"), isFloat);
  run("picture.sendBackward", layer("backward"), isFloat);
  run("picture.sendToBack", layer("back"), isFloat);
  run("picture.sendBehindText", layer("behindText"), isFloat);

  let alignToPage = false;
  const alignMenu = (): MenuEntry[] => {
    const f = !isFloat();
    const a = (label: string, x?: "left" | "center" | "right", y?: "top" | "center" | "bottom"): MenuEntry =>
      ({ label, disabled: f, run: () => { ed()?.alignPicture(x, y, alignToPage); api.focus(); } });
    return [a("Linksbündig", "left"), a("Horizontal zentrieren", "center"), a("Rechtsbündig", "right"), "-",
      a("Oben ausrichten", undefined, "top"), a("Vertikal zentrieren", undefined, "center"), a("Unten ausrichten", undefined, "bottom"), "-",
      { label: "Horizontal verteilen", disabled: true, run: () => {} }, { label: "Vertikal verteilen", disabled: true, run: () => {} }, "-",
      { label: "An Seite ausrichten", checked: alignToPage, run: () => { alignToPage = true; } },
      { label: "Am Seitenrand ausrichten", checked: !alignToPage, run: () => { alignToPage = false; } },
      { label: "Ausgewählte Objekte ausrichten", disabled: true, run: () => {} }];
  };
  c.implement("picture.align", { run: () => {}, enabled: picked });
  run("picture.rotateRight", () => ed()?.setPictureFormat((f) => ({ ...f, rotation: (((f.rotation ?? 0) + 90) % 360) || undefined })));
  run("picture.rotateLeft", () => ed()?.setPictureFormat((f) => ({ ...f, rotation: (((f.rotation ?? 0) + 270) % 360) || undefined })));
  run("picture.flipV", () => ed()?.setPictureFormat((f) => ({ ...f, flipV: f.flipV ? undefined : true })));
  run("picture.flipH", () => ed()?.setPictureFormat((f) => ({ ...f, flipH: f.flipH ? undefined : true })));
  c.implement("picture.rotate", { run: () => {}, enabled: picked });
  const rotateMenu: MenuEntry[] = [{ cmd: "picture.rotateRight" }, { cmd: "picture.rotateLeft" }, { cmd: "picture.flipV" }, { cmd: "picture.flipH" }, "-",
    { cmd: "picture.sizeDialog", label: "Weitere Drehungsoptionen..." }];

  // ----- Größe -------------------------------------------------------------------------
  run("picture.crop", () => ed()?.startCrop());
  c.implement("picture.cropMenu", { run: () => ed()?.startCrop(), enabled: picked });
  const ratios: [string, number][] = [["1:1", 1], ["2:3", 2 / 3], ["3:4", 3 / 4], ["3:5", 3 / 5], ["4:5", 4 / 5], ["3:2", 3 / 2], ["4:3", 4 / 3], ["5:3", 5 / 3], ["5:4", 5 / 4], ["16:9", 16 / 9], ["16:10", 16 / 10]];
  const SHAPES: [string, string][] = [["rect", "Rechteck"], ["roundRect", "Rechteck: abgerundete Ecken"], ["snip1Rect", "Rechteck: eine Ecke abgeschnitten"], ["snip2DiagRect", "Rechteck: diagonal liegende Ecken abgeschnitten"],
    ["round2DiagRect", "Rechteck: diagonal liegende Ecken abgerundet"], ["ellipse", "Ellipse"], ["triangle", "Gleichschenkliges Dreieck"], ["rtTriangle", "Rechtwinkliges Dreieck"],
    ["parallelogram", "Parallelogramm"], ["trapezoid", "Trapez"], ["diamond", "Raute"], ["pentagon", "Fünfeck"], ["hexagon", "Sechseck"], ["octagon", "Achteck"],
    ["star5", "Stern: 5 Zacken"], ["heart", "Herz"], ["cloud", "Wolke"], ["flowChartPunchedCard", "Flussdiagramm: Lochkarte"]];
  const cropMenu: MenuEntry[] = [{ cmd: "picture.crop", label: "Zuschneiden" },
    { label: "Auf Form zuschneiden", submenu: () => SHAPES.map(([g, l]) => ({ label: l, checked: (ed()?.pictureFormat()?.geometry ?? "rect") === g, run: () => set({ geometry: g === "rect" ? undefined : g, geometryAdj: undefined }) })) },
    { label: "Seitenverhältnis", submenu: () => [{ heading: "Quadrat" }, ...ratios.slice(0, 1).map(([l, r]) => ({ label: l, run: () => ed()?.cropToRatio(r) })),
      { heading: "Hochformat" }, ...ratios.slice(1, 5).map(([l, r]) => ({ label: l, run: () => ed()?.cropToRatio(r) })),
      { heading: "Querformat" }, ...ratios.slice(5).map(([l, r]) => ({ label: l, run: () => ed()?.cropToRatio(r) }))] },
    "-", { cmd: "picture.cropFill" }, { cmd: "picture.cropFit" }];
  run("picture.cropFill", () => ed()?.cropFill(false));
  run("picture.cropFit", () => ed()?.cropFill(true));

  const size = () => { const s = sel(); return s ? { w: s.obj.width ?? 0, h: s.obj.height ?? 0 } : null; };
  const setSide = (side: "w" | "h", pt: number) => {
    const s = size();
    if (!s || !(pt > 1) || !s.w || !s.h) return;
    const k = side === "w" ? pt / s.w : pt / s.h;
    ed()?.setImageSize(s.w * k, s.h * k);
  };
  const spinner = (id: string, label: string, key: string, side: "w" | "h"): Control => ({
    type: "spinner", cmd: id, label, key,
    value: () => { const s = size(); return s ? fmtCm(side === "w" ? s.w : s.h) : ""; },
    commit: (v) => setSide(side, parseNum(v) * CM),
    step: (d) => { const s = size(); if (s) setSide(side, (side === "w" ? s.w : s.h) + d * 0.1 * CM); },
  });

  // Layout dialog (Größe, Position, Textumbruch): "Weitere Layoutoptionen...".
  const layoutDialog = async (tabName: "size" | "position" | "wrap") => {
    const e = ed();
    const s = sel();
    if (!e || !s) return;
    const f = e.pictureFormat() ?? {};
    const img = s.obj.image;
    const info = img ? imageInfo(img.bytes) : null;
    const orig = info ? naturalSizePt(info) : null;
    const w = h("input", { value: fmtCm(s.obj.width ?? 0), style: "width:90px" });
    const hh = h("input", { value: fmtCm(s.obj.height ?? 0), style: "width:90px" });
    const rot = h("input", { type: "number", value: String(Math.round(f.rotation ?? 0)), style: "width:70px" });
    const lock = h("input", { type: "checkbox", checked: true });
    const sx = h("input", { value: orig ? `${Math.round(((s.obj.width ?? 0) / orig.w) * 100)} %` : "100 %", style: "width:70px" });
    const sy = h("input", { value: orig ? `${Math.round(((s.obj.height ?? 0) / orig.h) * 100)} %` : "100 %", style: "width:70px" });
    const body = h("div", { style: "width:380px;display:flex;flex-direction:column;gap:8px" },
      h("b", null, "Höhe"), h("div", { class: "grid2" }, "Absolut:", hh),
      h("b", null, "Breite"), h("div", { class: "grid2" }, "Absolut:", w),
      h("b", null, "Drehen"), h("div", { class: "grid2" }, "Drehung:", rot),
      h("b", null, "Skalieren"), h("div", { class: "grid2" }, "Höhe:", sy, "Breite:", sx),
      h("label", { class: "check" }, lock, "Seitenverhältnis sperren"),
      orig ? h("div", null, `Originalgröße: Höhe ${fmtCm(orig.h)}  Breite ${fmtCm(orig.w)}`) : null);
    void tabName;
    const ok = await showDialog({ title: "Layout", body, buttons: [{ label: "OK", value: true, primary: true }, { label: "Abbrechen", value: false }], cancel: false });
    if (!ok) return;
    const nw = parseNum(w.value) * CM, nh = parseNum(hh.value) * CM;
    if (nw > 1 && nh > 1) e.setImageSize(nw, nh);
    const r = Number(rot.value) || 0;
    e.setPictureFormat({ rotation: ((r % 360) + 360) % 360 || undefined });
  };
  c.implement("picture.sizeDialog", { run: () => void layoutDialog("size"), enabled: picked });
  c.implement("picture.moreLayout", { run: () => void layoutDialog("position"), enabled: picked });
  c.implement("picture.sizePosition", { run: () => void layoutDialog("size"), enabled: picked });

  // Grafik formatieren: the settings as numbers (Word's pane, as a dialog).
  c.implement("picture.formatPane", { run: () => void formatDialog(api), enabled: picked });

  const wrapMenu: MenuEntry[] = [
    { cmd: "picture.wrapInline" }, "-", { cmd: "picture.wrapSquare" }, { cmd: "picture.wrapTight" }, { cmd: "picture.wrapThrough" }, { cmd: "picture.wrapTopBottom" }, "-",
    { cmd: "picture.wrapBehind" }, { cmd: "picture.wrapFront" }, "-",
    { label: "Rahmenpunkte bearbeiten", disabled: true, run: () => {} }, "-",
    { label: "Mit Text verschieben", checked: true, disabled: !isFloat(), run: () => ed()?.setPictureMoveWithText(true) },
    { label: "Position auf Seite fixieren", disabled: !isFloat(), run: () => ed()?.setPictureMoveWithText(false) }, "-",
    { cmd: "picture.moreLayout" }, { cmd: "picture.setDefaultLayout" }];
  c.implement("picture.setDefaultLayout", { run: () => { const s = sel(); if (s) api.setDefaultPictureWrap?.(s.kind === "inline" ? "inline" : s.obj.wrap); }, enabled: picked });

  const S = (cmdId: string, key: string, menu: MenuEntry[]): Control => ({ type: "dropdown", cmd: cmdId, key, size: "small", menu });

  // The styles gallery in the ribbon: the selected picture in each style, live preview on hover.
  const styleItems = () => presetItems(P.PICTURE_STYLES, STYLE_KEYS, (i) => ({ preset: i }));
  let hoverBase: PictureFormat | null = null;
  const styleTiles: Control = {
    type: "tiles", cmd: "picture.styles", key: "K", count: P.PICTURE_STYLES.length, visible: 4, fill: true, size: [64, 48],
    render: (i) => {
      const img = sel()?.obj.image;
      if (!img) return null;
      const items = styleItems();
      return previewElement(imageUrl(img), 56, 42, withItem(hoverBase ?? ed()?.pictureFormat() ?? {}, items[i]));
    },
    tip: (i) => P.PICTURE_STYLES[i].label,
    pick: (i) => {
      // Back to the picture before the preview, so Rückgängig returns there.
      if (hoverBase) { ed()?.setPictureFormat(() => hoverBase!, false); hoverBase = null; }
      apply(styleItems()[i]);
      api.focus();
    },
    hover: (i) => {
      const e = ed();
      if (!e) return;
      if (i === null) { if (hoverBase) { e.setPictureFormat(() => hoverBase!, false); hoverBase = null; } return; }
      hoverBase ??= e.pictureFormat() ?? {};
      e.setPictureFormat(withItem(hoverBase, styleItems()[i]), false);
    },
    current: () => ed()?.pictureFormat()?.preset,
    redraw: () => { const img = sel()?.obj.image; return img ? `${img.bytes.length}:${img.bytes[100] ?? 0}` : ""; },
    panel: styles,
  };

  const big = (cmdId: string, key: string, popup: (close: () => void) => HTMLElement, label?: string): Control => ({ type: "dropdown", cmd: cmdId, key, size: "large", menu: [], popup, label });
  const mid = (cmdId: string, key: string, popup: (close: () => void) => HTMLElement, size: "medium" | "small" = "medium"): Control => ({ type: "dropdown", cmd: cmdId, key, size, menu: [], popup });
  const artistic = (size: "large" | "medium"): Control => ({ type: "dropdown", cmd: "picture.artistic", key: "F", size, menu: [] });
  const tools = (size: "medium" | "small"): Control => ({ type: "stack", items: [
    { type: "button", cmd: "picture.compress", key: "M", size, label: size === "medium" ? "Bilder komprimieren" : undefined },
    { type: "dropdown", cmd: "picture.change", key: "G", size, menu: [{ cmd: "picture.changeFromFile", label: "Aus einer Datei..." }, { cmd: "picture.changeFromClipboard", label: "Aus Zwischenablage" }] },
    { type: "split", cmd: "picture.reset", key: "Q", size, menu: [{ cmd: "picture.reset" }, { cmd: "picture.resetAll" }] }] });
  const styleTools = (size: "medium" | "small"): Control => ({ type: "stack", items: [
    { type: "split", cmd: "picture.border", key: "BR", size, menu: [], popup: border },
    { type: "dropdown", cmd: "picture.effects", key: "BE", size, menu: effectsMenu },
    { type: "dropdown", cmd: "picture.layout", key: "BL", size, menu: [] }] });
  const layers = (size: "large" | "medium"): Control[] => [
    { type: "split", cmd: "picture.bringForward", key: "AF", size, menu: [{ cmd: "picture.bringForward" }, { cmd: "picture.bringToFront" }, { cmd: "picture.bringInFrontOfText" }] },
    { type: "split", cmd: "picture.sendBackward", key: "AE", size, menu: [{ cmd: "picture.sendBackward" }, { cmd: "picture.sendToBack" }, { cmd: "picture.sendBehindText" }] },
    { type: "button", cmd: "layout.selectionPane", key: "AP", size }];
  const arrangeTools = (size: "medium" | "small"): Control => ({ type: "stack", items: [
    { type: "dropdown", cmd: "picture.align", key: "AA", size, menu: [], popup: (close) => menuPanel(alignMenu(), close) },
    { type: "dropdown", cmd: "layout.group", key: "GG", size, menu: GROUP_MENU },
    S("picture.rotate", "AY", rotateMenu)] });
  const sizeControls: Control[] = [{ type: "split", cmd: "picture.crop", key: "C", size: "large", menu: cropMenu },
    { type: "stack", items: [spinner("picture.height", "Höhe:", "H", "h"), spinner("picture.width", "Breite:", "W", "w")] }];
  const pos = big("picture.position", "PO", positionPanel);
  const wrap: Control = { type: "dropdown", cmd: "layout.wrapText", key: "TW", size: "large", menu: wrapMenu };

  // Layouts and reduction order as Word 365 at window widths from 2560 down to 800 px
  // (capture_ribbon_widths.py Bildformat).
  const tab: Tab = {
    id: "pictureFormat", label: "Bildformat", key: "JP", contextual: true,
    shrink: ["adjust", "arrange", "size", "adjust", "styles", "arrange", "size", "adjust"],
    groups: [
      { id: "adjust", label: "Anpassen", key: "ZP", icon: "corrections",
        controls: [{ type: "button", cmd: "picture.removeBackground", key: "E", size: "large" }, big("picture.corrections", "R", corrections),
          big("picture.color", "I", color), artistic("large"), big("picture.transparency", "T", transparency), tools("medium")],
        smaller: [
          [{ type: "button", cmd: "picture.removeBackground", key: "E", size: "large" }, big("picture.corrections", "R", corrections),
            { type: "stack", items: [mid("picture.color", "I", color), artistic("medium"), mid("picture.transparency", "T", transparency)] }, tools("medium")],
          [{ type: "button", cmd: "picture.removeBackground", key: "E", size: "large" }, big("picture.corrections", "R", corrections),
            { type: "stack", items: [mid("picture.color", "I", color), artistic("medium"), mid("picture.transparency", "T", transparency)] }, tools("small")],
        ] },
      { id: "styles", label: "Bildformatvorlagen", key: "ZS", icon: "picture-styles", launcher: { cmd: "picture.formatPane", key: "Ü" },
        controls: [styleTiles, styleTools("medium")],
        smaller: [[big("picture.styles", "K", styles, "Schnellformatvorlagen"), styleTools("small")]] },
      { id: "access", label: "Barrierefreiheit", key: "ZB", icon: "alt-text", controls: [{ type: "button", cmd: "picture.altText", key: "AT", size: "large", label: "Alternativtext" }] },
      { id: "arrange", label: "Anordnen", key: "ZA", icon: "wrap-text",
        controls: [pos, wrap, ...layers("large"), arrangeTools("medium")],
        smaller: [[pos, wrap, { type: "stack", items: layers("medium") }, arrangeTools("small")]] },
      { id: "size", label: "Größe", key: "ZG", icon: "image-size", launcher: { cmd: "picture.sizeDialog", key: "ÖG" },
        controls: sizeControls, smaller: [sizeControls] },
    ],
  };
  // Shared with the Formformat and Layout tabs (Anordnen).
  const arrange = {
    positionPanel, alignPanel: (close: () => void) => menuPanel(alignMenu(), close), rotateMenu, wrapMenu,
  };
  return { tab, arrange };

  /** A menu shown inside a dropdown popup (entries computed when it opens). */
  function menuPanel(entries: MenuEntry[], close: () => void): HTMLElement {
    const panel = h("div", { class: "rpanel wide" });
    for (const en of entries) {
      if (en === "-") { panel.append(h("hr")); continue; }
      if (!("run" in en)) continue;
      panel.append(h("button", { class: `link${en.checked ? " checked" : ""}`, disabled: !!en.disabled, onclick: () => { close(); en.run(); } }, (en.checked ? "✓ " : "") + en.label));
    }
    return panel;
  }
  void insertFromPath;
}

/** A small page preview for the Position gallery: lines of text and the picture's place. */
function pagePreview(x?: string, y?: string): HTMLElement {
  const box = h("div", { style: "position:relative;width:34px;height:44px;border:1px solid #9e9e9e;background:#fff" });
  for (let i = 0; i < 7; i++) box.append(h("div", { style: `position:absolute;left:4px;right:4px;top:${4 + i * 5.5}px;height:1px;background:#bdbdbd` }));
  const left = x === "left" ? 4 : x === "center" ? 11 : x === "right" ? 18 : 6;
  const top = y === "top" ? 4 : y === "center" ? 16 : y === "bottom" ? 28 : 16;
  box.append(h("div", { style: `position:absolute;left:${left}px;top:${top}px;width:12px;height:12px;background:#5b9bd5` }));
  return box;
}

/** Freistellen: pixels connected to the edges and close to the edge colour become transparent. */
async function removeBackground(bytes: Uint8Array, type: string): Promise<Uint8Array | null> {
  const url = URL.createObjectURL(new Blob([bytes as BlobPart], { type }));
  try {
    const img = await new Promise<HTMLImageElement>((resolve, reject) => { const i = new Image(); i.onload = () => resolve(i); i.onerror = reject; i.src = url; });
    const w = img.naturalWidth, hgt = img.naturalHeight;
    const cv = document.createElement("canvas");
    cv.width = w; cv.height = hgt;
    const g = cv.getContext("2d")!;
    g.drawImage(img, 0, 0);
    const data = g.getImageData(0, 0, w, hgt);
    const px = data.data;
    // The background colour: the average of the border pixels.
    let r = 0, gg = 0, b = 0, n = 0;
    for (let x = 0; x < w; x++) for (const y of [0, hgt - 1]) { const i = (y * w + x) * 4; r += px[i]; gg += px[i + 1]; b += px[i + 2]; n++; }
    for (let y = 0; y < hgt; y++) for (const x of [0, w - 1]) { const i = (y * w + x) * 4; r += px[i]; gg += px[i + 1]; b += px[i + 2]; n++; }
    r /= n; gg /= n; b /= n;
    const near = (i: number) => Math.abs(px[i] - r) + Math.abs(px[i + 1] - gg) + Math.abs(px[i + 2] - b) < 90;
    const seen = new Uint8Array(w * hgt);
    const stack: number[] = [];
    for (let x = 0; x < w; x++) { stack.push(x, (hgt - 1) * w + x); }
    for (let y = 0; y < hgt; y++) { stack.push(y * w, y * w + w - 1); }
    while (stack.length) {
      const p = stack.pop()!;
      if (seen[p]) continue;
      seen[p] = 1;
      if (!near(p * 4)) continue;
      px[p * 4 + 3] = 0;
      const x = p % w, y = (p / w) | 0;
      if (x > 0) stack.push(p - 1);
      if (x < w - 1) stack.push(p + 1);
      if (y > 0) stack.push(p - w);
      if (y < hgt - 1) stack.push(p + w);
    }
    g.putImageData(data, 0, 0);
    const blob = await new Promise<Blob | null>((resolve) => cv.toBlob(resolve, "image/png"));
    return blob ? new Uint8Array(await blob.arrayBuffer()) : null;
  } catch {
    return null;
  } finally {
    URL.revokeObjectURL(url);
  }
}

/** Grafik formatieren: Word's pane settings as numbers in a dialog. */
async function formatDialog(api: AppApi): Promise<void> {
  const e = api.editor();
  const f = e?.pictureFormat();
  if (!e || !f) return;
  const pctField = (v: number | undefined, def = 0) => h("input", { type: "number", value: String(Math.round(((v ?? def) * 100))), style: "width:70px" });
  const sharpen = pctField(f.sharpen), bright = pctField(f.brightness), contrast = pctField(f.contrast);
  const sat = pctField(f.saturation, 1), temp = h("input", { type: "number", value: String(f.temperature ?? 6500), style: "width:80px" });
  const transp = pctField(f.transparency);
  const lineW = h("input", { value: f.line ? `${f.line.width.toLocaleString("de-DE")} Pt.` : "", style: "width:80px" });
  const lineC = h("input", { type: "color", value: `#${f.line?.color ?? "000000"}` });
  const noLine = h("input", { type: "checkbox", checked: !f.line });
  const soft = h("input", { type: "number", value: String(f.softEdge ?? 0), style: "width:70px" });
  const crop = f.crop ?? { l: 0, t: 0, r: 0, b: 0 };
  const cl = pctField(crop.l), ct = pctField(crop.t), cr = pctField(crop.r), cb = pctField(crop.b);
  const ok = await showDialog({
    title: "Grafik formatieren",
    body: h("div", { style: "width:420px;display:flex;flex-direction:column;gap:6px" },
      h("b", null, "Linie"), h("label", { class: "check" }, noLine, "Keine Linie"), h("div", { class: "grid2" }, "Farbe:", lineC, "Breite:", lineW),
      h("b", null, "Weiche Kanten"), h("div", { class: "grid2" }, "Größe (Pt.):", soft),
      h("b", null, "Bildkorrekturen"), h("div", { class: "grid2" }, "Schärfe (%):", sharpen, "Helligkeit (%):", bright, "Kontrast (%):", contrast),
      h("b", null, "Bildfarbe"), h("div", { class: "grid2" }, "Sättigung (%):", sat, "Temperatur (K):", temp),
      h("b", null, "Bildtransparenz"), h("div", { class: "grid2" }, "Transparenz (%):", transp),
      h("b", null, "Zuschneiden (% des Bilds)"), h("div", { class: "grid2" }, "Links:", cl, "Oben:", ct, "Rechts:", cr, "Unten:", cb)),
    buttons: [{ label: "OK", value: true, primary: true }, { label: "Abbrechen", value: false }], cancel: false,
  });
  if (!ok) return;
  const n = (el: HTMLInputElement) => (Number(el.value) || 0) / 100;
  const width = parseNum(lineW.value);
  const cropNext = { l: n(cl), t: n(ct), r: n(cr), b: n(cb) };
  e.setPictureFormat({
    sharpen: n(sharpen) || undefined, brightness: n(bright) || undefined, contrast: n(contrast) || undefined,
    saturation: Math.abs(n(sat) - 1) < 1e-6 ? undefined : n(sat), temperature: Number(temp.value) === 6500 ? undefined : Number(temp.value) || undefined,
    transparency: n(transp) || undefined, softEdge: Number(soft.value) || undefined,
    line: noLine.checked ? (f.line ? null : f.line) : { ...(f.line ?? {}), color: lineC.value.slice(1).toUpperCase(), width: width > 0 ? width : f.line?.width ?? 0.75 },
    crop: cropNext.l || cropNext.t || cropNext.r || cropNext.b ? cropNext : undefined,
  });
  api.focus();
}
