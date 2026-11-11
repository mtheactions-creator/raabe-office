// RaabePresent's ribbon, following PowerPoint 365's German ribbon (D-070; tabs, groups, order and
// keytips as recorded in results/2026-11-06-powerpoint-inventory.json). Cloud features are left out
// (the buttons are missing, not grey): Designer, Diktieren, Copilot, Stockbilder, Onlinebilder,
// Piktogramme, Add-Ins, Cameo, Online vorführen, Untertitel, Übersetzen, Intelligente Suche, Teilen.

import { MORE_COLORS } from "./colors";
import type { Control, MenuEntry, Tab } from "@raabe/shell";

type Popup = (close: () => void) => HTMLElement;

export interface RibbonContext {
  /** The design's colours for the palettes (PowerPoint shows the presentation's design there). */
  paletteTheme?(): import("@raabe/shell").PaletteTheme | undefined;
  run(cmd: string): void;
  /** Überprüfen > Kommentare anzeigen > Markup anzeigen: the comment markers on the slides. */
  markup(): boolean; toggleMarkup(): void;
  /** PowerPoint's names for some highlight colours. */
  highlightNames?: Readonly<Record<string, string>>;
  fontName(): string;
  fontSize(): string;
  fontOptions(): { heading?: string; items: string[] }[];
  setFont(name: string): void;
  setSize(value: string): void;
  fontColor(): string;
  setFontColor(hex: string): void;
  fontColorExtra(): MenuEntry[];
  highlight(): string;
  setHighlight(hex: string): void;
  fillColor(): string;
  setFill(hex: string): void;
  lineColor(): string;
  setLine(hex: string): void;
  /** Keine Kontur, Stärke, Striche under the Formkontur palette. */
  fillExtra(): MenuEntry[];
  outlineExtra(): MenuEntry[];
  /** Entwurf > Designs: the gallery in the ribbon. */
  designTiles(): Control;
  /** Übergänge / Animationen: the effects as a gallery in the ribbon (PowerPoint). */
  fxTiles(kind: "trans" | "anim"): Control;
  preview: { font(name: string): void; size(value: string): void };
  menus: {
    paste: MenuEntry[]; layouts: () => MenuEntry[]; section: MenuEntry[]; spacing: MenuEntry[]; case: MenuEntry[];
    bullets: MenuEntry[]; numbering: MenuEntry[]; lineSpacing: MenuEntry[]; columns: MenuEntry[]; direction: MenuEntry[]; alignText: MenuEntry[];
    arrange: () => MenuEntry[]; find: MenuEntry[]; select: MenuEntry[]; table: MenuEntry[]; pictures: MenuEntry[]; textBox: MenuEntry[];
    slideSize: MenuEntry[]; showCustom: MenuEntry[]; record: MenuEntry[]; recordClear: MenuEntry[]; transOptions: () => MenuEntry[]; animOptions: () => MenuEntry[];
    animStart: MenuEntry[]; colors: () => MenuEntry[]; fonts: () => MenuEntry[]; themeEffects: () => MenuEntry[]; underline: MenuEntry[];
    outlineWeight: MenuEntry[]; outlineDash: MenuEntry[]; effects: MenuEntry[]; quickStyles: () => MenuEntry[]; zoomFit: MenuEntry[]; equations: () => MenuEntry[]; sounds: () => MenuEntry[]; windows: () => MenuEntry[]; monitors: () => MenuEntry[]; smartArtConvert: () => MenuEntry[]; triggers: () => MenuEntry[]; zooms: () => MenuEntry[]; bgStyles: () => MenuEntry[]; inkAdd: () => MenuEntry[]; customShows: () => MenuEntry[];
  };
  panels: { newSlide: Popup; layout: Popup; shapes: Popup; themes: Popup; transitions: Popup; animations: Popup; table: Popup; chart: Popup; addAnim: Popup; wordArt: Popup; quickStyles: Popup };
  spin: {
    transDuration: { value(): string; commit(v: string): void; step(d: 1 | -1): void };
    transAfter: { value(): string; commit(v: string): void; step(d: 1 | -1): void };
    animDuration: { value(): string; commit(v: string): void; step(d: 1 | -1): void };
    animDelay: { value(): string; commit(v: string): void; step(d: 1 | -1): void };
    shapeH: { value(): string; commit(v: string): void; step(d: 1 | -1): void };
    shapeW: { value(): string; commit(v: string): void; step(d: 1 | -1): void };
  };
}

const b = (cmd: string, key: string, size: "large" | "medium" | "small" = "small", label?: string): Control => ({ type: "button", cmd, key, size, label });
const dd = (cmd: string, key: string, size: "large" | "medium" | "small", menu: MenuEntry[], label?: string): Control => ({ type: "dropdown", cmd, key, size, label, menu });
const split = (cmd: string, key: string, size: "large" | "medium" | "small", menu: MenuEntry[], label?: string): Control => ({ type: "split", cmd, key, size, menu, label });
const panel = (cmd: string, key: string, size: "large" | "medium" | "small", popup: Popup, label?: string): Control => ({ type: "dropdown", cmd, key, size, label, menu: [], popup });
const splitPanel = (cmd: string, key: string, size: "large" | "medium" | "small", popup: Popup, label?: string): Control => ({ type: "split", cmd, key, size, label, menu: [], popup });
const stack = (...items: Control[]): Control => ({ type: "stack", items });
const rows = (...r: (Control | "gap")[][]): Control => ({ type: "rows", rows: r });
const check = (cmd: string, key: string): Control => ({ type: "check", cmd, key });
const lazy = (cmd: string, key: string, size: "large" | "medium" | "small", menu: () => MenuEntry[], label?: string): Control => ({ type: "dropdown", cmd, key, size, label, get menu() { return menu(); } });

export function ribbonTabs(ctx: RibbonContext): Tab[] {
  const m = ctx.menus, p = ctx.panels;
  const fontName: Control = { type: "combo", cmd: "font.name", width: 130, key: "SC", value: ctx.fontName, options: ctx.fontOptions, optionStyle: (f) => `font-family:"${f}";font-size:14px`, commit: ctx.setFont, preview: ctx.preview.font };
  const fontSize: Control = { type: "combo", cmd: "font.size", width: 44, key: "SG", value: ctx.fontSize,
    options: () => [{ items: ["8", "9", "10", "10,5", "11", "12", "14", "16", "18", "20", "24", "28", "32", "36", "40", "44", "48", "54", "60", "66", "72", "80", "88", "96"] }],
    commit: ctx.setSize, preview: ctx.preview.size };
  const color = (cmd: string, key: string, cur: () => string, pick: (c: string) => void, palette: "font" | "highlight" = "font", showLabel = false): Control =>
    ({ type: "color", cmd, key, palette, current: cur, pick, showLabel, theme: ctx.paletteTheme, more: MORE_COLORS[cmd] });
  const spin = (cmd: string, label: string, key: string, s: RibbonContext["spin"]["transDuration"], icon?: string): Control => ({ type: "spinner", cmd, label, key, icon, value: s.value, commit: s.commit, step: s.step });

  return [
    {
      id: "start", label: "Start", key: "R",
      // First from PowerPoint's widest layouts to its usual ones (D-159), from the right; then its sequence.
      shrink: ["editing", "drawing", "paragraph", "slides", "editing", "drawing", "paragraph", "font", "slides", "clipboard"],
      groups: [
        { id: "clipboard", label: "Zwischenablage", key: "ZZ", icon: "paste", wide: false, controls: [split("edit.paste", "V1", "large", m.paste), stack(b("edit.cut", "X"), split("edit.copy", "C", "small", [{ cmd: "edit.copy" }, { cmd: "edit.duplicate" }]), b("edit.formatPainter", "FO"))] },
        // The widest layouts are PowerPoint's at a wide window (labels beside the small icons, D-158); the ones before
        // are its first reductions.
        { id: "slides", label: "Folien", key: "ZF", icon: "slide-add",
          controls: [splitPanel("slide.new", "ÄN", "large", p.newSlide), stack(panel("slide.layout", "OA", "medium", p.layout), b("slide.reset", "FF", "medium"), dd("slide.section", "TA", "medium", m.section))],
          smaller: [[splitPanel("slide.new", "ÄN", "large", p.newSlide), stack(panel("slide.layout", "OA", "small", p.layout), b("slide.reset", "FF"), dd("slide.section", "TA", "small", m.section))]] },
        {
          id: "font", label: "Schriftart", key: "ZS", icon: "fonts", launcher: { cmd: "font.dialog", key: "TC" },
          controls: [rows(
            [fontName, fontSize, b("font.grow", "SV"), b("font.shrink", "J"), "gap", b("font.clear", "Q")],
            [b("font.bold", "1"), b("font.italic", "2"), b("font.underline", "3"), b("font.shadow", "5"), b("font.strike", "4"), dd("font.spacing", "6", "small", m.spacing), dd("font.case", "7", "small", m.case), "gap",
              { ...color("font.highlight", "TF", ctx.highlight, ctx.setHighlight, "highlight"), names: ctx.highlightNames } as Control, { ...color("font.color", "FC", ctx.fontColor, ctx.setFontColor), extra: () => ctx.fontColorExtra() } as Control])],
        },
        {
          id: "paragraph", label: "Absatz", key: "ZA", icon: "paragraph-spacing", launcher: { cmd: "para.dialog", key: "BA" },
          controls: [rows(
            [split("para.bullets", "ÖU", "small", m.bullets), split("para.numbering", "AN", "small", m.numbering), b("para.levelDown", "DI"), b("para.levelUp", "ÄI"), dd("para.lineSpacing", "IZ", "small", m.lineSpacing)],
            [b("para.left", "II"), b("para.center", "RZ"), b("para.right", "RE"), b("para.justify", "OL"), dd("para.columns", "HS", "small", m.columns)]),
          stack(dd("para.direction", "MT", "medium", m.direction), dd("para.alignText", "PE", "medium", m.alignText), lazy("para.smartArt", "NI", "medium", m.smartArtConvert))],
          smaller: [[rows(
            [split("para.bullets", "ÖU", "small", m.bullets), split("para.numbering", "AN", "small", m.numbering), b("para.levelDown", "DI"), b("para.levelUp", "ÄI"), dd("para.lineSpacing", "IZ", "small", m.lineSpacing), "gap", dd("para.direction", "MT", "small", m.direction)],
            [b("para.left", "II"), b("para.center", "RZ"), b("para.right", "RE"), b("para.justify", "OL"), dd("para.columns", "HS", "small", m.columns), "gap", dd("para.alignText", "PE", "small", m.alignText), lazy("para.smartArt", "NI", "small", m.smartArtConvert)])]],
        },
        {
          id: "drawing", label: "Zeichnen", key: "ZD", icon: "shapes", launcher: { cmd: "draw.format", key: "BB" },
          controls: [panel("draw.shapes", "23", "large", p.shapes), lazy("draw.arrange", "G", "large", m.arrange), panel("draw.quickStyles", "NC", "large", p.quickStyles),
            stack({ ...color("draw.fill", "Ü", ctx.fillColor, ctx.setFill, "font", true), extra: () => ctx.fillExtra() } as Control, { ...color("draw.outline", "V2", ctx.lineColor, ctx.setLine, "font", true), extra: () => ctx.outlineExtra() } as Control, dd("draw.effects", "MO", "medium", m.effects))],
          smaller: [[panel("draw.shapes", "23", "large", p.shapes), lazy("draw.arrange", "G", "large", m.arrange), panel("draw.quickStyles", "NC", "large", p.quickStyles),
            stack({ ...color("draw.fill", "Ü", ctx.fillColor, ctx.setFill, "font", true), extra: () => ctx.fillExtra() } as Control, { ...color("draw.outline", "V2", ctx.lineColor, ctx.setLine, "font", true), extra: () => ctx.outlineExtra() } as Control, dd("draw.effects", "MO", "small", m.effects))]],
        },
        { id: "editing", label: "Bearbeiten", key: "ZB", icon: "find",
          controls: [stack(split("edit.find", "FD", "medium", m.find), split("edit.replace", "R", "medium", [{ cmd: "edit.replace" }, { cmd: "edit.replaceFonts" }]), dd("edit.select", "SL", "medium", m.select))],
          smaller: [[stack(split("edit.find", "FD", "small", m.find), split("edit.replace", "R", "small", [{ cmd: "edit.replace" }, { cmd: "edit.replaceFonts" }]), dd("edit.select", "SL", "small", m.select))]] },
      ],
    },
    {
      id: "insert", label: "Einfügen", key: "I",
      shrink: ["media", "symbols", "text", "comments", "links", "illustrations", "images", "tables", "slides"],
      groups: [
        { id: "slides", label: "Folien", key: "ZF", icon: "slide-add", controls: [splitPanel("slide.new", "SI", "large", p.newSlide)] },
        { id: "tables", label: "Tabellen", key: "ZT", icon: "table", controls: [panel("insert.table", "LA", "large", p.table)] },
        { id: "images", label: "Bilder", key: "ZB", icon: "picture", controls: [dd("insert.pictures", "B", "large", m.pictures), b("insert.screenshot", "C1", "large"), split("insert.photoAlbum", "F", "large", [{ cmd: "insert.photoAlbum", label: "Neues Fotoalbum..." }])] },
        { id: "illustrations", label: "Illustrationen", key: "ZI", icon: "shapes", controls: [panel("insert.shapes", "2", "large", p.shapes), b("insert.smartArt", "M", "large"), panel("insert.chart", "I", "large", p.chart)] },
        { id: "links", label: "Links", key: "ZL", icon: "link", controls: [lazy("insert.zoom", "Y", "large", m.zooms), b("insert.link", "H1", "large"), b("insert.action", "T", "large")] },
        { id: "comments", label: "Kommentare", key: "ZK", icon: "comment", controls: [b("insert.comment", "N", "large")] },
        { id: "text", label: "Text", key: "ZX", icon: "text-box", controls: [split("insert.textBox", "LT", "large", m.textBox), b("insert.headerFooter", "KO", "large"), panel("insert.wordArt", "W", "large", p.wordArt),
          stack(b("insert.dateTime", "D"), b("insert.slideNumber", "H2"), b("insert.object", "9"))] },
        { id: "symbols", label: "Symbole", key: "ZS", icon: "equation", controls: [lazy("insert.symbols", "ZS", "large", () => [{ cmd: "insert.equation" }, { label: "Integriert", submenu: () => m.equations() }, { cmd: "insert.symbol" }])] },
        { id: "media", label: "Medien", key: "ZM", icon: "picture", controls: [dd("insert.media", "ZM", "large", [{ cmd: "insert.video" }, { cmd: "insert.audio" }, { cmd: "insert.screenRecording" }])] },
      ],
    },
    {
      id: "draw", label: "Zeichnen", key: "JI",
      shrink: ["convert", "tools", "undo"],
      groups: [
        { id: "undo", label: "Rückgängig", key: "ZR", icon: "undo", controls: [stack(b("edit.undo", "AZ"), b("edit.redo", "AQ"))] },
        { id: "tools", label: "Zeichentools", key: "ZZ", icon: "pen", controls: [b("ink.select", "G", "large"), b("ink.lasso", "G", "large"), b("ink.eraser", "G", "large"), b("ink.penBlack", "G", "large"), b("ink.penRed", "G", "large"), b("ink.pencil", "G", "large"), b("ink.highlighter", "G", "large"), lazy("ink.add", "A1", "large", m.inkAdd)] },
        { id: "stencils", label: "Schablonen", key: "ZS", icon: "ruler", controls: [b("ink.ruler", "L", "large")] },
        { id: "convert", label: "Konvertieren", key: "ZK", icon: "shapes", controls: [b("ink.toShape", "IS", "large"), b("ink.toMath", "F1", "large")] },
        { id: "replay", label: "Wiedergeben", key: "ZW", icon: "play", controls: [b("ink.replay", "W", "large")] },
      ],
    },
    {
      id: "design", label: "Entwurf", key: "U2",
      shrink: ["customize", "variants", "themes"],
      groups: [
        // PowerPoint shows the designs as a gallery in the ribbon and Anpassen as five large buttons (UI screenshot, D-072).
        { id: "themes", label: "Designs", key: "ZT", icon: "themes", controls: [ctx.designTiles()] },
        { id: "variants", label: "Anpassen", key: "ZV", icon: "colors", controls: [b("design.variants", "V", "large"), lazy("design.colors", "F", "large", m.colors), lazy("design.fonts", "6", "large", m.fonts), lazy("design.effects", "E", "large", m.themeEffects), lazy("design.background", "R", "large", m.bgStyles)] },
        { id: "customize", label: "Folien", key: "ZA", icon: "page-size", controls: [panel("slide.layout", "OA", "large", p.layout), dd("design.slideSize", "Ö", "large", m.slideSize)] },
      ],
    },
    {
      id: "transitions", label: "Übergänge", key: "Ä",
      shrink: ["timing", "gallery", "preview"],
      groups: [
        { id: "preview", label: "Vorschau", key: "ZV", icon: "eye", controls: [b("trans.preview", "V", "large")] },
        { id: "gallery", label: "Übergang zu dieser Folie", key: "ZU", icon: "transition", controls: [ctx.fxTiles("trans"), lazy("trans.options", "E", "large", m.transOptions)] },
        { id: "timing", label: "Anzeigedauer", key: "ZA", icon: "time", controls: [
          stack(lazy("trans.sound", "S", "medium", m.sounds), spin("trans.duration", "Dauer:", "D", ctx.spin.transDuration), b("trans.applyAll", "F", "medium")),
          stack({ type: "check", cmd: "trans.onClick", key: "B" }, rows([{ type: "check", cmd: "trans.after", key: "N" }, spin("trans.afterTime", "", "A", ctx.spin.transAfter)]))] },
      ],
    },
    {
      id: "animations", label: "Animationen", key: "V",
      shrink: ["timing", "advanced", "gallery", "preview"],
      groups: [
        { id: "preview", label: "Vorschau", key: "ZV", icon: "eye", controls: [b("anim.preview", "U", "large")] },
        { id: "gallery", label: "Animation", key: "ZA", icon: "animation", launcher: { cmd: "anim.more", key: "W" }, controls: [ctx.fxTiles("anim"), lazy("anim.options", "E", "large", m.animOptions)] },
        { id: "advanced", label: "Erweiterte Animation", key: "ZE", icon: "animation", controls: [panel("anim.add", "Q", "large", p.addAnim), stack(b("anim.pane", "I", "medium"), lazy("anim.trigger", "T", "medium", m.triggers), b("anim.painter", "M", "medium"))] },
        { id: "timing", label: "Anzeigedauer", key: "ZD", icon: "time", controls: [
          stack(dd("anim.start", "S", "medium", m.animStart, "Start:"), spin("anim.duration", "Dauer:", "D", ctx.spin.animDuration), spin("anim.delay", "Verzögerung:", "R", ctx.spin.animDelay)),
          { type: "stack", title: "Animation neu anordnen", items: [b("anim.earlier", "F", "medium"), b("anim.later", "P", "medium")] }] },
      ],
    },
    {
      id: "slideshow", label: "Bildschirmpräsentation", key: "L",
      shrink: ["monitors", "setup", "start"],
      groups: [
        { id: "start", label: "Bildschirmpräsentation starten", key: "ZB", icon: "slideshow", controls: [b("show.fromStart", "V", "large"), b("show.fromCurrent", "T", "large"), lazy("show.custom", "E", "large", m.customShows)] },
        { id: "setup", label: "Einrichten", key: "ZE", icon: "settings", controls: [b("show.setup", "I1", "large"), stack(b("show.hideSlide", "H", "medium"), b("show.rehearse", "U1", "medium"), dd("show.record", "N1", "medium", m.record)),
          stack(check("show.playNarrations", "K")), stack(check("show.useTimings", "N2"), check("show.mediaControls", "D"))] },
        { id: "monitors", label: "Bildschirme", key: "ZM", icon: "app", controls: [lazy("show.monitor", "ZM", "large", m.monitors)] },
      ],
    },
    {
      // Aufzeichnen (D-078): PowerPoint's groups and keytips (Cameo is a cloud feature and left out).
      id: "record", label: "Aufzeichnen", key: "P1",
      shrink: ["export", "edit", "rec"],
      groups: [
        { id: "preview", label: "Vorschau", key: "ZV", icon: "slideshow", controls: [b("rec.preview", "P", "large")] },
        { id: "rec", label: "Aufzeichnen", key: "ZA", icon: "dictate", controls: [b("rec.fromStart", "B", "large"), b("rec.fromCurrent", "CS", "large"), b("rec.screen", "R1", "large"), b("rec.audio", "O", "large")] },
        { id: "edit", label: "Bearbeiten", key: "ZB", icon: "close-x", controls: [dd("rec.clear", "CR", "large", m.recordClear)] },
        { id: "export", label: "Exportieren", key: "ZE", icon: "export", controls: [b("rec.saveShow", "S", "large"), b("rec.video", "Y", "large")] },
      ],
    },
    {
      id: "review", label: "Überprüfen", key: "P2",
      shrink: ["ink", "compare", "comments", "language", "accessibility", "proofing"],
      groups: [
        { id: "proofing", label: "Rechtschreibung", key: "ZD", icon: "spelling", controls: [b("review.spelling", "BE", "large", "Rechtschreibung"), b("review.thesaurus", "T", "large")] },
        { id: "accessibility", label: "Barrierefreiheit", key: "ZB", icon: "check", controls: [split("review.accessibility", "B1", "large", [{ cmd: "review.accessibility" }, { label: "Alternativtext", run: () => ctx.run("shape.altTextPane") }, { label: "Lesereihenfolgebereich", run: () => ctx.run("arrange.selectionPane") }, "-", { label: "Optionen: Barrierefreiheit", disabled: true, run: () => {} }])] },
        { id: "language", label: "Sprache", key: "ZS", icon: "globe", controls: [dd("review.language", "SP", "large", [{ cmd: "review.setLanguage" }, { cmd: "review.languageOptions" }])] },
        { id: "comments", label: "Kommentare", key: "ZK", icon: "comment", controls: [b("review.newComment", "N", "large"), b("review.deleteComment", "L", "large"), b("review.prevComment", "G", "large"), b("review.nextComment", "W", "large"), split("review.showComments", "O", "large", [{ label: "Kommentarbereich", run: () => ctx.run("review.showComments") }, { label: "Markup anzeigen", checked: ctx.markup(), run: () => ctx.toggleMarkup() }])] },
                { id: "ink", label: "Freihand", key: "ZF", icon: "pen", controls: [b("review.hideInk", "F", "large")] },
      ],
    },
    {
      id: "view", label: "Ansicht", key: "F",
      shrink: ["window", "color", "zoom", "show", "master", "views"],
      groups: [
        { id: "views", label: "Präsentationsansichten", key: "ZP", icon: "print-layout", controls: [b("view.normal", "K", "large"), b("view.outline", "I", "large"), b("view.sorter", "U", "large"), b("view.notesPage", "P", "large"), b("view.reading", "C", "large")] },
        { id: "master", label: "Masteransichten", key: "ZM", icon: "slide-layout", controls: [stack(b("view.slideMaster", "M", "medium"), b("view.handoutMaster", "B", "medium"), b("view.notesMaster", "Ä", "medium"))] },
        { id: "show", label: "Anzeigen", key: "ZA", icon: "ruler", launcher: { cmd: "view.gridDialog", key: "R" }, controls: [stack(check("view.ruler", "L"), check("view.gridlines", "G"), check("view.guides", "H")), b("view.notes", "O", "large")] },
        { id: "zoom", label: "Zoom", key: "ZZ", icon: "zoom", controls: [b("view.zoom", "DO", "large"), b("view.fit", "F", "large")] },
        { id: "color", label: "Farbe/Graustufe", key: "ZF", icon: "colors", controls: [stack(b("view.color", "Q", "medium"), b("view.grayscale", "ß", "medium"), b("view.bw", "W", "medium"))] },
        { id: "window", label: "Fenster", key: "ZW", icon: "app", controls: [b("view.newWindow", "N", "large"), stack(b("view.arrangeAll", "A"), b("view.cascade", "Ü"), b("view.moveSplit", "T")), lazy("view.switchWindow", "E", "large", m.windows)] },
        { id: "macros", label: "Makros", key: "Z0", icon: "settings", controls: [b("view.macros", "0", "large")] },
      ],
    },
  ];
}
