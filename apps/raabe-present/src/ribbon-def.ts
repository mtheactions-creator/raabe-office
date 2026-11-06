// RaabePresent's ribbon, following PowerPoint 365's German ribbon (D-070; tabs, groups, order and
// keytips as recorded in results/2026-11-06-powerpoint-inventory.json). Cloud features are left out
// (the buttons are missing, not grey): Designer, Diktieren, Copilot, Stockbilder, Onlinebilder,
// Piktogramme, Add-Ins, Cameo, Online vorführen, Untertitel, Übersetzen, Intelligente Suche, Teilen.

import type { Control, MenuEntry, Tab } from "@raabe/shell";

type Popup = (close: () => void) => HTMLElement;

export interface RibbonContext {
  fontName(): string;
  fontSize(): string;
  fontOptions(): { heading?: string; items: string[] }[];
  setFont(name: string): void;
  setSize(value: string): void;
  fontColor(): string;
  setFontColor(hex: string): void;
  highlight(): string;
  setHighlight(hex: string): void;
  fillColor(): string;
  setFill(hex: string): void;
  lineColor(): string;
  setLine(hex: string): void;
  preview: { font(name: string): void; size(value: string): void };
  menus: {
    paste: MenuEntry[]; layouts: () => MenuEntry[]; section: MenuEntry[]; spacing: MenuEntry[]; case: MenuEntry[];
    bullets: MenuEntry[]; numbering: MenuEntry[]; lineSpacing: MenuEntry[]; columns: MenuEntry[]; direction: MenuEntry[]; alignText: MenuEntry[];
    arrange: () => MenuEntry[]; find: MenuEntry[]; select: MenuEntry[]; table: MenuEntry[]; pictures: MenuEntry[]; textBox: MenuEntry[];
    slideSize: MenuEntry[]; showCustom: MenuEntry[]; record: MenuEntry[]; transOptions: () => MenuEntry[]; animOptions: () => MenuEntry[];
    animStart: MenuEntry[]; colors: () => MenuEntry[]; fonts: () => MenuEntry[]; underline: MenuEntry[];
    outlineWeight: MenuEntry[]; outlineDash: MenuEntry[]; effects: MenuEntry[]; quickStyles: () => MenuEntry[]; zoomFit: MenuEntry[];
  };
  panels: { newSlide: Popup; layout: Popup; shapes: Popup; themes: Popup; transitions: Popup; animations: Popup; table: Popup; chart: Popup; addAnim: Popup; wordArt: Popup };
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
    ({ type: "color", cmd, key, palette, current: cur, pick, showLabel });
  const spin = (cmd: string, label: string, key: string, s: RibbonContext["spin"]["transDuration"], icon?: string): Control => ({ type: "spinner", cmd, label, key, icon, value: s.value, commit: s.commit, step: s.step });

  return [
    {
      id: "start", label: "Start", key: "R",
      shrink: ["editing", "drawing", "paragraph", "font", "slides", "clipboard"],
      groups: [
        { id: "clipboard", label: "Zwischenablage", key: "ZZ", icon: "paste", controls: [split("edit.paste", "V1", "large", m.paste), stack(b("edit.cut", "X"), b("edit.copy", "C"), b("edit.formatPainter", "FO"))] },
        { id: "slides", label: "Folien", key: "ZF", icon: "slide-add", controls: [splitPanel("slide.new", "ÄN", "large", p.newSlide), stack(panel("slide.layout", "OA", "small", p.layout), b("slide.reset", "FF"), dd("slide.section", "TA", "small", m.section))] },
        {
          id: "font", label: "Schriftart", key: "ZS", icon: "fonts", launcher: { cmd: "font.dialog", key: "TC" },
          controls: [rows(
            [fontName, fontSize, b("font.grow", "SV"), b("font.shrink", "J"), "gap", b("font.clear", "Q")],
            [b("font.bold", "1"), b("font.italic", "2"), b("font.underline", "3"), b("font.shadow", "5"), b("font.strike", "4"), dd("font.spacing", "6", "small", m.spacing), dd("font.case", "7", "small", m.case), "gap",
              color("font.highlight", "TF", ctx.highlight, ctx.setHighlight, "highlight"), color("font.color", "FC", ctx.fontColor, ctx.setFontColor)])],
        },
        {
          id: "paragraph", label: "Absatz", key: "ZA", icon: "paragraph-spacing", launcher: { cmd: "para.dialog", key: "BA" },
          controls: [rows(
            [split("para.bullets", "ÖU", "small", m.bullets), split("para.numbering", "AN", "small", m.numbering), b("para.levelDown", "DI"), b("para.levelUp", "ÄI"), dd("para.lineSpacing", "IZ", "small", m.lineSpacing), "gap", dd("para.direction", "MT", "small", m.direction)],
            [b("para.left", "II"), b("para.center", "RZ"), b("para.right", "RE"), b("para.justify", "OL"), dd("para.columns", "HS", "small", m.columns), "gap", dd("para.alignText", "PE", "small", m.alignText)])],
        },
        {
          id: "drawing", label: "Zeichnen", key: "ZD", icon: "shapes", launcher: { cmd: "draw.format", key: "BB" },
          controls: [panel("draw.shapes", "23", "large", p.shapes), lazy("draw.arrange", "G", "large", m.arrange), lazy("draw.quickStyles", "NC", "large", m.quickStyles),
            stack(color("draw.fill", "Ü", ctx.fillColor, ctx.setFill, "font", true), color("draw.outline", "V2", ctx.lineColor, ctx.setLine, "font", true), dd("draw.effects", "MO", "small", m.effects))],
        },
        { id: "editing", label: "Bearbeiten", key: "ZB", icon: "find", controls: [stack(split("edit.find", "FD", "small", m.find), split("edit.replace", "R", "small", [{ cmd: "edit.replace" }, { cmd: "edit.replaceFonts" }]), dd("edit.select", "SL", "small", m.select))] },
      ],
    },
    {
      id: "insert", label: "Einfügen", key: "I",
      shrink: ["media", "symbols", "text", "comments", "links", "illustrations", "images", "tables", "slides"],
      groups: [
        { id: "slides", label: "Folien", key: "ZF", icon: "slide-add", controls: [splitPanel("slide.new", "SI", "large", p.newSlide)] },
        { id: "tables", label: "Tabellen", key: "ZT", icon: "table", controls: [panel("insert.table", "LA", "large", p.table)] },
        { id: "images", label: "Bilder", key: "ZB", icon: "picture", controls: [dd("insert.pictures", "B", "large", m.pictures), b("insert.screenshot", "C1", "large"), b("insert.photoAlbum", "F", "large")] },
        { id: "illustrations", label: "Illustrationen", key: "ZI", icon: "shapes", controls: [panel("insert.shapes", "2", "large", p.shapes), b("insert.smartArt", "M", "large"), panel("insert.chart", "I", "large", p.chart)] },
        { id: "links", label: "Links", key: "ZL", icon: "link", controls: [b("insert.zoom", "Y", "large"), b("insert.link", "H1", "large"), b("insert.action", "T", "large")] },
        { id: "comments", label: "Kommentare", key: "ZK", icon: "comment", controls: [b("insert.comment", "N", "large")] },
        { id: "text", label: "Text", key: "ZX", icon: "text-box", controls: [split("insert.textBox", "LT", "large", m.textBox), b("insert.headerFooter", "KO", "large"), panel("insert.wordArt", "W", "large", p.wordArt),
          stack(b("insert.dateTime", "D"), b("insert.slideNumber", "H2"), b("insert.object", "9"))] },
        { id: "symbols", label: "Symbole", key: "ZS", icon: "equation", controls: [dd("insert.symbols", "ZS", "large", [{ cmd: "insert.equation" }, { cmd: "insert.symbol" }])] },
        { id: "media", label: "Medien", key: "ZM", icon: "picture", controls: [dd("insert.media", "ZM", "large", [{ cmd: "insert.video" }, { cmd: "insert.audio" }, { cmd: "insert.screenRecording" }])] },
      ],
    },
    {
      id: "design", label: "Entwurf", key: "U2",
      shrink: ["customize", "variants", "themes"],
      groups: [
        { id: "themes", label: "Designs", key: "ZT", icon: "themes", controls: [panel("design.themes", "G", "large", p.themes)] },
        { id: "variants", label: "Anpassen", key: "ZV", icon: "colors", controls: [b("design.variants", "V", "large"), stack(lazy("design.colors", "F", "small", m.colors), lazy("design.fonts", "6", "small", m.fonts), b("design.effects", "E"), b("design.background", "R"))] },
        { id: "customize", label: "Folien", key: "ZA", icon: "page-size", controls: [panel("slide.layout", "OA", "large", p.layout), dd("design.slideSize", "Ö", "large", m.slideSize)] },
      ],
    },
    {
      id: "transitions", label: "Übergänge", key: "Ä",
      shrink: ["timing", "gallery", "preview"],
      groups: [
        { id: "preview", label: "Vorschau", key: "ZV", icon: "eye", controls: [b("trans.preview", "V", "large")] },
        { id: "gallery", label: "Übergang zu dieser Folie", key: "ZU", icon: "transition", controls: [panel("trans.gallery", "Ü", "large", p.transitions), lazy("trans.options", "E", "large", m.transOptions)] },
        { id: "timing", label: "Anzeigedauer", key: "ZA", icon: "time", controls: [
          stack(b("trans.sound", "S"), spin("trans.duration", "Dauer:", "D", ctx.spin.transDuration), b("trans.applyAll", "F")),
          stack({ type: "check", cmd: "trans.onClick", key: "B" }, rows([{ type: "check", cmd: "trans.after", key: "N" }, spin("trans.afterTime", "", "A", ctx.spin.transAfter)]))] },
      ],
    },
    {
      id: "animations", label: "Animationen", key: "V",
      shrink: ["timing", "advanced", "gallery", "preview"],
      groups: [
        { id: "preview", label: "Vorschau", key: "ZV", icon: "eye", controls: [b("anim.preview", "U", "large")] },
        { id: "gallery", label: "Animation", key: "ZA", icon: "animation", launcher: { cmd: "anim.more", key: "W" }, controls: [panel("anim.gallery", "N", "large", p.animations), lazy("anim.options", "E", "large", m.animOptions)] },
        { id: "advanced", label: "Erweiterte Animation", key: "ZE", icon: "animation", controls: [panel("anim.add", "Q", "large", p.addAnim), stack(b("anim.pane", "I"), b("anim.trigger", "T"), b("anim.painter", "M"))] },
        { id: "timing", label: "Anzeigedauer", key: "ZD", icon: "time", controls: [
          stack(dd("anim.start", "S", "small", m.animStart, "Start:"), spin("anim.duration", "Dauer:", "D", ctx.spin.animDuration), spin("anim.delay", "Verzögerung:", "R", ctx.spin.animDelay)),
          stack(b("anim.earlier", "F"), b("anim.later", "P"))] },
      ],
    },
    {
      id: "slideshow", label: "Bildschirmpräsentation", key: "L",
      shrink: ["monitors", "setup", "start"],
      groups: [
        { id: "start", label: "Bildschirmpräsentation starten", key: "ZB", icon: "slideshow", controls: [b("show.fromStart", "V", "large"), b("show.fromCurrent", "T", "large"), dd("show.custom", "E", "large", m.showCustom)] },
        { id: "setup", label: "Einrichten", key: "ZE", icon: "settings", controls: [b("show.setup", "I1", "large"), b("show.hideSlide", "H", "large"), b("show.rehearse", "U1", "large"), dd("show.record", "N1", "large", m.record),
          stack(check("show.playNarrations", "K"), check("show.useTimings", "N2"), check("show.mediaControls", "D"))] },
        { id: "monitors", label: "Bildschirme", key: "ZM", icon: "app", controls: [dd("show.monitor", "ZM", "large", [{ cmd: "show.presenterView" }])] },
      ],
    },
    {
      id: "review", label: "Überprüfen", key: "P2",
      shrink: ["ink", "compare", "comments", "language", "accessibility", "proofing"],
      groups: [
        { id: "proofing", label: "Rechtschreibung", key: "ZD", icon: "spelling", controls: [b("review.spelling", "BE", "large"), b("review.thesaurus", "T", "large")] },
        { id: "accessibility", label: "Barrierefreiheit", key: "ZB", icon: "check", controls: [b("review.accessibility", "B1", "large")] },
        { id: "language", label: "Sprache", key: "ZS", icon: "globe", controls: [b("review.language", "SP", "large")] },
        { id: "comments", label: "Kommentare", key: "ZK", icon: "comment", controls: [b("review.newComment", "N", "large"), b("review.deleteComment", "L", "large"), b("review.prevComment", "G", "large"), b("review.nextComment", "W", "large"), b("review.showComments", "O", "large")] },
                { id: "ink", label: "Freihand", key: "ZF", icon: "pen", controls: [b("review.hideInk", "F", "large")] },
      ],
    },
    {
      id: "view", label: "Ansicht", key: "F",
      shrink: ["window", "color", "zoom", "show", "master", "views"],
      groups: [
        { id: "views", label: "Präsentationsansichten", key: "ZP", icon: "print-layout", controls: [b("view.normal", "K", "large"), b("view.outline", "I", "large"), b("view.sorter", "U", "large"), b("view.notesPage", "P", "large"), b("view.reading", "C", "large")] },
        { id: "master", label: "Masteransichten", key: "ZM", icon: "slide-layout", controls: [b("view.slideMaster", "M", "large"), b("view.handoutMaster", "B", "large"), b("view.notesMaster", "Ä", "large")] },
        { id: "show", label: "Anzeigen", key: "ZA", icon: "ruler", launcher: { cmd: "view.gridDialog", key: "R" }, controls: [stack(check("view.ruler", "L"), check("view.gridlines", "G"), check("view.guides", "H")), b("view.notes", "O", "large")] },
        { id: "zoom", label: "Zoom", key: "ZZ", icon: "zoom", controls: [b("view.zoom", "DO", "large"), b("view.fit", "F", "large")] },
        { id: "color", label: "Farbe/Graustufe", key: "ZF", icon: "colors", controls: [stack(b("view.color", "Q"), b("view.grayscale", "ß"), b("view.bw", "W"))] },
        { id: "window", label: "Fenster", key: "ZW", icon: "app", controls: [b("view.newWindow", "N", "large"), stack(b("view.arrangeAll", "A"), b("view.cascade", "Ü"), b("view.moveSplit", "T")), b("view.switchWindow", "E", "large")] },
        { id: "macros", label: "Makros", key: "Z0", icon: "settings", controls: [b("view.macros", "0", "large")] },
      ],
    },
  ];
}
