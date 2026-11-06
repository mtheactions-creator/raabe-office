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
  panels: { newSlide: Popup; layout: Popup; shapes: Popup; themes: Popup; transitions: Popup; animations: Popup; table: Popup; chart: Popup; addAnim: Popup };
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
        { id: "clipboard", label: "Zwischenablage", key: "ZZ", icon: "paste", controls: [split("edit.paste", "V", "large", m.paste), stack(b("edit.cut", "X"), b("edit.copy", "C"), b("edit.formatPainter", "FP"))] },
        { id: "slides", label: "Folien", key: "ZF", icon: "slide-add", controls: [splitPanel("slide.new", "I", "large", p.newSlide), stack(panel("slide.layout", "L", "small", p.layout), b("slide.reset", "A"), dd("slide.section", "T", "small", m.section))] },
        {
          id: "font", label: "Schriftart", key: "ZS", icon: "fonts", launcher: { cmd: "font.dialog", key: "FN" },
          controls: [rows(
            [fontName, fontSize, b("font.grow", "SV"), b("font.shrink", "SK"), "gap", b("font.clear", "E")],
            [b("font.bold", "1"), b("font.italic", "2"), b("font.underline", "3"), b("font.shadow", "5"), b("font.strike", "4"), dd("font.spacing", "ZA", "small", m.spacing), dd("font.case", "7", "small", m.case), "gap",
              color("font.highlight", "HL", ctx.highlight, ctx.setHighlight, "highlight"), color("font.color", "FC", ctx.fontColor, ctx.setFontColor)])],
        },
        {
          id: "paragraph", label: "Absatz", key: "ZA", icon: "paragraph-spacing", launcher: { cmd: "para.dialog", key: "PG" },
          controls: [rows(
            [split("para.bullets", "U", "small", m.bullets), split("para.numbering", "N", "small", m.numbering), b("para.levelDown", "AO"), b("para.levelUp", "AI"), dd("para.lineSpacing", "K", "small", m.lineSpacing), "gap", dd("para.direction", "AX", "small", m.direction)],
            [b("para.left", "AL"), b("para.center", "AC"), b("para.right", "AR"), b("para.justify", "AJ"), dd("para.columns", "AM", "small", m.columns), "gap", dd("para.alignText", "AT", "small", m.alignText)])],
        },
        {
          id: "drawing", label: "Zeichnung", key: "ZD", icon: "shapes", launcher: { cmd: "draw.format", key: "SA" },
          controls: [panel("draw.shapes", "SH", "large", p.shapes), lazy("draw.arrange", "G", "large", m.arrange), lazy("draw.quickStyles", "Q", "large", m.quickStyles),
            stack(color("draw.fill", "SF", ctx.fillColor, ctx.setFill, "font", true), color("draw.outline", "SO", ctx.lineColor, ctx.setLine, "font", true), dd("draw.effects", "SE", "small", m.effects))],
        },
        { id: "editing", label: "Bearbeiten", key: "ZB", icon: "find", controls: [stack(split("edit.find", "FD", "small", m.find), split("edit.replace", "R", "small", [{ cmd: "edit.replace" }, { cmd: "edit.replaceFonts" }]), dd("edit.select", "SL", "small", m.select))] },
      ],
    },
    {
      id: "insert", label: "Einfügen", key: "I",
      shrink: ["media", "symbols", "text", "comments", "links", "illustrations", "images", "tables", "slides"],
      groups: [
        { id: "slides", label: "Folien", key: "ZF", icon: "slide-add", controls: [splitPanel("slide.new", "I", "large", p.newSlide)] },
        { id: "tables", label: "Tabellen", key: "ZT", icon: "table", controls: [panel("insert.table", "T", "large", p.table)] },
        { id: "images", label: "Bilder", key: "ZB", icon: "picture", controls: [dd("insert.pictures", "P", "large", m.pictures), b("insert.screenshot", "SC", "large"), b("insert.photoAlbum", "A", "large")] },
        { id: "illustrations", label: "Illustrationen", key: "ZI", icon: "shapes", controls: [panel("insert.shapes", "SH", "large", p.shapes), b("insert.smartArt", "M", "large"), panel("insert.chart", "C", "large", p.chart)] },
        { id: "links", label: "Links", key: "ZL", icon: "link", controls: [b("insert.zoom", "Z", "large"), b("insert.link", "H", "large"), b("insert.action", "K", "large")] },
        { id: "comments", label: "Kommentare", key: "ZK", icon: "comment", controls: [b("insert.comment", "L", "large")] },
        { id: "text", label: "Text", key: "ZX", icon: "text-box", controls: [split("insert.textBox", "X", "large", m.textBox), b("insert.headerFooter", "HF", "large"), b("insert.wordArt", "W", "large"),
          stack(b("insert.dateTime", "D"), b("insert.slideNumber", "SN"), b("insert.object", "J"))] },
        { id: "symbols", label: "Symbole", key: "ZS", icon: "equation", controls: [b("insert.equation", "E", "large"), b("insert.symbol", "U", "large")] },
        { id: "media", label: "Medien", key: "ZM", icon: "picture", controls: [b("insert.video", "V", "large"), b("insert.audio", "O", "large"), b("insert.screenRecording", "R", "large")] },
      ],
    },
    {
      id: "design", label: "Entwurf", key: "G",
      shrink: ["customize", "variants", "themes"],
      groups: [
        { id: "themes", label: "Designs", key: "ZT", icon: "themes", controls: [panel("design.themes", "H", "large", p.themes)] },
        { id: "variants", label: "Varianten", key: "ZV", icon: "colors", controls: [stack(lazy("design.colors", "C", "small", m.colors), lazy("design.fonts", "F", "small", m.fonts), b("design.effects", "E"), b("design.background", "B"))] },
        { id: "customize", label: "Anpassen", key: "ZA", icon: "page-size", controls: [dd("design.slideSize", "S", "large", m.slideSize), b("design.formatBackground", "G", "large")] },
      ],
    },
    {
      id: "transitions", label: "Übergänge", key: "K",
      shrink: ["timing", "gallery", "preview"],
      groups: [
        { id: "preview", label: "Vorschau", key: "ZV", icon: "eye", controls: [b("trans.preview", "P", "large")] },
        { id: "gallery", label: "Übergang zu dieser Folie", key: "ZU", icon: "transition", controls: [panel("trans.gallery", "T", "large", p.transitions), lazy("trans.options", "E", "large", m.transOptions)] },
        { id: "timing", label: "Anzeigedauer", key: "ZA", icon: "time", controls: [
          stack(b("trans.sound", "U"), spin("trans.duration", "Dauer:", "D", ctx.spin.transDuration), b("trans.applyAll", "L")),
          stack({ type: "check", cmd: "trans.onClick", key: "M" }, rows([{ type: "check", cmd: "trans.after", key: "F" }, spin("trans.afterTime", "", "N", ctx.spin.transAfter)]))] },
      ],
    },
    {
      id: "animations", label: "Animationen", key: "A",
      shrink: ["timing", "advanced", "gallery", "preview"],
      groups: [
        { id: "preview", label: "Vorschau", key: "ZV", icon: "eye", controls: [b("anim.preview", "P", "large")] },
        { id: "gallery", label: "Animation", key: "ZA", icon: "animation", controls: [panel("anim.gallery", "S", "large", p.animations), lazy("anim.options", "M", "large", m.animOptions)] },
        { id: "advanced", label: "Erweiterte Animation", key: "ZE", icon: "animation", controls: [panel("anim.add", "AA", "large", p.addAnim), stack(b("anim.pane", "C"), b("anim.trigger", "T"), b("anim.painter", "AP"))] },
        { id: "timing", label: "Anzeigedauer", key: "ZD", icon: "time", controls: [
          stack(dd("anim.start", "TS", "small", m.animStart, "Start:"), spin("anim.duration", "Dauer:", "TD", ctx.spin.animDuration), spin("anim.delay", "Verzögerung:", "TL", ctx.spin.animDelay)),
          stack(b("anim.earlier", "E"), b("anim.later", "L"))] },
      ],
    },
    {
      id: "slideshow", label: "Bildschirmpräsentation", key: "N",
      shrink: ["monitors", "setup", "start"],
      groups: [
        { id: "start", label: "Bildschirmpräsentation starten", key: "ZB", icon: "slideshow", controls: [b("show.fromStart", "B", "large"), b("show.fromCurrent", "C", "large"), dd("show.custom", "W", "large", m.showCustom)] },
        { id: "setup", label: "Einrichten", key: "ZE", icon: "settings", controls: [b("show.setup", "S", "large"), b("show.hideSlide", "H", "large"), b("show.rehearse", "T", "large"), dd("show.record", "N", "large", m.record),
          stack(check("show.playNarrations", "P"), check("show.useTimings", "U"), check("show.mediaControls", "W"))] },
        { id: "monitors", label: "Bildschirme", key: "ZM", icon: "app", controls: [b("show.monitor", "M", "small"), check("show.presenterView", "V")] },
      ],
    },
    {
      id: "review", label: "Überprüfen", key: "P",
      shrink: ["ink", "compare", "comments", "language", "accessibility", "proofing"],
      groups: [
        { id: "proofing", label: "Dokumentprüfung", key: "ZD", icon: "spelling", controls: [b("review.spelling", "S", "large"), b("review.thesaurus", "E", "large")] },
        { id: "accessibility", label: "Barrierefreiheit", key: "ZB", icon: "check", controls: [b("review.accessibility", "A", "large")] },
        { id: "language", label: "Sprache", key: "ZS", icon: "globe", controls: [b("review.language", "U", "large")] },
        { id: "comments", label: "Kommentare", key: "ZK", icon: "comment", controls: [b("review.newComment", "C", "large"), b("review.deleteComment", "D", "large"), b("review.prevComment", "V", "large"), b("review.nextComment", "N", "large"), b("review.showComments", "P", "large")] },
        { id: "compare", label: "Vergleichen", key: "ZV", icon: "document", controls: [b("review.compare", "H", "large")] },
        { id: "ink", label: "Freihand", key: "ZF", icon: "pen", controls: [b("review.hideInk", "I", "large")] },
      ],
    },
    {
      id: "view", label: "Ansicht", key: "F",
      shrink: ["window", "color", "zoom", "show", "master", "views"],
      groups: [
        { id: "views", label: "Präsentationsansichten", key: "ZP", icon: "print-layout", controls: [b("view.normal", "N", "large"), b("view.outline", "G", "large"), b("view.sorter", "S", "large"), b("view.notesPage", "T", "large"), b("view.reading", "D", "large")] },
        { id: "master", label: "Masteransichten", key: "ZM", icon: "slide-layout", controls: [b("view.slideMaster", "M", "large"), b("view.handoutMaster", "H", "large"), b("view.notesMaster", "I", "large")] },
        { id: "show", label: "Anzeigen", key: "ZA", icon: "ruler", launcher: { cmd: "view.gridDialog", key: "GD" }, controls: [stack(check("view.ruler", "R"), check("view.gridlines", "GL"), check("view.guides", "GU")), b("view.notes", "NO", "large")] },
        { id: "zoom", label: "Zoom", key: "ZZ", icon: "zoom", controls: [b("view.zoom", "Q", "large"), b("view.fit", "W", "large")] },
        { id: "color", label: "Farbe/Graustufe", key: "ZF", icon: "colors", controls: [stack(b("view.color", "C"), b("view.grayscale", "U"), b("view.bw", "B"))] },
        { id: "window", label: "Fenster", key: "ZW", icon: "app", controls: [b("view.newWindow", "F", "large")] },
      ],
    },
  ];
}
