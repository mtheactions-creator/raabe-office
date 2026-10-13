// RaabeWrite's ribbon, following Word 365's German ribbon as captured on 2026-10-13
// (tools/feel-lab/capture_word_ui.py): same tabs, groups, order, sizes and keytips.
// Left out on purpose: Sendungen, Entwicklertools, Hilfe, and groups for features we do not
// plan (cloud, add-ins, SmartArt, 3D models, macros ...). See docs/SCOPE.md.

import type { Control, MenuEntry, Tab } from "@raabe/shell";

/** What the ribbon needs from the app to show values and apply changes. */
export interface RibbonContext {
  fontName(): string;
  fontSize(): string;
  fontOptions(): { heading?: string; items: string[] }[];
  setFont(name: string): void;
  setSize(value: string): void;
  fontColor(): string;
  setFontColor(hex: string): void;
  highlight(): string;
  setHighlight(name: string): void;
  indent(which: "left" | "right"): string;
  setIndent(which: "left" | "right", value: string): void;
  stepIndent(which: "left" | "right", dir: 1 | -1): void;
  spacing(which: "before" | "after"): string;
  setSpacing(which: "before" | "after", value: string): void;
  stepSpacing(which: "before" | "after", dir: 1 | -1): void;
  /** Schattierung (paragraph or cells). */
  shading(): string;
  setShading(hex: string): void;
  /** Galleries and menus of the Step 3 features. */
  panels: {
    bullets: Popup; numbers: Popup; multilevel: Popup; insertTable: Popup; margins: Popup; size: Popup;
    columns: MenuEntry[]; breaks: MenuEntry[]; header: MenuEntry[]; footer: MenuEntry[]; pageNumber: MenuEntry[];
    pictures: MenuEntry[]; borders: MenuEntry[]; wrap: MenuEntry[];
  };
}

type Popup = (close: () => void) => HTMLElement;

const b = (cmd: string, key: string, size: "large" | "medium" | "small" = "small", label?: string): Control =>
  ({ type: "button", cmd, key, size, label });
const dd = (cmd: string, key: string, size: "large" | "medium" | "small" = "small", label?: string, menu?: MenuEntry[]): Control =>
  ({ type: "dropdown", cmd, key, size, label, menu: menu ?? [{ cmd }] });
const split = (cmd: string, key: string, size: "large" | "medium" | "small" = "small", menu: MenuEntry[] = [{ cmd }], label?: string, popup?: Popup): Control =>
  ({ type: "split", cmd, key, size, menu, label, popup });
const panel = (cmd: string, key: string, size: "large" | "medium" | "small", popup: Popup): Control =>
  ({ type: "dropdown", cmd, key, size, menu: [], popup });
const stack = (...items: Control[]): Control => ({ type: "stack", items });

/** Gallery entries in Word's order: built-in style name, command, preview CSS. */
export const GALLERY: { name: string; cmd: string; preview: string }[] = [
  { name: "Normal", cmd: "style.normal", preview: "font:14.667px Calibri,Carlito,sans-serif" },
  { name: "No Spacing", cmd: "style.noSpacing", preview: "font:14.667px Calibri,Carlito,sans-serif" },
  { name: "heading 1", cmd: "style.h1", preview: "font:21.333px 'Calibri Light',Carlito,sans-serif;color:#2f5496" },
  { name: "heading 2", cmd: "style.h2", preview: "font:17.333px 'Calibri Light',Carlito,sans-serif;color:#2f5496" },
  { name: "Title", cmd: "style.title", preview: "font:28px 'Calibri Light',Carlito,sans-serif;letter-spacing:-0.5px" },
  { name: "Subtitle", cmd: "style.subtitle", preview: "font:14.667px Calibri,Carlito,sans-serif;color:#5a5a5a;letter-spacing:1px" },
  { name: "Subtle Emphasis", cmd: "style.subtleEmphasis", preview: "font:italic 14.667px Calibri,Carlito,sans-serif;color:#404040" },
  { name: "Emphasis", cmd: "style.emphasis", preview: "font:italic 14.667px Calibri,Carlito,sans-serif" },
  { name: "Intense Emphasis", cmd: "style.intenseEmphasis", preview: "font:italic 14.667px Calibri,Carlito,sans-serif;color:#4472c4" },
  { name: "Strong", cmd: "style.strong", preview: "font:bold 14.667px Calibri,Carlito,sans-serif" },
  { name: "Quote", cmd: "style.quote", preview: "font:italic 14.667px Calibri,Carlito,sans-serif;color:#404040" },
  { name: "Intense Quote", cmd: "style.intenseQuote", preview: "font:italic 14.667px Calibri,Carlito,sans-serif;color:#4472c4" },
  { name: "Subtle Reference", cmd: "style.subtleReference", preview: "font:14.667px Calibri,Carlito,sans-serif;font-variant:small-caps;color:#5a5a5a" },
  { name: "Intense Reference", cmd: "style.intenseReference", preview: "font:bold 14.667px Calibri,Carlito,sans-serif;font-variant:small-caps;color:#4472c4" },
  { name: "Book Title", cmd: "style.bookTitle", preview: "font:bold italic 14.667px Calibri,Carlito,sans-serif" },
  { name: "List Paragraph", cmd: "style.listParagraph", preview: "font:14.667px Calibri,Carlito,sans-serif" },
];

export function ribbonTabs(ctx: RibbonContext): Tab[] {
  return [
    {
      id: "start", label: "Start", key: "R",
      shrink: ["styles", "voice", "editing", "font", "paragraph", "clipboard"],
      groups: [
        {
          id: "clipboard", label: "Zwischenablage", key: "ZL", icon: "paste",
          launcher: { cmd: "edit.clipboardPane", key: "WF" },
          controls: [
            split("edit.paste", "V", "large", [{ cmd: "edit.paste", label: "Ursprüngliche Formatierung beibehalten" }, { cmd: "edit.pasteText" }]),
            stack(b("edit.cut", "X"), b("edit.copy", "C"), b("edit.formatPainter", "FO")),
          ],
        },
        {
          id: "font", label: "Schriftart", key: "ZF", icon: "fonts",
          launcher: { cmd: "font.dialog", key: "ÄC" },
          controls: [{
            type: "rows", rows: [
              [{ type: "combo", cmd: "font.name", width: 153, key: "SC", value: ctx.fontName, options: ctx.fontOptions,
                optionStyle: (f) => `font-family:"${f.replace(/ \(.*\)$/, "")}";font-size:14px`, commit: ctx.setFont },
                { type: "combo", cmd: "font.size", width: 63, key: "HC", value: ctx.fontSize,
                  options: () => [{ items: ["8", "9", "10", "10,5", "11", "12", "14", "16", "18", "20", "22", "24", "26", "28", "36", "48", "72"] }],
                  commit: ctx.setSize }],
              [b("font.bold", "1"), b("font.italic", "2"),
                split("font.underline", "3", "small", [{ cmd: "underline.single" }, { cmd: "underline.double" }, { cmd: "underline.thick" },
                  { cmd: "underline.dotted" }, { cmd: "underline.dash" }, { cmd: "underline.wave" }]),
                b("font.strike", "4"), b("font.subscript", "S1"), b("font.superscript", "6"), "gap", b("font.clear", "Q")],
              [dd("font.effects", "D1"),
                { type: "color", cmd: "font.highlight", key: "GE", palette: "highlight", current: ctx.highlight, pick: ctx.setHighlight },
                { type: "color", cmd: "font.color", key: "AC", palette: "font", current: ctx.fontColor, pick: ctx.setFontColor },
                dd("font.case", "7", "small", undefined, [{ cmd: "case.sentence" }, { cmd: "case.lower" }, { cmd: "case.upper" }, { cmd: "case.title" }, { cmd: "case.toggle" }]),
                "gap", b("font.grow", "SV"), b("font.shrink", "HS")],
            ],
          }],
        },
        {
          id: "paragraph", label: "Absatz", key: "ZP", icon: "align-left",
          launcher: { cmd: "para.dialog", key: "BB" },
          controls: [{
            type: "rows", rows: [
              [split("para.bullets", "AA", "small", [], undefined, ctx.panels.bullets), split("para.numbering", "GN", "small", [], undefined, ctx.panels.numbers),
                panel("para.multilevel", "TL", "small", ctx.panels.multilevel), "gap",
                b("para.indentDec", "ÖI"), b("para.indentInc", "PI")],
              [b("para.alignLeft", "II"), b("para.alignCenter", "RZ"), b("para.alignRight", "RE"), b("para.justify", "OL"), "gap",
                dd("para.lineSpacing", "NZ", "small", undefined, [
                  { cmd: "spacing.1" }, { cmd: "spacing.1.15" }, { cmd: "spacing.1.5" }, { cmd: "spacing.2" }, { cmd: "spacing.2.5" }, { cmd: "spacing.3" }, "-",
                  { cmd: "para.dialog", label: "Zeilenabstandsoptionen..." }, "-",
                  { cmd: "spacing.toggleBefore" }, { cmd: "spacing.toggleAfter" }])],
              [{ type: "color", cmd: "para.shading", key: "NC", palette: "font", current: ctx.shading, pick: ctx.setShading },
                split("para.borders", "MR", "small", ctx.panels.borders), "gap", b("para.sort", "J"), b("para.showMarks", "8")],
            ],
          }],
        },
        {
          id: "styles", label: "Formatvorlagen", key: "ZS", icon: "fonts",
          launcher: { cmd: "style.pane", key: "RF" },
          controls: [{ type: "gallery", key: "K", visible: 4, items: GALLERY.map((g) => ({ cmd: g.cmd, preview: g.preview })) }],
        },
        {
          id: "editing", label: "Bearbeiten", key: "ZN", icon: "find",
          controls: [stack(
            split("edit.find", "UU", "medium", [{ cmd: "edit.find" }, { cmd: "edit.goto" }]),
            b("edit.replace", "TR", "medium", "Ersetzen"),
            { type: "dropdown", cmd: "edit.selectMenu", key: "ß", size: "medium", menu: [{ cmd: "edit.selectAll" }] },
          )],
        },
        {
          id: "voice", label: "Sprache", key: "Z2", icon: "dictate",
          controls: [split("edit.dictate", "D2", "large")],
        },
      ],
    },
    {
      id: "insert", label: "Einfügen", key: "I",
      shrink: ["text", "links", "headerFooter", "symbols", "illustrations", "pages", "comments", "tables"],
      groups: [
        { id: "pages", label: "Seiten", key: "ZP", icon: "cover-page",
          controls: [stack(dd("insert.coverPage", "EE", "medium"), b("insert.blankPage", "EL", "medium"), b("insert.pageBreak", "SU", "medium"))] },
        { id: "tables", label: "Tabellen", key: "ZB", icon: "table", controls: [panel("insert.table", "LA", "large", ctx.panels.insertTable)] },
        { id: "illustrations", label: "Illustrationen", key: "ZI", icon: "picture",
          controls: [dd("insert.pictures", "B1", "large", undefined, ctx.panels.pictures), stack(dd("insert.shapes", "2", "medium"), b("insert.chart", "I", "medium"))] },
        { id: "links", label: "Links", key: "ZL", icon: "link",
          controls: [stack(split("insert.link", "H", "medium"), b("insert.bookmark", "K", "medium"), b("insert.crossRef", "Q", "medium"))] },
        { id: "comments", label: "Kommentare", key: "Z1", icon: "comment", controls: [b("insert.comment", "N", "large")] },
        { id: "headerFooter", label: "Kopf- und Fußzeile", key: "ZH", icon: "header",
          controls: [stack(dd("insert.header", "OK", "medium", undefined, ctx.panels.header), dd("insert.footer", "LF", "medium", undefined, ctx.panels.footer),
            dd("insert.pageNumber", "T", "medium", undefined, ctx.panels.pageNumber))] },
        { id: "text", label: "Text", key: "ZT", icon: "text-box",
          controls: [dd("insert.textBox", "X", "large"), stack(b("insert.dateTime", "D"))] },
        { id: "symbols", label: "Symbole", key: "ZS", icon: "equation",
          controls: [stack(split("insert.equation", "A", "medium"), dd("insert.symbol", "SY", "medium"))] },
      ],
    },
    {
      id: "design", label: "Entwurf", key: "Ä",
      shrink: ["docFormat", "pageBackground"],
      groups: [
        { id: "docFormat", label: "Dokumentformatierung", key: "Z1", icon: "themes",
          controls: [dd("design.themes", "N", "large"), dd("design.colors", "F", "large"), dd("design.fonts", "6", "large"),
            stack(dd("design.paragraphSpacing", "BA", "medium"))] },
        { id: "pageBackground", label: "Seitenhintergrund", key: "ZB", icon: "page-color",
          controls: [dd("design.watermark", "W", "large"), dd("design.pageColor", "PC", "large"), b("design.pageBorders", "T", "large")] },
      ],
    },
    {
      id: "layout", label: "Layout", key: "S",
      shrink: ["arrange", "paragraph", "pageSetup"],
      groups: [
        { id: "pageSetup", label: "Seite einrichten", key: "ZS", icon: "margins", launcher: { cmd: "layout.pageSetup", key: "T" },
          controls: [panel("layout.margins", "8", "large", ctx.panels.margins),
            dd("layout.orientation", "4", "large", undefined, [{ cmd: "orientation.portrait" }, { cmd: "orientation.landscape" }]),
            panel("layout.size", "R", "large", ctx.panels.size),
            dd("layout.columns", "HS", "large", undefined, ctx.panels.columns),
            stack(dd("layout.breaks", "U", "medium", undefined, ctx.panels.breaks), dd("layout.hyphenation", "I", "medium"))] },
        { id: "paragraph", label: "Absatz", key: "ZG", icon: "line-spacing", launcher: { cmd: "para.dialog", key: "BB" },
          controls: [
            stack({ type: "spinner", cmd: "layout.indentLeft", label: "Links:", icon: "indent-increase", key: "L",
              value: () => ctx.indent("left"), commit: (v) => ctx.setIndent("left", v), step: (d) => ctx.stepIndent("left", d) },
            { type: "spinner", cmd: "layout.indentRight", label: "Rechts:", icon: "indent-decrease", key: "C",
              value: () => ctx.indent("right"), commit: (v) => ctx.setIndent("right", v), step: (d) => ctx.stepIndent("right", d) }),
            stack({ type: "spinner", cmd: "layout.spaceBefore", label: "Vor:", icon: "line-spacing", key: "V",
              value: () => ctx.spacing("before"), commit: (v) => ctx.setSpacing("before", v), step: (d) => ctx.stepSpacing("before", d) },
            { type: "spinner", cmd: "layout.spaceAfter", label: "Nach:", icon: "line-spacing", key: "N",
              value: () => ctx.spacing("after"), commit: (v) => ctx.setSpacing("after", v), step: (d) => ctx.stepSpacing("after", d) }),
          ] },
        { id: "arrange", label: "Anordnen", key: "ZA", icon: "position",
          controls: [dd("layout.position", "PP", "large"), dd("layout.wrapText", "J", "large", undefined, ctx.panels.wrap), split("layout.bringForward", "GE", "large"),
            split("layout.sendBackward", "GH", "large"), stack(dd("layout.align", "DA", "medium"), dd("layout.rotate", "DR", "medium"))] },
      ],
    },
    {
      id: "references", label: "Referenzen", key: "C",
      shrink: ["captions", "citations", "toc", "footnotes"],
      groups: [
        { id: "toc", label: "Inhaltsverzeichnis", key: "ZI", icon: "toc",
          controls: [dd("ref.toc", "I", "large"), stack(dd("ref.addText", "X", "medium"), b("ref.updateToc", "H", "medium", "Tabelle aktualisieren"))] },
        { id: "footnotes", label: "Fußnoten", key: "ZF", icon: "footnote", launcher: { cmd: "ref.footnoteDialog", key: "O" },
          controls: [b("ref.footnote", "P", "large"), stack(b("ref.endnote", "N", "medium"), split("ref.nextFootnote", "Ä", "medium"))] },
        { id: "citations", label: "Zitate und Literaturverzeichnis", key: "ZA", icon: "citation",
          controls: [dd("ref.citation", "Ü", "large"), stack(b("ref.sources", "V", "medium"), dd("ref.bibliography", "L", "medium"))] },
        { id: "captions", label: "Beschriftungen", key: "ZB", icon: "caption",
          controls: [b("ref.caption", "C", "large"), stack(b("ref.crossRef", "Q", "medium"))] },
      ],
    },
    {
      id: "review", label: "Überprüfen", key: "P",
      shrink: ["tracking", "comments", "language", "speech", "proofing"],
      groups: [
        { id: "proofing", label: "Rechtschreibung", key: "ZP", icon: "spelling",
          controls: [split("review.spelling", "S1", "large"), stack(b("review.thesaurus", "T", "medium"), b("review.wordCount", "W", "medium"))] },
        { id: "speech", label: "Sprache", key: "Z1", icon: "read-aloud", controls: [b("review.readAloud", "V1", "large")] },
        { id: "language", label: "Sprache", key: "ZL", icon: "translate",
          controls: [dd("review.translate", "Ü", "large"), dd("review.language", "SP", "large")] },
        { id: "comments", label: "Kommentare", key: "ZC", icon: "comment",
          controls: [b("review.newComment", "N", "large"), split("review.deleteComment", "L", "large"),
            b("review.prevComment", "Q", "large"), b("review.nextComment", "ß", "large")] },
        { id: "tracking", label: "Nachverfolgung", key: "ZT", icon: "track-changes", launcher: { cmd: "review.trackOptions", key: "C" },
          controls: [split("review.trackChanges", "DN", "large"), stack(dd("review.showMarkup", "G", "medium")),
            split("review.accept", "AA", "large"), split("review.reject", "FB", "large"),
            stack(b("review.prevChange", "U", "medium"), b("review.nextChange", "Ä", "medium"))] },
      ],
    },
    {
      id: "view", label: "Ansicht", key: "F",
      shrink: ["show", "zoom", "views"],
      groups: [
        { id: "views", label: "Ansichten", key: "ZV", icon: "print-layout",
          controls: [b("view.readMode", "E", "large"), b("view.printLayout", "D1", "large")] },
        { id: "show", label: "Anzeigen", key: "ZS", icon: "ruler",
          controls: [stack({ type: "check", cmd: "view.ruler", key: "L2" }, { type: "check", cmd: "view.navPane", key: "V2" })] },
        { id: "zoom", label: "Zoom", key: "ZZ", icon: "zoom",
          controls: [b("view.zoomDialog", "DO", "large"), b("view.zoom100", "0", "large"),
            stack(b("view.onePage", "1", "medium"), b("view.multiplePages", "2", "medium"), b("view.pageWidth", "Ö", "medium"))] },
      ],
    },
  ];
}
