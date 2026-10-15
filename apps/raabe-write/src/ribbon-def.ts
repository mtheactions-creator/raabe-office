// RaabeWrite's ribbon, following Word 365's German ribbon as captured on 2026-10-13
// (tools/feel-lab/capture_word_ui.py): same tabs, groups, order, sizes and keytips.
// Left out on purpose: Sendungen, Entwicklertools, Hilfe, and groups for features we do not
// plan (cloud, add-ins, SmartArt, 3D models, macros ...). See docs/SCOPE.md.

import { GROUP_MENU } from "./features/groups";
import { HYPHENATION_MENU } from "./features/hyphenation";
import type { Control, MenuEntry, Tab } from "@raabe/shell";

/** What the ribbon needs from the app to show values and apply changes. */
export interface RibbonContext {
  /** Referenzen > Formatvorlage (the document's citation style). */
  citationStyle: { value(): string; options(): { heading?: string; items: string[] }[]; commit(v: string): void };
  fontName(): string;
  /** Formatvorlagen: Livevorschau of the gallery entry under the mouse (null ends it). */
  styleHover?(cmd: string | null): void;
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
    columns: MenuEntry[]; breaks: MenuEntry[]; lineNumbers: MenuEntry[]; header: MenuEntry[]; footer: MenuEntry[]; pageNumber: MenuEntry[];
    pictures: MenuEntry[]; borders: MenuEntry[]; wrap: MenuEntry[]; shapes: Popup; textBox: MenuEntry[];
    position: Popup; align: Popup; rotate: MenuEntry[]; forward: MenuEntry[]; backward: MenuEntry[];
    toc: MenuEntry[]; addText: MenuEntry[]; nextNote: MenuEntry[]; equation: Popup; symbol: Popup; citation: MenuEntry[]; bibliography: Popup;
    track: MenuEntry[]; deleteComment: MenuEntry[]; accept: MenuEntry[]; reject: MenuEntry[];
    showComments: MenuEntry[]; filterMarkup: MenuEntry[]; reviewingPane: MenuEntry[];
    fontEffects: MenuEntry[]; themes: Popup; styleSet: Popup; colors: Popup; fonts: Popup; effects: Popup; paragraphSpacing: MenuEntry[]; watermark: Popup; pageColor: Popup; coverPage: Popup;
  };
  /** Markup für Überarbeitung anzeigen (Einfaches Markup, Alle Markups …). */
  markup: { value(): string; set(v: string): void; options(): { heading?: string; items: string[] }[] };
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
const panelL = (cmd: string, key: string, label: string, popup: Popup): Control =>
  ({ type: "dropdown", cmd, key, size: "large", label, menu: [], popup });

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

/**
 * Word's mini toolbar above the context menu (recorded 2026-10-14): font, size, bigger and
 * smaller, format painter, bold, italic, underline, colours and lists; in a table Einfügen and
 * Löschen; for a picture its styles and cropping.
 */
export function miniToolbar(ctx: RibbonContext, kind: "text" | "table" | "picture"): Control[] {
  if (kind === "picture") return [dd("picture.styles", "", "large", "Formatvorlage"), b("picture.crop", "", "large")];
  const fontName: Control = {
    type: "combo", cmd: "font.name", width: 104, value: ctx.fontName, options: ctx.fontOptions,
    optionStyle: (f) => `font-family:"${f.replace(/ \(.*\)$/, "")}";font-size:14px`, commit: ctx.setFont,
  };
  const fontSize: Control = {
    type: "combo", cmd: "font.size", width: 50, value: ctx.fontSize,
    options: () => [{ items: ["8", "9", "10", "10,5", "11", "12", "14", "16", "18", "20", "22", "24", "26", "28", "36", "48", "72"] }],
    commit: ctx.setSize,
  };
  const highlight: Control = { type: "color", cmd: "font.highlight", palette: "highlight", current: ctx.highlight, pick: ctx.setHighlight };
  const fontColor: Control = { type: "color", cmd: "font.color", palette: "font", current: ctx.fontColor, pick: ctx.setFontColor };
  const bullets = split("para.bullets", "", "small", [], undefined, ctx.panels.bullets);
  const numbering = split("para.numbering", "", "small", [], undefined, ctx.panels.numbers);
  const comment = b("review.newComment", "", "large");
  if (kind === "table") {
    const shading: Control = { type: "color", cmd: "para.shading", palette: "font", current: ctx.shading, pick: ctx.setShading };
    return [
      { type: "rows", rows: [
        [fontName, fontSize, b("font.grow", ""), b("font.shrink", ""), bullets, numbering],
        [b("font.bold", ""), b("font.italic", ""), b("para.alignCenter", ""), highlight, fontColor, shading, split("para.borders", "", "small", ctx.panels.borders), b("edit.formatPainter", "")],
      ] },
      dd("mini.tableInsert", "", "large", "Einfügen", [
        { cmd: "table.insertAbove", label: "Oberhalb einfügen" }, { cmd: "table.insertBelow", label: "Unterhalb einfügen" },
        { cmd: "table.insertLeft", label: "Links einfügen" }, { cmd: "table.insertRight", label: "Rechts einfügen" }]),
      dd("mini.tableDelete", "", "large", "Löschen", [
        { cmd: "table.deleteCells", label: "Zellen löschen..." }, { cmd: "table.deleteColumns" }, { cmd: "table.deleteRows" }, { cmd: "table.deleteTable" }]),
      comment,
    ];
  }
  return [
    { type: "rows", rows: [
      [fontName, fontSize, b("font.grow", ""), b("font.shrink", ""), b("edit.formatPainter", "")],
      [b("font.bold", ""), b("font.italic", ""), b("font.underline", ""), highlight, fontColor, bullets, numbering],
    ] },
    dd("mini.styles", "", "large", "Formatvorlagen", GALLERY.map((g) => ({ cmd: g.cmd }))),
    comment,
  ];
}

export function ribbonTabs(ctx: RibbonContext): Tab[] {
  // ----- shared controls ------------------------------------------------------------------
  const fontName = (width: number): Control => ({
    type: "combo", cmd: "font.name", width, key: "SC", value: ctx.fontName, options: ctx.fontOptions,
    optionStyle: (f) => `font-family:"${f.replace(/ \(.*\)$/, "")}";font-size:14px`, commit: ctx.setFont,
  });
  const fontSize: Control = {
    type: "combo", cmd: "font.size", width: 63, key: "HC", value: ctx.fontSize,
    options: () => [{ items: ["8", "9", "10", "10,5", "11", "12", "14", "16", "18", "20", "22", "24", "26", "28", "36", "48", "72"] }],
    commit: ctx.setSize,
  };
  const underline = split("font.underline", "3", "small", [{ cmd: "underline.single" }, { cmd: "underline.double" }, { cmd: "underline.thick" },
    { cmd: "underline.dotted" }, { cmd: "underline.dash" }, { cmd: "underline.wave" }]);
  const highlight: Control = { type: "color", cmd: "font.highlight", key: "GE", palette: "highlight", current: ctx.highlight, pick: ctx.setHighlight };
  const fontColor: Control = { type: "color", cmd: "font.color", key: "AC", palette: "font", current: ctx.fontColor, pick: ctx.setFontColor };
  const fontCase = dd("font.case", "7", "small", undefined, [{ cmd: "case.sentence" }, { cmd: "case.lower" }, { cmd: "case.upper" }, { cmd: "case.title" }, { cmd: "case.toggle" }]);
  const bullets = split("para.bullets", "AA", "small", [], undefined, ctx.panels.bullets);
  const numbering = split("para.numbering", "GN", "small", [], undefined, ctx.panels.numbers);
  const multilevel = panel("para.multilevel", "TL", "small", ctx.panels.multilevel);
  const lineSpacing = dd("para.lineSpacing", "NZ", "small", undefined, [
    { cmd: "spacing.1" }, { cmd: "spacing.1.15" }, { cmd: "spacing.1.5" }, { cmd: "spacing.2" }, { cmd: "spacing.2.5" }, { cmd: "spacing.3" }, "-",
    { cmd: "para.dialog", label: "Zeilenabstandsoptionen..." }, "-",
    { cmd: "spacing.toggleBefore" }, { cmd: "spacing.toggleAfter" }]);
  const shading: Control = { type: "color", cmd: "para.shading", key: "NC", palette: "font", current: ctx.shading, pick: ctx.setShading };
  const borders = split("para.borders", "MR", "small", ctx.panels.borders);
  const paste = split("edit.paste", "V", "large", [{ cmd: "edit.paste", label: "Ursprüngliche Formatierung beibehalten" }, { cmd: "edit.pasteText" }]);
  const find = (size: "medium" | "small") => split("edit.find", "UU", size, [{ cmd: "edit.find" }, { cmd: "edit.goto" }]);
  const select: Control = { type: "dropdown", cmd: "edit.selectMenu", key: "ß", size: "medium", menu: [{ cmd: "edit.selectAll" }, { cmd: "edit.selectObjects" }, { cmd: "edit.selectSimilar" }, { cmd: "layout.selectionPane", label: "Auswahlbereich..." }] };
  const spinners = (title: string, a: Control, b2: Control): Control => ({ type: "stack", title, items: [a, b2] });
  const indentLeft: Control = { type: "spinner", cmd: "layout.indentLeft", label: "Links:", icon: "indent-increase", key: "L",
    value: () => ctx.indent("left"), commit: (v) => ctx.setIndent("left", v), step: (d) => ctx.stepIndent("left", d) };
  const indentRight: Control = { type: "spinner", cmd: "layout.indentRight", label: "Rechts:", icon: "indent-decrease", key: "C",
    value: () => ctx.indent("right"), commit: (v) => ctx.setIndent("right", v), step: (d) => ctx.stepIndent("right", d) };
  const spaceBefore: Control = { type: "spinner", cmd: "layout.spaceBefore", label: "Vor:", icon: "line-spacing", key: "V",
    value: () => ctx.spacing("before"), commit: (v) => ctx.setSpacing("before", v), step: (d) => ctx.stepSpacing("before", d) };
  const spaceAfter: Control = { type: "spinner", cmd: "layout.spaceAfter", label: "Nach:", icon: "line-spacing", key: "N",
    value: () => ctx.spacing("after"), commit: (v) => ctx.setSpacing("after", v), step: (d) => ctx.stepSpacing("after", d) };
  const markupMode: Control = { type: "combo", cmd: "review.markupMode", width: 155, key: "TD", value: ctx.markup.value, options: ctx.markup.options, commit: ctx.markup.set };
  const showComments = split("review.showComments", "KA", "large", ctx.panels.showComments);
  const styleMenu: MenuEntry[] = GALLERY.map((g) => ({ cmd: g.cmd }));

  // Layouts and reduction sequences follow Word 365 at window widths from 2560 down to 800 px
  // (tools/feel-lab/capture_ribbon_widths.py, ribbon_variants.py; D-038).
  return [
    {
      id: "start", label: "Start", key: "R",
      shrink: ["clipboard", "font", "paragraph", "editing", "styles", "paragraph", "font", "voice"],
      groups: [
        {
          id: "clipboard", label: "Zwischenablage", key: "ZL", icon: "paste",
          launcher: { cmd: "edit.clipboardPane", key: "WF" },
          controls: [paste, stack(b("edit.cut", "X", "medium"), b("edit.copy", "C", "medium"), b("edit.formatPainter", "FO", "medium"))],
          smaller: [[paste, stack(b("edit.cut", "X"), b("edit.copy", "C"), b("edit.formatPainter", "FO"))]],
        },
        {
          id: "font", label: "Schriftart", key: "ZF", icon: "fonts",
          launcher: { cmd: "font.dialog", key: "ÄC" },
          controls: [{
            type: "rows", rows: [
              [fontName(104), fontSize, "gap", b("font.grow", "SV"), b("font.shrink", "HS"), "gap", fontCase, "gap", b("font.clear", "Q")],
              [b("font.bold", "1"), b("font.italic", "2"), underline, b("font.strike", "4"), b("font.subscript", "S1"), b("font.superscript", "6"), "gap",
                dd("font.effects", "D1", "small", undefined, ctx.panels.fontEffects), highlight, fontColor],
            ],
          }],
          smaller: [[{
            type: "rows", rows: [
              [fontName(153), fontSize],
              [b("font.bold", "1"), b("font.italic", "2"), underline, b("font.strike", "4"), b("font.subscript", "S1"), b("font.superscript", "6"), "gap", b("font.clear", "Q")],
              [dd("font.effects", "D1", "small", undefined, ctx.panels.fontEffects), highlight, fontColor, fontCase, "gap", b("font.grow", "SV"), b("font.shrink", "HS")],
            ],
          }]],
        },
        {
          id: "paragraph", label: "Absatz", key: "ZP", icon: "align-left",
          launcher: { cmd: "para.dialog", key: "BB" },
          controls: [{
            type: "rows", rows: [
              [bullets, numbering, multilevel, "gap", b("para.indentDec", "ÖI"), b("para.indentInc", "PI"), "gap", b("para.sort", "J"), "gap", b("para.showMarks", "8")],
              [b("para.alignLeft", "II"), b("para.alignCenter", "RZ"), b("para.alignRight", "RE"), b("para.justify", "OL"), "gap", lineSpacing, "gap", shading, borders],
            ],
          }],
          smaller: [[{
            type: "rows", rows: [
              [bullets, numbering, multilevel, "gap", b("para.indentDec", "ÖI"), b("para.indentInc", "PI")],
              [b("para.alignLeft", "II"), b("para.alignCenter", "RZ"), b("para.alignRight", "RE"), b("para.justify", "OL"), "gap", lineSpacing],
              [shading, borders, "gap", b("para.sort", "J"), b("para.showMarks", "8")],
            ],
          }]],
        },
        {
          id: "styles", label: "Formatvorlagen", key: "ZS", icon: "fonts",
          launcher: { cmd: "style.pane", key: "RF" },
          controls: [{ type: "gallery", key: "K", visible: 4, fill: true, items: GALLERY.map((g) => ({ cmd: g.cmd, preview: g.preview })), hover: ctx.styleHover ? (c) => ctx.styleHover!(c) : undefined }],
          smaller: [[{ type: "dropdown", cmd: "style.pane", key: "K", size: "large", label: "Formatvorlagen", menu: styleMenu }]],
        },
        {
          id: "editing", label: "Bearbeiten", key: "ZN", icon: "find",
          controls: [stack(find("medium"), b("edit.replace", "TR", "medium", "Ersetzen"), select)],
        },
        {
          id: "voice", label: "Sprache", key: "Z2", icon: "dictate",
          controls: [split("edit.dictate", "D2", "large")],
        },
      ],
    },
    {
      id: "insert", label: "Einfügen", key: "I",
      shrink: ["pages", "links", "symbols", "illustrations", "headerFooter", "text", "pages", "links", "illustrations", "symbols", "headerFooter", "text", "illustrations"],
      groups: [
        { id: "pages", label: "Seiten", key: "ZP", icon: "cover-page",
          controls: [panel("insert.coverPage", "EE", "large", ctx.panels.coverPage), b("insert.blankPage", "EL", "large"), b("insert.pageBreak", "SU", "large")],
          smaller: [[stack(panel("insert.coverPage", "EE", "medium", ctx.panels.coverPage), b("insert.blankPage", "EL", "medium"), b("insert.pageBreak", "SU", "medium"))]] },
        { id: "tables", label: "Tabellen", key: "ZB", icon: "table", controls: [panel("insert.table", "LA", "large", ctx.panels.insertTable)] },
        { id: "illustrations", label: "Illustrationen", key: "ZI", icon: "picture",
          controls: [dd("insert.pictures", "B1", "large", undefined, ctx.panels.pictures), panel("insert.shapes", "2", "large", ctx.panels.shapes), b("insert.chart", "I", "large")],
          smaller: [
            [dd("insert.pictures", "B1", "large", undefined, ctx.panels.pictures), stack(panel("insert.shapes", "2", "medium", ctx.panels.shapes), b("insert.chart", "I", "medium"))],
            [dd("insert.pictures", "B1", "large", undefined, ctx.panels.pictures), stack(panel("insert.shapes", "2", "medium", ctx.panels.shapes), b("insert.chart", "I"))],
          ] },
        { id: "links", label: "Links", key: "ZL", icon: "link",
          controls: [split("insert.link", "H", "large"), b("insert.bookmark", "K", "large"), b("insert.crossRef", "Q", "large")],
          smaller: [[stack(split("insert.link", "H", "medium"), b("insert.bookmark", "K", "medium"), b("insert.crossRef", "Q", "medium"))]] },
        { id: "comments", label: "Kommentare", key: "Z1", icon: "comment", controls: [b("insert.comment", "N", "large")] },
        { id: "headerFooter", label: "Kopf- und Fußzeile", key: "ZH", icon: "header",
          controls: [dd("insert.header", "OK", "large", undefined, ctx.panels.header), dd("insert.footer", "LF", "large", undefined, ctx.panels.footer),
            dd("insert.pageNumber", "T", "large", undefined, ctx.panels.pageNumber)],
          smaller: [[stack(dd("insert.header", "OK", "medium", undefined, ctx.panels.header), dd("insert.footer", "LF", "medium", undefined, ctx.panels.footer),
            dd("insert.pageNumber", "T", "medium", undefined, ctx.panels.pageNumber))]] },
        { id: "text", label: "Text", key: "ZT", icon: "text-box",
          controls: [dd("insert.textBox", "X", "large", undefined, ctx.panels.textBox), stack(b("insert.dateTime", "D", "medium"))],
          smaller: [[dd("insert.textBox", "X", "large", undefined, ctx.panels.textBox), stack(b("insert.dateTime", "D"))]] },
        { id: "symbols", label: "Symbole", key: "ZS", icon: "equation",
          controls: [split("insert.equation", "A", "large", [], undefined, ctx.panels.equation), panel("insert.symbol", "SY", "large", ctx.panels.symbol)],
          smaller: [[stack(split("insert.equation", "A", "medium", [], undefined, ctx.panels.equation), panel("insert.symbol", "SY", "medium", ctx.panels.symbol))]] },
      ],
    },
    {
      id: "design", label: "Entwurf", key: "Ä",
      shrink: ["docFormat", "pageBackground"],
      groups: [
        { id: "docFormat", label: "Dokumentformatierung", key: "Z1", icon: "themes",
          controls: [panel("design.themes", "N", "large", ctx.panels.themes), panel("design.styleSet", "J", "large", ctx.panels.styleSet),
            panel("design.colors", "F", "large", ctx.panels.colors), panel("design.fonts", "6", "large", ctx.panels.fonts),
            stack(dd("design.paragraphSpacing", "BA", "medium", undefined, ctx.panels.paragraphSpacing), panel("design.effects", "E", "medium", ctx.panels.effects), b("design.setDefault", "S", "medium"))] },
        { id: "pageBackground", label: "Seitenhintergrund", key: "ZB", icon: "page-color",
          controls: [panel("design.watermark", "W", "large", ctx.panels.watermark), panel("design.pageColor", "PC", "large", ctx.panels.pageColor), b("design.pageBorders", "T", "large")] },
      ],
    },
    {
      id: "layout", label: "Layout", key: "S",
      shrink: ["pageSetup", "arrange", "arrange", "paragraph", "pageSetup"],
      groups: [
        { id: "pageSetup", label: "Seite einrichten", key: "ZS", icon: "margins", launcher: { cmd: "layout.pageSetup", key: "T" },
          controls: [panelL("layout.margins", "8", "Seiten-\nränder", ctx.panels.margins),
            dd("layout.orientation", "4", "large", undefined, [{ cmd: "orientation.portrait" }, { cmd: "orientation.landscape" }]),
            panel("layout.size", "R", "large", ctx.panels.size),
            dd("layout.columns", "HS", "large", undefined, ctx.panels.columns),
            stack(dd("layout.breaks", "U", "medium", undefined, ctx.panels.breaks), dd("layout.lineNumbers", "LN", "medium", undefined, ctx.panels.lineNumbers), dd("layout.hyphenation", "I", "medium", undefined, HYPHENATION_MENU))],
          smaller: [[panelL("layout.margins", "8", "Seiten-\nränder", ctx.panels.margins),
            stack(dd("layout.orientation", "4", "medium", undefined, [{ cmd: "orientation.portrait" }, { cmd: "orientation.landscape" }]),
              panel("layout.size", "R", "medium", ctx.panels.size), dd("layout.columns", "HS", "medium", undefined, ctx.panels.columns)),
            stack(dd("layout.breaks", "U", "medium", undefined, ctx.panels.breaks), dd("layout.lineNumbers", "LN", "medium", undefined, ctx.panels.lineNumbers), dd("layout.hyphenation", "I", "medium", undefined, HYPHENATION_MENU))]] },
        { id: "paragraph", label: "Absatz", key: "ZG", icon: "line-spacing", launcher: { cmd: "para.dialog", key: "BB" },
          controls: [spinners("Einzug", indentLeft, indentRight), spinners("Abstand", spaceBefore, spaceAfter)] },
        { id: "arrange", label: "Anordnen", key: "ZA", icon: "position",
          controls: [panel("layout.position", "PP", "large", ctx.panels.position), dd("layout.wrapText", "J", "large", undefined, ctx.panels.wrap),
            split("layout.bringForward", "GE", "large", ctx.panels.forward), split("layout.sendBackward", "GH", "large", ctx.panels.backward), b("layout.selectionPane", "AP", "large"),
            stack(panel("layout.align", "DA", "medium", ctx.panels.align), dd("layout.group", "GG", "medium", undefined, GROUP_MENU), dd("layout.rotate", "DR", "medium", undefined, ctx.panels.rotate))],
          smaller: [[panel("layout.position", "PP", "large", ctx.panels.position), dd("layout.wrapText", "J", "large", undefined, ctx.panels.wrap),
            stack(split("layout.bringForward", "GE", "medium", ctx.panels.forward), split("layout.sendBackward", "GH", "medium", ctx.panels.backward), b("layout.selectionPane", "AP", "medium")),
            stack(panel("layout.align", "DA", "medium", ctx.panels.align), dd("layout.group", "GG", "medium", undefined, GROUP_MENU), dd("layout.rotate", "DR", "medium", undefined, ctx.panels.rotate))]] },
      ],
    },
    {
      id: "references", label: "Referenzen", key: "C",
      shrink: ["captions", "toc", "footnotes", "citations", "captions", "toc", "footnotes"],
      groups: [
        { id: "toc", label: "Inhaltsverzeichnis", key: "ZI", icon: "toc",
          controls: [dd("ref.toc", "I", "large", "Inhalts-\nverzeichnis", ctx.panels.toc),
            stack(dd("ref.addText", "X", "medium", undefined, ctx.panels.addText), b("ref.updateToc", "H", "medium"))],
          smaller: [[dd("ref.toc", "I", "large", "Inhalts-\nverzeichnis", ctx.panels.toc), stack(dd("ref.addText", "X", "medium", undefined, ctx.panels.addText), b("ref.updateToc", "H"))]] },
        { id: "footnotes", label: "Fußnoten", key: "ZF", icon: "footnote", launcher: { cmd: "ref.footnoteDialog", key: "O" },
          controls: [b("ref.footnote", "P", "large"), stack(b("ref.endnote", "N", "medium"), split("ref.nextFootnote", "Ä", "medium", ctx.panels.nextNote), b("ref.showNotes", "H", "medium"))],
          smaller: [[b("ref.footnote", "P", "large"), stack(b("ref.endnote", "N"), split("ref.nextFootnote", "Ä", "medium", ctx.panels.nextNote), b("ref.showNotes", "H"))]] },
        { id: "citations", label: "Zitate und Literaturverzeichnis", key: "ZA", icon: "citation",
          controls: [{ type: "dropdown", cmd: "ref.citation", key: "Ü", size: "large", get menu() { return ctx.panels.citation; } },
            stack(b("ref.sources", "V", "medium"),
              { type: "combo", cmd: "ref.citationStyle", width: 80, key: "L1", label: "Formatvorlage:", icon: "citation-style", value: ctx.citationStyle.value, options: ctx.citationStyle.options, commit: ctx.citationStyle.commit },
              panel("ref.bibliography", "L", "medium", ctx.panels.bibliography))] },
        { id: "captions", label: "Beschriftungen", key: "ZB", icon: "caption",
          controls: [b("ref.caption", "C", "large"), stack(b("ref.tableOfFigures", "G", "medium"), b("ref.updateTableOfFigures", "V1", "medium"), b("ref.crossRef", "Q", "medium"))],
          smaller: [[b("ref.caption", "C", "large"), stack(b("ref.tableOfFigures", "G"), b("ref.updateTableOfFigures", "V1"), b("ref.crossRef", "Q"))]] },
      ],
    },
    {
      id: "review", label: "Überprüfen", key: "P",
      shrink: ["proofing", "language", "comments", "tracking", "comments", "markup", "proofing", "speech"],
      groups: [
        { id: "proofing", label: "Rechtschreibung", key: "ZP", icon: "spelling",
          controls: [split("review.spelling", "S1", "large"), b("review.thesaurus", "T", "large"), b("review.wordCount", "W", "large")],
          smaller: [[stack(split("review.spelling", "S1", "medium"), b("review.thesaurus", "T", "medium"), b("review.wordCount", "W", "medium"))]] },
        { id: "speech", label: "Sprache", key: "Z1", icon: "read-aloud", controls: [b("review.readAloud", "V1", "large")] },
        { id: "language", label: "Sprache", key: "ZL", icon: "translate",
          controls: [dd("review.translate", "Ü", "large"), dd("review.language", "SP", "large", undefined, [{ cmd: "review.languageProofing" }, { cmd: "review.languageOptions" }])] },
        { id: "comments", label: "Kommentare", key: "ZC", icon: "comment",
          controls: [b("review.newComment", "N", "large"), split("review.deleteComment", "L", "large", ctx.panels.deleteComment),
            b("review.prevComment", "Q", "large"), b("review.nextComment", "ß", "large"), showComments],
          smaller: [[stack(b("review.newComment", "N", "medium"), split("review.deleteComment", "L", "medium", ctx.panels.deleteComment), b("review.prevComment", "Q", "medium")),
            stack(b("review.nextComment", "ß", "medium"), split("review.showComments", "KA", "medium", ctx.panels.showComments))]] },
        { id: "markup", label: "Markup", key: "ZM", icon: "markup", launcher: { cmd: "review.trackOptions", key: "C" },
          controls: [dd("review.filterMarkup", "F1", "large", undefined, ctx.panels.filterMarkup),
            stack(markupMode, dd("review.showMarkup", "G", "medium", undefined, ctx.panels.filterMarkup), split("review.reviewingPane", "U1", "medium", ctx.panels.reviewingPane))] },
        { id: "tracking", label: "Nachverfolgung", key: "ZT", icon: "track-changes",
          controls: [split("review.trackChanges", "DN", "large", ctx.panels.track),
            split("review.accept", "AA", "large", ctx.panels.accept), split("review.reject", "FB", "large", ctx.panels.reject),
            stack(b("review.prevChange", "U", "medium"), b("review.nextChange", "Ä", "medium"))] },
      ],
    },
    {
      id: "view", label: "Ansicht", key: "F",
      shrink: ["zoom", "show", "zoom", "views"],
      groups: [
        { id: "views", label: "Ansichten", key: "ZV", icon: "print-layout",
          controls: [b("view.readMode", "E", "large"), b("view.printLayout", "D1", "large"), b("view.webLayout", "W1", "large")] },
        { id: "show", label: "Anzeigen", key: "ZS", icon: "ruler",
          controls: [stack({ type: "check", cmd: "view.ruler", key: "L2" }, { type: "check", cmd: "view.gridlines", key: "G1" }, { type: "check", cmd: "view.navPane", key: "V2" })] },
        { id: "zoom", label: "Zoom", key: "ZZ", icon: "zoom",
          controls: [b("view.zoomDialog", "DO", "large"), b("view.zoom100", "0", "large"),
            stack(b("view.onePage", "1", "medium"), b("view.multiplePages", "2", "medium"), b("view.pageWidth", "Ö", "medium"))],
          smaller: [[b("view.zoomDialog", "DO", "large"), b("view.zoom100", "0", "large"),
            stack(b("view.onePage", "1"), b("view.multiplePages", "2"), b("view.pageWidth", "Ö"))]] },
      ],
    },
  ];
}
