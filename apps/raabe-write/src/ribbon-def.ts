// RaabeWrite's ribbon, following Word 365's German ribbon as captured on 2026-10-13
// (tools/feel-lab/capture_word_ui.py): same tabs, groups, order, sizes and keytips.
// Left out on purpose: Sendungen, Entwicklertools, Hilfe, and groups for features we do not
// plan (cloud, add-ins, SmartArt, 3D models, macros ...). See docs/SCOPE.md.

import type { Control, Tab } from "@raabe/shell";
import type { Editor } from "@raabe/text-engine";

const b = (cmd: string, key: string, size: "large" | "medium" | "small" = "small", label?: string): Control =>
  ({ type: "button", cmd, key, size, label });
const dd = (cmd: string, key: string, size: "large" | "medium" | "small" = "small", label?: string): Control =>
  ({ type: "dropdown", cmd, key, size, label, menu: [{ cmd }] });
const split = (cmd: string, key: string, size: "large" | "medium" | "small" = "small", menu: { cmd: string }[] = [{ cmd }], label?: string): Control =>
  ({ type: "split", cmd, key, size, menu, label });
const stack = (...items: Control[]): Control => ({ type: "stack", items });

export function ribbonTabs(active: () => Editor | null): Tab[] {
  const font = () => (active()?.currentStyle() === "h1" ? "Calibri Light" : "Calibri (Textkörper)");
  const size = () => (active()?.currentStyle() === "h1" ? "16" : "11");
  return [
    {
      id: "start", label: "Start", key: "R",
      shrink: ["styles", "voice", "editing", "font", "paragraph", "clipboard"],
      groups: [
        {
          id: "clipboard", label: "Zwischenablage", key: "ZL", icon: "paste",
          launcher: { cmd: "edit.clipboardPane", key: "WF" },
          controls: [
            split("edit.paste", "V", "large", [{ cmd: "edit.paste" }, { cmd: "edit.pasteText" }]),
            stack(b("edit.cut", "X"), b("edit.copy", "C"), b("edit.formatPainter", "FO")),
          ],
        },
        {
          id: "font", label: "Schriftart", key: "ZF", icon: "fonts",
          launcher: { cmd: "font.dialog", key: "ÄC" },
          controls: [{
            type: "rows", rows: [
              [{ type: "combo", cmd: "font.name", width: 153, key: "SC", value: font },
                { type: "combo", cmd: "font.size", width: 63, key: "HC", value: size }],
              [b("font.bold", "1"), b("font.italic", "2"), split("font.underline", "3"), b("font.strike", "4"),
                b("font.subscript", "S1"), b("font.superscript", "6"), "gap", b("font.clear", "Q")],
              [dd("font.effects", "D1"), split("font.highlight", "GE"), split("font.color", "AC"), dd("font.case", "7"),
                "gap", b("font.grow", "SV"), b("font.shrink", "HS")],
            ],
          }],
        },
        {
          id: "paragraph", label: "Absatz", key: "ZP", icon: "align-left",
          launcher: { cmd: "para.dialog", key: "BB" },
          controls: [{
            type: "rows", rows: [
              [split("para.bullets", "AA"), split("para.numbering", "GN"), dd("para.multilevel", "TL"), "gap",
                b("para.indentDec", "ÖI"), b("para.indentInc", "PI")],
              [b("para.alignLeft", "II"), b("para.alignCenter", "RZ"), b("para.alignRight", "RE"), b("para.justify", "OL"), "gap",
                dd("para.lineSpacing", "NZ")],
              [split("para.shading", "NC"), split("para.borders", "MR"), "gap", b("para.sort", "J"), b("para.showMarks", "8")],
            ],
          }],
        },
        {
          id: "styles", label: "Formatvorlagen", key: "ZS", icon: "fonts",
          launcher: { cmd: "style.pane", key: "RF" },
          controls: [{
            type: "gallery", key: "K", visible: 4, items: [
              { cmd: "style.normal", preview: "font:14.667px Calibri,Carlito,sans-serif" },
              { cmd: "style.noSpacing", preview: "font:14.667px Calibri,Carlito,sans-serif" },
              { cmd: "style.h1", preview: "font:21.333px 'Calibri Light',Carlito,sans-serif;color:#2f5496" },
              { cmd: "style.h2", preview: "font:17.333px 'Calibri Light',Carlito,sans-serif;color:#2f5496" },
              { cmd: "style.title", preview: "font:28px 'Calibri Light',Carlito,sans-serif" },
              { cmd: "style.subtitle", preview: "font:14.667px Calibri,Carlito,sans-serif;color:#5a5a5a" },
              { cmd: "style.quote", preview: "font:italic 14.667px Calibri,Carlito,sans-serif;color:#404040" },
            ],
          }],
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
        { id: "tables", label: "Tabellen", key: "ZB", icon: "table", controls: [dd("insert.table", "LA", "large")] },
        { id: "illustrations", label: "Illustrationen", key: "ZI", icon: "picture",
          controls: [dd("insert.pictures", "B1", "large"), stack(dd("insert.shapes", "2", "medium"), b("insert.chart", "I", "medium"))] },
        { id: "links", label: "Links", key: "ZL", icon: "link",
          controls: [stack(split("insert.link", "H", "medium"), b("insert.bookmark", "K", "medium"), b("insert.crossRef", "Q", "medium"))] },
        { id: "comments", label: "Kommentare", key: "Z1", icon: "comment", controls: [b("insert.comment", "N", "large")] },
        { id: "headerFooter", label: "Kopf- und Fußzeile", key: "ZH", icon: "header",
          controls: [stack(dd("insert.header", "OK", "medium"), dd("insert.footer", "LF", "medium"), dd("insert.pageNumber", "T", "medium"))] },
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
          controls: [dd("layout.margins", "8", "large"), dd("layout.orientation", "4", "large"), dd("layout.size", "R", "large"),
            dd("layout.columns", "HS", "large"), stack(dd("layout.breaks", "U", "medium"), dd("layout.hyphenation", "I", "medium"))] },
        { id: "paragraph", label: "Absatz", key: "ZG", icon: "line-spacing", launcher: { cmd: "para.dialog", key: "BB" },
          controls: [
            stack({ type: "spinner", cmd: "layout.indentLeft", label: "Links:", icon: "indent-increase", key: "L", value: "0 cm" },
              { type: "spinner", cmd: "layout.indentRight", label: "Rechts:", icon: "indent-decrease", key: "C", value: "0 cm" }),
            stack({ type: "spinner", cmd: "layout.spaceBefore", label: "Vor:", icon: "line-spacing", key: "V", value: "0 Pt." },
              { type: "spinner", cmd: "layout.spaceAfter", label: "Nach:", icon: "line-spacing", key: "N", value: "8 Pt." }),
          ] },
        { id: "arrange", label: "Anordnen", key: "ZA", icon: "position",
          controls: [dd("layout.position", "PP", "large"), dd("layout.wrapText", "J", "large"), split("layout.bringForward", "GE", "large"),
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
