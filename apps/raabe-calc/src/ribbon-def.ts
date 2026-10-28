// RaabeCalc's ribbon, following Excel 365's German ribbon as recorded on 2026-10-27
// (tools/feel-lab/capture_excel_ui.py, window of 1333 CSS px): same tabs, groups, order and
// keytips. Left out on purpose, as in RaabeWrite: Zeichnen, Automatisieren, Hilfe and add-in
// tabs, cloud and online features. See docs/SCOPE.md.

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
  fillColor(): string;
  setFillColor(hex: string): void;
  numberFormat(): string;
  numberFormats(): { heading?: string; items: string[] }[];
  setNumberFormat(label: string): void;
  preview: { font(name: string): void; size(value: string): void; color(hex: string): void; fill(hex: string): void };
  menus: {
    borders: MenuEntry[]; merge: MenuEntry[]; orientation: MenuEntry[]; accounting: MenuEntry[]; underline: MenuEntry[];
    conditional: MenuEntry[]; insertCells: MenuEntry[]; deleteCells: MenuEntry[]; format: MenuEntry[];
    autosum: MenuEntry[]; fill: MenuEntry[]; clear: MenuEntry[]; sortFilter: MenuEntry[]; find: MenuEntry[];
    paste: MenuEntry[]; margins: MenuEntry[]; orientationPage: MenuEntry[]; size: MenuEntry[]; printArea: MenuEntry[];
    freeze: MenuEntry[]; calcOptions: MenuEntry[]; notes: MenuEntry[]; names: MenuEntry[];
    fnLogical: MenuEntry[]; fnText: MenuEntry[]; fnDate: MenuEntry[]; fnLookup: MenuEntry[]; fnMath: MenuEntry[]; fnFinancial: MenuEntry[]; fnMore: MenuEntry[]; fnRecent: MenuEntry[];
    chartColumn: MenuEntry[]; chartLine: MenuEntry[]; chartPie: MenuEntry[]; chartScatter: MenuEntry[]; chartStat: MenuEntry[];
    fitWidth: MenuEntry[]; fitHeight: MenuEntry[];
  };
  panels: { cellStyles: Popup; tableStyles: Popup };
  scale: { value(): string; commit(v: string): void; step(d: 1 | -1): void };
}

const b = (cmd: string, key: string, size: "large" | "medium" | "small" = "small", label?: string): Control =>
  ({ type: "button", cmd, key, size, label });
const dd = (cmd: string, key: string, size: "large" | "medium" | "small", menu: MenuEntry[], label?: string): Control =>
  ({ type: "dropdown", cmd, key, size, label, menu });
const split = (cmd: string, key: string, size: "large" | "medium" | "small", menu: MenuEntry[], label?: string): Control =>
  ({ type: "split", cmd, key, size, menu, label });
const panel = (cmd: string, key: string, size: "large" | "medium" | "small", popup: Popup, label?: string): Control =>
  ({ type: "dropdown", cmd, key, size, label, menu: [], popup });
const stack = (...items: Control[]): Control => ({ type: "stack", items });
const check = (cmd: string, key: string): Control => ({ type: "check", cmd, key });

export function ribbonTabs(ctx: RibbonContext): Tab[] {
  const m = ctx.menus;
  const fontName: Control = {
    type: "combo", cmd: "font.name", width: 121, key: "SC", value: ctx.fontName, options: ctx.fontOptions,
    optionStyle: (f) => `font-family:"${f}";font-size:14px`, commit: ctx.setFont, preview: ctx.preview.font,
  };
  const fontSize: Control = {
    type: "combo", cmd: "font.size", width: 43, key: "HC", value: ctx.fontSize,
    options: () => [{ items: ["8", "9", "10", "11", "12", "14", "16", "18", "20", "22", "24", "26", "28", "36", "48", "72"] }],
    commit: ctx.setSize, preview: ctx.preview.size,
  };
  const numberFormat: Control = { type: "combo", cmd: "number.format", width: 91, key: "AZ", value: ctx.numberFormat, options: ctx.numberFormats, commit: ctx.setNumberFormat };
  const fontColor: Control = { type: "color", cmd: "font.color", key: "AC", palette: "font", current: ctx.fontColor, pick: ctx.setFontColor, preview: ctx.preview.color };
  const fill: Control = { type: "color", cmd: "font.fill", key: "H1", palette: "font", current: ctx.fillColor, pick: ctx.setFillColor, preview: ctx.preview.fill };
  const paste = split("edit.paste", "V", "large", m.paste);
  const autosum = split("edit.autosum", "KA", "small", m.autosum, "AutoSumme");

  // The shrink orders follow Excel's own at window widths from 2000 to 800 px (capture_excel_ui.py).
  return [
    {
      id: "start", label: "Start", key: "R",
      // Recorded at window widths from 2000 to 800 px: Formatvorlagen never collapses.
      shrink: ["editing", "cells", "number", "alignment", "font", "clipboard"],
      groups: [
        {
          id: "clipboard", label: "Zwischenablage", key: "ZL", icon: "paste",
          launcher: { cmd: "edit.clipboardPane", key: "WF" },
          controls: [paste, stack(b("edit.cut", "X"), split("edit.copy", "C", "small", [{ cmd: "edit.copy" }, { cmd: "edit.copyPicture" }]), b("edit.formatPainter", "FO"))],
        },
        {
          id: "font", label: "Schriftart", key: "ZF", icon: "fonts",
          launcher: { cmd: "font.dialog", key: "HZ" },
          controls: [{
            type: "rows", rows: [
              [fontName, fontSize, b("font.grow", "SV"), b("font.shrink", "J")],
              [b("font.bold", "1"), b("font.italic", "2"), split("font.underline", "3", "small", m.underline), "gap", split("border.menu", "RR", "small", m.borders), "gap", fill, fontColor],
            ],
          }],
        },
        {
          id: "alignment", label: "Ausrichtung", key: "ZA", icon: "align-left",
          launcher: { cmd: "align.dialog", key: "LZ" },
          controls: [{
            type: "rows", rows: [
              [b("align.top", "KB"), b("align.middle", "D"), b("align.bottom", "TN"), "gap", dd("align.orientation", "UA", "small", m.orientation), "gap", b("align.wrap", "NE", "small", "Textumbruch")],
              [b("align.left", "II"), b("align.center", "RZ"), b("align.right", "RE"), "gap", b("align.indentLess", "5"), b("align.indentMore", "6"), "gap", split("align.mergeCenter", "ß", "small", m.merge, "Verbinden und zentrieren")],
            ],
          }],
        },
        {
          id: "number", label: "Zahl", key: "ZN", icon: "number-format",
          launcher: { cmd: "number.dialog", key: "NZ" },
          controls: [{
            type: "rows", rows: [
              [numberFormat],
              [split("number.accounting", "N1", "small", m.accounting), b("number.percent", "P"), b("number.thousands", "Ö"), "gap", b("number.decimalMore", "0"), b("number.decimalLess", "9")],
            ],
          }],
        },
        {
          id: "styles", label: "Formatvorlagen", key: "ZS", icon: "cell-styles",
          controls: [stack(dd("styles.conditional", "8", "medium", m.conditional), panel("styles.table", "G", "medium", ctx.panels.tableStyles), panel("styles.cell", "M", "medium", ctx.panels.cellStyles))],
        },
        {
          id: "cells", label: "Zellen", key: "ZC", icon: "cells-insert",
          controls: [stack(split("cells.insert", "IE", "medium", m.insertCells, "Einfügen"), split("cells.delete", "EL", "medium", m.deleteCells, "Löschen"), dd("cells.format", "FF", "medium", m.format))],
        },
        {
          id: "editing", label: "Bearbeiten", key: "ZE", icon: "autosum",
          controls: [{
            type: "rows", rows: [
              [autosum],
              [dd("edit.fillDown", "Ü", "small", m.fill, "Ausfüllen")],
              [dd("edit.clearAll", "LÖ", "small", m.clear, "Löschen")],
            ],
          }, stack(dd("data.sort", "OS", "large", m.sortFilter, "Sortieren und Filtern")), stack(dd("edit.find", "UU", "large", m.find, "Suchen und Auswählen"))],
        },
      ],
    },
    {
      id: "insert", label: "Einfügen", key: "I",
      shrink: ["illustrations", "symbols", "text", "filters", "sparklines", "tables"],
      groups: [
        {
          id: "tables", label: "Tabellen", key: "ZB", icon: "table",
          controls: [b("insert.pivot", "V", "large"), b("insert.pivotRecommended", "FM", "large"), b("insert.table", "ß", "large")],
        },
        {
          id: "illustrations", label: "Illustrationen", key: "ZI", icon: "picture",
          controls: [stack(b("insert.pictures", "P"), b("insert.shapes", "SH"), b("insert.icons", "Y1"))],
        },
        {
          id: "charts", label: "Diagramme", key: "ZC", icon: "chart-column",
          launcher: { cmd: "insert.chartRecommended", key: "K2" },
          controls: [b("insert.chartRecommended", "LM", "large"), {
            type: "rows", rows: [
              [dd("insert.chartColumn", "D", "small", m.chartColumn), dd("insert.chartHierarchy", "E", "small", []), dd("insert.chartStat", "SA", "small", m.chartStat)],
              [dd("insert.chartLine", "IL", "small", m.chartLine), dd("insert.chartPie", "KK", "small", m.chartPie), dd("insert.chartScatter", "PU", "small", m.chartScatter)],
              [dd("insert.chartOther", "W1", "small", []), dd("insert.chartCombo", "N1", "small", []), b("insert.maps", "K1")],
            ],
          }, b("insert.pivotChart", "II", "large")],
        },
        {
          id: "sparklines", label: "Sparklines", key: "ZK", icon: "chart-line",
          controls: [stack(b("insert.sparkLine", "LL"), b("insert.sparkColumn", "Q"), b("insert.sparkWinLoss", "GG"))],
        },
        {
          id: "filters", label: "Filter", key: "ZF", icon: "filter",
          controls: [stack(b("insert.slicer", "91"), b("insert.timeline", "Ö"))],
        },
        { id: "links", label: "Links", key: "ZL", icon: "link", controls: [b("insert.link", "H", "large")] },
        { id: "comments", label: "Kommentare", key: "Z2", icon: "comment", controls: [b("insert.comment", "N2", "large")] },
        {
          id: "text", label: "Text", key: "ZT", icon: "text-box",
          controls: [stack(b("insert.textBox", "X"), b("insert.headerFooter", "H1"), b("insert.wordArt", "W"))],
        },
        { id: "symbols", label: "Symbole", key: "ZS", icon: "symbol", controls: [b("insert.equation", "E2", "large"), b("insert.symbol", "SY", "large")] },
      ],
    },
    {
      id: "layout", label: "Seitenlayout", key: "S",
      shrink: ["arrange", "sheetOptions", "scale", "themes"],
      groups: [
        {
          id: "themes", label: "Designs", key: "ZT", icon: "themes",
          controls: [b("layout.themes", "N", "large"), stack(b("layout.colors", "F"), b("layout.fonts", "6"), b("layout.effects", "E"))],
        },
        {
          id: "pageSetup", label: "Seite einrichten", key: "ZS", icon: "margins",
          launcher: { cmd: "layout.pageSetup", key: "K" },
          controls: [dd("layout.margins", "8", "large", m.margins), dd("layout.orientation", "4", "large", m.orientationPage), dd("layout.size", "R", "large", m.size),
            dd("layout.printArea", "H", "large", m.printArea), b("layout.breaks", "Ü", "large"), b("layout.background", "M", "large"), b("layout.printTitles", "I", "large")],
        },
        {
          id: "scale", label: "An Format anpassen", key: "ZC", icon: "zoom",
          launcher: { cmd: "layout.pageSetup", key: "K" },
          controls: [stack(
            dd("layout.fitWidth", "P", "small", m.fitWidth),
            dd("layout.fitHeight", "O", "small", m.fitHeight),
            { type: "spinner", cmd: "layout.scale", label: "Skalierung:", key: "L", value: ctx.scale.value, commit: ctx.scale.commit, step: ctx.scale.step },
          )],
        },
        {
          id: "sheetOptions", label: "Blattoptionen", key: "ZO", icon: "grid",
          launcher: { cmd: "layout.pageSetup", key: "B" },
          controls: [
            { type: "stack", title: "Gitternetzlinien", items: [check("layout.gridlinesView", "YX"), check("layout.gridlinesPrint", "YU")] },
            { type: "stack", title: "Überschriften", items: [check("layout.headingsView", "YA"), check("layout.headingsPrint", "YC")] },
          ],
        },
        {
          id: "arrange", label: "Anordnen", key: "ZA", icon: "bring-forward",
          controls: [b("layout.forward", "GE", "large"), b("layout.backward", "GH", "large"), b("layout.selectionPane", "Q", "large"),
            stack(b("layout.align", "DA"), b("layout.group", "GR"), b("layout.rotate", "DR"))],
        },
      ],
    },
    {
      id: "formulas", label: "Formeln", key: "O",
      shrink: ["names", "auditing", "library"],
      groups: [
        {
          id: "library", label: "Funktionsbibliothek", key: "ZF", icon: "fx",
          controls: [b("formulas.insertFunction", "U", "large"), split("edit.autosum", "KA", "large", m.autosum, "AutoSumme"),
            dd("formulas.recent", "Ä", "large", m.fnRecent), dd("formulas.financial", "I", "large", m.fnFinancial), dd("formulas.logical", "L", "large", m.fnLogical),
            dd("formulas.text", "X", "large", m.fnText), dd("formulas.dateTime", "9", "large", m.fnDate), dd("formulas.lookup", "C", "large", m.fnLookup),
            dd("formulas.math", "T", "large", m.fnMath), dd("formulas.more", "M", "large", m.fnMore)],
        },
        {
          id: "names", label: "Definierte Namen", key: "ZD", icon: "names",
          controls: [b("formulas.nameManager", "N", "large"), stack(b("formulas.defineName", "MD"), b("formulas.createFromSelection", "A"))],
        },
        {
          id: "auditing", label: "Formelüberwachung", key: "ZO", icon: "fx",
          controls: [stack(b("formulas.tracePrecedents", "R"), b("formulas.traceDependents", "D"), b("formulas.removeArrows", "AA")), stack(b("formulas.showFormulas", "H"), b("formulas.errorCheck", "K"), b("formulas.evaluate", "V"))],
        },
        {
          id: "calculation", label: "Berechnung", key: "ZB", icon: "calculate",
          controls: [dd("formulas.calcOptions", "Q", "large", m.calcOptions), stack(b("formulas.calcNow", "N2"), b("formulas.calcSheet", "B"))],
        },
      ],
    },
    {
      id: "data", label: "Daten", key: "V",
      shrink: ["outline", "tools", "forecast", "sortFilter"],
      groups: [
        {
          id: "get", label: "Daten abrufen und transformieren", key: "Z1", icon: "document",
          controls: [b("data.get", "D1", "large"), stack(b("data.fromText", "FT"), b("data.fromWeb", "M"))],
        },
        { id: "queries", label: "Abfragen & Verbindungen", key: "ZC", icon: "reapply", controls: [b("data.refreshAll", "K1", "large")] },
        {
          id: "sortFilter", label: "Sortieren und Filtern", key: "ZS", icon: "sort-az",
          controls: [stack(b("data.sortAsc", "U"), b("data.sortDesc", "F1")), b("data.sort", "S", "large"), b("data.filter", "1", "large"),
            stack(b("data.filterClear", "L"), b("data.filterReapply", "E"), b("data.advancedFilter", "W"))],
        },
        {
          id: "tools", label: "Datentools", key: "ZD", icon: "text-columns",
          controls: [b("data.textToColumns", "T", "large"), stack(b("edit.flashFill", "B"), b("data.removeDuplicates", "4"), split("data.validation", "Ü", "small", [{ cmd: "data.validation" }, { cmd: "data.circleInvalid" }, { cmd: "data.clearCircles" }])), b("data.consolidate", "8", "large")],
        },
        { id: "forecast", label: "Prognose", key: "Z3", icon: "trend", controls: [b("data.whatIf", "5", "large"), b("data.forecast", "3", "large")] },
        { id: "outline", label: "Gliederung", key: "ZO", icon: "outline", controls: [b("data.group", "G", "large"), b("data.ungroup", "H", "large"), b("data.subtotal", "B", "large")] },
      ],
    },
    {
      id: "review", label: "Überprüfen", key: "P",
      shrink: ["notes", "protect", "comments", "proofing"],
      groups: [
        { id: "proofing", label: "Rechtschreibung", key: "ZP", icon: "spelling", controls: [b("review.spelling", "BE", "large"), stack(b("review.thesaurus", "T"), b("review.statistics", "S1"))] },
        { id: "accessibility", label: "Barrierefreiheit", key: "Z2", icon: "check", controls: [b("review.accessibility", "B1", "large")] },
        { id: "language", label: "Sprache", key: "ZL", icon: "translate", controls: [b("review.translate", "Ü", "large")] },
        { id: "comments", label: "Kommentare", key: "ZC", icon: "comment", controls: [b("review.newComment", "N1", "large")] },
        { id: "notes", label: "Notizen", key: "Z4", icon: "note", controls: [dd("review.notes", "N3", "large", m.notes)] },
        { id: "protect", label: "Schützen", key: "Z5", icon: "protect", controls: [b("review.protectSheet", "S2", "large"), b("review.protectBook", "E", "large")] },
      ],
    },
    {
      id: "view", label: "Ansicht", key: "F",
      shrink: ["show", "views", "window", "zoom"],
      groups: [
        { id: "views", label: "Arbeitsmappenansichten", key: "ZV", icon: "document", controls: [b("view.normal", "C", "large"), b("view.pageBreak", "V1", "large"), b("view.pageLayout", "I", "large")] },
        {
          id: "show", label: "Anzeigen", key: "ZS", icon: "eye",
          controls: [stack(check("view.gridlines", "X"), check("view.formulaBar", "B"), check("view.headings", "A1"))],
        },
        { id: "zoom", label: "Zoom", key: "ZZ", icon: "zoom", controls: [b("view.zoomDialog", "DO", "large", "Zoom"), b("view.zoom100", "0", "large"), b("view.zoomSelection", "W", "large")] },
        {
          id: "window", label: "Fenster", key: "ZW", icon: "freeze",
          controls: [b("view.newWindow", "N2", "large"), b("view.arrange", "A2", "large"), dd("view.freeze", "P", "large", m.freeze), stack(b("view.split", "T1"))],
        },
      ],
    },
  ];
}
