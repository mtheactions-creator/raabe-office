// The contextual tabs as PowerPoint 365 shows them (D-070; groups, labels and keytips recorded in
// results/2026-11-06-powerpoint-inventory.json): Formformat, Bildformat, Tabellenentwurf, Tabellenlayout,
// Diagrammentwurf and Format. Their commands are registered here with PowerPoint's labels and attached
// in app.ts / contextcmds.ts; those without behaviour stay greyed with "Noch nicht verfügbar".

import type { Command, Control, Group, MenuEntry, Tab } from "@raabe/shell";

type C = Omit<Command, "id">;
const cmd = (id: string, label: string, icon?: string, extra: Partial<C> = {}): Command => ({ id, label, icon, ...extra });

export const CONTEXT_COMMANDS: Command[] = [
  cmd("shape.editShape", "Form bearbeiten", "edit-shape"),
  cmd("shape.merge", "Formen zusammenführen", "shapes"),
  cmd("shape.quickStyles", "Schnellformatvorlagen", "styles"),
  cmd("wordart.quickStyles", "Schnellformatvorlagen", "wordart"),
  cmd("wordart.fill", "Textfüllung", "text-fill"),
  cmd("wordart.outline", "Textkontur", "text-outline"),
  cmd("wordart.effects", "Texteffekte", "text-effects"),
  cmd("wordart.dialog", "Texteffekte formatieren...", "launcher"),
  cmd("shape.altTextPane", "Alternativtext", "alt-text"),
  cmd("shape.heightBox", "Formenhöhe", "shape-height"),
  cmd("shape.widthBox", "Formenbreite", "shape-width"),
  cmd("shape.sizeDialog", "Größe und Position...", "launcher"),
  cmd("smartart.edit2d", "In 2D bearbeiten", "edit-shape"),
  cmd("smartart.changeShape", "Form ändern", "edit-shape"),
  cmd("smartart.larger", "Größer", "font-grow"),
  cmd("smartart.smaller", "Kleiner", "font-shrink"),
  cmd("smartart.shapeStyles", "Schnellformatvorlagen", "styles"),
  cmd("pic.edit", "Bild bearbeiten", "edit"),
  cmd("pic.removeBg", "Hintergrund entfernen", "remove-background"),
  cmd("pic.corrections", "Korrekturen", "corrections"),
  cmd("pic.color", "Farbe", "picture-color"),
  cmd("pic.artistic", "Künstlerische Effekte", "artistic"),
  cmd("pic.transparency", "Transparenz", "transparency"),
  cmd("pic.compress", "Bilder komprimieren...", "compress"),
  cmd("pic.quickStyles", "Schnellformatvorlagen", "picture-styles"),
  cmd("pic.border", "Bildrahmen", "picture-border"),
  cmd("pic.resetSize", "Bild und Größe zurücksetzen", "reset-picture"),
  cmd("pic.changeClipboard", "Aus Zwischenablage...", "paste"),
  cmd("pic.effects", "Bildeffekte", "picture-effects"),
  cmd("pic.layout", "Bildlayout", "picture-layout"),
  cmd("pic.dialog", "Bild...", "launcher"),
  cmd("table.headerRow", "Überschrift", "check"),
  cmd("table.styles", "Tabellenformate", "table-style"),
  cmd("table.shading", "Schattierung", "shading"),
  cmd("table.borders", "Rahmen", "borders"),
  cmd("table.effects", "Effekte", "effects"),
  cmd("table.penStyle", "Stiftart", "line-style"),
  cmd("chart.elements", "Diagrammelemente", "chart"),
  cmd("shape.changeShapeChart", "Form ändern", "edit-shape"),
  cmd("table.penWeight", "Stiftstärke", "line-weight-pen"),
  cmd("table.penColor", "Stiftfarbe", "pen"),
  cmd("table.draw", "Tabelle zeichnen", "table-draw"),
  cmd("table.eraser", "Radierer", "table-eraser"),
  cmd("table.select", "Auswählen", "table-select"),
  cmd("table.gridlines", "Gitternetzlinien anzeigen", "gridlines"),
  cmd("table.deleteMenu", "Löschen", "table-delete"),
  cmd("table.rowAbove", "Zeile Oberhalb", "table-above"),
  cmd("table.rowBelow", "Zeile unterhalb einfügen", "table-below"),
  cmd("table.colLeft", "Spalte Links", "table-left"),
  cmd("table.colRight", "Spalte rechts einfügen", "table-right"),
  cmd("table.rowHeight", "Tabellenzeilenhöhe", "shape-height"),
  cmd("table.colWidth", "Tabellenspaltenbreite", "shape-width"),
  cmd("table.distRows", "Zeilen verteilen", "table-rows-even"),
  cmd("table.distCols", "Spalten verteilen", "table-cols-even"),
  cmd("table.alignLeft", "Linksbündig", "align-left"),
  cmd("table.alignCenter", "Zentriert", "align-center"),
  cmd("table.alignRight", "Rechtsbündig", "align-right"),
  cmd("table.top", "Oben ausrichten", "cell-top-center"),
  cmd("table.middle", "Vertikal zentrieren", "cell-middle-center"),
  cmd("table.bottom", "Unten ausrichten", "cell-bottom-center"),
  cmd("table.direction", "Textrichtung", "text-direction"),
  cmd("table.margins", "Zellenbegrenzungen", "cell-margins"),
  cmd("table.customMargins", "Benutzerdefinierte Seitenränder...", "cell-margins", { deferred: true }),
  cmd("table.height", "Höhe", "shape-height"),
  cmd("table.width", "Breite", "shape-width"),
  cmd("table.lockRatio", "Seitenverhältnis sperren", "check"),
  cmd("table.arrange", "Anordnen", "bring-forward"),
  cmd("chart.addElement", "Diagrammelement hinzufügen", "chart-recommended"),
  cmd("chart.quickLayout", "Schnelllayout", "chart-recommended"),
  cmd("chart.colors", "Farben ändern", "colors"),
  cmd("chart.styles", "Schnellformatvorlagen", "styles"),
  cmd("chart.switchRowCol", "Zeile/Spalte tauschen", "table"),
  cmd("chart.selectData", "Daten auswählen...", "table"),
  cmd("chart.refresh", "Daten aktualisieren", "update"),
];

const b = (cmd: string, key: string, size: "large" | "medium" | "small" = "small"): Control => ({ type: "button", cmd, key, size });
const dd = (cmd: string, key: string, size: "large" | "medium" | "small", menu: MenuEntry[] | (() => MenuEntry[])): Control =>
  typeof menu === "function" ? { type: "dropdown", cmd, key, size, get menu() { return menu(); } } : { type: "dropdown", cmd, key, size, menu };
const stack = (...items: Control[]): Control => ({ type: "stack", items });
const check = (cmd: string, key: string): Control => ({ type: "check", cmd, key });

export interface ContextTabHost {
  shapeGallery(close: () => void): HTMLElement;
  color(cur: () => string, pick: (hex: string) => void, cmd: string, key: string, label?: boolean, extra?: () => MenuEntry[]): Control;
  outlineExtra(): MenuEntry[];
  quickStyles(close: () => void): HTMLElement;
  /** Formformat: shapes and styles as galleries in the ribbon (PowerPoint). */
  shapeTiles(cols?: 3 | 6): Control;
  quickStyleTiles(visible?: number): Control;
  picStyleTiles(visible?: number): Control;
  tableStyleTiles(): Control;
  setShading(hex: string | null): void; shadingExtra(): MenuEntry[]; tableEffects(): MenuEntry[];
  wordArtTiles(): Control;
  cellDirections(): MenuEntry[];
  cellMargins(): MenuEntry[];
  tableBorders(): MenuEntry[];
  textEffects(): MenuEntry[];
  chartElements(): MenuEntry[];
  editShapeMenu(): MenuEntry[];
  tableStyles(close: () => void): HTMLElement;
  picStyles(close: () => void): HTMLElement;
  textOutlineColor(): string; setTextOutline(hex: string): void; textOutlineExtra(): MenuEntry[];
  fillColor(): string; setFill(hex: string): void;
  lineColor(): string; setLine(hex: string): void;
  fontColor(): string; setFontColor(hex: string): void; fontColorExtra(): MenuEntry[]; textFillExtra(): MenuEntry[];
  effects(): MenuEntry[];
  alignMenu(): MenuEntry[];
  /** `labelled`: "Höhe:" / "Breite:" beside the box (the widest layout). */
  sizeSpin(kind: "h" | "w", labelled?: boolean): Control;
  pictureColor(): MenuEntry[];
  pictureTransparency(): MenuEntry[];
  pictureCorrections(): MenuEntry[];
  cropMenu(): MenuEntry[];
  formatChartElement(name: string): void;
  pictureArtistic(): MenuEntry[];
  lineWeights(): MenuEntry[];
  fillExtra(): MenuEntry[];
  penStyles(): MenuEntry[]; penWeights(): MenuEntry[]; penColor(): string; setPenColor(c: string): void;
  chartKinds(): MenuEntry[];
  quickLayouts(): MenuEntry[];
  chartColors(): MenuEntry[];
  /** Diagrammentwurf > Diagrammformatvorlagen in the ribbon (D-187). */
  chartStyleTiles(): Control;
  tableSpin(kind: "row" | "col" | "h" | "w", labelled?: boolean): Control;
}

const arrange = (pic: boolean): Control[] => [stack(
  { type: "split", cmd: "arrange.forward", key: "GE", size: "medium", menu: [{ cmd: "arrange.forward" }, { cmd: "arrange.front" }] },
  { type: "split", cmd: "arrange.backward", key: "GH", size: "medium", menu: [{ cmd: "arrange.backward" }, { cmd: "arrange.back" }] },
  b("arrange.selectionPane", pic ? "Q2" : "Q", "medium")), stack(
  { type: "dropdown", cmd: "arrange.align", key: "DA", size: "small", menu: [{ cmd: "arrange.alignLeft" }, { cmd: "arrange.alignCenter" }, { cmd: "arrange.alignRight" }, "-", { cmd: "arrange.alignTop" }, { cmd: "arrange.alignMiddle" }, { cmd: "arrange.alignBottom" }, "-", { cmd: "arrange.distH" }, { cmd: "arrange.distV" }, "-", { cmd: "arrange.toSlide" }, { cmd: "arrange.toSelected" }] },
  { type: "dropdown", cmd: "arrange.group", key: "GR", size: "small", menu: [{ cmd: "arrange.group" }, { cmd: "arrange.regroup" }, { cmd: "arrange.ungroup" }] },
  { type: "dropdown", cmd: "arrange.rotate", key: "DR", size: "small", menu: [{ cmd: "arrange.rotRight" }, { cmd: "arrange.rotLeft" }, { cmd: "arrange.flipV" }, { cmd: "arrange.flipH" }, "-", { cmd: "shape.sizePosition", label: "Weitere Drehungsoptionen..." }] })];

/** Anordnen at a wide window: six large buttons, as PowerPoint's (D-186); `arrange` is the reduction after it. */
const arrangeWide = (pic: boolean): Control[] => [
  { type: "split", cmd: "arrange.forward", key: "GE", size: "large", menu: [{ cmd: "arrange.forward" }, { cmd: "arrange.front" }] },
  { type: "split", cmd: "arrange.backward", key: "GH", size: "large", menu: [{ cmd: "arrange.backward" }, { cmd: "arrange.back" }] },
  { type: "button", cmd: "arrange.selectionPane", key: pic ? "Q2" : "Q", size: "large", label: "Auswahlbereich" },
  { type: "dropdown", cmd: "arrange.align", key: "DA", size: "large", menu: [{ cmd: "arrange.alignLeft" }, { cmd: "arrange.alignCenter" }, { cmd: "arrange.alignRight" }, "-", { cmd: "arrange.alignTop" }, { cmd: "arrange.alignMiddle" }, { cmd: "arrange.alignBottom" }, "-", { cmd: "arrange.distH" }, { cmd: "arrange.distV" }, "-", { cmd: "arrange.toSlide" }, { cmd: "arrange.toSelected" }] },
  { type: "dropdown", cmd: "arrange.group", key: "GR", size: "large", menu: [{ cmd: "arrange.group" }, { cmd: "arrange.regroup" }, { cmd: "arrange.ungroup" }] },
  { type: "dropdown", cmd: "arrange.rotate", key: "DR", size: "large", menu: [{ cmd: "arrange.rotRight" }, { cmd: "arrange.rotLeft" }, { cmd: "arrange.flipV" }, { cmd: "arrange.flipH" }, "-", { cmd: "shape.sizePosition", label: "Weitere Drehungsoptionen..." }] }];
const arrangeGroup = (key: string, pic: boolean): Group => ({ id: "arrange", label: "Anordnen", key, icon: "bring-forward", controls: arrangeWide(pic), smaller: [arrange(pic)] });

/** Größe: "Höhe:" and "Breite:" beside the boxes at a wide window (PowerPoint), the icons only after that (D-186). */
const sizeGroup = (h: ContextTabHost, key: string, launcher: string, before: Control[] = []): Group => ({
  id: "size", label: "Größe", key, icon: before.length ? "crop" : "size-position", launcher: { cmd: launcher, key: "ß" },
  controls: [...before, stack(h.sizeSpin("h", true), h.sizeSpin("w", true))], smaller: [[...before, stack(h.sizeSpin("h"), h.sizeSpin("w"))]] });

/** Formenarten in Formformat and the chart's Format tab: seven styles at a wide window, three after that (D-186). */
const shapeStyles = (h: ContextTabHost, smaller: Control[][] = []): Group => {
  const tools = () => stack(h.color(() => h.fillColor(), (c) => h.setFill(c), "draw.fill", "Ü", true, () => h.fillExtra()), h.color(() => h.lineColor(), (c) => h.setLine(c), "draw.outline", "V", true, () => h.outlineExtra()), dd("draw.effects", "MO", "medium", () => h.effects()));
  return { id: "styles", label: "Formenarten", key: "ZF", icon: "shape-fill", launcher: { cmd: "draw.format", key: "BB" }, controls: [h.quickStyleTiles(7), tools()],
    smaller: [[h.quickStyleTiles(3), tools()], ...smaller] };
};

/**
 * WordArt-Formate in Tabellenentwurf: PowerPoint shows the styles as one large button there and the three text buttons
 * as icons, even at a wide window (D-186).
 */
const wordArtCompact = (h: ContextTabHost): Control[] => {
  const tiles = h.wordArtTiles() as Extract<Control, { type: "tiles" }>;
  return [{ type: "dropdown", cmd: "wordart.quickStyles", key: "C", size: "large", menu: [], popup: (close) => tiles.panel(close) },
    stack(h.color(() => h.fontColor(), (c) => h.setFontColor(c), "wordart.fill", "4", false, () => h.textFillExtra()), h.color(() => h.textOutlineColor(), (c) => h.setTextOutline(c), "wordart.outline", "TE", false, () => h.textOutlineExtra()), dd("wordart.effects", "IE", "small", () => h.textEffects()))];
};

const wordArt = (h: ContextTabHost): Control[] => [h.wordArtTiles(),
  stack(h.color(() => h.fontColor(), (c) => h.setFontColor(c), "wordart.fill", "4", true, () => h.textFillExtra()), h.color(() => h.textOutlineColor(), (c) => h.setTextOutline(c), "wordart.outline", "TE", true, () => h.textOutlineExtra()), dd("wordart.effects", "IE", "small", () => h.textEffects()))];

export function shapeFormatTab(h: ContextTabHost): Tab {
  return {
    id: "shapeFormat", label: "Formformat", key: "JD", contextual: true, shrink: [],
    groups: [
      { id: "insert", label: "Formen einfügen", key: "ZE", icon: "shapes", controls: [h.shapeTiles(), stack(dd("shape.editShape", "F", "small", () => h.editShapeMenu()), b("insert.textBox", "LT"), dd("shape.merge", "Y", "small", []))] },
      shapeStyles(h),
      { id: "wordart", label: "WordArt-Formate", key: "ZW", icon: "wordart", launcher: { cmd: "wordart.dialog", key: "O" }, controls: wordArt(h) },
      { id: "access", label: "Barrierefreiheit", key: "ZB", icon: "alt-text", controls: [b("shape.altTextPane", "AT", "large")] },
      arrangeGroup("ZA", false),
      sizeGroup(h, "ZG", "shape.sizeDialog"),
    ],
  };
}

/**
 * SmartArt's Format tab (SmartArt-Tools), recorded in PowerPoint (probe_powerpoint_followup.py, D-154): Formen (In 2D
 * bearbeiten, Form ändern, Größer, Kleiner), Formenarten, WordArt-Formate, Barrierefreiheit, Anordnen, Größe. With the
 * whole SmartArt selected PowerPoint greys Formen and the shape styles; Fülleffekt and Formkontur colour the graphic's
 * background and outline.
 */
export function smartArtFormatTab(h: ContextTabHost): Tab {
  return {
    id: "smartartFormat", label: "Format", key: "JO", contextual: true, shrink: [],
    groups: [
      { id: "shapes", label: "Formen", key: "ZM", icon: "shapes", controls: [b("smartart.edit2d", "A1", "large"), stack(dd("smartart.changeShape", "Ä", "small", []), b("smartart.larger", "Ö"), b("smartart.smaller", "L"))] },
      { id: "styles", label: "Formenarten", key: "ZF", icon: "shape-fill", launcher: { cmd: "draw.format", key: "BB" }, controls: [{ type: "dropdown", cmd: "smartart.shapeStyles", key: "NC", size: "large", menu: [] },
        stack(h.color(() => h.fillColor(), (c) => h.setFill(c), "draw.fill", "Ü", true, () => h.fillExtra()), h.color(() => h.lineColor(), (c) => h.setLine(c), "draw.outline", "V", true, () => h.outlineExtra()), dd("draw.effects", "MO", "medium", () => h.effects()))] },
      { id: "wordart", label: "WordArt-Formate", key: "ZW", icon: "wordart", launcher: { cmd: "wordart.dialog", key: "O" }, controls: wordArt(h) },
      { id: "access", label: "Barrierefreiheit", key: "ZB", icon: "alt-text", controls: [b("shape.altTextPane", "AT", "large")] },
      arrangeGroup("ZA", false),
      sizeGroup(h, "ZZ", "shape.sizeDialog"),
    ],
  };
}

export function pictureFormatTab(h: ContextTabHost): Tab {
  return {
    id: "pictureFormat", label: "Bildformat", key: "JP", contextual: true, shrink: [],
    groups: [
      { id: "edit", label: "Bearbeiten", key: "ZE", icon: "edit", controls: [b("pic.edit", "Y1", "large"), b("pic.removeBg", "N", "large")] },
      { id: "adjust", label: "Anpassen", key: "ZA", icon: "corrections", controls: [dd("pic.corrections", "O", "large", () => h.pictureCorrections()), dd("pic.color", "F", "large", () => h.pictureColor()), dd("pic.artistic", "Ü", "large", () => h.pictureArtistic()), dd("pic.transparency", "T", "large", () => h.pictureTransparency()),
        stack(b("pic.compress", "M"), dd("pic.change", "Q1", "small", [{ cmd: "pic.change", label: "Dieses Gerät..." }, { cmd: "pic.changeClipboard" }]), { type: "split", cmd: "pic.reset", key: "U", size: "small", menu: [{ cmd: "pic.reset" }, { cmd: "pic.resetSize" }] })],
        smaller: [[dd("pic.corrections", "O", "large", () => h.pictureCorrections()), stack(dd("pic.color", "F", "medium", () => h.pictureColor()), dd("pic.artistic", "Ü", "medium", () => h.pictureArtistic()), dd("pic.transparency", "T", "medium", () => h.pictureTransparency())),
          stack(b("pic.compress", "M"), dd("pic.change", "Q1", "small", [{ cmd: "pic.change", label: "Dieses Gerät..." }, { cmd: "pic.changeClipboard" }]), { type: "split", cmd: "pic.reset", key: "U", size: "small", menu: [{ cmd: "pic.reset" }, { cmd: "pic.resetSize" }] })]] },
      { id: "styles", label: "Bildformatvorlagen", key: "ZB", icon: "picture-styles", launcher: { cmd: "pic.dialog", key: "P" }, controls: [h.picStyleTiles(8),
        stack(h.color(() => h.lineColor(), (c) => h.setLine(c), "pic.border", "V", true, () => h.outlineExtra()), dd("pic.effects", "E", "medium", () => h.effects()), b("pic.layout", "A1", "medium"))],
        smaller: [[h.picStyleTiles(3), stack(h.color(() => h.lineColor(), (c) => h.setLine(c), "pic.border", "V", true, () => h.outlineExtra()), dd("pic.effects", "E", "medium", () => h.effects()), b("pic.layout", "A1", "medium"))]] },
      { id: "access", label: "Barrierefreiheit", key: "ZR", icon: "alt-text", controls: [b("shape.altTextPane", "AT", "large")] },
      arrangeGroup("ZO", true),
      sizeGroup(h, "ZG", "shape.sizeDialog", [{ type: "split", cmd: "pic.crop", key: "C", size: "large", get menu() { return h.cropMenu(); } } as Control]),
    ],
  };
}

export function tableDesignTab(h: ContextTabHost): Tab {
  return {
    id: "tableDesign", label: "Tabellenentwurf", key: "JT", contextual: true, shrink: [],
    groups: [
      { id: "options", label: "Tabellenformatoptionen", key: "ZO", icon: "table-style", controls: [stack(check("table.firstRow", "Ü"), check("table.lastRow", "Ä"), check("table.bandRow", "U")), stack(check("table.firstCol", "P"), check("table.lastCol", "L"), check("table.bandCol", "V"))] },
      { id: "styles", label: "Tabellenformatvorlagen", key: "ZT", icon: "table-style", controls: [h.tableStyleTiles(), stack(h.color(() => h.fillColor(), (c) => h.setShading(c), "table.shading", "NC", false, () => h.shadingExtra()), dd("table.borders", "DA", "small", () => h.tableBorders()), dd("table.effects", "F", "small", () => h.tableEffects()))] },
      { id: "wordart", label: "WordArt-Formate", key: "ZW", icon: "wordart", launcher: { cmd: "wordart.dialog", key: "O" }, wide: false, controls: wordArtCompact(h) },
      { id: "draw", label: "Rahmenlinien zeichnen", key: "ZR", icon: "table-draw", controls: [stack(dd("table.penStyle", "B", "medium", () => h.penStyles()), dd("table.penWeight", "6", "medium", () => h.penWeights()), h.color(() => h.penColor(), (c) => h.setPenColor(c), "table.penColor", "G", true)), b("table.draw", "AA", "large"), b("table.eraser", "R", "large")] },
    ],
  };
}

export function tableLayoutTab(h: ContextTabHost): Tab {
  return {
    id: "tableLayout", label: "Tabellenlayout", key: "JL", contextual: true, shrink: [],
    groups: [
      { id: "table", label: "Tabelle", key: "ZT", icon: "table", controls: [dd("table.select", "W", "large", [{ cmd: "table.selectAll" }, { cmd: "table.selectCol" }, { cmd: "table.selectRow" }]), b("table.gridlines", "ß", "large")] },
      { id: "rows", label: "Zeilen und Spalten", key: "ZZ", icon: "table-insert-row", sepAfter: [0], controls: [dd("table.deleteMenu", "Ö", "large", [{ cmd: "table.deleteCols" }, { cmd: "table.deleteRows" }, { cmd: "table.delete" }]),
        b("table.rowAbove", "Y1", "large"), b("table.rowBelow", "E", "large"), b("table.colLeft", "Y2", "large"), b("table.colRight", "I1", "large")],
      smaller: [[dd("table.deleteMenu", "Ö", "large", [{ cmd: "table.deleteCols" }, { cmd: "table.deleteRows" }, { cmd: "table.delete" }]),
        b("table.rowAbove", "Y1", "large"), stack(b("table.rowBelow", "E", "medium"), b("table.colLeft", "Y2", "medium"), b("table.colRight", "I1", "medium"))]] },
      { id: "merge", label: "Zusammenführen", key: "ZU", icon: "table-merge", controls: [b("table.merge", "LE", "large"), b("table.split", "NE", "large")] },
      { id: "cellSize", label: "Zellengröße", key: "ZG", icon: "table-autofit", controls: [stack(h.tableSpin("row"), h.tableSpin("col")), stack(b("table.distRows", "V", "medium"), b("table.distCols", "SP", "medium"))],
        smaller: [[stack(h.tableSpin("row", false), h.tableSpin("col", false)), stack(b("table.distRows", "V"), b("table.distCols", "SP"))]] },
      { id: "align", label: "Ausrichtung", key: "ZA", icon: "cell-middle-center", wide: false, controls: [stack(b("table.alignLeft", "II"), b("table.top", "C")), stack(b("table.alignCenter", "RZ"), b("table.middle", "K")), stack(b("table.alignRight", "RE"), b("table.bottom", "A")),
        dd("table.direction", "MT", "large", () => h.cellDirections()), dd("table.margins", "NZ", "large", () => h.cellMargins())] },
      { id: "tableSize", label: "Tabellengröße", key: "ZS", icon: "size-position", controls: [stack(h.tableSpin("h"), h.tableSpin("w"), check("table.lockRatio", "T"))] },
      { ...arrangeGroup("ZO", false), smaller: [[dd("table.arrange", "ZA", "large", [{ cmd: "arrange.forward" }, { cmd: "arrange.backward" }, { cmd: "arrange.selectionPane" }, "-", { label: "Ausrichten", submenu: () => h.alignMenu() },
        { label: "Gruppieren", submenu: [{ cmd: "arrange.group" }, { cmd: "arrange.regroup" }, { cmd: "arrange.ungroup" }] },
        { label: "Drehen", submenu: [{ cmd: "arrange.rotRight" }, { cmd: "arrange.rotLeft" }, { cmd: "arrange.flipV" }, { cmd: "arrange.flipH" }] }])]] },
    ],
  };
}

export function chartDesignTab(h: ContextTabHost): Tab {
  return {
    id: "chartDesign", label: "Diagrammentwurf", key: "JC", contextual: true, shrink: [],
    groups: [
      { id: "layouts", label: "Diagrammlayouts", key: "ZL", icon: "chart-recommended", controls: [dd("chart.addElement", "D", "large", () => h.chartElements()), dd("chart.quickLayout", "C", "large", () => h.quickLayouts())] },
      { id: "styles", label: "Diagrammformatvorlagen", key: "ZS", icon: "styles", controls: [dd("chart.colors", "F", "large", () => h.chartColors()), h.chartStyleTiles()],
        smaller: [[dd("chart.colors", "F", "large", () => h.chartColors()), { type: "dropdown", cmd: "chart.styles", key: "S", size: "large", menu: [], popup: (c) => (h.chartStyleTiles() as Extract<Control, { type: "tiles" }>).panel(c) }]] },
      { id: "data", label: "Daten", key: "ZD", icon: "table", controls: [b("chart.switchRowCol", "E", "large"), b("chart.selectData", "T", "large"), b("chart.editData", "A", "large"), b("chart.refresh", "N", "large")] },
      { id: "type", label: "Typ", key: "ZT", icon: "chart-column", controls: [dd("chart.changeType", "I", "large", () => h.chartKinds())] },
    ],
  };
}

export function chartFormatTab(h: ContextTabHost): Tab {
  return {
    id: "chartFormat", label: "Format", key: "JY", contextual: true, shrink: [],
    groups: [
      // Aktuelle Auswahl as PowerPoint's: the chart's elements (the format pane opens for the chart), Auswahl formatieren, Auf Formatvorl. zurücks.
      { id: "current", label: "Aktuelle Auswahl", key: "ZA", icon: "format-picture", controls: [stack(
        { type: "dropdown", cmd: "chart.elements", key: "A1", size: "small", menu: ["Diagrammbereich", "Zeichnungsfläche", "Diagrammtitel", "Legende", "Horizontal (Kategorie) Achse", "Vertikal (Wert) Achse", "Vertikal (Wert) Achse Hauptgitternetz"].map((label) => ({ label, run: () => h.formatChartElement(label) })) },
        { type: "button", cmd: "draw.format", key: "U", size: "small", label: "Auswahl formatieren" }, b("chart.reset", "V1"))] },
      { id: "insert", label: "Formen einfügen", key: "ZE", icon: "shapes", controls: [h.shapeTiles(3), dd("shape.changeShapeChart", "Ä", "large", [])],
        smaller: [[{ type: "dropdown", cmd: "insert.shapes", key: "2", size: "large", menu: [], popup: (c) => h.shapeGallery(c) }, b("shape.changeShapeChart", "Ä")]] },
      shapeStyles(h, [[{ type: "dropdown", cmd: "shape.quickStyles", key: "NC", size: "large", menu: [], popup: (c) => h.quickStyles(c) }, stack(h.color(() => h.fillColor(), (c) => h.setFill(c), "draw.fill", "Ü", true), h.color(() => h.lineColor(), (c) => h.setLine(c), "draw.outline", "V", true, () => h.outlineExtra()), dd("draw.effects", "MO", "small", () => h.effects()))]]),
      { id: "wordart", label: "WordArt-Formate", key: "ZW", icon: "wordart", launcher: { cmd: "wordart.dialog", key: "O" }, controls: wordArt(h) },
      { id: "access", label: "Barrierefreiheit", key: "ZB", icon: "alt-text", controls: [b("shape.altTextPane", "AT", "large")] },
      arrangeGroup("ZO", false),
      sizeGroup(h, "ZG", "shape.sizePosition"),
    ],
  };
}
