// The workbook model of RaabeCalc. It follows SpreadsheetML closely (D-004: OOXML is the
// document), so reading, writing and keeping unknown parts stay simple:
// - a cell holds its value, its formula in the file's form (English names, "," between
//   arguments, "." as decimal point, no leading "=") and a style index;
// - styles are whole cell formats interned in one table per workbook (Excel's cellXfs);
// - columns and rows hold widths, heights, hidden flags and their own style;
// - whatever the .xlsx layer does not model is kept in `raw` fields and written back.

import { cellKey, keyCell, type Rect } from "./address";
import type { ChartModel } from "@raabe/charts";

/** Error values as stored in the file; the UI shows them in German (`errorText`). */
export type ErrorCode = "#NULL!" | "#DIV/0!" | "#VALUE!" | "#REF!" | "#NAME?" | "#NUM!" | "#N/A" | "#SPILL!" | "#CALC!" | "#GETTING_DATA";

export class CellError {
  constructor(readonly code: ErrorCode) {}
  toString(): string { return this.code; }
}

const ERRORS = new Map<ErrorCode, CellError>();
export function err(code: ErrorCode): CellError {
  let e = ERRORS.get(code);
  if (!e) ERRORS.set(code, (e = new CellError(code)));
  return e;
}

/** Excel's German names of the error values. */
export const ERROR_TEXT_DE: Record<ErrorCode, string> = {
  "#NULL!": "#NULL!", "#DIV/0!": "#DIV/0!", "#VALUE!": "#WERT!", "#REF!": "#BEZUG!", "#NAME?": "#NAME?",
  "#NUM!": "#ZAHL!", "#N/A": "#NV", "#SPILL!": "#ÜBERLAUF!", "#CALC!": "#CALC!", "#GETTING_DATA": "#DATEN_ABRUFEN",
};

export function errorText(e: CellError): string {
  return ERROR_TEXT_DE[e.code];
}

export type Scalar = number | string | boolean | CellError;
/** null is an empty cell. */
export type Value = Scalar | null;

export interface Font {
  name: string;
  size: number;
  bold?: boolean;
  italic?: boolean;
  underline?: "single" | "double" | "singleAccounting" | "doubleAccounting";
  strike?: boolean;
  /** "#RRGGBB"; undefined is automatic (black). */
  color?: string;
  vertAlign?: "superscript" | "subscript";
  /** Theme font scheme reference ("minor", "major") as in the file. */
  scheme?: string;
  /** A design colour (the file's theme index, tint): `color` follows it when the design changes. */
  theme?: ThemeRef;
}

/** A design colour as the .xlsx file refers to it: 0 Hintergrund 1, 1 Text 1, 2 Hintergrund 2, 3 Text 2, 4-9 Akzent 1-6, 10 Link, 11 besuchter Link. */
export interface ThemeRef { t: number; tint?: number }

/** The theme's colour keys by the file's theme index. */
export const THEME_KEYS = ["lt1", "dk1", "lt2", "dk2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6", "hlink", "folHlink"] as const;

export interface BorderSide {
  /** Excel's styles: thin, medium, thick, dashed, dotted, double, hair, mediumDashed, dashDot … */
  style: string;
  color?: string;
}

export interface Borders {
  left?: BorderSide; right?: BorderSide; top?: BorderSide; bottom?: BorderSide;
  diagonal?: BorderSide; diagonalUp?: boolean; diagonalDown?: boolean;
}

export type HAlign = "general" | "left" | "center" | "right" | "fill" | "justify" | "centerContinuous" | "distributed";
export type VAlign = "top" | "center" | "bottom" | "justify" | "distributed";

/** A whole cell format (font, fill, borders, number format, alignment, protection). */
export interface CellStyle {
  font: Font;
  /** Solid fill colour "#RRGGBB"; undefined: no fill. */
  fill?: string;
  /** The fill's design colour, when it is one (`fill` follows it when the design changes). */
  fillTheme?: ThemeRef;
  /** Fills we do not draw (patterns, gradients) as the file's XML, written back unchanged. */
  fillXml?: string;
  borders: Borders;
  /** Excel number format code in the file's (English) form, "General" for Standard. */
  numFmt: string;
  h?: HAlign;
  v?: VAlign;
  wrap?: boolean;
  indent?: number;
  /** 0-180 as in the file (91-180 = downwards), 255 = vertical text. */
  rotation?: number;
  shrink?: boolean;
  /** Protection; Excel's default is locked. */
  locked?: boolean;
  hidden?: boolean;
  /** The cell style (Zellenformatvorlage) it is based on, by name. */
  styleName?: string;
  /**
   * The .xlsx layer's index of the file's cell format this style was read from: an unchanged
   * style is written back as that format (theme colours and all), a changed one reuses what
   * still matches. Copy it along when deriving a style.
   */
  xf?: number;
}

export const DEFAULT_FONT: Font = { name: "Calibri", size: 11, scheme: "minor" };

export function defaultStyle(): CellStyle {
  return { font: { ...DEFAULT_FONT }, borders: {}, numFmt: "General" };
}

/** Interned cell formats: index 0 is the workbook's default. */
export class StyleTable {
  readonly list: CellStyle[] = [];
  private ids = new Map<string, number>();

  constructor(base: CellStyle = defaultStyle()) {
    this.id(base);
  }

  get(id: number | undefined): CellStyle {
    return this.list[id ?? 0] ?? this.list[0];
  }

  /** The index of a format, adding it when it is new. */
  id(style: CellStyle): number {
    const key = styleKey(style);
    let id = this.ids.get(key);
    if (id === undefined) {
      id = this.list.length;
      this.list.push(style);
      this.ids.set(key, id);
    }
    return id;
  }

  /** The style `id` with some fields changed. */
  derive(id: number | undefined, change: (s: CellStyle) => void): number {
    const s = structuredClone(this.get(id));
    change(s);
    return this.id(s);
  }

  /** Changes every style in place (a new design's colours and fonts); ids stay as they are. */
  remap(change: (s: CellStyle) => void): void {
    this.list.forEach((s) => change(s));
    this.ids.clear();
    this.list.forEach((s, i) => { const k = styleKey(s); if (!this.ids.has(k)) this.ids.set(k, i); });
  }
}

function styleKey(s: CellStyle): string {
  return JSON.stringify(s, Object.keys(flatKeys(s)).sort());
}

function flatKeys(o: unknown, out: Record<string, true> = {}): Record<string, true> {
  if (o && typeof o === "object") {
    for (const [k, v] of Object.entries(o)) {
      out[k] = true;
      flatKeys(v, out);
    }
  }
  return out;
}

export interface Cell {
  v: Value;
  /** Formula in the file's form, without "=". */
  f?: string;
  /** Array formula (Strg+Umschalt+Eingabe) covering this range; set on its top-left cell. */
  array?: Rect;
  /** Dynamic array formula: the range its result last spilled into (anchor cell only; the
   *  other cells hold the values, as in the file). Kept up to date by the formula engine. */
  spill?: Rect;
  s?: number;
  /** Rich text runs as the file's XML (shared string <si> content), kept until the text is edited. */
  richXml?: string;
  /** Hyperlink target. */
  link?: string;
  /** Note (Notiz) text. */
  note?: string;
  /** A Datentabelle's result cell: the formula the formula bar shows ("MEHRFACHOPERATION(B1;A1)"; not saved). */
  dt?: string;
  /** Fehler ignorieren was chosen for this cell's error indicator (session only, not saved). */
  ignoreError?: boolean;
  /** The .xlsx layer's parts of the cell it does not model (attributes such as cm/vm, data-table formulas, the formula as the file spelt it). */
  raw?: { attrs?: string; f?: string; file?: string };
}

export interface ColInfo { width?: number; hidden?: boolean; s?: number; outline?: number; bestFit?: boolean; collapsed?: boolean }
export interface RowInfo {
  height?: number; hidden?: boolean; s?: number; outline?: number; custom?: boolean; collapsed?: boolean;
  /** The .xlsx layer's attributes of the row it does not model (thickTop, thickBot, x14ac:dyDescent …): Excel sizes rows by them. */
  raw?: string;
}

export interface View {
  /** Active cell and selection. */
  active: { row: number; col: number };
  selection: Rect[];
  /** Top-left visible cell of the scrollable part. */
  top: number;
  left: number;
  zoom: number;
  showGrid: boolean;
  showHeaders: boolean;
  /** Fenster fixieren: rows above `row` and columns left of `col` stay put (0 = none). */
  freeze?: { row: number; col: number };
  /**
   * Teilen (D-093): the split bars' places as the file's pane xSplit/ySplit, in 1/20 pt from the window's
   * left/top edge with the headers (0: no bar); `top`/`left` are the top-left pane's first row and column
   * (`top`/`left` of the view are the bottom-right pane's); `pane` is the active pane.
   */
  split?: { x: number; y: number; top: number; left: number; pane?: "topLeft" | "topRight" | "bottomLeft" | "bottomRight";
    /** The other panes' selections as the file has them (Excel keeps each pane's own). */
    sels?: Record<string, { ac: string; sqref: string }> };
  /** Fixed by Fenster fixieren while split: Fixierung aufheben gives the split back (recorded). */
  splitFrozen?: boolean;
  /**
   * Seitenlayout view (sheetView view="pageLayout"): `zoom` is then the view's own zoom (zoomScalePageLayoutView)
   * and `normalZoom` the Normal view's (zoomScaleNormal); Lineal off (showRuler="0").
   */
  layout?: boolean;
  normalZoom?: number;
  noRuler?: boolean;
}

/** An ink drawing (Zeichnen, D-093): one xdr:contentPart with its InkML part, as Excel stores ink. */
export interface SheetInk {
  from: { col: number; row: number; x: number; y: number };
  to: { col: number; row: number; x: number; y: number };
  /** Its box in px at 100 % (the strokes are drawn into it). */
  size: { w: number; h: number };
  /** The strokes: points in px at 100 % inside the box, t in ms from the drawing's start (Freihandwiedergabe). */
  strokes: { pts: { x: number; y: number; t?: number }[]; brush: InkBrush }[];
  name?: string;
  /** The InkML's timestamp (timeString). */
  time?: string;
  /** The Fallback picture as a data: URL (PNG), for programs without ink. */
  fallback?: string;
  /** The anchor's XML and its InkML part as read: kept while unchanged. */
  raw?: string;
  part?: string;
  rawKey?: string;
}

export interface InkBrush {
  color: string;
  /** Pen width in mm. */
  width: number;
  /** Textmarker (tip rectangle, rasterOp maskPen) and Bleistift (pencil texture). */
  kind?: "pen" | "highlighter" | "pencil";
  /** Further brush properties as the file has them (transparency, ignorePressure …). */
  extra?: Record<string, string>;
}

/** One column's filter: chosen values, or Zahlenfilter/Textfilter (custom), Top 10, above/below the average. */
export interface FilterColumn {
  values?: string[];
  blank?: boolean;
  /** Benutzerdefinierter AutoFilter: one or two conditions (operator as in the file, value with * and ? wildcards). */
  custom?: { and?: boolean; items: { op: string; val: string }[] };
  /** Top 10: the top or bottom items or percent; filterVal is the limit found when it was applied. */
  top10?: { top: boolean; percent: boolean; val: number; filterVal?: number };
  /** Über/Unter dem Durchschnitt; val is the average when it was applied. */
  dynamic?: { type: "aboveAverage" | "belowAverage"; val?: number };
  /** Nach Farbe filtern: a fill colour (null: Keine Füllung) or a font colour (null: Automatisch). */
  color?: { fill?: string | null; font?: string | null };
}

export interface AutoFilter {
  ref: Rect;
  /** Per column (0 = first column of `ref`): the texts shown (as the cells show them); absent = all. */
  columns: Record<number, FilterColumn>;
  /** The file's <autoFilter> children the app does not model (custom, top 10 …), kept as XML. */
  rawXml?: string;
}

/** One rule of a conditional format, as SpreadsheetML's cfRule. */
export interface CondRule {
  type: "cellIs" | "containsText" | "notContainsText" | "beginsWith" | "endsWith" | "duplicateValues" | "uniqueValues" | "top10" | "aboveAverage" | "dataBar" | "colorScale" | "expression" | "containsBlanks" | "timePeriod" | "iconSet" | "other";
  priority: number;
  /** cellIs: between, notBetween, equal, notEqual, greaterThan, lessThan, greaterThanOrEqual, lessThanOrEqual. */
  operator?: string;
  /** Formulas in the file's form (cellIs values, expression). */
  formulas?: string[];
  text?: string;
  rank?: number;
  percent?: boolean;
  bottom?: boolean;
  /** aboveAverage: false = below. */
  above?: boolean;
  stopIfTrue?: boolean;
  /** The format applied (from the file's dxf): font colour, bold/italic, fill. */
  format?: { color?: string; fill?: string; bold?: boolean; italic?: boolean; underline?: boolean; strike?: boolean; border?: string };
  /** dataBar: one colour; colorScale: two or three colours (min, [mid], max). */
  colors?: string[];
  /**
   * Symbolsatz: the set's file name (3TrafficLights1, 3Arrows …), reversed order, value hidden,
   * and the limits above the first icon (percent, num, percentile or formula; gte false = ">").
   */
  iconSet?: { set: string; reverse?: boolean; showValue?: boolean; cfvo: { type: string; val: string; gte?: boolean }[] };
  /** The whole cfRule XML when the app cannot write it back from these fields (icon sets …). */
  rawXml?: string;
}

/** One comment of a thread (Excel 365's Kommentar): author, ISO date, text. */
export interface ThreadComment { id: string; author: string; personId?: string; date: string; text: string }

/** A Kommentar thread on a cell: the first comment, its replies, and whether it is resolved. */
export interface CommentThread extends ThreadComment { row: number; col: number; replies: ThreadComment[]; done?: boolean }

/** A sparkline colour: "#RRGGBB", and the design colour it was picked as (it follows the design). */
export interface SparkColor { rgb: string; theme?: ThemeRef }

/** The eight colours of a sparkline group, as x14:sparklineGroup has them (colorSeries … colorLow). */
export type SparkColorKey = "series" | "negative" | "axis" | "markers" | "first" | "last" | "high" | "low";

/**
 * Sparklines (D-085): a group of small charts in cells, each drawn from a range ("Tabelle1!A2:E2"); every
 * setting belongs to the group, as in Excel (x14:sparklineGroup). Absent flags are off; absent values are
 * Excel's defaults (lineWeight 0.75, axis types "individual").
 */
export interface SparklineGroup {
  type: "line" | "column" | "stacked";
  colors: Record<SparkColorKey, SparkColor>;
  /** Anzeigen: Datenpunkte (markers, lines only), Höchstpunkt, Tiefpunkt, Erster Punkt, Letzter Punkt, Negative Punkte. */
  markers?: boolean; high?: boolean; low?: boolean; first?: boolean; last?: boolean; negative?: boolean;
  /** Achse anzeigen (the horizontal axis at zero). */
  displayXAxis?: boolean;
  /** Leere Zellen anzeigen als: Lücken, Null, Datenpunkte mit einer Linie verbinden (the file's default is "zero"). */
  displayEmptyCellsAs: "gap" | "zero" | "span";
  /** Daten in ausgeblendeten Zeilen und Spalten anzeigen. */
  displayHidden?: boolean;
  /** Stärke of the line in pt. */
  lineWeight?: number;
  /** Vertical axis: Automatisch für jede Sparkline (individual), Identisch für alle Sparklines (group), Benutzerdefinierter Wert. */
  minAxisType?: "individual" | "group" | "custom";
  maxAxisType?: "individual" | "group" | "custom";
  manualMin?: number;
  manualMax?: number;
  /** Daten von rechts nach links darstellen. */
  rightToLeft?: boolean;
  /** Datumsachsentyp: the range holding the dates ("Tabelle1!A7:E7"), x14:sparklineGroup dateAxis with its xm:f. */
  dateRange?: string;
  items: { f: string; row: number; col: number }[];
}

/** A shape of the sheet's drawing (Einfügen > Formen, Textfeld; D-066). Unset fill and line: the
 *  theme's shape style (Akzent 1 with a darker outline; text boxes white with a grey outline). */
export interface SheetShape {
  from: { col: number; row: number; x: number; y: number };
  to: { col: number; row: number; x: number; y: number };
  /** DrawingML preset geometry ("rect", "ellipse", "rightArrow", "line" …). */
  geom: string;
  /**
   * The shape style's line reference (xdr:style lnRef): which of the theme's line styles and how
   * much Akzent 1 is shaded. Excel 365 gives shapes style 2 shaded to 15 %, lines style 1 (Office
   * 2013-2022 theme) or 2 (the current Office theme).
   */
  style?: { ln: number; shade?: number };
  /**
   * Gruppieren: a group (geom "group") holds its shapes; each child's box is `rel`, in px at 100 %
   * in the group's own coordinates, whose extent is `frame` (the group's size when it was made).
   */
  children?: SheetShape[];
  rel?: { x: number; y: number; w: number; h: number };
  frame?: { w: number; h: number };
  /** Ausgeblendet in the Auswahlbereich (cNvPr hidden="1"). */
  hidden?: boolean;
  /** Freihandform, Kurve, Skizze: custom geometry (a:custGeom) in text-engine's definition syntax. */
  custom?: string;
  /** The preset's adjust values that differ from its defaults (a:avLst, the yellow handles). */
  adjust?: Record<string, number>;
  /** Lines and arrows (xdr:cxnSp). */
  connector?: boolean;
  /** Its size in px at 100 % (the file's extent; the anchors decide where it is). */
  size?: { w: number; h: number };
  fill?: { type: "none" | "solid"; color?: string; transparency?: number };
  line?: { type: "none" | "solid"; color?: string; width?: number; dash?: string; head?: string; tail?: string };
  /** The text, lines separated by "
". */
  text?: string;
  font?: { size?: number; bold?: boolean; italic?: boolean; underline?: boolean; color?: string };
  align?: "l" | "ctr" | "r";
  anchor?: "t" | "ctr" | "b";
  /** Drehung in degrees, Spiegeln. */
  rot?: number;
  flipH?: boolean;
  flipV?: boolean;
  /** Textfeld (white with a grey outline, text at the top left). */
  textBox?: boolean;
  name?: string;
  /** The anchor's XML as read and what it held then: written back as it is while unchanged. */
  raw?: string;
  rawKey?: string;
}

/** A picture in the sheet's drawing, shown as the file places it (the file keeps it as it was). */
export interface SheetPicture {
  /** Top-left: the cell and the offset inside it in px at 100 %. */
  from: { col: number; row: number; x: number; y: number };
  /** Bottom-right as a cell anchor (two-cell anchors) … */
  to?: { col: number; row: number; x: number; y: number };
  /** … or the size in px at 100 % (one-cell anchors). */
  size?: { w: number; h: number };
  /** The picture as a data: URL. */
  src: string;
  name?: string;
  /** The anchor's XML as the file had it (kept as it is while the picture is not moved). */
  raw?: string;
  /**
   * Bildformat (D-087): text-engine's PictureFormat (crop, colour settings, rotation, line,
   * effects, frame shape), as read from the picture's DrawingML.
   */
  format?: object;
  /** Alternativtext (cNvPr descr, title); Ausgeblendet in the Auswahlbereich (hidden="1"). */
  alt?: string;
  title?: string;
  hidden?: boolean;
  /** Seitenverhältnis sperren (picLocks noChangeAspect="1"). */
  lockAspect?: boolean;
  /**
   * The Office 2010 effects (Korrekturen, Sättigung, Farbton) already worked into an image, as
   * a key (JSON of those settings), and that image when it is not `src` (Excel stores pictures
   * with the effects applied and draws them as stored).
   */
  baked?: string;
  bakedSrc?: string;
  /** The file's relationships for the picture (blip, a14 image layer) and the image they hold. */
  media?: { rid: string; layer?: string; src: string };
  /** What the settings above were when read: the anchor is rewritten only when they change. */
  lookKey?: string;
  /** A picture inside a group (grpSp): its box as shares of the group's anchored box; the group is kept as it is. */
  group?: { x: number; y: number; w: number; h: number };
}

export interface CondFormat {
  ref: Rect[];
  rules: CondRule[];
}

/** Zulassen in Datenüberprüfung; "any" is Jeden Wert (no type in the file). */
export type ValidationType = "any" | "whole" | "decimal" | "list" | "date" | "time" | "textLength" | "custom";
export type ValidationOperator = "between" | "notBetween" | "equal" | "notEqual" | "greaterThan" | "lessThan" | "greaterThanOrEqual" | "lessThanOrEqual";

/**
 * Datenüberprüfung of some ranges, as SpreadsheetML's <dataValidation> (and the x14 form Excel
 * writes into the extLst when a formula refers to another sheet). Flags are false when absent.
 */
export interface Validation {
  ref: Rect[];
  type: ValidationType;
  /** Daten; absent: between (the file's default). */
  operator?: ValidationOperator;
  /**
   * Criteria in the file's form without "=": numbers with ".", dates and times as serial
   * numbers, a literal list as a quoted string ("\"Ja,Nein\""), a source range or name, a
   * formula with English names. Relative references belong to the top-left cell of `ref[0]`.
   */
  formula1?: string;
  formula2?: string;
  /** A literal list's entries when one holds a comma or quote (Excel's x12ac:list); formula1 is then Excel's fallback. */
  items?: string[];
  /** Leere Zellen ignorieren. */
  allowBlank?: boolean;
  /** Excel's inverted flag: true HIDES the cell's dropdown arrow (Zellendropdown off). */
  showDropDown?: boolean;
  showInputMessage?: boolean;
  promptTitle?: string;
  prompt?: string;
  showErrorMessage?: boolean;
  /** Typ of the Fehlermeldung; absent: stop (Stopp). */
  errorStyle?: "stop" | "warning" | "information";
  errorTitle?: string;
  error?: string;
  /** The element as the file had it when the fields above cannot express it (imeMode …); written back with the new sqref. */
  rawXml?: string;
}

export interface PageSetup {
  orientation?: "portrait" | "landscape";
  /** Excel's paperSize code (9 = A4, 1 = Letter, 11 = A5, 8 = A3). */
  paperSize?: number;
  /** Margins in inches. */
  margins?: { left: number; right: number; top: number; bottom: number; header: number; footer: number };
  scale?: number;
  fitWidth?: number;
  fitHeight?: number;
  gridLines?: boolean;
  headings?: boolean;
  /** The print area (the defined name _xlnm.Print_Area). */
  printArea?: Rect;
  /** Manual page breaks: a new page starts at these rows / columns (0-based; the file's brk id). */
  rowBreaks?: number[];
  colBreaks?: number[];
  /** Kopf- und Fußzeile in Excel's codes ("&LText&C&P&R&D"; &P page, &N pages, &D date, &T time, &F file, &A sheet). */
  header?: string;
  footer?: string;
  /** Drucktitel: rows repeated at the top and columns at the left of every page (_xlnm.Print_Titles). */
  titleRows?: [number, number];
  titleCols?: [number, number];
  /** Auf der Seite zentrieren: horizontally / vertically between the margins (printOptions). */
  centerH?: boolean;
  centerV?: boolean;
}

/**
 * Szenario-Manager: one scenario of a sheet (SpreadsheetML's CT_Scenario). The values are kept as
 * the file holds them, the text typed into Szenariowerte (German Excel writes "0,25"); Anzeigen
 * enters them as if typed.
 */
export interface Scenario {
  name: string;
  /** Veränderbare Zellen in the file's order (they need not be next to each other), with their values. */
  cells: { row: number; col: number; val: string; numFmtId?: number; deleted?: boolean; undone?: boolean }[];
  comment?: string;
  /** The author as the file has it (Excel's user name). */
  user?: string;
  /** Schutz: Änderungen verhindern (Excel's default) and Ausblenden; they count while the sheet is protected. */
  locked?: boolean;
  hidden?: boolean;
}

/** A sheet's scenarios (CT_Scenarios): the list, the one selected and the one shown last, the last Ergebniszellen. */
export interface Scenarios {
  list: Scenario[];
  current?: number;
  show?: number;
  sqref?: Rect[];
}

/** Konsolidieren's function as the file names it (CT_DataConsolidate function; absent: sum). */
export type ConsolidateFunction = "average" | "count" | "countNums" | "max" | "min" | "product" | "stdDev" | "stdDevp" | "sum" | "var" | "varp";

/** Konsolidieren's settings, kept by Excel in the destination sheet (CT_DataConsolidate) for the dialog's next use. */
export interface Consolidation {
  fn?: ConsolidateFunction;
  /** Beschriftung aus: Oberster Zeile, Linker Spalte; Verknüpfungen mit Quelldaten. */
  top?: boolean;
  left?: boolean;
  link?: boolean;
  startLabels?: boolean;
  /** Vorhandene Verweise: a range on a sheet (A1 form without $), or a defined name. */
  refs: { ref?: string; sheet?: string; name?: string; rid?: string }[];
}

export class Sheet {
  cells = new Map<number, Cell>();
  cols = new Map<number, ColInfo>();
  rows = new Map<number, RowInfo>();
  merges: Rect[] = [];
  /** The sheet is filtered by a Spezialfilter in place (sheetPr filterMode; an AutoFilter says so itself). */
  filterMode?: boolean;
  /** Width of columns without their own, in characters of the default font (Excel: 8.43). */
  defaultColWidth?: number;
  /** Height of rows without their own, in points (Calibri 11: 15). */
  defaultRowHeight = 15;
  view: View = { active: { row: 0, col: 0 }, selection: [{ r0: 0, c0: 0, r1: 0, c1: 0 }], top: 0, left: 0, zoom: 100, showGrid: true, showHeaders: true };
  tabColor?: string;
  /** Charts on the sheet (the drawing's chart frames), in drawing order. */
  charts?: ChartModel[];
  hidden?: boolean | "very";
  /** AutoFilter: its range and the values shown per column (column index inside the range). */
  autoFilter?: AutoFilter;
  /** Sparklines of the sheet (Einfügen > Sparklines). */
  sparklines?: SparklineGroup[];
  /** Kommentare (threaded comments) of the sheet's cells. */
  threads?: CommentThread[];
  /** Pictures of the drawing (read only: shown, kept in the file as they are). */
  pictures?: SheetPicture[];
  /** Excel tables (Tabellen) on this sheet. */
  tables?: import("./table").SheetTable[];
  /** PivotTables whose report is on this sheet. */
  pivots?: import("./pivot").PivotTable[];
  /** Datenschnitte on the sheet (D-092). */
  slicers?: import("./slicer").Slicer[];
  /** Zeitachsen on the sheet (D-121). */
  timelines?: import("./timeline").Timeline[];
  /** Shapes and text boxes of the drawing. */
  shapes?: SheetShape[];
  /** Ink drawn with the Zeichnen tab (D-093). */
  inks?: SheetInk[];
  /** Neues Fenster (D-093): this sheet's view in the workbook's further windows (views[i - 1] for window i). */
  views?: View[];
  /** Benutzerdefinierte Ansichten: the sheet's customSheetView elements by guid, as the file has them. */
  customViews?: { guid: string; xml: string }[];
  /** Blatt schützen: the attributes of <sheetProtection> as the file has them (sheet="1" when on). */
  protection?: Record<string, string>;
  /** Bedingte Formatierung, in the file's order (priority decides). */
  conditional?: CondFormat[];
  /** Datenüberprüfung, in the file's order (the first one covering a cell applies). */
  validations?: Validation[];
  /** Szenario-Manager: the sheet's scenarios. */
  scenarios?: Scenarios;
  /** Konsolidieren: the settings last used with this sheet as the destination. */
  consolidate?: Consolidation;
  /** Seite einrichten (pageSetup, pageMargins, printOptions, the print area). */
  page?: PageSetup;
  /** The .xlsx layer's parts of the sheet it does not model, written back unchanged. */
  raw?: unknown;

  constructor(public name: string) {}

  get(row: number, col: number): Cell | undefined {
    return this.cells.get(cellKey(row, col));
  }

  value(row: number, col: number): Value {
    return this.cells.get(cellKey(row, col))?.v ?? null;
  }

  set(row: number, col: number, cell: Cell | undefined): void {
    const k = cellKey(row, col);
    if (!cell || (cell.v === null && cell.f === undefined && cell.s === undefined && !cell.note && !cell.link)) this.cells.delete(k);
    else this.cells.set(k, cell);
  }

  /** The smallest rectangle holding every cell with content or format (Excel's used range). */
  usedRange(): Rect | null {
    let r0 = Infinity, c0 = Infinity, r1 = -1, c1 = -1;
    for (const k of this.cells.keys()) {
      const { row, col } = keyCell(k);
      if (row < r0) r0 = row;
      if (row > r1) r1 = row;
      if (col < c0) c0 = col;
      if (col > c1) c1 = col;
    }
    return r1 < 0 ? null : { r0, c0, r1, c1 };
  }

  /** The merged area starting at or covering a cell. */
  mergeAt(row: number, col: number): Rect | undefined {
    return this.merges.find((m) => row >= m.r0 && row <= m.r1 && col >= m.c0 && col <= m.c1);
  }
}

export interface DefinedName {
  name: string;
  /** Formula in the file's form ("Tabelle1!$A$1:$B$4"). */
  ref: string;
  /** Sheet index for a name local to one sheet. */
  sheet?: number;
  hidden?: boolean;
}

export class Workbook {
  sheets: Sheet[] = [];
  active = 0;
  styles = new StyleTable();
  names: DefinedName[] = [];
  /** Dates count from 1904 (Mac workbooks) instead of 1900. */
  date1904 = false;
  /** Berechnungsoptionen: Automatisch (absent), Automatisch außer bei Datentabellen, Manuell. */
  calcMode?: "autoNoTable" | "manual";
  /** Arbeitsmappe schützen: <workbookProtection>'s attributes as the file has them (lockStructure="1" …). */
  protection?: Record<string, string>;
  /** Theme colours (dk1, lt1, dk2, lt2, accent1-6, hlink, folHlink) as "#RRGGBB". */
  theme: Record<string, string> = {};
  /** The widths of the theme's line styles in pt (shapes' outlines refer to them). */
  themeLines: number[] = [0.5, 1, 1.5];
  /** The theme's fonts (Überschriften, Textkörper): cells with font.scheme use them. */
  themeFonts: { major: string; minor: string } = { major: "Calibri Light", minor: "Calibri" };
  /** The names of the design and its colour and font sets as the theme part has them. */
  themeNames: { theme: string; colors: string; fonts: string } = { theme: "Office", colors: "Office", fonts: "Office" };
  /** The number of the last sheet Excel would have named "TabelleN" (new sheets count on from it). */
  sheetSeq?: number;
  /** The VBA project (xl/vbaProject.bin) of a workbook with macros (.xlsm); written back while it is set. */
  vba?: Uint8Array;
  /**
   * Neues Fenster (D-093): every window of the workbook (bookViews' workbookView) when there are several or one
   * is hidden; window 0's active sheet is `active`, the others' `active` here. `attrs` are the file's attributes.
   */
  windows?: { active: number; hidden?: boolean; attrs?: Record<string, string> }[];
  /**
   * The window whose views the sheets' `view` and `active` hold while it is shown (the app swaps them, see
   * swapWindow); the file is written with window 0's in place.
   */
  shownWindow?: number;
  /** Benutzerdefinierte Ansichten (customWorkbookView): name, guid, Druckeinstellungen, Ausgeblendete Zeilen/Spalten. */
  customViews?: { name: string; guid: string; print: boolean; rowCol: boolean; attrs: Record<string, string> }[];
  /** The document's name without extension as the title bar shows it (ZELLE("Adresse") on another sheet). */
  docName?: string;
  /** The file's full path once saved or opened (ZELLE("Dateiname"), INFO("Verzeichnis")); never written to the file. */
  docPath?: string;
  /** The workbook's slicer caches (Datenschnitte, D-092). */
  slicerCaches?: import("./slicer").SlicerCache[];
  /** The workbook's timeline caches (Zeitachsen, D-121). */
  timelineCaches?: import("./timeline").TimelineCache[];
  /** The .xlsx layer's package and parts it does not model. */
  raw?: unknown;

  sheet(i = this.active): Sheet {
    return this.sheets[i];
  }

  sheetByName(name: string): Sheet | undefined {
    const n = name.toLowerCase();
    return this.sheets.find((s) => s.name.toLowerCase() === n);
  }

  static blank(sheetName = "Tabelle1"): Workbook {
    const wb = new Workbook();
    wb.sheets.push(new Sheet(sheetName));
    return wb;
  }
}

/**
 * Neues Fenster (D-093): swaps window k's views and active sheet with window 0's in place (`view`, `active`);
 * doing it twice restores them.
 */
export function swapWindow(wb: Workbook, k: number): void {
  const w = wb.windows?.[k];
  if (!k || !w) return;
  for (const s of wb.sheets) {
    s.views ??= [];
    const other = s.views[k - 1] ?? JSON.parse(JSON.stringify(s.view));
    s.views[k - 1] = s.view;
    s.view = other;
  }
  const a = w.active;
  w.active = wb.active;
  wb.active = Math.min(a, wb.sheets.length - 1);
}
