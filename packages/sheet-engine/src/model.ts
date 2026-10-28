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
  "#NUM!": "#ZAHL!", "#N/A": "#NV", "#SPILL!": "#ÜBERLAUF!", "#CALC!": "#KALK!", "#GETTING_DATA": "#DATEN_ABRUFEN",
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
}

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
}

export class Sheet {
  cells = new Map<number, Cell>();
  cols = new Map<number, ColInfo>();
  rows = new Map<number, RowInfo>();
  merges: Rect[] = [];
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
  /** Pictures of the drawing (read only: shown, kept in the file as they are). */
  pictures?: SheetPicture[];
  /** Blatt schützen: the attributes of <sheetProtection> as the file has them (sheet="1" when on). */
  protection?: Record<string, string>;
  /** Bedingte Formatierung, in the file's order (priority decides). */
  conditional?: CondFormat[];
  /** Datenüberprüfung, in the file's order (the first one covering a cell applies). */
  validations?: Validation[];
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
  /** Arbeitsmappe schützen: <workbookProtection>'s attributes as the file has them (lockStructure="1" …). */
  protection?: Record<string, string>;
  /** Theme colours (dk1, lt1, dk2, lt2, accent1-6, hlink, folHlink) as "#RRGGBB". */
  theme: Record<string, string> = {};
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
