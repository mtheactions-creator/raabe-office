// Excel tables (Einfügen > Tabelle, Start > Als Tabelle formatieren; ListObjects; D-066): a named
// range with a header row, an optional Ergebniszeile, its columns' names and the table style with
// its options. The cells hold the data as usual; the table only describes them.

import type { Rect } from "./address";

export interface TableColumn {
  /** The header text (unique within the table). */
  name: string;
  /** Ergebniszeile: Excel's function name ("sum", "average", "count", "countNums", "max", "min", "stdDev", "var", "custom") … */
  totalsFunction?: string;
  /** … or a text in the total row (the first column's "Ergebnis"). */
  totalsLabel?: string;
  /** Berechnete Spalte: the column's formula in the file's form (structured references). */
  formula?: string;
}

export interface SheetTable {
  /** Tabellenname (also its display name): unique in the workbook. */
  name: string;
  /** The whole table: header row, data and total row. */
  ref: Rect;
  /** Kopfzeile. */
  header: boolean;
  /** Ergebniszeile. */
  totals: boolean;
  columns: TableColumn[];
  /** Built-in style name (TableStyleMedium2 …) or "" for none. */
  style: string;
  firstColumn: boolean;
  lastColumn: boolean;
  rowStripes: boolean;
  columnStripes: boolean;
  /** Schaltfläche "Filter". */
  filterButton: boolean;
  /** The AutoFilter's chosen values per column (0-based; "" for empty cells), as Datenschnitte set them (D-092). */
  filters?: Record<number, string[]>;
  /** Attributes of a query's table (tableType, connectionId) and of its columns, kept for the file. */
  fileAttrs?: string;
  columnAttrs?: string[];
  /** The file's table id (unique in the workbook). */
  id?: number;
  /** The .xlsx layer's part and what the model held when read (written back verbatim while unchanged). */
  raw?: string;
  rawKey?: string;
}

/** What a table holds (everything but its file part), to see whether it changed. */
export function tableKey(t: SheetTable): string {
  const { raw: _r, rawKey: _k, ...rest } = t;
  void _r; void _k;
  return JSON.stringify(rest);
}

/** The data rows of a table (without header and total row). */
export function tableData(t: SheetTable): Rect {
  return { r0: t.ref.r0 + (t.header ? 1 : 0), c0: t.ref.c0, r1: t.ref.r1 - (t.totals ? 1 : 0), c1: t.ref.c1 };
}

/** Excel's new table name: "Tabelle" with the next free number. */
export function nextTableName(taken: Iterable<string>): string {
  const used = new Set([...taken].map((n) => n.toLowerCase()));
  for (let i = 1; ; i++) if (!used.has(`tabelle${i}`)) return `Tabelle${i}`;
}

/** Whether a name may be a table name (Excel's rules: a letter, _ or \ first, no spaces, not a cell address). */
export function validTableName(name: string): boolean {
  if (!name || name.length > 255) return false;
  if (!/^[\p{L}_\\][\p{L}\p{N}_.\\]*$/u.test(name)) return false;
  if (/^[A-Za-z]{1,3}\d+$/.test(name) || /^[RrCcZzSs]$/.test(name)) return false;
  return true;
}
