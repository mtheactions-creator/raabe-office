// PivotTables (Einfügen > PivotTable, D-066): the model and the table they produce. The source is a
// range whose first row names the fields; rows and columns group by the fields put there, values
// are summarised (Summe, Anzahl, Mittelwert …), filters keep chosen items. The output is laid out as
// Excel lays out its Kompaktformat (row labels in one column), Gliederungsformat or Tabellenformat,
// with Excel's German labels ("Zeilenbeschriftungen", "Gesamtergebnis", "Summe von Wert").

export type PivotFunc = "sum" | "count" | "average" | "max" | "min" | "product" | "countNums" | "stdDev" | "stdDevp" | "var" | "varp";

/** Werte zusammenfassen nach: the labels Excel shows (and puts in front of the field's name). */
export const PIVOT_FUNCS: { f: PivotFunc; label: string }[] = [
  { f: "sum", label: "Summe" }, { f: "count", label: "Anzahl" }, { f: "average", label: "Mittelwert" }, { f: "max", label: "Max." },
  { f: "min", label: "Min." }, { f: "product", label: "Produkt" }, { f: "countNums", label: "Anzahl Zahlen" }, { f: "stdDev", label: "StdAbw" },
  { f: "stdDevp", label: "StdAbwN" }, { f: "var", label: "Varianz" }, { f: "varp", label: "Varianzen" },
];

export interface PivotValueField { field: number; func: PivotFunc; name?: string; numFmt?: string }

export interface PivotTable {
  name: string;
  /** The source range on a sheet (first row: field names). */
  source: { sheet: string; r0: number; c0: number; r1: number; c1: number };
  /** Where the report starts (top-left cell of its filters, or of the table without filters). */
  at: { row: number; col: number };
  rows: number[];
  cols: number[];
  values: PivotValueField[];
  /** Report filters: field, and the items shown (undefined: all). */
  filters: { field: number; items?: string[] }[];
  /** Items hidden in row and column fields (Element filter of the field). */
  hidden?: Record<number, string[]>;
  layout?: "compact" | "outline" | "tabular";
  /** Gesamtergebnisse für Zeilen / Spalten (both on by default). */
  rowGrand?: boolean;
  colGrand?: boolean;
  /** Teilergebnisse oben in der Gruppe anzeigen (compact/outline); false: unten; "none": keine. */
  subtotals?: "top" | "bottom" | "none";
  /** Excel's style name (PivotStyleLight16 is the default). */
  style?: string;
  /** The cells the report took last time it was written (to clear it, and to know a cell is part of it). */
  area?: { r0: number; c0: number; r1: number; c1: number };
  /** The file's table part it was read from, and what it held then (kept while unchanged). */
  raw?: string;
  rawKey?: string;
}

export type PivotValue = string | number | boolean | null;

/** One cell of the produced report: its value and what kind of cell it is (for the look). */
export interface PivotCell {
  row: number;
  col: number;
  v: PivotValue;
  kind: "header" | "rowLabel" | "colLabel" | "value" | "subtotalLabel" | "subtotal" | "grandLabel" | "grand" | "filter" | "corner";
  /** Outline level of a row label (indent in the compact form). */
  level?: number;
  numFmt?: string;
}

export interface PivotOutput { cells: PivotCell[]; rows: number; cols: number }

/** The field's caption in the value area: "Summe von Wert" unless named. */
export function valueCaption(names: string[], v: PivotValueField, all?: PivotValueField[]): string {
  if (v.name) return v.name;
  // A field summarised a second time gets a number: "Anzahl von Note2" (recorded).
  const k = all ? all.slice(0, all.indexOf(v)).filter((x) => x.field === v.field).length : 0;
  return `${PIVOT_FUNCS.find((x) => x.f === v.func)!.label} von ${names[v.field]}${k ? k + 1 : ""}`;
}

function summarise(xs: PivotValue[], f: PivotFunc): PivotValue {
  const nums = xs.filter((x): x is number => typeof x === "number");
  switch (f) {
    case "count": return xs.filter((x) => x !== null && x !== "").length;
    case "countNums": return nums.length;
    case "sum": return nums.reduce((a, b) => a + b, 0);
    case "product": return nums.length ? nums.reduce((a, b) => a * b, 1) : 0;
    case "average": return nums.length ? nums.reduce((a, b) => a + b, 0) / nums.length : null;
    case "max": return nums.length ? Math.max(...nums) : 0;
    case "min": return nums.length ? Math.min(...nums) : 0;
    default: {
      const n = nums.length;
      const pop = f === "stdDevp" || f === "varp";
      if (n < (pop ? 1 : 2)) return null;
      const m = nums.reduce((a, b) => a + b, 0) / n;
      const v = nums.reduce((a, b) => a + (b - m) ** 2, 0) / (pop ? n : n - 1);
      return f === "var" || f === "varp" ? v : Math.sqrt(v);
    }
  }
}

const itemText = (v: PivotValue) => (v === null || v === "" ? "(Leer)" : typeof v === "boolean" ? (v ? "WAHR" : "FALSCH") : String(v));

/** Items in Excel's order: numbers ascending, then text alphabetically, "(Leer)" last. */
function sortItems(items: PivotValue[]): PivotValue[] {
  const rank = (v: PivotValue) => (v === null || v === "" ? 3 : typeof v === "number" ? 0 : typeof v === "boolean" ? 2 : 1);
  return [...items].sort((a, b) => rank(a) - rank(b) || (typeof a === "number" && typeof b === "number" ? a - b : String(a).localeCompare(String(b), "de", { sensitivity: "base" })));
}

/** The report of a PivotTable over source records (first record: field names). */
export function computePivot(p: PivotTable, data: PivotValue[][]): PivotOutput {
  const names = (data[0] ?? []).map((v) => itemText(v));
  let records = data.slice(1).filter((r) => r.some((v) => v !== null && v !== ""));
  for (const f of p.filters) if (f.items) records = records.filter((r) => f.items!.includes(itemText(r[f.field] ?? null)));
  for (const [k, hid] of Object.entries(p.hidden ?? {})) records = records.filter((r) => !hid.includes(itemText(r[Number(k)] ?? null)));
  const cells: PivotCell[] = [];
  const at = p.at;
  // Report filters above the table with an empty row between (the table stays where it is).
  const top = at.row;
  p.filters.forEach((f, i) => {
    const r = at.row - p.filters.length - 1 + i;
    cells.push({ row: r, col: at.col, v: names[f.field], kind: "filter" });
    cells.push({ row: r, col: at.col + 1, v: !f.items ? "(Alle)" : f.items.length === 1 ? f.items[0] : "(Mehrere Elemente)", kind: "filter" });
  });
  const vals = p.values;
  const multiVal = vals.length > 1;
  // Column keys: combinations of column field items; with several value fields "Werte" is a column field too.
  const keyOf = (r: PivotValue[], fields: number[]) => fields.map((f) => itemText(r[f] ?? null));
  const tree = (fields: number[], recs: PivotValue[][]): { key: string[]; recs: PivotValue[][] }[] => {
    const out: { key: string[]; recs: PivotValue[][] }[] = [];
    const walk = (level: number, prefix: string[], rs: PivotValue[][]) => {
      if (level === fields.length) { out.push({ key: prefix, recs: rs }); return; }
      const items = sortItems([...new Map(rs.map((r) => [itemText(r[fields[level]] ?? null), r[fields[level]] ?? null])).values()]);
      for (const it of items) walk(level + 1, [...prefix, itemText(it)], rs.filter((r) => itemText(r[fields[level]] ?? null) === itemText(it)));
    };
    walk(0, [], recs);
    return out;
  };
  const colGroups = tree(p.cols, records);
  const colKeys = p.cols.length ? colGroups : [{ key: [] as string[], recs: records }];
  const valCols: { key: string[]; v: number }[] = [];
  for (const g of colKeys) for (let v = 0; v < Math.max(1, vals.length); v++) valCols.push({ key: g.key, v });
  const rowGrand = p.rowGrand !== false && p.cols.length > 0, colGrand = p.colGrand !== false && p.rows.length > 0;
  const compact = (p.layout ?? "compact") === "compact";
  const labelCols = compact ? 1 : Math.max(1, p.rows.length);
  const headRows = p.cols.length ? 2 : 1;
  // Header rows.
  const c0 = at.col;
  if (!vals.length && !p.rows.length && !p.cols.length) return { cells, rows: top - at.row, cols: 0 };
  if (p.cols.length) {
    // Above the row labels the value field's caption (only with row fields), beside it the column header.
    if (p.rows.length) cells.push({ row: top, col: c0, v: multiVal ? "Werte" : vals[0] ? valueCaption(names, vals[0], vals) : "", kind: "corner" });
    cells.push({ row: top, col: c0 + labelCols, v: "Spaltenbeschriftungen", kind: "header" });
  }
  const hr = top + headRows - 1;
  if (p.rows.length) {
    if (compact) cells.push({ row: hr, col: c0, v: "Zeilenbeschriftungen", kind: "header" });
    else p.rows.forEach((f, i) => cells.push({ row: hr, col: c0 + i, v: names[f], kind: "header" }));
  } else if (p.cols.length || vals.length) {
    if (!p.cols.length && vals.length) { /* values only: their captions as the header */ }
  }
  let col = c0 + labelCols;
  valCols.forEach((vc) => {
    const label = p.cols.length ? (multiVal ? `${vc.key.join(" ")} ${valueCaption(names, vals[vc.v], vals)}`.trim() : vc.key.join(" ")) : vals[vc.v] ? valueCaption(names, vals[vc.v], vals) : "";
    cells.push({ row: hr, col, v: label, kind: p.cols.length ? "colLabel" : "header" });
    col++;
  });
  if (rowGrand) for (let v = 0; v < Math.max(1, vals.length); v++) {
    cells.push({ row: hr, col, v: multiVal ? `${valueCaption(names, vals[v], vals)} insgesamt` : "Gesamtergebnis", kind: "grandLabel" });
    col++;
  }
  const width = col - c0;
  // Body rows.
  let row = hr + 1;
  const valueRow = (recs: PivotValue[][], kind: "value" | "subtotal" | "grand") => {
    let cc = c0 + labelCols;
    for (const vc of valCols) {
      const rs = p.cols.length ? recs.filter((r) => keyOf(r, p.cols).join("\u0000") === vc.key.join("\u0000")) : recs;
      const vf = vals[vc.v];
      if (vf) cells.push({ row, col: cc, v: rs.length ? summarise(rs.map((r) => r[vf.field] ?? null), vf.func) : null, kind, numFmt: vf.numFmt });
      cc++;
    }
    if (rowGrand) for (const vf of vals) { cells.push({ row, col: cc, v: summarise(recs.map((r) => r[vf.field] ?? null), vf.func), kind: kind === "value" ? "grand" : kind, numFmt: vf.numFmt }); cc++; }
  };
  const subtotals = p.subtotals ?? "top";
  const emit = (level: number, recs: PivotValue[][], prefix: string[]) => {
    if (level === p.rows.length) return;
    const f = p.rows[level];
    const items = sortItems([...new Map(recs.map((r) => [itemText(r[f] ?? null), r[f] ?? null])).values()]);
    for (const it of items) {
      const sub = recs.filter((r) => itemText(r[f] ?? null) === itemText(it));
      const last = level === p.rows.length - 1;
      const labelCol = compact ? c0 : c0 + level;
      cells.push({ row, col: labelCol, v: it === null || it === "" ? "(Leer)" : it, kind: last ? "rowLabel" : "subtotalLabel", level });
      if (!last && p.layout === "tabular") {
        // Tabellenformat: the first child on the same row.
        emit(level + 1, sub, [...prefix, itemText(it)]);
        if (subtotals !== "none") { cells.push({ row, col: labelCol, v: `${itemText(it)} Ergebnis`, kind: "subtotalLabel", level }); valueRow(sub, "subtotal"); row++; }
        continue;
      }
      if (last || subtotals === "top") valueRow(sub, last ? "value" : "subtotal");
      row++;
      if (!last) {
        emit(level + 1, sub, [...prefix, itemText(it)]);
        if (subtotals === "bottom") { cells.push({ row, col: labelCol, v: `${itemText(it)} Ergebnis`, kind: "subtotalLabel", level }); valueRow(sub, "subtotal"); row++; }
      }
    }
  };
  if (p.rows.length) emit(0, records, []);
  else if (vals.length) {
    // Columns or values only: the value's caption labels its row.
    if (p.cols.length) cells.push({ row, col: c0, v: valueCaption(names, vals[0], vals), kind: "rowLabel" });
    valueRow(records, "value");
    row++;
  }
  if (colGrand && p.rows.length) {
    cells.push({ row, col: c0, v: "Gesamtergebnis", kind: "grandLabel" });
    valueRow(records, "grand");
    row++;
  }
  return { cells, rows: row - at.row, cols: width };
}
