// How Excel turns a selected block of cells into series, as recorded with COM (probe section
// "series"):
// - more rows than columns: one series per column; otherwise (also when square) one per row;
// - an empty top-left cell makes the first row and column labels;
// - otherwise the first row holds the series names when its cells (after the first) are all
//   text, and the first column holds the categories when its cells (below the first) are;
//   with one row or column the test covers all its cells;
// - scatter charts take the categories as X values; without them the first of several data
//   lines becomes the X values;
// - histograms take every number of the block as one series.

import type { ChartKind, ChartModel, ChartSeries } from "./model";
import { kindInfo } from "./kinds";

export type CellValue = string | number | boolean | null;

export interface Block {
  sheet: string;
  r0: number; c0: number; r1: number; c1: number;
  get(row: number, col: number): CellValue;
}

export function colName(col: number): string {
  let s = "";
  for (let c = col + 1; c > 0; c = Math.floor((c - 1) / 26)) s = String.fromCharCode(65 + ((c - 1) % 26)) + s;
  return s;
}

function colIndex(name: string): number {
  let n = 0;
  for (const ch of name.toUpperCase()) n = n * 26 + ch.charCodeAt(0) - 64;
  return n - 1;
}

/** A sheet name as formulas write it ('Meine Daten'!…). */
export function quoteSheet(name: string): string {
  return /^[A-Za-z_À-ɏ][\w.À-ɏ]*$/.test(name) && !/^[A-Za-z]{1,3}\d+$/.test(name) ? name : `'${name.replace(/'/g, "''")}'`;
}

/** "Tabelle1!$B$2:$B$6" ("Tabelle1!$B$1" for one cell). */
export function refText(sheet: string, r0: number, c0: number, r1 = r0, c1 = c0): string {
  const a = `$${colName(c0)}$${r0 + 1}`;
  return `${quoteSheet(sheet)}!${a}${r0 === r1 && c0 === c1 ? "" : `:$${colName(c1)}$${r1 + 1}`}`;
}

export interface Ref { sheet: string; r0: number; c0: number; r1: number; c1: number }

/** Reads "Tabelle1!$B$2:$B$6" or "'My sheet'!B2"; null for anything else (names, several areas). */
export function parseRef(ref: string): Ref | null {
  const m = /^(?:'((?:[^']|'')+)'|([^!'"]+))!\$?([A-Za-z]{1,3})\$?(\d+)(?::\$?([A-Za-z]{1,3})\$?(\d+))?$/.exec(ref.trim().replace(/^\(|\)$/g, ""));
  if (!m) return null;
  const sheet = m[1] !== undefined ? m[1].replace(/''/g, "'") : m[2];
  const a = { r: Number(m[4]) - 1, c: colIndex(m[3]) };
  const b = m[5] ? { r: Number(m[6]) - 1, c: colIndex(m[5]) } : a;
  return { sheet, r0: Math.min(a.r, b.r), c0: Math.min(a.c, b.c), r1: Math.max(a.r, b.r), c1: Math.max(a.c, b.c) };
}

const isText = (v: CellValue) => typeof v === "string" && v !== "";
const num = (v: CellValue): number | null => (typeof v === "number" && isFinite(v) ? v : null);

/** The series Excel makes of a block for a chart of `kind`. */
export function seriesFromBlock(b: Block, kind: ChartKind, orientation?: "rows" | "cols"): ChartSeries[] {
  const rows = b.r1 - b.r0 + 1, cols = b.c1 - b.c0 + 1;
  const fam = kindInfo(kind).family;
  if (fam === "treemap" || fam === "sunburst") return hierarchySeries(b);
  // Wasserfall, Trichter and Pareto show one series (the first value column).
  if (fam === "waterfall" || fam === "funnel" || fam === "pareto") return seriesFromBlock(b, "column", orientation).slice(0, 1);
  if (fam === "bubble") {
    // Blase: X values, then Y values and Blasengröße in pairs (Excel: one series of three columns).
    const xy = seriesFromBlock(b, "scatter", orientation);
    const out: ChartSeries[] = [];
    for (let i = 0; i < xy.length; i += 2) {
      const s = xy[i];
      if (xy[i + 1]) s.size = { ref: xy[i + 1].val.ref, cache: xy[i + 1].val.cache };
      out.push(s);
    }
    return out;
  }
  if (fam === "stock" && cols > 1 && rows > 1 && !orientation) {
    // Kurs: the first column holds the dates (numbers) when there are more columns than the type has prices.
    const need = { stockHLC: 3, stockOHLC: 4, stockVHLC: 4, stockVOHLC: 5 }[kind as "stockHLC"] ?? 3;
    if (cols > need) {
      const out = seriesFromBlock({ ...b, c0: b.c0 + 1 }, "line", "cols");
      const head = isText(b.get(b.r0, b.c1));
      const r0 = head ? b.r0 + 1 : b.r0;
      const cache: (string | number | null)[] = [];
      for (let r = r0; r <= b.r1; r++) { const v = b.get(r, b.c0); cache.push(v === null ? null : typeof v === "number" ? v : String(v)); }
      for (const s of out) s.cat = { ref: refText(b.sheet, r0, b.c0, b.r1, b.c0), cache: [...cache] };
      return out;
    }
  }
  if (kind === "histogram") {
    const vals: number[] = [];
    let name: ChartSeries["name"];
    for (let r = b.r0; r <= b.r1; r++) {
      for (let c = b.c0; c <= b.c1; c++) {
        const v = b.get(r, c);
        if (typeof v === "number") vals.push(v);
        else if (r === b.r0 && c === b.c0 && isText(v)) name = { ref: refText(b.sheet, r, c), text: String(v) };
      }
    }
    const first = name ? (cols === 1 ? b.r0 + 1 : b.r0) : b.r0;
    return [{ name, val: { ref: refText(b.sheet, first, b.c0, b.r1, b.c1), cache: vals } }];
  }
  // Series in columns when the block is taller than wide (Excel), unless Zeile/Spalte wechseln chose.
  const byCols = orientation ? orientation === "cols" : rows > cols;
  let headRow = false, headCol = false;
  if (rows > 1 && cols > 1 && (b.get(b.r0, b.c0) === null || b.get(b.r0, b.c0) === "")) {
    headRow = headCol = true;
  } else {
    if (rows > 1) {
      headRow = true;
      for (let c = cols > 1 ? b.c0 + 1 : b.c0; c <= b.c1; c++) if (!isText(b.get(b.r0, c))) headRow = false;
    }
    if (cols > 1) {
      headCol = true;
      for (let r = rows > 1 ? b.r0 + 1 : b.r0; r <= b.r1; r++) if (!isText(b.get(r, b.c0))) headCol = false;
    }
  }
  // Names run along the edge across the series, categories along the series.
  const nameEdge = byCols ? headRow : headCol;
  const catEdge = byCols ? headCol : headRow;
  const lines: number[] = [];
  const first = byCols ? b.c0 + (headCol ? 1 : 0) : b.r0 + (headRow ? 1 : 0);
  const last = byCols ? b.c1 : b.r1;
  for (let i = first; i <= last; i++) lines.push(i);
  const p0 = byCols ? b.r0 + (headRow ? 1 : 0) : b.c0 + (headCol ? 1 : 0);
  const p1 = byCols ? b.r1 : b.c1;
  const cell = (line: number, p: number) => (byCols ? b.get(p, line) : b.get(line, p));
  const ref = (line: number, q0: number, q1: number) => (byCols ? refText(b.sheet, q0, line, q1, line) : refText(b.sheet, line, q0, line, q1));
  const points = (line: number) => {
    const out: (number | null)[] = [];
    for (let p = p0; p <= p1; p++) out.push(num(cell(line, p)));
    return out;
  };
  let cat: ChartSeries["cat"];
  if (catEdge) {
    const line = byCols ? b.c0 : b.r0;
    const cache: (string | number | null)[] = [];
    for (let p = p0; p <= p1; p++) {
      const v = byCols ? b.get(p, line) : b.get(line, p);
      cache.push(v === null ? null : typeof v === "number" ? v : String(v));
    }
    cat = { ref: ref(line, p0, p1), cache };
  } else if (kindInfo(kind).family === "scatter" && lines.length > 1) {
    const line = lines.shift()!;
    cat = { ref: ref(line, p0, p1), cache: points(line) };
  }
  return lines.map((line) => {
    const s: ChartSeries = { val: { ref: ref(line, p0, p1), cache: points(line) } };
    if (cat) s.cat = { ref: cat.ref, cache: [...cat.cache] };
    if (nameEdge) {
      const nr = byCols ? b.r0 : line, nc = byCols ? line : b.c0;
      const v = b.get(nr, nc);
      s.name = { ref: refText(b.sheet, nr, nc), text: v === null ? "" : String(v) };
    }
    return s;
  });
}

/**
 * Treemap and Sunburst: the text columns are the category levels (outermost first, a blank cell
 * repeats the one above), the last numeric column the values; a text first row holds the names.
 */
function hierarchySeries(b: Block): ChartSeries[] {
  const head = b.r1 > b.r0 && isText(b.get(b.r0, b.c1));
  const r0 = head ? b.r0 + 1 : b.r0;
  let vc = b.c1;
  while (vc > b.c0 && ![...Array(b.r1 - r0 + 1)].some((_, k) => typeof b.get(r0 + k, vc) === "number")) vc--;
  const levels: (string | null)[][] = [];
  for (let c = b.c0; c < vc; c++) {
    const lv: (string | null)[] = [];
    let last: string | null = null;
    for (let r = r0; r <= b.r1; r++) {
      const v = b.get(r, c);
      if (v !== null && v !== "") last = String(v);
      lv.push(v === null || v === "" ? (c < vc - 1 ? last : null) : String(v));
    }
    levels.push(lv);
  }
  const vals: (number | null)[] = [];
  for (let r = r0; r <= b.r1; r++) vals.push(num(b.get(r, vc)));
  const s: ChartSeries = { val: { ref: refText(b.sheet, r0, vc, b.r1, vc), cache: vals } };
  if (levels.length) {
    s.levels = levels;
    s.cat = { ref: refText(b.sheet, r0, b.c0, b.r1, vc - 1), cache: levels.at(-1)! };
  }
  if (head) s.name = { ref: refText(b.sheet, b.r0, vc), text: String(b.get(b.r0, vc)) };
  return [s];
}

/** The block a chart's series come from (the bounding range of their names, categories and values) and how they run. */
export function sourceOf(series: ChartSeries[]): { ref: Ref; orientation: "rows" | "cols" } | null {
  const refs = series.flatMap((x) => [x.name?.ref, x.cat?.ref, x.val.ref]).filter((r): r is string => !!r).map(parseRef).filter((r): r is Ref => !!r);
  if (!refs.length) return null;
  const sheet = refs[0].sheet;
  const same = refs.filter((r) => r.sheet === sheet);
  const ref = { sheet, r0: Math.min(...same.map((r) => r.r0)), c0: Math.min(...same.map((r) => r.c0)), r1: Math.max(...same.map((r) => r.r1)), c1: Math.max(...same.map((r) => r.c1)) };
  const v = parseRef(series[0].val.ref ?? "");
  const orientation = v && v.c0 === v.c1 && v.r1 > v.r0 ? "cols" : v && v.r0 === v.r1 && v.c1 > v.c0 ? "rows" : "cols";
  return { ref, orientation };
}

/** Cells of a range as read from the workbook: row by row, `cols` across. */
export interface RangeRead { vals: CellValue[]; format?: string; cols: number }

/**
 * Re-reads a chart's names, categories, values and bubble sizes through `read` (Excel's caches).
 * Treemap and Sunburst categories of several columns become their levels (a blank cell of an
 * outer level repeats the one above).
 */
export function refreshCaches(m: ChartModel, read: (ref: string) => RangeRead | null, nameText: (v: CellValue, format?: string) => string): void {
  const hier = m.kind === "treemap" || m.kind === "sunburst";
  for (const s of m.series) {
    if (s.name?.ref) {
      const n = read(s.name.ref);
      if (n) s.name.text = n.vals[0] === null || n.vals[0] === undefined ? "" : nameText(n.vals[0], n.format);
    }
    if (s.cat?.ref) {
      const c = read(s.cat.ref);
      if (c) {
        const cells = c.vals.map((v) => (typeof v === "boolean" ? (v ? "WAHR" : "FALSCH") : v));
        if (hier && c.cols > 1) {
          const rows = Math.ceil(cells.length / c.cols);
          const levels: (string | null)[][] = [];
          for (let k = 0; k < c.cols; k++) {
            let last: string | null = null;
            levels.push(Array.from({ length: rows }, (_, r) => {
              const v = cells[r * c.cols + k];
              if (v !== null && v !== undefined && v !== "") last = String(v);
              return v === null || v === undefined || v === "" ? (k < c.cols - 1 ? last : null) : String(v);
            }));
          }
          s.levels = levels;
          s.cat.cache = levels.at(-1)!;
        } else {
          s.cat.cache = cells;
          if (hier) delete s.levels;
        }
        s.cat.format = c.format;
      }
    }
    const v = read(s.val.ref ?? "");
    if (v) {
      s.val.cache = v.vals.map((x) => (typeof x === "number" ? x : null));
      s.val.format = v.format;
    }
    if (s.size?.ref) {
      const z = read(s.size.ref);
      if (z) s.size.cache = z.vals.map((x) => (typeof x === "number" ? x : null));
    }
    for (const d of [s.errorBars?.plus, s.errorBars?.minus]) {
      const r = d?.ref ? read(d.ref) : null;
      if (d && r) d.cache = r.vals.map((x) => (typeof x === "number" ? x : null));
    }
  }
}
