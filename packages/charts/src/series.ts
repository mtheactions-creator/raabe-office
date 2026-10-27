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

import type { ChartKind, ChartSeries } from "./model";

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
export function seriesFromBlock(b: Block, kind: ChartKind): ChartSeries[] {
  const rows = b.r1 - b.r0 + 1, cols = b.c1 - b.c0 + 1;
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
  const byCols = rows > cols;
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
  } else if ((kind === "scatter" || kind === "scatterSmooth" || kind === "scatterLines") && lines.length > 1) {
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
