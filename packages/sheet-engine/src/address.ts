// Cell addresses as Excel writes them: columns A..XFD, rows 1..1048576. Inside the engine
// rows and columns are 0-based numbers.

export const MAX_ROWS = 1048576;
export const MAX_COLS = 16384;

export interface CellRef { row: number; col: number }
/** A rectangle of cells, inclusive on both ends. */
export interface Rect { r0: number; c0: number; r1: number; c1: number }

/** 0 → "A", 25 → "Z", 26 → "AA". */
export function colName(col: number): string {
  let s = "";
  let n = col + 1;
  while (n > 0) {
    const m = (n - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

/** "A" → 0, "aa" → 26; -1 when not a column name. */
export function colIndex(name: string): number {
  let n = 0;
  for (const ch of name.toUpperCase()) {
    const c = ch.charCodeAt(0);
    if (c < 65 || c > 90) return -1;
    n = n * 26 + (c - 64);
  }
  return n - 1;
}

export function cellName(row: number, col: number): string {
  return colName(col) + (row + 1);
}

/** "B3" → {row: 2, col: 1}; also "$B$3". null when it is not a cell address. */
export function parseCell(text: string): CellRef | null {
  const m = /^\$?([A-Za-z]{1,3})\$?(\d{1,7})$/.exec(text.trim());
  if (!m) return null;
  const col = colIndex(m[1]);
  const row = Number(m[2]) - 1;
  if (col < 0 || col >= MAX_COLS || row < 0 || row >= MAX_ROWS) return null;
  return { row, col };
}

/** "A1:C3", "B2", "A:C", "2:5" → a normalised rectangle; null otherwise. */
export function parseRange(text: string): Rect | null {
  const t = text.trim();
  const parts = t.split(":");
  if (parts.length === 1) {
    const c = parseCell(t);
    return c ? { r0: c.row, c0: c.col, r1: c.row, c1: c.col } : null;
  }
  if (parts.length !== 2) return null;
  const [a, b] = parts.map((p) => p.replace(/\$/g, ""));
  if (/^[A-Za-z]{1,3}$/.test(a) && /^[A-Za-z]{1,3}$/.test(b)) {
    const c0 = colIndex(a), c1 = colIndex(b);
    if (c0 < 0 || c1 < 0 || c0 >= MAX_COLS || c1 >= MAX_COLS) return null;
    return { r0: 0, c0: Math.min(c0, c1), r1: MAX_ROWS - 1, c1: Math.max(c0, c1) };
  }
  if (/^\d+$/.test(a) && /^\d+$/.test(b)) {
    const r0 = Number(a) - 1, r1 = Number(b) - 1;
    if (r0 < 0 || r1 < 0 || r0 >= MAX_ROWS || r1 >= MAX_ROWS) return null;
    return { r0: Math.min(r0, r1), c0: 0, r1: Math.max(r0, r1), c1: MAX_COLS - 1 };
  }
  const p = parseCell(a), q = parseCell(b);
  if (!p || !q) return null;
  return normRect(p.row, p.col, q.row, q.col);
}

export function normRect(ra: number, ca: number, rb: number, cb: number): Rect {
  return { r0: Math.min(ra, rb), c0: Math.min(ca, cb), r1: Math.max(ra, rb), c1: Math.max(ca, cb) };
}

/** Excel's way of writing a range: "B2", "B2:D5", "A:A", "3:3". */
export function rangeName(r: Rect): string {
  if (r.r0 === 0 && r.r1 === MAX_ROWS - 1) return `${colName(r.c0)}:${colName(r.c1)}`;
  if (r.c0 === 0 && r.c1 === MAX_COLS - 1) return `${r.r0 + 1}:${r.r1 + 1}`;
  const a = cellName(r.r0, r.c0);
  return r.r0 === r.r1 && r.c0 === r.c1 ? a : `${a}:${cellName(r.r1, r.c1)}`;
}

export function inRect(r: Rect, row: number, col: number): boolean {
  return row >= r.r0 && row <= r.r1 && col >= r.c0 && col <= r.c1;
}

/** Key of a cell in a sheet's map (row-major, so sorting keys sorts by row, then column). */
export function cellKey(row: number, col: number): number {
  return row * MAX_COLS + col;
}

export function keyCell(key: number): CellRef {
  return { row: Math.floor(key / MAX_COLS), col: key % MAX_COLS };
}
