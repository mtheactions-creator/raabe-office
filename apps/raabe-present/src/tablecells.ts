// Table cell ranges, Zellen verbinden and Zellen teilen on a:tbl (D-072). PowerPoint keeps one a:tc per
// grid column in every row; a merged area is its first cell with gridSpan / rowSpan and the covered cells
// marked hMerge (to the right) and vMerge (below). Ranges always grow to whole merged areas, as PowerPoint's
// selection does.

import { PptxDom } from "@raabe/ooxml";
import * as T from "./textops";

const { kid, kids, num } = PptxDom;

export interface CellRange { r0: number; c0: number; r1: number; c1: number }

export function grid(tbl: Element): Element[][] {
  return kids(tbl, "tr").map((tr) => kids(tr, "tc"));
}

/** The first cell of the merged area that covers (r, c). */
export function origin(g: Element[][], r: number, c: number): { r: number; c: number } {
  for (;;) {
    const tc = g[r]?.[c];
    if (!tc) return { r, c };
    if (tc.getAttribute("hMerge") === "1" && c > 0) { c--; continue; }
    if (tc.getAttribute("vMerge") === "1" && r > 0) { r--; continue; }
    return { r, c };
  }
}

/** The area covered by the cell whose first cell is (r, c). */
export function extent(g: Element[][], r: number, c: number): CellRange {
  const tc = g[r][c];
  return { r0: r, c0: c, r1: r + (num(tc, "rowSpan") ?? 1) - 1, c1: c + (num(tc, "gridSpan") ?? 1) - 1 };
}

/** The rectangle spanned by two cells, grown until no merged area crosses its border. */
export function normalize(g: Element[][], a: { r: number; c: number }, b: { r: number; c: number }): CellRange {
  const x: CellRange = { r0: Math.min(a.r, b.r), c0: Math.min(a.c, b.c), r1: Math.max(a.r, b.r), c1: Math.max(a.c, b.c) };
  for (let changed = true; changed;) {
    changed = false;
    for (let r = x.r0; r <= x.r1; r++) for (let c = x.c0; c <= x.c1; c++) {
      const o = origin(g, r, c), e = extent(g, o.r, o.c);
      if (e.r0 < x.r0 || e.c0 < x.c0 || e.r1 > x.r1 || e.c1 > x.c1) {
        x.r0 = Math.min(x.r0, e.r0); x.c0 = Math.min(x.c0, e.c0); x.r1 = Math.max(x.r1, e.r1); x.c1 = Math.max(x.c1, e.c1);
        changed = true;
      }
    }
  }
  return x;
}

export const inRange = (x: CellRange | undefined, r: number, c: number): boolean => !!x && r >= x.r0 && r <= x.r1 && c >= x.c0 && c <= x.c1;

/** More than one cell: what Zellen verbinden needs. */
export function isMulti(g: Element[][], x: CellRange): boolean {
  const e = extent(g, x.r0, x.c0);
  return e.r1 !== x.r1 || e.c1 !== x.c1;
}

function bodyOf(tc: Element): Element | undefined { return kid(tc, "txBody"); }
const hasText = (tc: Element) => { const b = bodyOf(tc); return !!b && T.paras(b).some((p) => T.paraText(p) !== ""); };

/** Zellen verbinden: the texts of the non-empty cells become paragraphs of the first. */
export function merge(tbl: Element, x: CellRange): void {
  const g = grid(tbl);
  const first = g[x.r0][x.c0];
  let target = bodyOf(first);
  for (let r = x.r0; r <= x.r1; r++) for (let c = x.c0; c <= x.c1; c++) {
    const tc = g[r][c];
    if (tc !== first && hasText(tc)) {
      const b = bodyOf(tc)!;
      if (!target) { target = b.cloneNode(true) as Element; first.insertBefore(target, first.firstChild); T.replaceAll(target, ""); }
      const paras = T.paras(target);
      if (paras.length === 1 && T.paraText(paras[0]) === "") paras[0].remove();
      for (const p of T.paras(b)) if (T.paraText(p) !== "" || T.paras(b).length > 1) target.append(p.cloneNode(true));
      T.replaceAll(b, "");
    }
    for (const a of ["gridSpan", "rowSpan", "hMerge", "vMerge"]) tc.removeAttribute(a);
    if (c > x.c0) tc.setAttribute("hMerge", "1");
    if (r > x.r0) tc.setAttribute("vMerge", "1");
    if (c === x.c0 && x.c1 > x.c0) tc.setAttribute("gridSpan", String(x.c1 - x.c0 + 1));
  }
  // The cells below the first in its column carry gridSpan too (as PowerPoint writes them); rowSpan only the first.
  if (x.r1 > x.r0) first.setAttribute("rowSpan", String(x.r1 - x.r0 + 1));
}

function blank(tc: Element): Element {
  const n = tc.cloneNode(true) as Element;
  for (const a of ["gridSpan", "rowSpan", "hMerge", "vMerge"]) n.removeAttribute(a);
  const b = bodyOf(n);
  if (b) T.replaceAll(b, "");
  return n;
}

/** Inserts `k` grid columns after column j, splitting its width; rows outside `rows` stay one cell wide there. */
function insertCols(tbl: Element, j: number, k: number, rows: [number, number]): void {
  const cols = kids(kid(tbl, "tblGrid")!, "gridCol");
  const w = num(cols[j], "w") ?? 0, part = Math.round(w / (k + 1));
  cols[j].setAttribute("w", String(w - part * k));
  for (let i = 0; i < k; i++) { const n = cols[j].cloneNode(true) as Element; n.setAttribute("w", String(part)); cols[j].after(n); }
  const g = grid(tbl);
  g.forEach((row, r) => {
    const tc = row[j];
    const inside = r >= rows[0] && r <= rows[1];
    if (!inside && tc.getAttribute("vMerge") !== "1") {
      const o = origin(g, r, j);
      const first = g[o.r][o.c];
      first.setAttribute("gridSpan", String((num(first, "gridSpan") ?? 1) + k));
      // The rows below inside its area carry the same gridSpan on their first cell.
      const e = extent(g, o.r, o.c);
      for (let rr = o.r + 1; rr <= e.r1; rr++) g[rr][o.c].setAttribute("gridSpan", first.getAttribute("gridSpan")!);
    }
    for (let i = 0; i < k; i++) {
      const n = blank(tc);
      if (!inside) { n.setAttribute("hMerge", "1"); if (tc.getAttribute("vMerge") === "1") n.setAttribute("vMerge", "1"); }
      tc.after(n);
    }
  });
}

/** Inserts `k` rows after row i, splitting its height; columns outside `cols` stay one cell high there. */
function insertRows(tbl: Element, i: number, k: number, cols: [number, number]): void {
  const trs = kids(tbl, "tr");
  const h = num(trs[i], "h") ?? 0, part = Math.round(h / (k + 1));
  trs[i].setAttribute("h", String(h - part * k));
  const g = grid(tbl);
  const done = new Set<Element>();
  for (let n = 0; n < k; n++) {
    const tr = trs[i].cloneNode(false) as Element;
    tr.setAttribute("h", String(part));
    g[i].forEach((tc, c) => {
      const inside = c >= cols[0] && c <= cols[1];
      const x = blank(tc);
      if (!inside) {
        x.setAttribute("vMerge", "1");
        if (tc.getAttribute("hMerge") === "1") x.setAttribute("hMerge", "1");
        const span = tc.getAttribute("gridSpan");
        if (span && tc.getAttribute("hMerge") !== "1") x.setAttribute("gridSpan", span);
        if (tc.getAttribute("hMerge") !== "1") {
          const o = origin(g, i, c);
          const first = g[o.r][o.c];
          if (!done.has(first)) { done.add(first); first.setAttribute("rowSpan", String((num(first, "rowSpan") ?? 1) + k)); }
        }
      }
      tr.append(x);
    });
    // Rows after the table's own extension elements keep their order: insert right after row i.
    trs[i].after(tr);
  }
}

/**
 * Zellen teilen: the cell (r, c) becomes `cols` × `rows` cells. A merged cell is first taken apart; more
 * columns or rows than it covers are made by splitting its last column or row (the rest of the table keeps
 * its look, the cells beside grow over the new grid lines).
 */
export function split(tbl: Element, r: number, c: number, cols: number, rows: number): void {
  const g = grid(tbl);
  const o = origin(g, r, c);
  const area = extent(g, o.r, o.c);
  // Take the area apart.
  for (let rr = area.r0; rr <= area.r1; rr++) for (let cc = area.c0; cc <= area.c1; cc++) for (const a of ["gridSpan", "rowSpan", "hMerge", "vMerge"]) g[rr][cc].removeAttribute(a);
  let nc = area.c1 - area.c0 + 1, nr = area.r1 - area.r0 + 1;
  if (cols > nc) { insertCols(tbl, area.c1, cols - nc, [area.r0, area.r1]); nc = cols; }
  if (rows > nr) { insertRows(tbl, area.r1, rows - nr, [area.c0, area.c0 + nc - 1]); nr = rows; }
  // Group the area's grid columns and rows into cols × rows cells.
  const groups = (n: number, parts: number) => { const out: [number, number][] = []; let s = 0; for (let p = 0; p < parts; p++) { const len = Math.floor(n / parts) + (p < n % parts ? 1 : 0); out.push([s, s + len - 1]); s += len; } return out; };
  for (const [ra, rb] of groups(nr, Math.min(rows, nr))) for (const [ca, cb] of groups(nc, Math.min(cols, nc))) {
    const x = { r0: area.r0 + ra, r1: area.r0 + rb, c0: area.c0 + ca, c1: area.c0 + cb };
    if (x.r1 > x.r0 || x.c1 > x.c0) merge(tbl, x);
  }
}

/** Entf / Rücktaste over a cell range: empties every cell. */
export function clear(tbl: Element, x: CellRange): void {
  const g = grid(tbl);
  for (let r = x.r0; r <= x.r1; r++) for (let c = x.c0; c <= x.c1; c++) { const b = bodyOf(g[r][c]); if (b) T.replaceAll(b, ""); }
}
