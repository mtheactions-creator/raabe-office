// Document operations behind Step 3's commands: lists, tables, section breaks, fields and
// pictures. They work on the active story of a Doc and replace paragraph objects (never
// mutate them), so the editor can take an undo snapshot before calling them.

import { Doc, nextTableId, type Pos } from "./doc";
import {
  EMPTY, OBJ, intern, joinParas, makePara, type CellDef, type CellRef, type CharProps, type ImageData, type InlineObject, type NumLevel,
  type Numbering, type Para, type ParaProps, type RowDef, type Section, type TableDef,
} from "./model";
import { addList, bulletLevels, levelOf, multilevelLevels, numberLevels, numberingOf } from "./numbering";

// ----- lists ---------------------------------------------------------------------------------

export type ListKind = "bullet" | "number" | "multilevel";

/** What kind of list a paragraph is in, or null. */
export function listKind(doc: Doc, p: Para): ListKind | null {
  const n = numberingOf(p.props, doc.styles);
  if (!n) return null;
  const found = levelOf(doc.numbering, n.numId, n.ilvl);
  if (!found) return null;
  if (found.level.format === "bullet") return "bullet";
  const lvl1 = found.abstract.levels[1];
  return lvl1 && /%1/.test(lvl1.text) ? "multilevel" : "number";
}

/**
 * Aufzählungszeichen / Nummerierung / Liste mit mehreren Ebenen. As in Word: paragraphs that
 * are all in such a list already leave it; otherwise they join the list of the paragraph
 * before (when it is the same kind) or start a new one, and Standard paragraphs get the
 * Listenabsatz style (measured with Word's FormatBulletDefault).
 */
export function toggleList(doc: Doc, a: number, b: number, kind: ListKind, listStyle: string, levels?: NumLevel[]): void {
  const paras = doc.paras;
  const all = paras.slice(a, b + 1).filter((p) => !p.block);
  if (!levels && all.length && all.every((p) => listKind(doc, p) === kind)) { removeList(doc, a, b, listStyle); return; }
  let numId: number | undefined;
  const prev = paras[a - 1];
  if (!levels && prev && listKind(doc, prev) === kind) numId = numberingOf(prev.props, doc.styles)!.numId;
  if (numId === undefined) {
    const made = addList(doc.numbering, levels ?? (kind === "bullet" ? bulletLevels() : kind === "number" ? numberLevels() : multilevelLevels()));
    doc.setNumbering(made.numbering);
    numId = made.numId;
  }
  const normal = doc.styles.defaultParagraphStyle;
  doc.mapParas(a, b, (old) => {
    const style = old.style ?? normal;
    // The level's indents apply (direct indents go, as in Word).
    const { indLeft: _l, indFirst: _f, ...rest } = old;
    return { ...rest, style: style === normal ? listStyle : style, numId, ilvl: old.ilvl ?? 0 };
  });
}

/** Removes numbering from paragraphs a..b (the Listenabsatz style returns to Standard). */
export function removeList(doc: Doc, a: number, b: number, listStyle?: string): void {
  const normal = doc.styles.defaultParagraphStyle;
  doc.mapParas(a, b, (old) => {
    const viaStyle = old.numId === undefined && !!doc.styles.numberingOfStyle(old.style ?? normal);
    const { numId: _n, ilvl: _i, indLeft: _l, indFirst: _f, ...rest } = old;
    const style = listStyle && old.style === listStyle ? undefined : old.style;
    return viaStyle ? { ...rest, style, numId: 0 } : { ...rest, style };
  });
}

/** Tab / Umschalt+Tab at the start of a list item: one level deeper or higher. */
export function changeListLevel(doc: Doc, a: number, b: number, delta: 1 | -1): boolean {
  let changed = false;
  doc.mapParas(a, b, (old, p) => {
    const n = numberingOf(p.props, doc.styles);
    if (!n) return old;
    const ilvl = Math.max(0, Math.min(8, n.ilvl + delta));
    if (ilvl === n.ilvl) return old;
    changed = true;
    return { ...old, numId: old.numId ?? n.numId, ilvl };
  });
  return changed;
}

/**
 * Backspace at the start of a list item: the number goes, the text stays where it was (the
 * paragraph keeps the level's indent).
 */
export function dropListLabel(doc: Doc, p: number): void {
  const para = doc.paras[p];
  const n = numberingOf(para.props, doc.styles);
  if (!n) return;
  const lvl = levelOf(doc.numbering, n.numId, n.ilvl)?.level;
  const left = lvl?.pPr.indLeft ?? doc.styles.resolvePara(para.props).indLeft;
  doc.mapParas(p, p, (old) => ({ ...old, numId: 0, ilvl: undefined, indLeft: left, indFirst: 0 }));
}

/**
 * AutoFormat beim Eingeben: a label typed at a paragraph start (and a space or Tab) begins a
 * list. Measured in Word 365 (2026-10-14): bullets from "*" (Symbol), ">" "->" "=>"
 * (Wingdings) and the dashes "-" "–" "—" (the typed character in the text's font); numbers
 * only from the first value (1, a, A, i, I) followed by "." ")" "-" or in brackets "(1)".
 * Roman numerals get a wider hanging indent (1.27 cm at 1.9 cm).
 */
export function autoListLevels(prefix: string, textFont?: string): { levels: NumLevel[]; start: number } | null {
  const lvl = (format: string, text: string, font?: string, wide = false): NumLevel => ({
    start: 1, format, text, align: "left", suffix: "tab", pPr: wide ? { indLeft: 54, indFirst: -36 } : { indLeft: 36, indFirst: -18 }, rPr: font ? { font } : {},
  });
  const with9 = (first: NumLevel) => {
    const base = first.format === "bullet" ? bulletLevels() : numberLevels();
    return { levels: [first, ...base.slice(1)], start: 1 };
  };
  const bullets: Record<string, [string, string]> = {
    "*": ["\uf0b7", "Symbol"], ">": ["\uf0d8", "Wingdings"], "->": ["\uf0e8", "Wingdings"], "=>": ["\uf0f0", "Wingdings"],
  };
  if (bullets[prefix]) return with9(lvl("bullet", bullets[prefix][0], bullets[prefix][1]));
  if (prefix === "-" || prefix === "\u2013" || prefix === "\u2014") return with9(lvl("bullet", prefix, textFont ?? "Calibri"));
  const m = /^(\()?(1|a|A|i|I)(?:(\))|([.)-]))$/.exec(prefix);
  if (!m) return null;
  const v = m[2];
  const format = v === "1" ? "decimal" : v === "a" ? "lowerLetter" : v === "A" ? "upperLetter" : v === "i" ? "lowerRoman" : "upperRoman";
  const text = m[1] ? "(%1)" : `%1${m[3] ?? m[4]}`;
  if (m[1] && !m[3]) return null;
  return with9(lvl(format, text, undefined, v === "i" || v === "I"));
}

/**
 * Tab (Umschalt+Tab) at the start of a list's first item moves the whole list: the number
 * goes to the next (previous) stop, stops being the default tab stops and 0.635 cm after
 * each (Word 365, measured: 0.63 → 1.25 → 1.88 → 2.5 cm). All levels move by as much.
 */
export function shiftList(doc: Doc, p: number, back: boolean): boolean {
  const para = doc.paras[p];
  const n = numberingOf(para.props, doc.styles);
  if (!n) return false;
  const found = levelOf(doc.numbering, n.numId, n.ilvl);
  if (!found) return false;
  const rp = doc.styles.resolvePara(para.props, { level: found.level.pPr });
  const tw = (x: number) => Math.round(x * 20);
  const pos = tw(rp.indLeft + Math.min(0, rp.indFirst));
  const step = Math.max(1, tw(doc.settings.defaultTab));
  let next: number | null = null;
  for (let k = 0; k * step <= pos + 2 * step + 720; k++) {
    for (const s of [k * step, k * step + 360]) {
      if (!back && s > pos && (next === null || s < next)) next = s;
      if (back && s < pos && s >= 0 && (next === null || s > next)) next = s;
    }
  }
  if (next === null) return false;
  const delta = (next - pos) / 20;
  const abs = found.abstract;
  const levels = abs.levels.map((l) => ({ ...l, pPr: { ...l.pPr, indLeft: Math.max(0, (l.pPr.indLeft ?? 0) + delta) }, data: undefined }));
  const abstracts = new Map(doc.numbering.abstracts);
  abstracts.set(abs.id, { ...abs, levels, data: undefined });
  doc.setNumbering({ ...doc.numbering, abstracts });
  // Paragraphs of the list with their own indent move along.
  for (let i = 0; i < doc.paras.length; i++) {
    const q = doc.paras[i];
    if (q.props.indLeft === undefined) continue;
    const m = numberingOf(q.props, doc.styles);
    if (!m || levelOf(doc.numbering, m.numId, m.ilvl)?.abstract.id !== abs.id) continue;
    doc.mapParas(i, i, (old) => ({ ...old, indLeft: Math.max(0, (old.indLeft ?? 0) + delta) }));
  }
  return true;
}

/** True when no paragraph before p belongs to p's list (p is its first item). */
export function firstOfList(doc: Doc, p: number): boolean {
  const n = numberingOf(doc.paras[p].props, doc.styles);
  if (!n) return false;
  for (let i = p - 1; i >= 0; i--) {
    const m = numberingOf(doc.paras[i].props, doc.styles);
    if (m && m.numId === n.numId) return false;
  }
  return true;
}

// ----- tables ---------------------------------------------------------------------------------

/** The table around paragraph p: its range and definition. */
export function tableAt(doc: Doc, p: number): { a: number; b: number; def: TableDef } | null {
  const cell = doc.paras[p]?.cell;
  if (!cell) return null;
  const r = doc.tableRange(p, cell.table);
  return r ? { a: r[0], b: r[1], def: cell.table } : null;
}

type CellParas = Para[][][];

/** The paragraphs of each cell (rows × cells), without their cell references. */
function cellParas(doc: Doc, a: number, b: number, def: TableDef): CellParas {
  const out: CellParas = def.rows.map((row) => row.cells.map(() => []));
  for (let i = a; i <= b; i++) {
    const p = doc.paras[i];
    const c = p.cell!;
    out[c.row]?.[c.col]?.push({ ...p, cell: undefined });
  }
  for (const row of out) for (let c = 0; c < row.length; c++) if (!row[c].length) row[c] = [makePara("")];
  return out;
}

/** Replaces the table a..b with `def` and its cells' paragraphs. */
function rebuild(doc: Doc, a: number, b: number, def: TableDef, cells: CellParas): void {
  const paras: Para[] = [];
  cells.forEach((row, r) => row.forEach((list, c) => {
    const ref: CellRef = { table: def, row: r, col: c };
    for (const p of list.length ? list : [makePara("")]) paras.push({ ...p, cell: ref });
  }));
  doc.splice(a, b, paras);
}

/** An empty paragraph like the first one of a cell (same style, same mark). */
const emptyLike = (list: Para[]): Para => {
  const p = list[0];
  return p ? { ...makePara("", { style: p.props.style }, EMPTY, p.mark) } : makePara("");
};

const gridStart = (row: RowDef, col: number) => row.cells.slice(0, col).reduce((s, c) => s + Math.max(1, c.span), 0);

/** Index of the cell of `row` that covers grid column gx. */
function cellAtGrid(row: RowDef, gx: number): number {
  let x = 0;
  for (let c = 0; c < row.cells.length; c++) {
    const span = Math.max(1, row.cells[c].span);
    if (gx >= x && gx < x + span) return c;
    x += span;
  }
  return row.cells.length - 1;
}

/**
 * Inserts a rows × cols table at the caret ("Tabelle einfügen"): before the paragraph when the
 * caret is at its start, else after splitting it. Returns the position in the first cell.
 */
export function insertTable(doc: Doc, at: Pos, rows: number, cols: number, width: number, style: string | undefined, mark: CharProps): Pos {
  const paras = doc.paras;
  const src = paras[at.p];
  if (src.cell || src.block) return at;
  const w = width / cols;
  const def: TableDef = {
    id: nextTableId(), grid: Array.from({ length: cols }, () => w), style,
    rows: Array.from({ length: rows }, (): RowDef => ({ cells: Array.from({ length: cols }, (): CellDef => ({ span: 1, width: w })) })),
    look: { firstRow: true, lastRow: false, firstColumn: true, lastColumn: false, noHBand: false, noVBand: true },
  };
  const cells: CellParas = def.rows.map((row) => row.cells.map(() => [makePara("", {}, EMPTY, mark)]));
  const table: Para[] = [];
  cells.forEach((row, r) => row.forEach((list, c) => { for (const p of list) table.push({ ...p, cell: { table: def, row: r, col: c } }); }));
  let insertAt = at.p;
  if (at.o > 0) {
    // Split the paragraph: its first part stays before the table.
    doc.insertText(at, at, "\n", mark);
    insertAt = at.p + 1;
  }
  doc.splice(insertAt, insertAt - 1, table);
  return { p: insertAt, o: 0 };
}

export type RowWhere = "above" | "below";
export type ColWhere = "left" | "right";

/** Zeilen darüber/darunter einfügen (one per selected row, as Word). */
export function insertRows(doc: Doc, p: number, q: number, where: RowWhere): Pos | null {
  const t = tableAt(doc, p);
  if (!t) return null;
  const r1 = doc.paras[p].cell!.row, r2 = doc.paras[q]?.cell?.table.id === t.def.id ? doc.paras[q].cell!.row : r1;
  const lo = Math.min(r1, r2), hi = Math.max(r1, r2);
  const count = hi - lo + 1;
  const cells = cellParas(doc, t.a, t.b, t.def);
  const base = where === "above" ? lo : hi;
  const at = where === "above" ? lo : hi + 1;
  const tmpl = t.def.rows[base];
  const below = t.def.rows[at];
  const newRows: RowDef[] = [], newCells: Para[][][] = [];
  for (let k = 0; k < count; k++) {
    newRows.push({
      ...tmpl, header: where === "above" ? tmpl.header : undefined,
      cells: tmpl.cells.map((c, ci) => {
        // Inside a vertical merge the new cell continues it.
        const gx = gridStart(tmpl, ci);
        const nb = below ? below.cells[cellAtGrid(below, gx)] : undefined;
        const inside = c.vMerge && nb?.vMerge === "continue" && where === "below" ? "continue" as const : c.vMerge === "continue" && where === "above" ? "continue" as const : undefined;
        return { ...c, vMerge: inside };
      }),
    });
    newCells.push(tmpl.cells.map((_, ci) => [emptyLike(cells[base][ci])]));
  }
  const rows = [...t.def.rows.slice(0, at), ...newRows, ...t.def.rows.slice(at)];
  const all = [...cells.slice(0, at), ...newCells, ...cells.slice(at)];
  const def: TableDef = { ...t.def, rows };
  rebuild(doc, t.a, t.b, def, all);
  return firstOfCell(doc, t.a, at, 0);
}

/** Spalten links/rechts einfügen; the table keeps its width (columns share it). */
export function insertColumns(doc: Doc, p: number, q: number, where: ColWhere): Pos | null {
  const t = tableAt(doc, p);
  if (!t) return null;
  const c1 = doc.paras[p].cell!, c2 = doc.paras[q]?.cell?.table.id === t.def.id ? doc.paras[q].cell! : c1;
  const g1 = gridStart(t.def.rows[c1.row], c1.col), g2 = gridStart(t.def.rows[c2.row], c2.col);
  const lo = Math.min(g1, g2);
  const hiCell = g1 >= g2 ? c1 : c2;
  const hi = gridStart(t.def.rows[hiCell.row], hiCell.col) + Math.max(1, t.def.rows[hiCell.row].cells[hiCell.col].span) - 1;
  const count = hi - lo + 1;
  const gi = where === "left" ? lo : hi + 1;
  const cells = cellParas(doc, t.a, t.b, t.def);
  const total = t.def.grid.reduce((s, x) => s + x, 0);
  const added = Array.from({ length: count }, (_, k) => t.def.grid[Math.min(t.def.grid.length - 1, (where === "left" ? lo : hi - count + 1) + k)]);
  const grid0 = [...t.def.grid.slice(0, gi), ...added, ...t.def.grid.slice(gi)];
  const scale = total / grid0.reduce((s, x) => s + x, 0);
  const grid = grid0.map((x) => x * scale);
  const rows = t.def.rows.map((row, r) => {
    let x = 0;
    const out: CellDef[] = [];
    const outCells: Para[][] = [];
    let inserted = false;
    row.cells.forEach((c, ci) => {
      const span = Math.max(1, c.span);
      if (!inserted && x === gi) {
        for (let k = 0; k < count; k++) { out.push({ span: 1 }); outCells.push([emptyLike(cells[r][ci])]); }
        inserted = true;
      }
      if (!inserted && gi > x && gi < x + span) { out.push({ ...c, span: span + count, width: undefined }); outCells.push(cells[r][ci]); inserted = true; }
      else { out.push({ ...c, width: undefined }); outCells.push(cells[r][ci]); }
      x += span;
    });
    if (!inserted) for (let k = 0; k < count; k++) { out.push({ span: 1 }); outCells.push([emptyLike(cells[r][row.cells.length - 1])]); }
    cells[r] = outCells;
    return { ...row, cells: out };
  });
  rebuild(doc, t.a, t.b, { ...t.def, grid, rows }, cells);
  const r0 = c1.row;
  return firstOfCell(doc, t.a, r0, cellAtGrid(rows[r0], gi));
}

/** Zeilen löschen: the rows from p's to q's. Returns where the caret goes. */
export function deleteRows(doc: Doc, p: number, q: number): Pos | null {
  const t = tableAt(doc, p);
  if (!t) return null;
  const r1 = doc.paras[p].cell!.row, r2 = doc.paras[q]?.cell?.table.id === t.def.id ? doc.paras[q].cell!.row : r1;
  const lo = Math.min(r1, r2), hi = Math.max(r1, r2);
  if (lo === 0 && hi === t.def.rows.length - 1) return deleteTable(doc, p);
  const cells = cellParas(doc, t.a, t.b, t.def);
  // A merge that starts in a deleted row restarts in the first remaining row.
  const rows = t.def.rows.map((row, r) => {
    if (r !== hi + 1) return row;
    return { ...row, cells: row.cells.map((c, ci) => {
      if (c.vMerge !== "continue") return c;
      const gx = gridStart(row, ci);
      const above = lo > 0 ? t.def.rows[lo - 1] : undefined;
      const a = above ? above.cells[cellAtGrid(above, gx)] : undefined;
      return a?.vMerge ? c : { ...c, vMerge: "restart" as const };
    }) };
  }).filter((_, r) => r < lo || r > hi);
  const all = cells.filter((_, r) => r < lo || r > hi);
  rebuild(doc, t.a, t.b, { ...t.def, rows }, all);
  return firstOfCell(doc, t.a, Math.min(lo, rows.length - 1), 0);
}

/** Spalten löschen: the grid columns covered by p's and q's cells. */
export function deleteColumns(doc: Doc, p: number, q: number): Pos | null {
  const t = tableAt(doc, p);
  if (!t) return null;
  const c1 = doc.paras[p].cell!, c2 = doc.paras[q]?.cell?.table.id === t.def.id ? doc.paras[q].cell! : c1;
  const range = (c: CellRef) => { const g = gridStart(t.def.rows[c.row], c.col); return [g, g + Math.max(1, t.def.rows[c.row].cells[c.col].span) - 1]; };
  const [a1, b1] = range(c1), [a2, b2] = range(c2);
  const lo = Math.min(a1, a2), hi = Math.max(b1, b2);
  if (lo === 0 && hi === t.def.grid.length - 1) return deleteTable(doc, p);
  const cells = cellParas(doc, t.a, t.b, t.def);
  const rows: RowDef[] = t.def.rows.map((row, r) => {
    let x = 0;
    const out: CellDef[] = [], outCells: Para[][] = [];
    row.cells.forEach((c, ci) => {
      const span = Math.max(1, c.span);
      const from = Math.max(x, lo), to = Math.min(x + span - 1, hi);
      const removed = to >= from ? to - from + 1 : 0;
      if (removed < span) { out.push({ ...c, span: span - removed, width: undefined }); outCells.push(cells[r][ci]); }
      x += span;
    });
    if (!out.length) { out.push({ span: 1 }); outCells.push([makePara("")]); }
    cells[r] = outCells;
    return { ...row, cells: out };
  });
  const grid = t.def.grid.filter((_, g) => g < lo || g > hi);
  rebuild(doc, t.a, t.b, { ...t.def, grid, rows }, cells);
  return firstOfCell(doc, t.a, c1.row, Math.min(cellAtGrid(rows[c1.row], Math.min(lo, grid.length - 1)), rows[c1.row].cells.length - 1));
}

/** A block of cells (row and cell indices). */
export interface CellBlock { r1: number; r2: number; c1: number; c2: number }

/** Entf on selected cells: their contents go, the cells stay (Word). */
export function clearCells(doc: Doc, p: number, block: CellBlock): void {
  const t = tableAt(doc, p);
  if (!t) return;
  const cells = cellParas(doc, t.a, t.b, t.def);
  for (let r = block.r1; r <= block.r2; r++) {
    for (let c = block.c1; c <= Math.min(block.c2, cells[r].length - 1); c++) cells[r][c] = [emptyLike(cells[r][c])];
  }
  rebuild(doc, t.a, t.b, t.def, cells);
}

/**
 * Zellen löschen (Word's dialog, recorded 2026-10-15): "left" takes the cells out of their
 * rows (the cells after them move left, the rows get shorter); "up" moves the contents below
 * up into them, the cells left at the bottom are empty.
 */
export function deleteCells(doc: Doc, p: number, block: CellBlock, mode: "left" | "up"): Pos | null {
  const t = tableAt(doc, p);
  if (!t) return null;
  const cells = cellParas(doc, t.a, t.b, t.def);
  let rows = t.def.rows;
  if (mode === "left") {
    rows = rows.map((row, r) => {
      if (r < block.r1 || r > block.r2) return row;
      const hi = Math.min(block.c2, row.cells.length - 1);
      if (block.c1 > hi) return row;
      // A row keeps at least one cell.
      const keep = row.cells.filter((_, c) => c < block.c1 || c > hi);
      cells[r] = keep.length ? cells[r].filter((_, c) => c < block.c1 || c > hi) : [[makePara("")]];
      return { ...row, cells: keep.length ? keep : [{ span: 1 }] };
    });
  } else {
    const n = block.r2 - block.r1 + 1;
    for (let c = block.c1; c <= block.c2; c++) {
      for (let r = block.r1; r < rows.length; r++) {
        if (c >= cells[r].length) continue;
        const from = r + n < rows.length && c < cells[r + n].length ? cells[r + n][c] : null;
        cells[r][c] = from ?? [emptyLike(cells[r][c])];
      }
    }
  }
  rebuild(doc, t.a, t.b, { ...t.def, rows }, cells);
  return firstOfCell(doc, t.a, block.r1, Math.min(block.c1, rows[block.r1].cells.length - 1));
}

/**
 * Zellen einfügen (Word's dialog, recorded 2026-10-15): "right" puts new empty cells before the
 * block in its rows (the rows get longer); "down" puts them in the block's place and moves the
 * contents below down, the table getting as many rows at the bottom as the block is high.
 */
export function insertCells(doc: Doc, p: number, block: CellBlock, mode: "right" | "down"): Pos | null {
  const t = tableAt(doc, p);
  if (!t) return null;
  const cells = cellParas(doc, t.a, t.b, t.def);
  let rows = t.def.rows;
  let grid = t.def.grid;
  if (mode === "right") {
    let most = 0;
    rows = rows.map((row, r) => {
      if (r < block.r1 || r > block.r2) return row;
      const c1 = Math.min(block.c1, row.cells.length);
      const hi = Math.min(block.c2, row.cells.length - 1);
      const n = Math.max(1, hi - c1 + 1);
      const fresh = Array.from({ length: n }, (_, i): CellDef => ({ ...row.cells[Math.min(c1 + i, row.cells.length - 1)], span: 1, vMerge: undefined }));
      cells[r] = [...cells[r].slice(0, c1), ...fresh.map(() => [emptyLike(cells[r][c1] ?? [])]), ...cells[r].slice(c1)];
      const out = { ...row, cells: [...row.cells.slice(0, c1), ...fresh, ...row.cells.slice(c1)] };
      most = Math.max(most, out.cells.reduce((s, c) => s + Math.max(1, c.span), 0));
      return out;
    });
    // Longer rows reach past the grid: it gets columns as wide as the new cells (Word: the rows stand out).
    if (most > grid.length) {
      const w = grid[Math.min(block.c1, grid.length - 1)] ?? 36;
      grid = [...grid, ...Array.from({ length: most - grid.length }, () => w)];
    }
  } else {
    const n = block.r2 - block.r1 + 1;
    const last = rows[rows.length - 1];
    const added = Array.from({ length: n }, (): RowDef => ({ ...last, header: undefined, cells: last.cells.map((c) => ({ ...c, vMerge: undefined })) }));
    rows = [...rows, ...added];
    for (let k = 0; k < n; k++) cells.push(last.cells.map((_, c) => [emptyLike(cells[cells.length - 1 - k]?.[c] ?? [])]));
    for (let c = block.c1; c <= block.c2; c++) {
      for (let r = rows.length - 1; r >= block.r1; r--) {
        if (c >= cells[r].length) continue;
        const from = r - n >= block.r1 && c < cells[r - n].length ? cells[r - n][c] : null;
        cells[r][c] = from ?? [emptyLike(cells[r][c])];
      }
    }
  }
  rebuild(doc, t.a, t.b, { ...t.def, grid, rows }, cells);
  return firstOfCell(doc, t.a, block.r1, Math.min(block.c1, rows[block.r1].cells.length - 1));
}

/**
 * Tabelle teilen (Word, recorded): the rows from p's row on become a table of their own, with an
 * empty "Standard" paragraph between; in the first row the empty paragraph goes before the table.
 * The caret goes into the empty paragraph.
 */
export function splitTable(doc: Doc, p: number): Pos | null {
  const t = tableAt(doc, p);
  if (!t) return null;
  const row = doc.paras[p].cell!.row;
  const empty = makePara("");
  if (row === 0) {
    doc.splice(t.a, t.a - 1, [empty]);
    return { p: t.a, o: 0 };
  }
  const cells = cellParas(doc, t.a, t.b, t.def);
  const top: TableDef = { ...t.def, rows: t.def.rows.slice(0, row) };
  const bottom: TableDef = { ...t.def, id: nextTableId(), rows: t.def.rows.slice(row), float: undefined };
  const paras: Para[] = [];
  const add = (def: TableDef, list: CellParas) => list.forEach((cs, r) => cs.forEach((ps, c) => {
    for (const q of ps.length ? ps : [makePara("")]) paras.push({ ...q, cell: { table: def, row: r, col: c } });
  }));
  add(top, cells.slice(0, row));
  const at = t.a + paras.length;
  paras.push(empty);
  add(bottom, cells.slice(row));
  doc.splice(t.a, t.b, paras);
  return { p: at, o: 0 };
}

/**
 * In Text konvertieren (Word's "Tabelle in Text umwandeln", recorded): each row becomes text, its
 * cells separated by `sep` ("\n" makes every cell a paragraph of its own); a cell with several
 * paragraphs keeps them. The caret goes to the start of the text.
 */
export function tableToText(doc: Doc, p: number, sep: string): Pos | null {
  const t = tableAt(doc, p);
  if (!t) return null;
  const cells = cellParas(doc, t.a, t.b, t.def);
  const out: Para[] = [];
  for (const row of cells) {
    let line: Para | null = null;
    for (const ps of row) {
      const list = ps.map((q) => ({ ...q, cell: undefined }));
      if (line === null || sep === "\n") {
        if (line) out.push(line);
        out.push(...list.slice(0, -1));
        line = list[list.length - 1];
        continue;
      }
      const glue = makePara(sep, {}, line.runs[line.runs.length - 1]?.props ?? line.mark, line.mark);
      line = joinParas(joinParas(line, glue), list[0]);
      if (list.length > 1) {
        out.push(line, ...list.slice(1, -1));
        line = list[list.length - 1];
      }
    }
    if (line) out.push(line);
  }
  doc.splice(t.a, t.b, out);
  return { p: t.a, o: 0 };
}

/** Tabelle löschen. */
export function deleteTable(doc: Doc, p: number): Pos | null {
  const t = tableAt(doc, p);
  if (!t) return null;
  doc.splice(t.a, t.b, []);
  return { p: Math.min(t.a, doc.paras.length - 1), o: 0 };
}

/** Zellen verbinden: the cells between p's and q's cell become one. */
export function mergeCells(doc: Doc, p: number, q: number): Pos | null {
  const t = tableAt(doc, p);
  if (!t || doc.paras[q]?.cell?.table.id !== t.def.id) return null;
  const c1 = doc.paras[p].cell!, c2 = doc.paras[q].cell!;
  const r1 = Math.min(c1.row, c2.row), r2 = Math.max(c1.row, c2.row);
  const g1 = gridStart(t.def.rows[c1.row], c1.col), g2 = gridStart(t.def.rows[c2.row], c2.col);
  const lo = Math.min(g1, g2);
  const hiC = g1 >= g2 ? c1 : c2;
  const hi = gridStart(t.def.rows[hiC.row], hiC.col) + Math.max(1, t.def.rows[hiC.row].cells[hiC.col].span) - 1;
  const cells = cellParas(doc, t.a, t.b, t.def);
  const content: Para[] = [];
  const keep = (list: Para[]) => { for (const x of list) if (x.text.length || x.block) content.push(x); };
  const rows = t.def.rows.map((row, r) => {
    if (r < r1 || r > r2) return row;
    let x = 0;
    const out: CellDef[] = [], outCells: Para[][] = [];
    let merged = false;
    row.cells.forEach((c, ci) => {
      const span = Math.max(1, c.span);
      if (x >= lo && x + span - 1 <= hi) {
        keep(cells[r][ci]);
        if (!merged) {
          out.push({ ...c, span: hi - lo + 1, width: undefined, vMerge: r2 > r1 ? (r === r1 ? "restart" : "continue") : c.vMerge });
          outCells.push(r === r1 ? [] : [emptyLike(cells[r][ci])]);
          merged = true;
        }
      } else { out.push(c); outCells.push(cells[r][ci]); }
      x += span;
    });
    cells[r] = outCells;
    return { ...row, cells: out };
  });
  const firstIdx = cellAtGrid(rows[r1], lo);
  cells[r1][firstIdx] = content.length ? content : [emptyLike(cells[r1][firstIdx] ?? [])];
  rebuild(doc, t.a, t.b, { ...t.def, rows }, cells);
  return firstOfCell(doc, t.a, r1, firstIdx);
}

/** Zellen teilen: a cell into n columns (the grid column is divided, other rows span it). */
export function splitCell(doc: Doc, p: number, n: number): Pos | null {
  const t = tableAt(doc, p);
  if (!t || n < 2) return null;
  const c = doc.paras[p].cell!;
  const row0 = t.def.rows[c.row];
  const gx = gridStart(row0, c.col);
  const span = Math.max(1, row0.cells[c.col].span);
  const cells = cellParas(doc, t.a, t.b, t.def);
  let grid = [...t.def.grid];
  let rows = [...t.def.rows];
  if (span >= n) {
    // Enough grid columns already: just divide the spans.
    const parts = Array.from({ length: n }, (_, k) => Math.floor(span / n) + (k < span % n ? 1 : 0));
    rows = rows.map((row, r) => r !== c.row ? row : { ...row, cells: [...row.cells.slice(0, c.col), ...parts.map((s, k) => ({ ...row.cells[c.col], span: s, vMerge: undefined, width: undefined, ...(k ? {} : {}) })), ...row.cells.slice(c.col + 1)] });
    cells[c.row] = [...cells[c.row].slice(0, c.col), cells[c.row][c.col], ...parts.slice(1).map(() => [emptyLike(cells[c.row][c.col])]), ...cells[c.row].slice(c.col + 1)];
  } else {
    // Divide the cell's last grid column into n - span + 1 columns.
    const extra = n - span;
    const g = gx + span - 1;
    const w = grid[g] / (extra + 1);
    grid = [...grid.slice(0, g), ...Array.from({ length: extra + 1 }, () => w), ...grid.slice(g + 1)];
    rows = rows.map((row, r) => {
      if (r === c.row) {
        return { ...row, cells: [...row.cells.slice(0, c.col), ...Array.from({ length: n }, (_, k) => ({ ...row.cells[c.col], span: k < span - 1 ? 1 : 1, vMerge: undefined, width: undefined })), ...row.cells.slice(c.col + 1)] };
      }
      const ci = cellAtGrid(row, g);
      return { ...row, cells: row.cells.map((cc, k) => (k === ci ? { ...cc, span: Math.max(1, cc.span) + extra, width: undefined } : cc)) };
    });
    cells[c.row] = [...cells[c.row].slice(0, c.col), cells[c.row][c.col], ...Array.from({ length: n - 1 }, () => [emptyLike(cells[c.row][c.col])]), ...cells[c.row].slice(c.col + 1)];
  }
  rebuild(doc, t.a, t.b, { ...t.def, grid, rows }, cells);
  return firstOfCell(doc, t.a, c.row, c.col);
}

/** Changes the table definition (style, borders, look, widths …) keeping all content. */
export function updateTable(doc: Doc, p: number, change: (def: TableDef) => TableDef): void {
  const t = tableAt(doc, p);
  if (!t) return;
  rebuild(doc, t.a, t.b, change(t.def), cellParas(doc, t.a, t.b, t.def));
}

/** Changes the definitions of the cells between p's and q's cell (shading, borders, alignment). */
export function updateCells(doc: Doc, p: number, q: number, change: (c: CellDef, row: number, col: number) => CellDef): void {
  const t = tableAt(doc, p);
  if (!t) return;
  const c1 = doc.paras[p].cell!, c2 = doc.paras[q]?.cell?.table.id === t.def.id ? doc.paras[q].cell! : c1;
  const r1 = Math.min(c1.row, c2.row), r2 = Math.max(c1.row, c2.row);
  const g1 = gridStart(t.def.rows[c1.row], c1.col), g2 = gridStart(t.def.rows[c2.row], c2.col);
  const lo = Math.min(g1, g2), hi = Math.max(g1, g2);
  const rows = t.def.rows.map((row, r) => (r < r1 || r > r2 ? row : {
    ...row, cells: row.cells.map((c, ci) => { const gx = gridStart(row, ci); return gx + Math.max(1, c.span) - 1 >= lo && gx <= hi ? change(c, r, ci) : c; }),
  }));
  updateTable(doc, p, (def) => ({ ...def, rows }));
}

/** Changes the rows between p's and q's row (height, header row, may break). */
export function updateRows(doc: Doc, p: number, q: number, change: (r: RowDef) => RowDef): void {
  const t = tableAt(doc, p);
  if (!t) return;
  const r1 = doc.paras[p].cell!.row, r2 = doc.paras[q]?.cell?.table.id === t.def.id ? doc.paras[q].cell!.row : r1;
  const lo = Math.min(r1, r2), hi = Math.max(r1, r2);
  updateTable(doc, p, (def) => ({ ...def, rows: def.rows.map((r, i) => (i >= lo && i <= hi ? change(r) : r)) }));
}

/** Spalten gleichmäßig verteilen. */
export function distributeColumns(doc: Doc, p: number): void {
  updateTable(doc, p, (def) => {
    const total = def.grid.reduce((s, x) => s + x, 0);
    return { ...def, grid: def.grid.map(() => total / def.grid.length), rows: def.rows.map((r) => ({ ...r, cells: r.cells.map((c) => ({ ...c, width: undefined })) })) };
  });
}

/** Moves the boundary after grid column g by dx points (dragging a column border). */
export function resizeColumn(doc: Doc, p: number, g: number, dx: number): void {
  updateTable(doc, p, (def) => {
    const grid = [...def.grid];
    const min = 12;
    if (g + 1 < grid.length) {
      const d = Math.max(min - grid[g], Math.min(grid[g + 1] - min, dx));
      grid[g] += d;
      grid[g + 1] -= d;
    } else grid[g] = Math.max(min, grid[g] + dx);
    return { ...def, grid, rows: def.rows.map((r) => ({ ...r, cells: r.cells.map((c) => ({ ...c, width: undefined })) })) };
  });
}

/** Sortieren in a table: each row's cell texts (a cell's paragraphs joined by spaces). */
export function tableRowTexts(doc: Doc, p: number): string[][] {
  const t = tableAt(doc, p);
  if (!t) return [];
  return cellParas(doc, t.a, t.b, t.def).map((row) => row.map((list) => list.map((q) => q.text).join(" ")));
}

/** Puts rows lo … lo + order.length − 1 in the given order (indices relative to lo). */
export function reorderRows(doc: Doc, p: number, lo: number, order: readonly number[]): void {
  const t = tableAt(doc, p);
  if (!t) return;
  const cells = cellParas(doc, t.a, t.b, t.def);
  const rows = [...t.def.rows], all = [...cells];
  order.forEach((src, i) => { rows[lo + i] = t.def.rows[lo + src]; all[lo + i] = cells[lo + src]; });
  rebuild(doc, t.a, t.b, { ...t.def, rows }, all);
}

/** Position in the first paragraph of cell (row, col) of the table starting at paragraph a. */
export function firstOfCell(doc: Doc, a: number, row: number, col: number): Pos {
  const paras = doc.paras;
  const id = paras[a]?.cell?.table.id;
  for (let i = a; i < paras.length && paras[i].cell?.table.id === id; i++) {
    const c = paras[i].cell!;
    if (c.row === row && c.col === col) return { p: i, o: 0 };
  }
  return { p: a, o: 0 };
}

/** The paragraphs of the cell of paragraph p: first and last index. */
export function cellRange(doc: Doc, p: number): [number, number] {
  const paras = doc.paras;
  const c = paras[p].cell;
  if (!c) return [p, p];
  const same = (i: number) => { const x = paras[i]?.cell; return !!x && x.table.id === c.table.id && x.row === c.row && x.col === c.col; };
  let a = p, b = p;
  while (a > 0 && same(a - 1)) a--;
  while (b < paras.length - 1 && same(b + 1)) b++;
  return [a, b];
}

/** Next (or previous) cell in reading order, skipping merged-away cells; null at the end. */
export function adjacentCell(doc: Doc, p: number, dir: 1 | -1): [number, number] | null {
  const paras = doc.paras;
  const [a, b] = cellRange(doc, p);
  const id = paras[p].cell!.table.id;
  let i = dir > 0 ? b + 1 : a - 1;
  while (i >= 0 && i < paras.length && paras[i].cell?.table.id === id) {
    const c = paras[i].cell!;
    if (c.table.rows[c.row].cells[c.col].vMerge !== "continue") return cellRange(doc, i);
    i = dir > 0 ? cellRange(doc, i)[1] + 1 : cellRange(doc, i)[0] - 1;
  }
  return null;
}

// ----- sections -------------------------------------------------------------------------------

/** Index of the section paragraph p of the main text belongs to. */
export function sectionIndex(doc: Doc, p: number): number {
  let k = 0;
  const main = doc.main;
  for (let i = 0; i < Math.min(p, main.length - 1); i++) if (main[i].sect) k++;
  return k;
}

/** Number of sections. */
export function sectionCount(doc: Doc): number {
  return doc.main.slice(0, -1).filter((p) => p.sect).length + 1;
}

/**
 * Abschnittsumbruch at the caret: the text before it ends a section with the current
 * section's properties; the new section starts as `start` with the same properties.
 */
export function insertSectionBreak(doc: Doc, at: Pos, start: Section["start"], mark: CharProps): Pos {
  if (doc.activeStory !== "main") return at;
  const cur = doc.sectionOf(at.p);
  doc.insertText(at, at, "\n", mark);
  const paras = doc.paras;
  // The paragraph that now ends before the break carries the section that ends there; the
  // following section (whose properties are stored at its own end) starts with `start`.
  const next = [...paras];
  next[at.p] = { ...paras[at.p], sect: { ...cur } };
  doc.paras = next;
  setSectionStart(doc, at.p + 1, start);
  return { p: at.p + 1, o: 0 };
}

/** Sets how the section of paragraph p starts. */
function setSectionStart(doc: Doc, p: number, start: Section["start"]): void {
  const main = doc.main;
  for (let i = p; i < main.length - 1; i++) {
    if (main[i].sect) {
      const next = [...main];
      next[i] = { ...main[i], sect: { ...main[i].sect!, start } };
      doc.setStory("main", next);
      return;
    }
  }
  doc.setSettings({ ...doc.settings, section: { ...doc.settings.section, start } });
}

// ----- fields and objects -----------------------------------------------------------------

/** A field object (Seitenzahl, Anzahl der Seiten, Datum …) created in RaabeWrite. */
export function fieldObject(instr: string, display: string): InlineObject {
  return { kind: "field", data: undefined, display, instr };
}

/** An inline picture created in RaabeWrite (size in points). */
export function imageObject(image: ImageData, width: number, height: number): InlineObject {
  return { kind: "image", data: undefined, display: "", width, height, image };
}

/** Inserts objects (with the formatting `props`) at the caret, replacing [from, to). */
export function insertObject(doc: Doc, from: Pos, to: Pos, obj: InlineObject, props: CharProps): Pos {
  const piece: Para = { text: OBJ, runs: [{ len: 1, props: intern(props) }], props: {}, mark: EMPTY, objects: [obj] };
  return doc.replace(from, to, [piece]);
}

/** Numbering used by a paragraph style list or direct numbering, for exports. */
export function numberingOfPara(p: ParaProps, doc: Doc): ReturnType<typeof numberingOf> {
  return numberingOf(p, doc.styles);
}

export type { Numbering };
