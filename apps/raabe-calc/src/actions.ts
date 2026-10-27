// What Excel's commands do to the cells: entering input, clearing, AutoAusfüllen (series as
// Excel continues them), moving and copying, inserting and deleting cells, rows and columns.
// Every function runs inside the caller's undo step (History.step).

import {
  MAX_COLS, MAX_ROWS, adjustFormula, cellKey, shiftFormula, type Cell, type CellStyle, type Rect, type Sheet, type Workbook,
} from "@raabe/sheet-engine";
import { readInput } from "./engine";
import type { History } from "./ops";

export interface Ctx {
  wb: Workbook;
  history: History;
}

const each = (r: Rect, fn: (row: number, col: number) => void) => {
  for (let row = r.r0; row <= r.r1; row++) for (let col = r.c0; col <= r.c1; col++) fn(row, col);
};

/** Typed input into a cell (or every cell of a range with Strg+Eingabe). */
export function enter(ctx: Ctx, sheet: Sheet, targets: Rect, text: string, origin = { row: targets.r0, col: targets.c0 }): void {
  ctx.history.cells(sheet, targets);
  each(targets, (row, col) => {
    const old = sheet.get(row, col);
    const style = ctx.wb.styles.get(old?.s ?? sheet.rows.get(row)?.s ?? sheet.cols.get(col)?.s);
    if (text === "") {
      sheet.set(row, col, old?.s !== undefined ? { v: null, s: old.s } : undefined);
      return;
    }
    const r = readInput(text, style);
    const cell: Cell = { v: r.v, s: old?.s ?? sheet.rows.get(row)?.s ?? sheet.cols.get(col)?.s };
    // A line break typed with Alt+Eingabe turns on Textumbruch (Excel).
    if (typeof r.v === "string" && r.v.includes(String.fromCharCode(10)) && !style.wrap) cell.s = ctx.wb.styles.derive(cell.s, (s) => { s.wrap = true; });
    if (r.f !== undefined) cell.f = row === origin.row && col === origin.col ? r.f : shiftFormula(r.f, row - origin.row, col - origin.col);
    if (r.numFmt && (style.numFmt === "General" || isDateLike(r.numFmt) !== isDateLike(style.numFmt) || style.numFmt === "0%" && r.numFmt !== "0%")) {
      cell.s = ctx.wb.styles.derive(cell.s, (s) => { s.numFmt = r.numFmt!; });
    }
    sheet.set(row, col, cell);
  });
}

function isDateLike(code: string): boolean {
  return /[dy]|h|s|m{3,}/i.test(code.replace(/"[^"]*"|\[[^\]]*\]/g, ""));
}

/** Inhalte löschen (Entf): values and formulas go, formats stay. */
export function clearContents(ctx: Ctx, sheet: Sheet, ranges: Rect[]): void {
  for (const r of ranges) {
    ctx.history.cells(sheet, r);
    for (const [k, c] of [...sheet.cells]) {
      const row = Math.floor(k / MAX_COLS), col = k % MAX_COLS;
      if (row < r.r0 || row > r.r1 || col < r.c0 || col > r.c1) continue;
      sheet.set(row, col, c.s !== undefined || c.note ? { v: null, s: c.s, note: c.note } : undefined);
    }
  }
}

/** Alle löschen / Formate löschen. */
export function clearAll(ctx: Ctx, sheet: Sheet, ranges: Rect[], what: "all" | "formats" | "notes" | "links"): void {
  for (const r of ranges) {
    ctx.history.cells(sheet, r);
    for (const [k, c] of [...sheet.cells]) {
      const row = Math.floor(k / MAX_COLS), col = k % MAX_COLS;
      if (row < r.r0 || row > r.r1 || col < r.c0 || col > r.c1) continue;
      if (what === "all") sheet.set(row, col, undefined);
      else if (what === "formats") sheet.set(row, col, { ...c, s: undefined });
      else if (what === "notes") sheet.set(row, col, { ...c, note: undefined });
      else sheet.set(row, col, { ...c, link: undefined });
    }
    if (what === "all" || what === "formats") {
      ctx.history.merges(sheet);
      sheet.merges = sheet.merges.filter((m) => m.r1 < r.r0 || m.r0 > r.r1 || m.c1 < r.c0 || m.c0 > r.c1);
    }
  }
}

/** Changes the format of every cell in the ranges (whole columns/rows: their own format too). */
export function format(ctx: Ctx, sheet: Sheet, ranges: Rect[], change: (s: CellStyle) => void): void {
  const ids = new Map<number | undefined, number>();
  const derive = (id: number | undefined) => {
    let n = ids.get(id);
    if (n === undefined) { n = ctx.wb.styles.derive(id, change); ids.set(id, n); }
    return n;
  };
  for (const r of ranges) {
    const wholeCols = r.r0 === 0 && r.r1 === MAX_ROWS - 1;
    const wholeRows = r.c0 === 0 && r.c1 === MAX_COLS - 1;
    if (wholeCols || wholeRows) {
      ctx.history.cells(sheet, r);
      if (wholeCols) for (let c = r.c0; c <= r.c1; c++) { ctx.history.col(sheet, c); const info = sheet.cols.get(c) ?? {}; sheet.cols.set(c, { ...info, s: derive(info.s) }); }
      if (wholeRows && !wholeCols) for (let rr = r.r0; rr <= r.r1; rr++) { ctx.history.row(sheet, rr); const info = sheet.rows.get(rr) ?? {}; sheet.rows.set(rr, { ...info, s: derive(info.s), custom: info.custom }); }
      for (const [k, c] of [...sheet.cells]) {
        const row = Math.floor(k / MAX_COLS), col = k % MAX_COLS;
        if (row >= r.r0 && row <= r.r1 && col >= r.c0 && col <= r.c1) sheet.set(row, col, { ...c, s: derive(c.s) });
      }
      continue;
    }
    ctx.history.cells(sheet, r);
    each(r, (row, col) => {
      const c = sheet.get(row, col);
      const base = c?.s ?? sheet.rows.get(row)?.s ?? sheet.cols.get(col)?.s;
      sheet.set(row, col, { ...(c ?? { v: null }), s: derive(base) });
    });
  }
}

// ----- AutoAusfüllen ------------------------------------------------------------------------

const LISTS: string[][] = [
  ["Mo", "Di", "Mi", "Do", "Fr", "Sa", "So"],
  ["Montag", "Dienstag", "Mittwoch", "Donnerstag", "Freitag", "Samstag", "Sonntag"],
  ["Jan", "Feb", "Mrz", "Apr", "Mai", "Jun", "Jul", "Aug", "Sep", "Okt", "Nov", "Dez"],
  ["Januar", "Februar", "März", "April", "Mai", "Juni", "Juli", "August", "September", "Oktober", "November", "Dezember"],
];

function listOf(v: string): { list: string[]; i: number } | null {
  const t = v.toLowerCase();
  for (const list of LISTS) {
    const i = list.findIndex((x) => x.toLowerCase() === t);
    if (i >= 0) return { list, i };
  }
  return null;
}

function matchCase(model: string, word: string): string {
  if (model === model.toUpperCase() && model !== model.toLowerCase()) return word.toUpperCase();
  if (model === model.toLowerCase()) return word.toLowerCase();
  return word;
}

/**
 * Fills `target` (which contains `source` and extends it in one direction) as Excel's fill
 * handle does: numbers and dates continue their trend (one number alone is copied, Strg makes
 * it a series), text with a number at its end counts on, lists (Mo, Januar …) continue,
 * formulas move their relative references.
 */
export function autoFill(ctx: Ctx, sheet: Sheet, source: Rect, target: Rect, opts: { series?: boolean; copy?: boolean; formatsOnly?: boolean; valuesOnly?: boolean } = {}): void {
  ctx.history.cells(sheet, target);
  const down = target.r1 > source.r1, up = target.r0 < source.r0, right = target.c1 > source.c1, left = target.c0 < source.c0;
  const vertical = down || up;
  const lanes = vertical ? range(source.c0, source.c1) : range(source.r0, source.r1);
  for (const lane of lanes) {
    const srcCells = vertical ? range(source.r0, source.r1).map((r) => ({ row: r, col: lane })) : range(source.c0, source.c1).map((c) => ({ row: lane, col: c }));
    const n = srcCells.length;
    const values = srcCells.map((p) => sheet.get(p.row, p.col));
    const dests: { row: number; col: number; k: number }[] = [];
    if (down) for (let r = source.r1 + 1, k = n; r <= target.r1; r++, k++) dests.push({ row: r, col: lane, k });
    if (up) for (let r = source.r0 - 1, k = -1; r >= target.r0; r--, k--) dests.push({ row: r, col: lane, k });
    if (right) for (let c = source.c1 + 1, k = n; c <= target.c1; c++, k++) dests.push({ row: lane, col: c, k });
    if (left) for (let c = source.c0 - 1, k = -1; c >= target.c0; c--, k--) dests.push({ row: lane, col: c, k });
    const plan = seriesPlan(values, opts);
    for (const d of dests) {
      const srcIdx = ((d.k % n) + n) % n;
      const src = values[srcIdx];
      const sp = srcCells[srcIdx];
      const style = src?.s;
      if (opts.formatsOnly) {
        const old = sheet.get(d.row, d.col);
        sheet.set(d.row, d.col, { ...(old ?? { v: null }), s: style });
        continue;
      }
      let cell: Cell | undefined;
      if (src?.f !== undefined && !opts.valuesOnly) cell = { v: null, f: shiftFormula(src.f, d.row - sp.row, d.col - sp.col), s: style };
      else cell = plan(d.k, srcIdx, src);
      if (opts.valuesOnly) cell = { ...(cell ?? { v: null }), s: sheet.get(d.row, d.col)?.s };
      sheet.set(d.row, d.col, cell);
    }
  }
}

function range(a: number, b: number): number[] {
  const out: number[] = [];
  for (let i = a; i <= b; i++) out.push(i);
  return out;
}

/** How the k-th cell of a filled lane gets its value (k counts from the source's first cell). */
function seriesPlan(values: (Cell | undefined)[], opts: { series?: boolean; copy?: boolean }): (k: number, srcIdx: number, src: Cell | undefined) => Cell | undefined {
  const copy = (_k: number, _i: number, src: Cell | undefined) => (src ? { ...src } : undefined);
  if (opts.copy) return copy;
  const nums = values.map((c) => (c && c.f === undefined && typeof c.v === "number" ? c.v : null));
  const n = values.length;
  // All numbers: a linear trend (least squares, as Excel's fill handle does); one number: copy, or +1 with Strg.
  if (nums.every((v) => v !== null)) {
    const ys = nums as number[];
    if (n === 1) {
      if (!opts.series) return copy;
      return (k, _i, src) => ({ ...src!, v: ys[0] + k });
    }
    const xm = (n - 1) / 2, ym = ys.reduce((a, b) => a + b, 0) / n;
    let sxy = 0, sxx = 0;
    ys.forEach((y, x) => { sxy += (x - xm) * (y - ym); sxx += (x - xm) ** 2; });
    const slope = sxx ? sxy / sxx : 0;
    const icept = ym - slope * xm;
    return (k, i, src) => ({ ...(src ?? values[i])!, v: roundSeries(icept + slope * k) });
  }
  // Text: lists (Mo, Januar …) and a number at the end ("Klasse 5a"? no: "Woche 1") continue; mixed lanes per cell.
  return (k, i, src) => {
    if (!src || src.f !== undefined || typeof src.v !== "string") {
      if (src && typeof src.v === "number" && src.f === undefined && opts.series) return { ...src, v: src.v + Math.floor(k / n) * 1 };
      return src ? { ...src } : undefined;
    }
    const t = src.v;
    const l = listOf(t);
    const cycle = Math.floor((k - i) / n);
    if (l) {
      // A list in the source with a step: Mo, Mi → Fr.
      const firstList = values[0] && typeof values[0].v === "string" ? listOf(values[0].v) : null;
      const secondList = n > 1 && values[1] && typeof values[1].v === "string" ? listOf(values[1].v) : null;
      const step = firstList && secondList && firstList.list === secondList.list ? (secondList.i - firstList.i + firstList.list.length) % firstList.list.length || 1 : 1;
      const idx = (l.i + cycle * n * (n > 1 ? step : 1) + (n === 1 ? k * step : 0) - (n === 1 ? 0 : 0)) % l.list.length;
      const j = n === 1 ? (l.i + k) % l.list.length : (idx + l.list.length) % l.list.length;
      return { ...src, v: matchCase(t, l.list[(j + l.list.length) % l.list.length]) };
    }
    const m = /^(.*?)(\d+)(\D*)$/.exec(t);
    if (m) {
      const num = Number(m[2]);
      const nums2 = values.map((c) => (c && typeof c.v === "string" ? /^(.*?)(\d+)(\D*)$/.exec(c.v) : null));
      const step = n > 1 && nums2.every((x) => x && x[1] === m[1] && x[3] === m[3]) ? Number(nums2[1]![2]) - Number(nums2[0]![2]) : 1;
      const delta = n > 1 && step !== 0 ? (k - i) * step : (n === 1 ? k : cycle * n) * (step || 1);
      const v = Math.max(0, num + delta);
      return { ...src, v: m[1] + String(v).padStart(m[2].startsWith("0") ? m[2].length : 0, "0") + m[3] };
    }
    return { ...src };
  };
}

function roundSeries(v: number): number {
  return Number(v.toPrecision(15));
}

// ----- moving, copying, inserting ------------------------------------------------------------

/** Copies cells (Strg+C / Strg+V inside RaabeCalc): values, formulas shifted, formats. */
export function copyCells(ctx: Ctx, from: { sheet: Sheet; rect: Rect }, to: Sheet, at: { row: number; col: number }, what: "all" | "values" | "formats" | "formulas" = "all", transpose = false): Rect {
  const h = from.rect.r1 - from.rect.r0 + 1, w = from.rect.c1 - from.rect.c0 + 1;
  const dst = transpose ? { r0: at.row, c0: at.col, r1: at.row + w - 1, c1: at.col + h - 1 } : { r0: at.row, c0: at.col, r1: at.row + h - 1, c1: at.col + w - 1 };
  const src: { row: number; col: number; cell: Cell | undefined }[] = [];
  for (let r = from.rect.r0; r <= from.rect.r1; r++) for (let c = from.rect.c0; c <= from.rect.c1; c++) src.push({ row: r, col: c, cell: from.sheet.get(r, c) });
  ctx.history.cells(to, dst);
  for (const s of src) {
    const dr = transpose ? at.row + (s.col - from.rect.c0) : at.row + (s.row - from.rect.r0);
    const dc = transpose ? at.col + (s.row - from.rect.r0) : at.col + (s.col - from.rect.c0);
    if (dr >= MAX_ROWS || dc >= MAX_COLS) continue;
    const old = to.get(dr, dc);
    const c = s.cell;
    let cell: Cell | undefined;
    if (what === "formats") cell = { ...(old ?? { v: null }), s: c?.s };
    else if (what === "values") cell = c ? { v: c.v, s: old?.s } : old?.s !== undefined ? { v: null, s: old.s } : undefined;
    else if (what === "formulas") cell = c ? { v: c.v, f: c.f !== undefined ? shiftFormula(c.f, dr - s.row, dc - s.col) : undefined, s: old?.s } : undefined;
    else cell = c ? { ...c, f: c.f !== undefined ? shiftFormula(c.f, dr - s.row, dc - s.col) : undefined } : undefined;
    if (cell && cell.f === undefined) delete cell.f;
    to.set(dr, dc, cell);
  }
  if (what === "all" || what === "formats") {
    const ms = from.sheet.merges.filter((m) => m.r0 >= from.rect.r0 && m.r1 <= from.rect.r1 && m.c0 >= from.rect.c0 && m.c1 <= from.rect.c1);
    if (ms.length || to.merges.some((m) => overlaps(m, dst))) {
      ctx.history.merges(to);
      to.merges = to.merges.filter((m) => !overlaps(m, dst));
      for (const m of ms) to.merges.push({ r0: m.r0 - from.rect.r0 + dst.r0, c0: m.c0 - from.rect.c0 + dst.c0, r1: m.r1 - from.rect.r0 + dst.r0, c1: m.c1 - from.rect.c0 + dst.c0 });
    }
  }
  return dst;
}

export function overlaps(a: Rect, b: Rect): boolean {
  return a.r0 <= b.r1 && a.r1 >= b.r0 && a.c0 <= b.c1 && a.c1 >= b.c0;
}

/**
 * Moves cells (dragging the selection's border, Ausschneiden + Einfügen): the cells keep their
 * formulas as they are, and formulas anywhere that pointed to the moved cells follow them.
 */
export function moveCells(ctx: Ctx, sheet: Sheet, from: Rect, at: { row: number; col: number }): Rect {
  const dr = at.row - from.r0, dc = at.col - from.c0;
  const dst = { r0: from.r0 + dr, c0: from.c0 + dc, r1: from.r1 + dr, c1: from.c1 + dc };
  ctx.history.cells(sheet, from);
  ctx.history.cells(sheet, dst);
  const moved: { row: number; col: number; cell: Cell | undefined }[] = [];
  each(from, (row, col) => moved.push({ row, col, cell: sheet.get(row, col) }));
  each(from, (row, col) => sheet.set(row, col, undefined));
  for (const m of moved) sheet.set(m.row + dr, m.col + dc, m.cell);
  // References to the moved block follow it (Excel); done by the engine on every formula.
  for (const s of ctx.wb.sheets) {
    for (const [k, c] of s.cells) {
      if (c.f === undefined) continue;
      const nf = adjustFormula(c.f, { sheet: sheet.name, kind: "move", from, dRow: dr, dCol: dc, home: s.name });
      if (nf !== c.f) {
        const row = Math.floor(k / MAX_COLS), col = k % MAX_COLS;
        ctx.history.cell(s, row, col);
        s.cells.set(k, { ...c, f: nf });
      }
    }
  }
  if (sheet.merges.some((m) => overlaps(m, from) || overlaps(m, dst))) {
    ctx.history.merges(sheet);
    sheet.merges = sheet.merges.filter((m) => !overlaps(m, dst) || overlaps(m, from)).map((m) =>
      m.r0 >= from.r0 && m.r1 <= from.r1 && m.c0 >= from.c0 && m.c1 <= from.c1 ? { r0: m.r0 + dr, c0: m.c0 + dc, r1: m.r1 + dr, c1: m.c1 + dc } : m);
  }
  return dst;
}

/** Inserts (count > 0) or deletes (count < 0) whole rows or columns at `at`. */
export function insertDelete(ctx: Ctx, sheet: Sheet, kind: "rows" | "cols", at: number, count: number): void {
  ctx.history.whole(sheet);
  const isRows = kind === "rows";
  const next = new Map<number, Cell>();
  for (const [k, c] of sheet.cells) {
    const row = Math.floor(k / MAX_COLS), col = k % MAX_COLS;
    const p = isRows ? row : col;
    if (count < 0 && p >= at && p < at - count) continue;
    const np = p >= at ? p + count : p;
    if (np < 0 || np >= (isRows ? MAX_ROWS : MAX_COLS)) continue;
    next.set(isRows ? cellKey(np, col) : cellKey(row, np), c);
  }
  // Inserted rows/columns take the format of the one before them (Excel: Format von oben/links).
  if (count > 0 && at > 0) {
    for (const [k, c] of sheet.cells) {
      const row = Math.floor(k / MAX_COLS), col = k % MAX_COLS;
      if ((isRows ? row : col) !== at - 1 || c.s === undefined) continue;
      for (let i = 0; i < count; i++) next.set(isRows ? cellKey(at + i, col) : cellKey(row, at + i), { v: null, s: c.s });
    }
  }
  sheet.cells = next;
  const shiftInfo = <T>(m: Map<number, T>) => {
    const out = new Map<number, T>();
    for (const [i, v] of m) {
      if (count < 0 && i >= at && i < at - count) continue;
      out.set(i >= at ? i + count : i, v);
    }
    if (count > 0 && at > 0 && m.has(at - 1)) for (let i = 0; i < count; i++) out.set(at + i, { ...m.get(at - 1)! });
    return out;
  };
  if (isRows) sheet.rows = shiftInfo(sheet.rows);
  else sheet.cols = shiftInfo(sheet.cols);
  sheet.merges = sheet.merges.flatMap((m) => {
    const a0 = isRows ? m.r0 : m.c0, a1 = isRows ? m.r1 : m.c1;
    let n0 = a0, n1 = a1;
    if (count > 0) { if (a0 >= at) { n0 += count; n1 += count; } else if (a1 >= at) n1 += count; }
    else {
      const d0 = at, d1 = at - count - 1;
      const cut = (p: number) => (p < d0 ? p : p > d1 ? p + count : d0);
      n0 = cut(a0); n1 = a1 > d1 ? a1 + count : Math.min(a1, d0 - 1);
      if (a0 >= d0 && a1 <= d1) return [];
      if (n1 < n0) return [];
    }
    const r = isRows ? { ...m, r0: n0, r1: n1 } : { ...m, c0: n0, c1: n1 };
    return r.r0 === r.r1 && r.c0 === r.c1 ? [] : [r];
  });
  // Formulas on every sheet follow (references grow, shrink or become #BEZUG!).
  for (const s of ctx.wb.sheets) {
    if (s !== sheet) ctx.history.whole(s);
    for (const [k, c] of s.cells) {
      if (c.f === undefined) continue;
      const nf = adjustFormula(c.f, { sheet: sheet.name, kind, at, count, home: s.name });
      if (nf !== c.f) s.cells.set(k, { ...c, f: nf });
    }
  }
}

/** Zellen einfügen/löschen with shifting (rechts/unten verschieben, nach links/oben verschieben). */
export function shiftCells(ctx: Ctx, sheet: Sheet, r: Rect, dir: "right" | "down" | "left" | "up"): void {
  ctx.history.whole(sheet);
  const next = new Map<number, Cell>();
  const h = r.r1 - r.r0 + 1, w = r.c1 - r.c0 + 1;
  for (const [k, c] of sheet.cells) {
    const row = Math.floor(k / MAX_COLS), col = k % MAX_COLS;
    let nr = row, nc = col;
    if (dir === "right" && row >= r.r0 && row <= r.r1 && col >= r.c0) nc = col + w;
    if (dir === "down" && col >= r.c0 && col <= r.c1 && row >= r.r0) nr = row + h;
    if (dir === "left" && row >= r.r0 && row <= r.r1 && col >= r.c0) { if (col <= r.c1) continue; nc = col - w; }
    if (dir === "up" && col >= r.c0 && col <= r.c1 && row >= r.r0) { if (row <= r.r1) continue; nr = row - h; }
    if (nr < MAX_ROWS && nc < MAX_COLS) next.set(cellKey(nr, nc), c);
  }
  sheet.cells = next;
}
