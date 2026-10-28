// What Excel's commands do to the cells: entering input, clearing, AutoAusfüllen (series as
// Excel continues them), moving and copying, inserting and deleting cells, rows and columns.
// Every function runs inside the caller's undo step (History.step).

import {
  MAX_COLS, MAX_ROWS, colName, adjustFormula, affectedByShift, shiftValidations, cellKey, formulaRefs, shiftFormula, toLocalFormula, type Cell, type CellStyle, type Rect, type Sheet, type Validation, type Workbook,
} from "@raabe/sheet-engine";
import { readInput } from "./engine";
import type { History } from "./ops";
import { applyValidation } from "./validation";

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
    // A formula in a Standard cell shows its result in the format of a referenced cell (the first
    // one that has one), or as a date for HEUTE/JETZT/DATUM, as Excel does.
    if (cell.f !== undefined && style.numFmt === "General") {
      const fmt = formulaFormat(ctx.wb, sheet, cell.f);
      if (fmt) cell.s = ctx.wb.styles.derive(cell.s, (s) => { s.numFmt = fmt; });
    }
    if (r.numFmt && (style.numFmt === "General" || isDateLike(r.numFmt) !== isDateLike(style.numFmt) || style.numFmt === "0%" && r.numFmt !== "0%")) {
      cell.s = ctx.wb.styles.derive(cell.s, (s) => { s.numFmt = r.numFmt!; });
    }
    sheet.set(row, col, cell);
  });
}

function formulaFormat(wb: import("@raabe/sheet-engine").Workbook, sheet: Sheet, file: string): string | undefined {
  const upper = file.toUpperCase();
  if (/^(TODAY|DATE|EDATE|EOMONTH|WORKDAY)\(/.test(upper)) return "m/d/yyyy";
  if (/^NOW\(/.test(upper)) return "m/d/yyyy h:mm";
  try {
    for (const ref of formulaRefs(toLocalFormula(file))) {
      if (ref.sheet) continue;
      for (let row = ref.ref.r0; row <= Math.min(ref.ref.r1, ref.ref.r0 + 200); row++) for (let col = ref.ref.c0; col <= Math.min(ref.ref.c1, ref.ref.c0 + 50); col++) {
        const c = sheet.get(row, col);
        if (!c) continue;
        const fmt = wb.styles.get(c.s).numFmt;
        if (fmt !== "General" && fmt !== "@") return fmt;
      }
    }
  } catch { /* no references */ }
  return undefined;
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
  if (!opts.formatsOnly && !opts.valuesOnly && sheet.validations?.length) fillValidations(ctx, sheet, source, target);
}

/** AutoAusfüllen repeats the source's Datenüberprüfung over the filled cells, block by block (Excel). */
function fillValidations(ctx: Ctx, sheet: Sheet, source: Rect, target: Rect): void {
  const vertical = target.r0 !== source.r0 || target.r1 !== source.r1;
  const n = vertical ? source.r1 - source.r0 + 1 : source.c1 - source.c0 + 1;
  const lo = vertical ? target.r0 : target.c0, hi = vertical ? target.r1 : target.c1, s0 = vertical ? source.r0 : source.c0;
  for (let a = s0 - Math.ceil((s0 - lo) / n) * n; a <= hi; a += n) {
    if (a === s0) continue;
    const t0 = Math.max(a, lo), t1 = Math.min(a + n - 1, hi);
    const src = vertical ? { ...source, r0: s0 + (t0 - a), r1: s0 + (t1 - a) } : { ...source, c0: s0 + (t0 - a), c1: s0 + (t1 - a) };
    const dst = vertical ? { ...source, r0: t0, r1: t1 } : { ...source, c0: t0, c1: t1 };
    pasteValidations(ctx, { sheet, rect: src }, sheet, dst);
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
  if (what === "all" && !transpose) pasteValidations(ctx, from, to, dst);
  return dst;
}

/**
 * Einfügen takes the copied cells' Datenüberprüfung along and replaces the target's (Excel);
 * relative references in criteria move with the cells.
 */
function pasteValidations(ctx: Ctx, from: { sheet: Sheet; rect: Rect }, to: Sheet, dst: Rect): void {
  const dr = dst.r0 - from.rect.r0, dc = dst.c0 - from.rect.c0;
  const moved: Validation[] = [];
  for (const v of from.sheet.validations ?? []) {
    const parts = v.ref.map((r) => ({ r0: Math.max(r.r0, from.rect.r0), c0: Math.max(r.c0, from.rect.c0), r1: Math.min(r.r1, from.rect.r1), c1: Math.min(r.c1, from.rect.c1) }))
      .filter((r) => r.r0 <= r.r1 && r.c0 <= r.c1).map((r) => ({ r0: r.r0 + dr, c0: r.c0 + dc, r1: r.r1 + dr, c1: r.c1 + dc }));
    if (!parts.length) continue;
    // Criteria are relative to the first range's top-left cell.
    const o = v.ref[0], n = parts[0];
    const sr = n.r0 - o.r0, sc = n.c0 - o.c0;
    const shift = (f?: string) => (f !== undefined && (sr || sc) ? shiftFormula(f, sr, sc) : f);
    moved.push({ ...v, ref: parts, formula1: shift(v.formula1), formula2: shift(v.formula2) });
  }
  if (!moved.length && !(to.validations ?? []).some((v) => v.ref.some((r) => overlaps(r, dst)))) return;
  ctx.history.validations(to);
  applyValidation(to, [dst], null);
  if (moved.length) to.validations = [...(to.validations ?? []), ...moved];
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
/**
 * Kommentare and pictures sit at a row and column of their own: they move with inserted and
 * deleted rows or columns as the cells do (a thread on a deleted cell goes with it; a picture
 * whose top-left cell is deleted moves to the first row or column after the gap), as in Excel.
 */
function shiftAnchored(sheet: Sheet, kind: "rows" | "cols", at: number, count: number): void {
  const rows = kind === "rows";
  const move = (p: number) => (p < at ? p : count < 0 && p < at - count ? null : p + count);
  if (sheet.threads?.length) {
    const next = sheet.threads.flatMap((t) => {
      const p = move(rows ? t.row : t.col);
      return p === null ? [] : [{ ...t, row: rows ? p : t.row, col: rows ? t.col : p }];
    });
    sheet.threads = next.length ? next : undefined;
  }
  if (sheet.pictures?.length) {
    const shift = <T extends { row: number; col: number; x: number; y: number }>(a: T): T => {
      const p = move(rows ? a.row : a.col);
      const q = p === null ? at : p;
      return { ...a, row: rows ? q : a.row, col: rows ? a.col : q, x: rows || p !== null ? a.x : 0, y: !rows || p !== null ? a.y : 0 };
    };
    sheet.pictures = sheet.pictures.map((pic) => ({ ...pic, from: shift(pic.from), to: pic.to ? shift(pic.to) : undefined }));
  }
}

export function insertDelete(ctx: Ctx, sheet: Sheet, kind: "rows" | "cols", at: number, count: number): void {
  ctx.history.whole(sheet);
  // Datenüberprüfung ranges move with the cells (also those on other sheets that point here).
  for (const s of affectedByShift(ctx.wb, sheet)) ctx.history.validations(s);
  shiftValidations(ctx.wb, sheet, kind, at, count);
  shiftAnchored(sheet, kind, at, count);
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
  // Charts follow too: their corners on this sheet and their series references anywhere (Excel).
  for (const s of ctx.wb.sheets) {
    if (!s.charts?.length) continue;
    ctx.history.charts(s);
    for (const ch of s.charts) {
      if (s === sheet) {
        for (const corner of [ch.from, ch.to]) {
          const p = isRows ? corner.row : corner.col;
          const np = p >= at ? Math.max(at, p + count) : p;
          if (isRows) corner.row = np; else corner.col = np;
        }
      }
      for (const ser of ch.series) {
        for (const d of [ser.name, ser.cat, ser.val]) {
          if (!d?.ref) continue;
          const nf = adjustFormula(d.ref, { sheet: sheet.name, kind, at, count, home: s.name });
          d.ref = nf;
        }
      }
    }
  }
}

/** Zellen einfügen/löschen with shifting (rechts/unten verschieben, nach links/oben verschieben). */
/**
 * Umschalt+drag of the selection's border: the cells are taken out (the rest closes the gap) and
 * inserted before row or column `at` (the cells there move aside), as Excel's insert-move.
 */
export function moveInsert(ctx: Ctx, sheet: Sheet, src: Rect, at: { row: number; col: number }, rows: boolean): Rect {
  const h = src.r1 - src.r0 + 1, w = src.c1 - src.c0 + 1;
  // Where the gap closed, the cells after it moved up (or left): the place counts from there.
  const overlap = rows ? at.col <= src.c1 && at.col + w - 1 >= src.c0 : at.row <= src.r1 && at.row + h - 1 >= src.r0;
  const target = rows
    ? { r0: overlap && at.row > src.r1 ? at.row - h : at.row, c0: at.col, r1: 0, c1: at.col + w - 1 }
    : { r0: at.row, c0: overlap && at.col > src.c1 ? at.col - w : at.col, r1: at.row + h - 1, c1: 0 };
  if (rows) target.r1 = target.r0 + h - 1;
  else target.c1 = target.c0 + w - 1;
  // Done as Excel's references see it: the cells go aside (references follow them), the gap
  // closes, room opens at the target, and the cells come back into it.
  const aside = { row: Math.max(src.r1, target.r1) + 1 + h, col: MAX_COLS - w };
  ctx.history.whole(sheet);
  moveCells(ctx, sheet, src, aside);
  shiftCells(ctx, sheet, src, rows ? "up" : "left");
  shiftCells(ctx, sheet, target, rows ? "down" : "right");
  moveCells(ctx, sheet, { r0: aside.row, c0: aside.col, r1: aside.row + h - 1, c1: aside.col + w - 1 }, { row: target.r0, col: target.c0 });
  return target;
}

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
  // Kommentare in the shifted block move with their cells; those in deleted cells go.
  if (sheet.threads?.length) {
    const h2 = r.r1 - r.r0 + 1, w2 = r.c1 - r.c0 + 1;
    const next2 = sheet.threads.flatMap((t) => {
      let row = t.row, col = t.col;
      if (dir === "right" && row >= r.r0 && row <= r.r1 && col >= r.c0) col += w2;
      if (dir === "down" && col >= r.c0 && col <= r.c1 && row >= r.r0) row += h2;
      if (dir === "left" && row >= r.r0 && row <= r.r1 && col >= r.c0) { if (col <= r.c1) return []; col -= w2; }
      if (dir === "up" && col >= r.c0 && col <= r.c1 && row >= r.r0) { if (row <= r.r1) return []; row -= h2; }
      return [{ ...t, row, col }];
    });
    sheet.threads = next2.length ? next2 : undefined;
  }
  // Formulas on every sheet follow the shifted cells (references to deleted cells become #BEZUG!).
  for (const s of ctx.wb.sheets) {
    let saved = s === sheet;
    for (const [k, c] of s.cells) {
      if (c.f === undefined) continue;
      const nf = adjustFormula(c.f, { kind: "cells", sheet: sheet.name, rect: r, dir, home: s.name });
      if (nf === c.f) continue;
      if (!saved) { ctx.history.whole(s); saved = true; }
      s.cells.set(k, { ...c, f: nf });
    }
  }
}

/** Excel's day number → a date (serial 1 = 1.1.1900, with Excel's 29.2.1900). */
const toDate = (n: number) => new Date(Date.UTC(1899, 11, 30) + Math.round(n) * 86400000);
const toSerial = (d: Date) => Math.round((d.getTime() - Date.UTC(1899, 11, 30)) / 86400000);

/**
 * Tage / Arbeitstage / Monate / Jahre ausfüllen (right-drag menu): each lane continues from its
 * last date by one unit (two dates: by their distance in that unit), keeping the format.
 */
export function fillDates(ctx: Ctx, sheet: Sheet, source: Rect, target: Rect, unit: "days" | "workdays" | "months" | "years"): void {
  ctx.history.cells(sheet, target);
  const vertical = target.r1 > source.r1 || target.r0 < source.r0;
  const lanes = vertical ? range(source.c0, source.c1) : range(source.r0, source.r1);
  const step = (d: Date, k: number): Date => {
    const x = new Date(d.getTime());
    if (unit === "months" || unit === "years") {
      // As EDATUM: the same day in the month k months on, the month's last day where it is shorter.
      const m = d.getUTCMonth() + (unit === "months" ? k : 12 * k);
      const y = d.getUTCFullYear() + Math.floor(m / 12), mm = ((m % 12) + 12) % 12;
      const last = new Date(Date.UTC(y, mm + 1, 0)).getUTCDate();
      return new Date(Date.UTC(y, mm, Math.min(d.getUTCDate(), last)));
    }
    else if (unit === "days") x.setUTCDate(x.getUTCDate() + k);
    else {
      let left = Math.abs(k);
      const dir = Math.sign(k);
      while (left > 0) { x.setUTCDate(x.getUTCDate() + dir); if (x.getUTCDay() !== 0 && x.getUTCDay() !== 6) left--; }
    }
    return x;
  };
  for (const lane of lanes) {
    const at = (i: number) => (vertical ? { row: i, col: lane } : { row: lane, col: i });
    const first = vertical ? source.r0 : source.c0, last = vertical ? source.r1 : source.c1;
    const start = sheet.get(at(first).row, at(first).col);
    if (typeof start?.v !== "number") continue;
    const end = sheet.get(at(last).row, at(last).col);
    const d0 = toDate(start.v);
    let by = 1;
    if (last > first && typeof end?.v === "number") {
      const d1 = toDate(end.v);
      let total = unit === "months" ? (d1.getUTCFullYear() - d0.getUTCFullYear()) * 12 + d1.getUTCMonth() - d0.getUTCMonth()
        : unit === "years" ? d1.getUTCFullYear() - d0.getUTCFullYear() : Math.round(end.v - start.v);
      if (unit === "workdays") {
        // Workdays between the two dates, not calendar days.
        total = 0;
        const dir = Math.sign(end.v - start.v);
        for (const x = new Date(d0.getTime()); dir && Math.sign(d1.getTime() - x.getTime()) === dir;) {
          x.setUTCDate(x.getUTCDate() + dir);
          if (x.getUTCDay() !== 0 && x.getUTCDay() !== 6) total += dir;
        }
      }
      by = total / (last - first) || 1;
    }
    const lo = vertical ? target.r0 : target.c0, hi = vertical ? target.r1 : target.c1;
    for (let i = lo; i <= hi; i++) {
      if (i >= first && i <= last) continue;
      const p = at(i);
      sheet.set(p.row, p.col, { v: toSerial(step(d0, Math.round((i - first) * by))), s: start.s });
    }
  }
}

/**
 * Linearer / Exponentieller Trend (right-drag menu): the best-fitting straight line (or growth
 * curve) through the source values, continued over the target.
 */
export function fillTrend(ctx: Ctx, sheet: Sheet, source: Rect, target: Rect, exponential: boolean): void {
  ctx.history.cells(sheet, target);
  const vertical = target.r1 > source.r1 || target.r0 < source.r0;
  const lanes = vertical ? range(source.c0, source.c1) : range(source.r0, source.r1);
  for (const lane of lanes) {
    const at = (i: number) => (vertical ? { row: i, col: lane } : { row: lane, col: i });
    const first = vertical ? source.r0 : source.c0, last = vertical ? source.r1 : source.c1;
    const xs: number[] = [], ys: number[] = [];
    for (let i = first; i <= last; i++) {
      const v = sheet.get(at(i).row, at(i).col)?.v;
      if (typeof v === "number" && (!exponential || v > 0)) { xs.push(i); ys.push(exponential ? Math.log(v) : v); }
    }
    if (!xs.length) continue;
    const n = xs.length, mx = xs.reduce((a, b) => a + b, 0) / n, my = ys.reduce((a, b) => a + b, 0) / n;
    let sxy = 0, sxx = 0;
    for (let k = 0; k < n; k++) { sxy += (xs[k] - mx) * (ys[k] - my); sxx += (xs[k] - mx) ** 2; }
    // One value: Excel steps by 1 (linear) or keeps it (growth).
    const slope = n > 1 ? sxy / sxx : exponential ? 0 : 1;
    const style = sheet.get(at(last).row, at(last).col)?.s;
    const lo = vertical ? target.r0 : target.c0, hi = vertical ? target.r1 : target.c1;
    for (let i = lo; i <= hi; i++) {
      if (i >= first && i <= last) continue;
      const y = my + slope * (i - mx);
      sheet.set(at(i).row, at(i).col, { v: exponential ? Math.exp(y) : y, s: style });
    }
  }
}

/** Verknüpfung hier erstellen: formulas pointing to the source cells (one cell: absolute, as Excel's Verknüpfung einfügen). */
export function linkCells(ctx: Ctx, sheet: Sheet, source: Rect, to: { row: number; col: number }): void {
  const h = source.r1 - source.r0, w = source.c1 - source.c0;
  ctx.history.cells(sheet, { r0: to.row, c0: to.col, r1: to.row + h, c1: to.col + w });
  const single = h === 0 && w === 0;
  for (let r = 0; r <= h; r++) for (let c = 0; c <= w; c++) {
    const ref = single ? `$${colName(source.c0)}$${source.r0 + 1}` : `${colName(source.c0 + c)}${source.r0 + r + 1}`;
    enter(ctx, sheet, { r0: to.row + r, c0: to.col + c, r1: to.row + r, c1: to.col + c }, `=${ref}`);
  }
}

/** Link hier erstellen: the source's values with a link to the source cell inside the workbook. */
export function hyperlinkCells(ctx: Ctx, sheet: Sheet, source: Rect, to: { row: number; col: number }): void {
  const h = source.r1 - source.r0, w = source.c1 - source.c0;
  ctx.history.cells(sheet, { r0: to.row, c0: to.col, r1: to.row + h, c1: to.col + w });
  const q = /^[\p{L}_][\p{L}\p{N}_.]*$/u.test(sheet.name) ? sheet.name : `'${sheet.name.replace(/'/g, "''")}'`;
  for (let r = 0; r <= h; r++) for (let c = 0; c <= w; c++) {
    const src = sheet.get(source.r0 + r, source.c0 + c);
    const old = sheet.get(to.row + r, to.col + c);
    sheet.set(to.row + r, to.col + c, { v: src?.v ?? null, s: old?.s, link: `#${q}!${colName(source.c0 + c)}${source.r0 + r + 1}` });
  }
}

/**
 * Teilergebnis's functions: the dialog's name, TEILERGEBNIS's number, and the labels Excel writes
 * after a group's value and in the last row (recorded: probe_excel_subtotal.py).
 */
export const SUBTOTALS: { name: string; code: number; label: string; grand: string }[] = [
  { name: "Summe", code: 9, label: "Ergebnis", grand: "Gesamtergebnis" },
  { name: "Anzahl", code: 3, label: "Anzahl", grand: "Gesamtanzahl" },
  { name: "Mittelwert", code: 1, label: "Mittelwert", grand: "Gesamtmittelwert" },
  { name: "Maximum", code: 4, label: "Max.", grand: "Gesamtmaximum" },
  { name: "Minimum", code: 5, label: "Min.", grand: "Gesamtminimum" },
  { name: "Produkt", code: 6, label: "Produkt", grand: "Gesamtprodukt" },
  { name: "Anzahl Zahlen", code: 2, label: "Anzahl", grand: "Gesamtanzahl" },
  { name: "Standardabweichung (Stichprobe)", code: 7, label: "STABW", grand: "Gesamtstandardabweichung (Stichprobe)" },
  { name: "Standardabweichung (Grundgesamtheit)", code: 8, label: "Standardabweichung (Grundgesamtheit)", grand: "Gesamtstandardabweichung (Grundgesamtheit)" },
  { name: "Varianz (Stichprobe)", code: 10, label: "Var", grand: "Gesamtvarianz" },
  { name: "Varianz (Grundgesamtheit)", code: 11, label: "Varianz (Grundgesamtheit)", grand: "Gesamtvarianz" },
];

/**
 * Teilergebnis as Excel inserts it: after each run of equal values in the group column a row with
 * "<value> <label>" (bold) and TEILERGEBNIS over the run in each chosen column, at the end a
 * row with the total over everything; data rows get outline level 2, result rows 1.
 * `region` includes the header row. Returns the new region.
 */
export function subtotal(ctx: Ctx, sheet: Sheet, region: Rect, groupCol: number, fn: (typeof SUBTOTALS)[number], cols: number[]): Rect {
  ctx.history.whole(sheet);
  const text = (r: number, c: number) => { const v = sheet.get(r, c)?.v; return v === null || v === undefined ? "" : String(v); };
  const first = region.r0 + 1;
  let row = first, end = region.r1;
  const bold = (r: number, c: number) => {
    const cell = sheet.get(r, c) ?? { v: null };
    sheet.set(r, c, { ...cell, s: ctx.wb.styles.derive(cell.s, (st) => { st.font.bold = true; }) });
  };
  const resultRow = (at: number, label: string, from: number, to: number) => {
    insertDelete(ctx, sheet, "rows", at, 1);
    // An inserted row takes the format of the one above (Excel does not here): start clean.
    for (let c = region.c0; c <= region.c1; c++) sheet.set(at, c, undefined);
    enter(ctx, sheet, { r0: at, c0: groupCol, r1: at, c1: groupCol }, label);
    bold(at, groupCol);
    for (const c of cols) {
      const col = colName(c);
      enter(ctx, sheet, { r0: at, c0: c, r1: at, c1: c }, `=TEILERGEBNIS(${fn.code};${col}${from + 1}:${col}${to + 1})`);
    }
    const info = sheet.rows.get(at);
    sheet.rows.set(at, { ...(info ?? {}), outline: label === fn.grand ? undefined : 1 });
  };
  let lastData = first;
  while (row <= end) {
    const value = text(row, groupCol);
    let last = row;
    while (last + 1 <= end && text(last + 1, groupCol) === value) last++;
    for (let r = row; r <= last; r++) sheet.rows.set(r, { ...(sheet.rows.get(r) ?? {}), outline: 2 });
    resultRow(last + 1, `${value} ${fn.label}`, row, last);
    lastData = last;
    end++;
    row = last + 2;
  }
  // The total runs from the first to the last data row (Excel's range leaves out the last result row).
  resultRow(end + 1, fn.grand, first, lastData);
  return { ...region, r1: end + 1 };
}

/** Teilergebnis > Alle entfernen: the result rows (with TEILERGEBNIS) go, and the outline with them. */
export function removeSubtotals(ctx: Ctx, sheet: Sheet, region: Rect): Rect {
  ctx.history.whole(sheet);
  let r1 = region.r1;
  for (let r = region.r1; r > region.r0; r--) {
    let has = false;
    for (let c = region.c0; c <= region.c1 && !has; c++) has = /^SUBTOTAL\(/i.test(sheet.get(r, c)?.f ?? "");
    if (has) { insertDelete(ctx, sheet, "rows", r, -1); r1--; }
  }
  for (let r = region.r0; r <= r1; r++) {
    const info = sheet.rows.get(r);
    if (info?.outline || info?.collapsed || info?.hidden) {
      const { outline: _o, collapsed: _c, hidden: _h, ...rest } = info;
      if (Object.keys(rest).length) sheet.rows.set(r, rest);
      else sheet.rows.delete(r);
    }
  }
  return { ...region, r1 };
}
