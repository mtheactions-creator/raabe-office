// Changes to a workbook as undoable steps. Each step records the cells, columns, rows,
// merges and sheets it touches before changing them; undo puts those back, redo applies the
// recorded "after" state. One user action = one step, as in Excel.

import { cellKey, type Cell, type ColInfo, type Rect, type RowInfo, type Sheet, type Workbook } from "@raabe/sheet-engine";

interface SheetState {
  sheet: Sheet;
  cells: Map<number, Cell | undefined>;
  cols: Map<number, ColInfo | undefined>;
  rows: Map<number, RowInfo | undefined>;
  merges?: Rect[];
  /** The sheet's charts (inserted, moved, sized, formatted, deleted). */
  charts?: Sheet["charts"];
  /** Whole-sheet copies for structural changes (insert/delete rows and columns, sorting). */
  whole?: { cells: Map<number, Cell>; cols: Map<number, ColInfo>; rows: Map<number, RowInfo>; merges: Rect[] };
  defaults?: { col?: number; row: number };
}

interface Step {
  label: string;
  before: SheetState[];
  after: SheetState[];
  sheetsBefore?: { list: Sheet[]; active: number };
  sheetsAfter?: { list: Sheet[]; active: number };
  /** Selection to restore with undo/redo. */
  selBefore?: { sheet: number; sel: Rect[]; active: { row: number; col: number } };
  selAfter?: { sheet: number; sel: Rect[]; active: { row: number; col: number } };
}

export class History {
  private undoStack: Step[] = [];
  private redoStack: Step[] = [];
  private open: Step | null = null;
  private depth = 0;
  /** Called after every change (undo/redo included) with the sheets whose cells changed. */
  onChange: (sheets: Set<Sheet>) => void = () => {};

  constructor(private wb: Workbook) {}

  get canUndo(): boolean { return this.undoStack.length > 0; }
  get canRedo(): boolean { return this.redoStack.length > 0; }
  get undoLabel(): string | undefined { return this.undoStack.at(-1)?.label; }

  /** Runs `fn` as one undo step. Nested calls join the outer step. */
  step<T>(label: string, fn: () => T): T {
    const outer = this.open === null;
    if (outer) this.open = { label, before: [], after: [], selBefore: this.selection() };
    this.depth++;
    try {
      return fn();
    } finally {
      this.depth--;
      if (outer) {
        const st = this.open!;
        this.open = null;
        st.selAfter = this.selection();
        st.after = st.before.map((b) => this.capture(b));
        if (st.sheetsBefore) st.sheetsAfter = { list: [...this.wb.sheets], active: this.wb.active };
        if (st.before.length || st.sheetsBefore) {
          this.undoStack.push(st);
          if (this.undoStack.length > 100) this.undoStack.shift();
          this.redoStack = [];
          this.onChange(new Set(st.before.map((b) => b.sheet)));
        }
      }
    }
  }

  private selection() {
    const s = this.wb.sheet();
    return { sheet: this.wb.active, sel: s.view.selection.map((r) => ({ ...r })), active: { ...s.view.active } };
  }

  private stateFor(sheet: Sheet): SheetState {
    if (!this.open) throw new Error("change outside an undo step");
    let st = this.open.before.find((b) => b.sheet === sheet);
    if (!st) {
      st = { sheet, cells: new Map(), cols: new Map(), rows: new Map(), defaults: { col: sheet.defaultColWidth, row: sheet.defaultRowHeight } };
      this.open.before.push(st);
    }
    return st;
  }

  /** Records cells before they change. */
  cells(sheet: Sheet, r: Rect): void {
    const st = this.stateFor(sheet);
    if (st.whole) return;
    if ((r.r1 - r.r0 + 1) * (r.c1 - r.c0 + 1) > 20000) {
      // Large areas (whole columns, rows): record the cells that exist.
      for (const [k, c] of sheet.cells) {
        const row = Math.floor(k / 16384), col = k % 16384;
        if (row >= r.r0 && row <= r.r1 && col >= r.c0 && col <= r.c1 && !st.cells.has(k)) st.cells.set(k, clone(c));
      }
      st.cells.set(-1 - st.cells.size, undefined);
      (st as SheetState & { areas?: Rect[] }).areas = [...((st as SheetState & { areas?: Rect[] }).areas ?? []), r];
      return;
    }
    for (let row = r.r0; row <= r.r1; row++) {
      for (let col = r.c0; col <= r.c1; col++) {
        const k = cellKey(row, col);
        if (!st.cells.has(k)) st.cells.set(k, clone(sheet.cells.get(k)));
      }
    }
  }

  cell(sheet: Sheet, row: number, col: number): void {
    this.cells(sheet, { r0: row, c0: col, r1: row, c1: col });
  }

  col(sheet: Sheet, c: number): void {
    const st = this.stateFor(sheet);
    if (!st.cols.has(c)) st.cols.set(c, clone(sheet.cols.get(c)));
  }

  row(sheet: Sheet, r: number): void {
    const st = this.stateFor(sheet);
    if (!st.rows.has(r)) st.rows.set(r, clone(sheet.rows.get(r)));
  }

  /** Records the sheet's charts before they change. */
  charts(sheet: Sheet): void {
    const st = this.stateFor(sheet);
    if (!st.charts) st.charts = structuredClone(sheet.charts ?? []);
  }

  merges(sheet: Sheet): void {
    const st = this.stateFor(sheet);
    if (!st.merges) st.merges = sheet.merges.map((m) => ({ ...m }));
  }

  /** Records the whole sheet (structural changes). */
  whole(sheet: Sheet): void {
    const st = this.stateFor(sheet);
    if (!st.whole) st.whole = snapshot(sheet);
  }

  /** Records the list of sheets (insert, delete, move, rename …). */
  sheets(): void {
    if (!this.open) throw new Error("change outside an undo step");
    if (!this.open.sheetsBefore) this.open.sheetsBefore = { list: [...this.wb.sheets], active: this.wb.active };
  }

  private capture(b: SheetState): SheetState {
    const s = b.sheet;
    const out: SheetState = { sheet: s, cells: new Map(), cols: new Map(), rows: new Map(), defaults: { col: s.defaultColWidth, row: s.defaultRowHeight } };
    if (b.whole) out.whole = snapshot(s);
    const areas = (b as SheetState & { areas?: Rect[] }).areas;
    if (areas) {
      (out as SheetState & { areas?: Rect[] }).areas = areas;
      for (const [k, c] of s.cells) {
        const row = Math.floor(k / 16384), col = k % 16384;
        if (areas.some((r) => row >= r.r0 && row <= r.r1 && col >= r.c0 && col <= r.c1)) out.cells.set(k, clone(c));
      }
    }
    for (const k of b.cells.keys()) if (k >= 0) out.cells.set(k, clone(s.cells.get(k)));
    for (const k of b.cols.keys()) out.cols.set(k, clone(s.cols.get(k)));
    for (const k of b.rows.keys()) out.rows.set(k, clone(s.rows.get(k)));
    if (b.merges) out.merges = s.merges.map((m) => ({ ...m }));
    if (b.charts) out.charts = structuredClone(s.charts ?? []);
    return out;
  }

  private apply(states: SheetState[]): void {
    for (const st of states) {
      const s = st.sheet;
      if (st.whole) {
        s.cells = new Map([...st.whole.cells].map(([k, c]) => [k, clone(c)!]));
        s.cols = new Map([...st.whole.cols].map(([k, c]) => [k, { ...c }]));
        s.rows = new Map([...st.whole.rows].map(([k, c]) => [k, { ...c }]));
        s.merges = st.whole.merges.map((m) => ({ ...m }));
      }
      const areas = (st as SheetState & { areas?: Rect[] }).areas;
      if (areas && !st.whole) {
        for (const k of [...s.cells.keys()]) {
          const row = Math.floor(k / 16384), col = k % 16384;
          if (areas.some((r) => row >= r.r0 && row <= r.r1 && col >= r.c0 && col <= r.c1)) s.cells.delete(k);
        }
      }
      for (const [k, c] of st.cells) {
        if (k < 0) continue;
        if (c) s.cells.set(k, clone(c)!);
        else s.cells.delete(k);
      }
      for (const [k, c] of st.cols) { if (c) s.cols.set(k, { ...c }); else s.cols.delete(k); }
      for (const [k, c] of st.rows) { if (c) s.rows.set(k, { ...c }); else s.rows.delete(k); }
      if (st.merges) s.merges = st.merges.map((m) => ({ ...m }));
      if (st.charts) s.charts = structuredClone(st.charts);
      if (st.defaults) { s.defaultColWidth = st.defaults.col; s.defaultRowHeight = st.defaults.row; }
    }
  }

  undo(): boolean {
    const st = this.undoStack.pop();
    if (!st) return false;
    this.apply(st.before);
    if (st.sheetsBefore) { this.wb.sheets = [...st.sheetsBefore.list]; this.wb.active = st.sheetsBefore.active; }
    this.restoreSel(st.selBefore);
    this.redoStack.push(st);
    this.onChange(new Set(st.before.map((b) => b.sheet)));
    return true;
  }

  redo(): boolean {
    const st = this.redoStack.pop();
    if (!st) return false;
    this.apply(st.after);
    if (st.sheetsAfter) { this.wb.sheets = [...st.sheetsAfter.list]; this.wb.active = st.sheetsAfter.active; }
    this.restoreSel(st.selAfter);
    this.undoStack.push(st);
    this.onChange(new Set(st.after.map((b) => b.sheet)));
    return true;
  }

  private restoreSel(sel: Step["selBefore"]): void {
    if (!sel || !this.wb.sheets[sel.sheet]) return;
    this.wb.active = sel.sheet;
    const v = this.wb.sheets[sel.sheet].view;
    v.selection = sel.sel.map((r) => ({ ...r }));
    v.active = { ...sel.active };
  }

  clear(): void {
    this.undoStack = [];
    this.redoStack = [];
  }
}

/** Shallow copies: values are primitives or shared, immutable CellError objects (structuredClone would lose their class). */
function clone<T>(v: T): T {
  if (v === undefined || v === null) return v;
  const c = { ...(v as object) } as T & { array?: Rect };
  if (c.array) c.array = { ...c.array };
  return c;
}

function snapshot(s: Sheet) {
  return {
    cells: new Map([...s.cells].map(([k, c]) => [k, clone(c)])),
    cols: new Map([...s.cols].map(([k, c]) => [k, { ...c }])),
    rows: new Map([...s.rows].map(([k, c]) => [k, { ...c }])),
    merges: s.merges.map((m) => ({ ...m })),
  };
}
