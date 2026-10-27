// Recalculation of a workbook (D-062): a dependency graph from every formula to the cells and
// ranges it reads, so that after an edit only the formulas depending on it are computed again.
// Volatile functions (JETZT, HEUTE, ZUFALLSZAHL, ZUFALLSBEREICH, ZUFALLSMATRIX, INDIREKT,
// BEREICH.VERSCHIEBEN) are computed on every recalculation, as in Excel. A circular reference
// gives 0, as Excel shows it with iteration switched off, and is reported in `circular`.
// Dynamic arrays spill into the cells below and to the right, or give #ÜBERLAUF! when a cell
// there is not empty.

import { MAX_COLS, MAX_ROWS, cellKey, keyCell, type Rect } from "../address";
import type { Cell, Value, Workbook } from "../model";
import { serialOfDate } from "./dates";
import { Ctx, FUNCS, evalFormula, type Host } from "./eval";
import "./functions";
import { arrayRows } from "./functions";
import { FormulaError, parseFormula, type Node } from "./parse";
import { Arr, E, type V } from "./values";

const CLEAN = 0, DIRTY = 1, BUSY = 2;

/** One formula cell. */
interface FNode {
  sheet: number;
  key: number;
  row: number;
  col: number;
  f: string;
  node: Node | null;
  volatile: boolean;
  deps: { sheet: number; rect: Rect }[];
  state: number;
  /** Dynamic array: the rectangle it wanted to fill (also when blocked) and what it wrote. */
  wanted?: Rect;
  spill?: Rect;
  written?: V[];
  /** The cell is in a circular reference and keeps 0 for this pass. */
  frozen?: boolean;
}

/** Thrown when a formula needs one that is not computed yet and the recursion is deep. */
class NeedFirst {
  constructor(readonly node: FNode) {}
}

export interface CellAt { sheet: number; row: number; col: number }

const BUCKET = 5; // 32 × 32 cells
const BUCKET_COLS = MAX_COLS >> BUCKET;
const MAX_DEPTH = 300;

/** Range dependencies of one sheet, found by cell. */
class RangeIndex {
  private buckets = new Map<number, FNode[]>();
  private big: { rect: Rect; n: FNode }[] = [];
  private rects = new Map<FNode, Rect[]>();

  add(n: FNode, r: Rect): void {
    let list = this.rects.get(n);
    if (!list) this.rects.set(n, (list = []));
    list.push(r);
    const b0 = r.r0 >> BUCKET, b1 = r.r1 >> BUCKET, c0 = r.c0 >> BUCKET, c1 = r.c1 >> BUCKET;
    if ((b1 - b0 + 1) * (c1 - c0 + 1) > 64) {
      this.big.push({ rect: r, n });
      return;
    }
    for (let b = b0; b <= b1; b++) {
      for (let c = c0; c <= c1; c++) {
        const k = b * BUCKET_COLS + c;
        const l = this.buckets.get(k);
        if (l) { if (l[l.length - 1] !== n) l.push(n); } else this.buckets.set(k, [n]);
      }
    }
  }

  remove(n: FNode): void {
    const list = this.rects.get(n);
    if (!list) return;
    this.rects.delete(n);
    for (const r of list) {
      const b0 = r.r0 >> BUCKET, b1 = r.r1 >> BUCKET, c0 = r.c0 >> BUCKET, c1 = r.c1 >> BUCKET;
      if ((b1 - b0 + 1) * (c1 - c0 + 1) > 64) {
        this.big = this.big.filter((x) => x.n !== n);
        continue;
      }
      for (let b = b0; b <= b1; b++) {
        for (let c = c0; c <= c1; c++) {
          const k = b * BUCKET_COLS + c;
          const l = this.buckets.get(k);
          if (!l) continue;
          const i = l.indexOf(n);
          if (i >= 0) l.splice(i, 1);
          if (!l.length) this.buckets.delete(k);
        }
      }
    }
  }

  /** Formulas whose ranges contain the cell. */
  each(row: number, col: number, f: (n: FNode) => void): void {
    const l = this.buckets.get((row >> BUCKET) * BUCKET_COLS + (col >> BUCKET));
    if (l) {
      for (const n of l) {
        for (const r of this.rects.get(n)!) {
          if (row >= r.r0 && row <= r.r1 && col >= r.c0 && col <= r.c1) { f(n); break; }
        }
      }
    }
    for (const { rect: r, n } of this.big) if (row >= r.r0 && row <= r.r1 && col >= r.c0 && col <= r.c1) f(n);
  }
}

export interface CalcOptions {
  /** The current time for JETZT and HEUTE (default: the clock). */
  now?: () => Date;
  random?: () => number;
}

export class Calc implements Host {
  private nodes = new Map<number, Map<number, FNode>>();
  private cellDeps = new Map<number, Map<number, Set<FNode>>>();
  private rangeDeps = new Map<number, RangeIndex>();
  private bnds = new Map<number, Rect | null>();
  private nowSerial = 0;
  private depth = 0;
  private path: FNode[] = [];
  private loops = new Set<FNode>();
  private changedCells = new Map<string, CellAt>();
  /** Spill children: sheet → cell key → the formula that wrote it. */
  private owner = new Map<number, Map<number, FNode>>();
  /** Formulas with a dynamic array result (spilled or blocked). */
  private anchors = new Set<FNode>();
  private volatiles = new Set<FNode>();
  private recomputed = new Set<FNode>();
  private newLoops: CellAt[] = [];
  private pending: FNode[] = [];
  private rounds = 0;

  /** Cells in circular references found by the last recalculation. */
  circular: CellAt[] = [];

  constructor(readonly wb: Workbook, private opts: CalcOptions = {}) {}

  // ----- Host --------------------------------------------------------------------------------

  value(sheet: number, row: number, col: number): V {
    const key = cellKey(row, col);
    // A cell a dynamic array spills into: that formula first.
    const a = this.owner.get(sheet)?.get(key);
    if (a && a.state === DIRTY) {
      if (this.depth >= MAX_DEPTH) throw new NeedFirst(a);
      this.compute(a);
    }
    const n = this.nodes.get(sheet)?.get(key);
    if (n) {
      if (n.frozen) return 0;
      if (n.state === BUSY) {
        // A circular reference: every formula on the path back to this one is in the loop.
        const i = this.path.indexOf(n);
        for (const m of i >= 0 ? this.path.slice(i) : [n]) this.loops.add(m);
        return 0;
      }
      if (n.state === DIRTY) {
        if (this.depth >= MAX_DEPTH) throw new NeedFirst(n);
        this.compute(n);
      }
    }
    return this.wb.sheets[sheet]?.value(row, col) ?? null;
  }

  bounds(sheet: number): Rect | null {
    if (!this.bnds.has(sheet)) this.bnds.set(sheet, this.wb.sheets[sheet]?.usedRange() ?? null);
    return this.bnds.get(sheet)!;
  }

  now(): number {
    return this.nowSerial;
  }

  random(): number {
    return (this.opts.random ?? Math.random)();
  }

  formula(sheet: number, row: number, col: number): string | undefined {
    return this.wb.sheets[sheet]?.get(row, col)?.f;
  }

  spill(sheet: number, row: number, col: number): Rect | null {
    const n = this.nodes.get(sheet)?.get(cellKey(row, col));
    if (n && n.state === DIRTY) this.value(sheet, row, col);
    return n?.spill ?? null;
  }

  // ----- Public API --------------------------------------------------------------------------

  /** Computes every formula of the workbook (after loading, or after sheets were added, renamed or removed). */
  recalc(): void {
    this.begin();
    this.nodes.clear();
    this.cellDeps.clear();
    this.rangeDeps.clear();
    this.owner.clear();
    this.anchors.clear();
    this.volatiles.clear();
    this.circular = [];
    this.wb.sheets.forEach((sh, s) => {
      for (const [k, c] of sh.cells) if (c.f !== undefined) this.addNode(s, k, c);
    });
    // Cells a dynamic array spilled into when the file was saved belong to it.
    for (const [, m] of this.nodes) for (const n of m.values()) this.adoptSpill(n);
    const all: FNode[] = [];
    for (const [, m] of this.nodes) for (const n of m.values()) { n.state = DIRTY; all.push(n); }
    this.run(all);
    this.finish();
  }

  /**
   * After the app changed cells (values or formulas) of one sheet: recomputes what depends on
   * them and every volatile formula, and returns every cell whose value changed (formula cells
   * and cells a dynamic array spills into; not the edited cells themselves).
   */
  changed(sheetIndex: number, cells: { row: number; col: number }[]): CellAt[] {
    this.begin();
    const sheet = this.wb.sheets[sheetIndex];
    const start: FNode[] = [];
    for (const { row, col } of cells) {
      const k = cellKey(row, col);
      const c = sheet.cells.get(k);
      const old = this.nodes.get(sheetIndex)?.get(k);
      if (old && (c?.f === undefined || c.f !== old.f)) {
        this.removeNode(old);
        if (old.spill) this.clearSpill(old);
      }
      if (c?.f !== undefined && (!old || c.f !== old.f)) {
        const n = this.addNode(sheetIndex, k, c);
        this.adoptSpill(n);
        start.push(n);
      } else if (old && c?.f !== undefined) {
        start.push(old);
      }
      // A dynamic array that wants this cell computes again (it may spill now or no longer).
      for (const n of this.anchors) {
        const w = n.wanted;
        if (!w || n.sheet !== sheetIndex || row < w.r0 || row > w.r1 || col < w.c0 || col > w.c1 || (n.row === row && n.col === col)) continue;
        if (this.owner.get(sheetIndex)?.get(k) === n && this.isOwn(n, row, col, c)) continue;
        start.push(n);
      }
    }
    const dirty: FNode[] = [];
    const mark = (n: FNode) => {
      if (n.state === DIRTY) return;
      n.state = DIRTY;
      dirty.push(n);
    };
    for (const n of start) mark(n);
    for (const { row, col } of cells) this.dependents(sheetIndex, row, col, mark);
    for (const n of this.volatiles) mark(n);
    // Everything downstream of a dirty formula is dirty too.
    for (let i = 0; i < dirty.length; i++) {
      const n = dirty[i];
      this.dependents(n.sheet, n.row, n.col, mark);
      if (n.spill) this.eachIn(n.spill, (r, c) => this.dependents(n.sheet, r, c, mark));
    }
    this.run(dirty);
    this.finish();
    return [...this.changedCells.values()];
  }

  /** The formula whose dynamic array covers this cell (for greyed formulas in the Bearbeitungsleiste). */
  spillAnchor(sheet: number, row: number, col: number): { row: number; col: number } | null {
    const n = this.owner.get(sheet)?.get(cellKey(row, col));
    return n ? { row: n.row, col: n.col } : null;
  }

  /** Evaluates a formula (file form) as if it stood in a cell, with the current values. */
  evaluate(sheet: number, row: number, col: number, file: string): Value | Value[][] {
    let node: Node;
    try {
      node = parseFormula(file.startsWith("=") ? file.slice(1) : file);
    } catch (e) {
      if (e instanceof FormulaError) return E.name;
      throw e;
    }
    this.bnds.clear();
    const r = evalFormula(node, new Ctx(this, sheet, row, col));
    return r instanceof Arr ? arrayRows(r) : r === null ? 0 : r;
  }

  // ----- Graph -------------------------------------------------------------------------------

  private begin(): void {
    this.bnds.clear();
    this.loops.clear();
    this.changedCells.clear();
    this.recomputed.clear();
    this.newLoops = [];
    this.rounds = 0;
    this.nowSerial = serialOfDate((this.opts.now ?? (() => new Date()))());
  }

  private addNode(sheet: number, key: number, c: Cell): FNode {
    const { row, col } = keyCell(key);
    let node: Node | null = null;
    try {
      node = parseFormula(c.f!);
    } catch (e) {
      if (!(e instanceof FormulaError)) throw e;
    }
    const n: FNode = { sheet, key, row, col, f: c.f!, node, volatile: false, deps: [], state: CLEAN };
    if (c.spill) n.spill = c.spill;
    if (node) this.collect(node, n, new Set());
    if (n.volatile) this.volatiles.add(n);
    let m = this.nodes.get(sheet);
    if (!m) this.nodes.set(sheet, (m = new Map()));
    m.set(key, n);
    for (const d of n.deps) {
      const r = d.rect;
      if (r.r0 === r.r1 && r.c0 === r.c1) {
        let cm = this.cellDeps.get(d.sheet);
        if (!cm) this.cellDeps.set(d.sheet, (cm = new Map()));
        const k = cellKey(r.r0, r.c0);
        let set = cm.get(k);
        if (!set) cm.set(k, (set = new Set()));
        set.add(n);
      } else {
        let ri = this.rangeDeps.get(d.sheet);
        if (!ri) this.rangeDeps.set(d.sheet, (ri = new RangeIndex()));
        ri.add(n, r);
      }
    }
    return n;
  }

  private removeNode(n: FNode): void {
    this.nodes.get(n.sheet)?.delete(n.key);
    this.anchors.delete(n);
    this.volatiles.delete(n);
    for (const d of n.deps) {
      const r = d.rect;
      if (r.r0 === r.r1 && r.c0 === r.c1) this.cellDeps.get(d.sheet)?.get(cellKey(r.r0, r.c0))?.delete(n);
    }
    for (const [, ri] of this.rangeDeps) ri.remove(n);
  }

  /** The references a formula reads (through defined names too), and whether it is volatile. */
  private collect(node: Node, n: FNode, seen: Set<string>): void {
    switch (node.t) {
      case "ref": {
        const s = node.sheet === undefined ? n.sheet : this.sheetIndex(node.sheet);
        if (s < 0) return;
        const s2 = node.sheet2 === undefined ? s : this.sheetIndex(node.sheet2);
        for (let i = Math.min(s, s2); i <= Math.max(s, s2); i++) n.deps.push({ sheet: i, rect: { r0: node.r0, c0: node.c0, r1: node.r1, c1: node.c1 } });
        return;
      }
      case "name": {
        const l = node.name.toLowerCase();
        if (seen.has(l)) return;
        seen.add(l);
        const home = node.sheet !== undefined ? this.sheetIndex(node.sheet) : n.sheet;
        const dn = this.wb.names.find((d) => d.name.toLowerCase() === l && d.sheet === home) ?? this.wb.names.find((d) => d.name.toLowerCase() === l && d.sheet === undefined);
        if (!dn) return;
        try {
          this.collect(parseFormula(dn.ref.startsWith("=") ? dn.ref.slice(1) : dn.ref), n, seen);
        } catch {
          // An unreadable name gives #NAME? and depends on nothing.
        }
        return;
      }
      case "fn":
        if (FUNCS.get(node.name)?.volatile) n.volatile = true;
        for (const a of node.args) this.collect(a, n, seen);
        return;
      case "un":
        // A1# reads whatever the formula in A1 spills: recomputed every time, like Excel's volatile ones.
        if (node.op === "#") n.volatile = true;
        this.collect(node.a, n, seen);
        return;
      case "bin":
        this.collect(node.a, n, seen);
        this.collect(node.b, n, seen);
        return;
      case "paren":
        this.collect(node.a, n, seen);
        return;
      default:
        return;
    }
  }

  private sheetIndex(name: string): number {
    const l = name.toLowerCase();
    return this.wb.sheets.findIndex((s) => s.name.toLowerCase() === l);
  }

  private dependents(sheet: number, row: number, col: number, f: (n: FNode) => void): void {
    const set = this.cellDeps.get(sheet)?.get(cellKey(row, col));
    if (set) for (const n of set) f(n);
    this.rangeDeps.get(sheet)?.each(row, col, f);
  }

  private eachIn(r: Rect, f: (row: number, col: number) => void): void {
    for (let row = r.r0; row <= r.r1; row++) for (let col = r.c0; col <= r.c1; col++) f(row, col);
  }

  // ----- Computing ---------------------------------------------------------------------------

  /** Computes the dirty formulas; deep chains are taken one level at a time (NeedFirst). */
  private run(list: FNode[]): void {
    for (const first of list) {
      if (first.state !== DIRTY) continue;
      const stack = [first];
      while (stack.length) {
        const n = stack[stack.length - 1];
        if (n.state !== DIRTY) { stack.pop(); continue; }
        try {
          this.depth = 0;
          this.path = [];
          this.compute(n);
          stack.pop();
        } catch (e) {
          if (!(e instanceof NeedFirst)) throw e;
          const i = stack.indexOf(e.node);
          if (i >= 0) {
            // A loop longer than the recursion limit.
            for (const m of stack.slice(i)) this.loops.add(m);
            this.freeze(stack.slice(i));
            stack.length = i;
          } else {
            stack.push(e.node);
          }
        }
      }
    }
    if (this.loops.size) this.settleLoops();
    // Bounded: a formula reading its own spill range would otherwise go round for ever.
    if (this.pending.length && this.rounds++ < 100) this.run(this.pending.splice(0));
    this.pending.length = 0;
  }

  private compute(n: FNode): void {
    this.recomputed.add(n);
    n.state = BUSY;
    this.depth++;
    this.path.push(n);
    try {
      const sheet = this.wb.sheets[n.sheet];
      const cell = sheet.cells.get(n.key);
      if (!cell) { n.state = CLEAN; return; }
      let r: V | Arr;
      if (!n.node) r = E.name;
      else r = evalFormula(n.node, new Ctx(this, n.sheet, n.row, n.col));
      if (cell.array) this.fillArray(n, cell, r);
      else if (r instanceof Arr) this.spillOut(n, cell, r);
      else {
        if (n.spill) this.clearSpill(n);
        n.wanted = undefined;
        this.anchors.delete(n);
        this.store(n.sheet, n.row, n.col, cell, r === null ? 0 : r);
      }
      n.state = CLEAN;
    } catch (e) {
      n.state = DIRTY;
      throw e;
    } finally {
      this.depth--;
      this.path.pop();
    }
  }

  private store(sheet: number, row: number, col: number, cell: Cell, v: V): void {
    if (cell.v !== v && !(typeof v === "number" && typeof cell.v === "number" && Object.is(v, cell.v))) {
      cell.v = v;
      this.touched(sheet, row, col);
    }
  }

  /** A value changed during the pass: formulas that already read it compute again. */
  private touched(sheet: number, row: number, col: number): void {
    this.changedCells.set(`${sheet}:${row}:${col}`, { sheet, row, col });
    this.dependents(sheet, row, col, (m) => {
      if (m.state !== CLEAN || m.frozen) return;
      m.state = DIRTY;
      this.pending.push(m);
    });
  }

  /** Whether a cell inside a formula's spill range still holds what the formula wrote there. */
  private isOwn(n: FNode, row: number, col: number, c: Cell | undefined): boolean {
    if (!n.spill || !n.written) return false;
    if (!c) return true;
    if (c.f !== undefined) return false;
    const i = (row - n.spill.r0) * (n.spill.c1 - n.spill.c0 + 1) + (col - n.spill.c0);
    return c.v === n.written[i] || c.v === null;
  }

  private spillOut(n: FNode, cell: Cell, a: Arr): void {
    const sheet = this.wb.sheets[n.sheet];
    const want = { r0: n.row, c0: n.col, r1: n.row + a.rows - 1, c1: n.col + a.cols - 1 };
    n.wanted = want;
    this.anchors.add(n);
    let blocked = want.r1 >= MAX_ROWS || want.c1 >= MAX_COLS;
    if (!blocked) {
      const own = this.owner.get(n.sheet);
      for (let r = want.r0; r <= want.r1 && !blocked; r++) {
        for (let c = want.c0; c <= want.c1; c++) {
          if (r === n.row && c === n.col) continue;
          const k = cellKey(r, c);
          const x = sheet.cells.get(k);
          if (!x || (x.v === null && x.f === undefined)) continue;
          if (own?.get(k) === n && this.isOwn(n, r, c, x)) continue;
          blocked = true;
          break;
        }
      }
    }
    if (n.spill) this.clearSpill(n, blocked ? undefined : want);
    if (blocked) {
      cell.spill = undefined;
      this.store(n.sheet, n.row, n.col, cell, E.spill);
      return;
    }
    let own = this.owner.get(n.sheet);
    if (!own) this.owner.set(n.sheet, (own = new Map()));
    const written: V[] = [];
    for (let r = 0; r < a.rows; r++) {
      for (let c = 0; c < a.cols; c++) {
        const v = a.d[r * a.cols + c] ?? 0;
        const vv = v === null ? 0 : v;
        written.push(vv);
        if (r === 0 && c === 0) continue;
        const row = n.row + r, col = n.col + c;
        const k = cellKey(row, col);
        let x = sheet.cells.get(k);
        if (!x) {
          x = { v: null };
          sheet.cells.set(k, x);
          this.grow(n.sheet, row, col);
        }
        own.set(k, n);
        this.store(n.sheet, row, col, x, vv);
      }
    }
    n.spill = want;
    n.written = written;
    cell.spill = want;
    this.store(n.sheet, n.row, n.col, cell, written[0]);
  }

  /** Empties the cells a formula spilled into (keeping those of `keep`). */
  private clearSpill(n: FNode, keep?: Rect): void {
    const sp = n.spill;
    if (!sp) return;
    const sheet = this.wb.sheets[n.sheet];
    const own = this.owner.get(n.sheet);
    this.eachIn(sp, (row, col) => {
      if (row === n.row && col === n.col) return;
      if (keep && row >= keep.r0 && row <= keep.r1 && col >= keep.c0 && col <= keep.c1) return;
      const k = cellKey(row, col);
      if (own?.get(k) !== n) return;
      own.delete(k);
      const x = sheet.cells.get(k);
      if (x && this.isOwn(n, row, col, x)) {
        if (x.v !== null) {
          x.v = null;
          this.touched(n.sheet, row, col);
        }
        if (x.s === undefined && !x.note && !x.link) sheet.cells.delete(k);
      }
    });
    n.spill = undefined;
    const cell = sheet.cells.get(n.key);
    if (cell && !keep) cell.spill = undefined;
  }

  /** Cells of a spill range saved in the file belong to their formula. */
  private adoptSpill(n: FNode): void {
    if (!n.spill) return;
    const sheet = this.wb.sheets[n.sheet];
    let own = this.owner.get(n.sheet);
    if (!own) this.owner.set(n.sheet, (own = new Map()));
    const written: V[] = [];
    this.eachIn(n.spill, (row, col) => {
      const k = cellKey(row, col);
      const x = sheet.cells.get(k);
      written.push(x && x.f === undefined ? x.v : null);
      if (!(row === n.row && col === n.col) && (!x || x.f === undefined)) own!.set(k, n);
    });
    n.written = written;
    n.wanted = n.spill;
    this.anchors.add(n);
  }

  /** An array formula entered with Strg+Umschalt+Eingabe fills its fixed range (#NV beyond the result). */
  private fillArray(n: FNode, cell: Cell, r: V | Arr): void {
    const rect = cell.array!;
    const sheet = this.wb.sheets[n.sheet];
    const a = r instanceof Arr ? r : new Arr(1, 1, [r]);
    for (let row = rect.r0; row <= rect.r1; row++) {
      for (let col = rect.c0; col <= rect.c1; col++) {
        const i = row - rect.r0, j = col - rect.c0;
        const ii = a.rows === 1 ? 0 : i, jj = a.cols === 1 ? 0 : j;
        const v = ii < a.rows && jj < a.cols ? a.d[ii * a.cols + jj] ?? 0 : E.na;
        const k = cellKey(row, col);
        let x = sheet.cells.get(k);
        if (!x) {
          x = { v: null };
          sheet.cells.set(k, x);
          this.grow(n.sheet, row, col);
        }
        this.store(n.sheet, row, col, x, v === null ? 0 : v);
      }
    }
  }

  private grow(sheet: number, row: number, col: number): void {
    if (!this.bnds.has(sheet)) return;
    const b = this.bnds.get(sheet);
    this.bnds.set(sheet, b ? { r0: Math.min(b.r0, row), c0: Math.min(b.c0, col), r1: Math.max(b.r1, row), c1: Math.max(b.c1, col) } : { r0: row, c0: col, r1: row, c1: col });
  }

  private freeze(list: FNode[]): void {
    for (const n of list) {
      n.frozen = true;
      n.state = CLEAN;
      const cell = this.wb.sheets[n.sheet].cells.get(n.key);
      if (cell) this.store(n.sheet, n.row, n.col, cell, 0);
    }
  }

  /** Cells in a loop show 0 (Excel without iteration); what depends on them is computed again. */
  private settleLoops(): void {
    const loop = [...this.loops];
    this.freeze(loop);
    const again: FNode[] = [];
    const mark = (m: FNode) => {
      if (m.frozen || m.state === DIRTY) return;
      m.state = DIRTY;
      again.push(m);
    };
    for (const n of loop) this.dependents(n.sheet, n.row, n.col, mark);
    for (let i = 0; i < again.length; i++) this.dependents(again[i].sheet, again[i].row, again[i].col, mark);
    this.loops.clear();
    this.newLoops.push(...loop.map((n) => ({ sheet: n.sheet, row: n.row, col: n.col })));
    this.run(again);
  }

  private finish(): void {
    const loopNodes = new Set<FNode>();
    for (const [, m] of this.nodes) for (const n of m.values()) if (n.frozen) { n.frozen = false; loopNodes.add(n); }
    // Loops found earlier stay reported until one of their formulas is computed again.
    const keep = this.circular.filter((c) => {
      const n = this.nodes.get(c.sheet)?.get(cellKey(c.row, c.col));
      return n !== undefined && !this.recomputed.has(n) && !loopNodes.has(n);
    });
    const seen = new Set<string>();
    this.circular = [...keep, ...this.newLoops].filter((c) => {
      const k = `${c.sheet}:${c.row}:${c.col}`;
      if (seen.has(k)) return false;
      seen.add(k);
      return true;
    });
  }
}
