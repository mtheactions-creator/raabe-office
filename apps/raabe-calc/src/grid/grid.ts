// The sheet window: canvas, scroll bars, the cell editor, Namenfeld and Bearbeitungsleiste,
// and Excel's controls on them (D-061): the cell pointer and selections, Eingeben, Bearbeiten
// and Zeigen modes, Strg+arrow jumps, the fill handle, moving cells by their border, header
// sizing, AutoVervollständigen and pointing at cells while a formula is typed.

import {
  MAX_COLS, MAX_ROWS, cellName, colName, formulaRefs, inRect, normRect, parseRange, rangeName,
  type Rect, type Sheet, type Workbook,
} from "@raabe/sheet-engine";
import { h } from "@raabe/shell";
import { editText } from "../engine";
import { cssFont, draw, errorIndicator, expandToMerges, filterButton, placed, type DrawState, type Pane } from "./render";
import { CondEval } from "./cond";
import { POINTERS } from "./pointers";
import { pxToChars, setDeviceScale, sheetAxes, type Axis } from "./metrics";
import type { Shown } from "../engine";
import { LANE, barSize, groups, maxLevel, showLevel, toggle, type Axis as OutlineAxis } from "./outline";

export type Mode = "ready" | "enter" | "edit" | "point";

export interface GridHost {
  workbook(): Workbook;
  /** Typed input committed into `targets` (Strg+Eingabe: every selected cell). */
  commit(targets: Rect, text: string, opts: { array?: boolean }): void;
  clearContents(): void;
  /** The fill handle dragged back into its selection: these cells are emptied. */
  clearRange(r: Rect): void;
  /** Fill handle released: fill from `source` to `target` (Strg held: series/copy toggled). */
  fill(source: Rect, target: Rect, ctrl: boolean): void;
  /** Umschalt+drag of the fill handle: empty cells inserted ("down"/"right") or cells deleted ("up"/"left"). */
  shiftFill?(rect: Rect, dir: "down" | "right" | "up" | "left"): void;
  /** Selection border dragged: move (or copy with Strg) to a new top-left cell. */
  move(source: Rect, to: { row: number; col: number }, copy: boolean): void;
  /** A drag with the right mouse button ended: Excel's menu of what to do (fill handle or border). */
  rightDrag?(kind: "fill" | "move", source: Rect, target: Rect, at: { x: number; y: number }): void;
  /** Umschalt+drag of the border: the cells are taken out and inserted at `at` (before that row, or column). */
  moveInsert?(source: Rect, at: { row: number; col: number }, rows: boolean): Rect | undefined;
  /** `before`: when the border was dragged past the start, the columns (rows) ahead shrunk too. */
  resizeCols(cols: number[], px: number, before?: { i: number; size: number }[]): void;
  resizeRows(rows: number[], px: number, before?: { i: number; size: number }[]): void;
  autofitCols(cols: number[]): void;
  autofitRows(rows: number[]): void;
  selectionChanged(): void;
  modeChanged(mode: Mode): void;
  /** Copied cells wait to be pasted (the status bar says so instead of the cell mode). */
  copyMode?(on: boolean): void;
  /** The status bar's mode texts besides the cell mode: [field, text] (Endmodus, Auswahlmodus). */
  flagsChanged?(flags: [string, string][]): void;
  contextMenu(e: MouseEvent, kind: "cell" | "col" | "row"): void;
  /** Keys the grid does not handle itself (ribbon shortcuts …); true when used. */
  key(e: KeyboardEvent): boolean;
  /** An AutoFilter button was clicked (its column and where the button is on the page). */
  filterMenu(col: number, anchor: DOMRect): void;
  /** Formeln anzeigen (Strg+#). */
  showFormulas(): boolean;
  /** Datenüberprüfung: "ok", or the user's answer to the error box (retry: edit again; cancel: drop it). */
  validate?(row: number, col: number, text: string): "ok" | Promise<"ok" | "retry" | "cancel">;
  /** After every paint (things placed over the grid follow scrolling and zoom). */
  painted?(): void;
  /** Blatt schützen: false (after telling the user) where the cell may not be changed. */
  canEdit?(row: number, col: number): boolean;
  /** Blatt schützen: Tab goes to the next unlocked cell (null: the sheet is not protected). */
  nextUnlocked?(row: number, col: number, back: boolean): { row: number; col: number } | null;
  /** Ungültige Daten einkreisen: the circled cells of the sheet. */
  circles?(): { row: number; col: number }[];
  /** Umbruchvorschau: the printed area and pages to draw, or null in Normal view. */
  pageBreaks?(): { area: Rect; pages: { rows: [number, number]; cols: [number, number] }[]; manualRows: number[]; manualCols: number[] } | null;
  /** Umbruchvorschau: a page break dragged from before row/column `from` to before `to`. */
  movePageBreak?(rows: boolean, from: number, to: number): void;
  /** A change to the Gliederung (a level button or −/+): done as one undo step by the host. */
  outline?(axis: "rows" | "cols", change: () => void): void;
  /** Spur zum Vorgänger/Nachfolger arrows of the sheet. */
  arrows?(): { from: Rect | null; to: { row: number; col: number }; error?: boolean }[];
  /** The warning button beside a cell with an error indicator was pressed. */
  errorMenu(kind: "text-number" | "error", anchor: DOMRect): void;
  /** Enter while copied cells are marked: paste them; true when it did. */
  enterPaste(): boolean;
  /** The Namenfeld's arrow: the workbook's names to go to. */
  namesMenu(anchor: DOMRect): void;
  /** A link was clicked: a web address, a file or "#Sheet!A1" inside the workbook. */
  followLink(target: string): void;
}

const REF_COLORS = ["#0070c0", "#c00000", "#7030a0", "#00b050", "#c55a11", "#2e75b6", "#bf8f00"];
const PT = 96 / 72;

export class Grid {
  readonly el: HTMLElement;
  readonly formulaBar: HTMLElement;
  readonly hscroll: HTMLElement;
  private canvas = h("canvas", { class: "grid-canvas" }) as HTMLCanvasElement;
  private ctx = this.canvas.getContext("2d", { alpha: false })!;
  private vscroll = h("div", { class: "vscroll" }, h("div"));
  private editor = h("textarea", { class: "cell-editor", spellcheck: "false", "aria-label": "Zelleninhalt", wrap: "off" }) as HTMLTextAreaElement;
  private nameBox = h("input", { class: "namebox", "aria-label": "Namenfeld", spellcheck: "false" }) as HTMLInputElement;
  private fx = h("textarea", { class: "fx-input", "aria-label": "Bearbeitungsleiste", spellcheck: "false", rows: 1 }) as HTMLTextAreaElement;
  private ac = h("div", { class: "fn-ac", hidden: true });
  private rows!: Axis;
  private cols!: Axis;
  private zoom = 1;
  private width = 0;
  private height = 0;
  private dpr = devicePixelRatio || 1;
  private textCache = new Map<number, Shown>();
  private frame = 0;
  mode: Mode = "ready";
  private editing: { row: number; col: number; original: string; viaBar: boolean } | null = null;
  /** Where Shift+arrows and Shift+click extend to (the active cell is the other corner). */
  private head = { row: 0, col: 0 };
  /** Tab moves remember their starting column; Enter returns to it (Excel). */
  private tabStart: number | null = null;
  private drag: null | { kind: "cells" | "cols" | "rows" | "fill" | "move" | "resizeCol" | "resizeRow" | "point" | "pageBreak"; start: { row: number; col: number }; x: number; y: number; extra?: unknown } = null;
  private ghost: Rect | null = null;
  marquee: Rect | null = null;
  private marqueePhase = 0;
  private marqueeTimer = 0;
  /** The reference being pointed at while a formula is typed (Zeigen). */
  private pointRef: { start: number; end: number; anchor: { row: number; col: number }; head: { row: number; col: number } } | null = null;
  private suppressScroll = false;

  constructor(private host: GridHost) {
    const cancel = h("button", { class: "fx-btn", "aria-label": "Abbrechen", "data-tip": "Abbrechen", onclick: () => this.cancelEdit() }, "✕");
    const ok = h("button", { class: "fx-btn", "aria-label": "Eingeben", "data-tip": "Eingeben", onclick: () => this.commitEdit("none") }, "✓");
    const insertFn = h("button", { class: "fx-btn fx-fn", "aria-label": "Funktion einfügen", "data-tip": "Funktion einfügen", onclick: () => this.startFormulaFromButton() }, h("i", {}, "fx"));
    this.formulaBar = h("div", { class: "fbar" },
      h("div", { class: "namebox-wrap" }, this.nameBox, h("button", { class: "namebox-dd", "aria-label": "Namenfeld", tabindex: "-1",
        onmousedown: (e: MouseEvent) => { e.preventDefault(); this.host.namesMenu((e.currentTarget as HTMLElement).closest(".namebox-wrap")!.getBoundingClientRect()); } }, "⌄")),
      h("div", { class: "fbar-sep" }, "⋮"),
      h("div", { class: "fx-btns" }, cancel, ok, insertFn),
      h("div", { class: "fx-wrap" }, this.fx));
    this.hscroll = h("div", { class: "hscroll" }, h("div"));
    this.el = h("div", { class: "gridwrap" }, this.canvas, this.editor, this.vscroll, this.ac);
    this.installEvents();
    new ResizeObserver(() => this.resize()).observe(this.el);
  }

  // ----- state -----------------------------------------------------------------------------

  get wb(): Workbook { return this.host.workbook(); }
  get sheet(): Sheet { return this.wb.sheet(); }
  get view() { return this.sheet.view; }
  get selection(): Rect[] { return this.view.selection; }
  get active() { return this.view.active; }
  get isEditing(): boolean { return this.editing !== null; }

  /** Sizes changed (columns, rows, sheet switched): rebuild the axes. */
  refreshLayout(): void {
    this.usedCache = null;
    setDeviceScale(this.dpr);
    const a = sheetAxes(this.sheet);
    this.rows = a.rows;
    this.cols = a.cols;
    this.zoom = this.view.zoom / 100;
    this.textCache.clear();
    this.head = { ...this.active };
    this.updateScrollbars();
    this.invalidate();
  }

  private previewState: DrawState["preview"] = null;

  /** Livevorschau of a format on the selection (null ends it); nothing is changed. */
  preview(change: ((s: import("@raabe/sheet-engine").CellStyle) => void) | null): void {
    this.previewState = change ? { ranges: this.selection, change } : null;
    this.textCache.clear();
    this.invalidate();
  }

  private cond: CondEval | null = null;
  private condFor: Sheet | null = null;

  /** Cell contents changed: redraw (formatted texts are cached per cell). */
  contentChanged(): void {
    this.textCache.clear();
    this.usedCache = null;
    this.cond = null;
    this.invalidate();
    this.syncBar();
  }

  setZoom(z: number): void {
    this.view.zoom = Math.max(10, Math.min(400, Math.round(z)));
    this.zoom = this.view.zoom / 100;
    this.textCache.clear();
    this.updateScrollbars();
    this.invalidate();
  }

  focus(): void {
    if (this.editing?.viaBar) this.fx.focus();
    else this.editor.focus({ preventScroll: true });
  }

  invalidate(): void {
    if (this.frame) return;
    this.frame = requestAnimationFrame(() => {
      this.frame = 0;
      this.paint();
    });
  }

  private resize(): void {
    const r = this.el.getBoundingClientRect();
    const dpr = devicePixelRatio || 1;
    const rescale = dpr !== this.dpr;
    this.dpr = dpr;
    if (rescale && this.rows) this.refreshLayout();
    this.width = Math.max(0, r.width - 17);
    this.height = Math.max(0, r.height);
    this.canvas.width = Math.round(this.width * this.dpr);
    this.canvas.height = Math.round(this.height * this.dpr);
    this.canvas.style.width = `${this.canvas.width / this.dpr}px`;
    this.canvas.style.height = `${this.canvas.height / this.dpr}px`;
    if (this.rows) {
      this.updateScrollbars();
      this.paint();
    }
  }

  /** The Gliederung bar left of the row headers (width) and above the column headers (height). */
  private outlineW(): number {
    return this.view.showHeaders ? barSize(this.sheet, "rows", this.zoom) : 0;
  }

  private outlineH(): number {
    return this.view.showHeaders ? barSize(this.sheet, "cols", this.zoom) : 0;
  }

  private headerW(): number {
    if (!this.view.showHeaders) return 0;
    const last = this.rows.at(this.rows.pos(this.view.top) + this.height / this.zoom) + 1;
    const digits = Math.max(3, String(last).length);
    return Math.round((digits * 7 + 6) * this.zoom) + this.outlineW();
  }

  private headerH(): number {
    return this.view.showHeaders ? Math.round(this.rows.def * this.zoom) + this.outlineH() : 0;
  }

  private panes(): Pane[] {
    const hw = this.headerW(), hh = this.headerH();
    const fr = this.view.freeze?.row ?? 0, fc = this.view.freeze?.col ?? 0;
    const fy = fr ? this.rows.pos(fr) * this.zoom : 0;
    const fx = fc ? this.cols.pos(fc) * this.zoom : 0;
    const top = Math.max(this.view.top, fr), left = Math.max(this.view.left, fc);
    const W = this.width - hw - fx, H = this.height - hh - fy;
    const out: Pane[] = [];
    if (fr && fc) out.push({ row: 0, col: 0, x: hw, y: hh, w: fx, h: fy });
    if (fr) out.push({ row: 0, col: left, x: hw + fx, y: hh, w: W, h: fy });
    if (fc) out.push({ row: top, col: 0, x: hw, y: hh + fy, w: fx, h: H });
    out.push({ row: top, col: left, x: hw + fx, y: hh + fy, w: W, h: H });
    return out;
  }

  private paint(): void {
    if (!this.rows || !this.width) return;
    const st: DrawState = {
      sheet: this.sheet, styles: this.wb.styles, rows: this.rows, cols: this.cols, zoom: this.zoom, dpr: this.dpr,
      width: this.width, height: this.height, headerW: this.headerW(), headerH: this.headerH(), panes: this.panes(),
      outlineW: this.outlineW(), outlineH: this.outlineH(),
      selection: this.selection, active: this.active, editing: this.editing ?? undefined, marquee: this.marquee,
      marqueePhase: this.marqueePhase, ghost: this.ghost, insertLine: this.insertLine, refs: this.refHighlights(), showGrid: this.view.showGrid,
      showHeaders: this.view.showHeaders, textCache: this.previewState ? new Map() : this.textCache, cond: this.condEval(), showFormulas: this.host.showFormulas(),
      preview: this.previewState, arrows: this.host.arrows?.(), circles: this.host.circles?.(), pageBreaks: this.host.pageBreaks?.(),
    };
    draw(this.ctx, st);
    if (this.editing && !this.editing.viaBar) this.placeEditor();
    this.placeOptions();
    this.updateWarning();
    this.host.painted?.();
  }

  private condEval(): CondEval | null {
    if (!this.sheet.conditional?.length) return null;
    if (!this.cond || this.condFor !== this.sheet) {
      this.cond = new CondEval(this.wb, this.sheet, this.wb.active);
      this.condFor = this.sheet;
    }
    return this.cond;
  }

  // ----- geometry ----------------------------------------------------------------------------

  /** The cell under a point of the canvas (CSS px), with the pane it is in. */
  hit(x: number, y: number): { area: "corner" | "col" | "row" | "cell" | "none"; row: number; col: number; pane?: Pane } {
    const hw = this.headerW(), hh = this.headerH();
    if (x < hw && y < hh) return { area: "corner", row: 0, col: 0 };
    const panes = this.panes();
    const px = panes.find((p) => x >= p.x && x < p.x + p.w) ?? panes[panes.length - 1];
    const py = panes.find((p) => y >= p.y && y < p.y + p.h) ?? panes[panes.length - 1];
    const col = this.cols.at(this.cols.pos(px.col) + (x - px.x) / this.zoom);
    const row = this.rows.at(this.rows.pos(py.row) + (y - py.y) / this.zoom);
    if (y < hh) return { area: "col", row, col };
    if (x < hw) return { area: "row", row, col };
    const pane = panes.find((p) => x >= p.x && x < p.x + p.w && y >= p.y && y < p.y + p.h);
    return { area: pane ? "cell" : "none", row, col, pane };
  }

  /** A range's rectangle on the canvas (CSS px); null when it is not visible. */
  rectOnCanvas(r: Rect): { x: number; y: number; w: number; h: number } | null {
    const fr = this.view.freeze?.row ?? 0, fc = this.view.freeze?.col ?? 0;
    const panes = this.panes();
    const pane = panes.find((p) => (r.r0 < fr ? p.row === 0 && fr > 0 : p.row !== 0 || fr === 0) && (r.c0 < fc ? p.col === 0 && fc > 0 : p.col !== 0 || fc === 0)) ?? panes[panes.length - 1];
    const x = pane.x + (this.cols.pos(r.c0) - this.cols.pos(pane.col)) * this.zoom;
    const y = pane.y + (this.rows.pos(r.r0) - this.rows.pos(pane.row)) * this.zoom;
    const w = (this.cols.pos(r.c1 + 1) - this.cols.pos(r.c0)) * this.zoom;
    const hh = (this.rows.pos(r.r1 + 1) - this.rows.pos(r.r0)) * this.zoom;
    return { x, y, w, h: hh };
  }

  private visibleRows(): number {
    const p = this.panes().at(-1)!;
    return placed(this.rows, p.row, p.y, p.h, this.zoom).length;
  }

  private visibleCols(): number {
    const p = this.panes().at(-1)!;
    return placed(this.cols, p.col, p.x, p.w, this.zoom).length;
  }

  /** Scrolls so that a cell is fully visible (Excel scrolls by whole rows and columns). */
  reveal(row: number, col: number): void {
    const fr = this.view.freeze?.row ?? 0, fc = this.view.freeze?.col ?? 0;
    const p = this.panes().at(-1)!;
    if (row >= fr) {
      if (row < this.view.top) this.view.top = row;
      else {
        const bottom = this.rows.pos(this.view.top) * this.zoom + p.h;
        if ((this.rows.pos(row + 1)) * this.zoom > bottom) {
          let top = this.view.top;
          while (top < row && (this.rows.pos(row + 1) - this.rows.pos(top)) * this.zoom > p.h) top = this.rows.step(top, 1);
          this.view.top = top;
        }
      }
    }
    if (col >= fc) {
      if (col < this.view.left) this.view.left = col;
      else if ((this.cols.pos(col + 1) - this.cols.pos(this.view.left)) * this.zoom > p.w) {
        let left = this.view.left;
        while (left < col && (this.cols.pos(col + 1) - this.cols.pos(left)) * this.zoom > p.w) left = this.cols.step(left, 1);
        this.view.left = left;
      }
    }
    this.updateScrollbars();
    this.invalidate();
  }

  /** The used range for the scroll bars, cached until cells change (a scan of every cell). */
  private usedCache: { sheet: Sheet; used: Rect | null } | null = null;

  private used(): Rect | null {
    if (!this.usedCache || this.usedCache.sheet !== this.sheet) this.usedCache = { sheet: this.sheet, used: this.sheet.usedRange() };
    return this.usedCache.used;
  }

  private updateScrollbars(): void {
    if (!this.rows) return;
    const used = this.used();
    const vis = Math.max(1, this.visibleRows()), visC = Math.max(1, this.visibleCols());
    const maxRow = Math.min(MAX_ROWS - 1, Math.max(used?.r1 ?? 0, this.view.top + vis * 2, this.active.row + vis));
    const maxCol = Math.min(MAX_COLS - 1, Math.max(used?.c1 ?? 0, this.view.left + visC * 2, this.active.col + visC));
    const vInner = this.vscroll.firstElementChild as HTMLElement;
    const hInner = this.hscroll.firstElementChild as HTMLElement;
    vInner.style.height = `${this.rows.pos(maxRow) * this.zoom + this.vscroll.clientHeight}px`;
    hInner.style.width = `${this.cols.pos(maxCol) * this.zoom + this.hscroll.clientWidth}px`;
    this.suppressScroll = true;
    this.vscroll.scrollTop = this.rows.pos(this.view.top) * this.zoom;
    this.hscroll.scrollLeft = this.cols.pos(this.view.left) * this.zoom;
    requestAnimationFrame(() => { this.suppressScroll = false; });
  }

  // ----- selection -----------------------------------------------------------------------------

  select(rect: Rect, active?: { row: number; col: number }, add = false): void {
    this.tabStart = null;
    const a = active ?? { row: rect.r0, col: rect.c0 };
    const full = expandToMerges(this.sheet, rect);
    if (add) this.view.selection = [...this.selection, full];
    else this.view.selection = [full];
    this.view.active = { ...a };
    this.head = { row: rect.r0 === a.row ? rect.r1 : rect.r0, col: rect.c0 === a.col ? rect.c1 : rect.c0 };
    this.afterSelect();
  }

  /**
   * A click in the Gliederung bar: a level button or a group's −/+ box (`probe` only says whether
   * there is one under the point).
   */
  private outlineClick(x: number, y: number, probe: boolean): boolean {
    const z = this.zoom, ow = this.outlineW(), oh = this.outlineH(), hw = this.headerW(), hh = this.headerH();
    const lane = (v: number) => Math.floor(v / (LANE * z)) + 1;
    const act = (axis: OutlineAxis, fn: () => void) => { if (!probe) this.host.outline?.(axis, fn); return true; };
    // Level buttons: for rows in the band beside the column headers, for columns above the row headers.
    if (x < ow && y >= oh && y < hh) { const n = lane(x); return n <= maxLevel(this.sheet, "rows") + 1 && act("rows", () => showLevel(this.sheet, "rows", n)); }
    if (y < oh && x >= ow && x < hw) { const n = lane(y); return n <= maxLevel(this.sheet, "cols") + 1 && act("cols", () => showLevel(this.sheet, "cols", n)); }
    // A group's box sits at its summary row (column), in the lane of its level.
    if (x < ow && y >= hh) {
      const level = lane(x), row = this.hit(hw + 1, y).row;
      const g = groups(this.sheet, "rows", level).find((r) => r[1] + 1 === row);
      return !!g && act("rows", () => toggle(this.sheet, "rows", g, level));
    }
    if (y < oh && x >= hw) {
      const level = lane(y), col = this.hit(x, hh + 1).col;
      const g = groups(this.sheet, "cols", level).find((c) => c[1] + 1 === col);
      return !!g && act("cols", () => toggle(this.sheet, "cols", g, level));
    }
    return false;
  }

  /** The cell under a point of the canvas (CSS px), for things placed over the grid. */
  hitCell(x: number, y: number): { row: number; col: number } {
    const h = this.hit(Math.max(this.headerW() + 1, x), Math.max(this.headerH() + 1, y));
    return { row: h.row, col: h.col };
  }

  /** Umbruchvorschau: the page break line under the pointer (within 3 px), as "before row/column at". */
  private breakAt(x: number, y: number): { rows: boolean; at: number } | null {
    const pb = this.host.pageBreaks?.();
    if (!pb || x < this.headerW() || y < this.headerH()) return null;
    for (const pg of pb.pages) {
      if (pg.rows[1] < pb.area.r1) {
        const b = this.rectOnCanvas({ r0: pg.rows[1], c0: pg.cols[0], r1: pg.rows[1], c1: pg.cols[1] });
        if (b && Math.abs(y - (b.y + b.h)) <= 3 && x >= b.x && x <= b.x + b.w) return { rows: true, at: pg.rows[1] + 1 };
      }
      if (pg.cols[1] < pb.area.c1) {
        const b = this.rectOnCanvas({ r0: pg.rows[0], c0: pg.cols[1], r1: pg.rows[1], c1: pg.cols[1] });
        if (b && Math.abs(x - (b.x + b.w)) <= 3 && y >= b.y && y <= b.y + b.h) return { rows: false, at: pg.cols[1] + 1 };
      }
    }
    return null;
  }

  /** Tab on a protected sheet: to the next unlocked cell (Excel); false where the sheet is not protected. */
  private protectedTab(back: boolean): boolean {
    const n = this.host.nextUnlocked?.(this.active.row, this.active.col, back);
    if (!n) return false;
    this.selectCell(n.row, n.col);
    return true;
  }

  /** Several ranges at once (Inhalte auswählen, Suchen und Auswählen); the first holds the active cell. */
  selectMany(ranges: Rect[]): void {
    if (!ranges.length) return;
    this.tabStart = null;
    this.view.selection = ranges.map((r) => ({ ...r }));
    this.view.active = { row: ranges[0].r0, col: ranges[0].c0 };
    this.head = { ...this.view.active };
    this.reveal(ranges[0].r0, ranges[0].c0);
    this.afterSelect();
  }

  selectCell(row: number, col: number): void {
    const m = this.sheet.mergeAt(row, col);
    this.select(m ?? { r0: row, c0: col, r1: row, c1: col }, { row: m?.r0 ?? row, col: m?.c0 ?? col });
    this.head = { row, col };
    this.reveal(row, col);
  }

  private warnBtn: HTMLElement | null = null;

  /** The warning button left of the active cell when it shows an error indicator (Excel). */
  private updateWarning(): void {
    const c = this.sheet.get(this.active.row, this.active.col);
    const kind = c && !this.editing ? errorIndicator(c) : null;
    if (!kind) { this.warnBtn?.remove(); this.warnBtn = null; return; }
    if (!this.warnBtn) {
      this.warnBtn = h("button", { class: "warn-btn", "aria-label": "Fehlerüberprüfung" }, "!", h("span", { class: "dd" }, "⌄"));
      this.warnBtn.addEventListener("mousedown", (e) => {
        e.preventDefault();
        const k = errorIndicator(this.sheet.get(this.active.row, this.active.col) ?? { v: null });
        if (k) this.host.errorMenu(k, this.warnBtn!.getBoundingClientRect());
      });
      this.el.append(this.warnBtn);
    }
    this.warnBtn.dataset.tip = kind === "text-number" ? "Die Zahl in dieser Zelle ist als Text formatiert oder ihr ist ein Apostroph vorangestellt." : "Die Formel in dieser Zelle ergibt einen Fehler.";
    const b = this.rectOnCanvas({ r0: this.active.row, c0: this.active.col, r1: this.active.row, c1: this.active.col });
    if (b) { this.warnBtn.style.left = `${b.x - 30}px`; this.warnBtn.style.top = `${b.y}px`; }
  }

  private afterSelect(): void {
    this.updateWarning();
    this.syncBar();
    this.host.selectionChanged();
    this.invalidate();
  }

  /** Extends the last range from the active cell to `to` (Shift+arrow, Shift+click). */
  private extendTo(to: { row: number; col: number }): void {
    this.head = to;
    const r = expandToMerges(this.sheet, normRect(this.active.row, this.active.col, to.row, to.col));
    this.view.selection = [...this.selection.slice(0, -1), r];
    // A selection over whole columns (rows) does not scroll along them (recorded in Excel).
    this.reveal(r.r0 === 0 && r.r1 === MAX_ROWS - 1 ? -1 : to.row, r.c0 === 0 && r.c1 === MAX_COLS - 1 ? -1 : to.col);
    this.afterSelect();
  }

  private isEmpty(row: number, col: number): boolean {
    const c = this.sheet.get(row, col);
    return !c || c.v === null || c.v === "";
  }

  /** Strg+arrow: to the edge of the data block, or to the next block, or to the sheet's edge. */
  private jump(from: { row: number; col: number }, dr: number, dc: number): { row: number; col: number } {
    const maxR = MAX_ROWS - 1, maxC = MAX_COLS - 1;
    const step = (p: { row: number; col: number }) => ({ row: p.row + dr, col: p.col + dc });
    const inside = (p: { row: number; col: number }) => p.row >= 0 && p.row <= maxR && p.col >= 0 && p.col <= maxC;
    let p = from;
    const next = step(p);
    if (!inside(next)) return p;
    const used = this.used();
    const beyond = (q: { row: number; col: number }) => !used || (dr > 0 && q.row > used.r1) || (dc > 0 && q.col > used.c1);
    if (!this.isEmpty(p.row, p.col) && !this.isEmpty(next.row, next.col)) {
      while (inside(step(p)) && !this.isEmpty(step(p).row, step(p).col)) p = step(p);
      return p;
    }
    p = next;
    while (inside(step(p)) && this.isEmpty(p.row, p.col)) {
      if (beyond(p)) return { row: dr > 0 ? maxR : dr < 0 ? 0 : p.row, col: dc > 0 ? maxC : dc < 0 ? 0 : p.col };
      p = step(p);
    }
    return p;
  }

  /** The block of data around a cell bounded by empty rows and columns (Strg+A, AutoSumme). */
  currentRegion(row: number, col: number): Rect {
    let r = { r0: row, c0: col, r1: row, c1: col };
    let grown = true;
    const any = (r0: number, c0: number, r1: number, c1: number) => {
      for (let i = Math.max(0, r0); i <= Math.min(MAX_ROWS - 1, r1); i++) for (let j = Math.max(0, c0); j <= Math.min(MAX_COLS - 1, c1); j++) if (!this.isEmpty(i, j)) return true;
      return false;
    };
    while (grown) {
      grown = false;
      if (r.r0 > 0 && any(r.r0 - 1, r.c0 - 1, r.r0 - 1, r.c1 + 1)) { r.r0--; grown = true; }
      if (r.r1 < MAX_ROWS - 1 && any(r.r1 + 1, r.c0 - 1, r.r1 + 1, r.c1 + 1)) { r.r1++; grown = true; }
      if (r.c0 > 0 && any(r.r0 - 1, r.c0 - 1, r.r1 + 1, r.c0 - 1)) { r.c0--; grown = true; }
      if (r.c1 < MAX_COLS - 1 && any(r.r0 - 1, r.c1 + 1, r.r1 + 1, r.c1 + 1)) { r.c1++; grown = true; }
      if (r.r1 - r.r0 > 100000) break;
    }
    return r;
  }

  /** Excel's keyboard modes: Ende (the next arrow jumps), F8 Auswahl erweitern, Umschalt+F8 Abschnitt hinzufügen. */
  private keyMode: { end: boolean; select: "extend" | "add" | null } = { end: false, select: null };

  /** Leaves Ende, Auswahl erweitern and Abschnitt hinzufügen (a fresh workbook starts without them). */
  resetKeyModes(): void {
    this.setKeyMode(false, null);
  }

  private setKeyMode(end: boolean, select: "extend" | "add" | null): void {
    if (this.keyMode.end === end && this.keyMode.select === select) return;
    this.keyMode = { end, select };
    const flags: [string, string][] = [];
    if (select === "extend") flags.push(["Auswahlmodus", "Auswahl erweitern"]);
    if (select === "add") flags.push(["Auswahlmodus", "Abschnitt hinzufügen oder entfernen"]);
    if (end) flags.push(["Endmodus", "Endmodus"]);
    this.host.flagsChanged?.(flags);
  }

  private move(dr: number, dc: number, extend: boolean, ctrl: boolean): void {
    const from = extend ? this.head : this.active;
    let to: { row: number; col: number };
    if (ctrl) to = this.jump(from, dr, dc);
    else {
      to = { row: from.row, col: from.col };
      if (dr) to.row = this.rows.step(from.row, dr > 0 ? 1 : -1);
      if (dc) to.col = this.cols.step(from.col, dc > 0 ? 1 : -1);
      // Leaving a merged area moves past it.
      const m = this.sheet.mergeAt(from.row, from.col);
      if (m && !extend) {
        if (dr > 0) to.row = Math.min(MAX_ROWS - 1, m.r1 + 1);
        if (dc > 0) to.col = Math.min(MAX_COLS - 1, m.c1 + 1);
        if (dr < 0) to.row = Math.max(0, m.r0 - 1);
        if (dc < 0) to.col = Math.max(0, m.c0 - 1);
      }
    }
    if (extend) this.extendTo(to);
    else this.selectCell(to.row, to.col);
  }

  /** Enter/Tab inside a selection move the active cell through it (Excel). */
  private moveWithin(dr: number, dc: number): void {
    const sel = this.selection;
    const single = sel.length === 1 && sel[0].r0 === sel[0].r1 && sel[0].c0 === sel[0].c1;
    const m = this.sheet.mergeAt(this.active.row, this.active.col);
    const isMergeOnly = sel.length === 1 && m && m.r0 === sel[0].r0 && m.r1 === sel[0].r1 && m.c0 === sel[0].c0 && m.c1 === sel[0].c1;
    if (single || isMergeOnly) {
      // Tab remembers the column it started in; Enter returns there (Excel).
      if (dc !== 0) {
        const start = this.tabStart ?? this.active.col;
        this.move(dr, dc, false, false);
        this.tabStart = start;
        return;
      }
      if (dr > 0 && this.tabStart !== null) {
        const col = this.tabStart;
        const row = m ? m.r1 + 1 : this.rows.step(this.active.row, 1);
        this.selectCell(row, col);
        return;
      }
      this.move(dr, dc, false, false);
      return;
    }
    let { row, col } = this.active;
    let ri = sel.length - 1;
    for (let i = sel.length - 1; i >= 0; i--) if (row >= sel[i].r0 && row <= sel[i].r1 && col >= sel[i].c0 && col <= sel[i].c1) { ri = i; break; }
    // Past the area's last cell the next area begins, before its first the previous one ends (Excel).
    const next = () => { ri = (ri + 1) % sel.length; row = sel[ri].r0; col = sel[ri].c0; };
    const prev = () => { ri = (ri - 1 + sel.length) % sel.length; row = sel[ri].r1; col = sel[ri].c1; };
    const advance = () => {
      const r = sel[ri];
      if (dr) {
        row += dr;
        if (row > r.r1) { row = r.r0; col++; if (col > r.c1) next(); }
        else if (row < r.r0) { row = r.r1; col--; if (col < r.c0) prev(); }
      } else {
        col += dc;
        if (col > r.c1) { col = r.c0; row++; if (row > r.r1) next(); }
        else if (col < r.c0) { col = r.c1; row--; if (row < r.r0) prev(); }
      }
    };
    advance();
    const mm = this.sheet.mergeAt(row, col);
    if (mm && (mm.r0 !== row || mm.c0 !== col)) advance();
    this.view.active = { row, col };
    this.reveal(row, col);
    this.afterSelect();
  }

  // ----- editing --------------------------------------------------------------------------------

  /** Eingeben (typing replaces) or Bearbeiten (F2, double-click: the content stays). */
  startEdit(mode: "enter" | "edit", text?: string, caret?: number, viaBar = false): void {
    if (this.editing) return;
    this.showOptions(null);
    const { row, col } = this.active;
    if (this.host.canEdit && !this.host.canEdit(row, col)) return;
    const cell = this.sheet.get(row, col);
    const style = this.wb.styles.get(cell?.s ?? this.sheet.rows.get(row)?.s ?? this.sheet.cols.get(col)?.s);
    const original = editText(cell, style);
    this.editing = { row, col, original, viaBar };
    this.reveal(row, col);
    const value = text ?? (mode === "edit" ? original : "");
    this.editor.value = value;
    this.fx.value = value;
    this.setMode(mode);
    this.editorStyle(style);
    this.editor.classList.add("on");
    this.placeEditor();
    const target = viaBar ? this.fx : this.editor;
    target.focus({ preventScroll: true });
    const pos = caret ?? value.length;
    target.setSelectionRange(pos, pos);
    this.invalidate();
  }

  private setMode(mode: Mode): void {
    if (this.mode === mode) return;
    this.mode = mode;
    this.host.modeChanged(mode);
  }

  private editorStyle(style: ReturnType<Workbook["styles"]["get"]>): void {
    const f = style.font;
    const e = this.editor.style;
    e.font = `${f.italic ? "italic " : ""}${f.bold ? "bold " : ""}${(f.size * PT * this.zoom).toFixed(3)}px "${f.name}", Carlito, sans-serif`;
    e.color = f.color ?? "#000";
    e.textDecoration = [f.underline ? "underline" : "", f.strike ? "line-through" : ""].join(" ").trim() || "none";
    e.background = style.fill ?? "#fff";
    this.editor.dataset.align = style.h === "right" ? "right" : style.h === "center" ? "center" : "left";
    this.editor.dataset.wrap = style.wrap ? "1" : "";
  }

  /** The editor covers the cell and grows to the right as the text gets longer (Excel). */
  private placeEditor(): void {
    if (!this.editing) return;
    const { row, col } = this.editing;
    const m = this.sheet.mergeAt(row, col) ?? { r0: row, c0: col, r1: row, c1: col };
    const b = this.rectOnCanvas(m);
    if (!b) return;
    const e = this.editor;
    const lines = e.value.split("\n").length;
    const ctx = this.ctx;
    ctx.font = e.style.font;
    const textW = Math.max(...e.value.split("\n").map((l) => ctx.measureText(l).width)) + 8 * this.zoom;
    let w = b.w, c = m.c1;
    while (w < textW && c < MAX_COLS - 1 && b.x + w < this.width) { c++; w += this.cols.size(c) * this.zoom; }
    const lineH = parseFloat(e.style.font) * 1.22;
    const hgt = Math.max(b.h, lines * lineH + 2);
    e.style.left = `${b.x}px`;
    e.style.top = `${b.y}px`;
    e.style.width = `${Math.max(b.w, w) - 1}px`;
    e.style.height = `${hgt - 1}px`;
    e.style.lineHeight = `${lineH}px`;
    e.style.paddingTop = `${Math.max(0, hgt - lines * lineH - 2)}px`;
  }

  private stopEdit(): void {
    this.editing = null;
    this.pointRef = null;
    this.editor.classList.remove("on");
    this.editor.value = "";
    // The hidden editor keeps the keyboard focus but must not cover the cell it edited.
    this.editor.style.left = "-10000px";
    this.editor.style.top = "0";
    this.ac.hidden = true;
    this.setMode("ready");
    this.syncBar();
    this.invalidate();
  }

  cancelEdit(): void {
    if (!this.editing) return;
    this.stopEdit();
    this.editor.focus({ preventScroll: true });
  }

  /** Commits the edit, then moves as Enter/Tab/arrows do ("none": stays, as the ✓ button). */
  commitEdit(moveBy: "down" | "up" | "right" | "left" | "none", opts: { all?: boolean; array?: boolean; arrow?: boolean } = {}): void {
    if (!this.editing) return;
    let text = (this.editing.viaBar ? this.fx : this.editor).value;
    // Excel closes brackets a formula left open.
    if (text.startsWith("=")) {
      let open = 0;
      let quoted = false;
      for (const ch of text) { if (ch === '"') quoted = !quoted; else if (!quoted && ch === "(") open++; else if (!quoted && ch === ")") open--; }
      if (open > 0) text += ")".repeat(open);
    }
    const { row, col, viaBar } = this.editing;
    const changed = text !== this.editing.original;
    // Datenüberprüfung: the entry is checked first; the editor stays while the box is open (Excel).
    if (changed && !opts.array && this.host.validate) {
      const v = this.host.validate(row, col, text);
      if (v !== "ok") {
        void v.then((r) => {
          if (r === "ok") this.finishCommit(row, col, text, changed, moveBy, opts);
          else if (r === "retry") {
            this.stopEdit();
            this.startEdit("edit", text, undefined, viaBar);
            (viaBar ? this.fx : this.editor).select();
          } else {
            this.stopEdit();
            this.editor.focus({ preventScroll: true });
            this.afterSelect();
          }
        });
        return;
      }
    }
    this.finishCommit(row, col, text, changed, moveBy, opts);
  }

  private finishCommit(row: number, col: number, text: string, changed: boolean, moveBy: "down" | "up" | "right" | "left" | "none", opts: { all?: boolean; array?: boolean; arrow?: boolean }): void {
    this.stopEdit();
    const targets = opts.all ? this.selection : [{ r0: row, c0: col, r1: row, c1: col }];
    if (changed || opts.all) for (const t of targets) this.host.commit(t, text, { array: opts.array });
    this.editor.focus({ preventScroll: true });
    if (opts.arrow) {
      const [dr, dc] = moveBy === "down" ? [1, 0] : moveBy === "up" ? [-1, 0] : moveBy === "right" ? [0, 1] : [0, -1];
      this.move(dr, dc, false, false);
    } else if (moveBy === "down") this.moveWithin(1, 0);
    else if (moveBy === "up") this.moveWithin(-1, 0);
    else if (moveBy === "right") this.moveWithin(0, 1);
    else if (moveBy === "left") this.moveWithin(0, -1);
    else this.afterSelect();
  }

  private startFormulaFromButton(): void {
    if (!this.editing) this.startEdit("enter", "=", 1, true);
    else if (!this.fx.value.startsWith("=")) { this.fx.value = "=" + this.fx.value; this.editor.value = this.fx.value; }
  }

  /** Namenfeld and Bearbeitungsleiste show the active cell. */
  syncBar(): void {
    if (this.editing) return;
    const sel = this.selection.at(-1)!;
    const multi = this.drag?.kind === "cells" && (sel.r0 !== sel.r1 || sel.c0 !== sel.c1);
    this.nameBox.value = multi ? `${sel.r1 - sel.r0 + 1}Z x ${sel.c1 - sel.c0 + 1}S` : cellName(this.active.row, this.active.col);
    const cell = this.sheet.get(this.active.row, this.active.col);
    const style = this.wb.styles.get(cell?.s ?? this.sheet.rows.get(this.active.row)?.s ?? this.sheet.cols.get(this.active.col)?.s);
    this.fx.value = editText(cell, style);
  }

  // ----- formulas: Zeigen and coloured references ---------------------------------------------

  private editorText(): string {
    return this.editing ? (this.editing.viaBar ? this.fx : this.editor).value : "";
  }

  private refHighlights(): { rect: Rect; color: string }[] | undefined {
    const t = this.editorText();
    if (!t.startsWith("=")) return undefined;
    try {
      const refs = formulaRefs(t.slice(1)).filter((r) => !r.sheet || r.sheet.toLowerCase() === this.sheet.name.toLowerCase());
      return refs.map((r, i) => ({ rect: r.ref, color: REF_COLORS[i % REF_COLORS.length] }));
    } catch {
      return undefined;
    }
  }

  /** Whether arrow keys and clicks point at cells (Excel: after an operator, "(" or ";"). */
  private canPoint(): boolean {
    if (!this.editing || this.mode === "edit") return false;
    const input = this.editing.viaBar ? this.fx : this.editor;
    const t = input.value;
    if (!t.startsWith("=") && !t.startsWith("+")) return false;
    const pos = input.selectionStart ?? t.length;
    if (input.selectionEnd !== pos && !this.pointRef) return false;
    if (this.pointRef && this.pointRef.end === pos) return true;
    if (pos !== t.length && !/^\s*[;)]/.test(t.slice(pos))) return false;
    return /[=(;+\-*/^&<>:,]\s*$/.test(t.slice(0, pos));
  }

  /** AutoSumme: the formula is typed with its range marked as pointed (Zeigen), as Excel does. */
  startPointing(text: string, start: number, end: number, ref: Rect | null): void {
    this.startEdit("enter", text, end);
    if (!ref) return;
    this.editor.setSelectionRange(start, end);
    this.pointRef = { start, end, anchor: { row: ref.r0, col: ref.c0 }, head: { row: ref.r1, col: ref.c1 } };
    this.setMode("point");
    this.invalidate();
  }

  /** Puts a reference to the pointed cell or range into the formula. */
  private point(anchor: { row: number; col: number }, head: { row: number; col: number }, sheetPrefix = ""): void {
    const input = this.editing!.viaBar ? this.fx : this.editor;
    const t = input.value;
    const start = this.pointRef?.start ?? input.selectionStart ?? t.length;
    const end = this.pointRef?.end ?? start;
    const ref = sheetPrefix + (anchor.row === head.row && anchor.col === head.col ? cellName(anchor.row, anchor.col) : rangeName(normRect(anchor.row, anchor.col, head.row, head.col)));
    input.value = t.slice(0, start) + ref + t.slice(end);
    const pos = start + ref.length;
    input.setSelectionRange(pos, pos);
    this.pointRef = { start, end: pos, anchor, head };
    (input === this.fx ? this.editor : this.fx).value = input.value;
    this.setMode("point");
    this.reveal(head.row, head.col);
    this.placeEditor();
    this.invalidate();
  }

  // ----- AutoVervollständigen -----------------------------------------------------------------

  /** Completes text typed into a cell from the column above and below (Excel's AutoVervollständigen). */
  private autoComplete(): void {
    if (!this.editing || this.mode !== "enter") return;
    const e = this.editor;
    const t = e.value;
    if (!t || t.startsWith("=") || /^[\d,.\-+ ]+$/.test(t) || e.selectionStart !== t.length) return;
    const { row, col } = this.editing;
    const seen = new Set<string>();
    const scan = (dir: 1 | -1) => {
      for (let r = row + dir; r >= 0 && r < MAX_ROWS; r += dir) {
        const c = this.sheet.get(r, col);
        if (!c || c.v === null || c.v === "") break;
        if (typeof c.v === "string" && c.f === undefined) seen.add(c.v);
      }
    };
    scan(-1);
    scan(1);
    const low = t.toLowerCase();
    const matches = [...seen].filter((s) => s.toLowerCase().startsWith(low) && s.length > t.length);
    const distinct = new Set(matches.map((m) => m.toLowerCase()));
    if (distinct.size !== 1) return;
    // The completed text takes the stored entry's spelling, typed letters included (Excel).
    const m = matches[0];
    e.value = m;
    e.setSelectionRange(t.length, e.value.length);
    this.fx.value = e.value;
  }

  // ----- events ---------------------------------------------------------------------------------

  private installEvents(): void {
    this.vscroll.addEventListener("scroll", () => {
      if (this.suppressScroll) return;
      this.view.top = Math.max(this.view.freeze?.row ?? 0, this.rows.at(this.vscroll.scrollTop / this.zoom + 0.5));
      this.invalidate();
      this.maybeGrowScroll();
    });
    this.hscroll.addEventListener("scroll", () => {
      if (this.suppressScroll) return;
      this.view.left = Math.max(this.view.freeze?.col ?? 0, this.cols.at(this.hscroll.scrollLeft / this.zoom + 0.5));
      this.invalidate();
      this.maybeGrowScroll();
    });
    this.el.addEventListener("wheel", (e) => {
      if (e.ctrlKey) return;
      e.preventDefault();
      const lines = Math.sign(e.deltaY) * Math.max(1, Math.round(Math.abs(e.deltaY) / 100 * 3));
      // Umschalt+wheel does not scroll in Excel (recorded); a sideways wheel or touchpad does.
      if (e.shiftKey && !e.deltaX) return;
      if (e.deltaX) {
        const n = Math.sign(e.deltaX) * 2;
        let left = this.view.left;
        for (let i = 0; i < Math.abs(n); i++) left = this.cols.step(left, n > 0 ? 1 : -1);
        this.view.left = Math.max(this.view.freeze?.col ?? 0, left);
      } else {
        let top = this.view.top;
        for (let i = 0; i < Math.abs(lines); i++) top = this.rows.step(top, lines > 0 ? 1 : -1);
        this.view.top = Math.max(this.view.freeze?.row ?? 0, top);
      }
      this.updateScrollbars();
      this.invalidate();
    }, { passive: false });

    this.canvas.addEventListener("pointerdown", (e) => this.onDown(e));
    this.canvas.addEventListener("pointermove", (e) => this.onMove(e));
    this.canvas.addEventListener("pointerup", (e) => this.onUp(e));
    this.canvas.addEventListener("dblclick", (e) => this.onDouble(e));
    this.canvas.addEventListener("contextmenu", (e) => {
      e.preventDefault();
      if (this.rightDragEnded) { this.rightDragEnded = false; return; }
      const p = this.local(e);
      const hit = this.hit(p.x, p.y);
      if (hit.area === "cell" && !this.selection.some((r) => inRect(r, hit.row, hit.col))) this.selectCell(hit.row, hit.col);
      // A header outside the selected whole columns (rows) selects its column (row), as a click does (recorded).
      if (hit.area === "col" && !this.selection.some((r) => r.r0 === 0 && r.r1 === MAX_ROWS - 1 && hit.col >= r.c0 && hit.col <= r.c1))
        this.select({ r0: 0, c0: hit.col, r1: MAX_ROWS - 1, c1: hit.col }, { row: this.view.top, col: hit.col });
      if (hit.area === "row" && !this.selection.some((r) => r.c0 === 0 && r.c1 === MAX_COLS - 1 && hit.row >= r.r0 && hit.row <= r.r1))
        this.select({ r0: hit.row, c0: 0, r1: hit.row, c1: MAX_COLS - 1 }, { row: hit.row, col: this.view.left });
      this.host.contextMenu(e, hit.area === "col" ? "col" : hit.area === "row" ? "row" : "cell");
    });

    this.editor.addEventListener("keydown", (e) => this.onKey(e, this.editor));
    this.fx.addEventListener("keydown", (e) => this.onKey(e, this.fx));
    this.editor.addEventListener("input", (e) => {
      if (!this.editing) {
        // A character typed on the grid starts Eingeben with it (Excel).
        const t = this.editor.value;
        this.editor.value = "";
        this.startEdit("enter", t);
        this.autoComplete();
        this.functionHint();
        return;
      }
      this.pointRef = null;
      if (this.mode === "point") this.setMode("enter");
      this.fx.value = this.editor.value;
      if ((e as InputEvent).inputType === "insertText") this.autoComplete();
      this.functionHint();
      this.placeEditor();
      this.invalidate();
    });
    this.fx.addEventListener("focus", () => {
      if (!this.editing) this.startEdit("edit", undefined, undefined, true);
      else this.editing.viaBar = true;
    });
    this.fx.addEventListener("pointerdown", () => {
      if (this.editing && !this.editing.viaBar) {
        this.editing.viaBar = true;
        this.setMode("edit");
      }
    });
    this.fx.addEventListener("input", () => {
      this.pointRef = null;
      this.editor.value = this.fx.value;
      this.functionHint();
      this.placeEditor();
      this.invalidate();
    });
    this.nameBox.addEventListener("focus", () => this.nameBox.select());
    this.nameBox.addEventListener("keydown", (e) => {
      if (e.key === "Enter") {
        e.preventDefault();
        this.goTo(this.nameBox.value);
        this.editor.focus({ preventScroll: true });
      } else if (e.key === "Escape") {
        this.syncBar();
        this.editor.focus({ preventScroll: true });
      }
    });
  }

  /** Namenfeld and Gehe zu: a cell, a range or a defined name. */
  goTo(text: string): boolean {
    const t = text.trim();
    const name = this.wb.names.find((n) => n.name.toLowerCase() === t.toLowerCase());
    let target = t;
    if (name) {
      const m = /^(?:'?([^'!]+)'?!)?(.+)$/.exec(name.ref);
      if (m?.[1]) {
        const i = this.wb.sheets.findIndex((s) => s.name.toLowerCase() === m[1].toLowerCase());
        if (i >= 0 && i !== this.wb.active) { this.wb.active = i; this.refreshLayout(); }
      }
      target = m?.[2] ?? t;
    }
    const r = parseRange(target.replace(/\$/g, ""));
    if (!r) {
      this.syncBar();
      return false;
    }
    this.select(r, { row: r.r0, col: r.c0 });
    this.reveal(r.r0, r.c0);
    return true;
  }

  private maybeGrowScroll(): void {
    const vInner = this.vscroll.firstElementChild as HTMLElement;
    if (this.vscroll.scrollTop + this.vscroll.clientHeight * 1.5 > vInner.offsetHeight) this.updateScrollbars();
  }

  private local(e: MouseEvent): { x: number; y: number } {
    const r = this.canvas.getBoundingClientRect();
    return { x: e.clientX - r.left, y: e.clientY - r.top };
  }

  /** What the pointer is over: a header edge, the fill handle, the selection's border or a cell. */
  private target(x: number, y: number): { kind: "resizeCol" | "resizeRow" | "fill" | "move" | "cells" | "cols" | "rows" | "corner" | "outline" | "pageBreak" | "none"; row: number; col: number } {
    if ((x < this.outlineW() || y < this.outlineH()) && this.outlineClick(x, y, true)) return { kind: "outline", row: 0, col: 0 };
    const pb = this.breakAt(x, y);
    if (pb) return { kind: "pageBreak", row: pb.rows ? pb.at : -1, col: pb.rows ? -1 : pb.at };
    if (x < this.outlineW() || y < this.outlineH()) return { kind: "none", row: 0, col: 0 };
    const hit = this.hit(x, y);
    const z = this.zoom;
    if (hit.area === "col") {
      const b = this.rectOnCanvas({ r0: 0, c0: hit.col, r1: 0, c1: hit.col })!;
      if (Math.abs(x - (b.x + b.w)) <= 3) return { kind: "resizeCol", row: 0, col: hit.col };
      if (Math.abs(x - b.x) <= 3 && hit.col > 0) return { kind: "resizeCol", row: 0, col: this.cols.step(hit.col, -1) };
      return { kind: "cols", row: 0, col: hit.col };
    }
    if (hit.area === "row") {
      const b = this.rectOnCanvas({ r0: hit.row, c0: 0, r1: hit.row, c1: 0 })!;
      if (Math.abs(y - (b.y + b.h)) <= 2) return { kind: "resizeRow", row: hit.row, col: 0 };
      if (Math.abs(y - b.y) <= 2 && hit.row > 0) return { kind: "resizeRow", row: this.rows.step(hit.row, -1), col: 0 };
      return { kind: "rows", row: hit.row, col: 0 };
    }
    if (hit.area === "corner") return { kind: "corner", row: 0, col: 0 };
    if (hit.area !== "cell") return { kind: "none", row: hit.row, col: hit.col };
    if (this.selection.length === 1 && !this.editing) {
      const sel = this.selection[0];
      const b = this.rectOnCanvas(sel);
      if (b) {
        const hx = b.x + b.w, hy = b.y + b.h;
        if (Math.abs(x - hx) <= 4 * Math.max(1, z) && Math.abs(y - hy) <= 4 * Math.max(1, z)) return { kind: "fill", row: hit.row, col: hit.col };
        const nearV = (Math.abs(x - b.x) <= 3 || Math.abs(x - hx) <= 3) && y >= b.y - 3 && y <= hy + 3;
        const nearH = (Math.abs(y - b.y) <= 3 || Math.abs(y - hy) <= 3) && x >= b.x - 3 && x <= hx + 3;
        if (nearV || nearH) return { kind: "move", row: hit.row, col: hit.col };
      }
    }
    return { kind: "cells", row: hit.row, col: hit.col };
  }

  /** Where the last press was: a drag is never a double-click (Excel). */
  private downAt = { x: 0, y: 0, moved: false };

  /** A finger on the sheet: dragging scrolls, a tap selects (Excel's touch mode). */
  private touch: { x: number; y: number; top: number; left: number; panned: boolean; id: number } | null = null;

  private onTouchDown(e: PointerEvent): void {
    this.touch = { x: e.clientX, y: e.clientY, top: this.view.top, left: this.view.left, panned: false, id: e.pointerId };
    this.canvas.setPointerCapture(e.pointerId);
  }

  private onTouchMove(e: PointerEvent): void {
    const t = this.touch!;
    const dx = e.clientX - t.x, dy = e.clientY - t.y;
    if (!t.panned && Math.hypot(dx, dy) < 8) return;
    t.panned = true;
    const fr = this.view.freeze?.row ?? 0, fc = this.view.freeze?.col ?? 0;
    this.view.top = Math.max(fr, this.rows.at(this.rows.pos(t.top) - dy / this.zoom));
    this.view.left = Math.max(fc, this.cols.at(this.cols.pos(t.left) - dx / this.zoom));
    this.updateScrollbars();
    this.invalidate();
  }

  private onTouchUp(e: PointerEvent): void {
    const t = this.touch!;
    this.touch = null;
    if (t.panned) return;
    const p = this.local(e);
    const hit = this.hit(p.x, p.y);
    if (hit.area === "cell") {
      if (this.editing) this.commitEdit("none");
      this.selectCell(hit.row, hit.col);
    }
  }

  private onDown(e: PointerEvent): void {
    if (e.pointerType === "touch") { this.onTouchDown(e); return; }
    this.downAt = { x: e.clientX, y: e.clientY, moved: false };
    if (e.button !== 0) {
      if (e.button === 2) {
        this.editor.focus({ preventScroll: true });
        // The right button on the fill handle or the selection's border drags too; the menu
        // comes when it is released (Excel).
        const p = this.local(e);
        const t = this.target(p.x, p.y);
        if (!this.editing && (t.kind === "fill" || t.kind === "move") && this.host.rightDrag) {
          this.canvas.setPointerCapture(e.pointerId);
          this.drag = { kind: t.kind, start: { row: t.row, col: t.col }, x: p.x, y: p.y, extra: "right" };
        }
      }
      return;
    }
    const p = this.local(e);
    const t = this.target(p.x, p.y);
    if (this.editing && t.kind === "cells" && this.canPoint() || (this.pointRef && t.kind === "cells")) {
      e.preventDefault();
      this.canvas.setPointerCapture(e.pointerId);
      if (e.shiftKey && this.pointRef) this.point(this.pointRef.anchor, { row: t.row, col: t.col });
      else this.point({ row: t.row, col: t.col }, { row: t.row, col: t.col });
      this.drag = { kind: "point", start: { row: t.row, col: t.col }, x: p.x, y: p.y };
      return;
    }
    if (this.editing) this.commitEdit("none");
    e.preventDefault();
    this.editor.focus({ preventScroll: true });
    const af = this.sheet.autoFilter;
    if (af && t.kind === "cells" && t.row === af.ref.r0 && t.col >= af.ref.c0 && t.col <= af.ref.c1) {
      const b = this.rectOnCanvas({ r0: t.row, c0: t.col, r1: t.row, c1: t.col })!;
      const fb = filterButton(b.x, b.w, b.y, b.h, this.zoom);
      if (p.x >= fb.x && p.x <= fb.x + fb.w && p.y >= fb.y && p.y <= fb.y + fb.h) {
        this.selectCell(t.row, t.col);
        const c = this.canvas.getBoundingClientRect();
        this.host.filterMenu(t.col, new DOMRect(c.left + fb.x, c.top + fb.y, fb.w, fb.h));
        return;
      }
    }
    if (t.kind === "outline") { this.outlineClick(p.x, p.y, false); return; }
    if (t.kind === "pageBreak") {
      this.canvas.setPointerCapture(e.pointerId);
      this.drag = { kind: "pageBreak", start: { row: t.row, col: t.col }, x: p.x, y: p.y };
      return;
    }
    this.canvas.setPointerCapture(e.pointerId);
    this.tabStart = null;
    const add = e.ctrlKey;
    switch (t.kind) {
      case "corner":
        this.select({ r0: 0, c0: 0, r1: MAX_ROWS - 1, c1: MAX_COLS - 1 }, { row: this.view.top, col: this.view.left });
        return;
      case "resizeCol":
      case "resizeRow":
        this.liveBefore = [];
        this.drag = { kind: t.kind, start: { row: t.row, col: t.col }, x: p.x, y: p.y, extra: t.kind === "resizeCol" ? this.cols.size(t.col) : this.rows.size(t.row) };
        return;
      case "fill":
      case "move":
        this.drag = { kind: t.kind, start: { row: t.row, col: t.col }, x: p.x, y: p.y };
        return;
      case "cols":
        if (e.shiftKey) this.extendTo({ row: MAX_ROWS - 1, col: t.col }), this.fullCols(this.active.col, t.col);
        else this.select({ r0: 0, c0: t.col, r1: MAX_ROWS - 1, c1: t.col }, { row: this.view.top, col: t.col }, add);
        this.drag = { kind: "cols", start: { row: 0, col: e.shiftKey ? this.active.col : t.col }, x: p.x, y: p.y };
        return;
      case "rows":
        if (e.shiftKey) this.fullRows(this.active.row, t.row);
        else this.select({ r0: t.row, c0: 0, r1: t.row, c1: MAX_COLS - 1 }, { row: t.row, col: this.view.left }, add);
        this.drag = { kind: "rows", start: { row: e.shiftKey ? this.active.row : t.row, col: 0 }, x: p.x, y: p.y };
        return;
      case "cells":
        if (e.shiftKey || this.keyMode.select === "extend") this.extendTo({ row: t.row, col: t.col });
        else if (add || this.keyMode.select === "add") this.select({ r0: t.row, c0: t.col, r1: t.row, c1: t.col }, { row: t.row, col: t.col }, true);
        else this.selectCell(t.row, t.col);
        this.drag = { kind: "cells", start: { ...this.active }, x: p.x, y: p.y };
        return;
    }
  }

  private fullCols(a: number, b: number): void {
    const r = { r0: 0, c0: Math.min(a, b), r1: MAX_ROWS - 1, c1: Math.max(a, b) };
    this.view.selection = [...this.selection.slice(0, -1), r];
    this.afterSelect();
  }

  private fullRows(a: number, b: number): void {
    const r = { r0: Math.min(a, b), c0: 0, r1: Math.max(a, b), c1: MAX_COLS - 1 };
    this.view.selection = [...this.selection.slice(0, -1), r];
    this.afterSelect();
  }

  private onMove(e: PointerEvent): void {
    if (this.touch && e.pointerId === this.touch.id) { this.onTouchMove(e); return; }
    const p = this.local(e);
    if (e.buttons && Math.hypot(e.clientX - this.downAt.x, e.clientY - this.downAt.y) > 4) this.downAt.moved = true;
    if (!this.drag) {
      const t = this.target(p.x, p.y);
      this.canvas.style.cursor = POINTERS[t.kind === "pageBreak" ? (t.row >= 0 ? "resizeRow" : "resizeCol") : t.kind] ?? "default";
      this.canvas.dataset.pointer = t.kind;
      this.noteHover(t.kind === "cells" ? t : null, p);
      return;
    }
    const d = this.drag;
    this.autoScroll(p);
    const hit = this.hit(Math.max(this.headerW() + 1, Math.min(this.width - 1, p.x)), Math.max(this.headerH() + 1, Math.min(this.height - 1, p.y)));
    switch (d.kind) {
      case "point":
        this.point(this.pointRef?.anchor ?? d.start, { row: hit.row, col: hit.col });
        return;
      case "cells":
        if (hit.row !== this.head.row || hit.col !== this.head.col) this.extendTo({ row: hit.row, col: hit.col });
        return;
      case "cols":
        this.fullCols(d.start.col, hit.col);
        return;
      case "rows":
        this.fullRows(d.start.row, hit.row);
        return;
      case "resizeCol": {
        const raw = (d.extra as number) + (p.x - d.x) / this.zoom;
        const w = Math.max(0, raw);
        this.liveBefore = this.overshoot("col", d.start.col, -raw);
        this.canvas.dataset.tip = `Breite: ${pxToChars(w).toFixed(2).replace(".", ",")} (${Math.round(w * this.dpr)} Pixel)`;
        this.ghostSize("col", d.start.col, w);
        return;
      }
      case "resizeRow": {
        const raw = (d.extra as number) + (p.y - d.y) / this.zoom;
        const hgt = Math.max(0, raw);
        this.liveBefore = this.overshoot("row", d.start.row, -raw);
        this.canvas.dataset.tip = `Höhe: ${(hgt * 0.75).toFixed(2).replace(".", ",")} (${Math.round(hgt)} Pixel)`;
        this.ghostSize("row", d.start.row, hgt);
        return;
      }
      case "fill": {
        const sel = this.selection[0];
        // Upwards and leftwards a row or column counts once the pointer is past its middle (Excel).
        const cellBox = this.rectOnCanvas({ r0: hit.row, c0: hit.col, r1: hit.row, c1: hit.col });
        if (cellBox && hit.row < sel.r0 && p.y > cellBox.y + cellBox.h / 2) hit.row++;
        if (cellBox && hit.col < sel.c0 && p.x > cellBox.x + cellBox.w / 2) hit.col++;
        const dr = hit.row > sel.r1 ? hit.row - sel.r1 : hit.row < sel.r0 ? hit.row - sel.r0 : 0;
        const dc = hit.col > sel.c1 ? hit.col - sel.c1 : hit.col < sel.c0 ? hit.col - sel.c0 : 0;
        // The fill goes along the axis the pointer moved further on (Excel); back inside the
        // selection it marks the cells that will be emptied.
        if (Math.abs(dr) >= Math.abs(dc) && dr) this.ghost = dr > 0 ? { ...sel, r1: hit.row } : { ...sel, r0: hit.row };
        else if (dc) this.ghost = dc > 0 ? { ...sel, c1: hit.col } : { ...sel, c0: hit.col };
        else if (sel.r1 > sel.r0 && hit.row < sel.r1 && hit.row >= sel.r0) this.ghost = { ...sel, r1: hit.row };
        else if (sel.c1 > sel.c0 && hit.col < sel.c1 && hit.col >= sel.c0) this.ghost = { ...sel, c1: hit.col };
        else this.ghost = null;
        this.invalidate();
        return;
      }
      case "pageBreak": {
        // The line follows the gridline nearest the pointer.
        const rows = d.start.row >= 0;
        const cell = this.rectOnCanvas({ r0: hit.row, c0: hit.col, r1: hit.row, c1: hit.col });
        const after = cell ? (rows ? p.y > cell.y + cell.h / 2 : p.x > cell.x + cell.w / 2) : false;
        const at = Math.max(1, (rows ? hit.row : hit.col) + (after ? 1 : 0));
        this.insertLine = rows ? { r0: at, c0: 0, r1: at - 1, c1: this.cols.at(this.cols.pos(this.view.left) + this.width / this.zoom) } : { r0: 0, c0: at, r1: this.rows.at(this.rows.pos(this.view.top) + this.height / this.zoom), c1: at - 1 };
        this.invalidate();
        return;
      }
      case "move": {
        const sel = this.selection[0];
        const dr = hit.row - d.start.row, dc = hit.col - d.start.col;
        // Umschalt: an insertion line between rows, under the cell at the pointer when the pointer
        // is in its lower half (the middle included), above it otherwise; whole columns go between
        // columns (recorded in Excel, 7 cases).
        if (e.shiftKey && this.host.moveInsert) {
          const rows = !(sel.r0 === 0 && sel.r1 === MAX_ROWS - 1);
          const cell = this.rectOnCanvas({ r0: hit.row, c0: hit.col, r1: hit.row, c1: hit.col });
          const after = cell ? (rows ? p.y >= cell.y + cell.h / 2 : p.x >= cell.x + cell.w / 2) : false;
          const at = (rows ? hit.row : hit.col) + (after ? 1 : 0);
          const c0 = Math.max(0, Math.min(MAX_COLS - 1 - (sel.c1 - sel.c0), sel.c0 + dc));
          const r0 = Math.max(0, Math.min(MAX_ROWS - 1 - (sel.r1 - sel.r0), sel.r0 + dr));
          this.insertLine = rows ? { r0: at, c0, r1: at - 1, c1: c0 + sel.c1 - sel.c0 } : { r0, c0: at, r1: r0 + sel.r1 - sel.r0, c1: at - 1 };
          this.ghost = null;
          this.invalidate();
          return;
        }
        this.insertLine = null;
        const r0 = Math.max(0, Math.min(MAX_ROWS - 1 - (sel.r1 - sel.r0), sel.r0 + dr));
        const c0 = Math.max(0, Math.min(MAX_COLS - 1 - (sel.c1 - sel.c0), sel.c0 + dc));
        this.ghost = { r0, c0, r1: r0 + sel.r1 - sel.r0, c1: c0 + sel.c1 - sel.c0 };
        this.nameBox.value = rangeName(this.ghost);
        this.invalidate();
        return;
      }
    }
  }

  /** The context menu after a right-button drag is the drag's own menu, not the cell menu. */
  private rightDragEnded = false;

  /** Umschalt+drag: where the moved cells will be inserted (r1 < r0: between rows, c1 < c0: between columns). */
  private insertLine: Rect | null = null;

  private optBtn: HTMLElement | null = null;
  private optRect: Rect | null = null;

  /** Einfügeoptionen / AutoAusfülloptionen: the small button at the bottom right of the range (Excel). */
  showOptions(r: Rect | null, icon?: string, open?: (anchor: DOMRect) => void): void {
    this.optBtn?.remove();
    this.optBtn = null;
    this.optRect = r;
    if (!r || !open) return;
    this.optBtn = h("button", { class: "opt-btn", "aria-label": icon === "fill" ? "AutoAusfülloptionen" : "Einfügeoptionen", "data-tip": icon === "fill" ? "AutoAusfülloptionen" : "Einfügeoptionen (Strg)" },
      h("span", { class: icon === "fill" ? "opt-fill" : "opt-paste" }), "⌄");
    this.optBtn.addEventListener("mousedown", (e) => { e.preventDefault(); open(this.optBtn!.getBoundingClientRect()); });
    this.el.append(this.optBtn);
    this.placeOptions();
  }

  /** Whether the Einfüge-/AutoAusfülloptionen button is showing (Schnellanalyse then stays away). */
  optionsShown(): boolean {
    return !!this.optBtn;
  }

  private placeOptions(): void {
    if (!this.optBtn || !this.optRect) return;
    const b = this.rectOnCanvas(this.optRect);
    if (!b) return;
    this.optBtn.style.left = `${b.x + b.w + 2}px`;
    this.optBtn.style.top = `${b.y + b.h + 2}px`;
  }

  private noteEl: HTMLElement | null = null;
  /** Hovering a cell with a note shows it beside the cell (Excel). */
  private noteHover(t: { row: number; col: number } | null, _p: { x: number; y: number }): void {
    const note = t ? this.sheet.get(t.row, t.col)?.note : undefined;
    if (!note || !t) { this.noteEl?.remove(); this.noteEl = null; return; }
    const b = this.rectOnCanvas({ r0: t.row, c0: t.col, r1: t.row, c1: t.col });
    if (!b) return;
    if (!this.noteEl) { this.noteEl = h("div", { class: "note-pop" }); this.el.append(this.noteEl); }
    this.noteEl.textContent = note;
    this.noteEl.style.left = `${b.x + b.w + 8}px`;
    this.noteEl.style.top = `${b.y}px`;
  }

  /**
   * A border dragged past its column's start: the column is hidden and the drag goes on into the
   * visible columns before it, which shrink to the pointer (recorded in Excel). Rows likewise.
   */
  private overshoot(kind: "col" | "row", i: number, over: number): { i: number; size: number }[] {
    const out: { i: number; size: number }[] = [];
    const a = sheetAxes(this.sheet);
    const axis = kind === "col" ? a.cols : a.rows;
    for (let j = i - 1; over > 0 && j >= 0; j--) {
      const s = axis.size(j);
      if (s <= 0) continue;
      out.push({ i: j, size: Math.max(0, Math.round(s - over)) });
      over -= s;
    }
    return out;
  }

  private liveBefore: { i: number; size: number }[] = [];

  private ghostSize(kind: "col" | "row", i: number, size: number): void {
    // Live: the column or row changes while dragging, as in Excel 365.
    this.liveSize = { kind, i, size };
    const a = sheetAxes(this.sheet);
    this.rows = a.rows;
    this.cols = a.cols;
    const axis = kind === "col" ? this.cols : this.rows;
    const all = new Map<number, number>();
    const keys = kind === "col" ? [...this.sheet.cols.keys()] : [...this.sheet.rows.keys()];
    for (const k of keys) all.set(k, axis.size(k));
    all.set(i, Math.round(size));
    for (const b of this.liveBefore) all.set(b.i, b.size);
    axis.set(all, axis.def);
    this.textCache.clear();
    this.invalidate();
  }

  private liveSize: { kind: "col" | "row"; i: number; size: number } | null = null;

  private onUp(e: PointerEvent): void {
    if (this.touch && e.pointerId === this.touch.id) { this.onTouchUp(e); return; }
    const d = this.drag;
    this.drag = null;
    this.stopAutoScroll();
    if (!d) return;
    delete this.canvas.dataset.tip;
    if (d.extra === "right" && (d.kind === "fill" || d.kind === "move")) {
      const g = this.ghost;
      this.ghost = null;
      this.insertLine = null;
      this.invalidate();
      if (!g || !this.downAt.moved) return;
      this.rightDragEnded = true;
      setTimeout(() => { this.rightDragEnded = false; }, 400);
      this.host.rightDrag!(d.kind, { ...this.selection[0] }, g, { x: e.clientX, y: e.clientY });
      return;
    }
    if (d.kind === "resizeCol" || d.kind === "resizeRow") {
      const live = this.liveSize;
      this.liveSize = null;
      if (live) {
        const kind = d.kind === "resizeCol" ? "cols" : "rows";
        const sel = this.selection.at(-1)!;
        const whole = kind === "cols" ? sel.r0 === 0 && sel.r1 === MAX_ROWS - 1 && live.i >= sel.c0 && live.i <= sel.c1
          : sel.c0 === 0 && sel.c1 === MAX_COLS - 1 && live.i >= sel.r0 && live.i <= sel.r1;
        const list: number[] = [];
        if (whole) for (let i = kind === "cols" ? sel.c0 : sel.r0; i <= (kind === "cols" ? sel.c1 : sel.r1); i++) list.push(i);
        else list.push(live.i);
        const before = whole ? [] : this.liveBefore;
        this.liveBefore = [];
        if (kind === "cols") this.host.resizeCols(list, Math.round(live.size), before);
        else this.host.resizeRows(list, Math.round(live.size), before);
      }
      return;
    }
    if (d.kind === "fill" && this.ghost) {
      const g = this.ghost;
      this.ghost = null;
      const src = this.selection[0];
      // Dragging the handle back into the selection empties the cells left out (Excel).
      const down = g.r1 > src.r1, right = g.c1 > src.c1;
      // With Umschalt the handle inserts empty cells, or deletes them when dragged back (recorded in Excel).
      if (e.shiftKey && this.host.shiftFill && (down || right || g.r1 < src.r1 || g.c1 < src.c1) && g.r0 >= src.r0 && g.c0 >= src.c0) {
        if (down || right) this.host.shiftFill(down ? { ...src, r0: src.r1 + 1, r1: g.r1 } : { ...src, c0: src.c1 + 1, c1: g.c1 }, down ? "down" : "right");
        else this.host.shiftFill(g.r1 < src.r1 ? { ...src, r0: g.r1 + 1 } : { ...src, c0: g.c1 + 1 }, g.r1 < src.r1 ? "up" : "left");
        this.select(g, { ...this.active });
        return;
      }
      if (g.r0 >= src.r0 && g.r1 <= src.r1 && g.c0 >= src.c0 && g.c1 <= src.c1) {
        const cleared = g.r1 < src.r1 ? { ...src, r0: g.r1 + 1 } : { ...src, c0: g.c1 + 1 };
        this.host.clearRange(cleared);
        this.select(g, { ...this.active });
        return;
      }
      this.host.fill(src, g, e.ctrlKey);
      this.select(g, { ...this.active });
      return;
    }
    if (d.kind === "pageBreak") {
      const line = this.insertLine;
      this.insertLine = null;
      this.invalidate();
      if (!line) return;
      const rows = d.start.row >= 0;
      const to = rows ? line.r0 : line.c0, from = rows ? d.start.row : d.start.col;
      if (to !== from) this.host.movePageBreak?.(rows, from, to);
      return;
    }
    if (d.kind === "move" && this.insertLine) {
      const line = this.insertLine;
      this.insertLine = null;
      const src = this.selection[0];
      const rows = line.r1 < line.r0;
      const at = { row: line.r0, col: line.c0 };
      const same = rows ? line.c0 === src.c0 : line.r0 === src.r0;
      const lo = rows ? src.r0 : src.c0, hi = rows ? src.r1 : src.c1, pos = rows ? at.row : at.col;
      if (!same || pos < lo || pos > hi + 1) {
        const dst = this.host.moveInsert!(src, at, rows);
        if (dst) this.select(dst, { row: dst.r0 + (this.active.row - src.r0), col: dst.c0 + (this.active.col - src.c0) });
      }
      this.invalidate();
      return;
    }
    if (d.kind === "move" && this.ghost) {
      const g = this.ghost;
      this.ghost = null;
      const src = this.selection[0];
      if (g.r0 !== src.r0 || g.c0 !== src.c0) {
        this.host.move(src, { row: g.r0, col: g.c0 }, e.ctrlKey);
        this.select(g, { row: g.r0 + (this.active.row - src.r0), col: g.c0 + (this.active.col - src.c0) });
      }
      this.invalidate();
      return;
    }
    if (d.kind === "point") {
      this.editing?.viaBar ? this.fx.focus() : this.editor.focus({ preventScroll: true });
      return;
    }
    // A click on a link's text follows it (Excel); a click beside the text only selects.
    if (d.kind === "cells" && !this.downAt.moved && !e.shiftKey && !e.ctrlKey) {
      const c = this.sheet.get(this.active.row, this.active.col);
      if (c?.link && this.overText(e, this.active.row, this.active.col)) this.host.followLink(c.link);
    }
    this.syncBar();
  }

  /** Whether the pointer is over the text a cell shows. */
  private overText(e: MouseEvent, row: number, col: number): boolean {
    const b = this.rectOnCanvas({ r0: row, c0: col, r1: row, c1: col });
    const c = this.sheet.get(row, col);
    if (!b || !c) return false;
    const st = this.wb.styles.get(c.s);
    this.ctx.font = cssFont(st, this.zoom);
    const w = this.ctx.measureText(String(c.v ?? "")).width;
    const p = this.local(e);
    const right = st.h === "right" || (st.h === undefined && typeof c.v === "number");
    const x0 = right ? b.x + b.w - w - 3 : st.h === "center" ? b.x + (b.w - w) / 2 : b.x + 2;
    return p.x >= x0 && p.x <= x0 + w + 2;
  }

  private scrollTimer = 0;
  private autoScroll(p: { x: number; y: number }): void {
    const hw = this.headerW(), hh = this.headerH();
    const dy = p.y > this.height ? 1 : p.y < hh ? -1 : 0;
    const dx = p.x > this.width ? 1 : p.x < hw ? -1 : 0;
    if (!dx && !dy) { this.stopAutoScroll(); return; }
    if (this.scrollTimer) return;
    this.scrollTimer = window.setInterval(() => {
      if (dy) this.view.top = Math.max(this.view.freeze?.row ?? 0, this.rows.step(this.view.top, dy as 1 | -1));
      if (dx) this.view.left = Math.max(this.view.freeze?.col ?? 0, this.cols.step(this.view.left, dx as 1 | -1));
      this.updateScrollbars();
      this.invalidate();
    }, 50);
  }

  private stopAutoScroll(): void {
    clearInterval(this.scrollTimer);
    this.scrollTimer = 0;
  }

  private onDouble(e: MouseEvent): void {
    if (this.downAt.moved) return;
    const p = this.local(e);
    const t = this.target(p.x, p.y);
    if (t.kind === "resizeCol") {
      const sel = this.selection.at(-1)!;
      const whole = sel.r0 === 0 && sel.r1 === MAX_ROWS - 1 && t.col >= sel.c0 && t.col <= sel.c1;
      const list: number[] = [];
      if (whole) for (let c = sel.c0; c <= sel.c1; c++) list.push(c);
      else list.push(t.col);
      this.host.autofitCols(list);
      return;
    }
    if (t.kind === "resizeRow") {
      this.host.autofitRows([t.row]);
      return;
    }
    if (t.kind === "fill") {
      // Double-click on the fill handle: down as far as the column beside has data.
      const sel = this.selection[0];
      const besideCol = sel.c0 > 0 && !this.isEmpty(sel.r1 + 1, sel.c0 - 1) ? sel.c0 - 1 : !this.isEmpty(sel.r1 + 1, sel.c1 + 1) ? sel.c1 + 1 : -1;
      if (besideCol < 0) return;
      let end = sel.r1;
      while (end + 1 < MAX_ROWS && !this.isEmpty(end + 1, besideCol)) end++;
      if (end > sel.r1) {
        const g = { ...sel, r1: end };
        this.host.fill(sel, g, false);
        this.select(g, { ...this.active });
      }
      return;
    }
    if (t.kind === "move") {
      // Double-click on the border: Strg+arrow in that direction.
      const sel = this.selection[0];
      const b = this.rectOnCanvas(sel)!;
      const dir = Math.abs(p.y - (b.y + b.h)) <= 3 ? [1, 0] : Math.abs(p.y - b.y) <= 3 ? [-1, 0] : Math.abs(p.x - b.x) <= 3 ? [0, -1] : [0, 1];
      this.move(dir[0], dir[1], false, true);
      return;
    }
    if (t.kind === "cells") {
      // Bearbeiten with the caret where the double-click was.
      this.startEdit("edit");
      const b = this.rectOnCanvas({ r0: t.row, c0: t.col, r1: t.row, c1: t.col });
      if (b) {
        const ctx = this.ctx;
        ctx.font = this.editor.style.font;
        const text = this.editor.value;
        const startX = this.editor.dataset.align === "right" ? b.x + b.w - ctx.measureText(text).width - 3 : b.x + 3;
        let pos = 0;
        while (pos < text.length && startX + ctx.measureText(text.slice(0, pos + 1)).width - ctx.measureText(text[pos]).width / 2 < p.x) pos++;
        this.editor.setSelectionRange(pos, pos);
      }
    }
  }

  // ----- keyboard ------------------------------------------------------------------------------

  private onKey(e: KeyboardEvent, input: HTMLTextAreaElement): void {
    if (e.isComposing) return;
    if (this.acKey(e)) return;
    const ctrl = e.ctrlKey && !e.altKey, shift = e.shiftKey;
    if (this.editing) {
      this.editKey(e, input, ctrl, shift);
      return;
    }
    if (this.host.key(e)) { e.preventDefault(); return; }
    const k = e.key;
    // Alt+Unten on a filter's header cell opens its dropdown (Excel).
    const af = this.sheet.autoFilter;
    if (e.altKey && k === "ArrowDown" && af && this.active.row === af.ref.r0 && this.active.col >= af.ref.c0 && this.active.col <= af.ref.c1) {
      e.preventDefault();
      const b = this.rectOnCanvas({ r0: this.active.row, c0: this.active.col, r1: this.active.row, c1: this.active.col })!;
      const fb = filterButton(b.x, b.w, b.y, b.h, this.zoom);
      const c = this.canvas.getBoundingClientRect();
      this.host.filterMenu(this.active.col, new DOMRect(c.left + fb.x, c.top + fb.y, fb.w, fb.h));
      return;
    }
    // Alt+Unten elsewhere: the Dropdown-Auswahlliste of the column's texts (Excel).
    if (e.altKey && k === "ArrowDown") {
      e.preventDefault();
      this.pickList();
      return;
    }
    const arrows: Record<string, [number, number]> = { ArrowDown: [1, 0], ArrowUp: [-1, 0], ArrowRight: [0, 1], ArrowLeft: [0, -1] };
    // F8: Auswahl erweitern (arrows extend until F8 or Esc); Umschalt+F8: Abschnitt hinzufügen
    // (the next click adds a range). Both are shown in the status bar (recorded).
    if (k === "F8" && !ctrl && !e.altKey) {
      e.preventDefault();
      const want = shift ? "add" : "extend";
      this.setKeyMode(false, this.keyMode.select === want ? null : want);
      return;
    }
    // Ende: the next arrow jumps as Strg+arrow does; Ende again leaves the mode (recorded).
    if (k === "End" && !ctrl && !e.altKey && !shift) {
      e.preventDefault();
      this.setKeyMode(!this.keyMode.end, this.keyMode.select);
      return;
    }
    if (arrows[k] && !e.altKey) {
      e.preventDefault();
      this.tabStart = null;
      const jump = ctrl || this.keyMode.end;
      const extend = shift || this.keyMode.select === "extend";
      if (this.keyMode.end) this.setKeyMode(false, this.keyMode.select);
      this.move(arrows[k][0], arrows[k][1], extend, jump);
      return;
    }
    if (this.keyMode.end && !["Shift", "Control", "Alt"].includes(k)) this.setKeyMode(false, this.keyMode.select);
    if (k === "Escape" && this.keyMode.select) this.setKeyMode(false, null);
    // Strg+Rücktaste: back to where the active cell is.
    if (ctrl && k === "Backspace") {
      e.preventDefault();
      this.reveal(this.active.row, this.active.col);
      return;
    }
    switch (k) {
      case "Enter":
        e.preventDefault();
        // With copied cells waiting (marching ants), Enter pastes them and ends the copy (Excel).
        if (this.marquee && !shift && this.host.enterPaste()) return;
        this.moveWithin(shift ? -1 : 1, 0);
        return;
      case "Tab":
        e.preventDefault();
        if (!this.protectedTab(shift)) this.moveWithin(0, shift ? -1 : 1);
        return;
      case "F2":
        e.preventDefault();
        this.startEdit("edit");
        return;
      case "Delete":
        e.preventDefault();
        this.host.clearContents();
        return;
      case "Backspace":
        e.preventDefault();
        // Umschalt+Rücktaste: the selection shrinks to the active cell (recorded in Excel).
        if (shift) { this.selectCell(this.active.row, this.active.col); return; }
        // Excel: Rücktaste empties the active cell and starts Eingeben.
        this.startEdit("enter", "");
        return;
      case "Escape":
        if (this.marquee) { this.setMarquee(null); e.preventDefault(); }
        return;
      case "Home":
        e.preventDefault();
        if (ctrl) { const fr = this.view.freeze; this.go(fr?.row ?? 0, fr?.col ?? 0, shift); }
        else this.go(this.active.row, this.view.freeze?.col ?? 0, shift);
        return;
      case "End":
        if (ctrl) {
          e.preventDefault();
          const u = this.sheet.usedRange();
          this.go(u?.r1 ?? 0, u?.c1 ?? 0, shift);
        }
        return;
      case "PageDown":
      case "PageUp": {
        e.preventDefault();
        const dir = k === "PageDown" ? 1 : -1;
        if (ctrl) { this.host.key(new KeyboardEvent("keydown", { key: k, ctrlKey: true })); return; }
        if (e.altKey) {
          const n = this.visibleCols();
          let c = this.active.col;
          for (let i = 0; i < n; i++) c = this.cols.step(c, dir as 1 | -1);
          let left = this.view.left;
          for (let i = 0; i < n; i++) left = this.cols.step(left, dir as 1 | -1);
          this.view.left = Math.max(this.view.freeze?.col ?? 0, left);
          this.go(this.active.row, c, shift);
          return;
        }
        // Excel moves one row more than the window shows (recorded: A1 → A56 with 54 rows in view).
        const n = this.visibleRows() + 1;
        let r = shift ? this.head.row : this.active.row;
        for (let i = 0; i < n; i++) r = this.rows.step(r, dir as 1 | -1);
        let top = this.view.top;
        for (let i = 0; i < n; i++) top = this.rows.step(top, dir as 1 | -1);
        this.view.top = Math.max(this.view.freeze?.row ?? 0, top);
        this.go(r, shift ? this.head.col : this.active.col, shift);
        return;
      }
      case " ":
        if (ctrl && shift) { e.preventDefault(); this.select({ r0: 0, c0: 0, r1: MAX_ROWS - 1, c1: MAX_COLS - 1 }, { ...this.active }); return; }
        if (ctrl) { e.preventDefault(); const s = this.selection.at(-1)!; this.select({ r0: 0, c0: s.c0, r1: MAX_ROWS - 1, c1: s.c1 }, { ...this.active }); return; }
        if (shift) { e.preventDefault(); const s = this.selection.at(-1)!; this.select({ r0: s.r0, c0: 0, r1: s.r1, c1: MAX_COLS - 1 }, { ...this.active }); return; }
        return;
    }
    if (ctrl && (k === "a" || k === "A") && !shift) {
      e.preventDefault();
      const region = this.currentRegion(this.active.row, this.active.col);
      const cur = this.selection.at(-1)!;
      const isRegion = cur.r0 === region.r0 && cur.c0 === region.c0 && cur.r1 === region.r1 && cur.c1 === region.c1;
      const empty = region.r0 === region.r1 && region.c0 === region.c1 && this.isEmpty(region.r0, region.c0);
      if (isRegion || empty) this.select({ r0: 0, c0: 0, r1: MAX_ROWS - 1, c1: MAX_COLS - 1 }, { ...this.active });
      else this.select(region, { ...this.active });
    }
  }

  private go(row: number, col: number, extend: boolean): void {
    if (extend) this.extendTo({ row, col });
    else this.selectCell(row, col);
  }

  private editKey(e: KeyboardEvent, input: HTMLTextAreaElement, ctrl: boolean, shift: boolean): void {
    const k = e.key;
    if (k === "Enter" && e.altKey) {
      // Alt+Eingabe: a line break inside the cell.
      e.preventDefault();
      const s = input.selectionStart ?? 0, en = input.selectionEnd ?? s;
      input.value = input.value.slice(0, s) + "\n" + input.value.slice(en);
      input.setSelectionRange(s + 1, s + 1);
      (input === this.fx ? this.editor : this.fx).value = input.value;
      this.placeEditor();
      return;
    }
    if (ctrl && shift && (k === "a" || k === "A")) {
      // Strg+Umschalt+A: the function's argument names after its name, as Excel writes them.
      e.preventDefault();
      const pos = input.selectionStart ?? input.value.length;
      const m = /([A-Za-zÄÖÜäöü][A-Za-zÄÖÜäöü0-9.]*)(\(?)$/.exec(input.value.slice(0, pos));
      const fn = m && this.functions.find((f) => f.name.toUpperCase() === m[1].toUpperCase());
      if (!m || !fn) return;
      const ins = (m[2] ? "" : "(") + fn.args.replace(/[[\]]/g, "").replace(/;\s*/g, ";") + ")";
      input.value = input.value.slice(0, pos) + ins + input.value.slice(pos);
      input.setSelectionRange(pos + ins.length, pos + ins.length);
      (input === this.fx ? this.editor : this.fx).value = input.value;
      this.ac.hidden = true;
      this.acItems = [];
      this.placeEditor();
      return;
    }
    if (e.altKey && k === "ArrowDown" && this.mode !== "point") {
      // Alt+Unten while typing opens the Auswahlliste too; the choice goes into the cell's text.
      e.preventDefault();
      this.pickList(true);
      return;
    }
    if (k === "Enter") {
      e.preventDefault();
      if (ctrl && shift) this.commitEdit("none", { array: true });
      else if (ctrl) this.commitEdit("none", { all: true });
      else this.commitEdit(shift ? "up" : "down");
      return;
    }
    if (k === "Tab") {
      e.preventDefault();
      if (this.host.nextUnlocked?.(this.active.row, this.active.col, shift)) {
        this.commitEdit("none");
        this.protectedTab(shift);
      } else this.commitEdit(shift ? "left" : "right");
      return;
    }
    if (k === "Escape") {
      e.preventDefault();
      this.cancelEdit();
      return;
    }
    if (k === "F2") {
      e.preventDefault();
      this.setMode(this.mode === "edit" ? "enter" : "edit");
      this.pointRef = null;
      return;
    }
    if (k === "F4" && input.value.startsWith("=")) {
      e.preventDefault();
      this.cycleAbsolute(input);
      return;
    }
    const arrows: Record<string, [number, number]> = { ArrowDown: [1, 0], ArrowUp: [-1, 0], ArrowRight: [0, 1], ArrowLeft: [0, -1] };
    if (arrows[k] && !e.altKey) {
      const [dr, dc] = arrows[k];
      if (this.mode !== "edit" && (this.canPoint() || this.pointRef)) {
        e.preventDefault();
        const base = this.pointRef ? (shift ? this.pointRef.head : this.pointRef.head) : { ...this.editing!.row !== undefined ? { row: this.editing!.row, col: this.editing!.col } : this.active };
        const next = ctrl ? this.jump(base, dr, dc) : { row: Math.max(0, Math.min(MAX_ROWS - 1, base.row + dr)), col: Math.max(0, Math.min(MAX_COLS - 1, base.col + dc)) };
        if (shift && this.pointRef) this.point(this.pointRef.anchor, next);
        else this.point(next, next);
        return;
      }
      if (this.mode === "enter") {
        // Eingeben: arrows commit and move one cell (also out of a selection).
        e.preventDefault();
        this.commitEdit(dr > 0 ? "down" : dr < 0 ? "up" : dc > 0 ? "right" : "left", { arrow: true });
        return;
      }
      return;
    }
    if (ctrl && (k === "z" || k === "Z")) {
      // Strg+Z while typing takes back the typing but stays in the cell (recorded in Excel).
      e.preventDefault();
      const back = this.mode === "edit" ? this.editing!.original : "";
      this.editor.value = back;
      this.fx.value = back;
      input.setSelectionRange(back.length, back.length);
      this.ac.hidden = true;
      this.acItems = [];
      this.placeEditor();
      return;
    }
  }

  /** F4 while editing a formula: A1 → $A$1 → A$1 → $A1 → A1 for the reference at the caret. */
  private cycleAbsolute(input: HTMLTextAreaElement): void {
    const t = input.value;
    const pos = input.selectionStart ?? t.length;
    const re = /(\$?)([A-Za-z]{1,3})(\$?)(\d+)(?::(\$?)([A-Za-z]{1,3})(\$?)(\d+))?/g;
    let m: RegExpExecArray | null;
    while ((m = re.exec(t))) {
      const start = m.index, end = start + m[0].length;
      if (pos < start || pos > end) continue;
      const cycle = (a: string, b: string) => (!a && !b ? ["$", "$"] : a && b ? ["", "$"] : !a && b ? ["$", ""] : ["", ""]);
      // A pointed range changes as a whole; in a typed one only the cell at the caret (recorded in Excel).
      const pointed = !!this.pointRef && this.pointRef.start <= start && this.pointRef.end >= end;
      const firstEnd = start + m[1].length + m[2].length + m[3].length + m[4].length;
      const doFirst = !m[6] || pointed || pos <= firstEnd, doSecond = !!m[6] && (pointed || pos > firstEnd);
      const [c1, r1] = doFirst ? cycle(m[1], m[3]) : [m[1], m[3]];
      let out = `${c1}${m[2]}${r1}${m[4]}`;
      if (m[6]) {
        const [c2, r2] = doSecond ? cycle(m[5], m[7]) : [m[5], m[7]];
        out += `:${c2}${m[6]}${r2}${m[8]}`;
      }
      input.value = t.slice(0, start) + out + t.slice(end);
      input.setSelectionRange(start + out.length, start + out.length);
      (input === this.fx ? this.editor : this.fx).value = input.value;
      this.pointRef = null;
      this.invalidate();
      return;
    }
  }

  /** Dropdown-Auswahlliste: the different texts above and below in the column, sorted; the chosen one is entered. */
  private pickList(typing = false): void {
    const { row, col } = this.active;
    const seen = new Map<string, string>();
    const scan = (dir: 1 | -1) => {
      for (let r = row + dir; r >= 0 && r < MAX_ROWS; r += dir) {
        const c = this.sheet.get(r, col);
        if (!c || c.v === null || c.v === "") break;
        if (typeof c.v === "string" && c.f === undefined && !seen.has(c.v.toLowerCase())) seen.set(c.v.toLowerCase(), c.v);
      }
    };
    scan(-1);
    scan(1);
    const items = [...seen.values()].sort((a, b) => a.localeCompare(b, "de"));
    if (!items.length) return;
    const b = this.rectOnCanvas({ r0: row, c0: col, r1: row, c1: col })!;
    const list = h("div", { class: "pick-list", role: "listbox" });
    // Excel opens the list with nothing highlighted; the first Down picks the first entry. While
    // typing, the entry that starts with the typed text is highlighted.
    const typed = typing ? this.editor.value.toLowerCase() : "";
    let index = typed ? items.findIndex((t) => t.toLowerCase().startsWith(typed)) : -1;
    const close = () => { list.remove(); removeEventListener("keydown", keys, true); removeEventListener("pointerdown", outside, true); this.editor.focus({ preventScroll: true }); };
    const take = (i: number) => {
      close();
      if (typing && this.editing) {
        // The choice replaces the typed text; the cell stays in Eingeben.
        this.editor.value = items[i];
        this.fx.value = items[i];
        this.editor.setSelectionRange(items[i].length, items[i].length);
        this.placeEditor();
        return;
      }
      this.host.commit({ r0: row, c0: col, r1: row, c1: col }, items[i], {});
      this.syncBar();
    };
    const render = () => list.replaceChildren(...items.map((t, i) => h("div", { class: `pick-item${i === index ? " on" : ""}`, onmousedown: (ev: MouseEvent) => { ev.preventDefault(); take(i); } }, t)));
    const keys = (ev: KeyboardEvent) => {
      if (ev.key === "ArrowDown" || ev.key === "ArrowUp") { ev.preventDefault(); ev.stopPropagation(); index = ev.key === "ArrowDown" ? Math.min(items.length - 1, index + 1) : Math.max(0, index - 1); render(); }
      else if (ev.key === "Enter") { ev.preventDefault(); ev.stopPropagation(); if (index >= 0) take(index); else close(); }
      else if (ev.key === "Escape") { ev.preventDefault(); ev.stopPropagation(); close(); }
    };
    const outside = (ev: Event) => { if (!list.contains(ev.target as Node)) close(); };
    render();
    list.style.left = `${b.x}px`;
    list.style.top = `${b.y + b.h}px`;
    list.style.minWidth = `${b.w}px`;
    this.el.append(list);
    setTimeout(() => { addEventListener("keydown", keys, true); addEventListener("pointerdown", outside, true); });
  }

  // ----- function AutoComplete --------------------------------------------------------------------

  private acItems: { name: string; args: string }[] = [];
  private acIndex = 0;
  private functions: { name: string; args: string }[] = [];

  setFunctions(list: { name: string; args: string }[]): void {
    this.functions = list;
  }

  /** "=SU" shows the functions starting with SU; Tab takes the selected one (Excel). */
  private functionHint(): void {
    const input = this.editing?.viaBar ? this.fx : this.editor;
    const t = input.value;
    const pos = input.selectionStart ?? t.length;
    const m = t.startsWith("=") ? /([A-Za-zÄÖÜäöüß][A-Za-zÄÖÜäöüß0-9.]*)$/.exec(t.slice(0, pos)) : null;
    if (!m || /[A-Za-z]\$?\d/.test(m[1])) { this.ac.hidden = true; this.acItems = []; return; }
    const prefix = m[1].toUpperCase();
    this.acItems = this.functions.filter((f) => f.name.startsWith(prefix)).sort((a, b) => a.name.localeCompare(b.name, "de")).slice(0, 12);
    if (!this.acItems.length) { this.ac.hidden = true; return; }
    this.acIndex = 0;
    this.renderAc();
    const b = this.editing ? this.rectOnCanvas({ r0: this.editing.row, c0: this.editing.col, r1: this.editing.row, c1: this.editing.col }) : null;
    this.ac.style.left = `${b?.x ?? 0}px`;
    this.ac.style.top = `${(b ? b.y + b.h : 0) + 2}px`;
    this.ac.hidden = false;
  }

  private renderAc(): void {
    this.ac.replaceChildren(...this.acItems.map((f, i) => h("div", {
      class: `fn-ac-item${i === this.acIndex ? " on" : ""}`,
      onmousedown: (e: MouseEvent) => { e.preventDefault(); this.acIndex = i; this.takeAc(); },
    }, h("span", { class: "fn-ac-icon" }, "ƒ"), f.name)));
  }

  private acKey(e: KeyboardEvent): boolean {
    if (this.ac.hidden || !this.acItems.length) return false;
    if (e.key === "ArrowDown" || e.key === "ArrowUp") {
      e.preventDefault();
      this.acIndex = (this.acIndex + (e.key === "ArrowDown" ? 1 : -1) + this.acItems.length) % this.acItems.length;
      this.renderAc();
      return true;
    }
    if (e.key === "Tab") {
      e.preventDefault();
      this.takeAc();
      return true;
    }
    if (e.key === "Escape") {
      e.preventDefault();
      this.ac.hidden = true;
      return true;
    }
    return false;
  }

  private takeAc(): void {
    const input = this.editing?.viaBar ? this.fx : this.editor;
    const f = this.acItems[this.acIndex];
    const t = input.value;
    const pos = input.selectionStart ?? t.length;
    const m = /([A-Za-zÄÖÜäöüß][A-Za-zÄÖÜäöüß0-9.]*)$/.exec(t.slice(0, pos))!;
    const start = pos - m[1].length;
    input.value = t.slice(0, start) + f.name + "(" + t.slice(pos);
    const np = start + f.name.length + 1;
    input.setSelectionRange(np, np);
    (input === this.fx ? this.editor : this.fx).value = input.value;
    this.ac.hidden = true;
    this.placeEditor();
  }

  // ----- marquee -------------------------------------------------------------------------------

  setMarquee(r: Rect | null): void {
    if (!!r !== !!this.marquee) this.host.copyMode?.(!!r);
    this.marquee = r;
    clearInterval(this.marqueeTimer);
    if (r) this.marqueeTimer = window.setInterval(() => { this.marqueePhase = (this.marqueePhase + 1) % 16; this.invalidate(); }, 120);
    this.invalidate();
  }

  /** The column letter label for menus ("Spalte B"). */
  static colLabel(c: number): string {
    return colName(c);
  }
}
