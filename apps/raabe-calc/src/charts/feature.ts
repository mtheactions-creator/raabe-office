// Charts in RaabeCalc (Step 7 E, D-061). A chart is made from the selection (or the current
// region around the active cell) as Excel splits it into series, placed in the middle of the
// visible cells at Excel's default size (480 x 288 px), and kept in `sheet.charts` with a
// two-cell anchor as the file has it. Charts float over the grid as SVG in a layer of their own,
// follow scrolling and zoom, and re-read their cells after every change. A chart is selected
// with a click (frame with 8 handles as in Excel), moved and sized with the mouse, deleted with
// Entf. Beside a selected chart stand Excel's three buttons (Diagrammelemente, Diagrammformatvorlagen,
// Diagrammfilter, buttons.ts); its source ranges are framed in the sheet in Excel's colours and can be
// dragged (ranges below); a right-click gives the element's context menu. Every change is one undo
// step (History.charts).

import {
  QUICK_LAYOUTS, isPie, newChart, parseRef, refText, refreshCaches, renderChart, seriesFromBlock, seriesName, sourceOf, titleText, visibleModel,
  type CellAnchor, type CellValue, type ChartKind as Kind, type ChartModel, type ChartSeries, type Ref,
} from "@raabe/charts";
import { CellError, formatValue, type Rect, type Sheet, type Workbook } from "@raabe/sheet-engine";
import { h, showDialog, showMenu, type MenuEntry } from "@raabe/shell";
import type { CalcApp } from "../app";
import { chartElements, elementName, elementOfPart, formatPane, partsSelector, type ChartElement } from "@raabe/chart-ui";
import { sheetAxes, type Axis } from "../grid/metrics";
import { chartButtons } from "@raabe/chart-ui";
import { choiceLabel, elementEntries } from "@raabe/chart-ui";
import { selectDataDialog, type SelectDataHost } from "@raabe/chart-ui";
import { applyStyle, chartStyles, setStyleTheme, styleIndex } from "@raabe/chart-ui";
import { chartMarks, drawMarks, elementMarks, marksLayer } from "@raabe/chart-ui";

export type ChartKind = Kind;

/** Excel's size of a new chart: 360 x 216 pt (12,7 x 7,62 cm). */
const NEW_W = 480, NEW_H = 288;
/** Smallest size while sizing with the mouse (px at 100 %). */
const MIN_SIZE = 20;
const SCROLLBAR = 17;
/** The source ranges' colours while a chart is selected (Excel: names red, categories purple, values blue). */
const RANGE_COLORS = { name: "#E0301E", cat: "#7030A0", val: "#0070C0" } as const;

interface View { el: HTMLDivElement; key: string; zoomKey: string }
interface Box { x: number; y: number; w: number; h: number }
type Handle = "nw" | "n" | "ne" | "e" | "se" | "s" | "sw" | "w";
const HANDLES: Handle[] = ["nw", "n", "ne", "e", "se", "s", "sw", "w"];

const CSS = `
.chart-layer { position: absolute; overflow: hidden; pointer-events: none; z-index: 4; }
.chart-obj { position: absolute; pointer-events: auto; outline: none; cursor: default; touch-action: none; }
.chart-obj > svg { display: block; position: absolute; left: 0; top: 0; }
.chart-obj.dragging { cursor: move; }
.chart-handle { position: absolute; width: 12px; height: 12px; margin: -6px 0 0 -6px; border-radius: 50%; display: none; z-index: 2; }
.chart-obj.sel .chart-handle { display: block; }
.chart-handle[data-h="nw"], .chart-handle[data-h="se"] { cursor: nwse-resize; }
.chart-handle[data-h="ne"], .chart-handle[data-h="sw"] { cursor: nesw-resize; }
.chart-handle[data-h="n"], .chart-handle[data-h="s"] { cursor: ns-resize; }
.chart-handle[data-h="e"], .chart-handle[data-h="w"] { cursor: ew-resize; }
.chart-obj:not(.sel) > .chart-marks { display: none; }
.chart-obj svg [data-fade] { opacity: .25; }
.chart-ranges { position: absolute; inset: 0; pointer-events: none; }
.chart-range { position: absolute; box-sizing: border-box; border: 2px solid; pointer-events: none; }
.chart-range .edge { position: absolute; pointer-events: auto; cursor: move; }
.chart-range .edge.t { left: -3px; right: -3px; top: -4px; height: 6px; } .chart-range .edge.b { left: -3px; right: -3px; bottom: -4px; height: 6px; }
.chart-range .edge.l { top: -3px; bottom: -3px; left: -4px; width: 6px; } .chart-range .edge.r { top: -3px; bottom: -3px; right: -4px; width: 6px; }
.chart-range .corner { position: absolute; width: 6px; height: 6px; pointer-events: auto; border: 1px solid #fff; }
.chart-range .corner.nw { left: -5px; top: -5px; cursor: nwse-resize; } .chart-range .corner.se { right: -5px; bottom: -5px; cursor: nwse-resize; }
.chart-range .corner.ne { right: -5px; top: -5px; cursor: nesw-resize; } .chart-range .corner.sw { left: -5px; bottom: -5px; cursor: nesw-resize; }
.chart-range.pick { border-style: dashed; border-color: #107c41; }
`;

export function chartsFeature(app: CalcApp) {
  let layer: HTMLDivElement | null = null;
  let rangesEl: HTMLDivElement | null = null;
  let views = new Map<ChartModel, View>();
  let selected: ChartModel | null = null;
  let axes: { rows: Axis; cols: Axis } | null = null;
  let viewSig = "";
  let shownSheet: Sheet | null = null;
  let shownList: ChartModel[] | undefined;
  /** What a drag needs of a pointer event (a modifier key mid-drag replays the last one, D-128). */
  type DragPoint = { clientX: number; clientY: number; shiftKey: boolean; ctrlKey: boolean; altKey: boolean };
  let drag: { m: ChartModel; handle: Handle | null; x: number; y: number; box: Box; draft: Box; moved: boolean; last?: DragPoint } | null = null;
  /** The selected element of the selected chart (Excel: Diagrammelemente on the Format tab). */
  let element: ChartElement = { kind: "chart" };
  /** Livevorschau and Datenquelle auswählen: what is drawn in place of a chart. */
  const looks = new Map<ChartModel, ChartModel>();
  /** Diagrammfilter: the series or category under the mouse (the rest fades). */
  let fade: { m: ChartModel; series?: number; cat?: number } | null = null;
  /** A range pointed at in a dialog (marquee). */
  let pickRect: Rect | null = null;

  const zoom = () => app.sheet.view.zoom / 100;

  // ----- set-up --------------------------------------------------------------------------------

  function ensureLayer(): HTMLDivElement {
    if (layer) return layer;
    if (!document.getElementById("raabe-charts-css")) document.head.append(h("style", { id: "raabe-charts-css" }, CSS));
    layer = h("div", { class: "chart-layer" });
    rangesEl = h("div", { class: "chart-ranges" });
    layer.append(rangesEl);
    const canvas = app.grid.el.querySelector(".grid-canvas");
    app.grid.el.insertBefore(layer, canvas ? canvas.nextSibling : app.grid.el.firstChild);
    // A click anywhere else in the grid deselects the chart (capture: before the grid sees it).
    app.grid.el.addEventListener("pointerdown", (e) => {
      if (!(e.target as HTMLElement).closest?.(".chart-obj, .chart-range")) select(null);
    }, true);
    const tick = () => {
      watch();
      requestAnimationFrame(tick);
    };
    requestAnimationFrame(tick);
    return layer;
  }

  /** Follows scrolling, zoom, sheet and document switches (app.ts calls refresh only after changes). */
  function watch(): void {
    if (!layer) return;
    const sheet = app.sheet;
    if (sheet !== shownSheet || sheet.charts !== shownList) { refresh(); return; }
    if (!views.size && !pickRect) return;
    const v = sheet.view;
    const sig = `${v.top},${v.left},${v.zoom},${v.freeze?.row ?? 0},${v.freeze?.col ?? 0},${app.grid.el.clientWidth},${app.grid.el.clientHeight}`;
    if (sig !== viewSig) place();
  }

  // ----- cells and anchors ----------------------------------------------------------------------

  function cellValue(v: unknown): CellValue {
    if (v instanceof CellError || v === undefined) return null;
    return v as CellValue;
  }

  /** The block of cells Excel charts: the selection, or the current region around one cell. */
  function sourceRect(): Rect | null {
    const sel = app.grid.selection.at(-1);
    if (!sel) return null;
    let r: Rect = sel;
    if (sel.r0 === sel.r1 && sel.c0 === sel.c1) r = app.grid.currentRegion(app.grid.active.row, app.grid.active.col);
    // Whole rows or columns: only the used part.
    const used = app.sheet.usedRange();
    if (!used) return r.r0 === r.r1 && r.c0 === r.c1 && app.sheet.value(r.r0, r.c0) === null ? null : r;
    const out = { r0: Math.max(r.r0, used.r0), c0: Math.max(r.c0, used.c0), r1: Math.min(r.r1, used.r1), c1: Math.min(r.c1, used.c1) };
    return out.r0 > out.r1 || out.c0 > out.c1 ? null : out;
  }

  /** Re-reads names, categories and values from the workbook (Excel's caches). */
  function readCaches(wb: Workbook, m: ChartModel): void {
    const read = (ref0: string | undefined) => {
      // Excel 2016's chart types refer to hidden names (_xlchart.v1.0 …) that hold the range.
      const ref = ref0 && !parseRef(ref0) ? wb.names.find((n) => n.name === ref0)?.ref ?? ref0 : ref0;
      const r = ref ? parseRef(ref) : null;
      const sheet = r ? wb.sheetByName(r.sheet) : undefined;
      if (!r || !sheet) return null;
      const vals: CellValue[] = [];
      let format: string | undefined, cols = 0;
      for (let row = r.r0; row <= r.r1; row++) {
        // Data in hidden rows and columns (also those a filter hides) is left out, as Excel does by default
        // (Daten in ausgeblendeten Zeilen und Spalten anzeigen keeps it).
        if (!m.showHidden && sheet.rows.get(row)?.hidden && r.r1 > r.r0) continue;
        cols = 0;
        for (let col = r.c0; col <= r.c1; col++) {
          if (!m.showHidden && sheet.cols.get(col)?.hidden && r.c1 > r.c0) continue;
          const cell = sheet.get(row, col);
          if (format === undefined && cell) format = wb.styles.get(cell.s).numFmt;
          const v = cell?.v ?? null;
          // Leere Zellen als Null (Lücken and Datenpunkte mit Linie verbinden are the renderer's empty points).
          vals.push(v === null && m.dispBlanksAs === "zero" ? 0 : cellValue(v));
          cols++;
        }
      }
      return { vals, format, cols: Math.max(1, cols) };
    };
    refreshCaches(m, read, (v, fmt) => (typeof v === "number" ? formatValue(v, fmt ?? "General").text : String(v)));
  }

  function getAxes(): { rows: Axis; cols: Axis } {
    if (!axes) axes = sheetAxes(app.sheet);
    return axes;
  }

  /** The chart's rectangle in sheet px at 100 %. */
  function sheetBox(m: ChartModel): Box {
    const { rows, cols } = getAxes();
    const x = cols.pos(m.from.col) + m.from.x, y = rows.pos(m.from.row) + m.from.y;
    const x2 = cols.pos(m.to.col) + m.to.x, y2 = rows.pos(m.to.row) + m.to.y;
    return { x, y, w: Math.max(1, x2 - x), h: Math.max(1, y2 - y) };
  }

  function anchorAt(x: number, y: number): CellAnchor {
    const { rows, cols } = getAxes();
    const col = cols.at(Math.max(0, x)), row = rows.at(Math.max(0, y));
    return { col, row, x: Math.max(0, x - cols.pos(col)), y: Math.max(0, y - rows.pos(row)) };
  }

  /** The visible cells in sheet px at 100 % (the scrolling pane). */
  function visibleBox(): Box {
    const { rows, cols } = getAxes();
    const v = app.sheet.view, z = zoom();
    const top = Math.max(v.top, v.freeze?.row ?? 0), left = Math.max(v.left, v.freeze?.col ?? 0);
    const tl = app.grid.rectOnCanvas({ r0: top, c0: left, r1: top, c1: left });
    const w = (app.grid.el.clientWidth - SCROLLBAR - (tl?.x ?? 0)) / z, hh = (app.grid.el.clientHeight - (tl?.y ?? 0)) / z;
    return { x: cols.pos(left), y: rows.pos(top), w: Math.max(0, w), h: Math.max(0, hh) };
  }

  // ----- drawing --------------------------------------------------------------------------------

  function accents(): string[] | undefined {
    const t = app.wb.theme;
    const a = [1, 2, 3, 4, 5, 6].map((i) => t[`accent${i}`]);
    return a.every((c) => typeof c === "string" && /^#[0-9a-f]{6}$/i.test(c)) ? a : undefined;
  }

  const format = (v: number, code?: string) => formatValue(v, code ?? "General").text;

  function viewOf(m: ChartModel): View {
    let v = views.get(m);
    if (v) return v;
    const el = h("div", { class: "chart-obj", tabindex: "-1", role: "img" }) as HTMLDivElement;
    el.addEventListener("pointerdown", (e) => onDown(e, m));
    el.addEventListener("pointermove", (e) => onMove(e));
    el.addEventListener("pointerup", (e) => onUp(e));
    el.addEventListener("pointercancel", () => dropDrag());
    el.addEventListener("dblclick", (e) => {
      e.preventDefault();
      e.stopPropagation();
      if ((e.target as HTMLElement).closest?.(".chart-btn")) return;
      select(m);
      // Excel: a double-click on an element opens its Format pane.
      setElement(m, elementAt(m, e.target as Element, e) ?? element);
      openPane(m);
    });
    el.addEventListener("contextmenu", (e) => {
      e.preventDefault();
      e.stopPropagation();
      if ((e.target as HTMLElement).closest?.(".chart-btn")) return;
      select(m);
      const at = elementAt(m, e.target as Element, e) ?? { kind: "chart" as const };
      setElement(m, at);
      // After the ribbon has followed the selection (its update closes open popups).
      const x = e.clientX, y = e.clientY;
      setTimeout(() => contextMenu(m, at, x, y), 0);
    });
    el.addEventListener("keydown", (e) => onKey(e, m));
    // A modifier key pressed or let go mid-drag applies at once (D-128).
    el.addEventListener("keyup", (e) => { if (drag && (e.key === "Shift" || e.key === "Control" || e.key === "Alt")) { e.preventDefault(); replay(e); } });
    el.addEventListener("wheel", () => buttons.close(), { passive: true });
    v = { el, key: "", zoomKey: "" };
    views.set(m, v);
    return v;
  }

  /** Draws a chart at a size (px at 100 %) unless it is already drawn so. */
  function draw(m: ChartModel, v: View, box: Box): void {
    const z = zoom();
    const look = looks.get(m) ?? m;
    const { from: _f, to: _t, raw: _r, ...content } = look;
    void _f; void _t; void _r;
    const fd = fade?.m === m ? fade : null;
    const key = JSON.stringify([content, Math.round(box.w), Math.round(box.h), app.wb.theme, fd && [fd.series, fd.cat]]);
    if (key !== v.key) {
      const svg = renderChart(look, box.w, box.h, { accents: accents(), format });
      svg.setAttribute("aria-label", titleText(m));
      v.el.querySelector(":scope > svg:not(.chart-marks)")?.remove();
      v.el.prepend(svg);
      if (fd) fadeParts(look, svg, fd);
      if (!v.el.querySelector(".chart-handle")) {
        for (const hd of HANDLES) v.el.append(h("div", { class: "chart-handle", "data-h": hd }));
        buttons.attach(v.el, () => m);
      }
      v.key = key;
      v.zoomKey = "";
    }
    const zk = `${z},${box.w},${box.h}`;
    if (zk !== v.zoomKey) {
      const svg = v.el.querySelector<SVGSVGElement>(":scope > svg:not(.chart-marks)")!;
      svg.setAttribute("width", String(box.w * z));
      svg.setAttribute("height", String(box.h * z));
      for (const hd of v.el.querySelectorAll<HTMLElement>(".chart-handle")) {
        const k = hd.dataset.h as Handle;
        hd.style.left = `${k.includes("w") ? 0 : k.includes("e") ? box.w * z : (box.w * z) / 2}px`;
        hd.style.top = `${k.includes("n") ? 0 : k.includes("s") ? box.h * z : (box.h * z) / 2}px`;
      }
      v.zoomKey = zk;
    }
  }

  /** Fades every series but one (or every point but one category), as Excel shows the filter entry under the mouse. */
  function fadeParts(m: ChartModel, svg: SVGSVGElement, f: { series?: number; cat?: number }): void {
    const { index } = visibleModel(m);
    const hiddenBefore = (cat: number) => (m.hiddenCats ?? []).filter((c) => c < cat).length;
    for (const el of svg.querySelectorAll<SVGElement>('[data-part^="series:"]')) {
      const [, si, pi] = (el.dataset.part ?? "").split(":");
      const full = index[Number(si)];
      if (f.series !== undefined && full !== f.series) el.dataset.fade = "1";
      if (f.cat !== undefined && pi !== undefined && Number(pi) !== f.cat - hiddenBefore(f.cat)) el.dataset.fade = "1";
      if (f.cat !== undefined && (m.hiddenCats ?? []).includes(f.cat)) el.dataset.fade = "1";
    }
  }

  /** Puts every chart where its anchor is on the canvas. */
  function place(): void {
    if (!layer) return;
    const v = app.sheet.view;
    viewSig = `${v.top},${v.left},${v.zoom},${v.freeze?.row ?? 0},${v.freeze?.col ?? 0},${app.grid.el.clientWidth},${app.grid.el.clientHeight}`;
    const z = zoom();
    // The layer covers the cells (not the headers or the scroll bar).
    const fr = v.freeze?.row ?? 0, fc = v.freeze?.col ?? 0;
    const tl = app.grid.rectOnCanvas({ r0: fr ? 0 : v.top, c0: fc ? 0 : v.left, r1: fr ? 0 : v.top, c1: fc ? 0 : v.left });
    const lx = tl?.x ?? 0, ly = tl?.y ?? 0;
    Object.assign(layer.style, { left: `${lx}px`, top: `${ly}px`, right: `${SCROLLBAR}px`, bottom: "0px" });
    for (const [m, view] of views) {
      const box = drag?.m === m ? drag.draft : sheetBox(m);
      const a = anchorAt(box.x, box.y);
      const cell = app.grid.rectOnCanvas({ r0: a.row, c0: a.col, r1: a.row, c1: a.col });
      if (!cell) { view.el.style.display = "none"; continue; }
      view.el.style.display = "";
      view.el.style.left = `${cell.x + a.x * z - lx}px`;
      view.el.style.top = `${cell.y + a.y * z - ly}px`;
      view.el.style.width = `${box.w * z}px`;
      view.el.style.height = `${box.h * z}px`;
      draw(m, view, box);
      if (m === selected) showElement(m);
    }
    showRanges(lx, ly);
  }

  /** Rebuilds the charts of the active sheet (after changes, sheet or document switches). */
  function refresh(): void {
    if (!app.grid) return;
    const l = ensureLayer();
    const sheet = app.sheet;
    const list = sheet.charts ?? [];
    axes = null;
    // Undo/redo replaces the models: the selection stays on the chart at the same place in the list (Excel).
    const was = selected && sheet === shownSheet ? (shownList ?? []).indexOf(selected) : -1;
    shownSheet = sheet;
    shownList = sheet.charts;
    for (const m of list) readCaches(app.wb, m);
    const next = new Map<ChartModel, View>();
    for (const m of list) {
      const v = viewOf(m);
      next.set(m, v);
      if (v.el.parentElement !== l) l.append(v.el);
    }
    for (const [m, v] of views) if (!next.has(m)) v.el.remove();
    views = next;
    if (selected && !next.has(selected)) {
      const again = was >= 0 ? list[was] : undefined;
      const flyout = buttons.isOpen() ? buttons.which() : null;
      buttons.close();
      if (again) {
        selected = again;
        if (paneChart) { paneChart = again; openPane(again); }
        if (flyout) queueMicrotask(() => { const v = views.get(again); if (v && selected === again) buttons.show(flyout, again, v.el); });
      } else {
        selected = null;
        if (paneChart) { paneChart = null; app.taskPane.close(); }
      }
      queueMicrotask(() => app.chartSelectionChanged());
    }
    for (const [m, v] of views) v.el.classList.toggle("sel", m === selected);
    // Drawing order: later charts on top.
    list.forEach((m, i) => { views.get(m)!.el.style.zIndex = String(i + 1); });
    place();
  }

  // ----- the source ranges in the sheet -----------------------------------------------------------

  interface RangeBox { kind: "name" | "cat" | "val"; rect: Rect; series: number[] }

  /** The ranges of the selected chart (or its selected series) on the shown sheet, merged where series line up. */
  function sourceRanges(m: ChartModel): RangeBox[] {
    const only = element.kind === "series" ? element.index : null;
    const out: RangeBox[] = [];
    const add = (kind: RangeBox["kind"], ref: string | undefined, i: number) => {
      const r = ref ? parseRef(ref) : null;
      if (!r || r.sheet !== app.sheet.name) return;
      const rect = { r0: r.r0, c0: r.c0, r1: r.r1, c1: r.c1 };
      // A block of neighbouring series' values (or names) makes one frame, as Excel draws it.
      const prev = out.find((b) => b.kind === kind && (kind === "cat" ? JSON.stringify(b.rect) === JSON.stringify(rect)
        : (b.rect.r0 === rect.r0 && b.rect.r1 === rect.r1 && b.rect.c1 + 1 === rect.c0) || (b.rect.c0 === rect.c0 && b.rect.c1 === rect.c1 && b.rect.r1 + 1 === rect.r0)));
      if (prev) { if (kind !== "cat") { prev.rect.r1 = Math.max(prev.rect.r1, rect.r1); prev.rect.c1 = Math.max(prev.rect.c1, rect.c1); } prev.series.push(i); return; }
      out.push({ kind, rect, series: [i] });
    };
    m.series.forEach((s, i) => {
      if (only !== null && i !== only) return;
      add("name", s.name?.ref, i);
      add("cat", s.cat?.ref, i);
      add("val", s.val.ref, i);
    });
    return out;
  }

  let rangeDrag: { m: ChartModel; box: RangeBox; mode: "move" | "nw" | "ne" | "sw" | "se"; start: { row: number; col: number }; draft: Rect } | null = null;

  function showRanges(lx: number, ly: number): void {
    if (!rangesEl) return;
    const boxes: { rect: Rect; color: string; box?: RangeBox }[] = [];
    if (selected && views.has(selected) && !looks.has(selected)) {
      for (const b of sourceRanges(selected)) boxes.push({ rect: rangeDrag?.box === b ? rangeDrag.draft : b.rect, color: RANGE_COLORS[b.kind], box: b });
      if (rangeDrag) {
        const i = boxes.findIndex((x) => x.box && x.box.kind === rangeDrag!.box.kind && JSON.stringify(x.box.series) === JSON.stringify(rangeDrag!.box.series));
        if (i >= 0) boxes[i].rect = rangeDrag.draft;
      }
    }
    if (pickRect) boxes.push({ rect: pickRect, color: "#107c41" });
    const els: HTMLElement[] = [];
    for (const b of boxes) {
      const r = app.grid.rectOnCanvas(b.rect);
      if (!r) continue;
      const el = h("div", { class: `chart-range${b.box ? "" : " pick"}`, style: `left:${r.x - lx}px;top:${r.y - ly}px;width:${r.w + 1}px;height:${r.h + 1}px;${b.box ? `border-color:${b.color}` : ""}`, "data-kind": b.box?.kind ?? "pick" });
      if (b.box) {
        const box = b.box;
        for (const e of ["t", "b", "l", "r"]) el.append(h("div", { class: `edge ${e}` }));
        for (const c of ["nw", "ne", "sw", "se"]) el.append(h("div", { class: `corner ${c}`, style: `background:${b.color}`, "data-corner": c }));
        el.addEventListener("pointerdown", (e) => startRangeDrag(e, box));
      }
      els.push(el);
    }
    rangesEl.replaceChildren(...els);
  }

  function cellAtClient(x: number, y: number): { row: number; col: number } {
    const g = app.grid.el.getBoundingClientRect();
    return app.grid.hitCell(x - g.left, y - g.top);
  }

  function startRangeDrag(e: PointerEvent, box: RangeBox): void {
    const m = selected;
    if (!m || e.button !== 0) return;
    e.preventDefault();
    e.stopPropagation();
    const corner = (e.target as HTMLElement).dataset.corner as "nw" | "ne" | "sw" | "se" | undefined;
    rangeDrag = { m, box, mode: corner ?? "move", start: cellAtClient(e.clientX, e.clientY), draft: { ...box.rect } };
    const move = (ev: PointerEvent) => {
      if (!rangeDrag) return;
      const c = cellAtClient(ev.clientX, ev.clientY);
      const r = box.rect;
      if (rangeDrag.mode === "move") {
        const dr = c.row - rangeDrag.start.row, dc = c.col - rangeDrag.start.col;
        rangeDrag.draft = { r0: Math.max(0, r.r0 + dr), c0: Math.max(0, r.c0 + dc), r1: Math.max(0, r.r1 + dr), c1: Math.max(0, r.c1 + dc) };
      } else {
        const k = rangeDrag.mode;
        // The opposite corner stays.
        const fixed = { row: k.includes("n") ? r.r1 : r.r0, col: k.includes("w") ? r.c1 : r.c0 };
        rangeDrag.draft = { r0: Math.min(fixed.row, c.row), r1: Math.max(fixed.row, c.row), c0: Math.min(fixed.col, c.col), c1: Math.max(fixed.col, c.col) };
      }
      place();
    };
    const up = () => {
      removeEventListener("pointermove", move, true);
      removeEventListener("pointerup", up, true);
      const d = rangeDrag;
      rangeDrag = null;
      if (d && JSON.stringify(d.draft) !== JSON.stringify(box.rect)) applyRangeDrag(d.m, box, d.draft, d.mode === "move");
      else place();
    };
    addEventListener("pointermove", move, true);
    addEventListener("pointerup", up, true);
  }

  /** A dragged frame: moved, the series' references move along; sized, the chart's data block grows or shrinks (Excel). */
  function applyRangeDrag(m: ChartModel, box: RangeBox, to: Rect, moved: boolean): void {
    const sheetName = app.sheet.name;
    const src = sourceOf(m.series);
    if (!moved && src && box.kind === "val" && box.series.length === m.series.length && src.ref.sheet === sheetName) {
      // The whole chart's values sized: the block keeps its name row and category column, the values end where dragged.
      const v = box.rect;
      const block: Ref = { sheet: sheetName, r0: src.ref.r0 + (to.r0 - v.r0), c0: src.ref.c0 + (to.c0 - v.c0), r1: src.ref.r1 + (to.r1 - v.r1), c1: src.ref.c1 + (to.c1 - v.c1) };
      if (block.r1 < block.r0 || block.c1 < block.c0) { place(); return; }
      edit(m, (c) => { c.series = seriesOf(c, block, src.orientation); delete c.raw; delete c.hiddenCats; readCaches(app.wb, c); }, "Datenquelle ändern");
      return;
    }
    // Otherwise each series' range of that kind becomes its part of the new frame.
    const dr0 = to.r0 - box.rect.r0, dc0 = to.c0 - box.rect.c0;
    edit(m, (c) => {
      for (const i of box.series) {
        const s = c.series[i];
        const d = box.kind === "name" ? s.name : box.kind === "cat" ? s.cat : s.val;
        const r = d?.ref ? parseRef(d.ref) : null;
        if (!d || !r) continue;
        let nr: Rect;
        if (moved) nr = { r0: r.r0 + dr0, c0: r.c0 + dc0, r1: r.r1 + dr0, c1: r.c1 + dc0 };
        else if (box.series.length === 1 || box.kind === "cat") nr = to;
        else {
          // A block of several series sized: each keeps its row/column, the length follows.
          const vertical = r.c0 === r.c1 && box.rect.c1 > box.rect.c0;
          nr = vertical ? { r0: to.r0, r1: to.r1, c0: r.c0, c1: r.c1 } : { r0: r.r0, r1: r.r1, c0: to.c0, c1: to.c1 };
        }
        d.ref = refText(sheetName, nr.r0, nr.c0, nr.r1, nr.c1);
      }
      delete c.raw;
      delete c.hiddenCats;
      readCaches(app.wb, c);
    }, "Datenquelle ändern");
  }

  // ----- selection, moving and sizing ------------------------------------------------------------------

  function select(m: ChartModel | null): void {
    queueMicrotask(() => app.chartSelectionChanged());
    if (m === selected) return;
    selected = m;
    element = { kind: "chart" };
    if (!m && app.taskPane.isOpen && paneChart) { paneChart = null; app.taskPane.close(); }
    for (const [k, v] of views) v.el.classList.toggle("sel", k === m);
    if (!m) buttons.close();
    else views.get(m)?.el.focus({ preventScroll: true });
    place();
  }

  function onDown(e: PointerEvent, m: ChartModel): void {
    if (e.button !== 0) return;
    if ((e.target as HTMLElement).closest?.(".chart-btn")) return;
    e.stopPropagation();
    e.preventDefault();
    if (app.grid.isEditing) app.grid.commitEdit("none");
    buttons.close();
    select(m);
    // A click on an element selects it; a click on the chart's handles keeps the chart area.
    if (!(e.target as HTMLElement).closest?.(".chart-handle")) setElement(m, elementAt(m, e.target as Element, e) ?? { kind: "chart" });
    const view = views.get(m)!;
    view.el.focus({ preventScroll: true });
    const handle = ((e.target as HTMLElement).closest?.(".chart-handle") as HTMLElement | null)?.dataset.h as Handle | undefined;
    const box = sheetBox(m);
    drag = { m, handle: handle ?? null, x: e.clientX, y: e.clientY, box, draft: { ...box }, moved: false };
    try { view.el.setPointerCapture(e.pointerId); } catch { /* synthetic events */ }
  }

  /** The nearest cell edge to a place in sheet px at 100 % (Alt, Excel). */
  function gridEdge(v: number, axis: "x" | "y"): number {
    const { rows, cols } = getAxes();
    const a = axis === "x" ? cols : rows;
    const i = a.at(Math.max(0, v)), lo = a.pos(i), hi = a.pos(i + 1);
    return v - lo < (hi - lo) / 2 ? lo : hi;
  }

  function replay(e: KeyboardEvent): void {
    if (!drag?.last) return;
    onMove({ ...drag.last, shiftKey: e.shiftKey, ctrlKey: e.ctrlKey, altKey: e.altKey });
  }

  /** The drag ends without effect: the pointer was taken away, or a move came without the button (D-127). */
  function dropDrag(): void {
    if (!drag) return;
    views.get(drag.m)?.el.classList.remove("dragging");
    drag = null;
    place();
  }

  function onMove(e: DragPoint): void {
    if (!drag) return;
    if (e instanceof PointerEvent && !e.buttons && e.pointerType === "mouse") { dropDrag(); return; }
    drag.last = { clientX: e.clientX, clientY: e.clientY, shiftKey: e.shiftKey, ctrlKey: e.ctrlKey, altKey: e.altKey };
    const z = zoom();
    let dx = (e.clientX - drag.x) / z, dy = (e.clientY - drag.y) / z;
    if (!drag.moved && Math.abs(e.clientX - drag.x) + Math.abs(e.clientY - drag.y) < 3) return;
    drag.moved = true;
    const b = drag.box;
    let { x, y, w, h: hh } = b;
    const k = drag.handle;
    if (!k) {
      // Umschalt: along one axis; Alt: the corner to the cell grid (Excel; D-128).
      if (e.shiftKey) { if (Math.abs(dx) >= Math.abs(dy)) dy = 0; else dx = 0; }
      x = Math.max(0, b.x + dx);
      y = Math.max(0, b.y + dy);
      if (e.altKey) { x = gridEdge(x, "x"); y = gridEdge(y, "y"); }
      views.get(drag.m)?.el.classList.add("dragging");
    } else {
      if (k.includes("e")) w = b.w + dx;
      if (k.includes("s")) hh = b.h + dy;
      if (k.includes("w")) { w = b.w - dx; x = b.x + dx; }
      if (k.includes("n")) { hh = b.h - dy; y = b.y + dy; }
      if (e.altKey) {
        if (k.includes("e")) w = gridEdge(x + w, "x") - x;
        if (k.includes("s")) hh = gridEdge(y + hh, "y") - y;
        if (k.includes("w")) { const nx = gridEdge(x, "x"); w += x - nx; x = nx; }
        if (k.includes("n")) { const ny = gridEdge(y, "y"); hh += y - ny; y = ny; }
      }
      // Umschalt on a corner keeps the proportions (Excel).
      if (e.shiftKey && k.length === 2) {
        const s = Math.max(w / b.w, hh / b.h);
        w = b.w * s;
        hh = b.h * s;
        if (k.includes("w")) x = b.x + b.w - w;
        if (k.includes("n")) y = b.y + b.h - hh;
      }
      if (e.ctrlKey) {
        // Strg: the size changes around the middle (Excel; D-128).
        const gw = w - b.w, gh = hh - b.h;
        w = b.w + 2 * gw; hh = b.h + 2 * gh; x = b.x - gw; y = b.y - gh;
      }
      if (w < MIN_SIZE) { if (k.includes("w")) x -= MIN_SIZE - w; w = MIN_SIZE; }
      if (hh < MIN_SIZE) { if (k.includes("n")) y -= MIN_SIZE - hh; hh = MIN_SIZE; }
      x = Math.max(0, x);
      y = Math.max(0, y);
    }
    drag.draft = { x, y, w, h: hh };
    place();
  }

  function onUp(e: PointerEvent): void {
    if (!drag) return;
    const d = drag;
    drag = null;
    const view = views.get(d.m);
    view?.el.classList.remove("dragging");
    try { if (view?.el.hasPointerCapture(e.pointerId)) view.el.releasePointerCapture(e.pointerId); } catch { /* synthetic events */ }
    if (!d.moved) { place(); return; }
    const sheet = app.sheet;
    if (e.ctrlKey && !d.handle) {
      // Strg held when the button goes up: a copy goes there, the chart stays (Excel; D-128).
      const copy = structuredClone({ ...d.m, raw: undefined }) as ChartModel;
      if (d.m.raw !== undefined) copy.raw = d.m.raw;
      copy.from = anchorAt(d.draft.x, d.draft.y);
      copy.to = anchorAt(d.draft.x + d.draft.w, d.draft.y + d.draft.h);
      copy.name = `Diagramm ${nextNumber(app.wb)}`;
      app.change("Kopieren", () => {
        app.history.charts(sheet);
        sheet.charts = [...(sheet.charts ?? []), copy];
      });
      refresh();
      select(copy);
      return;
    }
    app.change(d.handle ? "Größe ändern" : "Verschieben", () => {
      app.history.charts(sheet);
      d.m.from = anchorAt(d.draft.x, d.draft.y);
      d.m.to = anchorAt(d.draft.x + d.draft.w, d.draft.y + d.draft.h);
    });
    // The undo step copies the models; keep the moved chart selected.
    const again = (app.sheet.charts ?? []).find((c) => c === d.m);
    if (again) select(again);
  }

  function onKey(e: KeyboardEvent, m: ChartModel): void {
    if (e.key === "Escape" && drag) {
      // Esc while moving or sizing: the chart goes back to where it was (Excel; D-096).
      e.preventDefault();
      e.stopPropagation();
      views.get(drag.m)?.el.classList.remove("dragging");
      drag = null;
      place();
      return;
    }
    if (e.ctrlKey && (e.key === "1" || e.code === "Digit1")) {
      // Strg+1: the Format pane of the selected element (Excel).
      e.preventDefault();
      e.stopPropagation();
      openPane(m);
      return;
    }
    if (e.key === "Delete" || e.key === "Backspace") {
      e.preventDefault();
      e.stopPropagation();
      // Entf on an element deletes the element (Excel), on the chart area the chart.
      if (element.kind !== "chart" && element.kind !== "plot") deleteElement(m, element);
      else remove(m);
    } else if (drag && (e.key === "Shift" || e.key === "Control" || e.key === "Alt")) {
      e.preventDefault();
      replay(e);
    } else if (e.key === "Escape") {
      e.preventDefault();
      e.stopPropagation();
      if (buttons.isOpen()) buttons.close();
      else if (element.kind !== "chart") setElement(m, { kind: "chart" });
      else { select(null); app.grid.focus(); }
    } else if (e.key.startsWith("Arrow")) {
      e.preventDefault();
      e.stopPropagation();
    }
  }

  function remove(m: ChartModel): void {
    const sheet = app.sheet;
    select(null);
    app.change("Löschen", () => {
      app.history.charts(sheet);
      sheet.charts = (sheet.charts ?? []).filter((c) => c !== m);
    });
    app.grid.focus();
  }

  /** Löschen of an element (context menu, Entf): the element goes, as Excel removes it. */
  function deleteElement(m: ChartModel, e: ChartElement): void {
    edit(m, (c) => {
      const f = (c.format ??= {});
      switch (e.kind) {
        case "title": c.title.show = false; break;
        case "legend": c.legend = "none"; break;
        case "valAxis": f.valAxis = { ...(f.valAxis ?? {}), visible: false }; break;
        case "catAxis": f.catAxis = { ...(f.catAxis ?? {}), visible: false }; break;
        case "majorGrid": c.gridlines = false; break;
        case "minorGrid": delete f.minorGrid; break;
        case "catGrid": delete f.catGrid; break;
        case "axisTitle": delete c.axisTitles[e.axis]; break;
        case "labels": c.dataLabels = false; break;
        case "trend": delete c.series[e.index].trendline; break;
        case "errorBars": delete c.series[e.index].errorBars; break;
        case "series": if (c.series.length > 1) c.series.splice(e.index, 1); delete c.raw; break;
        default: break;
      }
      if (!Object.keys(f).length) delete c.format;
    }, "Löschen");
    setElement(m, { kind: "chart" });
  }

  // ----- elements and the Format pane --------------------------------------------------------------

  /**
   * The element under the pointer, from the drawing's data-part marks (series by their index in the full chart;
   * a second click on a selected series: its point, D-111).
   */
  function elementAt(m: ChartModel, target: Element, at?: { clientX: number; clientY: number }): ChartElement | null {
    const part = (target.closest?.("[data-part]") as SVGElement | null)?.dataset.part;
    if (!part) return null;
    const [kind, a, b] = part.split(":");
    // The chart area's background inside the plot's box is the Zeichnungsfläche (Excel).
    if (kind === "chart" && at) {
      const svg = target.closest("svg") as SVGSVGElement | null;
      const box = svg?.dataset.plot?.split(",").map(Number);
      const vb = svg?.viewBox.baseVal;
      if (svg && box?.length === 4 && vb && vb.width) {
        const r = svg.getBoundingClientRect();
        const x = ((at.clientX - r.left) / r.width) * vb.width, y = ((at.clientY - r.top) / r.height) * vb.height;
        if (x >= box[0] && x <= box[0] + box[2] && y >= box[1] && y <= box[1] + box[3]) return { kind: "plot" };
      }
    }
    const full = (i: number) => visibleModel(m).index[i] ?? i;
    const indexed = ["series", "trend", "trendLabel", "errbars"].includes(kind) && a !== undefined;
    return elementOfPart(indexed ? [kind, full(Number(a)), ...(b !== undefined ? [b] : [])].join(":") : part, element);
  }

  /** The SVG parts that make up an element (for its selection frame). */
  function partsOf(m: ChartModel, e: ChartElement): string {
    const shown = (i: number) => visibleModel(m).index.indexOf(i);
    return partsSelector("index" in e ? { ...e, index: shown(e.index) } as ChartElement : e);
  }

  function setElement(m: ChartModel, e: ChartElement): void {
    element = e;
    showElement(m);
    if (paneChart === m && app.taskPane.isOpen) openPane(m);
    queueMicrotask(() => app.chartSelectionChanged());
    place();
  }

  /** The selection marks as Excel draws them: the chart's frame and sizing handles, and the selected element's (D-170). */
  function showElement(m: ChartModel): void {
    const v = views.get(m);
    const svg = v?.el.querySelector<SVGSVGElement>(":scope > svg:not(.chart-marks)");
    if (!v || !svg || m !== selected) return;
    const o = v.el.getBoundingClientRect();
    const own = elementMarks(svg, element, partsOf(m, element), o, { pie: isPie(m.kind) });
    drawMarks(marksLayer(v.el), [chartMarks(o.width, o.height), own]);
  }

  let paneChart: ChartModel | null = null;

  /** Opens (or follows) the Format pane of the selected element. */
  function openPane(m: ChartModel): void {
    paneChart = m;
    const valid = chartElements(m).some((x) => JSON.stringify(x) === JSON.stringify(element));
    if (!valid) element = { kind: "chart" };
    app.taskPane.onClose = () => { paneChart = null; };
    // Below the formula bar, as Excel docks its panes.
    document.documentElement.style.setProperty("--pane-top", `${document.querySelector(".calc-main")?.getBoundingClientRect().top ?? 150}px`);
    app.taskPane.show(formatPane(m, element, {
      edit: (fn) => { edit(m, fn); showElement(m); },
      select: (x) => setElement(m, x),
      size: () => { const b = sheetBox(m); return { w: Math.round((b.w / 96) * 2.54 * 100) / 100, h: Math.round((b.h / 96) * 2.54 * 100) / 100 }; },
      resize: (w, hh) => {
        const b = sheetBox(m);
        const to = anchorAt(b.x + (w / 2.54) * 96, b.y + (hh / 2.54) * 96);
        edit(m, (c) => { c.to = to; });
      },
      data: selectHost,
    }));
  }

  /** "Weitere Optionen...": the Format pane of an element (the pane opens with the element selected). */
  function paneFor(m: ChartModel, e: ChartElement): void {
    select(m);
    const valid = chartElements(m).some((x) => JSON.stringify(x) === JSON.stringify(e));
    element = valid ? e : { kind: "chart" };
    showElement(m);
    openPane(m);
  }

  /** Changes the selected chart as one undo step and redraws the flyout. */
  function edit(m: ChartModel, fn: (m: ChartModel) => void, label = "Diagrammelemente"): void {
    const sheet = app.sheet;
    app.change(label, () => {
      app.history.charts(sheet);
      fn(m);
    });
    if (buttons.isOpen()) buttons.refresh(m);
    if (paneChart === m && app.taskPane.isOpen) app.taskPane.refresh();
    place();
  }

  // ----- the buttons beside the chart ------------------------------------------------------------

  const buttons = chartButtons({
    edit: (m, fn, label) => edit(m, fn, label),
    preview: (m, look) => {
      if (look) { const l = structuredClone({ ...look, raw: undefined }); readCaches(app.wb, l); looks.set(m, l); } else looks.delete(m);
      place();
    },
    highlight: (m, what) => { fade = what ? { m, ...what } : null; place(); },
    openPane: (m, e) => paneFor(m, e),
    selectData: (m) => { select(m); void selectData(); },
    accents: () => accents() ?? ["#4472C4", "#ED7D31", "#A5A5A5", "#FFC000", "#5B9BD5", "#70AD47"],
    tile: (m, w, hh) => tileOf(m, w, hh),
    pickSeries: (m) => pickSeries(m),
    nameSources: (m) => nameSources(m),
    setNameSources: (m, series, cats) => setNameSources(m, series, cats),
  });

  /** A gallery tile: the chart drawn at its own size and scaled down (Excel's tiles are small pictures of the chart). */
  function tileOf(m: ChartModel, w: number, hh: number): SVGSVGElement {
    readCaches(app.wb, m);
    const sel = selected && views.has(selected) ? sheetBox(selected) : { w: NEW_W, h: NEW_H };
    const svg = renderChart(m, sel.w, sel.h, { accents: accents(), format }) as SVGSVGElement;
    const s = Math.min(w / sel.w, hh / sel.h);
    svg.setAttribute("width", String(Math.round(sel.w * s)));
    svg.setAttribute("height", String(Math.round(sel.h * s)));
    return svg;
  }

  /** Trendlinie hinzufügen: Excel asks which series when there are several. */
  async function pickSeries(m: ChartModel): Promise<number | null> {
    const list = h("select", { size: 6, style: "width:260px", "aria-label": "Reihen" }) as HTMLSelectElement;
    m.series.forEach((s, i) => { if (!s.filtered) list.append(h("option", { value: String(i) }, seriesName(s, i))); });
    list.selectedIndex = 0;
    const ok = await showDialog({ title: "Trendlinie hinzufügen", body: h("div", null, h("div", { style: "margin-bottom:6px" }, "Trendlinie hinzufügen basierend auf Reihe:"), list),
      buttons: [{ label: "OK", value: true, primary: true }, { label: "Abbrechen", value: false }], cancel: false });
    return ok && list.value !== "" ? Number(list.value) : null;
  }

  /** Diagrammfilter > Namen: the row (column) the series names come from and the column (row) of the categories. */
  function nameSources(m: ChartModel) {
    const nameRef = m.series.map((s) => s.name?.ref ? parseRef(s.name.ref) : null);
    const catRef = m.series[0]?.cat?.ref ? parseRef(m.series[0].cat.ref) : null;
    const colName = (c: number) => refText("x", 0, c).replace(/^x!\$|\$\d+$/g, "");
    let series: string | null = null;
    const n0 = nameRef[0];
    if (n0 && nameRef.every((r) => r)) series = nameRef.every((r) => r!.r0 === n0.r0) ? `Zeile ${n0.r0 + 1}` : nameRef.every((r) => r!.c0 === n0.c0) ? `Spalte ${colName(n0.c0)}` : null;
    const cats = catRef ? (catRef.c0 === catRef.c1 ? `Spalte ${colName(catRef.c0)}` : `Zeile ${catRef.r0 + 1}`) : null;
    return {
      series, cats,
      seriesList: m.series.map((s, i) => seriesName(s, i)).join("; "),
      catList: (m.series[0]?.cat?.cache ?? []).map((v) => (v === null ? "" : String(v))).join("; "),
    };
  }

  function setNameSources(m: ChartModel, series: boolean, cats: boolean): void {
    const src = sourceOf(m.series);
    edit(m, (c) => {
      if (src) {
        const fresh = seriesOf(c, src.ref, src.orientation);
        c.series.forEach((s, i) => {
          if (series) { if (fresh[i]?.name) s.name = fresh[i].name; } else delete s.name;
          if (cats) { if (fresh[i]?.cat) s.cat = fresh[i].cat; } else delete s.cat;
        });
      } else if (!series || !cats) c.series.forEach((s) => { if (!series) delete s.name; if (!cats) delete s.cat; });
      delete c.raw;
      readCaches(app.wb, c);
    }, "Diagrammfilter");
  }

  // ----- context menus ----------------------------------------------------------------------------------

  function contextMenu(m: ChartModel, e: ChartElement, x: number, y: number): void {
    const labels: Record<string, string> = { "chart.changeType": "Diagrammtyp ändern...", "chart.selectData": "Daten auswählen...", "chart.move": "Diagramm verschieben..." };
    const cmd = (id: string) => ({ cmd: id, label: labels[id] });
    const reset = { label: "Auf Formatvorl. zurücks.", run: () => resetElement(m, e) };
    const font = { label: "Schriftart...", run: () => openPane(m) };
    const common: MenuEntry[] = [cmd("chart.changeType"), cmd("chart.selectData"), { label: "3D-Drehung...", run: () => {}, disabled: true }];
    const del = { label: "Löschen", run: () => deleteElement(m, e) };
    const fmt = (label: string) => ({ label, run: () => openPane(m) });
    let entries: MenuEntry[];
    switch (e.kind) {
      case "title":
        entries = [del, reset, "-", { label: "Text bearbeiten", run: () => void editTitle(m) }, font, "-", ...common, "-", fmt("Diagrammtitel formatieren...")];
        break;
      case "legend":
        entries = [del, reset, "-", font, "-", ...common, "-", fmt("Legende formatieren...")];
        break;
      case "valAxis": case "catAxis": {
        const minorKey = e.kind === "valAxis" ? "minorGrid" : "catMinorGrid";
        const hasMinor = !!m.format?.[minorKey]?.show;
        const hasMajor = e.kind === "valAxis" ? m.gridlines : !!m.format?.catGrid?.show;
        entries = [del, reset, "-", font, "-", ...common, "-",
          hasMajor ? { label: "Hauptgitternetz formatieren...", run: () => paneFor(m, e.kind === "valAxis" ? { kind: "majorGrid" } : { kind: "catGrid" }) }
            : { label: "Hauptgitternetz hinzufügen", run: () => edit(m, (c) => { if (e.kind === "valAxis") c.gridlines = true; else { c.format ??= {}; c.format.catGrid = { ...(c.format.catGrid ?? {}), show: true }; } }) },
          hasMinor ? { label: "Hilfsgitternetz formatieren...", run: () => paneFor(m, e.kind === "valAxis" ? { kind: "minorGrid" } : { kind: "catGrid" }) }
            : { label: "Hilfsgitternetz hinzufügen", run: () => edit(m, (c) => { c.format ??= {}; c.format[minorKey] = { ...(c.format[minorKey] ?? {}), show: true }; }) },
          fmt("Achse formatieren...")];
        break;
      }
      case "majorGrid": case "minorGrid": case "catGrid":
        entries = [del, reset, "-", ...common, "-", fmt("Gitternetzlinien formatieren..."),
          { label: "Achse formatieren...", run: () => paneFor(m, e.kind === "catGrid" ? { kind: "catAxis" } : { kind: "valAxis" }) }];
        break;
      case "series": {
        const s = m.series[e.index];
        const labelsOn = m.dataLabels;
        entries = [del, reset, "-", { label: "Datenreihen-Diagrammtyp ändern...", run: () => void app.commands.run("chart.changeType") }, cmd("chart.selectData"),
          { label: "3D-Drehung...", run: () => {}, disabled: true }, "-",
          labelsOn ? { label: "Datenbeschriftungen formatieren...", run: () => paneFor(m, { kind: "labels" }) }
            : { label: "Datenbeschriftungen hinzufügen", submenu: [
              { label: "Datenbeschriftungen hinzufügen", run: () => edit(m, (c) => { c.dataLabels = true; }) },
              { label: "Datenlegenden hinzufügen", run: () => edit(m, (c) => { c.dataLabels = true; c.format ??= {}; c.format.dataLabels = { ...(c.format.dataLabels ?? {}), showCat: true }; }) }] },
          s?.trendline ? { label: "Trendlinie formatieren...", run: () => paneFor(m, { kind: "trend", index: e.index }) }
            : { label: "Trendlinie hinzufügen...", run: () => { edit(m, (c) => { c.series[e.index].trendline = { type: "linear", showEq: false, showR2: false }; }); paneFor(m, { kind: "trend", index: e.index }); } },
          fmt("Datenreihen formatieren...")];
        break;
      }
      case "plot":
        entries = [del, reset, "-", cmd("chart.changeType"), { label: "Als Vorlage speichern...", run: () => {}, disabled: true }, cmd("chart.selectData"),
          { label: "3D-Drehung...", run: () => {}, disabled: true }, "-", fmt("Zeichnungsfläche formatieren...")];
        break;
      case "labels":
        entries = [del, reset, "-", font, "-", ...common, "-", fmt("Datenbeschriftungen formatieren...")];
        break;
      case "axisTitle":
        entries = [del, reset, "-", font, "-", ...common, "-", fmt("Achsentitel formatieren...")];
        break;
      case "trend":
        entries = [del, reset, "-", fmt("Trendlinie formatieren...")];
        break;
      case "errorBars":
        entries = [del, reset, "-", fmt("Fehlerindikatoren formatieren...")];
        break;
      default:
        // The chart area: Excel's object menu.
        entries = [cmd("edit.cut"), cmd("edit.copy"), { label: "Inhalte einfügen...", run: () => {}, disabled: true }, "-", reset, font, cmd("chart.changeType"),
          { label: "Als Vorlage speichern...", run: () => {}, disabled: true }, cmd("chart.selectData"), cmd("chart.move"),
          { label: "3D-Drehung...", run: () => {}, disabled: true }, "-", { label: "Gruppieren", run: () => {}, disabled: true },
          // Greyed out while the chart is the sheet's only chart (Excel).
          { label: "In den Vordergrund", run: () => arrange(m, "front"), disabled: (app.sheet.charts ?? []).length < 2 },
          { label: "In den Hintergrund", run: () => arrange(m, "back"), disabled: (app.sheet.charts ?? []).length < 2 }, "-",
          { label: "Als Grafik speichern...", run: () => {}, disabled: true }, { label: "Makro zuweisen...", run: () => {}, disabled: true },
          { label: "Alternativtext anzeigen...", run: () => {}, disabled: true },
          { label: "Diagrammbereich formatieren...", run: () => paneFor(m, { kind: "chart" }) }, { label: "PivotChart-Optionen...", run: () => {}, disabled: true }];
    }
    showMenu(app.commands, entries, new DOMRect(x, y, 0, 0));
  }

  async function editTitle(m: ChartModel): Promise<void> {
    const input = h("input", { type: "text", value: titleText(m), style: "width:300px", "aria-label": "Diagrammtitel" }) as HTMLInputElement;
    const ok = await showDialog({ title: "Diagrammtitel", body: h("label", { class: "field" }, "Text:", input), buttons: [{ label: "OK", value: true, primary: true }, { label: "Abbrechen", value: false }], cancel: false });
    if (ok) edit(m, (c) => { c.title.show = true; c.title.text = input.value; });
  }

  /** In den Vordergrund / In den Hintergrund (and Ebene nach vorne / hinten): the chart's place in the sheet's drawing order. */
  function arrange(m: ChartModel, how: "front" | "back" | "forward" | "backward"): void {
    const sheet = app.sheet;
    const list = sheet.charts ?? [];
    const i = list.indexOf(m);
    if (i < 0) return;
    const j = how === "front" ? list.length - 1 : how === "back" ? 0 : how === "forward" ? Math.min(list.length - 1, i + 1) : Math.max(0, i - 1);
    if (i === j) return;
    app.change("Anordnen", () => {
      app.history.charts(sheet);
      const next = list.filter((c) => c !== m);
      next.splice(j, 0, m);
      sheet.charts = next;
    });
    refresh();
    select(m);
  }

  // ----- inserting ----------------------------------------------------------------------------------------

  function insert(kind: ChartKind, tweak?: (m: ChartModel) => void): void {
    if (app.grid.isEditing) app.grid.commitEdit("none");
    ensureLayer();
    const sheet = app.sheet;
    axes = null;
    const rect = sourceRect();
    const series = rect
      ? seriesFromBlock({ sheet: sheet.name, ...rect, get: (r, c) => cellValue(sheet.value(r, c)) }, kind)
      : [];
    // Excel puts a new chart in the middle of the visible cells.
    const vis = visibleBox();
    const x = vis.x + Math.max(0, (vis.w - NEW_W) / 2), y = vis.y + Math.max(0, (vis.h - NEW_H) / 2);
    const m = newChart(kind, series, anchorAt(x, y), anchorAt(x + NEW_W, y + NEW_H));
    m.name = `Diagramm ${nextNumber(app.wb)}`;
    tweak?.(m);
    readCaches(app.wb, m);
    app.change("Diagramm einfügen", () => {
      app.history.charts(sheet);
      sheet.charts = [...(sheet.charts ?? []), m];
    });
    refresh();
    select(m);
  }

  // ----- Schnelllayout ------------------------------------------------------------------------------

  /** The recorded layouts that fit the chart's kind (Excel shows the same eleven for each family). */
  function layoutsFor(m: ChartModel) {
    const fam: Record<string, string> = { column: "column", columnStacked: "column", bar: "bar", barStacked: "bar", line: "line", lineMarkers: "lineMarkers",
      area: "area", pie: "pie", doughnut: "pie", scatter: "scatter", scatterSmooth: "scatter", scatterLines: "scatter" };
    return QUICK_LAYOUTS[fam[m.kind] ?? ""] ?? [];
  }

  /** Diagrammentwurf > Schnelllayout > Layout n: the elements Excel's layout sets (recorded). */
  function quickLayout(n: number, target?: ChartModel): void {
    const m = target ?? selected;
    const l = m ? layoutsFor(m)[n - 1] : null;
    if (!m || !l) return;
    const apply = (c: ChartModel) => {
      c.title.show = l.title;
      c.legend = l.legend;
      c.dataLabels = l.labels;
      const f = (c.format ??= {});
      if (l.labels) {
        const d: NonNullable<typeof f.dataLabels> = { ...(f.dataLabels ?? {}) };
        delete d.pos; delete d.showCat; delete d.showPercent;
        const defPos = isPie(c.kind) ? "bestFit" : c.kind.endsWith("Stacked") ? "ctr" : c.kind === "column" || c.kind === "bar" ? "outEnd" : "r";
        if (l.labelPos && l.labelPos !== defPos) d.pos = l.labelPos as import("@raabe/charts").DataLabelOptions["pos"];
        if (l.showCat) d.showCat = true;
        if (l.showPct) d.showPercent = true;
        if (Object.keys(d).length) f.dataLabels = d; else delete f.dataLabels;
      }
      c.axisTitles = { ...(l.catTitle ? { cat: c.axisTitles.cat ?? "" } : {}), ...(l.valTitle ? { val: c.axisTitles.val ?? "" } : {}) };
      c.gridlines = !!l.majorGrid;
      if (l.minorGrid) f.minorGrid = { ...(f.minorGrid ?? {}), show: true }; else delete f.minorGrid;
      if (l.catGrid) f.catGrid = { ...(f.catGrid ?? {}), show: true }; else delete f.catGrid;
      if (l.dataTable) c.dataTable = { legendKeys: true }; else delete c.dataTable;
      const defGap = c.kind.endsWith("Stacked") ? 150 : c.kind === "bar" ? 182 : 219;
      if (l.gap !== undefined && (c.kind === "column" || c.kind === "bar" || c.kind.endsWith("Stacked"))) c.gapWidth = l.gap === defGap ? undefined : l.gap;
      if (!Object.keys(f).length) delete c.format;
    };
    if (target) apply(target); else edit(m, apply, "Schnelllayout");
  }

  // ----- Daten: Zeile/Spalte tauschen, Daten auswählen -------------------------------------------

  /** The series of a range, as Excel makes them (`orientation` forced by Zeile/Spalte tauschen). */
  function seriesOf(m: ChartModel, ref: { sheet: string; r0: number; c0: number; r1: number; c1: number }, orientation?: "rows" | "cols"): ChartSeries[] {
    const sheet = app.wb.sheets.find((s) => s.name === ref.sheet) ?? app.sheet;
    return seriesFromBlock({ ...ref, sheet: sheet.name, get: (r, c) => cellValue(sheet.value(r, c)) }, m.kind, orientation);
  }

  /** Diagrammentwurf > Zeile/Spalte tauschen: the series run the other way through the same range. */
  function switchRowCol(): void {
    const m = selected;
    if (!m) return;
    const src = sourceOf(m.series);
    if (!src) return;
    edit(m, (c) => {
      c.series = seriesOf(c, src.ref, src.orientation === "cols" ? "rows" : "cols");
      delete c.raw;
      delete c.hiddenCats;
      readCaches(app.wb, c);
    }, "Zeile/Spalte tauschen");
  }

  const selectHost: SelectDataHost = {
    sheetName: () => app.sheet.name,
    seriesOf: (m, ref, o) => seriesOf(m, ref, o),
    read: (m) => readCaches(app.wb, m),
    preview: (ref) => {
      const r = parseRef(ref);
      const sheet = r ? app.wb.sheetByName(r.sheet) : undefined;
      if (!r || !sheet) return null;
      const out: string[] = [];
      for (let row = r.r0; row <= r.r1; row++) for (let col = r.c0; col <= r.c1; col++) {
        const cell = sheet.get(row, col);
        const v = cell?.v ?? null;
        out.push(v === null ? "" : typeof v === "number" ? formatValue(v, app.wb.styles.get(cell!.s).numFmt).text : String(v));
        if (out.length > 30) break;
      }
      return out.join("; ");
    },
    show: (m, draft) => {
      if (draft) { const l = structuredClone({ ...draft, raw: undefined }); readCaches(app.wb, l); looks.set(m, l); } else looks.delete(m);
      place();
    },
    commit: (m, draft) => edit(m, (c) => {
      c.series = draft.series;
      if (draft.hiddenCats?.length) c.hiddenCats = draft.hiddenCats; else delete c.hiddenCats;
      for (const k of ["dispBlanksAs", "showHidden", "naAsBlank"] as const) { if (draft[k] !== undefined) (c as unknown as Record<string, unknown>)[k] = draft[k]; else delete c[k]; }
      delete c.raw;
      readCaches(app.wb, c);
    }, "Daten auswählen"),
    grid: () => app.grid.el,
    cellAt: (x, y) => cellAtClient(x, y),
    marquee: (r) => { pickRect = r; place(); },
  };

  /** Diagrammentwurf > Daten auswählen: Excel's "Datenquelle auswählen" (selectdata.ts). */
  async function selectData(): Promise<void> {
    const m = selected;
    if (!m) return;
    buttons.close();
    await selectDataDialog(selectHost, m);
  }

  /** Diagrammentwurf > Diagramm verschieben: as an object onto another sheet (chart sheets: EXCEL-INVENTORY). */
  async function moveChart(): Promise<void> {
    const m = selected;
    if (!m) return;
    const sel = h("select", null, ...app.wb.sheets.map((s) => h("option", { value: s.name, selected: s === app.sheet }, s.name))) as HTMLSelectElement;
    const ok = await showDialog({ title: "Diagramm verschieben", body: h("div", null, h("p", null, "Wählen Sie, wo das Diagramm platziert werden soll:"), h("label", { class: "field" }, "Objekt in:", sel)),
      buttons: [{ label: "OK", value: true, primary: true }, { label: "Abbrechen", value: false }], cancel: false });
    if (!ok || sel.value === app.sheet.name) return;
    const to = app.wb.sheets.find((s) => s.name === sel.value)!;
    const from = app.sheet;
    select(null);
    app.change("Diagramm verschieben", () => {
      app.history.charts(from);
      app.history.charts(to);
      from.charts = (from.charts ?? []).filter((c) => c !== m);
      to.charts = [...(to.charts ?? []), m];
    });
    refresh();
  }

  function nextNumber(wb: Workbook): number {
    let n = 0;
    for (const s of wb.sheets) for (const c of s.charts ?? []) {
      const k = /(\d+)$/.exec(c.name ?? "");
      if (k) n = Math.max(n, Number(k[1]));
    }
    return n + 1;
  }

  /** Auf Formatvorlage zurücksetzen: the element's own formatting goes. */
  function resetElement(m: ChartModel, e: ChartElement): void {
    edit(m, (c) => {
      const f = c.format;
      const k = e.kind;
      if (k === "series") { const s = c.series[e.index]; delete s.fill; delete s.line; delete s.marker; return; }
      if (k === "chart") {
        // The chart area resets every element (Excel).
        delete c.format;
        for (const s of c.series) { delete s.fill; delete s.line; delete s.marker; }
        return;
      }
      if (!f) return;
      const key = ({ plot: "plot", title: "title", legend: "legend", valAxis: "valAxis", catAxis: "catAxis", majorGrid: "majorGrid", minorGrid: "minorGrid", catGrid: "catGrid", labels: "dataLabels" } as const)[k as "plot"];
      if (key) {
        const keep = (f as Record<string, { show?: boolean; overlay?: boolean; visible?: boolean } | undefined>)[key];
        const kept = keep ? { ...(keep.show !== undefined ? { show: keep.show } : {}), ...(keep.overlay ? { overlay: true } : {}), ...(keep.visible === false ? { visible: false } : {}) } : {};
        if (Object.keys(kept).length) (f as Record<string, unknown>)[key] = kept; else delete (f as Record<string, unknown>)[key];
      }
      if (!Object.keys(f).length) delete c.format;
    }, "Auf Formatvorlage zurücksetzen");
  }

  /** Diagrammentwurf > Diagrammelement hinzufügen: Excel's menu (each element's submenu with "Keine" and "Weitere ...optionen..."). */
  function elementsMenu(): MenuEntry[] {
    const m = selected;
    if (!m) return [];
    return elementEntries(m, true).map((e): MenuEntry => {
      if (e.disabled && !e.choices.length) return { label: e.label, submenu: [], disabled: true };
      return {
        label: e.label, disabled: e.disabled,
        submenu: () => {
          const cur = selected ?? m;
          const items: MenuEntry[] = [];
          if (e.none) items.push({ label: "Keine", checked: !e.on(cur), run: () => edit(cur, (c) => e.toggle(c, false)) });
          for (const ch of e.choices) {
            const on = ch.on(cur);
            items.push({
              label: choiceLabel(ch, true), checked: on,
              hover: () => { const look = structuredClone({ ...cur, raw: undefined }); ch.set(look, ch.check ? !on : true); readCaches(app.wb, look); looks.set(cur, look); place(); },
              run: () => { looks.delete(cur); void chooseRibbon(cur, e.id, ch, on); },
            });
          }
          items.push("-", { label: e.more, run: () => paneFor(cur, e.pane(cur)) });
          return items;
        },
      };
    });
  }

  async function chooseRibbon(m: ChartModel, id: string, ch: import("@raabe/chart-ui").ElementChoice, on: boolean): Promise<void> {
    if (id === "trendline" && m.series.filter((s) => !s.filtered).length > 1) {
      const i = await pickSeries(m);
      if (i === null) return;
      edit(m, (c) => ch.set({ ...c, series: [c.series[i]] } as ChartModel, true));
      return;
    }
    edit(m, (c) => ch.set(c, ch.check ? !on : true));
  }

  /** Format > Fülleffekt / Formkontur / Textfüllung on the selected element (the whole chart: the chart area). */
  function setElementColor(what: "fill" | "line" | "text", color: string | null): void {
    const m = selected;
    if (!m) return;
    const e = element;
    edit(m, (c) => {
      const f = (c.format ??= {});
      const target = ((): { fill?: unknown; line?: unknown; text?: { color?: string } } | null => {
        switch (e.kind) {
          case "series": return c.series[e.index] as never;
          case "chart": return (f.chart ??= {});
          case "plot": return (f.plot ??= {});
          case "title": return (f.title ??= {});
          case "legend": return (f.legend ??= {});
          case "valAxis": return (f.valAxis ??= {});
          case "catAxis": return (f.catAxis ??= {});
          case "majorGrid": return (f.majorGrid ??= {});
          case "labels": return (f.dataLabels ??= {});
          case "axisTitle": { const at = (f.axisTitles ??= {}); return (at[e.axis] ??= {}); }
          default: return null;
        }
      })();
      if (!target) return;
      if (what === "fill") target.fill = color === null ? { type: "none" } : { type: "solid", color };
      else if (what === "line") target.line = color === null ? { type: "none" } : { type: "solid", color };
      else target.text = { ...(target.text ?? {}), ...(color ? { color } : {}) };
    }, what === "fill" ? "Fülleffekt" : what === "line" ? "Formkontur" : "Textfüllung");
  }

  return {
    insert,
    refresh,
    /** A chart drawn for printing at a size in page px (print.ts, D-094). */
    printSvg: (m: ChartModel, w: number, hh: number): SVGSVGElement => {
      readCaches(app.wb, m);
      return renderChart(m, w, hh, { accents: accents(), format }) as SVGSVGElement;
    },
    /** The selected chart, if any (for the ribbon's contextual tabs). */
    selected: (): ChartModel | null => selected,
    /** A cell was selected: the chart is no longer (Excel). */
    deselect: (): void => { if (selected) select(null); },
    deleteSelected: (): void => { if (selected) remove(selected); },
    /** Diagrammentwurf > Diagrammelement hinzufügen: the menu of elements. */
    elementsMenu,
    /** The flyout of one of the three buttons beside the selected chart. */
    showButton: (which: "elements" | "styles" | "filter"): void => {
      if (!selected) return;
      const v = views.get(selected);
      if (v) buttons.show(which, selected, v.el);
    },
    /** Diagrammentwurf > Diagrammtyp ändern. */
    changeKind: (kind: ChartKind): void => { if (selected) edit(selected, (m) => { m.kind = kind; for (const x of m.series) { delete x.kind; delete x.secondary; } }); },
    /** Diagrammtyp ändern (dialog): any change of the selected chart as one undo step. */
    changeSelected: (fn: (m: ChartModel) => void): void => { if (selected) edit(selected, fn, "Diagrammtyp ändern"); },
    /** The series a chart of `kind` would get from the selection (the dialog's preview). */
    previewSeries: (kind: ChartKind): ChartModel["series"] => {
      const rect = sourceRect(), sheet = app.sheet;
      const series = rect ? seriesFromBlock({ sheet: sheet.name, ...rect, get: (r, c) => cellValue(sheet.value(r, c)) }, kind) : [];
      const m = newChart(kind, series, { col: 0, row: 0, x: 0, y: 0 }, { col: 1, row: 1, x: 0, y: 0 });
      readCaches(app.wb, m);
      return m.series;
    },
    switchRowCol,
    quickLayout: (n: number) => quickLayout(n),
    /** Schnelllayout's Livevorschau: layout n drawn on the chart (null ends it). */
    previewLayout: (n: number | null): void => {
      const m = selected;
      if (!m) return;
      if (n === null) looks.delete(m);
      else { const look = structuredClone({ ...m, raw: undefined }); quickLayout(n, look); readCaches(app.wb, look); looks.set(m, look); }
      place();
    },
    /** Diagrammentwurf > Farben ändern: a colour scheme (Excel's 10-26; series with their own colour keep it). */
    setColors: (k: number): void => { if (selected) edit(selected, (c) => { c.colors = k === 10 ? undefined : k; for (const s of c.series) delete s.color; }, "Farben ändern"); },
    previewColors: (k: number | null): void => {
      const m = selected;
      if (!m) return;
      if (k === null) looks.delete(m);
      else { const look = structuredClone({ ...m, raw: undefined }); look.colors = k === 10 ? undefined : k; for (const s of look.series) delete s.color; readCaches(app.wb, look); looks.set(m, look); }
      place();
    },
    colors: (): number => selected?.colors ?? 10,
    layoutCount: (): number => (selected ? layoutsFor(selected).filter(Boolean).length : 0),
    /** Diagrammformatvorlagen (ribbon gallery): the styles, the current one, a tile, applying, the preview. */
    styleCount: (): number => (selected ? chartStyles(selected).length : 0),
    styleCurrent: (): number | undefined => { if (!selected) return undefined; const i = styleIndex(selected); return i >= 0 ? i : undefined; },
    styleTile: (i: number, w: number, hh: number): SVGSVGElement | null => {
      if (!selected) return null;
      setStyleTheme(app.wb.theme);
      const look = structuredClone({ ...selected, raw: undefined });
      applyStyle(look, i);
      return tileOf(look, w, hh);
    },
    applyStyle: (i: number): void => { if (selected) { setStyleTheme(app.wb.theme); looks.delete(selected); edit(selected, (c) => applyStyle(c, i), "Diagrammformatvorlage"); } },
    previewStyle: (i: number | null): void => {
      const m = selected;
      if (!m) return;
      if (i === null) looks.delete(m);
      else { setStyleTheme(app.wb.theme); const look = structuredClone({ ...m, raw: undefined }); applyStyle(look, i); readCaches(app.wb, look); looks.set(m, look); }
      place();
    },
    styleKey: (): string => (selected ? `${selected.kind},${selected.series.length},${selected.colors ?? 10},${JSON.stringify(app.wb.theme)}` : ""),
    selectData: (): void => void selectData(),
    moveChart: (): void => void moveChart(),
    /** Format > Aktuelle Auswahl: the elements of the selected chart, the selected one, choosing one. */
    elementNames: (): string[] => (selected ? chartElements(selected).map((x) => elementName(selected!, x)) : []),
    elementName: (): string => (selected ? elementName(selected, element) : ""),
    chooseElement: (name: string): void => {
      if (!selected) return;
      const x = chartElements(selected).find((e) => elementName(selected!, e) === name);
      if (x) setElement(selected, x);
    },
    /** Format > Auswahl formatieren (Strg+1). */
    formatSelection: (): void => { if (selected) openPane(selected); },
    /** Format > Auf Formatvorlage zurücksetzen: the selected element's own formatting goes. */
    resetElement: (): void => { if (selected) resetElement(selected, element); },
    /** Format > Fülleffekt, Formkontur, Textfüllung (null: Keine Füllung / Kein Rahmen). */
    setElementColor,
    /** Format > Anordnen. */
    arrange: (how: "front" | "back" | "forward" | "backward"): void => { if (selected) arrange(selected, how); },
    /** Format > Größe: the chart's height and width in cm. */
    sizeCm: (): { w: number; h: number } | null => {
      if (!selected) return null;
      const b = sheetBox(selected);
      return { w: Math.round((b.w / 96) * 2.54 * 100) / 100, h: Math.round((b.h / 96) * 2.54 * 100) / 100 };
    },
    setSizeCm: (w: number | null, hh: number | null): void => {
      const m = selected;
      if (!m) return;
      const b = sheetBox(m);
      const nw = w !== null ? (w / 2.54) * 96 : b.w, nh = hh !== null ? (hh / 2.54) * 96 : b.h;
      if (!(nw > 0 && nh > 0)) return;
      edit(m, (c) => { c.to = anchorAt(b.x + nw, b.y + nh); }, "Größe ändern");
    },
    /** For tests and the harness: the source frames on the sheet. */
    rangeFrames: (): { kind: string; rect: Rect }[] => (selected ? sourceRanges(selected).map((b) => ({ kind: b.kind, rect: b.rect })) : []),
  };
}
