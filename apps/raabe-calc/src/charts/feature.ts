// Charts in RaabeCalc (Step 7 E, D-061). A chart is made from the selection (or the current
// region around the active cell) as Excel splits it into series, placed in the middle of the
// visible cells at Excel's default size (480 x 288 px), and kept in `sheet.charts` with a
// two-cell anchor as the file has it. Charts float over the grid as SVG in a layer of their own,
// follow scrolling and zoom, and re-read their cells after every change. A chart is selected
// with a click (frame with 8 handles as in Excel), moved and sized with the mouse, deleted with
// Entf; the "+" button and a double-click open Excel's Diagrammelemente list (Achsentitel,
// Diagrammtitel, Datenbeschriftungen, Gitternetzlinien, Legende, Trendlinie with Formel and
// Bestimmtheitsmaß). Every change is one undo step (History.charts).

import {
  QUICK_LAYOUTS, isBar, isPie, newChart, parseRef, refText, refreshCaches, renderChart, seriesFromBlock, seriesName, sourceOf, titleText,
  type CellAnchor, type CellValue, type ChartKind as Kind, type ChartModel, type LegendPos,
} from "@raabe/charts";
import { CellError, formatValue, type Rect, type Sheet, type Workbook } from "@raabe/sheet-engine";
import { h, messageBox, showDialog } from "@raabe/shell";
import type { CalcApp } from "../app";
import { chartElements, elementName, formatPane, type ChartElement } from "./formatpane";
import { sheetAxes, type Axis } from "../grid/metrics";

export type ChartKind = Kind;

/** Excel's size of a new chart: 360 x 216 pt (12,7 x 7,62 cm). */
const NEW_W = 480, NEW_H = 288;
/** Smallest size while sizing with the mouse (px at 100 %). */
const MIN_SIZE = 20;
const SCROLLBAR = 17;

interface View { el: HTMLDivElement; key: string; zoomKey: string }
interface Box { x: number; y: number; w: number; h: number }
type Handle = "nw" | "n" | "ne" | "e" | "se" | "s" | "sw" | "w";
const HANDLES: Handle[] = ["nw", "n", "ne", "e", "se", "s", "sw", "w"];

const CSS = `
.chart-layer { position: absolute; overflow: hidden; pointer-events: none; z-index: 4; }
.chart-obj { position: absolute; pointer-events: auto; outline: none; cursor: default; touch-action: none; }
.chart-obj > svg { display: block; position: absolute; left: 0; top: 0; }
.chart-obj.sel { box-shadow: 0 0 0 1px #a6a6a6; }
.chart-obj.dragging { cursor: move; }
.chart-handle { position: absolute; width: 7px; height: 7px; margin: -5px 0 0 -5px; border: 1px solid #5f5f5f; border-radius: 50%; background: #fff; display: none; z-index: 1; }
.chart-obj.sel .chart-handle { display: block; }
.chart-handle[data-h="nw"], .chart-handle[data-h="se"] { cursor: nwse-resize; }
.chart-handle[data-h="ne"], .chart-handle[data-h="sw"] { cursor: nesw-resize; }
.chart-handle[data-h="n"], .chart-handle[data-h="s"] { cursor: ns-resize; }
.chart-handle[data-h="e"], .chart-handle[data-h="w"] { cursor: ew-resize; }
.chart-plus { position: absolute; top: 0; right: -30px; width: 24px; height: 24px; display: none; align-items: center; justify-content: center; padding: 0;
  background: #fff; border: 1px solid #c6c6c6; color: #444; font: 18px/1 var(--font, "Segoe UI", sans-serif); cursor: default; }
.chart-plus:hover { background: #f3f3f3; }
.chart-obj.sel .chart-plus { display: flex; }
.chart-flyout { position: fixed; z-index: 1000; min-width: 210px; padding: 6px 0; background: #fff; border: 1px solid #c6c6c6; box-shadow: 0 2px 8px rgba(0,0,0,.2); font: 12px var(--font, "Segoe UI", sans-serif); color: #262626; }
.chart-flyout-head { padding: 2px 12px 6px; font-weight: 600; }
.chart-flyout label { display: flex; align-items: center; gap: 8px; padding: 3px 12px; }
.chart-flyout label:hover { background: #f0f0f0; }
.chart-flyout label.sub { padding-left: 34px; }
.chart-flyout input[type=text] { flex: 1; min-width: 0; font: inherit; padding: 1px 4px; border: 1px solid #c6c6c6; }
.chart-flyout select { margin-left: auto; font: inherit; }
.chart-elsel { position: absolute; pointer-events: none; border: 1px dashed #5f5f5f; display: none; }
.chart-obj.sel .chart-elsel.on { display: block; }
`;

export function chartsFeature(app: CalcApp) {
  let layer: HTMLDivElement | null = null;
  let views = new Map<ChartModel, View>();
  let selected: ChartModel | null = null;
  let flyout: HTMLElement | null = null;
  /** Closes the flyout on a click elsewhere. */
  let away: ((e: PointerEvent) => void) | null = null;
  let axes: { rows: Axis; cols: Axis } | null = null;
  let viewSig = "";
  let shownSheet: Sheet | null = null;
  let shownList: ChartModel[] | undefined;
  let drag: { m: ChartModel; handle: Handle | null; x: number; y: number; box: Box; draft: Box; moved: boolean } | null = null;
  /** The selected element of the selected chart (Excel: Diagrammelemente on the Format tab). */
  let element: ChartElement = { kind: "chart" };

  const zoom = () => app.sheet.view.zoom / 100;

  // ----- set-up --------------------------------------------------------------------------------

  function ensureLayer(): HTMLDivElement {
    if (layer) return layer;
    if (!document.getElementById("raabe-charts-css")) document.head.append(h("style", { id: "raabe-charts-css" }, CSS));
    layer = h("div", { class: "chart-layer" });
    const canvas = app.grid.el.querySelector(".grid-canvas");
    app.grid.el.insertBefore(layer, canvas ? canvas.nextSibling : app.grid.el.firstChild);
    // A click anywhere else in the grid deselects the chart (capture: before the grid sees it).
    app.grid.el.addEventListener("pointerdown", (e) => {
      if (!(e.target as HTMLElement).closest?.(".chart-obj")) select(null);
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
    if (!views.size) return;
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
        // Data in hidden rows and columns (also those a filter hides) is left out, as Excel does by default.
        if (sheet.rows.get(row)?.hidden && r.r1 > r.r0) continue;
        cols = 0;
        for (let col = r.c0; col <= r.c1; col++) {
          if (sheet.cols.get(col)?.hidden && r.c1 > r.c0) continue;
          const cell = sheet.get(row, col);
          if (format === undefined && cell) format = wb.styles.get(cell.s).numFmt;
          vals.push(cellValue(cell?.v ?? null));
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
    el.addEventListener("dblclick", (e) => {
      e.preventDefault();
      e.stopPropagation();
      if ((e.target as HTMLElement).closest?.(".chart-plus")) return;
      select(m);
      // Excel: a double-click on an element opens its Format pane.
      setElement(m, elementAt(e.target as Element) ?? element);
      openPane(m);
    });
    el.addEventListener("contextmenu", (e) => { e.preventDefault(); e.stopPropagation(); select(m); openFlyout(m); });
    el.addEventListener("keydown", (e) => onKey(e, m));
    el.addEventListener("wheel", () => closeFlyout(), { passive: true });
    v = { el, key: "", zoomKey: "" };
    views.set(m, v);
    return v;
  }

  /** Draws a chart at a size (px at 100 %) unless it is already drawn so. */
  function draw(m: ChartModel, v: View, box: Box): void {
    const z = zoom();
    const { from: _f, to: _t, raw: _r, ...content } = m;
    void _f; void _t; void _r;
    const key = JSON.stringify([content, Math.round(box.w), Math.round(box.h), app.wb.theme]);
    if (key !== v.key) {
      const svg = renderChart(m, box.w, box.h, { accents: accents(), format });
      svg.setAttribute("aria-label", titleText(m));
      v.el.querySelector("svg")?.remove();
      v.el.prepend(svg);
      if (!v.el.querySelector(".chart-handle")) {
        for (const hd of HANDLES) v.el.append(h("div", { class: "chart-handle", "data-h": hd }));
        const plus = h("button", { class: "chart-plus", "aria-label": "Diagrammelemente", "data-tip": "Diagrammelemente", tabindex: "-1" }, "+");
        plus.addEventListener("pointerdown", (e) => { e.stopPropagation(); e.preventDefault(); });
        plus.addEventListener("click", (e) => { e.stopPropagation(); flyout ? closeFlyout() : openFlyout(m); });
        v.el.append(plus);
      }
      v.key = key;
      v.zoomKey = "";
    }
    const zk = `${z},${box.w},${box.h}`;
    if (zk !== v.zoomKey) {
      const svg = v.el.querySelector("svg")!;
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
    }
  }

  /** Rebuilds the charts of the active sheet (after changes, sheet or document switches). */
  function refresh(): void {
    if (!app.grid) return;
    const l = ensureLayer();
    const sheet = app.sheet;
    const list = sheet.charts ?? [];
    axes = null;
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
      // Undo/redo replaces the models: keep the selection on the same chart by position.
      selected = null;
      closeFlyout();
      if (paneChart) { paneChart = null; app.taskPane.close(); }
    }
    for (const [m, v] of views) v.el.classList.toggle("sel", m === selected);
    // Drawing order: later charts on top.
    list.forEach((m, i) => { views.get(m)!.el.style.zIndex = String(i + 1); });
    place();
  }

  // ----- selection, moving and sizing ------------------------------------------------------------------

  function select(m: ChartModel | null): void {
    queueMicrotask(() => app.chartSelectionChanged());
    if (m === selected) return;
    selected = m;
    element = { kind: "chart" };
    if (!m && app.taskPane.isOpen && paneChart) { paneChart = null; app.taskPane.close(); }
    for (const [k, v] of views) v.el.classList.toggle("sel", k === m);
    if (!m) closeFlyout();
    else views.get(m)?.el.focus({ preventScroll: true });
  }

  function onDown(e: PointerEvent, m: ChartModel): void {
    if (e.button !== 0) return;
    e.stopPropagation();
    e.preventDefault();
    if (app.grid.isEditing) app.grid.commitEdit("none");
    closeFlyout();
    select(m);
    // A click on an element selects it; a click on the chart's handles keeps the chart area.
    if (!(e.target as HTMLElement).closest?.(".chart-handle")) setElement(m, elementAt(e.target as Element) ?? { kind: "chart" });
    const view = views.get(m)!;
    view.el.focus({ preventScroll: true });
    const handle = ((e.target as HTMLElement).closest?.(".chart-handle") as HTMLElement | null)?.dataset.h as Handle | undefined;
    const box = sheetBox(m);
    drag = { m, handle: handle ?? null, x: e.clientX, y: e.clientY, box, draft: { ...box }, moved: false };
    try { view.el.setPointerCapture(e.pointerId); } catch { /* synthetic events */ }
  }

  function onMove(e: PointerEvent): void {
    if (!drag) return;
    const z = zoom();
    const dx = (e.clientX - drag.x) / z, dy = (e.clientY - drag.y) / z;
    if (!drag.moved && Math.abs(e.clientX - drag.x) + Math.abs(e.clientY - drag.y) < 3) return;
    drag.moved = true;
    const b = drag.box;
    let { x, y, w, h: hh } = b;
    const k = drag.handle;
    if (!k) {
      x = Math.max(0, b.x + dx);
      y = Math.max(0, b.y + dy);
      views.get(drag.m)?.el.classList.add("dragging");
    } else {
      if (k.includes("e")) w = b.w + dx;
      if (k.includes("s")) hh = b.h + dy;
      if (k.includes("w")) { w = b.w - dx; x = b.x + dx; }
      if (k.includes("n")) { hh = b.h - dy; y = b.y + dy; }
      // Umschalt on a corner keeps the proportions (Excel).
      if (e.shiftKey && k.length === 2) {
        const s = Math.max(w / b.w, hh / b.h);
        w = b.w * s;
        hh = b.h * s;
        if (k.includes("w")) x = b.x + b.w - w;
        if (k.includes("n")) y = b.y + b.h - hh;
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
      remove(m);
    } else if (e.key === "Escape") {
      e.preventDefault();
      e.stopPropagation();
      if (flyout) closeFlyout();
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

  // ----- elements and the Format pane --------------------------------------------------------------

  /** The element under the pointer, from the drawing's data-part marks. */
  function elementAt(target: Element): ChartElement | null {
    const part = (target.closest?.("[data-part]") as SVGElement | null)?.dataset.part;
    if (!part) return null;
    const [kind, a] = part.split(":");
    switch (kind) {
      case "chart": return { kind: "chart" };
      case "plot": return { kind: "plot" };
      case "title": return { kind: "title" };
      case "legend": return { kind: "legend" };
      case "axis": return a === "val" ? { kind: "valAxis" } : { kind: "catAxis" };
      case "axes": return { kind: "catAxis" };
      case "grid": return { kind: "majorGrid" };
      case "minorGrid": return { kind: "minorGrid" };
      case "catGrid": return { kind: "catGrid" };
      case "axisTitle": return { kind: "axisTitle", axis: a === "val" ? "val" : "cat" };
      case "series": return { kind: "series", index: Number(a) };
      case "labels": return { kind: "labels" };
      case "trend": case "trendLabel": return { kind: "trend", index: Number(a) };
      case "errbars": return { kind: "errorBars", index: Number(a) };
      default: return null;
    }
  }

  /** The SVG parts that make up an element (for its selection frame). */
  function partsOf(e: ChartElement): string {
    switch (e.kind) {
      case "chart": return "";
      case "valAxis": return '[data-part="axis:val"]';
      case "catAxis": return '[data-part="axis:cat"],[data-part="axes"]';
      case "majorGrid": return '[data-part="grid"]';
      case "series": return `[data-part^="series:${e.index}"]`;
      case "labels": return '[data-part^="labels:"]';
      case "trend": return `[data-part="trend:${e.index}"]`;
      case "errorBars": return `[data-part="errbars:${e.index}"]`;
      case "axisTitle": return `[data-part="axisTitle:${e.axis}"]`;
      default: return `[data-part="${e.kind}"]`;
    }
  }

  function setElement(m: ChartModel, e: ChartElement): void {
    element = e;
    showElement(m);
    if (paneChart === m && app.taskPane.isOpen) openPane(m);
    queueMicrotask(() => app.chartSelectionChanged());
  }

  /** The dashed frame around the selected element (the chart area has the chart's own frame). */
  function showElement(m: ChartModel): void {
    const v = views.get(m);
    if (!v) return;
    let box = v.el.querySelector<HTMLElement>(".chart-elsel");
    if (!box) { box = h("div", { class: "chart-elsel" }); v.el.append(box); }
    const sel = partsOf(element);
    const els = sel ? [...v.el.querySelectorAll<SVGGraphicsElement>(sel)] : [];
    if (!els.length) { box.classList.remove("on"); return; }
    const o = v.el.getBoundingClientRect();
    let l = Infinity, t = Infinity, r = -Infinity, b = -Infinity;
    for (const el of els) {
      const rr = el.getBoundingClientRect();
      l = Math.min(l, rr.left); t = Math.min(t, rr.top); r = Math.max(r, rr.right); b = Math.max(b, rr.bottom);
    }
    Object.assign(box.style, { left: `${l - o.left - 2}px`, top: `${t - o.top - 2}px`, width: `${r - l + 4}px`, height: `${b - t + 4}px` });
    box.classList.add("on");
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
    }));
  }

  // ----- Diagrammelemente -----------------------------------------------------------------------------

  function closeFlyout(): void {
    flyout?.remove();
    flyout = null;
    if (away) removeEventListener("pointerdown", away, true);
    away = null;
  }

  /** Changes the selected chart as one undo step and redraws the list. */
  function edit(m: ChartModel, fn: (m: ChartModel) => void): void {
    const sheet = app.sheet;
    app.change("Diagrammelemente", () => {
      app.history.charts(sheet);
      fn(m);
    });
    if (flyout) openFlyout(m);
    if (paneChart === m && app.taskPane.isOpen) app.taskPane.refresh();
  }

  function openFlyout(m: ChartModel): void {
    closeFlyout();
    const view = views.get(m);
    if (!view) return;
    const pie = isPie(m.kind), hist = m.kind === "histogram";
    const check = (label: string, on: boolean, change: (on: boolean) => void, sub = false) => {
      const box = h("input", { type: "checkbox" }) as HTMLInputElement;
      box.checked = on;
      box.addEventListener("change", () => change(box.checked));
      return h("label", { class: sub ? "sub" : undefined }, box, label);
    };
    const text = (label: string, value: string, change: (v: string) => void) => {
      const input = h("input", { type: "text", "aria-label": label }) as HTMLInputElement;
      input.value = value;
      const commit = () => { if (input.value !== value) change(input.value); };
      input.addEventListener("keydown", (e) => {
        e.stopPropagation();
        if (e.key === "Enter") { e.preventDefault(); commit(); }
        if (e.key === "Escape") { e.preventDefault(); closeFlyout(); view.el.focus(); }
      });
      input.addEventListener("change", commit);
      return h("label", { class: "sub" }, input);
    };
    const rows: HTMLElement[] = [h("div", { class: "chart-flyout-head" }, "Diagrammelemente")];
    if (!pie) {
      const on = m.axisTitles.cat !== undefined || m.axisTitles.val !== undefined;
      rows.push(check("Achsentitel", on, (v) => edit(m, (c) => { c.axisTitles = v ? { cat: "", val: "" } : {}; })));
      if (on) {
        if (m.axisTitles.cat !== undefined) rows.push(text("Primär horizontal", m.axisTitles.cat || "Achsentitel", (t) => edit(m, (c) => { c.axisTitles.cat = t; })));
        if (m.axisTitles.val !== undefined) rows.push(text("Primär vertikal", m.axisTitles.val || "Achsentitel", (t) => edit(m, (c) => { c.axisTitles.val = t; })));
      }
    }
    rows.push(check("Diagrammtitel", m.title.show, (v) => edit(m, (c) => { c.title.show = v; })));
    if (m.title.show) rows.push(text("Diagrammtitel", titleText(m), (t) => edit(m, (c) => { c.title.text = t; })));
    rows.push(check("Datenbeschriftungen", m.dataLabels, (v) => edit(m, (c) => { c.dataLabels = v; })));
    if (!pie) rows.push(check("Gitternetzlinien", m.gridlines, (v) => edit(m, (c) => { c.gridlines = v; })));
    if (!hist) {
      const pos = h("select", { "aria-label": "Legende" }) as HTMLSelectElement;
      for (const [val, label] of [["r", "Rechts"], ["t", "Oben"], ["l", "Links"], ["b", "Unten"]] as [LegendPos, string][]) {
        pos.append(h("option", { value: val }, label));
      }
      pos.value = m.legend === "none" ? "r" : m.legend;
      pos.disabled = m.legend === "none";
      pos.addEventListener("change", () => edit(m, (c) => { c.legend = pos.value as LegendPos; }));
      const legend = check("Legende", m.legend !== "none", (v) => edit(m, (c) => { c.legend = v ? "r" : "none"; }));
      legend.append(pos);
      rows.push(legend);
    }
    if (!pie && !hist && !isBar(m.kind) && m.series.length) {
      const tl = m.series[0].trendline;
      rows.push(check("Trendlinie", !!tl, (v) => edit(m, (c) => {
        for (const s of c.series) s.trendline = v ? { type: "linear", showEq: false, showR2: false } : undefined;
      })));
      if (tl) {
        rows.push(check("Formel im Diagramm anzeigen", tl.showEq, (v) => edit(m, (c) => { for (const s of c.series) if (s.trendline) s.trendline.showEq = v; }), true));
        rows.push(check("Bestimmtheitsmaß im Diagramm anzeigen", tl.showR2, (v) => edit(m, (c) => { for (const s of c.series) if (s.trendline) s.trendline.showR2 = v; }), true));
      }
    }
    flyout = h("div", { class: "chart-flyout", role: "dialog", "aria-label": "Diagrammelemente" }, ...rows);
    flyout.addEventListener("pointerdown", (e) => e.stopPropagation());
    flyout.addEventListener("keydown", (e) => {
      if (e.key === "Escape") { e.preventDefault(); e.stopPropagation(); closeFlyout(); view.el.focus(); }
    });
    document.body.append(flyout);
    const r = view.el.getBoundingClientRect();
    const fw = flyout.offsetWidth, fh = flyout.offsetHeight;
    let left = r.right + 36;
    if (left + fw > innerWidth - 4) left = Math.max(4, r.left - fw - 8);
    flyout.style.left = `${left}px`;
    flyout.style.top = `${Math.max(4, Math.min(r.top, innerHeight - fh - 4))}px`;
    away = (e: PointerEvent) => {
      if (flyout?.contains(e.target as Node) || (e.target as HTMLElement).closest?.(".chart-plus")) return;
      closeFlyout();
    };
    addEventListener("pointerdown", away, true);
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
  function quickLayout(n: number): void {
    const m = selected;
    const l = m ? layoutsFor(m)[n - 1] : null;
    if (!m || !l) return;
    edit(m, (c) => {
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
      const defGap = c.kind.endsWith("Stacked") ? 150 : c.kind === "bar" ? 182 : 219;
      if (l.gap !== undefined && (c.kind === "column" || c.kind === "bar" || c.kind.endsWith("Stacked"))) c.gapWidth = l.gap === defGap ? undefined : l.gap;
      if (!Object.keys(f).length) delete c.format;
    });
  }

  // ----- Daten: Zeile/Spalte wechseln, Daten auswählen -------------------------------------------

  /** The series of a range, as Excel makes them (`orientation` forced by Zeile/Spalte wechseln). */
  function seriesOf(m: ChartModel, ref: { sheet: string; r0: number; c0: number; r1: number; c1: number }, orientation?: "rows" | "cols") {
    const sheet = app.wb.sheets.find((s) => s.name === ref.sheet) ?? app.sheet;
    return seriesFromBlock({ ...ref, sheet: sheet.name, get: (r, c) => cellValue(sheet.value(r, c)) }, m.kind, orientation);
  }

  /** Diagrammentwurf > Zeile/Spalte wechseln: the series run the other way through the same range. */
  function switchRowCol(): void {
    const m = selected;
    if (!m) return;
    const src = sourceOf(m.series);
    if (!src) return;
    edit(m, (c) => {
      c.series = seriesOf(c, src.ref, src.orientation === "cols" ? "rows" : "cols");
      delete c.raw;
      readCaches(app.wb, c);
    });
  }

  /** Diagrammentwurf > Daten auswählen: Excel's "Datenquelle auswählen". */
  async function selectData(): Promise<void> {
    const m = selected;
    if (!m) return;
    const src = sourceOf(m.series);
    const range = h("input", { type: "text", style: "width:260px", value: src ? refText(src.ref.sheet, src.ref.r0, src.ref.c0, src.ref.r1, src.ref.c1) : "" }) as HTMLInputElement;
    const list = h("select", { size: 6, style: "width:220px" }) as HTMLSelectElement;
    const cats = h("div", { style: "width:220px;height:96px;overflow:auto;border:1px solid #c6c6c6;padding:2px 4px" });
    let draft = m.series.map((s) => structuredClone(s));
    let orientation = src?.orientation;
    const fill = () => {
      const keep = list.selectedIndex;
      list.replaceChildren(...draft.map((s, i) => h("option", { value: String(i) }, seriesName(s, i))));
      list.selectedIndex = Math.min(Math.max(0, keep), draft.length - 1);
      cats.replaceChildren(...(draft[0]?.cat?.cache ?? []).map((v) => h("div", null, v === null ? "" : String(v))));
    };
    const fromRange = (o?: "rows" | "cols") => {
      const r = parseRef(range.value.replace(/^=/, ""));
      if (!r) return;
      draft = seriesOf(m, r, o);
      fill();
    };
    range.addEventListener("change", () => fromRange());
    const btn = (label: string, run: () => void) => h("button", { type: "button", class: "dlg-btn", onclick: run }, label);
    const editSeries = async (i: number | null) => {
      const s = i === null ? { val: { cache: [] as (number | null)[] } } : draft[i];
      const name = h("input", { type: "text", style: "width:240px", value: s.name?.ref ? "=" + s.name.ref : s.name?.text ?? "" }) as HTMLInputElement;
      const vals = h("input", { type: "text", style: "width:240px", value: s.val.ref ? "=" + s.val.ref : "={1}" }) as HTMLInputElement;
      const ok = await showDialog({ title: "Datenreihe bearbeiten", body: h("div", null, h("label", { class: "field" }, "Reihenname:", name), h("label", { class: "field" }, "Reihenwerte:", vals)),
        buttons: [{ label: "OK", value: true, primary: true }, { label: "Abbrechen", value: false }], cancel: false });
      if (!ok) return;
      const n = name.value.startsWith("=") ? { ref: name.value.slice(1) } : { text: name.value };
      const next = { ...s, name: n, val: { ...s.val, ref: vals.value.replace(/^=/, "") } };
      if (i === null) draft.push(next); else draft[i] = next;
      fill();
    };
    fill();
    const body = h("div", { style: "display:flex;flex-direction:column;gap:10px" },
      h("label", { class: "field" }, "Diagrammdatenbereich:", range),
      h("div", null, btn("Zeile/Spalte wechseln", () => { orientation = orientation === "cols" ? "rows" : "cols"; fromRange(orientation); })),
      h("div", { style: "display:flex;gap:16px" },
        h("div", null, h("div", null, "Legendeneinträge (Reihen)"),
          h("div", { style: "display:flex;gap:4px;margin:4px 0" },
            btn("Hinzufügen", () => void editSeries(null)), btn("Bearbeiten", () => list.selectedIndex >= 0 && void editSeries(list.selectedIndex)),
            btn("Entfernen", () => { if (list.selectedIndex >= 0) { draft.splice(list.selectedIndex, 1); fill(); } }),
            btn("\u25b2", () => { const i = list.selectedIndex; if (i > 0) { [draft[i - 1], draft[i]] = [draft[i], draft[i - 1]]; fill(); list.selectedIndex = i - 1; } }),
            btn("\u25bc", () => { const i = list.selectedIndex; if (i >= 0 && i < draft.length - 1) { [draft[i + 1], draft[i]] = [draft[i], draft[i + 1]]; fill(); list.selectedIndex = i + 1; } })),
          list),
        h("div", null, h("div", null, "Horizontale Achsenbeschriftungen (Rubrik)"), h("div", { style: "height:26px" }), cats)));
    const ok = await showDialog({ title: "Datenquelle auswählen", body, buttons: [{ label: "OK", value: true, primary: true }, { label: "Abbrechen", value: false }], cancel: false });
    if (!ok) return;
    if (!draft.length) { await messageBox("RaabeCalc", "Ein Diagramm braucht mindestens eine Datenreihe."); return; }
    edit(m, (c) => { c.series = draft; delete c.raw; readCaches(app.wb, c); });
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
    /** Diagrammentwurf > Diagrammelement hinzufügen: the elements list of the selected chart. */
    elements: (): void => { if (selected) openFlyout(selected); },
    /** Diagrammentwurf > Diagrammtyp ändern. */
    changeKind: (kind: ChartKind): void => { if (selected) edit(selected, (m) => { m.kind = kind; for (const x of m.series) { delete x.kind; delete x.secondary; } }); },
    /** Diagrammtyp ändern (dialog): any change of the selected chart as one undo step. */
    changeSelected: (fn: (m: ChartModel) => void): void => { if (selected) edit(selected, fn); },
    /** The series a chart of `kind` would get from the selection (the dialog's preview). */
    previewSeries: (kind: ChartKind): ChartModel["series"] => {
      const rect = sourceRect(), sheet = app.sheet;
      const series = rect ? seriesFromBlock({ sheet: sheet.name, ...rect, get: (r, c) => cellValue(sheet.value(r, c)) }, kind) : [];
      const m = newChart(kind, series, { col: 0, row: 0, x: 0, y: 0 }, { col: 1, row: 1, x: 0, y: 0 });
      readCaches(app.wb, m);
      return m.series;
    },
    switchRowCol,
    quickLayout,
    /** Diagrammentwurf > Farben ändern: a colour scheme (Excel's 10-26; series with their own colour keep it). */
    setColors: (k: number): void => { if (selected) edit(selected, (c) => { c.colors = k === 10 ? undefined : k; for (const s of c.series) delete s.color; }); },
    layoutCount: (): number => (selected ? layoutsFor(selected).filter(Boolean).length : 0),
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
    resetElement: (): void => {
      if (!selected) return;
      const e = element;
      edit(selected, (m) => {
        const f = m.format;
        const k = e.kind;
        if (k === "series") { const s = m.series[e.index]; delete s.fill; delete s.line; delete s.marker; return; }
        if (!f) return;
        const key = ({ chart: "chart", plot: "plot", title: "title", legend: "legend", valAxis: "valAxis", catAxis: "catAxis", majorGrid: "majorGrid", minorGrid: "minorGrid", catGrid: "catGrid", labels: "dataLabels" } as const)[k as "chart"];
        if (key) delete (f as Record<string, unknown>)[key];
      });
    },
  };
}
