// Datenschnitte (slicers) in RaabeCalc (D-092): Tabellenentwurf > Datenschnitt einfügen, Einfügen > Datenschnitt,
// PivotTable-Analyse > Datenschnitt einfügen; the slicer on the sheet with one button per item (a click
// selects one item, Strg+click or Mehrfachauswahl several, Filter löschen all), items with and without data
// drawn in Excel's default slicer style (and the 13 others) as recorded through COM (probe_excel_slicers.py,
// slicer-styles.gen.ts); filtering a table hides its rows and keeps the chosen values in the table's AutoFilter,
// filtering PivotTables hides their items. The contextual tab Datenschnitt (labels and keytips recorded,
// probe_excel_slicer_ui.py): Datenschnittbeschriftung, Datenschnitteinstellungen, Berichtsverbindungen,
// Datenschnitt-Formatvorlagen, Anordnen, Schaltflächen (Spalten, Höhe, Breite), Größe. Geometry measured on
// Excel's rendering at 96 dpi: header 21 pt, margins 6.75 pt, buttons 0.7 cm with 2 pt between rows and 2.25 pt
// between columns. Each change is one undo step.

import { h, messageBox, showDialog, showMenu, type Control, type MenuEntry, type Tab } from "@raabe/shell";
import {
  BLANK_ITEM, SLICER_ROW_HEIGHT, nextCacheName, nextSlicerName, passesTableFilters, pivotHiddenAfter, pivotSlicerRows, slicerItems, tableData,
  type PivotTable, type Sheet, type SheetTable, type Slicer, type SlicerCache, type SlicerItem,
} from "@raabe/sheet-engine";
import type { CalcApp } from "./app";
import { shown } from "./engine";
import { SLICER_STYLES, type SlicerElement } from "./slicer-styles.gen";
import { OFFICE_2013, tinted } from "./tables";
import type { TsColor } from "./table-styles.gen";

const PX_PER_PT = 96 / 72;
const EMU_PER_PX = 9525;
const PX_PER_CM = 96 / 2.54;
/** Excel's measures (pt): header height, margin left/right, gap after the header, gaps between buttons. */
const HEAD = 21, MARGIN = 6.75, AFTER_HEAD = 3, GAP_ROW = 2, GAP_COL = 2.25;
/** A new slicer: 144 x 206.25 pt (recorded default). */
const NEW_W = 144 * PX_PER_PT, NEW_H = 206.25 * PX_PER_PT;

const CSS = `
.slicer-layer { position: absolute; right: 17px; bottom: 0; overflow: hidden; pointer-events: none; z-index: 4; }
.slicer-layer .slicer { position: absolute; pointer-events: auto; box-sizing: border-box; background: #fff; transform-origin: 0 0; outline: none; user-select: none; }
.slicer-layer .slicer.on { box-shadow: 0 0 0 1px #7a7a7a; }
.slicer-layer .sl-head { position: absolute; left: 0; right: 0; top: 0; box-sizing: border-box; display: flex; align-items: center; padding-left: 9px; cursor: move; }
.slicer-layer .sl-cap { flex: 1; overflow: hidden; white-space: nowrap; text-overflow: ellipsis; }
.slicer-layer .sl-ic:last-child { margin-right: 7px; }
.slicer-layer .sl-ic { width: 20px; height: 20px; margin-right: 13px; display: flex; align-items: center; justify-content: center; border-radius: 2px; cursor: default; }
.slicer-layer .sl-ic:hover { background: #e6e6e6; }
.slicer-layer .sl-ic.pressed { background: #cce4f7; }
.slicer-layer .sl-body { position: absolute; left: 0; right: 0; bottom: 0; overflow-x: hidden; overflow-y: auto; }
.slicer-layer .sl-btn { position: absolute; box-sizing: border-box; border: 1px solid; border-radius: 3px; display: flex; align-items: center; padding-left: 8px; overflow: hidden; white-space: nowrap; cursor: default; }
.slicer-layer .sl-h { position: absolute; width: 8px; height: 8px; margin: -5px 0 0 -5px; border: 1px solid #5f5f5f; border-radius: 50%; background: #fff; display: none; z-index: 2; }
.slicer-layer .slicer.on .sl-h { display: block; }
.slicer-layer .sl-h.nw, .slicer-layer .sl-h.se { cursor: nwse-resize; }
.slicer-layer .sl-h.ne, .slicer-layer .sl-h.sw { cursor: nesw-resize; }
.slicer-layer .sl-h.n, .slicer-layer .sl-h.s { cursor: ns-resize; }
.slicer-layer .sl-h.e, .slicer-layer .sl-h.w { cursor: ew-resize; }
.slicer-gallery { display: grid; grid-template-columns: repeat(6, 64px); gap: 6px; padding: 8px; }
.slicer-gallery .sg-head { grid-column: 1 / -1; font-weight: 600; font-size: 12px; margin-top: 4px; }
.slicer-gallery .sg-tile { width: 60px; height: 60px; border: 1px solid #ddd; padding: 3px; box-sizing: border-box; cursor: pointer; }
.slicer-gallery .sg-tile:hover, .slicer-gallery .sg-tile.on { border-color: #6b9bd1; background: #e5f0fb; }
.slicer-dlg label { display: flex; align-items: center; gap: 6px; }
.slicer-dlg fieldset { border: 1px solid #d0d0d0; padding: 6px 10px; margin: 0; display: flex; flex-direction: column; gap: 4px; }
`;

type Handle = "nw" | "n" | "ne" | "e" | "se" | "s" | "sw" | "w";
const HANDLES: Handle[] = ["nw", "n", "ne", "e", "se", "s", "sw", "w"];
const POS: Record<Handle, [number, number]> = { nw: [0, 0], n: [0.5, 0], ne: [1, 0], e: [1, 0.5], se: [1, 1], s: [0.5, 1], sw: [0, 1], w: [0, 0.5] };
type Box = { x: number; y: number; w: number; h: number };
type Run = (id: string, fn: () => void | Promise<void>, enabled?: () => boolean, checked?: () => boolean) => void;

/** The commands of this feature (registered by `implement`). */
export const SLICER_COMMANDS = [
  { id: "insert.slicer", label: "Datenschnitt", icon: "filter" },
  { id: "table.slicer", label: "Datenschnitt einfügen", icon: "filter" },
  { id: "pivot.slicer", label: "Datenschnitt einfügen", icon: "filter" },
  { id: "slicer.caption", label: "Datenschnittbeschriftung" },
  { id: "slicer.settings", label: "Datenschnitteinstellungen...", icon: "settings" },
  { id: "slicer.connections", label: "Berichtsverbindungen...", icon: "table" },
  { id: "slicer.styles", label: "Schnellformatvorlagen", icon: "styles" },
  { id: "slicer.forward", label: "Ebene nach vorne", icon: "bring-forward" },
  { id: "slicer.toFront", label: "In den Vordergrund", icon: "bring-forward" },
  { id: "slicer.backward", label: "Ebene nach hinten", icon: "send-backward" },
  { id: "slicer.toBack", label: "In den Hintergrund", icon: "send-backward" },
  { id: "slicer.selectionPane", label: "Auswahlbereich...", icon: "select-objects" },
  { id: "slicer.align", label: "Ausrichten", icon: "align-left" },
  { id: "slicer.group", label: "Gruppieren", icon: "group" },
  { id: "slicer.rotate", label: "Drehen", icon: "rotate" },
  { id: "slicer.columns", label: "Spalten" },
  { id: "slicer.buttonHeight", label: "Höhe" },
  { id: "slicer.buttonWidth", label: "Breite" },
  { id: "slicer.height", label: "Formenhöhe" },
  { id: "slicer.width", label: "Formenbreite" },
  { id: "slicer.sizeDialog", label: "Größe und Eigenschaften..." },
  { id: "slicer.clear", label: "Filter löschen", icon: "filter-clear", shortcut: "Alt+C" },
  { id: "slicer.multi", label: "Mehrfachauswahl", icon: "select-objects", shortcut: "Alt+S" },
  { id: "slicer.delete", label: "Datenschnitt entfernen", icon: "delete" },
];

const cm = (px: number) => `${(px / PX_PER_CM).toLocaleString("de-DE", { maximumFractionDigits: 2 })} cm`;
const parseCm = (v: string): number | null => {
  const x = Number(v.replace(/cm/i, "").trim().replace(",", "."));
  return Number.isFinite(x) && x > 0 ? x * PX_PER_CM : null;
};

export function slicersFeature(app: CalcApp) {
  const layer = h("div", { class: "slicer-layer" });
  let selected: Slicer | null = null;
  let multi = false;
  let shownList: Slicer[] | undefined;
  const els: HTMLElement[] = [];
  let live: { s: Slicer; box: Box } | null = null;

  const z = () => app.sheet.view.zoom / 100;
  const caches = () => app.wb.slicerCaches ?? [];
  const cacheOf = (s: Slicer) => caches().find((c) => c.name === s.cache);
  const allSlicers = () => app.wb.sheets.flatMap((sh) => sh.slicers ?? []);

  function css(): void {
    if (!document.getElementById("raabe-slicer-css")) document.head.append(h("style", { id: "raabe-slicer-css" }, CSS));
  }

  // ----- sources --------------------------------------------------------------------------------

  function tableOf(c: SlicerCache): { sheet: Sheet; t: SheetTable } | null {
    if (!c.table) return null;
    for (const sheet of app.wb.sheets) {
      const t = sheet.tables?.find((x) => x.name.toLowerCase() === c.table!.name.toLowerCase());
      if (t) return { sheet, t };
    }
    return null;
  }
  function pivotsOf(c: SlicerCache): { sheet: Sheet; p: PivotTable }[] {
    return (c.pivots ?? []).flatMap((r) => {
      const sheet = app.wb.sheets.find((s) => s.name === r.sheet);
      const p = sheet?.pivots?.find((x) => x.name === r.name);
      return sheet && p ? [{ sheet, p }] : [];
    });
  }
  const cellText = (sheet: Sheet, row: number, col: number) => {
    const c = sheet.get(row, col);
    return c && c.v !== null && c.v !== "" ? shown(c, app.wb.styles.get(c.s)).text : "";
  };

  /** The slicer's items as Excel shows them. */
  /** Items are worked out again after each change (undo step) of the workbook; paints while scrolling reuse them. */
  const memo = new Map<SlicerCache, { stamp: unknown; items: SlicerItem[] }>();
  const stamp = () => {
    const h = app.history as unknown as { undoStack?: unknown[]; redoStack?: unknown[] };
    return `${app.wb.sheets.length}:${h.undoStack?.length}:${h.redoStack?.length}:${String(h.undoStack?.length ? memoIds.get(h.undoStack.at(-1) as object) ?? memoIds.set(h.undoStack.at(-1) as object, ++memoSeq).get(h.undoStack.at(-1) as object) : 0)}`;
  };
  const memoIds = new WeakMap<object, number>();
  let memoSeq = 0;
  function items(c: SlicerCache): SlicerItem[] {
    const m = memo.get(c), now = stamp();
    if (m && m.stamp === now) return m.items;
    const list = computeItems(c);
    memo.set(c, { stamp: now, items: list });
    return list;
  }
  function computeItems(c: SlicerCache): SlicerItem[] {
    const tab = tableOf(c);
    if (tab) {
      const { sheet, t } = tab;
      const d = tableData(t);
      const rows: { text: string; num: number | null }[][] = [];
      for (let r = d.r0; r <= d.r1; r++) {
        const row: { text: string; num: number | null }[] = [];
        for (let k = 0; k <= d.c1 - d.c0; k++) {
          const v = sheet.value(r, d.c0 + k);
          row.push({ text: cellText(sheet, r, d.c0 + k), num: typeof v === "number" ? v : null });
        }
        rows.push(row);
      }
      const f = t.filters ?? {};
      const others = Object.entries(f).filter(([k]) => Number(k) !== c.table!.column).map(([k, v]) => ({ col: Number(k), values: v.map((x) => (x === "" ? BLANK_ITEM : x)) }));
      return slicerItems(rows, c.table!.column, f[c.table!.column]?.map((x) => (x === "" ? BLANK_ITEM : x)), others, c);
    }
    const ps = pivotsOf(c);
    if (!ps.length) return [];
    const { p } = ps[0];
    const data = app.pivots.data(p);
    const field = fieldOf(c, p);
    if (field < 0) return [];
    const rows = pivotSlicerRows(data);
    const hiddenOf = (f: number) => p.hidden?.[f];
    const allOf = (f: number) => [...new Set(rows.map((r) => (r[f]?.text === "" ? BLANK_ITEM : r[f]?.text ?? BLANK_ITEM)))];
    const others = Object.keys(p.hidden ?? {}).map(Number).filter((f) => f !== field).map((f) => ({ col: f, values: allOf(f).filter((x) => !hiddenOf(f)!.includes(x)) }));
    const hid = hiddenOf(field);
    return slicerItems(rows, field, hid ? allOf(field).filter((x) => !hid.includes(x)) : undefined, others, c);
  }
  function fieldOf(c: SlicerCache, p: PivotTable): number {
    const names = (app.pivots.data(p)[0] ?? []).map((v) => (v === null ? "" : String(v)));
    return names.findIndex((n) => n.toLowerCase() === c.sourceName.toLowerCase());
  }
  const filtered = (c: SlicerCache) => items(c).some((i) => !i.selected);

  // ----- filtering --------------------------------------------------------------------------------

  /** Applies a selection (null: all) to the cache's table or PivotTables, as one undo step. */
  function select(c: SlicerCache, chosen: string[] | null, label = "Datenschnitt"): void {
    memo.clear();
    const all = items(c).map((i) => i.text);
    if (chosen && chosen.length >= all.length && all.every((x) => chosen!.includes(x))) chosen = null;
    const tab = tableOf(c);
    if (tab) {
      const { sheet } = tab;
      app.change(label, () => {
        app.history.whole(sheet);
        const t = sheet.tables!.find((x) => x.name === tab.t.name)!;
        const f = { ...(t.filters ?? {}) };
        if (chosen === null) delete f[c.table!.column];
        else f[c.table!.column] = chosen.map((x) => (x === BLANK_ITEM ? "" : x));
        const next: SheetTable = { ...t, filters: Object.keys(f).length ? f : undefined };
        sheet.tables = sheet.tables!.map((x) => (x === t ? next : x));
        hideRows(sheet, next);
      });
    } else {
      const ps = pivotsOf(c);
      if (!ps.length) return;
      app.change(label, () => {
        for (const { sheet, p } of ps) {
          app.history.whole(sheet);
          const field = fieldOf(c, p);
          if (field < 0) continue;
          const allItems = [...new Set(pivotSlicerRows(app.pivots.data(p)).map((r) => (r[field]?.text || BLANK_ITEM)))];
          p.hidden = pivotHiddenAfter(p, field, allItems, chosen);
          app.pivots.write(sheet, p);
        }
      });
    }
    refresh();
  }

  /** The table's data rows shown or hidden by its filters. */
  function hideRows(sheet: Sheet, t: SheetTable): void {
    const d = tableData(t);
    for (let r = d.r0; r <= d.r1; r++) {
      const show = passesTableFilters(t.filters, (k) => cellText(sheet, r, d.c0 + k));
      const info = sheet.rows.get(r);
      if (!!info?.hidden === !show) continue;
      sheet.rows.set(r, { ...(info ?? {}), hidden: show ? undefined : true });
    }
  }

  /** The item Umschalt+click starts from: the one clicked last (D-122). */
  let anchorItem: { s: Slicer; i: number } | null = null;

  /** The items from i to j (in the order shown), with `keep` (Strg: the items chosen before). */
  function rangeSelect(c: SlicerCache, list: SlicerItem[], i: number, j: number, keep: string[]): void {
    const [a, b] = [Math.min(i, j), Math.max(i, j)];
    const span = list.slice(a, b + 1).map((x) => x.text);
    select(c, [...new Set([...keep, ...span])]);
  }

  /** A click on a button: one item, or (Strg / Mehrfachauswahl) the item added or taken away. */
  function clickItem(s: Slicer, it: SlicerItem, add: boolean): void {
    const c = cacheOf(s);
    if (!c) return;
    const list = items(c);
    const isFiltered = list.some((i) => !i.selected);
    let chosen: string[];
    if (add || multi) {
      // Unfiltered, the first click with Mehrfachauswahl takes the item away (Excel).
      const sel = list.filter((i) => i.selected).map((i) => i.text);
      chosen = sel.includes(it.text) ? sel.filter((x) => x !== it.text) : [...sel, it.text];
      if (!chosen.length) return;
    } else {
      if (isFiltered && list.filter((i) => i.selected).length === 1 && it.selected) return;
      chosen = [it.text];
    }
    select(c, chosen);
  }

  // ----- places ---------------------------------------------------------------------------------

  function anchorAt(x: number, y: number): Slicer["from"] {
    const hit = app.grid.hitCell(Math.max(0, x), Math.max(0, y));
    const c = app.grid.rectOnCanvas({ r0: hit.row, c0: hit.col, r1: hit.row, c1: hit.col })!;
    return { col: hit.col, row: hit.row, x: Math.max(0, (x - c.x) / z()), y: Math.max(0, (y - c.y) / z()) };
  }
  function boxOf(s: Slicer): Box | null {
    const a = app.grid.rectOnCanvas({ r0: s.from.row, c0: s.from.col, r1: s.from.row, c1: s.from.col });
    const b = app.grid.rectOnCanvas({ r0: s.to.row, c0: s.to.col, r1: s.to.row, c1: s.to.col });
    if (!a || !b) return null;
    const x = a.x + s.from.x * z(), y = a.y + s.from.y * z();
    return { x, y, w: b.x + s.to.x * z() - x, h: b.y + s.to.y * z() - y };
  }
  function placed(s: Slicer, b: Box): Slicer {
    return { ...s, from: anchorAt(b.x, b.y), to: anchorAt(b.x + b.w, b.y + b.h), size: { w: b.w / z(), h: b.h / z() } };
  }
  /** Size at 100 % (px). */
  function sizeOf(s: Slicer): { w: number; h: number } {
    const b = boxOf(s);
    return b ? { w: b.w / z(), h: b.h / z() } : s.size ?? { w: NEW_W, h: NEW_H };
  }

  // ----- changes --------------------------------------------------------------------------------

  /** Changes a slicer of the active sheet (one undo step); `fn` returns the new slicer. */
  function edit(label: string, s: Slicer | null, fn: (s: Slicer) => Slicer): void {
    if (!s) return;
    const sheet = app.sheet;
    let next: Slicer | null = null;
    app.change(label, () => {
      app.history.whole(sheet);
      next = fn(structuredClone(s));
      sheet.slicers = (sheet.slicers ?? []).map((x) => (x === s ? next! : x));
    });
    if (selected === s) selected = next;
    refresh();
  }
  function editCache(label: string, c: SlicerCache, fn: (c: SlicerCache) => void): void {
    app.change(label, () => {
      app.history.whole(app.sheet);
      const next = structuredClone(c);
      fn(next);
      app.wb.slicerCaches = caches().map((x) => (x === c ? next : x));
    });
    refresh();
  }

  /** Adds slicers for the chosen columns (table) or fields (PivotTable), placed like Excel's, cascaded. */
  function add(sources: { sourceName: string; table?: { name: string; column: number }; pivot?: { sheet: string; name: string } }[]): void {
    if (!sources.length) return;
    const sheet = app.sheet;
    const view = app.grid.el.getBoundingClientRect();
    const v = sheet.view;
    const tl = app.grid.rectOnCanvas({ r0: v.top, c0: v.left, r1: v.top, c1: v.left }) ?? { x: 0, y: 0, w: 0, h: 0 };
    // The middle of the visible cells (Excel puts a new slicer there), each next one a little lower right.
    const cx = tl.x + (view.width - 60) / 2 - (NEW_W * z()) / 2, cy = tl.y + (view.height - 60) / 2 - (NEW_H * z()) / 2;
    let made: Slicer | null = null;
    app.change("Datenschnitt einfügen", () => {
      app.history.whole(sheet);
      const list = [...(sheet.slicers ?? [])];
      const cs = [...caches()];
      sources.forEach((src, k) => {
        const name = nextCacheName([...cs.map((c) => c.name), ...app.wb.names.map((n) => n.name)], src.sourceName);
        const c: SlicerCache = { name, sourceName: src.sourceName };
        if (src.table) c.table = src.table;
        if (src.pivot) c.pivots = [src.pivot];
        cs.push(c);
        app.wb.names.push({ name, ref: "#N/A" });
        const sname = nextSlicerName([...allSlicers(), ...list].map((x) => x.name), src.sourceName);
        const off = k * 26 * z();
        const s0: Slicer = { name: sname, cache: name, caption: src.sourceName, rowHeight: SLICER_ROW_HEIGHT, from: { col: 0, row: 0, x: 0, y: 0 }, to: { col: 0, row: 0, x: 0, y: 0 } };
        const s = placed(s0, { x: Math.max(tl.x, cx + off), y: Math.max(tl.y, cy + off), w: NEW_W * z(), h: NEW_H * z() });
        list.push(s);
        made = s;
      });
      app.wb.slicerCaches = cs;
      sheet.slicers = list;
    });
    selected = made;
    refresh();
    app.slicerSelectionChanged();
  }

  function remove(s: Slicer | null): void {
    if (!s) return;
    const sheet = app.sheet;
    app.change("Datenschnitt entfernen", () => {
      app.history.whole(sheet);
      sheet.slicers = (sheet.slicers ?? []).filter((x) => x !== s);
      if (!sheet.slicers.length) sheet.slicers = undefined;
      // A cache no other slicer uses goes, with its name; the filter it set stays (Excel).
      if (!allSlicers().some((x) => x.cache === s.cache)) {
        app.wb.slicerCaches = caches().filter((c) => c.name !== s.cache);
        if (!app.wb.slicerCaches.length) app.wb.slicerCaches = undefined;
        app.wb.names = app.wb.names.filter((n) => n.name !== s.cache);
      }
    });
    if (selected === s) selected = null;
    refresh();
    app.slicerSelectionChanged();
  }

  // ----- dialogs ----------------------------------------------------------------------------------

  /** Datenschnitt auswählen (Excel's dialog behind Datenschnitt einfügen): one box per column or field. */
  async function insertDialog(): Promise<void> {
    if (app.grid.isEditing) app.grid.commitEdit("none");
    const a = app.sheet.view.active;
    const t = app.tables.tableAt(a.row, a.col);
    const p = t ? null : app.pivots.current();
    if (!t && !p) {
      await messageBox("RaabeCalc", "Ein Datenschnitt filtert eine Tabelle oder eine PivotTable. Wählen Sie zuerst eine Zelle darin aus.");
      return;
    }
    const names = t ? t.columns.map((c) => c.name) : (app.pivots.data(p!)[0] ?? []).map((v) => (v === null ? "" : String(v)));
    const boxes = names.map((n) => h("input", { type: "checkbox", "aria-label": n }) as HTMLInputElement);
    const body = h("div", { class: "slicer-dlg", style: "display:flex;flex-direction:column;gap:4px;min-width:220px;max-height:340px;overflow:auto" },
      ...names.map((n, i) => h("label", null, boxes[i], n)));
    css();
    const okHolder: { btn?: HTMLButtonElement } = {};
    const update = () => { if (okHolder.btn) okHolder.btn.disabled = !boxes.some((b) => b.checked); };
    for (const b of boxes) b.addEventListener("change", update);
    queueMicrotask(() => {
      okHolder.btn = [...document.querySelectorAll<HTMLButtonElement>("[role=dialog] button")].find((x) => x.textContent === "OK");
      update();
    });
    const ok = await showDialog({ title: "Datenschnitt auswählen", body, buttons: [{ label: "OK", value: true, primary: true }, { label: "Abbrechen", value: false }], cancel: false });
    if (!ok) return;
    const chosen = names.map((n, i) => ({ n, i })).filter((x) => boxes[x.i].checked);
    if (!chosen.length) return;
    add(chosen.map((x) => (t ? { sourceName: x.n, table: { name: t.name, column: x.i } } : { sourceName: x.n, pivot: { sheet: app.sheet.name, name: p!.name } })));
  }

  /** Datenschnitteinstellungen (labels as recorded). */
  async function settings(s: Slicer | null = selected): Promise<void> {
    if (!s) return;
    const c = cacheOf(s);
    if (!c) return;
    const input = (v: string) => h("input", { type: "text", value: v, style: "width:220px" }) as HTMLInputElement;
    const check = (on: boolean) => h("input", { type: "checkbox", ...(on ? { checked: true } : {}) }) as HTMLInputElement;
    const radio = (on: boolean) => h("input", { type: "radio", name: "slsort", ...(on ? { checked: true } : {}) }) as HTMLInputElement;
    const name = input(s.name), caption = input(s.caption);
    const head = check(s.showCaption !== false);
    const asc = radio((c.sort ?? "ascending") !== "descending"), desc = radio(c.sort === "descending");
    const lists = check(c.customListSort !== false);
    const hide = check(!!c.hideNoData);
    const mark = check((c.crossFilter ?? "showItemsWithDataAtTop") !== "none");
    const last = check((c.crossFilter ?? "showItemsWithDataAtTop") === "showItemsWithDataAtTop");
    const sync = () => { mark.disabled = hide.checked; last.disabled = hide.checked || !mark.checked; caption.disabled = !head.checked; };
    for (const x of [hide, mark, head]) x.addEventListener("change", sync);
    sync();
    const body = h("div", { class: "slicer-dlg", style: "display:flex;flex-direction:column;gap:8px;min-width:420px" },
      h("div", null, "Quellname: ", c.sourceName),
      h("div", null, "In Formeln zu verwendender Name: ", c.name),
      h("label", null, "Name:", name),
      h("fieldset", null, h("legend", null, "Kopfzeile"), h("label", null, head, "Kopfzeile anzeigen"), h("label", null, "Beschriftung:", caption)),
      h("fieldset", null, h("legend", null, "Elementsortierung und -filterung"),
        h("label", null, asc, "Aufsteigend (A bis Z)"), h("label", null, desc, "Absteigend (Z bis A)"),
        h("label", null, lists, "Beim Sortieren benutzerdefinierte Listen verwenden"),
        h("label", null, hide, "Elemente ohne Daten ausblenden"),
        h("label", null, mark, "Elemente ohne Daten visuell kennzeichnen"),
        h("label", null, last, "Elemente ohne Daten zuletzt anzeigen")));
    const ok = await showDialog({ title: "Datenschnitteinstellungen", body, buttons: [{ label: "OK", value: true, primary: true }, { label: "Abbrechen", value: false }], cancel: false });
    if (!ok) return;
    const newName = name.value.trim() || s.name;
    if (newName !== s.name && allSlicers().some((x) => x !== s && x.name.toLowerCase() === newName.toLowerCase())) {
      await messageBox("RaabeCalc", "Diesen Namen trägt schon ein anderer Datenschnitt.");
      return;
    }
    const sheet = app.sheet;
    let next: Slicer | null = null;
    app.change("Datenschnitteinstellungen", () => {
      app.history.whole(sheet);
      next = { ...structuredClone(s), name: newName, caption: caption.value, showCaption: head.checked ? undefined : false };
      sheet.slicers = (sheet.slicers ?? []).map((x) => (x === s ? next! : x));
      const nc = structuredClone(c);
      nc.sort = desc.checked ? "descending" : undefined;
      nc.customListSort = lists.checked ? undefined : false;
      nc.hideNoData = hide.checked || undefined;
      nc.crossFilter = !mark.checked ? "none" : last.checked ? undefined : "showItemsWithNoData";
      app.wb.slicerCaches = caches().map((x) => (x === c ? nc : x));
    });
    if (selected === s) selected = next;
    refresh();
  }

  /** Berichtsverbindungen: which PivotTables of the same source the slicer filters (table slicers: greyed, as Excel). */
  async function connections(s: Slicer | null = selected): Promise<void> {
    const c = s ? cacheOf(s) : undefined;
    if (!c?.pivots) return;
    const first = pivotsOf(c)[0];
    if (!first) return;
    const same = (p: PivotTable) => JSON.stringify(p.source) === JSON.stringify(first.p.source);
    const all = app.wb.sheets.flatMap((sheet) => (sheet.pivots ?? []).filter(same).map((p) => ({ sheet, p })));
    const boxes = all.map((x) => h("input", { type: "checkbox", ...(c.pivots!.some((r) => r.sheet === x.sheet.name && r.name === x.p.name) ? { checked: true } : {}) }) as HTMLInputElement);
    const body = h("div", { class: "slicer-dlg", style: "display:grid;grid-template-columns:auto 1fr 1fr;gap:4px 10px;min-width:360px" },
      h("span"), h("b", null, "Name"), h("b", null, "Blatt"),
      ...all.flatMap((x, i) => [boxes[i], h("span", null, x.p.name), h("span", null, x.sheet.name)]));
    const ok = await showDialog({ title: `Berichtsverbindungen (${s!.name})`, body, buttons: [{ label: "OK", value: true, primary: true }, { label: "Abbrechen", value: false }], cancel: false });
    if (!ok) return;
    const chosen = all.filter((_, i) => boxes[i].checked).map((x) => ({ sheet: x.sheet.name, name: x.p.name }));
    if (!chosen.length) return;
    editCache("Berichtsverbindungen", c, (n) => { n.pivots = chosen; });
  }

  /** Größe und Eigenschaften: height and width of the frame in cm. */
  async function sizeDialog(): Promise<void> {
    const s = selected;
    if (!s) return;
    const sz = sizeOf(s);
    const hi = h("input", { type: "text", value: cm(sz.h), style: "width:90px" }) as HTMLInputElement;
    const wi = h("input", { type: "text", value: cm(sz.w), style: "width:90px" }) as HTMLInputElement;
    const body = h("div", { class: "slicer-dlg", style: "display:flex;flex-direction:column;gap:8px" }, h("label", null, "Höhe:", hi), h("label", null, "Breite:", wi));
    const ok = await showDialog({ title: "Datenschnitt formatieren", body, buttons: [{ label: "OK", value: true, primary: true }, { label: "Abbrechen", value: false }], cancel: false });
    if (!ok) return;
    const H = parseCm(hi.value), W = parseCm(wi.value);
    if (H || W) resize(s, W ?? sz.w, H ?? sz.h);
  }

  function resize(s: Slicer, w: number, hh: number): void {
    const b = boxOf(s);
    if (!b) return;
    edit("Größe", s, (x) => placed(x, { x: b.x, y: b.y, w: Math.max(20, w) * z(), h: Math.max(20, hh) * z() }));
  }

  // ----- drawing ----------------------------------------------------------------------------------

  const themeColor = (c: TsColor | undefined, fallback: string): string => {
    if (!c) return fallback;
    if (typeof c === "string") return c;
    return tinted(app.wb.theme[c.t] ?? OFFICE_2013[c.t] ?? "#000000", c.tint ?? 0);
  };
  const styleOf = (s: Slicer) => SLICER_STYLES.find((x) => x.name === (s.style ?? "SlicerStyleLight1")) ?? SLICER_STYLES[0];
  const el = (s: Slicer, k: string): SlicerElement => styleOf(s).elements[k] ?? {};
  const font = () => `"${app.wb.themeFonts?.minor ?? "Calibri"}", Calibri, Carlito, sans-serif`;

  /** The item buttons' geometry (px at 100 %). */
  function geometry(s: Slicer, w: number) {
    const cols = Math.max(1, s.columns ?? 1);
    const bh = ((s.rowHeight ?? SLICER_ROW_HEIGHT) / EMU_PER_PX);
    const inner = w - 2 * MARGIN * PX_PER_PT;
    const bw = Math.max(4, (inner - (cols - 1) * GAP_COL * PX_PER_PT) / cols);
    return { cols, bw, bh, top: (s.showCaption === false ? 0 : HEAD * PX_PER_PT) + AFTER_HEAD * PX_PER_PT, left: MARGIN * PX_PER_PT };
  }

  function render(s: Slicer, root: HTMLElement, size: { w: number; h: number }): void {
    const c = cacheOf(s);
    const whole = el(s, "wholeTable"), head = el(s, "headerRow");
    const border = themeColor(whole.borders?.left, "#156082");
    Object.assign(root.style, { width: `${size.w}px`, height: `${size.h}px`, border: `1px solid ${border}`, font: `${11 * PX_PER_PT}px ${font()}`, color: themeColor(whole.font, "#000") });
    root.replaceChildren();
    const its = c ? items(c) : [];
    const g = geometry(s, size.w);
    if (s.showCaption !== false) {
      const isOn = c ? its.some((i) => !i.selected) : false;
      const multiIc = h("span", { class: `sl-ic${multi ? " pressed" : ""}`, title: "Mehrfachauswahl (Alt+S)", "data-act": "multi" });
      multiIc.innerHTML = '<svg width="16" height="16" viewBox="0 0 16 16"><path d="M1 3l1.5 1.5L5 2M1 8l1.5 1.5L5 7M1 13l1.5 1.5L5 12" fill="none" stroke="#555" stroke-width="1.1"/><path d="M7 3.5h8M7 8.5h8M7 13.5h8" stroke="#555"/></svg>';
      const clear = h("span", { class: "sl-ic", title: "Filter löschen (Alt+C)", "data-act": "clear" });
      clear.innerHTML = `<svg width="16" height="16" viewBox="0 0 16 16"><path d="M1.5 2h11l-4.2 5v5.5L6 14V7z" fill="none" stroke="${isOn ? "#555" : "#b5b5b5"}" stroke-width="1.1"/>${isOn ? '<path d="M10 10l4.5 4.5M14.5 10L10 14.5" stroke="#d24726" stroke-width="1.5"/>' : '<path d="M10 10l4.5 4.5M14.5 10L10 14.5" stroke="#b5b5b5" stroke-width="1.2"/>'}</svg>`;
      root.append(h("div", { class: "sl-head", style: `height:${HEAD * PX_PER_PT}px;border-bottom:1px solid ${themeColor(head.borders?.bottom, border)};font-weight:${head.bold ? 700 : 400};color:${themeColor(head.font, "#000")}` },
        h("span", { class: "sl-cap" }, s.caption), multiIc, clear));
    }
    const body = h("div", { class: "sl-body", style: `top:${(s.showCaption === false ? 0 : HEAD * PX_PER_PT + 1)}px` });
    const step = g.bh + GAP_ROW * PX_PER_PT;
    const top0 = AFTER_HEAD * PX_PER_PT;
    its.forEach((it, i) => {
      const r = Math.floor(i / g.cols), k = i % g.cols;
      const kind = `${it.selected ? "selected" : "unselected"}ItemWith${it.hasData ? "" : "No"}Data`;
      const e = el(s, kind);
      const b = h("div", {
        class: "sl-btn", title: it.text,
        style: `left:${g.left + k * (g.bw + GAP_COL * PX_PER_PT)}px;top:${top0 + r * step}px;width:${g.bw}px;height:${g.bh}px;`
          + `background:${themeColor(e.fill, "#fff")};border-color:${themeColor(e.borders?.left, "#ccc")};color:${themeColor(e.font, "#000")}`,
      }, it.text);
      b.dataset.item = String(i);
      body.append(b);
    });
    // The space below the last row keeps the margin when the body scrolls.
    const rowsN = Math.ceil(its.length / g.cols);
    body.append(h("div", { style: `position:absolute;left:0;width:1px;height:1px;top:${top0 + rowsN * step + AFTER_HEAD * PX_PER_PT}px` }));
    root.append(body);
    for (const k of HANDLES) {
      const hd = h("span", { class: `sl-h ${k}`, "data-h": k, style: `left:${POS[k][0] * 100}%;top:${POS[k][1] * 100}%` });
      root.append(hd);
    }
    root.dataset.items = JSON.stringify(its.map((x) => x.text));
  }

  /** Called after every paint of the grid. */
  function place(): void {
    const list = app.sheet.slicers;
    if (!list?.length) {
      if (layer.isConnected) layer.remove();
      shownList = undefined;
      if (selected) { selected = null; app.slicerSelectionChanged(); }
      return;
    }
    css();
    if (!layer.isConnected) {
      const canvas = app.grid.el.querySelector(".grid-canvas");
      app.grid.el.insertBefore(layer, canvas ? canvas.nextSibling : app.grid.el.firstChild);
    }
    if (selected && !list.includes(selected)) {
      const name = selected.name;
      selected = list.find((x) => x.name === name) ?? null;
      if (!selected) queueMicrotask(() => app.slicerSelectionChanged());
    }
    if (shownList !== list) {
      layer.replaceChildren();
      els.length = 0;
      for (const s of list) {
        const e = h("div", { class: "slicer", tabindex: "-1" });
        e.addEventListener("pointerdown", (ev) => down(ev, s, e));
        e.addEventListener("contextmenu", (ev) => { ev.preventDefault(); contextMenu(s, ev); });
        e.addEventListener("keydown", (ev) => key(ev, s));
        els.push(e);
        layer.append(e);
      }
      shownList = list;
    }
    const v = app.sheet.view;
    const tl = app.grid.rectOnCanvas({ r0: v.top, c0: v.left, r1: v.top, c1: v.left });
    if (!tl) return;
    Object.assign(layer.style, { left: `${tl.x}px`, top: `${tl.y}px` });
    list.forEach((s, i) => {
      const e = els[i];
      const b = live?.s === s ? live.box : boxOf(s);
      if (!b || b.w <= 0 || b.h <= 0) { e.style.display = "none"; return; }
      const size = { w: b.w / z(), h: b.h / z() };
      const key = JSON.stringify([s, size.w.toFixed(1), size.h.toFixed(1), cacheOf(s) ? items(cacheOf(s)!) : null, multi, app.wb.themeFonts?.minor, app.wb.theme.accent1]);
      if (e.dataset.key !== key) { render(s, e, size); e.dataset.key = key; }
      e.classList.toggle("on", s === selected);
      Object.assign(e.style, { display: "", left: `${b.x - tl.x}px`, top: `${b.y - tl.y}px`, transform: `scale(${z()})` });
    });
  }

  function refresh(): void {
    shownList = undefined;
    memo.clear();
    app.grid.contentChanged();
  }

  function choose(s: Slicer | null): void {
    if (selected === s) return;
    selected = s;
    app.slicerSelectionChanged();
    app.grid.contentChanged();
  }

  // ----- mouse and keys ---------------------------------------------------------------------------

  function down(ev: PointerEvent, s: Slicer, e: HTMLElement): void {
    if (ev.button !== 0) { choose(s); return; }
    ev.stopPropagation();
    const t = ev.target as HTMLElement;
    const act = t.closest<HTMLElement>("[data-act]")?.dataset.act;
    const btn = t.closest<HTMLElement>(".sl-btn");
    const handle = t.closest<HTMLElement>(".sl-h")?.dataset.h as Handle | undefined;
    choose(s);
    e.focus({ preventScroll: true });
    const c = cacheOf(s);
    if (act === "multi") { multi = !multi; refresh(); return; }
    if (act === "clear") { if (c) select(c, null, "Filter löschen"); return; }
    if (btn && c) {
      // D-122: dragging over the buttons chooses the items between (with Strg: added to the choice); Umschalt+click
      // chooses from the item clicked last to this one.
      const list = items(c);
      const i0 = Number(btn.dataset.item);
      const before = list.filter((i) => i.selected).map((i) => i.text);
      const filtered = before.length < list.length;
      let i1 = i0, dragged = false;
      const at = (m: PointerEvent): number | null => {
        const el = (document.elementFromPoint(m.clientX, m.clientY) as HTMLElement | null)?.closest<HTMLElement>(".sl-btn");
        return el && e.contains(el) ? Number(el.dataset.item) : null;
      };
      const mark = () => {
        const [a, b] = [Math.min(i0, i1), Math.max(i0, i1)];
        e.querySelectorAll<HTMLElement>(".sl-btn").forEach((x) => x.classList.toggle("sl-dragging", Number(x.dataset.item) >= a && Number(x.dataset.item) <= b));
      };
      const mv = (m: PointerEvent) => { const j = at(m); if (j === null || j === i1) return; i1 = j; dragged = true; mark(); };
      const upb = () => {
        removeEventListener("pointermove", mv);
        removeEventListener("pointerup", upb);
        if (dragged) rangeSelect(c, list, i0, i1, ev.ctrlKey && filtered ? before : []);
        else if (ev.shiftKey && anchorItem && anchorItem.s === s) rangeSelect(c, list, anchorItem.i, i0, (ev.ctrlKey || multi) && filtered ? before : []);
        else { const it = list[i0]; if (it) clickItem(s, it, ev.ctrlKey); }
        if (!ev.shiftKey) anchorItem = { s, i: dragged ? i0 : i0 };
      };
      addEventListener("pointermove", mv);
      addEventListener("pointerup", upb);
      return;
    }
    // Header, border or handle: move or size.
    const start = boxOf(s);
    if (!start) return;
    const x0 = ev.clientX, y0 = ev.clientY;
    let moved = false;
    const move = (m: PointerEvent) => {
      const dx = m.clientX - x0, dy = m.clientY - y0;
      if (!moved && Math.abs(dx) + Math.abs(dy) < 3) return;
      moved = true;
      let b = { ...start };
      if (!handle) b = { ...b, x: start.x + dx, y: start.y + dy };
      else {
        const [hx, hy] = POS[handle];
        if (hx === 0) { b.x = start.x + Math.min(dx, start.w - 10); b.w = start.w - (b.x - start.x); }
        if (hx === 1) b.w = Math.max(10, start.w + dx);
        if (hy === 0) { b.y = start.y + Math.min(dy, start.h - 10); b.h = start.h - (b.y - start.y); }
        if (hy === 1) b.h = Math.max(10, start.h + dy);
      }
      live = { s, box: b };
      app.grid.contentChanged();
    };
    const up = () => {
      removeEventListener("pointermove", move);
      removeEventListener("pointerup", up);
      const done = live;
      live = null;
      if (moved && done) edit(handle ? "Größe" : "Verschieben", s, (x) => placed(x, done.box));
      else app.grid.contentChanged();
    };
    addEventListener("pointermove", move);
    addEventListener("pointerup", up);
  }

  function key(ev: KeyboardEvent, s: Slicer): void {
    const c = cacheOf(s);
    if (ev.key === "Delete") { ev.preventDefault(); remove(s); return; }
    if (ev.altKey && (ev.key === "c" || ev.key === "C") && c) { ev.preventDefault(); select(c, null, "Filter löschen"); return; }
    if (ev.altKey && (ev.key === "s" || ev.key === "S")) { ev.preventDefault(); multi = !multi; refresh(); return; }
    if (ev.key === "Escape") { ev.preventDefault(); choose(null); app.grid.focus(); return; }
    const d: Record<string, [number, number]> = { ArrowLeft: [-1, 0], ArrowRight: [1, 0], ArrowUp: [0, -1], ArrowDown: [0, 1] };
    if (d[ev.key]) {
      ev.preventDefault();
      const b = boxOf(s);
      if (b) edit("Verschieben", s, (x) => placed(x, { ...b, x: b.x + d[ev.key][0] * z(), y: b.y + d[ev.key][1] * z() }));
    }
  }

  function contextMenu(s: Slicer, ev: MouseEvent): void {
    choose(s);
    const c = cacheOf(s);
    if (!c) return;
    // Entries of Excel's slicer context menu we offer (not recorded through UI Automation: labels as Excel 365 shows them).
    const entries: MenuEntry[] = [
      { label: "Von A bis Z sortieren", icon: "sort-az", run: () => editCache("Sortieren", c, (n) => { n.sort = undefined; }) },
      { label: "Von Z bis A sortieren", icon: "sort-za", run: () => editCache("Sortieren", c, (n) => { n.sort = "descending"; }) },
      "-",
      { label: `Filter aus "${s.caption}" löschen`, icon: "filter-clear", run: () => select(c, null, "Filter löschen"), disabled: !filtered(c) },
      { label: "Mehrfachauswahl", checked: multi, run: () => { multi = !multi; refresh(); } },
      "-",
      { label: `"${s.name}" entfernen`, icon: "delete", run: () => remove(s) },
      "-",
      { label: "Datenschnitteinstellungen...", run: () => void settings(s) },
      { label: "Größe und Eigenschaften...", run: () => void sizeDialog() },
    ];
    showMenu(app.commands, entries, new DOMRect(ev.clientX, ev.clientY, 0, 0));
  }

  // A click elsewhere ends the selection.
  addEventListener("pointerdown", (e) => {
    const t = e.target as HTMLElement;
    if (!selected || t.closest?.(".slicer") || t.closest?.(".ribbon-card, .tab, .taskpane, .popup, .rpanel, .menu, .menu-root, .dialog-layer, [role=dialog], [role=menu]")) return;
    choose(null);
  }, true);

  // ----- ribbon ---------------------------------------------------------------------------------

  function stylesPanel(close: () => void): HTMLElement {
    css();
    const groups: [string, string[]][] = [["Hell", SLICER_STYLES.filter((x) => /Light/.test(x.name)).map((x) => x.name)],
      ["Sonstige", SLICER_STYLES.filter((x) => /Other/.test(x.name)).map((x) => x.name)], ["Dunkel", SLICER_STYLES.filter((x) => /Dark/.test(x.name)).map((x) => x.name)]];
    const box = h("div", { class: "slicer-gallery" });
    for (const [label, names] of groups) {
      box.append(h("div", { class: "sg-head" }, label));
      for (const n of names) {
        const st = SLICER_STYLES.find((x) => x.name === n)!;
        const e = (k: string) => st.elements[k] ?? {};
        const tile = h("div", { class: `sg-tile${(selected?.style ?? "SlicerStyleLight1") === n ? " on" : ""}`, title: st.local });
        const bd = themeColor(e("wholeTable").borders?.left, "#999");
        tile.innerHTML = `<svg width="52" height="52" viewBox="0 0 52 52"><rect x="0.5" y="0.5" width="51" height="51" fill="#fff" stroke="${bd}"/>`
          + `<rect x="4" y="4" width="26" height="4" fill="${themeColor(e("headerRow").font, "#333")}" opacity="0.6"/>`
          + [["selectedItemWithData", 13], ["selectedItemWithData", 24], ["unselectedItemWithData", 35]].map(([k, y]) => {
            const x = e(k as string);
            return `<rect x="4" y="${y}" width="44" height="9" rx="1.5" fill="${themeColor(x.fill, "#fff")}" stroke="${themeColor(x.borders?.left, "#ccc")}"/>`;
          }).join("") + "</svg>";
        tile.addEventListener("click", () => { close(); edit("Datenschnitt-Formatvorlage", selected, (x) => ({ ...x, style: n === "SlicerStyleLight1" ? undefined : n })); });
        box.append(tile);
      }
    }
    return box;
  }

  const has = () => !!selected;
  const sizeSpin = (cmd: string, which: "w" | "h", key: string): Control => ({
    type: "spinner", cmd, label: which === "h" ? "Höhe:" : "Breite:", key,
    value: () => (selected ? cm(sizeOf(selected)[which]) : ""),
    commit: (v) => { const s = selected; const x = parseCm(v); if (s && x) { const sz = sizeOf(s); resize(s, which === "w" ? x : sz.w, which === "h" ? x : sz.h); } },
    step: (d) => { const s = selected; if (s) { const sz = sizeOf(s); const step = 0.1 * PX_PER_CM * d; resize(s, which === "w" ? sz.w + step : sz.w, which === "h" ? sz.h + step : sz.h); } },
  });
  /** Schaltflächen > Breite: the buttons' width (the frame follows, as Excel's). */
  const buttonWidth = (s: Slicer) => geometry(s, sizeOf(s).w).bw;
  const setButtonWidth = (s: Slicer, bw: number) => {
    const g = geometry(s, sizeOf(s).w);
    resize(s, g.cols * bw + (g.cols - 1) * GAP_COL * PX_PER_PT + 2 * MARGIN * PX_PER_PT, sizeOf(s).h);
  };

  let tabDef: Tab | null = null;
  function tab(): Tab {
    const b = (cmd: string, key: string, size: "large" | "medium" | "small" = "small"): Control => ({ type: "button", cmd, key, size });
    return tabDef ??= {
      id: "slicer", label: "Datenschnitt", key: "Y3", contextual: true, shrink: [],
      groups: [
        { id: "slicer", label: "Datenschnitt", key: "ZS", icon: "filter", controls: [{ type: "stack", items: [
          { type: "combo", cmd: "slicer.caption", width: 130, key: "T", label: "Datenschnittbeschriftung:", value: () => selected?.caption ?? "",
            commit: (v) => edit("Datenschnittbeschriftung", selected, (x) => ({ ...x, caption: v })) },
          b("slicer.settings", "A", "medium"), b("slicer.connections", "E", "medium")] }] },
        { id: "styles", label: "Datenschnitt-Formatvorlagen", key: "ZQ", icon: "styles", controls: [
          { type: "dropdown", cmd: "slicer.styles", key: "S", size: "large", menu: [], popup: (close) => stylesPanel(close) }] },
        { id: "arrange", label: "Anordnen", key: "ZA", icon: "bring-forward", controls: [
          { type: "split", cmd: "slicer.forward", key: "GE", size: "large", menu: [{ cmd: "slicer.forward" }, { cmd: "slicer.toFront" }] },
          { type: "split", cmd: "slicer.backward", key: "GH", size: "large", menu: [{ cmd: "slicer.backward" }, { cmd: "slicer.toBack" }] },
          b("slicer.selectionPane", "Q", "large"),
          { type: "stack", items: [
            { type: "dropdown", cmd: "slicer.align", key: "DA", size: "small", menu: [{ label: "Am Raster ausrichten", run: () => {}, disabled: true }] },
            { type: "dropdown", cmd: "slicer.group", key: "GR", size: "small", menu: [] },
            { type: "dropdown", cmd: "slicer.rotate", key: "DR", size: "small", menu: [] }] }] },
        { id: "buttons", label: "Schaltflächen", key: "ZC", icon: "table", controls: [{ type: "stack", items: [
          { type: "spinner", cmd: "slicer.columns", label: "Spalten:", key: "N", value: () => String(selected?.columns ?? 1),
            commit: (v) => { const n = Math.round(Number(v)); if (n >= 1 && n <= 20000) edit("Spalten", selected, (x) => ({ ...x, columns: n > 1 ? n : undefined })); },
            step: (d) => { const n = (selected?.columns ?? 1) + d; if (n >= 1) edit("Spalten", selected, (x) => ({ ...x, columns: n > 1 ? n : undefined })); } },
          { type: "spinner", cmd: "slicer.buttonHeight", label: "Höhe:", key: "Ö", value: () => (selected ? cm((selected.rowHeight ?? SLICER_ROW_HEIGHT) / EMU_PER_PX) : ""),
            commit: (v) => { const x = parseCm(v); if (x) edit("Höhe", selected, (s) => ({ ...s, rowHeight: Math.round(x * EMU_PER_PX) })); },
            step: (d) => { const s = selected; if (s) edit("Höhe", s, (x) => ({ ...x, rowHeight: Math.max(9525, Math.round((x.rowHeight ?? SLICER_ROW_HEIGHT) + d * 0.1 * PX_PER_CM * EMU_PER_PX)) })); } },
          { type: "spinner", cmd: "slicer.buttonWidth", label: "Breite:", key: "B", value: () => (selected ? cm(buttonWidth(selected)) : ""),
            commit: (v) => { const x = parseCm(v); if (x && selected) setButtonWidth(selected, x); },
            step: (d) => { const s = selected; if (s) setButtonWidth(s, buttonWidth(s) + d * 0.1 * PX_PER_CM); } }] }] },
        { id: "size", label: "Größe", key: "ZZ", icon: "position", launcher: { cmd: "slicer.sizeDialog", key: "ß" }, controls: [{ type: "stack", items: [
          sizeSpin("slicer.height", "h", "H"), sizeSpin("slicer.width", "w", "RR")] }] },
      ],
    };
  }

  /** Ebene nach vorne / hinten among the sheet's slicers. */
  function reorder(s: Slicer | null, to: "forward" | "front" | "backward" | "back"): void {
    if (!s) return;
    const sheet = app.sheet;
    const list = [...(sheet.slicers ?? [])];
    const i = list.indexOf(s);
    if (i < 0) return;
    list.splice(i, 1);
    const j = to === "front" ? list.length : to === "back" ? 0 : to === "forward" ? Math.min(list.length, i + 1) : Math.max(0, i - 1);
    list.splice(j, 0, s);
    app.change("Anordnen", () => { app.history.whole(sheet); sheet.slicers = list; });
    refresh();
  }

  function implement(run: Run): void {
    app.commands.register(...SLICER_COMMANDS);
    const tableOrPivot = () => { const a = app.sheet.view.active; return !!app.tables.tableAt(a.row, a.col) || !!app.pivots.current(); };
    run("insert.slicer", () => void insertDialog(), () => !app.grid.isEditing);
    run("table.slicer", () => void insertDialog(), tableOrPivot);
    run("pivot.slicer", () => void insertDialog(), () => !!app.pivots.current());
    run("slicer.caption", () => {}, has);
    run("slicer.settings", () => void settings(), has);
    run("slicer.connections", () => void connections(), () => !!selected && !!cacheOf(selected)?.pivots);
    run("slicer.styles", () => {}, has);
    run("slicer.forward", () => reorder(selected, "forward"), () => has() && (app.sheet.slicers?.length ?? 0) > 1);
    run("slicer.toFront", () => reorder(selected, "front"), () => has() && (app.sheet.slicers?.length ?? 0) > 1);
    run("slicer.backward", () => reorder(selected, "backward"), () => has() && (app.sheet.slicers?.length ?? 0) > 1);
    run("slicer.toBack", () => reorder(selected, "back"), () => has() && (app.sheet.slicers?.length ?? 0) > 1);
    run("slicer.selectionPane", () => void app.commands.run("layout.selectionPane"), has);
    run("slicer.align", () => {}, has);
    run("slicer.columns", () => {}, has);
    run("slicer.buttonHeight", () => {}, has);
    run("slicer.buttonWidth", () => {}, has);
    run("slicer.height", () => {}, has);
    run("slicer.width", () => {}, has);
    run("slicer.sizeDialog", () => void sizeDialog(), has);
    run("slicer.clear", () => { const c = selected && cacheOf(selected); if (c) select(c, null, "Filter löschen"); }, () => !!selected && !!cacheOf(selected) && filtered(cacheOf(selected)!));
    run("slicer.multi", () => { multi = !multi; refresh(); }, has, () => multi);
    run("slicer.delete", () => remove(selected), has);
  }

  return {
    place, tab, implement, add, remove, select, items, insertDialog, settings,
    selected: () => selected,
    deselect: () => choose(null),
    /** Tests and the harness: the slicer of that name on the active sheet. */
    byName: (name: string) => app.sheet.slicers?.find((s) => s.name === name) ?? null,
    choose,
    cacheOf,
    /** A click on an item (as the mouse does); `add`: with Strg. */
    click: (s: Slicer, text: string, add = false) => { const c = cacheOf(s); const it = c ? items(c).find((i) => i.text === text) : undefined; if (it) clickItem(s, it, add); },
    setMulti: (on: boolean) => { multi = on; refresh(); },
    edit, editCache, sizeOf, resize,
    rendered: () => els.map((e) => ({ items: JSON.parse(e.dataset.items ?? "[]") as string[] })),
  };
}
