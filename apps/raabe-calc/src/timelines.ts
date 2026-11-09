// Zeitachsen (timelines) in RaabeCalc (D-121): Einfügen > Filter > Zeitachse and PivotTable-Analyse > Zeitachse
// einfügen open Zeitachsen einfügen (the PivotTable's date fields); the timeline on the sheet shows the field's time as
// periods at one Zeitebene (Jahre, Quartale, Monate, Tage) and filters the connected PivotTables to consecutive periods:
// a click takes one period, Umschalt+click extends, dragging across periods selects them, the selection's edges are
// dragged to widen or narrow it, dragging inside it moves it; Filter löschen in the header takes all; the Zeitebene
// menu, a scroll bar (arrows, track, thumb) and the mouse wheel move through time. The filter is the PivotTables' own
// date filter (dateBetween in whole days), as Excel keeps it. The contextual tab Zeitachse with Excel's groups and
// keytips (recorded through UI Automation, probe_excel_timeline_ui.py). Geometry measured on Excel's rendering at
// 96 dpi (probe_excel_timeline.py): 262.5 x 108 pt by default; header 28 px, label row 26 px, period area 66 px
// (labels, then the blocks 17 px high), scroll bar 24 px; a period is 54 / 34.67 / 42.67 / 30.67 px wide for Jahre /
// Quartale / Monate / Tage; the 12 built-in styles from COM (timeline-styles.gen.ts). Each change is one undo step.

import { h, messageBox, showDialog, showMenu, type Control, type MenuEntry, type Tab } from "@raabe/shell";
import {
  TIME_LEVELS, TIMELINE_H, TIMELINE_W, addPeriods, firstUpperLabel, isDateFormat, nextSlicerName, nextTimelineCacheName, periodEnd, periodLabels,
  periodStart, periodsBetween, pivotDateFilter, pivotDateFiltersAfter, selectionLabel, timelineBounds,
  type PivotTable, type Sheet, type TimeLevel, type Timeline, type TimelineCache,
} from "@raabe/sheet-engine";
import type { CalcApp } from "./app";
import { OFFICE_2013, tinted } from "./tables";
import type { TsColor } from "./table-styles.gen";
import { TIMELINE_STYLES, type TimelineElement } from "./timeline-styles.gen";

const PX_PER_CM = 96 / 2.54;
/** Heights (px at 100 %) of the parts, measured on Excel's rendering. */
const HEAD = 28, ROW = 26, AREA = 66, SCROLL = 24, MARGIN = 8, BLOCK_TOP = 42, BLOCK_H = 17;
/** A period's width per level (Jahre, Quartale, Monate, Tage). */
export const PITCH = [54, 34.67, 42.67, 30.67];

const CSS = `
.timeline-layer { position: absolute; right: 17px; bottom: 0; overflow: hidden; pointer-events: none; z-index: 4; }
.timeline-layer .timeline { position: absolute; pointer-events: auto; box-sizing: border-box; background: #fff; transform-origin: 0 0; outline: none; user-select: none; overflow: hidden; }
.timeline-layer .timeline.on { box-shadow: 0 0 0 1px #7a7a7a; }
.timeline-layer .tl-head { position: absolute; left: 0; right: 0; top: 0; box-sizing: border-box; display: flex; align-items: center; padding-left: 9px; cursor: move; }
.timeline-layer .tl-cap { flex: 1; overflow: hidden; white-space: nowrap; text-overflow: ellipsis; }
.timeline-layer .tl-ic { width: 20px; height: 20px; margin-right: 7px; display: flex; align-items: center; justify-content: center; border-radius: 2px; }
.timeline-layer .tl-ic:hover { background: #e6e6e6; }
.timeline-layer .tl-row { position: absolute; left: 0; right: 0; display: flex; align-items: center; padding: 0 9px; box-sizing: border-box; }
.timeline-layer .tl-sel { flex: 1; overflow: hidden; white-space: nowrap; }
.timeline-layer .tl-lvl { cursor: default; padding: 1px 3px; letter-spacing: .3px; }
.timeline-layer .tl-lvl:hover { background: #e6e6e6; }
.timeline-layer .tl-area { position: absolute; left: 0; right: 0; overflow: hidden; cursor: default; }
.timeline-layer .tl-lab { position: absolute; white-space: nowrap; transform: translateX(-50%); pointer-events: none; }
.timeline-layer .tl-blk { position: absolute; box-sizing: border-box; }
.timeline-layer .tl-edge { position: absolute; width: 6px; margin-left: -3px; cursor: ew-resize; z-index: 2; }
.timeline-layer .tl-scroll { position: absolute; left: ${MARGIN}px; right: ${MARGIN}px; height: 16px; box-sizing: border-box; }
.timeline-layer .tl-arrow { position: absolute; top: 0; width: 16px; height: 16px; box-sizing: border-box; border: 1px solid #c4c4c4; display: flex; align-items: center; justify-content: center; background: #fff; }
.timeline-layer .tl-track { position: absolute; left: 18px; right: 18px; top: 0; height: 16px; }
.timeline-layer .tl-span { position: absolute; top: 4px; height: 8px; }
.timeline-layer .tl-thumb { position: absolute; top: 0; height: 16px; box-sizing: border-box; border: 1px solid #c4c4c4; background: #fff; border-radius: 2px; }
.timeline-layer .tl-h { position: absolute; width: 8px; height: 8px; margin: -5px 0 0 -5px; border: 1px solid #5f5f5f; border-radius: 50%; background: #fff; display: none; z-index: 3; }
.timeline-layer .timeline.on .tl-h { display: block; }
.timeline-layer .tl-h.nw, .timeline-layer .tl-h.se { cursor: nwse-resize; }
.timeline-layer .tl-h.ne, .timeline-layer .tl-h.sw { cursor: nesw-resize; }
.timeline-layer .tl-h.n, .timeline-layer .tl-h.s { cursor: ns-resize; }
.timeline-layer .tl-h.e, .timeline-layer .tl-h.w { cursor: ew-resize; }
.timeline-gallery { display: grid; grid-template-columns: repeat(6, 64px); gap: 6px; padding: 8px; }
.timeline-gallery .tg-head { grid-column: 1 / -1; font-weight: 600; font-size: 12px; margin-top: 4px; }
.timeline-gallery .tg-tile { width: 60px; height: 44px; border: 1px solid #ddd; padding: 3px; box-sizing: border-box; cursor: pointer; }
.timeline-gallery .tg-tile:hover, .timeline-gallery .tg-tile.on { border-color: #6b9bd1; background: #e5f0fb; }
`;

type Handle = "nw" | "n" | "ne" | "e" | "se" | "s" | "sw" | "w";
const HANDLES: Handle[] = ["nw", "n", "ne", "e", "se", "s", "sw", "w"];
const POS: Record<Handle, [number, number]> = { nw: [0, 0], n: [0.5, 0], ne: [1, 0], e: [1, 0.5], se: [1, 1], s: [0.5, 1], sw: [0, 1], w: [0, 0.5] };
type Box = { x: number; y: number; w: number; h: number };
type Run = (id: string, fn: () => void | Promise<void>, enabled?: () => boolean, checked?: () => boolean) => void;
type Sel = { from: number; to: number };

export const TIMELINE_COMMANDS = [
  { id: "insert.timeline", label: "Zeitachse", icon: "date" },
  { id: "pivot.timeline", label: "Zeitachse einfügen", icon: "date" },
  { id: "timeline.caption", label: "Zeitachsenbeschriftung" },
  { id: "timeline.connections", label: "Berichtsverbindungen...", icon: "table" },
  { id: "timeline.styles", label: "Schnellformatvorlagen", icon: "styles" },
  { id: "timeline.forward", label: "Ebene nach vorne", icon: "bring-forward" },
  { id: "timeline.toFront", label: "In den Vordergrund", icon: "bring-forward" },
  { id: "timeline.backward", label: "Ebene nach hinten", icon: "send-backward" },
  { id: "timeline.toBack", label: "In den Hintergrund", icon: "send-backward" },
  { id: "timeline.selectionPane", label: "Auswahlbereich...", icon: "select-objects" },
  { id: "timeline.align", label: "Ausrichten", icon: "align-left" },
  { id: "timeline.group", label: "Gruppieren", icon: "group" },
  { id: "timeline.rotate", label: "Drehen", icon: "rotate" },
  { id: "timeline.height", label: "Formenhöhe" },
  { id: "timeline.width", label: "Formenbreite" },
  { id: "timeline.sizeDialog", label: "Größe und Eigenschaften..." },
  { id: "timeline.showHeader", label: "Kopfzeile" },
  { id: "timeline.showSelectionLabel", label: "Auswahlbeschriftung" },
  { id: "timeline.showScrollbar", label: "Scrollleiste" },
  { id: "timeline.showTimeLevel", label: "Zeitebene" },
  { id: "timeline.clear", label: "Filter löschen", icon: "filter-clear", shortcut: "Alt+C" },
  { id: "timeline.delete", label: "Zeitachse entfernen", icon: "delete" },
];

const cm = (px: number) => `${(px / PX_PER_CM).toLocaleString("de-DE", { maximumFractionDigits: 2 })} cm`;
const parseCm = (v: string): number | null => {
  const x = Number(v.replace(/cm/i, "").trim().replace(",", "."));
  return Number.isFinite(x) && x > 0 ? x * PX_PER_CM : null;
};

/** The parts' heights of a timeline (what is shown). */
export function timelineLayout(t: Pick<Timeline, "showHeader" | "showSelectionLabel" | "showTimeLevel" | "showHorizontalScrollbar">) {
  const head = t.showHeader === false ? 0 : HEAD;
  const row = t.showSelectionLabel === false && t.showTimeLevel === false ? 0 : ROW;
  const scroll = t.showHorizontalScrollbar === false ? 0 : SCROLL;
  return { head, row, area: head + row, scroll, need: head + row + AREA + scroll };
}

export function timelinesFeature(app: CalcApp) {
  const layer = h("div", { class: "timeline-layer" });
  let selected: Timeline | null = null;
  let shownList: Timeline[] | undefined;
  const els: HTMLElement[] = [];
  let live: { t: Timeline; box: Box } | null = null;
  /** A selection being dragged (shown before it filters). */
  let preview: { t: Timeline; sel: Sel } | null = null;
  /** Scrolling while it is dragged or by the wheel (first visible period), kept per timeline until saved. */
  const z = () => app.sheet.view.zoom / 100;
  const caches = () => app.wb.timelineCaches ?? [];
  const cacheOf = (t: Timeline) => caches().find((c) => c.name === t.cache);
  const allTimelines = () => app.wb.sheets.flatMap((s) => s.timelines ?? []);

  function css(): void {
    if (!document.getElementById("raabe-timeline-css")) document.head.append(h("style", { id: "raabe-timeline-css" }, CSS));
  }

  // ----- source and state ----------------------------------------------------------------------

  function pivotsOf(c: TimelineCache): { sheet: Sheet; p: PivotTable }[] {
    return c.pivots.flatMap((r) => {
      const sheet = app.wb.sheets.find((s) => s.name === r.sheet);
      const p = sheet?.pivots?.find((x) => x.name === r.name);
      return sheet && p ? [{ sheet, p }] : [];
    });
  }
  function fieldOf(c: TimelineCache, p: PivotTable): number {
    const names = (app.pivots.data(p)[0] ?? []).map((v) => (v === null ? "" : String(v)));
    return names.findIndex((n) => n.toLowerCase() === c.sourceName.toLowerCase());
  }
  /** The field's dates in the first PivotTable's source. */
  function datesOf(c: TimelineCache): number[] {
    const first = pivotsOf(c)[0];
    if (!first) return [];
    const f = fieldOf(c, first.p);
    return f < 0 ? [] : app.pivots.data(first.p).slice(1).map((r) => r[f]).filter((v): v is number => typeof v === "number");
  }
  /** The selection (null: all periods), from the first PivotTable's date filter. */
  function selectionOf(c: TimelineCache): Sel | null {
    const first = pivotsOf(c)[0];
    if (!first) return null;
    return pivotDateFilter(first.p, fieldOf(c, first.p));
  }
  /** The date fields of a PivotTable's source: numbers formatted as dates. */
  function dateFields(p: PivotTable): string[] {
    const src = app.wb.sheets.find((s) => s.name === p.source.sheet);
    if (!src) return [];
    const out: string[] = [];
    for (let c = p.source.c0; c <= p.source.c1; c++) {
      let dates = 0, other = 0;
      for (let r = p.source.r0 + 1; r <= p.source.r1; r++) {
        const cell = src.get(r, c);
        if (!cell || cell.v === null || cell.v === "") continue;
        if (typeof cell.v === "number" && isDateFormat(app.wb.styles.get(cell.s).numFmt)) dates++;
        else other++;
      }
      if (dates && !other) out.push(String(src.value(p.source.r0, c) ?? ""));
    }
    return out;
  }

  // ----- filtering --------------------------------------------------------------------------------

  /** Applies a selection (null: all) to the cache's PivotTables, as one undo step; the timeline keeps the level it was made at. */
  function select(t: Timeline, sel: Sel | null, label = "Zeitachse"): void {
    const c = cacheOf(t);
    if (!c) return;
    const ps = pivotsOf(c);
    if (!ps.length) return;
    const sheet = app.sheet;
    app.change(label, () => {
      for (const { sheet: s, p } of ps) {
        app.history.whole(s);
        const f = fieldOf(c, p);
        if (f < 0) continue;
        p.dateFilters = pivotDateFiltersAfter(p, f, sel);
        app.pivots.write(s, p);
      }
      app.history.whole(sheet);
      const next = { ...structuredClone(t), selectionLevel: sel ? t.level : t.selectionLevel };
      if (next.selectionLevel === next.level) delete next.selectionLevel;
      sheet.timelines = (sheet.timelines ?? []).map((x) => (x === t ? next : x));
      if (selected === t) selected = next;
    });
    refresh();
  }

  // ----- places ---------------------------------------------------------------------------------

  function anchorAt(x: number, y: number): Timeline["from"] {
    const hit = app.grid.hitCell(Math.max(0, x), Math.max(0, y));
    const c = app.grid.rectOnCanvas({ r0: hit.row, c0: hit.col, r1: hit.row, c1: hit.col })!;
    return { col: hit.col, row: hit.row, x: Math.max(0, (x - c.x) / z()), y: Math.max(0, (y - c.y) / z()) };
  }
  function boxOf(t: Timeline): Box | null {
    const a = app.grid.rectOnCanvas({ r0: t.from.row, c0: t.from.col, r1: t.from.row, c1: t.from.col });
    const b = app.grid.rectOnCanvas({ r0: t.to.row, c0: t.to.col, r1: t.to.row, c1: t.to.col });
    if (!a || !b) return null;
    const x = a.x + t.from.x * z(), y = a.y + t.from.y * z();
    return { x, y, w: b.x + t.to.x * z() - x, h: b.y + t.to.y * z() - y };
  }
  function placed(t: Timeline, b: Box): Timeline {
    return { ...t, from: anchorAt(b.x, b.y), to: anchorAt(b.x + b.w, b.y + b.h), size: { w: b.w / z(), h: b.h / z() } };
  }
  function sizeOf(t: Timeline): { w: number; h: number } {
    const b = boxOf(t);
    return b ? { w: b.w / z(), h: b.h / z() } : t.size ?? { w: TIMELINE_W, h: TIMELINE_H };
  }

  // ----- changes --------------------------------------------------------------------------------

  function edit(label: string, t: Timeline | null, fn: (t: Timeline) => Timeline): void {
    if (!t) return;
    const sheet = app.sheet;
    let next: Timeline | null = null;
    app.change(label, () => {
      app.history.whole(sheet);
      next = fn(structuredClone(t));
      sheet.timelines = (sheet.timelines ?? []).map((x) => (x === t ? next! : x));
    });
    if (selected === t) selected = next;
    refresh();
  }

  /** Anzeigen: an option on or off; the frame grows or shrinks by that part's height (as Excel's does). */
  function toggle(t: Timeline | null, key: "showHeader" | "showSelectionLabel" | "showTimeLevel" | "showHorizontalScrollbar", label: string): void {
    if (!t) return;
    const b = boxOf(t);
    edit(label, t, (x) => {
      const before = timelineLayout(x).need;
      if (x[key] === false) delete x[key]; else x[key] = false;
      const dh = (timelineLayout(x).need - before) * z();
      return b ? placed(x, { ...b, h: Math.max(20, b.h + dh) }) : x;
    });
  }

  function add(p: PivotTable, fields: string[]): void {
    if (!fields.length) return;
    const sheet = app.sheet;
    const view = app.grid.el.getBoundingClientRect();
    const v = sheet.view;
    const tl = app.grid.rectOnCanvas({ r0: v.top, c0: v.left, r1: v.top, c1: v.left }) ?? { x: 0, y: 0, w: 0, h: 0 };
    const cx = tl.x + (view.width - 60) / 2 - (TIMELINE_W * z()) / 2, cy = tl.y + (view.height - 60) / 2 - (TIMELINE_H * z()) / 2;
    const pivotSheet = app.wb.sheets.find((s) => s.pivots?.includes(p))?.name ?? sheet.name;
    let made: Timeline | null = null;
    app.change("Zeitachse einfügen", () => {
      app.history.whole(sheet);
      const list = [...(sheet.timelines ?? [])];
      const cs = [...caches()];
      fields.forEach((field, k) => {
        const name = nextTimelineCacheName([...cs.map((c) => c.name), ...app.wb.names.map((n) => n.name)], field);
        cs.push({ name, sourceName: field, pivots: [{ sheet: pivotSheet, name: p.name }] });
        app.wb.names.push({ name, ref: "#N/A" });
        const taken = [...app.wb.sheets.flatMap((s) => [...(s.slicers ?? []), ...(s.timelines ?? [])]), ...list].map((x) => x.name);
        const tname = nextSlicerName(taken, field);
        const off = k * 26 * z();
        const t0: Timeline = { name: tname, cache: name, caption: field, level: 2, from: { col: 0, row: 0, x: 0, y: 0 }, to: { col: 0, row: 0, x: 0, y: 0 } };
        const t = placed(t0, { x: Math.max(tl.x, cx + off), y: Math.max(tl.y, cy + off), w: TIMELINE_W * z(), h: TIMELINE_H * z() });
        list.push(t);
        made = t;
      });
      app.wb.timelineCaches = cs;
      sheet.timelines = list;
    });
    selected = made;
    refresh();
    app.timelineSelectionChanged();
  }

  function remove(t: Timeline | null): void {
    if (!t) return;
    const sheet = app.sheet;
    app.change("Zeitachse entfernen", () => {
      app.history.whole(sheet);
      sheet.timelines = (sheet.timelines ?? []).filter((x) => x !== t);
      if (!sheet.timelines.length) sheet.timelines = undefined;
      // A cache no other timeline uses goes, with its name; the PivotTable's filter stays (as Excel).
      if (!allTimelines().some((x) => x.cache === t.cache)) {
        app.wb.timelineCaches = caches().filter((c) => c.name !== t.cache);
        if (!app.wb.timelineCaches.length) app.wb.timelineCaches = undefined;
        app.wb.names = app.wb.names.filter((n) => n.name !== t.cache);
      }
    });
    if (selected === t) selected = null;
    refresh();
    app.timelineSelectionChanged();
  }

  // ----- dialogs ----------------------------------------------------------------------------------

  /** Zeitachsen einfügen: one box per date field of the PivotTable. */
  async function insertDialog(): Promise<void> {
    if (app.grid.isEditing) app.grid.commitEdit("none");
    const p = app.pivots.current();
    if (!p) {
      await messageBox("RaabeCalc", "Eine Zeitachse filtert eine PivotTable nach einem Datumsfeld. Wählen Sie zuerst eine Zelle in einer PivotTable aus.");
      return;
    }
    const names = dateFields(p);
    if (!names.length) {
      await messageBox("RaabeCalc", "Diese PivotTable hat kein Feld, das nur Datumswerte enthält. Eine Zeitachse braucht ein solches Feld.");
      return;
    }
    css();
    const boxes = names.map((n) => h("input", { type: "checkbox", "aria-label": n }) as HTMLInputElement);
    const body = h("div", { class: "timeline-dlg", style: "display:flex;flex-direction:column;gap:4px;min-width:220px;max-height:340px;overflow:auto" },
      ...names.map((n, i) => h("label", { style: "display:flex;align-items:center;gap:6px" }, boxes[i], n)));
    const ok = await showDialog({ title: "Zeitachsen einfügen", body, buttons: [{ label: "OK", value: true, primary: true }, { label: "Abbrechen", value: false }], cancel: false });
    if (!ok) return;
    add(p, names.filter((_, i) => boxes[i].checked));
  }

  async function connections(t: Timeline | null = selected): Promise<void> {
    const c = t ? cacheOf(t) : undefined;
    if (!c) return;
    const first = pivotsOf(c)[0];
    if (!first) return;
    const same = (p: PivotTable) => JSON.stringify(p.source) === JSON.stringify(first.p.source);
    const all = app.wb.sheets.flatMap((sheet) => (sheet.pivots ?? []).filter(same).map((p) => ({ sheet, p })));
    const boxes = all.map((x) => h("input", { type: "checkbox", ...(c.pivots.some((r) => r.sheet === x.sheet.name && r.name === x.p.name) ? { checked: true } : {}) }) as HTMLInputElement);
    const body = h("div", { style: "display:grid;grid-template-columns:auto 1fr 1fr;gap:4px 10px;min-width:360px" },
      h("span"), h("b", null, "Name"), h("b", null, "Blatt"),
      ...all.flatMap((x, i) => [boxes[i], h("span", null, x.p.name), h("span", null, x.sheet.name)]));
    const ok = await showDialog({ title: `Berichtsverbindungen (${t!.name})`, body, buttons: [{ label: "OK", value: true, primary: true }, { label: "Abbrechen", value: false }], cancel: false });
    if (!ok) return;
    const chosen = all.filter((_, i) => boxes[i].checked);
    if (!chosen.length) return;
    const sel = selectionOf(c);
    app.change("Berichtsverbindungen", () => {
      // A newly connected PivotTable takes the timeline's filter; a disconnected one keeps its own.
      for (const { sheet, p } of chosen) {
        if (c.pivots.some((r) => r.sheet === sheet.name && r.name === p.name)) continue;
        app.history.whole(sheet);
        const f = fieldOf(c, p);
        if (f >= 0) { p.dateFilters = pivotDateFiltersAfter(p, f, sel); app.pivots.write(sheet, p); }
      }
      app.wb.timelineCaches = caches().map((x) => (x === c ? { ...x, pivots: chosen.map((y) => ({ sheet: y.sheet.name, name: y.p.name })) } : x));
    });
    refresh();
  }

  async function sizeDialog(): Promise<void> {
    const t = selected;
    if (!t) return;
    const sz = sizeOf(t);
    const hi = h("input", { type: "text", value: cm(sz.h), style: "width:90px" }) as HTMLInputElement;
    const wi = h("input", { type: "text", value: cm(sz.w), style: "width:90px" }) as HTMLInputElement;
    const body = h("div", { style: "display:flex;flex-direction:column;gap:8px" }, h("label", null, "Höhe: ", hi), h("label", null, "Breite: ", wi));
    const ok = await showDialog({ title: "Zeitachse formatieren", body, buttons: [{ label: "OK", value: true, primary: true }, { label: "Abbrechen", value: false }], cancel: false });
    if (!ok) return;
    const H = parseCm(hi.value), W = parseCm(wi.value);
    if (H || W) resize(t, W ?? sz.w, H ?? sz.h);
  }

  function resize(t: Timeline, w: number, hh: number): void {
    const b = boxOf(t);
    if (!b) return;
    edit("Größe", t, (x) => placed(x, { x: b.x, y: b.y, w: Math.max(20, w) * z(), h: Math.max(20, hh) * z() }));
  }

  // ----- drawing ----------------------------------------------------------------------------------

  const themeColor = (c: TsColor | undefined, fallback: string): string => {
    if (!c) return fallback;
    if (typeof c === "string") return c;
    return tinted(app.wb.theme[c.t] ?? OFFICE_2013[c.t] ?? "#000000", c.tint ?? 0);
  };
  const styleOf = (t: Timeline) => TIMELINE_STYLES.find((x) => x.name === (t.style ?? "TimeSlicerStyleLight1")) ?? TIMELINE_STYLES[0];
  const el = (t: Timeline, k: string): TimelineElement => styleOf(t).elements[k] ?? {};
  const font = () => `"${app.wb.themeFonts?.minor ?? "Calibri"}", Calibri, Carlito, sans-serif`;
  const fill = (e: TimelineElement, fb: string) => (e.fill2 ? `linear-gradient(${themeColor(e.fill, fb)}, ${themeColor(e.fill2, fb)})` : themeColor(e.fill, fb));

  /** What the timeline shows: bounds, the first visible period, how many fit, the selection. */
  function view(t: Timeline, w: number) {
    const c = cacheOf(t);
    const bounds = c ? timelineBounds(datesOf(c)) : null;
    const sel = preview?.t === t ? preview.sel : c ? selectionOf(c) : null;
    const pitch = PITCH[t.level];
    const fit = Math.max(1, Math.floor((w - 2 * MARGIN) / pitch));
    const total = bounds ? periodsBetween(bounds.start, bounds.end - 1, t.level) + 1 : 0;
    // Unscrolled: the selection's start, or (as a new timeline in Excel) the latest periods with data.
    const dates = c ? datesOf(c) : [];
    const latest = dates.length ? addPeriods(Math.max(...dates), t.level, -(fit - 1)) : undefined;
    let first = bounds ? periodStart(t.scrollPosition ?? sel?.from ?? latest ?? bounds.start, t.level) : 0;
    if (bounds) {
      // Never past the end: the last periods fill the strip.
      const maxStart = addPeriods(bounds.start, t.level, Math.max(0, total - fit));
      if (first > maxStart) first = maxStart;
      if (first < bounds.start) first = bounds.start;
    }
    return { c, bounds, sel, pitch, fit, total, first };
  }

  function render(t: Timeline, root: HTMLElement, size: { w: number; h: number }): void {
    const whole = el(t, "wholeTable"), head = el(t, "headerRow");
    const border = themeColor(whole.borders?.left, "#156082");
    Object.assign(root.style, { width: `${size.w}px`, height: `${size.h}px`, border: `1px solid ${border}`, font: `${11 * 96 / 72}px ${font()}`, color: themeColor(whole.font, "#000") });
    root.replaceChildren();
    const L = timelineLayout(t);
    const v = view(t, size.w);
    if (L.head) {
      const clear = h("span", { class: "tl-ic", title: "Filter löschen (Alt+C)", "data-act": "clear" });
      const on = !!v.sel;
      clear.innerHTML = `<svg width="16" height="16" viewBox="0 0 16 16"><path d="M1.5 2h11l-4.2 5v5.5L6 14V7z" fill="none" stroke="${on ? "#555" : "#b5b5b5"}" stroke-width="1.1"/><path d="M10 10l4.5 4.5M14.5 10L10 14.5" stroke="${on ? "#d24726" : "#b5b5b5"}" stroke-width="1.4"/></svg>`;
      root.append(h("div", { class: "tl-head", style: `height:${HEAD}px;border-bottom:1px solid ${border};font-weight:${head.bold ? 700 : 400};color:${themeColor(head.font, "#000")}` },
        h("span", { class: "tl-cap" }, t.caption), clear));
    }
    if (L.row) {
      const se = el(t, "selectionLabel"), le = el(t, "timeLevel");
      const row = h("div", { class: "tl-row", style: `top:${L.head}px;height:${ROW}px` },
        h("span", { class: "tl-sel", "data-part": "label", style: `visibility:${t.showSelectionLabel === false ? "hidden" : "visible"};font-size:${(se.size ?? 10) * 96 / 72}px;color:${themeColor(se.font, "#104861")}` },
          selectionLabel(v.sel, (t.selectionLevel ?? t.level) as TimeLevel)));
      if (t.showTimeLevel !== false) {
        const lv = h("span", { class: "tl-lvl", "data-act": "level", style: `font-size:${(le.size ?? 9) * 96 / 72}px;color:${themeColor(le.font, "#808080")}` }, `${TIME_LEVELS[t.level].toUpperCase()} ▾`);
        row.append(lv);
      }
      root.append(row);
    }
    const area = h("div", { class: "tl-area", "data-part": "area", style: `top:${L.area}px;height:${AREA}px` });
    const p1 = el(t, "periodLabel1"), p2 = el(t, "periodLabel2"), on = el(t, "selectedTimeBlock"), off = el(t, "unselectedTimeBlock");
    if (v.bounds) {
      const visible = Math.min(v.fit + 1, v.total);
      let upperShown = false;
      for (let i = 0; i < visible; i++) {
        const start = addPeriods(v.first, t.level, i);
        if (start >= v.bounds.end) break;
        const end = periodEnd(start, t.level);
        const x = MARGIN + i * v.pitch;
        const lab = periodLabels(start, t.level);
        const upper = lab.upper ?? (!upperShown && t.level !== 0 ? firstUpperLabel(start, t.level) : undefined);
        if (upper) {
          upperShown = true;
          area.append(h("span", { class: "tl-lab", style: `left:${x + 2}px;top:2px;transform:none;font-size:${(p1.size ?? 9) * 96 / 72}px;color:${themeColor(p1.font, "#808080")}` }, upper));
        }
        if (t.level === 0) area.append(h("span", { class: "tl-lab", style: `left:${x + 2}px;top:22px;transform:none;font-size:${(p1.size ?? 9) * 96 / 72}px;color:${themeColor(p1.font, "#808080")}` }, lab.lower));
        else area.append(h("span", { class: "tl-lab", style: `left:${x + v.pitch / 2}px;top:22px;font-size:${(p2.size ?? 9) * 96 / 72}px;color:${themeColor(p2.font, "#808080")}` }, lab.lower));
        const isOn = !v.sel || (start <= v.sel.to && end >= v.sel.from);
        const nextOn = !v.sel || (addPeriods(start, t.level, 1) <= v.sel.to && periodEnd(addPeriods(start, t.level, 1), t.level) >= v.sel.from);
        const blk = h("div", { class: "tl-blk", "data-period": String(start), "data-on": isOn ? "1" : "0",
          style: `left:${x}px;top:${BLOCK_TOP}px;width:${isOn && nextOn ? v.pitch : v.pitch - 1}px;height:${BLOCK_H}px;background:${isOn ? fill(on, "#156082") : fill(off, "#d9d9d9")}` });
        area.append(blk);
      }
      // The selection's edges (dragged to widen or narrow it).
      if (v.sel) {
        const a = periodsBetween(v.first, v.sel.from, t.level), b = periodsBetween(v.first, v.sel.to, t.level) + 1;
        for (const [k, i] of [["from", a], ["to", b]] as const) {
          if (i < 0 || i > v.fit + 1) continue;
          area.append(h("div", { class: "tl-edge", "data-edge": k, style: `left:${MARGIN + i * v.pitch}px;top:${BLOCK_TOP - 2}px;height:${BLOCK_H + 4}px` }));
        }
      }
    }
    root.append(area);
    if (L.scroll) {
      const top = L.area + AREA + 2;
      const sc = h("div", { class: "tl-scroll", "data-part": "scroll", style: `top:${top}px` });
      const arrow = (dir: -1 | 1) => {
        const a = h("div", { class: "tl-arrow", "data-act": dir < 0 ? "left" : "right", style: dir < 0 ? "left:0" : "right:0" });
        a.innerHTML = `<svg width="8" height="8" viewBox="0 0 8 8"><path d="${dir < 0 ? "M5.5 1L2 4l3.5 3" : "M2.5 1L6 4 2.5 7"}" fill="#333"/></svg>`;
        return a;
      };
      const track = h("div", { class: "tl-track", "data-part": "track" });
      if (v.bounds && v.total) {
        const tw = Math.max(1, size.w - 2 * MARGIN - 36);
        const firstIdx = periodsBetween(v.bounds.start, v.first, t.level);
        const thumbW = Math.max(12, (Math.min(v.fit, v.total) / v.total) * tw);
        const thumbX = v.total > v.fit ? (firstIdx / (v.total - v.fit)) * (tw - thumbW) : 0;
        if (v.sel) {
          const sa = Math.max(0, periodsBetween(v.bounds.start, v.sel.from, t.level)) / v.total * tw;
          const sb = Math.min(v.total, periodsBetween(v.bounds.start, v.sel.to, t.level) + 1) / v.total * tw;
          track.append(h("div", { class: "tl-span", style: `left:${sa}px;width:${Math.max(2, sb - sa)}px;background:${themeColor(el(t, "selectedTimeBlockSpace").fill, "#cfe6f3")};opacity:.45` }));
        }
        track.append(h("div", { class: "tl-thumb", "data-act": "thumb", style: `left:${thumbX}px;width:${thumbW}px` }));
      }
      sc.append(arrow(-1), track, arrow(1));
      root.append(sc);
    }
    for (const k of HANDLES) root.append(h("span", { class: `tl-h ${k}`, "data-h": k, style: `left:${POS[k][0] * 100}%;top:${POS[k][1] * 100}%` }));
    root.dataset.first = String(v.first);
    root.dataset.sel = v.sel ? `${v.sel.from}:${v.sel.to}` : "";
  }

  function place(): void {
    const list = app.sheet.timelines;
    if (!list?.length) {
      if (layer.isConnected) layer.remove();
      shownList = undefined;
      if (selected) { selected = null; app.timelineSelectionChanged(); }
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
      if (!selected) queueMicrotask(() => app.timelineSelectionChanged());
    }
    if (shownList !== list) {
      layer.replaceChildren();
      els.length = 0;
      for (const t of list) {
        const e = h("div", { class: "timeline", tabindex: "-1" });
        e.addEventListener("pointerdown", (ev) => down(ev, list[els.indexOf(e)] ?? t, e));
        e.addEventListener("contextmenu", (ev) => { ev.preventDefault(); contextMenu(list[els.indexOf(e)] ?? t, ev); });
        e.addEventListener("keydown", (ev) => key(ev, list[els.indexOf(e)] ?? t));
        e.addEventListener("wheel", (ev) => { ev.preventDefault(); ev.stopPropagation(); scrollBy(list[els.indexOf(e)] ?? t, ev.deltaY > 0 || ev.deltaX > 0 ? 1 : -1); }, { passive: false });
        e.dataset.name = t.name;
        els.push(e);
        layer.append(e);
      }
      shownList = list;
    }
    const v = app.sheet.view;
    const tl = app.grid.rectOnCanvas({ r0: v.top, c0: v.left, r1: v.top, c1: v.left });
    if (!tl) return;
    Object.assign(layer.style, { left: `${tl.x}px`, top: `${tl.y}px` });
    list.forEach((t, i) => {
      const e = els[i];
      const b = live?.t === t ? live.box : boxOf(t);
      if (!b || b.w <= 0 || b.h <= 0) { e.style.display = "none"; return; }
      const size = { w: b.w / z(), h: b.h / z() };
      const c = cacheOf(t);
      const key = JSON.stringify([t, size.w.toFixed(1), size.h.toFixed(1), c ? [selectionOf(c), datesOf(c).length] : null, preview?.t === t ? preview.sel : null, app.wb.themeFonts?.minor, app.wb.theme.accent1]);
      if (e.dataset.key !== key) { render(t, e, size); e.dataset.key = key; }
      e.dataset.name = t.name;
      e.classList.toggle("on", t === selected);
      Object.assign(e.style, { display: "", left: `${b.x - tl.x}px`, top: `${b.y - tl.y}px`, transform: `scale(${z()})` });
    });
  }

  function refresh(): void {
    shownList = undefined;
    app.grid.contentChanged();
  }

  function choose(t: Timeline | null): void {
    if (selected === t) return;
    selected = t;
    app.timelineSelectionChanged();
    app.grid.contentChanged();
  }

  // ----- mouse and keys ---------------------------------------------------------------------------

  /** Scrolling moves the first visible period (not an undo step of its own: it is kept with the next change). */
  function scrollBy(t: Timeline, n: number): void {
    const w = sizeOf(t).w;
    const v = view(t, w);
    if (!v.bounds) return;
    const next = addPeriods(v.first, t.level, n);
    const sheet = app.sheet;
    const nt = { ...t, scrollPosition: Math.max(v.bounds.start, next) };
    sheet.timelines = (sheet.timelines ?? []).map((x) => (x === t ? nt : x));
    if (selected === t) selected = nt;
    refresh();
  }

  function setLevel(t: Timeline, level: TimeLevel): void {
    edit("Zeitebene", t, (x) => {
      x.level = level;
      // The view stays near the selection's start (or what was first shown).
      const c = cacheOf(x);
      const sel = c ? selectionOf(c) : null;
      const keep = sel?.from ?? x.scrollPosition;
      if (keep === undefined) delete x.scrollPosition; else x.scrollPosition = periodStart(keep, level);
      if ((x.selectionLevel ?? t.level) === level) delete x.selectionLevel; else x.selectionLevel ??= t.level;
      return x;
    });
  }

  function down(ev: PointerEvent, t: Timeline, e: HTMLElement): void {
    if (ev.button !== 0) { choose(t); return; }
    ev.stopPropagation();
    const target = ev.target as HTMLElement;
    const act = target.closest<HTMLElement>("[data-act]")?.dataset.act;
    const handle = target.closest<HTMLElement>(".tl-h")?.dataset.h as Handle | undefined;
    const edge = target.closest<HTMLElement>(".tl-edge")?.dataset.edge as "from" | "to" | undefined;
    const part = target.closest<HTMLElement>("[data-part]")?.dataset.part;
    choose(t);
    e.focus({ preventScroll: true });
    const rect = e.getBoundingClientRect();
    const localX = (x: number) => (x - rect.left) / z();
    const w = sizeOf(t).w;
    const v = view(t, w);
    const periodAt = (clientX: number) => {
      const i = Math.floor((localX(clientX) - MARGIN) / v.pitch);
      const s = addPeriods(v.first, t.level, Math.max(0, Math.min(v.fit, i)));
      return v.bounds ? Math.max(v.bounds.start, Math.min(addPeriods(v.bounds.end - 1, t.level, 0), s)) : s;
    };
    if (act === "clear") { select(t, null, "Filter löschen"); return; }
    if (act === "level") {
      const r = target.getBoundingClientRect();
      showMenu(app.commands, TIME_LEVELS.map((label, i) => ({ label, checked: t.level === i, run: () => setLevel(t, i as TimeLevel) })), new DOMRect(r.left, r.bottom, 0, 0));
      return;
    }
    if (act === "left" || act === "right") { scrollBy(t, act === "left" ? -1 : 1); return; }
    if (part === "track" && act !== "thumb") {
      const tr = target.getBoundingClientRect();
      const thumb = e.querySelector<HTMLElement>(".tl-thumb")?.getBoundingClientRect();
      scrollBy(t, thumb && ev.clientX < thumb.left ? -v.fit : v.fit);
      void tr;
      return;
    }
    if (act === "thumb" && v.bounds) {
      const x0 = ev.clientX, first0 = v.first;
      const tw = Math.max(1, w - 2 * MARGIN - 36);
      const per = v.total > v.fit ? tw / v.total : 0;
      const move = (m: PointerEvent) => {
        if (!per) return;
        const n = Math.round((m.clientX - x0) / z() / per);
        const sheet = app.sheet;
        const cur = sheet.timelines?.find((x) => x.name === t.name);
        if (!cur) return;
        const nt = { ...cur, scrollPosition: Math.max(v.bounds!.start, addPeriods(first0, t.level, n)) };
        sheet.timelines = sheet.timelines!.map((x) => (x === cur ? nt : x));
        if (selected === cur) selected = nt;
        refresh();
      };
      const up = () => { removeEventListener("pointermove", move); removeEventListener("pointerup", up); };
      addEventListener("pointermove", move);
      addEventListener("pointerup", up);
      return;
    }
    if (part === "area" || edge) {
      const c = cacheOf(t);
      if (!c || !v.bounds) return;
      const cur = selectionOf(c);
      const p0 = periodAt(ev.clientX);
      let mode: "new" | "from" | "to" | "move" = "new";
      let anchor = p0;
      if (edge && cur) { mode = edge; anchor = edge === "from" ? periodStart(cur.to, t.level) : periodStart(cur.from, t.level); }
      else if (ev.shiftKey && cur) { mode = "new"; anchor = periodStart(cur.from, t.level); }
      else if (cur && p0 >= periodStart(cur.from, t.level) && p0 <= cur.to) mode = "move";
      const span = cur ? periodsBetween(cur.from, cur.to, t.level) : 0;
      const range = (a: number, b: number): Sel => ({ from: Math.min(a, b), to: periodEnd(Math.max(a, b), t.level) });
      const at = (clientX: number): Sel => {
        const p = periodAt(clientX);
        if (mode === "move" && cur) {
          const d = periodsBetween(p0, p, t.level);
          let from = addPeriods(cur.from, t.level, d);
          const last = addPeriods(v.bounds!.end - 1, t.level, 0);
          if (from < v.bounds!.start) from = v.bounds!.start;
          if (addPeriods(from, t.level, span) > last) from = addPeriods(last, t.level, -span);
          return { from, to: periodEnd(addPeriods(from, t.level, span), t.level) };
        }
        return range(anchor, p);
      };
      let moved = false;
      preview = { t, sel: at(ev.clientX) };
      const startSel = preview.sel;
      refresh();
      const move = (m: PointerEvent) => { const s = at(m.clientX); if (s.from !== preview?.sel.from || s.to !== preview?.sel.to) moved = true; preview = { t, sel: s }; refresh(); };
      const up = (u: PointerEvent) => {
        removeEventListener("pointermove", move);
        removeEventListener("pointerup", up);
        const s = preview?.sel ?? at(u.clientX);
        preview = null;
        // A click inside the selection without moving takes that one period (as a click elsewhere does).
        const finalSel = mode === "move" && !moved ? range(p0, p0) : s;
        void startSel;
        const cur2 = selectionOf(c);
        if (cur2 && cur2.from === finalSel.from && cur2.to === finalSel.to && t.level === (t.selectionLevel ?? t.level)) { refresh(); return; }
        const tt = app.sheet.timelines?.find((x) => x.name === t.name) ?? t;
        select(tt, finalSel);
      };
      addEventListener("pointermove", move);
      addEventListener("pointerup", up);
      return;
    }
    // Header, border or handle: move or size.
    const start = boxOf(t);
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
      live = { t, box: b };
      app.grid.contentChanged();
    };
    const up = () => {
      removeEventListener("pointermove", move);
      removeEventListener("pointerup", up);
      const done = live;
      live = null;
      if (moved && done) edit(handle ? "Größe" : "Verschieben", t, (x) => placed(x, done.box));
      else app.grid.contentChanged();
    };
    addEventListener("pointermove", move);
    addEventListener("pointerup", up);
  }

  function key(ev: KeyboardEvent, t: Timeline): void {
    if (ev.key === "Delete") { ev.preventDefault(); remove(t); return; }
    if (ev.altKey && (ev.key === "c" || ev.key === "C")) { ev.preventDefault(); select(t, null, "Filter löschen"); return; }
    if (ev.key === "Escape") { ev.preventDefault(); choose(null); app.grid.focus(); return; }
    const d: Record<string, [number, number]> = { ArrowLeft: [-1, 0], ArrowRight: [1, 0], ArrowUp: [0, -1], ArrowDown: [0, 1] };
    if (d[ev.key]) {
      ev.preventDefault();
      const b = boxOf(t);
      if (b) edit("Verschieben", t, (x) => placed(x, { ...b, x: b.x + d[ev.key][0] * z(), y: b.y + d[ev.key][1] * z() }));
    }
  }

  function contextMenu(t: Timeline, ev: MouseEvent): void {
    choose(t);
    const c = cacheOf(t);
    const entries: MenuEntry[] = [
      { label: `Filter aus "${t.caption}" löschen`, icon: "filter-clear", run: () => select(t, null, "Filter löschen"), disabled: !c || !selectionOf(c) },
      "-",
      { label: `"${t.name}" entfernen`, icon: "delete", run: () => remove(t) },
      "-",
      { label: "Berichtsverbindungen...", run: () => void connections(t) },
      { label: "Größe und Eigenschaften...", run: () => void sizeDialog() },
    ];
    showMenu(app.commands, entries, new DOMRect(ev.clientX, ev.clientY, 0, 0));
  }

  addEventListener("pointerdown", (e) => {
    const t = e.target as HTMLElement;
    if (!selected || t.closest?.(".timeline") || t.closest?.(".ribbon-card, .tab, .taskpane, .popup, .rpanel, .menu, .menu-root, .dialog-layer, [role=dialog], [role=menu]")) return;
    choose(null);
  }, true);

  // ----- ribbon ---------------------------------------------------------------------------------

  function stylesPanel(close: () => void): HTMLElement {
    css();
    const box = h("div", { class: "timeline-gallery" });
    for (const [label, re] of [["Hell", /Light/], ["Dunkel", /Dark/]] as const) {
      box.append(h("div", { class: "tg-head" }, label));
      for (const st of TIMELINE_STYLES.filter((x) => re.test(x.name))) {
        const e = (k: string) => st.elements[k] ?? {};
        const tile = h("div", { class: `tg-tile${(selected?.style ?? "TimeSlicerStyleLight1") === st.name ? " on" : ""}`, title: st.local });
        const bd = themeColor(e("wholeTable").borders?.left, "#999");
        tile.innerHTML = `<svg width="52" height="36" viewBox="0 0 52 36"><rect x="0.5" y="0.5" width="51" height="35" fill="#fff" stroke="${bd}"/>`
          + `<rect x="4" y="5" width="22" height="3" fill="${themeColor(e("selectionLabel").font, "#333")}"/>`
          + [0, 1, 2, 3].map((k) => `<rect x="${4 + k * 11}" y="20" width="10" height="6" fill="${themeColor(k === 1 || k === 2 ? e("selectedTimeBlock").fill2 : e("unselectedTimeBlock").fill, "#ccc")}"/>`).join("") + "</svg>";
        tile.addEventListener("click", () => { close(); edit("Zeitachsenformatvorlage", selected, (x) => ({ ...x, style: st.name === "TimeSlicerStyleLight1" ? undefined : st.name })); });
        box.append(tile);
      }
    }
    return box;
  }

  const has = () => !!selected;
  const sizeSpin = (cmd: string, which: "w" | "h", key: string): Control => ({
    type: "spinner", cmd, label: which === "h" ? "Höhe:" : "Breite:", key,
    value: () => (selected ? cm(sizeOf(selected)[which]) : ""),
    commit: (v) => { const t = selected; const x = parseCm(v); if (t && x) { const sz = sizeOf(t); resize(t, which === "w" ? x : sz.w, which === "h" ? x : sz.h); } },
    step: (d) => { const t = selected; if (t) { const sz = sizeOf(t); const step = 0.1 * PX_PER_CM * d; resize(t, which === "w" ? sz.w + step : sz.w, which === "h" ? sz.h + step : sz.h); } },
  });

  let tabDef: Tab | null = null;
  function tab(): Tab {
    const b = (cmd: string, key: string, size: "large" | "medium" | "small" = "small"): Control => ({ type: "button", cmd, key, size });
    const ck = (cmd: string, key: string): Control => ({ type: "check", cmd, key });
    // Keytips as recorded from Excel 365 (probe_excel_timeline_ui.py): tab J L, groups Z 1, Z Q, Z A, Z Z, Z 2.
    return tabDef ??= {
      id: "timeline", label: "Zeitachse", key: "JL", contextual: true, shrink: [],
      groups: [
        { id: "timeline", label: "Zeitachse", key: "Z1", icon: "date", controls: [{ type: "stack", items: [
          { type: "combo", cmd: "timeline.caption", width: 130, key: "T", label: "Zeitachsenbeschriftung:", value: () => selected?.caption ?? "",
            commit: (v) => edit("Zeitachsenbeschriftung", selected, (x) => ({ ...x, caption: v })) },
          b("timeline.connections", "E", "medium")] }] },
        { id: "styles", label: "Zeitachsen-Formatvorlagen", key: "ZQ", icon: "styles", controls: [
          { type: "dropdown", cmd: "timeline.styles", key: "S", size: "large", menu: [], popup: (close) => stylesPanel(close) }] },
        { id: "arrange", label: "Anordnen", key: "ZA", icon: "bring-forward", controls: [
          { type: "split", cmd: "timeline.forward", key: "GE", size: "large", menu: [{ cmd: "timeline.forward" }, { cmd: "timeline.toFront" }] },
          { type: "split", cmd: "timeline.backward", key: "GH", size: "large", menu: [{ cmd: "timeline.backward" }, { cmd: "timeline.toBack" }] },
          b("timeline.selectionPane", "Q", "large"),
          { type: "stack", items: [
            { type: "dropdown", cmd: "timeline.align", key: "DA", size: "small", menu: [{ label: "Am Raster ausrichten", run: () => {}, disabled: true }] },
            { type: "dropdown", cmd: "timeline.group", key: "GR", size: "small", menu: [] },
            { type: "dropdown", cmd: "timeline.rotate", key: "DR", size: "small", menu: [] }] }] },
        { id: "size", label: "Größe", key: "ZZ", icon: "position", launcher: { cmd: "timeline.sizeDialog", key: "ß" }, controls: [{ type: "stack", items: [
          sizeSpin("timeline.height", "h", "H"), sizeSpin("timeline.width", "w", "RR")] }] },
        { id: "show", label: "Anzeigen", key: "Z2", icon: "eye", controls: [
          { type: "stack", items: [ck("timeline.showHeader", "K"), ck("timeline.showSelectionLabel", "A")] },
          { type: "stack", items: [ck("timeline.showScrollbar", "O"), ck("timeline.showTimeLevel", "I")] }] },
      ],
    };
  }

  function reorder(t: Timeline | null, to: "forward" | "front" | "backward" | "back"): void {
    if (!t) return;
    const sheet = app.sheet;
    const list = [...(sheet.timelines ?? [])];
    const i = list.indexOf(t);
    if (i < 0) return;
    list.splice(i, 1);
    const j = to === "front" ? list.length : to === "back" ? 0 : to === "forward" ? Math.min(list.length, i + 1) : Math.max(0, i - 1);
    list.splice(j, 0, t);
    app.change("Anordnen", () => { app.history.whole(sheet); sheet.timelines = list; });
    refresh();
  }

  function implement(run: Run): void {
    app.commands.register(...TIMELINE_COMMANDS);
    const many = () => has() && (app.sheet.timelines?.length ?? 0) > 1;
    run("insert.timeline", () => void insertDialog(), () => !app.grid.isEditing);
    run("pivot.timeline", () => void insertDialog(), () => !!app.pivots.current());
    run("timeline.caption", () => {}, has);
    run("timeline.connections", () => void connections(), has);
    run("timeline.styles", () => {}, has);
    run("timeline.forward", () => reorder(selected, "forward"), many);
    run("timeline.toFront", () => reorder(selected, "front"), many);
    run("timeline.backward", () => reorder(selected, "backward"), many);
    run("timeline.toBack", () => reorder(selected, "back"), many);
    run("timeline.selectionPane", () => void app.commands.run("layout.selectionPane"), has);
    run("timeline.align", () => {}, has);
    run("timeline.height", () => {}, has);
    run("timeline.width", () => {}, has);
    run("timeline.sizeDialog", () => void sizeDialog(), has);
    run("timeline.showHeader", () => toggle(selected, "showHeader", "Kopfzeile"), has, () => selected?.showHeader !== false);
    run("timeline.showSelectionLabel", () => toggle(selected, "showSelectionLabel", "Auswahlbeschriftung"), has, () => selected?.showSelectionLabel !== false);
    run("timeline.showScrollbar", () => toggle(selected, "showHorizontalScrollbar", "Scrollleiste"), has, () => selected?.showHorizontalScrollbar !== false);
    run("timeline.showTimeLevel", () => toggle(selected, "showTimeLevel", "Zeitebene"), has, () => selected?.showTimeLevel !== false);
    run("timeline.clear", () => { if (selected) select(selected, null, "Filter löschen"); }, () => !!selected && !!cacheOf(selected) && !!selectionOf(cacheOf(selected)!));
    run("timeline.delete", () => remove(selected), has);
  }

  return {
    place, tab, implement, add, remove, select, insertDialog, dateFields, setLevel, scrollBy, toggle,
    selected: () => selected,
    deselect: () => choose(null),
    byName: (name: string) => app.sheet.timelines?.find((t) => t.name === name) ?? null,
    choose, cacheOf, sizeOf, resize, edit,
    selection: (t: Timeline) => { const c = cacheOf(t); return c ? selectionOf(c) : null; },
    /** Tests and the harness: where a period's block is (CSS px from the viewport's top left), and the label shown. */
    blockPoint: (name: string, start: number) => {
      const e = els.find((x) => x.dataset.name === name);
      const b = e?.querySelector<HTMLElement>(`.tl-blk[data-period="${start}"]`)?.getBoundingClientRect();
      return b ? { x: b.left + b.width / 2, y: b.top + b.height / 2, left: b.left, right: b.right } : null;
    },
    label: (name: string) => els.find((x) => x.dataset.name === name)?.querySelector(".tl-sel")?.textContent ?? null,
  };
}
