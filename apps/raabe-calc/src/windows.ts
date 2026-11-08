// Ansicht > Fenster and Arbeitsmappenansichten > Benutzerdefinierte Ansichten (D-093), recorded in our own
// invisible Excel (probe_excel_windows.py, results/2026-11-08-excel-windows.json):
// - Neues Fenster: a further window of the same workbook with its own active sheet, selection, scroll and zoom
//   (edits show in all). With several windows each has the caption "Mappe1  -  2" (Excel 365's Window.Caption;
//   older Excel wrote "Mappe1:2"), shown as a document tab; closing one numbers the rest again and the last one
//   left is called "Mappe1" again (recorded). The file holds one workbookView per window and one sheetView per
//   window on every sheet (written by ooxml/views.ts).
// - Alle anordnen: Excel's dialog; the windows are laid out inside our window (Unterteilt, Horizontal, Vertikal,
//   Überlappend), the other windows drawn by grids of their own that a click makes the active one.
// - Ausblenden / Einblenden..., Nebeneinander anzeigen (two windows above each other, as Excel arranges them),
//   Synchrones Scrollen, Fensterposition zurücksetzen, Fenster wechseln (the numbered list of windows).
// - Benutzerdefinierte Ansichten: Anzeigen, Schließen, Hinzufügen..., Löschen; a view keeps per sheet the zoom,
//   the selection and the top-left cell, with "Ausgeblendete Zeilen, Spalten und Filtereinstellungen" the hidden
//   rows and columns (as Excel: names Z_<guid>_.wvu.Rows/.Cols), with "Druckeinstellungen" the page setup.
//   Excel refuses it while the workbook has a table (recorded), so the button is greyed then.

import { h, showDialog, type MenuEntry } from "@raabe/shell";
import { cellName, colName, parseRange, swapWindow, type Sheet, type View, type Workbook } from "@raabe/sheet-engine";
import { Grid, type GridHost } from "./grid/grid";
import type { CalcApp, Doc } from "./app";

type Mode = "tiled" | "horizontal" | "vertical" | "cascade";
const clone = <T>(v: T): T => JSON.parse(JSON.stringify(v)) as T;

const CSS = `
.win-tiles { flex: 1; min-height: 0; display: grid; gap: 4px; background: #c8c8c8; padding: 4px; position: relative; }
.win-tiles.cascade { display: block; }
.win-tile { display: flex; flex-direction: column; min-height: 0; min-width: 0; background: #fff; border: 1px solid #a0a0a0; position: relative; }
.win-tiles.cascade .win-tile { position: absolute; width: 70%; height: 70%; }
.win-tile.on { border-color: #217346; }
.win-cap { height: 22px; line-height: 22px; padding: 0 8px; font-size: 12px; background: #f3f3f3; color: #444; border-bottom: 1px solid #d4d4d4; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; flex: none; }
.win-tile.on .win-cap { background: #e1efe6; color: #1e1e1e; }
.win-body { flex: 1; min-height: 0; position: relative; display: flex; }
.win-body > .gridwrap { flex: 1; }
.win-shield { position: absolute; inset: 0; z-index: 9; cursor: default; }
body.no-window .calc-main > * { visibility: hidden; }
.cv-list { width: 260px; height: 140px; border: 1px solid #ababab; overflow-y: auto; margin: 4px 0 8px; }
.cv-list div { padding: 1px 4px; cursor: default; white-space: nowrap; }
.cv-list div.on { background: #cce4f7; }
.cv-row { display: flex; gap: 12px; align-items: flex-start; }
.cv-btns { display: flex; flex-direction: column; gap: 6px; margin-top: 4px; }
.cv-btns button { min-width: 92px; }
`;

export function windowsFeature(app: CalcApp) {
  let arranged: { mode: Mode; activeOnly: boolean } | null = null;
  let side: { other: string } | null = null;
  let sync = true;
  let tiles: HTMLElement | null = null;
  let passive: { grid: Grid; id: string; doc: Doc; k: number }[] = [];
  let lastScroll: { top: number; left: number; sheet: Sheet } | null = null;

  const css = () => { if (!document.getElementById("raabe-windows-css")) document.head.append(h("style", { id: "raabe-windows-css" }, CSS)); };
  const count = (wb: Workbook) => wb.windows?.length ?? 1;
  const shown = (wb: Workbook) => wb.shownWindow ?? 0;
  const hidden = (wb: Workbook, k: number) => !!wb.windows?.[k]?.hidden;

  /** Window k's view of a sheet and its active sheet, whatever window is shown. */
  function viewOf(wb: Workbook, k: number, sheet: Sheet): View {
    const s = shown(wb);
    if (k === s) return sheet.view;
    return (k === 0 ? sheet.views?.[s - 1] : sheet.views?.[k - 1]) ?? sheet.view;
  }
  function activeOf(wb: Workbook, k: number): number {
    const s = shown(wb);
    if (k === s) return wb.active;
    return (k === 0 ? wb.windows?.[s]?.active : wb.windows?.[k]?.active) ?? wb.active;
  }

  /** Shows window k of a workbook in place (its views into the sheets' `view`). */
  function show(wb: Workbook, k: number): void {
    const s = shown(wb);
    if (s === k) return;
    swapWindow(wb, s);
    swapWindow(wb, k);
    wb.shownWindow = k || undefined;
  }

  const caption = (d: Doc, k: number) => (count(d.wb) > 1 ? `${d.name}  -  ${k + 1}` : d.name);
  const idOf = (d: Doc, k: number) => (count(d.wb) > 1 ? `${d.id}#${k}` : d.id);
  const parse = (id: string): { doc: Doc | undefined; k: number } => {
    const [docId, k] = id.split("#");
    return { doc: app.docList.find((d) => d.id === docId), k: Number(k ?? 0) };
  };

  /** The document tabs: one per visible window. */
  function tabs(d: Doc): { id: string; name: string; dirty: boolean }[] {
    const dirty = d.version !== d.savedVersion;
    const out = [];
    for (let k = 0; k < count(d.wb); k++) if (!hidden(d.wb, k)) out.push({ id: idOf(d, k), name: caption(d, k), dirty });
    return out;
  }

  const activeId = (d: Doc) => idOf(d, shown(d.wb));
  const title = (d: Doc) => caption(d, shown(d.wb));

  /** Makes a window the active one (document tab, Fenster wechseln, a click on an arranged window). */
  function select(id: string): void {
    const { doc, k } = parse(id);
    if (!doc) return;
    if (app.grid.isEditing) app.grid.commitEdit("none");
    show(doc.wb, Math.min(k, count(doc.wb) - 1));
    app.activate(doc);
    if (tiles) layout();
  }

  // ----- Neues Fenster, closing, Ausblenden ----------------------------------------------------

  function newWindow(): void {
    const wb = app.wb;
    if (app.grid.isEditing) app.grid.commitEdit("none");
    if (!wb.windows) wb.windows = [{ active: wb.active }];
    wb.windows.push({ active: wb.active });
    const k = wb.windows.length - 1;
    for (const s of wb.sheets) {
      s.views ??= [];
      // The new window starts as the window it was made from (recorded).
      s.views[k - 1] = clone(s.view);
    }
    show(wb, k);
    app.markChanged();
    app.activate(app.activeDoc);
    if (tiles) layout();
  }

  /** Closes window k of a document with several windows; false when it has only one (the document closes). */
  function closeWindow(id: string): boolean {
    const { doc, k } = parse(id);
    if (!doc || count(doc.wb) < 2) return false;
    const wb = doc.wb;
    show(wb, k === 0 ? 1 : 0);
    if (k === 0) {
      // Window 1 (now shown in place) takes window 0's place; window 0's views (swapped into slot 1) go.
      for (const s of wb.sheets) s.views?.shift();
      wb.windows!.splice(1, 1);
      wb.shownWindow = undefined;
    } else {
      for (const s of wb.sheets) s.views?.splice(k - 1, 1);
      wb.windows!.splice(k, 1);
    }
    if (wb.windows!.length < 2 && !wb.windows![0].hidden) {
      delete wb.windows;
      for (const s of wb.sheets) delete s.views;
    }
    app.markChanged();
    if (doc === app.activeDoc) app.activate(doc);
    else app.refreshTabs();
    if (tiles) layout();
    return true;
  }

  function hide(): void {
    const d = app.activeDoc, wb = d.wb, k = shown(wb);
    if (!wb.windows) wb.windows = [{ active: wb.active }];
    wb.windows[k].hidden = true;
    app.markChanged();
    const next = visibleIds().find((x) => x !== idOf(d, k));
    if (next) { document.body.classList.remove("no-window"); select(next); return; }
    // No window left to show: Excel shows its empty frame.
    document.body.classList.add("no-window");
    app.refreshTabs();
  }

  function visibleIds(): string[] {
    return app.docList.flatMap((d) => tabs(d).map((t) => t.id));
  }
  function hiddenList(): { id: string; name: string }[] {
    return app.docList.flatMap((d) => [...Array(count(d.wb)).keys()].filter((k) => hidden(d.wb, k)).map((k) => ({ id: `${d.id}#${k}`, name: caption(d, k) })));
  }

  async function unhide(): Promise<void> {
    const list = hiddenList();
    if (!list.length) return;
    let pick = 0;
    const box = h("div", { class: "cv-list", role: "listbox" });
    const draw = () => box.replaceChildren(...list.map((w, i) => h("div", { class: i === pick ? "on" : "", role: "option", onclick: () => { pick = i; draw(); } }, w.name)));
    draw();
    css();
    const ok = await showDialog({ title: "Einblenden", body: h("div", null, h("div", null, "Arbeitsmappe einblenden:"), box), buttons: [{ label: "OK", value: true, primary: true }, { label: "Abbrechen", value: false }], cancel: false });
    if (!ok) return;
    const { doc, k } = parse(list[pick].id);
    if (!doc) return;
    delete doc.wb.windows![k].hidden;
    if (doc.wb.windows!.length < 2) { delete doc.wb.windows; for (const s of doc.wb.sheets) delete s.views; }
    document.body.classList.remove("no-window");
    app.markChanged();
    select(idOf(doc, k));
  }

  function switchMenu(): MenuEntry[] {
    const cur = activeId(app.activeDoc);
    return app.docList.flatMap((d) => tabs(d)).map((t, i) => ({ label: `${i + 1} ${t.name}`, checked: t.id === cur, run: () => select(t.id) }));
  }

  // ----- Alle anordnen, Nebeneinander anzeigen ------------------------------------------------------

  async function arrangeDialog(): Promise<void> {
    css();
    const opt = (v: Mode, label: string) => h("label", { class: "check" }, h("input", { type: "radio", name: "arr", value: v, checked: (arranged?.mode ?? "tiled") === v ? true : undefined }), label);
    const only = h("input", { type: "checkbox", checked: arranged?.activeOnly ? true : undefined }) as HTMLInputElement;
    const body = h("div", null, h("fieldset", { class: "group" }, h("legend", null, "Anordnen"), opt("tiled", "Unterteilt"), opt("horizontal", "Horizontal"), opt("vertical", "Vertikal"), opt("cascade", "Überlappend")),
      h("label", { class: "check" }, only, "Fenster der aktiven Arbeitsmappe"));
    const ok = await showDialog({ title: "Fenster anordnen", body, buttons: [{ label: "OK", value: true, primary: true }, { label: "Abbrechen", value: false }], cancel: false });
    if (!ok) return;
    const mode = (body.querySelector<HTMLInputElement>("input[name=arr]:checked")?.value ?? "tiled") as Mode;
    arrange(mode, only.checked);
  }

  /** The windows laid out inside our window (one of them alone: back to the normal layout). */
  function arrange(mode: Mode | null, activeOnly = false): void {
    arranged = mode ? { mode, activeOnly } : null;
    if (!arranged) side = null;
    layout();
  }

  function windowsToShow(): { doc: Doc; k: number; id: string }[] {
    const d = app.activeDoc;
    if (side) {
      const o = parse(side.other);
      const list = [{ doc: d, k: shown(d.wb), id: activeId(d) }];
      if (o.doc) list.push({ doc: o.doc, k: o.k, id: side.other });
      return list;
    }
    const docs = arranged?.activeOnly ? [d] : app.docList;
    return docs.flatMap((doc) => [...Array(count(doc.wb)).keys()].filter((k) => !hidden(doc.wb, k)).map((k) => ({ doc, k, id: idOf(doc, k) })));
  }

  /** Builds the tiles (or takes them down); the main grid goes into the active window's tile. */
  function layout(): void {
    const gridEl = app.grid.el;
    const list = arranged || side ? windowsToShow() : [];
    for (const p of passive) p.grid.el.remove();
    passive = [];
    if (list.length < 2) {
      if (tiles) { tiles.replaceWith(gridEl); tiles = null; }
      app.grid.refreshLayout();
      return;
    }
    css();
    if (!tiles) { tiles = h("div", { class: "win-tiles" }); gridEl.replaceWith(tiles); }
    const mode: Mode = side ? "horizontal" : arranged!.mode;
    tiles.className = `win-tiles ${mode}`;
    const n = list.length;
    const cols = mode === "vertical" ? n : mode === "horizontal" ? 1 : Math.ceil(Math.sqrt(n));
    const rows = mode === "cascade" ? 1 : Math.ceil(n / cols);
    tiles.style.gridTemplateColumns = mode === "cascade" ? "" : `repeat(${cols}, minmax(0, 1fr))`;
    tiles.style.gridTemplateRows = mode === "cascade" ? "" : `repeat(${rows}, minmax(0, 1fr))`;
    const cur = activeId(app.activeDoc);
    const items: HTMLElement[] = [];
    list.forEach((w, i) => {
      const on = w.id === cur;
      const body = h("div", { class: "win-body" });
      const tile = h("div", { class: `win-tile${on ? " on" : ""}` }, h("div", { class: "win-cap" }, caption(w.doc, w.k)), body);
      if (mode === "cascade") Object.assign(tile.style, { left: `${i * 24}px`, top: `${i * 24}px`, zIndex: String(on ? 100 : i) });
      if (on) body.append(gridEl);
      else {
        const g = passiveGrid(w.doc, w.k);
        body.append(g.el, h("div", { class: "win-shield", onpointerdown: (e: PointerEvent) => { e.preventDefault(); select(w.id); } }));
        passive.push({ grid: g, id: w.id, doc: w.doc, k: w.k });
      }
      items.push(tile);
    });
    tiles.replaceChildren(...items);
    requestAnimationFrame(() => { app.grid.refreshLayout(); for (const p of passive) p.grid.refreshLayout(); });
  }

  /** A grid that draws another window: the workbook seen through that window's views, nothing editable. */
  function passiveGrid(doc: Doc, k: number): Grid {
    const wb = doc.wb;
    const sheets = wb.sheets.map((s) => Object.create(s, { view: { get: () => viewOf(wb, k, s), configurable: true } }) as Sheet);
    const book = Object.create(wb, {
      active: { get: () => activeOf(wb, k), set: () => {}, configurable: true },
      sheets: { value: sheets, configurable: true },
    }) as Workbook;
    const noop = () => undefined;
    const host = new Proxy({} as GridHost, {
      get: (_t, key) => (key === "workbook" ? () => book : key === "showFormulas" || key === "key" || key === "enterPaste" ? () => false
        : key === "tableStyle" ? app.tableStyleOf : key === "canEdit" ? () => false : noop),
    });
    const g = new Grid(host);
    g.el.style.pointerEvents = "none";
    return g;
  }

  /** After every paint of the main grid: the other windows follow edits, Synchrones Scrollen moves the other one. */
  let pending = 0;
  function painted(): void {
    if (!passive.length || pending) return;
    pending = requestAnimationFrame(() => {
      pending = 0;
      const v = app.sheet.view;
      if (side && sync && lastScroll && lastScroll.sheet === app.sheet) {
        const dt = v.top - lastScroll.top, dl = v.left - lastScroll.left;
        if (dt || dl) for (const p of passive) {
          const ov = viewOf(p.doc.wb, p.k, p.doc.wb.sheets[activeOf(p.doc.wb, p.k)]);
          ov.top = Math.max(0, ov.top + dt);
          ov.left = Math.max(0, ov.left + dl);
          p.grid.refreshLayout();
        }
      }
      lastScroll = { top: v.top, left: v.left, sheet: app.sheet };
      for (const p of passive) p.grid.contentChanged();
    });
  }

  async function sideBySide(): Promise<void> {
    if (side) { side = null; layout(); return; }
    const cur = activeId(app.activeDoc);
    const others = app.docList.flatMap((d) => tabs(d)).filter((t) => t.id !== cur);
    if (!others.length) return;
    let pick = 0;
    if (others.length > 1) {
      css();
      const box = h("div", { class: "cv-list", role: "listbox" });
      const draw = () => box.replaceChildren(...others.map((w, i) => h("div", { class: i === pick ? "on" : "", role: "option", onclick: () => { pick = i; draw(); } }, w.name)));
      draw();
      const ok = await showDialog({ title: "Nebeneinander vergleichen", body: h("div", null, h("div", null, `Nebeneinander vergleichen mit: ${app.activeDoc.name}`), box), buttons: [{ label: "OK", value: true, primary: true }, { label: "Abbrechen", value: false }], cancel: false });
      if (!ok) return;
    }
    side = { other: others[pick].id };
    sync = true;
    lastScroll = null;
    layout();
  }

  // ----- Benutzerdefinierte Ansichten -----------------------------------------------------------

  const guid = () => {
    const b = crypto.getRandomValues(new Uint8Array(16));
    const x = [...b].map((v) => v.toString(16).padStart(2, "0").toUpperCase()).join("");
    return `{${x.slice(0, 8)}-${x.slice(8, 12)}-${x.slice(12, 16)}-${x.slice(16, 20)}-${x.slice(20)}}`;
  };
  const nameKey = (g: string) => `Z_${g.slice(1, -1).replace(/-/g, "_")}_`;
  const hasTables = () => app.wb.sheets.some((s) => s.tables?.length);

  function spans(list: number[]): [number, number][] {
    const out: [number, number][] = [];
    for (const i of [...list].sort((a, b) => a - b)) {
      const last = out[out.length - 1];
      if (last && last[1] === i - 1) last[1] = i; else out.push([i, i]);
    }
    return out;
  }

  /** A sheet's customSheetView in Excel's form (recorded: scale, showPageBreaks, hiddenRows, selection, page setup). */
  function sheetViewXml(s: Sheet, g: string, print: boolean, rowCol: boolean): string {
    const v = s.view;
    const hiddenRows = rowCol && [...s.rows.values()].some((r) => r.hidden);
    const hiddenCols = rowCol && [...s.cols.values()].some((c) => c.hidden);
    const tl = v.top || v.left ? ` topLeftCell="${cellName(v.top, v.left)}"` : "";
    const attrs = `guid="${g}"${v.zoom !== 100 ? ` scale="${v.zoom}"` : ""}${v.zoom !== 100 || hiddenRows ? ' showPageBreaks="1"' : ""}${!v.showGrid ? ' showGridLines="0"' : ""}${hiddenRows ? ' hiddenRows="1"' : ""}${hiddenCols ? ' hiddenColumns="1"' : ""}${tl}`;
    const ac = cellName(v.active.row, v.active.col);
    const sq = v.selection.map((r) => (r.r0 === r.r1 && r.c0 === r.c1 ? cellName(r.r0, r.c0) : `${cellName(r.r0, r.c0)}:${cellName(r.r1, r.c1)}`)).join(" ") || ac;
    const sel = ac === "A1" && sq === "A1" ? "" : `<selection activeCell="${ac}" sqref="${sq}"/>`;
    const p = s.page;
    const page = print ? `<pageMargins left="${p?.margins?.left ?? 0.7}" right="${p?.margins?.right ?? 0.7}" top="${p?.margins?.top ?? 0.78740157499999996}" bottom="${p?.margins?.bottom ?? 0.78740157499999996}" header="${p?.margins?.header ?? 0.3}" footer="${p?.margins?.footer ?? 0.3}"/>`
      + (p?.orientation === "landscape" ? '<pageSetup paperSize="9" orientation="landscape"/>' : "") : "";
    return sel || page ? `<customSheetView ${attrs}>${sel}${page}</customSheetView>` : `<customSheetView ${attrs}/>`;
  }

  function addView(name: string, print: boolean, rowCol: boolean): void {
    const wb = app.wb;
    const g = guid();
    app.change("Ansicht hinzufügen", () => {
      // A name used again replaces that view (Excel asks first; the dialog did).
      const old = wb.customViews?.find((c) => c.name.toLowerCase() === name.toLowerCase());
      if (old) removeView(old.guid);
      // windowWidth and windowHeight are required; Excel writes its window's size there (a tenth of the workbookView's, recorded).
      const w0 = wb.windows?.[0]?.attrs ?? {};
      const tenth = (v: string | undefined, d: number) => String(v !== undefined ? Math.round(Number(v) / 10) : d);
      const attrs = { xWindow: tenth(w0.xWindow, -11), yWindow: tenth(w0.yWindow, -11), windowWidth: tenth(w0.windowWidth, 1942), windowHeight: tenth(w0.windowHeight, 1042), activeSheetId: String(wb.active + 1) };
      wb.customViews = [{ name, guid: g, print, rowCol, attrs }, ...(wb.customViews ?? [])];
      wb.sheets.forEach((s, i) => {
        s.customViews = [{ guid: g, xml: sheetViewXml(s, g, print, rowCol) }, ...(s.customViews ?? [])];
        if (!rowCol) return;
        const rows = spans([...s.rows].filter(([, r]) => r.hidden).map(([k]) => k));
        const cols = spans([...s.cols].filter(([, c]) => c.hidden).map(([k]) => k));
        const q = /[^A-Za-z0-9_.]/.test(s.name) ? `'${s.name.replace(/'/g, "''")}'` : s.name;
        if (rows.length) wb.names.push({ name: `${nameKey(g)}.wvu.Rows`, ref: rows.map(([a, b]) => `${q}!$${a + 1}:$${b + 1}`).join(","), sheet: i, hidden: true });
        if (cols.length) wb.names.push({ name: `${nameKey(g)}.wvu.Cols`, ref: cols.map(([a, b]) => `${q}!$${colName(a)}:$${colName(b)}`).join(","), sheet: i, hidden: true });
      });
    });
  }

  function removeView(g: string): void {
    const wb = app.wb;
    wb.customViews = wb.customViews?.filter((c) => c.guid !== g);
    for (const s of wb.sheets) if (s.customViews) s.customViews = s.customViews.filter((c) => c.guid !== g);
    wb.names = wb.names.filter((n) => !n.name.startsWith(nameKey(g)));
  }

  /** Anzeigen: each sheet's zoom, selection and top-left cell, the hidden rows and columns, the page orientation. */
  function showView(g: string): void {
    const wb = app.wb;
    const cv = wb.customViews?.find((c) => c.guid === g);
    if (!cv) return;
    app.change("Ansicht anzeigen", () => {
      wb.sheets.forEach((s, i) => {
        const x = s.customViews?.find((c) => c.guid === g)?.xml;
        if (!x) return;
        app.history.whole(s);
        const a = (n: string) => new RegExp(`\\b${n}="([^"]*)"`).exec(/<customSheetView\b[^>]*>/.exec(x)?.[0] ?? "")?.[1];
        const v = s.view;
        v.zoom = Number(a("scale") ?? 100);
        const tl = parseRange(a("topLeftCell") ?? "A1");
        if (tl) { v.top = tl.r0; v.left = tl.c0; }
        const sel = /<selection\b[^>]*activeCell="([^"]+)"[^>]*sqref="([^"]+)"/.exec(x);
        const ac = parseRange(sel?.[1] ?? "A1")!;
        v.active = { row: ac.r0, col: ac.c0 };
        v.selection = (sel?.[2] ?? "A1").split(" ").map((t) => parseRange(t)).filter((r): r is NonNullable<typeof r> => !!r);
        if (cv.rowCol) {
          for (const r of s.rows.values()) delete r.hidden;
          for (const c of s.cols.values()) delete c.hidden;
          for (const n of wb.names.filter((q) => q.sheet === i && q.name.startsWith(nameKey(g)))) {
            for (const part of n.ref.split(",")) {
              const m = /!\$?([A-Z]+|\d+):\$?([A-Z]+|\d+)$/.exec(part);
              if (!m) continue;
              if (/^\d+$/.test(m[1])) for (let r = Number(m[1]) - 1; r <= Number(m[2]) - 1; r++) s.rows.set(r, { ...(s.rows.get(r) ?? {}), hidden: true });
              else {
                const c0 = parseRange(`${m[1]}1`)!.c0, c1 = parseRange(`${m[2]}1`)!.c0;
                for (let c = c0; c <= c1; c++) s.cols.set(c, { ...(s.cols.get(c) ?? {}), hidden: true });
              }
            }
          }
        }
        if (cv.print) {
          const land = /orientation="landscape"/.test(x);
          s.page = { ...(s.page ?? {}), orientation: land ? "landscape" : "portrait" };
        }
      });
      const id = Number(cv.attrs.activeSheetId ?? 1) - 1;
      if (id >= 0 && id < wb.sheets.length) wb.active = id;
    });
    app.activate(app.activeDoc);
  }

  async function customViews(): Promise<void> {
    css();
    const wb = app.wb;
    let pick = 0;
    const box = h("div", { class: "cv-list", role: "listbox" });
    const showBtn = h("button", { class: "primary" }, "Anzeigen") as HTMLButtonElement;
    const delBtn = h("button", null, "Löschen") as HTMLButtonElement;
    const list = () => [...(wb.customViews ?? [])].sort((a, b) => a.name.localeCompare(b.name, "de"));
    const draw = () => {
      const l = list();
      box.replaceChildren(...l.map((c, i) => h("div", { class: i === pick ? "on" : "", role: "option", onclick: () => { pick = i; draw(); }, ondblclick: () => { pick = i; showBtn.click(); } }, c.name)));
      showBtn.disabled = delBtn.disabled = !l.length;
    };
    draw();
    let close: (v: string) => void = () => {};
    const body = h("div", { class: "cv-row" }, h("div", null, h("div", null, "Ansichten:"), box),
      h("div", { class: "cv-btns" }, showBtn, h("button", { onclick: () => close("close") }, "Schließen"), h("button", { onclick: () => close("add") }, "Hinzufügen..."), delBtn));
    showBtn.onclick = () => close("show");
    delBtn.onclick = async () => {
      const c = list()[pick];
      if (!c) return;
      const yes = await showDialog({ title: "RaabeCalc", body: `Möchten Sie die Ansicht „${c.name}“ löschen?`, buttons: [{ label: "Ja", value: true, primary: true }, { label: "Nein", value: false }], cancel: false });
      if (!yes) return;
      app.change("Ansicht löschen", () => removeView(c.guid));
      pick = 0;
      draw();
    };
    const r = await new Promise<string>((resolve) => {
      close = (v) => { (document.querySelector(".dialog-layer .dlg-title button") as HTMLButtonElement | null)?.click(); resolve(v); };
      void showDialog({ title: "Benutzerdefinierte Ansichten", body, buttons: [], cancel: "close" }).then(() => resolve("close"));
    });
    if (r === "show") { const c = list()[pick]; if (c) showView(c.guid); }
    if (r === "add") await addDialog();
  }

  async function addDialog(): Promise<void> {
    const name = h("input", { type: "text", style: "width:240px" }) as HTMLInputElement;
    const print = h("input", { type: "checkbox", checked: true }) as HTMLInputElement;
    const rowCol = h("input", { type: "checkbox", checked: true }) as HTMLInputElement;
    const body = h("div", null, h("label", { class: "field" }, "Name: ", name),
      h("fieldset", { class: "group" }, h("legend", null, "In Ansicht einbeziehen"),
        h("label", { class: "check" }, print, "Druckeinstellungen"), h("label", { class: "check" }, rowCol, "Ausgeblendete Zeilen, Spalten und Filtereinstellungen")));
    setTimeout(() => name.focus(), 0);
    const ok = await showDialog({ title: "Ansicht hinzufügen", body, buttons: [{ label: "OK", value: true, primary: true }, { label: "Abbrechen", value: false }], cancel: false });
    const n = name.value.trim();
    if (!ok || !n) return;
    if (app.wb.customViews?.some((c) => c.name.toLowerCase() === n.toLowerCase())) {
      const yes = await showDialog({ title: "RaabeCalc", body: `Die Ansicht „${n}“ gibt es schon. Soll sie ersetzt werden?`, buttons: [{ label: "Ja", value: true, primary: true }, { label: "Nein", value: false }], cancel: false });
      if (!yes) return;
    }
    addView(n, print.checked, rowCol.checked);
  }

  // ----- commands ---------------------------------------------------------------------------------

  type Run = (id: string, fn: () => void | Promise<void>, enabled?: () => boolean, checked?: () => boolean) => void;
  function implement(run: Run): void {
    const many = () => app.docList.reduce((n, d) => n + tabs(d).length, 0) > 1;
    run("view.newWindow", () => newWindow());
    run("view.arrange", () => void arrangeDialog());
    run("view.hide", () => hide(), () => !document.body.classList.contains("no-window"));
    run("view.unhide", () => void unhide(), () => hiddenList().length > 0);
    run("view.sideBySide", () => void sideBySide(), many, () => !!side);
    run("view.syncScroll", () => { sync = !sync; lastScroll = null; }, () => !!side, () => !!side && sync);
    run("view.resetPosition", () => layout(), () => !!side);
    run("view.switchWindows", () => {});
    run("view.customViews", () => void customViews(), () => !hasTables());
  }

  return { tabs, activeId, title, select, closeWindow, newWindow, hide, unhide, switchMenu, arrange, layout, painted, implement, customViews, addView, showView, removeView, show, viewOf, activeOf, get arranged() { return arranged; } };
}
