// Datenquelle auswählen (Diagrammentwurf > Daten auswählen..., D-117), as German Excel 365 shows it
// (results/2026-11-09-excel-chart-buttons.json, inventory 2026-11-05):
// - Diagrammdatenbereich with a range field: type a range, or point at cells in the sheet while the field has the
//   focus (the dialog lets clicks through to the cells); the button at its right end shrinks the dialog to the
//   field (Dialog reduzieren) and grows it again.
// - Zeile/Spalte tauschen; Legendeneinträge (Reihen): Hinzufügen, Bearbeiten (Datenreihe bearbeiten: Reihenname,
//   Reihenwerte), Entfernen, Nach oben, Nach unten, a check box per series (cleared: hidden by the filter);
//   Horizontale Achsenbeschriftungen (Rubrik): Bearbeiten (Achsenbeschriftungen: Achsenbeschriftungsbereich) and a
//   check box per category; Ausgeblendete und leere Zellen (Lücken / Null / Datenpunkte mit Linie verbinden,
//   #NV als leere Zelle anzeigen, Daten in ausgeblendeten Zeilen und Spalten anzeigen).
// The chart follows the dialog while it is open; OK keeps it as one undo step, Abbrechen puts it back.

import { kindInfo, isScatter, parseRef, refText, seriesName, sourceOf, type ChartModel, type ChartSeries, type Ref } from "@raabe/charts";
import { h, messageBox, showDialog } from "@raabe/shell";
import { catNames } from "./buttons";

export interface SelectDataHost {
  /** The workbook's sheet names (the first is used for refs without one). */
  sheetName(): string;
  /** The series Excel makes from a range (orientation forced by Zeile/Spalte tauschen). */
  seriesOf(m: ChartModel, ref: Ref, orientation?: "rows" | "cols"): ChartSeries[];
  /** Re-reads names, categories and values of a draft from the cells. */
  read(m: ChartModel): void;
  /** Values of a range as Excel shows them after "=" ("3; 5; 2; 4"), or null. */
  preview(ref: string): string | null;
  /** Shows the draft on the chart (null: the chart itself again). */
  show(m: ChartModel, draft: ChartModel | null): void;
  /** Keeps the draft as one undo step. */
  commit(m: ChartModel, draft: ChartModel): void;
  /** Range picking: the grid element, the cell under a point (client px) and a marquee around a range. */
  grid(): HTMLElement;
  cellAt(clientX: number, clientY: number): { row: number; col: number } | null;
  marquee(rect: { r0: number; c0: number; r1: number; c1: number } | null): void;
}

const CSS = `
.sd-dialog { min-width: 560px; }
.sd-dialog .dlg-body { padding: 6px 18px 8px; font-size: 12px; }
.sd-row { display: flex; align-items: center; gap: 8px; margin: 6px 0; }
.sd-row > span:first-child { white-space: nowrap; }
.rf { display: inline-flex; align-items: stretch; border: 1px solid #c7c7c7; border-radius: 2px; background: #fff; flex: 1; min-width: 0; }
.dialog .rf input[type=text] { border: 0; height: 24px; flex: 1; min-width: 0; border-radius: 0; }
.rf button { width: 24px; border: 0; border-left: 1px solid #e1e1e1; background: #fff; display: flex; align-items: center; justify-content: center; padding: 0; }
.rf button:hover { background: #e6f2ea; }
.rf.picking { outline: 2px solid #107c41; }
.sd-center { display: flex; justify-content: center; margin: 10px 0; }
.sd-center button, .sd-tool button, .sd-foot button { font: inherit; font-size: 12px; height: 26px; padding: 0 10px; border: 1px solid #c7c7c7; border-radius: 3px; background: #fff; display: inline-flex; align-items: center; gap: 5px; }
.sd-center button:hover, .sd-tool button:hover:not(:disabled), .sd-foot button:hover { background: #f0f0f0; }
.sd-tool button:disabled { color: #a6a6a6; }
.sd-panes { display: flex; gap: 14px; }
.sd-pane { flex: 1; display: flex; flex-direction: column; min-width: 0; border: 1px solid #d6d6d6; }
.sd-pane-title { font-weight: 600; padding: 5px 6px; border-bottom: 1px solid #e6e6e6; }
.sd-tool { display: flex; gap: 4px; padding: 5px 6px; border-bottom: 1px solid #e6e6e6; flex-wrap: wrap; }
.sd-list { height: 150px; overflow: auto; padding: 2px 0; outline: none; }
.sd-item { display: flex; align-items: center; gap: 6px; padding: 2px 6px; height: 20px; cursor: default; }
.sd-item.sel { background: #cfe3d7; }
.sd-list:focus .sd-item.sel { background: #b5d6c2; }
.sd-item input { margin: 0; }
.sd-foot { display: flex; align-items: center; margin-top: 10px; }
.sd-preview { color: #595959; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; max-width: 160px; }
.dialog.rf-collapsed .dlg-body > :not(.rf-bar), .dialog.rf-collapsed .dlg-buttons { display: none !important; }
.rf-bar { display: none; }
.dialog.rf-collapsed .rf-bar { display: flex; }
.dialog.rf-collapsed { min-width: 420px !important; }
.dialog-layer.rf-through { pointer-events: none; background: transparent; }
.dialog-layer.rf-through .dialog { pointer-events: auto; }
`;

const COLLAPSE = '<svg width="14" height="12" viewBox="0 0 14 12"><rect x=".5" y=".5" width="13" height="11" fill="#fff" stroke="#444"/><path d="M2 3h10M2 9h6" stroke="#107c41" stroke-width="1.2"/><path d="m9 6 2 2.5L13 6" fill="none" stroke="#444"/></svg>';
const EXPAND = '<svg width="14" height="12" viewBox="0 0 14 12"><rect x=".5" y=".5" width="13" height="11" fill="#fff" stroke="#444"/><path d="M2 3h10M2 6h10M2 9h10" stroke="#107c41" stroke-width="1.2"/></svg>';

/** The range a series' formula field refers to, or literal values ("={1;2;3}"). */
export function parseValues(text: string): { ref?: string; values?: (number | null)[] } | null {
  const t = text.trim().replace(/^=/, "").trim();
  if (!t) return null;
  const lit = /^\{(.*)\}$/.exec(t);
  if (lit) {
    const vals = lit[1].split(/[;.]/).map((x) => x.trim()).map((x) => (x === "" ? null : Number(x.replace(",", "."))));
    return vals.some((v) => v !== null && !Number.isFinite(v)) ? null : { values: vals };
  }
  return parseRef(t) ? { ref: t } : null;
}

/** A series name typed into Reihenname: a reference, or text ("Nord" or ="Nord"). */
export function parseName(text: string): { ref?: string; text?: string } | null {
  const t = text.trim();
  if (!t) return {};
  if (t.startsWith("=")) {
    const r = t.slice(1).trim();
    const q = /^"(.*)"$/.exec(r);
    if (q) return { text: q[1].replace(/""/g, '"') };
    return parseRef(r) ? { ref: r } : null;
  }
  return { text: t };
}

const withSheet = (ref: string, sheet: string) => (ref.includes("!") ? ref : `${/[^A-Za-z0-9_.]/.test(sheet) ? `'${sheet.replace(/'/g, "''")}'` : sheet}!${ref}`);

/** A field for a range with Excel's Dialog reduzieren button; pointing at cells writes the range. */
export function refField(_host: SelectDataHost, value: string, label: string, opts: { onChange?: (v: string) => void; dialog: () => HTMLElement | null }) {
  const input = h("input", { type: "text", value, "aria-label": label, spellcheck: "false" }) as HTMLInputElement;
  const btn = h("button", { type: "button", "aria-label": "Dialog reduzieren", "data-tip": "Dialog reduzieren", tabindex: "-1" });
  btn.innerHTML = COLLAPSE;
  const box = h("span", { class: "rf" }, input, btn);
  let home: { parent: Node; next: Node | null } | null = null;
  const collapse = () => {
    const d = opts.dialog();
    if (!d) return;
    // The dialogs below (Datenquelle auswählen under Datenreihe bearbeiten) go out of the way too, as in Excel.
    const others = [...document.querySelectorAll<HTMLElement>(".dialog-layer")].filter((l) => !l.contains(d));
    if (d.classList.contains("rf-collapsed")) {
      for (const l of others) l.style.visibility = "";
      d.classList.remove("rf-collapsed");
      Object.assign(d.style, { position: "", left: "", top: "", margin: "" });
      if (home) home.parent.insertBefore(box, home.next);
      home = null;
      btn.innerHTML = COLLAPSE;
      btn.setAttribute("aria-label", "Dialog erweitern");
    } else {
      home = { parent: box.parentNode!, next: box.nextSibling };
      let bar = d.querySelector<HTMLElement>(".rf-bar");
      if (!bar) { bar = h("div", { class: "rf-bar sd-row" }); d.querySelector(".dlg-body")!.prepend(bar); }
      bar.replaceChildren(box);
      // The shrunk dialog keeps its top left corner where the dialog stood (Excel; D-155).
      const at = d.getBoundingClientRect();
      Object.assign(d.style, { position: "fixed", left: `${Math.round(at.left)}px`, top: `${Math.round(at.top)}px`, margin: "0" });
      d.classList.add("rf-collapsed");
      for (const l of others) l.style.visibility = "hidden";
      btn.innerHTML = EXPAND;
      btn.setAttribute("aria-label", "Dialog erweitern");
    }
    input.focus();
  };
  btn.addEventListener("click", collapse);
  input.addEventListener("input", () => opts.onChange?.(input.value));
  // Enter in a shrunk dialog grows it again (Excel).
  input.addEventListener("keydown", (e) => {
    if (e.key === "Enter" && opts.dialog()?.classList.contains("rf-collapsed")) { e.preventDefault(); e.stopPropagation(); collapse(); }
  });
  input.addEventListener("focus", () => { active = { input, set: (v) => { input.value = v; opts.onChange?.(v); } }; });
  return { el: box, input, collapse };
}

/** The range field that gets what the mouse points at (the one focused last in the open dialog). */
let active: { input: HTMLInputElement; set: (v: string) => void } | null = null;

/**
 * Lets clicks through to the cells while a dialog with range fields is open: a press on a cell starts a range in
 * the focused field (Excel writes it absolute, with the sheet's name); presses elsewhere outside the dialog are
 * swallowed (the dialog stays modal for the ribbon and the rest).
 */
export function pointThrough(host: SelectDataHost, layer: HTMLElement, dialog: HTMLElement): () => void {
  layer.classList.add("rf-through");
  let drag: { r: number; c: number } | null = null;
  const grid = host.grid();
  const write = (r0: number, c0: number, r1: number, c1: number) => {
    const rect = { r0: Math.min(r0, r1), c0: Math.min(c0, c1), r1: Math.max(r0, r1), c1: Math.max(c0, c1) };
    host.marquee(rect);
    active?.set("=" + refText(host.sheetName(), rect.r0, rect.c0, rect.r1, rect.c1));
  };
  // Inside any dialog (this one or one opened from it) nothing is swallowed.
  const inDialog = (t: EventTarget | null) => !!(t as Element | null)?.closest?.(".dialog");
  const down = (e: PointerEvent) => {
    const t = e.target as Node;
    if (inDialog(t) || !layer.isConnected) return;
    e.preventDefault();
    e.stopPropagation();
    if (!grid.contains(t) || !active || !dialog.contains(active.input)) return;
    const cell = host.cellAt(e.clientX, e.clientY);
    if (!cell) return;
    drag = { r: cell.row, c: cell.col };
    write(cell.row, cell.col, cell.row, cell.col);
    active.input.closest(".rf")?.classList.add("picking");
  };
  const move = (e: PointerEvent) => {
    if (!drag) return;
    e.preventDefault();
    e.stopPropagation();
    const cell = host.cellAt(e.clientX, e.clientY);
    if (cell) write(drag.r, drag.c, cell.row, cell.col);
  };
  const up = (e: PointerEvent) => {
    if (!drag) return;
    e.preventDefault();
    e.stopPropagation();
    drag = null;
    active?.input.closest(".rf")?.classList.remove("picking");
    active?.input.focus();
  };
  const swallow = (e: Event) => { if (!inDialog(e.target) && layer.isConnected) { e.stopPropagation(); e.preventDefault(); } };
  addEventListener("pointerdown", down, true);
  addEventListener("pointermove", move, true);
  addEventListener("pointerup", up, true);
  addEventListener("mousedown", swallow, true);
  addEventListener("click", swallow, true);
  addEventListener("dblclick", swallow, true);
  addEventListener("contextmenu", swallow, true);
  const wheel = (e: WheelEvent) => { if (!grid.contains(e.target as Node) && !inDialog(e.target)) e.preventDefault(); };
  addEventListener("wheel", wheel, { capture: true, passive: false });
  return () => {
    removeEventListener("pointerdown", down, true);
    removeEventListener("pointermove", move, true);
    removeEventListener("pointerup", up, true);
    removeEventListener("mousedown", swallow, true);
    removeEventListener("click", swallow, true);
    removeEventListener("dblclick", swallow, true);
    removeEventListener("contextmenu", swallow, true);
    removeEventListener("wheel", wheel, true);
    host.marquee(null);
  };
}

/** Opens a dialog made by showDialog and returns its elements (the dialog is the last one in the page). */
function openDialog<T>(opts: Parameters<typeof showDialog<T>>[0], cls: string): { done: Promise<T>; dialog: HTMLElement; layer: HTMLElement } {
  const done = showDialog(opts);
  const layers = document.querySelectorAll<HTMLElement>(".dialog-layer");
  const layer = layers[layers.length - 1];
  const dialog = layer.querySelector<HTMLElement>(".dialog")!;
  dialog.classList.add(cls);
  return { done, dialog, layer };
}

function ensureCss(): void {
  if (!document.getElementById("raabe-selectdata-css")) document.head.append(h("style", { id: "raabe-selectdata-css" }, CSS));
}

/** The draft's Diagrammdatenbereich: the block its references span (Excel shows nothing when they are scattered). */
function rangeText(m: ChartModel): string {
  const live = m.series.filter((s) => s.val.ref);
  const src = sourceOf(live.length ? live : m.series);
  return src ? "=" + refText(src.ref.sheet, src.ref.r0, src.ref.c0, src.ref.r1, src.ref.c1) : "";
}

type Amount = { ref?: string; cache: (number | null)[] };
const amountText = (d: Amount | undefined) => (d?.ref ? "=" + d.ref : `={${(d?.cache.length ? d.cache : [1]).map((v) => v ?? "").join(";")}}`);

/**
 * Fehlerindikatoren > Benutzerdefiniert > Wert angeben (D-155): Excel's "Benutzerdefinierte Fehlerindikatoren" with
 * Positiver Fehlerwert and Negativer Fehlerwert, each a range (pointed at in the sheet, or typed) or values ("={1}").
 * Resolves to the two amounts, or null for Abbrechen.
 */
export async function customErrorDialog(host: SelectDataHost, plus: Amount | undefined, minus: Amount | undefined): Promise<{ plus: Amount; minus: Amount } | null> {
  ensureCss();
  let sub: HTMLElement | null = null;
  const pf = refField(host, amountText(plus), "Positiver Fehlerwert", { dialog: () => sub });
  const mf = refField(host, amountText(minus), "Negativer Fehlerwert", { dialog: () => sub });
  const body = h("div", null,
    h("div", null, "Positiver Fehlerwert"), h("div", { class: "sd-row" }, pf.el),
    h("div", { style: "margin-top:8px" }, "Negativer Fehlerwert"), h("div", { class: "sd-row" }, mf.el));
  const o = openDialog({ title: "Benutzerdefinierte Fehlerindikatoren", body, buttons: [{ label: "OK", value: true, primary: true }, { label: "Abbrechen", value: false }], cancel: false }, "sd-dialog");
  sub = o.dialog;
  sub.style.minWidth = "300px";
  sub.classList.add("ce-dialog");
  const stop = pointThrough(host, o.layer, o.dialog);
  pf.input.focus();
  pf.input.select();
  const ok = await o.done;
  stop();
  if (!ok) return null;
  const sheet = host.sheetName();
  const read = (text: string): Amount | null => {
    const v = parseValues(text);
    if (!v) return null;
    return v.ref ? { ref: withSheet(v.ref, sheet), cache: [] } : { cache: v.values ?? [] };
  };
  const p = read(pf.input.value), q = read(mf.input.value);
  if (!p || !q) {
    await messageBox("RaabeCalc", "Der eingegebene Bezug ist ungültig.");
    return null;
  }
  return { plus: p, minus: q };
}

export async function selectDataDialog(host: SelectDataHost, m: ChartModel): Promise<void> {
  ensureCss();
  let draft: ChartModel = structuredClone(m);
  let orientation = sourceOf(draft.series)?.orientation;
  let selSeries = 0, selCat = 0;
  let dialogEl: HTMLElement | null = null;

  const range = refField(host, rangeText(draft), "Diagrammdatenbereich", {
    dialog: () => dialogEl,
    onChange: (v) => {
      const p = parseValues(v);
      if (!p?.ref) return;
      const r = parseRef(withSheet(p.ref, host.sheetName()));
      if (!r) return;
      fromRange(r, undefined);
    },
  });
  const seriesList = h("div", { class: "sd-list", tabindex: "0", role: "listbox", "aria-label": "Legendeneinträge (Reihen)" });
  const catList = h("div", { class: "sd-list", tabindex: "0", role: "listbox", "aria-label": "Horizontale Achsenbeschriftungen (Rubrik)" });
  const btn = (label: string, run: () => void, tip?: string) => {
    const b = h("button", { type: "button", "aria-label": tip ?? label, "data-tip": tip ?? label }, label) as HTMLButtonElement;
    b.addEventListener("click", run);
    return b;
  };

  const update = () => {
    host.read(draft);
    host.show(m, draft);
    fill();
  };
  const fromRange = (r: Ref, o: "rows" | "cols" | undefined) => {
    const keep = draft;
    draft = { ...structuredClone(keep), series: host.seriesOf(keep, r, o) };
    delete draft.hiddenCats;
    orientation = sourceOf(draft.series)?.orientation ?? o;
    update();
  };

  const up = btn("▲", () => move(-1), "Nach oben");
  const down = btn("▼", () => move(1), "Nach unten");
  const edit = btn("Bearbeiten", () => void editSeries(selSeries), "Datenreihe bearbeiten");
  const remove = btn("Entfernen", () => {
    if (!draft.series.length) return;
    draft.series.splice(selSeries, 1);
    selSeries = Math.max(0, Math.min(selSeries, draft.series.length - 1));
    range.input.value = rangeText(draft);
    update();
  }, "Datenreihe entfernen");
  const catEdit = btn("Bearbeiten", () => void editCats(), "Kategorie bearbeiten");

  function move(d: -1 | 1): void {
    const i = selSeries, j = i + d;
    if (j < 0 || j >= draft.series.length) return;
    [draft.series[i], draft.series[j]] = [draft.series[j], draft.series[i]];
    selSeries = j;
    update();
  }

  function fill(): void {
    seriesList.replaceChildren(...draft.series.map((s, i) => {
      const box = h("input", { type: "checkbox", tabindex: "-1", "aria-label": seriesName(s, i) }) as HTMLInputElement;
      box.checked = !s.filtered;
      box.addEventListener("change", () => { if (box.checked) delete draft.series[i].filtered; else draft.series[i].filtered = true; host.show(m, draft); });
      const row = h("div", { class: `sd-item${i === selSeries ? " sel" : ""}`, role: "option", "aria-selected": String(i === selSeries) }, box, seriesName(s, i));
      row.addEventListener("mousedown", () => { selSeries = i; fill(); });
      row.addEventListener("dblclick", () => void editSeries(i));
      return row;
    }));
    const hidden = new Set(draft.hiddenCats ?? []);
    catList.replaceChildren(...catNames(draft).map((c, i) => {
      const box = h("input", { type: "checkbox", tabindex: "-1", "aria-label": c }) as HTMLInputElement;
      box.checked = !hidden.has(i);
      box.addEventListener("change", () => {
        const set = new Set(draft.hiddenCats ?? []);
        if (box.checked) set.delete(i); else set.add(i);
        if (set.size) draft.hiddenCats = [...set].sort((a, b) => a - b); else delete draft.hiddenCats;
        host.show(m, draft);
      });
      const row = h("div", { class: `sd-item${i === selCat ? " sel" : ""}`, role: "option" }, box, c);
      row.addEventListener("mousedown", () => { selCat = i; fill(); });
      return row;
    }));
    up.disabled = selSeries <= 0;
    down.disabled = selSeries >= draft.series.length - 1;
    edit.disabled = remove.disabled = !draft.series.length;
    catEdit.disabled = !draft.series.length;
  }
  seriesList.addEventListener("keydown", (e) => {
    if (e.key === "ArrowDown" || e.key === "ArrowUp") { e.preventDefault(); selSeries = Math.max(0, Math.min(draft.series.length - 1, selSeries + (e.key === "ArrowDown" ? 1 : -1))); fill(); }
  });

  /** Datenreihe bearbeiten: Reihenname and Reihenwerte (null: Hinzufügen). */
  async function editSeries(i: number | null): Promise<void> {
    const s: ChartSeries | null = i === null ? null : draft.series[i];
    if (i !== null && !s) return;
    let sub: HTMLElement | null = null;
    const namePrev = h("span", { class: "sd-preview" });
    const valPrev = h("span", { class: "sd-preview" });
    const nameF = refField(host, s ? (s.name?.ref ? "=" + s.name.ref : s.name?.text ?? "") : "", "Reihenname", { dialog: () => sub, onChange: () => show() });
    const valsF = refField(host, s ? (s.val.ref ? "=" + s.val.ref : `={${s.val.cache.map((v) => v ?? "").join(";")}}`) : "={1}", "Reihenwerte", { dialog: () => sub, onChange: () => show() });
    const show = () => {
      const n = parseName(nameF.input.value);
      namePrev.textContent = n?.ref ? `= ${host.preview(withSheet(n.ref, host.sheetName())) ?? ""}` : n?.text !== undefined && n.text !== "" ? `= ${n.text}` : "";
      const v = parseValues(valsF.input.value);
      valPrev.textContent = v?.ref ? `= ${host.preview(withSheet(v.ref, host.sheetName())) ?? ""}` : v?.values ? `= ${v.values.map((x) => x ?? "").join("; ")}` : "";
    };
    show();
    const body = h("div", null,
      h("div", null, "Reihenname:"), h("div", { class: "sd-row" }, nameF.el, namePrev),
      h("div", { style: "margin-top:8px" }, "Reihenwerte:"), h("div", { class: "sd-row" }, valsF.el, valPrev));
    const o = openDialog({ title: "Datenreihe bearbeiten", body, buttons: [{ label: "OK", value: true, primary: true }, { label: "Abbrechen", value: false }], cancel: false }, "sd-dialog");
    sub = o.dialog;
    sub.style.minWidth = "420px";
    const stop = pointThrough(host, o.layer, o.dialog);
    nameF.input.focus();
    const ok = await o.done;
    stop();
    if (!ok) return;
    const n = parseName(nameF.input.value);
    const v = parseValues(valsF.input.value);
    if (n === null || v === null) {
      await messageBox("RaabeCalc", "Der eingegebene Bezug ist ungültig.");
      return;
    }
    const sheet = host.sheetName();
    const next: ChartSeries = { ...(s ?? { val: { cache: [] } }), name: n.ref ? { ref: withSheet(n.ref, sheet) } : n.text !== undefined ? { text: n.text } : undefined,
      val: v.ref ? { ref: withSheet(v.ref, sheet), cache: [] } : { cache: v.values ?? [] } };
    if (!next.name) delete next.name;
    if (!s) {
      // A new series takes the categories of the first one (Excel).
      const cat = draft.series[0]?.cat;
      if (cat) next.cat = structuredClone(cat);
      draft.series.push(next);
      selSeries = draft.series.length - 1;
    } else draft.series[i!] = next;
    range.input.value = rangeText(draft);
    update();
  }

  /** Achsenbeschriftungen: the range of the category labels (for every series). */
  async function editCats(): Promise<void> {
    let sub: HTMLElement | null = null;
    const first = draft.series[0];
    const prev = h("span", { class: "sd-preview" });
    const f = refField(host, first?.cat?.ref ? "=" + first.cat.ref : "", "Achsenbeschriftungsbereich", { dialog: () => sub, onChange: (v) => {
      const p = parseValues(v);
      prev.textContent = p?.ref ? `= ${host.preview(withSheet(p.ref, host.sheetName())) ?? ""}` : "";
    } });
    const p0 = parseValues(f.input.value);
    prev.textContent = p0?.ref ? `= ${host.preview(withSheet(p0.ref, host.sheetName())) ?? ""}` : "";
    const o = openDialog({ title: "Achsenbeschriftungen", body: h("div", null, h("div", null, "Achsenbeschriftungsbereich:"), h("div", { class: "sd-row" }, f.el, prev)),
      buttons: [{ label: "OK", value: true, primary: true }, { label: "Abbrechen", value: false }], cancel: false }, "sd-dialog");
    sub = o.dialog;
    sub.style.minWidth = "420px";
    const stop = pointThrough(host, o.layer, o.dialog);
    f.input.focus();
    const ok = await o.done;
    stop();
    if (!ok) return;
    const p = parseValues(f.input.value);
    const t = f.input.value.trim();
    if (t && t !== "=" && !p?.ref) { await messageBox("RaabeCalc", "Der eingegebene Bezug ist ungültig."); return; }
    for (const s of draft.series) {
      if (p?.ref) s.cat = { ref: withSheet(p.ref, host.sheetName()), cache: [] };
      else delete s.cat;
    }
    delete draft.hiddenCats;
    range.input.value = rangeText(draft);
    update();
  }

  /** Ausgeblendete und leere Zellen. */
  async function hiddenCells(): Promise<void> {
    const lineLike = ["line", "scatter", "radar"].includes(kindInfo(draft.kind).family) || isScatter(draft.kind);
    const radio = (label: string, v: "gap" | "zero" | "span", disabled = false) => {
      const r = h("input", { type: "radio", name: "sd-blanks", value: v, disabled }) as HTMLInputElement;
      r.checked = (draft.dispBlanksAs ?? "gap") === v;
      return { r, el: h("label", { class: "check", style: "margin:4px 0 4px 14px" }, r, label) };
    };
    const gap = radio("Lücken", "gap"), zero = radio("Null", "zero"), span = radio("Datenpunkte mit einer Linie verbinden", "span", !lineLike);
    const na = h("input", { type: "checkbox" }) as HTMLInputElement;
    na.checked = draft.naAsBlank ?? true;
    const hid = h("input", { type: "checkbox" }) as HTMLInputElement;
    hid.checked = !!draft.showHidden;
    const body = h("div", null, h("div", null, "Leere Zellen anzeigen als:"), gap.el, zero.el, span.el,
      h("label", { class: "check", style: "margin:4px 0 4px 14px" }, na, '"#N/V" als leere Zelle anzeigen'),
      h("label", { class: "check" }, hid, "Daten in ausgeblendeten Zeilen und Spalten anzeigen"));
    const ok = await showDialog({ title: "Ausgeblendete und leere Zelleneinstellungen", body, buttons: [{ label: "OK", value: true, primary: true }, { label: "Abbrechen", value: false }], cancel: false });
    if (!ok) return;
    const v = [gap, zero, span].find((x) => x.r.checked)?.r.value as "gap" | "zero" | "span" | undefined;
    if (!v || v === "gap") delete draft.dispBlanksAs; else draft.dispBlanksAs = v;
    if (na.checked) delete draft.naAsBlank; else draft.naAsBlank = false;
    if (hid.checked) draft.showHidden = true; else delete draft.showHidden;
    update();
  }

  fill();
  const add = btn("Hinzufügen", () => void editSeries(null), "Datenreihe hinzufügen");
  const swap = btn("Zeile/Spalte tauschen", () => {
    const src = sourceOf(draft.series);
    if (!src) return;
    fromRange(src.ref, src.orientation === "cols" ? "rows" : "cols");
    range.input.value = rangeText(draft);
  });
  const body = h("div", null,
    h("div", { class: "sd-row" }, h("span", null, "Diagrammdatenbereich:"), range.el),
    h("div", { class: "sd-center" }, swap),
    h("div", { class: "sd-panes" },
      h("div", { class: "sd-pane" }, h("div", { class: "sd-pane-title" }, "Legendeneinträge (Reihen)"), h("div", { class: "sd-tool" }, add, edit, remove, up, down), seriesList),
      h("div", { class: "sd-pane" }, h("div", { class: "sd-pane-title" }, "Horizontale Achsenbeschriftungen (Rubrik)"), h("div", { class: "sd-tool" }, catEdit), catList)),
    h("div", { class: "sd-foot" }, btn("Ausgeblendete und leere Zellen", () => void hiddenCells())));
  void orientation;
  const o = openDialog({ title: "Datenquelle auswählen", body, buttons: [{ label: "OK", value: true, primary: true }, { label: "Abbrechen", value: false }], cancel: false }, "sd-dialog");
  dialogEl = o.dialog;
  const stop = pointThrough(host, o.layer, o.dialog);
  range.input.focus();
  range.input.select();
  const ok = await o.done;
  stop();
  host.show(m, null);
  if (!ok) return;
  if (!draft.series.length) { await messageBox("RaabeCalc", "Ein Diagramm braucht mindestens eine Datenreihe."); return; }
  host.commit(m, draft);
}
