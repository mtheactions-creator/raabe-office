// Spezialfilter (Daten > Sortieren und Filtern > Erweitert...; D-090): the dialog, filtering the list in place (rows
// hidden, the sheet in filter mode) or copying the records to another place, with Keine Duplikate. As Excel, the
// ranges are kept as the sheet's names _FilterDatenbank (list), Kriterien and Ziel (the copy's heading row), so
// Löschen and Erneut anwenden work on them and Excel finds them in the file.

import { h, messageBox, showDialog } from "@raabe/shell";
import { advancedFilter, parseRange, rangeName, type Rect, type Sheet } from "@raabe/sheet-engine";
import type { CalcApp } from "./app";

const FILTER_DB = "_xlnm._FilterDatabase", CRITERIA = "_xlnm.Criteria", EXTRACT = "_xlnm.Extract";

export function advancedFilterFeature(app: CalcApp) {
  const si = () => app.wb.sheets.indexOf(app.sheet);
  const q = (s: Sheet) => (/^[\p{L}_][\p{L}\p{N}_.]*$/u.test(s.name) ? s.name : `'${s.name.replace(/'/g, "''")}'`);
  const abs = (r: Rect) => rangeName(r).replace(/([A-Z]+)(\d+)/g, "$$$1$$$2");

  /** A range the sheet's name refers to ("Tabelle1!$A$1:$D$13"), or null. */
  function named(name: string): Rect | null {
    const n = app.wb.names.find((x) => x.name === name && x.sheet === si());
    return n ? parseRange(n.ref.replace(/^=/, "").replace(/^.*!/, "").replace(/\$/g, "")) : null;
  }

  function setName(name: string, r: Rect | null): void {
    const i = si();
    app.wb.names = app.wb.names.filter((x) => !(x.name === name && x.sheet === i));
    if (r) app.wb.names.push({ name, ref: `${q(app.sheet)}!${abs(r)}`, sheet: i, ...(name === FILTER_DB ? { hidden: true } : {}) });
  }

  /** The list around the active cell (its current region), as Excel proposes it. */
  function region(): Rect {
    const a = app.sheet.view.active;
    const sel = app.sheet.view.selection[0];
    if (sel && (sel.r1 > sel.r0 || sel.c1 > sel.c0)) return sel;
    const has = (r: number, c: number) => { const v = app.sheet.value(r, c); return v !== null && v !== undefined && v !== ""; };
    let r0 = a.row, r1 = a.row, c0 = a.col, c1 = a.col;
    for (let grew = true; grew;) {
      grew = false;
      const any = (r: number, ca: number, cb: number) => { for (let c = ca; c <= cb; c++) if (has(r, c)) return true; return false; };
      const anyC = (c: number, ra: number, rb: number) => { for (let r = ra; r <= rb; r++) if (has(r, c)) return true; return false; };
      if (r0 > 0 && any(r0 - 1, Math.max(0, c0 - 1), c1 + 1)) { r0--; grew = true; }
      if (any(r1 + 1, Math.max(0, c0 - 1), c1 + 1)) { r1++; grew = true; }
      if (c0 > 0 && anyC(c0 - 1, Math.max(0, r0 - 1), r1 + 1)) { c0--; grew = true; }
      if (anyC(c1 + 1, Math.max(0, r0 - 1), r1 + 1)) { c1++; grew = true; }
    }
    return { r0, c0, r1, c1 };
  }

  function evaluate(f: string, row: number, col: number) {
    return app.calc.evaluate(si(), row, col, f);
  }

  /** Filters (or copies); false when a range is not usable. */
  function run(list: Rect, crit: Rect, dest: Rect | null, unique: boolean): boolean {
    const sheet = app.sheet;
    const rows = advancedFilter(app.wb, sheet, list, crit, unique, evaluate);
    app.change("Spezialfilter", () => {
      app.history.whole(sheet);
      setName(FILTER_DB, list);
      setName(CRITERIA, crit);
      if (!dest) {
        const keep = new Set(rows);
        for (let r = list.r0 + 1; r <= list.r1; r++) {
          const info = sheet.rows.get(r);
          const hide = !keep.has(r);
          if (!!info?.hidden !== hide) sheet.rows.set(r, { ...(info ?? {}), hidden: hide ? true : undefined });
        }
        sheet.filterMode = rows.length < list.r1 - list.r0;
        return;
      }
      // Copy: the list's headings (or the ones already in the target row, in their order) and the records under them.
      const width = list.c1 - list.c0 + 1;
      const ownHeads: number[] = [];
      for (let c = dest.c0; c <= dest.c1 && dest.c1 > dest.c0; c++) {
        const t = String(sheet.value(dest.r0, c) ?? "").trim().toLowerCase();
        for (let k = 0; k < width; k++) if (String(sheet.value(list.r0, list.c0 + k) ?? "").trim().toLowerCase() === t) { ownHeads.push(k); break; }
      }
      const cols = ownHeads.length ? ownHeads : Array.from({ length: width }, (_, k) => k);
      // Excel empties what was below the target before (as far as the old copy went).
      const old = named(EXTRACT);
      if (old) for (let r = old.r0; r <= old.r0 + 1 + (list.r1 - list.r0); r++) for (let c = dest.c0; c < dest.c0 + cols.length; c++) sheet.set(r, c, undefined);
      cols.forEach((k, i) => { const c = sheet.get(list.r0, list.c0 + k); sheet.set(dest.r0, dest.c0 + i, c ? { ...c } : undefined); });
      rows.forEach((r, j) => cols.forEach((k, i) => { const c = sheet.get(r, list.c0 + k); sheet.set(dest.r0 + 1 + j, dest.c0 + i, c ? { ...c, f: undefined } : undefined); }));
      setName(EXTRACT, { r0: dest.r0, c0: dest.c0, r1: dest.r0, c1: dest.c0 + cols.length - 1 });
    });
    return true;
  }

  async function dialog(): Promise<void> {
    const list = named(FILTER_DB) ?? region();
    const crit = named(CRITERIA);
    const extract = named(EXTRACT);
    const field = (label: string, value: string) => {
      const input = h("input", { type: "text", value, style: "width:220px" }) as HTMLInputElement;
      return { input, el: h("label", { class: "field" }, label, input) };
    };
    const radio = (label: string, checked: boolean) => {
      const input = h("input", { type: "radio", name: "advf", checked }) as HTMLInputElement;
      return { input, el: h("label", { class: "check" }, input, label) };
    };
    const inPlace = radio("Liste an gleicher Stelle filtern", true), copy = radio("An eine andere Stelle kopieren", false);
    const lr = field("Listenbereich:", abs(list)), cr = field("Kriterienbereich:", crit ? abs(crit) : ""), to = field("Kopieren nach:", extract ? abs(extract) : "");
    to.input.disabled = true;
    inPlace.input.addEventListener("change", () => { to.input.disabled = inPlace.input.checked; });
    copy.input.addEventListener("change", () => { to.input.disabled = !copy.input.checked; });
    const unique = h("input", { type: "checkbox" }) as HTMLInputElement;
    const body = h("div", { class: "advf" }, h("fieldset", null, h("legend", null, "Aktion"), inPlace.el, copy.el), lr.el, cr.el, to.el,
      h("label", { class: "check" }, unique, "Keine Duplikate"));
    const ok = await showDialog({ title: "Spezialfilter", body, buttons: [{ label: "OK", value: true, primary: true }, { label: "Abbrechen", value: false }], cancel: false });
    if (!ok) return;
    const read = (t: string) => parseRange(t.replace(/\$/g, "").replace(/^.*!/, "").trim());
    const l = read(lr.input.value), c = read(cr.input.value);
    if (!l || l.r1 <= l.r0) { await messageBox("RaabeCalc", "Der Listenbereich ist ungültig."); return; }
    if (cr.input.value.trim() && (!c || c.r1 <= c.r0)) { await messageBox("RaabeCalc", "Der Kriterienbereich ist ungültig."); return; }
    let dest: Rect | null = null;
    if (copy.input.checked) {
      dest = read(to.input.value);
      if (!dest) { await messageBox("RaabeCalc", "Der Bereich 'Kopieren nach' ist ungültig."); return; }
    }
    run(l, c ?? { r0: l.r0, c0: l.c0, r1: l.r0, c1: l.c0 }, dest, unique.checked);
  }

  /** Whether the sheet has a Spezialfilter in place (Löschen, Erneut anwenden). */
  function active(): boolean {
    return !app.sheet.autoFilter && !!named(FILTER_DB) && !!app.sheet.filterMode;
  }

  /** Löschen: all records shown again. */
  function clear(): void {
    const list = named(FILTER_DB);
    if (!list) return;
    const sheet = app.sheet;
    app.change("Filter löschen", () => {
      app.history.whole(sheet);
      for (let r = list.r0 + 1; r <= list.r1; r++) { const info = sheet.rows.get(r); if (info?.hidden) sheet.rows.set(r, { ...info, hidden: undefined }); }
      sheet.filterMode = undefined;
    });
  }

  /** Erneut anwenden: the same list and criteria again (in place). */
  function reapply(): void {
    const list = named(FILTER_DB), crit = named(CRITERIA);
    if (list && crit) run(list, crit, null, false);
  }

  return { dialog, run, active, clear, reapply };
}
