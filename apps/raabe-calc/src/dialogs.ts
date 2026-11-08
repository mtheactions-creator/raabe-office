// RaabeCalc's dialogs and the menus behind them, with Excel's German titles and fields:
// Zellen formatieren (Strg+1), Spaltenbreite, Zeilenhöhe, Zellen einfügen/löschen, Gehe zu,
// Suchen und Ersetzen, Sortieren, AutoFilter, Text in Spalten, Duplikate entfernen, Namen,
// Notizen, Funktion einfügen, Zoom, Inhalte einfügen, bedingte Formatierung.

import {
  MAX_COLS, MAX_ROWS, cellName, colName, FUNCTIONS_DE, FORMAT_CATEGORIES, localFormat, fileFormat, parseRange, rangeName, shiftFormula,
  type Cell, type CondRule, type Rect, type Sheet,
} from "@raabe/sheet-engine";
import { colorGrid, h, messageBox, showDialog, showMenu, type MenuEntry } from "@raabe/shell";
import type { CalcApp } from "./app";
import * as A from "./actions";
import { shown, editText } from "./engine";
import { DEFAULT_COL_CHARS, charsToPx, shownChars, shownToChars } from "./grid/metrics";
import type { Settings } from "./store";
import { CUSTOM_OPS, NUMBER_FILTERS, TEXT_FILTERS, fromCustom, passes, prepare as prepareFilter, toCustom } from "./filters";

const field = (label: string, input: HTMLElement) => h("label", { class: "field" }, label, input);
const radio = (name: string, label: string, value: string, checked = false) =>
  h("label", { class: "check" }, h("input", { type: "radio", name, value, checked: checked || undefined }), label);
const checkbox = (label: string, checked = false) => {
  const input = h("input", { type: "checkbox", checked: checked || undefined }) as HTMLInputElement;
  return { el: h("label", { class: "check" }, input, label), input };
};
const text = (value = "", width = 200) => h("input", { type: "text", value, style: `width:${width}px` }) as HTMLInputElement;
const picked = (root: HTMLElement, name: string) => root.querySelector<HTMLInputElement>(`input[name=${name}]:checked`)?.value ?? "";
const de = (n: number) => String(Math.round(n * 100) / 100).replace(".", ",");
const num = (s: string) => Number(s.replace(",", ".").trim());

export function dialogs(app: CalcApp) {
  const sel = () => app.grid.selection.at(-1)!;
  const sheet = () => app.sheet;

  // ----- sizes --------------------------------------------------------------------------------

  async function colWidth(): Promise<void> {
    const r = sel();
    const w = sheet().cols.get(r.c0)?.width;
    const cur = shownChars(w ?? sheet().defaultColWidth ?? DEFAULT_COL_CHARS);
    const input = text(de(Math.round(cur * 100) / 100), 80);
    const ok = await showDialog({ title: "Spaltenbreite", body: field("Spaltenbreite:", input), buttons: [{ label: "OK", value: true, primary: true }, { label: "Abbrechen", value: false }], cancel: false });
    if (!ok) return;
    const v = num(input.value);
    if (!(v >= 0 && v <= 255)) return void messageBox("RaabeCalc", "Die Spaltenbreite muss zwischen 0 und 255 liegen.");
    app.setColWidths(range(r.c0, r.c1), v === 0 ? 0 : charsToPx(shownToChars(v)));
  }

  async function defaultWidth(): Promise<void> {
    const input = text(de(sheet().defaultColWidth ?? DEFAULT_COL_CHARS), 80);
    const ok = await showDialog({ title: "Standardbreite", body: field("Standardspaltenbreite:", input), buttons: [{ label: "OK", value: true, primary: true }, { label: "Abbrechen", value: false }], cancel: false });
    if (!ok) return;
    const v = num(input.value);
    if (!(v > 0 && v <= 255)) return;
    app.change("Standardbreite", () => { app.history.whole(sheet()); sheet().defaultColWidth = v; });
  }

  async function rowHeight(): Promise<void> {
    const r = sel();
    const cur = sheet().rows.get(r.r0)?.height ?? sheet().defaultRowHeight;
    const input = text(de(cur), 80);
    const ok = await showDialog({ title: "Zeilenhöhe", body: field("Zeilenhöhe:", input), buttons: [{ label: "OK", value: true, primary: true }, { label: "Abbrechen", value: false }], cancel: false });
    if (!ok) return;
    const v = num(input.value);
    if (!(v >= 0 && v <= 409)) return void messageBox("RaabeCalc", "Die Zeilenhöhe muss zwischen 0 und 409 liegen.");
    app.setRowHeights(range(r.r0, r.r1), (v * 96) / 72);
  }

  // ----- cells ---------------------------------------------------------------------------------

  async function insertCells(): Promise<void> {
    const body = h("div", null, radio("ins", "Zellen nach rechts verschieben", "right"), radio("ins", "Zellen nach unten verschieben", "down", true),
      radio("ins", "Ganze Zeile", "row"), radio("ins", "Ganze Spalte", "col"));
    const ok = await showDialog({ title: "Zellen einfügen", body, buttons: [{ label: "OK", value: true, primary: true }, { label: "Abbrechen", value: false }], cancel: false });
    if (!ok) return;
    const r = sel();
    const v = picked(body, "ins");
    if (v === "row") return void app.commands.run("cells.insertRows");
    if (v === "col") return void app.commands.run("cells.insertCols");
    app.change("Zellen einfügen", () => A.shiftCells(app.ctx, sheet(), r, v as "right" | "down"));
  }

  async function deleteCells(): Promise<void> {
    const body = h("div", null, radio("del", "Zellen nach links verschieben", "left", true), radio("del", "Zellen nach oben verschieben", "up"),
      radio("del", "Ganze Zeile", "row"), radio("del", "Ganze Spalte", "col"));
    const ok = await showDialog({ title: "Löschen", body, buttons: [{ label: "OK", value: true, primary: true }, { label: "Abbrechen", value: false }], cancel: false });
    if (!ok) return;
    const r = sel();
    const v = picked(body, "del");
    if (v === "row") return void app.commands.run("cells.deleteRows");
    if (v === "col") return void app.commands.run("cells.deleteCols");
    app.change("Zellen löschen", () => A.shiftCells(app.ctx, sheet(), r, v as "left" | "up"));
  }

  // ----- Zellen formatieren ---------------------------------------------------------------------

  async function formatCells(tab: "Zahlen" | "Ausrichtung" | "Schrift" | "Rahmen" | "Ausfüllen" | "Schutz"): Promise<void> {
    const st = structuredClone(app.activeStyle());
    const { row, col } = app.sheet.view.active;
    const sample = app.sheet.get(row, col);
    const tabs = ["Zahlen", "Ausrichtung", "Schrift", "Rahmen", "Ausfüllen", "Schutz"] as const;
    const panels: Record<string, HTMLElement> = {};
    const nav = h("div", { class: "fc-tabs", role: "tablist" });
    const body = h("div", { class: "fc" }, nav);

    // Zahlen: categories on the left, the sample and the category's options on the right.
    const cats = FORMAT_CATEGORIES;
    const catList = h("select", { size: 12, class: "fc-cats" }) as HTMLSelectElement;
    for (const c of cats) catList.append(h("option", { value: c.name }, c.name));
    const codeBox = text(localFormat(st.numFmt), 260);
    const codeList = h("select", { size: 8, class: "fc-codes" }) as HTMLSelectElement;
    const preview = h("div", { class: "fc-sample" });
    const showSample = () => {
      try { preview.textContent = shown(sample, { ...st, numFmt: fileFormat(codeBox.value) }).text || " "; } catch { preview.textContent = "Ungültiges Format"; }
    };
    const fillCodes = () => {
      const c = cats.find((x) => x.name === catList.value);
      codeList.replaceChildren(...(c?.codes ?? []).map((f) => h("option", { value: f }, localFormat(f))));
      if (c && c.codes.length) {
        const cur = c.codes.find((f) => f === st.numFmt) ?? c.codes[0];
        codeList.value = cur;
        codeBox.value = localFormat(cur);
      }
      if (c?.name === "Standard") codeBox.value = "Standard";
      if (c?.name === "Text") codeBox.value = "@";
      showSample();
    };
    const curCat = cats.find((c) => c.codes.includes(st.numFmt))?.name ?? (st.numFmt === "General" ? "Standard" : "Benutzerdefiniert");
    catList.value = curCat;
    catList.onchange = fillCodes;
    codeList.onchange = () => { codeBox.value = localFormat(codeList.value); showSample(); };
    codeBox.oninput = showSample;
    if (curCat !== "Benutzerdefiniert") fillCodes();
    else { codeList.replaceChildren(...(cats.find((c) => c.name === "Benutzerdefiniert")?.codes ?? []).map((f) => h("option", { value: f }, localFormat(f)))); showSample(); }
    panels.Zahlen = h("div", { class: "fc-num" }, h("div", null, h("div", { class: "fc-label" }, "Kategorie:"), catList),
      h("div", { class: "fc-right" }, h("div", { class: "fc-label" }, "Beispiel"), preview, h("div", { class: "fc-label" }, "Typ:"), codeBox, codeList));

    // Ausrichtung.
    const hSel = h("select", null, ...[["general", "Standard"], ["left", "Links (Einzug)"], ["center", "Zentriert"], ["right", "Rechts (Einzug)"], ["fill", "Ausfüllen"], ["justify", "Blocksatz"], ["centerContinuous", "Über Auswahl zentrieren"], ["distributed", "Verteilt (Einzug)"]]
      .map(([v, l]) => h("option", { value: v, selected: (st.h ?? "general") === v || undefined }, l))) as HTMLSelectElement;
    const vSel = h("select", null, ...[["top", "Oben"], ["center", "Zentriert"], ["bottom", "Unten"], ["justify", "Blocksatz"], ["distributed", "Verteilt"]]
      .map(([v, l]) => h("option", { value: v, selected: (st.v ?? "bottom") === v || undefined }, l))) as HTMLSelectElement;
    const indent = h("input", { type: "number", min: "0", max: "250", value: String(st.indent ?? 0), style: "width:60px" }) as HTMLInputElement;
    const rot = h("input", { type: "number", min: "-90", max: "90", value: String(st.rotation === undefined ? 0 : st.rotation > 90 && st.rotation <= 180 ? 90 - st.rotation : st.rotation === 255 ? 0 : st.rotation), style: "width:60px" }) as HTMLInputElement;
    const wrap = checkbox("Zeilenumbruch", !!st.wrap), shrink = checkbox("An Zellgröße anpassen", !!st.shrink), mergeBox = checkbox("Zellen verbinden", app.sheet.merges.some((m) => A.overlaps(m, sel())));
    panels.Ausrichtung = h("div", null, h("div", { class: "fc-label" }, "Textausrichtung"), field("Horizontal:", hSel), field("Vertikal:", vSel), field("Einzug:", indent),
      h("div", { class: "fc-label" }, "Textsteuerung"), wrap.el, shrink.el, mergeBox.el, field("Ausrichtung (Grad):", rot));

    // Schrift.
    const fName = text(st.font.name, 220), fSize = text(de(st.font.size), 60);
    const fStyle = h("select", null, ...["Standard", "Kursiv", "Fett", "Fett Kursiv"].map((v) => h("option", { value: v }, v))) as HTMLSelectElement;
    fStyle.value = st.font.bold && st.font.italic ? "Fett Kursiv" : st.font.bold ? "Fett" : st.font.italic ? "Kursiv" : "Standard";
    const fUnder = h("select", null, ...[["", "Ohne"], ["single", "Einfach"], ["double", "Doppelt"], ["singleAccounting", "Einfach (Buchhaltung)"], ["doubleAccounting", "Doppelt (Buchhaltung)"]]
      .map(([v, l]) => h("option", { value: v, selected: (st.font.underline ?? "") === v || undefined }, l))) as HTMLSelectElement;
    let fColor = st.font.color ?? "";
    const fColorBtn = h("button", { type: "button", class: "fc-color", style: `background:${fColor || "#000"}` });
    fColorBtn.onclick = () => openColors(fColorBtn, (c) => { fColor = c; fColorBtn.style.background = c || "#000"; });
    const strike = checkbox("Durchgestrichen", !!st.font.strike), sup = checkbox("Hochgestellt", st.font.vertAlign === "superscript"), sub = checkbox("Tiefgestellt", st.font.vertAlign === "subscript");
    panels.Schrift = h("div", null, field("Schriftart:", fName), field("Schriftschnitt:", fStyle), field("Schriftgrad:", fSize), field("Unterstreichung:", fUnder), field("Farbe:", fColorBtn),
      h("div", { class: "fc-label" }, "Darstellung"), strike.el, sup.el, sub.el);

    // Rahmen: presets (Ohne, Außen, Innen) and the four sides.
    const borders = structuredClone(st.borders);
    let lineStyle = "thin";
    const styleSel = h("select", null, ...[["thin", "Dünn"], ["medium", "Mittel"], ["thick", "Dick"], ["double", "Doppelt"], ["dotted", "Gepunktet"], ["dashed", "Gestrichelt"], ["hair", "Haarlinie"]]
      .map(([v, l]) => h("option", { value: v }, l))) as HTMLSelectElement;
    styleSel.onchange = () => { lineStyle = styleSel.value; };
    const sideBox = (k: "top" | "bottom" | "left" | "right", label: string) => {
      const c = checkbox(label, !!borders[k]);
      c.input.onchange = () => { borders[k] = c.input.checked ? { style: lineStyle } : undefined; };
      return c;
    };
    const sides = [sideBox("top", "Oben"), sideBox("bottom", "Unten"), sideBox("left", "Links"), sideBox("right", "Rechts")];
    let preset: "none" | "outline" | "inside" | null = null;
    panels.Rahmen = h("div", null, field("Linienart:", styleSel), h("div", { class: "fc-label" }, "Voreinstellungen"),
      h("div", { class: "fc-presets" },
        h("button", { type: "button", onclick: () => { preset = "none"; for (const s of sides) s.input.checked = false; for (const k of ["top", "bottom", "left", "right"] as const) borders[k] = undefined; } }, "Keine"),
        h("button", { type: "button", onclick: () => { preset = "outline"; for (const s of sides) s.input.checked = true; } }, "Außen"),
        h("button", { type: "button", onclick: () => { preset = "inside"; } }, "Innen")),
      h("div", { class: "fc-label" }, "Rahmen"), ...sides.map((s) => s.el));

    // Ausfüllen.
    let fill = st.fill ?? "";
    const fillSample = h("div", { class: "fc-fill-sample", style: `background:${fill || "#fff"}` });
    const grid = colorGrid((c) => { fill = c; fillSample.style.background = c || "#fff"; });
    panels.Ausfüllen = h("div", null, h("button", { type: "button", onclick: () => { fill = ""; fillSample.style.background = "#fff"; } }, "Keine Farbe"),
      h("div", { class: "fc-palette" }, ...grid), h("div", { class: "fc-label" }, "Beispiel"), fillSample);

    // Schutz.
    const locked = checkbox("Gesperrt", st.locked !== false), hidden = checkbox("Ausgeblendet", !!st.hidden);
    panels.Schutz = h("div", null, locked.el, hidden.el,
      h("p", { class: "muted" }, "Das Sperren von Zellen oder Ausblenden von Formeln wirkt erst, wenn das Blatt geschützt ist."));

    const show = (t: string) => {
      for (const [k, p] of Object.entries(panels)) p.hidden = k !== t;
      for (const b of nav.children) b.classList.toggle("on", (b as HTMLElement).dataset.tab === t);
    };
    for (const t of tabs) nav.append(h("button", { type: "button", "data-tab": t, role: "tab", onclick: () => show(t) }, t));
    body.append(...Object.values(panels));
    show(tab);
    // What the dialog showed when it opened: OK changes only what the user changed in it, every other
    // property keeps each selected cell's own value (Excel; a mixed selection must not take the active
    // cell's format, D-142).
    const shownAt = () => ({
      code: codeBox.value, h: hSel.value, v: vSel.value, indent: indent.value, rot: rot.value, wrap: wrap.input.checked, shrink: shrink.input.checked,
      name: fName.value.trim(), size: fSize.value.trim(), style: fStyle.value, under: fUnder.value, color: fColor, strike: strike.input.checked,
      sup: sup.input.checked, sub: sub.input.checked, fill, locked: locked.input.checked, hidden: hidden.input.checked,
    });
    const initial = shownAt();
    const ok = await showDialog({ title: "Zellen formatieren", body, buttons: [{ label: "OK", value: true, primary: true }, { label: "Abbrechen", value: false }], cancel: false });
    if (!ok) return;
    const now = shownAt();
    const changed = (k: keyof typeof now) => now[k] !== initial[k];
    let code: string;
    try { code = codeBox.value === "Standard" ? "General" : fileFormat(codeBox.value); } catch { code = st.numFmt; }
    const before = app.activeStyle();
    const changedBorders = JSON.stringify(borders) !== JSON.stringify(before.borders) || preset !== null;
    app.change("Zellen formatieren", () => {
      A.format(app.ctx, sheet(), app.grid.selection, (s) => {
        if (changed("code")) s.numFmt = code;
        if (changed("h")) s.h = hSel.value === "general" ? undefined : hSel.value as never;
        if (changed("v")) s.v = vSel.value === "bottom" ? undefined : vSel.value as never;
        if (changed("indent")) s.indent = Number(indent.value) || undefined;
        const deg = Number(rot.value) || 0;
        if (changed("rot")) s.rotation = deg === 0 ? undefined : deg > 0 ? deg : 90 - deg;
        if (changed("wrap")) s.wrap = wrap.input.checked || undefined;
        if (changed("shrink")) s.shrink = shrink.input.checked || undefined;
        if (changed("name")) s.font.name = fName.value.trim() || s.font.name;
        const size = num(fSize.value);
        if (changed("size") && size >= 1 && size <= 409) s.font.size = size;
        if (changed("style")) {
          s.font.bold = fStyle.value.includes("Fett") || undefined;
          s.font.italic = fStyle.value.includes("Kursiv") || undefined;
        }
        if (changed("under")) s.font.underline = (fUnder.value || undefined) as never;
        if (changed("color")) s.font.color = fColor || undefined;
        if (changed("strike")) s.font.strike = strike.input.checked || undefined;
        if (changed("sup") || changed("sub")) s.font.vertAlign = sup.input.checked ? "superscript" : sub.input.checked ? "subscript" : undefined;
        if (changed("fill")) { s.fill = fill || undefined; s.fillXml = undefined; }
        if (changed("locked")) s.locked = locked.input.checked ? undefined : false;
        if (changed("hidden")) s.hidden = hidden.input.checked || undefined;
      });
      if (changedBorders) applyBorderSides(borders, preset, lineStyle);
      // Taller text (Schriftgrad, Zeilenumbruch, Drehung): the rows fit in the same step (Excel).
      if (changed("size") || changed("wrap") || changed("rot") || changed("name")) for (const r of app.grid.selection) app.autoRowHeights(r);
      const merged = app.sheet.merges.some((m) => A.overlaps(m, sel()));
      if (mergeBox.input.checked !== merged) {
        app.history.merges(sheet());
        if (mergeBox.input.checked) sheet().merges.push(sel());
        else sheet().merges = sheet().merges.filter((m) => !A.overlaps(m, sel()));
      }
    });
  }

  function applyBorderSides(b: { top?: { style: string }; bottom?: { style: string }; left?: { style: string }; right?: { style: string } }, preset: string | null, style: string): void {
    const r = sel();
    app.history.cells(sheet(), r);
    for (let row = r.r0; row <= r.r1; row++) for (let col = r.c0; col <= r.c1; col++) {
      const c = sheet().get(row, col) ?? { v: null };
      const s = app.wb.styles.derive(c.s, (st) => {
        if (preset === "none") { st.borders = {}; return; }
        const outer = { top: row === r.r0, bottom: row === r.r1, left: col === r.c0, right: col === r.c1 };
        for (const k of ["top", "bottom", "left", "right"] as const) {
          if (preset === "inside") { if (!outer[k]) st.borders[k] = { style }; continue; }
          if (outer[k]) st.borders[k] = b[k] ? { style: b[k]!.style } : undefined;
        }
      });
      sheet().set(row, col, { ...c, s });
    }
  }

  function openColors(anchor: HTMLElement, pick: (hex: string) => void): void {
    const pop = h("div", { class: "menu fc-pop" }, h("button", { type: "button", onclick: () => { pick(""); pop.remove(); } }, "Automatisch"),
      h("div", { class: "fc-palette" }, ...colorGrid((c) => { pick(c); pop.remove(); })));
    const r = anchor.getBoundingClientRect();
    pop.style.cssText = `position:fixed;left:${r.left}px;top:${r.bottom + 2}px;z-index:90`;
    document.body.append(pop);
    setTimeout(() => addEventListener("pointerdown", function close(e) { if (!pop.contains(e.target as Node)) { pop.remove(); removeEventListener("pointerdown", close, true); } }, true));
  }

  // ----- Suchen und Ersetzen, Gehe zu -----------------------------------------------------------

  let lastFind = "";
  async function find(mode: "Suchen" | "Ersetzen"): Promise<void> {
    const what = text(lastFind, 260), repl = text("", 260);
    const matchCase = checkbox("Groß-/Kleinschreibung beachten"), whole = checkbox("Gesamten Zellinhalt vergleichen");
    const inSel = h("select", null, h("option", { value: "sheet" }, "Blatt"), h("option", { value: "book" }, "Arbeitsmappe")) as HTMLSelectElement;
    const lookIn = h("select", null, h("option", { value: "formulas" }, "Formeln"), h("option", { value: "values" }, "Werte")) as HTMLSelectElement;
    const status = h("div", { class: "muted" });
    const matches = (c: Cell, s: Sheet) => {
      const t = lookIn.value === "values" ? shown(c, app.wb.styles.get(c.s)).text : editText(c, app.wb.styles.get(c.s));
      void s;
      const a = matchCase.input.checked ? t : t.toLowerCase(), b = matchCase.input.checked ? what.value : what.value.toLowerCase();
      return whole.input.checked ? a === b : a.includes(b);
    };
    const all = () => {
      const out: { sheet: number; row: number; col: number }[] = [];
      const sheets = inSel.value === "book" ? app.wb.sheets.map((_, i) => i) : [app.wb.active];
      for (const i of sheets) {
        const s = app.wb.sheets[i];
        const keys = [...s.cells.keys()].sort((x, y) => x - y);
        for (const k of keys) {
          const c = s.cells.get(k)!;
          if (c.v === null && !c.f) continue;
          if (matches(c, s)) out.push({ sheet: i, row: Math.floor(k / MAX_COLS), col: k % MAX_COLS });
        }
      }
      return out;
    };
    const next = () => {
      lastFind = what.value;
      if (!what.value) return;
      const list = all();
      if (!list.length) { status.textContent = "Die gesuchten Daten wurden nicht gefunden."; return; }
      const a = app.sheet.view.active;
      const cur = app.wb.active;
      const after = list.find((m) => m.sheet > cur || (m.sheet === cur && (m.row > a.row || (m.row === a.row && m.col > a.col)))) ?? list[0];
      if (after.sheet !== app.wb.active) { app.wb.active = after.sheet; app.grid.refreshLayout(); }
      app.grid.selectCell(after.row, after.col);
      status.textContent = "";
    };
    const replaceOne = () => {
      const a = app.sheet.view.active;
      const c = app.sheet.get(a.row, a.col);
      if (c && matches(c, app.sheet)) {
        const t = editText(c, app.wb.styles.get(c.s));
        const re = new RegExp(what.value.replace(/[.*+?^${}()|[\]\\]/g, "\\$&"), matchCase.input.checked ? "g" : "gi");
        app.change("Ersetzen", () => A.enter(app.ctx, app.sheet, { r0: a.row, c0: a.col, r1: a.row, c1: a.col }, t.replace(re, repl.value)));
      }
      next();
    };
    const replaceAll = () => {
      const list = all();
      const re = new RegExp(what.value.replace(/[.*+?^${}()|[\]\\]/g, "\\$&"), matchCase.input.checked ? "g" : "gi");
      app.change("Ersetzen", () => {
        for (const m of list) {
          const s = app.wb.sheets[m.sheet];
          const c = s.get(m.row, m.col)!;
          A.enter({ wb: app.wb, history: app.history }, s, { r0: m.row, c0: m.col, r1: m.row, c1: m.col }, editText(c, app.wb.styles.get(c.s)).replace(re, repl.value));
        }
      });
      void messageBox("RaabeCalc", `Fertig. Es wurden ${list.length} Ersetzungen vorgenommen.`);
    };
    const body = h("div", { class: "find" }, field("Suchen nach:", what), mode === "Ersetzen" ? field("Ersetzen durch:", repl) : h("span"),
      h("div", { class: "find-opts" }, field("Innerhalb:", inSel), field("Suchen in:", lookIn), matchCase.el, whole.el),
      h("div", { class: "find-btns" },
        mode === "Ersetzen" ? h("button", { type: "button", onclick: replaceAll }, "Alle ersetzen") : h("span"),
        mode === "Ersetzen" ? h("button", { type: "button", onclick: replaceOne }, "Ersetzen") : h("span"),
        h("button", { type: "button", onclick: () => { const l = all(); status.textContent = `${l.length} Zelle(n) gefunden: ` + l.slice(0, 8).map((m) => `${app.wb.sheets[m.sheet].name}!${cellName(m.row, m.col)}`).join(", "); } }, "Alle suchen"),
        h("button", { type: "button", class: "primary", onclick: next }, "Weitersuchen")),
      status);
    await showDialog({ title: "Suchen und Ersetzen", body, buttons: [{ label: "Schließen", value: undefined }], cancel: undefined });
  }

  async function goto(): Promise<void> {
    const ref = text("", 240);
    const names = h("select", { size: 8, style: "width:240px" }, ...app.wb.names.filter((n) => !n.name.startsWith("_xlnm")).map((n) => h("option", { value: n.name }, n.name))) as HTMLSelectElement;
    names.onchange = () => { ref.value = names.value; };
    const ok = await showDialog({ title: "Gehe zu", body: h("div", null, field("Gehe zu:", names), field("Verweis:", ref)), buttons: [{ label: "OK", value: true, primary: true }, { label: "Abbrechen", value: false }], cancel: false });
    if (ok && ref.value.trim() && !app.grid.goTo(ref.value)) void messageBox("RaabeCalc", "Der eingegebene Bezug ist ungültig.");
  }

  // ----- sorting and filtering ----------------------------------------------------------------------

  /** The range to sort: the selection, or the current region around the active cell, with Excel's header guess. */
  function sortRange(from: Rect = sel()): { rect: Rect; header: boolean } {
    let r = from;
    if (r.r0 === r.r1 && r.c0 === r.c1) r = app.grid.currentRegion(r.r0, r.c0);
    // Whole rows or columns: only as far as there is data.
    const used = sheet().usedRange();
    if (used && (r.r1 > used.r1 || r.c1 > used.c1)) r = { ...r, r1: Math.max(r.r0, Math.min(r.r1, used.r1)), c1: Math.max(r.c0, Math.min(r.c1, used.c1)) };
    const firstRowText = range(r.c0, r.c1).every((c) => typeof sheet().get(r.r0, c)?.v === "string");
    const secondRowHasNumbers = range(r.c0, r.c1).some((c) => typeof sheet().get(r.r0 + 1, c)?.v === "number");
    return { rect: r, header: firstRowText && (secondRowHasNumbers || r.r1 > r.r0) && r.r1 > r.r0 && !range(r.c0, r.c1).every((c) => typeof sheet().get(r.r0 + 1, c)?.v === "string" && firstRowText && false) };
  }

  function doSort(rect: Rect, keys: { col: number; asc: boolean }[]): void {
    app.change("Sortieren", () => {
      app.history.cells(sheet(), rect);
      const rows: Cell[][] = [];
      for (let r = rect.r0; r <= rect.r1; r++) rows.push(range(rect.c0, rect.c1).map((c) => sheet().get(r, c) ?? { v: null }));
      const origin = new Map(rows.map((cells, i) => [cells, rect.r0 + i]));
      const rank = (v: Cell["v"]) => (v === null || v === "" ? 5 : typeof v === "number" ? 0 : typeof v === "string" ? 1 : typeof v === "boolean" ? 2 : 3);
      rows.sort((a, b) => {
        for (const k of keys) {
          const x = a[k.col - rect.c0].v, y = b[k.col - rect.c0].v;
          const rx = rank(x), ry = rank(y);
          // Empty cells always last, in both directions (Excel).
          if (rx === 5 || ry === 5) { if (rx !== ry) return rx === 5 ? 1 : -1; continue; }
          let d = 0;
          if (rx !== ry) d = rx - ry;
          else if (typeof x === "number" && typeof y === "number") d = x - y;
          else if (typeof x === "string" && typeof y === "string") d = x.localeCompare(y, "de", { sensitivity: "base" });
          else if (typeof x === "boolean" && typeof y === "boolean") d = Number(x) - Number(y);
          if (d) return k.asc ? d : -d;
        }
        return 0;
      });
      placeRows(rect, rows, origin);
      moveThreads(rect, rows.map((cells) => origin.get(cells)!));
    });
  }

  /**
   * Puts sorted rows into place. A formula moves with its row as if copied there: its relative
   * references follow (=B3*2 from row 3 becomes =B2*2 in row 2), absolute ones stay (Excel, D-142).
   */
  function placeRows(rect: Rect, rows: Cell[][], origin: Map<Cell[], number>): void {
    rows.forEach((cells, i) => {
      const row = rect.r0 + i, d = row - origin.get(cells)!;
      cells.forEach((c, j) => {
        const col = rect.c0 + j;
        sheet().set(row, col, c.v === null && c.s === undefined && !c.f ? undefined : { ...c, ...(c.f !== undefined && d ? { f: shiftFormula(c.f, d, 0) } : {}) });
      });
    });
  }

  /** Sorting: Kommentare in the sorted range go with their rows (`order` lists the old row of each new one). */
  function moveThreads(rect: Rect, order: number[]): void {
    const th = sheet().threads;
    if (!th?.length) return;
    app.history.whole(sheet());
    const newRow = new Map(order.map((old, i) => [old, rect.r0 + i]));
    sheet().threads = th.map((t) => (t.col >= rect.c0 && t.col <= rect.c1 && newRow.has(t.row) ? { ...t, row: newRow.get(t.row)! } : t));
  }

  /** Nach Farbe sortieren: the rows with that colour first, the others after them in their order (Excel's "Oben"). */
  function sortByColor(rect: Rect, col: number, color: { fill?: string | null; font?: string | null }): void {
    app.change("Sortieren", () => {
      app.history.cells(sheet(), rect);
      const rows: Cell[][] = [];
      for (let r = rect.r0; r <= rect.r1; r++) rows.push(range(rect.c0, rect.c1).map((c) => sheet().get(r, c) ?? { v: null }));
      const hit = (cells: Cell[]) => {
        const st = app.wb.styles.get(cells[col - rect.c0].s);
        return passes({ color }, cells[col - rect.c0], "", { fill: st.fill, color: st.font.color });
      };
      const origin = new Map(rows.map((cells, i) => [cells, rect.r0 + i]));
      const sorted = [...rows.filter(hit), ...rows.filter((x) => !hit(x))];
      placeRows(rect, sorted, origin);
    });
  }

  /**
   * Sortierwarnung: a selection of several cells with data beside it. Excel asks whether to take the
   * neighbouring columns along (Markierung erweitern, the default) or to sort the selection alone,
   * which tears the rows apart (D-142). Null: Abbrechen.
   */
  async function sortTarget(): Promise<{ rect: Rect; header: boolean } | null> {
    const s = sel();
    if (s.r0 === s.r1 && s.c0 === s.c1) return sortRange();
    const used = sheet().usedRange();
    if (!used) return sortRange();
    const top = Math.min(s.r0, used.r1), left = Math.min(s.c0, used.c1);
    const near = app.grid.currentRegion(top, left);
    const region = { r0: Math.min(near.r0, s.r0), c0: Math.min(near.c0, s.c0), r1: Math.max(near.r1, Math.min(s.r1, used.r1)), c1: Math.max(near.c1, Math.min(s.c1, used.c1)) };
    if (region.c0 >= s.c0 && region.c1 <= s.c1) return sortRange();
    const body = h("div", null, h("p", null, "Neben der Markierung stehen weitere Daten. Wenn die Markierung nicht erweitert wird, bleiben sie, wo sie sind, und gehören danach nicht mehr zu ihren Zeilen."),
      h("div", { class: "fc-label" }, "Wie möchten Sie fortfahren?"),
      radio("sortw", "Markierung erweitern", "expand", true), radio("sortw", "Mit bestehender Markierung fortfahren", "keep"));
    const ok = await showDialog({ title: "Sortierwarnung", body, buttons: [{ label: "Sortieren", value: true, primary: true }, { label: "Abbrechen", value: false }], cancel: false });
    if (!ok) return null;
    if (picked(body, "sortw") !== "expand") return sortRange(s);
    app.grid.select(region, { ...app.sheet.view.active });
    return sortRange(region);
  }

  async function sortQuick(asc: boolean): Promise<void> {
    const target = await sortTarget();
    if (!target) return;
    const { rect, header } = target;
    const body = header ? { ...rect, r0: rect.r0 + 1 } : rect;
    if (body.r0 > body.r1) return;
    doSort(body, [{ col: app.sheet.view.active.col, asc }]);
  }

  async function sort(): Promise<void> {
    const target = await sortTarget();
    if (!target) return;
    const { rect, header } = target;
    const hasHeader = checkbox("Daten haben Überschriften", header);
    const colSel = h("select", null) as HTMLSelectElement;
    const fill = () => {
      colSel.replaceChildren(...range(rect.c0, rect.c1).map((c) => {
        const t = hasHeader.input.checked ? shown(sheet().get(rect.r0, c), app.wb.styles.get(sheet().get(rect.r0, c)?.s)).text : "";
        return h("option", { value: String(c) }, t || `Spalte ${colName(c)}`);
      }));
      colSel.value = String(app.sheet.view.active.col);
    };
    hasHeader.input.onchange = fill;
    fill();
    const order = h("select", null, h("option", { value: "asc" }, "A bis Z / Nach Größe (aufsteigend)"), h("option", { value: "desc" }, "Z bis A / Nach Größe (absteigend)")) as HTMLSelectElement;
    const col2 = h("select", null, h("option", { value: "" }, "(keine)"), ...range(rect.c0, rect.c1).map((c) => h("option", { value: String(c) }, `Spalte ${colName(c)}`))) as HTMLSelectElement;
    const order2 = h("select", null, h("option", { value: "asc" }, "Aufsteigend"), h("option", { value: "desc" }, "Absteigend")) as HTMLSelectElement;
    const body = h("div", null, hasHeader.el, field("Sortieren nach", colSel), field("Reihenfolge", order), field("Dann nach", col2), field("Reihenfolge", order2));
    const ok = await showDialog({ title: "Sortieren", body, buttons: [{ label: "OK", value: true, primary: true }, { label: "Abbrechen", value: false }], cancel: false });
    if (!ok) return;
    const r = hasHeader.input.checked ? { ...rect, r0: rect.r0 + 1 } : rect;
    const keys = [{ col: Number(colSel.value), asc: order.value === "asc" }];
    if (col2.value) keys.push({ col: Number(col2.value), asc: order2.value === "asc" });
    doSort(r, keys);
  }

  function hasFilter(): boolean {
    return !!app.sheet?.autoFilter;
  }

  function toggleFilter(): void {
    if (app.sheet.autoFilter) {
      app.change("Filtern", () => {
        app.history.whole(sheet());
        const f = sheet().autoFilter!;
        for (let r = f.ref.r0 + 1; r <= f.ref.r1; r++) { const info = sheet().rows.get(r); if (info?.hidden) sheet().rows.set(r, { ...info, hidden: undefined }); }
        sheet().autoFilter = undefined;
      });
      return;
    }
    const { rect } = sortRange();
    app.change("Filtern", () => { app.history.whole(sheet()); sheet().autoFilter = { ref: rect, columns: {} }; });
  }

  function applyFilter(): void {
    const f = sheet().autoFilter;
    if (!f) return;
    // Top 10 and the average need the column's numbers first.
    for (const [k, crit] of Object.entries(f.columns)) {
      if (!crit.top10 && !crit.dynamic) continue;
      const nums: number[] = [];
      for (let r = f.ref.r0 + 1; r <= f.ref.r1; r++) { const v = sheet().get(r, f.ref.c0 + Number(k))?.v; if (typeof v === "number") nums.push(v); }
      f.columns[Number(k)] = prepareFilter(crit, nums);
    }
    for (let r = f.ref.r0 + 1; r <= f.ref.r1; r++) {
      let show = true;
      for (const [k, crit] of Object.entries(f.columns)) {
        const c = sheet().get(r, f.ref.c0 + Number(k));
        const t = shown(c, app.wb.styles.get(c?.s)).text;
        const st = app.wb.styles.get(c?.s);
        if (!passes(crit, c, t, { fill: st.fill, color: st.font.color })) show = false;
      }
      const info = sheet().rows.get(r);
      if (!!info?.hidden === !show) continue;
      app.history.row(sheet(), r);
      sheet().rows.set(r, { ...(info ?? {}), hidden: show ? undefined : true });
    }
  }

  function clearFilter(): void {
    if (!sheet().autoFilter) return;
    app.change("Filter löschen", () => { app.history.whole(sheet()); sheet().autoFilter!.columns = {}; applyFilter(); });
  }

  function reapplyFilter(): void {
    app.change("Erneut anwenden", () => { app.history.whole(sheet()); applyFilter(); });
  }

  function filterByActive(): void {
    const a = app.sheet.view.active;
    if (!sheet().autoFilter) { const { rect } = sortRange(); app.change("Filtern", () => { app.history.whole(sheet()); sheet().autoFilter = { ref: rect, columns: {} }; }); }
    const f = sheet().autoFilter!;
    const c = sheet().get(a.row, a.col);
    const t = shown(c, app.wb.styles.get(c?.s)).text;
    app.change("Filtern", () => { app.history.whole(sheet()); f.columns[a.col - f.ref.c0] = { values: [t] }; applyFilter(); });
  }

  /**
   * The dropdown of a filter button as Excel's panel under the button (recorded with
   * probe_excel_filter.py): sorting labelled by the column's content, Filter entfernen aus "…",
   * the colour and custom filters (not built yet), a search box and the column's values with
   * check boxes, OK and Abbrechen.
   */
  async function filterMenu(col: number, anchor?: DOMRect): Promise<void> {
    const f = sheet().autoFilter;
    if (!f) return;
    // A protected sheet filters only when AutoFilter verwenden is allowed.
    if (app.isProtected() && !app.protectionAllows("autoFilter")) return;
    const values = new Set<string>();
    let blank = false, numbers = 0, texts = 0;
    for (let r = f.ref.r0 + 1; r <= f.ref.r1; r++) {
      const c = sheet().get(r, col);
      const t = shown(c, app.wb.styles.get(c?.s)).text;
      if (typeof c?.v === "number") numbers++; else if (typeof c?.v === "string") texts++;
      if (t === "") blank = true; else values.add(t);
    }
    const head = sheet().get(f.ref.r0, col);
    const name = head && head.v !== null ? shown(head, app.wb.styles.get(head.s)).text : `Spalte ${colName(col)}`;
    const numeric = numbers > texts;
    const crit = f.columns[col - f.ref.c0];
    // The colours in the column, in order of appearance (Nach Farbe sortieren / filtern).
    const fills: string[] = [], fonts: string[] = [];
    let noFill = false, autoFont = false;
    for (let r = f.ref.r0 + 1; r <= f.ref.r1; r++) {
      const st = app.wb.styles.get(sheet().get(r, col)?.s);
      if (st.fill) { if (!fills.some((x) => x.toUpperCase() === st.fill!.toUpperCase())) fills.push(st.fill); } else noFill = true;
      const fc = st.font.color;
      if (fc && fc.toUpperCase() !== "#000000") { if (!fonts.some((x) => x.toUpperCase() === fc.toUpperCase())) fonts.push(fc); } else autoFont = true;
    }
    const swatch = (c: string) => h("span", { class: "color-swatch", style: `background:${c}` });
    const colorMenu = (kind: "filter" | "sort", at: DOMRect) => {
      const entries: MenuEntry[] = [];
      const pick = (color: { fill?: string | null; font?: string | null }) => () => {
        close();
        if (kind === "filter") setColumn(col, { color });
        else sortByColor(body, col, color);
      };
      if (fills.length) {
        entries.push({ heading: kind === "filter" ? "Nach Zellenfarbe filtern" : "Nach Zellenfarbe sortieren" });
        for (const c of fills) entries.push({ label: "", preview: swatch(c), run: pick({ fill: c }) });
        if (noFill) entries.push({ label: "Keine Füllung", run: pick({ fill: null }) });
      }
      if (fonts.length) {
        entries.push({ heading: kind === "filter" ? "Nach Schriftfarbe filtern" : "Nach Schriftfarbe sortieren" });
        for (const c of fonts) entries.push({ label: "", preview: swatch(c), run: pick({ font: c }) });
        if (autoFont) entries.push({ label: "Automatisch", run: pick({ font: null }) });
      }
      if (kind === "sort") entries.push("-", { label: "Benutzerdefiniertes Sortieren...", run: () => { close(); void app.commands.run("data.sort"); } });
      showMenu(app.commands, entries, new DOMRect(at.right, at.top, 0, at.height));
    };
    const colorItem = (label: string, kind: "filter" | "sort") => h("button", { type: "button", class: "filter-item sub", disabled: fills.length || fonts.length ? undefined : true,
      onclick: (e: MouseEvent) => colorMenu(kind, (e.currentTarget as HTMLElement).getBoundingClientRect()) }, h("span", { class: "fi-icon" }), label, h("span", { class: "fi-arrow" }, "›"));
    const list = [...values].sort((a, b) => a.localeCompare(b, "de", { numeric: true }));
    const boxes = list.map((v) => checkbox(v, !crit?.values || crit.values.includes(v)));
    const blankBox = blank ? checkbox("(Leere)", !crit?.values || !!crit.blank) : null;
    const allBox = checkbox("(Alles auswählen)", !crit?.values);
    const all = () => [...boxes, ...(blankBox ? [blankBox] : [])];
    allBox.input.onchange = () => { for (const b of all()) b.input.checked = allBox.input.checked; };
    for (const b of all()) b.input.onchange = () => { allBox.input.checked = all().every((x) => x.input.checked); allBox.input.indeterminate = !allBox.input.checked && all().some((x) => x.input.checked); };
    const search = h("input", { type: "search", class: "filter-search", placeholder: "Suchen" }) as HTMLInputElement;
    search.oninput = () => {
      const q = search.value.toLowerCase();
      list.forEach((v, i) => { boxes[i].el.hidden = !!q && !v.toLowerCase().includes(q); boxes[i].input.checked = !q || v.toLowerCase().includes(q); });
    };
    const panel = h("div", { class: "filter-panel" });
    const close = () => { panel.remove(); removeEventListener("pointerdown", outside, true); removeEventListener("keydown", keys, true); app.grid.focus(); };
    const outside = (e: Event) => { if (!panel.contains(e.target as Node)) close(); };
    const keys = (e: KeyboardEvent) => { if (e.key === "Escape") { e.preventDefault(); e.stopPropagation(); close(); } };
    const item = (label: string, run: (() => void) | null, icon = "") =>
      h("button", { type: "button", class: "filter-item", disabled: run ? undefined : true, onclick: () => { close(); run?.(); } }, h("span", { class: "fi-icon" }, icon), label);
    const body = { ...f.ref, r0: f.ref.r0 + 1 };
    const apply = () => {
      const chosen = list.filter((_, i) => boxes[i].input.checked);
      close();
      app.change("Filtern", () => {
        app.history.whole(sheet());
        if (chosen.length === list.length && (!blankBox || blankBox.input.checked)) delete f.columns[col - f.ref.c0];
        else f.columns[col - f.ref.c0] = { values: chosen, blank: blankBox?.input.checked };
        applyFilter();
      });
    };
    panel.append(
      item(numeric ? "Nach Größe sortieren (aufsteigend)" : "Von A bis Z sortieren", () => doSort(body, [{ col, asc: true }]), "↑"),
      item(numeric ? "Nach Größe sortieren (absteigend)" : "Von Z bis A sortieren", () => doSort(body, [{ col, asc: false }]), "↓"),
      colorItem("Nach Farbe sortieren", "sort"),
      h("hr"),
      item(`Filter entfernen aus "${name}"`, crit ? () => app.change("Filter löschen", () => { app.history.whole(sheet()); delete f.columns[col - f.ref.c0]; applyFilter(); }) : null),
      colorItem("Nach Farbe filtern", "filter"),
      h("button", { type: "button", class: "filter-item sub", onclick: (e: MouseEvent) => {
        // Excel's submenu: the comparisons open Benutzerdefinierter AutoFilter with the operator chosen.
        const rect = (e.currentTarget as HTMLElement).getBoundingClientRect();
        const entries: MenuEntry[] = (numeric ? NUMBER_FILTERS : TEXT_FILTERS).map((x) => x === "-" ? "-" : ({ label: x.label, run: () => {
          close();
          if (x.kind === "top10") void top10Dialog(col);
          else if (x.kind === "above" || x.kind === "below") setColumn(col, { dynamic: { type: x.kind === "above" ? "aboveAverage" : "belowAverage" } });
          else void customFilterDialog(col, name, x.op, x.op2);
        } }));
        showMenu(app.commands, entries, new DOMRect(rect.right, rect.top, 0, rect.height));
      } }, h("span", { class: "fi-icon" }), numeric ? "Zahlenfilter" : "Textfilter", h("span", { class: "fi-arrow" }, "›")),
      search,
      h("div", { class: "filter-list" }, allBox.el, ...boxes.map((b) => b.el), ...(blankBox ? [blankBox.el] : [])),
      h("div", { class: "filter-buttons" },
        h("button", { type: "button", class: "primary", onclick: apply }, "OK"),
        h("button", { type: "button", onclick: close }, "Abbrechen")));
    document.body.append(panel);
    const r = anchor ?? new DOMRect(200, 200, 0, 0);
    panel.style.left = `${Math.max(4, Math.min(innerWidth - 271, r.left))}px`;
    panel.style.top = `${Math.min(innerHeight - panel.offsetHeight - 4, r.bottom + 1)}px`;
    setTimeout(() => { addEventListener("pointerdown", outside, true); addEventListener("keydown", keys, true); });
    search.focus();
  }

  /** One column's filter set (Über/Unter dem Durchschnitt, the dialogs below). */
  function setColumn(col: number, crit: import("@raabe/sheet-engine").FilterColumn): void {
    const f = sheet().autoFilter;
    if (!f) return;
    app.change("Filtern", () => { app.history.whole(sheet()); f.columns[col - f.ref.c0] = crit; applyFilter(); });
  }

  /** Benutzerdefinierter AutoFilter: two conditions joined by Und or Oder. */
  async function customFilterDialog(col: number, name: string, op = "eq", op2?: string): Promise<void> {
    const f = sheet().autoFilter;
    if (!f) return;
    const now = f.columns[col - f.ref.c0]?.custom;
    const start = now ? now.items.map(fromCustom) : [{ op, val: "" }, ...(op2 ? [{ op: op2, val: "" }] : [])];
    const values = new Set<string>();
    for (let r = f.ref.r0 + 1; r <= f.ref.r1; r++) { const c = sheet().get(r, col); const t = shown(c, app.wb.styles.get(c?.s)).text; if (t) values.add(t); }
    const listId = `cf-values-${Date.now()}`;
    const row = (i: number) => {
      const sel = h("select", null, h("option", { value: "" }, ""), ...CUSTOM_OPS.map(([k, label]) => h("option", { value: k, selected: start[i]?.op === k || undefined }, label))) as HTMLSelectElement;
      const val = h("input", { type: "text", list: listId, value: start[i]?.val ?? "", style: "width:180px" }) as HTMLInputElement;
      return { sel, val, el: h("div", { class: "cf-row" }, sel, val) };
    };
    const a = row(0), b = row(1);
    const and = h("input", { type: "radio", name: "cf-join", checked: !now || now.and ? true : undefined }) as HTMLInputElement;
    const or = h("input", { type: "radio", name: "cf-join", checked: now && !now.and ? true : undefined }) as HTMLInputElement;
    const body = h("div", { class: "cf-dialog" }, h("div", { class: "fc-label" }, "Zeilen anzeigen:"), h("div", null, name), a.el,
      h("div", { class: "cf-join" }, h("label", { class: "check" }, and, "Und"), h("label", { class: "check" }, or, "Oder")), b.el,
      h("datalist", { id: listId }, ...[...values].sort((x, y) => x.localeCompare(y, "de", { numeric: true })).map((v) => h("option", { value: v }))),
      h("p", { class: "muted" }, "? steht für ein beliebiges Zeichen, * für beliebig viele Zeichen."));
    setTimeout(() => a.val.focus());
    const ok = await showDialog({ title: "Benutzerdefinierter AutoFilter", body, buttons: [{ label: "OK", value: true, primary: true }, { label: "Abbrechen", value: false }], cancel: false });
    if (!ok) return;
    const items = [a, b].filter((r) => r.sel.value).map((r) => toCustom(r.sel.value, r.val.value));
    if (!items.length) return;
    setColumn(col, { custom: { and: (items.length > 1 && and.checked) || undefined, items } });
  }

  /** Top 10 AutoFilter: Obersten/Untersten, a number, Elemente/Prozent. */
  async function top10Dialog(col: number): Promise<void> {
    const f = sheet().autoFilter;
    if (!f) return;
    const now = f.columns[col - f.ref.c0]?.top10;
    const which = h("select", null, h("option", { value: "top", selected: !now || now.top || undefined }, "Obersten"), h("option", { value: "bottom", selected: now && !now.top || undefined }, "Untersten")) as HTMLSelectElement;
    const n = h("input", { type: "number", min: "1", max: "500", value: String(now?.val ?? 10), style: "width:60px" }) as HTMLInputElement;
    const unit = h("select", null, h("option", { value: "items", selected: !now?.percent || undefined }, "Elemente"), h("option", { value: "percent", selected: now?.percent || undefined }, "Prozent")) as HTMLSelectElement;
    const body = h("div", null, h("div", { class: "fc-label" }, "Anzeigen"), h("div", { class: "cf-row" }, which, n, unit));
    const ok = await showDialog({ title: "Top 10 AutoFilter", body, buttons: [{ label: "OK", value: true, primary: true }, { label: "Abbrechen", value: false }], cancel: false });
    if (!ok) return;
    const val = Math.max(1, Math.min(500, Math.round(Number(n.value) || 10)));
    setColumn(col, { top10: { top: which.value === "top", percent: unit.value === "percent", val } });
  }

  // ----- data tools ----------------------------------------------------------------------------------

  async function textToColumns(): Promise<void> {
    const r = sel();
    const sepTab = checkbox("Tabstopp", true), sepSemi = checkbox("Semikolon"), sepComma = checkbox("Komma"), sepSpace = checkbox("Leerzeichen"), other = text("", 40);
    const body = h("div", null, h("p", null, "Trennzeichen:"), sepTab.el, sepSemi.el, sepComma.el, sepSpace.el, field("Andere:", other));
    const ok = await showDialog({ title: "Textkonvertierungs-Assistent", body, buttons: [{ label: "Fertig stellen", value: true, primary: true }, { label: "Abbrechen", value: false }], cancel: false });
    if (!ok) return;
    const seps = [sepTab.input.checked ? "\t" : "", sepSemi.input.checked ? ";" : "", sepComma.input.checked ? "," : "", sepSpace.input.checked ? " " : "", other.value].filter(Boolean);
    if (!seps.length) return;
    const re = new RegExp(seps.map((s) => s.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")).join("|"));
    app.change("Text in Spalten", () => {
      for (let row = r.r0; row <= r.r1; row++) {
        const c = sheet().get(row, r.c0);
        if (typeof c?.v !== "string") continue;
        c.v.split(re).forEach((part, j) => A.enter(app.ctx, sheet(), { r0: row, c0: r.c0 + j, r1: row, c1: r.c0 + j }, part));
      }
    });
  }

  async function removeDuplicates(): Promise<void> {
    const { rect, header } = sortRange();
    const hasHeader = checkbox("Daten haben Überschriften", header);
    const cols = range(rect.c0, rect.c1).map((c) => ({ c, box: checkbox(`Spalte ${colName(c)}`, true) }));
    const body = h("div", null, hasHeader.el, h("p", null, "Spalten:"), ...cols.map((x) => x.box.el));
    const ok = await showDialog({ title: "Duplikate entfernen", body, buttons: [{ label: "OK", value: true, primary: true }, { label: "Abbrechen", value: false }], cancel: false });
    if (!ok) return;
    const r0 = hasHeader.input.checked ? rect.r0 + 1 : rect.r0;
    const keyCols = cols.filter((x) => x.box.input.checked).map((x) => x.c);
    const seen = new Set<string>();
    const keep: Cell[][] = [];
    let removed = 0;
    for (let row = r0; row <= rect.r1; row++) {
      const key = keyCols.map((c) => String(sheet().get(row, c)?.v ?? "").toLowerCase()).join("\u0001");
      if (seen.has(key)) { removed++; continue; }
      seen.add(key);
      keep.push(range(rect.c0, rect.c1).map((c) => sheet().get(row, c) ?? { v: null }));
    }
    app.change("Duplikate entfernen", () => {
      app.history.cells(sheet(), { ...rect, r0 });
      for (let row = r0; row <= rect.r1; row++) for (const c of range(rect.c0, rect.c1)) sheet().set(row, c, undefined);
      keep.forEach((cells, i) => cells.forEach((c, j) => sheet().set(r0 + i, rect.c0 + j, c.v === null && c.s === undefined ? undefined : c)));
    });
    void messageBox("RaabeCalc", removed ? `${removed} doppelte Werte gefunden und entfernt; ${keep.length} eindeutige Werte bleiben erhalten.` : "Es wurden keine doppelten Werte gefunden.");
  }

  async function series(): Promise<void> {
    const r = sel();
    const inRows = radio("dir", "Zeilen", "rows", r.c1 - r.c0 > r.r1 - r.r0), inCols = radio("dir", "Spalten", "cols", r.c1 - r.c0 <= r.r1 - r.r0);
    const step = text("1", 80), stop = text("", 80);
    const body = h("div", null, h("p", null, "Reihe in:"), inRows, inCols, field("Inkrement:", step), field("Endwert:", stop));
    const ok = await showDialog({ title: "Reihe", body, buttons: [{ label: "OK", value: true, primary: true }, { label: "Abbrechen", value: false }], cancel: false });
    if (!ok) return;
    const inc = num(step.value) || 1, end = stop.value.trim() ? num(stop.value) : Infinity;
    const rows = picked(body, "dir") === "rows";
    app.change("Reihe", () => {
      app.history.cells(sheet(), r);
      const lanes = rows ? range(r.r0, r.r1) : range(r.c0, r.c1);
      for (const lane of lanes) {
        const start = rows ? sheet().get(lane, r.c0) : sheet().get(r.r0, lane);
        if (typeof start?.v !== "number") continue;
        let v = start.v;
        const cells = rows ? range(r.c0 + 1, r.c1).map((c) => ({ row: lane, col: c })) : range(r.r0 + 1, r.r1).map((rr) => ({ row: rr, col: lane }));
        for (const p of cells) {
          v = Number((v + inc).toPrecision(15));
          if (inc > 0 ? v > end : v < end) break;
          sheet().set(p.row, p.col, { v, s: start.s });
        }
      }
    });
  }

  // ----- tables and styles ---------------------------------------------------------------------------

  /** Als Tabelle formatieren: Excel's table styles; in a table it changes that table's style. */
  function tableStylesPanel(close: () => void): HTMLElement {
    return app.tables.stylesPanel(close, (name) => {
      if (app.tables.current()) app.tables.setStyle(name);
      else void formatAsTable(name);
    });
  }

  /** Als Tabelle formatieren / Einfügen > Tabelle (Strg+T): Tabelle erstellen, then a real Excel table;
   *  `direct` (Schnellanalyse > Tabelle) skips the dialog as Excel does. */
  async function formatAsTable(style?: string, direct = false): Promise<void> {
    const { rect, header } = sortRange();
    const ref = text(rangeName(rect).replace(/([A-Z]+)(\d+)/g, "$$$1$$$2"), 160);
    const hasHeader = checkbox("Tabelle hat Überschriften", header);
    const ok = direct || await showDialog({ title: "Tabelle erstellen", body: h("div", null, field("Wo sind die Daten für die Tabelle?", ref), hasHeader.el), buttons: [{ label: "OK", value: true, primary: true }, { label: "Abbrechen", value: false }], cancel: false });
    if (!ok) return;
    const r = parseRange(ref.value.replace(/\$/g, "")) ?? rect;
    app.tables.create(r, hasHeader.input.checked, style);
  }

  const CELL_STYLES: { name: string; fill?: string; color?: string; bold?: boolean; italic?: boolean; size?: number; border?: string }[] = [
    { name: "Gut", fill: "#C6EFCE", color: "#006100" }, { name: "Neutral", fill: "#FFEB9C", color: "#9C5700" }, { name: "Schlecht", fill: "#FFC7CE", color: "#9C0006" },
    { name: "Ausgabe", fill: "#F2F2F2", color: "#3F3F3F", bold: true }, { name: "Berechnung", fill: "#F2F2F2", color: "#FA7D00", bold: true }, { name: "Eingabe", fill: "#FFCC99", color: "#3F3F76" },
    { name: "Notiz", fill: "#FFFFCC" }, { name: "Warnender Text", color: "#FF0000" }, { name: "Erklärender Text", color: "#7F7F7F", italic: true },
    { name: "Überschrift", color: "#44546A", bold: true, size: 18 }, { name: "Überschrift 1", color: "#44546A", bold: true, size: 15, border: "#4472C4" },
    { name: "Überschrift 2", color: "#44546A", bold: true, size: 13, border: "#A2B8E1" }, { name: "Ergebnis", bold: true, border: "#4472C4" },
  ];

  function cellStylesPanel(close: () => void): HTMLElement {
    return h("div", { class: "tile-panel" }, h("div", { class: "tiles cellstyles" }, ...CELL_STYLES.map((s) =>
      h("button", {
        class: "cellstyle-tile", style: `background:${s.fill ?? "#fff"};color:${s.color ?? "#000"};font-weight:${s.bold ? 700 : 400};font-style:${s.italic ? "italic" : "normal"}`,
        onclick: () => {
          close();
          app.formatSel("Zellenformatvorlage", (st) => {
            st.fill = s.fill; st.font.color = s.color; st.font.bold = s.bold; st.font.italic = s.italic; st.styleName = s.name;
            if (s.size) st.font.size = s.size;
            if (s.border) st.borders = { ...st.borders, bottom: { style: s.name === "Ergebnis" ? "double" : "thick", color: s.border } };
          });
        },
      }, s.name))), h("button", { class: "cellstyle-reset", onclick: () => { close(); app.formatSel("Zellenformatvorlage", (st) => { st.fill = undefined; st.font = { name: "Calibri", size: 11, scheme: "minor" }; st.borders = {}; st.styleName = undefined; }); } }, "Standard"));
  }

  // ----- bedingte Formatierung -------------------------------------------------------------------------

  async function addRule(rule: Omit<CondRule, "priority">): Promise<void> {
    const ref = app.grid.selection.map((r) => ({ ...r }));
    app.change("Bedingte Formatierung", () => {
      app.history.whole(sheet());
      const list = sheet().conditional ?? (sheet().conditional = []);
      const prio = Math.max(0, ...list.flatMap((f) => f.rules.map((r) => r.priority))) + 1;
      list.push({ ref, rules: [{ ...rule, priority: prio }] });
    });
  }

  const PRESETS: Record<string, CondRule["format"]> = {
    "Hellrote Füllung 2 mit dunkelrotem Text": { fill: "#FFC7CE", color: "#9C0006" },
    "Gelbe Füllung mit dunkelgelbem Text": { fill: "#FFEB9C", color: "#9C5700" },
    "Grüne Füllung mit dunkelgrünem Text": { fill: "#C6EFCE", color: "#006100" },
    "Hellrote Füllung": { fill: "#FFC7CE" },
    "Roter Text": { color: "#9C0006" },
    "Roter Rahmen": { border: "#FF0000" },
  };

  async function ruleDialog(title: string, prompt: string, values: number, make: (v: string[], format: CondRule["format"]) => Omit<CondRule, "priority">): Promise<void> {
    const inputs = Array.from({ length: values }, () => text("", 120));
    const presetSel = h("select", null, ...Object.keys(PRESETS).map((k) => h("option", { value: k }, k))) as HTMLSelectElement;
    const body = h("div", null, h("p", null, prompt), h("div", { style: "display:flex;gap:8px;align-items:center" }, ...inputs.flatMap((x, i) => (i ? [h("span", null, "und"), x] : [x])), h("span", null, "mit"), presetSel));
    const ok = await showDialog({ title, body, buttons: [{ label: "OK", value: true, primary: true }, { label: "Abbrechen", value: false }], cancel: false });
    if (!ok) return;
    void addRule(make(inputs.map((x) => x.value.trim()), PRESETS[presetSel.value]));
  }

  const toFile = (v: string) => (/^-?[\d.]*,?\d+$/.test(v) ? v.replace(/\./g, "").replace(",", ".") : v.startsWith("=") ? v.slice(1) : `"${v.replace(/"/g, '""')}"`);

  function conditionalHighlightMenu(): MenuEntry[] {
    return [
      { label: "Größer als...", run: () => void ruleDialog("Größer als", "Zellen formatieren, die GRÖSSER SIND ALS:", 1, (v, f) => ({ type: "cellIs", operator: "greaterThan", formulas: [toFile(v[0])], format: f })) },
      { label: "Kleiner als...", run: () => void ruleDialog("Kleiner als", "Zellen formatieren, die KLEINER SIND ALS:", 1, (v, f) => ({ type: "cellIs", operator: "lessThan", formulas: [toFile(v[0])], format: f })) },
      { label: "Zwischen...", run: () => void ruleDialog("Zwischen", "Zellen formatieren, die ZWISCHEN folgenden Werten liegen:", 2, (v, f) => ({ type: "cellIs", operator: "between", formulas: [toFile(v[0]), toFile(v[1])], format: f })) },
      { label: "Gleich...", run: () => void ruleDialog("Gleich", "Zellen formatieren, die GLEICH SIND:", 1, (v, f) => ({ type: "cellIs", operator: "equal", formulas: [toFile(v[0])], format: f })) },
      { label: "Textinhalt...", run: () => void ruleDialog("Textinhalt", "Zellen formatieren, die folgenden Text enthalten:", 1, (v, f) => ({ type: "containsText", operator: "containsText", text: v[0], format: f })) },
      { label: "Datum...", run: () => void dateRuleDialog() },
      { label: "Doppelte Werte...", run: () => void duplicateDialog() },
    ];
  }

  /** The preset list of the rule dialogs ("Zu übernehmendes Format"). */
  const presetSelect = () => h("select", { "aria-label": "Zu übernehmendes Format" }, ...Object.keys(PRESETS).map((k) => h("option", { value: k }, k))) as HTMLSelectElement;
  const ruleOk = (title: string, body: HTMLElement) => showDialog({ title, body, buttons: [{ label: "OK", value: true, primary: true }, { label: "Abbrechen", value: false }], cancel: false });
  const row = (...kids: (HTMLElement | string)[]) => h("div", { style: "display:flex;gap:8px;align-items:center" }, ...kids);

  /** Obere/Untere 10 Elemente / %: how many (1 to 1000 items, 1 to 100 %), then the format (Excel's dialogs, D-143). */
  async function topDialog(bottom: boolean, percent: boolean): Promise<void> {
    const title = `${bottom ? "Untere" : "Obere"} 10 ${percent ? "%" : "Elemente"}`;
    const n = h("input", { type: "number", min: "1", max: percent ? "100" : "1000", value: "10", style: "width:64px", "aria-label": bottom ? "Zellen mit den niedrigsten Werten formatieren" : "Zellen mit den höchsten Werten formatieren" }) as HTMLInputElement;
    const preset = presetSelect();
    const body = h("div", null, h("p", null, bottom ? "Zellen mit den niedrigsten Werten formatieren:" : "Zellen mit den höchsten Werten formatieren:"), row(n, h("span", null, percent ? "% mit" : "mit"), preset));
    if (!await ruleOk(title, body)) return;
    const v = Math.round(num(n.value));
    if (!(v >= 1 && v <= (percent ? 100 : 1000))) return void messageBox("RaabeCalc", percent ? "Geben Sie eine ganze Zahl zwischen 1 und 100 ein." : "Geben Sie eine ganze Zahl zwischen 1 und 1000 ein.");
    void addRule({ type: "top10", rank: v, ...(percent ? { percent: true } : {}), ...(bottom ? { bottom: true } : {}), format: PRESETS[preset.value] });
  }

  async function averageDialog(above: boolean): Promise<void> {
    const preset = presetSelect();
    const body = h("div", null, h("p", null, above ? "Zellen formatieren, die ÜBER DEM DURCHSCHNITT liegen:" : "Zellen formatieren, die UNTER DEM DURCHSCHNITT liegen:"), row(h("span", null, "Für ausgewählten Bereich mit"), preset));
    if (!await ruleOk(above ? "Über dem Durchschnitt" : "Unter dem Durchschnitt", body)) return;
    void addRule({ type: "aboveAverage", above, format: PRESETS[preset.value] });
  }

  async function duplicateDialog(): Promise<void> {
    const kind = h("select", { "aria-label": "Zellen formatieren, die folgende Werte enthalten" }, h("option", { value: "dup" }, "Doppelte"), h("option", { value: "uni" }, "Eindeutige")) as HTMLSelectElement;
    const preset = presetSelect();
    const body = h("div", null, h("p", null, "Zellen formatieren, die folgende Werte enthalten:"), row(kind, h("span", null, "Werte mit"), preset));
    if (!await ruleOk("Doppelte Werte", body)) return;
    void addRule({ type: kind.value === "uni" ? "uniqueValues" : "duplicateValues", format: PRESETS[preset.value] });
  }

  /** Datum: the periods of Excel's list and the formula Excel writes for each (top-left cell `c`). */
  const PERIODS: [string, string, (c: string) => string][] = [
    ["yesterday", "Gestern", (c) => `FLOOR(${c},1)=TODAY()-1`],
    ["today", "Heute", (c) => `FLOOR(${c},1)=TODAY()`],
    ["tomorrow", "Morgen", (c) => `FLOOR(${c},1)=TODAY()+1`],
    ["last7Days", "In den letzten 7 Tagen", (c) => `AND(TODAY()-FLOOR(${c},1)<=6,FLOOR(${c},1)<=TODAY())`],
    ["lastWeek", "Letzte Woche", (c) => `AND(TODAY()-ROUNDDOWN(${c},0)>=(WEEKDAY(TODAY())),TODAY()-ROUNDDOWN(${c},0)<(WEEKDAY(TODAY())+7))`],
    ["thisWeek", "Diese Woche", (c) => `AND(TODAY()-ROUNDDOWN(${c},0)<=WEEKDAY(TODAY())-1,ROUNDDOWN(${c},0)-TODAY()<=7-WEEKDAY(TODAY()))`],
    ["nextWeek", "Nächste Woche", (c) => `AND(ROUNDDOWN(${c},0)-TODAY()>(7-WEEKDAY(TODAY())),ROUNDDOWN(${c},0)-TODAY()<(15-WEEKDAY(TODAY())))`],
    ["lastMonth", "Letzter Monat", (c) => `AND(MONTH(${c})=MONTH(EDATE(TODAY(),0-1)),YEAR(${c})=YEAR(EDATE(TODAY(),0-1)))`],
    ["thisMonth", "Dieser Monat", (c) => `AND(MONTH(${c})=MONTH(TODAY()),YEAR(${c})=YEAR(TODAY()))`],
    ["nextMonth", "Nächster Monat", (c) => `AND(MONTH(${c})=MONTH(EDATE(TODAY(),0+1)),YEAR(${c})=YEAR(EDATE(TODAY(),0+1)))`],
  ];

  async function dateRuleDialog(): Promise<void> {
    const period = h("select", { "aria-label": "Zellen mit Datum formatieren" }, ...PERIODS.map(([v, l]) => h("option", { value: v }, l))) as HTMLSelectElement;
    const preset = presetSelect();
    const body = h("div", null, h("p", null, "Zellen mit Datum formatieren:"), row(period, h("span", null, "mit"), preset));
    if (!await ruleOk("Datum", body)) return;
    const p = PERIODS.find((x) => x[0] === period.value)!;
    const tl = app.grid.selection[0];
    void addRule({ type: "timePeriod", timePeriod: p[0], formulas: [p[2](cellName(tl.r0, tl.c0))], format: PRESETS[preset.value] });
  }

  function conditionalTopMenu(): MenuEntry[] {
    return [
      { label: "Obere 10 Elemente...", run: () => void topDialog(false, false) },
      { label: "Obere 10 %...", run: () => void topDialog(false, true) },
      { label: "Untere 10 Elemente...", run: () => void topDialog(true, false) },
      { label: "Untere 10 %...", run: () => void topDialog(true, true) },
      { label: "Über dem Durchschnitt...", run: () => void averageDialog(true) },
      { label: "Unter dem Durchschnitt...", run: () => void averageDialog(false) },
    ];
  }

  function conditionalBarMenu(): MenuEntry[] {
    const bars: [string, string][] = [["Blaue Datenleiste", "#638EC6"], ["Grüne Datenleiste", "#63BE7B"], ["Rote Datenleiste", "#FF555A"], ["Orangefarbene Datenleiste", "#FFB628"], ["Hellblaue Datenleiste", "#008AEF"], ["Lila Datenleiste", "#D6007B"]];
    return bars.map(([l, c]) => ({ label: l, run: () => void addRule({ type: "dataBar", colors: [c] }) }));
  }

  function conditionalScaleMenu(): MenuEntry[] {
    const scales: [string, string[]][] = [
      ["Grün - Gelb - Rot", ["#F8696B", "#FFEB84", "#63BE7B"].reverse()], ["Rot - Gelb - Grün", ["#F8696B", "#FFEB84", "#63BE7B"]],
      ["Grün - Weiß - Rot", ["#63BE7B", "#FCFCFF", "#F8696B"]], ["Rot - Weiß - Grün", ["#F8696B", "#FCFCFF", "#63BE7B"]],
      ["Weiß - Rot", ["#FCFCFF", "#F8696B"]], ["Grün - Weiß", ["#63BE7B", "#FCFCFF"]],
    ];
    return scales.map(([l, c]) => ({ label: `Farbskala ${l}`, run: () => void addRule({ type: "colorScale", colors: c }) }));
  }

  /** Symbolsätze, in the gallery's groups (Richtungen, Formen, Indikatoren, Bewertungen). */
  function conditionalIconMenu(): MenuEntry[] {
    const out: MenuEntry[] = [];
    for (const [group, sets] of ICON_SETS) {
      out.push({ heading: group });
      for (const [label, set] of sets) out.push({ label, run: () => void addRule(iconRule(set)) });
    }
    return out;
  }

  function clearConditional(selection: boolean): void {
    app.change("Regeln löschen", () => {
      app.history.whole(sheet());
      if (!selection) { sheet().conditional = undefined; return; }
      sheet().conditional = (sheet().conditional ?? []).filter((f) => !f.ref.some((r) => app.grid.selection.some((s) => A.overlaps(r, s))));
    });
  }

  // ----- names, notes, links, symbols, functions ----------------------------------------------------------

  async function defineName(): Promise<void> {
    const r = sel();
    const above = sheet().get(r.r0, r.c0 - 1);
    const name = text(typeof above?.v === "string" ? above.v.replace(/\s+/g, "_") : "", 220);
    const ref = text(`=${quoteSheet(sheet().name)}!${rangeName(r).replace(/([A-Z]+)(\d+)/g, "$$$1$$$2")}`, 260);
    const ok = await showDialog({ title: "Neuer Name", body: h("div", null, field("Name:", name), field("Bezieht sich auf:", ref)), buttons: [{ label: "OK", value: true, primary: true }, { label: "Abbrechen", value: false }], cancel: false });
    if (!ok) return;
    const n = name.value.trim();
    if (!/^[A-Za-zÄÖÜäöüß_\\][A-Za-z0-9ÄÖÜäöüß_.]*$/.test(n) || /^[A-Za-z]{1,3}\d+$/.test(n)) return void messageBox("RaabeCalc", "Der eingegebene Name ist ungültig.");
    app.change("Namen definieren", () => {
      app.history.sheets();
      app.wb.names = [...app.wb.names.filter((x) => x.name.toLowerCase() !== n.toLowerCase()), { name: n, ref: ref.value.replace(/^=/, "") }];
    });
  }

  async function nameManager(): Promise<void> {
    const list = h("table", { class: "props" }, h("tr", null, h("th", null, "Name"), h("th", null, "Bezieht sich auf"), h("th", null, "")),
      ...app.wb.names.filter((n) => !n.name.startsWith("_xlnm")).map((n) => h("tr", null, h("td", null, n.name), h("td", null, "=" + n.ref),
        h("td", null, h("button", { type: "button", onclick: (e: MouseEvent) => {
          app.change("Name löschen", () => { app.history.sheets(); app.wb.names = app.wb.names.filter((x) => x !== n); });
          (e.target as HTMLElement).closest("tr")?.remove();
        } }, "Löschen")))));
    await showDialog({ title: "Namens-Manager", body: h("div", null, h("button", { type: "button", onclick: () => void defineName() }, "Neu..."), list), buttons: [{ label: "Schließen", value: undefined }], cancel: undefined });
  }

  async function note(): Promise<void> {
    const a = app.sheet.view.active;
    const c = sheet().get(a.row, a.col);
    const area = h("textarea", { rows: 5, style: "width:300px;font:13px var(--font)" }, c?.note ?? "") as HTMLTextAreaElement;
    const ok = await showDialog({ title: `Notiz: ${cellName(a.row, a.col)}`, body: area, buttons: [{ label: "OK", value: true, primary: true }, { label: "Abbrechen", value: false }], cancel: false });
    if (!ok) return;
    app.change("Notiz", () => {
      app.history.cell(sheet(), a.row, a.col);
      sheet().set(a.row, a.col, { ...(c ?? { v: null }), note: area.value || undefined });
    });
  }

  async function link(): Promise<void> {
    const a = app.sheet.view.active;
    const c = sheet().get(a.row, a.col);
    const shownText = text(c && typeof c.v === "string" ? c.v : "", 280), address = text(c?.link ?? "", 280);
    const ok = await showDialog({ title: "Link einfügen", body: h("div", null, field("Anzuzeigender Text:", shownText), field("Adresse:", address)), buttons: [{ label: "OK", value: true, primary: true }, { label: "Abbrechen", value: false }], cancel: false });
    if (!ok || !address.value.trim()) return;
    app.change("Link", () => {
      app.history.cell(sheet(), a.row, a.col);
      const s = app.wb.styles.derive(c?.s, (st) => { st.font.color = "#0563C1"; st.font.theme = { t: 10 }; st.font.underline = "single"; });
      sheet().set(a.row, a.col, { ...(c ?? { v: null }), v: shownText.value || address.value, link: address.value.trim(), s });
    });
  }

  async function symbol(): Promise<void> {
    const chars = "€£¥©®™°±×÷≠≤≥≈∞√∑∏πΩαβγδΔλμσφ½¼¾²³¹←→↑↓↔•...–—‰§¶✓✗".split("");
    let chosen = "";
    const grid = h("div", { class: "symbols" }, ...chars.map((ch) => h("button", { type: "button", onclick: () => { chosen = ch; for (const b of grid.children) b.classList.toggle("on", b.textContent === ch); } }, ch)));
    const ok = await showDialog({ title: "Symbol", body: grid, buttons: [{ label: "Einfügen", value: true, primary: true }, { label: "Abbrechen", value: false }], cancel: false });
    if (!ok || !chosen) return;
    if (app.grid.isEditing) {
      const input = document.querySelector<HTMLTextAreaElement>(".cell-editor.on")!;
      const s = input.selectionStart ?? input.value.length;
      input.value = input.value.slice(0, s) + chosen + input.value.slice(s);
      input.dispatchEvent(new Event("input"));
    } else app.grid.startEdit("edit", editText(sheet().get(app.sheet.view.active.row, app.sheet.view.active.col), app.activeStyle()) + chosen);
  }

  async function insertFunction(): Promise<void> {
    const search = text("", 260);
    const list = h("select", { size: 12, style: "width:300px" }) as HTMLSelectElement;
    const hint = h("div", { class: "muted", style: "min-height:2em" });
    const fill = () => {
      const q = search.value.trim().toUpperCase();
      list.replaceChildren(...FUNCTIONS_DE.filter((f) => !q || f.name.includes(q)).map((f) => h("option", { value: f.name }, f.name)));
      list.selectedIndex = 0;
      list.dispatchEvent(new Event("change"));
    };
    list.onchange = () => { const f = FUNCTIONS_DE.find((x) => x.name === list.value); hint.textContent = f ? `${f.name}(${f.args})` : ""; };
    search.oninput = fill;
    fill();
    const ok = await showDialog({ title: "Funktion einfügen", body: h("div", null, field("Funktion suchen:", search), field("Funktion auswählen:", list), hint), buttons: [{ label: "OK", value: true, primary: true }, { label: "Abbrechen", value: false }], cancel: false });
    if (!ok || !list.value) return;
    if (app.grid.isEditing) {
      const input = document.querySelector<HTMLTextAreaElement>(".cell-editor.on") ?? document.querySelector<HTMLTextAreaElement>(".fx-input")!;
      const s = input.selectionStart ?? input.value.length;
      input.value = input.value.slice(0, s) + list.value + "(" + input.value.slice(s);
      input.dispatchEvent(new Event("input"));
    } else app.grid.startEdit("enter", `=${list.value}(`);
  }

  // ----- sheets, zoom, statistics, options, paste special ---------------------------------------------------

  async function moveSheet(): Promise<void> {
    const before = h("select", { size: 8, style: "width:240px" }, ...app.wb.sheets.map((s, i) => h("option", { value: String(i) }, s.name)), h("option", { value: String(app.wb.sheets.length) }, "(ans Ende stellen)")) as HTMLSelectElement;
    before.value = String(app.wb.active);
    const copy = checkbox("Kopie erstellen");
    const ok = await showDialog({ title: "Verschieben oder kopieren", body: h("div", null, field("Einfügen vor:", before), copy.el), buttons: [{ label: "OK", value: true, primary: true }, { label: "Abbrechen", value: false }], cancel: false });
    if (!ok) return;
    let to = Number(before.value);
    if (!copy.input.checked && to > app.wb.active) to--;
    (app as unknown as { moveSheet(f: number, t: number, c: boolean): void }).moveSheet(app.wb.active, to, copy.input.checked);
  }

  async function unhideSheet(): Promise<void> {
    const list = h("select", { size: 6, style: "width:240px" }, ...app.wb.sheets.map((s, i) => (s.hidden ? h("option", { value: String(i) }, s.name) : null)).filter(Boolean) as HTMLElement[]) as HTMLSelectElement;
    list.selectedIndex = 0;
    const ok = await showDialog({ title: "Einblenden", body: field("Blatt einblenden:", list), buttons: [{ label: "OK", value: true, primary: true }, { label: "Abbrechen", value: false }], cancel: false });
    if (!ok || !list.value) return;
    const i = Number(list.value);
    app.change("Einblenden", () => { app.history.sheets(); app.wb.sheets[i].hidden = undefined; app.wb.active = i; });
  }

  async function zoom(): Promise<void> {
    const cur = app.sheet.view.zoom;
    const body = h("div", null, ...[200, 100, 75, 50, 25].map((z) => radio("z", `${z} %`, String(z), z === cur)), radio("z", "Markierung", "fit"),
      h("label", { class: "check" }, h("input", { type: "radio", name: "z", value: "custom", checked: [200, 100, 75, 50, 25].includes(cur) ? undefined : true }), "Benutzerdefiniert:", text(String(cur), 50)));
    const ok = await showDialog({ title: "Zoom", body, buttons: [{ label: "OK", value: true, primary: true }, { label: "Abbrechen", value: false }], cancel: false });
    if (!ok) return;
    const v = picked(body, "z");
    if (v === "fit") return void app.commands.run("view.zoomSelection");
    const z = v === "custom" ? num(body.querySelector<HTMLInputElement>("input[type=text]")!.value) : Number(v);
    if (z >= 10 && z <= 400) app.setZoom(z);
  }

  async function statistics(): Promise<void> {
    let cells = 0, formulas = 0, notes = 0;
    for (const c of sheet().cells.values()) { if (c.v !== null) cells++; if (c.f) formulas++; if (c.note) notes++; }
    const used = sheet().usedRange();
    const t = h("table", { class: "props" },
      h("tr", null, h("th", { colspan: "2" }, "Aktuelles Blatt")),
      h("tr", null, h("td", null, "Ende des Blatts"), h("td", null, used ? cellName(used.r1, used.c1) : "A1")),
      h("tr", null, h("td", null, "Zellen mit Daten"), h("td", null, String(cells))),
      h("tr", null, h("td", null, "Formeln"), h("td", null, String(formulas))),
      h("tr", null, h("td", null, "Notizen"), h("td", null, String(notes))),
      h("tr", null, h("th", { colspan: "2" }, "Arbeitsmappe")),
      h("tr", null, h("td", null, "Blätter"), h("td", null, String(app.wb.sheets.length))),
      h("tr", null, h("td", null, "Namen"), h("td", null, String(app.wb.names.length))));
    await showDialog({ title: "Arbeitsmappenstatistiken", body: t, buttons: [{ label: "OK", value: undefined, primary: true }], cancel: undefined });
  }

  async function options(settings: Settings, save: (s: Settings) => void): Promise<void> {
    const user = text(settings.userName, 240), auto = text(String(settings.autoRecoverSeconds), 60), touch = checkbox("Fingereingabemodus (größere Schaltflächen)", settings.touchMode);
    const ok = await showDialog({ title: "Optionen", body: h("div", null, field("Benutzername:", user), field("AutoWiederherstellen-Informationen speichern alle (Sekunden):", auto), touch.el), buttons: [{ label: "OK", value: true, primary: true }, { label: "Abbrechen", value: false }], cancel: false });
    if (ok) save({ userName: user.value.trim(), autoRecoverSeconds: Math.max(10, Number(auto.value) || 30), touchMode: touch.input.checked });
  }

  async function pasteSpecial(): Promise<void> {
    const body = h("div", null, h("p", null, "Einfügen"), radio("ps", "Alles", "all", true), radio("ps", "Formeln", "formulas"), radio("ps", "Werte", "values"), radio("ps", "Formate", "formats"), h("hr"), checkbox("Transponieren").el);
    const ok = await showDialog({ title: "Inhalte einfügen", body, buttons: [{ label: "OK", value: true, primary: true }, { label: "Abbrechen", value: false }], cancel: false });
    if (!ok) return;
    const transpose = body.querySelector<HTMLInputElement>("input[type=checkbox]")!.checked;
    void app.paste(transpose ? "transpose" : (picked(body, "ps") as "all" | "values" | "formats" | "formulas"));
  }

  return {
    colWidth, defaultWidth, rowHeight, insertCells, deleteCells, formatCells, find, goto, sortQuick, sort, hasFilter, toggleFilter, clearFilter,
    reapplyFilter, filterByActive, filterMenu, textToColumns, removeDuplicates, series, tableStylesPanel, formatAsTable, cellStylesPanel,
    addRule, conditionalIconMenu, conditionalHighlightMenu, conditionalTopMenu, conditionalBarMenu, conditionalScaleMenu, clearConditional, defineName, nameManager, note, link,
    symbol, insertFunction, moveSheet, unhideSheet, zoom, statistics, options, pasteSpecial,
  };
}

function range(a: number, b: number): number[] {
  const out: number[] = [];
  for (let i = a; i <= Math.min(b, a + 20000); i++) out.push(i);
  return out;
}

function quoteSheet(n: string): string {
  return /^[A-Za-z_][A-Za-z0-9_.]*$/.test(n) ? n : `'${n.replace(/'/g, "''")}'`;
}

void MAX_ROWS;

/** Excel's icon sets by gallery group (recorded: probe_excel_iconsets.py): German name and the file's name. */
export const ICON_SETS: [string, [string, string][]][] = [
  ["Richtungen", [["3 Pfeile (farbig)", "3Arrows"], ["3 Pfeile (grau)", "3ArrowsGray"], ["3 Dreiecke", "3Triangles"], ["4 Pfeile (grau)", "4ArrowsGray"], ["4 Pfeile (farbig)", "4Arrows"], ["5 Pfeile (grau)", "5ArrowsGray"], ["5 Pfeile (farbig)", "5Arrows"]]],
  ["Formen", [["3 Ampeln (ohne Rand)", "3TrafficLights1"], ["3 Ampeln (mit Rand)", "3TrafficLights2"], ["3 Zeichen", "3Signs"], ["4 Ampeln", "4TrafficLights"], ["Rot/Schwarz", "4RedToBlack"]]],
  ["Indikatoren", [["3 Symbole (mit Kreis)", "3Symbols"], ["3 Symbole (ohne Kreis)", "3Symbols2"], ["3 Kennzeichen", "3Flags"]]],
  ["Bewertungen", [["3 Sterne", "3Stars"], ["4 Bewertungen", "4Rating"], ["5 Viertel", "5Quarters"], ["5 Bewertungen", "5Rating"], ["5 Kästchen", "5Boxes"]]],
];

/** A Symbolsatz rule with Excel's default limits (equal shares of the range in percent). */
export function iconRule(set: string): Omit<CondRule, "priority"> {
  const n = Number(set[0]) || 3;
  const cfvo = Array.from({ length: n }, (_, i) => ({ type: "percent", val: String(Math.round((100 * i) / n)) }));
  return { type: "iconSet", iconSet: { set, cfvo } };
}
