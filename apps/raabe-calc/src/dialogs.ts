// RaabeCalc's dialogs and the menus behind them, with Excel's German titles and fields:
// Zellen formatieren (Strg+1), Spaltenbreite, Zeilenhöhe, Zellen einfügen/löschen, Gehe zu,
// Suchen und Ersetzen, Sortieren, AutoFilter, Text in Spalten, Duplikate entfernen, Namen,
// Notizen, Funktion einfügen, Zoom, Inhalte einfügen, bedingte Formatierung.

import {
  MAX_COLS, MAX_ROWS, cellName, colName, FUNCTIONS_DE, FORMAT_CATEGORIES, localFormat, fileFormat, parseRange, rangeName, shiftFormula,
  type Cell, type CondRule, type Rect, type Sheet,
} from "@raabe/sheet-engine";
import { colorGrid, h, messageBox, pressOutside, showDialog, showMenu, type MenuEntry } from "@raabe/shell";
import type { CalcApp } from "./app";
import * as A from "./actions";
import { shown, editText } from "./engine";
import { DEFAULT_COL_CHARS, charsToPx, shownChars, shownToChars } from "./grid/metrics";
import type { Settings } from "./store";
import { condRules, hexColor, textRule } from "./condrules";
import { PATTERNS, SHADINGS, cssFill, fillXmlOf, parseFill, shadingFill, type FancyFill, type GradientFill, type Shading } from "./fills";
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
    // Dezimalstellen and 1.000er-Trennzeichen of the number categories (Excel's Zahlen tab, D-142): they
    // rewrite the code picked in the list.
    const DEC_CATS = ["Zahl", "Währung", "Buchhaltung", "Prozent", "Wissenschaft"];
    const decimals = h("input", { type: "number", min: "0", max: "30", value: "2", style: "width:60px", "aria-label": "Dezimalstellen:" }) as HTMLInputElement;
    const thousands = checkbox("1.000er-Trennzeichen verwenden (.)");
    const decRow = h("div", null, field("Dezimalstellen:", decimals));
    const applyDigits = (code: string) => {
      const n = Math.max(0, Math.min(30, Math.round(num(decimals.value) || 0)));
      // The first number of each section only (not the exponent of 0,00E+00).
      let out = code.split(";").map((sec) => sec.replace(/((?:#,##)?0)(\.0+)?/, (_m, i: string) => i + (n ? "." + "0".repeat(n) : ""))).join(";").replace(/"-"\?+/g, `"-"${"?".repeat(n)}`);
      if (catList.value === "Zahl") out = thousands.input.checked ? out.replace(/(^|[^#,0.])0(?=[.;_ \\]|$)/g, "$1#,##0") : out.replace(/#,##0/g, "0");
      return out;
    };
    const showDigits = () => {
      decRow.hidden = !DEC_CATS.includes(catList.value);
      thousands.el.hidden = catList.value !== "Zahl";
    };
    const digitsChanged = () => { if (codeList.value) { codeBox.value = localFormat(applyDigits(codeList.value)); showSample(); } };
    decimals.oninput = digitsChanged;
    thousands.input.onchange = digitsChanged;
    const fillCodes = () => {
      const c = cats.find((x) => x.name === catList.value);
      codeList.replaceChildren(...(c?.codes ?? []).map((f) => h("option", { value: f }, localFormat(f))));
      if (c && c.codes.length) {
        const cur = c.codes.find((f) => f === st.numFmt) ?? c.codes[0];
        codeList.value = cur;
        codeBox.value = localFormat(cur);
        const m = /0\.(0+)/.exec(cur.split(";")[0]);
        decimals.value = String(m ? m[1].length : 0);
        thousands.input.checked = cur.includes("#,##0");
      }
      if (c?.name === "Standard") codeBox.value = "Standard";
      if (c?.name === "Text") codeBox.value = "@";
      showDigits();
      showSample();
    };
    // A code that is a category's code with other decimals (0,000) belongs to that category.
    let curCat = cats.find((c) => c.codes.includes(st.numFmt))?.name ?? (st.numFmt === "General" ? "Standard" : "Benutzerdefiniert");
    let curBase: string | null = null;
    if (curCat === "Benutzerdefiniert") {
      const m = /0\.(0+)/.exec(st.numFmt.split(";")[0]);
      decimals.value = String(m ? m[1].length : 0);
      for (const c of cats.filter((x) => DEC_CATS.includes(x.name))) {
        catList.value = c.name;
        for (const th of [false, true]) {
          thousands.input.checked = th;
          const hit = c.codes.find((f) => applyDigits(f) === st.numFmt);
          if (hit && !curBase) { curCat = c.name; curBase = hit; }
        }
      }
    }
    catList.value = curCat;
    catList.onchange = fillCodes;
    codeList.onchange = () => { codeBox.value = localFormat(DEC_CATS.includes(catList.value) ? applyDigits(codeList.value) : codeList.value); showSample(); };
    codeBox.oninput = showSample;
    if (curBase) {
      const keep = decimals.value, th = st.numFmt.includes("#,##0");
      fillCodes();
      codeList.value = curBase;
      decimals.value = keep;
      thousands.input.checked = th;
      codeBox.value = localFormat(st.numFmt);
      showSample();
    } else if (curCat !== "Benutzerdefiniert") fillCodes();
    else { codeList.replaceChildren(...(cats.find((c) => c.name === "Benutzerdefiniert")?.codes ?? []).map((f) => h("option", { value: f }, localFormat(f)))); showDigits(); showSample(); }
    panels.Zahlen = h("div", { class: "fc-num" }, h("div", null, h("div", { class: "fc-label" }, "Kategorie:"), catList),
      h("div", { class: "fc-right" }, h("div", { class: "fc-label" }, "Beispiel"), preview, decRow, thousands.el, h("div", { class: "fc-label" }, "Typ:"), codeBox, codeList));

    // Ausrichtung: Textausrichtung, Textsteuerung and Excel's orientation box: the vertical "Text" (Vertikaler Text)
    // and the dial whose pointer is dragged between -90 and 90 degrees, with the Grad spinner (D-148).
    const hSel = h("select", { "aria-label": "Horizontal:" }, ...[["general", "Standard"], ["left", "Links (Einzug)"], ["center", "Zentriert"], ["right", "Rechts (Einzug)"], ["fill", "Ausfüllen"], ["justify", "Blocksatz"], ["centerContinuous", "Über Auswahl zentrieren"], ["distributed", "Verteilt (Einzug)"]]
      .map(([v, l]) => h("option", { value: v, selected: (st.h ?? "general") === v || undefined }, l))) as HTMLSelectElement;
    const vSel = h("select", { "aria-label": "Vertikal:" }, ...[["top", "Oben"], ["center", "Zentriert"], ["bottom", "Unten"], ["justify", "Blocksatz"], ["distributed", "Verteilt"]]
      .map(([v, l]) => h("option", { value: v, selected: (st.v ?? "bottom") === v || undefined }, l))) as HTMLSelectElement;
    const indent = h("input", { type: "number", min: "0", max: "250", value: String(st.indent ?? 0), style: "width:60px", "aria-label": "Einzug:" }) as HTMLInputElement;
    let vertical = st.rotation === 255;
    let deg = st.rotation === undefined || st.rotation === 255 ? 0 : st.rotation > 90 && st.rotation <= 180 ? 90 - st.rotation : st.rotation;
    const rot = h("input", { type: "number", min: "-90", max: "90", value: String(deg), style: "width:56px", "aria-label": "Grad" }) as HTMLInputElement;
    const vbox = h("button", { type: "button", class: "fc-vtext", role: "checkbox", "aria-label": "Vertikaler Text" }, ...[..."Text"].map((ch) => h("span", null, ch)));
    const dial = h("div", { class: "fc-dial", role: "slider", tabindex: "0", "aria-label": "Ausrichtung", "aria-valuemin": "-90", "aria-valuemax": "90" });
    const CX = 14, CY = 74, R = 60;
    const paintDial = () => {
      const a = (deg * Math.PI) / 180, x = CX + R * Math.cos(a), y = CY - R * Math.sin(a);
      const dots = range(-6, 6).map((k) => { const b = (k * 15 * Math.PI) / 180, big = k % 3 === 0; return `<circle cx="${CX + R * Math.cos(b)}" cy="${CY - R * Math.sin(b)}" r="${big ? 2.6 : 1.5}" fill="${big ? "#000" : "#555"}"/>`; }).join("");
      const tx = CX + 22 * Math.cos(a), ty = CY - 22 * Math.sin(a);
      dial.innerHTML = `<svg width="90" height="148" viewBox="0 0 90 148" aria-hidden="true">${dots}<line x1="${CX + 44 * Math.cos(a)}" y1="${CY - 44 * Math.sin(a)}" x2="${x}" y2="${y}" stroke="#c00000" stroke-width="1.5"/>` +
        `<rect x="${x - 3}" y="${y - 3}" width="6" height="6" fill="#c00000"/><text x="${tx}" y="${ty}" font-size="12" dominant-baseline="middle" transform="rotate(${-deg} ${tx} ${ty})">Text</text></svg>`;
      dial.setAttribute("aria-valuenow", String(deg));
      dial.classList.toggle("off", vertical);
      vbox.setAttribute("aria-checked", String(vertical));
      vbox.classList.toggle("on", vertical);
      rot.disabled = vertical;
      if (Number(rot.value) !== deg) rot.value = String(deg);
    };
    const setDeg = (d: number) => { if (vertical) return; deg = Math.max(-90, Math.min(90, Math.round(d))); paintDial(); };
    const fromPointer = (e: PointerEvent) => {
      const r = dial.getBoundingClientRect();
      setDeg((Math.atan2(r.top + CY - e.clientY, e.clientX - (r.left + CX)) * 180) / Math.PI);
    };
    dial.addEventListener("pointerdown", (e) => { if (vertical) return; dial.setPointerCapture(e.pointerId); fromPointer(e); });
    dial.addEventListener("pointermove", (e) => { if (dial.hasPointerCapture(e.pointerId)) fromPointer(e); });
    dial.addEventListener("keydown", (e) => { if (e.key === "ArrowUp" || e.key === "ArrowRight") { setDeg(deg + 1); e.preventDefault(); } else if (e.key === "ArrowDown" || e.key === "ArrowLeft") { setDeg(deg - 1); e.preventDefault(); } });
    rot.oninput = () => { const v = Number(rot.value); if (Number.isFinite(v) && rot.value.trim() !== "" && rot.value !== "-") { deg = Math.max(-90, Math.min(90, Math.round(v))); paintDial(); } };
    vbox.onclick = () => { vertical = !vertical; paintDial(); };
    paintDial();
    const wrap = checkbox("Zeilenumbruch", !!st.wrap), shrink = checkbox("An Zellgröße anpassen", !!st.shrink), mergeBox = checkbox("Zellen verbinden", app.sheet.merges.some((m) => A.overlaps(m, sel())));
    panels.Ausrichtung = h("div", { class: "fc-align" },
      h("div", null, h("div", { class: "fc-label" }, "Textausrichtung"), field("Horizontal:", hSel), field("Vertikal:", vSel), field("Einzug:", indent),
        h("div", { class: "fc-label" }, "Textsteuerung"), wrap.el, shrink.el, mergeBox.el),
      h("fieldset", { class: "fc-orient" }, h("legend", null, "Ausrichtung"), h("div", { class: "fc-orient-box" }, vbox, dial), h("label", { class: "field" }, rot, "Grad")));

    // Schrift: Schriftart, Schriftschnitt and Schriftgrad each a field with its list, Unterstreichung, Farbe,
    // Standardschrift, Effekte and Vorschau, as Excel's page (D-148).
    const strip = (n: string) => n.replace(/ \((Überschriften|Textkörper)\)$/, "");
    const themed = [`${app.wb.themeFonts.major} (Überschriften)`, `${app.wb.themeFonts.minor} (Textkörper)`];
    const all = [...new Set(app.fonts().length ? app.fonts() : ["Arial", "Calibri", "Calibri Light", "Cambria", "Times New Roman"])].sort((a, b) => a.localeCompare(b, "de"));
    const pair = (label: string, value: string, items: string[], width: number) => {
      const input = h("input", { type: "text", value, "aria-label": label, style: `width:${width}px` }) as HTMLInputElement;
      const list = h("select", { size: 6, class: "fc-list", "aria-label": label, style: `width:${width + 6}px` }, ...items.map((v) => h("option", { value: v }, v))) as HTMLSelectElement;
      const sync = () => { const t = input.value.trim().toLowerCase(); const i = items.findIndex((v) => strip(v).toLowerCase() === t) ; const j = i >= 0 ? i : items.findIndex((v) => v.toLowerCase().startsWith(t)); list.selectedIndex = t ? j : -1; };
      list.onchange = () => { input.value = label === "Schriftart:" ? strip(list.value) : list.value; input.dispatchEvent(new Event("input")); };
      input.addEventListener("input", sync);
      sync();
      return { input, list, el: h("div", { class: "fc-pair" }, h("div", null, label), input, list) };
    };
    const fStyleName = st.font.bold && st.font.italic ? "Fett Kursiv" : st.font.bold ? "Fett" : st.font.italic ? "Kursiv" : "Standard";
    const fontP = pair("Schriftart:", st.font.name, [...themed, ...all], 200);
    const styleP = pair("Schriftschnitt:", fStyleName, ["Standard", "Kursiv", "Fett", "Fett Kursiv"], 110);
    const sizeP = pair("Schriftgrad:", de(st.font.size), ["8", "9", "10", "11", "12", "14", "16", "18", "20", "22", "24", "26", "28", "36", "48", "72"], 56);
    const fName = fontP.input, fStyle = styleP.input, fSize = sizeP.input;
    const fUnder = h("select", { "aria-label": "Unterstreichung:" }, ...[["", "Ohne"], ["single", "Einfach"], ["double", "Doppelt"], ["singleAccounting", "Einfach (Buchhaltung)"], ["doubleAccounting", "Doppelt (Buchhaltung)"]]
      .map(([v, l]) => h("option", { value: v, selected: (st.font.underline ?? "") === v || undefined }, l))) as HTMLSelectElement;
    let fColor = st.font.color ?? "";
    const fColorBtn = h("button", { type: "button", class: "fc-color", "aria-label": "Farbe:", style: `background:${fColor || "#000"}` });
    const strike = checkbox("Durchgestrichen", !!st.font.strike), sup = checkbox("Hochgestellt", st.font.vertAlign === "superscript"), sub = checkbox("Tiefgestellt", st.font.vertAlign === "subscript");
    const normal = checkbox("Standardschrift");
    const fPreview = h("div", { class: "fc-font-preview" }, "AaBbCcYyZz");
    const paintFont = () => {
      const size = num(fSize.value);
      fPreview.style.fontFamily = `"${strip(fName.value.trim()) || st.font.name}", Carlito, sans-serif`;
      fPreview.style.fontSize = `${Math.min(48, size >= 1 ? size : st.font.size) * 96 / 72}px`;
      fPreview.style.fontWeight = /fett/i.test(fStyle.value) ? "700" : "400";
      fPreview.style.fontStyle = /kursiv/i.test(fStyle.value) ? "italic" : "normal";
      fPreview.style.textDecoration = [fUnder.value ? "underline" : "", strike.input.checked ? "line-through" : ""].filter(Boolean).join(" ") || "none";
      fPreview.style.textDecorationStyle = fUnder.value.startsWith("double") ? "double" : "solid";
      fPreview.style.color = fColor || "#000";
      fPreview.style.verticalAlign = "baseline";
    };
    const fontChanged = () => { normal.input.checked = false; paintFont(); };
    for (const el of [fName, fStyle, fSize]) el.addEventListener("input", fontChanged);
    fUnder.onchange = fontChanged;
    strike.input.onchange = fontChanged;
    // Hochgestellt and Tiefgestellt exclude each other.
    sup.input.onchange = () => { if (sup.input.checked) sub.input.checked = false; fontChanged(); };
    sub.input.onchange = () => { if (sub.input.checked) sup.input.checked = false; fontChanged(); };
    fColorBtn.onclick = () => openColors(fColorBtn, (c) => { fColor = c; fColorBtn.style.background = c || "#000"; fontChanged(); });
    // Standardschrift: the workbook's standard font with no effects.
    normal.input.onchange = () => {
      if (!normal.input.checked) return;
      const d = app.wb.styles.get(undefined).font;
      fName.value = d.name; fStyle.value = "Standard"; fSize.value = de(d.size); fUnder.value = ""; fColor = ""; fColorBtn.style.background = "#000";
      strike.input.checked = sup.input.checked = sub.input.checked = false;
      for (const el of [fName, fStyle, fSize]) el.dispatchEvent(new Event("input"));
      normal.input.checked = true;
      paintFont();
    };
    paintFont();
    panels.Schrift = h("div", { class: "fc-font" },
      h("div", { class: "fc-pairs" }, fontP.el, styleP.el, sizeP.el),
      h("div", { class: "fc-font-row" }, field("Unterstreichung:", fUnder), field("Farbe:", fColorBtn), normal.el),
      h("div", { class: "fc-font-row" }, h("fieldset", null, h("legend", null, "Effekte"), strike.el, sup.el, sub.el), h("fieldset", { class: "fc-preview-box" }, h("legend", null, "Vorschau"), fPreview)));

    // Rahmen: Linie (Art and Farbe), Voreinstellungen (Keine, Außen, Innen) and the sides as buttons around a
    // preview that can be clicked too, both diagonals included (D-147). Only the sides touched in the dialog change.
    type SideKey = "top" | "insideH" | "bottom" | "diagUp" | "left" | "insideV" | "right" | "diagDown";
    const LINE_STYLES: [string, string][] = [["none", "Keine"], ["hair", "Haarstrich"], ["dotted", "Punktiert"], ["dashDotDot", "StrichPunktPunkt"], ["dashDot", "StrichPunkt"], ["dashed", "Gestrichelt"], ["thin", "Dünn"],
      ["mediumDashDotDot", "StrichPunktPunkt Mittel"], ["slantDashDot", "SchrägstrichStrichPunkt"], ["mediumDashDot", "StrichPunkt Mittel"], ["mediumDashed", "Gestrichelt Mittel"], ["medium", "Mittel"], ["thick", "Dick"], ["double", "Doppelt"]];
    const pen: { style: string; color?: string } = { style: "thin" };
    const many = app.grid.selection.some((r) => r.r1 > r.r0 || r.c1 > r.c0);
    const rowsMany = app.grid.selection.some((r) => r.r1 > r.r0), colsMany = app.grid.selection.some((r) => r.c1 > r.c0);
    const b0 = st.borders;
    const sides: Partial<Record<SideKey, { style: string; color?: string }>> = {
      top: b0.top, bottom: b0.bottom, left: b0.left, right: b0.right,
      diagUp: b0.diagonalUp ? b0.diagonal : undefined, diagDown: b0.diagonalDown ? b0.diagonal : undefined,
    };
    const touched = new Set<SideKey>();
    const sameSide = (a?: { style: string; color?: string }, b?: { style: string; color?: string }) => !!a && !!b && a.style === b.style && (a.color ?? "") === (b.color ?? "");
    const dashOf = (style: string) => ({ hair: "1 1", dotted: "1 2", dashDotDot: "6 2 2 2 2 2", dashDot: "6 2 2 2", dashed: "3 2", mediumDashDotDot: "6 2 2 2 2 2", slantDashDot: "8 2 2 2", mediumDashDot: "6 2 2 2", mediumDashed: "6 2" } as Record<string, string>)[style] ?? "";
    const widthOf = (style: string) => (/^(medium|mediumDash|slant)/.test(style) ? 2 : style === "thick" ? 3 : 1);
    const svgLine = (x1: number, y1: number, x2: number, y2: number, side: { style: string; color?: string }) => {
      const c = side.color ?? "#000";
      if (side.style === "double") {
        const dx = y1 === y2 ? 0 : 1.5, dy = y1 === y2 ? 1.5 : 0;
        return `<line x1="${x1 - dx}" y1="${y1 - dy}" x2="${x2 - dx}" y2="${y2 - dy}" stroke="${c}"/><line x1="${x1 + dx}" y1="${y1 + dy}" x2="${x2 + dx}" y2="${y2 + dy}" stroke="${c}"/>`;
      }
      return `<line x1="${x1}" y1="${y1}" x2="${x2}" y2="${y2}" stroke="${c}" stroke-width="${widthOf(side.style)}" stroke-dasharray="${dashOf(side.style)}"/>`;
    };
    const styleBtns = LINE_STYLES.map(([v, label]) => {
      const b = h("button", { type: "button", class: "fc-linestyle", role: "radio", "aria-label": label, "data-style": v });
      b.innerHTML = v === "none" ? "<span>Keine</span>" : `<svg width="64" height="10" aria-hidden="true">${svgLine(2, 5, 62, 5, { style: v })}</svg>`;
      b.onclick = () => { pen.style = v; paintBorders(); };
      return b;
    });
    const penColorBtn = h("button", { type: "button", class: "fc-color", "aria-label": "Farbe Automatisch", style: "background:#000" });
    penColorBtn.onclick = () => openColors(penColorBtn, (c) => {
      pen.color = c || undefined;
      penColorBtn.style.background = c || "#000";
      penColorBtn.setAttribute("aria-label", c ? `Farbe ${c}` : "Farbe Automatisch");
      paintBorders();
    });
    const toggle = (k: SideKey) => {
      if ((k === "insideH" && !rowsMany) || (k === "insideV" && !colsMany)) return;
      sides[k] = pen.style === "none" || sameSide(sides[k], pen) ? undefined : { style: pen.style, ...(pen.color ? { color: pen.color } : {}) };
      touched.add(k);
      paintBorders();
    };
    const preset = (which: "none" | "outline" | "inside") => {
      const keys: SideKey[] = which === "none" ? ["top", "bottom", "left", "right", "insideH", "insideV", "diagUp", "diagDown"] : which === "outline" ? ["top", "bottom", "left", "right"] : ["insideH", "insideV"];
      for (const k of keys) {
        if ((k === "insideH" && !rowsMany) || (k === "insideV" && !colsMany)) continue;
        sides[k] = which === "none" || pen.style === "none" ? undefined : { style: pen.style, ...(pen.color ? { color: pen.color } : {}) };
        touched.add(k);
      }
      paintBorders();
    };
    const sideBtn = (k: SideKey, label: string) => {
      const b = h("button", { type: "button", class: `fc-side fc-side-${k}`, role: "checkbox", "aria-label": label, disabled: (k === "insideH" && !rowsMany) || (k === "insideV" && !colsMany) || undefined });
      b.onclick = () => toggle(k);
      return b;
    };
    const sideBtns: Record<SideKey, HTMLElement> = {
      top: sideBtn("top", "Oben"), insideH: sideBtn("insideH", "Horizontal"), bottom: sideBtn("bottom", "Unten"), diagUp: sideBtn("diagUp", "Diagonal nach oben"),
      left: sideBtn("left", "Links"), insideV: sideBtn("insideV", "Vertikal"), right: sideBtn("right", "Rechts"), diagDown: sideBtn("diagDown", "Diagonal nach unten"),
    };
    const PW = 150, PH = 110, PX = 10, PY = 10;
    const borderPreview = h("div", { class: "fc-border-preview", role: "img", "aria-label": "Rahmen" });
    const paintBorders = () => {
      for (const b of styleBtns) b.classList.toggle("on", b.dataset.style === pen.style), b.setAttribute("aria-checked", String(b.dataset.style === pen.style));
      for (const [k, b] of Object.entries(sideBtns)) { b.classList.toggle("on", !!sides[k as SideKey]); b.setAttribute("aria-checked", String(!!sides[k as SideKey])); }
      const x0 = PX, y0 = PY, x1 = PX + PW, y1 = PY + PH, xm = (x0 + x1) / 2, ym = (y0 + y1) / 2;
      let svg = `<svg width="${PW + 2 * PX}" height="${PH + 2 * PY}" aria-hidden="true">`;
      // Corner marks and the sample text, as Excel's preview.
      for (const [cx, cy] of [[x0, y0], [x1, y0], [x0, y1], [x1, y1], ...(colsMany ? [[xm, y0], [xm, y1]] : []), ...(rowsMany ? [[x0, ym], [x1, ym]] : [])]) svg += `<path d="M${cx - 6} ${cy}h4M${cx + 2} ${cy}h4M${cx} ${cy - 6}v4M${cx} ${cy + 2}v4" stroke="#999"/>`;
      const cellsX = colsMany ? [x0, xm] : [x0], cellsY = rowsMany ? [y0, ym] : [y0], cw = colsMany ? PW / 2 : PW, ch = rowsMany ? PH / 2 : PH;
      for (const cx of cellsX) for (const cy of cellsY) svg += `<text x="${cx + cw / 2}" y="${cy + ch / 2}" font-size="11" fill="#666" text-anchor="middle" dominant-baseline="middle">Text</text>`;
      if (sides.top) svg += svgLine(x0, y0, x1, y0, sides.top);
      if (sides.bottom) svg += svgLine(x0, y1, x1, y1, sides.bottom);
      if (sides.left) svg += svgLine(x0, y0, x0, y1, sides.left);
      if (sides.right) svg += svgLine(x1, y0, x1, y1, sides.right);
      if (sides.insideH && rowsMany) svg += svgLine(x0, ym, x1, ym, sides.insideH);
      if (sides.insideV && colsMany) svg += svgLine(xm, y0, xm, y1, sides.insideV);
      for (const cx of cellsX) for (const cy of cellsY) {
        if (sides.diagDown) svg += svgLine(cx, cy, cx + cw, cy + ch, sides.diagDown);
        if (sides.diagUp) svg += svgLine(cx, cy + ch, cx + cw, cy, sides.diagUp);
      }
      borderPreview.innerHTML = svg + "</svg>";
    };
    // A click in the preview: the nearest line within a few pixels (sides, inner lines, diagonals of a cell).
    borderPreview.addEventListener("click", (e) => {
      const r = borderPreview.getBoundingClientRect();
      const x = e.clientX - r.left - PX, y = e.clientY - r.top - PY;
      const cw = colsMany ? PW / 2 : PW, ch = rowsMany ? PH / 2 : PH;
      const near: [SideKey, number][] = [["top", Math.abs(y)], ["bottom", Math.abs(y - PH)], ["left", Math.abs(x)], ["right", Math.abs(x - PW)]];
      if (rowsMany) near.push(["insideH", Math.abs(y - PH / 2)]);
      if (colsMany) near.push(["insideV", Math.abs(x - PW / 2)]);
      const lx = ((x % cw) + cw) % cw, ly = ((y % ch) + ch) % ch, len = Math.hypot(cw, ch);
      near.push(["diagDown", Math.abs(ch * lx - cw * ly) / len], ["diagUp", Math.abs(ch * lx + cw * ly - cw * ch) / len]);
      const [k, d] = near.sort((a, b) => a[1] - b[1])[0];
      if (d <= 8) toggle(k);
    });
    paintBorders();
    panels.Rahmen = h("div", { class: "fc-borders" },
      h("fieldset", { class: "fc-line" }, h("legend", null, "Linie"), h("div", null, "Art:"), h("div", { class: "fc-linestyles", role: "radiogroup", "aria-label": "Art:" }, ...styleBtns), h("div", null, "Farbe:"), penColorBtn),
      h("div", null,
        h("div", { class: "fc-label" }, "Voreinstellungen"),
        h("div", { class: "fc-presets" },
          h("button", { type: "button", "aria-label": "Keine", onclick: () => preset("none") }, "Keine"),
          h("button", { type: "button", "aria-label": "Außen", onclick: () => preset("outline") }, "Außen"),
          h("button", { type: "button", "aria-label": "Innen", disabled: !many || undefined, onclick: () => preset("inside") }, "Innen")),
        h("div", { class: "fc-label" }, "Rahmen"),
        h("div", { class: "fc-border-grid" },
          h("div", { class: "fc-side-col" }, sideBtns.top, sideBtns.insideH, sideBtns.bottom),
          borderPreview,
          h("div", { class: "fc-side-row" }, sideBtns.diagUp, sideBtns.left, sideBtns.insideV, sideBtns.right, sideBtns.diagDown))));

    // Ausfüllen: Hintergrundfarbe (Keine Farbe, the palette, Fülleffekte...), Musterfarbe and Musterformat, and the
    // Beispiel (D-147). A solid colour stays the cell's fill; a pattern or a gradient is written as Excel writes it.
    let fill = st.fill ?? "";
    let fancy: FancyFill | undefined = st.fill ? undefined : parseFill(st.fillXml);
    let patColor = fancy?.kind === "pattern" && fancy.type !== "solid" ? fancy.fg : "";
    let patType = fancy?.kind === "pattern" && fancy.type !== "solid" ? fancy.type : "";
    if (fancy?.kind === "pattern") fill = fancy.type === "solid" ? fancy.fg : fancy.bg ?? "";
    const fillSample = h("div", { class: "fc-fill-sample", role: "img", "aria-label": "Beispiel" });
    const currentFill = (): FancyFill | undefined => fancy?.kind === "gradient" ? fancy : patType ? { kind: "pattern", type: patType, fg: patColor || "#000000", ...(fill ? { bg: fill } : {}) } : undefined;
    const paintFill = () => { fillSample.style.background = cssFill(currentFill(), fill || "#fff"); patBtn.style.background = cssFill(patType ? { kind: "pattern", type: patType, fg: patColor || "#000000" } : undefined, "#fff"); };
    const grid = colorGrid((c) => { fill = hexColor(c); if (fancy?.kind === "gradient") fancy = undefined; paintFill(); });
    const patColorBtn = h("button", { type: "button", class: "fc-color", "aria-label": "Musterfarbe Automatisch", style: `background:${patColor || "#000"}` });
    patColorBtn.onclick = () => openColors(patColorBtn, (c) => {
      patColor = c; patColorBtn.style.background = c || "#000"; patColorBtn.setAttribute("aria-label", c ? `Musterfarbe ${c}` : "Musterfarbe Automatisch");
      if (fancy?.kind === "gradient") fancy = undefined;
      paintFill();
    });
    const patBtn = h("button", { type: "button", class: "fc-pattern-btn", "aria-label": "Musterformat:" });
    patBtn.onclick = () => {
      const pop = h("div", { class: "menu fc-pop fc-patterns", role: "listbox", "aria-label": "Musterformat:" },
        // Uni (the first entry) is no pattern: the cell has its Hintergrundfarbe only.
        ...PATTERNS.map(([type, label]) => h("button", { type: "button", role: "option", "aria-label": label, title: label, class: "fc-pattern",
          style: `background:${cssFill({ kind: "pattern", type, fg: "#000000" })}`, onclick: () => { patType = type === "solid" ? "" : type; if (fancy?.kind === "gradient") fancy = undefined; pop.remove(); paintFill(); } })));
      const r = patBtn.getBoundingClientRect();
      pop.style.cssText = `position:fixed;left:${r.left}px;top:${r.bottom + 2}px;z-index:90`;
      document.body.append(pop);
      setTimeout(() => addEventListener("pointerdown", function close(e) { if (!pop.contains(e.target as Node)) { pop.remove(); removeEventListener("pointerdown", close, true); } }, true));
    };
    const effects = h("button", { type: "button", onclick: async () => {
      const g = await fillEffects(fancy?.kind === "gradient" ? fancy : undefined, fill || "#FFFFFF");
      if (g) { fancy = g; patType = ""; paintFill(); }
    } }, "Fülleffekte...");
    const moreColors = h("input", { type: "color", "aria-label": "Weitere Farben...", value: fill || "#ffffff" }) as HTMLInputElement;
    moreColors.oninput = () => { fill = moreColors.value.toUpperCase(); if (fancy?.kind === "gradient") fancy = undefined; paintFill(); };
    paintFill();
    const fillKey = () => JSON.stringify([fill, currentFill() ?? null]);
    panels.Ausfüllen = h("div", { class: "fc-fillpage" },
      h("div", null, h("div", { class: "fc-label" }, "Hintergrundfarbe:"),
        h("button", { type: "button", onclick: () => { fill = ""; patType = ""; fancy = undefined; paintFill(); } }, "Keine Farbe"),
        h("div", { class: "fc-palette" }, ...grid), h("div", { class: "fc-presets" }, effects, h("label", { class: "field" }, "Weitere Farben...", moreColors))),
      h("div", null, field("Musterfarbe:", patColorBtn), field("Musterformat:", patBtn), h("div", { class: "fc-label" }, "Beispiel"), fillSample));

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
      code: codeBox.value, h: hSel.value, v: vSel.value, indent: indent.value, rot: vertical ? "v" : String(deg), wrap: wrap.input.checked, shrink: shrink.input.checked,
      name: strip(fName.value.trim()), size: fSize.value.trim(), style: fStyle.value.trim(), under: fUnder.value, color: fColor, strike: strike.input.checked,
      sup: sup.input.checked, sub: sub.input.checked, fill: fillKey(), locked: locked.input.checked, hidden: hidden.input.checked,
    });
    const initial = shownAt();
    const ok = await showDialog({ title: "Zellen formatieren", body, buttons: [{ label: "OK", value: true, primary: true }, { label: "Abbrechen", value: false }], cancel: false });
    if (!ok) return;
    const now = shownAt();
    const changed = (k: keyof typeof now) => now[k] !== initial[k];
    let code: string;
    try { code = codeBox.value === "Standard" ? "General" : fileFormat(codeBox.value); } catch { code = st.numFmt; }
    const changedBorders = touched.size > 0;
    app.change("Zellen formatieren", () => {
      A.format(app.ctx, sheet(), app.grid.selection, (s) => {
        if (changed("code")) s.numFmt = code;
        if (changed("h")) s.h = hSel.value === "general" ? undefined : hSel.value as never;
        if (changed("v")) s.v = vSel.value === "bottom" ? undefined : vSel.value as never;
        if (changed("indent")) s.indent = Number(indent.value) || undefined;
        if (changed("rot")) s.rotation = vertical ? 255 : deg === 0 ? undefined : deg > 0 ? deg : 90 - deg;
        if (changed("wrap")) s.wrap = wrap.input.checked || undefined;
        if (changed("shrink")) s.shrink = shrink.input.checked || undefined;
        if (changed("name")) s.font.name = strip(fName.value.trim()) || s.font.name;
        const size = num(fSize.value);
        if (changed("size") && size >= 1 && size <= 409) s.font.size = size;
        if (changed("style")) {
          s.font.bold = /fett/i.test(fStyle.value) || undefined;
          s.font.italic = /kursiv/i.test(fStyle.value) || undefined;
        }
        if (changed("under")) s.font.underline = (fUnder.value || undefined) as never;
        if (changed("color")) s.font.color = fColor || undefined;
        if (changed("strike")) s.font.strike = strike.input.checked || undefined;
        if (changed("sup") || changed("sub")) s.font.vertAlign = sup.input.checked ? "superscript" : sub.input.checked ? "subscript" : undefined;
        if (changed("fill")) {
          const f = currentFill();
          s.fill = f ? undefined : fill || undefined;
          s.fillXml = f ? fillXmlOf(f) : undefined;
          delete s.fillTheme;
        }
        if (changed("locked")) s.locked = locked.input.checked ? undefined : false;
        if (changed("hidden")) s.hidden = hidden.input.checked || undefined;
      });
      if (changedBorders) applyBorderSides(sides, touched);
      if (changed("code")) app.autoWiden(app.grid.selection);
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

  /**
   * Rahmen of Zellen formatieren: the sides touched in the dialog, on every cell of the selection: the outer ones on
   * the edge cells, Horizontal and Vertikal between the cells (on both cells, as Alle Rahmenlinien), the diagonals in
   * every cell (the file has one diagonal line style per cell, with flags for the directions).
   */
  function applyBorderSides(sides: Partial<Record<string, { style: string; color?: string }>>, touched: Set<string>): void {
    const side = (k: string) => (sides[k] ? { ...sides[k]! } : undefined);
    for (const r of app.grid.selection) {
      app.history.cells(sheet(), r);
      for (let row = r.r0; row <= r.r1; row++) for (let col = r.c0; col <= r.c1; col++) {
        const c = sheet().get(row, col) ?? { v: null };
        const s = app.wb.styles.derive(c.s, (st) => {
          const b = st.borders;
          const put = (k: "top" | "bottom" | "left" | "right", from: string) => { if (touched.has(from)) b[k] = side(from); };
          put("top", row === r.r0 ? "top" : "insideH");
          put("bottom", row === r.r1 ? "bottom" : "insideH");
          put("left", col === r.c0 ? "left" : "insideV");
          put("right", col === r.c1 ? "right" : "insideV");
          if (touched.has("diagUp") || touched.has("diagDown")) {
            const up = touched.has("diagUp") ? side("diagUp") : b.diagonalUp ? b.diagonal : undefined;
            const down = touched.has("diagDown") ? side("diagDown") : b.diagonalDown ? b.diagonal : undefined;
            b.diagonal = down ?? up;
            b.diagonalUp = up ? true : undefined;
            b.diagonalDown = down ? true : undefined;
            if (!b.diagonal) { delete b.diagonal; delete b.diagonalUp; delete b.diagonalDown; }
          }
          for (const k of ["top", "bottom", "left", "right"] as const) if (!b[k]) delete b[k];
        });
        sheet().set(row, col, { ...c, s });
      }
    }
  }

  /** Fülleffekte (Ausfüllen): Excel's dialog with Farben, Schattierungsarten, Varianten and Beispiel (D-147). */
  async function fillEffects(start: GradientFill | undefined, base: string): Promise<GradientFill | null> {
    let c1 = start?.stops[0]?.color ?? base, c2 = start?.stops.at(-1)?.color ?? "#4472C4";
    let kind: Shading = start?.path ? (start.path.left === 0.5 ? "Aus der Mitte" : "Aus der Ecke") : start?.degree === 0 ? "Vertikal" : start?.degree === 45 ? "Diagonal oben" : start?.degree === 135 ? "Diagonal unten" : "Horizontal";
    let variant = 0;
    let colors: "Einfarbig" | "Zweifarbig" = "Zweifarbig";
    let light = 0.5;
    const second = () => (colors === "Einfarbig" ? mixHex(c1, light >= 0.5 ? "#FFFFFF" : "#000000", Math.abs(light - 0.5) * 2) : c2);
    const sample = h("div", { class: "fc-fill-sample", role: "img", "aria-label": "Beispiel:" });
    const variants = h("div", { class: "fc-variants", role: "radiogroup", "aria-label": "Varianten" });
    const c1Btn = h("button", { type: "button", class: "fc-color", "aria-label": "Füllfarbe 1" });
    const c2Btn = h("button", { type: "button", class: "fc-color", "aria-label": "Füllfarbe 2" });
    const lightIn = h("input", { type: "range", min: "0", max: "100", value: "50", "aria-label": "Dunkel Hell" }) as HTMLInputElement;
    const c2Row = h("div", null, h("span", null, "Farbe 2:"), c2Btn);
    const lightRow = h("div", null, h("span", null, "Dunkel"), lightIn, h("span", null, "Hell"));
    const paint = () => {
      c1Btn.style.background = c1;
      c2Btn.style.background = c2;
      c2Row.hidden = colors !== "Zweifarbig";
      lightRow.hidden = colors !== "Einfarbig";
      const n = kind === "Aus der Mitte" ? 2 : 4;
      if (variant >= n) variant = 0;
      variants.replaceChildren(...range(0, n - 1).map((i) => h("button", { type: "button", role: "radio", "aria-label": `Variante ${i + 1}`, "aria-checked": String(i === variant), class: `fc-variant${i === variant ? " on" : ""}`,
        style: `background:${cssFill(shadingFill(kind, i, c1, second()))}`, onclick: () => { variant = i; paint(); } })));
      sample.style.background = cssFill(shadingFill(kind, variant, c1, second()));
    };
    c1Btn.onclick = () => openColors(c1Btn, (c) => { c1 = c || "#FFFFFF"; paint(); });
    c2Btn.onclick = () => openColors(c2Btn, (c) => { c2 = c || "#FFFFFF"; paint(); });
    lightIn.oninput = () => { light = Number(lightIn.value) / 100; paint(); };
    const radioOf = (group: string, label: string, on: boolean, pick: () => void, disabled = false) => {
      const input = h("input", { type: "radio", name: group, checked: on || undefined, disabled: disabled || undefined }) as HTMLInputElement;
      input.onchange = () => { if (input.checked) { pick(); paint(); } };
      return h("label", { class: "check" }, input, label);
    };
    paint();
    const body = h("div", { class: "fc-effects" },
      h("div", { class: "fc-tabs", role: "tablist" }, h("button", { type: "button", role: "tab", class: "on" }, "Graduell")),
      h("div", { class: "fc-effects-top" },
        h("fieldset", null, h("legend", null, "Farben"),
          radioOf("fe-colors", "Einfarbig", false, () => { colors = "Einfarbig"; }), radioOf("fe-colors", "Zweifarbig", true, () => { colors = "Zweifarbig"; }),
          // Voreinstellung (Office's ready-made gradients) is not built (D-147).
          radioOf("fe-colors", "Voreinstellung", false, () => {}, true)),
        h("div", null, h("div", null, h("span", null, "Farbe 1:"), c1Btn), c2Row, lightRow)),
      h("div", { class: "fc-effects-top" },
        h("fieldset", null, h("legend", null, "Schattierungsarten"), ...SHADINGS.map((k) => radioOf("fe-shade", k, k === kind, () => { kind = k; }))),
        h("fieldset", null, h("legend", null, "Varianten"), variants),
        h("div", null, h("div", null, "Beispiel:"), sample)));
    const ok = await showDialog({ title: "Fülleffekte", body, buttons: [{ label: "OK", value: true, primary: true }, { label: "Abbrechen", value: false }], cancel: false });
    return ok ? shadingFill(kind, variant, c1, second()) : null;
  }

  function openColors(anchor: HTMLElement, pick: (hex: string) => void): void {
    const pop = h("div", { class: "menu fc-pop" }, h("button", { type: "button", onclick: () => { pick(""); pop.remove(); } }, "Automatisch"),
      h("div", { class: "fc-palette" }, ...colorGrid((c) => { pick(hexColor(c)); pop.remove(); })));
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
    app.recorder.sort(sheet(), rect, keys);
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
        sheet().set(row, col, c.v === null && c.s === undefined && !c.f ? undefined : { ...c, ...(c.f !== undefined && d ? { f: shiftFormula(c.f, d, 0, { clipRanges: true }) } : {}) });
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
    const c = sheet().get(a.row, a.col);
    const t = shown(c, app.wb.styles.get(c?.s)).text;
    // One undo step, as in Excel: switching the AutoFilter on (when it was off) and filtering by the value (D-142).
    app.change("Filtern", () => {
      app.history.whole(sheet());
      if (!sheet().autoFilter) sheet().autoFilter = { ref: sortRange().rect, columns: {} };
      const f = sheet().autoFilter!;
      f.columns[a.col - f.ref.c0] = { values: [t] };
      applyFilter();
    });
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
    // A press outside closes the panel as Office closes its menus (D-156).
    const outside = (e: PointerEvent) => { if (!panel.contains(e.target as Node) && !(e.target as Element).closest?.(".menu-root")) pressOutside(e, close); };
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
      // A new rule goes to the top of the order (priority 1) and the others move down, as Excel adds rules (D-146).
      for (const f of list) for (const r of f.rules) r.priority += 1;
      list.push({ ref, rules: [{ ...rule, priority: 1 }] });
    });
  }

  /** Neue Regel..., Weitere Regeln... and Regeln verwalten... (condrules.ts, D-143). */
  const cr = condRules(app, addRule);

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
      { label: "Textinhalt...", run: () => void ruleDialog("Textinhalt", "Zellen formatieren, die folgenden Text enthalten:", 1, (v, f) => textRule("containsText", v[0], app.grid.selection[0], f)) },
      { label: "Datum...", run: () => void dateRuleDialog() },
      { label: "Doppelte Werte...", run: () => void duplicateDialog() },
      "-",
      { label: "Weitere Regeln...", run: () => void cr.newRule(1) },
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
      "-",
      { label: "Weitere Regeln...", run: () => void cr.newRule(2) },
    ];
  }

  function conditionalBarMenu(): MenuEntry[] {
    // Excel's names (recorded in the inventory, D-143): Farbverlauf, then Einfarbige Füllung with the same six
    // colours (recorded with real clicks, probe_excel_followup_ui.py, D-146).
    const bars: [string, string][] = [["Blauer Datenbalken", "#638EC6"], ["Grüner Datenbalken", "#63C384"], ["Roter Datenbalken", "#FF555A"], ["Orangefarbener Datenbalken", "#FFB628"], ["Hellblauer Datenbalken", "#008AEF"], ["Lilafarbener Datenbalken", "#D6007B"]];
    return [{ heading: "Farbverlauf" }, ...bars.map(([l, c]): MenuEntry => ({ label: l, run: () => void addRule({ type: "dataBar", colors: [c] }) })),
      { heading: "Einfarbige Füllung" }, ...bars.map(([l, c]): MenuEntry => ({ label: l, run: () => void addRule({ type: "dataBar", colors: [c], solid: true }) })),
      "-", { label: "Weitere Regeln...", run: () => void cr.newRule(0) }];
  }

  function conditionalScaleMenu(): MenuEntry[] {
    // Excel's twelve scales by its names. A name lists the colours from the highest value down; the rule keeps
    // them from the lowest up (Grün-Gelb-Rot: lowest red, highest green, D-143).
    const R = "#F8696B", Y = "#FFEB84", G = "#63BE7B", W = "#FCFCFF", B = "#5A8AC6", Y2 = "#FFEF9C";
    const scales: [string, string[]][] = [
      ["Grün-Gelb-Rot", [R, Y, G]], ["Rot-Gelb-Grün", [G, Y, R]], ["Grün-Weiß-Rot", [R, W, G]], ["Rot-Weiß-Grün", [G, W, R]],
      ["Blau-Weiß-Rot", [R, W, B]], ["Rot-Weiß-Blau", [B, W, R]], ["Weiß-Rot", [R, W]], ["Rot-Weiß", [W, R]],
      ["Grün-Weiß", [W, G]], ["Weiß-Grün", [G, W]], ["Grün-Gelb", [Y2, G]], ["Gelb-Grün", [G, Y2]],
    ];
    return [...scales.map(([l, c]): MenuEntry => ({ label: `${l}-Farbskala`, run: () => void addRule({ type: "colorScale", colors: c }) })), "-", { label: "Weitere Regeln...", run: () => void cr.newRule(0) }];
  }

  /** Symbolsätze, in the gallery's groups (Richtungen, Formen, Indikatoren, Bewertungen). */
  function conditionalIconMenu(): MenuEntry[] {
    const out: MenuEntry[] = [];
    for (const [group, sets] of ICON_SETS) {
      out.push({ heading: group });
      for (const [label, set] of sets) out.push({ label, run: () => void addRule(iconRule(set)) });
    }
    out.push("-", { label: "Weitere Regeln...", run: () => void cr.newRule(0) });
    return out;
  }

  function clearConditional(selection: boolean): void {
    // Nothing to delete: no undo step (Excel).
    const list = sheet().conditional ?? [];
    if (!list.some((f) => !selection || f.ref.some((r) => app.grid.selection.some((s) => A.overlaps(r, s))))) return;
    app.change("Regeln löschen", () => {
      app.history.whole(sheet());
      if (!selection) { sheet().conditional = undefined; return; }
      // Excel takes the selected cells out of each rule's range; a rule left without cells goes.
      const out = [];
      for (const f of sheet().conditional ?? []) {
        let ref = f.ref;
        for (const s of app.grid.selection) ref = ref.flatMap((r) => subtract(r, s));
        if (!ref.length) continue;
        // The formulas are written for the range's top-left cell: when that moves, they move with it.
        const dr = ref[0].r0 - f.ref[0].r0, dc = ref[0].c0 - f.ref[0].c0;
        const rules = dr || dc ? f.rules.map((r) => (r.formulas ? { ...r, formulas: r.formulas.map((x) => shiftFormula(x, dr, dc)) } : r)) : f.rules;
        out.push({ ...f, ref, rules });
      }
      sheet().conditional = out.length ? out : undefined;
    });
  }

  /** A rectangle without the cells of another: up to four rectangles (above, below, left, right). */
  function subtract(r: Rect, s: Rect): Rect[] {
    if (!A.overlaps(r, s)) return [r];
    const out: Rect[] = [];
    if (s.r0 > r.r0) out.push({ ...r, r1: s.r0 - 1 });
    if (s.r1 < r.r1) out.push({ ...r, r0: s.r1 + 1 });
    const r0 = Math.max(r.r0, s.r0), r1 = Math.min(r.r1, s.r1);
    if (s.c0 > r.c0) out.push({ r0, r1, c0: r.c0, c1: s.c0 - 1 });
    if (s.c1 < r.c1) out.push({ r0, r1, c0: s.c1 + 1, c1: r.c1 });
    return out;
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

  /**
   * Aus Auswahl erstellen (Strg+Umschalt+F3): names from the labels in the selection's top row, left
   * column, bottom row or right column, each naming the cells beside it (Excel's dialog, D-142).
   */
  async function createFromSelection(): Promise<void> {
    const r = sel();
    const isText = (row: number, col: number) => typeof sheet().get(row, col)?.v === "string";
    const multiRow = r.r1 > r.r0, multiCol = r.c1 > r.c0;
    const topGuess = multiRow && range(r.c0 + (multiCol ? 1 : 0), r.c1).some((c) => isText(r.r0, c));
    const leftGuess = multiCol && range(r.r0 + (multiRow ? 1 : 0), r.r1).some((row) => isText(row, r.c0));
    const top = checkbox("Oberster Zeile", topGuess), left = checkbox("Linker Spalte", leftGuess), bottom = checkbox("Unterster Zeile"), right = checkbox("Rechter Spalte");
    const body = h("div", null, h("div", { class: "fc-label" }, "Namen erstellen aus den Werten in:"), top.el, left.el, bottom.el, right.el);
    const ok = await showDialog({ title: "Namen aus Auswahl erstellen", body, buttons: [{ label: "OK", value: true, primary: true }, { label: "Abbrechen", value: false }], cancel: false });
    if (!ok) return;
    const t = top.input.checked, l = left.input.checked, b = bottom.input.checked, rt = right.input.checked;
    if (!(t || l || b || rt)) return;
    const body0 = { r0: r.r0 + (t ? 1 : 0), r1: r.r1 - (b ? 1 : 0), c0: r.c0 + (l ? 1 : 0), c1: r.c1 - (rt ? 1 : 0) };
    if (body0.r0 > body0.r1 || body0.c0 > body0.c1) return void messageBox("RaabeCalc", "Aus dieser Auswahl lassen sich keine Namen bilden.");
    const clean = (v: unknown) => {
      let n = String(v ?? "").trim().replace(/\s+/g, "_").replace(/[^A-Za-z0-9ÄÖÜäöüß_.\\]/g, "");
      if (!n) return "";
      if (/^\d/.test(n) || /^[A-Za-z]{1,3}\d+$/.test(n) || /^[rcRC]$/.test(n)) n = "_" + n;
      return n;
    };
    const abs = (x: Rect) => `${quoteSheet(sheet().name)}!${rangeName(x).replace(/([A-Z]+)(\d+)/g, "$$$1$$$2")}`;
    const made: { name: string; ref: string }[] = [];
    const add = (label: unknown, x: Rect) => { const n = clean(label); if (n && !made.some((m) => m.name.toLowerCase() === n.toLowerCase())) made.push({ name: n, ref: abs(x) }); };
    for (let c = body0.c0; c <= body0.c1; c++) {
      if (t) add(sheet().get(r.r0, c)?.v, { r0: body0.r0, r1: body0.r1, c0: c, c1: c });
      if (b) add(sheet().get(r.r1, c)?.v, { r0: body0.r0, r1: body0.r1, c0: c, c1: c });
    }
    for (let row = body0.r0; row <= body0.r1; row++) {
      if (l) add(sheet().get(row, r.c0)?.v, { r0: row, r1: row, c0: body0.c0, c1: body0.c1 });
      if (rt) add(sheet().get(row, r.c1)?.v, { r0: row, r1: row, c0: body0.c0, c1: body0.c1 });
    }
    if (!made.length) return;
    app.change("Namen erstellen", () => {
      app.history.sheets();
      app.wb.names = [...app.wb.names.filter((x) => !made.some((m) => m.name.toLowerCase() === x.name.toLowerCase())), ...made];
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
    addRule, newRule: cr.newRule, manageRules: cr.manager, conditionalIconMenu, conditionalHighlightMenu, conditionalTopMenu, conditionalBarMenu, conditionalScaleMenu, clearConditional, defineName, createFromSelection, nameManager, note, link,
    symbol, insertFunction, moveSheet, unhideSheet, zoom, statistics, options, pasteSpecial,
  };
}

function mixHex(a: string, b: string, t: number): string {
  const p = (x: string, i: number) => parseInt(x.replace("#", "").slice(i * 2, i * 2 + 2), 16);
  return "#" + [0, 1, 2].map((i) => Math.round(p(a, i) + (p(b, i) - p(a, i)) * t).toString(16).padStart(2, "0")).join("").toUpperCase();
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
