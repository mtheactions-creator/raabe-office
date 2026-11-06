// Excel tables in RaabeCalc (Einfügen > Tabelle, Strg+T / Strg+L, Start > Als Tabelle formatieren,
// the contextual tab Tabellenentwurf; D-066). The table style is drawn under the cells' own
// formatting, element by element in Excel's order (whole table, column stripes, row stripes, last and
// first column, header and total row, their corner cells), from the 60 built-in styles recorded in
// Excel (table-styles.gen.ts) with theme-relative colours.

import { h, messageBox, showDialog } from "@raabe/shell";
import {
  nextTableName, parseRange, rangeName, tableData, validTableName,
  type BorderSide, type Borders, type CellStyle, type Rect, type Sheet, type SheetTable,
} from "@raabe/sheet-engine";
import type { CalcApp } from "./app";
import { DEFAULT_TABLE_STYLE, TABLE_STYLE_DEFS, type TsColor, type TsElement } from "./table-styles.gen";

// Windows' integer HLS, as Excel tints theme colours (the same as the .xlsx reader's).
const HLS = 240, RGB = 255;
const idiv = (a: number, b: number) => Math.trunc(a / b);
function rgbToHls(r: number, g: number, b: number): [number, number, number] {
  const max = Math.max(r, g, b), min = Math.min(r, g, b);
  const l = idiv((max + min) * HLS + RGB, 2 * RGB);
  if (max === min) return [idiv(HLS * 2, 3), l, 0];
  const s = l <= HLS / 2 ? idiv((max - min) * HLS + idiv(max + min, 2), max + min) : idiv((max - min) * HLS + idiv(2 * RGB - max - min, 2), 2 * RGB - max - min);
  const d = (c: number) => idiv((max - c) * (HLS / 6) + idiv(max - min, 2), max - min);
  let hh = r === max ? d(b) - d(g) : g === max ? HLS / 3 + d(r) - d(b) : (2 * HLS) / 3 + d(g) - d(r);
  if (hh < 0) hh += HLS;
  if (hh > HLS) hh -= HLS;
  return [hh, l, s];
}
function hueToRgb(n1: number, n2: number, hh: number): number {
  if (hh < 0) hh += HLS;
  if (hh > HLS) hh -= HLS;
  if (hh < HLS / 6) return n1 + idiv((n2 - n1) * hh + HLS / 12, HLS / 6);
  if (hh < HLS / 2) return n2;
  if (hh < (HLS * 2) / 3) return n1 + idiv((n2 - n1) * ((HLS * 2) / 3 - hh) + HLS / 12, HLS / 6);
  return n1;
}
function hlsToRgb(hh: number, l: number, s: number): [number, number, number] {
  if (s === 0) { const v = idiv(l * RGB, HLS); return [v, v, v]; }
  const m2 = l <= HLS / 2 ? idiv(l * (HLS + s) + HLS / 2, HLS) : l + s - idiv(l * s + HLS / 2, HLS);
  const m1 = 2 * l - m2;
  return [hueToRgb(m1, m2, hh + HLS / 3), hueToRgb(m1, m2, hh), hueToRgb(m1, m2, hh - HLS / 3)].map((x) => idiv(x * RGB + HLS / 2, HLS)) as [number, number, number];
}
function tinted(hex: string, tint: number): string {
  const v = hex.replace("#", "");
  if (!tint) return "#" + v.toUpperCase();
  const [r, g, b] = [0, 2, 4].map((i) => parseInt(v.slice(i, i + 2), 16));
  const [hh, l0, s] = rgbToHls(r, g, b);
  const l = Math.round(tint < 0 ? l0 * (1 + tint) : l0 * (1 - tint) + HLS * tint);
  return "#" + hlsToRgb(hh, Math.max(0, Math.min(HLS, l)), s).map((x) => Math.max(0, Math.min(255, x)).toString(16).padStart(2, "0")).join("").toUpperCase();
}

const OFFICE_2013: Record<string, string> = {
  dk1: "#000000", lt1: "#FFFFFF", dk2: "#44546A", lt2: "#E7E6E6", accent1: "#4472C4", accent2: "#ED7D31", accent3: "#A5A5A5",
  accent4: "#FFC000", accent5: "#5B9BD5", accent6: "#70AD47", hlink: "#0563C1", folHlink: "#954F72",
};

/** Excel's order: later elements win. */
const ORDER = ["wholeTable", "firstColumnStripe", "secondColumnStripe", "firstRowStripe", "secondRowStripe", "lastColumn", "firstColumn",
  "headerRow", "totalRow", "firstHeaderCell", "lastHeaderCell", "firstTotalCell", "lastTotalCell"] as const;

interface Look { fill?: string; color?: string; bold?: boolean; borders: Partial<Record<"left" | "right" | "top" | "bottom", BorderSide>> }

export function tablesFeature(app: CalcApp) {
  const color = (c: TsColor): string => (typeof c === "string" ? c : tinted((app.wb.theme[c.t] ?? OFFICE_2013[c.t] ?? "#000000"), c.tint ?? 0));
  const def = (name: string) => TABLE_STYLE_DEFS.find((d) => d.name === name);

  /** The table a cell is in. */
  const tableAt = (sheet: Sheet, row: number, col: number) =>
    sheet.tables?.find((t) => row >= t.ref.r0 && row <= t.ref.r1 && col >= t.ref.c0 && col <= t.ref.c1);

  /** The look the table style gives one cell. */
  function look(t: SheetTable, row: number, col: number): Look | null {
    const d = def(t.style);
    if (!d) return null;
    const out: Look = { borders: {} };
    const data = tableData(t);
    const inData = row >= data.r0 && row <= data.r1;
    const regions: [string, Rect][] = [["wholeTable", t.ref]];
    const stripe = (el: TsElement | undefined) => el?.stripe ?? 1;
    if (t.columnStripes && inData) {
      const s1 = stripe(d.elements.firstColumnStripe), s2 = stripe(d.elements.secondColumnStripe);
      const k = (col - t.ref.c0) % (s1 + s2);
      const first = k < s1;
      const c0 = col - (first ? k : k - s1);
      regions.push([first ? "firstColumnStripe" : "secondColumnStripe", { r0: data.r0, r1: data.r1, c0, c1: c0 + (first ? s1 : s2) - 1 }]);
    }
    if (t.rowStripes && inData) {
      const s1 = stripe(d.elements.firstRowStripe), s2 = stripe(d.elements.secondRowStripe);
      const k = (row - data.r0) % (s1 + s2);
      const first = k < s1;
      const r0 = row - (first ? k : k - s1);
      regions.push([first ? "firstRowStripe" : "secondRowStripe", { r0, r1: r0 + (first ? s1 : s2) - 1, c0: t.ref.c0, c1: t.ref.c1 }]);
    }
    if (t.lastColumn) regions.push(["lastColumn", { ...t.ref, c0: t.ref.c1 }]);
    if (t.firstColumn) regions.push(["firstColumn", { ...t.ref, c1: t.ref.c0 }]);
    if (t.header) regions.push(["headerRow", { ...t.ref, r1: t.ref.r0 }]);
    if (t.totals) regions.push(["totalRow", { ...t.ref, r0: t.ref.r1 }]);
    if (t.header && t.firstColumn) regions.push(["firstHeaderCell", { r0: t.ref.r0, r1: t.ref.r0, c0: t.ref.c0, c1: t.ref.c0 }]);
    if (t.header && t.lastColumn) regions.push(["lastHeaderCell", { r0: t.ref.r0, r1: t.ref.r0, c0: t.ref.c1, c1: t.ref.c1 }]);
    if (t.totals && t.firstColumn) regions.push(["firstTotalCell", { r0: t.ref.r1, r1: t.ref.r1, c0: t.ref.c0, c1: t.ref.c0 }]);
    if (t.totals && t.lastColumn) regions.push(["lastTotalCell", { r0: t.ref.r1, r1: t.ref.r1, c0: t.ref.c1, c1: t.ref.c1 }]);
    regions.sort((a, b) => ORDER.indexOf(a[0] as (typeof ORDER)[number]) - ORDER.indexOf(b[0] as (typeof ORDER)[number]));
    for (const [name, r] of regions) {
      if (row < r.r0 || row > r.r1 || col < r.c0 || col > r.c1) continue;
      const el = d.elements[name];
      if (!el) continue;
      if (el.fill) out.fill = color(el.fill);
      if (el.font) out.color = color(el.font);
      if (el.bold) out.bold = true;
      const b = el.borders;
      if (b) {
        const side = (s: { style: string; color: TsColor }): BorderSide => ({ style: s.style, color: color(s.color) });
        if (b.left && col === r.c0) out.borders.left = side(b.left);
        if (b.right && col === r.c1) out.borders.right = side(b.right);
        if (b.top && row === r.r0) out.borders.top = side(b.top);
        if (b.bottom && row === r.r1) out.borders.bottom = side(b.bottom);
        if (b.insideV) { if (col > r.c0) out.borders.left = side(b.insideV); if (col < r.c1) out.borders.right = side(b.insideV); }
        if (b.insideH) { if (row > r.r0) out.borders.top = side(b.insideH); if (row < r.r1) out.borders.bottom = side(b.insideH); }
      }
    }
    return out;
  }

  const cache = new Map<string, CellStyle>();
  let cacheKey = "";

  /** The cell's style with the table style under its own formatting (for the renderer). */
  function styleFor(row: number, col: number, st: CellStyle): CellStyle {
    const sheet = app.sheet;
    const t = tableAt(sheet, row, col);
    if (!t || !t.style) return st;
    const key = JSON.stringify([app.wb.theme, sheet.tables]);
    if (key !== cacheKey) { cache.clear(); cacheKey = key; }
    const lk = look(t, row, col);
    if (!lk) return st;
    const k = `${JSON.stringify(lk)}|${JSON.stringify(st)}`;
    let out = cache.get(k);
    if (!out) {
      out = structuredClone(st);
      if (!out.fill && lk.fill) out.fill = lk.fill;
      if (!out.font.color && lk.color) out.font.color = lk.color;
      if (lk.bold && !out.font.bold) out.font.bold = true;
      const b: Borders = { ...out.borders };
      for (const s of ["left", "right", "top", "bottom"] as const) if (!b[s] && lk.borders[s]) b[s] = lk.borders[s];
      out.borders = b;
      cache.set(k, out);
    }
    return out;
  }

  // ----- making and changing tables --------------------------------------------------------------

  const allNames = () => app.wb.sheets.flatMap((s) => (s.tables ?? []).map((t) => t.name));

  /** Column names: the header cells' texts, made unique and filled in as Excel does ("Spalte1" …). */
  function columnNames(sheet: Sheet, r: Rect, header: boolean): string[] {
    const names: string[] = [];
    for (let c = r.c0; c <= r.c1; c++) {
      let n = header ? String(sheet.value(r.r0, c) ?? "").trim() : "";
      if (!n) n = `Spalte${c - r.c0 + 1}`;
      let k = 2;
      const base = n;
      while (names.some((x) => x.toLowerCase() === n.toLowerCase())) n = `${base}${k++}`;
      names.push(n);
    }
    return names;
  }

  /** Makes a table over r (with the header row as given, or one inserted above). */
  function create(r: Rect, header: boolean, style = DEFAULT_TABLE_STYLE): SheetTable | null {
    const sheet = app.sheet;
    if (sheet.tables?.some((t) => t.ref.r0 <= r.r1 && t.ref.r1 >= r.r0 && t.ref.c0 <= r.c1 && t.ref.c1 >= r.c0)) {
      void messageBox("RaabeCalc", "Eine Tabelle darf sich nicht mit einer anderen Tabelle überschneiden.");
      return null;
    }
    let ref = { ...r };
    const t: SheetTable = { name: nextTableName(allNames()), ref, header: true, totals: false, columns: [], style,
      firstColumn: false, lastColumn: false, rowStripes: true, columnStripes: false, filterButton: true };
    app.change("Tabelle erstellen", () => {
      app.history.whole(sheet);
      if (!header) {
        // Without headers Excel puts a row of "Spalte1", "Spalte2" … above the data (shifting it down).
        ref = { ...r, r1: r.r1 + 1 };
        for (let row = r.r1; row >= r.r0; row--) for (let c = r.c0; c <= r.c1; c++) sheet.set(row + 1, c, sheet.get(row, c));
        for (let c = r.c0; c <= r.c1; c++) sheet.set(r.r0, c, { v: `Spalte${c - r.c0 + 1}` });
      }
      const names = columnNames(sheet, ref, true);
      // Header cells hold their names as text.
      names.forEach((n, i) => { const c = sheet.get(ref.r0, ref.c0 + i); if (c?.v !== n) sheet.set(ref.r0, ref.c0 + i, { ...(c ?? {}), v: n, f: undefined }); });
      t.ref = ref;
      t.columns = names.map((name) => ({ name }));
      sheet.tables = [...(sheet.tables ?? []), t];
      // The sheet's AutoFilter gives way: a table has its own.
      if (sheet.autoFilter && sheet.autoFilter.ref.r0 <= ref.r1 && sheet.autoFilter.ref.r1 >= ref.r0 && sheet.autoFilter.ref.c0 <= ref.c1 && sheet.autoFilter.ref.c1 >= ref.c0) sheet.autoFilter = undefined;
    });
    return t;
  }

  /** Changes the selected table (one undo step). */
  function edit(label: string, fn: (t: SheetTable) => void): void {
    const t = current();
    if (!t) return;
    const sheet = app.sheet;
    app.change(label, () => {
      app.history.whole(sheet);
      const i = sheet.tables!.indexOf(t);
      const next: SheetTable = structuredClone(t);
      fn(next);
      sheet.tables = sheet.tables!.map((x, k) => (k === i ? next : x));
    });
  }

  /** The table the active cell is in. */
  function current(): SheetTable | undefined {
    const a = app.sheet.view.active;
    return tableAt(app.sheet, a.row, a.col);
  }

  function setOption(k: "header" | "totals" | "firstColumn" | "lastColumn" | "rowStripes" | "columnStripes" | "filterButton", on: boolean): void {
    const t = current();
    if (!t) return;
    const sheet = app.sheet;
    if (k === "totals") { setTotals(on); return; }
    if (k === "header") {
      edit(on ? "Kopfzeile" : "Kopfzeile", (x) => {
        if (on === x.header) return;
        if (on) {
          // The header row comes back above the data (into the row above the table).
          if (x.ref.r0 === 0) return;
          x.ref = { ...x.ref, r0: x.ref.r0 - 1 };
          x.columns.forEach((c, i) => sheet.set(x.ref.r0, x.ref.c0 + i, { v: c.name }));
        } else {
          x.columns.forEach((_, i) => sheet.set(x.ref.r0, x.ref.c0 + i, undefined));
          x.ref = { ...x.ref, r0: x.ref.r0 + 1 };
        }
        x.header = on;
      });
      return;
    }
    edit("Tabellenformatoptionen", (x) => { x[k] = on; });
  }

  /** Ergebniszeile: a row below the data with "Ergebnis" and the last column's sum (Excel's default). */
  function setTotals(on: boolean): void {
    const sheet = app.sheet;
    edit("Ergebniszeile", (x) => {
      if (on === x.totals) return;
      if (on) {
        const row = x.ref.r1 + 1;
        x.ref = { ...x.ref, r1: row };
        x.totals = true;
        const last = x.columns.length - 1;
        x.columns = x.columns.map((c, i) => (i === 0 && last > 0 ? { ...c, totalsLabel: "Ergebnis" } : i === last ? { ...c, totalsFunction: "sum" } : c));
        writeTotals(sheet, x);
      } else {
        for (let c = x.ref.c0; c <= x.ref.c1; c++) sheet.set(x.ref.r1, c, undefined);
        x.ref = { ...x.ref, r1: x.ref.r1 - 1 };
        x.totals = false;
        x.columns = x.columns.map(({ totalsFunction: _f, totalsLabel: _l, ...c }) => { void _f; void _l; return c; });
      }
    });
  }

  /** The total row's cells from the columns' functions (TEILERGEBNIS with Excel's 10x codes). */
  function writeTotals(sheet: Sheet, t: SheetTable): void {
    const code: Record<string, number> = { average: 101, count: 103, countNums: 102, max: 104, min: 105, stdDev: 107, sum: 109, var: 110 };
    const data = tableData(t);
    t.columns.forEach((c, i) => {
      const col = t.ref.c0 + i;
      const range = rangeName({ r0: data.r0, c0: col, r1: data.r1, c1: col });
      if (c.totalsFunction && code[c.totalsFunction]) sheet.set(t.ref.r1, col, { v: null, f: `SUBTOTAL(${code[c.totalsFunction]},${range})` });
      else if (c.totalsLabel !== undefined) sheet.set(t.ref.r1, col, { v: c.totalsLabel });
      else sheet.set(t.ref.r1, col, undefined);
    });
  }

  function rename(name: string): boolean {
    const t = current();
    if (!t) return false;
    if (!validTableName(name)) { void messageBox("RaabeCalc", "Der Tabellenname ist ungültig. Er muss mit einem Buchstaben oder Unterstrich beginnen und darf keine Leerzeichen enthalten."); return false; }
    if (allNames().some((n) => n.toLowerCase() === name.toLowerCase() && n !== t.name)) { void messageBox("RaabeCalc", "Dieser Name wird bereits verwendet."); return false; }
    edit("Tabellenname", (x) => { x.name = name; });
    return true;
  }

  async function resizeDialog(): Promise<void> {
    const t = current();
    if (!t) return;
    const input = h("input", { type: "text", value: rangeName(t.ref).replace(/([A-Z]+)(\d+)/g, "$$$1$$$2"), style: "width:200px" }) as HTMLInputElement;
    const ok = await showDialog({ title: "Tabellengröße ändern", body: h("div", null, h("label", { class: "field" }, "Wählen Sie den neuen Datenbereich für die Tabelle aus:", input)),
      buttons: [{ label: "OK", value: true, primary: true }, { label: "Abbrechen", value: false }], cancel: false });
    if (!ok) return;
    const r = parseRange(input.value.replace(/\$/g, ""));
    if (!r || r.r0 !== t.ref.r0 || r.r1 < r.r0 + (t.header ? 1 : 0)) { await messageBox("RaabeCalc", "Die Kopfzeile muss in derselben Zeile bleiben, und der neue Bereich muss die Tabelle überlappen."); return; }
    const sheet = app.sheet;
    edit("Tabellengröße ändern", (x) => {
      x.ref = r;
      const names = columnNames(sheet, r, x.header);
      x.columns = names.map((name) => x.columns.find((c) => c.name === name) ?? { name });
      if (x.header) names.forEach((n, i) => { if (sheet.value(r.r0, r.c0 + i) !== n) sheet.set(r.r0, r.c0 + i, { ...(sheet.get(r.r0, r.c0 + i) ?? {}), v: n }); });
    });
  }

  /** In Bereich konvertieren: the cells keep the table's look as their own formatting (Excel). */
  async function convert(): Promise<void> {
    const t = current();
    if (!t) return;
    const ok = await showDialog({ title: "RaabeCalc", body: h("p", null, "Möchten Sie die Tabelle in einen normalen Bereich umwandeln?"),
      buttons: [{ label: "Ja", value: true, primary: true }, { label: "Nein", value: false }], cancel: false });
    if (!ok) return;
    const sheet = app.sheet;
    app.change("In Bereich konvertieren", () => {
      app.history.whole(sheet);
      for (let r = t.ref.r0; r <= t.ref.r1; r++) for (let c = t.ref.c0; c <= t.ref.c1; c++) {
        const cell = sheet.get(r, c) ?? { v: null };
        const st = app.wb.styles.get(cell.s);
        const look = styleFor(r, c, st);
        if (look === st) continue;
        const s = app.wb.styles.derive(cell.s, (x) => { x.fill = look.fill; x.font = { ...look.font }; x.borders = { ...look.borders }; });
        sheet.set(r, c, { ...cell, s });
      }
      sheet.tables = sheet.tables!.filter((x) => x !== t);
      if (!sheet.tables.length) sheet.tables = undefined;
    });
  }

  function setStyle(name: string): void { edit("Tabellenformatvorlage", (x) => { x.style = name; }); }

  /** Tabellenformatvorlagen: Hell, Mittel, Dunkel as tiles drawn from the style data. */
  function stylesPanel(close: () => void, pick: (name: string) => void = setStyle): HTMLElement {
    const group = (prefix: string, title: string) => {
      const tiles = TABLE_STYLE_DEFS.filter((d) => d.name.startsWith(prefix)).map((d) => {
        const fake: SheetTable = { name: "x", ref: { r0: 0, c0: 0, r1: 4, c1: 4 }, header: true, totals: false, columns: [], style: d.name,
          firstColumn: false, lastColumn: false, rowStripes: true, columnStripes: false, filterButton: false };
        const cells: HTMLElement[] = [];
        for (let r = 0; r <= 4; r++) for (let c = 0; c <= 4; c++) {
          const lk = look(fake, r, c);
          const b = lk?.borders ?? {};
          const line = (s?: BorderSide) => (s ? `1px solid ${s.color ?? "#000"}` : "0");
          cells.push(h("i", { style: `background:${lk?.fill ?? "#fff"};border-top:${line(b.top)};border-bottom:${line(b.bottom)};border-left:${line(b.left)};border-right:${line(b.right)}` }));
        }
        return h("button", { class: "ts-tile", "data-tip": d.local.replace(/^Tabellenformat - /, ""), "aria-label": d.local, onclick: () => { close(); pick(d.name); } }, h("span", { class: "ts-grid" }, ...cells));
      });
      return [h("div", { class: "palette-title" }, title), h("div", { class: "ts-tiles" }, ...tiles)];
    };
    return h("div", { class: "tile-panel ts-panel" }, ...group("TableStyleLight", "Hell"), ...group("TableStyleMedium", "Mittel"), ...group("TableStyleDark", "Dunkel"));
  }

  return { styleFor, tableAt: (row: number, col: number) => tableAt(app.sheet, row, col), current, create, setOption, setTotals, rename, resizeDialog, convert, setStyle, stylesPanel, look };
}
