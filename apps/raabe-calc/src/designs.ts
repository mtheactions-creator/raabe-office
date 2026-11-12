// Seitenlayout > Designs (D-083): Designs, Farben and Schriftarten as in Excel. A design sets the workbook's
// theme colours and its heading and body fonts; cells whose colours are design colours (picked from the
// palette's Designfarben or read so from a file) and whose fonts are the theme's (Textkörper, Überschriften)
// follow, as do table styles, which are drawn from the theme already.

import { DESIGN_GROUPS, OFFICE_COLOR_SETS, OFFICE_DESIGNS, OFFICE_FONT_SETS, SCHOOL_DESIGNS, h } from "@raabe/shell";
import { THEME_KEYS, type CellStyle, type Sheet, type ThemeRef } from "@raabe/sheet-engine";
import { resolveXlsxColor } from "@raabe/ooxml";
import type { CalcApp } from "./app";

/** The theme index of a palette key (Hintergrund 1 = bg1 = lt1 …). */
const INDEX: Record<string, number> = { bg1: 0, lt1: 0, tx1: 1, dk1: 1, bg2: 2, lt2: 2, tx2: 3, dk2: 3, accent1: 4, accent2: 5, accent3: 6, accent4: 7, accent5: 8, accent6: 9, hlink: 10, folHlink: 11 };

/** Excel's palette tints: a share of 32767 (heller 80 % is 26213/32767 in Excel's files). */
const tintOf = (share: number) => Math.sign(share) * Math.floor(Math.abs(share) * 32767) / 32767;

export function designsFeature(app: CalcApp) {
  const colors = () => Object.fromEntries(Object.entries(app.wb.theme).map(([k, v]) => [k, v.replace("#", "")]));

  /** "#RRGGBB" of a design colour in the workbook's current theme. */
  function resolve(ref: ThemeRef, theme = colors()): string {
    return resolveXlsxColor({ theme: String(ref.t), ...(ref.tint ? { tint: String(ref.tint) } : {}) }, { theme, indexed: [] }) ?? "#000000";
  }

  /**
   * A palette pick: "RRGGBB" or "#RRGGBB" for a standard or own colour, or a design colour as the palette
   * hands it over ("accent1", "accent1,20000,80000" for heller 80 %, "accent1,75000" for dunkler 25 %).
   */
  function pick(value: string): { hex: string; theme?: ThemeRef } | null {
    if (!value) return null;
    if (/^#?[0-9A-Fa-f]{6}$/.test(value)) return { hex: "#" + value.replace("#", "").toUpperCase() };
    const [key, mod, off] = value.split(",");
    const t = INDEX[key];
    if (t === undefined) return null;
    let ref: ThemeRef = { t };
    if (off !== undefined) ref = { t, tint: tintOf(Number(off) / 100000) };
    else if (mod !== undefined) ref = { t, tint: tintOf(-(1 - Number(mod) / 100000)) };
    return { hex: resolve(ref), theme: ref };
  }

  /**
   * A style with the current design applied (null when nothing in it follows the design). A design colour
   * counts only while the style still shows it in the old design: a colour set some other way since drops it.
   */
  function restyled(s: CellStyle, old?: Record<string, string>): CellStyle | null {
    let changed = false;
    const n = structuredClone(s);
    if (n.font.theme) {
      if (old && n.font.color !== resolve(n.font.theme, old)) { delete n.font.theme; changed = true; }
      else { const c = resolve(n.font.theme); if (c !== n.font.color) { n.font.color = c; changed = true; } }
    }
    if (n.fillTheme) {
      if (old && n.fill !== resolve(n.fillTheme, old)) { delete n.fillTheme; changed = true; }
      else { const c = resolve(n.fillTheme); if (c !== n.fill) { n.fill = c; changed = true; } }
    }
    if (n.font.scheme === "minor" || n.font.scheme === "major") {
      const f = n.font.scheme === "major" ? app.wb.themeFonts.major : app.wb.themeFonts.minor;
      if (f !== n.font.name) { n.font.name = f; changed = true; }
    }
    return changed ? n : null;
  }

  /** Every cell, column and row style brought to the current design (one undo step with the design itself). */
  function restyleAll(old: Record<string, string>): void {
    // The Standard style (cells without their own format) follows in place: it is derived from the design alone.
    app.wb.styles.remap((s) => { if (s === app.wb.styles.list[0]) { const n = restyled(s); if (n) Object.assign(s, n); } });
    const map = new Map<number, number>();
    const idFor = (id: number | undefined): number | undefined => {
      if (!id) return id;
      if (!map.has(id)) { const n = restyled(app.wb.styles.get(id), old); map.set(id, n ? app.wb.styles.id(n) : id); }
      return map.get(id);
    };
    for (const sh of app.wb.sheets) {
      let touched = false;
      const mark = (s: Sheet) => { if (!touched) { app.history.whole(s); touched = true; } };
      for (const [k, c] of sh.cells) { const s = idFor(c.s); if (s !== c.s) { mark(sh); sh.cells.set(k, { ...c, s }); } }
      for (const [k, c] of sh.cols) { const s = idFor(c.s); if (s !== c.s) { mark(sh); sh.cols.set(k, { ...c, s }); } }
      for (const [k, r] of sh.rows) { const s = idFor(r.s); if (s !== r.s) { mark(sh); sh.rows.set(k, { ...r, s }); } }
    }
  }

  /** The Standard style after undo/redo of a design (it is changed in place). */
  function syncDefault(): void {
    const s = app.wb.styles.list[0];
    const n = restyled(s);
    if (n) app.wb.styles.remap((x) => { if (x === s) Object.assign(x, n); });
  }

  function setDesign(label: string, change: () => void): void {
    app.change(label, () => {
      app.history.design();
      const old = colors();
      change();
      restyleAll(old);
    });
    app.grid.contentChanged();
  }

  const withHash = (c: Readonly<Record<string, string>>) => Object.fromEntries(THEME_KEYS.map((k) => [k, "#" + (c[k] ?? "000000").replace("#", "").toUpperCase()]));

  /** Office's designs and our own (the school's, D-196: the same colours and fonts as RaabePresent's designs). */
  const ALL_DESIGNS = [...OFFICE_DESIGNS, ...SCHOOL_DESIGNS];

  function applyDesign(name: string): void {
    const d = ALL_DESIGNS.find((x) => x.name === name);
    if (!d) return;
    setDesign("Design", () => {
      app.wb.theme = withHash(d.colors);
      app.wb.themeFonts = { major: d.major, minor: d.minor };
      app.wb.themeNames = { theme: d.name, colors: d.name, fonts: d.name };
    });
  }

  function applyColors(name: string): void {
    const c = OFFICE_COLOR_SETS.find((x) => x.name === name);
    if (!c) return;
    setDesign("Designfarben", () => {
      app.wb.theme = withHash(c.colors);
      app.wb.themeNames = { ...app.wb.themeNames, colors: c.name };
    });
  }

  function applyFonts(name: string): void {
    const f = OFFICE_FONT_SETS.find((x) => x.name === name);
    if (!f) return;
    setDesign("Designschriftarten", () => {
      app.wb.themeFonts = { major: f.major, minor: f.minor };
      app.wb.themeNames = { ...app.wb.themeNames, fonts: f.name };
    });
  }

  const sw = (hex: string) => h("i", { class: "dz-sw", style: `background:#${hex.replace("#", "")}` });

  /** Designs: a tile per design ("Aa" in its heading font over its accent colours); Office's, then ours by group. */
  function designsPanel(close: () => void): HTMLElement {
    const tile = (d: (typeof ALL_DESIGNS)[number]) => h("button", {
      class: "dz-tile" + (d.name === app.wb.themeNames.theme ? " on" : ""), "data-tip": d.name, "aria-label": d.name,
      onclick: () => { close(); applyDesign(d.name); },
    }, h("span", { class: "dz-aa", style: `font-family:'${d.major}';color:#${d.colors.dk2}` }, "Aa"),
    h("span", { class: "dz-bars" }, ...["accent1", "accent2", "accent3", "accent4"].map((k) => sw(d.colors[k as keyof typeof d.colors]))));
    const groups = DESIGN_GROUPS.flatMap((g) => [h("div", { class: "palette-title" }, g), h("div", { class: "dz-tiles" }, ...SCHOOL_DESIGNS.filter((d) => d.group === g).map(tile))]);
    return h("div", { class: "tile-panel dz-panel dz-scroll" }, h("div", { class: "palette-title" }, "Office"), h("div", { class: "dz-tiles" }, ...OFFICE_DESIGNS.map(tile)), ...groups);
  }

  /** Farben: one row per colour set (its eight colours and its name). */
  function colorsPanel(close: () => void): HTMLElement {
    const rows = OFFICE_COLOR_SETS.map((c) => h("button", {
      class: "dz-row" + (c.name === app.wb.themeNames.colors ? " on" : ""), "aria-label": c.name,
      onclick: () => { close(); applyColors(c.name); },
    }, h("span", { class: "dz-sws" }, ...["dk2", "lt2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6"].map((k) => sw(c.colors[k]))), h("span", null, c.name)));
    return h("div", { class: "tile-panel dz-list" }, h("div", { class: "palette-title" }, "Office"), ...rows);
  }

  /** Schriftarten: one row per font set ("Aa", heading font, body font). */
  function fontsPanel(close: () => void): HTMLElement {
    const rows = OFFICE_FONT_SETS.map((f) => h("button", {
      class: "dz-row dz-font" + (f.name === app.wb.themeNames.fonts ? " on" : ""), "data-tip": f.label, "aria-label": f.label,
      onclick: () => { close(); applyFonts(f.name); },
    }, h("span", { class: "dz-aa" }, "Aa"), h("span", { class: "dz-names" }, h("b", null, f.name), h("span", { style: `font-family:'${f.major}'` }, f.major), h("span", { style: `font-family:'${f.minor}'` }, f.minor))));
    return h("div", { class: "tile-panel dz-list" }, h("div", { class: "palette-title" }, "Office"), ...rows);
  }

  /** The palette's design for the colour buttons of the ribbon. */
  const paletteTheme = () => ({ colors: colors() });

  return { pick, resolve, applyDesign, applyColors, applyFonts, designsPanel, colorsPanel, fontsPanel, syncDefault, paletteTheme };
}
