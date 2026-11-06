// Formenarten > Schnellformatvorlagen (D-072): PowerPoint's 77 entries, 11 rows of 7 colours (Dunkel 1
// and Akzent 1-6). The six rows of Designformatvorlagen are shape styles (p:style with references into the
// theme's format matrix, as PowerPoint writes them), the five rows of Voreinstellungen set the fill and the
// outline on the shape itself. Names as PowerPoint's (recorded), the colour word from the design's colour
// (PowerPoint names its colours after the design too: "Dunkelblaugrün, Akzent 1" in Office, "Blau, Akzent 1"
// in Office 2013-2022). How each style looks in detail is our own reading of PowerPoint's gallery.

import { h } from "@raabe/shell";
import { PptxDom, type ShapeView } from "@raabe/ooxml";
import type { PresentApp } from "./app";

const { kid, kids, fragment } = PptxDom;

const COLS = ["dk1", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6"];
const clr = (v: string, mods = "") => `<a:schemeClr val="${v}">${mods}</a:schemeClr>`;
const style = (ln: [number, string], fill: [number, string], fx: [number, string], font: string) =>
  `<p:style><a:lnRef idx="${ln[0]}">${ln[1]}</a:lnRef><a:fillRef idx="${fill[0]}">${fill[1]}</a:fillRef><a:effectRef idx="${fx[0]}">${fx[1]}</a:effectRef><a:fontRef idx="minor">${font}</a:fontRef></p:style>`;
const onFill = (_c: string) => clr("lt1");

interface Row { label: string; suffix?: string; preset?: boolean; style: (c: string) => string; spPr?: (c: string) => string }

export const QUICK_ROWS: Row[] = [
  { label: "Farbige Kontur", style: (c) => style([2, clr(c)], [1, clr("lt1")], [0, clr(c)], clr(c)) },
  { label: "Farbige Füllung", style: (c) => style([2, clr(c, '<a:shade val="15000"/>')], [1, clr(c)], [0, clr(c)], onFill(c)) },
  { label: "Hell 1 Kontur, Farbige Füllung", style: (c) => style([3, clr("lt1")], [1, clr(c)], [1, clr(c)], onFill(c)) },
  { label: "Subtiler Effekt", style: (c) => style([1, clr(c)], [3, clr(c, '<a:lumMod val="20000"/><a:lumOff val="80000"/>')], [0, clr(c)], clr("dk1")) },
  { label: "Moderater Effekt", style: (c) => style([1, clr(c)], [2, clr(c)], [1, clr(c)], onFill(c)) },
  { label: "Intensiver Effekt", style: (c) => style([0, clr(c)], [3, clr(c)], [2, clr(c)], onFill(c)) },
  { label: "Transparent", preset: true, style: (c) => style([2, clr(c)], [1, clr(c)], [0, clr(c)], clr(c)), spPr: () => '<a:noFill/><a:ln><a:noFill/></a:ln>' },
  { label: "Transparent, Farbige Kontur", preset: true, style: (c) => style([2, clr(c)], [1, clr(c)], [0, clr(c)], clr(c)), spPr: (c) => `<a:noFill/><a:ln w="12700"><a:solidFill>${clr(c)}</a:solidFill></a:ln>` },
  { label: "Halbtransparent", suffix: ", Keine Kontur", preset: true, style: (c) => style([2, clr(c)], [1, clr(c)], [0, clr(c)], clr("dk1")), spPr: (c) => `<a:solidFill>${clr(c, '<a:alpha val="40000"/>')}</a:solidFill><a:ln><a:noFill/></a:ln>` },
  { label: "Farbige Füllung", suffix: ", Keine Kontur", preset: true, style: (c) => style([2, clr(c)], [1, clr(c)], [0, clr(c)], onFill(c)), spPr: (c) => `<a:solidFill>${clr(c)}</a:solidFill><a:ln><a:noFill/></a:ln>` },
  { label: "Farbverlauf", suffix: ", Keine Kontur", preset: true, style: (c) => style([2, clr(c)], [1, clr(c)], [0, clr(c)], onFill(c)),
    spPr: (c) => `<a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0">${clr(c, '<a:lumMod val="60000"/><a:lumOff val="40000"/>')}</a:gs><a:gs pos="100000">${clr(c)}</a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill><a:ln><a:noFill/></a:ln>` },
];

/** A German colour word for a design colour, as PowerPoint's names go ("Blau", "Dunkelblaugrün", "Gold"). */
export function colorWord(hex: string): string {
  const n = parseInt(hex, 16), r = ((n >> 16) & 255) / 255, g = ((n >> 8) & 255) / 255, b = (n & 255) / 255;
  const max = Math.max(r, g, b), min = Math.min(r, g, b), l = (max + min) / 2, d = max - min;
  if (l < 0.1) return "Schwarz";
  if (l > 0.93) return "Weiß";
  const s = d === 0 ? 0 : d / (1 - Math.abs(2 * l - 1));
  if (s < 0.15) return l < 0.35 ? "Dunkelgrau" : l > 0.75 ? "Hellgrau" : "Grau";
  let hue = max === r ? ((g - b) / d) % 6 : max === g ? (b - r) / d + 2 : (r - g) / d + 4;
  hue = (hue * 60 + 360) % 360;
  const dark = l < 0.33;
  const word = hue < 15 || hue >= 345 ? "Rot" : hue < 40 ? "Orange" : hue < 52 ? "Gold" : hue < 70 ? "Gelb" : hue < 165 ? "Grün"
    : hue < 205 ? (dark ? "Blaugrün" : "Türkis") : hue < 255 ? "Blau" : hue < 290 ? "Lila" : "Pflaume";
  return dark && word !== "Pflaume" ? `Dunkel${word.toLowerCase()}` : word;
}

export function colName(col: number, colors: Record<string, string>): string {
  const ref = COLS[col];
  return `${colorWord(colors[ref === "dk1" ? "tx1" : ref] ?? colors[ref] ?? "000000")}, ${col === 0 ? "Dunkel 1" : `Akzent ${col}`}`;
}

export function quickName(row: Row, col: number, colors: Record<string, string>): string {
  return `${row.label} - ${colName(col, colors)}${row.suffix ?? ""}`;
}

/** Applies a gallery entry to a shape (p:sp): its style, and for Voreinstellungen its own fill and outline. */
export function applyQuickStyle(sp: Element, row: Row, col: number): void {
  const doc = sp.ownerDocument, c = COLS[col];
  const spPr = kid(sp, "spPr");
  if (!spPr) return;
  for (const x of kids(spPr)) if (["noFill", "solidFill", "gradFill", "pattFill", "blipFill", "grpFill", "ln", "effectLst", "effectDag"].includes(x.localName)) x.remove();
  if (row.spPr) {
    const geom = kids(spPr).find((x) => x.localName === "prstGeom" || x.localName === "custGeom");
    let after: Element | null = geom ?? kid(spPr, "xfrm") ?? null;
    for (const x of fragment(doc, row.spPr(c))) { spPr.insertBefore(x, after ? after.nextSibling : spPr.firstChild); after = x; }
  }
  kid(sp, "style")?.remove();
  const st = fragment(doc, row.style(c))[0];
  spPr.after(st);
}

/** The gallery: Designformatvorlagen and Voreinstellungen, each tile drawn in the design's colours. */
export function quickStylesPanel(app: PresentApp, close: () => void): HTMLElement {
  const view = app.slidePath() ? app.deck.view(app.slidePath()!) : null;
  const colors = view?.theme.colors ?? {};
  const hex = (ref: string) => `#${colors[ref === "dk1" ? "tx1" : ref === "lt1" ? "bg1" : ref] ?? colors[ref] ?? "000000"}`;
  const wrap = h("div", { class: "menu quick-styles" });
  const section = (title: string, rows: Row[], start: number) => {
    wrap.append(h("div", { class: "gal-head" }, title));
    const grid = h("div", { class: "qs-grid" });
    rows.forEach((row, ri) => COLS.forEach((c, ci) => {
      const name = quickName(row, ci, colors);
      const tile = h("button", { class: "qs-tile", "aria-label": name, "data-tip": name, onclick: () => { close(); apply(app, start + ri, ci); } }, "Abc");
      const col = hex(c);
      const look: Record<string, string> = {
        "Farbige Kontur": `background:${hex("lt1")};border:2px solid ${col};color:${col}`,
        "Farbige Füllung": row.suffix ? `background:${col};border:2px solid transparent;color:#fff` : `background:${col};border:2px solid ${col};color:#fff`,
        "Hell 1 Kontur, Farbige Füllung": `background:${col};border:2px solid ${hex("lt1")};outline:1px solid ${col};color:#fff`,
        "Subtiler Effekt": `background:color-mix(in srgb, ${col} 20%, white);border:1px solid ${col};color:${hex("dk1")}`,
        "Moderater Effekt": `background:linear-gradient(${col}, color-mix(in srgb, ${col} 80%, black));border:1px solid ${col};color:#fff;box-shadow:0 1px 2px rgba(0,0,0,.4)`,
        "Intensiver Effekt": `background:linear-gradient(color-mix(in srgb, ${col} 70%, white), ${col});border:none;color:#fff;box-shadow:0 2px 3px rgba(0,0,0,.5)`,
        "Transparent": `background:transparent;border:1px dashed transparent;color:${col}`,
        "Transparent, Farbige Kontur": `background:transparent;border:1px solid ${col};color:${col}`,
        "Halbtransparent": `background:color-mix(in srgb, ${col} 40%, transparent);border:none;color:${hex("dk1")}`,
        "Farbverlauf": `background:linear-gradient(color-mix(in srgb, ${col} 60%, white), ${col});border:none;color:#fff`,
      };
      tile.style.cssText = look[row.label] ?? "";
      grid.append(tile);
    }));
    wrap.append(grid);
  };
  section("Designformatvorlagen", QUICK_ROWS.slice(0, 6), 0);
  section("Voreinstellungen", QUICK_ROWS.slice(6), 6);
  return wrap;
}

function apply(app: PresentApp, row: number, col: number): void {
  const ed = app.editor;
  const views = ed.selectedViews.filter((v: ShapeView) => v.kind === "sp");
  if (!views.length) return;
  ed.edit("Schnellformatvorlage", () => { for (const v of views) { const el = ed.shapeEl(v.id); if (el) applyQuickStyle(el, QUICK_ROWS[row], col); } });
}
