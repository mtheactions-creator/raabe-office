// WordArt (D-072): PowerPoint's 20 WordArt styles of WordArt-Formate > Schnellformatvorlagen, the 14 of
// Einfügen > WordArt among them, named as PowerPoint names them with the design's colour words. Each style
// is run properties (outline, fill, effects), applied to the selected text or the selected shapes' text.
// Bevels are not drawn (no 3-D): those styles keep their fill.

import { h, type Control } from "@raabe/shell";
import { PptxDom } from "@raabe/ooxml";
import type { PresentApp } from "./app";
import { colorWord } from "./quickstyles";

const { kids, fragment } = PptxDom;

const c = (v: string, mods = "") => `<a:schemeClr val="${v}">${mods}</a:schemeClr>`;
const solid = (v: string, mods = "") => `<a:solidFill>${c(v, mods)}</a:solidFill>`;
const ln = (v: string, w = 12700) => `<a:ln w="${w}">${solid(v)}</a:ln>`;
const grad = (v: string) => `<a:gradFill><a:gsLst><a:gs pos="0">${c(v, '<a:lumMod val="60000"/><a:lumOff val="40000"/>')}</a:gs><a:gs pos="100000">${c(v, '<a:lumMod val="75000"/>')}</a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill>`;
const patt = (prst: string, fg: string) => `<a:pattFill prst="${prst}"><a:fgClr>${c(fg)}</a:fgClr><a:bgClr>${c("bg1")}</a:bgClr></a:pattFill>`;
const SHADOW = '<a:outerShdw blurRad="38100" dist="38100" dir="2700000" algn="tl"><a:srgbClr val="000000"><a:alpha val="43137"/></a:srgbClr></a:outerShdw>';
const hard = (v: string) => `<a:outerShdw dist="38100" dir="2700000" algn="tl" rotWithShape="0">${c(v)}</a:outerShdw>`;
const INNER = '<a:innerShdw blurRad="63500" dist="50800" dir="13500000"><a:srgbClr val="000000"><a:alpha val="50000"/></a:srgbClr></a:innerShdw>';
const REFLECT = '<a:reflection blurRad="6350" stA="53000" endA="300" endPos="35500" dir="5400000" sy="-90000" algn="bl" rotWithShape="0"/>';
const glow = (v: string) => `<a:glow rad="63500">${c(v, '<a:alpha val="40000"/>')}</a:glow>`;
const fx = (...e: string[]) => `<a:effectLst>${e.join("")}</a:effectLst>`;

/** Name templates: {tx1}, {bg2}, {aN} become the design's colour word with PowerPoint's suffix. */
const STYLES: [string, string][] = [
  ["Füllung: {tx1}; Schatten", solid("tx1") + fx(SHADOW)],
  ["Füllung: {a1}; Schatten", solid("accent1") + fx(SHADOW)],
  ["Füllung: {a2}; Kontur: {a2}", ln("accent2") + solid("accent2", '<a:lumMod val="40000"/><a:lumOff val="60000"/>')],
  ["Füllung: Weiß; Kontur: {a5}; Schatten", ln("accent5") + solid("bg1") + fx(SHADOW)],
  ["Füllung: {a4}; Sanfte Abschrägung", solid("accent4")],
  ["Farbverlauf, Grau", grad("tx1")],
  ["Farbverlauf: {a5}; Spiegelung", grad("accent5") + fx(REFLECT)],
  ["Farbverlauf: {a4}; Kontur: {a4}", ln("accent4") + grad("accent4")],
  ["Füllung: Weiß; Kontur: {a1}; Leuchteffekt: {a1}", ln("accent1") + solid("bg1") + fx(glow("accent1"))],
  ["Füllung: {a3}; Steile Abschrägung", solid("accent3")],
  ["Füllung: {tx1}; Kontur: Weiß, Hintergrundfarbe 1, Harter Schatten: Weiß, Hintergrundfarbe 1", ln("bg1") + solid("tx1") + fx(hard("bg1"))],
  ["Füllung: {tx1}; Kontur: Weiß, Hintergrundfarbe 1; Harter Schatten: {a5}", ln("bg1") + solid("tx1") + fx(hard("accent5"))],
  ["Füllung: {a5}; Kontur: Weiß, Hintergrundfarbe 1; Harter Schatten: {a5}", ln("bg1") + solid("accent5") + fx(hard("accent5"))],
  ["Füllung: Weiß; Kontur: {a2}; Harter Schatten: {a2}", ln("accent2") + solid("bg1") + fx(hard("accent2"))],
  ["Füllung: {bg2}; Innerer Schatten", solid("bg2") + fx(INNER)],
  ["Musterfüllung: Weiß, Streifen diagonal dunkel nach oben; Schatten", ln("tx1", 9525) + patt("dkUpDiag", "tx1") + fx(SHADOW)],
  ["Musterfüllung: {a3}, Dünner horizontaler Streifen; Innerer Schatten", patt("ltHorz", "accent3") + fx(INNER)],
  ["Musterfüllung: {a1}, 50 %; Harter Schatten: {a1}", patt("pct50", "accent1") + fx(hard("accent1"))],
  ["Musterfüllung: {a5}, Streifen diagonal hell nach unten; Kontur: {a5}", ln("accent5") + patt("ltDnDiag", "accent5")],
  ["Musterfüllung: {tx2:plain}, Streifen diagonal dunkel nach oben; Harter Schatten", patt("dkUpDiag", "tx2") + fx(hard("tx1"))],
];
/** Einfügen > WordArt offers these of them (recorded). */
const INSERT = [0, 1, 2, 3, 4, 5, 6, 7, 9, 13, 14, 15, 16, 19];

export function wordArtNames(colors: Record<string, string>): string[] {
  const word = (ref: string) => colorWord(colors[ref] ?? "000000");
  return STYLES.map(([n]) => n.replace(/\{(tx1|tx2|bg2|a\d)(:plain)?\}/g, (_m, r: string, plain?: string) => {
    if (plain) return word(r);
    if (r === "tx1") return `${word("tx1")}, Textfarbe 1`;
    if (r === "bg2") return `${word("bg2")}, Hintergrundfarbe 2`;
    return `${word(`accent${r.slice(1)}`)}, Akzentfarbe ${r.slice(1)}`;
  }));
}

/** Puts a style's outline, fill and effects into a run's rPr (schema order: ln, fill, effects first). */
export function applyWordArt(rPr: Element, i: number): void {
  for (const x of kids(rPr)) if (["ln", "noFill", "solidFill", "gradFill", "pattFill", "blipFill", "grpFill", "effectLst", "effectDag"].includes(x.localName)) x.remove();
  const els = fragment(rPr.ownerDocument, STYLES[i][1]);
  for (const el of els.reverse()) rPr.insertBefore(el, rPr.firstChild);
}

export function styleXml(i: number): string { return STYLES[i][1]; }

const colorsOf = (app: PresentApp) => (app.slidePath() ? app.deck.view(app.slidePath()!).theme.colors : {}) as Record<string, string>;

function tile(app: PresentApp, i: number): HTMLElement {
  const colors = colorsOf(app);
  const hex = (ref: string) => `#${colors[ref] ?? "000000"}`;
  const xml = STYLES[i][1];
  const fill = /pattFill|gradFill/.test(xml) ? undefined : (xml.match(/<a:solidFill><a:schemeClr val="(\w+)"/) ?? [])[1];
  const fillRef = fill ?? (xml.match(/<a:(?:fgClr|gs pos="0")>?<a:schemeClr val="(\w+)"/) ?? [])[1] ?? "tx1";
  const line = (xml.match(/<a:ln[^>]*><a:solidFill><a:schemeClr val="(\w+)"/) ?? [])[1];
  const a = h("span", { class: "wa-a" }, "A");
  a.style.cssText = `color:${hex(fillRef)};${line ? `-webkit-text-stroke:1px ${hex(line)};` : ""}${/outerShdw/.test(xml) ? "text-shadow:2px 2px 1px rgba(0,0,0,.4);" : ""}`;
  return a;
}

function apply(app: PresentApp, i: number): void {
  app.formatText("WordArt-Formatvorlage", (r) => applyWordArt(r, i));
}

/** WordArt-Formate > Schnellformatvorlagen in the ribbon. */
export function wordArtTiles(app: PresentApp): Control {
  return {
    type: "tiles", cmd: "wordart.quickStyles", key: "C", count: STYLES.length, visible: 3, size: [44, 44], cls: "wa-tiles",
    render: (i) => tile(app, i),
    tip: (i) => wordArtNames(colorsOf(app))[i],
    pick: (i) => apply(app, i),
    redraw: () => app.designName(),
    panel: (close) => panel(app, close, STYLES.map((_, i) => i), (i) => apply(app, i)),
  };
}

/** Einfügen > WordArt: the 14 styles; the new shape takes the chosen one. */
export function wordArtInsertPanel(app: PresentApp, close: () => void, insert: (rpr: string) => void): HTMLElement {
  return panel(app, close, INSERT, (i) => insert(STYLES[i][1]));
}

function panel(app: PresentApp, close: () => void, which: number[], pick: (i: number) => void): HTMLElement {
  const names = wordArtNames(colorsOf(app));
  const grid = h("div", { class: "wa-grid" });
  for (const i of which) grid.append(h("button", { class: "wa-item", "aria-label": names[i], "data-tip": names[i], onclick: () => { close(); pick(i); } }, tile(app, i)));
  return h("div", { class: "menu wa-panel" }, grid);
}
