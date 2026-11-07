// Tabellenentwurf > Tabellenformate (D-072): PowerPoint's 74 entries, named as recorded, in its families
// (Designformatvorlage 1 and 2, Helle 1-3, Mittlere 1-4, Dunkle 1 and 2, each plain and with Akzent 1-6).
// Each style is written into the presentation's tableStyles.xml when it is used, as PowerPoint writes its
// built-in ones; their look is our reading of PowerPoint's gallery. PowerPoint's own style ids are only
// known for its default (Mittlere Formatvorlage 2 - Akzent 1); the others get ids of our own, so
// PowerPoint shows them as custom styles of the presentation.

import { h } from "@raabe/shell";
import { PptxDom } from "@raabe/ooxml";
import type { PresentApp } from "./app";
import { DEFAULT_TABLE_STYLE_ID, DEFAULT_TABLE_STYLE_XML } from "./render/table";

const { kid, kids, fragment } = PptxDom;

export interface TableStyleDef { name: string; id: string; xml: string; look: { head: string; headText: string; band: string; body: string; line: string } }

const SIDES = ["left", "right", "top", "bottom", "insideH", "insideV"];
const c = (v: string, mod = "") => `<a:schemeClr val="${v}">${mod}</a:schemeClr>`;
const tint = (v: string, t: number) => c(v, `<a:tint val="${t}"/>`);
const shade = (v: string, s: number) => c(v, `<a:shade val="${s}"/>`);
const ln = (side: string, color: string | null, w = 12700) => `<a:${side}><a:ln w="${w}" cmpd="sng">${color ? `<a:solidFill>${color}</a:solidFill>` : "<a:noFill/>"}</a:ln></a:${side}>`;
const bdr = (sides: Record<string, string | null | [string, number]>) => `<a:tcBdr>${Object.entries(sides).map(([s, v]) => Array.isArray(v) ? ln(s, v[0], v[1]) : ln(s, v)).join("")}</a:tcBdr>`;
const fill = (color: string | null) => (color ? `<a:fill><a:solidFill>${color}</a:solidFill></a:fill>` : "");
const tx = (color: string, bold = false) => `<a:tcTxStyle${bold ? ' b="on"' : ""}><a:fontRef idx="minor"><a:prstClr val="black"/></a:fontRef>${color}</a:tcTxStyle>`;
const part = (name: string, body: { text?: string; bdr?: string; fill?: string }) => `<a:${name}>${body.text ?? ""}<a:tcStyle>${body.bdr ?? "<a:tcBdr/>"}${body.fill ?? ""}</a:tcStyle></a:${name}>`;

/** A style id of our own, the same for the same name (so a file keeps one entry per style). */
function ownId(name: string): string {
  let a = 0x811c9dc5, b = 0x1234567;
  for (const ch of name) { a = Math.imul(a ^ ch.charCodeAt(0), 0x01000193) >>> 0; b = Math.imul(b + ch.charCodeAt(0), 0x5bd1e995) >>> 0; }
  const hx = (n: number, l: number) => (n >>> 0).toString(16).toUpperCase().padStart(8, "0").slice(0, l);
  return `{${hx(a, 8)}-${hx(b, 4)}-4${hx(a ^ b, 3)}-A${hx(b >>> 3, 3)}-${hx(a >>> 1, 8)}${hx(b >>> 5, 4)}}`;
}

function style(name: string, parts: string): string {
  return `<a:tblStyle styleId="${ownId(name)}" styleName="${name}">${parts}</a:tblStyle>`;
}

type Family = (acc: string, name: string) => { xml: string; look: TableStyleDef["look"] };

const FAMILIES: [string, Family][] = [
  ["Designformatvorlage 1", (a, n) => ({ xml: style(n,
    part("wholeTbl", { text: tx(c("dk1")), bdr: bdr(Object.fromEntries(SIDES.map((s) => [s, c(a)]))) }) +
    part("band1H", { fill: fill(tint(a, 20000)) }) + part("firstRow", { text: tx(c("lt1"), true), fill: fill(c(a)) }) + part("lastRow", { text: tx(c("dk1"), true), bdr: bdr({ top: [c(a), 38100] }) })),
    look: { head: a, headText: "lt1", band: `${a}:20`, body: "none", line: a } })],
  ["Designformatvorlage 2", (a, n) => ({ xml: style(n,
    part("wholeTbl", { text: tx(c("lt1")), bdr: bdr({ insideH: tint(a, 50000), insideV: tint(a, 50000) }), fill: fill(c(a)) }) +
    part("band1H", { fill: fill(tint(a, 80000)) }) + part("firstRow", { text: tx(c("lt1"), true), bdr: bdr({ bottom: [c("lt1"), 38100] }) })),
    look: { head: a, headText: "lt1", band: `${a}:80`, body: a, line: "lt1" } })],
  ["Helle Formatvorlage 1", (a, n) => ({ xml: style(n,
    part("wholeTbl", { text: tx(c("dk1")), bdr: bdr({ top: c(a), bottom: c(a) }) }) +
    part("band1H", { fill: fill(tint(a, 20000)) }) + part("firstRow", { text: tx(c("dk1"), true), bdr: bdr({ bottom: c(a) }) })),
    look: { head: "none", headText: "dk1", band: `${a}:20`, body: "none", line: a } })],
  ["Helle Formatvorlage 2", (a, n) => ({ xml: style(n,
    part("wholeTbl", { text: tx(c("dk1")), bdr: bdr({ left: c(a), right: c(a), top: c(a), bottom: c(a) }) }) +
    part("band1H", { bdr: bdr({ top: c(a), bottom: c(a) }) }) + part("firstRow", { text: tx(c("lt1"), true), fill: fill(c(a)) })),
    look: { head: a, headText: "lt1", band: "none", body: "none", line: a } })],
  ["Helle Formatvorlage 3", (a, n) => ({ xml: style(n,
    part("wholeTbl", { text: tx(c("dk1")), bdr: bdr(Object.fromEntries(SIDES.map((s) => [s, c(a)]))) }) +
    part("band1H", { fill: fill(tint(a, 20000)) }) + part("firstRow", { text: tx(c("dk1"), true), bdr: bdr({ bottom: [c(a), 25400] }) })),
    look: { head: "none", headText: "dk1", band: `${a}:20`, body: "none", line: a } })],
  ["Mittlere Formatvorlage 1", (a, n) => ({ xml: style(n,
    part("wholeTbl", { text: tx(c("dk1")), bdr: bdr({ left: c(a), right: c(a), top: c(a), bottom: c(a), insideH: c(a) }), fill: fill(c("lt1")) }) +
    part("band1H", { fill: fill(tint(a, 20000)) }) + part("firstRow", { text: tx(c("lt1"), true), fill: fill(c(a)) })),
    look: { head: a, headText: "lt1", band: `${a}:20`, body: "lt1", line: a } })],
  ["Mittlere Formatvorlage 2", (a, n) => ({ xml: style(n,
    part("wholeTbl", { text: tx(c("dk1")), bdr: bdr(Object.fromEntries(SIDES.map((s) => [s, c("lt1")]))), fill: fill(tint(a, 20000)) }) +
    part("band1H", { fill: fill(tint(a, 40000)) }) + part("firstRow", { text: tx(c("lt1"), true), bdr: bdr({ bottom: [c("lt1"), 38100] }), fill: fill(c(a)) })),
    look: { head: a, headText: "lt1", band: `${a}:40`, body: `${a}:20`, line: "lt1" } })],
  ["Mittlere Formatvorlage 3", (a, n) => ({ xml: style(n,
    part("wholeTbl", { text: tx(c("dk1")), bdr: bdr({ top: [c("dk1"), 25400], bottom: [c("dk1"), 25400] }), fill: fill(c("lt1")) }) +
    part("band1H", { fill: fill(tint("dk1", 20000)) }) + part("firstRow", { text: tx(c("lt1"), true), bdr: bdr({ bottom: [c("dk1"), 25400] }), fill: fill(c(a)) })),
    look: { head: a, headText: "lt1", band: "dk1:20", body: "lt1", line: "dk1" } })],
  ["Mittlere Formatvorlage 4", (a, n) => ({ xml: style(n,
    part("wholeTbl", { text: tx(c("dk1")), bdr: bdr(Object.fromEntries(SIDES.map((s) => [s, c(a)]))), fill: fill(tint(a, 20000)) }) +
    part("band1H", { fill: fill(tint(a, 40000)) }) + part("firstRow", { text: tx(c("dk1"), true), fill: fill(tint(a, 20000)) })),
    look: { head: `${a}:20`, headText: "dk1", band: `${a}:40`, body: `${a}:20`, line: a } })],
  ["Dunkle Formatvorlage 1", (a, n) => ({ xml: style(n,
    part("wholeTbl", { text: tx(c("lt1")), fill: fill(shade(a, 20000 + 60000)) }) +
    part("band1H", { fill: fill(shade(a, 60000)) }) + part("firstRow", { text: tx(c("lt1"), true), bdr: bdr({ bottom: [c("lt1"), 38100] }), fill: fill(c("dk1")) })),
    look: { head: "dk1", headText: "lt1", band: `${a}:s60`, body: `${a}:s80`, line: "lt1" } })],
];

const ACCENT_NAME = (n: number) => `Akzent ${n}`;

/** All 74 entries in PowerPoint's order. */
export function tableStyles(): TableStyleDef[] {
  const out: TableStyleDef[] = [];
  const add = (name: string, xml: string, look: TableStyleDef["look"], id?: string) => out.push({ name, id: id ?? ownId(name), xml, look });
  const none = (name: string, grid: boolean) => add(name, style(name, part("wholeTbl", { text: tx(c("tx1")), bdr: grid ? bdr(Object.fromEntries(SIDES.map((s) => [s, c("tx1")]))) : "<a:tcBdr/>" })),
    { head: "none", headText: "dk1", band: "none", body: "none", line: grid ? "dk1" : "none" });
  // Row 1 and 2 start with the two "Keine Formatvorlage" entries, then each family as in the gallery.
  for (const [fam, f] of FAMILIES) {
    if (fam === "Designformatvorlage 1") none("Keine Formatvorlage, kein Raster", false);
    if (fam === "Designformatvorlage 2") none("Keine Formatvorlage, Tabellenraster", true);
    const themed = fam.startsWith("Design");
    if (!themed) { const r = f("dk1", fam); add(fam, r.xml, r.look); }
    for (let n = 1; n <= 6; n++) {
      const name = `${fam} - ${ACCENT_NAME(n)}`;
      if (name === "Mittlere Formatvorlage 2 - Akzent 1") { out.push({ name, id: DEFAULT_TABLE_STYLE_ID, xml: DEFAULT_TABLE_STYLE_XML, look: f("accent1", name).look }); continue; }
      const r = f(`accent${n}`, name);
      add(name, r.xml, r.look);
    }
  }
  // Dunkle Formatvorlage 2: plain and three accent pairs.
  const dark2 = (name: string, a: string, b: string) => add(name, style(name,
    part("wholeTbl", { text: tx(c("dk1")), fill: fill(tint(a, 20000)) }) + part("band1H", { fill: fill(tint(a, 40000)) }) +
    part("firstRow", { text: tx(c("lt1"), true), fill: fill(c(b)) })), { head: b, headText: "lt1", band: `${a}:40`, body: `${a}:20`, line: "none" });
  dark2("Dunkle Formatvorlage 2", "dk1", "dk1");
  for (const [x, y] of [[1, 2], [3, 4], [5, 6]]) dark2(`Dunkle Formatvorlage 2 - Akzent ${x}/Akzent ${y}`, `accent${x}`, `accent${y}`);
  return out;
}

/** Applies a style to the selected tables: written into tableStyles.xml when missing, referred to by id. */
export function applyTableStyle(app: PresentApp, def: TableStyleDef): void {
  const ed = app.editor;
  const views = ed.selectedViews.filter((v) => v.frame === "table");
  if (!views.length) return;
  const pptx = app.deck.pptx;
  const stylesPath = pptx.related(pptx.presPath, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/tableStyles");
  const paths = [app.slidePath()!, ...(stylesPath ? [stylesPath] : [])];
  app.deck.change("Tabellenformatvorlage", paths, () => {
    if (stylesPath) {
      const doc = pptx.doc(stylesPath);
      if (!kids(doc.documentElement, "tblStyle").some((s) => s.getAttribute("styleId") === def.id)) doc.documentElement.append(...fragment(doc, def.xml));
    }
    for (const v of views) {
      const tbl = ed.shapeEl(v.id)?.getElementsByTagNameNS(PptxDom.NS_A, "tbl")[0];
      const pr = tbl && kid(tbl, "tblPr");
      if (!pr) continue;
      kid(pr, "tableStyleId")?.remove();
      const idEl = fragment(pr.ownerDocument, `<a:tableStyleId>${def.id}</a:tableStyleId>`)[0];
      pr.append(idEl);
    }
  });
  ed.render();
}

/** The gallery: one small table per style, drawn in the design's colours. */
export function tableStylesPanel(app: PresentApp, close: () => void): HTMLElement {
  const colors = app.slidePath() ? app.deck.view(app.slidePath()!).theme.colors : {};
  const col = (spec: string): string => {
    if (spec === "none") return "transparent";
    const [ref, amt] = spec.split(":");
    const hex = `#${colors[ref === "dk1" ? "tx1" : ref === "lt1" ? "bg1" : ref] ?? colors[ref] ?? "000000"}`;
    if (!amt) return hex;
    return amt.startsWith("s") ? `color-mix(in srgb, ${hex} ${amt.slice(1)}%, black)` : `color-mix(in srgb, ${hex} ${amt}%, white)`;
  };
  const wrap = h("div", { class: "menu table-styles" });
  const grid = h("div", { class: "ts-grid" });
  for (const def of tableStyles()) {
    const t = h("button", { class: "ts-tile", "aria-label": def.name, "data-tip": def.name, onclick: () => { close(); applyTableStyle(app, def); } });
    const line = col(def.look.line);
    for (let r = 0; r < 5; r++) {
      const row = h("div", { class: "ts-row" });
      row.style.background = r === 0 ? col(def.look.head) : r % 2 === 1 ? col(def.look.band) : col(def.look.body);
      row.style.borderBottom = `1px solid ${line}`;
      t.append(row);
    }
    grid.append(t);
  }
  wrap.append(grid);
  return wrap;
}

/** Tabellenentwurf > Tabellenformatvorlagen in the ribbon (seven visible, as PowerPoint's). */
export function tableStyleTiles(app: PresentApp): import("@raabe/shell").Control {
  const all = tableStyles();
  return {
    type: "tiles", cmd: "table.styles", key: "M", count: all.length, visible: 7, size: [62, 44], cls: "ts-tiles",
    render: (i) => tableTile(app, all[i]),
    tip: (i) => all[i].name,
    pick: (i) => applyTableStyle(app, all[i]),
    redraw: () => app.designName(),
    panel: (close) => tableStylesPanel(app, close),
  };
}

function tableTile(app: PresentApp, def: TableStyleDef): HTMLElement {
  const colors = app.slidePath() ? app.deck.view(app.slidePath()!).theme.colors : {};
  const col = (spec: string): string => {
    if (spec === "none") return "transparent";
    const [ref, amt] = spec.split(":");
    const hex = `#${colors[ref === "dk1" ? "tx1" : ref === "lt1" ? "bg1" : ref] ?? colors[ref] ?? "000000"}`;
    if (!amt) return hex;
    return amt.startsWith("s") ? `color-mix(in srgb, ${hex} ${amt.slice(1)}%, black)` : `color-mix(in srgb, ${hex} ${amt}%, white)`;
  };
  const t = h("div", { class: "ts-rtile" });
  for (let r = 0; r < 5; r++) {
    const row = h("div", { class: "ts-row" });
    row.style.background = r === 0 ? col(def.look.head) : r % 2 === 1 ? col(def.look.band) : col(def.look.body);
    row.style.borderBottom = `1px solid ${col(def.look.line)}`;
    t.append(row);
  }
  return t;
}
