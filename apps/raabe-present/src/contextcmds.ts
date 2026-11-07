// What the contextual tabs' commands do (D-070): picture colour, transparency and corrections as
// DrawingML blip effects (a:grayscl, a:biLevel, a:lum, a:alphaModFix, as PowerPoint writes them),
// Alternativtext (descr), table cells (alignment, vertical anchor, rows and columns at the cell,
// distributing, sizes), charts through chartops.ts.

import { colorXml } from "./colors";
import { closePopup, h, showDialog, type Control, type MenuEntry } from "@raabe/shell";
import { PptxDom, type ShapeView } from "@raabe/ooxml";
import type { PresentApp } from "./app";
import { LINE_WEIGHTS } from "./lines";
import { colorWord, quickStyleTiles, quickStylesPanel } from "./quickstyles";
import { tableStyleTiles, tableStylesPanel } from "./tablestyles";
import { wordArtTiles } from "./wordart";
import { picStyleTiles, picStylesPanel } from "./picstyles";
import type { ContextTabHost } from "./contexttabs";
import * as T from "./textops";
import * as S from "./shapeops";
import * as CH from "./chartops";
import * as TC from "./tablecells";
import { artisticEntries, colorEntries, sharpenEntries } from "./picfx";
import { textEffectEntries, textThreeD, transformMenu } from "./shapefx";

const { kid, kids, make, num, fragment } = PptxDom;

export function contextHost(app: PresentApp): ContextTabHost {
  const ctx = app.ribbonContext();
  const spin = (value: () => string, commit: (v: string) => void, step: (d: 1 | -1) => void, label: string, cmd: string, icon: string, key: string): Control =>
    ({ type: "spinner", cmd, label, icon, key, value, commit, step });
  return {
    shapeGallery: (c) => app.shapeGallery(c),
    color: (cur, pick, cmd, key, label, extra) => ({ type: "color", cmd, key, palette: "font", current: cur, pick, showLabel: label, extra, theme: () => app.paletteTheme() }),
    outlineExtra: () => app.outlineExtra(),
    quickStyles: (c) => quickStylesPanel(app, c),
    quickStyleTiles: () => quickStyleTiles(app),
    picStyleTiles: () => picStyleTiles(app),
    tableStyleTiles: () => tableStyleTiles(app),
    tableBorders: () => tableBorders(app),
    textEffects: () => textEffects(app),
    chartElements: () => CH.chartElementsMenu(app),
    editShapeMenu: () => [{ label: "Form ändern", flyout: () => app.shapeGallery(() => closePopup(), (g) => app.changeShape(g)) },
      { label: "Punkte bearbeiten", run: () => {}, disabled: true }, { label: "Verbindungen umleiten", run: () => {}, disabled: true }],
    wordArtTiles: () => wordArtTiles(app),
    // Textrichtung and Zellenbegrenzungen of the cells (the one being edited, the range, else all), as PowerPoint's entries.
    cellDirections: () => ([["Horizontal", null], ["Text um 90° drehen", "vert"], ["Text um 270° drehen", "vert270"], ["Gestapelt", "wordArtVert"]] as [string, string | null][])
      .map(([label, v]): MenuEntry => ({ label, run: () => cellPr(app, "Textrichtung", (pr) => { if (v) pr.setAttribute("vert", v); else pr.removeAttribute("vert"); }) }))
      .concat([{ label: "Weitere Optionen...", run: () => void app.commands.run("draw.format") }]),
    cellMargins: () => [...([["Normal", 91440, 45720], ["Keine", 0, 0], ["Schmal", 45720, 45720], ["Breit", 137160, 137160]] as [string, number, number][])
      .map(([label, lr, tb]) => ({ label, run: () => cellPr(app, "Zellenbegrenzungen", (pr) => { for (const [a, v] of [["marL", lr], ["marR", lr], ["marT", tb], ["marB", tb]] as [string, number][]) { if (label === "Normal") pr.removeAttribute(a); else pr.setAttribute(a, String(v)); } }) })),
      "-" as const, { cmd: "table.customMargins" }],
    shapeTiles: () => app.shapeTiles(),
    tableStyles: (c) => tableStylesPanel(app, c),
    picStyles: (c) => picStylesPanel(app, c),
    textOutlineColor: () => app.textOutlineColor(), setTextOutline: (hex) => app.setTextOutline({ color: hex }), textOutlineExtra: () => app.textOutlineExtra(),
    fillColor: ctx.fillColor, setFill: ctx.setFill, lineColor: ctx.lineColor, setLine: ctx.setLine, fontColor: ctx.fontColor, setFontColor: ctx.setFontColor,
    effects: () => ctx.menus.effects,
    alignMenu: () => [{ cmd: "arrange.alignLeft" }, { cmd: "arrange.alignCenter" }, { cmd: "arrange.alignRight" }, "-", { cmd: "arrange.alignTop" }, { cmd: "arrange.alignMiddle" }, { cmd: "arrange.alignBottom" }],
    sizeSpin: (k) => spin(() => sizeText(app, k), (v) => setSize(app, k, cmText(v)), (d) => { const v = app.editor.selectedViews[0]; if (v) setSize(app, k, Math.max(0, (k === "h" ? v.h : v.w) + d * 36000)); },
      k === "h" ? "Formenhöhe:" : "Formenbreite:", k === "h" ? "shape.heightBox" : "shape.widthBox", k === "h" ? "shape-height" : "shape-width", k === "h" ? "H" : "RR"),
    pictureColor: () => pictureColor(app),
    pictureArtistic: () => artisticEntries(app),
    pictureTransparency: () => [0, 15, 30, 50, 65, 80, 95].map((p) => ({ label: `Transparenz: ${p}%`, run: () => blipEffect(app, "Transparenz", (blip) => { remove(blip, "alphaModFix"); if (p) blip.prepend(make(blip.ownerDocument, "a:alphaModFix", { amt: (100 - p) * 1000 })); }) })),
    pictureCorrections: () => {
      const out: MenuEntry[] = [...sharpenEntries(app), { heading: "Helligkeit/Kontrast" }];
      // PowerPoint's names: "Helligkeit: 0 % (Normal) Kontrast: +20 %".
      const pct = (v: number) => `${v > 0 ? "+" : ""}${v} %${v === 0 ? " (Normal)" : ""}`;
      for (const c of [-40, -20, 0, 20, 40]) for (const br of [-40, -20, 0, 20, 40]) out.push({ label: `Helligkeit: ${pct(br)} Kontrast: ${pct(c)}`, run: () => blipEffect(app, "Korrekturen", (blip) => { remove(blip, "lum"); if (br || c) blip.append(make(blip.ownerDocument, "a:lum", { bright: br * 1000, contrast: c * 1000 })); }) });
      out.push("-", { label: "Bildkorrekturoptionen...", run: () => void app.commands.run("draw.format") });
      return out;
    },
    lineWeights: () => LINE_WEIGHTS.map(([label, w]) => ({ label, run: () => app.setLineWidth(w) })),
    fillExtra: () => app.fillExtra(),
    cropMenu: () => app.cropMenu(),
    penStyles: () => penMenus(app).styles(), penWeights: () => penMenus(app).weights(),
    penColor: () => `#${app.tablePen.color.length === 6 ? app.tablePen.color : "000000"}`,
    setPenColor: (c) => { app.tablePen.color = c.replace("#", "").toUpperCase(); app.editor.endText(); app.editor.tool = { kind: "tableDraw" }; app.editor.render(); app.updateChrome(); },
    chartKinds: () => CH.CHART_KINDS.map(([label, k]) => ({ label, run: () => CH.changeType(app, k) })),
    quickLayouts: () => Array.from({ length: 11 }, (_, i) => ({ label: `Layout ${i + 1}`, run: () => CH.quickLayout(app, i + 1) })),
    chartColors: () => [{ heading: "Bunt" }, ...[10, 11, 12, 13].map((k) => ({ label: `Farbige Palette ${k - 9}`, run: () => CH.setColors(app, k) })),
      { heading: "Einfarbig" }, ...Array.from({ length: 13 }, (_, i) => ({ label: `Monochrome Palette ${i + 1}`, run: () => CH.setColors(app, 14 + i) }))],
    tableSpin: (k) => {
      const labels = { row: "Höhe:", col: "Breite:", h: "Höhe:", w: "Breite:" }, cmds = { row: "table.rowHeight", col: "table.colWidth", h: "table.height", w: "table.width" };
      return spin(() => tableSizeText(app, k), (v) => setTableSize(app, k, cmText(v)), (d) => setTableSize(app, k, Math.max(36000, tableSize(app, k) + d * 36000)), labels[k], cmds[k], k === "row" || k === "h" ? "shape-height" : "shape-width", { row: "P", col: "B", h: "X", w: "LB" }[k]);
    },
  };
}

const cm = (emu: number) => `${(emu / 360000).toFixed(2).replace(".", ",")} cm`;
const cmText = (v: string) => Math.round(Number(v.replace(",", ".").replace(/[^\d.]/g, "")) * 360000);
function remove(el: Element, name: string): void { for (const c of kids(el, name)) c.remove(); }

function sizeText(app: PresentApp, k: "h" | "w"): string {
  const v = app.editor.selectedViews[0];
  return v ? cm(k === "h" ? v.h : v.w) : "";
}

function setSize(app: PresentApp, k: "h" | "w", emu: number): void {
  if (!(emu > 0)) return;
  const views = app.editor.selectedViews;
  app.editor.edit("Größe", () => {
    for (const v of views) {
      const el = app.editor.shapeEl(v.id);
      if (!el) continue;
      // Pictures keep their proportions (Seitenverhältnis sperren is on for them, as in PowerPoint).
      const keep = v.kind === "pic";
      const r = v.h ? v.w / v.h : 1;
      S.setXfrm(el, { x: v.x, y: v.y, w: k === "w" ? emu : keep ? emu * r : v.w, h: k === "h" ? emu : keep ? emu / r : v.h });
    }
  });
}

function blipEffect(app: PresentApp, label: string, fn: (blip: Element) => void): void {
  const views = app.editor.selectedViews.filter((v) => v.kind === "pic");
  app.editor.edit(label, () => {
    for (const v of views) {
      const blip = PptxDom.at(app.editor.shapeEl(v.id)!, "blipFill/blip");
      if (blip) fn(blip);
    }
  });
}

/** Farbe > Neu einfärben as PowerPoint's (names recorded; the theme colours named after the design, as there). */
function pictureColor(app: PresentApp): MenuEntry[] {
  const clear = (blip: Element) => { for (const n of ["grayscl", "biLevel", "duotone", "clrChange", "lum"]) remove(blip, n); };
  const set = (xml: string) => blipEffect(app, "Farbe", (b) => { clear(b); for (const x of fragment(b.ownerDocument, xml)) b.append(x); });
  const colors = app.slidePath() ? app.deck.view(app.slidePath()!).theme.colors : {};
  const word = (ref: string) => colorWord(colors[ref] ?? "000000");
  const dark = (ref: string) => `<a:duotone><a:schemeClr val="${ref}"><a:shade val="45000"/><a:satMod val="135000"/></a:schemeClr><a:prstClr val="white"/></a:duotone>`;
  const light = (ref: string) => `<a:duotone><a:prstClr val="black"/><a:schemeClr val="${ref}"><a:tint val="45000"/><a:satMod val="400000"/></a:schemeClr></a:duotone>`;
  const accents = [1, 2, 3, 4, 5, 6];
  return [...colorEntries(app), { heading: "Neu einfärben" },
    { label: "Keine Neueinfärbung", run: () => blipEffect(app, "Farbe", clear) },
    { label: "Graustufe", run: () => set("<a:grayscl/>") },
    { label: "Sepia", run: () => set('<a:duotone><a:prstClr val="black"/><a:srgbClr val="D9C3A5"><a:tint val="50000"/><a:satMod val="180000"/></a:srgbClr></a:duotone>') },
    { label: "Ausgeblichen", run: () => set('<a:lum bright="70000" contrast="-70000"/>') },
    ...[25, 50, 75].map((t) => ({ label: `Schwarzweiß: ${t} %`, run: () => set(`<a:biLevel thresh="${t * 1000}"/>`) })),
    { label: `${word("tx2")}, Textfarbe 2 dunkel`, run: () => set(dark("tx2")) },
    ...accents.map((n) => ({ label: `${word(`accent${n}`)}, Akzentfarbe ${n} dunkel`, run: () => set(dark(`accent${n}`)) })),
    { label: `${word("bg2")}, Hintergrundfarbe 2 hell`, run: () => set(light("bg2")) },
    ...accents.map((n) => ({ label: `${word(`accent${n}`)}, Akzentfarbe ${n} hell`, run: () => set(light(`accent${n}`)) })),
    "-", { label: "Bildfarboptionen...", run: () => void app.commands.run("draw.format") }];
}

// ----- tables --------------------------------------------------------------------------------------

function tableOf(app: PresentApp): { v: ShapeView; el: Element; tbl: Element } | null {
  const v = app.editor.selectedViews.find((x) => x.frame === "table");
  const el = v ? app.editor.shapeEl(v.id) : undefined;
  const tbl = el?.getElementsByTagNameNS(PptxDom.NS_A, "tbl")[0];
  return v && el && tbl ? { v, el, tbl } : null;
}

/** The cells a command acts on: the range, the one being edited, else all. */
function cells(app: PresentApp, tbl: Element): Element[] {
  const c = app.editor.text?.cell, x = app.editor.text?.range;
  const rows = kids(tbl, "tr");
  if (x) return rows.flatMap((r, ri) => kids(r, "tc").filter((_, ci) => TC.inRange(x, ri, ci)));
  if (c) { const tc = kids(rows[c.r], "tc")[c.c]; return tc ? [tc] : []; }
  return rows.flatMap((r) => kids(r, "tc"));
}

export function tableAlign(app: PresentApp, algn: "l" | "ctr" | "r"): void {
  const t = tableOf(app);
  if (!t) return;
  app.editor.edit("Ausrichten", () => { for (const tc of cells(app, t.tbl)) { const body = kid(tc, "txBody"); if (body) T.paraProps(body, 0, T.paras(body).length - 1, (p) => p.setAttribute("algn", algn)); } });
}

/** A change to the a:tcPr of the cells a command acts on. */
function cellPr(app: PresentApp, label: string, fn: (pr: Element) => void): void {
  const t = tableOf(app);
  if (!t) return;
  app.editor.edit(label, () => {
    for (const tc of cells(app, t.tbl)) {
      let pr = kid(tc, "tcPr");
      if (!pr) { pr = make(tc.ownerDocument, "a:tcPr"); tc.append(pr); }
      fn(pr);
    }
  });
}

/**
 * Tabellenentwurf > Rahmen (PowerPoint's 13 entries): the cells' own lines (a:lnL ... a:lnBlToTr) over the
 * cell range being edited, the cell, or the whole table, in the pen colour.
 */
const LINE_TAGS = ["lnL", "lnR", "lnT", "lnB", "lnTlToBr", "lnBlToTr"];
export function setCellLine(pr: Element, tag: string, color: string | null, pen: { w: number; dash: string } = { w: 12700, dash: "solid" }): void {
  kid(pr, tag)?.remove();
  const el = fragment(pr.ownerDocument, color ? `<a:${tag} w="${pen.w}" cap="flat" cmpd="sng" algn="ctr"><a:solidFill>${colorXml(color)}</a:solidFill><a:prstDash val="${pen.dash}"/><a:round/></a:${tag}>` : `<a:${tag} w="${pen.w}" cmpd="sng"><a:noFill/></a:${tag}>`)[0];
  const after = LINE_TAGS.slice(LINE_TAGS.indexOf(tag) + 1);
  const before = kids(pr).find((c) => after.includes(c.localName) || !LINE_TAGS.includes(c.localName)) ?? null;
  pr.insertBefore(el, before);
}

/** Tabellenentwurf > Stiftart and Stiftstärke: PowerPoint's entries; choosing one takes up the table pen. */
export const PEN_STYLES: [string, string | null][] = [["Kein Rahmen", null], ["Durchgezogene Linie", "solid"], ["Gestrichelt", "sysDash"], ["Punktiert", "sysDot"],
  ["Strich und Punkt", "sysDashDot"], ["Strich und 2 Punkte", "sysDashDotDot"], ["Punktiert (weiter Abstand)", "dot"], ["Gestrichelt (weiter Abstand)", "dash"],
  ["Lang gestrichelt (w. Abstand)", "lgDash"], ["Strich und Punkt (w. Abstand)", "dashDot"], ["Lng. Strich u.Punkt (w. Abst.)", "lgDashDot"]];
export const PEN_WEIGHTS: [string, number][] = [["¼ pt", 3175], ["½ pt", 6350], ["¾ pt", 9525], ["1 pt", 12700], ["1½ pt", 19050], ["2¼ pt", 28575], ["3 pt", 38100], ["4½ pt", 57150], ["6 pt", 76200]];

export function penMenus(app: PresentApp): { styles: () => MenuEntry[]; weights: () => MenuEntry[] } {
  const take = () => { app.editor.endText(); app.editor.tool = { kind: "tableDraw" }; app.editor.render(); app.updateChrome(); };
  return {
    styles: () => PEN_STYLES.map(([label, dash]) => ({ label, checked: app.tablePen.dash === dash, run: () => { app.tablePen.dash = dash; take(); } })),
    weights: () => PEN_WEIGHTS.map(([label, w]) => ({ label, checked: app.tablePen.w === w, run: () => { app.tablePen.w = w; take(); } })),
  };
}

export function tableBorders(app: PresentApp): MenuEntry[] {
  const apply = (label: string, sides: (edge: { top: boolean; bottom: boolean; left: boolean; right: boolean }) => Partial<Record<string, boolean | null>>) => () => {
    const t = tableOf(app);
    if (!t) return;
    const rows = kids(t.tbl, "tr"), cols = kids(rows[0], "tc").length;
    const x = app.editor.text?.range, c = app.editor.text?.cell;
    const blk = x ?? (c ? { r0: c.r, c0: c.c, r1: c.r, c1: c.c } : { r0: 0, c0: 0, r1: rows.length - 1, c1: cols - 1 });
    const color = app.ribbonContext().lineColor().replace("#", "").toUpperCase();
    app.editor.edit(label, () => {
      rows.forEach((tr, r) => kids(tr, "tc").forEach((tc, ci) => {
        if (r < blk.r0 || r > blk.r1 || ci < blk.c0 || ci > blk.c1) return;
        let pr = kid(tc, "tcPr");
        if (!pr) { pr = make(tc.ownerDocument, "a:tcPr"); tc.append(pr); }
        const want = sides({ top: r === blk.r0, bottom: r === blk.r1, left: ci === blk.c0, right: ci === blk.c1 });
        for (const [tag, on] of Object.entries(want)) if (on !== undefined) setCellLine(pr, tag, on ? color : null);
      }));
    });
  };
  const all = { lnL: true, lnR: true, lnT: true, lnB: true };
  return [
    { label: "Kein Rahmen", run: apply("Rahmen", () => ({ lnL: false, lnR: false, lnT: false, lnB: false })) },
    { label: "Alle Rahmenlinien", run: apply("Rahmen", () => all) },
    { label: "Rahmenlinien außen", run: apply("Rahmen", (e) => ({ lnL: e.left || undefined, lnR: e.right || undefined, lnT: e.top || undefined, lnB: e.bottom || undefined })) },
    { label: "Rahmenlinien innen", run: apply("Rahmen", (e) => ({ lnL: !e.left || undefined, lnR: !e.right || undefined, lnT: !e.top || undefined, lnB: !e.bottom || undefined })) },
    { label: "Rahmenlinie oben", run: apply("Rahmen", (e) => ({ lnT: e.top || undefined })) },
    { label: "Rahmenlinie unten", run: apply("Rahmen", (e) => ({ lnB: e.bottom || undefined })) },
    { label: "Rahmenlinie links", run: apply("Rahmen", (e) => ({ lnL: e.left || undefined })) },
    { label: "Rahmenlinie rechts", run: apply("Rahmen", (e) => ({ lnR: e.right || undefined })) },
    { label: "Horizontale Rahmenlinie innen", run: apply("Rahmen", (e) => ({ lnT: !e.top || undefined, lnB: !e.bottom || undefined })) },
    { label: "Vertikale Rahmenlinie innen", run: apply("Rahmen", (e) => ({ lnL: !e.left || undefined, lnR: !e.right || undefined })) },
    { label: "Rahmenlinien diagonal nach unten", run: apply("Rahmen", () => ({ lnTlToBr: true })) },
    { label: "Rahmenlinien diagonal nach oben", run: apply("Rahmen", () => ({ lnBlToTr: true })) },
  ];
}

/** Texteffekte (WordArt-Formate): Schatten, Spiegelung, Leuchten on the text's runs (a:effectLst in a:rPr). */
export function textEffects(app: PresentApp): MenuEntry[] {
  // Schatten, Spiegelung, Leuchten: Office's presets on the runs; 3-D and Transformieren on the text body (shapefx.ts).
  return [
    ...textEffectEntries(app),
    { label: "Abschrägung", submenu: () => textThreeD(app, "bevel") }, { label: "3D-Drehung", submenu: () => textThreeD(app, "rotation") }, { label: "Transformieren", submenu: () => transformMenu(app) },
  ];
}

export function tableAnchor(app: PresentApp, anchor: "t" | "ctr" | "b"): void {
  const t = tableOf(app);
  if (!t) return;
  app.editor.edit("Ausrichten", () => {
    for (const tc of cells(app, t.tbl)) {
      let pr = kid(tc, "tcPr");
      if (!pr) { pr = make(tc.ownerDocument, "a:tcPr"); tc.append(pr); }
      if (anchor === "t") pr.removeAttribute("anchor"); else pr.setAttribute("anchor", anchor);
    }
  });
}

/** Zeile oberhalb / unterhalb, Spalte links / rechts of the cell being edited. */
export function tableInsertAt(app: PresentApp, where: "above" | "below" | "left" | "right"): void {
  const t = tableOf(app);
  if (!t) return;
  const c = app.editor.text?.cell ?? { r: where === "above" ? 0 : kids(t.tbl, "tr").length - 1, c: where === "left" ? 0 : kids(kids(t.tbl, "tr")[0], "tc").length - 1 };
  app.editor.edit(where === "above" || where === "below" ? "Zeile einfügen" : "Spalte einfügen", () => {
    const rows = kids(t.tbl, "tr");
    const empty = (tc: Element) => { const x = tc.cloneNode(true) as Element; const b2 = kid(x, "txBody"); if (b2) T.replaceAll(b2, ""); x.removeAttribute("gridSpan"); x.removeAttribute("hMerge"); return x; };
    const ext = kid(kid(t.el, "xfrm"), "ext");
    if (where === "above" || where === "below") {
      const ref = rows[c.r];
      const nr = ref.cloneNode(false) as Element;
      for (const tc of kids(ref, "tc")) nr.append(empty(tc));
      if (where === "above") ref.before(nr); else ref.after(nr);
      if (ext) ext.setAttribute("cy", String((num(ext, "cy") ?? 0) + (num(ref, "h") ?? 370840)));
    } else {
      const col = kids(kid(t.tbl, "tblGrid"), "gridCol")[c.c];
      const nc = col.cloneNode(false) as Element;
      if (where === "left") col.before(nc); else col.after(nc);
      for (const tr of rows) { const tc = kids(tr, "tc")[c.c]; if (!tc) continue; const x = empty(tc); if (where === "left") tc.before(x); else tc.after(x); }
      if (ext) ext.setAttribute("cx", String((num(ext, "cx") ?? 0) + (num(col, "w") ?? 0)));
    }
  });
}

export function distribute(app: PresentApp, what: "rows" | "cols"): void {
  const t = tableOf(app);
  if (!t) return;
  app.editor.edit(what === "rows" ? "Zeilen verteilen" : "Spalten verteilen", () => {
    const list = what === "rows" ? kids(t.tbl, "tr") : kids(kid(t.tbl, "tblGrid"), "gridCol");
    const attr = what === "rows" ? "h" : "w";
    const total = list.reduce((s, x) => s + (num(x, attr) ?? 0), 0);
    for (const x of list) x.setAttribute(attr, String(Math.round(total / list.length)));
  });
}

function tableSize(app: PresentApp, k: "row" | "col" | "h" | "w"): number {
  const t = tableOf(app);
  if (!t) return 0;
  const c = app.editor.text?.cell ?? { r: 0, c: 0 };
  if (k === "row") return num(kids(t.tbl, "tr")[c.r], "h") ?? 0;
  if (k === "col") return num(kids(kid(t.tbl, "tblGrid"), "gridCol")[c.c], "w") ?? 0;
  return k === "h" ? t.v.h : t.v.w;
}

function tableSizeText(app: PresentApp, k: "row" | "col" | "h" | "w"): string {
  return tableOf(app) ? cm(tableSize(app, k)) : "";
}

function setTableSize(app: PresentApp, k: "row" | "col" | "h" | "w", emu: number): void {
  const t = tableOf(app);
  if (!t || !(emu > 0)) return;
  const c = app.editor.text?.cell ?? { r: 0, c: 0 };
  app.editor.edit("Größe", () => {
    const rows = kids(t.tbl, "tr"), cols = kids(kid(t.tbl, "tblGrid"), "gridCol");
    const ext = kid(kid(t.el, "xfrm"), "ext");
    if (k === "row") rows[c.r]?.setAttribute("h", String(emu));
    else if (k === "col") cols[c.c]?.setAttribute("w", String(emu));
    else if (k === "h") { const f = emu / Math.max(1, t.v.h); for (const r of rows) r.setAttribute("h", String(Math.round((num(r, "h") ?? 0) * f))); }
    else { const f = emu / Math.max(1, t.v.w); for (const g of cols) g.setAttribute("w", String(Math.round((num(g, "w") ?? 0) * f))); }
    if (ext) { ext.setAttribute("cx", String(cols.reduce((s, g) => s + (num(g, "w") ?? 0), 0))); ext.setAttribute("cy", String(rows.reduce((s, r) => s + (num(r, "h") ?? 0), 0))); }
  });
}

/** Alternativtext: the description PowerPoint keeps in cNvPr/@descr (and "Als dekorativ markieren"). */
export async function altText(app: PresentApp): Promise<void> {
  const v = app.editor.selectedViews[0];
  if (!v) return;
  const c = S.cNvPr(v.el);
  const text = h("textarea", { style: "width:320px;height:90px" }) as HTMLTextAreaElement;
  text.value = c?.getAttribute("descr") ?? "";
  const ok = await showDialog({ title: "Alternativtext", body: h("div", { class: "form" }, h("label", null, "Beschreiben Sie das Objekt für Menschen, die es nicht sehen können:"), text), buttons: [{ label: "OK", value: true, primary: true }, { label: "Abbrechen", value: false }], cancel: false });
  if (!ok) return;
  app.editor.edit("Alternativtext", () => { const el = S.cNvPr(app.editor.shapeEl(v.id)!); if (!el) return; if (text.value) el.setAttribute("descr", text.value); else el.removeAttribute("descr"); });
}

/** Zellen teilen: PowerPoint's dialog with Spaltenanzahl and Zeilenanzahl. */
async function splitCells(app: PresentApp): Promise<void> {
  const t = tableOf(app), cell = app.editor.text?.cell;
  if (!t || !cell) return;
  const field = (v: number) => h("input", { type: "number", min: "1", max: "75", value: String(v), style: "width:70px" }) as HTMLInputElement;
  const g = TC.grid(t.tbl), o = TC.origin(g, cell.r, cell.c), e = TC.extent(g, o.r, o.c);
  const merged = e.r1 > e.r0 || e.c1 > e.c0;
  const cols = field(merged ? e.c1 - e.c0 + 1 : 2), rows = field(merged ? e.r1 - e.r0 + 1 : 1);
  const ok = await showDialog({ title: "Zellen teilen", body: h("div", { class: "form" },
    h("label", null, h("u", null, "S"), "paltenanzahl: ", cols), h("label", null, h("u", null, "Z"), "eilenanzahl: ", rows)),
    buttons: [{ label: "OK", value: true, primary: true }, { label: "Abbrechen", value: false }], cancel: false });
  if (!ok) return;
  const nc = Math.max(1, Math.min(75, Math.round(Number(cols.value) || 1))), nr = Math.max(1, Math.min(75, Math.round(Number(rows.value) || 1)));
  app.editor.edit("Zellen teilen", () => TC.split(t.tbl, o.r, o.c, nc, nr));
  app.editor.enterText(t.v.id, undefined, o);
}

export function implementContext(app: PresentApp): void {
  const c = app.commands;
  const impl = (id: string, run: () => void | Promise<void>, enabled?: () => boolean, checked?: () => boolean) => c.implement(id, { run, enabled, checked });
  const shapes = () => app.editor.sel.length > 0;
  const pic = () => app.editor.selectedViews.some((v) => v.kind === "pic");
  const table = () => !!tableOf(app);
  const chart = () => CH.hasChart(app);
  impl("shape.altTextPane", () => altText(app), shapes);
  impl("shape.altText", () => altText(app), shapes);
  impl("shape.sizeDialog", () => void app.commands.run("draw.format"), shapes);
  impl("shape.heightBox", () => {}, shapes);
  impl("shape.widthBox", () => {}, shapes);
  impl("shape.quickStyles", () => {}, shapes);
  impl("wordart.quickStyles", () => {}, () => shapes() || !!app.editor.text);
  impl("wordart.fill", () => void app.commands.run("font.color"), shapes);
  impl("wordart.outline", () => app.setTextOutline({ color: app.textOutlineColor() === "none" ? "000000" : app.textOutlineColor() }), shapes);
  impl("pic.corrections", () => {}, pic);
  impl("pic.color", () => {}, pic);
  impl("pic.artistic", () => {}, pic);
  // Hintergrund entfernen: PowerPoint's automatic removal at once (its default rectangle, 10 % in from each side).
  impl("pic.removeBg", () => void import("./picfx").then((m) => m.setEffect(app, "backgroundRemoval", { name: "backgroundRemoval", attrs: { t: "10000", b: "90000", l: "10000", r: "90000" } }, "Hintergrund entfernen")), pic);
  impl("pic.transparency", () => {}, pic);
  impl("pic.border", () => void app.commands.run("draw.outline"), pic);
  impl("pic.effects", () => {}, pic);
  impl("table.alignLeft", () => tableAlign(app, "l"), table);
  impl("table.alignCenter", () => tableAlign(app, "ctr"), table);
  impl("table.alignRight", () => tableAlign(app, "r"), table);
  impl("table.top", () => tableAnchor(app, "t"), table);
  impl("table.middle", () => tableAnchor(app, "ctr"), table);
  impl("table.bottom", () => tableAnchor(app, "b"), table);
  impl("table.rowAbove", () => tableInsertAt(app, "above"), table);
  impl("table.rowBelow", () => tableInsertAt(app, "below"), table);
  impl("table.colLeft", () => tableInsertAt(app, "left"), table);
  impl("table.colRight", () => tableInsertAt(app, "right"), table);
  impl("table.insertAbove", () => tableInsertAt(app, "above"), table);
  impl("table.insertBelow", () => tableInsertAt(app, "below"), table);
  impl("table.insertLeft", () => tableInsertAt(app, "left"), table);
  impl("table.insertRight", () => tableInsertAt(app, "right"), table);
  impl("table.distRows", () => distribute(app, "rows"), table);
  impl("table.distCols", () => distribute(app, "cols"), table);
  impl("table.deleteMenu", () => {}, table);
  impl("table.select", () => {}, table);
  impl("table.styles", () => {}, table);
  impl("table.borders", () => {}, table);
  impl("chart.addElement", () => {}, chart);
  impl("shape.editShape", () => {}, () => app.editor.selectedViews.some((v) => v.kind === "sp"));
  impl("wordart.effects", () => {}, () => shapes() || !!app.editor.text);
  impl("table.direction", () => {}, table);
  impl("table.margins", () => {}, table);
  impl("pic.quickStyles", () => {}, pic);
  impl("pic.compress", () => import("./compress").then((m) => m.compressDialog(app)), pic);
  const editingCell = () => !!app.editor.text?.cell;
  // Tabelle zeichnen and Radierer stay on until Esc, another tool, or the button again.
  const tableTool = (kind: "tableDraw" | "tableEraser") => () => { app.editor.endText(); app.editor.tool = app.editor.tool?.kind === kind ? null : { kind }; app.editor.render(); app.updateChrome(); };
  impl("table.draw", tableTool("tableDraw"), () => !!app.slidePath(), () => app.editor.tool?.kind === "tableDraw");
  impl("table.eraser", tableTool("tableEraser"), () => !!app.slidePath(), () => app.editor.tool?.kind === "tableEraser");
  impl("table.merge", () => {
    const t = tableOf(app), x = app.editor.text?.range;
    if (!t || !x) return;
    app.editor.edit("Zellen verbinden", () => TC.merge(t.tbl, x));
    app.editor.enterText(t.v.id, undefined, { r: x.r0, c: x.c0 });
  }, () => !!app.editor.text?.range);
  impl("table.split", () => splitCells(app), editingCell);
  impl("table.selectCol", () => app.editor.selectLine("col"), editingCell);
  impl("table.selectRow", () => app.editor.selectLine("row"), editingCell);
  impl("table.shading", () => {}, table);
  impl("table.penStyle", () => {}, table);
  impl("table.penWeight", () => {}, table);
  impl("table.penColor", () => {}, table);
  for (const id of ["table.rowHeight", "table.colWidth", "table.height", "table.width"]) impl(id, () => {}, table);
  impl("table.gridlines", () => { document.body.classList.toggle("table-gridlines"); }, table, () => document.body.classList.contains("table-gridlines"));
  impl("chart.changeType", () => {}, chart);
  impl("chart.quickLayout", () => {}, chart);
  impl("chart.colors", () => {}, chart);
  impl("chart.switchRowCol", () => CH.switchRowCol(app), chart);
  impl("chart.editData", () => CH.editData(app), chart);
  impl("chart.selectData", () => CH.editData(app), chart);
}
