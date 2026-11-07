// Tabellenentwurf > Effekte (D-072): Zellenabschrägung, Schatten, Spiegelung, as PowerPoint writes them. The
// shadow and reflection belong to the table (a:tblPr > a:effectLst, before the table style), the bevel to each
// cell being worked on (a:tcPr > a:cell3D with a:bevel and the light rig the schema requires). The presets are
// Office's (effect-presets.ts, the same galleries as Formeffekte); the gallery names are PowerPoint's.
// The table's shadow and reflection are drawn (render/table.ts); the cell bevels are written, not drawn.

import type { MenuEntry } from "@raabe/shell";
import { PptxDom } from "@raabe/ooxml";
import type { PresentApp } from "./app";
import { BEVELS, REFLECTIONS, SHADOWS, type PicturePreset } from "./effect-presets";

const { kid, kids, make, fragment } = PptxDom;

function tableEl(app: PresentApp): { tbl: Element; id: number } | null {
  const v = app.editor.selectedViews.find((x) => x.frame === "table");
  const el = v ? app.editor.shapeEl(v.id) : undefined;
  const tbl = el?.getElementsByTagNameNS(PptxDom.NS_A, "tbl")[0];
  return v && tbl ? { tbl, id: v.id } : null;
}

/** The preset's effect elements (outerShdw, innerShdw, reflection) from its spPr XML. */
function presetFx(doc: Document, p: PicturePreset, names: string[]): Element[] {
  const wrap = fragment(doc, `<a:x xmlns:a="${PptxDom.NS_A}">${p.xml}</a:x>`)[0];
  return Array.from(wrap?.getElementsByTagNameNS(PptxDom.NS_A, "*") ?? []).filter((e) => names.includes(e.localName) && e.parentElement?.localName === "effectLst");
}

const FX_ORDER = ["blur", "fillOverlay", "glow", "innerShdw", "outerShdw", "prstShdw", "reflection", "softEdge"];

function tableFx(app: PresentApp, names: string[], p: PicturePreset | null, label: string): void {
  const t = tableEl(app);
  if (!t) return;
  app.editor.edit(label, () => {
    const doc = t.tbl.ownerDocument;
    let pr = kid(t.tbl, "tblPr");
    if (!pr) { pr = make(doc, "a:tblPr"); t.tbl.insertBefore(pr, t.tbl.firstChild); }
    let fx = kid(pr, "effectLst");
    if (!fx) { fx = make(doc, "a:effectLst"); pr.insertBefore(fx, kids(pr).find((c) => ["tableStyle", "tableStyleId", "extLst"].includes(c.localName)) ?? null); }
    for (const c of kids(fx)) if (names.includes(c.localName)) c.remove();
    for (const e of p ? presetFx(doc, p, names) : []) fx.insertBefore(doc.importNode(e, true), kids(fx).find((c) => FX_ORDER.indexOf(c.localName) > FX_ORDER.indexOf(e.localName)) ?? null);
    if (!fx.children.length) fx.remove();
  });
}

/** Zellenabschrägung: the cells being edited (the cell range, the cell) or the whole table. */
function cellBevel(app: PresentApp, p: PicturePreset | null, label: string, cells: () => Element[]): void {
  const bevel = p ? fragment(document.implementation.createDocument(null, null), `<a:x xmlns:a="${PptxDom.NS_A}">${p.xml}</a:x>`)[0]?.getElementsByTagNameNS(PptxDom.NS_A, "bevelT")[0] : undefined;
  app.editor.edit(label, () => {
    for (const tc of cells()) {
      let pr = kid(tc, "tcPr");
      if (!pr) { pr = make(tc.ownerDocument, "a:tcPr"); tc.append(pr); }
      kid(pr, "cell3D")?.remove();
      if (!bevel) continue;
      const attrs = ["w", "h", "prst"].filter((a) => bevel.hasAttribute(a)).map((a) => ` ${a}="${bevel.getAttribute(a)}"`).join("");
      const c3 = fragment(tc.ownerDocument, `<a:cell3D><a:bevel${attrs}/><a:lightRig rig="threePt" dir="t"/></a:cell3D>`)[0];
      // After the cell lines, before the fill (CT_TableCellProperties).
      pr.insertBefore(c3, kids(pr).find((c) => !["lnL", "lnR", "lnT", "lnB", "lnTlToBr", "lnBlToTr"].includes(c.localName)) ?? null);
    }
  });
}

/** The Effekte menu of Tabellenentwurf (PowerPoint's three entries). */
export function tableEffectsMenu(app: PresentApp, cells: () => Element[]): MenuEntry[] {
  const list = (presets: readonly (PicturePreset & { id: string })[], run: (p: PicturePreset | null, label: string) => void): MenuEntry[] =>
    presets.map((p, i) => ({ label: p.label, run: () => run(i === 0 ? null : p, p.label) }));
  return [
    { label: "Zellenabschrägung", submenu: () => list(BEVELS, (p, l) => cellBevel(app, p, `Zellenabschrägung: ${l}`, cells)) },
    { label: "Schatten", submenu: () => list(SHADOWS, (p, l) => tableFx(app, ["outerShdw", "innerShdw", "prstShdw"], p, `Schatten: ${l}`)) },
    { label: "Spiegelung", submenu: () => list(REFLECTIONS, (p, l) => tableFx(app, ["reflection"], p, `Spiegelung: ${l}`)) },
  ];
}
