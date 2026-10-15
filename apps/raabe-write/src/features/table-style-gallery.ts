// Tabellenentwurf > Tabellenformatvorlagen as Word 365 has it (Step 4c D, recorded 2026-10-15):
// the 105 styles of Word's gallery in its order and groups (Einfache Tabellen, Rastertabellen,
// Listentabellen), each drawn from its definition with the table's Tabellenformatoptionen; the
// ribbon shows a row of them, hovering one previews it on the table (Livevorschau).

import { h, type Control } from "@raabe/shell";
import { tableCellLook, type BorderSpec, type Editor, type StyleDef, type TableDef, type TableStyleDef } from "@raabe/text-engine";
import { styleFromXml } from "@raabe/ooxml";
import { TABLE_STYLE_GROUPS, TABLE_STYLE_PRESETS } from "./table-styles.gen";

type Look = NonNullable<TableDef["look"]>;
const DEFAULT_LOOK: Look = { firstRow: true, lastRow: false, firstColumn: true, lastColumn: false, noHBand: false, noVBand: true };

const visible = (b: BorderSpec | undefined): b is BorderSpec => !!b && b.style !== "nil" && b.style !== "none" && b.size > 0;
const css = (b: BorderSpec | undefined) => (visible(b) ? `${b.style === "double" ? 3 : 1}px ${b.style === "double" ? "double" : /dot/.test(b.style) ? "dotted" : /dash/.test(b.style) ? "dashed" : "solid"} ${b.color === "auto" ? "#000" : `#${b.color}`}` : "none");

/** Word's gallery picture of a table style: 5 × 5 cells, a dash of "text" in each. */
export function tableStylePreview(def: StyleDef, look: Look, w: number, hgt: number): HTMLElement {
  const style: TableStyleDef = { ...(def.table ?? {}), rPr: def.rPr };
  const nRows = 5, nCols = 5;
  const table = h("table", { class: "tstyle-preview", style: `width:${w}px;height:${hgt}px` });
  for (let r = 0; r < nRows; r++) {
    const tr = h("tr");
    for (let c = 0; c < nCols; c++) {
      const { parts, edges: edge } = tableCellLook(style, look, r, nRows, c, 1, nCols);
      const fill = [...parts].reverse().map((p) => p.shading?.fill).find((f) => f && f !== "auto") ?? (style.shading?.fill && style.shading.fill !== "auto" ? style.shading.fill : undefined);
      const rpr = [...parts].reverse().map((p) => p.rPr).find((x) => x?.color) ?? style.rPr;
      const color = rpr?.color && rpr.color !== "auto" ? `#${rpr.color}` : "#000";
      const bold = parts.some((p) => p.rPr?.bold) || !!style.rPr?.bold;
      tr.append(h("td", {
        style: `background:${fill ? `#${fill}` : "#fff"};border-top:${css(edge.top)};border-bottom:${css(edge.bottom)};border-left:${css(edge.left)};border-right:${css(edge.right)}`,
      }, h("i", { style: `background:${color};height:${bold ? 2 : 1}px` })));
    }
    table.append(tr);
  }
  return table;
}

/** The gallery for the tables feature: the ribbon control and the full dropdown. */
export function tableStyleGallery(ed: () => Editor | undefined, done: () => void, more: { clear: () => void }) {
  const cache = new Map<string, StyleDef | undefined>();
  const colors = () => ed()?.doc.styles.theme.colors ?? {};
  const themeKey = () => Object.values(colors()).join("");
  const defOf = (i: number): StyleDef | undefined => {
    const k = `${themeKey()}|${i}`;
    if (!cache.has(k)) cache.set(k, styleFromXml(TABLE_STYLE_PRESETS[i].xml, colors()));
    return cache.get(k);
  };
  const look = () => ed()?.tableDef()?.look ?? DEFAULT_LOOK;
  /** The gallery entry of the table's style (by its name, as documents name Word's styles). */
  const current = () => {
    const e = ed();
    const id = e?.tableDef()?.style;
    if (!e || !id) return undefined;
    const name = e.doc.styles.get(id)?.name ?? id;
    const i = TABLE_STYLE_PRESETS.findIndex((p) => p.name === name || p.id === id);
    return i >= 0 ? i : undefined;
  };
  let previewing = false;
  const hover = (i: number | null) => {
    const e = ed();
    if (!e) return;
    if (i === null) { if (previewing) { e.previewTableStyle(null); previewing = false; } return; }
    const def = defOf(i);
    if (!def) return;
    e.previewTableStyle(def);
    previewing = true;
  };
  const pick = (i: number) => { previewing = false; const def = defOf(i); if (def) ed()?.applyTableStyle(def); done(); };
  const tile = (i: number, w: number, hh: number) => {
    const def = defOf(i);
    return def ? tableStylePreview(def, look(), w, hh) : h("span");
  };

  const panel = (close: () => void) => {
    const box = h("div", { class: "rpanel tstyle-gallery" });
    let at = 0;
    const cur = current();
    for (const [title, n] of TABLE_STYLE_GROUPS) {
      box.append(h("div", { class: "title" }, title));
      const grid = h("div", { class: "tiles", style: "grid-template-columns:repeat(7, auto)" });
      for (let i = at; i < at + n; i++) {
        const b = h("button", {
          class: `tile${cur === i ? " on" : ""}`, "data-tip": TABLE_STYLE_PRESETS[i].label, "aria-label": TABLE_STYLE_PRESETS[i].label,
          onclick: () => { close(); pick(i); },
        }, tile(i, 60, 44));
        b.addEventListener("mouseenter", () => hover(i));
        b.addEventListener("mouseleave", () => hover(null));
        grid.append(b);
      }
      box.append(grid);
      at += n;
    }
    box.append(h("hr"),
      h("button", { class: "link", disabled: true }, "Tabellenformatvorlage ändern..."),
      h("button", { class: "link", onclick: () => { close(); hover(null); more.clear(); } }, "Löschen"),
      h("button", { class: "link", disabled: true }, "Neue Tabellenformatvorlage..."));
    // The popup closing any other way ends the preview.
    const watch = new MutationObserver(() => { if (!box.isConnected) { hover(null); watch.disconnect(); } });
    setTimeout(() => watch.observe(document.body, { childList: true, subtree: true }));
    return box;
  };

  const control: Control = {
    type: "tiles", cmd: "table.styles", key: "O", count: TABLE_STYLE_PRESETS.length, visible: 7, fill: true, size: [72, 56], cls: "tstyle-tiles",
    render: (i) => tile(i, 62, 46),
    tip: (i) => TABLE_STYLE_PRESETS[i].label,
    pick,
    hover,
    current,
    redraw: () => { const l = look(); return `${themeKey()}|${Object.values(l).map((v) => (v ? 1 : 0)).join("")}`; },
    panel,
  };
  return { control, panel };
}
