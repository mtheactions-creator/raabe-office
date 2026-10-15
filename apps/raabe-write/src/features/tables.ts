// Tables: Einfügen > Tabelle (the grid and the dialog) and the contextual tabs
// Tabellenentwurf and Layout, with Word's groups, labels and keytips.

import { PT, type BorderSpec, type Borders, type CellDef, type TableDef } from "@raabe/text-engine";
import { closePopup, h, icon, showDialog, showPopup, type Command, type Control, type MenuEntry, type Tab } from "@raabe/shell";
import { wrapIcon } from "./pictures";
import { CM, fmtCm, parseNum, type AppApi } from "./api";
import { tableStyleGallery } from "./table-style-gallery";

const cmd = (id: string, label: string, icon?: string, extra: Partial<Command> = {}): Command => ({ id, label, icon, ...extra });

export const TABLE_COMMANDS: Command[] = [
  cmd("table.insertDialog", "Tabelle einfügen...", "table"),
  cmd("table.headerRow", "Kopfzeile", undefined, { tip: "Kopfzeile\nDie erste Zeile der Tabelle besonders formatieren." }),
  cmd("table.totalRow", "Ergebniszeile", undefined, { tip: "Ergebniszeile\nDie letzte Zeile der Tabelle besonders formatieren." }),
  cmd("table.bandedRows", "Gebänderte Zeilen", undefined),
  cmd("table.firstColumn", "Erste Spalte", undefined),
  cmd("table.lastColumn", "Letzte Spalte", undefined),
  cmd("table.bandedColumns", "Gebänderte Spalten", undefined),
  cmd("table.styles", "Tabellenformatvorlagen", "table-style"),
  cmd("table.shading", "Schattierung", "shading"),
  cmd("table.borders", "Rahmen", "borders"),
  cmd("table.insertCells", "Zellen einfügen...", "table-insert-cells"),
  cmd("table.deleteCells", "Zellen löschen...", "table-delete"),
  cmd("table.autofit", "AutoAnpassen", "table-autofit"),
  cmd("table.autofitContents", "Automatisch an Inhalt anpassen", "table-autofit"),
  cmd("table.autofitWindow", "Automatisch an Fenster anpassen", "table-autofit"),
  cmd("table.fixedWidth", "Feste Spaltenbreite", "table-autofit"),
  // Turned text in cells needs vertical lines in the layout engine.
  cmd("table.textDirection", "Textrichtung", "text-direction", { deferred: true }),
  cmd("table.splitTable", "Tabelle teilen", "table-split-table"),
  cmd("table.cellMargins", "Zellenbegrenzungen...", "cell-margins"),
  cmd("table.toText", "In Text konvertieren...", "table-to-text"),
  cmd("table.eraser", "Radierer", "table-eraser", { deferred: true }),
  cmd("table.formula", "Formel...", "table-formula", { deferred: true }),
  cmd("table.select", "Auswählen", "table-select"),
  cmd("table.selectCell", "Zelle auswählen", "table-select"),
  cmd("table.selectColumn", "Spalte auswählen", "table-select"),
  cmd("table.selectRow", "Zeile auswählen", "table-select"),
  cmd("table.selectTable", "Tabelle auswählen", "table-select"),
  cmd("table.gridlines", "Gitternetzlinien anzeigen", "gridlines"),
  cmd("table.properties", "Eigenschaften", "table-properties"),
  cmd("table.delete", "Löschen", "table-delete"),
  cmd("table.deleteColumns", "Spalten löschen", "table-delete-column"),
  cmd("table.deleteRows", "Zeilen löschen", "table-delete-row"),
  cmd("table.deleteTable", "Tabelle löschen", "table-delete"),
  // Word's labels on the ribbon (recorded 2026-10-15).
  cmd("table.insertAbove", "Zeile Oberhalb", "table-above"),
  cmd("table.insertBelow", "Zeile unterhalb einfügen", "table-below"),
  cmd("table.insertLeft", "Spalte Links", "table-left"),
  cmd("table.insertRight", "Spalte rechts einfügen", "table-right"),
  cmd("table.merge", "Zellen verbinden", "table-merge"),
  cmd("table.split", "Zellen teilen", "table-split"),
  cmd("table.height", "Höhe der Tabellenzeile", undefined),
  cmd("table.width", "Breite der Tabellenspalte", undefined),
  cmd("table.distributeRows", "Zeilen verteilen", "table-rows-even"),
  cmd("table.distributeColumns", "Spalten verteilen", "table-cols-even"),
  cmd("table.repeatHeader", "Überschriften wiederholen", "table-repeat-header", { tip: "Überschriften wiederholen\nDie Kopfzeilen auf jeder Seite wiederholen, wenn die Tabelle über mehrere Seiten geht." }),
  ...([["TopLeft", "Oben links ausrichten"], ["TopCenter", "Oben zentriert ausrichten"], ["TopRight", "Oben rechts ausrichten"],
    ["MiddleLeft", "Mitte links ausrichten"], ["MiddleCenter", "Zentriert ausrichten"], ["MiddleRight", "Mitte rechts ausrichten"],
    ["BottomLeft", "Unten links ausrichten"], ["BottomCenter", "Unten Mitte ausrichten"], ["BottomRight", "Unten rechts ausrichten"]] as const)
    .map(([k, label]) => cmd(`table.align${k}`, label, `cell-${k.replace(/([A-Z])/g, "-$1").slice(1).toLowerCase()}`)),
  cmd("border.bottom", "Rahmenlinie unten", "border-bottom"),
  cmd("border.top", "Rahmenlinie oben", "border-top"),
  cmd("border.left", "Rahmenlinie links", "border-left"),
  cmd("border.right", "Rahmenlinie rechts", "border-right"),
  cmd("border.none", "Kein Rahmen", "border-none"),
  cmd("border.all", "Alle Rahmenlinien", "border-all"),
  cmd("border.outside", "Rahmenlinien außen", "border-outside"),
  cmd("border.inside", "Rahmenlinien innen", "border-inside"),
  cmd("border.insideH", "Horizontale Rahmenlinie innen", "border-inside-h"),
  cmd("border.insideV", "Vertikale Rahmenlinie innen", "border-inside-v"),
  cmd("border.diagDown", "Rahmenlinien diagonal nach unten", "border-diag-down", { deferred: true }),
  cmd("border.diagUp", "Rahmenlinien diagonal nach oben", "border-diag-up", { deferred: true }),
  cmd("border.horizontalLine", "Horizontale Linie", "line-style", { deferred: true }),
  cmd("table.draw", "Tabelle zeichnen", "table-draw", { deferred: true }),
  cmd("table.borderStyles", "Rahmenarten", "border-styles", { tip: "Rahmenarten\nEine Rahmenart auswählen: Linie, Stärke und Farbe zusammen. Danach mit Rahmen übertragen auf die Rahmenlinien klicken." }),
  cmd("table.penStyle", "Stiftart", "line-style", { tip: "Stiftart\nDie Art der Linie für Rahmen ändern." }),
  cmd("table.penWeight", "Stiftstärke", "line-weight-pen", { tip: "Stiftstärke\nDie Breite der Linie für Rahmen ändern." }),
  cmd("table.penColor", "Stiftfarbe", "pen", { tip: "Stiftfarbe\nDie Farbe des Stifts ändern." }),
  cmd("table.borderPainter", "Rahmen übertragen", "border-painter", { tip: "Rahmen übertragen\nDie ausgewählte Rahmenart auf bestimmte Rahmenlinien in der Tabelle anwenden.\n\nKlicken Sie auf die Rahmenlinien, auf die Sie die Formatvorlage anwenden möchten." }),
  cmd("table.borderSampler", "Rahmenpipette", "pipette"),
];

/** Stiftart: Word's 25 line styles in its order (recorded 2026-10-15). */
export const PEN_STYLES: [string, string][] = [
  ["nil", "Kein Rand"], ["single", "Einfache einfarbige Linie"], ["dotted", "Punktiert"], ["dashSmallGap", "Gestrichelt (kleiner Abstand)"],
  ["dashed", "Gestrichelt (große Zwischenräume)"], ["dotDash", "Strichpunkt"], ["dotDotDash", "Strichpunkt-Punkt"], ["double", "Doppelte einfarbige Linien"],
  ["triple", "Dreifache einfarbige Linien"], ["thickThinSmallGap", "Kleiner Abstand dick-dünn"], ["thinThickSmallGap", "Kleiner Abstand dünn-dick"],
  ["thinThickThinSmallGap", "Kleiner Abstand dünn-dick-dünn"], ["thickThinMediumGap", "Mittlerer Abstand dick-dünn"], ["thinThickMediumGap", "Mittlerer Abstand dünn-dick"],
  ["thinThickThinMediumGap", "Mittlerer Abstand dünn-dick-dünn"], ["thickThinLargeGap", "Großer Abstand dick-dünn"], ["thinThickLargeGap", "Großer Abstand dünn-dick"],
  ["thinThickThinLargeGap", "Großer Abstand dünn-dick-dünn"], ["wave", "Einzelne Welle"], ["doubleWave", "Doppelt gewellt"], ["dashDotStroked", "Strichpunkt (gestrichelt)"],
  ["threeDEmboss", "Relief 3D"], ["threeDEngrave", "Gravur 3D"], ["outset", "Anfang"], ["inset", "Abgesenkt"],
];
/** Stiftstärke: Word's widths. */
export const PEN_WEIGHTS: [number, string][] = [[0.25, "¼ Pt."], [0.5, "½ Pt."], [0.75, "¾ Pt."], [1, "1 Pt."], [1.5, "1 ½ Pt."], [2.25, "2 ¼ Pt."], [3, "3 Pt."], [4.5, "4 ½ Pt."], [6, "6 Pt."]];

/** CSS for a line sample of a border style. */
const SAMPLE_CSS: Record<string, string> = {
  single: "solid", dotted: "dotted", dashSmallGap: "dashed", dashed: "dashed", dotDash: "dashed", dotDotDash: "dotted", double: "double", triple: "double",
  wave: "solid", doubleWave: "double", dashDotStroked: "dashed", threeDEmboss: "ridge", threeDEngrave: "groove", outset: "outset", inset: "inset",
};
const sampleCss = (style: string) => SAMPLE_CSS[style] ?? (/thin|thick/.test(style) ? "double" : "solid");

/** A drawn line sample (Stiftart, Stiftstärke, Rahmenarten). */
export function penSample(style: string, size: number, color: string, width = 110): HTMLElement {
  if (style === "nil" || style === "none") return h("span", { class: "pen-sample", style: `width:${width}px;justify-content:center;color:#616161` }, "Kein Rand");
  const multi = sampleCss(style) === "double" || ["ridge", "groove", "outset", "inset"].includes(sampleCss(style));
  const px = Math.max(multi ? 3 : 1, Math.round(size * (multi ? 2.5 : 1.33)));
  return h("span", { class: "pen-sample", style: `width:${width}px` },
    h("span", { style: `flex:1;border-top:${px}px ${sampleCss(style)} ${color === "auto" ? "#000" : `#${color}`}` }));
}

const NIL: BorderSpec = { style: "nil", size: 0, color: "auto", space: 0 };

export function tablesFeature(api: AppApi) {
  const ed = () => api.editor();
  const inTable = () => !!ed()?.inTable();
  const c = api.commands;
  const run = (id: string, fn: () => void, enabled = inTable, checked?: () => boolean) =>
    c.implement(id, { run: () => { fn(); api.focus(); }, enabled, checked });

  // ----- inserting ------------------------------------------------------------------------
  const insert = (rows: number, cols: number) => { ed()?.insertTable(rows, cols); api.focus(); };
  const insertDialog = async () => {
    const cols = h("input", { type: "number", min: "1", max: "63", value: "5", style: "width:70px" });
    const rows = h("input", { type: "number", min: "1", max: "500", value: "2", style: "width:70px" });
    const ok = await showDialog({
      title: "Tabelle einfügen",
      body: h("div", { style: "width:280px" }, h("fieldset", null, h("legend", null, "Tabellengröße"),
        h("div", { class: "grid2" }, "Spaltenanzahl:", cols, "Zeilenanzahl:", rows))),
      buttons: [{ label: "OK", value: true, primary: true }, { label: "Abbrechen", value: false }], cancel: false,
    });
    if (ok) insert(Math.max(1, Math.min(500, Number(rows.value) || 2)), Math.max(1, Math.min(63, Number(cols.value) || 5)));
  };
  c.implement("insert.table", { run: () => {}, enabled: () => !!ed() && !inTable() });
  c.implement("table.insertDialog", { run: insertDialog, enabled: () => !!ed() && !inTable() });

  /** Word's grid: hover to choose rows × columns ("3x4 Tabelle"). */
  const insertPanel = (close: () => void) => {
    const COLS = 10, ROWS = 8;
    const title = h("div", { class: "title" }, "Tabelle einfügen");
    const grid = h("div", { class: "grid", style: `grid-template-columns:repeat(${COLS},16px)` });
    const cells: HTMLElement[] = [];
    for (let r = 0; r < ROWS; r++) {
      for (let k = 0; k < COLS; k++) {
        const cell = h("div", { class: "cell" });
        cell.onmouseenter = () => {
          cells.forEach((x, i) => x.classList.toggle("on", Math.floor(i / COLS) <= r && i % COLS <= k));
          title.textContent = `Tabelle ${k + 1}x${r + 1}`;
        };
        cell.onclick = () => { close(); insert(r + 1, k + 1); };
        cells.push(cell);
        grid.append(cell);
      }
    }
    grid.onmouseleave = () => { cells.forEach((x) => x.classList.remove("on")); title.textContent = "Tabelle einfügen"; };
    return h("div", { class: "rpanel" }, title, grid, h("hr"),
      h("button", { class: "link", onclick: () => { close(); void insertDialog(); } }, "Tabelle einfügen..."));
  };

  // ----- Tabellenentwurf --------------------------------------------------------------------
  const look = (key: keyof NonNullable<TableDef["look"]>, invert = false) => {
    const get = () => { const l = ed()?.tableDef()?.look; return l ? (invert ? !l[key] : l[key]) : false; };
    return {
      run: () => ed()?.updateTable((def) => {
        const l = def.look ?? { firstRow: true, lastRow: false, firstColumn: true, lastColumn: false, noHBand: false, noVBand: true };
        return { ...def, look: { ...l, [key]: !l[key] } };
      }),
      checked: get,
    };
  };
  for (const [id, key, inv] of [["table.headerRow", "firstRow", false], ["table.totalRow", "lastRow", false], ["table.bandedRows", "noHBand", true],
    ["table.firstColumn", "firstColumn", false], ["table.lastColumn", "lastColumn", false], ["table.bandedColumns", "noVBand", true]] as const) {
    const l = look(key, inv);
    run(id, l.run, inTable, l.checked);
  }
  // Tabellenformatvorlagen: Word's gallery (table-style-gallery.ts).
  c.implement("table.styles", { run: () => {}, enabled: inTable });
  const styles = tableStyleGallery(ed, () => api.focus(), { clear: () => { ed()?.applyTableStyle(null); api.focus(); } });

  let lastShading = "FFFF00";
  const shade = (fill: string) => ed()?.updateCells((cd) => ({ ...cd, shading: fill === "auto" ? undefined : { fill, val: "clear", color: "auto" } }));
  run("table.shading", () => shade(lastShading));

  // Rahmen for cells: the edges of the selected block of cells.
  const cellBorders = (kind: string) => {
    const e = ed();
    if (!e) return;
    const sel = e.cellBlock();
    if (!sel) return;
    const line = pen.style === "nil" ? NIL : penSpec();
    e.updateCells((cd, r, col) => {
      const b: Borders = { ...(cd.borders ?? {}) };
      const top = r === sel.r1, bottom = r === sel.r2, left = col === sel.c1, right = col === sel.c2;
      const set = (side: keyof Borders, on: boolean) => { (b as Record<string, BorderSpec>)[side] = on ? line : NIL; };
      switch (kind) {
        case "all": set("top", true); set("bottom", true); set("left", true); set("right", true); break;
        case "none": set("top", false); set("bottom", false); set("left", false); set("right", false); break;
        case "outside": if (top) set("top", true); if (bottom) set("bottom", true); if (left) set("left", true); if (right) set("right", true); break;
        case "inside": if (!top) set("top", true); if (!bottom) set("bottom", true); if (!left) set("left", true); if (!right) set("right", true); break;
        case "insideH": if (!top) set("top", true); if (!bottom) set("bottom", true); break;
        case "insideV": if (!left) set("left", true); if (!right) set("right", true); break;
        case "top": if (top) set("top", true); break;
        case "bottom": if (bottom) set("bottom", true); break;
        case "left": if (left) set("left", true); break;
        case "right": if (right) set("right", true); break;
      }
      return { ...cd, borders: b };
    });
  };

  // Rahmen: Word's pen (Stiftart, Stiftstärke, Stiftfarbe; Rahmenarten sets all three). Choosing
  // any of them starts Rahmen übertragen; the Rahmen menu draws with the pen too.
  const pen: BorderSpec = { style: "single", size: 0.5, color: "auto", space: 0 };
  const penSpec = (): BorderSpec => ({ ...pen });
  const startPainter = () => { const e = ed(); if (e && inTable()) e.startBorderPainter(penSpec()); api.refresh(); };
  const setPen = (patch: Partial<BorderSpec>) => { Object.assign(pen, patch); startPainter(); };
  run("table.borderPainter", () => { const e = ed(); if (!e) return; if (e.borderPainting) e.stopBorderPainter(); else e.startBorderPainter(penSpec()); }, inTable, () => !!ed()?.borderPainting);
  run("table.borderSampler", () => ed()?.startBorderSampler((b) => { Object.assign(pen, { style: b.style === "none" ? "nil" : b.style, size: b.size || 0.5, color: b.color }); startPainter(); }),
    inTable, () => !!ed()?.borderSampling);
  c.implement("table.borderStyles", { run: startPainter, enabled: inTable, checked: () => !!ed()?.borderPainting });
  c.implement("table.penStyle", { run: () => {}, enabled: inTable });
  c.implement("table.penWeight", { run: () => {}, enabled: inTable });
  c.implement("table.penColor", { run: startPainter, enabled: inTable });
  const themeColor = (k: string) => (ed()?.doc.styles.theme.colors ?? {})[k] ?? "000000";
  /** Rahmenarten > Designrahmen: three lines (½ Pt., 1 ½ Pt., double ½ Pt.) in Automatisch and Akzent 1–6. */
  const designBorders = (): { label: string; spec: BorderSpec }[] => {
    const out: { label: string; spec: BorderSpec }[] = [];
    for (const [style, size, name] of [["single", 0.5, "Einfache einfarbige Linie, 1/2 Pt."], ["single", 1.5, "Einfache einfarbige Linie, 1 1/2 Pt."], ["double", 0.5, "Doppelte einfarbige Linien, 1/2 Pt."]] as const) {
      out.push({ label: name, spec: { style, size, color: "auto", space: 0 } });
      for (let a = 1; a <= 6; a++) out.push({ label: `${name}, Akzent ${a}`, spec: { style, size, color: themeColor(`accent${a}`), space: 0 } });
    }
    return out;
  };
  const borderStylesPanel = (close: () => void) => h("div", { class: "rpanel" }, h("div", { class: "title" }, "Designrahmen"),
    h("div", { class: "tiles", style: "grid-template-columns:repeat(7, auto)" }, ...designBorders().map((d) => h("button", {
      class: `tile${d.spec.style === pen.style && d.spec.size === pen.size && d.spec.color === pen.color ? " on" : ""}`, "data-tip": d.label, "aria-label": d.label,
      style: "width:64px;height:28px", onclick: () => { close(); setPen(d.spec); },
    }, penSample(d.spec.style, d.spec.size, d.spec.color, 52)))),
    h("hr"),
    h("button", { class: "link", onclick: () => { close(); void c.run("table.borderSampler"); } }, "Rahmenpipette"));

  // ----- Layout ----------------------------------------------------------------------------
  run("table.selectCell", () => ed()?.selectTablePart("cell"));
  run("table.selectColumn", () => ed()?.selectTablePart("column"));
  run("table.selectRow", () => ed()?.selectTablePart("row"));
  run("table.selectTable", () => ed()?.selectTablePart("table"));
  c.implement("table.select", { run: () => {}, enabled: inTable });
  c.implement("table.delete", { run: () => {}, enabled: inTable });
  run("table.deleteColumns", () => ed()?.deleteColumns());
  // Zellen löschen: Word's dialog (recorded 2026-10-15), "Zellen nach links verschieben" first.
  c.implement("table.deleteCells", { run: async () => {
    const e = ed();
    if (!e?.inTable()) return;
    const choices: ["left" | "up" | "row" | "column", string][] = [["left", "Zellen nach links verschieben"], ["up", "Zellen nach oben verschieben"],
      ["row", "Ganze Zeile löschen"], ["column", "Ganze Spalte löschen"]];
    const radios = choices.map(([v, l], i) => ({ v, el: h("input", { type: "radio", name: "delcells", checked: i === 0 }) as HTMLInputElement, l }));
    const ok = await showDialog({
      title: "Zellen löschen",
      body: h("div", { style: "display:flex;flex-direction:column;gap:6px;min-width:240px" }, ...radios.map((r) => h("label", { class: "check" }, r.el, r.l))),
      buttons: [{ label: "OK", value: true, primary: true }, { label: "Abbrechen", value: false }], cancel: false,
    });
    if (ok) e.deleteCells(radios.find((r) => r.el.checked)?.v ?? "left");
    api.focus();
  }, enabled: () => !!ed()?.inTable() });
  run("table.deleteRows", () => ed()?.deleteRows());
  run("table.deleteTable", () => ed()?.deleteTable());
  run("table.insertAbove", () => ed()?.insertRows("above"));
  run("table.insertBelow", () => ed()?.insertRows("below"));
  run("table.insertLeft", () => ed()?.insertColumns("left"));
  run("table.insertRight", () => ed()?.insertColumns("right"));
  run("table.merge", () => ed()?.mergeCells(), () => !!ed()?.cellBlock()?.multi);
  run("table.split", () => void splitDialog());
  run("table.splitTable", () => ed()?.splitTable());
  // Zellen einfügen: Word's dialog (recorded 2026-10-15), "Zellen nach unten verschieben" chosen.
  c.implement("table.insertCells", { run: async () => {
    const e = ed();
    if (!e?.inTable()) return;
    const choices: ["right" | "down" | "row" | "column", string][] = [["right", "Zellen nach rechts verschieben"], ["down", "Zellen nach unten verschieben"],
      ["row", "Ganze Zeile einfügen"], ["column", "Ganze Spalte einfügen"]];
    const radios = choices.map(([v, l], i) => ({ v, el: h("input", { type: "radio", name: "inscells", checked: i === 1 }) as HTMLInputElement, l }));
    const ok = await showDialog({
      title: "Zellen einfügen",
      body: h("div", { style: "display:flex;flex-direction:column;gap:6px;min-width:240px" }, ...radios.map((r) => h("label", { class: "check" }, r.el, r.l))),
      buttons: [{ label: "OK", value: true, primary: true }, { label: "Abbrechen", value: false }], cancel: false,
    });
    if (ok) e.insertCells(radios.find((r) => r.el.checked)?.v ?? "down");
    api.focus();
  }, enabled: inTable });
  // AutoAnpassen.
  c.implement("table.autofit", { run: () => {}, enabled: inTable });
  run("table.autofitContents", () => ed()?.autofitTable("contents"));
  run("table.autofitWindow", () => ed()?.autofitTable("window"));
  run("table.fixedWidth", () => ed()?.autofitTable("fixed"));
  // Zellenbegrenzungen: Word's dialog Tabellenoptionen (recorded); Abstand zwischen Zellen is not drawn.
  run("table.cellMargins", async () => {
    const e = ed();
    const o = e?.tableOptions();
    if (!e || !o) return;
    const field = (v: number) => h("input", { type: "text", value: fmtCm(v), style: "width:90px" }) as HTMLInputElement;
    const top = field(o.top), bottom = field(o.bottom), left = field(o.left), right = field(o.right);
    const spacing = h("input", { type: "checkbox", disabled: true }) as HTMLInputElement;
    const auto = h("input", { type: "checkbox", checked: o.autoResize }) as HTMLInputElement;
    const ok = await showDialog({
      title: "Tabellenoptionen",
      body: h("div", { style: "width:380px" },
        h("fieldset", null, h("legend", null, "Standardzellenbegrenzungen"),
          h("div", { style: "display:grid;grid-template-columns:auto auto auto auto;gap:6px 10px;align-items:center" }, "Oben:", top, "Links:", left, "Unten:", bottom, "Rechts:", right)),
        h("fieldset", null, h("legend", null, "Standardzellenabstand"),
          h("label", { class: "check", "data-tip": "Abstand zwischen Zellen zulassen\nRaabeWrite zeichnet Tabellen ohne Abstand zwischen den Zellen." }, spacing, "Abstand zwischen Zellen zulassen")),
        h("fieldset", null, h("legend", null, "Optionen"), h("label", { class: "check" }, auto, "Automatische Größenänderung zulassen"))),
      buttons: [{ label: "OK", value: true, primary: true }, { label: "Abbrechen", value: false }], cancel: false,
    });
    if (!ok) return;
    const val = (f: HTMLInputElement, old: number) => { const v = parseNum(f.value); return Number.isFinite(v) && v >= 0 ? v * CM : old; };
    e.setTableOptions({ top: val(top, o.top), bottom: val(bottom, o.bottom), left: val(left, o.left), right: val(right, o.right), autoResize: auto.checked });
    api.focus();
  });
  // In Text konvertieren: Word's dialog Tabelle in Text umwandeln (recorded), Tabstopps chosen.
  run("table.toText", async () => {
    const e = ed();
    if (!e?.inTable()) return;
    const choices: [string, string][] = [["\n", "Absatzmarken"], ["\t", "Tabstopps"], ["; ", "Semikolons"], ["", "Andere:"]];
    const radios = choices.map(([v, l], i) => ({ v, el: h("input", { type: "radio", name: "totext", checked: i === 1 }) as HTMLInputElement, l }));
    const other = h("input", { type: "text", value: "-", maxlength: "1", style: "width:36px" }) as HTMLInputElement;
    other.addEventListener("focus", () => { radios[3].el.checked = true; });
    const nested = h("input", { type: "checkbox", checked: true, disabled: true }) as HTMLInputElement;
    const ok = await showDialog({
      title: "Tabelle in Text umwandeln",
      body: h("div", { style: "width:300px" },
        h("fieldset", null, h("legend", null, "Text trennen durch"),
          h("div", { style: "display:flex;flex-direction:column;gap:6px" },
            ...radios.map((r, i) => h("label", { class: "check" }, r.el, r.l, ...(i === 3 ? [other] : []))))),
        h("label", { class: "check" }, nested, "Geschachtelte Tabellen umwandeln")),
      buttons: [{ label: "OK", value: true, primary: true }, { label: "Abbrechen", value: false }], cancel: false,
    });
    if (!ok) return;
    const pick = radios.find((r) => r.el.checked);
    e.tableToText(pick?.l === "Andere:" ? other.value || "-" : pick?.v ?? "\t");
    api.focus();
  });
  run("table.distributeColumns", () => ed()?.distributeColumns());
  run("table.distributeRows", () => ed()?.updateRows((r) => ({ ...r, height: undefined, heightRule: undefined })));
  run("table.repeatHeader", () => { const on = !!ed()?.rowDef()?.header; ed()?.updateRows((r) => ({ ...r, header: !on || undefined })); }, inTable, () => !!ed()?.rowDef()?.header);
  c.implement("table.gridlines", { run: () => { const e = ed(); if (e) e.setShowGridlines(!e.showGridlines); }, enabled: inTable, checked: () => !!ed()?.showGridlines });
  run("table.properties", () => void propertiesDialog());
  const aligns: [string, "top" | "center" | "bottom", "left" | "center" | "right"][] = [
    ["TopLeft", "top", "left"], ["TopCenter", "top", "center"], ["TopRight", "top", "right"],
    ["MiddleLeft", "center", "left"], ["MiddleCenter", "center", "center"], ["MiddleRight", "center", "right"],
    ["BottomLeft", "bottom", "left"], ["BottomCenter", "bottom", "center"], ["BottomRight", "bottom", "right"],
  ];
  for (const [k, v, hz] of aligns) {
    run(`table.align${k}`, () => {
      const e = ed();
      if (!e) return;
      e.updateCells((cd) => ({ ...cd, vAlign: v === "top" ? undefined : v }));
      e.setParaProps({ align: hz });
    }, inTable, () => { const e = ed(); return !!e && (e.cellDef()?.vAlign ?? "top") === v && e.paraState().align === hz; });
  }
  for (const k of ["bottom", "top", "left", "right", "none", "all", "outside", "inside"]) {
    run(`border.${k}`, () => (inTable() ? cellBorders(k) : ed()?.paragraphBorders(k)), () => !!ed());
  }
  run("border.insideH", () => cellBorders("insideH"));
  run("border.insideV", () => cellBorders("insideV"));
  c.implement("table.borders", { run: () => (inTable() ? cellBorders("all") : ed()?.paragraphBorders("bottom")), enabled: () => !!ed() });

  const splitDialog = async () => {
    const cols = h("input", { type: "number", min: "2", max: "20", value: "2", style: "width:70px" });
    const ok = await showDialog({
      title: "Zellen teilen", body: h("div", { class: "grid2", style: "width:240px" }, "Spaltenanzahl:", cols),
      buttons: [{ label: "OK", value: true, primary: true }, { label: "Abbrechen", value: false }], cancel: false,
    });
    if (ok) ed()?.splitCell(Math.max(2, Math.min(20, Number(cols.value) || 2)));
  };

  const propertiesDialog = async () => {
    const e = ed();
    const def = e?.tableDef();
    if (!e || !def) return;
    const align = h("select", null, ...[["left", "Links"], ["center", "Zentriert"], ["right", "Rechts"]].map(([v, l]) => h("option", { value: v }, l)));
    align.value = def.float ? def.float.xAlign ?? "left" : def.align ?? "left";
    const wrapNone = h("input", { type: "radio", name: "tp-wrap", checked: !def.float }) as HTMLInputElement;
    const wrapAround = h("input", { type: "radio", name: "tp-wrap", checked: !!def.float }) as HTMLInputElement;
    const indent = h("input", { type: "text", value: fmtCm(def.indent ?? 0), style: "width:90px" });
    const row = e.rowDef();
    const height = h("input", { type: "text", value: row?.height !== undefined ? fmtCm(row.height) : "", style: "width:90px", placeholder: "automatisch" });
    const rule = h("select", null, h("option", { value: "atLeast" }, "Mindestens"), h("option", { value: "exact" }, "Genau"));
    rule.value = row?.heightRule ?? "atLeast";
    const cantSplit = h("input", { type: "checkbox", checked: !row?.cantSplit });
    const header = h("input", { type: "checkbox", checked: !!row?.header });
    const ok = await showDialog({
      title: "Tabelleneigenschaften",
      body: h("div", { style: "width:380px" },
        h("fieldset", null, h("legend", null, "Tabelle"), h("div", { class: "grid2" }, "Ausrichtung:", align, "Einzug von links:", indent),
          h("div", { style: "display:flex;gap:16px;margin-top:6px" }, h("span", null, "Textumbruch:"),
            h("label", { class: "check" }, wrapNone, "Ohne"), h("label", { class: "check" }, wrapAround, "Umgebend"))),
        h("fieldset", null, h("legend", null, "Zeile"), h("div", { class: "grid2" }, "Höhe angeben:", height, "Zeilenhöhe:", rule),
          h("label", { class: "check" }, cantSplit, "Zeilenwechsel auf Seiten zulassen"),
          h("label", { class: "check" }, header, "Gleiche Kopfzeile auf jeder Seite wiederholen"))),
      buttons: [{ label: "OK", value: true, primary: true }, { label: "Abbrechen", value: false }], cancel: false,
    });
    if (!ok) return;
    e.setTableWrap(wrapAround.checked ? "around" : "none");
    if (wrapAround.checked) e.setTableAlign(align.value as "left" | "center" | "right");
    else e.updateTable((d) => ({ ...d, align: align.value === "left" ? undefined : align.value as "center" | "right", indent: parseNum(indent.value) * CM || undefined }));
    const hv = parseNum(height.value);
    e.updateRows((r) => ({
      ...r, height: Number.isFinite(hv) && hv > 0 ? hv * CM : undefined, heightRule: Number.isFinite(hv) && hv > 0 ? rule.value as "atLeast" | "exact" : undefined,
      cantSplit: !cantSplit.checked || undefined, header: header.checked || undefined,
    }));
  };

  // Zellengröße: row height and column width of the caret's cell.
  const heightValue = () => { const r = ed()?.rowDef(); return r?.height !== undefined ? fmtCm(r.height) : ""; };
  const widthValue = () => { const w = ed()?.cellWidth(); return w ? fmtCm(w) : ""; };
  c.implement("table.height", { run: () => {}, enabled: inTable });
  c.implement("table.width", { run: () => {}, enabled: inTable });
  const setHeight = (v: string) => { const n = parseNum(v); if (Number.isFinite(n) && n > 0) ed()?.updateRows((r) => ({ ...r, height: n * CM, heightRule: r.heightRule ?? "atLeast" })); };
  const setWidth = (v: string) => { const n = parseNum(v); if (Number.isFinite(n) && n > 0) ed()?.setCellWidth(n * CM); };
  const stepHeight = (d: 1 | -1) => { const r = ed()?.rowDef(); const cur = (r?.height ?? 12.95) / CM; setHeight(String(Math.max(0.1, Math.round((cur + d * 0.1) * 100) / 100))); };
  const stepWidth = (d: 1 | -1) => { const w = (ed()?.cellWidth() ?? 0) / CM; setWidth(String(Math.max(0.2, Math.round((w + d * 0.1) * 100) / 100))); };

  const b = (id: string, key: string, size: "large" | "medium" | "small" = "small", label?: string): Control => ({ type: "button", cmd: id, key, size, label });
  const check = (id: string, key: string): Control => ({ type: "check", cmd: id, key });
  const dd = (id: string, key: string, size: "large" | "medium" | "small", menu: MenuEntry[], popup?: (close: () => void) => HTMLElement): Control =>
    ({ type: "dropdown", cmd: id, key, size, menu, popup });
  const stack = (...items: Control[]): Control => ({ type: "stack", items });
  const bordersMenu: MenuEntry[] = [
    ...["bottom", "top", "left", "right"].map((k): MenuEntry => ({ cmd: `border.${k}` })), "-",
    ...["none", "all", "outside", "inside"].map((k): MenuEntry => ({ cmd: `border.${k}` })), "-",
    { cmd: "border.insideH" }, { cmd: "border.insideV" }, { cmd: "border.diagDown" }, { cmd: "border.diagUp" }, "-",
    { cmd: "border.horizontalLine" }, "-",
    { cmd: "table.draw" }, { cmd: "table.gridlines" }, { cmd: "border.dialog", label: "Rahmen und Schattierung..." },
  ];
  const penCombo = (cmdId: string, key: string, options: () => string[], value: () => string, pick: (v: string) => void, sample: (v: string) => HTMLElement): Control => ({
    type: "combo", cmd: cmdId, key, width: 140, value, options: () => [{ items: options() }], commit: pick, renderOption: sample, renderValue: () => sample(value()),
  });
  const styleCombo = penCombo("table.penStyle", "B", () => PEN_STYLES.map(([, l]) => l), () => PEN_STYLES.find(([v]) => v === pen.style)?.[1] ?? PEN_STYLES[1][1],
    (l) => { const v = PEN_STYLES.find(([, x]) => x === l)?.[0]; if (v) setPen({ style: v }); }, (l) => penSample(PEN_STYLES.find(([, x]) => x === l)?.[0] ?? "single", 0.75, "auto", 118));
  const weightCombo = penCombo("table.penWeight", "6", () => PEN_WEIGHTS.map(([, l]) => l), () => PEN_WEIGHTS.find(([v]) => v === pen.size)?.[1] ?? `${pen.size} Pt.`,
    (l) => { const v = PEN_WEIGHTS.find(([, x]) => x === l)?.[0]; if (v) setPen({ size: v }); },
    (l) => { const v = PEN_WEIGHTS.find(([, x]) => x === l)?.[0] ?? 0.5; const el = penSample(pen.style === "nil" ? "single" : pen.style, v, "auto", 70); el.prepend(h("span", { style: "width:42px;flex:none" }, l)); return el; });

  const shading: Control = { type: "color", cmd: "table.shading", key: "NC", palette: "font", showLabel: true, current: () => lastShading, pick: (hex) => { if (hex !== "auto") lastShading = hex; shade(hex); }, preview: (hex) => shade(hex) };
  const designTab: Tab = {
    id: "tableDesign", label: "Tabellenentwurf", key: "JT", contextual: true, shrink: ["borders", "styles", "options"],
    groups: [
      { id: "options", label: "Tabellenformatoptionen", key: "ZL", icon: "table-style",
        controls: [stack(check("table.headerRow", "Ü"), check("table.totalRow", "G"), check("table.bandedRows", "R1")),
          stack(check("table.firstColumn", "T"), check("table.lastColumn", "L"), check("table.bandedColumns", "U"))] },
      { id: "styles", label: "Tabellenformatvorlagen", key: "ZS", icon: "table-style",
        controls: [styles.control, shading],
        smaller: [[dd("table.styles", "O", "large", [], styles.panel), shading]] },
      { id: "borders", label: "Rahmen", key: "ZB", icon: "borders", launcher: { cmd: "border.dialog", key: "R2" },
        controls: [{ type: "split", cmd: "table.borderStyles", key: "H", size: "large", menu: [], popup: borderStylesPanel },
          stack(styleCombo, weightCombo, { type: "color", cmd: "table.penColor", key: "F", palette: "font", showLabel: true, current: () => pen.color, pick: (hex) => setPen({ color: hex }) }),
          { type: "split", cmd: "table.borders", key: "MR", size: "large", menu: bordersMenu },
          b("table.borderPainter", "A", "large")] },
    ],
  };

  // Word's Tabellenlayout (recorded 2026-10-15 at a wide window, probe_word_table_layout.py tab):
  // labels, keytips and layouts; the smaller ones are Word's at 2000 px.
  const select = (size: "large" | "medium") => dd("table.select", "W", size, [{ cmd: "table.selectCell" }, { cmd: "table.selectColumn" }, { cmd: "table.selectRow" }, { cmd: "table.selectTable" }]);
  const del = dd("table.delete", "Ö", "large", [{ cmd: "table.deleteCells" }, { cmd: "table.deleteColumns" }, { cmd: "table.deleteRows" }, { cmd: "table.deleteTable" }]);
  const autofit = (size: "large" | "medium") => dd("table.autofit", "AU", size, [{ cmd: "table.autofitContents" }, { cmd: "table.autofitWindow" }, { cmd: "table.fixedWidth" }]);
  const height: Control = { type: "spinner", cmd: "table.height", label: "Höhe:", key: "P", value: heightValue, commit: setHeight, step: stepHeight };
  const width: Control = { type: "spinner", cmd: "table.width", label: "Breite:", key: "B", value: widthValue, commit: setWidth, step: stepWidth };
  const layoutTab: Tab = {
    id: "tableLayout", label: "Layout", key: "JL", contextual: true,
    shrink: ["table", "draw", "merge", "rowsCols", "size", "data", "alignment", "size", "merge", "draw", "rowsCols", "table"],
    groups: [
      { id: "table", label: "Tabelle", key: "ZT", icon: "table",
        controls: [select("large"), b("table.gridlines", "ß", "large"), b("table.properties", "IE", "large")],
        smaller: [[stack(select("medium"), b("table.gridlines", "ß", "medium"), b("table.properties", "IE", "medium"))]] },
      { id: "draw", label: "Zeichnen", key: "ZD1", icon: "table-draw",
        controls: [b("table.draw", "AA", "large"), b("table.eraser", "R", "large")],
        smaller: [[stack(b("table.draw", "AA", "medium"), b("table.eraser", "R", "medium"))]] },
      { id: "rowsCols", label: "Zeilen und Spalten", key: "ZR", icon: "table-insert-row", launcher: { cmd: "table.insertCells", key: "EZ" },
        controls: [del, b("table.insertAbove", "Ü", "large"), b("table.insertBelow", "FA", "large"), b("table.insertLeft", "Q", "large"), b("table.insertRight", "K", "large")],
        smaller: [[del, b("table.insertAbove", "Ü", "large"), stack(b("table.insertBelow", "FA", "medium"), b("table.insertLeft", "Q", "medium"), b("table.insertRight", "K", "medium"))]] },
      { id: "merge", label: "Zusammenführen", key: "ZM", icon: "table-merge",
        controls: [b("table.merge", "LE", "large"), b("table.split", "NE", "large"), b("table.splitTable", "TT", "large")],
        smaller: [[stack(b("table.merge", "LE", "medium"), b("table.split", "NE", "medium"), b("table.splitTable", "TT", "medium"))]] },
      { id: "size", label: "Zellengröße", key: "ZS", icon: "table-cols-even", launcher: { cmd: "table.properties", key: "IE" },
        controls: [autofit("large"), stack(height, width), stack(b("table.distributeRows", "V", "medium"), b("table.distributeColumns", "SP", "medium"))],
        smaller: [[stack(height, width, autofit("medium")), stack(b("table.distributeRows", "V"), b("table.distributeColumns", "SP"))]] },
      { id: "alignment", label: "Ausrichtung", key: "ZA", icon: "cell-top-left",
        controls: [{ type: "rows", rows: [
          [b("table.alignTopLeft", "O"), b("table.alignTopCenter", "ÄB"), b("table.alignTopRight", "EB")],
          [b("table.alignMiddleLeft", "II"), b("table.alignMiddleCenter", "H"), b("table.alignMiddleRight", "M")],
          [b("table.alignBottomLeft", "TN"), b("table.alignBottomCenter", "CN"), b("table.alignBottomRight", "U")],
        ] }, b("table.textDirection", "CT", "large", "Text-\nrichtung"), b("table.cellMargins", "G", "large", "Zellenbe-\ngrenzungen")] },
      { id: "data", label: "Daten", key: "ZD2", icon: "table-repeat-header",
        controls: [b("para.sort", "J", "large"), b("table.repeatHeader", "D", "large"), b("table.toText", "X", "large"), b("table.formula", "FF", "large")] },
    ],
  };

  // Our own addition (D-053): the table's layout options beside its top right corner, like a
  // picture's: Textumbruch (Ohne / Umgebend) and Ausrichtung, with Tabelleneigenschaften below.
  const layoutOptions = (rect: DOMRect) => {
    const e = ed();
    const def = e?.tableDef();
    if (!e || !def) return;
    const floating = !!def.float;
    const align = def.float ? def.float.xAlign ?? "left" : def.align ?? "left";
    const tile = (label: string, on: boolean, body: HTMLElement, run: () => void) =>
      h("button", { class: `tile${on ? " on" : ""}`, "data-tip": label, "aria-label": label, style: "width:42px;height:42px", onclick: () => { closePopup(); run(); api.focus(); } }, body);
    const alignIcon = (a: string) => icon(a === "center" ? "align-center" : a === "right" ? "align-right" : "align-left");
    const panel = h("div", { class: "rpanel layout-options", style: "width:220px" },
      h("div", { class: "title" }, "Layoutoptionen"),
      h("div", { class: "subtitle" }, "Textumbruch"),
      h("div", { class: "tiles", style: "grid-template-columns:repeat(2, auto)" },
        tile("Ohne", !floating, wrapIcon("top-bottom"), () => e.setTableWrap("none")),
        tile("Umgebend", floating, wrapIcon("square"), () => e.setTableWrap("around"))),
      h("div", { class: "subtitle" }, "Ausrichtung"),
      h("div", { class: "tiles", style: "grid-template-columns:repeat(3, auto)" },
        ...([["left", "Links"], ["center", "Zentriert"], ["right", "Rechts"]] as const).map(([a, l]) => tile(l, align === a, alignIcon(a), () => e.setTableAlign(a)))),
      h("hr"),
      h("button", { class: "link", onclick: () => { closePopup(); void c.run("table.properties"); } }, "Tabelleneigenschaften..."));
    showPopup(panel, new DOMRect(rect.left, rect.bottom, 0, 0));
  };
  const attach = (e: import("@raabe/text-engine").Editor) => { e.onTableLayout = layoutOptions; };

  return { insertPanel, designTab, layoutTab, attach, cellWidthPt: () => (ed()?.cellWidth() ?? 0) / PT };
}

export type { CellDef };
