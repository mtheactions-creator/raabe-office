// Tables: Einfügen > Tabelle (the grid and the dialog) and the contextual tabs
// Tabellenentwurf and Layout, with Word's groups, labels and keytips.

import { PT, type BorderSpec, type Borders, type CellDef, type TableDef } from "@raabe/text-engine";
import { h, showDialog, type Command, type Control, type MenuEntry, type Tab } from "@raabe/shell";
import { CM, fmtCm, parseNum, type AppApi } from "./api";

const cmd = (id: string, label: string, icon?: string, extra: Partial<Command> = {}): Command => ({ id, label, icon, ...extra });

export const TABLE_COMMANDS: Command[] = [
  cmd("table.insertDialog", "Tabelle einfügen...", "table"),
  cmd("table.styleOptions", "Optionen für Tabellenformat", "table-style", { hidden: true }),
  cmd("table.headerRow", "Überschrift", undefined, { tip: "Überschrift\nDie erste Zeile der Tabelle besonders formatieren." }),
  cmd("table.totalRow", "Ergebniszeile", undefined, { tip: "Ergebniszeile\nDie letzte Zeile der Tabelle besonders formatieren." }),
  cmd("table.bandedRows", "Gebänderte Zeilen", undefined),
  cmd("table.firstColumn", "Erste Spalte", undefined),
  cmd("table.lastColumn", "Letzte Spalte", undefined),
  cmd("table.bandedColumns", "Gebänderte Spalten", undefined),
  cmd("table.styles", "Tabellenformatvorlagen", "table-style"),
  cmd("table.shading", "Schattierung", "shading"),
  cmd("table.borders", "Rahmen", "borders"),
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
  cmd("table.insertAbove", "Darüber einfügen", "table-above", { tip: "Darüber einfügen\nEine neue Zeile direkt über der aktuellen Zeile einfügen." }),
  cmd("table.insertBelow", "Darunter einfügen", "table-below", { tip: "Darunter einfügen\nEine neue Zeile direkt unter der aktuellen Zeile einfügen." }),
  cmd("table.insertLeft", "Links einfügen", "table-left", { tip: "Links einfügen\nEine neue Spalte direkt links von der aktuellen Spalte einfügen." }),
  cmd("table.insertRight", "Rechts einfügen", "table-right", { tip: "Rechts einfügen\nEine neue Spalte direkt rechts von der aktuellen Spalte einfügen." }),
  cmd("table.merge", "Zellen verbinden", "table-merge"),
  cmd("table.split", "Zellen teilen", "table-split"),
  cmd("table.height", "Höhe der Tabellenzeile", undefined),
  cmd("table.width", "Breite der Tabellenspalte", undefined),
  cmd("table.distributeRows", "Zeilen verteilen", "table-rows-even"),
  cmd("table.distributeColumns", "Spalten verteilen", "table-cols-even"),
  cmd("table.repeatHeader", "Überschriften wiederholen", "table-repeat-header", { tip: "Überschriften wiederholen\nDie Kopfzeilen auf jeder Seite wiederholen, wenn die Tabelle über mehrere Seiten geht." }),
  ...([["TopLeft", "Oben links ausrichten"], ["TopCenter", "Oben zentriert ausrichten"], ["TopRight", "Oben rechts ausrichten"],
    ["MiddleLeft", "Zentriert links ausrichten"], ["MiddleCenter", "Zentriert ausrichten"], ["MiddleRight", "Zentriert rechts ausrichten"],
    ["BottomLeft", "Unten links ausrichten"], ["BottomCenter", "Unten zentriert ausrichten"], ["BottomRight", "Unten rechts ausrichten"]] as const)
    .map(([k, label]) => cmd(`table.align${k}`, label, `cell-${k.replace(/([A-Z])/g, "-$1").slice(1).toLowerCase()}`)),
  cmd("border.bottom", "Rahmenlinie unten", "border-bottom"),
  cmd("border.top", "Rahmenlinie oben", "border-top"),
  cmd("border.left", "Rahmenlinie links", "border-left"),
  cmd("border.right", "Rahmenlinie rechts", "border-right"),
  cmd("border.none", "Kein Rahmen", "border-none"),
  cmd("border.all", "Alle Rahmenlinien", "border-all"),
  cmd("border.outside", "Rahmenlinien außen", "border-outside"),
  cmd("border.inside", "Rahmenlinien innen", "border-inside"),
];

const LINE: BorderSpec = { style: "single", size: 0.5, color: "auto", space: 0 };
const NIL: BorderSpec = { style: "nil", size: 0, color: "auto", space: 0 };

/** Table styles offered in the gallery (built-in Word styles; definitions in the engine). */
export const TABLE_STYLES: { name: string; label: string; preview: { head: string; band: string; line: string; headText: string } }[] = [
  { name: "Table Grid", label: "Tabellenraster", preview: { head: "#fff", band: "#fff", line: "#000", headText: "#000" } },
  { name: "Table Grid Light", label: "Tabellenraster hell", preview: { head: "#fff", band: "#fff", line: "#bfbfbf", headText: "#000" } },
  { name: "Plain Table 1", label: "Einfache Tabelle 1", preview: { head: "#fff", band: "#f2f2f2", line: "#bfbfbf", headText: "#000" } },
  { name: "Grid Table 1 Light", label: "Gitternetztabelle 1 hell", preview: { head: "#fff", band: "#fff", line: "#999", headText: "#000" } },
  { name: "Grid Table 4 Accent 1", label: "Gitternetztabelle 4 – Akzent 1", preview: { head: "#4472c4", band: "#d9e2f3", line: "#8eaadb", headText: "#fff" } },
  { name: "Grid Table 5 Dark Accent 1", label: "Gitternetztabelle 5 dunkel – Akzent 1", preview: { head: "#4472c4", band: "#b4c6e7", line: "#fff", headText: "#fff" } },
];

function stylePreview(p: (typeof TABLE_STYLES)[number]["preview"]): HTMLElement {
  const rows = [0, 1, 2, 3, 4].map((r) => h("div", { style: `display:flex;background:${r === 0 ? p.head : r % 2 ? p.band : "#fff"}` },
    ...[0, 1, 2, 3].map(() => h("div", { style: `width:14px;height:8px;border:1px solid ${p.line};margin:-1px 0 0 -1px` }))));
  return h("div", { style: "padding:4px" }, ...rows);
}

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
  const setStyle = (name: string) => {
    const e = ed();
    if (!e) return;
    const id = e.styleId(name);
    e.ensureStyle(id);
    e.updateTable((def) => ({ ...def, style: id }));
  };
  c.implement("table.styles", { run: () => {}, enabled: inTable });
  const stylesPanel = (close: () => void) => h("div", { class: "rpanel" }, h("div", { class: "title" }, "Tabellenformatvorlagen"),
    h("div", { class: "tiles" }, ...TABLE_STYLES.map((s) => h("button", {
      class: `tile${ed()?.tableDef()?.style === ed()?.styleId(s.name) ? " on" : ""}`, "data-tip": s.label,
      onclick: () => { close(); setStyle(s.name); api.focus(); },
    }, stylePreview(s.preview)))));

  let lastShading = "FFFF00";
  const shade = (fill: string) => ed()?.updateCells((cd) => ({ ...cd, shading: fill === "auto" ? undefined : { fill, val: "clear", color: "auto" } }));
  run("table.shading", () => shade(lastShading));

  // Rahmen for cells: the edges of the selected block of cells.
  const cellBorders = (kind: string) => {
    const e = ed();
    if (!e) return;
    const sel = e.cellBlock();
    if (!sel) return;
    e.updateCells((cd, r, col) => {
      const b: Borders = { ...(cd.borders ?? {}) };
      const top = r === sel.r1, bottom = r === sel.r2, left = col === sel.c1, right = col === sel.c2;
      const set = (side: keyof Borders, on: boolean) => { (b as Record<string, BorderSpec>)[side] = on ? LINE : NIL; };
      switch (kind) {
        case "all": set("top", true); set("bottom", true); set("left", true); set("right", true); break;
        case "none": set("top", false); set("bottom", false); set("left", false); set("right", false); break;
        case "outside": if (top) set("top", true); if (bottom) set("bottom", true); if (left) set("left", true); if (right) set("right", true); break;
        case "inside": if (!top) set("top", true); if (!bottom) set("bottom", true); if (!left) set("left", true); if (!right) set("right", true); break;
        case "top": if (top) set("top", true); break;
        case "bottom": if (bottom) set("bottom", true); break;
        case "left": if (left) set("left", true); break;
        case "right": if (right) set("right", true); break;
      }
      return { ...cd, borders: b };
    });
  };

  // ----- Layout ----------------------------------------------------------------------------
  run("table.selectCell", () => ed()?.selectTablePart("cell"));
  run("table.selectColumn", () => ed()?.selectTablePart("column"));
  run("table.selectRow", () => ed()?.selectTablePart("row"));
  run("table.selectTable", () => ed()?.selectTablePart("table"));
  c.implement("table.select", { run: () => {}, enabled: inTable });
  c.implement("table.delete", { run: () => {}, enabled: inTable });
  run("table.deleteColumns", () => ed()?.deleteColumns());
  run("table.deleteRows", () => ed()?.deleteRows());
  run("table.deleteTable", () => ed()?.deleteTable());
  run("table.insertAbove", () => ed()?.insertRows("above"));
  run("table.insertBelow", () => ed()?.insertRows("below"));
  run("table.insertLeft", () => ed()?.insertColumns("left"));
  run("table.insertRight", () => ed()?.insertColumns("right"));
  run("table.merge", () => ed()?.mergeCells(), () => !!ed()?.cellBlock()?.multi);
  run("table.split", () => void splitDialog());
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
    align.value = def.align ?? "left";
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
        h("fieldset", null, h("legend", null, "Tabelle"), h("div", { class: "grid2" }, "Ausrichtung:", align, "Einzug von links:", indent)),
        h("fieldset", null, h("legend", null, "Zeile"), h("div", { class: "grid2" }, "Höhe angeben:", height, "Zeilenhöhe:", rule),
          h("label", { class: "check" }, cantSplit, "Zeilenwechsel auf Seiten zulassen"),
          h("label", { class: "check" }, header, "Gleiche Kopfzeile auf jeder Seite wiederholen"))),
      buttons: [{ label: "OK", value: true, primary: true }, { label: "Abbrechen", value: false }], cancel: false,
    });
    if (!ok) return;
    e.updateTable((d) => ({ ...d, align: align.value === "left" ? undefined : align.value as "center" | "right", indent: parseNum(indent.value) * CM || undefined }));
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
  const bordersMenu: MenuEntry[] = ["bottom", "top", "left", "right", "none", "all", "outside", "inside"].map((k) => ({ cmd: `border.${k}` }));

  const designTab: Tab = {
    id: "tableDesign", label: "Tabellenentwurf", key: "JT", contextual: true, shrink: ["borders", "styles", "options"],
    groups: [
      { id: "options", label: "Optionen für Tabellenformat", key: "ZO", icon: "table-style",
        controls: [stack(check("table.headerRow", "T"), check("table.totalRow", "E"), check("table.bandedRows", "R")),
          stack(check("table.firstColumn", "S"), check("table.lastColumn", "L"), check("table.bandedColumns", "B"))] },
      { id: "styles", label: "Tabellenformatvorlagen", key: "ZT", icon: "table-style",
        controls: [dd("table.styles", "Q", "large", [], stylesPanel),
          { type: "color", cmd: "table.shading", key: "H", palette: "font", current: () => lastShading, pick: (hex) => { if (hex !== "auto") lastShading = hex; shade(hex); } }] },
      { id: "borders", label: "Rahmen", key: "ZR", icon: "borders",
        controls: [{ type: "split", cmd: "table.borders", key: "B", size: "large", menu: bordersMenu }] },
    ],
  };

  const layoutTab: Tab = {
    id: "tableLayout", label: "Layout", key: "JL", contextual: true, shrink: ["data", "alignment", "size", "merge", "rowsCols", "table"],
    groups: [
      { id: "table", label: "Tabelle", key: "ZT", icon: "table",
        controls: [stack(dd("table.select", "K", "medium", [{ cmd: "table.selectCell" }, { cmd: "table.selectColumn" }, { cmd: "table.selectRow" }, { cmd: "table.selectTable" }]),
          b("table.gridlines", "TG", "medium"), b("table.properties", "O", "medium"))] },
      { id: "rowsCols", label: "Zeilen und Spalten", key: "ZZ", icon: "table-insert-row",
        controls: [dd("table.delete", "D", "large", [{ cmd: "table.deleteColumns" }, { cmd: "table.deleteRows" }, { cmd: "table.deleteTable" }]),
          b("table.insertAbove", "A", "large"), b("table.insertBelow", "BE", "large"),
          stack(b("table.insertLeft", "L", "medium"), b("table.insertRight", "R", "medium"))] },
      { id: "merge", label: "Zusammenführen", key: "ZM", icon: "table-merge",
        controls: [stack(b("table.merge", "M", "medium"), b("table.split", "P", "medium"))] },
      { id: "size", label: "Zellengröße", key: "ZG", icon: "table-cols-even",
        controls: [stack(
          { type: "spinner", cmd: "table.height", label: "Höhe:", key: "H", value: heightValue, commit: setHeight, step: stepHeight },
          { type: "spinner", cmd: "table.width", label: "Breite:", key: "W", value: widthValue, commit: setWidth, step: stepWidth }),
          stack(b("table.distributeRows", "HZ", "medium"), b("table.distributeColumns", "HS", "medium"))] },
      { id: "alignment", label: "Ausrichtung", key: "ZA", icon: "cell-top-left",
        controls: [{ type: "rows", rows: [
          [b("table.alignTopLeft", "1"), b("table.alignTopCenter", "2"), b("table.alignTopRight", "3")],
          [b("table.alignMiddleLeft", "4"), b("table.alignMiddleCenter", "5"), b("table.alignMiddleRight", "6")],
          [b("table.alignBottomLeft", "7"), b("table.alignBottomCenter", "8"), b("table.alignBottomRight", "9")],
        ] }] },
      { id: "data", label: "Daten", key: "ZD", icon: "table-repeat-header",
        controls: [b("table.repeatHeader", "W", "large")] },
    ],
  };

  return { insertPanel, designTab, layoutTab, cellWidthPt: () => (ed()?.cellWidth() ?? 0) / PT };
}

export type { CellDef };
