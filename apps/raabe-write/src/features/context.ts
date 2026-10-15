// The context menu (right click, menu key) as Word 365 shows it (recorded 2026-10-14 with
// tools/feel-lab/probe_word_context.py): a search box, Ausschneiden/Kopieren, the paste
// options, then entries for what is under the mouse (text, list, table cell, several cells,
// picture) with Word's submenus. Spelling suggestions come first for a misspelled word.

import type { Editor, Pos } from "@raabe/text-engine";
import { closePopup, h, showDialog, showMenu, type Command, type MenuEntry } from "@raabe/shell";
import { CM, fmtCm, parseNum, type AppApi } from "./api";

const cmd = (id: string, label: string, icon?: string, extra: Partial<Command> = {}): Command => ({ id, label, icon, ...extra });

export const CONTEXT_COMMANDS: Command[] = [
  cmd("edit.pasteKeep", "Ursprüngliche Formatierung beibehalten (K)", "paste-keep"),
  cmd("edit.pasteMerge", "Formatierung zusammenführen (Z)", "paste-merge"),
  cmd("list.adjustIndent", "Listeneinzug anpassen...", "list-indent"),
  cmd("list.separate", "Separate Liste"),
  cmd("list.restart", "Neu beginnen mit 1"),
  cmd("list.joinPrevious", "Mit vorheriger Liste verknüpfen"),
  cmd("list.continue", "Nummerierung fortsetzen"),
  cmd("list.setValue", "Nummerierungswert festlegen...", "number-value"),
  cmd("mini.styles", "Formatvorlagen", "styles"),
  cmd("mini.tableInsert", "Einfügen", "table-insert-row"),
  cmd("mini.tableDelete", "Löschen", "table-delete"),
];

/** Submenu entries of commands (the command's own label). */
const cmds = (...ids: string[]): MenuEntry[] => ids.map((id) => (id === "-" ? "-" : { cmd: id }));

export function contextFeature(api: AppApi, spellEntries: (e: Editor, pos: Pos) => Promise<MenuEntry[]>,
  toolbar: (kind: "text" | "table" | "picture") => HTMLElement | null,
  /** Word's menu inside an equation being edited (null outside one). */
  mathEntries: (x: number, y: number) => MenuEntry[] | null = () => null,
  /** Synonyme: the thesaurus' words for the word clicked. */
  synonymEntries: (e: Editor, pos: Pos) => Promise<MenuEntry[]> = async () => []) {
  const c = api.commands;
  const ed = () => api.editor();
  const has = () => !!ed();
  c.implement("mini.styles", { run: () => {}, enabled: has });
  c.implement("mini.tableInsert", { run: () => {}, enabled: () => !!ed()?.inTable() });
  c.implement("mini.tableDelete", { run: () => {}, enabled: () => !!ed()?.inTable() });
  const inList = () => !!ed()?.listKind();

  const paste = (mode: "keep" | "merge" | "text") => async () => {
    const e = ed();
    if (!e) return;
    const clip = await api.clipboard();
    if (mode === "text") { if (clip.text) e.pasteText(clip.text, true); }
    else if (mode === "merge") e.pasteMerged(clip.text, clip.html);
    else e.paste(clip.text, clip.html, clip.image);
    api.focus();
  };
  c.implement("edit.pasteKeep", { run: paste("keep"), enabled: has });
  c.implement("edit.pasteMerge", { run: paste("merge"), enabled: has });

  // ----- lists ---------------------------------------------------------------------------
  c.implement("list.separate", { run: () => { ed()?.restartNumbering(); api.focus(); }, enabled: inList });
  c.implement("list.restart", { run: () => { ed()?.restartNumbering(); api.focus(); }, enabled: inList });
  c.implement("list.joinPrevious", { run: () => { ed()?.continueNumbering(); api.focus(); }, enabled: () => !!ed()?.canContinueNumbering() });
  c.implement("list.continue", { run: () => { ed()?.continueNumbering(); api.focus(); }, enabled: () => !!ed()?.canContinueNumbering() });

  /** Listeneinzug anpassen: number position, text indent, what follows the number. */
  const indentDialog = async () => {
    const e = ed();
    const cur = e?.listLevelIndent();
    if (!e || !cur) return;
    const pos = h("input", { type: "text", value: fmtCm(cur.numberPos), style: "width:90px" });
    const text = h("input", { type: "text", value: fmtCm(cur.textIndent), style: "width:90px" });
    const after = h("select", null, ...([["tab", "Tabstoppzeichen"], ["space", "Leerzeichen"], ["nothing", "Nichts"]] as const)
      .map(([v, l]) => h("option", { value: v, selected: cur.suffix === v }, l)));
    const ok = await showDialog({
      title: "Listeneinzug anpassen",
      body: h("div", { class: "grid2", style: "width:320px" },
        cur.bullet ? "Position des Aufzählungszeichens:" : "Position der Nummer:", pos, "Texteinzug:", text, "Text danach:", after),
      buttons: [{ label: "OK", value: true, primary: true }, { label: "Abbrechen", value: false }], cancel: false,
    });
    if (!ok) return;
    const cm = (s: string, d: number) => { const v = parseNum(s); return Number.isFinite(v) ? Math.max(0, v * CM) : d; };
    e.setListLevelIndent(cm(pos.value, cur.numberPos), cm(text.value, cur.textIndent), after.value as "tab" | "space" | "nothing");
    api.focus();
  };
  c.implement("list.adjustIndent", { run: () => void indentDialog(), enabled: inList });

  /** Nummerierungswert festlegen: a new list from a value, or continuing the one before. */
  const valueDialog = async () => {
    const e = ed();
    if (!e) return;
    const fresh = h("input", { type: "radio", name: "numval", checked: true });
    const cont = h("input", { type: "radio", name: "numval", disabled: !e.canContinueNumbering() });
    const value = h("input", { type: "number", min: "0", max: "9999", value: "1", style: "width:80px" });
    const ok = await showDialog({
      title: "Nummerierungswert festlegen",
      body: h("div", { style: "width:300px;display:flex;flex-direction:column;gap:8px" },
        h("label", null, fresh, " Neue Liste beginnen"), h("label", null, cont, " Vorherige Liste fortsetzen"),
        h("div", { class: "grid2" }, "Wert festlegen auf:", value)),
      buttons: [{ label: "OK", value: true, primary: true }, { label: "Abbrechen", value: false }], cancel: false,
    });
    if (!ok) return;
    if (cont.checked) e.continueNumbering();
    else e.restartNumbering(Math.max(0, Math.round(Number(value.value) || 1)));
    api.focus();
  };
  c.implement("list.setValue", { run: () => void valueDialog(), enabled: () => ed()?.listKind() === "number" || ed()?.listKind() === "multilevel" });

  // ----- the menu ---------------------------------------------------------------------------
  const pasteOptions: MenuEntry = {
    label: "Einfügeoptionen:", icons: [
      { icon: "paste-keep", tip: "Ursprüngliche Formatierung beibehalten (K)", run: () => void c.run("edit.pasteKeep") },
      { icon: "paste-merge", tip: "Formatierung zusammenführen (Z)", run: () => void c.run("edit.pasteMerge") },
      { icon: "paste-text", tip: "Nur den Text übernehmen (T)", run: () => void c.run("edit.pasteText") },
    ],
  };

  const entriesFor = (e: Editor, synonymList: MenuEntry[]): MenuEntry[] => {
    const out: MenuEntry[] = [];
    const image = e.selectedImage();
    const block = e.cellBlock();
    const cells = !!block?.multi;
    const synonyms: MenuEntry = { label: "Synonyme", submenu: () => synonymList.length ? synonymList : [{ label: "(Keine Vorschläge)", run: () => {}, disabled: true }, "-", { cmd: "review.thesaurus", label: "Thesaurus..." }] };
    const insertSub: MenuEntry = { label: "Einfügen", submenu: [
      { cmd: "table.insertLeft", label: "Spalten links einfügen" }, { cmd: "table.insertRight", label: "Spalten rechts einfügen" },
      { cmd: "table.insertAbove", label: "Zeilen oberhalb einfügen" }, { cmd: "table.insertBelow", label: "Zeilen unterhalb einfügen" },
      { cmd: "table.insertCells", label: "Zellen einfügen..." },
    ] };
    const borders: MenuEntry = { label: "Rahmenarten", submenu: cmds("border.bottom", "border.top", "border.left", "border.right", "border.none", "border.all", "border.outside", "border.inside") };
    out.push({ cmd: "edit.cut" }, { cmd: "edit.copy" }, pasteOptions, "-");
    // A shape or text box (its frame selected): Word's menu (recorded 2026-10-15,
    // probe_word_context.py objects).
    const shape = e.selectedShape();
    if (shape && !shape.editing) {
      const wrap = cmds("picture.wrapInline", "-", "picture.wrapSquare", "picture.wrapTight", "picture.wrapThrough", "picture.wrapTopBottom", "-", "picture.wrapBehind", "picture.wrapFront");
      const hasText = !!shape.obj.shape?.story;
      const textbox = shape.obj.shape?.kind === "textbox";
      out.push(
        { label: hasText ? "Text bearbeiten" : "Text hinzufügen", icon: "text-box", run: () => { e.editShapeText(shape.p, shape.index, "end"); api.refresh(); }, disabled: shape.obj.shape?.kind === "connector" },
        { label: "Punkte bearbeiten", icon: "edit-shape", run: () => {}, disabled: true },
        "-",
        { label: "Gruppieren", icon: "group-objects", submenu: cmds("layout.group") },
        { label: "In den Vordergrund", icon: "bring-forward", submenu: cmds("layout.bringForward") },
        { label: "In den Hintergrund", icon: "send-backward", submenu: cmds("layout.sendBackward") },
        "-",
        { cmd: "insert.link", label: "Link" },
        { cmd: "ref.caption", label: "Beschriftung einfügen..." },
        { label: "Textumbruch", icon: "wrap-text", submenu: [...wrap, "-", { cmd: "picture.moreLayout", label: "Weitere Layoutoptionen..." }] },
        "-",
        { cmd: "picture.altText", label: "Alternativtext anzeigen..." },
        { label: textbox ? "Als Standardtextfeld-Format festlegen" : "Als Standardform festlegen", run: () => {}, disabled: true },
        { cmd: "picture.moreLayout", label: "Weitere Layoutoptionen..." },
        { cmd: "shape.formatPane", label: "Form formatieren..." },
      );
      return out;
    }
    if (image) {
      out.push(
        { label: "Bild ändern", icon: "change-picture", submenu: [{ cmd: "picture.changeFromFile", label: "Dieses Gerät..." }, { cmd: "picture.changeFromClipboard", label: "Aus Zwischenablage..." }] },
        "-",
        { label: "Gruppieren", icon: "group-objects", submenu: cmds("layout.group"), disabled: image.kind === "inline" },
        { label: "In den Vordergrund", icon: "bring-forward", submenu: cmds("layout.bringForward"), disabled: image.kind === "inline" },
        { label: "In den Hintergrund", icon: "send-backward", submenu: cmds("layout.sendBackward"), disabled: image.kind === "inline" },
        "-",
        { cmd: "insert.link", label: "Link" },
        { cmd: "picture.saveAs", label: "Als Grafik speichern..." },
        "-",
        { cmd: "ref.caption", label: "Beschriftung einfügen..." },
        { label: "Textumbruch", icon: "wrap-text", submenu: [
          ...cmds("picture.wrapInline", "-", "picture.wrapSquare", "picture.wrapTight", "picture.wrapThrough", "picture.wrapTopBottom", "-", "picture.wrapBehind", "picture.wrapFront"),
          "-", { cmd: "picture.moreLayout", label: "Weitere Layoutoptionen..." }, { cmd: "picture.setDefaultLayout", label: "Als Standardlayout festlegen" },
        ] },
        "-",
        { cmd: "picture.altText", label: "Alternativtext anzeigen..." },
        { cmd: "picture.sizePosition", label: "Größe und Position..." },
        { cmd: "picture.formatPane", label: "Grafik formatieren..." },
      );
      return out;
    }
    if (cells) {
      out.push(insertSub, { cmd: "table.deleteTable", label: "Tabelle löschen" });
      if (!e.wholeTableSelected()) out.splice(out.length - 1, 1, { cmd: "table.deleteCells", label: "Zellen löschen..." }, { cmd: "table.merge", label: "Zellen verbinden" });
      out.push({ cmd: "table.distributeRows", label: "Zeilen gleichmäßig verteilen" }, { cmd: "table.distributeColumns", label: "Spalten gleichmäßig verteilen" }, "-",
        borders,
        { label: "AutoAnpassen", submenu: [{ cmd: "table.autofitContents", label: "AutoAnpassen an Inhalt" }, { cmd: "table.autofitWindow", label: "Größe an Fenster anpassen" }, { cmd: "table.fixedWidth", label: "Feste Spaltenbreite" }] },
        { cmd: "table.textDirection", label: "Textrichtung..." },
        { cmd: "ref.caption", label: "Beschriftung einfügen..." },
        { cmd: "table.properties", label: "Tabelleneigenschaften..." }, "-",
        { cmd: "review.newComment", label: "Neuer Kommentar" });
      return out;
    }
    if (e.inTable()) {
      out.push(synonyms, { cmd: "review.readAloud", label: "Laut vorlesen" }, { cmd: "review.translate", label: "Übersetzen" }, "-",
        insertSub, { cmd: "table.deleteCells", label: "Zellen löschen..." }, { cmd: "table.split", label: "Zellen teilen..." }, "-",
        borders, { cmd: "table.textDirection", label: "Textrichtung..." }, { cmd: "table.properties", label: "Tabelleneigenschaften..." }, "-",
        { cmd: "insert.link", label: "Link" }, { cmd: "review.newComment", label: "Neuer Kommentar" });
      return out;
    }
    const kind = e.listKind();
    if (kind === "bullet") out.push({ cmd: "list.adjustIndent" }, { cmd: "list.separate" }, { cmd: "list.joinPrevious" }, "-");
    else if (kind) out.push({ cmd: "list.adjustIndent" }, { cmd: "list.restart" }, { cmd: "list.continue" }, { cmd: "list.setValue" }, "-");
    out.push({ cmd: "font.dialog", label: "Schriftart..." }, { cmd: "para.dialog", label: "Absatz..." }, "-",
      synonyms, { cmd: "review.readAloud", label: "Laut vorlesen" }, { cmd: "review.translate", label: "Übersetzen" }, "-",
      { cmd: "insert.link", label: "Link" }, { cmd: "review.newComment", label: "Neuer Kommentar" });
    return out;
  };

  const flat = (list: MenuEntry[]): MenuEntry[] => list.flatMap((x) => {
    if (x === "-" || "heading" in x || "icons" in x || "flyout" in x || "block" in x) return [];
    if ("submenu" in x) return flat(typeof x.submenu === "function" ? x.submenu() : x.submenu);
    return [x];
  });
  const labelOf = (x: MenuEntry) => (typeof x === "object" && "cmd" in x ? x.label ?? c.get(x.cmd).label : typeof x === "object" && "label" in x ? x.label : "");

  /** Shows the menu; "Menüs durchsuchen" filters it to the matching entries (submenus included). */
  const open = (all: MenuEntry[], anchor: DOMRect, query = "", mini: HTMLElement | null = null) => {
    const q = query.trim().toLowerCase();
    const list = q ? flat(all).filter((x) => labelOf(x).toLowerCase().includes(q)) : all;
    const input = h("input", { type: "search", class: "menu-search", placeholder: "Menüs durchsuchen", "aria-label": "Menüs durchsuchen" });
    input.value = query;
    input.addEventListener("input", () => open(all, anchor, input.value, mini));
    showMenu(c, list.length ? list : [{ label: "Keine Ergebnisse", run: () => {}, disabled: true }], anchor, mini ?? undefined, h("div", { class: "menu-search-row" }, input));
    if (query) { input.focus(); input.setSelectionRange(query.length, query.length); }
  };

  const show = async (e: Editor, x: number, y: number, pos: Pos) => {
    const math = mathEntries(x, y);
    if (math) { open(math, new DOMRect(x, y, 0, 0), "", toolbar("text")); return; }
    const [spell, synonyms] = await Promise.all([spellEntries(e, pos), synonymEntries(e, pos)]);
    const kind = e.selectedImage() ? "picture" : e.inTable() ? "table" : "text";
    open([...spell, ...entriesFor(e, synonyms)], new DOMRect(x, y, 0, 0), "", toolbar(kind));
  };

  const attach = (e: Editor) => {
    e.onContextMenu = (x, y, pos) => void show(e, x, y, pos);
  };

  return { attach, close: closePopup };
}
