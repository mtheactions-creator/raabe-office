// Referenzen: Inhaltsverzeichnis (gallery, Text hinzufügen, Tabelle aktualisieren) and
// Fußnoten (Fußnote/Endnote einfügen, Nächste Fußnote).

import { h, showDialog, type Command, type MenuEntry } from "@raabe/shell";
import type { AppApi } from "./api";

const cmd = (id: string, label: string, icon?: string, extra: Partial<Command> = {}): Command => ({ id, label, icon, ...extra });

export const REFERENCE_COMMANDS: Command[] = [
  cmd("toc.auto1", "Automatische Tabelle 1", "toc"),
  cmd("toc.auto2", "Automatische Tabelle 2", "toc"),
  cmd("toc.remove", "Inhaltsverzeichnis entfernen", "remove"),
  cmd("addText.none", "Nicht im Inhaltsverzeichnis anzeigen"),
  cmd("addText.1", "Ebene 1"),
  cmd("addText.2", "Ebene 2"),
  cmd("addText.3", "Ebene 3"),
  cmd("ref.prevFootnote", "Vorherige Fußnote", "previous"),
  cmd("ref.nextEndnote", "Nächste Endnote", "next"),
  cmd("ref.prevEndnote", "Vorherige Endnote", "previous"),
];

/** A small picture of a table of contents for the gallery (as Word's previews). */
function tocPreview(title: string): HTMLElement {
  const row = (text: string, indent: number, size: number) => h("div", {
    style: `display:flex;gap:2px;padding-left:${indent}px;font:${size}px Calibri,Carlito,sans-serif;line-height:1.5`,
  }, h("span", null, text), h("span", { style: "flex:1;border-bottom:1px dotted #000;margin-bottom:4px" }), h("span", null, "1"));
  return h("div", { style: "width:230px;padding:6px 10px;background:#fff;border:1px solid #e1dfdd" },
    h("div", { style: "font:16px 'Calibri Light',Carlito,sans-serif;color:#2f5496;margin-bottom:4px" }, title),
    row("Überschrift 1", 0, 11), row("Überschrift 2", 8, 11), row("Überschrift 3", 16, 11));
}

export function referencesFeature(api: AppApi) {
  const ed = () => api.editor();
  const c = api.commands;
  const has = () => !!ed();
  const run = (id: string, fn: () => void, enabled = has, checked?: () => boolean) => c.implement(id, { run: () => { fn(); api.focus(); }, enabled, checked });
  const inMain = () => !!ed() && !ed()!.isHeaderFooter() && ed()!.doc.activeStory === "main";

  // ----- table of contents ---------------------------------------------------------------
  c.implement("ref.toc", { run: () => {}, enabled: has });
  run("toc.auto1", () => ed()?.insertToc("Inhalt"), inMain);
  run("toc.auto2", () => ed()?.insertToc("Inhaltsverzeichnis"), inMain);
  run("toc.remove", () => ed()?.removeToc(), () => !!ed()?.hasToc());
  c.implement("ref.updateToc", { run: () => void updateDialog(), enabled: () => !!ed()?.hasToc() });
  const updateDialog = async () => {
    const e = ed();
    if (!e) return;
    const pages = h("input", { type: "radio", name: "tocupd", checked: true });
    const all = h("input", { type: "radio", name: "tocupd" });
    const ok = await showDialog({
      title: "Inhaltsverzeichnis aktualisieren",
      body: h("div", { style: "width:330px" }, h("p", { style: "margin:0 0 8px" }, "Inhaltsverzeichnis wird aktualisiert. Wählen Sie eine der folgenden Optionen:"),
        h("label", { class: "check" }, pages, "Nur Seitenzahlen aktualisieren"),
        h("label", { class: "check" }, all, "Gesamtes Verzeichnis aktualisieren")),
      buttons: [{ label: "OK", value: true, primary: true }, { label: "Abbrechen", value: false }], cancel: false,
    });
    if (ok) e.updateToc(pages.checked);
    api.focus();
  };

  // Text hinzufügen: the paragraph's level in the table of contents (a heading style).
  c.implement("ref.addText", { run: () => {}, enabled: inMain });
  const level = () => {
    const e = ed();
    if (!e) return 0;
    const hd = e.headings().find((x) => x.p === e.sel.head.p);
    return hd ? hd.level : 0;
  };
  run("addText.none", () => ed()?.applyParaStyle(ed()!.styleId("Normal")), inMain, () => level() === 0);
  for (const k of [1, 2, 3]) run(`addText.${k}`, () => ed()?.applyParaStyle(ed()!.styleId(`heading ${k}`)), inMain, () => level() === k);

  // ----- notes -------------------------------------------------------------------------------
  run("ref.footnote", () => ed()?.insertNote("fn"), inMain);
  run("ref.endnote", () => ed()?.insertNote("en"), inMain);
  run("ref.nextFootnote", () => ed()?.nextNote(1, "fn"));
  run("ref.prevFootnote", () => ed()?.nextNote(-1, "fn"));
  run("ref.nextEndnote", () => ed()?.nextNote(1, "en"));
  run("ref.prevEndnote", () => ed()?.nextNote(-1, "en"));

  const tocMenu: MenuEntry[] = [
    { heading: "Integriert" },
    { label: "Automatische Tabelle 1", run: () => void c.run("toc.auto1"), preview: tocPreview("Inhalt") },
    { label: "Automatische Tabelle 2", run: () => void c.run("toc.auto2"), preview: tocPreview("Inhaltsverzeichnis") },
    "-",
    { cmd: "toc.remove" },
  ];
  const checkedEntry = (id: string): MenuEntry => ({
    label: c.get(id).label, run: () => void c.run(id),
    get checked() { return c.isChecked(id); },
    get disabled() { return !c.isEnabled(id); },
  });
  const addTextMenu: MenuEntry[] = [checkedEntry("addText.none"), checkedEntry("addText.1"), checkedEntry("addText.2"), checkedEntry("addText.3")];
  const nextNoteMenu: MenuEntry[] = [{ cmd: "ref.nextFootnote" }, { cmd: "ref.prevFootnote" }, { cmd: "ref.nextEndnote" }, { cmd: "ref.prevEndnote" }];
  return { tocMenu, addTextMenu, nextNoteMenu };
}
