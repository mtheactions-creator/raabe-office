// Referenzen > Fußnoten > Fuß- und Endnote (dialog launcher), as Word's dialog (recorded with
// probe_word_dialogs.py): Speicherort, Konvertieren, Format (Zahlenformat, Beginnen mit,
// Nummerierung), Einfügen and Übernehmen. Settings are the document's (w:footnotePr/endnotePr).

import { h } from "@raabe/shell";
import type { NoteProps } from "@raabe/text-engine";
import type { AppApi } from "./api";
import { openModal } from "./modal";

const FORMATS: [string, string][] = [["decimal", "1, 2, 3, …"], ["lowerLetter", "a, b, c, …"], ["upperLetter", "A, B, C, …"], ["lowerRoman", "i, ii, iii, …"], ["upperRoman", "I, II, III, …"], ["chicago", "*, †, ‡, §, …"]];

const select = (opts: [string, string][], value: string) => {
  const s = h("select", {}, ...opts.map(([v, l]) => h("option", { value: v }, l))) as HTMLSelectElement;
  s.value = value;
  return s;
};

export function notesDialogFeature(api: AppApi) {
  const ed = () => api.editor();

  const dialog = () => {
    const e = ed();
    if (!e) return;
    const cur = e.noteSettings();
    const props = { fn: { ...cur.footnote }, en: { ...cur.endnote } };
    let kind: "fn" | "en" = "fn";
    const fnRadio = h("input", { type: "radio", name: "notekind", checked: true }) as HTMLInputElement;
    const enRadio = h("input", { type: "radio", name: "notekind" }) as HTMLInputElement;
    const fnPos = select([["pageBottom", "Seitenende"], ["beneathText", "Unterhalb des Texts"]], props.fn.pos ?? "pageBottom");
    const enPos = select([["docEnd", "Ende des Dokuments"], ["sectEnd", "Ende des Abschnitts"]], props.en.pos ?? "docEnd");
    const columns = select([["auto", "An Abschnittslayout anpassen"]], "auto");
    const fmt = select(FORMATS, "decimal");
    const custom = h("input", { type: "text", style: "width:80px", disabled: true, "data-tip": "Benutzerdefiniert\nRaabeWrite nummeriert alle Notizen." }) as HTMLInputElement;
    const start = h("input", { type: "number", min: "1", style: "width:70px" }) as HTMLInputElement;
    const restart = select([["continuous", "Fortlaufend"], ["eachSect", "Bei jedem Abschnitt neu beginnen"], ["eachPage", "Bei jeder Seite neu beginnen"]], "continuous");
    const apply = select([["doc", "Gesamtes Dokument"]], "doc");
    const show = () => {
      const p = props[kind];
      fnPos.disabled = kind !== "fn";
      enPos.disabled = kind !== "en";
      fmt.value = p.numFmt ?? (kind === "fn" ? "decimal" : "lowerRoman");
      start.value = String(p.numStart ?? 1);
      // Endnotes cannot restart on every page (Word offers it for footnotes only).
      restart.replaceChildren(...([["continuous", "Fortlaufend"], ["eachSect", "Bei jedem Abschnitt neu beginnen"], ...(kind === "fn" ? [["eachPage", "Bei jeder Seite neu beginnen"]] : [])] as [string, string][])
        .map(([v, l]) => h("option", { value: v }, l)));
      restart.value = p.numRestart ?? "continuous";
    };
    const keep = () => {
      const p = props[kind];
      p.numFmt = fmt.value;
      p.numStart = Math.max(1, Number(start.value) || 1);
      p.numRestart = restart.value as NoteProps["numRestart"];
      props.fn.pos = fnPos.value as NoteProps["pos"];
      props.en.pos = enPos.value as NoteProps["pos"];
    };
    fnRadio.onchange = () => { keep(); kind = "fn"; show(); };
    enRadio.onchange = () => { keep(); kind = "en"; show(); };
    const clean = (p: NoteProps, k: "fn" | "en"): NoteProps | undefined => {
      const out: NoteProps = {
        ...(p.pos && p.pos !== (k === "fn" ? "pageBottom" : "docEnd") ? { pos: p.pos } : {}),
        ...(p.numFmt && p.numFmt !== (k === "fn" ? "decimal" : "lowerRoman") ? { numFmt: p.numFmt } : {}),
        ...(p.numStart && p.numStart !== 1 ? { numStart: p.numStart } : {}),
        ...(p.numRestart && p.numRestart !== "continuous" ? { numRestart: p.numRestart } : {}),
      };
      return Object.keys(out).length ? out : undefined;
    };
    const save = () => {
      keep();
      e.setNoteSettings({ footnote: clean(props.fn, "fn"), endnote: clean(props.en, "en") });
      api.refresh();
    };
    const convert = () => {
      const a = h("input", { type: "radio", name: "conv", checked: true }) as HTMLInputElement;
      const b = h("input", { type: "radio", name: "conv" }) as HTMLInputElement;
      const c = h("input", { type: "radio", name: "conv" }) as HTMLInputElement;
      openModal("Notizen konvertieren", h("div", {},
        h("label", { class: "check" }, a, "Alle Fußnoten in Endnoten umwandeln"),
        h("label", { class: "check" }, b, "Alle Endnoten in Fußnoten umwandeln"),
        h("label", { class: "check" }, c, "Fuß- und Endnoten austauschen")), [
        { label: "OK", primary: true, run: () => { e.convertNotes(a.checked ? "toEndnotes" : b.checked ? "toFootnotes" : "swap"); api.refresh(); return true; } },
        { label: "Abbrechen", run: () => true },
      ], { width: 380 });
    };
    const body = h("div", { class: "notedialog" },
      h("fieldset", {}, h("legend", {}, "Speicherort"),
        h("div", { class: "grid2" }, h("label", { class: "check" }, fnRadio, "Fußnoten:"), fnPos, h("label", { class: "check" }, enRadio, "Endnoten:"), enPos),
        h("div", { class: "row end" }, h("button", { onclick: convert }, "Konvertieren..."))),
      h("fieldset", {}, h("legend", {}, "Fußnotenlayout"), h("label", { class: "field row" }, "Spalten:", columns)),
      h("fieldset", {}, h("legend", {}, "Format"),
        h("div", { class: "grid2" }, "Zahlenformat:", fmt, "Benutzerdefiniert:", h("div", { class: "row" }, custom, h("button", { disabled: true }, "Symbol...")),
          "Beginnen mit:", start, "Nummerierung:", restart)),
      h("fieldset", {}, h("legend", {}, "Änderungen übernehmen"), h("label", { class: "field row" }, "Änderungen übernehmen für:", apply)));
    openModal("Fuß- und Endnote", body, [
      { label: "Einfügen", primary: true, run: () => { save(); e.insertNote(kind); api.refresh(); return true; } },
      { label: "Abbrechen", run: () => true },
      { label: "Übernehmen", run: () => { save(); return true; } },
    ], { width: 520, cls: "notedlg" });
    show();
  };

  api.commands.implement("ref.footnoteDialog", { run: dialog, enabled: () => !!ed() });

  // Notizen anzeigen: with both kinds in the document, Word asks which area to show.
  const showNotes = () => {
    const e = ed();
    if (!e) return;
    const story = e.doc.activeStory;
    const k = e.noteKinds();
    if (story.startsWith("fn:") || story.startsWith("en:") || !(k.fn && k.en)) {
      e.showNotes(k.fn ? "fn" : "en");
      api.refresh();
      api.focus();
      return;
    }
    const fn = h("input", { type: "radio", name: "shownotes", checked: true }) as HTMLInputElement;
    const en = h("input", { type: "radio", name: "shownotes" }) as HTMLInputElement;
    openModal("Notizen anzeigen", h("div", {},
      h("p", { style: "margin:0 0 8px" }, "Das Dokument enthält Fußnoten und Endnoten."),
      h("label", { class: "check" }, fn, "Fußnotenbereich anzeigen"),
      h("label", { class: "check" }, en, "Endnotenbereich anzeigen")), [
      { label: "OK", primary: true, run: () => { e.showNotes(fn.checked ? "fn" : "en"); api.refresh(); api.focus(); return true; } },
      { label: "Abbrechen", run: () => true },
    ], { width: 360 });
  };
  api.commands.implement("ref.showNotes", { run: showNotes, enabled: () => { const k = ed()?.noteKinds(); return !!k && (k.fn || k.en); } });
  return { dialog };
}
