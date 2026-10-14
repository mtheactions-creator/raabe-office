// Start > Absatz > Sortieren: Word's dialog Text sortieren (recorded 2026-10-14) with
// Sortieren nach / Dann nach / Dann nach (field, Typ, Aufsteigend/Absteigend), Meine Liste
// enthält (Überschrift / Keine Überschrift) and Optionen... (Trennzeichen, Groß-/Kleinschreibung,
// Sortiersprache).

import { h, messageBox } from "@raabe/shell";
import { sortDate, sortNumber, type SortKey, type SortType } from "@raabe/text-engine";
import type { AppApi } from "./api";
import { openModal } from "./modal";

const TYPES: [SortType, string][] = [["text", "Text"], ["number", "Zahl"], ["date", "Datum"]];
const LANGS: [string, string][] = [["de-DE", "Deutsch (Deutschland)"], ["en-US", "Englisch (USA)"], ["en-GB", "Englisch (Vereinigtes Königreich)"], ["fr-FR", "Französisch (Frankreich)"], ["es-ES", "Spanisch (Spanien)"]];

export function sortFeature(api: AppApi) {
  const ed = () => api.editor();
  // Options last used in the session, as Word keeps them.
  const opts = { separator: "\t", caseSensitive: false, lang: "de-DE" };

  const dialog = () => {
    const e = ed();
    if (!e) return;
    let preview = e.sortPreview(opts.separator);
    let header = preview.headerRow;
    const fieldSel = () => h("select", { style: "width:250px" }) as HTMLSelectElement;
    const typeSel = () => h("select", { style: "width:150px" }, ...TYPES.map(([v, l]) => h("option", { value: v }, l))) as HTMLSelectElement;
    const rows = [0, 1, 2].map((i) => {
      const field = fieldSel(), type = typeSel();
      const name = `sortdir${i}`;
      const asc = h("input", { type: "radio", name, checked: true }) as HTMLInputElement;
      const desc = h("input", { type: "radio", name }) as HTMLInputElement;
      const el = h("fieldset", {}, h("legend", {}, i ? "Dann nach" : "Sortieren nach"),
        h("div", { class: "sortrow" }, field, h("label", { class: "field row" }, "Typ:", type),
          h("div", {}, h("label", { class: "check" }, asc, "Aufsteigend"), h("label", { class: "check" }, desc, "Absteigend"))));
      return { field, type, asc, desc, el };
    });
    const withHeader = h("input", { type: "radio", name: "sorthead" }) as HTMLInputElement;
    const noHeader = h("input", { type: "radio", name: "sorthead" }) as HTMLInputElement;
    const count = () => Math.max(1, ...preview.entries.map((x) => x.length - 1));
    // The fields: Absätze or Feld n for paragraphs, Spalte n in a table; with Überschrift the
    // first entry's texts name them, as in Word.
    const names = (): [string, string][] => {
      const n = count();
      const first = preview.entries[0] ?? [];
      const label = (i: number) => (header && first[i]?.trim() ? first[i].trim() : preview.table ? `Spalte ${i}` : `Feld ${i}`);
      const out: [string, string][] = preview.table ? [] : [["0", "Absätze"]];
      if (preview.table || n > 1) for (let i = 1; i <= n; i++) out.push([String(i), label(i)]);
      return out;
    };
    const guess = (field: number): SortType => {
      const vals = preview.entries.slice(header ? 1 : 0).map((x) => (x[field] ?? "").trim()).filter(Boolean);
      if (vals.length && vals.every((v) => !Number.isNaN(sortDate(v)))) return "date";
      if (vals.length && vals.every((v) => !Number.isNaN(sortNumber(v)) && /^[\s\d.,+\-−%€$]+$/.test(v))) return "number";
      return "text";
    };
    const fill = () => {
      const list = names();
      rows.forEach((r, i) => {
        const keep = r.field.value;
        r.field.replaceChildren(...(i ? [h("option", { value: "" }, "")] : []), ...list.map(([v, l]) => h("option", { value: v }, l)));
        r.field.value = list.some(([v]) => v === keep) ? keep : i ? "" : list[0]?.[0] ?? "0";
        // Dann nach: only with fields (Word greys it for whole paragraphs).
        const off = i > 0 && (list.length < 2 || !rows[i - 1].field.value);
        for (const x of [r.field, r.type, r.asc, r.desc]) x.disabled = off;
        r.el.classList.toggle("disabled", off);
        if (!r.field.value && i) r.type.value = "text";
      });
      withHeader.checked = header;
      noHeader.checked = !header;
    };
    rows.forEach((r) => r.field.addEventListener("change", () => { if (r.field.value) r.type.value = guess(Number(r.field.value)); fill(); }));
    withHeader.onchange = () => { header = true; fill(); };
    noHeader.onchange = () => { header = false; fill(); };
    fill();
    rows[0].type.value = guess(Number(rows[0].field.value));

    const options = () => {
      const tabs = h("input", { type: "radio", name: "sortsep", checked: opts.separator === "\t" }) as HTMLInputElement;
      const commas = h("input", { type: "radio", name: "sortsep", checked: opts.separator === "," }) as HTMLInputElement;
      const other = h("input", { type: "radio", name: "sortsep", checked: opts.separator !== "\t" && opts.separator !== "," }) as HTMLInputElement;
      const otherText = h("input", { type: "text", maxLength: 1, style: "width:40px", value: opts.separator !== "\t" && opts.separator !== "," ? opts.separator : "-" }) as HTMLInputElement;
      const caseBox = h("input", { type: "checkbox", checked: opts.caseSensitive }) as HTMLInputElement;
      const column = h("input", { type: "checkbox", disabled: true }) as HTMLInputElement;
      const lang = h("select", {}, ...LANGS.map(([v, l]) => h("option", { value: v }, l))) as HTMLSelectElement;
      lang.value = opts.lang;
      openModal("Sortieroptionen", h("div", { class: "sortopts" },
        h("fieldset", {}, h("legend", {}, "Trennzeichen"),
          h("div", { class: "row" }, h("label", { class: "check" }, tabs, "Tabstopps"), h("label", { class: "check" }, commas, "Kommas"),
            h("label", { class: "check" }, other, "Andere:"), otherText)),
        h("fieldset", {}, h("legend", {}, "Sortieroptionen"),
          h("label", { class: "check disabled", "data-tip": "Nur Spalte sortieren\nNur bei einer markierten Spalte." }, column, "Nur Spalte sortieren"),
          h("label", { class: "check" }, caseBox, "Groß-/Kleinschreibung beachten")),
        h("fieldset", {}, h("legend", {}, "Sortiersprache"), lang)), [
        { label: "OK", primary: true, run: () => {
          opts.separator = tabs.checked ? "\t" : commas.checked ? "," : otherText.value.slice(0, 1) || "-";
          opts.caseSensitive = caseBox.checked;
          opts.lang = lang.value;
          preview = e.sortPreview(opts.separator);
          fill();
          return true;
        } },
        { label: "Abbrechen", run: () => true },
      ], { width: 420 });
    };

    const body = h("div", { class: "sortdlg" }, ...rows.map((r) => r.el),
      h("fieldset", {}, h("legend", {}, "Meine Liste enthält"),
        h("div", { class: "row" }, h("label", { class: "check" }, withHeader, "Überschrift"), h("label", { class: "check" }, noHeader, "Keine Überschrift"))));
    openModal("Text sortieren", body, [
      { label: "Optionen...", left: true, run: () => { options(); } },
      { label: "OK", primary: true, run: () => {
        const keys: SortKey[] = rows.filter((r) => !r.field.disabled && r.field.value !== "")
          .map((r) => ({ field: Number(r.field.value), type: r.type.value as SortType, desc: r.desc.checked }));
        const why = e.sortText({ keys, header, separator: opts.separator, caseSensitive: opts.caseSensitive, lang: opts.lang });
        if (why) void messageBox("RaabeWrite", why);
        api.refresh();
        api.focus();
        return true;
      } },
      { label: "Abbrechen", run: () => true },
    ], { width: 640, cls: "sortdialog" });
  };

  api.commands.implement("para.sort", { run: dialog, enabled: () => !!ed() && ed()!.doc.activeStory === "main" });
  return { dialog };
}
