// Referenzen > Beschriftungen as in Word (recorded with probe_word_dialogs.py): Beschriftung
// einfügen (Bezeichnung, Position, Bezeichnung nicht verwenden, Neue Bezeichnung, Bezeichnung
// löschen, Nummerierung), Abbildungsverzeichnis einfügen (Seitenzahlen anzeigen, rechtsbündig,
// Füllzeichen, Links anstelle von Seitenzahlen, Beschriftung, Bezeichnung und Nummer
// einschließen) and aktualisieren; Felder aktualisieren (F9).

import { h } from "@raabe/shell";
import type { AppApi } from "./api";
import { openModal } from "./modal";

const BUILT_IN = ["Abbildung", "Formel", "Tabelle"];
const FORMATS: [string, string][] = [["ARABIC", "1, 2, 3, …"], ["alphabetic", "a, b, c, …"], ["ALPHABETIC", "A, B, C, …"], ["roman", "i, ii, iii, …"], ["ROMAN", "I, II, III, …"]];

function ownLabels(): string[] {
  try { const v = JSON.parse(localStorage.getItem("raabe.caption.labels") ?? "[]"); return Array.isArray(v) ? v.filter((x) => typeof x === "string") : []; } catch { return []; }
}
function saveLabels(list: string[]): void {
  try { localStorage.setItem("raabe.caption.labels", JSON.stringify(list)); } catch { /* private mode */ }
}
const labels = () => [...BUILT_IN, ...ownLabels()].sort((a, b) => a.localeCompare(b, "de"));

const select = (opts: [string, string][], value: string) => {
  const s = h("select", {}, ...opts.map(([v, l]) => h("option", { value: v }, l))) as HTMLSelectElement;
  s.value = value;
  return s;
};
const check = (label: string, on: boolean) => {
  const input = h("input", { type: "checkbox", checked: on }) as HTMLInputElement;
  return { input, el: h("label", { class: "check" }, input, label) };
};

export function captionsFeature(api: AppApi) {
  const c = api.commands;
  const ed = () => api.editor();
  let lastLabel = "Abbildung";
  const formatOf = new Map<string, string>();

  const captionDialog = () => {
    const e = ed();
    if (!e) return;
    const object = !!e.selectedImage() || e.inTable();
    const label = select(labels().map((l) => [l, l]), lastLabel);
    const position = select([["above", "Über dem ausgewählten Element"], ["below", "Unter dem ausgewählten Element"]], "below");
    position.disabled = !object;
    const without = check("Bezeichnung nicht in der Beschriftung verwenden", false);
    const number = () => {
      const n = e.referenceTargets({ caption: label.value }).length + 1;
      const f = formatOf.get(label.value) ?? "ARABIC";
      return f === "ARABIC" ? String(n) : f === "alphabetic" ? String.fromCharCode(96 + Math.min(26, n)) : f === "ALPHABETIC" ? String.fromCharCode(64 + Math.min(26, n))
        : f === "roman" ? ["i", "ii", "iii", "iv", "v", "vi", "vii", "viii", "ix", "x"][n - 1] ?? String(n) : ["I", "II", "III", "IV", "V", "VI", "VII", "VIII", "IX", "X"][n - 1] ?? String(n);
    };
    const fixed = h("span", { class: "capfixed" });
    const text = h("input", { type: "text", class: "captext" }) as HTMLInputElement;
    const sync = () => {
      fixed.textContent = `${without.input.checked ? "" : `${label.value} `}${number()}`;
      deleteBtn.disabled = BUILT_IN.includes(label.value);
    };
    const deleteBtn = h("button", { onclick: () => {
      saveLabels(ownLabels().filter((l) => l !== label.value));
      label.replaceChildren(...labels().map((l) => h("option", { value: l }, l)));
      label.value = "Abbildung";
      sync();
    } }, "Bezeichnung löschen") as HTMLButtonElement;
    const newBtn = h("button", { onclick: () => {
      const input = h("input", { type: "text", style: "width:300px" }) as HTMLInputElement;
      openModal("Neue Bezeichnung", h("label", { class: "field" }, "Bezeichnung:", input), [
        { label: "OK", primary: true, run: () => {
          const v = input.value.trim();
          if (!v) return true;
          if (!labels().includes(v)) saveLabels([...ownLabels(), v]);
          label.replaceChildren(...labels().map((l) => h("option", { value: l }, l)));
          label.value = v;
          sync();
          return true;
        } },
        { label: "Abbrechen", run: () => true },
      ], { width: 360 });
    } }, "Neue Bezeichnung...");
    const numberingBtn = h("button", { onclick: () => {
      const fmt = select(FORMATS, formatOf.get(label.value) ?? "ARABIC");
      const chapter = check("Kapitelnummer einbeziehen", false);
      chapter.input.disabled = true;
      chapter.el.classList.add("disabled");
      chapter.el.setAttribute("data-tip", "Kapitelnummer einbeziehen\nIn RaabeWrite werden Beschriftungen durchgehend nummeriert.");
      openModal("Beschriftungsnummerierung", h("div", {}, h("label", { class: "field row" }, "Format:", fmt), chapter.el), [
        { label: "OK", primary: true, run: () => { formatOf.set(label.value, fmt.value); sync(); return true; } },
        { label: "Abbrechen", run: () => true },
      ], { width: 380 });
    } }, "Nummerierung...");
    const autoBtn = h("button", { disabled: true, "data-tip": "AutoBeschriftung...\nBeschriftungen werden in RaabeWrite nicht automatisch eingefügt." }, "AutoBeschriftung...");
    label.onchange = sync;
    without.input.onchange = sync;
    const body = h("div", { class: "capdialog" },
      h("div", {}, "Beschriftung:"), h("div", { class: "capline" }, fixed, text),
      h("fieldset", {}, h("legend", {}, "Optionen"),
        h("label", { class: "field row" }, "Bezeichnung:", label),
        h("label", { class: `field row${object ? "" : " disabled"}` }, "Position:", position)),
      without.el,
      h("div", { class: "row" }, newBtn, deleteBtn, numberingBtn),
      h("div", { class: "row" }, autoBtn));
    openModal("Beschriftung", body, [
      { label: "OK", primary: true, run: () => {
        lastLabel = label.value;
        const fmt = formatOf.get(label.value) ?? "ARABIC";
        e.insertCaption(label.value, text.value, position.value !== "above", without.input.checked, fmt);
        api.refresh();
        api.focus();
        return true;
      } },
      { label: "Abbrechen", run: () => true },
    ], { width: 520, cls: "capdlg" });
    sync();
    text.focus();
  };

  // ----- Abbildungsverzeichnis --------------------------------------------------------------------
  const tofDialog = () => {
    const e = ed();
    if (!e) return;
    const numbers = check("Seitenzahlen anzeigen", true);
    const right = check("Seitenzahlen rechtsbündig", true);
    const leader = select([["none", "(ohne)"], ["dot", "......."], ["hyphen", "-------"], ["underscore", "_______"]], "dot");
    const links = check("Links anstelle von Seitenzahlen", true);
    const formats = select([["template", "Von Vorlage"]], "template");
    const label = select([...labels().map((l): [string, string] => [l, l])], lastLabel);
    const include = check("Bezeichnung und Nummer einschließen", true);
    const preview = h("div", { class: "tofpreview" });
    const web = h("div", { class: "tofpreview web" });
    const draw = () => {
      const lines = [1, 2, 3, 4].map((n) => `${include.input.checked ? `${label.value} ${n}: ` : ""}Text`);
      const dots = leader.value === "dot" ? "." : leader.value === "hyphen" ? "-" : leader.value === "underscore" ? "_" : " ";
      preview.replaceChildren(...lines.map((l, i) => h("div", { class: "tofline" }, h("span", {}, l),
        numbers.input.checked ? h("span", { class: right.input.checked ? "lead" : "" }, right.input.checked ? dots.repeat(40) : "") : "",
        numbers.input.checked ? h("span", {}, String(i * 2 + 1)) : "")));
      web.replaceChildren(...lines.map((l) => h("div", { class: links.input.checked ? "tofweb" : "" }, l)));
      right.input.disabled = !numbers.input.checked;
      leader.disabled = !numbers.input.checked || !right.input.checked;
    };
    for (const x of [numbers.input, right.input, links.input, include.input]) x.onchange = draw;
    leader.onchange = draw;
    label.onchange = draw;
    const body = h("div", { class: "tofdialog" },
      h("div", { class: "toftabs" }, h("span", { class: "on" }, "Abbildungsverzeichnis")),
      h("div", { class: "tofcols" },
        h("div", {}, h("div", {}, "Seitenansicht"), preview, numbers.el, right.el, h("label", { class: "field row" }, "Füllzeichen:", leader)),
        h("div", {}, h("div", {}, "Webvorschau"), web, links.el)),
      h("fieldset", {}, h("legend", {}, "Allgemein"),
        h("label", { class: "field row" }, "Formate:", formats),
        h("label", { class: "field row" }, "Beschriftung:", label), include.el));
    openModal("Abbildungsverzeichnis", body, [
      { label: "OK", primary: true, run: () => {
        lastLabel = label.value;
        e.insertTableOfFigures(label.value, {
          pageNumbers: numbers.input.checked, rightAlign: right.input.checked, links: links.input.checked,
          leader: leader.value as "dot" | "hyphen" | "underscore" | "none",
        }, include.input.checked);
        api.refresh();
        api.focus();
        return true;
      } },
      { label: "Abbrechen", run: () => true },
    ], { width: 720, cls: "tofdlg" });
    draw();
  };

  c.implement("ref.caption", { run: captionDialog, enabled: () => !!ed() });
  c.implement("ref.tableOfFigures", { run: tofDialog, enabled: () => !!ed() });
  c.implement("ref.updateTableOfFigures", { run: () => { ed()?.updateTablesOfFigures(); api.refresh(); }, enabled: () => !!ed()?.hasTableOfFigures() });
  return { captionDialog, updateFields: () => { ed()?.updateAllFields(); api.refresh(); } };
}
