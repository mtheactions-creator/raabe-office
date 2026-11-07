// Überprüfen > Sprache and Barrierefreiheit überprüfen (D-072).
// Sprache für die Korrekturhilfen festlegen: the language of the selected text, or of all text on the slide
// or in the presentation (PowerPoint's "Was aktualisieren?": Dokument / Folie), written as lang on the runs.
// Barrierefreiheit überprüfen: our own short check (pictures, charts and tables without Alternativtext,
// slides without a title); a click on an entry goes to the slide and selects the object.

import { h, showDialog } from "@raabe/shell";
import { PptxDom } from "@raabe/ooxml";
import type { PresentApp } from "./app";
import * as T from "./textops";

export const LANGUAGES: [string, string][] = [
  ["de-DE", "Deutsch (Deutschland)"], ["de-AT", "Deutsch (Österreich)"], ["de-CH", "Deutsch (Schweiz)"],
  ["en-GB", "Englisch (Vereinigtes Königreich)"], ["en-US", "Englisch (Vereinigte Staaten)"], ["fr-FR", "Französisch (Frankreich)"],
  ["es-ES", "Spanisch (Spanien)"], ["it-IT", "Italienisch (Italien)"], ["la", "Latein"], ["nl-NL", "Niederländisch (Niederlande)"],
  ["pl-PL", "Polnisch"], ["ru-RU", "Russisch"], ["tr-TR", "Türkisch"], ["el-GR", "Griechisch"], ["pt-PT", "Portugiesisch (Portugal)"],
  ["ar-SA", "Arabisch (Saudi-Arabien)"], ["zh-CN", "Chinesisch (Volksrepublik China)"], ["ja-JP", "Japanisch"],
];

function setLangIn(doc: Document, lang: string): void {
  for (const tag of ["rPr", "endParaRPr"]) for (const r of Array.from(doc.getElementsByTagNameNS(PptxDom.NS_A, tag))) r.setAttribute("lang", lang);
}

export async function languageDialog(app: PresentApp): Promise<void> {
  const ed = app.editor;
  const hasText = !!ed.text && T.cmpPos(ed.text.caret, ed.text.anchor) !== 0;
  const list = h("select", { size: "10", style: "width:320px" }) as HTMLSelectElement;
  const cur = ed.text ? (ed.textBody(ed.text.id)?.getElementsByTagNameNS(PptxDom.NS_A, "rPr")[0]?.getAttribute("lang") ?? "de-DE") : "de-DE";
  for (const [code, name] of LANGUAGES) { const o = h("option", { value: code }, name) as HTMLOptionElement; o.selected = code === cur; list.append(o); }
  const radio = (label: string, on: boolean) => { const r = h("input", { type: "radio", name: "scope" }) as HTMLInputElement; r.checked = on; return { r, el: h("label", { class: "dlg-check" }, r, ` ${label}`) }; };
  const docR = radio("Dokument", false), slideR = radio("Folie", true);
  const body = h("div", { class: "form" },
    ...(hasText ? [] : [h("b", null, "Was aktualisieren?"), docR.el, slideR.el]),
    h("label", null, hasText ? "Markierten Text markieren als:" : "Gesamten Text auf dieser Folie markieren als:"), list);
  const ok = await showDialog({ title: "Sprache", body, buttons: [{ label: "OK", value: true, primary: true }, { label: "Abbrechen", value: false }], cancel: false });
  if (!ok) return;
  const lang = list.value;
  if (hasText) { app.formatText("Sprache", (r) => r.setAttribute("lang", lang)); return; }
  const paths = docR.r.checked ? app.deck.slides.map((s) => s.path) : [app.slidePath()!];
  app.deck.change("Sprache", paths, () => { for (const p of paths) setLangIn(app.deck.pptx.doc(p), lang); });
  ed.render();
}

interface Finding { slide: number; id?: number; text: string }

function findings(app: PresentApp): Finding[] {
  const out: Finding[] = [];
  app.deck.slides.forEach((s, i) => {
    const v = app.deck.view(s.path);
    if (!v.shapes.some((x) => (x.ph?.type === "title" || x.ph?.type === "ctrTitle") && x.text && !x.text.empty)) out.push({ slide: i, text: `Folie ${i + 1}: kein Folientitel` });
    for (const x of v.shapes) {
      const descr = PptxDom.kids(x.el).find((c) => c.localName.startsWith("nv"))?.firstElementChild?.getAttribute("descr");
      const kind = x.kind === "pic" ? "Bild" : x.frame === "chart" ? "Diagramm" : x.frame === "table" ? "Tabelle" : x.frame === "diagram" ? "SmartArt" : null;
      if (kind && !descr) out.push({ slide: i, id: x.id, text: `Folie ${i + 1}: ${kind} „${x.name}“ ohne Alternativtext` });
    }
  });
  return out;
}

export async function accessibilityCheck(app: PresentApp): Promise<void> {
  const found = findings(app);
  const list = h("div", { class: "a11y-list", style: "max-height:320px;overflow:auto;min-width:380px" });
  if (!found.length) list.append(h("p", null, "Keine Probleme gefunden."));
  for (const f of found) list.append(h("button", { class: "menu-item", style: "display:block;width:100%;text-align:left", onclick: () => { app.setCurrent(f.slide); if (f.id !== undefined) app.editor.select([f.id]); } }, f.text));
  await showDialog({ title: "Barrierefreiheit", body: h("div", { class: "form" }, h("b", null, "Ergebnisse der Prüfung"), list),
    buttons: [{ label: "Schließen", value: true, primary: true }], cancel: true });
}
