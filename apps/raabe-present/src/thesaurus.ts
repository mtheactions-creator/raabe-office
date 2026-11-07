// Überprüfen > Thesaurus (Umschalt+F7, D-072): RaabeWrite's open thesauri (German, English, French, Spanish;
// D-049) looked up by the platform. A box with the word at the caret (or the selected word), its meanings and
// their synonyms; a click on a synonym puts it in place of the word, another word can be looked up.

import { h, showDialog } from "@raabe/shell";
import type { PresentApp } from "./app";
import * as T from "./textops";

export async function thesaurusDialog(app: PresentApp): Promise<void> {
  const ed = app.editor, t = ed.text;
  const body = t ? ed.textBody(t.id) : undefined;
  let range: { p: number; from: number; to: number } | null = null;
  let word = "";
  if (t && body) {
    const p = T.paras(body)[t.caret.p];
    if (T.cmpPos(t.caret, t.anchor) !== 0 && t.caret.p === t.anchor.p) {
      const from = Math.min(t.caret.o, t.anchor.o), to = Math.max(t.caret.o, t.anchor.o);
      range = { p: t.caret.p, from, to };
    } else if (p) { const w = T.wordAt(p, t.caret.o); if (w.to > w.from) range = { p: t.caret.p, ...w }; }
    if (range && p) word = T.paraText(p).slice(range.from, range.to).trim();
  }
  const lang = (body?.getElementsByTagNameNS("http://schemas.openxmlformats.org/drawingml/2006/main", "rPr")[0]?.getAttribute("lang")) ?? "de-DE";
  const input = h("input", { type: "text", value: word, style: "width:260px" }) as HTMLInputElement;
  const list = h("div", { class: "thes-list", style: "max-height:320px;overflow:auto;min-width:340px;margin-top:8px" });
  const look = async () => {
    list.replaceChildren(h("p", null, "…"));
    const r = await app.platform.thesaurus(lang, input.value.trim()).catch(() => null);
    list.replaceChildren();
    if (!r || !r.meanings.length) { list.append(h("p", null, "Keine Ergebnisse gefunden.")); return; }
    for (const m of r.meanings) {
      list.append(h("div", { class: "thes-head", style: "font-weight:600;margin-top:6px" }, `${m.terms[0]?.replace(/^!/, "") ?? r.word}${m.pos ? ` (${m.pos})` : ""}`));
      for (const term of m.terms) {
        const anti = term.startsWith("!"), text = term.replace(/^!/, "").replace(/\s*\([^)]*\)/g, "").trim();
        list.append(h("button", { class: "menu-item", style: "display:block;width:100%;text-align:left;padding-left:16px", ondblclick: () => { input.value = text; void look(); },
          onclick: () => {
            if (!range || !body || !t) return;
            const rg = range;
            ed.edit("Thesaurus", () => { const pos = T.deleteRange(body, { p: rg.p, o: rg.from }, { p: rg.p, o: rg.to }); const end = T.insertText(body, pos, text); t.caret = end; t.anchor = end; });
            (list.closest(".dialog")?.querySelector(".dlg-buttons button.primary") as HTMLElement | null)?.click();
          } }, `${text}${anti ? " (Antonym)" : ""}`));
      }
    }
  };
  input.addEventListener("keydown", (e) => { if (e.key === "Enter") { e.preventDefault(); void look(); } });
  void look();
  await showDialog({ title: "Thesaurus", body: h("div", { class: "form" }, h("label", null, "Suchen nach: ", input), list),
    buttons: [{ label: "Schließen", value: false, primary: true }], cancel: false });
}
