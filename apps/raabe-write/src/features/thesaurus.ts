// Überprüfen > Thesaurus (Umschalt+F7) and Synonyme in the context menu, as Word shows them
// (recorded 2026-10-14, probe_word_languages.py): the pane "Thesaurus" with Zurück, the search
// box and the meanings, each a heading "wundervoll (Adj.)" with its synonyms and the antonyms
// marked "(Antonym)" at the end; a click on a word looks it up, its arrow offers Einfügen and
// Kopieren; the language below. The context menu lists five synonyms of the first meaning and
// Thesaurus... The data is open (OpenThesaurus, WordNet, D-049), looked up by the platform.

import { h, icon, showMenu, type MenuEntry, type ThesaurusResult } from "@raabe/shell";
import type { Editor, Pos } from "@raabe/text-engine";
import type { AppApi } from "./api";
import type { Dock } from "./dock";
import { languageName } from "./languages";

/** The proofing languages the thesaurus has, as the pane's language list offers them. */
const THES_LANGS = ["de-DE", "en-US", "en-GB", "fr-FR", "es-ES"];

/** A term as written in the text: without notes such as "(ugs.)" or optional parts "(ein)". */
export const plainTerm = (t: string) => t.replace(/^!/, "").replace(/\s*\([^)]*\)/g, "").replace(/\s+/g, " ").trim();

/** The word at a position (or the selected word) with its range. */
export function wordAt(e: Editor, pos: Pos): { a: number; b: number; text: string } | null {
  const para = e.doc.paras[pos.p];
  if (!para) return null;
  const sel = e.hasSelection ? e.selectedText() : "";
  if (sel && /^[\p{L}\p{M}'’-]+$/u.test(sel.trim())) {
    const [from] = e.selectionRange();
    return { a: from.o, b: from.o + sel.length, text: sel.trim() };
  }
  const re = /[\p{L}\p{M}]+(?:['’-][\p{L}\p{M}]+)*/gu;
  for (const m of para.text.matchAll(re)) {
    const a = m.index!, b = a + m[0].length;
    if (pos.o >= a && pos.o <= b) return { a, b, text: m[0] };
  }
  return null;
}

/** Keeps the capital at the start of the replaced word (Schön → Hübsch). */
const matchCase = (original: string, term: string) =>
  /^\p{Lu}/u.test(original) && !/^\p{Lu}/u.test(term) ? term[0].toLocaleUpperCase("de-DE") + term.slice(1) : term;

interface Group { heading: string; items: { term: string; antonym: boolean }[] }

/** Word's groups: the meaning's first other word and its part of speech, then its terms. */
function groups(r: ThesaurusResult): Group[] {
  const self = r.word.toLocaleLowerCase();
  return r.meanings.map((m) => {
    const terms = m.terms.filter((t) => !t.startsWith("!"));
    const others = terms.filter((t) => plainTerm(t).toLocaleLowerCase() !== self);
    // Single words first, as Word's lists are; the heading is the first of them.
    const ordered = [...others.filter((t) => !plainTerm(t).includes(" ")), ...others.filter((t) => plainTerm(t).includes(" "))];
    const lead = ordered.find((t) => /\(Hauptform\)/.test(t)) ?? ordered[0] ?? terms[0] ?? "";
    const heading = `${plainTerm(lead)}${m.pos ? ` (${m.pos})` : ""}`;
    const items = [
      ...ordered.map((t) => ({ term: t.replace(/\s*\(Hauptform\)/, ""), antonym: false })),
      ...m.terms.filter((t) => t.startsWith("!")).map((t) => ({ term: t.slice(1), antonym: true })),
    ];
    return { heading, items };
  }).filter((g) => g.items.length);
}

export function thesaurusFeature(api: AppApi, docArea: HTMLElement, dock: Dock) {
  const ed = () => api.editor();
  const c = api.commands;
  const langAt = (e: Editor) => e.proofingAt().lang;

  // ----- Synonyme (context menu) ---------------------------------------------------------------
  const synonymEntries = async (e: Editor, pos: Pos): Promise<MenuEntry[]> => {
    const w = wordAt(e, pos);
    const lang = langAt(e);
    const r = w ? await api.platform.thesaurus(lang, w.text).catch(() => null) : null;
    const out: MenuEntry[] = [];
    if (w && r) {
      const self = r.word.toLocaleLowerCase();
      const pick: string[] = [];
      for (const m of r.meanings) {
        const terms = m.terms.filter((t) => !t.startsWith("!") && plainTerm(t).toLocaleLowerCase() !== self).map(plainTerm);
        // Single words first, as Word's lists are.
        for (const t of [...terms.filter((t) => !t.includes(" ")), ...terms.filter((t) => t.includes(" "))]) if (!pick.includes(t) && pick.length < 5) pick.push(t);
        if (pick.length) break;
      }
      for (const t of pick) out.push({ label: t, run: () => { e.replaceRange(pos.p, w.a, w.b, matchCase(w.text, t)); api.focus(); } });
    }
    if (!out.length) out.push({ label: "(Keine Vorschläge)", run: () => {}, disabled: true });
    out.push("-", { label: "Thesaurus...", run: () => openPane(w?.text ?? "", lang) });
    return out;
  };

  // ----- the pane --------------------------------------------------------------------------------
  const pane = h("div", { class: "tpane thespane", hidden: true, role: "complementary", "aria-label": "Thesaurus" });
  docArea.append(pane);
  dock.add(pane, "right", 220);
  let open = false;
  let lang = "de-DE";
  let history: string[] = [];
  let current = "";
  let result: ThesaurusResult | null = null;
  let token = 0;

  const close = () => { open = false; pane.hidden = true; api.refresh(); api.focus(); };

  const lookup = async (word: string, push = true) => {
    if (push && current && current !== word) history.push(current);
    current = word;
    const my = ++token;
    result = word.trim() ? await api.platform.thesaurus(lang, word.trim()).catch(() => null) : null;
    if (my === token) render();
  };

  const insert = (term: string) => {
    const e = ed();
    if (!e) return;
    const w = wordAt(e, e.sel.head);
    const t = plainTerm(term);
    if (e.hasSelection || !w) e.insertText(t);
    else e.replaceRange(e.sel.head.p, w.a, w.b, matchCase(w.text, t));
    api.focus();
  };

  const render = () => {
    if (!open) return;
    const head = h("div", { class: "tp-head" }, h("span", null, "Thesaurus"),
      h("div", null,
        h("button", { "aria-label": "Optionen für den Aufgabenbereich", "data-tip": "Optionen für den Aufgabenbereich",
          onclick: (ev: MouseEvent) => showMenu(c, [{ label: "Schließen", run: close }], (ev.currentTarget as HTMLElement).getBoundingClientRect()) }, icon("chevron-down")),
        h("button", { "aria-label": "Bereich schließen", "data-tip": "Bereich schließen", onclick: close }, icon("close"))));
    const input = h("input", { type: "text", class: "thes-input", value: current, "aria-label": "Ausgangswort", spellcheck: false }) as HTMLInputElement;
    input.addEventListener("keydown", (ev) => { ev.stopPropagation(); if (ev.key === "Enter") { ev.preventDefault(); void lookup(input.value); } });
    const back = h("button", { class: "thes-back", "aria-label": "Zurück", "data-tip": "Zurück", disabled: !history.length,
      onclick: () => { const prev = history.pop(); if (prev !== undefined) void lookup(prev, false); } }, icon("back"));
    const search = h("button", { class: "thes-search", "aria-label": "Suchen", "data-tip": "Suchen", onclick: () => void lookup(input.value) }, icon("search"));
    const list = h("div", { class: "thes-list", role: "tree", "aria-label": "Synonyme" });
    if (current && !result) list.append(h("div", { class: "thes-empty" }, "Keine Ergebnisse gefunden."));
    for (const g of result ? groups(result) : []) {
      const items = h("div", { class: "thes-items" });
      const toggle = h("button", { class: "thes-heading", "aria-expanded": "true" }, icon("chevron-down"), h("span", null, g.heading));
      toggle.onclick = () => { const on = items.hidden; items.hidden = !on; toggle.classList.toggle("closed", !on); toggle.setAttribute("aria-expanded", String(on)); };
      for (const it of g.items) {
        const label = `${it.term}${it.antonym ? " (Antonym)" : ""}`;
        const menu = h("button", { class: "thes-more", "aria-label": `Optionen für ${it.term}` }, icon("chevron-down"));
        menu.onclick = (ev: MouseEvent) => {
          ev.stopPropagation();
          showMenu(c, [{ label: "Einfügen", run: () => insert(it.term) }, { label: "Kopieren", run: () => void api.platform.clipboardWrite(plainTerm(it.term)) }],
            (ev.currentTarget as HTMLElement).getBoundingClientRect());
        };
        const row = h("div", { class: "thes-item", role: "treeitem", tabIndex: 0 }, h("span", null, label), menu);
        row.onclick = () => void lookup(plainTerm(it.term));
        items.append(row);
      }
      list.append(h("div", { class: "thes-group", role: "group", "aria-label": g.heading }, toggle, items));
    }
    const sel = h("select", { class: "thes-lang", "aria-label": "Sprache" }, ...THES_LANGS.map((l) => h("option", { value: l }, languageName(l)))) as HTMLSelectElement;
    sel.value = THES_LANGS.includes(lang) ? lang : THES_LANGS.find((l) => l.slice(0, 2) === lang.slice(0, 2)) ?? "de-DE";
    sel.onchange = () => { lang = sel.value; void lookup(current, false); };
    pane.replaceChildren(head, h("div", { class: "thes-bar" }, back, h("div", { class: "thes-box" }, input, search)), list, h("div", { class: "thes-foot" }, sel));
  };

  const openPane = (word: string, language: string) => {
    open = true;
    pane.hidden = false;
    lang = language;
    history = [];
    current = "";
    result = null;
    render();
    api.refresh();
    void lookup(word, false);
  };

  c.implement("review.thesaurus", {
    run: () => {
      const e = ed();
      if (!e) return;
      openPane(wordAt(e, e.sel.head)?.text ?? "", langAt(e));
    },
    enabled: () => !!ed(),
    checked: () => open,
  });

  return { synonymEntries, isOpen: () => open };
}
