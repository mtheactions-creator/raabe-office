// Überprüfen > Rechtschreibung und Grammatik (F7): Word 365's pane "Editor" (recorded
// 2026-10-14, probe_word_languages.py): "1 verbleibend", a card with the kind of issue
// (Rechtschreibung: Nicht im Wörterbuch), the sentence with the word in red, Betrachten: the
// suggestions, each with synonyms "ähnlich wie" from the thesaurus, Weitere Optionen (Alle
// ändern), then Einmal ignorieren, Alle ignorieren, Zum Wörterbuch hinzufügen and Optionen.
// Grammar issues (grammar.ts) come into the same pane.

import { h, icon, messageBox, showMenu } from "@raabe/shell";
import type { Editor } from "@raabe/text-engine";
import type { AppApi } from "./api";
import type { Dock } from "./dock";
import { plainTerm } from "./thesaurus";

/** One issue of the document: a misspelled or repeated word, or a grammar finding. */
export interface Issue {
  p: number;
  a: number;
  b: number;
  kind: "spell" | "repeat" | "grammar";
  text: string;
  lang: string;
  /** Grammar: the rule's heading and explanation. */
  title?: string;
  message?: string;
  suggestions?: string[];
  /** Grammar: an id to ignore the rule's finding. */
  rule?: string;
}

export interface IssueSource {
  /** Checks the whole document (waits until every word is known). */
  prepare(e: Editor): Promise<void>;
  issues(e: Editor): Issue[];
  suggestions(i: Issue): Promise<string[]>;
  ignoreOnce(e: Editor, i: Issue): void;
  ignoreAll(e: Editor, i: Issue): void;
  /** Zum Wörterbuch hinzufügen (spelling only). */
  add?(i: Issue): void;
  /** Applies a suggestion (Wiederholtes Wort löschen …). */
  apply?(e: Editor, i: Issue, s: string): boolean;
}

export function editorPaneFeature(api: AppApi, docArea: HTMLElement, dock: Dock, sources: IssueSource[], openOptions: () => void) {
  const ed = () => api.editor();
  const c = api.commands;
  const pane = h("div", { class: "tpane edpane", hidden: true, role: "complementary", "aria-label": "Editor" });
  docArea.append(pane);
  dock.add(pane, "right", 340);
  let open = false;
  let at = { p: 0, o: 0 };
  let token = 0;

  const close = () => { open = false; pane.hidden = true; api.refresh(); api.focus(); };

  const all = (e: Editor) => sources.flatMap((s) => s.issues(e).map((i) => ({ i, s })))
    .sort((x, y) => x.i.p - y.i.p || x.i.a - y.i.a);

  /** The next issue at or after the position (from the start again at the end). */
  const next = (e: Editor) => {
    const list = all(e);
    const after = list.find(({ i }) => i.p > at.p || (i.p === at.p && i.a >= at.o));
    return { cur: after ?? list[0], count: list.length };
  };

  const sentence = (e: Editor, i: Issue): HTMLElement => {
    const t = e.doc.paras[i.p]?.text ?? "";
    let a = i.a, b = i.b;
    while (a > 0 && !/[.!?]\s/.test(t.slice(a - 2, a))) a--;
    while (b < t.length && !/[.!?]/.test(t[b - 1] ?? "")) b++;
    return h("div", { class: "ed-context" }, t.slice(a, i.a), h("b", { class: "ed-word" }, t.slice(i.a, i.b)), t.slice(i.b, b));
  };

  const show = async () => {
    const e = ed();
    if (!e || !open) return;
    const my = ++token;
    const { cur, count } = next(e);
    if (!cur) {
      close();
      await messageBox("RaabeWrite", "Die Rechtschreib- und Grammatikprüfung ist abgeschlossen.");
      return;
    }
    const { i, s } = cur;
    e.selectRange(i.p, i.a, i.b);
    const sugg = await s.suggestions(i);
    // Synonyms of each suggestion ("ähnlich wie"), as Word shows them.
    const similar = await Promise.all(sugg.slice(0, 3).map(async (x) => {
      if (i.kind !== "spell") return "";
      const r = await api.platform.thesaurus(i.lang, x).catch(() => null);
      const words = (r?.meanings[0]?.terms ?? []).filter((t) => !t.startsWith("!")).map(plainTerm).filter((t) => t.toLocaleLowerCase() !== x.toLocaleLowerCase() && !t.includes(" "));
      return words.slice(0, 3).join(", ");
    }));
    if (my !== token || !open) return;
    const advance = (to: { p: number; o: number }) => { at = to; void show(); };
    const apply = (x: string, everywhere = false) => {
      if (s.apply?.(e, i, x)) { advance({ p: i.p, o: i.a }); return; }
      if (everywhere) e.replaceAll(i.text, x, { matchCase: true, wholeWord: true });
      else e.replaceRange(i.p, i.a, i.b, x);
      advance({ p: i.p, o: i.a + x.length });
    };
    const head = h("div", { class: "tp-head" }, h("span", null, "Editor"),
      h("div", null,
        h("button", { "aria-label": "Optionen für den Aufgabenbereich", "data-tip": "Optionen für den Aufgabenbereich",
          onclick: (ev: MouseEvent) => showMenu(c, [{ label: "Schließen", run: close }], (ev.currentTarget as HTMLElement).getBoundingClientRect()) }, icon("chevron-down")),
        h("button", { "aria-label": "Bereich schließen", "data-tip": "Bereich schließen", onclick: close }, icon("close"))));
    const bar = h("div", { class: "ed-bar" }, h("button", { class: "ed-back", "aria-label": "Zurück", "data-tip": "Zurück", onclick: close }, icon("close")),
      h("span", null, `${count} verbleibend`));
    const title = i.kind === "grammar" ? i.title ?? "Grammatik" : "Rechtschreibung";
    const sub = i.kind === "repeat" ? "Wiederholtes Wort" : i.kind === "grammar" ? i.message ?? "" : "Nicht im Wörterbuch";
    const cards = h("div", { class: "ed-sugg" });
    if (!sugg.length) cards.append(h("div", { class: "ed-none" }, "Keine Vorschläge"));
    sugg.slice(0, 5).forEach((x, k) => {
      const more = h("button", { class: "ed-more", "aria-label": `Weitere Optionen für ${x}`, "data-tip": `Weitere Optionen für ${x}` }, icon("chevron-down"));
      more.onclick = (ev: MouseEvent) => {
        ev.stopPropagation();
        showMenu(c, [{ label: "Alle ändern", run: () => apply(x, true), disabled: i.kind !== "spell" }], (ev.currentTarget as HTMLElement).getBoundingClientRect());
      };
      const card = h("button", { class: "ed-card" }, h("span", { class: "ed-s" }, x || "(löschen)"), similar[k] ? h("span", { class: "ed-sim" }, similar[k]) : null);
      card.onclick = () => apply(x);
      cards.append(h("div", { class: "ed-cardrow" }, card, more));
    });
    const action = (ic: string, label: string, run: () => void, disabled = false) =>
      h("button", { class: "ed-action", disabled, onclick: run }, icon(ic), h("span", null, label));
    const actions = h("div", { class: "ed-actions" },
      action("ignore", "Einmal ignorieren", () => { s.ignoreOnce(e, i); advance({ p: i.p, o: i.b }); }),
      action("ignore-all", "Alle ignorieren", () => { s.ignoreAll(e, i); advance({ p: i.p, o: i.b }); }),
      i.kind === "spell" && s.add ? action("dictionary-add", "Zum Wörterbuch hinzufügen", () => { s.add!(i); advance({ p: i.p, o: i.b }); }) : null);
    const card = h("div", { class: "ed-issue" },
      h("div", { class: "ed-title" }, title), h("div", { class: "ed-sub" }, sub),
      h("div", { class: "ed-ctxrow" }, sentence(e, i), h("button", { class: "ed-read", disabled: true, "aria-label": "Kontext lesen", "data-tip": "Kontext lesen\nZurückgestellt" }, icon("read-aloud"))),
      h("div", { class: "ed-label" }, "Betrachten:"), cards, actions);
    const optionsLink = h("button", { class: "ed-options", onclick: () => openOptions() }, icon("editor"), h("span", null, "Optionen"));
    pane.replaceChildren(head, h("div", { class: "ed-body" }, bar, card, optionsLink));
    api.refresh();
  };

  const run = async () => {
    const e = ed();
    if (!e) return;
    open = true;
    pane.hidden = false;
    pane.replaceChildren(h("div", { class: "tp-head" }, h("span", null, "Editor")), h("div", { class: "ed-body muted" }, "Das Dokument wird überprüft..."));
    const [from] = e.selectionRange();
    at = { p: from.p, o: from.o };
    for (const s of sources) await s.prepare(e);
    await show();
  };

  c.implement("review.spelling", { run: () => void run(), enabled: () => !!ed(), checked: () => open });
  return { isOpen: () => open, refresh: () => { if (open) void show(); } };
}
