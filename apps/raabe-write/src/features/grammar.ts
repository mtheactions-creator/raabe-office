// Grammatikprüfung while typing (Word's blue double underline) with RaabeWrite's own checker
// (Rust, D-040, D-051): paragraphs are checked in the background, a moment after typing stops;
// results are cached per paragraph text. The context menu offers the suggestions and Einmal
// ignorieren; F7's Editor pane lists the findings with their explanation. German, English, French
// and Spanish, each switched on once it measured well against Word (D-040, D-051).

import type { Editor, Para, TextMark } from "@raabe/text-engine";
import type { GrammarIssue, MenuEntry } from "@raabe/shell";
import type { AppApi } from "./api";
import type { Issue, IssueSource } from "./editor-pane";

/** The proofing languages the checker has rules for. */
const LANGS = new Set(["de", "en", "fr", "es"]);

/** Word's names for the kinds of findings (LanguageTool's categories). */
const TITLES: Record<string, string> = {
  GRAMMAR: "Grammatik", PUNCTUATION: "Zeichensetzung", HILFESTELLUNG_KOMMASETZUNG: "Zeichensetzung", CASING: "Groß-/Kleinschreibung",
  COMPOUNDING: "Getrennt- und Zusammenschreibung", CONFUSED_WORDS: "Verwechselte Wörter", TYPOS: "Tippfehler", TYPOGRAPHY: "Typografie",
  IDIOMS: "Redewendungen", SEMANTICS: "Bedeutung", CORRESPONDENCE: "Briefe", PROPER_NOUNS: "Eigennamen", MULTITOKEN_SPELLING: "Rechtschreibung",
};

/** The kind of finding (the context menu's note); Word heads every grammar card "Grammatik". */
export const grammarKind = (category: string) => TITLES[category] ?? "Grammatik";

interface Spell {
  langAt(e: Editor, p: Para, o: number): { lang: string; noProof: boolean };
  misspelledIn(e: Editor, para: Para, a: number, b: number): boolean;
  options(): { grammarAsYouType: boolean };
}

export function grammarFeature(api: AppApi, spell: Spell) {
  /** Findings per paragraph text and language ("de\u0001text"). */
  const cache = new Map<string, GrammarIssue[]>();
  const pending = new Map<string, Set<string>>();
  const ignoredOnce = new Set<string>();
  const ignoredRules = new Set<string>();
  let timer = 0;
  let generation = 0;
  const keyOf = (lang: string, text: string) => `${lang}\u0001${text}`;

  const proofLang = (e: Editor, para: Para): string | null => {
    if (!para.text || para.block) return null;
    const at = spell.langAt(e, para, 0);
    const l = at.lang.slice(0, 2).toLowerCase();
    return !at.noProof && LANGS.has(l) ? at.lang : null;
  };

  const flush = async () => {
    timer = 0;
    const jobs = [...pending.entries()];
    pending.clear();
    let any = false;
    // Versions typed in the meantime are no longer in the document.
    const live = new Set(api.editor()?.doc.paras.map((p) => p.text) ?? []);
    await Promise.all(jobs.map(async ([lang, set]) => {
      const texts = [...set].filter((t) => live.has(t));
      if (!texts.length) return;
      const res = await api.platform.grammarCheck(lang, texts).catch(() => texts.map(() => [] as GrammarIssue[]));
      texts.forEach((t, i) => cache.set(keyOf(lang, t), res[i] ?? []));
      if (res.some((r) => r.length)) any = true;
    }));
    if (cache.size > 4000) for (const k of [...cache.keys()].slice(0, 1000)) cache.delete(k);
    if (any) { generation++; api.editor()?.view.refresh(); }
  };

  const ask = (lang: string, text: string, now = false) => {
    const set = pending.get(lang) ?? new Set<string>();
    // Repaints ask again for what is already waiting: only new text restarts the pause.
    if (set.has(text) && !now) return;
    set.add(text);
    pending.set(lang, set);
    if (timer) clearTimeout(timer);
    timer = window.setTimeout(() => void flush(), now ? 0 : 700);
  };

  const issueKey = (text: string, f: GrammarIssue) => `${text}\u0001${f.start}\u0001${f.rule}`;
  const visible = (text: string, list: GrammarIssue[]) => {
    // One finding per place (the first rule's), not ignored ones.
    const out: GrammarIssue[] = [];
    for (const f of list) {
      if (ignoredRules.has(f.rule) || ignoredOnce.has(issueKey(text, f))) continue;
      if (out.some((g) => g.start < f.end && f.start < g.end)) continue;
      out.push(f);
    }
    return out;
  };

  /** The findings of a paragraph, when known (undefined: still being checked). */
  const findings = (e: Editor, para: Para): GrammarIssue[] | undefined => {
    const lang = proofLang(e, para);
    if (!lang) return [];
    const hit = cache.get(keyOf(lang, para.text));
    if (hit) return visible(para.text, hit).filter((f) => !spell.misspelledIn(e, para, f.start, f.end));
    ask(lang, para.text);
    return undefined;
  };

  const marksFor = (e: Editor, para: Para): readonly TextMark[] => {
    if (!spell.options().grammarAsYouType) return [];
    return (findings(e, para) ?? []).map((f) => ({ start: f.start, end: f.end, kind: "grammar" as const }));
  };

  /** Adds grammar marks to an editor's spelling marks. */
  const attach = (e: Editor) => {
    const prev = e.extraMarks;
    e.extraMarks = (story, para) => {
      const a = prev?.(story, para);
      const b = marksFor(e, para);
      return b.length ? [...(a ?? []), ...b] : a;
    };
  };

  // ----- context menu -----------------------------------------------------------------------------
  const entries = (e: Editor, pos: { p: number; o: number }): MenuEntry[] => {
    const para = e.doc.paras[pos.p];
    if (!para || !spell.options().grammarAsYouType) return [];
    const f = (findings(e, para) ?? []).find((x) => pos.o >= x.start && pos.o < x.end);
    if (!f) return [];
    const out: MenuEntry[] = f.suggestions.slice(0, 5).map((s) => ({ label: s || "(Löschen)", bold: true, run: () => { e.replaceRange(pos.p, f.start, f.end, s); api.focus(); } }));
    out.push({ label: f.short || grammarKind(f.category), run: () => {}, disabled: true },
      { label: "Einmal ignorieren", run: () => { ignoredOnce.add(issueKey(para.text, f)); generation++; e.view.refresh(); } }, "-");
    return out;
  };

  // ----- for the Editor pane (F7) --------------------------------------------------------------------
  const source: IssueSource = {
    prepare: async (e) => {
      for (const para of e.doc.paras) {
        const lang = proofLang(e, para);
        if (lang && !cache.has(keyOf(lang, para.text))) ask(lang, para.text, true);
      }
      if (timer) { clearTimeout(timer); await flush(); }
    },
    issues: (e) => {
      const out: Issue[] = [];
      e.doc.paras.forEach((para, p) => {
        const lang = proofLang(e, para);
        if (!lang) return;
        for (const f of visible(para.text, cache.get(keyOf(lang, para.text)) ?? [])) {
          if (spell.misspelledIn(e, para, f.start, f.end)) continue;
          out.push({ p, a: f.start, b: f.end, kind: "grammar", text: para.text.slice(f.start, f.end), lang, title: "Grammatik",
            message: f.message, suggestions: f.suggestions, rule: f.rule });
        }
      });
      return out;
    },
    suggestions: async (i) => i.suggestions ?? [],
    ignoreOnce: (e, i) => {
      const para = e.doc.paras[i.p];
      const f = para && (cache.get(keyOf(i.lang, para.text)) ?? []).find((x) => x.start === i.a && x.rule === i.rule);
      if (f) ignoredOnce.add(issueKey(para.text, f));
      generation++;
      e.view.refresh();
    },
    ignoreAll: (e, i) => { if (i.rule) ignoredRules.add(i.rule); generation++; e.view.refresh(); },
  };

  return { attach, entries, source, generation: () => generation };
}
