// Rechtschreibprüfung while typing (red wavy lines), the context menu with suggestions, the
// language of the text (Überprüfen > Sprache, the status bar) and Word's options for proofing
// (Datei > Optionen > Dokumentprüfung, Sprache; recorded 2026-10-14). Words are checked in the
// background by the platform (Hunspell dictionaries for German, English, French and Spanish in
// Rust, D-031, D-049); results are cached per word and language. As in Word, words in
// CAPITALS, words with digits and Internet addresses are ignored, repeated words are marked,
// text marked "Rechtschreibung und Grammatik nicht prüfen" is left out, and the word being typed
// is not marked until the caret leaves it. Sprache automatisch erkennen (detect.ts) gives
// sentences in another proofing language that language for checking.

import type { Editor, Para, TextMark } from "@raabe/text-engine";
import { h, messageBox, showDialog, type MenuEntry } from "@raabe/shell";
import type { AppApi } from "./api";
import { detectSentences } from "./detect";
import { LANGUAGES, PREFERRED, languageName, proofLang, tagFor } from "./languages";

export { languageName };

export interface Word { a: number; b: number; text: string; lang: string; repeat?: boolean }

const USER_DICT = "raabe.userDictionary";
const OPTIONS = "raabe.write.proofing";

/** Datei > Optionen > Dokumentprüfung (Word's defaults) and Sprache automatisch erkennen. */
export interface ProofingOptions {
  ignoreCaps: boolean;
  ignoreNumbers: boolean;
  ignoreAddresses: boolean;
  flagRepeated: boolean;
  checkAsYouType: boolean;
  grammarAsYouType: boolean;
  detect: boolean;
  /** Bevorzugte Erstellungssprache: the language of new documents. */
  preferred: string;
}

const DEFAULTS: ProofingOptions = { ignoreCaps: true, ignoreNumbers: true, ignoreAddresses: true, flagRepeated: true, checkAsYouType: true, grammarAsYouType: true, detect: true, preferred: "de-DE" };

export function loadProofing(): ProofingOptions {
  try { return { ...DEFAULTS, ...JSON.parse(localStorage.getItem(OPTIONS) ?? "{}") }; } catch { return { ...DEFAULTS }; }
}

/** Damerau-Levenshtein distance (changes, insertions, deletions, swaps). */
function distance(a: string, b: string): number {
  const d: number[][] = Array.from({ length: a.length + 1 }, (_, i) => [i, ...new Array<number>(b.length).fill(0)]);
  for (let j = 1; j <= b.length; j++) d[0][j] = j;
  for (let i = 1; i <= a.length; i++) {
    for (let j = 1; j <= b.length; j++) {
      const cost = a[i - 1] === b[j - 1] ? 0 : 1;
      d[i][j] = Math.min(d[i - 1][j] + 1, d[i][j - 1] + 1, d[i - 1][j - 1] + cost);
      if (i > 1 && j > 1 && a[i - 1] === b[j - 2] && a[i - 2] === b[j - 1]) d[i][j] = Math.min(d[i][j], d[i - 2][j - 2] + 1);
    }
  }
  return d[a.length][b.length];
}

const WORD_RE = /[\p{L}\p{M}\u00ad]+(?:['’\u2011-][\p{L}\p{M}\u00ad]+)*/gu;

export function spellFeature(api: AppApi) {
  const cache = new Map<string, Map<string, boolean>>();
  const pending = new Map<string, Set<string>>();
  const ignored = new Set<string>();
  /** Einmal ignorieren: the places of single words left alone (paragraph object and offset). */
  const ignoredOnce = new WeakMap<Para, Set<number>>();
  let userDict = new Set<string>();
  try { userDict = new Set(JSON.parse(localStorage.getItem(USER_DICT) ?? "[]")); } catch { /* none yet */ }
  let options = loadProofing();
  const marks = new WeakMap<Para, { key: string; marks: TextMark[]; complete: boolean }>();
  const detected = new WeakMap<Para, ReturnType<typeof detectSentences>>();
  let timer = 0;
  const editors = new WeakSet<Editor>();
  let generation = 0;

  const known = (lang: string, w: string): boolean | undefined => {
    if (ignored.has(w) || userDict.has(w)) return true;
    if (!proofLang(lang)) return true;
    return cache.get(lang)?.get(w);
  };

  const flush = async () => {
    timer = 0;
    const jobs = [...pending.entries()];
    pending.clear();
    let any = false;
    await Promise.all(jobs.map(async ([lang, set]) => {
      const list = [...set];
      const ok = await api.platform.spellCheck(lang, list).catch(() => list.map(() => true));
      const m = cache.get(lang) ?? new Map<string, boolean>();
      list.forEach((w, i) => m.set(w, ok[i] ?? true));
      cache.set(lang, m);
      if (ok.some((x) => !x)) any = true;
    }));
    if (any) api.editor()?.view.refresh();
  };
  const ask = (lang: string, w: string) => {
    const set = pending.get(lang) ?? new Set<string>();
    set.add(w);
    pending.set(lang, set);
    if (!timer) timer = window.setTimeout(() => void flush(), 150);
  };

  /** The detected language of each sentence of a paragraph (cached per paragraph object). */
  const sentences = (p: Para) => {
    let s = detected.get(p);
    if (!s) { s = detectSentences(p.text); detected.set(p, s); }
    return s;
  };

  /** The language of the text at an offset: its own, or the detected one for text in the default language. */
  const langAt = (e: Editor, p: Para, o: number): { lang: string; noProof: boolean } => {
    let pos = 0;
    let run = p.runs[0];
    for (const r of p.runs) { if (o < pos + r.len) { run = r; break; } pos += r.len; run = r; }
    const props = run?.props ?? p.mark;
    const rc = e.doc.styles.resolveChar(p.props.style, props);
    const own = rc.lang ?? "de-DE";
    if (!options.detect || props.lang !== undefined) return { lang: own, noProof: !!rc.noProof || !!rc.hidden };
    const d = sentences(p).find((s) => o >= s.a && o < s.b)?.lang;
    return { lang: d ? tagFor(d, own) : own, noProof: !!rc.noProof || !!rc.hidden };
  };

  /** The words of a paragraph to check, with their language. */
  const words = (e: Editor, p: Para): Word[] => {
    const out: Word[] = [];
    let prev: { text: string; b: number } | null = null;
    for (const m of p.text.matchAll(WORD_RE)) {
      const a = m.index!, b = a + m[0].length;
      const text = m[0].replace(/\u00ad/g, "").replace(/\u2011/g, "-");
      const { lang, noProof } = langAt(e, p, a);
      const before = prev;
      prev = { text, b };
      if (noProof || !proofLang(lang)) continue;
      // Wiederholte Wörter kennzeichnen: "der der" (only spaces between).
      if (options.flagRepeated && before && before.text.toLocaleLowerCase() === text.toLocaleLowerCase() && /^ +$/.test(p.text.slice(before.b, a)) && !/^\p{N}/u.test(text)) {
        out.push({ a, b, text, lang, repeat: true });
        continue;
      }
      // Internet and e-mail addresses, words next to digits (12a) and words in capitals are left out.
      let s0 = a, e0 = b;
      while (s0 > 0 && !/\s/.test(p.text[s0 - 1])) s0--;
      while (e0 < p.text.length && !/\s/.test(p.text[e0])) e0++;
      if (options.ignoreAddresses && /:\/\/|^www\.|@|\\/i.test(p.text.slice(s0, e0))) continue;
      if (options.ignoreNumbers && (/\p{N}/u.test(p.text[a - 1] ?? "") || /\p{N}/u.test(p.text[b] ?? ""))) continue;
      if (m[0].length < 2 || (options.ignoreCaps && m[0] === m[0].toLocaleUpperCase("de-DE"))) continue;
      out.push({ a, b, text, lang });
    }
    return out;
  };

  /** Whether a word is wrong (false), right (true) or not yet known (undefined); asks for unknown ones. */
  const judge = (w: Word, askMissing: boolean): boolean | undefined => {
    if (w.repeat) return false;
    const k = known(w.lang, w.text);
    if (k !== false) {
      if (k === undefined && askMissing) ask(w.lang, w.text);
      return k;
    }
    // Hyphenated words are right when their parts are (Kohlenstoff-Atom).
    const parts = w.text.split(/[-]/).filter((x) => x.length > 1 && x !== x.toLocaleUpperCase("de-DE"));
    if (parts.length < 2) return false;
    const each = parts.map((x) => known(w.lang, x));
    if (each.every((x) => x === true)) return true;
    if (each.some((x) => x === undefined)) {
      if (askMissing) parts.forEach((x, i) => { if (each[i] === undefined) ask(w.lang, x); });
      return undefined;
    }
    return false;
  };

  /** Spelling marks of a paragraph (cached until its text or the caret's word changes). */
  const marksFor = (e: Editor, story: string, para: Para): readonly TextMark[] | undefined => {
    if (!options.checkAsYouType || para.block || !para.text) return undefined;
    // The word with the caret is not judged while it is being typed.
    const h0 = e.sel.head;
    const caretPara = story === (e.doc.activeStory === "main" ? "main" : e.doc.activeStory) && e.doc.paras[h0.p] === para && !e.hasSelection ? h0.o : -1;
    const key = `${caretPara}|${generation}`;
    const hit = marks.get(para);
    if (hit && hit.complete && hit.key === key) return hit.marks;
    const out: TextMark[] = [];
    let complete = true;
    const once = ignoredOnce.get(para);
    for (const w of words(e, para)) {
      if (caretPara >= w.a && caretPara <= w.b) continue;
      if (once?.has(w.a)) continue;
      const k = judge(w, true);
      if (k === false) out.push({ start: w.a, end: w.b, kind: "spell" });
      else if (k === undefined) complete = false;
    }
    marks.set(para, { key, marks: out, complete });
    return out;
  };

  const refresh = () => { generation++; api.editor()?.view.refresh(); };

  /** Connects an editor (each document's). */
  const attach = (e: Editor) => {
    if (editors.has(e)) return;
    editors.add(e);
    e.extraMarks = (story, para) => marksFor(e, story, para);
  };

  const misspelled = (e: Editor, p: number, a: number): Word | null => {
    const para = e.doc.paras[p];
    for (const w of words(e, para)) if (a >= w.a && a < w.b && judge(w, false) === false && !ignoredOnce.get(para)?.has(w.a)) return w;
    return null;
  };

  const addToDictionary = (w: string) => {
    userDict.add(w);
    try { localStorage.setItem(USER_DICT, JSON.stringify([...userDict])); } catch { /* best effort */ }
    refresh();
  };
  const ignoreAll = (w: string) => { ignored.add(w); refresh(); };
  const ignoreOnce = (e: Editor, p: number, a: number) => {
    const para = e.doc.paras[p];
    const set = ignoredOnce.get(para) ?? new Set<number>();
    set.add(a);
    ignoredOnce.set(para, set);
    refresh();
  };
  /** Wiederholtes Wort löschen: the word and the space before it. */
  const deleteRepeat = (e: Editor, p: number, w: Word) => {
    let a = w.a;
    while (a > 0 && e.doc.paras[p].text[a - 1] === " ") a--;
    e.replaceRange(p, a, w.b, "");
  };

  /**
   * Suggestions, most likely first: Hunspell's list re-ranked as Word's reads (the same first
   * letter, then the fewest changes; "Feher" → Fehler, Feier … before Eher).
   */
  const suggest = async (w: Word) => {
    const list = await api.platform.spellSuggest(w.lang, w.text).catch(() => [] as string[]);
    const score = (s: string) => (s[0] === w.text[0] ? 0 : s[0]?.toLocaleLowerCase() === w.text[0]?.toLocaleLowerCase() ? 1 : 3) + distance(s, w.text);
    return list.map((s, i) => ({ s, i, k: score(s) })).sort((a, b) => a.k - b.k || a.i - b.i).map((x) => x.s);
  };

  // ----- context menu: suggestions for a misspelled word -------------------------------------
  const spellEntries = async (e: Editor, pos: { p: number; o: number }): Promise<MenuEntry[]> => {
    const w = misspelled(e, pos.p, pos.o);
    if (!w) return [];
    if (w.repeat) {
      return [{ label: "Wiederholtes Wort löschen", bold: true, run: () => { deleteRepeat(e, pos.p, w); api.focus(); } },
        { label: "Ignorieren", run: () => ignoreOnce(e, pos.p, w.a) }, "-"];
    }
    const sugg = await suggest(w);
    const out: MenuEntry[] = sugg.slice(0, 5).map((s) => ({ label: s, bold: true, run: () => { e.replaceRange(pos.p, w.a, w.b, s); api.focus(); } }));
    if (!sugg.length) out.push({ label: "(Keine Vorschläge)", run: () => {}, disabled: true });
    out.push("-", { label: "Alle ignorieren", run: () => ignoreAll(w.text) },
      { label: "Zum Wörterbuch hinzufügen", run: () => addToDictionary(w.text) }, "-");
    return out;
  };

  // ----- for the Editor pane (F7) ----------------------------------------------------------------
  /** Checks every word of the document (waits for the dictionary). */
  const checkAll = async (e: Editor) => {
    for (const p of e.doc.paras) if (!p.block) for (const w of words(e, p)) judge(w, true);
    clearTimeout(timer);
    await flush();
  };

  /** The spelling issues from a position on (in document order). */
  const issues = (e: Editor): { p: number; w: Word }[] => {
    const out: { p: number; w: Word }[] = [];
    e.doc.paras.forEach((para, p) => {
      if (para.block) return;
      const once = ignoredOnce.get(para);
      for (const w of words(e, para)) if (!once?.has(w.a) && judge(w, false) === false) out.push({ p, w });
    });
    return out;
  };

  // ----- Sprache (Korrekturhilfesprache) ------------------------------------------------------
  const languageDialog = async () => {
    const e = api.editor();
    if (!e) return;
    const cur = e.proofingAt();
    const shown = (() => {
      const p = e.doc.paras[e.sel.head.p];
      return p ? langAt(e, p, Math.max(0, Math.min(e.sel.head.o, p.text.length - 1))).lang : cur.lang;
    })();
    const radio = (checked: boolean) => h("input", { type: "radio", name: "langscope", checked }) as HTMLInputElement;
    const scopeSel = radio(true), scopeDoc = radio(false);
    // Word's list: the editing languages first, then every language (proofing ones marked).
    const item = (tag: string, name: string) => h("option", { value: tag, class: proofLang(tag) ? "proof" : "" }, `${proofLang(tag) ? "\u2713 " : "\u2003 "}${name}`);
    const sel = h("select", { size: 8, class: "langlist" },
      ...PREFERRED.map((t) => item(t, languageName(t))),
      h("option", { disabled: true, class: "sep" }, "\u2500".repeat(30)),
      ...LANGUAGES.map(([t, n]) => item(t, n))) as HTMLSelectElement;
    sel.value = shown;
    if (sel.selectedIndex < 0) sel.value = cur.lang;
    const noProof = h("input", { type: "checkbox", checked: cur.noProof }) as HTMLInputElement;
    const detect = h("input", { type: "checkbox", checked: options.detect }) as HTMLInputElement;
    const result = await showDialog<"ok" | "default" | false>({
      title: "Sprache",
      body: h("div", { class: "langdlg" },
        h("div", null, "Sprache für die Korrekturhilfen ändern für:"),
        h("div", { class: "row" }, h("label", { class: "check" }, scopeSel, "Markierter Text"), h("label", { class: "check" }, scopeDoc, "Aktives Dokument")),
        h("div", null, "Ausgewählten Text markieren als:"), sel,
        h("div", { class: "muted" }, "Alle Korrekturhilfen verwenden automatisch die Wörterbücher der ausgewählten Sprache."),
        h("label", { class: "check" }, noProof, "Rechtschreibung und Grammatik nicht prüfen"),
        h("label", { class: "check" }, detect, "Sprache automatisch erkennen")),
      buttons: [{ label: "Als Standard festlegen", value: "default", left: true }, { label: "OK", value: "ok", primary: true }, { label: "Abbrechen", value: false }],
      cancel: false,
    });
    if (!result) { api.focus(); return; }
    const tag = sel.value || cur.lang;
    if (result === "default") {
      const yes = await showDialog({
        title: "RaabeWrite",
        body: `Möchten Sie die Standardsprache in ${languageName(tag)} ändern?\nDiese Änderung wirkt sich auf alle neuen Dokumente aus.`,
        buttons: [{ label: "Ja", value: true, primary: true }, { label: "Nein", value: false }], cancel: false,
      });
      if (!yes) { api.focus(); return; }
      setOptions({ ...options, preferred: tag });
      e.setDocumentProofing(tag, noProof.checked);
    } else if (scopeDoc.checked) e.setDocumentProofing(tag, noProof.checked);
    else e.setProofing(tag, noProof.checked);
    if (detect.checked !== options.detect) setOptions({ ...options, detect: detect.checked });
    refresh();
    api.refresh();
    api.focus();
  };

  const setOptions = (o: ProofingOptions) => {
    options = o;
    try { localStorage.setItem(OPTIONS, JSON.stringify(o)); } catch { /* best effort */ }
    refresh();
  };

  // ----- Datei > Optionen: Dokumentprüfung and Sprache ----------------------------------------
  const optionPages = () => {
    const box = (checked: boolean, label: string) => {
      const input = h("input", { type: "checkbox", checked }) as HTMLInputElement;
      return { input, el: h("label", { class: "check" }, input, label) };
    };
    const caps = box(options.ignoreCaps, "Wörter in GROSSBUCHSTABEN ignorieren");
    const nums = box(options.ignoreNumbers, "Wörter mit Zahlen ignorieren");
    const addr = box(options.ignoreAddresses, "Internet- und Dateiadressen ignorieren");
    const rep = box(options.flagRepeated, "Wiederholte Wörter kennzeichnen");
    const asType = box(options.checkAsYouType, "Rechtschreibung während der Eingabe überprüfen");
    const gram = box(options.grammarAsYouType, "Grammatikfehler während der Eingabe markieren");
    const dictBtn = h("button", { onclick: () => void dictionaryDialog() }, "Benutzerwörterbücher...");
    const recheck = h("button", { onclick: () => { ignored.clear(); refresh(); void messageBox("RaabeWrite", "Die Rechtschreibung und Grammatik wird erneut überprüft. Ignorierte Wörter werden wieder markiert."); } }, "Dokument überprüfen");
    const proofing = h("div", { class: "panel" },
      h("h4", null, "Bei der Rechtschreibkorrektur in RaabeWrite"), caps.el, nums.el, addr.el, rep.el, h("div", null, dictBtn),
      h("h4", null, "Bei der Rechtschreib- und Grammatikkorrektur in RaabeWrite"), asType.el, gram.el, h("div", null, recheck));
    let preferred = options.preferred;
    const list = h("div", { class: "langprefs" });
    const paint = () => list.replaceChildren(...PREFERRED.map((t) => h("div", { class: `langpref${t === preferred ? " on" : ""}`, onclick: () => { preferred = t; paint(); } },
      h("span", null, `${languageName(t)}${t === preferred ? " <bevorzugt>" : ""}`), h("span", null, "Korrekturhilfen installiert"))));
    paint();
    const language = h("div", { class: "panel" },
      h("h4", null, "Erstellungssprachen und Korrekturhilfen"),
      h("p", { class: "muted" }, "Verwalten von Sprachen, die zum Erstellen und Bearbeiten von Dokumenten verwendet werden, einschließlich Korrekturhilfen wie Rechtschreib- und Grammatikprüfung."),
      list);
    const apply = () => setOptions({ ...options, ignoreCaps: caps.input.checked, ignoreNumbers: nums.input.checked, ignoreAddresses: addr.input.checked,
      flagRepeated: rep.input.checked, checkAsYouType: asType.input.checked, grammarAsYouType: gram.input.checked, preferred });
    return [{ label: "Dokumentprüfung", el: proofing }, { label: "Sprache", el: language, apply }];
  };

  /** Benutzerwörterbücher: the words added with Zum Wörterbuch hinzufügen. */
  const dictionaryDialog = async () => {
    const sel = h("select", { size: 10, style: "width:100%" }, ...[...userDict].sort((a, b) => a.localeCompare(b, "de")).map((w) => h("option", { value: w }, w))) as HTMLSelectElement;
    const word = h("input", { type: "text", style: "width:100%" }) as HTMLInputElement;
    const add = h("button", { onclick: () => { const w = word.value.trim(); if (w && !userDict.has(w)) { userDict.add(w); sel.append(h("option", { value: w }, w)); word.value = ""; } } }, "Hinzufügen");
    const del = h("button", { onclick: () => { const w = sel.value; if (w) { userDict.delete(w); sel.querySelector(`option[value="${CSS.escape(w)}"]`)?.remove(); } } }, "Löschen");
    const all = h("button", { onclick: () => { userDict.clear(); sel.replaceChildren(); } }, "Alle löschen");
    const ok = await showDialog({
      title: "Benutzerwörterbuch", body: h("div", { class: "dictdlg" }, h("label", { class: "field" }, "Wort:", word), sel, h("div", { class: "row" }, add, del, all)),
      buttons: [{ label: "OK", value: true, primary: true }, { label: "Abbrechen", value: false }], cancel: false,
    });
    if (ok) {
      try { localStorage.setItem(USER_DICT, JSON.stringify([...userDict])); } catch { /* best effort */ }
    } else {
      try { userDict = new Set(JSON.parse(localStorage.getItem(USER_DICT) ?? "[]")); } catch { /* keep */ }
    }
    refresh();
  };

  // ----- commands ---------------------------------------------------------------------------------
  const c = api.commands;
  c.implement("review.language", { run: () => {}, enabled: () => !!api.editor() });
  c.implement("review.languageProofing", { run: languageDialog, enabled: () => !!api.editor() });

  /** The language shown in the status bar (detected or the text's own). */
  const statusLanguage = (e: Editor): string => {
    const p = e.doc.paras[e.sel.head.p];
    if (!p) return languageName(e.proofingAt().lang);
    const at = langAt(e, p, Math.max(0, Math.min(e.sel.head.o, p.text.length - 1)));
    return languageName(at.lang);
  };

  return { attach, languageName, spellEntries, checkAll, issues, suggest, ignoreAll, ignoreOnce, addToDictionary, deleteRepeat, optionPages,
    statusLanguage, options: () => options, languageDialog };
}
