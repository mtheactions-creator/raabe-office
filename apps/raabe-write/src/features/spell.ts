// Rechtschreibprüfung while typing (red wavy lines), the context menu with suggestions and
// the dialog "Rechtschreibung und Grammatik" (F7). Words are checked in the background by
// the platform (Hunspell dictionaries in Rust); results are cached per word and language.
// As in Word, words in CAPITALS and words with digits are ignored, and the word being typed
// is not marked until the caret leaves it.

import type { Editor, Para, TextMark } from "@raabe/text-engine";
import { closePopup, h, icon, messageBox, showDialog, showPopup } from "@raabe/shell";
import type { AppApi } from "./api";

interface Word { a: number; b: number; text: string; lang: string }

const LANGS: [string, string][] = [["de-DE", "Deutsch (Deutschland)"], ["en-US", "Englisch (USA)"], ["en-GB", "Englisch (Großbritannien)"]];
export const languageName = (lang: string) => LANGS.find(([l]) => l.toLowerCase() === lang.toLowerCase())?.[1] ?? lang;

const USER_DICT = "raabe.userDictionary";

/** The words of a paragraph with their language (runs' lang, else the document default). */
function words(p: Para, fallback: string): Word[] {
  const out: Word[] = [];
  const re = /[\p{L}\p{M}]+(?:['’-][\p{L}\p{M}]+)*/gu;
  const langAt: string[] = [];
  let pos = 0;
  for (const r of p.runs) {
    const l = r.props.lang ?? fallback;
    for (let i = 0; i < r.len; i++) langAt[pos + i] = l;
    pos += r.len;
  }
  for (const m of p.text.matchAll(re)) {
    const a = m.index!, b = a + m[0].length;
    // Internet and e-mail addresses, words next to digits (12a) and words in capitals are left out.
    let s0 = a, e0 = b;
    while (s0 > 0 && !/\s/.test(p.text[s0 - 1])) s0--;
    while (e0 < p.text.length && !/\s/.test(p.text[e0])) e0++;
    if (/:\/\/|^www\.|@/i.test(p.text.slice(s0, e0))) continue;
    if (/\p{N}/u.test(p.text[a - 1] ?? "") || /\p{N}/u.test(p.text[b] ?? "")) continue;
    if (m[0].length < 2 || m[0] === m[0].toLocaleUpperCase("de-DE")) continue;
    out.push({ a, b, text: m[0], lang: langAt[a] ?? fallback });
  }
  return out;
}

export function spellFeature(api: AppApi) {
  const cache = new Map<string, Map<string, boolean>>();
  const pending = new Map<string, Set<string>>();
  const ignored = new Set<string>();
  let userDict = new Set<string>();
  try { userDict = new Set(JSON.parse(localStorage.getItem(USER_DICT) ?? "[]")); } catch { /* none yet */ }
  const marks = new WeakMap<Para, { key: string; marks: TextMark[]; complete: boolean }>();
  let timer = 0;
  let enabled = true;
  const editors = new WeakSet<Editor>();

  const known = (lang: string, w: string): boolean | undefined => {
    if (ignored.has(w) || userDict.has(w)) return true;
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

  const fallbackLang = (e: Editor) => e.doc.styles.defaultsR.lang ?? "de-DE";

  /** Spelling marks of a paragraph (cached until its text or the caret's word changes). */
  const marksFor = (e: Editor, story: string, para: Para): readonly TextMark[] | undefined => {
    if (!enabled || para.block || !para.text) return undefined;
    // The word with the caret is not judged while it is being typed.
    const h0 = e.sel.head;
    const caretPara = story === (e.doc.activeStory === "main" ? "main" : e.doc.activeStory) && e.doc.paras[h0.p] === para && !e.hasSelection ? h0.o : -1;
    const key = `${caretPara}|${ignored.size}|${userDict.size}`;
    const hit = marks.get(para);
    if (hit && hit.complete && hit.key === key) return hit.marks;
    const out: TextMark[] = [];
    let complete = true;
    for (const w of words(para, fallbackLang(e))) {
      if (caretPara >= w.a && caretPara <= w.b) continue;
      const parts = w.text.split(/[-]/).filter((x) => x.length > 1 && x !== x.toLocaleUpperCase("de-DE"));
      const k = known(w.lang, w.text);
      if (k === true) continue;
      if (k === false) {
        // Hyphenated words are right when their parts are (Kohlenstoff-Atom).
        const partsOk = parts.length > 1 && parts.every((x) => known(w.lang, x) === true);
        if (!partsOk) out.push({ start: w.a, end: w.b, kind: "spell" });
        if (parts.length > 1) for (const x of parts) if (known(w.lang, x) === undefined) { ask(w.lang, x); complete = false; }
        continue;
      }
      ask(w.lang, w.text);
      complete = false;
    }
    marks.set(para, { key, marks: out, complete });
    return out;
  };

  /** Connects an editor (each document's). */
  const attach = (e: Editor) => {
    if (editors.has(e)) return;
    editors.add(e);
    e.extraMarks = (story, para) => marksFor(e, story, para);
    e.onContextMenu = (x, y, pos) => void contextMenu(e, x, y, pos);
  };

  const misspelled = (e: Editor, p: number, a: number): Word | null => {
    const para = e.doc.paras[p];
    for (const w of words(para, fallbackLang(e))) if (a >= w.a && a < w.b && known(w.lang, w.text) === false) return w;
    return null;
  };

  const addToDictionary = (w: string) => {
    userDict.add(w);
    try { localStorage.setItem(USER_DICT, JSON.stringify([...userDict])); } catch { /* best effort */ }
    api.editor()?.view.refresh();
  };

  // ----- context menu ------------------------------------------------------------------------
  const contextMenu = async (e: Editor, x: number, y: number, pos: { p: number; o: number }) => {
    const c = api.commands;
    const menu = h("div", { class: "menu", role: "menu" });
    const item = (label: string, run: () => void, ic?: string, disabled = false) => {
      menu.append(h("button", { role: "menuitem", disabled, onclick: () => { closePopup(); run(); api.focus(); } }, icon(ic), label));
    };
    const cmdItem = (id: string, label?: string) => {
      const d = c.get(id);
      menu.append(h("button", { role: "menuitem", disabled: !c.isEnabled(id), onclick: () => { closePopup(); void c.run(id); } }, icon(d.icon), label ?? d.label));
    };
    const w = misspelled(e, pos.p, pos.o);
    if (w) {
      const sugg = await api.platform.spellSuggest(w.lang, w.text).catch(() => [] as string[]);
      for (const s of sugg.slice(0, 5)) {
        menu.append(h("button", { role: "menuitem", style: "font-weight:600", onclick: () => { closePopup(); e.replaceRange(pos.p, w.a, w.b, s); api.focus(); } }, icon(undefined), s));
      }
      if (!sugg.length) menu.append(h("button", { role: "menuitem", disabled: true }, icon(undefined), "(Keine Vorschläge)"));
      menu.append(h("hr"));
      item("Alle ignorieren", () => { ignored.add(w.text); e.view.refresh(); });
      item("Zum Wörterbuch hinzufügen", () => addToDictionary(w.text));
      menu.append(h("hr"));
    }
    cmdItem("edit.cut"); cmdItem("edit.copy"); cmdItem("edit.paste", "Einfügen");
    menu.append(h("hr"));
    if (e.inTable()) {
      cmdItem("table.insertAbove", "Zeilen oberhalb einfügen"); cmdItem("table.insertBelow", "Zeilen unterhalb einfügen");
      cmdItem("table.insertLeft", "Spalten links einfügen"); cmdItem("table.insertRight", "Spalten rechts einfügen");
      cmdItem("table.deleteRows", "Zeilen löschen"); cmdItem("table.deleteColumns", "Spalten löschen");
      if (e.cellBlock()?.multi) cmdItem("table.merge");
      menu.append(h("hr"));
    }
    if (e.selectedImage()) { cmdItem("layout.wrapText", "Textumbruch"); menu.append(h("hr")); }
    cmdItem("font.dialog", "Schriftart..."); cmdItem("para.dialog", "Absatz...");
    if (c.has("insert.comment")) cmdItem("insert.comment", "Neuer Kommentar");
    showPopup(menu, new DOMRect(x, y, 0, 0));
  };

  // ----- F7: Rechtschreibung und Grammatik -------------------------------------------------
  const allWordsChecked = async (e: Editor) => {
    const fb = fallbackLang(e);
    const todo = new Map<string, Set<string>>();
    for (const p of e.doc.paras) {
      if (p.block) continue;
      for (const w of words(p, fb)) if (known(w.lang, w.text) === undefined) {
        const set = todo.get(w.lang) ?? new Set<string>();
        set.add(w.text);
        todo.set(w.lang, set);
      }
    }
    for (const [lang, set] of todo) for (const w of set) ask(lang, w);
    clearTimeout(timer);
    await flush();
  };

  const nextError = (e: Editor, from: { p: number; o: number }): { p: number; w: Word } | null => {
    const fb = fallbackLang(e);
    const paras = e.doc.paras;
    for (let p = from.p; p < paras.length; p++) {
      if (paras[p].block) continue;
      for (const w of words(paras[p], fb)) {
        if (p === from.p && w.a < from.o) continue;
        const parts = w.text.split("-").filter((x) => x.length > 1);
        if (known(w.lang, w.text) === false && !(parts.length > 1 && parts.every((x) => known(w.lang, x) === true))) return { p, w };
      }
    }
    return null;
  };

  const dialog = async () => {
    const e = api.editor();
    if (!e) return;
    await allWordsChecked(e);
    let at = { p: e.sel.head.p, o: Math.min(e.sel.head.o, e.sel.anchor.o) };
    let wrapped = false;
    for (;;) {
      let found = nextError(e, at);
      if (!found && !wrapped && at.p > 0) { wrapped = true; at = { p: 0, o: 0 }; found = nextError(e, at); }
      if (!found) { await messageBox("Rechtschreibung", "Die Rechtschreibprüfung ist abgeschlossen."); api.focus(); return; }
      const { p, w } = found;
      e.selectRange(p, w.a, w.b);
      const sugg = await api.platform.spellSuggest(w.lang, w.text).catch(() => [] as string[]);
      const notIn = h("div", { class: "spell-word" }, w.text);
      const list = h("select", { size: 6, style: "width:100%" }, ...(sugg.length ? sugg : ["(Keine Vorschläge)"]).map((s) => h("option", { value: s }, s)));
      list.selectedIndex = 0;
      const choice = await showDialog<string>({
        title: `Rechtschreibung: ${languageName(w.lang)}`,
        body: h("div", { style: "width:420px" }, h("div", { style: "margin-bottom:4px" }, "Nicht im Wörterbuch:"), notIn,
          h("div", { style: "margin:10px 0 4px" }, "Vorschläge:"), list),
        buttons: [
          { label: "Ändern", value: "change", primary: true }, { label: "Alle ändern", value: "changeAll" },
          { label: "Einmal ignorieren", value: "skip" }, { label: "Alle ignorieren", value: "ignoreAll" },
          { label: "Hinzufügen", value: "add" }, { label: "Abbrechen", value: "cancel" },
        ],
        cancel: "cancel",
      });
      const pick = sugg.length ? list.value : "";
      if (choice === "cancel") { api.focus(); return; }
      if (choice === "change" && pick) { e.replaceRange(p, w.a, w.b, pick); at = { p, o: w.a + pick.length }; continue; }
      if (choice === "changeAll" && pick) { e.replaceAll(w.text, pick, { matchCase: true, wholeWord: true }); at = { p, o: w.a + pick.length }; continue; }
      if (choice === "ignoreAll") ignored.add(w.text);
      if (choice === "add") addToDictionary(w.text);
      at = { p, o: w.b };
    }
  };

  const c = api.commands;
  c.implement("review.spelling", { run: dialog, enabled: () => !!api.editor() });

  // ----- Sprache (Korrekturhilfesprache) ------------------------------------------------------
  const languageDialog = async () => {
    const e = api.editor();
    if (!e) return;
    const sel = h("select", { size: 4, style: "width:100%" }, ...LANGS.map(([v, l]) => h("option", { value: v }, l)));
    sel.value = e.charLang() ?? fallbackLang(e);
    const off = h("input", { type: "checkbox", checked: !enabled });
    const ok = await showDialog({
      title: "Sprache", body: h("div", { style: "width:320px" }, h("div", { style: "margin-bottom:6px" }, "Markierten Text kennzeichnen als:"), sel,
        h("label", { class: "check", style: "margin-top:8px" }, off, "Rechtschreibung nicht prüfen")),
      buttons: [{ label: "OK", value: true, primary: true }, { label: "Abbrechen", value: false }], cancel: false,
    });
    if (!ok) return;
    enabled = !off.checked;
    e.setLang(sel.value);
    e.view.refresh();
    api.focus();
  };
  c.implement("review.language", { run: languageDialog, enabled: () => !!api.editor() });

  return { attach, languageName };
}
