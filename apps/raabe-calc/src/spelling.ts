// Rechtschreibung (Überprüfen > Rechtschreibung, F7) as Excel checks a sheet: the text cells (not
// formulas) from the active cell on, row by row, then from the top; for each word the dictionary
// does not know Excel's dialog: Nicht im Wörterbuch, Vorschläge, Einmal ignorieren, Alle
// ignorieren, Zum Wörterbuch hinzufügen, Ändern, Alle ändern, Wörterbuchsprache. Words in
// CAPITALS, words with digits and Internet addresses are left out (Office's defaults). The
// dictionaries are RaabeWrite's (Hunspell, compiled in).

import { cellName } from "@raabe/sheet-engine";
import { h, icon, messageBox } from "@raabe/shell";
import type { CalcApp } from "./app";

const USER_DICT = "raabe.userDictionary";
const LANGS: [string, string][] = [["de-DE", "Deutsch (Deutschland)"], ["en-GB", "Englisch (Vereinigtes Königreich)"], ["en-US", "Englisch (USA)"], ["fr-FR", "Französisch (Frankreich)"], ["es-ES", "Spanisch (Spanien)"]];
const WORD = /[\p{L}][\p{L}\p{M}'’-]*[\p{L}]|[\p{L}]/gu;

function userWords(): Set<string> {
  try { return new Set(JSON.parse(localStorage.getItem(USER_DICT) ?? "[]") as string[]); } catch { return new Set(); }
}

function addUserWord(w: string): void {
  try { const s = userWords(); s.add(w); localStorage.setItem(USER_DICT, JSON.stringify([...s])); } catch { /* stays for this run */ }
}

/** The words of a text worth checking, with their places. */
export function wordsOf(text: string): { word: string; at: number }[] {
  if (/\b(https?:\/\/|www\.)|@\S+\.\S/.test(text)) text = text.replace(/\b(https?:\/\/|www\.)\S+|\S+@\S+\.\S+/g, (m) => " ".repeat(m.length));
  const out: { word: string; at: number }[] = [];
  for (const m of text.matchAll(WORD)) {
    const w = m[0];
    const before = text[m.index! - 1] ?? "", after = text[m.index! + w.length] ?? "";
    if (/\d/.test(before) || /\d/.test(after)) continue;
    if (w.length > 1 && w === w.toLocaleUpperCase("de") && w !== w.toLocaleLowerCase("de")) continue;
    out.push({ word: w, at: m.index! });
  }
  return out;
}

export function spellingFeature(app: CalcApp) {
  async function run(): Promise<void> {
    const sheet = app.sheet;
    const start = { ...sheet.view.active };
    let lang = "de-DE";
    const ignored = new Set<string>();
    const changeAll = new Map<string, string>();
    const known = userWords();
    // The text cells in reading order from the active cell, then from the top.
    const cells = [...sheet.cells].filter(([, c]) => typeof c.v === "string" && c.f === undefined)
      .map(([k]) => ({ row: Math.floor(k / 16384), col: k % 16384 })).sort((a, b) => a.row - b.row || a.col - b.col);
    const from = cells.findIndex((p) => p.row > start.row || (p.row === start.row && p.col >= start.col));
    const order = from > 0 ? [...cells.slice(from), ...cells.slice(0, from)] : cells;
    let cancelled = false;
    for (const p of order) {
      if (cancelled) break;
      let text = sheet.get(p.row, p.col)?.v;
      if (typeof text !== "string") continue;
      let offset = 0;
      for (const w of wordsOf(text)) {
        const at = w.at + offset;
        if (ignored.has(w.word) || known.has(w.word)) continue;
        const replace = (by: string) => {
          text = (text as string).slice(0, at) + by + (text as string).slice(at + w.word.length);
          offset += by.length - w.word.length;
          const cur = sheet.get(p.row, p.col)!;
          app.change("Rechtschreibung", () => { app.history.cell(sheet, p.row, p.col); sheet.set(p.row, p.col, { ...cur, v: text as string }); });
        };
        if (changeAll.has(w.word)) { replace(changeAll.get(w.word)!); continue; }
        const [ok] = await app.platform.spellCheck(lang, [w.word]);
        if (ok) continue;
        app.grid.selectCell(p.row, p.col);
        const answer = await ask(w.word, await app.platform.spellSuggest(lang, w.word), lang, cellName(p.row, p.col));
        lang = answer.lang;
        if (answer.kind === "cancel") { cancelled = true; break; }
        if (answer.kind === "ignoreAll") ignored.add(w.word);
        if (answer.kind === "add") { addUserWord(w.word); known.add(w.word); }
        if (answer.kind === "change" || answer.kind === "changeAll") {
          if (answer.kind === "changeAll") changeAll.set(w.word, answer.to);
          replace(answer.to);
        }
      }
    }
    app.grid.focus();
    if (!cancelled) await messageBox("RaabeCalc", "Die Rechtschreibprüfung für dieses Blatt ist fertig.");
  }

  type Answer = { kind: "ignore" | "ignoreAll" | "add" | "change" | "changeAll" | "cancel"; to: string; lang: string };

  function ask(word: string, suggestions: string[], lang: string, where: string): Promise<Answer> {
    return new Promise((resolve) => {
      const notIn = h("input", { type: "text", value: word, class: "sp-word", "aria-label": "Nicht im Wörterbuch" }) as HTMLInputElement;
      const list = h("select", { size: 6, class: "sp-list", "aria-label": "Vorschläge" }, ...(suggestions.length ? suggestions : ["(Keine Vorschläge)"]).map((s, i) => h("option", { value: suggestions.length ? s : "", selected: i === 0 || undefined }, s))) as HTMLSelectElement;
      const langSel = h("select", null, ...LANGS.map(([v, l]) => h("option", { value: v, selected: v === lang || undefined }, l))) as HTMLSelectElement;
      const target = () => (notIn.value !== word ? notIn.value : list.value || word);
      const layer = h("div", { class: "dialog-layer" });
      const done = (kind: Answer["kind"]) => { layer.remove(); removeEventListener("keydown", keys, true); resolve({ kind, to: target(), lang: langSel.value }); };
      const keys = (e: KeyboardEvent) => { if (e.key === "Escape") { e.preventDefault(); e.stopPropagation(); done("cancel"); } };
      const btn = (label: string, kind: Answer["kind"] | null) => h("button", { type: "button", disabled: kind ? undefined : true, onclick: () => kind && done(kind) }, label);
      const title = () => `Rechtschreibung: ${LANGS.find(([v]) => v === langSel.value)?.[1] ?? ""}`;
      const head = h("span", { class: "t" }, title());
      langSel.onchange = () => { head.textContent = title(); };
      const dialog = h("div", { class: "dialog sp-dialog", role: "dialog", "aria-label": "Rechtschreibung" },
        h("div", { class: "dlg-title" }, head, h("button", { "aria-label": "Schließen", onclick: () => done("cancel") }, icon("close"))),
        h("div", { class: "dlg-body sp-body" },
          h("div", null, h("div", { class: "fc-label" }, `Nicht im Wörterbuch (${where}):`), notIn, h("div", { class: "fc-label" }, "Vorschläge:"), list,
            h("label", { class: "field" }, "Wörterbuchsprache:", langSel)),
          h("div", { class: "sp-buttons" }, btn("Einmal ignorieren", "ignore"), btn("Alle ignorieren", "ignoreAll"), btn("Zum Wörterbuch hinzufügen", "add"),
            btn("Ändern", "change"), btn("Alle ändern", "changeAll"), btn("AutoKorrektur", null))),
        h("div", { class: "dlg-buttons" }, btn("Optionen...", null), btn("Rückgängig", null), btn("Abbrechen", "cancel")));
      layer.append(dialog);
      document.body.append(layer);
      addEventListener("keydown", keys, true);
      list.ondblclick = () => done("change");
      (dialog.querySelector(".sp-buttons button:nth-child(4)") as HTMLButtonElement).focus();
    });
  }

  return { run };
}
