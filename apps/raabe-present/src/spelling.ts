// Rechtschreibung (F7, D-070): the slides' text from the current slide on, word by word with
// RaabeWrite's dictionaries (Hunspell, compiled into the app as in RaabeCalc, D-065); each unknown word
// is selected on its slide and offered with PowerPoint's choices: Ignorieren, Alle ignorieren,
// Hinzufügen, Ändern, Alle ändern. Capitals, words with digits and addresses are left out (Office's
// defaults). PowerPoint shows these in a pane; ours is a box beside the slide (to be made a pane).

import { h, icon, messageBox } from "@raabe/shell";
import { spTreeOf } from "@raabe/ooxml";
import type { PresentApp } from "./app";
import * as S from "./shapeops";
import * as T from "./textops";

const USER_DICT = "raabe.userDictionary";
const WORD = /[\p{L}][\p{L}\p{M}'’-]*[\p{L}]|[\p{L}]/gu;

function userWords(): Set<string> {
  try { return new Set(JSON.parse(localStorage.getItem(USER_DICT) ?? "[]") as string[]); } catch { return new Set(); }
}

function addUserWord(w: string): void {
  try { const s = userWords(); s.add(w); localStorage.setItem(USER_DICT, JSON.stringify([...s])); } catch { /* this run only */ }
}

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

type Answer = { kind: "ignore" | "ignoreAll" | "add" | "change" | "changeAll" | "cancel"; to: string };

export async function checkSpelling(app: PresentApp): Promise<void> {
  const deck = app.deck;
  const n = deck.slides.length;
  const ignored = new Set<string>(), known = userWords();
  const changeAll = new Map<string, string>();
  const lang = "de-DE";
  const start = app.current();
  for (let k = 0; k < n; k++) {
    const si = (start + k) % n;
    const path = deck.slides[si].path;
    const shapes = S.shapeEls(spTreeOf(deck.pptx.doc(path))!).filter((s) => S.txBodyOf(s));
    for (const sh of shapes) {
      const id = S.shapeId(sh);
      const body = S.txBodyOf(sh)!;
      for (let pi = 0; pi < T.paras(body).length; pi++) {
        let offset = 0;
        for (const w of wordsOf(T.paraText(T.paras(body)[pi]))) {
          const at = w.at + offset;
          if (ignored.has(w.word) || known.has(w.word)) continue;
          const replace = (by: string) => {
            deck.change("Rechtschreibung", [path], () => { T.deleteRange(body, { p: pi, o: at }, { p: pi, o: at + w.word.length }); T.insertText(body, { p: pi, o: at }, by); });
            offset += by.length - w.word.length;
          };
          if (changeAll.has(w.word)) { replace(changeAll.get(w.word)!); continue; }
          const [ok] = await app.platform.spellCheck(lang, [w.word]).catch(() => [true]);
          if (ok) continue;
          if (app.current() !== si) app.setCurrent(si);
          app.editor.enterText(id, { p: pi, o: at + w.word.length });
          if (app.editor.text) { app.editor.text.anchor = { p: pi, o: at }; app.editor.render(); }
          const a = await ask(w.word, await app.platform.spellSuggest(lang, w.word).catch(() => []));
          if (a.kind === "cancel") { app.editor.focus(); return; }
          if (a.kind === "ignoreAll") ignored.add(w.word);
          if (a.kind === "add") { addUserWord(w.word); known.add(w.word); }
          if (a.kind === "change" || a.kind === "changeAll") {
            if (a.kind === "changeAll") changeAll.set(w.word, a.to);
            replace(a.to);
          }
        }
      }
    }
  }
  app.editor.endText();
  app.editor.render();
  await messageBox("RaabePresent", "Die Rechtschreibprüfung ist abgeschlossen.");
}

function ask(word: string, suggestions: string[]): Promise<Answer> {
  return new Promise((resolve) => {
    const list = h("select", { size: 6, class: "sp-list", "aria-label": "Vorschläge" }, ...(suggestions.length ? suggestions : ["(Keine Vorschläge)"]).map((s, i) => h("option", { value: suggestions.length ? s : "", selected: i === 0 || undefined }, s))) as HTMLSelectElement;
    const layer = h("div", { class: "find-layer" });
    const done = (kind: Answer["kind"]) => { layer.remove(); removeEventListener("keydown", keys, true); resolve({ kind, to: list.value || word }); };
    const keys = (e: KeyboardEvent) => { if (e.key === "Escape") { e.preventDefault(); e.stopPropagation(); done("cancel"); } };
    const btn = (label: string, kind: Answer["kind"]) => h("button", { type: "button", onclick: () => done(kind) }, label);
    const dialog = h("div", { class: "dialog find-dialog", role: "dialog", "aria-label": "Rechtschreibung" },
      h("div", { class: "dlg-title" }, h("span", { class: "t" }, "Rechtschreibung"), h("button", { "aria-label": "Schließen", onclick: () => done("cancel") }, icon("close"))),
      h("div", { class: "find-body" },
        h("div", { class: "find-fields" }, h("b", null, word), list),
        h("div", { class: "find-buttons" }, btn("Ignorieren", "ignore"), btn("Alle ignorieren", "ignoreAll"), btn("Hinzufügen", "add"), btn("Ändern", "change"), btn("Alle ändern", "changeAll"))),
      h("div", { class: "muted", style: "padding:0 12px 10px" }, "Deutsch (Deutschland)"));
    layer.append(dialog);
    document.body.append(layer);
    addEventListener("keydown", keys, true);
    list.ondblclick = () => done("change");
    (dialog.querySelector(".find-buttons button:nth-child(4)") as HTMLButtonElement).focus();
  });
}
