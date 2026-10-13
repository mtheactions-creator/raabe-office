// Suchen (Strg+F) in Word's Navigationsbereich, with the tabs Überschriften and Ergebnisse,
// and the dialog Suchen und Ersetzen (Strg+H) with Suchen, Ersetzen and Gehe zu (Strg+G).

import { h, icon, messageBox, type Command } from "@raabe/shell";
import type { AppApi } from "./api";

type Match = { p: number; a: number; b: number };

export const FIND_COMMANDS: Command[] = [
  { id: "find.next", label: "Weitersuchen", icon: "find", hidden: true },
];

export function findFeature(api: AppApi, docArea: HTMLElement) {
  const ed = () => api.editor();
  const c = api.commands;
  let open = false;
  let matches: Match[] = [];
  let current = -1;
  let tab: "headings" | "results" = "results";
  let opts = { matchCase: false, wholeWord: false };

  // ----- the pane ---------------------------------------------------------------------------
  const input = h("input", { type: "search", placeholder: "Dokument durchsuchen", spellcheck: false, "aria-label": "Dokument durchsuchen" });
  const count = h("div", { class: "nav-count" });
  const list = h("div", { class: "nav-list" });
  const tabs = h("div", { class: "nav-tabs" });
  const tabBtn = (id: "headings" | "results", label: string) => h("button", { "data-k": id, onclick: () => { tab = id; render(); } }, label);
  tabs.append(tabBtn("headings", "Überschriften"), tabBtn("results", "Ergebnisse"));
  const pane = h("div", { class: "navpane", hidden: true },
    h("div", { class: "nav-head" }, h("span", null, "Navigation"),
      h("button", { "aria-label": "Schließen", "data-tip": "Schließen", onclick: () => setOpen(false) }, icon("close"))),
    h("div", { class: "nav-search" }, icon("search"), input),
    count, tabs, list);
  docArea.append(pane);

  const setOpen = (on: boolean) => {
    open = on;
    pane.hidden = !on;
    docArea.classList.toggle("with-nav", on);
    if (!on) { ed()?.showFindMarks(null); matches = []; api.focus(); }
    api.refresh();
  };

  const search = () => {
    const e = ed();
    const q = input.value;
    matches = e && q ? e.findAll(q, opts) : [];
    if (current >= matches.length) current = matches.length ? 0 : -1;
    e?.showFindMarks(matches.length ? matches : null, matches[current] ? { p: matches[current].p, a: matches[current].a } : undefined);
    if (q) tab = "results";
    render();
  };

  const go = (i: number) => {
    const e = ed();
    if (!e || !matches.length) return;
    current = (i + matches.length) % matches.length;
    const m = matches[current];
    e.selectRange(m.p, m.a, m.b);
    e.showFindMarks(matches, { p: m.p, a: m.a });
    render();
  };

  const render = () => {
    tabs.querySelectorAll("button").forEach((b) => b.classList.toggle("active", (b as HTMLElement).dataset.k === tab));
    const e = ed();
    const q = input.value;
    count.textContent = !q ? "" : matches.length ? `Ergebnis ${current >= 0 ? current + 1 : 0} von ${matches.length}` : "Keine Ergebnisse";
    list.replaceChildren();
    if (!e) return;
    if (tab === "headings") {
      const hs = e.headings();
      if (!hs.length) list.append(h("div", { class: "nav-empty" }, "Erstellen Sie eine interaktive Gliederung Ihres Dokuments: Wenden Sie Überschriften-Formatvorlagen auf den Text an."));
      for (const x of hs) {
        list.append(h("button", { class: "nav-heading", style: `padding-left:${8 + (x.level - 1) * 14}px`, onclick: () => e.goToParagraph(x.p) }, x.text));
      }
      return;
    }
    const paras = e.doc.paras;
    matches.slice(0, 500).forEach((m, i) => {
      const t = paras[m.p].text.replace(/[\t\u000b\ufffc]/g, " ");
      const before = t.slice(Math.max(0, m.a - 30), m.a), after = t.slice(m.b, m.b + 40);
      list.append(h("button", { class: `nav-result${i === current ? " active" : ""}`, onclick: () => go(i) },
        (m.a > 30 ? "…" : "") + before, h("b", null, t.slice(m.a, m.b)), after));
    });
  };

  let timer = 0;
  input.addEventListener("input", () => { clearTimeout(timer); timer = window.setTimeout(() => { current = -1; search(); }, 120); });
  input.addEventListener("keydown", (ev) => {
    ev.stopPropagation();
    if (ev.key === "Enter") { ev.preventDefault(); if (!matches.length) search(); go(ev.shiftKey ? current - 1 : current + 1); input.focus(); }
    else if (ev.key === "Escape") { ev.preventDefault(); api.focus(); }
  });

  c.implement("view.navPane", { run: () => setOpen(!open), enabled: () => !!ed(), checked: () => open });
  c.implement("edit.find", {
    run: () => { if (!open) setOpen(true); tab = "results"; render(); input.focus(); input.select(); },
    enabled: () => !!ed(),
  });

  // ----- Suchen und Ersetzen -------------------------------------------------------------------
  let dialog: HTMLElement | null = null;
  const replaceDialog = (page: "find" | "replace" | "goto") => {
    dialog?.remove();
    const what = h("input", { type: "text", value: input.value || ed()?.selectedText().split("\n")[0] || "", style: "width:100%" });
    const by = h("input", { type: "text", style: "width:100%" });
    const pageNo = h("input", { type: "text", style: "width:100%" });
    const matchCase = h("input", { type: "checkbox", checked: opts.matchCase });
    const wholeWord = h("input", { type: "checkbox", checked: opts.wholeWord });
    const status = h("div", { class: "muted", style: "min-height:18px" });
    const readOpts = () => { opts = { matchCase: matchCase.checked, wholeWord: wholeWord.checked }; };
    let pos = -1;
    const next = () => {
      readOpts();
      const e = ed();
      if (!e || !what.value) return;
      const all = e.findAll(what.value, opts);
      if (!all.length) { status.textContent = "Das Suchelement wurde nicht gefunden."; return; }
      const sel = e.sel;
      const after = all.findIndex((m) => m.p > sel.head.p || (m.p === sel.head.p && m.a >= Math.max(sel.head.o, sel.anchor.o)));
      pos = after >= 0 ? after : 0;
      if (after < 0) status.textContent = "Ende des Dokuments erreicht – die Suche beginnt am Anfang.";
      else status.textContent = "";
      e.selectRange(all[pos].p, all[pos].a, all[pos].b);
    };
    const replaceOne = () => {
      readOpts();
      const e = ed();
      if (!e || !what.value) return;
      const sel = e.selectedText();
      const same = opts.matchCase ? sel === what.value : sel.toLocaleLowerCase("de-DE") === what.value.toLocaleLowerCase("de-DE");
      if (same && e.sel.head.p === e.sel.anchor.p) {
        const a = Math.min(e.sel.anchor.o, e.sel.head.o), b = Math.max(e.sel.anchor.o, e.sel.head.o);
        e.replaceRange(e.sel.head.p, a, b, by.value);
      }
      next();
    };
    const replaceAll = async () => {
      readOpts();
      const e = ed();
      if (!e || !what.value) return;
      const n = e.replaceAll(what.value, by.value, opts);
      await messageBox("RaabeWrite", n ? `Alle erledigt. ${n} Ersetzungen wurden vorgenommen.` : "Das Suchelement wurde nicht gefunden.");
    };
    const goto = () => {
      const n = Number(pageNo.value);
      if (Number.isFinite(n) && n > 0) ed()?.goToPage(n);
    };
    const pages: Record<string, HTMLElement> = {
      find: h("div", null, h("label", { class: "field" }, "Suchen nach:", what)),
      replace: h("div", null, h("label", { class: "field" }, "Ersetzen durch:", by)),
      goto: h("div", null, h("label", { class: "field" }, "Seitenzahl eingeben:", pageNo)),
    };
    const options = h("fieldset", null, h("legend", null, "Suchoptionen"),
      h("label", { class: "check" }, matchCase, "Groß-/Kleinschreibung beachten"), h("label", { class: "check" }, wholeWord, "Nur ganzes Wort suchen"));
    const btn = (label: string, fn: () => void, primary = false) => h("button", { class: primary ? "primary" : "", onclick: fn }, label);
    const buttons = h("div", { class: "dlg-buttons" });
    const tabsEl = h("div", { class: "tabs" });
    const show = (k: "find" | "replace" | "goto") => {
      pages.find.hidden = k === "goto";
      pages.replace.hidden = k !== "replace";
      pages.goto.hidden = k !== "goto";
      options.hidden = k === "goto";
      tabsEl.querySelectorAll("button").forEach((b) => b.classList.toggle("active", (b as HTMLElement).dataset.k === k));
      buttons.replaceChildren(...(k === "goto" ? [btn("Gehe zu", goto, true)]
        : k === "replace" ? [btn("Ersetzen", replaceOne), btn("Alle ersetzen", () => void replaceAll()), btn("Weitersuchen", next, true)]
        : [btn("Weitersuchen", next, true)]), btn("Abbrechen", close));
      (k === "goto" ? pageNo : what).focus();
    };
    for (const [k, label] of [["find", "Suchen"], ["replace", "Ersetzen"], ["goto", "Gehe zu"]] as const) {
      tabsEl.append(h("button", { "data-k": k, onclick: () => show(k) }, label));
    }
    const close = () => { dialog?.remove(); dialog = null; api.focus(); };
    dialog = h("div", { class: "dialog floating", role: "dialog", "aria-label": "Suchen und Ersetzen" },
      h("div", { class: "dlg-title" }, h("span", { class: "t" }, "Suchen und Ersetzen"), h("button", { "aria-label": "Schließen", onclick: close }, icon("close"))),
      h("div", { class: "dlg-body", style: "width:440px" }, tabsEl, pages.find, pages.replace, pages.goto, options, status), buttons);
    dialog.addEventListener("keydown", (ev) => {
      ev.stopPropagation();
      if (ev.key === "Escape") { ev.preventDefault(); close(); }
      else if (ev.key === "Enter" && !(ev.target instanceof HTMLButtonElement)) { ev.preventDefault(); if (pages.goto.hidden) next(); else goto(); }
    });
    document.body.append(dialog);
    show(page);
    what.select();
  };
  c.implement("edit.replace", { run: () => replaceDialog("replace"), enabled: () => !!ed() });
  c.implement("edit.goto", { run: () => replaceDialog("goto"), enabled: () => !!ed() });

  /** After edits: the result list follows the text (as in Word). */
  const refresh = () => { if (open && input.value) search(); else if (open && tab === "headings") render(); };
  return { refresh, isOpen: () => open };
}
