// Einfügen > Links as in Word (recorded with probe_word_dialogs.py): Link einfügen (Strg+K;
// Datei oder Webseite, Aktuelles Dokument, Neues Dokument erstellen, E-Mail-Adresse,
// QuickInfo, Textmarke, Zielframe; Link bearbeiten and Link entfernen on an existing link),
// Textmarke (Hinzufügen, Löschen, Gehe zu, Sortieren nach Namen/Ort, Ausgeblendete
// Textmarken) and Querverweis (Verweistyp, Verweisen auf, Als Link einfügen, Oben/unten
// einschließen). Links are followed with Strg+Klick.

import { baseName, dirName, h, icon } from "@raabe/shell";
import type { Editor } from "@raabe/text-engine";
import type { AppApi } from "./api";
import { openModal } from "./modal";
import { loadRecent } from "../store";

const radio = (name: string, label: string, on = false) => {
  const input = h("input", { type: "radio", name, checked: on }) as HTMLInputElement;
  return { input, el: h("label", { class: "check" }, input, label) };
};
const check = (label: string, on = false, disabled = false) => {
  const input = h("input", { type: "checkbox", checked: on, disabled }) as HTMLInputElement;
  return { input, el: h("label", { class: `check${disabled ? " disabled" : ""}` }, input, label) };
};
const text = (value = "", width = 0) => h("input", { type: "text", value, style: width ? `width:${width}px` : "" }) as HTMLInputElement;

/** A list box of entries (Word's list controls): click selects, double click runs. */
function listBox(cls = ""): { el: HTMLElement; set(items: { label: string; value: string; indent?: number; head?: boolean }[]): void; value(): string | null; select(v: string): void; onPick(fn: (v: string) => void): void; onOpen(fn: (v: string) => void): void } {
  const el = h("div", { class: `lb ${cls}`, tabIndex: 0 });
  let items: { label: string; value: string; indent?: number; head?: boolean }[] = [];
  let current: string | null = null;
  let pick = (_v: string) => {}, open = (_v: string) => {};
  const render = () => el.replaceChildren(...items.map((x) => h("div", {
    class: `lbi${x.value === current ? " on" : ""}${x.head ? " head" : ""}`, style: x.indent ? `padding-left:${8 + x.indent * 18}px` : "",
    onmousedown: () => { if (x.head && !x.value) return; current = x.value; render(); pick(x.value); },
    ondblclick: () => open(x.value),
  }, x.label)));
  el.addEventListener("keydown", (e) => {
    const i = items.findIndex((x) => x.value === current);
    const step = e.key === "ArrowDown" ? 1 : e.key === "ArrowUp" ? -1 : 0;
    if (!step) return;
    e.preventDefault();
    const next = items[Math.max(0, Math.min(items.length - 1, i + step))];
    if (next) { current = next.value; render(); pick(next.value); }
  });
  return {
    el, set: (list) => { items = list; render(); }, value: () => current,
    select: (v) => { current = v; render(); }, onPick: (fn) => { pick = fn; }, onOpen: (fn) => { open = fn; },
  };
}

export function linksFeature(api: AppApi) {
  const c = api.commands;
  const ed = () => api.editor();

  // ----- the places in the document (Aktuelles Dokument, Textmarke...) ---------------------------
  const placeItems = (e: Editor) => {
    const items: { label: string; value: string; indent?: number; head?: boolean }[] = [{ label: "Dokumentanfang", value: "_top" }, { label: "Überschriften", value: "", head: true }];
    for (const t of e.referenceTargets("heading")) items.push({ label: t.label, value: `heading:${t.key}`, indent: Math.min(4, t.level ?? 1) });
    items.push({ label: "Textmarken", value: "", head: true });
    for (const b of e.bookmarks().filter((x) => !x.name.startsWith("_"))) items.push({ label: b.name, value: `bm:${b.name}`, indent: 1 });
    return items;
  };
  /** A heading becomes a link target through a hidden _Toc bookmark (Word's way). */
  const anchorFor = (e: Editor, value: string): string => {
    if (value.startsWith("bm:")) return value.slice(3);
    if (value.startsWith("heading:")) return e.headingBookmark(Number(value.slice(8)));
    return value;
  };

  // ----- Link einfügen / Link bearbeiten ---------------------------------------------------------
  const linkDialog = async () => {
    const e = ed();
    if (!e) return;
    const old = e.linkAt();
    const shown = text(old?.text ?? e.selectedText().split("\n")[0] ?? "", 520);
    let tip = old?.tip ?? "";
    let frame = old?.frame ?? "";
    let anchor = old?.anchor && old.href ? old.anchor : "";
    const address = text(old?.href ?? (old?.anchor ? `#${old.anchor}` : ""), 520);
    address.setAttribute("list", "raabe-link-history");
    const kinds = ["file", "place", "new", "mail"] as const;
    let kind: (typeof kinds)[number] = old?.href?.startsWith("mailto:") ? "mail" : old?.anchor && !old.href ? "place" : "file";

    // Datei oder Webseite: the folder's contents, visited pages, recent files.
    let folder = api.documentPath() ? dirName(api.documentPath()!) : await api.platform.documentsDir?.().catch(() => "") ?? "";
    const folderName = h("span", { class: "lkfolder" });
    const files = listBox("lkfiles");
    let view: "folder" | "web" | "recent" = "folder";
    const views = h("div", { class: "lkviews" });
    const showFiles = async () => {
      folderName.textContent = baseName(folder) || folder;
      if (view === "folder") {
        const entries = await api.platform.listDir?.(folder).catch(() => []) ?? [];
        files.set(entries.map((x) => ({ label: `${x.dir ? "📁 " : ""}${x.name}`, value: `${folder}\\${x.name}${x.dir ? "\\" : ""}` })));
      } else if (view === "recent") files.set(loadRecent().map((r) => ({ label: `${baseName(r.path)}  (${dirName(r.path)})`, value: r.path })));
      else files.set(webHistory().map((u) => ({ label: u, value: u })));
      for (const b of Array.from(views.children)) b.classList.toggle("on", b.getAttribute("data-v") === view);
    };
    views.append(...([["folder", "Aktueller Ordner"], ["web", "Besuchte Webseiten"], ["recent", "Zuletzt verwendet"]] as const).map(([v, l]) =>
      h("button", { "data-v": v, onclick: () => { view = v; void showFiles(); } }, l)));
    files.onPick((v) => { if (!v.endsWith("\\")) address.value = v; if (!shown.value) shown.value = v; });
    files.onOpen((v) => { if (v.endsWith("\\")) { folder = v.slice(0, -1); void showFiles(); } });
    const up = h("button", { class: "iconbtn", "data-tip": "Eine Ebene höher", onclick: () => { const d = dirName(folder); if (d && d !== folder) { folder = d; void showFiles(); } } }, icon("arrow-up"));
    const browse = h("button", { class: "iconbtn", "data-tip": "Nach Datei suchen", onclick: async () => {
      const path = await api.platform.openFileDialog().catch(() => null);
      if (path) { address.value = path; if (!shown.value) shown.value = baseName(path); }
    } }, icon("open"));
    const pageFile = h("div", { class: "lkpage" },
      h("div", { class: "row" }, h("span", {}, "Suchen in:"), h("div", { class: "lkfolderbox" }, icon("open"), folderName), up, browse),
      h("div", { class: "lkbrowse" }, views, files.el),
      h("label", { class: "field row" }, "Adresse:", address));

    // Aktuelles Dokument.
    const places = listBox("lkplaces");
    places.set(placeItems(e));
    if (old?.anchor && !old.href) places.select(old.anchor === "_top" ? "_top" : `bm:${old.anchor}`);
    const pagePlace = h("div", { class: "lkpage", hidden: true }, h("div", {}, "Wählen Sie eine Stelle im Dokument aus:"), places.el);

    // Neues Dokument erstellen.
    const newName = text("", 360);
    let newFolder = folder;
    const newPath = h("div", { class: "lkpath" }, newFolder);
    const later = radio("lknew", "Neues Dokument später bearbeiten", true), now = radio("lknew", "Neues Dokument jetzt bearbeiten");
    const pageNew = h("div", { class: "lkpage", hidden: true },
      h("label", { class: "field" }, "Name des neuen Dokuments:", newName),
      h("div", { class: "row" }, h("div", {}, h("div", {}, "Vollständiger Pfad:"), newPath),
        h("button", { onclick: async () => { const p = await api.platform.saveFileDialog(newName.value || "Dokument", "docx").catch(() => null); if (p) { newFolder = dirName(p); newName.value = baseName(p); newPath.textContent = newFolder; } } }, "Ändern...")),
      h("div", {}, "Wann soll bearbeitet werden:"), later.el, now.el);

    // E-Mail-Adresse.
    const mail = text(old?.href?.startsWith("mailto:") ? decodeURIComponent(old.href.slice(7).split("?")[0]) : "", 420);
    const subject = text(old?.href?.includes("subject=") ? decodeURIComponent(old.href.split("subject=")[1]) : "", 420);
    const mails = listBox("lkmails");
    mails.set(mailHistory().map((m) => ({ label: `mailto:${m}`, value: m })));
    mails.onPick((v) => { mail.value = v; });
    const pageMail = h("div", { class: "lkpage", hidden: true },
      h("label", { class: "field" }, "E-Mail-Adresse:", mail), h("label", { class: "field" }, "Betreff:", subject),
      h("div", {}, "Zuletzt verwendete E-Mail-Adressen:"), mails.el);

    const pages = { file: pageFile, place: pagePlace, new: pageNew, mail: pageMail };
    const side = h("div", { class: "lkside" }, h("div", { class: "lkhead" }, "Link zu:"), ...([
      ["file", "Datei oder Webseite", "globe"], ["place", "Aktuelles Dokument", "document"], ["new", "Neues Dokument erstellen", "new"], ["mail", "E-Mail-Adresse", "mail"],
    ] as const).map(([k, l, ic]) => h("button", { "data-k": k, onclick: () => setKind(k) }, icon(ic, "l"), h("span", {}, l))));
    const setKind = (k: typeof kind) => {
      kind = k;
      for (const [n, p] of Object.entries(pages)) p.hidden = n !== k;
      for (const b of Array.from(side.querySelectorAll("button"))) b.classList.toggle("on", b.getAttribute("data-k") === k);
    };
    const sideButtons = h("div", { class: "lkbuttons" },
      h("button", { onclick: () => void quickInfo() }, "QuickInfo..."),
      h("button", { onclick: () => void pickPlace() }, "Textmarke..."),
      h("button", { onclick: () => void pickFrame() }, "Zielframe..."));
    const quickInfo = async () => {
      const t = text(tip, 300);
      openModal("QuickInfo für Link festlegen", h("label", { class: "field" }, "QuickInfo-Text:", t), [
        { label: "OK", primary: true, run: () => { tip = t.value; return true; } }, { label: "Abbrechen", run: () => true }], { width: 380 });
    };
    const pickPlace = async () => {
      const list = listBox("lkplaces");
      list.set(placeItems(e));
      openModal("Textmarke im Dokument auswählen", h("div", {}, h("div", {}, "Wählen Sie eine vorhandene Stelle im Dokument aus:"), list.el), [
        { label: "OK", primary: true, run: () => { const v = list.value(); if (v) { anchor = anchorFor(e, v); if (kind === "file" && !address.value) setKind("place"); places.select(v); } return true; } },
        { label: "Abbrechen", run: () => true }], { width: 420 });
    };
    const pickFrame = async () => {
      const sel = h("select", {}, ...([["", "Standardeinstellung für Seite (keiner)"], ["_self", "Gleicher Frame"], ["_top", "Ganzes Fenster"], ["_blank", "Neues Fenster"], ["_parent", "Übergeordneter Frame"]] as const)
        .map(([v, l]) => h("option", { value: v }, l))) as HTMLSelectElement;
      sel.value = frame;
      openModal("Zielframe festlegen", h("label", { class: "field" }, "Wählen Sie den Frame, in dem das Dokument erscheinen soll:", sel), [
        { label: "OK", primary: true, run: () => { frame = sel.value; return true; } }, { label: "Abbrechen", run: () => true }], { width: 420 });
    };
    const body = h("div", { class: "lkdialog" }, side,
      h("div", { class: "lkmain" }, h("label", { class: "field row" }, "Anzuzeigender Text:", shown), pageFile, pagePlace, pageNew, pageMail),
      sideButtons);

    const ok = async (): Promise<boolean> => {
      const label = shown.value.trim();
      if (kind === "mail") {
        if (!mail.value.trim()) return false;
        const addr = mail.value.trim().replace(/^mailto:/i, "");
        rememberMail(addr);
        e.insertLink({ text: label || `mailto:${addr}`, href: `mailto:${addr}${subject.value ? `?subject=${encodeURIComponent(subject.value)}` : ""}`, tip });
      } else if (kind === "place") {
        const v = places.value();
        if (!v) return false;
        const a = anchorFor(e, v);
        e.insertLink({ text: label || (v.startsWith("bm:") ? v.slice(3) : a), anchor: a, tip, frame });
      } else if (kind === "new") {
        const name = newName.value.trim();
        if (!name) return false;
        const path = `${newFolder}\\${/\.docx$/i.test(name) ? name : `${name}.docx`}`;
        await api.createEmptyDocument(path, now.input.checked);
        e.insertLink({ text: label || path, href: path, tip, frame });
      } else {
        let target = address.value.trim();
        if (!target) return false;
        if (target.startsWith("#")) { e.insertLink({ text: label || target.slice(1), anchor: target.slice(1), tip, frame }); return true; }
        // "www.schule.de" is a web address (Word adds http://).
        if (/^www\./i.test(target)) target = `http://${target}`;
        if (/^https?:/i.test(target)) rememberWeb(target);
        e.insertLink({ text: label || target, href: target, anchor: anchor || undefined, tip, frame });
      }
      api.refresh();
      api.focus();
      return true;
    };
    const buttons = [
      ...(old ? [{ label: "Link entfernen", run: () => { e.removeLink(); api.refresh(); api.focus(); return true; } }] : []),
      { label: "OK", primary: true, run: ok },
      { label: "Abbrechen", run: () => true },
    ];
    openModal(old ? "Link bearbeiten" : "Link einfügen", body, buttons, { width: 760, cls: "linkdlg" });
    setKind(kind);
    await showFiles();
  };

  // ----- Textmarke ------------------------------------------------------------------------------
  const bookmarkDialog = () => {
    const e = ed();
    if (!e) return;
    const name = text("", 0);
    const list = listBox("bmlist");
    const byName = radio("bmsort", "Namen", true), byPlace = radio("bmsort", "Ort");
    const hidden = check("Ausgeblendete Textmarken");
    const fill = () => {
      let marks = e.bookmarks().filter((b) => hidden.input.checked || !b.name.startsWith("_"));
      if (byName.input.checked) marks = [...marks].sort((a, b) => a.name.localeCompare(b.name, "de"));
      list.set(marks.map((b) => ({ label: b.name, value: b.name })));
      sync();
    };
    const sync = () => {
      const v = name.value.trim();
      const exists = e.bookmarks().some((b) => b.name === v);
      m.button("Hinzufügen")!.disabled = !v || !/^[\p{L}][\p{L}\p{N}_]{0,39}$/u.test(v);
      m.button("Löschen")!.disabled = !exists;
      m.button("Gehe zu")!.disabled = !exists;
    };
    list.onPick((v) => { name.value = v; sync(); });
    list.onOpen((v) => { e.goToBookmark(v); api.focus(); });
    name.addEventListener("input", () => { list.select(name.value); sync(); });
    for (const r of [byName, byPlace, hidden]) r.input.onchange = fill;
    const body = h("div", { class: "bmdialog" },
      h("div", { class: "bmleft" }, h("label", { class: "field" }, "Textmarkenname:", name), list.el,
        h("div", { class: "row" }, h("span", {}, "Sortieren nach:"), h("div", {}, byName.el, byPlace.el)), hidden.el));
    const m = openModal("Textmarke", body, [
      { label: "Hinzufügen", primary: true, run: () => { if (!e.addBookmark(name.value.trim())) return false; api.refresh(); api.focus(); return true; } },
      { label: "Löschen", run: () => { e.deleteBookmark(name.value.trim()); name.value = ""; fill(); m.button("Abbrechen")!.textContent = "Schließen"; } },
      { label: "Gehe zu", run: () => { e.goToBookmark(name.value.trim()); m.button("Abbrechen")!.textContent = "Schließen"; } },
      { label: "Abbrechen", run: () => true },
    ], { width: 560, cls: "bmdlg" });
    // Word suggests the selected word as the name.
    const sel = e.selectedText().trim();
    if (sel && /^[\p{L}][\p{L}\p{N}_]*$/u.test(sel)) name.value = sel.slice(0, 40);
    fill();
  };

  // ----- Querverweis ------------------------------------------------------------------------------
  type RefKind = Parameters<Editor["referenceTargets"]>[0];
  const REF_TYPES: { label: string; kind: RefKind; list: string; what: [string, string][] }[] = [
    { label: "Nummeriertes Element", kind: "numbered", list: "Für welches nummerierte Element:", what: [["number", "Absatznummer"], ["number", "Absatznummer (kein Kontext)"], ["number", "Absatznummer (vollständiger Kontext)"], ["text", "Absatztext"], ["page", "Seitenzahl"], ["aboveBelow", "Oben/unten"]] },
    { label: "Überschrift", kind: "heading", list: "Für welche Überschrift:", what: [["text", "Überschriftentext"], ["page", "Seitenzahl"], ["number", "Überschriftennummer"], ["number", "Überschriftennummer (kein Kontext)"], ["number", "Überschriftennummer (vollständiger Kontext)"], ["aboveBelow", "Oben/unten"]] },
    { label: "Textmarke", kind: "bookmark", list: "Für welche Textmarke:", what: [["text", "Textmarkentext"], ["page", "Seitenzahl"], ["number", "Absatznummer"], ["number", "Absatznummer (kein Kontext)"], ["number", "Absatznummer (vollständiger Kontext)"], ["aboveBelow", "Oben/unten"]] },
    { label: "Fußnote", kind: "footnote", list: "Für welche Fußnote:", what: [["noteNumber", "Fußnotennummer"], ["page", "Seitenzahl"], ["aboveBelow", "Oben/unten"], ["noteNumber", "Fußnotennummer (formatiert)"]] },
    { label: "Endnote", kind: "endnote", list: "Für welche Endnote:", what: [["noteNumber", "Endnotennummer"], ["page", "Seitenzahl"], ["aboveBelow", "Oben/unten"], ["noteNumber", "Endnotennummer (formatiert)"]] },
    ...["Abbildung", "Formel", "Tabelle"].map((l) => ({
      label: l, kind: { caption: l } as RefKind, list: `Für welche Beschriftung:`,
      what: [["whole", "Gesamte Beschriftung"], ["labelNumber", "Nur Kategorie und Nummer"], ["captionText", "Nur Beschriftungstext"], ["page", "Seitenzahl"], ["aboveBelow", "Oben/unten"]] as [string, string][],
    })),
  ];
  let lastType = 0;
  const crossRefDialog = () => {
    const e = ed();
    if (!e) return;
    const type = h("select", {}, ...REF_TYPES.map((t, i) => h("option", { value: String(i) }, t.label))) as HTMLSelectElement;
    const what = h("select", {}) as HTMLSelectElement;
    const asLink = check("Als Link einfügen", true);
    const aboveBelow = check("Oben/unten einschließen");
    const sep = check("Zahlen trennen mit", false, true);
    const listLabel = h("div", {});
    const list = listBox("xrlist");
    const fill = () => {
      const t = REF_TYPES[Number(type.value)];
      what.replaceChildren(...t.what.map(([v, l], i) => h("option", { value: `${i}:${v}` }, l)));
      listLabel.textContent = t.list;
      list.set(e.referenceTargets(t.kind).map((x) => ({ label: x.label, value: x.key, indent: x.level ? x.level - 1 : 0 })));
      sync();
    };
    const sync = () => {
      const w = what.value.split(":")[1];
      aboveBelow.input.disabled = w === "aboveBelow";
      aboveBelow.el.classList.toggle("disabled", w === "aboveBelow");
      m.button("Einfügen")!.disabled = !list.value();
    };
    type.onchange = () => { lastType = Number(type.value); fill(); };
    what.onchange = sync;
    list.onPick(sync);
    const insert = () => {
      const t = REF_TYPES[Number(type.value)];
      const key = list.value();
      if (!key) return;
      e.insertCrossReference(t.kind, key, what.value.split(":")[1], asLink.input.checked, aboveBelow.input.checked && !aboveBelow.input.disabled);
      m.button("Abbrechen")!.textContent = "Schließen";
      api.refresh();
    };
    list.onOpen(insert);
    const body = h("div", { class: "xrdialog" },
      h("div", { class: "row two" }, h("label", { class: "field" }, "Verweistyp:", type), h("label", { class: "field" }, "Verweisen auf:", what)),
      h("div", { class: "row two" }, asLink.el, aboveBelow.el),
      h("div", { class: "row" }, sep.el, h("input", { type: "text", disabled: true, style: "width:40px" })),
      listLabel, list.el);
    const m = openModal("Querverweis", body, [
      { label: "Einfügen", primary: true, run: insert },
      { label: "Abbrechen", run: () => true },
    ], { width: 560, cls: "xrdlg", onClose: () => api.focus() });
    type.value = String(lastType);
    fill();
  };

  // ----- following links (Strg+Klick) ---------------------------------------------------------------
  const follow = (target: { href?: string; anchor?: string }) => {
    const e = ed();
    if (target.anchor && !target.href) {
      if (target.anchor === "_top") e?.caretAt({ p: 0, o: 0 });
      else e?.goToBookmark(target.anchor);
      return;
    }
    if (target.href) void api.platform.openExternal?.(target.href).catch(() => {});
  };

  c.implement("insert.link", { run: () => void linkDialog(), enabled: () => !!ed() });
  c.implement("insert.bookmark", { run: bookmarkDialog, enabled: () => !!ed() });
  c.implement("insert.crossRef", { run: crossRefDialog, enabled: () => !!ed() });
  c.implement("ref.crossRef", { run: crossRefDialog, enabled: () => !!ed() });
  return { linkDialog, follow, attach: (e: Editor) => { e.onFollowLink = follow; } };
}

// ----- what the user typed before (Besuchte Webseiten, Zuletzt verwendete E-Mail-Adressen) -------

function history(key: string): string[] {
  try { const v = JSON.parse(localStorage.getItem(key) ?? "[]"); return Array.isArray(v) ? v.filter((x) => typeof x === "string") : []; } catch { return []; }
}
function push(key: string, v: string): void {
  try { localStorage.setItem(key, JSON.stringify([v, ...history(key).filter((x) => x !== v)].slice(0, 20))); } catch { /* private mode */ }
}
const webHistory = () => history("raabe.links.web");
const mailHistory = () => history("raabe.links.mail");
const rememberWeb = (v: string) => push("raabe.links.web", v);
const rememberMail = (v: string) => push("raabe.links.mail", v);
