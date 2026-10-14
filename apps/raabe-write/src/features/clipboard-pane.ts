// Start > Zwischenablage (dialog launcher): Word's pane Zwischenablage (recorded 2026-10-14),
// the Office clipboard of RaabeWrite: up to 24 copied or cut items, newest first, collected
// while the pane is open (or always with "Sammeln ohne Anzeige der Office-Zwischenablage").
// A click pastes an item; its menu pastes or deletes it; Alle einfügen, Alle löschen, Optionen.

import { h, icon, showMenu, type MenuEntry } from "@raabe/shell";
import type { AppApi } from "./api";

interface Item {
  text: string;
  html: string;
  /** A picture or shape copied as itself (no text). */
  object: boolean;
}

const MAX = 24;
const OPTS_KEY = "raabe.clipboard.options";

export function clipboardPaneFeature(api: AppApi, docArea: HTMLElement) {
  const ed = () => api.editor();
  const c = api.commands;
  const items: Item[] = [];
  let open = false;
  let opts = { autoShow: false, ctrlCTwice: false, collectHidden: false };
  try { opts = { ...opts, ...JSON.parse(localStorage.getItem(OPTS_KEY) ?? "{}") }; } catch { /* defaults */ }
  const saveOpts = () => { try { localStorage.setItem(OPTS_KEY, JSON.stringify(opts)); } catch { /* not kept */ } };

  const pane = h("div", { class: "clippane", hidden: true, role: "complementary", "aria-label": "Zwischenablage" });
  docArea.append(pane);

  const setOpen = (on: boolean) => {
    open = on;
    pane.hidden = !on;
    render();
    api.refresh();
    if (!on) api.focus();
  };

  let lastCopy = 0;
  /** Called for every copy or cut in RaabeWrite. */
  const collect = (text: string, html: string, object: boolean) => {
    const now = Date.now();
    // Strg+C zweimal: the pane opens (Word's option).
    if (opts.ctrlCTwice && now - lastCopy < 600 && !open) setOpen(true);
    lastCopy = now;
    if (!open && !opts.collectHidden && !opts.autoShow) return;
    if (!text && !object) return;
    const same = items.findIndex((x) => x.text === text && x.html === html);
    if (same >= 0) items.splice(same, 1);
    items.unshift({ text, html, object });
    if (items.length > MAX) items.length = MAX;
    // Office-Zwischenablage automatisch anzeigen: with two items collected.
    if (!open && opts.autoShow && items.length >= 2) setOpen(true);
    else render();
  };

  const paste = (x: Item) => {
    const e = ed();
    if (!e) return;
    e.paste(x.text, x.html);
    api.focus();
  };

  const render = () => {
    if (!open) return;
    const head = h("div", { class: "clip-head" }, h("span", null, "Zwischenablage"),
      h("div", null,
        h("button", { "aria-label": "Optionen für den Aufgabenbereich", "data-tip": "Optionen für den Aufgabenbereich",
          onclick: (ev: MouseEvent) => showMenu(c, [{ label: "Schließen", run: () => setOpen(false) }], (ev.currentTarget as HTMLElement).getBoundingClientRect()) }, icon("chevron-down")),
        h("button", { "aria-label": "Bereich schließen", "data-tip": "Bereich schließen", onclick: () => setOpen(false) }, icon("close"))));
    const tools = h("div", { class: "clip-tools" },
      h("button", { disabled: !items.length, onclick: () => { for (const x of [...items].reverse()) paste(x); } }, "Alle einfügen"),
      h("button", { disabled: !items.length, onclick: () => { items.length = 0; render(); } }, "Alle löschen"));
    const list = h("div", { class: "clip-list" });
    if (!items.length) {
      list.append(h("div", { class: "clip-hint" }, "Klicken Sie auf ein Element, um es einzufügen:"),
        h("div", { class: "clip-empty" }, "Die Zwischenablage ist leer.", h("br"), "Verwenden Sie \"Kopieren\" oder \"Ausschneiden\", um Elemente aufzunehmen."));
    } else {
      list.append(h("div", { class: "clip-hint" }, "Klicken Sie auf ein Element, um es einzufügen:"));
      for (const x of items) {
        const preview = x.object && !x.text ? "Grafik" : x.text.replace(/\s+/g, " ").trim().slice(0, 160);
        const menu: MenuEntry[] = [
          { label: "Einfügen", icon: "paste", run: () => paste(x) },
          { label: "Löschen", icon: "close", run: () => { const i = items.indexOf(x); if (i >= 0) items.splice(i, 1); render(); } },
        ];
        const arrow = h("button", { class: "clip-arrow", "aria-label": "Optionen", tabIndex: -1 }, icon("chevron-down"));
        const row = h("div", { class: "clip-item", role: "button", tabIndex: 0 }, h("span", { class: "clip-icon" }, icon(x.object && !x.text ? "picture" : "document")), h("span", { class: "clip-text" }, preview), arrow);
        row.onclick = (ev) => { if ((ev.target as Element).closest(".clip-arrow")) return; paste(x); };
        arrow.onclick = (ev) => { ev.stopPropagation(); showMenu(c, menu, row.getBoundingClientRect()); };
        row.oncontextmenu = (ev) => { ev.preventDefault(); showMenu(c, menu, new DOMRect(ev.clientX, ev.clientY, 0, 0)); };
        list.append(row);
      }
    }
    const optEntry = (label: string, key: keyof typeof opts): MenuEntry => ({
      label, get checked() { return opts[key]; }, run: () => { opts[key] = !opts[key]; saveOpts(); },
    });
    const options = h("button", { class: "clip-options", onclick: (ev: MouseEvent) => showMenu(c, [
      optEntry("Office-Zwischenablage automatisch anzeigen", "autoShow"),
      optEntry("Office-Zwischenablage anzeigen, wenn STRG+C zweimal gedrückt wurde", "ctrlCTwice"),
      optEntry("Sammeln ohne Anzeige der Office-Zwischenablage", "collectHidden"),
      { label: "Office-Zwischenablagensymbol auf Taskleiste anzeigen", run: () => {}, disabled: true },
      { label: "Status beim Kopieren in der Taskleiste anzeigen", run: () => {}, disabled: true },
    ], (ev.currentTarget as HTMLElement).getBoundingClientRect()) }, "Optionen", icon("chevron-up"));
    pane.replaceChildren(head, tools, list, h("div", { class: "clip-foot" }, options));
  };

  c.implement("edit.clipboardPane", { run: () => setOpen(!open), enabled: () => !!ed(), checked: () => open });
  return { collect, isOpen: () => open, pane };
}
