// Layout > Anordnen > Auswahlbereich (also Start > Markieren > Auswahlbereich...): Word's pane
// "Auswahl" listing the floating objects front first. A click selects an object, a double click
// renames it, the eye hides or shows it; Alle anzeigen / Alle ausblenden, and the arrows change
// the order (Ebene nach vorne / Ebene nach hinten).

import { h, icon, showMenu } from "@raabe/shell";
import type { AppApi } from "./api";

export function selectionPaneFeature(api: AppApi, docArea: HTMLElement) {
  const ed = () => api.editor();
  const c = api.commands;
  let open = false;
  let renaming: string | null = null;
  const pane = h("div", { class: "selpane", hidden: true, role: "complementary", "aria-label": "Auswahl" });
  docArea.append(pane);

  const setOpen = (on: boolean) => {
    open = on;
    pane.hidden = !on;
    lastKey = "";
    render();
    api.refresh();
    if (!on) api.focus();
  };

  let lastKey = "";
  const render = () => {
    if (!open) return;
    const e = ed();
    const list = e?.floatList() ?? [];
    const key = JSON.stringify([list, renaming]);
    if (key === lastKey) return;
    lastKey = key;
    const head = h("div", { class: "sel-head" }, h("span", null, "Auswahl"),
      h("div", null,
        h("button", { "aria-label": "Optionen für den Aufgabenbereich", "data-tip": "Optionen für den Aufgabenbereich",
          onclick: (ev: MouseEvent) => showMenu(c, [{ label: "Schließen", run: () => setOpen(false) }], (ev.currentTarget as HTMLElement).getBoundingClientRect()) }, icon("chevron-down")),
        h("button", { "aria-label": "Bereich schließen", "data-tip": "Bereich schließen", onclick: () => setOpen(false) }, icon("close"))));
    const all = (hidden: boolean) => () => {
      const x = ed();
      if (!x) return;
      for (const o of x.floatList()) if (o.hidden !== hidden) x.setFloatAlt(o.p, o.index, { hidden });
      lastKey = "";
      render();
    };
    const sel = list.find((o) => o.selected);
    const tools = h("div", { class: "sel-tools" },
      h("button", { onclick: all(false), disabled: !list.length }, "Alle anzeigen"),
      h("button", { onclick: all(true), disabled: !list.length }, "Alle ausblenden"));
    const order = h("div", { class: "sel-order" },
      h("button", { class: "icon", "aria-label": "Ebene nach vorne", "data-tip": "Ebene nach vorne", disabled: !sel, onclick: () => { c.run("picture.bringForward"); lastKey = ""; render(); } }, icon("chevron-up")),
      h("button", { class: "icon", "aria-label": "Ebene nach hinten", "data-tip": "Ebene nach hinten", disabled: !sel, onclick: () => { c.run("picture.sendBackward"); lastKey = ""; render(); } }, icon("chevron-down")));
    const rows = h("div", { class: "sel-list" });
    if (!list.length) rows.append(h("div", { class: "sel-empty" }, "Auf dieser Seite gibt es keine Formen."));
    for (const o of list) {
      const id = `${o.p}:${o.index}`;
      const name = renaming === id ? renameBox(o.p, o.index, o.name) : h("span", { class: "sel-name" }, o.name);
      const eye = h("button", { class: "icon sel-eye", "aria-label": o.hidden ? "Anzeigen" : "Ausblenden", "data-tip": o.hidden ? "Anzeigen" : "Ausblenden",
        onclick: (ev: MouseEvent) => { ev.stopPropagation(); ed()?.setFloatAlt(o.p, o.index, { hidden: !o.hidden }); lastKey = ""; render(); } }, icon(o.hidden ? "eye-off" : "eye"));
      const row = h("div", { class: `sel-row${o.selected ? " active" : ""}${o.hidden ? " hidden-obj" : ""}`, role: "button", tabIndex: 0 }, name, eye);
      row.onclick = () => { if (renaming) return; if (!o.hidden) { ed()?.selectFloat(o.p, o.index); api.refresh(); } };
      row.ondblclick = () => { renaming = id; lastKey = ""; render(); };
      rows.append(row);
    }
    pane.replaceChildren(head, tools, order, rows);
    pane.querySelector<HTMLInputElement>(".sel-rename")?.select();
  };

  const renameBox = (p: number, index: number, value: string) => {
    const input = h("input", { class: "sel-rename", type: "text", value, spellcheck: false }) as HTMLInputElement;
    const done = (keep: boolean) => {
      if (renaming === null) return;
      renaming = null;
      if (keep && input.value.trim() && input.value.trim() !== value) ed()?.setFloatAlt(p, index, { name: input.value.trim() });
      lastKey = "";
      render();
    };
    input.addEventListener("keydown", (ev) => {
      ev.stopPropagation();
      if (ev.key === "Enter") { ev.preventDefault(); done(true); }
      else if (ev.key === "Escape") { ev.preventDefault(); done(false); }
    });
    input.addEventListener("blur", () => done(true));
    return input;
  };

  c.implement("layout.selectionPane", { run: () => setOpen(!open), enabled: () => !!ed(), checked: () => open });
  return { render, isOpen: () => open };
}
