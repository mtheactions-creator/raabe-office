// Office's task pane (Excel's "Diagrammbereich formatieren", "Form formatieren" …): docked at the right
// of the window, with a title, an optional switch between option sets ("Diagrammoptionen ▾",
// "Textoptionen"), a row of icon tabs and collapsible sections of controls. The pane is plain data;
// every control reads its value when drawn and writes it back at once, as Office's panes apply
// changes immediately. Its contents come from what Excel shows (D-066).

import { h, icon } from "./dom";
import { colorGrid } from "./palette";
import { showPopup, closePopup } from "./popup";

export type PaneControl =
  | { type: "radio"; label: string; checked: () => boolean; pick: () => void; disabled?: () => boolean }
  | { type: "check"; label: string; value: () => boolean; set: (on: boolean) => void; disabled?: () => boolean }
  | { type: "color"; label: string; value: () => string | null; set: (hex: string | null) => void; auto?: string; disabled?: () => boolean }
  | {
      type: "number"; label: string; value: () => number | null; set: (v: number) => void;
      /** Shown after the number ("%", " pt", "°"); the box shows it as Excel does. */
      unit?: string; min?: number; max?: number; step?: number; decimals?: number; disabled?: () => boolean;
      /** A slider beside the box (Transparenz). */
      slider?: boolean;
      /** "Automatisch" / "Zurücksetzen" button beside the box (axis bounds). */
      reset?: { label: string; isAuto: () => boolean; run: () => void };
    }
  | { type: "select"; label: string; options: { label: string; value: string }[]; value: () => string; set: (v: string) => void; disabled?: () => boolean }
  | { type: "text"; label: string; value: () => string; set: (v: string) => void; disabled?: () => boolean }
  | { type: "button"; label: string; run: () => void; disabled?: () => boolean }
  | { type: "heading"; label: string }
  /** Several controls on one row (Breite / Höhe …). */
  | { type: "row"; items: PaneControl[] }
  /**
   * Farbverlaufsstopps: the gradient bar with a handle per stop, Farbverlaufstopp hinzufügen /
   * entfernen; the chosen stop's Farbe, Position, Transparenz follow as ordinary controls.
   */
  | {
      type: "gradient"; label: string; stops: () => { pos: number; color: string; transparency?: number }[]; selected: () => number;
      select: (i: number) => void; add: () => void; remove: () => void; move: (i: number, pos: number) => void;
    }
  /** Musterfüllung: a grid of patterns (each drawn by `draw` as an SVG path in an 8 x 8 cell grid). */
  | { type: "pattern"; label: string; options: { value: string; label: string; path: string }[]; value: () => string; fg: () => string; bg: () => string; set: (v: string) => void }
  /** Free content (a preview, a gallery). */
  | { type: "custom"; render: () => HTMLElement };

export interface PaneSection { title: string; open?: boolean; controls: () => PaneControl[] }
export interface PaneTab { id: string; label: string; icon: string; sections: () => PaneSection[] }
export interface PaneSet { label: string; tabs: PaneTab[] }

export interface PaneSpec {
  /** The pane's title, e.g. "Diagrammbereich formatieren". */
  title: string;
  /** The option sets ("Diagrammoptionen", "Textoptionen"); the first one's label opens `choices`. */
  sets: PaneSet[];
  /** The elements the first set's dropdown offers (Diagrammbereich, Zeichnungsfläche, Achsen …). */
  choices?: { label: string; pick: () => void }[];
}

export class TaskPane {
  readonly el = h("aside", { class: "taskpane", role: "complementary", hidden: true });
  private spec: PaneSpec | null = null;
  private set = 0;
  private tab = 0;
  private closed = new Set<string>();
  private opened = new Set<string>();
  onClose: (() => void) | null = null;

  constructor(private host: HTMLElement = document.body) {
    this.host.append(this.el);
  }

  get isOpen(): boolean { return !this.el.hidden; }

  /** Shows a spec (keeping the chosen set and tab when the title's kind of pane stays the same). */
  show(spec: PaneSpec, keep = true): void {
    if (!keep || this.spec?.title !== spec.title) { this.set = 0; this.tab = 0; }
    this.spec = spec;
    this.el.hidden = false;
    this.el.setAttribute("aria-label", spec.title);
    document.body.classList.add("taskpane-open");
    this.render();
  }

  /** Shows a spec at one tab (by id) with one section opened (by title): the menus' "…optionen..." entries. */
  showAt(spec: PaneSpec, tabId: string, section?: string): void {
    this.show(spec);
    spec.sets.forEach((s, si) => s.tabs.forEach((t, ti) => {
      if (t.id !== tabId) return;
      this.set = si; this.tab = ti;
      if (section) { const key = `${t.id}/${section}`; this.closed.delete(key); this.opened.add(key); }
    }));
    this.render();
  }

  close(): void {
    if (this.el.hidden) return;
    this.el.hidden = true;
    this.spec = null;
    document.body.classList.remove("taskpane-open");
    this.onClose?.();
  }

  /** Draws the values again (after a change made elsewhere: undo, the ribbon). */
  refresh(): void {
    if (this.spec && !this.el.hidden) this.render();
  }

  private render(): void {
    const spec = this.spec!;
    const set = spec.sets[Math.min(this.set, spec.sets.length - 1)];
    const tab = set.tabs[Math.min(this.tab, set.tabs.length - 1)];
    const scroll = this.el.querySelector(".tp-body")?.scrollTop ?? 0;
    const sets = h("div", { class: "tp-sets" }, ...spec.sets.map((s, i) => {
      const b = h("button", { type: "button", class: `tp-set${i === this.set ? " on" : ""}` }, s.label);
      if (i === 0 && spec.choices?.length) {
        b.append(h("span", { class: "tp-chev" }, "˅"));
        b.addEventListener("click", () => {
          if (this.set !== 0) { this.set = 0; this.tab = 0; this.render(); return; }
          const list = h("div", { class: "menu" }, ...spec.choices!.map((c) => h("button", { type: "button", class: "item", onclick: () => { closePopup(); c.pick(); } }, c.label)));
          showPopup(list, b.getBoundingClientRect());
        });
      } else b.addEventListener("click", () => { this.set = i; this.tab = 0; this.render(); });
      return b;
    }));
    const tabs = h("div", { class: "tp-tabs", role: "tablist" }, ...set.tabs.map((t, i) =>
      h("button", { type: "button", role: "tab", class: `tp-tab${i === this.tab ? " on" : ""}`, "aria-label": t.label, "data-tip": t.label,
        "aria-selected": String(i === this.tab), onclick: () => { this.tab = i; this.render(); } }, icon(t.icon, "l"))));
    const body = h("div", { class: "tp-body" }, ...tab.sections().map((s) => this.section(tab.id, s)));
    this.el.replaceChildren(
      h("div", { class: "tp-head" }, h("div", { class: "tp-title" }, spec.title),
        h("button", { type: "button", class: "tp-close", "aria-label": "Schließen", "data-tip": "Schließen", onclick: () => this.close() }, "✕")),
      sets, tabs, body);
    body.scrollTop = scroll;
  }

  private section(tab: string, s: PaneSection): HTMLElement {
    const key = `${tab}/${s.title}`;
    const open = this.opened.has(key) || (!!s.open && !this.closed.has(key));
    const head = h("button", { type: "button", class: "tp-sec-head", "aria-expanded": String(open), onclick: () => {
      if (open) { this.opened.delete(key); this.closed.add(key); } else { this.closed.delete(key); this.opened.add(key); }
      this.render();
    } }, h("span", { class: "tp-tri" }, open ? "◢" : "▷"), s.title);
    const box = h("div", { class: "tp-sec" }, head);
    if (open) box.append(h("div", { class: "tp-sec-body" }, ...s.controls().map((c) => this.control(c))));
    return box;
  }

  private control(c: PaneControl): HTMLElement {
    const dis = "disabled" in c && c.disabled ? c.disabled() : false;
    const changed = () => queueMicrotask(() => this.refresh());
    switch (c.type) {
      case "heading":
        return h("div", { class: "tp-heading" }, c.label);
      case "radio": {
        const input = h("input", { type: "radio", checked: c.checked(), disabled: dis, onchange: () => { c.pick(); changed(); } }) as HTMLInputElement;
        return h("label", { class: "tp-line tp-choice" }, input, c.label);
      }
      case "check": {
        const input = h("input", { type: "checkbox", checked: c.value(), disabled: dis, onchange: () => { c.set(input.checked); changed(); } }) as HTMLInputElement;
        return h("label", { class: "tp-line tp-choice" }, input, c.label);
      }
      case "color": {
        const v = c.value();
        const btn = h("button", { type: "button", class: "tp-color", disabled: dis, "aria-label": c.label },
          h("span", { class: "tp-swatch", style: `background:${v ?? c.auto ?? "transparent"}` }), h("span", { class: "tp-chev" }, "˅"));
        btn.addEventListener("click", () => {
          const pop = h("div", { class: "menu palette" });
          if (c.auto) pop.append(h("button", { type: "button", class: "item", onclick: () => { closePopup(); c.set(null); changed(); } }, "Automatisch"));
          pop.append(...colorGrid((hex) => { closePopup(); c.set(hex.startsWith("#") ? hex.toUpperCase() : `#${hex.toUpperCase()}`); changed(); }));
          showPopup(pop, btn.getBoundingClientRect());
        });
        return h("div", { class: "tp-line" }, h("span", { class: "tp-label" }, c.label), btn);
      }
      case "number": {
        const v = c.value();
        const fmt = (x: number | null) => (x === null ? "" : (c.decimals !== undefined ? x.toFixed(c.decimals) : String(Math.round(x * 100) / 100)).replace(".", ",") + (c.unit ?? ""));
        const input = h("input", { type: "text", class: "tp-num", value: fmt(v), disabled: dis }) as HTMLInputElement;
        const commit = () => {
          const n = parseFloat(input.value.replace(",", ".").replace(/[^\d.-]/g, ""));
          if (!Number.isFinite(n)) { input.value = fmt(c.value()); return; }
          // Enter commits, and Chromium fires change after it too: the same text is not set twice (one undo step).
          if (input.dataset.done === input.value || n === c.value()) { input.value = fmt(n); return; }
          input.dataset.done = input.value;
          c.set(Math.max(c.min ?? -Infinity, Math.min(c.max ?? Infinity, n)));
          changed();
        };
        input.addEventListener("change", commit);
        input.addEventListener("keydown", (e) => { if (e.key === "Enter") { e.preventDefault(); commit(); } });
        const step = (d: number) => { const cur = c.value() ?? 0; c.set(Math.max(c.min ?? -Infinity, Math.min(c.max ?? Infinity, Math.round((cur + d * (c.step ?? 1)) * 1000) / 1000))); changed(); };
        const spin = h("span", { class: "tp-spin" },
          h("button", { type: "button", tabindex: "-1", disabled: dis, "aria-label": "Höher", onclick: () => step(1) }, "˄"),
          h("button", { type: "button", tabindex: "-1", disabled: dis, "aria-label": "Niedriger", onclick: () => step(-1) }, "˅"));
        const parts: HTMLElement[] = [h("span", { class: "tp-label" }, c.label)];
        if (c.slider) {
          const range = h("input", { type: "range", class: "tp-slider", min: String(c.min ?? 0), max: String(c.max ?? 100), step: String(c.step ?? 1), value: String(v ?? 0), disabled: dis }) as HTMLInputElement;
          range.addEventListener("change", () => { c.set(Number(range.value)); changed(); });
          parts.push(range);
        }
        parts.push(h("span", { class: "tp-numbox" }, input, spin));
        if (c.reset) parts.push(h("button", { type: "button", class: "tp-reset", disabled: c.reset.isAuto(), onclick: () => { c.reset!.run(); changed(); } }, c.reset.label));
        return h("div", { class: "tp-line" }, ...parts);
      }
      case "select": {
        const sel = h("select", { class: "tp-select", disabled: dis }, ...c.options.map((o) => h("option", { value: o.value, selected: o.value === c.value() }, o.label))) as HTMLSelectElement;
        sel.addEventListener("change", () => { c.set(sel.value); changed(); });
        return h("div", { class: "tp-line" }, h("span", { class: "tp-label" }, c.label), sel);
      }
      case "text": {
        const input = h("input", { type: "text", class: "tp-text", value: c.value(), disabled: dis }) as HTMLInputElement;
        input.addEventListener("change", () => { c.set(input.value); changed(); });
        return h("div", { class: "tp-line" }, h("span", { class: "tp-label" }, c.label), input);
      }
      case "button":
        return h("div", { class: "tp-line" }, h("button", { type: "button", class: "tp-btn", disabled: dis, onclick: () => { c.run(); changed(); } }, c.label));
      case "row":
        return h("div", { class: "tp-row" }, ...c.items.map((i) => this.control(i)));
      case "gradient": {
        const stops = c.stops();
        const sel = c.selected();
        const css = [...stops].sort((a, b) => a.pos - b.pos).map((s) => `${s.color} ${s.pos}%`).join(", ");
        const bar = h("div", { class: "tp-grad-bar", style: `background:linear-gradient(to right, ${css || "#fff, #fff"})`, "aria-label": c.label, role: "group" });
        stops.forEach((s, i) => {
          const handle = h("button", { type: "button", class: `tp-grad-stop${i === sel ? " on" : ""}`, style: `left:${s.pos}%`, "aria-label": `Stopp ${i + 1} von ${stops.length}`, "data-tip": `Stopp ${i + 1} von ${stops.length}` },
            h("span", { style: `background:${s.color}` }));
          handle.addEventListener("pointerdown", (e) => {
            e.preventDefault();
            c.select(i);
            const r = bar.getBoundingClientRect();
            let moved = false;
            const move = (ev: PointerEvent) => { moved = true; handle.style.left = `${Math.max(0, Math.min(100, ((ev.clientX - r.left) / r.width) * 100))}%`; };
            const up = (ev: PointerEvent) => {
              removeEventListener("pointermove", move);
              removeEventListener("pointerup", up);
              if (moved) c.move(i, Math.round(Math.max(0, Math.min(100, ((ev.clientX - r.left) / r.width) * 100))));
              changed();
            };
            addEventListener("pointermove", move);
            addEventListener("pointerup", up);
          });
          bar.append(handle);
        });
        const add = h("button", { type: "button", class: "tp-icon-btn", "aria-label": "Farbverlaufstopp hinzufügen", "data-tip": "Farbverlaufstopp hinzufügen", onclick: () => { c.add(); changed(); } }, "+");
        const del = h("button", { type: "button", class: "tp-icon-btn", "aria-label": "Farbverlaufstopp entfernen", "data-tip": "Farbverlaufstopp entfernen", disabled: stops.length <= 2, onclick: () => { c.remove(); changed(); } }, "✕");
        return h("div", { class: "tp-grad" }, h("div", { class: "tp-line" }, h("span", { class: "tp-label" }, c.label), add, del), bar);
      }
      case "pattern": {
        const cur = c.value(), fg = c.fg(), bg = c.bg();
        const grid = h("div", { class: "tp-patterns", role: "listbox", "aria-label": c.label }, ...c.options.map((o) => {
          const svg = `<svg width="20" height="20" viewBox="0 0 6 6" aria-hidden="true"><rect width="6" height="6" fill="${bg}"/><path d="${o.path}" fill="${fg}"/></svg>`;
          const b = h("button", { type: "button", class: `tp-pattern${o.value === cur ? " on" : ""}`, role: "option", "aria-selected": String(o.value === cur), "aria-label": o.label, "data-tip": o.label, "data-value": o.value, onclick: () => { c.set(o.value); changed(); } });
          b.innerHTML = svg;
          return b;
        }));
        return h("div", { class: "tp-pattern-box" }, h("div", { class: "tp-heading" }, c.label), grid);
      }
      case "custom":
        return c.render();
    }
  }
}
