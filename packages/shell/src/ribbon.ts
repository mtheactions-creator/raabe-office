// The ribbon: tab row plus the ribbon card. Tabs and groups are plain data (see the app's
// ribbon definition); this module renders them like Word, shrinks groups in Word's order
// when the window gets narrower, and exposes keytips (Alt shortcuts) for keytips.ts.

import { h, icon } from "./dom";
import type { Commands } from "./commands";
import { closePopup, showMenu, showPopup, type MenuEntry } from "./popup";
import { highlightHex, showFontColors, showHighlights } from "./palette";

export type Size = "large" | "medium" | "small";

export type Control =
  | { type: "button"; cmd: string; size?: Size; key?: string; label?: string }
  | { type: "split"; cmd: string; size?: Size; key?: string; label?: string; menu: MenuEntry[] }
  | { type: "dropdown"; cmd: string; size?: Size; key?: string; label?: string; menu: MenuEntry[] }
  | {
      type: "combo"; cmd: string; width: number; key?: string; value: () => string;
      /** Entries of the dropdown list, grouped under headings. */
      options?: () => { heading?: string; items: string[] }[];
      /** CSS for an entry (e.g. the font itself for "Schriftart"). */
      optionStyle?: (o: string) => string;
      commit?: (v: string) => void;
    }
  | { type: "gallery"; key?: string; items: { cmd: string; preview: string }[]; visible: number }
  | {
      type: "spinner"; cmd: string; label: string; icon?: string; key?: string;
      value?: () => string; commit?: (v: string) => void; step?: (dir: 1 | -1) => void;
    }
  | { type: "color"; cmd: string; key?: string; palette: "font" | "highlight"; current: () => string; pick: (c: string) => void }
  | { type: "check"; cmd: string; key?: string }
  | { type: "rows"; rows: (Control | "gap")[][] }
  | { type: "stack"; items: Control[] };

export interface Group {
  id: string;
  label: string;
  /** Keytip of the collapsed group button, e.g. "ZF". */
  key?: string;
  icon: string;
  launcher?: { cmd: string; key?: string };
  controls: Control[];
}

export interface Tab {
  id: string;
  label: string;
  key: string;
  groups: Group[];
  /** Group ids in the order Word collapses them as the window narrows. */
  shrink: string[];
}

type Updater = () => void;

/** Office's dialog launcher: a corner with a diagonal arrow (drawn here; no Fluent equivalent). */
const LAUNCHER_SVG = '<svg viewBox="0 0 10 10" width="10" height="10" aria-hidden="true"><path d="M1 1v8h8" fill="none" ' +
  'stroke="currentColor" stroke-width="1"/><path d="M3.5 3.5l4.5 4.5M8 4.5V8H4.5" fill="none" stroke="currentColor" stroke-width="1"/></svg>';

export interface RibbonOptions {
  onFile: () => void;
  /** Called after a ribbon input (font box, spinner …) is done, to give focus back to the document. */
  onDone?: () => void;
  /** Extra controls at the right end of the tab row (e.g. "Kommentare"). */
  tabRowRight?: HTMLElement[];
}

export class Ribbon {
  readonly tabRow: HTMLElement;
  readonly card: HTMLElement;
  private current: Tab;
  private updaters: Updater[] = [];
  private tabButtons = new Map<string, HTMLElement>();
  private naturalWidths = new Map<string, number[]>();
  private collapsedGroups = new Set<string>();
  private minimized = false;

  constructor(private commands: Commands, private tabs: Tab[], private opts: RibbonOptions) {
    this.current = tabs[0];
    const file = h("button", {
      class: "tab", "data-keytip": "D", "data-keytip-scope": "root", "data-tip": "Datei\nDateien öffnen, speichern, drucken und mehr.",
      onclick: () => opts.onFile(),
    }, "Datei");
    this.tabRow = h("div", { class: "tabrow", role: "tablist" }, file);
    for (const t of tabs) {
      const b = h("button", {
        class: "tab", role: "tab", "data-keytip": t.key, "data-keytip-scope": "root",
        onclick: () => this.select(t.id),
        ondblclick: () => this.toggleMinimized(),
      }, t.label);
      this.tabButtons.set(t.id, b);
      this.tabRow.append(b);
    }
    this.tabRow.append(h("div", { class: "tabrow-right" }, ...(opts.tabRowRight ?? [])));
    this.card = h("div", { class: "ribbon-card" });
    commands.onRefresh(() => this.refresh());
    new ResizeObserver(() => this.fit()).observe(this.card);
    this.select(tabs[0].id);
  }

  get activeTab(): string {
    return this.current.id;
  }

  select(id: string): void {
    const tab = this.tabs.find((t) => t.id === id);
    if (!tab) return;
    this.current = tab;
    for (const [tid, b] of this.tabButtons) {
      b.classList.toggle("active", tid === id);
      b.setAttribute("aria-selected", String(tid === id));
    }
    if (this.minimized) {
      // "Nur Registerkarten anzeigen": clicking a tab shows the ribbon until the next click.
      this.card.classList.remove("hidden");
      const away = (e: MouseEvent) => {
        if (this.card.contains(e.target as Node) || this.tabRow.contains(e.target as Node)) return;
        if (this.minimized) this.card.classList.add("hidden");
        removeEventListener("mousedown", away, true);
      };
      setTimeout(() => addEventListener("mousedown", away, true));
    }
    this.collapsedGroups.clear();
    this.render();
    this.fit();
  }

  toggleMinimized(): void {
    this.minimized = !this.minimized;
    this.card.classList.toggle("hidden", this.minimized);
    if (!this.minimized) this.fit();
  }

  get isMinimized(): boolean {
    return this.minimized;
  }

  refresh(): void {
    for (const u of this.updaters) u();
  }

  // ----- rendering ----------------------------------------------------------------------

  private render(): void {
    closePopup();
    this.updaters = [];
    this.card.replaceChildren();
    for (const g of this.current.groups) {
      this.card.append(this.collapsedGroups.has(g.id) ? this.renderCollapsed(g) : this.renderGroup(g));
    }
    this.card.append(h("button", {
      class: "ribbon-collapse", "data-tip": "Menüband reduzieren (Strg+F1)", onclick: () => this.toggleMinimized(),
    }, icon("chevron-up")));
    this.refresh();
  }

  private renderGroup(g: Group): HTMLElement {
    const content = h("div", { class: "content" }, ...g.controls.map((c) => this.renderControl(c)));
    const label = h("div", { class: "label" }, g.label);
    if (g.launcher) {
      const cmd = g.launcher.cmd;
      const b = h("button", {
        class: "launcher", "data-tip": this.commands.tooltip(cmd), "aria-label": this.commands.get(cmd).label,
        "data-keytip": g.launcher.key, "data-keytip-scope": this.current.id,
        onclick: () => void this.commands.run(cmd),
      });
      b.innerHTML = LAUNCHER_SVG;
      this.updaters.push(() => { b.disabled = !this.commands.isEnabled(cmd); });
      label.append(b);
    }
    return h("div", { class: "rgroup", role: "group", "aria-label": g.label, "data-group": g.id }, content, label);
  }

  private renderCollapsed(g: Group): HTMLElement {
    const btn = h("button", {
      class: "rb rb-large", style: "height:92px", "data-keytip": g.key, "data-keytip-scope": this.current.id,
      "data-tip": g.label,
    }, icon(g.icon, "l"), h("span", { class: "txt" }, g.label), h("span", { class: "chev" }, icon("chevron-down")));
    btn.onclick = () => {
      const popup = h("div", { class: "group-popup" }, this.renderGroup(g));
      showPopup(popup, btn.getBoundingClientRect(), () => this.refresh());
      this.refresh();
    };
    return h("div", { class: "rgroup collapsed", "data-group": g.id }, h("div", { class: "content" }, btn));
  }

  private scope(): string {
    return this.current.id;
  }

  private bind(el: HTMLElement & { disabled?: boolean }, cmd: string): void {
    const c = this.commands;
    this.updaters.push(() => {
      const enabled = c.isEnabled(cmd);
      if ("disabled" in el) el.disabled = !enabled;
      el.classList.toggle("disabled", !enabled);
      el.classList.toggle("checked", c.isChecked(cmd));
      el.setAttribute("aria-pressed", String(c.isChecked(cmd)));
    });
  }

  private button(cmd: string, size: Size, key?: string, label?: string, chevron = false): HTMLButtonElement {
    const c = this.commands.get(cmd);
    const text = label ?? c.label;
    const parts: Node[] = [icon(c.icon, size === "large" ? "l" : "s")];
    if (size !== "small") parts.push(h("span", { class: "txt" }, text));
    if (chevron) parts.push(h("span", { class: "chev" }, icon("chevron-down")));
    const b = h("button", {
      class: `rb rb-${size}`, "data-tip": this.commands.tooltip(cmd), "aria-label": c.label,
      "data-keytip": key, "data-keytip-scope": this.scope(),
    }, ...parts);
    this.bind(b, cmd);
    return b;
  }

  private renderControl(c: Control | "gap"): HTMLElement {
    if (c === "gap") return h("span", { class: "gap" });
    switch (c.type) {
      case "button": {
        const b = this.button(c.cmd, c.size ?? "small", c.key, c.label);
        b.onclick = () => void this.commands.run(c.cmd);
        return b;
      }
      case "dropdown": {
        const b = this.button(c.cmd, c.size ?? "small", c.key, c.label, true);
        // A dropdown's command is enabled when anything in its menu is.
        this.updaters.push(() => {
          const any = c.menu.some((e) => e !== "-" && ("cmd" in e ? this.commands.isEnabled(e.cmd) : !e.disabled));
          b.disabled = !any;
          b.classList.toggle("disabled", !any);
          b.dataset.tip = this.commands.tooltip(c.cmd, any);
        });
        b.onclick = () => showMenu(this.commands, c.menu, b.getBoundingClientRect());
        return b;
      }
      case "split": {
        const size = c.size ?? "small";
        const main = this.button(c.cmd, size, undefined, c.label);
        main.onclick = () => void this.commands.run(c.cmd);
        if (size === "large") main.querySelector(".txt")?.remove();
        const arrowParts: Node[] = [];
        if (size === "large") arrowParts.push(h("span", { class: "txt" }, c.label ?? this.commands.get(c.cmd).label));
        arrowParts.push(h("span", { class: "chev" }, icon("chevron-down")));
        const arrow = h("button", {
          class: `rb rb-${size === "large" ? "large" : "small"} arrow`, "aria-label": "Weitere Optionen",
          "data-keytip": c.key, "data-keytip-scope": this.scope(),
        }, ...arrowParts);
        this.updaters.push(() => {
          const any = c.menu.some((e) => e !== "-" && ("cmd" in e ? this.commands.isEnabled(e.cmd) : !e.disabled));
          arrow.disabled = !any && !this.commands.isEnabled(c.cmd);
        });
        arrow.onclick = () => showMenu(this.commands, c.menu, arrow.getBoundingClientRect());
        return h("div", { class: `rsplit${size === "large" ? " rsplit-large" : ""}` }, main, arrow);
      }
      case "combo": {
        const c2 = c;
        const input = h("input", { type: "text", "aria-label": this.commands.get(c.cmd).label, spellcheck: false });
        const done = () => this.opts.onDone?.();
        const commit = () => { if (c2.commit && input.value.trim()) c2.commit(input.value.trim()); done(); };
        input.addEventListener("keydown", (e) => {
          if (e.key === "Enter") { e.preventDefault(); commit(); }
          else if (e.key === "Escape") { e.preventDefault(); input.value = c2.value(); done(); }
          e.stopPropagation();
        });
        input.addEventListener("focus", () => input.select());
        const openList = () => {
          const groups = c2.options?.() ?? [];
          const list = h("div", { class: "menu combo-list", style: `min-width:${Math.max(c2.width, 180)}px` });
          for (const g of groups) {
            if (g.heading) list.append(h("div", { class: "palette-title" }, g.heading));
            for (const item of g.items) {
              list.append(h("button", {
                style: c2.optionStyle?.(item) ?? "",
                onclick: () => { closePopup(); input.value = item; commit(); },
              }, item));
            }
          }
          showPopup(list, wrap.getBoundingClientRect());
          list.querySelector<HTMLElement>("button")?.scrollIntoView({ block: "nearest" });
        };
        const arrow = h("button", { tabIndex: -1, "aria-label": "Öffnen", onclick: () => openList() }, icon("chevron-down"));
        const wrap = h("div", {
          class: "rcombo", style: `width:${c.width}px`, "data-tip": this.commands.tooltip(c.cmd),
          "data-keytip": c.key, "data-keytip-scope": this.scope(),
          onclick: (e: MouseEvent) => { if (e.target === wrap) input.focus(); },
        }, input, arrow);
        this.updaters.push(() => {
          if (document.activeElement !== input) input.value = c2.value();
          const on = this.commands.isEnabled(c2.cmd);
          wrap.classList.toggle("disabled", !on);
          input.readOnly = !on || !c2.commit;
          arrow.disabled = !on || !c2.options;
        });
        return wrap;
      }
      case "gallery": {
        const items = h("div", { class: "items" });
        for (const it of c.items.slice(0, c.visible)) {
          const b = h("button", {
            class: "item", "data-tip": this.commands.tooltip(it.cmd), style: "",
            onclick: () => void this.commands.run(it.cmd),
          }, h("span", { style: it.preview }, this.commands.get(it.cmd).label));
          this.bind(b, it.cmd);
          items.append(b);
        }
        const more = h("button", { "aria-label": "Weitere", "data-keytip": c.key, "data-keytip-scope": this.scope() }, icon("chevron-down"));
        more.onclick = () => showMenu(this.commands, c.items.map((i) => ({ cmd: i.cmd })), more.getBoundingClientRect());
        return h("div", { class: "rgallery", role: "listbox" }, items,
          h("div", { class: "scroll" }, h("button", { disabled: true }, icon("chevron-up")), more));
      }
      case "spinner": {
        const c2 = c;
        const input = h("input", { type: "text", "aria-label": c.label, spellcheck: false });
        const done = () => this.opts.onDone?.();
        input.addEventListener("keydown", (e) => {
          if (e.key === "Enter") { e.preventDefault(); c2.commit?.(input.value); done(); }
          else if (e.key === "Escape") { e.preventDefault(); done(); }
          else if (e.key === "ArrowUp" || e.key === "ArrowDown") { e.preventDefault(); c2.step?.(e.key === "ArrowUp" ? 1 : -1); }
          e.stopPropagation();
        });
        input.addEventListener("focus", () => input.select());
        const arrows = h("span", { class: "spin" },
          h("button", { tabIndex: -1, "aria-label": "Mehr", onclick: () => c2.step?.(1) }, icon("chevron-up")),
          h("button", { tabIndex: -1, "aria-label": "Weniger", onclick: () => c2.step?.(-1) }, icon("chevron-down")));
        const el = h("div", { class: "rspinner", "data-tip": this.commands.tooltip(c.cmd), "data-keytip": c.key, "data-keytip-scope": this.scope() },
          icon(c.icon), h("span", { class: "lbl" }, c.label), h("span", { class: "spinbox" }, input, arrows));
        this.updaters.push(() => {
          const on = this.commands.isEnabled(c2.cmd);
          input.disabled = !on;
          arrows.querySelectorAll("button").forEach((b) => { b.disabled = !on; });
          if (document.activeElement !== input && c2.value) input.value = c2.value();
        });
        return el;
      }
      case "color": {
        const c2 = c;
        const cmdDef = this.commands.get(c.cmd);
        const bar = h("span", { class: "colorbar" });
        const main = h("button", {
          class: "rb rb-small color-main", "data-tip": this.commands.tooltip(c.cmd), "aria-label": cmdDef.label,
          onclick: () => { if (this.commands.isEnabled(c2.cmd)) { c2.pick(c2.current()); this.opts.onDone?.(); } },
        }, icon(cmdDef.icon), bar);
        const arrow = h("button", {
          class: "rb rb-small arrow", "aria-label": "Weitere Optionen", "data-keytip": c.key, "data-keytip-scope": this.scope(),
          onclick: () => {
            const pick = (v: string) => { c2.pick(v); this.opts.onDone?.(); };
            if (c2.palette === "font") showFontColors(arrow.getBoundingClientRect(), pick);
            else showHighlights(arrow.getBoundingClientRect(), pick);
          },
        }, h("span", { class: "chev" }, icon("chevron-down")));
        this.updaters.push(() => {
          const on = this.commands.isEnabled(c2.cmd);
          main.disabled = !on; arrow.disabled = !on;
          const cur = c2.current();
          const hex = c2.palette === "highlight" ? highlightHex(cur) : cur === "auto" ? "000000" : cur;
          bar.style.background = cur === "none" ? "transparent" : `#${hex}`;
        });
        return h("div", { class: "rsplit" }, main, arrow);
      }
      case "check": {
        const box = h("span", { class: "box" });
        const el = h("button", {
          class: "rcheck", "data-tip": this.commands.tooltip(c.cmd), "data-keytip": c.key, "data-keytip-scope": this.scope(),
          onclick: () => void this.commands.run(c.cmd),
        }, box, this.commands.get(c.cmd).label);
        this.bind(el, c.cmd);
        this.updaters.push(() => { box.replaceChildren(this.commands.isChecked(c.cmd) ? icon("check") : ""); });
        return el;
      }
      case "rows":
        return h("div", { class: "rrows" }, ...c.rows.map((row) => h("div", { class: "rrow" }, ...row.map((x) => this.renderControl(x)))));
      case "stack":
        return h("div", { class: "rstack" }, ...c.items.map((x) => this.renderControl(x)));
    }
  }

  // ----- shrinking ----------------------------------------------------------------------

  /** Collapses groups in Word's order until the tab fits the window. */
  private fit(): void {
    if (this.minimized || !this.card.isConnected) return;
    const tab = this.current;
    let widths = this.naturalWidths.get(tab.id);
    if (!widths) {
      if (this.collapsedGroups.size) { this.collapsedGroups.clear(); this.render(); }
      const groups = Array.from(this.card.querySelectorAll<HTMLElement>(":scope > .rgroup"));
      if (!groups.length || groups[0].offsetWidth === 0) return;
      widths = groups.map((g) => g.offsetWidth);
      this.naturalWidths.set(tab.id, widths);
    }
    const available = this.card.clientWidth - 34;
    const collapsedWidth = 72;
    const want = new Set<string>();
    let total = widths.reduce((a, b) => a + b, 0);
    for (const id of tab.shrink) {
      if (total <= available) break;
      const i = tab.groups.findIndex((g) => g.id === id);
      if (i < 0) continue;
      want.add(id);
      total -= widths[i] - collapsedWidth;
    }
    const same = want.size === this.collapsedGroups.size && [...want].every((x) => this.collapsedGroups.has(x));
    if (!same) {
      this.collapsedGroups = want;
      this.render();
    }
  }
}
