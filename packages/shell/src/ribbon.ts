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
  | { type: "split"; cmd: string; size?: Size; key?: string; label?: string; menu: MenuEntry[]; popup?: (close: () => void) => HTMLElement }
  | {
      type: "dropdown"; cmd: string; size?: Size; key?: string; label?: string; menu: MenuEntry[];
      /** Own content instead of a menu (the table grid, galleries …); `close` closes it. */
      popup?: (close: () => void) => HTMLElement;
    }
  | {
      type: "combo"; cmd: string; width: number; key?: string; value: () => string;
      /** Entries of the dropdown list, grouped under headings. */
      options?: () => { heading?: string; items: string[] }[];
      /** CSS for an entry (e.g. the font itself for "Schriftart"). */
      optionStyle?: (o: string) => string;
      commit?: (v: string) => void;
    }
  | {
      type: "gallery"; key?: string; items: { cmd: string; preview: string }[];
      /** Items shown at least; with `fill` the gallery grows into free width, as Word's does. */
      visible: number; fill?: boolean;
    }
  | {
      type: "spinner"; cmd: string; label: string; icon?: string; key?: string;
      value?: () => string; commit?: (v: string) => void; step?: (dir: 1 | -1) => void;
    }
  | { type: "color"; cmd: string; key?: string; palette: "font" | "highlight"; current: () => string; pick: (c: string) => void }
  | { type: "check"; cmd: string; key?: string }
  | { type: "rows"; rows: (Control | "gap")[][] }
  /** A column of controls; `title` is a small heading above it (Layout > Absatz: "Einzug"). */
  | { type: "stack"; items: Control[]; title?: string };

export interface Group {
  id: string;
  label: string;
  /** Keytip of the collapsed group button, e.g. "ZF". */
  key?: string;
  icon: string;
  launcher?: { cmd: string; key?: string };
  /** The largest layout (Word's at a wide window). */
  controls: Control[];
  /** Word's smaller layouts of the group, in the order it switches to them as the window narrows. */
  smaller?: Control[][];
}

export interface Tab {
  id: string;
  label: string;
  key: string;
  groups: Group[];
  /**
   * Word's reduction sequence as the window narrows: each entry steps that group down one
   * layout (`smaller`), and after its smallest one collapses it to a single button.
   */
  shrink: string[];
  /** Shown only in a context (table, header, picture …), after the other tabs. */
  contextual?: boolean;
}

type Updater = () => void;

/** Office's small chevron next to and below ribbon controls (8 x 5 px at 100 %, measured). */
const CHEVRON_SVG = '<svg viewBox="0 0 8 5" width="8" height="5" aria-hidden="true"><path d="M.75.75 4 4.25 7.25.75" fill="none" ' +
  'stroke="currentColor" stroke-width="1" stroke-linecap="round" stroke-linejoin="round"/></svg>';

function chev(): HTMLElement {
  const span = h("span", { class: "chev" });
  span.innerHTML = CHEVRON_SVG;
  return span;
}

/**
 * Word's large-button labels: two words go on two lines, split where the lines are most even;
 * a long hyphenated word breaks after the hyphen ("3D-" / "Modelle"). The chevron of a menu
 * button follows the second line, or stands on its own line under a one-line label.
 */
function largeLabel(text: string, withChevron: boolean): Node[] {
  let lines = text.includes("\n") ? text.split("\n") : [text];
  if (lines.length === 1) {
    const words = text.split(" ");
    if (words.length > 1) {
      let best = 1, score = Infinity;
      for (let i = 1; i < words.length; i++) {
        const a = words.slice(0, i).join(" ").length, b = words.slice(i).join(" ").length;
        // On a tie Word puts more on the first line ("Ebene nach" / "vorne").
        if (Math.max(a, b) <= score) { score = Math.max(a, b); best = i; }
      }
      lines = [words.slice(0, best).join(" "), words.slice(best).join(" ")];
    } else if (text.length > 8 && text.indexOf("-") > 0 && text.indexOf("-") < text.length - 1) {
      const i = text.indexOf("-") + 1;
      lines = [text.slice(0, i), text.slice(i)];
    }
  }
  const txt = h("span", { class: "txt" }, lines.join("\n"));
  if (!withChevron) return [txt];
  if (lines.length > 1) { txt.append(chev()); return [txt]; }
  return [txt, chev()];
}

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
  /** Layout per group of the current tab: 0 = largest, smaller ones after, then collapsed. */
  private levels = new Map<string, number>();
  /** Measured group widths: "tab|group|level" → px. */
  private widths = new Map<string, number>();
  /** Free width the style gallery may grow into. */
  private galleryRoom = 0;
  private minimized = false;
  private contextual: Tab[] = [];
  private contextStart: HTMLElement;

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
    this.contextStart = h("span", { class: "ctx-start" });
    this.tabRow.append(this.contextStart, h("div", { class: "tabrow-right" }, ...(opts.tabRowRight ?? [])));
    this.card = h("div", { class: "ribbon-card" });
    commands.onRefresh(() => this.refresh());
    new ResizeObserver(() => this.fit()).observe(this.card);
    this.select(tabs[0].id);
  }

  get activeTab(): string {
    return this.current.id;
  }

  /**
   * Shows the contextual tabs for where the caret is (Tabellenentwurf, Kopf- und Fußzeile …).
   * Tabs that disappear while selected hand back to Start.
   */
  setContextual(tabs: Tab[]): void {
    const key = tabs.map((t) => t.id).join(",");
    if (key === this.contextual.map((t) => t.id).join(",")) {
      // Same tabs: refresh their definitions (values) without re-rendering.
      this.contextual = tabs;
      return;
    }
    for (const t of this.contextual) { this.tabButtons.get(t.id)?.remove(); this.tabButtons.delete(t.id); }
    this.contextual = tabs;
    for (const t of tabs) {
      const b = h("button", {
        class: "tab contextual", role: "tab", "data-keytip": t.key, "data-keytip-scope": "root",
        onclick: () => this.select(t.id),
        ondblclick: () => this.toggleMinimized(),
      }, t.label);
      this.tabButtons.set(t.id, b);
      this.tabRow.insertBefore(b, this.contextStart.nextSibling);
    }
    // Keep their order: insertBefore above reverses, so put them back in sequence.
    let at: ChildNode = this.contextStart;
    for (const t of tabs) { const b = this.tabButtons.get(t.id)!; at.after(b); at = b; }
    if (this.current.contextual && !tabs.some((t) => t.id === this.current.id)) this.select(this.tabs[0].id);
  }

  select(id: string): void {
    const tab = this.tabs.find((t) => t.id === id) ?? this.contextual.find((t) => t.id === id);
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
    this.levels = new Map();
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
      const layouts = this.layoutsOf(g);
      const level = this.levels.get(g.id) ?? 0;
      this.card.append(level >= layouts.length ? this.renderCollapsed(g) : this.renderGroup(g, layouts[level]));
    }
    this.card.append(h("button", {
      class: "ribbon-collapse", "data-tip": "Menüband reduzieren (Strg+F1)", onclick: () => this.toggleMinimized(),
    }, icon("chevron-up")));
    this.refresh();
  }

  /**
   * Controls outside the ribbon (Word's mini toolbar), bound to the commands like the
   * ribbon's own; `refresh` updates their state.
   */
  toolbar(controls: Control[]): { el: HTMLElement; refresh: () => void } {
    const saved = this.updaters;
    this.updaters = [];
    const el = h("div", { class: "mini-toolbar", role: "toolbar" }, ...controls.map((c) => this.renderControl(c)));
    const own = this.updaters;
    this.updaters = saved;
    const refresh = () => { for (const u of own) u(); };
    refresh();
    return { el, refresh };
  }

  private layoutsOf(g: Group): Control[][] {
    return [g.controls, ...(g.smaller ?? [])];
  }

  private renderGroup(g: Group, controls: Control[] = g.controls): HTMLElement {
    const content = h("div", { class: "content" }, ...controls.map((c) => this.renderControl(c)));
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
    }, icon(g.icon, "l"), ...largeLabel(g.label, true));
    btn.onclick = () => {
      // Word's popup shows the group in its smallest layout.
      const layouts = this.layoutsOf(g);
      const popup = h("div", { class: "group-popup" }, this.renderGroup(g, layouts[layouts.length - 1]));
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
    if (size === "large") parts.push(...largeLabel(text, chevron));
    else {
      if (size !== "small") parts.push(h("span", { class: "txt" }, text));
      if (chevron) parts.push(chev());
    }
    const b = h("button", {
      class: `rb rb-${size}${chevron ? " has-chev" : ""}`, "data-tip": this.commands.tooltip(cmd), "aria-label": c.label,
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
          const any = c.popup ? this.commands.isEnabled(c.cmd) : c.menu.some((e) => e !== "-" && ("cmd" in e ? this.commands.isEnabled(e.cmd) : "run" in e && !e.disabled));
          b.disabled = !any;
          b.classList.toggle("disabled", !any);
          b.dataset.tip = this.commands.tooltip(c.cmd, any);
        });
        b.onclick = () => {
          if (c.popup) showPopup(c.popup(closePopup), b.getBoundingClientRect(), () => this.opts.onDone?.());
          else showMenu(this.commands, c.menu, b.getBoundingClientRect());
        };
        return b;
      }
      case "split": {
        const size = c.size ?? "small";
        const main = this.button(c.cmd, size, undefined, c.label);
        main.onclick = () => void this.commands.run(c.cmd);
        if (size === "large") main.querySelector(".txt")?.remove();
        const arrowParts: Node[] = size === "large" ? largeLabel(c.label ?? this.commands.get(c.cmd).label, true) : [chev()];
        const arrow = h("button", {
          class: `rb rb-${size === "large" ? "large" : "small"} arrow`, "aria-label": "Weitere Optionen",
          "data-keytip": c.key, "data-keytip-scope": this.scope(),
        }, ...arrowParts);
        this.updaters.push(() => {
          const any = c.menu.some((e) => e !== "-" && ("cmd" in e ? this.commands.isEnabled(e.cmd) : "run" in e && !e.disabled));
          arrow.disabled = !any && !this.commands.isEnabled(c.cmd);
        });
        arrow.onclick = () => {
          if (c.popup) showPopup(c.popup(closePopup), arrow.getBoundingClientRect(), () => this.opts.onDone?.());
          else showMenu(this.commands, c.menu, arrow.getBoundingClientRect());
        };
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
        const el = h("div", { class: `rgallery${c.fill ? " fill" : ""}`, role: "listbox", "data-min": String(c.visible) }, items,
          h("div", { class: "scroll" }, h("button", { disabled: true }, icon("chevron-up")), more));
        if (c.fill) {
          // All items are there; fitGallery() shows as many as the free width allows.
          for (const it of c.items.slice(c.visible)) {
            const b = h("button", {
              class: "item extra", "data-tip": this.commands.tooltip(it.cmd), style: "display:none",
              onclick: () => void this.commands.run(it.cmd),
            }, h("span", { style: it.preview }, this.commands.get(it.cmd).label));
            this.bind(b, it.cmd);
            items.append(b);
          }
        }
        return el;
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
        }, chev());
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
        // Two rows (Word's wide Schriftart and Absatz) stand apart; three are packed.
        return h("div", { class: `rrows${c.rows.length === 2 ? " rrows-2" : ""}` },
          ...c.rows.map((row) => h("div", { class: "rrow" }, ...row.map((x) => this.renderControl(x)))));
      case "stack":
        return h("div", { class: `rstack${c.title ? " titled" : ""}` }, ...(c.title ? [h("div", { class: "rstack-title" }, c.title)] : []),
          ...c.items.map((x) => this.renderControl(x)));
    }
  }

  // ----- shrinking ----------------------------------------------------------------------

  /** Forget measured widths (touch mode and other size changes). */
  invalidate(): void {
    this.widths.clear();
    this.fit();
  }

  /** Group widths for these layouts; renders the groups once when a width is not known yet. */
  private measure(tab: Tab, levels: Map<string, number>): number[] {
    const key = (g: Group) => `${tab.id}|${g.id}|${levels.get(g.id) ?? 0}`;
    if (tab.groups.some((g) => !this.widths.has(key(g)))) {
      this.levels = new Map(levels);
      this.galleryRoom = 0;
      this.render();
      const els = Array.from(this.card.querySelectorAll<HTMLElement>(":scope > .rgroup"));
      if (els.length !== tab.groups.length || els[0].offsetWidth === 0) return [];
      tab.groups.forEach((g, i) => this.widths.set(key(g), els[i].offsetWidth));
    }
    return tab.groups.map((g) => this.widths.get(key(g)) ?? 0);
  }

  /** Steps groups down in Word's order until the tab fits the window. */
  private fit(): void {
    if (this.minimized || !this.card.isConnected) return;
    const tab = this.current;
    const available = this.card.clientWidth - 34;
    if (available <= 0) return;
    const levels = new Map(tab.groups.map((g) => [g.id, 0]));
    // Word's sequence, then everything still open from the right (a fallback for narrow windows).
    const steps = [...tab.shrink, ...tab.groups.map((g) => g.id).reverse(), ...tab.groups.map((g) => g.id).reverse()];
    let widths = this.measure(tab, levels);
    if (!widths.length) return;
    let total = widths.reduce((a, b) => a + b, 0);
    while (total > available && steps.length) {
      const id = steps.shift()!;
      const g = tab.groups.find((x) => x.id === id);
      if (!g) continue;
      const lv = levels.get(id)!;
      if (lv >= this.layoutsOf(g).length) continue;
      levels.set(id, lv + 1);
      widths = this.measure(tab, levels);
      if (!widths.length) return;
      total = widths.reduce((a, b) => a + b, 0);
    }
    const room = Math.max(0, available - total);
    const same = tab.groups.every((g) => (this.levels.get(g.id) ?? 0) === levels.get(g.id));
    if (!same) {
      this.levels = levels;
      this.render();
    }
    this.galleryRoom = room;
    this.fitGallery();
  }

  /** Word's style gallery shows as many styles as the free width allows. */
  private fitGallery(): void {
    for (const gal of Array.from(this.card.querySelectorAll<HTMLElement>(".rgallery.fill"))) {
      const items = Array.from(gal.querySelectorAll<HTMLElement>(".item"));
      const min = Number(gal.dataset.min) || 1;
      const itemW = items[0]?.offsetWidth || 108;
      const extra = Math.max(0, Math.floor(this.galleryRoom / itemW));
      items.forEach((it, i) => { it.style.display = i < min + extra ? "" : "none"; });
    }
  }
}

