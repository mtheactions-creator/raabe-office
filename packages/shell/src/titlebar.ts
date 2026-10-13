// Title bar: app icon, Quick Access Toolbar, document tabs (D-010), the "Suchen" box
// (finds and runs commands, Alt+M) and the window buttons. The empty areas drag the window.

import { h, icon } from "./dom";
import type { Commands } from "./commands";
import type { Platform } from "./platform";
import { showMenu, type MenuEntry } from "./popup";

export interface DocTab {
  id: string;
  name: string;
  dirty: boolean;
}

export interface TitleBarOptions {
  appIcon: string;
  qat: { cmd: string; key: string; menu?: MenuEntry[] }[];
  onSelectDoc: (id: string) => void;
  onCloseDoc: (id: string) => void;
  onNewDoc: () => void;
  /** Called when the search box is left, to give focus back to the document. */
  onSearchDone: () => void;
}

function fold(s: string): string {
  return s.toLowerCase().normalize("NFD").replace(/[̀-ͯ]/g, "").replace(/ß/g, "ss");
}

export class TitleBar {
  readonly el: HTMLElement;
  private tabsEl: HTMLElement;
  private maxBtn: HTMLButtonElement;
  private searchInput: HTMLInputElement;
  private results: HTMLElement | null = null;
  private resultIndex = 0;

  constructor(private commands: Commands, private platform: Platform, private opts: TitleBarOptions) {
    const qat = h("div", { class: "qat", role: "toolbar", "aria-label": "Symbolleiste für den Schnellzugriff" });
    for (const q of opts.qat) {
      const c = commands.get(q.cmd);
      const b = h("button", {
        "aria-label": c.label, "data-tip": commands.tooltip(q.cmd), "data-keytip": q.key, "data-keytip-scope": "root",
        onclick: () => void commands.run(q.cmd),
      }, icon(c.icon));
      commands.onRefresh(() => { b.disabled = !commands.isEnabled(q.cmd); });
      if (q.menu) {
        const more = h("button", { class: "more", "aria-label": "Weitere Optionen" }, icon("chevron-down"));
        const menu = q.menu;
        more.onclick = () => showMenu(commands, menu, more.getBoundingClientRect());
        commands.onRefresh(() => { more.disabled = !commands.isEnabled(q.cmd); });
        qat.append(h("span", { class: "split" }, b, more));
      } else {
        qat.append(b);
      }
    }

    this.tabsEl = h("div", { class: "doctabs", role: "tablist", "aria-label": "Geöffnete Dokumente", "data-tauri-drag-region": "" });

    this.searchInput = h("input", {
      type: "search", placeholder: "Suchen (Alt+M)", "aria-label": "Suchen",
      oninput: () => this.showResults(),
      onfocus: () => this.showResults(),
      onblur: () => setTimeout(() => this.hideResults(), 150),
      onkeydown: (e: KeyboardEvent) => this.onSearchKey(e),
    });
    const search = h("div", { class: "search", "data-keytip": "M", "data-keytip-scope": "root", onclick: () => this.searchInput.focus() },
      icon("search"), this.searchInput);

    this.maxBtn = h("button", { "aria-label": "Maximieren", onclick: () => platform.window.toggleMaximize() }, icon("maximize"));
    const win = h("div", { class: "winbtns" },
      h("button", { "aria-label": "Minimieren", onclick: () => platform.window.minimize() }, icon("minimize")),
      this.maxBtn,
      h("button", { class: "close", "aria-label": "Schließen", onclick: () => platform.window.close() }, icon("close")));
    platform.window.onResized(() => void this.updateMaximized());
    void this.updateMaximized();

    this.el = h("div", { class: "titlebar", "data-tauri-drag-region": "" },
      h("img", { class: "app-icon", src: opts.appIcon, alt: "", draggable: false }),
      qat, this.tabsEl, search, h("div", { class: "titlebar-spacer", "data-tauri-drag-region": "" }), win);
  }

  focusSearch(): void {
    this.searchInput.focus();
    this.searchInput.select();
  }

  setDocs(docs: DocTab[], activeId: string): void {
    this.tabsEl.replaceChildren();
    for (const d of docs) {
      const close = h("button", {
        class: "x", "aria-label": `${d.name} schließen`, "data-tip": "Schließen",
        onclick: (e: MouseEvent) => { e.stopPropagation(); this.opts.onCloseDoc(d.id); },
      }, icon("close"));
      const tab = h("div", {
        class: `doctab${d.id === activeId ? " active" : ""}`, role: "tab", "aria-selected": String(d.id === activeId),
        "data-tip": d.name, onmousedown: (e: MouseEvent) => {
          if (e.button === 1) { e.preventDefault(); this.opts.onCloseDoc(d.id); }
        },
        onclick: () => this.opts.onSelectDoc(d.id),
      }, h("span", { class: "name" }, d.name), d.dirty ? h("span", { class: "dirty", "aria-label": "Nicht gespeichert" }) : null, close);
      this.tabsEl.append(tab);
    }
    this.tabsEl.append(h("button", { class: "doctab-new", "aria-label": "Neues Dokument", "data-tip": "Neues Dokument (Strg+N)", onclick: () => this.opts.onNewDoc() }, "+"));
  }

  private async updateMaximized(): Promise<void> {
    const max = await this.platform.window.isMaximized();
    this.maxBtn.replaceChildren(icon(max ? "restore" : "maximize"));
    this.maxBtn.setAttribute("aria-label", max ? "Verkleinern" : "Maximieren");
    document.body.style.borderWidth = max ? "0" : "";
  }

  // ----- search -------------------------------------------------------------------------

  private matches(): string[] {
    const q = fold(this.searchInput.value.trim());
    if (!q) return [];
    return this.commands.all()
      .filter((c) => !c.hidden && [c.label, c.tip ?? "", ...(c.keywords ?? [])].some((t) => fold(t).includes(q)))
      .sort((a, b) => Number(this.commands.isEnabled(b.id)) - Number(this.commands.isEnabled(a.id)) ||
        Number(fold(b.label).startsWith(q)) - Number(fold(a.label).startsWith(q)))
      .slice(0, 12)
      .map((c) => c.id);
  }

  private showResults(): void {
    this.hideResults();
    if (!this.searchInput.value.trim()) return;
    const ids = this.matches();
    this.resultIndex = 0;
    const box = h("div", { class: "search-results", role: "listbox" }, h("div", { class: "hdr" }, "Aktionen"));
    if (!ids.length) box.append(h("div", { class: "empty" }, "Keine Ergebnisse"));
    ids.forEach((id, i) => {
      const c = this.commands.get(id);
      box.append(h("button", {
        class: i === 0 ? "active" : "", disabled: !this.commands.isEnabled(id), "data-id": id,
        onmousedown: (e: MouseEvent) => { e.preventDefault(); this.runResult(id); },
      }, icon(c.icon), c.label, this.commands.isEnabled(id) ? null : h("span", { class: "muted", style: "margin-left:auto;color:#9e9e9e" }, "noch nicht verfügbar")));
    });
    this.searchInput.parentElement!.append(box);
    this.results = box;
  }

  private hideResults(): void {
    this.results?.remove();
    this.results = null;
  }

  private runResult(id: string): void {
    this.searchInput.value = "";
    this.hideResults();
    this.opts.onSearchDone();
    void this.commands.run(id);
  }

  private onSearchKey(e: KeyboardEvent): void {
    const buttons = this.results ? Array.from(this.results.querySelectorAll<HTMLButtonElement>("button")) : [];
    if (e.key === "Escape") {
      e.preventDefault();
      this.searchInput.value = "";
      this.hideResults();
      this.opts.onSearchDone();
    } else if (e.key === "ArrowDown" || e.key === "ArrowUp") {
      e.preventDefault();
      if (!buttons.length) return;
      this.resultIndex = (this.resultIndex + (e.key === "ArrowDown" ? 1 : buttons.length - 1)) % buttons.length;
      buttons.forEach((b, i) => b.classList.toggle("active", i === this.resultIndex));
    } else if (e.key === "Enter") {
      e.preventDefault();
      const b = buttons[this.resultIndex];
      if (b && !b.disabled) this.runResult(b.dataset.id!);
    }
  }
}
