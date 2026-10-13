// Backstage ("Datei"): a left navigation and one page at a time. Entries either show a page
// or run a command directly (e.g. "Speichern unter" opens the Windows dialog at once,
// without Word's extra location page).

import { h, icon } from "./dom";
import type { Commands } from "./commands";

export interface BackstageEntry {
  id: string;
  label: string;
  icon: string;
  /** Renders the page; omit for entries that run `cmd` and close the backstage. */
  page?: () => HTMLElement;
  cmd?: string;
  divider?: boolean;
  bottom?: boolean;
}

export class Backstage {
  readonly el: HTMLElement;
  private nav: HTMLElement;
  private body: HTMLElement;
  private current = "";
  private onCloseCb: () => void = () => {};

  constructor(private commands: Commands, private entries: BackstageEntry[]) {
    this.nav = h("nav", { "aria-label": "Datei" });
    this.body = h("div", { class: "page-body" });
    this.el = h("div", { class: "backstage", hidden: true, role: "dialog", "aria-label": "Datei" }, this.nav, this.body);
    addEventListener("keydown", (e) => {
      if (!this.isOpen) return;
      if (e.key === "Escape") { e.preventDefault(); e.stopPropagation(); this.close(); }
    }, true);
    commands.onRefresh(() => { if (this.isOpen) this.renderNav(); });
  }

  get isOpen(): boolean {
    return !this.el.hidden;
  }

  onClose(cb: () => void): void {
    this.onCloseCb = cb;
  }

  open(pageId = "home"): void {
    this.el.hidden = false;
    this.show(pageId);
  }

  close(): void {
    if (!this.isOpen) return;
    this.el.hidden = true;
    this.onCloseCb();
  }

  /** Re-renders the current page (e.g. after the recent-file list changed). */
  refreshPage(): void {
    if (this.isOpen) this.show(this.current);
  }

  private show(id: string): void {
    const entry = this.entries.find((e) => e.id === id);
    if (!entry?.page) return;
    this.current = id;
    this.renderNav();
    this.body.replaceChildren(entry.page());
    this.body.scrollTop = 0;
  }

  private renderNav(): void {
    this.nav.replaceChildren(h("button", { class: "back", "aria-label": "Zurück", "data-tip": "Zurück", onclick: () => this.close() }, icon("back")));
    for (const e of this.entries) {
      if (e.bottom && !this.nav.querySelector(".spacer")) this.nav.append(h("div", { class: "spacer" }), h("hr"));
      if (e.divider) this.nav.append(h("hr"));
      const disabled = e.cmd ? !this.commands.isEnabled(e.cmd) : false;
      this.nav.append(h("button", {
        class: `item${e.id === this.current ? " active" : ""}`, disabled,
        onclick: () => {
          if (e.page) this.show(e.id);
          else if (e.cmd) { this.close(); void this.commands.run(e.cmd); }
        },
      }, icon(e.icon), e.label));
    }
  }
}
