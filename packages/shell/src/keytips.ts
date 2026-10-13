// Keytips: press and release Alt (or F10), letters appear on tabs and the Quick Access
// Toolbar; typing a tab's letter shows the letters of that tab's commands, as in Office.
// Elements opt in with data-keytip="R" and data-keytip-scope="root" | <tab id>.
// AltGr (Ctrl+Alt on German keyboards, needed for @, €, {) never triggers keytips.

import { h } from "./dom";

function norm(key: string): string | null {
  if (key.length !== 1) return null;
  return key === "ß" ? "ß" : key.toUpperCase();
}

export class Keytips {
  private active = false;
  private scope = "root";
  private typed = "";
  private badges: HTMLElement[] = [];
  private altAlone = false;

  /** `scopeAfter(el)` tells which scope to continue in after activating `el` (a ribbon tab). */
  constructor(private scopeAfter: (el: HTMLElement) => string | null, private onExit: () => void) {
    addEventListener("keydown", (e) => this.onKeyDown(e), true);
    addEventListener("keyup", (e) => this.onKeyUp(e), true);
    addEventListener("mousedown", () => this.exit(), true);
    addEventListener("blur", () => this.exit());
    addEventListener("resize", () => this.exit());
  }

  get isActive(): boolean {
    return this.active;
  }

  private onKeyDown(e: KeyboardEvent): void {
    if (e.key === "Alt") {
      this.altAlone = !e.ctrlKey && !e.repeat;
      if (this.active) { e.preventDefault(); }
      return;
    }
    if (e.key === "Control" || e.key === "AltGraph" || e.ctrlKey) { this.altAlone = false; return; }
    if (e.key === "F10" && !e.shiftKey) {
      e.preventDefault();
      this.active ? this.exit() : this.enter();
      return;
    }
    if (!this.active) {
      // Alt+letter pressed together jumps straight to that tab (Alt+Umschalt+… are shortcuts,
      // such as Alt+Umschalt+0 for a new equation).
      if (e.altKey && !e.ctrlKey && !e.shiftKey && norm(e.key)) {
        this.altAlone = false;
        e.preventDefault();
        e.stopPropagation();
        this.enter();
        this.type(norm(e.key)!);
      } else {
        this.altAlone = false;
      }
      return;
    }
    e.preventDefault();
    e.stopPropagation();
    if (e.key === "Escape") {
      if (this.typed) { this.typed = ""; this.show(); }
      else if (this.scope !== "root") { this.scope = "root"; this.show(); }
      else this.exit();
      return;
    }
    const k = norm(e.key);
    if (k) this.type(k);
  }

  private onKeyUp(e: KeyboardEvent): void {
    if (e.key !== "Alt") return;
    if (this.altAlone) {
      e.preventDefault();
      this.active ? this.exit() : this.enter();
    }
    this.altAlone = false;
  }

  private enter(): void {
    this.active = true;
    this.scope = "root";
    this.typed = "";
    this.show();
  }

  exit(): void {
    if (!this.active) return;
    this.active = false;
    this.clear();
    this.onExit();
  }

  private targets(): HTMLElement[] {
    return Array.from(document.querySelectorAll<HTMLElement>(`[data-keytip][data-keytip-scope="${this.scope}"]`))
      .filter((el) => el.offsetParent !== null && el.dataset.keytip);
  }

  private type(k: string): void {
    const next = this.typed + k;
    const matches = this.targets().filter((el) => el.dataset.keytip!.replace(/\s/g, "").toUpperCase().startsWith(next));
    if (!matches.length) return;
    const exact = matches.find((el) => el.dataset.keytip!.replace(/\s/g, "").toUpperCase() === next);
    if (exact && matches.length === 1) {
      this.typed = "";
      exact.click();
      // Ask after the click: activating a ribbon tab changes which scope comes next.
      const nextScope = this.scopeAfter(exact);
      if (exact instanceof HTMLInputElement || exact.querySelector("input")) (exact.querySelector("input") ?? exact).focus();
      if (nextScope) {
        this.scope = nextScope;
        requestAnimationFrame(() => this.show());
      } else {
        this.exit();
      }
      return;
    }
    this.typed = next;
    this.show();
  }

  private clear(): void {
    for (const b of this.badges) b.remove();
    this.badges = [];
  }

  private show(): void {
    this.clear();
    if (!this.active) return;
    for (const el of this.targets()) {
      const key = el.dataset.keytip!.replace(/\s/g, "");
      if (!key.toUpperCase().startsWith(this.typed)) continue;
      const r = el.getBoundingClientRect();
      const small = r.height <= 30 && r.width <= 40;
      const x = r.left + r.width / 2;
      const y = small ? r.top + r.height * 0.75 : r.bottom - 4;
      const b = h("div", { class: "keytip", style: `left:${x}px;top:${y}px` }, key);
      document.body.append(b);
      this.badges.push(b);
    }
  }
}
