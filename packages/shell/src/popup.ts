// Menus, popups and tooltips. One popup is open at a time; it closes on outside click,
// Escape, window blur or when another opens, as in Office.

import { h, icon } from "./dom";
import type { Commands } from "./commands";

let open: { el: HTMLElement; close: () => void } | null = null;

export function closePopup(): void {
  const o = open;
  open = null;
  o?.close();
}

/** Shows `el` as a popup anchored below (or beside) `anchor`. */
export function showPopup(el: HTMLElement, anchor: DOMRect, onClose?: () => void): void {
  closePopup();
  document.body.append(el);
  const r = el.getBoundingClientRect();
  let x = anchor.left;
  let y = anchor.bottom + 2;
  if (x + r.width > innerWidth - 4) x = Math.max(4, innerWidth - r.width - 4);
  if (y + r.height > innerHeight - 4) y = Math.max(4, anchor.top - r.height - 2);
  el.style.left = `${x}px`;
  el.style.top = `${y}px`;
  const outside = (e: MouseEvent) => { if (!el.contains(e.target as Node)) closePopup(); };
  const key = (e: KeyboardEvent) => { if (e.key === "Escape") { e.preventDefault(); closePopup(); } };
  setTimeout(() => {
    addEventListener("mousedown", outside, true);
    addEventListener("keydown", key, true);
    addEventListener("blur", closePopup);
  });
  open = {
    el,
    close: () => {
      removeEventListener("mousedown", outside, true);
      removeEventListener("keydown", key, true);
      removeEventListener("blur", closePopup);
      el.remove();
      onClose?.();
    },
  };
}

export function isPopupOpen(): boolean {
  return open !== null;
}

export type MenuEntry = { cmd: string; label?: string } | { label: string; icon?: string; run: () => void; disabled?: boolean } | "-";

export function showMenu(commands: Commands, entries: MenuEntry[], anchor: DOMRect): void {
  const menu = h("div", { class: "menu", role: "menu" });
  for (const e of entries) {
    if (e === "-") { menu.append(h("hr")); continue; }
    if ("cmd" in e) {
      const c = commands.get(e.cmd);
      menu.append(h("button", {
        role: "menuitem", disabled: !commands.isEnabled(e.cmd), title: commands.tooltip(e.cmd),
        onclick: () => { closePopup(); void commands.run(e.cmd); },
      }, icon(c.icon), e.label ?? c.label));
    } else {
      menu.append(h("button", {
        role: "menuitem", disabled: !!e.disabled,
        onclick: () => { closePopup(); e.run(); },
      }, icon(e.icon), e.label));
    }
  }
  showPopup(menu, anchor);
}

// ----- tooltips ---------------------------------------------------------------------------

let tipEl: HTMLElement | null = null;
let tipTimer = 0;

/** Office-style tooltips for elements with data-tip ("Title\nDescription"). */
export function installTooltips(): void {
  const hide = () => { clearTimeout(tipTimer); tipEl?.remove(); tipEl = null; };
  document.addEventListener("mouseover", (e) => {
    const target = (e.target as Element).closest<HTMLElement>("[data-tip]");
    hide();
    if (!target || isPopupOpen()) return;
    tipTimer = window.setTimeout(() => {
      const [title, ...rest] = (target.dataset.tip ?? "").split("\n");
      tipEl = h("div", { class: "tooltip" }, h("b", null, title), rest.join("\n"));
      document.body.append(tipEl);
      const r = target.getBoundingClientRect();
      const t = tipEl.getBoundingClientRect();
      tipEl.style.left = `${Math.min(r.left, innerWidth - t.width - 4)}px`;
      tipEl.style.top = `${r.bottom + 6 + t.height > innerHeight ? r.top - t.height - 6 : r.bottom + 6}px`;
    }, 600);
  });
  document.addEventListener("mousedown", hide, true);
  document.addEventListener("keydown", hide, true);
}
