// Menus, popups and tooltips. One popup is open at a time; it closes on a press outside it,
// Escape, window blur or when another opens, as in Office (D-156).

import { h, icon } from "./dom";
import type { Commands } from "./commands";

let open: { el: HTMLElement; close: () => void } | null = null;

/** The control pressed last (the one that opens a popup), and when. */
let lastDown: { el: Element; at: number } | null = null;
addEventListener("pointerdown", (e) => {
  const el = (e.target as Element | null)?.closest?.("button, [role=button], .rcombo");
  lastDown = el ? { el, at: performance.now() } : null;
}, true);

/**
 * The app's frame around the document: title bar, ribbon, status bar, task panes, dialogs, backstage. A press there
 * while a popup is open closes the popup and still acts (Excel, Word and PowerPoint, recorded 2026-11-11: a ribbon
 * tab switches, Fett applies, the zoom button zooms).
 */
export const FRAME_SELECTOR = ".titlebar, .tabrow, .ribbon-card, .mini-toolbar, .statusbar, .backstage, [role=complementary], [role=dialog]";

/**
 * A press outside an open popup, as Office handles it (D-156). The popup closes in every case (`close`). A left press
 * in the document (a cell, a header, a sheet tab, the formula bar, the text, the slide, a thumbnail) does nothing
 * else: Office swallows it, the selection stays. A right press there acts (it selects and opens the next context
 * menu), and so does any press in the app's frame. Listen with `pointerdown` in the capture phase: a surface that
 * cancels `pointerdown` (the grid does) stops the `mousedown` that would follow, so a `mousedown` listener never
 * hears that press.
 */
export function pressOutside(e: PointerEvent, close: () => void): void {
  close();
  if (e.button !== 0) return;
  const t = e.target as Element | null;
  if (t?.closest?.(FRAME_SELECTOR)) return;
  swallowPress(e);
}

/** Swallows a whole press: this pointerdown and the mouse and click events that follow it; focus stays where it was. */
export function swallowPress(e: PointerEvent): void {
  e.preventDefault();
  e.stopImmediatePropagation();
  const before = document.activeElement as HTMLElement | null;
  // The click ends the swallowed press (its listener must be this one: it stops every later listener).
  const block = (ev: Event) => { ev.preventDefault(); ev.stopImmediatePropagation(); if (ev.type === "click") setTimeout(end); };
  const types = ["mousedown", "pointerup", "mouseup", "click", "dblclick", "auxclick"] as const;
  for (const t of types) addEventListener(t, block, true);
  let done = false;
  const end = () => {
    if (done) return;
    done = true;
    for (const t of types) removeEventListener(t, block, true);
    removeEventListener("pointerdown", end, true);
    // Chromium may still move the focus for a cancelled press (an input under the pointer): put it back.
    if (before?.isConnected && document.activeElement !== before) before.focus({ preventScroll: true });
  };
  // The next press, or the timeout, ends it when no click follows (the pointer was released outside the window).
  addEventListener("pointerdown", end, true);
  setTimeout(end, 800);
}

/** The registry whose Livevorschau a menu started (it ends when the menu closes). */
let hovered: Commands | null = null;

export function closePopup(): void {
  const o = open;
  open = null;
  hovered?.hover(null);
  hovered = null;
  o?.close();
}

/** Livevorschau for a menu or gallery entry: shown while the mouse is over `el`. */
export function hoverPreview(commands: Commands, el: HTMLElement, target: string | (() => void)): void {
  el.addEventListener("mouseenter", () => { hovered = commands; commands.hover(target); });
  el.addEventListener("mouseleave", () => commands.hover(null));
  // A click ends the preview first, so the command itself makes the change (one undo step).
  el.addEventListener("mousedown", () => commands.hover(null), true);
}

/** Shows `el` as a popup anchored below (or beside) `anchor`. */
export function showPopup(el: HTMLElement, anchor: DOMRect, onClose?: () => void): void {
  closePopup();
  document.body.append(el);
  // Every popup floats at its anchor. Panels without a look of their own (a gallery's tile panel) get the card of a
  // menu; without this they were laid out at the bottom of the page.
  const cs = getComputedStyle(el);
  if (cs.position === "static") { el.style.position = "fixed"; el.style.zIndex = "70"; }
  if (cs.backgroundColor === "rgba(0, 0, 0, 0)" && !el.classList.contains("menu-root")) el.classList.add("popup-card");
  const r = el.getBoundingClientRect();
  let x = anchor.left;
  let y = anchor.bottom + 2;
  if (x + r.width > innerWidth - 4) x = Math.max(4, innerWidth - r.width - 4);
  if (y + r.height > innerHeight - 4) y = Math.max(4, anchor.top - r.height - 2);
  el.style.left = `${x}px`;
  el.style.top = `${y}px`;
  // The control that opened the popup: pressing it again closes the popup (Office), and the
  // click that follows must not open it anew.
  const opener = lastDown && performance.now() - lastDown.at < 1500 ? lastDown.el : null;
  const outside = (e: PointerEvent) => {
    if (el.contains(e.target as Node)) return;
    if (opener && opener.contains(e.target as Node)) {
      // Pressing the opener again only closes the popup (Office): the press must not open it anew.
      closePopup();
      swallowPress(e);
      return;
    }
    pressOutside(e, closePopup);
  };
  const key = (e: KeyboardEvent) => { if (e.key === "Escape") { e.preventDefault(); closePopup(); } };
  setTimeout(() => {
    addEventListener("pointerdown", outside, true);
    addEventListener("keydown", key, true);
    addEventListener("blur", closePopup);
  });
  // For tests and styles: whether a popup is open (check_menuclose_*.py).
  document.documentElement.setAttribute("data-popup-open", "");
  open = {
    el,
    close: () => {
      document.documentElement.removeAttribute("data-popup-open");
      removeEventListener("pointerdown", outside, true);
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

export type MenuEntry =
  /** `key`: the entry's keytip once the menu was opened by keytips (Alt, F, M, R). */
  | { cmd: string; label?: string; key?: string }
  | { label: string; icon?: string; run: () => void; disabled?: boolean; checked?: boolean; preview?: HTMLElement; bold?: boolean; key?: string;
      /** Livevorschau while the mouse is over the entry (the change `run` makes, shown without an undo step). */
      hover?: () => void }
  | { heading: string }
  /** A submenu opening to the right (Word's "Einfügen ▸" in a table's context menu). */
  | { label: string; icon?: string; submenu: MenuEntry[] | (() => MenuEntry[]); disabled?: boolean }
  /** A submenu with own content (a gallery of tiles: Bildeffekte > Schatten). */
  | { label: string; icon?: string; flyout: () => HTMLElement; disabled?: boolean }
  /** Own content inside the menu (the presets gallery at the top of Word's Texteffekte menu). */
  | { block: () => HTMLElement }
  /** A row of icon buttons under a label (Word's "Einfügeoptionen:"). */
  | { label: string; icons: { icon: string; tip: string; run: () => void; disabled?: boolean }[] }
  | "-";

const SUB_CHEVRON = '<svg viewBox="0 0 5 8" width="5" height="8" aria-hidden="true"><path d="M.75.75 4.25 4 .75 7.25" fill="none" stroke="currentColor" stroke-width="1" stroke-linecap="round" stroke-linejoin="round"/></svg>';

function buildMenu(commands: Commands, entries: MenuEntry[], root: HTMLElement, level: number): HTMLElement {
  const menu = h("div", { class: `menu${level ? " submenu" : ""}`, role: "menu" });
  let openSub: HTMLElement | null = null;
  const closeSub = () => { openSub?.remove(); openSub = null; };
  for (const e of entries) {
    if (e === "-") { menu.append(h("hr")); continue; }
    if ("heading" in e) { menu.append(h("div", { class: "menu-heading" }, e.heading)); continue; }
    if ("block" in e) { menu.append(e.block()); continue; }
    if ("icons" in e) {
      const row = h("div", { class: "menu-icons" }, ...e.icons.map((ic) => h("button", {
        class: "menu-icon", "data-tip": ic.tip, "aria-label": ic.tip, disabled: !!ic.disabled,
        onclick: () => { closePopup(); ic.run(); },
      }, icon(ic.icon))));
      menu.append(h("div", { class: "menu-icons-label" }, icon(undefined), e.label), row);
      continue;
    }
    if ("submenu" in e || "flyout" in e) {
      const item = h("button", { role: "menuitem", class: "has-sub", disabled: !!e.disabled, "aria-haspopup": "true" }, icon(e.icon), e.label);
      const chev = h("span", { class: "sub-chev" });
      chev.innerHTML = SUB_CHEVRON;
      item.append(chev);
      const open = () => {
        if (e.disabled) return;
        closeSub();
        let sub: HTMLElement;
        if ("flyout" in e) {
          sub = h("div", { class: "menu submenu flyout" }, e.flyout());
        } else {
          const items = typeof e.submenu === "function" ? e.submenu() : e.submenu;
          sub = buildMenu(commands, items, root, level + 1);
        }
        root.append(sub);
        const r = item.getBoundingClientRect();
        sub.style.left = `${r.right - 2}px`;
        sub.style.top = `${r.top - 4}px`;
        const sr = sub.getBoundingClientRect();
        if (sr.right > innerWidth - 4) sub.style.left = `${r.left - sr.width + 2}px`;
        if (sr.bottom > innerHeight - 4) sub.style.top = `${Math.max(4, innerHeight - sr.height - 4)}px`;
        openSub = sub;
      };
      item.addEventListener("mouseenter", open);
      item.addEventListener("click", open);
      menu.append(item);
      continue;
    }
    if ("cmd" in e) {
      const c = commands.get(e.cmd);
      // A command without an icon of its own that is on shows the check (Relative Verweise verwenden).
      const checked = !c.icon && !!c.checked && commands.isChecked(e.cmd);
      const b = h("button", {
        role: "menuitem", disabled: !commands.isEnabled(e.cmd), title: commands.tooltip(e.cmd), class: checked ? "checked" : "",
        "aria-checked": c.checked ? String(checked) : undefined,
        "data-keytip": e.key, "data-keytip-scope": e.key ? "menu" : undefined,
        onclick: () => { commands.hover(null); closePopup(); void commands.run(e.cmd); },
      }, icon(checked ? "check" : c.icon), e.label ?? c.label);
      b.addEventListener("mouseenter", closeSub);
      if (c.preview) hoverPreview(commands, b, e.cmd);
      menu.append(b);
    } else {
      const b = h("button", {
        role: "menuitem", disabled: !!e.disabled, class: e.preview ? "with-preview" : e.checked ? "checked" : "",
        style: e.bold ? "font-weight:600" : undefined, "data-keytip": e.key, "data-keytip-scope": e.key ? "menu" : undefined,
        onclick: () => { commands.hover(null); closePopup(); e.run(); },
      }, e.preview ?? icon(e.checked ? "check" : e.icon), e.preview ? h("span", null, e.label) : e.label);
      b.addEventListener("mouseenter", closeSub);
      if (e.hover && !e.disabled) hoverPreview(commands, b, e.hover);
      menu.append(b);
    }
  }
  return menu;
}

/**
 * A menu at `anchor`. `before` stands above the menu (Word's mini toolbar), `header` at the
 * top inside it (the search box of the context menu).
 */
export function showMenu(commands: Commands, entries: MenuEntry[], anchor: DOMRect, before?: HTMLElement, header?: HTMLElement): void {
  // The popup holds the menu and its submenus (fixed positions), so a click in a submenu
  // does not count as a click outside.
  const root = h("div", { class: "menu-root" });
  if (before) root.append(before);
  const menu = buildMenu(commands, entries, root, 0);
  if (header) menu.prepend(header);
  root.append(menu);
  showPopup(root, anchor);
  // Word: the menu starts at the mouse, the toolbar sits above it.
  if (before) {
    const lift = before.getBoundingClientRect().height + 4;
    const top = parseFloat(root.style.top);
    if (top - lift >= 4 && top >= anchor.bottom) root.style.top = `${top - lift}px`;
  }
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
    // While a menu or gallery is open, its own entries get tooltips (as in Word), nothing else.
    if (!target || (isPopupOpen() && !(open?.el.contains(target) || target.closest(".menu, .rpanel")))) return;
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
  document.addEventListener("pointerdown", hide, true);
  document.addEventListener("keydown", hide, true);
}
