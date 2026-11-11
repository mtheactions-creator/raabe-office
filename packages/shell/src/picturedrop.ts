// Office's picture dropdown (Step 11, D-161): a box that shows a picture of the current value, and when opened a
// list or grid of pictures (Strichtyp, Verbundtyp, the arrows, Markierungstyp, the Voreinstellungen of the effects,
// Voreingestellte Farbverläufe, Richtung, Textur). One control for every pane and dialog of the three apps. The
// pictures are drawn by us (pictures.ts), never copied from Office. The list opens with showPopup, so it closes by
// the Office rule of D-156; keyboard as in Office: Enter, Leertaste, Alt+Nach-unten or F4 open it, the arrow keys move
// (in a grid by place), Pos1 / Ende jump, Enter or Leertaste picks, Esc closes and gives the focus back to the box.

import { h } from "./dom";
import { closePopup, hoverPreview, showPopup } from "./popup";
import type { Commands } from "./commands";

/**
 * Livevorschau for picture entries (D-166): while the mouse is over an entry the app shows what it would do, through
 * the shell's hoverPreview, so it ends when the mouse leaves, on a click (the pick is one undo step) and when the menu
 * closes. `preview` applies the entry; the app's hover handler runs it without an undo step.
 */
export interface PictureLive { commands: Commands; preview: (v: string) => void }

export interface PictureEntry {
  value: string;
  /** The entry's name: its tooltip and accessible name (Office's labels), drawn beside the picture with `text`. */
  label: string;
  /** The picture: SVG markup of our own (a whole <svg>). */
  svg: string;
  /** Or a picture drawn as an element (an effect preset drawn by the app's own renderer); wins over `svg`. */
  node?: () => Element;
  /** A heading above this entry starts a new group (Schatten: Außen, Innen, Perspektive). */
  group?: string;
  /** A group of its own with this many columns (the "Kein Schatten" tile above the grid). */
  columns?: number;
}

export interface PictureDropSpec {
  label: string;
  entries: PictureEntry[];
  value: () => string;
  set: (v: string) => void;
  disabled?: boolean;
  /** 1: a list (Strichtyp); more: a grid (Voreingestellte Farbverläufe: 6). */
  columns: number;
  /** The picture's size in px [width, height], in the box and in the open list. */
  size: [number, number];
  /** Entries draw their label beside the picture (Office's lists with text). */
  text?: boolean;
  /** The closed box shows the current value's name beside its picture. */
  boxText?: boolean;
  /** Commands under the entries ("Weitere Pfeile..."). */
  more?: { label: string; run: () => void }[];
  /** Livevorschau while the mouse is over an entry (D-166). */
  live?: PictureLive;
}

let openFor: HTMLElement | null = null;

/**
 * The same pictures inside a menu (a ribbon submenu such as Formkontur > Striche, which Office shows as pictures):
 * a block of picture buttons for a `{ block }` menu entry; a pick closes the menu.
 */
export function pictureMenu(entries: PictureEntry[], columns: number, size: [number, number], value: string | undefined, pick: (v: string) => void, live?: PictureLive): HTMLElement {
  const grid = h("div", { class: "pd-grid pd-menu", role: "listbox", style: `grid-template-columns:repeat(${columns}, auto)` });
  for (const e of entries) {
    const pic = h("span", { class: "pd-pic", style: `width:${size[0]}px;height:${size[1]}px` });
    if (e.node) pic.append(e.node()); else pic.innerHTML = e.svg;
    const on = e.value === value;
    const b = h("button", { type: "button", role: "option", class: `pd-item${on ? " on" : ""}`, "aria-selected": String(on), "aria-label": e.label, "data-tip": e.label,
      "data-value": e.value, onclick: () => { closePopup(); pick(e.value); } }, pic);
    if (live) hoverPreview(live.commands, b, () => live.preview(e.value));
    grid.append(b);
  }
  return grid;
}

/**
 * How an entry of a menu, gallery or list shows itself (the sweep of D-163): "picture" (a drawing, its name only a
 * tooltip), "text" (a name, maybe with a small icon) or "both" (a drawing with its name, as Office's Stärke menu).
 */
export function entryKind(el: Element): "picture" | "text" | "both" {
  const text = (el.textContent ?? "").replace(/\s+/g, " ").trim();
  const graphic = [...el.querySelectorAll("svg, img, canvas, .pd-pic, .swatch, [style*='background']")].some((g) => {
    // A menu's 16 px icon is not a picture of the entry (Office's text menus have icons too).
    const w = Number(g.getAttribute("width")) || (g as HTMLElement).offsetWidth || 0;
    return !g.closest(".ico") && !(g.tagName.toLowerCase() === "svg" && w > 0 && w <= 20 && text);
  }) || (!text && el.children.length > 0) || el.classList.contains("swatch") || /background/.test(el.getAttribute("style") ?? "");
  // An entry without a name of its own on it (a cell of the table grid) or with a one- or two-letter glyph (the WordArt
  // and Texteffekte previews, an "A" in the style) is a picture too.
  if (!text || text.length <= 2) return "picture";
  if (graphic && text) return "both";
  return graphic ? "picture" : "text";
}

/** A line of a weight (Stärke menus: Office shows the line and its name). */
export function weightPicture(pt: number, w = 64, hgt = 14): HTMLElement {
  const s = h("span", { class: "pd-pic", style: `width:${w}px;height:${hgt}px` });
  const px = Math.max(0.75, pt * (96 / 72));
  s.innerHTML = `<svg xmlns="http://www.w3.org/2000/svg" width="${w}" height="${hgt}" aria-hidden="true"><rect x="0" y="${(hgt - px) / 2}" width="${w}" height="${px}" fill="#404040"/></svg>`;
  return s;
}

/** The closed box (a button). */
export function pictureDropdown(spec: PictureDropSpec): HTMLButtonElement {
  const cur = spec.entries.find((e) => e.value === spec.value());
  const pic = h("span", { class: "pd-pic", style: `width:${spec.size[0]}px;height:${spec.size[1]}px` });
  if (cur?.node) pic.append(cur.node()); else pic.innerHTML = cur?.svg ?? "";
  const box = h("button", {
    type: "button", class: "picdrop", disabled: !!spec.disabled, "aria-haspopup": "listbox", "aria-expanded": "false",
    "aria-label": spec.label, "data-tip": cur ? `${spec.label}: ${cur.label}` : spec.label, "data-value": cur?.value ?? "",
  }, pic, ...(spec.boxText && cur ? [h("span", { class: "pd-text" }, cur.label)] : []), h("span", { class: "pd-chev", "aria-hidden": "true" })) as HTMLButtonElement;
  box.addEventListener("click", () => {
    if (openFor === box) { closePopup(); return; }
    openList(box, spec);
  });
  box.addEventListener("keydown", (e) => {
    if (e.key === "F4" || (e.altKey && e.key === "ArrowDown") || e.key === "Enter" || e.key === " ") {
      e.preventDefault();
      openList(box, spec);
    }
  });
  return box;
}

function openList(box: HTMLButtonElement, spec: PictureDropSpec): void {
  const value = spec.value();
  const list = h("div", { class: "menu picdrop-list", role: "listbox", "aria-label": spec.label, tabindex: "-1" });
  const items: HTMLButtonElement[] = [];
  let grid: HTMLElement | null = null;
  let group: string | undefined;
  for (const e of spec.entries) {
    if (!grid || e.group !== group || e.columns) {
      group = e.group;
      if (e.group) list.append(h("div", { class: "menu-heading" }, e.group));
      const cols = e.columns ?? spec.columns;
      grid = h("div", { class: `pd-grid${spec.text ? " with-text" : ""}`, style: `grid-template-columns:repeat(${cols}, auto)` });
      list.append(grid);
    }
    const pic = h("span", { class: "pd-pic", style: `width:${spec.size[0]}px;height:${spec.size[1]}px` });
    if (e.node) pic.append(e.node()); else pic.innerHTML = e.svg;
    const on = e.value === value;
    const b = h("button", {
      type: "button", role: "option", class: `pd-item${on ? " on" : ""}`, "aria-selected": String(on), "aria-label": e.label,
      "data-tip": spec.text ? undefined : e.label, "data-value": e.value, tabindex: "-1",
      onclick: () => { closePopup(); if (e.value !== spec.value()) spec.set(e.value); },
    }, pic, ...(spec.text ? [h("span", { class: "pd-text" }, e.label)] : [])) as HTMLButtonElement;
    if (spec.live) { const live = spec.live; hoverPreview(live.commands, b, () => live.preview(e.value)); }
    items.push(b);
    grid.append(b);
  }
  if (spec.more?.length) {
    list.append(h("hr"), ...spec.more.map((m) => h("button", { type: "button", class: "item pd-more", onclick: () => { closePopup(); m.run(); } }, m.label)));
  }
  const focus = (i: number) => {
    const b = items[Math.max(0, Math.min(items.length - 1, i))];
    if (!b) return;
    for (const x of items) x.classList.toggle("focus", x === b);
    b.focus({ preventScroll: true });
    b.scrollIntoView?.({ block: "nearest" });
  };
  /** The entry nearest in a direction (grids with groups of other widths: by place, not by index). */
  const step = (from: number, dx: number, dy: number): number => {
    if (spec.columns === 1 && !dx) return from + dy;
    const r = items[from].getBoundingClientRect();
    const cx = r.left + r.width / 2, cy = r.top + r.height / 2;
    let best = from, bestD = Infinity;
    items.forEach((b, i) => {
      if (i === from) return;
      const q = b.getBoundingClientRect();
      const qx = q.left + q.width / 2, qy = q.top + q.height / 2;
      const ax = qx - cx, ay = qy - cy;
      if (dx && (Math.sign(ax) !== dx || Math.abs(ay) > r.height / 2)) return;
      if (dy && (Math.sign(ay) !== dy)) return;
      const d = dx ? Math.abs(ax) : Math.abs(ay) * 1000 + Math.abs(ax);
      if (d < bestD) { bestD = d; best = i; }
    });
    // Left and right run on to the row before or after at the ends of a row (Office's galleries).
    if (best === from && dx) return Math.max(0, Math.min(items.length - 1, from + dx));
    return best;
  };
  list.addEventListener("keydown", (ev) => {
    const at = Math.max(0, items.findIndex((b) => b === document.activeElement));
    const moves: Record<string, () => number> = {
      ArrowDown: () => step(at, 0, 1), ArrowUp: () => step(at, 0, -1), ArrowRight: () => step(at, 1, 0), ArrowLeft: () => step(at, -1, 0),
      Home: () => 0, End: () => items.length - 1,
    };
    if (moves[ev.key]) { ev.preventDefault(); focus(moves[ev.key]()); return; }
    if (ev.key === "Enter" || ev.key === " ") { ev.preventDefault(); items[at]?.click(); return; }
    if (ev.key === "Tab") { ev.preventDefault(); closePopup(); }
  });
  const r = box.getBoundingClientRect();
  showPopup(list, r, () => {
    openFor = null;
    box.setAttribute("aria-expanded", "false");
    // Esc and a pick give the focus back to the box (as Office), unless the focus went elsewhere.
    if (!document.activeElement || document.activeElement === document.body || list.contains(document.activeElement)) box.focus({ preventScroll: true });
  });
  openFor = box;
  box.setAttribute("aria-expanded", "true");
  focus(Math.max(0, items.findIndex((b) => b.classList.contains("on"))));
}
