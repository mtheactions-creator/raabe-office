// Office's picture dropdown (Step 11, D-161): a box that shows a picture of the current value, and when opened a
// list or grid of pictures (Strichtyp, Verbundtyp, the arrows, Markierungstyp, the Voreinstellungen of the effects,
// Voreingestellte Farbverläufe, Richtung, Textur). One control for every pane and dialog of the three apps. The
// pictures are drawn by us (pictures.ts), never copied from Office. The list opens with showPopup, so it closes by
// the Office rule of D-156; keyboard as in Office: Enter, Leertaste, Alt+Nach-unten or F4 open it, the arrow keys move
// (in a grid by place), Pos1 / Ende jump, Enter or Leertaste picks, Esc closes and gives the focus back to the box.

import { h } from "./dom";
import { closePopup, showPopup } from "./popup";

export interface PictureEntry {
  value: string;
  /** The entry's name: its tooltip and accessible name (Office's labels), drawn beside the picture with `text`. */
  label: string;
  /** The picture: SVG markup of our own (a whole <svg>). */
  svg: string;
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
}

let openFor: HTMLElement | null = null;

/** The closed box (a button). */
export function pictureDropdown(spec: PictureDropSpec): HTMLButtonElement {
  const cur = spec.entries.find((e) => e.value === spec.value());
  const pic = h("span", { class: "pd-pic", style: `width:${spec.size[0]}px;height:${spec.size[1]}px` });
  pic.innerHTML = cur?.svg ?? "";
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
    pic.innerHTML = e.svg;
    const on = e.value === value;
    const b = h("button", {
      type: "button", role: "option", class: `pd-item${on ? " on" : ""}`, "aria-selected": String(on), "aria-label": e.label,
      "data-tip": spec.text ? undefined : e.label, "data-value": e.value, tabindex: "-1",
      onclick: () => { closePopup(); if (e.value !== spec.value()) spec.set(e.value); },
    }, pic, ...(spec.text ? [h("span", { class: "pd-text" }, e.label)] : [])) as HTMLButtonElement;
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
