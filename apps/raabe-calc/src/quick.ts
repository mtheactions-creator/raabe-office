// Schnellanalyse as Excel shows it: a small button at the bottom right of a selection of several
// cells with data (22 px, 2 px beyond the corner; recorded), Strg+Q or a click opens the panel
// under it with the tabs Formatierung, Diagramme, Ergebnisse, Tabellen and Sparklines. Each tile
// does what Excel's does (probe_excel_quickanalysis.py, results/2026-10-28-excel-quickanalysis.json).

import { h, icon } from "@raabe/shell";
import type { Rect } from "@raabe/sheet-engine";
import type { CalcApp } from "./app";

export interface QuickTile {
  label: string;
  icon: string;
  run: (r: Rect) => void;
  disabled?: boolean;
}

export function quickFeature(app: CalcApp, tiles: () => Record<string, QuickTile[]>) {
  const button = h("button", { class: "quick-btn", "aria-label": "Schnellanalyse", "data-tip": "Schnellanalyse (Strg+Q)", hidden: true, onclick: () => open() }, icon("quick-analysis"));
  let panel: HTMLElement | null = null;
  let tab = "Formatierung";

  /** The selection Schnellanalyse works on: one range of at least two cells, one of them not empty. */
  function target(): Rect | null {
    const sel = app.grid.selection;
    if (sel.length !== 1 || app.grid.isEditing) return null;
    const r = sel[0];
    if (r.r0 === r.r1 && r.c0 === r.c1) return null;
    if ((r.r1 - r.r0 + 1) * (r.c1 - r.c0 + 1) > 1_000_000) return null;
    for (const [key, c] of app.sheet.cells) {
      const row = Math.floor(key / 16384), col = key % 16384;
      if (row >= r.r0 && row <= r.r1 && col >= r.c0 && col <= r.c1 && c.v !== null && c.v !== "") return r;
    }
    return null;
  }

  function place(): void {
    if (!button.isConnected) app.grid.el.append(button);
    const r = target();
    const b = r ? app.grid.rectOnCanvas(r) : null;
    const hidden = !r || !b || app.grid.optionsShown();
    button.hidden = hidden;
    if (hidden) { close(); return; }
    button.style.left = `${b!.x + b!.w + 2}px`;
    button.style.top = `${b!.y + b!.h + 2}px`;
  }

  function close(): void {
    panel?.remove();
    panel = null;
    removeEventListener("pointerdown", outside, true);
    removeEventListener("keydown", keys, true);
  }
  const outside = (e: Event) => { if (panel && !panel.contains(e.target as Node) && e.target !== button && !button.contains(e.target as Node)) close(); };
  const keys = (e: KeyboardEvent) => {
    if (!panel) return;
    if (e.key === "Escape") { e.preventDefault(); e.stopPropagation(); close(); app.grid.focus(); }
    else if (e.key === "ArrowRight" || e.key === "ArrowLeft") {
      e.preventDefault(); e.stopPropagation();
      const names = Object.keys(tiles());
      tab = names[(names.indexOf(tab) + (e.key === "ArrowRight" ? 1 : -1) + names.length) % names.length];
      render();
    }
  };

  function render(): void {
    if (!panel) return;
    const r = target();
    const all = tiles();
    panel.replaceChildren(
      h("div", { class: "quick-tabs", role: "tablist" }, ...Object.keys(all).map((t) => h("button", { type: "button", role: "tab", class: t === tab ? "on" : "", onclick: () => { tab = t; render(); } }, t))),
      h("div", { class: "quick-tiles" }, ...all[tab].map((x) => h("button", { type: "button", class: "quick-tile", disabled: x.disabled || !r ? true : undefined,
        onclick: () => { close(); if (r) x.run(r); app.grid.focus(); } }, icon(x.icon, "l"), h("span", null, x.label)))));
  }

  /** Strg+Q or the button: the panel under the button. */
  function open(): boolean {
    place();
    if (button.hidden) return false;
    if (panel) { close(); return true; }
    panel = h("div", { class: "quick-panel", role: "dialog", "aria-label": "Schnellanalyse" });
    render();
    app.grid.el.append(panel);
    panel.style.left = button.style.left;
    panel.style.top = `${parseFloat(button.style.top) + 24}px`;
    setTimeout(() => { addEventListener("pointerdown", outside, true); addEventListener("keydown", keys, true); });
    return true;
  }

  return { place, open, close };
}
