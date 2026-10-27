// The sheet tabs below the grid, as in Excel: the scroll arrows, one tab per visible sheet
// (active one white with an accent underline), Neues Blatt (+), and the horizontal scroll bar
// on the right with a splitter between them. Click switches, Strg+click selects several,
// double-click renames in place, dragging moves (Strg: copies), right-click opens the menu.

import { h, icon } from "@raabe/shell";
import type { Workbook } from "@raabe/sheet-engine";

export interface SheetBarHost {
  workbook(): Workbook;
  activate(index: number): void;
  add(): void;
  rename(index: number, name: string): string | null;
  moveSheet(from: number, to: number, copy: boolean): void;
  contextMenu(e: MouseEvent, index: number): void;
}

export class SheetBar {
  readonly el: HTMLElement;
  private tabs = h("div", { class: "sheettabs", role: "tablist", "aria-label": "Blätter" });
  private first = 0;
  private dragFrom: number | null = null;

  constructor(private host: SheetBarHost, hscroll: HTMLElement) {
    const nav = h("div", { class: "sheetnav" },
      h("button", { "aria-label": "Zum vorherigen Blatt scrollen", "data-tip": "Strg+Klick: zum ersten Blatt", onclick: (e: MouseEvent) => this.scrollTabs(e.ctrlKey ? -999 : -1) }, icon("chevron-left")),
      h("button", { "aria-label": "Zum nächsten Blatt scrollen", "data-tip": "Strg+Klick: zum letzten Blatt", onclick: (e: MouseEvent) => this.scrollTabs(e.ctrlKey ? 999 : 1) }, icon("chevron-right")));
    const add = h("button", { class: "sheetadd", "aria-label": "Neues Blatt", "data-tip": "Neues Blatt (Umschalt+F11)", onclick: () => host.add() }, "+");
    const splitter = h("div", { class: "sheetsplit", "aria-hidden": "true" });
    const left = h("div", { class: "sheetleft" }, nav, this.tabs, add);
    this.el = h("div", { class: "sheetbar" }, left, splitter, hscroll);
    // The splitter shares the width between the tabs and the scroll bar (Excel).
    splitter.addEventListener("pointerdown", (e) => {
      e.preventDefault();
      splitter.setPointerCapture(e.pointerId);
      const start = e.clientX, w0 = left.getBoundingClientRect().width;
      const move = (ev: PointerEvent) => { left.style.flexBasis = `${Math.max(120, Math.min(this.el.clientWidth - 80, w0 + ev.clientX - start))}px`; };
      const up = () => { splitter.removeEventListener("pointermove", move); splitter.removeEventListener("pointerup", up); };
      splitter.addEventListener("pointermove", move);
      splitter.addEventListener("pointerup", up);
    });
  }

  private scrollTabs(d: number): void {
    const n = this.host.workbook().sheets.length;
    this.first = Math.max(0, Math.min(n - 1, this.first + d));
    this.render();
  }

  render(): void {
    const wb = this.host.workbook();
    this.tabs.replaceChildren();
    wb.sheets.forEach((s, i) => {
      if (s.hidden || i < this.first) return;
      const tab = h("div", {
        class: `sheettab${i === wb.active ? " on" : ""}`, role: "tab", "aria-selected": String(i === wb.active), draggable: "true",
        style: s.tabColor ? `--tab-color:${s.tabColor}` : "",
      }, s.name);
      if (s.tabColor) tab.classList.add("colored");
      tab.addEventListener("mousedown", (e) => {
        if (e.button === 0) this.host.activate(i);
      });
      tab.addEventListener("dblclick", () => this.renameInPlace(tab, i));
      tab.addEventListener("contextmenu", (e) => {
        e.preventDefault();
        this.host.activate(i);
        this.host.contextMenu(e, i);
      });
      tab.addEventListener("dragstart", (e) => { this.dragFrom = i; e.dataTransfer?.setData("text/plain", s.name); });
      tab.addEventListener("dragover", (e) => { e.preventDefault(); tab.classList.add("drop"); });
      tab.addEventListener("dragleave", () => tab.classList.remove("drop"));
      tab.addEventListener("drop", (e) => {
        e.preventDefault();
        tab.classList.remove("drop");
        if (this.dragFrom !== null && this.dragFrom !== i) this.host.moveSheet(this.dragFrom, i, e.ctrlKey);
        this.dragFrom = null;
      });
      this.tabs.append(tab);
    });
  }

  /** Double-click: the tab's name becomes editable; Enter keeps it, Esc goes back. */
  renameInPlace(tab: HTMLElement | null, index: number): void {
    tab ??= this.tabs.children[index - this.first] as HTMLElement | null;
    if (!tab) return;
    const wb = this.host.workbook();
    const input = h("input", { class: "sheetrename", value: wb.sheets[index].name, maxlength: "31" }) as HTMLInputElement;
    tab.replaceChildren(input);
    input.focus();
    input.select();
    let done = false;
    const finish = (keep: boolean) => {
      if (done) return;
      done = true;
      if (keep && input.value.trim() && input.value !== wb.sheets[index].name) {
        const err = this.host.rename(index, input.value.trim());
        if (err) { done = false; input.focus(); return; }
      }
      this.render();
    };
    input.addEventListener("keydown", (e) => {
      e.stopPropagation();
      if (e.key === "Enter") finish(true);
      if (e.key === "Escape") finish(false);
    });
    input.addEventListener("blur", () => finish(true));
  }
}
