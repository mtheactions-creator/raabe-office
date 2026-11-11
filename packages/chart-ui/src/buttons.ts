// The three buttons beside a selected chart (D-115, D-116), as German Excel 365 shows them
// (probe_excel_chart_buttons.py, results/2026-11-09-excel-chart-buttons.json): 28 x 28 px, 11 px right of the
// chart's frame, the first at its top and the others 38 px apart.
// - Diagrammelemente ("+"): the elements with a check box each and, on the hovered entry, an arrow at the right
//   end that opens its side menu (positions, sub-elements, "Weitere Optionen..." for the Format pane).
// - Diagrammformatvorlagen (brush): Formatvorlage (the chart drawn in each style) and Farbe (Bunt / Einfarbig
//   palettes); the entry under the mouse is shown on the chart (Livevorschau), a click applies it.
// - Diagrammfilter (funnel): Werte (Datenreihe and Kategorien with check boxes, hovering one highlights it in the
//   chart) and Namen (where series and category names come from); Anwenden applies, Daten auswählen... opens the
//   Datenquelle auswählen dialog.

import { isPie, kindInfo, paletteColor, schemeColor, seriesName, visibleModel, type ChartModel } from "@raabe/charts";
import { h, pressOutside } from "@raabe/shell";
import { choiceLabel, elementEntries, type ElementEntry } from "./elements";
import type { ChartElement } from "./formatpane";
import { chartStyles, applyStyle, styleIndex } from "./styles";

export interface ButtonsHost {
  /** Changes the chart as one undo step (and redraws it). */
  edit(m: ChartModel, fn: (m: ChartModel) => void, label?: string): void;
  /** Livevorschau: draws `look` in place of the chart (null ends it). */
  preview(m: ChartModel, look: ChartModel | null): void;
  /** Fades all but one series or category in the drawing (null ends it). */
  highlight(m: ChartModel, what: { series?: number; cat?: number } | null): void;
  /** Opens the Format pane of an element ("Weitere Optionen..."). */
  openPane(m: ChartModel, e: ChartElement): void;
  /** Datenquelle auswählen. */
  selectData(m: ChartModel): void;
  /** The theme's accents (for the colour rows and the style tiles). */
  accents(): string[];
  /** A small drawing of a chart (the style tiles). */
  tile(m: ChartModel, w: number, h: number): SVGSVGElement;
  /** Excel asks which series gets a trendline when there are several. */
  pickSeries(m: ChartModel): Promise<number | null>;
  /** The names Namen offers: the row/column the series names and categories come from (null: none). */
  nameSources(m: ChartModel): { series: string | null; cats: string | null; seriesList: string; catList: string };
  /** Namen: series names / categories from the cells, or none (Reihe1 …, 1, 2 …). */
  setNameSources(m: ChartModel, series: boolean, cats: boolean): void;
}

const BTN = 28, GAP_X = 11, PITCH = 38;

const CSS = `
.chart-btn { position: absolute; width: ${BTN}px; height: ${BTN}px; display: none; align-items: center; justify-content: center; padding: 0; box-sizing: border-box;
  background: #fff; border: 1px solid #8a8a8a; cursor: default; z-index: 2; }
.chart-btn:hover { background: #f3f3f3; }
.chart-btn.on { background: #d3f0e0; border-color: #107c41; }
.chart-btn svg { display: block; pointer-events: none; }
.chart-obj.sel .chart-btn { display: flex; }
.cb-fly { position: fixed; z-index: 1000; background: #fff; border: 1px solid #8a8a8a; box-shadow: 0 2px 6px rgba(0,0,0,.18); font: 12px var(--font, "Segoe UI", sans-serif); color: #262626; box-sizing: border-box; }
.cb-fly::before { content: ""; position: absolute; left: -7px; top: 10px; width: 12px; height: 12px; background: #fff; border-left: 1px solid #8a8a8a; border-bottom: 1px solid #8a8a8a; transform: rotate(45deg); }
.cb-fly.left::before { left: auto; right: -7px; border: 0; border-right: 1px solid #8a8a8a; border-top: 1px solid #8a8a8a; }
.cb-head { padding: 7px 10px 4px; font-weight: 700; font-size: 12px; color: #107c41; }
.cb-list { padding: 0 6px 7px; }
.cb-row { display: flex; align-items: center; height: 21px; position: relative; }
.cb-row label { flex: 1; display: flex; align-items: center; gap: 9px; height: 100%; padding-left: 4px; }
.cb-row.on label { background: #ebebeb; }
.cb-row:hover label { background: #e1e1e1; }
.cb-row input { margin: 0; width: 13px; height: 13px; appearance: none; -webkit-appearance: none; border: 1px solid #444; background: #fff; display: inline-grid; place-content: center; }
.cb-row input:checked::before { content: ""; width: 9px; height: 7px; background: no-repeat center/contain url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 9 7'%3E%3Cpath d='M.6 3.4 3.3 6 8.4.8' fill='none' stroke='%23262626' stroke-width='1.1'/%3E%3C/svg%3E"); }
.cb-arrow { width: 30px; height: 100%; border: 0; background: transparent; padding: 0; visibility: hidden; display: flex; align-items: center; justify-content: center; }
.cb-row:hover .cb-arrow, .cb-arrow.open { visibility: visible; }
.cb-arrow:hover, .cb-arrow.open { background: #e1e1e1; }
.cb-sub { position: fixed; z-index: 1001; min-width: 170px; padding: 4px 0; background: #fff; border: 1px solid #8a8a8a; box-shadow: 0 2px 6px rgba(0,0,0,.18); font: 12px var(--font, "Segoe UI", sans-serif); }
.cb-sub button { display: flex; align-items: center; gap: 8px; width: 100%; height: 26px; padding: 0 12px 0 8px; border: 0; background: none; font: inherit; text-align: left; white-space: nowrap; }
.cb-sub button:hover { background: #e1e1e1; }
.cb-sub .mark { width: 14px; display: inline-flex; justify-content: center; }
.cb-sub hr { border: 0; border-top: 1px solid #e1e1e1; margin: 3px 0; }
.cb-tabs { display: flex; gap: 0; padding: 4px 6px 0; }
.cb-tab { border: 0; background: none; font: inherit; font-size: 13px; padding: 6px 10px 5px; border-bottom: 2px solid transparent; color: #262626; }
.cb-tab.sel { font-weight: 600; border-bottom-color: #107c41; }
.cb-scroll { overflow-y: auto; padding: 4px 4px 4px 6px; }
.cb-tile { display: block; margin: 0 0 2px; padding: 2px; border: 1px solid #d6d6d6; background: #fff; }
.cb-tile:hover { border-color: #8a8a8a; background: #f3f3f3; }
.cb-tile.sel { border-color: #616161; background: #e1e1e1; }
.cb-pal-head { font-weight: 600; padding: 6px 4px 3px; }
.cb-pal { display: flex; gap: 3px; padding: 6px 6px; border: 1px solid transparent; }
.cb-pal:hover { border-color: #8a8a8a; background: #f3f3f3; }
.cb-pal.sel { border-color: #616161; background: #e1e1e1; }
.cb-pal span { width: 22px; height: 18px; display: block; }
.cb-group-head { display: flex; align-items: center; gap: 6px; height: 21px; width: 100%; border: 0; background: none; font: inherit; font-weight: 700; padding: 0 4px; }
.cb-filter-row { display: flex; align-items: center; gap: 8px; height: 29px; padding-left: 22px; }
.cb-filter-row:hover { background: #ebebeb; }
.cb-filter-row input { margin: 0; width: 14px; height: 14px; accent-color: #107c41; }
.cb-swatch { width: 6px; height: 13px; display: inline-block; margin-right: 2px; }
.cb-filter-row .lbl { white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
.cb-radio { display: flex; align-items: flex-start; gap: 8px; padding: 4px 4px 4px 22px; }
.cb-radio small { display: block; color: #595959; }
.cb-foot { display: flex; align-items: center; justify-content: space-between; border-top: 1px solid #d6d6d6; margin: 0 10px; padding: 6px 0 8px; }
.cb-foot button.apply { font: inherit; padding: 4px 9px; border: 1px solid #262626; background: #fff; }
.cb-foot button.apply:disabled { border-color: #c6c6c6; color: #a6a6a6; }
.cb-foot a { color: #107c41; text-decoration: underline; cursor: pointer; }
`;

const ICON = {
  plus: '<svg width="20" height="20" viewBox="0 0 20 20"><path d="M10 2v16M2 10h16" stroke="#107c41" stroke-width="1.3" fill="none"/></svg>',
  brush: '<svg width="20" height="20" viewBox="0 0 20 20"><path d="M18 2 9.2 11.4" stroke="#444" stroke-width="1.2" fill="none"/><path d="M8.4 10.6c-2-.2-3.4 1-3.6 2.7-.2 1.6-1 2.8-2.8 3.7 2.6.9 6 .4 7.4-1.4 1-1.3 1-3.2-1-5Z" fill="#4a90d9" stroke="#2b5797" stroke-width=".8"/></svg>',
  funnel: '<svg width="20" height="20" viewBox="0 0 20 20"><path d="M2.5 3h15l-6 7.5V17l-3-1.6v-4.9Z" fill="none" stroke="#444" stroke-width="1.2" stroke-linejoin="round"/></svg>',
  arrow: '<svg width="6" height="9" viewBox="0 0 6 9"><path d="M1 .8 4.8 4.5 1 8.2" fill="none" stroke="#444" stroke-width="1.1"/></svg>',
  chevron: '<svg width="9" height="6" viewBox="0 0 9 6"><path d="M.8 1 4.5 4.8 8.2 1" fill="none" stroke="#444" stroke-width="1.2"/></svg>',
  chevronR: '<svg width="6" height="9" viewBox="0 0 6 9"><path d="M1 .8 4.8 4.5 1 8.2" fill="none" stroke="#444" stroke-width="1.2"/></svg>',
  check: '<svg width="12" height="10" viewBox="0 0 12 10"><path d="M1 5.2 4.3 8.5 11 1.5" fill="none" stroke="#262626" stroke-width="1.3"/></svg>',
  dot: '<svg width="8" height="8" viewBox="0 0 8 8"><circle cx="4" cy="4" r="3" fill="#262626"/></svg>',
};

const svgEl = (markup: string) => { const s = h("span"); s.innerHTML = markup; return s.firstElementChild as SVGElement; };

type Which = "elements" | "styles" | "filter";

export function chartButtons(host: ButtonsHost) {
  let open: { which: Which; m: ChartModel; el: HTMLElement; btn: HTMLElement; sub?: HTMLElement } | null = null;
  let away: ((e: PointerEvent) => void) | null = null;
  let keyAway: ((e: KeyboardEvent) => void) | null = null;

  if (!document.getElementById("raabe-chart-buttons-css")) document.head.append(h("style", { id: "raabe-chart-buttons-css" }, CSS));

  /** Adds the three buttons to a chart's element (once). */
  function attach(view: HTMLElement, m: () => ChartModel): void {
    if (view.querySelector(".chart-btn")) return;
    const defs: [Which, string, string, string][] = [
      ["elements", "Diagrammelemente", "Element hinzufügen", ICON.plus],
      ["styles", "Diagrammformatvorlagen", "Formatvorlagen", ICON.brush],
      ["filter", "Diagrammfilter", "Datenlayout", ICON.funnel],
    ];
    defs.forEach(([which, tip, _uia, icon], i) => {
      void _uia;
      const b = h("button", { class: "chart-btn", "data-which": which, "aria-label": tip, "data-tip": tip, tabindex: "-1" });
      b.append(svgEl(icon));
      b.style.top = `${i * PITCH}px`;
      b.style.right = `${-(GAP_X + BTN)}px`;
      b.addEventListener("pointerdown", (e) => { e.stopPropagation(); e.preventDefault(); });
      b.addEventListener("dblclick", (e) => e.stopPropagation());
      b.addEventListener("click", (e) => {
        e.stopPropagation();
        if (open?.which === which && open.m === m()) close();
        else show(which, m(), b);
      });
      view.append(b);
    });
  }

  function close(): void {
    if (!open) return;
    host.preview(open.m, null);
    host.highlight(open.m, null);
    open.sub?.remove();
    open.el.remove();
    open.btn.classList.remove("on");
    open = null;
    if (away) removeEventListener("pointerdown", away, true);
    if (keyAway) removeEventListener("keydown", keyAway, true);
    away = keyAway = null;
  }

  function show(which: Which, m: ChartModel, btn: HTMLElement): void {
    close();
    const el = which === "elements" ? elementsFlyout(m) : which === "styles" ? stylesFlyout(m) : filterFlyout(m);
    el.classList.add("cb-fly");
    el.dataset.which = which;
    el.addEventListener("pointerdown", (e) => e.stopPropagation());
    document.body.append(el);
    btn.classList.add("on");
    open = { which, m, el, btn };
    place();
    away = (e: PointerEvent) => {
      const t = e.target as Node;
      if (open?.el.contains(t) || open?.sub?.contains(t) || (t as HTMLElement).closest?.(".chart-btn")) return;
      // As Office's menus (D-156): a left press in the document is swallowed, the frame and a right press act.
      pressOutside(e, close);
    };
    keyAway = (e: KeyboardEvent) => {
      if (e.key !== "Escape" || !open) return;
      e.preventDefault();
      e.stopPropagation();
      // Esc closes a side menu first, then the flyout (Excel).
      if (open.sub) { open.sub.remove(); open.sub = undefined; open.el.querySelector(".cb-arrow.open")?.classList.remove("open"); return; }
      const b = open.btn;
      close();
      (b.closest(".chart-obj") as HTMLElement | null)?.focus();
    };
    addEventListener("pointerdown", away, true);
    addEventListener("keydown", keyAway, true);
  }

  /** Beside the button, the notch at the button's middle; to the left of the chart when there is no room. */
  function place(): void {
    if (!open) return;
    const r = open.btn.getBoundingClientRect();
    const fw = open.el.offsetWidth, fh = open.el.offsetHeight;
    let left = r.right + 8;
    open.el.classList.toggle("left", left + fw > innerWidth - 4);
    if (left + fw > innerWidth - 4) left = Math.max(4, (open.btn.closest(".chart-obj")?.getBoundingClientRect().left ?? r.left) - fw - 8);
    const top = Math.max(4, Math.min(r.top - 6, innerHeight - fh - 4));
    open.el.style.left = `${left}px`;
    open.el.style.top = `${top}px`;
  }

  /** Redraws the open flyout after a change (the chart model may be a new object after undo steps). */
  function refresh(m: ChartModel): void {
    if (!open) return;
    const { which, btn } = open;
    const scroll = open.el.querySelector(".cb-scroll")?.scrollTop ?? 0;
    const tab = open.el.querySelector(".cb-tab.sel")?.textContent;
    show(which, m, btn);
    if (tab && open) (Array.from(open.el.querySelectorAll<HTMLElement>(".cb-tab")).find((t) => t.textContent === tab))?.click();
    const sc = open?.el.querySelector(".cb-scroll");
    if (sc) sc.scrollTop = scroll;
  }

  // ----- Diagrammelemente ------------------------------------------------------------------------

  function elementsFlyout(m: ChartModel): HTMLElement {
    const list = h("div", { class: "cb-list", role: "list" });
    for (const e of elementEntries(m).filter((x) => !x.ribbonOnly && !x.disabled)) {
      const box = h("input", { type: "checkbox", tabindex: "-1" }) as HTMLInputElement;
      box.checked = e.on(m);
      const row = h("div", { class: `cb-row${box.checked ? " on" : ""}`, role: "listitem", "data-id": e.id });
      const label = h("label", null, box, e.label);
      box.addEventListener("change", () => void toggle(m, e, box.checked));
      const arrow = h("button", { class: "cb-arrow", "aria-label": `${e.label} Optionen`, tabindex: "-1" });
      arrow.append(svgEl(ICON.arrow));
      arrow.addEventListener("click", (ev) => { ev.stopPropagation(); sideMenu(m, e, arrow); });
      row.append(label, arrow);
      list.append(row);
    }
    return h("div", { role: "dialog", "aria-label": "Diagrammelemente", style: "width:189px" }, h("div", { class: "cb-head" }, "Diagrammelemente"), list);
  }

  async function toggle(m: ChartModel, e: ElementEntry, on: boolean): Promise<void> {
    if (e.id === "trendline" && on && m.series.filter((s) => !s.filtered).length > 1) {
      const i = await host.pickSeries(m);
      if (i === null) { refresh(m); return; }
      host.edit(m, (c) => { c.series[i].trendline = { type: "linear", showEq: false, showR2: false }; });
      return;
    }
    host.edit(m, (c) => e.toggle(c, on));
  }

  function sideMenu(m: ChartModel, e: ElementEntry, arrow: HTMLElement): void {
    if (!open) return;
    const wasOpen = arrow.classList.contains("open");
    open.sub?.remove();
    open.sub = undefined;
    open.el.querySelectorAll(".cb-arrow.open").forEach((a) => a.classList.remove("open"));
    if (wasOpen) return;
    const sub = h("div", { class: "cb-sub", role: "menu", "aria-label": e.label });
    for (const c of e.choices) {
      const on = c.on(m);
      const mark = h("span", { class: "mark" });
      if (on) mark.append(svgEl(c.check ? ICON.check : ICON.dot));
      const b = h("button", { role: c.check ? "menuitemcheckbox" : "menuitemradio", "aria-checked": String(on) }, mark, choiceLabel(c, false));
      b.addEventListener("click", () => void choose(m, e, c.check ? !on : true, e.choices.indexOf(c)));
      sub.append(b);
    }
    sub.append(h("hr"));
    const more = h("button", { role: "menuitem" }, h("span", { class: "mark" }), "Weitere Optionen...");
    more.addEventListener("click", () => { const pane = e.pane(m); close(); host.openPane(m, pane); });
    sub.append(more);
    sub.addEventListener("pointerdown", (ev) => ev.stopPropagation());
    document.body.append(sub);
    const r = arrow.getBoundingClientRect();
    let left = r.right + 2;
    if (left + sub.offsetWidth > innerWidth - 4) left = Math.max(4, (open.el.getBoundingClientRect().left) - sub.offsetWidth - 2);
    sub.style.left = `${left}px`;
    sub.style.top = `${Math.max(4, Math.min(r.top - 5, innerHeight - sub.offsetHeight - 4))}px`;
    arrow.classList.add("open");
    open.sub = sub;
  }

  async function choose(m: ChartModel, e: ElementEntry, on: boolean, i: number): Promise<void> {
    const c = e.choices[i];
    if (e.id === "trendline" && m.series.filter((s) => !s.filtered).length > 1) {
      const si = await host.pickSeries(m);
      if (si === null) return;
      host.edit(m, (x) => {
        const one = { ...x, series: [x.series[si]] } as ChartModel;
        c.set(one, true);
      });
      return;
    }
    host.edit(m, (x) => c.set(x, on));
  }

  // ----- Diagrammformatvorlagen ----------------------------------------------------------------------

  function stylesFlyout(m: ChartModel): HTMLElement {
    const tabs = h("div", { class: "cb-tabs", role: "tablist" });
    const body = h("div", { class: "cb-scroll", style: "height:338px" });
    const mk = (label: string, fill: () => void) => {
      const t = h("button", { class: "cb-tab", role: "tab" }, label);
      t.addEventListener("click", () => { tabs.querySelectorAll(".cb-tab").forEach((x) => x.classList.toggle("sel", x === t)); body.scrollTop = 0; fill(); });
      tabs.append(t);
      return t;
    };
    const first = mk("Formatvorlage", () => fillStyles(m, body));
    mk("Farbe", () => fillColors(m, body));
    first.classList.add("sel");
    fillStyles(m, body);
    return h("div", { role: "dialog", "aria-label": "Diagrammformatvorlagen", style: "width:225px" }, tabs, body);
  }

  function fillStyles(m: ChartModel, body: HTMLElement): void {
    body.replaceChildren();
    const styles = chartStyles(m);
    const cur = styleIndex(m);
    styles.forEach((_, i) => {
      const look = structuredClone(m);
      applyStyle(look, i);
      const tile = h("button", { class: `cb-tile${cur === i ? " sel" : ""}`, role: "option", "aria-label": `Formatvorlage ${i + 1}`, "data-tip": `Formatvorlage ${i + 1}` });
      tile.append(host.tile(look, 186, 113));
      tile.addEventListener("mouseenter", () => host.preview(m, look));
      tile.addEventListener("mouseleave", () => host.preview(m, null));
      tile.addEventListener("click", () => { host.preview(m, null); host.edit(m, (c) => applyStyle(c, i), "Diagrammformatvorlage"); });
      body.append(tile);
    });
  }

  function fillColors(m: ChartModel, body: HTMLElement): void {
    body.replaceChildren();
    const accents = host.accents();
    const theme: Record<string, string> = { dk1: "#000000", lt1: "#FFFFFF" };
    accents.forEach((a, k) => { theme[`accent${k + 1}`] = a; });
    const row = (id: number, label: string) => {
      const sw = (n: number) => Array.from({ length: n }, (_, i) => h("span", { style: `background:${id === 10 ? paletteColor({ ...m, colors: 10 }, i, n, accents) : schemeColor(id, i, n, theme)}` }));
      const el = h("button", { class: `cb-pal${(m.colors ?? 10) === id ? " sel" : ""}`, role: "option", "aria-label": label, "data-tip": label, style: "width:100%;border-radius:0;background:none" }, ...sw(6));
      const look = { ...structuredClone(m), colors: id === 10 ? undefined : id };
      for (const s of look.series) delete s.color;
      el.addEventListener("mouseenter", () => host.preview(m, look));
      el.addEventListener("mouseleave", () => host.preview(m, null));
      el.addEventListener("click", () => { host.preview(m, null); host.edit(m, (c) => { c.colors = id === 10 ? undefined : id; for (const s of c.series) delete s.color; }, "Farben ändern"); });
      return el;
    };
    body.append(h("div", { class: "cb-pal-head" }, "Bunt"));
    [10, 11, 12, 13].forEach((id, i) => body.append(row(id, `Farbige Palette ${i + 1}`)));
    body.append(h("div", { class: "cb-pal-head" }, "Einfarbig"));
    for (let i = 0; i < 13; i++) body.append(row(14 + i, `Monochrome Palette ${i + 1}`));
  }

  // ----- Diagrammfilter ----------------------------------------------------------------------------

  function filterFlyout(m: ChartModel): HTMLElement {
    const tabs = h("div", { class: "cb-tabs", role: "tablist" });
    const body = h("div", { class: "cb-scroll", style: "height:330px;padding:2px 6px" });
    const apply = h("button", { class: "apply", disabled: true }, "Anwenden") as HTMLButtonElement;
    const link = h("a", { role: "link", tabindex: "0" }, "Daten auswählen...");
    link.addEventListener("click", () => { close(); host.selectData(m); });
    // Drafts: nothing changes before Anwenden (Excel).
    const series = m.series.map((s) => !s.filtered);
    const cats = catNames(m).map((_, i) => !(m.hiddenCats ?? []).includes(i));
    const src = host.nameSources(m);
    let names = { series: src.series !== null, cats: src.cats !== null };
    const dirty = () => {
      const changed = series.some((on, i) => on === !!m.series[i].filtered) || cats.some((on, i) => on === (m.hiddenCats ?? []).includes(i))
        || names.series !== (src.series !== null) || names.cats !== (src.cats !== null);
      apply.disabled = !changed;
    };
    apply.addEventListener("click", () => {
      const hidden = cats.map((on, i) => (on ? -1 : i)).filter((i) => i >= 0);
      const nameChange = names.series !== (src.series !== null) || names.cats !== (src.cats !== null);
      host.highlight(m, null);
      host.edit(m, (c) => {
        c.series.forEach((s, i) => { if (series[i]) delete s.filtered; else s.filtered = true; });
        if (hidden.length) c.hiddenCats = hidden; else delete c.hiddenCats;
      }, "Diagrammfilter");
      if (nameChange) host.setNameSources(m, names.series, names.cats);
    });
    const werte = () => {
      body.replaceChildren();
      const group = (title: string, items: { label: string; color?: string }[], state: boolean[], hover: (i: number | null) => void) => {
        const list = h("div");
        const head = h("button", { class: "cb-group-head", "aria-expanded": "true" });
        const chev = svgEl(ICON.chevron);
        head.append(chev, title);
        head.addEventListener("click", () => {
          const shut = list.hidden = !list.hidden;
          head.setAttribute("aria-expanded", String(!shut));
          chev.replaceWith(svgEl(shut ? ICON.chevronR : ICON.chevron));
        });
        const all = h("input", { type: "checkbox", tabindex: "-1" }) as HTMLInputElement;
        const boxes: HTMLInputElement[] = [];
        const sync = () => { const n = state.filter(Boolean).length; all.checked = n === state.length; all.indeterminate = n > 0 && n < state.length; };
        all.addEventListener("change", () => { state.fill(all.checked); boxes.forEach((b) => { b.checked = all.checked; }); sync(); dirty(); });
        list.append(h("label", { class: "cb-filter-row" }, all, h("span", { class: "lbl" }, "(Alle auswählen)")));
        items.forEach((it, i) => {
          const b = h("input", { type: "checkbox", tabindex: "-1" }) as HTMLInputElement;
          b.checked = state[i];
          b.addEventListener("change", () => { state[i] = b.checked; sync(); dirty(); });
          boxes.push(b);
          const row = h("label", { class: "cb-filter-row", "data-index": String(i) }, b,
            ...(it.color ? [h("span", { class: "cb-swatch", style: `background:${it.color}` })] : []), h("span", { class: "lbl", title: it.label }, it.label));
          row.addEventListener("mouseenter", () => hover(i));
          row.addEventListener("mouseleave", () => hover(null));
          list.append(row);
        });
        sync();
        return h("div", { role: "group", "aria-label": title }, head, list);
      };
      const accents = host.accents();
      const colors = m.series.map((s, i) => s.color ?? paletteColor(m, i, m.series.length, accents));
      body.append(group("Datenreihe", m.series.map((s, i) => ({ label: seriesName(s, i), color: colors[i] })), series,
        (i) => host.highlight(m, i === null ? null : { series: i })));
      if (!isPie(m.kind) || true) {
        body.append(group("Kategorien", catNames(m).map((label) => ({ label })), cats, (i) => host.highlight(m, i === null ? null : { cat: i })));
      }
    };
    const namen = () => {
      body.replaceChildren();
      const radios = (title: string, from: string | null, list: string, key: "series" | "cats") => {
        const nm = `cbn-${key}`;
        const r1 = h("input", { type: "radio", name: nm, tabindex: "-1" }) as HTMLInputElement;
        const r2 = h("input", { type: "radio", name: nm, tabindex: "-1" }) as HTMLInputElement;
        r1.checked = names[key];
        r2.checked = !names[key];
        r1.disabled = from === null;
        r1.addEventListener("change", () => { names = { ...names, [key]: true }; dirty(); });
        r2.addEventListener("change", () => { names = { ...names, [key]: false }; dirty(); });
        return h("div", { role: "group", "aria-label": title },
          h("div", { class: "cb-group-head" }, svgEl(ICON.chevron), title),
          h("label", { class: "cb-radio" }, r1, h("span", null, from ?? "", h("small", null, list))),
          h("label", { class: "cb-radio" }, r2, h("span", null, "(Keine)")));
      };
      body.append(radios("Datenreihe", src.series, src.seriesList, "series"), radios("Kategorien", src.cats, src.catList, "cats"));
    };
    const mk = (label: string, fill: () => void) => {
      const t = h("button", { class: "cb-tab", role: "tab" }, label);
      t.addEventListener("click", () => { tabs.querySelectorAll(".cb-tab").forEach((x) => x.classList.toggle("sel", x === t)); fill(); });
      tabs.append(t);
      return t;
    };
    mk("Werte", werte).classList.add("sel");
    mk("Namen", namen);
    werte();
    return h("div", { role: "dialog", "aria-label": "Diagrammfilter", style: "width:222px" }, tabs, body, h("div", { class: "cb-foot" }, apply, link));
  }

  return {
    attach,
    close,
    refresh,
    isOpen: () => !!open,
    which: (): Which | null => open?.which ?? null,
    openFor: (): ChartModel | null => open?.m ?? null,
    /** The flyout of a button, for the ribbon or a test (as a click on the button). */
    show: (which: Which, m: ChartModel, view: HTMLElement) => {
      const b = view.querySelector<HTMLElement>(`.chart-btn[data-which="${which}"]`);
      if (b) show(which, m, b);
    },
  };
}

/** The category names of the full range (filtered ones included), as the filter lists them. */
export function catNames(m: ChartModel): string[] {
  const s = m.series[0];
  if (!s) return [];
  const n = Math.max(s.val.cache.length, s.cat?.cache.length ?? 0);
  const fam = kindInfo(m.kind).family;
  void fam;
  return Array.from({ length: n }, (_, i) => {
    const c = s.cat?.cache[i];
    return c === null || c === undefined ? String(i + 1) : String(c);
  });
}

/** For tests: whether a view of the chart drops what the filter hides. */
export const visibleSeriesCount = (m: ChartModel) => visibleModel(m).model.series.length;
