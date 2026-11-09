// Everything Word 365 shows around a selected chart (D-129, D-130), with the chart UI RaabeCalc already has
// (@raabe/chart-ui): the contextual tabs Diagrammentwurf (Alt J C) and Format (Alt J Y) with the groups, controls and
// keytips recorded in our own Word (probe_word_chart_ui.py, results/2026-11-09-word-chart-ui.json); the buttons beside
// the chart (Layoutoptionen from the pictures, then Diagrammelemente, Diagrammformatvorlagen and Diagrammfilter as in
// RaabeCalc); the Format pane of the chart's elements. A chart in the text keeps its chart part as XML: each change
// parses it once (cached), changes the model, writes the part again and draws the new picture, as one undo step.

import {
  applyColors, applyQuickLayout, applyStyle, chartButtons, chartElements, chartStyles, chartTypeDialog, choiceLabel, elementEntries,
  elementName, formatPane, layoutCount, setStyleTheme, styleIndex, type ChartElement, type ElementChoice,
} from "@raabe/chart-ui";
import { renderChart, seriesName, visibleModel, type ChartKind, type ChartModel } from "@raabe/charts";
import { DEFAULT_THEME_COLORS, chartFromXml, chartImage, chartSvg, chartToXml } from "@raabe/ooxml";
import { TaskPane, h, showDialog, type Command, type Control, type MenuEntry, type Tab } from "@raabe/shell";
import type { Dock } from "./dock";
import { CM, type AppApi } from "./api";
import { GROUP_MENU } from "./groups";
import { tableOf, workbookOf } from "./charts";

const cmd = (id: string, label: string, icon?: string, extra: Partial<Command> = {}): Command => ({ id, label, icon, ...extra });
const later = (id: string, label: string, icon?: string): Command => cmd(id, label, icon, { deferred: true });

/** The tabs' commands with Word's labels (results/2026-11-09-word-chart-ui.json). */
export const CHART_TOOLS_COMMANDS: Command[] = [
  // Diagrammentwurf
  cmd("wchart.addElement", "Diagrammelement hinzufügen", "chart-recommended"),
  cmd("wchart.quickLayout", "Schnelllayout", "chart-recommended"),
  cmd("wchart.colors", "Farben ändern", "colors"),
  cmd("wchart.styles", "Schnellformatvorlagen", "styles"),
  cmd("wchart.switchRowCol", "Zeile/Spalte tauschen", "table-switch"),
  cmd("wchart.selectData", "Daten auswählen...", "table"),
  cmd("wchart.editData", "Daten bearbeiten", "table"),
  cmd("wchart.refreshData", "Daten aktualisieren", "reapply"),
  cmd("wchart.changeType", "Diagrammtyp ändern...", "chart-column"),
  // Format
  cmd("wchart.element", "Diagrammelemente", "chart-recommended"),
  cmd("wchart.formatSelection", "Auswahl formatieren", "format-picture", { shortcut: "Strg+1" }),
  cmd("wchart.resetStyle", "Auf Formatvorl. zurücks.", "reset-picture"),
  later("wchart.insertShape", "Formen", "shapes"),
  later("wchart.changeShape", "Form ändern", "shapes"),
  later("wchart.shapeStyles", "Schnellformatvorlagen", "styles"),
  cmd("wchart.fill", "Fülleffekt", "shape-fill"),
  cmd("wchart.outline", "Formkontur", "shape-outline"),
  later("wchart.shapeEffects", "Formeffekte", "effects"),
  cmd("wchart.formatObject", "Objekt formatieren...", "format-picture"),
  later("wchart.wordArtStyles", "Schnellformatvorlagen", "styles"),
  cmd("wchart.textFill", "Textfüllung", "font-color"),
  later("wchart.textOutline", "Textkontur", "shape-outline"),
  later("wchart.textEffects", "Texteffekte", "effects"),
  cmd("wchart.formatText", "Texteffekte formatieren...", "format-picture"),
  cmd("wchart.height", "Formenhöhe", "shape-height"),
  cmd("wchart.width", "Formenbreite", "shape-width"),
  cmd("wchart.sizeDialog", "Größe und Position...", "size-position"),
];

/** Diagrammelemente and the colour scheme names as Office has them. */
const COLOR_SETS: [number, string][] = [[10, "Farbige Palette 1"], [11, "Farbige Palette 2"], [12, "Farbige Palette 3"], [13, "Farbige Palette 4"],
  ...Array.from({ length: 13 }, (_, i): [number, string] => [14 + i, `Monochrome Palette ${i + 1}`])];

/** Word's chart buttons in Word's blue instead of Excel's green (the rest as RaabeCalc draws them). */
const CSS = `
.wchart-overlay { position: fixed; pointer-events: none; z-index: 6; }
.wchart-overlay .chart-btn { pointer-events: auto; }
.wchart-overlay > svg { position: absolute; left: 0; top: 0; opacity: 0; pointer-events: none; }
.wchart-overlay .chart-btn.on { background: #dbe7f6; border-color: #2b579a; }
.cb-fly .cb-head { color: #2b579a; }
.cb-fly .cb-tab.sel { border-bottom-color: #2b579a; }
.cb-fly .cb-foot a { color: #2b579a; }
.cb-fly .cb-filter-row input { accent-color: #2b579a; }
.docarea > .taskpane { position: absolute; top: 0; bottom: 0; }
`;

export interface ChartToolsDeps {
  /** Daten bearbeiten (the chart's table, D-091). */
  editData(): Promise<void> | void;
  /** Bildformat's Anordnen: Position, Textumbruch, Ausrichten, Drehen (shared with Formformat). */
  positionPanel: (close: () => void) => HTMLElement;
  alignPanel: (close: () => void) => HTMLElement;
  rotateMenu: MenuEntry[];
  wrapMenu: MenuEntry[];
}

export function chartToolsFeature(api: AppApi, docArea: HTMLElement, dock: Dock, deps: ChartToolsDeps) {
  const c = api.commands;
  const ed = () => api.editor();
  if (!document.getElementById("raabe-wchart-css")) document.head.append(h("style", { id: "raabe-wchart-css" }, CSS));

  // ----- the selected chart and its model ------------------------------------------------------------------
  type Chart = { readonly xml: string; readonly workbook?: Uint8Array };
  const models = new WeakMap<Chart, ChartModel>();
  const sel = () => { const s = ed()?.selectedImage(); return s && s.obj.chart ? s : null; };
  function modelOf(ch: Chart): ChartModel {
    let m = models.get(ch);
    if (!m) { m = chartFromXml(ch.xml); models.set(ch, m); }
    return m;
  }
  /** The selected chart's model (the same object while the chart is not changed elsewhere). */
  const selected = (): ChartModel | null => { const s = sel(); return s ? modelOf(s.obj.chart!) : null; };
  const theme = (): Record<string, string> => {
    const t: Record<string, string> = { ...DEFAULT_THEME_COLORS, ...(ed()?.doc.styles.theme.colors ?? {}) };
    return Object.fromEntries(Object.entries(t).map(([k, v]) => [k, "#" + String(v).replace("#", "")]));
  };
  const accents = () => { const t = theme(); return [1, 2, 3, 4, 5, 6].map((i) => t[`accent${i}`]); };
  const sizePt = () => { const s = sel(); return s ? { w: s.obj.width ?? 432, h: s.obj.height ?? 252 } : { w: 432, h: 252 }; };

  /** The chart's new part and picture after `m` changed (the embedded workbook stays, or is made from the part's data). */
  function commit(m: ChartModel, snapshot = true): void {
    const e = ed(), s = sel();
    if (!e || !s?.obj.chart) return;
    const old = s.obj.chart;
    const workbook = old.workbook ?? workbookOf(tableOf(old.xml));
    const next: Chart = { xml: chartToXml(m), workbook };
    models.delete(old);
    models.set(next, m);
    const { w, h: hh } = sizePt();
    e.setSelectedChart(next, chartImage(m, w, hh, accents()), snapshot);
    drawnSize = `${w}x${hh}`;
  }

  /** Changes the selected chart as one undo step (Livevorschau ends first). */
  function edit(m: ChartModel, fn: (m: ChartModel) => void): void {
    const e = ed();
    if (!e) return;
    e.preview(null);
    const cur = selected();
    const target = cur ?? m;
    fn(target);
    commit(target);
    if (buttons.isOpen()) buttons.refresh(target);
    if (paneOpen) showPane(target);
    api.refresh();
  }

  /** Livevorschau: `look` drawn in place of the chart (null ends it). */
  function preview(look: ChartModel | null): void {
    const e = ed(), s = sel();
    if (!e) return;
    if (!look || !s?.obj.chart) { e.preview(null); return; }
    const ch = s.obj.chart;
    const { w, h: hh } = sizePt();
    const img = chartImage(look, w, hh, accents());
    e.preview(() => e.setSelectedChart(ch, img));
  }

  /** Diagrammfilter: all but one series or category faded, as the entry under the mouse shows it. */
  function highlight(m: ChartModel, what: { series?: number; cat?: number } | null): void {
    const e = ed(), s = sel();
    if (!e || !s?.obj.chart) return;
    if (!what) { e.preview(null); return; }
    const { w, h: hh } = sizePt();
    const doc = new DOMParser().parseFromString(chartSvg(m, w, hh, accents()), "image/svg+xml");
    const { index } = visibleModel(m);
    const hiddenBefore = (cat: number) => (m.hiddenCats ?? []).filter((x) => x < cat).length;
    for (const el of Array.from(doc.querySelectorAll('[data-part^="series:"]'))) {
      const [, si, pi] = (el.getAttribute("data-part") ?? "").split(":");
      const fade = (what.series !== undefined && index[Number(si)] !== what.series)
        || (what.cat !== undefined && pi !== undefined && Number(pi) !== what.cat - hiddenBefore(what.cat));
      if (fade) el.setAttribute("opacity", "0.25");
    }
    const bytes = new TextEncoder().encode(new XMLSerializer().serializeToString(doc));
    const ch = s.obj.chart;
    e.preview(() => e.setSelectedChart(ch, { bytes, type: "image/svg+xml" }));
  }

  /** A small drawing of the chart (style tiles). */
  function tile(m: ChartModel, w: number, hh: number): SVGSVGElement {
    const { w: cw, h: ch } = sizePt();
    const pw = (cw * 96) / 72, ph = (ch * 96) / 72;
    const svg = renderChart(m, pw, ph, { accents: accents() }) as SVGSVGElement;
    const k = Math.min(w / pw, hh / ph);
    svg.setAttribute("width", String(Math.round(pw * k)));
    svg.setAttribute("height", String(Math.round(ph * k)));
    return svg;
  }

  async function pickSeries(m: ChartModel): Promise<number | null> {
    const list = h("select", { size: 6, style: "width:260px", "aria-label": "Reihen" }) as HTMLSelectElement;
    m.series.forEach((s, i) => { if (!s.filtered) list.append(h("option", { value: String(i) }, seriesName(s, i))); });
    list.selectedIndex = 0;
    const ok = await showDialog({ title: "Trendlinie hinzufügen", body: h("div", null, h("div", { style: "margin-bottom:6px" }, "Trendlinie hinzufügen basierend auf Reihe:"), list),
      buttons: [{ label: "OK", value: true, primary: true }, { label: "Abbrechen", value: false }], cancel: false });
    return ok && list.value !== "" ? Number(list.value) : null;
  }

  // ----- the buttons beside the chart -----------------------------------------------------------------------
  const buttons = chartButtons({
    edit: (m, fn) => edit(m, fn),
    preview: (_m, look) => preview(look),
    highlight: (m, what) => highlight(m, what),
    openPane: (m, e) => { element = e; showPane(m); },
    selectData: () => void deps.editData(),
    accents,
    tile,
    pickSeries,
    // Word's data sheet: names in row 1, categories in column A (D-091).
    nameSources: (m) => ({
      series: m.series.some((s) => s.name) ? "Zeile 1" : null, cats: m.series[0]?.cat ? "Spalte A" : null,
      seriesList: m.series.map((s, i) => seriesName(s, i)).join("; "),
      catList: (m.series[0]?.cat?.cache ?? []).map((v) => (v === null ? "" : String(v))).join("; "),
    }),
    setNameSources: (m, series, cats) => {
      const t = tableOf(sel()?.obj.chart?.xml ?? "");
      edit(m, (x) => x.series.forEach((s, i) => {
        if (series) s.name = { ref: `Tabelle1!$${String.fromCharCode(66 + i)}$1`, text: t.series[i]?.name ?? `Datenreihe ${i + 1}` }; else delete s.name;
        if (cats) s.cat = { ref: `Tabelle1!$A$2:$A$${t.categories.length + 1}`, cache: [...t.categories] }; else delete s.cat;
      }));
    },
  });

  // An overlay over the selected chart holds the three buttons (below Word's layout options button) and an invisible
  // drawing of the chart that tells which element a click is on.
  const overlay = h("div", { class: "wchart-overlay chart-obj sel", hidden: true }) as HTMLDivElement;
  document.body.append(overlay);
  buttons.attach(overlay, () => selected()!);
  let overlayKey = "";
  let drawnSize = "";
  let element: ChartElement = { kind: "chart" };
  let lastChart: Chart | null = null;

  function sync(): void {
    const e = ed();
    const s = sel();
    const r = s ? e?.view.selectedObjectRect() : null;
    if (!e || !s || !r || r.width <= 0) {
      overlay.hidden = true;
      if (buttons.isOpen() && !e?.previewing) buttons.close();
      lastChart = null;
      return;
    }
    if (s.obj.chart !== lastChart && !e.previewing) {
      // Another chart, or this one changed elsewhere (undo): the element choice starts over unless it still exists.
      const m = modelOf(s.obj.chart!);
      if (!chartElements(m).some((x) => JSON.stringify(x) === JSON.stringify(element))) element = { kind: "chart" };
      lastChart = s.obj.chart!;
    }
    overlay.hidden = false;
    startFollowing();
    // The layout options button (22 px) stands at the top; the chart buttons follow below it, 38 px apart.
    Object.assign(overlay.style, { left: `${r.left}px`, top: `${r.top + 38}px`, width: `${r.width}px`, height: `${Math.max(0, r.height - 38)}px` });
    const m = modelOf(s.obj.chart!);
    const key = `${s.obj.chart!.xml.length}:${r.width}x${r.height}:${models.get(s.obj.chart!) === m}`;
    if (key !== overlayKey && !e.previewing) {
      overlay.querySelector(":scope > svg")?.remove();
      const svg = renderChart(m, r.width, r.height, { accents: accents() });
      svg.style.top = "-38px";
      overlay.prepend(svg);
      overlayKey = key;
    }
    // A chart sized with the mouse is drawn again at its new size (Word lays the chart out anew), in the same step.
    const { w, h: hh } = sizePt();
    const size = `${w}x${hh}`;
    const img = s.obj.image;
    const svgSize = img ? /width="([\d.]+)" height="([\d.]+)"/.exec(new TextDecoder().decode(img.bytes.slice(0, 200))) : null;
    const drawn = svgSize ? `${Math.round((Number(svgSize[1]) * 72) / 96 * 10) / 10}x${Math.round((Number(svgSize[2]) * 72) / 96 * 10) / 10}` : "";
    const want = `${Math.round(w * 10) / 10}x${Math.round(hh * 10) / 10}`;
    if (!e.previewing && drawn && drawn !== want && size !== drawnSize) {
      drawnSize = size;
      queueMicrotask(() => { const ch = sel()?.obj.chart; if (ch) e.setSelectedChart(ch, chartImage(modelOf(ch), w, hh, accents()), false); });
    }
  }
  // While a chart is selected the buttons follow it every frame (panes opening, rulers, zoom move the page).
  let following = false;
  const follow = () => {
    if (overlay.hidden) { following = false; return; }
    const r = ed()?.view.selectedObjectRect();
    const o = overlay.getBoundingClientRect();
    if (r && (Math.abs(r.left - o.left) > 0.5 || Math.abs(r.top + 38 - o.top) > 0.5 || Math.abs(r.width - o.width) > 0.5)) sync();
    requestAnimationFrame(follow);
  };
  const startFollowing = () => { if (!following && !overlay.hidden) { following = true; requestAnimationFrame(follow); } };
  addEventListener("scroll", () => { if (!overlay.hidden) sync(); }, true);
  addEventListener("resize", () => sync());

  /** The chart element under a client point (from the invisible drawing's data-part marks). */
  function elementAtPoint(x: number, y: number): ChartElement | null {
    const svg = overlay.querySelector<SVGSVGElement>(":scope > svg");
    const m = selected();
    if (!svg || !m || overlay.hidden) return null;
    svg.style.pointerEvents = "auto";
    const hit = document.elementFromPoint(x, y);
    svg.style.pointerEvents = "";
    if (!hit || !svg.contains(hit)) return null;
    const part = (hit.closest("[data-part]") as SVGElement | null)?.dataset.part;
    if (!part) return { kind: "chart" };
    const [kind, a] = part.split(":");
    const full = (i: number) => visibleModel(m).index[i] ?? i;
    switch (kind) {
      case "plot": return { kind: "plot" };
      case "title": return { kind: "title" };
      case "legend": return { kind: "legend" };
      case "axis": return a === "val" ? { kind: "valAxis" } : { kind: "catAxis" };
      case "axes": return { kind: "catAxis" };
      case "grid": return { kind: "majorGrid" };
      case "minorGrid": return { kind: "minorGrid" };
      case "catGrid": return { kind: "catGrid" };
      case "axisTitle": return { kind: "axisTitle", axis: a === "val" ? "val" : "cat" };
      case "series": return { kind: "series", index: full(Number(a)) };
      case "labels": return { kind: "labels" };
      case "trend": case "trendLabel": return { kind: "trend", index: full(Number(a)) };
      case "errbars": return { kind: "errorBars", index: full(Number(a)) };
      default: return { kind: "chart" };
    }
  }
  const inChart = (x: number, y: number) => {
    const r = sel() ? ed()?.view.selectedObjectRect() : null;
    return !!r && x >= r.left && x <= r.right && y >= r.top && y <= r.bottom;
  };
  // A click on a selected chart chooses the element under it (Aktuelle Auswahl); a double-click opens its Format pane.
  addEventListener("pointerdown", (ev) => {
    if (ev.button !== 0 || !inChart(ev.clientX, ev.clientY) || (ev.target as HTMLElement).closest?.(".chart-btn")) return;
    const at = elementAtPoint(ev.clientX, ev.clientY);
    if (at) { element = at; if (paneOpen) showPane(selected()!); api.refresh(); }
  }, true);
  addEventListener("dblclick", (ev) => {
    if (!inChart(ev.clientX, ev.clientY) || (ev.target as HTMLElement).closest?.(".chart-btn")) return;
    ev.preventDefault();
    ev.stopPropagation();
    element = elementAtPoint(ev.clientX, ev.clientY) ?? element;
    const m = selected();
    if (m) showPane(m);
  }, true);

  // ----- the Format pane ---------------------------------------------------------------------------------
  const pane = new TaskPane(docArea);
  dock.add(pane.el, "right", 300);
  let paneOpen = false;
  pane.onClose = () => { paneOpen = false; dock.layout(); };
  function showPane(m: ChartModel): void {
    if (!chartElements(m).some((x) => JSON.stringify(x) === JSON.stringify(element))) element = { kind: "chart" };
    pane.show(formatPane(m, element, {
      edit: (fn) => edit(m, fn),
      select: (x) => { element = x; showPane(selected() ?? m); },
      size: () => { const s = sizePt(); return { w: Math.round((s.w / CM) * 100) / 100, h: Math.round((s.h / CM) * 100) / 100 }; },
      resize: (w, hh) => { ed()?.setImageSize(w * CM, hh * CM); api.refresh(); },
    }));
    paneOpen = true;
    dock.layout();
  }

  // ----- ribbon: Diagrammentwurf and Format ----------------------------------------------------------------
  const on = () => !!sel();
  const run = (id: string, fn: () => void | Promise<void>, enabled: () => boolean = on) => c.implement(id, { run: async () => { await fn(); }, enabled });

  function elementsMenu(): MenuEntry[] {
    const m = selected();
    if (!m) return [];
    return elementEntries(m, true).map((e): MenuEntry => {
      if (e.disabled && !e.choices.length) return { label: e.label, submenu: [], disabled: true };
      return {
        label: e.label, disabled: e.disabled,
        submenu: () => {
          const cur = selected() ?? m;
          const items: MenuEntry[] = [];
          if (e.none) items.push({ label: "Keine", checked: !e.on(cur), run: () => edit(cur, (x) => e.toggle(x, false)) });
          for (const ch of e.choices) {
            const isOn = ch.on(cur);
            items.push({
              label: choiceLabel(ch, true), checked: isOn,
              hover: () => { const look = structuredClone(cur); ch.set(look, ch.check ? !isOn : true); preview(look); },
              run: () => { preview(null); void choose(cur, e.id, ch, isOn); },
            });
          }
          items.push("-", { label: e.more, run: () => { element = e.pane(cur); showPane(cur); } });
          return items;
        },
      };
    });
  }
  async function choose(m: ChartModel, id: string, ch: ElementChoice, isOn: boolean): Promise<void> {
    if (id === "trendline" && m.series.filter((s) => !s.filtered).length > 1) {
      const i = await pickSeries(m);
      if (i === null) return;
      edit(m, (x) => ch.set({ ...x, series: [x.series[i]] } as ChartModel, true));
      return;
    }
    edit(m, (x) => ch.set(x, ch.check ? !isOn : true));
  }

  const styleLook = (i: number) => { const m = selected(); if (!m) return null; setStyleTheme(theme()); const look = structuredClone(m); applyStyle(look, i); return look; };
  const layoutLook = (n: number) => { const m = selected(); if (!m) return null; const look = structuredClone(m); applyQuickLayout(look, n); return look; };
  const colorLook = (k: number) => { const m = selected(); if (!m) return null; const look = structuredClone(m); applyColors(look, k); return look; };

  function setElementColor(what: "fill" | "line" | "text", color: string | null): void {
    const m = selected();
    if (!m) return;
    const e = element;
    edit(m, (x) => {
      const f = (x.format ??= {});
      const target = ((): { fill?: unknown; line?: unknown; text?: { color?: string } } | null => {
        switch (e.kind) {
          case "series": return x.series[e.index] as never;
          case "chart": return (f.chart ??= {});
          case "plot": return (f.plot ??= {});
          case "title": return (f.title ??= {});
          case "legend": return (f.legend ??= {});
          case "valAxis": return (f.valAxis ??= {});
          case "catAxis": return (f.catAxis ??= {});
          case "majorGrid": return (f.majorGrid ??= {});
          case "labels": return (f.dataLabels ??= {});
          case "axisTitle": { const at = (f.axisTitles ??= {}); return (at[e.axis] ??= {}); }
          default: return null;
        }
      })();
      if (!target) return;
      if (what === "fill") target.fill = color === null ? { type: "none" } : { type: "solid", color };
      else if (what === "line") target.line = color === null ? { type: "none" } : { type: "solid", color };
      else target.text = { ...(target.text ?? {}), ...(color ? { color } : {}) };
    });
  }

  function resetElement(): void {
    const m = selected();
    if (!m) return;
    const e = element;
    edit(m, (x) => {
      if (e.kind === "series") { const s = x.series[e.index]; delete s.fill; delete s.line; delete s.marker; return; }
      if (e.kind === "chart") { delete x.format; for (const s of x.series) { delete s.fill; delete s.line; delete s.marker; } return; }
      const f = x.format;
      if (!f) return;
      const key = ({ plot: "plot", title: "title", legend: "legend", valAxis: "valAxis", catAxis: "catAxis", majorGrid: "majorGrid", minorGrid: "minorGrid", catGrid: "catGrid", labels: "dataLabels" } as const)[e.kind as "plot"];
      if (key) {
        const keep = (f as Record<string, { show?: boolean; overlay?: boolean; visible?: boolean } | undefined>)[key];
        const kept = keep ? { ...(keep.show !== undefined ? { show: keep.show } : {}), ...(keep.overlay ? { overlay: true } : {}), ...(keep.visible === false ? { visible: false } : {}) } : {};
        if (Object.keys(kept).length) (f as Record<string, unknown>)[key] = kept; else delete (f as Record<string, unknown>)[key];
      }
      if (!Object.keys(f).length) delete x.format;
    });
  }

  run("wchart.addElement", () => {});
  run("wchart.quickLayout", () => {}, () => { const m = selected(); return !!m && layoutCount(m) > 0; });
  run("wchart.colors", () => {});
  run("wchart.styles", () => {});
  // Word greys Zeile/Spalte tauschen out while the chart's data sheet is closed.
  run("wchart.switchRowCol", () => {}, () => false);
  run("wchart.selectData", () => deps.editData());
  run("wchart.editData", () => deps.editData());
  run("wchart.refreshData", () => {}, () => false);
  run("wchart.changeType", () => chartTypeDialog({
    insert: () => {},
    selected: () => selected(),
    change: (fn) => { const m = selected(); if (m) edit(m, fn); },
    previewSeries: () => selected()?.series ?? [],
  }, "change"));
  run("wchart.element", () => {});
  run("wchart.formatSelection", () => { const m = selected(); if (m) showPane(m); });
  run("wchart.formatObject", () => { const m = selected(); if (m) { element = { kind: "chart" }; showPane(m); } });
  run("wchart.formatText", () => { const m = selected(); if (m) showPane(m); });
  run("wchart.resetStyle", resetElement);
  run("wchart.fill", () => {});
  run("wchart.outline", () => {});
  run("wchart.textFill", () => {});
  run("wchart.height", () => {});
  run("wchart.width", () => {});
  run("wchart.sizeDialog", () => c.run("picture.moreLayout"));

  const fmtCm = (pt: number) => `${(pt / CM).toLocaleString("de-DE", { minimumFractionDigits: 2, maximumFractionDigits: 2 })} cm`;
  const parseCm = (v: string) => Number(v.replace(/[^\d,.-]/g, "").replace(",", "."));
  const setSize = (side: "w" | "h", cm: number) => {
    const s = sizePt();
    if (!(cm > 0)) return;
    ed()?.setImageSize(side === "w" ? cm * CM : s.w, side === "h" ? cm * CM : s.h);
    api.refresh();
  };
  const size = (id: string, label: string, key: string, side: "w" | "h"): Control => ({
    type: "spinner", cmd: id, label, key, icon: side === "h" ? "shape-height" : "shape-width",
    value: () => { if (!sel()) return ""; const s = sizePt(); return fmtCm(side === "w" ? s.w : s.h); },
    commit: (v) => setSize(side, parseCm(v)),
    step: (d) => { const s = sizePt(); const cur = (side === "w" ? s.w : s.h) / CM; setSize(side, Math.max(0.1, Math.round((cur + d * 0.1) * 100) / 100)); },
  });
  const noFill = (what: "fill" | "line"): MenuEntry[] => [{ label: what === "fill" ? "Keine Füllung" : "Keine Kontur", run: () => setElementColor(what, null) }];
  const colorMenu = (k: number, label: string): MenuEntry => {
    const m = selected();
    return { label, checked: (m?.colors ?? 10) === k, hover: () => preview(colorLook(k)), run: () => { preview(null); const x = selected(); if (x) edit(x, (y) => applyColors(y, k)); } };
  };

  const designTab: Tab = {
    id: "chartDesign", label: "Diagrammentwurf", key: "JC", contextual: true, shrink: [],
    groups: [
      { id: "layouts", label: "Diagrammlayouts", key: "ZL", icon: "chart-recommended", controls: [
        { type: "dropdown", cmd: "wchart.addElement", key: "D", size: "large", get menu() { return elementsMenu(); } },
        { type: "dropdown", cmd: "wchart.quickLayout", key: "C", size: "large", get menu() {
          const m = selected();
          return m ? Array.from({ length: layoutCount(m) }, (_, i): MenuEntry => ({ label: `Layout ${i + 1}`, hover: () => preview(layoutLook(i + 1)),
            run: () => { preview(null); const x = selected(); if (x) edit(x, (y) => { applyQuickLayout(y, i + 1); }); } })) : [];
        } }] },
      { id: "styles", label: "Diagrammformatvorlagen", key: "ZS", icon: "styles", controls: [
        { type: "dropdown", cmd: "wchart.colors", key: "F", size: "large", get menu() {
          return [{ heading: "Bunt" }, ...COLOR_SETS.slice(0, 4).map(([k, l]) => colorMenu(k, l)), { heading: "Einfarbig" }, ...COLOR_SETS.slice(4).map(([k, l]) => colorMenu(k, l))];
        } },
        { type: "tiles", cmd: "wchart.styles", key: "S", get count() { const m = selected(); return m ? chartStyles(m).length : 0; }, visible: 6, fill: true, size: [62, 46],
          render: (i) => { const look = styleLook(i); return look ? (tile(look, 58, 42) as unknown as HTMLElement) : null; },
          tip: (i) => `Formatvorlage ${i + 1}`,
          pick: (i) => { preview(null); const m = selected(); if (m) { setStyleTheme(theme()); edit(m, (x) => applyStyle(x, i)); } },
          hover: (i) => preview(i === null ? null : styleLook(i)),
          current: () => { const m = selected(); if (!m) return undefined; const i = styleIndex(m); return i >= 0 ? i : undefined; },
          redraw: () => { const m = selected(); return m ? `${m.kind},${m.series.length},${m.colors ?? 10}` : ""; },
          panel: (close) => {
            const grid = h("div", { class: "chart-style-panel", style: "display:grid;grid-template-columns:repeat(4,92px);gap:4px;padding:6px" });
            const m = selected();
            for (let i = 0; m && i < chartStyles(m).length; i++) {
              const t = h("button", { class: "tile", "aria-label": `Formatvorlage ${i + 1}`, "data-tip": `Formatvorlage ${i + 1}`, style: "padding:2px;border:1px solid #d6d6d6;background:#fff" });
              const look = styleLook(i);
              if (look) t.append(tile(look, 86, 62));
              t.addEventListener("mouseenter", () => preview(styleLook(i)));
              t.addEventListener("mouseleave", () => preview(null));
              t.addEventListener("click", () => { preview(null); close(); const x = selected(); if (x) { setStyleTheme(theme()); edit(x, (y) => applyStyle(y, i)); } });
              grid.append(t);
            }
            return grid;
          } } as Control] },
      { id: "data", label: "Daten", key: "ZD", icon: "table", controls: [
        { type: "button", cmd: "wchart.switchRowCol", key: "E", size: "large" },
        { type: "button", cmd: "wchart.selectData", key: "T", size: "large" },
        { type: "split", cmd: "wchart.editData", key: "A", size: "large", menu: [{ cmd: "wchart.editData" }] },
        { type: "button", cmd: "wchart.refreshData", key: "N", size: "large" }] },
      { id: "type", label: "Typ", key: "ZT", icon: "chart-column", controls: [{ type: "button", cmd: "wchart.changeType", key: "I", size: "large" }] },
    ],
  };

  // Anordnen with the keytips recorded in Word's chart Format tab (Position P P, Textumbruch J, Ebene nach vorne G E,
  // Ebene nach hinten G H, Auswahlbereich Q, Ausrichten D A, Gruppieren G R, Drehen D R).
  const layers: Control[] = [
    { type: "split", cmd: "picture.bringForward", key: "GE", size: "small", menu: [{ cmd: "picture.bringForward" }, { cmd: "picture.bringToFront" }, { cmd: "picture.bringInFrontOfText" }] },
    { type: "split", cmd: "picture.sendBackward", key: "GH", size: "small", menu: [{ cmd: "picture.sendBackward" }, { cmd: "picture.sendToBack" }, { cmd: "picture.sendBehindText" }] },
    { type: "button", cmd: "layout.selectionPane", key: "Q", size: "small" }];
  const formatTab: Tab = {
    id: "chartFormat", label: "Format", key: "JY", contextual: true, shrink: [],
    groups: [
      { id: "current", label: "Aktuelle Auswahl", key: "ZL", icon: "format-picture", controls: [{ type: "stack", items: [
        { type: "combo", cmd: "wchart.element", width: 160, key: "A1", value: () => { const m = selected(); return m ? elementName(m, element) : ""; },
          options: () => { const m = selected(); return [{ items: m ? chartElements(m).map((x) => elementName(m, x)) : [] }]; },
          commit: (v) => { const m = selected(); const x = m ? chartElements(m).find((e) => elementName(m, e) === v) : undefined; if (x) { element = x; if (paneOpen) showPane(m!); api.refresh(); } } },
        { type: "button", cmd: "wchart.formatSelection", key: "U", size: "small" },
        { type: "button", cmd: "wchart.resetStyle", key: "V1", size: "small" },
      ] }] },
      { id: "insertShapes", label: "Formen einfügen", key: "ZH", icon: "shapes", controls: [
        { type: "button", cmd: "wchart.insertShape", key: "2", size: "large" },
        { type: "button", cmd: "wchart.changeShape", key: "Ä", size: "small" }] },
      { id: "shapeStyles", label: "Formenarten", key: "ZY", icon: "shape-fill", launcher: { cmd: "wchart.formatObject", key: "BB" }, controls: [
        { type: "button", cmd: "wchart.shapeStyles", key: "NC", size: "large" },
        { type: "stack", items: [
          { type: "color", cmd: "wchart.fill", key: "Ü", palette: "font", current: () => "#E97132", pick: (col) => setElementColor("fill", col), showLabel: true, extra: () => noFill("fill") },
          { type: "color", cmd: "wchart.outline", key: "V2", palette: "font", current: () => "#156082", pick: (col) => setElementColor("line", col), showLabel: true, extra: () => noFill("line") },
          { type: "button", cmd: "wchart.shapeEffects", key: "MO", size: "small" },
        ] }] },
      { id: "wordArt", label: "WordArt-Formate", key: "ZS", icon: "font-color", launcher: { cmd: "wchart.formatText", key: "O" }, controls: [
        { type: "button", cmd: "wchart.wordArtStyles", key: "C", size: "large" },
        { type: "stack", items: [
          { type: "color", cmd: "wchart.textFill", key: "4", palette: "font", current: () => "#000000", pick: (col) => setElementColor("text", col), showLabel: true },
          { type: "button", cmd: "wchart.textOutline", key: "TE", size: "small" },
          { type: "button", cmd: "wchart.textEffects", key: "IE", size: "small" },
        ] }] },
      { id: "access", label: "Barrierefreiheit", key: "Z1", icon: "alt-text", controls: [{ type: "button", cmd: "picture.altText", key: "AT", size: "large", label: "Alternativtext" }] },
      { id: "arrange", label: "Anordnen", key: "ZA", icon: "wrap-text", controls: [
        { type: "dropdown", cmd: "picture.position", key: "PP", size: "large", menu: [], popup: deps.positionPanel },
        { type: "dropdown", cmd: "layout.wrapText", key: "J", size: "large", menu: deps.wrapMenu },
        { type: "stack", items: layers },
        { type: "stack", items: [
          { type: "dropdown", cmd: "picture.align", key: "DA", size: "small", menu: [], popup: deps.alignPanel },
          { type: "dropdown", cmd: "layout.group", key: "GR", size: "small", menu: GROUP_MENU },
          { type: "dropdown", cmd: "picture.rotate", key: "DR", size: "small", menu: deps.rotateMenu }] }] },
      { id: "size", label: "Größe", key: "ZZ", icon: "shape-height", launcher: { cmd: "wchart.sizeDialog", key: "RG" }, controls: [{ type: "stack", items: [
        size("wchart.height", "Höhe:", "H", "h"), size("wchart.width", "Breite:", "RR", "w")] }] },
    ],
  };

  return {
    designTab, formatTab,
    /** A chart is selected (the tabs show). */
    selected: () => !!sel(),
    /** After every change: the buttons follow the chart, a sized chart is drawn anew. */
    sync,
    /** For tests and the harness. */
    model: () => selected(),
    element: () => element,
    setElement: (e: ChartElement) => { element = e; },
    showButton: (which: "elements" | "styles" | "filter") => {
      const b = overlay.querySelector<HTMLElement>(`.chart-btn[data-which="${which}"]`);
      b?.click();
    },
    paneOpen: () => paneOpen && pane.isOpen,
    edit: (fn: (m: ChartModel) => void) => { const m = selected(); if (m) edit(m, fn); },
    kind: (): ChartKind | null => selected()?.kind ?? null,
  };
}
