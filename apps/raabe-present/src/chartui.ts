// The chart UI shared with RaabeCalc and RaabeWrite (@raabe/chart-ui, D-129) on slides (D-132): the three buttons
// beside a selected chart as PowerPoint shows them (Diagrammelemente, Diagrammformatvorlagen, Diagrammfilter) and the
// Format pane of the chart's elements, which "Objekt formatieren..." (and Diagrammbereich formatieren...) opens for a
// chart instead of the shape pane. The chart part is read into the chart engine's model, changed in place and written
// back as one undo step (the slide and the chart part, as chartops.ts does).

import { applyStyle, chartButtons, chartElements, chartStyles, elementName, formatPane, styleIndex, type ChartElement } from "@raabe/chart-ui";
import { renderChart, seriesName, type ChartModel } from "@raabe/charts";
import { chartFromXml, chartToXml } from "@raabe/ooxml";
import { TaskPane, h, showDialog, type Control } from "@raabe/shell";
import type { PresentApp } from "./app";

const EMU_CM = 360000;

const CSS = `
.pchart-overlay { position: fixed; pointer-events: none; z-index: 14; }
.pchart-overlay .chart-btn { pointer-events: auto; }
.pchart-overlay .chart-btn.on { background: #fbe3d6; border-color: #c43e1c; }
`;

export function presentChartUi(app: PresentApp) {
  if (!document.getElementById("raabe-pchart-css")) document.head.append(h("style", { id: "raabe-pchart-css" }, CSS));
  const view = () => app.editor.selectedViews.find((v) => v.frame === "chart" && v.chart);
  const models = new Map<string, ChartModel>();
  /** The selected chart's model, the same object while its part does not change elsewhere (undo reads it anew). */
  function selected(): ChartModel | null {
    const v = view();
    const xml = v?.chart ? app.deck.pptx.text(v.chart) : null;
    if (!v || !xml) return null;
    const key = `${v.chart}\u0001${xml}`;
    let m = models.get(key);
    if (!m) { models.clear(); m = chartFromXml(xml); models.set(key, m); }
    return m;
  }
  let element: ChartElement = { kind: "chart" };

  function edit(m: ChartModel, fn: (m: ChartModel) => void, label = "Diagrammelemente"): void {
    const v = view();
    if (!v?.chart) return;
    const path = v.chart;
    const cur = selected() ?? m;
    fn(cur);
    const xml = chartToXml(cur);
    const slide = app.slidePath();
    app.deck.change(label, slide ? [slide, path] : [path], () => app.deck.pptx.put(path, xml));
    app.deck.markChanged([path]);
    models.clear();
    models.set(`${path}\u0001${xml}`, cur);
    app.editor.render();
    app.selectionChanged();
    if (buttons.isOpen()) buttons.refresh(cur);
    if (pane.isOpen) showPane(cur);
  }

  const accents = (): string[] => {
    const t = (app as unknown as { paletteTheme(): { colors?: Record<string, string> } | undefined }).paletteTheme?.()?.colors ?? {};
    return [1, 2, 3, 4, 5, 6].map((i, k) => "#" + String(t[`accent${i}`] ?? ["4472C4", "ED7D31", "A5A5A5", "FFC000", "5B9BD5", "70AD47"][k]).replace("#", ""));
  };

  const buttons = chartButtons({
    edit: (m, fn, label) => edit(m, fn, label),
    // The slide draws the chart from its part; the flyouts' Livevorschau and highlighting are not shown on slides yet.
    preview: () => {},
    highlight: () => {},
    openPane: (m, e) => { element = e; showPane(m); },
    selectData: () => void app.commands.run("chart.editData"),
    accents,
    tile: (m, w, hh) => {
      const svg = renderChart(m, 480, 288, { accents: accents() }) as SVGSVGElement;
      const k = Math.min(w / 480, hh / 288);
      svg.setAttribute("width", String(Math.round(480 * k)));
      svg.setAttribute("height", String(Math.round(288 * k)));
      return svg;
    },
    pickSeries: async (m) => {
      const list = h("select", { size: 6, style: "width:260px", "aria-label": "Reihen" }) as HTMLSelectElement;
      m.series.forEach((s, i) => { if (!s.filtered) list.append(h("option", { value: String(i) }, seriesName(s, i))); });
      list.selectedIndex = 0;
      const ok = await showDialog({ title: "Trendlinie hinzufügen", body: h("div", null, h("div", { style: "margin-bottom:6px" }, "Trendlinie hinzufügen basierend auf Reihe:"), list),
        buttons: [{ label: "OK", value: true, primary: true }, { label: "Abbrechen", value: false }], cancel: false });
      return ok && list.value !== "" ? Number(list.value) : null;
    },
    nameSources: (m) => ({
      series: m.series.some((s) => s.name) ? "Zeile 1" : null, cats: m.series[0]?.cat ? "Spalte A" : null,
      seriesList: m.series.map((s, i) => seriesName(s, i)).join("; "),
      catList: (m.series[0]?.cat?.cache ?? []).map((v) => (v === null ? "" : String(v))).join("; "),
    }),
    setNameSources: () => {},
  });

  const overlay = h("div", { class: "pchart-overlay chart-obj sel", hidden: true }) as HTMLDivElement;
  document.body.append(overlay);
  buttons.attach(overlay, () => selected()!);

  /** The selected chart's frame on the screen (the slide element with the shape's id). */
  function rect(): DOMRect | null {
    const v = view();
    if (!v) return null;
    const el = document.querySelector<HTMLElement>(`.stage [data-id="${v.id}"]`);
    const r = el?.getBoundingClientRect();
    return r && r.width > 0 ? r : null;
  }
  let following = false;
  function sync(): void {
    const r = rect();
    if (!r || app.editor.text) {
      overlay.hidden = true;
      if (buttons.isOpen()) buttons.close();
      return;
    }
    overlay.hidden = false;
    Object.assign(overlay.style, { left: `${r.left}px`, top: `${r.top}px`, width: `${r.width}px`, height: `${r.height}px` });
    if (!following) {
      following = true;
      const follow = () => {
        if (overlay.hidden) { following = false; return; }
        const q = rect();
        const o = overlay.getBoundingClientRect();
        if (!q) { sync(); following = false; return; }
        if (Math.abs(q.left - o.left) > 0.5 || Math.abs(q.top - o.top) > 0.5 || Math.abs(q.width - o.width) > 0.5) sync();
        requestAnimationFrame(follow);
      };
      requestAnimationFrame(follow);
    }
  }

  const pane = new TaskPane();
  function showPane(m: ChartModel): void {
    if (!chartElements(m).some((x) => JSON.stringify(x) === JSON.stringify(element))) element = { kind: "chart" };
    pane.show(formatPane(m, element, {
      edit: (fn) => edit(m, fn, "Formatieren"),
      select: (x) => { element = x; showPane(selected() ?? m); },
      size: () => {
        const v = view();
        const cx = v?.w ?? 0, cy = v?.h ?? 0;
        return { w: Math.round((cx / EMU_CM) * 100) / 100, h: Math.round((cy / EMU_CM) * 100) / 100 };
      },
      resize: () => {},
    }));
  }

  /** WordArt-Formate on a selected chart: the colour of the chart's text (title, legend, axes, labels; PowerPoint). */
  function setTextColor(hex: string): boolean {
    const m = selected();
    if (!m) return false;
    edit(m, (c) => {
      const f = (c.format ??= {});
      for (const k of ["title", "legend", "valAxis", "catAxis", "dataLabels"] as const) {
        const el = ((f as Record<string, { text?: { color?: string } }>)[k] ??= {});
        el.text = { ...(el.text ?? {}), color: hex.startsWith("#") ? hex : "#" + hex };
      }
    }, "Textfüllung");
    return true;
  }

  /** A drawing of `m` in a w x hh box (the style tiles). */
  const tileOf = (m: ChartModel, w: number, hh: number): SVGSVGElement => {
    const svg = renderChart(m, 480, 288, { accents: accents() }) as SVGSVGElement;
    const k = Math.min(w / 480, hh / 288);
    svg.setAttribute("width", String(Math.round(480 * k)));
    svg.setAttribute("height", String(Math.round(288 * k)));
    return svg;
  };
  const styleLook = (i: number): ChartModel | null => { const m = selected(); if (!m) return null; const look = structuredClone(m); applyStyle(look, i); return look; };

  /**
   * Diagrammentwurf > Diagrammformatvorlagen: the styles as a gallery in the ribbon, as PowerPoint shows them (D-187):
   * eight 88 px tiles at a wide window, at least four. No Livevorschau on slides yet.
   */
  function styleTiles(): Control {
    return {
      type: "tiles", cmd: "chart.styles", key: "S", get count() { const m = selected(); return m ? chartStyles(m).length : 0; }, visible: 4, fill: true, max: 8, size: [79, 48],
      render: (i) => { const look = styleLook(i); return look ? (tileOf(look, 76, 44) as unknown as HTMLElement) : null; },
      tip: (i) => `Formatvorlage ${i + 1}`,
      pick: (i) => { const m = selected(); if (m) edit(m, (x) => applyStyle(x, i), "Diagrammformatvorlage"); },
      current: () => { const m = selected(); if (!m) return undefined; const i = styleIndex(m); return i >= 0 ? i : undefined; },
      redraw: () => { const m = selected(); return m ? `${m.kind},${m.series.length},${m.colors ?? 10}` : ""; },
      panel: (close) => {
        const grid = h("div", { class: "chart-style-panel", style: "display:grid;grid-template-columns:repeat(4,92px);gap:4px;padding:6px" });
        const m = selected();
        for (let i = 0; m && i < chartStyles(m).length; i++) {
          const t = h("button", { class: "tile", "aria-label": `Formatvorlage ${i + 1}`, "data-tip": `Formatvorlage ${i + 1}`, style: "padding:2px;border:1px solid #d6d6d6;background:#fff" });
          const look = styleLook(i);
          if (look) t.append(tileOf(look, 86, 62));
          t.addEventListener("click", () => { close(); const x = selected(); if (x) edit(x, (y) => applyStyle(y, i), "Diagrammformatvorlage"); });
          grid.append(t);
        }
        return grid;
      },
    } as Control;
  }

  return {
    sync,
    styleTiles,
    /** Objekt formatieren... with a chart selected: the chart's Format pane (Diagrammbereich, or the chosen element). */
    openPane: (name?: string): boolean => {
      const m = selected();
      if (!m) return false;
      // Aktuelle Auswahl's entries by name (PowerPoint's names are the pane's).
      element = (name ? chartElements(m).find((e) => elementName(m, e) === name) : undefined) ?? { kind: "chart" };
      showPane(m);
      return true;
    },
    element: () => element,
    setTextColor,
    selected: () => !!selected(),
    model: () => selected(),
    paneOpen: () => pane.isOpen,
    paneTitle: () => pane.el.querySelector(".tp-title")?.textContent ?? "",
  };
}
