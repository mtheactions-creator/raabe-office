// Status bar: page and word counts on the left, view buttons and the zoom slider on the
// right, laid out like Word's. The zoom slider maps 10-500 % with 100 % in the middle.

import { h, icon } from "./dom";
import type { Commands } from "./commands";

export interface StatusInfo {
  page: number;
  pages: number;
  words: number | null;
  zoom: number;
  /** Replaces "Seite 1 von 3" (Lesemodus: "Bildschirm 1 von 6"). */
  pageLabel?: string;
}

function sliderToZoom(v: number): number {
  return v <= 50 ? 0.1 + (v / 50) * 0.9 : 1 + ((v - 50) / 50) * 4;
}

function zoomToSlider(z: number): number {
  return z <= 1 ? ((z - 0.1) / 0.9) * 50 : 50 + ((z - 1) / 4) * 50;
}

export class StatusBar {
  readonly el: HTMLElement;
  private page = h("button", { class: "cell", "data-tip": "Seitenzahl im Dokument" });
  private words = h("span", { class: "cell", "data-tip": "Wortanzahl" });
  private slider = h("input", { type: "range", min: "0", max: "100", step: "0.5", "aria-label": "Zoom" });
  private pct = h("button", { class: "cell pct", "data-tip": "Zoomfaktor" });

  constructor(commands: Commands, private onZoom: (z: number) => void, views: { cmd: string }[], zoomDialogCmd: string) {
    const viewButtons = views.map((v) => {
      const c = commands.get(v.cmd);
      const b = h("button", { class: "cell viewbtn", "aria-label": c.label, "data-tip": commands.tooltip(v.cmd), onclick: () => void commands.run(v.cmd) }, icon(c.icon));
      commands.onRefresh(() => {
        b.disabled = !commands.isEnabled(v.cmd);
        b.classList.toggle("checked", commands.isChecked(v.cmd));
      });
      return b;
    });
    this.slider.oninput = () => this.onZoom(Math.round(sliderToZoom(Number(this.slider.value)) * 100) / 100);
    this.pct.onclick = () => void commands.run(zoomDialogCmd);
    const step = (d: number) => () => {
      const z = sliderToZoom(Number(this.slider.value));
      this.onZoom(Math.max(0.1, Math.min(5, Math.round((z + d) * 10) / 10)));
    };
    this.el = h("div", { class: "statusbar", role: "status" },
      this.page, this.words,
      h("span", { class: "cell", "data-tip": "Sprache: Deutsch (Deutschland)" }, "Deutsch (Deutschland)"),
      h("div", { class: "right" }, ...viewButtons,
        h("div", { class: "zoom" },
          h("button", { "aria-label": "Verkleinern", "data-tip": "Verkleinern (Strg -)", onclick: step(-0.1) }, "−"),
          this.slider,
          h("button", { "aria-label": "Vergrößern", "data-tip": "Vergrößern (Strg +)", onclick: step(0.1) }, "+")),
        this.pct));
  }

  update(s: StatusInfo): void {
    this.page.textContent = s.pageLabel ?? `Seite ${s.page} von ${s.pages}`;
    this.words.textContent = s.words === null ? "" : `${s.words.toLocaleString("de-DE")} ${s.words === 1 ? "Wort" : "Wörter"}`;
    if (document.activeElement !== this.slider) this.slider.value = String(zoomToSlider(s.zoom));
    this.pct.textContent = `${Math.round(s.zoom * 100)} %`;
  }
}
