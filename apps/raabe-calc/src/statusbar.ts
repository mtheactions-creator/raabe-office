// Excel's status bar: the cell mode on the left (Bereit, Eingeben, Bearbeiten, Zeigen), the
// statistics of the selection (Mittelwert, Anzahl, Summe when at least two cells hold values),
// the view buttons and the zoom slider on the right.

import { h, icon, type Commands } from "@raabe/shell";

function sliderToZoom(v: number): number {
  return v <= 50 ? 10 + (v / 50) * 90 : 100 + ((v - 50) / 50) * 300;
}

function zoomToSlider(z: number): number {
  return z <= 100 ? ((z - 10) / 90) * 50 : 50 + ((z - 100) / 300) * 50;
}

export interface Stats {
  count: number;
  numbers: number;
  sum: number;
  min: number;
  max: number;
}

export class CalcStatusBar {
  readonly el: HTMLElement;
  private mode = h("span", { class: "cell mode" }, "Bereit");
  private stats = h("div", { class: "stats" });
  private slider = h("input", { type: "range", min: "0", max: "100", step: "0.5", "aria-label": "Zoom" }) as HTMLInputElement;
  private pct = h("button", { class: "cell pct", "data-tip": "Zoom" });

  constructor(commands: Commands, onZoom: (z: number) => void, private format: (v: number) => string) {
    const views = ["view.normal", "view.pageLayout", "view.pageBreak"].map((id) => {
      const c = commands.get(id);
      const b = h("button", { class: "cell viewbtn", "aria-label": c.label, "data-tip": commands.tooltip(id), onclick: () => void commands.run(id) }, icon(c.icon));
      commands.onRefresh(() => { b.disabled = !commands.isEnabled(id); b.classList.toggle("checked", commands.isChecked(id)); });
      return b;
    });
    this.slider.oninput = () => onZoom(Math.round(sliderToZoom(Number(this.slider.value))));
    this.pct.onclick = () => void commands.run("view.zoomDialog");
    const step = (d: number) => () => {
      const z = sliderToZoom(Number(this.slider.value));
      onZoom(Math.max(10, Math.min(400, Math.round((z + d) / 10) * 10)));
    };
    this.el = h("div", { class: "statusbar", role: "status" }, this.mode,
      h("div", { class: "right" }, this.stats, ...views,
        h("div", { class: "zoom" },
          h("button", { "aria-label": "Verkleinern", "data-tip": "Verkleinern", onclick: step(-10) }, "−"),
          this.slider,
          h("button", { "aria-label": "Vergrößern", "data-tip": "Vergrößern", onclick: step(10) }, "+")),
        this.pct));
  }

  setMode(text: string): void {
    this.mode.textContent = text;
  }

  update(zoom: number, s: Stats | null): void {
    if (document.activeElement !== this.slider) this.slider.value = String(zoomToSlider(zoom));
    this.pct.textContent = `${zoom} %`;
    const parts: string[] = [];
    if (s && s.count >= 2) {
      if (s.numbers > 0) parts.push(`Mittelwert: ${this.format(s.sum / s.numbers)}`);
      parts.push(`Anzahl: ${s.count}`);
      if (s.numbers > 0) parts.push(`Summe: ${this.format(s.sum)}`);
    }
    this.stats.replaceChildren(...parts.map((p) => h("span", { class: "cell stat" }, p)));
  }
}
