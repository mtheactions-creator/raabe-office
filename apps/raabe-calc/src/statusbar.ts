// Excel's status bar: the cell mode on the left (Bereit, Eingeben, Bearbeiten, Zeigen), the
// statistics of the selection (Mittelwert, Anzahl, Summe when at least two cells hold values),
// the view buttons and the zoom slider on the right.

import { h, icon, showMenu, type Commands } from "@raabe/shell";

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
  private flags = h("span", { class: "flags" });
  private message = h("span", { class: "cell flag message", "aria-label": "Statusmeldung", hidden: true });
  private stats = h("div", { class: "stats" });
  private slider = h("input", { type: "range", min: "0", max: "100", step: "0.5", "aria-label": "Zoom" }) as HTMLInputElement;
  private pct = h("button", { class: "cell pct", "data-tip": "Zoom" });

  /** What the statistics show; Excel's default is Mittelwert, Anzahl and Summe. */
  private shown: Record<string, boolean> = { avg: true, count: true, numbers: false, min: false, max: false, sum: true };
  private last: { zoom: number; s: Stats | null } = { zoom: 100, s: null };

  constructor(private commands: Commands, onZoom: (z: number) => void, private format: (v: number) => string) {
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
    this.el = h("div", { class: "statusbar", role: "status", oncontextmenu: (e: MouseEvent) => this.menu(e) }, this.mode, this.message, this.flags,
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

  /** A message in place of the cell mode (Excel's copy mode); null shows the mode again. */
  setMessage(text: string | null): void {
    this.message.textContent = text ?? "";
    this.message.hidden = !text;
    this.mode.hidden = !!text;
  }

  /** Mode texts beside the cell mode (Endmodus, Auswahl erweitern …), named as Excel names them. */
  setFlags(flags: [string, string][]): void {
    this.modeFlags = flags.map(([field, text]) => [`${field} ${text}`, text]);
    this.renderFlags();
  }

  /** A note of the workbook beside them (Berechnen, Zirkelbezüge: A1), as [accessible name, text]; null removes it. */
  setNote(key: string, note: [string, string] | null): void {
    if (note) this.notes.set(key, note);
    else this.notes.delete(key);
    this.renderFlags();
  }

  private modeFlags: [string, string][] = [];
  private notes = new Map<string, [string, string]>();

  private renderFlags(): void {
    this.flags.replaceChildren(...[...this.modeFlags, ...this.notes.values()].map(([name, text]) => h("span", { class: "cell flag", "aria-label": name }, text)));
  }

  /** Statusleiste anpassen: which statistics are shown (right click, as in Excel). */
  private menu(e: MouseEvent): void {
    e.preventDefault();
    const entry = (key: string, label: string) => ({ label, checked: this.shown[key], run: () => { this.shown[key] = !this.shown[key]; this.update(this.last.zoom, this.last.s); } });
    showMenu(this.commands, [{ heading: "Statusleiste anpassen" }, entry("avg", "Mittelwert"), entry("count", "Anzahl"), entry("numbers", "Numerische Zählung"),
      entry("min", "Minimum"), entry("max", "Maximum"), entry("sum", "Summe")], new DOMRect(e.clientX, e.clientY, 0, 0));
  }

  update(zoom: number, s: Stats | null): void {
    this.last = { zoom, s };
    if (document.activeElement !== this.slider) this.slider.value = String(zoomToSlider(zoom));
    this.pct.textContent = `${zoom} %`;
    const parts: string[] = [];
    const on = this.shown;
    if (s && s.count >= 2) {
      if (on.avg && s.numbers > 0) parts.push(`Mittelwert: ${this.format(s.sum / s.numbers)}`);
      if (on.count) parts.push(`Anzahl: ${s.count}`);
      if (on.numbers && s.numbers > 0) parts.push(`Numerische Zählung: ${s.numbers}`);
      if (on.min && s.numbers > 0) parts.push(`Min: ${this.format(s.min)}`);
      if (on.max && s.numbers > 0) parts.push(`Max: ${this.format(s.max)}`);
      if (on.sum && s.numbers > 0) parts.push(`Summe: ${this.format(s.sum)}`);
    }
    this.stats.replaceChildren(...parts.map((p) => h("span", { class: "cell stat" }, p)));
  }
}
