// Office's "Farben" dialog behind every "Weitere Farben..." (D-134): the tabs Standard (a grid of colours) and
// Benutzerdefiniert (Rot, Grün, Blau and Hex with the new and the current colour side by side). It replaces the
// browser's own colour input, whose system dialog the app's web view did not open. Dialogs can also be moved by
// their title bar, as Office's can (installDialogDrag).

import { h } from "./dom";
import { showDialog } from "./dialog";

const hex2 = (n: number) => Math.max(0, Math.min(255, Math.round(n))).toString(16).padStart(2, "0").toUpperCase();

/** HSL (h 0..360, s and l 0..1) to RRGGBB. */
function hsl(hue: number, s: number, l: number): string {
  const c = (1 - Math.abs(2 * l - 1)) * s;
  const x = c * (1 - Math.abs(((hue / 60) % 2) - 1));
  const m = l - c / 2;
  const [r, g, b] = hue < 60 ? [c, x, 0] : hue < 120 ? [x, c, 0] : hue < 180 ? [0, c, x] : hue < 240 ? [0, x, c] : hue < 300 ? [x, 0, c] : [c, 0, x];
  return hex2((r + m) * 255) + hex2((g + m) * 255) + hex2((b + m) * 255);
}

/** The Standard tab's colours: twelve hues in six shades, then a row of greys. */
function standardColors(): string[] {
  const out: string[] = [];
  for (const l of [0.85, 0.7, 0.55, 0.45, 0.35, 0.22]) for (let i = 0; i < 12; i++) out.push(hsl(i * 30, l > 0.8 ? 0.7 : 0.85, l));
  for (let i = 0; i < 12; i++) out.push(hex2(255 - (i * 255) / 11).repeat(3));
  return out;
}

/** "Farben": resolves to RRGGBB, or null for Abbrechen. `current` is the colour shown as Aktuell. */
export async function colorDialog(current = "000000"): Promise<string | null> {
  let value = /^[0-9A-F]{6}$/i.test(current) ? current.toUpperCase() : "000000";
  const num = (label: string) => h("input", { type: "number", min: "0", max: "255", step: "1", style: "width:70px", "aria-label": label }) as HTMLInputElement;
  const r = num("Rot:"), g = num("Grün:"), b = num("Blau:");
  const hex = h("input", { type: "text", style: "width:70px", "aria-label": "Hex:", maxlength: "7" }) as HTMLInputElement;
  const fresh = h("div", { style: "height:28px;border:1px solid #888;border-bottom:none" });
  const cur = h("div", { style: `height:28px;border:1px solid #888;background:#${value}` });
  const sync = (from?: HTMLInputElement) => {
    if (from !== r && from !== g && from !== b) {
      r.value = String(parseInt(value.slice(0, 2), 16)); g.value = String(parseInt(value.slice(2, 4), 16)); b.value = String(parseInt(value.slice(4, 6), 16));
    }
    if (from !== hex) hex.value = `#${value}`;
    fresh.style.background = `#${value}`;
    grid.querySelectorAll<HTMLElement>("button").forEach((x) => x.classList.toggle("on", x.dataset.hex === value));
  };
  const fromRgb = () => { value = hex2(+r.value || 0) + hex2(+g.value || 0) + hex2(+b.value || 0); sync(r); };
  for (const x of [r, g, b]) x.addEventListener("input", fromRgb);
  hex.addEventListener("input", () => { const m = /^#?([0-9a-f]{6})$/i.exec(hex.value.trim()); if (m) { value = m[1].toUpperCase(); sync(hex); } });
  const grid = h("div", { class: "swatches", style: "grid-template-columns:repeat(12, 20px);gap:2px" }, ...standardColors().map((c) =>
    h("button", { class: "swatch", "data-hex": c, "aria-label": `#${c}`, style: `width:20px;height:20px;background:#${c};border:1px solid rgba(0,0,0,.15)`, onclick: () => { value = c; sync(); } })));
  const pageStd = h("div", { style: "padding:8px 0" }, h("div", { style: "margin-bottom:6px" }, "Farben:"), grid);
  const pageCustom = h("div", { style: "padding:8px 0;display:grid;grid-template-columns:auto auto;gap:6px 10px;align-items:center;justify-content:start" },
    "Farbmodell:", h("select", { "aria-label": "Farbmodell:" }, h("option", null, "RGB")), "Rot:", r, "Grün:", g, "Blau:", b, "Hex:", hex);
  pageCustom.hidden = true;
  const tabs = h("div", { class: "tabs", role: "tablist", style: "display:flex;gap:4px;border-bottom:1px solid #ddd" });
  const tab = (label: string, page: HTMLElement, other: HTMLElement) => {
    const t = h("button", { role: "tab", style: "padding:4px 10px", onclick: () => { page.hidden = false; other.hidden = true; tabs.querySelectorAll("button").forEach((x) => x.classList.toggle("active", x === t)); } }, label);
    return t;
  };
  const tStd = tab("Standard", pageStd, pageCustom);
  const tCustom = tab("Benutzerdefiniert", pageCustom, pageStd);
  tStd.classList.add("active");
  tabs.append(tStd, tCustom);
  const body = h("div", { style: "display:grid;grid-template-columns:1fr 70px;gap:16px;width:420px" },
    h("div", null, tabs, pageStd, pageCustom),
    h("div", { style: "align-self:end;text-align:center;font-size:12px" }, "Neu", fresh, cur, "Aktuell"));
  sync();
  const ok = await showDialog({ title: "Farben", body, buttons: [{ label: "OK", value: true, primary: true }, { label: "Abbrechen", value: false }], cancel: false });
  return ok ? value : null;
}

let dragInstalled = false;

/** Office's dialogs move with their title bar: any `.dialog .dlg-title` dragged moves its dialog. */
export function installDialogDrag(): void {
  if (dragInstalled) return;
  dragInstalled = true;
  addEventListener("mousedown", (e) => {
    const title = (e.target as Element | null)?.closest?.(".dialog .dlg-title");
    if (!title || e.button !== 0 || (e.target as Element).closest("button")) return;
    const dlg = title.closest<HTMLElement>(".dialog");
    if (!dlg) return;
    e.preventDefault();
    const r = dlg.getBoundingClientRect();
    const dx = e.clientX - r.left, dy = e.clientY - r.top;
    const move = (m: MouseEvent) => {
      // Fixed position from the first move on; the title stays on screen.
      dlg.style.position = "fixed";
      dlg.style.margin = "0";
      dlg.style.left = `${Math.max(-r.width + 80, Math.min(innerWidth - 80, m.clientX - dx))}px`;
      dlg.style.top = `${Math.max(0, Math.min(innerHeight - 36, m.clientY - dy))}px`;
    };
    const up = () => { removeEventListener("mousemove", move, true); removeEventListener("mouseup", up, true); };
    addEventListener("mousemove", move, true);
    addEventListener("mouseup", up, true);
  }, true);
}
