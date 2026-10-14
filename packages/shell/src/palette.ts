// Word's colour pickers: "Schriftfarbe" (theme colours with tints and shades, standard
// colours, Automatisch) and "Texthervorhebungsfarbe" (the 15 highlight colours).

import { h } from "./dom";
import { closePopup, showPopup } from "./popup";

/** Office 2013–2022 theme colours, in Word's palette order. */
const THEME = [
  ["FFFFFF", "Weiß, Hintergrund 1"], ["000000", "Schwarz, Text 1"], ["E7E6E6", "Hellgrau, Hintergrund 2"], ["44546A", "Blaugrau, Text 2"],
  ["4472C4", "Blau, Akzent 1"], ["ED7D31", "Orange, Akzent 2"], ["A5A5A5", "Grau, Akzent 3"], ["FFC000", "Gold, Akzent 4"],
  ["5B9BD5", "Blau, Akzent 5"], ["70AD47", "Grün, Akzent 6"],
] as const;

const STANDARD = [
  ["C00000", "Dunkelrot"], ["FF0000", "Rot"], ["FFC000", "Orange"], ["FFFF00", "Gelb"], ["92D050", "Hellgrün"],
  ["00B050", "Grün"], ["00B0F0", "Hellblau"], ["0070C0", "Blau"], ["002060", "Dunkelblau"], ["7030A0", "Lila"],
] as const;

export const HIGHLIGHTS = [
  ["yellow", "FFFF00", "Gelb"], ["green", "00FF00", "Hellgrün"], ["cyan", "00FFFF", "Türkis"], ["magenta", "FF00FF", "Rosa"],
  ["blue", "0000FF", "Blau"], ["red", "FF0000", "Rot"], ["darkBlue", "000080", "Dunkelblau"], ["darkCyan", "008080", "Blaugrün"],
  ["darkGreen", "008000", "Grün"], ["darkMagenta", "800080", "Violett"], ["darkRed", "800000", "Dunkelrot"],
  ["darkYellow", "808000", "Dunkelgelb"], ["darkGray", "808080", "Grau 50 %"], ["lightGray", "C0C0C0", "Grau 25 %"], ["black", "000000", "Schwarz"],
] as const;

// HSL luminance changes, as Word computes tints and shades.
function toHsl(hex: string): [number, number, number] {
  const n = parseInt(hex, 16);
  const r = (n >> 16) / 255, g = ((n >> 8) & 255) / 255, b = (n & 255) / 255;
  const max = Math.max(r, g, b), min = Math.min(r, g, b);
  const l = (max + min) / 2;
  if (max === min) return [0, 0, l];
  const d = max - min;
  const s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
  let hue = max === r ? (g - b) / d + (g < b ? 6 : 0) : max === g ? (b - r) / d + 2 : (r - g) / d + 4;
  hue /= 6;
  return [hue, s, l];
}

function fromHsl(hu: number, s: number, l: number): string {
  const f = (p: number, q: number, t: number) => {
    if (t < 0) t += 1;
    if (t > 1) t -= 1;
    if (t < 1 / 6) return p + (q - p) * 6 * t;
    if (t < 1 / 2) return q;
    if (t < 2 / 3) return p + (q - p) * (2 / 3 - t) * 6;
    return p;
  };
  let r = l, g = l, b = l;
  if (s) {
    const q = l < 0.5 ? l * (1 + s) : l + s - l * s, p = 2 * l - q;
    r = f(p, q, hu + 1 / 3); g = f(p, q, hu); b = f(p, q, hu - 1 / 3);
  }
  return [r, g, b].map((x) => Math.round(x * 255).toString(16).padStart(2, "0")).join("").toUpperCase();
}

function variants(hex: string): { hex: string; label: string }[] {
  const [hu, s, l] = toHsl(hex);
  const lighter = (p: number) => ({ hex: fromHsl(hu, s, l + (1 - l) * p), label: `heller ${Math.round(p * 100)} %` });
  const darker = (p: number) => ({ hex: fromHsl(hu, s, l * (1 - p)), label: `dunkler ${Math.round(p * 100)} %` });
  if (l > 0.9) return [darker(0.05), darker(0.15), darker(0.25), darker(0.35), darker(0.5)];
  if (l < 0.1) return [lighter(0.5), lighter(0.35), lighter(0.25), lighter(0.15), lighter(0.05)];
  return [lighter(0.8), lighter(0.6), lighter(0.4), darker(0.25), darker(0.5)];
}

function swatch(hex: string, label: string, pick: () => void): HTMLElement {
  return h("button", {
    class: "swatch", style: `background:#${hex}`, "aria-label": label, "data-tip": label,
    onclick: () => { closePopup(); pick(); },
  });
}

/** Word's colour grid (Designfarben with tints and shades, Standardfarben), for other menus. */
export function colorGrid(pick: (hex: string) => void): HTMLElement[] {
  const grid = h("div", { class: "swatches", style: "grid-template-columns:repeat(10, 18px)" });
  for (const [hex, label] of THEME) grid.append(swatch(hex, label, () => pick(hex)));
  const vars = THEME.map(([hex]) => variants(hex));
  for (let row = 0; row < 5; row++) {
    for (let col = 0; col < THEME.length; col++) {
      const v = vars[col][row];
      grid.append(swatch(v.hex, `${THEME[col][1]}, ${v.label}`, () => pick(v.hex)));
    }
  }
  const std = h("div", { class: "swatches", style: "grid-template-columns:repeat(10, 18px)" },
    ...STANDARD.map(([hex, label]) => swatch(hex, label, () => pick(hex))));
  return [h("div", { class: "palette-title" }, "Designfarben"), grid, h("div", { class: "palette-title" }, "Standardfarben"), std];
}

export function showFontColors(anchor: DOMRect, pick: (hex: string) => void): void {
  const grid = h("div", { class: "swatches", style: "grid-template-columns:repeat(10, 18px)" });
  for (const [hex, label] of THEME) grid.append(swatch(hex, label, () => pick(hex)));
  const vars = THEME.map(([hex]) => variants(hex));
  for (let row = 0; row < 5; row++) {
    for (let col = 0; col < THEME.length; col++) {
      const v = vars[col][row];
      grid.append(swatch(v.hex, `${THEME[col][1]}, ${v.label}`, () => pick(v.hex)));
    }
  }
  const std = h("div", { class: "swatches", style: "grid-template-columns:repeat(10, 18px)" },
    ...STANDARD.map(([hex, label]) => swatch(hex, label, () => pick(hex))));
  const custom = h("input", { type: "color", style: "position:absolute;opacity:0;width:0;height:0" });
  custom.addEventListener("change", () => { closePopup(); pick(custom.value.slice(1).toUpperCase()); });
  const box = h("div", { class: "menu palette" },
    h("button", { onclick: () => { closePopup(); pick("auto"); } }, h("span", { class: "auto-swatch" }), "Automatisch"),
    h("div", { class: "palette-title" }, "Designfarben"), grid,
    h("div", { class: "palette-title" }, "Standardfarben"), std,
    h("hr"),
    h("button", { onclick: () => custom.click() }, "Weitere Farben..."), custom);
  showPopup(box, anchor);
}

export function showHighlights(anchor: DOMRect, pick: (name: string) => void): void {
  const grid = h("div", { class: "swatches", style: "grid-template-columns:repeat(5, 22px)" },
    ...HIGHLIGHTS.map(([name, hex, label]) => swatch(hex, label, () => pick(name))));
  const box = h("div", { class: "menu palette" }, grid, h("hr"),
    h("button", { onclick: () => { closePopup(); pick("none"); } }, "Keine Farbe"));
  showPopup(box, anchor);
}

export function highlightHex(name: string): string {
  return HIGHLIGHTS.find((x) => x[0] === name)?.[1] ?? "FFFF00";
}
