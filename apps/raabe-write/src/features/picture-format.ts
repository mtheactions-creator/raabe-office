// The Bildformat tab's galleries as Word 365 has them (recorded with
// tools/feel-lab/probe_word_picture.py): Korrekturen, Farbe, Transparenz, Bildformatvorlagen,
// Bildrahmen, Bildeffekte, Drehen. Each tile shows the selected picture with the setting, as
// Word's galleries do; hovering a tile previews it in the document (Livevorschau).

import { closePopup, h, type MenuEntry } from "@raabe/shell";
import { previewElement, shapeElement, type Editor, type PictureFormat } from "@raabe/text-engine";
import { imageUrl } from "@raabe/text-engine";

export interface GalleryItem {
  label: string;
  /** The settings the entry gives (merged into the picture's). */
  format: PictureFormat;
}

export interface GallerySection {
  title?: string;
  items: GalleryItem[];
  /** Tiles per row and tile size (CSS px). */
  columns: number;
  tile: [number, number];
}

const pct = (x: number) => `${Math.round(x * 100)} %`;
const sign = (x: number) => (x > 0 ? "+" : x < 0 ? "−" : "") + pct(Math.abs(x)).replace(" %", " %");

// ----- Korrekturen -----------------------------------------------------------------------
export const SHARPEN: GalleryItem[] = [-0.5, -0.25, 0, 0.25, 0.5].map((v) => ({
  label: v < 0 ? `Weichzeichnen: ${pct(-v)}` : v > 0 ? `Schärfen: ${pct(v)}` : "Schärfen und Weichzeichnen: 0 %",
  format: { sharpen: v || undefined },
}));
export const BRIGHT_CONTRAST: GalleryItem[] = [-0.4, -0.2, 0, 0.2, 0.4].flatMap((c) => [-0.4, -0.2, 0, 0.2, 0.4].map((b) => ({
  label: `Helligkeit: ${sign(b)}${b === 0 ? " (Normal)" : ""} Kontrast: ${sign(c)}${c === 0 ? " (Normal)" : ""}`,
  format: { brightness: b || undefined, contrast: c || undefined },
})));

// ----- Farbe -----------------------------------------------------------------------------
export const SATURATION: GalleryItem[] = [0, 0.33, 0.66, 1, 2, 3, 4].map((v) => ({
  label: `Sättigung: ${pct(v)}`, format: { saturation: v === 1 ? undefined : v },
}));
export const TEMPERATURE: GalleryItem[] = [4700, 5300, 5900, 6500, 7200, 8800, 11200].map((k) => ({
  label: `Temperatur: ${k.toLocaleString("de-DE")} K`, format: { temperature: k === 6500 ? undefined : k },
}));

// ----- Transparenz -----------------------------------------------------------------------
export const TRANSPARENCY: GalleryItem[] = [0, 0.15, 0.3, 0.5, 0.65, 0.8, 0.95].map((t) => ({
  label: `Transparenz: ${pct(t)}`, format: { transparency: t || undefined },
}));

/** Neu einfärben: Word's entries for the document's theme colours. */
export function recolorGallery(t: Readonly<Record<string, string>>): GalleryItem[] {
  const names = ["Text 2", "Akzent 1", "Akzent 2", "Akzent 3", "Akzent 4", "Akzent 5", "Akzent 6"];
  const keys = ["dk2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6"];
  const lightBg = ["lt2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6"];
  const mix = (hex: string, k: number, toWhite: boolean) => {
    const v = [0, 2, 4].map((i) => parseInt(hex.slice(i, i + 2), 16));
    const out = v.map((x) => Math.round(toWhite ? x + (255 - x) * (1 - k) : x * k));
    return out.map((x) => x.toString(16).padStart(2, "0")).join("").toUpperCase();
  };
  return [
    { label: "Keine Neueinfärbung", format: { recolor: undefined, lum: undefined } },
    { label: "Graustufe", format: { recolor: { kind: "grayscale" } } },
    { label: "Sepia", format: { recolor: { kind: "sepia" } } },
    { label: "Ausgewaschen", format: { recolor: { kind: "washout" } } },
    { label: "Schwarzweiß: 25 %", format: { recolor: { kind: "bw", threshold: 0.25 } } },
    { label: "Schwarzweiß: 50 %", format: { recolor: { kind: "bw", threshold: 0.5 } } },
    { label: "Schwarzweiß: 75 %", format: { recolor: { kind: "bw", threshold: 0.75 } } },
    ...keys.map((k, i) => ({ label: `${names[i]}, dunkle Farbe`, format: { recolor: { kind: "duotone" as const, dark: "000000", light: mix(t[k] ?? "4472C4", 0.45, true) } } })),
    ...lightBg.map((k, i) => ({ label: `${names[i] === "Text 2" ? "Hintergrund 2" : names[i]}, helle Farbe`, format: { recolor: { kind: "duotone" as const, dark: mix(t[k] ?? "4472C4", 0.45, false), light: "FFFFFF" } } })),
  ];
}

/** The picture's current settings with an entry's applied (for previews and applying). */
export function withItem(f: PictureFormat, item: GalleryItem): PictureFormat {
  return { ...f, ...item.format };
}

/** Is the entry what the picture has now (Word frames it in the gallery)? */
function current(f: PictureFormat, item: GalleryItem): boolean {
  return Object.entries(item.format).every(([k, v]) => JSON.stringify((f as Record<string, unknown>)[k]) === JSON.stringify(v));
}

/**
 * A gallery popup for the selected picture: sections of tiles, each the picture with the
 * entry; hover previews it in the document, a click applies it.
 */
export function galleryPanel(e: Editor, sections: GallerySection[], links: MenuEntry[], close: () => void, apply: (item: GalleryItem) => void): HTMLElement {
  const sel = e.selectedImage();
  const image = sel?.obj.image;
  const shape = sel?.obj.shape;
  const base = e.pictureFormat() ?? {};
  const panel = h("div", { class: "rpanel picture-gallery" });
  let previewing = false;
  const restore = () => { if (previewing) { e.setPictureFormat(() => base, false); previewing = false; } };
  for (const s of sections) {
    if (s.title) panel.append(h("div", { class: "title" }, s.title));
    const grid = h("div", { class: "tiles", style: `grid-template-columns:repeat(${s.columns}, auto)` });
    for (const item of s.items) {
      const [w, hh] = s.tile;
      const tile = h("button", {
        class: `tile pic-tile${current(base, item) ? " on" : ""}`, "data-tip": item.label, "aria-label": item.label,
        style: `width:${w + 8}px;height:${hh + 8}px`,
        onclick: () => { restore(); close(); apply(item); },
      });
      if (image) {
        tile.append(previewElement(imageUrl(image), w, hh, withItem(base, item)));
      } else if (shape) {
        // A shape's tiles show a small rectangle-ish preview of its own geometry.
        const el = shapeElement(w * 0.75, hh * 0.75, withItem(base, item), shape);
        el.style.position = "relative";
        el.style.margin = `${hh * 0.125}px ${w * 0.125}px`;
        tile.append(el);
      }
      tile.addEventListener("mouseenter", () => { e.setPictureFormat(withItem(base, item), false); previewing = true; });
      tile.addEventListener("mouseleave", restore);
      grid.append(tile);
    }
    panel.append(grid);
  }
  // The popup closing any other way ends the preview too.
  const watch = new MutationObserver(() => { if (!panel.isConnected) { restore(); watch.disconnect(); } });
  setTimeout(() => watch.observe(document.body, { childList: true, subtree: true }));
  if (links.length) {
    panel.append(h("hr"));
    for (const l of links) {
      if (l === "-" || !("run" in l)) continue;
      panel.append(h("button", { class: "link", disabled: !!l.disabled, onclick: () => { restore(); closePopup(); l.run(); } }, l.label));
    }
  }
  return panel;
}
