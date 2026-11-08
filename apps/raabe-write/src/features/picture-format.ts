// The Bildformat tab's galleries as Word 365 has them (recorded with
// tools/feel-lab/probe_word_picture.py): Korrekturen, Farbe, Transparenz, Bildformatvorlagen,
// Bildrahmen, Bildeffekte, Drehen. Each tile shows the selected picture with the setting, as
// Word's galleries do; hovering a tile previews it in the document (Livevorschau).

import { closePopup, h, type MenuEntry } from "@raabe/shell";
import { previewElement, shapeElement, type Editor, type GalleryItem, type PictureFormat } from "@raabe/text-engine";
import { imageUrl } from "@raabe/text-engine";

// The gallery entries are shared with RaabeCalc (packages/text-engine/src/picture-galleries.ts).
export { BRIGHT_CONTRAST, SATURATION, SHARPEN, TEMPERATURE, TRANSPARENCY, recolorGallery, type GalleryItem } from "@raabe/text-engine";

export interface GallerySection {
  title?: string;
  items: GalleryItem[];
  /** Tiles per row and tile size (CSS px). */
  columns: number;
  tile: [number, number];
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
