// Pictures of a sheet (the school logo in a template …) shown where the file anchors them: over
// the cells, following scrolling and zoom, behind charts. They are shown only; the file keeps
// them as they were.

import { h } from "@raabe/shell";
import type { SheetPicture } from "@raabe/sheet-engine";
import type { CalcApp } from "./app";

export function picturesFeature(app: CalcApp) {
  const layer = h("div", { class: "picture-layer" });
  let shown: SheetPicture[] | undefined;
  const imgs: HTMLImageElement[] = [];

  /** Called after every paint of the grid. */
  function place(): void {
    const list = app.sheet.pictures;
    if (!list?.length) { if (layer.isConnected) layer.remove(); shown = undefined; return; }
    if (!layer.isConnected) {
      const canvas = app.grid.el.querySelector(".grid-canvas");
      app.grid.el.insertBefore(layer, canvas ? canvas.nextSibling : app.grid.el.firstChild);
    }
    if (shown !== list) {
      layer.replaceChildren();
      imgs.length = 0;
      for (const p of list) {
        const img = h("img", { src: p.src, alt: p.name ?? "", draggable: "false" }) as HTMLImageElement;
        imgs.push(img);
        layer.append(img);
      }
      shown = list;
    }
    const v = app.sheet.view;
    const z = v.zoom / 100;
    // The layer covers the cells only (not the headers).
    const tl = app.grid.rectOnCanvas({ r0: v.top, c0: v.left, r1: v.top, c1: v.left });
    if (!tl) return;
    Object.assign(layer.style, { left: `${tl.x}px`, top: `${tl.y}px` });
    list.forEach((p, i) => {
      const img = imgs[i];
      const a = app.grid.rectOnCanvas({ r0: p.from.row, c0: p.from.col, r1: p.from.row, c1: p.from.col });
      if (!a) { img.style.display = "none"; return; }
      const x = a.x + p.from.x * z, y = a.y + p.from.y * z;
      let w = (p.size?.w ?? 0) * z, hh = (p.size?.h ?? 0) * z;
      if (p.to) {
        const b = app.grid.rectOnCanvas({ r0: p.to.row, c0: p.to.col, r1: p.to.row, c1: p.to.col });
        if (b) { w = b.x + p.to.x * z - x; hh = b.y + p.to.y * z - y; }
      }
      Object.assign(img.style, { display: w > 0 && hh > 0 ? "" : "none", left: `${x - tl.x}px`, top: `${y - tl.y}px`, width: `${w}px`, height: `${hh}px` });
    });
  }

  return { place };
}
