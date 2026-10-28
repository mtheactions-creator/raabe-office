// Pictures of a sheet (the school logo in a template …) shown where the file anchors them: over
// the cells, following scrolling and zoom, behind charts. Einfügen > Bilder > Dieses Gerät adds
// one at the active cell at its natural size (as Excel); a click selects it (frame and corner
// handles), dragging moves it, a corner sizes it keeping its proportions, Entf deletes it. Each
// change is one undo step; the file writes moved pictures with new corners and keeps the rest.

import { h, messageBox } from "@raabe/shell";
import type { SheetPicture } from "@raabe/sheet-engine";
import type { CalcApp } from "./app";

const MIME: Record<string, string> = { png: "image/png", jpg: "image/jpeg", jpeg: "image/jpeg", gif: "image/gif", bmp: "image/bmp", svg: "image/svg+xml" };

export function picturesFeature(app: CalcApp) {
  const layer = h("div", { class: "picture-layer" });
  let shown: SheetPicture[] | undefined;
  const imgs: HTMLElement[] = [];
  let selected: SheetPicture | null = null;
  let live: { p: SheetPicture; dx: number; dy: number; w: number; h: number } | null = null;

  const z = () => app.sheet.view.zoom / 100;

  /** A point on the canvas (CSS px) as a cell and the offset in it at 100 %. */
  function anchorAt(x: number, y: number): { col: number; row: number; x: number; y: number } {
    const hit = app.grid.hitCell(x, y);
    const c = app.grid.rectOnCanvas({ r0: hit.row, c0: hit.col, r1: hit.row, c1: hit.col })!;
    return { col: hit.col, row: hit.row, x: Math.max(0, (x - c.x) / z()), y: Math.max(0, (y - c.y) / z()) };
  }

  /** Where a picture is on the canvas (CSS px). */
  function boxOf(p: SheetPicture): { x: number; y: number; w: number; h: number } | null {
    const a = app.grid.rectOnCanvas({ r0: p.from.row, c0: p.from.col, r1: p.from.row, c1: p.from.col });
    if (!a) return null;
    const x = a.x + p.from.x * z(), y = a.y + p.from.y * z();
    let w = (p.size?.w ?? 0) * z(), hh = (p.size?.h ?? 0) * z();
    if (p.to) {
      const b = app.grid.rectOnCanvas({ r0: p.to.row, c0: p.to.col, r1: p.to.row, c1: p.to.col });
      if (b) { w = b.x + p.to.x * z() - x; hh = b.y + p.to.y * z() - y; }
    }
    return { x, y, w, h: hh };
  }

  /** The picture put at a box on the canvas (both corners, and its size at 100 %). */
  function placeAt(p: SheetPicture, b: { x: number; y: number; w: number; h: number }): SheetPicture {
    return { ...p, from: anchorAt(b.x, b.y), to: anchorAt(b.x + b.w, b.y + b.h), size: { w: b.w / z(), h: b.h / z() } };
  }

  function change(label: string, fn: (list: SheetPicture[]) => SheetPicture[]): void {
    const sheet = app.sheet;
    app.change(label, () => {
      app.history.whole(sheet);
      const next = fn([...(sheet.pictures ?? [])]);
      sheet.pictures = next.length ? next : undefined;
    });
  }

  /** Called after every paint of the grid. */
  function place(): void {
    const list = app.sheet.pictures;
    if (!list?.length) { if (layer.isConnected) layer.remove(); shown = undefined; selected = null; return; }
    if (!layer.isConnected) {
      const canvas = app.grid.el.querySelector(".grid-canvas");
      app.grid.el.insertBefore(layer, canvas ? canvas.nextSibling : app.grid.el.firstChild);
    }
    if (shown !== list) {
      if (selected && !list.includes(selected)) selected = list.find((p) => p.raw === selected!.raw && p.src === selected!.src) ?? null;
      layer.replaceChildren();
      imgs.length = 0;
      for (const p of list) {
        const el = h("div", { class: "pic" }, h("img", { src: p.src, alt: p.name ?? "", draggable: "false" }),
          ...["nw", "ne", "sw", "se"].map((c) => h("span", { class: `pic-handle ${c}`, "data-corner": c })));
        el.addEventListener("pointerdown", (e) => down(e, p, el));
        imgs.push(el);
        layer.append(el);
      }
      shown = list;
    }
    const v = app.sheet.view;
    const tl = app.grid.rectOnCanvas({ r0: v.top, c0: v.left, r1: v.top, c1: v.left });
    if (!tl) return;
    Object.assign(layer.style, { left: `${tl.x}px`, top: `${tl.y}px` });
    list.forEach((p, i) => {
      const el = imgs[i];
      const b = boxOf(p);
      if (!b) { el.style.display = "none"; return; }
      const d = live?.p === p ? live : { dx: 0, dy: 0, w: b.w, h: b.h };
      Object.assign(el.style, { display: b.w > 0 && b.h > 0 ? "" : "none", left: `${b.x + d.dx - tl.x}px`, top: `${b.y + d.dy - tl.y}px`, width: `${d.w}px`, height: `${d.h}px` });
      el.classList.toggle("on", p === selected);
    });
  }

  function select(p: SheetPicture | null): void {
    selected = p;
    shown = undefined;
    app.grid.contentChanged();
  }

  /** Press on a picture: select it, then move it or size it by a corner. */
  function down(e: PointerEvent, p: SheetPicture, el: HTMLElement): void {
    if (e.button !== 0) return;
    e.preventDefault();
    e.stopPropagation();
    if (selected !== p) { selected = p; for (const x of imgs) x.classList.toggle("on", x === el); }
    const b = boxOf(p);
    if (!b) return;
    const corner = (e.target as HTMLElement).dataset.corner;
    const sx = e.clientX, sy = e.clientY;
    el.setPointerCapture(e.pointerId);
    live = { p, dx: 0, dy: 0, w: b.w, h: b.h };
    const move = (ev: PointerEvent) => {
      const dx = ev.clientX - sx, dy = ev.clientY - sy;
      if (!corner) live = { p, dx, dy, w: b.w, h: b.h };
      else {
        // Corners keep the proportions (Excel's default for pictures).
        const sw = corner.includes("e") ? b.w + dx : b.w - dx;
        const k = Math.max(8 / b.w, sw / b.w);
        const w = b.w * k, hh = b.h * k;
        live = { p, dx: corner.includes("w") ? b.w - w : 0, dy: corner.includes("n") ? b.h - hh : 0, w, h: hh };
      }
      place();
    };
    const up = () => {
      el.removeEventListener("pointermove", move);
      el.removeEventListener("pointerup", up);
      const l = live;
      live = null;
      if (!l || (Math.abs(l.dx) < 1 && Math.abs(l.dy) < 1 && Math.abs(l.w - b.w) < 1)) { place(); return; }
      const moved = placeAt(p, { x: b.x + l.dx, y: b.y + l.dy, w: l.w, h: l.h });
      change(corner ? "Größe ändern" : "Verschieben", (list) => list.map((q) => (q === p ? moved : q)));
      selected = moved;
    };
    el.addEventListener("pointermove", move);
    el.addEventListener("pointerup", up);
  }

  // A click elsewhere ends the selection; Entf deletes the selected picture.
  addEventListener("pointerdown", (e) => { if (selected && !(e.target as HTMLElement).closest?.(".pic")) select(null); }, true);
  addEventListener("keydown", (e) => {
    if (!selected || app.grid.isEditing) return;
    if (e.key === "Delete" || e.key === "Backspace") {
      e.preventDefault();
      e.stopPropagation();
      const p = selected;
      selected = null;
      change("Bild löschen", (list) => list.filter((q) => q !== p));
    } else if (e.key === "Escape") select(null);
  }, true);

  /** Einfügen > Bilder > Dieses Gerät: the picture at the active cell, at its natural size. */
  async function insertFromFile(): Promise<void> {
    const path = await app.platform.openImageDialog();
    if (!path) return;
    try {
      const bytes = await app.platform.readFile(path);
      const ext = path.split(".").pop()?.toLowerCase() ?? "";
      const type = bytes[0] === 0x89 && bytes[1] === 0x50 ? "image/png" : bytes[0] === 0xff && bytes[1] === 0xd8 ? "image/jpeg" : MIME[ext] ?? "image/png";
      let bin = "";
      for (let i = 0; i < bytes.length; i += 0x8000) bin += String.fromCharCode(...bytes.subarray(i, i + 0x8000));
      const src = `data:${type};base64,${btoa(bin)}`;
      const size = await new Promise<{ w: number; h: number }>((resolve, reject) => {
        const img = new Image();
        img.onload = () => resolve({ w: img.naturalWidth, h: img.naturalHeight });
        img.onerror = () => reject(new Error("Kein lesbares Bild."));
        img.src = src;
      });
      const a = app.sheet.view.active;
      const c = app.grid.rectOnCanvas({ r0: a.row, c0: a.col, r1: a.row, c1: a.col });
      const name = `Grafik ${(app.sheet.pictures?.length ?? 0) + 1}`;
      let p: SheetPicture = { from: { col: a.col, row: a.row, x: 0, y: 0 }, size, src, name };
      if (c) p = placeAt(p, { x: c.x, y: c.y, w: size.w * z(), h: size.h * z() });
      change("Bild einfügen", (list) => [...list, p]);
      selected = p;
    } catch (e) {
      await messageBox("RaabeCalc", `Das Bild konnte nicht eingefügt werden.\n\n${e instanceof Error ? e.message : String(e)}`);
    }
  }

  return { place, insertFromFile, insertPicture: (p: SheetPicture) => change("Bild einfügen", (list) => [...list, p]) };
}
