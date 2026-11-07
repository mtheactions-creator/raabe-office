// Datei > Exportieren > Dateityp ändern (D-072): the presentation types (via Speichern unter) and the slides as
// pictures, PNG or JPEG, as PowerPoint offers them: this slide, or all slides (PowerPoint puts those into a
// folder; we write them beside the chosen name as "Name-Folie1.png" …). The pictures are our renderer's drawing
// of each slide (the same as Folienzoom's previews), 1920 pixels wide.

import { h, showDialog } from "@raabe/shell";
import type { PresentApp } from "./app";
import { slidePng } from "./zoom";

async function toJpeg(png: Uint8Array): Promise<Uint8Array | null> {
  const url = URL.createObjectURL(new Blob([png as BlobPart], { type: "image/png" }));
  try {
    const img = await new Promise<HTMLImageElement | null>((res) => { const i = new Image(); i.onload = () => res(i); i.onerror = () => res(null); i.src = url; });
    if (!img) return null;
    const c = document.createElement("canvas");
    c.width = img.naturalWidth; c.height = img.naturalHeight;
    const g = c.getContext("2d");
    if (!g) return null;
    g.fillStyle = "#fff"; g.fillRect(0, 0, c.width, c.height); g.drawImage(img, 0, 0);
    const blob = await new Promise<Blob | null>((r) => c.toBlob(r, "image/jpeg", 0.92));
    return blob ? new Uint8Array(await blob.arrayBuffer()) : null;
  } finally { URL.revokeObjectURL(url); }
}

/** The slides as pictures: asks which slides, then where. */
export async function exportPictures(app: PresentApp, kind: "png" | "jpg"): Promise<void> {
  const which = await showDialog({ title: "RaabePresent", body: "Welche Folien möchten Sie exportieren?",
    buttons: [{ label: "Alle Folien", value: "all", primary: true }, { label: "Nur diese", value: "one" }, { label: "Abbrechen", value: "" }], cancel: "" });
  if (!which) return;
  const name = app.docName();
  const filters = [kind === "png" ? { name: "PNG (Portable Network Graphics)", extensions: ["png"] } : { name: "JPEG-Dateiaustauschformat", extensions: ["jpg"] }];
  const target = app.platform.saveAnyDialog ? await app.platform.saveAnyDialog("Exportieren", `${name}.${kind}`, filters) : null;
  if (!target) return;
  const base = target.replace(/\.(png|jpe?g)$/i, "");
  const slides = which === "all" ? app.deck.slides.map((s, i) => ({ path: s.path, n: i + 1 })) : [{ path: app.deck.slides[app.slideNumber() - 1].path, n: app.slideNumber() }];
  for (const s of slides) {
    const png = await slidePng(app, s.path, 1920);
    if (!png) continue;
    const bytes = kind === "png" ? png : await toJpeg(png);
    if (!bytes) continue;
    await app.platform.writeFile(which === "all" ? `${base}-Folie${s.n}.${kind}` : `${base}.${kind}`, bytes);
  }
}

/** The Exportieren page's Dateityp ändern list. */
export function fileTypeSection(app: PresentApp): HTMLElement {
  const item = (label: string, note: string, run: () => void) => h("button", { class: "bigbtn", style: "width:360px;text-align:left", onclick: run }, h("b", null, label), h("div", { class: "muted" }, note));
  return h("div", null, h("h2", null, "Dateityp ändern"),
    item("Präsentation (*.pptx)", "Das übliche Format für Präsentationen.", () => void app.saveAs()),
    item("PowerPoint-Vorlage (*.potx)", "Ein Ausgangspunkt für neue Präsentationen.", () => void app.saveAs()),
    item("PowerPoint-Bildschirmpräsentation (*.ppsx)", "Beginnt beim Öffnen gleich mit der Bildschirmpräsentation.", () => void app.saveAs()),
    item("PNG (Portable Network Graphics) (*.png)", "Jede Folie als Bilddatei in hoher Qualität.", () => void exportPictures(app, "png")),
    item("JPEG-Dateiaustauschformat (*.jpg)", "Jede Folie als kleinere Bilddatei.", () => void exportPictures(app, "jpg")));
}
