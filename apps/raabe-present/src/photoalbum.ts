// Einfügen > Fotoalbum (D-072): pictures chosen together become a new presentation, a title slide
// "Fotoalbum" first, then the pictures one, two or four to a slide (Bildlayout), each fitted into its place.
// Our dialog, with PowerPoint's choices for the layout.

import { h, showDialog } from "@raabe/shell";
import type { PresentApp } from "./app";
import * as S from "./shapeops";
import * as T from "./textops";

const TYPES: Record<string, string> = { png: "image/png", jpg: "image/jpeg", jpeg: "image/jpeg", gif: "image/gif", bmp: "image/bmp", svg: "image/svg+xml" };

async function sizeOf(bytes: Uint8Array, type: string): Promise<{ w: number; h: number }> {
  const url = URL.createObjectURL(new Blob([bytes as BlobPart], { type }));
  try {
    return await new Promise((res) => { const img = new Image(); img.onload = () => res({ w: img.naturalWidth || 1, h: img.naturalHeight || 1 }); img.onerror = () => res({ w: 4, h: 3 }); img.src = url; });
  } finally { URL.revokeObjectURL(url); }
}

export async function photoAlbumDialog(app: PresentApp): Promise<void> {
  const files: string[] = [];
  const list = h("select", { size: "8", style: "width:320px" }) as HTMLSelectElement;
  const add = h("button", { onclick: async () => { const r = (await app.platform.openImagesDialog?.("Neue Bilder einfügen")) ?? []; for (const f of r) { files.push(f); list.append(h("option", null, f.split(/[/\\]/).pop() ?? f)); } } }, "Datei/Datenträger...");
  const layout = h("select", null, h("option", { value: "1" }, "An Folie anpassen"), h("option", { value: "1b" }, "1 Bild"), h("option", { value: "2" }, "2 Bilder"), h("option", { value: "4" }, "4 Bilder")) as HTMLSelectElement;
  const ok = await showDialog({ title: "Fotoalbum", body: h("div", { class: "form" }, h("b", null, "Bild einfügen aus:"), add, h("label", null, "Bilder im Album:"), list, h("label", null, "Bildlayout: ", layout)),
    buttons: [{ label: "Erstellen", value: true, primary: true }, { label: "Abbrechen", value: false }], cancel: false });
  if (!ok || !files.length) return;
  const per = layout.value === "4" ? 4 : layout.value === "2" ? 2 : 1;
  const fit = layout.value === "1";
  app.newDoc();
  const deck = app.deck;
  const size = deck.pptx.slideSize;
  // The title slide: "Fotoalbum".
  deck.change("Fotoalbum", [deck.slides[0].path], () => {
    const body = S.shapeEls(S.treeOf(deck.pptx.doc(deck.slides[0].path))).map((e) => S.txBodyOf(e)).find((b): b is Element => !!b);
    if (body) T.insertText(body, { p: 0, o: 0 }, "Fotoalbum");
  });
  const blank = deck.layouts().find((l) => l.name === "Leer")?.path ?? deck.layouts()[6].path;
  for (let i = 0; i < files.length; i += per) {
    const group = files.slice(i, i + per);
    const slide = deck.addSlide(deck.slides.length - 1, blank);
    const pics = await Promise.all(group.map(async (f) => { const bytes = await app.platform.readFile(f); const ext = f.slice(f.lastIndexOf(".") + 1).toLowerCase(); const type = TYPES[ext] ?? "image/png"; return { bytes, type, ext: type === "image/jpeg" ? "jpeg" : ext, ...(await sizeOf(bytes, type)) }; }));
    deck.change("Fotoalbum", [slide], () => {
      const doc = deck.pptx.doc(slide), tree = S.treeOf(doc);
      const margin = fit ? 0 : 457200, gap = 228600;
      const cols = per === 1 ? 1 : 2, rows = per === 4 ? 2 : 1;
      const cw = (size.cx - margin * 2 - gap * (cols - 1)) / cols, ch = (size.cy - margin * 2 - gap * (rows - 1)) / rows;
      pics.forEach((pic, k) => {
        const media = deck.pptx.freeName(`ppt/media/image#.${pic.ext}`);
        deck.pptx.put(media, pic.bytes);
        deck.pptx.setType(media, pic.type, true);
        const rid = deck.pptx.link(slide, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image", media);
        const sc = Math.min(cw / pic.w, ch / pic.h);
        const w = pic.w * sc, hh = pic.h * sc;
        const x = margin + (k % cols) * (cw + gap) + (cw - w) / 2, y = margin + Math.floor(k / cols) * (ch + gap) + (ch - hh) / 2;
        tree.append(S.newPicture(doc, S.maxId(tree) + 1, rid, { x: Math.round(x), y: Math.round(y), w: Math.round(w), h: Math.round(hh) }));
      });
    });
  }
  app.setCurrent(0);
}
