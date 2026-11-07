// Bildformat > Bilder komprimieren (D-072): pictures are resampled to the chosen resolution at the size they
// are shown (HD 330, Drucken 220, Web 150, E-Mail 96 ppi), cropped areas deleted if wanted, for the selected
// picture or all pictures of the presentation. A resampled picture becomes a new media part (JPEG stays JPEG,
// everything else PNG); parts nothing refers to any more are left out when the file is written.

import { h, showDialog } from "@raabe/shell";
import { PptxDom } from "@raabe/ooxml";
import type { PresentApp } from "./app";

const { kid } = PptxDom;
const RT_IMAGE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";

async function decode(bytes: Uint8Array, type: string): Promise<HTMLImageElement | null> {
  const url = URL.createObjectURL(new Blob([bytes as BlobPart], { type }));
  try {
    return await new Promise((res) => { const img = new Image(); img.onload = () => res(img); img.onerror = () => res(null); img.src = url; });
  } finally { setTimeout(() => URL.revokeObjectURL(url), 1000); }
}

export async function compressDialog(app: PresentApp): Promise<void> {
  const selPics = app.editor.selectedViews.filter((v) => v.kind === "pic");
  const check = (label: string, on: boolean) => { const c = h("input", { type: "checkbox" }) as HTMLInputElement; c.checked = on; return { c, el: h("label", { class: "dlg-check" }, c, ` ${label}`) }; };
  const only = check("Nur für dieses Bild übernehmen", selPics.length > 0), crop = check("Zugeschnittene Bildbereiche löschen", true);
  only.c.disabled = !selPics.length;
  const opts: [number, string][] = [[330, "HD (330 ppi)"], [220, "Drucken (220 ppi)"], [150, "Web (150 ppi)"], [96, "E-Mail (96 ppi)"], [220, "Standardauflösung verwenden"]];
  const radios = opts.map(([ppi, label], i) => { const r = h("input", { type: "radio", name: "ppi" }) as HTMLInputElement; r.checked = i === 1; return { r, ppi, el: h("label", { class: "dlg-check" }, r, ` ${label}`) }; });
  const ok = await showDialog({ title: "Bilder komprimieren", body: h("div", { class: "form" }, h("b", null, "Komprimierungsoptionen"), only.el, crop.el, h("b", null, "Auflösung"), ...radios.map((x) => x.el)),
    buttons: [{ label: "OK", value: true, primary: true }, { label: "Abbrechen", value: false }], cancel: false });
  if (!ok) return;
  const ppi = radios.find((x) => x.r.checked)?.ppi ?? 220;
  await compress(app, { ppi, deleteCrop: crop.c.checked, only: only.c.checked ? selPics.map((v) => v.id) : null });
}

export async function compress(app: PresentApp, o: { ppi: number; deleteCrop: boolean; only: number[] | null }): Promise<number> {
  const pptx = app.deck.pptx;
  const paths = o.only ? [app.slidePath()!] : app.deck.slides.map((s) => s.path);
  type Job = { slide: string; blip: Element; srcRect?: Element; bytes: Uint8Array; type: string; w: number; h: number; crop: { l: number; t: number; r: number; b: number } };
  const jobs: Job[] = [];
  for (const path of paths) {
    for (const pic of Array.from(pptx.doc(path).getElementsByTagNameNS(PptxDom.NS_P, "pic"))) {
      const id = Number(pic.getElementsByTagNameNS(PptxDom.NS_P, "cNvPr")[0]?.getAttribute("id") ?? 0);
      if (o.only && !o.only.includes(id)) continue;
      const blipFill = kid(pic, "blipFill"), blip = kid(blipFill, "blip"), xf = pic.getElementsByTagNameNS(PptxDom.NS_A, "ext")[0];
      const rid = blip?.getAttributeNS(PptxDom.NS_R, "embed");
      const media = rid ? pptx.target(path, rid) : undefined;
      const bytes = media ? pptx.files[media] : undefined;
      if (!blip || !media || !bytes || !xf || /\.(svg|emf|wmf)$/i.test(media)) continue;
      const src = kid(blipFill, "srcRect");
      const pc = (a: string) => Number(src?.getAttribute(a) ?? 0) / 100000;
      jobs.push({ slide: path, blip, srcRect: src, bytes, type: /\.jpe?g$/i.test(media) ? "image/jpeg" : "image/png",
        w: Number(xf.getAttribute("cx")) / 914400, h: Number(xf.getAttribute("cy")) / 914400, crop: { l: pc("l"), t: pc("t"), r: pc("r"), b: pc("b") } });
    }
  }
  const done: { job: Job; data: Uint8Array }[] = [];
  for (const job of jobs) {
    const img = await decode(job.bytes, job.type);
    if (!img) continue;
    const vis = { x: img.naturalWidth * job.crop.l, y: img.naturalHeight * job.crop.t, w: img.naturalWidth * (1 - job.crop.l - job.crop.r), h: img.naturalHeight * (1 - job.crop.t - job.crop.b) };
    const keepCrop = !o.deleteCrop || (!job.crop.l && !job.crop.t && !job.crop.r && !job.crop.b);
    // The whole picture's pixels at the target resolution (the visible part shown at w x h inches).
    const scale = Math.min(1, (job.w * o.ppi) / Math.max(1, vis.w));
    if (scale > 0.95 && keepCrop) continue;
    const from = keepCrop ? { x: 0, y: 0, w: img.naturalWidth, h: img.naturalHeight } : vis;
    const cw = Math.max(1, Math.round(from.w * scale)), ch = Math.max(1, Math.round(from.h * scale));
    const c = document.createElement("canvas");
    c.width = cw; c.height = ch;
    const g = c.getContext("2d");
    if (!g) continue;
    g.imageSmoothingQuality = "high";
    g.drawImage(img, from.x, from.y, from.w, from.h, 0, 0, cw, ch);
    const blob = await new Promise<Blob | null>((r) => c.toBlob(r, job.type, 0.9));
    if (!blob) continue;
    const data = new Uint8Array(await blob.arrayBuffer());
    if (data.length >= job.bytes.length && keepCrop) continue;
    done.push({ job, data });
  }
  if (!done.length) return 0;
  app.deck.change("Bilder komprimieren", [...new Set(done.map((d) => d.job.slide))], () => {
    for (const { job, data } of done) {
      const media = pptx.freeName(`ppt/media/image#.${job.type === "image/jpeg" ? "jpeg" : "png"}`);
      pptx.put(media, data);
      pptx.setType(media, job.type, true);
      const rid = pptx.link(job.slide, RT_IMAGE, media);
      job.blip.setAttributeNS(PptxDom.NS_R, "r:embed", rid);
      if (o.deleteCrop) job.srcRect?.remove();
    }
  });
  app.editor.render();
  return done.length;
}
