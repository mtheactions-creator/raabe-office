// Video und Audio from files on this device (D-070): embedded as PowerPoint 2010+ embeds them (the file
// in ppt/media with an a:videoFile / a:audioFile link and a p14:media embed, a poster picture as the
// shape's picture: the video's first frame, a speaker symbol for audio). In Normalansicht the poster is
// shown; in the show the file plays with its controls (PowerPoint's own playback bar is not copied).

import { PptxDom } from "@raabe/ooxml";
import type { PresentApp } from "./app";
import * as S from "./shapeops";

const RT = {
  media: "http://schemas.microsoft.com/office/2007/relationships/media",
  video: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/video",
  audio: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/audio",
  image: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
};
const TYPES: Record<string, string> = { mp4: "video/mp4", m4v: "video/mp4", mov: "video/quicktime", webm: "video/webm", wmv: "video/x-ms-wmv", avi: "video/x-msvideo",
  mp3: "audio/mpeg", m4a: "audio/mp4", wav: "audio/wav", wma: "audio/x-ms-wma", ogg: "audio/ogg" };

async function posterOfVideo(url: string): Promise<{ png: Uint8Array; w: number; h: number }> {
  const v = document.createElement("video");
  v.muted = true;
  v.src = url;
  await new Promise<void>((res, rej) => { v.onloadeddata = () => res(); v.onerror = () => rej(new Error("Das Video kann nicht gelesen werden.")); });
  v.currentTime = Math.min(0.1, v.duration || 0);
  await new Promise((r) => { v.onseeked = r; setTimeout(r, 500); });
  const c = document.createElement("canvas");
  c.width = v.videoWidth || 640; c.height = v.videoHeight || 360;
  c.getContext("2d")!.drawImage(v, 0, 0, c.width, c.height);
  return { png: await canvasPng(c), w: c.width, h: c.height };
}

async function speakerIcon(): Promise<Uint8Array> {
  const c = document.createElement("canvas");
  c.width = c.height = 96;
  const g = c.getContext("2d")!;
  g.fillStyle = "#4472C4";
  g.beginPath(); g.moveTo(14, 36); g.lineTo(34, 36); g.lineTo(58, 14); g.lineTo(58, 82); g.lineTo(34, 60); g.lineTo(14, 60); g.closePath(); g.fill();
  g.strokeStyle = "#4472C4"; g.lineWidth = 6;
  for (const r of [14, 26]) { g.beginPath(); g.arc(58, 48, r, -0.9, 0.9); g.stroke(); }
  return canvasPng(c);
}

async function canvasPng(c: HTMLCanvasElement): Promise<Uint8Array> {
  const blob = await new Promise<Blob>((r) => c.toBlob((b) => r(b!), "image/png"));
  return new Uint8Array(await blob.arrayBuffer());
}

export async function insertMedia(app: PresentApp, kind: "video" | "audio"): Promise<void> {
  const exts = kind === "video" ? ["mp4", "m4v", "mov", "webm", "wmv", "avi"] : ["mp3", "m4a", "wav", "wma", "ogg"];
  const path = await app.platform.openAnyDialog?.(kind === "video" ? "Video einfügen" : "Audio einfügen", [{ name: kind === "video" ? "Videodateien" : "Audiodateien", extensions: exts }]);
  if (!path) return;
  const bytes = await app.platform.readFile(path);
  const ext = path.slice(path.lastIndexOf(".") + 1).toLowerCase();
  const type = TYPES[ext] ?? "application/octet-stream";
  const url = URL.createObjectURL(new Blob([bytes as BlobPart], { type }));
  let poster: { png: Uint8Array; w: number; h: number };
  try {
    poster = kind === "video" ? await posterOfVideo(url) : { png: await speakerIcon(), w: 64, h: 64 };
  } catch {
    poster = { png: await speakerIcon(), w: 640, h: 360 };
  }
  URL.revokeObjectURL(url);
  const pptx = app.deck.pptx;
  const slide = app.slidePath()!;
  const media = pptx.freeName(`ppt/media/media#.${ext}`);
  const image = pptx.freeName("ppt/media/image#.png");
  const size = pptx.slideSize;
  const k = Math.min(1, (size.cx * 0.8) / (poster.w * 9525), (size.cy * 0.8) / (poster.h * 9525));
  const w = Math.round(poster.w * 9525 * k), h = Math.round(poster.h * 9525 * k);
  const base = kind === "video" ? "Video" : "Audio";
  let id = 0;
  app.editor.edit(kind === "video" ? "Video einfügen" : "Audio einfügen", () => {
    pptx.put(media, bytes);
    pptx.setType(media, type, true);
    pptx.put(image, poster.png);
    pptx.setType(image, "image/png", true);
    const rels = pptx.rels(slide);
    const rMedia = rels.ensure(RT.media, `../media/${media.split("/").pop()}`);
    const rLink = rels.ensure(kind === "video" ? RT.video : RT.audio, `../media/${media.split("/").pop()}`);
    const rImg = pptx.link(slide, RT.image, image);
    const doc = pptx.doc(slide);
    const tree = S.treeOf(doc);
    id = S.maxId(tree) + 1;
    const file = kind === "video" ? "videoFile" : "audioFile";
    tree.append(PptxDom.fragment(doc, `<p:pic><p:nvPicPr><p:cNvPr id="${id}" name="${base} ${id - 1}"><a:hlinkClick r:id="" action="ppaction://media"/></p:cNvPr><p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr>` +
      `<p:nvPr><a:${file} r:link="${rLink}"/><p:extLst><p:ext uri="{DAA4B4D4-6D71-4841-9C94-3DE7FCFB9230}"><p14:media xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" r:embed="${rMedia}"/></p:ext></p:extLst></p:nvPr></p:nvPicPr>` +
      `<p:blipFill><a:blip r:embed="${rImg}"/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr><a:xfrm><a:off x="${Math.round((size.cx - w) / 2)}" y="${Math.round((size.cy - h) / 2)}"/><a:ext cx="${w}" cy="${h}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>`)[0]);
  });
  app.editor.select([id]);
}

/** The media file of a picture shape (video or audio), if it has one. */
export function mediaOf(app: PresentApp, slidePath: string, el: Element): { path: string; kind: "video" | "audio" } | null {
  const nvPr = PptxDom.at(el, "nvPicPr/nvPr");
  const file = PptxDom.kid(nvPr, "videoFile") ?? PptxDom.kid(nvPr, "audioFile");
  if (!file) return null;
  const embed = nvPr?.getElementsByTagNameNS("http://schemas.microsoft.com/office/powerpoint/2010/main", "media")[0];
  const rid = embed?.getAttributeNS(PptxDom.NS_R, "embed") ?? file.getAttributeNS(PptxDom.NS_R, "link") ?? "";
  const path = app.deck.pptx.target(slidePath, rid);
  return path ? { path, kind: file.localName === "videoFile" ? "video" : "audio" } : null;
}

export function mediaType(path: string): string {
  return TYPES[path.slice(path.lastIndexOf(".") + 1).toLowerCase()] ?? "application/octet-stream";
}
