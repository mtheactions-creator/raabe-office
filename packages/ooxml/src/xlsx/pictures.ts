// Pictures of a sheet's drawing with their Bildformat (D-087): the DrawingML of an xdr:pic read into
// text-engine's PictureFormat (the shared reader in ../drawingml.ts, as RaabeWrite reads Word's) and a
// changed picture written back the way Excel 365 writes it (recorded with
// tools/feel-lab/probe_excel_pictures.py, fixtures xlsx-pictures.xlsx and excel-pictures.json).
//
// Excel stores the Office 2010 effects (a14:imgEffect: Korrekturen, Sättigung, Farbton, artistic
// effects, background removal) worked into the picture's pixels, with the untouched original as an
// HD Photo beside it (a14:imgLayer r:embed), and draws the stored picture as it is: a file whose
// effects do not match its picture shows the picture, whatever the effects say (recorded).

import type { PictureFormat } from "@raabe/text-engine";
import type { SheetPicture } from "@raabe/sheet-engine";
import { blipEffectsXml, imgEffectsXml, readPicture, shapePropsXml } from "../drawingml";
import { canon } from "./styles";

const EMU_PER_PX = 9525;
const NS = {
  xdr: "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
  a: "http://schemas.openxmlformats.org/drawingml/2006/main",
  r: "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
  a14: "http://schemas.microsoft.com/office/drawing/2010/main",
  mc: "http://schemas.openxmlformats.org/markup-compatibility/2006",
};

/** The settings of PictureFormat that Office stores as a14 effects (worked into the pixels). */
export const BAKED_KEYS = ["brightness", "contrast", "saturation", "temperature", "sharpen"] as const;

/** The a14 part of a picture's format as a key ("" when it has none). */
export function bakedKey(f: PictureFormat | undefined): string {
  if (!f) return "";
  const o: Record<string, number> = {};
  for (const k of BAKED_KEYS) {
    const v = f[k];
    if (v === undefined || (k === "saturation" && v === 1) || (k === "temperature" && v === 6500) || (k !== "saturation" && k !== "temperature" && !v)) continue;
    o[k] = Math.round(v * 1000) / 1000;
  }
  return Object.keys(o).length ? canon(o) : "";
}

/** What decides whether a picture's anchor must be rewritten (not its place: that is patched). */
export function lookKey(p: SheetPicture): string {
  return canon({
    format: p.format ?? null, alt: p.alt ?? null, title: p.title ?? null, hidden: !!p.hidden, lock: !!p.lockAspect, name: p.name ?? null,
    same: !p.media || p.media.src === p.src, baked: p.baked ?? null, bakedSrc: p.bakedSrc ? p.bakedSrc.length : null,
  });
}

const attr = (s: string, name: string) => new RegExp(`\\b${name}="([^"]*)"`).exec(s)?.[1];
const unesc = (s: string) => s.replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&amp;/g, "&");
const esc = (s: string) => s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");

/**
 * The Bildformat of an anchor holding an xdr:pic: its format (parsed with the shared DrawingML
 * reader where a DOM parser exists), Alternativtext, Ausgeblendet, Seitenverhältnis sperren and
 * the relationships of its picture. `decls` are the drawing root's namespace declarations.
 */
export function readPictureLook(ax: string, decls: string, theme: Record<string, string>): Partial<SheetPicture> & { embed?: string } {
  const out: Partial<SheetPicture> & { embed?: string } = {};
  const cnv = /<xdr:cNvPr\b[^>]*>/.exec(ax)?.[0] ?? "";
  const descr = attr(cnv, "descr"), title = attr(cnv, "title");
  if (descr) out.alt = unesc(descr);
  if (title) out.title = unesc(title);
  if (attr(cnv, "hidden") === "1" || attr(cnv, "hidden") === "true") out.hidden = true;
  if (/<a:picLocks\b[^>]*\bnoChangeAspect="1"/.test(ax)) out.lockAspect = true;
  const blip = /<a:blip\b[^>]*>/.exec(ax)?.[0] ?? "";
  out.embed = attr(blip, "r:embed");
  const layer = /<a14:imgLayer\b[^>]*\br:embed="([^"]*)"/.exec(ax)?.[1];
  if (typeof DOMParser !== "undefined") {
    try {
      const doc = new DOMParser().parseFromString(`<xdr:wsDr ${decls}>${ax}</xdr:wsDr>`, "application/xml");
      const pic = doc.getElementsByTagNameNS(NS.xdr, "pic")[0];
      if (pic && !doc.getElementsByTagName("parsererror").length) {
        const f = readPicture(pic, theme) as Record<string, unknown>;
        // The frame's shape alone is no format; a plain rectangle is none.
        for (const k of Object.keys(f)) if (f[k] === undefined) delete f[k];
        if (Object.keys(f).length) out.format = f;
      }
    } catch {
      // Unreadable DrawingML: the picture is shown plainly and written back as it was.
    }
  }
  // Effects worked into the stored picture: shown as stored (a14 image layer present).
  if (layer || /<a14:imgEffect\b/.test(ax)) {
    const k = bakedKey(out.format as PictureFormat | undefined);
    if (k || /<a14:imgEffect\b/.test(ax)) out.baked = k;
  }
  if (out.embed) out.media = { rid: out.embed, layer, src: "" };
  return out;
}

/** The drawing root's namespace declarations (xmlns:… attributes), with the ones a picture needs. */
export function rootDecls(xml: string): string {
  const head = /<xdr:wsDr\b[^>]*>/.exec(xml)?.[0] ?? "";
  const found = new Map<string, string>();
  for (const m of head.matchAll(/\bxmlns(?::(\w+))?="([^"]*)"/g)) found.set(m[1] ?? "", m[2]);
  for (const [k, v] of Object.entries(NS)) if (!found.has(k)) found.set(k, v);
  return [...found].map(([k, v]) => (k ? `xmlns:${k}="${v}"` : `xmlns="${v}"`)).join(" ");
}

const pointXml = (tag: string, p: { col: number; row: number; x: number; y: number }) =>
  `<xdr:${tag}><xdr:col>${p.col}</xdr:col><xdr:colOff>${Math.round(p.x * EMU_PER_PX)}</xdr:colOff><xdr:row>${p.row}</xdr:row><xdr:rowOff>${Math.round(p.y * EMU_PER_PX)}</xdr:rowOff></xdr:${tag}>`;

const LOCAL_DPI = '<a:ext uri="{28A0092B-C50C-407E-A947-70E740481C1C}"><a14:useLocalDpi xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main" val="0"/></a:ext>';
const IMG_PROPS = /<a:ext uri="\{BEBA8EAE-BF5A-486C-A8C5-ECC9F3942E4B\}">[\s\S]*?<\/a:ext>/;

/**
 * A picture's anchor as Excel 365 writes it. `rid` is the picture's relationship (the blip);
 * `keepFx` keeps the a14 effects of the read anchor as they were (their picture is unchanged), else
 * they are written from the format, with the layer relationship `layer` when the original is known.
 */
export function pictureAnchorXml(p: SheetPicture, id: number, rid: string, size: { w: number; h: number }, opts: { keepFx?: boolean; layer?: string } = {}): string {
  const f = (p.format ?? {}) as PictureFormat;
  const raw = p.raw ?? "";
  const cx = Math.round(size.w * EMU_PER_PX), cy = Math.round(size.h * EMU_PER_PX);
  const open = raw ? /^<xdr:(twoCellAnchor|oneCellAnchor|absoluteAnchor)\b[^>]*>/.exec(raw)?.[0] : undefined;
  const kind = p.to ? "twoCellAnchor" : "oneCellAnchor";
  const head = open && open.startsWith(`<xdr:${kind}`) ? open : kind === "twoCellAnchor" ? '<xdr:twoCellAnchor editAs="oneCell">' : "<xdr:oneCellAnchor>";
  const place = p.to ? pointXml("from", p.from) + pointXml("to", p.to) : pointXml("from", p.from) + `<xdr:ext cx="${cx}" cy="${cy}"/>`;
  // cNvPr: the kept creation id (a16) of the read anchor.
  const cnvExt = /<xdr:cNvPr\b[^>]*>([\s\S]*?)<\/xdr:cNvPr>/.exec(raw)?.[1] ?? "";
  const name = p.name ?? `Grafik ${id - 1}`;
  const cnv = `<xdr:cNvPr id="${id}" name="${esc(name)}"${p.alt ? ` descr="${esc(p.alt)}"` : ""}${p.title ? ` title="${esc(p.title)}"` : ""}${p.hidden ? ' hidden="1"' : ""}` +
    (cnvExt ? `>${cnvExt}</xdr:cNvPr>` : "/>");
  const locks = p.lockAspect ? '<a:picLocks noChangeAspect="1"/>' : "<a:picLocks/>";
  const kept = opts.keepFx ? IMG_PROPS.exec(raw)?.[0] ?? "" : "";
  const fx = kept || imgEffectsXml(f, opts.layer);
  const blip = `<a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="${rid}">${blipEffectsXml(f)}<a:extLst>${fx}${LOCAL_DPI}</a:extLst></a:blip>`;
  const c = f.crop;
  const pc = (v: number) => Math.round(v * 100000);
  const src = c && (c.l || c.t || c.r || c.b)
    ? `<a:srcRect${c.l ? ` l="${pc(c.l)}"` : ""}${c.t ? ` t="${pc(c.t)}"` : ""}${c.r ? ` r="${pc(c.r)}"` : ""}${c.b ? ` b="${pc(c.b)}"` : ""}/>` : "";
  // The shape's transform keeps Excel's offset (Excel places the picture by the anchor).
  const off = /<a:xfrm\b[^>]*>\s*(<a:off\b[^>]*\/>)/.exec(raw)?.[1];
  let sp = shapePropsXml(f, cx, cy);
  if (off) sp = sp.replace('<a:off x="0" y="0"/>', off);
  return `${head}${place}<xdr:pic><xdr:nvPicPr>${cnv}<xdr:cNvPicPr>${locks}</xdr:cNvPicPr></xdr:nvPicPr>` +
    `<xdr:blipFill>${blip}${src}<a:stretch><a:fillRect/></a:stretch></xdr:blipFill><xdr:spPr>${sp}</xdr:spPr></xdr:pic><xdr:clientData/></xdr:${head.slice(5, head.search(/[\s>]/))}>`;
}
