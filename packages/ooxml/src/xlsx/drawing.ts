// Charts of a worksheet: read from its drawing part (xl/drawings/drawingN.xml) into the
// sheet's chart models and written back. The drawing and every chart part stay as Excel wrote
// them while the charts are unchanged; pictures and shapes in the drawing are always kept.

import { strFromU8 } from "fflate";
import type { ChartModel } from "@raabe/charts";
import type { Sheet } from "@raabe/sheet-engine";
import { PackageWriter, Rels, relsPath, resolveTarget } from "../package";
import {
  CHART_CONTENT_TYPE, CHART_REL_TYPE, CHARTEX_CONTENT_TYPE, CHARTEX_REL_TYPE, DRAWING_CONTENT_TYPE, DRAWING_REL_TYPE,
  anchorFromXml, anchorXml, chartFromXml, chartToXml, drawingXml, isChartEx,
} from "./chart";
import { isRel } from "./read";
import { canon } from "./styles";
import { children, outer, root } from "./xml";

/** The sheet's charts (the model's optional `charts` field). */
export function chartsOf(s: Sheet): ChartModel[] | undefined {
  return (s as unknown as { charts?: ChartModel[] }).charts;
}

function setCharts(s: Sheet, charts: ChartModel[]): void {
  (s as unknown as { charts?: ChartModel[] }).charts = charts;
}

/** A chart as read: its part, its anchor's XML and the model's corners. */
interface ChartRead { xml: string; part: string; rId: string; anchor: string; place: string }

export interface DrawingRaw {
  part: string;
  xml: string;
  charts: ChartRead[];
  snap: string;
}

const rawXml = (m: ChartModel) => (m.raw as { xml?: unknown } | undefined)?.xml;
const placeOf = (m: ChartModel) => canon([m.from, m.to, m.name ?? null]);

/** Reads the charts of a sheet's drawing; undefined when it has none. */
export function readCharts(files: Record<string, Uint8Array>, sheet: Sheet, sheetPart: string, sheetRels: Rels): DrawingRaw | undefined {
  const rel = sheetRels.list.find((r) => isRel(r.type, DRAWING_REL_TYPE));
  if (!rel) return undefined;
  const part = resolveTarget(sheetPart, rel.target);
  if (!files[part]) return undefined;
  const xml = strFromU8(files[part]);
  const rels = new Rels(files[relsPath(part)] ? strFromU8(files[relsPath(part)]) : null);
  const r = root(xml);
  const charts: ChartModel[] = [];
  const read: ChartRead[] = [];
  for (const a of children(xml, r.open, r.close)) {
    const ax = outer(xml, a);
    if (!/<(?:c|cx):chart\b/.test(ax)) continue;
    const info = anchorFromXml(ax);
    const cr = info && rels.get(info.rId);
    const cpart = cr ? resolveTarget(part, cr.target) : "";
    if (!info || !cr || !files[cpart]) continue;
    try {
      const m = chartFromXml(strFromU8(files[cpart]), { from: info.from, to: info.to, name: info.name });
      charts.push(m);
      read.push({ xml: String(rawXml(m) ?? ""), part: cpart, rId: info.rId, anchor: ax, place: placeOf(m) });
    } catch {
      // A chart the model cannot read stays in the drawing as it is (it is not a model chart).
    }
  }
  if (!charts.length) return undefined;
  setCharts(sheet, charts);
  return { part, xml, charts: read, snap: canon(charts) };
}

/**
 * Writes the sheet's charts when they changed: kept charts reuse their parts and anchors,
 * new ones get parts of their own, removed ones lose their anchor (their parts go unused).
 * Returns the <drawing> element for the sheet, null to remove it, undefined to keep it.
 */
export function writeCharts(pw: PackageWriter, sheet: Sheet, sheetPart: string, raw: DrawingRaw | undefined, current: string | undefined, picsSnap?: string): string | null | undefined {
  const charts = chartsOf(sheet) ?? [];
  const pics = sheet.pictures ?? [];
  const picsSame = picturesKey(sheet) === (picsSnap ?? canon([]));
  if ((raw ? canon(charts) === raw.snap : !charts.length) && picsSame) return undefined;
  const sheetRels = pw.relsOf(sheetPart);
  let drel = sheetRels.list.find((r) => isRel(r.type, DRAWING_REL_TYPE));
  let part = drel ? resolveTarget(sheetPart, drel.target) : undefined;
  if (part && !pw.files[part]) part = undefined;
  const old = part ? strFromU8(pw.files[part]) : undefined;
  // The drawing's anchors in their order: pictures and shapes stay, a chart's slot is filled by
  // its model (or dropped); new charts come last.
  const slots: { xml: string; chart?: ChartRead }[] = [];
  let head = "", tail = "";
  if (old) {
    const r = root(old);
    const kids = children(old, r.open, r.close);
    head = old.slice(0, kids.length ? kids[0].start : r.close);
    tail = old.slice(kids.length ? kids[kids.length - 1].end : r.close);
    for (const a of kids) {
      const ax = outer(old, a);
      // Pictures: kept as they are, with new corners when moved, gone when deleted.
      if (/<xdr:pic\b/.test(ax)) {
        const p = pics.find((q) => q.raw === ax);
        if (!p) continue;
        const moved = canon([p.from, p.to ?? null]) !== canon([point(ax, "from") ?? null, p.to ? point(ax, "to") ?? null : null]);
        slots.push({ xml: moved ? movedPicture(p) : ax });
        continue;
      }
      slots.push({ xml: ax, chart: /<(?:c|cx):chart\b/.test(ax) ? raw?.charts.find((c) => c.anchor === ax) : undefined });
    }
  }
  const fresh = pics.filter((p) => !p.raw || !old?.includes(p.raw));
  const others = [...slots.filter((x) => !x.chart), ...fresh];
  if (!charts.length && !others.length) {
    if (drel) sheetRels.list.splice(sheetRels.list.indexOf(drel), 1), (sheetRels.changed = true);
    return null;
  }
  if (!part) part = pw.freeName("xl/drawings/drawing");
  const drels = pw.relsOf(part);
  let id = Math.max(1, ...slots.map((k) => Number(/<xdr:cNvPr\b[^>]*\bid="(\d+)"/.exec(k.xml)?.[1] ?? 1))) + 1;
  const placed = new Map<ChartRead, string>();
  const added: string[] = [];
  for (const m of charts) {
    const orig = raw?.charts.find((c) => !placed.has(c) && c.xml !== "" && c.xml === rawXml(m) && pw.files[c.part]);
    if (orig) {
      const xml = chartToXml(m);
      if (xml !== strFromU8(pw.files[orig.part])) pw.setPart(orig.part, xml);
      placed.set(orig, placeOf(m) === orig.place ? orig.anchor
        : anchorXml({ from: m.from, to: m.to, id: id++, name: m.name ?? `Diagramm ${id - 1}`, rId: orig.rId, chartEx: isChartEx(m) }));
      continue;
    }
    // A new chart (or a copy): a part of its own. Kept XML that refers to relationships of
    // another part is not reused.
    const ex = isChartEx(m);
    let xml = chartToXml(m);
    if (/\br:id="/.test(xml)) xml = chartToXml({ ...m, raw: undefined });
    const cpart = pw.freeName(ex ? "xl/charts/chartEx" : "xl/charts/chart");
    pw.setPart(cpart, xml, ex ? CHARTEX_CONTENT_TYPE : CHART_CONTENT_TYPE);
    const rId = pw.link(part, ex ? CHARTEX_REL_TYPE : CHART_REL_TYPE, cpart);
    added.push(anchorXml({ from: m.from, to: m.to, id: id++, name: m.name ?? `Diagramm ${id - 1}`, rId, chartEx: ex }));
  }
  // New pictures: their image as a media part of the package.
  for (const p of fresh) {
    const m = /^data:([^;]+);base64,(.*)$/.exec(p.src);
    if (!m) continue;
    const bin = atob(m[2]);
    const bytes = new Uint8Array(bin.length);
    for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
    const media = pw.addMedia(bytes, m[1], "xl/media/image");
    const rId = pw.link(part, IMAGE_REL_TYPE, media);
    added.push(newPicture(p, id++, rId));
  }
  const all = [...slots.map((x) => (x.chart ? placed.get(x.chart) ?? "" : x.xml)), ...added].join("");
  // Relationships of removed charts.
  for (const r of [...drels.list]) {
    if ((isRel(r.type, CHART_REL_TYPE) || isRel(r.type, CHARTEX_REL_TYPE)) && !all.includes(`r:id="${r.id}"`)) {
      drels.list.splice(drels.list.indexOf(r), 1);
      drels.changed = true;
    }
  }
  pw.setPart(part, old ? head + all + tail : drawingXml([all]), old ? undefined : DRAWING_CONTENT_TYPE);
  if (!drel) {
    pw.link(sheetPart, DRAWING_REL_TYPE, part);
    drel = sheetRels.list.find((r) => isRel(r.type, DRAWING_REL_TYPE))!;
  }
  return current && current.includes(`r:id="${drel.id}"`) ? current : `<drawing r:id="${drel.id}"/>`;
}

const EMU_PER_PX = 9525;
const MIME: Record<string, string> = { png: "image/png", jpg: "image/jpeg", jpeg: "image/jpeg", gif: "image/gif", bmp: "image/bmp", svg: "image/svg+xml", emf: "image/emf", wmf: "image/wmf" };

/**
 * The pictures of a sheet's drawing (xdr:pic in two- or one-cell anchors), with their image as a
 * data: URL, for showing them; the drawing itself is written back as it was.
 */
function point(ax: string, tag: string): { col: number; row: number; x: number; y: number } | undefined {
  const m = new RegExp(`<xdr:${tag}>([\\s\\S]*?)</xdr:${tag}>`).exec(ax);
  if (!m) return undefined;
  const v = (n: string) => Number(new RegExp(`<xdr:${n}>([^<]*)</xdr:${n}>`).exec(m[1])?.[1] ?? 0);
  return { col: v("col"), row: v("row"), x: v("colOff") / EMU_PER_PX, y: v("rowOff") / EMU_PER_PX };
}

/** What decides whether the drawing's pictures changed (their places, names and identity, not the image data). */
export function picturesKey(sheet: Sheet): string {
  return canon((sheet.pictures ?? []).map((p) => ({ from: p.from, to: p.to ?? null, size: p.size ?? null, name: p.name ?? null, raw: p.raw ?? null })));
}

const IMAGE_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";

const pointXml = (tag: string, p: { col: number; row: number; x: number; y: number }) =>
  `<xdr:${tag}><xdr:col>${p.col}</xdr:col><xdr:colOff>${Math.round(p.x * EMU_PER_PX)}</xdr:colOff><xdr:row>${p.row}</xdr:row><xdr:rowOff>${Math.round(p.y * EMU_PER_PX)}</xdr:rowOff></xdr:${tag}>`;

/** A picture's anchor after it moved: the kept XML with new corners (and size). */
function movedPicture(p: NonNullable<Sheet["pictures"]>[number]): string {
  let x = p.raw!.replace(/<xdr:from>[\s\S]*?<\/xdr:from>/, pointXml("from", p.from));
  if (p.to) x = x.replace(/<xdr:to>[\s\S]*?<\/xdr:to>/, pointXml("to", p.to));
  if (p.size) x = x.replace(/<xdr:ext\b[^>]*\/>/, `<xdr:ext cx="${Math.round(p.size.w * EMU_PER_PX)}" cy="${Math.round(p.size.h * EMU_PER_PX)}"/>`);
  return x;
}

/** A new picture's anchor (Excel's form for an inserted picture: two cells, sized with them). */
function newPicture(p: NonNullable<Sheet["pictures"]>[number], id: number, rId: string): string {
  const w = Math.round((p.size?.w ?? 100) * EMU_PER_PX), h = Math.round((p.size?.h ?? 100) * EMU_PER_PX);
  const name = p.name ?? `Grafik ${id - 1}`;
  return `<xdr:twoCellAnchor editAs="oneCell">${pointXml("from", p.from)}${pointXml("to", p.to ?? p.from)}<xdr:pic><xdr:nvPicPr><xdr:cNvPr id="${id}" name="${name.replace(/[<&"]/g, "")}"/><xdr:cNvPicPr><a:picLocks noChangeAspect="1"/></xdr:cNvPicPr></xdr:nvPicPr>` +
    `<xdr:blipFill><a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="${rId}"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill>` +
    `<xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${w}" cy="${h}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic><xdr:clientData/></xdr:twoCellAnchor>`;
}

export function readPictures(files: Record<string, Uint8Array>, sheet: Sheet, sheetPart: string, sheetRels: Rels): string {
  const rel = sheetRels.list.find((r) => isRel(r.type, DRAWING_REL_TYPE));
  if (!rel) return picturesKey(sheet);
  const part = resolveTarget(sheetPart, rel.target);
  if (!files[part]) return picturesKey(sheet);
  const xml = strFromU8(files[part]);
  const rels = new Rels(files[relsPath(part)] ? strFromU8(files[relsPath(part)]) : null);
  const r = root(xml);
  const out: NonNullable<Sheet["pictures"]> = [];
  for (const a of children(xml, r.open, r.close)) {
    const ax = outer(xml, a);
    if (!/<xdr:pic\b/.test(ax)) continue;
    const from = point(ax, "from");
    const embed = /<a:blip\b[^>]*r:embed="([^"]*)"/.exec(ax)?.[1];
    const target = embed ? rels.get(embed) : undefined;
    if (!from || !target) continue;
    const media = resolveTarget(part, target.target);
    const bytes = files[media];
    const ext = media.split(".").pop()?.toLowerCase() ?? "";
    if (!bytes || !MIME[ext] || ext === "emf" || ext === "wmf") continue;
    let bin = "";
    for (let i = 0; i < bytes.length; i += 0x8000) bin += String.fromCharCode(...bytes.subarray(i, i + 0x8000));
    const src = `data:${MIME[ext]};base64,${btoa(bin)}`;
    const to = point(ax, "to");
    const extEl = /<xdr:ext\b[^>]*cx="(\d+)"[^>]*cy="(\d+)"/.exec(ax);
    const name = /<xdr:cNvPr\b[^>]*name="([^"]*)"/.exec(ax)?.[1];
    out.push({ from, to: a.name === "xdr:twoCellAnchor" ? to : undefined, size: a.name !== "xdr:twoCellAnchor" && extEl ? { w: Number(extEl[1]) / EMU_PER_PX, h: Number(extEl[2]) / EMU_PER_PX } : undefined, src, name, raw: ax });
  }
  if (out.length) sheet.pictures = out;
  return picturesKey(sheet);
}
