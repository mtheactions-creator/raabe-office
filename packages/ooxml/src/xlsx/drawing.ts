// Charts of a worksheet: read from its drawing part (xl/drawings/drawingN.xml) into the
// sheet's chart models and written back. The drawing and every chart part stay as Excel wrote
// them while the charts are unchanged; pictures and shapes in the drawing are always kept.

import { shapeFromXml, shapeKey, shapeToXml } from "./shapes";
import { bakedKey, lookKey, pictureAnchorXml, readPictureLook, rootDecls } from "./pictures";
import type { PictureFormat } from "@raabe/text-engine";
import { strFromU8 } from "fflate";
import type { ChartModel } from "@raabe/charts";
import type { Sheet } from "@raabe/sheet-engine";
import { PackageWriter, Rels, relsPath, resolveTarget } from "../package";
import {
  CHART_COLORS_CONTENT_TYPE, CHART_COLORS_REL_TYPE, CHART_CONTENT_TYPE, CHART_REL_TYPE, CHART_STYLE_CONTENT_TYPE, CHART_STYLE_REL_TYPE,
  CHARTEX_CONTENT_TYPE, CHARTEX_REL_TYPE, DRAWING_CONTENT_TYPE, DRAWING_REL_TYPE, anchorFromXml, chartColorsXml, chartStyleXml, anchorXml, chartFromXml, chartToXml, chartToXmlWithImages, drawingXml, isChartEx,
} from "./chart";
import { isRel } from "./read";
import { canon } from "./styles";
import { children, outer, root } from "./xml";
import { inkKey, isInkAnchor, readInk, writeInk } from "./ink";

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
const placeOf = (m: ChartModel) => canon([m.from, m.to, m.name ?? null, m.props ?? null]);

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
      // Bild- oder Texturfüllungen: the chart part's pictures as data URLs.
      const crels = new Rels(files[relsPath(cpart)] ? strFromU8(files[relsPath(cpart)]) : null);
      const image = (rId: string) => {
        const ir = crels.get(rId);
        const ipart = ir ? resolveTarget(cpart, ir.target) : "";
        const bytes = ipart ? files[ipart] : undefined;
        if (!bytes) return undefined;
        let bin = "";
        for (let i = 0; i < bytes.length; i += 0x8000) bin += String.fromCharCode(...bytes.subarray(i, i + 0x8000));
        const ext = (/\.(\w+)$/.exec(ipart)?.[1] ?? "png").toLowerCase();
        return `data:${MIME[ext] ?? "image/png"};base64,${btoa(bin)}`;
      };
      const m = chartFromXml(strFromU8(files[cpart]), { from: info.from, to: info.to, name: info.name, image });
      if (info.props) m.props = info.props;
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
  // The anchors the reader turned into pictures; others (EMF logos …) are kept as they are.
  const readRaws = new Set<string>((JSON.parse(picsSnap ?? "[]") as { raw?: string | null }[]).map((p) => p.raw ?? "").filter(Boolean));
  if ((raw ? canon(charts) === raw.snap : !charts.length) && picsSame) return undefined;
  const sheetRels = pw.relsOf(sheetPart);
  let drel = sheetRels.list.find((r) => isRel(r.type, DRAWING_REL_TYPE));
  let part = drel ? resolveTarget(sheetPart, drel.target) : undefined;
  if (part && !pw.files[part]) part = undefined;
  const old = part ? strFromU8(pw.files[part]) : undefined;
  // The drawing's anchors in their order: pictures and shapes stay, a chart's slot is filled by
  // its model (or dropped); new charts come last.
  const slots: { xml: string; chart?: ChartRead; ink?: NonNullable<Sheet["inks"]>[number] }[] = [];
  const inkRaws = new Set((JSON.parse(picsSnap ?? "[]") as { ink?: string; raw?: string | null }[]).filter((p) => p.ink).map((p) => p.raw ?? ""));
  let head = "", tail = "";
  const keptPics = pics.filter((p) => p.raw && !p.group && old?.includes(p.raw));
  let nextPic = 0;
  // A picture's image: the media part it has (same picture) or a new one.
  const imageRid = (src: string, p: NonNullable<Sheet["pictures"]>[number] | null, at: string): string | undefined => {
    if (p?.media && p.media.src === src && p.media.rid) return p.media.rid;
    const m = /^data:([^;]+);base64,(.*)$/.exec(src);
    if (!m) return undefined;
    const bin = atob(m[2]);
    const bytes = new Uint8Array(bin.length);
    for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
    const media = pw.addMedia(bytes, m[1], "xl/media/image");
    return pw.link(at, IMAGE_REL_TYPE, media);
  };
  /** A picture with its Bildformat as Excel writes it; `fresh`: one not read from this drawing. */
  const formatted = (p: NonNullable<Sheet["pictures"]>[number], id: number, fresh: boolean, at: string): string | undefined => {
    const f = p.format as PictureFormat | undefined;
    const fx = bakedKey(f);
    let image = p.src, keepFx = false;
    if (fx && p.baked === fx) { image = p.bakedSrc ?? p.src; keepFx = !fresh && !p.bakedSrc && !!p.media && p.media.src === p.src; }
    const rid = imageRid(image, fresh ? null : p, at);
    if (!rid) return undefined;
    const ext = p.raw ? /<a:xfrm\b[^>]*>\s*<a:off\b[^>]*\/>\s*<a:ext cx="(\d+)" cy="(\d+)"/.exec(p.raw) : null;
    const size = p.size ?? (ext ? { w: Number(ext[1]) / EMU_PER_PX, h: Number(ext[2]) / EMU_PER_PX } : { w: 100, h: 100 });
    return pictureAnchorXml(fresh ? { ...p, lockAspect: p.lockAspect ?? true } : p, id, rid, size, { keepFx });
  };
  const pictureXml = (p: NonNullable<Sheet["pictures"]>[number]): string => {
    const ax = p.raw!;
    if (lookKey(p) !== p.lookKey && part) {
      const sid = Number(/<xdr:cNvPr\b[^>]*\bid="(\d+)"/.exec(ax)?.[1] ?? 2);
      const xml = formatted(p, sid, false, part);
      if (xml) return xml;
    }
    const moved = canon([p.from, p.to ?? null]) !== canon([point(ax, "from") ?? null, p.to ? point(ax, "to") ?? null : null]);
    return moved ? movedPicture(p) : ax;
  };
  if (old) {
    const r = root(old);
    const kids = children(old, r.open, r.close);
    head = old.slice(0, kids.length ? kids[0].start : r.close);
    tail = old.slice(kids.length ? kids[kids.length - 1].end : r.close);
    for (const a of kids) {
      const ax = outer(old, a);
      // Ink (D-093): the model's version, gone when deleted (written once the drawing's part is known).
      if (isInkAnchor(ax)) {
        const k = (sheet.inks ?? []).find((q) => q.raw === ax);
        if (k) slots.push({ xml: "", ink: k });
        else if (!inkRaws.has(ax)) slots.push({ xml: ax });
        continue;
      }
      // Pictures: kept as they are, with new corners when moved, with Excel's DrawingML when their
      // Bildformat changed, gone when deleted; the model's order fills the pictures' places (Ebene
      // nach vorne / nach hinten).
      // A group holding pictures: kept as it is (its pictures are shown, not changed on their own).
      if (/<xdr:grpSp\b/.test(ax) && pics.some((q) => q.raw === ax && q.group)) { slots.push({ xml: ax }); continue; }
      if (/<xdr:pic\b/.test(ax) && !/<xdr:grpSp\b/.test(ax)) {
        if (!pics.some((q) => q.raw === ax)) { if (!readRaws.has(ax)) slots.push({ xml: ax }); continue; }
        const p = keptPics[nextPic++];
        if (p) slots.push({ xml: pictureXml(p) });
        continue;
      }
      // Shapes: the model's version (the same XML while unchanged), gone when deleted.
      if (/^<xdr:twoCellAnchor\b/.test(ax) && /<xdr:(?:sp|cxnSp|grpSp)\b/.test(ax)) {
        const s = (sheet.shapes ?? []).find((q) => q.raw === ax);
        if (!s) { if (!readRaws.has(ax)) slots.push({ xml: ax }); continue; }
        const sid = Number(/<xdr:cNvPr\b[^>]*\bid="(\d+)"/.exec(ax)?.[1] ?? 2);
        slots.push({ xml: shapeToXml(s, sid, s.size ?? { w: 120, h: 80 }) });
        continue;
      }
      slots.push({ xml: ax, chart: /<(?:c|cx):chart\b/.test(ax) ? raw?.charts.find((c) => c.anchor === ax) : undefined });
    }
  }
  /** A chart part's XML with its picture fills linked as media of that part (D-111). */
  const chartPart = (m: ChartModel, cpart: string): string => {
    const { xml, images } = chartToXmlWithImages(m);
    let out = xml;
    images.forEach((src, k) => {
      const mm = /^data:([^;]+);base64,(.*)$/.exec(src);
      if (!mm) return;
      const bin = atob(mm[2]);
      const bytes = new Uint8Array(bin.length);
      for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
      const media = pw.addMedia(bytes, mm[1], "xl/media/image");
      out = out.split(`r:embed="@img${k}"`).join(`r:embed="${pw.link(cpart, IMAGE_REL_TYPE, media)}"`);
    });
    if (!isChartEx(m)) chartStyleParts(m, cpart);
    return out;
  };
  // The chart's style and colour parts (xl/charts/styleN.xml, colorsN.xml): Excel reports the chosen
  // Diagrammformatvorlage (Chart.ChartStyle) from the style part's id (D-155). A part Excel wrote is kept and only
  // its id follows the chart's style; a chart without one gets ours.
  const chartStyleParts = (m: ChartModel, cpart: string): void => {
    const rels = pw.relsOf(cpart);
    const id = m.style ?? 201;
    const sr = rels.list.find((r) => isRel(r.type, CHART_STYLE_REL_TYPE));
    const spath = sr ? resolveTarget(cpart, sr.target) : undefined;
    if (spath && pw.files[spath]) {
      const old = strFromU8(pw.files[spath]);
      const next = old.replace(/(<cs:chartStyle\b[^>]*?\bid=")\d+(")/, `$1${id}$2`);
      if (next !== old) pw.setPart(spath, next);
      return;
    }
    const style = pw.freeName("xl/charts/style");
    pw.setPart(style, chartStyleXml(id), CHART_STYLE_CONTENT_TYPE);
    pw.link(cpart, CHART_STYLE_REL_TYPE, style);
    if (!rels.list.some((r) => isRel(r.type, CHART_COLORS_REL_TYPE))) {
      const colors = pw.freeName("xl/charts/colors");
      pw.setPart(colors, chartColorsXml(), CHART_COLORS_CONTENT_TYPE);
      pw.link(cpart, CHART_COLORS_REL_TYPE, colors);
    }
  };
  const fresh = pics.filter((p) => !p.raw || !old?.includes(p.raw));
  const freshShapes = (sheet.shapes ?? []).filter((s) => !s.raw || !old?.includes(s.raw));
  const freshInks = (sheet.inks ?? []).filter((k) => !k.raw || !old?.includes(k.raw));
  const others = [...slots.filter((x) => !x.chart), ...fresh, ...freshShapes, ...freshInks];
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
      const xml = chartPart(m, orig.part);
      if (xml !== strFromU8(pw.files[orig.part])) pw.setPart(orig.part, xml);
      placed.set(orig, placeOf(m) === orig.place ? orig.anchor
        : anchorXml({ from: m.from, to: m.to, id: id++, name: m.name ?? `Diagramm ${id - 1}`, rId: orig.rId, chartEx: isChartEx(m), props: m.props }));
      continue;
    }
    // A new chart (or a copy): a part of its own. Kept XML that refers to relationships of
    // another part is not reused.
    const ex = isChartEx(m);
    const cpart = pw.freeName(ex ? "xl/charts/chartEx" : "xl/charts/chart");
    // Kept XML that refers to relationships of another part is not reused.
    const xml = chartPart(/\br:(id|embed)="/.test(chartToXml(m)) ? { ...m, raw: undefined } : m, cpart);
    pw.setPart(cpart, xml, ex ? CHARTEX_CONTENT_TYPE : CHART_CONTENT_TYPE);
    const rId = pw.link(part, ex ? CHARTEX_REL_TYPE : CHART_REL_TYPE, cpart);
    added.push(anchorXml({ from: m.from, to: m.to, id: id++, name: m.name ?? `Diagramm ${id - 1}`, rId, chartEx: ex, props: m.props }));
  }
  // New pictures: their image as a media part of the package.
  for (const p of fresh) {
    const xml = formatted(p, id, true, part);
    if (!xml) continue;
    id++;
    added.push(xml);
  }
  // New shapes and text boxes.
  for (const s of freshShapes) added.push(shapeToXml({ ...s, raw: undefined }, id++, s.size ?? { w: 120, h: 80 }));
  // Ink: kept anchors stay, changed and new ones get their InkML part (and picture) as Excel writes them.
  for (const x of slots) if (x.ink) x.xml = writeInk(pw, part, x.ink, Number(/<xdr14:cNvPr\b[^>]*\bid="(\d+)"/.exec(x.ink.raw ?? "")?.[1] ?? id++), { x: 0, y: 0 });
  for (const k of freshInks) added.push(writeInk(pw, part, { ...k, raw: undefined }, id++, { x: 0, y: 0 }));
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
  return canon([
    ...(sheet.pictures ?? []).map((p) => ({ from: p.from, to: p.to ?? null, size: p.size ?? null, name: p.name ?? null, raw: p.raw ?? null, look: lookKey(p) })),
    ...(sheet.shapes ?? []).map((s) => ({ shape: shapeKey(s), raw: s.raw ?? null })),
    ...(sheet.inks ?? []).map((k) => ({ ink: inkKey(k), raw: k.raw ?? null })),
  ]);
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

export function readPictures(files: Record<string, Uint8Array>, sheet: Sheet, sheetPart: string, sheetRels: Rels, theme: Record<string, string> = {}): string {
  const rel = sheetRels.list.find((r) => isRel(r.type, DRAWING_REL_TYPE));
  if (!rel) return picturesKey(sheet);
  const part = resolveTarget(sheetPart, rel.target);
  if (!files[part]) return picturesKey(sheet);
  const xml = strFromU8(files[part]);
  const rels = new Rels(files[relsPath(part)] ? strFromU8(files[relsPath(part)]) : null);
  const r = root(xml);
  const decls = rootDecls(xml);
  const out: NonNullable<Sheet["pictures"]> = [];
  const shapes: NonNullable<Sheet["shapes"]> = [];
  const inks: NonNullable<Sheet["inks"]> = [];
  for (const a of children(xml, r.open, r.close)) {
    const ax = outer(xml, a);
    // Ink (D-093): read before pictures, as its Fallback holds a picture.
    if (isInkAnchor(ax)) {
      const k = readInk(ax, part, rels, files);
      if (k) inks.push(k);
      continue;
    }
    const sh = shapeFromXml(ax);
    if (sh) {
      const ext = /<a:ext cx="(\d+)" cy="(\d+)"/.exec(ax);
      if (ext) { sh.size = { w: Number(ext[1]) / EMU_PER_PX, h: Number(ext[2]) / EMU_PER_PX }; sh.rawKey = shapeKey(sh); }
      shapes.push(sh);
      continue;
    }
    if (!/<xdr:pic\b/.test(ax)) continue;
    const from = point(ax, "from");
    if (!from) continue;
    const to = point(ax, "to");
    const extEl = /<xdr:ext\b[^>]*cx="(\d+)"[^>]*cy="(\d+)"/.exec(ax);
    const size = a.name !== "xdr:twoCellAnchor" && extEl ? { w: Number(extEl[1]) / EMU_PER_PX, h: Number(extEl[2]) / EMU_PER_PX } : undefined;
    // A group holding pictures: each picture is shown in its place inside the group's box (as part
    // of the group: not selected on its own); the group's anchor stays as the file has it.
    const group = /<xdr:grpSp\b/.test(ax);
    const units = group ? [...ax.matchAll(/<xdr:pic\b[\s\S]*?<\/xdr:pic>/g)].map((m) => m[0]) : [ax];
    const g = group ? /<xdr:grpSpPr\b[\s\S]*?<a:chOff x="(-?\d+)" y="(-?\d+)"\/>\s*<a:chExt cx="(\d+)" cy="(\d+)"\/>/.exec(ax) : null;
    for (const unit of units) {
      const embed = /<a:blip\b[^>]*r:embed="([^"]*)"/.exec(unit)?.[1];
      const target = embed ? rels.get(embed) : undefined;
      if (!target) continue;
      const media = resolveTarget(part, target.target);
      const bytes = files[media];
      const ext = media.split(".").pop()?.toLowerCase() ?? "";
      if (!bytes || !MIME[ext] || ext === "emf" || ext === "wmf") continue;
      let bin = "";
      for (let i = 0; i < bytes.length; i += 0x8000) bin += String.fromCharCode(...bytes.subarray(i, i + 0x8000));
      const src = `data:${MIME[ext]};base64,${btoa(bin)}`;
      const name = /<xdr:cNvPr\b[^>]*name="([^"]*)"/.exec(unit)?.[1];
      const pic: NonNullable<Sheet["pictures"]>[number] = { from, to: a.name === "xdr:twoCellAnchor" ? to : undefined, size, src, name, raw: ax };
      if (group) {
        const off = /<a:xfrm\b[^>]*>\s*<a:off x="(-?\d+)" y="(-?\d+)"\/>\s*<a:ext cx="(\d+)" cy="(\d+)"\/>/.exec(unit);
        if (!g || !off) continue;
        const [cx, cy, cw, ch] = g.slice(1).map(Number);
        pic.group = { x: (Number(off[1]) - cx) / (cw || 1), y: (Number(off[2]) - cy) / (ch || 1), w: Number(off[3]) / (cw || 1), h: Number(off[4]) / (ch || 1) };
      }
      // Bildformat (D-087): the picture's DrawingML as text-engine's format.
      const { embed: _e, ...look } = readPictureLook(group ? `<xdr:oneCellAnchor>${unit}</xdr:oneCellAnchor>` : ax, decls, theme);
      void _e;
      Object.assign(pic, look);
      if (pic.media) pic.media.src = src;
      if (pic.name) pic.name = pic.name.replace(/&quot;/g, '"').replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&amp;/g, "&");
      pic.lookKey = lookKey(pic);
      out.push(pic);
    }
  }
  if (out.length) sheet.pictures = out;
  if (shapes.length) sheet.shapes = shapes;
  if (inks.length) sheet.inks = inks;
  return picturesKey(sheet);
}
