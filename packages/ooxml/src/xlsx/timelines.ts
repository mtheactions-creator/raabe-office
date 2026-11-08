// Zeitachsen (timelines) in .xlsx (D-121), as German Excel 365 writes them (recorded with probe_excel_timeline.py,
// fixture excel-timeline.xlsx):
// - xl/timelineCaches/timelineCacheN.xml per cache (relationship .../2011/relationships/timelineCache from the
//   workbook, listed in the workbook's extLst as x15:timelineCacheRefs): the PivotTables it filters, and a state
//   with the PivotTable cache's x14 pivotCacheId, filterType (dateBetween or unknown), the selection and the bounds;
// - xl/timelines/timelineN.xml per sheet (relationship .../2011/relationships/timeline, x15:timelineRefs in the
//   sheet's extLst): each <timeline> names its cache, the level, the selection's level, the scroll position and the
//   Anzeigen options that are off;
// - a two-cell anchor (editAs oneCell) in the sheet's drawing with an mc:AlternateContent graphicFrame naming the
//   timeline (tsle:timeslicer), a plain shape as fallback;
// - a workbook name per cache ("NativeZeitachse_Datum", #NV), kept in the model's names;
// - the filter itself is the PivotTables' dateBetween filter (pivots.ts).
// While nothing about the timelines (and their PivotTables) changed every part stays as read; after a change all
// timeline parts are written anew from the model.

import { strFromU8 } from "fflate";
import {
  dateToSerial, serialToDate, timelineBounds, timelineKey,
  type PivotTable, type Sheet, type TimeLevel, type Timeline, type TimelineCache, type Workbook,
} from "@raabe/sheet-engine";
import { PackageWriter, Rels, relsPath, resolveTarget } from "../package";
import { isRel, type BookRaw, type SheetRaw } from "./read";
import { canon } from "./styles";
import { esc, unesc } from "./xml";

export const RT_TIMELINE = "http://schemas.microsoft.com/office/2011/relationships/timeline";
export const RT_TIMELINE_CACHE = "http://schemas.microsoft.com/office/2011/relationships/timelineCache";
export const CT_TIMELINE = "application/vnd.ms-excel.timeline+xml";
export const CT_TIMELINE_CACHE = "application/vnd.ms-excel.timelineCache+xml";
const RT_DRAWING = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing";
const CT_DRAWING = "application/vnd.openxmlformats-officedocument.drawing+xml";
const HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n';
const X15 = "http://schemas.microsoft.com/office/spreadsheetml/2010/11/main";
const X14 = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main";
const NS_TAIL = `xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x xr10" xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:xr10="http://schemas.microsoft.com/office/spreadsheetml/2016/revision10"`;
export const TL_URI = { book: "{D0CA8CA8-9F24-4464-BF8E-62219DCF47F9}", sheet: "{7E03D99C-DC04-49d9-9315-930204A7B6E9}", pivotCacheId: "{725AE2AE-9491-48be-B2B4-4EB974FC3084}" };
const EMU_PER_PX = 9525;

const attr = (tag: string, name: string) => { const v = new RegExp(`\\b${name}="([^"]*)"`).exec(tag)?.[1]; return v === undefined ? undefined : unesc(v); };
const str = (files: Record<string, Uint8Array>, p: string | undefined) => (p && files[p] ? strFromU8(files[p]) : "");

/** "2025-10-01T00:00:00" to a serial and back. */
export function isoToSerial(iso: string): number {
  const m = /^(\d{4})-(\d{2})-(\d{2})/.exec(iso);
  return m ? dateToSerial(Number(m[1]), Number(m[2]), Number(m[3])) : NaN;
}
export function serialToIso(serial: number): string {
  const d = serialToDate(Math.floor(serial));
  return `${String(d.y).padStart(4, "0")}-${String(d.m).padStart(2, "0")}-${String(d.d).padStart(2, "0")}T00:00:00`;
}

/** What decides whether the timelines changed. */
export function timelineSnapshot(wb: Workbook): string {
  return canon({ caches: wb.timelineCaches ?? [], sheets: wb.sheets.map((s) => (s.timelines ?? []).map((x) => JSON.parse(timelineKey(x)))) });
}

export function timelineAnchorName(ax: string): string | undefined {
  const m = /<tsle:timeslicer\b[^>]*\bname="([^"]*)"/.exec(ax);
  return m ? unesc(m[1]) : undefined;
}

function point(ax: string, tag: string): { col: number; row: number; x: number; y: number } | undefined {
  const m = new RegExp(`<xdr:${tag}>([\\s\\S]*?)</xdr:${tag}>`).exec(ax);
  if (!m) return undefined;
  const v = (n: string) => Number(new RegExp(`<xdr:${n}>([^<]*)</xdr:${n}>`).exec(m[1])?.[1] ?? 0);
  return { col: v("col"), row: v("row"), x: v("colOff") / EMU_PER_PX, y: v("rowOff") / EMU_PER_PX };
}

function anchors(xml: string): string[] {
  return [...xml.matchAll(/<xdr:(twoCellAnchor|oneCellAnchor|absoluteAnchor)\b[\s\S]*?<\/xdr:\1>/g)].map((m) => m[0]);
}

interface TimelineSnap { snap: string }

/** The workbook's timeline caches and every sheet's timelines with their drawing anchors. */
export function readTimelines(wb: Workbook): void {
  const raw = wb.raw as BookRaw | undefined;
  if (!raw) return;
  const files = raw.files;
  const wbRels = new Rels(str(files, relsPath(raw.wbPath)) || null);
  const sheetById = new Map<number, Sheet>();
  for (const s of wb.sheets) { const r = s.raw as SheetRaw | undefined; if (r) sheetById.set(r.sheetId, s); }
  const caches: TimelineCache[] = [];
  for (const r of wbRels.list.filter((x) => isRel(x.type, RT_TIMELINE_CACHE))) {
    const xml = str(files, resolveTarget(raw.wbPath, r.target));
    const top = /<timelineCacheDefinition\b[^>]*>/.exec(xml)?.[0];
    if (!top) continue;
    const c: TimelineCache = {
      name: attr(top, "name") ?? "", sourceName: attr(top, "sourceName") ?? "",
      pivots: [...xml.matchAll(/<pivotTable\b([^>]*)\/>/g)].map((m) => ({ sheet: sheetById.get(Number(attr(m[1], "tabId")))?.name ?? "", name: attr(m[1], "name") ?? "" })),
    };
    const id = Number(attr(/<state\b[^>]*>/.exec(xml)?.[0] ?? "", "pivotCacheId") ?? 0);
    if (id) c.pivotCacheId = id;
    caches.push(c);
  }
  if (caches.length) wb.timelineCaches = caches;
  for (const s of wb.sheets) {
    const sr = s.raw as SheetRaw | undefined;
    if (!sr) continue;
    const rels = new Rels(str(files, relsPath(sr.part)) || null);
    const list: Timeline[] = [];
    for (const r of rels.list.filter((x) => isRel(x.type, RT_TIMELINE))) {
      const xml = str(files, resolveTarget(sr.part, r.target));
      for (const m of xml.matchAll(/<timeline\b([^>]*?)\/?>/g)) {
        const a = m[1];
        const lvl = (v: string | undefined, d: TimeLevel): TimeLevel => (v !== undefined && /^[0-3]$/.test(v) ? (Number(v) as TimeLevel) : d);
        const t: Timeline = { name: attr(a, "name") ?? "", cache: attr(a, "cache") ?? "", caption: attr(a, "caption") ?? attr(a, "name") ?? "",
          level: lvl(attr(a, "level"), 2), from: { col: 0, row: 0, x: 0, y: 0 }, to: { col: 5, row: 7, x: 0, y: 0 } };
        const sl = attr(a, "selectionLevel");
        if (sl !== undefined) t.selectionLevel = lvl(sl, t.level);
        const sp = attr(a, "scrollPosition");
        if (sp) t.scrollPosition = isoToSerial(sp);
        for (const k of ["showHeader", "showSelectionLabel", "showTimeLevel", "showHorizontalScrollbar"] as const) if (attr(a, k) === "0") t[k] = false;
        const st = attr(a, "style");
        if (st) t.style = st;
        list.push(t);
      }
    }
    if (!list.length) continue;
    const drel = rels.list.find((x) => isRel(x.type, RT_DRAWING));
    const dxml = drel ? str(files, resolveTarget(sr.part, drel.target)) : "";
    for (const ax of anchors(dxml)) {
      const name = timelineAnchorName(ax);
      const t = list.find((x) => x.name === name);
      if (!t) continue;
      t.from = point(ax, "from") ?? t.from;
      t.to = point(ax, "to") ?? t.to;
      t.raw = ax;
    }
    for (const t of list) if (t.raw) t.rawKey = timelineKey(t);
    s.timelines = list;
  }
  (raw as BookRaw & { timelines?: TimelineSnap }).timelines = { snap: timelineSnapshot(wb) };
}

// ----- writing ------------------------------------------------------------------------------

export interface TimelinePlan {
  changed: boolean;
  /** For writePivots: the date fields (by name) a timeline filters, and the x14 pivotCacheId it needs. */
  dateFields(sheet: string, pivot: string): string[];
  pivotCacheId(sheet: string, pivot: string): number | undefined;
  sheetXml(sheet: Sheet, part: string, xml: string): string;
  finish(wbXml: string, sheetIds: Map<string, number>): string;
}

const rewritten = (p: { raw?: string; rawKey?: string }, key: (p: never) => string) => !p.raw || p.rawKey !== key(p as never);

/** Decides before the sheets are written whether the timeline parts are rewritten; `idOf` gives a slicer's id first. */
export function planTimelines(wb: Workbook, pw: PackageWriter, wbPath: string, pivotKey: (p: never) => string, idOf: (sheet: string, pivot: string) => number | undefined): TimelinePlan {
  const raw = wb.raw as (BookRaw & { timelines?: TimelineSnap }) | undefined;
  const caches = wb.timelineCaches ?? [];
  const pivotOf = (sheet: string, name: string) => wb.sheets.find((s) => s.name === sheet)?.pivots?.find((p) => p.name === name);
  const anyPivotRewritten = caches.some((c) => c.pivots.some((r) => { const p = pivotOf(r.sheet, r.name); return !!p && rewritten(p, pivotKey); }));
  const had = !!raw?.timelines;
  const changed = (caches.length > 0 || wb.sheets.some((s) => s.timelines?.length) || had)
    && (!had || raw!.timelines!.snap !== timelineSnapshot(wb) || anyPivotRewritten);
  let nextId = Math.max(2000, ...Object.keys(pw.files).filter((p) => p.includes("pivotCache")).map((p) => Number(/pivotCacheId="(\d+)"/.exec(str(pw.files, p))?.[1] ?? 0))) + 1;
  const ids = new Map<string, number>();
  const has = (sheet: string, pivot: string) => caches.some((c) => c.pivots.some((r) => r.sheet === sheet && r.name === pivot));
  const pivotCacheId = (sheet: string, pivot: string) => {
    if (!has(sheet, pivot)) return undefined;
    const k = `${sheet}\u0000${pivot}`;
    if (!ids.has(k)) ids.set(k, idOf(sheet, pivot) ?? nextId++);
    return ids.get(k);
  };
  const parts = new Map<Sheet, string>();
  return {
    changed,
    dateFields: (sheet, pivot) => caches.filter((c) => c.pivots.some((r) => r.sheet === sheet && r.name === pivot)).map((c) => c.sourceName),
    pivotCacheId: (sheet, pivot) => (changed ? pivotCacheId(sheet, pivot) : undefined),
    sheetXml: (sheet, part, xml) => { parts.set(sheet, part); return changed ? writeSheetTimelines(pw, sheet, part, xml) : xml; },
    finish: (wbXml, sheetIds) => (changed ? writeCaches(pw, wb, wbPath, wbXml, sheetIds, pivotCacheId, parts) : wbXml),
  };
}

function patchExt(xml: string, close: string, uri: string, add: string | null): string {
  const m = /<extLst>([\s\S]*?)<\/extLst>/.exec(xml);
  let inner = m ? m[1] : "";
  inner = inner.replace(new RegExp(`<(\\w+:)?ext\\b[^>]*uri="${uri.replace(/[{}]/g, "\\$&")}"[\\s\\S]*?<\\/\\1?ext>`), "");
  if (add) inner += add;
  if (m) return inner ? xml.replace(m[0], `<extLst>${inner}</extLst>`) : xml.replace(m[0], "");
  return inner ? xml.replace(close, `<extLst>${inner}</extLst>${close}`) : xml;
}

const guid = (seed: string) => {
  let h = 2166136261;
  const hex: string[] = [];
  for (let k = 0; k < 4; k++) {
    for (const ch of seed + k) h = Math.imul(h ^ ch.charCodeAt(0), 16777619) >>> 0;
    hex.push(h.toString(16).toUpperCase().padStart(8, "0"));
  }
  const x = hex.join("");
  return `{${x.slice(0, 8)}-${x.slice(8, 12)}-4${x.slice(13, 16)}-A${x.slice(17, 20)}-${x.slice(20, 32)}}`;
};

/** A timeline's drawing anchor as Excel writes it (editAs oneCell, tsle:timeslicer, a plain shape as fallback). */
export function timelineAnchorXml(t: Timeline, id: number): string {
  const pt = (tag: string, p: { col: number; row: number; x: number; y: number }) =>
    `<xdr:${tag}><xdr:col>${p.col}</xdr:col><xdr:colOff>${Math.round(p.x * EMU_PER_PX)}</xdr:colOff><xdr:row>${p.row}</xdr:row><xdr:rowOff>${Math.round(p.y * EMU_PER_PX)}</xdr:rowOff></xdr:${tag}>`;
  const w = Math.round((t.size?.w ?? 350) * EMU_PER_PX), h = Math.round((t.size?.h ?? 144) * EMU_PER_PX);
  const note = "Hier steht eine Zeitachse. Dieses Programm kann Zeitachsen nicht anzeigen.";
  return `<xdr:twoCellAnchor editAs="oneCell">${pt("from", t.from)}${pt("to", t.to)}`
    + `<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:Choice xmlns:tsle="http://schemas.microsoft.com/office/drawing/2012/timeslicer" Requires="tsle">`
    + `<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="${id}" name="${esc(t.name)}"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm>`
    + `<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/drawing/2012/timeslicer"><tsle:timeslicer xmlns:tsle="http://schemas.microsoft.com/office/drawing/2012/timeslicer" name="${esc(t.name)}"/></a:graphicData></a:graphic></xdr:graphicFrame></mc:Choice>`
    + `<mc:Fallback><xdr:sp macro="" textlink=""><xdr:nvSpPr><xdr:cNvPr id="0" name=""/><xdr:cNvSpPr><a:spLocks noTextEdit="1"/></xdr:cNvSpPr></xdr:nvSpPr><xdr:spPr><a:xfrm><a:off x="${Math.round((t.from.col * 64 + t.from.x) * EMU_PER_PX)}" y="${Math.round((t.from.row * 20 + t.from.y) * EMU_PER_PX)}"/><a:ext cx="${w}" cy="${h}"/></a:xfrm>`
    + `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:prstClr val="white"/></a:solidFill><a:ln w="1"><a:solidFill><a:prstClr val="green"/></a:solidFill></a:ln></xdr:spPr>`
    + `<xdr:txBody><a:bodyPr vertOverflow="clip" horzOverflow="clip"/><a:lstStyle/><a:p><a:r><a:rPr lang="de-DE" sz="1100"/><a:t>${note}</a:t></a:r></a:p></xdr:txBody></xdr:sp></mc:Fallback></mc:AlternateContent><xdr:clientData/></xdr:twoCellAnchor>`;
}

function timelineXml(t: Timeline): string {
  const a = [`name="${esc(t.name)}"`, `xr10:uid="${guid(`timeline:${t.name}`)}"`, `cache="${esc(t.cache)}"`, `caption="${esc(t.caption)}"`];
  for (const k of ["showHeader", "showSelectionLabel", "showTimeLevel", "showHorizontalScrollbar"] as const) if (t[k] === false) a.push(`${k}="0"`);
  a.push(`level="${t.level}"`, `selectionLevel="${t.selectionLevel ?? t.level}"`);
  if (t.scrollPosition !== undefined && Number.isFinite(t.scrollPosition)) a.push(`scrollPosition="${serialToIso(t.scrollPosition)}"`);
  if (t.style && t.style !== "TimeSlicerStyleLight1") a.push(`style="${esc(t.style)}"`);
  return `<timeline ${a.join(" ")}/>`;
}

function writeSheetTimelines(pw: PackageWriter, sheet: Sheet, part: string, xml: string): string {
  const list = sheet.timelines ?? [];
  const rels = pw.relsOf(part);
  for (const r of rels.list.filter((x) => isRel(x.type, RT_TIMELINE))) { rels.list.splice(rels.list.indexOf(r), 1); rels.changed = true; }
  let ext: string | null = null;
  if (list.length) {
    const tp = pw.freeName("xl/timelines/timeline");
    pw.setPart(tp, HEAD + `<timelines xmlns="${X15}" ${NS_TAIL}>${list.map(timelineXml).join("")}</timelines>`, CT_TIMELINE);
    const rid = pw.link(part, RT_TIMELINE, tp);
    ext = `<ext uri="${TL_URI.sheet}" xmlns:x15="${X15}"><x15:timelineRefs><x15:timelineRef r:id="${rid}"/></x15:timelineRefs></ext>`;
  }
  let out = patchExt(xml, "</worksheet>", TL_URI.sheet, ext);
  // The drawing: timeline anchors kept (moved ones with their new corners), removed or added.
  let drel = rels.list.find((r) => isRel(r.type, RT_DRAWING));
  let dpart = drel ? resolveTarget(part, drel.target) : undefined;
  let dxml = dpart ? str(pw.files, dpart) : "";
  const kept = new Set<Timeline>();
  let maxId = Math.max(1, ...[...dxml.matchAll(/<xdr:cNvPr\b[^>]*\bid="(\d+)"/g)].map((m) => Number(m[1])));
  if (dxml) {
    dxml = dxml.replace(/<xdr:(twoCellAnchor|oneCellAnchor|absoluteAnchor)\b[\s\S]*?<\/xdr:\1>/g, (ax) => {
      if (!timelineAnchorName(ax)) return ax;
      const t = list.find((x) => x.raw === ax && !kept.has(x));
      if (!t) return "";
      kept.add(t);
      if (t.rawKey === timelineKey(t)) return ax;
      const id = Number(/<xdr:cNvPr\b[^>]*\bid="(\d+)"/.exec(ax)?.[1] ?? ++maxId);
      return timelineAnchorXml(t, id);
    });
  }
  const fresh = list.filter((t) => !kept.has(t)).map((t) => timelineAnchorXml(t, ++maxId)).join("");
  if (fresh || (dxml && dpart)) {
    if (!dxml) {
      dpart = pw.freeName("xl/drawings/drawing");
      dxml = HEAD + `<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"></xdr:wsDr>`;
      const rid = pw.link(part, RT_DRAWING, dpart);
      drel = rels.list.find((r) => r.id === rid);
      const at = /<(legacyDrawing|legacyDrawingHF|drawingHF|picture|oleObjects|controls|webPublishItems|tableParts|extLst)\b/.exec(out);
      out = at ? out.slice(0, at.index) + `<drawing r:id="${rid}"/>` + out.slice(at.index) : out.replace("</worksheet>", `<drawing r:id="${rid}"/></worksheet>`);
    }
    dxml = dxml.replace(/<\/xdr:wsDr>\s*$/, `${fresh}</xdr:wsDr>`);
    if (!/<xdr:(twoCellAnchor|oneCellAnchor|absoluteAnchor)\b/.test(dxml) && drel) {
      rels.list.splice(rels.list.indexOf(drel), 1);
      rels.changed = true;
      out = out.replace(new RegExp(`<drawing r:id="${drel.id}"\\s*/>`), "");
    } else pw.setPart(dpart!, dxml, CT_DRAWING);
  }
  return out;
}

/** The pivot cache part of a PivotTable on a written sheet. */
function pivotCachePart(pw: PackageWriter, sheetPart: string, pivot: string): string | null {
  for (const r of pw.relsOf(sheetPart).list.filter((x) => x.type.endsWith("/pivotTable"))) {
    const tp = resolveTarget(sheetPart, r.target);
    if (attr(/<pivotTableDefinition\b[^>]*>/.exec(str(pw.files, tp))?.[0] ?? "", "name") !== pivot) continue;
    const cr = pw.relsOf(tp).list.find((x) => x.type.endsWith("/pivotCacheDefinition"));
    return cr ? resolveTarget(tp, cr.target) : null;
  }
  return null;
}

/** The dates of a PivotTable's field in its source (serials). */
function fieldDates(wb: Workbook, p: PivotTable, field: string): number[] {
  const src = wb.sheets.find((s) => s.name === p.source.sheet);
  if (!src) return [];
  let col = -1;
  for (let c = p.source.c0; c <= p.source.c1; c++) if (String(src.value(p.source.r0, c) ?? "").toLowerCase() === field.toLowerCase()) col = c;
  if (col < 0) return [];
  const out: number[] = [];
  for (let r = p.source.r0 + 1; r <= p.source.r1; r++) { const v = src.value(r, col); if (typeof v === "number") out.push(v); }
  return out;
}

function writeCaches(pw: PackageWriter, wb: Workbook, wbPath: string, wbXml: string, sheetIds: Map<string, number>, pivotCacheId: (sheet: string, pivot: string) => number | undefined, parts: Map<Sheet, string>): string {
  const wbRels = pw.relsOf(wbPath);
  for (const r of wbRels.list.filter((x) => isRel(x.type, RT_TIMELINE_CACHE))) { wbRels.list.splice(wbRels.list.indexOf(r), 1); wbRels.changed = true; }
  const ids: string[] = [];
  for (const c of wb.timelineCaches ?? []) {
    if (!c.pivots.length) continue;
    const first = c.pivots[0];
    const sheet = wb.sheets.find((s) => s.name === first.sheet);
    const p = sheet?.pivots?.find((x) => x.name === first.name);
    const sp = sheet ? parts.get(sheet) : undefined;
    if (!p || !sp) continue;
    // The PivotTable cache must carry the id the timeline refers to.
    let id = pivotCacheId(first.sheet, first.name) ?? c.pivotCacheId ?? 1;
    const cp = pivotCachePart(pw, sp, first.name);
    if (cp) {
      const cx = str(pw.files, cp);
      const own = /<x14:pivotCacheDefinition\b[^>]*pivotCacheId="(\d+)"/.exec(cx)?.[1];
      if (own) id = Number(own);
      else pw.setPart(cp, patchExt(cx, "</pivotCacheDefinition>", TL_URI.pivotCacheId, `<ext uri="${TL_URI.pivotCacheId}" xmlns:x14="${X14}"><x14:pivotCacheDefinition pivotCacheId="${id}"/></ext>`));
    }
    let field = -1;
    const src = wb.sheets.find((s) => s.name === p.source.sheet);
    for (let k = p.source.c0; src && k <= p.source.c1; k++) if (String(src.value(p.source.r0, k) ?? "").toLowerCase() === c.sourceName.toLowerCase()) field = k - p.source.c0;
    const sel = p.dateFilters?.find((d) => d.field === field);
    const b = timelineBounds(fieldDates(wb, p, c.sourceName));
    const bounds = b ? `<bounds startDate="${serialToIso(b.start)}" endDate="${serialToIso(b.end)}"/>` : "";
    const selection = sel ? `<selection startDate="${serialToIso(sel.from)}" endDate="${serialToIso(sel.to)}"/>` : "";
    const tabs = c.pivots.map((r) => `<pivotTable tabId="${sheetIds.get(r.sheet) ?? 1}" name="${esc(r.name)}"/>`).join("");
    const part = pw.freeName("xl/timelineCaches/timelineCache");
    pw.setPart(part, HEAD + `<timelineCacheDefinition xmlns="${X15}" xmlns:x15="${X15}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="xr10" xmlns:xr10="http://schemas.microsoft.com/office/spreadsheetml/2016/revision10" name="${esc(c.name)}" xr10:uid="${guid(`tlcache:${c.name}`)}" sourceName="${esc(c.sourceName)}">`
      + `<pivotTables>${tabs}</pivotTables><state minimalRefreshVersion="6" lastRefreshVersion="6" pivotCacheId="${id}" filterType="${sel ? "dateBetween" : "unknown"}">${selection}${bounds}</state></timelineCacheDefinition>`, CT_TIMELINE_CACHE);
    ids.push(pw.link(wbPath, RT_TIMELINE_CACHE, part));
  }
  const add = ids.length ? `<ext uri="${TL_URI.book}" xmlns:x15="${X15}"><x15:timelineCacheRefs>${ids.map((id) => `<x15:timelineCacheRef r:id="${id}"/>`).join("")}</x15:timelineCacheRefs></ext>` : null;
  return patchExt(wbXml, "</workbook>", TL_URI.book, add);
}

export const _test = { patchExt, anchors };
