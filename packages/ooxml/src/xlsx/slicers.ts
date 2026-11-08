// Datenschnitte (slicers) in .xlsx (D-092), as German Excel 365 writes them (recorded with
// probe_excel_slicers.py, fixture excel-slicers.xlsx):
// - xl/slicerCaches/slicerCacheN.xml per cache, linked from the workbook and listed in its extLst
//   (x14:slicerCaches for PivotTable slicers, x15:slicerCaches for table slicers); a table slicer's
//   cache names the table and its column (x15:tableSlicerCache), a PivotTable slicer's cache names
//   the PivotTables and lists the items of the PivotTable cache with the selected ones marked;
// - xl/slicers/slicerN.xml per sheet (one for table slicers, one for PivotTable slicers), linked
//   from the sheet and listed in its extLst; each <slicer> names its cache;
// - a two-cell anchor in the sheet's drawing with an mc:AlternateContent graphicFrame naming the
//   slicer (a plain shape as fallback for programs without slicers);
// - a workbook name per cache ("Datenschnitt_Fach", #NV), kept in the model's names.
// The selection itself is the table's AutoFilter or the PivotTable's hidden items.
// While nothing about the slicers changed every part stays as it was read; after a change all slicer
// parts are written anew from the model.

import { strFromU8 } from "fflate";
import {
  BLANK_ITEM, compareItems, slicerKey,
  type Sheet, type Slicer, type SlicerCache, type Workbook,
} from "@raabe/sheet-engine";
import { PackageWriter, Rels, relsPath, resolveTarget } from "../package";
import { isRel, type BookRaw, type SheetRaw } from "./read";
import { canon } from "./styles";
import { esc, unesc } from "./xml";

export const RT_SLICER = "http://schemas.microsoft.com/office/2007/relationships/slicer";
export const RT_SLICER_CACHE = "http://schemas.microsoft.com/office/2007/relationships/slicerCache";
export const CT_SLICER = "application/vnd.ms-excel.slicer+xml";
export const CT_SLICER_CACHE = "application/vnd.ms-excel.slicerCache+xml";
const RT_DRAWING = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing";
const CT_DRAWING = "application/vnd.openxmlformats-officedocument.drawing+xml";
const HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n';
const X14 = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main";
const X15 = "http://schemas.microsoft.com/office/spreadsheetml/2010/11/main";
const SLICER_NS = `xmlns="${X14}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x xr10" xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:xr10="http://schemas.microsoft.com/office/spreadsheetml/2016/revision10"`;
/** extLst entries: workbook (PivotTable / table caches), sheet (PivotTable / table slicers), cache (table, hide items without data). */
export const URI = {
  bookPivot: "{BBE1A952-AA13-448e-AADC-164F8A28A991}",
  bookTable: "{46BE6895-7355-4a93-B00E-2C351335B9C9}",
  sheetPivot: "{A8765BA9-456A-4dab-B4F3-ACF838C121DE}",
  sheetTable: "{3A4CF648-6AED-40f4-86FF-DC5316D8AED3}",
  tableCache: "{2F2917AC-EB37-4324-AD4E-5DD8C200BD13}",
  hideNoData: "{470722E0-AACD-4C17-9CDC-17EF765DBC7E}",
  pivotCacheId: "{725AE2AE-9491-48be-B2B4-4EB974FC3084}",
};
const EMU_PER_PX = 9525;

const attr = (tag: string, name: string) => { const v = new RegExp(`\\b${name}="([^"]*)"`).exec(tag)?.[1]; return v === undefined ? undefined : unesc(v); };
const str = (files: Record<string, Uint8Array>, p: string | undefined) => (p && files[p] ? strFromU8(files[p]) : "");

interface SlicerSnap { snap: string }

/** What decides whether the slicers changed. */
export function slicerSnapshot(wb: Workbook): string {
  return canon({ caches: wb.slicerCaches ?? [], sheets: wb.sheets.map((s) => (s.slicers ?? []).map((x) => JSON.parse(slicerKey(x)))) });
}

/** The anchor of a drawing child that holds a slicer (sle:slicer), with the slicer's name. */
export function slicerAnchorName(ax: string): string | undefined {
  const m = /<sle:slicer\b[^>]*\bname="([^"]*)"/.exec(ax);
  return m ? unesc(m[1]) : undefined;
}

function point(ax: string, tag: string): { col: number; row: number; x: number; y: number } | undefined {
  const m = new RegExp(`<xdr:${tag}>([\\s\\S]*?)</xdr:${tag}>`).exec(ax);
  if (!m) return undefined;
  const v = (n: string) => Number(new RegExp(`<xdr:${n}>([^<]*)</xdr:${n}>`).exec(m[1])?.[1] ?? 0);
  return { col: v("col"), row: v("row"), x: v("colOff") / EMU_PER_PX, y: v("rowOff") / EMU_PER_PX };
}

/** The workbook's slicer caches and every sheet's slicers (with their drawing anchors). */
export function readSlicers(wb: Workbook): void {
  const raw = wb.raw as BookRaw | undefined;
  if (!raw) return;
  const files = raw.files;
  const wbRels = new Rels(str(files, relsPath(raw.wbPath)) || null);
  const caches: SlicerCache[] = [];
  const sheetById = new Map<number, Sheet>();
  for (const s of wb.sheets) { const r = s.raw as SheetRaw | undefined; if (r) sheetById.set(r.sheetId, s); }
  for (const r of wbRels.list.filter((x) => isRel(x.type, RT_SLICER_CACHE))) {
    const xml = str(files, resolveTarget(raw.wbPath, r.target));
    const top = /<slicerCacheDefinition\b[^>]*>/.exec(xml)?.[0];
    if (!top) continue;
    const c: SlicerCache = { name: attr(top, "name") ?? "", sourceName: attr(top, "sourceName") ?? "" };
    const tc = /<x15:tableSlicerCache\b[^>]*>/.exec(xml)?.[0];
    const tab = /<tabular\b[^>]*>/.exec(xml)?.[0];
    const opts = tc ?? tab ?? "";
    if (tc) {
      const id = Number(attr(tc, "tableId"));
      const colId = Number(attr(tc, "column"));
      for (const s of wb.sheets) for (const t of s.tables ?? []) {
        if (t.id !== id) continue;
        // The column is the table column's id: its position among the part's columns.
        const part = t.raw ? str(files, t.raw) : "";
        const ids = [...part.matchAll(/<tableColumn\b[^>]*\bid="(\d+)"/g)].map((m) => Number(m[1]));
        const at = ids.indexOf(colId);
        c.table = { name: t.name, column: at >= 0 ? at : colId - 1 };
      }
      if (!c.table) continue;
    } else if (tab) {
      c.pivots = [...xml.matchAll(/<pivotTable\b([^>]*)\/>/g)].map((m) => ({ sheet: sheetById.get(Number(attr(m[1], "tabId")))?.name ?? "", name: attr(m[1], "name") ?? "" }));
      const id = Number(attr(tab, "pivotCacheId"));
      if (id) c.pivotCacheId = id;
    } else continue;
    const sort = attr(opts, "sortOrder");
    if (sort === "descending" || sort === "dataSourceOrder") c.sort = sort;
    if (attr(opts, "customListSort") === "0") c.customListSort = false;
    const cross = attr(opts, "crossFilter");
    if (cross === "none" || cross === "showItemsWithNoData") c.crossFilter = cross;
    if (/<x15:slicerCacheHideItemsWithNoData\b/.test(xml)) c.hideNoData = true;
    caches.push(c);
  }
  if (caches.length) wb.slicerCaches = caches;
  for (const s of wb.sheets) {
    const sr = s.raw as SheetRaw | undefined;
    if (!sr) continue;
    const rels = new Rels(str(files, relsPath(sr.part)) || null);
    const list: Slicer[] = [];
    for (const r of rels.list.filter((x) => isRel(x.type, RT_SLICER))) {
      const xml = str(files, resolveTarget(sr.part, r.target));
      for (const m of xml.matchAll(/<slicer\b([^>]*?)\/?>/g)) {
        const a = m[1];
        const sl: Slicer = { name: attr(a, "name") ?? "", cache: attr(a, "cache") ?? "", caption: attr(a, "caption") ?? attr(a, "name") ?? "",
          from: { col: 0, row: 0, x: 0, y: 0 }, to: { col: 2, row: 10, x: 0, y: 0 } };
        if (attr(a, "showCaption") === "0") sl.showCaption = false;
        const cols = Number(attr(a, "columnCount") ?? 1);
        if (cols > 1) sl.columns = cols;
        const rh = Number(attr(a, "rowHeight") ?? 0);
        if (rh) sl.rowHeight = rh;
        const st = attr(a, "style");
        if (st) sl.style = st;
        list.push(sl);
      }
    }
    if (!list.length) continue;
    // Their places: the drawing's anchors naming them.
    const drel = rels.list.find((x) => isRel(x.type, RT_DRAWING));
    const dxml = drel ? str(files, resolveTarget(sr.part, drel.target)) : "";
    for (const ax of anchors(dxml)) {
      const name = slicerAnchorName(ax);
      const sl = list.find((x) => x.name === name);
      if (!sl) continue;
      sl.from = point(ax, "from") ?? sl.from;
      sl.to = point(ax, "to") ?? sl.to;
      sl.raw = ax;
    }
    for (const sl of list) if (sl.raw) sl.rawKey = slicerKey(sl);
    s.slicers = list;
  }
  (raw as BookRaw & { slicers?: SlicerSnap }).slicers = { snap: slicerSnapshot(wb) };
}

/** The drawing's top-level anchors as text. */
function anchors(xml: string): string[] {
  const out: string[] = [];
  const re = /<xdr:(twoCellAnchor|oneCellAnchor|absoluteAnchor)\b[\s\S]*?<\/xdr:\1>/g;
  for (const m of xml.matchAll(re)) out.push(m[0]);
  return out;
}

// ----- writing ------------------------------------------------------------------------------

export interface SlicerPlan {
  changed: boolean;
  /** For writePivots: the slicer fields (by name) a PivotTable's cache must list, and its x14 pivotCacheId. */
  pivotFields(sheet: string, pivot: string): string[];
  pivotCacheId(sheet: string, pivot: string): number | undefined;
  /** The sheet's XML after its slicer parts, drawing anchors and extLst entries were written. */
  sheetXml(sheet: Sheet, part: string, xml: string): string;
  /** Writes the caches; returns workbook.xml with its extLst entries. */
  finish(wbXml: string, sheetIds: Map<string, number>): string;
}

const pivotRewritten = (p: { raw?: string; rawKey?: string }, key: (p: never) => string) => !p.raw || p.rawKey !== key(p as never);

/** Decides before the sheets are written whether the slicer parts are rewritten. */
export function planSlicers(wb: Workbook, pw: PackageWriter, wbPath: string, pivotKey: (p: never) => string): SlicerPlan {
  const raw = wb.raw as (BookRaw & { slicers?: SlicerSnap }) | undefined;
  const caches = wb.slicerCaches ?? [];
  const pivotOf = (sheet: string, name: string) => wb.sheets.find((s) => s.name === sheet)?.pivots?.find((p) => p.name === name);
  const anyPivotRewritten = caches.some((c) => (c.pivots ?? []).some((r) => { const p = pivotOf(r.sheet, r.name); return !!p && pivotRewritten(p, pivotKey); }));
  const had = !!raw?.slicers;
  const changed = (caches.length > 0 || wb.sheets.some((s) => s.slicers?.length) || had)
    && (!had || raw!.slicers!.snap !== slicerSnapshot(wb) || anyPivotRewritten);
  // New x14 pivotCacheIds count on from the file's own.
  let nextId = Math.max(1000, ...Object.keys(pw.files).filter((p) => p.includes("pivotCache")).map((p) => Number(/pivotCacheId="(\d+)"/.exec(str(pw.files, p))?.[1] ?? 0))) + 1;
  const ids = new Map<string, number>();
  const pivotCacheId = (sheet: string, pivot: string) => {
    const k = `${sheet}\u0000${pivot}`;
    if (!caches.some((c) => c.pivots?.some((r) => r.sheet === sheet && r.name === pivot))) return undefined;
    if (!ids.has(k)) ids.set(k, nextId++);
    return ids.get(k);
  };
  const parts = new Map<Sheet, string>();
  const plan: SlicerPlan = {
    changed,
    pivotFields: (sheet, pivot) => (changed ? caches.filter((c) => c.pivots?.some((r) => r.sheet === sheet && r.name === pivot)).map((c) => c.sourceName) : []),
    pivotCacheId: (sheet, pivot) => (changed ? pivotCacheId(sheet, pivot) : undefined),
    sheetXml: (sheet, part, xml) => { parts.set(sheet, part); return changed ? writeSheetSlicers(pw, wb, sheet, part, xml) : xml; },
    finish: (wbXml, sheetIds) => (changed ? writeCaches(pw, wb, wbPath, wbXml, sheetIds, pivotCacheId, parts) : wbXml),
  };
  return plan;
}

/** Removes the extLst entries with the given uris and adds `add` (an extLst made when there was none). */
function patchExt(xml: string, close: string, uris: string[], add: string[], firstUris: string[] = []): string {
  const m = /<extLst>([\s\S]*?)<\/extLst>/.exec(xml);
  let inner = m ? m[1] : "";
  for (const u of uris) inner = inner.replace(new RegExp(`<(\\w+:)?ext\\b[^>]*uri="${u.replace(/[{}]/g, "\\$&")}"[\\s\\S]*?<\\/\\1?ext>`), "");
  // Excel's order: the slicer entries first in the workbook (PivotTable caches before the rest).
  const first = add.filter((a) => firstUris.some((u) => a.includes(u)));
  const rest = add.filter((a) => !first.includes(a));
  inner = first.join("") + inner + rest.join("");
  if (m) return inner ? xml.replace(m[0], `<extLst>${inner}</extLst>`) : xml.replace(m[0], "");
  return inner ? xml.replace(close, `<extLst>${inner}</extLst>${close}`) : xml;
}

/** A slicer's drawing anchor as Excel writes a new one (absolute anchor for table slicers, as recorded). */
export function slicerAnchorXml(s: Slicer, id: number, table: boolean): string {
  const pt = (tag: string, p: { col: number; row: number; x: number; y: number }) =>
    `<xdr:${tag}><xdr:col>${p.col}</xdr:col><xdr:colOff>${Math.round(p.x * EMU_PER_PX)}</xdr:colOff><xdr:row>${p.row}</xdr:row><xdr:rowOff>${Math.round(p.y * EMU_PER_PX)}</xdr:rowOff></xdr:${tag}>`;
  const w = Math.round((s.size?.w ?? 192) * EMU_PER_PX), h = Math.round((s.size?.h ?? 275) * EMU_PER_PX);
  // The fallback shape for programs without slicers: our own words (D-058).
  const note = "Hier steht ein Datenschnitt. Dieses Programm kann Datenschnitte nicht anzeigen.";
  return `<xdr:twoCellAnchor editAs="${table ? "absolute" : "oneCell"}">${pt("from", s.from)}${pt("to", s.to)}`
    + `<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:Choice xmlns:sle15="http://schemas.microsoft.com/office/drawing/2012/slicer" Requires="sle15">`
    + `<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="${id}" name="${esc(s.name)}"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm>`
    + `<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/drawing/2010/slicer"><sle:slicer xmlns:sle="http://schemas.microsoft.com/office/drawing/2010/slicer" name="${esc(s.name)}"/></a:graphicData></a:graphic></xdr:graphicFrame></mc:Choice>`
    + `<mc:Fallback><xdr:sp macro="" textlink=""><xdr:nvSpPr><xdr:cNvPr id="0" name=""/><xdr:cNvSpPr><a:spLocks noTextEdit="1"/></xdr:cNvSpPr></xdr:nvSpPr><xdr:spPr><a:xfrm><a:off x="${Math.round((s.from.col * 64 + s.from.x) * EMU_PER_PX)}" y="${Math.round((s.from.row * 20 + s.from.y) * EMU_PER_PX)}"/><a:ext cx="${w}" cy="${h}"/></a:xfrm>`
    + `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:prstClr val="white"/></a:solidFill><a:ln w="1"><a:solidFill><a:prstClr val="green"/></a:solidFill></a:ln></xdr:spPr>`
    + `<xdr:txBody><a:bodyPr vertOverflow="clip" horzOverflow="clip"/><a:lstStyle/><a:p><a:r><a:rPr lang="de-DE" sz="1100"/><a:t>${note}</a:t></a:r></a:p></xdr:txBody></xdr:sp></mc:Fallback></mc:AlternateContent><xdr:clientData/></xdr:twoCellAnchor>`;
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

/** One sheet: its slicer parts, extLst entries and drawing anchors. */
function writeSheetSlicers(pw: PackageWriter, wb: Workbook, sheet: Sheet, part: string, xml: string): string {
  const list = sheet.slicers ?? [];
  const cacheOf = (s: Slicer) => wb.slicerCaches?.find((c) => c.name === s.cache);
  const rels = pw.relsOf(part);
  for (const r of rels.list.filter((x) => isRel(x.type, RT_SLICER))) { rels.list.splice(rels.list.indexOf(r), 1); rels.changed = true; }
  const ext: string[] = [];
  const slicerXml = (s: Slicer) => {
    const a = [`name="${esc(s.name)}"`, `xr10:uid="${guid(`slicer:${s.name}`)}"`, `cache="${esc(s.cache)}"`, `caption="${esc(s.caption)}"`];
    if ((s.columns ?? 1) > 1) a.push(`columnCount="${s.columns}"`);
    if (s.showCaption === false) a.push('showCaption="0"');
    if (s.style && s.style !== "SlicerStyleLight1") a.push(`style="${esc(s.style)}"`);
    a.push(`rowHeight="${Math.round(s.rowHeight ?? 251883)}"`);
    return `<slicer ${a.join(" ")}/>`;
  };
  for (const kind of ["pivot", "table"] as const) {
    const these = list.filter((s) => { const c = cacheOf(s); return c && (kind === "table" ? !!c.table : !!c.pivots); });
    if (!these.length) continue;
    const sp = pw.freeName("xl/slicers/slicer");
    pw.setPart(sp, HEAD + `<slicers ${SLICER_NS}>${these.map(slicerXml).join("")}</slicers>`, CT_SLICER);
    const rid = pw.link(part, RT_SLICER, sp);
    ext.push(kind === "table"
      ? `<ext uri="${URI.sheetTable}" xmlns:x15="${X15}"><x14:slicerList xmlns:x14="${X14}"><x14:slicer r:id="${rid}"/></x14:slicerList></ext>`
      : `<ext uri="${URI.sheetPivot}" xmlns:x14="${X14}"><x14:slicerList><x14:slicer r:id="${rid}"/></x14:slicerList></ext>`);
  }
  let out = patchExt(xml, "</worksheet>", [URI.sheetPivot, URI.sheetTable], ext);
  // The drawing: slicer anchors kept (moved ones with their new corners), removed or added.
  let drel = rels.list.find((r) => isRel(r.type, RT_DRAWING));
  let dpart = drel ? resolveTarget(part, drel.target) : undefined;
  let dxml = dpart ? str(pw.files, dpart) : "";
  const kept = new Set<Slicer>();
  let maxId = Math.max(1, ...[...dxml.matchAll(/<xdr:cNvPr\b[^>]*\bid="(\d+)"/g)].map((m) => Number(m[1])));
  if (dxml) {
    dxml = dxml.replace(/<xdr:(twoCellAnchor|oneCellAnchor|absoluteAnchor)\b[\s\S]*?<\/xdr:\1>/g, (ax) => {
      if (!slicerAnchorName(ax)) return ax;
      const s = list.find((x) => x.raw === ax && !kept.has(x));
      if (!s) return "";
      kept.add(s);
      if (s.rawKey === slicerKey(s)) return ax;
      const id = Number(/<xdr:cNvPr\b[^>]*\bid="(\d+)"/.exec(ax)?.[1] ?? ++maxId);
      return slicerAnchorXml(s, id, !!cacheOf(s)?.table);
    });
  }
  const fresh = list.filter((s) => !kept.has(s)).map((s) => slicerAnchorXml(s, ++maxId, !!cacheOf(s)?.table)).join("");
  if (fresh || (dxml && dpart)) {
    if (!dxml) {
      dpart = pw.freeName("xl/drawings/drawing");
      dxml = HEAD + `<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"></xdr:wsDr>`;
      const rid = pw.link(part, RT_DRAWING, dpart);
      drel = rels.list.find((r) => r.id === rid);
      // <drawing> goes before legacyDrawing, …, tableParts and extLst (CT_Worksheet's order).
      const at = /<(legacyDrawing|legacyDrawingHF|drawingHF|picture|oleObjects|controls|webPublishItems|tableParts|extLst)\b/.exec(out);
      out = at ? out.slice(0, at.index) + `<drawing r:id="${rid}"/>` + out.slice(at.index) : out.replace("</worksheet>", `<drawing r:id="${rid}"/></worksheet>`);
    }
    dxml = dxml.replace(/<\/xdr:wsDr>\s*$/, `${fresh}</xdr:wsDr>`);
    if (!/<xdr:(twoCellAnchor|oneCellAnchor|absoluteAnchor)\b/.test(dxml) && drel) {
      // Nothing left in a drawing that held only slicers: it goes.
      rels.list.splice(rels.list.indexOf(drel), 1);
      rels.changed = true;
      out = out.replace(new RegExp(`<drawing r:id="${drel.id}"\\s*/>`), "");
    } else pw.setPart(dpart!, dxml, CT_DRAWING);
  }
  return out;
}

/** The items of a PivotTable cache's field (texts in the cache's order) and its x14 id, from the written package. */
function pivotCacheField(pw: PackageWriter, sheetPart: string, pivot: string, field: string): { cachePart: string; index: number; items: string[]; numbers: (number | null)[] } | null {
  for (const r of pw.relsOf(sheetPart).list.filter((x) => x.type.endsWith("/pivotTable"))) {
    const tp = resolveTarget(sheetPart, r.target);
    const txml = str(pw.files, tp);
    if (attr(/<pivotTableDefinition\b[^>]*>/.exec(txml)?.[0] ?? "", "name") !== pivot) continue;
    const cr = pw.relsOf(tp).list.find((x) => x.type.endsWith("/pivotCacheDefinition"));
    if (!cr) continue;
    const cachePart = resolveTarget(tp, cr.target);
    const cxml = str(pw.files, cachePart);
    let index = -1;
    for (const m of cxml.matchAll(/<cacheField\b([^>]*?)(?:\/>|>([\s\S]*?)<\/cacheField>)/g)) {
      index++;
      if (attr(m[1], "name") !== field) continue;
      const its = [...(m[2] ?? "").matchAll(/<(s|n|b|m|d|e)(?:\s+v="([^"]*)")?\s*\/>/g)];
      return {
        cachePart, index,
        items: its.map((x) => (x[1] === "m" ? BLANK_ITEM : x[1] === "b" ? (x[2] === "1" ? "WAHR" : "FALSCH") : unesc(x[2] ?? ""))),
        numbers: its.map((x) => (x[1] === "n" ? Number(x[2]) : null)),
      };
    }
  }
  return null;
}

function writeCaches(pw: PackageWriter, wb: Workbook, wbPath: string, wbXml: string, sheetIds: Map<string, number>, pivotCacheId: (sheet: string, pivot: string) => number | undefined, parts: Map<Sheet, string>): string {
  const wbRels = pw.relsOf(wbPath);
  for (const r of wbRels.list.filter((x) => isRel(x.type, RT_SLICER_CACHE))) { wbRels.list.splice(wbRels.list.indexOf(r), 1); wbRels.changed = true; }
  const pivotIds: string[] = [], tableIds: string[] = [];
  const partOfSheet = (name: string) => { const s = wb.sheets.find((x) => x.name === name); return s ? parts.get(s) : undefined; };
  for (const c of wb.slicerCaches ?? []) {
    const opts: string[] = [];
    if (c.sort && c.sort !== "ascending") opts.push(`sortOrder="${c.sort}"`);
    if (c.customListSort === false) opts.push('customListSort="0"');
    const cross = c.crossFilter && c.crossFilter !== "showItemsWithDataAtTop" ? ` crossFilter="${c.crossFilter}"` : "";
    const hide = c.hideNoData ? `<x:ext uri="${URI.hideNoData}" xmlns:x15="${X15}"><x15:slicerCacheHideItemsWithNoData/></x:ext>` : "";
    let body = "";
    if (c.table) {
      // The table's id and the column's id as the written table part has them.
      let tableId = 0, colId = c.table.column + 1;
      for (const s of wb.sheets) {
        const sp = parts.get(s);
        if (!(s.tables ?? []).some((t) => t.name === c.table!.name) || !sp) continue;
        for (const r of pw.relsOf(sp).list.filter((x) => x.type.endsWith("/table"))) {
          const txml = str(pw.files, resolveTarget(sp, r.target));
          const top = /<table\b[^>]*>/.exec(txml)?.[0] ?? "";
          if (attr(top, "name") !== c.table.name && attr(top, "displayName") !== c.table.name) continue;
          tableId = Number(attr(top, "id") ?? 0);
          const cols = [...txml.matchAll(/<tableColumn\b[^>]*>/g)].map((m) => ({ id: Number(attr(m[0], "id")), name: attr(m[0], "name") }));
          colId = cols.find((x) => x.name === c.sourceName)?.id ?? cols[c.table.column]?.id ?? colId;
        }
      }
      if (!tableId) continue;
      body = `<extLst><x:ext uri="${URI.tableCache}" xmlns:x15="${X15}"><x15:tableSlicerCache tableId="${tableId}" column="${colId}"${opts.length ? " " + opts.join(" ") : ""}${cross}/></x:ext>${hide}</extLst>`;
    } else if (c.pivots?.length) {
      const first = c.pivots[0];
      const sp = partOfSheet(first.sheet);
      const field = sp ? pivotCacheField(pw, sp, first.name, c.sourceName) : null;
      const id = pivotCacheId(first.sheet, first.name) ?? c.pivotCacheId ?? 1;
      const p = wb.sheets.find((s) => s.name === first.sheet)?.pivots?.find((x) => x.name === first.name);
      const hidden = new Set<string>();
      // The field's hidden items in the PivotTable (the slicer's selection).
      if (p && field) for (const t of p.hidden?.[field.index] ?? []) hidden.add(t.toLowerCase());
      let items = "";
      if (field) {
        // The cache must carry the id the slicer refers to.
        const cx = str(pw.files, field.cachePart);
        if (!new RegExp(`pivotCacheId="${id}"`).test(cx)) {
          const e = `<ext uri="${URI.pivotCacheId}" xmlns:x14="${X14}"><x14:pivotCacheDefinition pivotCacheId="${id}"/></ext>`;
          pw.setPart(field.cachePart, patchExt(cx, "</pivotCacheDefinition>", [URI.pivotCacheId], [e]));
        }
        const order = field.items.map((t, k) => ({ text: t, num: field.numbers[k], k }));
        if (c.sort !== "dataSourceOrder") order.sort((a, b) => compareItems(a, b, c.sort ?? "ascending"));
        items = `<items count="${order.length}">${order.map((o) => `<i x="${o.k}"${hidden.has(o.text.toLowerCase()) ? "" : ' s="1"'}/>`).join("")}</items>`;
      }
      const tabs = c.pivots.map((r) => `<pivotTable tabId="${sheetIds.get(r.sheet) ?? 1}" name="${esc(r.name)}"/>`).join("");
      body = `<pivotTables>${tabs}</pivotTables><data><tabular pivotCacheId="${id}"${opts.length ? " " + opts.join(" ") : ""}${cross}>${items}</tabular></data>${hide ? `<extLst>${hide}</extLst>` : ""}`;
    } else continue;
    const part = pw.freeName("xl/slicerCaches/slicerCache");
    pw.setPart(part, HEAD + `<slicerCacheDefinition ${SLICER_NS} name="${esc(c.name)}" xr10:uid="${guid(`cache:${c.name}`)}" sourceName="${esc(c.sourceName)}">${body}</slicerCacheDefinition>`, CT_SLICER_CACHE);
    (c.table ? tableIds : pivotIds).push(pw.link(wbPath, RT_SLICER_CACHE, part));
  }
  const add: string[] = [];
  if (pivotIds.length) add.push(`<ext uri="${URI.bookPivot}" xmlns:x14="${X14}"><x14:slicerCaches>${pivotIds.map((id) => `<x14:slicerCache r:id="${id}"/>`).join("")}</x14:slicerCaches></ext>`);
  if (tableIds.length) add.push(`<ext uri="${URI.bookTable}" xmlns:x15="${X15}"><x15:slicerCaches xmlns:x14="${X14}">${tableIds.map((id) => `<x14:slicerCache r:id="${id}"/>`).join("")}</x15:slicerCaches></ext>`);
  return patchExt(wbXml, "</workbook>", [URI.bookPivot, URI.bookTable], add, [URI.bookPivot]);
}

/** Whether a drawing anchor is a slicer (or timeline) frame, which the shape and picture readers leave alone. */
export function isSlicerAnchor(ax: string): boolean {
  return /<(sle|tsle):(slicer|timeslicer)\b/.test(ax);
}

export const _test = { patchExt, anchors };
