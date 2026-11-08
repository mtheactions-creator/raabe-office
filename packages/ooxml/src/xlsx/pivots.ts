// PivotTables in .xlsx (D-066): a PivotTable made in RaabeCalc is written as Excel's parts, a cache
// definition (xl/pivotCache/pivotCacheDefinitionN.xml, refreshed by Excel when the file opens, no
// records stored) and a table definition (xl/pivotTables/pivotTableN.xml) linked from the sheet, so
// Excel opens it as a PivotTable. A PivotTable read from Excel's file keeps its parts as they are.

import { serialToDate, valueCaption, type PivotTable, type PivotValue, type Sheet } from "@raabe/sheet-engine";
import { strFromU8 } from "fflate";
import { Rels, relsPath, resolveTarget, type PackageWriter } from "../package";

export const RT_PIVOT_CACHE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition";
export const RT_PIVOT_TABLE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable";
const CT_CACHE = "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml";
const CT_TABLE = "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml";
const NS = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
const HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n';

const esc = (s: string) => s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
const col = (c: number) => { let s = ""; c++; while (c > 0) { const m = (c - 1) % 26; s = String.fromCharCode(65 + m) + s; c = Math.floor((c - 1) / 26); } return s; };
const ref = (r0: number, c0: number, r1: number, c1: number) => `${col(c0)}${r0 + 1}:${col(c1)}${r1 + 1}`;
const itemText = (v: PivotValue) => (v === null || v === "" ? "" : String(v));

/** The workbook's pivot caches written by this save (cacheId, relationship id). */
export interface PivotCacheRef { id: number; rid: string }

/** Writes the sheet's new PivotTables; returns their caches for workbook.xml. */
export function writePivots(pw: PackageWriter, sheet: Sheet, sheetPart: string, wbPath: string, data: (p: PivotTable) => PivotValue[][], nextId: () => number,
  slicers?: { fields(p: PivotTable): string[]; cacheId(p: PivotTable): number | undefined; dates?(p: PivotTable): string[] }): PivotCacheRef[] {
  const out: PivotCacheRef[] = [];
  for (const p of sheet.pivots ?? []) {
    const kept = p as PivotTable & { raw?: string; rawKey?: string };
    if (kept.raw && kept.rawKey === pivotKey(p)) continue;
    if (kept.raw) {
      // Changed in RaabeCalc: the sheet no longer points at the old table (its cache stays unused).
      const rels = pw.relsOf(sheetPart);
      const old = rels.list.find((r) => r.type.endsWith("/pivotTable") && resolveTarget(sheetPart, r.target) === kept.raw);
      if (old) { rels.list.splice(rels.list.indexOf(old), 1); rels.changed = true; }
    }
    const rows = data(p);
    if (!rows.length) continue;
    const names = rows[0].map((v) => itemText(v) || "Feld");
    const records = rows.slice(1);
    const axis = new Set([...p.rows, ...p.cols, ...p.filters.map((f) => f.field)]);
    // Fields a Datenschnitt filters (D-092): their items are listed in the cache and the table, as Excel lists them.
    const sliced = new Set((slicers?.fields(p) ?? []).map((n) => names.indexOf(n)).filter((f) => f >= 0));
    for (const f of sliced) axis.add(f);
    // Date fields a Zeitachse filters (D-121): listed as dates (<d>), as Excel lists them, so the timeline finds them.
    const dated = new Set([...(slicers?.dates?.(p) ?? []).map((n) => names.indexOf(n)), ...(p.dateFilters ?? []).map((d) => d.field)].filter((f) => f >= 0));
    for (const f of dated) { axis.add(f); sliced.add(f); }
    // Shared items of the fields on an axis (their items are referred to by index).
    const items = names.map((_, f) => (axis.has(f) ? [...new Map(records.map((r) => [itemText(r[f] ?? null), r[f] ?? null])).values()] : []));
    // The field lists its items sorted as Excel sorts a new PivotTable's items (numbers, text, empty last).
    const rank = (v: PivotValue) => (v === null || v === "" ? 3 : typeof v === "number" ? 0 : typeof v === "boolean" ? 2 : 1);
    const order = items.map((list) => list.map((_, k) => k).sort((a, b) => {
      const x = list[a], y = list[b];
      return rank(x) - rank(y) || (typeof x === "number" && typeof y === "number" ? x - y : String(x).localeCompare(String(y), "de", { sensitivity: "base" }));
    }));
    const cacheFields = names.map((n, f) => {
      const vals = records.map((r) => r[f] ?? null);
      const nums = vals.filter((v): v is number => typeof v === "number");
      const allNum = nums.length > 0 && vals.every((v) => v === null || typeof v === "number");
      if (!axis.has(f)) {
        return allNum
          ? `<cacheField name="${esc(n)}" numFmtId="0"><sharedItems containsSemiMixedTypes="0" containsString="0" containsNumber="1"${nums.every(Number.isInteger) ? ' containsInteger="1"' : ""} minValue="${Math.min(...nums)}" maxValue="${Math.max(...nums)}"/></cacheField>`
          : `<cacheField name="${esc(n)}" numFmtId="0"><sharedItems/></cacheField>`;
      }
      if (dated.has(f) && allNum) {
        const ds = items[f].filter((v): v is number => typeof v === "number");
        return `<cacheField name="${esc(n)}" numFmtId="14"><sharedItems containsSemiMixedTypes="0" containsNonDate="0" containsDate="1" containsString="0" minDate="${isoDate(Math.min(...ds))}" maxDate="${isoDate(Math.max(...ds) + 1)}" count="${items[f].length}">${items[f].map((v) => (typeof v === "number" ? `<d v="${isoDate(v)}"/>` : "<m/>")).join("")}</sharedItems></cacheField>`;
      }
      const it = items[f].map((v) => (v === null || v === "" ? "<m/>" : typeof v === "number" ? `<n v="${v}"/>` : typeof v === "boolean" ? `<b v="${v ? 1 : 0}"/>` : `<s v="${esc(String(v))}"/>`)).join("");
      const attrs = allNum ? ` containsSemiMixedTypes="0" containsString="0" containsNumber="1"${nums.every(Number.isInteger) ? ' containsInteger="1"' : ""} minValue="${Math.min(...nums)}" maxValue="${Math.max(...nums)}"` : "";
      return `<cacheField name="${esc(n)}" numFmtId="0"><sharedItems${attrs} count="${items[f].length}">${it}</sharedItems></cacheField>`;
    }).join("");
    const cachePart = pw.freeName("xl/pivotCache/pivotCacheDefinition");
    const s = p.source;
    pw.setPart(cachePart, HEAD + `<pivotCacheDefinition ${NS} refreshOnLoad="1" refreshedBy="RaabeCalc" createdVersion="8" refreshedVersion="8" minRefreshableVersion="3" recordCount="${records.length}" saveData="0">`
      + `<cacheSource type="worksheet">${p.sourceTable ? `<worksheetSource name="${esc(p.sourceTable)}"/>` : `<worksheetSource ref="${ref(s.r0, s.c0, s.r1, s.c1)}" sheet="${esc(s.sheet)}"/>`}</cacheSource><cacheFields count="${names.length}">${cacheFields}</cacheFields>`
      + (slicers?.cacheId(p) ? `<extLst><ext uri="{725AE2AE-9491-48be-B2B4-4EB974FC3084}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"><x14:pivotCacheDefinition pivotCacheId="${slicers.cacheId(p)}"/></ext></extLst>` : "")
      + `</pivotCacheDefinition>`, CT_CACHE);
    const rid = pw.link(wbPath, RT_PIVOT_CACHE, cachePart);
    const id = nextId();
    out.push({ id, rid });
    // The table: its fields, where it is, its data fields and layout.
    const layout = p.layout ?? "compact";
    // The layout is on each field (Excel: compact="0" for Gliederungsformat, also outline="0" for Tabellenformat).
    const fieldAttrs = layout === "compact" ? "" : layout === "outline" ? ' compact="0"' : ' compact="0" outline="0"';
    const sub = p.subtotals === "none" ? ' defaultSubtotal="0"' : p.subtotals === "bottom" && layout !== "tabular" ? ' subtotalTop="0"' : "";
    const pivotFields = names.map((_, f) => {
      const ax = p.rows.includes(f) ? "axisRow" : p.cols.includes(f) ? "axisCol" : p.filters.some((x) => x.field === f) ? "axisPage" : null;
      const data = p.values.some((v) => v.field === f) ? ' dataField="1"' : "";
      if (!ax && !sliced.has(f)) return `<pivotField${data}${fieldAttrs} showAll="0"/>`;
      const hidden = new Set((p.hidden?.[f] ?? []));
      const filt = p.filters.find((x) => x.field === f);
      const its = order[f].map((k) => { const v = items[f][k]; return `<item${hidden.has(itemText(v)) || (filt?.items && !filt.items.includes(itemText(v) || "(Leer)")) ? ' h="1"' : ""} x="${k}"/>`; }).join("");
      return `<pivotField${ax ? ` axis="${ax}"` : ""}${data}${dated.has(f) ? ' numFmtId="14"' : ""}${fieldAttrs}${ax && ax !== "axisPage" ? sub : ""} showAll="0"><items count="${items[f].length + 1}">${its}<item t="default"/></items></pivotField>`;
    }).join("");
    const area = p.area ?? { r0: p.at.row, c0: p.at.col, r1: p.at.row + 1, c1: p.at.col + 1 };
    const top = p.at.row;
    const headRows = p.cols.length ? 2 : 1;
    // Several value fields without a column field: no header row above them (firstHeaderRow 0, recorded).
    const firstHeader = !p.cols.length && p.values.length > 1 ? 0 : 1;
    const loc = `<location ref="${ref(top, p.at.col, Math.max(top + headRows, area.r1), Math.max(p.at.col + 1, area.c1))}" firstHeaderRow="${firstHeader}" firstDataRow="${headRows}" firstDataCol="${layout === "compact" ? 1 : Math.max(1, p.rows.length)}"${p.filters.length ? ` rowPageCount="${p.filters.length}" colPageCount="1"` : ""}/>`;
    const fields = (tag: string, list: number[]) => (list.length ? `<${tag} count="${list.length}">${list.map((f) => `<field x="${f}"/>`).join("")}</${tag}>` : "");
    const colList = p.values.length > 1 ? [...p.cols, -2] : p.cols;
    // Only the value fields across: their column items, or Excel adds a "Werte" header row on refresh (recorded).
    const valueCols = !p.cols.length && p.values.length > 1
      ? `<colItems count="${p.values.length}">${p.values.map((_, k) => (k ? `<i i="${k}"><x v="${k}"/></i>` : "<i><x/></i>")).join("")}</colItems>` : "";
    const pageFields = p.filters.length ? `<pageFields count="${p.filters.length}">${p.filters.map((f) => {
      const one = f.items?.length === 1 ? order[f.field].findIndex((k) => (itemText(items[f.field][k]) || "(Leer)") === f.items![0]) : -1;
      return `<pageField fld="${f.field}"${one >= 0 ? ` item="${one}"` : ""} hier="-1"/>`;
    }).join("")}</pageFields>` : "";
    const dataFields = p.values.length ? `<dataFields count="${p.values.length}">${p.values.map((v) =>
      `<dataField name="${esc(valueCaption(names, v, p.values))}" fld="${v.field}"${v.func !== "sum" ? ` subtotal="${v.func}"` : ""} baseField="0" baseItem="0"/>`).join("")}</dataFields>` : "";
    const tableAttrs = `${layout !== "compact" ? ' compact="0" compactData="0"' : ""} outline="1" outlineData="1"${p.colGrand === false ? ' rowGrandTotals="0"' : ""}${p.rowGrand === false ? ' colGrandTotals="0"' : ""}`;
    const table = HEAD + `<pivotTableDefinition ${NS} name="${esc(p.name)}" cacheId="${id}" applyNumberFormats="0" applyBorderFormats="0" applyFontFormats="0" applyPatternFormats="0" applyAlignmentFormats="0" applyWidthHeightFormats="1" dataCaption="Werte" updatedVersion="8" minRefreshableVersion="3" useAutoFormatting="1" itemPrintTitles="1" createdVersion="8" indent="0"${tableAttrs} multipleFieldFilters="0">`
      + loc + `<pivotFields count="${names.length}">${pivotFields}</pivotFields>` + fields("rowFields", p.rows) + fields("colFields", colList) + valueCols + pageFields + dataFields
      + `<pivotTableStyleInfo name="${esc(p.style ?? "PivotStyleLight16")}" showRowHeaders="1" showColHeaders="1" showRowStripes="0" showColStripes="0" showLastColumn="1"/>`
      + dateFiltersXml(p, names)
      // Excel 2010+ hides the "Werte" row of the value fields; without this it shows one (recorded).
      + `<extLst><ext uri="{962EF5D1-5CA2-4c93-8EF4-DBF5C05439D2}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"><x14:pivotTableDefinition hideValuesRow="1" xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main"/></ext></extLst></pivotTableDefinition>`;
    const tablePart = pw.freeName("xl/pivotTables/pivotTable");
    pw.setPart(tablePart, table, CT_TABLE);
    pw.link(tablePart, RT_PIVOT_CACHE, cachePart);
    pw.link(sheetPart, RT_PIVOT_TABLE, tablePart);
  }
  return out;
}

const attr = (tag: string, name: string) => new RegExp(`\\b${name}="([^"]*)"`).exec(tag)?.[1];
const unesc = (s: string) => s.replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, "&");
function cellRef(a: string): { row: number; col: number } {
  const m = /^\$?([A-Z]+)\$?(\d+)$/.exec(a)!;
  let c = 0;
  for (const ch of m[1]) c = c * 26 + ch.charCodeAt(0) - 64;
  return { row: Number(m[2]) - 1, col: c - 1 };
}

/** What a PivotTable's model holds (everything but its file parts), to see whether it changed. */
export function pivotKey(p: PivotTable): string {
  const { area: _a, raw: _r, rawKey: _k, ...rest } = p as PivotTable & { raw?: string; rawKey?: string };
  void _a; void _r; void _k;
  return JSON.stringify(rest);
}

/** PivotTables whose source is a table's name get the table's range (once every sheet's tables are read). */
export function resolvePivotSources(sheets: Sheet[]): void {
  for (const s of sheets) for (const p of s.pivots ?? []) {
    if (!p.sourceTable) continue;
    for (const ts of sheets) {
      const t = ts.tables?.find((x) => x.name.toLowerCase() === p.sourceTable!.toLowerCase());
      if (!t) continue;
      p.source = { sheet: ts.name, r0: t.ref.r0, c0: t.ref.c0, r1: t.ref.r1 - (t.totals ? 1 : 0), c1: t.ref.c1 };
      const k = p as PivotTable & { raw?: string; rawKey?: string };
      if (k.raw) k.rawKey = pivotKey(p);
    }
  }
}

/** The sheet's PivotTables as models (their parts stay in the file while unchanged). */
export function readPivots(files: Record<string, Uint8Array>, sheet: Sheet, sheetPart: string, sheetRels: Rels): void {
  const out: PivotTable[] = [];
  for (const r of sheetRels.list.filter((x) => x.type.endsWith("/pivotTable"))) {
    const part = resolveTarget(sheetPart, r.target);
    if (!files[part]) continue;
    const xml = strFromU8(files[part]);
    const prels = new Rels(files[relsPath(part)] ? strFromU8(files[relsPath(part)]) : null);
    const crel = prels.list.find((x) => x.type.endsWith("/pivotCacheDefinition"));
    const cpart = crel ? resolveTarget(part, crel.target) : undefined;
    const cache = cpart && files[cpart] ? strFromU8(files[cpart]) : "";
    const ws = /<worksheetSource\b[^>]*>/.exec(cache)?.[0] ?? "";
    const srcRef = attr(ws, "ref") ?? (attr(ws, "name") ? "A1" : undefined), srcSheet = attr(ws, "sheet");
    if (!srcRef) continue;
    // A source given by a table's name (Excel's default for a PivotTable made from a table): resolved once all
    // sheets and their tables are read (resolvePivotSources).
    const srcTable = attr(ws, "ref") ? undefined : unesc(attr(ws, "name") ?? "");
    const [a, b] = srcRef.split(":");
    const s0 = cellRef(a), s1 = cellRef(b ?? a);
    // Shared items of each cache field (for filters and hidden items).
    const fieldItems = [...cache.matchAll(/<cacheField\b[^>]*>([\s\S]*?)<\/cacheField>|<cacheField\b[^>]*\/>/g)].map((m) =>
      [...(m[1] ?? "").matchAll(/<(s|n|b|m|d|e)(?:\s+v="([^"]*)")?\s*\/>/g)].map((x) => (x[1] === "m" ? "(Leer)" : x[1] === "b" ? (x[2] === "1" ? "WAHR" : "FALSCH") : unesc(x[2] ?? ""))));
    const top = /<pivotTableDefinition\b[^>]*>/.exec(xml)?.[0] ?? "";
    const loc = /<location\b[^>]*>/.exec(xml)?.[0] ?? "";
    const [l0, l1] = (attr(loc, "ref") ?? "A1").split(":");
    const lt = cellRef(l0), lb = cellRef(l1 ?? l0);
    const fields = (tag: string) => [...(new RegExp(`<${tag}\\b[^>]*>([\\s\\S]*?)</${tag}>`).exec(xml)?.[1] ?? "").matchAll(/<field x="(-?\d+)"/g)].map((m) => Number(m[1])).filter((x) => x >= 0);
    const pivotFields = [...xml.matchAll(/<pivotField\b([^>]*?)(?:\/>|>([\s\S]*?)<\/pivotField>)/g)];
    // Each field's items in its own order: the cache index of each (filters and hidden items refer to these).
    const fieldOrder = pivotFields.map((m) => [...(m[2] ?? "").matchAll(/<item\b([^>]*)\/>/g)].filter((x) => !/\bt="/.test(x[1])).map((x) => Number(attr(x[1], "x") ?? 0)));
    const hidden: Record<number, string[]> = {};
    pivotFields.forEach((m, f) => {
      const h = [...(m[2] ?? "").matchAll(/<item\b[^>]*\bh="1"[^>]*\bx="(\d+)"|<item\b[^>]*\bx="(\d+)"[^>]*\bh="1"/g)].map((x) => fieldItems[f]?.[Number(x[1] ?? x[2])]).filter((x): x is string => x !== undefined);
      if (h.length) hidden[f] = h;
    });
    const p: PivotTable & { raw?: string; rawKey?: string } = {
      name: unesc(attr(top, "name") ?? "PivotTable"),
      source: { sheet: unesc(srcSheet ?? sheet.name), r0: s0.row, c0: s0.col, r1: s1.row, c1: s1.col },
      at: { row: lt.row, col: lt.col },
      rows: fields("rowFields"), cols: fields("colFields"),
      values: [...xml.matchAll(/<dataField\b([^>]*)\/?>/g)].map((m) => {
        const fn = attr(m[1], "subtotal") ?? "sum";
        const v: PivotTable["values"][number] = { field: Number(attr(m[1], "fld") ?? 0), func: fn as PivotTable["values"][number]["func"] };
        const name = attr(m[1], "name");
        if (name) v.name = unesc(name);
        return v;
      }),
      filters: [...xml.matchAll(/<pageField\b([^>]*)\/?>/g)].map((m) => {
        const f = Number(attr(m[1], "fld") ?? 0), item = attr(m[1], "item");
        const k = item !== undefined ? fieldOrder[f]?.[Number(item)] ?? Number(item) : undefined;
        return k !== undefined ? { field: f, items: [fieldItems[f]?.[k] ?? ""] } : { field: f };
      }),
      area: { r0: lt.row - (/<pageFields\b/.test(xml) ? fields("pageFields").length + 1 : 0), c0: lt.col, r1: lb.row, c1: lb.col },
    };
    if (Object.keys(hidden).length) p.hidden = hidden;
    // Date filters (a Zeitachse, D-121): <filter type="dateBetween"> with two custom filters in whole days.
    const dates: NonNullable<PivotTable["dateFilters"]> = [];
    for (const m of xml.matchAll(/<filter\b([^>]*)>([\s\S]*?)<\/filter>/g)) {
      if (attr(m[1], "type") !== "dateBetween") continue;
      const lo = /<customFilter\b[^>]*operator="greaterThanOrEqual"[^>]*val="([^"]*)"/.exec(m[2])?.[1];
      const hi = /<customFilter\b[^>]*operator="lessThanOrEqual"[^>]*val="([^"]*)"/.exec(m[2])?.[1];
      if (lo === undefined || hi === undefined) continue;
      dates.push({ field: Number(attr(m[1], "fld") ?? 0), from: Math.floor(Number(lo)), to: Math.floor(Number(hi)) });
    }
    if (dates.length) p.dateFilters = dates;
    if (srcTable) p.sourceTable = srcTable;
    // The layout from the first row field (Excel keeps it on the fields).
    const rowField = pivotFields[p.rows[0] ?? -1]?.[1] ?? pivotFields[0]?.[1] ?? "";
    if (attr(rowField, "compact") === "0") p.layout = attr(rowField, "outline") === "0" ? "tabular" : "outline";
    if (attr(top, "rowGrandTotals") === "0") p.colGrand = false;
    if (attr(top, "colGrandTotals") === "0") p.rowGrand = false;
    if (/defaultSubtotal="0"/.test(xml)) p.subtotals = "none";
    else if (/subtotalTop="0"/.test(xml)) p.subtotals = "bottom";
    // Names Excel gave automatically ("Summe von Note") are left to the model.
    p.values.forEach((v) => {
      const names = fieldItems.length ? [...cache.matchAll(/<cacheField\b[^>]*name="([^"]*)"/g)].map((m) => unesc(m[1])) : [];
      const label = { sum: "Summe", count: "Anzahl", average: "Mittelwert", max: "Max.", min: "Min.", product: "Produkt", countNums: "Anzahl Zahlen", stdDev: "StdAbw", stdDevp: "StdAbwN", var: "Varianz", varp: "Varianzen" }[v.func];
      if (v.name && new RegExp(`^${label.replace(".", "\\.")} von ${(names[v.field] ?? "").replace(/[.*+?^${}()|[\]\\]/g, "\\$&")}\\d*$`).test(v.name)) delete v.name;
    });
    p.raw = part;
    p.rawKey = pivotKey(p);
    out.push(p);
  }
  if (out.length) sheet.pivots = out;
}

/** A serial as Excel's date-time text in the pivot cache ("2025-10-01T00:00:00"). */
export function isoDate(serial: number): string {
  const d = serialToDate(Math.floor(serial));
  return `${String(d.y).padStart(4, "0")}-${String(d.m).padStart(2, "0")}-${String(d.d).padStart(2, "0")}T00:00:00`;
}

/** The PivotTable's date filters (a Zeitachse's selection) as Excel writes them: dateBetween in whole days. */
function dateFiltersXml(p: PivotTable, names: string[]): string {
  const list = p.dateFilters ?? [];
  if (!list.length) return "";
  return `<filters count="${list.length}">${list.map((d, i) => `<filter fld="${d.field}" type="dateBetween" evalOrder="-1" id="${i + 2}" name="${esc(names[d.field] ?? "")}"><autoFilter ref="A1"><filterColumn colId="0"><customFilters and="1">`
    + `<customFilter operator="greaterThanOrEqual" val="${d.from}"/><customFilter operator="lessThanOrEqual" val="${d.to}"/></customFilters></filterColumn></autoFilter>`
    + `<extLst><ext uri="{0605FD5F-26C8-4aeb-8148-2DB25E43C511}" xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main"><x15:pivotFilter useWholeDay="1"/></ext></extLst></filter>`).join("")}</filters>`;
}
