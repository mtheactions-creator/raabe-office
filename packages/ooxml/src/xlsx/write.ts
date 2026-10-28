// Writing RaabeCalc's workbook model as .xlsx. The package read from the file is the base:
// every part and every worksheet element the model does not hold is written back unchanged
// (D-024), modelled elements the user did not change keep Excel's text, the rest is written
// in schema order the way Excel writes it. Shared formulas are written per cell, strings into
// a fresh shared-string table; calcChain.xml is dropped (Excel rebuilds it).

import { strFromU8, strToU8, zipSync } from "fflate";
import {
  CellError, cellName, fromXlsxFormula, keyCell, rangeName, toXlsxFormula,
  type Cell, type ColInfo, type DefinedName, type RowInfo, type Sheet, type View, type Workbook,
} from "@raabe/sheet-engine";
import { PackageWriter, Rels, relsPath, resolveTarget } from "../package";
import { RT, isRel, nameKey, readPackage, sheetNamesSnap, sheetSnap, type BookRaw, type SheetRaw } from "./read";
import { canon, StylesWriter } from "./styles";
import { commentsXml, vmlBlocks, vmlXml, type NoteIn } from "./notes";
import { colorXml } from "./colors";
import { writeCharts } from "./drawing";
import { newPackage } from "./new";
import { numText, widthFromBase, widthToFile } from "./units";
import { esc, escText, encodeX, joinSplit, place, placeAll, setAttrs, split, tXml, type Split, XML_HEAD } from "./xml";
import { FILTER_DB, PRINT_AREA, PRINT_TITLES, titlesRef, breaksXml, readPage, headerFooterXml, autoFilterXml, conditionalXml, fitToPage, pageXml, sheetRef, validationsXml } from "./sheetparts";

export const CT = {
  worksheet: "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml",
  sharedStrings: "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml",
  comments: "application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml",
  vml: "application/vnd.openxmlformats-officedocument.vmlDrawing",
};

/** CT_Worksheet's order of children. */
export const SHEET_ORDER = [
  "sheetPr", "dimension", "sheetViews", "sheetFormatPr", "cols", "sheetData", "sheetCalcPr", "sheetProtection", "protectedRanges",
  "scenarios", "autoFilter", "sortState", "dataConsolidate", "customSheetViews", "mergeCells", "phoneticPr", "conditionalFormatting",
  "dataValidations", "hyperlinks", "printOptions", "pageMargins", "pageSetup", "headerFooter", "rowBreaks", "colBreaks",
  "customProperties", "cellWatches", "ignoredErrors", "smartTags", "drawing", "legacyDrawing", "legacyDrawingHF", "drawingHF",
  "picture", "oleObjects", "controls", "webPublishItems", "tableParts", "extLst",
];

/** CT_Workbook's order of children. */
const BOOK_ORDER = [
  "fileVersion", "fileSharing", "workbookPr", "workbookProtection", "bookViews", "sheets", "functionGroups", "externalReferences",
  "definedNames", "calcPr", "oleSize", "customWorkbookViews", "pivotCaches", "smartTagPr", "smartTagTypes", "webPublishing",
  "fileRecoveryPr", "webPublishObjects", "extLst",
];

/** Elements that refer to parts of their own sheet: left out when a sheet is a copy. */
const PART_REFS = new Set(["drawing", "legacyDrawing", "legacyDrawingHF", "drawingHF", "picture", "oleObjects", "controls", "tableParts"]);

export interface XlsxWriteOptions {
  /** Author of new notes and last author in the document properties. */
  author?: string;
  now?: Date;
}

export function writeXlsx(wb: Workbook, options: XlsxWriteOptions = {}): Uint8Array {
  return zipSync(writePackage(wb, options), { level: 6 });
}

/** The package's files for a workbook (before zipping). */
export function writePackage(wb: Workbook, options: XlsxWriteOptions = {}): Record<string, Uint8Array> {
  const raw = (wb.raw as BookRaw | undefined) ?? (readPackage(newPackage(false)).raw as BookRaw);
  const pw = new PackageWriter(raw.files);
  const before = reachable(raw.files);
  const sw = new StylesWriter(raw.styles, wb.styles, raw.theme);
  const sst = new SharedStrings();
  const wbRels = pw.relsOf(raw.wbPath);
  const used = new Set<string>();
  const entries: { name: string; sheetId: number; state?: string; rid: string; attrs: Record<string, string> }[] = [];
  let maxId = Math.max(0, ...wb.sheets.map((s) => (s.raw as SheetRaw | undefined)?.sheetId ?? 0), ...raw.others.map((o) => Number(/sheetId="(\d+)"/.exec(o.xml)?.[1] ?? 0)));
  let needsCalc = false;
  const vmlIds = vmlBlocks(Object.keys(raw.files).filter((p) => p.endsWith(".vml")).map((p) => strFromU8(raw.files[p])));

  wb.sheets.forEach((sheet, i) => {
    const sr = sheet.raw as SheetRaw | undefined;
    const own = !!sr && !used.has(sr.part) && !!raw.files[sr.part];
    const part = own ? sr!.part : pw.freeName("xl/worksheets/sheet");
    used.add(part);
    const w = new SheetWriter(sheet, i === wb.active, own ? sr : undefined, sr, part, pw, sw, sst, raw, options.author ?? "", vmlIds);
    pw.setPart(part, w.xml(), CT.worksheet);
    if (w.formulaWithoutValue) needsCalc = true;
    let rid = own ? wbRels.list.find((r) => isRel(r.type, RT.worksheet) && resolveTarget(raw.wbPath, r.target) === part)?.id : undefined;
    if (!rid) rid = pw.link(raw.wbPath, RT.worksheet, part);
    const sheetId = own ? sr!.sheetId : ++maxId;
    const state = sheet.hidden === "very" ? "veryHidden" : sheet.hidden ? "hidden" : undefined;
    entries.push({ name: sheet.name, sheetId, state, rid, attrs: own ? sr!.attrs : {} });
  });

  // Worksheets that were deleted, and the calculation chain, lose their relationships.
  for (const r of [...wbRels.list]) {
    if ((isRel(r.type, RT.worksheet) && !used.has(resolveTarget(raw.wbPath, r.target))) || isRel(r.type, RT.calcChain)) removeRel(wbRels, r.id);
  }

  // Shared strings.
  const sstPath = raw.sstPath ?? "xl/sharedStrings.xml";
  if (sst.list.length) {
    pw.setPart(sstPath, sst.xml(), CT.sharedStrings);
    pw.link(raw.wbPath, RT.sharedStrings, sstPath);
  } else {
    for (const r of [...wbRels.list]) if (isRel(r.type, RT.sharedStrings)) removeRel(wbRels, r.id);
  }

  // Styles (after the sheets, which add the formats they use).
  const stylesRel = wbRels.list.find((r) => isRel(r.type, RT.styles));
  const stylesXml = sw.xml();
  if (stylesRel && stylesXml !== raw.styles.xml) pw.setPart(resolveTarget(raw.wbPath, stylesRel.target), stylesXml);

  pw.setPart(raw.wbPath, workbookXml(wb, raw, entries, needsCalc));
  props(pw, wb, raw, options);

  const files = pw.finish();
  // Parts only the removed sheets (or calcChain) used go too; orphans of the original stay.
  const after = reachable(files);
  for (const p of before) {
    if (after.has(p)) continue;
    delete files[p];
    delete files[relsPath(p)];
    dropOverride(files, p);
  }
  return files;
}

function removeRel(rels: Rels, id: string): void {
  const i = rels.list.findIndex((r) => r.id === id);
  if (i >= 0) {
    rels.list.splice(i, 1);
    rels.changed = true;
  }
}

/** Every part reachable through relationships from the package root. */
function reachable(files: Record<string, Uint8Array>): Set<string> {
  const out = new Set<string>();
  const visit = (part: string) => {
    const f = files[relsPath(part)] ?? (part === "" ? files["_rels/.rels"] : undefined);
    if (!f) return;
    for (const r of new Rels(strFromU8(f)).list) {
      if (r.external) continue;
      const target = resolveTarget(part, r.target);
      if (out.has(target) || !files[target]) continue;
      out.add(target);
      visit(target);
    }
  };
  visit("");
  return out;
}

function dropOverride(files: Record<string, Uint8Array>, part: string): void {
  const ct = files["[Content_Types].xml"];
  if (!ct) return;
  const xml = strFromU8(ct);
  const re = new RegExp(`<Override PartName="/${part.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")}"[^>]*/>`);
  if (re.test(xml)) files["[Content_Types].xml"] = strToU8(xml.replace(re, ""));
}

class SharedStrings {
  readonly list: string[] = [];
  private ids = new Map<string, number>();
  count = 0;

  /** Index of a string (<si> content as XML). */
  id(si: string): number {
    this.count++;
    let i = this.ids.get(si);
    if (i === undefined) {
      i = this.list.length;
      this.list.push(si);
      this.ids.set(si, i);
    }
    return i;
  }

  xml(): string {
    return `${XML_HEAD}<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="${this.count}" uniqueCount="${this.list.length}">` +
      this.list.map((s) => `<si>${s}</si>`).join("") + "</sst>";
  }
}

const NEW_SHEET = `${XML_HEAD}<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac"><dimension ref="A1"/><sheetViews><sheetView workbookViewId="0"/></sheetViews><sheetFormatPr baseColWidth="10" defaultRowHeight="15" x14ac:dyDescent="0.25"/><sheetData/><pageMargins left="0.7" right="0.7" top="0.78740157499999996" bottom="0.78740157499999996" header="0.3" footer="0.3"/></worksheet>`;

class SheetWriter {
  formulaWithoutValue = false;
  private sp: Split;
  private readonly rels: Rels;

  constructor(
    private s: Sheet, private selected: boolean, private own: SheetRaw | undefined, private src: SheetRaw | undefined, private part: string,
    private pw: PackageWriter, private sw: StylesWriter, private sst: SharedStrings, private book: BookRaw, private author: string, private vmlIds: number[],
  ) {
    this.sp = src ? { ...src.xml, parts: src.xml.parts.map((p) => ({ ...p })) } : split(NEW_SHEET);
    if (src && !own) {
      // A copy of another sheet: no parts of its own yet.
      this.sp.parts = this.sp.parts.filter((p) => !PART_REFS.has(p.name));
      for (const p of this.sp.parts) if (p.name === "pageSetup") p.xml = setAttrs(p.xml, { "r:id": null });
    }
    this.rels = pw.relsOf(part);
  }

  private get(name: string): string | undefined {
    return this.sp.parts.find((p) => p.name === name)?.xml;
  }

  private put(name: string, xml: string | null): void {
    place(this.sp, name, xml, SHEET_ORDER);
  }

  xml(): string {
    const s = this.s;
    const snap = this.own ? this.own.snap : undefined;
    const now = sheetSnap(s);
    // Tab colour.
    if (!snap || now.tab !== snap.tab) {
      let pr = this.get("sheetPr") ?? "<sheetPr/>";
      pr = pr.replace(/<tabColor\b[^>]*\/>/, "");
      if (s.tabColor) {
        const tab = colorXml("tabColor", s.tabColor);
        pr = pr.endsWith("/>") ? `${pr.slice(0, -2).replace(/\s+$/, "")}>${tab}</sheetPr>` : pr.replace(/^(<sheetPr\b[^>]*>)/, `$1${tab}`);
      }
      this.put("sheetPr", /^<sheetPr\s*\/>$|^<sheetPr>\s*<\/sheetPr>$/.test(pr) ? null : pr);
    }
    const used = s.usedRange();
    this.put("dimension", `<dimension ref="${used ? rangeName(used) : "A1"}"/>`);
    // View.
    let views = this.get("sheetViews");
    if (!views || !snap || now.view !== snap.view) views = viewsXml(s.view, views, this.src);
    views = views.replace(/<sheetView\b[^>]*>/, (tag) => {
      const on = /\btabSelected="(1|true)"/.test(tag);
      return on === this.selected ? tag : setAttrs(tag, { tabSelected: this.selected ? "1" : null });
    });
    this.put("sheetViews", views);
    // Default sizes.
    if (!snap || now.defaults !== snap.defaults) {
      let fmt = this.get("sheetFormatPr") ?? '<sheetFormatPr baseColWidth="10" defaultRowHeight="15"/>';
      const base = Number(/baseColWidth="([^"]*)"/.exec(fmt)?.[1] ?? 8);
      const dw = s.defaultColWidth;
      const fromBase = dw === undefined || Math.abs(dw - widthFromBase(base, this.book.mdw)) < 1e-9;
      fmt = setAttrs(fmt, {
        defaultColWidth: fromBase ? null : String(widthToFile(dw!, this.book.mdw)),
        defaultRowHeight: String(s.defaultRowHeight),
        customHeight: Math.abs(s.defaultRowHeight - 15) > 1e-9 && /customHeight/.test(fmt) ? "1" : undefined,
      });
      this.put("sheetFormatPr", fmt);
    }
    // Gliederung: Excel shows the outline bar from these two levels.
    {
      const lv = (m: Map<number, { outline?: number }>) => Math.max(0, ...[...m.values()].map((i) => i.outline ?? 0));
      const fmt = this.get("sheetFormatPr") ?? '<sheetFormatPr baseColWidth="10" defaultRowHeight="15"/>';
      const want = { outlineLevelRow: lv(s.rows) ? String(lv(s.rows)) : null, outlineLevelCol: lv(s.cols) ? String(lv(s.cols)) : null };
      const next = setAttrs(fmt, want);
      if (next !== fmt) this.put("sheetFormatPr", next);
    }
    // Columns.
    if (!snap || now.cols !== snap.cols) this.put("cols", this.colsXml());
    this.put("sheetData", this.dataXml());
    if (!snap || now.merges !== snap.merges) {
      this.put("mergeCells", s.merges.length ? `<mergeCells count="${s.merges.length}">${s.merges.map((m) => `<mergeCell ref="${rangeName(m)}"/>`).join("")}</mergeCells>` : null);
    }
    if (!snap || now.links !== snap.links) this.hyperlinks();
    if (!snap || now.notes !== snap.notes) this.notes();
    if (!snap || now.filter !== snap.filter) {
      const af = s.autoFilter;
      this.put("autoFilter", af ? autoFilterXml(af, this.get("autoFilter"), this.sw) : null);
      // sheetPr filterMode: whether the filter hides rows.
      const filtering = !!af && (Object.keys(af.columns).length > 0 || /<filterColumn\b/.test(af.rawXml ?? ""));
      const pr = this.get("sheetPr") ?? "<sheetPr/>";
      const next = pr.replace(/^<sheetPr\b[^>]*?\/?>/, (t) => setAttrs(t, { filterMode: filtering ? "1" : null }));
      this.put("sheetPr", /^<sheetPr\s*\/>$|^<sheetPr>\s*<\/sheetPr>$/.test(next) ? null : next);
    }
    const drawing = writeCharts(this.pw, s, this.part, this.own?.drawing, this.get("drawing"));
    if (drawing !== undefined) this.put("drawing", drawing);
    if (!snap || now.cf !== snap.cf) placeAll(this.sp, "conditionalFormatting", conditionalXml(s.conditional ?? [], this.own?.cf, this.sw), SHEET_ORDER);
    if (!snap || now.dv !== snap.dv) {
      const dv = validationsXml(s.validations ?? [], this.own?.dv, this.get("dataValidations"), this.get("extLst"));
      this.put("dataValidations", dv.main);
      this.put("extLst", dv.ext);
    }
    if (!snap || now.prot !== snap.prot) {
      const p = s.protection;
      this.put("sheetProtection", p && Object.keys(p).length ? `<sheetProtection${Object.entries(p).map(([k, v]) => ` ${k}="${esc(v)}"`).join("")}/>` : null);
    }
    if (!snap || now.page !== snap.page) {
      const pg = pageXml(s.page, this.sp.parts);
      this.put("pageMargins", pg.margins);
      this.put("pageSetup", pg.setup);
      this.put("printOptions", pg.options);
      const same = (a?: number[], b?: number[]) => JSON.stringify(a ?? []) === JSON.stringify(b ?? []);
      const old = this.own ? readPage(this.sp.parts) : undefined;
      if (!this.own || !same(old?.rowBreaks, s.page?.rowBreaks)) this.put("rowBreaks", breaksXml(s.page?.rowBreaks, true));
      if (!this.own || !same(old?.colBreaks, s.page?.colBreaks)) this.put("colBreaks", breaksXml(s.page?.colBreaks, false));
      if ((s.page?.header ?? "") !== (this.own ? readHf(this.get("headerFooter"), "oddHeader") : "") || (s.page?.footer ?? "") !== (this.own ? readHf(this.get("headerFooter"), "oddFooter") : "")) {
        this.put("headerFooter", headerFooterXml(s.page, this.get("headerFooter")));
      }
      this.put("sheetPr", fitToPage(this.get("sheetPr"), pg.fit));
    }
    return joinSplit(this.sp);
  }

  private colsXml(): string | null {
    const cols = [...this.s.cols].filter(([, c]) => Object.keys(c).length).sort((a, b) => a[0] - b[0]);
    if (!cols.length) return null;
    const out: string[] = [];
    let i = 0;
    while (i < cols.length) {
      const [lo, info] = cols[i];
      const key = canon(info);
      let j = i;
      while (j + 1 < cols.length && cols[j + 1][0] === cols[j][0] + 1 && canon(cols[j + 1][1]) === key) j++;
      out.push(this.colXml(lo, cols[j][0], info));
      i = j + 1;
    }
    return `<cols>${out.join("")}</cols>`;
  }

  private colXml(lo: number, hi: number, c: ColInfo): string {
    const mdw = this.book.mdw;
    const width = c.width !== undefined ? widthToFile(c.width, mdw) : c.hidden ? 0 : widthToFile(this.s.defaultColWidth ?? widthFromBase(10, mdw), mdw);
    const xf = c.s ? this.sw.xf(c.s) : 0;
    return `<col min="${lo + 1}" max="${hi + 1}" width="${width}"${xf ? ` style="${xf}"` : ""}${c.hidden ? ' hidden="1"' : ""}${c.bestFit ? ' bestFit="1"' : ""}` +
      `${c.width !== undefined || c.hidden ? ' customWidth="1"' : ""}${c.outline ? ` outlineLevel="${c.outline}"` : ""}${c.collapsed ? ' collapsed="1"' : ""}/>`;
  }

  private dataXml(): string {
    const s = this.s;
    const keys = [...s.cells.keys()].sort((a, b) => a - b);
    const rows = [...s.rows.keys()].sort((a, b) => a - b);
    const out: string[] = ["<sheetData>"];
    let ki = 0, ri = 0;
    while (ki < keys.length || ri < rows.length) {
      const kr = ki < keys.length ? keyCell(keys[ki]).row : Infinity;
      const rr = ri < rows.length ? rows[ri] : Infinity;
      const row = Math.min(kr, rr);
      if (rr === row) ri++;
      const cells: string[] = [];
      while (ki < keys.length && keyCell(keys[ki]).row === row) {
        const k = keys[ki++];
        const x = this.cellXml(s.cells.get(k)!, k);
        if (x) cells.push(x);
      }
      const info = s.rows.get(row);
      if (!cells.length && !info) continue;
      const tag = rowTag(row, info, info?.s ? this.sw.xf(info.s) : 0);
      out.push(cells.length ? `${tag}${cells.join("")}</row>` : tag.replace(/>$/, "/>"));
    }
    out.push("</sheetData>");
    return out.length === 2 ? "<sheetData/>" : out.join("");
  }

  private cellXml(c: Cell, key: number): string {
    const { row, col } = keyCell(key);
    const xf = c.s ? this.sw.xf(c.s) : 0;
    let t = "";
    let body = "";
    const v = c.v;
    if (c.raw?.f) body += c.raw.f;
    else if (c.f !== undefined) {
      const file = c.raw?.file !== undefined && fromXlsxFormula(c.raw.file) === c.f ? c.raw.file : toXlsxFormula(c.f);
      body += c.array ? `<f t="array" ref="${rangeName(c.array)}">${escText(file)}</f>` : `<f>${escText(file)}</f>`;
      if (v === null) this.formulaWithoutValue = true;
    }
    const formula = c.f !== undefined || !!c.raw?.f;
    if (typeof v === "number") body += Number.isFinite(v) ? `<v>${numText(v)}</v>` : ((t = "e"), "<v>#NUM!</v>");
    else if (typeof v === "string") {
      if (formula) { t = "str"; body += `<v>${escText(encodeX(v))}</v>`; }
      else {
        t = "s";
        const rich = c.richXml && richMatches(c.richXml, v) ? c.richXml : tXml(v);
        body += `<v>${this.sst.id(rich)}</v>`;
      }
    } else if (typeof v === "boolean") { t = "b"; body += `<v>${v ? 1 : 0}</v>`; }
    else if (v instanceof CellError) { t = "e"; body += `<v>${esc(v.code)}</v>`; }
    if (!body && !xf && !c.raw?.attrs) return "";
    const head = `<c r="${cellName(row, col)}"${xf ? ` s="${xf}"` : ""}${t ? ` t="${t}"` : ""}${c.raw?.attrs ?? ""}`;
    return body ? `${head}>${body}</c>` : `${head}/>`;
  }

  private hyperlinks(): void {
    const own = this.own;
    const out: string[] = [];
    const keys = [...this.s.cells].filter(([, c]) => c.link !== undefined).sort((a, b) => a[0] - b[0]);
    for (const [k, c] of keys) {
      const { row, col } = keyCell(k);
      const ref = cellName(row, col);
      const old = own?.hyperlinks.find((h) => h.link === c.link && h.ref.split(":")[0] === ref);
      if (old) { out.push(old.xml); continue; }
      const link = c.link!;
      if (link.startsWith("#")) out.push(`<hyperlink ref="${ref}" location="${esc(link.slice(1))}" display="${esc(display(c))}"/>`);
      else out.push(`<hyperlink ref="${ref}" r:id="${this.rels.ensure(RT.hyperlink, link, true)}"/>`);
    }
    this.put("hyperlinks", out.length ? `<hyperlinks>${out.join("")}</hyperlinks>` : null);
    // Relationships of links that are gone.
    const ids = new Set([...out.join("").matchAll(/r:id="([^"]+)"/g)].map((m) => m[1]));
    for (const r of [...this.rels.list]) if (isRel(r.type, RT.hyperlink) && !ids.has(r.id)) removeRel(this.rels, r.id);
  }

  private notes(): void {
    const notes: NoteIn[] = [...this.s.cells].filter(([, c]) => c.note !== undefined).sort((a, b) => a[0] - b[0])
      .map(([k, c]) => ({ ...keyCell(k), text: c.note! }));
    const orig = this.own?.notes;
    const cRel = this.rels.list.find((r) => isRel(r.type, RT.comments));
    const vRel = this.rels.list.find((r) => isRel(r.type, RT.vmlDrawing));
    const vmlPart = vRel ? resolveTarget(this.part, vRel.target) : undefined;
    const vmlOld = vmlPart && this.pw.files[vmlPart] ? strFromU8(this.pw.files[vmlPart]) : null;
    if (!notes.length) {
      if (cRel) removeRel(this.rels, cRel.id);
      const vml = vmlOld ? vmlXml([], vmlOld, 1) : null;
      if (vRel && vml === null) {
        removeRel(this.rels, vRel.id);
        this.put("legacyDrawing", null);
      } else if (vml && vmlPart) this.pw.setPart(vmlPart, vml);
      return;
    }
    const cPart = cRel ? resolveTarget(this.part, cRel.target) : this.pw.freeName("xl/comments");
    this.pw.setPart(cPart, commentsXml(notes, this.author, orig && orig.part === cPart ? orig : undefined), CT.comments);
    if (!cRel) this.pw.link(this.part, RT.comments, cPart);
    let block = 1;
    while (this.vmlIds.includes(block)) block++;
    const vml = vmlXml(notes, vmlOld, vmlOld ? 1 : block)!;
    if (!vmlOld) this.vmlIds.push(block);
    const vPart = vmlPart ?? this.pw.freeName("xl/drawings/vmlDrawing", "vml");
    this.pw.setPart(vPart, vml, vmlPart ? undefined : CT.vml);
    const vid = vRel?.id ?? this.pw.link(this.part, RT.vmlDrawing, vPart);
    if (!this.get("legacyDrawing")) this.put("legacyDrawing", `<legacyDrawing r:id="${vid}"/>`);
  }
}

function display(c: Cell): string {
  return typeof c.v === "string" ? c.v : c.link!.replace(/^#/, "");
}

/** Whether rich text XML still spells the cell's text (it is dropped once the text is edited). */
function richMatches(xml: string, text: string): boolean {
  let out = "";
  for (const m of xml.matchAll(/<(r|t|rPh)\b[^>]*?(?:\/>|>([\s\S]*?)<\/\1>)/g)) {
    if (m[1] === "rPh") continue;
    if (m[1] === "t") out += m[2] ?? "";
    else {
      const t = /<t\b[^>]*?(?:\/>|>([\s\S]*?)<\/t>)/.exec(m[2] ?? "");
      out += t?.[1] ?? "";
    }
  }
  const dec = out.replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&#x([0-9a-fA-F]+);/g, (_, h) => String.fromCodePoint(parseInt(h, 16)))
    .replace(/&#(\d+);/g, (_, d) => String.fromCodePoint(Number(d))).replace(/&amp;/g, "&").replace(/_x([0-9A-Fa-f]{4})_/g, (_, h: string) => String.fromCharCode(parseInt(h, 16)));
  return dec === text;
}

function rowTag(row: number, info: RowInfo | undefined, xf: number): string {
  let a = ` r="${row + 1}"`;
  if (xf) a += ` s="${xf}" customFormat="1"`;
  if (info?.height !== undefined) a += ` ht="${numText(info.height)}"`;
  if (info?.hidden) a += ' hidden="1"';
  if (info?.custom) a += ' customHeight="1"';
  if (info?.outline) a += ` outlineLevel="${info.outline}"`;
  if (info?.collapsed) a += ' collapsed="1"';
  return `<row${a}${info?.raw ?? ""}>`;
}

const PANES = { both: "bottomRight", rows: "bottomLeft", cols: "topRight" } as const;

/** <sheetViews> for a view: the original's attributes and extra children are kept. */
function viewsXml(v: View, old: string | undefined, raw: SheetRaw | undefined): string {
  const oldTag = old ? /<sheetView\b[^>]*?\/?>/.exec(old)?.[0] : undefined;
  let tag = oldTag ? oldTag.replace(/\/>$/, ">") : '<sheetView workbookViewId="0">';
  const tl = v.freeze ? raw?.frozenTopLeft ?? { row: 0, col: 0 } : { row: v.top, col: v.left };
  tag = setAttrs(tag, {
    showGridLines: v.showGrid ? null : "0", showRowColHeaders: v.showHeaders ? null : "0",
    topLeftCell: tl.row || tl.col ? cellName(tl.row, tl.col) : null,
    zoomScale: v.zoom !== 100 ? String(v.zoom) : null, zoomScaleNormal: v.zoom !== 100 ? String(v.zoom) : null,
  });
  // Children other than pane and selection (pivotSelection, extLst) stay.
  const inner = old && oldTag && !oldTag.endsWith("/>") ? /<sheetView\b[^>]*>([\s\S]*?)<\/sheetView>/.exec(old)?.[1] ?? "" : "";
  const rest = inner.replace(/<pane\b[^>]*\/>/g, "").replace(/<selection\b[^>]*\/>/g, "");
  const oldPane = /<pane\b[^>]*\/>/.exec(inner)?.[0];
  let pane = "";
  let active = "topLeft";
  if (v.freeze && (v.freeze.row > tl.row || v.freeze.col > tl.col)) {
    const ys = v.freeze.row - tl.row, xs = v.freeze.col - tl.col;
    active = ys && xs ? PANES.both : ys ? PANES.rows : PANES.cols;
    const top = Math.max(v.top, v.freeze.row), left = Math.max(v.left, v.freeze.col);
    pane = `<pane${xs ? ` xSplit="${xs}"` : ""}${ys ? ` ySplit="${ys}"` : ""} topLeftCell="${cellName(top, left)}" activePane="${active}" state="frozen"/>`;
  } else if (oldPane && !/state="frozen/.test(oldPane)) {
    // A split window (not modelled) stays as it was.
    pane = oldPane;
    active = /activePane="([^"]+)"/.exec(oldPane)?.[1] ?? "topLeft";
  }
  const sqref = v.selection.map(rangeName).join(" ") || cellName(v.active.row, v.active.col);
  const ac = cellName(v.active.row, v.active.col);
  const sel = `<selection${active !== "topLeft" ? ` pane="${active}"` : ""}${ac === "A1" && sqref === "A1" && !pane ? "" : ` activeCell="${ac}" sqref="${sqref}"`}/>`;
  const body = pane + (sel === "<selection/>" ? "" : sel) + rest;
  const head = old ? old.slice(0, old.indexOf(">") + 1) : "<sheetViews>";
  const others = old ? (/<\/sheetView>([\s\S]*)<\/sheetViews>/.exec(old)?.[1] ?? /<sheetView\b[^>]*\/>([\s\S]*)<\/sheetViews>/.exec(old)?.[1] ?? "") : "";
  return `${head}${body ? `${tag}${body}</sheetView>` : tag.replace(/>$/, "/>")}${others}</sheetViews>`;
}

function workbookXml(wb: Workbook, raw: BookRaw, entries: { name: string; sheetId: number; state?: string; rid: string; attrs: Record<string, string> }[], needsCalc: boolean): string {
  const sp: Split = { ...raw.wb, parts: raw.wb.parts.map((p) => ({ ...p })) };
  const get = (n: string) => sp.parts.find((p) => p.name === n)?.xml;
  const put = (n: string, xml: string | null) => place(sp, n, xml, BOOK_ORDER);
  // File order: worksheets with the other sheets at their old positions.
  const order: (number | { xml: string })[] = wb.sheets.map((_, i) => i);
  for (const o of [...raw.others].sort((a, b) => a.index - b.index)) order.splice(Math.min(o.index, order.length), 0, { xml: o.xml });
  const fileIndex = new Map<number, number>();
  order.forEach((x, i) => { if (typeof x === "number") fileIndex.set(x, i); });
  if (wb.date1904 !== raw.snap.date1904) put("workbookPr", setAttrs(get("workbookPr") ?? "<workbookPr/>", { date1904: wb.date1904 ? "1" : null }));
  const active = fileIndex.get(wb.active) ?? 0;
  const views = get("bookViews");
  if (views) {
    put("bookViews", views.replace(/<workbookView\b[^>]*?\/?>/, (tag) => {
      const first = Number(/firstSheet="(\d+)"/.exec(tag)?.[1] ?? 0);
      return setAttrs(tag, {
        activeTab: active ? String(active) : null,
        firstSheet: first > active || first >= order.length ? null : undefined,
      });
    }));
  } else put("bookViews", `<bookViews><workbookView${active ? ` activeTab="${active}"` : ""}/></bookViews>`);
  const sheetXml = order.map((x) => {
    if (typeof x !== "number") return x.xml;
    const e = entries[x];
    const extra = Object.entries(e.attrs).map(([k, v]) => ` ${k}="${esc(v)}"`).join("");
    return `<sheet name="${esc(e.name)}" sheetId="${e.sheetId}"${e.state ? ` state="${e.state}"` : ""} r:id="${e.rid}"${extra}/>`;
  });
  put("sheets", `<sheets>${sheetXml.join("")}</sheets>`);
  // Defined names, with each sheet's print area and AutoFilter range; Excel's text is kept
  // when the set of names did not change.
  const names = wb.names.map((n) => nameXml(n, raw, fileIndex));
  wb.sheets.forEach((s, i) => {
    const local = fileIndex.get(i);
    if (s.page?.printArea) names.push(`<definedName name="${PRINT_AREA}" localSheetId="${local}">${escText(sheetRef(s.name, s.page.printArea))}</definedName>`);
    if (s.page?.titleRows || s.page?.titleCols) names.push(`<definedName name="${PRINT_TITLES}" localSheetId="${local}">${escText(titlesRef(s.name, s.page.titleRows, s.page.titleCols))}</definedName>`);
    if (s.autoFilter) names.push(`<definedName name="${FILTER_DB}" localSheetId="${local}" hidden="1">${escText(sheetRef(s.name, s.autoFilter.ref))}</definedName>`);
  });
  for (const o of raw.otherNames) {
    const at = order.findIndex((x) => typeof x !== "number" && raw.others.find((q) => q.index === o.index)?.xml === x.xml);
    if (at >= 0) names.push(o.xml.replace(/localSheetId="\d+"/, `localSheetId="${at}"`));
  }
  if (canon([...names].sort()) !== canon([...raw.defined].sort())) {
    // Excel lists names alphabetically.
    names.sort((a, b) => nameSort(a).localeCompare(nameSort(b)));
    put("definedNames", names.length ? `<definedNames>${names.join("")}</definedNames>` : null);
  }
  if (needsCalc) put("calcPr", setAttrs(get("calcPr") ?? '<calcPr calcId="191029"/>', { fullCalcOnLoad: "1" }));
  return joinSplit(sp);
}

function nameXml(n: DefinedName, raw: BookRaw, fileIndex: Map<number, number>): string {
  const local = n.sheet !== undefined ? fileIndex.get(n.sheet) : undefined;
  return `<definedName name="${esc(n.name)}"${raw.nameAttrs[nameKey(n)] ?? ""}${local !== undefined ? ` localSheetId="${local}"` : ""}${n.hidden ? ' hidden="1"' : ""}>${escText(toXlsxFormula(n.ref))}</definedName>`;
}

const nameSort = (xml: string) => (/name="([^"]*)"/.exec(xml)?.[1] ?? "").replace(/^_xlnm\./, "").toLowerCase();

/** Document properties: modification time and last author; sheet titles go when they changed. */
function props(pw: PackageWriter, wb: Workbook, raw: BookRaw, o: XlsxWriteOptions): void {
  const rels = new Rels(pw.files["_rels/.rels"] ? strFromU8(pw.files["_rels/.rels"]) : null);
  const core = rels.list.find((r) => isRel(r.type, RT.coreProps));
  const corePath = core ? resolveTarget("", core.target) : undefined;
  if (corePath && pw.files[corePath]) {
    const t = (o.now ?? new Date()).toISOString().replace(/\.\d+Z$/, "Z");
    let xml = strFromU8(pw.files[corePath]).replace(/(<dcterms:modified[^>]*>)[^<]*(<\/dcterms:modified>)/, `$1${t}$2`);
    if (o.author) xml = xml.replace(/(<cp:lastModifiedBy>)[^<]*(<\/cp:lastModifiedBy>)/, `$1${esc(o.author)}$2`);
    pw.files[corePath] = strToU8(xml);
  }
  const app = rels.list.find((r) => isRel(r.type, RT.extProps));
  const appPath = app ? resolveTarget("", app.target) : undefined;
  if (appPath && pw.files[appPath] && sheetNamesSnap(wb) !== raw.snap.sheetNames) {
    const xml = strFromU8(pw.files[appPath]).replace(/<HeadingPairs>[\s\S]*?<\/HeadingPairs>/, "").replace(/<TitlesOfParts>[\s\S]*?<\/TitlesOfParts>/, "");
    pw.files[appPath] = strToU8(xml);
  }
}


/** The text of one part of <headerFooter> as the model holds it. */
function readHf(xml: string | undefined, name: string): string {
  const m = xml ? new RegExp(`<${name}>([\\s\\S]*?)</${name}>`).exec(xml) : null;
  return m ? m[1].replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, "&") : "";
}
