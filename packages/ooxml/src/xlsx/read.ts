// Reading .xlsx into RaabeCalc's workbook model. Values, formulas (shared formulas expanded
// per cell), formats, sizes, merges, views, names, notes and hyperlinks are modelled; every
// other part and element is kept as Excel wrote it (D-024) in the `raw` fields, with a
// snapshot of what was read, so the writer can tell what the user changed.

import { strFromU8, unzipSync } from "fflate";
import {
  Sheet, Workbook, cellKey, err, fromXlsxFormula, parseCell, parseRange, toXlsxFormula,
  type Cell, type ColInfo, type DefinedName, type ErrorCode, type Rect, type RowInfo, type View,
} from "@raabe/sheet-engine";
import { Rels, relsPath, resolveTarget } from "../package";
import { readTheme, resolveColor, type ColorContext } from "./colors";
import { shiftFormula } from "./shift";
import { readStyles, canon, type StylesRaw } from "./styles";
import { MDW, widthFromBase, widthFromFile } from "./units";
import { readCharts, readPictures, type DrawingRaw } from "./drawing";
import { FILTER_DB, PRINT_AREA, PRINT_TITLES, ownRange, readTitles, readAutoFilter, readConditional, readPage, readValidations, type CfRaw, type DvRaw } from "./sheetparts";
import { attrs, child, children, decodeX, element, outer, root, split, text, unesc, type El, type Split } from "./xml";

/** Whether a relationship type is `want`, also in Strict OOXML (same last segment). */
export function isRel(type: string, want: string): boolean {
  return type === want || type.slice(type.lastIndexOf("/")) === want.slice(want.lastIndexOf("/"));
}

export const RT = {
  officeDocument: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument",
  worksheet: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet",
  styles: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles",
  theme: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme",
  sharedStrings: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings",
  calcChain: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain",
  comments: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments",
  vmlDrawing: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing",
  hyperlink: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
  coreProps: "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties",
  extProps: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties",
};

/** A sheet of the workbook that is not a worksheet (chart sheet, dialog sheet …): kept as is. */
export interface OtherSheet {
  /** Position among all sheets of the file. */
  index: number;
  /** The <sheet> element in workbook.xml. */
  xml: string;
}

/** The workbook's package and what the model does not hold (Workbook.raw). */
export interface BookRaw {
  files: Record<string, Uint8Array>;
  wbPath: string;
  wb: Split;
  styles: StylesRaw;
  theme: { major: string; minor: string };
  sstPath?: string;
  others: OtherSheet[];
  /** Attributes of each defined name the model does not hold, by name and scope. */
  nameAttrs: Record<string, string>;
  /** The file's <definedName> elements, to keep them as written when nothing changed. */
  defined: string[];
  /** Names scoped to non-worksheets ("<definedName …>" as written), by OtherSheet index. */
  otherNames: { index: number; xml: string }[];
  mdw: number;
  /** What was read, to see changes. */
  snap: { names: string; date1904: boolean; active: number; sheetNames: string; prot: string };
}

/** A worksheet's part and what the model does not hold (Sheet.raw). */
export interface SheetRaw {
  part: string;
  sheetId: number;
  /** The <sheet> element's attributes the model does not hold (e.g. xr:uid). */
  attrs: Record<string, string>;
  xml: Split;
  /** Top-left cell of the frozen part (sheetView topLeftCell) when panes are frozen. */
  frozenTopLeft?: { row: number; col: number };
  /** Base column width (sheetFormatPr baseColWidth). */
  baseColWidth?: number;
  hyperlinks: { ref: string; xml: string; link: string }[];
  notes?: { part: string; vml?: string; xml: string; byRef: Record<string, { text: string; xml: string }> };
  /** Charts of the sheet's drawing as read. */
  drawing?: DrawingRaw;
  /** Conditional formats as read. */
  cf?: CfRaw;
  /** Data validations as read. */
  dv?: DvRaw;
  /** The pictures as read (picturesKey), to see whether they changed. */
  pics?: string;
  snap: { view: string; cols: string; merges: string; tab: string; defaults: string; notes: string; links: string; filter: string; cf: string; page: string; dv: string; prot: string };
}

const ERRORS = new Set(["#NULL!", "#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#NUM!", "#N/A", "#SPILL!", "#CALC!", "#GETTING_DATA"]);

export function readXlsx(bytes: Uint8Array): Workbook {
  return readPackage(unzipSync(bytes));
}

export function readPackage(files: Record<string, Uint8Array>): Workbook {
  const str = (p: string | undefined) => (p && files[p] ? strFromU8(files[p]) : null);
  const rootRels = new Rels(str("_rels/.rels"));
  const main = rootRels.list.find((r) => isRel(r.type, RT.officeDocument));
  const wbPath = main ? resolveTarget("", main.target) : "xl/workbook.xml";
  const wbXml = str(wbPath);
  if (!wbXml) throw new Error("Keine Excel-Arbeitsmappe: xl/workbook.xml fehlt.");
  const rels = new Rels(str(relsPath(wbPath)));
  const partOf = (type: string) => {
    const r = rels.list.find((x) => isRel(x.type, type));
    return r ? resolveTarget(wbPath, r.target) : undefined;
  };

  const wb = new Workbook();
  const theme = readTheme(str(partOf(RT.theme)));
  wb.theme = Object.fromEntries(Object.entries(theme.colors).map(([k, v]) => [k, "#" + v]));
  const { table, raw: stylesRaw } = readStyles(str(partOf(RT.styles)) || NO_STYLES, theme.colors);
  wb.styles = table;
  const ctx = stylesRaw.ctx;
  const sstPath = partOf(RT.sharedStrings);
  const sst = readSst(str(sstPath));

  const wbSplit = split(wbXml);
  const wbRoot = root(wbXml);
  const parts = children(wbXml, wbRoot.open, wbRoot.close);
  const find = (name: string) => parts.find((p) => p.name === name);
  const pr = find("workbookPr");
  wb.date1904 = pr ? /^(1|true)$/.test(attrs(wbXml, pr).date1904 ?? "") : false;
  const wp = find("workbookProtection");
  if (wp) wb.protection = attrs(wbXml, wp);
  const sheetsEl = find("sheets");
  const entries = sheetsEl ? children(wbXml, sheetsEl.open, sheetsEl.close) : [];
  const others: OtherSheet[] = [];
  const fileToModel = new Map<number, number>();
  entries.forEach((e, index) => {
    const a = attrs(wbXml, e);
    const rid = a["r:id"] ?? "";
    const rel = rels.get(rid);
    const part = rel ? resolveTarget(wbPath, rel.target) : "";
    const xml = str(part);
    if (!rel || !isRel(rel.type, RT.worksheet) || xml === null) {
      others.push({ index, xml: outer(wbXml, e) });
      return;
    }
    const sheet = readSheet(xml, a.name ?? `Tabelle${index + 1}`, { files, part, sst, ctx, date1904: wb.date1904, styles: wb.styles.list.length, dxfs: stylesRaw.dxfs });
    const keep: Record<string, string> = {};
    for (const [k, v] of Object.entries(a)) if (!/^(name|sheetId|state|r:id)$/.test(k)) keep[k] = v;
    (sheet.raw as SheetRaw).sheetId = Number(a.sheetId ?? index + 1);
    (sheet.raw as SheetRaw).attrs = keep;
    if (a.state === "hidden") sheet.hidden = true;
    else if (a.state === "veryHidden") sheet.hidden = "very";
    fileToModel.set(index, wb.sheets.length);
    wb.sheets.push(sheet);
  });
  if (!wb.sheets.length) wb.sheets.push(new Sheet("Tabelle1"));

  const view = find("bookViews") && children(wbXml, find("bookViews")!.open, find("bookViews")!.close)[0];
  const activeTab = view ? Number(attrs(wbXml, view).activeTab ?? 0) : 0;
  wb.active = fileToModel.get(activeTab) ?? 0;

  // Defined names.
  const nameAttrs: Record<string, string> = {};
  const defined: string[] = [];
  const otherNames: BookRaw["otherNames"] = [];
  const namesEl = find("definedNames");
  for (const n of namesEl ? children(wbXml, namesEl.open, namesEl.close) : []) {
    const a = attrs(wbXml, n);
    const dn: DefinedName = { name: a.name ?? "", ref: fromXlsxFormula(unesc(text(wbXml, n))) };
    if (a.localSheetId !== undefined) {
      const m = fileToModel.get(Number(a.localSheetId));
      if (m === undefined) {
        otherNames.push({ index: Number(a.localSheetId), xml: outer(wbXml, n) });
        continue;
      }
      dn.sheet = m;
    }
    if (a.hidden === "1" || a.hidden === "true") dn.hidden = true;
    defined.push(outer(wbXml, n));
    // The print area and the AutoFilter's range belong to the sheet's model.
    if (dn.sheet !== undefined && (dn.name === PRINT_AREA || dn.name === FILTER_DB || dn.name === PRINT_TITLES) && Object.keys(a).every((k) => /^(name|localSheetId|hidden)$/.test(k))) {
      const sheet = wb.sheets[dn.sheet];
      if (dn.name === PRINT_TITLES) {
        const t = readTitles(dn.ref, sheet.name);
        if (t) {
          sheet.page = { ...sheet.page, titleRows: t.rows, titleCols: t.cols };
          continue;
        }
      }
      const rect = ownRange(dn.ref, sheet.name);
      if (rect && dn.name === PRINT_AREA) {
        sheet.page = { ...sheet.page, printArea: rect };
        continue;
      }
      if (rect && dn.name === FILTER_DB && sheet.autoFilter && canon(rect) === canon(sheet.autoFilter.ref)) continue;
    }
    const extra = Object.entries(a).filter(([k]) => !/^(name|localSheetId|hidden)$/.test(k)).map(([k, v]) => ` ${k}="${v.replace(/&/g, "&amp;").replace(/"/g, "&quot;").replace(/</g, "&lt;")}"`).join("");
    if (extra) nameAttrs[nameKey(dn)] = extra;
    wb.names.push(dn);
  }

  const raw: BookRaw = {
    files, wbPath, wb: wbSplit, styles: stylesRaw, theme: { major: theme.major, minor: theme.minor }, sstPath, others, nameAttrs, otherNames, defined, mdw: MDW,
    snap: { names: canon(wb.names), date1904: wb.date1904, active: wb.active, sheetNames: "", prot: canon([wb.protection ?? null]) },
  };
  raw.snap.sheetNames = sheetNamesSnap(wb);
  // Snapshots again: print areas came from the names.
  for (const s of wb.sheets) if (s.raw) (s.raw as SheetRaw).snap = sheetSnap(s);
  wb.raw = raw;
  return wb;
}

export function nameKey(n: DefinedName): string {
  return `${n.name}\u0001${n.sheet ?? ""}`;
}

export function sheetNamesSnap(wb: Workbook): string {
  return JSON.stringify(wb.sheets.map((s) => [s.name, s.hidden ?? false, (s.raw as SheetRaw | undefined)?.part ?? ""]));
}

const NO_STYLES = `<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs></styleSheet>`;

interface SstItem { text: string; rich?: string }

/** Text of an <si>/<is>/<text>: its own <t> and those of its runs, without phonetic runs. */
export function richText(xml: string, el: El): string {
  let out = "";
  for (const c of children(xml, el.open, el.close)) {
    if (c.name === "t") out += decodeX(unesc(text(xml, c)));
    else if (c.name === "r") {
      const t = child(xml, c, "t");
      if (t) out += decodeX(unesc(text(xml, t)));
    }
  }
  return out;
}

function readSst(xml: string | null): SstItem[] {
  if (!xml) return [];
  const r = root(xml);
  return children(xml, r.open, r.close).filter((e) => e.name === "si").map((si) => {
    const inner = text(xml, si);
    const item: SstItem = { text: richText(xml, si) };
    if (/<(r|rPh|phoneticPr)\b/.test(inner)) item.rich = inner;
    return item;
  });
}

interface SheetContext {
  files: Record<string, Uint8Array>;
  part: string;
  sst: SstItem[];
  ctx: ColorContext;
  date1904: boolean;
  styles: number;
  dxfs: string[];
  /** The sheet's default x14ac:dyDescent (rows with it need not keep their own). */
  dyDescent?: string;
}

const style = (v: string | undefined, n: number) => {
  const s = Number(v ?? 0);
  return s > 0 && s < n ? s : undefined;
};

function readSheet(xml: string, name: string, c: SheetContext): Sheet {
  const sheet = new Sheet(name);
  const sp = split(xml);
  const r = root(xml);
  const els = children(xml, r.open, r.close);
  const get = (n: string) => els.find((e) => e.name === n);
  const rels = new Rels(c.files[relsPath(c.part)] ? strFromU8(c.files[relsPath(c.part)]) : null);
  const raw: SheetRaw = { part: c.part, sheetId: 0, attrs: {}, xml: sp, hyperlinks: [], snap: { view: "", cols: "", merges: "", tab: "", defaults: "", notes: "", links: "", filter: "", cf: "", page: "", dv: "", prot: "" } };

  const prEl = get("sheetPr");
  const tab = prEl && child(xml, prEl, "tabColor");
  if (tab) sheet.tabColor = resolveColor(attrs(xml, tab), c.ctx);

  // Default sizes.
  const fmt = get("sheetFormatPr");
  const fa = fmt ? attrs(xml, fmt) : {};
  if (fa.defaultRowHeight) sheet.defaultRowHeight = Number(fa.defaultRowHeight);
  if (fa.baseColWidth) raw.baseColWidth = Number(fa.baseColWidth);
  sheet.defaultColWidth = fa.defaultColWidth ? widthFromFile(Number(fa.defaultColWidth)) : widthFromBase(Number(fa.baseColWidth ?? 8));

  // Columns.
  const colsEl = get("cols");
  for (const col of colsEl ? children(xml, colsEl.open, colsEl.close) : []) {
    const a = attrs(xml, col);
    const info: ColInfo = {};
    if (a.width !== undefined && (a.customWidth === "1" || a.customWidth === "true" || a.hidden !== "1")) info.width = widthFromFile(Number(a.width));
    if (a.hidden === "1" || a.hidden === "true") info.hidden = true;
    const s = style(a.style, c.styles);
    if (s) info.s = s;
    if (a.outlineLevel && a.outlineLevel !== "0") info.outline = Number(a.outlineLevel);
    if (a.bestFit === "1" || a.bestFit === "true") info.bestFit = true;
    if (a.collapsed === "1" || a.collapsed === "true") info.collapsed = true;
    const lo = Number(a.min) - 1, hi = Math.min(Number(a.max) - 1, 16383);
    for (let i = lo; i <= hi; i++) sheet.cols.set(i, { ...info });
  }

  // Cells.
  const data = get("sheetData");
  if (data) readCells(xml, data, sheet, { ...c, dyDescent: fa["x14ac:dyDescent"] });

  const mc = get("mergeCells");
  for (const m of mc ? children(xml, mc.open, mc.close) : []) {
    const rect = parseRange(attrs(xml, m).ref ?? "");
    if (rect) sheet.merges.push(rect);
  }

  // Hyperlinks.
  const hl = get("hyperlinks");
  for (const h of hl ? children(xml, hl.open, hl.close) : []) {
    const a = attrs(xml, h);
    const rel = a["r:id"] ? rels.get(a["r:id"]) : undefined;
    const link = rel ? rel.target + (a.location ? "#" + a.location : "") : "#" + (a.location ?? "");
    const ref = a.ref ?? "";
    raw.hyperlinks.push({ ref, xml: outer(xml, h), link });
    const at = parseRange(ref);
    if (at) {
      const k = cellKey(at.r0, at.c0);
      const cell = sheet.cells.get(k) ?? { v: null };
      cell.link = link;
      sheet.cells.set(k, cell);
    }
  }

  // Notes (legacy comments).
  const cRel = rels.list.find((x) => isRel(x.type, RT.comments));
  if (cRel) {
    const part = resolveTarget(c.part, cRel.target);
    const cx = c.files[part] ? strFromU8(c.files[part]) : "";
    const vRel = rels.list.find((x) => isRel(x.type, RT.vmlDrawing));
    const notes: NonNullable<SheetRaw["notes"]> = { part, vml: vRel ? resolveTarget(c.part, vRel.target) : undefined, xml: cx, byRef: {} };
    const cr = cx ? root(cx) : undefined;
    const list = cr ? child(cx, cr, "commentList") : undefined;
    for (const cm of list ? children(cx, list.open, list.close) : []) {
      const ref = attrs(cx, cm).ref ?? "";
      const t = child(cx, cm, "text");
      const note = t ? richText(cx, t) : "";
      notes.byRef[ref] = { text: note, xml: outer(cx, cm) };
      const at = parseCell(ref);
      if (at) {
        const k = cellKey(at.row, at.col);
        const cell = sheet.cells.get(k) ?? { v: null };
        cell.note = note;
        sheet.cells.set(k, cell);
      }
    }
    raw.notes = notes;
  }

  const af = get("autoFilter");
  if (af) sheet.autoFilter = readAutoFilter(xml, af, c.dxfs, c.ctx);
  // Blatt schützen: the attributes as they are (sheet="1" when on, the password hash, what is allowed).
  const prot = get("sheetProtection");
  if (prot) sheet.protection = attrs(xml, prot);
  const cfs = els.filter((e) => e.name === "conditionalFormatting");
  if (cfs.length) {
    raw.cf = { xml: [], rules: [] };
    sheet.conditional = readConditional(xml, cfs, c.ctx, c.dxfs, raw.cf);
  }
  const dvs = get("dataValidations");
  const ext = get("extLst");
  if (dvs || (ext && text(xml, ext).includes("x14:dataValidation"))) {
    raw.dv = { items: [] };
    const list = readValidations(xml, dvs, ext, raw.dv);
    if (list.length) sheet.validations = list;
  }
  sheet.page = readPage(sp.parts);

  const charts = readCharts(c.files, sheet, c.part, rels);
  raw.pics = readPictures(c.files, sheet, c.part, rels);
  if (charts) raw.drawing = charts;

  // View.
  const sv = get("sheetViews");
  const v0 = sv ? children(xml, sv.open, sv.close)[0] : undefined;
  if (v0) readView(xml, v0, sheet, raw);

  sheet.raw = raw;
  raw.snap = sheetSnap(sheet);
  return sheet;
}

/** The modelled parts of a sheet that the writer compares with what was read. */
export function sheetSnap(s: Sheet): SheetRaw["snap"] {
  const notes: Record<string, string> = {};
  const links: Record<string, string> = {};
  for (const [k, cell] of s.cells) {
    if (cell.note !== undefined) notes[k] = cell.note;
    if (cell.link !== undefined) links[k] = cell.link;
  }
  return {
    view: canon(s.view), cols: canon([...s.cols].sort((a, b) => a[0] - b[0])), merges: canon(s.merges), tab: canon([s.tabColor ?? null]),
    defaults: canon([s.defaultColWidth ?? null, s.defaultRowHeight]), notes: canon(notes), links: canon(links),
    filter: canon([s.autoFilter ?? null]), cf: canon(s.conditional ?? []), page: canon([s.page ?? null]), dv: canon(s.validations ?? []), prot: canon([s.protection ?? null]),
  };
}

function readView(xml: string, el: El, sheet: Sheet, raw: SheetRaw): void {
  const a = attrs(xml, el);
  const kids = children(xml, el.open, el.close);
  const pane = kids.find((k) => k.name === "pane");
  const pa = pane ? attrs(xml, pane) : {};
  const tl = parseCell(a.topLeftCell ?? "A1") ?? { row: 0, col: 0 };
  const v: View = {
    active: { row: 0, col: 0 }, selection: [], top: tl.row, left: tl.col, zoom: Number(a.zoomScale ?? 100) || 100,
    showGrid: a.showGridLines !== "0" && a.showGridLines !== "false", showHeaders: a.showRowColHeaders !== "0" && a.showRowColHeaders !== "false",
  };
  if (pane && (pa.state === "frozen" || pa.state === "frozenSplit")) {
    const ys = Number(pa.ySplit ?? 0), xs = Number(pa.xSplit ?? 0);
    v.freeze = { row: tl.row + ys, col: tl.col + xs };
    raw.frozenTopLeft = tl;
    const ptl = parseCell(pa.topLeftCell ?? "") ?? { row: tl.row + ys, col: tl.col + xs };
    v.top = ptl.row;
    v.left = ptl.col;
  }
  const activePane = pa.activePane ?? "topLeft";
  const sels = kids.filter((k) => k.name === "selection").map((k) => attrs(xml, k));
  const sel = sels.find((s) => (s.pane ?? "topLeft") === activePane) ?? sels.find((s) => !s.pane) ?? sels[0];
  const ac = parseCell(sel?.activeCell ?? "A1") ?? { row: 0, col: 0 };
  v.active = ac;
  const ranges = (sel?.sqref ?? "").split(/\s+/).map((t) => parseRange(t)).filter((x): x is Rect => !!x);
  v.selection = ranges.length ? ranges : [{ r0: ac.row, c0: ac.col, r1: ac.row, c1: ac.col }];
  sheet.view = v;
}

function readCells(xml: string, data: El, sheet: Sheet, c: SheetContext): void {
  const shared = new Map<string, { f: string; row: number; col: number }>();
  let rowIndex = -1;
  let i = xml.indexOf("<", data.open);
  while (i >= 0 && i < data.close) {
    const rowEl = element(xml, i);
    i = xml.indexOf("<", rowEl.end);
    if (rowEl.name !== "row") continue;
    const ra = attrs(xml, rowEl);
    rowIndex = ra.r ? Number(ra.r) - 1 : rowIndex + 1;
    const info: RowInfo = {};
    if (ra.ht !== undefined) info.height = Number(ra.ht);
    if (ra.customHeight === "1" || ra.customHeight === "true") info.custom = true;
    if (ra.hidden === "1" || ra.hidden === "true") info.hidden = true;
    if (ra.outlineLevel && ra.outlineLevel !== "0") info.outline = Number(ra.outlineLevel);
    if (ra.collapsed === "1" || ra.collapsed === "true") info.collapsed = true;
    if (ra.customFormat === "1" || ra.customFormat === "true") {
      const s = style(ra.s, c.styles);
      if (s) info.s = s;
    }
    const extra = Object.entries(ra).filter(([k, v]) => !ROW_ATTRS.test(k) && !(k === "x14ac:dyDescent" && v === c.dyDescent))
      .map(([k, v]) => ` ${k}="${v.replace(/&/g, "&amp;").replace(/"/g, "&quot;")}"`).join("");
    if (extra) info.raw = extra;
    if (Object.keys(info).length) sheet.rows.set(rowIndex, info);
    let colIndex = -1;
    for (const ce of children(xml, rowEl.open, rowEl.close)) {
      if (ce.name !== "c") continue;
      const a = attrs(xml, ce);
      const at = a.r ? parseCell(a.r) : null;
      colIndex = at ? at.col : colIndex + 1;
      const cell = readCell(xml, ce, a, rowIndex, colIndex, shared, c);
      if (cell) sheet.cells.set(cellKey(rowIndex, colIndex), cell);
    }
  }
}

const CELL_ATTRS = /^(r|s|t)$/;
const ROW_ATTRS = /^(r|spans|s|customFormat|ht|customHeight|hidden|outlineLevel|collapsed)$/;

function readCell(xml: string, ce: El, a: Record<string, string>, row: number, col: number, shared: Map<string, { f: string; row: number; col: number }>, c: SheetContext): Cell | null {
  const cell: Cell = { v: null };
  const s = style(a.s, c.styles);
  if (s) cell.s = s;
  const extra = Object.entries(a).filter(([k]) => !CELL_ATTRS.test(k)).map(([k, v]) => ` ${k}="${v.replace(/&/g, "&amp;").replace(/"/g, "&quot;")}"`).join("");
  if (extra) cell.raw = { attrs: extra };
  let vText: string | null = null;
  let isEl: El | undefined;
  if (ce.open !== ce.end) {
    for (const k of children(xml, ce.open, ce.close)) {
      if (k.name === "v") vText = text(xml, k);
      else if (k.name === "is") isEl = k;
      else if (k.name === "f") {
        const fa = attrs(xml, k);
        const body = unesc(text(xml, k));
        const t = fa.t ?? "normal";
        if (t === "shared" && fa.si !== undefined) {
          const master = shared.get(fa.si);
          if (body || !master) {
            shared.set(fa.si, { f: body, row, col });
            cell.f = body;
          } else cell.f = shiftFormula(master.f, row - master.row, col - master.col);
        } else if (t === "array") {
          cell.f = body;
          const rect = parseRange(fa.ref ?? "");
          if (rect) cell.array = rect;
        } else if (t === "normal") cell.f = body;
        else cell.raw = { ...cell.raw, f: outer(xml, k) };
      }
    }
  }
  if (cell.f !== undefined) {
    // The model's form has no _xlfn. prefixes; keep the file's text when they would not come back alike.
    const file = cell.f;
    cell.f = fromXlsxFormula(file);
    if (toXlsxFormula(cell.f) !== file) cell.raw = { ...cell.raw, file };
  }
  const t = a.t ?? "n";
  if (t === "s") {
    const item = vText === null ? undefined : c.sst[Number(vText)];
    if (item) {
      cell.v = item.text;
      if (item.rich) cell.richXml = item.rich;
    }
  } else if (t === "inlineStr") {
    if (isEl) {
      cell.v = richText(xml, isEl);
      const inner = text(xml, isEl);
      if (/<(r|rPh)\b/.test(inner)) cell.richXml = inner;
    }
  } else if (vText !== null) {
    const v = unesc(vText);
    if (t === "str") cell.v = decodeX(v);
    else if (t === "b") cell.v = v === "1" || v === "true";
    else if (t === "e") cell.v = err((ERRORS.has(v) ? v : "#VALUE!") as ErrorCode);
    else if (t === "d") cell.v = isoToSerial(v, c.date1904);
    else if (v !== "") cell.v = Number(v);
  } else if (t === "str" && cell.f !== undefined) cell.v = "";
  if (cell.v === null && cell.f === undefined && cell.s === undefined && !cell.raw) return null;
  return cell;
}

/** An ISO 8601 date (cell type "d") as a serial number. */
function isoToSerial(iso: string, date1904: boolean): number {
  const ms = Date.parse(/[zZ]|[+-]\d\d:\d\d$/.test(iso) ? iso : iso + "Z");
  if (Number.isNaN(ms)) return 0;
  const days = ms / 86400000;
  // 1900 system: day 1 is 1900-01-01 and Excel counts the phantom 29 Feb 1900.
  return date1904 ? days + 24107 : days + 25569;
}
