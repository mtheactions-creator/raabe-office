// AutoFilter, conditional formatting and page setup of a worksheet: read into the model and
// written back. Elements and rules the user did not change keep Excel's XML; what the model
// cannot express stays in `rawXml` fields.

import { MAX_COLS, MAX_ROWS, fromXlsxFormula, parseRange, rangeName, toXlsxFormula, type AutoFilter, type FilterColumn, type CondFormat, type CondRule, type PageSetup, type Rect } from "@raabe/sheet-engine";
import { resolveColor, type ColorContext } from "./colors";
import { canon, readDxf, type StylesWriter } from "./styles";
import { attrs, child, children, decodeX, encodeX, esc, escText, local, outer, root, setAttrs, text, unesc, type El } from "./xml";
import type { Validation } from "@raabe/sheet-engine";

/** Excel's built-in names for a sheet's print area and AutoFilter range. */
export const PRINT_AREA = "_xlnm.Print_Area";
export const PRINT_TITLES = "_xlnm.Print_Titles";
export const FILTER_DB = "_xlnm._FilterDatabase";

// ----- AutoFilter -----------------------------------------------------------------------

export function readAutoFilter(xml: string, el: El): AutoFilter | undefined {
  const ref = parseRange(attrs(xml, el).ref ?? "");
  if (!ref) return undefined;
  const af: AutoFilter = { ref, columns: {} };
  let rest = "";
  for (const c of children(xml, el.open, el.close)) {
    if (c.name === "filterColumn") {
      const a = attrs(xml, c);
      const kids = children(xml, c.open, c.close);
      const filters = kids.length === 1 && kids[0].name === "filters" ? kids[0] : undefined;
      const plain = filters && children(xml, filters.open, filters.close).every((f) => f.name === "filter")
        && Object.keys(attrs(xml, filters)).every((k) => k === "blank") && Object.keys(a).every((k) => k === "colId");
      // Benutzerdefinierter AutoFilter, Top 10, Über/Unter dem Durchschnitt.
      const only = kids.length === 1 && Object.keys(a).every((k) => k === "colId") ? kids[0] : undefined;
      if (only?.name === "customFilters") {
        const items = children(xml, only.open, only.close).filter((f) => f.name === "customFilter").map((f) => { const fa = attrs(xml, f); return { op: fa.operator ?? "equal", val: fa.val ?? "" }; });
        const and = attrs(xml, only).and;
        if (items.length && items.length <= 2) { af.columns[Number(a.colId)] = { custom: { and: and === "1" || and === "true" || undefined, items } }; continue; }
      }
      if (only?.name === "top10") {
        const t = attrs(xml, only);
        af.columns[Number(a.colId)] = { top10: { top: t.top !== "0" && t.top !== "false", percent: t.percent === "1" || t.percent === "true", val: Number(t.val ?? 10), filterVal: t.filterVal !== undefined ? Number(t.filterVal) : undefined } };
        continue;
      }
      if (only?.name === "dynamicFilter") {
        const t = attrs(xml, only);
        if (t.type === "aboveAverage" || t.type === "belowAverage") { af.columns[Number(a.colId)] = { dynamic: { type: t.type, val: t.val !== undefined ? Number(t.val) : undefined } }; continue; }
      }
      if (filters && plain) {
        const fa = attrs(xml, filters);
        const col: AutoFilter["columns"][number] = {};
        const values = children(xml, filters.open, filters.close).map((f) => attrs(xml, f).val ?? "");
        if (values.length) col.values = values;
        if (fa.blank === "1" || fa.blank === "true") col.blank = true;
        af.columns[Number(a.colId)] = col;
        continue;
      }
    }
    rest += outer(xml, c);
  }
  if (rest) af.rawXml = rest;
  return af;
}

/** One column's criteria as Excel writes them. */
function filterXml(c: FilterColumn): string {
  if (c.custom) return `<customFilters${c.custom.and ? ' and="1"' : ""}>${c.custom.items.map((i) => `<customFilter${i.op !== "equal" ? ` operator="${i.op}"` : ""} val="${esc(i.val)}"/>`).join("")}</customFilters>`;
  if (c.top10) {
    const t = c.top10;
    return `<top10${t.top ? "" : ' top="0"'}${t.percent ? ' percent="1"' : ""} val="${t.val}"${t.filterVal !== undefined ? ` filterVal="${t.filterVal}"` : ""}/>`;
  }
  if (c.dynamic) return `<dynamicFilter type="${c.dynamic.type}"${c.dynamic.val !== undefined ? ` val="${c.dynamic.val}"` : ""}/>`;
  return `<filters${c.blank ? ' blank="1"' : ""}>${(c.values ?? []).map((v) => `<filter val="${esc(v)}"/>`).join("")}</filters>`;
}

export function autoFilterXml(af: AutoFilter, old: string | undefined): string {
  const tag = old ? setAttrs(/^<autoFilter\b[^>]*?\/?>/.exec(old)![0].replace(/\/>$/, ">"), { ref: rangeName(af.ref) }) : `<autoFilter ref="${rangeName(af.ref)}">`;
  const raw = af.rawXml ?? "";
  const rawCols = [...raw.matchAll(/<filterColumn\b[\s\S]*?(?:<\/filterColumn>|\/>)/g)].map((m) => ({ col: Number(/colId="(\d+)"/.exec(m[0])?.[1] ?? 0), xml: m[0] }));
  const rest = raw.replace(/<filterColumn\b[\s\S]*?(?:<\/filterColumn>|\/>)/g, "");
  const cols = Object.entries(af.columns).map(([k, c]) => ({
    col: Number(k),
    xml: `<filterColumn colId="${k}">${filterXml(c)}</filterColumn>`,
  }));
  const all = [...cols, ...rawCols.filter((r) => !(r.col in af.columns))].sort((a, b) => a.col - b.col);
  const body = all.map((c) => c.xml).join("") + rest;
  return body ? `${tag}${body}</autoFilter>` : tag.replace(/>$/, "/>");
}

// ----- Conditional formatting -------------------------------------------------------------

const TYPES = new Set<CondRule["type"]>(["cellIs", "containsText", "notContainsText", "beginsWith", "endsWith", "duplicateValues", "uniqueValues",
  "top10", "aboveAverage", "dataBar", "colorScale", "expression", "containsBlanks", "timePeriod", "iconSet"]);
/** cfRule attributes the model holds. */
const MODELLED = new Set(["type", "dxfId", "priority", "operator", "text", "rank", "percent", "bottom", "aboveAverage", "stopIfTrue"]);

export interface CfRaw {
  /** The original <conditionalFormatting> elements. */
  xml: string[];
  /** Original rules by their model form (without priority), with their dxf index. */
  rules: { key: string; xml: string; dxf?: number }[];
}

const yes = (v: string | undefined) => v === "1" || v === "true";

export function readConditional(xml: string, els: El[], ctx: ColorContext, dxfs: string[], raw: CfRaw): CondFormat[] {
  const out: CondFormat[] = [];
  for (const el of els) {
    raw.xml.push(outer(xml, el));
    const ref = (attrs(xml, el).sqref ?? "").split(/\s+/).map((t) => parseRange(t)).filter((r): r is Rect => !!r);
    const cf: CondFormat = { ref, rules: [] };
    for (const r of children(xml, el.open, el.close)) {
      if (r.name !== "cfRule") continue;
      const a = attrs(xml, r);
      const type = (TYPES.has(a.type as CondRule["type"]) ? a.type : "other") as CondRule["type"];
      const rule: CondRule = { type, priority: Number(a.priority ?? 1) };
      if (a.operator) rule.operator = a.operator;
      if (a.text !== undefined) rule.text = a.text;
      if (a.rank !== undefined) rule.rank = Number(a.rank);
      if (yes(a.percent)) rule.percent = true;
      if (yes(a.bottom)) rule.bottom = true;
      if (a.aboveAverage !== undefined) rule.above = yes(a.aboveAverage);
      if (yes(a.stopIfTrue)) rule.stopIfTrue = true;
      const kids = children(xml, r.open, r.close);
      const formulas = kids.filter((k) => k.name === "formula").map((k) => fromXlsxFormula(unesc(text(xml, k))));
      if (formulas.length) rule.formulas = formulas;
      const dxf = a.dxfId !== undefined ? Number(a.dxfId) : undefined;
      if (dxf !== undefined && dxfs[dxf]) rule.format = readDxf(dxfs[dxf], ctx);
      let plain = type !== "other" && type !== "iconSet" && type !== "timePeriod" && Object.keys(a).every((k) => MODELLED.has(k));
      for (const k of kids) {
        if (k.name === "formula") continue;
        if (k.name === "colorScale" || k.name === "dataBar") {
          const parts = children(xml, k.open, k.close);
          rule.colors = parts.filter((p) => p.name === "color").map((p) => resolveColor(attrs(xml, p), ctx) ?? "#000000");
          const cfvo = parts.filter((p) => p.name === "cfvo").map((p) => canon(attrs(xml, p)));
          const std = cfvo.length === 3 ? [canon({ type: "min" }), canon({ type: "percentile", val: "50" }), canon({ type: "max" })] : [canon({ type: "min" }), canon({ type: "max" })];
          if (canon(cfvo) !== canon(std) || Object.keys(attrs(xml, k)).length) plain = false;
        } else plain = false;
      }
      if (!plain) rule.rawXml = outer(xml, r);
      raw.rules.push({ key: ruleKey(rule), xml: outer(xml, r), dxf });
      cf.rules.push(rule);
    }
    out.push(cf);
  }
  return out;
}

function ruleKey(r: CondRule): string {
  return canon({ ...r, priority: undefined });
}

/** The <conditionalFormatting> elements for the model's formats. */
export function conditionalXml(list: CondFormat[], raw: CfRaw | undefined, sw: StylesWriter): string[] {
  return list.filter((cf) => cf.rules.length && cf.ref.length).map((cf) => {
    const rules = cf.rules.map((r) => {
      const prio = String(r.priority);
      if (r.rawXml) return r.rawXml.replace(/^<cfRule\b[^>]*>/, (t) => setAttrs(t, { priority: prio }));
      const old = raw?.rules.find((o) => o.key === ruleKey(r));
      if (old) {
        return old.xml.replace(/^<cfRule\b[^>]*?\/?>/, (t) => {
          const dxf = r.format ? sw.dxf(r.format, old.dxf) : undefined;
          return setAttrs(t, { priority: prio, dxfId: dxf === undefined ? undefined : String(dxf) });
        });
      }
      return ruleXml(r, sw);
    });
    return `<conditionalFormatting sqref="${cf.ref.map(rangeName).join(" ")}">${rules.join("")}</conditionalFormatting>`;
  });
}

function ruleXml(r: CondRule, sw: StylesWriter): string {
  let a = ` type="${r.type}"`;
  if (r.format && r.type !== "dataBar" && r.type !== "colorScale") a += ` dxfId="${sw.dxf(r.format)}"`;
  a += ` priority="${r.priority}"`;
  if (r.stopIfTrue) a += ' stopIfTrue="1"';
  if (r.above === false) a += ' aboveAverage="0"';
  if (r.percent) a += ' percent="1"';
  if (r.bottom) a += ' bottom="1"';
  if (r.rank !== undefined) a += ` rank="${r.rank}"`;
  if (r.operator) a += ` operator="${esc(r.operator)}"`;
  if (r.text !== undefined) a += ` text="${esc(r.text)}"`;
  let body = (r.formulas ?? []).map((f) => `<formula>${escText(toXlsxFormula(f))}</formula>`).join("");
  const colors = r.colors ?? [];
  const rgb = (c: string) => `<color rgb="FF${c.replace("#", "").toUpperCase()}"/>`;
  if (r.type === "colorScale") {
    const cfvo = colors.length === 3 ? '<cfvo type="min"/><cfvo type="percentile" val="50"/><cfvo type="max"/>' : '<cfvo type="min"/><cfvo type="max"/>';
    body = `<colorScale>${cfvo}${colors.map(rgb).join("")}</colorScale>`;
  } else if (r.type === "dataBar") {
    body = `<dataBar><cfvo type="min"/><cfvo type="max"/>${rgb(colors[0] ?? "#638EC6")}</dataBar>`;
  }
  return body ? `<cfRule${a}>${body}</cfRule>` : `<cfRule${a}/>`;
}

// ----- Datenüberprüfung -------------------------------------------------------------------
// <dataValidations> in the sheet, and Excel's x14 form in the extLst for criteria that refer
// to other sheets. Literal lists whose entries hold a comma or quote come as Excel 365 writes
// them: an mc:AlternateContent with x12ac:list (the entries) and formula1 as the fallback.

/** The extLst entry of x14 data validations. */
export const DV_EXT = "{CCE6A557-97BC-4b89-ADB6-D9C93CAAB3DF}";
const CF_EXT = "{78C0D931-6437-407d-A8EE-F0AAD7539E65}";
const X14_NS = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main";
const XM_NS = "http://schemas.microsoft.com/office/excel/2006/main";
const X12AC = 'xmlns:x12ac="http://schemas.microsoft.com/office/spreadsheetml/2011/1/ac" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"';

const DV_TYPES = new Set<Validation["type"]>(["whole", "decimal", "list", "date", "time", "textLength", "custom"]);
const DV_OPS = new Set(["between", "notBetween", "equal", "notEqual", "greaterThan", "lessThan", "greaterThanOrEqual", "lessThanOrEqual"]);
/** dataValidation attributes the model holds (xr:uid stays with Excel's own text). */
const DV_ATTRS = new Set(["type", "errorStyle", "operator", "allowBlank", "showDropDown", "showInputMessage", "showErrorMessage", "errorTitle", "error", "promptTitle", "prompt", "sqref", "xr:uid"]);

export interface DvRaw {
  /** The file's elements by their model form (without ref), and whether they were in the extLst. */
  items: { key: string; xml: string; x14: boolean }[];
}

/** A validation's model form without its ranges (flags that are false count as absent). */
const dvKey = (v: Validation) =>
  canon(Object.fromEntries(Object.entries(v).filter(([k, x]) => k !== "ref" && x !== false && x !== undefined)));

/** The validations of a sheet: `main` is <dataValidations>, `ext` the sheet's <extLst>. */
export function readValidations(xml: string, main: El | undefined, ext: El | undefined, raw: DvRaw): Validation[] {
  const out: Validation[] = [];
  for (const el of main ? children(xml, main.open, main.close) : []) {
    if (el.name === "dataValidation") out.push(readValidation(xml, el, false, raw));
  }
  for (const e of ext ? children(xml, ext.open, ext.close) : []) {
    if (e.name !== "ext" || attrs(xml, e).uri !== DV_EXT) continue;
    for (const list of children(xml, e.open, e.close)) {
      if (list.name !== "x14:dataValidations") continue;
      for (const el of children(xml, list.open, list.close)) if (el.name === "x14:dataValidation") out.push(readValidation(xml, el, true, raw));
    }
  }
  return out;
}

function readValidation(xml: string, el: El, x14: boolean, raw: DvRaw): Validation {
  const a = attrs(xml, el);
  const type = a.type === undefined || a.type === "none" ? "any" : DV_TYPES.has(a.type as Validation["type"]) ? (a.type as Validation["type"]) : "any";
  const v: Validation = { ref: [], type };
  let plain = a.type === undefined || a.type === "none" || DV_TYPES.has(a.type as Validation["type"]);
  if (a.operator !== undefined) {
    if (DV_OPS.has(a.operator)) v.operator = a.operator as Validation["operator"];
    else plain = false;
  }
  if (yes(a.allowBlank)) v.allowBlank = true;
  if (yes(a.showDropDown)) v.showDropDown = true;
  if (yes(a.showInputMessage)) v.showInputMessage = true;
  if (yes(a.showErrorMessage)) v.showErrorMessage = true;
  if (a.errorStyle === "stop" || a.errorStyle === "warning" || a.errorStyle === "information") v.errorStyle = a.errorStyle;
  else if (a.errorStyle !== undefined) plain = false;
  if (a.errorTitle !== undefined) v.errorTitle = decodeX(a.errorTitle);
  if (a.error !== undefined) v.error = decodeX(a.error);
  if (a.promptTitle !== undefined) v.promptTitle = decodeX(a.promptTitle);
  if (a.prompt !== undefined) v.prompt = decodeX(a.prompt);
  if (!Object.keys(a).every((k) => DV_ATTRS.has(k))) plain = false;
  let sqref = a.sqref ?? "";
  const formula = (k: El) => {
    if (!x14) return fromXlsxFormula(unesc(text(xml, k)));
    const f = child(xml, k, "xm:f");
    return f ? fromXlsxFormula(unesc(text(xml, f))) : "";
  };
  for (const k of children(xml, el.open, el.close)) {
    const name = local(k.name);
    if (name === "formula1") v.formula1 = formula(k);
    else if (name === "formula2") v.formula2 = formula(k);
    else if (k.name === "xm:sqref") sqref = text(xml, k);
    else if (k.name === "mc:AlternateContent") {
      const list = /<x12ac:list>([\s\S]*?)<\/x12ac:list>/.exec(outer(xml, k));
      const fb = /<formula1>([\s\S]*?)<\/formula1>/.exec(outer(xml, k));
      if (list && fb) {
        v.items = csvItems(unesc(list[1]));
        v.formula1 = fromXlsxFormula(unesc(fb[1]));
      } else plain = false;
    } else plain = false;
  }
  v.ref = sqref.trim().split(/\s+/).map((t) => parseRange(t)).filter((r): r is Rect => !!r);
  if (!plain) v.rawXml = outer(xml, el);
  raw.items.push({ key: dvKey(v), xml: outer(xml, el), x14 });
  return v;
}

/** Excel's x12ac:list: entries separated by commas, quoted ("") when they hold one. */
function csvItems(s: string): string[] {
  const out: string[] = [];
  let i = 0;
  for (;;) {
    let item = "";
    if (s[i] === '"') {
      i++;
      for (; i < s.length; i++) {
        if (s[i] === '"') {
          if (s[i + 1] === '"') { item += '"'; i++; } else { i++; break; }
        } else item += s[i];
      }
    }
    const comma = s.indexOf(",", i);
    item += s.slice(i, comma < 0 ? s.length : comma);
    out.push(item);
    if (comma < 0) return out;
    i = comma + 1;
  }
}

const csvItem = (s: string) => (/[,"]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s);

/** Whether a criterion names another sheet (Excel then writes the validation in the extLst). */
const otherSheet = (f: string | undefined) => !!f && f.replace(/"(?:[^"]|"")*"/g, "").includes("!");

/**
 * <dataValidations> and the sheet's <extLst> for the model's validations: elements Excel wrote
 * keep their text (with the new sqref), new ones are written as Excel writes them.
 */
export function validationsXml(list: Validation[], raw: DvRaw | undefined, oldMain: string | undefined, oldExt: string | undefined): { main: string | null; ext: string | null } {
  const used = new Set<number>();
  const main: string[] = [];
  const x14: string[] = [];
  for (const v of list.filter((x) => x.ref.length)) {
    const sqref = v.ref.map(rangeName).join(" ");
    const key = dvKey(v);
    const at = raw ? raw.items.findIndex((o, i) => !used.has(i) && o.key === key) : -1;
    let xml: string;
    let inExt: boolean;
    if (at >= 0) {
      used.add(at);
      inExt = raw!.items[at].x14;
      xml = withSqref(raw!.items[at].xml, sqref, inExt);
    } else if (v.rawXml) {
      inExt = v.rawXml.startsWith("<x14:");
      xml = withSqref(v.rawXml.replace(/\sxr:uid="[^"]*"/, ""), sqref, inExt);
    } else {
      inExt = otherSheet(v.formula1) || otherSheet(v.formula2);
      xml = validationXml(v, sqref, inExt);
    }
    (inExt ? x14 : main).push(xml);
  }
  let mainXml: string | null = null;
  if (main.length) {
    const tag = oldMain ? /^<dataValidations\b[^>]*?\/?>/.exec(oldMain)![0].replace(/\/>$/, ">") : "<dataValidations>";
    mainXml = `${setAttrs(tag, { count: String(main.length) })}${main.join("")}</dataValidations>`;
  }
  return { main: mainXml, ext: extXml(oldExt, x14) };
}

function withSqref(xml: string, sqref: string, x14: boolean): string {
  if (x14) return xml.replace(/<xm:sqref>[\s\S]*?<\/xm:sqref>/, `<xm:sqref>${sqref}</xm:sqref>`);
  return xml.replace(/^<dataValidation\b[^>]*?\/?>/, (t) => setAttrs(t, { sqref }));
}

function validationXml(v: Validation, sqref: string, x14: boolean): string {
  let a = "";
  if (v.type !== "any") a += ` type="${v.type}"`;
  if (v.errorStyle && v.errorStyle !== "stop") a += ` errorStyle="${v.errorStyle}"`;
  if (v.operator && v.operator !== "between" && v.type !== "list" && v.type !== "custom" && v.type !== "any") a += ` operator="${v.operator}"`;
  if (v.allowBlank) a += ' allowBlank="1"';
  if (v.showDropDown) a += ' showDropDown="1"';
  if (v.showInputMessage) a += ' showInputMessage="1"';
  if (v.showErrorMessage) a += ' showErrorMessage="1"';
  // Line breaks as character references (a literal one would read back as a space).
  for (const k of ["errorTitle", "error", "promptTitle", "prompt"] as const) if (v[k] !== undefined) a += ` ${k}="${esc(encodeX(v[k]!)).replace(/\r?\n/g, "&#10;")}"`;
  const two = v.type !== "list" && v.type !== "custom" && v.type !== "any" && (!v.operator || v.operator === "between" || v.operator === "notBetween");
  const fs = v.type === "any" ? [] : [v.formula1, two ? v.formula2 : undefined];
  if (x14) {
    const body = fs.map((f, i) => (f === undefined || f === "" ? "" : `<x14:formula${i + 1}><xm:f>${escText(toXlsxFormula(f))}</xm:f></x14:formula${i + 1}>`)).join("");
    return `<x14:dataValidation${a}>${body}<xm:sqref>${sqref}</xm:sqref></x14:dataValidation>`;
  }
  a += ` sqref="${sqref}"`;
  let body = fs.map((f, i) => (f === undefined || f === "" ? "" : `<formula${i + 1}>${escText(toXlsxFormula(f))}</formula${i + 1}>`)).join("");
  if (v.type === "list" && v.items) {
    body = `<mc:AlternateContent ${X12AC}><mc:Choice Requires="x12ac"><x12ac:list>${escText(v.items.map(csvItem).join(","))}</x12ac:list></mc:Choice>` +
      `<mc:Fallback><formula1>${escText(v.formula1 ?? "")}</formula1></mc:Fallback></mc:AlternateContent>`;
  }
  return body ? `<dataValidation${a}>${body}</dataValidation>` : `<dataValidation${a}/>`;
}

/** The sheet's <extLst> with the x14 validations' entry replaced, added (after conditional formats) or removed. */
function extXml(old: string | undefined, x14: string[]): string | null {
  const re = new RegExp(`<ext\\b[^>]*uri="${DV_EXT.replace(/[{}]/g, "\\$&")}"[^>]*>[\\s\\S]*?</ext>`);
  const oldEntry = old ? re.exec(old)?.[0] : undefined;
  let entry = "";
  if (x14.length) {
    const head = oldEntry ? /^<ext\b[^>]*>/.exec(oldEntry)![0] : `<ext uri="${DV_EXT}" xmlns:x14="${X14_NS}">`;
    const oldList = oldEntry ? /<x14:dataValidations\b[^>]*>/.exec(oldEntry)?.[0] : undefined;
    const listTag = oldList ? setAttrs(oldList, { count: String(x14.length) }) : `<x14:dataValidations count="${x14.length}" xmlns:xm="${XM_NS}">`;
    entry = `${head}${listTag}${x14.join("")}</x14:dataValidations></ext>`;
  }
  if (!old) return entry ? `<extLst>${entry}</extLst>` : null;
  let out: string;
  if (oldEntry) out = old.replace(oldEntry, entry);
  else if (!entry) return old;
  else {
    const cf = new RegExp(`<ext\\b[^>]*uri="${CF_EXT.replace(/[{}]/g, "\\$&")}"[^>]*>[\\s\\S]*?</ext>`).exec(old)?.[0];
    out = cf ? old.replace(cf, cf + entry) : old.replace(/^<extLst\b[^>]*>/, (t) => t + entry);
  }
  return /^<extLst\b[^>]*>\s*<\/extLst>$/.test(out) ? null : out;
}

// ----- Page setup -----------------------------------------------------------------------

export function readPage(parts: { name: string; xml: string }[]): PageSetup | undefined {
  const get = (n: string) => parts.find((p) => p.name === n)?.xml;
  const at = (x: string | undefined) => (x ? attrs(x, root(x)) : undefined);
  const m = at(get("pageMargins")), ps = at(get("pageSetup")), po = at(get("printOptions"));
  const pr = get("sheetPr");
  const fit = pr ? /<pageSetUpPr\b[^>]*fitToPage="(1|true)"/.test(pr) : false;
  if (!m && !ps && !po) return undefined;
  const page: PageSetup = {};
  if (ps?.orientation === "portrait" || ps?.orientation === "landscape") page.orientation = ps.orientation;
  if (ps?.paperSize) page.paperSize = Number(ps.paperSize);
  if (m) page.margins = { left: n(m.left, 0.7), right: n(m.right, 0.7), top: n(m.top, 0.75), bottom: n(m.bottom, 0.75), header: n(m.header, 0.3), footer: n(m.footer, 0.3) };
  if (ps?.scale) page.scale = Number(ps.scale);
  if (fit) {
    page.fitWidth = n(ps?.fitToWidth, 1);
    page.fitHeight = n(ps?.fitToHeight, 1);
  }
  if (yes(po?.gridLines)) page.gridLines = true;
  if (yes(po?.headings)) page.headings = true;
  const hf = get("headerFooter");
  if (hf) {
    const part = (name: string) => { const m = new RegExp(`<${name}>([\\s\\S]*?)</${name}>`).exec(hf); return m ? unesc(m[1]) : undefined; };
    page.header = part("oddHeader");
    page.footer = part("oddFooter");
  }
  return page;
}

/** <headerFooter> with the model's odd header and footer; Excel's other parts (first page, even pages) stay. */
export function headerFooterXml(p: PageSetup | undefined, old: string | undefined): string | null {
  let x = old ?? "<headerFooter/>";
  if (x.endsWith("/>")) x = x.slice(0, -2).replace(/\s+$/, "") + "></headerFooter>";
  for (const [name, text] of [["oddHeader", p?.header], ["oddFooter", p?.footer]] as const) {
    const el = text ? `<${name}>${escText(text)}</${name}>` : "";
    const re = new RegExp(`<${name}>[\\s\\S]*?</${name}>|<${name}/>`);
    if (re.test(x)) x = x.replace(re, el);
    else if (el) {
      // oddHeader first, oddFooter right after it (CT_HeaderFooter order).
      const after = name === "oddFooter" ? /<\/oddHeader>/.exec(x) : null;
      const at = after ? after.index + after[0].length : /^<headerFooter[^>]*>/.exec(x)![0].length;
      x = x.slice(0, at) + el + x.slice(at);
    }
  }
  return /^<headerFooter[^>]*><\/headerFooter>$/.test(x) ? null : x;
}

const n = (v: string | undefined, d: number) => (v === undefined || v === "" ? d : Number(v));
const num = (v: number) => String(v);

/** The page elements for a changed PageSetup (attributes the model does not hold are kept). */
export function pageXml(p: PageSetup | undefined, parts: { name: string; xml: string }[]): { margins: string | null; setup: string | null; options: string | null; fit: boolean } {
  const get = (name: string) => parts.find((x) => x.name === name)?.xml;
  const m = p?.margins;
  const margins = m ? `<pageMargins left="${num(m.left)}" right="${num(m.right)}" top="${num(m.top)}" bottom="${num(m.bottom)}" header="${num(m.header)}" footer="${num(m.footer)}"/>`
    : get("pageMargins") ?? null;
  const fit = p?.fitWidth !== undefined || p?.fitHeight !== undefined;
  const change = {
    orientation: p?.orientation ?? null, paperSize: p?.paperSize !== undefined ? String(p.paperSize) : null, scale: p?.scale !== undefined ? String(p.scale) : null,
    fitToWidth: fit && p?.fitWidth !== 1 ? String(p?.fitWidth ?? 1) : null, fitToHeight: fit && p?.fitHeight !== 1 ? String(p?.fitHeight ?? 1) : null,
  };
  const oldSetup = get("pageSetup");
  let setup: string | null = setAttrs(oldSetup ?? "<pageSetup/>", change);
  if (setup === "<pageSetup/>") setup = null;
  let options: string | null = setAttrs(get("printOptions") ?? "<printOptions/>", { gridLines: p?.gridLines ? "1" : null, headings: p?.headings ? "1" : null });
  if (options === "<printOptions/>") options = null;
  return { margins, setup, options, fit };
}

/** sheetPr with pageSetUpPr fitToPage set or cleared. */
export function fitToPage(pr: string | undefined, fit: boolean): string | null {
  let x = pr ?? "<sheetPr/>";
  const has = /<pageSetUpPr\b[^>]*fitToPage="(1|true)"/.test(x);
  if (has === fit) return pr ?? null;
  if (/<pageSetUpPr\b/.test(x)) x = x.replace(/<pageSetUpPr\b[^>]*\/>/, (t) => setAttrs(t, { fitToPage: fit ? "1" : null }));
  else if (fit) {
    const el = '<pageSetUpPr fitToPage="1"/>';
    // After tabColor and outlinePr (CT_SheetPr order).
    if (x.endsWith("/>")) x = `${x.slice(0, -2).replace(/\s+$/, "")}>${el}</sheetPr>`;
    else x = x.replace(/<\/sheetPr>$/, `${el}</sheetPr>`);
  }
  x = x.replace(/<pageSetUpPr\s*\/>/, "");
  return /^<sheetPr\s*\/>$|^<sheetPr>\s*<\/sheetPr>$/.test(x) ? null : x;
}

// Letters beyond ASCII (Ü …) need no quotes either.
const WIDE = `${String.fromCharCode(0xc0)}-${String.fromCharCode(0xffef)}`;
const BARE_NAME = new RegExp(`^[A-Za-z_${WIDE}][A-Za-z0-9_.${WIDE}]*$`);

/** A reference to a range of a sheet as names write it: 'Mein Blatt'!$A$1:$B$4. */
export function sheetRef(sheet: string, r: Rect): string {
  const q = BARE_NAME.test(sheet) && !/^[A-Za-z]{1,3}\d+$/.test(sheet) && !/^(TRUE|FALSE|R|C|RC)$/i.test(sheet) ? sheet : `'${sheet.replace(/'/g, "''")}'`;
  const abs = rangeName(r).replace(/([A-Z]+)(\d+)/g, "$$$1$$$2").replace(/^([A-Z]+):([A-Z]+)$/, "$$$1:$$$2").replace(/^(\d+):(\d+)$/, "$$$1:$$$2");
  return `${q}!${abs}`;
}

/** Drucktitel from the name's formula ("Tabelle1!$A:$A,Tabelle1!$1:$2"), else null. */
export function readTitles(ref: string, sheet: string): { rows?: [number, number]; cols?: [number, number] } | null {
  const out: { rows?: [number, number]; cols?: [number, number] } = {};
  for (const part of ref.split(/,(?=(?:[^']*'[^']*')*[^']*$)/)) {
    const r = ownRange(part, sheet);
    if (!r) return null;
    if (r.c0 === 0 && r.c1 === MAX_COLS - 1) out.rows = [r.r0, r.r1];
    else if (r.r0 === 0 && r.r1 === MAX_ROWS - 1) out.cols = [r.c0, r.c1];
    else return null;
  }
  return out.rows || out.cols ? out : null;
}

/** The name's formula for Drucktitel, columns first as Excel writes it. */
export function titlesRef(sheet: string, rows?: [number, number], cols?: [number, number]): string {
  const parts: string[] = [];
  if (cols) parts.push(sheetRef(sheet, { r0: 0, c0: cols[0], r1: MAX_ROWS - 1, c1: cols[1] }));
  if (rows) parts.push(sheetRef(sheet, { r0: rows[0], c0: 0, r1: rows[1], c1: MAX_COLS - 1 }));
  return parts.join(",");
}

/** The single range of a name's formula on its own sheet ("Tabelle1!$A$1:$B$4"), else null. */
export function ownRange(ref: string, sheet: string): Rect | null {
  const m = /^(?:'((?:[^']|'')+)'|([^'!]+))!(\$?[A-Z]+\$?\d*(?::\$?[A-Z]+\$?\d*)?|\$?\d+:\$?\d+)$/.exec(ref.trim());
  if (!m) return null;
  const name = m[1] !== undefined ? m[1].replace(/''/g, "'") : m[2];
  if (name !== sheet) return null;
  return parseRange(m[3]);
}
