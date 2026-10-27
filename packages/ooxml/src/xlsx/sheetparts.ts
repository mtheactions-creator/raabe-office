// AutoFilter, conditional formatting and page setup of a worksheet: read into the model and
// written back. Elements and rules the user did not change keep Excel's XML; what the model
// cannot express stays in `rawXml` fields.

import { fromXlsxFormula, parseRange, rangeName, toXlsxFormula, type AutoFilter, type CondFormat, type CondRule, type PageSetup, type Rect } from "@raabe/sheet-engine";
import { resolveColor, type ColorContext } from "./colors";
import { canon, readDxf, type StylesWriter } from "./styles";
import { attrs, children, esc, escText, outer, root, setAttrs, text, unesc, type El } from "./xml";

/** Excel's built-in names for a sheet's print area and AutoFilter range. */
export const PRINT_AREA = "_xlnm.Print_Area";
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

export function autoFilterXml(af: AutoFilter, old: string | undefined): string {
  const tag = old ? setAttrs(/^<autoFilter\b[^>]*?\/?>/.exec(old)![0].replace(/\/>$/, ">"), { ref: rangeName(af.ref) }) : `<autoFilter ref="${rangeName(af.ref)}">`;
  const raw = af.rawXml ?? "";
  const rawCols = [...raw.matchAll(/<filterColumn\b[\s\S]*?(?:<\/filterColumn>|\/>)/g)].map((m) => ({ col: Number(/colId="(\d+)"/.exec(m[0])?.[1] ?? 0), xml: m[0] }));
  const rest = raw.replace(/<filterColumn\b[\s\S]*?(?:<\/filterColumn>|\/>)/g, "");
  const cols = Object.entries(af.columns).map(([k, c]) => ({
    col: Number(k),
    xml: `<filterColumn colId="${k}"><filters${c.blank ? ' blank="1"' : ""}>${(c.values ?? []).map((v) => `<filter val="${esc(v)}"/>`).join("")}</filters></filterColumn>`,
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
  return page;
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

/** The single range of a name's formula on its own sheet ("Tabelle1!$A$1:$B$4"), else null. */
export function ownRange(ref: string, sheet: string): Rect | null {
  const m = /^(?:'((?:[^']|'')+)'|([^'!]+))!(\$?[A-Z]+\$?\d*(?::\$?[A-Z]+\$?\d*)?|\$?\d+:\$?\d+)$/.exec(ref.trim());
  if (!m) return null;
  const name = m[1] !== undefined ? m[1].replace(/''/g, "'") : m[2];
  if (name !== sheet) return null;
  return parseRange(m[3]);
}
