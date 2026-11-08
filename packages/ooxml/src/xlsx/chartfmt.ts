// The Format pane's settings of a chart element in DrawingML (D-066): fills and lines (c:spPr), text
// (c:txPr), and the small value elements of axes, series, labels, trendlines and error bars. Written
// in the order the schema wants (Excel refuses parts out of order) and read back from the same XML.

import type { AxisOptions, DataLabelOptions, Dash, ElementFormat, ErrorBars, Fill, Line, MarkerOptions, TextFormat, Trendline } from "@raabe/charts";
import { attrs, children, esc, type El } from "./xml";

const EMU_PT = 12700;

function srgb(color: string, transparency?: number): string {
  const hex = color.replace("#", "").slice(0, 6).toUpperCase();
  return transparency ? `<a:srgbClr val="${hex}"><a:alpha val="${Math.round((100 - transparency) * 1000)}"/></a:srgbClr>` : `<a:srgbClr val="${hex}"/>`;
}

/** A fill element, or "" to keep the automatic one (`auto` is what Excel writes then). */
export function fillXml(f: Fill | undefined, auto = ""): string {
  if (!f || f.type === "auto") return auto;
  if (f.type === "none") return "<a:noFill/>";
  return `<a:solidFill>${srgb(f.color ?? "#FFFFFF", f.transparency)}</a:solidFill>`;
}

/** An a:ln element; `auto` is Excel's automatic line for the element. */
export function lineXml(l: Line | undefined, auto = ""): string {
  if (!l || (l.type === "auto" && !l.width && !l.dash)) return auto;
  if (l.type === "none") return "<a:ln><a:noFill/></a:ln>";
  const w = l.width !== undefined ? ` w="${Math.round(l.width * EMU_PT)}"` : "";
  const fill = l.type === "solid" ? `<a:solidFill>${srgb(l.color ?? "#000000", l.transparency)}</a:solidFill>` : "";
  const dash = l.dash ? `<a:prstDash val="${l.dash}"/>` : "";
  return `<a:ln${w} cap="flat" cmpd="sng" algn="ctr">${fill}${dash}<a:round/></a:ln>`;
}

/** c:spPr from a format; `autoFill` / `autoLine` are Excel's defaults for the element. */
export function spPrXml(fmt: { fill?: Fill; line?: Line } | undefined, autoFill: string, autoLine: string): string {
  return `<c:spPr>${fillXml(fmt?.fill, autoFill)}${lineXml(fmt?.line, autoLine)}<a:effectLst/></c:spPr>`;
}

/** The run properties of a text format inside a:defRPr (size in pt, colour). */
export function defRPrAttrs(t: TextFormat | undefined, size: number): { attrs: string; fill: string | null } {
  const sz = t?.size !== undefined ? Math.round(t.size * 100) : size;
  return { attrs: ` sz="${sz}" b="${t?.bold ? 1 : 0}" i="${t?.italic ? 1 : 0}"`, fill: t?.color ? `<a:solidFill>${srgb(t.color)}</a:solidFill>` : null };
}

export function tickXml(tag: "majorTickMark" | "minorTickMark", v: AxisOptions["majorTick"]): string {
  return `<c:${tag} val="${v ?? "none"}"/>`;
}

/** c:scaling: Basis, orientation, Maximum, Minimum (schema order). */
export function scalingXml(o: AxisOptions | undefined): string {
  return `<c:scaling>${o?.logBase ? `<c:logBase val="${o.logBase}"/>` : ""}<c:orientation val="${o?.reverse ? "maxMin" : "minMax"}"/>${o?.max !== undefined ? `<c:max val="${o.max}"/>` : ""}${o?.min !== undefined ? `<c:min val="${o.min}"/>` : ""}</c:scaling>`;
}

export function crossesXml(o: AxisOptions | undefined): string {
  if (typeof o?.crosses === "number") return `<c:crossesAt val="${o.crosses}"/>`;
  return `<c:crosses val="${o?.crosses === "max" ? "max" : "autoZero"}"/>`;
}

/** The flags and position of data labels (c:dLbls without the shape and text parts). */
export function labelFlags(o: DataLabelOptions | undefined, defPos: string | undefined, leader: boolean): string {
  const pos = o?.pos ?? defPos;
  const b = (v: boolean | undefined, d: boolean) => ((v ?? d) ? 1 : 0);
  return `${pos ? `<c:dLblPos val="${pos}"/>` : ""}<c:showLegendKey val="${b(o?.showLegendKey, false)}"/><c:showVal val="${b(o?.showVal, true)}"/><c:showCatName val="${b(o?.showCat, false)}"/><c:showSerName val="${b(o?.showSer, false)}"/><c:showPercent val="${b(o?.showPercent, false)}"/><c:showBubbleSize val="0"/>${o?.separator !== undefined ? `<c:separator>${esc(o.separator)}</c:separator>` : ""}${leader ? '<c:showLeaderLines val="1"/>' : ""}`;
}

export function markerXml(mk: MarkerOptions | undefined, auto: string, autoFill: string, autoLine: string): string {
  if (!mk || (!mk.symbol && mk.size === undefined && !mk.fill && !mk.line)) return auto;
  const sym = mk.symbol && mk.symbol !== "auto" ? mk.symbol : "circle";
  if (sym === "none") return '<c:marker><c:symbol val="none"/></c:marker>';
  return `<c:marker><c:symbol val="${sym}"/><c:size val="${mk.size ?? 5}"/>${spPrXml(mk, autoFill, autoLine)}</c:marker>`;
}

export function trendlineXml(t: Trendline, autoLine: string, label: string): string {
  const type = t.type;
  const extra = (type === "poly" ? `<c:order val="${t.order ?? 2}"/>` : "") + (type === "movingAvg" ? `<c:period val="${t.period ?? 2}"/>` : "");
  const fc = (t.forward ? `<c:forward val="${t.forward}"/>` : "") + (t.backward ? `<c:backward val="${t.backward}"/>` : "") + (t.intercept !== undefined ? `<c:intercept val="${t.intercept}"/>` : "");
  return `<c:trendline>${t.name ? `<c:name>${esc(t.name)}</c:name>` : ""}<c:spPr>${lineXml(t.line, autoLine)}<a:effectLst/></c:spPr><c:trendlineType val="${type}"/>${extra}${fc}<c:dispRSqr val="${t.showR2 ? 1 : 0}"/><c:dispEq val="${t.showEq ? 1 : 0}"/>${label}</c:trendline>`;
}

/** `data` writes a custom amount (c:plus / c:minus with its numRef and cache), as Excel's Prognoseblatt does (D-125). */
export function errBarsXml(e: ErrorBars, scatter: boolean, autoLine: string, data?: (tag: string, d: { ref?: string; cache: (number | null)[]; format?: string }) => string): string {
  const cust = e.type === "cust" && data ? `${e.plus ? data("plus", e.plus) : ""}${e.minus ? data("minus", e.minus) : ""}` : "";
  return `<c:errBars>${scatter ? '<c:errDir val="y"/>' : ""}<c:errBarType val="${e.dir}"/><c:errValType val="${e.type}"/><c:noEndCap val="${e.noEndCap ? 1 : 0}"/>${cust}${e.type === "fixedVal" || e.type === "percentage" || e.type === "stdDev" ? `<c:val val="${e.value ?? (e.type === "percentage" ? 5 : 1)}"/>` : ""}<c:spPr>${e.type === "cust" ? "<a:noFill/>" : ""}${lineXml(e.line, autoLine)}<a:effectLst/></c:spPr></c:errBars>`;
}

// ----- reading ---------------------------------------------------------------------------------------

type Find = (el: El | undefined, ...path: string[]) => El | undefined;

export interface Reader { xml: string; find: Find; val: (el: El | undefined) => string | undefined }

function colorOf(r: Reader, holder: El | undefined): { color?: string; transparency?: number } {
  const c = r.find(holder, "a:srgbClr");
  if (!c) return {};
  const hex = attrs(r.xml, c).val;
  const alpha = r.find(c, "a:alpha");
  const a = alpha ? Number(attrs(r.xml, alpha).val) : undefined;
  return { color: hex ? "#" + hex.toUpperCase() : undefined, ...(a !== undefined && a < 100000 ? { transparency: Math.round(100 - a / 1000) } : {}) };
}

/** Fill and line of a c:spPr, as far as they are explicit colours (scheme colours stay automatic). */
export function readShape(r: Reader, spPr: El | undefined): { fill?: Fill; line?: Line } {
  if (!spPr) return {};
  const out: { fill?: Fill; line?: Line } = {};
  const kids = children(r.xml, spPr.open, spPr.close);
  const nf = kids.find((k) => k.name === "a:noFill"), sf = kids.find((k) => k.name === "a:solidFill");
  if (nf) out.fill = { type: "none" };
  else if (sf && r.find(sf, "a:srgbClr")) out.fill = { type: "solid", ...colorOf(r, sf) };
  const ln = kids.find((k) => k.name === "a:ln");
  if (ln) {
    const lk = children(r.xml, ln.open, ln.close);
    const w = attrs(r.xml, ln).w;
    const dash = r.val(lk.find((k) => k.name === "a:prstDash"));
    if (lk.some((k) => k.name === "a:noFill")) out.line = { type: "none" };
    else {
      const lf = lk.find((k) => k.name === "a:solidFill");
      const explicit = lf && r.find(lf, "a:srgbClr");
      if (explicit || dash) {
        out.line = { type: explicit ? "solid" : "auto", ...(explicit ? colorOf(r, lf) : {}) };
        if (w) out.line.width = Math.round((Number(w) / EMU_PT) * 100) / 100;
        if (dash && dash !== "solid") out.line.dash = dash as Dash;
      }
    }
  }
  return out;
}

/** Size (pt), bold, italic and an explicit colour of a c:txPr's first a:defRPr. */
export function readText(r: Reader, txPr: El | undefined, defaultSize: number): TextFormat | undefined {
  const d = r.find(txPr, "a:p", "a:pPr", "a:defRPr");
  if (!d) return undefined;
  const a = attrs(r.xml, d);
  const out: TextFormat = {};
  if (a.sz && Number(a.sz) !== defaultSize) out.size = Number(a.sz) / 100;
  if (a.b === "1") out.bold = true;
  if (a.i === "1") out.italic = true;
  const sf = r.find(d, "a:solidFill");
  if (sf && r.find(sf, "a:srgbClr")) out.color = colorOf(r, sf).color;
  return Object.keys(out).length ? out : undefined;
}

export function readFormat(r: Reader, spPr: El | undefined, txPr: El | undefined, size: number): ElementFormat | undefined {
  const s = readShape(r, spPr);
  const t = readText(r, txPr, size);
  const out: ElementFormat = { ...s, ...(t ? { text: t } : {}) };
  return Object.keys(out).length ? out : undefined;
}

export function readAxis(r: Reader, ax: El, isVal: boolean): AxisOptions | undefined {
  const o: AxisOptions = {};
  const sc = r.find(ax, "c:scaling");
  const num = (el: El | undefined) => { const v = r.val(el); return v === undefined ? undefined : Number(v); };
  if (r.val(r.find(sc, "c:orientation")) === "maxMin") o.reverse = true;
  const mx = num(r.find(sc, "c:max")), mn = num(r.find(sc, "c:min")), lb = num(r.find(sc, "c:logBase"));
  if (mx !== undefined) o.max = mx;
  if (mn !== undefined) o.min = mn;
  if (lb !== undefined) o.logBase = lb;
  if (r.val(r.find(ax, "c:delete")) === "1") o.visible = false;
  const nf = r.find(ax, "c:numFmt");
  if (nf && attrs(r.xml, nf).sourceLinked === "0") o.numFmt = attrs(r.xml, nf).formatCode;
  const mt = r.val(r.find(ax, "c:majorTickMark")), nt = r.val(r.find(ax, "c:minorTickMark"));
  if (mt && mt !== "none") o.majorTick = mt as AxisOptions["majorTick"];
  if (nt && nt !== "none") o.minorTick = nt as AxisOptions["minorTick"];
  const lp = r.val(r.find(ax, "c:tickLblPos"));
  if (lp && lp !== "nextTo") o.labelPos = lp as AxisOptions["labelPos"];
  const ca = num(r.find(ax, "c:crossesAt"));
  if (ca !== undefined) o.crosses = ca;
  else if (r.val(r.find(ax, "c:crosses")) === "max") o.crosses = "max";
  if (isVal) {
    const mu = num(r.find(ax, "c:majorUnit")), nu = num(r.find(ax, "c:minorUnit"));
    if (mu !== undefined) o.majorUnit = mu;
    if (nu !== undefined) o.minorUnit = nu;
  }
  const f = readFormat(r, r.find(ax, "c:spPr"), r.find(ax, "c:txPr"), 900);
  if (f?.line) o.line = f.line;
  if (f?.text) o.text = f.text;
  return Object.keys(o).length ? o : undefined;
}
