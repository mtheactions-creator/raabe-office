// Tables on slides (a:tbl in a graphicFrame) with PowerPoint's table styles: the style's parts
// (whole table, banded rows and columns, first and last row and column) in PowerPoint's order of
// precedence, then each cell's own fill and borders. Rows grow to fit their text as PowerPoint's do.

import { PptxDom, readBodyProps, readColor, readText, readThemeInfo, SlideResolver, type ShapeView } from "@raabe/ooxml";
import { EMU_PX, renderText } from "./text";
import type { RenderOptions } from "./slide";

const { kid, kids, num, at, NS_A } = PptxDom;

/** PowerPoint's default table style (Mittlere Formatvorlage 2 - Akzent 1), as it writes it into tableStyles.xml. */
export const DEFAULT_TABLE_STYLE_ID = "{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}";
const ln = (side: string, w = 12700) => `<a:${side}><a:ln w="${w}" cmpd="sng"><a:solidFill><a:schemeClr val="lt1"/></a:solidFill></a:ln></a:${side}>`;
const txt = (b: boolean, c: string) => `<a:tcTxStyle${b ? ' b="on"' : ""}><a:fontRef idx="minor"><a:prstClr val="black"/></a:fontRef><a:schemeClr val="${c}"/></a:tcTxStyle>`;
const fill = (tint?: number) => `<a:fill><a:solidFill><a:schemeClr val="accent1">${tint ? `<a:tint val="${tint}"/>` : ""}</a:schemeClr></a:solidFill></a:fill>`;
export const DEFAULT_TABLE_STYLE_XML = `<a:tblStyle styleId="${DEFAULT_TABLE_STYLE_ID}" styleName="Mittlere Formatvorlage 2 - Akzent 1">` +
  `<a:wholeTbl>${txt(false, "dk1")}<a:tcStyle><a:tcBdr>${["left", "right", "top", "bottom", "insideH", "insideV"].map((s) => ln(s)).join("")}</a:tcBdr>${fill(20000)}</a:tcStyle></a:wholeTbl>` +
  `<a:band1H><a:tcStyle><a:tcBdr/>${fill(40000)}</a:tcStyle></a:band1H><a:band2H><a:tcStyle><a:tcBdr/></a:tcStyle></a:band2H>` +
  `<a:band1V><a:tcStyle><a:tcBdr/>${fill(40000)}</a:tcStyle></a:band1V><a:band2V><a:tcStyle><a:tcBdr/></a:tcStyle></a:band2V>` +
  `<a:lastCol>${txt(true, "lt1")}<a:tcStyle><a:tcBdr/>${fill()}</a:tcStyle></a:lastCol><a:firstCol>${txt(true, "lt1")}<a:tcStyle><a:tcBdr/>${fill()}</a:tcStyle></a:firstCol>` +
  `<a:lastRow>${txt(true, "lt1")}<a:tcStyle><a:tcBdr>${ln("top", 38100)}</a:tcBdr>${fill()}</a:tcStyle></a:lastRow>` +
  `<a:firstRow>${txt(true, "lt1")}<a:tcStyle><a:tcBdr>${ln("bottom", 38100)}</a:tcBdr>${fill()}</a:tcStyle></a:firstRow></a:tblStyle>`;

let builtIn: Element | null = null;

/** The style element for an id: the presentation's tableStyles.xml, else PowerPoint's default. */
function styleOf(o: RenderOptions, id: string | undefined): Element | null {
  const path = o.pptx.related(o.pptx.presPath, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/tableStyles");
  if (path && o.pptx.has(path)) {
    const doc = o.pptx.doc(path);
    const want = id ?? doc.documentElement.getAttribute("def") ?? undefined;
    const hit = kids(doc.documentElement, "tblStyle").find((s) => s.getAttribute("styleId") === want);
    if (hit) return hit;
  }
  if (id && id !== DEFAULT_TABLE_STYLE_ID) return null;
  builtIn ??= PptxDom.parseXml(`<a:tblStyleLst xmlns:a="${NS_A}">${DEFAULT_TABLE_STYLE_XML}</a:tblStyleLst>`).documentElement.firstElementChild;
  return builtIn;
}

interface CellLook { fill?: string; borders: Record<string, { color: string; width: number } | null>; bold?: boolean; color?: string }

function lineOf(lnEl: Element | undefined, colors: Record<string, string>): { color: string; width: number } | null | undefined {
  if (!lnEl) return undefined;
  if (kid(lnEl, "noFill")) return null;
  const c = readColor(kid(lnEl, "solidFill"), colors);
  return c ? { color: c.hex, width: (num(lnEl, "w") ?? 12700) / EMU_PX } : undefined;
}

function applyPart(look: CellLook, part: Element | undefined, colors: Record<string, string>, edges: { top: boolean; bottom: boolean; left: boolean; right: boolean }): void {
  if (!part) return;
  const tx = kid(part, "tcTxStyle");
  if (tx) {
    if (tx.getAttribute("b") === "on") look.bold = true;
    if (tx.getAttribute("b") === "off") look.bold = false;
    const c = readColor(tx, colors);
    if (c) look.color = c.hex;
  }
  const st = kid(part, "tcStyle");
  const f = kid(kid(st, "fill"), "solidFill");
  if (f) look.fill = readColor(f, colors)?.hex;
  if (kid(kid(st, "fill"), "noFill")) look.fill = undefined;
  const bdr = kid(st, "tcBdr");
  const side = (name: string) => lineOf(kid(kid(bdr, name), "ln"), colors);
  const set = (k: string, v: ReturnType<typeof side>) => { if (v !== undefined) look.borders[k] = v; };
  set("top", edges.top ? side("top") : side("insideH"));
  set("bottom", edges.bottom ? side("bottom") : side("insideH"));
  set("left", edges.left ? side("left") : side("insideV"));
  set("right", edges.right ? side("right") : side("insideV"));
}

export function renderTable(s: ShapeView, o: RenderOptions): HTMLElement {
  const tbl = at(s.el, "graphic/graphicData/tbl");
  const box = document.createElement("div");
  box.className = "tbl";
  box.style.cssText = "position:absolute;left:0;top:0";
  if (!tbl) return box;
  const slidePath = o.pptx.slides.find((x) => o.pptx.doc(x.path) === s.el.ownerDocument)?.path;
  const theme = slidePath ? new SlideResolver(o.pptx, slidePath).theme : readThemeInfo(o.pptx, undefined, undefined);
  const pr = kid(tbl, "tblPr");
  const flag = (n: string) => pr?.getAttribute(n) === "1";
  const style = styleOf(o, kid(pr, "tableStyleId")?.textContent ?? undefined);
  const cols = kids(kid(tbl, "tblGrid"), "gridCol").map((c) => (num(c, "w") ?? 0) / EMU_PX);
  const rows = kids(tbl, "tr");
  const table = document.createElement("table");
  table.style.cssText = `border-collapse:collapse;table-layout:fixed;width:${cols.reduce((a, b) => a + b, 0)}px`;
  const cg = document.createElement("colgroup");
  for (const w of cols) { const c = document.createElement("col"); c.style.width = `${w}px`; cg.append(c); }
  table.append(cg);
  const defaults = at(o.pptx.presentation, "defaultTextStyle");
  rows.forEach((tr, ri) => {
    const row = document.createElement("tr");
    row.style.height = `${(num(tr, "h") ?? 0) / EMU_PX}px`;
    const cells = kids(tr, "tc");
    cells.forEach((tc, ci) => {
      if (tc.getAttribute("hMerge") === "1" || tc.getAttribute("vMerge") === "1") return;
      const span = num(tc, "gridSpan") ?? 1, rspan = num(tc, "rowSpan") ?? 1;
      const look: CellLook = { borders: {} };
      const edges = { top: ri === 0, bottom: ri + rspan >= rows.length, left: ci === 0, right: ci + span >= cols.length };
      const part = (n: string) => kid(style ?? undefined, n);
      applyPart(look, part("wholeTbl"), theme.colors, edges);
      const firstRow = flag("firstRow"), lastRow = flag("lastRow"), firstCol = flag("firstCol"), lastCol = flag("lastCol");
      if (flag("bandCol")) { const k = ci - (firstCol ? 1 : 0); if (k >= 0) applyPart(look, part(k % 2 === 0 ? "band1V" : "band2V"), theme.colors, edges); }
      if (flag("bandRow")) { const k = ri - (firstRow ? 1 : 0); if (k >= 0) applyPart(look, part(k % 2 === 0 ? "band1H" : "band2H"), theme.colors, edges); }
      if (lastCol && edges.right) applyPart(look, part("lastCol"), theme.colors, edges);
      if (firstCol && ci === 0) applyPart(look, part("firstCol"), theme.colors, edges);
      if (lastRow && edges.bottom) applyPart(look, part("lastRow"), theme.colors, edges);
      if (firstRow && ri === 0) applyPart(look, part("firstRow"), theme.colors, edges);
      // The cell's own formatting.
      const tcPr = kid(tc, "tcPr");
      const own = kid(tcPr, "solidFill");
      if (own) look.fill = readColor(own, theme.colors)?.hex;
      if (kid(tcPr, "noFill")) look.fill = undefined;
      for (const [side, name] of [["left", "lnL"], ["right", "lnR"], ["top", "lnT"], ["bottom", "lnB"]]) {
        const v = lineOf(kid(tcPr, name), theme.colors);
        if (v !== undefined) look.borders[side] = v;
      }
      const td = document.createElement("td");
      if (span > 1) td.colSpan = span;
      if (rspan > 1) td.rowSpan = rspan;
      const bd = (k: string) => { const b = look.borders[k]; return b ? `${Math.max(b.width, 0.5)}px solid #${b.color}` : "none"; };
      const mar = (n: string, d: number) => (num(tcPr, n) ?? d) / EMU_PX;
      td.style.cssText = `padding:0;vertical-align:top;position:relative;background:${look.fill ? `#${look.fill}` : "transparent"};border-left:${bd("left")};border-right:${bd("right")};border-top:${bd("top")};border-bottom:${bd("bottom")}`;
      const anchor = tcPr?.getAttribute("anchor");
      const body = readBodyProps([kid(kid(tc, "txBody"), "bodyPr")]);
      body.anchor = anchor === "ctr" ? "ctr" : anchor === "b" ? "b" : "t";
      const lstOverride = PptxDom.parseXml(`<a:lstStyle xmlns:a="${NS_A}"><a:lvl1pPr><a:defRPr${look.bold ? ' b="1"' : ""}/></a:lvl1pPr></a:lstStyle>`).documentElement;
      const text = readText(kid(tc, "txBody"), { theme, lists: [defaults, lstOverride, kid(kid(tc, "txBody"), "lstStyle")], fontColor: look.color }, body);
      const w = cols.slice(ci, ci + span).reduce((a, b) => a + b, 0);
      const inner = document.createElement("div");
      const l = mar("marL", 91440), r = mar("marR", 91440), t = mar("marT", 45720), b = mar("marB", 45720);
      inner.style.cssText = `position:relative;padding:${t}px ${r}px ${b}px ${l}px;min-height:${Math.max(0, (num(tr, "h") ?? 0) / EMU_PX - t - b)}px;box-sizing:content-box`;
      const tEl = renderText(text, w - l - r, 0, { prompt: false, slideNumber: o.slideNumber });
      tEl.style.position = "relative";
      tEl.style.height = "auto";
      inner.append(tEl);
      td.append(inner);
      row.append(td);
    });
    table.append(row);
  });
  box.append(table);
  return box;
}
