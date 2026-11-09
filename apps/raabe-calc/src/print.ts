// Printing and PDF as Excel lays sheets out on paper (D-063, D-094): from A1 to the end of the used range and the
// drawing objects (or the Druckbereich), cut into pages first down, then across, at the page's printable size, Skalierung or An
// Format anpassen as Excel's whole percent, Auf der Seite zentrieren, optional gridlines and headings, Drucktitel
// repeated on every page. Columns and rows take Excel's printed sizes (print-metrics.ts); the cells are drawn by
// the grid's own renderer into SVG (print-svg.ts), so the page shows what the screen shows, with real text.

import { MAX_COLS, MAX_ROWS, colName, parseRange, type PageSetup, type Rect, type Sheet } from "@raabe/sheet-engine";
import { h, icon, showDialog, type MenuEntry } from "@raabe/shell";
import { pictureElement, type PictureFormat } from "@raabe/text-engine";
import { pictureBakedKey } from "@raabe/ooxml";
import type { CalcApp } from "./app";
import { CondEval } from "./grid/cond";
import { Axis, sheetAxes } from "./grid/metrics";
import { cssFont, draw, fitCtx, wrapLines, type Pane } from "./grid/render";
import { PRN, printColPx, printRowPx, printerDigit, rowHeights } from "./print-metrics";
import { svgContext } from "./print-svg";

const PAPERS: { code: number; name: string; w: number; h: number }[] = [
  { code: 9, name: "A4", w: 8.27, h: 11.69 },
  { code: 8, name: "A3", w: 11.69, h: 16.54 },
  { code: 11, name: "A5", w: 5.83, h: 8.27 },
  { code: 1, name: "Letter", w: 8.5, h: 11 },
  { code: 5, name: "Legal", w: 8.5, h: 14 },
];

const MARGINS: Record<string, NonNullable<PageSetup["margins"]>> = {
  "Normale Seitenränder": { left: 0.7, right: 0.7, top: 0.75, bottom: 0.75, header: 0.3, footer: 0.3 },
  "Breite Seitenränder": { left: 1, right: 1, top: 1, bottom: 1, header: 0.5, footer: 0.5 },
  "Schmale Seitenränder": { left: 0.25, right: 0.25, top: 0.75, bottom: 0.75, header: 0.3, footer: 0.3 },
};

/** The renderer's device scale for paper: Excel lays pages out at 600 dpi. */
const PRINT_DPR = 600 / 96;

const NL = String.fromCharCode(10);

/**
 * The page as the browser's PDF can hold it without shrinking: it lays paper out in 1/300 inch and makes A4
 * 594,96 x 841,92 pt; a page even slightly larger is scaled down (by 0,9992 for 8,27 in) and moved, so the
 * pages are asked for two such units smaller than the paper (recorded, D-094). Everything is placed from the
 * top left, so nothing moves.
 */
export function sheetSize(sz: { w: number; h: number }): { w: number; h: number } {
  return { w: (Math.floor(sz.w * 300 + 1e-6) - 2) / 300, h: (Math.floor(sz.h * 300 + 1e-6) - 2) / 300 };
}

const cm = (inch: number) => `${(inch * 2.54).toFixed(1).replace(".", ",")} cm`;

type Page = { rows: [number, number]; cols: [number, number]; index?: number; count?: number };

/** A sheet's printed geometry: Excel's sizes on paper (page px at 96 dpi, before the zoom). */
interface Geometry {
  rows: Axis;
  cols: Axis;
  /** The rows' and columns' own printed sizes (the axes' exceptions), for axes that end with a page. */
  own: { rows: Map<number, number>; cols: Map<number, number> };
  /** The screen's sizes (CSS px at 100 %), where drawing objects keep their offsets. */
  screen: { rows: Axis; cols: Axis };
  /** Skalierung as a fraction (An Format anpassen: Excel's whole percent). */
  zoom: number;
  area: Rect | null;
  /** Printable width and height (page px). */
  pw: number;
  ph: number;
}

/**
 * Kopf- und Fußzeile: Excel's codes split into left, centre and right with the fields filled in
 * (&P page, &N pages, &D date, &T time, &F file, &A sheet, &Z path; format codes are left out).
 */
export function headerSections(code: string, f: { page: number; pages: number; sheet: string; file: string; now?: Date }): { L: string; C: string; R: string } {
  const out = { L: "", C: "", R: "" };
  let sec: "L" | "C" | "R" = "C";
  const now = f.now ?? new Date();
  const two = (n: number) => String(n).padStart(2, "0");
  for (let i = 0; i < code.length; i++) {
    const ch = code[i];
    if (ch !== "&") { out[sec] += ch; continue; }
    const k = code[++i];
    if (k === undefined) break;
    if (k === "&") out[sec] += "&";
    else if (k === "L" || k === "C" || k === "R") sec = k;
    else if (k === "P") out[sec] += String(f.page);
    else if (k === "N") out[sec] += String(f.pages);
    else if (k === "D") out[sec] += `${two(now.getDate())}.${two(now.getMonth() + 1)}.${now.getFullYear()}`;
    else if (k === "T") out[sec] += `${two(now.getHours())}:${two(now.getMinutes())}`;
    else if (k === "F") out[sec] += f.file;
    else if (k === "A") out[sec] += f.sheet;
    else if (k === '"') { const j = code.indexOf('"', i + 1); i = j < 0 ? code.length : j; }
    else if (/\d/.test(k)) { while (/\d/.test(code[i + 1] ?? "")) i++; }
    else if (k === "K") i += 6;
  }
  return out;
}

/** The three sections of a header code, codes kept (for editing). */
function splitSections(code: string): { L: string; C: string; R: string } {
  const out = { L: "", C: "", R: "" };
  let sec: "L" | "C" | "R" = "C";
  for (let i = 0; i < code.length; i++) {
    if (code[i] === "&" && /[LCR]/.test(code[i + 1] ?? "")) { sec = code[++i] as "L" | "C" | "R"; continue; }
    if (code[i] === "&" && code[i + 1] === "&") { out[sec] += "&&"; i++; continue; }
    out[sec] += code[i];
  }
  return out;
}

/** Excel's German field names in the Kopf-/Fußzeile fields ⇄ its codes. */
const FIELDS: [string, string][] = [["&[Seite]", "&P"], ["&[Seiten]", "&N"], ["&[Datum]", "&D"], ["&[Zeit]", "&T"], ["&[Pfad]", "&Z"], ["&[Datei]", "&F"], ["&[Register]", "&A"]];
const toFields = (s: string) => FIELDS.reduce((t, [f, c]) => t.split(c).join(f), s);
const toCodes = (s: string) => FIELDS.reduce((t, [f, c]) => t.split(f).join(c), s);

/** "$1:$2" / "1" → rows, "$A:$B" / "A" → columns (Wiederholungszeilen/-spalten). */
function readSpan(text: string, rows: boolean): [number, number] | undefined {
  const t = text.trim().replace(/^.*!/, "");
  if (!t) return undefined;
  const r = parseRange(t.includes(":") ? t : `${t}:${t}`);
  if (!r) return undefined;
  return rows ? (r.c0 === 0 && r.c1 === MAX_COLS - 1 ? [r.r0, r.r1] : undefined) : (r.r1 - r.r0 > 1000000 ? [r.c0, r.c1] : undefined);
}

/** A drawing object's anchors (pictures, charts, shapes) and how it is drawn on paper. */
type Anchor = { col: number; row: number; x: number; y: number };
type Drawing = { from: Anchor; to?: Anchor; size?: { w: number; h: number }; kind: "picture" | "chart" | "shape"; obj: unknown; rot?: number };

/**
 * Excel's fit to pages: the largest whole percent (at most 100, at least 10) at which the area takes no more
 * pages across and down than asked for (FitToPagesWide / FitToPagesTall; 0 or none: no limit that way).
 */
export function fitZoom(total: { w: number; h: number }, room: { w: number; h: number }, pages: { wide?: number; tall?: number }): number {
  let z = 100;
  if (pages.wide) z = Math.min(z, Math.floor((100 * pages.wide * room.w) / total.w + 1e-9));
  if (pages.tall) z = Math.min(z, Math.floor((100 * pages.tall * room.h) / total.h + 1e-9));
  return Math.max(10, z);
}

/** Cuts indices into pages of at most `room` (sizes add up in order; manual breaks start a page). */
export function cutPages(from: number, to: number, size: (i: number) => number, room: number, breaks?: number[], titleRoom = 0, title?: [number, number]): [number, number][] {
  const out: [number, number][] = [];
  const manual = new Set(breaks ?? []);
  let start = from, used = 0;
  const roomFor = (s: number) => (title && s > title[1] ? room - titleRoom : room);
  for (let i = from; i <= to; i++) {
    const s = size(i);
    if (i > start && (manual.has(i) || used + s > roomFor(start) + 1e-6)) { out.push([start, i - 1]); start = i; used = 0; }
    used += s;
  }
  out.push([start, to]);
  return out;
}

export function printFeature(app: CalcApp) {
  const page = (): PageSetup => app.sheet.page ?? {};
  const setPage = (label: string, fn: (p: PageSetup) => void) => app.change(label, () => {
    app.history.whole(app.sheet);
    const p = { ...(app.sheet.page ?? {}) };
    fn(p);
    app.sheet.page = p;
  });

  function paper(p: PageSetup) {
    const sz = PAPERS.find((x) => x.code === (p.paperSize ?? 9)) ?? PAPERS[0];
    return p.orientation === "landscape" ? { ...sz, w: sz.h, h: sz.w } : sz;
  }

  /** Pictures, charts and shapes of a sheet with their anchors (hidden ones are not printed). */
  function drawings(sheet: Sheet): Drawing[] {
    const out: Drawing[] = [];
    for (const pic of sheet.pictures ?? []) if (!pic.hidden) out.push({ from: pic.from, to: pic.to, size: pic.size, kind: "picture", obj: pic, rot: (pic.format as PictureFormat | undefined)?.rotation });
    for (const ch of sheet.charts ?? []) out.push({ from: ch.from, to: ch.to, kind: "chart", obj: ch });
    for (const s of sheet.shapes ?? []) if (!(s as { hidden?: boolean }).hidden) out.push({ from: s.from, to: s.to, kind: "shape", obj: s, rot: s.rot });
    return out;
  }

  /** Where a drawing ends (its bottom-right anchor; one-cell anchors through the screen's sizes). */
  function endOf(d: Drawing, screen: { rows: Axis; cols: Axis }): Anchor {
    if (d.to) return d.to;
    const x = screen.cols.pos(d.from.col) + d.from.x + (d.size?.w ?? 0), y = screen.rows.pos(d.from.row) + d.from.y + (d.size?.h ?? 0);
    const col = screen.cols.at(x), row = screen.rows.at(y);
    return { col, row, x: x - screen.cols.pos(col), y: y - screen.rows.pos(row) };
  }

  /** Excel's printed sizes of a sheet's rows and columns, the printed area and the zoom. */
  function geometry(sheet: Sheet): Geometry {
    const p = sheet.page ?? {};
    const screen = sheetAxes(sheet);
    const std = app.wb.styles.get(undefined).font;
    const digit = printerDigit(std.name, std.size);
    const rowStd = rowHeights(std.name, std.size, sheet.defaultRowHeight);
    const cols = new Axis(printColPx(screen.cols.def, digit), MAX_COLS);
    const cx = new Map<number, number>();
    for (const [c] of sheet.cols) {
      const w = printColPx(screen.cols.size(c), digit);
      if (w !== cols.def) cx.set(c, w);
    }
    cols.set(cx, cols.def);
    const rows = new Axis(printRowPx(undefined, rowStd), MAX_ROWS);
    const rx = new Map<number, number>();
    for (const [r, info] of sheet.rows) {
      const hgt = info.hidden ? 0 : printRowPx(info.height, rowStd);
      if (hgt !== rows.def) rx.set(r, hgt);
    }
    // Rows with wrapped text and no height of their own are sized again (Excel does so when it opens a file).
    for (const [r, hgt] of autoRows(sheet, screen.cols, rowStd)) if (!sheet.rows.get(r)?.hidden) rx.set(r, hgt);
    rows.set(rx, rows.def);
    // Without a Druckbereich: the used cells and the cells under drawing objects (Excel prints those too).
    let area: Rect | null = selectionArea ?? p.printArea ?? sheet.usedRange();
    if (!p.printArea && !selectionArea) for (const d of drawings(sheet)) {
      const end = endOf(d, screen);
      const r = { r0: d.from.row, c0: d.from.col, r1: end.row, c1: end.col };
      area = area ? { r0: Math.min(area.r0, r.r0), c0: Math.min(area.c0, r.c0), r1: Math.max(area.r1, r.r1), c1: Math.max(area.c1, r.c1) } : r;
    }
    // Excel prints from A1, whatever cell the used range starts at (recorded: a table from B2 keeps column A and row 1).
    if (area && !p.printArea && !selectionArea) area = { ...area, r0: 0, c0: 0 };
    const sz = paper(p);
    const m = p.margins ?? MARGINS["Normale Seitenränder"];
    const pw = (sz.w - m.left - m.right) * 96, ph = (sz.h - m.top - m.bottom) * 96;
    let zoom = (p.scale ?? 100) / 100;
    if (area && (p.fitWidth || p.fitHeight)) {
      const tw = cols.pos(area.c1 + 1) - cols.pos(area.c0), th = rows.pos(area.r1 + 1) - rows.pos(area.r0);
      zoom = fitZoom({ w: tw, h: th }, { w: pw, h: ph }, { wide: p.fitWidth, tall: p.fitHeight }) / 100;
    }
    return { rows, cols, own: { rows: rx, cols: cx }, screen, zoom, area, pw, ph };
  }

  /**
   * Printed heights (page px) of rows with wrapped text that Excel sizes by their content: Excel sizes them again
   * when it opens a file (the file's height can be from other column widths), as lines at the screen's widths
   * times the font's screen line height, and prints that as a row of that height (recorded: five wrapped lines
   * in Aptos Narrow 11 print 76,2 pt where the file says 130,5 pt). Other rows print as their height says.
   */
  function autoRows(sheet: Sheet, cols: Axis, std: { printed: number; screen: number }): Map<number, number> {
    const out = new Map<number, number>();
    const auto = new Set<number>();
    // Rows grown for thick borders (thickTop/thickBot) keep their height.
    for (const [r, info] of sheet.rows) if (info.height !== undefined && !info.custom && !info.hidden && !/thick(Top|Bot)="(1|true)"/.test(info.raw ?? "")) auto.add(r);
    if (!auto.size) return out;
    const need = new Map<number, number>();
    const wrapped = new Set<number>();
    const keep = new Set<number>();
    const ctx = fitCtx();
    for (const [k, cell] of sheet.cells) {
      const row = Math.floor(k / MAX_COLS);
      if (!auto.has(row)) continue;
      const col = k % MAX_COLS;
      const merge = sheet.mergeAt(row, col);
      if (merge && merge.r0 !== merge.r1) continue;
      const st = app.wb.styles.get(cell.s ?? sheet.rows.get(row)?.s ?? sheet.cols.get(col)?.s);
      if (st.rotation) { keep.add(row); continue; }
      const lineH = rowHeights(st.font.name, st.font.size, (st.font.size * 15) / 11).screen;
      let lines = 1;
      if (typeof cell.v === "string" && cell.v !== "") {
        if (st.wrap && !merge) {
          ctx.font = cssFont(st, 1);
          lines = cell.v.split(NL).reduce((n, para) => n + Math.max(1, wrapLines(ctx, para, cols.size(col) - 5).length), 0);
          wrapped.add(row);
        } else lines = cell.v.split(NL).length;
      } else if (cell.v === null || cell.v === "") {
        if (st.font.size <= 11) continue;
      }
      need.set(row, Math.max(need.get(row) ?? 0, lines * lineH));
    }
    for (const r of wrapped) {
      if (keep.has(r)) continue;
      out.set(r, printRowPx(Math.max(std.screen, need.get(r) ?? 0), std));
    }
    return out;
  }

  /** The pages of a sheet: row and column ranges, Excel's order (down, then across). */
  function paginate(sheet: Sheet, g = geometry(sheet)): Page[] {
    const p = sheet.page ?? {};
    const area = g.area;
    if (!area) return [];
    const span = (t: [number, number] | undefined, ax: Axis) => (t ? ax.pos(t[1] + 1) - ax.pos(t[0]) : 0);
    // Zeilen- und Spaltenüberschriften take their room on every page.
    const hw = p.headings ? 30 : 0, hh = p.headings ? g.rows.def : 0;
    const colPages = cutPages(area.c0, area.c1, (c) => g.cols.size(c), g.pw / g.zoom - hw, p.colBreaks, span(p.titleCols, g.cols), p.titleCols);
    const rowPages = cutPages(area.r0, area.r1, (r) => g.rows.size(r), g.ph / g.zoom - hh, p.rowBreaks, span(p.titleRows, g.rows), p.titleRows);
    const pages: Page[] = [];
    for (const cols of colPages) for (const rows of rowPages) pages.push({ rows, cols });
    pages.forEach((pg, i) => { pg.index = i; pg.count = pages.length; });
    return pages;
  }

  /** One page at its printed size: the cells drawn by the grid's renderer, then the drawing objects. */
  function renderPage(sheet: Sheet, pg: Page, g: Geometry): HTMLElement {
    const p = sheet.page ?? {};
    const sz = paper(p);
    const m = p.margins ?? MARGINS["Normale Seitenränder"];
    const z = g.zoom;
    const needRows = p.titleRows && pg.rows[0] > p.titleRows[1] ? p.titleRows : undefined;
    const needCols = p.titleCols && pg.cols[0] > p.titleCols[1] ? p.titleCols : undefined;
    const ext = (t: [number, number] | undefined, ax: Axis) => (t ? (ax.pos(t[1] + 1) - ax.pos(t[0])) * z : 0);
    const th = ext(needRows, g.rows), tw = ext(needCols, g.cols);
    const head = p.headings ? { w: 30 * z, h: (g.rows.def) * z } : { w: 0, h: 0 };
    const bw = (g.cols.pos(pg.cols[1] + 1) - g.cols.pos(pg.cols[0])) * z, bh = (g.rows.pos(pg.rows[1] + 1) - g.rows.pos(pg.rows[0])) * z;
    // Printed gridlines: a line in the first printer pixel of every cell and one after the last (the area grows by it).
    const grid = !!p.gridLines, line = grid ? PRN : 0;
    const W = head.w + tw + bw + line, H = head.h + th + bh + line;
    // Lines on the area's edges reach a little beyond it (Excel's thick lines: 8 printer pixels before the edge,
    // 16 after): the drawing has room around the cells, and the axes end with the page's last row and column.
    const L = 2, E = 3;
    const rowsAx = new Axis(g.rows.def, pg.rows[1] + 1), colsAx = new Axis(g.cols.def, pg.cols[1] + 1);
    rowsAx.set(g.own.rows, g.rows.def);
    colsAx.set(g.own.cols, g.cols.def);
    const panes: Pane[] = [];
    const x0 = L + head.w, y0 = L + head.h;
    const eps = 1e-3;
    if (needRows && needCols) panes.push({ row: needRows[0], col: needCols[0], x: x0, y: y0, w: tw - eps, h: th - eps });
    if (needRows) panes.push({ row: needRows[0], col: pg.cols[0], x: x0 + tw, y: y0, w: bw + E, h: th - eps });
    if (needCols) panes.push({ row: pg.rows[0], col: needCols[0], x: x0, y: y0 + th, w: tw - eps, h: bh + E });
    panes.push({ row: pg.rows[0], col: pg.cols[0], x: x0 + tw, y: y0 + th, w: bw + E, h: bh + E });
    const SW = L + W + E, SH = L + H + E;
    const { ctx, svg } = svgContext(SW * PRINT_DPR, SH * PRINT_DPR, `p${pg.index ?? 0}`, true);
    svg.setAttribute("width", String(SW));
    svg.setAttribute("height", String(SH));
    draw(ctx, {
      sheet, styles: app.wb.styles, rows: rowsAx, cols: colsAx, zoom: z, dpr: PRINT_DPR, width: SW, height: SH, headerW: head.w, headerH: head.h, panes,
      selection: [], active: { row: -1, col: -1 }, showGrid: !!p.gridLines, showHeaders: !!p.headings, textCache: new Map(),
      cond: sheet.conditional?.length ? new CondEval(app.wb, sheet, app.wb.sheets.indexOf(sheet)) : null,
      tableStyle: sheet.tables?.length && sheet === app.sheet ? (r, c, st) => app.tables.styleFor(r, c, st) : undefined, book: app.wb, print: true,
      // Excel's text sits 6 printer pixels further right and 3 lower than the screen's padding puts it; printed
      // gridlines move the cells by 4 pixels but not the text (recorded, D-094).
      textShift: { x: (grid ? 2 : 6) * PRN, y: (grid ? -1 : 3) * PRN },
    });
    if (grid) {
      // Excel closes the grid and frames the printed area with a line of eight printer pixels around its edge.
      ctx.fillStyle = "#000000";
      const x1 = Math.round((L + W - line) * PRINT_DPR) / PRINT_DPR, y1 = Math.round((L + H - line) * PRINT_DPR) / PRINT_DPR;
      const f = 4 * PRN;
      ctx.fillRect(x1, y0, line, y1 - y0);
      ctx.fillRect(x0, y1, x1 - x0 + line, line);
      ctx.fillRect(x0 - f, y0 - f, x1 - x0 + 2 * f, 2 * f);
      ctx.fillRect(x0 - f, y1 - f, x1 - x0 + 2 * f, 2 * f);
      ctx.fillRect(x0 - f, y0 - f, 2 * f, y1 - y0 + 2 * f);
      ctx.fillRect(x1 - f, y0 - f, 2 * f, y1 - y0 + 2 * f);
    }
    // Auf der Seite zentrieren.
    // Excel starts printed gridlines four printer pixels inside the margins and centres the area with the
    // lines around it (recorded); its places are whole printer pixels. The browser puts the drawing at whole
    // CSS pixels, so the rest of the way is inside the drawing (lx, ly).
    const offX = (p.centerH ? Math.max(0, (g.pw - W - 16 * PRN) / 2) : 0) + (grid ? 4 * PRN : 0), offY = (p.centerV ? Math.max(0, (g.ph - H - 16 * PRN) / 2) : 0) + (grid ? 4 * PRN : 0);
    const ox = Math.round((m.left * 96 + offX) / PRN) * PRN, oy = Math.round((m.top * 96 + offY) / PRN) * PRN;
    const bx = Math.floor(ox - L), by = Math.floor(oy - L);
    svg.setAttribute("viewBox", `${(L - (ox - bx)) * PRINT_DPR} ${(L - (oy - by)) * PRINT_DPR} ${SW * PRINT_DPR} ${SH * PRINT_DPR}`);
    const body = h("div", { style: `position:absolute;left:${bx}px;top:${by}px;width:${SW}px;height:${SH}px;overflow:hidden` }, svg);
    svg.style.position = "absolute";
    svg.style.left = svg.style.top = "0";
    const out = sheetSize(sz);
    const el = h("div", { class: "print-page", style: `position:relative;width:${out.w}in;height:${out.h}in` }, body);
    // Drawing objects over the cells, at their cells' printed places (a part on each page they reach).
    const X = (a: Anchor) => (ox - bx) + head.w + tw + ((g.cols.pos(a.col) - g.cols.pos(pg.cols[0])) + a.x * ratio(g.cols, g.screen.cols, a.col)) * z;
    const Y = (a: Anchor) => (oy - by) + head.h + th + ((g.rows.pos(a.row) - g.rows.pos(pg.rows[0])) + a.y * ratio(g.rows, g.screen.rows, a.row)) * z;
    for (const d of drawings(sheet)) {
      const end = endOf(d, g.screen);
      if (end.row < pg.rows[0] || d.from.row > pg.rows[1] || end.col < pg.cols[0] || d.from.col > pg.cols[1]) continue;
      const x = X(d.from), y = Y(d.from), w = Math.max(1, X(end) - x), hh = Math.max(1, Y(end) - y);
      const holder = h("div", { style: `position:absolute;left:${x}px;top:${y}px;width:${w}px;height:${hh}px;${d.rot && d.kind !== "shape" ? `transform:rotate(${d.rot}deg);` : ""}` });
      const look = h("div", { style: `position:absolute;left:0;top:0;width:${w / z}px;height:${hh / z}px;transform:scale(${z});transform-origin:0 0` });
      if (d.kind === "chart") look.append(app.charts.printSvg(d.obj as Parameters<typeof app.charts.printSvg>[0], w / z, hh / z));
      else if (d.kind === "shape") look.append(app.shapes.printElement(d.obj as Parameters<typeof app.shapes.printElement>[0], w / z, hh / z));
      else look.append(pictureLook(d.obj as NonNullable<Sheet["pictures"]>[number], w / z, hh / z));
      holder.append(look);
      body.append(holder);
    }
    // Kopf- und Fußzeile at Excel's header and footer distances, between the side margins (Excel's PDF: the
    // text 0,96 pt inside the margins, the header's 1,48 pt below its distance, the footer's 1,5 pt above).
    const info = { page: (pg.index ?? 0) + 1, pages: pg.count ?? 1, sheet: sheet.name, file: document.title.replace(/ - RaabeCalc$/, "") };
    const std = app.wb.styles.get(undefined).font;
    const band = (code: string | undefined, top: boolean) => {
      if (!code) return;
      const s = headerSections(code, info);
      el.append(h("div", { class: "print-hf", style: `left:calc(${m.left}in + 0.96pt);right:calc(${m.right}in + 0.96pt);font-family:"${std.name}",Carlito,sans-serif;${top ? `top:calc(${m.header}in + 1.48pt)` : `bottom:calc(${m.footer}in + 1.5pt)`}` },
        h("span", { style: "text-align:left" }, s.L), h("span", { style: "text-align:center" }, s.C), h("span", { style: "text-align:right" }, s.R)));
    };
    band(p.header, true);
    band(p.footer, false);
    return el;
  }

  /** How much longer a row or column is on paper than on screen (offsets inside it scale alike). */
  function ratio(print: Axis, screen: Axis, i: number): number {
    const s = screen.size(i);
    return s > 0 ? print.size(i) / s : 1;
  }

  /** A picture with its Bildformat (D-087) at a size in px at 100 %. */
  function pictureLook(pic: NonNullable<Sheet["pictures"]>[number], w: number, hh: number): HTMLElement {
    let bx = 0, by = 0, bw = w, bh = hh;
    if (pic.group) { bx = pic.group.x * w; by = pic.group.y * hh; bw = pic.group.w * w; bh = pic.group.h * hh; }
    const f = (pic.format ?? {}) as PictureFormat;
    const baked = pic.baked !== undefined && pictureBakedKey(f) === pic.baked;
    const el = pictureElement(baked ? pic.bakedSrc ?? pic.src : pic.src, Math.max(1, bw), Math.max(1, bh),
      { ...f, rotation: undefined, ...(baked ? { brightness: undefined, contrast: undefined, saturation: undefined, temperature: undefined, sharpen: undefined } : {}) });
    return h("div", { style: `position:absolute;left:${bx}px;top:${by}px;width:${bw}px;height:${bh}px` }, el);
  }

  /** Datei > Drucken's choices for this printing (not saved, as in Excel): what and which pages. */
  const job: { what: "active" | "book" | "selection"; from?: number; to?: number } = { what: "active" };
  /** "Auswahl drucken": the selection stands in for the Druckbereich while the pages are made. */
  let selectionArea: Rect | null = null;

  function pages(): HTMLElement[] {
    const sheets = job.what === "book" ? app.wb.sheets.filter((s) => !s.hidden) : [app.sheet];
    const out: HTMLElement[] = [];
    for (const sheet of sheets) {
      selectionArea = job.what === "selection" && sheet === app.sheet ? app.grid.selection.at(-1)! : null;
      try {
        const g = geometry(sheet);
        out.push(...paginate(sheet, g).map((pg) => renderPage(sheet, pg, g)));
      } finally {
        selectionArea = null;
      }
    }
    // Seiten von … bis: the pages of the whole job, counted from 1.
    const from = Math.max(1, job.from ?? 1), to = Math.min(out.length, job.to ?? out.length);
    return from > 1 || to < out.length ? out.slice(from - 1, Math.max(from - 1, to)) : out;
  }

  async function exportPdf(to?: string): Promise<void> {
    const path = to ?? await app.platform.saveFileDialog(`${document.title.replace(/ - RaabeCalc$/, "")}.pdf`, "pdf");
    if (!path) return;
    const root = document.getElementById("print-root") ?? document.body.appendChild(h("div", { id: "print-root" }));
    root.replaceChildren(...pages());
    const sz = sheetSize(paper(page()));
    // The page size for the PDF (the shell's print style sheet would otherwise make it A4 portrait).
    const style = h("style", null, `@page { size: ${sz.w}in ${sz.h}in; margin: 0 }`);
    document.head.append(style);
    document.body.classList.add("printing");
    try {
      await app.platform.printToPdf(path, { width: sz.w, height: sz.h });
    } finally {
      document.body.classList.remove("printing");
      root.replaceChildren();
      style.remove();
    }
  }

  function printNow(): void {
    const root = document.getElementById("print-root") ?? document.body.appendChild(h("div", { id: "print-root" }));
    root.replaceChildren(...pages());
    const sz = sheetSize(paper(page()));
    const style = h("style", null, `@page { size: ${sz.w}in ${sz.h}in; margin: 0 }`);
    document.head.append(style);
    document.body.classList.add("printing");
    addEventListener("afterprint", () => { document.body.classList.remove("printing"); root.replaceChildren(); style.remove(); }, { once: true });
    print();
  }

  /** Datei > Drucken: settings left, preview right. */
  function printPage(): HTMLElement {
    const list = pages();
    let i = 0;
    const preview = h("div", { class: "print-preview" });
    const label = h("span");
    const show = () => {
      preview.replaceChildren();
      const pg = list[i];
      if (!pg) { preview.append(h("p", { class: "muted" }, "Es gibt nichts zu drucken.")); label.textContent = "0 von 0"; return; }
      const sz = paper(page());
      const scale = Math.min(0.7, (innerHeight - 260) / (sz.h * 96), 560 / (sz.w * 96));
      pg.style.transform = `scale(${scale})`;
      pg.style.transformOrigin = "top left";
      preview.append(h("div", { class: "sheet-holder", style: `width:${sz.w * 96 * scale}px;height:${sz.h * 96 * scale}px;overflow:hidden;background:#fff;box-shadow:0 1px 4px rgba(0,0,0,.25)` }, pg));
      label.textContent = `${i + 1} von ${list.length}`;
    };
    show();
    const p = page();
    const orient = h("select", { onchange: (e: Event) => { setOrientation((e.target as HTMLSelectElement).value as "portrait" | "landscape"); app.commands.run("file.print"); } },
      h("option", { value: "portrait", selected: p.orientation !== "landscape" || undefined }, "Hochformat"), h("option", { value: "landscape", selected: p.orientation === "landscape" || undefined }, "Querformat"));
    const paperSel = h("select", { onchange: (e: Event) => { setPage("Format", (x) => { x.paperSize = Number((e.target as HTMLSelectElement).value); }); app.commands.run("file.print"); } },
      ...PAPERS.map((x) => h("option", { value: String(x.code), selected: (p.paperSize ?? 9) === x.code || undefined }, `${x.name} (${(x.w * 2.54).toFixed(1).replace(".", ",")} x ${(x.h * 2.54).toFixed(1).replace(".", ",")} cm)`)));
    // Excel's other settings on this page (D-142): what to print and which pages (for this printing only),
    // Seitenränder and Skalierung (the sheet's page setup, as on the Seitenlayout tab).
    const refresh = () => void app.commands.run("file.print");
    const whatSel = h("select", { "aria-label": "Druckbereich", onchange: (e: Event) => { job.what = (e.target as HTMLSelectElement).value as typeof job.what; refresh(); } },
      ...([["active", "Aktive Blätter drucken"], ["book", "Gesamte Arbeitsmappe drucken"], ["selection", "Auswahl drucken"]] as const)
        .map(([v, l]) => h("option", { value: v, selected: job.what === v || undefined }, l)));
    const pageNo = (v: number | undefined, set: (n: number | undefined) => void, label: string) => {
      const input = h("input", { type: "number", min: "1", value: v ? String(v) : "", style: "width:52px", "aria-label": label }) as HTMLInputElement;
      input.onchange = () => { const n = Math.round(Number(input.value)); set(n >= 1 ? n : undefined); refresh(); };
      return input;
    };
    const fromTo = h("div", { class: "print-pages", style: "display:flex;gap:6px;align-items:center" }, h("span", null, "Seiten:"), pageNo(job.from, (n) => { job.from = n; }, "Seiten von"),
      h("span", null, "bis"), pageNo(job.to, (n) => { job.to = n; }, "Seiten bis"));
    const curMargins = Object.entries(MARGINS).find(([, mm]) => JSON.stringify(mm) === JSON.stringify(p.margins ?? MARGINS["Normale Seitenränder"]))?.[0] ?? "";
    const marginSel = h("select", { "aria-label": "Seitenränder", onchange: (e: Event) => {
      const v = (e.target as HTMLSelectElement).value;
      if (v === "custom") return void pageSetupDialog().then(refresh);
      setPage("Seitenränder", (x) => { x.margins = { ...MARGINS[v] }; });
      refresh();
    } }, ...Object.keys(MARGINS).map((k) => h("option", { value: k, selected: curMargins === k || undefined }, k)),
      h("option", { value: "custom", selected: !curMargins || undefined }, "Benutzerdefinierte Seitenränder"));
    const fit = p.fitWidth === 1 && p.fitHeight === 1 ? "sheet" : p.fitWidth === 1 && !p.fitHeight ? "cols" : !p.fitWidth && p.fitHeight === 1 ? "rows" : !p.fitWidth && !p.fitHeight && (p.scale ?? 100) === 100 ? "none" : "custom";
    const scaleSel = h("select", { "aria-label": "Skalierung", onchange: (e: Event) => {
      const v = (e.target as HTMLSelectElement).value;
      if (v === "custom") return void pageSetupDialog().then(refresh);
      setPage("Skalierung", (x) => {
        x.scale = undefined;
        x.fitWidth = v === "sheet" || v === "cols" ? 1 : undefined;
        x.fitHeight = v === "sheet" || v === "rows" ? 1 : undefined;
      });
      refresh();
    } }, ...([["none", "Keine Skalierung"], ["sheet", "Blatt auf einer Seite darstellen"], ["cols", "Alle Spalten auf einer Seite darstellen"], ["rows", "Alle Zeilen auf einer Seite darstellen"], ["custom", "Benutzerdefinierte Skalierungsoptionen..."]] as const)
      .map(([v, l]) => h("option", { value: v, selected: fit === v || undefined }, l)));
    return h("div", { class: "print-layout" },
      h("div", { class: "print-settings" }, h("h1", null, "Drucken"),
        h("button", { class: "bigbtn", onclick: printNow }, icon("print", "l"), "Drucken"),
        h("h2", null, "Einstellungen"), whatSel, fromTo, orient, paperSel, marginSel, scaleSel,
        h("button", { onclick: () => void pageSetupDialog() }, "Seite einrichten")),
      h("div", { class: "print-right" }, preview,
        h("div", { class: "print-nav" }, h("button", { onclick: () => { if (i > 0) { i--; show(); } } }, "◀"), label, h("button", { onclick: () => { if (i < list.length - 1) { i++; show(); } } }, "▶"))));
  }

  function setOrientation(o: "portrait" | "landscape"): void {
    setPage("Ausrichtung", (p) => { p.orientation = o; });
  }

  async function pageSetupDialog(focusTitles: boolean | "header" = false): Promise<void> {
    const p = page();
    const m = { ...(p.margins ?? MARGINS["Normale Seitenränder"]) };
    const inputs = (["top", "bottom", "left", "right"] as const).map((k) => {
      const input = h("input", { type: "text", value: (m[k] * 2.54).toFixed(1).replace(".", ","), style: "width:60px" }) as HTMLInputElement;
      return { k, input, el: h("label", { class: "field" }, { top: "Oben:", bottom: "Unten:", left: "Links:", right: "Rechts:" }[k], input) };
    });
    const scale = h("input", { type: "text", value: String(p.scale ?? 100), style: "width:60px" }) as HTMLInputElement;
    const grid = h("input", { type: "checkbox", checked: p.gridLines || undefined }) as HTMLInputElement;
    const heads = h("input", { type: "checkbox", checked: p.headings || undefined }) as HTMLInputElement;
    const centerH = h("input", { type: "checkbox", checked: p.centerH || undefined }) as HTMLInputElement;
    const centerV = h("input", { type: "checkbox", checked: p.centerV || undefined }) as HTMLInputElement;
    const span = (t: [number, number] | undefined, rows: boolean) => (t ? (rows ? `$${t[0] + 1}:$${t[1] + 1}` : `$${colName(t[0])}:$${colName(t[1])}`) : "");
    const titleRows = h("input", { type: "text", value: span(p.titleRows, true), style: "width:120px" }) as HTMLInputElement;
    const titleCols = h("input", { type: "text", value: span(p.titleCols, false), style: "width:120px" }) as HTMLInputElement;
    // Kopfzeile and Fußzeile: the three sections of Excel's Benutzerdefinierte Kopfzeile, with its field names.
    const hfInputs = (code: string | undefined) => {
      const s = code ? splitSections(code) : { L: "", C: "", R: "" };
      return (["L", "C", "R"] as const).map((k) => h("input", { type: "text", value: toFields(s[k]), style: "width:100%", "aria-label": { L: "Linker Abschnitt", C: "Mittlerer Abschnitt", R: "Rechter Abschnitt" }[k] }) as HTMLInputElement);
    };
    const head = hfInputs(p.header), foot = hfInputs(p.footer);
    const hfRow = (label: string, inputs: HTMLInputElement[]) => h("div", null, h("div", { class: "fc-label" }, label), h("div", { class: "hf-row" }, ...inputs));
    if (focusTitles === "header") setTimeout(() => head[1].focus());
    const body = h("div", null, h("div", { class: "fc-label" }, "Seitenränder (cm)"), ...inputs.map((x) => x.el),
      h("div", { class: "fc-label" }, "Auf der Seite zentrieren"),
      h("label", { class: "check" }, centerH, "Horizontal"), h("label", { class: "check" }, centerV, "Vertikal"),
      h("label", { class: "field" }, "Verkleinern/Vergrößern (%):", scale),
      h("div", { class: "fc-label" }, "Drucktitel"),
      h("label", { class: "field" }, "Wiederholungszeilen oben:", titleRows), h("label", { class: "field" }, "Wiederholungsspalten links:", titleCols),
      hfRow("Kopfzeile", head), hfRow("Fußzeile", foot),
      h("p", { class: "muted" }, "Felder: &[Seite], &[Seiten], &[Datum], &[Zeit], &[Datei], &[Register]"),
      h("label", { class: "check" }, grid, "Gitternetzlinien drucken"), h("label", { class: "check" }, heads, "Zeilen- und Spaltenüberschriften drucken"));
    if (focusTitles === true) setTimeout(() => titleRows.focus());
    const ok = await showDialog({ title: "Seite einrichten", body, buttons: [{ label: "OK", value: true, primary: true }, { label: "Abbrechen", value: false }], cancel: false });
    if (!ok) return;
    setPage("Seite einrichten", (x) => {
      const nm = { ...m };
      for (const i of inputs) { const v = Number(i.input.value.replace(",", ".")); if (v >= 0) nm[i.k] = v / 2.54; }
      x.margins = nm;
      const s = Number(scale.value);
      if (s >= 10 && s <= 400) { x.scale = s; x.fitWidth = x.fitHeight = undefined; }
      x.gridLines = grid.checked || undefined;
      x.headings = heads.checked || undefined;
      x.centerH = centerH.checked || undefined;
      x.centerV = centerV.checked || undefined;
      x.titleRows = readSpan(titleRows.value, true);
      x.titleCols = readSpan(titleCols.value, false);
      const code = (inputs: HTMLInputElement[]) => {
        const [l, c, r] = inputs.map((i) => toCodes(i.value));
        return (l ? `&L${l}` : "") + (c ? `&C${c}` : "") + (r ? `&R${r}` : "") || undefined;
      };
      x.header = code(head);
      x.footer = code(foot);
    });
  }

  return {
    /** The printed pages of a sheet (Umbruchvorschau): the area and each page's rows and columns. */
    layout: (sheet: Sheet) => {
      const pages = paginate(sheet);
      const area = pages.length ? { r0: Math.min(...pages.map((p) => p.rows[0])), c0: Math.min(...pages.map((p) => p.cols[0])), r1: Math.max(...pages.map((p) => p.rows[1])), c1: Math.max(...pages.map((p) => p.cols[1])) } : null;
      return { area, pages };
    },
    exportPdf: (to?: string) => exportPdf(to),
    page: printPage,
    orientation: () => page().orientation ?? "portrait",
    setOrientation,
    setPrintArea: (r: Rect | null) => setPage("Druckbereich", (p) => { p.printArea = r ?? undefined; }),
    /** Seitenumbruch einfügen / entfernen at the active cell (above its row, left of its column), or all reset. */
    breaks: (what: "insert" | "remove" | "reset") => setPage("Seitenumbruch", (p) => {
      const { row, col } = app.sheet.view.active;
      if (what === "reset") { p.rowBreaks = p.colBreaks = undefined; return; }
      const edit = (list: number[] | undefined, at: number) => {
        const set = new Set(list ?? []);
        if (at > 0) { if (what === "insert") set.add(at); else set.delete(at); }
        return set.size ? [...set].sort((a, b) => a - b) : undefined;
      };
      p.rowBreaks = edit(p.rowBreaks, row);
      p.colBreaks = edit(p.colBreaks, col);
    }),
    pageSetupDialog,
    printOption: (k: "gridLines" | "headings") => !!page()[k],
    togglePrintOption: (k: "gridLines" | "headings") => setPage("Blattoptionen", (p) => { p[k] = !p[k] || undefined; }),
    marginsMenu: (): MenuEntry[] => [
      ...Object.entries(MARGINS).map(([name, mm]): MenuEntry => ({
        label: `${name}  (Oben: ${cm(mm.top)}  Unten: ${cm(mm.bottom)}  Links: ${cm(mm.left)}  Rechts: ${cm(mm.right)})`,
        run: () => setPage("Seitenränder", (p) => { p.margins = { ...mm }; }),
      })),
      { label: "Benutzerdefinierte Seitenränder...", run: () => void pageSetupDialog() },
    ],
    sizeMenu: (): MenuEntry[] => PAPERS.map((x) => ({ label: `${x.name}  ${(x.w * 2.54).toFixed(1).replace(".", ",")} cm x ${(x.h * 2.54).toFixed(1).replace(".", ",")} cm`, run: () => setPage("Format", (p) => { p.paperSize = x.code; }) })),
    setFit: (k: "width" | "height", n: number) => setPage("An Format anpassen", (p) => { if (k === "width") p.fitWidth = n || undefined; else p.fitHeight = n || undefined; }),
    scale: () => page().scale ?? 100,
    setScale: (n: number) => setPage("Skalierung", (p) => { p.scale = Math.max(10, Math.min(400, n)); p.fitWidth = p.fitHeight = undefined; }),
  };
}
