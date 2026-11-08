// Ansicht > Seitenlayout (D-093): the sheet shown as pages, as Excel's Seitenlayout view (sheetView
// view="pageLayout", its own zoom zoomScalePageLayoutView, Lineal showRuler; recorded with
// probe_excel_windows.py). The cells stay editable as in Normal: the rows and columns are laid out with the
// pages' margins and the space between pages added (an axis with gaps, so pointing, selecting and scrolling
// work as before); outside the pages it is grey. The header and footer areas of each page show their text,
// and a click there opens Seite einrichten at Kopfzeile/Fußzeile (Excel edits in place; its prompt texts are
// its own prose, so ours differ). Lineal: centimetre rulers along the top and left edge of the sheet area.

import type { PageSetup, Sheet } from "@raabe/sheet-engine";
import { h } from "@raabe/shell";
import { Axis } from "./grid/metrics";
import type { CalcApp } from "./app";
import { cutPages, headerSections } from "./print";

/** Paper sizes in inches (Seite einrichten's list). */
const PAPERS: Record<number, { w: number; h: number }> = { 9: { w: 8.27, h: 11.69 }, 8: { w: 11.69, h: 16.54 }, 11: { w: 5.83, h: 8.27 }, 1: { w: 8.5, h: 11 }, 5: { w: 8.5, h: 14 } };
const NORMAL = { left: 0.7, right: 0.7, top: 0.75, bottom: 0.75, header: 0.3, footer: 0.3 };
/** Grey between pages and above the first one, in px at 100 % (our own). */
const GAP = 14;
const TOP = 26;
const RULER = 16;

/** An axis whose rows (columns) are spread over pages: margins and the space between pages added. */
export class GapAxis extends Axis {
  /** Page starts (first index of each page) and the room added before each. */
  constructor(private base: Axis, readonly starts: number[], gaps: number[]) {
    super(base.def, base.count);
    let sum = 0;
    this.cum = gaps.map((g) => (sum += g));
  }
  private cum: number[];
  size(i: number): number { return this.base.size(i); }
  /** The room added before index i (all page starts up to it). */
  gapAt(i: number): number {
    let lo = 0, hi = this.starts.length - 1, k = -1;
    while (lo <= hi) { const m = (lo + hi) >> 1; if (this.starts[m] <= i) { k = m; lo = m + 1; } else hi = m - 1; }
    return k < 0 ? 0 : this.cum[k];
  }
  pos(i: number): number { return this.base.pos(i) + this.gapAt(i); }
}

export interface Pages {
  rows: GapAxis; cols: GapAxis;
  /** Each page row's / column's paper start (in the gapped axis, px at 100 %) and its size. */
  paperY: number[]; paperX: number[]; paperH: number; paperW: number;
  margins: typeof NORMAL;
}

/** The pages of a sheet in the Seitenlayout view: page breaks as printing makes them (manual breaks too). */
export function layoutPages(sheet: Sheet, base: { rows: Axis; cols: Axis }): Pages {
  const p: PageSetup = sheet.page ?? {};
  const sz0 = PAPERS[p.paperSize ?? 9] ?? PAPERS[9];
  const sz = p.orientation === "landscape" ? { w: sz0.h, h: sz0.w } : sz0;
  const m = { ...NORMAL, ...(p.margins ?? {}) };
  const ph = sz.h * 96, pw = sz.w * 96;
  const roomH = (sz.h - m.top - m.bottom) * 96, roomW = (sz.w - m.left - m.right) * 96;
  const used = sheet.usedRange();
  const lastR = Math.min(base.rows.count - 1, Math.max((used?.r1 ?? 0) + 400, 2000));
  const lastC = Math.min(base.cols.count - 1, Math.max((used?.c1 ?? 0) + 40, 80));
  const pr = cutPages(0, lastR, (i) => base.rows.size(i), roomH, p.rowBreaks);
  const pc = cutPages(0, lastC, (i) => base.cols.size(i), roomW, p.colBreaks);
  const gapsOf = (pages: [number, number][], axis: Axis, room: number, before: number, after: number, first: number) =>
    pages.map((_pg, k) => {
      if (k === 0) return first + before;
      const prev = pages[k - 1];
      const usedPx = axis.pos(prev[1] + 1) - axis.pos(prev[0]);
      return Math.max(0, room - usedPx) + after + GAP + before;
    });
  const rows = new GapAxis(base.rows, pr.map((x) => x[0]), gapsOf(pr, base.rows, roomH, m.top * 96, m.bottom * 96, TOP));
  const cols = new GapAxis(base.cols, pc.map((x) => x[0]), gapsOf(pc, base.cols, roomW, m.left * 96, m.right * 96, GAP));
  return {
    rows, cols, paperH: ph, paperW: pw, margins: m,
    paperY: pr.map((x) => rows.pos(x[0]) - m.top * 96),
    paperX: pc.map((x) => cols.pos(x[0]) - m.left * 96),
  };
}

const keyOf = (s: Sheet) => JSON.stringify([s.page ?? null, [...s.rows].filter(([, r]) => r.height !== undefined || r.hidden), [...s.cols], s.defaultRowHeight, s.defaultColWidth, s.usedRange()]);

export function pageLayoutFeature(app: CalcApp) {
  const overlay = h("canvas", { class: "pl-overlay", style: "position:absolute;left:0;top:0;pointer-events:none;z-index:2" }) as HTMLCanvasElement;
  let cache: { sheet: Sheet; key: string; pages: Pages } | null = null;
  let hover: { page: [number, number]; foot: boolean } | null = null;

  const on = () => !!app.sheet.view.layout;

  function pagesOf(sheet: Sheet, base: { rows: Axis; cols: Axis }): Pages {
    const key = keyOf(sheet);
    if (!cache || cache.sheet !== sheet || cache.key !== key) cache = { sheet, key, pages: layoutPages(sheet, base) };
    return cache.pages;
  }

  /** The grid's axes: with gaps while the Seitenlayout view is on. */
  function axes(sheet: Sheet, base: { rows: Axis; cols: Axis }): { rows: Axis; cols: Axis } {
    if (!sheet.view.layout) return base;
    const p = pagesOf(sheet, base);
    return { rows: p.rows, cols: p.cols };
  }

  /** Where position 0 of the gapped axes is on the canvas (the grid shows the first margin at the top, D-093). */
  function origin(g: NonNullable<ReturnType<typeof app.grid.geometry>>, v: { top: number; left: number }): { ox: number; oy: number } {
    const z = g.zoom;
    return { oy: g.headerH - (v.top ? g.rows.pos(v.top) : 0) * z, ox: g.headerW - (v.left ? g.cols.pos(v.left) : 0) * z };
  }

  /** Header/footer zones of the visible pages on the canvas (CSS px). */
  function zones(): { x: number; y: number; w: number; h: number; page: [number, number]; foot: boolean }[] {
    const g = app.grid.geometry();
    if (!g || !cache || cache.sheet !== app.sheet) return [];
    const p = cache.pages, v = app.sheet.view, z = g.zoom;
    const { ox, oy } = origin(g, v);
    const out = [];
    for (let r = 0; r < p.paperY.length; r++) {
      const y = oy + p.paperY[r] * z;
      if (y > g.height || y + p.paperH * z < 0) continue;
      for (let c = 0; c < p.paperX.length; c++) {
        const x = ox + p.paperX[c] * z;
        if (x > g.width || x + p.paperW * z < 0) continue;
        const left = x + p.margins.left * 96 * z, w = (p.paperW - (p.margins.left + p.margins.right) * 96) * z;
        out.push({ x: left, y: y + p.margins.header * 96 * z, w, h: Math.max(8, (p.margins.top - p.margins.header) * 96 * z), page: [r, c] as [number, number], foot: false });
        out.push({ x: left, y: y + (p.paperH - p.margins.bottom * 96) * z, w, h: Math.max(8, (p.margins.bottom - p.margins.footer) * 96 * z), page: [r, c] as [number, number], foot: true });
      }
    }
    return out;
  }

  /** After every paint: grey around the pages, the header and footer text, the rulers. */
  function paint(): void {
    const g = app.grid.geometry();
    if (!on() || !g) { if (overlay.isConnected) overlay.remove(); return; }
    if (!overlay.isConnected) app.grid.el.insertBefore(overlay, app.grid.el.querySelector(".grid-canvas")?.nextSibling ?? null);
    const dpr = devicePixelRatio || 1;
    if (overlay.width !== Math.round(g.width * dpr) || overlay.height !== Math.round(g.height * dpr)) {
      overlay.width = Math.round(g.width * dpr);
      overlay.height = Math.round(g.height * dpr);
      overlay.style.width = `${g.width}px`;
      overlay.style.height = `${g.height}px`;
    }
    const ctx = overlay.getContext("2d")!;
    ctx.setTransform(dpr, 0, 0, dpr, 0, 0);
    ctx.clearRect(0, 0, g.width, g.height);
    pagesOf(app.sheet, { rows: g.rows, cols: g.cols });
    const p = cache!.pages, v = app.sheet.view, z = g.zoom;
    const { ox, oy } = origin(g, v);
    // Grey outside the pages (the cells area only).
    ctx.save();
    ctx.beginPath();
    ctx.rect(g.headerW, g.headerH, g.width - g.headerW, g.height - g.headerH);
    ctx.clip();
    ctx.beginPath();
    ctx.rect(g.headerW, g.headerH, g.width - g.headerW, g.height - g.headerH);
    const papers: { x: number; y: number; r: number; c: number }[] = [];
    for (let r = 0; r < p.paperY.length; r++) {
      const y = oy + p.paperY[r] * z;
      if (y > g.height || y + p.paperH * z < g.headerH) continue;
      for (let c = 0; c < p.paperX.length; c++) {
        const x = ox + p.paperX[c] * z;
        if (x > g.width || x + p.paperW * z < g.headerW) continue;
        papers.push({ x, y, r, c });
        ctx.rect(x, y, p.paperW * z, p.paperH * z);
      }
    }
    ctx.fillStyle = "#e6e6e6";
    ctx.fill("evenodd");
    // The margins are paper, not cells: white over whatever the grid drew there (gridlines).
    ctx.fillStyle = "#ffffff";
    for (const q of papers) {
      const W = p.paperW * z, H = p.paperH * z, mt = p.margins.top * 96 * z, mb = p.margins.bottom * 96 * z, ml = p.margins.left * 96 * z, mr = p.margins.right * 96 * z;
      ctx.fillRect(q.x, q.y, W, mt);
      ctx.fillRect(q.x, q.y + H - mb, W, mb);
      ctx.fillRect(q.x, q.y, ml, H);
      ctx.fillRect(q.x + W - mr, q.y, mr, H);
      ctx.strokeStyle = "#c6c6c6";
      ctx.lineWidth = 1;
      ctx.strokeRect(q.x + 0.5, q.y + 0.5, W - 1, H - 1);
    }
    // Header and footer text (Seite einrichten's), and the zone under the pointer.
    const page = app.sheet.page;
    ctx.font = `${Math.max(7, 11 * 96 / 72 * z * 0.9)}px Calibri, Carlito, sans-serif`;
    ctx.textBaseline = "middle";
    const all = p.paperY.length * p.paperX.length;
    for (const zn of zones()) {
      const code = zn.foot ? page?.footer : page?.header;
      const n = zn.page[1] * p.paperY.length + zn.page[0] + 1;
      const sec = code ? headerSections(code, { page: n, pages: all, sheet: app.sheet.name, file: app.wb.docName ?? "" }) : null;
      const yy = zn.y + zn.h / 2;
      ctx.fillStyle = "#333";
      if (sec) {
        ctx.textAlign = "left"; ctx.fillText(sec.L, zn.x, yy);
        ctx.textAlign = "center"; ctx.fillText(sec.C, zn.x + zn.w / 2, yy);
        ctx.textAlign = "right"; ctx.fillText(sec.R, zn.x + zn.w, yy);
      }
      if (hover && hover.page[0] === zn.page[0] && hover.page[1] === zn.page[1] && hover.foot === zn.foot) {
        ctx.setLineDash([3, 2]);
        ctx.strokeStyle = "#8a8a8a";
        ctx.strokeRect(zn.x + 0.5, zn.y + 0.5, zn.w - 1, zn.h - 1);
        ctx.setLineDash([]);
        if (!sec || !(sec.L + sec.C + sec.R).trim()) {
          ctx.fillStyle = "#7a7a7a";
          ctx.textAlign = "center";
          ctx.fillText(zn.foot ? "Fußzeile bearbeiten" : "Kopfzeile bearbeiten", zn.x + zn.w / 2, yy);
        }
      }
    }
    ctx.restore();
    if (!v.noRuler) rulers(ctx, g, papers, p, z);
  }

  /** Lineal: centimetres from each page's printable edge, the margins shaded. */
  function rulers(ctx: CanvasRenderingContext2D, g: NonNullable<ReturnType<typeof app.grid.geometry>>, papers: { x: number; y: number; r: number; c: number }[], p: Pages, z: number): void {
    ctx.save();
    ctx.fillStyle = "#f3f3f3";
    ctx.fillRect(g.headerW, g.headerH, g.width - g.headerW, RULER);
    ctx.fillRect(g.headerW, g.headerH, RULER, g.height - g.headerH);
    ctx.font = "9px Segoe UI, sans-serif";
    ctx.fillStyle = "#666";
    ctx.strokeStyle = "#999";
    ctx.lineWidth = 1;
    const cmPx = 96 / 2.54 * z;
    const seenX = new Set<number>(), seenY = new Set<number>();
    for (const q of papers) {
      if (!seenX.has(q.c)) {
        seenX.add(q.c);
        const x0 = q.x + p.margins.left * 96 * z, x1 = q.x + (p.paperW - p.margins.right * 96) * z;
        ctx.fillStyle = "#ffffff";
        ctx.fillRect(Math.max(g.headerW + RULER, x0), g.headerH + 2, Math.max(0, x1 - Math.max(g.headerW + RULER, x0)), RULER - 4);
        ctx.fillStyle = "#666";
        ctx.textAlign = "center";
        for (let k = 0; x0 + k * cmPx <= x1; k++) {
          const x = x0 + k * cmPx;
          if (x < g.headerW + RULER) continue;
          if (k) ctx.fillText(String(k), x, g.headerH + RULER / 2 + 1);
          ctx.beginPath(); ctx.moveTo(Math.round(x) + 0.5, g.headerH + RULER - 4); ctx.lineTo(Math.round(x) + 0.5, g.headerH + RULER); ctx.stroke();
        }
      }
      if (!seenY.has(q.r)) {
        seenY.add(q.r);
        const y0 = q.y + p.margins.top * 96 * z, y1 = q.y + (p.paperH - p.margins.bottom * 96) * z;
        ctx.fillStyle = "#ffffff";
        ctx.fillRect(g.headerW + 2, Math.max(g.headerH + RULER, y0), RULER - 4, Math.max(0, y1 - Math.max(g.headerH + RULER, y0)));
        ctx.fillStyle = "#666";
        ctx.textAlign = "center";
        for (let k = 0; y0 + k * cmPx <= y1; k++) {
          const y = y0 + k * cmPx;
          if (y < g.headerH + RULER) continue;
          if (k) ctx.fillText(String(k), g.headerW + RULER / 2, y + 1);
          ctx.beginPath(); ctx.moveTo(g.headerW + RULER - 4, Math.round(y) + 0.5); ctx.lineTo(g.headerW + RULER, Math.round(y) + 0.5); ctx.stroke();
        }
      }
    }
    ctx.restore();
  }

  function zoneAt(e: MouseEvent) {
    const r = app.grid.el.getBoundingClientRect();
    const x = e.clientX - r.left, y = e.clientY - r.top;
    return zones().find((zn) => x >= zn.x && x <= zn.x + zn.w && y >= zn.y && y <= zn.y + zn.h) ?? null;
  }

  function install(): void {
    app.grid.el.addEventListener("pointerdown", (e) => {
      if (!on() || e.button !== 0) return;
      const zn = zoneAt(e);
      if (!zn) return;
      e.preventDefault();
      e.stopPropagation();
      void app.print.pageSetupDialog("header");
    }, true);
    app.grid.el.addEventListener("pointermove", (e) => {
      if (!on()) return;
      const zn = zoneAt(e);
      const next = zn ? { page: zn.page, foot: zn.foot } : null;
      if (JSON.stringify(next) !== JSON.stringify(hover)) { hover = next; paint(); }
    });
  }

  /** Normal / Seitenlayout: each keeps its own zoom (recorded: zoomScaleNormal, zoomScalePageLayoutView). */
  function set(layout: boolean): void {
    const v = app.sheet.view;
    if (!!v.layout === layout) return;
    if (layout) {
      const z = v.zoom;
      v.layout = true;
      v.zoom = v.normalZoom && v.normalZoom !== z ? v.zoom : z;
      v.normalZoom = z;
      // Fenster fixieren and Teilen have no effect in the Seitenlayout view (Excel greys them).
    } else {
      delete v.layout;
      v.zoom = v.normalZoom ?? 100;
      delete v.normalZoom;
    }
    cache = null;
  }

  return { axes, paint, install, set, on, get pages() { return cache?.pages ?? null; } };
}
