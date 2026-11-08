// Drawing sparklines as Excel 365 draws them (D-085, measured on Excel's CopyPicture of cells at 150 %):
// - the plot area is the cell less 1.25 pt left and right and 2 pt above and below; the n values sit in n
//   equal slots, line points at the slots' centres, columns 72 % of a slot wide and centred in it;
// - columns rise from zero when the values cross it, else from the bottom (all positive: the smallest
//   value is a bar of one pixel) or the top (all negative); Gewinn/Verlust fills the upper half for a
//   positive value, the lower half for a negative one, nothing for zero;
// - a point's colour: Höchstpunkt, Tiefpunkt, Erster Punkt, Letzter Punkt, Negative Punkte, then
//   Datenpunkte (markers) and the series colour; on a line each such point is a diamond of
//   2 pt + 4/3 of the line's Stärke across;
// - the horizontal axis (Achse anzeigen) is drawn at zero only when zero lies inside the vertical range;
// - empty cells: Lücken break the line, Null counts them as 0, Datenpunkte mit einer Linie verbinden
//   skips them; values in hidden rows and columns are left out unless displayHidden is set.

import { parseRange, type Sheet, type SparkColor, type SparkColorKey, type SparklineGroup, type Workbook } from "@raabe/sheet-engine";
import { resolveXlsxColor } from "@raabe/ooxml";

const THEME_KEYS = ["lt1", "dk1", "lt2", "dk2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6", "hlink", "folHlink"];

/** One sparkline's values (null: empty) and, with a date axis, the date of each. */
export interface SparkSeries { values: (number | null)[]; dates?: (number | null)[] }

/** A range reference "Tabelle1!A1:E1" ("'Mein Blatt'!A1:E1") as sheet name and rectangle. */
export function splitRef(f: string, home: string): { sheet: string; rect: ReturnType<typeof parseRange> } {
  const m = /^(?:'((?:[^']|'')+)'|([^!]+))!(.+)$/.exec(f.trim());
  return { sheet: m ? (m[1] ?? m[2]).replace(/''/g, "'") : home, rect: parseRange((m ? m[3] : f).replace(/\$/g, "")) };
}

function cellsOf(book: Workbook | undefined, home: Sheet, f: string, hidden: boolean): (number | null)[] | null {
  const ref = splitRef(f, home.name);
  const sheet = ref.sheet.toLowerCase() === home.name.toLowerCase() ? home : book?.sheetByName(ref.sheet);
  if (!sheet || !ref.rect) return null;
  const out: (number | null)[] = [];
  const r = ref.rect;
  for (let rr = r.r0; rr <= r.r1; rr++) {
    if (!hidden && sheet.rows.get(rr)?.hidden) continue;
    for (let cc = r.c0; cc <= r.c1; cc++) {
      if (!hidden && sheet.cols.get(cc)?.hidden) continue;
      const v = sheet.get(rr, cc)?.v;
      out.push(typeof v === "number" ? v : null);
    }
  }
  return out;
}

/** The values a sparkline shows (empty cells as Excel's setting says: null for a gap or a span, 0 for Null). */
export function sparkSeries(book: Workbook | undefined, home: Sheet, g: SparklineGroup, f: string): SparkSeries | null {
  const raw = cellsOf(book, home, f, !!g.displayHidden);
  if (!raw) return null;
  const values = g.displayEmptyCellsAs === "zero" ? raw.map((v) => v ?? 0) : raw;
  const dates = g.dateRange ? cellsOf(book, home, g.dateRange, !!g.displayHidden) ?? undefined : undefined;
  return { values, dates };
}

/** The vertical range of each sparkline of a group (automatic, the same for all, or a custom value). */
export function sparkRanges(g: SparklineGroup, series: (SparkSeries | null)[]): ({ lo: number; hi: number } | null)[] {
  const nums = (s: SparkSeries | null) => (s?.values ?? []).filter((v): v is number => v !== null);
  const all = series.flatMap(nums);
  return series.map((s) => {
    const own = nums(s);
    if (!own.length) return null;
    const pick = (type: SparklineGroup["minAxisType"], manual: number | undefined, f: (xs: number[]) => number) =>
      type === "custom" && manual !== undefined ? manual : type === "group" ? f(all) : f(own);
    return { lo: pick(g.minAxisType, g.manualMin, (xs) => Math.min(...xs)), hi: pick(g.maxAxisType, g.manualMax, (xs) => Math.max(...xs)) };
  });
}

/** "#RRGGBB" of a sparkline colour in the workbook's design (a design colour follows Seitenlayout > Designs). */
export function sparkRgb(c: SparkColor, theme?: Record<string, string>): string {
  if (!c.theme || !theme || !Object.keys(theme).length) return c.rgb;
  const t = Object.fromEntries(Object.entries(theme).map(([k, v]) => [k, v.replace("#", "")]));
  if (!t[THEME_KEYS[c.theme.t]]) return c.rgb;
  return resolveXlsxColor({ theme: String(c.theme.t), ...(c.theme.tint ? { tint: String(c.theme.tint) } : {}) }, { theme: t, indexed: [] }) ?? c.rgb;
}

/** Which colour a point takes (null: the series colour, no marker). */
export function pointColor(g: SparklineGroup, values: (number | null)[], i: number): SparkColorKey | null {
  const v = values[i];
  if (v === null) return null;
  const nums = values.filter((x): x is number => x !== null);
  const firstI = values.findIndex((x) => x !== null);
  let lastI = values.length - 1;
  while (lastI >= 0 && values[lastI] === null) lastI--;
  if (g.high && v === Math.max(...nums)) return "high";
  if (g.low && v === Math.min(...nums)) return "low";
  if (g.first && i === firstI) return "first";
  if (g.last && i === lastI) return "last";
  if (g.negative && v < 0) return "negative";
  if (g.markers && g.type === "line") return "markers";
  return null;
}

export interface SparkBox { x: number; y: number; w: number; h: number }

/**
 * Draws one sparkline into a cell's box (CSS px). `pt` is the CSS px of one point at the current zoom,
 * `dpr` the device pixel ratio, `theme` the workbook's design colours.
 */
export function drawSparkline(ctx: CanvasRenderingContext2D, g: SparklineGroup, s: SparkSeries, range: { lo: number; hi: number }, box: SparkBox,
  pt: number, dpr: number, theme?: Record<string, string>): void {
  const n = s.values.length;
  if (!n) return;
  const color = (k: SparkColorKey) => sparkRgb(g.colors[k], theme);
  const left = box.x + 1.25 * pt, top = box.y + 2 * pt, w = box.w - 2.5 * pt, h = box.h - 4 * pt;
  if (w <= 0 || h <= 0) return;
  const slot = w / n;
  // Horizontal positions: slot centres, or by date between the first and the last slot's centre.
  const xs: number[] = [];
  const dates = s.dates && s.dates.length === n && s.dates.every((d) => d !== null) ? (s.dates as number[]) : null;
  const d0 = dates ? Math.min(...dates) : 0, d1 = dates ? Math.max(...dates) : 0;
  for (let i = 0; i < n; i++) {
    const t = dates && d1 > d0 ? (dates[i] - d0) / (d1 - d0) : n > 1 ? i / (n - 1) : 0.5;
    const x = left + slot / 2 + t * (w - slot);
    xs.push(g.rightToLeft ? 2 * left + w - x : x);
  }
  const { lo, hi } = range;
  const y = (v: number) => (hi === lo ? top + h / 2 : top + h - ((v - lo) / (hi - lo)) * h);
  const onePx = 1 / dpr;
  ctx.save();
  ctx.beginPath();
  ctx.rect(box.x, box.y, box.w, box.h);
  ctx.clip();
  // Achse anzeigen: at zero, when zero lies within the vertical range (Gewinn/Verlust: always the middle).
  const axisAt = g.type === "stacked" ? top + h / 2 : lo <= 0 && hi >= 0 && lo < hi ? y(0) : null;
  const drawAxis = () => {
    if (!g.displayXAxis || axisAt === null) return;
    ctx.fillStyle = color("axis");
    const t = Math.max(onePx, Math.round(pt * dpr) / dpr);
    ctx.fillRect(box.x, Math.round((axisAt - t / 2) * dpr) / dpr, box.w, t);
  };
  if (g.type === "line") {
    drawAxis();
    const lw = (g.lineWeight ?? 0.75) * pt;
    ctx.strokeStyle = color("series");
    ctx.lineWidth = Math.max(onePx, lw);
    ctx.lineJoin = "round";
    ctx.lineCap = "round";
    ctx.beginPath();
    let started = false;
    for (let i = 0; i < n; i++) {
      const v = s.values[i];
      if (v === null) { if (g.displayEmptyCellsAs !== "span") started = false; continue; }
      if (started) ctx.lineTo(xs[i], y(v)); else { ctx.moveTo(xs[i], y(v)); started = true; }
    }
    ctx.stroke();
    const r = (2 + (4 / 3) * (g.lineWeight ?? 0.75)) * pt / 2;
    for (let i = 0; i < n; i++) {
      const k = pointColor(g, s.values, i);
      if (!k) continue;
      const cx = xs[i], cy = y(s.values[i]!);
      ctx.fillStyle = color(k);
      ctx.beginPath();
      ctx.moveTo(cx, cy - r);
      ctx.lineTo(cx + r, cy);
      ctx.lineTo(cx, cy + r);
      ctx.lineTo(cx - r, cy);
      ctx.closePath();
      ctx.fill();
    }
  } else {
    // Columns 72 % of a slot; with a date axis as wide as the closest two dates allow.
    let bw = slot * 0.72;
    if (dates && d1 > d0) {
      const sorted = [...dates].sort((a, b) => a - b);
      let gap = Infinity;
      for (let i = 1; i < sorted.length; i++) if (sorted[i] > sorted[i - 1]) gap = Math.min(gap, sorted[i] - sorted[i - 1]);
      if (Number.isFinite(gap)) bw = Math.min(bw, (gap / (d1 - d0)) * (w - slot) * 0.72);
    }
    const base = lo <= 0 && hi >= 0 ? y(0) : lo > 0 ? top + h : top;
    for (let i = 0; i < n; i++) {
      const v = s.values[i];
      if (v === null) continue;
      const k = pointColor(g, s.values, i);
      ctx.fillStyle = color(k && k !== "markers" ? k : "series");
      const x0 = Math.round((xs[i] - bw / 2) * dpr) / dpr, x1 = Math.round((xs[i] + bw / 2) * dpr) / dpr;
      if (g.type === "stacked") {
        if (v === 0) continue;
        const mid = Math.round((top + h / 2) * dpr) / dpr;
        if (v > 0) ctx.fillRect(x0, Math.round(top * dpr) / dpr, x1 - x0, mid - Math.round(top * dpr) / dpr);
        else ctx.fillRect(x0, mid, x1 - x0, Math.round((top + h) * dpr) / dpr + onePx - mid);
        continue;
      }
      const yv = Math.max(top, Math.min(top + h, y(v)));
      let a = Math.round(Math.min(base, yv) * dpr) / dpr, b = Math.round(Math.max(base, yv) * dpr) / dpr;
      // Excel's bars take the pixel row of their lower end too (measured: bottoms at the plot's last row).
      b += onePx;
      ctx.fillRect(x0, a, x1 - x0, b - a);
    }
    drawAxis();
  }
  ctx.restore();
}

/** Draws every sparkline of a sheet whose cell passes `visible`, into the boxes `boxOf` gives. */
export function drawSparklines(ctx: CanvasRenderingContext2D, book: Workbook | undefined, sheet: Sheet, zoom: number, dpr: number,
  visible: (row: number, col: number) => boolean, boxOf: (row: number, col: number) => SparkBox): void {
  const pt = (96 / 72) * zoom;
  for (const g of sheet.sparklines ?? []) {
    const shown = g.items.filter((it) => visible(it.row, it.col));
    if (!shown.length) continue;
    // A range "Identisch für alle Sparklines" needs every sparkline of the group.
    const needAll = g.minAxisType === "group" || g.maxAxisType === "group";
    const items = needAll ? g.items : shown;
    const series = items.map((it) => sparkSeries(book, sheet, g, it.f));
    const ranges = sparkRanges(g, series);
    items.forEach((it, i) => {
      const s = series[i], r = ranges[i];
      if (!s || !r || !visible(it.row, it.col)) return;
      drawSparkline(ctx, g, s, r, boxOf(it.row, it.col), pt, dpr, book?.theme);
    });
  }
}
