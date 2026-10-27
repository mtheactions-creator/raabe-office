// Excel's automatic value-axis scaling, as recorded in Excel 365 (probe_excel_charts.py,
// results/2026-10-27-excel-charts.json):
// - an axis starts at 0 when all values are positive and the smallest is below 5/6 of the
//   largest (mirrored for negative values); otherwise it starts half the data range below the
//   smallest value;
// - the far end is the largest value plus 5 % of (largest - axis start); with values of both
//   signs both ends get 5 % of the data range;
// - the major unit is the smallest of 1, 2, 5 x 10^n that divides that span into at most L
//   intervals, where L depends on the axis length (10 for a default chart);
// - the ends are rounded outwards to whole major units.

export interface AxisScale { min: number; max: number; major: number }

/** The nice units Excel uses, ascending from well below `span`. */
function units(span: number): number[] {
  const e = Math.floor(Math.log10(span)) - 2;
  const out: number[] = [];
  for (let k = e; k <= e + 4; k++) for (const m of [1, 2, 5]) out.push(Number((m * 10 ** k).toPrecision(12)));
  return out;
}

/**
 * The largest number of intervals Excel allows on a value axis `length` pt long (labels of 9 pt).
 * Vertical axes: one more interval per 10 pt above 19.7 pt; horizontal axes as recorded at
 * several widths. Never more than 10.
 */
export function maxIntervals(length: number, horizontal: boolean): number {
  if (!horizontal) return Math.max(1, Math.min(10, (length - 19.7) / 10));
  const knots: [number, number][] = [[0, 1], [107, 3], [147, 4.05], [187, 5.1], [227, 5.9], [267, 8.1], [307, 9.2], [330, 9.35], [340, 10]];
  if (length >= 340) return 10;
  for (let i = 1; i < knots.length; i++) {
    const [x1, y1] = knots[i], [x0, y0] = knots[i - 1];
    if (length <= x1) return y0 + ((length - x0) / (x1 - x0)) * (y1 - y0);
  }
  return 10;
}

/** Excel's automatic scale for values between `lo` and `hi` with at most `limit` intervals. */
export function autoScale(lo: number, hi: number, limit = 10): AxisScale {
  if (!isFinite(lo) || !isFinite(hi)) { lo = 0; hi = 1; }
  if (lo > hi) [lo, hi] = [hi, lo];
  let a: number, b: number;
  if (lo === hi) {
    // One value: from 0 to the value (plus 5 %); 0 alone gives 0 to 1.
    if (lo === 0) { a = 0; b = 1; }
    else if (lo > 0) { a = 0; b = lo * 1.05; }
    else { b = 0; a = lo * 1.05; }
  } else if (lo >= 0) {
    a = lo < (hi * 5) / 6 ? 0 : lo - (hi - lo) / 2;
    b = hi + 0.05 * (hi - a);
  } else if (hi <= 0) {
    b = hi > (lo * 5) / 6 ? 0 : hi + (hi - lo) / 2;
    a = lo - 0.05 * (b - lo);
  } else {
    a = lo - 0.05 * (hi - lo);
    b = hi + 0.05 * (hi - lo);
  }
  const span = b - a;
  let major = 1;
  for (const u of units(span)) {
    major = u;
    if (span / u <= limit + 1e-9) break;
  }
  const min = round(Math.floor(a / major + 1e-9) * major, major);
  const max = round(Math.ceil(b / major - 1e-9) * major, major);
  return { min, max: max === min ? round(min + major, major) : max, major };
}

/** A multiple of a unit without floating-point noise. */
function round(v: number, unit: number): number {
  const d = Math.max(0, -Math.floor(Math.log10(unit)) + 1);
  return Number(v.toFixed(Math.min(20, d)));
}

/** The tick values of a scale. */
export function ticks(s: AxisScale): number[] {
  const out: number[] = [];
  const n = Math.round((s.max - s.min) / s.major);
  for (let i = 0; i <= n && i < 1000; i++) out.push(round(s.min + i * s.major, s.major));
  return out;
}
