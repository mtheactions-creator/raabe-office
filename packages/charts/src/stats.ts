// Trendline and histogram maths as Excel computes them.

export interface LinearFit { slope: number; intercept: number; /** null when y does not vary (Excel: #N/A). */ r2: number | null }

/** Least-squares line through the points with numeric x and y. */
export function linearFit(xs: number[], ys: number[]): LinearFit | null {
  const pts: [number, number][] = [];
  for (let i = 0; i < Math.min(xs.length, ys.length); i++) {
    if (isFinite(xs[i]) && isFinite(ys[i])) pts.push([xs[i], ys[i]]);
  }
  const n = pts.length;
  if (n < 2) return null;
  const mx = pts.reduce((s, p) => s + p[0], 0) / n, my = pts.reduce((s, p) => s + p[1], 0) / n;
  let sxx = 0, sxy = 0, syy = 0;
  for (const [x, y] of pts) {
    sxx += (x - mx) ** 2;
    sxy += (x - mx) * (y - my);
    syy += (y - my) ** 2;
  }
  if (sxx === 0) return null;
  const slope = sxy / sxx, intercept = my - slope * mx;
  const r2 = syy === 0 ? null : (sxy * sxy) / (sxx * syy);
  return { slope, intercept, r2 };
}

export interface Bin { from: number; to: number; count: number }

/** Rounds to two significant digits (Excel's automatic bin width). */
function twoDigits(v: number): number {
  return Number(v.toPrecision(2));
}

/**
 * Excel's automatic histogram bins: width by Scott's rule, 3.5 s / n^(1/3) with the sample
 * standard deviation, rounded to two significant digits; bins from the smallest value, the
 * first closed on both ends ([a, b]), the others open on the left ((a, b]).
 */
export function histogramBins(values: number[]): Bin[] {
  const v = values.filter((x) => isFinite(x));
  if (!v.length) return [];
  const n = v.length;
  const lo = Math.min(...v), hi = Math.max(...v);
  const mean = v.reduce((s, x) => s + x, 0) / n;
  const sd = n > 1 ? Math.sqrt(v.reduce((s, x) => s + (x - mean) ** 2, 0) / (n - 1)) : 0;
  let width = twoDigits((3.5 * sd) / Math.cbrt(n));
  // All values equal: one bin as wide as the value (recorded: 5 → [5, 10]).
  if (!(width > 0)) width = Math.abs(lo) || 1;
  const count = Math.max(1, Math.ceil((hi - lo) / width - 1e-9));
  const bins: Bin[] = [];
  for (let i = 0; i < count; i++) {
    bins.push({ from: Number((lo + i * width).toPrecision(12)), to: Number((lo + (i + 1) * width).toPrecision(12)), count: 0 });
  }
  for (const x of v) {
    let i = Math.ceil((x - lo) / width - 1e-9) - 1;
    if (i < 0) i = 0;
    if (i >= count) i = count - 1;
    bins[i].count++;
  }
  return bins;
}
