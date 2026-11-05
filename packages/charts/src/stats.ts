import { equationText, formatCoefficient } from "./format";

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

export interface TrendFit {
  /** The fitted value at x. */
  at: (x: number) => number;
  /** The label's equation as Excel writes it (none for moving averages). */
  equation?: string;
  r2?: number | null;
  /** Moving averages: the averaged points themselves. */
  points?: [number, number][];
}

/** R² of a fit computed on transformed values (Excel uses the transformed y for exp and power). */
function rSquared(ys: number[], fitted: number[]): number | null {
  const n = ys.length;
  const my = ys.reduce((s, y) => s + y, 0) / n;
  const tot = ys.reduce((s, y) => s + (y - my) ** 2, 0);
  if (tot === 0) return null;
  const res = ys.reduce((s, y, i) => s + (y - fitted[i]) ** 2, 0);
  return 1 - res / tot;
}

/** Least squares for a polynomial of `order` (normal equations, Gaussian elimination). */
function polyFit(xs: number[], ys: number[], order: number): number[] | null {
  const k = order + 1;
  if (xs.length < k) return null;
  const a: number[][] = Array.from({ length: k }, () => new Array(k + 1).fill(0));
  for (let i = 0; i < xs.length; i++) {
    for (let r = 0; r < k; r++) {
      for (let c = 0; c < k; c++) a[r][c] += xs[i] ** (r + c);
      a[r][k] += ys[i] * xs[i] ** r;
    }
  }
  for (let c = 0; c < k; c++) {
    let p = c;
    for (let r = c + 1; r < k; r++) if (Math.abs(a[r][c]) > Math.abs(a[p][c])) p = r;
    if (Math.abs(a[p][c]) < 1e-12) return null;
    [a[c], a[p]] = [a[p], a[c]];
    for (let r = 0; r < k; r++) {
      if (r === c) continue;
      const m = a[r][c] / a[c][c];
      for (let j = c; j <= k; j++) a[r][j] -= m * a[c][j];
    }
  }
  return a.map((row, i) => row[k] / row[i]);
}

/** A trendline of any of Excel's types through the points (D-066). */
export function fitTrend(t: { type: string; order?: number; period?: number; intercept?: number }, xs: number[], ys: number[]): TrendFit | null {
  const pts = xs.map((x, i) => [x, ys[i]] as [number, number]).filter(([x, y]) => isFinite(x) && isFinite(y));
  const X = pts.map((p) => p[0]), Y = pts.map((p) => p[1]);
  if (pts.length < 2) return null;
  switch (t.type) {
    case "exp": {
      if (Y.some((y) => y <= 0)) return null;
      const l = linearFit(X, Y.map(Math.log));
      if (!l) return null;
      const a = Math.exp(l.intercept), b = l.slope;
      return { at: (x) => a * Math.exp(b * x), equation: `y = ${formatCoefficient(a)}e${b < 0 ? "-" : ""}${formatCoefficient(Math.abs(b))}x`, r2: l.r2 };
    }
    case "log": {
      if (X.some((x) => x <= 0)) return null;
      const l = linearFit(X.map(Math.log), Y);
      if (!l) return null;
      const sign = l.intercept < 0 ? "-" : "+";
      return { at: (x) => l.slope * Math.log(x) + l.intercept, equation: `y = ${formatCoefficient(l.slope)}ln(x) ${sign} ${formatCoefficient(Math.abs(l.intercept))}`, r2: l.r2 };
    }
    case "power": {
      if (X.some((x) => x <= 0) || Y.some((y) => y <= 0)) return null;
      const l = linearFit(X.map(Math.log), Y.map(Math.log));
      if (!l) return null;
      const a = Math.exp(l.intercept), b = l.slope;
      return { at: (x) => a * x ** b, equation: `y = ${formatCoefficient(a)}x${formatCoefficient(b)}`, r2: l.r2 };
    }
    case "poly": {
      const order = Math.max(2, Math.min(6, t.order ?? 2));
      const c = polyFit(X, Y, order);
      if (!c) return null;
      const at = (x: number) => c.reduce((s, k, i) => s + k * x ** i, 0);
      const SUP = ["", "", "²", "³", "⁴", "⁵", "⁶"];
      let eq = "";
      for (let i = order; i >= 0; i--) {
        const k = c[i];
        const v = formatCoefficient(Math.abs(k));
        if (v === "0") continue;
        const term = i === 0 ? v : `${v === "1" ? "" : v}x${SUP[i]}`;
        eq += eq ? ` ${k < 0 ? "-" : "+"} ${term}` : `${k < 0 ? "-" : ""}${term}`;
      }
      return { at, equation: `y = ${eq || "0"}`, r2: rSquared(Y, X.map(at)) };
    }
    case "movingAvg": {
      const p = Math.max(2, t.period ?? 2);
      const points: [number, number][] = [];
      for (let i = p - 1; i < pts.length; i++) points.push([X[i], Y.slice(i - p + 1, i + 1).reduce((s, y) => s + y, 0) / p]);
      return { at: () => NaN, points };
    }
    default: {
      const l = linearFit(X, Y);
      if (!l) return null;
      return { at: (x) => l.slope * x + l.intercept, equation: equationText(l.slope, l.intercept), r2: l.r2 };
    }
  }
}
