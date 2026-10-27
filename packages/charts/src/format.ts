// Numbers as German Excel shows them in charts. Axis and data labels follow the source cells'
// number format ("Mit Quelle verknüpft"); the app passes its own formatter for that, and this
// file's General format (decimal comma, up to 10 significant digits) is the fallback.

export type NumberFormatter = (v: number, format?: string) => string;

/** Excel's General format with a decimal comma (no thousands separator). */
export function formatGeneral(v: number): string {
  if (!isFinite(v)) return "#ZAHL!";
  if (v === 0) return "0";
  const a = Math.abs(v);
  let s: string;
  if (a >= 1e11 || a < 1e-9) {
    const [m, e] = v.toExponential(5).split("e");
    const mant = m.replace(/\.?0+$/, "");
    const exp = Number(e);
    s = `${mant}E${exp < 0 ? "-" : "+"}${String(Math.abs(exp)).padStart(2, "0")}`;
  } else {
    s = String(Number(v.toPrecision(10)));
    if (/e/i.test(s)) s = v.toFixed(10).replace(/\.?0+$/, "");
  }
  return s.replace(".", ",");
}

/**
 * A coefficient of the trendline label as Excel writes it: five significant digits, at most
 * four decimals ("99,714", "2,6667", "0,0013", "11111").
 */
export function formatCoefficient(v: number): string {
  const a = Math.abs(v);
  if (a !== 0 && a < 0.00005) return formatGeneral(Number(v.toPrecision(5)));
  const intDigits = a >= 1 ? Math.floor(Math.log10(a)) + 1 : 1;
  const decimals = Math.min(4, Math.max(0, 5 - intDigits));
  let s = v.toFixed(decimals);
  if (s.includes(".")) s = s.replace(/0+$/, "").replace(/\.$/, "");
  if (/^-0$/.test(s)) s = "0";
  return s.replace(".", ",");
}

/** "y = 1,93x + 0,23", "y = -2x + 11", "y = 0,084x - 161,92", "y = 5". */
export function equationText(slope: number, intercept: number): string {
  const k = formatCoefficient(slope), d = formatCoefficient(Math.abs(intercept));
  if (k === "0") return `y = ${formatCoefficient(intercept)}`;
  const x = `${k === "1" ? "" : k === "-1" ? "-" : k}x`;
  if (d === "0") return `y = ${x}`;
  return `y = ${x} ${intercept < 0 ? "-" : "+"} ${d}`;
}

/** "R² = 0,9963", "R² = 1"; "#N/A" (in English, as recorded) when the values do not vary. */
export function r2Text(r2: number | null): string {
  if (r2 === null || !isFinite(r2)) return "R² = #N/A";
  let s = r2.toFixed(4).replace(/0+$/, "").replace(/\.$/, "");
  if (s === "-0") s = "0";
  return `R² = ${s.replace(".", ",")}`;
}
