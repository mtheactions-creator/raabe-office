// Column widths and row heights on paper as Excel prints them (D-094), recorded from Excel's PDF export in our
// own Excel (tools/feel-lab/probe_excel_print.py, results/2026-11-08-excel-print.json). Excel lays the printed page
// out at 600 dpi:
// - a column is the screen's width in digits (its device pixels over the screen's digit width) times the digit
//   width of the standard font at 600 dpi (Aptos Narrow 11 and Calibri 11: 47 px), with Excel's rounding;
// - a row of the standard height prints at the standard font's printed line height (Aptos Narrow 11: 127 px =
//   15,24 pt, Calibri 11: 121 px = 14,52 pt); a row of its own height h at h times printed / screen standard
//   height, rounded down to whole 600-dpi pixels (Aptos Narrow 11: a 30-pt row prints 31,44 pt).
// So a sheet in Aptos Narrow prints about 5 % taller than its row heights say, one in Calibri as they say.

import { device } from "./grid/metrics";

/** Printed standard row height (pt) and Excel's screen standard height at 150 % (pt), per font and size. */
const ROWS: Record<string, [number, number, number][]> = {
  // [size, printed, screen]
  "aptos narrow": [[9, 12.72, 12], [10, 13.92, 13], [11, 15.24, 14.5], [12, 16.44, 16], [14, 19.08, 18.5]],
  aptos: [[9, 12.72, 12], [10, 13.92, 13], [11, 15.24, 14.5], [12, 16.44, 16], [14, 19.08, 18.5]],
  calibri: [[9, 11.88, 12], [10, 13.08, 13], [11, 14.52, 14.5], [12, 15.6, 15.5], [14, 18.0, 18.5]],
  "calibri light": [[9, 11.88, 12], [10, 13.08, 13], [11, 14.52, 14.5], [12, 15.6, 15.5], [14, 18.0, 18.5]],
  arial: [[9, 11.16, 11.5], [10, 12.36, 12.5], [11, 13.68, 14], [12, 14.76, 15.5], [14, 17.16, 17.5]],
  "times new roman": [[9, 11.52, 11.5], [10, 12.72, 13], [11, 13.92, 14], [12, 15.12, 15.5], [14, 17.4, 18]],
  cambria: [[9, 11.52, 11.5], [10, 12.6, 12.5], [11, 13.8, 14], [12, 15.0, 15], [14, 17.4, 17.5]],
  "segoe ui": [[9, 13.68, 14], [10, 14.88, 16], [11, 16.56, 16.5], [12, 17.76, 17.5], [14, 20.52, 21]],
  verdana: [[9, 12.0, 11.5], [10, 12.96, 13.5], [11, 14.64, 13.5], [12, 15.48, 15], [14, 18.0, 17.5]],
  tahoma: [[9, 11.76, 11.5], [10, 12.96, 12.5], [11, 14.28, 14], [12, 15.48, 15], [14, 17.88, 17.5]],
  consolas: [[9, 11.52, 12], [10, 12.6, 13], [11, 13.92, 14.5], [12, 15.0, 15.5], [14, 17.4, 18]],
  georgia: [[9, 11.28, 11.5], [10, 12.48, 13], [11, 13.56, 14.5], [12, 14.64, 15.5], [14, 16.92, 17.5]],
};

const DPI = 600;
/** 600-dpi printer pixels → page px (96 dpi). */
export const PRN = 96 / DPI;

/** Printed and screen standard row heights (pt) of a font; unknown fonts print as their rows say. */
export function rowHeights(font: string, size: number, fileDefault: number): { printed: number; screen: number } {
  const t = ROWS[font.toLowerCase()];
  if (!t) return { printed: fileDefault, screen: fileDefault };
  // Between the recorded sizes and beyond them: along the line through the nearest two (Excel's line height
  // grows by about 1,3 printer pixels per pixel of font size; Aptos Narrow 16 prints 21,6 pt as recorded).
  let a = size < t[0][0] ? t[0] : t[t.length - 2], b = size < t[0][0] ? t[1] : t[t.length - 1];
  for (let i = 0; i < t.length - 1; i++) if (size >= t[i][0] && size <= t[i + 1][0]) { a = t[i]; b = t[i + 1]; break; }
  const k = b[0] === a[0] ? 0 : (size - a[0]) / (b[0] - a[0]);
  return { printed: a[1] + (b[1] - a[1]) * k, screen: a[2] + (b[2] - a[2]) * k };
}

let ruler: CanvasRenderingContext2D | null = null;
const digitCache = new Map<string, number>();

/** The standard font's widest digit at 600 dpi (Excel: the font at whole pixels, the digit rounded). */
export function printerDigit(font: string, size: number): number {
  const key = `${font}|${size}`;
  let w = digitCache.get(key);
  if (w !== undefined) return w;
  const em = Math.round((size * DPI) / 72);
  let frac = 0;
  try {
    ruler ??= document.createElement("canvas").getContext("2d");
    if (ruler) {
      ruler.font = `1000px "${font}", Carlito, sans-serif`;
      frac = Math.max(...[..."0123456789"].map((d) => ruler!.measureText(d).width)) / 1000;
    }
  } catch { /* no canvas (tests) */ }
  if (!(frac > 0.2 && frac < 1)) frac = 0.507; // Calibri's digits
  w = Math.round(frac * em);
  digitCache.set(key, w);
  return w;
}

/** Excel's width rounding: characters (with padding) times a digit width, in whole pixels. */
function widthPx(chars: number, mdw: number): number {
  return Math.trunc(((256 * chars + Math.trunc(128 / mdw)) / 256) * mdw);
}

/** A column's printed width in page px from its width on screen (CSS px at 100 %). */
export function printColPx(screenPx: number, digit600: number): number {
  if (screenPx <= 0) return 0;
  const d = device();
  const chars = (screenPx * d.dpr) / d.mdw;
  return widthPx(chars, digit600) * PRN;
}

/** A row's printed height in page px: the standard height prints as the font's printed line height. */
export function printRowPx(pt: number | undefined, std: { printed: number; screen: number }): number {
  if (pt === undefined) return Math.round((std.printed * DPI) / 72) * PRN;
  if (pt <= 0) return 0;
  // The printed line height is whole pixels rounded down from a little more (recorded: with Aptos Narrow 11 a
  // 21-pt row prints 184 px and a 15,75-pt row 138 px, which 127 / 120,83 would make 183 and 137).
  const k = (Math.round((std.printed * DPI) / 72) + 0.08) / ((std.screen * DPI) / 72);
  return Math.floor(((pt * DPI) / 72) * k + 1e-6) * PRN;
}
