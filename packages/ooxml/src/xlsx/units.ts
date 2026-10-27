// Column widths: the file stores characters of the default font's widest digit plus 5 pixels
// of padding, in 1/256 steps; the model (and Excel's Spaltenbreite at 100 % and 96 dpi)
// counts characters without the padding. Calibri 11's digits are 7 pixels wide.

export const MDW = 7;

/** Model width (characters) of a file width. */
export function widthFromFile(w: number, mdw = MDW): number {
  return w <= 0 ? 0 : Math.max(0, w - 5 / mdw);
}

/** File width of a model width, as Excel computes it (truncated to 1/256). */
export function widthToFile(chars: number, mdw = MDW): number {
  if (chars <= 0) return 0;
  const x = (chars + 5 / mdw) * 256;
  // Widths read from a file come back exactly; widths typed in are truncated as by Excel.
  const r = Math.round(x);
  return (Math.abs(x - r) < 1e-6 ? r : Math.floor(x)) / 256;
}

/** The standard width from sheetFormatPr's baseColWidth: padded and rounded up to 8 pixels. */
export function widthFromBase(base: number, mdw = MDW): number {
  const px = Math.ceil((base * mdw + 5) / 8) * 8;
  return (px - 5) / mdw;
}

/** A number as Excel writes it in XML (shortest round-trip form, "E" exponent). */
export function numText(n: number): string {
  return String(n).replace("e", "E");
}
