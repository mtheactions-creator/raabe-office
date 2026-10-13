// Page and style geometry, taken from Word's own defaults on the development PC
// (read via COM on 2026-10-13): A4, margins 2.5 / 2.0 / 2.5 / 2.5 cm, Normal style
// Calibri 11 pt, line spacing "multiple 1.08", 8 pt after. All values in CSS px at 100 %
// zoom (1 pt = 4/3 px), so the page has the same physical size as in Word.

export const PT = 4 / 3;

export const PAGE = {
  width: 595.3 * PT,
  height: 841.9 * PT,
  marginTop: 70.85 * PT,
  marginBottom: 56.7 * PT,
  marginLeft: 70.85 * PT,
  marginRight: 70.85 * PT,
  gap: 16,
};

export const CONTENT_WIDTH = PAGE.width - PAGE.marginLeft - PAGE.marginRight;
export const CONTENT_HEIGHT = PAGE.height - PAGE.marginTop - PAGE.marginBottom;

export type StyleId = "normal" | "h1";

export interface ParaStyle {
  family: string;
  sizePt: number;
  bold: boolean;
  color: string;
  /** Word's "single" line height as a multiple of the font size (winAscent + winDescent). */
  singleEm: number;
  /** Ascent part of singleEm, used to place the baseline. */
  ascentEm: number;
  lineMultiple: number;
  spaceBeforePt: number;
  spaceAfterPt: number;
  keepWithNext: boolean;
}

// Calibri and Calibri Light: winAscent 1950, winDescent 550, unitsPerEm 2048.
const CALIBRI_SINGLE = (1950 + 550) / 2048;
const CALIBRI_ASCENT = 1950 / 2048;

export const STYLES: Record<StyleId, ParaStyle> = {
  normal: {
    family: "Calibri, Carlito, sans-serif", sizePt: 11, bold: false, color: "#000000",
    singleEm: CALIBRI_SINGLE, ascentEm: CALIBRI_ASCENT, lineMultiple: 1.08,
    spaceBeforePt: 0, spaceAfterPt: 8, keepWithNext: false,
  },
  h1: {
    family: "'Calibri Light', Carlito, sans-serif", sizePt: 16, bold: false, color: "#2F5496",
    singleEm: CALIBRI_SINGLE, ascentEm: CALIBRI_ASCENT, lineMultiple: 1.08,
    spaceBeforePt: 12, spaceAfterPt: 0, keepWithNext: true,
  },
};

export function fontCss(s: ParaStyle): string {
  return `${s.bold ? "bold " : ""}${s.sizePt * PT}px ${s.family}`;
}

export function lineHeightPx(s: ParaStyle): number {
  return s.sizePt * PT * s.singleEm * s.lineMultiple;
}
