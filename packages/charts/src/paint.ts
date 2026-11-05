// The Format pane's fills and lines as SVG paint (D-066): an element's own formatting where it has
// one, else Excel's automatic look passed in by the renderer.

import { autoScale, type AxisScale } from "./scale";
import type { AxisOptions, Dash, Fill, Line, TextFormat } from "./model";

/** A colour with Transparenz (0-100 %) as #RRGGBBAA. */
export function withAlpha(color: string, transparency?: number): string {
  if (!transparency) return color;
  const a = Math.round((1 - Math.max(0, Math.min(100, transparency)) / 100) * 255);
  return color.slice(0, 7) + a.toString(16).padStart(2, "0").toUpperCase();
}

/** The fill to draw: "none", the element's colour, or the automatic one. */
export function fillPaint(fill: Fill | undefined, auto: string | undefined): string | undefined {
  if (!fill || fill.type === "auto") return auto;
  if (fill.type === "none") return "none";
  return withAlpha(fill.color ?? auto ?? "#FFFFFF", fill.transparency);
}

export interface Stroke { stroke?: string; sw?: number; dash?: string }

/** The line to draw (null: no line), with DrawingML's preset dashes scaled to the width as Excel draws them. */
export function linePaint(line: Line | undefined, auto: { color?: string; width: number } | null): Stroke | null {
  if (line?.type === "none") return null;
  if (!line || line.type === "auto") {
    if (!auto?.color) return null;
    return { stroke: auto.color, sw: auto.width, dash: line?.dash ? dashArray(line.dash, auto.width) : undefined };
  }
  const sw = line.width ?? auto?.width ?? 0.75;
  return { stroke: withAlpha(line.color ?? auto?.color ?? "#000000", line.transparency), sw, dash: line.dash ? dashArray(line.dash, sw) : undefined };
}

/** DrawingML's preset dashes (lengths in line widths). */
export function dashArray(d: Dash, w: number): string | undefined {
  const pat: Record<Dash, number[]> = {
    solid: [], dot: [1, 3], sysDot: [1, 1], sysDash: [3, 1], sysDashDot: [3, 1, 1, 1], sysDashDotDot: [3, 1, 1, 1, 1, 1], dash: [4, 3], dashDot: [4, 3, 1, 3],
    lgDash: [8, 3], lgDashDot: [8, 3, 1, 3], lgDashDotDot: [8, 3, 1, 3, 1, 3],
  };
  // A preset the file has that is not in the list draws solid.
  const p = pat[d] ?? [];
  return p.length ? p.map((x) => String(Math.round(x * w * 100) / 100)).join(" ") : undefined;
}

/** Text paint: own size, colour and weight, else the automatic ones. */
export function textLook(t: TextFormat | undefined, auto: { size: number; color: string }): { size: number; fill: string; bold?: boolean; italic?: boolean } {
  return { size: t?.size ?? auto.size, fill: t?.color ?? auto.color, bold: t?.bold, italic: t?.italic };
}

/**
 * A value axis' scale with the pane's fixed Begrenzungen and Einheiten: what is fixed stays, the
 * rest is worked out around it as Excel does (the unit from the range, the bounds on the unit).
 */
export function fixedScale(lo: number, hi: number, o: AxisOptions | undefined, limit: number): AxisScale {
  const auto = autoScale(o?.min ?? lo, o?.max ?? hi, limit);
  if (!o || (o.min === undefined && o.max === undefined && o.majorUnit === undefined)) return auto;
  const major = o.majorUnit ?? (o.min !== undefined || o.max !== undefined ? autoScale(o.min ?? Math.min(lo, 0), o.max ?? hi, limit).major : auto.major);
  const round = (v: number) => Number(v.toFixed(10));
  const min = o.min ?? round(Math.floor(Math.min(lo, o.max ?? lo) / major + 1e-9) * major);
  let max = o.max ?? round(Math.ceil(Math.max(hi, min) / major - 1e-9) * major);
  if (max <= min) max = round(min + major);
  return { min, max, major };
}
