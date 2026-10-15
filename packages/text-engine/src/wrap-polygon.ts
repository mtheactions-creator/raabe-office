// Wrap polygons for Textumbruch "Eng" and "Transparent" (wp:wrapPolygon), made as Word makes
// them: around a shape's outline, around the visible part of a picture with a transparent
// background, else around the box. Points are in 21600ths of the object's unturned box; the
// layout turns them with the object (measured, D-054). Only the left and right edge of each
// band matter for the text, so the outline is taken band by band.

import { flattenPath, presetGeometry } from "./geometry";
import type { FloatObject, ImageData } from "./model";
import { shapeGeometry } from "./shape";

export type Polygon = [number, number][];

const BANDS = 48;
const UNIT = 21600;

/** A polygon through the leftmost and rightmost point of each band's middle (null rows are empty). */
function fromBands(rows: ([number, number] | null)[]): Polygon | null {
  const n = rows.length;
  const left: [number, number][] = [], right: [number, number][] = [];
  let first = -1, last = -1;
  rows.forEach((r, i) => {
    if (!r) return;
    if (first < 0) first = i;
    last = i;
    const y = ((i + 0.5) / n) * UNIT;
    left.push([r[0], y]);
    right.push([r[1], y]);
  });
  if (first < 0) return null;
  // The outline's top and bottom edges.
  const top = (first / n) * UNIT, bottom = ((last + 1) / n) * UNIT;
  left.unshift([rows[first]![0], top]);
  right.unshift([rows[first]![1], top]);
  left.push([rows[last]![0], bottom]);
  right.push([rows[last]![1], bottom]);
  const pts = simplify([...left, ...right.reverse()]);
  return [...pts, pts[0]];
}

/** Drops points that lie (nearly) on the line between their neighbours. */
function simplify(pts: Polygon, tol = 60): Polygon {
  const out: Polygon = [];
  for (const p of pts) {
    if (out.length && out[out.length - 1][0] === p[0] && out[out.length - 1][1] === p[1]) continue;
    out.push(p);
    while (out.length >= 3) {
      const [a, b, c] = out.slice(-3);
      const dx = c[0] - a[0], dy = c[1] - a[1];
      const len = Math.hypot(dx, dy) || 1;
      if (Math.abs((b[0] - a[0]) * dy - (b[1] - a[1]) * dx) / len > tol) break;
      out.splice(out.length - 2, 1);
    }
  }
  return out.map(([x, y]) => [Math.round(x), Math.round(y)]);
}

/** The box (Word's default for pictures without transparency). */
export const BOX: Polygon = [[0, 0], [0, UNIT], [UNIT, UNIT], [UNIT, 0], [0, 0]];

/** The outline of a shape or of a picture cut to a shape (Auf Form zuschneiden); null: the box. */
export function outlinePolygon(f: FloatObject): Polygon | null {
  const w = Math.max(1, f.width), h = Math.max(1, f.height);
  const geom = f.shape ? shapeGeometry(f.pic, f.shape, w, h)
    : f.pic?.geometry && f.pic.geometry !== "rect" ? presetGeometry(f.pic.geometry, w, h) : null;
  if (!geom) return null;
  const pts: [number, number][] = [];
  for (const p of geom.paths) if (p.fill !== "none" || p.stroke) pts.push(...flattenPath(p.d, 12));
  if (pts.length < 3) return null;
  // Band by band: the outline's extent between the band's top and bottom.
  const rows: ([number, number] | null)[] = [];
  for (let i = 0; i < BANDS; i++) {
    const y0 = (i / BANDS) * h, y1 = ((i + 1) / BANDS) * h;
    let lo = Infinity, hi = -Infinity;
    for (let k = 0; k + 1 < pts.length; k++) {
      const [ax, ay] = pts[k], [bx, by] = pts[k + 1];
      if (Math.max(ay, by) < y0 || Math.min(ay, by) > y1) continue;
      let t0 = 0, t1 = 1;
      if (by !== ay) {
        const ta = (y0 - ay) / (by - ay), tb = (y1 - ay) / (by - ay);
        t0 = Math.max(0, Math.min(ta, tb)); t1 = Math.min(1, Math.max(ta, tb));
        if (t0 > t1) continue;
      }
      for (const t of [t0, t1]) { const x = ax + (bx - ax) * t; lo = Math.min(lo, x); hi = Math.max(hi, x); }
    }
    rows.push(lo <= hi ? [(lo / w) * UNIT, (hi / w) * UNIT] : null);
  }
  return fromBands(rows);
}

/**
 * The visible part of a picture (pixels more than a tenth opaque), when it has a transparent
 * background; null for pictures without transparency (the box) or that cannot be read here.
 */
export async function alphaPolygon(image: ImageData, crop?: { l: number; t: number; r: number; b: number }): Promise<Polygon | null> {
  if (!/png|gif|webp/i.test(image.type) || typeof createImageBitmap !== "function") return null;
  let bmp: ImageBitmap;
  try {
    bmp = await createImageBitmap(new Blob([image.bytes as BlobPart], { type: image.type }));
  } catch {
    return null;
  }
  const c = crop ?? { l: 0, t: 0, r: 0, b: 0 };
  const sx = c.l * bmp.width, sy = c.t * bmp.height;
  const sw = Math.max(1, bmp.width * (1 - c.l - c.r)), sh = Math.max(1, bmp.height * (1 - c.t - c.b));
  const W = 216, H = Math.max(1, Math.round((216 * sh) / sw));
  const canvas = typeof OffscreenCanvas !== "undefined" ? new OffscreenCanvas(W, H) : Object.assign(document.createElement("canvas"), { width: W, height: H });
  const g = canvas.getContext("2d") as CanvasRenderingContext2D | OffscreenCanvasRenderingContext2D | null;
  if (!g) return null;
  g.drawImage(bmp, sx, sy, sw, sh, 0, 0, W, H);
  bmp.close?.();
  const px = g.getImageData(0, 0, W, H).data;
  let transparent = false;
  const rows: ([number, number] | null)[] = [];
  for (let i = 0; i < BANDS; i++) {
    const y0 = Math.floor((i / BANDS) * H), y1 = Math.max(y0 + 1, Math.floor(((i + 1) / BANDS) * H));
    let lo = Infinity, hi = -Infinity;
    for (let y = y0; y < y1 && y < H; y++) {
      for (let x = 0; x < W; x++) {
        const a = px[(y * W + x) * 4 + 3];
        if (a < 255) transparent = true;
        if (a > 25) { lo = Math.min(lo, x); hi = Math.max(hi, x + 1); }
      }
    }
    rows.push(lo <= hi ? [(lo / W) * UNIT, (hi / W) * UNIT] : null);
  }
  if (!transparent) return null;
  // Word's polygon stands about 3 % outside the visible part (measured on a cut-out circle:
  // 295–21207 for a circle from 1080 to 20520).
  const pad = 650;
  const first = rows.findIndex((r) => r), last = rows.length - 1 - [...rows].reverse().findIndex((r) => r);
  const grow = Math.max(1, Math.round((pad / UNIT) * BANDS));
  const out = rows.map((r) => (r ? [Math.max(0, r[0] - pad), Math.min(UNIT, r[1] + pad)] as [number, number] : null));
  for (let k = 1; k <= grow && first >= 0; k++) {
    if (first - k >= 0 && !out[first - k]) out[first - k] = out[first];
    if (last + k < out.length && !out[last + k]) out[last + k] = out[last];
  }
  return fromBands(out);
}
