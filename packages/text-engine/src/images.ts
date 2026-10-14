// Picture headers: pixel size and resolution of PNG, JPEG, GIF and BMP files, read without
// decoding the picture. Word sizes an inserted picture from its resolution (PNG pHYs, JPEG
// JFIF density, BMP pixels per metre) and assumes 96 dpi when there is none.

export interface ImageInfo {
  readonly type: string;
  readonly px: { readonly w: number; readonly h: number };
  /** Dots per inch; 96 when the file says nothing. */
  readonly dpi: { readonly x: number; readonly y: number };
}

const u16be = (b: Uint8Array, i: number) => (b[i] << 8) | b[i + 1];
const u32be = (b: Uint8Array, i: number) => ((b[i] << 24) >>> 0) + (b[i + 1] << 16) + (b[i + 2] << 8) + b[i + 3];
const u16le = (b: Uint8Array, i: number) => b[i] | (b[i + 1] << 8);
const i32le = (b: Uint8Array, i: number) => b[i] | (b[i + 1] << 8) | (b[i + 2] << 16) | (b[i + 3] << 24);

/** Resolutions outside this range are treated as missing (broken files say 1 or 0 dpi). */
const plausible = (d: number) => d >= 24 && d <= 2400;

export function imageInfo(b: Uint8Array): ImageInfo | null {
  let dx = 96, dy = 96;
  // PNG: IHDR, then pHYs (unit 1 = per metre) somewhere before IDAT.
  if (b.length > 24 && b[0] === 0x89 && b[1] === 0x50 && b[2] === 0x4e && b[3] === 0x47) {
    const w = u32be(b, 16), h = u32be(b, 20);
    for (let i = 8; i + 8 <= b.length;) {
      const n = u32be(b, i);
      const t = String.fromCharCode(b[i + 4], b[i + 5], b[i + 6], b[i + 7]);
      if (t === "pHYs" && n >= 9 && b[i + 16] === 1) {
        const x = u32be(b, i + 8) * 0.0254, y = u32be(b, i + 12) * 0.0254;
        if (plausible(x) && plausible(y)) { dx = x; dy = y; }
      }
      if (t === "IDAT" || t === "IEND") break;
      i += 12 + n;
    }
    return { type: "image/png", px: { w, h }, dpi: { x: dx, y: dy } };
  }
  // JPEG: JFIF density in APP0, the size in the first SOFn segment.
  if (b.length > 4 && b[0] === 0xff && b[1] === 0xd8) {
    for (let i = 2; i + 4 <= b.length;) {
      if (b[i] !== 0xff) { i++; continue; }
      const m = b[i + 1];
      if (m === 0xd8 || m === 0x01 || (m >= 0xd0 && m <= 0xd7)) { i += 2; continue; }
      const len = u16be(b, i + 2);
      if (m === 0xe0 && b[i + 4] === 0x4a && b[i + 5] === 0x46 && b[i + 6] === 0x49 && b[i + 7] === 0x46) {
        const unit = b[i + 11], x = u16be(b, i + 12), y = u16be(b, i + 14);
        const k = unit === 1 ? 1 : unit === 2 ? 2.54 : 0;
        if (k && plausible(x * k) && plausible(y * k)) { dx = x * k; dy = y * k; }
      }
      if (m >= 0xc0 && m <= 0xcf && m !== 0xc4 && m !== 0xc8 && m !== 0xcc) {
        return { type: "image/jpeg", px: { w: u16be(b, i + 7), h: u16be(b, i + 5) }, dpi: { x: dx, y: dy } };
      }
      i += 2 + len;
    }
    return null;
  }
  // GIF: logical screen size, no resolution.
  if (b.length > 10 && b[0] === 0x47 && b[1] === 0x49 && b[2] === 0x46) {
    return { type: "image/gif", px: { w: u16le(b, 6), h: u16le(b, 8) }, dpi: { x: 96, y: 96 } };
  }
  // BMP: BITMAPINFOHEADER with pixels per metre.
  if (b.length > 46 && b[0] === 0x42 && b[1] === 0x4d) {
    const w = Math.abs(i32le(b, 18)), h = Math.abs(i32le(b, 22));
    const x = i32le(b, 38) * 0.0254, y = i32le(b, 42) * 0.0254;
    if (plausible(x) && plausible(y)) { dx = x; dy = y; }
    return { type: "image/bmp", px: { w, h }, dpi: { x: dx, y: dy } };
  }
  return null;
}

/** A picture's natural size in points, as Word inserts it. */
export function naturalSizePt(info: ImageInfo): { w: number; h: number } {
  return { w: (info.px.w * 72) / info.dpi.x, h: (info.px.h * 72) / info.dpi.y };
}
