// Our own animated GIF writer (D-074): GIF89a with the NETSCAPE2.0 loop extension, one local colour table of up to
// 256 colours per frame (median cut on the frame's pixels), LZW-compressed image data, and only the changed
// rectangle of each frame after the first (unchanged pixels transparent, disposal "keep"). A frame equal to the
// last one only lengthens the last one's delay. "Hintergrund transparent machen": pixels of the slide's
// background colour become transparent (PowerPoint's option on its GIF page).

export interface GifFrame { rgba: Uint8ClampedArray; /** Hundredths of a second. */ delay: number }

const enc = new TextEncoder();

export class GifWriter {
  private parts: Uint8Array[] = [];
  private prev: Uint8ClampedArray | null = null;
  /** The last frame's graphic control extension, to lengthen its delay. */
  private lastGce: Uint8Array | null = null;
  private lastDelay = 0;

  constructor(private width: number, private height: number, loop = true, private transparentBg: [number, number, number] | null = null) {
    const head = new Uint8Array(13);
    head.set(enc.encode("GIF89a"));
    const v = new DataView(head.buffer);
    v.setUint16(6, width, true); v.setUint16(8, height, true);
    head[10] = 0x70; // no global colour table, colour resolution 8 bit
    this.parts.push(head);
    if (loop) this.parts.push(Uint8Array.of(0x21, 0xff, 0x0b, ...enc.encode("NETSCAPE2.0"), 0x03, 0x01, 0x00, 0x00, 0x00));
  }

  add(frame: GifFrame): void {
    const { width: W, height: H } = this;
    const px = frame.rgba;
    // Unchanged frame: the last one stays longer.
    if (this.prev && this.lastGce && equal(this.prev, px)) {
      this.lastDelay = Math.min(65535, this.lastDelay + frame.delay);
      new DataView(this.lastGce.buffer, this.lastGce.byteOffset).setUint16(4, this.lastDelay, true);
      return;
    }
    // The changed rectangle (all of it for the first frame).
    let x0 = 0, y0 = 0, x1 = W - 1, y1 = H - 1;
    if (this.prev && !this.transparentBg) {
      x0 = W; y0 = H; x1 = -1; y1 = -1;
      for (let y = 0; y < H; y++) for (let x = 0; x < W; x++) {
        const i = (y * W + x) * 4;
        if (px[i] !== this.prev[i] || px[i + 1] !== this.prev[i + 1] || px[i + 2] !== this.prev[i + 2] || px[i + 3] !== this.prev[i + 3]) {
          if (x < x0) x0 = x; if (x > x1) x1 = x; if (y < y0) y0 = y; if (y > y1) y1 = y;
        }
      }
    }
    const w = x1 - x0 + 1, h = y1 - y0 + 1;
    // Pixels of the rectangle: transparent where unchanged (after the first frame) or of the background colour.
    const keep = new Uint8Array(w * h);
    const rgb: number[] = [];
    for (let y = 0; y < h; y++) for (let x = 0; x < w; x++) {
      const i = ((y + y0) * W + (x + x0)) * 4;
      const same = !!this.prev && px[i] === this.prev[i] && px[i + 1] === this.prev[i + 1] && px[i + 2] === this.prev[i + 2];
      const bg = !!this.transparentBg && px[i] === this.transparentBg[0] && px[i + 1] === this.transparentBg[1] && px[i + 2] === this.transparentBg[2];
      const clear = px[i + 3] < 128;
      if ((same && !this.transparentBg) || bg || clear) continue;
      keep[y * w + x] = 1;
      rgb.push((px[i] << 16) | (px[i + 1] << 8) | px[i + 2]);
    }
    const palette = medianCut(rgb, 255);
    const lookup = new Map<number, number>();
    const index = new Uint8Array(w * h);
    const TRANSPARENT = palette.length;
    for (let y = 0; y < h; y++) for (let x = 0; x < w; x++) {
      const k = y * w + x;
      if (!keep[k]) { index[k] = TRANSPARENT; continue; }
      const i = ((y + y0) * W + (x + x0)) * 4;
      const c = (px[i] << 16) | (px[i + 1] << 8) | px[i + 2];
      let n = lookup.get(c);
      if (n === undefined) { n = nearest(palette, c); lookup.set(c, n); }
      index[k] = n;
    }
    // Table size: a power of two holding the palette and the transparent entry.
    let bits = 1;
    while (1 << bits < palette.length + 1) bits++;
    // Disposal: keep the picture (changed rectangles), or clear it when the background is transparent.
    const gce = Uint8Array.of(0x21, 0xf9, 0x04, ((this.transparentBg ? 2 : 1) << 2) | 1, 0, 0, TRANSPARENT, 0x00);
    new DataView(gce.buffer).setUint16(4, Math.max(2, frame.delay), true);
    const desc = new Uint8Array(10);
    const dv = new DataView(desc.buffer);
    desc[0] = 0x2c; dv.setUint16(1, x0, true); dv.setUint16(3, y0, true); dv.setUint16(5, w, true); dv.setUint16(7, h, true);
    desc[9] = 0x80 | (bits - 1);
    const table = new Uint8Array(3 << bits);
    palette.forEach((c, i) => { table[i * 3] = (c >> 16) & 255; table[i * 3 + 1] = (c >> 8) & 255; table[i * 3 + 2] = c & 255; });
    const minCode = Math.max(2, bits);
    this.parts.push(gce, desc, table, Uint8Array.of(minCode), ...subBlocks(lzw(index, minCode)));
    this.lastGce = gce;
    this.lastDelay = Math.max(2, frame.delay);
    // What the viewer shows now: the previous picture with this rectangle drawn over it.
    if (!this.prev || this.transparentBg) this.prev = new Uint8ClampedArray(px);
    else for (let y = y0; y <= y1; y++) this.prev.set(px.subarray((y * W + x0) * 4, (y * W + x1 + 1) * 4), (y * W + x0) * 4);
  }

  finish(): Uint8Array {
    this.parts.push(Uint8Array.of(0x3b));
    const out = new Uint8Array(this.parts.reduce((n, p) => n + p.length, 0));
    let o = 0;
    for (const p of this.parts) { out.set(p, o); o += p.length; }
    return out;
  }
}

function equal(a: Uint8ClampedArray, b: Uint8ClampedArray): boolean {
  if (a.length !== b.length) return false;
  for (let i = 0; i < a.length; i++) if (a[i] !== b[i]) return false;
  return true;
}

/** Median cut: boxes of colours split at the median of their widest channel until there are `max` boxes. */
export function medianCut(colors: number[], max: number): number[] {
  if (!colors.length) return [0];
  // Count each colour once, with its weight.
  const counts = new Map<number, number>();
  for (const c of colors) counts.set(c, (counts.get(c) ?? 0) + 1);
  if (counts.size <= max) return [...counts.keys()];
  type Box = { cols: number[] };
  const boxes: Box[] = [{ cols: [...counts.keys()] }];
  const ch = (c: number, k: number) => (c >> (16 - k * 8)) & 255;
  const range = (b: Box) => {
    let best = 0, axis = 0;
    for (let k = 0; k < 3; k++) {
      let lo = 255, hi = 0;
      for (const c of b.cols) { const v = ch(c, k); if (v < lo) lo = v; if (v > hi) hi = v; }
      if (hi - lo > best) { best = hi - lo; axis = k; }
    }
    return { best, axis };
  };
  while (boxes.length < max) {
    // Split the box with the most pixels among those that can still be split.
    let pick = -1, weight = -1;
    for (let i = 0; i < boxes.length; i++) {
      if (boxes[i].cols.length < 2) continue;
      const wgt = boxes[i].cols.reduce((n, c) => n + counts.get(c)!, 0) * (range(boxes[i]).best + 1);
      if (wgt > weight) { weight = wgt; pick = i; }
    }
    if (pick < 0) break;
    const b = boxes[pick];
    const { axis } = range(b);
    b.cols.sort((p, q) => ch(p, axis) - ch(q, axis));
    const total = b.cols.reduce((n, c) => n + counts.get(c)!, 0);
    let acc = 0, cut = 1;
    for (let i = 0; i < b.cols.length - 1; i++) { acc += counts.get(b.cols[i])!; if (acc >= total / 2) { cut = i + 1; break; } }
    boxes.splice(pick, 1, { cols: b.cols.slice(0, cut) }, { cols: b.cols.slice(cut) });
  }
  // Each box's colour: the weighted mean.
  return boxes.map((b) => {
    let r = 0, g = 0, bl = 0, n = 0;
    for (const c of b.cols) { const w = counts.get(c)!; r += ((c >> 16) & 255) * w; g += ((c >> 8) & 255) * w; bl += (c & 255) * w; n += w; }
    return (Math.round(r / n) << 16) | (Math.round(g / n) << 8) | Math.round(bl / n);
  });
}

function nearest(palette: number[], c: number): number {
  const r = (c >> 16) & 255, g = (c >> 8) & 255, b = c & 255;
  let best = 0, dist = Infinity;
  for (let i = 0; i < palette.length; i++) {
    const p = palette[i];
    const dr = ((p >> 16) & 255) - r, dg = ((p >> 8) & 255) - g, db = (p & 255) - b;
    const d = dr * dr * 2 + dg * dg * 4 + db * db * 3;
    if (d < dist) { dist = d; best = i; if (d === 0) break; }
  }
  return best;
}

/** GIF's variable-length LZW, codes packed least significant bit first. */
export function lzw(index: Uint8Array, minCode: number): Uint8Array {
  const clear = 1 << minCode, eoi = clear + 1;
  const out: number[] = [];
  let cur = 0, curBits = 0;
  let size = minCode + 1;
  const emit = (code: number) => {
    cur |= code << curBits; curBits += size;
    while (curBits >= 8) { out.push(cur & 255); cur >>>= 8; curBits -= 8; }
  };
  let dict = new Map<number, number>();
  let next = eoi + 1;
  emit(clear);
  let prefix = index.length ? index[0] : 0;
  for (let i = 1; i < index.length; i++) {
    const k = index[i];
    const key = prefix * 4096 + k;
    const found = dict.get(key);
    if (found !== undefined) { prefix = found; continue; }
    emit(prefix);
    if (next < 4096) {
      dict.set(key, next++);
      if (next > 1 << size && size < 12) size++;
    } else {
      emit(clear);
      dict = new Map(); next = eoi + 1; size = minCode + 1;
    }
    prefix = k;
  }
  if (index.length) emit(prefix);
  emit(eoi);
  if (curBits > 0) out.push(cur & 255);
  return Uint8Array.from(out);
}

function subBlocks(data: Uint8Array): Uint8Array[] {
  const out: Uint8Array[] = [];
  for (let o = 0; o < data.length; o += 255) { const n = Math.min(255, data.length - o); out.push(Uint8Array.of(n), data.subarray(o, o + n)); }
  out.push(Uint8Array.of(0));
  return out;
}
