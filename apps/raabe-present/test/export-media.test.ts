// Video and GIF export (D-074): our MP4 and GIF writers.
import { describe, expect, it } from "vitest";
import { Mp4Writer, readBoxes, type Mp4Box } from "../src/mp4";
import { GifWriter, lzw, medianCut } from "../src/gif";

const find = (boxes: Mp4Box[], path: string[]): Mp4Box[] => {
  const [head, ...rest] = path;
  const hits = boxes.filter((b) => b.type === head);
  return rest.length ? hits.flatMap((b) => find(b.children ?? [], rest)) : hits;
};

describe("MP4 writer", () => {
  const w = new Mp4Writer();
  const v = w.addVideo(640, 360, Uint8Array.of(1, 0x64, 0, 0x1f, 0xff, 0xe0, 0));
  const a = w.addAudio(48000, 2, Uint8Array.of(0x11, 0x90));
  for (let i = 0; i < 30; i++) w.addSample(v, { data: new Uint8Array(100 + i), time: (i * 1e6) / 30, duration: 1e6 / 30, key: i % 10 === 0 });
  for (let i = 0; i < 47; i++) w.addSample(a, { data: new Uint8Array(20), time: (i * 1024 * 1e6) / 48000, duration: (1024 * 1e6) / 48000, key: true });
  const file = w.finish();
  const boxes = readBoxes(file);

  it("has ftyp, mdat and moov in this order, covering the whole file", () => {
    expect(boxes.map((b) => b.type)).toEqual(["ftyp", "mdat", "moov"]);
    expect(boxes.reduce((n, b) => n + b.size, 0)).toBe(file.length);
  });

  it("has a video and an audio track with their sample entries", () => {
    expect(find(boxes, ["moov", "trak"]).length).toBe(2);
    const stbl = find(boxes, ["moov", "trak", "mdia", "minf", "stbl"]);
    expect(stbl.length).toBe(2);
    const types = stbl.map((s) => (s.children ?? []).map((c) => c.type).join(","));
    expect(types[0]).toBe("stsd,stts,stss,stsc,stsz,stco");
    expect(types[1]).toBe("stsd,stts,stsc,stsz,stco");
    const text = new TextDecoder("latin1").decode(file);
    expect(text).toContain("avc1");
    expect(text).toContain("avcC");
    expect(text).toContain("mp4a");
    expect(text).toContain("esds");
  });

  it("points each chunk offset at its sample's bytes in mdat", () => {
    const stco = find(boxes, ["moov", "trak", "mdia", "minf", "stbl", "stco"])[0];
    const dv = new DataView(file.buffer);
    const n = dv.getUint32(stco.start + 12);
    expect(n).toBe(30);
    const mdat = boxes[1];
    const first = dv.getUint32(stco.start + 16), second = dv.getUint32(stco.start + 20);
    expect(first).toBe(mdat.start + 8);
    expect(second - first).toBe(100);
  });

  it("gives the movie the length of its samples", () => {
    const mvhd = find(boxes, ["moov", "mvhd"])[0];
    const dv = new DataView(file.buffer);
    // version 0: timescale at +20, duration at +24 (1000 per second).
    expect(dv.getUint32(mvhd.start + 20)).toBe(1000);
    expect(dv.getUint32(mvhd.start + 24)).toBe(1003);
  });
});

/** A small GIF decoder for the test: frames as index arrays, with their rectangles and delays. */
function decodeGif(b: Uint8Array) {
  let o = 13;
  const frames: { x: number; y: number; w: number; h: number; delay: number; pixels: number[]; table: number[] }[] = [];
  let delay = 0;
  let loop = false;
  while (o < b.length) {
    const t = b[o];
    if (t === 0x3b) break;
    if (t === 0x21) {
      const label = b[o + 1];
      if (label === 0xf9) delay = b[o + 4] | (b[o + 5] << 8);
      if (label === 0xff) loop = true;
      o += 2;
      while (b[o]) o += b[o] + 1;
      o++;
      continue;
    }
    if (t !== 0x2c) throw new Error(`bad block ${t} at ${o}`);
    const x = b[o + 1] | (b[o + 2] << 8), y = b[o + 3] | (b[o + 4] << 8), w = b[o + 5] | (b[o + 6] << 8), h = b[o + 7] | (b[o + 8] << 8);
    const packed = b[o + 9];
    o += 10;
    const size = 3 << ((packed & 7) + 1);
    const table = Array.from(b.subarray(o, o + size));
    o += size;
    const min = b[o++];
    const data: number[] = [];
    while (b[o]) { data.push(...b.subarray(o + 1, o + 1 + b[o])); o += b[o] + 1; }
    o++;
    // LZW decode.
    const clear = 1 << min, eoi = clear + 1;
    let size2 = min + 1, dict: number[][] = [], prev: number[] | null = null;
    const reset = () => { dict = Array.from({ length: eoi + 1 }, (_, i) => [i]); size2 = min + 1; prev = null; };
    reset();
    let bit = 0;
    const out: number[] = [];
    for (;;) {
      let code = 0;
      for (let i = 0; i < size2; i++) { code |= ((data[(bit >> 3)] >> (bit & 7)) & 1) << i; bit++; }
      if (code === clear) { reset(); continue; }
      if (code === eoi) break;
      let entry: number[];
      if (code < dict.length) entry = dict[code];
      else entry = [...prev!, prev![0]];
      out.push(...entry);
      if (prev) dict.push([...prev, entry[0]]);
      prev = entry;
      if (dict.length === 1 << size2 && size2 < 12) size2++;
    }
    frames.push({ x, y, w, h, delay, pixels: out, table });
  }
  return { frames, loop };
}

describe("GIF writer", () => {
  it("LZW round-trips long runs and mixed data", () => {
    const data = Uint8Array.from({ length: 20000 }, (_, i) => (i % 97 < 50 ? 3 : (i * 7) % 5));
    const gw = new GifWriter(200, 100, false);
    const rgba = new Uint8ClampedArray(200 * 100 * 4);
    data.forEach((v, i) => { rgba[i * 4] = v * 50; rgba[i * 4 + 1] = 0; rgba[i * 4 + 2] = 0; rgba[i * 4 + 3] = 255; });
    gw.add({ rgba, delay: 10 });
    const g = decodeGif(gw.finish());
    expect(g.frames.length).toBe(1);
    const f = g.frames[0];
    expect(f.pixels.length).toBe(20000);
    const red = f.pixels.map((p) => f.table[p * 3]);
    expect(red).toEqual(Array.from(data, (v) => v * 50));
    expect(lzw(Uint8Array.of(), 2).length).toBeGreaterThan(0);
  });

  it("writes only the changed rectangle and merges equal frames into a longer delay", () => {
    const W = 40, H = 30;
    const mk = (fill: (x: number, y: number) => [number, number, number]) => {
      const a = new Uint8ClampedArray(W * H * 4);
      for (let y = 0; y < H; y++) for (let x = 0; x < W; x++) { const [r, g, b] = fill(x, y); const i = (y * W + x) * 4; a[i] = r; a[i + 1] = g; a[i + 2] = b; a[i + 3] = 255; }
      return a;
    };
    const white = mk(() => [255, 255, 255]);
    const box = mk((x, y) => (x >= 10 && x < 20 && y >= 5 && y < 9 ? [255, 0, 0] : [255, 255, 255]));
    const gw = new GifWriter(W, H, true);
    gw.add({ rgba: white, delay: 50 });
    gw.add({ rgba: white, delay: 50 });
    gw.add({ rgba: box, delay: 20 });
    const g = decodeGif(gw.finish());
    expect(g.loop).toBe(true);
    expect(g.frames.length).toBe(2);
    expect(g.frames[0].delay).toBe(100);
    expect([g.frames[1].x, g.frames[1].y, g.frames[1].w, g.frames[1].h]).toEqual([10, 5, 10, 4]);
  });

  it("median cut keeps few colours exactly and reduces many to the limit", () => {
    expect(medianCut([0xff0000, 0x00ff00, 0xff0000], 255).sort()).toEqual([0x00ff00, 0xff0000].sort());
    const many = Array.from({ length: 5000 }, (_, i) => (i * 2654435761) >>> 8 & 0xffffff);
    expect(medianCut(many, 255).length).toBe(255);
  });
});
