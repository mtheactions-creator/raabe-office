// Our own MP4 writer (D-073, D-074): the ISO base media file format (ISO/IEC 14496-12) with one H.264 video
// track (avc1 + avcC, 14496-15) and optionally one AAC audio track (mp4a + esds, 14496-14). The samples come
// from the WebView's encoders (WebCodecs: VideoEncoder with avc format "avc", AudioEncoder "mp4a.40.2"), which
// also give the decoder configurations (avcC record, AudioSpecificConfig). Layout: ftyp, mdat, moov; sample
// tables with one chunk per sample (stco offsets into mdat), 64-bit offsets when the file passes 4 GB.

export interface Mp4Sample { data: Uint8Array; /** Microseconds. */ time: number; duration: number; key: boolean }

interface Track {
  kind: "video" | "audio";
  timescale: number;
  samples: Mp4Sample[];
  /** Video: width, height and the avcC record; audio: rate, channels and the AudioSpecificConfig. */
  width?: number; height?: number; config: Uint8Array; rate?: number; channels?: number;
}

const enc = new TextEncoder();

/** A box: size, type and the payload parts. */
function box(type: string, ...parts: Uint8Array[]): Uint8Array {
  const len = 8 + parts.reduce((n, p) => n + p.length, 0);
  const out = new Uint8Array(len);
  const v = new DataView(out.buffer);
  v.setUint32(0, len);
  out.set(enc.encode(type), 4);
  let o = 8;
  for (const p of parts) { out.set(p, o); o += p.length; }
  return out;
}

/** A full box: version and flags before the payload. */
const full = (type: string, version: number, flags: number, ...parts: Uint8Array[]) => box(type, u8(version), u24(flags), ...parts);

const u8 = (n: number) => Uint8Array.of(n & 255);
const u16 = (n: number) => Uint8Array.of((n >> 8) & 255, n & 255);
const u24 = (n: number) => Uint8Array.of((n >> 16) & 255, (n >> 8) & 255, n & 255);
const u32 = (n: number) => { const b = new Uint8Array(4); new DataView(b.buffer).setUint32(0, n >>> 0); return b; };
const u64 = (n: number) => { const b = new Uint8Array(8); const v = new DataView(b.buffer); v.setUint32(0, Math.floor(n / 2 ** 32)); v.setUint32(4, n >>> 0); return b; };
const zeros = (n: number) => new Uint8Array(n);
const concat = (parts: Uint8Array[]) => { const out = new Uint8Array(parts.reduce((n, p) => n + p.length, 0)); let o = 0; for (const p of parts) { out.set(p, o); o += p.length; } return out; };

/** The unity matrix of tkhd and mvhd (16.16 and 2.30 fixed point). */
const MATRIX = concat([u32(0x00010000), u32(0), u32(0), u32(0), u32(0x00010000), u32(0), u32(0), u32(0), u32(0x40000000)]);

export class Mp4Writer {
  private tracks: Track[] = [];

  /** Adds the video track; returns its number for addSample. */
  addVideo(width: number, height: number, avcC: Uint8Array, timescale = 90000): number {
    this.tracks.push({ kind: "video", timescale, samples: [], width, height, config: avcC });
    return this.tracks.length - 1;
  }

  addAudio(rate: number, channels: number, audioSpecificConfig: Uint8Array): number {
    this.tracks.push({ kind: "audio", timescale: rate, samples: [], rate, channels, config: audioSpecificConfig });
    return this.tracks.length - 1;
  }

  addSample(track: number, s: Mp4Sample): void { this.tracks[track].samples.push(s); }

  /** The whole file. */
  finish(): Uint8Array {
    const ftyp = box("ftyp", enc.encode("isom"), u32(512), enc.encode("isom"), enc.encode("iso2"), enc.encode("avc1"), enc.encode("mp41"));
    const payload = this.tracks.flatMap((t) => t.samples.map((s) => s.data));
    const mdatSize = 8 + payload.reduce((n, p) => n + p.length, 0);
    const large = ftyp.length + mdatSize + 16 > 0xffffffff;
    // mdat: with a 64-bit size when large (size field 1, then the size).
    const mdatHead = large ? concat([u32(1), enc.encode("mdat"), u64(mdatSize + 8)]) : concat([u32(mdatSize), enc.encode("mdat")]);
    let offset = ftyp.length + mdatHead.length;
    const offsets: number[][] = this.tracks.map((t) => t.samples.map((s) => { const o = offset; offset += s.data.length; return o; }));
    const moov = this.moov(offsets, large);
    return concat([ftyp, mdatHead, ...payload, moov]);
  }

  private moov(offsets: number[][], large: boolean): Uint8Array {
    const MOVIE_SCALE = 1000;
    const durUs = Math.max(0, ...this.tracks.map((t) => trackDurationUs(t)));
    const mvhd = full("mvhd", 0, 0, u32(0), u32(0), u32(MOVIE_SCALE), u32(Math.round((durUs / 1e6) * MOVIE_SCALE)), u32(0x00010000), u16(0x0100), zeros(10), MATRIX, zeros(24), u32(this.tracks.length + 1));
    const traks = this.tracks.map((t, i) => this.trak(t, i + 1, offsets[i], large, MOVIE_SCALE));
    return box("moov", mvhd, ...traks);
  }

  private trak(t: Track, id: number, offsets: number[], large: boolean, movieScale: number): Uint8Array {
    const durUs = trackDurationUs(t);
    const video = t.kind === "video";
    const tkhd = full("tkhd", 0, 3, u32(0), u32(0), u32(id), u32(0), u32(Math.round((durUs / 1e6) * movieScale)), zeros(8), u16(0), u16(video ? 0 : 1), u16(video ? 0 : 0x0100), u16(0), MATRIX,
      u32(video ? (t.width! << 16) >>> 0 : 0), u32(video ? (t.height! << 16) >>> 0 : 0));
    const mdhd = full("mdhd", 0, 0, u32(0), u32(0), u32(t.timescale), u32(Math.round((durUs / 1e6) * t.timescale)), u16(0x55c4) /* und */, u16(0));
    const hdlr = full("hdlr", 0, 0, u32(0), enc.encode(video ? "vide" : "soun"), zeros(12), enc.encode(video ? "VideoHandler\0" : "SoundHandler\0"));
    const head = video ? full("vmhd", 0, 1, zeros(8)) : full("smhd", 0, 0, zeros(4));
    const dinf = box("dinf", full("dref", 0, 0, u32(1), full("url ", 0, 1)));
    const stbl = box("stbl", this.stsd(t), stts(t), ...(video ? [stss(t)] : []), stsc(t), stsz(t), stco(offsets, large));
    const minf = box("minf", head, dinf, stbl);
    return box("trak", tkhd, box("mdia", mdhd, hdlr, minf));
  }

  private stsd(t: Track): Uint8Array {
    if (t.kind === "video") {
      const avc1 = box("avc1", zeros(6), u16(1), zeros(16), u16(t.width!), u16(t.height!), u32(0x00480000), u32(0x00480000), u32(0), u16(1), zeros(32), u16(0x0018), u16(0xffff), box("avcC", t.config));
      return full("stsd", 0, 0, u32(1), avc1);
    }
    const asc = t.config;
    // esds: ES_Descriptor > DecoderConfigDescriptor (AAC, audio stream) > DecoderSpecificInfo, then SLConfig.
    const desc = (tag: number, body: Uint8Array) => concat([u8(tag), u8(body.length), body]);
    const dsi = desc(0x05, asc);
    const dcd = desc(0x04, concat([u8(0x40), u8(0x15), u24(0), u32(0), u32(0), dsi]));
    const es = desc(0x03, concat([u16(1), u8(0), dcd, desc(0x06, u8(2))]));
    const mp4a = box("mp4a", zeros(6), u16(1), zeros(8), u16(t.channels!), u16(16), u16(0), u16(0), u32((t.rate! << 16) >>> 0), full("esds", 0, 0, es));
    return full("stsd", 0, 0, u32(1), mp4a);
  }
}

const ticks = (t: Track, us: number) => Math.round((us / 1e6) * t.timescale);

function trackDurationUs(t: Track): number {
  const last = t.samples[t.samples.length - 1];
  return last ? last.time + last.duration - (t.samples[0]?.time ?? 0) : 0;
}

/** Decoding times: run-length of sample durations (from the samples' own times, so rounding never drifts). */
function stts(t: Track): Uint8Array {
  const runs: [number, number][] = [];
  const t0 = t.samples[0]?.time ?? 0;
  t.samples.forEach((s, i) => {
    const start = ticks(t, s.time - t0);
    const end = i + 1 < t.samples.length ? ticks(t, t.samples[i + 1].time - t0) : start + Math.max(1, ticks(t, s.duration));
    const d = Math.max(1, end - start);
    const last = runs[runs.length - 1];
    if (last && last[1] === d) last[0]++; else runs.push([1, d]);
  });
  return full("stts", 0, 0, u32(runs.length), ...runs.flatMap(([n, d]) => [u32(n), u32(d)]));
}

function stss(t: Track): Uint8Array {
  const keys = t.samples.map((s, i) => (s.key ? i + 1 : 0)).filter((n) => n > 0);
  return full("stss", 0, 0, u32(keys.length), ...keys.map(u32));
}

/** One sample per chunk. */
const stsc = (_t: Track) => full("stsc", 0, 0, u32(1), u32(1), u32(1), u32(1));

function stsz(t: Track): Uint8Array {
  return full("stsz", 0, 0, u32(0), u32(t.samples.length), ...t.samples.map((s) => u32(s.data.length)));
}

function stco(offsets: number[], large: boolean): Uint8Array {
  return large ? full("co64", 0, 0, u32(offsets.length), ...offsets.map(u64)) : full("stco", 0, 0, u32(offsets.length), ...offsets.map(u32));
}

/** Reads the boxes of a file (tests and our own check): type, offset and size, nested for container boxes. */
export interface Mp4Box { type: string; start: number; size: number; children?: Mp4Box[] }
const CONTAINERS = new Set(["moov", "trak", "mdia", "minf", "stbl", "dinf"]);
export function readBoxes(b: Uint8Array, start = 0, end = b.length): Mp4Box[] {
  const v = new DataView(b.buffer, b.byteOffset, b.byteLength);
  const out: Mp4Box[] = [];
  let o = start;
  while (o + 8 <= end) {
    let size = v.getUint32(o);
    const type = String.fromCharCode(b[o + 4], b[o + 5], b[o + 6], b[o + 7]);
    let head = 8;
    if (size === 1) { size = v.getUint32(o + 8) * 2 ** 32 + v.getUint32(o + 12); head = 16; }
    if (size < head) break;
    const bx: Mp4Box = { type, start: o, size };
    if (CONTAINERS.has(type)) bx.children = readBoxes(b, o + head, o + size);
    out.push(bx);
    o += size;
  }
  return out;
}
