// Datei > Exportieren > Video erstellen and Animiertes GIF erstellen (D-074). The show is played on a timeline
// of its own: each shown slide comes in with its transition, its animations run (those after the previous at
// once, the click steps spread over the slide's time) and it stays for "In jeder Folie verbrachte Sekunden"
// (or its recorded time, Nach, when Aufgezeichnete Anzeigedauern are used). Each frame is our renderer's
// drawing of the slide at that moment (the show's own keyframes, held at the frame's time), drawn into a canvas;
// videos on a slide are drawn into their frame, sounds (transition sounds, audio and video sound) mixed into one
// track. The MP4 is H.264 + AAC from the WebView's encoders in our own file writer (mp4.ts); the GIF is our own
// writer (gif.ts).

import type { PresentApp } from "./app";
import { renderSlide } from "./render/slide";
import { EMU_PX } from "./render/text";
import { mainSequence, samplePath, type Step } from "./animations";
import { readTransition, transitionSound } from "./transitions";
import { effectKeyframes, transitionKeyframes } from "./show";
import { mediaOf, mediaType } from "./media";
import { Mp4Writer } from "./mp4";
import { GifWriter } from "./gif";

export interface MediaOptions {
  /** Height in pixels (the width follows the slide's proportions). */
  height: number;
  fps: number;
  /** In jeder Folie verbrachte Sekunden. */
  secs: number;
  /** Aufgezeichnete Anzeigedauern verwenden: a slide's Nach time replaces the seconds. */
  useTimings: boolean;
  /** Slides, 1-based, inclusive (GIF: Zu exportierende Folien). */
  from?: number;
  to?: number;
  /** GIF: Hintergrund transparent machen. */
  transparent?: boolean;
  /** GIF: the width rounded to the nearest pixel (PowerPoint's GIF 853 wide at 480, its video 852). */
  roundWidth?: boolean;
  /** GIF: PowerPoint's GIF timing (measured, D-107): transitions are cuts at their middle, and a slide lasts one
   *  frame longer, one more per animation step, and the last slide one more. */
  gifTiming?: boolean;
  /** Called with the share done (0..1); returning false stops the export. */
  progress?: (done: number) => boolean | void;
}

interface Act { at: number; step: Step }
interface MediaShape { id: number; path: string; kind: "video" | "audio"; x: number; y: number; w: number; h: number }
/** One slide on the timeline; times in milliseconds. */
export interface Segment { index: number; path: string; start: number; length: number; trans: { kind: string; dir?: string; dur: number; thruBlk?: boolean } | null; acts: Act[]; hidden: Set<number>; media: MediaShape[] }

const stepLength = (s: Step) => Math.max(0, ...s.effects.map((e) => e.delay + e.dur));

/**
 * The slides on the timeline, as PowerPoint's video (measured, probe_powerpoint_video.py): a slide takes its
 * transition plus its seconds (In jeder Folie verbrachte Sekunden, or Nach with recorded timings); the steps
 * after the previous start with it, the click steps come at equal shares of the seconds (two clicks in 2 s: at
 * 0.67 s and 1.33 s), and the animations' own durations add nothing unless they run past the slide's end.
 */
export function timeline(app: PresentApp, o: Pick<MediaOptions, "secs" | "useTimings" | "from" | "to"> & Partial<Pick<MediaOptions, "fps" | "gifTiming">>): Segment[] {
  const out: Segment[] = [];
  const slides = app.deck.slides;
  const first = Math.max(0, (o.from ?? 1) - 1), last = Math.min(slides.length - 1, (o.to ?? slides.length) - 1);
  let t = 0;
  for (let i = first; i <= last; i++) {
    const path = slides[i].path;
    const root = app.deck.pptx.doc(path).documentElement;
    if (root.getAttribute("show") === "0") continue;
    const tr = readTransition(root);
    const trans = tr && tr.kind !== "none" && tr.kind !== "cut" ? { kind: tr.kind, dir: tr.dir, dur: tr.dur, thruBlk: tr.thruBlk } : null;
    const td = trans?.dur ?? 0;
    const seq = mainSequence(root);
    const acts: Act[] = [];
    const timed = o.useTimings && tr?.advTm !== undefined;
    let hold = timed ? tr!.advTm! : o.secs * 1000;
    // PowerPoint's GIF: one frame more per slide and per step, one more on the last slide (measured, D-107).
    if (o.gifTiming && o.fps) hold += (1000 / o.fps) * (1 + seq.steps.length + (seq.auto ? 1 : 0) + (i === last ? 1 : 0));
    let end = td + hold;
    if (seq.auto) { acts.push({ at: td, step: seq.auto }); end = Math.max(end, td + stepLength(seq.auto)); }
    const gap = hold / (seq.steps.length + 1);
    seq.steps.forEach((s, k) => { const at = td + gap * (k + 1); acts.push({ at, step: s }); end = Math.max(end, at + stepLength(s)); });
    const length = Math.max(1, end);
    const media: MediaShape[] = [];
    for (const s of app.deck.view(path).shapes) {
      if (s.kind !== "pic") continue;
      const m = mediaOf(app, path, s.el);
      if (m) media.push({ id: s.id, path: m.path, kind: m.kind, x: s.x, y: s.y, w: s.w, h: s.h });
    }
    out.push({ index: i, path, start: t, length, trans, acts, hidden: seq.hidden, media });
    t += length;
  }
  return out;
}

export const totalLength = (segs: Segment[]) => (segs.length ? segs[segs.length - 1].start + segs[segs.length - 1].length : 0);

// ----- drawing a frame ------------------------------------------------------------------------------------

/** A slide drawn once, with its animations held (paused) so that any moment can be shown. */
class SlideStage {
  readonly el: HTMLElement;
  private shapes: Map<number, HTMLElement>;
  private anims: { el: HTMLElement; a: Animation; at: number; cls: string; delay: number; dur: number }[] = [];
  private enter: Animation | null = null;

  constructor(app: PresentApp, readonly seg: Segment, host: HTMLElement) {
    const r = renderSlide(app.deck.view(seg.path), app.deck.pptx.slideSize, { editing: false, slideNumber: seg.index + 1, image: (p) => app.image(p), pptx: app.deck.pptx });
    this.el = r.el;
    this.shapes = r.shapes;
    this.el.style.position = "absolute";
    this.el.style.left = "0"; this.el.style.top = "0";
    host.append(this.el);
    const size = app.deck.pptx.slideSize, W = size.cx / 9525, H = size.cy / 9525;
    for (const act of seg.acts) for (const fx of act.step.effects) {
      const el = this.shapes.get(fx.spid);
      if (!el || fx.cls === "mediacall") continue;
      const kf = fx.path ? samplePath(fx.path).map(([px, py]) => ({ translate: `${px * W}px ${py * H}px` })) : effectKeyframes(fx.cls, fx.effect);
      const a = el.animate(kf, { duration: Math.max(1, fx.dur), fill: "both", easing: fx.path ? "ease-in-out" : "ease-out" });
      a.pause();
      this.anims.push({ el, a, at: act.at, cls: fx.path ? "path" : fx.cls, delay: fx.delay, dur: Math.max(1, fx.dur) });
    }
    if (seg.trans) {
      this.enter = this.el.animate(transitionKeyframes(seg.trans.kind, seg.trans.dir, seg.trans.thruBlk), { duration: Math.max(1, seg.trans.dur), fill: "both", easing: "ease-in-out" });
      this.enter.pause();
    }
  }

  /**
   * The slide at slide time t: each effect before its start (an entrance hidden), running, or done (an exit
   * hidden). Returns a key of what is shown and whether anything moves at t.
   */
  seek(t: number): { key: string; moving: boolean } {
    let moving = false;
    const key: string[] = [];
    // Visibility in time order: shapes with an entrance first start hidden; an entrance shows its shape,
    // a finished exit hides it.
    const vis = new Map<HTMLElement, boolean>();
    for (const [id, el] of this.shapes) if (this.seg.hidden.has(id)) vis.set(el, false);
    for (const x of [...this.anims].sort((p, q) => p.at + p.delay - (q.at + q.delay))) {
      const local = t - x.at - x.delay;
      const phase = local < 0 ? 0 : local >= x.dur ? 2 : 1;
      if (phase === 1) moving = true;
      key.push(String(phase));
      x.a.currentTime = Math.min(Math.max(local, 0), x.dur);
      if (phase > 0 && x.cls === "entr") vis.set(x.el, true);
      if (x.cls === "exit" && phase > 0) vis.set(x.el, phase === 1);
    }
    for (const [el, v] of vis) el.style.visibility = v ? "visible" : "hidden";
    if (this.enter) {
      const d = Math.max(1, this.seg.trans!.dur);
      this.enter.currentTime = Math.min(Math.max(t, 0), d);
      if (t < d) moving = true;
      key.push(t < d ? "tr" : "in");
    }
    if (moving) key.push(String(Math.round(t)));
    return { key: key.join(""), moving };
  }

  /** Copies the animated values into the elements' own style (what the picture shows). */
  bake(): () => void {
    const els = new Set<HTMLElement>([this.el, ...this.anims.map((x) => x.el)]);
    const saved: [HTMLElement, string][] = [];
    for (const el of els) {
      const cs = getComputedStyle(el);
      saved.push([el, el.style.cssText]);
      el.style.opacity = cs.opacity;
      el.style.transform = cs.transform === "none" ? "" : cs.transform;
      el.style.clipPath = cs.clipPath === "none" ? "" : cs.clipPath;
      el.style.translate = cs.translate === "none" ? "" : cs.translate;
      el.style.filter = cs.filter === "none" ? "" : cs.filter;
    }
    return () => { for (const [el, css] of saved) el.style.cssText = css; };
  }

  remove(): void { for (const x of this.anims) x.a.cancel(); this.enter?.cancel(); this.el.remove(); }
}

/** Turns the slide DOM into a picture: the page's styles and the pictures as data inside an SVG. */
class Rasterizer {
  private css: string;
  private data = new Map<string, string>();
  constructor(private W: number, private H: number) {
    this.css = Array.from(document.styleSheets).flatMap((ss) => { try { return Array.from(ss.cssRules).map((x) => x.cssText); } catch { return []; } }).join("\n").replace(/</g, "\\3c ");
  }

  async draw(nodes: HTMLElement[], g: CanvasRenderingContext2D, width: number, height: number): Promise<boolean> {
    let html = nodes.map((n) => new XMLSerializer().serializeToString(n)).join("");
    for (const u of new Set(html.match(/blob:[^"')\s&]+/g) ?? [])) {
      let d = this.data.get(u);
      if (!d) { try { d = await toDataUrl(u); this.data.set(u, d); } catch { continue; } }
      html = html.split(u).join(d);
    }
    const scale = width / this.W;
    const svg = `<svg xmlns="http://www.w3.org/2000/svg" width="${width}" height="${height}"><foreignObject width="${this.W}" height="${this.H}" transform="scale(${scale})">` +
      `<div xmlns="http://www.w3.org/1999/xhtml" style="width:${this.W}px;height:${this.H}px;position:relative;overflow:hidden;background:#000"><style>${this.css}</style>${html}</div></foreignObject></svg>`;
    const img = new Image();
    const ok = await new Promise<boolean>((res) => { img.onload = () => res(true); img.onerror = () => res(false); img.src = "data:image/svg+xml;charset=utf-8," + encodeURIComponent(svg); });
    if (!ok) return false;
    g.drawImage(img, 0, 0, width, height);
    return true;
  }
}

async function toDataUrl(url: string): Promise<string> {
  const blob = await (await fetch(url)).blob();
  return await new Promise((res) => { const r = new FileReader(); r.onload = () => res(String(r.result)); r.readAsDataURL(blob); });
}

/** The videos of a slide, each seeked to the frame's time and drawn into its box. */
class Videos {
  private els = new Map<string, HTMLVideoElement>();
  constructor(private app: PresentApp) {}

  async draw(seg: Segment, t: number, g: CanvasRenderingContext2D, scale: number): Promise<boolean> {
    let playing = false;
    const td = seg.trans?.dur ?? 0;
    for (const m of seg.media) {
      if (m.kind !== "video") continue;
      let v = this.els.get(m.path);
      if (!v) {
        const bytes = this.app.deck.pptx.files[m.path];
        if (!bytes) continue;
        v = document.createElement("video");
        v.muted = true; v.preload = "auto";
        v.src = URL.createObjectURL(new Blob([bytes as BlobPart], { type: mediaType(m.path) }));
        await new Promise<void>((res) => { v!.onloadeddata = () => res(); v!.onerror = () => res(); });
        this.els.set(m.path, v);
      }
      if (!v.duration || !Number.isFinite(v.duration)) continue;
      const at = Math.max(0, (t - td) / 1000);
      const want = Math.min(at, Math.max(0, v.duration - 0.04));
      if (at < v.duration) playing = true;
      if (Math.abs(v.currentTime - want) > 0.001) await new Promise<void>((res) => { v!.onseeked = () => res(); v!.currentTime = want; setTimeout(res, 2000); });
      g.drawImage(v, (m.x / EMU_PX) * scale, (m.y / EMU_PX) * scale, (m.w / EMU_PX) * scale, (m.h / EMU_PX) * scale);
    }
    return playing;
  }

  dispose(): void { for (const v of this.els.values()) { URL.revokeObjectURL(v.src); v.remove(); } }
}

/**
 * Plays the timeline frame by frame into a canvas; `frame` gets each picture with its time and duration
 * (microseconds) and whether it equals the one before (nothing moved).
 */
async function playFrames(app: PresentApp, o: MediaOptions, width: number, height: number, frame: (c: HTMLCanvasElement, time: number, dur: number, same: boolean) => Promise<void>): Promise<boolean> {
  const segs = timeline(app, o);
  const total = totalLength(segs);
  const size = app.deck.pptx.slideSize, W = size.cx / EMU_PX, H = size.cy / EMU_PX;
  const host = document.createElement("div");
  host.style.cssText = `position:fixed;left:-${W * 3}px;top:0;width:${W}px;height:${H}px;overflow:hidden;pointer-events:none`;
  document.body.append(host);
  const canvas = document.createElement("canvas");
  canvas.width = width; canvas.height = height;
  const g = canvas.getContext("2d", { willReadFrequently: true })!;
  const ras = new Rasterizer(W, H);
  const videos = new Videos(app);
  const n = Math.max(1, Math.ceil((total / 1000) * o.fps));
  let cur: SlideStage | null = null, prev: SlideStage | null = null;
  let lastKey = "";
  try {
    for (let f = 0; f < n; f++) {
      const t = (f * 1000) / o.fps;
      const seg = segs.find((s) => t >= s.start && t < s.start + s.length) ?? segs[segs.length - 1];
      if (!cur || cur.seg !== seg) {
        prev?.remove();
        prev = cur;
        cur = new SlideStage(app, seg, host);
        if (prev) { prev.seek(prev.seg.length); }
      }
      const local = t - seg.start;
      const st = cur.seek(local);
      // During a transition the slide before stays underneath; PowerPoint's GIF shows no transition, only a cut
      // at its middle (the slide before until then).
      const cutBefore = !!(o.gifTiming && seg.trans && local < seg.trans.dur / 2 && prev);
      const under = seg.trans && local < seg.trans.dur && prev && !o.gifTiming ? [prev] : [];
      if (!under.length && !cutBefore && prev) { prev.remove(); prev = null; }
      const key = cutBefore ? `${prev!.seg.index}:end` : `${seg.index}:${o.gifTiming && seg.trans && local < seg.trans.dur ? "in" : st.key}:${under.length}`;
      const hasVideo = seg.media.some((m) => m.kind === "video");
      const same = key === lastKey && !hasVideo;
      if (!same) {
        g.fillStyle = "#000";
        g.fillRect(0, 0, width, height);
        const shown = cutBefore ? [prev!] : [...under, cur];
        if (o.gifTiming && !cutBefore && seg.trans && local < seg.trans.dur) cur.seek(seg.trans.dur);
        const restore = shown.map((s) => s.bake());
        const nodes = shown.map((s) => s.el);
        try { await ras.draw(nodes, g, width, height); } finally { for (const r of restore) r(); }
        if (hasVideo) await videos.draw(seg, local, g, width / W);
      }
      lastKey = key;
      await frame(canvas, Math.round((f * 1e6) / o.fps), Math.round(1e6 / o.fps), same);
      if (o.progress && o.progress((f + 1) / n) === false) return false;
    }
    return true;
  } finally {
    cur?.remove(); prev?.remove();
    videos.dispose();
    host.remove();
  }
}

/** Width for a height: the slide's proportions (H.264 needs even sizes). */
export function frameSize(app: PresentApp, height: number, round = false): { width: number; height: number } {
  const s = app.deck.pptx.slideSize;
  if (round) return { width: Math.round((height * s.cx) / s.cy), height };
  // PowerPoint's width: rounded down to a multiple of 4 (480p 16:9: 852).
  const w = Math.floor((height * s.cx) / s.cy / 4) * 4;
  return { width: w, height: height - (height % 2) };
}

// ----- sound ----------------------------------------------------------------------------------------------

/** The show's sounds mixed into one stereo track at 48 kHz, or null when there is none. */
async function mixSound(app: PresentApp, segs: Segment[]): Promise<AudioBuffer | null> {
  const total = totalLength(segs) / 1000;
  const events: { at: number; bytes: Uint8Array; loopUntil?: number }[] = [];
  for (const seg of segs) {
    const snd = transitionSound(app, seg.path);
    const sb = snd?.media ? app.deck.pptx.files[snd.media] : undefined;
    if (sb) events.push({ at: seg.start / 1000, bytes: sb, loopUntil: snd?.loop ? (seg.start + seg.length) / 1000 : undefined });
    for (const m of seg.media) {
      const b = app.deck.pptx.files[m.path];
      if (b) events.push({ at: (seg.start + (seg.trans?.dur ?? 0)) / 1000, bytes: b });
    }
  }
  if (!events.length || total <= 0) return null;
  const rate = 48000;
  const ctx = new OfflineAudioContext(2, Math.ceil(total * rate), rate);
  let any = false;
  const decoded = new Map<Uint8Array, AudioBuffer | null>();
  for (const e of events) {
    let buf = decoded.get(e.bytes);
    if (buf === undefined) {
      try { buf = await ctx.decodeAudioData(e.bytes.slice().buffer as ArrayBuffer); } catch { buf = null; }
      decoded.set(e.bytes, buf);
    }
    if (!buf) continue;
    const src = ctx.createBufferSource();
    src.buffer = buf;
    if (e.loopUntil !== undefined) { src.loop = true; src.start(e.at); src.stop(e.loopUntil); } else src.start(e.at);
    src.connect(ctx.destination);
    any = true;
  }
  return any ? await ctx.startRendering() : null;
}

// ----- MP4 ------------------------------------------------------------------------------------------------

/** H.264 profile and level for the picture size (High profile; level by the number of macroblocks). */
function avcCodec(width: number, height: number): string {
  const mbs = Math.ceil(width / 16) * Math.ceil(height / 16);
  const level = mbs <= 3600 ? "1F" : mbs <= 8192 ? "28" : mbs <= 8704 ? "2A" : "33";
  return `avc1.6400${level}`;
}

const BITRATES: [number, number][] = [[2160, 20_000_000], [1080, 8_000_000], [720, 5_000_000], [480, 2_500_000]];

export function videoSupported(): boolean { return typeof VideoEncoder !== "undefined" && typeof OfflineAudioContext !== "undefined"; }

/** The presentation as an MP4 (null when stopped or when the WebView cannot encode). */
export async function makeVideo(app: PresentApp, o: MediaOptions): Promise<Uint8Array | null> {
  if (!videoSupported()) return null;
  const { width, height } = frameSize(app, o.height);
  const bitrate = BITRATES.find(([h]) => o.height >= h)?.[1] ?? 2_500_000;
  const base: VideoEncoderConfig = { codec: avcCodec(width, height), width, height, bitrate, framerate: o.fps, avc: { format: "avc" }, latencyMode: "quality" };
  let config: VideoEncoderConfig | null = null;
  for (const c of [base, { ...base, hardwareAcceleration: "prefer-software" as const }, { ...base, codec: "avc1.4D0033" }, { ...base, codec: "avc1.42E033", latencyMode: "realtime" as const }]) {
    try { if ((await VideoEncoder.isConfigSupported(c)).supported) { config = c; break; } } catch { /* next */ }
  }
  if (!config) return null;
  const mp4 = new Mp4Writer();
  let vtrack = -1;
  let failure: unknown = null;
  const encoder = new VideoEncoder({
    output: (chunk, meta) => {
      if (vtrack < 0) {
        const desc = meta?.decoderConfig?.description;
        const avcC = desc ? new Uint8Array(desc instanceof ArrayBuffer ? desc : (desc as ArrayBufferView).buffer.slice((desc as ArrayBufferView).byteOffset, (desc as ArrayBufferView).byteOffset + (desc as ArrayBufferView).byteLength)) : new Uint8Array();
        vtrack = mp4.addVideo(width, height, avcC);
      }
      const data = new Uint8Array(chunk.byteLength);
      chunk.copyTo(data);
      mp4.addSample(vtrack, { data, time: chunk.timestamp, duration: chunk.duration ?? Math.round(1e6 / o.fps), key: chunk.type === "key" });
    },
    error: (e) => { failure = e; },
  });
  encoder.configure(config);
  const keyEvery = o.fps * 2;
  let f = 0;
  const done = await playFrames(app, o, width, height, async (canvas, time, dur) => {
    const vf = new VideoFrame(canvas, { timestamp: time, duration: dur });
    encoder.encode(vf, { keyFrame: f % keyEvery === 0 });
    vf.close();
    f++;
    while (encoder.encodeQueueSize > 4) await new Promise((r) => setTimeout(r, 5));
    if (failure) throw failure;
  });
  await encoder.flush();
  encoder.close();
  if (!done || failure) return null;
  // Sound.
  const sound = await mixSound(app, timeline(app, o));
  if (sound) await encodeSound(mp4, sound);
  return mp4.finish();
}

async function encodeSound(mp4: Mp4Writer, buf: AudioBuffer): Promise<void> {
  if (typeof AudioEncoder === "undefined") return;
  const config: AudioEncoderConfig = { codec: "mp4a.40.2", sampleRate: buf.sampleRate, numberOfChannels: 2, bitrate: 160_000 };
  try { if (!(await AudioEncoder.isConfigSupported(config)).supported) return; } catch { return; }
  let track = -1;
  const enc = new AudioEncoder({
    output: (chunk, meta) => {
      if (track < 0) {
        const d = meta?.decoderConfig?.description;
        const asc = d ? new Uint8Array(d instanceof ArrayBuffer ? d : (d as ArrayBufferView).buffer.slice((d as ArrayBufferView).byteOffset, (d as ArrayBufferView).byteOffset + (d as ArrayBufferView).byteLength)) : Uint8Array.of(0x11, 0x90);
        track = mp4.addAudio(buf.sampleRate, 2, asc);
      }
      const data = new Uint8Array(chunk.byteLength);
      chunk.copyTo(data);
      mp4.addSample(track, { data, time: chunk.timestamp, duration: chunk.duration ?? Math.round((1024 * 1e6) / buf.sampleRate), key: true });
    },
    error: () => {},
  });
  enc.configure(config);
  const n = 4800;
  const l = buf.getChannelData(0), r = buf.numberOfChannels > 1 ? buf.getChannelData(1) : l;
  for (let i = 0; i < buf.length; i += n) {
    const m = Math.min(n, buf.length - i);
    const planar = new Float32Array(m * 2);
    planar.set(l.subarray(i, i + m), 0);
    planar.set(r.subarray(i, i + m), m);
    const ad = new AudioData({ format: "f32-planar", sampleRate: buf.sampleRate, numberOfFrames: m, numberOfChannels: 2, timestamp: Math.round((i * 1e6) / buf.sampleRate), data: planar });
    enc.encode(ad);
    ad.close();
  }
  await enc.flush();
  enc.close();
}

// ----- GIF ------------------------------------------------------------------------------------------------

/** The presentation as an animated GIF (null when stopped). */
export async function makeGif(app: PresentApp, o: MediaOptions): Promise<Uint8Array | null> {
  const { width, height } = frameSize(app, o.height, o.roundWidth);
  // A frame is written when the next differs (equal frames lengthen its delay).
  const st: { writer: GifWriter | null; pending: { rgba: Uint8ClampedArray; delay: number } | null } = { writer: null, pending: null };
  const done = await playFrames(app, o, width, height, async (canvas, _time, dur, same) => {
    const delay = dur / 10000;
    if (same && st.pending) { st.pending.delay += delay; return; }
    const rgba = canvas.getContext("2d")!.getImageData(0, 0, width, height).data;
    // Hintergrund transparent machen: the colour at the first frame's corner is the slide's background.
    if (!st.writer) st.writer = new GifWriter(width, height, true, o.transparent ? [rgba[0], rgba[1], rgba[2]] : null);
    if (st.pending) st.writer.add({ rgba: st.pending.rgba, delay: Math.round(st.pending.delay) });
    st.pending = { rgba: new Uint8ClampedArray(rgba), delay };
  });
  if (!done || !st.writer) return null;
  if (st.pending) st.writer.add({ rgba: st.pending.rgba, delay: Math.round(st.pending.delay) });
  return st.writer.finish();
}
