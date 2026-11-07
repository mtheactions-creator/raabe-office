// Aufzeichnen (D-078): Von Beginn an / Ab aktueller Folie record the presenter while the show runs: the microphone
// (and the camera, when switched on) per slide and the time spent on each slide. At the end each recorded slide
// gets its narration (a WAV, or with the camera an MP4 written by our own writer with the WebView's encoders) as a
// media object that plays when the slide starts, as PowerPoint stores narrations, and the slide's time as its
// advance time (Nach), with Anzeigedauern verwenden on. Audio aufnehmen and Bildschirmaufzeichnung use the same
// pieces. Everything stays on the device.

import { h, showDialog } from "@raabe/shell";
import { PptxDom } from "@raabe/ooxml";
import type { PresentApp } from "./app";
import * as S from "./shapeops";
import { Mp4Writer } from "./mp4";

const RT = {
  media: "http://schemas.microsoft.com/office/2007/relationships/media",
  audio: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/audio",
  video: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/video",
  image: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
};

/** A recorded piece: PCM samples (mono, at `rate`) and, with the camera or screen, an MP4. */
export interface Take { slide: number; ms: number; pcm?: Float32Array; rate?: number; mp4?: Uint8Array }

/** WAV (16-bit mono) from PCM. */
export function wav(pcm: Float32Array, rate: number): Uint8Array {
  const out = new Uint8Array(44 + pcm.length * 2);
  const v = new DataView(out.buffer);
  const s = (o: number, t: string) => { for (let i = 0; i < t.length; i++) out[o + i] = t.charCodeAt(i); };
  s(0, "RIFF"); v.setUint32(4, 36 + pcm.length * 2, true); s(8, "WAVE"); s(12, "fmt ");
  v.setUint32(16, 16, true); v.setUint16(20, 1, true); v.setUint16(22, 1, true); v.setUint32(24, rate, true); v.setUint32(28, rate * 2, true); v.setUint16(32, 2, true); v.setUint16(34, 16, true);
  s(36, "data"); v.setUint32(40, pcm.length * 2, true);
  for (let i = 0; i < pcm.length; i++) v.setInt16(44 + i * 2, Math.max(-1, Math.min(1, pcm[i])) * 32767, true);
  return out;
}

/** Records a stream's sound (and its video track into an MP4) between start and stop. */
export class StreamRecorder {
  private rec: MediaRecorder | null = null;
  private chunks: Blob[] = [];
  private video: { enc: VideoEncoder; mp4: Mp4Writer; track: number; reader: ReadableStreamDefaultReader<VideoFrame>; t0: number; w: number; h: number; stop: boolean; done: Promise<void> } | null = null;

  constructor(private stream: MediaStream) {}

  start(): void {
    this.chunks = [];
    const audio = new MediaStream(this.stream.getAudioTracks());
    if (audio.getAudioTracks().length) {
      this.rec = new MediaRecorder(audio);
      this.rec.ondataavailable = (e) => { if (e.data.size) this.chunks.push(e.data); };
      this.rec.start();
    }
    const vt = this.stream.getVideoTracks()[0];
    const Processor = (globalThis as unknown as { MediaStreamTrackProcessor?: new (o: { track: MediaStreamTrack }) => { readable: ReadableStream<VideoFrame> } }).MediaStreamTrackProcessor;
    if (vt && Processor && typeof VideoEncoder !== "undefined") {
      const st = vt.getSettings();
      const w = Math.floor((st.width ?? 1280) / 2) * 2, hh = Math.floor((st.height ?? 720) / 2) * 2;
      const mp4 = new Mp4Writer();
      const v = { mp4, track: -1, t0: -1, w, h: hh, stop: false } as NonNullable<StreamRecorder["video"]>;
      v.enc = new VideoEncoder({
        output: (chunk, meta) => {
          if (v.track < 0) { const d = meta?.decoderConfig?.description; v.track = mp4.addVideo(w, hh, d ? new Uint8Array(d instanceof ArrayBuffer ? d : (d as ArrayBufferView).buffer) : new Uint8Array()); }
          const data = new Uint8Array(chunk.byteLength); chunk.copyTo(data);
          mp4.addSample(v.track, { data, time: chunk.timestamp, duration: chunk.duration ?? 33333, key: chunk.type === "key" });
        },
        error: () => { v.stop = true; },
      });
      v.enc.configure({ codec: "avc1.64001F", width: w, height: hh, bitrate: 2_500_000, framerate: 30, avc: { format: "avc" }, latencyMode: "realtime" });
      v.reader = new Processor({ track: vt }).readable.getReader();
      let n = 0;
      v.done = (async () => {
        for (;;) {
          const { value: f, done } = await v.reader.read();
          if (done || !f) break;
          if (v.stop) { f.close(); break; }
          if (v.t0 < 0) v.t0 = f.timestamp;
          const frame = new VideoFrame(f, { timestamp: f.timestamp - v.t0 });
          f.close();
          if (v.enc.encodeQueueSize < 8) v.enc.encode(frame, { keyFrame: n++ % 60 === 0 });
          frame.close();
        }
      })();
      this.video = v;
    }
  }

  pause(): void { if (this.rec?.state === "recording") this.rec.pause(); }
  resume(): void { if (this.rec?.state === "paused") this.rec.resume(); }

  /** Ends the piece: its sound as PCM and, with video, the MP4 (sound in it as AAC). */
  async stop(): Promise<{ pcm?: Float32Array; rate?: number; mp4?: Uint8Array }> {
    const out: { pcm?: Float32Array; rate?: number; mp4?: Uint8Array } = {};
    if (this.rec && this.rec.state !== "inactive") {
      const ended = new Promise<void>((r) => { this.rec!.onstop = () => r(); });
      this.rec.stop();
      await ended;
      try {
        const buf = await new AudioContext().decodeAudioData(await new Blob(this.chunks).arrayBuffer());
        out.pcm = buf.getChannelData(0).slice();
        out.rate = buf.sampleRate;
      } catch { /* nothing recorded */ }
    }
    if (this.video) {
      const v = this.video;
      v.stop = true;
      void v.reader.cancel().catch(() => {});
      await v.done.catch(() => {});
      await v.enc.flush().catch(() => {});
      v.enc.close();
      if (out.pcm && out.rate) await addSound(v.mp4, out.pcm, out.rate);
      out.mp4 = v.track >= 0 ? v.mp4.finish() : undefined;
      this.video = null;
    }
    return out;
  }
}

/** The recorded sound into the MP4 as AAC. */
async function addSound(mp4: Mp4Writer, pcm: Float32Array, rate: number): Promise<void> {
  if (typeof AudioEncoder === "undefined") return;
  const config: AudioEncoderConfig = { codec: "mp4a.40.2", sampleRate: rate, numberOfChannels: 1, bitrate: 96_000 };
  try { if (!(await AudioEncoder.isConfigSupported(config)).supported) return; } catch { return; }
  let track = -1;
  const enc = new AudioEncoder({
    output: (chunk, meta) => {
      if (track < 0) { const d = meta?.decoderConfig?.description; track = mp4.addAudio(rate, 1, d ? new Uint8Array(d instanceof ArrayBuffer ? d : (d as ArrayBufferView).buffer) : Uint8Array.of(0x11, 0x88)); }
      const data = new Uint8Array(chunk.byteLength); chunk.copyTo(data);
      mp4.addSample(track, { data, time: chunk.timestamp, duration: chunk.duration ?? Math.round((1024 * 1e6) / rate), key: true });
    },
    error: () => {},
  });
  enc.configure(config);
  for (let i = 0; i < pcm.length; i += 4800) {
    const part = pcm.slice(i, i + 4800);
    const ad = new AudioData({ format: "f32-planar", sampleRate: rate, numberOfFrames: part.length, numberOfChannels: 1, timestamp: Math.round((i * 1e6) / rate), data: part });
    enc.encode(ad);
    ad.close();
  }
  await enc.flush();
  enc.close();
}

/** A media object on a slide (narration: small, bottom right, hidden in the show unless it is a video), playing when the slide starts. */
export function addMediaToSlide(app: PresentApp, slidePath: string, bytes: Uint8Array, kind: "audio" | "video", o: { narration: boolean; poster: Uint8Array; box?: { x: number; y: number; w: number; h: number }; name: string }): number {
  const pptx = app.deck.pptx;
  const ext = kind === "audio" ? "wav" : "mp4";
  const media = pptx.freeName(`ppt/media/media#.${ext}`);
  const image = pptx.freeName("ppt/media/image#.png");
  pptx.put(media, bytes);
  pptx.setType(media, kind === "audio" ? "audio/wav" : "video/mp4", true);
  pptx.put(image, o.poster);
  pptx.setType(image, "image/png", true);
  const rels = pptx.rels(slidePath);
  const rMedia = rels.ensure(RT.media, `../media/${media.split("/").pop()}`);
  const rLink = rels.ensure(kind === "video" ? RT.video : RT.audio, `../media/${media.split("/").pop()}`);
  const rImg = pptx.link(slidePath, RT.image, image);
  const doc = pptx.doc(slidePath);
  const tree = S.treeOf(doc);
  const id = S.maxId(tree) + 1;
  const size = pptx.slideSize;
  const box = o.box ?? (kind === "audio" ? { x: size.cx - 609600, y: size.cy - 609600, w: 487680, h: 487680 } : { x: size.cx - 3048000, y: size.cy - 1828800, w: 2926080, h: 1645920 });
  const file = kind === "video" ? "videoFile" : "audioFile";
  tree.append(PptxDom.fragment(doc, `<p:pic xmlns:p="${PptxDom.NS_P}" xmlns:a="${PptxDom.NS_A}" xmlns:r="${PptxDom.NS_R}"><p:nvPicPr><p:cNvPr id="${id}" name="${PptxDom.esc(o.name)}"><a:hlinkClick r:id="" action="ppaction://media"/></p:cNvPr><p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr>` +
    `<p:nvPr><a:${file} r:link="${rLink}"/><p:extLst><p:ext uri="{DAA4B4D4-6D71-4841-9C94-3DE7FCFB9230}"><p14:media xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" r:embed="${rMedia}"/></p:ext></p:extLst></p:nvPr></p:nvPicPr>` +
    `<p:blipFill><a:blip r:embed="${rImg}"/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr><a:xfrm><a:off x="${box.x}" y="${box.y}"/><a:ext cx="${box.w}" cy="${box.h}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>`)[0]);
  autoplay(doc.documentElement, id, kind, o.narration);
  return id;
}

/**
 * Plays a media object when its slide starts (PowerPoint's timing: a playFrom command after the previous at the
 * main sequence's start, and the media node; a narration's audio node carries isNarration and is hidden while
 * not playing).
 */
function autoplay(root: Element, spid: number, kind: "audio" | "video", narration: boolean): void {
  const P = PptxDom.NS_P;
  const doc = root.ownerDocument;
  let timing = Array.from(root.children).find((c) => c.localName === "timing");
  if (!timing) {
    timing = PptxDom.fragment(doc, `<p:timing xmlns:p="${P}"><p:tnLst><p:par><p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot"><p:childTnLst/></p:cTn></p:par></p:tnLst></p:timing>`)[0];
    const ext = Array.from(root.children).find((c) => c.localName === "extLst");
    root.insertBefore(timing, ext ?? null);
  }
  const ids = Array.from(timing.getElementsByTagNameNS(P, "cTn")).map((c) => Number(c.getAttribute("id")) || 0);
  let next = Math.max(1, ...ids) + 1;
  const n = () => next++;
  const rootCtn = Array.from(timing.getElementsByTagNameNS(P, "cTn")).find((c) => c.getAttribute("nodeType") === "tmRoot")!;
  let list = Array.from(rootCtn.children).find((c) => c.localName === "childTnLst");
  if (!list) { list = doc.createElementNS(P, "p:childTnLst"); rootCtn.append(list); }
  let main = Array.from(timing.getElementsByTagNameNS(P, "cTn")).find((c) => c.getAttribute("nodeType") === "mainSeq");
  if (!main) {
    const seq = PptxDom.fragment(doc, `<p:seq xmlns:p="${P}" concurrent="1" nextAc="seek"><p:cTn id="${n()}" dur="indefinite" nodeType="mainSeq"><p:childTnLst/></p:cTn><p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst><p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:nextCondLst></p:seq>`)[0];
    list.prepend(seq);
    main = seq.firstElementChild!;
  }
  const mainList = Array.from(main.children).find((c) => c.localName === "childTnLst")!;
  // The play command at the start of the main sequence (with the slide, not on a click).
  mainList.prepend(PptxDom.fragment(doc, `<p:par xmlns:p="${P}"><p:cTn id="${n()}" fill="hold"><p:stCondLst><p:cond delay="indefinite"/><p:cond evt="onBegin" delay="0"><p:tn val="2"/></p:cond></p:stCondLst><p:childTnLst>` +
    `<p:par><p:cTn id="${n()}" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst><p:par><p:cTn id="${n()}" presetID="1" presetClass="mediacall" presetSubtype="0" fill="hold" nodeType="afterEffect"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>` +
    `<p:cmd type="call" cmd="playFrom(0.0)"><p:cBhvr><p:cTn id="${n()}" dur="1" fill="hold"/><p:tgtEl><p:spTgt spid="${spid}"/></p:tgtEl></p:cBhvr></p:cmd></p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par>`)[0]);
  list.append(PptxDom.fragment(doc, `<p:${kind} xmlns:p="${P}"${narration && kind === "audio" ? ' isNarration="1"' : ""}><p:cMediaNode vol="80000"${narration && kind === "audio" ? ' showWhenStopped="0"' : ""}><p:cTn id="${n()}" fill="hold" display="0"><p:stCondLst><p:cond delay="indefinite"/></p:stCondLst>` +
    `<p:endCondLst><p:cond evt="onStopAudio" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:endCondLst></p:cTn><p:tgtEl><p:spTgt spid="${spid}"/></p:tgtEl></p:cMediaNode></p:${kind}>`)[0]);
}

/** A small poster for a narration: a microphone-like circle (our own drawing), or a camera frame. */
export async function posterPng(kind: "audio" | "video", frame?: CanvasImageSource): Promise<Uint8Array> {
  const c = document.createElement("canvas");
  c.width = kind === "audio" ? 64 : 320; c.height = kind === "audio" ? 64 : 180;
  const g = c.getContext("2d")!;
  if (kind === "video" && frame) g.drawImage(frame, 0, 0, c.width, c.height);
  else { g.fillStyle = "#B7472A"; g.beginPath(); g.arc(32, 32, 28, 0, Math.PI * 2); g.fill(); g.fillStyle = "#fff"; g.fillRect(27, 14, 10, 24); g.fillRect(22, 42, 20, 4); g.fillRect(30, 38, 4, 6); }
  const blob = await new Promise<Blob | null>((r) => c.toBlob(r, "image/png"));
  return blob ? new Uint8Array(await blob.arrayBuffer()) : new Uint8Array();
}

/** Writes recorded takes into their slides: media, the slide's time (Nach) and Anzeigedauern verwenden. */
export async function saveTakes(app: PresentApp, takes: Take[]): Promise<void> {
  const deck = app.deck;
  const byslide = new Map<number, Take[]>();
  for (const t of takes) byslide.set(t.slide, [...(byslide.get(t.slide) ?? []), t]);
  const paths = [...byslide.keys()].map((i) => deck.slides[i]?.path).filter((p): p is string => !!p);
  const audioPoster = await posterPng("audio");
  deck.change("Aufzeichnung", paths, () => {
    for (const [i, list] of byslide) {
      const path = deck.slides[i]?.path;
      if (!path) continue;
      const ms = list.reduce((s, t) => s + t.ms, 0);
      // Nach: the time the slide was shown while recording.
      const root = deck.pptx.doc(path).documentElement;
      let tr = Array.from(root.getElementsByTagNameNS(PptxDom.NS_P, "transition"))[0];
      if (!tr) { tr = deck.pptx.doc(path).createElementNS(PptxDom.NS_P, "p:transition"); root.insertBefore(tr, Array.from(root.children).find((c) => c.localName === "timing" || c.localName === "extLst") ?? null); }
      for (const t of Array.from(root.getElementsByTagNameNS(PptxDom.NS_P, "transition"))) t.setAttribute("advTm", String(Math.round(ms)));
      const t = list[list.length - 1];
      if (t.mp4) addMediaToSlide(app, path, t.mp4, "video", { narration: true, poster: audioPoster, name: "Aufzeichnung" });
      else if (t.pcm && t.rate) addMediaToSlide(app, path, wav(joinPcm(list), t.rate), "audio", { narration: true, poster: audioPoster, name: "Aufzeichnung" });
    }
  });
  const { setShowFlag } = await import("./showsetup");
  setShowFlag(deck, "timings", true);
  deck.invalidate?.(paths[0] ?? "");
}

function joinPcm(list: Take[]): Float32Array {
  const n = list.reduce((s, t) => s + (t.pcm?.length ?? 0), 0);
  const out = new Float32Array(n);
  let o = 0;
  for (const t of list) if (t.pcm) { out.set(t.pcm, o); o += t.pcm.length; }
  return out;
}

/** Aufzeichnung löschen: narrations (and their timings) of this slide or of all slides. */
export function clearRecording(app: PresentApp, all: boolean, what: "both" | "timing" | "narration" = "both"): void {
  const deck = app.deck;
  const paths = all ? deck.slides.map((s) => s.path) : [app.slidePath()].filter((p): p is string => !!p);
  deck.change("Aufzeichnung löschen", paths, () => {
    for (const p of paths) {
      const root = deck.pptx.doc(p).documentElement;
      if (what !== "narration") for (const t of Array.from(root.getElementsByTagNameNS(PptxDom.NS_P, "transition"))) t.removeAttribute("advTm");
      if (what === "timing") continue;
      for (const a of Array.from(root.getElementsByTagNameNS(PptxDom.NS_P, "audio")).concat(Array.from(root.getElementsByTagNameNS(PptxDom.NS_P, "video")))) {
        if (a.getAttribute("isNarration") !== "1" && a.localName === "audio") continue;
        const spid = a.getElementsByTagNameNS(PptxDom.NS_P, "spTgt")[0]?.getAttribute("spid");
        const pic = Array.from(root.getElementsByTagNameNS(PptxDom.NS_P, "cNvPr")).find((c) => c.getAttribute("id") === spid && c.getAttribute("name") === "Aufzeichnung")?.closest("pic");
        if (!pic && a.localName === "video") continue;
        pic?.remove();
        for (const cmd of Array.from(root.getElementsByTagNameNS(PptxDom.NS_P, "spTgt")).filter((t) => t.getAttribute("spid") === spid)) {
          const par = cmd.closest("par");
          if (par && par.parentElement?.closest("cTn")?.getAttribute("nodeType") !== "tmRoot") par.remove();
        }
        a.remove();
      }
    }
  });
}

/** Asks for the microphone (and camera): the stream, or null with a message when it is refused. */
export async function getStream(camera: boolean, screen = false): Promise<MediaStream | null> {
  try {
    if (screen) {
      const s = await navigator.mediaDevices.getDisplayMedia({ video: true, audio: false });
      try { const mic = await navigator.mediaDevices.getUserMedia({ audio: true }); for (const t of mic.getAudioTracks()) s.addTrack(t); } catch { /* without sound */ }
      return s;
    }
    return await navigator.mediaDevices.getUserMedia({ audio: true, video: camera ? { width: 1280, height: 720 } : false });
  } catch {
    await showDialog({ title: "RaabePresent", body: h("div", null, camera || screen ? "Kamera, Bildschirm oder Mikrofon sind nicht verfügbar oder wurden nicht freigegeben." : "Das Mikrofon ist nicht verfügbar oder wurde nicht freigegeben."),
      buttons: [{ label: "OK", value: true, primary: true }], cancel: true });
    return null;
  }
}
