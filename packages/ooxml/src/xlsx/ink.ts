// Ink (Zeichnen, D-093) as Excel 365 stores it, recorded with probe_excel_ink.py: each drawing is a
// twoCellAnchor (editAs="oneCell") holding mc:AlternateContent; its Choice (Requires="xdr14") has an
// xdr:contentPart whose relationship (type customXml) points to an InkML part (xl/ink/inkN.xml,
// application/inkml+xml) and xdr14 children (cNvPr "Freihand N", xfrm = the box of the strokes'
// centre lines); its Fallback is a picture of the ink. The InkML's points are in 1/1000 cm (X, Y)
// and ms (T); Excel maps the strokes' box onto the xfrm and writes the points as InkML differences
// (' first, " second differences).

import type { InkBrush, Sheet, SheetInk } from "@raabe/sheet-engine";
import { strFromU8 } from "fflate";
import { PackageWriter, Rels, resolveTarget } from "../package";
import { esc } from "./xml";

export const INK_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml";
export const INK_CONTENT_TYPE = "application/inkml+xml";
const IMAGE_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";
const EMU_PER_PX = 9525;
/** 1/1000 cm per px at 96 dpi: 9525 EMU per px, 360 EMU per 1/1000 cm. */
const HIM_PER_PX = EMU_PER_PX / 360;

// ----- InkML -------------------------------------------------------------------------------------

/** The values of a trace: InkML's explicit (!), first-difference (') and second-difference (") qualifiers. */
export function decodeTrace(text: string, channels: number): number[][] {
  const out: number[][] = [];
  const mode: string[] = new Array(channels).fill("!");
  const prev: number[] = new Array(channels).fill(0);
  const vel: number[] = new Array(channels).fill(0);
  for (const raw of text.split(",")) {
    const s = raw.trim();
    if (!s) continue;
    const tokens = s.match(/[!'"*?]|[-+]?(?:\d+\.?\d*|\.\d+)(?:[eE][-+]?\d+)?/g) ?? [];
    const point: number[] = [];
    let q: string | null = null;
    for (const t of tokens) {
      if (/^[!'"*?]$/.test(t)) { q = t; continue; }
      const c = point.length;
      if (c >= channels) break;
      if (q) { mode[c] = q === "*" || q === "?" ? mode[c] : q; }
      const v = Number(t);
      let val: number;
      if (mode[c] === "'") { vel[c] = v; val = prev[c] + v; }
      else if (mode[c] === '"') { vel[c] += v; val = prev[c] + vel[c]; }
      else { vel[c] = out.length ? v - prev[c] : 0; val = v; }
      point.push(val);
      prev[c] = val;
      q = null;
    }
    if (point.length) out.push(point);
  }
  return out;
}

/** Points written as Excel writes them: the first explicit, the second as first differences, the rest as second differences. */
export function encodeTrace(points: number[][]): string {
  const join = (vals: number[], q: string) => vals.map((v, i) => (q ? q : i && !(v < 0) ? " " : "") + String(v)).join("");
  const parts: string[] = [];
  let mode = "";
  for (let i = 0; i < points.length; i++) {
    const p = points[i];
    if (i === 0) { parts.push(join(p, "")); continue; }
    const d1 = p.map((v, c) => v - points[i - 1][c]);
    if (i === 1) { parts.push(join(d1, "'")); mode = "'"; continue; }
    const d0 = points[i - 1].map((v, c) => v - points[i - 2][c]);
    const d2 = d1.map((v, c) => v - d0[c]);
    parts.push(mode === '"' ? join(d2, "") : join(d2, '"'));
    mode = '"';
  }
  return parts.join(",");
}

interface Parsed { strokes: { pts: number[][]; brush: InkBrush }[]; time?: string; channels: string[] }

/** An InkML part: its traces (X, Y and T when present) with their brushes. */
export function parseInkml(xml: string): Parsed {
  const channels = [...(/<inkml:traceFormat>([\s\S]*?)<\/inkml:traceFormat>/.exec(xml)?.[1] ?? "").matchAll(/<inkml:channel\b[^>]*name="([^"]+)"/g)].map((m) => m[1]);
  if (!channels.length) channels.push("X", "Y");
  const brushes = new Map<string, InkBrush>();
  for (const m of xml.matchAll(/<inkml:brush\b[^>]*xml:id="([^"]+)"[^>]*>([\s\S]*?)<\/inkml:brush>/g)) {
    const props: Record<string, string> = {};
    for (const p of m[2].matchAll(/<inkml:brushProperty\b[^>]*name="([^"]+)"[^>]*value="([^"]*)"/g)) props[p[1]] = p[2];
    const kind = props.tip === "rectangle" || props.rasterOp === "maskPen" ? "highlighter" : props.inkEffects === "pencil" ? "pencil" : "pen";
    const extra: Record<string, string> = {};
    for (const [k, v] of Object.entries(props)) if (!["width", "height", "color", "tip", "rasterOp"].includes(k)) extra[k] = v;
    brushes.set(m[1], { color: (props.color ?? "#000000").toUpperCase(), width: Math.round(Number(props.width ?? 0.035) * 10 * 1e6) / 1e6, kind, ...(Object.keys(extra).length ? { extra } : {}) });
  }
  const strokes: Parsed["strokes"] = [];
  const ix = channels.indexOf("X"), iy = channels.indexOf("Y"), it = channels.indexOf("T");
  for (const m of xml.matchAll(/<inkml:trace\b([^>]*)>([\s\S]*?)<\/inkml:trace>/g)) {
    const ref = /brushRef="#?([^"]+)"/.exec(m[1])?.[1];
    const brush = (ref && brushes.get(ref)) || [...brushes.values()][0] || { color: "#000000", width: 0.35, kind: "pen" as const };
    const vals = decodeTrace(m[2], channels.length);
    strokes.push({ pts: vals.map((v) => [v[ix] ?? 0, v[iy] ?? 0, it >= 0 ? v[it] ?? 0 : 0]), brush });
  }
  return { strokes, time: /timeString="([^"]+)"/.exec(xml)?.[1], channels };
}

const CONTEXT = '<inkml:context xml:id="ctx0"><inkml:inkSource xml:id="inkSrc0"><inkml:traceFormat>'
  + '<inkml:channel name="X" type="integer" max="32767" units="cm"/><inkml:channel name="Y" type="integer" max="32767" units="cm"/>'
  + '<inkml:channel name="T" type="integer" max="2.14748E9" units="dev"/></inkml:traceFormat><inkml:channelProperties>'
  + '<inkml:channelProperty channel="X" name="resolution" value="1000" units="1/cm"/><inkml:channelProperty channel="Y" name="resolution" value="1000" units="1/cm"/>'
  + '<inkml:channelProperty channel="T" name="resolution" value="1" units="1/dev"/></inkml:channelProperties></inkml:inkSource>';

const cm = (mm: number) => String(Math.round(mm * 1000) / 10000);

function brushXml(id: string, b: InkBrush): string {
  let props = `<inkml:brushProperty name="width" value="${cm(b.width)}" units="cm"/><inkml:brushProperty name="height" value="${cm(b.width)}" units="cm"/>`;
  // Excel leaves the colour out for black (recorded).
  if (b.color.toUpperCase() !== "#000000") props += `<inkml:brushProperty name="color" value="${esc(b.color.toUpperCase())}"/>`;
  if (b.kind === "highlighter") props += '<inkml:brushProperty name="tip" value="rectangle"/><inkml:brushProperty name="rasterOp" value="maskPen"/>';
  for (const [k, v] of Object.entries(b.extra ?? {})) props += `<inkml:brushProperty name="${esc(k)}" value="${esc(v)}"/>`;
  return `<inkml:brush xml:id="${id}">${props}</inkml:brush>`;
}

/** The InkML part of a drawing: its points in 1/1000 cm of the box at 100 %, as Excel saves them. */
export function inkmlXml(ink: SheetInk): string {
  const brushes: InkBrush[] = [];
  const keyOf = (b: InkBrush) => JSON.stringify(b);
  for (const s of ink.strokes) if (!brushes.some((b) => keyOf(b) === keyOf(s.brush))) brushes.push(s.brush);
  const time = (ink.time ?? new Date().toISOString()).replace(/Z$/, "").replace(/\.\d+$/, "");
  const traces = ink.strokes.map((s) => {
    const pts = s.pts.map((p) => [Math.round(p.x * HIM_PER_PX), Math.round(p.y * HIM_PER_PX), Math.round(p.t ?? 0)]);
    return `<inkml:trace contextRef="#ctx0" brushRef="#br${brushes.findIndex((b) => keyOf(b) === keyOf(s.brush))}">${encodeTrace(pts)}</inkml:trace>`;
  }).join("");
  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
    + `<inkml:ink xmlns:inkml="http://www.w3.org/2003/InkML"><inkml:definitions>${CONTEXT}<inkml:timestamp xml:id="ts0" timeString="${time}"/></inkml:context>`
    + brushes.map((b, i) => brushXml(`br${i}`, b)).join("") + `</inkml:definitions>${traces}</inkml:ink>`;
}

// ----- drawing -----------------------------------------------------------------------------------

function point(ax: string, tag: string): { col: number; row: number; x: number; y: number } | undefined {
  const m = new RegExp(`<xdr:${tag}>([\\s\\S]*?)</xdr:${tag}>`).exec(ax);
  if (!m) return undefined;
  const v = (n: string) => Number(new RegExp(`<xdr:${n}>([^<]*)</xdr:${n}>`).exec(m[1])?.[1] ?? 0);
  return { col: v("col"), row: v("row"), x: v("colOff") / EMU_PER_PX, y: v("rowOff") / EMU_PER_PX };
}

/** What decides whether an ink drawing changed (its place, name and strokes). */
export function inkKey(k: SheetInk): string {
  return JSON.stringify([k.from, k.to, k.size, k.name ?? null, k.strokes]);
}

export const isInkAnchor = (ax: string) => /<xdr:contentPart\b/.test(ax);

/** An ink anchor of the drawing as the model's ink (null when its InkML part is missing). */
export function readInk(ax: string, part: string, rels: Rels, files: Record<string, Uint8Array>): SheetInk | null {
  const rid = /<xdr:contentPart\b[^>]*r:id="([^"]+)"/.exec(ax)?.[1];
  const rel = rid ? rels.get(rid) : undefined;
  if (!rel) return null;
  const inkPart = resolveTarget(part, rel.target);
  if (!files[inkPart]) return null;
  const from = point(ax, "from"), to = point(ax, "to");
  if (!from || !to) return null;
  const parsed = parseInkml(strFromU8(files[inkPart]));
  const ext = /<xdr14:xfrm>[\s\S]*?<a:ext cx="(\d+)" cy="(\d+)"/.exec(ax);
  const size = { w: ext ? Number(ext[1]) / EMU_PER_PX : 1, h: ext ? Number(ext[2]) / EMU_PER_PX : 1 };
  const all = parsed.strokes.flatMap((s) => s.pts);
  const xs = all.map((p) => p[0]), ys = all.map((p) => p[1]);
  const x0 = Math.min(...xs), y0 = Math.min(...ys), bw = Math.max(...xs) - x0, bh = Math.max(...ys) - y0;
  // The strokes' box is mapped onto the drawing's box (a line without width or height keeps its scale).
  const sx = bw > 0 ? size.w / bw : 1 / HIM_PER_PX, sy = bh > 0 ? size.h / bh : 1 / HIM_PER_PX;
  const t0 = Math.min(...all.map((p) => p[2]));
  const ink: SheetInk = {
    from, to, size,
    strokes: parsed.strokes.map((s) => ({ brush: s.brush, pts: s.pts.map((p) => ({ x: (p[0] - x0) * sx, y: (p[1] - y0) * sy, t: p[2] - t0 })) })),
    name: /<xdr14:cNvPr\b[^>]*name="([^"]*)"/.exec(ax)?.[1],
    time: parsed.time, raw: ax, part: inkPart,
  };
  const blip = /<mc:Fallback>[\s\S]*?<a:blip\b[^>]*r:embed="([^"]+)"/.exec(ax)?.[1];
  const img = blip ? rels.get(blip) : undefined;
  if (img) {
    const bytes = files[resolveTarget(part, img.target)];
    if (bytes) {
      let bin = "";
      for (let i = 0; i < bytes.length; i += 0x8000) bin += String.fromCharCode(...bytes.subarray(i, i + 0x8000));
      ink.fallback = `data:image/png;base64,${btoa(bin)}`;
    }
  }
  ink.rawKey = inkKey(ink);
  return ink;
}

const pointXml = (tag: string, p: { col: number; row: number; x: number; y: number }) =>
  `<xdr:${tag}><xdr:col>${p.col}</xdr:col><xdr:colOff>${Math.round(p.x * EMU_PER_PX)}</xdr:colOff><xdr:row>${p.row}</xdr:row><xdr:rowOff>${Math.round(p.y * EMU_PER_PX)}</xdr:rowOff></xdr:${tag}>`;

const NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

/** The anchor of an ink drawing in Excel's form; `fallbackRid`: the picture of the ink, if any. */
export function inkAnchorXml(k: SheetInk, id: number, rid: string, fallbackRid: string | undefined, off: { x: number; y: number }): string {
  const name = esc(k.name ?? `Freihand ${id - 1}`);
  const w = Math.round(k.size.w * EMU_PER_PX), hh = Math.round(k.size.h * EMU_PER_PX);
  const x = Math.round(off.x * EMU_PER_PX), y = Math.round(off.y * EMU_PER_PX);
  // The picture covers the strokes with half the widest pen around them (recorded: 0.035 cm pen, 6121 EMU each side).
  const pad = Math.round((Math.max(...k.strokes.map((s) => s.brush.width), 0) / 2) * 36000);
  const fallback = fallbackRid
    ? `<mc:Fallback><xdr:pic><xdr:nvPicPr><xdr:cNvPr id="${id}" name="${name}"/><xdr:cNvPicPr/></xdr:nvPicPr><xdr:blipFill><a:blip xmlns:r="${NS_R}" r:embed="${fallbackRid}"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill>`
      + `<xdr:spPr><a:xfrm><a:off x="${x - pad}" y="${y - pad}"/><a:ext cx="${w + 2 * pad}" cy="${hh + 2 * pad}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic></mc:Fallback>`
    : "<mc:Fallback/>";
  return `<xdr:twoCellAnchor editAs="oneCell">${pointXml("from", k.from)}${pointXml("to", k.to)}`
    + `<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:Choice xmlns:xdr14="http://schemas.microsoft.com/office/excel/2010/spreadsheetDrawing" Requires="xdr14">`
    + `<xdr:contentPart xmlns:r="${NS_R}" r:id="${rid}"><xdr14:nvContentPartPr><xdr14:cNvPr id="${id}" name="${name}"/><xdr14:cNvContentPartPr/></xdr14:nvContentPartPr>`
    + `<xdr14:nvPr macro=""/><xdr14:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${w}" cy="${hh}"/></xdr14:xfrm></xdr:contentPart></mc:Choice>${fallback}</mc:AlternateContent><xdr:clientData/></xdr:twoCellAnchor>`;
}

/** A moved ink drawing whose strokes did not change: the kept anchor with new corners. */
export function movedInkXml(k: SheetInk): string {
  return k.raw!.replace(/<xdr:from>[\s\S]*?<\/xdr:from>/, pointXml("from", k.from)).replace(/<xdr:to>[\s\S]*?<\/xdr:to>/, pointXml("to", k.to));
}

/** Writes an ink drawing's parts (InkML, picture) and returns its anchor. */
export function writeInk(pw: PackageWriter, drawingPart: string, k: SheetInk, id: number, off: { x: number; y: number }): string {
  const changedStrokes = !k.raw || !k.rawKey || JSON.stringify(JSON.parse(k.rawKey)[4]) !== JSON.stringify(k.strokes);
  if (k.raw && !changedStrokes) {
    const r = k.rawKey ? JSON.parse(k.rawKey) : null;
    const same = r && JSON.stringify([r[0], r[1], r[2], r[3]]) === JSON.stringify([k.from, k.to, k.size, k.name ?? null]);
    if (same) return k.raw;
    if (r && JSON.stringify(r[2]) === JSON.stringify(k.size) && r[3] === (k.name ?? null)) return movedInkXml(k);
  }
  const inkPart = pw.freeName("xl/ink/ink");
  pw.setPart(inkPart, inkmlXml(k), INK_CONTENT_TYPE);
  const rid = pw.link(drawingPart, INK_REL_TYPE, inkPart);
  let fallbackRid: string | undefined;
  const m = k.fallback ? /^data:([^;]+);base64,(.*)$/.exec(k.fallback) : null;
  if (m) {
    const bin = atob(m[2]);
    const bytes = new Uint8Array(bin.length);
    for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
    fallbackRid = pw.link(drawingPart, IMAGE_REL_TYPE, pw.addMedia(bytes, m[1], "xl/media/image"));
  }
  return inkAnchorXml(k, id, rid, fallbackRid, off);
}

/** The sheet's ink drawings (the model's `inks`). */
export const inksOf = (s: Sheet): SheetInk[] => s.inks ?? [];
