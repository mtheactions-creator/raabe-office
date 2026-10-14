// Writing the model back to .docx. With the original package, parts whose content did not
// change (headers, notes, numbering …) are written back byte for byte, all preserved XML
// goes back in its place, and new content gets new parts and relationships.

import { blipEffectsXml, imgEffectsXml, shapePropsXml } from "./drawingml";
import { strFromU8, strToU8, zipSync } from "fflate";
import {
  COLUMN_BREAK, LINE_BREAK, OBJ, PAGE_BREAK, effectExtent, type AltText, type CommentMeta, type PictureFormat, type DocState, type FloatObject, type InlineObject,
  type Para, type Revision, type Section, type StyleDef, type TableDef,
} from "@raabe/text-engine";
import { latexToOmml, M_NS } from "./math";
import { writeNumbering } from "./numbering";
import { CT, PackageWriter, REL, relink } from "./package";
import type { DocxPackage, ObjectData } from "./read";
import { bordersXml, revisionAttrs, shadingXml, writePPr, writeRPr, type WriteOptions } from "./props";
import { esc, R, W } from "./xml";

const NBHYPHEN = String.fromCharCode(0x2011);
const SOFTHYPHEN = String.fromCharCode(0xad);
const EMU = 12700;
const NS_WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
const NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main";
const NS_PIC = "http://schemas.openxmlformats.org/drawingml/2006/picture";
const W14 = "http://schemas.microsoft.com/office/word/2010/wordml";
const W15 = "http://schemas.microsoft.com/office/word/2012/wordml";

const XML_HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
const DEFAULT_ROOT = (tag: string) => `<w:${tag} xmlns:w="${W}" xmlns:r="${R}" xmlns:wp="${NS_WP}" xmlns:w14="${W14}" xmlns:m="${M_NS}">`;
/** Declares the math namespace on a part's root element when its content uses it. */
const withMath = (root: string, body: string) => {
  if (body.includes("<m:") && !root.includes("xmlns:m=")) root = root.replace(/>$/, ` xmlns:m="${M_NS}">`);
  // Emoji (w16se:symEx inside mc:AlternateContent).
  if (body.includes("<w16se:")) {
    if (!root.includes("xmlns:mc=")) root = root.replace(/>$/, ' xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">');
    if (!root.includes("xmlns:w16se=")) root = root.replace(/>$/, ' xmlns:w16se="http://schemas.microsoft.com/office/word/2015/wordml/symex">');
    const ign = /mc:Ignorable="([^"]*)"/.exec(root);
    if (!ign) root = root.replace(/>$/, ' mc:Ignorable="w16se">');
    else if (!ign[1].split(/\s+/).includes("w16se")) root = root.replace(ign[0], `mc:Ignorable="${ign[1]} w16se"`);
  }
  return root;
};

function splitExtra(list: readonly unknown[] | undefined): { name: string; xml: string }[] {
  return (list ?? []).map((x) => {
    const s = String(x);
    const i = s.indexOf("\u0000");
    return i < 0 ? { name: "zzz", xml: s } : { name: s.slice(0, i), xml: s.slice(i + 1) };
  });
}

function ordered(parts: { name: string; xml: string }[], order: string[]): string {
  const idx = (n: string) => { const i = order.indexOf(n); return i < 0 ? order.length : i; };
  return parts.map((p, i) => ({ ...p, i })).sort((a, b) => idx(a.name) - idx(b.name) || a.i - b.i).map((p) => p.xml).join("");
}

const SECT_ORDER = ["headerReference", "footerReference", "footnotePr", "endnotePr", "type", "pgSz", "pgMar", "paperSrc", "pgBorders",
  "lnNumType", "pgNumType", "cols", "formProt", "vAlign", "noEndnote", "titlePg", "textDirection", "bidi", "rtlGutter", "docGrid",
  "printerSettings", "sectPrChange"];
const TBLPR_ORDER = ["tblStyle", "tblpPr", "tblOverlap", "bidiVisual", "tblStyleRowBandSize", "tblStyleColBandSize", "tblW", "jc",
  "tblCellSpacing", "tblInd", "tblBorders", "shd", "tblLayout", "tblCellMar", "tblLook", "tblCaption", "tblDescription", "tblPrChange"];
const TRPR_ORDER = ["cnfStyle", "divId", "gridBefore", "gridAfter", "wBefore", "wAfter", "cantSplit", "trHeight", "tblHeader",
  "tblCellSpacing", "jc", "hidden", "ins", "del", "trPrChange"];
const TCPR_ORDER = ["cnfStyle", "tcW", "gridSpan", "hMerge", "vMerge", "tcBorders", "shd", "noWrap", "tcMar", "textDirection",
  "tcFitText", "vAlign", "hideMark", "headers", "cellIns", "cellDel", "cellMerge", "tcPrChange"];

const tw = (pt: number) => Math.round(pt * 20);

function isData(d: unknown): d is ObjectData {
  return typeof d === "object" && d !== null && ("raw" in d || "inner" in d || "drawing" in d);
}

/** Preserved XML of a deleted revision uses w:delText for its text. */
function asDeleted(xml: string): string {
  return xml.replace(/<w:t(\s|>)/g, "<w:delText$1").replace(/<\/w:t>/g, "</w:delText>")
    .replace(/<w:instrText(\s|>)/g, "<w:delInstrText$1").replace(/<\/w:instrText>/g, "</w:delInstrText>");
}

class DocWriter {
  private docPrId = 1000;
  private paraIdSeq = 0x10000000;
  /** Story ids of headers/footers → part path. */
  readonly storyPaths = new Map<string, string>();

  constructor(readonly pw: PackageWriter, readonly state: DocState, readonly pkg: DocxPackage | null, readonly opts: WriteOptions) {}

  // ----- blocks ---------------------------------------------------------------------------

  /** Paragraphs, tables and content-control groups of a story, as body XML. */
  blocks(paras: readonly Para[], part: string, paraAttrs?: (i: number) => string): string {
    const fields = this.validFields(paras);
    let out = "";
    let group: Para["group"] | undefined;
    for (let i = 0; i < paras.length; i++) {
      const p = paras[i];
      if (p.group?.id !== group?.id) {
        if (group) out += String(group.close);
        group = p.group;
        if (group) out += String(group.open);
      }
      if (p.cell) {
        let j = i;
        while (j + 1 < paras.length && paras[j + 1].cell?.table.id === p.cell.table.id) j++;
        out += this.table(paras.slice(i, j + 1), part, fields);
        i = j;
        continue;
      }
      out += this.para(p, part, fields, paraAttrs?.(i) ?? "");
    }
    if (group) out += String(group.close);
    return out;
  }

  /** Ids of multi-paragraph fields whose start and end are both still there (in order). */
  private validFields(paras: readonly Para[]): Set<string> {
    const started = new Set<string>(), ok = new Set<string>();
    for (const p of paras) {
      for (const a of [...(p.anchors ?? [])].sort((x, y) => x.at - y.at)) {
        if (!a.field) continue;
        if (a.field.edge === "start") started.add(a.field.id);
        else if (started.has(a.field.id)) ok.add(a.field.id);
      }
    }
    return ok;
  }

  para(p: Para, part: string, fields: Set<string>, attrs = ""): string {
    const leading = (p.leading ?? []).map(String).join("");
    if (p.block) return leading + String(p.block.data);
    const sect = p.sect ? this.section(p.sect, part) : "";
    return `${leading}<w:p${attrs}>${writePPr(p.props, p.mark, this.opts, sect)}${this.runs(p, part, fields)}</w:p>`;
  }

  private runs(p: Para, part: string, fields: Set<string>): string {
    let out = "";
    const anchors = [...(p.anchors ?? [])].sort((a, b) => a.at - b.at);
    const floats = [...(p.floats ?? [])].sort((a, b) => a.at - b.at);
    // Hyperlinks must start and end in one paragraph.
    const links = new Map<string, number>();
    for (const a of anchors) if (a.link) links.set(a.link.id, (links.get(a.link.id) ?? 0) + 1);
    let ai = 0, fi = 0, objIdx = 0, pos = 0;
    let rev: Revision | undefined;
    const closeRev = () => { if (rev) { out += `</w:${rev.kind}>`; rev = undefined; } };
    const openRev = (r: Revision | undefined) => {
      if (r === rev) return;
      closeRev();
      if (r) { out += `<w:${r.kind} ${revisionAttrs(r)}>`; rev = r; }
    };
    const markers = (i: number, rPr: string) => {
      let s = "";
      while (ai < anchors.length && anchors[ai].at <= i) s += this.anchor(anchors[ai++], fields, links);
      while (fi < floats.length && floats[fi].at <= i) s += `<w:r>${rPr}${this.floatXml(floats[fi++], part)}</w:r>`;
      if (s) { closeRev(); out += s; }
    };
    for (const run of p.runs) {
      const rPr = writeRPr(run.props, this.opts);
      const del = run.props.rev?.kind === "del";
      let inner = "";
      let text = "";
      const flushText = () => {
        if (text) { inner += del ? `<w:delText xml:space="preserve">${esc(text)}</w:delText>` : `<w:t xml:space="preserve">${esc(text)}</w:t>`; text = ""; }
      };
      const flushRun = () => { flushText(); if (inner) { openRev(run.props.rev); out += `<w:r>${rPr}${inner}</w:r>`; inner = ""; } };
      for (let i = pos; i < pos + run.len; i++) {
        if ((ai < anchors.length && anchors[ai].at <= i) || (fi < floats.length && floats[fi].at <= i)) { flushRun(); markers(i, rPr); }
        const ch = p.text[i];
        if (ch === OBJ) {
          flushRun();
          openRev(run.props.rev);
          const xml = this.object(p.objects?.[objIdx++], rPr, part, p.text === OBJ);
          out += del ? asDeleted(xml) : xml;
          continue;
        }
        if (ch === "\t") { flushText(); inner += "<w:tab/>"; }
        else if (ch === LINE_BREAK) { flushText(); inner += "<w:br/>"; }
        else if (ch === PAGE_BREAK) { flushText(); inner += '<w:br w:type="page"/>'; }
        else if (ch === COLUMN_BREAK) { flushText(); inner += '<w:br w:type="column"/>'; }
        else if (ch === NBHYPHEN) { flushText(); inner += "<w:noBreakHyphen/>"; }
        else if (ch === SOFTHYPHEN) { flushText(); inner += "<w:softHyphen/>"; }
        else if (run.props.font === "Segoe UI Emoji" && !del && (p.text.codePointAt(i) ?? 0) >= 0x2600) {
          // Emoji the way Word writes them, so Word treats them as its own symbols.
          flushText();
          const cp = p.text.codePointAt(i)!;
          const str = String.fromCodePoint(cp);
          inner += `<mc:AlternateContent><mc:Choice Requires="w16se"><w16se:symEx w16se:font="Segoe UI Emoji" w16se:char="${cp.toString(16).toUpperCase()}"/></mc:Choice>` +
            `<mc:Fallback><w:t>${str}</w:t></mc:Fallback></mc:AlternateContent>`;
          if (str.length === 2) i++;
        }
        else text += ch;
      }
      flushRun();
      pos += run.len;
    }
    closeRev();
    markers(Number.MAX_SAFE_INTEGER, writeRPr(p.mark, this.opts));
    return out;
  }

  private anchor(a: NonNullable<Para["anchors"]>[number], fields: Set<string>, links: Map<string, number>): string {
    if (a.field && !fields.has(a.field.id)) return "";
    if (a.link && links.get(a.link.id) !== 2) return "";
    if (a.comment && !this.state.stories?.has(a.comment.id)) return "";
    const d = a.data;
    if (isData(d) && "raw" in d) return d.raw;
    if (typeof d === "string") return d;
    if (a.comment) {
      const n = a.comment.id.slice(4);
      if (a.comment.edge === "start") return `<w:commentRangeStart w:id="${n}"/>`;
      if (a.comment.edge === "end") return `<w:commentRangeEnd w:id="${n}"/>`;
      const style = this.styleId("annotation reference");
      return `<w:r>${style ? `<w:rPr><w:rStyle w:val="${esc(style)}"/></w:rPr>` : ""}<w:commentReference w:id="${n}"/></w:r>`;
    }
    return "";
  }

  private styleId(name: string): string | undefined {
    for (const s of this.state.styles.styles.values()) if (s.name.toLowerCase() === name) return s.id;
    return undefined;
  }

  private object(o: InlineObject | undefined, rPr: string, part: string, alone = false): string {
    if (!o) return "";
    const d = o.data;
    if (typeof d === "string") return d;
    if (isData(d)) {
      if ("raw" in d) return d.raw;
      if ("inner" in d) return `<w:r>${rPr}${d.inner}</w:r>`;
      return `<w:r>${rPr}${relink(this.repictured(this.resized(d.drawing, o.width, o.height), o, d.picRead), d.part, part, this.pw)}</w:r>`;
    }
    // Content created in RaabeWrite.
    if (o.kind === "image" && o.image) return `<w:r>${rPr}${this.newDrawing(o, part)}</w:r>`;
    // An equation alone in its paragraph is a display equation (m:oMathPara), as in Word.
    if (o.kind === "math" && o.latex !== undefined) return latexToOmml(o.latex, alone);
    if (o.kind === "note" && o.note) {
      const [kind, n] = o.note.split(":");
      const tag = kind === "en" ? "endnote" : "footnote";
      return `<w:r>${rPr}<w:${tag}Reference w:id="${esc(n)}"/></w:r>`;
    }
    if (o.kind === "field" && o.instr) {
      const r = (inner: string) => `<w:r>${rPr}${inner}</w:r>`;
      return r('<w:fldChar w:fldCharType="begin"/>') + r(`<w:instrText xml:space="preserve"> ${esc(o.instr)} </w:instrText>`) +
        r('<w:fldChar w:fldCharType="separate"/>') + r(`<w:t xml:space="preserve">${esc(o.display)}</w:t>`) + r('<w:fldChar w:fldCharType="end"/>');
    }
    return o.display ? `<w:r>${rPr}<w:t xml:space="preserve">${esc(o.display)}</w:t></w:r>` : "";
  }

  /**
   * Writes changed Bildformat settings and Alternativtext into a picture's XML (blip effects,
   * cropping, shape properties, docPr); unchanged pictures keep their XML as read.
   */
  private repictured(xml: string, o: { pic?: PictureFormat; alt?: AltText }, read: string | undefined): string {
    if (JSON.stringify({ pic: o.pic, alt: o.alt }) === (read ?? "{}")) return xml;
    const f = o.pic ?? {};
    let out = xml.replace(/<a:blip ([^>]*?)(?:\/>|>([\s\S]*?)<\/a:blip>)/, (_m, attrs: string, inner: string | undefined) => {
      const exts = (inner ?? "").match(/<a:ext uri="[^"]*">[\s\S]*?<\/a:ext>/g) ?? [];
      const kept = exts.filter((x) => !x.includes("{BEBA8EAE-BF5A-486C-A8C5-ECC9F3942E4B}")).join("");
      const ext = kept + imgEffectsXml(f);
      return `<a:blip ${attrs}>${blipEffectsXml(f)}${ext ? `<a:extLst>${ext}</a:extLst>` : ""}</a:blip>`;
    });
    out = out.replace(/<a:srcRect[^>]*\/>/, "");
    const c = f.crop;
    if (c && (c.l || c.t || c.r || c.b)) {
      const v = (x: number) => Math.round(x * 100000);
      out = out.replace(/(<\/a:blip>)/, `$1<a:srcRect${c.l ? ` l="${v(c.l)}"` : ""}${c.t ? ` t="${v(c.t)}"` : ""}${c.r ? ` r="${v(c.r)}"` : ""}${c.b ? ` b="${v(c.b)}"` : ""}/>`);
    }
    out = out.replace(/<pic:spPr([^>]*)>([\s\S]*?)<\/pic:spPr>/, (_m, attrs: string, inner: string) => {
      const ext = /<a:ext cx="(\d+)" cy="(\d+)"/.exec(inner);
      return `<pic:spPr${attrs}>${shapePropsXml(f, Number(ext?.[1] ?? 0), Number(ext?.[2] ?? 0))}</pic:spPr>`;
    });
    const ext = /<wp:extent cx="(\d+)" cy="(\d+)"/.exec(out);
    const ex = effectExtent(f, Number(ext?.[2] ?? 0) / 12700);
    const emu = (x: number) => Math.round(x * 12700);
    out = out.replace(/<wp:effectExtent [^>]*\/>/, `<wp:effectExtent l="${emu(ex.l)}" t="${emu(ex.t)}" r="${emu(ex.r)}" b="${emu(ex.b)}"/>`);
    out = out.replace(/<wp:docPr ([^>]*?)(\/?)>/, (_m, attrs: string, close: string) => {
      let a = attrs.replace(/\s*(descr|title)="[^"]*"/g, "");
      if (o.alt?.descr) a += ` descr="${esc(o.alt.descr)}"`;
      if (o.alt?.title) a += ` title="${esc(o.alt.title)}"`;
      return `<wp:docPr ${a.trim()}${close}>`;
    });
    return out;
  }

  /** Updates the size in a picture's XML when it was resized. */
  private resized(xml: string, width: number | undefined, height: number | undefined): string {
    if (width === undefined || height === undefined) return xml;
    const cx = Math.round(width * EMU), cy = Math.round(height * EMU);
    const m = /<wp:extent cx="(\d+)" cy="(\d+)"/.exec(xml);
    if (!m || (Math.abs(Number(m[1]) - cx) <= 1 && Math.abs(Number(m[2]) - cy) <= 1)) return xml;
    let out = xml.replace(/<wp:extent cx="\d+" cy="\d+"/, `<wp:extent cx="${cx}" cy="${cy}"`);
    const sp = out.indexOf("<pic:spPr");
    if (sp >= 0) out = out.slice(0, sp) + out.slice(sp).replace(/<a:ext cx="\d+" cy="\d+"/, `<a:ext cx="${cx}" cy="${cy}"`);
    return out;
  }

  private newDrawing(o: { image?: InlineObject["image"]; width?: number; height?: number; pic?: PictureFormat; alt?: AltText }, part: string, anchor?: FloatObject): string {
    const img = o.image!;
    const media = this.pw.addMedia(img.bytes, img.type);
    const rid = this.pw.link(part, REL.image, media);
    const cx = Math.round((o.width ?? 100) * EMU), cy = Math.round((o.height ?? 100) * EMU);
    const id = this.docPrId++;
    const f = o.pic ?? {};
    const fx = imgEffectsXml(f);
    const crop = f.crop && (f.crop.l || f.crop.t || f.crop.r || f.crop.b)
      ? `<a:srcRect l="${Math.round(f.crop.l * 100000)}" t="${Math.round(f.crop.t * 100000)}" r="${Math.round(f.crop.r * 100000)}" b="${Math.round(f.crop.b * 100000)}"/>` : "";
    const graphic = `<a:graphic xmlns:a="${NS_A}"><a:graphicData uri="${NS_PIC}"><pic:pic xmlns:pic="${NS_PIC}">` +
      `<pic:nvPicPr><pic:cNvPr id="0" name="Bild ${id}"${o.alt?.descr ? ` descr="${esc(o.alt.descr)}"` : ""}/><pic:cNvPicPr/></pic:nvPicPr>` +
      `<pic:blipFill><a:blip r:embed="${rid}">${blipEffectsXml(f)}${fx ? `<a:extLst>${fx}</a:extLst>` : ""}</a:blip>${crop}<a:stretch><a:fillRect/></a:stretch></pic:blipFill>` +
      `<pic:spPr>${shapePropsXml(f, cx, cy)}</pic:spPr></pic:pic></a:graphicData></a:graphic>`;
    const ex = effectExtent(f, (o.height ?? 0));
    const common = `<wp:extent cx="${cx}" cy="${cy}"/><wp:effectExtent l="${Math.round(ex.l * EMU)}" t="${Math.round(ex.t * EMU)}" r="${Math.round(ex.r * EMU)}" b="${Math.round(ex.b * EMU)}"/>`;
    const docPr = `<wp:docPr id="${id}" name="Grafik ${id}"${o.alt?.descr ? ` descr="${esc(o.alt.descr)}"` : ""}${o.alt?.title ? ` title="${esc(o.alt.title)}"` : ""}/><wp:cNvGraphicFramePr><a:graphicFrameLocks xmlns:a="${NS_A}" noChangeAspect="1"/></wp:cNvGraphicFramePr>`;
    if (!anchor) return `<w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0" xmlns:wp="${NS_WP}">${common}${docPr}${graphic}</wp:inline></w:drawing>`;
    const e = (pt: number) => Math.round(pt * EMU);
    const pos = (tag: string, p: FloatObject["x"]) => `<wp:${tag} relativeFrom="${esc(p.rel)}">${p.align ? `<wp:align>${esc(p.align)}</wp:align>` : `<wp:posOffset>${e(p.offset)}</wp:posOffset>`}</wp:${tag}>`;
    const wrap = anchor.wrap === "front" || anchor.wrap === "behind" || anchor.wrap === "none" ? "<wp:wrapNone/>"
      : anchor.wrap === "topAndBottom" ? "<wp:wrapTopAndBottom/>"
      : `<wp:wrap${anchor.wrap === "tight" ? "Tight" : anchor.wrap === "through" ? "Through" : "Square"} wrapText="${esc(anchor.side)}"${anchor.wrap === "square" ? "" : "><wp:wrapPolygon edited=\"0\"><wp:start x=\"0\" y=\"0\"/><wp:lineTo x=\"0\" y=\"21600\"/><wp:lineTo x=\"21600\" y=\"21600\"/><wp:lineTo x=\"21600\" y=\"0\"/><wp:lineTo x=\"0\" y=\"0\"/></wp:wrapPolygon></wp:wrap" + (anchor.wrap === "tight" ? "Tight" : "Through")}/>`;
    const [t, b, l, r] = anchor.dist.map(e);
    return `<w:drawing><wp:anchor distT="${t}" distB="${b}" distL="${l}" distR="${r}" simplePos="0" relativeHeight="${anchor.z ?? id}" behindDoc="${anchor.wrap === "behind" ? 1 : 0}" locked="0" layoutInCell="1" allowOverlap="1" xmlns:wp="${NS_WP}">` +
      `<wp:simplePos x="0" y="0"/>${pos("positionH", anchor.x)}${pos("positionV", anchor.y)}${common}${wrap}${docPr}${graphic}</wp:anchor></w:drawing>`;
  }

  private floatXml(f: FloatObject, part: string): string {
    const d = f.data;
    if (isData(d) && "drawing" in d) return relink(this.repictured(this.repositioned(this.resized(d.drawing, f.width, f.height), f), f, d.picRead), d.part, part, this.pw);
    if (typeof d === "string") return d;
    if (f.image) return this.newDrawing(f, part, f);
    return "";
  }

  /** Updates the offsets of a floating picture that was moved. */
  private repositioned(xml: string, f: FloatObject): string {
    let out = xml;
    if (f.z !== undefined) out = out.replace(/(<wp:anchor [^>]*?relativeHeight=")\d+(")/, `$1${f.z}$2`);
    const behind = f.wrap === "behind" ? "1" : "0";
    out = out.replace(/(<wp:anchor [^>]*?behindDoc=")[01](")/, `$1${behind}$2`);
    for (const [tag, p] of [["positionH", f.x], ["positionV", f.y]] as const) {
      if (p.align) continue;
      const re = new RegExp(`(<wp:${tag}[^>]*>\\s*<wp:posOffset>)(-?\\d+)(</wp:posOffset>)`);
      const m = re.exec(out);
      const v = Math.round(p.offset * EMU);
      if (m && Math.abs(Number(m[2]) - v) > 1) out = out.replace(re, `$1${v}$3`);
    }
    return out;
  }

  // ----- sections -------------------------------------------------------------------------

  section(s: Section, part: string): string {
    const parts: { name: string; xml: string }[] = [];
    const extra = splitExtra(s.extra);
    const add = (name: string, xml: string) => parts.push({ name, xml });
    const refs = (kind: "header" | "footer", r: Section["headers"]) => {
      for (const type of ["default", "first", "even"] as const) {
        const id = r[type];
        if (!id || !this.state.stories?.has(id)) continue;
        const path = this.storyPath(id);
        add(`${kind}Reference`, `<w:${kind}Reference w:type="${type}" r:id="${this.pw.link(part, REL[kind], path)}"/>`);
      }
    };
    refs("header", s.headers);
    refs("footer", s.footers);
    if (s.start !== "nextPage") add("type", `<w:type w:val="${s.start}"/>`);
    const pg = s.page;
    const orient = extra.find((e) => e.name === "pgSzOrient")?.xml ?? (pg.width > pg.height ? "landscape" : undefined);
    const landscape = pg.width > pg.height;
    const code = extra.find((e) => e.name === "pgSzCode")?.xml;
    add("pgSz", `<w:pgSz w:w="${tw(pg.width)}" w:h="${tw(pg.height)}"${(orient === "landscape") === landscape && orient ? ` w:orient="${orient}"` : landscape ? ' w:orient="landscape"' : ""}${code ? ` w:code="${esc(code)}"` : ""}/>`);
    add("pgMar", `<w:pgMar w:top="${tw(pg.marginTop)}" w:right="${tw(pg.marginRight)}" w:bottom="${tw(pg.marginBottom)}" w:left="${tw(pg.marginLeft)}" w:header="${tw(pg.header)}" w:footer="${tw(pg.footer)}" w:gutter="${tw(pg.gutter)}"/>`);
    if (!extra.some((e) => e.name === "pgNumType") && (s.pageStart !== undefined || s.pageFormat)) {
      add("pgNumType", `<w:pgNumType${s.pageFormat ? ` w:fmt="${esc(s.pageFormat)}"` : ""}${s.pageStart !== undefined ? ` w:start="${s.pageStart}"` : ""}/>`);
    }
    const c = s.columns;
    const cols = c.cols ? c.cols.map((x) => `<w:col w:w="${tw(x.width)}" w:space="${tw(x.space)}"/>`).join("") : "";
    add("cols", `<w:cols${c.count > 1 ? ` w:num="${c.count}"` : ""} w:space="${tw(c.space)}"${c.sep ? ' w:sep="1"' : ""}${c.cols ? ' w:equalWidth="0"' : ""}${cols ? `>${cols}</w:cols>` : "/>"}`);
    if (s.titlePage) add("titlePg", "<w:titlePg/>");
    for (const e of extra) if (e.name !== "pgSzOrient" && e.name !== "pgSzCode") parts.push(e);
    return `<w:sectPr>${ordered(parts, SECT_ORDER)}</w:sectPr>`;
  }

  storyPath(id: string): string {
    let path = this.storyPaths.get(id) ?? this.pkg?.storyParts.get(id)?.path;
    if (!path) path = this.pw.freeName(id.startsWith("hdr:") ? "word/header" : "word/footer");
    this.storyPaths.set(id, path);
    return path;
  }

  // ----- tables ---------------------------------------------------------------------------

  private table(paras: readonly Para[], part: string, fields: Set<string>): string {
    const def: TableDef = paras[0].cell!.table;
    const extra = splitExtra(def.extra);
    const pr: { name: string; xml: string }[] = [];
    if (def.style) pr.push({ name: "tblStyle", xml: `<w:tblStyle w:val="${esc(def.style)}"/>` });
    if (!extra.some((e) => e.name === "tblW")) pr.push({ name: "tblW", xml: '<w:tblW w:w="0" w:type="auto"/>' });
    if (def.align && def.align !== "left") pr.push({ name: "jc", xml: `<w:jc w:val="${def.align}"/>` });
    if (def.indent !== undefined) pr.push({ name: "tblInd", xml: `<w:tblInd w:w="${tw(def.indent)}" w:type="dxa"/>` });
    if (def.borders) pr.push({ name: "tblBorders", xml: bordersXml("tblBorders", def.borders, "table") });
    if (def.shading) pr.push({ name: "shd", xml: shadingXml(def.shading) });
    if (def.cellMargins) pr.push({ name: "tblCellMar", xml: marginsXml("tblCellMar", def.cellMargins) });
    if (def.look) {
      const l = def.look;
      const v = (l.firstRow ? 0x20 : 0) | (l.lastRow ? 0x40 : 0) | (l.firstColumn ? 0x80 : 0) | (l.lastColumn ? 0x100 : 0) | (l.noHBand ? 0x200 : 0) | (l.noVBand ? 0x400 : 0);
      const b = (x: boolean) => (x ? "1" : "0");
      pr.push({ name: "tblLook", xml: `<w:tblLook w:val="${v.toString(16).toUpperCase().padStart(4, "0")}" w:firstRow="${b(l.firstRow)}" w:lastRow="${b(l.lastRow)}" w:firstColumn="${b(l.firstColumn)}" w:lastColumn="${b(l.lastColumn)}" w:noHBand="${b(l.noHBand)}" w:noVBand="${b(l.noVBand)}"/>` });
    }
    pr.push(...extra);
    let out = `<w:tbl><w:tblPr>${ordered(pr, TBLPR_ORDER)}</w:tblPr><w:tblGrid>${def.grid.map((g) => `<w:gridCol w:w="${tw(g)}"/>`).join("")}</w:tblGrid>`;
    def.rows.forEach((row, r) => {
      const rowExtra = splitExtra(row.extra);
      const prEx = rowExtra.find((e) => e.name === "tblPrEx")?.xml ?? "";
      const trPr: { name: string; xml: string }[] = rowExtra.filter((e) => e.name !== "tblPrEx");
      if (row.cantSplit) trPr.push({ name: "cantSplit", xml: "<w:cantSplit/>" });
      if (row.height !== undefined) trPr.push({ name: "trHeight", xml: `<w:trHeight w:val="${tw(row.height)}"${row.heightRule === "exact" ? ' w:hRule="exact"' : ""}/>` });
      if (row.header) trPr.push({ name: "tblHeader", xml: "<w:tblHeader/>" });
      out += `<w:tr>${prEx}${trPr.length ? `<w:trPr>${ordered(trPr, TRPR_ORDER)}</w:trPr>` : ""}`;
      let gx = 0;
      row.cells.forEach((cd, c) => {
        const span = Math.max(1, cd.span);
        const cellExtra = splitExtra(cd.extra);
        const tc: { name: string; xml: string }[] = [];
        const gridWidth = def.grid.slice(gx, gx + span).reduce((s, x) => s + x, 0);
        const width = cd.width ?? gridWidth;
        const tcW = cellExtra.find((e) => e.name === "tcW");
        const keepW = tcW && !/w:type="dxa"/.test(tcW.xml) && cd.width === undefined;
        tc.push(keepW ? tcW! : { name: "tcW", xml: `<w:tcW w:w="${tw(width)}" w:type="dxa"/>` });
        if (span > 1) tc.push({ name: "gridSpan", xml: `<w:gridSpan w:val="${span}"/>` });
        if (cd.vMerge) tc.push({ name: "vMerge", xml: cd.vMerge === "restart" ? '<w:vMerge w:val="restart"/>' : "<w:vMerge/>" });
        if (cd.borders) tc.push({ name: "tcBorders", xml: bordersXml("tcBorders", cd.borders, "table") });
        if (cd.shading) tc.push({ name: "shd", xml: shadingXml(cd.shading) });
        if (cd.margins) tc.push({ name: "tcMar", xml: marginsXml("tcMar", cd.margins) });
        if (cd.vAlign && cd.vAlign !== "top") tc.push({ name: "vAlign", xml: `<w:vAlign w:val="${cd.vAlign}"/>` });
        tc.push(...cellExtra.filter((e) => e.name !== "tcW"));
        const content = paras.filter((p) => p.cell!.row === r && p.cell!.col === c);
        let inner = content.map((p) => (p.block ? String(p.block.data) : this.para(p, part, fields))).join("");
        if (!content.length || content[content.length - 1].block) inner += "<w:p/>";
        out += `<w:tc><w:tcPr>${ordered(tc, TCPR_ORDER)}</w:tcPr>${inner}</w:tc>`;
        gx += span;
      });
      out += "</w:tr>";
    });
    return `${out}</w:tbl>`;
  }

  // ----- the parts ------------------------------------------------------------------------

  /** Headers and footers: changed or new stories get (new) parts. */
  writeHeaders(): void {
    for (const [id, paras] of this.state.stories ?? []) {
      if (!id.startsWith("hdr:") && !id.startsWith("ftr:")) continue;
      const orig = this.pkg?.storyParts.get(id);
      const path = this.storyPath(id);
      if (orig && this.pkg?.originalStories?.get(id) === paras) continue;
      const tag = id.startsWith("hdr:") ? "hdr" : "ftr";
      const root = orig?.rootTag ?? DEFAULT_ROOT(tag);
      const xml = this.blocks(paras, path);
      this.pw.setPart(path, `${XML_HEAD}${withMath(root, xml)}${xml}</w:${tag}>`, tag === "hdr" ? CT.header : CT.footer);
    }
  }

  /** Foot- and endnotes: the part is rewritten when a note changed, was added or removed. */
  writeNotes(kind: "fn" | "en"): void {
    const name = kind === "fn" ? "footnotes" : "endnotes";
    const tag = kind === "fn" ? "footnote" : "endnote";
    // Notes whose reference was deleted are left out (as Word does).
    const referenced = new Set<string>();
    for (const p of this.state.paras) for (const o of p.objects ?? []) if (o.note) referenced.add(o.note);
    const ids = [...(this.state.stories?.keys() ?? [])].filter((k) => k.startsWith(`${kind}:`) && referenced.has(k));
    const orig = this.pkg?.notes[name];
    const origIds = [...(this.pkg?.originalStories?.keys() ?? [])].filter((k) => k.startsWith(`${kind}:`));
    const unchanged = ids.length === origIds.length && ids.every((k) => this.pkg?.originalStories?.get(k) === this.state.stories?.get(k));
    if (unchanged && (orig || !ids.length)) return;
    const path = orig?.path ?? `word/${name}.xml`;
    const root = orig?.rootTag ?? DEFAULT_ROOT(name);
    const special = orig?.special ?? [
      `<w:${tag} w:type="separator" w:id="-1"><w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:separator/></w:r></w:p></w:${tag}>`,
      `<w:${tag} w:type="continuationSeparator" w:id="0"><w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:continuationSeparator/></w:r></w:p></w:${tag}>`,
    ];
    const notes = ids.sort((a, b) => Number(a.slice(3)) - Number(b.slice(3)))
      .map((id) => `<w:${tag} w:id="${esc(id.slice(3))}">${this.blocks(this.state.stories!.get(id)!, path)}</w:${tag}>`);
    this.pw.setPart(path, `${XML_HEAD}${withMath(root, notes.join(""))}${special.join("")}${notes.join("")}</w:${name}>`, kind === "fn" ? CT.footnotes : CT.endnotes);
    this.pw.link("word/document.xml", kind === "fn" ? REL.footnotes : REL.endnotes, path);
  }

  writeComments(): void {
    const ids = [...(this.state.stories?.keys() ?? [])].filter((k) => k.startsWith("cmt:"));
    const origIds = [...(this.pkg?.originalStories?.keys() ?? [])].filter((k) => k.startsWith("cmt:"));
    const meta = this.state.comments ?? new Map<string, CommentMeta>();
    if (!ids.length && !this.pkg?.comments) return;
    const om = this.pkg?.originalComments;
    const unchanged = ids.length === origIds.length && ids.every((k) => this.pkg?.originalStories?.get(k) === this.state.stories?.get(k)) &&
      !!om && om.size === meta.size && [...meta].every(([k, v]) => om.get(k) === v);
    if (unchanged && (this.pkg?.comments || !ids.length)) return;
    const path = this.pkg?.comments?.path ?? "word/comments.xml";
    const root = this.pkg?.comments?.rootTag ?? DEFAULT_ROOT("comments");
    const paraIds = new Map<string, string>();
    const body = ids.sort((a, b) => Number(a.slice(4)) - Number(b.slice(4))).map((id) => {
      const m = meta.get(id);
      const paras = this.state.stories!.get(id)!;
      const pid = (this.paraIdSeq++).toString(16).toUpperCase();
      paraIds.set(id, pid);
      const attrs = (i: number) => (i === paras.length - 1 ? ` w14:paraId="${pid}"` : "");
      return `<w:comment w:id="${esc(id.slice(4))}" w:author="${esc(m?.author ?? "")}"${m?.date ? ` w:date="${esc(m.date)}"` : ""} w:initials="${esc(m?.initials ?? "")}">` +
        `${this.blocks(paras, path, attrs)}</w:comment>`;
    });
    const rootWith14 = root.includes("xmlns:w14=") ? root : root.replace(/>$/, ` xmlns:w14="${W14}">`);
    this.pw.setPart(path, `${XML_HEAD}${withMath(rootWith14, body.join(""))}${body.join("")}</w:comments>`, CT.comments);
    this.pw.link("word/document.xml", REL.comments, path);
    // Replies and resolved comments (Word 2013+): commentsExtended.xml.
    if ([...meta.values()].some((m) => m.parent || m.done)) {
      const ex = ids.map((id) => {
        const m = meta.get(id);
        const parent = m?.parent ? paraIds.get(m.parent) : undefined;
        return `<w15:commentEx w15:paraId="${paraIds.get(id)}"${parent ? ` w15:paraIdParent="${parent}"` : ""} w15:done="${m?.done ? 1 : 0}"/>`;
      });
      const exPath = "word/commentsExtended.xml";
      this.pw.setPart(exPath, `${XML_HEAD}<w15:commentsEx xmlns:w15="${W15}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">${ex.join("")}</w15:commentsEx>`, CT.commentsExtended);
      this.pw.link("word/document.xml", REL.commentsExtended, exPath);
    }
  }

  writeNumbering(): void {
    const n = this.state.numbering;
    if (!n || (this.pkg && this.pkg.numbering === n)) return;
    if (!n.abstracts.size && !this.pkg?.numbering) return;
    const rel = this.pw.relsOf("word/document.xml").list.find((r) => r.type === REL.numbering);
    const path = rel ? `word/${rel.target.replace(/^\//, "").replace(/^word\//, "")}` : "word/numbering.xml";
    this.pw.setPart(path, writeNumbering(n, this.pkg?.numberingRoot ?? null, this.opts), CT.numbering);
    this.pw.link("word/document.xml", REL.numbering, path);
  }
}

function marginsXml(tag: string, m: Partial<{ left: number; right: number; top: number; bottom: number }>): string {
  const one = (n: string, v: number | undefined) => (v === undefined ? "" : `<w:${n} w:w="${tw(v)}" w:type="dxa"/>`);
  return `<w:${tag}>${one("top", m.top)}${one("left", m.left)}${one("bottom", m.bottom)}${one("right", m.right)}</w:${tag}>`;
}

// ----- styles -----------------------------------------------------------------------------

export function styleXml(s: StyleDef, opts: WriteOptions): string {
  if (s.data && typeof s.data === "string") return s.data;
  const parts = [`<w:name w:val="${esc(s.name)}"/>`];
  if (s.basedOn) parts.push(`<w:basedOn w:val="${esc(s.basedOn)}"/>`);
  if (s.next) parts.push(`<w:next w:val="${esc(s.next)}"/>`);
  if (s.link) parts.push(`<w:link w:val="${esc(s.link)}"/>`);
  if (s.uiPriority !== undefined) parts.push(`<w:uiPriority w:val="${s.uiPriority}"/>`);
  if (s.semiHidden) parts.push("<w:semiHidden/><w:unhideWhenUsed/>");
  if (s.qFormat) parts.push("<w:qFormat/>");
  if (s.type === "paragraph" || s.type === "table") parts.push(writePPr(s.pPr, {}, opts));
  parts.push(writeRPr(s.rPr, opts));
  if (s.type === "table" && s.table) {
    const t = s.table;
    const tblPr = (t.borders ? bordersXml("tblBorders", t.borders, "table") : "") + (t.cellMargins ? marginsXml("tblCellMar", t.cellMargins) : "");
    if (tblPr) parts.push(`<w:tblPr>${tblPr}</w:tblPr>`);
    for (const [type, part] of Object.entries(t.cond ?? {})) {
      parts.push(`<w:tblStylePr w:type="${esc(type)}">${part.pPr ? writePPr(part.pPr, {}, opts) : ""}${part.rPr ? writeRPr(part.rPr, opts) : ""}` +
        `${part.borders || part.shading ? `<w:tcPr>${part.borders ? bordersXml("tcBorders", part.borders, "table") : ""}${shadingXml(part.shading)}</w:tcPr>` : ""}</w:tblStylePr>`);
    }
  }
  return `<w:style w:type="${s.type}"${s.isDefault ? ' w:default="1"' : ""} w:styleId="${esc(s.id)}">${parts.join("")}</w:style>`;
}

function stylesXml(state: DocState, opts: WriteOptions): string {
  const st = state.styles;
  const defs = [...st.styles.values()].map((s) => styleXml(s, opts)).join("");
  return `${XML_HEAD}<w:styles xmlns:w="${W}"><w:docDefaults>` +
    `<w:rPrDefault>${writeRPr(st.defaultsR, opts) || "<w:rPr/>"}</w:rPrDefault>` +
    `<w:pPrDefault>${writePPr(st.defaultsP, {}, opts) || "<w:pPr/>"}</w:pPrDefault></w:docDefaults>${defs}</w:styles>`;
}

/** The original styles.xml plus definitions for styles the document now uses but lacked. */
function mergedStyles(state: DocState, pkg: DocxPackage, opts: WriteOptions): string | null {
  if (!pkg.stylesXml) return null;
  const added = [...state.styles.styles.values()].filter((s) => !pkg.styleIds.has(s.id));
  if (!added.length) return pkg.stylesXml;
  const i = pkg.stylesXml.lastIndexOf("</w:styles>");
  if (i < 0) return pkg.stylesXml;
  return pkg.stylesXml.slice(0, i) + added.map((s) => styleXml(s, opts)).join("") + pkg.stylesXml.slice(i);
}

// ----- settings -----------------------------------------------------------------------------

/** Elements that precede w:evenAndOddHeaders / w:trackRevisions in CT_Settings. */
const BEFORE_EVEN_ODD = ["writeProtection", "view", "zoom", "removePersonalInformation", "removeDateAndTime", "doNotDisplayPageBoundaries",
  "displayBackgroundShape", "printPostScriptOverText", "printFractionalCharacterWidth", "printFormsData", "embedTrueTypeFonts",
  "embedSystemFonts", "saveSubsetFonts", "saveFormsData", "mirrorMargins", "alignBordersAndEdges", "bordersDoNotSurroundHeader",
  "bordersDoNotSurroundFooter", "gutterAtTop", "hideSpellingErrors", "hideGrammaticalErrors", "activeWritingStyle", "proofState",
  "formsDesign", "attachedTemplate", "linkStyles", "stylePaneFormatFilter", "stylePaneSortMethod", "documentType", "mailMerge",
  "revisionView", "trackRevisions", "doNotTrackMoves", "doNotTrackFormatting", "documentProtection", "autoFormatOverride",
  "styleLockTheme", "styleLockQFSet", "defaultTabStop", "autoHyphenation", "consecutiveHyphenLimit", "hyphenationZone",
  "doNotHyphenateCaps", "showEnvelope", "summaryLength", "clickAndTypeStyle", "defaultTableStyle"];
const BEFORE_TRACK = BEFORE_EVEN_ODD.slice(0, BEFORE_EVEN_ODD.indexOf("trackRevisions"));

/** Sets or removes an on/off element in settings.xml, at its place in the schema. */
function setFlag(xml: string, name: string, on: boolean, before: string[]): string {
  const re = new RegExp(`<w:${name}(\\s[^>]*)?/>`);
  if (re.test(xml) === on) return xml;
  if (!on) return xml.replace(re, "");
  let pos = -1;
  for (const b of before) {
    const m = new RegExp(`<w:${b}[\\s/>][\\s\\S]*?(?:</w:${b}>|/>)`).exec(xml);
    if (m) pos = Math.max(pos, m.index + m[0].length);
  }
  if (pos < 0) pos = xml.indexOf(">", xml.indexOf("<w:settings")) + 1;
  return `${xml.slice(0, pos)}<w:${name}/>${xml.slice(pos)}`;
}

function settingsXml(state: DocState, base: string | null): string {
  let xml = base ?? `${XML_HEAD}<w:settings xmlns:w="${W}"><w:defaultTabStop w:val="${tw(state.settings.defaultTab)}"/><w:hyphenationZone w:val="425"/>` +
    '<w:characterSpacingControl w:val="doNotCompress"/><w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>';
  xml = setFlag(xml, "evenAndOddHeaders", !!state.settings.evenOdd, BEFORE_EVEN_ODD);
  xml = setFlag(xml, "trackRevisions", !!state.settings.trackChanges, BEFORE_TRACK);
  return xml;
}

// ----- the package --------------------------------------------------------------------------

const CONTENT_TYPES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>`;

const ROOT_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>`;

const DOCUMENT_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>
</Relationships>`;

function coreXml(author: string, now: Date): string {
  const t = now.toISOString().replace(/\.\d+Z$/, "Z");
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
<dc:creator>${esc(author)}</dc:creator><cp:lastModifiedBy>${esc(author)}</cp:lastModifiedBy>
<dcterms:created xsi:type="dcterms:W3CDTF">${t}</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">${t}</dcterms:modified>
</cp:coreProperties>`;
}

/** Updates the modification date (and last author) in an existing core.xml. */
function touchCore(xml: string, author: string, now: Date): string {
  const t = now.toISOString().replace(/\.\d+Z$/, "Z");
  let out = xml.replace(/(<dcterms:modified[^>]*>)[^<]*(<\/dcterms:modified>)/, `$1${t}$2`);
  if (author) out = out.replace(/(<cp:lastModifiedBy>)[^<]*(<\/cp:lastModifiedBy>)/, `$1${esc(author)}$2`);
  return out;
}

const APP_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>RaabeWrite</Application></Properties>`;

export function writeDocx(state: DocState, pkg: DocxPackage | null, options: { author?: string; now?: Date } = {}): Uint8Array {
  const now = options.now ?? new Date();
  const author = options.author ?? "";
  const opts: WriteOptions = pkg
    ? { hasTheme: pkg.hasTheme, themeFonts: pkg.hasTheme ? undefined : { major: state.styles.theme.majorFont, minor: state.styles.theme.minorFont } }
    // New documents carry no theme part, so fonts are written by name.
    : { hasTheme: false, themeFonts: { major: state.styles.theme.majorFont, minor: state.styles.theme.minorFont } };
  const base: Record<string, Uint8Array> = pkg ? { ...pkg.files } : {
    "[Content_Types].xml": strToU8(CONTENT_TYPES),
    "_rels/.rels": strToU8(ROOT_RELS),
    "docProps/core.xml": strToU8(coreXml(author, now)),
    "docProps/app.xml": strToU8(APP_XML),
    "word/_rels/document.xml.rels": strToU8(DOCUMENT_RELS),
  };
  const pw = new PackageWriter(base);
  const w = new DocWriter(pw, state, pkg, opts);
  const body = w.blocks(state.paras, "word/document.xml");
  const trailing = (pkg?.trailing ?? []).join("");
  const sect = w.section(state.settings.section, "word/document.xml");
  pw.setPart("word/document.xml", `${XML_HEAD}${withMath(pkg?.rootTag ?? DEFAULT_ROOT("document"), body)}<w:body>${body}${trailing}${sect}</w:body></w:document>`);
  w.writeHeaders();
  w.writeNotes("fn");
  w.writeNotes("en");
  w.writeComments();
  w.writeNumbering();
  if (pkg) {
    const styles = mergedStyles(state, pkg, opts);
    if (styles && styles !== pkg.stylesXml) pw.setPart("word/styles.xml", styles);
    const settings = settingsXml(state, pkg.settingsXml);
    if (settings !== pkg.settingsXml) pw.setPart("word/settings.xml", settings);
    if (pw.files["docProps/core.xml"]) pw.files["docProps/core.xml"] = strToU8(touchCore(strFromU8(pw.files["docProps/core.xml"]), author, now));
  } else {
    pw.setPart("word/styles.xml", stylesXml(state, opts));
    pw.setPart("word/settings.xml", settingsXml(state, null));
  }
  return zipSync(pw.finish(), { level: 6 });
}
