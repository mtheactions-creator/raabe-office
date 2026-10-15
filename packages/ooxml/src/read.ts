// Reading .docx into the text engine's model: the main text, headers and footers, foot- and
// endnotes and comments as stories; sections, tables, lists, pictures, fields and tracked
// changes as the model describes them. Everything RaabeWrite does not interpret is kept as
// its original XML in place, so writing back loses nothing (D-004).

import { readExtent, readPicture, readShapeProps, themeColors } from "./drawingml";
import { NS_WPS, custGeomDef, formatMatrix, readAdjust, readBody, resolveStyle, type FormatMatrix } from "./shapes";
import { strFromU8, unzipSync } from "fflate";
import {
  A4, COLUMN_BREAK, DEFAULT_SECTION, LINE_BREAK, OBJ, ONE_COLUMN, PAGE_BREAK, StyleSheet, defaultStyleSheet, intern, nextTableId, plainPicture, seqToLatex,
  type Anchor, type BlockObject, type CellDef, type CellRef, type CharProps, type CommentMeta, type DocState, type FloatObject,
  type AltText, type DocSettings, type GroupMember, type HeaderRefs, type Watermark, type ImageData, type InlineObject, type MathSettings, type NoteProps, type Numbering, type Para, type ParaGroup, type ParaProps, type PictureFormat, type Revision, type ShapeInfo,
  type RowDef, type Run, type Section, type StyleDef, type StyleType, type TableDef, type TableStyleDef, type TableStylePart,
} from "@raabe/text-engine";
import { mathParaJc, ommlToSeq, readMathPr } from "./math";
import { readSources, type Bibliography } from "@raabe/text-engine";
import { readNumbering } from "./numbering";
import { REL, Rels, mimeOf, relsPath, resolveTarget } from "./package";
import { readBorders, readPPr, readRPr, readShading, type RevisionOf } from "./props";
import { attr, child, kids, num, onOff, outer, parse, textOf, W } from "./xml";

const WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
const A = "http://schemas.openxmlformats.org/drawingml/2006/main";
const NS_WPG = "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup";
const RNS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const W14 = "http://schemas.microsoft.com/office/word/2010/wordml";
const W15 = "http://schemas.microsoft.com/office/word/2012/wordml";
const W16SE = "http://schemas.microsoft.com/office/word/2015/wordml/symex";

/** Everything needed to write the file back: original parts and wrappers. */
export interface DocxPackage {
  files: Record<string, Uint8Array>;
  /** Start tag of w:document with all namespace declarations. */
  rootTag: string;
  /** Opaque body content after the last paragraph (e.g. trailing bookmarks). */
  trailing: string[];
  stylesXml: string | null;
  styleIds: Set<string>;
  hasTheme: boolean;
  /** The theme and the document defaults as read (JSON): changed ones are written. */
  themeRead?: string;
  defaultsRead?: string;
  /** The watermark as read (JSON) and w:background as read: headers and the background are rewritten when they change. */
  watermarkRead?: string;
  backgroundXml?: string | null;
  /** Header/footer stories: their part and the part's root start tag. */
  storyParts: Map<string, { path: string; rootTag: string }>;
  /** Numbering as read (written back unchanged while it is the same object). */
  numbering: Numbering | null;
  numberingRoot: string | null;
  /** Foot-/endnote parts: separators and other special notes, kept as XML. */
  notes: Partial<Record<"footnotes" | "endnotes", { path: string; rootTag: string; special: string[] }>>;
  comments: { path: string; rootTag: string } | null;
  settingsXml: string | null;
  /** Stories and comment data as read: parts whose stories are unchanged are kept as they are. */
  originalStories: Map<string, Para[]>;
  originalComments: ReadonlyMap<string, CommentMeta> | null;
  /** Word's sources part (customXml/itemN.xml with b:Sources) and what was read from it. */
  bibliography?: { path: string; read: Bibliography };
}

export interface ReadResult {
  state: DocState;
  pkg: DocxPackage;
  /** German names of content RaabeWrite shows only partly (it is still saved unchanged). */
  partial: string[];
}

const LABEL = {
  shapes: "Textfelder und Formen",
  math: "Formeln",
  controls: "Inhaltssteuerelemente",
  nested: "verschachtelte Tabellen",
  charts: "Diagramme und eingebettete Objekte",
};

const twips = (v: string | null, fallback: number) => (num(v) ?? fallback * 20) / 20;
const EMU = 12700;

/** Preserved inline content: complete XML, run content to wrap in a run, or a picture. */
export type ObjectData = { raw: string } | { inner: string } | { drawing: string; part: string; picRead?: string };

function themeFonts(xml: string | null): { major: string; minor: string; name?: string; colorName?: string; fontName?: string } {
  if (!xml) return { major: "Calibri Light", minor: "Calibri" };
  const doc = parse(xml);
  const pick = (tag: string, fallback: string) => {
    const el = doc.getElementsByTagNameNS(A, tag)[0];
    const latin = el ? Array.from(el.children).find((c) => c.localName === "latin") : undefined;
    return latin?.getAttribute("typeface") || fallback;
  };
  // The names Entwurf shows for the theme, its colours and its fonts.
  const nameOf = (tag: string) => doc.getElementsByTagNameNS(A, tag)[0]?.getAttribute("name") || undefined;
  return { major: pick("majorFont", "Calibri Light"), minor: pick("minorFont", "Calibri"), name: nameOf("theme"), colorName: nameOf("clrScheme"), fontName: nameOf("fontScheme") };
}

function cellMargins(el: Element | null): Partial<{ left: number; right: number; top: number; bottom: number }> | undefined {
  if (!el) return undefined;
  const w = (n: string, alt?: string) => {
    const c = child(el, n) ?? (alt ? child(el, alt) : null);
    return c ? (num(attr(c, "w")) ?? 0) / 20 : undefined;
  };
  const m = { left: w("left", "start"), right: w("right", "end"), top: w("top"), bottom: w("bottom") };
  return Object.fromEntries(Object.entries(m).filter(([, v]) => v !== undefined));
}

function tableStylePart(el: Element): TableStylePart {
  const tcPr = child(el, "tcPr"), tblPr = child(el, "tblPr");
  return {
    pPr: child(el, "pPr") ? readPPr(child(el, "pPr")).props : undefined,
    rPr: child(el, "rPr") ? readRPr(child(el, "rPr")) : undefined,
    borders: readBorders(child(tcPr ?? el, "tcBorders")) ?? readBorders(child(tblPr ?? el, "tblBorders")),
    shading: readShading(child(tcPr ?? el, "shd")) ?? readShading(child(tblPr ?? el, "shd")),
  };
}

function readTableStyle(s: Element): TableStyleDef {
  const tblPr = child(s, "tblPr");
  const cond: Record<string, TableStylePart> = {};
  for (const c of kids(s).filter((e) => e.localName === "tblStylePr")) cond[attr(c, "type") ?? ""] = tableStylePart(c);
  const base = tableStylePart(s);
  return {
    pPr: undefined, rPr: undefined,
    borders: readBorders(child(tblPr ?? s, "tblBorders")) ?? base.borders,
    shading: base.shading,
    cellMargins: cellMargins(child(tblPr ?? s, "tblCellMar")),
    rowBand: num(attr(child(tblPr ?? s, "tblStyleRowBandSize"), "val")),
    colBand: num(attr(child(tblPr ?? s, "tblStyleColBandSize"), "val")),
    cond: Object.keys(cond).length ? cond : undefined,
  };
}

function readStyles(xml: string | null, theme: { major: string; minor: string; colors?: Record<string, string>; name?: string; colorName?: string; fontName?: string }): StyleSheet {
  if (!xml) return defaultStyleSheet();
  const doc = parse(xml);
  const root = doc.documentElement;
  const dd = child(root, "docDefaults");
  const rDefault = dd ? readRPr(child(child(dd, "rPrDefault") ?? dd, "rPr")) : {};
  const pDefault = dd ? readPPr(child(child(dd, "pPrDefault") ?? dd, "pPr")).props : {};
  // ECMA-376 defaults when the file does not set them.
  const defaultsR: CharProps = { size: 10, font: rDefault.fontTheme ? undefined : "Times New Roman", ...rDefault };
  const defaultsP: ParaProps = { after: 0, line: 1, lineRule: "auto", widow: true, ...pDefault };
  const styles = new Map<string, StyleDef>();
  for (const s of kids(root).filter((e) => e.localName === "style")) {
    const id = attr(s, "styleId") ?? "";
    const type = (attr(s, "type") ?? "paragraph") as StyleType;
    styles.set(id, {
      id,
      name: attr(child(s, "name"), "val") ?? id,
      type,
      basedOn: attr(child(s, "basedOn"), "val") ?? undefined,
      next: attr(child(s, "next"), "val") ?? undefined,
      link: attr(child(s, "link"), "val") ?? undefined,
      isDefault: attr(s, "default") === "1" || attr(s, "default") === "true",
      uiPriority: num(attr(child(s, "uiPriority"), "val")),
      qFormat: !!child(s, "qFormat"),
      semiHidden: !!child(s, "semiHidden"),
      pPr: readPPr(child(s, "pPr")).props,
      rPr: readRPr(child(s, "rPr")),
      data: outer(s),
      table: type === "table" ? readTableStyle(s) : undefined,
    });
  }
  const names = { ...(theme.name ? { name: theme.name } : {}), ...(theme.colorName ? { colorName: theme.colorName } : {}), ...(theme.fontName ? { fontName: theme.fontName } : {}) };
  return new StyleSheet(styles, defaultsR, defaultsP, { majorFont: theme.major, minorFont: theme.minor, colors: theme.colors, ...names });
}

// ----- fields -------------------------------------------------------------------------------

interface FieldCollect {
  depth: number;
  /** XML from the begin to the separator (inclusive). */
  code: string[];
  /** Elements of the result (after the separator). */
  result: Element[];
  xml: string[];
  instr: string;
  display: string;
  inResult: boolean;
  props: CharProps;
  /** Offset in the paragraph where the field began. */
  at: number;
  rev?: Revision;
}

/** Updates a field's nesting depth for the fldChar elements inside `el`. */
function trackField(el: Element, f: FieldCollect): void {
  const walk = (e: Element) => {
    for (const c of Array.from(e.children)) {
      if (c.localName === "fldChar") {
        const t = attr(c, "fldCharType");
        if (t === "begin") f.depth++;
        else if (t === "separate" && f.depth === 1) f.inResult = true;
        else if (t === "end") f.depth--;
      } else if ((c.localName === "instrText" || c.localName === "delInstrText") && f.depth === 1 && !f.inResult) {
        f.instr += c.textContent ?? "";
      } else if ((c.localName === "t" || c.localName === "delText") && f.inResult && f.depth >= 1) {
        f.display += c.textContent ?? "";
      } else if (c.localName === "tab" && f.inResult) {
        f.display += "\t";
      } else {
        walk(c);
      }
    }
  };
  walk(el);
}

function hasFldChar(el: Element, type: string): boolean {
  for (const c of Array.from(el.getElementsByTagNameNS(W, "fldChar"))) if (attr(c, "fldCharType") === type) return true;
  return false;
}

// ----- paragraphs ---------------------------------------------------------------------------

class ParaBuilder {
  text = "";
  runs: Run[] = [];
  objects: InlineObject[] = [];
  anchors: Anchor[] = [];
  floats: FloatObject[] = [];

  addText(s: string, props: CharProps): void {
    if (!s) return;
    this.text += s;
    const last = this.runs[this.runs.length - 1];
    if (last && last.props === props) this.runs[this.runs.length - 1] = { len: last.len + s.length, props };
    else this.runs.push({ len: s.length, props });
  }

  addObject(o: InlineObject, props: CharProps): void {
    this.objects.push(o);
    this.addText(OBJ, props);
  }

  addAnchor(a: Omit<Anchor, "at">): void {
    this.anchors.push({ ...a, at: this.text.length });
  }
}

const ANCHOR_ELEMENTS = new Set(["bookmarkStart", "bookmarkEnd", "commentRangeStart", "commentRangeEnd", "permStart", "permEnd",
  "moveFromRangeStart", "moveFromRangeEnd", "moveToRangeStart", "moveToRangeEnd", "customXmlInsRangeStart",
  "customXmlInsRangeEnd", "customXmlDelRangeStart", "customXmlDelRangeEnd"]);

function mathText(el: Element): string {
  let out = "";
  for (const t of Array.from(el.getElementsByTagName("*"))) if (t.localName === "t") out += t.textContent ?? "";
  return out || "[Formel]";
}

function tableBlock(tbl: Element, label: string): BlockObject {
  const rows = kids(tbl).filter((e) => e.localName === "tr");
  const lines = rows.slice(0, 8).map((r) => kids(r).filter((c) => c.localName === "tc").map((c) => textOf(c).trim()).join("  |  "));
  return { kind: "table", data: outer(tbl), label, lines, height: Math.max(18, rows.length * 15.5) };
}

/** Shared state while reading a package. */
class Reader {
  readonly partial = new Set<string>();
  /** The document theme's colours (for pictures and shapes). */
  themeColors: Record<string, string> = themeColors(null);
  /** The theme's format matrix, which shape styles refer to. */
  matrix: FormatMatrix | null = null;
  /** Formeloptionen of the document (m:mathPr). */
  mathSettings: MathSettings | undefined;
  /** Fuß- und Endnote settings (w:footnotePr, w:endnotePr in settings.xml). */
  noteSettings: DocSettings["notes"];
  /** The watermark found in a header (Word puts it into every header). */
  watermark: Watermark | undefined;
  /** Linked text boxes: the story of each chain (its first box's text). */
  readonly chainStories = new Map<number, string>();
  readonly stories = new Map<string, Para[]>();
  private textboxSeq = 0;

  /** A new text box story id ("tx:1" …). */
  nextTextbox(): string {
    return `tx:${++this.textboxSeq}`;
  }
  readonly storyParts = new Map<string, { path: string; rootTag: string }>();
  private readonly storyOfPath = new Map<string, string>();
  private readonly revisions = new Map<string, Revision>();
  private readonly images = new Map<string, ImageData>();
  private fieldSeq = 0;
  private groupSeq = 0;

  constructor(readonly files: Record<string, Uint8Array>) {}

  text(path: string): string | null {
    return this.files[path] ? strFromU8(this.files[path]) : null;
  }

  rels(part: string): Rels {
    return new Rels(this.text(relsPath(part)));
  }

  revisionOf: RevisionOf = (el, kind) => {
    const id = num(attr(el, "id")) ?? 0;
    const author = attr(el, "author") ?? "";
    const date = attr(el, "date") ?? "";
    const key = `${kind}:${id}:${author}:${date}`;
    let r = this.revisions.get(key);
    if (!r) { r = { kind, id, author, date }; this.revisions.set(key, r); }
    return r;
  };

  image(part: string, rels: Rels, id: string | null): ImageData | undefined {
    if (!id) return undefined;
    const rel = rels.get(id);
    if (!rel || rel.external) return undefined;
    const path = resolveTarget(part, rel.target);
    let img = this.images.get(path);
    if (!img && this.files[path]) {
      img = { bytes: this.files[path], type: mimeOf(path) };
      this.images.set(path, img);
    }
    return img;
  }

  nextField(): string {
    return `f${++this.fieldSeq}`;
  }

  nextGroup(): number {
    return ++this.groupSeq;
  }

  /** The story of a header or footer part (read once). */
  headerStory(part: string, kind: "hdr" | "ftr"): string | null {
    const hit = this.storyOfPath.get(part);
    if (hit) return hit;
    const xml = this.text(part);
    if (!xml) return null;
    const m = /(\d+)\.xml$/.exec(part);
    let n = m ? Number(m[1]) : 1;
    while (this.stories.has(`${kind}:${n}`)) n++;
    const id = `${kind}:${n}`;
    this.storyOfPath.set(part, id);
    const root = parse(xml).documentElement;
    const pr = new PartReader(this, part);
    // Word's watermark (an sdt from the Watermarks gallery) becomes the document's setting.
    const els = kids(root).filter((el) => {
      if (el.localName !== "sdt" || !/docPartGallery w:val="Watermarks"/.test(outer(child(el, "sdtPr") ?? el))) return true;
      this.watermark ??= readWatermark(el, (rid) => this.image(part, pr.rels, rid));
      return !this.watermark;
    });
    this.stories.set(id, ensurePara(pr.blocks(els)));
    this.storyParts.set(id, { path: part, rootTag: rootTagOf(root) });
    return id;
  }
}

function rootTagOf(root: Element): string {
  return outer(root.cloneNode(false) as Element).replace(/\/>\s*$/, ">");
}

function ensurePara(paras: Para[]): Para[] {
  if (!paras.length || paras[paras.length - 1].cell) paras.push({ text: "", runs: [], props: {}, mark: intern({}) });
  return paras;
}

/** Reads the content of one part (document, header, note …). */
class PartReader {
  readonly rels: Rels;
  private fld: FieldCollect | null = null;
  /** Fields whose result spans paragraphs and whose end has not come yet. */
  private multi: { id: string; instr: string }[] = [];
  /** The note story being read (for its own reference mark). */
  noteId: string | null = null;
  /** Body-level markers waiting for the next paragraph. */
  leading: string[] = [];
  /** The final section of the main text. */
  finalSect: Section | null = null;

  constructor(readonly r: Reader, readonly part: string) {
    this.rels = r.rels(part);
  }

  // ----- blocks ---------------------------------------------------------------------------

  blocks(els: Element[], cell?: CellRef, group?: ParaGroup): Para[] {
    const out: Para[] = [];
    const push = (p: Para) => {
      let q = p;
      if (this.leading.length && !cell) { q = { ...q, leading: this.leading }; this.leading = []; }
      if (cell) q = { ...q, cell };
      if (group) q = { ...q, group };
      out.push(q);
    };
    for (const el of els) {
      if (el.namespaceURI !== W) { this.leading.push(outer(el)); continue; }
      switch (el.localName) {
        case "p": push(this.paragraph(el)); break;
        case "tbl":
          if (cell) {
            this.r.partial.add(LABEL.nested);
            push(this.blockPara({ ...tableBlock(el, "Verschachtelte Tabelle"), kind: "table" }));
          } else {
            for (const p of this.table(el)) out.push(group ? { ...p, group } : p);
          }
          break;
        case "sdt": {
          const content = child(el, "sdtContent");
          const inner = content ? kids(content) : [];
          const simple = !cell && !group && inner.length > 0 && inner.every((c) => c.namespaceURI === W && ["p", "tbl", "bookmarkStart", "bookmarkEnd"].includes(c.localName))
            && inner.some((c) => c.localName === "p" || c.localName === "tbl");
          if (simple) {
            const pr = child(el, "sdtPr"), endPr = child(el, "sdtEndPr");
            const g: ParaGroup = {
              id: this.r.nextGroup(), kind: "sdt",
              open: `<w:sdt>${pr ? outer(pr) : ""}${endPr ? outer(endPr) : ""}<w:sdtContent>`, close: "</w:sdtContent></w:sdt>",
            };
            out.push(...this.blocks(inner, undefined, g));
          } else {
            this.r.partial.add(LABEL.controls);
            push(this.blockPara({ kind: "other", data: outer(el), label: "Inhaltssteuerelement", lines: [textOf(el)], height: 14.5 * Math.max(1, inner.length) }));
          }
          break;
        }
        case "sectPr":
          this.finalSect = this.section(el);
          break;
        case "customXml": case "altChunk":
          push(this.blockPara({ kind: "other", data: outer(el), label: "Eingebetteter Inhalt", lines: [textOf(el)], height: 18 }));
          break;
        default:
          // Body-level markers (bookmarks …) travel with the next paragraph.
          this.leading.push(outer(el));
      }
    }
    return out;
  }

  private blockPara(block: BlockObject): Para {
    return { text: OBJ, runs: [{ len: 1, props: intern({}) }], props: {}, mark: intern({}), block };
  }

  paragraph(p: Element): Para {
    const pPrEl = child(p, "pPr");
    const { props, mark, sect } = readPPr(pPrEl, this.r.revisionOf);
    const b = new ParaBuilder();
    const fieldAtStart = this.multi[0]?.id;
    this.inline(kids(p).filter((c) => c !== pPrEl), b);
    this.endParagraph(b);
    const field = fieldAtStart ?? this.multi[0]?.id ?? b.anchors.find((a) => a.field)?.field?.id;
    return {
      text: b.text,
      runs: b.runs,
      props,
      mark: intern(mark),
      objects: b.objects.length ? b.objects : undefined,
      anchors: b.anchors.length ? b.anchors : undefined,
      floats: b.floats.length ? b.floats : undefined,
      sect: sect ? this.section(sect) : undefined,
      field,
    };
  }

  // ----- inline content -------------------------------------------------------------------

  private inline(els: Element[], b: ParaBuilder, rev?: Revision): void {
    for (const el of els) {
      if (this.fld) { this.collect(el, b); continue; }
      if (el.namespaceURI !== W && el.localName !== "oMath" && el.localName !== "oMathPara") {
        b.addObject({ kind: "other", data: { raw: outer(el) }, display: textOf(el) }, intern({ rev }));
        continue;
      }
      const name = el.localName;
      switch (name) {
        case "r":
          if (hasFldChar(el, "begin")) { this.startField(el, b, rev); break; }
          if (hasFldChar(el, "end") && this.multi.length) {
            const f = this.multi.pop()!;
            b.addAnchor({ data: { raw: outer(el) }, field: { id: f.id, edge: "end" } });
            break;
          }
          this.run(el, b, rev);
          break;
        case "hyperlink": {
          const id = this.r.nextField();
          const open = outer(el.cloneNode(false) as Element).replace(/\/>\s*$/, ">");
          // Where it goes: an external relationship (web, file, mail) or a bookmark (w:anchor).
          const rid = el.getAttributeNS(RNS, "id");
          const href = rid ? this.rels.list.find((r) => r.id === rid)?.target : undefined;
          b.addAnchor({ data: { raw: open }, link: { id, edge: "start", href, anchor: attr(el, "anchor") ?? undefined, tip: attr(el, "tooltip") ?? undefined, frame: attr(el, "tgtFrame") ?? undefined } });
          this.inline(kids(el), b, rev);
          b.addAnchor({ data: { raw: "</w:hyperlink>" }, link: { id, edge: "end" } });
          break;
        }
        case "fldSimple":
          b.addObject({ kind: "field", data: { raw: outer(el) }, display: textOf(el), instr: attr(el, "instr") ?? "" }, this.firstRunProps(el, rev));
          break;
        case "ins": case "moveTo":
          this.inline(kids(el), b, this.r.revisionOf(el, "ins"));
          break;
        case "del": case "moveFrom":
          this.inline(kids(el), b, this.r.revisionOf(el, "del"));
          break;
        case "smartTag": case "customXml":
          this.inline(kids(el).filter((c) => c.localName !== "smartTagPr" && c.localName !== "customXmlPr"), b, rev);
          break;
        case "oMath": case "oMathPara": {
          // Editable as LaTeX when every part converts; the XML is written back unchanged
          // as long as the equation is not edited.
          const math = ommlToSeq(el, this.r.mathSettings ?? {}) ?? undefined;
          const latex = math ? seqToLatex(math) : undefined;
          const mathJc = mathParaJc(el);
          // An m:oMath (not in m:oMathPara) is shown inline even alone in its paragraph.
          const mathInline = el.localName === "oMath" || undefined;
          b.addObject({ kind: "math", data: { raw: outer(el) }, display: mathText(el), latex, math, mathJc, mathInline }, intern({ ...this.firstRunProps(el, rev), italic: true }));
          if (latex === undefined) this.r.partial.add(LABEL.math);
          break;
        }
          break;
        case "sdt": {
          const content = child(el, "sdtContent");
          b.addObject({ kind: "other", data: { raw: outer(el) }, display: content ? textOf(content) : "" }, this.firstRunProps(el, rev));
          this.r.partial.add(LABEL.controls);
          break;
        }
        case "proofErr": break;
        case "bookmarkStart":
          b.addAnchor({ data: { raw: outer(el) }, bookmark: attr(el, "name") ?? undefined });
          break;
        case "commentRangeStart": case "commentRangeEnd":
          b.addAnchor({ data: { raw: outer(el) }, comment: { id: `cmt:${attr(el, "id") ?? "0"}`, edge: name === "commentRangeStart" ? "start" : "end" } });
          break;
        default:
          if (ANCHOR_ELEMENTS.has(name)) b.addAnchor({ data: { raw: outer(el) } });
          else b.addObject({ kind: "other", data: { raw: outer(el) }, display: textOf(el) }, this.firstRunProps(el, rev));
      }
    }
  }

  private firstRunProps(el: Element, rev?: Revision): CharProps {
    const r = el.getElementsByTagNameNS(W, "r")[0];
    return intern({ ...readRPr(r ? child(r, "rPr") : null), rev });
  }

  private startField(el: Element, b: ParaBuilder, rev?: Revision): void {
    this.fld = {
      depth: 0, code: [], result: [], xml: [], instr: "", display: "", inResult: false,
      props: intern({ ...readRPr(child(el, "rPr")), rev }), at: b.text.length, rev,
    };
    this.collect(el, b);
  }

  private collect(el: Element, b: ParaBuilder): void {
    const f = this.fld!;
    if (f.inResult) f.result.push(el); else f.code.push(outer(el));
    f.xml.push(outer(el));
    trackField(el, f);
    if (f.depth <= 0) {
      this.fld = null;
      b.addObject({ kind: "field", data: { raw: f.xml.join("") }, display: f.display, instr: f.instr.trim() }, f.props);
    }
  }

  /**
   * A field still open at the end of a paragraph spans paragraphs (a table of contents): its
   * code becomes a start marker and its result is read as ordinary, editable content.
   */
  private endParagraph(b: ParaBuilder): void {
    for (let guard = 0; this.fld && guard < 8; guard++) {
      const f = this.fld;
      this.fld = null;
      if (!f.inResult) {
        // Field code continuing into the next paragraph: kept as it is.
        b.addObject({ kind: "other", data: { raw: f.xml.join("") }, display: "" }, f.props);
        return;
      }
      const id = this.r.nextField();
      const anchor: Anchor = { at: f.at, data: { raw: f.code.join("") }, field: { id, edge: "start", instr: f.instr.trim() } };
      b.anchors.push(anchor);
      this.multi.push({ id, instr: f.instr.trim() });
      this.inline(f.result, b, f.rev);
    }
  }

  private run(r: Element, b: ParaBuilder, rev?: Revision): void {
    const rPrEl = child(r, "rPr");
    const props = intern({ ...readRPr(rPrEl), rev });
    const rPrXml = rPrEl ? outer(rPrEl) : "";
    let pending = "";
    const flush = () => { b.addText(pending, props); pending = ""; };
    const obj = (o: InlineObject) => { flush(); b.addObject(o, props); };
    for (const c of kids(r)) {
      switch (c.localName) {
        case "rPr": case "lastRenderedPageBreak": break;
        case "t": case "delText": pending += c.textContent ?? ""; break;
        case "tab": pending += "\t"; break;
        case "cr": pending += LINE_BREAK; break;
        case "noBreakHyphen": pending += "\u2011"; break;
        case "softHyphen": pending += "\u00ad"; break;
        case "br": {
          const type = attr(c, "type");
          pending += type === "page" ? PAGE_BREAK : type === "column" ? COLUMN_BREAK : LINE_BREAK;
          break;
        }
        case "footnoteReference": case "endnoteReference": {
          const kind = c.localName === "footnoteReference" ? "fn" : "en";
          obj({ kind: "note", data: { inner: outer(c) }, display: "", note: `${kind}:${attr(c, "id") ?? "0"}` });
          break;
        }
        case "footnoteRef": case "endnoteRef":
          obj({ kind: "note", data: { inner: outer(c) }, display: "", note: this.noteId ?? undefined });
          break;
        case "commentReference":
          flush();
          b.addAnchor({ data: { raw: `<w:r>${rPrXml}${outer(c)}</w:r>` }, comment: { id: `cmt:${attr(c, "id") ?? "0"}`, edge: "ref" } });
          break;
        case "annotationRef":
          flush();
          b.addAnchor({ data: { raw: `<w:r>${rPrXml}${outer(c)}</w:r>` } });
          break;
        case "AlternateContent": {
          // Emoji as Word writes them (w16se:symEx, the character as fallback text).
          const sym = c.getElementsByTagNameNS(W16SE, "symEx")[0];
          if (sym) {
            const code = parseInt(sym.getAttributeNS(W16SE, "char") ?? "", 16);
            if (Number.isFinite(code)) {
              flush();
              const font = sym.getAttributeNS(W16SE, "font") || undefined;
              b.addText(String.fromCodePoint(code), font ? intern({ ...props, font, fontTheme: undefined }) : props);
              break;
            }
          }
          flush();
          this.drawing(c, b, props);
          break;
        }
        case "drawing":
          flush();
          this.drawing(c, b, props);
          break;
        case "pict": case "object":
          flush();
          b.addAnchor({ data: { raw: `<w:r>${rPrXml}${outer(c)}</w:r>` } });
          this.r.partial.add(c.localName === "object" ? LABEL.charts : LABEL.shapes);
          break;
        case "sym": {
          const code = parseInt(attr(c, "char") ?? "", 16);
          // Symbol fonts map their characters to U+F000-U+F0FF; the font is the sym's own.
          const ch = Number.isFinite(code) ? String.fromCharCode(code < 0x100 ? code + 0xf000 : code) : "?";
          obj({ kind: "other", data: { inner: outer(c) }, display: ch, font: attr(c, "font") ?? undefined });
          break;
        }
        case "ptab":
          obj({ kind: "other", data: { inner: outer(c) }, display: " " });
          break;
        default:
          obj({ kind: "other", data: { inner: outer(c) }, display: textOf(c) });
      }
    }
    flush();
  }

  /** A shape or text box (wps:wsp): its look, its style, its text frame and its text. */
  private shapeOf(wsp: Element): { pic: PictureFormat | undefined; shape: ShapeInfo } {
    const theme = this.r.themeColors;
    const sp = kids(wsp).find((x) => x.localName === "spPr");
    const styleEl = kids(wsp).find((x) => x.localName === "style");
    this.r.matrix ??= formatMatrix(null);
    const styled = resolveStyle(styleEl, theme, this.r.matrix);
    const read = sp ? readShapeProps(sp, theme, styled.styleLine) : {};
    const pic = Object.values(read).some((v) => v !== undefined) ? read : undefined;
    const cnv = kids(wsp).find((x) => x.localName === "cNvSpPr" || x.localName === "cNvCnPr");
    const kind = cnv?.localName === "cNvCnPr" ? "connector" : cnv?.getAttribute("txBox") === "1" ? "textbox" : "shape";
    const prst = sp ? kids(sp).find((x) => x.localName === "prstGeom") : undefined;
    const cust = sp ? kids(sp).find((x) => x.localName === "custGeom") : undefined;
    const txbx = kids(wsp).find((x) => x.localName === "txbx");
    const content = txbx ? kids(txbx).find((x) => x.localName === "txbxContent") : undefined;
    let story: string | undefined;
    if (content) {
      story = this.r.nextTextbox();
      this.r.stories.set(story, ensurePara(new PartReader(this.r, this.part).blocks(kids(content))));
    }
    // Verknüpfte Textfelder: the first box's txbx has the chain's id, the others wps:linkedTxbx.
    const linked = kids(wsp).find((x) => x.localName === "linkedTxbx");
    const chainId = txbx?.getAttribute("id") ?? linked?.getAttribute("id");
    const chain = chainId !== null && chainId !== undefined && chainId !== "" ? { id: Number(chainId), seq: linked ? Number(linked.getAttribute("seq") ?? 1) : 0 } : undefined;
    if (chain && story && chain.seq === 0) this.r.chainStories.set(chain.id, story);
    const shape: ShapeInfo = {
      kind, adjust: readAdjust(prst), ...styled, chain,
      story, body: readBody(kids(wsp).find((x) => x.localName === "bodyPr"), kind === "textbox"),
      styleXml: styleEl ? outer(styleEl) : undefined, custom: cust ? custGeomDef(cust) : undefined,
    };
    return { pic, shape: Object.fromEntries(Object.entries(shape).filter(([, v]) => v !== undefined)) as unknown as ShapeInfo };
  }

  /** A picture or shape: inline (an object in the text) or anchored (positioned on the page). */
  private drawing(c: Element, b: ParaBuilder, props: CharProps): void {
    const inline = c.getElementsByTagNameNS(WP, "inline")[0];
    const anchor = c.getElementsByTagNameNS(WP, "anchor")[0];
    // A group (wpg:wgp) is one object with its shapes and pictures inside.
    const wgp = c.getElementsByTagNameNS(NS_WPG, "wgp")[0];
    if (wgp && anchor) {
      this.groupDrawing(c, wgp, b, anchor);
      return;
    }
    const wsp = wgp ? undefined : c.getElementsByTagNameNS(NS_WPS, "wsp")[0];
    if (wsp && (inline || anchor)) {
      this.shapeDrawing(c, wsp, b, props, inline, anchor);
      return;
    }
    const blip = c.getElementsByTagNameNS(A, "blip")[0];
    const image = blip ? this.r.image(this.part, this.rels, blip.getAttributeNS(RNS, "embed")) : undefined;
    // Bildformat settings and Alternativtext; the XML is only rewritten when they change.
    const picEl = c.getElementsByTagNameNS("http://schemas.openxmlformats.org/drawingml/2006/picture", "pic")[0];
    const holder = inline ?? anchor;
    const extentEl = holder ? Array.from(holder.children).find((x) => x.localName === "effectExtent") : undefined;
    const extent = readExtent(extentEl);
    const read = picEl ? { ...readPicture(picEl, this.r.themeColors), extent } : undefined;
    const pic = read && !plainPicture(read) ? read : undefined;
    const docPr = c.getElementsByTagNameNS(WP, "docPr")[0];
    const descr = docPr?.getAttribute("descr") ?? undefined, title = docPr?.getAttribute("title") ?? undefined;
    const decorative = !!docPr?.getElementsByTagNameNS("http://schemas.microsoft.com/office/drawing/2017/decorative", "decorative")[0];
    const alt = docPrAlt(docPr, { descr: descr || undefined, title: title || undefined, decorative: decorative || undefined });
    const data: ObjectData = { drawing: outer(c), part: this.part, picRead: JSON.stringify({ pic, alt }) };
    const ext = (el: Element) => {
      const e = Array.from(el.children).find((x) => x.localName === "extent");
      return { width: Number(e?.getAttribute("cx") ?? 0) / EMU, height: Number(e?.getAttribute("cy") ?? 0) / EMU };
    };
    if (!image && !c.getElementsByTagNameNS("http://schemas.openxmlformats.org/drawingml/2006/picture", "pic")[0]) {
      this.r.partial.add(c.getElementsByTagNameNS(A, "graphicData")[0]?.getAttribute("uri")?.includes("chart") ? LABEL.charts : LABEL.shapes);
    }
    if (inline || !anchor) {
      const size = inline ? ext(inline) : { width: 0, height: 0 };
      b.addObject({ kind: "image", data, display: "", width: size.width, height: size.height, image, pic, alt }, props);
      return;
    }
    const size = ext(anchor);
    const pos = (tag: string) => {
      const el = Array.from(anchor.children).find((x) => x.localName === tag);
      const align = el ? Array.from(el.children).find((x) => x.localName === "align")?.textContent ?? undefined : undefined;
      const off = el ? Array.from(el.children).find((x) => x.localName === "posOffset")?.textContent : undefined;
      return { rel: el?.getAttribute("relativeFrom") ?? "column", align, offset: Number(off ?? 0) / EMU };
    };
    const wrapEl = Array.from(anchor.children).find((x) => x.localName.startsWith("wrap"));
    const wrapName = wrapEl?.localName ?? "wrapNone";
    const behind = anchor.getAttribute("behindDoc") === "1";
    const wrap = wrapName === "wrapNone" ? (behind ? "behind" : "front")
      : wrapName === "wrapSquare" ? "square" : wrapName === "wrapTight" ? "tight" : wrapName === "wrapThrough" ? "through" : "topAndBottom";
    const d = (n: string) => Number(anchor.getAttribute(n) ?? 0) / EMU;
    b.floats.push({
      at: b.text.length, data, width: size.width, height: size.height, image,
      x: pos("positionH"), y: pos("positionV"), wrap, side: wrapEl?.getAttribute("wrapText") ?? "bothSides",
      dist: [d("distT"), d("distB"), d("distL"), d("distR")], polygon: wrapPolygon(wrapEl), pic, alt,
      z: Number(anchor.getAttribute("relativeHeight") ?? 0) || undefined,
    });
  }

  /**
   * Word's group: the members' places come from the group's child coordinates (a:chOff,
   * a:chExt) scaled into its box, nested groups included; each member keeps its XML.
   */
  private groupDrawing(c: Element, wgp: Element, b: ParaBuilder, a: Element): void {
    const members: GroupMember[] = [];
    const num = (el: Element | undefined, n: string) => Number(el?.getAttribute(n) ?? 0);
    const xf = (el: Element | undefined) => {
      const x = el ? Array.from(el.children).find((k) => k.localName === "xfrm") : undefined;
      const k = (n: string) => (x ? Array.from(x.children).find((e) => e.localName === n) : undefined);
      return { off: [num(k("off"), "x"), num(k("off"), "y")], ext: [num(k("ext"), "cx"), num(k("ext"), "cy")], chOff: [num(k("chOff"), "x"), num(k("chOff"), "y")], chExt: [num(k("chExt"), "cx"), num(k("chExt"), "cy")] };
    };
    type Map = (x: number, y: number, w: number, h: number) => [number, number, number, number];
    const walk = (grp: Element, outerMap: Map) => {
      const g = xf(kids(grp).find((k) => k.localName === "grpSpPr"));
      const sx = g.chExt[0] ? g.ext[0] / g.chExt[0] : 1, sy = g.chExt[1] ? g.ext[1] / g.chExt[1] : 1;
      const map: Map = (x, y, w, h) => outerMap(g.off[0] + (x - g.chOff[0]) * sx, g.off[1] + (y - g.chOff[1]) * sy, w * sx, h * sy);
      for (const k of kids(grp)) {
        if (k.localName === "grpSp") { walk(k, map); continue; }
        if (k.localName !== "wsp" && k.localName !== "pic") continue;
        const sp = kids(k).find((e) => e.localName === "spPr");
        const t = xf(sp);
        const [mx, my, mw, mh] = map(t.off[0], t.off[1], t.ext[0], t.ext[1]);
        const cnv = kids(k).find((e) => e.localName === "cNvPr") ?? (k.localName === "pic" ? kids(kids(k).find((e) => e.localName === "nvPicPr") ?? k).find((e) => e.localName === "cNvPr") : undefined);
        const name = cnv?.getAttribute("name") || undefined, descr = cnv?.getAttribute("descr") || undefined;
        const alt = name || descr ? { ...(name ? { name } : {}), ...(descr ? { descr } : {}) } : undefined;
        const pos = { x: mx / EMU, y: my / EMU, width: mw / EMU, height: mh / EMU };
        if (k.localName === "wsp") {
          const { pic, shape } = this.shapeOf(k);
          members.push({ ...pos, shape, pic, alt, data: { xml: outer(k), read: JSON.stringify({ pic, alt, shape: { ...shape, story: undefined } }) } });
        } else {
          const blip = k.getElementsByTagNameNS(A, "blip")[0];
          const image = blip ? this.r.image(this.part, this.rels, blip.getAttributeNS(RNS, "embed")) : undefined;
          const read = readPicture(k, this.r.themeColors);
          const p2 = plainPicture(read) ? undefined : read;
          members.push({ ...pos, image, pic: p2, alt, data: { xml: outer(k), read: JSON.stringify({ pic: p2, alt }) } });
        }
      }
    };
    // The root group's own offset is the object's box (0, 0 within it).
    const root = xf(kids(wgp).find((k) => k.localName === "grpSpPr"));
    walk(wgp, (x, y, w, h) => [x - root.off[0], y - root.off[1], w, h]);
    const docPr = c.getElementsByTagNameNS(WP, "docPr")[0];
    const alt = docPrAlt(docPr, { descr: docPr?.getAttribute("descr") || undefined, title: docPr?.getAttribute("title") || undefined });
    const snapshot = members.map(({ data: _d, shape, ...m }) => ({ ...m, shape: shape ? { ...shape, story: undefined } : undefined }));
    const data: ObjectData = { drawing: outer(c), part: this.part, picRead: JSON.stringify({ alt, members: snapshot }) };
    const e = Array.from(a.children).find((x) => x.localName === "extent");
    const pos = (tag: string) => {
      const el = Array.from(a.children).find((x) => x.localName === tag);
      const align = el ? Array.from(el.children).find((x) => x.localName === "align")?.textContent ?? undefined : undefined;
      const off = el ? Array.from(el.children).find((x) => x.localName === "posOffset")?.textContent : undefined;
      return { rel: el?.getAttribute("relativeFrom") ?? "column", align, offset: Number(off ?? 0) / EMU };
    };
    const wrapEl = Array.from(a.children).find((x) => x.localName.startsWith("wrap"));
    const wrapName = wrapEl?.localName ?? "wrapNone";
    const behind = a.getAttribute("behindDoc") === "1";
    const wrap = wrapName === "wrapNone" ? (behind ? "behind" : "front")
      : wrapName === "wrapSquare" ? "square" : wrapName === "wrapTight" ? "tight" : wrapName === "wrapThrough" ? "through" : "topAndBottom";
    const d = (n: string) => Number(a.getAttribute(n) ?? 0) / EMU;
    b.floats.push({
      at: b.text.length, data, width: Number(e?.getAttribute("cx") ?? 0) / EMU, height: Number(e?.getAttribute("cy") ?? 0) / EMU,
      x: pos("positionH"), y: pos("positionV"), wrap, side: wrapEl?.getAttribute("wrapText") ?? "bothSides",
      dist: [d("distT"), d("distB"), d("distL"), d("distR")], polygon: wrapPolygon(wrapEl), alt, z: Number(a.getAttribute("relativeHeight") ?? 0) || undefined,
      group: { members },
    });
  }

  private shapeDrawing(c: Element, wsp: Element, b: ParaBuilder, props: CharProps, inline: Element | undefined, anchor: Element | undefined): void {
    const read = this.shapeOf(wsp);
    const shape = read.shape;
    // The room of the outline and effects (wp:effectExtent): Word aligns shapes with it.
    const fx = readExtent(Array.from((inline ?? anchor!).children).find((x) => x.localName === "effectExtent"));
    const pic = fx && (fx.l || fx.t || fx.r || fx.b) ? { ...read.pic, extent: fx } : read.pic;
    const docPr = c.getElementsByTagNameNS(WP, "docPr")[0];
    const descr = docPr?.getAttribute("descr") ?? undefined, title = docPr?.getAttribute("title") ?? undefined;
    const alt = docPrAlt(docPr, { descr: descr || undefined, title: title || undefined });
    // The XML is written back unchanged unless the shape changes (snapshot without the story).
    const data: ObjectData = { drawing: outer(c), part: this.part, picRead: JSON.stringify({ pic, alt, shape: { ...shape, story: undefined, chain: shape.chain ? { id: shape.chain.id, seq: shape.chain.seq } : undefined } }) };
    const holder = inline ?? anchor!;
    const e = Array.from(holder.children).find((x) => x.localName === "extent");
    const size = { width: Number(e?.getAttribute("cx") ?? 0) / EMU, height: Number(e?.getAttribute("cy") ?? 0) / EMU };
    if (inline) {
      b.addObject({ kind: "image", data, display: "", width: size.width, height: size.height, pic, alt, shape }, props);
      return;
    }
    const a = anchor!;
    const pos = (tag: string) => {
      const el = Array.from(a.children).find((x) => x.localName === tag);
      const align = el ? Array.from(el.children).find((x) => x.localName === "align")?.textContent ?? undefined : undefined;
      const off = el ? Array.from(el.children).find((x) => x.localName === "posOffset")?.textContent : undefined;
      return { rel: el?.getAttribute("relativeFrom") ?? "column", align, offset: Number(off ?? 0) / EMU };
    };
    const wrapEl = Array.from(a.children).find((x) => x.localName.startsWith("wrap"));
    const wrapName = wrapEl?.localName ?? "wrapNone";
    const behind = a.getAttribute("behindDoc") === "1";
    const wrap = wrapName === "wrapNone" ? (behind ? "behind" : "front")
      : wrapName === "wrapSquare" ? "square" : wrapName === "wrapTight" ? "tight" : wrapName === "wrapThrough" ? "through" : "topAndBottom";
    const d = (n: string) => Number(a.getAttribute(n) ?? 0) / EMU;
    b.floats.push({
      at: b.text.length, data, width: size.width, height: size.height,
      x: pos("positionH"), y: pos("positionV"), wrap, side: wrapEl?.getAttribute("wrapText") ?? "bothSides",
      dist: [d("distT"), d("distB"), d("distL"), d("distR")], polygon: wrapPolygon(wrapEl), pic, alt, shape,
      z: Number(a.getAttribute("relativeHeight") ?? 0) || undefined,
    });
  }

  // ----- tables ---------------------------------------------------------------------------

  table(tbl: Element): Para[] {
    const tblPr = child(tbl, "tblPr");
    const extra: string[] = [];
    let style: string | undefined, align: TableDef["align"], indent: number | undefined;
    let look: TableDef["look"];
    let float: TableDef["float"], overlap: boolean | undefined;
    if (tblPr) {
      for (const c of kids(tblPr)) {
        switch (c.localName) {
          case "tblStyle": style = attr(c, "val") ?? undefined; break;
          case "jc": { const v = attr(c, "val"); align = v === "center" ? "center" : v === "right" || v === "end" ? "right" : "left"; break; }
          case "tblInd": indent = (num(attr(c, "w")) ?? 0) / 20; break;
          case "tblpPr": {
            const tw = (n: string) => (num(attr(c, n)) ?? 0) / 20;
            const xs = attr(c, "tblpXSpec"), ys = attr(c, "tblpYSpec");
            const ha = attr(c, "horzAnchor"), va = attr(c, "vertAnchor");
            float = {
              x: tw("tblpX"), y: tw("tblpY"),
              xAlign: xs === "left" || xs === "center" || xs === "right" || xs === "inside" || xs === "outside" ? xs : undefined,
              yAlign: ys === "top" || ys === "center" || ys === "bottom" || ys === "inside" || ys === "outside" || ys === "inline" ? ys : undefined,
              hAnchor: ha === "margin" || ha === "page" ? ha : "text", vAnchor: va === "text" || va === "page" ? va : "margin",
              left: tw("leftFromText"), right: tw("rightFromText"), top: tw("topFromText"), bottom: tw("bottomFromText"),
            };
            break;
          }
          case "tblOverlap": overlap = attr(c, "val") !== "never"; break;
          case "tblBorders": case "shd": case "tblCellMar": break;
          case "tblLook": {
            const hex = parseInt(attr(c, "val") ?? "0", 16) || 0;
            const f = (n: string, bit: number) => (attr(c, n) !== null ? attr(c, n) === "1" || attr(c, n) === "true" : !!(hex & bit));
            look = { firstRow: f("firstRow", 0x20), lastRow: f("lastRow", 0x40), firstColumn: f("firstColumn", 0x80), lastColumn: f("lastColumn", 0x100), noHBand: f("noHBand", 0x200), noVBand: f("noVBand", 0x400) };
            break;
          }
          default: extra.push(`${c.localName}\u0000${outer(c)}`);
        }
      }
    }
    const grid = kids(child(tbl, "tblGrid") ?? tbl).filter((g) => g.localName === "gridCol").map((g) => (num(attr(g, "w")) ?? 0) / 20);
    const rowEls = kids(tbl).filter((e) => e.localName === "tr");
    const cellEls: Element[][] = [];
    const rows: RowDef[] = rowEls.map((tr) => {
      const trPr = child(tr, "trPr");
      const rowExtra: string[] = [];
      let height: number | undefined, heightRule: RowDef["heightRule"], header = false, cantSplit = false;
      for (const c of trPr ? kids(trPr) : []) {
        if (c.localName === "trHeight") {
          height = (num(attr(c, "val")) ?? 0) / 20;
          heightRule = attr(c, "hRule") === "exact" ? "exact" : "atLeast";
        } else if (c.localName === "tblHeader") header = onOff(c) ?? false;
        else if (c.localName === "cantSplit") cantSplit = onOff(c) ?? false;
        else rowExtra.push(`${c.localName}\u0000${outer(c)}`);
      }
      const prEx = child(tr, "tblPrEx");
      if (prEx) rowExtra.push(`tblPrEx\u0000${outer(prEx)}`);
      const tcs = kids(tr).filter((e) => e.localName === "tc");
      cellEls.push(tcs);
      const cells: CellDef[] = tcs.map((tc) => {
        const tcPr = child(tc, "tcPr");
        const cellExtra: string[] = [];
        let span = 1, vMerge: CellDef["vMerge"], width: number | undefined, vAlign: CellDef["vAlign"];
        for (const c of tcPr ? kids(tcPr) : []) {
          switch (c.localName) {
            case "gridSpan": span = num(attr(c, "val")) ?? 1; break;
            case "vMerge": vMerge = attr(c, "val") === "restart" ? "restart" : "continue"; break;
            case "tcW": if (attr(c, "type") === "dxa" || !attr(c, "type")) width = (num(attr(c, "w")) ?? 0) / 20; cellExtra.push(`tcW\u0000${outer(c)}`); break;
            case "vAlign": { const v = attr(c, "val"); vAlign = v === "center" ? "center" : v === "bottom" ? "bottom" : "top"; break; }
            case "tcBorders": case "shd": case "tcMar": break;
            default: cellExtra.push(`${c.localName}\u0000${outer(c)}`);
          }
        }
        return {
          span, vMerge, width, vAlign,
          shading: readShading(child(tcPr ?? tc, "shd")),
          borders: readBorders(child(tcPr ?? tc, "tcBorders")),
          margins: cellMargins(child(tcPr ?? tc, "tcMar")),
          extra: cellExtra.length ? cellExtra : undefined,
        };
      });
      return { cells, height, heightRule, header: header || undefined, cantSplit: cantSplit || undefined, extra: rowExtra.length ? rowExtra : undefined };
    });
    const def: TableDef = {
      id: nextTableId(), grid, rows, style, align, indent,
      borders: readBorders(child(tblPr ?? tbl, "tblBorders")),
      shading: readShading(child(tblPr ?? tbl, "shd")),
      cellMargins: cellMargins(child(tblPr ?? tbl, "tblCellMar")),
      look, extra: extra.length ? extra : undefined,
      ...(float ? { float: overlap === false ? { ...float, overlap: false } : float } : {}),
    };
    const out: Para[] = [];
    cellEls.forEach((tcs, r) => tcs.forEach((tc, c) => {
      const cell: CellRef = { table: def, row: r, col: c };
      const paras = this.blocks(kids(tc).filter((e) => e.localName !== "tcPr"), cell);
      out.push(...(paras.length ? paras : [{ text: "", runs: [], props: {}, mark: intern({}), cell }]));
    }));
    return out;
  }

  // ----- sections -------------------------------------------------------------------------

  section(el: Element): Section {
    const extra: string[] = [];
    const headers: Record<string, string> = {}, footers: Record<string, string> = {};
    let page = { ...A4 };
    let columns = ONE_COLUMN;
    let start: Section["start"] = "nextPage";
    let titlePage = false;
    let pageStart: number | undefined, pageFormat: string | undefined;
    let lineNumbers: Section["lineNumbers"];
    let pageBorders: Section["pageBorders"];
    for (const c of kids(el)) {
      switch (c.localName) {
        case "headerReference": case "footerReference": {
          const rel = this.rels.get(c.getAttributeNS(RNS, "id") ?? "");
          const kind = c.localName === "headerReference" ? "hdr" : "ftr";
          const id = rel ? this.r.headerStory(resolveTarget(this.part, rel.target), kind) : null;
          if (id) (kind === "hdr" ? headers : footers)[attr(c, "type") ?? "default"] = id;
          break;
        }
        case "pgSz":
          page = { ...page, width: twips(attr(c, "w"), A4.width), height: twips(attr(c, "h"), A4.height) };
          if (attr(c, "orient")) extra.push(`pgSzOrient\u0000${attr(c, "orient")}`);
          if (attr(c, "code")) extra.push(`pgSzCode\u0000${attr(c, "code")}`);
          break;
        case "pgMar":
          page = {
            ...page,
            marginTop: Math.abs(twips(attr(c, "top"), A4.marginTop)), marginBottom: Math.abs(twips(attr(c, "bottom"), A4.marginBottom)),
            marginLeft: twips(attr(c, "left") ?? attr(c, "start"), A4.marginLeft), marginRight: twips(attr(c, "right") ?? attr(c, "end"), A4.marginRight),
            header: twips(attr(c, "header"), A4.header), footer: twips(attr(c, "footer"), A4.footer), gutter: twips(attr(c, "gutter"), 0),
          };
          break;
        case "cols": {
          const n = num(attr(c, "num")) ?? 1;
          const cols = kids(c).filter((x) => x.localName === "col").map((x) => ({ width: (num(attr(x, "w")) ?? 0) / 20, space: (num(attr(x, "space")) ?? 0) / 20 }));
          columns = {
            count: Math.max(1, n), space: (num(attr(c, "space")) ?? 720) / 20, sep: onOff(child(c, "sep")) ?? (attr(c, "sep") === "1" || attr(c, "sep") === "true"),
            cols: attr(c, "equalWidth") === "0" && cols.length ? cols : undefined,
          };
          break;
        }
        case "type": {
          const v = attr(c, "val");
          start = v === "continuous" || v === "evenPage" || v === "oddPage" || v === "nextColumn" ? v : "nextPage";
          break;
        }
        case "titlePg": titlePage = onOff(c) ?? false; break;
        case "pgBorders": {
          const b = readBorders(c);
          const disp = attr(c, "display");
          pageBorders = {
            top: b?.top, left: b?.left, bottom: b?.bottom, right: b?.right, offsetFrom: attr(c, "offsetFrom") === "page" ? "page" : "text",
            ...(disp === "firstPage" || disp === "notFirstPage" ? { display: disp } : {}),
          };
          if (attr(c, "zOrder")) extra.push(`pgBordersZ\u0000${attr(c, "zOrder")}`);
          break;
        }
        case "lnNumType": {
          const by = num(attr(c, "countBy")) ?? 0;
          const r = attr(c, "restart");
          if (by > 0) lineNumbers = { countBy: by, start: (num(attr(c, "start")) ?? 0) + 1, distance: attr(c, "distance") ? twips(attr(c, "distance"), 0) : undefined,
            restart: r === "newSection" || r === "continuous" ? r : "newPage" };
          else extra.push(`${c.localName}\u0000${outer(c)}`);
          break;
        }
        case "pgNumType":
          pageStart = num(attr(c, "start"));
          pageFormat = attr(c, "fmt") ?? undefined;
          if (attr(c, "chapStyle") || attr(c, "chapSep")) extra.push(`${c.localName}\u0000${outer(c)}`);
          break;
        default:
          extra.push(`${c.localName}\u0000${outer(c)}`);
      }
    }
    return {
      page, columns, start, titlePage, headers: headers as HeaderRefs, footers: footers as HeaderRefs, pageStart, pageFormat,
      ...(lineNumbers ? { lineNumbers } : {}),
      ...(pageBorders ? { pageBorders } : {}),
      extra: extra.length ? extra : undefined,
    };
  }
}

// ----- notes and comments -------------------------------------------------------------------

function readNotes(r: Reader, path: string, kind: "fn" | "en"): { rootTag: string; special: string[] } | null {
  const xml = r.text(path);
  if (!xml) return null;
  const root = parse(xml).documentElement;
  const special: string[] = [];
  const pr = new PartReader(r, path);
  for (const n of kids(root)) {
    const type = attr(n, "type");
    if (type && type !== "normal") { special.push(outer(n)); continue; }
    const id = `${kind}:${attr(n, "id") ?? "0"}`;
    pr.noteId = id;
    r.stories.set(id, ensurePara(pr.blocks(kids(n))));
  }
  return { rootTag: rootTagOf(root), special };
}

function readComments(r: Reader, path: string, extPath: string | null): { rootTag: string; meta: Map<string, CommentMeta> } | null {
  const xml = r.text(path);
  if (!xml) return null;
  const root = parse(xml).documentElement;
  const meta = new Map<string, CommentMeta>();
  const byPara = new Map<string, string>();
  const pr = new PartReader(r, path);
  for (const c of kids(root).filter((e) => e.localName === "comment")) {
    const id = `cmt:${attr(c, "id") ?? "0"}`;
    r.stories.set(id, ensurePara(pr.blocks(kids(c))));
    const ps = kids(c).filter((e) => e.localName === "p");
    const paraId = ps.length ? ps[ps.length - 1].getAttributeNS(W14, "paraId") : null;
    if (paraId) byPara.set(paraId, id);
    meta.set(id, { id, author: attr(c, "author") ?? "", initials: attr(c, "initials") ?? "", date: attr(c, "date") ?? "" });
  }
  const ext = extPath ? r.text(extPath) : null;
  if (ext) {
    for (const e of Array.from(parse(ext).getElementsByTagNameNS(W15, "commentEx"))) {
      const id = byPara.get(e.getAttributeNS(W15, "paraId") ?? "");
      if (!id) continue;
      const parent = byPara.get(e.getAttributeNS(W15, "paraIdParent") ?? "");
      const m = meta.get(id)!;
      meta.set(id, { ...m, parent, done: e.getAttributeNS(W15, "done") === "1" || undefined });
    }
  }
  return { rootTag: rootTagOf(root), meta };
}

// ----- the package --------------------------------------------------------------------------

/**
 * Word's watermark: VML WordArt (v:textpath, PowerPlusWaterMarkObject) or a picture
 * (v:imagedata, WordPictureWatermark) inside the header's Watermarks sdt.
 */
function readWatermark(sdt: Element, image: (rid: string) => ImageData | undefined): Watermark | undefined {
  const xml = outer(sdt);
  const shape = /<v:shape\b([^>]*)>([\s\S]*?)<\/v:shape>/.exec(xml);
  if (!shape) return undefined;
  const attrs = shape[1], inner = shape[2];
  const style = /style="([^"]*)"/.exec(attrs)?.[1] ?? "";
  const css = (k: string) => new RegExp(`(?:^|;)${k}:([^;]*)`).exec(style)?.[1];
  const pt = (v: string | undefined) => (v ? parseFloat(v) * (v.endsWith("in") ? 72 : v.endsWith("cm") ? 28.3465 : 1) : 0);
  const width = pt(css("width")), height = pt(css("height"));
  const textpath = /<v:textpath\b([^>]*)\/?>/.exec(inner);
  if (textpath) {
    const tp = textpath[1];
    const text = (/string="([^"]*)"/.exec(tp)?.[1] ?? "").replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&amp;/g, "&");
    const tstyle = (/style="([^"]*)"/.exec(tp)?.[1] ?? "").replace(/&quot;/g, '"');
    const font = /font-family:"?([^";]*)"?/.exec(tstyle)?.[1] ?? "Calibri";
    const fsize = parseFloat(/font-size:([\d.]+)pt/.exec(tstyle)?.[1] ?? "1");
    const fill = /fillcolor="([^"]*)"/.exec(attrs)?.[1] ?? "silver";
    const named: Record<string, string> = { silver: "C0C0C0", gray: "808080", black: "000000", red: "FF0000", blue: "0000FF", white: "FFFFFF" };
    const color = fill.startsWith("#") ? fill.slice(1).toUpperCase() : named[fill] ?? "C0C0C0";
    const semi = /<v:fill\b[^>]*opacity="([^"]*)"/.test(inner);
    return { kind: "text", text, font, ...(fsize > 1 ? { size: fsize } : {}), color, semi, diagonal: /rotation:315/.test(style), width, height };
  }
  const data = /<v:imagedata\b([^>]*)\/?>/.exec(inner);
  const rid = data ? /r:id="([^"]*)"/.exec(data[1])?.[1] : undefined;
  const img = rid ? image(rid) : undefined;
  if (!img) return undefined;
  return { kind: "picture", image: img, washout: !!data && /gain=/.test(data[1]), width, height };
}

/** Alternativtext plus the drawing's name and Ausblenden (wp:docPr), or undefined when empty. */
function docPrAlt(docPr: Element | undefined, alt: AltText): AltText | undefined {
  const name = docPr?.getAttribute("name") || undefined;
  const h = docPr?.getAttribute("hidden");
  const out: AltText = { ...alt, ...(name ? { name } : {}), ...(h === "1" || h === "true" ? { hidden: true } : {}) };
  const clean = Object.fromEntries(Object.entries(out).filter(([, v]) => v !== undefined)) as AltText;
  return Object.keys(clean).length ? clean : undefined;
}

/**
 * Entwurf > Designs > Nach Designs suchen: a theme file (.thmx: its theme1.xml), or Word's
 * colour or font scheme files (.xml: a:clrScheme, a:fontScheme). What it does not contain is
 * left out.
 */
// ----- styles from Word's XML (the table style gallery) -----------------------------------------------

/** Theme colour names of WordprocessingML (w:themeColor) and the theme's keys. */
const THEME_KEYS: Record<string, string> = {
  text1: "dk1", background1: "lt1", text2: "dk2", background2: "lt2", dark1: "dk1", light1: "lt1", dark2: "dk2", light2: "lt2",
  hyperlink: "hlink", followedHyperlink: "folHlink",
};

/**
 * A theme colour with Word's w:themeTint / w:themeShade (hex 00–FF): the HSL lightness is
 * lifted towards white (L·t + 1 − t) or lowered (L·s), as Word computes it (checked against
 * Word's own values, e.g. Akzent 1 with tint 99).
 */
export function themedHex(base: string, tint?: string, shade?: string): string {
  const rgb = [0, 2, 4].map((i) => parseInt(base.slice(i, i + 2), 16) / 255);
  const max = Math.max(...rgb), min = Math.min(...rgb);
  let l = (max + min) / 2;
  const d = max - min;
  const s = d === 0 ? 0 : d / (1 - Math.abs(2 * l - 1));
  let hue = 0;
  if (d) {
    const [r, g, b] = rgb;
    hue = max === r ? ((g - b) / d) % 6 : max === g ? (b - r) / d + 2 : (r - g) / d + 4;
    hue *= 60;
    if (hue < 0) hue += 360;
  }
  if (tint) { const t = parseInt(tint, 16) / 255; l = l * t + (1 - t); }
  if (shade) l *= parseInt(shade, 16) / 255;
  const c = (1 - Math.abs(2 * l - 1)) * s, x = c * (1 - Math.abs(((hue / 60) % 2) - 1)), m = l - c / 2;
  const [r1, g1, b1] = hue < 60 ? [c, x, 0] : hue < 120 ? [x, c, 0] : hue < 180 ? [0, c, x] : hue < 240 ? [0, x, c] : hue < 300 ? [x, 0, c] : [c, 0, x];
  return [r1, g1, b1].map((v) => Math.round(Math.min(1, Math.max(0, v + m)) * 255).toString(16).padStart(2, "0")).join("").toUpperCase();
}

/** Word's XML with every theme-referencing colour (w:color/w:fill) computed from `colors`. */
export function rethemeXml(xml: string, colors: Readonly<Record<string, string>>): string {
  return xml.replace(/<w:[A-Za-z]+\b[^>]*>/g, (tag) => {
    const get = (a: string) => new RegExp(`w:${a}="([^"]*)"`).exec(tag)?.[1];
    let t = tag;
    const tc = get("themeColor"), tf = get("themeFill");
    const base = (k: string | undefined) => (k ? colors[THEME_KEYS[k] ?? k] : undefined);
    if (base(tc)) t = t.replace(/w:color="[^"]*"/, `w:color="${themedHex(base(tc)!, get("themeTint"), get("themeShade"))}"`);
    if (base(tf)) t = t.replace(/w:fill="[^"]*"/, `w:fill="${themedHex(base(tf)!, get("themeFillTint"), get("themeFillShade"))}"`);
    return t;
  });
}

/** A style from Word's w:style XML (a gallery table style), with the document's theme colours. */
export function styleFromXml(styleXml: string, colors?: Readonly<Record<string, string>>): StyleDef | undefined {
  const xml = colors ? rethemeXml(styleXml, colors) : styleXml;
  const sheet = readStyles(`<w:styles xmlns:w="${W}">${xml}</w:styles>`, { major: "Calibri Light", minor: "Calibri" });
  const id = /w:styleId="([^"]+)"/.exec(xml)?.[1];
  return id ? sheet.get(id) : undefined;
}

export function readThemeFile(bytes: Uint8Array): { colors?: Record<string, string>; colorName?: string; major?: string; minor?: string; fontName?: string; name?: string } | null {
  let xml: string | null = null;
  try {
    if (bytes[0] === 0x50 && bytes[1] === 0x4b) {
      const files = unzipSync(bytes);
      const path = Object.keys(files).find((f) => /theme\/theme\/theme\d*\.xml$/.test(f)) ?? Object.keys(files).find((f) => /theme\d*\.xml$/.test(f));
      xml = path ? strFromU8(files[path]) : null;
    } else xml = strFromU8(bytes);
  } catch { return null; }
  if (!xml) return null;
  const doc = parse(xml);
  const A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main";
  const scheme = doc.getElementsByTagNameNS(A_NS, "clrScheme")[0];
  const fonts = doc.getElementsByTagNameNS(A_NS, "fontScheme")[0];
  const latin = (tag: string) => {
    const el = fonts ? Array.from(fonts.getElementsByTagNameNS(A_NS, tag))[0] : undefined;
    return el ? Array.from(el.children).find((c) => c.localName === "latin")?.getAttribute("typeface") || undefined : undefined;
  };
  const out = {
    ...(scheme ? { colors: themeColors(doc), colorName: scheme.getAttribute("name") || undefined } : {}),
    ...(fonts ? { major: latin("majorFont"), minor: latin("minorFont"), fontName: fonts.getAttribute("name") || undefined } : {}),
    ...(doc.documentElement.localName === "theme" ? { name: doc.documentElement.getAttribute("name") || undefined } : {}),
  };
  return Object.keys(out).length ? out : null;
}

/** Linked text boxes after their first box: the chain's story, so the text flows through them. */
function linkChains(paras: Para[], stories: ReadonlyMap<number, string>): void {
  if (!stories.size) return;
  const fix = (s: ShapeInfo | undefined) => (s?.chain && s.chain.seq > 0 && stories.has(s.chain.id) ? { ...s, chain: { ...s.chain, story: stories.get(s.chain.id) } } : s);
  paras.forEach((p, i) => {
    if (!p.floats?.some((f) => f.shape?.chain || f.group)) return;
    paras[i] = { ...p, floats: p.floats.map((f) => ({ ...f, shape: fix(f.shape), ...(f.group ? { group: { members: f.group.members.map((m) => ({ ...m, shape: fix(m.shape) })) } } : {}) })) };
  });
}

/** w:footnotePr and w:endnotePr of settings.xml (only what the dialog Fuß- und Endnote sets). */
export function readNoteSettings(settings: Element): DocSettings["notes"] {
  const one = (name: string): NoteProps | undefined => {
    const pr = child(settings, name);
    if (!pr) return undefined;
    const v = (k: string) => attr(child(pr, k), "val");
    const pos = v("pos"), fmt = v("numFmt"), start = v("numStart"), restart = v("numRestart");
    const out: NoteProps = {
      ...(pos === "pageBottom" || pos === "beneathText" || pos === "sectEnd" || pos === "docEnd" ? { pos } : {}),
      ...(fmt ? { numFmt: fmt } : {}), ...(start && Number(start) > 0 ? { numStart: Number(start) } : {}),
      ...(restart === "continuous" || restart === "eachSect" || restart === "eachPage" ? { numRestart: restart } : {}),
    };
    return Object.keys(out).length ? out : undefined;
  };
  const footnote = one("footnotePr"), endnote = one("endnotePr");
  return footnote || endnote ? { footnote, endnote } : undefined;
}

export function readDocx(bytes: Uint8Array): ReadResult {
  const files = unzipSync(bytes);
  const r = new Reader(files);
  const main = r.text("word/document.xml");
  if (!main) throw new Error("Die Datei ist kein Word-Dokument (word/document.xml fehlt).");
  const themeXml = r.text("word/theme/theme1.xml");
  r.themeColors = themeColors(themeXml ? parse(themeXml) : null);
  r.matrix = formatMatrix(themeXml ? parse(themeXml) : null);
  const theme = { ...themeFonts(themeXml), colors: r.themeColors };
  const stylesXml = r.text("word/styles.xml");
  const styles = readStyles(stylesXml, theme);
  const settingsXml = r.text("word/settings.xml");
  let defaultTab = 36;
  let evenOdd = false;
  let hyphenation: DocState["settings"]["hyphenation"];
  let trackChanges = false;
  if (settingsXml) {
    const s = parse(settingsXml).documentElement;
    defaultTab = twips(attr(child(s, "defaultTabStop"), "val"), 36);
    evenOdd = onOff(child(s, "evenAndOddHeaders")) ?? false;
    trackChanges = onOff(child(s, "trackRevisions")) ?? false;
    // Layout > Silbentrennung.
    const auto = onOff(child(s, "autoHyphenation")) ?? false;
    const limit = num(attr(child(s, "consecutiveHyphenLimit"), "val")) ?? 0;
    const zone = num(attr(child(s, "hyphenationZone"), "val"));
    const caps = !(onOff(child(s, "doNotHyphenateCaps")) ?? false);
    if (auto || limit || !caps) hyphenation = { auto, caps, limit, ...(zone !== undefined ? { zone: zone / 20 } : {}) };
    r.mathSettings = readMathPr(s);
    r.noteSettings = readNoteSettings(s);
  }
  // Word keeps the document's sources in a customXml part.
  let bibliography: DocxPackage["bibliography"];
  for (const path of Object.keys(files).filter((f) => /^customXml\/item\d+\.xml$/i.test(f)).sort()) {
    const xml = r.text(path);
    const read = xml && xml.includes("bibliography") ? readSources(parse(xml)) : null;
    if (read) { bibliography = { path, read }; break; }
  }
  const docRels = r.rels("word/document.xml");
  const partOf = (type: string) => {
    const rel = docRels.list.find((x) => x.type === type);
    return rel ? resolveTarget("word/document.xml", rel.target) : null;
  };
  const numberingPath = partOf(REL.numbering);
  const numberingXml = numberingPath ? r.text(numberingPath) : null;
  const numbering = readNumbering(numberingXml, styles);

  // Notes and comments first: the main text refers to them.
  const notes: DocxPackage["notes"] = {};
  const fnPath = partOf(REL.footnotes), enPath = partOf(REL.endnotes);
  const fn = fnPath ? readNotes(r, fnPath, "fn") : null;
  if (fn && fnPath) notes.footnotes = { path: fnPath, ...fn };
  const en = enPath ? readNotes(r, enPath, "en") : null;
  if (en && enPath) notes.endnotes = { path: enPath, ...en };
  const cmPath = partOf(REL.comments);
  const cm = cmPath ? readComments(r, cmPath, partOf(REL.commentsExtended)) : null;

  const doc = parse(main);
  const root = doc.documentElement;
  const body = child(root, "body");
  // Seitenfarbe.
  const bgEl = child(root, "background");
  const bgColor = bgEl ? attr(bgEl, "color") : null;
  const bgTheme = bgEl ? ["themeColor", "themeShade", "themeTint"].filter((a) => attr(bgEl, a) !== null).map((a) => `w:${a}="${attr(bgEl, a)}"`).join(" ") : "";
  const background = bgColor && bgColor !== "auto" ? { color: bgColor.toUpperCase(), ...(bgTheme ? { colorTheme: bgTheme } : {}) } : undefined;
  const pr = new PartReader(r, "word/document.xml");
  const paras = pr.blocks(body ? kids(body) : []);
  const trailing = pr.leading;
  if (!paras.length || paras[paras.length - 1].cell) paras.push({ text: "", runs: [], props: {}, mark: intern({}) });
  const section = pr.finalSect ?? DEFAULT_SECTION;

  const styleIds = new Set<string>(styles.styles.keys());
  linkChains(paras, r.chainStories);
  return {
    state: {
      paras, styles, settings: { section, defaultTab, evenOdd: evenOdd || undefined, trackChanges: trackChanges || undefined, hyphenation, math: r.mathSettings, bibliography: bibliography?.read, notes: r.noteSettings, watermark: r.watermark, background }, stories: r.stories, numbering,
      comments: cm?.meta,
    },
    pkg: {
      files, rootTag: rootTagOf(root), trailing, stylesXml, styleIds, hasTheme: !!themeXml, storyParts: r.storyParts,
      themeRead: JSON.stringify(styles.theme), defaultsRead: JSON.stringify([styles.defaultsR, styles.defaultsP]),
      watermarkRead: JSON.stringify(r.watermark ?? null), backgroundXml: bgEl ? outer(bgEl) : null,
      numbering, numberingRoot: numberingXml ? rootTagOf(parse(numberingXml).documentElement) : null, notes,
      comments: cm && cmPath ? { path: cmPath, rootTag: cm.rootTag } : null, settingsXml,
      originalStories: new Map(r.stories), originalComments: cm?.meta ?? null, bibliography,
    },
    partial: [...r.partial],
  };
}

/** The points of a wp:wrapPolygon (Eng, Transparent), in 21600ths of the object's box. */
function wrapPolygon(wrapEl: Element | undefined): [number, number][] | undefined {
  const poly = wrapEl ? Array.from(wrapEl.children).find((x) => x.localName === "wrapPolygon") : undefined;
  if (!poly) return undefined;
  const pts = Array.from(poly.children).filter((x) => x.localName === "start" || x.localName === "lineTo")
    .map((x) => [Number(x.getAttribute("x") ?? 0), Number(x.getAttribute("y") ?? 0)] as [number, number]);
  return pts.length >= 3 ? pts : undefined;
}
