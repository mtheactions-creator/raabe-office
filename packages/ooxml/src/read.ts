// Reading .docx into the text engine's model: the main text, headers and footers, foot- and
// endnotes and comments as stories; sections, tables, lists, pictures, fields and tracked
// changes as the model describes them. Everything RaabeWrite does not interpret is kept as
// its original XML in place, so writing back loses nothing (D-004).

import { readExtent, readPicture, themeColors } from "./drawingml";
import { strFromU8, unzipSync } from "fflate";
import {
  A4, COLUMN_BREAK, DEFAULT_SECTION, LINE_BREAK, OBJ, ONE_COLUMN, PAGE_BREAK, StyleSheet, defaultStyleSheet, intern, nextTableId, plainPicture,
  type Anchor, type BlockObject, type CellDef, type CellRef, type CharProps, type CommentMeta, type DocState, type FloatObject,
  type HeaderRefs, type ImageData, type InlineObject, type Numbering, type Para, type ParaGroup, type ParaProps, type Revision,
  type RowDef, type Run, type Section, type StyleDef, type StyleType, type TableDef, type TableStyleDef, type TableStylePart,
} from "@raabe/text-engine";
import { ommlToLatex } from "./math";
import { readNumbering } from "./numbering";
import { REL, Rels, mimeOf, relsPath, resolveTarget } from "./package";
import { readBorders, readPPr, readRPr, readShading, type RevisionOf } from "./props";
import { attr, child, kids, num, onOff, outer, parse, textOf, W } from "./xml";

const WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
const A = "http://schemas.openxmlformats.org/drawingml/2006/main";
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

function themeFonts(xml: string | null): { major: string; minor: string } {
  if (!xml) return { major: "Calibri Light", minor: "Calibri" };
  const doc = parse(xml);
  const pick = (tag: string, fallback: string) => {
    const el = doc.getElementsByTagNameNS(A, tag)[0];
    const latin = el ? Array.from(el.children).find((c) => c.localName === "latin") : undefined;
    return latin?.getAttribute("typeface") || fallback;
  };
  return { major: pick("majorFont", "Calibri Light"), minor: pick("minorFont", "Calibri") };
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

function readStyles(xml: string | null, theme: { major: string; minor: string; colors?: Record<string, string> }): StyleSheet {
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
  return new StyleSheet(styles, defaultsR, defaultsP, { majorFont: theme.major, minorFont: theme.minor, colors: theme.colors });
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
  readonly stories = new Map<string, Para[]>();
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
    this.stories.set(id, ensurePara(pr.blocks(kids(root))));
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
          b.addAnchor({ data: { raw: open }, link: { id, edge: "start" } });
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
          const latex = ommlToLatex(el) ?? undefined;
          b.addObject({ kind: "math", data: { raw: outer(el) }, display: mathText(el), latex }, intern({ ...this.firstRunProps(el, rev), italic: true }));
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

  /** A picture: inline (an object in the text) or anchored (positioned on the page). */
  private drawing(c: Element, b: ParaBuilder, props: CharProps): void {
    const inline = c.getElementsByTagNameNS(WP, "inline")[0];
    const anchor = c.getElementsByTagNameNS(WP, "anchor")[0];
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
    const alt = descr || title || decorative ? { descr: descr || undefined, title: title || undefined, decorative: decorative || undefined } : undefined;
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
      dist: [d("distT"), d("distB"), d("distL"), d("distR")], pic, alt,
      z: Number(anchor.getAttribute("relativeHeight") ?? 0) || undefined,
    });
  }

  // ----- tables ---------------------------------------------------------------------------

  table(tbl: Element): Para[] {
    const tblPr = child(tbl, "tblPr");
    const extra: string[] = [];
    let style: string | undefined, align: TableDef["align"], indent: number | undefined;
    let look: TableDef["look"];
    if (tblPr) {
      for (const c of kids(tblPr)) {
        switch (c.localName) {
          case "tblStyle": style = attr(c, "val") ?? undefined; break;
          case "jc": { const v = attr(c, "val"); align = v === "center" ? "center" : v === "right" || v === "end" ? "right" : "left"; break; }
          case "tblInd": indent = (num(attr(c, "w")) ?? 0) / 20; break;
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

export function readDocx(bytes: Uint8Array): ReadResult {
  const files = unzipSync(bytes);
  const r = new Reader(files);
  const main = r.text("word/document.xml");
  if (!main) throw new Error("Die Datei ist kein Word-Dokument (word/document.xml fehlt).");
  const themeXml = r.text("word/theme/theme1.xml");
  r.themeColors = themeColors(themeXml ? parse(themeXml) : null);
  const theme = { ...themeFonts(themeXml), colors: r.themeColors };
  const stylesXml = r.text("word/styles.xml");
  const styles = readStyles(stylesXml, theme);
  const settingsXml = r.text("word/settings.xml");
  let defaultTab = 36;
  let evenOdd = false;
  let trackChanges = false;
  if (settingsXml) {
    const s = parse(settingsXml).documentElement;
    defaultTab = twips(attr(child(s, "defaultTabStop"), "val"), 36);
    evenOdd = onOff(child(s, "evenAndOddHeaders")) ?? false;
    trackChanges = onOff(child(s, "trackRevisions")) ?? false;
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
  const pr = new PartReader(r, "word/document.xml");
  const paras = pr.blocks(body ? kids(body) : []);
  const trailing = pr.leading;
  if (!paras.length || paras[paras.length - 1].cell) paras.push({ text: "", runs: [], props: {}, mark: intern({}) });
  const section = pr.finalSect ?? DEFAULT_SECTION;

  const styleIds = new Set<string>(styles.styles.keys());
  return {
    state: {
      paras, styles, settings: { section, defaultTab, evenOdd: evenOdd || undefined, trackChanges: trackChanges || undefined }, stories: r.stories, numbering,
      comments: cm?.meta,
    },
    pkg: {
      files, rootTag: rootTagOf(root), trailing, stylesXml, styleIds, hasTheme: !!themeXml, storyParts: r.storyParts,
      numbering, numberingRoot: numberingXml ? rootTagOf(parse(numberingXml).documentElement) : null, notes,
      comments: cm && cmPath ? { path: cmPath, rootTag: cm.rootTag } : null, settingsXml,
      originalStories: new Map(r.stories), originalComments: cm?.meta ?? null,
    },
    partial: [...r.partial],
  };
}
