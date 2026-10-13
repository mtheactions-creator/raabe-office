// Reading .docx into the text engine's model. Everything RaabeWrite cannot edit yet is kept
// as its original XML (inline objects, anchors, blocks) so writing back loses nothing (D-004).

import { unzipSync, strFromU8 } from "fflate";
import {
  A4, OBJ, LINE_BREAK, PAGE_BREAK, StyleSheet, defaultStyleSheet, intern,
  type Anchor, type BlockObject, type CharProps, type DocState, type InlineObject, type Para, type ParaProps,
  type Run, type StyleDef, type StyleType,
} from "@raabe/text-engine";
import { readPPr, readRPr } from "./props";
import { attr, child, kids, num, outer, parse, textOf, W } from "./xml";

const WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
const A = "http://schemas.openxmlformats.org/drawingml/2006/main";

/** Everything needed to write the file back: original parts and wrappers. */
export interface DocxPackage {
  files: Record<string, Uint8Array>;
  /** Start tag of w:document with all namespace declarations. */
  rootTag: string;
  /** The final section properties (page size, margins, header/footer references). */
  sectPr: string | null;
  /** Opaque body content after the last paragraph (e.g. trailing bookmarks). */
  trailing: string[];
  stylesXml: string | null;
  styleIds: Set<string>;
  hasTheme: boolean;
}

export interface ReadResult {
  state: DocState;
  pkg: DocxPackage;
  /** German names of content RaabeWrite shows only partly (it is still saved unchanged). */
  partial: string[];
}

const LABEL = {
  tables: "Tabellen",
  images: "Bilder und Formen",
  lists: "Aufzählungen und Nummerierungen",
  headers: "Kopf- und Fußzeilen",
  notes: "Fußnoten",
  comments: "Kommentare",
  revisions: "nachverfolgte Änderungen",
  math: "Formeln",
  fields: "Felder (z. B. Inhaltsverzeichnis, Seitenzahlen)",
  sections: "Abschnittswechsel",
  columns: "mehrspaltiger Text",
  borders: "Rahmen und Schattierung",
  controls: "Inhaltssteuerelemente",
  links: "Links",
};

const twips = (v: string | null, fallback: number) => (num(v) ?? fallback * 20) / 20;

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

function readStyles(xml: string | null, theme: { major: string; minor: string }): StyleSheet {
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
    styles.set(id, {
      id,
      name: attr(child(s, "name"), "val") ?? id,
      type: (attr(s, "type") ?? "paragraph") as StyleType,
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
    });
  }
  return new StyleSheet(styles, defaultsR, defaultsP, { majorFont: theme.major, minorFont: theme.minor });
}

// ----- paragraphs ---------------------------------------------------------------------------

interface FieldState {
  depth: number;
  xml: string[];
  display: string;
  inResult: boolean;
  props: CharProps;
  instr: string;
}

interface Ctx {
  partial: Set<string>;
  noteNo: number;
}

/** Updates a field's nesting depth for the fldChar elements inside `el`. */
function trackField(el: Element, f: FieldState): void {
  const walk = (e: Element) => {
    for (const c of Array.from(e.children)) {
      if (c.localName === "fldChar") {
        const t = attr(c, "fldCharType");
        if (t === "begin") f.depth++;
        else if (t === "separate" && f.depth === 1) f.inResult = true;
        else if (t === "end") f.depth--;
      } else if (c.localName === "instrText" && f.depth === 1 && !f.inResult) {
        f.instr += c.textContent ?? "";
      } else if (c.localName === "t" && f.inResult && f.depth >= 1) {
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

function containsFieldBegin(el: Element): boolean {
  for (const c of Array.from(el.getElementsByTagNameNS(W, "fldChar"))) if (attr(c, "fldCharType") === "begin") return true;
  return false;
}

class ParaBuilder {
  text = "";
  runs: Run[] = [];
  objects: InlineObject[] = [];
  anchors: Anchor[] = [];

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

  addAnchor(xml: string): void {
    this.anchors.push({ at: this.text.length, data: xml });
  }
}

function inlineImage(el: Element): { width: number; height: number } | null {
  const inline = el.getElementsByTagNameNS(WP, "inline")[0];
  if (!inline) return null;
  const ext = Array.from(inline.children).find((c) => c.localName === "extent");
  const cx = Number(ext?.getAttribute("cx") ?? 0), cy = Number(ext?.getAttribute("cy") ?? 0);
  return { width: cx / 12700, height: cy / 12700 };
}

function readRun(r: Element, b: ParaBuilder, ctx: Ctx): void {
  const rPrEl = child(r, "rPr");
  const props = intern(readRPr(rPrEl));
  const rPrXml = rPrEl ? outer(rPrEl) : "";
  let pending = "";
  const flush = () => { b.addText(pending, props); pending = ""; };
  const single = (c: Element) => `<w:r xmlns:w="${W}">${rPrXml}${outer(c)}</w:r>`;
  if (rPrEl && (child(rPrEl, "shd") || child(rPrEl, "bdr"))) ctx.partial.add(LABEL.borders);
  for (const c of kids(r)) {
    switch (c.localName) {
      case "rPr": case "lastRenderedPageBreak": break;
      case "t": pending += c.textContent ?? ""; break;
      case "tab": pending += "\t"; break;
      case "cr": pending += LINE_BREAK; break;
      case "noBreakHyphen": pending += "\u2011"; break;
      case "softHyphen": pending += "\u00ad"; break;
      case "br": {
        const type = attr(c, "type");
        if (type === "page") pending += PAGE_BREAK;
        else if (type === "column") { flush(); b.addObject({ kind: "other", data: single(c), display: "" }, props); ctx.partial.add(LABEL.columns); }
        else pending += LINE_BREAK;
        break;
      }
      case "footnoteReference": case "endnoteReference":
        flush();
        ctx.noteNo++;
        b.addObject({ kind: "note", data: single(c), display: String(ctx.noteNo) }, props);
        ctx.partial.add(LABEL.notes);
        break;
      case "commentReference": case "annotationRef":
        flush(); b.addAnchor(single(c)); ctx.partial.add(LABEL.comments); break;
      case "drawing": case "pict": case "object": case "AlternateContent": {
        flush();
        const img = c.localName === "drawing" || c.localName === "AlternateContent" ? inlineImage(c) : null;
        if (img) b.addObject({ kind: "image", data: single(c), display: "", width: img.width, height: img.height }, props);
        else b.addAnchor(single(c));
        ctx.partial.add(LABEL.images);
        break;
      }
      case "sym": {
        flush();
        const code = parseInt(attr(c, "char") ?? "", 16);
        const ch = Number.isFinite(code) ? String.fromCharCode(code >= 0xf000 ? code - 0xf000 : code) : "?";
        b.addObject({ kind: "other", data: single(c), display: ch }, props);
        break;
      }
      default:
        flush();
        b.addObject({ kind: "other", data: single(c), display: textOf(c) }, props);
    }
  }
  flush();
}

const ANCHOR_ELEMENTS = new Set(["bookmarkStart", "bookmarkEnd", "commentRangeStart", "commentRangeEnd", "permStart", "permEnd",
  "moveFromRangeStart", "moveFromRangeEnd", "moveToRangeStart", "moveToRangeEnd", "customXmlInsRangeStart",
  "customXmlInsRangeEnd", "customXmlDelRangeStart", "customXmlDelRangeEnd"]);

function firstRunProps(el: Element): CharProps {
  const r = el.getElementsByTagNameNS(W, "r")[0];
  return intern(readRPr(r ? child(r, "rPr") : null));
}

/** Parses a w:p. Returns an open field when a field continues into the next paragraph. */
function readParagraph(p: Element, ctx: Ctx): { para: Para; open: FieldState | null } {
  const pPrEl = child(p, "pPr");
  const { props, mark } = readPPr(pPrEl);
  if (pPrEl) {
    if (child(pPrEl, "numPr")) ctx.partial.add(LABEL.lists);
    if (child(pPrEl, "sectPr")) ctx.partial.add(LABEL.sections);
    if (child(pPrEl, "pBdr") || child(pPrEl, "shd")) ctx.partial.add(LABEL.borders);
  }
  const b = new ParaBuilder();
  let field: FieldState | null = null;
  for (const c of kids(p)) {
    if (c === pPrEl) continue;
    if (field) {
      field.xml.push(outer(c));
      trackField(c, field);
      if (field.depth <= 0) {
        b.addObject({ kind: "field", data: field.xml.join(""), display: field.display }, field.props);
        field = null;
      }
      continue;
    }
    const name = c.localName;
    if (name === "r" && containsFieldBegin(c)) {
      field = { depth: 0, xml: [outer(c)], display: "", inResult: false, props: intern(readRPr(child(c, "rPr"))), instr: "" };
      trackField(c, field);
      ctx.partial.add(LABEL.fields);
      if (field.depth <= 0) { b.addObject({ kind: "field", data: field.xml.join(""), display: field.display }, field.props); field = null; }
      continue;
    }
    if (name === "r") { readRun(c, b, ctx); continue; }
    if (name === "proofErr") continue;
    if (ANCHOR_ELEMENTS.has(name)) {
      b.addAnchor(outer(c));
      if (name.startsWith("comment")) ctx.partial.add(LABEL.comments);
      continue;
    }
    switch (name) {
      case "hyperlink":
        b.addObject({ kind: "link", data: outer(c), display: textOf(c) }, firstRunProps(c));
        ctx.partial.add(LABEL.links);
        break;
      case "fldSimple":
        b.addObject({ kind: "field", data: outer(c), display: textOf(c) }, firstRunProps(c));
        ctx.partial.add(LABEL.fields);
        break;
      case "ins": case "moveTo":
        b.addObject({ kind: "revision", data: outer(c), display: textOf(c) }, intern({ ...firstRunProps(c), underline: "single", color: "C00000" }));
        ctx.partial.add(LABEL.revisions);
        break;
      case "del": case "moveFrom":
        b.addObject({ kind: "revision", data: outer(c), display: textOf(c, true) }, intern({ ...firstRunProps(c), strike: true, color: "C00000" }));
        ctx.partial.add(LABEL.revisions);
        break;
      case "oMath": case "oMathPara":
        b.addObject({ kind: "other", data: outer(c), display: mathText(c) }, intern({ ...mark, italic: true }));
        ctx.partial.add(LABEL.math);
        break;
      case "sdt":
        b.addObject({ kind: "other", data: outer(c), display: textOf(c) }, firstRunProps(c));
        ctx.partial.add(LABEL.controls);
        break;
      default:
        b.addObject({ kind: "other", data: outer(c), display: textOf(c) }, firstRunProps(c));
    }
  }
  const para: Para = {
    text: b.text,
    runs: b.runs,
    props,
    mark: intern(mark),
    objects: b.objects.length ? b.objects : undefined,
    anchors: b.anchors.length ? b.anchors : undefined,
  };
  return { para, open: field };
}

function mathText(el: Element): string {
  let out = "";
  for (const t of Array.from(el.getElementsByTagName("*"))) if (t.localName === "t") out += t.textContent ?? "";
  return out || "[Formel]";
}

function tableBlock(tbl: Element): BlockObject {
  const rows = kids(tbl).filter((e) => e.localName === "tr");
  const cols = Math.max(0, ...rows.map((r) => kids(r).filter((c) => c.localName === "tc").length));
  const lines: string[] = [];
  let height = 0;
  for (const r of rows) {
    const cells = kids(r).filter((c) => c.localName === "tc");
    const paraCount = Math.max(1, ...cells.map((c) => kids(c).filter((e) => e.localName === "p").length));
    height += paraCount * 14.5 + 1;
    if (lines.length < 8) lines.push(cells.map((c) => textOf(c).trim()).join("  |  "));
  }
  return { kind: "table", data: outer(tbl), label: `Tabelle (${rows.length} × ${cols}) – wird ab Schritt 3 bearbeitbar angezeigt`, lines, height: Math.max(height, 18) };
}

// ----- the package --------------------------------------------------------------------------

export function readDocx(bytes: Uint8Array): ReadResult {
  const files = unzipSync(bytes);
  const text = (path: string) => (files[path] ? strFromU8(files[path]) : null);
  const main = text("word/document.xml");
  if (!main) throw new Error("Die Datei ist kein Word-Dokument (word/document.xml fehlt).");
  const themeXml = text("word/theme/theme1.xml");
  const theme = themeFonts(themeXml);
  const stylesXml = text("word/styles.xml");
  const styles = readStyles(stylesXml, theme);
  const settingsXml = text("word/settings.xml");
  let defaultTab = 36;
  if (settingsXml) {
    const s = parse(settingsXml).documentElement;
    defaultTab = twips(attr(child(s, "defaultTabStop"), "val"), 36);
  }

  const doc = parse(main);
  const root = doc.documentElement;
  const shallow = root.cloneNode(false) as Element;
  const rootTag = outer(shallow).replace(/\/>\s*$/, ">");
  const body = child(root, "body");
  const ctx: Ctx = { partial: new Set(), noteNo: 0 };
  const paras: Para[] = [];
  let leading: string[] = [];
  let sectPr: string | null = null;
  let page = A4;
  const pushPara = (p: Para) => {
    paras.push(leading.length ? { ...p, leading } : p);
    leading = [];
  };
  const els = body ? kids(body) : [];
  for (let i = 0; i < els.length; i++) {
    const el = els[i];
    const name = el.localName;
    if (name === "p") {
      const { para, open } = readParagraph(el, ctx);
      if (!open) { pushPara(para); continue; }
      // A field that runs over several paragraphs (e.g. a table of contents) becomes one block.
      const xml = [outer(el)];
      const lines = [textOf(el)];
      const f = open;
      while (f.depth > 0 && i + 1 < els.length) {
        i++;
        xml.push(outer(els[i]));
        lines.push(textOf(els[i]));
        trackField(els[i], f);
      }
      const isToc = /^\s*TOC\b/i.test(f.instr);
      pushPara({
        text: OBJ, runs: [{ len: 1, props: intern({}) }], props: para.props, mark: para.mark,
        block: { kind: "field", data: xml.join(""), label: isToc ? "Inhaltsverzeichnis" : "Feld", lines, height: lines.length * 14.5 },
      });
    } else if (name === "tbl") {
      ctx.partial.add(LABEL.tables);
      pushPara({ text: OBJ, runs: [{ len: 1, props: intern({}) }], props: {}, mark: intern({}), block: tableBlock(el) });
    } else if (name === "sectPr") {
      sectPr = outer(el);
      const pgSz = child(el, "pgSz"), pgMar = child(el, "pgMar");
      page = {
        width: twips(attr(pgSz, "w"), A4.width), height: twips(attr(pgSz, "h"), A4.height),
        marginTop: Math.abs(twips(attr(pgMar, "top"), A4.marginTop)), marginBottom: Math.abs(twips(attr(pgMar, "bottom"), A4.marginBottom)),
        marginLeft: twips(attr(pgMar, "left"), A4.marginLeft), marginRight: twips(attr(pgMar, "right"), A4.marginRight),
      };
      if (child(el, "headerReference") || child(el, "footerReference")) ctx.partial.add(LABEL.headers);
      const cols = child(el, "cols");
      if (cols && (num(attr(cols, "num")) ?? 1) > 1) ctx.partial.add(LABEL.columns);
    } else if (name === "sdt") {
      ctx.partial.add(LABEL.controls);
      pushPara({ text: OBJ, runs: [{ len: 1, props: intern({}) }], props: {}, mark: intern({}),
        block: { kind: "other", data: outer(el), label: "Inhaltssteuerelement", lines: [textOf(el)], height: 14.5 * Math.max(1, kids(el).length) } });
    } else {
      // Body-level markers (bookmarks …) travel with the next paragraph.
      leading.push(outer(el));
    }
  }
  const trailing = leading;
  if (!paras.length) paras.push({ text: "", runs: [], props: {}, mark: intern({}) });

  const styleIds = new Set<string>(styles.styles.keys());
  return {
    state: { paras, styles, settings: { page, defaultTab } },
    pkg: { files, rootTag, sectPr, trailing, stylesXml, styleIds, hasTheme: !!themeXml },
    partial: [...ctx.partial],
  };
}
