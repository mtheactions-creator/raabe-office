// The document model. A paragraph is its text plus a partition of that text into runs of
// identical character formatting, direct paragraph properties and the properties of the
// paragraph mark. Content the engine does not understand (fields, hyperlinks, images,
// bookmarks, tables …) travels along as opaque payloads so that files can be written back
// without loss (D-004). All values are immutable: edits create new paragraph objects, which
// keeps undo cheap and lets the layout cache per paragraph object.

import type { PictureFormat } from "./picture";

export type Underline = "single" | "double" | "thick" | "dotted" | "dash" | "wave" | "words" | "none" | string;

/** Character formatting. `undefined` means "inherit from styles". Sizes in points. */
export interface CharProps {
  readonly font?: string;
  /** Theme font reference ("minorHAnsi", "majorHAnsi") when the font comes from the theme. */
  readonly fontTheme?: string;
  readonly size?: number;
  readonly bold?: boolean;
  readonly italic?: boolean;
  readonly underline?: Underline;
  readonly strike?: boolean;
  readonly dstrike?: boolean;
  readonly vertAlign?: "superscript" | "subscript" | "baseline";
  /** "RRGGBB" or "auto". */
  readonly color?: string;
  /** Theme colour attributes as read from the file, kept while the colour is unchanged. */
  readonly colorTheme?: string;
  /** OOXML highlight name ("yellow", "green" …) or "none". */
  readonly highlight?: string;
  readonly caps?: boolean;
  readonly smallCaps?: boolean;
  readonly hidden?: boolean;
  /** Extra character spacing in points (Laufweite). */
  readonly spacing?: number;
  /** Kerning applies from this font size (points) upwards; 0 = off. */
  readonly kern?: number;
  readonly lang?: string;
  /** Sprache > Rechtschreibung und Grammatik nicht prüfen (w:noProof). */
  readonly noProof?: boolean;
  /** Character style id. */
  readonly rStyle?: string;
  /** Run properties the engine does not interpret, kept for writing back. */
  readonly extra?: readonly unknown[];
  /** Texteffekte und Typografie: Word's w14 elements (glow, shadow, textOutline …) as XML. */
  readonly textFx?: string;
  /** Tracked change this text belongs to (Änderungen nachverfolgen). */
  readonly rev?: Revision;
}

/** A tracked insertion or deletion. Shared objects: runs of one change point to the same one. */
export interface Revision {
  readonly kind: "ins" | "del";
  readonly author: string;
  /** ISO date as in the file ("2026-10-13T10:00:00Z"); may be empty. */
  readonly date: string;
  readonly id: number;
}

export type Align = "left" | "center" | "right" | "both";
export type LineRule = "auto" | "exact" | "atLeast";

/** One border line (paragraph, cell, table, page). Lengths in points. */
export interface BorderSpec {
  /** OOXML border style: "single", "double", "dotted", "dashed", "thick", "nil", "none" … */
  readonly style: string;
  /** Line width in points (w:sz is eighths of a point). */
  readonly size: number;
  /** "RRGGBB" or "auto". */
  readonly color: string;
  /** Distance to the text in points. */
  readonly space: number;
}

export interface Borders {
  readonly top?: BorderSpec;
  readonly bottom?: BorderSpec;
  readonly left?: BorderSpec;
  readonly right?: BorderSpec;
  /** Between paragraphs of a bordered group, or between table rows. */
  readonly between?: BorderSpec;
  readonly insideH?: BorderSpec;
  readonly insideV?: BorderSpec;
}

/** Background shading; `fill` is "RRGGBB" or "auto" (none). */
export interface Shading {
  readonly fill: string;
  /** Pattern and pattern colour, kept for writing back. */
  readonly val?: string;
  readonly color?: string;
}

// ----- sections ------------------------------------------------------------------------------

/** Page geometry. Points. */
export interface PageSetup {
  readonly width: number;
  readonly height: number;
  readonly marginTop: number;
  readonly marginBottom: number;
  readonly marginLeft: number;
  readonly marginRight: number;
  /** Distance of the header from the top edge and of the footer from the bottom edge. */
  readonly header: number;
  readonly footer: number;
  readonly gutter: number;
}

export const A4: PageSetup = {
  width: 595.3, height: 841.9, marginTop: 70.85, marginBottom: 56.7, marginLeft: 70.85, marginRight: 70.85,
  header: 35.4, footer: 35.4, gutter: 0,
};

export interface Columns {
  readonly count: number;
  /** Space between columns (equal columns), points. */
  readonly space: number;
  /** Vertical line between columns. */
  readonly sep: boolean;
  /** Individual widths when the columns are not equal. */
  readonly cols?: readonly { readonly width: number; readonly space: number }[];
}

/** Story ids of the headers or footers of a section. */
export interface HeaderRefs {
  readonly default?: string;
  readonly first?: string;
  readonly even?: string;
}

/** A section: page geometry, columns, headers and footers, page numbering. */
export interface Section {
  readonly page: PageSetup;
  readonly columns: Columns;
  /** How the section starts (the type stored on the section that begins here). */
  readonly start: "nextPage" | "continuous" | "evenPage" | "oddPage" | "nextColumn";
  /** Different first page header/footer ("Erste Seite anders"). */
  readonly titlePage: boolean;
  readonly headers: HeaderRefs;
  readonly footers: HeaderRefs;
  /** Page number to restart at, and its format ("decimal", "upperRoman", "lowerRoman" …). */
  readonly pageStart?: number;
  readonly pageFormat?: string;
  /** Layout > Zeilennummern (w:lnNumType): every countBy-th line numbered from `start`. */
  readonly lineNumbers?: LineNumbers;
  /** Entwurf > Seitenränder (w:pgBorders). */
  readonly pageBorders?: PageBorders;
  /** Other sectPr children, kept as XML for writing back. */
  readonly extra?: readonly unknown[];
}

/** Seitenränder (page borders) of a section; `space` of each side is its distance in points. */
export interface PageBorders {
  readonly top?: BorderSpec;
  readonly left?: BorderSpec;
  readonly bottom?: BorderSpec;
  readonly right?: BorderSpec;
  /** Gemessen von: Seitenrand (page edge) or Text. */
  readonly offsetFrom: "page" | "text";
  /** Übernehmen für: all pages, only the first, all but the first. */
  readonly display?: "allPages" | "firstPage" | "notFirstPage";
}

/** Zeilennummern of a section (Word's dialog Zeilennummern). */
export interface LineNumbers {
  readonly countBy: number;
  /** Beginnen mit (1-based; the file stores it 0-based). */
  readonly start: number;
  /** Abstand vom Text in points; undefined = Auto. */
  readonly distance?: number;
  readonly restart: "newPage" | "newSection" | "continuous";
}

export const ONE_COLUMN: Columns = { count: 1, space: 35.4, sep: false };

export const DEFAULT_SECTION: Section = {
  page: A4, columns: ONE_COLUMN, start: "nextPage", titlePage: false, headers: {}, footers: {},
};

// ----- tables --------------------------------------------------------------------------------

export interface CellMargins {
  readonly left: number;
  readonly right: number;
  readonly top: number;
  readonly bottom: number;
}

export interface CellDef {
  /** Grid columns this cell spans (w:gridSpan). */
  readonly span: number;
  readonly vMerge?: "restart" | "continue";
  /** Preferred width in points (w:tcW, dxa); undefined = from the grid. */
  readonly width?: number;
  readonly shading?: Shading;
  readonly borders?: Borders;
  readonly vAlign?: "top" | "center" | "bottom";
  readonly margins?: Partial<CellMargins>;
  /** Other tcPr children (XML). */
  readonly extra?: readonly unknown[];
}

export interface RowDef {
  readonly cells: readonly CellDef[];
  /** Row height in points with its rule; undefined = automatic. */
  readonly height?: number;
  readonly heightRule?: "atLeast" | "exact";
  /** Repeat as header row on every page (Überschrift wiederholen). */
  readonly header?: boolean;
  /** Row may not break across pages. */
  readonly cantSplit?: boolean;
  /** Other trPr / tblPrEx content (XML). */
  readonly extra?: readonly unknown[];
}

/**
 * A table. Its cells' paragraphs are ordinary paragraphs of the story, in row-major order,
 * each pointing to the cell it belongs to (Para.cell). Like Word's own model, the text stays
 * one sequence, so caret, selection, undo and search work unchanged.
 */
export interface TableDef {
  /** Distinguishes adjacent tables. */
  readonly id: number;
  /** Grid column widths in points (w:tblGrid). */
  readonly grid: readonly number[];
  readonly rows: readonly RowDef[];
  readonly style?: string;
  readonly align?: "left" | "center" | "right";
  /** Indent from the margin in points (w:tblInd). */
  readonly indent?: number;
  readonly borders?: Borders;
  readonly shading?: Shading;
  readonly cellMargins?: Partial<CellMargins>;
  /** w:tblLook flags (first row, banded rows …), as hex or attributes. */
  readonly look?: { firstRow: boolean; lastRow: boolean; firstColumn: boolean; lastColumn: boolean; noHBand: boolean; noVBand: boolean };
  /** Preferred table width in points; undefined = automatic. */
  readonly width?: number;
  readonly fixed?: boolean;
  /** Other tblPr children (XML). */
  readonly extra?: readonly unknown[];
}

export interface CellRef {
  readonly table: TableDef;
  readonly row: number;
  /** Index of the cell within its row (not the grid column). */
  readonly col: number;
}

// ----- numbering -----------------------------------------------------------------------------

export interface NumLevel {
  readonly start: number;
  /** "decimal", "lowerLetter", "upperLetter", "lowerRoman", "upperRoman", "bullet", "none" … */
  readonly format: string;
  /** Level text with %1…%9 placeholders ("%1.", "%1.%2"), or the bullet character. */
  readonly text: string;
  readonly align: "left" | "center" | "right";
  readonly suffix: "tab" | "space" | "nothing";
  readonly pPr: ParaProps;
  readonly rPr: CharProps;
  /** Restart after this level (1-based; 0 = never); undefined = after any higher level. */
  readonly restart?: number;
  /** Show all levels as arabic numbers (w:isLgl). */
  readonly legal?: boolean;
  /** Original w:lvl XML. */
  readonly data?: unknown;
}

export interface AbstractNum {
  readonly id: number;
  readonly levels: readonly NumLevel[];
  /** Original XML, written back unless the definition changed. */
  readonly data?: unknown;
}

export interface NumInstance {
  readonly id: number;
  readonly abstractId: number;
  /** Per level: restart value and/or replaced level definition. */
  readonly overrides?: ReadonlyMap<number, { readonly start?: number; readonly level?: NumLevel }>;
  readonly data?: unknown;
}

export interface Numbering {
  readonly abstracts: ReadonlyMap<number, AbstractNum>;
  readonly nums: ReadonlyMap<number, NumInstance>;
  /** Parts of numbering.xml the model does not cover (picture bullets …), kept as XML. */
  readonly extra?: readonly unknown[];
}

export const EMPTY_NUMBERING: Numbering = { abstracts: new Map(), nums: new Map() };

// ----- floating objects and comments ---------------------------------------------------------

/** A picture or shape anchored to a paragraph and positioned on the page (wp:anchor). */
export interface FloatObject {
  /** Offset in the paragraph the object is anchored at. */
  readonly at: number;
  /** Original XML (w:r with w:drawing …), written back unchanged unless edited. */
  readonly data: unknown;
  /** Size in points. */
  readonly width: number;
  readonly height: number;
  readonly image?: ImageData;
  readonly x: { readonly rel: string; readonly align?: string; readonly offset: number };
  readonly y: { readonly rel: string; readonly align?: string; readonly offset: number };
  /** "front", "behind", "square", "tight", "through", "topAndBottom". */
  readonly wrap: string;
  /** "bothSides", "left", "right", "largest". */
  readonly side: string;
  /** Distance from the text: top, bottom, left, right (points). */
  readonly dist: readonly [number, number, number, number];
  readonly pic?: PictureFormat;
  readonly alt?: AltText;
  /** Mit Text verschieben (false: Position auf Seite fixieren). */
  readonly moveWithText?: boolean;
  /** Stacking order among floating objects (wp:anchor relativeHeight); higher is in front. */
  readonly z?: number;
  /** Shapes and text boxes (wps:wsp); their outline, fill and effects are in `pic`. */
  readonly shape?: ShapeInfo;
  /** A group (Gruppieren, wpg:wgp): its shapes and pictures inside the object's box. */
  readonly group?: GroupInfo;
}

/** A shape, text box or picture of a group; position and size in points within the group's box. */
export interface GroupMember {
  readonly x: number;
  readonly y: number;
  readonly width: number;
  readonly height: number;
  readonly shape?: ShapeInfo;
  readonly pic?: PictureFormat;
  readonly image?: ImageData;
  readonly alt?: AltText;
  /** The member's XML as read or as it was a single object (wps:wsp or pic:pic). */
  readonly data?: unknown;
}

export interface GroupInfo {
  readonly members: readonly GroupMember[];
}

/** A shape's text frame (wps:bodyPr): insets in points, vertical anchor, wrapping, autofit. */
export interface ShapeBody {
  readonly l: number;
  readonly t: number;
  readonly r: number;
  readonly b: number;
  readonly anchor: "t" | "ctr" | "b";
  readonly wrap: boolean;
  /** "shape": Form dem Text anpassen (spAutoFit); "normal": shrink text (normAutofit). */
  readonly autofit: "none" | "shape" | "normal";
  /** Textrichtung: "horz", "vert" (270°: "vert270"), as DrawingML. */
  readonly vert?: string;
}

/**
 * A shape or text box: what DrawingML's shape style (wps:style) gives it, resolved against the
 * document theme, and its text. Geometry, fill, outline and effects set on the shape itself
 * are in the float's `pic`; where those are missing the style's apply.
 */
export interface ShapeInfo {
  readonly kind: "shape" | "textbox" | "connector";
  /** Adjust values of the preset geometry (a:avLst), by name. */
  readonly adjust?: Readonly<Record<string, number>>;
  readonly styleFill?: import("./picture").Fill;
  readonly styleLine?: { readonly color: string; readonly width: number };
  readonly styleShadow?: import("./picture").PictureShadow;
  /** Text colour the style gives (fontRef), e.g. white on Word's blue default shape. */
  readonly fontColor?: string;
  /** The text story ("tx:3"), when the shape has text. */
  readonly story?: string;
  /**
   * Verknüpfte Textfelder: the chain's id and this box's place in it (0: the first, which owns
   * the text); later boxes name the chain's story, through which the text flows on.
   */
  readonly chain?: { readonly id: number; readonly seq: number; readonly story?: string };
  readonly body: ShapeBody;
  /** wps:style as XML, written back unless a Formenart replaced it. */
  readonly styleXml?: string;
  /** Custom geometry (Freihandform, a:custGeom) in the definition syntax of geometry.ts. */
  readonly custom?: string;
}

export const DEFAULT_SHAPE_BODY: ShapeBody = { l: 7.2, t: 3.6, r: 7.2, b: 3.6, anchor: "ctr", wrap: true, autofit: "none" };

export interface ImageData {
  readonly bytes: Uint8Array;
  /** MIME type ("image/png" …). */
  readonly type: string;
}

export interface CommentMeta {
  readonly id: string;
  readonly author: string;
  readonly initials: string;
  readonly date: string;
  /** Id of the comment this one replies to. */
  readonly parent?: string;
  /** Marked as resolved (erledigt). */
  readonly done?: boolean;
  readonly data?: unknown;
}

export interface TabStop {
  /** Position in points from the left margin. */
  readonly pos: number;
  /** "num": the tab stop a list level sets after its number (acts like "left"). */
  readonly align: "left" | "center" | "right" | "decimal" | "bar" | "clear" | "num";
  readonly leader?: "none" | "dot" | "hyphen" | "underscore" | "heavy" | "middleDot";
}

/** Paragraph formatting. Lengths in points. */
export interface ParaProps {
  readonly style?: string;
  readonly align?: Align;
  readonly indLeft?: number;
  readonly indRight?: number;
  /** First-line indent; negative = hanging indent. */
  readonly indFirst?: number;
  readonly before?: number;
  readonly after?: number;
  readonly beforeAuto?: boolean;
  readonly afterAuto?: boolean;
  /** With lineRule "auto": multiple of single spacing (1 = single). Otherwise points. */
  readonly line?: number;
  readonly lineRule?: LineRule;
  readonly contextual?: boolean;
  readonly keepNext?: boolean;
  readonly keepLines?: boolean;
  readonly widow?: boolean;
  readonly pageBreakBefore?: boolean;
  /** Zeilennummern: Für aktuellen Absatz unterdrücken. */
  readonly suppressLineNumbers?: boolean;
  readonly tabs?: readonly TabStop[];
  readonly outline?: number;
  /** List: numbering instance (0 = explicitly none) and level (0-based). */
  readonly numId?: number;
  readonly ilvl?: number;
  readonly borders?: Borders;
  readonly shading?: Shading;
  /** Paragraph properties the engine does not interpret, kept for writing back. */
  readonly extra?: readonly unknown[];
}

export interface Run {
  readonly len: number;
  readonly props: CharProps;
}

/** Something inline the engine shows but does not edit inside: a field, a link, an image … */
export interface InlineObject {
  readonly kind: "field" | "link" | "image" | "note" | "revision" | "math" | "other";
  /** Opaque payload for the file format layer. */
  readonly data: unknown;
  /** Text shown in place of the object (field result, link text …). */
  readonly display: string;
  /** Size in points (images, equations). */
  readonly width?: number;
  readonly height?: number;
  /** Depth below the baseline in points (equations). */
  readonly depth?: number;
  readonly image?: ImageData;
  /** Field code ("PAGE", "NUMPAGES", "TOC \\o …"): page fields are shown per page. */
  readonly instr?: string;
  /** Foot- or endnote reference: the note's story id ("fn:2", "en:1"). */
  readonly note?: string;
  /** Equation as LaTeX: derived from `math` (a key for caches) or read from a template. */
  readonly latex?: string;
  /** Equation: its tree (what is edited and written as OMML). */
  readonly math?: import("./math/ast").Seq;
  /** An equation alone in its paragraph shown inline (an m:oMath without m:oMathPara). */
  readonly mathInline?: boolean;
  /** Display equation: its own alignment (m:oMathParaPr/m:jc); else the document's. */
  readonly mathJc?: "left" | "right" | "center" | "centerGroup";
  /** Link target (hyperlinks). */
  readonly href?: string;
  /** Font of a symbol object (w:sym, e.g. Wingdings); its display is in that font. */
  readonly font?: string;
  /** Pictures: Bildformat settings (corrections, colour, crop, border, effects …). */
  readonly pic?: PictureFormat;
  /** Pictures: Alternativtext. */
  readonly alt?: AltText;
  /** Shapes in line with the text (wps:wsp in wp:inline). */
  readonly shape?: ShapeInfo;
}

/** Alternativtext of a picture or shape (wp:docPr descr/title, decorative). */
export interface AltText {
  readonly descr?: string;
  readonly title?: string;
  readonly decorative?: boolean;
  /** The object's name (wp:docPr name: "Rechteck 3", "Grafik 1"), shown in the Auswahlbereich. */
  readonly name?: string;
  /** Hidden in the Auswahlbereich (wp:docPr hidden): not drawn, text flows as without it. */
  readonly hidden?: boolean;
}

/** A zero-width marker kept at an offset (bookmark, comment range …). */
export interface Anchor {
  readonly at: number;
  readonly data: unknown;
  /** Comment ranges: which edge of which comment. */
  readonly comment?: { readonly id: string; readonly edge: "start" | "end" | "ref" };
  /**
   * Start and end of a field whose result spans paragraphs (Inhaltsverzeichnis) or of a
   * hyperlink; the file format layer writes only complete pairs.
   */
  readonly field?: { readonly id: string; readonly edge: "start" | "end"; readonly instr?: string };
  /** A hyperlink's edges; the start says where it goes (web address, file, mailto:, or a bookmark in the document). */
  readonly link?: { readonly id: string; readonly edge: "start" | "end"; readonly href?: string; readonly anchor?: string; readonly tip?: string; readonly frame?: string };
  /** Bookmark name (bookmarkStart), for cross references and the table of contents. */
  readonly bookmark?: string;
}

/** A block the engine does not edit (table, multi-paragraph field …), shown as a placeholder. */
export interface BlockObject {
  readonly kind: "table" | "field" | "other";
  readonly data: unknown;
  readonly label: string;
  /** Text lines shown read-only inside the placeholder. */
  readonly lines: readonly string[];
  /** Estimated height in points (used for pagination until the block is supported). */
  readonly height: number;
}

export interface Para {
  /** Text; "\t" tab, "\u000b" line break, "\u000c" page break, "\ufffc" object. */
  readonly text: string;
  /** Formatting runs covering `text` exactly (empty for an empty paragraph). */
  readonly runs: readonly Run[];
  readonly props: ParaProps;
  /** Formatting of the paragraph mark (height of empty paragraphs, typing at the start). */
  readonly mark: CharProps;
  /** One entry per "\ufffc" in `text`, in order. */
  readonly objects?: readonly InlineObject[];
  readonly anchors?: readonly Anchor[];
  /** When set, the whole paragraph is this read-only block. */
  readonly block?: BlockObject;
  /** Opaque content that preceded this paragraph in the file (body-level bookmarks …). */
  readonly leading?: readonly unknown[];
  /** The section that ends with this paragraph (Abschnittswechsel). */
  readonly sect?: Section;
  /** The table cell this paragraph is in. */
  readonly cell?: CellRef;
  /** Pictures and shapes anchored to this paragraph. */
  readonly floats?: readonly FloatObject[];
  /** The content control (block-level w:sdt) the paragraph is in: shared by its paragraphs. */
  readonly group?: ParaGroup;
  /** The multi-paragraph field (Inhaltsverzeichnis) whose result this paragraph is part of. */
  readonly field?: string;
}

/** Paragraphs wrapped together in the file (a content control around a table of contents …). */
export interface ParaGroup {
  readonly id: number;
  readonly kind: "sdt";
  /** XML before the first and after the last paragraph. */
  readonly open: unknown;
  readonly close: unknown;
}

export const OBJ = "\ufffc";
export const LINE_BREAK = "\u000b";
export const PAGE_BREAK = "\u000c";
/** Column break (Spaltenumbruch); a control character Word does not use in text. */
export const COLUMN_BREAK = "\u000e";
export const EMPTY: CharProps = Object.freeze({});

// ----- interning --------------------------------------------------------------------------

const interned = new Map<string, CharProps>();

function propsKey(p: CharProps): string {
  const keys = Object.keys(p).filter((k) => (p as Record<string, unknown>)[k] !== undefined).sort();
  return keys.map((k) => {
    const v = (p as Record<string, unknown>)[k];
    if (k === "rev") { const r = v as Revision; return `rev=${r.kind}:${r.id}:${r.author}:${r.date}`; }
    return `${k}=${k === "extra" ? (v as unknown[]).map(String).join("\u0001") : String(v)}`;
  }).join("|");
}

/** Returns the canonical (shared, frozen) object for these properties, so runs compare by identity. */
export function intern(p: CharProps): CharProps {
  const clean: Record<string, unknown> = {};
  for (const [k, v] of Object.entries(p)) if (v !== undefined) clean[k] = v;
  const key = propsKey(clean as CharProps);
  let hit = interned.get(key);
  if (!hit) {
    hit = Object.freeze(clean) as CharProps;
    interned.set(key, hit);
  }
  return hit;
}

// ----- paragraph helpers ------------------------------------------------------------------

export function makePara(text: string, props: ParaProps = {}, runProps: CharProps = EMPTY, mark: CharProps = EMPTY): Para {
  return { text, runs: text ? [{ len: text.length, props: intern(runProps) }] : [], props, mark: intern(mark) };
}

/** Merges neighbouring runs with identical properties and drops empty ones. */
export function normalizeRuns(runs: readonly Run[]): Run[] {
  const out: Run[] = [];
  for (const r of runs) {
    if (r.len <= 0) continue;
    const last = out[out.length - 1];
    if (last && last.props === r.props) out[out.length - 1] = { len: last.len + r.len, props: r.props };
    else out.push(r);
  }
  return out;
}

/** Character properties of the character at `offset` (the paragraph mark's at the end). */
export function propsAt(p: Para, offset: number): CharProps {
  let pos = 0;
  for (const r of p.runs) {
    if (offset < pos + r.len) return r.props;
    pos += r.len;
  }
  return p.mark;
}

/** Formatting new text takes at `offset`: that of the character before, as in Word. */
export function insertionProps(p: Para, offset: number): CharProps {
  if (!p.text.length) return p.mark;
  if (offset > 0) {
    // After a line break or object, continue with the following text's formatting.
    const before = p.text[offset - 1];
    if (before !== OBJ) return propsAt(p, offset - 1);
  }
  return offset < p.text.length ? propsAt(p, offset) : propsAt(p, offset - 1);
}

function sliceRuns(runs: readonly Run[], from: number, to: number): Run[] {
  const out: Run[] = [];
  let pos = 0;
  for (const r of runs) {
    const a = Math.max(from, pos), b = Math.min(to, pos + r.len);
    if (b > a) out.push({ len: b - a, props: r.props });
    pos += r.len;
    if (pos >= to) break;
  }
  return out;
}

function countObj(text: string, from: number, to: number): number {
  let n = 0;
  for (let i = from; i < to; i++) if (text.charCodeAt(i) === 0xfffc) n++;
  return n;
}

/**
 * The part [from, to) of a paragraph, keeping runs, objects and anchors aligned. Anchors at
 * `to` belong to the piece only at the paragraph end or with `inclusiveEnd` (a head piece
 * keeps the markers at the cut, the tail piece does not repeat them).
 */
export function slicePara(p: Para, from: number, to: number, inclusiveEnd = false, exclusiveStart = false): Para {
  const before = countObj(p.text, 0, from);
  const inside = countObj(p.text, from, to);
  const objects = p.objects?.slice(before, before + inside);
  const anchors = p.anchors?.filter((a) => (exclusiveStart ? a.at > from : a.at >= from) && (a.at < to || (a.at === to && (inclusiveEnd || to === p.text.length))))
    .map((a) => ({ ...a, at: a.at - from }));
  const floats = p.floats?.filter((f) => f.at >= from && (f.at < to || (f.at === to && to === p.text.length))).map((f) => ({ ...f, at: f.at - from }));
  return {
    ...p,
    text: p.text.slice(from, to),
    runs: sliceRuns(p.runs, from, to),
    objects: objects?.length ? objects : undefined,
    anchors: anchors?.length ? anchors : undefined,
    floats: floats?.length ? floats : undefined,
    // The section break belongs to the paragraph mark, which only the end piece keeps.
    sect: to >= p.text.length ? p.sect : undefined,
  };
}

/** Joins two paragraph pieces; the result keeps `a`'s paragraph properties and `b`'s mark. */
export function joinParas(a: Para, b: Para): Para {
  const shift = a.text.length;
  const objects = [...(a.objects ?? []), ...(b.objects ?? [])];
  const anchors = [...(a.anchors ?? []), ...(b.anchors ?? []).map((x) => ({ ...x, at: x.at + shift }))];
  const floats = [...(a.floats ?? []), ...(b.floats ?? []).map((x) => ({ ...x, at: x.at + shift }))];
  return {
    ...a,
    text: a.text + b.text,
    runs: normalizeRuns([...a.runs, ...b.runs]),
    mark: b.mark,
    objects: objects.length ? objects : undefined,
    anchors: anchors.length ? anchors : undefined,
    floats: floats.length ? floats : undefined,
    // Joining removes a's paragraph mark; the section break of b's mark survives (Word).
    sect: b.sect,
  };
}

/** Text with one formatting, as a paragraph piece (for inserting). */
export function textPiece(text: string, props: CharProps): Para {
  return { text, runs: text ? [{ len: text.length, props: intern(props) }] : [], props: {}, mark: EMPTY };
}

/** Visits every paragraph of a story with its index (helper for numbering and tables). */
export function cellKey(c: CellRef): string {
  return `${c.table.id}:${c.row}:${c.col}`;
}

/** Changes character properties on [from, to) with `change(old) => new`. */
export function mapRuns(p: Para, from: number, to: number, change: (old: CharProps) => CharProps): Para {
  const out: Run[] = [];
  let pos = 0;
  for (const r of p.runs) {
    const start = pos, end = pos + r.len;
    pos = end;
    const a = Math.max(from, start), b = Math.min(to, end);
    if (b <= a) { out.push(r); continue; }
    if (a > start) out.push({ len: a - start, props: r.props });
    out.push({ len: b - a, props: intern(change(r.props)) });
    if (end > b) out.push({ len: end - b, props: r.props });
  }
  return { ...p, runs: normalizeRuns(out) };
}

/** Replaces the text of [from, to) with same-length text (change case), keeping formatting. */
export function replaceSameLength(p: Para, from: number, text: string): Para {
  return { ...p, text: p.text.slice(0, from) + text + p.text.slice(from + text.length) };
}

/** Checks the invariants; used by tests. */
export function validPara(p: Para): boolean {
  const total = p.runs.reduce((n, r) => n + r.len, 0);
  return total === p.text.length && countObj(p.text, 0, p.text.length) === (p.objects?.length ?? 0) &&
    p.runs.every((r) => r.len > 0);
}
