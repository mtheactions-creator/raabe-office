// The editor: a hidden textarea receives keyboard and IME input, mouse handling implements
// Word's selection rules, and all edits go through the document model. The shell calls the
// public commands below and listens for changes with onChange().

import { DEFAULT_MATH, Doc, MAIN, cmp, maxPos, minPos, type DocSettings, type DocSnapshot, type DocState, type MathSettings, type Pos } from "./doc";
import { Fonts, PT } from "./fonts";
import { layoutDocument, type FlowInput, type Layout } from "./flow";
import { Layouter, type LayoutContext } from "./layout";
import { markInputHandled } from "./probe";
import {
  EMPTY, LINE_BREAK, OBJ, PAGE_BREAK, insertionProps, intern, joinParas, makePara, propsAt, type Align, type CharProps, type LineRule,
  type Para, type ParaProps, type Section,
} from "./model";
import { GERMAN_NAMES, builtinStyle, defaultStyleSheet, type ResolvedChar, type ResolvedPara, type StyleDef, type Theme } from "./styles";
import { View, imageUrl, storyIdOf, type RenderOptions, type Selection, type TextMark } from "./view";
import * as E from "./edits";
import { fromHtml, toHtml } from "./html";
import { imageInfo, naturalSizePt } from "./images";

/**
 * A picture scaled down to `ppi` at its size on the page (points), optionally with its
 * cropped parts removed; null when nothing would change.
 */
async function resample(image: ImageData, w: number, h: number, ppi: number, cut?: { l: number; t: number; r: number; b: number }, crop?: { l: number; t: number; r: number; b: number }): Promise<ImageData | null> {
  const url = URL.createObjectURL(new Blob([image.bytes as BlobPart], { type: image.type }));
  try {
    const el = await new Promise<HTMLImageElement>((resolve, reject) => { const i = new Image(); i.onload = () => resolve(i); i.onerror = reject; i.src = url; });
    const c = crop ?? { l: 0, t: 0, r: 0, b: 0 };
    const k = cut ?? { l: 0, t: 0, r: 0, b: 0 };
    // The whole picture's size on the page.
    const fullW = w / Math.max(0.01, 1 - c.l - c.r), fullH = h / Math.max(0.01, 1 - c.t - c.b);
    const sx = Math.max(0, k.l) * el.naturalWidth, sy = Math.max(0, k.t) * el.naturalHeight;
    const sw = el.naturalWidth * (1 - Math.max(0, k.l) - Math.max(0, k.r)), sh = el.naturalHeight * (1 - Math.max(0, k.t) - Math.max(0, k.b));
    const showW = cut ? w : fullW, showH = cut ? h : fullH;
    const tw = Math.round(Math.min(sw, (showW / 72) * ppi)), th = Math.round(Math.min(sh, (showH / 72) * ppi));
    if (!cut && tw >= el.naturalWidth - 1) return null;
    const cv = document.createElement("canvas");
    cv.width = Math.max(1, tw); cv.height = Math.max(1, th);
    const g = cv.getContext("2d")!;
    g.imageSmoothingQuality = "high";
    g.drawImage(el, sx, sy, sw, sh, 0, 0, cv.width, cv.height);
    const type = image.type === "image/jpeg" ? "image/jpeg" : "image/png";
    const blob = await new Promise<Blob | null>((resolve) => cv.toBlob(resolve, type, 0.9));
    return blob ? { bytes: new Uint8Array(await blob.arrayBuffer()), type } : null;
  } catch {
    return null;
  } finally {
    URL.revokeObjectURL(url);
  }
}

/** The visible text of an HTML fragment. */
function htmlText(html: string): string {
  return new DOMParser().parseFromString(html, "text/html").body.textContent ?? "";
}
import * as R from "./revisions";
import { INSTANT_TRIGGERS, TRIGGERS, borderLine, instantFix, smartQuote, wordFixes, type Fix } from "./autocorrect";
import * as S from "./school";
import * as Refs from "./refs";
import * as Bib from "./bibliography";
import { sortOrder, type SortSpec } from "./sort";
import { COLUMN_BREAK, type CellDef, type ImageData, type RowDef, type TableDef } from "./model";
import { addList, bulletLevels as bulletLevelsOf, levelOf, listLabels, numberLevels as numberLevelsOf, numberingOf, restartList } from "./numbering";
import { DEFAULT_SHAPE_BODY, type AltText, type BorderSpec, type Borders, type FloatObject, type InlineObject, type NumLevel, type ShapeInfo } from "./model";
import { SHAPE_NAMES, shadeLinear, shapeElement, shapeGeometry, shapeTextRect } from "./shape";
import { parseLatex, seqToLatex, type Seq } from "./math/ast";
import { presetAdjust, presetGeometry } from "./geometry";
import { effectExtent, type PictureFormat } from "./picture";
import { countWords, isWordChar, nextTokenStart, prevTokenStart, sentenceAt, tokenAt } from "./words";

/** Formatting text typed at a position gets: as in Word, never hidden (hidden text is not shown). */
function typedProps(p: Para, o: number): CharProps {
  const c = insertionProps(p, o);
  return c.hidden ? intern({ ...c, hidden: undefined }) : c;
}

type Granularity = "char" | "word" | "para";

/** Length of the character before / after an offset: 2 for an emoji's surrogate pair. */
function charBefore(text: string, o: number): number {
  const lo = text.charCodeAt(o - 1), hi = text.charCodeAt(o - 2);
  return lo >= 0xdc00 && lo <= 0xdfff && hi >= 0xd800 && hi <= 0xdbff ? 2 : 1;
}
function charAfter(text: string, o: number): number {
  const hi = text.charCodeAt(o), lo = text.charCodeAt(o + 1);
  return hi >= 0xd800 && hi <= 0xdbff && lo >= 0xdc00 && lo <= 0xdfff ? 2 : 1;
}

interface Snapshot {
  state: DocSnapshot;
  /** Story with the caret and its layout key. */
  story: string;
  key: string;
  anchor: Pos;
  head: Pos;
}

export interface ChangeEvent {
  /** The document content changed (not just the selection). */
  content: boolean;
}

export interface EditorStats {
  pageCount: number;
  /** Zero-based page of the caret. */
  caretPage: number;
}

/** Character formatting of the selection; `undefined` where the selection is mixed. */
export interface CharState {
  font?: string;
  size?: number;
  bold?: boolean;
  italic?: boolean;
  underline?: string;
  strike?: boolean;
  dstrike?: boolean;
  superscript?: boolean;
  subscript?: boolean;
  caps?: boolean;
  smallCaps?: boolean;
  hidden?: boolean;
  color?: string;
  highlight?: string;
  charStyle?: string;
}

export type CaseMode = "sentence" | "lower" | "upper" | "title" | "toggle";

/** Copied formatting (Format übertragen). */
export interface CopiedFormat {
  char: CharProps;
  para: ParaProps | null;
}

/** Word's font size steps for "Schriftgrad vergrößern/verkleinern". */
export const FONT_STEPS = [8, 9, 10, 10.5, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72];
/** Schriftgrad vergrößern/verkleinern skip 10.5 (measured: 11 shrinks to 10). */
const GROW_STEPS = FONT_STEPS.filter((x) => x !== 10.5);

let clipboardFragment: {
  text: string; paras: Para[]; source: Editor;
  /** A copied floating picture or shape, with the text stories of its shapes. */
  float?: { obj: FloatObject; stories: Map<string, Para[]> };
} | null = null;

/**
 * Word's pointer in the selection bar (left margin): the Windows arrow, mirrored so it points
 * up and to the right. Drawn at the display's scale so it is as sharp as the system cursors.
 */
const MIRRORED_ARROW = (() => {
  const svg = (px: number) => `url("data:image/svg+xml,${encodeURIComponent(
    `<svg xmlns="http://www.w3.org/2000/svg" width="${px}" height="${px}" viewBox="0 0 32 32">` +
    '<path d="M19.5 1.5V18.5L15.5 14.5L12.5 21.5L10 20.5L13 13.5H7.5Z" fill="white" stroke="black" stroke-width="1" stroke-linejoin="round"/></svg>')}")`;
  const set = `image-set(${svg(32)} 1x, ${svg(48)} 1.5x, ${svg(64)} 2x) 19 1, default`;
  return typeof CSS !== "undefined" && CSS.supports("cursor", set) ? set : `${svg(32)} 19 1, default`;
})();

const BRUSH_CURSOR = `url("data:image/svg+xml,${encodeURIComponent(
  '<svg xmlns="http://www.w3.org/2000/svg" width="20" height="24"><path d="M2 2v20M2 2h3M2 22h3" stroke="black" stroke-width="1.5" fill="none"/>' +
  '<path d="M8 4h9v5H8z" fill="white" stroke="black"/><path d="M12 9v4h1v6" stroke="black" fill="none"/></svg>')}") 2 12, text`;

/** A tool of the Formen gallery: preset geometry (or a freeform tool) and what it starts with. */
export interface DrawTool {
  readonly geometry: string;
  readonly kind: ShapeInfo["kind"];
  readonly pic?: PictureFormat;
  readonly adjust?: Record<string, number>;
  /** Kurve, Freihandform: Form, Freihand: Skizze. */
  readonly freeform?: "curve" | "polygon" | "scribble";
}

/** SVG path data through points: straight segments, or a smooth curve (Catmull-Rom). */
function freeformPath(pts: readonly [number, number][], smooth: boolean): string {
  if (!pts.length) return "";
  let d = `M${pts[0][0]} ${pts[0][1]}`;
  for (let i = 1; i < pts.length; i++) {
    if (!smooth) { d += ` L${pts[i][0]} ${pts[i][1]}`; continue; }
    const p0 = pts[Math.max(0, i - 2)], p1 = pts[i - 1], p2 = pts[i], p3 = pts[Math.min(pts.length - 1, i + 1)];
    d += ` C${p1[0] + (p2[0] - p0[0]) / 6} ${p1[1] + (p2[1] - p0[1]) / 6} ${p2[0] - (p3[0] - p1[0]) / 6} ${p2[1] - (p3[1] - p1[1]) / 6} ${p2[0]} ${p2[1]}`;
  }
  return d;
}

export class Editor {
  doc: Doc;
  layout!: Layout;
  sel: Selection;
  readonly view: View;
  readonly layouter: Layouter;
  private readonly input: HTMLTextAreaElement;
  private goalX: number | null = null;
  private undoStack: Snapshot[] = [];
  private redoStack: Snapshot[] = [];
  private typing = false;
  private drag: { granularity: Granularity; origin: [Pos, Pos] } | null = null;
  private listeners: ((e: ChangeEvent) => void)[] = [];
  private lastParas: Para[] | null = null;
  private lastStories: ReadonlyMap<string, Para[]> | null = null;
  /** Layout key of the story with the caret ("main", "fn:1", "hdr:2@0" …). */
  storyKey = MAIN;
  /** Measures equations (LaTeX → points); set by the host page. */
  measureMath: LayoutContext["measureMath"];
  /** Formatting for the next typed text at a collapsed caret (e.g. after Strg+Umschalt+F). */
  private pending: CharProps | null = null;
  private painter: { fmt: CopiedFormat; sticky: boolean } | null = null;

  constructor(readonly workspace: HTMLElement, state: Partial<DocState>, readonly fonts: Fonts) {
    this.doc = new Doc(state);
    this.sel = { anchor: { p: 0, o: 0 }, head: { p: 0, o: 0 }, headUpstream: false };
    this.layouter = new Layouter(this.context());
    this.view = new View(workspace, fonts);
    this.installMarks();
    this.requireFonts();

    this.input = document.createElement("textarea");
    this.input.setAttribute("autocomplete", "off");
    this.input.setAttribute("autocorrect", "off");
    this.input.spellcheck = false;
    this.input.style.cssText = "position:absolute;z-index:-1;opacity:0;width:1px;height:1px;padding:0;" +
      "border:0;resize:none;overflow:hidden;pointer-events:none;white-space:pre";
    this.view.container.append(this.input);

    this.input.addEventListener("beforeinput", (e) => this.onBeforeInput(e));
    this.input.addEventListener("compositionend", (e) => this.onCompositionEnd(e));
    this.input.addEventListener("keydown", (e) => this.onKeyDown(e));
    this.input.addEventListener("copy", (e) => this.onCopy(e, false));
    this.input.addEventListener("cut", (e) => this.onCopy(e, true));
    this.input.addEventListener("paste", (e) => this.onPaste(e));
    workspace.addEventListener("mousedown", (e) => this.onMouseDown(e));
    workspace.addEventListener("mousemove", (e) => { if (!e.buttons) this.hoverCursor(e); });
    workspace.addEventListener("contextmenu", (e) => { if (!this.view.markupLayer.contains(e.target as Node)) this.onContext(e); });
    this.update();
  }

  private context(): LayoutContext {
    return {
      styles: this.doc.styles, fonts: this.fonts, defaultTab: this.doc.settings.defaultTab, measureMath: this.measureMath, markup: this.markupMode,
      mathJc: this.doc.settings.math?.defJc,
    };
  }

  /** Asks the platform for metrics of every font the document uses. */
  private requireFonts(): void {
    const families = new Set<string>([this.doc.styles.theme.minorFont, this.doc.styles.theme.majorFont]);
    const scan = (paras: readonly Para[]) => {
      for (const p of paras) {
        const rp = this.doc.styles.resolvePara(p.props);
        for (const r of p.runs) families.add(this.doc.styles.resolveChar(rp.style, r.props).font);
        families.add(this.doc.styles.resolveChar(rp.style, p.mark).font);
      }
    };
    scan(this.doc.main);
    for (const paras of this.doc.stories.values()) scan(paras);
    for (const a of this.doc.numbering.abstracts.values()) for (const l of a.levels) if (l?.rPr.font) families.add(l.rPr.font);
    this.fonts.require(families);
  }

  /** Re-lays out everything (fonts arrived, styles changed). */
  relayout(): void {
    this.layouter.setContext(this.context());
    this.update();
  }

  // ----- equations ------------------------------------------------------------------------

  /** How the host measures (points) and draws equations; MathML in Cambria Math in the app. */
  useMath(measure: NonNullable<LayoutContext["measureMath"]>, render: (o: InlineObject, sizePx: number, display: boolean) => HTMLElement): void {
    this.measureMath = measure;
    this.view.measureMath = measure;
    this.view.renderMath = render;
    this.relayout();
  }

  /** Called when the user clicks into an equation or moves the caret into one (host opens its editor). */
  onEditMath: ((pos: Pos, enter: "start" | "end" | "click", point?: { x: number; y: number }) => void) | null = null;

  /** The formula editor takes keys, typed text and mouse presses first while it is open. */
  interceptKey: ((e: KeyboardEvent) => boolean) | null = null;
  interceptText: ((text: string) => boolean) | null = null;
  interceptMouse: ((e: MouseEvent) => boolean) | null = null;

  // ----- bookmarks, cross references, captions, fields ------------------------------------

  /** Einfügen > Textmarke: the document's bookmarks (hidden ones start with "_"). */
  bookmarks(): Refs.Bookmark[] {
    return Refs.bookmarks(this.doc);
  }

  /** Textmarke > Hinzufügen: a bookmark over the selection (main text). */
  addBookmark(name: string): boolean {
    if (this.storyKey !== MAIN || !Refs.validBookmarkName(name)) return false;
    const [from, to] = this.range();
    this.snapshot();
    Refs.addBookmark(this.doc, name, from, to);
    this.update();
    return true;
  }

  deleteBookmark(name: string): void {
    this.snapshot();
    Refs.deleteBookmark(this.doc, name);
    this.update();
  }

  /** Textmarke > Gehe zu: the bookmark's text is selected. */
  goToBookmark(name: string): void {
    const b = Refs.bookmarks(this.doc).find((x) => x.name === name);
    if (!b) return;
    this.setSel(b.start, b.end);
    this.update({ scroll: true });
  }

  /** What a cross reference can point to (Querverweis > Verweistyp). */
  referenceTargets(kind: "heading" | "bookmark" | "footnote" | "endnote" | "numbered" | { caption: string }): { label: string; key: string; level?: number }[] {
    const labels = listLabels(this.doc.main, this.doc.numbering, this.doc.styles);
    if (kind === "heading") {
      return this.doc.main.flatMap((p, i) => {
        const level = p.props.outline ?? this.doc.styles.outlineLevel(this.doc.styles.resolvePara(p.props).style);
        if (level === undefined || level > 8 || !p.text.trim()) return [];
        const label = labels[i]?.text ? `${labels[i]!.text} ` : "";
        return [{ label: `${label}${p.text.replace(/\ufffc/g, "")}`, key: String(i), level: level + 1 }];
      });
    }
    if (kind === "numbered") {
      return this.doc.main.flatMap((p, i) => (labels[i]?.text && p.text.trim() ? [{ label: `${labels[i]!.text} ${p.text.replace(/\ufffc/g, "")}`, key: String(i) }] : []));
    }
    if (kind === "bookmark") return Refs.bookmarks(this.doc).filter((b) => !b.name.startsWith("_")).map((b) => ({ label: b.name, key: b.name }));
    if (kind === "footnote" || kind === "endnote") {
      const out: { label: string; key: string }[] = [];
      this.doc.main.forEach((p, i) => {
        let k = 0;
        for (let o = 0; o < p.text.length; o++) {
          if (p.text[o] !== OBJ) continue;
          const obj = p.objects?.[k++];
          if (!obj?.note || !obj.note.startsWith(kind === "footnote" ? "fn:" : "en:")) continue;
          const text = (this.doc.stories.get(obj.note) ?? []).map((x) => x.text).join(" ").replace(/\ufffc/g, "").trim();
          out.push({ label: `${this.layout.noteNumbers.get(obj.note) ?? ""} ${text}`, key: `${i}:${o}` });
        }
      });
      return out;
    }
    return Refs.captions(this.doc, kind.caption).map((c) => ({ label: c.text, key: String(c.p) }));
  }

  /**
   * Querverweis einfügen: a REF, PAGEREF or NOTEREF field to a hidden _Ref bookmark around the
   * target, with Word's result. `what`: text, page, number, aboveBelow, noteNumber; for
   * captions whole, labelNumber, captionText.
   */
  insertCrossReference(kind: Parameters<Editor["referenceTargets"]>[0], key: string, what: string, asLink: boolean, aboveBelow: boolean): void {
    if (this.storyKey !== MAIN) return;
    this.snapshot();
    let from: Pos, to: Pos;
    if (kind === "bookmark") {
      const b = Refs.bookmarks(this.doc).find((x) => x.name === key);
      if (!b) return;
      from = b.start; to = b.end;
    } else if (kind === "footnote" || kind === "endnote") {
      const [p, o] = key.split(":").map(Number);
      from = { p, o }; to = { p, o: o + 1 };
    } else {
      const p = Number(key);
      const para = this.doc.main[p];
      from = { p, o: 0 }; to = { p, o: para.text.length };
      if (typeof kind === "object") {
        const seq = Refs.seqFields(this.doc, kind.caption).find((x) => x.p === p);
        if (seq && what === "labelNumber") to = { p, o: seq.o + 1 };
        if (seq && what === "captionText") {
          const rest = para.text.slice(seq.o + 1);
          const skip = /^[\s:.\-–—]*/.exec(rest)?.[0].length ?? 0;
          from = { p, o: seq.o + 1 + skip };
        }
      }
    }
    const name = kind === "bookmark" ? key : Refs.refBookmark(this.doc, from, to);
    const sw = `${asLink ? " \\h" : ""}${aboveBelow ? " \\p" : ""}`;
    const instr = what === "page" ? `PAGEREF ${name}${sw}` : what === "noteNumber" ? `NOTEREF ${name}${sw}`
      : what === "number" ? `REF ${name} \\r${sw}` : what === "aboveBelow" ? `REF ${name} \\p${asLink ? " \\h" : ""}` : `REF ${name}${sw}`;
    const [a, b] = this.range();
    const end = E.insertObject(this.doc, a, b, E.fieldObject(instr, ""), this.typingProps());
    this.setSel(end, end);
    this.update();
    this.refreshFields();
    this.update({ scroll: true });
  }

  /**
   * Beschriftung einfügen: "Abbildung 1: Text" in the style Beschriftung, the number a SEQ
   * field, as its own paragraph above or below the selected picture/table (else at the caret's
   * paragraph). Numbers of all captions of the label are updated.
   */
  insertCaption(label: string, text: string, below = true, withoutLabel = false, format = "ARABIC"): void {
    if (this.storyKey !== MAIN) return;
    this.snapshot();
    this.typing = false;
    const style = this.styleId("caption");
    this.ensureStyle(style);
    const at = this.sel.head;
    const paras = [...this.doc.main];
    const p = Math.min(at.p, paras.length - 1);
    const lead = withoutLabel ? "" : `${label} `;
    const tail = text ? (text.startsWith(":") || text.startsWith(" ") ? text : `: ${text}`) : "";
    const textAll = `${lead}${OBJ}${tail}`;
    const caption: Para = {
      text: textAll, runs: [{ len: textAll.length, props: EMPTY }], props: { style }, mark: EMPTY,
      objects: [E.fieldObject(`SEQ ${label} \\* ${format}`, "1")],
    };
    paras.splice(below ? p + 1 : p, 0, caption);
    this.doc.setStory(MAIN, paras);
    const cp = below ? p + 1 : p;
    this.setSel({ p: cp, o: textAll.length }, { p: cp, o: textAll.length });
    this.refreshFields();
    this.update({ scroll: true });
  }

  /** Felder aktualisieren (F9): SEQ numbers, cross references, page references. */
  updateAllFields(): boolean {
    this.snapshot();
    const changed = this.refreshFields();
    this.refreshCitations();
    this.update();
    return changed;
  }

  private refreshFields(): boolean {
    this.update();
    const labels = listLabels(this.doc.main, this.doc.numbering, this.doc.styles);
    return Refs.updateFields(this.doc, this.pageNumberOf, (story) => this.layout.noteNumbers.get(story) ?? null, (p) => labels[p]?.text ?? "");
  }

  // ----- Zitate und Literaturverzeichnis ---------------------------------------------------

  /** The document's sources and citation style (APA when it has none). */
  bibliography(): Bib.Bibliography {
    return this.doc.settings.bibliography ?? { sources: [], style: "APA" };
  }

  /** Quellen verwalten, Formatvorlage: new sources or style; citations and lists follow. */
  setBibliography(b: Bib.Bibliography): void {
    this.snapshot();
    this.doc.setSettings({ ...this.doc.settings, bibliography: b });
    this.refreshCitations();
    this.update();
  }

  /** The tags cited in the main text, in order of first citation. */
  citedTags(): string[] {
    const out: string[] = [];
    for (const p of this.doc.main) for (const o of p.objects ?? []) {
      const m = o.kind === "field" ? /^\s*CITATION\s+(\S+)/i.exec(o.instr ?? "") : null;
      if (m && !out.includes(m[1])) out.push(m[1]);
    }
    return out;
  }

  /** Zitat einfügen: a CITATION field for a source (\\l 1031: German), with Word's text. */
  insertCitation(tag: string, pages = ""): void {
    if (this.inBlock()) return;
    this.snapshot();
    const instr = `CITATION ${tag} \\l 1031${pages ? ` \\p ${pages}` : ""}`;
    const [from, to] = this.range();
    const end = E.insertObject(this.doc, from, to, E.fieldObject(instr, `(${tag})`), this.typingProps());
    this.setSel(end, end);
    this.refreshCitations();
    this.update({ scroll: true });
  }

  /**
   * Literaturverzeichnis: a heading ("Literaturverzeichnis", "Verweise", or none) and a
   * BIBLIOGRAPHY field with the cited sources, formatted in the document's style.
   */
  insertBibliography(heading = ""): void {
    if (this.doc.activeStory !== MAIN || this.inTable()) return;
    this.snapshot();
    const style = this.styleId("Bibliography");
    this.ensureStyle(style);
    const h1 = this.styleId("heading 1");
    this.ensureStyle(h1);
    // An empty paragraph at the caret is used; otherwise the list goes after the caret's paragraph.
    const cur = this.sel.head.p;
    const empty = !this.doc.main[cur]?.text.length;
    const at = empty ? cur : cur + 1;
    const paras: Para[] = [...(heading ? [makePara(heading, { style: h1 })] : []), ...S.bibliographyParagraphs([], style)];
    this.doc.splice(at, empty ? at : at - 1, paras);
    this.refreshCitations();
    this.setSel({ p: at, o: 0 }, { p: at, o: 0 });
    this.update({ scroll: true });
  }

  /** Zitate and Literaturverzeichnisse in the document's style, numbered by first citation. */
  private refreshCitations(): void {
    const b = this.bibliography();
    const cited = this.citedTags();
    const byTag = new Map(b.sources.map((x) => [x.tag, x]));
    const main = this.doc.main.map((p) => {
      if (!p.objects?.some((o) => o.kind === "field" && /^\s*CITATION\b/i.test(o.instr ?? ""))) return p;
      const objects = p.objects.map((o) => {
        const m = o.kind === "field" ? /^\s*CITATION\s+(\S+)(?:.*?\\p\s+(\S+))?/i.exec(o.instr ?? "") : null;
        if (!m) return o;
        const display = Bib.formatCitation(byTag.get(m[1]), m[1], b.style, cited.indexOf(m[1]) + 1, m[2] ?? "");
        return display === o.display ? o : { ...o, display, data: undefined };
      });
      return { ...p, objects };
    });
    this.doc.setStory(MAIN, main);
    const style = this.styleId("Bibliography");
    for (const f of S.tocFields(this.doc, /^\s*BIBLIOGRAPHY\b/i)) {
      const sources = Bib.sortedSources(b, cited).filter((x) => cited.includes(x.tag));
      const entries = sources.map((x) => Bib.formatEntry(x, b.style, cited.indexOf(x.tag) + 1));
      this.doc.splice(f.a, f.b, S.bibliographyParagraphs(entries as { text: string; italic?: boolean }[][], style, f.id));
    }
  }

  /** Literaturverzeichnis aktualisieren (and F9). */
  updateCitations(): void {
    this.snapshot();
    this.refreshCitations();
    this.update();
  }

  // ----- links ----------------------------------------------------------------------------

  /** The link at the caret (or around the selection): its text range and where it goes. */
  linkAt(pos: Pos = this.sel.head): { id: string; from: Pos; to: Pos; text: string; href?: string; anchor?: string; tip?: string; frame?: string } | null {
    const para = this.doc.paras[pos.p];
    const anchors = para?.anchors ?? [];
    for (const start of anchors) {
      if (start.link?.edge !== "start") continue;
      const end = anchors.find((x) => x.link?.id === start.link!.id && x.link.edge === "end");
      if (!end || pos.o < start.at || pos.o > end.at) continue;
      return { id: start.link.id, from: { p: pos.p, o: start.at }, to: { p: pos.p, o: end.at }, text: para.text.slice(start.at, end.at),
        href: start.link.href, anchor: start.link.anchor, tip: start.link.tip, frame: start.link.frame };
    }
    return null;
  }

  /**
   * Link einfügen (Strg+K): the text to show becomes a link (web address, file, mailto: or a
   * place in the document). An existing link at the caret is replaced; a link stays in one
   * paragraph. The text gets Word's character style Hyperlink.
   */
  insertLink(link: { text: string; href?: string; anchor?: string; tip?: string; frame?: string }): void {
    if (!link.href && !link.anchor) return;
    const old = this.linkAt();
    let [from, to] = old ? [old.from, old.to] : this.range();
    if (from.p !== to.p) to = { p: from.p, o: this.doc.paras[from.p].text.length };
    this.snapshot();
    this.typing = false;
    const style = this.styleId("hyperlink");
    this.ensureStyle(style);
    if (old) this.dropLink(old.id, from.p);
    const text = link.text || link.href || link.anchor || "";
    if (text !== this.doc.paras[from.p].text.slice(from.o, to.o)) {
      const props = intern({ ...typedProps(this.doc.paras[from.p], from.o), rStyle: style, rev: undefined });
      to = this.replaceTracked(from, to, [{ text, runs: [{ len: text.length, props }], props: {}, mark: EMPTY }]);
    } else {
      this.doc.mapChars(from, to, (c) => ({ ...c, rStyle: style }));
    }
    const id = `link:${++this.linkSeq}`;
    const paras = [...this.doc.paras];
    let para = paras[from.p];
    para = S.withAnchor(para, { at: from.o, data: undefined, link: { id, edge: "start", href: link.href, anchor: link.anchor, tip: link.tip || undefined, frame: link.frame || undefined } });
    para = S.withAnchor(para, { at: to.o, data: undefined, link: { id, edge: "end" } });
    paras[from.p] = para;
    this.doc.paras = paras;
    this.setSel(to, to);
    this.update({ scroll: true });
  }

  /** A heading as a link target: its hidden _Toc bookmark (added if missing), as Word does. */
  headingBookmark(p: number): string {
    const before = this.doc.main[p]?.anchors;
    const name = S.tocBookmark(this.doc, p);
    if (this.doc.main[p]?.anchors !== before) this.update();
    return name;
  }

  /** Strg+Klick on a link (the host opens web pages, files and mail; bookmarks are gone to). */
  onFollowLink: ((link: { href?: string; anchor?: string }) => void) | null = null;

  /** Link entfernen: the text stays, without the link and its style. */
  removeLink(): void {
    const l = this.linkAt();
    if (!l) return;
    this.snapshot();
    this.dropLink(l.id, l.from.p);
    const style = this.styleId("hyperlink");
    this.doc.mapChars(l.from, l.to, (c) => (c.rStyle === style ? { ...c, rStyle: undefined } : c));
    this.update();
  }

  private linkSeq = 0;

  private dropLink(id: string, p: number): void {
    const paras = [...this.doc.paras];
    paras[p] = { ...paras[p], anchors: paras[p].anchors?.filter((a) => a.link?.id !== id) };
    this.doc.paras = paras;
  }

  /**
   * Einfügen > Symbol: a character at the caret, in its own font when one was chosen. A symbol
   * font's character (U+F020-U+F0FF) is written as Word writes it, a w:sym.
   */
  insertSymbol(ch: string, font?: string, symbolFont = false): void {
    if (this.inBlock()) return;
    this.snapshot();
    this.typing = false;
    const [from, to] = this.range();
    const base: CharProps = { ...typedProps(this.doc.paras[from.p], from.o), rev: undefined };
    let piece: Para;
    if (symbolFont && font) {
      const hex = (ch.codePointAt(0) ?? 0).toString(16).toUpperCase().padStart(4, "0");
      const safe = font.replace(/&/g, "&amp;").replace(/"/g, "&quot;").replace(/</g, "&lt;");
      const obj: InlineObject = { kind: "other", data: { inner: `<w:sym w:font="${safe}" w:char="${hex}"/>` }, display: ch, font };
      piece = { text: OBJ, runs: [{ len: 1, props: intern(base) }], props: {}, mark: EMPTY, objects: [obj] };
    } else {
      piece = { text: ch, runs: [{ len: ch.length, props: intern(font ? { ...base, font } : base) }], props: {}, mark: EMPTY };
    }
    const end = this.replaceTracked(from, to, [piece]);
    this.setSel(end, end);
    this.update({ scroll: true });
  }

  /** Formel einfügen: a new equation at the caret; returns where it is. */
  insertMath(latex = "", seq?: Seq): Pos | null {
    if (this.inBlock()) return null;
    this.snapshot();
    this.typing = false;
    const [from, to] = this.range();
    const props = intern({ ...typedProps(this.doc.paras[from.p], from.o), rev: undefined });
    const math = seq ?? parseLatex(latex);
    const obj: InlineObject = { kind: "math", data: undefined, display: "", latex: seqToLatex(math), math };
    const piece: Para = { text: OBJ, runs: [{ len: 1, props }], props: {}, mark: EMPTY, objects: [obj] };
    const end = this.replaceTracked(from, to, [piece]);
    this.setSel(end, end);
    this.update({ scroll: true });
    return { p: end.p, o: end.o - 1 };
  }

  /** The equation at a position (its object character), if any. */
  mathAt(pos: Pos): InlineObject | null {
    const p = this.doc.paras[pos.p];
    if (!p || p.text[pos.o] !== OBJ) return null;
    const o = this.objectAt(p, pos.o);
    return o?.kind === "math" ? o : null;
  }

  /** Changes an equation; an empty one is removed. `record` adds an undo step. */
  setMathLatex(pos: Pos, latex: string, record = true): void {
    this.setMath(pos, latex.trim() ? parseLatex(latex) : [], record);
  }

  /** Changes an equation's tree; an empty one is removed. `record` adds an undo step. */
  setMath(pos: Pos, seq: Seq, record = true, keepEmpty = false): void {
    const o = this.mathAt(pos);
    if (!o) return;
    const latex = seqToLatex(seq);
    if (o.latex === latex && JSON.stringify(o.math) === JSON.stringify(seq) && (latex || keepEmpty)) return;
    if (record) this.snapshot();
    if (!seq.length && !keepEmpty) {
      this.doc.delete(pos, { p: pos.p, o: pos.o + 1 });
      this.setSel(pos, pos);
    } else this.replaceObject(pos.p, pos.o, { ...o, latex, math: seq, data: undefined, display: "", width: undefined, height: undefined, depth: undefined });
    this.update();
  }

  /** Formel > Konvertieren > Alle: changes every equation of the document (one undo step). */
  mapAllMath(fn: (seq: Seq) => Seq): void {
    const mapParas = (paras: readonly Para[]): Para[] | null => {
      let out: Para[] | null = null;
      paras.forEach((para, i) => {
        if (!para.objects?.some((o) => o.kind === "math")) return;
        const objects = para.objects.map((o) => {
          if (o.kind !== "math") return o;
          const seq = fn(o.math ?? parseLatex(o.latex ?? ""));
          const latex = seqToLatex(seq);
          if (latex === o.latex) return o;
          return { ...o, latex, math: seq, data: undefined, display: "", width: undefined, height: undefined, depth: undefined };
        });
        if (objects.every((o, k) => o === para.objects![k])) return;
        out ??= [...paras];
        out[i] = { ...para, objects };
      });
      return out;
    };
    const updates: [string, Para[]][] = [];
    const m = mapParas(this.doc.main);
    if (m) updates.push([MAIN, m]);
    for (const [id, paras] of this.doc.stories) { const r = mapParas(paras); if (r) updates.push([id, r]); }
    if (!updates.length) return;
    this.snapshot();
    for (const [id, paras] of updates) this.doc.setStory(id, paras);
    this.update();
  }

  /** Where an object is, in the view's container coordinates (unzoomed). */
  objectBox(pos: Pos): { x: number; y: number; w: number; h: number; baseline: number; size: number } | null {
    if (this.doc.paras[pos.p]?.text[pos.o] !== OBJ) return null;
    const { pl } = this.view.placedFor(pos, false);
    if (!pl) return null;
    const line = this.view.lineOf(pl);
    const i = pos.o - line.start;
    if (i < 0 || i >= line.end - line.start) return null;
    const x = this.view.pageLeft(pl.page) + pl.x + line.x0 + line.charX[i];
    const w = line.charX[i + 1] - line.charX[i];
    const top = this.view.pageTop(pl.page) + pl.y;
    const rp = this.view.storyLayout().paras[pos.p]?.rp;
    const size = this.doc.styles.resolveChar(rp?.style, propsAt(this.doc.paras[pos.p], pos.o)).size;
    return { x, y: top, w, h: line.height, baseline: top + line.baseline, size };
  }

  /** The drawn equation at a position (the formula editor reads its geometry) and its page. */
  mathElement(pos: Pos): { el: Element; page: number } | null {
    const para = this.doc.paras[pos.p];
    return para ? this.view.mathElementFor(para, pos.o) : null;
  }

  /**
   * Formeloptionen > In 'Inline' ändern / In 'Anzeige' ändern: a display equation is shown inline
   * in its paragraph; an inline one gets a paragraph of its own. Returns where it is then.
   */
  setMathDisplay(pos: Pos, display: boolean): Pos | null {
    const o = this.mathAt(pos);
    if (!o?.math || this.inBlock()) return null;
    this.snapshot();
    this.typing = false;
    const p0 = this.doc.paras[pos.p];
    const blank = (): Para => ({ text: "", runs: [], props: p0.props, mark: p0.mark });
    let at = pos;
    if (display && p0.text !== OBJ) {
      if (pos.o + 1 < p0.text.length) this.doc.replace({ p: pos.p, o: pos.o + 1 }, { p: pos.p, o: pos.o + 1 }, [blank(), blank()]);
      if (pos.o > 0) { this.doc.replace(pos, pos, [blank(), blank()]); at = { p: pos.p + 1, o: 0 }; }
    }
    this.replaceObject(at.p, at.o, { ...o, mathInline: display ? undefined : true, data: undefined, width: undefined, height: undefined, depth: undefined });
    this.setSel({ p: at.p, o: at.o + 1 }, { p: at.p, o: at.o + 1 });
    this.update({ scroll: true });
    return at;
  }

  /** Formeloptionen > Ausrichtung of a display equation (one undo step). */
  setMathJc(pos: Pos, jc: "left" | "right" | "center" | "centerGroup"): void {
    const o = this.mathAt(pos);
    if (!o?.math || o.mathJc === jc) return;
    this.snapshot();
    this.replaceObject(pos.p, pos.o, { ...o, mathJc: jc, data: undefined });
    this.update();
  }

  /** Whether the line holding a position is the first or the last of its story (Up/Down cannot leave it). */
  lineEdges(pos: Pos): { first: boolean; last: boolean } {
    const { index } = this.view.placedFor(pos, false);
    const n = this.view.storyLayout().placed.length;
    return { first: index <= 0, last: index >= n - 1 };
  }

  /** The text position under a client point (where something dragged would go). */
  posAtPoint(clientX: number, clientY: number): Pos {
    const hit = this.view.hitLine(clientX, clientY);
    return this.view.posInLine(hit.index, hit.x).pos;
  }

  /**
   * The equation's move handle (Feld zum Verschieben): the object at `from` goes to `to`
   * (one undo step). Returns its new position.
   */
  moveObject(from: Pos, to: Pos): Pos | null {
    const para = this.doc.paras[from.p];
    if (!para || para.text[from.o] !== OBJ || this.inBlock()) return null;
    if (to.p === from.p && (to.o === from.o || to.o === from.o + 1)) return from;
    const obj = this.objectAt(para, from.o);
    if (!obj) return null;
    const props = propsAt(para, from.o);
    this.snapshot();
    this.typing = false;
    this.doc.delete(from, { p: from.p, o: from.o + 1 });
    const at = to.p === from.p && to.o > from.o ? { p: to.p, o: to.o - 1 } : to;
    const piece: Para = { text: OBJ, runs: [{ len: 1, props }], props: {}, mark: EMPTY, objects: [obj] };
    const end = this.replaceTracked(at, at, [piece]);
    this.setSel(end, end);
    this.update({ scroll: true });
    return { p: end.p, o: end.o - 1 };
  }

  /** Puts the caret before or after an object (leaving the equation editor). */
  caretAt(pos: Pos): void {
    this.setSel(pos, pos);
    this.update({ scroll: true });
    this.focus();
  }

  onChange(listener: (e: ChangeEvent) => void): () => void {
    this.listeners.push(listener);
    return () => { this.listeners = this.listeners.filter((l) => l !== listener); };
  }

  /** The main text's paragraphs (word count, file name suggestions). */
  paragraphs(): readonly Para[] {
    return this.doc.main;
  }

  /** Identity of the whole content: changes whenever any story, style or setting does. */
  contentStamp(): readonly unknown[] {
    const d = this.doc;
    return [d.main, d.stories, d.styles, d.settings, d.numbering, d.comments];
  }

  stats(): EditorStats {
    return { pageCount: this.layout.pageCount, caretPage: this.view.caretBox(this.sel.head, this.sel.headUpstream).page };
  }

  countWords(): number {
    return countWords(this.doc.paras.filter((p) => !p.block).map((p) => p.text.replace(/[\u000b\u000c]/g, " ")));
  }

  get canUndo(): boolean { return this.undoStack.length > 0; }
  get canRedo(): boolean { return this.redoStack.length > 0; }
  get hasSelection(): boolean { return cmp(this.sel.anchor, this.sel.head) !== 0; }
  get zoom(): number { return this.view.zoom; }
  get showMarks(): boolean { return this.view.showMarks; }
  get painterActive(): boolean { return this.painter !== null; }

  setZoom(z: number): void {
    this.view.setZoom(z);
    // Weblayout and Lesemodus reflow to the window at the new zoom.
    if (this.viewMode !== "print" && this.computeOverride()) this.update();
    this.emit(false);
  }

  // ----- Ansicht: Drucklayout, Weblayout, Lesemodus ----------------------------------------------

  /** The view: Drucklayout (pages), Weblayout (text as wide as the window) or Lesemodus. */
  viewMode: "print" | "web" | "read" = "print";
  /** Lesemodus > Ansicht: Spaltenbreite and Layout (Spaltenlayout or Papierlayout). */
  readOptions: { column: "narrow" | "standard" | "wide"; paper: boolean } = { column: "standard", paper: false };
  private pageOverride: FlowInput["pageOverride"] = undefined;
  private overrideKey = "";
  private resizeWatch: ResizeObserver | null = null;
  private resizeTimer = 0;

  /** Lesemodus: reading, selecting and copying only. */
  get readOnly(): boolean {
    return this.viewMode === "read";
  }

  setViewMode(mode: Editor["viewMode"]): void {
    if (mode !== "print" && this.doc.activeStory !== MAIN) this.switchStory(MAIN, MAIN);
    this.viewMode = mode;
    this.view.setArrangement(mode === "web" ? "flat" : mode === "read" ? "row" : "pages");
    if (mode !== "print" && !this.resizeWatch && typeof ResizeObserver !== "undefined") {
      this.resizeWatch = new ResizeObserver(() => {
        clearTimeout(this.resizeTimer);
        this.resizeTimer = window.setTimeout(() => { if (this.viewMode !== "print" && this.computeOverride()) this.update(); }, 80);
      });
      this.resizeWatch.observe(this.workspace);
    }
    this.computeOverride();
    this.update({ scroll: true });
    if (mode === "read") this.view.showScreen(this.view.screenInfo().page);
    this.emit(false);
  }

  setReadOptions(o: Partial<Editor["readOptions"]>): void {
    this.readOptions = { ...this.readOptions, ...o };
    this.view.setArrangement(this.viewMode === "read" ? "row" : this.view.arrangement);
    if (this.computeOverride() || this.viewMode === "read") this.update({ scroll: true });
  }

  /** Lesemodus: pages per screen and the screen shown. */
  readScreens(): { page: number; count: number; perScreen: number } {
    const info = this.view.screenInfo();
    return { ...info, perScreen: this.perScreen };
  }

  /** Lesemodus: the next or previous screen. */
  turnScreen(dir: 1 | -1): void {
    const { page } = this.view.screenInfo();
    const per = this.perScreen;
    this.view.showScreen(Math.floor(page / per) * per + dir * per);
    this.emit(false);
  }

  private perScreen = 1;

  /** The page Weblayout and Lesemodus lay out on, from the window and zoom; true when it changed. */
  private computeOverride(): boolean {
    const ws = this.workspace;
    const z = this.view.zoom;
    let o: FlowInput["pageOverride"];
    this.perScreen = 1;
    if (this.viewMode === "web") {
      const w = Math.max(144, (ws.clientWidth - 4) / z / PT);
      o = { page: { width: w, height: 100000, marginLeft: 14.4, marginRight: 14.4, marginTop: 14.4, marginBottom: 14.4, header: 0, footer: 0, gutter: 0 }, fitHeight: true };
    } else if (this.viewMode === "read" && !this.readOptions.paper) {
      // Columns of about Word's widths (Schmal, Standard, Breit), as many as fit on a screen.
      // A screen shows `per` pages and their gaps exactly (px at 100 %, the view scales them).
      const gap = 16;
      const col = { narrow: 330, standard: 470, wide: 650 }[this.readOptions.column];
      const W = Math.max(300, ws.clientWidth) / z, H = Math.max(300, ws.offsetHeight - 18) / z;
      const per = Math.max(1, Math.floor(W / PT / (col + 24)));
      this.perScreen = per;
      const w = ((W - gap) / per - gap) / PT, hgt = (H - 2 * gap - 2) / PT;
      o = { page: { width: w, height: hgt, marginLeft: 28, marginRight: 28, marginTop: 24, marginBottom: 24, header: 0, footer: 0, gutter: 0 } };
    }
    const key = JSON.stringify(o ?? null);
    if (key === this.overrideKey) return false;
    this.overrideKey = key;
    this.pageOverride = o;
    return true;
  }

  setShowMarks(on: boolean): void {
    this.view.setShowMarks(on);
    this.emit(false);
  }

  /** Replaces the whole document (open file, feel-lab fixture). */
  load(state: Partial<DocState>): void {
    this.doc = new Doc(state);
    this.storyKey = MAIN;
    this.sel = { anchor: { p: 0, o: 0 }, head: { p: 0, o: 0 }, headUpstream: false };
    this.undoStack = [];
    this.redoStack = [];
    this.pending = null;
    this.workspace.scrollTop = 0;
    this.layouter.setContext(this.context());
    this.requireFonts();
    this.update();
  }

  focus(): void {
    this.input.focus({ preventScroll: true });
  }

  // ----- state updates ------------------------------------------------------------------

  update(opts: { scroll?: boolean; inputTs?: number } = {}): void {
    const ctx = this.layouter.context;
    if (ctx.styles !== this.doc.styles || ctx.defaultTab !== this.doc.settings.defaultTab || ctx.measureMath !== this.measureMath || ctx.markup !== this.markupMode ||
      ctx.mathJc !== this.doc.settings.math?.defJc) {
      this.layouter.setContext(this.context());
    }
    this.layout = layoutDocument(this.layouter, {
      main: this.doc.main, stories: this.doc.stories, settings: this.doc.settings, numbering: this.doc.numbering,
      pageOverride: this.pageOverride,
    });
    this.fixStoryKey();
    this.clampSelection();
    const ropts = { active: this.storyKey, cells: this.selectedCells(), image: this.imageOverlay(), table: this.handleTable() };
    this.view.render(this.doc, this.layout, this.sel, ropts);
    // The picture's frame comes from the view's new layout: draw it again when it moved.
    const image = this.imageOverlay();
    if (JSON.stringify(image) !== JSON.stringify(ropts.image)) { ropts.image = image; this.view.render(this.doc, this.layout, this.sel, ropts); }
    if (opts.scroll) {
      this.view.scrollIntoView(this.sel.head, this.sel.headUpstream);
      this.view.render(this.doc, this.layout, this.sel, ropts);
    }
    this.placeInput();
    if (opts.inputTs !== undefined) markInputHandled(opts.inputTs);
    const content = this.doc.main !== this.lastParas || this.doc.stories !== this.lastStories;
    this.lastParas = this.doc.main;
    this.lastStories = this.doc.stories;
    this.emit(content);
  }

  /**
   * The layout key of the story with the caret: a header or footer is shown once per page,
   * so the caret sits in one of its instances. When that instance is gone (fewer pages, a
   * different header), another one is used, or the main text.
   */
  private fixStoryKey(): void {
    const id = this.doc.activeStory;
    if (id === MAIN) { this.storyKey = MAIN; return; }
    if (this.layout.stories.has(this.storyKey) && storyIdOf(this.storyKey) === id) return;
    const prefix = `${id}@`;
    for (const k of this.layout.stories.keys()) {
      if (k === id || k.startsWith(prefix)) { this.storyKey = k; return; }
    }
    // Not shown anywhere (e.g. a header of a page that no longer exists): back to the text.
    this.doc.setActive(MAIN);
    this.storyKey = MAIN;
    this.sel = { anchor: { p: 0, o: 0 }, head: { p: 0, o: 0 }, headUpstream: false };
  }

  private clampSelection(): void {
    const n = this.doc.paras.length;
    const fix = (p: Pos): Pos => {
      const pi = Math.min(Math.max(0, p.p), n - 1);
      return { p: pi, o: Math.min(Math.max(0, p.o), this.doc.paras[pi].text.length) };
    };
    const a = fix(this.sel.anchor), h = fix(this.sel.head);
    if (cmp(a, this.sel.anchor) || cmp(h, this.sel.head)) this.sel = { ...this.sel, anchor: a, head: h };
  }

  /** Cells selected as a whole: when a selection spans cells (Word selects whole cells). */
  private selectedCells(): Set<string> | undefined {
    const [from, to0] = this.range();
    const paras = this.doc.paras;
    // A selection that ends at the start of a cell ends after the cell before it.
    let to = to0;
    const tc = paras[to.p]?.cell, fc = paras[from.p]?.cell;
    if (to.o === 0 && tc && to.p > from.p && !(fc && fc.table.id === tc.table.id && fc.row === tc.row && fc.col === tc.col)) {
      to = { p: to.p - 1, o: paras[to.p - 1].text.length };
    }
    const a = paras[from.p]?.cell, b = paras[to.p]?.cell;
    if (from.p === to.p || (!a && !b)) return undefined;
    if (a && b && a.table.id === b.table.id && a.row === b.row && a.col === b.col) return undefined;
    if (a && b && a.table.id === b.table.id) {
      // Word selects a block: the rows and columns between the two cells.
      const r1 = Math.min(a.row, b.row), r2 = Math.max(a.row, b.row), c1 = Math.min(a.col, b.col), c2 = Math.max(a.col, b.col);
      const block = new Set<string>();
      for (let r = r1; r <= r2; r++) for (let c = c1; c <= Math.min(c2, a.table.rows[r].cells.length - 1); c++) block.add(`${a.table.id}:${r}:${c}`);
      return block;
    }
    const out = new Set<string>();
    for (let i = from.p; i <= to.p; i++) {
      const c = this.doc.paras[i].cell;
      if (c) out.add(`${c.table.id}:${c.row}:${c.col}`);
    }
    return out.size ? out : undefined;
  }

  private emit(content: boolean): void {
    for (const l of this.listeners) l({ content });
  }

  private placeInput(): void {
    const c = this.view.caretBox(this.sel.head, this.sel.headUpstream);
    this.input.style.left = `${this.view.pageLeft(c.page) + c.x}px`;
    this.input.style.top = `${this.view.pageTop(c.page) + c.y}px`;
    // Keep the selected text in the textarea so the browser's copy/cut events fire.
    const selected = this.hasSelection ? this.selectedText() : "";
    const value = selected.length > 0 && selected.length < 100_000 ? selected : "";
    if (this.input.value !== value) this.input.value = value;
    if (value) this.input.select();
  }

  select(anchor: Pos, head: Pos): void {
    this.setSel(anchor, head);
    this.update({ scroll: true });
  }

  private setSel(anchor: Pos, head: Pos, upstream = false, keepGoal = false): void {
    this.sel = { anchor, head, headUpstream: upstream };
    this.cellAnchor = null;
    this.floatSel = null;
    if (!keepGoal) this.goalX = null;
    this.typing = false;
    this.pending = null;
  }

  private range(): [Pos, Pos] {
    return [minPos(this.sel.anchor, this.sel.head), maxPos(this.sel.anchor, this.sel.head)];
  }

  selectedText(): string {
    const [from, to] = this.range();
    return this.doc.slice(from, to);
  }

  // ----- queries for the ribbon ----------------------------------------------------------

  private resolvedAt(p: number, props: CharProps): ResolvedChar {
    return this.doc.styles.resolveChar(this.doc.styles.resolvePara(this.doc.paras[p].props).style, props);
  }

  /** Raw formatting the next typed character gets at the caret. */
  typingProps(): CharProps {
    const h = this.sel.head;
    return this.pending ?? typedProps(this.doc.paras[h.p], h.o);
  }

  /** Formatting shown in the ribbon (font box, Fett button …). */
  charState(): CharState {
    const list: { c: ResolvedChar; raw: CharProps; space?: boolean }[] = [];
    if (!this.hasSelection) {
      const h = this.sel.head;
      const raw = this.pending ?? typedProps(this.doc.paras[h.p], h.o);
      list.push({ c: this.resolvedAt(h.p, raw), raw });
    } else {
      const [from, to] = this.range();
      for (let i = from.p; i <= to.p && list.length < 2000; i++) {
        const p = this.doc.paras[i];
        if (p.block) continue;
        let pos = 0;
        for (const r of p.runs) {
          const a = pos, b = pos + r.len;
          pos = b;
          if ((i === from.p && b <= from.o) || (i === to.p && a >= to.o)) continue;
          const lo = i === from.p ? Math.max(a, from.o) : a, hi = i === to.p ? Math.min(b, to.o) : b;
          list.push({ c: this.resolvedAt(i, r.props), raw: r.props, space: !p.text.slice(lo, hi).trim() });
        }
      }
      if (!list.length) list.push({ c: this.resolvedAt(from.p, this.doc.paras[from.p].mark), raw: this.doc.paras[from.p].mark });
    }
    // As in Word, spaces do not count when deciding whether a selection is uniform.
    const counted = list.some((x) => !x.space) ? list.filter((x) => !x.space) : list;
    const same = <T>(f: (x: { c: ResolvedChar; raw: CharProps }) => T): T | undefined => {
      const v = f(counted[0]);
      return counted.every((x) => f(x) === v) ? v : undefined;
    };
    return {
      font: same((x) => x.c.font),
      size: same((x) => x.c.size),
      bold: same((x) => x.c.bold),
      italic: same((x) => x.c.italic),
      underline: same((x) => x.c.underline),
      strike: same((x) => x.c.strike),
      dstrike: same((x) => x.c.dstrike),
      superscript: same((x) => x.c.vertAlign === "superscript"),
      subscript: same((x) => x.c.vertAlign === "subscript"),
      caps: same((x) => x.c.caps),
      smallCaps: same((x) => x.c.smallCaps),
      hidden: same((x) => x.c.hidden),
      color: same((x) => x.c.color),
      highlight: same((x) => x.c.highlight),
      charStyle: same((x) => x.raw.rStyle ?? ""),
    };
  }

  /** Resolved paragraph formatting of the paragraph with the caret. */
  paraState(): ResolvedPara {
    return this.doc.styles.resolvePara(this.doc.paras[this.sel.head.p].props);
  }

  /** Style of the paragraph holding the caret (for the style gallery). */
  currentStyle(): string {
    return this.paraState().style;
  }

  // ----- undo ---------------------------------------------------------------------------

  private current(): Snapshot {
    return { state: this.doc.snapshot(), story: this.doc.activeStory, key: this.storyKey, anchor: this.sel.anchor, head: this.sel.head };
  }

  private snapshot(): void {
    this.undoStack.push(this.current());
    if (this.undoStack.length > 500) this.undoStack.shift();
    this.redoStack = [];
  }

  private restoreSnapshot(s: Snapshot): void {
    this.doc.restore(s.state);
    this.doc.activeStory = s.story === MAIN || this.doc.stories.has(s.story) ? s.story : MAIN;
    this.storyKey = this.doc.activeStory === MAIN ? MAIN : s.key;
    this.setSel(s.anchor, s.head);
    this.update({ scroll: true });
  }

  undo(): void {
    const s = this.undoStack.pop();
    if (!s) return;
    this.redoStack.push(this.current());
    this.restoreSnapshot(s);
  }

  redo(): void {
    const s = this.redoStack.pop();
    if (!s) return;
    this.undoStack.push(this.current());
    this.restoreSnapshot(s);
  }

  /** Runs a formatting change as one undo step. */
  private edit(fn: () => void, scroll = true): void {
    this.snapshot();
    this.typing = false;
    fn();
    this.requireFonts();
    this.update({ scroll });
  }

  // ----- text editing -------------------------------------------------------------------

  private inBlock(): boolean {
    return !!this.doc.paras[this.sel.head.p].block && !this.hasSelection;
  }

  private replaceSelection(text: string, inputTs?: number, continueTyping = false, props?: CharProps): void {
    if (this.inBlock()) return;
    const continuing = continueTyping && this.typing;
    if (!continuing) this.snapshot();
    const [from, to] = this.range();
    let fmt = props ?? this.pending ?? (cmp(from, to) !== 0
      ? (from.o < this.doc.paras[from.p].text.length ? this.doc.charPropsAt(from) : typedProps(this.doc.paras[from.p], from.o))
      : typedProps(this.doc.paras[from.p], from.o));
    const pending = this.pending;
    // Tracked changes: new text is an insertion of this author; what it replaces stays as a
    // deletion before it. Without tracking, typed text never inherits a revision.
    const tracking = !!this.doc.settings.trackChanges;
    if (tracking && (!continuing || !this.insRev)) this.insRev = R.newRevision("ins", this.author);
    fmt = intern({ ...fmt, rev: tracking ? this.insRev! : undefined });
    let end: Pos;
    if (tracking && cmp(from, to) !== 0) {
      end = this.doc.insertText(to, to, text, fmt);
      end = this.keepPos(end, () => R.markDeleted(this.doc, from, to, R.newRevision("del", this.author)));
    } else end = this.doc.insertText(from, to, text, fmt);
    this.setSel(end, end);
    // Formatting chosen at the caret applies to what is typed next; Enter drops it (Word: the
    // new paragraph takes its formatting from the paragraph mark).
    if (pending) this.pending = null;
    this.typing = continueTyping;
    this.update({ scroll: true, inputTs });
  }

  insertText(text: string, inputTs?: number): void {
    const shape = this.selectedShape();
    if (shape && !shape.editing && shape.obj.shape?.kind !== "connector" && (shape.obj.pic?.geometry ?? "rect") !== "line") {
      this.editShapeText(shape.p, shape.index, "end");
    }
    this.smartTrimSelection();
    this.replaceSelection(text, inputTs, true);
    if (text === " " || text === "\t") this.autoList();
    // Entries such as (c), --> and :) are replaced as soon as their last character is typed.
    if (text.length === 1 && INSTANT_TRIGGERS.has(text) && this.instantCorrect()) return;
    if (text.length === 1 && TRIGGERS.has(text)) this.autoCorrect(text);
    if (text === "*" || text === "~") this.markdownInline(text);
  }

  /** Caps Lock state from the last key event (Word corrects "hALLO" only while it is on). */
  private capsLock = false;

  private instantCorrect(): boolean {
    if (!this.autoCorrectOn || this.hasSelection) return false;
    const h = this.sel.head;
    const para = this.doc.paras[h.p];
    if (!para || para.block) return false;
    const f = instantFix(para.text, h.o);
    if (!f || para.text.slice(f.a, f.b).includes(OBJ)) return false;
    this.applyFixes(h.p, [f]);
    return true;
  }

  /**
   * Word 365's Markdown while typing: **text** becomes bold and ~~text~~ struck through when
   * the closing marks are typed (measured 2026-10-14; *text* and _text_ stay as typed).
   */
  private markdownInline(mark: string): void {
    if (!this.autoCorrectOn || this.hasSelection) return;
    const h = this.sel.head;
    const para = this.doc.paras[h.p];
    if (!para || para.block) return;
    const before = para.text.slice(0, h.o);
    const re = mark === "*" ? /(^|[\s(„“"'‚])\*\*([^*\s](?:[^*]*[^*\s])?)\*\*$/u : /(^|[\s(„“"'‚])~~([^~\s](?:[^~]*[^~\s])?)~~$/u;
    const m = re.exec(before);
    if (!m || m[2].includes(OBJ)) return;
    const a = h.o - m[2].length - 4;
    this.snapshot();
    this.typing = false;
    const props = typedProps(para, h.o);
    this.doc.delete({ p: h.p, o: h.o - 2 }, { p: h.p, o: h.o });
    this.doc.delete({ p: h.p, o: a }, { p: h.p, o: a + 2 });
    const end = { p: h.p, o: a + m[2].length };
    this.doc.mapChars({ p: h.p, o: a }, end, (old) => intern(mark === "*" ? { ...old, bold: true } : { ...old, strike: true }));
    this.setSel(end, end);
    this.pending = props;
    this.update({ scroll: true });
  }

  /** AutoKorrektur ein (Word's defaults); a setting so tests and future options can turn it off. */
  autoCorrectOn = true;

  /** Language of the text at the caret, for quotes and word rules. */
  private langAt(pos: Pos): string {
    const para = this.doc.paras[pos.p];
    return propsAt(para, Math.max(0, pos.o - 1)).lang ?? this.doc.styles.defaultsR.lang ?? "de-DE";
  }

  /**
   * After a typed word end (space, punctuation, quote …): Word's corrections of the word just
   * finished, and typographic quotes. Each correction is its own undo step.
   */
  private autoCorrect(trigger: string): void {
    if (!this.autoCorrectOn || this.hasSelection) return;
    const h = this.sel.head;
    const para = this.doc.paras[h.p];
    if (!para || para.block) return;
    const lang = this.langAt(h);
    const fixes: Fix[] = [...wordFixes(para.text, h.o - 1, lang, { cellStart: !!para.cell, capsLock: this.capsLock })];
    if (trigger === "\"" || trigger === "'") {
      const q = smartQuote(para.text, h.o, lang);
      if (q) fixes.push(q);
    }
    this.applyFixes(h.p, fixes);
  }

  /** Before Enter splits the paragraph: corrections of the last word, border lines ("---"). */
  private autoCorrectEnter(): boolean {
    if (!this.autoCorrectOn || this.hasSelection) return false;
    const h = this.sel.head;
    const para = this.doc.paras[h.p];
    if (!para || para.block) return false;
    if (h.o === para.text.length && !para.cell) {
      const line = borderLine(para.text);
      if (line) {
        // AutoFormat: the characters go, the paragraph gets a bottom border.
        this.snapshot();
        this.doc.delete({ p: h.p, o: 0 }, { p: h.p, o: para.text.length });
        this.doc.mapParas(h.p, h.p, (old) => ({ ...old, borders: { ...old.borders, bottom: { style: line.style, size: line.width, space: 1, color: "auto" } } }));
        this.setSel({ p: h.p, o: 0 }, { p: h.p, o: 0 });
        this.update();
        return true;
      }
    }
    this.applyFixes(h.p, wordFixes(para.text, h.o, this.langAt(h), { cellStart: !!para.cell, capsLock: this.capsLock }));
    return false;
  }

  private applyFixes(p: number, fixes: Fix[]): void {
    if (!fixes.length) return;
    this.snapshot();
    this.typing = false;
    const caret = { ...this.sel.head };
    let after: CharProps | null = null;
    for (const f of [...fixes].sort((x, y) => y.a - x.a)) {
      const para = this.doc.paras[p];
      const props = propsAt(para, f.a);
      // Text typed after a symbol or emoji is in the text's font again.
      if ((f.symbol || f.font) && caret.p === p && caret.o === f.b) after = props;
      let len = f.text.length;
      if (f.symbol) {
        const obj: InlineObject = { kind: "other", data: { inner: `<w:sym w:font="${f.symbol.font}" w:char="${f.symbol.char.toString(16).toUpperCase()}"/>` },
          display: String.fromCharCode(f.symbol.char), font: f.symbol.font };
        this.doc.replace({ p, o: f.a }, { p, o: f.b }, [{ text: OBJ, runs: [{ len: 1, props: intern({ ...props, font: f.symbol.font, fontTheme: undefined }) }], props: {}, mark: EMPTY, objects: [obj] }]);
        len = 1;
      } else this.doc.insertText({ p, o: f.a }, { p, o: f.b }, f.text, f.font ? intern({ ...props, font: f.font, fontTheme: undefined }) : props);
      if (caret.p === p && caret.o >= f.b) caret.o += len - (f.b - f.a);
    }
    this.setSel(caret, caret);
    if (after) this.pending = after;
    this.update({ scroll: true });
  }

  /**
   * AutoFormat während der Eingabe: "1. ", "a) ", "* " or "- " typed at the start of a
   * paragraph turn it into a list item (Strg+Z brings the typed text back).
   */
  private autoList(): void {
    if (!this.autoCorrectOn) return;
    const h = this.sel.head;
    const para = this.doc.paras[h.p];
    if (para.block || numberingOf(para.props, this.doc.styles) || h.o < 2 || h.o > 7) return;
    const prefix = para.text.slice(0, h.o - 1);
    if (para.text.slice(h.o - 1, h.o) !== " " && para.text.slice(h.o - 1, h.o) !== "\t") return;
    // Markdown: "# " … "###### " make the paragraph a heading (Word 365).
    if (/^#{1,6}$/.test(prefix) && para.text.slice(h.o - 1, h.o) === " " && !para.cell) {
      this.snapshot();
      this.typing = false;
      this.doc.delete({ p: h.p, o: 0 }, { p: h.p, o: h.o });
      const id = this.styleId(`heading ${prefix.length}`);
      this.ensureStyle(id);
      this.doc.mapParas(h.p, h.p, (old) => ({ style: id, extra: old.extra }));
      this.setSel({ p: h.p, o: 0 }, { p: h.p, o: 0 });
      this.requireFonts();
      this.update({ scroll: true });
      return;
    }
    const font = this.resolvedAt(h.p, propsAt(para, 0)).font;
    const auto = E.autoListLevels(prefix, font);
    if (!auto) return;
    this.snapshot();
    this.typing = false;
    this.doc.delete({ p: h.p, o: 0 }, { p: h.p, o: h.o });
    // Continue a list just above with the same kind of label, else start a new one.
    E.toggleList(this.doc, h.p, h.p, auto.levels[0].format === "bullet" ? "bullet" : "number", this.listStyleId(), auto.levels);
    this.setSel({ p: h.p, o: 0 }, { p: h.p, o: 0 });
    this.update({ scroll: true });
  }

  private listStyleId(): string {
    const id = this.styleId("List Paragraph");
    this.ensureStyle(id);
    return id;
  }

  /** Word's smart cut and paste: replacing whole words keeps their trailing space. */
  private smartTrimSelection(): void {
    const [from, to] = this.range();
    if (from.p !== to.p || cmp(from, to) === 0) return;
    const text = this.doc.paras[from.p].text;
    const wholeWords = tokenAt(text, from.o)[0] === from.o && (to.o === text.length || tokenAt(text, to.o)[0] === to.o);
    if (!wholeWords || text[to.o - 1] !== " ") return;
    let end = to.o;
    while (end > from.o && text[end - 1] === " ") end--;
    if (end === from.o) return;
    const head = cmp(this.sel.head, this.sel.anchor) >= 0;
    const trimmed = { p: to.p, o: end };
    this.sel = head ? { ...this.sel, head: trimmed } : { ...this.sel, anchor: trimmed };
  }

  private deleteRange(from: Pos, to: Pos, inputTs?: number, forward = false): void {
    if (cmp(from, to) === 0) return;
    this.snapshot();
    if (this.doc.settings.trackChanges) {
      // The text stays, marked deleted; the caret moves over it (Entf: to its end).
      const count = this.doc.paras.length, before = this.doc.paras[to.p].text.length;
      R.markDeleted(this.doc, from, to, R.newRevision("del", this.author));
      const kept = this.doc.paras.length === count && this.doc.paras[to.p].text.length === before;
      const at = forward && kept ? to : from;
      this.setSel(at, at);
    } else {
      this.doc.delete(from, to);
      this.setSel(from, from);
    }
    this.insRev = null;
    this.update({ scroll: true, inputTs });
  }

  /** Replaces [a, b) by pieces; with tracking the pieces are insertions and [a, b) stays deleted. */
  private replaceTracked(a: Pos, b: Pos, pieces: Para[]): Pos {
    if (!this.doc.settings.trackChanges) return this.doc.replace(a, b, pieces);
    const rev = R.newRevision("ins", this.author);
    const mark = (props: CharProps) => intern({ ...props, rev });
    const tracked = pieces.map((p, i) => ({
      ...p, runs: p.runs.map((r) => ({ ...r, props: mark(r.props) })),
      // The last piece's mark is the target paragraph's own (not inserted).
      mark: i < pieces.length - 1 ? mark(p.mark) : p.mark,
    }));
    const end = this.doc.replace(b, b, tracked);
    return cmp(a, b) === 0 ? end : this.keepPos(end, () => R.markDeleted(this.doc, a, b, R.newRevision("del", this.author)));
  }

  /** Runs a change that only touches text before `pos`; returns where `pos` is afterwards. */
  private keepPos(pos: Pos, change: () => void): Pos {
    const fromEnd = this.doc.paras.length - pos.p, tail = this.doc.paras[pos.p].text.length - pos.o;
    change();
    const p = Math.max(0, this.doc.paras.length - fromEnd);
    return { p, o: Math.max(0, this.doc.paras[p].text.length - tail) };
  }

  /** Author of tracked changes and comments (Optionen > Benutzername). */
  author = "Benutzer";
  authorInitials = "B";
  private insRev: import("./model").Revision | null = null;

  /**
   * Word's smart cut: deleting whole words also removes the space that would be left doubled
   * (or a space before punctuation).
   */
  private smartRange(from: Pos, to: Pos): [Pos, Pos] {
    if (from.p !== to.p || cmp(from, to) === 0) return [from, to];
    const text = this.doc.paras[from.p].text;
    const startsWord = from.o === 0 || !isWordChar(text[from.o - 1]);
    const endsWord = to.o === text.length || !isWordChar(text[to.o]);
    if (!startsWord || !endsWord || !isWordChar(text[from.o]) || !isWordChar(text[to.o - 1])) return [from, to];
    if (text[from.o - 1] === " " && text[to.o] === " ") return [from, { p: to.p, o: to.o + 1 }];
    if (text[from.o - 1] === " " && /[.,;:!?)\]]/.test(text[to.o] ?? "")) return [{ p: from.p, o: from.o - 1 }, to];
    return [from, to];
  }

  private deleteBackward(word: boolean, inputTs: number): void {
    const [from, to] = this.range();
    if (cmp(from, to) !== 0) { const [a, b] = this.smartRange(from, to); return this.deleteRange(a, b, inputTs); }
    const h = this.sel.head;
    if (h.o === 0) {
      const paras = this.doc.paras;
      // At the start of a list item the number goes first (the text stays indented).
      if (numberingOf(paras[h.p].props, this.doc.styles) && !paras[h.p].block) {
        this.edit(() => E.dropListLabel(this.doc, h.p));
        return;
      }
      if (this.backspaceIndent(h.p)) return;
      if (h.p === 0 || paras[h.p - 1].block || paras[h.p].block) return;
      const cur = paras[h.p].cell, prev = paras[h.p - 1].cell;
      if (cur || prev) {
        const same = cur && prev && cur.table.id === prev.table.id && cur.row === prev.row && cur.col === prev.col;
        if (!same) {
          // Paragraphs of different cells never merge (Word). After a table, an empty paragraph goes.
          if (!cur && prev && !paras[h.p].text.length && h.p < paras.length - 1) {
            this.snapshot();
            this.doc.splice(h.p, h.p, []);
            const end = { p: h.p - 1, o: this.doc.paras[h.p - 1].text.length };
            this.setSel(end, end);
            this.update({ scroll: true, inputTs });
          } else if (!cur && prev) {
            const end = { p: h.p - 1, o: paras[h.p - 1].text.length };
            this.setSel(end, end);
            this.update({ scroll: true, inputTs });
          }
          return;
        }
      }
      return this.deleteRange({ p: h.p - 1, o: this.doc.paras[h.p - 1].text.length }, h, inputTs);
    }
    const text = this.doc.paras[h.p].text;
    const start = word ? prevTokenStart(text, h.o) : h.o - charBefore(text, h.o);
    this.deleteRange({ p: h.p, o: start }, h, inputTs);
  }

  /** Effective paragraph formatting including the list level's indents. */
  private effectivePara(p: number): ResolvedPara {
    const para = this.doc.paras[p];
    const n = numberingOf(para.props, this.doc.styles);
    const lvl = n ? levelOf(this.doc.numbering, n.numId, n.ilvl)?.level : undefined;
    return this.doc.styles.resolvePara(para.props, lvl ? { level: lvl.pPr } : undefined);
  }

  /**
   * Rücktaste at the start of an indented paragraph (Word 365, measured): a first-line indent
   * goes first; then the left indent goes back to the previous paragraph's text or number
   * position below it, else to 0 (a "Listenabsatz" becomes "Standard"); only then do the
   * paragraphs join.
   */
  private backspaceIndent(p: number): boolean {
    if (!this.autoCorrectOn || this.hasSelection) return false;
    const para = this.doc.paras[p];
    if (para.block || this.doc.settings.trackChanges) return false;
    const rp = this.doc.styles.resolvePara(para.props);
    if (rp.indFirst > 0.01) {
      this.edit(() => this.doc.mapParas(p, p, (old) => ({ ...old, indFirst: 0 })));
      return true;
    }
    if (rp.indLeft <= 0.01 && rp.indFirst >= -0.01) return false;
    let left = 0;
    const prev = p > 0 ? this.doc.paras[p - 1] : null;
    if (prev && !prev.block && !!prev.cell === !!para.cell) {
      const pr = this.effectivePara(p - 1);
      const stops = [pr.indLeft, pr.indLeft + Math.min(0, pr.indFirst)].filter((x) => x >= 0 && x < rp.indLeft - 0.05);
      if (stops.length) left = Math.max(...stops);
    }
    const listPara = para.props.style === this.styleId("List Paragraph");
    this.edit(() => this.doc.mapParas(p, p, (old) => {
      const next = { ...old, style: listPara ? undefined : old.style, indLeft: undefined, indFirst: undefined };
      const base = this.doc.styles.resolvePara(next);
      return { ...next, indLeft: Math.abs(base.indLeft - left) > 0.01 ? left : undefined, indFirst: Math.abs(base.indFirst) > 0.01 ? 0 : undefined };
    }));
    return true;
  }

  private deleteForward(word: boolean, inputTs: number): void {
    const [from, to] = this.range();
    if (cmp(from, to) !== 0) { const [a, b] = this.smartRange(from, to); return this.deleteRange(a, b, inputTs); }
    const h = this.sel.head;
    const text = this.doc.paras[h.p].text;
    if (h.o === text.length) {
      const paras = this.doc.paras;
      if (h.p === paras.length - 1 || paras[h.p + 1].block || paras[h.p].block) return;
      const cur = paras[h.p].cell, next = paras[h.p + 1].cell;
      if (cur || next) {
        const same = cur && next && cur.table.id === next.table.id && cur.row === next.row && cur.col === next.col;
        if (!same) {
          // An empty paragraph before a table goes; otherwise its text moves into the first
          // cell, in front of the cell's text (Word 365, measured).
          if (!cur && next && !text.length) {
            this.snapshot();
            this.doc.splice(h.p, h.p, []);
            this.setSel({ p: h.p, o: 0 }, { p: h.p, o: 0 });
            this.update({ scroll: true, inputTs });
          } else if (!cur && next && !this.doc.settings.trackChanges) {
            this.snapshot();
            const b = paras[h.p + 1];
            const merged = { ...joinParas(paras[h.p], b), props: b.props, cell: b.cell, mark: b.mark };
            this.doc.splice(h.p, h.p + 1, [merged]);
            this.setSel({ p: h.p, o: text.length }, { p: h.p, o: text.length });
            this.update({ scroll: true, inputTs });
          }
          return;
        }
      }
      return this.deleteRange(h, { p: h.p + 1, o: 0 }, inputTs, true);
    }
    let end = h.o + charAfter(text, h.o);
    if (word) {
      // Inside a word, Strg+Entf deletes only the rest of the word, not the space after it.
      end = nextTokenStart(text, h.o);
      if (isWordChar(text[h.o - 1]) && isWordChar(text[h.o])) {
        end = h.o;
        while (end < text.length && isWordChar(text[end])) end++;
      }
    }
    this.deleteRange(h, { p: h.p, o: end }, inputTs, true);
  }

  // ----- commands for the shell -----------------------------------------------------------

  selectAll(): void {
    this.setSel({ p: 0, o: 0 }, this.doc.end());
    this.update();
  }

  /** Puts the selection on the internal clipboard (with formatting) and returns its text. */
  copySelection(): string {
    const [from, to] = this.range();
    let text = this.selectedText();
    const paras = this.doc.fragment(from, to).map((p) => this.portable(p));
    // Whole paragraphs up to the end of the text: Word copies the last paragraph mark too, so
    // pasting makes them paragraphs of their own.
    const end = this.doc.end();
    if (from.o === 0 && cmp(to, end) === 0 && this.doc.paras[to.p].text.length && !this.doc.paras[to.p].cell) {
      paras.push({ ...makePara(""), props: {}, mark: this.doc.paras[to.p].mark });
      text += "\n";
    }
    clipboardFragment = { text, paras, source: this };
    return text;
  }

  /**
   * A copy of a paragraph piece that is safe to paste elsewhere: no bookmarks, and objects
   * that must stay unique (comment references, revisions …) become their plain text.
   */
  private portable(p: Para, keep = (k: string) => k === "field" || k === "link" || k === "image"): Para {
    if (!(p.objects ?? []).some((o) => !keep(o.kind))) return { ...p, anchors: undefined };
    let text = "";
    const runs: { len: number; props: CharProps }[] = [];
    const objects: NonNullable<Para["objects"]>[number][] = [];
    let k = 0, pos = 0;
    for (const r of p.runs) {
      let piece = "";
      for (let i = pos; i < pos + r.len; i++) {
        const ch = p.text[i];
        if (ch !== OBJ) { piece += ch; continue; }
        const o = p.objects?.[k++];
        if (o && keep(o.kind)) { piece += OBJ; objects.push(o); }
        else piece += (o?.display ?? "").split(OBJ).join("");
      }
      pos += r.len;
      if (piece) { text += piece; runs.push({ len: piece.length, props: r.props }); }
    }
    return { ...p, text, runs, objects: objects.length ? objects : undefined, anchors: undefined };
  }

  /**
   * Content copied in another document: images and links refer to that file's parts, so
   * they become their text; styles this document lacks are copied over, as Word does.
   */
  private adopt(p: Para, source: Editor): Para {
    const out = this.portable(p, (k) => k === "field");
    const ids = new Set<string>();
    if (out.props.style) ids.add(out.props.style);
    for (const r of out.runs) if (r.props.rStyle) ids.add(r.props.rStyle);
    if (out.mark.rStyle) ids.add(out.mark.rStyle);
    let sheet = this.doc.styles;
    for (const id of ids) {
      for (let st = source.doc.styles.get(id); st && !sheet.get(st.id); st = source.doc.styles.get(st.basedOn)) sheet = sheet.withStyle(st);
    }
    if (sheet !== this.doc.styles) this.doc.setStyles(sheet);
    return out;
  }

  /** Removes the selection and returns its text (for Ausschneiden). */
  cutSelection(): string {
    const text = this.copySelection();
    const [from, to] = this.smartRange(...this.range());
    this.deleteRange(from, to);
    return text;
  }

  /** Inserts clipboard text; formatted when it came from RaabeWrite, with Word's smart spacing. */
  pasteText(text: string, plainOnly = false): void {
    text = text.replace(/\r\n?/g, "\n");
    if (!text || this.inBlock()) return;
    this.typing = false;
    const [from] = this.range();
    const before = this.doc.paras[from.p].text[from.o - 1];
    const smartSpace = isWordChar(before) && isWordChar(text[0]);
    const frag = clipboardFragment;
    if (!plainOnly && frag && frag.text === text) {
      this.snapshot();
      const [a, b] = this.range();
      const pieces = frag.paras.map((p) => (frag.source === this ? { ...p } : this.adopt(p, frag.source)));
      // The last piece merges into the target paragraph and takes its formatting (as in Word).
      if (pieces.length > 1) pieces[pieces.length - 1] = { ...pieces[pieces.length - 1], props: this.doc.paras[b.p].props };
      const ownFirst = a.o === 0 && pieces.length > 1 ? pieces[0].props : null;
      let end: Pos;
      if (this.trackChanges) {
        // Inserted after the text it replaces, which stays as a deletion.
        let at = b;
        if (smartSpace) at = this.doc.insertText(b, b, " ", intern({ ...typedProps(this.doc.paras[b.p], b.o), rev: R.newRevision("ins", this.author) }));
        end = this.replaceTracked(at, at, pieces);
        if (cmp(a, b) !== 0) end = this.keepPos(end, () => R.markDeleted(this.doc, a, b, R.newRevision("del", this.author)));
      } else {
        let at = a;
        if (smartSpace) at = this.doc.insertText(a, b, " ", typedProps(this.doc.paras[a.p], a.o));
        end = this.doc.replace(at, smartSpace ? at : b, pieces);
      }
      // Pasted at a paragraph start, the first pasted paragraph keeps its own formatting.
      if (ownFirst) this.doc.mapParas(a.p, a.p, () => ownFirst);
      end = this.smartPasteEnd(end);
      this.setSel(end, end);
      this.requireFonts();
      this.update({ scroll: true });
      return;
    }
    this.replaceSelection(smartSpace ? ` ${text}` : text);
    const fixed = this.smartPasteEnd(this.sel.head);
    if (cmp(fixed, this.sel.head) !== 0) { this.setSel(fixed, fixed); this.update({ scroll: true }); }
  }

  /** Smart paste: a pasted trailing space before punctuation or another space goes (Word). */
  private smartPasteEnd(end: Pos): Pos {
    const text = this.doc.paras[end.p]?.text ?? "";
    if (end.o > 0 && text[end.o - 1] === " " && /^[ .,;:!?)\]]/.test(text[end.o] ?? "") && !this.doc.settings.trackChanges) {
      this.doc.delete({ p: end.p, o: end.o - 1 }, end);
      return { p: end.p, o: end.o - 1 };
    }
    return end;
  }

  /**
   * Inserts a line break (Umschalt+Eingabe), page break (Strg+Eingabe) or column break
   * (Strg+Umschalt+Eingabe), as Word 365 does (measured): a page break stands in a paragraph
   * of its own and the caret goes to the start of the text after it, on the next page; a
   * column break starts the paragraph after the caret, the caret behind it.
   */
  insertBreak(kind: "line" | "page" | "column"): void {
    this.typing = false;
    if (kind === "line") { this.replaceSelection(LINE_BREAK); return; }
    const atStart = this.range()[0].o === 0;
    if (kind === "page") this.replaceSelection(atStart ? `${PAGE_BREAK}\n` : `\n${PAGE_BREAK}\n`);
    else this.replaceSelection(atStart ? COLUMN_BREAK : `\n${COLUMN_BREAK}`);
  }

  /**
   * AutoFormat "Absatzeinzug mit Tab- und Rücktaste festlegen": Tab at the start of a
   * paragraph that has text indents its first line by one default tab stop instead of
   * typing a tab; again, the whole paragraph.
   */
  private firstIndentTab(): boolean {
    if (!this.autoCorrectOn || this.hasSelection) return false;
    const h = this.sel.head;
    const para = this.doc.paras[h.p];
    if (h.o !== 0 || !para.text.length || para.cell || para.block || numberingOf(para.props, this.doc.styles)) return false;
    const rp = this.doc.styles.resolvePara(para.props);
    const step = this.doc.settings.defaultTab;
    this.snapshot();
    if (rp.indFirst < step - 0.01) this.doc.mapParas(h.p, h.p, (old) => ({ ...old, indFirst: step }));
    else this.doc.mapParas(h.p, h.p, (old) => ({ ...old, indLeft: rp.indLeft + step }));
    this.update({ scroll: true });
    return true;
  }

  /**
   * Tab: in a table to the next cell (its text selected; a new row after the last cell),
   * at the start of a list item one level deeper. Umschalt+Tab goes back. Returns false when
   * a tab character should be typed instead.
   */
  private tabKey(back: boolean): boolean {
    const h = this.sel.head;
    const paras = this.doc.paras;
    if (paras[h.p].cell) {
      let next = E.adjacentCell(this.doc, h.p, back ? -1 : 1);
      if (!next && !back) {
        this.snapshot();
        const pos = E.insertRows(this.doc, h.p, h.p, "below");
        if (pos) next = E.cellRange(this.doc, pos.p);
      }
      if (next) {
        const a = { p: next[0], o: 0 }, b = { p: next[1], o: this.doc.paras[next[1]].text.length };
        this.setSel(a, b);
        this.update({ scroll: true });
      }
      return true;
    }
    const [from, to] = this.range();
    const atStart = from.o === 0 && (from.p !== to.p || to.o === 0 || to.o === paras[to.p].text.length);
    if (atStart && numberingOf(paras[from.p].props, this.doc.styles)) {
      this.snapshot();
      // At the list's first item Word moves the whole list instead of changing the level.
      const moved = from.p === to.p && E.firstOfList(this.doc, from.p) ? E.shiftList(this.doc, from.p, back) : E.changeListLevel(this.doc, from.p, to.p, back ? -1 : 1);
      if (!moved) this.undoStack.pop();
      this.update({ scroll: true });
      return true;
    }
    return false;
  }

  // ----- lists ------------------------------------------------------------------------------

  /** Aufzählungszeichen, Nummerierung, Liste mit mehreren Ebenen (with optional levels). */
  toggleList(kind: E.ListKind, levels?: import("./model").NumLevel[]): void {
    const [a, b] = this.paraRangeIdx();
    this.edit(() => E.toggleList(this.doc, a, b, kind, this.listStyleId(), levels), false);
  }

  /**
   * Strg+Umschalt+L: the style "Aufzählungszeichen" (List Bullet), with the single-level
   * bullet numbering Word gives it (Symbol bullet, 0.63 cm hanging).
   */
  applyListBullet(): void {
    const [a, b] = this.paraRangeIdx();
    this.snapshot();
    this.typing = false;
    const id = this.styleId("List Bullet");
    let def = this.doc.styles.get(id);
    let numId = def?.pPr.numId;
    if (!def || numId === undefined || !this.doc.numbering.nums.has(numId)) {
      const made = addList(this.doc.numbering, [{ start: 1, format: "bullet", text: "\uf0b7", align: "left", suffix: "tab", pPr: { indLeft: 18, indFirst: -18, tabs: [{ pos: 18, align: "num" }] }, rPr: { font: "Symbol" } }]);
      this.doc.setNumbering(made.numbering);
      numId = made.numId;
      def = { id, name: "List Bullet", type: "paragraph", basedOn: this.styleId("Normal"), uiPriority: 99, pPr: { ...(def?.pPr ?? {}), numId, contextual: true }, rPr: def?.rPr ?? {} };
      this.doc.setStyles(this.doc.styles.withStyle(def));
    }
    this.doc.mapParas(a, b, (old) => ({ style: id, extra: old.extra }));
    this.update({ scroll: true });
  }

  listKind(): E.ListKind | null {
    return E.listKind(this.doc, this.doc.paras[this.sel.head.p]);
  }

  /** Listenebene ändern: the selected list items go to level `ilvl` (0-8). */
  setListLevel(ilvl: number): void {
    const [a, b] = this.paraRangeIdx();
    this.edit(() => this.doc.mapParas(a, b, (old, para) => {
      const n = numberingOf(para.props, this.doc.styles);
      return n ? { ...old, numId: old.numId ?? n.numId, ilvl: Math.max(0, Math.min(8, ilvl)) } : old;
    }), false);
  }

  changeListLevel(delta: 1 | -1): void {
    const [a, b] = this.paraRangeIdx();
    this.edit(() => E.changeListLevel(this.doc, a, b, delta), false);
  }

  // ----- tables ------------------------------------------------------------------------------

  /** Width of the text area at the caret (points), for new tables. */
  private textWidth(): number {
    const pl = this.view.placedFor(this.sel.head, this.sel.headUpstream).pl;
    const fr = this.layout.frames[pl.frame];
    return (fr ? fr.w : 453.6 * PT) / PT;
  }

  inTable(): boolean {
    return !!this.doc.paras[this.sel.head.p]?.cell;
  }

  /** The table definition at the caret. */
  tableDef(): TableDef | undefined {
    return this.doc.paras[this.sel.head.p]?.cell?.table;
  }

  /** The cell definition at the caret. */
  cellDef(): CellDef | undefined {
    const c = this.doc.paras[this.sel.head.p]?.cell;
    return c ? c.table.rows[c.row].cells[c.col] : undefined;
  }

  rowDef(): RowDef | undefined {
    const c = this.doc.paras[this.sel.head.p]?.cell;
    return c ? c.table.rows[c.row] : undefined;
  }

  /** Tabelle einfügen: rows × cols in "Tabellenraster", caret in the first cell. */
  insertTable(rows: number, cols: number): void {
    if (this.inTable()) return;
    this.snapshot();
    const style = this.styleId("Table Grid");
    this.ensureStyle(style);
    const [from, to] = this.range();
    if (cmp(from, to) !== 0) this.doc.delete(from, to);
    const at = E.insertTable(this.doc, from, rows, cols, this.textWidth(), style, this.typingProps());
    this.setSel(at, at);
    this.requireFonts();
    this.update({ scroll: true });
  }

  /** Runs a table operation as one undo step; it returns where the caret goes. */
  private tableOp(fn: (p: number, q: number) => Pos | null | void): void {
    const [from, to] = this.range();
    if (!this.doc.paras[from.p]?.cell) return;
    this.snapshot();
    const pos = fn(from.p, to.p);
    if (pos) this.setSel(pos, pos);
    this.update({ scroll: true });
  }

  insertRows(where: E.RowWhere): void { this.tableOp((p, q) => E.insertRows(this.doc, p, q, where)); }
  insertColumns(where: E.ColWhere): void { this.tableOp((p, q) => E.insertColumns(this.doc, p, q, where)); }
  deleteRows(): void { this.tableOp((p, q) => E.deleteRows(this.doc, p, q)); }
  deleteColumns(): void { this.tableOp((p, q) => E.deleteColumns(this.doc, p, q)); }
  deleteTable(): void { this.tableOp((p) => E.deleteTable(this.doc, p)); }
  mergeCells(): void { this.tableOp((p, q) => E.mergeCells(this.doc, p, q)); }
  splitCell(n: number): void { this.tableOp((p) => E.splitCell(this.doc, p, n)); }
  distributeColumns(): void { this.tableOp((p) => { E.distributeColumns(this.doc, p); }); }

  /** Changes the selected cells' definitions (Schattierung, Rahmen, Ausrichtung). */
  updateCells(change: (c: CellDef, row: number, col: number) => CellDef): void {
    const keep = this.sel;
    this.tableOp((p, q) => { E.updateCells(this.doc, p, q, change); });
    this.sel = keep;
    this.update();
  }

  updateRows(change: (r: RowDef) => RowDef): void {
    const keep = this.sel;
    this.tableOp((p, q) => { E.updateRows(this.doc, p, q, change); });
    this.sel = keep;
    this.update();
  }

  updateTable(change: (def: TableDef) => TableDef): void {
    const keep = this.sel;
    this.tableOp((p) => { E.updateTable(this.doc, p, change); });
    this.sel = keep;
    this.requireFonts();
    this.update();
  }

  /**
   * Start > Sortieren: the selected paragraphs (all of the text without a selection) or, in a
   * table, its rows (the selected ones, else all). Section breaks stay where they are.
   * Returns why nothing was sorted, or null.
   */
  /**
   * What Sortieren would sort, for the dialog: table rows or paragraphs, each as
   * [whole text, field 1, …]; `headerRow` when the table's first row repeats as a heading.
   */
  sortPreview(separator = "\t"): { table: boolean; entries: string[][]; headerRow: boolean } {
    const [from, to] = this.range();
    const paras = this.doc.paras;
    if (paras[from.p]?.cell) {
      const t = E.tableAt(this.doc, from.p);
      if (!t) return { table: true, entries: [], headerRow: false };
      const r1 = paras[from.p].cell!.row, r2 = paras[to.p]?.cell?.table.id === t.def.id ? paras[to.p].cell!.row : r1;
      const [lo, hi] = r1 === r2 ? [0, t.def.rows.length - 1] : [Math.min(r1, r2), Math.max(r1, r2)];
      const entries = E.tableRowTexts(this.doc, from.p).slice(lo, hi + 1).map((cells) => [cells.join(" "), ...cells]);
      return { table: true, entries, headerRow: lo === 0 && !!t.def.rows[0]?.header };
    }
    let a = from.p, b = to.p;
    if (cmp(from, to) === 0) { a = 0; b = paras.length - 1; }
    else if (to.o === 0 && b > a) b--;
    return { table: false, entries: paras.slice(a, b + 1).filter((p) => !p.cell && !p.block).map((p) => [p.text, ...p.text.split(separator)]), headerRow: false };
  }

  sortText(spec: SortSpec): string | null {
    const [from, to] = this.range();
    const paras = this.doc.paras;
    const sep = spec.separator || "\t";
    if (paras[from.p]?.cell) {
      const t = E.tableAt(this.doc, from.p);
      if (!t) return null;
      if (t.def.rows.some((r) => r.cells.some((c) => c.vMerge))) return "Tabellen mit vertikal verbundenen Zellen können nicht sortiert werden.";
      const r1 = paras[from.p].cell!.row, r2 = paras[to.p]?.cell?.table.id === t.def.id ? paras[to.p].cell!.row : r1;
      const [lo, hi] = r1 === r2 ? [0, t.def.rows.length - 1] : [Math.min(r1, r2), Math.max(r1, r2)];
      const texts = E.tableRowTexts(this.doc, from.p).slice(lo, hi + 1);
      const order = sortOrder(texts.map((cells) => [cells.join(" "), ...cells]), spec);
      if (order.every((x, i) => x === i)) return null;
      this.snapshot();
      E.reorderRows(this.doc, from.p, lo, order);
      const pos = E.firstOfCell(this.doc, t.a, lo, 0);
      this.setSel(pos, pos);
      this.update({ scroll: true });
      return null;
    }
    let a = from.p, b = to.p;
    if (cmp(from, to) === 0) { a = 0; b = paras.length - 1; }
    else if (to.o === 0 && b > a) b--;
    for (let i = a; i <= b; i++) {
      if (paras[i].cell || paras[i].block) return "Die Markierung enthält eine Tabelle. Markieren Sie nur Absätze oder nur Zeilen einer Tabelle.";
    }
    const list = paras.slice(a, b + 1);
    const order = sortOrder(list.map((p) => [p.text, ...p.text.split(sep)]), spec);
    if (order.every((x, i) => x === i)) return null;
    this.snapshot();
    // Sections end where they ended: the breaks stay at their places.
    const sects = list.map((p) => p.sect);
    this.doc.splice(a, b, order.map((src, i) => {
      const { sect: _s, ...rest } = list[src];
      return sects[i] ? { ...rest, sect: sects[i] } : rest;
    }));
    this.setSel({ p: a, o: 0 }, { p: b, o: this.doc.paras[b].text.length });
    this.update({ scroll: true });
    return null;
  }

  /** Start > Markieren > Objekte auswählen: the pointer selects objects (also behind the text) only; Esc ends it. */
  objectSelect = false;

  setObjectSelect(on: boolean): void {
    this.objectSelect = on;
    this.workspace.classList.toggle("objselect", on);
    this.emit(false);
  }

  /** Selects the whole table / the current row / column / cell ("Auswählen"). */
  selectTablePart(part: "table" | "row" | "column" | "cell"): void {
    const h = this.sel.head;
    const t = E.tableAt(this.doc, h.p);
    if (!t) return;
    const c = this.doc.paras[h.p].cell!;
    let a = t.a, b = t.b;
    if (part === "cell") [a, b] = E.cellRange(this.doc, h.p);
    else if (part === "row") {
      a = t.b; b = t.a;
      for (let i = t.a; i <= t.b; i++) if (this.doc.paras[i].cell!.row === c.row) { a = Math.min(a, i); b = Math.max(b, i); }
    } else if (part === "column") {
      a = t.b; b = t.a;
      for (let i = t.a; i <= t.b; i++) if (this.doc.paras[i].cell!.col === c.col) { a = Math.min(a, i); b = Math.max(b, i); }
    }
    this.setSel({ p: a, o: 0 }, { p: b, o: this.doc.paras[b].text.length });
    this.update();
  }

  /** Moves a column border by dx points (dragging it with the mouse). */
  resizeColumn(p: number, g: number, dx: number, snapshot: boolean): void {
    if (snapshot) this.snapshot();
    E.resizeColumn(this.doc, p, g, dx);
    this.update();
  }

  // ----- sections, fields, pictures --------------------------------------------------------------

  /** Abschnittsumbrüche: Nächste Seite, Fortlaufend, Gerade Seite, Ungerade Seite. */
  insertSectionBreak(start: import("./model").Section["start"]): void {
    if (this.doc.activeStory !== MAIN || this.inTable()) return;
    this.snapshot();
    const [from, to] = this.range();
    if (cmp(from, to) !== 0) this.doc.delete(from, to);
    const at = E.insertSectionBreak(this.doc, from, start, this.typingProps());
    this.setSel(at, at);
    this.update({ scroll: true });
  }

  /** The section of the caret (main text) or of the page shown (header/footer). */
  sectionIndex(): number {
    if (this.doc.activeStory === MAIN) return E.sectionIndex(this.doc, this.sel.head.p);
    const page = Number(this.storyKey.split("@")[1] ?? 0);
    return this.layout.pages[page]?.section ?? 0;
  }

  sectionCount(): number {
    return E.sectionCount(this.doc);
  }

  /** The current section's properties. */
  section(): import("./model").Section {
    const si = this.sectionIndex();
    let k = 0;
    const main = this.doc.main;
    for (let i = 0; i < main.length - 1; i++) {
      if (!main[i].sect) continue;
      if (k === si) return main[i].sect!;
      k++;
    }
    return this.doc.settings.section;
  }

  /** Seite einrichten: changes the current section, or all sections. */
  changeSections(change: (s: import("./model").Section) => import("./model").Section, all: boolean): void {
    this.snapshot();
    if (all) for (let si = 0; si < this.sectionCount(); si++) this.updateSection(si, change);
    else this.updateSection(this.sectionIndex(), change);
    this.update({ scroll: true });
  }

  /** Erste Seite anders / Untersch. gerade und ungerade Seiten. */
  setTitlePage(on: boolean): void {
    this.changeSections((s) => ({ ...s, titlePage: on }), false);
  }

  /** Formeloptionen of the document (Word's defaults when the file has none). */
  mathSettings(): MathSettings {
    return this.doc.settings.math ?? DEFAULT_MATH;
  }

  setMathSettings(math: MathSettings, record = true): void {
    const before = this.mathSettings();
    if (record) this.snapshot();
    this.doc.setSettings({ ...this.doc.settings, math });
    // Nested fractions drawn smaller or not: every equation is measured again.
    if (before.smallFrac !== math.smallFrac) this.relayout();
    else this.update();
  }

  setEvenOdd(on: boolean): void {
    this.snapshot();
    this.doc.setSettings({ ...this.doc.settings, evenOdd: on || undefined });
    this.update();
  }

  /** Inserts a field (Seitenzahl, Seitenanzahl, Datum …) at the caret. */
  insertField(instr: string, display: string): void {
    this.snapshot();
    const [from, to] = this.range();
    const end = E.insertObject(this.doc, from, to, E.fieldObject(instr, display), this.typingProps());
    this.setSel(end, end);
    this.update({ scroll: true });
  }

  /** Inserts a picture in the line at the caret, scaled down to fit the text width. */
  insertImage(image: ImageData, width: number, height: number): void {
    this.snapshot();
    const max = this.textWidth();
    const k = width > max ? max / width : 1;
    const [from, to] = this.range();
    const end = E.insertObject(this.doc, from, to, E.imageObject(image, width * k, height * k), this.typingProps());
    this.setSel({ p: end.p, o: end.o - 1 }, end);
    this.update({ scroll: true });
  }

  /** Header/footer story of the current page, created when missing, as Word's gallery does. */
  ensureHeaderFooter(kind: "header" | "footer"): string {
    const page = this.isHeaderFooter() ? Number(this.storyKey.split("@")[1] ?? 0) : this.view.caretBox(this.sel.head).page;
    const pg = this.layout.pages[page];
    const existing = kind === "header" ? pg.header : pg.footer;
    if (existing) return existing;
    this.snapshot();
    return this.newHeaderFooter(pg.section, kind, page);
  }

  /** Replaces a header's or footer's text (gallery entries, Seitenzahlen). */
  setStoryParas(id: string, paras: Para[]): void {
    this.snapshot();
    this.doc.setStory(id, paras);
    this.requireFonts();
    this.update();
  }

  /** Renders pages (all, or the given zero-based ones) into `target` for printing. */
  renderForPrint(target: HTMLElement, pages?: readonly number[]): void {
    this.view.renderForPrint(target, pages);
  }

  // ----- character formatting -------------------------------------------------------------

  /**
   * The range a character format applies to: the selection, or – with a collapsed caret
   * inside a word – that word (Word's rule). null means "format the next typed text".
   */
  private formatRange(withSpaces = false): [Pos, Pos] | null {
    if (this.hasSelection) return this.range();
    const h = this.sel.head;
    const text = this.doc.paras[h.p].text;
    if (!isWordChar(text[h.o - 1]) || !isWordChar(text[h.o])) return null;
    let a = h.o, b = h.o;
    while (a > 0 && isWordChar(text[a - 1])) a--;
    while (b < text.length && isWordChar(text[b])) b++;
    // Hiding a word also hides the spaces after it (Word), so no double space remains.
    if (withSpaces) while (b < text.length && text[b] === " ") b++;
    return [{ p: h.p, o: a }, { p: h.p, o: b }];
  }

  /** Sets one character property, removing direct formatting where it equals the style. */
  private setCharProp<K extends keyof CharProps>(key: K, value: CharProps[K] | undefined): void {
    const change = (p: number) => (old: CharProps): CharProps => {
      const without = { ...old, [key]: undefined } as CharProps;
      if (value === undefined) return without;
      const styleVal = this.resolvedAt(p, without)[key as keyof ResolvedChar];
      const resolvedNew = this.resolvedAt(p, { ...without, [key]: value })[key as keyof ResolvedChar];
      return styleVal === resolvedNew && key !== "font" && key !== "size" ? without : { ...without, [key]: value };
    };
    const r = this.formatRange(key === "hidden");
    if (!r) {
      const h = this.sel.head;
      const base = this.pending ?? typedProps(this.doc.paras[h.p], h.o);
      const pending = intern(change(h.p)(base));
      this.pending = pending;
      if (!this.doc.paras[h.p].text.length) {
        // In an empty paragraph the mark carries the formatting (and the line height).
        this.edit(() => this.doc.mapChars(h, h, () => pending), false);
        this.pending = pending;
      }
      this.emit(false);
      return;
    }
    const [from, to] = r;
    this.edit(() => this.doc.mapChars(from, to, (old, p) => change(p)(old)), false);
  }

  toggleChar(key: "bold" | "italic" | "strike" | "dstrike" | "caps" | "smallCaps" | "hidden"): void {
    const s = this.charState();
    this.setCharProp(key, !s[key]);
  }

  /** Unterstreichen: toggles, or sets a specific kind from the menu. */
  setUnderline(kind?: string): void {
    const s = this.charState();
    this.setCharProp("underline", kind ?? (s.underline && s.underline !== "none" ? "none" : "single"));
  }

  toggleVertAlign(kind: "superscript" | "subscript"): void {
    const s = this.charState();
    this.setCharProp("vertAlign", s[kind] ? "baseline" : kind);
  }

  setFont(name: string): void {
    this.setCharProp("fontTheme", undefined);
    this.setCharPropKeep("font", name);
  }

  /** Font name changes always become direct formatting (and drop theme references). */
  private setCharPropKeep(key: "font" | "size", value: string | number): void {
    const r = this.formatRange();
    const change = (old: CharProps): CharProps => ({ ...old, fontTheme: key === "font" ? undefined : old.fontTheme, [key]: value });
    if (!r) {
      const h = this.sel.head;
      this.pending = intern(change(this.pending ?? typedProps(this.doc.paras[h.p], h.o)));
      if (!this.doc.paras[h.p].text.length) this.edit(() => this.doc.mapChars(h, h, change), false);
      this.pending = this.pending ?? null;
      this.fonts.require([String(value)]);
      this.emit(false);
      return;
    }
    this.edit(() => this.doc.mapChars(r[0], r[1], change), false);
  }

  setSize(pt: number): void {
    if (!(pt >= 1 && pt <= 1638)) return;
    this.setCharPropKeep("size", Math.round(pt * 2) / 2);
  }

  /** Schriftgrad vergrößern/verkleinern in Word's steps (each run separately). */
  growFont(dir: 1 | -1, onePoint = false): void {
    const step = (s: number) => {
      if (onePoint) return Math.max(1, s + dir);
      if (dir > 0) return GROW_STEPS.find((x) => x > s) ?? Math.floor(s / 10) * 10 + 10;
      return [...GROW_STEPS].reverse().find((x) => x < s) ?? Math.max(1, s - 1);
    };
    const r = this.formatRange();
    if (!r) {
      const h = this.sel.head;
      const base = this.pending ?? typedProps(this.doc.paras[h.p], h.o);
      this.pending = intern({ ...base, size: step(this.resolvedAt(h.p, base).size) });
      this.emit(false);
      return;
    }
    this.edit(() => this.doc.mapChars(r[0], r[1], (old, p) => ({ ...old, size: step(this.resolvedAt(p, old).size) })), false);
  }

  setColor(hex: string): void {
    this.setCharProp("colorTheme", undefined);
    this.setCharProp("color", hex);
  }

  /** Language of the text at the caret (Korrekturhilfesprache), when set directly. */
  charLang(): string | undefined {
    return this.typingProps().lang;
  }

  /** Marks the selection (or what is typed next) as being in a language ("de-DE", "en-GB" …). */
  setLang(lang: string): void {
    this.setCharPropKeepAll("lang", lang);
  }

  /** Sets a property as direct formatting on the selection, or for the next typed text. */
  private setCharPropKeepAll<K extends keyof CharProps>(key: K, value: CharProps[K]): void {
    if (!this.hasSelection) {
      const h = this.sel.head;
      this.pending = intern({ ...(this.pending ?? typedProps(this.doc.paras[h.p], h.o)), [key]: value });
      this.emit(false);
      return;
    }
    const [from, to] = this.range();
    this.edit(() => this.doc.mapChars(from, to, (old) => ({ ...old, [key]: value })), false);
  }

  setHighlight(name: string): void {
    this.setCharProp("highlight", name);
  }

  /** Strg+Leertaste: removes direct character formatting (character styles stay). */
  resetChar(): void {
    const r = this.formatRange();
    const keep = (old: CharProps): CharProps => intern({ rStyle: old.rStyle, lang: old.lang, extra: old.extra });
    if (!r) { this.pending = intern({}); this.emit(false); return; }
    this.edit(() => this.doc.mapChars(r[0], r[1], keep), false);
  }

  applyCharStyle(id: string): void {
    this.ensureStyle(id);
    const r = this.formatRange();
    if (!r) {
      const h = this.sel.head;
      this.pending = intern({ ...(this.pending ?? typedProps(this.doc.paras[h.p], h.o)), rStyle: id });
      this.emit(false);
      return;
    }
    this.edit(() => this.doc.mapChars(r[0], r[1], (old) => ({ ...old, rStyle: id })), false);
  }

  changeCase(mode: CaseMode | "cycle"): void {
    const r = this.formatRange() ?? null;
    if (!r) return;
    const [from, to] = r;
    const sample = this.doc.slice(from, to);
    if (mode === "cycle") {
      const letters = sample.replace(/[^\p{L}]/gu, "");
      const upper = letters === letters.toUpperCase();
      const lower = letters === letters.toLowerCase();
      mode = upper ? "lower" : lower ? "title" : "upper";
    }
    const m = mode;
    const per = (ch: string, up: boolean) => {
      const x = up ? ch.toUpperCase() : ch.toLowerCase();
      return x.length === 1 ? x : ch;
    };
    let wordStart = true;
    let sentenceStart = from.o === 0 || /[.!?]\s*$/.test(this.doc.paras[from.p].text.slice(0, from.o));
    this.edit(() => this.doc.mapText(from, to, (s) => {
      let out = "";
      for (const ch of s) {
        let o = ch;
        if (m === "upper") o = per(ch, true);
        else if (m === "lower") o = per(ch, false);
        else if (m === "toggle") o = per(ch, ch === ch.toLowerCase());
        else if (m === "title") o = per(ch, wordStart);
        else if (m === "sentence") o = per(ch, sentenceStart && /\p{L}/u.test(ch));
        if (/\p{L}/u.test(ch)) { wordStart = false; sentenceStart = false; }
        if (/\s/.test(ch)) wordStart = true;
        if (/[.!?]/.test(ch)) sentenceStart = true;
        out += o;
      }
      return out;
    }), false);
    // With the caret inside a word Word changes the word but leaves the caret where it was.
    if (this.hasSelection) this.setSel(r[0], r[1]);
    this.update();
  }

  // ----- paragraph formatting -------------------------------------------------------------

  private paraRangeIdx(): [number, number] {
    const [from, to] = this.range();
    return [from.p, to.p];
  }

  /** Sets paragraph properties on the selected paragraphs. */
  setParaProps(patch: Partial<ParaProps> | ((rp: ResolvedPara, old: ParaProps) => Partial<ParaProps>)): void {
    const [a, b] = this.paraRangeIdx();
    this.edit(() => this.doc.mapParas(a, b, (old) => {
      const rp = this.doc.styles.resolvePara(old);
      const changes = typeof patch === "function" ? patch(rp, old) : patch;
      return { ...old, ...changes };
    }), false);
  }

  setAlign(align: Align): void {
    // Word's toggles (real keystrokes, keys_check.py): the active alignment again returns to
    // left, and left again switches to Blocksatz.
    const cur = this.paraState().align;
    this.setParaProps({ align: cur !== align ? align : align === "left" ? "both" : "left" });
  }

  /** Einzug vergrößern/verkleinern: to the next/previous multiple of the default tab stop. */
  indent(dir: 1 | -1): void {
    const step = this.doc.settings.defaultTab;
    this.setParaProps((rp) => {
      const n = dir > 0 ? Math.floor(rp.indLeft / step + 1e-6) + 1 : Math.ceil(rp.indLeft / step - 1e-6) - 1;
      return { indLeft: Math.max(0, n * step) };
    });
  }

  /** Strg+T / Strg+Umschalt+T: hanging indent by one tab step. */
  hanging(dir: 1 | -1): void {
    const step = this.doc.settings.defaultTab;
    this.setParaProps((rp) => {
      const left = Math.max(0, rp.indLeft + dir * step);
      const firstAbs = rp.indLeft + rp.indFirst;
      return { indLeft: left, indFirst: Math.min(0, firstAbs - left) };
    });
  }

  setLineSpacing(rule: LineRule, value: number): void {
    this.setParaProps({ lineRule: rule, line: value });
  }

  setSpacing(before?: number, after?: number): void {
    const patch: { before?: number; after?: number } = {};
    if (before !== undefined) patch.before = before;
    if (after !== undefined) patch.after = after;
    this.setParaProps(patch);
  }

  /** Adds a built-in style the document lacks (e.g. "Überschrift 2" in an old file), with its base styles. */
  ensureStyle(id: string): void {
    for (let cur: string | undefined = id, guard = 0; cur && guard < 10; guard++) {
      if (this.doc.styles.get(cur)) { cur = this.doc.styles.get(cur)!.basedOn; continue; }
      const builtin = builtinStyle(BUILTIN_NAME_BY_ID[cur] ?? cur);
      if (!builtin) break;
      this.doc.setStyles(this.doc.styles.withStyle(builtin));
      cur = builtin.basedOn;
    }
  }

  /** Style id in this document for a built-in style name ("heading 1"). */
  styleId(name: string): string {
    return this.doc.styles.idForName(name) ?? builtinStyle(name)?.id ?? name;
  }

  /** Applies a paragraph style; direct paragraph formatting is removed (numbering stays). */
  applyParaStyle(id: string): void {
    const [a, b] = this.paraRangeIdx();
    this.snapshot();
    this.typing = false;
    this.ensureStyle(id);
    this.doc.mapParas(a, b, (old) => ({ style: id, extra: old.extra, outline: undefined }));
    this.requireFonts();
    this.update({ scroll: true });
  }

  // ----- Formatvorlagen (pane, Formatvorlage ändern, Neue Formatvorlage) ------------------------

  /**
   * The styles the Formatvorlagen pane lists: the document's and Word's built-ins it lacks,
   * with type, whether text uses them, Word's recommendation (priority) and the linked flag.
   */
  styleCatalog(): { id: string; name: string; type: StyleDef["type"]; inUse: boolean; recommended: boolean; builtin: boolean; priority: number; linked: boolean; qFormat: boolean }[] {
    const sheet = this.doc.styles;
    const builtins = defaultStyleSheet();
    const used = new Set<string>([sheet.defaultParagraphStyle]);
    const scan = (paras: readonly Para[]) => {
      for (const p of paras) {
        if (p.props.style) used.add(p.props.style);
        for (const r of p.runs) if (r.props.rStyle) used.add(r.props.rStyle);
      }
    };
    scan(this.doc.main);
    for (const ps of this.doc.stories.values()) scan(ps);
    const out = new Map<string, ReturnType<Editor["styleCatalog"]>[number]>();
    const add = (s: StyleDef, inDoc: boolean) => {
      if (s.type === "numbering" as string || out.has(s.id)) return;
      const builtin = !!builtins.idForName(s.name) || !!GERMAN_NAMES[s.name.toLowerCase()];
      out.set(s.id, {
        id: s.id, name: inDoc ? sheet.displayName(s.id) : builtins.displayName(s.id), type: s.type, inUse: inDoc && used.has(s.id),
        recommended: !s.semiHidden || used.has(s.id), builtin, priority: s.uiPriority ?? (s.isDefault ? 0 : builtin ? 99 : 100), linked: !!s.link && s.type === "paragraph", qFormat: !!s.qFormat,
      });
    };
    for (const s of sheet.styles.values()) add(s, true);
    for (const s of builtins.styles.values()) if (!sheet.idForName(s.name) && !sheet.get(s.id)) add(s, false);
    return [...out.values()];
  }

  /** A style's definition (a built-in one when the document lacks it). */
  styleDef(id: string): StyleDef | undefined {
    return this.doc.styles.get(id) ?? defaultStyleSheet().get(id);
  }

  /** What a style looks like: its paragraph and character formatting, resolved. */
  styleLook(id: string): { para: ResolvedPara; char: ResolvedChar } {
    this.ensureStyleQuietly(id);
    const sheet = this.doc.styles;
    const def = this.styleDef(id);
    if (def?.type === "character") {
      const para = sheet.resolvePara({ style: sheet.defaultParagraphStyle } as ParaProps);
      return { para, char: sheet.resolveChar(sheet.defaultParagraphStyle, { rStyle: id }) };
    }
    return { para: sheet.resolvePara({ style: id } as ParaProps), char: sheet.resolveChar(id, {}) };
  }

  /** The look of a style being edited (Formatvorlage ändern's preview), before it is stored. */
  styleLookOf(def: StyleDef): { para: ResolvedPara; char: ResolvedChar } {
    if (def.basedOn) this.ensureStyleQuietly(def.basedOn);
    const sheet = this.doc.styles.withStyle({ ...def, data: undefined });
    if (def.type === "character") {
      return { para: sheet.resolvePara({ style: sheet.defaultParagraphStyle } as ParaProps), char: sheet.resolveChar(sheet.defaultParagraphStyle, { rStyle: def.id }) };
    }
    return { para: sheet.resolvePara({ style: def.id } as ParaProps), char: sheet.resolveChar(def.id, {}) };
  }

  /** Formatvorlageninspektor: the caret paragraph's style and direct formatting, and its character's. */
  directFormatting(): { paraStyle: string; para: ParaProps; charStyle?: string; char: CharProps } {
    const p = this.doc.paras[this.sel.head.p];
    const { style, extra: _e, ...para } = p.props;
    const o = Math.max(0, Math.min(this.sel.head.o - (this.sel.head.o > 0 ? 1 : 0), p.text.length - 1));
    const run = propsAt(p, Math.max(0, o));
    const { rStyle, rev: _r, extra: _x, lang: _l, ...char } = run;
    return { paraStyle: style ?? this.doc.styles.defaultParagraphStyle, para, charStyle: rStyle, char };
  }

  /** Formatvorlageninspektor: removes the caret paragraph's direct paragraph or character formatting. */
  clearDirect(kind: "para" | "char"): void {
    const [a, b] = this.paraRangeIdx();
    if (kind === "para") { this.resetPara(); return; }
    this.edit(() => this.doc.mapChars({ p: a, o: 0 }, { p: b, o: this.doc.paras[b].text.length }, (old) => intern({ lang: old.lang, extra: old.extra, rStyle: old.rStyle, rev: old.rev })), false);
  }

  private ensureStyleQuietly(id: string): void {
    if (!this.doc.styles.get(id) && defaultStyleSheet().get(id)) this.ensureStyle(id);
  }

  /**
   * Formatvorlage ändern / Neue Formatvorlage: stores a style (written from the model from now
   * on). `apply` gives the selection the new style.
   */
  defineStyle(def: StyleDef, apply = false): void {
    this.snapshot();
    if (def.basedOn) this.ensureStyle(def.basedOn);
    this.doc.setStyles(this.doc.styles.withStyle({ ...def, data: undefined }));
    if (apply) {
      if (def.type === "character") this.doc.mapChars(...this.range(), (old) => intern({ ...old, rStyle: def.id }));
      else { const [a, b] = this.paraRangeIdx(); this.doc.mapParas(a, b, (old) => ({ style: def.id, extra: old.extra, numId: old.numId, ilvl: old.ilvl })); }
    }
    this.requireFonts();
    this.relayout();
  }

  /** A free style id for a new style's name ("Meine Formatvorlage" → "MeineFormatvorlage"). */
  newStyleId(name: string): string {
    const base = name.replace(/[^A-Za-z0-9ÄÖÜäöüß]/g, "") || "Formatvorlage";
    let id = base;
    for (let n = 2; this.doc.styles.get(id); n++) id = `${base}${n}`;
    return id;
  }

  /** Formatvorlage löschen: text in it gets the style it was based on (Standard). */
  removeStyle(id: string): void {
    const def = this.doc.styles.get(id);
    if (!def || def.isDefault) return;
    this.snapshot();
    const to = def.type === "paragraph" ? (def.basedOn && this.doc.styles.get(def.basedOn) ? def.basedOn : this.doc.styles.defaultParagraphStyle) : undefined;
    const fix = (paras: readonly Para[]): Para[] => paras.map((p) => {
      const runs = p.runs.some((r) => r.props.rStyle === id) ? p.runs.map((r) => (r.props.rStyle === id ? { ...r, props: intern({ ...r.props, rStyle: undefined }) } : r)) : p.runs;
      const props = p.props.style === id ? { ...p.props, style: to } : p.props;
      return runs === p.runs && props === p.props ? p : { ...p, runs, props };
    });
    const story = this.doc.activeStory;
    this.doc.setStory(MAIN, fix(this.doc.main));
    for (const [k, ps] of this.doc.stories) this.doc.setStory(k, fix(ps));
    this.doc.activeStory = story;
    // Styles based on it inherit what it was based on.
    let sheet = this.doc.styles.without(id);
    for (const s of sheet.styles.values()) {
      if (s.basedOn === id || s.next === id) sheet = sheet.withStyle({ ...s, basedOn: s.basedOn === id ? def.basedOn : s.basedOn, next: s.next === id ? undefined : s.next, data: undefined });
    }
    if (def.link) { const l = sheet.get(def.link); if (l) sheet = sheet.without(def.link); }
    this.doc.setStyles(sheet);
    this.relayout();
  }

  /**
   * Aktualisieren, um der Auswahl zu entsprechen: the style takes the caret paragraph's direct
   * formatting (and its first character's), which then is no longer direct.
   */
  matchStyleToSelection(id: string): void {
    this.ensureStyle(id);
    const def = this.doc.styles.get(id);
    if (!def) return;
    const p = this.doc.paras[this.sel.head.p];
    const { style: _s, numId: _n, ilvl: _i, extra: _e, ...direct } = p.props as ParaProps & { numId?: number; ilvl?: number };
    const run = p.runs.find((r) => r.len > 0)?.props ?? {};
    const { rStyle: _r, rev: _v, extra: _x, lang: _l, ...rdirect } = run;
    this.snapshot();
    const clean = <T extends object>(o: T) => Object.fromEntries(Object.entries(o).filter(([, v]) => v !== undefined)) as T;
    const next: StyleDef = def.type === "character"
      ? { ...def, rPr: clean({ ...def.rPr, ...rdirect }), data: undefined }
      : { ...def, pPr: clean({ ...def.pPr, ...direct }), rPr: clean({ ...def.rPr, ...rdirect }), data: undefined };
    this.doc.setStyles(this.doc.styles.withStyle(next));
    const [a, b] = this.paraRangeIdx();
    if (def.type === "paragraph") {
      this.doc.mapParas(a, b, (old) => ({ style: id, extra: old.extra, numId: (old as { numId?: number }).numId, ilvl: (old as { ilvl?: number }).ilvl } as ParaProps));
      const end = this.doc.paras[b].text.length;
      this.doc.mapChars({ p: a, o: 0 }, { p: b, o: end }, (old) => intern({ lang: old.lang, extra: old.extra, rStyle: old.rStyle, rev: old.rev }));
    }
    this.requireFonts();
    this.relayout();
  }

  // ----- Entwurf: Designs, Formatvorlagensatz, Farben, Schriftarten, Absatzabstand ---------------

  /** The document's theme (colours, fonts, names). */
  get theme(): Theme {
    return this.doc.styles.theme;
  }

  /** Designs / Farben / Schriftarten: a new theme; theme fonts and colours follow it. */
  setTheme(theme: Theme): void {
    this.snapshot();
    this.doc.setStyles(this.doc.styles.withTheme(theme));
    this.requireFonts();
    this.relayout();
  }

  /**
   * Formatvorlagensatz: the set's document defaults and styles replace the document's (by
   * their English names; styles the set lacks stay). The language of the text stays.
   */
  applyStyleSet(set: { defaultsR: CharProps; defaultsP: ParaProps; styles: Readonly<Record<string, { type: string; basedOn?: string; next?: string; pPr: ParaProps; rPr: CharProps }>> }): void {
    this.snapshot();
    let sheet = this.doc.styles;
    sheet = sheet.withDefaults({ ...set.defaultsR, lang: sheet.defaultsR.lang }, set.defaultsP);
    const idOf = (name: string | undefined) => (name ? sheet.idForName(name) ?? builtinStyle(name)?.id : undefined);
    for (const [name, s] of Object.entries(set.styles)) {
      const id = idOf(name);
      const cur = (id ? sheet.get(id) : undefined) ?? builtinStyle(name);
      if (!cur) continue;
      for (const n of [s.basedOn, s.next]) if (n && !sheet.idForName(n)) { const b = builtinStyle(n); if (b) sheet = sheet.withStyle(b); }
      sheet = sheet.withStyle({ ...cur, pPr: s.pPr, rPr: s.rPr, basedOn: s.basedOn ? sheet.idForName(s.basedOn) ?? cur.basedOn : undefined, next: s.next ? sheet.idForName(s.next) ?? cur.next : cur.next, data: undefined });
    }
    this.doc.setStyles(sheet);
    this.requireFonts();
    this.relayout();
  }

  /**
   * Absatzabstand: spacing of the document defaults (the Standard style keeps none of its
   * own), as Word's presets do.
   */
  setParagraphSpacing(sp: { before: number; after: number; line: number; lineRule: LineRule }): void {
    this.snapshot();
    let sheet = this.doc.styles;
    sheet = sheet.withDefaults(sheet.defaultsR, { ...sheet.defaultsP, before: sp.before, after: sp.after, line: sp.line, lineRule: sp.lineRule });
    const normal = sheet.get(sheet.defaultParagraphStyle);
    if (normal) {
      const { before: _b, after: _a, line: _l, lineRule: _r, ...pPr } = normal.pPr;
      sheet = sheet.withStyle({ ...normal, pPr, data: undefined });
    }
    this.doc.setStyles(sheet);
    this.relayout();
  }

  /** Wasserzeichen: set or remove; every section gets headers to carry it, as in Word. */
  setWatermark(wm: DocSettings["watermark"]): void {
    this.snapshot();
    if (wm) {
      // The first section needs its own headers (later ones inherit them).
      const firstSect = this.doc.main.slice(0, -1).find((p) => p.sect)?.sect ?? this.doc.settings.section;
      const need: ("default" | "first" | "even")[] = ["default", ...(firstSect.titlePage ? ["first" as const] : []), ...(this.doc.settings.evenOdd ? ["even" as const] : [])];
      const missing = need.filter((slot) => !firstSect.headers[slot] || !this.doc.stories.has(firstSect.headers[slot]!));
      if (missing.length) {
        const style = this.styleId("header");
        this.ensureStyle(style);
        const refs = { ...firstSect.headers };
        for (const slot of missing) {
          let n = 1;
          while (this.doc.stories.has(`hdr:${n}`)) n++;
          const id = `hdr:${n}`;
          this.doc.setStory(id, [{ ...makePara("", { style }), mark: EMPTY }]);
          refs[slot] = id;
        }
        this.updateSection(0, (s) => ({ ...s, headers: refs }));
      }
    }
    this.doc.setSettings({ ...this.doc.settings, watermark: wm });
    this.update();
  }

  /** Seitenfarbe: the page colour (undefined: Keine Farbe). */
  setBackground(bg: DocSettings["background"]): void {
    this.snapshot();
    this.doc.setSettings({ ...this.doc.settings, background: bg });
    this.update();
  }

  /** Strg+Q: removes direct paragraph formatting. */
  resetPara(): void {
    const [a, b] = this.paraRangeIdx();
    this.edit(() => this.doc.mapParas(a, b, (old) => ({ style: old.style, extra: old.extra })), false);
  }

  /** Alle Formatierungen löschen: Standard style, no direct formatting. */
  clearAllFormatting(): void {
    const r = this.hasSelection ? this.range() : null;
    const [a, b] = this.paraRangeIdx();
    this.edit(() => {
      if (r) this.doc.mapChars(r[0], r[1], (old) => intern({ lang: old.lang, extra: old.extra }));
      this.doc.mapParas(a, b, (old) => ({ style: this.doc.styles.defaultParagraphStyle, extra: old.extra }));
    }, false);
  }

  // ----- more list, table and paragraph commands -------------------------------------------

  /** Überschriften nummerieren: a list linked to Überschrift 1-3 (1 / 1.1 / 1.1.1). */
  numberHeadings(levels: NumLevel[]): void {
    this.snapshot();
    const made = addList(this.doc.numbering, levels);
    this.doc.setNumbering(made.numbering);
    for (let k = 1; k <= 3; k++) {
      const id = this.styleId(`heading ${k}`);
      this.ensureStyle(id);
      const def = this.doc.styles.get(id);
      if (def) this.doc.setStyles(this.doc.styles.withStyle({ ...def, pPr: { ...def.pPr, numId: made.numId, ilvl: k - 1 }, data: undefined }));
    }
    this.layouter.invalidate();
    this.update({ scroll: true });
  }

  /** Neu beginnen mit 1 / Separate Liste: the item at the caret and the rest of its list start again. */
  restartNumbering(start = 1): void {
    const h = this.sel.head;
    const n = numberingOf(this.doc.paras[h.p].props, this.doc.styles);
    if (!n) return;
    this.snapshot();
    const made = restartList(this.doc.numbering, n.numId, n.ilvl, start);
    this.doc.setNumbering(made.numbering);
    this.doc.mapParas(h.p, this.doc.paras.length - 1, (old, para) =>
      numberingOf(para.props, this.doc.styles)?.numId === n.numId ? { ...old, numId: made.numId, ilvl: old.ilvl ?? n.ilvl } : old);
    this.update({ scroll: true });
  }

  /** The list paragraph before the caret's list (of the same kind), for "Nummerierung fortsetzen". */
  private previousList(): { from: number; numId: number; prevNum: number } | null {
    const h = this.sel.head;
    const paras = this.doc.paras;
    const n = numberingOf(paras[h.p].props, this.doc.styles);
    if (!n) return null;
    const kind = E.listKind(this.doc, paras[h.p]);
    let i = h.p;
    while (i > 0 && numberingOf(paras[i - 1].props, this.doc.styles)?.numId === n.numId) i--;
    for (let k = i - 1; k >= 0; k--) {
      const m = numberingOf(paras[k].props, this.doc.styles);
      if (m && m.numId !== n.numId && E.listKind(this.doc, paras[k]) === kind) return { from: h.p, numId: n.numId, prevNum: m.numId };
    }
    return null;
  }

  /** Mit vorheriger Liste verknüpfen / Nummerierung fortsetzen. */
  canContinueNumbering(): boolean {
    return !!this.previousList();
  }

  continueNumbering(): void {
    const prev = this.previousList();
    if (!prev) return;
    this.edit(() => this.doc.mapParas(prev.from, this.doc.paras.length - 1, (old, para) =>
      numberingOf(para.props, this.doc.styles)?.numId === prev.numId ? { ...old, numId: prev.prevNum } : old));
  }

  /** The caret's list level: number position, text indent (points) and what follows the number. */
  listLevelIndent(): { numberPos: number; textIndent: number; suffix: "tab" | "space" | "nothing"; bullet: boolean } | null {
    const para = this.doc.paras[this.sel.head.p];
    const n = numberingOf(para.props, this.doc.styles);
    const found = n ? levelOf(this.doc.numbering, n.numId, n.ilvl) : null;
    if (!found) return null;
    const left = found.level.pPr.indLeft ?? 0, first = found.level.pPr.indFirst ?? 0;
    return { numberPos: left + first, textIndent: left, suffix: found.level.suffix, bullet: found.level.format === "bullet" };
  }

  /** Listeneinzug anpassen: changes the caret's list level (all items of that level). */
  setListLevelIndent(numberPos: number, textIndent: number, suffix: "tab" | "space" | "nothing"): void {
    const para = this.doc.paras[this.sel.head.p];
    const n = numberingOf(para.props, this.doc.styles);
    const found = n ? levelOf(this.doc.numbering, n.numId, n.ilvl) : null;
    if (!n || !found) return;
    this.edit(() => {
      const abs = found.abstract;
      const levels = abs.levels.map((l, i) => (i === n.ilvl ? { ...l, suffix, pPr: { ...l.pPr, indLeft: textIndent, indFirst: numberPos - textIndent }, data: undefined } : l));
      const abstracts = new Map(this.doc.numbering.abstracts);
      abstracts.set(abs.id, { ...abs, levels, data: undefined });
      this.doc.setNumbering({ ...this.doc.numbering, abstracts });
      // Items of the level with their own indent follow the new one.
      this.doc.mapParas(0, this.doc.paras.length - 1, (old, q) => {
        const m = numberingOf(q.props, this.doc.styles);
        return m && m.numId === n.numId && m.ilvl === n.ilvl && (old.indLeft !== undefined || old.indFirst !== undefined) ? { ...old, indLeft: undefined, indFirst: undefined } : old;
      });
    });
  }

  /** The selection covers every cell of its table. */
  wholeTableSelected(): boolean {
    const b = this.cellBlock();
    const def = this.doc.paras[this.sel.head.p]?.cell?.table;
    if (!b || !def) return false;
    return b.r1 === 0 && b.c1 === 0 && b.r2 === def.rows.length - 1 && b.c2 >= Math.max(...def.rows.map((r) => r.cells.length)) - 1;
  }

  /** The selected block of cells (cell indices), when the selection is in one table. */
  cellBlock(): { r1: number; r2: number; c1: number; c2: number; multi: boolean } | null {
    const a = this.doc.paras[this.sel.anchor.p]?.cell, b = this.doc.paras[this.sel.head.p]?.cell;
    if (!a || !b || a.table.id !== b.table.id) return null;
    const r1 = Math.min(a.row, b.row), r2 = Math.max(a.row, b.row), c1 = Math.min(a.col, b.col), c2 = Math.max(a.col, b.col);
    return { r1, r2, c1, c2, multi: r1 !== r2 || c1 !== c2 };
  }

  get showGridlines(): boolean {
    return this.view.showGridlines;
  }

  setShowGridlines(on: boolean): void {
    this.view.showGridlines = on;
    this.view.refresh();
    this.emit(false);
  }

  /** Width of the caret's cell (points). */
  cellWidth(): number {
    const c = this.doc.paras[this.sel.head.p]?.cell;
    if (!c) return 0;
    const row = c.table.rows[c.row];
    let gx = 0;
    for (let i = 0; i < c.col; i++) gx += Math.max(1, row.cells[i].span);
    return c.table.grid.slice(gx, gx + Math.max(1, row.cells[c.col].span)).reduce((s, x) => s + x, 0);
  }

  /** Breite: the caret's column gets this width (the table gets wider or narrower). */
  setCellWidth(pt: number): void {
    const c = this.doc.paras[this.sel.head.p]?.cell;
    if (!c) return;
    const row = c.table.rows[c.row];
    let gx = 0;
    for (let i = 0; i < c.col; i++) gx += Math.max(1, row.cells[i].span);
    const span = Math.max(1, row.cells[c.col].span);
    const cur = this.cellWidth();
    const k = pt / (cur || pt);
    this.updateTable((def) => ({
      ...def, grid: def.grid.map((g, i) => (i >= gx && i < gx + span ? g * k : g)),
      rows: def.rows.map((r) => ({ ...r, cells: r.cells.map((cell) => ({ ...cell, width: undefined })) })),
    }));
  }

  /** Rahmen for paragraphs: one side toggles, "all"/"outside" draw a box, "inside" the lines between. */
  paragraphBorders(kind: string): void {
    const [a, b] = this.paraRangeIdx();
    const line = (space: number): BorderSpec => ({ style: "single", size: 0.5, color: "auto", space });
    const sides = (kind === "all" || kind === "outside" ? ["top", "bottom", "left", "right"] : kind === "inside" ? ["between"] : [kind]) as (keyof Borders)[];
    const rps = this.doc.paras.slice(a, b + 1).map((p) => this.doc.styles.resolvePara(p.props));
    const allOn = kind !== "none" && sides.every((side) => rps.every((rp) => rp.borders?.[side] && rp.borders[side]!.style !== "nil"));
    this.setParaProps((rp) => {
      if (kind === "none") return { borders: undefined };
      const next: Record<string, BorderSpec | undefined> = { ...(rp.borders ?? {}) };
      for (const side of sides) next[side] = allOn && kind !== "all" && kind !== "outside" ? undefined : line(side === "left" || side === "right" ? 4 : 1);
      return { borders: Object.values(next).some((v) => v) ? next as Borders : undefined };
    });
  }

  /** Schattierung for paragraphs (or the selected cells in a table). */
  paragraphShading(fill: string): void {
    if (this.inTable()) {
      this.updateCells((cd) => ({ ...cd, shading: fill === "auto" ? undefined : { fill, val: "clear", color: "auto" } }));
      return;
    }
    this.setParaProps({ shading: fill === "auto" ? undefined : { fill, val: "clear", color: "auto" } });
  }

  // ----- headers and footers: contents -------------------------------------------------------

  /** The page the caret is on (or the header/footer instance being edited). */
  /** The page with the caret (or the header/footer being edited). */
  currentPage(): number {
    return this.isHeaderFooter() ? Number(this.storyKey.split("@")[1] ?? 0) : this.view.caretBox(this.sel.head, this.sel.headUpstream).page;
  }

  /** Opens the header or footer of the current page for editing. */
  editHeaderFooter(kind: "header" | "footer"): void {
    this.openHeaderFooter(this.currentPage(), kind);
  }

  /**
   * Replaces the header or footer of the current page with one paragraph of text and fields
   * (Word's gallery entries and Seitenzahlen) and opens it, caret at the end.
   */
  setHeaderFooterContent(kind: "header" | "footer", align: Align, parts: (string | { field: string; display: string })[]): void {
    const page = this.currentPage();
    const before = this.undoStack.length;
    const id = this.ensureHeaderFooter(kind);
    if (this.undoStack.length === before) this.snapshot();
    const style = this.styleId(kind);
    this.ensureStyle(style);
    let text = "";
    const objects: InlineObject[] = [];
    for (const part of parts) {
      if (typeof part === "string") text += part;
      else { text += OBJ; objects.push(E.fieldObject(part.field, part.display)); }
    }
    const para: Para = { ...makePara(text, { style, align: align === "left" ? undefined : align }), objects: objects.length ? objects : undefined };
    this.doc.setStory(id, [para]);
    this.requireFonts();
    this.update();
    const end = { p: 0, o: text.includes("[Hier eingeben]") ? 0 : text.length };
    this.switchStory(id, `${id}@${page}`, { anchor: end, head: text.includes("[Hier eingeben]") ? { p: 0, o: 15 } : end, headUpstream: false });
    this.emit(true);
  }

  /** Kopfzeile/Fußzeile entfernen: its text goes (the section keeps an empty one). */
  removeHeaderFooter(kind: "header" | "footer"): void {
    const page = this.currentPage();
    const pg = this.layout.pages[page];
    const id = kind === "header" ? pg?.header : pg?.footer;
    if (!id) return;
    this.snapshot();
    const style = this.styleId(kind);
    this.doc.setStory(id, [makePara("", { style })]);
    if (this.isHeaderFooter() && storyIdOf(this.storyKey) === id) this.sel = { anchor: { p: 0, o: 0 }, head: { p: 0, o: 0 }, headUpstream: false };
    this.update();
  }

  /** Seitenzahlen entfernen: page number fields leave every header and footer. */
  removePageNumbers(): void {
    this.snapshot();
    for (const [id, paras] of this.doc.stories) {
      if (!id.startsWith("hdr:") && !id.startsWith("ftr:")) continue;
      const next = paras.map((p) => {
        if (!p.objects?.some((o) => o.kind === "field" && /^(PAGE|NUMPAGES|SECTIONPAGES)\b/i.test((o.instr ?? "").trim()))) return p;
        let text = "", k = 0;
        const objects: InlineObject[] = [];
        for (const ch of p.text) {
          if (ch !== OBJ) { text += ch; continue; }
          const o = p.objects![k++];
          if (o.kind === "field" && /^(PAGE|NUMPAGES|SECTIONPAGES)\b/i.test((o.instr ?? "").trim())) continue;
          text += OBJ; objects.push(o);
        }
        if (/^\s*(Seite\s*(von\s*)?)?\s*$/.test(text)) { text = ""; objects.length = 0; }
        return { ...p, text, runs: text ? [{ len: text.length, props: p.runs[0]?.props ?? EMPTY }] : [], objects: objects.length ? objects : undefined };
      });
      this.doc.setStory(id, next);
    }
    if (this.isHeaderFooter()) this.clampSelection();
    this.update();
  }

  // ----- pictures ---------------------------------------------------------------------------

  /** A selected floating picture: its paragraph (main text) and index. */
  private floatSel: { p: number; index: number } | null = null;

  /** The selected picture: an inline one (one object selected) or a floating one. */
  selectedImage(): { kind: "inline"; p: number; o: number; obj: InlineObject } | { kind: "float"; p: number; index: number; obj: FloatObject } | null {
    if (this.floatSel && this.doc.activeStory === MAIN) {
      const f = this.doc.main[this.floatSel.p]?.floats?.[this.floatSel.index];
      if (f) return { kind: "float", p: this.floatSel.p, index: this.floatSel.index, obj: f };
    }
    const [from, to] = this.range();
    if (from.p !== to.p || to.o - from.o !== 1) return null;
    const para = this.doc.paras[from.p];
    if (para.text[from.o] !== OBJ) return null;
    const obj = this.objectAt(para, from.o);
    return obj?.kind === "image" ? { kind: "inline", p: from.p, o: from.o, obj } : null;
  }

  private objectAt(p: Para, o: number): InlineObject | undefined {
    let k = 0;
    for (let i = 0; i < o; i++) if (p.text[i] === OBJ) k++;
    return p.objects?.[k];
  }

  private replaceObject(p: number, o: number, obj: InlineObject): void {
    const para = this.doc.paras[p];
    let k = 0;
    for (let i = 0; i < o; i++) if (para.text[i] === OBJ) k++;
    const objects = [...(para.objects ?? [])];
    objects[k] = obj;
    const next = [...this.doc.paras];
    next[p] = { ...para, objects };
    this.doc.paras = next;
  }

  private replaceFloat(p: number, index: number, f: FloatObject | null): void {
    const para = this.doc.main[p];
    const floats = [...(para.floats ?? [])];
    if (f) floats[index] = f; else floats.splice(index, 1);
    const next = [...this.doc.main];
    next[p] = { ...para, floats: floats.length ? floats : undefined };
    this.doc.setStory(MAIN, next);
  }

  /** Größe: width and height of the selected picture (points). */
  setImageSize(width: number, height: number, snapshot = true): void {
    const img = this.selectedImage();
    // Lines may be flat (no height or no width); pictures need a few points.
    const min = img?.obj.shape ? 0 : 2;
    if (!img || !(width > min || (min === 0 && width >= 0)) || !(height > min || (min === 0 && height >= 0))) return;
    if (snapshot) this.snapshot();
    if (img.kind === "inline") this.replaceObject(img.p, img.o, { ...img.obj, width, height });
    else this.replaceFloat(img.p, img.index, { ...img.obj, width, height });
    this.update();
  }

  /** Bildformat settings of the selected picture ({} when none are set). */
  pictureFormat(): PictureFormat | null {
    const img = this.selectedImage();
    return img ? img.obj.pic ?? {} : null;
  }

  /**
   * Changes the selected picture's Bildformat settings: `change` gives new values (undefined
   * removes one) or maps the old settings. One undo step.
   */
  setPictureFormat(change: PictureFormat | ((f: PictureFormat) => PictureFormat), snapshot = true): void {
    const img = this.selectedImage();
    if (!img) return;
    const old = img.obj.pic ?? {};
    let merged: PictureFormat = typeof change === "function" ? change(old) : { ...old, ...change };
    // Changed effects take new room beside the picture (wp:effectExtent).
    const effects = ["line", "shadow", "glow", "reflection", "softEdge", "bevel", "preset", "rotation"] as const;
    if (typeof change !== "function" && !("extent" in change) && effects.some((k) => k in change)) {
      merged = { ...merged, extent: undefined };
      const ex = effectExtent(merged, img.obj.height ?? 0);
      if (ex.l || ex.t || ex.r || ex.b) merged = { ...merged, extent: ex };
    }
    const clean = Object.fromEntries(Object.entries(merged).filter(([, v]) => v !== undefined)) as PictureFormat;
    const pic = Object.keys(clean).length ? clean : undefined;
    if (snapshot) this.snapshot();
    this.typing = false;
    if (img.kind === "inline") this.replaceObject(img.p, img.o, { ...img.obj, pic });
    else this.replaceFloat(img.p, img.index, { ...img.obj, pic });
    this.update();
  }

  /** Alternativtext of the selected picture. */
  pictureAlt(): AltText | null {
    const img = this.selectedImage();
    return img ? img.obj.alt ?? {} : null;
  }

  setPictureAlt(alt: AltText): void {
    const img = this.selectedImage();
    if (!img) return;
    this.snapshot();
    // The name and Ausblenden (Auswahlbereich) stay.
    const keep = { name: img.obj.alt?.name, hidden: img.obj.alt?.hidden };
    const merged = Object.fromEntries(Object.entries({ ...alt, ...keep }).filter(([, v]) => v !== undefined && v !== "")) as AltText;
    const clean = Object.keys(merged).length ? merged : undefined;
    if (img.kind === "inline") this.replaceObject(img.p, img.o, { ...img.obj, alt: clean });
    else this.replaceFloat(img.p, img.index, { ...img.obj, alt: clean });
    this.update();
  }

  /**
   * Bild ändern: another picture in the place of the selected one, at the old width (Word
   * keeps the width and the new picture's proportions); settings except cropping stay.
   */
  replacePicture(image: ImageData, naturalW: number, naturalH: number): void {
    const img = this.selectedImage();
    if (!img) return;
    const w = (img.kind === "inline" ? img.obj.width : img.obj.width) ?? naturalW;
    const h = naturalW > 0 ? (w * naturalH) / naturalW : naturalH;
    const pic = img.obj.pic ? { ...img.obj.pic, crop: undefined } : undefined;
    this.snapshot();
    if (img.kind === "inline") this.replaceObject(img.p, img.o, { ...img.obj, image, width: w, height: h, pic, data: undefined });
    else this.replaceFloat(img.p, img.index, { ...img.obj, image, width: w, height: h, pic, data: undefined });
    this.update();
  }

  /** Bild zurücksetzen (settings) or Bild und Größe zurücksetzen (also the natural size). */
  resetPicture(naturalSize?: { w: number; h: number }): void {
    const img = this.selectedImage();
    if (!img) return;
    this.snapshot();
    const size = naturalSize ? { width: naturalSize.w, height: naturalSize.h } : {};
    if (img.kind === "inline") this.replaceObject(img.p, img.o, { ...img.obj, pic: undefined, ...size });
    else this.replaceFloat(img.p, img.index, { ...img.obj, pic: undefined, ...size });
    this.update();
  }

  /** The selected picture's overlay: its frame, and in Zuschneiden the whole picture. */
  private imageOverlay(): RenderOptions["image"] {
    const se = this.selectedShape();
    if (se?.editing) {
      const b = this.floatBox(se.p, se.index);
      return b ? { ...b, picture: true, editing: true } : undefined;
    }
    const box = this.imageBox();
    if (!box) { this.cropping = false; return undefined; }
    const img = this.selectedImage();
    const isPicture = img?.obj.image !== undefined;
    if (img?.obj.shape) {
      const ends = this.lineEnds(img.obj, box);
      return ends ? { ...box, picture: false, adjust: this.adjustHandles(img.obj, box), ends } : { ...box, picture: true, adjust: this.adjustHandles(img.obj, box) };
    }
    if (!this.cropping || !img?.obj.image) return { ...box, picture: isPicture };
    const c = img.obj.pic?.crop ?? { l: 0, t: 0, r: 0, b: 0 };
    const fw = box.w / Math.max(0.01, 1 - c.l - c.r), fh = box.h / Math.max(0.01, 1 - c.t - c.b);
    return { ...box, picture: true, crop: { x: box.x - c.l * fw, y: box.y - c.t * fh, w: fw, h: fh, url: imageUrl(img.obj.image) } };
  }

  /** Zuschneiden: the crop marks are shown until Esc, Enter or a click elsewhere. */
  private cropping = false;

  startCrop(): void {
    if (!this.selectedImage()?.obj.image) return;
    this.cropping = !this.cropping;
    this.update();
  }

  get isCropping(): boolean {
    return this.cropping;
  }

  endCrop(): void {
    if (!this.cropping) return;
    this.cropping = false;
    this.update();
  }

  /** Dragging a crop mark (or the picture inside the frame, `null`) while cropping. */
  private dragCrop(e: MouseEvent, handle: string | null): void {
    const img = this.selectedImage();
    const box0 = this.imageBox();
    if (!img || !box0) return;
    const f0 = img.obj.pic ?? {};
    const c0 = f0.crop ?? { l: 0, t: 0, r: 0, b: 0 };
    const w0 = (img.obj.width ?? 0), h0 = (img.obj.height ?? 0);
    const fw = w0 / Math.max(0.01, 1 - c0.l - c0.r), fh = h0 / Math.max(0.01, 1 - c0.t - c0.b);
    const start = { x: e.clientX, y: e.clientY };
    let first = true;
    const move = (ev: MouseEvent) => {
      const z = this.view.zoom;
      const dx = (ev.clientX - start.x) / z / PT, dy = (ev.clientY - start.y) / z / PT;
      const c = { ...c0 };
      if (handle) {
        if (handle.includes("w")) c.l = Math.min(c0.l + dx / fw, 1 - c.r - 0.02);
        if (handle.includes("e")) c.r = Math.min(c0.r - dx / fw, 1 - c.l - 0.02);
        if (handle.startsWith("n")) c.t = Math.min(c0.t + dy / fh, 1 - c.b - 0.02);
        if (handle.startsWith("s")) c.b = Math.min(c0.b - dy / fh, 1 - c.t - 0.02);
      } else {
        // The picture moves under the frame.
        c.l = c0.l - dx / fw; c.r = c0.r + dx / fw; c.t = c0.t - dy / fh; c.b = c0.b + dy / fh;
      }
      const w = fw * (1 - c.l - c.r), h = fh * (1 - c.t - c.b);
      if (first) this.snapshot();
      first = false;
      const pic = { ...f0, crop: c.l || c.t || c.r || c.b ? c : undefined };
      if (img.kind === "inline") this.replaceObject(img.p, img.o, { ...img.obj, width: w, height: h, pic });
      else this.replaceFloat(img.p, img.index, { ...img.obj, width: w, height: h, pic });
      this.update();
    };
    const up = () => { removeEventListener("mousemove", move, true); removeEventListener("mouseup", up, true); this.emit(true); };
    addEventListener("mousemove", move, true);
    addEventListener("mouseup", up, true);
  }

  /** The rotation handle: drag around the picture's centre (Umschalt: 15° steps). */
  private dragRotate(): void {
    const box = this.imageBox();
    if (!box) return;
    const ws = this.workspace.getBoundingClientRect();
    const z = this.view.zoom;
    const cx = ws.left + (this.view.pageLeft(box.page) + box.x + box.w / 2) * z - this.workspace.scrollLeft;
    const cy = ws.top + (this.view.pageTop(box.page) + box.y + box.h / 2) * z - this.workspace.scrollTop;
    let first = true;
    const move = (ev: MouseEvent) => {
      let a = (Math.atan2(ev.clientY - cy, ev.clientX - cx) * 180) / Math.PI + 90;
      if (ev.shiftKey) a = Math.round(a / 15) * 15;
      a = ((Math.round(a) % 360) + 360) % 360;
      this.setPictureFormat({ rotation: a || undefined }, first);
      first = false;
    };
    const up = () => { removeEventListener("mousemove", move, true); removeEventListener("mouseup", up, true); this.emit(true); };
    addEventListener("mousemove", move, true);
    addEventListener("mouseup", up, true);
  }

  /** The host shows Word's layout options (wrapping) next to the picture. */
  onLayoutOptions: ((rect: DOMRect) => void) | null = null;

  /** New picture data in place of the selected picture's (Freistellen, Bilder komprimieren). */
  replaceImageData(image: ImageData, snapshot = true): void {
    const img = this.selectedImage();
    if (!img) return;
    if (snapshot) this.snapshot();
    if (img.kind === "inline") this.replaceObject(img.p, img.o, { ...img.obj, image, data: undefined });
    else this.replaceFloat(img.p, img.index, { ...img.obj, image, data: undefined });
    this.update();
  }

  /**
   * Bilder komprimieren: pictures larger than `ppi` at their size on the page are scaled down;
   * cropped parts can go. Only the selected picture or all of the main text's pictures.
   */
  async compressPictures(ppi: number, deleteCropped: boolean, onlySelected: boolean): Promise<void> {
    const sel = this.selectedImage();
    const targets: { p: number; o?: number; index?: number }[] = [];
    if (onlySelected && sel) targets.push(sel.kind === "inline" ? { p: sel.p, o: sel.o } : { p: sel.p, index: sel.index });
    else {
      this.doc.main.forEach((para, p) => {
        for (let o = 0; o < para.text.length; o++) if (para.text[o] === OBJ && this.objectAt(para, o)?.kind === "image") targets.push({ p, o });
        (para.floats ?? []).forEach((f, index) => { if (f.image) targets.push({ p, index }); });
      });
    }
    let snap = true;
    for (const t of targets) {
      const para = this.doc.main[t.p];
      const obj = t.o !== undefined ? this.objectAt(para, t.o) : para.floats?.[t.index!];
      if (!obj?.image || !obj.width || !obj.height) continue;
      const crop = deleteCropped ? obj.pic?.crop : undefined;
      const out = await resample(obj.image, obj.width, obj.height, ppi, crop, obj.pic?.crop);
      if (!out) continue;
      if (snap) { this.snapshot(); snap = false; }
      const pic = crop && obj.pic ? { ...obj.pic, crop: undefined } : obj.pic;
      if (t.o !== undefined) this.replaceObject(t.p, t.o, { ...(obj as InlineObject), image: out, pic, data: undefined });
      else this.replaceFloat(t.p, t.index!, { ...(obj as FloatObject), image: out, pic, data: undefined });
    }
    this.update();
  }

  /** Position: the picture floats at one of nine places of the page margins, Quadrat wrapping. */
  positionPicture(x: "left" | "center" | "right", y: "top" | "center" | "bottom"): void {
    let img = this.selectedImage();
    if (!img || this.doc.activeStory !== MAIN) return;
    if (img.kind === "inline") { this.setImageWrap("square"); img = this.selectedImage(); }
    else this.snapshot();
    if (!img || img.kind !== "float") return;
    this.replaceFloat(img.p, img.index, { ...img.obj, x: { rel: "margin", align: x, offset: 0 }, y: { rel: "margin", align: y, offset: 0 }, wrap: img.obj.wrap === "inline" ? "square" : img.obj.wrap, data: img.obj.image ? undefined : img.obj.data });
    this.update({ scroll: true });
  }

  /** Ebene nach vorne / hinten, In den Vorder- / Hintergrund, Vor / Hinter den Text. */
  pictureLayer(dir: "forward" | "front" | "frontText" | "backward" | "back" | "behindText"): void {
    const img = this.selectedImage();
    if (!img || img.kind !== "float") return;
    if (dir === "frontText" || dir === "behindText") { this.setImageWrap(dir === "frontText" ? "front" : "behind"); return; }
    const zs: number[] = [];
    this.doc.main.forEach((para) => (para.floats ?? []).forEach((f, i) => zs.push(f.z ?? i + 1)));
    const cur = img.obj.z ?? img.index + 1;
    const others = zs.filter((z) => z !== cur).sort((a, b) => a - b);
    let z = cur;
    if (dir === "front") z = (others[others.length - 1] ?? cur) + 1;
    else if (dir === "back") z = Math.max(1, (others[0] ?? cur) - 1);
    else if (dir === "forward") { const n = others.find((v) => v > cur); z = n !== undefined ? n + 1 : cur; }
    else { const n = [...others].reverse().find((v) => v < cur); z = n !== undefined ? Math.max(1, n - 1) : cur; }
    this.snapshot();
    this.replaceFloat(img.p, img.index, { ...img.obj, z });
    this.update();
  }

  /**
   * Auswahlbereich: the main text's floating objects, front first, with their names (Word's
   * "Grafik 1", "Rechteck 2" … for objects without one) and whether they are hidden.
   */
  floatList(): { p: number; index: number; name: string; hidden: boolean; selected: boolean }[] {
    const all: { p: number; index: number; f: FloatObject }[] = [];
    this.doc.main.forEach((para, p) => (para.floats ?? []).forEach((f, index) => all.push({ p, index, f })));
    const sel = this.floatSel;
    let n = 0;
    const named = all.map((x) => {
      n++;
      const base = x.f.shape ? (x.f.shape.kind === "textbox" ? "Textfeld" : x.f.shape.kind === "connector" ? "Gerade Verbindung" : "Form") : "Grafik";
      return { p: x.p, index: x.index, z: x.f.z ?? x.index + 1, name: x.f.alt?.name ?? `${base} ${n}`, hidden: !!x.f.alt?.hidden, selected: !!sel && sel.p === x.p && sel.index === x.index };
    });
    return named.sort((a, b) => b.z - a.z).map(({ z: _z, ...rest }) => rest);
  }

  /** Word's name for a new drawing: the kind and the next number in the document ("Rechteck 3"). */
  private newObjectName(base: string): string {
    let n = 0;
    for (const p of this.doc.main) {
      n += p.floats?.length ?? 0;
      for (const o of p.objects ?? []) if (o.kind === "image") n++;
    }
    return `${base} ${n + 1}`;
  }

  /** Auswahlbereich: renames or hides/shows a floating object. */
  setFloatAlt(p: number, index: number, change: { name?: string; hidden?: boolean }): void {
    const f = this.doc.main[p]?.floats?.[index];
    if (!f) return;
    this.snapshot();
    const alt = Object.fromEntries(Object.entries({ ...f.alt, ...change }).filter(([, v]) => v !== undefined && v !== false && v !== "")) as AltText;
    this.replaceFloat(p, index, { ...f, alt: Object.keys(alt).length ? alt : undefined });
    if (change.hidden && this.floatSel?.p === p && this.floatSel.index === index) this.floatSel = null;
    this.update();
  }

  /** Ausrichten: a floating picture to the margin's (or page's) left, centre, right, top … */
  alignPicture(x: "left" | "center" | "right" | undefined, y: "top" | "center" | "bottom" | undefined, toPage: boolean): void {
    const img = this.selectedImage();
    if (!img || img.kind !== "float") return;
    this.snapshot();
    const rel = toPage ? "page" : "margin";
    this.replaceFloat(img.p, img.index, {
      ...img.obj, x: x ? { rel, align: x, offset: 0 } : img.obj.x, y: y ? { rel, align: y, offset: 0 } : img.obj.y,
    });
    this.update();
  }

  /** Mit Text verschieben, or Position auf Seite fixieren (the picture keeps its place). */
  setPictureMoveWithText(withText: boolean): void {
    const img = this.selectedImage();
    const box = this.imageBox();
    if (!img || img.kind !== "float" || !box) return;
    this.snapshot();
    const firstLine = this.layout.placed[this.layout.firstPlaced[img.p]];
    const y = withText ? { rel: "paragraph", offset: (box.y - (firstLine?.y ?? 0)) / PT } : { rel: "page", offset: box.y / PT };
    this.replaceFloat(img.p, img.index, { ...img.obj, y, moveWithText: withText ? undefined : false });
    this.update();
  }

  /** Seitenverhältnis: the largest frame of that ratio in the picture, centred; then Zuschneiden. */
  cropToRatio(ratio: number): void {
    const img = this.selectedImage();
    if (!img?.obj.image) return;
    const f = img.obj.pic ?? {};
    const c = f.crop ?? { l: 0, t: 0, r: 0, b: 0 };
    const fw = (img.obj.width ?? 0) / Math.max(0.01, 1 - c.l - c.r), fh = (img.obj.height ?? 0) / Math.max(0.01, 1 - c.t - c.b);
    let w = fw, h = fw / ratio;
    if (h > fh) { h = fh; w = fh * ratio; }
    const cl = (fw - w) / 2 / fw, ct = (fh - h) / 2 / fh;
    const crop = cl || ct ? { l: cl, t: ct, r: cl, b: ct } : undefined;
    this.snapshot();
    if (img.kind === "inline") this.replaceObject(img.p, img.o, { ...img.obj, width: w, height: h, pic: { ...f, crop } });
    else this.replaceFloat(img.p, img.index, { ...img.obj, width: w, height: h, pic: { ...f, crop } });
    this.cropping = true;
    this.update();
  }

  /** Ausfüllen (the picture covers the frame) or Einpassen (the whole picture in the frame). */
cropFill(fit: boolean): void {
    const img = this.selectedImage();
    if (!img?.obj.image) return;
    const info = imageInfo(img.obj.image.bytes);
    if (!info) return;
    const a = info.px.w / info.px.h;
    const F = (img.obj.width ?? 1) / (img.obj.height ?? 1);
    const side = (1 - F / a) / 2, topBottom = (1 - a / F) / 2;
    // Ausfüllen cuts the longer side, Einpassen pads the shorter one (negative crop).
    const crop = (a > F) !== fit ? { l: side, t: 0, r: side, b: 0 } : { l: 0, t: topBottom, r: 0, b: topBottom };
    this.setPictureFormat({ crop: Math.abs(side) < 1e-6 && Math.abs(topBottom) < 1e-6 ? undefined : crop });
  }

    /** Where the selected picture is drawn (page px), from the layout. */
  imageBox(): { page: number; x: number; y: number; w: number; h: number } | null {
    const img = this.selectedImage();
    if (!img) return null;
    if (img.kind === "float") {
      for (const pg of this.layout.pages) {
        for (const d of pg.deco) if (d.kind === "image" && d.story === "main" && d.para === img.p && d.index === img.index) return { page: pg.index, x: d.x, y: d.y, w: d.w, h: d.h };
      }
      return null;
    }
    const a = this.view.caretBox({ p: img.p, o: img.o }), b = this.view.caretBox({ p: img.p, o: img.o + 1 });
    const { pl } = this.view.placedFor({ p: img.p, o: img.o });
    const line = this.view.lineOf(pl);
    const w = (img.obj.width ?? 0) * PT, hh = (img.obj.height ?? 0) * PT;
    void b;
    const ex = img.obj.pic?.extent;
    return { page: a.page, x: a.x + (ex ? ex.l * PT : 0), y: pl.y + line.baseline - hh - (ex ? ex.b * PT : 0), w, h: hh };
  }

  /**
   * Textumbruch of the selected picture: "inline" (Mit Text in Zeile), "square" (Quadrat),
   * "tight", "through", "topAndBottom" (Oben und unten), "behind" (Hinter den Text),
   * "front" (Vor den Text). A picture leaving the line keeps its place on the page.
   */
  setImageWrap(wrap: string): void {
    const img = this.selectedImage();
    if (!img || this.doc.activeStory !== MAIN) return;
    const box = this.imageBox();
    this.snapshot();
    if (img.kind === "inline") {
      if (wrap === "inline") return;
      const { pl } = this.view.placedFor({ p: img.p, o: img.o });
      const fr = this.layout.frames[pl.frame];
      const paraTop = this.layout.placed[this.layout.firstPlaced[img.p]]?.y ?? pl.y;
      const f: FloatObject = {
        at: img.o, data: undefined, width: img.obj.width ?? 50, height: img.obj.height ?? 50, image: img.obj.image,
        x: { rel: "column", offset: box && fr ? (box.x - fr.x) / PT : 0 },
        y: { rel: "paragraph", offset: box ? (box.y - paraTop) / PT : 0 },
        wrap, side: "bothSides", dist: [0, 0, 9, 9], pic: img.obj.pic, alt: img.obj.alt,
      };
      this.doc.delete({ p: img.p, o: img.o }, { p: img.p, o: img.o + 1 });
      const para = this.doc.main[img.p];
      const next = [...this.doc.main];
      next[img.p] = { ...para, floats: [...(para.floats ?? []), f] };
      this.doc.setStory(MAIN, next);
      this.floatSel = { p: img.p, index: (para.floats ?? []).length };
      this.setSelKeepFloat({ p: img.p, o: img.o });
    } else if (wrap === "inline") {
      const f = img.obj;
      this.replaceFloat(img.p, img.index, null);
      const at = { p: img.p, o: Math.min(f.at, this.doc.main[img.p].text.length) };
      const obj: InlineObject = { kind: "image", data: undefined, display: "", width: f.width, height: f.height, image: f.image, pic: f.pic, alt: f.alt };
      E.insertObject(this.doc, at, at, obj, this.typingPropsAt(at));
      this.floatSel = null;
      this.setSel(at, { p: at.p, o: at.o + 1 });
    } else {
      this.replaceFloat(img.p, img.index, { ...img.obj, wrap, data: img.obj.image ? undefined : img.obj.data });
    }
    this.update({ scroll: true });
  }

  private typingPropsAt(at: Pos): CharProps {
    return typedProps(this.doc.paras[at.p], at.o);
  }

  private setSelKeepFloat(pos: Pos): void {
    const keep = this.floatSel;
    this.setSel(pos, pos);
    this.floatSel = keep;
  }

  /** Moves the selected floating picture by dx, dy points. */
  moveFloat(dx: number, dy: number, snapshot = true): void {
    const img = this.selectedImage();
    const box = this.imageBox();
    if (!img || img.kind !== "float" || !box) return;
    if (snapshot) this.snapshot();
    const f = img.obj;
    const pg = this.layout.pages[box.page];
    const firstLine = this.layout.placed[this.layout.firstPlaced[img.p]];
    const fr = firstLine ? this.layout.frames[firstLine.frame] : undefined;
    const baseX = f.x.rel === "page" ? 0 : f.x.rel === "column" && fr ? fr.x : pg.setup.marginLeft * PT;
    const baseY = f.y.rel === "page" ? 0 : f.y.rel === "margin" ? pg.setup.marginTop * PT : firstLine?.y ?? 0;
    const x = { rel: f.x.rel === "character" ? "column" : f.x.rel, offset: (box.x - baseX) / PT + dx };
    const y = { rel: f.y.rel === "line" ? "paragraph" : f.y.rel, offset: (box.y - baseY) / PT + dy };
    this.replaceFloat(img.p, img.index, { ...f, x, y, data: f.image ? undefined : f.data });
    this.update();
  }

  /** Copies the selected floating picture or shape (to paste it with Strg+V). */
  copyFloat(): boolean {
    const img = this.selectedImage();
    if (!img || img.kind !== "float") return false;
    const stories = new Map<string, Para[]>();
    const story = img.obj.shape?.story;
    if (story && this.doc.stories.has(story)) stories.set(story, this.doc.stories.get(story)!);
    clipboardFragment = { text: "", paras: [], source: this, float: { obj: img.obj, stories } };
    return true;
  }

  /**
   * Pastes a copied floating picture or shape at the same place, anchored to the paragraph
   * with the caret; a shape's text becomes a story of its own.
   */
  pasteFloat(): void {
    const frag = clipboardFragment?.float;
    if (!frag) return;
    if (this.inShapeText) this.leaveShapeText(false);
    this.snapshot();
    let obj: FloatObject = { ...frag.obj, data: frag.obj.image ? undefined : frag.obj.data };
    if (obj.shape) {
      // A new drawing (the old XML has the source's ids), with its own text.
      obj = { ...obj, data: undefined };
      const story = obj.shape!.story;
      if (story) {
        const id = this.newShapeStory();
        this.doc.setStory(id, frag.stories.get(story) ?? [{ ...makePara(""), mark: EMPTY }]);
        obj = { ...obj, shape: { ...obj.shape!, story: id } };
      }
    }
    let z = 251659264;
    for (const q of this.doc.main) for (const f of q.floats ?? []) z = Math.max(z, (f.z ?? 0) + 1024);
    obj = { ...obj, z };
    const p = this.floatSel ? this.floatSel.p : Math.min(this.sel.head.p, this.doc.main.length - 1);
    const para = this.doc.main[p];
    const floats = [...(para.floats ?? []), obj];
    const next = [...this.doc.main];
    next[p] = { ...para, floats };
    this.doc.setStory(MAIN, next);
    this.floatSel = { p, index: floats.length - 1 };
    this.update({ scroll: true });
    this.emit(true);
  }

  /** Deletes the selected floating picture. */
  deleteFloat(): boolean {
    const img = this.selectedImage();
    if (!img || img.kind !== "float") return false;
    this.snapshot();
    this.replaceFloat(img.p, img.index, null);
    if (img.obj.shape?.story) this.doc.deleteStory(img.obj.shape.story);
    this.floatSel = null;
    this.update();
    return true;
  }

  selectFloat(p: number, index: number): void {
    this.floatSel = { p, index };
    this.sel = { anchor: { p, o: 0 }, head: { p, o: 0 }, headUpstream: false };
    this.update();
  }


  // ----- shapes and text boxes ------------------------------------------------------------

  /** Arrow keys are moving a float (one undo step for a run of them). */
  private nudging = false;

  /** The shape whose text is being edited (its float in the main text). */
  private shapeEdit: { p: number; index: number } | null = null;

  /** The selected shape, or the one whose text is being edited. */
  selectedShape(): { p: number; index: number; obj: FloatObject; editing: boolean } | null {
    const img = this.selectedImage();
    if (img?.kind === "float" && img.obj.shape) return { p: img.p, index: img.index, obj: img.obj, editing: false };
    const se = this.shapeEdit;
    if (se && this.doc.activeStory !== MAIN) {
      const f = this.doc.main[se.p]?.floats?.[se.index];
      if (f?.shape && f.shape.story === this.doc.activeStory) return { p: se.p, index: se.index, obj: f, editing: true };
    }
    return null;
  }

  /** True while the caret is in the text of a shape or text box. */
  get inShapeText(): boolean {
    return this.selectedShape()?.editing ?? false;
  }

  /** Changes the selected shape (or the one being edited); one undo step. */
  setShape(change: Partial<ShapeInfo> | ((s: ShapeInfo) => ShapeInfo), snapshot = true): void {
    const s = this.selectedShape();
    if (!s) return;
    const old = s.obj.shape!;
    const next = typeof change === "function" ? change(old) : { ...old, ...change };
    if (snapshot) this.snapshot();
    this.replaceFloat(s.p, s.index, { ...s.obj, shape: next });
    this.update();
  }

  /** Changes the look (spPr: geometry, fill, outline, effects) of the selected shape. */
  setShapeFormat(change: PictureFormat | ((f: PictureFormat) => PictureFormat), snapshot = true): void {
    const s = this.selectedShape();
    if (!s) return;
    const old = s.obj.pic ?? {};
    const merged = typeof change === "function" ? change(old) : { ...old, ...change };
    const clean = Object.fromEntries(Object.entries(merged).filter(([, v]) => v !== undefined)) as PictureFormat;
    if (snapshot) this.snapshot();
    this.replaceFloat(s.p, s.index, { ...s.obj, pic: Object.keys(clean).length ? clean : undefined });
    this.update();
  }

  /** A new story id for a shape's text ("tx:…"). */
  private newShapeStory(): string {
    let n = 1;
    while (this.doc.stories.has(`tx:${n}`)) n++;
    return `tx:${n}`;
  }

  /**
   * Puts the caret into a shape's text (creating the text when the shape has none), at the
   * end or at a client point.
   */
  editShapeText(p: number, index: number, at?: { clientX: number; clientY: number } | "end", snapshot = true): void {
    const f = this.doc.main[p]?.floats?.[index];
    if (!f?.shape) return;
    let story = f.shape.story;
    if (!story || !this.doc.stories.has(story)) {
      if (snapshot) this.snapshot();
      story = this.newShapeStory();
      // Text typed into a shape is centred (Word); a text box starts like the body text.
      const props = f.shape.kind === "textbox" ? {} : { align: "center" as const };
      this.doc.setStory(story, [{ ...makePara("", props), mark: EMPTY }]);
      this.replaceFloat(p, index, { ...f, shape: { ...f.shape, story } });
    }
    this.floatSel = null;
    this.switchStory(story, story);
    this.shapeEdit = { p, index };
    if (at && at !== "end") {
      const hit = this.view.hitLine(at.clientX, at.clientY);
      const pos = this.view.posInLine(hit.index, hit.x);
      this.sel = { anchor: pos.pos, head: pos.pos, headUpstream: pos.upstream };
    } else {
      const end = this.doc.end();
      this.sel = { anchor: end, head: end, headUpstream: false };
    }
    this.update();
    this.emit(false);
  }

  /** Leaves a shape's text; with `select`, the shape stays selected (Esc in Word). */
  leaveShapeText(select: boolean): void {
    const se = this.shapeEdit;
    if (!se || this.doc.activeStory === MAIN) { this.shapeEdit = null; return; }
    this.switchStory(MAIN, MAIN);
    this.shapeEdit = null;
    if (select) this.floatSel = { p: se.p, index: se.index };
    this.update();
    this.emit(false);
  }

  /** Positions (page px) of a shape's yellow handles; rotation and flips applied. */
  private adjustHandles(f: FloatObject | InlineObject, box: { x: number; y: number; w: number; h: number }): { x: number; y: number }[] {
    if (!f.shape) return [];
    const g = shapeGeometry(f.pic, f.shape, box.w, box.h);
    return (g?.handles ?? []).map((hd) => this.fromShape(f.pic, box, hd.x, hd.y));
  }

  /** The two ends of a line or connector (page px), or null for other shapes. */
  private lineEnds(f: FloatObject | InlineObject, box: { x: number; y: number; w: number; h: number }): [{ x: number; y: number }, { x: number; y: number }] | null {
    const s = f.shape;
    const geom = f.pic?.geometry ?? "rect";
    if (!s || !(s.kind === "connector" || geom === "line" || geom === "straightConnector1" || geom.endsWith("Connector2") || geom.endsWith("Connector3") || geom.endsWith("Connector4") || geom.endsWith("Connector5"))) return null;
    return [this.fromShape(f.pic, box, 0, 0), this.fromShape(f.pic, box, box.w, box.h)];
  }

  /** Drags one end of a line: the other stays, the box and flips follow (as in Word). */
  private dragLineEnd(which: 0 | 1): void {
    const s = this.selectedShape();
    const box = s ? this.floatBox(s.p, s.index) : null;
    if (!s || !box) return;
    const ends = this.lineEnds(s.obj, box);
    if (!ends) return;
    const fixed = ends[1 - which];
    let first = true;
    const move = (ev: MouseEvent) => {
      const pt = this.view.pointAt(ev.clientX, ev.clientY);
      let mx = pt.x, my = pt.y;
      // Umschalt: multiples of 45°.
      if (ev.shiftKey) {
        const len = Math.hypot(mx - fixed.x, my - fixed.y), a = Math.round(Math.atan2(my - fixed.y, mx - fixed.x) / (Math.PI / 4)) * (Math.PI / 4);
        mx = fixed.x + len * Math.cos(a); my = fixed.y + len * Math.sin(a);
      }
      const start = which === 0 ? { x: mx, y: my } : fixed, end = which === 0 ? fixed : { x: mx, y: my };
      const cur = this.floatBox(s.p, s.index);
      if (!cur) return;
      const nx = Math.min(start.x, end.x), ny = Math.min(start.y, end.y);
      if (first) this.snapshot();
      this.moveFloat((nx - cur.x) / PT, (ny - cur.y) / PT, false);
      this.setImageSize(Math.abs(end.x - start.x) / PT, Math.abs(end.y - start.y) / PT, false);
      this.setShapeFormat((f) => ({ ...f, rotation: undefined, flipH: end.x < start.x ? true : undefined, flipV: end.y < start.y ? true : undefined }), false);
      first = false;
    };
    const up = () => {
      removeEventListener("mousemove", move, true);
      removeEventListener("mouseup", up, true);
      this.emit(true);
    };
    addEventListener("mousemove", move, true);
    addEventListener("mouseup", up, true);
  }

  /** A point in shape coordinates (px) to page px, with the shape's rotation and flips. */
  private fromShape(pic: PictureFormat | undefined, box: { x: number; y: number; w: number; h: number }, x: number, y: number): { x: number; y: number } {
    let dx = x - box.w / 2, dy = y - box.h / 2;
    if (pic?.flipH) dx = -dx;
    if (pic?.flipV) dy = -dy;
    const a = ((pic?.rotation ?? 0) * Math.PI) / 180;
    return { x: box.x + box.w / 2 + dx * Math.cos(a) - dy * Math.sin(a), y: box.y + box.h / 2 + dx * Math.sin(a) + dy * Math.cos(a) };
  }

  private toShape(pic: PictureFormat | undefined, box: { x: number; y: number; w: number; h: number }, x: number, y: number): { x: number; y: number } {
    const a = (-(pic?.rotation ?? 0) * Math.PI) / 180;
    const ox = x - box.x - box.w / 2, oy = y - box.y - box.h / 2;
    let dx = ox * Math.cos(a) - oy * Math.sin(a), dy = ox * Math.sin(a) + oy * Math.cos(a);
    if (pic?.flipH) dx = -dx;
    if (pic?.flipV) dy = -dy;
    return { x: dx + box.w / 2, y: dy + box.h / 2 };
  }

  /**
   * Drags a yellow handle: the adjust value whose handle lands nearest the mouse (searched
   * in the evaluated geometry, so every preset behaves as its formulas say).
   */
  private dragAdjust(i: number): void {
    const s = this.selectedShape();
    const box = s ? this.floatBox(s.p, s.index) : null;
    if (!s || !box) return;
    const f = s.obj;
    const shape = f.shape!;
    const base = { ...presetAdjust(f.pic?.geometry ?? "rect"), ...(shape.adjust ?? {}) };
    const g0 = shapeGeometry(f.pic, shape, box.w, box.h);
    const hd = g0?.handles[i];
    if (!hd) return;
    let first = true;
    const search = (adj: Record<string, number>, key: string, lo: number, hi: number, pos: (a: Record<string, number>) => number, target: number) => {
      // Coarse samples, then a finer look around the best one.
      let best = adj[key] ?? lo, bestD = Infinity;
      const steps = 48;
      let a = Math.min(lo, hi), b = Math.max(lo, hi);
      for (let round = 0; round < 3; round++) {
        for (let k = 0; k <= steps; k++) {
          const v = a + ((b - a) * k) / steps;
          const d = Math.abs(pos({ ...adj, [key]: v }) - target);
          if (d < bestD) { bestD = d; best = v; }
        }
        const span = (b - a) / steps;
        a = Math.max(Math.min(lo, hi), best - span);
        b = Math.min(Math.max(lo, hi), best + span);
      }
      return Math.round(best);
    };
    const move = (ev: MouseEvent) => {
      const pt = this.view.pointAt(ev.clientX, ev.clientY);
      const q = this.toShape(f.pic, box, pt.x, pt.y);
      const adj: Record<string, number> = { ...base };
      const at = (a: Record<string, number>) => presetGeometry(f.pic?.geometry ?? "rect", box.w, box.h, a)?.handles[i];
      if (hd.kind === "polar") {
        // gdX is the radius, gdY the angle (60000ths of a degree).
        if (hd.gdY) {
          let ang = (Math.atan2(q.y - box.h / 2, q.x - box.w / 2) * 180) / Math.PI;
          if (ang < 0) ang += 360;
          adj[hd.gdY] = Math.max(hd.minY ?? 0, Math.min(hd.maxY ?? 21599999, Math.round(ang * 60000)));
        }
        if (hd.gdX) {
          const r = Math.hypot(q.x - box.w / 2, q.y - box.h / 2);
          adj[hd.gdX] = search(adj, hd.gdX, hd.minX ?? 0, hd.maxX ?? 50000, (a) => { const h = at(a); return h ? Math.hypot(h.x - box.w / 2, h.y - box.h / 2) : 0; }, r);
        }
      } else {
        const big = (v: number | undefined, fallback: number) => (v === undefined || Math.abs(v) > 1e8 ? fallback : v);
        if (hd.gdX) adj[hd.gdX] = search(adj, hd.gdX, big(hd.minX, -400000), big(hd.maxX, 400000), (a) => at(a)?.x ?? 0, q.x);
        if (hd.gdY) adj[hd.gdY] = search(adj, hd.gdY, big(hd.minY, -400000), big(hd.maxY, 400000), (a) => at(a)?.y ?? 0, q.y);
      }
      this.setShape((sh) => ({ ...sh, adjust: adj }), first);
      first = false;
    };
    const up = () => {
      removeEventListener("mousemove", move, true);
      removeEventListener("mouseup", up, true);
      this.emit(true);
    };
    addEventListener("mousemove", move, true);
    addEventListener("mouseup", up, true);
  }

  /** The page box of a float (px), from the layout. */
  private floatBox(p: number, index: number): { page: number; x: number; y: number; w: number; h: number } | null {
    for (const pg of this.layout.pages) {
      for (const d of pg.deco) if (d.kind === "image" && d.story === "main" && d.para === p && d.index === index) return { page: pg.index, x: d.x, y: d.y, w: d.w, h: d.h };
    }
    return null;
  }

  /** True when a page point lies in a shape's text area (where a click edits its text). */
  private inShapeTextArea(p: number, index: number, page: number, x: number, y: number): boolean {
    const f = this.doc.main[p]?.floats?.[index];
    const box = this.floatBox(p, index);
    if (!f?.shape || !box || box.page !== page) return false;
    if (f.shape.kind === "connector" || (f.pic?.geometry ?? "rect") === "line") return false;
    const tr = shapeTextRect(f.pic, f.shape, box.w, box.h);
    const inRect = x >= box.x + tr.x && x <= box.x + tr.x + tr.w && y >= box.y + tr.y - 2 && y <= box.y + tr.y + tr.h + 2;
    if (!inRect) return false;
    // Text boxes take every click inside; other shapes only clicks on their text.
    if (f.shape.kind === "textbox") return true;
    const story = f.shape.story ? this.layout.stories.get(f.shape.story) : undefined;
    const paras = f.shape.story ? this.doc.stories.get(f.shape.story) : undefined;
    if (!story || !paras || !paras.some((q) => q.text.length)) return false;
    return story.placed.some((pl) => {
      const ln = story.paras[pl.para]?.lines[pl.line];
      return ln && pl.page === page && y >= pl.y && y <= pl.y + ln.height && x >= pl.x + ln.x0 - 4 && x <= pl.x + ln.x0 + ln.width + 4;
    });
  }

  /**
   * Einfügen > Formen / Textfeld: a new shape at a page rectangle (px), anchored to the
   * paragraph next to its top, in front of the text, with Word's default style.
   */
  insertShape(geometry: string, kind: ShapeInfo["kind"], page: number, x: number, y: number, w: number, h: number,
    extra?: { pic?: PictureFormat; adjust?: Record<string, number>; text?: string; custom?: string; line?: boolean; name?: string }): void {
    if (this.doc.activeStory !== MAIN) this.leaveShapeText(false);
    const pg = this.layout.pages[page];
    if (!pg) return;
    // The paragraph whose line is nearest above the shape's top on that page.
    let anchor = -1;
    for (const pl of this.layout.placed) {
      if (pl.page !== page || pl.ghost || pl.hide || this.doc.main[pl.para]?.cell) continue;
      if (pl.y <= y || anchor < 0) anchor = pl.para;
    }
    if (anchor < 0) anchor = Math.max(0, this.doc.main.length - 1);
    const firstLine = this.layout.placed[this.layout.firstPlaced[anchor]];
    const fr = firstLine ? this.layout.frames[firstLine.frame] : undefined;
    const colX = fr ? fr.x : pg.setup.marginLeft * PT;
    const paraTop = firstLine?.y ?? pg.setup.marginTop * PT;
    const theme = this.doc.styles.theme.colors ?? {};
    const accent = theme.accent1 ?? "4472C4";
    const line = kind === "connector" || geometry === "line" || !!extra?.line;
    const shape: ShapeInfo = {
      kind, adjust: extra?.adjust, custom: extra?.custom,
      ...(kind === "textbox" ? {} : line
        ? { styleLine: { color: accent, width: 0.5 }, fontColor: theme.dk1 ?? "000000" }
        : { styleFill: { kind: "solid", color: accent }, styleLine: { color: shadeLinear(accent, 0.15), width: 1 }, fontColor: theme.lt1 ?? "FFFFFF" }),
      body: kind === "textbox" ? { ...DEFAULT_SHAPE_BODY, anchor: "t" } : DEFAULT_SHAPE_BODY,
    };
    // Word's text box: white with a thin black outline.
    const pic: PictureFormat | undefined = kind === "textbox"
      ? { fill: { kind: "solid", color: theme.lt1 ?? "FFFFFF" }, fillXml: '<a:solidFill><a:schemeClr val="lt1"/></a:solidFill>', line: { color: "000000", width: 0.5 }, ...(extra?.pic ?? {}) }
      : geometry === "rect" && !extra?.pic ? undefined : { geometry: geometry === "rect" ? undefined : geometry, ...(extra?.pic ?? {}) };
    let z = 251659264;
    for (const q of this.doc.main) for (const f of q.floats ?? []) z = Math.max(z, (f.z ?? 0) + 1024);
    const f: FloatObject = {
      at: 0, data: undefined, width: w / PT, height: h / PT,
      x: { rel: "column", offset: (x - colX) / PT }, y: { rel: "paragraph", offset: (y - paraTop) / PT },
      wrap: "front", side: "bothSides", dist: [0, 0, 9, 9], pic: pic && Object.keys(pic).length ? pic : undefined, shape, z,
      alt: { name: this.newObjectName(extra?.name ?? (kind === "textbox" ? "Textfeld" : SHAPE_NAMES[geometry] ?? "Form")) },
    };
    this.snapshot();
    const para = this.doc.main[anchor];
    const floats = [...(para.floats ?? []), f];
    const next = [...this.doc.main];
    next[anchor] = { ...para, floats };
    this.doc.setStory(MAIN, next);
    this.floatSel = { p: anchor, index: floats.length - 1 };
    this.sel = { anchor: { p: anchor, o: 0 }, head: { p: anchor, o: 0 }, headUpstream: false };
    // A text box is ready for typing; other shapes stay selected.
    if (kind === "textbox" || extra?.text !== undefined) {
      this.editShapeText(anchor, floats.length - 1, "end", false);
      if (extra?.text) this.insertText(extra.text);
    } else this.update({ scroll: true });
    this.emit(true);
  }


  // ----- drawing shapes (Einfügen > Formen) ---------------------------------------------------

  /** The shape tool chosen in the Formen gallery: the next drag on a page draws it. */
  private drawTool: DrawTool | null = null;
  /** Called when a drawing ends (inserted or cancelled). */
  onDrawEnd: (() => void) | null = null;

  /** Einfügen > Formen: the next drag (or click) on a page draws the shape; Esc cancels. */
  startDrawing(tool: DrawTool): void {
    if (this.inShapeText) this.leaveShapeText(false);
    this.drawTool = tool;
    this.workspace.style.cursor = "crosshair";
    this.focus();
  }

  get drawing(): boolean {
    return this.drawTool !== null;
  }

  cancelDrawing(): void {
    if (!this.drawTool) return;
    this.drawTool = null;
    this.workspace.style.cursor = "";
    this.view.setDraft(null);
    this.onDrawEnd?.();
  }

  /** A drag on a page with a shape tool: the rubber band, then the new shape. */
  private drawShape(e: MouseEvent): void {
    const tool = this.drawTool!;
    if (tool.freeform) { this.drawFreeform(e, tool); return; }
    const start = this.view.pointAt(e.clientX, e.clientY);
    const line = tool.kind === "connector" || tool.geometry === "line";
    let end = { x: start.x, y: start.y };
    const rect = (shift: boolean) => {
      let dx = end.x - start.x, dy = end.y - start.y;
      if (shift && line) {
        // Lines snap to multiples of 45°.
        const len = Math.hypot(dx, dy), a = Math.round(Math.atan2(dy, dx) / (Math.PI / 4)) * (Math.PI / 4);
        dx = len * Math.cos(a); dy = len * Math.sin(a);
      } else if (shift) {
        const s = Math.max(Math.abs(dx), Math.abs(dy));
        dx = Math.sign(dx || 1) * s; dy = Math.sign(dy || 1) * s;
      }
      return { x: Math.min(start.x, start.x + dx), y: Math.min(start.y, start.y + dy), w: Math.abs(dx), h: Math.abs(dy), flipH: dx < 0, flipV: dy < 0 };
    };
    const preview = (shift: boolean) => {
      const r = rect(shift);
      const look: PictureFormat = { ...(tool.pic ?? {}), geometry: tool.geometry === "rect" ? undefined : tool.geometry, flipH: line && r.flipH ? true : undefined, flipV: line && r.flipV ? true : undefined };
      const info: ShapeInfo = { kind: tool.kind, adjust: tool.adjust, body: DEFAULT_SHAPE_BODY, styleLine: { color: "6F8DB9", width: 0.75 }, styleFill: line ? undefined : { kind: "solid", color: "DEE6F2", transparency: 0.4 } };
      const el = shapeElement(Math.max(r.w, 1), Math.max(r.h, 1), look, info);
      el.style.left = `${r.x}px`;
      el.style.top = `${r.y}px`;
      this.view.setDraft({ page: start.page, el });
    };
    const move = (ev: MouseEvent) => {
      const p = this.view.pointAt(ev.clientX, ev.clientY);
      end = { x: p.x, y: p.y };
      preview(ev.shiftKey);
    };
    const up = (ev: MouseEvent) => {
      removeEventListener("mousemove", move, true);
      removeEventListener("mouseup", up, true);
      this.view.setDraft(null);
      this.drawTool = null;
      this.workspace.style.cursor = "";
      let r = rect(ev.shiftKey);
      // A click without dragging: Word's default size (2.54 cm, lines 2.54 cm long).
      if (r.w < 3 && r.h < 3) r = { x: start.x, y: start.y, w: 72 * PT, h: line ? 0 : 72 * PT, flipH: false, flipV: false };
      const pic: PictureFormat = { ...(tool.pic ?? {}), flipH: line && r.flipH ? true : undefined, flipV: line && r.flipV ? true : undefined };
      this.insertShape(tool.geometry, tool.kind, start.page, r.x, r.y, r.w, r.h, { pic, adjust: tool.adjust });
      this.onDrawEnd?.();
    };
    addEventListener("mousemove", move, true);
    addEventListener("mouseup", up, true);
  }

  /**
   * Kurve (clicked points, smoothed), Freihandform: Form (clicked corners; dragging draws
   * freehand), Freihand: Skizze (one drag). A double click (or a click on the first point)
   * ends; a shape that ends where it began is closed and filled.
   */
  private drawFreeform(e: MouseEvent, tool: DrawTool): void {
    const page = this.view.pointAt(e.clientX, e.clientY).page;
    const pts: [number, number][] = [];
    const at = (ev: MouseEvent): [number, number] => { const p = this.view.pointAt(ev.clientX, ev.clientY); return [p.x, p.y]; };
    pts.push(at(e));
    let hover: [number, number] | null = null;
    let down = true;
    const preview = () => {
      const list = hover ? [...pts, hover] : pts;
      const d = freeformPath(list, tool.freeform === "curve");
      const svg = document.createElementNS("http://www.w3.org/2000/svg", "svg");
      svg.setAttribute("style", "position:absolute;left:0;top:0;width:1px;height:1px;overflow:visible");
      svg.innerHTML = `<path d="${d}" fill="none" stroke="#6F8DB9" stroke-width="1.5"/>`;
      const holder = document.createElement("div");
      holder.style.cssText = "position:absolute;left:0;top:0";
      holder.append(svg);
      this.view.setDraft({ page, el: holder });
    };
    const finish = () => {
      removeEventListener("mousemove", move, true);
      removeEventListener("mouseup", up, true);
      removeEventListener("mousedown", next, true);
      removeEventListener("dblclick", dbl, true);
      removeEventListener("keydown", key, true);
      this.view.setDraft(null);
      this.drawTool = null;
      this.workspace.style.cursor = "";
      const clean = pts.filter((p, i) => i === 0 || Math.hypot(p[0] - pts[i - 1][0], p[1] - pts[i - 1][1]) > 0.5);
      if (clean.length < 2) { this.onDrawEnd?.(); return; }
      const first = clean[0], last = clean[clean.length - 1];
      const closed = clean.length > 2 && Math.hypot(first[0] - last[0], first[1] - last[1]) < 6;
      if (closed) clean[clean.length - 1] = first;
      const xs = clean.map((p) => p[0]), ys = clean.map((p) => p[1]);
      const x = Math.min(...xs), y = Math.min(...ys), w = Math.max(1, Math.max(...xs) - x), h = Math.max(1, Math.max(...ys) - y);
      const local = clean.map(([px, py]) => [px - x, py - y] as [number, number]);
      this.insertFreeform(page, x, y, w, h, local, tool.freeform === "curve", closed);
      this.onDrawEnd?.();
    };
    const move = (ev: MouseEvent) => {
      const p = at(ev);
      // Dragging adds points (Skizze always, Form while the button is down).
      if (down && tool.freeform !== "curve") {
        const lastP = pts[pts.length - 1];
        if (Math.hypot(p[0] - lastP[0], p[1] - lastP[1]) > 2) pts.push(p);
        hover = null;
      } else hover = p;
      preview();
    };
    const up = () => {
      down = false;
      if (tool.freeform === "scribble") finish();
    };
    const next = (ev: MouseEvent) => {
      if (ev.button !== 0) return;
      ev.preventDefault();
      ev.stopPropagation();
      const p = at(ev);
      // A click on the first point closes the shape.
      if (pts.length > 2 && Math.hypot(p[0] - pts[0][0], p[1] - pts[0][1]) < 6) { pts.push(pts[0]); finish(); return; }
      pts.push(p);
      down = true;
    };
    const dbl = (ev: MouseEvent) => { ev.preventDefault(); ev.stopPropagation(); finish(); };
    const key = (ev: KeyboardEvent) => {
      if (ev.key === "Escape" || ev.key === "Enter") { ev.preventDefault(); ev.stopPropagation(); finish(); }
    };
    addEventListener("mousemove", move, true);
    addEventListener("mouseup", up, true);
    if (tool.freeform !== "scribble") {
      setTimeout(() => addEventListener("mousedown", next, true), 0);
      addEventListener("dblclick", dbl, true);
      addEventListener("keydown", key, true);
    }
  }

  /** A freeform shape (custGeom) from points in its own box (px). */
  private insertFreeform(page: number, x: number, y: number, w: number, h: number, pts: [number, number][], smooth: boolean, closed: boolean): void {
    // In EMU-like units of the box, as Word writes custom geometry (path w/h = the shape's size).
    const W = Math.round((w / PT) * 12700), H = Math.round((h / PT) * 12700);
    const sx = W / w, sy = H / h;
    const q = pts.map(([px, py]) => [Math.round(px * sx), Math.round(py * sy)] as [number, number]);
    let cmds = `M ${q[0][0]} ${q[0][1]}`;
    if (smooth) {
      for (let i = 0; i < q.length - 1; i++) {
        const p0 = q[Math.max(0, i - 1)], p1 = q[i], p2 = q[i + 1], p3 = q[Math.min(q.length - 1, i + 2)];
        const c1 = [Math.round(p1[0] + (p2[0] - p0[0]) / 6), Math.round(p1[1] + (p2[1] - p0[1]) / 6)];
        const c2 = [Math.round(p2[0] - (p3[0] - p1[0]) / 6), Math.round(p2[1] - (p3[1] - p1[1]) / 6)];
        cmds += ` C ${c1[0]} ${c1[1]} ${c2[0]} ${c2[1]} ${p2[0]} ${p2[1]}`;
      }
    } else for (const p of q.slice(1)) cmds += ` L ${p[0]} ${p[1]}`;
    if (closed) cmds += " Z";
    const custom = `rect l t r b\npath w=${W} h=${H}${closed ? "" : " fill=none"}: ${cmds}`;
    this.insertShape("rect", "shape", page, x, y, w, h, { custom, line: !closed });
  }

  // ----- find and replace ------------------------------------------------------------------

  private findMarks: WeakMap<Para, TextMark[]> | null = null;
  /** Extra marks from the host (spelling errors), by story and paragraph. */
  extraMarks: ((story: string, para: Para) => readonly TextMark[] | undefined) | null = null;
  private mergedMarks = new WeakMap<Para, { f?: readonly TextMark[]; x?: readonly TextMark[]; out: readonly TextMark[] }>();

  private commentMarks: { main: Para[]; active: string | null; map: WeakMap<Para, TextMark[]> } | null = null;

  /** Commented text (light), the comment at the caret (darker), per paragraph. */
  private commentMarksFor(para: Para): TextMark[] | undefined {
    if (!this.doc.comments.size) return undefined;
    const active = this.focusedComment ?? this.commentAtCaret();
    if (!this.commentMarks || this.commentMarks.main !== this.doc.main || this.commentMarks.active !== active) {
      const map = new WeakMap<Para, TextMark[]>();
      for (const r of S.commentRanges(this.doc)) {
        if (this.doc.comments.get(r.id)?.parent) continue;
        for (let p = r.start.p; p <= r.end.p; p++) {
          const para2 = this.doc.main[p];
          const a = p === r.start.p ? r.start.o : 0, b = p === r.end.p ? r.end.o : para2.text.length;
          if (b <= a) continue;
          const list = map.get(para2) ?? [];
          list.push({ start: a, end: b, kind: r.id === active ? "commentActive" : "comment" });
          map.set(para2, list);
        }
      }
      this.commentMarks = { main: this.doc.main, active, map };
    }
    return this.commentMarks.map.get(para);
  }

  private installMarks(): void {
    this.view.marksFor = (story, para) => {
      const f0 = story === this.doc.activeStory || (story === "main" && this.doc.activeStory === MAIN) ? this.findMarks?.get(para) : undefined;
      const cm = story === "main" ? this.commentMarksFor(para) : undefined;
      const f = cm && f0 ? [...cm, ...f0] : cm ?? f0;
      const x = this.extraMarks?.(story, para);
      if (!f) return x;
      if (!x) return f;
      const hit = this.mergedMarks.get(para);
      if (hit && hit.f === f && hit.x === x) return hit.out;
      const out = [...x, ...f];
      this.mergedMarks.set(para, { f, x, out });
      return out;
    };
  }

  /** All matches of `query` in the story with the caret (paragraph, start, end). */
  findAll(query: string, opts: { matchCase?: boolean; wholeWord?: boolean } = {}): { p: number; a: number; b: number }[] {
    const out: { p: number; a: number; b: number }[] = [];
    if (!query) return out;
    const q = opts.matchCase ? query : query.toLocaleLowerCase("de-DE");
    this.doc.paras.forEach((para, p) => {
      if (para.block) return;
      const text = opts.matchCase ? para.text : para.text.toLocaleLowerCase("de-DE");
      if (text.length !== para.text.length) return;
      for (let i = text.indexOf(q); i >= 0; i = text.indexOf(q, i + Math.max(1, q.length))) {
        if (opts.wholeWord && (isWordChar(para.text[i - 1]) || isWordChar(para.text[i + q.length]))) continue;
        out.push({ p, a: i, b: i + q.length });
      }
    });
    return out;
  }

  /** Highlights matches (Navigationsbereich: all in yellow, the current one darker). */
  showFindMarks(matches: { p: number; a: number; b: number }[] | null, current?: { p: number; a: number }): void {
    if (!matches) { this.findMarks = null; this.view.refresh(); return; }
    const map = new WeakMap<Para, TextMark[]>();
    const byPara = new Map<number, TextMark[]>();
    for (const m of matches) {
      const list = byPara.get(m.p) ?? [];
      list.push({ start: m.a, end: m.b, kind: current && current.p === m.p && current.a === m.a ? "findCurrent" : "find" });
      byPara.set(m.p, list);
    }
    for (const [p, list] of byPara) map.set(this.doc.paras[p], list);
    this.findMarks = map;
    this.view.refresh();
  }

  /** Selects a match and scrolls to it. */
  selectRange(p: number, a: number, b: number): void {
    if (this.doc.activeStory !== MAIN) this.switchStory(MAIN, MAIN);
    this.setSel({ p, o: a }, { p, o: b });
    this.update({ scroll: true });
  }

  /** Replaces [a, b) of paragraph p with text in the formatting of its first character. */
  replaceRange(p: number, a: number, b: number, text: string): void {
    this.snapshot();
    const props = this.doc.charPropsAt({ p, o: a });
    const end = this.doc.insertText({ p, o: a }, { p, o: b }, text, props);
    this.setSel(end, end);
    this.update({ scroll: true });
  }

  /** Alle ersetzen: one undo step; returns the number of replacements. */
  replaceAll(query: string, text: string, opts: { matchCase?: boolean; wholeWord?: boolean } = {}): number {
    const matches = this.findAll(query, opts);
    if (!matches.length) return 0;
    this.snapshot();
    // From the end, so earlier offsets stay valid.
    for (let i = matches.length - 1; i >= 0; i--) {
      const m = matches[i];
      const props = this.doc.charPropsAt({ p: m.p, o: m.a });
      this.doc.insertText({ p: m.p, o: m.a }, { p: m.p, o: m.b }, text, props);
    }
    const h = this.sel.head;
    this.setSel(this.fixPos(h), this.fixPos(h));
    this.update({ scroll: true });
    return matches.length;
  }

  private fixPos(p: Pos): Pos {
    const pi = Math.min(p.p, this.doc.paras.length - 1);
    return { p: pi, o: Math.min(p.o, this.doc.paras[pi].text.length) };
  }

  /** Headings of the main text (Navigationsbereich > Überschriften). */
  headings(): { p: number; level: number; text: string; label?: string }[] {
    const out: { p: number; level: number; text: string; label?: string }[] = [];
    this.doc.main.forEach((para, p) => {
      const rp = this.doc.styles.resolvePara(para.props);
      const lvl = para.props.outline ?? this.doc.styles.outlineLevel(rp.style);
      if (lvl === undefined || lvl >= 9 || !para.text.trim()) return;
      const label = this.layout?.paras[p]?.lines[0]?.label?.text.replace(/\.$/, "");
      out.push({ p, level: lvl + 1, text: para.text.replace(/[\u000b\ufffc]/g, " ").trim(), label: label || undefined });
    });
    return out;
  }

  /** Forgets undo history and marks the content as the saved state (new document from a template). */
  clearHistory(): void {
    this.undoStack = [];
    this.redoStack = [];
  }

  /** Puts the caret at the start of a main-text paragraph and scrolls there. */
  goToParagraph(p: number): void {
    if (this.doc.activeStory !== MAIN) this.switchStory(MAIN, MAIN);
    this.setSel({ p, o: 0 }, { p, o: 0 });
    this.update({ scroll: true });
  }

  /** Gehe zu Seite n (1-based). */
  goToPage(n: number): void {
    const page = Math.max(0, Math.min(this.layout.pageCount - 1, n - 1));
    const i = this.layout.placed.findIndex((pl) => pl.page === page && !pl.ghost && !pl.hide);
    if (i < 0) return;
    const pl = this.layout.placed[i];
    if (this.doc.activeStory !== MAIN) this.switchStory(MAIN, MAIN);
    const line = this.layout.paras[pl.para].lines[pl.line];
    this.setSel({ p: pl.para, o: line.start }, { p: pl.para, o: line.start });
    this.view.scrollToPage(page);
    this.update();
  }

  // ----- foot- and endnotes ------------------------------------------------------------------

  /** Fußnote einfügen (Strg+Alt+F) / Endnote einfügen (Strg+Alt+D): the caret goes into the note. */
  /** Fuß- und Endnote: the document's note settings (Word's defaults when it has none). */
  noteSettings(): NonNullable<DocSettings["notes"]> {
    return this.doc.settings.notes ?? {};
  }

  setNoteSettings(notes: NonNullable<DocSettings["notes"]>): void {
    this.snapshot();
    this.doc.setSettings({ ...this.doc.settings, notes });
    this.relayout();
  }

  /**
   * Fuß- und Endnote > Konvertieren: all footnotes become endnotes, all endnotes footnotes, or
   * both swap. References, note texts and their styles follow.
   */
  convertNotes(mode: "toEndnotes" | "toFootnotes" | "swap"): void {
    if (this.storyKey !== MAIN) return;
    const map = new Map<string, string>();
    const next = { fn: 0, en: 0 };
    for (const k of this.doc.stories.keys()) {
      const kind = k.slice(0, 2) as "fn" | "en";
      if (kind === "fn" || kind === "en") next[kind] = Math.max(next[kind], Number(k.slice(3)) || 0);
    }
    for (const k of S.referencedNotes(this.doc)) {
      const kind = k.slice(0, 2);
      const to = mode === "swap" ? (kind === "fn" ? "en" : "fn") : mode === "toEndnotes" ? "en" : "fn";
      if (kind === to) continue;
      map.set(k, `${to}:${++next[to as "fn" | "en"]}`);
    }
    if (!map.size) return;
    this.snapshot();
    const style = (kind: string, what: "reference" | "text") => {
      const id = this.styleId(`${kind === "fn" ? "footnote" : "endnote"} ${what}`);
      this.ensureStyle(id);
      return id;
    };
    const refOf = new Map([["fn", style("fn", "reference")], ["en", style("en", "reference")]]);
    const textOf = new Map([["fn", style("fn", "text")], ["en", style("en", "text")]]);
    const swapRuns = (p: Para): Para => {
      if (!p.objects?.some((o) => o.note && map.has(o.note))) return p;
      let k = 0;
      const runs: Para["runs"][number][] = [];
      let pos = 0;
      for (const r of p.runs) {
        // A run over a converted reference gets the other reference style.
        let props = r.props;
        for (let i = pos; i < pos + r.len; i++) {
          if (p.text[i] !== OBJ) continue;
          const o = p.objects?.[k++];
          const to = o?.note ? map.get(o.note) : undefined;
          if (to && props.rStyle === refOf.get(o!.note!.slice(0, 2))) props = intern({ ...props, rStyle: refOf.get(to.slice(0, 2)) });
        }
        runs.push(props === r.props ? r : { ...r, props });
        pos += r.len;
      }
      const objects = p.objects.map((o) => {
        const to = o.note ? map.get(o.note) : undefined;
        if (!to) return o;
        const data = o.data && typeof o.data === "object" && "inner" in (o.data as object) ? { inner: to.startsWith("en") ? "<w:endnoteRef/>" : "<w:footnoteRef/>" } : undefined;
        return { ...o, note: to, data };
      });
      return { ...p, runs, objects };
    };
    this.doc.setStory(MAIN, this.doc.main.map(swapRuns));
    for (const [from, to] of map) {
      const paras = this.doc.stories.get(from);
      if (!paras) continue;
      const fromText = textOf.get(from.slice(0, 2)), toText = textOf.get(to.slice(0, 2));
      this.doc.setStory(to, paras.map((q) => swapRuns({ ...q, props: q.props.style === fromText ? { ...q.props, style: toText } : q.props })));
      this.doc.deleteStory(from);
    }
    this.update();
  }

  insertNote(kind: "fn" | "en"): void {
    if (this.doc.activeStory !== MAIN) return;
    this.snapshot();
    const refStyle = this.styleId(kind === "fn" ? "footnote reference" : "endnote reference");
    const textStyle = this.styleId(kind === "fn" ? "footnote text" : "endnote text");
    this.ensureStyle(refStyle);
    this.ensureStyle(textStyle);
    const [from, to] = this.range();
    const at = cmp(from, to) !== 0 ? to : from;
    const { id } = S.insertNote(this.doc, at, kind, refStyle, textStyle);
    this.requireFonts();
    this.update();
    this.switchStory(id, id, { anchor: { p: 0, o: 2 }, head: { p: 0, o: 2 }, headUpstream: false });
    this.update({ scroll: true });
  }

  /** Nächste/Vorherige Fuß- oder Endnote: to the next reference in the text. */
  /** The note references of the main text, in order. */
  private noteRefs(kind?: "fn" | "en"): { at: Pos; id: string }[] {
    const out: { at: Pos; id: string }[] = [];
    this.doc.main.forEach((p, i) => {
      if (!p.objects?.some((x) => x.note)) return;
      let k = 0;
      for (let o = p.text.indexOf(OBJ); o >= 0; o = p.text.indexOf(OBJ, o + 1)) {
        const id = p.objects[k++]?.note;
        if (id && (!kind || id.startsWith(`${kind}:`)) && this.doc.stories.has(id)) out.push({ at: { p: i, o }, id });
      }
    });
    return out;
  }

  /** Which kinds of notes the document has (Notizen anzeigen). */
  noteKinds(): { fn: boolean; en: boolean } {
    const refs = this.noteRefs();
    return { fn: refs.some((r) => r.id.startsWith("fn:")), en: refs.some((r) => r.id.startsWith("en:")) };
  }

  /**
   * Referenzen > Notizen anzeigen: from a note's text back to its reference; from the text to
   * the notes area (the note at or after the caret, else the last one before it).
   */
  showNotes(kind: "fn" | "en" = "fn"): boolean {
    const story = this.doc.activeStory;
    if (story.startsWith("fn:") || story.startsWith("en:")) {
      const ref = this.noteRefs().find((r) => r.id === story);
      this.switchStory(MAIN, MAIN);
      if (ref) this.selectRange(ref.at.p, ref.at.o + 1, ref.at.o + 1);
      return true;
    }
    if (story !== MAIN) this.switchStory(MAIN, MAIN);
    const refs = this.noteRefs(kind);
    if (!refs.length) return false;
    const h = this.sel.head;
    const ref = refs.find((r) => cmp(r.at, h) >= 0) ?? refs[refs.length - 1];
    const paras = this.doc.stories.get(ref.id)!;
    const first = paras[0]?.text ?? "";
    // After the note's own number and its space, as Word puts the caret.
    const o = first[0] === OBJ ? (first[1] === " " ? 2 : 1) : 0;
    this.switchStory(ref.id, ref.id, { anchor: { p: 0, o }, head: { p: 0, o }, headUpstream: false });
    this.update({ scroll: true });
    return true;
  }

  nextNote(dir: 1 | -1, kind: "fn" | "en" = "fn"): void {
    if (this.doc.activeStory !== MAIN) this.switchStory(MAIN, MAIN);
    const h = this.sel.head;
    const refs: Pos[] = [];
    this.doc.main.forEach((p, i) => {
      let k = 0;
      for (let o = 0; o < p.text.length; o++) {
        if (p.text[o] !== OBJ) continue;
        if (p.objects?.[k++]?.note?.startsWith(`${kind}:`)) refs.push({ p: i, o });
      }
    });
    const next = dir > 0 ? refs.find((r) => cmp(r, h) > 0) : [...refs].reverse().find((r) => cmp(r, h) < 0);
    if (next) this.selectRange(next.p, next.o, next.o);
  }

  // ----- table of contents ----------------------------------------------------------------------

  private tocWidth(): number {
    const s = this.section().page;
    return s.width - s.marginLeft - s.marginRight - s.gutter;
  }

  private pageNumberOf = (p: number): string => {
    const i = this.layout.firstPlaced[p];
    const pl = i !== undefined && i >= 0 ? this.layout.placed[i] : undefined;
    return pl ? this.layout.pages[pl.page].numberText : "1";
  };

  /**
   * Inhaltsverzeichnis einfügen ("Automatische Tabelle 1/2"): a heading and one linked entry
   * per Überschrift 1-3, page numbers from the layout (laid out twice, so the numbers count
   * the table itself).
   */
  insertToc(title = "Inhalt"): void {
    if (this.doc.activeStory !== MAIN || this.inTable()) return;
    this.snapshot();
    const styles = [1, 2, 3].map((k) => { const id = this.styleId(`toc ${k}`); this.ensureStyle(id); return id; });
    const head = this.styleId("TOC Heading");
    this.ensureStyle(head);
    const h = this.sel.head;
    // At the start of a new paragraph before the caret's.
    const at = { p: h.p, o: 0 };
    const entries = this.headings().filter((x) => x.level <= 3);
    // numId 0: the heading keeps no number when the headings are numbered.
    const heading: Para = makePara(title, { style: head, numId: 0 });
    const paras = S.tocParagraphs(this.doc, entries, () => "1", styles, this.tocWidth());
    this.doc.splice(at.p, at.p - 1, [heading, ...paras]);
    this.update();
    this.refreshToc(S.tocFields(this.doc).find((f) => f.a === at.p + 1)?.id);
    this.setSel({ p: at.p, o: 0 }, { p: at.p, o: 0 });
    this.update({ scroll: true });
  }

  /**
   * Abbildungsverzeichnis einfügen: a TOC field with \\c "Abbildung" (or Formel, Tabelle)
   * listing the captions with dotted page numbers, at the caret's paragraph.
   */
  insertTableOfFigures(label = "Abbildung", look: S.TocLook = {}, includeLabel = true): void {
    if (this.doc.activeStory !== MAIN || this.inTable()) return;
    this.snapshot();
    const style = this.styleId("table of figures");
    this.ensureStyle(style);
    const cur = this.sel.head.p;
    const empty = !this.doc.main[cur]?.text.length;
    const at = empty ? cur : cur + 1;
    // Word's switches: \\h links, \\n no page numbers, \\c with "Abbildung 1", \\a the caption text only.
    const instr = `TOC ${look.links !== false ? "\\h " : ""}\\z ${look.pageNumbers === false ? "\\n " : ""}${includeLabel ? "\\c" : "\\a"} "${label}"`;
    const entries = Refs.captions(this.doc, label).map((x) => ({ p: x.p, level: 1, text: includeLabel ? x.text : x.rest }));
    const paras = S.tocParagraphs(this.doc, entries, () => "1", [style], this.tocWidth(), undefined, instr, look);
    this.doc.splice(at, empty ? at : at - 1, paras);
    this.update();
    this.refreshToc(S.tocFields(this.doc).find((f) => f.a === at)?.id);
    this.setSel({ p: at, o: 0 }, { p: at, o: 0 });
    this.update({ scroll: true });
  }

  /** Abbildungsverzeichnis aktualisieren: every table of figures is built again. */
  updateTablesOfFigures(): void {
    const ids = S.tocFields(this.doc).filter((f) => /\\[ca]\s+"/i.test(f.instr)).map((f) => f.id);
    if (!ids.length) return;
    this.snapshot();
    for (const id of ids) this.refreshToc(id);
    this.update({ scroll: true });
  }

  /** Whether the document has a table of figures (Abbildungsverzeichnis aktualisieren). */
  hasTableOfFigures(): boolean {
    return S.tocFields(this.doc).some((f) => /\\[ca]\s+"/i.test(f.instr));
  }

  /** Replaces paragraph `p` by a table of contents (the heading above stays as it is). */
  insertTocAt(p: number): void {
    if (p < 0 || p >= this.doc.main.length) return;
    const styles = [1, 2, 3].map((k) => { const id = this.styleId(`toc ${k}`); this.ensureStyle(id); return id; });
    this.snapshot();
    const entries = this.headings().filter((x) => x.level <= 3 && x.p !== p);
    this.doc.splice(p, p, S.tocParagraphs(this.doc, entries, () => "1", styles, this.tocWidth()));
    this.update();
    this.refreshToc(S.tocFields(this.doc).find((f) => f.a === p)?.id);
    this.update();
  }

  hasToc(): boolean {
    return S.tocFields(this.doc).length > 0;
  }

  /**
   * Tabelle aktualisieren: "Nur Seitenzahlen aktualisieren" keeps the entries and corrects
   * their numbers; otherwise the whole table is rebuilt from the headings.
   */
  updateToc(pagesOnly = false): void {
    if (!S.tocFields(this.doc).length) return;
    this.snapshot();
    if (pagesOnly) this.refreshTocPages(); else this.refreshToc();
    this.update({ scroll: true });
  }

  private refreshTocPages(): void {
    const where = new Map<string, number>();
    this.doc.main.forEach((p, i) => { for (const a of p.anchors ?? []) if (a.bookmark) where.set(a.bookmark, i); });
    const main = [...this.doc.main];
    let changed = false;
    for (const f of S.tocFields(this.doc)) {
      for (let i = f.a; i <= f.b; i++) {
        const p = main[i];
        if (!p.objects?.length) continue;
        const objects = p.objects.map((o) => {
          const m = o.kind === "field" ? /^\s*PAGEREF\s+(\S+)/i.exec(o.instr ?? "") : null;
          const at = m ? where.get(m[1]) : undefined;
          if (at === undefined) return o;
          const display = this.pageNumberOf(at);
          return display === o.display ? o : { ...o, display, data: undefined };
        });
        if (objects.some((o, k) => o !== p.objects![k])) { main[i] = { ...p, objects }; changed = true; }
      }
    }
    if (changed) this.doc.setStory(MAIN, main);
  }

  /**
   * Rebuilds tables of contents until their page numbers are stable (a longer table can push
   * headings to later pages).
   */
  private refreshToc(only?: string): void {
    const styles = [1, 2, 3].map((k) => { const id = this.styleId(`toc ${k}`); this.ensureStyle(id); return id; });
    const sig = (ps: readonly Para[]) => ps.map((p) => p.text + (p.objects?.map((o) => o.display).join() ?? "")).join("|");
    const ids = S.tocFields(this.doc).map((f) => f.id).filter((id) => !only || id === only);
    for (let pass = 0; pass < 4; pass++) {
      let changed = false;
      for (const id of ids) {
        const f = S.tocFields(this.doc).find((x) => x.id === id);
        if (!f) continue;
        // An Abbildungsverzeichnis (TOC \\c "Abbildung") lists the captions of its label.
        const label = /\\[ca]\s+"([^"]+)"/i.exec(f.instr)?.[1];
        const withLabel = !/\\a\s+"/i.test(f.instr);
        const outside = (p: number) => p < f.a || p > f.b;
        const entries = label
          ? Refs.captions(this.doc, label).filter((x) => outside(x.p)).map((x) => ({ p: x.p, level: 1, text: withLabel ? x.text : x.rest }))
          : this.headings().filter((x) => x.level <= 3 && outside(x.p));
        // The entries keep their look: page numbers, the right tab and its leader, links.
        const first = this.doc.main[f.a];
        const tab = first?.props.tabs?.find((t) => t.align === "right");
        const look: S.TocLook = label ? {
          pageNumbers: !/\\n\b/.test(f.instr), links: /\\h\b/.test(f.instr),
          rightAlign: !!tab || !first?.objects?.length, leader: (tab?.leader as S.TocLook["leader"]) ?? "dot",
        } : {};
        const tof = this.styleId("table of figures");
        if (label) this.ensureStyle(tof);
        const before = sig(this.doc.main.slice(f.a, f.b + 1));
        const paras = S.tocParagraphs(this.doc, entries, this.pageNumberOf, label ? [tof] : styles, this.tocWidth(), f.id, f.instr, look);
        if (sig(paras) === before) continue;
        changed = true;
        S.stripTocTail(this.doc, f);
        this.doc.splice(f.a, f.b, paras);
      }
      if (!changed) break;
      this.update();
    }
  }

  removeToc(): void {
    const fields = S.tocFields(this.doc);
    if (!fields.length) return;
    this.snapshot();
    for (const f of [...fields].reverse()) {
      // The heading Word puts above (Inhaltsverzeichnisüberschrift) goes too.
      const headStyle = this.styleId("TOC Heading");
      const a = f.a > 0 && this.doc.main[f.a - 1].props.style === headStyle ? f.a - 1 : f.a;
      S.stripTocTail(this.doc, f);
      const main = [...this.doc.main];
      main.splice(a, f.b - a + 1);
      this.doc.setStory(MAIN, main.length ? main : [makePara("")]);
    }
    this.clampSelection();
    this.update({ scroll: true });
  }

  // ----- comments ---------------------------------------------------------------------------------

  /** Neuer Kommentar (Strg+Alt+K) on the selection or the word at the caret; returns its id. */
  addComment(text = ""): string | null {
    if (this.doc.activeStory !== MAIN) return null;
    let [from, to] = this.range();
    if (cmp(from, to) === 0) {
      const w = this.wordAt(from);
      if (w) { from = { p: w.p, o: w.a }; to = { p: w.p, o: w.b }; }
    }
    this.snapshot();
    const textStyle = this.styleId("annotation text"), refStyle = this.styleId("annotation reference");
    this.ensureStyle(textStyle);
    this.ensureStyle(refStyle);
    const id = S.addComment(this.doc, from, to, { author: this.author, initials: this.authorInitials, date: new Date().toISOString().replace(/\.\d+Z$/, "Z") }, textStyle, refStyle, text);
    this.update();
    return id;
  }

  /** Every comment with its text and where it is (for the comment pane). */
  comments(): { id: string; author: string; initials: string; date: string; text: string; parent?: string; done?: boolean; start: Pos; end: Pos }[] {
    const ranges = S.commentRanges(this.doc);
    return ranges.map((r) => {
      const m = this.doc.comments.get(r.id);
      return { id: r.id, author: m?.author ?? "", initials: m?.initials ?? "", date: m?.date ?? "", text: S.commentText(this.doc, r.id), parent: m?.parent, done: m?.done, start: r.start, end: r.end };
    });
  }

  setCommentText(id: string, text: string): void {
    if (S.commentText(this.doc, id) === text) return;
    this.snapshot();
    S.setCommentText(this.doc, id, text, this.styleId("annotation text"), this.styleId("annotation reference"));
    this.update();
  }

  deleteComment(id: string): void {
    this.snapshot();
    S.deleteComment(this.doc, id);
    this.update();
  }

  replyComment(id: string, text: string): string {
    this.snapshot();
    const r = S.addReply(this.doc, id, { author: this.author, initials: this.authorInitials, date: new Date().toISOString().replace(/\.\d+Z$/, "Z") }, text,
      this.styleId("annotation text"), this.styleId("annotation reference"));
    this.update();
    return r;
  }

  /** Erledigt / wieder öffnen. */
  setCommentDone(id: string, done: boolean): void {
    const m = this.doc.comments.get(id);
    if (!m) return;
    this.snapshot();
    this.doc.setComment({ ...m, done: done || undefined }, id);
    this.update();
  }

  /** Selects a comment's text (Zurück/Weiter, a click on its card). */
  selectComment(id: string): void {
    const r = S.commentRanges(this.doc).find((x) => x.id === id);
    if (!r) return;
    if (this.doc.activeStory !== MAIN) this.switchStory(MAIN, MAIN);
    this.setSel(r.start, r.end);
    this.update({ scroll: true });
  }

  /** A comment card the user clicked: shown active until the caret moves in the text. */
  private focusedComment: string | null = null;

  focusComment(id: string | null): void {
    this.focusedComment = id;
    this.update();
  }

  /** The comment shown active: a clicked card, else the one around the caret. */
  activeComment(): string | null {
    return this.focusedComment ?? this.commentAtCaret();
  }

  private revCache: { main: Para[]; has: boolean } | null = null;

  /** Whether the main text has tracked changes (Annehmen/Ablehnen enabled). */
  hasRevisions(): boolean {
    const main = this.doc.main;
    if (this.revCache?.main !== main) this.revCache = { main, has: R.findRevision(main, { p: 0, o: -1 }, 1) !== null };
    return this.revCache.has;
  }

  /** The comment at the caret (its text contains the caret). */
  commentAtCaret(): string | null {
    if (this.doc.activeStory !== MAIN) return null;
    const h = this.sel.head;
    const hit = S.commentRanges(this.doc).filter((r) => cmp(r.start, h) <= 0 && cmp(h, r.end) <= 0);
    return hit.length ? hit[hit.length - 1].id : null;
  }

  /** Where each comment's text starts, in the view's container coordinates (for the cards). */
  commentTops(): Map<string, number> {
    const out = new Map<string, number>();
    for (const r of S.commentRanges(this.doc)) out.set(r.id, this.view.mainLineTop(r.start));
    return out;
  }

  /** The markup area right of the pages (comment cards); `width` 0 removes it. */
  get markupLayer(): HTMLElement {
    return this.view.markupLayer;
  }

  setMarkupWidth(width: number): void {
    this.view.setMarkupWidth(width);
  }

  // ----- tracked changes -------------------------------------------------------------------------

  get trackChanges(): boolean {
    return !!this.doc.settings.trackChanges;
  }

  /** Änderungen nachverfolgen (Strg+Umschalt+E). */
  setTrackChanges(on: boolean): void {
    this.doc.setSettings({ ...this.doc.settings, trackChanges: on || undefined });
    this.insRev = null;
    this.emit(true);
  }

  /** Annehmen / Ablehnen: the selection's changes, the change at the caret, or all. */
  resolveChanges(accept: boolean, scope: "selection" | "all" | "next" = "selection"): void {
    const paras = this.doc.paras;
    let [from, to] = this.range();
    if (scope === "all") { from = { p: 0, o: 0 }; to = this.doc.end(); }
    else if (cmp(from, to) === 0) {
      // The change at the caret (or the next one).
      const prev = R.findRevision(paras, { p: from.p, o: from.o + 1 }, -1);
      const span = prev && cmp(prev.from, from) <= 0 && cmp(from, prev.to) <= 0 ? prev : R.findRevision(paras, from, 1);
      if (!span) return;
      [from, to] = [span.from, span.to];
    }
    if (!R.hasRevisions(paras, from, to)) return;
    this.snapshot();
    R.resolveRevisions(this.doc, from, to, accept);
    const at = this.fixPos(from);
    this.setSel(at, at);
    this.update({ scroll: true });
    if (scope === "next") this.nextChange(1);
  }

  /** Zurück / Weiter: selects the previous or next change. */
  nextChange(dir: 1 | -1): boolean {
    const h = dir > 0 ? maxPos(this.sel.anchor, this.sel.head) : minPos(this.sel.anchor, this.sel.head);
    const span = R.findRevision(this.doc.paras, dir > 0 ? { p: h.p, o: h.o - 1 } : h, dir);
    if (!span) return false;
    this.setSel(span.from, span.to);
    this.update({ scroll: true });
    return true;
  }

  /** Überarbeitungsbereich: the main text's tracked changes in order (author, kind, text). */
  revisionList(): { from: Pos; to: Pos; kind: "ins" | "del"; author: string; date: string; text: string }[] {
    return R.listRevisions(this.doc.main).map((s) => ({ from: s.from, to: s.to, kind: s.rev.kind, author: s.rev.author, date: s.rev.date, text: s.text }));
  }

  /** Markup anzeigen: all changes marked, the result without marks, or the original. */
  setMarkup(markup: "all" | "simple" | "none" | "original"): void {
    this.markupMode = markup;
    this.view.markupMode = markup;
    this.layouter.setContext(this.context());
    this.update();
  }

  markupMode: "all" | "simple" | "none" | "original" = "all";

  // ----- format painter -------------------------------------------------------------------

  copyFormat(): CopiedFormat {
    const [from, to] = this.range();
    const p = this.doc.paras[from.p];
    const char = from.o < p.text.length ? this.doc.charPropsAt(from) : insertionProps(p, from.o);
    // Paragraph formatting is copied when the selection includes a paragraph mark or is empty.
    const withPara = !this.hasSelection || to.p > from.p || to.o >= this.doc.paras[to.p].text.length;
    return { char, para: withPara ? p.props : null };
  }

  pasteFormat(fmt: CopiedFormat): void {
    const r = this.formatRange();
    const [a, b] = this.paraRangeIdx();
    this.edit(() => {
      if (r) this.doc.mapChars(r[0], r[1], () => fmt.char);
      if (fmt.para) this.doc.mapParas(a, b, () => fmt.para!);
    }, false);
  }

  /** Format übertragen: the next selection made with the mouse gets the copied formatting. */
  startPainter(sticky: boolean): void {
    this.painter = { fmt: this.copyFormat(), sticky };
    this.workspace.style.cursor = BRUSH_CURSOR;
    this.emit(false);
  }

  stopPainter(): void {
    this.painter = null;
    this.workspace.style.cursor = "";
    this.emit(false);
  }

  // ----- caret movement -----------------------------------------------------------------

  private charLeft(p: Pos): Pos {
    if (p.o > 0) return { p: p.p, o: p.o - charBefore(this.doc.paras[p.p].text, p.o) };
    return p.p > 0 ? { p: p.p - 1, o: this.doc.paras[p.p - 1].text.length } : p;
  }

  private charRight(p: Pos): Pos {
    const text = this.doc.paras[p.p].text;
    if (p.o < text.length) return { p: p.p, o: p.o + charAfter(text, p.o) };
    return p.p < this.doc.paras.length - 1 ? { p: p.p + 1, o: 0 } : p;
  }

  private wordLeft(p: Pos): Pos {
    if (p.o === 0) return p.p > 0 ? { p: p.p - 1, o: this.doc.paras[p.p - 1].text.length } : p;
    return { p: p.p, o: prevTokenStart(this.doc.paras[p.p].text, p.o) };
  }

  private wordRight(p: Pos): Pos {
    const text = this.doc.paras[p.p].text;
    if (p.o === text.length) return p.p < this.doc.paras.length - 1 ? { p: p.p + 1, o: 0 } : p;
    return { p: p.p, o: nextTokenStart(text, p.o) };
  }

  private vertical(dir: -1 | 1, extend: boolean): void {
    const caret = this.view.caretBox(this.sel.head, this.sel.headUpstream);
    const goal = this.goalX ?? caret.x;
    const { pl } = this.view.placedFor(this.sel.head, this.sel.headUpstream);
    const target = this.lineNeighbour(pl.para, pl.line, dir, goal);
    let res: { pos: Pos; upstream: boolean };
    if (target === null) res = dir < 0 ? { pos: { p: 0, o: 0 }, upstream: false } : { pos: this.doc.end(), upstream: false };
    else res = this.view.posInLine(target, goal);
    if (extend && (this.doc.paras[res.pos.p]?.cell || this.doc.paras[(this.cellAnchor ?? this.sel.anchor).p]?.cell)) this.extendTo(res.pos, dir > 0);
    else this.setSel(extend ? this.sel.anchor : res.pos, res.pos, res.upstream, true);
    this.goalX = goal;
  }

  /** Placed index of a paragraph's first (`last` false) or last line, skipping hidden text. */
  private lineIndex(p: number, last: boolean): number | null {
    const idx = this.view.storyLayout().index[p];
    if (!idx?.length) return null;
    const i = last ? idx[idx.length - 1] : idx[0];
    return this.view.storyLayout().placed[i]?.hide ? null : i;
  }

  /**
   * The line above or below a line, as Word moves with the arrow keys: in text order, but
   * from a table cell to the cell above or below (in the column under the caret), and into
   * a table at the cell under the caret.
   */
  private lineNeighbour(p: number, line: number, dir: -1 | 1, goal: number): number | null {
    const story = this.view.storyLayout();
    const paras = this.doc.paras;
    const idx = story.index[p] ?? [];
    if (idx[line + dir] !== undefined) return idx[line + dir];
    const cell = paras[p].cell;
    let q = p + dir;
    while (q >= 0 && q < paras.length && this.lineIndex(q, dir < 0) === null) q += dir;
    if (q < 0 || q >= paras.length) return null;
    const next = paras[q].cell;
    if (cell && next && next.table.id === cell.table.id && next.row === cell.row && next.col === cell.col) return this.lineIndex(q, dir < 0);
    if (cell) {
      // Leaving a cell: the cell above/below in the same column, else out of the table.
      const t = cell.table;
      const range = this.doc.tableRange(p, t)!;
      let row = cell.row + dir;
      while (row >= 0 && row < t.rows.length) {
        const target = this.cellAt(range, t.id, row, goal);
        if (target !== null) return this.lineIndex(target[dir < 0 ? target.length - 1 : 0], dir < 0);
        row += dir;
      }
      const out = dir > 0 ? range[1] + 1 : range[0] - 1;
      return out >= 0 && out < paras.length ? this.lineIndex(out, dir < 0) : null;
    }
    if (next) {
      // Entering a table: the first or last row, at the cell under the caret.
      const t = next.table;
      const range = this.doc.tableRange(q, t)!;
      const rows = dir > 0 ? Array.from(t.rows.keys()) : Array.from(t.rows.keys()).reverse();
      for (const row of rows) {
        const target = this.cellAt(range, t.id, row, goal);
        if (target !== null) return this.lineIndex(target[dir < 0 ? target.length - 1 : 0], dir < 0);
      }
    }
    return this.lineIndex(q, dir < 0);
  }

  /** Paragraphs of the visible cell of a row under x (page px), or null. */
  private cellAt(range: [number, number], table: number, row: number, x: number): number[] | null {
    const story = this.view.storyLayout();
    let best: { col: number; d: number } | null = null;
    for (const fr of this.layout.frames) {
      if (fr.kind !== "cell" || !fr.cell || fr.story !== this.storyKey) continue;
      const [tid, r, c] = fr.cell.split(":").map(Number);
      if (tid !== table || r !== row) continue;
      const d = x < fr.x ? fr.x - x : x > fr.x + fr.w ? x - (fr.x + fr.w) : 0;
      if (!best || d < best.d) best = { col: c, d };
    }
    if (!best) return null;
    const out: number[] = [];
    for (let i = range[0]; i <= range[1]; i++) {
      const c = this.doc.paras[i].cell;
      if (c && c.row === row && c.col === best.col) out.push(i);
    }
    if (!out.length || out.every((i) => story.placed[story.index[i]?.[0] ?? -1]?.hide)) return null;
    return out;
  }

  private lineEdge(end: boolean): { pos: Pos; upstream: boolean } {
    const { pl } = this.view.placedFor(this.sel.head, this.sel.headUpstream);
    const lines = this.view.storyLayout().paras[pl.para].lines;
    const ln = lines[pl.line];
    if (!end) return { pos: { p: pl.para, o: ln.start }, upstream: false };
    const isLast = pl.line === lines.length - 1;
    const text = this.doc.paras[pl.para].text;
    let o = ln.end;
    if (o > ln.start && (text[o - 1] === LINE_BREAK || text[o - 1] === PAGE_BREAK)) return { pos: { p: pl.para, o: o - 1 }, upstream: false };
    return { pos: { p: pl.para, o }, upstream: !isLast };
  }

  private page(dir: -1 | 1, extend: boolean): void {
    const pt = this.view.clientPoint(this.sel.head, this.sel.headUpstream);
    const ws = this.workspace;
    const before = ws.scrollTop;
    ws.scrollTop += dir * ws.clientHeight;
    const dy = ws.scrollTop - before;
    this.update();
    const hit = this.view.hitLine(pt.x, pt.y + pt.h / 2 + (dy === 0 ? dir * ws.clientHeight : 0));
    const res = this.view.posInLine(hit.index, hit.x);
    this.setSel(extend ? this.sel.anchor : res.pos, res.pos, res.upstream);
  }

  private collapseOrMove(extend: boolean, moved: Pos, collapseTo: "start" | "end"): void {
    const [from, to] = this.range();
    if (!extend && cmp(from, to) !== 0) {
      const p = collapseTo === "start" ? from : to;
      this.setSel(p, p);
    } else if (extend) {
      this.extendTo(moved, cmp(moved, this.sel.head) > 0);
    } else {
      this.setSel(moved, moved);
    }
  }

  /** The anchor a keyboard selection had before it was widened to whole cells. */
  private cellAnchor: Pos | null = null;

  /**
   * Extends the selection; once it reaches another cell of the same table it covers whole
   * cells, as in Word (the end then sits at the start of the cell after the last one).
   */
  private extendTo(moved: Pos, forward: boolean): void {
    const anchor = this.cellAnchor ?? this.sel.anchor;
    const paras = this.doc.paras;
    const ac = paras[anchor.p]?.cell, hc = paras[moved.p]?.cell;
    type C = NonNullable<typeof ac>;
    const sameCell = (i: number, c: C) => { const x = paras[i]?.cell; return !!x && x.table.id === c.table.id && x.row === c.row && x.col === c.col; };
    const sameRow = (i: number, c: C) => { const x = paras[i]?.cell; return !!x && x.table.id === c.table.id && x.row === c.row; };
    const first = (i: number, same: (i: number, c: C) => boolean): Pos => { const c = paras[i].cell!; while (i > 0 && same(i - 1, c)) i--; return { p: i, o: 0 }; };
    const after = (i: number, same: (i: number, c: C) => boolean): Pos => {
      const c = paras[i].cell!;
      while (i + 1 < paras.length && same(i + 1, c)) i++;
      return i + 1 < paras.length ? { p: i + 1, o: 0 } : { p: i, o: paras[i].text.length };
    };
    const down = cmp(moved, anchor) > 0;
    const set = (a: Pos, h: Pos) => { this.setSel(a, h); this.cellAnchor = anchor; };
    if (ac && hc && ac.table.id === hc.table.id && !(ac.row === hc.row && ac.col === hc.col)) {
      // Whole cells. Reaching a cell's start going forward ends the selection after the cell before.
      const a = down ? first(anchor.p, sameCell) : after(anchor.p, sameCell);
      const atCellStart = moved.o === 0 && cmp(first(moved.p, sameCell), moved) === 0;
      const h = forward ? (atCellStart ? moved : after(moved.p, sameCell)) : first(moved.p, sameCell);
      set(a, h);
      return;
    }
    if (ac && (!hc || hc.table.id !== ac.table.id)) {
      // From a cell out of the table: whole rows from the anchor's row.
      set(down ? first(anchor.p, sameRow) : after(anchor.p, sameRow), moved);
      return;
    }
    if (!ac && hc) {
      // From outside into a table: whole rows up to the head's row.
      set(anchor, down ? after(moved.p, sameRow) : first(moved.p, sameRow));
      return;
    }
    this.setSel(anchor, moved);
    this.cellAnchor = null;
  }

  // ----- keyboard -----------------------------------------------------------------------

  private onBeforeInput(e: InputEvent): void {
    if (e.inputType === "insertCompositionText" && !this.readOnly) return; // let the IME work in the textarea
    e.preventDefault();
    if (this.readOnly) return;
    if (e.inputType === "insertText" && e.data && this.interceptText?.(e.data)) return;
    // Typing in the text ends the focus on a clicked comment card.
    this.focusedComment = null;
    if (e.inputType === "insertText" && e.data) this.insertText(e.data, e.timeStamp);
  }

  private onCompositionEnd(e: CompositionEvent): void {
    this.input.value = "";
    if (e.data && this.interceptText?.(e.data)) return;
    if (e.data) this.insertText(e.data, e.timeStamp);
  }

  private onKeyDown(e: KeyboardEvent): void {
    if (e.isComposing) return;
    if (this.interceptKey?.(e)) { e.preventDefault(); return; }
    // Lesemodus: moving, selecting and copying only; the arrows and Bild ab/auf turn screens.
    if (this.readOnly) {
      const nav = /^(Arrow|Home$|End$|Page)/.test(e.key);
      const keep = (e.ctrlKey && /^[acf]$/i.test(e.key)) || e.key === "Escape" || /^F\d+$/.test(e.key);
      if (!e.shiftKey && !e.ctrlKey && (e.key === "ArrowRight" || e.key === "ArrowLeft" || e.key === "PageDown" || e.key === "PageUp") && !this.hasSelection) {
        e.preventDefault();
        this.turnScreen(e.key === "ArrowRight" || e.key === "PageDown" ? 1 : -1);
        return;
      }
      if (!nav && !keep) { e.preventDefault(); return; }
    }
    this.capsLock = !!e.getModifierState?.("CapsLock");
    const ctrl = e.ctrlKey && !e.altKey;
    const shift = e.shiftKey;
    const ts = e.timeStamp;
    const h = this.sel.head;
    let handled = true;
    // The arrow keys enter an equation next to the caret, as in Word.
    if ((e.key === "ArrowLeft" || e.key === "ArrowRight") && !shift && !ctrl && !e.altKey && this.onEditMath && !this.hasSelection) {
      const at = e.key === "ArrowRight" ? h : { p: h.p, o: h.o - 1 };
      if (at.o >= 0 && this.mathAt(at)) {
        e.preventDefault();
        this.onEditMath(at, e.key === "ArrowRight" ? "start" : "end");
        return;
      }
    }
    // A selected floating picture or shape moves with the arrow keys (0.35 pt, measured in Word).
    if (this.floatSel && !shift && e.key.startsWith("Arrow")) {
      const step = 0.35;
      const d = e.key === "ArrowLeft" ? [-step, 0] : e.key === "ArrowRight" ? [step, 0] : e.key === "ArrowUp" ? [0, -step] : [0, step];
      e.preventDefault();
      this.moveFloat(d[0], d[1], !this.nudging);
      this.nudging = true;
      this.emit(true);
      return;
    }
    this.nudging = false;
    switch (e.key) {
      case "ArrowLeft":
        this.collapseOrMove(shift, ctrl ? this.wordLeft(h) : this.charLeft(h), "start"); break;
      case "ArrowRight":
        this.collapseOrMove(shift, ctrl ? this.wordRight(h) : this.charRight(h), "end"); break;
      case "ArrowUp":
        if (ctrl) {
          const t = h.o > 0 ? { p: h.p, o: 0 } : { p: Math.max(0, h.p - 1), o: 0 };
          this.setSel(shift ? this.sel.anchor : t, t);
        } else this.vertical(-1, shift);
        break;
      case "ArrowDown":
        if (ctrl) {
          const t = h.p < this.doc.paras.length - 1 ? { p: h.p + 1, o: 0 } : this.doc.end();
          this.setSel(shift ? this.sel.anchor : t, t);
        } else this.vertical(1, shift);
        break;
      case "Home": {
        const t = ctrl ? { pos: { p: 0, o: 0 }, upstream: false } : this.lineEdge(false);
        this.setSel(shift ? this.sel.anchor : t.pos, t.pos, t.upstream);
        break;
      }
      case "End": {
        const t = ctrl ? { pos: this.doc.end(), upstream: false } : this.lineEdge(true);
        this.setSel(shift ? this.sel.anchor : t.pos, t.pos, t.upstream);
        break;
      }
      case "PageUp": this.page(-1, shift); break;
      case "PageDown": this.page(1, shift); break;
      case "Backspace": if (!this.deleteFloat()) this.deleteBackward(ctrl, ts); return void e.preventDefault();
      case "Delete": if (!this.deleteFloat()) this.deleteForward(ctrl, ts); return void e.preventDefault();
      case "Enter":
        e.preventDefault();
        this.typing = false;
        if (shift && !e.ctrlKey) this.insertBreak("line");
        else if (ctrl && !shift) this.insertBreak("page");
        else if (e.ctrlKey && shift) this.insertBreak("column");
        else if (!e.ctrlKey) {
          // Enter in an empty list item ends the list (Word).
          const p = this.doc.paras[h.p];
          if (!this.hasSelection && !p.text.length && numberingOf(p.props, this.doc.styles)) {
            this.edit(() => E.removeList(this.doc, h.p, h.p, this.styleId("List Paragraph")));
            return;
          }
          this.autoCorrectEnter();
          this.replaceSelection("\n", ts);
        }
        return;
      case "Tab":
        e.preventDefault();
        if (e.altKey) return;
        if (e.ctrlKey) { if (this.doc.paras[h.p].cell) this.replaceSelection("\t", ts); return; }
        if (this.tabKey(shift)) return;
        if (!shift && this.firstIndentTab()) return;
        this.replaceSelection("\t", ts);
        this.autoList();
        this.autoCorrect("\t");
        return;
      case "Escape":
        if (this.drawTool) { this.cancelDrawing(); return void e.preventDefault(); }
        if (this.cropping) { this.endCrop(); return void e.preventDefault(); }
        if (this.painter) { this.stopPainter(); return void e.preventDefault(); }
        if (this.objectSelect) { this.setObjectSelect(false); return void e.preventDefault(); }
        if (this.isHeaderFooter()) { this.closeHeaderFooter(); return void e.preventDefault(); }
        if (this.inShapeText) { this.leaveShapeText(true); return void e.preventDefault(); }
        if (this.floatSel) { this.floatSel = null; this.update(); return void e.preventDefault(); }
        handled = false;
        break;
      default:
        handled = false;
        if (ctrl && !shift) {
          const k = e.key.toLowerCase();
          if (k === "a") { this.setSel({ p: 0, o: 0 }, this.doc.end()); handled = true; }
          else if (k === "z") { this.undo(); return void e.preventDefault(); }
          else if (k === "y") { this.redo(); return void e.preventDefault(); }
        }
    }
    if (handled) {
      e.preventDefault();
      this.update({ scroll: true, inputTs: ts });
    }
  }

  // ----- clipboard ----------------------------------------------------------------------

  /**
   * The host's clipboard access (the desktop app): the whole HTML document both ways, so Word's
   * list definitions and RaabeWrite's arrive intact. Without it the browser's clipboard events
   * are used.
   */
  clipboardHost: { readHtml(): Promise<string>; readImage?(): Promise<Blob | null>; write(text: string, html: string): Promise<void> } | null = null;

  /** Each copy or cut (the Office-Zwischenablage collects them): text, HTML, or an object. */
  onCopied: ((text: string, html: string, object: boolean) => void) | null = null;

  private onCopy(e: ClipboardEvent, cut: boolean): void {
    e.preventDefault();
    // A selected floating picture or shape is copied as itself.
    if (this.floatSel && this.copyFloat()) {
      if (this.clipboardHost) void this.clipboardHost.write("", "").catch(() => undefined);
      else e.clipboardData?.setData("text/plain", "");
      this.onCopied?.("", "", true);
      if (cut) this.deleteFloat();
      return;
    }
    if (!this.hasSelection) return;
    const html = this.selectionHtml();
    const text = this.copySelection().replace(/\n/g, "\r\n");
    this.onCopied?.(text, html, false);
    if (this.clipboardHost) void this.clipboardHost.write(text, html).catch(() => undefined);
    else {
      e.clipboardData?.setData("text/plain", text);
      // Formatting for Word and other programs.
      e.clipboardData?.setData("text/html", html);
    }
    if (cut && !this.readOnly) {
      const [from, to] = this.smartRange(...this.range());
      this.deleteRange(from, to);
    }
  }

  private onPaste(e: ClipboardEvent): void {
    e.preventDefault();
    if (this.readOnly) return;
    const dt = e.clipboardData;
    const text = dt?.getData("text/plain") ?? "";
    const html = dt?.getData("text/html") ?? "";
    const item = Array.from(dt?.items ?? []).find((i) => i.kind === "file" && i.type.startsWith("image/"));
    const image = item?.getAsFile() ?? null;
    const host = this.clipboardHost;
    if (host && (html || (!image && !text.trim() && host.readImage))) {
      void (async () => {
        const full = html ? await host.readHtml().catch(() => "") : "";
        // Pictures the web view does not hand out (Word's, the "PNG" format) come from the host.
        const picture = image ?? (!text.trim() && host.readImage ? await host.readImage().catch(() => null) : null);
        this.paste(text, full || html, picture);
      })();
      return;
    }
    this.paste(text, html, image);
  }

  /**
   * A picture from the clipboard (a screenshot, a picture copied in Word or a browser), at
   * the size the HTML gives it, else its own (resolution as in the file, else 96 dpi).
   */
  async pasteImage(blob: Blob, html = ""): Promise<boolean> {
    let bytes = new Uint8Array(await blob.arrayBuffer());
    let info = imageInfo(bytes);
    if (!info) return false;
    // A bitmap from the clipboard is kept as PNG (smaller, and what Word does).
    if (info.type === "image/bmp") {
      try {
        const bmp = await createImageBitmap(new Blob([bytes as BlobPart], { type: "image/bmp" }));
        const cv = new OffscreenCanvas(bmp.width, bmp.height);
        cv.getContext("2d")!.drawImage(bmp, 0, 0);
        const png = new Uint8Array(await (await cv.convertToBlob({ type: "image/png" })).arrayBuffer());
        const pinfo = imageInfo(png);
        if (pinfo) { bytes = png; info = { ...pinfo, dpi: info.dpi }; }
      } catch { /* keep the bitmap */ }
    }
    let { w, h } = naturalSizePt(info);
    const img = html ? new DOMParser().parseFromString(html, "text/html").querySelector("img") : null;
    const hw = Number(img?.getAttribute("width")), hh = Number(img?.getAttribute("height"));
    if (hw > 0 && hh > 0) { w = hw * 0.75; h = hh * 0.75; }
    this.insertImage({ bytes, type: info.type }, w, h);
    return true;
  }

  /** The selection as HTML (for the clipboard). */
  selectionHtml(): string {
    const [from, to] = this.range();
    const labels = listLabels(this.doc.paras, this.doc.numbering, this.doc.styles).slice(from.p, to.p + 1);
    return toHtml(this.doc.fragment(from, to), this.doc.styles, this.doc.numbering, labels);
  }

  /**
   * Einfügen: content copied in RaabeWrite keeps all its formatting; HTML from other programs
   * (Word, browsers) keeps character formatting, headings, lists and tables; otherwise text.
   */
  paste(text: string, html: string, image: Blob | null = null): void {
    const frag = clipboardFragment;
    const t = text.replace(/\r\n?/g, "\n");
    if (frag?.float && !t.trim() && !image && !(html && htmlText(html).trim())) { this.pasteFloat(); return; }
    // Only a picture (no text around it): it goes in as a picture.
    if (image && !this.inBlock() && !t.trim() && !(html && htmlText(html).trim())) { void this.pasteImage(image, html); return; }
    if (frag && frag.text === t) return this.pasteText(t);
    if (html && !this.inBlock()) { this.pasteHtml(html); return; }
    if (t) this.pasteText(t);
  }

  /**
   * Einfügeoptionen > Formatierung zusammenführen: the pasted text takes the formatting at the
   * caret and the target paragraph's style; bold, italic, underlining and super-/subscript of
   * the source stay (Word).
   */
  pasteMerged(text: string, html: string): void {
    if (this.inBlock()) return;
    const t = text.replace(/\r\n?/g, "\n");
    const frag = clipboardFragment;
    let pieces: Para[] | null = null;
    if (frag && frag.text === t) pieces = frag.paras.map((x) => this.portable(x, (k) => k === "field" || k === "link"));
    else if (html) {
      const style = (name: string) => { const id = this.styleId(name); this.ensureStyle(id); return id; };
      pieces = fromHtml(html, style).paras.map(({ list: _l, level: _v, listKey: _k, ...x }) => x);
    }
    if (!pieces?.length) { if (t) this.pasteText(t, true); return; }
    const [a, b] = this.range();
    const base = typedProps(this.doc.paras[a.p], a.o);
    const keep = (r: CharProps) => intern({ ...base, bold: r.bold, italic: r.italic, underline: r.underline, strike: r.strike, vertAlign: r.vertAlign, rev: undefined });
    const target = this.doc.paras[a.p].props;
    const merged = pieces.map((x) => ({
      ...x, cell: undefined, block: undefined, props: { style: target.style }, mark: keep(x.mark),
      runs: x.runs.map((r) => ({ ...r, props: keep(r.props) })),
    }));
    this.snapshot();
    this.typing = false;
    let end = this.replaceTracked(a, b, merged);
    end = this.smartPasteEnd(end);
    this.setSel(end, end);
    this.update({ scroll: true });
  }

  /** Inserts HTML from another program. */
  pasteHtml(html: string): void {
    const style = (name: string) => { const id = this.styleId(name); this.ensureStyle(id); return id; };
    const { paras, lists } = fromHtml(html, style);
    if (!paras.length) return;
    // Lists: Word's own definitions (symbols, fonts, indents) where the HTML has them, else one
    // numbering per run of list paragraphs of the same kind.
    let numbering = this.doc.numbering;
    let curKind: string | undefined, curNum = 0;
    const byKey = new Map<string, number>();
    const listStyle = this.listStyleId();
    const levelsFor = (kind: "bullet" | "number", key?: string): NumLevel[] => {
      const base = kind === "bullet" ? bulletLevelsOf() : numberLevelsOf();
      const defs = key ? lists.get(key) : undefined;
      if (!defs) return base;
      return base.map((b, i) => {
        const d = defs[i];
        if (!d) return b;
        const format = d.format ?? b.format;
        return {
          ...b, format, text: d.text ?? (format === "bullet" ? b.text : `%${i + 1}.`), start: d.start ?? 1, align: d.align ?? b.align,
          suffix: d.suffix ?? b.suffix,
          pPr: { ...b.pPr, indLeft: d.indLeft ?? b.pPr.indLeft, indFirst: d.indFirst ?? b.pPr.indFirst },
          rPr: d.font ? { ...b.rPr, font: d.font } : format === "bullet" ? b.rPr : {},
        };
      });
    };
    const pieces: Para[] = paras.map((p) => {
      const { list, level, listKey, ...para } = p;
      if (!list) { curKind = undefined; return para; }
      if (listKey && byKey.has(listKey)) curNum = byKey.get(listKey)!;
      else if (listKey || list !== curKind) {
        const made = addList(numbering, levelsFor(list, listKey));
        numbering = made.numbering;
        curNum = made.numId;
        if (listKey) byKey.set(listKey, curNum);
      }
      curKind = list;
      return { ...para, props: { ...para.props, style: para.props.style ?? listStyle, numId: curNum, ilvl: level ?? 0 } };
    });
    // A pasted table needs a paragraph before and after it inside the insertion.
    if (pieces[0].cell) pieces.unshift(makePara(""));
    if (pieces[pieces.length - 1].cell) pieces.push(makePara(""));
    this.snapshot();
    this.doc.setNumbering(numbering);
    const [a, b] = this.range();
    if (this.doc.paras[a.p].cell && pieces.some((x) => x.cell)) {
      // No tables inside table cells: their text comes in as paragraphs.
      for (let i = 0; i < pieces.length; i++) pieces[i] = { ...pieces[i], cell: undefined };
    }
    const end = this.replaceTracked(a, b, pieces);
    this.setSel(end, end);
    this.requireFonts();
    this.update({ scroll: true });
  }

  // ----- mouse --------------------------------------------------------------------------

  private wordRange(pos: Pos): [Pos, Pos] {
    const text = this.doc.paras[pos.p].text;
    if (!text.length) return [pos, pos];
    const [a, b] = tokenAt(text, Math.min(pos.o, text.length - 1));
    return [{ p: pos.p, o: a }, { p: pos.p, o: b }];
  }

  private paraRange(p: number): [Pos, Pos] {
    const end = p < this.doc.paras.length - 1 ? { p: p + 1, o: 0 } : this.doc.end();
    return [{ p, o: 0 }, end];
  }

  private onMouseDown(e: MouseEvent): void {
    if (e.button !== 0) return;
    if (this.interceptMouse?.(e)) { e.preventDefault(); return; }
    // Einfügen > Formen: the drag draws the chosen shape.
    if (this.drawTool) { e.preventDefault(); this.focus(); this.drawShape(e); return; }
    // Comment cards handle their own clicks.
    if (this.view.markupLayer.contains(e.target as Node)) return;
    this.focusedComment = null;
    const ws = this.workspace;
    if (e.target === ws && (e.offsetX >= ws.clientWidth || e.offsetY >= ws.clientHeight)) return; // scrollbar
    e.preventDefault();
    this.focus();
    // In a shape's text: clicks inside the shape stay there, others leave it.
    const se = this.selectedShape();
    if (se?.editing) {
      const pt = this.view.pointAt(e.clientX, e.clientY);
      const box = this.floatBox(se.p, se.index);
      const inside = box && pt.page === box.page && pt.x >= box.x - 3 && pt.x <= box.x + box.w + 3 && pt.y >= box.y - 3 && pt.y <= box.y + box.h + 3;
      if (!inside || this.view.handleAt(e.clientX, e.clientY)) this.leaveShapeText(!!inside);
    }
    if (this.storyKey === MAIN && !e.shiftKey && e.detail < 2 && !this.readOnly) {
      const th = this.view.tableHandleAt(e.clientX, e.clientY);
      if (th) { this.dragTableHandle(e, th.kind, th.table); return; }
      // Pictures and table borders take the click before the text.
      const handle = this.view.handleAt(e.clientX, e.clientY);
      if (handle === "rot") { this.dragRotate(); return; }
      if (handle?.startsWith("adj-")) { this.dragAdjust(Number(handle.slice(4))); return; }
      if (handle === "end-0" || handle === "end-1") { this.dragLineEnd(handle === "end-1" ? 1 : 0); return; }
      if (handle === "layout") { const r = this.view.layoutButtonRect(); if (r) this.onLayoutOptions?.(r); return; }
      if (handle?.startsWith("crop-")) { this.dragCrop(e, handle.slice(5)); return; }
      if (this.cropping) {
        const box = this.imageOverlay();
        const pt = this.view.pointAt(e.clientX, e.clientY);
        const inside = box?.crop && pt.page === box.page && pt.x >= box.crop.x && pt.x <= box.crop.x + box.crop.w && pt.y >= box.crop.y && pt.y <= box.crop.y + box.crop.h;
        if (inside) { this.dragCrop(e, null); return; }
        this.cropping = false;
      }
      if (handle) { this.dragImage(e, handle); return; }
      const border = this.view.cellBorderAt(e.clientX, e.clientY);
      if (border) { this.dragColumn(e, border.cell); return; }
      const rowBorder = this.view.rowBorderAt(e.clientX, e.clientY);
      if (rowBorder) { this.dragRow(e, rowBorder.cell, rowBorder.h); return; }
      const deco = this.view.imageAt(e.clientX, e.clientY);
      if (deco) {
        const pt = this.view.pointAt(e.clientX, e.clientY);
        if (deco.shape && !this.objectSelect && this.inShapeTextArea(deco.para, deco.index, pt.page, pt.x, pt.y)) {
          this.editShapeText(deco.para, deco.index, { clientX: e.clientX, clientY: e.clientY });
          return;
        }
        this.selectFloat(deco.para, deco.index);
        this.dragImage(e, null);
        return;
      }
    }
    // Objekte auswählen: clicks pick objects only, the text stays as it is.
    if (this.objectSelect && this.storyKey === MAIN) { if (this.floatSel) { this.floatSel = null; this.update(); } return; }
    if (this.storyKey === MAIN && e.detail === 2 && !this.readOnly) {
      const deco = this.view.imageAt(e.clientX, e.clientY);
      if (deco?.shape && deco.shape.kind !== "connector") { this.editShapeText(deco.para, deco.index, "end"); return; }
    }
    if (!this.storyKey.startsWith("tx:") && !this.chooseStory(e)) return;
    const hit = this.view.hitLine(e.clientX, e.clientY);
    const at = this.view.posInLine(hit.index, hit.x);
    const under = this.view.posInLine(hit.index, hit.x, true).pos;
    const pl = this.view.storyLayout().placed[hit.index];
    // Strg+Klick follows a link (Word).
    if ((e.ctrlKey || e.metaKey) && !e.shiftKey && this.onFollowLink && !hit.inLeftMargin) {
      const link = this.linkAt(under);
      if (link && under.o < link.to.o) { this.onFollowLink(link); return; }
    }

    if (hit.inLeftMargin) {
      // Selection bar: line, paragraph, whole document.
      if (e.detail >= 3) this.setSel({ p: 0, o: 0 }, this.doc.end());
      else if (e.detail === 2) { const [a, b] = this.paraRange(pl.para); this.setSel(a, b); }
      else {
        const lines = this.view.storyLayout().paras[pl.para].lines;
        const ln = lines[pl.line];
        const last = pl.line === lines.length - 1;
        const end = last ? this.paraRange(pl.para)[1] : { p: pl.para, o: ln.end };
        this.setSel({ p: pl.para, o: ln.start }, end, !last);
      }
      this.drag = null;
    } else if (e.shiftKey) {
      this.setSel(this.sel.anchor, at.pos, at.upstream);
      this.drag = { granularity: "char", origin: [this.sel.anchor, this.sel.anchor] };
    } else if (e.detail === 2) {
      const r = this.wordRange(under);
      this.setSel(r[0], r[1]);
      this.drag = { granularity: "word", origin: r };
    } else if (e.detail >= 3) {
      const r = this.paraRange(under.p);
      this.setSel(r[0], r[1]);
      this.drag = { granularity: "para", origin: r };
    } else if (e.ctrlKey) {
      const [a, b] = sentenceAt(this.doc.paras[under.p].text, under.o);
      this.setSel({ p: under.p, o: a }, { p: under.p, o: b });
      this.drag = null;
    } else if (this.onEditMath && this.mathAt(under) && this.overObject(hit.index, hit.x, under)) {
      // A click on an equation opens it for editing.
      this.setSel(under, { p: under.p, o: under.o + 1 });
      this.update();
      this.onEditMath(under, "click", { x: e.clientX, y: e.clientY });
      return;
    } else if (this.isImageAt(under) && this.overObject(hit.index, hit.x, under)) {
      // A click on a picture in the text selects it.
      this.setSel(under, { p: under.p, o: under.o + 1 });
      this.drag = null;
    } else {
      this.setSel(at.pos, at.pos, at.upstream);
      this.drag = { granularity: "char", origin: [at.pos, at.pos] };
    }
    this.update({ inputTs: e.timeStamp });

    const finishPainter = () => {
      if (!this.painter) return;
      const fmt = this.painter.fmt;
      if (!this.painter.sticky) this.stopPainter();
      this.pasteFormat(fmt);
    };
    if (!this.drag) { finishPainter(); return; }
    let lastX = e.clientX, lastY = e.clientY;
    let raf = 0;
    const move = (ev: MouseEvent) => { lastX = ev.clientX; lastY = ev.clientY; this.dragTo(lastX, lastY, ev.timeStamp); };
    const autoScroll = () => {
      const r = ws.getBoundingClientRect();
      const dy = lastY < r.top ? lastY - r.top : lastY > r.bottom ? lastY - r.bottom : 0;
      if (dy) { ws.scrollTop += dy / 2; this.dragTo(lastX, lastY); }
      raf = requestAnimationFrame(autoScroll);
    };
    const up = () => {
      removeEventListener("mousemove", move, true);
      removeEventListener("mouseup", up, true);
      cancelAnimationFrame(raf);
      this.drag = null;
      finishPainter();
    };
    addEventListener("mousemove", move, true);
    addEventListener("mouseup", up, true);
    raf = requestAnimationFrame(autoScroll);
  }

  /**
   * Switches the story a click goes to, as Word does: a double click in the header or footer
   * area opens it for editing and a double click on the body closes it again (single clicks
   * on the dimmed body are ignored); footnotes are edited directly. Returns false when the
   * click should do nothing else.
   */
  private chooseStory(e: MouseEvent): boolean {
    const region = this.view.regionAt(e.clientX, e.clientY);
    const inHf = this.isHeaderFooter();
    if (inHf) {
      if (region.kind === "header" || region.kind === "footer") {
        this.openHeaderFooter(region.page, region.kind);
        return true;
      }
      if (e.detail >= 2) { this.closeHeaderFooter(); return true; }
      return false;
    }
    if ((region.kind === "header" || region.kind === "footer") && e.detail >= 2) {
      this.openHeaderFooter(region.page, region.kind);
      return true;
    }
    if (region.kind === "note" && region.key && this.storyKey !== region.key) {
      this.switchStory(region.key, region.key);
    } else if (region.kind !== "note" && this.storyKey !== MAIN && !inHf) {
      this.switchStory(MAIN, MAIN);
    }
    return true;
  }

  /** True while a header or footer is being edited. */
  isHeaderFooter(): boolean {
    return this.storyKey.startsWith("hdr:") || this.storyKey.startsWith("ftr:");
  }

  /** Where the caret was in the body before a header/footer or note was opened. */
  private bodySel: Selection | null = null;

  private switchStory(id: string, key: string, sel?: Selection): void {
    if (!id.startsWith("tx:")) this.shapeEdit = null;
    if (this.doc.activeStory === MAIN && id !== MAIN) this.bodySel = this.sel;
    this.doc.setActive(id);
    this.storyKey = key;
    const restore = id === MAIN ? this.bodySel : null;
    this.sel = sel ?? restore ?? { anchor: { p: 0, o: 0 }, head: { p: 0, o: 0 }, headUpstream: false };
    if (id === MAIN) this.bodySel = null;
    this.pending = null;
    this.goalX = null;
    this.update();
  }

  /**
   * Opens the header or footer of a page for editing, creating it (for the page's section)
   * when the document has none yet.
   */
  openHeaderFooter(page: number, kind: "header" | "footer"): void {
    const pg = this.layout.pages[page];
    if (!pg) return;
    let id = kind === "header" ? pg.header : pg.footer;
    if (!id) {
      this.snapshot();
      id = this.newHeaderFooter(pg.section, kind, pg.index);
    }
    this.switchStory(id, `${id}@${page}`, { anchor: { p: 0, o: 0 }, head: { p: 0, o: 0 }, headUpstream: false });
    this.emit(false);
  }

  closeHeaderFooter(): void {
    if (!this.isHeaderFooter()) return;
    this.switchStory(MAIN, MAIN);
    this.emit(false);
  }

  /** A new empty header/footer story referenced by section `si` for its pages. */
  private newHeaderFooter(si: number, kind: "header" | "footer", pageIndex: number): string {
    const prefix = kind === "header" ? "hdr:" : "ftr:";
    let n = 1;
    while (this.doc.stories.has(`${prefix}${n}`)) n++;
    const id = `${prefix}${n}`;
    const style = this.styleId(kind);
    this.ensureStyle(style);
    this.doc.setStory(id, [{ ...makePara("", { style }), mark: EMPTY }]);
    // Which slot: the section's first page, even pages or the default.
    const page = this.layout.pages[pageIndex];
    const first = this.layout.pages.findIndex((x) => x.section === si) === pageIndex;
    this.updateSection(si, (sect) => {
      const refs = { ...(kind === "header" ? sect.headers : sect.footers) };
      const slot = first && sect.titlePage ? "first" : this.doc.settings.evenOdd && page.number % 2 === 0 ? "even" : "default";
      refs[slot] = id;
      return kind === "header" ? { ...sect, headers: refs } : { ...sect, footers: refs };
    });
    return id;
  }

  /** Changes section `si` (counted in the main text; the last one lives in the settings). */
  updateSection(si: number, change: (s: Section) => Section): void {
    const main = this.doc.main;
    let k = 0;
    for (let i = 0; i < main.length - 1; i++) {
      if (!main[i].sect) continue;
      if (k === si) {
        const next = [...main];
        next[i] = { ...main[i], sect: change(main[i].sect!) };
        this.doc.setStory(MAIN, next);
        return;
      }
      k++;
    }
    this.doc.setSettings({ ...this.doc.settings, section: change(this.doc.settings.section) });
  }

  private isImageAt(pos: Pos): boolean {
    const para = this.doc.paras[pos.p];
    return para?.text[pos.o] === OBJ && this.objectAt(para, pos.o)?.kind === "image";
  }

  /** Whether x (page px) is horizontally over the object at pos in the line `index`. */
  private overObject(index: number, x: number, pos: Pos): boolean {
    const pl = this.view.storyLayout().placed[index];
    if (!pl || pl.para !== pos.p) return false;
    const line = this.view.lineOf(pl);
    const a = pl.x + line.x0 + line.charX[pos.o - line.start], b = pl.x + line.x0 + line.charX[pos.o + 1 - line.start];
    return x >= a && x <= b;
  }

  /** Resizes (from a handle) or moves (floating pictures) the selected picture with the mouse. */
  private dragImage(e: MouseEvent, handle: string | null): void {
    const box0 = this.imageBox();
    if (!box0) return;
    const start = { x: e.clientX, y: e.clientY };
    let first = true, lastDx = 0, lastDy = 0;
    const w0 = box0.w / PT, h0 = box0.h / PT;
    const sel0 = this.selectedImage();
    const isShape = !!sel0?.obj.shape, floating = sel0?.kind === "float";
    let movedX = 0, movedY = 0;
    const move = (ev: MouseEvent) => {
      const z = this.view.zoom;
      const dx = (ev.clientX - start.x) / z / PT, dy = (ev.clientY - start.y) / z / PT;
      if (Math.abs(dx) + Math.abs(dy) < 0.5 && first) return;
      if (handle) {
        let w = w0, hh = h0;
        if (handle.includes("e")) w += dx;
        if (handle.includes("w")) w -= dx;
        if (handle.startsWith("s")) hh += dy;
        if (handle.startsWith("n")) hh -= dy;
        // Corner handles keep a picture's proportions; a shape's only with Umschalt (Word).
        if (handle.length === 2 && isShape === ev.shiftKey) { const k = Math.max(w / w0, hh / h0); w = w0 * k; hh = h0 * k; }
        w = Math.max(isShape ? 1 : 6, w);
        hh = Math.max(isShape ? 1 : 6, hh);
        this.setImageSize(w, hh, first);
        // A floating object keeps the opposite edge where it was.
        if (floating) {
          const mx = handle.includes("w") ? w0 - w : 0, my = handle.startsWith("n") ? h0 - hh : 0;
          if (mx !== movedX || my !== movedY) { this.moveFloat(mx - movedX, my - movedY, false); movedX = mx; movedY = my; }
        }
      } else {
        this.moveFloat(dx - lastDx, dy - lastDy, first);
        lastDx = dx; lastDy = dy;
      }
      first = false;
    };
    const up = () => { removeEventListener("mousemove", move, true); removeEventListener("mouseup", up, true); this.emit(true); };
    addEventListener("mousemove", move, true);
    addEventListener("mouseup", up, true);
    this.update();
  }

  /** Drags a table's column border (the one right of `cellKey`'s cell). */
  private dragColumn(e: MouseEvent, cellKey: string): void {
    const [tid, row, col] = cellKey.split(":").map(Number);
    const p = this.doc.paras.findIndex((x) => x.cell && x.cell.table.id === tid && x.cell.row === row && x.cell.col === col);
    if (p < 0) return;
    const def = this.doc.paras[p].cell!.table;
    const r = def.rows[row];
    let g = -1;
    for (let i = 0; i <= col; i++) g += Math.max(1, r.cells[i].span);
    const start = e.clientX;
    let first = true, last = 0;
    const move = (ev: MouseEvent) => {
      const dx = (ev.clientX - start) / this.view.zoom / PT;
      if (Math.abs(dx - last) < 0.1) return;
      this.resizeColumn(p, g, dx - last, first);
      first = false;
      last = dx;
    };
    const up = () => { removeEventListener("mousemove", move, true); removeEventListener("mouseup", up, true); this.emit(true); };
    addEventListener("mousemove", move, true);
    addEventListener("mouseup", up, true);
  }

  /**
   * Zeilenhöhe by dragging the row's bottom border: the row gets at least that height
   * ("Mindestens"), as Word sets it; one undo step for the whole drag.
   */
  private dragRow(e: MouseEvent, cellKey: string, heightPx: number): void {
    const [tid, row, col] = cellKey.split(":").map(Number);
    const p = this.doc.paras.findIndex((x) => x.cell && x.cell.table.id === tid && x.cell.row === row && x.cell.col === col);
    if (p < 0) return;
    const start = e.clientY;
    let first = true;
    const move = (ev: MouseEvent) => {
      const h = Math.max(4, (heightPx + (ev.clientY - start) / this.view.zoom) / PT);
      if (first) this.snapshot();
      first = false;
      const at = this.doc.paras.findIndex((x) => x.cell && x.cell.table.id === tid && x.cell.row === row);
      if (at < 0) return;
      E.updateRows(this.doc, at, at, (r) => ({ ...r, height: Math.round(h * 20) / 20, heightRule: "atLeast" }));
      this.update();
    };
    const up = () => { removeEventListener("mousemove", move, true); removeEventListener("mouseup", up, true); this.emit(true); };
    addEventListener("mousemove", move, true);
    addEventListener("mouseup", up, true);
  }

  /** The host's context menu (right click or the menu key), at a client point and text position. */
  onContextMenu: ((clientX: number, clientY: number, pos: Pos) => void) | null = null;

  private onContext(e: MouseEvent): void {
    e.preventDefault();
    if (!this.onContextMenu) return;
    const hit = this.view.hitLine(e.clientX, e.clientY);
    const under = this.view.posInLine(hit.index, hit.x, true).pos;
    const deco = this.storyKey === MAIN ? this.view.imageAt(e.clientX, e.clientY) : null;
    if (deco) {
      this.selectFloat(deco.para, deco.index);
      this.onContextMenu(e.clientX, e.clientY, under);
      return;
    }
    const para = this.doc.paras[under.p];
    if (para?.text[under.o] === OBJ && this.objectAt(para, under.o)?.kind === "image") {
      this.setSel(under, { p: under.p, o: under.o + 1 });
      this.update();
      this.onContextMenu(e.clientX, e.clientY, under);
      return;
    }
    const [from, to] = this.range();
    // A right click outside the selection moves the caret there (Word).
    const inside = cmp(under, from) >= 0 && cmp(under, to) < 0;
    if (!inside) {
      const at = this.view.posInLine(hit.index, hit.x);
      this.setSel(at.pos, at.pos, at.upstream);
      this.update();
    }
    this.onContextMenu(e.clientX, e.clientY, under);
  }

  /** The word at a position (letters, digits, inner hyphens and apostrophes). */
  wordAt(pos: Pos): { p: number; a: number; b: number; text: string } | null {
    const text = this.doc.paras[pos.p]?.text ?? "";
    const isW = (ch: string | undefined) => !!ch && /[\p{L}\p{N}'’-]/u.test(ch);
    if (!isW(text[pos.o]) && !isW(text[pos.o - 1])) return null;
    let a = pos.o, b = pos.o;
    while (a > 0 && isW(text[a - 1])) a--;
    while (b < text.length && isW(text[b])) b++;
    while (a < b && /['’-]/.test(text[a])) a++;
    while (b > a && /['’-]/.test(text[b - 1])) b--;
    return b > a ? { p: pos.p, a, b, text: text.slice(a, b) } : null;
  }

  /**
   * Mouse pointer as in Word: the I-beam over the page, the mirrored arrow in the selection
   * bar, the arrow over selected text (it can be dragged) and outside the pages, resize and
   * move pointers on pictures, handles and table borders.
   */
  private hoverCursor(e: MouseEvent): void {
    if (this.painter) return;
    let cursor = "default";
    // Word shows the table handles while the mouse is over a table.
    const tableNow = this.storyKey === MAIN ? this.view.tableAt(e.clientX, e.clientY) : null;
    if (tableNow !== this.hoverTable) { this.hoverTable = tableNow; this.view.render(this.doc, this.layout, this.sel, { active: this.storyKey, cells: this.selectedCells(), image: this.imageOverlay(), table: this.handleTable() }); }
    const th = this.storyKey === MAIN ? this.view.tableHandleAt(e.clientX, e.clientY) : null;
    if (th) { const c = th.kind === "move" ? "move" : "nwse-resize"; if (this.workspace.style.cursor !== c) this.workspace.style.cursor = c; return; }
    if (this.view.pageIndexAt(e.clientX, e.clientY) >= 0) {
      cursor = "text";
      const hd = this.storyKey === MAIN ? this.view.handleAt(e.clientX, e.clientY) : null;
      if (hd) cursor = hd === "n" || hd === "s" ? "ns-resize" : hd === "e" || hd === "w" ? "ew-resize" : hd === "nw" || hd === "se" ? "nwse-resize" : "nesw-resize";
      else if (this.storyKey === MAIN && this.view.cellBorderAt(e.clientX, e.clientY)) cursor = "col-resize";
      else if (this.storyKey === MAIN && this.view.rowBorderAt(e.clientX, e.clientY)) cursor = "row-resize";
      else if (this.storyKey === MAIN && this.view.imageAt(e.clientX, e.clientY)) cursor = "move";
      else {
        const hit = this.view.hitLine(e.clientX, e.clientY);
        if (hit.inLeftMargin) cursor = MIRRORED_ARROW;
        else if (this.hasSelection) {
          const at = this.view.posInLine(hit.index, hit.x, true).pos;
          const [a, b] = this.range();
          if (cmp(a, at) <= 0 && cmp(at, b) < 0) cursor = "default";
        }
      }
    }
    if (this.workspace.style.cursor !== cursor) this.workspace.style.cursor = cursor;
  }

  /** Mouse over a table (Word shows its handles then, and while the caret is in it). */
  private hoverTable: number | null = null;

  private handleTable(): number | undefined {
    if (this.storyKey !== MAIN) return undefined;
    const cell = this.doc.paras[this.sel.head.p]?.cell;
    return cell ? cell.table.id : this.hoverTable ?? undefined;
  }

  private paraOfTable(tid: number): number {
    return this.doc.paras.findIndex((x) => x.cell && x.cell.table.id === tid);
  }

  /**
   * The table handles: a click on the move handle selects the table, dragging it moves the
   * table sideways (its indent); the sizing handle scales columns and rows together.
   */
  private dragTableHandle(e: MouseEvent, kind: "move" | "size", tid: number): void {
    const p = this.paraOfTable(tid);
    if (p < 0) return;
    if (kind === "move") {
      this.setSel({ p, o: 0 }, { p, o: 0 });
      this.selectTablePart("table");
    }
    const def0 = this.doc.paras[p].cell!.table;
    const width0 = def0.grid.reduce((a, b) => a + b, 0);
    const pl = this.layout.placed[this.layout.firstPlaced[p]];
    const pageH = pl ? this.view.tableBounds(this.layout.pages[pl.page], tid) : null;
    const height0 = pageH ? pageH.h / PT : 0;
    const sx = e.clientX, sy = e.clientY;
    let started = false;
    const move = (ev: MouseEvent) => {
      const dx = (ev.clientX - sx) / this.view.zoom / PT, dy = (ev.clientY - sy) / this.view.zoom / PT;
      if (!started && Math.abs(dx) < 1 && Math.abs(dy) < 1) return;
      if (!started) { this.snapshot(); started = true; }
      const at = this.paraOfTable(tid);
      if (at < 0) return;
      if (kind === "move") {
        E.updateTable(this.doc, at, (d) => ({ ...d, indent: Math.round(((def0.indent ?? 0) + dx) * 20) / 20 }));
      } else {
        const fx = Math.max(0.1, (width0 + dx) / width0);
        const fy = height0 > 0 ? Math.max(0.1, (height0 + dy) / height0) : 1;
        E.updateTable(this.doc, at, (d) => ({
          ...d, grid: def0.grid.map((g) => g * fx),
          rows: def0.rows.map((r, i) => ({
            ...r, cells: r.cells.map((c) => (c.width !== undefined ? { ...c, width: c.width * fx } : c)),
            height: fy !== 1 ? Math.max(4, ((r.height ?? 0) || this.rowHeightPt(tid, i)) * fy) : r.height,
            heightRule: fy !== 1 ? (r.heightRule ?? "atLeast") : r.heightRule,
          })),
        }));
      }
      this.update();
    };
    const up = () => { removeEventListener("mousemove", move, true); removeEventListener("mouseup", up, true); this.emit(true); };
    addEventListener("mousemove", move, true);
    addEventListener("mouseup", up, true);
  }

  /** A row's laid-out height in points (from its first cell's frame). */
  private rowHeightPt(tid: number, row: number): number {
    for (const f of this.layout.frames) {
      if (f.kind === "cell" && f.cell?.startsWith(`${tid}:${row}:`)) return f.h / PT;
    }
    return 14;
  }

  private dragTo(clientX: number, clientY: number, ts?: number): void {
    const d = this.drag;
    if (!d) return;
    const hit = this.view.hitLine(clientX, clientY);
    const at = this.view.posInLine(hit.index, hit.x);
    const under = this.view.posInLine(hit.index, hit.x, true).pos;
    const [o1, o2] = d.origin;
    if (d.granularity === "para") {
      const r = this.paraRange(under.p);
      if (cmp(r[0], o1) >= 0) this.setSel(o1, r[1]); else this.setSel(o2, r[0]);
    } else if (d.granularity === "word") {
      const r = this.wordRange(under);
      if (cmp(r[0], o1) >= 0) this.setSel(o1, r[1]); else this.setSel(o2, r[0]);
    } else {
      // Word's "automatically select entire word": once the selection leaves the word it
      // started in, both ends snap to whole words.
      const anchor = o1;
      const [ws, we] = this.wordRange(anchor);
      const inside = at.pos.p === anchor.p && at.pos.o >= ws.o && at.pos.o <= we.o;
      if (inside) this.setSel(anchor, at.pos, at.upstream);
      else if (cmp(at.pos, anchor) > 0) this.setSel(ws, this.wordRange(under)[1]);
      else this.setSel(we, this.wordRange(under)[0]);
    }
    this.update({ inputTs: ts });
  }
}

/** Ids of built-in styles as German Word writes them, mapped to their built-in names. */
const BUILTIN_NAME_BY_ID: Record<string, string> = {
  Standard: "Normal", berschrift1: "heading 1", berschrift2: "heading 2", berschrift3: "heading 3", Titel: "Title",
  Untertitel: "Subtitle", Zitat: "Quote", IntensivesZitat: "Intense Quote", KeinLeerraum: "No Spacing",
  Listenabsatz: "List Paragraph", SchwacheHervorhebung: "Subtle Emphasis", Hervorhebung: "Emphasis",
  IntensiveHervorhebung: "Intense Emphasis", Fett: "Strong", SchwacherVerweis: "Subtle Reference",
  IntensiverVerweis: "Intense Reference", Buchtitel: "Book Title",
};
