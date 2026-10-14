// A document: its stories (main text, headers and footers, foot- and endnotes, comments),
// styles, numbering definitions and settings. All edits go through these methods and
// replace paragraph objects (never mutate them), so a snapshot of the state is an undo step.
//
// The editing methods work on the *active* story (`paras`), which is the main text unless
// the caret is in a header, footnote or comment.

import {
  DEFAULT_SECTION, EMPTY, EMPTY_NUMBERING, OBJ, intern, joinParas, makePara, mapRuns, normalizeRuns, propsAt, slicePara,
  textPiece, type CellRef, type CharProps, type CommentMeta, type Numbering, type Para, type ParaProps, type Section,
  type TableDef, type ImageData,
} from "./model";
import { defaultStyleSheet, type StyleSheet } from "./styles";

export interface Pos {
  p: number;
  o: number;
}

export const cmp = (a: Pos, b: Pos): number => a.p - b.p || a.o - b.o;
export const minPos = (a: Pos, b: Pos): Pos => (cmp(a, b) <= 0 ? a : b);
export const maxPos = (a: Pos, b: Pos): Pos => (cmp(a, b) <= 0 ? b : a);

export interface DocSettings {
  /** The last section (w:body/w:sectPr); earlier sections end at paragraphs (Para.sect). */
  section: Section;
  /** Points. */
  defaultTab: number;
  /** Different headers on even and odd pages (w:evenAndOddHeaders). */
  evenOdd?: boolean;
  /** Änderungen nachverfolgen is on (w:trackRevisions). */
  trackChanges?: boolean;
  /** Formeloptionen (m:mathPr); Word's defaults when absent. */
  math?: MathSettings;
  /** Entwurf > Wasserzeichen: VML in every header, as Word writes it. */
  watermark?: Watermark;
  /** Entwurf > Seitenfarbe (w:background). */
  background?: { readonly color: string; readonly colorTheme?: string };
  /** Fuß- und Endnote: where notes go and how they are numbered (w:footnotePr, w:endnotePr). */
  notes?: { readonly footnote?: NoteProps; readonly endnote?: NoteProps };
  /** Zitate und Literaturverzeichnis: the document's sources and citation style (customXml b:Sources). */
  bibliography?: import("./bibliography").Bibliography;
}

/**
 * A watermark (Wasserzeichen): text as WordArt or a picture, centred on the margins behind the
 * text of every page; size in points as the shape has it.
 */
export type Watermark =
  | { readonly kind: "text"; readonly text: string; readonly font: string; readonly size?: number; readonly color: string; readonly semi: boolean;
      readonly diagonal: boolean; readonly width: number; readonly height: number }
  | { readonly kind: "picture"; readonly image: ImageData; readonly scale?: number; readonly washout: boolean; readonly width: number; readonly height: number };

/** Numbering and place of foot- or endnotes (Word's dialog Fuß- und Endnote). */
export interface NoteProps {
  /** Footnotes: pageBottom (Seitenende) or beneathText; endnotes: docEnd or sectEnd. */
  readonly pos?: "pageBottom" | "beneathText" | "docEnd" | "sectEnd";
  /** decimal, lowerLetter, upperLetter, lowerRoman, upperRoman, chicago (*, †, ‡, §). */
  readonly numFmt?: string;
  readonly numStart?: number;
  readonly numRestart?: "continuous" | "eachSect" | "eachPage";
}

/** Formeloptionen of a document (m:mathPr in settings.xml). Lengths in points. */
export interface MathSettings {
  readonly font: string;
  /** Größe von verschachtelten Brüchen für Anzeigeformeln reduzieren. */
  readonly smallFrac: boolean;
  /** Zeilen mit binären und relationalen Operatoren umbrechen: vor, nach, doppelt. */
  readonly brkBin: "before" | "after" | "repeat";
  /** Operatoren für die Subtraktion duplizieren als. */
  readonly brkBinSub: "--" | "-+" | "+-";
  /** Die folgenden Einstellungen für Mathematik in einer eigenen Zeile verwenden. */
  readonly dispDef: boolean;
  readonly lMargin: number;
  readonly rMargin: number;
  /** Ausrichtung of display equations. */
  readonly defJc: "left" | "right" | "center" | "centerGroup";
  /** Einzug für umbrochene Zeilen, or wrapped lines right-aligned. */
  readonly wrapIndent: number;
  readonly wrapRight: boolean;
  /** Integralgrenzwerte / N-äre Grenzwerte: beside (subSup) or above and below (undOvr). */
  readonly intLim: "subSup" | "undOvr";
  readonly naryLim: "subSup" | "undOvr";
}

/** What Word's Normal template sets (a new document). */
export const DEFAULT_MATH: MathSettings = {
  font: "Cambria Math", smallFrac: false, brkBin: "before", brkBinSub: "--", dispDef: true, lMargin: 0, rMargin: 0,
  defJc: "centerGroup", wrapIndent: 72, wrapRight: false, intLim: "subSup", naryLim: "undOvr",
};

export const MAIN = "main";

/** What the file-format layer hands over and gets back. */
export interface DocState {
  paras: Para[];
  styles: StyleSheet;
  settings: DocSettings;
  /** Header/footer ("hdr:…", "ftr:…"), note ("fn:…", "en:…") and comment ("cmt:…") texts. */
  stories?: ReadonlyMap<string, Para[]>;
  numbering?: Numbering;
  comments?: ReadonlyMap<string, CommentMeta>;
}

/** Everything an undo step restores. */
export interface DocSnapshot {
  main: Para[];
  stories: ReadonlyMap<string, Para[]>;
  styles: StyleSheet;
  settings: DocSettings;
  numbering: Numbering;
  comments: ReadonlyMap<string, CommentMeta>;
}

let tableSeq = 1;
/** A fresh table id (ids only need to tell adjacent tables apart). */
export function nextTableId(): number {
  return tableSeq++;
}

const sameCell = (a: CellRef | undefined, b: CellRef | undefined) =>
  a === b || (!!a && !!b && a.table.id === b.table.id && a.row === b.row && a.col === b.col);

export class Doc {
  private mainParas: Para[];
  stories: Map<string, Para[]>;
  styles: StyleSheet;
  settings: DocSettings;
  numbering: Numbering;
  comments: Map<string, CommentMeta>;
  /** The story the editing methods work on. */
  activeStory = MAIN;
  private starts: number[] | null = null;

  constructor(state: Partial<DocState> = {}) {
    this.mainParas = state.paras?.length ? state.paras : [makePara("")];
    this.stories = new Map(state.stories ?? []);
    this.styles = state.styles ?? defaultStyleSheet();
    this.settings = state.settings ?? { section: DEFAULT_SECTION, defaultTab: 35.4 };
    this.numbering = state.numbering ?? EMPTY_NUMBERING;
    this.comments = new Map(state.comments ?? []);
  }

  /** Paragraphs of the active story. */
  get paras(): Para[] {
    return this.activeStory === MAIN ? this.mainParas : (this.stories.get(this.activeStory) ?? this.mainParas);
  }

  set paras(v: Para[]) {
    if (this.activeStory === MAIN) this.mainParas = v;
    else {
      const next = new Map(this.stories);
      next.set(this.activeStory, v);
      this.stories = next;
    }
  }

  /** The main text, whatever story is active. */
  get main(): Para[] {
    return this.mainParas;
  }

  story(id: string): Para[] | undefined {
    return id === MAIN ? this.mainParas : this.stories.get(id);
  }

  /** Makes `id` the story the editing methods work on (creating it if needed). */
  setActive(id: string): void {
    if (id !== MAIN && !this.stories.has(id)) {
      const next = new Map(this.stories);
      next.set(id, [makePara("")]);
      this.stories = next;
    }
    this.activeStory = id;
    this.changed();
  }

  setStory(id: string, paras: Para[]): void {
    if (id === MAIN) { this.mainParas = paras; this.changed(); return; }
    const next = new Map(this.stories);
    next.set(id, paras);
    this.stories = next;
    this.changed();
  }

  deleteStory(id: string): void {
    if (!this.stories.has(id)) return;
    const next = new Map(this.stories);
    next.delete(id);
    this.stories = next;
    if (this.activeStory === id) this.activeStory = MAIN;
    this.changed();
  }

  state(): DocState {
    return {
      paras: this.mainParas, styles: this.styles, settings: this.settings, stories: this.stories,
      numbering: this.numbering, comments: this.comments,
    };
  }

  snapshot(): DocSnapshot {
    return {
      main: this.mainParas, stories: this.stories, styles: this.styles, settings: this.settings,
      numbering: this.numbering, comments: this.comments,
    };
  }

  restore(s: DocSnapshot): void {
    this.mainParas = s.main;
    // Maps are never changed in place (edits make new ones), so the snapshot's can be shared;
    // keeping their identity lets "undo back to the saved state" count as unchanged.
    this.stories = s.stories as Map<string, Para[]>;
    this.styles = s.styles;
    this.settings = s.settings;
    this.numbering = s.numbering;
    this.comments = s.comments as Map<string, CommentMeta>;
    if (this.activeStory !== MAIN && !this.stories.has(this.activeStory)) this.activeStory = MAIN;
    this.changed();
  }

  /** The section each main paragraph belongs to. */
  sectionOf(p: number): Section {
    const paras = this.mainParas;
    for (let i = p; i < paras.length; i++) if (paras[i].sect) return paras[i].sect!;
    return this.settings.section;
  }

  end(): Pos {
    const last = this.paras.length - 1;
    return { p: last, o: this.paras[last].text.length };
  }

  private changed(): void {
    this.starts = null;
  }

  private offsets(): number[] {
    if (!this.starts) {
      const paras = this.paras;
      this.starts = new Array(paras.length);
      let acc = 0;
      for (let i = 0; i < paras.length; i++) {
        this.starts[i] = acc;
        acc += paras[i].text.length + 1;
      }
    }
    return this.starts;
  }

  toOffset(pos: Pos): number {
    return this.offsets()[pos.p] + pos.o;
  }

  fromOffset(offset: number): Pos {
    const s = this.offsets();
    let lo = 0, hi = s.length - 1;
    while (lo < hi) {
      const mid = (lo + hi + 1) >> 1;
      if (s[mid] <= offset) lo = mid; else hi = mid - 1;
    }
    return { p: lo, o: Math.min(offset - s[lo], this.paras[lo].text.length) };
  }

  /** Plain text; paragraphs joined by "\n", objects shown by their display text. */
  text(): string {
    return this.paras.map((p) => p.text).join("\n");
  }

  slice(from: Pos, to: Pos): string {
    const display = (p: Para, a: number, b: number) => {
      let out = "";
      let obj = 0;
      for (let i = 0; i < p.text.length && i < b; i++) {
        const ch = p.text[i];
        if (ch === OBJ) { if (i >= a) out += p.objects?.[obj]?.display ?? ""; obj++; }
        else if (i >= a) out += ch === "\u000b" ? "\n" : ch === "\u000c" || ch === "\u000e" ? "" : ch;
      }
      return out;
    };
    const paras = this.paras;
    if (from.p === to.p) return display(paras[from.p], from.o, to.o);
    const parts = [display(paras[from.p], from.o, paras[from.p].text.length)];
    for (let i = from.p + 1; i < to.p; i++) {
      // Cells of one row are separated by tabs when copied as text, rows by line breaks.
      parts.push(display(paras[i], 0, paras[i].text.length));
    }
    parts.push(display(paras[to.p], 0, to.o));
    return parts.join("\n");
  }

  /**
   * Paragraphs covering [from, to) as pieces (for copying formatted content). Tables that
   * are copied completely stay tables (as a new table); cells of partly copied tables become
   * plain paragraphs.
   */
  fragment(from: Pos, to: Pos): Para[] {
    const paras = this.paras;
    const out: Para[] = from.p === to.p
      ? [slicePara(paras[from.p], from.o, to.o)]
      : [slicePara(paras[from.p], from.o, paras[from.p].text.length), ...paras.slice(from.p + 1, to.p), slicePara(paras[to.p], 0, to.o)];
    const whole = new Map<number, TableDef>();
    for (const t of new Set(out.map((p) => p.cell?.table).filter((t): t is TableDef => !!t))) {
      const range = this.tableRange(from.p, t);
      if (range && range[0] >= from.p && range[1] <= to.p && (range[0] > from.p || from.o === 0) &&
          (range[1] < to.p || to.o >= paras[range[1]].text.length)) {
        whole.set(t.id, { ...t, id: nextTableId() });
      }
    }
    return out.map((p) => {
      if (!p.cell) return p;
      const t = whole.get(p.cell.table.id);
      return t ? { ...p, cell: { ...p.cell, table: t } } : { ...p, cell: undefined };
    });
  }

  /** First and last paragraph index of the table `t` (searching around `near`). */
  tableRange(near: number, t: TableDef): [number, number] | null {
    const paras = this.paras;
    let i = Math.min(near, paras.length - 1);
    // Find any paragraph of the table, then its ends.
    let hit = -1;
    for (let d = 0; d < paras.length; d++) {
      if (i - d >= 0 && paras[i - d].cell?.table.id === t.id) { hit = i - d; break; }
      if (i + d < paras.length && paras[i + d].cell?.table.id === t.id) { hit = i + d; break; }
    }
    if (hit < 0) return null;
    i = hit;
    let a = i, b = i;
    while (a > 0 && paras[a - 1].cell?.table.id === t.id) a--;
    while (b < paras.length - 1 && paras[b + 1].cell?.table.id === t.id) b++;
    return [a, b];
  }

  /**
   * Replaces [from, to) with `pieces`: the first piece joins the text before, the last
   * piece joins the text after, pieces in between become whole paragraphs. Returns the
   * position after the inserted content.
   *
   * Inside one table cell (or in plain text) this is a simple splice. A range that crosses
   * cell or table boundaries follows Word: tables covered completely are removed, covered
   * cells of other tables are emptied, and paragraphs of different cells never merge.
   */
  replace(from: Pos, to: Pos, pieces: Para[]): Pos {
    const paras = this.paras;
    const fromCell = paras[from.p].cell, toCell = paras[to.p].cell;
    let crosses = !sameCell(fromCell, toCell);
    if (!crosses && !fromCell) for (let i = from.p + 1; i < to.p && !crosses; i++) crosses = !!paras[i].cell;
    if (crosses) return this.replaceAcross(from, to, pieces);
    // Markers at the cut stay before inserted text, and none is lost or doubled.
    const head = slicePara(paras[from.p], 0, from.o, true);
    const tailSrc = paras[to.p];
    const tail = slicePara(tailSrc, to.o, tailSrc.text.length, false, from.p === to.p && from.o === to.o);
    // joinParas(a, b) keeps a's paragraph properties and b's paragraph mark, so: the first
    // paragraph keeps the original's properties and gets the first piece's mark; the last
    // gets the last piece's properties and the original tail's mark.
    const created: Para[] = [];
    if (pieces.length <= 1) {
      const mid = pieces[0] ?? textPiece("", EMPTY);
      created.push(joinParas(joinParas(head, mid), tail));
    } else {
      created.push(joinParas(head, pieces[0]));
      for (let i = 1; i < pieces.length - 1; i++) created.push(pieces[i]);
      created.push(joinParas(pieces[pieces.length - 1], tail));
    }
    // New paragraphs stay in the cell they were typed in; inserted whole tables keep theirs.
    const placed = created.map((p, i) => {
      if (fromCell) return { ...p, cell: fromCell };
      if (i === 0 || i === created.length - 1) return { ...p, cell: undefined };
      return p;
    });
    // A pasted table must not end up as the last thing in a story (Word needs a paragraph after it).
    this.paras = [...paras.slice(0, from.p), ...placed, ...paras.slice(to.p + 1)];
    this.ensureParagraphAfterTables();
    this.changed();
    const last = pieces.length ? pieces[pieces.length - 1] : textPiece("", EMPTY);
    return pieces.length <= 1 ? { p: from.p, o: from.o + last.text.length } : { p: from.p + pieces.length - 1, o: last.text.length };
  }

  private replaceAcross(from: Pos, to: Pos, pieces: Para[]): Pos {
    const paras = this.paras;
    // Group the covered paragraphs: runs of plain paragraphs and runs of one table.
    type Group = { a: number; b: number; table?: TableDef };
    const groups: Group[] = [];
    for (let i = from.p; i <= to.p; i++) {
      const t = paras[i].cell?.table;
      const g = groups[groups.length - 1];
      if (g && (g.table?.id ?? -1) === (t?.id ?? -1)) g.b = i;
      else groups.push({ a: i, b: i, table: t });
    }
    // Items of the result for the covered range, in order.
    type Item = { kind: "head" | "tail" | "cell"; para: Para };
    const items: Item[] = [];
    groups.forEach((g, gi) => {
      if (!g.table) {
        if (gi === 0) items.push({ kind: "head", para: slicePara(paras[g.a], 0, from.o) });
        if (gi === groups.length - 1) items.push({ kind: "tail", para: slicePara(paras[g.b], to.o, paras[g.b].text.length) });
        return;
      }
      const [ta, tb] = this.tableRange(g.a, g.table)!;
      const coversStart = ta > from.p || (ta === from.p && from.o === 0);
      const coversEnd = tb < to.p || (tb === to.p && to.o >= paras[tb].text.length);
      if (coversStart && coversEnd) return; // the whole table goes
      // Every covered cell keeps one empty paragraph (Word empties selected cells).
      let last: CellRef | undefined;
      for (let i = g.a; i <= g.b; i++) {
        const c = paras[i].cell!;
        if (sameCell(c, last)) continue;
        last = c;
        items.push({ kind: "cell", para: { ...makePara("", paras[i].props, EMPTY, paras[i].mark), cell: c } });
      }
    });
    // Join head and tail when nothing remains between them; insert the pieces after the head,
    // or into the first emptied cell when the range began in a table.
    const out: Para[] = [];
    let caret: Pos | null = null;
    const base = from.p;
    for (let k = 0; k < items.length; k++) {
      const it = items[k];
      if (it.kind === "head") {
        const next = items[k + 1];
        const tail = next?.kind === "tail" ? next.para : textPiece("", EMPTY);
        const joined = this.withPieces(it.para, pieces, tail);
        if (next?.kind !== "tail") joined.paras[joined.paras.length - 1] = { ...joined.paras[joined.paras.length - 1], mark: paras[from.p].mark };
        caret = { p: base + out.length + joined.caret.p, o: joined.caret.o };
        out.push(...joined.paras);
        if (next?.kind === "tail") k++;
        continue;
      }
      if (it.kind === "cell" && caret === null && pieces.length) {
        const joined = this.withPieces(it.para, pieces, textPiece("", EMPTY));
        caret = { p: base + out.length + joined.caret.p, o: joined.caret.o };
        out.push(...joined.paras.map((x) => ({ ...x, cell: it.para.cell })));
        continue;
      }
      if (caret === null) caret = { p: base + out.length, o: 0 };
      out.push(it.para);
    }
    const result = [...paras.slice(0, from.p), ...out, ...paras.slice(to.p + 1)];
    if (!result.length) result.push(makePara(""));
    this.paras = result;
    this.ensureParagraphAfterTables();
    this.changed();
    const c = caret ?? { p: Math.min(from.p, this.paras.length - 1), o: 0 };
    return { p: Math.min(c.p, this.paras.length - 1), o: c.o };
  }

  /** head + pieces + tail as paragraphs, with the caret after the pieces. */
  private withPieces(head: Para, pieces: Para[], tail: Para): { paras: Para[]; caret: Pos } {
    if (pieces.length <= 1) {
      const mid = pieces[0] ?? textPiece("", EMPTY);
      return { paras: [{ ...joinParas(joinParas(head, mid), tail), cell: undefined }], caret: { p: 0, o: head.text.length + mid.text.length } };
    }
    const out = [{ ...joinParas(head, pieces[0]), cell: undefined }, ...pieces.slice(1, -1), { ...joinParas(pieces[pieces.length - 1], tail), cell: undefined }];
    return { paras: out, caret: { p: out.length - 1, o: pieces[pieces.length - 1].text.length } };
  }

  /** Word keeps a paragraph after a table at the end of a story or cell-less region. */
  private ensureParagraphAfterTables(): void {
    const paras = this.paras;
    if (paras.length && paras[paras.length - 1].cell) {
      this.paras = [...paras, makePara("")];
    }
  }

  /**
   * Inserts plain text with one formatting; "\n" starts new paragraphs. Word's rules (measured
   * with tools/feel-lab/run.py formatting): text typed at the end of a paragraph also formats
   * its paragraph mark; Enter inserts a new mark with the formatting at the caret, which ends
   * the first paragraph, while the original mark moves on with the text after the caret.
   */
  insertText(from: Pos, to: Pos, text: string, props: CharProps): Pos {
    const lines = text.split("\n");
    if (lines.length === 1) {
      const atEnd = to.o >= this.paras[to.p].text.length;
      const end = this.replace(from, to, [textPiece(text, props)]);
      if (atEnd && text) this.paras[end.p] = { ...this.paras[end.p], mark: intern({ ...props, rev: this.paras[end.p].mark.rev }) };
      return end;
    }
    const src = this.paras[from.p];
    const atEnd = to.p === from.p ? to.o >= src.text.length : to.o >= this.paras[to.p].text.length;
    // Enter at the end of a paragraph continues with the style's "next" style (heading →
    // Standard), dropping direct paragraph formatting of the heading; elsewhere both halves
    // keep the paragraph's formatting. Numbering continues (Word copies numPr).
    const style = this.styles.resolvePara(src.props).style;
    const next = this.styles.get(style)?.next;
    const nextProps: ParaProps = atEnd && next && next !== style
      ? { style: next, numId: src.props.numId, ilvl: src.props.ilvl }
      : src.props;
    // Inserted marks take the inserted formatting; the last paragraph keeps the original mark
    // (replace() joins it with the tail), unless text typed at its end formats it.
    const pieces: Para[] = lines.map((line, i) => ({
      ...textPiece(line, props),
      props: i === 0 ? src.props : nextProps,
      mark: intern(props),
    }));
    // The section break stays with the last paragraph (it belongs to the original mark).
    const end = this.replace(from, to, pieces);
    if (atEnd && lines[lines.length - 1]) this.paras[end.p] = { ...this.paras[end.p], mark: intern({ ...props, rev: this.paras[end.p].mark.rev }) };
    return end;
  }

  delete(from: Pos, to: Pos): void {
    this.replace(from, to, []);
  }

  /** Changes character formatting on [from, to); paragraph marks inside the range change too. */
  mapChars(from: Pos, to: Pos, change: (old: CharProps, para: number) => CharProps): void {
    const out = [...this.paras];
    for (let i = from.p; i <= to.p; i++) {
      const p = out[i];
      if (p.block) continue;
      const a = i === from.p ? from.o : 0;
      const b = i === to.p ? to.o : p.text.length;
      const ch = (old: CharProps) => change(old, i);
      let np = mapRuns(p, a, b, ch);
      // The paragraph mark changes only when the range includes it (Word: not when just the
      // text of a paragraph is selected), and for empty paragraphs.
      if (i < to.p || !p.text.length) np = { ...np, mark: intern(ch(p.mark)) };
      out[i] = np;
    }
    this.paras = out;
    this.changed();
  }

  /** Changes paragraph properties of paragraphs from..to. */
  mapParas(from: number, to: number, change: (old: ParaProps, p: Para) => ParaProps): void {
    this.paras = this.paras.map((p, i) => (i >= from && i <= to && !p.block ? { ...p, props: change(p.props, p) } : p));
    this.changed();
  }

  /** Replaces whole paragraphs from..to with `next` (table and list operations). */
  splice(from: number, to: number, next: Para[]): void {
    const paras = this.paras;
    this.paras = [...paras.slice(0, from), ...next, ...paras.slice(to + 1)];
    if (!this.paras.length) this.paras = [makePara("")];
    this.ensureParagraphAfterTables();
    this.changed();
  }

  /** Replaces same-length text (change case) keeping formatting. */
  mapText(from: Pos, to: Pos, change: (s: string) => string): void {
    const out = [...this.paras];
    for (let i = from.p; i <= to.p; i++) {
      const p = out[i];
      const a = i === from.p ? from.o : 0;
      const b = i === to.p ? to.o : p.text.length;
      const seg = p.text.slice(a, b);
      const changed = change(seg);
      if (changed.length !== seg.length) continue;
      out[i] = { ...p, text: p.text.slice(0, a) + changed + p.text.slice(b) };
    }
    this.paras = out;
    this.changed();
  }

  setStyles(styles: StyleSheet): void {
    this.styles = styles;
  }

  setSettings(settings: DocSettings): void {
    this.settings = settings;
  }

  setNumbering(n: Numbering): void {
    this.numbering = n;
  }

  setComment(meta: CommentMeta | null, id: string): void {
    const next = new Map(this.comments);
    if (meta) next.set(id, meta); else next.delete(id);
    this.comments = next;
  }

  /** Formatting that new text takes at a position. */
  charPropsAt(pos: Pos): CharProps {
    const p = this.paras[pos.p];
    return pos.o < p.text.length ? propsAt(p, pos.o) : p.mark;
  }

  /** Whole paragraphs for a range of indices (used for normalising after edits). */
  normalize(i: number): void {
    const p = this.paras[i];
    this.paras[i] = { ...p, runs: normalizeRuns(p.runs) };
  }
}
