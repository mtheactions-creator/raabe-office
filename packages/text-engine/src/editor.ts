// The editor: a hidden textarea receives keyboard and IME input, mouse handling implements
// Word's selection rules, and all edits go through the document model. The shell calls the
// public commands below and listens for changes with onChange().

import { Doc, MAIN, cmp, maxPos, minPos, type DocSnapshot, type DocState, type Pos } from "./doc";
import { Fonts, PT } from "./fonts";
import { layoutDocument, type Layout } from "./flow";
import { Layouter, type LayoutContext } from "./layout";
import { markInputHandled } from "./probe";
import {
  EMPTY, LINE_BREAK, OBJ, PAGE_BREAK, insertionProps, intern, makePara, type Align, type CharProps, type LineRule,
  type Para, type ParaProps, type Section,
} from "./model";
import { builtinStyle, type ResolvedChar, type ResolvedPara } from "./styles";
import { View, storyIdOf, type Selection, type TextMark } from "./view";
import * as E from "./edits";
import { fromHtml, toHtml } from "./html";
import { COLUMN_BREAK, type CellDef, type ImageData, type RowDef, type TableDef } from "./model";
import { addList, bulletLevels as bulletLevelsOf, numberLevels as numberLevelsOf, numberingOf, restartList } from "./numbering";
import type { BorderSpec, Borders, FloatObject, InlineObject, NumLevel } from "./model";
import { countWords, isWordChar, nextTokenStart, prevTokenStart, sentenceAt, tokenAt } from "./words";

/** Formatting text typed at a position gets: as in Word, never hidden (hidden text is not shown). */
function typedProps(p: Para, o: number): CharProps {
  const c = insertionProps(p, o);
  return c.hidden ? intern({ ...c, hidden: undefined }) : c;
}

type Granularity = "char" | "word" | "para";

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

let clipboardFragment: { text: string; paras: Para[]; source: Editor } | null = null;

const BRUSH_CURSOR = `url("data:image/svg+xml,${encodeURIComponent(
  '<svg xmlns="http://www.w3.org/2000/svg" width="20" height="24"><path d="M2 2v20M2 2h3M2 22h3" stroke="black" stroke-width="1.5" fill="none"/>' +
  '<path d="M8 4h9v5H8z" fill="white" stroke="black"/><path d="M12 9v4h1v6" stroke="black" fill="none"/></svg>')}") 2 12, text`;

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
    workspace.addEventListener("contextmenu", (e) => this.onContext(e));
    this.update();
  }

  private context(): LayoutContext {
    return { styles: this.doc.styles, fonts: this.fonts, defaultTab: this.doc.settings.defaultTab, measureMath: this.measureMath };
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
    this.emit(false);
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
    if (ctx.styles !== this.doc.styles || ctx.defaultTab !== this.doc.settings.defaultTab || ctx.measureMath !== this.measureMath) {
      this.layouter.setContext(this.context());
    }
    this.layout = layoutDocument(this.layouter, {
      main: this.doc.main, stories: this.doc.stories, settings: this.doc.settings, numbering: this.doc.numbering,
    });
    this.fixStoryKey();
    this.clampSelection();
    const ropts = { active: this.storyKey, cells: this.selectedCells(), image: this.imageBox() ?? undefined };
    this.view.render(this.doc, this.layout, this.sel, ropts);
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
    const [from, to] = this.range();
    const a = this.doc.paras[from.p]?.cell, b = this.doc.paras[to.p]?.cell;
    if (from.p === to.p || (!a && !b)) return undefined;
    if (a && b && a.table.id === b.table.id && a.row === b.row && a.col === b.col) return undefined;
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
    const list: { c: ResolvedChar; raw: CharProps }[] = [];
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
          list.push({ c: this.resolvedAt(i, r.props), raw: r.props });
        }
      }
      if (!list.length) list.push({ c: this.resolvedAt(from.p, this.doc.paras[from.p].mark), raw: this.doc.paras[from.p].mark });
    }
    const same = <T>(f: (x: { c: ResolvedChar; raw: CharProps }) => T): T | undefined => {
      const v = f(list[0]);
      return list.every((x) => f(x) === v) ? v : undefined;
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
    if (!(continueTyping && this.typing)) this.snapshot();
    const [from, to] = this.range();
    const fmt = props ?? this.pending ?? (cmp(from, to) !== 0
      ? (from.o < this.doc.paras[from.p].text.length ? this.doc.charPropsAt(from) : typedProps(this.doc.paras[from.p], from.o))
      : typedProps(this.doc.paras[from.p], from.o));
    const pending = this.pending;
    const end = this.doc.insertText(from, to, text, fmt);
    this.setSel(end, end);
    // Formatting chosen at the caret applies to what is typed next; Enter drops it (Word: the
    // new paragraph takes its formatting from the paragraph mark).
    if (pending) this.pending = null;
    this.typing = continueTyping;
    this.update({ scroll: true, inputTs });
  }

  insertText(text: string, inputTs?: number): void {
    this.smartTrimSelection();
    this.replaceSelection(text, inputTs, true);
    if (text === " " || text === "\t") this.autoList();
  }

  /**
   * AutoFormat während der Eingabe: "1. ", "a) ", "* " or "- " typed at the start of a
   * paragraph turn it into a list item (Strg+Z brings the typed text back).
   */
  private autoList(): void {
    const h = this.sel.head;
    const para = this.doc.paras[h.p];
    if (para.block || numberingOf(para.props, this.doc.styles) || h.o < 2 || h.o > 5) return;
    const prefix = para.text.slice(0, h.o - 1);
    if (para.text.slice(h.o - 1, h.o) !== " " && para.text.slice(h.o - 1, h.o) !== "\t") return;
    const auto = E.autoListLevels(prefix);
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

  private deleteRange(from: Pos, to: Pos, inputTs?: number): void {
    if (cmp(from, to) === 0) return;
    this.snapshot();
    this.doc.delete(from, to);
    this.setSel(from, from);
    this.update({ scroll: true, inputTs });
  }

  private deleteBackward(word: boolean, inputTs: number): void {
    const [from, to] = this.range();
    if (cmp(from, to) !== 0) return this.deleteRange(from, to, inputTs);
    const h = this.sel.head;
    if (h.o === 0) {
      const paras = this.doc.paras;
      // At the start of a list item the number goes first (the text stays indented).
      if (numberingOf(paras[h.p].props, this.doc.styles) && !paras[h.p].block) {
        this.edit(() => E.dropListLabel(this.doc, h.p));
        return;
      }
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
    const start = word ? prevTokenStart(text, h.o) : h.o - 1;
    this.deleteRange({ p: h.p, o: start }, h, inputTs);
  }

  private deleteForward(word: boolean, inputTs: number): void {
    const [from, to] = this.range();
    if (cmp(from, to) !== 0) return this.deleteRange(from, to, inputTs);
    const h = this.sel.head;
    const text = this.doc.paras[h.p].text;
    if (h.o === text.length) {
      const paras = this.doc.paras;
      if (h.p === paras.length - 1 || paras[h.p + 1].block || paras[h.p].block) return;
      const cur = paras[h.p].cell, next = paras[h.p + 1].cell;
      if (cur || next) {
        const same = cur && next && cur.table.id === next.table.id && cur.row === next.row && cur.col === next.col;
        if (!same) {
          // An empty paragraph before a table goes; otherwise nothing happens (Word).
          if (!cur && next && !text.length) {
            this.snapshot();
            this.doc.splice(h.p, h.p, []);
            this.setSel({ p: h.p, o: 0 }, { p: h.p, o: 0 });
            this.update({ scroll: true, inputTs });
          }
          return;
        }
      }
      return this.deleteRange(h, { p: h.p + 1, o: 0 }, inputTs);
    }
    let end = h.o + 1;
    if (word) {
      // Inside a word, Strg+Entf deletes only the rest of the word, not the space after it.
      end = nextTokenStart(text, h.o);
      if (isWordChar(text[h.o - 1]) && isWordChar(text[h.o])) {
        end = h.o;
        while (end < text.length && isWordChar(text[end])) end++;
      }
    }
    this.deleteRange(h, { p: h.p, o: end }, inputTs);
  }

  // ----- commands for the shell -----------------------------------------------------------

  selectAll(): void {
    this.setSel({ p: 0, o: 0 }, this.doc.end());
    this.update();
  }

  /** Puts the selection on the internal clipboard (with formatting) and returns its text. */
  copySelection(): string {
    const [from, to] = this.range();
    const text = this.selectedText();
    const paras = this.doc.fragment(from, to).map((p) => this.portable(p));
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
    const [from, to] = this.range();
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
      let at = a;
      if (smartSpace) at = this.doc.insertText(a, b, " ", typedProps(this.doc.paras[a.p], a.o));
      const end = this.doc.replace(at, smartSpace ? at : b, pieces);
      this.setSel(end, end);
      this.requireFonts();
      this.update({ scroll: true });
      return;
    }
    this.replaceSelection(smartSpace ? ` ${text}` : text);
  }

  /** Inserts a line break (Umschalt+Eingabe), page break (Strg+Eingabe) or column break. */
  insertBreak(kind: "line" | "page" | "column"): void {
    this.typing = false;
    this.replaceSelection(kind === "line" ? LINE_BREAK : kind === "page" ? PAGE_BREAK : COLUMN_BREAK);
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
      if (!E.changeListLevel(this.doc, from.p, to.p, back ? -1 : 1)) this.undoStack.pop();
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

  listKind(): E.ListKind | null {
    return E.listKind(this.doc, this.doc.paras[this.sel.head.p]);
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

  /** Neu beginnen mit 1: the list item at the caret and the rest of its list start again. */
  restartNumbering(): void {
    const h = this.sel.head;
    const n = numberingOf(this.doc.paras[h.p].props, this.doc.styles);
    if (!n) return;
    this.snapshot();
    const made = restartList(this.doc.numbering, n.numId, n.ilvl, 1);
    this.doc.setNumbering(made.numbering);
    this.doc.mapParas(h.p, this.doc.paras.length - 1, (old, para) =>
      numberingOf(para.props, this.doc.styles)?.numId === n.numId ? { ...old, numId: made.numId, ilvl: old.ilvl ?? n.ilvl } : old);
    this.update({ scroll: true });
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
  private currentPage(): number {
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
    if (!img || !(width > 2) || !(height > 2)) return;
    if (snapshot) this.snapshot();
    if (img.kind === "inline") this.replaceObject(img.p, img.o, { ...img.obj, width, height });
    else this.replaceFloat(img.p, img.index, { ...img.obj, width, height });
    this.update();
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
    return { page: a.page, x: a.x, y: pl.y + line.baseline - hh, w, h: hh };
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
        wrap, side: "bothSides", dist: [0, 0, 9, 9],
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
      const obj: InlineObject = { kind: "image", data: undefined, display: "", width: f.width, height: f.height, image: f.image };
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

  /** Deletes the selected floating picture. */
  deleteFloat(): boolean {
    const img = this.selectedImage();
    if (!img || img.kind !== "float") return false;
    this.snapshot();
    this.replaceFloat(img.p, img.index, null);
    this.floatSel = null;
    this.update();
    return true;
  }

  selectFloat(p: number, index: number): void {
    this.floatSel = { p, index };
    this.sel = { anchor: { p, o: 0 }, head: { p, o: 0 }, headUpstream: false };
    this.update();
  }

  // ----- find and replace ------------------------------------------------------------------

  private findMarks: WeakMap<Para, TextMark[]> | null = null;
  /** Extra marks from the host (spelling errors), by story and paragraph. */
  extraMarks: ((story: string, para: Para) => readonly TextMark[] | undefined) | null = null;
  private mergedMarks = new WeakMap<Para, { f?: readonly TextMark[]; x?: readonly TextMark[]; out: readonly TextMark[] }>();

  private installMarks(): void {
    this.view.marksFor = (story, para) => {
      const f = story === this.doc.activeStory || (story === "main" && this.doc.activeStory === MAIN) ? this.findMarks?.get(para) : undefined;
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
  headings(): { p: number; level: number; text: string }[] {
    const out: { p: number; level: number; text: string }[] = [];
    this.doc.main.forEach((para, p) => {
      const rp = this.doc.styles.resolvePara(para.props);
      const lvl = para.props.outline ?? this.doc.styles.outlineLevel(rp.style);
      if (lvl !== undefined && lvl < 9 && para.text.trim()) out.push({ p, level: lvl + 1, text: para.text.replace(/[\u000b\ufffc]/g, " ").trim() });
    });
    return out;
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
    if (p.o > 0) return { p: p.p, o: p.o - 1 };
    return p.p > 0 ? { p: p.p - 1, o: this.doc.paras[p.p - 1].text.length } : p;
  }

  private charRight(p: Pos): Pos {
    if (p.o < this.doc.paras[p.p].text.length) return { p: p.p, o: p.o + 1 };
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
    this.setSel(extend ? this.sel.anchor : res.pos, res.pos, res.upstream, true);
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
    } else {
      this.setSel(extend ? this.sel.anchor : moved, moved);
    }
  }

  // ----- keyboard -----------------------------------------------------------------------

  private onBeforeInput(e: InputEvent): void {
    if (e.inputType === "insertCompositionText") return; // let the IME work in the textarea
    e.preventDefault();
    if (e.inputType === "insertText" && e.data) this.insertText(e.data, e.timeStamp);
  }

  private onCompositionEnd(e: CompositionEvent): void {
    this.input.value = "";
    if (e.data) this.insertText(e.data, e.timeStamp);
  }

  private onKeyDown(e: KeyboardEvent): void {
    if (e.isComposing) return;
    const ctrl = e.ctrlKey && !e.altKey;
    const shift = e.shiftKey;
    const ts = e.timeStamp;
    const h = this.sel.head;
    let handled = true;
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
          this.replaceSelection("\n", ts);
        }
        return;
      case "Tab":
        e.preventDefault();
        if (e.altKey) return;
        if (e.ctrlKey) { if (this.doc.paras[h.p].cell) this.replaceSelection("\t", ts); return; }
        if (this.tabKey(shift)) return;
        this.replaceSelection("\t", ts);
        return;
      case "Escape":
        if (this.painter) { this.stopPainter(); return void e.preventDefault(); }
        if (this.isHeaderFooter()) { this.closeHeaderFooter(); return void e.preventDefault(); }
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

  private onCopy(e: ClipboardEvent, cut: boolean): void {
    e.preventDefault();
    if (!this.hasSelection) return;
    const text = this.copySelection();
    e.clipboardData?.setData("text/plain", text.replace(/\n/g, "\r\n"));
    // Formatting for Word and other programs.
    e.clipboardData?.setData("text/html", this.selectionHtml());
    if (cut) {
      const [from, to] = this.range();
      this.deleteRange(from, to);
    }
  }

  private onPaste(e: ClipboardEvent): void {
    e.preventDefault();
    const text = e.clipboardData?.getData("text/plain") ?? "";
    const html = e.clipboardData?.getData("text/html") ?? "";
    this.paste(text, html);
  }

  /** The selection as HTML (for the clipboard). */
  selectionHtml(): string {
    const [from, to] = this.range();
    return toHtml(this.doc.fragment(from, to), this.doc.styles);
  }

  /**
   * Einfügen: content copied in RaabeWrite keeps all its formatting; HTML from other programs
   * (Word, browsers) keeps character formatting, headings, lists and tables; otherwise text.
   */
  paste(text: string, html: string): void {
    const frag = clipboardFragment;
    const t = text.replace(/\r\n?/g, "\n");
    if (frag && frag.text === t) return this.pasteText(t);
    if (html && !this.inBlock()) { this.pasteHtml(html); return; }
    if (t) this.pasteText(t);
  }

  /** Inserts HTML from another program. */
  pasteHtml(html: string): void {
    const style = (name: string) => { const id = this.styleId(name); this.ensureStyle(id); return id; };
    const { paras } = fromHtml(html, style);
    if (!paras.length) return;
    // Lists: one numbering per run of list paragraphs of the same kind.
    let numbering = this.doc.numbering;
    let curKind: string | undefined, curNum = 0;
    const listStyle = this.listStyleId();
    const pieces: Para[] = paras.map((p) => {
      const { list, level, ...para } = p;
      if (!list) { curKind = undefined; return para; }
      if (list !== curKind) {
        const made = addList(numbering, list === "bullet" ? bulletLevelsOf() : numberLevelsOf());
        numbering = made.numbering;
        curNum = made.numId;
        curKind = list;
      }
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
    const end = this.doc.replace(a, b, pieces);
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
    const ws = this.workspace;
    if (e.target === ws && (e.offsetX >= ws.clientWidth || e.offsetY >= ws.clientHeight)) return; // scrollbar
    e.preventDefault();
    this.focus();
    if (this.storyKey === MAIN && !e.shiftKey && e.detail < 2) {
      // Pictures and table borders take the click before the text.
      const handle = this.view.handleAt(e.clientX, e.clientY);
      if (handle) { this.dragImage(e, handle); return; }
      const border = this.view.cellBorderAt(e.clientX, e.clientY);
      if (border) { this.dragColumn(e, border.cell); return; }
      const deco = this.view.imageAt(e.clientX, e.clientY);
      if (deco) { this.selectFloat(deco.para, deco.index); this.dragImage(e, null); return; }
    }
    if (!this.chooseStory(e)) return;
    const hit = this.view.hitLine(e.clientX, e.clientY);
    const at = this.view.posInLine(hit.index, hit.x);
    const under = this.view.posInLine(hit.index, hit.x, true).pos;
    const pl = this.view.storyLayout().placed[hit.index];

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
        // Corner handles keep the picture's proportions (as in Word).
        if (handle.length === 2) { const k = Math.max(w / w0, hh / h0); w = w0 * k; hh = h0 * k; }
        this.setImageSize(Math.max(6, w), Math.max(6, hh), first);
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

  /** The host's context menu (right click or the menu key), at a client point and text position. */
  onContextMenu: ((clientX: number, clientY: number, pos: Pos) => void) | null = null;

  private onContext(e: MouseEvent): void {
    e.preventDefault();
    if (!this.onContextMenu) return;
    const hit = this.view.hitLine(e.clientX, e.clientY);
    const under = this.view.posInLine(hit.index, hit.x, true).pos;
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

  /** Mouse pointer over pictures, handles and table borders. */
  private hoverCursor(e: MouseEvent): void {
    if (this.painter) return;
    let cursor = "";
    if (this.storyKey === MAIN) {
      const hd = this.view.handleAt(e.clientX, e.clientY);
      if (hd) cursor = hd === "n" || hd === "s" ? "ns-resize" : hd === "e" || hd === "w" ? "ew-resize" : hd === "nw" || hd === "se" ? "nwse-resize" : "nesw-resize";
      else if (this.view.cellBorderAt(e.clientX, e.clientY)) cursor = "col-resize";
      else if (this.view.imageAt(e.clientX, e.clientY)) cursor = "move";
    }
    if (this.workspace.style.cursor !== cursor) this.workspace.style.cursor = cursor;
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
