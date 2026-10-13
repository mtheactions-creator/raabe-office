// The editor: a hidden textarea receives keyboard and IME input, mouse handling implements
// Word's selection rules, and all edits go through the document model. The shell calls the
// public commands below and listens for changes with onChange().

import { Doc, cmp, maxPos, minPos, type DocState, type Pos } from "./doc";
import { Fonts, PT } from "./fonts";
import { Layouter, type Layout } from "./layout";
import { markInputHandled } from "./probe";
import {
  LINE_BREAK, OBJ, PAGE_BREAK, insertionProps, intern, type Align, type CharProps, type LineRule,
  type Para, type ParaProps,
} from "./model";
import { builtinStyle, type ResolvedChar, type ResolvedPara } from "./styles";
import { View, type Selection } from "./view";
import { countWords, isWordChar, nextTokenStart, prevTokenStart, sentenceAt, tokenAt } from "./words";

/** Formatting text typed at a position gets: as in Word, never hidden (hidden text is not shown). */
function typedProps(p: Para, o: number): CharProps {
  const c = insertionProps(p, o);
  return c.hidden ? intern({ ...c, hidden: undefined }) : c;
}

type Granularity = "char" | "word" | "para";

interface Snapshot {
  state: { paras: Para[]; styles: Doc["styles"] };
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
  /** Formatting for the next typed text at a collapsed caret (e.g. after Strg+Umschalt+F). */
  private pending: CharProps | null = null;
  private painter: { fmt: CopiedFormat; sticky: boolean } | null = null;

  constructor(readonly workspace: HTMLElement, state: Partial<DocState>, readonly fonts: Fonts) {
    this.doc = new Doc(state);
    this.sel = { anchor: { p: 0, o: 0 }, head: { p: 0, o: 0 }, headUpstream: false };
    this.layouter = new Layouter(this.context());
    this.view = new View(workspace, fonts);
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
    this.update();
  }

  private context() {
    return { styles: this.doc.styles, fonts: this.fonts, page: this.doc.settings.page, defaultTab: this.doc.settings.defaultTab };
  }

  /** Asks the platform for metrics of every font the document uses. */
  private requireFonts(): void {
    const families = new Set<string>([this.doc.styles.theme.minorFont, this.doc.styles.theme.majorFont]);
    for (const p of this.doc.paras) {
      const rp = this.doc.styles.resolvePara(p.props);
      for (const r of p.runs) families.add(this.doc.styles.resolveChar(rp.style, r.props).font);
      families.add(this.doc.styles.resolveChar(rp.style, p.mark).font);
    }
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

  paragraphs(): readonly Para[] {
    return this.doc.paras;
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
    if (this.layouter.context.styles !== this.doc.styles || this.layouter.context.page !== this.doc.settings.page) {
      this.layouter.setContext(this.context());
    }
    this.layout = this.layouter.layout(this.doc.paras);
    this.view.render(this.doc, this.layout, this.sel);
    if (opts.scroll) {
      this.view.scrollIntoView(this.sel.head, this.sel.headUpstream);
      this.view.render(this.doc, this.layout, this.sel);
    }
    this.placeInput();
    if (opts.inputTs !== undefined) markInputHandled(opts.inputTs);
    const content = this.doc.paras !== this.lastParas;
    this.lastParas = this.doc.paras;
    this.emit(content);
  }

  private emit(content: boolean): void {
    for (const l of this.listeners) l({ content });
  }

  private placeInput(): void {
    const c = this.view.caretBox(this.sel.head, this.sel.headUpstream);
    this.input.style.left = `${16 + c.x}px`;
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

  private snapshot(): void {
    this.undoStack.push({ state: { paras: this.doc.paras, styles: this.doc.styles }, anchor: this.sel.anchor, head: this.sel.head });
    if (this.undoStack.length > 500) this.undoStack.shift();
    this.redoStack = [];
  }

  undo(): void {
    const s = this.undoStack.pop();
    if (!s) return;
    this.redoStack.push({ state: { paras: this.doc.paras, styles: this.doc.styles }, anchor: this.sel.anchor, head: this.sel.head });
    this.doc.restore(s.state);
    this.setSel(s.anchor, s.head);
    this.update({ scroll: true });
  }

  redo(): void {
    const s = this.redoStack.pop();
    if (!s) return;
    this.undoStack.push({ state: { paras: this.doc.paras, styles: this.doc.styles }, anchor: this.sel.anchor, head: this.sel.head });
    this.doc.restore(s.state);
    this.setSel(s.anchor, s.head);
    this.update({ scroll: true });
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
      if (h.p === 0 || this.doc.paras[h.p - 1].block || this.doc.paras[h.p].block) return;
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
      if (h.p === this.doc.paras.length - 1 || this.doc.paras[h.p + 1].block || this.doc.paras[h.p].block) return;
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

  /** Inserts a line break (Umschalt+Eingabe) or page break (Strg+Eingabe). */
  insertBreak(kind: "line" | "page"): void {
    this.typing = false;
    this.replaceSelection(kind === "line" ? LINE_BREAK : PAGE_BREAK);
  }

  /** Renders all pages into `target` for printing. */
  renderForPrint(target: HTMLElement): void {
    this.view.renderForPrint(target);
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

  /** Adds a built-in style the document lacks (e.g. "Überschrift 2" in an old file). */
  private ensureStyle(id: string): void {
    if (this.doc.styles.get(id)) return;
    const builtin = builtinStyle(BUILTIN_NAME_BY_ID[id] ?? id);
    if (builtin) this.doc.setStyles(this.doc.styles.withStyle(builtin));
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
    const { index } = this.view.placedFor(this.sel.head, this.sel.headUpstream);
    const caret = this.view.caretBox(this.sel.head, this.sel.headUpstream);
    const marginLeft = this.doc.settings.page.marginLeft * PT;
    const goal = this.goalX ?? caret.x;
    const target = index + dir;
    let res: { pos: Pos; upstream: boolean };
    if (target < 0) res = { pos: { p: 0, o: 0 }, upstream: false };
    else if (target >= this.layout.placed.length) res = { pos: this.doc.end(), upstream: false };
    else res = this.view.posInLine(target, goal - marginLeft);
    this.setSel(extend ? this.sel.anchor : res.pos, res.pos, res.upstream, true);
    this.goalX = goal;
  }

  private lineEdge(end: boolean): { pos: Pos; upstream: boolean } {
    const { pl } = this.view.placedFor(this.sel.head, this.sel.headUpstream);
    const lines = this.layout.paras[pl.para].lines;
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
      case "Backspace": this.deleteBackward(ctrl, ts); return void e.preventDefault();
      case "Delete": this.deleteForward(ctrl, ts); return void e.preventDefault();
      case "Enter":
        e.preventDefault();
        this.typing = false;
        if (shift && !e.ctrlKey) this.insertBreak("line");
        else if (ctrl && !shift) this.insertBreak("page");
        else if (!e.ctrlKey) this.replaceSelection("\n", ts);
        return;
      case "Tab":
        e.preventDefault();
        if (!e.ctrlKey && !e.altKey) this.replaceSelection("\t", ts);
        return;
      case "Escape":
        if (this.painter) { this.stopPainter(); return void e.preventDefault(); }
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
    if (cut) {
      const [from, to] = this.range();
      this.deleteRange(from, to);
    }
  }

  private onPaste(e: ClipboardEvent): void {
    e.preventDefault();
    const text = e.clipboardData?.getData("text/plain");
    if (text) this.pasteText(text);
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
    const hit = this.view.hitLine(e.clientX, e.clientY);
    const at = this.view.posInLine(hit.index, hit.x);
    const under = this.view.posInLine(hit.index, hit.x, true).pos;
    const pl = this.layout.placed[hit.index];

    if (hit.inLeftMargin) {
      // Selection bar: line, paragraph, whole document.
      if (e.detail >= 3) this.setSel({ p: 0, o: 0 }, this.doc.end());
      else if (e.detail === 2) { const [a, b] = this.paraRange(pl.para); this.setSel(a, b); }
      else {
        const lines = this.layout.paras[pl.para].lines;
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
