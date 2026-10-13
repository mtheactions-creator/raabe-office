// The editor: a hidden textarea receives keyboard and IME input, mouse handling implements
// Word's selection rules, and all edits go through the model. The shell talks to it
// through the public methods below and listens for changes with onChange().

import { PAGE, type StyleId } from "./geometry";
import { markInputHandled } from "./probe";
import { Doc, cmp, maxPos, minPos, type Para, type Pos } from "./doc";
import { layout, type Layout } from "./layout";
import { View, type Selection } from "./view";
import { countWords, isWordChar, nextTokenStart, prevTokenStart, sentenceAt, tokenAt } from "./words";

type Granularity = "char" | "word" | "para";

interface Snapshot {
  paras: Para[];
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

export class Editor {
  doc: Doc;
  layout!: Layout;
  sel: Selection;
  readonly view: View;
  private readonly input: HTMLTextAreaElement;
  private goalX: number | null = null;
  private undoStack: Snapshot[] = [];
  private redoStack: Snapshot[] = [];
  private typing = false;
  private drag: { granularity: Granularity; origin: [Pos, Pos] } | null = null;
  private listeners: ((e: ChangeEvent) => void)[] = [];
  private lastParas: Para[] | null = null;

  constructor(readonly workspace: HTMLElement, paras: Para[]) {
    this.doc = new Doc(paras);
    this.sel = { anchor: { p: 0, o: 0 }, head: { p: 0, o: 0 }, headUpstream: false };
    this.view = new View(workspace);

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
    return countWords(this.doc.paras.map((p) => p.text));
  }

  get canUndo(): boolean {
    return this.undoStack.length > 0;
  }

  get canRedo(): boolean {
    return this.redoStack.length > 0;
  }

  get hasSelection(): boolean {
    return cmp(this.sel.anchor, this.sel.head) !== 0;
  }

  /** Style of the paragraph holding the caret (for the style gallery). */
  currentStyle(): StyleId {
    return this.doc.paras[this.sel.head.p].style;
  }

  get zoom(): number {
    return this.view.zoom;
  }

  setZoom(z: number): void {
    this.view.setZoom(z);
    this.emit(false);
  }

  load(paras: Para[]): void {
    this.doc = new Doc(paras);
    this.sel = { anchor: { p: 0, o: 0 }, head: { p: 0, o: 0 }, headUpstream: false };
    this.undoStack = [];
    this.redoStack = [];
    this.workspace.scrollTop = 0;
    this.update();
  }

  focus(): void {
    this.input.focus({ preventScroll: true });
  }

  // ----- state updates ------------------------------------------------------------------

  update(opts: { scroll?: boolean; inputTs?: number } = {}): void {
    this.layout = layout(this.doc.paras);
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
    this.input.style.left = `${PAGE.gap + c.x}px`;
    this.input.style.top = `${this.view.pageTop(c.page) + c.y}px`;
    // Keep the selected text in the textarea so the browser's copy/cut events fire.
    const selected = this.selectedText();
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
  }

  private range(): [Pos, Pos] {
    return [minPos(this.sel.anchor, this.sel.head), maxPos(this.sel.anchor, this.sel.head)];
  }

  selectedText(): string {
    const [from, to] = this.range();
    return this.doc.slice(from, to);
  }

  // ----- editing ------------------------------------------------------------------------

  private snapshot(): void {
    this.undoStack.push({ paras: this.doc.paras, anchor: this.sel.anchor, head: this.sel.head });
    if (this.undoStack.length > 500) this.undoStack.shift();
    this.redoStack = [];
  }

  private replaceSelection(text: string, inputTs?: number, continueTyping = false): void {
    if (!(continueTyping && this.typing)) this.snapshot();
    const [from, to] = this.range();
    const end = this.doc.replace(from, to, text);
    this.setSel(end, end);
    this.typing = continueTyping;
    this.update({ scroll: true, inputTs });
  }

  insertText(text: string, inputTs?: number): void {
    // Consecutive typing forms one undo step, like Word's "Eingabe".
    this.smartTrimSelection();
    this.replaceSelection(text, inputTs, true);
  }

  /**
   * Word's smart cut and paste: when whole words with their trailing space are replaced,
   * the space survives ("Algen " replaced by typing "Moose" gives "Moose ").
   */
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
    this.doc.replace(from, to, "");
    this.setSel(from, from);
    this.update({ scroll: true, inputTs });
  }

  private deleteBackward(word: boolean, inputTs: number): void {
    const [from, to] = this.range();
    if (cmp(from, to) !== 0) return this.deleteRange(from, to, inputTs);
    const h = this.sel.head;
    if (h.o === 0) {
      if (h.p === 0) return;
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
      if (h.p === this.doc.paras.length - 1) return;
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

  undo(): void {
    const s = this.undoStack.pop();
    if (!s) return;
    this.redoStack.push({ paras: this.doc.paras, anchor: this.sel.anchor, head: this.sel.head });
    this.doc.restore(s.paras);
    this.setSel(s.anchor, s.head);
    this.update({ scroll: true });
  }

  // ----- commands for the shell -----------------------------------------------------------

  selectAll(): void {
    this.setSel({ p: 0, o: 0 }, this.doc.end());
    this.update();
  }

  /** Removes the selection and returns its text (for Ausschneiden). */
  cutSelection(): string {
    const text = this.selectedText();
    const [from, to] = this.range();
    this.deleteRange(from, to);
    return text;
  }

  /** Inserts plain text with Word's smart spacing (for Einfügen). */
  pasteText(text: string): void {
    text = text.replace(/\r\n?/g, "\n");
    if (!text) return;
    this.typing = false;
    // Smart paste: a word pasted directly after a letter gets a separating space.
    const [from] = this.range();
    const before = this.doc.paras[from.p].text[from.o - 1];
    if (isWordChar(before) && isWordChar(text[0])) text = " " + text;
    this.replaceSelection(text);
  }

  /** Applies a paragraph style to every paragraph touched by the selection. */
  setParagraphStyle(style: StyleId): void {
    const [from, to] = this.range();
    this.snapshot();
    this.typing = false;
    this.doc.setStyle(from.p, to.p, style);
    this.update({ scroll: true });
  }

  /** Renders all pages into `target` for printing. */
  renderForPrint(target: HTMLElement): void {
    this.view.renderForPrint(target);
  }

  redo(): void {
    const s = this.redoStack.pop();
    if (!s) return;
    this.undoStack.push({ paras: this.doc.paras, anchor: this.sel.anchor, head: this.sel.head });
    this.doc.restore(s.paras);
    this.setSel(s.anchor, s.head);
    this.update({ scroll: true });
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
    const goal = this.goalX ?? caret.x;
    const target = index + dir;
    let res: { pos: Pos; upstream: boolean };
    if (target < 0) res = { pos: { p: 0, o: 0 }, upstream: false };
    else if (target >= this.layout.placed.length) res = { pos: this.doc.end(), upstream: false };
    else res = this.view.posInLine(target, goal - PAGE.marginLeft);
    this.setSel(extend ? this.sel.anchor : res.pos, res.pos, res.upstream, true);
    this.goalX = goal;
  }

  private lineEdge(end: boolean): { pos: Pos; upstream: boolean } {
    const { pl } = this.view.placedFor(this.sel.head, this.sel.headUpstream);
    const lines = this.layout.lines[pl.para];
    const ln = lines[pl.line];
    if (!end) return { pos: { p: pl.para, o: ln.start }, upstream: false };
    const isLast = pl.line === lines.length - 1;
    return { pos: { p: pl.para, o: ln.end }, upstream: !isLast };
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
      case "Enter": this.typing = false; this.replaceSelection("\n", ts); return void e.preventDefault();
      case "Tab": e.preventDefault(); return;
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
    const text = this.selectedText();
    e.preventDefault();
    if (!text) return;
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
        const ln = this.layout.lines[pl.para][pl.line];
        const last = pl.line === this.layout.lines[pl.para].length - 1;
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

    if (!this.drag) return;
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
