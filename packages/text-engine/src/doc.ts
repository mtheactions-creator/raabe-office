// A document: paragraphs plus styles, page setup and settings. All edits go through these
// methods and replace paragraph objects (never mutate them).

import { A4, type PageSetup } from "./layout";
import {
  EMPTY, OBJ, intern, joinParas, makePara, mapRuns, normalizeRuns, propsAt, slicePara, textPiece,
  type CharProps, type Para, type ParaProps,
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
  page: PageSetup;
  /** Points. */
  defaultTab: number;
}

/** What the file-format layer hands over and gets back. */
export interface DocState {
  paras: Para[];
  styles: StyleSheet;
  settings: DocSettings;
}

export class Doc {
  paras: Para[];
  styles: StyleSheet;
  settings: DocSettings;
  private starts: number[] | null = null;

  constructor(state: Partial<DocState> = {}) {
    this.paras = state.paras?.length ? state.paras : [makePara("")];
    this.styles = state.styles ?? defaultStyleSheet();
    this.settings = state.settings ?? { page: A4, defaultTab: 35.4 };
  }

  state(): DocState {
    return { paras: this.paras, styles: this.styles, settings: this.settings };
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
      this.starts = new Array(this.paras.length);
      let acc = 0;
      for (let i = 0; i < this.paras.length; i++) {
        this.starts[i] = acc;
        acc += this.paras[i].text.length + 1;
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
        else if (i >= a) out += ch === "\u000b" ? "\n" : ch === "\u000c" ? "" : ch;
      }
      return out;
    };
    if (from.p === to.p) return display(this.paras[from.p], from.o, to.o);
    const parts = [display(this.paras[from.p], from.o, this.paras[from.p].text.length)];
    for (let i = from.p + 1; i < to.p; i++) parts.push(display(this.paras[i], 0, this.paras[i].text.length));
    parts.push(display(this.paras[to.p], 0, to.o));
    return parts.join("\n");
  }

  /** Paragraphs covering [from, to) as pieces (for copying formatted content). */
  fragment(from: Pos, to: Pos): Para[] {
    if (from.p === to.p) return [slicePara(this.paras[from.p], from.o, to.o)];
    const out = [slicePara(this.paras[from.p], from.o, this.paras[from.p].text.length)];
    for (let i = from.p + 1; i < to.p; i++) out.push(this.paras[i]);
    out.push(slicePara(this.paras[to.p], 0, to.o));
    return out;
  }

  /**
   * Replaces [from, to) with `pieces`: the first piece joins the text before, the last
   * piece joins the text after, pieces in between become whole paragraphs. Returns the
   * position after the inserted content.
   */
  replace(from: Pos, to: Pos, pieces: Para[]): Pos {
    const head = slicePara(this.paras[from.p], 0, from.o);
    const tailSrc = this.paras[to.p];
    const tail = slicePara(tailSrc, to.o, tailSrc.text.length);
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
    this.paras = [...this.paras.slice(0, from.p), ...created, ...this.paras.slice(to.p + 1)];
    this.changed();
    const last = pieces.length ? pieces[pieces.length - 1] : textPiece("", EMPTY);
    return pieces.length <= 1 ? { p: from.p, o: from.o + last.text.length } : { p: from.p + pieces.length - 1, o: last.text.length };
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
      if (atEnd && text) this.paras[end.p] = { ...this.paras[end.p], mark: intern(props) };
      return end;
    }
    const src = this.paras[from.p];
    const atEnd = to.p === from.p ? to.o >= src.text.length : to.o >= this.paras[to.p].text.length;
    // Enter at the end of a paragraph continues with the style's "next" style (heading →
    // Standard), dropping direct paragraph formatting of the heading; elsewhere both halves
    // keep the paragraph's formatting.
    const style = this.styles.resolvePara(src.props).style;
    const next = this.styles.get(style)?.next;
    const nextProps: ParaProps = atEnd && next && next !== style ? { style: next } : src.props;
    // Inserted marks take the inserted formatting; the last paragraph keeps the original mark
    // (replace() joins it with the tail), unless text typed at its end formats it.
    const pieces: Para[] = lines.map((line, i) => ({
      ...textPiece(line, props),
      props: i === 0 ? src.props : nextProps,
      mark: intern(props),
    }));
    const end = this.replace(from, to, pieces);
    if (atEnd && lines[lines.length - 1]) this.paras[end.p] = { ...this.paras[end.p], mark: intern(props) };
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

  restore(state: { paras: Para[]; styles: StyleSheet }): void {
    this.paras = state.paras;
    this.styles = state.styles;
    this.changed();
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
