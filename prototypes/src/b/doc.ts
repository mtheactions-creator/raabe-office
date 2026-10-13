// Document model for prototype B: an array of immutable paragraphs. Edits replace
// paragraph objects, so the layout can cache line breaks per paragraph object.

import type { StyleId } from "../shared/geometry";

export interface Para {
  readonly text: string;
  readonly style: StyleId;
}

export interface Pos {
  p: number;
  o: number;
}

export const cmp = (a: Pos, b: Pos): number => a.p - b.p || a.o - b.o;
export const minPos = (a: Pos, b: Pos): Pos => (cmp(a, b) <= 0 ? a : b);
export const maxPos = (a: Pos, b: Pos): Pos => (cmp(a, b) <= 0 ? b : a);

const NEXT_STYLE: Record<StyleId, StyleId> = { normal: "normal", h1: "normal" };

export class Doc {
  paras: Para[];
  private starts: number[] | null = null;

  constructor(paras: Para[]) {
    this.paras = paras.length ? paras : [{ text: "", style: "normal" }];
  }

  end(): Pos {
    const last = this.paras.length - 1;
    return { p: last, o: this.paras[last].text.length };
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

  text(): string {
    return this.paras.map((p) => p.text).join("\n");
  }

  slice(from: Pos, to: Pos): string {
    if (from.p === to.p) return this.paras[from.p].text.slice(from.o, to.o);
    const parts = [this.paras[from.p].text.slice(from.o)];
    for (let i = from.p + 1; i < to.p; i++) parts.push(this.paras[i].text);
    parts.push(this.paras[to.p].text.slice(0, to.o));
    return parts.join("\n");
  }

  /** Replaces [from, to) with `text` ("\n" splits paragraphs). Returns the position after the insert. */
  replace(from: Pos, to: Pos, text: string): Pos {
    const first = this.paras[from.p];
    const prefix = first.text.slice(0, from.o);
    const suffix = this.paras[to.p].text.slice(to.o);
    const pieces = text.split("\n");
    const created: Para[] = [];
    if (pieces.length === 1) {
      created.push({ text: prefix + pieces[0] + suffix, style: first.style });
    } else {
      // Pressing Enter at the end of a heading continues with the heading's "next" style.
      const tailStyle = suffix.length === 0 ? NEXT_STYLE[first.style] : first.style;
      created.push({ text: prefix + pieces[0], style: first.style });
      for (let i = 1; i < pieces.length - 1; i++) created.push({ text: pieces[i], style: tailStyle });
      created.push({ text: pieces[pieces.length - 1] + suffix, style: tailStyle });
    }
    this.paras = [...this.paras.slice(0, from.p), ...created, ...this.paras.slice(to.p + 1)];
    this.starts = null;
    const lastPiece = pieces[pieces.length - 1];
    return pieces.length === 1
      ? { p: from.p, o: from.o + lastPiece.length }
      : { p: from.p + pieces.length - 1, o: lastPiece.length };
  }

  restore(paras: Para[]): void {
    this.paras = paras;
    this.starts = null;
  }
}
