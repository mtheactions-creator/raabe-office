// The formula editor (D-039): editing an equation tree the way Word does, recorded with
// tools/feel-lab/catalogue_math.py. Positions are paths through the tree's slots; text is
// edited character by character. Word's rules:
// - Right/Left walk through every slot in reading order (a fraction: numerator, denominator;
//   scripts: base, subscript, superscript) and leave the equation at its ends; Up/Down move
//   between slots above and below each other.
// - Umschalt+arrow extends at the level of the selection: a structure is selected as a whole.
// - Backspace after a structure (or at the start of one of its slots) and Delete before one
//   select it first; the next press removes it. An equation emptied this way disappears.
// - Typing uses Word's linear format: an expression is built up when it ends (space or an
//   operator): a/b, x^2, x_1, √(x), ∑_(i=1)^n, (…), |…|, ■(a&b@c&d), function names, accents.
//   Math AutoCorrect turns \alpha, \cdot … and <=, ->, /= … into symbols.

import { ACCENTS, FUNCTIONS, LIMIT_FUNCTIONS, NARY, SYMBOLS, mergeText, parseLatex, type MNode, type Seq } from "./ast";

export type SlotKey = string;
export interface MathStep { readonly i: number; readonly slot: SlotKey }
export interface MathPos { readonly path: readonly MathStep[]; readonly o: number }
export interface MathRange { readonly path: readonly MathStep[]; readonly from: number; readonly to: number }

/** What an editing step asks of the host. */
export interface MathSignal {
  /** The caret leaves the equation (to the text before or after it). */
  readonly leave?: "before" | "after";
  /** The equation became empty and goes. */
  readonly remove?: boolean;
  /** Backspace at the very start (the host deletes before the equation). */
  readonly backspaceBefore?: boolean;
  readonly deleteAfter?: boolean;
}

type T = Extract<MNode, { k: "t" }>;
const isT = (n: MNode | undefined): n is T => n?.k === "t";

// ----- slots ---------------------------------------------------------------------------------

/** The slots of a structure in reading order (only those it has). */
export function slotsOf(n: MNode): SlotKey[] {
  switch (n.k) {
    case "t": return [];
    case "frac": return ["num", "den"];
    case "scr": {
      const s = [...(n.sub ? ["sub"] : []), ...(n.sup ? ["sup"] : [])];
      return n.pre ? [...s, "base"] : ["base", ...s];
    }
    case "rad": return n.deg ? ["deg", "e"] : ["e"];
    case "d": return n.items.map((_, i) => `item${i}`);
    case "nary": return [...(n.sub ? ["sub"] : []), ...(n.sup ? ["sup"] : []), "e"];
    case "func": return ["name", "e"];
    case "acc": case "bar": case "group": case "box": return ["e"];
    case "lim": return ["e", "lim"];
    case "m": return n.rows.flatMap((r, ri) => r.map((_, ci) => `cell${ri}_${ci}`));
    case "eqArr": return n.rows.map((_, i) => `row${i}`);
  }
}

export function getSlot(n: MNode, key: SlotKey): Seq {
  const any = n as unknown as Record<string, Seq>;
  if (key.startsWith("item") && n.k === "d") return n.items[Number(key.slice(4))] ?? [];
  if (key.startsWith("cell") && n.k === "m") { const [r, c] = key.slice(4).split("_").map(Number); return n.rows[r]?.[c] ?? []; }
  if (key.startsWith("row") && n.k === "eqArr") return n.rows[Number(key.slice(3))] ?? [];
  return any[key] ?? [];
}

function setSlot(n: MNode, key: SlotKey, seq: Seq): void {
  const any = n as unknown as Record<string, Seq>;
  if (key.startsWith("item") && n.k === "d") { n.items[Number(key.slice(4))] = seq; return; }
  if (key.startsWith("cell") && n.k === "m") { const [r, c] = key.slice(4).split("_").map(Number); n.rows[r][c] = seq; return; }
  if (key.startsWith("row") && n.k === "eqArr") { n.rows[Number(key.slice(3))] = seq; return; }
  any[key] = seq;
}

/** A deep copy with every text run split into single characters (positions count them). */
export function explode(seq: Seq): Seq {
  const out: Seq = [];
  for (const n of seq) {
    if (n.k === "t") { for (const ch of n.s) out.push({ ...n, s: ch }); continue; }
    const c = structuredClone(n) as MNode;
    for (const k of slotsOf(c)) setSlot(c, k, explode(getSlot(c, k)));
    out.push(c);
  }
  return out;
}

/** The canonical tree: neighbouring characters with the same style joined again. */
export function implode(seq: Seq): Seq {
  const out: Seq = [];
  for (const n of seq) {
    if (n.k === "t") { out.push(n); continue; }
    const c = { ...n } as MNode;
    if (c.k === "d") (c as { items: Seq[] }).items = [...c.items];
    if (c.k === "m") (c as { rows: Seq[][] }).rows = c.rows.map((r) => [...r]);
    if (c.k === "eqArr") (c as { rows: Seq[] }).rows = [...c.rows];
    for (const k of slotsOf(c)) setSlot(c, k, implode(getSlot(c, k)));
    out.push(c);
  }
  return mergeText(out);
}

// ----- characters ------------------------------------------------------------------------------

const ALNUM = /[\p{L}\p{N}]/u;
const LETTER = /\p{L}/u;
const BUILD_OPS = new Set(["^", "_", "/", "√", "∛", "∜", "■", "&", "@", "\\"]);
const OPEN: Record<string, string> = { "(": ")", "[": "]", "{": "}", "⟨": "⟩", "⌊": "⌋", "⌈": "⌉" };
const CLOSE = new Set(Object.values(OPEN));
/** Characters that end an operand (Word's operators). */
const OPERATORS = new Set(["+", "-", "−", "=", "<", ">", "≤", "≥", "≠", "±", "∓", "×", "⋅", "÷", ",", ";", ":", "→", "←", "⇒", "⇔", "≈", "≡", "∼", "≅", "∝", "∈", "∉", "⊂", "⊆", "∪", "∩", " ", "!", "?", "'", "§", "≪", "≫", "∗", "∘", "↔", "⟶", "∖"]);

/** Word's Math AutoCorrect for two or three typed characters. */
const PAIRS: Record<string, string> = {
  "<=": "≤", ">=": "≥", "/=": "≠", "!=": "≠", "->": "→", "<-": "←", "+-": "±", "-+": "∓", "~=": "≅", ":=": "≔",
  "<<": "≪", ">>": "≫", "...": "…", "::": "∷", "/<": "≮", "/>": "≯", "<->": "↔", "=>": "⇒",
};

/** Control words of Word's Math AutoCorrect beyond the LaTeX symbol names. */
const WORDS: Record<string, string> = {
  degree: "°", sqrt: "√", cbrt: "∛", qdrt: "∜", matrix: "■", inc: "∆", infty: "∞", deg: "°", cdots: "⋯", ldots: "…",
  hbar: "ℏ", ohm: "Ω", angstrom: "Å", euro: "€", celsius: "℃", box: "□", rect: "▭", dd: "ⅆ", ee: "ⅇ", ii: "ⅈ",
};

/** Word's Math AutoCorrect as a list (Formeloptionen > Math. AutoKorrektur): what is typed, what it becomes. */
export function mathAutoCorrect(): [string, string][] {
  const out = new Map<string, string>();
  for (const [k, v] of Object.entries(PAIRS)) out.set(k, v);
  for (const table of [SYMBOLS, NARY, ACCENTS, WORDS]) for (const [k, v] of Object.entries(table)) out.set(`\\${k}`, v);
  return [...out].sort((a, b) => a[0].localeCompare(b[0], "de"));
}

function symbolFor(word: string): string | null {
  if (NARY[word]) return NARY[word];
  if (WORDS[word]) return WORDS[word];
  if (ACCENTS[word]) return ACCENTS[word];
  if (SYMBOLS[word]) return SYMBOLS[word];
  return null;
}

const NARY_CHARS = new Set(Object.values(NARY));
const ACCENT_CHARS = new Set(Object.values(ACCENTS));

// ----- the editor ---------------------------------------------------------------------------------

export class MathEditor {
  root: Seq;
  anchor: MathPos;
  head: MathPos;
  /** Formel > Text: typed characters are plain text, nothing is built up. */
  normal = false;
  /** Formeloptionen: expressions are built up while typing (else only by Konvertieren). */
  autoBuild = true;
  /** Formeloptionen: limits of new integrals and other large operators above and below. */
  limits = { int: false, nary: true };
  /** Formel > LaTeX: typed LaTeX is built up on a space (\frac{a}{b}, x^{2}). */
  latex = false;

  constructor(seq: Seq, at: "start" | "end" = "end") {
    this.root = explode(seq);
    const o = at === "start" ? 0 : this.root.length;
    this.anchor = this.head = { path: [], o };
  }

  /** The equation as a canonical tree. */
  value(): Seq {
    return implode(this.root);
  }

  isEmpty(): boolean {
    return this.root.length === 0;
  }

  seqAt(path: readonly MathStep[]): Seq {
    let s = this.root;
    for (const st of path) s = getSlot(s[st.i], st.slot);
    return s;
  }

  /** The structure a path step points into. */
  private nodeAt(path: readonly MathStep[]): MNode | null {
    if (!path.length) return null;
    return this.seqAt(path.slice(0, -1))[path[path.length - 1].i] ?? null;
  }

  get hasSelection(): boolean {
    const r = this.range();
    return r.from !== r.to;
  }

  /** The selection, lifted to the sequence both ends are in (structures selected whole). */
  range(): MathRange {
    const a = this.anchor, h = this.head;
    let c = 0;
    while (c < a.path.length && c < h.path.length && a.path[c].i === h.path[c].i && a.path[c].slot === h.path[c].slot) c++;
    const path = a.path.slice(0, c);
    const at = (p: MathPos): [number, boolean] => (p.path.length > c ? [p.path[c].i, true] : [p.o, false]);
    const [ai, ain] = at(a), [hi, hin] = at(h);
    if (!ain && !hin) return { path, from: Math.min(ai, hi), to: Math.max(ai, hi) };
    // An end inside a structure takes the whole structure.
    const lo = Math.min(ai, hi), hiI = Math.max(ai, hi);
    const loInside = ai < hi ? ain : ai > hi ? hin : ain || hin;
    const hiInside = ai > hi ? ain : ai < hi ? hin : ain || hin;
    return { path, from: lo, to: hiI + (hiInside ? 1 : 0) - (loInside ? 0 : 0) };
  }

  private collapse(p: MathPos): void {
    this.anchor = this.head = p;
  }

  private select(path: readonly MathStep[], from: number, to: number): void {
    this.anchor = { path, o: from };
    this.head = { path, o: to };
  }

  // ----- movement ----------------------------------------------------------------------------

  private stepRight(p: MathPos, enter: boolean): MathPos | null {
    const s = this.seqAt(p.path);
    if (p.o < s.length) {
      const a = s[p.o];
      const slots = slotsOf(a);
      if (enter && slots.length) return { path: [...p.path, { i: p.o, slot: slots[0] }], o: 0 };
      return { path: p.path, o: p.o + 1 };
    }
    if (!p.path.length) return null;
    const last = p.path[p.path.length - 1], parent = p.path.slice(0, -1);
    const slots = slotsOf(this.seqAt(parent)[last.i]);
    const k = slots.indexOf(last.slot);
    if (enter && k + 1 < slots.length) return { path: [...parent, { i: last.i, slot: slots[k + 1] }], o: 0 };
    return { path: parent, o: last.i + 1 };
  }

  private stepLeft(p: MathPos, enter: boolean): MathPos | null {
    if (p.o > 0) {
      const a = this.seqAt(p.path)[p.o - 1];
      const slots = slotsOf(a);
      if (enter && slots.length) {
        const slot = slots[slots.length - 1];
        return { path: [...p.path, { i: p.o - 1, slot }], o: getSlot(a, slot).length };
      }
      return { path: p.path, o: p.o - 1 };
    }
    if (!p.path.length) return null;
    const last = p.path[p.path.length - 1], parent = p.path.slice(0, -1);
    const node = this.seqAt(parent)[last.i];
    const slots = slotsOf(node);
    const k = slots.indexOf(last.slot);
    if (enter && k > 0) return { path: [...parent, { i: last.i, slot: slots[k - 1] }], o: getSlot(node, slots[k - 1]).length };
    return { path: parent, o: last.i };
  }

  /** Right (Umschalt: extend). Leaving the equation is signalled. */
  right(shift = false): MathSignal {
    if (shift) return this.extend(1);
    if (this.hasSelection) { const r = this.range(); this.collapse({ path: r.path, o: r.to }); return {}; }
    const n = this.stepRight(this.head, true);
    if (!n) return { leave: "after" };
    this.collapse(n);
    return {};
  }

  left(shift = false): MathSignal {
    if (shift) return this.extend(-1);
    if (this.hasSelection) { const r = this.range(); this.collapse({ path: r.path, o: r.from }); return {}; }
    const n = this.stepLeft(this.head, true);
    if (!n) return { leave: "before" };
    this.collapse(n);
    return {};
  }

  /** Umschalt+arrow: the head moves at the selection's level, over structures as a whole. */
  private extend(dir: 1 | -1): MathSignal {
    const r = this.range();
    let head = this.head;
    if (head.path.length > r.path.length) {
      const atEnd = this.compare(this.head, this.anchor) >= 0;
      head = { path: r.path, o: atEnd ? r.to : r.from };
      this.anchor = { path: r.path, o: atEnd ? r.from : r.to };
    }
    const n = dir > 0 ? this.stepRight(head, false) : this.stepLeft(head, false);
    if (n) this.head = n;
    return {};
  }

  /** Order of two positions in reading order (-1, 0, 1). */
  compare(a: MathPos, b: MathPos): number {
    const n = Math.min(a.path.length, b.path.length);
    for (let k = 0; k < n; k++) {
      if (a.path[k].i !== b.path[k].i) return a.path[k].i < b.path[k].i ? -1 : 1;
      if (a.path[k].slot !== b.path[k].slot) {
        const slots = slotsOf(this.nodeAt(a.path.slice(0, k + 1))!);
        return slots.indexOf(a.path[k].slot) < slots.indexOf(b.path[k].slot) ? -1 : 1;
      }
    }
    const ai = a.path.length > n ? a.path[n].i : a.o, bi = b.path.length > n ? b.path[n].i : b.o;
    if (ai !== bi) return ai < bi ? -1 : 1;
    if (a.path.length > n) return 1; // inside the structure after the gap
    if (b.path.length > n) return -1;
    return 0;
  }

  home(shift = false): void {
    const p = { path: [], o: 0 };
    if (shift) { this.lift(); this.head = p; } else this.collapse(p);
  }

  end(shift = false): void {
    const p = { path: [], o: this.root.length };
    if (shift) { this.lift(); this.head = p; } else this.collapse(p);
  }

  /** The anchor at the top level, for a selection that reaches it (Umschalt+Pos1/Ende). */
  private lift(): void {
    let a = this.anchor;
    while (a.path.length) a = { path: a.path.slice(0, -1), o: a.path[a.path.length - 1].i + (this.compare(this.anchor, this.head) > 0 ? 1 : 0) };
    this.anchor = a;
  }

  /** Strg+Right / Strg+Left: over a run of letters and digits, or one other atom. */
  word(dir: 1 | -1, shift = false): MathSignal {
    const s = this.seqAt(this.head.path);
    let o = this.head.o;
    const alnum = (n: MNode | undefined) => isT(n) && ALNUM.test(n.s);
    if (dir > 0) {
      if (o >= s.length) return shift ? this.extend(1) : this.right();
      if (alnum(s[o])) while (o < s.length && alnum(s[o])) o++; else o++;
    } else {
      if (o <= 0) return shift ? this.extend(-1) : this.left();
      // Spaces before the caret are passed over first, then the word (Word: ab+cd |→ ab+|cd).
      const space = (n: MNode | undefined) => isT(n) && n.s === " ";
      while (o > 1 && space(s[o - 1])) o--;
      if (alnum(s[o - 1])) while (o > 0 && alnum(s[o - 1])) o--; else o--;
    }
    const p = { path: this.head.path, o };
    if (shift) this.head = p; else this.collapse(p);
    return {};
  }

  /**
   * Up/Down: to the slot above or below (numerator/denominator, scripts, limits, matrix rows,
   * lines); `pick` chooses the offset in the target slot (by x, from the host's layout).
   */
  vertical(dir: 1 | -1, pick?: (path: readonly MathStep[], seqLength: number) => number): MathSignal {
    let path = this.head.path;
    while (path.length) {
      const last = path[path.length - 1], parent = path.slice(0, -1);
      const node = this.seqAt(parent)[last.i];
      const target = verticalSlot(node, last.slot, dir);
      if (target) {
        const tp = [...parent, { i: last.i, slot: target }];
        const len = getSlot(node, target).length;
        const cur = this.seqAt(path).length;
        const o = pick ? pick(tp, len) : this.head.o >= cur ? len : Math.min(this.head.o, len);
        this.collapse({ path: tp, o });
        return {};
      }
      path = parent;
    }
    return {};
  }

  // ----- deleting -------------------------------------------------------------------------------

  private removeRange(r: MathRange): void {
    const s = this.seqAt(r.path);
    s.splice(r.from, r.to - r.from);
    this.collapse({ path: r.path, o: r.from });
  }

  /** Backspace (Rücktaste), as Word does it. */
  backspace(): MathSignal {
    if (this.hasSelection) {
      this.removeRange(this.range());
      return this.root.length ? {} : { remove: true };
    }
    const p = this.head;
    const s = this.seqAt(p.path);
    if (p.o > 0) {
      const prev = s[p.o - 1];
      if (isT(prev)) {
        s.splice(p.o - 1, 1);
        this.collapse({ path: p.path, o: p.o - 1 });
        return this.root.length ? {} : { remove: true };
      }
      // A structure is selected first.
      this.select(p.path, p.o - 1, p.o);
      return {};
    }
    if (!p.path.length) return this.root.length ? { backspaceBefore: true } : { remove: true };
    // At the start of a slot: the structure is selected.
    const last = p.path[p.path.length - 1];
    this.select(p.path.slice(0, -1), last.i, last.i + 1);
    return {};
  }

  /** Entf, as Word does it. */
  del(): MathSignal {
    if (this.hasSelection) {
      this.removeRange(this.range());
      return this.root.length ? {} : { remove: true };
    }
    const p = this.head;
    const s = this.seqAt(p.path);
    if (p.o < s.length) {
      const next = s[p.o];
      if (isT(next)) {
        s.splice(p.o, 1);
        return this.root.length ? {} : { remove: true };
      }
      this.select(p.path, p.o, p.o + 1);
      return {};
    }
    if (!p.path.length) return this.root.length ? { deleteAfter: true } : { remove: true };
    const last = p.path[p.path.length - 1];
    this.select(p.path.slice(0, -1), last.i, last.i + 1);
    return {};
  }

  selectAll(): void {
    this.select([], 0, this.root.length);
  }

  // ----- typing ---------------------------------------------------------------------------------

  /** Inserts atoms at the caret (replacing the selection). */
  insertAtoms(atoms: Seq): void {
    if (this.hasSelection) this.removeRange(this.range());
    const s = this.seqAt(this.head.path);
    s.splice(this.head.o, 0, ...atoms);
    this.collapse({ path: this.head.path, o: this.head.o + atoms.length });
  }

  /**
   * Inserts a structure (Formel tab): the selection goes into its first slot (as Word does
   * for brackets, roots …), the caret into the first empty slot.
   */
  insertStructure(node: MNode, wrapSelection = true): void {
    const n = explode([node])[0];
    let moved: Seq = [];
    if (this.hasSelection) {
      const r = this.range();
      moved = this.seqAt(r.path).slice(r.from, r.to);
      this.removeRange(r);
    }
    const slots = slotsOf(n);
    if (wrapSelection && moved.length && slots.length) {
      const target = n.k === "scr" && !n.pre ? "base" : slots[n.k === "nary" || n.k === "func" ? slots.length - 1 : 0];
      setSlot(n, target, [...moved, ...getSlot(n, target)]);
    }
    const s = this.seqAt(this.head.path);
    const at = this.head.o;
    s.splice(at, 0, n);
    const empty = slots.find((k) => getSlot(n, k).length === 0);
    if (empty) this.collapse({ path: [...this.head.path, { i: at, slot: empty }], o: 0 });
    else this.collapse({ path: this.head.path, o: at + 1 });
  }

  /** Where a quoted text was just closed (a space typed there is used up, as in Word). */
  private quoteClosed: MathPos | null = null;

  /** Types a character with Word's Math AutoCorrect and build-up. */
  type(ch: string): MathSignal {
    if (this.hasSelection) this.removeRange(this.range());
    if (this.normal) { this.insertAtoms([{ k: "t", s: ch, nor: true }]); return {}; }
    if (this.latex) return this.typeLatex(ch);
    // Quoted text ("wenn") becomes normal text when the closing quote is typed; inside the
    // quotes nothing is built up or corrected; a space right after it is used up; typing
    // inside or at the end of normal text continues it (Word, recorded 2026-10-15).
    {
      const qs = this.seqAt(this.head.path), qo = this.head.o;
      const closed = this.quoteClosed && samePos(this.quoteClosed, this.head) ? this.quoteClosed : null;
      this.quoteClosed = null;
      const open = openQuote(qs, qo);
      if (open >= 0) {
        if (ch !== '"') { this.insertAtoms([{ k: "t", s: ch }]); return {}; }
        const inner: Seq = qs.slice(open + 1, qo).map((n) => ({ k: "t", s: (n as T).s, nor: true }));
        qs.splice(open, qo - open, ...inner);
        this.collapse({ path: this.head.path, o: open + inner.length });
        this.quoteClosed = { path: this.head.path, o: open + inner.length };
        return {};
      }
      if (closed && ch === " ") { this.quoteClosed = closed; return {}; }
      const left = qs[qo - 1], right = qs[qo];
      if (isT(left) && left.nor && ((isT(right) && right.nor) || !closed)) { this.insertAtoms([{ k: "t", s: ch, nor: true }]); return {}; }
    }
    const s = this.seqAt(this.head.path);
    const o = this.head.o;
    const prev = s[o - 1];
    // 1. Two or three characters that are one symbol (<= → ≤ …).
    for (const len of [2, 1]) {
      const before = s.slice(o - len, o);
      if (before.length !== len || !before.every((n) => isT(n) && !n.nor)) continue;
      const key = before.map((n) => (n as T).s).join("") + ch;
      if (PAIRS[key]) {
        s.splice(o - len, len, { k: "t", s: PAIRS[key] });
        this.collapse({ path: this.head.path, o: o - len + 1 });
        return {};
      }
    }
    // 2. A control word ends (\alpha, \int, \bar …).
    if (!(LETTER.test(ch) && /^[A-Za-z]$/.test(ch))) {
      let k = o;
      while (k > 0 && isT(s[k - 1]) && /^[A-Za-z]$/.test((s[k - 1] as T).s)) k--;
      if (k > 0 && isT(s[k - 1]) && (s[k - 1] as T).s === "\\" && k < o) {
        const word = s.slice(k, o).map((n) => (n as T).s).join("");
        const sym = symbolFor(word);
        if (sym) {
          s.splice(k - 1, o - k + 1, /^[Α-Ω]$/.test(sym) ? { k: "t", s: sym, sty: "p" } : { k: "t", s: sym });
          this.collapse({ path: this.head.path, o: k });
          // An accent applies to what is before it at once (x\bar → x̄).
          if (ACCENT_CHARS.has(sym)) this.buildAccent(sym);
          if (ch === " ") {
            // A large operator, root or function built from the space.
            if (NARY_CHARS.has(sym)) this.buildUp(" ");
            return {};
          }
        }
      }
    }
    // 3. Build-up when an expression ends.
    if (!this.autoBuild) { this.insertAtoms([{ k: "t", s: ch }]); return {}; }
    const trigger = ch === " " || ch === ")" || ch === "]" || ch === "}" || (!ALNUM.test(ch) && !BUILD_OPS.has(ch) && !OPEN[ch] && ch !== "|" && !ACCENT_CHARS.has(ch));
    if (trigger && !(prev && isT(prev) && BUILD_OPS.has(prev.s) && ch !== " ")) {
      const built = this.buildUp(ch);
      if (built && ch === " ") return {};
    }
    // Closing brackets build what is inside them first.
    this.insertAtoms([{ k: "t", s: ch }]);
    return {};
  }

  /** LaTeX mode: the LaTeX before the caret is built up when a space ends it. */
  private typeLatex(ch: string): MathSignal {
    const s = this.seqAt(this.head.path);
    const o = this.head.o;
    if (ch === " ") {
      let k = o;
      while (k > 0 && isT(s[k - 1]) && (s[k - 1] as T).s !== " " && !(s[k - 1] as T).nor) k--;
      const src = s.slice(k, o).map((n) => (n as T).s).join("");
      let depth = 0;
      for (const c of src) depth += c === "{" ? 1 : c === "}" ? -1 : 0;
      if (src && /[\\^_{]/.test(src) && depth === 0) {
        const built = parseLatex(src);
        if (built.length && !(built.length === 1 && isT(built[0]) && built[0].s === src)) {
          const atoms = explode(built);
          s.splice(k, o - k, ...atoms);
          this.collapse({ path: this.head.path, o: k + atoms.length });
          return {};
        }
      }
    }
    this.insertAtoms([{ k: "t", s: ch }]);
    return {};
  }

  /** x then a combining accent → the accent over x. */
  private buildAccent(chr: string): void {
    const s = this.seqAt(this.head.path);
    const o = this.head.o; // after the accent character
    const i = o - 1;
    if (i < 1) return;
    const base = s[i - 1];
    s.splice(i - 1, 2, { k: "acc", chr, e: [base] });
    this.collapse({ path: this.head.path, o: i });
  }

  /**
   * Builds up the linear expression that ends at the caret (Word's linear format). Returns
   * whether anything was built.
   */
  private buildUp(trigger: string): boolean {
    const path = this.head.path;
    const s = this.seqAt(path);
    const end = this.head.o;
    const start = expressionStart(s, end, trigger);
    const atoms = s.slice(start, end);
    if (!atoms.length) return false;
    LIMITS = this.limits;
    const res = build(atoms, trigger);
    if (!res) return false;
    s.splice(start, end - start, ...res.seq);
    // The caret goes into the argument of a large operator or function, else after the result.
    if (res.into !== undefined) {
      const node = s[start + res.into];
      const slot = "e";
      this.collapse({ path: [...path, { i: start + res.into, slot }], o: getSlot(node, slot).length });
    } else this.collapse({ path, o: start + res.seq.length });
    return true;
  }

  /** Umschalt+Eingabe: a new line in the equation (the host leaves on Eingabe). */
  newLine(): void {
    if (this.head.path.length) this.collapse({ path: [], o: this.head.path[0].i + 1 });
    this.insertAtoms([{ k: "t", s: "\n" }]);
  }
}

/** The slot above (dir -1) or below (1) another in the same structure. */
function verticalSlot(n: MNode, slot: SlotKey, dir: 1 | -1): SlotKey | null {
  switch (n.k) {
    case "frac": return dir < 0 && slot === "den" ? "num" : dir > 0 && slot === "num" ? "den" : null;
    case "scr":
      // Recorded in Word: Up from the subscript goes to the base, Down from the superscript too.
      if (slot === "base") return dir < 0 ? (n.sup ? "sup" : null) : n.sub ? "sub" : null;
      if (slot === "sup" && dir > 0) return "base";
      if (slot === "sub" && dir < 0) return "base";
      return null;
    case "nary":
      // Up and Down in the operand stay where they are (Word); the limits reach each other.
      if (slot === "sup" && dir > 0) return n.sub ? "sub" : null;
      if (slot === "sub" && dir < 0) return n.sup ? "sup" : null;
      return null;
    case "lim": return n.upper ? (dir < 0 && slot === "e" ? null : null) : dir > 0 && slot === "e" ? "lim" : dir < 0 && slot === "lim" ? "e" : null;
    case "rad": return slot === "e" && dir < 0 && n.deg ? "deg" : slot === "deg" && dir > 0 ? "e" : null;
    case "m": {
      const [r, c] = slot.slice(4).split("_").map(Number);
      const nr = r + dir;
      return nr >= 0 && nr < n.rows.length ? `cell${nr}_${Math.min(c, n.rows[nr].length - 1)}` : null;
    }
    case "eqArr": {
      const r = Number(slot.slice(3)) + dir;
      return r >= 0 && r < n.rows.length ? `row${r}` : null;
    }
    default: return null;
  }
}

// ----- building up the linear format -----------------------------------------------------------

/** The unmatched opening quote before `o` in a sequence (its index), or -1. */
function openQuote(s: Seq, o: number): number {
  let at = -1;
  for (let i = 0; i < o; i++) {
    const n = s[i];
    if (isT(n) && !n.nor && n.s === '"') at = at < 0 ? i : -1;
  }
  return at;
}

const samePos = (a: MathPos, b: MathPos) => a.o === b.o && a.path.length === b.path.length && a.path.every((st, k) => st.i === b.path[k].i && st.slot === b.path[k].slot);

/** Where new large operators put their limits (the editor's Formeloptionen while it builds). */
let LIMITS = { int: false, nary: true };

/** Where the expression ending at `end` starts: after the last operator outside brackets. */
function expressionStart(s: Seq, end: number, trigger: string): number {
  let depth = 0;
  let i = end - 1;
  // A closing bracket typed now ends the expression inside its bracket.
  const inside = trigger === ")" || trigger === "]" || trigger === "}";
  for (; i >= 0; i--) {
    const n = s[i];
    if (!isT(n)) continue;
    const c = n.s;
    if (CLOSE.has(c)) { depth++; continue; }
    if (OPEN[c]) {
      if (depth === 0) break;
      depth--;
      continue;
    }
    if (depth > 0) continue;
    if (c === "\n") break;
    if (OPERATORS.has(c)) {
      // A minus right after ^, _ or / belongs to the operand (e^-x).
      const before = s[i - 1];
      // A decimal comma or point between digits is part of the number (0,5/2).
      const after = s[i + 1];
      if ((c === "," || c === ".") && isT(before) && /[0-9]/.test(before.s) && isT(after) && /[0-9]/.test(after.s)) continue;
      if ((c === "-" || c === "−") && (i === 0 || (isT(before) && BUILD_OPS.has(before.s)))) continue;
      break;
    }
  }
  void inside;
  return i + 1;
}

interface Unit { seq: Seq; bracket?: { open: string; close: string; inner: Seq }; letters?: boolean; digits?: boolean }

/** Splits atoms into units: letter runs, numbers, bracket groups, structures, other characters. */
function units(atoms: Seq): Unit[] | null {
  const out: Unit[] = [];
  for (let i = 0; i < atoms.length; i++) {
    const n = atoms[i];
    if (!isT(n)) { out.push({ seq: [n] }); continue; }
    const c = n.s;
    if (OPEN[c] || c === "|") {
      const close = c === "|" ? "|" : OPEN[c];
      let depth = 0, j = i + 1;
      for (; j < atoms.length; j++) {
        const m = atoms[j];
        if (!isT(m)) continue;
        if (m.s === close && depth === 0) break;
        if (c !== "|" && m.s === c) depth++;
        else if (c !== "|" && m.s === close) depth--;
      }
      if (j >= atoms.length) { out.push({ seq: [n] }); continue; }
      out.push({ seq: atoms.slice(i, j + 1), bracket: { open: c, close, inner: atoms.slice(i + 1, j) } });
      i = j;
      continue;
    }
    if (/[0-9]/.test(c)) {
      let j = i;
      while (j + 1 < atoms.length && isT(atoms[j + 1]) && /[0-9.,]/.test((atoms[j + 1] as T).s) && !(/[.,]/.test((atoms[j + 1] as T).s) && !(isT(atoms[j + 2]) && /[0-9]/.test((atoms[j + 2] as T).s)))) j++;
      out.push({ seq: atoms.slice(i, j + 1), digits: true });
      i = j;
      continue;
    }
    if (LETTER.test(c) && !n.nor) {
      let j = i;
      while (j + 1 < atoms.length && isT(atoms[j + 1]) && LETTER.test((atoms[j + 1] as T).s) && !(atoms[j + 1] as T).nor) j++;
      out.push({ seq: atoms.slice(i, j + 1), letters: true });
      i = j;
      continue;
    }
    out.push({ seq: [n] });
  }
  return out;
}

const charOf = (u: Unit | undefined): string | null => (u && u.seq.length === 1 && isT(u.seq[0]) && !u.bracket ? (u.seq[0] as T).s : null);

/** An operand as a script or fraction part: brackets around it go (x^(2n) → 2n). */
const operandSeq = (u: Unit): Seq => (u.bracket && (u.bracket.open === "(" || u.bracket.open === "[" || u.bracket.open === "{") ? buildAll(u.bracket.inner) : u.bracket ? [delim(u)] : u.seq);

function delim(u: Unit): MNode {
  const b = u.bracket!;
  const items: Seq[] = [[]];
  for (const n of buildAll(b.inner)) {
    if (isT(n) && n.s === "│") items.push([]);
    else items[items.length - 1].push(n);
  }
  return { k: "d", beg: b.open, end: b.close, items };
}

/** Builds everything in a finished part (inside brackets, operands). */
function buildAll(atoms: Seq): Seq {
  return build(atoms, " ")?.seq ?? atoms;
}

/**
 * Builds a linear expression: prefix operators (√, ■, large operators), scripts, fractions
 * and bracket groups. Returns the result and, for a large operator or function, which atom
 * the caret goes into; null when nothing is to be built.
 */
function build(atoms: Seq, trigger: string): { seq: Seq; into?: number } | null {
  const us = units(atoms);
  if (!us) return null;
  let changed = false;
  // Function names ("sin" + space): the function with an empty argument.
  const lastU = us[us.length - 1];
  if (trigger === " " && us.length === 1 && lastU.letters) {
    const name = lastU.seq.map((n) => (n as T).s).join("");
    if (FUNCTIONS.has(name) || LIMIT_FUNCTIONS.has(name)) return { seq: [{ k: "func", name: [{ k: "t", s: name, sty: "p" }], e: [] }], into: 0 };
  }
  if (trigger === " " && us.length === 3 && us[0].letters && (charOf(us[1]) === "_" || charOf(us[1]) === "^")) {
    const name = us[0].seq.map((n) => (n as T).s).join("");
    if (LIMIT_FUNCTIONS.has(name)) {
      const lim: MNode = { k: "lim", upper: charOf(us[1]) === "^", e: [{ k: "t", s: name, sty: "p" }], lim: operandSeq(us[2]) };
      return { seq: [{ k: "func", name: [lim], e: [] }], into: 0 };
    }
  }
  // Pass 1: prefix operators and their operands.
  const p1: (Unit | MNode)[] = [];
  let into: number | undefined;
  for (let i = 0; i < us.length; i++) {
    const c = charOf(us[i]);
    if (c === "√" || c === "∛" || c === "∜") {
      const next = us[i + 1];
      if (!next) { p1.push(us[i]); continue; }
      let deg = c === "∛" ? [{ k: "t", s: "3" } as MNode] : c === "∜" ? [{ k: "t", s: "4" } as MNode] : undefined;
      let e = operandSeq(next);
      const amp = c === "√" && next.bracket?.open === "(" ? next.bracket.inner.findIndex((n) => isT(n) && n.s === "&") : -1;
      if (amp >= 0) {
        deg = buildAll(next.bracket!.inner.slice(0, amp));
        e = buildAll(next.bracket!.inner.slice(amp + 1));
      }
      p1.push({ k: "rad", deg, e });
      i++;
      changed = true;
      continue;
    }
    if (c === "■") {
      const next = us[i + 1];
      if (!next?.bracket) { p1.push(us[i]); continue; }
      const rows: Seq[][] = [[[]]];
      for (const n of next.bracket.inner) {
        if (isT(n) && n.s === "@") rows.push([[]]);
        else if (isT(n) && n.s === "&") rows[rows.length - 1].push([]);
        else rows[rows.length - 1][rows[rows.length - 1].length - 1].push(n);
      }
      p1.push({ k: "m", rows: rows.map((r) => r.map(buildAll)) });
      i++;
      changed = true;
      continue;
    }
    if (c && NARY_CHARS.has(c)) {
      // ∑_(i=1)^n: the limits follow.
      let sub: Seq | undefined, sup: Seq | undefined;
      let j = i + 1;
      while (j + 1 < us.length + 1) {
        const op = charOf(us[j]);
        if ((op === "_" && !sub) || (op === "^" && !sup)) {
          const arg = us[j + 1];
          if (!arg) return null;
          if (op === "_") sub = operandSeq(arg); else sup = operandSeq(arg);
          j += 2;
          continue;
        }
        break;
      }
      // ▒ and an operand: the linear format's complete large operator.
      let e: Seq = [];
      const closed = charOf(us[j]) === "▒" && !!us[j + 1];
      if (closed) { e = operandSeq(us[j + 1]); j += 2; }
      if (j < us.length && !closed) { p1.push(us[i]); continue; }
      p1.push({ k: "nary", chr: c, sub, sup, e, und: /[∫∬∭∮∯∰∱∲∳]/.test(c) ? LIMITS.int : LIMITS.nary });
      if (!closed) into = p1.length - 1;
      i = j - 1;
      changed = true;
      continue;
    }
    // sin⁡x, lim┬(n→∞)⁡x: a function (name before the ⁡) applied to its argument.
    if (c === "⁡" && p1.length && us[i + 1]) {
      const prev = p1.pop()!;
      const name: Seq = "k" in prev ? [prev] : prev.letters ? [{ k: "t", s: prev.seq.map((n) => (n as T).s).join(""), sty: "p" }] : prev.seq;
      p1.push({ k: "func", name, e: operandSeq(us[i + 1]) });
      i += 1;
      changed = true;
      continue;
    }
    // lim┬(n→∞): a limit below (┴ above).
    if ((charOf(us[i + 1]) === "┬" || charOf(us[i + 1]) === "┴") && us[i + 2]) {
      const e = us[i].letters ? [{ ...(us[i].seq[0] as T), s: us[i].seq.map((n) => (n as T).s).join(""), sty: "p" as const }] : operandSeq(us[i]);
      p1.push({ k: "lim", upper: charOf(us[i + 1]) === "┴", e, lim: operandSeq(us[i + 2]) });
      i += 2;
      changed = true;
      continue;
    }
    p1.push(us[i]);
  }
  // Pass 2: scripts (^ and _) on the operand before them.
  const p2: (Unit | MNode)[] = [];
  const asUnit = (x: Unit | MNode): Unit => ("k" in x ? { seq: [x] } : x);
  for (let i = 0; i < p1.length; i++) {
    const x = p1[i];
    const op = "k" in x ? null : charOf(x);
    if ((op === "^" || op === "_") && p2.length) {
      let sub: Seq | undefined, sup: Seq | undefined;
      let j = i;
      for (;;) {
        const o = "k" in (p1[j] ?? {}) ? null : charOf(p1[j] as Unit);
        if ((o === "^" && !sup) || (o === "_" && !sub)) {
          let arg = p1[j + 1];
          if (!arg) return null;
          // A minus before the operand belongs to it (e^-x).
          let neg: Seq = [];
          if (!("k" in arg) && (charOf(arg) === "-" || charOf(arg) === "−")) {
            neg = arg.seq;
            arg = p1[j + 2];
            if (!arg) return null;
            j++;
          }
          const s = [...neg, ...operandSeq(asUnit(arg))];
          if (o === "^") sup = s; else sub = s;
          j += 2;
          continue;
        }
        break;
      }
      const baseX = p2.pop()!;
      let base: Seq;
      if (!("k" in baseX) && baseX.letters && baseX.seq.length > 1) {
        // Only the last letter carries the script.
        p2.push({ seq: baseX.seq.slice(0, -1), letters: true });
        base = baseX.seq.slice(-1);
      } else base = "k" in baseX ? [baseX] : baseX.bracket ? [delim(baseX)] : operandSeq(baseX);
      p2.push({ k: "scr", base, sub, sup });
      i = j - 1;
      changed = true;
      continue;
    }
    p2.push(x);
  }
  // Pass 3: fractions (left to right).
  const p3: (Unit | MNode)[] = [];
  for (let i = 0; i < p2.length; i++) {
    const x = p2[i];
    // On a space Word builds a fraction with an empty part too ("/ ", "a/ ": recorded 2026-10-16).
    const open = trigger === " ";
    if (!("k" in x) && (charOf(x) === "/" || charOf(x) === "¦") && ((p3.length && p2[i + 1]) || open)) {
      // Word's numerator: everything since the last operator; the denominator: everything up
      // to the next one (2x/3, 0,5/2, y=3a/4b, 2(a+b)/c).
      const ends = (u: Unit | MNode) => { const c = "k" in u ? null : charOf(u); return c !== null && (OPERATORS.has(c) || c === "/" || c === "¦"); };
      let k = p3.length;
      while (k > 0 && !ends(p3[k - 1])) k--;
      let j = i + 1;
      while (j < p2.length && !ends(p2[j])) j++;
      if (!open && (k === p3.length || j === i + 1)) { p3.push(x); continue; }
      const part = (us: (Unit | MNode)[]): Seq => us.length === 1 ? ("k" in us[0] ? [us[0]] : operandSeq(us[0]))
        : us.flatMap((u) => ("k" in u ? [u] : u.bracket ? [delim(u)] : u.seq));
      const num = part(p3.splice(k));
      const den = part(p2.slice(i + 1, j));
      p3.push({ k: "frac", num, den, ...(charOf(x) === "¦" ? { type: "noBar" as const } : {}) });
      i = j - 1;
      changed = true;
      continue;
    }
    p3.push(x);
  }
  // Bracket groups left over become delimiters (they were ended, so their content is built).
  const out: Seq = [];
  let intoOut: number | undefined;
  p3.forEach((x, i) => {
    if (i === into || (x as MNode).k === "nary" && into !== undefined && x === p1[into]) intoOut = out.length;
    if ("k" in x) { out.push(x); return; }
    if (x.bracket && trigger !== ")" && trigger !== "]" && trigger !== "}") { out.push(delim(x)); changed = true; return; }
    out.push(...x.seq);
  });
  if (!changed) return null;
  // The large operator's index after the passes.
  if (into !== undefined) intoOut = out.findIndex((n) => n.k === "nary");
  return { seq: out, into: intoOut !== undefined && intoOut >= 0 && intoOut === out.length - 1 ? intoOut : undefined };
}

/** Formel > Konvertieren > Professionell: a linear-format text built up as a whole. */
export function buildLinear(text: string): Seq {
  const atoms: Seq = [...text].map((ch) => ({ k: "t", s: ch }));
  return implode(build(atoms, " ")?.seq ?? atoms);
}
