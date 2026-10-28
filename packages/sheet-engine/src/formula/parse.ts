// Parser from tokens to a syntax tree. Precedence as in Excel, highest first: the reference
// operators ":" (range), " " (intersection) and "," (union, only inside parentheses), then
// negation, "%", "^", "* /", "+ -", "&", comparisons. "^" is left-associative and negation
// binds tighter than "^" (=-2^2 is 4), as in Excel.

import { MAX_COLS, MAX_ROWS } from "../address";
import type { ErrorCode, Scalar } from "../model";
import { err } from "../model";
import { lex, type Dialect, type RefTok, type Token } from "./lexer";

export type Node =
  | { t: "num"; v: number }
  | { t: "str"; v: string }
  | { t: "bool"; v: boolean }
  | { t: "err"; v: ErrorCode }
  | { t: "ref"; sheet?: string; sheet2?: string; r0: number; c0: number; r1: number; c1: number }
  | { t: "name"; name: string; sheet?: string }
  | { t: "fn"; name: string; args: Node[] }
  | { t: "bin"; op: string; a: Node; b: Node }
  | { t: "un"; op: "-" | "+" | "%" | "@" | "#"; a: Node }
  | { t: "arr"; rows: number; cols: number; v: Scalar[] }
  | { t: "paren"; a: Node }
  | { t: "miss" }
  | { t: "struct"; text: string };

export class FormulaError extends Error {}

/** The rectangle of a reference token (whole columns and rows spread to the sheet's edges). */
export function refRect(r: RefTok): { r0: number; c0: number; r1: number; c1: number } {
  const b = r.b ?? r.a;
  const r0a = r.a.row ?? 0, r1a = r.a.row ?? MAX_ROWS - 1;
  const r0b = b.row ?? 0, r1b = b.row ?? MAX_ROWS - 1;
  const c0a = r.a.col ?? 0, c1a = r.a.col ?? MAX_COLS - 1;
  const c0b = b.col ?? 0, c1b = b.col ?? MAX_COLS - 1;
  return { r0: Math.min(r0a, r0b), r1: Math.max(r1a, r1b), c0: Math.min(c0a, c0b), c1: Math.max(c1a, c1b) };
}

class Parser {
  private i = 0;
  /** Source offsets of the nodes (Formelauswertung), when asked for. */
  spans: Map<Node, [number, number]> | null = null;
  constructor(private toks: Token[]) {}

  /** Offset of the next significant token. */
  private here(): number {
    let j = this.i;
    while (j < this.toks.length && this.toks[j].k === "ws") j++;
    return this.toks[j]?.s ?? 0;
  }

  private spanned(start: number, n: Node): Node {
    if (this.spans && !this.spans.has(n)) this.spans.set(n, [start, this.toks[this.i - 1]?.e ?? start]);
    return n;
  }

  /** Next significant token (skipping whitespace). */
  peek(): Token | undefined {
    let j = this.i;
    while (j < this.toks.length && this.toks[j].k === "ws") j++;
    return this.toks[j];
  }

  next(): Token | undefined {
    while (this.i < this.toks.length && this.toks[this.i].k === "ws") this.i++;
    return this.toks[this.i++];
  }

  isOp(t: Token | undefined, ...ops: string[]): boolean {
    return !!t && t.k === "op" && ops.includes(t.text);
  }

  expect(k: string): Token {
    const t = this.next();
    if (!t || t.k !== k) throw new FormulaError(`expected ${k}`);
    return t;
  }

  parseFormula(): Node {
    if (!this.peek()) throw new FormulaError("empty");
    const n = this.expr(false);
    if (this.peek()) throw new FormulaError(`unexpected ${this.peek()!.text}`);
    return n;
  }

  expr(union: boolean): Node {
    return this.comparison(union);
  }

  comparison(u: boolean): Node {
    let a = this.concat(u);
    for (;;) {
      const t = this.peek();
      if (!this.isOp(t, "=", "<>", "<", ">", "<=", ">=")) return a;
      this.next();
      a = { t: "bin", op: t!.text, a, b: this.concat(u) };
    }
  }

  concat(u: boolean): Node {
    let a = this.additive(u);
    while (this.isOp(this.peek(), "&")) {
      this.next();
      a = { t: "bin", op: "&", a, b: this.additive(u) };
    }
    return a;
  }

  additive(u: boolean): Node {
    let a = this.mult(u);
    for (;;) {
      const t = this.peek();
      if (!this.isOp(t, "+", "-")) return a;
      this.next();
      a = { t: "bin", op: t!.text, a, b: this.mult(u) };
    }
  }

  mult(u: boolean): Node {
    let a = this.power(u);
    for (;;) {
      const t = this.peek();
      if (!this.isOp(t, "*", "/")) return a;
      this.next();
      a = { t: "bin", op: t!.text, a, b: this.power(u) };
    }
  }

  power(u: boolean): Node {
    let a = this.unary(u);
    while (this.isOp(this.peek(), "^")) {
      this.next();
      a = { t: "bin", op: "^", a, b: this.unary(u) };
    }
    return a;
  }

  unary(u: boolean): Node {
    const start = this.here();
    return this.spanned(start, this.unary0(u));
  }

  private unary0(u: boolean): Node {
    const t = this.peek();
    if (this.isOp(t, "-", "+")) {
      this.next();
      return { t: "un", op: t!.text as "-" | "+", a: this.unary(u) };
    }
    let a = this.refExpr(u);
    while (this.isOp(this.peek(), "%")) {
      this.next();
      a = { t: "un", op: "%", a };
    }
    return a;
  }

  /** Union (in parentheses only), intersection and range operators. */
  refExpr(u: boolean): Node {
    let a = this.intersect();
    if (!u) return a;
    while (this.peek()?.k === "sep") {
      this.next();
      a = { t: "bin", op: ",", a, b: this.intersect() };
    }
    return a;
  }

  intersect(): Node {
    let a = this.range();
    for (;;) {
      // A space between two operands is the intersection operator.
      const raw = this.toks[this.i];
      if (!raw || raw.k !== "ws") return a;
      const nx = this.peek();
      if (!nx || !(nx.k === "ref" || nx.k === "name" || nx.k === "func" || nx.k === "(" || (nx.k === "op" && nx.text === "@"))) return a;
      if (!isRefish(a)) return a;
      a = { t: "bin", op: " ", a, b: this.range() };
    }
  }

  range(): Node {
    let a = this.primary();
    while (this.isOp(this.peek(), ":")) {
      this.next();
      a = { t: "bin", op: ":", a, b: this.primary() };
    }
    return a;
  }

  primary(): Node {
    const start = this.here();
    return this.spanned(start, this.primary0());
  }

  private primary0(): Node {
    const t = this.next();
    if (!t) throw new FormulaError("unexpected end");
    switch (t.k) {
      case "num": return { t: "num", v: t.num! };
      case "str": return { t: "str", v: t.str! };
      case "bool": return { t: "bool", v: t.bool! };
      case "err": return { t: "err", v: t.err! };
      case "ref": {
        const r = refRect(t.ref!);
        return this.postfix({ t: "ref", sheet: t.ref!.sheet, sheet2: t.ref!.sheet2, ...r });
      }
      case "name": return this.postfix({ t: "name", name: t.name!, sheet: t.sheet });
      case "struct": return { t: "struct", text: t.text };
      case "func": {
        this.expect("(");
        const args: Node[] = [];
        if (this.peek()?.k === ")") {
          this.next();
        } else {
          for (;;) {
            const p = this.peek();
            if (p && (p.k === "sep" || p.k === ")")) args.push({ t: "miss" });
            else args.push(this.expr(false));
            const q = this.next();
            if (!q) throw new FormulaError("missing )");
            if (q.k === ")") break;
            if (q.k !== "sep") throw new FormulaError(`unexpected ${q.text}`);
          }
        }
        return this.postfix({ t: "fn", name: t.name!, args });
      }
      case "(": {
        const a = this.expr(true);
        this.expect(")");
        return this.postfix({ t: "paren", a });
      }
      case "{": return this.array();
      case "op":
        if (t.text === "@") return { t: "un", op: "@", a: this.range() };
        if (t.text === "-" || t.text === "+") return { t: "un", op: t.text, a: this.unary(false) };
        break;
    }
    throw new FormulaError(`unexpected ${t.text}`);
  }

  postfix(n: Node): Node {
    const raw = this.toks[this.i];
    if (raw && raw.k === "op" && raw.text === "#") {
      this.i++;
      return { t: "un", op: "#", a: n };
    }
    return n;
  }

  array(): Node {
    const rows: Scalar[][] = [[]];
    for (;;) {
      let t = this.next();
      if (!t) throw new FormulaError("missing }");
      let sign = 1;
      while (t && t.k === "op" && (t.text === "-" || t.text === "+")) {
        if (t.text === "-") sign = -sign;
        t = this.next();
      }
      if (!t) throw new FormulaError("missing }");
      let v: Scalar;
      if (t.k === "num") v = sign * t.num!;
      else if (sign !== 1) throw new FormulaError("sign");
      else if (t.k === "str") v = t.str!;
      else if (t.k === "bool") v = t.bool!;
      else if (t.k === "err") v = err(t.err!);
      else throw new FormulaError("array constant");
      rows[rows.length - 1].push(v);
      const s = this.next();
      if (!s) throw new FormulaError("missing }");
      if (s.k === "}") break;
      if (s.k === "arow") rows.push([]);
      else if (s.k !== "acol") throw new FormulaError("array separator");
    }
    const cols = rows[0].length;
    if (rows.some((r) => r.length !== cols)) throw new FormulaError("ragged array");
    return { t: "arr", rows: rows.length, cols, v: rows.flat() };
  }
}

function isRefish(n: Node): boolean {
  return n.t === "ref" || n.t === "name" || n.t === "fn" || n.t === "paren" || (n.t === "bin" && (n.op === ":" || n.op === " " || n.op === ",")) || (n.t === "un" && (n.op === "@" || n.op === "#"));
}

const CACHE = new Map<string, Node | FormulaError>();

/** The syntax tree of a formula in the file form (no "="); throws FormulaError when it is not one. */
export function parseFormula(file: string, dialect: Dialect = "file"): Node {
  const key = dialect + "\u0000" + file;
  let n = CACHE.get(key);
  if (!n) {
    try {
      n = new Parser(lex(file, dialect)).parseFormula();
    } catch (e) {
      if (!(e instanceof FormulaError)) throw e;
      n = e;
    }
    if (CACHE.size > 50000) CACHE.clear();
    CACHE.set(key, n);
  }
  if (n instanceof FormulaError) throw n;
  return n;
}

/**
 * A formula in the local form with the source offsets of its nodes (without "="; offsets count
 * from the text's start, "=" included when given). Operators get the span from their left
 * operand's start to their right operand's end.
 */
export function parseWithSpans(local: string): { node: Node; span: (n: Node) => [number, number] } {
  const p = new Parser(lex(local, "local"));
  p.spans = new Map();
  const node = p.parseFormula();
  const spans = p.spans;
  const span = (n: Node): [number, number] => {
    let s = spans.get(n);
    if (!s) {
      if (n.t === "bin") s = [span(n.a)[0], span(n.b)[1]];
      else if (n.t === "un" || n.t === "paren") s = span(n.a);
      else s = [0, local.length];
      spans.set(n, s);
    }
    return s;
  };
  return { node, span };
}

/** Whether the text is a formula Excel would accept (local form, with or without "="). */
export function isValidLocal(local: string): boolean {
  try {
    parseFormula(local.startsWith("=") ? local.slice(1) : local, "local");
    return true;
  } catch {
    return false;
  }
}
