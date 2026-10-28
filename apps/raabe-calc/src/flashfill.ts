// Blitzvorschau (Strg+E): from one or two typed examples, a program that builds the column's
// text out of the texts to its left, then applied to the other rows. The pieces a program is
// made of: a word of a source text (counted from the left, or the last word), its first letter,
// a stretch of a source text between delimiters (from the start or after a character, to the
// end or before a character), each with its case kept or changed, and literal text. The
// cheapest description of the example wins; on equal cost words before stretches and earlier
// columns before later ones. Checked against Excel's own results (probe_excel_flashfill.py).

export type Case = "keep" | "upper" | "lower" | "proper";

export type Piece =
  | { kind: "word"; src: number; index: number; case: Case }
  | { kind: "initial"; src: number; index: number; case: Case }
  | { kind: "span"; src: number; from: Bound; to: Bound; case: Case }
  | { kind: "const"; text: string };

/** A position in a source text: its start or end, or beside the k-th occurrence of a character (k < 0: from the right). */
export type Bound = { at: "start" } | { at: "end" } | { at: "char"; ch: string; k: number };

export type Program = Piece[];

const WORD = /[\p{L}\p{N}]+/gu;

function words(s: string): { text: string; start: number; end: number }[] {
  return [...s.matchAll(WORD)].map((m) => ({ text: m[0], start: m.index!, end: m.index! + m[0].length }));
}

function applyCase(s: string, c: Case): string {
  if (c === "upper") return s.toLocaleUpperCase("de");
  if (c === "lower") return s.toLocaleLowerCase("de");
  if (c === "proper") return s.charAt(0).toLocaleUpperCase("de") + s.slice(1).toLocaleLowerCase("de");
  return s;
}

/** The case change that turns `from` into `to`, if one does. */
function caseOf(from: string, to: string): Case | null {
  for (const c of ["keep", "upper", "lower", "proper"] as const) if (applyCase(from, c) === to) return c;
  return null;
}

function occurrences(s: string, ch: string): number[] {
  const out: number[] = [];
  for (let i = s.indexOf(ch); i >= 0; i = s.indexOf(ch, i + 1)) out.push(i);
  return out;
}

function boundAt(s: string, b: Bound, side: "from" | "to"): number | null {
  if (b.at === "start") return 0;
  if (b.at === "end") return s.length;
  const occ = occurrences(s, b.ch);
  const i = b.k > 0 ? occ[b.k - 1] : occ[occ.length + b.k];
  if (i === undefined) return null;
  return side === "from" ? i + b.ch.length : i;
}

/** The text a piece takes from a row's sources, or null where the row does not have it. */
export function pieceText(p: Piece, sources: string[]): string | null {
  if (p.kind === "const") return p.text;
  const s = sources[p.src] ?? "";
  if (p.kind === "span") {
    const a = boundAt(s, p.from, "from"), b = boundAt(s, p.to, "to");
    if (a === null || b === null || b <= a) return null;
    return applyCase(s.slice(a, b), p.case);
  }
  const ws = words(s);
  const w = p.index >= 0 ? ws[p.index] : ws[ws.length + p.index];
  if (!w) return null;
  return applyCase(p.kind === "initial" ? w.text.charAt(0) : w.text, p.case);
}

export function run(program: Program, sources: string[]): string | null {
  let out = "";
  for (const p of program) {
    const t = pieceText(p, sources);
    if (t === null) return null;
    out += t;
  }
  return out;
}

/** Ways to make `piece` (a part of the example) from the sources, cheapest first. */
function candidates(piece: string, sources: string[]): { p: Piece; cost: number }[] {
  const out: { p: Piece; cost: number }[] = [];
  sources.forEach((s, src) => {
    const ws = words(s);
    ws.forEach((w, i) => {
      // The first word counts from the left, the last from the right, the others from the left.
      const index = i === ws.length - 1 && i > 0 ? -1 : i;
      const c = caseOf(w.text, piece);
      if (c) out.push({ p: { kind: "word", src, index, case: c }, cost: 1 + src * 0.01 });
      const ci = piece.length === 1 ? caseOf(w.text.charAt(0), piece) : null;
      if (ci) out.push({ p: { kind: "initial", src, index, case: ci }, cost: 1 + src * 0.01 });
    });
    // Stretches between delimiters: what follows and precedes them in the source.
    for (let a = 0; a < s.length; a++) {
      for (let b = a + 1; b <= s.length; b++) {
        const c = caseOf(s.slice(a, b), piece);
        if (!c) continue;
        const from: Bound | null = a === 0 ? { at: "start" } : /[\p{L}\p{N}]/u.test(s[a - 1]) ? null : { at: "char", ch: s[a - 1], k: occurrences(s.slice(0, a), s[a - 1]).length };
        const to: Bound | null = b === s.length ? { at: "end" } : /[\p{L}\p{N}]/u.test(s[b]) ? null : { at: "char", ch: s[b], k: occurrences(s.slice(0, b + 1), s[b]).length };
        if (from && to) out.push({ p: { kind: "span", src, from, to, case: c }, cost: 1.1 + src * 0.01 });
      }
    }
  });
  return out.sort((x, y) => x.cost - y.cost);
}

/** The cheapest program that makes `output` from `sources` (dynamic programming over the example). */
function explain(sources: string[], output: string): Program | null {
  const n = output.length;
  const best: ({ cost: number; prog: Program } | null)[] = Array(n + 1).fill(null);
  best[0] = { cost: 0, prog: [] };
  for (let i = 0; i < n; i++) {
    const here = best[i];
    if (!here) continue;
    for (let j = i + 1; j <= n; j++) {
      const piece = output.slice(i, j);
      const cand = candidates(piece, sources);
      const options = cand.length ? [cand[0]] : [];
      // Literal text: punctuation and spaces cheaply, anything else at a price.
      options.push({ p: { kind: "const", text: piece }, cost: /^[^\p{L}\p{N}]+$/u.test(piece) ? 0.5 : 1.5 * piece.length });
      for (const o of options) {
        const cost = here.cost + o.cost;
        if (!best[j] || cost < best[j]!.cost - 1e-9) {
          const last = here.prog.at(-1);
          // Literal pieces next to each other are one.
          const prog = o.p.kind === "const" && last?.kind === "const" ? [...here.prog.slice(0, -1), { kind: "const" as const, text: last.text + o.p.text }] : [...here.prog, o.p];
          best[j] = { cost, prog };
        }
      }
    }
  }
  return best[n]?.prog ?? null;
}

/** A program that gives every example, or null (Excel: no pattern found). */
export function learn(examples: { sources: string[]; output: string }[]): Program | null {
  if (!examples.length) return null;
  for (const e of examples) {
    const prog = explain(e.sources, e.output);
    if (prog && examples.every((x) => run(prog, x.sources) === x.output)) {
      // A program made of literal text only has learnt nothing.
      return prog.some((p) => p.kind !== "const") ? prog : null;
    }
  }
  return null;
}
