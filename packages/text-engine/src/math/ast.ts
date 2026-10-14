// Equations as a tree shaped like Office Math (OMML): text runs and structures (fraction,
// scripts, root, brackets, large operators, functions, accents, bars, limits, group
// characters, matrices, equation arrays, boxes). The tree is what is stored, edited (edit.ts)
// and written as OMML; LaTeX is read for templates and the Formel galleries, and written as a
// canonical key and for the LaTeX conversion.

// ----- symbols ---------------------------------------------------------------------------------

/** LaTeX command → character (letters, operators, relations, arrows, misc). */
export const SYMBOLS: Record<string, string> = {
  alpha: "α", beta: "β", gamma: "γ", delta: "δ", epsilon: "ϵ", varepsilon: "ε", zeta: "ζ", eta: "η", theta: "θ",
  vartheta: "ϑ", iota: "ι", kappa: "κ", lambda: "λ", mu: "μ", nu: "ν", xi: "ξ", omicron: "ο", pi: "π", varpi: "ϖ",
  rho: "ρ", varrho: "ϱ", sigma: "σ", varsigma: "ς", tau: "τ", upsilon: "υ", phi: "ϕ", varphi: "φ", chi: "χ",
  psi: "ψ", omega: "ω", Gamma: "Γ", Delta: "Δ", Theta: "Θ", Lambda: "Λ", Xi: "Ξ", Pi: "Π", Sigma: "Σ",
  Upsilon: "Υ", Phi: "Φ", Psi: "Ψ", Omega: "Ω",
  cdot: "⋅", times: "×", div: "÷", pm: "±", mp: "∓", ast: "∗", star: "⋆", circ: "∘", bullet: "∙",
  oplus: "⊕", otimes: "⊗", setminus: "∖", cup: "∪", cap: "∩", wedge: "∧", land: "∧", vee: "∨", lor: "∨",
  le: "≤", leq: "≤", ge: "≥", geq: "≥", ne: "≠", neq: "≠", approx: "≈", equiv: "≡", sim: "∼", simeq: "≃",
  cong: "≅", propto: "∝", ll: "≪", gg: "≫", prec: "≺", succ: "≻", coloneq: "≔", coloneqq: "≔", doteq: "≐",
  in: "∈", notin: "∉", ni: "∋", subset: "⊂", subseteq: "⊆", supset: "⊃", supseteq: "⊇", perp: "⊥",
  parallel: "∥", mid: "∣", nmid: "∤",
  to: "→", rightarrow: "→", leftarrow: "←", gets: "←", leftrightarrow: "↔", Rightarrow: "⇒", Leftarrow: "⇐",
  Leftrightarrow: "⇔", implies: "⟹", impliedby: "⟸", iff: "⟺", longrightarrow: "⟶", longleftarrow: "⟵",
  mapsto: "↦", uparrow: "↑", downarrow: "↓", nearrow: "↗", searrow: "↘", rightleftharpoons: "⇌",
  infty: "∞", partial: "∂", nabla: "∇", forall: "∀", exists: "∃", nexists: "∄", neg: "¬", lnot: "¬",
  emptyset: "∅", varnothing: "∅", angle: "∠", measuredangle: "∡", triangle: "△", square: "□", degree: "°",
  prime: "′", ell: "ℓ", hbar: "ℏ", Re: "ℜ", Im: "ℑ", aleph: "ℵ", wp: "℘", therefore: "∴", because: "∵",
  cdots: "⋯", ldots: "…", dots: "…", vdots: "⋮", ddots: "⋱", langle: "⟨", rangle: "⟩", lfloor: "⌊",
  rfloor: "⌋", lceil: "⌈", rceil: "⌉", vert: "|", Vert: "‖", lvert: "|", rvert: "|", lVert: "‖", rVert: "‖",
  colon: ":", lbrace: "{", rbrace: "}", dagger: "†", celsius: "℃", euro: "€", checkmark: "✓",
};

/** Large operators: command → character. */
export const NARY: Record<string, string> = {
  sum: "∑", prod: "∏", coprod: "∐", int: "∫", iint: "∬", iiint: "∭", oint: "∮", oiint: "∯",
  bigcup: "⋃", bigcap: "⋂", bigvee: "⋁", bigwedge: "⋀", bigoplus: "⨁", bigotimes: "⨂",
};

export const FUNCTIONS = new Set(["sin", "cos", "tan", "cot", "sec", "csc", "arcsin", "arccos", "arctan", "sinh", "cosh", "tanh",
  "coth", "log", "ln", "lg", "exp", "lim", "max", "min", "sup", "inf", "det", "gcd", "deg", "arg", "dim", "ker", "Pr",
  "hom", "mod", "liminf", "limsup", "sgn"]);
/** Functions whose limits sit below them (lim, max …). */
export const LIMIT_FUNCTIONS = new Set(["lim", "max", "min", "sup", "inf", "liminf", "limsup", "det", "gcd", "Pr"]);

/** Accents: command → combining character (m:acc). */
export const ACCENTS: Record<string, string> = {
  hat: "\u0302", widehat: "\u0302", tilde: "\u0303", widetilde: "\u0303", dot: "\u0307", ddot: "\u0308",
  vec: "\u20d7", overrightarrow: "\u20d7", bar: "\u0305", check: "\u030c", breve: "\u0306", acute: "\u0301",
  grave: "\u0300", mathring: "\u030a", overleftarrow: "\u20d6",
};

/** Spaces: command → Unicode space (as Word writes them in m:t). */
export const SPACES: Record<string, string> = {
  ",": "\u2009", ":": "\u2005", ";": "\u2004", " ": " ", quad: "\u2003", qquad: "\u2003\u2003", enspace: "\u2002",
  thinspace: "\u2009", "!": "",
};

export const ESCAPED = new Set(["{", "}", "%", "#", "&", "_", "$"]);

/** Bracket characters for \left … \right. */
export const DELIMS: Record<string, string> = {
  "(": "(", ")": ")", "[": "[", "]": "]", "\\{": "{", "\\}": "}", "|": "|", ".": "", "\\|": "‖", "\\langle": "⟨",
  "\\rangle": "⟩", "\\lfloor": "⌊", "\\rfloor": "⌋", "\\lceil": "⌈", "\\rceil": "⌉", "\\vert": "|", "\\Vert": "‖",
  "\\lvert": "|", "\\rvert": "|", "\\lVert": "‖", "\\rVert": "‖", "\\lbrace": "{", "\\rbrace": "}", "/": "/",
};

export const MATRIX_ENVS: Record<string, [string, string]> = {
  matrix: ["", ""], pmatrix: ["(", ")"], bmatrix: ["[", "]"], Bmatrix: ["{", "}"], vmatrix: ["|", "|"],
  Vmatrix: ["‖", "‖"], smallmatrix: ["", ""], array: ["", ""],
};

// Characters that end the operand of a sum, integral or function (Word's "argument").
export const OPERAND_END = new Set(["+", "-", "−", "=", "<", ">", ",", ";", "±", "∓", "≤", "≥", "≠", "≈", "≡", "→", "⇒", "⇔", "∼", "≅", "∝", "×", "⋅", "÷"]);

// ----- AST ---------------------------------------------------------------------------------

export type Seq = MNode[];
export type MNode =
  | { k: "t"; s: string; sty?: "p" | "b" | "bi" | "i"; scr?: string; nor?: boolean }
  | { k: "frac"; num: Seq; den: Seq; type?: "noBar" | "lin" | "skw" }
  /** aln: Skripts ausrichten (m:alnScr). */
  | { k: "scr"; base: Seq; sub?: Seq; sup?: Seq; pre?: boolean; aln?: boolean }
  | { k: "rad"; deg?: Seq; e: Seq }
  /** Brackets: noGrow = Eckige Klammern dehnen off; shp "match" = an Argumenthöhe anpassen. */
  | { k: "d"; beg: string; end: string; items: Seq[]; noGrow?: boolean; shp?: "match" }
  /** grow: N-ären Operator dehnen. */
  | { k: "nary"; chr: string; sub?: Seq; sup?: Seq; e: Seq; und: boolean; grow?: boolean }
  | { k: "func"; name: Seq; e: Seq }
  | { k: "acc"; chr: string; e: Seq }
  | { k: "bar"; top: boolean; e: Seq }
  | { k: "lim"; upper: boolean; e: Seq; lim: Seq }
  | { k: "group"; chr: string; top: boolean; e: Seq }
  /** plcHide: Platzhalter ausblenden; baseJc: Matrixausrichtung; mcJc: Spaltenausrichtung per column. */
  | { k: "m"; rows: Seq[][]; plcHide?: boolean; baseJc?: "top" | "bot"; mcJc?: ("left" | "center" | "right")[] }
  | { k: "eqArr"; rows: Seq[] }
  | { k: "box"; e: Seq };

// ----- LaTeX → AST ---------------------------------------------------------------------------

type Tok = { t: "cmd"; v: string } | { t: "ch"; v: string } | { t: "{" } | { t: "}" } | { t: "^" } | { t: "_" } | { t: "&" } | { t: "[" } | { t: "]" };

function tokenize(src: string): Tok[] {
  const out: Tok[] = [];
  for (let i = 0; i < src.length; i++) {
    const c = src[i];
    if (c === "\\") {
      let j = i + 1;
      if (/[A-Za-z]/.test(src[j] ?? "")) { while (j < src.length && /[A-Za-z]/.test(src[j])) j++; out.push({ t: "cmd", v: src.slice(i + 1, j) }); i = j - 1; }
      else if (j < src.length) { out.push({ t: "cmd", v: src[j] }); i = j; }
      continue;
    }
    if (c === "#" && src[i + 1] === "?") { i++; continue; } // MathLive placeholder
    if (c === "{" || c === "}" || c === "^" || c === "_" || c === "&" || c === "[" || c === "]") { out.push({ t: c } as Tok); continue; }
    if (/\s/.test(c)) continue;
    if (c === "'") { out.push({ t: "ch", v: "′" }); continue; }
    // Surrogate pairs stay together.
    const cp = src.codePointAt(i)!;
    const ch = String.fromCodePoint(cp);
    if (ch.length === 2) i++;
    // LaTeX's "-" is a minus sign (Word's galleries write U+2212); typing goes into the tree directly.
    out.push({ t: "ch", v: ch === "-" ? "−" : ch === "*" ? "∗" : ch });
  }
  return out;
}

class Parser {
  i = 0;
  constructor(readonly toks: Tok[]) {}

  private peek(): Tok | undefined { return this.toks[this.i]; }
  private next(): Tok | undefined { return this.toks[this.i++]; }

  /** A sequence until `}` / end / a stop command (\right, \end, \\, &). */
  seq(stops: { right?: boolean; end?: boolean; rows?: boolean; bracket?: boolean } = {}): Seq {
    const out: Seq = [];
    for (;;) {
      const t = this.peek();
      if (!t || t.t === "}") break;
      if (stops.bracket && t.t === "]") break;
      if (t.t === "cmd" && ((stops.right && (t.v === "right" || t.v === "middle")) || (stops.end && t.v === "end") || (stops.rows && t.v === "\\"))) break;
      if (stops.rows && t.t === "&") break;
      this.atom(out);
    }
    return mergeText(out);
  }

  /** Argument of a command: a group or a single token. */
  arg(): Seq {
    const t = this.peek();
    if (!t) return [];
    if (t.t === "{") { this.next(); const s = this.seq(); if (this.peek()?.t === "}") this.next(); return s; }
    const out: Seq = [];
    this.atom(out, true);
    return out;
  }

  private optArg(): Seq | undefined {
    if (this.peek()?.t !== "[") return undefined;
    this.next();
    const s = this.seq({ bracket: true });
    if (this.peek()?.t === "]") this.next();
    return s;
  }

  private text(t: string, extra: Partial<Extract<MNode, { k: "t" }>> = {}): MNode {
    return { k: "t", s: t, ...extra };
  }

  /** Scripts after a base (x^2_i, \sum_a^b). */
  private scripts(): { sub?: Seq; sup?: Seq } {
    const r: { sub?: Seq; sup?: Seq } = {};
    for (;;) {
      const t = this.peek();
      if (t?.t === "cmd" && (t.v === "limits" || t.v === "nolimits")) { this.next(); continue; }
      if (t?.t === "^" && !r.sup) { this.next(); r.sup = this.arg(); continue; }
      if (t?.t === "_" && !r.sub) { this.next(); r.sub = this.arg(); continue; }
      return r;
    }
  }

  /** The operand of a large operator or function: up to the next operator at this level. */
  private operand(): Seq {
    const out: Seq = [];
    for (;;) {
      const t = this.peek();
      if (!t || t.t === "}" || t.t === "&" || t.t === "]") break;
      if (t.t === "ch" && OPERAND_END.has(t.v)) break;
      if (t.t === "cmd") {
        if (t.v === "right" || t.v === "middle" || t.v === "end" || t.v === "\\") break;
        if (SYMBOLS[t.v] && OPERAND_END.has(SYMBOLS[t.v])) break;
        if ((NARY[t.v] || FUNCTIONS.has(t.v)) && out.length) break;
      }
      this.atom(out);
    }
    return mergeText(out);
  }

  /** Parses one atom (with its scripts) and appends it. */
  atom(out: Seq, single = false): void {
    const t = this.next();
    if (!t) return;
    let node: MNode | Seq | null = null;
    switch (t.t) {
      case "{": { const s = this.seq(); if (this.peek()?.t === "}") this.next(); node = s; break; }
      case "ch": node = this.text(t.v); break;
      case "[": node = this.text("["); break;
      case "]": node = this.text("]"); break;
      case "&": out.push({ k: "t", s: "&" }); return; // an alignment mark (equation arrays)
      case "^": case "_": {
        // A script without a base ({}^{14}C).
        this.i--;
        const sc = this.scripts();
        out.push({ k: "scr", base: [], ...sc });
        return;
      }
      case "}": return;
      case "cmd": node = this.command(t.v); break;
    }
    if (node === null) return;
    if (single) { if (Array.isArray(node)) out.push(...node); else out.push(node); return; }
    if (!Array.isArray(node) && (node.k === "nary" || node.k === "func")) { out.push(node); return; }
    const sc = this.peek()?.t === "^" || this.peek()?.t === "_" ? this.scripts() : null;
    if (sc && (sc.sub || sc.sup)) {
      let base = Array.isArray(node) ? node : [node];
      // x^2 on a run "ab": only the last character carries the script.
      const last = base[base.length - 1];
      if (base.length === 1 && last?.k === "t" && [...last.s].length > 1 && !last.nor) {
        const chars = [...last.s];
        out.push({ ...last, s: chars.slice(0, -1).join("") });
        base = [{ ...last, s: chars[chars.length - 1] }];
      }
      // Overbrace with a label above becomes a limit, as Word builds it.
      if (base.length === 1 && base[0].k === "group" && ((base[0].top && sc.sup) || (!base[0].top && sc.sub))) {
        out.push({ k: "lim", upper: base[0].top, e: base, lim: (base[0].top ? sc.sup : sc.sub)! });
        return;
      }
      out.push({ k: "scr", base, sub: sc.sub, sup: sc.sup });
      return;
    }
    if (Array.isArray(node)) out.push(...node); else out.push(node);
  }

  private command(v: string): MNode | Seq | null {
    if (v === "frac" || v === "dfrac" || v === "tfrac" || v === "cfrac") return { k: "frac", num: this.arg(), den: this.arg() };
    if (v === "binom" || v === "dbinom" || v === "tbinom") return { k: "d", beg: "(", end: ")", items: [[{ k: "frac", num: this.arg(), den: this.arg(), type: "noBar" }]] };
    if (v === "genfrac") { this.arg(); this.arg(); const th = this.arg(); this.arg(); const num = this.arg(), den = this.arg(); return { k: "frac", num, den, type: th.length && seqText(th) === "0pt" ? "noBar" : undefined }; }
    if (v === "sqrt") { const deg = this.optArg(); return { k: "rad", deg, e: this.arg() }; }
    if (v === "hyphen") return this.text("-");
    if (v === "sfrac") return { k: "frac", num: this.arg(), den: this.arg(), type: "skw" };
    if (v === "left") return this.leftRight();
    if (/^(big|Big|bigg|Bigg)[lrm]?$/.test(v)) { const d = this.delimToken(); return d ? this.text(d) : null; }
    if (NARY[v]) {
      const sc = this.scripts();
      const und = !v.includes("int");
      return { k: "nary", chr: NARY[v], sub: sc.sub, sup: sc.sup, e: this.operand(), und };
    }
    if (FUNCTIONS.has(v) || v === "operatorname") {
      const nameText = v === "operatorname" ? seqText(this.arg()) : v;
      let name: Seq = [this.text(nameText, { sty: "p" })];
      const sc = this.scripts();
      if (sc.sub && LIMIT_FUNCTIONS.has(nameText) && !sc.sup) name = [{ k: "lim", upper: false, e: name, lim: sc.sub }];
      else if (sc.sub || sc.sup) name = [{ k: "scr", base: name, sub: sc.sub, sup: sc.sup }];
      return { k: "func", name, e: this.operand() };
    }
    if (ACCENTS[v]) return { k: "acc", chr: ACCENTS[v], e: this.arg() };
    if (v === "overline") return { k: "bar", top: true, e: this.arg() };
    if (v === "underline") return { k: "bar", top: false, e: this.arg() };
    if (v === "overbrace") return { k: "group", chr: "⏞", top: true, e: this.arg() };
    if (v === "underbrace") return { k: "group", chr: "⏟", top: false, e: this.arg() };
    if (v === "overset" || v === "stackrel") { const lim = this.arg(); return { k: "lim", upper: true, e: this.arg(), lim }; }
    if (v === "underset") { const lim = this.arg(); return { k: "lim", upper: false, e: this.arg(), lim }; }
    if (v === "boxed") return { k: "box", e: this.arg() };
    if (v === "text" || v === "textrm" || v === "mbox" || v === "textnormal" || v === "textup") return this.text(this.rawGroup(), { nor: true });
    if (v === "textbf") return this.text(this.rawGroup(), { nor: true, sty: "b" });
    if (v === "textit") return this.text(this.rawGroup(), { nor: true, sty: "i" });
    const styled = ({ mathrm: { sty: "p" }, mathup: { sty: "p" }, mathbf: { sty: "b" }, mathit: { sty: "i" }, boldsymbol: { sty: "bi" }, bm: { sty: "bi" },
      mathbb: { scr: "double-struck", sty: "p" }, mathcal: { scr: "script", sty: "p" }, mathscr: { scr: "script", sty: "p" }, mathfrak: { scr: "fraktur", sty: "p" },
      mathsf: { scr: "sans-serif", sty: "p" }, mathtt: { scr: "monospace", sty: "p" } } as Record<string, { sty?: "p" | "b" | "bi" | "i"; scr?: string }>)[v];
    if (styled) return this.arg().map((n) => (n.k === "t" ? { ...n, ...styled } : n));
    if (v === "begin") return this.environment();
    if (v === "\\") return this.text("\n"); // row break at the top level → equation array (see below)
    if (v in SPACES) return SPACES[v] ? this.text(SPACES[v]) : null;
    if (ESCAPED.has(v)) return this.text(v);
    if (v === "|") return this.text("‖");
    if (SYMBOLS[v]) return this.text(SYMBOLS[v]);
    if (v === "placeholder") { this.optArg(); this.arg(); return null; }
    if (["displaystyle", "textstyle", "scriptstyle", "limits", "nolimits", "right", "middle", "end", "mathstrut", "strut"].includes(v)) return null;
    if (v === "phantom" || v === "hphantom" || v === "vphantom") { this.arg(); return null; }
    if (v === "color" || v === "textcolor") { this.arg(); return v === "textcolor" ? this.arg() : null; }
    if (v === "mathop" || v === "mathbin" || v === "mathrel" || v === "mathord" || v === "mathopen" || v === "mathclose") return this.arg();
    return this.text(`\\${v}`);
  }

  /** The text of a {…} group verbatim (\text{…}). */
  private rawGroup(): string {
    if (this.peek()?.t !== "{") { const t = this.next(); return t && t.t === "ch" ? t.v : ""; }
    this.next();
    let depth = 1, s = "";
    for (;;) {
      const t = this.next();
      if (!t) break;
      if (t.t === "{") depth++;
      if (t.t === "}" && --depth === 0) break;
      s += t.t === "ch" ? t.v : t.t === "cmd" ? (SPACES[t.v] !== undefined ? " " : ESCAPED.has(t.v) ? t.v : SYMBOLS[t.v] ?? "") : t.t === "{" || t.t === "}" ? "" : t.t;
    }
    return s.replace(/−/g, "-");
  }

  /** Reads a delimiter after \left, \right, \big … */
  private delimToken(): string | null {
    const t = this.next();
    if (!t) return null;
    if (t.t === "ch") return DELIMS[t.v === "−" ? "-" : t.v] ?? t.v;
    if (t.t === "[" || t.t === "]") return t.t;
    if (t.t === "cmd") return DELIMS[`\\${t.v}`] ?? (t.v === "." ? "" : SYMBOLS[t.v] ?? "");
    return null;
  }

  private leftRight(): MNode {
    const beg = this.delimToken() ?? "";
    const items: Seq[] = [];
    let end = "";
    for (;;) {
      items.push(this.seq({ right: true }));
      const t = this.next();
      if (!t) break;
      if (t.t === "cmd" && t.v === "middle") { this.delimToken(); continue; }
      if (t.t === "cmd" && t.v === "right") { end = this.delimToken() ?? ""; break; }
    }
    return { k: "d", beg, end, items };
  }

  private environment(): MNode | null {
    const name = seqText(this.argRaw());
    if (name === "array") this.argRaw(); // column spec
    const rows: Seq[][] = [];
    let row: Seq[] = [];
    for (;;) {
      row.push(this.seq({ end: true, rows: true }));
      const t = this.next();
      if (!t) break;
      if (t.t === "&") continue;
      if (t.t === "cmd" && t.v === "\\") { rows.push(row); row = []; continue; }
      if (t.t === "cmd" && t.v === "end") { this.argRaw(); break; }
    }
    if (row.length && !(row.length === 1 && !row[0].length)) rows.push(row);
    if (name === "cases" || name === "dcases") {
      // Word: a brace before an equation array; "&" becomes a space.
      return { k: "d", beg: "{", end: "", items: [[{ k: "eqArr", rows: rows.map((r) => joinCells(r)) }]] };
    }
    if (name in MATRIX_ENVS) {
      const [b, e] = MATRIX_ENVS[name];
      const m: MNode = { k: "m", rows };
      return b || e ? { k: "d", beg: b, end: e, items: [[m]] } : m;
    }
    // aligned, align, gathered, split, …: one line per row.
    return { k: "eqArr", rows: rows.map((r) => joinCells(r)) };
  }

  private argRaw(): Seq {
    if (this.peek()?.t !== "{") return [];
    this.next();
    let s = "";
    while (this.peek() && this.peek()!.t !== "}") { const t = this.next()!; s += t.t === "ch" ? t.v : ""; }
    this.next();
    return [{ k: "t", s }];
  }
}

function joinCells(row: Seq[]): Seq {
  const out: Seq = [];
  row.forEach((c, i) => { if (i) out.push({ k: "t", s: "&" }); out.push(...c); });
  return mergeText(out);
}

export function seqText(s: Seq): string {
  return s.map((n) => (n.k === "t" ? n.s : "")).join("");
}

/** Joins neighbouring text nodes with the same style. */
export function mergeText(s: Seq): Seq {
  const out: Seq = [];
  for (const n of s) {
    const last = out[out.length - 1];
    if (n.k === "t" && last?.k === "t" && last.sty === n.sty && last.scr === n.scr && last.nor === n.nor && n.s !== "\n" && last.s !== "\n") {
      out[out.length - 1] = { ...last, s: last.s + n.s };
    } else if (!(n.k === "t" && n.s === "")) out.push(n);
  }
  return out;
}

export function parseLatex(latex: string): Seq {
  const p = new Parser(tokenize(latex));
  const out: Seq = [];
  while (p.i < p.toks.length) {
    const s = p.seq();
    out.push(...s);
    p.i++; // a stray "}"
  }
  return mergeText(out);
}


// ----- tree → LaTeX ----------------------------------------------------------------------------

const REVERSE = new Map<string, string>();
for (const [k, v] of Object.entries(SYMBOLS)) if (!REVERSE.has(v)) REVERSE.set(v, `\\${k}`);
REVERSE.set("−", "-");
REVERSE.set("∗", "*");
REVERSE.set("′", "'");
REVERSE.set("|", "|");
REVERSE.set("…", "\\ldots");
const REVERSE_NARY = new Map(Object.entries(NARY).map(([k, v]) => [v, `\\${k}`]));
export const REVERSE_ACCENT = new Map<string, string>([["̂", "hat"], ["̃", "tilde"], ["̇", "dot"], ["̈", "ddot"], ["⃗", "vec"],
  ["̅", "bar"], ["̌", "check"], ["̆", "breve"], ["́", "acute"], ["̀", "grave"], ["̊", "mathring"], ["⃖", "overleftarrow"]]);
const REVERSE_SPACES = new Map<string, string>([[" ", "\\,"], [" ", "\\:"], [" ", "\\;"], [" ", "\\quad "], [" ", "\\enspace "], [" ", "\\,"]]);
const LEFT: Record<string, string> = { "(": "(", ")": ")", "[": "[", "]": "]", "{": "\\{", "}": "\\}", "|": "|", "‖": "\\Vert", "⟨": "\\langle", "⟩": "\\rangle",
  "⌊": "\\lfloor", "⌋": "\\rfloor", "⌈": "\\lceil", "⌉": "\\rceil", "": ".", "〈": "\\langle", "〉": "\\rangle" };

/** A math run's text as LaTeX (spaces kept as "\ "). */
function latexText(s: string): string {
  let out = "";
  for (const ch of s) {
    if (REVERSE_SPACES.has(ch)) { out += REVERSE_SPACES.get(ch); continue; }
    if (ch === " ") { out += "\\ "; continue; }
    if (ch === "\n") { out += "\\\\"; continue; }
    if (ESCAPED.has(ch)) { out += `\\${ch}`; continue; }
    if (ch === "\\") { out += "\\backslash "; continue; }
    if (ch === "^") { out += "\\hat{}"; continue; }
    if (ch === "~") { out += "\\sim "; continue; }
    const r = REVERSE.get(ch);
    out += r ? (/[a-z]$/i.test(r) ? `${r} ` : r) : ch;
  }
  return out;
}

const grp = (s: string) => `{${s}}`;
function scriptBase(s: string): string {
  const t = s.trim();
  if ([...t].length === 1 || /^\\[A-Za-z]+\s*$/.test(t)) return t;
  return grp(t);
}

function runLatex(n: Extract<MNode, { k: "t" }>): string {
  if (n.nor) return `\\text{${n.s.replace(/[{}\\]/g, "")}}`;
  const body = latexText(n.s);
  if (n.scr) {
    const cmd = ({ "double-struck": "mathbb", script: "mathcal", fraktur: "mathfrak", "sans-serif": "mathsf", monospace: "mathtt" } as Record<string, string>)[n.scr];
    if (cmd) return `\\${cmd}{${body}}`;
  }
  if (n.sty === "p" && /[A-Za-z]/.test(n.s)) return FUNCTIONS.has(n.s) ? `\\${n.s}` : `\\mathrm{${body}}`;
  if (n.sty === "b") return `\\mathbf{${body}}`;
  if (n.sty === "bi") return `\\boldsymbol{${body}}`;
  return body;
}

/** Canonical LaTeX of an equation tree (also its key for caches). */
export function seqToLatex(s: Seq): string {
  let out = "";
  for (const n of mergeText(s)) {
    const part = nodeLatex(n);
    // A command name followed by a letter needs a separating space.
    if (/\\[A-Za-z]+$/.test(out) && /^[A-Za-z]/.test(part)) out += " ";
    out += part;
  }
  return out;
}

function nodeLatex(n: MNode): string {
  const L = seqToLatex;
  switch (n.k) {
    case "t": return runLatex(n);
    case "frac":
      if (n.type === "lin") return `${scriptBase(L(n.num))}/${scriptBase(L(n.den))}`;
      if (n.type === "noBar") return `\\genfrac{}{}{0pt}{}{${L(n.num)}}{${L(n.den)}}`;
      if (n.type === "skw") return `\\sfrac{${L(n.num)}}{${L(n.den)}}`;
      return `\\frac{${L(n.num)}}{${L(n.den)}}`;
    case "scr": {
      const sub = n.sub ? `_{${L(n.sub)}}` : "", sup = n.sup ? `^{${L(n.sup)}}` : "";
      return n.pre ? `{}${sub}${sup}${scriptBase(L(n.base))}` : `${n.base.length ? scriptBase(L(n.base)) : "{}"}${sub}${sup}`;
    }
    case "rad": return n.deg ? `\\sqrt[${L(n.deg)}]{${L(n.e)}}` : `\\sqrt{${L(n.e)}}`;
    case "d": {
      if (n.items.length === 1 && n.items[0].length === 1) {
        const only = n.items[0][0];
        if (only.k === "frac" && only.type === "noBar" && n.beg === "(" && n.end === ")") return `\\binom{${L(only.num)}}{${L(only.den)}}`;
        if (only.k === "m") {
          const env = ({ "()": "pmatrix", "[]": "bmatrix", "{}": "Bmatrix", "||": "vmatrix", "‖‖": "Vmatrix" } as Record<string, string>)[n.beg + n.end];
          if (env) return matrixLatex(only.rows, env);
        }
        if (only.k === "eqArr" && n.beg === "{" && n.end === "") return `\\begin{cases}${only.rows.map(L).join("\\\\")}\\end{cases}`;
      }
      const l = LEFT[n.beg] ?? latexText(n.beg), r = LEFT[n.end] ?? latexText(n.end);
      return `\\left${l}${n.items.map(L).join("\\middle|")}\\right${r}`;
    }
    case "nary": {
      const cmd = REVERSE_NARY.get(n.chr) ?? `\\mathop{${latexText(n.chr)}}`;
      const lim = cmd.includes("int") && n.und ? "\\limits" : !cmd.includes("int") && !n.und ? "\\nolimits" : "";
      return `${cmd}${lim}${n.sub ? `_{${L(n.sub)}}` : ""}${n.sup ? `^{${L(n.sup)}}` : ""}${grp(L(n.e))}`;
    }
    case "func": {
      const nm = n.name.length === 1 && n.name[0].k === "t" && !FUNCTIONS.has(n.name[0].s) ? `\\operatorname{${n.name[0].s}}` : L(n.name);
      return `${nm}${grp(L(n.e))}`;
    }
    case "acc": return `\\${REVERSE_ACCENT.get(n.chr) ?? "hat"}{${L(n.e)}}`;
    case "bar": return `\\${n.top ? "overline" : "underline"}{${L(n.e)}}`;
    case "lim": {
      const txt = n.e.length === 1 && n.e[0].k === "t" ? n.e[0].s : "";
      if (!n.upper && LIMIT_FUNCTIONS.has(txt)) return `\\${txt}_{${L(n.lim)}}`;
      if (n.e.length === 1 && n.e[0].k === "group") return `${L(n.e)}${n.upper ? "^" : "_"}{${L(n.lim)}}`;
      return `\\${n.upper ? "overset" : "underset"}{${L(n.lim)}}{${L(n.e)}}`;
    }
    case "group":
      if (n.chr === "⏟" || n.chr === "⏞") return `\\${n.top ? "overbrace" : "underbrace"}{${L(n.e)}}`;
      return `\\${n.top ? "overset" : "underset"}{${latexText(n.chr)}}{${L(n.e)}}`;
    case "m": return matrixLatex(n.rows, "matrix");
    case "eqArr": return `\\begin{aligned}${n.rows.map(L).join("\\\\")}\\end{aligned}`;
    case "box": return `\\boxed{${L(n.e)}}`;
  }
}

function matrixLatex(rows: Seq[][], env: string): string {
  return `\\begin{${env}}${rows.map((r) => r.map(seqToLatex).join("&")).join("\\\\")}\\end{${env}}`;
}
