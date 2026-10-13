// Equations: Office Math (OMML, m:oMath) in .docx and LaTeX in the editor (MathLive).
// The converter covers what school mathematics uses: fractions, scripts, roots, brackets,
// sums and integrals, functions, limits, accents, matrices and case distinctions. Anything
// else read from a file keeps its XML (the equation is shown but not editable then).

export const M_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math";

// ----- symbols ---------------------------------------------------------------------------------

/** LaTeX command → character (letters, operators, relations, arrows, misc). */
const SYMBOLS: Record<string, string> = {
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
const NARY: Record<string, string> = {
  sum: "∑", prod: "∏", coprod: "∐", int: "∫", iint: "∬", iiint: "∭", oint: "∮", oiint: "∯",
  bigcup: "⋃", bigcap: "⋂", bigvee: "⋁", bigwedge: "⋀", bigoplus: "⨁", bigotimes: "⨂",
};

const FUNCTIONS = new Set(["sin", "cos", "tan", "cot", "sec", "csc", "arcsin", "arccos", "arctan", "sinh", "cosh", "tanh",
  "coth", "log", "ln", "lg", "exp", "lim", "max", "min", "sup", "inf", "det", "gcd", "deg", "arg", "dim", "ker", "Pr",
  "hom", "mod", "liminf", "limsup", "sgn"]);
/** Functions whose limits sit below them (lim, max …). */
const LIMIT_FUNCTIONS = new Set(["lim", "max", "min", "sup", "inf", "liminf", "limsup", "det", "gcd", "Pr"]);

/** Accents: command → combining character (m:acc). */
const ACCENTS: Record<string, string> = {
  hat: "\u0302", widehat: "\u0302", tilde: "\u0303", widetilde: "\u0303", dot: "\u0307", ddot: "\u0308",
  vec: "\u20d7", overrightarrow: "\u20d7", bar: "\u0305", check: "\u030c", breve: "\u0306", acute: "\u0301",
  grave: "\u0300", mathring: "\u030a", overleftarrow: "\u20d6",
};

/** Spaces: command → Unicode space (as Word writes them in m:t). */
const SPACES: Record<string, string> = {
  ",": "\u2009", ":": "\u2005", ";": "\u2004", " ": " ", quad: "\u2003", qquad: "\u2003\u2003", enspace: "\u2002",
  thinspace: "\u2009", "!": "",
};

const ESCAPED = new Set(["{", "}", "%", "#", "&", "_", "$"]);

/** Bracket characters for \left … \right. */
const DELIMS: Record<string, string> = {
  "(": "(", ")": ")", "[": "[", "]": "]", "\\{": "{", "\\}": "}", "|": "|", ".": "", "\\|": "‖", "\\langle": "⟨",
  "\\rangle": "⟩", "\\lfloor": "⌊", "\\rfloor": "⌋", "\\lceil": "⌈", "\\rceil": "⌉", "\\vert": "|", "\\Vert": "‖",
  "\\lvert": "|", "\\rvert": "|", "\\lVert": "‖", "\\rVert": "‖", "\\lbrace": "{", "\\rbrace": "}", "/": "/",
};

const MATRIX_ENVS: Record<string, [string, string]> = {
  matrix: ["", ""], pmatrix: ["(", ")"], bmatrix: ["[", "]"], Bmatrix: ["{", "}"], vmatrix: ["|", "|"],
  Vmatrix: ["‖", "‖"], smallmatrix: ["", ""], array: ["", ""],
};

// Characters that end the operand of a sum, integral or function (Word's "argument").
const OPERAND_END = new Set(["+", "-", "−", "=", "<", ">", ",", ";", "±", "∓", "≤", "≥", "≠", "≈", "≡", "→", "⇒", "⇔", "∼", "≅", "∝", "×", "⋅", "÷"]);

// ----- AST ---------------------------------------------------------------------------------

type Seq = MNode[];
type MNode =
  | { k: "t"; s: string; sty?: "p" | "b" | "bi" | "i"; scr?: string; nor?: boolean }
  | { k: "frac"; num: Seq; den: Seq; type?: "noBar" | "lin" | "skw" }
  | { k: "scr"; base: Seq; sub?: Seq; sup?: Seq; pre?: boolean }
  | { k: "rad"; deg?: Seq; e: Seq }
  | { k: "d"; beg: string; end: string; items: Seq[] }
  | { k: "nary"; chr: string; sub?: Seq; sup?: Seq; e: Seq; und: boolean }
  | { k: "func"; name: Seq; e: Seq }
  | { k: "acc"; chr: string; e: Seq }
  | { k: "bar"; top: boolean; e: Seq }
  | { k: "lim"; upper: boolean; e: Seq; lim: Seq }
  | { k: "group"; chr: string; top: boolean; e: Seq }
  | { k: "m"; rows: Seq[][] }
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
      case "&": return;
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

function seqText(s: Seq): string {
  return s.map((n) => (n.k === "t" ? n.s : "")).join("");
}

/** Joins neighbouring text nodes with the same style. */
function mergeText(s: Seq): Seq {
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

// ----- AST → OMML ------------------------------------------------------------------------------

const xesc = (s: string) => s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
const RPR = '<w:rPr><w:rFonts w:ascii="Cambria Math" w:hAnsi="Cambria Math"/></w:rPr>';
const val = (tag: string, v: string) => `<m:${tag} m:val="${xesc(v)}"/>`;

function ommlRun(n: Extract<MNode, { k: "t" }>): string {
  let pr = "";
  if (n.nor) pr = `<m:nor/>`;
  else if (n.scr || n.sty) pr = (n.scr ? val("scr", n.scr) : "") + (n.sty ? val("sty", n.sty) : "");
  const wpr = n.nor
    ? (n.sty === "b" ? "<w:rPr><w:b/></w:rPr>" : n.sty === "i" ? "<w:rPr><w:i/></w:rPr>" : "")
    : RPR;
  return `<m:r>${pr ? `<m:rPr>${pr}</m:rPr>` : ""}${wpr}<m:t xml:space="preserve">${xesc(n.s)}</m:t></m:r>`;
}

const e = (s: Seq | undefined, tag = "e") => (s && s.length ? `<m:${tag}>${omml(s)}</m:${tag}>` : `<m:${tag}/>`);

function omml(s: Seq): string {
  let out = "";
  for (const n of s) {
    switch (n.k) {
      case "t": out += ommlRun(n); break;
      case "frac": out += `<m:f>${n.type ? `<m:fPr>${val("type", n.type)}</m:fPr>` : ""}${e(n.num, "num")}${e(n.den, "den")}</m:f>`; break;
      case "scr":
        if (n.sub && n.sup) out += `<m:sSubSup>${e(n.base)}${e(n.sub, "sub")}${e(n.sup, "sup")}</m:sSubSup>`;
        else if (n.sup) out += `<m:sSup>${e(n.base)}${e(n.sup, "sup")}</m:sSup>`;
        else out += `<m:sSub>${e(n.base)}${e(n.sub, "sub")}</m:sSub>`;
        break;
      case "rad": out += `<m:rad>${n.deg?.length ? "" : `<m:radPr>${val("degHide", "1")}</m:radPr>`}${e(n.deg, "deg")}${e(n.e)}</m:rad>`; break;
      case "d": {
        const pr = (n.beg !== "(" ? val("begChr", n.beg) : "") + (n.end !== ")" ? val("endChr", n.end) : "");
        out += `<m:d>${pr ? `<m:dPr>${pr}</m:dPr>` : ""}${n.items.map((x) => e(x)).join("")}</m:d>`;
        break;
      }
      case "nary": {
        const pr = (n.chr !== "∫" ? val("chr", n.chr) : "") + val("limLoc", n.und ? "undOvr" : "subSup") +
          (n.sub ? "" : val("subHide", "1")) + (n.sup ? "" : val("supHide", "1"));
        out += `<m:nary><m:naryPr>${pr}</m:naryPr>${e(n.sub, "sub")}${e(n.sup, "sup")}${e(n.e)}</m:nary>`;
        break;
      }
      case "func": out += `<m:func>${e(n.name, "fName")}${e(n.e)}</m:func>`; break;
      case "acc": out += `<m:acc><m:accPr>${val("chr", n.chr)}</m:accPr>${e(n.e)}</m:acc>`; break;
      case "bar": out += `<m:bar><m:barPr>${val("pos", n.top ? "top" : "bot")}</m:barPr>${e(n.e)}</m:bar>`; break;
      case "lim": out += `<m:${n.upper ? "limUpp" : "limLow"}>${e(n.e)}${e(n.lim, "lim")}</m:${n.upper ? "limUpp" : "limLow"}>`; break;
      case "group": out += `<m:groupChr><m:groupChrPr>${val("chr", n.chr)}${val("pos", n.top ? "top" : "bot")}</m:groupChrPr>${e(n.e)}</m:groupChr>`; break;
      case "m": {
        const cols = Math.max(1, ...n.rows.map((r) => r.length));
        out += `<m:m><m:mPr><m:mcs><m:mc><m:mcPr>${val("count", String(cols))}${val("mcJc", "center")}</m:mcPr></m:mc></m:mcs></m:mPr>` +
          n.rows.map((r) => `<m:mr>${Array.from({ length: cols }, (_x, i) => e(r[i])).join("")}</m:mr>`).join("") + "</m:m>";
        break;
      }
      case "eqArr": out += `<m:eqArr>${n.rows.map((r) => e(r)).join("")}</m:eqArr>`; break;
      case "box": out += `<m:borderBox>${e(n.e)}</m:borderBox>`; break;
    }
  }
  return out;
}

/** Splits a top-level sequence at "\\" (several lines → an equation array). */
function lines(s: Seq): Seq[] {
  const out: Seq[] = [[]];
  for (const n of s) {
    if (n.k === "t" && n.s === "\n") out.push([]);
    else out[out.length - 1].push(n);
  }
  return out;
}

/** OMML for an equation: m:oMathPara around it when it stands alone on its line. */
export function latexToOmml(latex: string, display: boolean): string {
  const rows = lines(parseLatex(latex));
  const body = rows.length > 1 ? omml([{ k: "eqArr", rows }]) : omml(rows[0]);
  const math = `<m:oMath>${body}</m:oMath>`;
  return display ? `<m:oMathPara><m:oMathParaPr>${val("jc", "centerGroup")}</m:oMathParaPr>${math}</m:oMathPara>` : math;
}

// ----- AST → MathML (display in the editor) -----------------------------------------------------
// Shown the way Word lays out the same OMML: letters italic (capital Greek upright), numbers
// and function names upright, sums with limits above and below in display equations,
// integrals with limits at the side, lim/max/min with the limit below.

const UPRIGHT_GREEK = /[Α-Ω]/;
const ACCENT_GLYPH: Record<string, string> = {
  "\u0302": "^", "\u0303": "~", "\u0307": "˙", "\u0308": "¨", "\u20d7": "→", "\u0305": "¯",
  "\u030c": "ˇ", "\u0306": "˘", "\u0301": "´", "\u0300": "`", "\u030a": "˚", "\u20d6": "←",
};
const SPACE_EM: Record<string, string> = { "\u2009": "0.1667em", "\u2005": "0.2222em", "\u2004": "0.2778em", "\u2003": "1em", "\u2002": "0.5em", "\u200a": "0.0833em" };
const PUNCT = new Set([",", ";", ":", ".", "!", "?"]);

const mesc = (s: string) => s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
const row = (parts: string[]) => (parts.length === 1 ? parts[0] : `<mrow>${parts.join("")}</mrow>`);

function mathmlRun(n: Extract<MNode, { k: "t" }>): string[] {
  if (n.nor) return [`<mtext${n.sty === "b" ? ' mathvariant="bold"' : ""}>${mesc(n.s)}</mtext>`];
  const out: string[] = [];
  const variant = n.scr === "double-struck" ? "double-struck" : n.scr === "script" ? "script" : n.scr === "fraktur" ? "fraktur"
    : n.scr === "sans-serif" ? "sans-serif" : n.scr === "monospace" ? "monospace" : n.sty === "b" ? "bold" : n.sty === "bi" ? "bold-italic"
      : n.sty === "p" ? "normal" : "";
  const chars = [...n.s];
  // Upright letters (sty p, e.g. a function name or \mathrm) stay one identifier.
  if (variant === "normal" && /^[A-Za-z]+$/.test(n.s)) return [`<mi${n.s.length === 1 ? ' mathvariant="normal"' : ""}>${mesc(n.s)}</mi>`];
  for (let i = 0; i < chars.length; i++) {
    const ch = chars[i];
    if (SPACE_EM[ch]) { out.push(`<mspace width="${SPACE_EM[ch]}"/>`); continue; }
    if (ch === " ") continue;
    if (/[0-9]/.test(ch)) {
      let num = ch;
      while (i + 1 < chars.length && /[0-9.,]/.test(chars[i + 1]) && /[0-9]/.test(chars[i + 2] ?? "")) num += chars[++i];
      out.push(`<mn>${num}</mn>`);
      continue;
    }
    if (/\p{L}/u.test(ch) || ch === "∞" || ch === "∂" || ch === "∇" || ch === "∅") {
      const v = variant && variant !== "normal" ? variant : variant === "normal" || UPRIGHT_GREEK.test(ch) || /[∞∂∇∅]/.test(ch) ? "normal" : "";
      out.push(`<mi${v ? ` mathvariant="${v}"` : ""}>${mesc(ch)}</mi>`);
      continue;
    }
    if (ch === "′") { out.push("<mo>′</mo>"); continue; }
    out.push(PUNCT.has(ch) ? `<mo separator="true">${mesc(ch)}</mo>` : `<mo>${mesc(ch)}</mo>`);
  }
  return out;
}

function mml(s: Seq | undefined, display: boolean): string {
  const parts: string[] = [];
  for (const n of s ?? []) parts.push(...mathmlNode(n, display));
  return row(parts.length ? parts : ["<mrow/>"]);
}

function mathmlNode(n: MNode, display: boolean): string[] {
  switch (n.k) {
    case "t": return mathmlRun(n);
    case "frac":
      if (n.type === "lin") return [mml(n.num, display), "<mo>/</mo>", mml(n.den, display)];
      return [`<mfrac${n.type === "noBar" ? ' linethickness="0"' : ""}>${mml(n.num, display)}${mml(n.den, display)}</mfrac>`];
    case "scr": {
      const b = mml(n.base, display);
      if (n.sub && n.sup) return [`<msubsup>${b}${mml(n.sub, display)}${mml(n.sup, display)}</msubsup>`];
      if (n.sup) return [`<msup>${b}${mml(n.sup, display)}</msup>`];
      return [`<msub>${b}${mml(n.sub, display)}</msub>`];
    }
    case "rad": return [n.deg?.length ? `<mroot>${mml(n.e, display)}${mml(n.deg, display)}</mroot>` : `<msqrt>${mml(n.e, display)}</msqrt>`];
    case "d": {
      const parts: string[] = [];
      if (n.beg) parts.push(`<mo fence="true" stretchy="true">${mesc(n.beg)}</mo>`);
      n.items.forEach((x, i) => { if (i) parts.push('<mo stretchy="true">|</mo>'); parts.push(mml(x, display)); });
      if (n.end) parts.push(`<mo fence="true" stretchy="true">${mesc(n.end)}</mo>`);
      return [`<mrow>${parts.join("")}</mrow>`];
    }
    case "nary": {
      const op = `<mo largeop="true" movablelimits="false">${mesc(n.chr)}</mo>`;
      const under = n.und && display;
      let head = op;
      if (n.sub && n.sup) head = under ? `<munderover>${op}${mml(n.sub, display)}${mml(n.sup, display)}</munderover>` : `<msubsup>${op}${mml(n.sub, display)}${mml(n.sup, display)}</msubsup>`;
      else if (n.sub) head = under ? `<munder>${op}${mml(n.sub, display)}</munder>` : `<msub>${op}${mml(n.sub, display)}</msub>`;
      else if (n.sup) head = under ? `<mover>${op}${mml(n.sup, display)}</mover>` : `<msup>${op}${mml(n.sup, display)}</msup>`;
      // Word's display sums and integrals take 0.08 em more above and below (word-step4.docx).
      if (display) head = `<mrow style="padding-block:0.08em">${head}</mrow>`;
      return [head, mml(n.e, display)];
    }
    case "func": return [mml(n.name, display), "<mo>\u2061</mo>", mml(n.e, display)];
    case "acc": return [`<mover accent="true">${mml(n.e, display)}<mo stretchy="true">${mesc(ACCENT_GLYPH[n.chr] ?? n.chr)}</mo></mover>`];
    case "bar": return [n.top ? `<mover accent="true">${mml(n.e, display)}<mo stretchy="true">‾</mo></mover>` : `<munder accentunder="true">${mml(n.e, display)}<mo stretchy="true">_</mo></munder>`];
    case "lim": return [n.upper ? `<mover>${mml(n.e, display)}${mml(n.lim, display)}</mover>` : `<munder>${mml(n.e, display)}${mml(n.lim, display)}</munder>`];
    case "group": return [n.top ? `<mover>${mml(n.e, display)}<mo stretchy="true">${n.chr}</mo></mover>` : `<munder>${mml(n.e, display)}<mo stretchy="true">${n.chr}</mo></munder>`];
    case "m": {
      const cols = Math.max(1, ...n.rows.map((r) => r.length));
      return [`<mtable>${n.rows.map((r) => `<mtr>${Array.from({ length: cols }, (_x, i) => `<mtd>${mml(r[i], display)}</mtd>`).join("")}</mtr>`).join("")}</mtable>`];
    }
    // Alignment marks (&) are not shown.
    case "eqArr": return [`<mtable>${n.rows.map((r) => `<mtr><mtd>${mml(r.map((x) => (x.k === "t" ? { ...x, s: x.s.replace(/&/g, "") } : x)), display)}</mtd></mtr>`).join("")}</mtable>`];
    case "box": return [`<mrow style="border:1px solid currentColor;padding:1px 2px">${mml(n.e, display)}</mrow>`];
  }
}

/** MathML for showing an equation (display: alone in its paragraph, in display style). */
export function latexToMathml(latex: string, display: boolean): string {
  const rows = lines(parseLatex(latex));
  const body = rows.length > 1 ? mathmlNode({ k: "eqArr", rows }, display).join("") : mml(rows[0], display);
  return `<math${display ? ' displaystyle="true"' : ""}>${body}</math>`;
}

// ----- OMML → LaTeX ------------------------------------------------------------------------------

const REVERSE = new Map<string, string>();
for (const [k, v] of Object.entries(SYMBOLS)) if (!REVERSE.has(v)) REVERSE.set(v, `\\${k}`);
REVERSE.set("−", "-");
REVERSE.set("∗", "*");
REVERSE.set("′", "'");
REVERSE.set("|", "|");
REVERSE.set("…", "\\ldots");
const REVERSE_NARY = new Map(Object.entries(NARY).map(([k, v]) => [v, `\\${k}`]));
const REVERSE_ACC = new Map<string, string>([["\u0302", "hat"], ["\u0303", "tilde"], ["\u0307", "dot"], ["\u0308", "ddot"], ["\u20d7", "vec"],
  ["\u0305", "bar"], ["\u030c", "check"], ["\u0306", "breve"], ["\u0301", "acute"], ["\u0300", "grave"], ["\u030a", "mathring"], ["\u20d6", "overleftarrow"]]);
const REVERSE_SPACES = new Map<string, string>([["\u2009", "\\,"], ["\u2005", "\\:"], ["\u2004", "\\;"], ["\u2003", "\\quad "], ["\u2002", "\\enspace "], ["\u200a", "\\,"]]);
const LEFT: Record<string, string> = { "(": "(", ")": ")", "[": "[", "]": "]", "{": "\\{", "}": "\\}", "|": "|", "‖": "\\Vert", "⟨": "\\langle", "⟩": "\\rangle",
  "⌊": "\\lfloor", "⌋": "\\rfloor", "⌈": "\\lceil", "⌉": "\\rceil", "": ".", "〈": "\\langle", "〉": "\\rangle" };

class Unsupported extends Error {}

const kidsOf = (el: Element) => Array.from(el.children);
const mchild = (el: Element, name: string) => kidsOf(el).find((c) => c.localName === name);
const mval = (el: Element | undefined, name: string): string | null => {
  const c = el ? mchild(el, name) : undefined;
  if (!c) return null;
  return c.getAttributeNS(M_NS, "val") ?? c.getAttribute("m:val") ?? "";
};
const on = (v: string | null) => v !== null && v !== "0" && v !== "off" && v !== "false";

/** Escapes text of a math run for LaTeX. */
function latexText(s: string): string {
  let out = "";
  for (const ch of s) {
    if (REVERSE_SPACES.has(ch)) { out += REVERSE_SPACES.get(ch); continue; }
    if (ch === " ") continue;
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
/** A script base: braces unless it is a single character or command. */
function base(s: string): string {
  const t = s.trim();
  if ([...t].length === 1 || /^\\[A-Za-z]+\s*$/.test(t)) return t;
  return grp(t);
}

function seqLatex(el: Element | undefined): string {
  if (!el) return "";
  let out = "";
  for (const c of kidsOf(el)) {
    const s = nodeLatex(c);
    // A command name followed by a letter needs a separating space.
    if (/\\[A-Za-z]+$/.test(out) && /^[A-Za-z]/.test(s)) out += " ";
    out += s;
  }
  return out;
}

function nodeLatex(el: Element): string {
  const name = el.localName;
  if (el.namespaceURI !== M_NS) {
    if (name === "bookmarkStart" || name === "bookmarkEnd" || name === "proofErr" || name === "rPr") return "";
    if (name === "r") return latexText(Array.from(el.getElementsByTagName("*")).filter((x) => x.localName === "t").map((x) => x.textContent ?? "").join(""));
    throw new Unsupported(name);
  }
  switch (name) {
    case "r": {
      const t = kidsOf(el).filter((x) => x.localName === "t").map((x) => x.textContent ?? "").join("");
      const pr = mchild(el, "rPr");
      if (pr && mchild(pr, "nor")) return `\\text{${t.replace(/[{}\\]/g, "")}}`;
      const sty = mval(pr, "sty"), scr = mval(pr, "scr");
      const body = latexText(t);
      if (scr) {
        const cmd = ({ "double-struck": "mathbb", script: "mathcal", fraktur: "mathfrak", "sans-serif": "mathsf", monospace: "mathtt" } as Record<string, string>)[scr];
        if (cmd) return `\\${cmd}{${body}}`;
      }
      if (sty === "p" && /[A-Za-z]/.test(t)) return FUNCTIONS.has(t) ? `\\${t}` : `\\mathrm{${body}}`;
      if (sty === "b") return `\\mathbf{${body}}`;
      if (sty === "bi") return `\\boldsymbol{${body}}`;
      return body;
    }
    case "f": {
      const type = mval(mchild(el, "fPr"), "type");
      const num = seqLatex(mchild(el, "num")), den = seqLatex(mchild(el, "den"));
      if (type === "lin") return `${base(num)}/${base(den)}`;
      if (type === "noBar") return `\\genfrac{}{}{0pt}{}{${num}}{${den}}`;
      return `\\frac{${num}}{${den}}`;
    }
    case "sSup": return `${base(seqLatex(mchild(el, "e")))}^{${seqLatex(mchild(el, "sup"))}}`;
    case "sSub": return `${base(seqLatex(mchild(el, "e")))}_{${seqLatex(mchild(el, "sub"))}}`;
    case "sSubSup": return `${base(seqLatex(mchild(el, "e")))}_{${seqLatex(mchild(el, "sub"))}}^{${seqLatex(mchild(el, "sup"))}}`;
    case "sPre": return `{}_{${seqLatex(mchild(el, "sub"))}}^{${seqLatex(mchild(el, "sup"))}}${base(seqLatex(mchild(el, "e")))}`;
    case "rad": {
      const hide = on(mval(mchild(el, "radPr"), "degHide"));
      const deg = seqLatex(mchild(el, "deg"));
      return hide || !deg ? `\\sqrt{${seqLatex(mchild(el, "e"))}}` : `\\sqrt[${deg}]{${seqLatex(mchild(el, "e"))}}`;
    }
    case "d": {
      const pr = mchild(el, "dPr");
      const beg = mval(pr, "begChr") ?? "(", end = mval(pr, "endChr") ?? ")", sep = mval(pr, "sepChr") ?? "|";
      const items = kidsOf(el).filter((x) => x.localName === "e");
      // \binom and the matrix environments.
      if (items.length === 1) {
        const only = kidsOf(items[0]);
        if (only.length === 1 && only[0].localName === "f" && mval(mchild(only[0], "fPr"), "type") === "noBar" && beg === "(" && end === ")") {
          return `\\binom{${seqLatex(mchild(only[0], "num"))}}{${seqLatex(mchild(only[0], "den"))}}`;
        }
        if (only.length === 1 && only[0].localName === "m") {
          const env = ({ "()": "pmatrix", "[]": "bmatrix", "{}": "Bmatrix", "||": "vmatrix", "‖‖": "Vmatrix" } as Record<string, string>)[beg + end];
          if (env) return matrixLatex(only[0], env);
        }
        if (only.length === 1 && only[0].localName === "eqArr" && beg === "{" && end === "") {
          return `\\begin{cases}${kidsOf(only[0]).filter((x) => x.localName === "e").map((x) => seqLatex(x)).join("\\\\")}\\end{cases}`;
        }
      }
      const l = LEFT[beg] ?? latexText(beg), r = LEFT[end] ?? latexText(end);
      return `\\left${l}${items.map((x) => seqLatex(x)).join(`\\middle${LEFT[sep] ?? sep}`)}\\right${r}`;
    }
    case "nary": {
      const pr = mchild(el, "naryPr");
      const chr = mval(pr, "chr") ?? "∫";
      const cmd = REVERSE_NARY.get(chr) ?? `\\mathop{${latexText(chr)}}`;
      const loc = mval(pr, "limLoc");
      const lim = cmd.includes("int") && loc === "undOvr" ? "\\limits" : "";
      const sub = on(mval(pr, "subHide")) ? "" : seqLatex(mchild(el, "sub"));
      const sup = on(mval(pr, "supHide")) ? "" : seqLatex(mchild(el, "sup"));
      return `${cmd}${lim}${sub ? `_{${sub}}` : ""}${sup ? `^{${sup}}` : ""}${grp(seqLatex(mchild(el, "e")))}`;
    }
    case "func": {
      const fn = mchild(el, "fName");
      let nm = seqLatex(fn);
      const plain = fn ? kidsOf(fn).map((x) => x.textContent ?? "").join("") : "";
      if (fn && kidsOf(fn).length === 1 && kidsOf(fn)[0].localName === "r" && !FUNCTIONS.has(plain)) nm = `\\operatorname{${plain}}`;
      return `${nm}${grp(seqLatex(mchild(el, "e")))}`;
    }
    case "acc": {
      const chr = mval(mchild(el, "accPr"), "chr") ?? "\u0302";
      const cmd = REVERSE_ACC.get(chr);
      if (!cmd) throw new Unsupported(`acc ${chr}`);
      return `\\${cmd}{${seqLatex(mchild(el, "e"))}}`;
    }
    case "bar": return `\\${mval(mchild(el, "barPr"), "pos") === "top" ? "overline" : "underline"}{${seqLatex(mchild(el, "e"))}}`;
    case "limLow": case "limUpp": {
      const eEl = mchild(el, "e");
      const txt = eEl ? kidsOf(eEl).map((x) => x.textContent ?? "").join("") : "";
      const lim = seqLatex(mchild(el, "lim"));
      if (name === "limLow" && LIMIT_FUNCTIONS.has(txt)) return `\\${txt}_{${lim}}`;
      const inner = eEl && kidsOf(eEl).length === 1 && kidsOf(eEl)[0].localName === "groupChr" ? seqLatex(eEl) : null;
      if (inner) return `${inner}${name === "limLow" ? "_" : "^"}{${lim}}`;
      return `\\${name === "limLow" ? "underset" : "overset"}{${lim}}{${seqLatex(eEl)}}`;
    }
    case "groupChr": {
      const pr = mchild(el, "groupChrPr");
      const chr = mval(pr, "chr") ?? "⏟";
      const top = (mval(pr, "pos") ?? "bot") === "top";
      if (chr === "⏟" || chr === "⏞") return `\\${top ? "overbrace" : "underbrace"}{${seqLatex(mchild(el, "e"))}}`;
      return `\\${top ? "overset" : "underset"}{${latexText(chr)}}{${seqLatex(mchild(el, "e"))}}`;
    }
    case "m": return matrixLatex(el, "matrix");
    case "eqArr": return `\\begin{aligned}${kidsOf(el).filter((x) => x.localName === "e").map((x) => seqLatex(x)).join("\\\\")}\\end{aligned}`;
    case "borderBox": return `\\boxed{${seqLatex(mchild(el, "e"))}}`;
    case "box": return seqLatex(mchild(el, "e"));
    case "phant": return `\\phantom{${seqLatex(mchild(el, "e"))}}`;
    case "e": return seqLatex(el);
    case "ctrlPr": case "argPr": case "oMathParaPr": return "";
    case "oMath": return seqLatex(el);
    default: throw new Unsupported(name);
  }
}

function matrixLatex(m: Element, env: string): string {
  const rows = kidsOf(m).filter((x) => x.localName === "mr").map((r) => kidsOf(r).filter((x) => x.localName === "e").map((x) => seqLatex(x)).join("&"));
  return `\\begin{${env}}${rows.join("\\\\")}\\end{${env}}`;
}

/** LaTeX of an m:oMath or m:oMathPara element, or null when it uses something unsupported. */
export function ommlToLatex(el: Element): string | null {
  try {
    if (el.localName === "oMathPara") {
      return kidsOf(el).filter((x) => x.localName === "oMath").map((x) => seqLatex(x)).join("\\\\");
    }
    return seqLatex(el);
  } catch (err) {
    if (err instanceof Unsupported) return null;
    throw err;
  }
}
