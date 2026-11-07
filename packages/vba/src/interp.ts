// Our VBA interpreter (D-079): the language as Office macros in school use it, run against an object model the
// app gives (RaabePresent's Application/ActivePresentation, RaabeCalc's later). Statements: Sub, Function, Dim,
// ReDim (Preserve), Const, Set, assignment, Call, If/ElseIf/Else (block and one-line), Select Case, For/Next,
// For Each, Do/Loop (While/Until, before or after), While/Wend, With, Exit Sub/Function/For/Do, On Error Resume
// Next / GoTo 0, Option lines; expressions with VBA's operators and precedence (^, unary -, * /, \, Mod, + -, &,
// comparisons and Like, Not, And, Or, Xor, Eqv, Imp) and its built-in functions. A macro cannot reach outside the
// app: there is no Shell, no file access, no CreateObject, no Declare (D-066).

export type VbaValue = string | number | boolean | null | undefined | Date | VbaArray | VbaObject;
export interface VbaArray { kind: "array"; lb: number; items: VbaValue[] }
/** An object of the app's model: properties, methods and (for collections) items. */
export interface VbaObject {
  kind: "object";
  get(name: string, args: VbaValue[]): VbaValue;
  set?(name: string, value: VbaValue, args: VbaValue[]): void;
  call?(name: string, args: VbaValue[]): VbaValue;
  /** For Each over a collection. */
  items?(): VbaValue[];
  /** The object's default member (Item for collections, Value, Text …). */
  default?(args: VbaValue[]): VbaValue;
}
export interface VbaHost {
  /** Global names of the app (Application, ActivePresentation, ActiveSheet …); undefined when unknown. */
  global(name: string): VbaValue;
  msgBox(prompt: string, buttons: number, title: string): Promise<number>;
  inputBox(prompt: string, title: string, def: string): Promise<string>;
}

export class VbaError extends Error {
  constructor(message: string, readonly number = 5, readonly line = 0) { super(message); }
}

// ----- lexer ------------------------------------------------------------------------------------------------

type Tok = { t: "id" | "num" | "str" | "op" | "nl" | "date"; v: string; line: number };

function lex(src: string): Tok[] {
  const out: Tok[] = [];
  let line = 1;
  const lines = src.replace(/\r\n?/g, "\n").split("\n");
  // Line continuations (" _" at the end) join lines.
  const joined: { text: string; line: number }[] = [];
  let acc = "", start = 1;
  lines.forEach((l, i) => {
    if (/\s_\s*$/.test(l)) { if (!acc) start = i + 1; acc += l.replace(/\s_\s*$/, " "); return; }
    joined.push({ text: acc + l, line: acc ? start : i + 1 });
    acc = "";
  });
  for (const { text, line: ln } of joined) {
    line = ln;
    let i = 0;
    while (i < text.length) {
      const c = text[i];
      if (c === " " || c === "\t") { i++; continue; }
      if (c === "'") break;
      if (/^rem\b/i.test(text.slice(i)) && (i === 0 || /[\s:]/.test(text[i - 1]))) break;
      if (c === ":") { out.push({ t: "nl", v: ":", line }); i++; continue; }
      if (c === '"') {
        let s = "";
        i++;
        while (i < text.length) { if (text[i] === '"') { if (text[i + 1] === '"') { s += '"'; i += 2; continue; } i++; break; } s += text[i++]; }
        out.push({ t: "str", v: s, line });
        continue;
      }
      if (c === "#") { const m = /^#([^#]+)#/.exec(text.slice(i)); if (m) { out.push({ t: "date", v: m[1], line }); i += m[0].length; continue; } }
      const num = /^(&H[0-9a-f]+&?|&O[0-7]+&?|\d*\.?\d+(e[+-]?\d+)?[!#@%&]?)/i.exec(text.slice(i));
      if (num && /[0-9&.]/.test(c) && !(c === "&" && !/^&[HO]/i.test(text.slice(i)))) { out.push({ t: "num", v: num[0], line }); i += num[0].length; continue; }
      const id = /^[A-Za-zÄÖÜäöüß_][\wÄÖÜäöüß]*[$%&!#@]?/.exec(text.slice(i));
      if (id) { out.push({ t: "id", v: id[0].replace(/[$%&!#@]$/, ""), line }); i += id[0].length; continue; }
      const op = /^(<>|<=|>=|:=|[-+*/\\^&=<>(),.!])/.exec(text.slice(i));
      if (op) { out.push({ t: "op", v: op[0], line }); i += op[0].length; continue; }
      throw new VbaError(`Unerwartetes Zeichen "${c}" in Zeile ${line}`, 1001, line);
    }
    out.push({ t: "nl", v: "\n", line });
  }
  return out;
}

// ----- syntax tree -----------------------------------------------------------------------------------------

type Expr =
  | { e: "lit"; v: VbaValue }
  | { e: "id"; name: string }
  | { e: "member"; obj: Expr | null; name: string }
  | { e: "call"; fn: Expr; args: (Expr | null)[] }
  | { e: "bin"; op: string; a: Expr; b: Expr }
  | { e: "un"; op: string; a: Expr }
  | { e: "new"; cls: string };
type Stmt = { line: number } & (
  | { s: "dim"; vars: { name: string; dims?: Expr[]; isNew?: boolean }[]; konst?: false }
  | { s: "const"; name: string; value: Expr }
  | { s: "redim"; preserve: boolean; vars: { name: string; dims: Expr[] }[] }
  | { s: "assign"; target: Expr; value: Expr; set: boolean }
  | { s: "expr"; value: Expr }
  | { s: "if"; branches: { cond: Expr; body: Stmt[] }[]; other: Stmt[] }
  | { s: "for"; v: string; from: Expr; to: Expr; step?: Expr; body: Stmt[] }
  | { s: "foreach"; v: string; coll: Expr; body: Stmt[] }
  | { s: "do"; cond?: Expr; until: boolean; post: boolean; body: Stmt[] }
  | { s: "select"; value: Expr; cases: { tests: ({ k: "eq"; v: Expr } | { k: "range"; a: Expr; b: Expr } | { k: "is"; op: string; v: Expr })[] | null; body: Stmt[] }[] }
  | { s: "with"; obj: Expr; body: Stmt[] }
  | { s: "exit"; what: string }
  | { s: "onerror"; resume: boolean });
export interface Proc { name: string; kind: "sub" | "function"; params: { name: string; byRef: boolean; optional: boolean; def?: Expr }[]; body: Stmt[]; isPrivate: boolean }

class Parser {
  i = 0;
  constructor(private toks: Tok[]) {}
  get cur(): Tok { return this.toks[this.i] ?? { t: "nl", v: "\n", line: 0 }; }
  is(v: string): boolean { const c = this.cur; return (c.t === "id" || c.t === "op") && c.v.toLowerCase() === v; }
  isAt(k: number, v: string): boolean { const c = this.toks[this.i + k]; return !!c && (c.t === "id" || c.t === "op") && c.v.toLowerCase() === v; }
  eat(v: string): boolean { if (this.is(v)) { this.i++; return true; } return false; }
  need(v: string): void { if (!this.eat(v)) throw new VbaError(`Erwartet: ${v} (Zeile ${this.cur.line})`, 1002, this.cur.line); }
  name(): string { const c = this.cur; if (c.t !== "id") throw new VbaError(`Erwartet: Bezeichner (Zeile ${c.line})`, 1003, c.line); this.i++; return c.v; }
  eol(): boolean { return this.cur.t === "nl"; }
  skipNl(): void { while (this.i < this.toks.length && this.cur.t === "nl") this.i++; }

  module(): { procs: Map<string, Proc>; decls: Stmt[] } {
    const procs = new Map<string, Proc>();
    const decls: Stmt[] = [];
    this.skipNl();
    while (this.i < this.toks.length) {
      if (this.is("option")) { while (!this.eol()) this.i++; this.skipNl(); continue; }
      if (this.is("attribute")) { while (!this.eol()) this.i++; this.skipNl(); continue; }
      let isPrivate = false;
      while (this.is("public") || this.is("private") || this.is("static") || this.is("global")) { if (this.is("private")) isPrivate = true; this.i++; }
      if (this.is("sub") || this.is("function")) { const p = this.proc(isPrivate); procs.set(p.name.toLowerCase(), p); this.skipNl(); continue; }
      if (this.is("declare")) throw new VbaError("Declare ist in RaabePresent nicht erlaubt (Makros bleiben in der App).", 1004, this.cur.line);
      if (this.is("dim") || this.is("const") || (this.cur.t === "id" && this.toks[this.i - 1]?.t !== "id")) { decls.push(this.statement()); this.skipNl(); continue; }
      this.statement();
      this.skipNl();
    }
    return { procs, decls };
  }

  proc(isPrivate: boolean): Proc {
    const kind = this.is("sub") ? "sub" : "function";
    this.i++;
    const name = this.name();
    const params: Proc["params"] = [];
    if (this.eat("(")) {
      while (!this.eat(")")) {
        let byRef = true, optional = false;
        for (;;) { if (this.eat("byval")) byRef = false; else if (this.eat("byref")) byRef = true; else if (this.eat("optional")) optional = true; else if (this.eat("paramarray")) { /* as an array */ } else break; }
        const pn = this.name();
        if (this.eat("(")) this.need(")");
        if (this.eat("as")) this.typeName();
        let def: Expr | undefined;
        if (this.eat("=")) def = this.expr();
        params.push({ name: pn, byRef, optional, def });
        this.eat(",");
      }
    }
    if (this.eat("as")) this.typeName();
    const body = this.block(() => this.is("end") && (this.isAt(1, "sub") || this.isAt(1, "function")));
    this.need("end"); this.i++;
    return { name, kind, params, body, isPrivate };
  }

  typeName(): void { this.eat("new"); this.name(); while (this.eat(".")) this.name(); if (this.eat("*")) this.expr(); }

  block(stop: () => boolean): Stmt[] {
    const out: Stmt[] = [];
    for (;;) {
      this.skipNl();
      if (this.i >= this.toks.length) throw new VbaError("Unerwartetes Ende des Moduls", 1005);
      if (stop()) return out;
      out.push(this.statement());
    }
  }

  statement(): Stmt {
    const line = this.cur.line;
    if (this.eat("dim") || this.eat("static") || this.eat("public") || this.eat("private")) {
      if (this.is("const")) return this.statement();
      const vars: { name: string; dims?: Expr[]; isNew?: boolean }[] = [];
      do {
        const name = this.name();
        let dims: Expr[] | undefined;
        if (this.eat("(")) { dims = []; while (!this.eat(")")) { const a = this.expr(); if (this.eat("to")) { const b = this.expr(); dims.push({ e: "bin", op: "range", a, b }); } else dims.push(a); this.eat(","); } }
        let isNew = false;
        if (this.eat("as")) { isNew = this.is("new"); this.typeName(); }
        vars.push({ name, dims, isNew });
      } while (this.eat(","));
      return { line, s: "dim", vars };
    }
    if (this.eat("const")) { const name = this.name(); if (this.eat("as")) this.typeName(); this.need("="); return { line, s: "const", name, value: this.expr() }; }
    if (this.eat("redim")) {
      const preserve = this.eat("preserve");
      const vars: { name: string; dims: Expr[] }[] = [];
      do { const name = this.name(); this.need("("); const dims: Expr[] = []; while (!this.eat(")")) { const a = this.expr(); if (this.eat("to")) { const b = this.expr(); dims.push({ e: "bin", op: "range", a, b }); } else dims.push(a); this.eat(","); } if (this.eat("as")) this.typeName(); vars.push({ name, dims }); } while (this.eat(","));
      return { line, s: "redim", preserve, vars };
    }
    if (this.eat("if")) {
      const cond = this.expr();
      this.need("then");
      if (!this.eol()) {
        // One-line If: statements until Else or the end of the line.
        const body: Stmt[] = [this.simple()];
        while (this.cur.v === ":" && !this.isAt(1, "else")) { this.i++; if (this.eol()) break; body.push(this.simple()); }
        const other: Stmt[] = [];
        if (this.cur.v === ":" && this.isAt(1, "else")) this.i++;
        if (this.eat("else")) { other.push(this.simple()); while (this.cur.v === ":") { this.i++; if (this.eol()) break; other.push(this.simple()); } }
        return { line, s: "if", branches: [{ cond, body }], other };
      }
      const branches = [{ cond, body: this.block(() => this.is("elseif") || this.is("else") || (this.is("end") && this.isAt(1, "if"))) }];
      let other: Stmt[] = [];
      for (;;) {
        if (this.eat("elseif")) { const c = this.expr(); this.need("then"); branches.push({ cond: c, body: this.block(() => this.is("elseif") || this.is("else") || (this.is("end") && this.isAt(1, "if"))) }); continue; }
        if (this.eat("else")) { other = this.block(() => this.is("end") && this.isAt(1, "if")); }
        this.need("end"); this.need("if");
        return { line, s: "if", branches, other };
      }
    }
    if (this.eat("for")) {
      if (this.eat("each")) { const v = this.name(); this.need("in"); const coll = this.expr(); const body = this.block(() => this.is("next")); this.need("next"); if (this.cur.t === "id") this.i++; return { line, s: "foreach", v, coll, body }; }
      const v = this.name(); this.need("="); const from = this.expr(); this.need("to"); const to = this.expr();
      const step = this.eat("step") ? this.expr() : undefined;
      const body = this.block(() => this.is("next"));
      this.need("next"); if (this.cur.t === "id") this.i++;
      return { line, s: "for", v, from, to, step, body };
    }
    if (this.eat("do")) {
      let cond: Expr | undefined, until = false, post = false;
      if (this.eat("while")) cond = this.expr(); else if (this.eat("until")) { cond = this.expr(); until = true; }
      const body = this.block(() => this.is("loop"));
      this.need("loop");
      if (!cond) { if (this.eat("while")) { cond = this.expr(); post = true; } else if (this.eat("until")) { cond = this.expr(); until = true; post = true; } }
      return { line, s: "do", cond, until, post, body };
    }
    if (this.eat("while")) { const cond = this.expr(); const body = this.block(() => this.is("wend")); this.need("wend"); return { line, s: "do", cond, until: false, post: false, body }; }
    if (this.eat("select")) {
      this.need("case");
      const value = this.expr();
      const cases: Extract<Stmt, { s: "select" }>["cases"] = [];
      this.skipNl();
      while (this.eat("case")) {
        let tests: Extract<Stmt, { s: "select" }>["cases"][number]["tests"] = [];
        if (this.eat("else")) tests = null;
        else do {
          if (this.eat("is")) { const op = this.cur.v; this.i++; tests!.push({ k: "is", op, v: this.expr() }); }
          else { const a = this.expr(); if (this.eat("to")) tests!.push({ k: "range", a, b: this.expr() }); else tests!.push({ k: "eq", v: a }); }
        } while (this.eat(","));
        cases.push({ tests, body: this.block(() => this.is("case") || (this.is("end") && this.isAt(1, "select"))) });
      }
      this.need("end"); this.need("select");
      return { line, s: "select", value, cases };
    }
    if (this.eat("with")) { const obj = this.expr(); const body = this.block(() => this.is("end") && this.isAt(1, "with")); this.need("end"); this.need("with"); return { line, s: "with", obj, body }; }
    return this.simple();
  }

  simple(): Stmt {
    const line = this.cur.line;
    if (this.eat("exit")) { const what = this.name().toLowerCase(); return { line, s: "exit", what }; }
    if (this.eat("on")) { this.need("error"); if (this.eat("resume")) { this.need("next"); return { line, s: "onerror", resume: true }; } this.need("goto"); this.expr(); return { line, s: "onerror", resume: false }; }
    if (this.eat("call")) { const value = this.expr(); return { line, s: "expr", value }; }
    if (this.is("end") && !this.isAt(1, "sub") && !this.isAt(1, "function") && !this.isAt(1, "if")) { this.i++; return { line, s: "exit", what: "end" }; }
    const set = this.eat("set");
    this.eat("let");
    const target = this.postfix();
    if (this.eat("=")) return { line, s: "assign", target, value: this.expr(), set };
    // A procedure call without parentheses: Name arg1, arg2.
    if (!this.eol() && !this.is("else") && this.cur.v !== ":") {
      const args: (Expr | null)[] = [];
      do { if (this.is(",")) { args.push(null); continue; } args.push(this.expr()); } while (this.eat(","));
      return { line, s: "expr", value: { e: "call", fn: target, args } };
    }
    return { line, s: "expr", value: target };
  }

  // Precedence: Imp < Eqv < Xor < Or < And < Not < comparisons < & < + - < Mod < \ < * / < unary - < ^
  expr(): Expr { return this.bin(0); }
  private static LEVELS: string[][] = [["imp"], ["eqv"], ["xor"], ["or"], ["and"]];
  bin(level: number): Expr {
    if (level < Parser.LEVELS.length) {
      let a = this.bin(level + 1);
      while (Parser.LEVELS[level].some((o) => this.is(o))) { const op = this.cur.v.toLowerCase(); this.i++; a = { e: "bin", op, a, b: this.bin(level + 1) }; }
      return a;
    }
    if (this.eat("not")) return { e: "un", op: "not", a: this.bin(level) };
    return this.compare();
  }
  compare(): Expr {
    let a = this.concat();
    while (["=", "<>", "<", ">", "<=", ">="].includes(this.cur.v) || this.is("is") || this.is("like")) { const op = this.cur.v.toLowerCase(); this.i++; a = { e: "bin", op, a, b: this.concat() }; }
    return a;
  }
  concat(): Expr { let a = this.additive(); while (this.cur.v === "&") { this.i++; a = { e: "bin", op: "&", a, b: this.additive() }; } return a; }
  additive(): Expr { let a = this.modulo(); while (this.cur.t === "op" && (this.cur.v === "+" || this.cur.v === "-")) { const op = this.cur.v; this.i++; a = { e: "bin", op, a, b: this.modulo() }; } return a; }
  modulo(): Expr { let a = this.intdiv(); while (this.is("mod")) { this.i++; a = { e: "bin", op: "mod", a, b: this.intdiv() }; } return a; }
  intdiv(): Expr { let a = this.mul(); while (this.cur.v === "\\") { this.i++; a = { e: "bin", op: "\\", a, b: this.mul() }; } return a; }
  mul(): Expr { let a = this.unary(); while (this.cur.t === "op" && (this.cur.v === "*" || this.cur.v === "/")) { const op = this.cur.v; this.i++; a = { e: "bin", op, a, b: this.unary() }; } return a; }
  unary(): Expr { if (this.cur.t === "op" && (this.cur.v === "-" || this.cur.v === "+")) { const op = this.cur.v; this.i++; const a = this.unary(); return op === "-" ? { e: "un", op: "neg", a } : a; } return this.power(); }
  power(): Expr { const a = this.postfix(); if (this.cur.v === "^") { this.i++; return { e: "bin", op: "^", a, b: this.unary() }; } return a; }
  postfix(): Expr {
    let x = this.primary();
    for (;;) {
      if (this.cur.v === "(" && this.cur.t === "op") {
        this.i++;
        const args: (Expr | null)[] = [];
        while (!this.eat(")")) { if (this.is(",")) { args.push(null); this.i++; continue; } if ((this.cur as Tok).t === "id" && this.isAt(1, ":=")) { this.i += 2; } args.push(this.expr()); this.eat(","); }
        x = { e: "call", fn: x, args };
        continue;
      }
      if (this.cur.v === "." || this.cur.v === "!") { this.i++; x = { e: "member", obj: x, name: this.name() }; continue; }
      return x;
    }
  }
  primary(): Expr {
    const c = this.cur;
    if (c.t === "num") { this.i++; return { e: "lit", v: parseNum(c.v) }; }
    if (c.t === "str") { this.i++; return { e: "lit", v: c.v }; }
    if (c.t === "date") { this.i++; return { e: "lit", v: new Date(c.v) }; }
    if (c.v === "(") { this.i++; const x = this.expr(); this.need(")"); return x; }
    if (c.v === ".") { this.i++; return { e: "member", obj: null, name: this.name() }; }
    if (this.eat("new")) return { e: "new", cls: this.name() };
    if (this.eat("true")) return { e: "lit", v: true };
    if (this.eat("false")) return { e: "lit", v: false };
    if (this.eat("nothing")) return { e: "lit", v: null };
    if (this.eat("empty")) return { e: "lit", v: undefined };
    if (this.eat("null")) return { e: "lit", v: null };
    if (c.t === "id") { this.i++; return { e: "id", name: c.v }; }
    throw new VbaError(`Unerwartet: "${c.v}" (Zeile ${c.line})`, 1006, c.line);
  }
}

function parseNum(s: string): number {
  const t = s.replace(/[!#@%&]$/, "");
  if (/^&H/i.test(t)) return parseInt(t.slice(2), 16);
  if (/^&O/i.test(t)) return parseInt(t.slice(2), 8);
  return Number(t);
}

/** A module's procedures and module-level declarations. */
export function parseModule(src: string): { procs: Map<string, Proc>; decls: Stmt[] } {
  return new Parser(lex(src)).module();
}

// ----- running ---------------------------------------------------------------------------------------------

class Exit { constructor(readonly what: string) {} }
type Ref = { get(): VbaValue; set(v: VbaValue): void };
class Scope {
  vars = new Map<string, VbaValue>();
  refs = new Map<string, Ref>();
  constructor(readonly parent: Scope | null) {}
}

const toNum = (v: VbaValue): number => {
  if (typeof v === "number") return v;
  if (typeof v === "boolean") return v ? -1 : 0;
  if (v === undefined || v === null) return 0;
  if (v instanceof Date) return v.getTime() / 86400000 + 25569;
  if (typeof v === "string") { const n = Number(v.trim().replace(",", ".")); if (v.trim() === "" || Number.isNaN(n)) throw new VbaError("Typen unverträglich", 13); return n; }
  throw new VbaError("Typen unverträglich", 13);
};
const toStr = (v: VbaValue): string => {
  if (typeof v === "string") return v;
  if (typeof v === "number") return String(v).replace(".", ",");
  if (typeof v === "boolean") return v ? "Wahr" : "Falsch";
  if (v === undefined || v === null) return "";
  if (v instanceof Date) return v.toLocaleString("de-DE");
  if ((v as VbaObject).kind === "object") { const d = (v as VbaObject).default?.([]); return d === undefined ? "" : toStr(d); }
  throw new VbaError("Typen unverträglich", 13);
};
const toBool = (v: VbaValue): boolean => (typeof v === "boolean" ? v : toNum(v) !== 0);
const isObj = (v: VbaValue): v is VbaObject => !!v && typeof v === "object" && (v as VbaObject).kind === "object";
const isArr = (v: VbaValue): v is VbaArray => !!v && typeof v === "object" && (v as VbaArray).kind === "array";
/** Rounds as VBA does (banker's rounding) for integer conversions. */
const bankers = (n: number) => { const f = Math.floor(n), d = n - f; return d > 0.5 ? f + 1 : d < 0.5 ? f : f % 2 === 0 ? f : f + 1; };

function makeArray(dims: [number, number][]): VbaArray {
  const [lb, ub] = dims[0];
  return { kind: "array", lb, items: Array.from({ length: Math.max(0, ub - lb + 1) }, () => (dims.length > 1 ? makeArray(dims.slice(1)) : undefined)) };
}

function like(s: string, p: string): boolean {
  let re = "";
  for (let i = 0; i < p.length; i++) {
    const c = p[i];
    if (c === "*") re += ".*"; else if (c === "?") re += "."; else if (c === "#") re += "\\d";
    else if (c === "[") { const j = p.indexOf("]", i); re += "[" + p.slice(i + 1, j).replace(/^!/, "^") + "]"; i = j; }
    else re += c.replace(/[.+^${}()|\\]/g, "\\$&");
  }
  return new RegExp(`^${re}$`, "s").test(s);
}

export class VbaRuntime {
  private procs = new Map<string, Proc>();
  private moduleScope = new Scope(null);
  private steps = 0;
  /** Debug.Print lines. */
  readonly printed: string[] = [];

  constructor(private host: VbaHost, sources: string[], private limit = 5_000_000) {
    for (const src of sources) {
      const { procs, decls } = parseModule(src);
      for (const [k, p] of procs) this.procs.set(k, p);
      this.pendingDecls.push(...decls);
    }
  }
  private pendingDecls: Stmt[] = [];

  /** The macros a user can run (public Subs without parameters), as PowerPoint's Makros dialog lists them. */
  macros(): string[] {
    return [...this.procs.values()].filter((p) => p.kind === "sub" && !p.isPrivate && p.params.every((x) => x.optional)).map((p) => p.name);
  }

  async run(name: string, args: VbaValue[] = []): Promise<VbaValue> {
    if (this.pendingDecls.length) { const d = this.pendingDecls; this.pendingDecls = []; await this.exec(d, this.moduleScope, { withObj: [], resume: false }); }
    const p = this.procs.get(name.toLowerCase().replace(/^.*\./, ""));
    if (!p) throw new VbaError(`Das Makro "${name}" wurde nicht gefunden.`, 35);
    this.steps = 0;
    return this.invoke(p, args.map((v) => ({ e: "lit", v }) as Expr), this.moduleScope, { withObj: [], resume: false });
  }

  private async invoke(p: Proc, argExprs: (Expr | null)[], caller: Scope, ctx: Ctx): Promise<VbaValue> {
    const scope = new Scope(this.moduleScope);
    for (let i = 0; i < p.params.length; i++) {
      const prm = p.params[i];
      const a = argExprs[i];
      if (!a) { scope.vars.set(prm.name.toLowerCase(), prm.def ? await this.eval(prm.def, scope, ctx) : undefined); continue; }
      // ByRef to a variable: the procedure works on the caller's variable.
      if (prm.byRef && a.e === "id" && this.lookupScope(a.name, caller)) {
        const owner = this.lookupScope(a.name, caller)!;
        const key = a.name.toLowerCase();
        scope.refs.set(prm.name.toLowerCase(), { get: () => owner.vars.get(key), set: (v) => owner.vars.set(key, v) });
      } else scope.vars.set(prm.name.toLowerCase(), await this.eval(a, caller, ctx));
    }
    if (p.kind === "function") scope.vars.set(p.name.toLowerCase(), undefined);
    try {
      await this.exec(p.body, scope, { withObj: [], resume: false });
    } catch (e) {
      if (!(e instanceof Exit) || (e.what !== "sub" && e.what !== "function")) throw e;
    }
    return p.kind === "function" ? scope.vars.get(p.name.toLowerCase()) : undefined;
  }

  private lookupScope(name: string, scope: Scope | null): Scope | null {
    const k = name.toLowerCase();
    for (let s = scope; s; s = s.parent) if (s.vars.has(k) || s.refs.has(k)) return s;
    return null;
  }

  private tick(line: number): void {
    if (++this.steps > this.limit) throw new VbaError("Das Makro wurde angehalten: es lief zu lange (Endlosschleife?).", 6, line);
  }

  private async exec(stmts: Stmt[], scope: Scope, ctx: Ctx): Promise<void> {
    for (const st of stmts) {
      this.tick(st.line);
      try {
        await this.stmt(st, scope, ctx);
      } catch (e) {
        if (e instanceof Exit) throw e;
        if (ctx.resume) continue;
        if (e instanceof VbaError && !e.line) throw new VbaError(e.message, e.number, st.line);
        throw e instanceof VbaError ? e : new VbaError(String((e as Error)?.message ?? e), 5, st.line);
      }
    }
  }

  private async stmt(st: Stmt, scope: Scope, ctx: Ctx): Promise<void> {
    switch (st.s) {
      case "dim":
        for (const v of st.vars) {
          let val: VbaValue = undefined;
          if (v.dims) val = v.dims.length ? makeArray(await this.bounds(v.dims, scope, ctx)) : { kind: "array", lb: 0, items: [] };
          scope.vars.set(v.name.toLowerCase(), val);
        }
        return;
      case "const": scope.vars.set(st.name.toLowerCase(), await this.eval(st.value, scope, ctx)); return;
      case "redim":
        for (const v of st.vars) {
          const dims = await this.bounds(v.dims, scope, ctx);
          const fresh = makeArray(dims);
          const owner = this.lookupScope(v.name, scope) ?? scope;
          const old = owner.vars.get(v.name.toLowerCase());
          if (st.preserve && isArr(old)) old.items.slice(0, fresh.items.length).forEach((x, i) => { fresh.items[i] = x; });
          owner.vars.set(v.name.toLowerCase(), fresh);
        }
        return;
      case "assign": await this.assign(st.target, await this.eval(st.value, scope, ctx), scope, ctx); return;
      case "expr": {
        const x = st.value;
        // A bare name is a call of a procedure without arguments.
        await this.eval(x.e === "id" || x.e === "member" ? { e: "call", fn: x, args: [] } : x, scope, ctx, true);
        return;
      }
      case "if":
        for (const b of st.branches) if (toBool(await this.eval(b.cond, scope, ctx))) { await this.exec(b.body, scope, ctx); return; }
        await this.exec(st.other, scope, ctx);
        return;
      case "for": {
        const from = toNum(await this.eval(st.from, scope, ctx)), to = toNum(await this.eval(st.to, scope, ctx));
        const step = st.step ? toNum(await this.eval(st.step, scope, ctx)) : 1;
        const set = (v: number) => this.assign({ e: "id", name: st.v }, v, scope, ctx);
        await set(from);
        try {
          for (let i = from; step >= 0 ? i <= to : i >= to; ) {
            await this.exec(st.body, scope, ctx);
            i = toNum(await this.eval({ e: "id", name: st.v }, scope, ctx)) + step;
            await set(i);
            this.tick(st.line);
          }
        } catch (e) { if (!(e instanceof Exit && e.what === "for")) throw e; }
        return;
      }
      case "foreach": {
        const coll = await this.eval(st.coll, scope, ctx);
        const list = isArr(coll) ? flat(coll) : isObj(coll) && coll.items ? coll.items() : [];
        try {
          for (const item of list) { await this.assign({ e: "id", name: st.v }, item, scope, ctx); await this.exec(st.body, scope, ctx); this.tick(st.line); }
        } catch (e) { if (!(e instanceof Exit && e.what === "for")) throw e; }
        return;
      }
      case "do":
        try {
          for (;;) {
            if (st.cond && !st.post) { const c = toBool(await this.eval(st.cond, scope, ctx)); if (st.until ? c : !c) break; }
            await this.exec(st.body, scope, ctx);
            this.tick(st.line);
            if (st.cond && st.post) { const c = toBool(await this.eval(st.cond, scope, ctx)); if (st.until ? c : !c) break; }
          }
        } catch (e) { if (!(e instanceof Exit && e.what === "do")) throw e; }
        return;
      case "select": {
        const v = await this.eval(st.value, scope, ctx);
        for (const c of st.cases) {
          let hit = c.tests === null;
          for (const t of c.tests ?? []) {
            if (t.k === "eq") hit = this.equal(v, await this.eval(t.v, scope, ctx));
            else if (t.k === "range") hit = this.cmp(v, await this.eval(t.a, scope, ctx)) >= 0 && this.cmp(v, await this.eval(t.b, scope, ctx)) <= 0;
            else hit = this.compareOp(t.op, v, await this.eval(t.v, scope, ctx));
            if (hit) break;
          }
          if (hit) { await this.exec(c.body, scope, ctx); return; }
        }
        return;
      }
      case "with": {
        const obj = await this.eval(st.obj, scope, ctx);
        await this.exec(st.body, scope, { ...ctx, withObj: [...ctx.withObj, obj] });
        return;
      }
      case "exit": throw new Exit(st.what === "end" ? "sub" : st.what === "function" ? "function" : st.what);
      case "onerror": ctx.resume = st.resume; return;
    }
  }

  private async bounds(dims: Expr[], scope: Scope, ctx: Ctx): Promise<[number, number][]> {
    const out: [number, number][] = [];
    for (const d of dims) {
      if (d.e === "bin" && d.op === "range") out.push([toNum(await this.eval(d.a, scope, ctx)), toNum(await this.eval(d.b, scope, ctx))]);
      else out.push([0, toNum(await this.eval(d, scope, ctx))]);
    }
    return out;
  }

  private async assign(target: Expr, value: VbaValue, scope: Scope, ctx: Ctx): Promise<void> {
    if (target.e === "id") {
      const k = target.name.toLowerCase();
      for (let s: Scope | null = scope; s; s = s.parent) {
        const r = s.refs.get(k);
        if (r) { r.set(value); return; }
        if (s.vars.has(k)) { s.vars.set(k, value); return; }
      }
      // A default property of a host object (Range = 5) is not supported for bare names: a new variable.
      scope.vars.set(k, value);
      return;
    }
    if (target.e === "call") {
      // Array element or a parameterised property (Cells(1, 2) = 5, a(3) = x).
      const base = target.fn;
      const args = await this.args(target.args, scope, ctx);
      if (base.e === "id") {
        const arr = await this.eval(base, scope, ctx);
        if (isArr(arr)) { setIn(arr, args.map(toNum), value); return; }
        if (isObj(arr)) { arr.set?.("", value, args); return; }
      }
      if (base.e === "member") {
        const obj = await this.objOf(base.obj, scope, ctx);
        if (isObj(obj)) { obj.set?.(base.name, value, args); return; }
      }
      throw new VbaError("Zuweisung nicht möglich", 424);
    }
    if (target.e === "member") {
      const obj = await this.objOf(target.obj, scope, ctx);
      if (!isObj(obj) || !obj.set) throw new VbaError(`Objekt erforderlich (.${target.name})`, 424);
      obj.set(target.name, value, []);
      return;
    }
    throw new VbaError("Zuweisung nicht möglich", 424);
  }

  private async objOf(e: Expr | null, scope: Scope, ctx: Ctx): Promise<VbaValue> {
    if (!e) { const w = ctx.withObj[ctx.withObj.length - 1]; if (w === undefined) throw new VbaError("Ungültiger Verweis: . ohne With", 91); return w; }
    return this.eval(e, scope, ctx);
  }

  private async args(list: (Expr | null)[], scope: Scope, ctx: Ctx): Promise<VbaValue[]> {
    const out: VbaValue[] = [];
    for (const a of list) out.push(a ? await this.eval(a, scope, ctx) : undefined);
    return out;
  }

  private equal(a: VbaValue, b: VbaValue): boolean { return this.cmp(a, b) === 0; }
  private cmp(a: VbaValue, b: VbaValue): number {
    if (typeof a === "string" && typeof b === "string") return a < b ? -1 : a > b ? 1 : 0;
    if ((typeof a === "string" && Number.isNaN(Number(a))) || (typeof b === "string" && Number.isNaN(Number(b)))) { const x = toStr(a), y = toStr(b); return x < y ? -1 : x > y ? 1 : 0; }
    const x = toNum(a), y = toNum(b);
    return x < y ? -1 : x > y ? 1 : 0;
  }
  private compareOp(op: string, a: VbaValue, b: VbaValue): boolean {
    const c = this.cmp(a, b);
    return op === "=" ? c === 0 : op === "<>" ? c !== 0 : op === "<" ? c < 0 : op === ">" ? c > 0 : op === "<=" ? c <= 0 : c >= 0;
  }

  private async eval(x: Expr, scope: Scope, ctx: Ctx, statement = false): Promise<VbaValue> {
    switch (x.e) {
      case "lit": return x.v;
      case "id": {
        const k = x.name.toLowerCase();
        for (let s: Scope | null = scope; s; s = s.parent) {
          const r = s.refs.get(k);
          if (r) return r.get();
          if (s.vars.has(k)) return s.vars.get(k);
        }
        const p = this.procs.get(k);
        if (p) return this.invoke(p, [], scope, ctx);
        const b = await this.builtin(k, [], statement);
        if (b !== NOT_BUILTIN) return b;
        const g = this.host.global(x.name);
        if (g !== undefined) return g;
        const c = CONSTANTS[k];
        if (c !== undefined) return c;
        return undefined; // an undeclared variable is Empty (Option Explicit is not enforced)
      }
      case "member": {
        const obj = await this.objOf(x.obj, scope, ctx);
        if (!isObj(obj)) throw new VbaError(`Objekt erforderlich (.${x.name})`, 424);
        return obj.get(x.name, []);
      }
      case "new": throw new VbaError(`Neue Objekte (New ${x.cls}) gibt es in RaabePresent nicht.`, 429);
      case "un": {
        const a = await this.eval(x.a, scope, ctx);
        if (x.op === "neg") return -toNum(a);
        return typeof a === "boolean" ? !a : ~toNum(a);
      }
      case "bin": {
        if (x.op === "range") throw new VbaError("Ungültiger Bereich", 5);
        const a = await this.eval(x.a, scope, ctx), b = await this.eval(x.b, scope, ctx);
        switch (x.op) {
          case "+": return typeof a === "string" && typeof b === "string" ? a + b : toNum(a) + toNum(b);
          case "-": return toNum(a) - toNum(b);
          case "*": return toNum(a) * toNum(b);
          case "/": { const d = toNum(b); if (d === 0) throw new VbaError("Division durch Null", 11); return toNum(a) / d; }
          case "\\": { const d = bankers(toNum(b)); if (d === 0) throw new VbaError("Division durch Null", 11); return Math.trunc(bankers(toNum(a)) / d); }
          case "mod": { const d = bankers(toNum(b)); if (d === 0) throw new VbaError("Division durch Null", 11); return bankers(toNum(a)) % d; }
          case "^": return Math.pow(toNum(a), toNum(b));
          case "&": return toStr(a) + toStr(b);
          case "=": case "<>": case "<": case ">": case "<=": case ">=": return this.compareOp(x.op, a, b);
          case "is": return a === b;
          case "like": return like(toStr(a), toStr(b));
          case "and": return typeof a === "boolean" && typeof b === "boolean" ? a && b : toNum(a) & toNum(b);
          case "or": return typeof a === "boolean" && typeof b === "boolean" ? a || b : toNum(a) | toNum(b);
          case "xor": return typeof a === "boolean" && typeof b === "boolean" ? a !== b : toNum(a) ^ toNum(b);
          case "eqv": return toBool(a) === toBool(b);
          case "imp": return !toBool(a) || toBool(b);
        }
        throw new VbaError(`Operator ${x.op}`, 5);
      }
      case "call": {
        const f = x.fn;
        if (f.e === "id") {
          const k = f.name.toLowerCase();
          // A variable: an array element or an object's default member.
          const owner = this.lookupScope(f.name, scope);
          if (owner) {
            const v = owner.refs.get(k)?.get() ?? owner.vars.get(k);
            const args = await this.args(x.args, scope, ctx);
            if (isArr(v)) return getIn(v, args.map(toNum));
            if (isObj(v)) return v.default ? v.default(args) : v.get("Item", args);
            throw new VbaError(`"${f.name}" ist kein Datenfeld`, 13);
          }
          const p = this.procs.get(k);
          if (p) return this.invoke(p, x.args, scope, ctx);
          const args = await this.args(x.args, scope, ctx);
          const b = await this.builtin(k, args, statement);
          if (b !== NOT_BUILTIN) return b;
          const g = this.host.global(f.name);
          if (isObj(g)) return g.default ? g.default(args) : g.get("Item", args);
          throw new VbaError(`Sub oder Function nicht definiert: ${f.name}`, 35);
        }
        if (f.e === "member") {
          const obj = await this.objOf(f.obj, scope, ctx);
          if (!isObj(obj)) throw new VbaError(`Objekt erforderlich (.${f.name})`, 424);
          const args = await this.args(x.args, scope, ctx);
          if (obj.call) { const r = obj.call(f.name, args); if (r !== NOT_A_METHOD) return r instanceof Promise ? await r : r; }
          // A collection property with an index (Slides(1), Shapes("Titel")): the collection's item.
          if (args.length) { const coll = obj.get(f.name, []); if (isObj(coll) && (coll.default || coll.items)) return coll.default ? coll.default(args) : coll.get("Item", args); }
          return obj.get(f.name, args);
        }
        const base = await this.eval(f, scope, ctx);
        const args = await this.args(x.args, scope, ctx);
        if (isArr(base)) return getIn(base, args.map(toNum));
        if (isObj(base)) return base.default ? base.default(args) : base.get("Item", args);
        throw new VbaError("Aufruf nicht möglich", 5);
      }
    }
  }

  /** VBA's built-in functions and statements (MsgBox, InputBox, Debug.Print through the host's Debug). */
  private async builtin(k: string, a: VbaValue[], statement: boolean): Promise<VbaValue | typeof NOT_BUILTIN> {
    const s = (i: number) => toStr(a[i]), n = (i: number) => toNum(a[i]);
    switch (k) {
      case "msgbox": { const r = await this.host.msgBox(s(0), a[1] === undefined ? 0 : n(1), a[2] === undefined ? "RaabePresent" : s(2)); return statement ? undefined : r; }
      case "inputbox": return this.host.inputBox(s(0), a[1] === undefined ? "RaabePresent" : s(1), a[2] === undefined ? "" : s(2));
      case "len": return isArr(a[0]) ? a[0].items.length : s(0).length;
      case "left": return s(0).slice(0, Math.max(0, n(1)));
      case "right": return n(1) <= 0 ? "" : s(0).slice(-n(1));
      case "mid": return a[2] === undefined ? s(0).slice(n(1) - 1) : s(0).substr(n(1) - 1, n(2));
      case "ucase": return s(0).toUpperCase();
      case "lcase": return s(0).toLowerCase();
      case "trim": return s(0).trim();
      case "ltrim": return s(0).replace(/^\s+/, "");
      case "rtrim": return s(0).replace(/\s+$/, "");
      case "instr": { if (a.length >= 3 && typeof a[0] === "number") return s(1).indexOf(s(2), n(0) - 1) + 1; return s(0).indexOf(s(1)) + 1; }
      case "instrrev": return s(0).lastIndexOf(s(1)) + 1;
      case "replace": return s(0).split(s(1)).join(s(2));
      case "space": return " ".repeat(Math.max(0, n(0)));
      case "string": return s(1).charAt(0).repeat(Math.max(0, n(0)));
      case "strreverse": return Array.from(s(0)).reverse().join("");
      case "strcomp": { const x = a[2] && n(2) === 1 ? s(0).toLowerCase() : s(0), y = a[2] && n(2) === 1 ? s(1).toLowerCase() : s(1); return x < y ? -1 : x > y ? 1 : 0; }
      case "chr": case "chrw": return String.fromCharCode(n(0));
      case "asc": case "ascw": return s(0).charCodeAt(0) || 0;
      case "cstr": case "str": return k === "str" && toNum(a[0]) >= 0 ? " " + toStr(a[0]).replace(",", ".") : toStr(a[0]);
      case "val": { const m = /^\s*[-+]?\d*\.?\d+(e[-+]?\d+)?/i.exec(s(0)); return m ? Number(m[0]) : 0; }
      case "cint": case "clng": case "clnglng": return bankers(n(0));
      case "cdbl": case "csng": case "ccur": case "cdec": return n(0);
      case "cbool": return toBool(a[0]);
      case "cdate": return new Date(s(0));
      case "int": return Math.floor(n(0));
      case "fix": return Math.trunc(n(0));
      case "abs": return Math.abs(n(0));
      case "sgn": return Math.sign(n(0));
      case "sqr": return Math.sqrt(n(0));
      case "exp": return Math.exp(n(0));
      case "log": return Math.log(n(0));
      case "sin": return Math.sin(n(0));
      case "cos": return Math.cos(n(0));
      case "tan": return Math.tan(n(0));
      case "atn": return Math.atan(n(0));
      case "round": { const d = a[1] === undefined ? 0 : n(1); const f = 10 ** d; return bankers(n(0) * f) / f; }
      case "rnd": return Math.random();
      case "randomize": return undefined;
      case "isnumeric": return a[0] !== "" && a[0] !== undefined && !Number.isNaN(Number(toStr(a[0]).replace(",", ".")));
      case "isempty": return a[0] === undefined;
      case "isnull": return a[0] === null;
      case "isarray": return isArr(a[0]);
      case "isobject": return isObj(a[0]) || a[0] === null;
      case "isdate": return a[0] instanceof Date || !Number.isNaN(Date.parse(s(0)));
      case "ubound": { const arr = a[0]; if (!isArr(arr)) throw new VbaError("Typen unverträglich", 13); return arr.lb + arr.items.length - 1; }
      case "lbound": { const arr = a[0]; if (!isArr(arr)) throw new VbaError("Typen unverträglich", 13); return arr.lb; }
      case "array": return { kind: "array", lb: 0, items: [...a] };
      case "split": return { kind: "array", lb: 0, items: s(0).split(a[1] === undefined ? " " : s(1)) };
      case "join": { const arr = a[0]; return isArr(arr) ? arr.items.map(toStr).join(a[1] === undefined ? " " : s(1)) : ""; }
      case "now": return new Date();
      case "date": return new Date(new Date().toDateString());
      case "time": return new Date();
      case "timer": { const d = new Date(); return d.getHours() * 3600 + d.getMinutes() * 60 + d.getSeconds() + d.getMilliseconds() / 1000; }
      case "year": return (a[0] instanceof Date ? a[0] : new Date(s(0))).getFullYear();
      case "month": return (a[0] instanceof Date ? a[0] : new Date(s(0))).getMonth() + 1;
      case "day": return (a[0] instanceof Date ? a[0] : new Date(s(0))).getDate();
      case "hour": return (a[0] instanceof Date ? a[0] : new Date(s(0))).getHours();
      case "minute": return (a[0] instanceof Date ? a[0] : new Date(s(0))).getMinutes();
      case "second": return (a[0] instanceof Date ? a[0] : new Date(s(0))).getSeconds();
      case "weekday": return (a[0] instanceof Date ? a[0] : new Date(s(0))).getDay() + 1;
      case "format": return format(a[0], a[1] === undefined ? "" : s(1));
      case "rgb": return (n(0) & 255) | ((n(1) & 255) << 8) | ((n(2) & 255) << 16);
      case "iif": return toBool(a[0]) ? a[1] : a[2];
      case "typename": return a[0] === undefined ? "Empty" : a[0] === null ? "Nothing" : typeof a[0] === "string" ? "String" : typeof a[0] === "boolean" ? "Boolean" : typeof a[0] === "number" ? (Number.isInteger(a[0]) ? "Integer" : "Double") : isArr(a[0]) ? "Variant()" : "Object";
      case "doevents": return 0;
      case "debug": return DEBUG(this.printed);
      case "shell": case "createobject": case "getobject": case "kill": case "mkdir": case "rmdir": case "filecopy": case "sendkeys": case "environ":
        throw new VbaError(`${k} ist in RaabePresent nicht erlaubt: Makros bleiben in der App.`, 70);
    }
    return NOT_BUILTIN;
  }
}

interface Ctx { withObj: VbaValue[]; resume: boolean }
const NOT_BUILTIN = Symbol("not builtin");
/** Returned by an object's call() when the name is a property, not a method. */
export const NOT_A_METHOD = Symbol("not a method") as unknown as VbaValue;

const CONSTANTS: Record<string, VbaValue> = {
  vbcrlf: "\r\n", vbcr: "\r", vblf: "\n", vbnewline: "\r\n", vbtab: "\t", vbnullstring: "",
  vbokonly: 0, vbokcancel: 1, vbabortretryignore: 2, vbyesnocancel: 3, vbyesno: 4, vbretrycancel: 5, vbcritical: 16, vbquestion: 32, vbexclamation: 48, vbinformation: 64,
  vbok: 1, vbcancel: 2, vbabort: 3, vbretry: 4, vbignore: 5, vbyes: 6, vbno: 7, vbtextcompare: 1, vbbinarycompare: 0,
  vbblack: 0, vbred: 255, vbgreen: 65280, vbyellow: 65535, vbblue: 16711680, vbmagenta: 16711935, vbcyan: 16776960, vbwhite: 16777215,
};

function DEBUG(printed: string[]): VbaObject {
  return { kind: "object", get: () => undefined, call: (name, args) => { if (name.toLowerCase() === "print") { printed.push(args.map(toStr).join(" ")); return undefined; } return NOT_A_METHOD; } };
}

function flat(a: VbaArray): VbaValue[] { return a.items.flatMap((x) => (isArr(x) ? flat(x) : [x])); }
function getIn(a: VbaArray, idx: number[]): VbaValue {
  let cur: VbaValue = a;
  for (const i of idx) { if (!isArr(cur)) throw new VbaError("Index außerhalb des gültigen Bereichs", 9); const arr: VbaArray = cur; const k: number = i - arr.lb; if (k < 0 || k >= arr.items.length) throw new VbaError("Index außerhalb des gültigen Bereichs", 9); cur = arr.items[k]; }
  return cur;
}
function setIn(a: VbaArray, idx: number[], v: VbaValue): void {
  let cur: VbaArray = a;
  idx.forEach((i, n) => {
    const k = i - cur.lb;
    if (k < 0 || k >= cur.items.length) throw new VbaError("Index außerhalb des gültigen Bereichs", 9);
    if (n === idx.length - 1) cur.items[k] = v; else cur = cur.items[k] as VbaArray;
  });
}

/** Format with the common named and numeric formats (German conventions). */
function format(v: VbaValue, f: string): string {
  const lf = f.toLowerCase();
  if (v instanceof Date || ["short date", "long date", "short time", "long time", "general date"].includes(lf) || /[dmyhns]{1,4}/i.test(f) && v instanceof Date) {
    const d = v instanceof Date ? v : new Date(toStr(v));
    if (lf === "short date" || !f) return d.toLocaleDateString("de-DE");
    if (lf === "long date") return d.toLocaleDateString("de-DE", { weekday: "long", day: "numeric", month: "long", year: "numeric" });
    if (lf === "short time") return d.toLocaleTimeString("de-DE", { hour: "2-digit", minute: "2-digit" });
    if (lf === "long time") return d.toLocaleTimeString("de-DE");
    const p = (x: number, l = 2) => String(x).padStart(l, "0");
    return f.replace(/yyyy|yy|mmmm|mmm|mm|dddd|ddd|dd|hh|nn|ss|m|d/gi, (t) => {
      switch (t.toLowerCase()) {
        case "yyyy": return String(d.getFullYear()); case "yy": return p(d.getFullYear() % 100);
        case "mmmm": return d.toLocaleDateString("de-DE", { month: "long" }); case "mmm": return d.toLocaleDateString("de-DE", { month: "short" });
        case "mm": return p(d.getMonth() + 1); case "m": return String(d.getMonth() + 1);
        case "dddd": return d.toLocaleDateString("de-DE", { weekday: "long" }); case "ddd": return d.toLocaleDateString("de-DE", { weekday: "short" });
        case "dd": return p(d.getDate()); case "d": return String(d.getDate());
        case "hh": return p(d.getHours()); case "nn": return p(d.getMinutes()); case "ss": return p(d.getSeconds());
      }
      return t;
    });
  }
  const n = toNum(v);
  if (lf === "percent") return (n * 100).toLocaleString("de-DE", { minimumFractionDigits: 2, maximumFractionDigits: 2 }) + " %";
  if (lf === "currency") return n.toLocaleString("de-DE", { style: "currency", currency: "EUR" });
  if (lf === "fixed") return n.toLocaleString("de-DE", { minimumFractionDigits: 2, maximumFractionDigits: 2, useGrouping: false });
  if (lf === "standard") return n.toLocaleString("de-DE", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  const m = /^([#0,]*)(?:\.([#0]+))?(%?)$/.exec(f);
  if (m) {
    const dec = (m[2] ?? "").length, min = (m[2] ?? "").replace(/#/g, "").length;
    const x = m[3] ? n * 100 : n;
    return x.toLocaleString("de-DE", { minimumFractionDigits: min, maximumFractionDigits: dec, useGrouping: m[1].includes(",") }) + (m[3] ? "%" : "");
  }
  return toStr(v);
}

export { toStr as vbaToString, toNum as vbaToNumber };
