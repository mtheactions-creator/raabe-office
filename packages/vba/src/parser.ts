// The VBA language as our interpreter reads it (D-079): lexer, syntax tree and parser. A module is parsed into its
// procedures (Sub, Function, Property Get/Let/Set), its module-level declarations (Dim, Public, Private, Const,
// Enum, Type, WithEvents variables), its events (Public Event), Implements lines, attributes (VB_Name,
// VB_PredeclaredId, a member's VB_UserMemId) and options (Explicit, Base, Compare). Every statement keeps its
// source line (the first line of a statement continued with " _"), which the debugger and error messages use.
// Conditional compilation (#If VBA7 Then … #Else … #End If) is resolved while lexing, as a 64-bit Office on
// Windows resolves it.

import { VbaError, type VbaValue } from "./values";

export type Tok = { t: "id" | "num" | "str" | "op" | "nl" | "date"; v: string; line: number; first?: boolean; br?: boolean };

/** The constants of conditional compilation, as in a 64-bit Office 365 on Windows. */
const CC: Record<string, number> = { vba7: -1, vba6: -1, win64: -1, win32: -1, win16: 0, mac: 0 };

/** Lines of inactive #If branches become empty (line numbers stay), #Const values are kept. */
function precompile(lines: string[]): string[] {
  const consts: Record<string, number> = { ...CC };
  const stack: { active: boolean; taken: boolean; outer: boolean }[] = [];
  const active = () => stack.every((s) => s.active);
  const cond = (s: string): boolean => {
    // A tiny evaluator: names, numbers, Not, And, Or, =, <>.
    const js = s.replace(/\bThen\b/i, "").trim()
      .replace(/[A-Za-z_]\w*/g, (w) => { const l = w.toLowerCase(); if (l === "not") return "!"; if (l === "and") return "&&"; if (l === "or") return "||"; return String(consts[l] ?? 0); })
      .replace(/<>/g, "!=").replace(/(^|[^!<>=])=(?!=)/g, "$1==");
    try { return !!Function(`"use strict";return (${js});`)(); } catch { throw new VbaError(`Ungültige Bedingung: #If ${s}`, 1010); }
  };
  return lines.map((l) => {
    const m = /^\s*#(if|elseif|else|end\s*if|const)\b(.*)$/i.exec(l);
    if (!m) return active() ? l : "";
    const k = m[1].toLowerCase().replace(/\s+/g, "");
    if (k === "if") { const outer = active(); const c = outer && cond(m[2]); stack.push({ active: c, taken: c, outer }); }
    else if (k === "elseif") { const top = stack[stack.length - 1]; if (top) { const c = top.outer && !top.taken && cond(m[2]); top.active = c; top.taken ||= c; } }
    else if (k === "else") { const top = stack[stack.length - 1]; if (top) { top.active = top.outer && !top.taken; top.taken = true; } }
    else if (k === "endif") stack.pop();
    else if (k === "const" && active()) { const c = /^\s*(\w+)\s*=\s*(.+)$/.exec(m[2]); if (c) consts[c[1].toLowerCase()] = cond(c[2]) ? -1 : Number(c[2]) || 0; }
    return "";
  });
}

export function lex(src: string): Tok[] {
  const out: Tok[] = [];
  const lines = precompile(src.replace(/\r\n?/g, "\n").split("\n"));
  // Line continuations (" _" at the end) join lines.
  const joined: { text: string; line: number }[] = [];
  let acc = "", start = 1;
  lines.forEach((l, i) => {
    if (/(^|\s)_\s*$/.test(l) && !/'/.test(l.replace(/"[^"]*"/g, ""))) { if (!acc) start = i + 1; acc += l.replace(/_\s*$/, " "); return; }
    joined.push({ text: acc + l, line: acc ? start : i + 1 });
    acc = "";
  });
  const NBSP = String.fromCharCode(160);
  const ID = /^[A-Za-zÄÖÜäöüß_][\wÄÖÜäöüß]*/;
  for (const { text, line } of joined) {
    let i = 0;
    let first = true;
    const push = (t: Tok) => { if (first) { t.first = true; first = false; } out.push(t); };
    while (i < text.length) {
      const c = text[i];
      if (c === " " || c === "\t" || c === NBSP) { i++; continue; }
      if (c === "'") break;
      if (/^rem\b/i.test(text.slice(i)) && (i === 0 || /[\s:]/.test(text[i - 1]))) break;
      if (c === ":" && text[i + 1] !== "=") { push({ t: "nl", v: ":", line }); i++; continue; }
      if (c === '"') {
        let s = "";
        i++;
        while (i < text.length) { if (text[i] === '"') { if (text[i + 1] === '"') { s += '"'; i += 2; continue; } i++; break; } s += text[i++]; }
        push({ t: "str", v: s, line });
        continue;
      }
      if (c === "#") {
        const m = /^#(\d{1,4}[-/.]\d{1,2}[-/.]\d{1,4}(?:\s+\d{1,2}:\d{2}(?::\d{2})?(?:\s*[AP]M)?)?|\d{1,2}:\d{2}(?::\d{2})?(?:\s*[AP]M)?)#/i.exec(text.slice(i));
        if (m) { push({ t: "date", v: m[1], line }); i += m[0].length; continue; }
        push({ t: "op", v: "#", line }); i++; continue;
      }
      if (c === "[") { const j = text.indexOf("]", i); if (j > i) { push({ t: "id", v: text.slice(i + 1, j), line, br: true }); i = j + 1; continue; } }
      const num = /^(&H[0-9a-f]+&?|&O[0-7]+&?|\d*\.?\d+(e[+-]?\d+)?[!#@%&]?)/i.exec(text.slice(i));
      if (num && /[0-9&.]/.test(c) && !(c === "&" && !/^&[HO]/i.test(text.slice(i))) && !(c === "." && !/^\.\d/.test(text.slice(i)))) { push({ t: "num", v: num[0], line }); i += num[0].length; continue; }
      const id = ID.exec(text.slice(i));
      if (id) {
        let w = id[0];
        i += w.length;
        // A type suffix (Left$, n%, x&); ! and # only when no name follows (rs!Feld is a bang access).
        if (/[$%&@]/.test(text[i] ?? "") || (/[!#]/.test(text[i] ?? "") && !/[A-Za-zÄÖÜäöüß_[\d]/.test(text[i + 1] ?? ""))) { if (!(text[i] === "&" && /[HO]/i.test(text[i + 1] ?? ""))) i++; }
        if (w === "_") continue;
        push({ t: "id", v: w, line });
        w = "";
        continue;
      }
      const op = /^(<>|<=|>=|:=|[-+*/\\^&=<>(),.!;])/.exec(text.slice(i));
      if (op) { push({ t: "op", v: op[0], line }); i += op[0].length; continue; }
      throw new VbaError(`Unerwartetes Zeichen "${c}" in Zeile ${line}`, 1001, line);
    }
    out.push({ t: "nl", v: "\n", line });
  }
  return out;
}

// ----- syntax tree -----------------------------------------------------------------------------------------

/** A variable reference passed ByRef from outside the language (the host's event arguments). */
export interface Ref { get(): VbaValue; set(v: VbaValue): void }

export type Expr =
  | { e: "lit"; v: VbaValue }
  | { e: "id"; name: string; br?: boolean }
  | { e: "member"; obj: Expr | null; name: string }
  | { e: "bang"; obj: Expr | null; name: string }
  | { e: "call"; fn: Expr; args: (Expr | null)[]; names?: (string | null)[] }
  | { e: "bin"; op: string; a: Expr; b: Expr }
  | { e: "un"; op: string; a: Expr }
  | { e: "paren"; a: Expr }
  | { e: "new"; cls: string }
  | { e: "typeof"; a: Expr; cls: string }
  | { e: "slot"; ref: Ref };

export interface VarDecl { name: string; dims?: Expr[]; isNew?: boolean; type?: string; withEvents?: boolean }
export interface PrintItem { x: Expr | null; sep: "," | ";" | "" }
export type CaseTest = { k: "eq"; v: Expr } | { k: "range"; a: Expr; b: Expr } | { k: "is"; op: string; v: Expr };
export type Stmt = { line: number } & (
  | { s: "dim"; vars: VarDecl[]; isPublic?: boolean; isStatic?: boolean }
  | { s: "const"; name: string; value: Expr; isPublic?: boolean }
  | { s: "redim"; preserve: boolean; vars: { target: Expr; dims: Expr[] }[] }
  | { s: "assign"; target: Expr; value: Expr; set: boolean }
  | { s: "expr"; value: Expr }
  | { s: "if"; branches: { cond: Expr; body: Stmt[] }[]; other: Stmt[] }
  | { s: "for"; v: Expr; from: Expr; to: Expr; step?: Expr; body: Stmt[] }
  | { s: "foreach"; v: Expr; coll: Expr; body: Stmt[] }
  | { s: "do"; cond?: Expr; until: boolean; post: boolean; body: Stmt[] }
  | { s: "select"; value: Expr; cases: { tests: CaseTest[] | null; body: Stmt[] }[] }
  | { s: "with"; obj: Expr; body: Stmt[] }
  | { s: "exit"; what: string }
  | { s: "end" }
  | { s: "stop" }
  | { s: "onerror"; mode: "next" | "zero" | "reset" | "label"; label?: string }
  | { s: "resume"; mode: "retry" | "next" | "label"; label?: string }
  | { s: "goto"; label: string }
  | { s: "label"; name: string }
  | { s: "print"; file: Expr | null; items: PrintItem[] }
  | { s: "write"; file: Expr; items: (Expr | null)[] }
  | { s: "open"; path: Expr; mode: "input" | "output" | "append" | "binary" | "random"; file: Expr }
  | { s: "close"; files: Expr[] }
  | { s: "lineinput"; file: Expr; target: Expr }
  | { s: "input"; file: Expr; targets: Expr[] }
  | { s: "name"; from: Expr; to: Expr }
  | { s: "assert"; cond: Expr }
  | { s: "raise"; name: string; args: (Expr | null)[]; names?: (string | null)[] }
  | { s: "erase"; targets: Expr[] });

export interface Param { name: string; byRef: boolean; optional: boolean; paramArray: boolean; type?: string; def?: Expr }
export interface Proc {
  name: string;
  kind: "sub" | "function" | "get" | "let" | "set";
  params: Param[];
  body: Stmt[];
  isPrivate: boolean;
  isStatic: boolean;
  /** The declared return type (lower case) of a Function or Property Get. */
  type?: string;
  /** The line of the Sub/Function/Property line. */
  line: number;
  /** Labels at the top level of the body (index of the statement after them). */
  labels: Map<string, number>;
  /** Member attributes (Attribute Item.VB_UserMemId = 0 makes the default member). */
  attrs: Record<string, string>;
}
export interface PropProcs { get?: Proc; let?: Proc; set?: Proc }
export interface ParsedModule {
  /** Subs and Functions by lower-case name. */
  procs: Map<string, Proc>;
  /** Property Get/Let/Set by lower-case name. */
  props: Map<string, PropProcs>;
  /** Module-level declarations in order (Dim, Public, Private, Const; Enum members as constants). */
  decls: Stmt[];
  events: Map<string, { name: string; params: Param[] }>;
  implements: string[];
  /** Module attributes by lower-case name without the VB_ prefix (name, predeclaredid, base …), as written. */
  attributes: Map<string, string>;
  options: { explicit: boolean; base: number; compareText: boolean };
  /** Enum names (lower case) with their member names. */
  enums: Map<string, string[]>;
  /** User-defined types (Type … End Type) by lower-case name. */
  types: Map<string, { name: string; fields: VarDecl[] }>;
}

const RESERVED_LABEL = new Set(["else", "end", "loop", "next", "wend", "case", "resume", "then"]);

class Parser {
  i = 0;
  constructor(private toks: Tok[]) {}
  get cur(): Tok { return this.toks[this.i] ?? { t: "nl", v: "\n", line: this.toks[this.toks.length - 1]?.line ?? 0 }; }
  at(k: number): Tok | undefined { return this.toks[this.i + k]; }
  is(v: string): boolean { const c = this.cur; return (c.t === "id" || c.t === "op") && !c.br && c.v.toLowerCase() === v; }
  isAt(k: number, v: string): boolean { const c = this.toks[this.i + k]; return !!c && (c.t === "id" || c.t === "op") && !c.br && c.v.toLowerCase() === v; }
  eat(v: string): boolean { if (this.is(v)) { this.i++; return true; } return false; }
  need(v: string): void { if (!this.eat(v)) throw new VbaError(`Erwartet: ${v} (Zeile ${this.cur.line})`, 1002, this.cur.line); }
  name(): string { const c = this.cur; if (c.t !== "id") throw new VbaError(`Erwartet: Bezeichner (Zeile ${c.line})`, 1003, c.line); this.i++; return c.v; }
  eol(): boolean { return this.cur.t === "nl"; }
  skipNl(): void { while (this.i < this.toks.length && this.cur.t === "nl") this.i++; }
  skipLine(): void { while (this.i < this.toks.length && !(this.cur.t === "nl" && this.cur.v === "\n")) this.i++; }

  module(): ParsedModule {
    const m: ParsedModule = { procs: new Map(), props: new Map(), decls: [], events: new Map(), implements: [], attributes: new Map(), options: { explicit: false, base: 0, compareText: false }, enums: new Map(), types: new Map() };
    this.skipNl();
    while (this.i < this.toks.length) {
      const line = this.cur.line;
      if (this.eat("option")) {
        if (this.eat("explicit")) m.options.explicit = true;
        else if (this.eat("base")) m.options.base = Number(this.cur.v) || 0;
        else if (this.eat("compare")) m.options.compareText = this.is("text");
        this.skipLine(); this.skipNl(); continue;
      }
      if (this.eat("attribute")) {
        const n = this.name();
        if (this.eat(".")) { this.skipLine(); this.skipNl(); continue; }
        this.eat("=");
        const v = this.cur.v; this.skipLine();
        m.attributes.set(n.toLowerCase().replace(/^vb_/, ""), v);
        this.skipNl(); continue;
      }
      if (/^def(int|lng|lnglng|lngptr|cur|sng|dbl|dec|date|str|obj|var|bool|byte)$/i.test(this.cur.v) && this.cur.t === "id") { this.skipLine(); this.skipNl(); continue; }
      let isPrivate = false, isPublic = false, isStatic = false;
      while (this.is("public") || this.is("private") || this.is("static") || this.is("global") || this.is("friend")) {
        if (this.is("private")) isPrivate = true; else if (this.is("static")) isStatic = true; else isPublic = true;
        this.i++;
      }
      if (this.is("sub") || this.is("function") || this.is("property")) {
        const p = this.proc(isPrivate, isStatic);
        if (p.kind === "sub" || p.kind === "function") m.procs.set(p.name.toLowerCase(), p);
        else { const k = p.name.toLowerCase(); const e = m.props.get(k) ?? {}; e[p.kind] = p; m.props.set(k, e); }
        this.skipNl(); continue;
      }
      if (this.is("declare")) throw new VbaError("Declare ist nicht erlaubt (Makros bleiben in der App).", 1004, this.cur.line);
      if (this.eat("event")) {
        const name = this.name();
        const params = this.is("(") ? this.params() : [];
        m.events.set(name.toLowerCase(), { name, params });
        this.skipNl(); continue;
      }
      if (this.eat("implements")) { let n = this.name(); while (this.eat(".")) n = this.name(); m.implements.push(n); this.skipNl(); continue; }
      if (this.eat("enum")) {
        const ename = this.name();
        const members: string[] = [];
        let prev: string | null = null;
        this.skipNl();
        while (!(this.is("end") && this.isAt(1, "enum"))) {
          if (this.i >= this.toks.length) throw new VbaError("Erwartet: End Enum", 1005, line);
          const l = this.cur.line;
          const n = this.name();
          const value: Expr = this.eat("=") ? this.expr() : prev ? { e: "bin", op: "+", a: { e: "id", name: prev }, b: { e: "lit", v: 1 } } : { e: "lit", v: 0 };
          m.decls.push({ line: l, s: "const", name: n, value, isPublic: !isPrivate });
          members.push(n);
          prev = n;
          this.skipNl();
        }
        this.i += 2;
        m.enums.set(ename.toLowerCase(), members);
        this.skipNl(); continue;
      }
      if (this.eat("type")) {
        const tname = this.name();
        const fields: VarDecl[] = [];
        this.skipNl();
        while (!(this.is("end") && this.isAt(1, "type"))) {
          if (this.i >= this.toks.length) throw new VbaError("Erwartet: End Type", 1005, line);
          fields.push(this.varDecl());
          this.skipNl();
        }
        this.i += 2;
        m.types.set(tname.toLowerCase(), { name: tname, fields });
        this.skipNl(); continue;
      }
      if (this.is("const")) { const st = this.statement(); if (st.s === "const") st.isPublic = isPublic; m.decls.push(st); this.skipNl(); continue; }
      if (this.eat("dim") || isPublic || isPrivate || isStatic) {
        const vars: VarDecl[] = [];
        do vars.push(this.varDecl()); while (this.eat(","));
        m.decls.push({ line, s: "dim", vars, isPublic });
        this.skipNl(); continue;
      }
      throw new VbaError(`Ungültig außerhalb einer Prozedur (Zeile ${line})`, 1011, line);
    }
    return m;
  }

  varDecl(): VarDecl {
    const withEvents = this.eat("withevents");
    const name = this.name();
    let dims: Expr[] | undefined;
    if (this.eat("(")) dims = this.dimList();
    let isNew = false, type: string | undefined;
    if (this.eat("as")) { isNew = this.is("new"); type = this.typeName(); }
    return { name, dims, isNew, type, withEvents };
  }

  dimList(): Expr[] {
    const dims: Expr[] = [];
    while (!this.eat(")")) { const a = this.expr(); if (this.eat("to")) { const b = this.expr(); dims.push({ e: "bin", op: "range", a, b }); } else dims.push(a); this.eat(","); }
    return dims;
  }

  params(): Param[] {
    const params: Param[] = [];
    this.need("(");
    while (!this.eat(")")) {
      let byRef = true, optional = false, paramArray = false;
      for (;;) { if (this.eat("byval")) byRef = false; else if (this.eat("byref")) byRef = true; else if (this.eat("optional")) optional = true; else if (this.eat("paramarray")) paramArray = true; else break; }
      const pn = this.name();
      if (this.eat("(")) this.need(")");
      let type: string | undefined;
      if (this.eat("as")) type = this.typeName();
      let def: Expr | undefined;
      if (this.eat("=")) def = this.expr();
      params.push({ name: pn, byRef, optional: optional || paramArray, paramArray, type, def });
      this.eat(",");
    }
    return params;
  }

  proc(isPrivate: boolean, isStatic: boolean): Proc {
    const line = this.cur.line;
    let kind: Proc["kind"];
    if (this.eat("property")) { const k = this.cur.v.toLowerCase(); if (k !== "get" && k !== "let" && k !== "set") throw new VbaError(`Erwartet: Get, Let oder Set (Zeile ${line})`, 1002, line); this.i++; kind = k; }
    else { kind = this.is("sub") ? "sub" : "function"; this.i++; }
    const name = this.name();
    const params = this.is("(") ? this.params() : [];
    let type: string | undefined;
    if (this.eat("as")) type = this.typeName();
    const endWord = kind === "sub" ? "sub" : kind === "function" ? "function" : "property";
    const attrs: Record<string, string> = {};
    const body = this.block(() => this.is("end") && this.isAt(1, endWord), attrs);
    this.need("end"); this.i++;
    const labels = new Map<string, number>();
    body.forEach((s, k) => { if (s.s === "label") labels.set(s.name.toLowerCase(), k); });
    return { name, kind, params, body, isPrivate, isStatic, type, line, labels, attrs };
  }

  /** A type name (lower case, without a library prefix: VBA.Collection is collection, MSForms.TextBox textbox). */
  typeName(): string {
    this.eat("new");
    let n = this.name();
    while (this.eat(".")) n = this.name();
    if (this.eat("*")) this.expr();
    return n.toLowerCase();
  }

  block(stop: () => boolean, attrs?: Record<string, string>): Stmt[] {
    const out: Stmt[] = [];
    for (;;) {
      this.skipNl();
      if (this.i >= this.toks.length) throw new VbaError("Unerwartetes Ende des Moduls", 1005, this.cur.line);
      if (stop()) return out;
      if (this.is("attribute")) {
        // Attribute Item.VB_UserMemId = 0 (and others) inside a procedure.
        this.i++;
        const parts: string[] = [];
        while (!this.eol() && !this.is("=")) parts.push(this.cur.v), this.i++;
        this.eat("=");
        let v = "";
        while (!this.eol()) { v += this.cur.v; this.i++; }
        if (attrs) attrs[(parts[parts.length - 1] ?? "").toLowerCase().replace(/^vb_/, "")] = v;
        continue;
      }
      out.push(this.statement());
    }
  }

  statement(): Stmt {
    const line = this.cur.line;
    const c = this.cur;
    // A line label (Fehler:) at the start of a line.
    if (c.t === "id" && c.first && !c.br && this.at(1)?.t === "nl" && this.at(1)?.v === ":" && !RESERVED_LABEL.has(c.v.toLowerCase())) { this.i += 2; return { line, s: "label", name: c.v }; }
    if (c.t === "num" && c.first && /^\d+$/.test(c.v)) { this.i++; if (this.cur.v === ":") this.i++; return { line, s: "label", name: c.v }; }
    if (this.is("dim") || this.is("static") || ((this.is("public") || this.is("private")) && !this.isAt(1, "="))) {
      const isStatic = this.is("static");
      this.i++;
      if (this.is("const")) return this.statement();
      const vars: VarDecl[] = [];
      do vars.push(this.varDecl()); while (this.eat(","));
      return { line, s: "dim", vars, isStatic };
    }
    if (this.eat("const")) { const name = this.name(); if (this.eat("as")) this.typeName(); this.need("="); const st: Stmt = { line, s: "const", name, value: this.expr() }; while (this.eat(",")) { /* Const a = 1, b = 2: the rest as further constants is rare; skip */ this.name(); if (this.eat("as")) this.typeName(); this.need("="); this.expr(); } return st; }
    if (this.eat("redim")) {
      const preserve = this.eat("preserve");
      const vars: { target: Expr; dims: Expr[] }[] = [];
      do {
        let target: Expr = { e: "id", name: this.name() };
        while (this.is(".")) { this.i++; target = { e: "member", obj: target, name: this.name() }; }
        this.need("(");
        const dims = this.dimList();
        if (this.eat("as")) this.typeName();
        vars.push({ target, dims });
      } while (this.eat(","));
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
        if (this.is("end") && this.isAt(1, "if")) this.i += 2;
        return { line, s: "if", branches: [{ cond, body }], other };
      }
      const stopIf = () => this.is("elseif") || this.is("else") || (this.is("end") && this.isAt(1, "if"));
      const branches = [{ cond, body: this.block(stopIf) }];
      let other: Stmt[] = [];
      for (;;) {
        if (this.eat("elseif")) { const cnd = this.expr(); this.need("then"); branches.push({ cond: cnd, body: this.block(stopIf) }); continue; }
        if (this.eat("else")) { other = this.block(() => this.is("end") && this.isAt(1, "if")); }
        this.need("end"); this.need("if");
        return { line, s: "if", branches, other };
      }
    }
    if (this.eat("for")) {
      if (this.eat("each")) { const v = this.lvalue(); this.need("in"); const coll = this.expr(); const body = this.block(() => this.is("next")); this.endNext(); return { line, s: "foreach", v, coll, body }; }
      const v = this.lvalue(); this.need("="); const from = this.expr(); this.need("to"); const to = this.expr();
      const step = this.eat("step") ? this.expr() : undefined;
      const body = this.block(() => this.is("next"));
      this.endNext();
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
        let tests: CaseTest[] | null = [];
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

  /** Next, Next i, Next i, j (one Next closing several loops is not supported). */
  endNext(): void { this.need("next"); if (this.cur.t === "id") { this.i++; while (this.is(".")) { this.i++; this.i++; } } }

  /** A variable or element as the target of For, Line Input …. */
  lvalue(): Expr { return this.postfix(); }

  fileNum(): Expr { this.eat("#"); return this.expr(); }

  simple(): Stmt {
    const line = this.cur.line;
    if (this.eat("exit")) { const what = this.name().toLowerCase(); return { line, s: "exit", what }; }
    if (this.is("end") && !this.at(1)?.br && !["sub", "function", "if", "property", "select", "with", "type", "enum"].some((w) => this.isAt(1, w))) { this.i++; return { line, s: "end" }; }
    if (this.is("stop") && (this.at(1)?.t === "nl")) { this.i++; return { line, s: "stop" }; }
    if (this.eat("goto")) { const l = this.cur.v; this.i++; return { line, s: "goto", label: l }; }
    if (this.is("on") && this.isAt(1, "error")) {
      this.i += 2;
      if (this.eat("resume")) { this.need("next"); return { line, s: "onerror", mode: "next" }; }
      this.need("goto");
      if (this.cur.t === "num" && this.cur.v === "0") { this.i++; return { line, s: "onerror", mode: "zero" }; }
      if (this.cur.v === "-" && this.at(1)?.v === "1") { this.i += 2; return { line, s: "onerror", mode: "reset" }; }
      const l = this.cur.v; this.i++;
      return { line, s: "onerror", mode: "label", label: l };
    }
    if (this.is("resume") && !this.isAt(1, "=")) {
      this.i++;
      if (this.eat("next")) return { line, s: "resume", mode: "next" };
      if (this.eol() || this.is("else")) return { line, s: "resume", mode: "retry" };
      const l = this.cur.v; this.i++;
      return l === "0" ? { line, s: "resume", mode: "retry" } : { line, s: "resume", mode: "label", label: l };
    }
    if (this.eat("raiseevent")) {
      const name = this.name();
      let args: (Expr | null)[] = [], names: (string | null)[] | undefined;
      if (this.is("(")) { const r = this.argList(); args = r.args; names = r.names; }
      return { line, s: "raise", name, args, names };
    }
    if (this.is("debug") && this.isAt(1, ".") && this.isAt(2, "print")) { this.i += 3; return { line, s: "print", file: null, items: this.printItems() }; }
    if (this.is("debug") && this.isAt(1, ".") && this.isAt(2, "assert")) { this.i += 3; return { line, s: "assert", cond: this.expr() }; }
    if (this.is("print") && this.isAt(1, "#")) { this.i += 2; const file = this.expr(); this.eat(","); return { line, s: "print", file, items: this.printItems() }; }
    if (this.is("write") && this.isAt(1, "#")) {
      this.i += 2; const file = this.expr(); this.eat(",");
      const items: (Expr | null)[] = [];
      while (!this.eol() && !this.is("else")) { if (this.is(",") || this.is(";")) { this.i++; continue; } items.push(this.expr()); }
      return { line, s: "write", file, items };
    }
    if (this.is("open") && !this.isAt(1, "=") && !this.isAt(1, ".") && !this.isAt(1, "(")) {
      this.i++;
      const path = this.expr();
      this.need("for");
      const mode = this.name().toLowerCase();
      if (!["input", "output", "append", "binary", "random"].includes(mode)) throw new VbaError(`Ungültiger Dateimodus ${mode} (Zeile ${line})`, 1002, line);
      while (!this.is("as")) { if (this.eol()) throw new VbaError(`Erwartet: As (Zeile ${line})`, 1002, line); this.i++; }
      this.need("as");
      const file = this.fileNum();
      if (this.eat("len")) { this.need("="); this.expr(); }
      return { line, s: "open", path, mode: mode as Extract<Stmt, { s: "open" }>["mode"], file };
    }
    if (this.is("close") && (this.at(1)?.t === "nl" || this.isAt(1, "#") || this.at(1)?.t === "num" || this.at(1)?.t === "id") && !this.isAt(1, "=")) {
      this.i++;
      const files: Expr[] = [];
      while (!this.eol() && !this.is("else")) { files.push(this.fileNum()); this.eat(","); }
      return { line, s: "close", files };
    }
    if (this.is("line") && this.isAt(1, "input")) { this.i += 2; this.need("#"); const file = this.expr(); this.need(","); return { line, s: "lineinput", file, target: this.lvalue() }; }
    if (this.is("input") && this.isAt(1, "#")) {
      this.i += 2; const file = this.expr(); this.need(",");
      const targets: Expr[] = [];
      do targets.push(this.lvalue()); while (this.eat(","));
      return { line, s: "input", file, targets };
    }
    if (this.eat("erase")) { const targets: Expr[] = []; do targets.push(this.lvalue()); while (this.eat(",")); return { line, s: "erase", targets }; }
    if (this.is("name") && !["=", ".", "(", "!"].some((v) => this.isAt(1, v)) && this.at(1)?.t !== "nl") { this.i++; const from = this.expr(); this.need("as"); return { line, s: "name", from, to: this.expr() }; }
    if (this.eat("call")) { const value = this.expr(); return { line, s: "expr", value }; }
    const set = this.eat("set");
    this.eat("let");
    let target = this.postfix();
    if (this.eat("=")) return { line, s: "assign", target, value: this.expr(), set };
    // A procedure call without parentheses: Name arg1, arg2 (named arguments with :=).
    if (!this.eol() && !this.is("else") && this.cur.v !== ":") {
      const args: (Expr | null)[] = [], names: (string | null)[] = [];
      let named = false;
      // Foo (1), 2: the parentheses belong to the first argument.
      if (target.e === "call" && target.args.length === 1 && target.args[0] && !target.names && this.is(",")) {
        args.push({ e: "paren", a: target.args[0] }); names.push(null);
        target = target.fn;
        this.i++;
      }
      for (;;) {
        if (this.is(",")) { args.push(null); names.push(null); this.i++; continue; }
        if (this.cur.t === "id" && this.isAt(1, ":=")) { names.push(this.cur.v); this.i += 2; named = true; } else names.push(null);
        args.push(this.expr());
        if (!this.eat(",")) break;
        if (this.eol()) { args.push(null); names.push(null); break; }
      }
      return { line, s: "expr", value: { e: "call", fn: target, args, names: named ? names : undefined } };
    }
    return { line, s: "expr", value: target };
  }

  printItems(): PrintItem[] {
    const items: PrintItem[] = [];
    while (!this.eol() && !this.is("else")) {
      if (this.cur.v === ";" || this.cur.v === ",") { items.push({ x: null, sep: this.cur.v as "," | ";" }); this.i++; continue; }
      const x = this.expr();
      let sep: PrintItem["sep"] = "";
      if (this.cur.v === ";" || this.cur.v === ",") { sep = this.cur.v as "," | ";"; this.i++; }
      items.push({ x, sep });
    }
    return items;
  }

  argList(): { args: (Expr | null)[]; names?: (string | null)[] } {
    this.need("(");
    const args: (Expr | null)[] = [], names: (string | null)[] = [];
    let named = false;
    while (!this.eat(")")) {
      if (this.is(",")) { args.push(null); names.push(null); this.i++; if (this.is(")")) { args.push(null); names.push(null); } continue; }
      if (this.cur.t === "id" && this.isAt(1, ":=")) { names.push(this.cur.v); this.i += 2; named = true; } else names.push(null);
      // ByVal in a call (ByVal x) is accepted and ignored.
      this.eat("byval");
      args.push(this.expr());
      if (!this.eat(",")) { this.need(")"); break; }
      if (this.is(")")) { args.push(null); names.push(null); }
    }
    return { args, names: named ? names : undefined };
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
    while ((this.cur.t === "op" && ["=", "<>", "<", ">", "<=", ">="].includes(this.cur.v)) || this.is("is") || this.is("like")) { const op = this.cur.v.toLowerCase(); this.i++; a = { e: "bin", op, a, b: this.concat() }; }
    return a;
  }
  concat(): Expr { let a = this.additive(); while (this.cur.t === "op" && this.cur.v === "&") { this.i++; a = { e: "bin", op: "&", a, b: this.additive() }; } return a; }
  additive(): Expr { let a = this.modulo(); while (this.cur.t === "op" && (this.cur.v === "+" || this.cur.v === "-")) { const op = this.cur.v; this.i++; a = { e: "bin", op, a, b: this.modulo() }; } return a; }
  modulo(): Expr { let a = this.intdiv(); while (this.is("mod")) { this.i++; a = { e: "bin", op: "mod", a, b: this.intdiv() }; } return a; }
  intdiv(): Expr { let a = this.mul(); while (this.cur.t === "op" && this.cur.v === "\\") { this.i++; a = { e: "bin", op: "\\", a, b: this.mul() }; } return a; }
  mul(): Expr { let a = this.unary(); while (this.cur.t === "op" && (this.cur.v === "*" || this.cur.v === "/")) { const op = this.cur.v; this.i++; a = { e: "bin", op, a, b: this.unary() }; } return a; }
  unary(): Expr { if (this.cur.t === "op" && (this.cur.v === "-" || this.cur.v === "+")) { const op = this.cur.v; this.i++; const a = this.unary(); return op === "-" ? { e: "un", op: "neg", a } : a; } return this.power(); }
  power(): Expr {
    let a = this.postfix();
    while (this.cur.t === "op" && this.cur.v === "^") {
      this.i++;
      // A unary minus after ^ binds to the exponent (2 ^ -1).
      const nx = this.cur as Tok;
      const b = nx.t === "op" && nx.v === "-" ? (this.i++, { e: "un", op: "neg", a: this.postfix() } as Expr) : this.postfix();
      a = { e: "bin", op: "^", a, b };
    }
    return a;
  }
  postfix(): Expr {
    let x = this.primary();
    for (;;) {
      if (this.cur.v === "(" && this.cur.t === "op") { const { args, names } = this.argList(); x = { e: "call", fn: x, args, names }; continue; }
      if (this.cur.t === "op" && this.cur.v === ".") { this.i++; x = { e: "member", obj: x, name: this.name() }; continue; }
      if (this.cur.t === "op" && this.cur.v === "!") { this.i++; x = { e: "bang", obj: x, name: this.name() }; continue; }
      return x;
    }
  }
  primary(): Expr {
    const c = this.cur;
    if (c.t === "num") { this.i++; return { e: "lit", v: parseNum(c.v) }; }
    if (c.t === "str") { this.i++; return { e: "lit", v: c.v }; }
    if (c.t === "date") { this.i++; return { e: "lit", v: parseDateLiteral(c.v) }; }
    if (c.t === "op" && c.v === "(") { this.i++; const x = this.expr(); this.need(")"); return { e: "paren", a: x }; }
    if (c.t === "op" && c.v === ".") { this.i++; return { e: "member", obj: null, name: this.name() }; }
    if (c.t === "op" && c.v === "!") { this.i++; return { e: "bang", obj: null, name: this.name() }; }
    if (c.t === "id" && !c.br) {
      const k = c.v.toLowerCase();
      if (k === "new") { this.i++; let n = this.name(); while (this.eat(".")) n += "." + this.name(); return { e: "new", cls: n }; }
      if (k === "typeof") { this.i++; const a = this.postfix(); this.need("is"); let n = this.name(); while (this.eat(".")) n = this.name(); return { e: "typeof", a, cls: n }; }
      if (k === "addressof") throw new VbaError(`AddressOf ist nicht erlaubt (Zeile ${c.line})`, 1004, c.line);
      if (k === "true") { this.i++; return { e: "lit", v: true }; }
      if (k === "false") { this.i++; return { e: "lit", v: false }; }
      if (k === "nothing") { this.i++; return { e: "lit", v: null }; }
      if (k === "empty") { this.i++; return { e: "lit", v: undefined }; }
      if (k === "null") { this.i++; return { e: "lit", v: null }; }
    }
    if (c.t === "id") { this.i++; return { e: "id", name: c.v, br: c.br }; }
    throw new VbaError(`Unerwartet: "${c.v}" (Zeile ${c.line})`, 1006, c.line);
  }
}

function parseNum(s: string): number {
  const t = s.replace(/[!#@%&]$/, "");
  if (/^&H/i.test(t)) { const n = parseInt(t.slice(2), 16); return t.length <= 6 && n > 0x7fff && !/&$/.test(s) ? n - 0x10000 : n > 0x7fffffff ? n - 0x100000000 : n; }
  if (/^&O/i.test(t)) return parseInt(t.slice(2), 8);
  return Number(t);
}

/** #12/31/2024# (month first, as VBA's date literals are) or #2024-12-31#, with an optional time. */
function parseDateLiteral(s: string): Date {
  const m = /^(\d{1,4})[-/.](\d{1,2})[-/.](\d{1,4})(?:\s+(\d{1,2}):(\d{2})(?::(\d{2}))?\s*([AP]M)?)?$/i.exec(s.trim());
  const t = /^(\d{1,2}):(\d{2})(?::(\d{2}))?\s*([AP]M)?$/i.exec(s.trim());
  const hour = (h: string | undefined, ap: string | undefined) => { let x = Number(h ?? 0); if (ap && /pm/i.test(ap) && x < 12) x += 12; if (ap && /am/i.test(ap) && x === 12) x = 0; return x; };
  if (m) {
    const [y, mo, d] = m[1].length === 4 ? [Number(m[1]), Number(m[2]), Number(m[3])] : [Number(m[3]) < 100 ? 2000 + Number(m[3]) : Number(m[3]), Number(m[1]), Number(m[2])];
    return new Date(y, mo - 1, d, hour(m[4], m[7]), Number(m[5] ?? 0), Number(m[6] ?? 0));
  }
  if (t) return new Date(1899, 11, 30, hour(t[1], t[4]), Number(t[2]), Number(t[3] ?? 0));
  return new Date(s);
}

/** A module's procedures, declarations, events and attributes. */
export function parseModule(src: string): ParsedModule {
  return new Parser(lex(src)).module();
}

/** One expression (the debugger's Direktbereich and watches). */
export function parseExpression(src: string): Expr {
  const p = new Parser(lex(src));
  const x = p.expr();
  if (!p.eol()) throw new VbaError(`Unerwartet: "${p.cur.v}"`, 1006);
  return x;
}

/** Statements outside a procedure (the debugger's Direktbereich). */
export function parseStatements(src: string): Stmt[] {
  const p = new Parser(lex(src + "\n"));
  const out: Stmt[] = [];
  p.skipNl();
  while (p.i < (p as unknown as { toks: Tok[] }).toks.length) { out.push(p.statement()); p.skipNl(); }
  return out;
}
