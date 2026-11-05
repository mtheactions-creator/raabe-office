// What Excel does with a formula when it is entered (recorded with alerts on, D-066): some slips it
// mends without asking (brackets of a function left open, "++"), for others it proposes a mended
// formula and asks (a string or a plain bracket left open, an operator at the end, a bracket too
// many, two operators in a row, "2(3)", "1 2"), and some it refuses (a formula that does not parse,
// a function with too few or too many arguments).

import { FUNCS } from "./eval";
import "./functions";
import { lex, type Token } from "./lexer";
import { parseFormula, type Node } from "./parse";

export type EntryCheck =
  | { kind: "ok"; text: string }
  /** Excel asks whether to take this mended formula instead. */
  | { kind: "propose"; text: string }
  /** Excel refuses the formula and the cell stays open. */
  | { kind: "error" };

/** Smallest number of arguments Excel accepts when a formula is typed, where it differs from evaluation. */
const ENTRY_MIN: Record<string, number> = { IF: 2 };

const BINARY_ONLY = new Set(["*", "/", "^", "&", "=", "<", ">", "<=", ">=", "<>"]);

export function checkEntry(text: string): EntryCheck {
  if (!text.startsWith("=") || text.length < 2) return { kind: "ok", text };
  let body = text.slice(1).replace(/\s+$/, "");
  let propose = false;
  if ((body.match(/"/g)?.length ?? 0) % 2 === 1) {
    body += '"';
    propose = true;
  }
  let toks: { k: Token["k"]; text: string }[];
  try {
    toks = lex(body, "local").map((t) => ({ k: t.k, text: body.slice(t.s, t.e) }));
  } catch {
    return { kind: "error" };
  }
  const out: { k: Token["k"]; text: string }[] = [];
  for (let i = 0; i < toks.length; i++) {
    const t = toks[i];
    const prev = out.at(-1);
    // "1 2" is "12" (spaces between two numbers go).
    if (t.k === "ws" && prev?.k === "num" && toks[i + 1]?.k === "num") { propose = true; continue; }
    if (t.k === "num" && prev?.k === "num") { prev.text += t.text; continue; }
    // "2(3)" is "2*(3)".
    if (t.k === "(" && prev?.k === "num") { out.push({ k: "op", text: "*" }); propose = true; }
    if (t.k === "op" && prev?.k === "op") {
      // "1++2" is "1+2" without asking; a second operator that cannot stand alone goes, with asking.
      if (t.text === "+" && prev.text === "+") continue;
      if (BINARY_ONLY.has(t.text)) { propose = true; continue; }
    }
    out.push({ ...t });
  }
  // Operators left at the end go ("=1+" is "=1"); a range colon does not.
  while (out.length) {
    const last = out.at(-1)!;
    if (last.k === "ws") { out.pop(); continue; }
    if (last.k === "op" && last.text !== ":" && last.text !== "%") { out.pop(); propose = true; continue; }
    break;
  }
  if (!out.length) return { kind: "error" };
  // Brackets: one too many goes (asking); open ones are closed, without asking when all belong to functions.
  const stack: boolean[] = [];
  const kept: typeof out = [];
  for (const t of out) {
    if (t.k === "(") stack.push(kept.at(-1)?.k === "func");
    if (t.k === ")") {
      if (!stack.length) { propose = true; continue; }
      stack.pop();
    }
    kept.push(t);
  }
  if (stack.some((fn) => !fn)) propose = true;
  let fixed = kept.map((t) => t.text).join("") + ")".repeat(stack.length);
  let node: Node;
  try {
    node = parseFormula(fixed, "local");
  } catch {
    return { kind: "error" };
  }
  if (!argumentsFit(node)) return { kind: "error" };
  fixed = "=" + fixed;
  return propose ? { kind: "propose", text: fixed } : { kind: "ok", text: fixed };
}

/** Whether every function has as many arguments as Excel accepts. */
function argumentsFit(n: Node): boolean {
  switch (n.t) {
    case "fn": {
      const d = FUNCS.get(n.name.toUpperCase());
      if (d) {
        const min = Math.max(d.min, ENTRY_MIN[n.name.toUpperCase()] ?? 0);
        if (n.args.length < min || n.args.length > d.max) return false;
      }
      return n.args.every(argumentsFit);
    }
    case "bin": return argumentsFit(n.a) && argumentsFit(n.b);
    case "un": case "paren": return argumentsFit(n.a);
    default: return true;
  }
}
