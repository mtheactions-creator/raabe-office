// Formelauswertung (Formel auswerten): which part of a formula Excel works out next. References
// to single cells first, then each operation and function once its operands are plain values,
// left to right, innermost first; WENN works out its condition first and then itself as a
// whole (only the branch it takes counts). Ranges stay as they are until their function is
// worked out.

import { parseWithSpans, type Node } from "./parse";

const LITERAL = new Set(["num", "str", "bool", "err", "arr", "miss"]);
const IF_LIKE = new Set(["IF", "IFERROR", "IFNA", "IFS", "SWITCH", "CHOOSE"]);

/** The offsets of the part to work out next in a local formula ("=" included), or null when it is a value. */
export function nextStep(local: string): { start: number; end: number } | null {
  let parsed: ReturnType<typeof parseWithSpans>;
  try {
    parsed = parseWithSpans(local);
  } catch {
    return null;
  }
  const { node, span } = parsed;
  const isRange = (n: Node) => (n.t === "ref" && (n.r0 !== n.r1 || n.c0 !== n.c1)) || (n.t === "bin" && n.op === ":");
  // A negative number is a value too (the result of a step can be one).
  const settled = (n: Node, asArg: boolean): boolean => LITERAL.has(n.t) || (n.t === "un" && n.op === "-" && n.a.t === "num") || (asArg && (isRange(n) || n.t === "name"));
  const find = (n: Node, asArg: boolean): Node | null => {
    if (settled(n, asArg)) return null;
    switch (n.t) {
      case "ref":
      case "name":
      case "struct":
        return n;
      case "un":
        return find(n.a, false) ?? n;
      case "paren":
        return find(n.a, false) ?? n;
      case "bin":
        if (n.op === ":") return n;
        return find(n.a, false) ?? find(n.b, false) ?? n;
      case "fn": {
        const args = IF_LIKE.has(n.name.toUpperCase()) ? n.args.slice(0, 1) : n.args;
        for (const a of args) {
          const f = find(a, true);
          if (f) return f;
        }
        return n;
      }
      default:
        return null;
    }
  };
  const hit = find(node, false);
  if (!hit) return null;
  const [start, end] = span(hit);
  return { start, end };
}
