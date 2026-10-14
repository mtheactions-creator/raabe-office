// Word's linear format (UnicodeMath, Unicode Technical Note #28), for Formel > Konvertieren:
// a/b, x^2, x_(i+1), √(n&x), ∑_(i=1)^n▒x, sin⁡x, lim┬(n→∞), ■(a&b@c&d). Operands of more than
// one letter run are put in parentheses; building the text up again (buildLinear) removes them.

import { mergeText, type MNode, type Seq } from "./ast";

const SIMPLE = /^[\p{L}\p{N}.,′∞]+$/u;

function op(s: Seq): string {
  const m = mergeText(s);
  const t = seqToLinear(m);
  if (m.length === 1 && m[0].k === "t" && !m[0].nor && SIMPLE.test(m[0].s)) return t;
  return `(${t})`;
}

function node(n: MNode): string {
  switch (n.k) {
    case "t": return n.nor ? `"${n.s}"` : n.s;
    case "frac": return `${op(n.num)}${n.type === "noBar" ? "¦" : "/"}${op(n.den)}`;
    case "scr": return `${op(n.base)}${n.sub ? `_${op(n.sub)}` : ""}${n.sup ? `^${op(n.sup)}` : ""}`;
    case "rad": return n.deg ? `√(${seqToLinear(n.deg)}&${seqToLinear(n.e)})` : `√${op(n.e)}`;
    case "nary": return `${n.chr}${n.sub ? `_${op(n.sub)}` : ""}${n.sup ? `^${op(n.sup)}` : ""}▒${op(n.e)}`;
    case "func": return `${seqToLinear(n.name)}⁡${op(n.e)}`;
    case "d": return `${n.beg || "├"}${n.items.map(seqToLinear).join("│")}${n.end || "┤"}`;
    case "acc": return `${op(n.e)}${n.chr}`;
    case "bar": return `${n.top ? "¯" : "▁"}${op(n.e)}`;
    case "lim": return `${seqToLinear(n.e)}${n.upper ? "┴" : "┬"}${op(n.lim)}`;
    case "group": return `${n.chr}${op(n.e)}`;
    case "m": return `■(${n.rows.map((r) => r.map(seqToLinear).join("&")).join("@")})`;
    case "eqArr": return `█(${n.rows.map(seqToLinear).join("@")})`;
    case "box": return `▭(${seqToLinear(n.e)})`;
  }
}

/** An equation in Word's linear format. */
export function seqToLinear(s: Seq): string {
  return mergeText(s).map(node).join("");
}
