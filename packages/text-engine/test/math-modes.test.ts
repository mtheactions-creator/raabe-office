// Formel > Konvertierungen: Word's linear format both ways, LaTeX typing and Text (plain text).

import { describe, expect, it } from "vitest";
import { MathEditor, buildLinear, parseLatex, seqToLatex, seqToLinear, type Seq } from "../src";

const typed = (keys: string, setup?: (me: MathEditor) => void): Seq => {
  const me = new MathEditor([]);
  setup?.(me);
  for (const ch of keys) me.type(ch);
  return me.value();
};

describe("linear format", () => {
  const cases: [string, string][] = [
    ["\\frac{a}{b}", "a/b"],
    ["\\frac{a+b}{c}", "(a+b)/c"],
    ["x^{2}", "x^2"],
    ["x_{i+1}", "x_(i+1)"],
    ["\\sqrt{x}", "√x"],
    ["\\sqrt[3]{x+1}", "√(3&x+1)"],
    ["\\sum_{i=1}^{n}{x_{i}}", "∑_(i=1)^n▒(x_i)"],
    ["\\sin{x}", "sin⁡x"],
    ["\\begin{matrix}a&b\\\\c&d\\end{matrix}", "■(a&b@c&d)"],
  ];
  it.each(cases)("%s is written as %s and built up again", (latex, linear) => {
    const seq = parseLatex(latex);
    expect(seqToLinear(seq)).toBe(linear);
    expect(seqToLatex(buildLinear(linear))).toBe(seqToLatex(seq));
  });
});

describe("typing modes", () => {
  it("LaTeX: a fraction built up on the space", () => {
    expect(seqToLatex(typed("\\frac{a}{b} ", (me) => { me.latex = true; }))).toBe("\\frac{a}{b}");
  });
  it("LaTeX: a/b stays as typed", () => {
    expect(seqToLatex(typed("a/b ", (me) => { me.latex = true; }))).toBe("a/b\\ ");
  });
  it("Text: nothing built, plain text", () => {
    const seq = typed("a/b ", (me) => { me.normal = true; });
    expect(seq).toEqual([{ k: "t", s: "a/b ", nor: true }]);
  });
});
