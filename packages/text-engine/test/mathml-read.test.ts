// MathML read into the equation model (D-079): Freihandgleichung and pasted MathML.
import { describe, expect, it } from "vitest";
import { mathmlToSeq } from "../src/math/mathml-read";
import { seqToLinear } from "../src";

const M = (body: string) => `<math xmlns="http://www.w3.org/1998/Math/MathML">${body}</math>`;

describe("MathML reader", () => {
  it("reads a fraction, a power and a root", () => {
    const seq = mathmlToSeq(M("<mrow><mfrac><mrow><mi>a</mi><mo>+</mo><mn>1</mn></mrow><mi>b</mi></mfrac><mo>=</mo><msup><mi>x</mi><mn>2</mn></msup><mo>+</mo><msqrt><mi>y</mi></msqrt></mrow>"))!;
    expect(seq.map((n) => n.k)).toEqual(["frac", "t", "scr", "t", "rad"]);
    expect(seqToLinear(seq)).toContain("^2");
  });

  it("reads a sum with limits and takes the next element as its operand", () => {
    const seq = mathmlToSeq(M("<mrow><munderover><mo>∑</mo><mrow><mi>i</mi><mo>=</mo><mn>1</mn></mrow><mi>n</mi></munderover><msub><mi>a</mi><mi>i</mi></msub></mrow>"))!;
    expect(seq.length).toBe(1);
    const n = seq[0] as Extract<(typeof seq)[number], { k: "nary" }>;
    expect(n.k).toBe("nary");
    expect(n.chr).toBe("∑");
    expect(n.e[0].k).toBe("scr");
  });

  it("reads a prefixed math element with text around it (the clipboard's HTML)", () => {
    const seq = mathmlToSeq(`<html><body><mml:math xmlns:mml="http://www.w3.org/1998/Math/MathML"><mml:mi>x</mml:mi><mml:mo>≤</mml:mo><mml:mn>3</mml:mn></mml:math></body></html>`);
    expect(seq?.length).toBe(3);
  });

  it("reads accents, fences and matrices", () => {
    const seq = mathmlToSeq(M("<mover accent='true'><mi>v</mi><mo>→</mo></mover><mfenced open='[' close=']'><mi>a</mi><mi>b</mi></mfenced><mtable><mtr><mtd><mn>1</mn></mtd><mtd><mn>0</mn></mtd></mtr><mtr><mtd><mn>0</mn></mtd><mtd><mn>1</mn></mtd></mtr></mtable>"))!;
    expect(seq.map((n) => n.k)).toEqual(["acc", "d", "m"]);
  });

  it("gives null for no MathML", () => {
    expect(mathmlToSeq("x+1")).toBeNull();
  });
});
