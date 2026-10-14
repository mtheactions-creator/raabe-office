import { readFileSync } from "node:fs";
import { describe, expect, it } from "vitest";
import { DEFAULT_MATH, Doc, OBJ, makePara, parseLatex } from "@raabe/text-engine";
import { latexToOmml, ommlToLatex, ommlToSeq, readMathPr, seqToOmml } from "../src/math";
import { KEEP_ALL, M_NS, norm } from "./omml-norm";
import { readDocx, writeDocx } from "../src";
import { strFromU8, unzipSync } from "fflate";

const M = "http://schemas.openxmlformats.org/officeDocument/2006/math";
const parseOmml = (xml: string) =>
  new DOMParser().parseFromString(`<root xmlns:m="${M}" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">${xml}</root>`, "application/xml").documentElement.firstElementChild!;

/** LaTeX → OMML → LaTeX → OMML: the second OMML must equal the first. */
function roundTrip(latex: string, display = false): { omml: string; back: string } {
  const omml = latexToOmml(latex, display);
  const back = ommlToLatex(parseOmml(omml));
  expect(back).not.toBeNull();
  expect(latexToOmml(back!, display)).toBe(omml);
  return { omml, back: back! };
}

describe("LaTeX → OMML", () => {
  it("fractions, scripts and roots", () => {
    const { omml } = roundTrip("x=\\frac{-b\\pm\\sqrt{b^2-4ac}}{2a}");
    expect(omml).toContain("<m:f><m:num>");
    expect(omml).toContain("<m:rad><m:radPr><m:degHide m:val=\"1\"/></m:radPr><m:deg/>");
    expect(omml).toContain("<m:sSup><m:e><m:r>");
    expect(omml).toContain(">−4ac<"); // minus sign, as Word writes it
    expect(roundTrip("\\sqrt[3]{x}").omml).toContain("<m:deg><m:r>");
    expect(roundTrip("x_{i}^{2}").omml).toContain("<m:sSubSup>");
    expect(roundTrip("a_n").omml).toContain("<m:sSub>");
  });

  it("sums, integrals and functions take their operand", () => {
    const sum = roundTrip("\\sum_{k=0}^{n}\\binom{n}{k}x^ka^{n-k}", true).omml;
    expect(sum).toContain("<m:oMathPara>");
    // As Word writes it: limLoc only where it differs from the default (here: none).
    expect(sum).toContain('<m:naryPr><m:chr m:val="∑"/></m:naryPr>');
    expect(sum).toContain('<m:type m:val="noBar"/>');
    const int = roundTrip("\\int_0^1 x^2\\,dx=\\frac{1}{3}").omml;
    expect(int).toMatch(/<m:nary><m:naryPr><\/m:naryPr><m:sub>.*<\/m:sub><m:sup>.*<\/m:sup><m:e>.*dx.*<\/m:e><\/m:nary><m:r>.*=/);
    const sin = roundTrip("\\sin^2\\alpha+\\cos^2\\alpha=1").omml;
    expect((sin.match(/<m:func>/g) ?? []).length).toBe(2);
    expect(sin).toContain('<m:fName><m:sSup><m:e><m:r><m:rPr><m:sty m:val="p"/></m:rPr>');
    const lim = roundTrip("\\lim_{x\\to 0}\\frac{\\sin x}{x}=1").omml;
    expect(lim).toContain("<m:fName><m:limLow>");
  });

  it("brackets, matrices, cases, accents and text", () => {
    expect(roundTrip("\\left(a+b\\right)^2").omml).toContain("<m:d><m:e>");
    expect(roundTrip("\\left[0;1\\right]").omml).toContain('<m:begChr m:val="["/><m:endChr m:val="]"/>');
    expect(roundTrip("\\left|x\\right|").omml).toContain('<m:begChr m:val="|"/>');
    const mat = roundTrip("A=\\begin{pmatrix}1&2\\\\3&4\\end{pmatrix}").omml;
    expect(mat).toContain('<m:count m:val="2"/>');
    expect((mat.match(/<m:mr>/g) ?? []).length).toBe(2);
    expect(roundTrip("f(x)=\\begin{cases}x&x\\ge 0\\\\-x&x<0\\end{cases}").omml).toContain("<m:eqArr>");
    expect(roundTrip("\\vec{v}+\\overline{AB}+\\hat{x}").omml).toContain('<m:chr m:val="⃗"/>');
    expect(roundTrip("v=\\frac{s}{t}\\quad\\text{mit }t>0").omml).toContain("<m:nor/>");
    expect(roundTrip("\\mathbb{R}").omml).toContain('<m:scr m:val="double-struck"/>');
    expect(roundTrip("\\alpha\\cdot\\beta\\le\\infty").omml).toContain(">α⋅β≤∞<");
  });

  it("several lines become an equation array", () => {
    const { omml } = roundTrip("a&=b\\\\c&=d", true);
    expect(omml).toContain("<m:eqArr>");
  });
});

describe("OMML written by Word", () => {
  it("reads a fraction with Word's run properties", () => {
    const xml = '<m:oMath><m:f><m:fPr><m:ctrlPr><w:rPr><w:rFonts w:ascii="Cambria Math"/><w:i/></w:rPr></m:ctrlPr></m:fPr>' +
      '<m:num><m:r><w:rPr><w:rFonts w:ascii="Cambria Math"/></w:rPr><m:t>a</m:t></m:r></m:num><m:den><m:r><m:t>b</m:t></m:r></m:den></m:f>' +
      '<m:r><m:t>=</m:t></m:r><m:sSup><m:sSupPr><m:ctrlPr/></m:sSupPr><m:e><m:r><m:t>x</m:t></m:r></m:e><m:sup><m:r><m:t>2</m:t></m:r></m:sup></m:sSup></m:oMath>';
    expect(ommlToLatex(parseOmml(xml))).toBe("\\frac{a}{b}=x^{2}");
  });

  it("gives up on unknown elements", () => {
    expect(ommlToLatex(parseOmml("<m:oMath><m:foo/></m:oMath>"))).toBeNull();
  });
});

describe("equations in .docx", () => {
  it("writes new equations as OMML and reads them back editable", () => {
    const inline = { kind: "math" as const, data: undefined, display: "", latex: "a^2+b^2=c^2" };
    const display = { kind: "math" as const, data: undefined, display: "", latex: "A=\\pi r^2" };
    const d = new Doc({
      paras: [
        { ...makePara(`Satz ${OBJ} gilt.`), objects: [inline] },
        { ...makePara(OBJ), objects: [display] },
      ],
    });
    const bytes = writeDocx(d.state(), null);
    const x = strFromU8(unzipSync(bytes)["word/document.xml"]);
    expect(x).toContain('xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"');
    expect(x).toMatch(/<w:t xml:space="preserve">Satz <\/w:t><\/w:r><m:oMath>/);
    expect(x).toContain("<m:oMathPara>");
    const back = readDocx(bytes);
    expect(back.partial).not.toContain("Formeln");
    const objs = back.state.paras.flatMap((p) => p.objects ?? []);
    expect(objs.map((o) => o.latex)).toEqual(["a^{2}+b^{2}=c^{2}", "A=\\pi r^{2}"]);
  });
});

describe("Formeloptionen (m:mathPr)", () => {
  it("a new document carries Word's defaults; changed ones are read back and decide limLoc", () => {
    const sum = parseLatex("\\sum_{i=1}^{n}i");
    const d = new Doc({ paras: [{ ...makePara(OBJ), objects: [{ kind: "math", data: undefined, display: "", math: sum }] }] });
    const first = unzipSync(writeDocx(d.state(), null));
    expect(strFromU8(first["word/settings.xml"])).toContain('<m:naryLim m:val="undOvr"/>');
    expect(readDocx(writeDocx(d.state(), null)).state.settings.math).toEqual(DEFAULT_MATH);
    // Limits beside sums by default: a sum with limits above and below now says so.
    d.setSettings({ ...d.settings, math: { ...DEFAULT_MATH, naryLim: "subSup", defJc: "left" } });
    const bytes = writeDocx(d.state(), null);
    const x = strFromU8(unzipSync(bytes)["word/document.xml"]);
    expect(x).toContain('<m:limLoc m:val="undOvr"/>');
    expect(strFromU8(unzipSync(bytes)["word/settings.xml"])).toContain('<m:defJc m:val="left"/>');
    const back = readDocx(bytes).state;
    expect(back.settings.math?.naryLim).toBe("subSup");
    expect(back.paras[0].objects?.[0].math?.[0]).toMatchObject({ k: "nary", und: true });
  });
});

describe("Word's equations written back (word-math.docx)", () => {
  it("every equation reads editable and is written with Word's structure", () => {
    const bytes = new Uint8Array(readFileSync("packages/ooxml/test/fixtures/word-math.docx"));
    const files = unzipSync(bytes);
    const settings = new DOMParser().parseFromString(strFromU8(files["word/settings.xml"]), "application/xml").documentElement;
    const opts = readMathPr(settings) ?? {};
    const doc = new DOMParser().parseFromString(strFromU8(files["word/document.xml"]), "application/xml");
    const maths = [...Array.from(doc.getElementsByTagNameNS(M_NS, "oMathPara")),
      ...Array.from(doc.getElementsByTagNameNS(M_NS, "oMath")).filter((m) => m.parentElement?.localName !== "oMathPara")];
    expect(maths.length).toBe(24);
    const diffs: string[] = [];
    for (const el of maths) {
      const seq = ommlToSeq(el, opts);
      if (!seq) { diffs.push(`unreadable: ${norm(el, KEEP_ALL)}`); continue; }
      const xml = seqToOmml(seq, el.localName === "oMathPara", opts);
      const back = new DOMParser().parseFromString(`<r xmlns:m="${M_NS}" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">${xml}</r>`, "application/xml").documentElement.firstElementChild!;
      const want = norm(el, KEEP_ALL), got = norm(back, KEEP_ALL);
      if (want !== got) diffs.push(`word: ${want}\nours: ${got}`);
    }
    expect(diffs.join("\n\n")).toBe("");
  });
});
