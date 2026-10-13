import { describe, expect, it } from "vitest";
import { countWords, nextTokenStart, prevTokenStart, sentenceAt, tokenAt } from "../src/words";
import { Doc } from "../src/doc";
import { makePara, validPara } from "../src/model";

const t = "Pflanzen, Algen und (Glucose) geht's.";

describe("Word's word rules", () => {
  it("double-click selects the word with its trailing space", () => {
    expect(t.slice(...tokenAt(t, t.indexOf("Algen") + 2))).toBe("Algen ");
  });
  it("a word before punctuation is selected without it", () => {
    expect(t.slice(...tokenAt(t, 2))).toBe("Pflanzen");
  });
  it("an apostrophe inside a word belongs to the word", () => {
    expect(t.slice(...tokenAt(t, t.indexOf("geht")))).toBe("geht's");
  });
  it("Strg+Rechts treats punctuation as its own word", () => {
    expect(nextTokenStart(t, 0)).toBe(t.indexOf(","));
    expect(nextTokenStart(t, t.indexOf(","))).toBe(t.indexOf("Algen"));
  });
  it("Strg+Links goes to the start of the current or previous word", () => {
    expect(prevTokenStart(t, t.indexOf("und") + 1)).toBe(t.indexOf("und"));
    expect(prevTokenStart(t, t.indexOf("und"))).toBe(t.indexOf("Algen"));
  });
  it("Strg+Klick selects the sentence with trailing spaces", () => {
    const s = "Erster Satz. Zweiter Satz! Dritter.";
    expect(s.slice(...sentenceAt(s, s.indexOf("Zweiter") + 3))).toBe("Zweiter Satz! ");
  });
});

describe("word count", () => {
  it("counts like Word: standalone dashes are not words, other symbols are", () => {
    expect(countWords(["a – b", "a — b"])).toBe(4);
    expect(countWords(["a - b", "a + b", "a → b", "6 CO2"])).toBe(11);
    expect(countWords(["", "  ", "Wort"])).toBe(1);
  });
});

const NL = String.fromCharCode(10);

describe("Doc", () => {
  const plain = (d: Doc) => d.paras.map((p) => ({ style: p.props.style ?? "Standard", text: p.text }));
  it("splits and joins paragraphs, and Enter after a heading continues with Standard", () => {
    const d = new Doc({ paras: [makePara("Titel", { style: "berschrift1" })] });
    const end = d.insertText({ p: 0, o: 5 }, { p: 0, o: 5 }, NL + "Text", {});
    expect(plain(d)).toEqual([{ style: "berschrift1", text: "Titel" }, { style: "Standard", text: "Text" }]);
    expect(end).toEqual({ p: 1, o: 4 });
    d.delete({ p: 0, o: 5 }, { p: 1, o: 0 });
    expect(d.text()).toBe("TitelText");
    expect(d.paras.every(validPara)).toBe(true);
  });
  it("splitting in the middle keeps the style in both halves", () => {
    const d = new Doc({ paras: [makePara("AbcDef", { style: "berschrift1" })] });
    d.insertText({ p: 0, o: 3 }, { p: 0, o: 3 }, NL, {});
    expect(plain(d)).toEqual([{ style: "berschrift1", text: "Abc" }, { style: "berschrift1", text: "Def" }]);
  });
  it("maps offsets to positions and back", () => {
    const d = new Doc({ paras: [makePara("ab"), makePara("cde")] });
    expect(d.fromOffset(4)).toEqual({ p: 1, o: 1 });
    expect(d.toOffset({ p: 1, o: 1 })).toBe(4);
  });
  it("formats runs and merges equal neighbours", () => {
    const d = new Doc({ paras: [makePara("Ein fettes Wort")] });
    d.mapChars({ p: 0, o: 4 }, { p: 0, o: 10 }, (o) => ({ ...o, bold: true }));
    expect(d.paras[0].runs.map((r) => [r.len, !!r.props.bold])).toEqual([[4, false], [6, true], [5, false]]);
    d.mapChars({ p: 0, o: 0 }, { p: 0, o: 15 }, (o) => ({ ...o, bold: undefined }));
    expect(d.paras[0].runs.length).toBe(1);
    expect(d.paras.every(validPara)).toBe(true);
  });
});
