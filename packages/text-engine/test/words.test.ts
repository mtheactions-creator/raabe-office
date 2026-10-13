import { describe, expect, it } from "vitest";
import { nextTokenStart, prevTokenStart, sentenceAt, tokenAt } from "../src/words";
import { Doc } from "../src/doc";

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

describe("Doc", () => {
  it("splits and joins paragraphs, and Enter after a heading continues with Standard", () => {
    const d = new Doc([{ style: "h1", text: "Titel" }]);
    const end = d.replace({ p: 0, o: 5 }, { p: 0, o: 5 }, "\nText");
    expect(d.paras).toEqual([{ style: "h1", text: "Titel" }, { style: "normal", text: "Text" }]);
    expect(end).toEqual({ p: 1, o: 4 });
    d.replace({ p: 0, o: 5 }, { p: 1, o: 0 }, "");
    expect(d.text()).toBe("TitelText");
  });
  it("maps offsets to positions and back", () => {
    const d = new Doc([{ style: "normal", text: "ab" }, { style: "normal", text: "cde" }]);
    expect(d.fromOffset(4)).toEqual({ p: 1, o: 1 });
    expect(d.toOffset({ p: 1, o: 1 })).toBe(4);
  });
});
