import { describe, expect, it } from "vitest";
import { Doc } from "../src/doc";
import { intern, makePara } from "../src/model";
import { defaultStyleSheet } from "../src/styles";

const NL = String.fromCharCode(10);

describe("style resolution", () => {
  const s = defaultStyleSheet();
  it("new documents use the theme's body font, headings the heading font", () => {
    expect(s.resolveChar("Standard", {}).font).toBe("Calibri");
    expect(s.resolveChar("berschrift1", {}).font).toBe("Calibri Light");
    expect(s.resolveChar("berschrift1", {}).size).toBe(16);
  });
  it("a font name set directly replaces the inherited theme font", () => {
    expect(s.resolveChar("Standard", intern({ font: "Georgia" })).font).toBe("Georgia");
    expect(s.resolveChar("berschrift1", intern({ font: "Arial" })).font).toBe("Arial");
  });
  it("a character style toggles bold off inside a bold paragraph style (as in Word)", () => {
    const sheet = defaultStyleSheet().withStyle({
      id: "Fettabsatz", name: "Fettabsatz", type: "paragraph", basedOn: "Standard", pPr: {}, rPr: { bold: true },
    });
    expect(sheet.resolveChar("Standard", intern({ rStyle: "Fett" })).bold).toBe(true);
    expect(sheet.resolveChar("Fettabsatz", intern({ rStyle: "Fett" })).bold).toBe(false);
    // Direct formatting is not a toggle: it sets the value.
    expect(sheet.resolveChar("Fettabsatz", intern({ bold: true })).bold).toBe(true);
  });
});

describe("paragraph marks", () => {
  const bold = (d: Doc, p: number) => !!d.paras[p].mark.bold;
  it("formatting the text of a whole paragraph leaves its mark alone (as in Word)", () => {
    const d = new Doc({ paras: [makePara("Hallo")] });
    d.mapChars({ p: 0, o: 0 }, { p: 0, o: 5 }, (o) => ({ ...o, bold: true }));
    expect(bold(d, 0)).toBe(false);
  });
  it("typing at a paragraph end formats the mark, and Enter copies it", () => {
    const d = new Doc({ paras: [makePara("Hallo")] });
    d.insertText({ p: 0, o: 5 }, { p: 0, o: 5 }, " fett", intern({ bold: true }));
    expect(bold(d, 0)).toBe(true);
    const end = d.insertText({ p: 0, o: 10 }, { p: 0, o: 10 }, NL, intern({ bold: true }));
    expect(bold(d, 0)).toBe(true);
    expect(bold(d, 1)).toBe(true);
    expect(d.charPropsAt(end).bold).toBe(true);
  });
  it("formatting part of a paragraph leaves the mark alone", () => {
    const d = new Doc({ paras: [makePara("Hallo")] });
    d.mapChars({ p: 0, o: 1 }, { p: 0, o: 5 }, (o) => ({ ...o, bold: true }));
    expect(bold(d, 0)).toBe(false);
  });
  it("a range running into the next paragraph formats the marks it passes", () => {
    const d = new Doc({ paras: [makePara("Eins"), makePara("Zwei")] });
    d.mapChars({ p: 0, o: 2 }, { p: 1, o: 2 }, (o) => ({ ...o, italic: true }));
    expect(!!d.paras[0].mark.italic).toBe(true);
    expect(!!d.paras[1].mark.italic).toBe(false);
  });
  it("Enter at the end: the new mark ends the first paragraph, the old one moves on", () => {
    const d = new Doc({ paras: [makePara("Hallo")] });
    const end = d.insertText({ p: 0, o: 5 }, { p: 0, o: 5 }, NL, intern({ bold: true }));
    expect(end).toEqual({ p: 1, o: 0 });
    expect(bold(d, 0)).toBe(true);
    expect(bold(d, 1)).toBe(false);
    expect(d.charPropsAt(end).bold).toBeFalsy();
  });
});
