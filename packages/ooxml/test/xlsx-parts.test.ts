import { describe, expect, it } from "vitest";
import { tinted } from "../src/xlsx/colors";
import { shiftFormula } from "../src/xlsx/shift";
import { widthFromBase, widthFromFile, widthToFile } from "../src/xlsx/units";
import { decodeX, encodeX, setAttrs, split, joinSplit, place } from "../src/xlsx/xml";
import { ownRange, sheetRef } from "../src/xlsx/sheetparts";

describe("theme tints", () => {
  it("lighten and darken as Excel computes them (Windows HLS)", () => {
    // Excel's own colours for these theme colours and tints (xlsx-styles.xlsx, xlsx-sheets.xlsx).
    expect(tinted("4472C4", 0.39988402966399123)).toBe("8EA9DB");
    expect(tinted("ED7D31", 0.39988402966399123)).toBe("F4B084");
    expect(tinted("000000", 0.49989318521683401)).toBe("808080");
    expect(tinted("A5A5A5", -0.24988555558946501)).toBe("7B7B7B");
    expect(tinted("4472C4", 0.499984740745262)).toBe("A2B8E1");
  });
});

describe("shared formulas", () => {
  it("move relative references only", () => {
    expect(shiftFormula("A1+$A$1+A$1+$A1", 2, 1)).toBe("B3+$A$1+B$1+$A3");
    expect(shiftFormula("SUM(Daten!A:A)+SUM(Daten!$1:1)", 3, 0)).toBe("SUM(Daten!A:A)+SUM(Daten!$1:4)");
    expect(shiftFormula('"A1"&Daten!A1', 1, 0)).toBe('"A1"&Daten!A2');
    expect(shiftFormula("LOG10(A1)+ATAN2(1,A1)", 1, 0)).toBe("LOG10(A2)+ATAN2(1,A2)");
    expect(shiftFormula("'Mein Blatt'!$A$1+H$1+Tabelle1!B2", 0, 1)).toBe("'Mein Blatt'!$A$1+I$1+Tabelle1!C2");
    expect(shiftFormula("SUM(Ergebnisse[Note])+1.5E+10+A1", 1, 0)).toBe("SUM(Ergebnisse[Note])+1.5E+10+A2");
    expect(shiftFormula("A1", -1, 0)).toBe("#REF!");
  });
});

describe("column widths", () => {
  it("convert between the file's width and characters", () => {
    expect(widthFromBase(10)).toBeCloseTo(10.714, 3); // German Excel's standard width (80 px)
    expect(widthFromBase(8)).toBeCloseTo(8.43, 2); // English Excel's (64 px)
    expect(widthToFile(30)).toBe(30.7109375);
    for (const w of [20.6328125, 3.1796875, 15.90625, 10.90625]) expect(widthToFile(widthFromFile(w))).toBe(w);
    expect(widthToFile(0)).toBe(0);
  });
});

describe("XML text", () => {
  it("escapes characters XML cannot hold as Excel does", () => {
    expect(encodeX("_x000D_ wörtlich")).toBe("_x005F_x000D_ wörtlich");
    expect(decodeX("_x005F_x000D_ wörtlich")).toBe("_x000D_ wörtlich");
    expect(encodeX(String.fromCharCode(1))).toBe("_x0001_");
  });

  it("sets attributes in place and puts parts in schema order", () => {
    expect(setAttrs('<pane xSplit="1" state="frozen"/>', { xSplit: null, ySplit: "2" })).toBe('<pane state="frozen" ySplit="2"/>');
    const s = split('<?xml version="1.0"?><r><a/><c/><x:y/><e/></r>');
    place(s, "d", "<d/>", ["a", "b", "c", "d", "e"]);
    place(s, "b", "<b/>", ["a", "b", "c", "d", "e"]);
    expect(joinSplit(s)).toBe('<?xml version="1.0"?><r><a/><b/><c/><x:y/><d/><e/></r>');
  });
});

describe("names of a sheet's ranges", () => {
  it("quote sheet names where Excel does", () => {
    expect(sheetRef("Tabelle1", { r0: 0, c0: 0, r1: 3, c1: 1 })).toBe("Tabelle1!$A$1:$B$4");
    expect(sheetRef("Mein Blatt", { r0: 1, c0: 0, r1: 1, c1: 0 })).toBe("'Mein Blatt'!$A$2");
    expect(sheetRef("Übersicht", { r0: 0, c0: 0, r1: 0, c1: 0 })).toBe("Übersicht!$A$1");
    expect(ownRange("'Mein Blatt'!$A$1:$B$4", "Mein Blatt")).toEqual({ r0: 0, c0: 0, r1: 3, c1: 1 });
    expect(ownRange("Erstes!$A$1:$B$4,Erstes!$D$1", "Erstes")).toBeNull();
  });
});
