// Design colours from the palette (D-072): written as PowerPoint writes them, shown as Office's palette shows them.
import { describe, expect, it } from "vitest";
import { colorXml, specHex } from "../src/colors";

const OFFICE_2013 = { lt1: "FFFFFF", dk1: "000000", lt2: "E7E6E6", dk2: "44546A", accent1: "4472C4", accent2: "ED7D31" };

describe("design colours", () => {
  it("are written as schemeClr with lumMod/lumOff", () => {
    expect(colorXml("accent1,40000,60000")).toBe('<a:schemeClr val="accent1"><a:lumMod val="40000"/><a:lumOff val="60000"/></a:schemeClr>');
    expect(colorXml("ACCENT1,75000")).toBe('<a:schemeClr val="accent1"><a:lumMod val="75000"/></a:schemeClr>');
    expect(colorXml("tx1")).toBe('<a:schemeClr val="tx1"></a:schemeClr>');
    expect(colorXml("ff0000")).toBe('<a:srgbClr val="FF0000"/>');
  });
  it("resolve to Office's palette colours", () => {
    expect(specHex("accent2,60000,40000", OFFICE_2013)).toBe("F4B183");
    expect(specHex("accent1,75000", OFFICE_2013)).toBe("2F5597");
    expect(specHex("bg1", OFFICE_2013)).toBe("FFFFFF");
  });
});
