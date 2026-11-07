// The galleries named as PowerPoint's recorded ones (D-072).

import { describe, expect, it } from "vitest";
import { PptxDom } from "@raabe/ooxml";
import { tableStyles } from "../src/tablestyles";
import { QUICK_ROWS, applyQuickStyle, colorWord, quickName } from "../src/quickstyles";
import { PIC_STYLES } from "../src/picstyles";

describe("galleries", () => {
  it("has PowerPoint's 74 table styles in its order, with unique ids", () => {
    const all = tableStyles();
    expect(all.length).toBe(74);
    expect(new Set(all.map((s) => s.id)).size).toBe(74);
    expect(all.slice(0, 8).map((s) => s.name)).toEqual(["Keine Formatvorlage, kein Raster", "Designformatvorlage 1 - Akzent 1", "Designformatvorlage 1 - Akzent 2",
      "Designformatvorlage 1 - Akzent 3", "Designformatvorlage 1 - Akzent 4", "Designformatvorlage 1 - Akzent 5", "Designformatvorlage 1 - Akzent 6", "Keine Formatvorlage, Tabellenraster"]);
    expect(all.at(-1)!.name).toBe("Dunkle Formatvorlage 2 - Akzent 5/Akzent 6");
    // PowerPoint's default keeps PowerPoint's own id.
    expect(all.find((s) => s.name === "Mittlere Formatvorlage 2 - Akzent 1")!.id).toBe("{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}");
    for (const s of all) expect(() => PptxDom.parseXml(`<a:tblStyleLst xmlns:a="${PptxDom.NS_A}">${s.xml}</a:tblStyleLst>`)).not.toThrow();
  });

  it("names quick styles after the design's colours, as PowerPoint does", () => {
    const office2013 = { tx1: "000000", accent1: "4472C4", accent2: "ED7D31", accent3: "A5A5A5", accent4: "FFC000", accent5: "5B9BD5", accent6: "70AD47" };
    expect(QUICK_ROWS.length * 7).toBe(77);
    expect(quickName(QUICK_ROWS[0], 0, office2013)).toBe("Farbige Kontur - Schwarz, Dunkel 1");
    expect(quickName(QUICK_ROWS[9], 3, office2013)).toBe("Farbige Füllung - Grau, Akzent 3, Keine Kontur");
    // PowerPoint's default design (Aptos): the names recorded in its gallery.
    expect(["156082", "E97132", "196B24", "0F9ED5", "A02B93", "4EA72E"].map(colorWord)).toEqual(["Dunkelblaugrün", "Orange", "Dunkelgrün", "Türkis", "Pflaume", "Grün"]);
  });

  it("writes a quick style as p:style after spPr and a preset's fill on the shape", () => {
    const sp = PptxDom.parseXml(`<p:sp xmlns:p="${PptxDom.NS_P}" xmlns:a="${PptxDom.NS_A}"><p:nvSpPr/><p:spPr><a:xfrm/><a:prstGeom prst="rect"/><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></p:spPr><p:txBody/></p:sp>`).documentElement;
    applyQuickStyle(sp, QUICK_ROWS[9], 1);
    const names = Array.from(sp.children).map((c) => c.localName);
    expect(names).toEqual(["nvSpPr", "spPr", "style", "txBody"]);
    expect(Array.from(PptxDom.kid(sp, "spPr")!.children).map((c) => c.localName)).toEqual(["xfrm", "prstGeom", "solidFill", "ln"]);
  });

  it("has PowerPoint's 28 picture styles", () => {
    expect(PIC_STYLES.length).toBe(28);
    expect(PIC_STYLES[0].name).toBe("Einfacher Rahmen, weiß");
    expect(PIC_STYLES.at(-1)!.name).toBe("Metallenes Oval");
  });
});
