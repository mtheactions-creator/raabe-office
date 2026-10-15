// Step 4c: wrap polygons (Eng, Transparent), Transparente Farbe bestimmen.
import { describe, expect, it } from "vitest";
import { readFileSync } from "node:fs";
import { join } from "node:path";
import { strFromU8, unzipSync } from "fflate";
import { readDocx, writeDocx } from "../src";
import { blipEffectsXml, readPicture } from "../src/drawingml";

const fixture = (name: string) => new Uint8Array(readFileSync(join(__dirname, "fixtures", name)));

describe("Textumbruch Eng (word-wrap.docx, made by Word)", () => {
  const pkg = readDocx(fixture("word-wrap.docx"));
  const floats = pkg.state.paras.flatMap((p) => p.floats ?? []);

  it("reads Word's wrap polygons of Eng, none for Quadrat", () => {
    expect(floats.length).toBe(9);
    expect(floats.filter((f) => f.wrap === "square").every((f) => !f.polygon)).toBe(true);
    const tight = floats.filter((f) => f.wrap === "tight");
    expect(tight.length).toBe(4);
    expect(tight.every((f) => (f.polygon?.length ?? 0) >= 4)).toBe(true);
    // The cut-out circle: its polygon stays within the box, first point near the top middle.
    const circle = tight.find((f) => f.width === 110)!;
    expect(circle.polygon![0][1]).toBeLessThan(1000);
    expect(Math.abs(circle.polygon![0][0] - 10800)).toBeLessThan(3000);
  });

  it("writes an edited Eng picture with its polygon", () => {
    const paras = pkg.state.paras.map((p) => (p.floats ? { ...p, floats: p.floats.map((f) => ({ ...f, data: f.image ? undefined : f.data })) } : p));
    const xml = strFromU8(unzipSync(writeDocx({ ...pkg.state, paras }, pkg.pkg))["word/document.xml"]);
    const circle = floats.find((f) => f.wrap === "tight" && f.width === 110)!;
    const [x, y] = circle.polygon![0];
    expect(xml).toContain(`<wp:wrapTight wrapText="bothSides"><wp:wrapPolygon edited="0"><wp:start x="${x}" y="${y}"/>`);
    const back = readDocx(writeDocx({ ...pkg.state, paras }, pkg.pkg)).state.paras.flatMap((p) => p.floats ?? []);
    expect(back.find((f) => f.wrap === "tight" && f.width === 110)?.polygon).toEqual(circle.polygon);
  });
});

describe("Transparente Farbe bestimmen", () => {
  it("writes and reads Word's a:clrChange", () => {
    const xml = blipEffectsXml({ transparentColor: "ebf5e6" });
    expect(xml).toBe('<a:clrChange><a:clrFrom><a:srgbClr val="EBF5E6"/></a:clrFrom><a:clrTo><a:srgbClr val="EBF5E6"><a:alpha val="0"/></a:srgbClr></a:clrTo></a:clrChange>');
    const pic = new DOMParser().parseFromString(
      `<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
      `<pic:blipFill><a:blip r:embed="rId1">${xml}</a:blip></pic:blipFill><pic:spPr/></pic:pic>`, "application/xml");
    expect(readPicture(pic.documentElement, {}).transparentColor).toBe("EBF5E6");
  });
});
