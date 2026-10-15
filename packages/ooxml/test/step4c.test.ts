// Step 4c: wrap polygons (Eng, Transparent), Transparente Farbe bestimmen, equations on the clipboard.
import { describe, expect, it } from "vitest";
import { readFileSync } from "node:fs";
import { join } from "node:path";
import { strFromU8, unzipSync } from "fflate";
import { seqToLinear } from "@raabe/text-engine";
import { readDocx, styleFromXml, themedHex, writeDocx } from "../src";
import { blipEffectsXml, readPicture } from "../src/drawingml";
import { ommlToSeq, ommlToWordHtml, parseLatex, seqToOmml, wordHtmlMath } from "../src/math";

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

describe("Formeln in der Zwischenablage (Word's HTML)", () => {
  const word = JSON.parse(readFileSync(join(__dirname, "fixtures", "word-math-clipboard.json"), "utf8")) as Record<string, string>;

  it("reads Word's equations: fraction and power, function and normal text, the quadratic formula", () => {
    const frac = wordHtmlMath(word.text_with_equation)!;
    expect(frac.localName).toBe("oMath");
    expect(seqToLinear(ommlToSeq(frac)!)).toBe("a/b+x^2");
    const plain = wordHtmlMath(word.plain_runs)!;
    expect(plain.localName).toBe("oMathPara");
    const seq = ommlToSeq(plain)!;
    expect(seq.some((n) => n.k === "func")).toBe(true);
    expect(seq.some((n) => n.k === "t" && n.nor && n.s === "wenn")).toBe(true);
    expect(seqToLinear(ommlToSeq(wordHtmlMath(word.display_equation)!)!)).toContain("√");
  });

  it("writes equations Word reads back to the same tree", () => {
    for (const latex of ["\frac{a}{b}+x^{2}", "\sqrt{x}=\sin x", "\sum_{i=1}^{n} i", "x_{1}"]) {
      const seq = parseLatex(latex);
      const html = ommlToWordHtml(seqToOmml(seq, false));
      expect(html).toContain("<m:oMath>");
      expect(html).not.toContain("<m:t");
      const back = ommlToSeq(wordHtmlMath(html)!)!;
      expect(seqToLinear(back)).toBe(seqToLinear(seq));
    }
  });
});

describe("Tabellenformatvorlagen (Word's theme colours)", () => {
  it("computes Word's tinted and shaded theme colours", () => {
    const samples: [string, string, string, string][] = [["156082", "", "", "156082"], ["156082", "Tint", "19", "E0F2FA"], ["156082", "Tint", "7F", "64BDE6"], ["156082", "Tint", "99", "45B0E1"],
      ["E97132", "Shade", "7F", "7F340D"], ["E97132", "Tint", "19", "FCF0EA"], ["E97132", "Tint", "7F", "F4B798"], ["196B24", "Shade", "7F", "0C3511"],
      ["196B24", "Tint", "19", "E0F8E3"], ["196B24", "Tint", "7F", "66DB75"], ["0F9ED5", "Shade", "7F", "074E69"], ["0F9ED5", "Tint", "19", "E5F6FD"],
      ["0F9ED5", "Tint", "7F", "7BD3F5"], ["A02B93", "Shade", "7F", "4F1548"], ["A02B93", "Tint", "19", "F8E7F6"], ["A02B93", "Tint", "7F", "DE86D4"],
      ["4EA72E", "Shade", "7F", "265317"], ["4EA72E", "Tint", "19", "ECF8E8"], ["4EA72E", "Tint", "7F", "A0DF8A"], ["FFFFFF", "Shade", "BF", "BFBFBF"],
      ["000000", "Shade", "7F", "000000"], ["000000", "Tint", "33", "CCCCCC"], ["000000", "Tint", "80", "7F7F7F"]];
    for (const [base, kind, v, want] of samples) {
      const got = themedHex(base, kind === "Tint" ? v : undefined, kind === "Shade" ? v : undefined);
      const diff = Math.max(...[0, 2, 4].map((i) => Math.abs(parseInt(got.slice(i, i + 2), 16) - parseInt(want.slice(i, i + 2), 16))));
      expect(diff, `${base} ${kind} ${v}: ${got} vs ${want}`).toBeLessThanOrEqual(1);
    }
  });

  it("reads a gallery style with the document's theme", () => {
    const xml = '<w:style w:type="table" w:styleId="X"><w:name w:val="Grid Table 4 Accent 1"/><w:tblPr><w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="45B0E1" w:themeColor="accent1" w:themeTint="99"/></w:tblBorders></w:tblPr>' +
      '<w:tblStylePr w:type="firstRow"><w:tcPr><w:shd w:val="clear" w:color="auto" w:fill="156082" w:themeFill="accent1"/></w:tcPr></w:tblStylePr></w:style>';
    const def = styleFromXml(xml, { accent1: "4472C4" })!;
    expect(def.name).toBe("Grid Table 4 Accent 1");
    expect(def.table?.cond?.firstRow?.shading?.fill).toBe("4472C4");
    expect(def.table?.borders?.top?.color).toBe(themedHex("4472C4", "99"));
  });
});

describe("Umgebende Tabellen (word-table-float.docx, made by Word)", () => {
  const pkg = readDocx(fixture("word-table-float.docx"));
  const tables = [...new Map(pkg.state.paras.filter((p) => p.cell).map((p) => [p.cell!.table.id, p.cell!.table])).values()];

  it("reads Word's w:tblpPr", () => {
    expect(tables.length).toBe(3);
    expect(tables[0].float).toMatchObject({ xAlign: "center", hAnchor: "margin", vAnchor: "text", left: 7.05, right: 7.05 });
    expect(tables[0].float!.y).toBeCloseTo(6.05, 2);
    expect(tables[1].float).toMatchObject({ x: 60.05, hAnchor: "margin" });
    expect(tables[2].float?.xAlign).toBe("right");
  });

  it("writes them back as Word wrote them", () => {
    const xml = strFromU8(unzipSync(writeDocx(pkg.state, pkg.pkg))["word/document.xml"]);
    expect(xml).toContain('<w:tblpPr w:leftFromText="141" w:rightFromText="141" w:vertAnchor="text" w:horzAnchor="margin" w:tblpXSpec="center" w:tblpY="121"/>');
    const back = readDocx(writeDocx(pkg.state, pkg.pkg)).state.paras.filter((p) => p.cell).map((p) => p.cell!.table.float);
    expect(back[0]).toEqual(tables[0].float);
  });
});
