import { readFileSync } from "node:fs";
import { describe, expect, it } from "vitest";
import { strToU8, unzipSync, zipSync, strFromU8 } from "fflate";
import { readDocx, writeDocx, type DocxParagraph } from "../src";
import fixtures from "../../../prototypes/src/shared/fixtures.json";

const sample: DocxParagraph[] = [
  { style: "h1", text: "Überschrift mit „Anführungszeichen“ & <Sonderzeichen>" },
  { style: "normal", text: "  Zwei führende Leerzeichen, ein\tTabulator und Umlaute: äöüß." },
  { style: "normal", text: "" },
];

describe("writeDocx / readDocx", () => {
  it("round-trips paragraphs, styles, spaces and tabs", () => {
    const back = readDocx(writeDocx(sample));
    expect(back.unsupported).toEqual([]);
    expect(back.paragraphs).toEqual(sample);
  });

  it("reads a document saved by Microsoft Word", () => {
    const bytes = new Uint8Array(readFileSync("packages/ooxml/test/fixtures/word-short.docx"));
    const result = readDocx(bytes);
    expect(result.unsupported).toEqual([]);
    expect(result.paragraphs).toEqual(fixtures.base);
  });

  it("reports content it cannot keep", () => {
    const files = unzipSync(writeDocx(sample));
    const xml = strFromU8(files["word/document.xml"]).replace(
      "<w:body>",
      "<w:body><w:tbl/><w:p><w:r><w:rPr><w:b/></w:rPr><w:t>fett</w:t></w:r></w:p>",
    );
    files["word/document.xml"] = strToU8(xml);
    const result = readDocx(zipSync(files));
    expect(result.unsupported).toContain("Tabellen");
    expect(result.unsupported).toContain("Zeichenformatierung (z. B. fett, kursiv, Farbe)");
  });

  it("rejects files that are not Word documents", () => {
    expect(() => readDocx(zipSync({ "hello.txt": strToU8("hi") }))).toThrow();
  });
});
