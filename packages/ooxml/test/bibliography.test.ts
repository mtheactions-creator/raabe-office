// Zitate und Literaturverzeichnis: Word's sources part, citations and entries per style.

import { describe, expect, it } from "vitest";
import { strFromU8, unzipSync } from "fflate";
import { Doc, formatCitation, formatEntry, makePara, parsePeople, suggestTag, type Bibliography, type Source } from "@raabe/text-engine";
import { readDocx, writeDocx } from "../src";

const book: Source = {
  tag: "Mus23", type: "Book", contributors: { Author: { people: parsePeople("Muster, Max; Beispiel, Erika") } },
  fields: { Title: "Photosynthese verstehen", Year: "2023", City: "Hannover", Publisher: "Schulbuchverlag" },
};

describe("citations", () => {
  it("are formatted per style", () => {
    expect(formatCitation(book, "Mus23", "APA", 1)).toBe("(Muster & Beispiel, 2023)");
    expect(formatCitation(book, "Mus23", "IEEE", 2)).toBe("[2]");
    expect(formatCitation(book, "Mus23", "MLA", 1, "12")).toBe("(Muster und Beispiel 12)");
    expect(formatCitation(undefined, "Platzhalter1", "APA", 1)).toBe("(Platzhalter1)");
    expect(formatEntry(book, "APA", 1).map((p) => p.text).join("")).toBe("Muster, M., & Beispiel, E. (2023). Photosynthese verstehen. Hannover: Schulbuchverlag.");
    expect(formatEntry(book, "APA", 1).find((p) => p.italic)?.text).toBe("Photosynthese verstehen.");
    expect(suggestTag(book, ["Mus23"])).toBe("Mus232");
  });

  it("sources travel in Word's customXml part and come back", () => {
    const bib: Bibliography = { sources: [book], style: "MLA" };
    const d = new Doc({ paras: [makePara("Text")] });
    d.setSettings({ ...d.settings, bibliography: bib });
    const bytes = writeDocx(d.state(), null);
    const files = unzipSync(bytes);
    const item = Object.keys(files).find((f) => /^customXml\/item\d+\.xml$/.test(f))!;
    expect(strFromU8(files[item])).toContain('StyleName="MLA"');
    expect(strFromU8(files["word/_rels/document.xml.rels"])).toContain("customXml/item1.xml");
    expect(strFromU8(files["[Content_Types].xml"])).toContain("customXmlProperties+xml");
    const back = readDocx(bytes).state.settings.bibliography!;
    expect(back.style).toBe("MLA");
    expect(back.sources[0]).toMatchObject({ tag: "Mus23", type: "Book", fields: book.fields });
    expect(back.sources[0].contributors?.Author.people?.map((p) => p.last)).toEqual(["Muster", "Beispiel"]);
  });
});
