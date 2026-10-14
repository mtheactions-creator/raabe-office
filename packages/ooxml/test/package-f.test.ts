// Step 4b F: language and proofing properties and Word's hyphenation settings, read from
// Word's files and written back as Word writes them.

import { describe, expect, it } from "vitest";
import { readFileSync } from "node:fs";
import { join } from "node:path";
import { strFromU8, unzipSync } from "fflate";
import { Doc, makePara } from "@raabe/text-engine";
import { readDocx, writeDocx } from "../src";

const fixture = (name: string) => new Uint8Array(readFileSync(join(__dirname, "fixtures", name)));
const part = (bytes: Uint8Array, path: string) => strFromU8(unzipSync(bytes)[path]);

describe("Silbentrennung", () => {
  it("reads Word's settings", () => {
    expect(readDocx(fixture("word-hyphen.docx")).state.settings.hyphenation).toEqual({ auto: true, caps: true, limit: 0, zone: 21.25 });
    expect(readDocx(fixture("word-hyphen-limit.docx")).state.settings.hyphenation).toEqual({ auto: true, caps: false, limit: 2, zone: 21.25 });
  });
  it("writes them where Word does and keeps the zone", () => {
    const pkg = readDocx(fixture("word-hyphen.docx"));
    const off = writeDocx({ ...pkg.state, settings: { ...pkg.state.settings, hyphenation: { auto: false, caps: true, limit: 0, zone: 21.25 } } }, pkg.pkg);
    expect(part(off, "word/settings.xml")).not.toContain("<w:autoHyphenation");
    const on = writeDocx({ ...pkg.state, settings: { ...pkg.state.settings, hyphenation: { auto: true, caps: false, limit: 3 } } }, pkg.pkg);
    const xml = part(on, "word/settings.xml");
    expect(xml).toMatch(/<w:defaultTabStop[^>]*\/><w:autoHyphenation\/><w:consecutiveHyphenLimit w:val="3"\/><w:hyphenationZone w:val="425"\/><w:doNotHyphenateCaps\/>/);
    expect(readDocx(on).state.settings.hyphenation).toEqual({ auto: true, caps: false, limit: 3, zone: 21.25 });
  });
  it("writes Keine Silbentrennung (w:suppressAutoHyphens)", () => {
    const d = new Doc({ paras: [makePara("Ohne Trennung", { suppressAutoHyphens: true })] });
    const xml = part(writeDocx(d.state(), null), "word/document.xml");
    expect(xml).toContain("<w:suppressAutoHyphens/>");
    expect(readDocx(writeDocx(d.state(), null)).state.paras[0].props.suppressAutoHyphens).toBe(true);
  });
});

describe("Sprache", () => {
  it("writes the language and Rechtschreibung nicht prüfen of runs", () => {
    const d = new Doc({ paras: [makePara("Hola amigos", {}, { lang: "es-ES" }), makePara("ohne Prüfung", {}, { noProof: true })] });
    const out = writeDocx(d.state(), null);
    const xml = part(out, "word/document.xml");
    expect(xml).toContain('<w:lang w:val="es-ES"/>');
    expect(xml).toContain("<w:noProof/>");
    const back = readDocx(out).state.paras;
    expect(back[0].runs[0].props.lang).toBe("es-ES");
    expect(back[1].runs[0].props.noProof).toBe(true);
  });
});
