// Step 4b E: Sortieren, Formatvorlagen written back into Word's files, theme and document
// defaults, the Auswahlbereich's names and Ausblenden (wp:docPr).

import { describe, expect, it } from "vitest";
import { readFileSync } from "node:fs";
import { join } from "node:path";
import { strFromU8, unzipSync } from "fflate";
import { OFFICE_2013_COLORS, Revisions, sortDate, sortNumber, sortOrder, themeColor, type StyleDef } from "@raabe/text-engine";
import { readDocx, writeDocx } from "../src";

const fixture = (name: string) => new Uint8Array(readFileSync(join(__dirname, "fixtures", name)));
const part = (bytes: Uint8Array, path: string) => strFromU8(unzipSync(bytes)[path]);

describe("Sortieren", () => {
  const spec = { keys: [{ field: 0, type: "text" as const, desc: false }], header: false, separator: "\t", caseSensitive: false };
  it("sorts text the German way, empty first", () => {
    const list = ["Zebra", "apfel", "Birne", "", "Ähre"].map((t) => [t]);
    expect(sortOrder(list, spec).map((i) => list[i][0])).toEqual(["", "Ähre", "apfel", "Birne", "Zebra"]);
    expect(sortOrder(list, { ...spec, keys: [{ field: 0, type: "text", desc: true }] }).map((i) => list[i][0])).toEqual(["Zebra", "Birne", "apfel", "Ähre", ""]);
  });
  it("reads numbers with a decimal comma and dates as Word does", () => {
    expect(sortNumber("Preis: 1.234,50 €")).toBe(1234.5);
    expect(sortNumber("-3")).toBe(-3);
    expect(Number.isNaN(sortNumber("keine Zahl"))).toBe(true);
    expect(sortDate("14.10.2026")).toBe(20261014);
    expect(sortDate("3. März 2025")).toBe(20250303);
    expect(sortDate("2024-01-05")).toBe(20240105);
  });
  it("sorts by a field, keeps a heading row and is stable", () => {
    const list = [["Name\tPunkte", "Name", "Punkte"], ["Anna\t12", "Anna", "12"], ["Ben\t9", "Ben", "9"], ["Cem\t12", "Cem", "12"]];
    const order = sortOrder(list, { ...spec, header: true, keys: [{ field: 2, type: "number", desc: true }] });
    expect(order).toEqual([0, 1, 3, 2]);
  });
});

describe("Formatvorlagen in Word's files", () => {
  it("writes a changed style from the model and keeps the others' XML", () => {
    const r = readDocx(fixture("word-formatting.docx"));
    const sheet = r.state.styles;
    const h1 = [...sheet.styles.values()].find((s) => s.name.toLowerCase() === "heading 1") ?? [...sheet.styles.values()].find((s) => s.type === "paragraph" && !s.isDefault)!;
    const changed: StyleDef = { ...h1, rPr: { ...h1.rPr, size: 30, bold: true }, pPr: { ...h1.pPr, align: "center" }, data: undefined };
    const out = writeDocx({ ...r.state, styles: sheet.withStyle(changed) }, r.pkg);
    const xml = part(out, "word/styles.xml");
    const before = part(fixture("word-formatting.docx"), "word/styles.xml");
    // Every other style element is byte for byte the original.
    const others = (x: string) => (x.match(/<w:style\b[\s\S]*?<\/w:style>/g) ?? []).filter((e) => !e.includes(`w:styleId="${h1.id}"`));
    expect(others(xml)).toEqual(others(before));
    const back = readDocx(out).state.styles.get(h1.id)!;
    expect(back.rPr.size).toBe(30);
    expect(back.rPr.bold).toBe(true);
    expect(back.pPr.align).toBe("center");
  });

  it("removes a deleted style and writes changed document defaults", () => {
    const r = readDocx(fixture("word-formatting.docx"));
    const sheet = r.state.styles;
    const victim = [...sheet.styles.values()].find((s) => s.type === "character" && !s.isDefault)!;
    const next = sheet.without(victim.id).withDefaults({ ...sheet.defaultsR, size: 12 }, { ...sheet.defaultsP, after: 6 });
    const out = writeDocx({ ...r.state, styles: next }, r.pkg);
    const back = readDocx(out).state.styles;
    expect(back.get(victim.id)).toBeUndefined();
    expect(back.defaultsR.size).toBe(12);
    expect(back.defaultsP.after).toBe(6);
  });

  it("writes theme colours and fonts into the file's theme part", () => {
    const r = readDocx(fixture("word-formatting.docx"));
    const t = r.state.styles.theme;
    const theme = { ...t, majorFont: "Georgia", minorFont: "Verdana", colors: { ...t.colors, accent1: "FF0000" }, colorName: "Rot", fontName: "Georgia-Verdana" };
    const out = writeDocx({ ...r.state, styles: r.state.styles.withTheme(theme) }, r.pkg);
    const xml = part(out, "word/theme/theme1.xml");
    expect(xml).toContain('<a:clrScheme name="Rot">');
    expect(xml).toContain('<a:accent1><a:srgbClr val="FF0000"/></a:accent1>');
    expect(xml).toMatch(/<a:majorFont>\s*<a:latin typeface="Georgia"/);
    expect(xml).toMatch(/<a:minorFont>\s*<a:latin typeface="Verdana"/);
    // The format matrix (effects, fills) stays Word's.
    expect(xml).toContain("<a:fmtScheme");
    const back = readDocx(out).state.styles.theme;
    expect([back.majorFont, back.minorFont, back.colors?.accent1, back.colorName]).toEqual(["Georgia", "Verdana", "FF0000", "Rot"]);
  });
});

describe("Auswahlbereich", () => {
  it("keeps names and writes Ausblenden as wp:docPr hidden", () => {
    const r = readDocx(fixture("word-shapes.docx"));
    const at = r.state.paras.findIndex((p) => p.floats?.length);
    expect(at).toBeGreaterThanOrEqual(0);
    const f = r.state.paras[at].floats![0];
    expect(f.alt?.name).toBeTruthy();
    const paras = [...r.state.paras];
    paras[at] = { ...paras[at], floats: [{ ...f, alt: { ...f.alt, name: "Mein Objekt", hidden: true } }, ...paras[at].floats!.slice(1)] };
    const out = writeDocx({ ...r.state, paras }, r.pkg);
    expect(part(out, "word/document.xml")).toMatch(/<wp:docPr [^>]*name="Mein Objekt"[^>]*hidden="1"/);
    const back = readDocx(out).state.paras[at].floats![0];
    expect(back.alt?.name).toBe("Mein Objekt");
    expect(back.alt?.hidden).toBe(true);
  });
});

describe("Überarbeitungsbereich", () => {
  it("lists each change once, also over paragraph marks", () => {
    const rev = { kind: "ins" as const, author: "A", date: "", id: 1 };
    const del = { kind: "del" as const, author: "B", date: "", id: 2 };
    const p = (text: string, runs: { len: number; props: object }[], mark: object = {}) => ({ text, runs, props: {}, mark }) as never;
    const list = Revisions.listRevisions([
      p("Hallo Welt", [{ len: 6, props: {} }, { len: 4, props: { rev } }], { rev }),
      p("neu", [{ len: 3, props: { rev } }]),
      p("alt", [{ len: 3, props: { rev: del } }]),
    ]);
    expect(list.map((x) => [x.rev.kind, x.text])).toEqual([["ins", "Welt\nneu"], ["del", "alt"]]);
  });
});

describe("Zeilennummern", () => {
  it("reads and writes w:lnNumType as Word does (start 0-based, distance in twips)", () => {
    const r = readDocx(fixture("word-short.docx"));
    const sect = r.state.settings.section;
    const paras = [...r.state.paras];
    paras[0] = { ...paras[0], props: { ...paras[0].props, suppressLineNumbers: true } };
    const state = { ...r.state, paras, settings: { ...r.state.settings, section: { ...sect, lineNumbers: { countBy: 2, start: 5, distance: 36, restart: "continuous" as const } } } };
    const out = writeDocx(state, r.pkg);
    const xml = part(out, "word/document.xml");
    expect(xml).toContain('<w:lnNumType w:countBy="2" w:start="4" w:distance="720" w:restart="continuous"/>');
    expect(xml).toContain("<w:suppressLineNumbers/>");
    const back = readDocx(out).state;
    expect(back.settings.section.lineNumbers).toEqual({ countBy: 2, start: 5, distance: 36, restart: "continuous" });
    expect(back.paras[0].props.suppressLineNumbers).toBe(true);
  });
});

describe("Designfarben", () => {
  it("resolve shades and tints as Word does", () => {
    expect(themeColor('w:themeColor="accent1" w:themeShade="BF"', OFFICE_2013_COLORS)).toBe("2F5496");
    // Word: 1F3763; within one step per channel (Word rounds in its own 8-bit HSL).
    const dark = themeColor('w:themeColor="accent1" w:themeShade="7F"', OFFICE_2013_COLORS)!;
    for (const [i, want] of [[0, 0x1f], [2, 0x37], [4, 0x63]]) expect(Math.abs(parseInt(dark.slice(i, i + 2), 16) - want)).toBeLessThanOrEqual(1);
    expect(themeColor('w:themeColor="text1" w:themeTint="A5"', OFFICE_2013_COLORS)).toBe("5A5A5A");
    expect(themeColor('w:themeColor="text1" w:themeTint="BF"', OFFICE_2013_COLORS)).toBe("404040");
    expect(themeColor('w:themeColor="accent1"', { ...OFFICE_2013_COLORS, accent1: "FF0000" })).toBe("FF0000");
  });
});

describe("Seitenhintergrund", () => {
  it("writes Word's watermark into every header and reads it back", () => {
    const r = readDocx(fixture("word-short.docx"));
    const wm = { kind: "text" as const, text: "ENTWURF", font: "Calibri", color: "C0C0C0", semi: true, diagonal: true, width: 412.41, height: 247.44 };
    // A document without headers gets one (the editor adds it; here by hand).
    const stories = new Map(r.state.stories ?? []);
    stories.set("hdr:9", [{ text: "", runs: [], props: {}, mark: {} } as never]);
    const sect = { ...r.state.settings.section, headers: { default: "hdr:9" } };
    const out = writeDocx({ ...r.state, stories, settings: { ...r.state.settings, section: sect, watermark: wm } }, r.pkg);
    const files = unzipSync(out);
    const hdr = Object.keys(files).filter((f) => /word\/header\d+\.xml$/.test(f)).map((f) => strFromU8(files[f]));
    expect(hdr.some((x) => x.includes('<w:docPartGallery w:val="Watermarks"/>') && x.includes('string="ENTWURF"') && x.includes("rotation:315") && x.includes('xmlns:v="urn:schemas-microsoft-com:vml"'))).toBe(true);
    const back = readDocx(out).state;
    expect(back.settings.watermark).toEqual(wm);
    // The header's own text is not the watermark.
    const h = back.stories?.get(back.settings.section.headers.default!);
    expect(h?.every((p) => !p.objects?.length)).toBe(true);
  });

  it("writes w:background with its theme colour and w:pgBorders", () => {
    const r = readDocx(fixture("word-short.docx"));
    const line = { style: "single", size: 1.5, color: "FF0000", space: 24 };
    const sect = { ...r.state.settings.section, pageBorders: { top: line, left: line, bottom: line, right: line, offsetFrom: "page" as const } };
    const out = writeDocx({ ...r.state, settings: { ...r.state.settings, section: sect, background: { color: "FAE2D5", colorTheme: 'w:themeColor="accent2" w:themeTint="33"' } } }, r.pkg);
    const xml = part(out, "word/document.xml");
    expect(xml).toContain('<w:background w:color="FAE2D5" w:themeColor="accent2" w:themeTint="33"/><w:body>');
    expect(xml).toMatch(/<w:pgBorders w:offsetFrom="page"><w:top w:val="single" w:sz="12" w:space="24" w:color="FF0000"\/>/);
    const back = readDocx(out).state.settings;
    expect(back.background).toEqual({ color: "FAE2D5", colorTheme: 'w:themeColor="accent2" w:themeTint="33"' });
    expect(back.section.pageBorders?.offsetFrom).toBe("page");
    expect(back.section.pageBorders?.left?.size).toBe(1.5);
  });
});
