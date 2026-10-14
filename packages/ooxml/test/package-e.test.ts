// Step 4b E: Sortieren, Formatvorlagen written back into Word's files, theme and document
// defaults, the Auswahlbereich's names and Ausblenden (wp:docPr).

import { describe, expect, it } from "vitest";
import { readFileSync } from "node:fs";
import { join } from "node:path";
import { strFromU8, unzipSync } from "fflate";
import { Revisions, sortDate, sortNumber, sortOrder, type StyleDef } from "@raabe/text-engine";
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
