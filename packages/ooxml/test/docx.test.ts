import { readFileSync, mkdirSync, writeFileSync } from "node:fs";
import { describe, expect, it } from "vitest";
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";
import { Doc, makePara, validPara, type DocState, type Para } from "@raabe/text-engine";
import { readDocx, writeDocx } from "../src";
import fixtures from "../../../prototypes/src/shared/fixtures.json";

const fixture = (name: string) => new Uint8Array(readFileSync(`packages/ooxml/test/fixtures/${name}`));

/** Comparable view of paragraphs: text, style, runs with their properties. */
function shape(paras: readonly Para[]) {
  return paras.map((p) => ({
    text: p.text,
    style: p.props.style,
    props: { ...p.props, extra: p.props.extra?.length },
    runs: p.runs.map((r) => ({ len: r.len, ...r.props, extra: r.props.extra?.length })),
    objects: p.objects?.map((o) => o.kind),
    anchors: p.anchors?.length,
    block: p.block?.kind,
  }));
}

const count = (xml: string, tag: string) => (xml.match(new RegExp(`<${tag}[ >/]`, "g")) ?? []).length;

describe("new documents", () => {
  it("round-trip text, styles and character formatting", () => {
    const d = new Doc({ paras: [makePara("Überschrift", { style: "berschrift1" }), makePara("Ein fettes Wort und \tTab", { align: "both", indLeft: 36, indFirst: -18 })] });
    d.mapChars({ p: 1, o: 4 }, { p: 1, o: 10 }, (o) => ({ ...o, bold: true, color: "FF0000", size: 14 }));
    const back = readDocx(writeDocx(d.state(), null)).state;
    expect(shape(back.paras).map((p) => p.text)).toEqual(["Überschrift", "Ein fettes Wort und \tTab"]);
    expect(back.paras[0].props.style).toBe("berschrift1");
    const r = back.paras[1].runs;
    expect(r.map((x) => [x.len, !!x.props.bold, x.props.color, x.props.size])).toEqual([[4, false, undefined, undefined], [6, true, "FF0000", 14], [14, false, undefined, undefined]]);
    expect(back.paras[1].props).toMatchObject({ align: "both", indLeft: 36, indFirst: -18 });
    expect(back.styles.get("berschrift1")?.rPr.size).toBe(16);
  });

  it("writes theme fonts as names (no theme part in new files)", () => {
    const xml = strFromU8(unzipSync(writeDocx(new Doc().state(), null))["word/styles.xml"]);
    expect(xml).toContain('w:ascii="Calibri"');
    expect(xml).not.toContain("asciiTheme");
  });
});

describe("documents saved by Microsoft Word", () => {
  it("reads the plain test document", () => {
    const { state, partial } = readDocx(fixture("word-short.docx"));
    expect(partial).toEqual([]);
    expect(state.paras.map((p) => p.text)).toEqual(fixtures.base.map((p) => p.text));
    expect(state.paras[0].props.style).toBe("berschrift1");
  });

  it("reads formatting", () => {
    const { state } = readDocx(fixture("word-formatting.docx"));
    const text = state.paras.map((p) => p.text);
    const p2 = state.paras[1];
    const bold = p2.runs.find((r) => r.props.bold);
    expect(bold && p2.text.substr(p2.runs.slice(0, p2.runs.indexOf(bold)).reduce((n, r) => n + r.len, 0), bold.len)).toBe("fetten");
    expect(p2.runs.some((r) => r.props.highlight === "yellow")).toBe(true);
    expect(p2.runs.some((r) => r.props.color === "FF0000")).toBe(true);
    expect(state.paras[2].runs.some((r) => r.props.vertAlign === "subscript")).toBe(true);
    expect(state.paras[2].runs.some((r) => r.props.underline === "double")).toBe(true);
    expect(state.paras[3].props.align).toBe("center");
    expect(state.paras[4].props.align).toBe("right");
    expect(state.paras[5].props.align).toBe("both");
    expect(state.paras[6].props).toMatchObject({ indLeft: 56.7, indFirst: -28.35 });
    expect(state.paras[7].props.tabs?.map((t) => t.align)).toEqual(["left", "right"]);
    expect(state.paras[9].props).toMatchObject({ lineRule: "exact", line: 20 });
    expect(text.some((t) => t.includes("\u000b"))).toBe(true);
    expect(text.some((t) => t.includes("\u000c"))).toBe(true);
    expect(state.paras.every(validPara)).toBe(true);
  });

  for (const name of ["word-short.docx", "word-formatting.docx", "word-rich.docx", "word-layout.docx", "word-fonts.docx", "word-breaks.docx", "word-top.docx"]) {
    it(`${name}: writing back keeps every part and every element`, () => {
      const bytes = fixture(name);
      const first = readDocx(bytes);
      const out = writeDocx(first.state, first.pkg);
      const a = unzipSync(bytes), b = unzipSync(out);
      expect(Object.keys(b).sort()).toEqual(Object.keys(a).sort());
      for (const k of Object.keys(a)) if (k !== "word/document.xml" && k !== "docProps/core.xml") expect(strFromU8(b[k])).toBe(strFromU8(a[k]));
      const x1 = strFromU8(a["word/document.xml"]), x2 = strFromU8(b["word/document.xml"]);
      for (const tag of ["w:tbl", "w:tr", "w:tc", "w:drawing", "w:footnoteReference", "w:commentReference", "w:commentRangeStart",
        "w:hyperlink", "w:fldChar", "w:instrText", "m:oMath", "w:ins", "w:del", "w:bookmarkStart", "w:bookmarkEnd", "w:sectPr", "w:numPr", "w:tab", "w:br"]) {
        expect([tag, count(x2, tag)]).toEqual([tag, count(x1, tag)]);
      }
      const second = readDocx(out);
      expect(shape(second.state.paras)).toEqual(shape(first.state.paras));
      if (process.env.ROUNDTRIP_OUT) {
        mkdirSync(process.env.ROUNDTRIP_OUT, { recursive: true });
        writeFileSync(`${process.env.ROUNDTRIP_OUT}/${name}`, out);
      }
    });
  }

  it("word-rich.docx: shows placeholders for what it cannot edit yet", () => {
    const { state, partial } = readDocx(fixture("word-rich.docx"));
    expect(state.paras.some((p) => p.block?.kind === "table")).toBe(true);
    expect(state.paras.some((p) => p.block?.kind === "field" && p.block.label === "Inhaltsverzeichnis")).toBe(true);
    const kinds = state.paras.flatMap((p) => p.objects?.map((o) => o.kind) ?? []);
    for (const k of ["note", "link", "image", "revision"]) expect(kinds).toContain(k);
    expect(partial).toEqual(expect.arrayContaining(["Tabellen", "Fußnoten", "Kommentare", "Bilder und Formen", "Aufzählungen und Nummerierungen", "Kopf- und Fußzeilen", "Formeln"]));
  });

  it("word-rich.docx: typing next to preserved content keeps it", () => {
    const { state, pkg } = readDocx(fixture("word-rich.docx"));
    const d = new Doc(state);
    const i = d.paras.findIndex((p) => p.text.startsWith("Lies den Text"));
    d.insertText({ p: i, o: 0 }, { p: i, o: 0 }, "Bitte: ", {});
    const out = writeDocx(d.state() as DocState, pkg);
    const x = strFromU8(unzipSync(out)["word/document.xml"]);
    expect(x).toContain("Bitte: ");
    expect(count(x, "w:footnoteReference")).toBe(1);
    expect(count(x, "w:commentReference")).toBe(1);
    expect(count(x, "w:hyperlink")).toBe(1);
  });
});

describe("damaged input", () => {
  it("rejects files that are not Word documents", () => {
    expect(() => readDocx(zipSync({ "hello.txt": strToU8("hi") }))).toThrow();
  });
});
