import { describe, expect, it } from "vitest";
import { strFromU8, unzipSync } from "fflate";
import { Doc, builtinStyle, defaultStyleSheet, makePara, validPara } from "@raabe/text-engine";
import * as R from "../../text-engine/src/revisions";
import * as S from "../../text-engine/src/school";
import { readDocx, writeDocx } from "../src";

const sheet = () => {
  let s = defaultStyleSheet();
  for (const n of ["footnote text", "footnote reference", "endnote text", "endnote reference", "annotation text", "annotation reference",
    "toc 1", "toc 2", "TOC Heading", "heading 1", "heading 2"]) {
    const b = builtinStyle(n);
    if (b) s = s.withStyle(b);
  }
  return s;
};
const id = (name: string) => builtinStyle(name)!.id;

describe("foot- and endnotes", () => {
  it("inserts notes, writes them with separators and reads them back", () => {
    const d = new Doc({ paras: [makePara("Satz mit Fußnote."), makePara("Und Endnote.")], styles: sheet() });
    const fn = S.insertNote(d, { p: 0, o: 17 }, "fn", id("footnote reference"), id("footnote text"));
    const en = S.insertNote(d, { p: 1, o: 12 }, "en", id("endnote reference"), id("endnote text"));
    expect(fn.id).toBe("fn:1");
    expect(en.id).toBe("en:1");
    d.setStory(fn.id, [{ ...d.story(fn.id)![0], text: d.story(fn.id)![0].text + "Quelle", runs: [...d.story(fn.id)![0].runs, { len: 6, props: {} }] }]);
    expect(d.main.every(validPara)).toBe(true);
    expect(S.referencedNotes(d)).toEqual(["fn:1", "en:1"]);
    const files = unzipSync(writeDocx(d.state(), null));
    const fx = strFromU8(files["word/footnotes.xml"]);
    expect(fx).toContain('w:type="separator"');
    expect(fx).toContain("<w:footnoteRef/>");
    expect(strFromU8(files["word/endnotes.xml"])).toContain("<w:endnoteRef/>");
    const doc = strFromU8(files["word/document.xml"]);
    expect(doc).toContain('<w:footnoteReference w:id="1"/>');
    expect(doc).toContain('<w:endnoteReference w:id="1"/>');
    const back = readDocx(writeDocx(d.state(), null)).state;
    const note = [...back.stories!.entries()].find(([k]) => k.startsWith("fn:"))!;
    expect(note[1][0].text.endsWith("Quelle")).toBe(true);
  });

  it("leaves out a note whose reference was deleted", () => {
    const d = new Doc({ paras: [makePara("Text")], styles: sheet() });
    S.insertNote(d, { p: 0, o: 4 }, "fn", id("footnote reference"), id("footnote text"));
    d.delete({ p: 0, o: 4 }, { p: 0, o: 5 });
    const files = unzipSync(writeDocx(d.state(), null));
    const fx = files["word/footnotes.xml"] ? strFromU8(files["word/footnotes.xml"]) : "";
    expect(fx).not.toContain('w:id="1"');
  });
});

describe("comments", () => {
  it("adds a comment with a reply, resolves it and round-trips", () => {
    const d = new Doc({ paras: [makePara("Hier ist ein Fehler.")], styles: sheet() });
    const meta = { author: "Lehrkraft", initials: "LK", date: "2026-10-13T10:00:00Z" };
    const c = S.addComment(d, { p: 0, o: 13 }, { p: 0, o: 19 }, meta, id("annotation text"), id("annotation reference"), "Rechtschreibung prüfen");
    const r = S.addReply(d, c, { ...meta, author: "Schülerin", initials: "S" }, "Erledigt", id("annotation text"), id("annotation reference"));
    d.setComment({ ...d.comments.get(c)!, done: true }, c);
    expect(S.commentRanges(d).map((x) => x.id)).toEqual([c, r]);
    const bytes = writeDocx(d.state(), null);
    const files = unzipSync(bytes);
    const cx = strFromU8(files["word/comments.xml"]);
    expect(cx).toContain('w:author="Lehrkraft"');
    expect(cx).toContain("<w:annotationRef/>");
    expect(strFromU8(files["word/commentsExtended.xml"])).toContain("w15:paraIdParent");
    const back = readDocx(bytes).state;
    const metas = [...back.comments!.values()];
    expect(metas.length).toBe(2);
    expect(metas.find((m) => m.author === "Schülerin")?.parent).toBeDefined();
    expect(metas.find((m) => m.author === "Lehrkraft")?.done).toBe(true);
    const bd = new Doc(back);
    const first = S.commentRanges(bd).find((x) => bd.comments.get(x.id)?.author === "Lehrkraft")!;
    expect(S.commentText(bd, first.id)).toBe("Rechtschreibung prüfen");
    expect(first.start).toEqual({ p: 0, o: 13 });
    S.deleteComment(bd, first.id);
    expect(bd.comments.size).toBe(0);
  });
});

describe("table of contents", () => {
  it("builds a TOC field with linked entries and page numbers", () => {
    const d = new Doc({
      paras: [makePara(""), makePara("Einleitung", { style: id("heading 1") }), makePara("Text"), makePara("Teil", { style: id("heading 2") })],
      styles: sheet(),
    });
    const heads = [{ p: 1, level: 1, text: "Einleitung" }, { p: 3, level: 2, text: "Teil" }];
    const paras = S.tocParagraphs(d, heads, (p) => String(p), [id("toc 1"), id("toc 2"), id("toc 2")], 450);
    d.splice(0, 0, paras);
    const f = S.tocFields(d);
    expect(f.length).toBe(1);
    expect(f[0]).toMatchObject({ a: 0, b: 1 });
    const bytes = writeDocx(d.state(), null);
    const x = strFromU8(unzipSync(bytes)["word/document.xml"]);
    expect(x).toMatch(/TOC \\o "1-3" \\h \\z \\u/);
    expect(x).toContain('w:anchor="_Toc');
    expect(x).toMatch(/PAGEREF _Toc\d+ \\h/);
    expect(x).toContain('<w:bookmarkStart w:id="');
    const back = new Doc(readDocx(bytes).state);
    expect(S.tocFields(back).length).toBe(1);
  });
});

describe("tracked changes", () => {
  it("marks deletions, drops own insertions, and accepts or rejects", () => {
    const d = new Doc({ paras: [makePara("Alter Text bleibt.")], styles: sheet() });
    const ins = R.newRevision("ins", "A");
    d.insertText({ p: 0, o: 10 }, { p: 0, o: 10 }, " neu", { rev: ins });
    R.markDeleted(d, { p: 0, o: 0 }, { p: 0, o: 6 }, R.newRevision("del", "A"));
    expect(d.main[0].text).toBe("Alter Text neu bleibt.");
    expect(d.main[0].runs.map((r) => r.props.rev?.kind ?? null)).toEqual(["del", null, "ins", null]);
    // Deleting one's own insertion removes it for real.
    R.markDeleted(d, { p: 0, o: 10 }, { p: 0, o: 14 }, R.newRevision("del", "A"));
    expect(d.main[0].text).toBe("Alter Text bleibt.");
    // Another author's insertion is marked deleted instead.
    d.insertText({ p: 0, o: 10 }, { p: 0, o: 10 }, " B", { rev: R.newRevision("ins", "B") });
    R.markDeleted(d, { p: 0, o: 10 }, { p: 0, o: 12 }, R.newRevision("del", "A"));
    expect(d.main[0].text).toBe("Alter Text B bleibt.");
    const bytes = writeDocx(d.state(), null);
    const x = strFromU8(unzipSync(bytes)["word/document.xml"]);
    expect(x).toContain("<w:del ");
    expect(x).toContain("<w:delText");
    const back = new Doc(readDocx(bytes).state);
    expect(R.hasRevisions(back.main, { p: 0, o: 0 }, back.end())).toBe(true);
    R.resolveRevisions(back, { p: 0, o: 0 }, back.end(), false);
    // A deletion of another author's insertion counts as a plain deletion (D-034), so
    // rejecting it keeps " B"; Word would also reject B's insertion.
    expect(back.main[0].text).toBe("Alter Text B bleibt.");
    expect(R.hasRevisions(back.main, { p: 0, o: 0 }, back.end())).toBe(false);
    R.resolveRevisions(d, { p: 0, o: 0 }, d.end(), true);
    expect(d.main[0].text).toBe("Text bleibt.");
    expect(d.main[0].runs.length).toBe(1);
  });

  it("tracks paragraph marks: a deleted mark joins on accept", () => {
    const d = new Doc({ paras: [makePara("Eins"), makePara("Zwei")], styles: sheet() });
    R.markDeleted(d, { p: 0, o: 4 }, { p: 1, o: 0 }, R.newRevision("del", "A"));
    expect(d.main.length).toBe(2);
    expect(d.main[0].mark.rev?.kind).toBe("del");
    const x = strFromU8(unzipSync(writeDocx(d.state(), null))["word/document.xml"]);
    expect(x).toMatch(/<w:rPr><w:del [^>]*\/><\/w:rPr><\/w:pPr>/);
    R.resolveRevisions(d, { p: 0, o: 0 }, d.end(), true);
    expect(d.main.map((p) => p.text)).toEqual(["EinsZwei"]);
  });
});

describe("word-step4.docx (made by Word)", () => {
  const load = async () => {
    const { readFileSync } = await import("node:fs");
    return readDocx(new Uint8Array(readFileSync("packages/ooxml/test/fixtures/word-step4.docx")));
  };

  it("reads its equations as LaTeX, notes, the comment and the table of contents", async () => {
    const r = await load();
    expect(r.partial).toEqual([]);
    const math = r.state.paras.flatMap((p) => (p.objects ?? []).filter((o) => o.kind === "math").map((o) => o.latex));
    expect(math).toEqual([
      "E=h\\,f",
      "x=\\frac{-b\\pm \\sqrt{b^{2}-4ac}}{2a}",
      "\\sum_{k=1}^{n}{k}=\\frac{n\\left(n+1\\right)}{2}",
      "\\int_{0}^{1}{x^{2}dx}=\\frac{1}{3}",
      "A=\\begin{pmatrix}1&2\\\\3&4\\end{pmatrix}",
    ]);
    const keys = [...(r.state.stories?.keys() ?? [])];
    expect(keys.filter((k) => k.startsWith("fn:")).length).toBe(2);
    expect(keys.filter((k) => k.startsWith("en:")).length).toBe(1);
    expect(r.state.comments?.size).toBe(1);
    const d = new Doc(r.state);
    const toc = S.tocFields(d);
    expect(toc.length).toBe(1);
    // Word ends the field at the start of the first heading: that paragraph is not an entry.
    expect(d.main.slice(toc[0].a, toc[0].b + 1).every((p) => p.props.style?.startsWith("Verzeichnis"))).toBe(true);
    expect(toc[0].tail).toBe(toc[0].b + 1);
  });

  it("a rebuilt table of contents keeps the heading after it", async () => {
    const d = new Doc((await load()).state);
    const f = S.tocFields(d)[0];
    const heading = d.main[f.tail!];
    const entries = [{ p: f.tail!, level: 1, text: heading.text }];
    const paras = S.tocParagraphs(d, entries, () => "1", ["Verzeichnis1", "Verzeichnis2", "Verzeichnis3"], 450, f.id);
    S.stripTocTail(d, f);
    d.splice(f.a, f.b, paras);
    const g = S.tocFields(d);
    expect(g.length).toBe(1);
    expect(g[0].b - g[0].a + 1).toBe(1);
    expect(d.main[g[0].b + 1].text).toBe(heading.text);
    expect(d.main[g[0].b + 1].anchors?.some((a) => a.field)).toBeFalsy();
    const x = strFromU8(unzipSync(writeDocx(d.state(), null))["word/document.xml"]);
    expect((x.match(/w:fldCharType="end"/g) ?? []).length).toBe((x.match(/w:fldCharType="begin"/g) ?? []).length);
  });
});
