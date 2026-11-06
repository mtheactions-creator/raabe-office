// Text editing on a:txBody and slide operations (D-070, D-071).

import { describe, expect, it } from "vitest";
import { Pptx, PptxDom, slideView } from "@raabe/ooxml";
import * as T from "../src/textops";
import { Deck } from "../src/deck";

const NS = `xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"`;
const body = (inner: string) => PptxDom.parseXml(`<p:txBody ${NS}><a:bodyPr/><a:lstStyle/>${inner}</p:txBody>`).documentElement;
const run = (t: string, attrs = "") => `<a:r><a:rPr lang="de-DE"${attrs}/><a:t>${t}</a:t></a:r>`;

describe("text editing", () => {
  it("types into the run before the caret, keeping its formatting", () => {
    const b = body(`<a:p>${run("fett", ' b="1"')}${run(" normal")}<a:endParaRPr lang="de-DE"/></a:p>`);
    const pos = T.insertText(b, { p: 0, o: 4 }, "er");
    expect(pos).toEqual({ p: 0, o: 6 });
    expect(T.paraText(T.paras(b)[0])).toBe("fetter normal");
    expect(T.formatAt(T.paras(b)[0], 6)?.getAttribute("b")).toBe("1");
  });

  it("splits and joins paragraphs (Eingabe, Rücktaste at the start)", () => {
    const b = body(`<a:p><a:pPr lvl="1"/>${run("Hallo Welt")}</a:p>`);
    const pos = T.splitPara(b, { p: 0, o: 5 });
    expect(pos).toEqual({ p: 1, o: 0 });
    expect(T.paras(b).map(T.paraText)).toEqual(["Hallo", " Welt"]);
    expect(T.paras(b)[1].getElementsByTagNameNS(PptxDom.NS_A, "pPr")[0]?.getAttribute("lvl")).toBe("1");
    T.deleteRange(b, { p: 0, o: 5 }, { p: 1, o: 0 });
    expect(T.paras(b).map(T.paraText)).toEqual(["Hallo Welt"]);
  });

  it("formats a range by splitting runs", () => {
    const b = body(`<a:p>${run("normal fett")}</a:p>`);
    T.formatRange(b, { p: 0, o: 7 }, { p: 0, o: 11 }, (r) => r.setAttribute("b", "1"));
    const runs = Array.from(T.paras(b)[0].getElementsByTagNameNS(PptxDom.NS_A, "r"));
    expect(runs.map((r) => [r.textContent, r.firstElementChild!.getAttribute("b")])).toEqual([["normal ", null], ["fett", "1"]]);
  });

  it("line breaks count one character", () => {
    const b = body(`<a:p>${run("eins")}<a:br><a:rPr lang="de-DE"/></a:br>${run("zwei")}</a:p>`);
    expect(T.paraText(T.paras(b)[0])).toBe("eins\nzwei");
    T.insertText(b, { p: 0, o: 5 }, "X");
    expect(T.paraText(T.paras(b)[0])).toBe("eins\nXzwei");
  });

  it("finds words and jumps like Strg+Pfeil", () => {
    const b = body(`<a:p>${run("Ein kleiner Test")}</a:p>`);
    expect(T.wordAt(T.paras(b)[0], 6)).toEqual({ from: 4, to: 11 });
    expect(T.wordJump(b, { p: 0, o: 0 }, 1)).toEqual({ p: 0, o: 4 });
    expect(T.wordJump(b, { p: 0, o: 16 }, -1)).toEqual({ p: 0, o: 12 });
  });
});

describe("slides", () => {
  it("adds, duplicates, moves and deletes slides with undo", () => {
    const d = new Deck();
    d.addSlide(0);
    expect(d.slides.length).toBe(2);
    // After a Titelfolie the new slide is "Titel und Inhalt" (PowerPoint).
    expect(d.layoutName(d.slides[1].path)).toBe("Titel und Inhalt");
    d.duplicateSlides([1]);
    expect(d.slides.length).toBe(3);
    d.moveSlides([2], 0);
    d.deleteSlides([0]);
    expect(d.slides.length).toBe(2);
    d.undo(); d.undo(); d.undo(); d.undo();
    expect(d.slides.length).toBe(1);
    d.redo();
    expect(d.slides.length).toBe(2);
    const back = Pptx.read(d.write());
    expect(back.slides.length).toBe(2);
  });

  it("writes notes and reads them back", () => {
    const d = new Deck();
    d.setNotes(d.slides[0].path, "Erste Zeile\nZweite Zeile");
    const back = new Deck(Pptx.read(d.write()));
    expect(back.notesText(back.slides[0].path)).toBe("Erste Zeile\nZweite Zeile");
  });

  it("changes a slide's layout and keeps the title text", () => {
    const d = new Deck();
    const path = d.slides[0].path;
    d.change("Text", [path], () => {
      const title = Array.from(d.pptx.doc(path).getElementsByTagNameNS(PptxDom.NS_P, "txBody"))[0];
      T.insertText(title, { p: 0, o: 0 }, "Titel");
    });
    d.setLayout([0], d.layouts()[1].path);
    const v = slideView(d.pptx, path);
    expect(d.layoutName(path)).toBe("Titel und Inhalt");
    expect(v.shapes.find((s) => s.ph?.type === "title")?.text?.paras[0].runs[0].text).toBe("Titel");
    expect(v.shapes.some((s) => s.ph?.idx === 1)).toBe(true);
  });

  it("keeps sections in presentation.xml", () => {
    const d = new Deck();
    d.addSlide(0);
    d.setSections([{ name: "A", id: "{1}", slides: [d.slides[0].id] }, { name: "B", id: "{2}", slides: [d.slides[1].id] }]);
    const back = new Deck(Pptx.read(d.write()));
    expect(back.sections().map((s) => s.name)).toEqual(["A", "B"]);
  });
});

describe("Raabeschule templates", () => {
  it("builds the Referat with its slides and texts in the Raabeschule design", async () => {
    const { PRESENT_TEMPLATES, templateDeck } = await import("../src/templates");
    const d = templateDeck(PRESENT_TEMPLATES[0]);
    expect(d.slides.length).toBe(6);
    const v = slideView(d.pptx, d.slides[1].path);
    expect(v.theme.name).toBe("Raabeschule");
    expect(v.shapes.map((s) => s.text?.paras.map((p) => p.runs.map((r) => r.text).join("")))).toEqual([["Gliederung"], ["Einleitung", "Hauptteil", "Ergebnis", "Quellen"]]);
    const back = Pptx.read(d.write());
    expect(back.slides.length).toBe(6);
  });
});
