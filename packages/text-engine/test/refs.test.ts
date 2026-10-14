// Textmarken, captions and the fields of cross references (refs.ts).

import { describe, expect, it } from "vitest";
import { Doc, OBJ, makePara, type InlineObject, type Para } from "../src";
import * as Refs from "../src/refs";

const field = (instr: string, display = ""): InlineObject => ({ kind: "field", data: undefined, display, instr });
const withField = (before: string, instr: string, after = ""): Para => ({ ...makePara(`${before}${OBJ}${after}`), objects: [field(instr)] });

describe("bookmarks", () => {
  it("are added, listed, moved and deleted", () => {
    const d = new Doc({ paras: [makePara("Einleitung"), makePara("Der Text.")] });
    Refs.addBookmark(d, "Start", { p: 1, o: 4 }, { p: 1, o: 8 });
    expect(Refs.bookmarks(d)).toEqual([{ name: "Start", start: { p: 1, o: 4 }, end: { p: 1, o: 8 } }]);
    Refs.addBookmark(d, "Start", { p: 0, o: 0 }, { p: 0, o: 10 });
    expect(Refs.bookmarks(d).map((b) => b.start.p)).toEqual([0]);
    Refs.deleteBookmark(d, "Start");
    expect(Refs.bookmarks(d)).toEqual([]);
    expect(Refs.validBookmarkName("Kapitel_2")).toBe(true);
    expect(Refs.validBookmarkName("2Kapitel")).toBe(false);
  });
});

describe("fields", () => {
  it("number captions per label and fill cross references", () => {
    const d = new Doc({ paras: [
      withField("Abbildung ", "SEQ Abbildung \\* ARABIC", ": Blatt"),
      makePara("Siehe"),
      withField("Tabelle ", "SEQ Tabelle \\* ARABIC", ": Werte"),
      withField("Abbildung ", "SEQ Abbildung \\* ARABIC", ": Wurzel"),
    ] });
    const name = Refs.refBookmark(d, { p: 3, o: 0 }, { p: 3, o: d.main[3].text.length });
    d.setStory("main", d.main.map((p, i) => (i === 1 ? { ...withField("Siehe ", `REF ${name} \\h`), } : p)));
    const page = Refs.updateFields(d, () => "1", () => null);
    expect(page).toBe(true);
    const shown = d.main.map((p) => p.objects?.[0]?.display);
    expect(shown).toEqual(["1", "Abbildung 2: Wurzel", "1", "2"]);
    expect(Refs.captions(d, "Abbildung").map((c) => c.text)).toEqual(["Abbildung 1: Blatt", "Abbildung 2: Wurzel"]);
  });
});
