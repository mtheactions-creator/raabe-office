import { describe, expect, it } from "vitest";
import { wordsOf } from "../src/spelling";

describe("Rechtschreibung: which words are checked", () => {
  it("leaves out capitals, words with digits and addresses (Office's defaults)", () => {
    expect(wordsOf("Die Klasse 7b schreibt eine Arbeit").map((w) => w.word)).toEqual(["Die", "Klasse", "schreibt", "eine", "Arbeit"]);
    expect(wordsOf("NRW und GmbH").map((w) => w.word)).toEqual(["und", "GmbH"]);
    expect(wordsOf("siehe www.raabeschule.de oder max@schule.de bitte").map((w) => w.word)).toEqual(["siehe", "oder", "bitte"]);
    expect(wordsOf("Anna-Lena's Heft").map((w) => w.word)).toEqual(["Anna-Lena's", "Heft"]);
  });
});
