// Silbentrennung: hyphenation points from LibreOffice's patterns (German with compounds).

import { describe, expect, it } from "vitest";
import { loadHyphenation } from "../src/hyphen";

const show = (word: string, points: readonly number[]) => {
  let out = "", last = 0;
  for (const p of points) { out += word.slice(last, p) + "-"; last = p; }
  return out + word.slice(last);
};

describe("Silbentrennung", () => {
  it("divides German words and compounds", async () => {
    const de = (await loadHyphenation("de-DE"))!;
    // Syllables (two letters on each side); at a line end Word keeps three after the hyphen.
    const h = (w: string) => show(w, de.hyphenate(w, 2, 2));
    expect(h("Silbentrennung")).toBe("Sil-ben-tren-nung");
    expect(h("Fotosynthese")).toBe("Fo-to-syn-the-se");
    expect(h("Kohlenstoffdioxid")).toBe("Koh-len-stoff-di-oxid");
    expect(h("mithilfe")).toBe("mit-hil-fe");
    expect(h("Wachstube")).toContain("-");
    // Recorded in Word (word-hyphen.docx): Pflanzen-teile, Verwaltungsbe-hörden, Mikroorganis-men.
    expect(h("Pflanzenteile")).toBe("Pflan-zen-tei-le");
    expect(h("Verwaltungsbehörden")).toBe("Ver-wal-tungs-be-hör-den");
    expect(h("Ei")).toBe("Ei");
    expect(show("Inhaltsangabe", de.hyphenate("Inhaltsangabe"))).toBe("In-halts-an-gabe");
    expect(show("seine", de.hyphenate("seine"))).toBe("seine");
  });
  it("divides English, French and Spanish words", async () => {
    const en = (await loadHyphenation("en-US"))!;
    expect(show("hyphenation", en.hyphenate("hyphenation", 2, 2))).toBe("hy-phen-ation");
    const fr = (await loadHyphenation("fr-FR"))!;
    expect(show("développement", fr.hyphenate("développement", 2, 2))).toBe("dé-ve-lop-pe-ment");
    const es = (await loadHyphenation("es-ES"))!;
    expect(show("biblioteca", es.hyphenate("biblioteca"))).toBe("bi-blio-te-ca");
  });
});
