// Sprache automatisch erkennen: sentences of school texts in the four proofing languages.

import { describe, expect, it } from "vitest";
import { detectLanguage, detectSentences } from "../src/features/detect";

describe("Sprache automatisch erkennen", () => {
  it("recognises German, English, French and Spanish sentences", () => {
    expect(detectLanguage("Die Fotosynthese ist ein Vorgang, bei dem Pflanzen Licht nutzen.")).toBe("de");
    expect(detectLanguage("We went to the market with our friends to buy some fresh vegetables.")).toBe("en");
    expect(detectLanguage("Nous sommes allés au marché avec nos amis pour acheter des légumes frais.")).toBe("fr");
    expect(detectLanguage("Mañana vamos a la playa con nuestros amigos porque hace mucho calor.")).toBe("es");
    expect(detectLanguage("J'aime l'école, mais je n'aime pas les devoirs.")).toBe("fr");
    expect(detectLanguage("¿Dónde está la biblioteca de la escuela?")).toBe("es");
  });
  it("leaves short or unclear text alone", () => {
    expect(detectLanguage("Hallo Welt")).toBeNull();
    expect(detectLanguage("Paris, Berlin, Madrid, London")).toBeNull();
    expect(detectLanguage("Kapitel 3: Ergebnisse")).toBeNull();
  });
  it("splits a paragraph into sentences", () => {
    const s = detectSentences("Das ist ein deutscher Satz mit vielen Wörtern. This is an English sentence with some words. Fin");
    expect(s.map((x) => x.lang)).toEqual(["de", "en", null]);
    expect(s[1].a).toBe(47);
  });
});
