import { describe, expect, it } from "vitest";
import { borderLine, instantFix, smartQuote, wordFixes } from "../src/autocorrect";

/** Applies the fixes for a word finished at `end` (where the trigger was typed). */
function fix(text: string, end = text.length - 1, lang = "de-DE", cellStart = false, capsLock = false): string {
  let out = text;
  for (const f of [...wordFixes(text, end, lang, { cellStart, capsLock })].sort((a, b) => b.a - a.a)) {
    out = out.slice(0, f.a) + (f.symbol ? "<sym>" : f.text) + out.slice(f.b);
  }
  return out;
}

describe("AutoKorrektur", () => {
  it("capitalises sentence starts, not after abbreviations or list labels", () => {
    expect(fix("der ")).toBe("Der ");
    expect(fix("Der test. das ")).toBe("Der test. Das ");
    expect(fix("Er kam, z.B. heute ")).toBe("Er kam, z.B. heute ");
    expect(fix("Wir sehen z.B. das ")).toBe("Wir sehen z.B. das ");
    expect(fix("Am 1. mai ")).toBe("Am 1. Mai "); // Word ends the sentence after "1."
    expect(fix("a)")).toBe("a)");
    expect(fix("a) ")).toBe("a) ");
    expect(fix("o ")).toBe("O ");
    expect(fix("Wirklich? ja ")).toBe("Wirklich? Ja ");
    expect(fix("Er sagte: „ja ")).toBe("Er sagte: „ja ");
    expect(fix("Ende\u2026 und ")).toBe("Ende\u2026 und ");
  });

  it("two initial capitals, Caps Lock, weekdays, typos", () => {
    expect(fix("Das ZWei ")).toBe("Das Zwei ");
    expect(fix("Das USA ")).toBe("Das USA ");
    expect(fix("Das hALLO ")).toBe("Das hALLO "); // only while Caps Lock is on
    expect(fix("Das hALLO ", undefined, "de-DE", false, true)).toBe("Das Hallo ");
    expect(fix("Wie geht\u2018s ")).toBe("Wie geht\u2019s ");
    expect(fix("Peters\u2018 ")).toBe("Peters\u2018 ");
    expect(fix("Am montag ")).toBe("Am Montag ");
    expect(fix("Und dei ")).toBe("Und die ");
  });

  it("symbols as soon as they are typed, fractions and dashes", () => {
    const now = (t: string) => { const f = instantFix(t, t.length); return f ? t.slice(0, f.a) + (f.symbol ? "<sym>" : f.text) + t.slice(f.b) : t; };
    expect(now("(c)")).toBe("\u00a9");
    expect(now("Text(C)")).toBe("Text\u00a9");
    expect(now("Ende...")).toBe("Ende\u2026");
    expect(now("A -->")).toBe("A <sym>");
    expect(now("Gut :)")).toBe("Gut \u{1F60A}");
    expect(now("f(x:)")).toBe("f(x:)");
    expect(instantFix("Gut :)", 6)?.font).toBe("Segoe UI Emoji");
    expect(fix("1/2 ")).toBe("\u00bd ");
    expect(fix("Haus - Garten ")).toBe("Haus – Garten ");
    expect(fix("Haus--Garten ")).toBe("Haus—Garten ");
  });

  it("typographic quotes by language", () => {
    const q = (t: string, lang = "de-DE") => { const f = smartQuote(t, t.length, lang)!; return t.slice(0, f.a) + f.text + t.slice(f.b); };
    expect(q('Er sagte "')).toBe("Er sagte „");
    expect(q('Er sagte „Hallo"')).toBe("Er sagte „Hallo“");
    expect(q("Er sagte '")).toBe("Er sagte ‚");
    expect(q("Er sagte ‚ja'")).toBe("Er sagte ‚ja‘");
    expect(q('He said "', "en-US")).toBe("He said “");
    expect(q('He said “yes"', "en-US")).toBe("He said “yes”");
  });

  it("border lines", () => {
    expect(borderLine("---")?.style).toBe("single");
    expect(borderLine("===")?.style).toBe("double");
    expect(borderLine("--")).toBeNull();
  });
});
