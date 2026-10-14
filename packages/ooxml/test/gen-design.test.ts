// Generates apps/raabe-write/src/features/design-data.gen.ts: what Word's Entwurf galleries
// apply. tools/feel-lab/probe_word_design.py applied every entry in Word and read the theme
// colours and fonts and the styles (results/<date>-word-design.json); the style sets
// (Formatvorlagensätze) are Word's QuickStyles files of the local Office installation, read
// with our reader and matched to the gallery's German names by the formatting the probe saw.
//
//   GEN_DESIGN=1 npx vitest run packages/ooxml/test/gen-design.test.ts

import { existsSync, readFileSync, readdirSync, writeFileSync } from "node:fs";
import { join } from "node:path";
import { describe, expect, it } from "vitest";
import type { CharProps, ParaProps, StyleSheet } from "@raabe/text-engine";
import { readDocx } from "../src";

const LAB = join(__dirname, "..", "..", "..", "tools", "feel-lab");
const RESULTS = join(LAB, "results");
const OUT = join(__dirname, "..", "..", "..", "apps", "raabe-write", "src", "features", "design-data.gen.ts");
const QS = "C:/Program Files/Microsoft Office/root/Office16/1031/QuickStyles";
const W2003 = "C:/Program Files/Microsoft Office/root/Templates/1031/Office Word 2003 Look.dotx";

/** The styles a style set brings (English names as in the files). */
const STYLE_NAMES = ["normal", "heading 1", "heading 2", "heading 3", "heading 4", "heading 5", "heading 6", "heading 7", "heading 8", "heading 9",
  "title", "subtitle", "quote", "intense quote", "subtle emphasis", "emphasis", "intense emphasis", "strong", "subtle reference", "intense reference",
  "book title", "list paragraph", "no spacing", "caption", "toc heading"];

interface ProbeStyle { size: number; before?: number; after?: number; line?: number; lineRule?: number; bold?: boolean; caps?: boolean; themeColor?: number }
interface ProbeEntry { name: string; theme: { colors: Record<string, string>; major: string; minor: string }; styles?: Record<string, ProbeStyle> }

const clean = <T extends object>(o: T): T => {
  const out = Object.fromEntries(Object.entries(o).filter(([k, v]) => v !== undefined && k !== "extra" && k !== "rev" && k !== "lang" && k !== "rStyle" && k !== "style"));
  return out as T;
};

function lineOf(e: ProbeStyle): { line: number; lineRule: "auto" | "atLeast" | "exact" } {
  switch (e.lineRule) {
    case 0: return { line: 1, lineRule: "auto" };
    case 1: return { line: 1.5, lineRule: "auto" };
    case 2: return { line: 2, lineRule: "auto" };
    case 3: return { line: e.line ?? 12, lineRule: "atLeast" };
    case 4: return { line: e.line ?? 12, lineRule: "exact" };
    default: return { line: Math.round(((e.line ?? 12) / 12) * 100) / 100, lineRule: "auto" };
  }
}

function setOf(sheet: StyleSheet) {
  const styles: Record<string, { type: string; basedOn?: string; next?: string; link?: string; pPr: ParaProps; rPr: CharProps }> = {};
  for (const n of STYLE_NAMES) {
    const id = sheet.idForName(n);
    const s = id ? sheet.get(id) : undefined;
    if (!s) continue;
    const name = (x?: string) => (x ? sheet.get(x)?.name.toLowerCase() : undefined);
    styles[n] = { type: s.type, basedOn: name(s.basedOn), next: name(s.next), pPr: clean({ ...s.pPr }), rPr: clean({ ...s.rPr }) };
  }
  return { defaultsR: clean({ ...sheet.defaultsR }), defaultsP: clean({ ...sheet.defaultsP }), styles };
}

/** What Word showed for a style set in the probe, as numbers to compare with a file. */
function signature(sheet: StyleSheet) {
  const look = (n: string) => {
    const id = sheet.idForName(n);
    const c = sheet.resolveChar(id ?? sheet.defaultParagraphStyle, {});
    const p = sheet.resolvePara({ style: id ?? sheet.defaultParagraphStyle } as ParaProps);
    const own = id ? sheet.get(id)?.rPr : undefined;
    return { size: c.size, before: p.before, after: p.after, caps: c.caps, bold: c.bold, accent: /accent/.test(own?.colorTheme ?? "") };
  };
  return { normal: look("normal"), h1: look("heading 1"), h2: look("heading 2"), title: look("title") };
}

describe.runIf(!!process.env.GEN_DESIGN)("generate design-data.gen.ts", () => {
  it("writes Word's Entwurf galleries", () => {
    const file = readdirSync(RESULTS).filter((f) => f.endsWith("-word-design.json")).sort().pop()!;
    const probe = JSON.parse(readFileSync(join(RESULTS, file), "utf-8")) as Record<string, ProbeEntry[]>;
    // Word's names: non-breaking spaces ("Word 2013") and trailing blanks as plain text.
    for (const list of Object.values(probe)) if (Array.isArray(list)) for (const e of list) e.name = e.name.replace(new RegExp(String.fromCharCode(0xa0), "g"), " ").trim();

    const colors = probe.Farben.map((e) => ({ name: e.name, colors: e.theme.colors }));
    const fonts = probe.Schriftarten.map((e) => {
      const m = /^(.*?), Überschriftenschriftart: (.*?), Textschriftart: (.*)$/.exec(e.name);
      return m ? { name: m[1], major: m[2], minor: m[3] } : { name: e.name, major: e.theme.major, minor: e.theme.minor };
    });
    // Entries the probe could not reach are named after their two fonts ("Arial-Times New Roman").
    const named = (n: string) => n.replace(/TrebuchetMs/g, "Trebuchet MS");
    for (const n of ["Arial-Times New Roman", "Century Gothic-Palatino Linotype", "Times New Roman-Arial", "Consolas-Verdana", "Garamond-TrebuchetMs", "TrebuchetMs"]) {
      if (fonts.some((f) => f.name === n)) continue;
      const [major, minor] = n.includes("-") ? n.split("-").map(named) : [named(n), named(n)];
      const at = probe.Schriftarten.findIndex((e) => e.name === n);
      fonts.splice(at >= 0 ? at : fonts.length, 0, { name: n, major, minor });
    }
    const seen = new Set<string>();
    const themes = probe.Designs.filter((e) => !seen.has(e.name) && seen.add(e.name))
      .map((e) => ({ name: e.name.replace(/-Design$/, "").trim(), colors: e.theme.colors, major: e.theme.major, minor: e.theme.minor }));
    const spacing = probe.Absatzabstand.map((e) => {
      const n = e.styles!.Normal;
      return { name: e.name, before: n.before ?? 0, after: n.after ?? 0, ...lineOf(n) };
    });
    const effects = probe.Effekte.map((e) => e.name);

    // Style sets: the files, matched to the gallery names by what the probe saw.
    const files = existsSync(QS) ? readdirSync(QS).filter((f) => f.endsWith(".dotx")).map((f) => join(QS, f)) : [];
    if (existsSync(W2003)) files.push(W2003);
    const sets = files.map((f) => {
      const sheet = readDocx(new Uint8Array(readFileSync(f))).state.styles;
      return { file: f, sheet, sig: signature(sheet) };
    });
    const entries = probe.Formatvorlagensatz.filter((e) => e.name !== "Dokument");
    const used = new Set<string>();
    const styleSets: { name: string; set: ReturnType<typeof setOf> }[] = [];
    const cost = (e: ProbeEntry, s: (typeof sets)[number]) => {
      const st = e.styles!;
      const d = (a: number | undefined, b: number) => Math.abs((a ?? 0) - b);
      return d(st.Normal.size, s.sig.normal.size) * 3 + d(st.Normal.after, s.sig.normal.after) + d(st["Heading 1"].size, s.sig.h1.size) * 2 +
        d(st["Heading 1"].before, s.sig.h1.before) + d(st["Heading 1"].after, s.sig.h1.after) + d(st.Title.size, s.sig.title.size) * 2 +
        d(st["Heading 2"].size, s.sig.h2.size) + (!!st["Heading 1"].caps !== s.sig.h1.caps ? 4 : 0) + (!!st["Heading 1"].bold !== s.sig.h1.bold ? 4 : 0) +
        // Word's theme colour index: 4 … 9 are the accents (WdThemeColorIndex).
        (((st["Heading 1"].themeColor ?? -1) >= 4 && (st["Heading 1"].themeColor ?? -1) <= 9) !== s.sig.h1.accent ? 6 : 0);
    };
    // Greedy by best match first.
    const pairs = entries.flatMap((e) => sets.map((s) => ({ e, s, c: cost(e, s) }))).sort((a, b) => a.c - b.c);
    const done = new Set<string>();
    for (const p of pairs) {
      if (done.has(p.e.name) || used.has(p.s.file)) continue;
      done.add(p.e.name);
      used.add(p.s.file);
      styleSets.push({ name: p.e.name, set: setOf(p.s.sheet) });
      console.log(p.e.name, "<-", p.s.file.split("/").pop(), "cost", p.c.toFixed(2));
    }
    styleSets.sort((a, b) => entries.findIndex((e) => e.name === a.name) - entries.findIndex((e) => e.name === b.name));

    const src = `// Generated by packages/ooxml/test/gen-design.test.ts from Word's Entwurf galleries
// (tools/feel-lab/probe_word_design.py, ${file}) and Word's style sets. Do not edit.
/* eslint-disable */
import type { CharProps, ParaProps } from "@raabe/text-engine";

export type ThemeColors = Readonly<Record<string, string>>;
export interface StyleSetData {
  readonly defaultsR: CharProps;
  readonly defaultsP: ParaProps;
  readonly styles: Readonly<Record<string, { readonly type: string; readonly basedOn?: string; readonly next?: string; readonly pPr: ParaProps; readonly rPr: CharProps }>>;
}

/** Entwurf > Designs: Word's themes (colours and fonts). */
export const THEMES: readonly { name: string; colors: ThemeColors; major: string; minor: string }[] = ${JSON.stringify(themes, null, 1)};

/** Entwurf > Farben. */
export const THEME_COLORS: readonly { name: string; colors: ThemeColors }[] = ${JSON.stringify(colors, null, 1)};

/** Entwurf > Schriftarten. */
export const THEME_FONTS: readonly { name: string; major: string; minor: string }[] = ${JSON.stringify(fonts, null, 1)};

/** Entwurf > Absatzabstand (the Normal style's spacing Word sets). */
export const PARAGRAPH_SPACING: readonly { name: string; before: number; after: number; line: number; lineRule: "auto" | "atLeast" | "exact" }[] = ${JSON.stringify(spacing, null, 1)};

/** Entwurf > Effekte: Word's names (only Office's own effects are part of RaabeWrite). */
export const EFFECTS: readonly string[] = ${JSON.stringify(effects)};

/** Entwurf > Dokumentformatierung: Word's style sets. */
export const STYLE_SETS: readonly { name: string; set: StyleSetData }[] = ${JSON.stringify(styleSets, null, 1)};
`;
    writeFileSync(OUT, src.replace(/\r\n/g, "\n"), "utf-8");
    expect(themes.length).toBeGreaterThan(20);
    expect(styleSets.length).toBe(entries.length);
  });
});
