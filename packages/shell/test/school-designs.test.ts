// Our own designs' theme data (D-190 to D-192): valid colours, readable text in every variant, four variants, unique names.

import { describe, expect, it } from "vitest";
import { COLOR_KEYS, DESIGN_GROUPS, MIN_CONTRAST, SCHOOL_DESIGNS, contrast, designVariants, schemeHex, surfaces } from "../src/school-designs";

describe("school designs", () => {
  it("has about fifty designs with unique names, in the four groups", () => {
    expect(SCHOOL_DESIGNS.length).toBeGreaterThanOrEqual(40);
    expect(new Set(SCHOOL_DESIGNS.map((d) => d.name)).size).toBe(SCHOOL_DESIGNS.length);
    for (const d of SCHOOL_DESIGNS) expect(DESIGN_GROUPS).toContain(d.group);
    for (const g of DESIGN_GROUPS) expect(SCHOOL_DESIGNS.some((d) => d.group === g)).toBe(true);
  });

  it("gives every design four variants with valid colours", () => {
    for (const d of SCHOOL_DESIGNS) {
      const vs = designVariants(d);
      expect(vs.length, d.name).toBe(4);
      expect(new Set(vs.map((v) => v.colorName)).size).toBe(4);
      for (const v of vs) for (const k of COLOR_KEYS) expect(v.colors[k], `${d.name} ${v.colorName} ${k}`).toMatch(/^[0-9A-F]{6}$/);
      // The variants differ from each other.
      expect(new Set(vs.map((v) => JSON.stringify(v.colors))).size, d.name).toBe(4);
    }
  });

  it("keeps text readable on the background in every variant (WCAG)", () => {
    const fails: string[] = [];
    for (const d of SCHOOL_DESIGNS) for (const v of designVariants(d)) {
      const s = surfaces(v.colors, v.dark);
      const check = (what: string, fg: string, bg: string, min: number) => { const r = contrast(fg, bg); if (r < min) fails.push(`${v.colorName}: ${what} ${r.toFixed(2)} < ${min}`); };
      check("body", s.text, s.bg, MIN_CONTRAST.body);
      check("text 2 on bg", s.text2, s.bg, MIN_CONTRAST.text);
      check("text on bg 2", s.text, s.bg2, MIN_CONTRAST.text);
      check("title", schemeHex(v.colors, v.dark, d.title ?? "tx1"), s.bg, MIN_CONTRAST.text);
      check("link", v.colors.hlink, s.bg, MIN_CONTRAST.text);
      // In RaabeCalc and RaabeWrite the colours are used light: dark 1 text on light 1.
      check("light use", v.colors.dk1, v.colors.lt1, MIN_CONTRAST.body);
    }
    expect(fails).toEqual([]);
  });
});
