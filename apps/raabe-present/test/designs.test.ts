// RaabePresent's own designs (D-190 to D-197): every design and variant makes well-formed parts with unique shape
// ids, its artwork and effect style exist, and text on layouts with their own background stays readable.

import { describe, expect, it } from "vitest";
import { newPptxFiles } from "@raabe/ooxml";
import { MIN_CONTRAST, SCHOOL_DESIGNS, contrast } from "@raabe/shell";
import { strFromU8 } from "fflate";
import { DESIGNS, DESIGN_HEADINGS, variantsOf } from "../src/designs";

const parse = (s: string) => new DOMParser().parseFromString(s, "application/xml");

describe("designs", () => {
  it("lists Office and every school design once, each under a gallery heading", () => {
    expect(DESIGNS[0].name).toBe("Office");
    expect(DESIGNS.length).toBe(SCHOOL_DESIGNS.length + 1);
    expect(DESIGNS.length).toBeGreaterThanOrEqual(40);
    expect(new Set(DESIGNS.map((d) => d.name)).size).toBe(DESIGNS.length);
    for (const d of DESIGNS) expect(DESIGN_HEADINGS).toContain(d.group);
  });

  it("gives each design of ours artwork of its own and a known effect style", () => {
    for (const d of DESIGNS.slice(1)) {
      expect(d.masterDeco ?? d.layouts?.title?.deco, d.name).toBeTruthy();
      if (d.school!.effects !== "Office") expect(d.effects, d.name).toContain("<a:fmtScheme");
      // No pictures of Office's (a texture fill) in an effect style of ours.
      expect(d.effects ?? "").not.toContain("blipFill");
    }
  });

  it("writes well-formed master, layouts and theme with unique shape ids for every variant", () => {
    for (const d of DESIGNS) for (const v of variantsOf(d)) {
      const files = newPptxFiles(v, new Date(0));
      for (const [name, bytes] of Object.entries(files)) {
        if (!name.endsWith(".xml")) continue;
        const doc = parse(strFromU8(bytes));
        expect(doc.getElementsByTagName("parsererror").length, `${v.colorName} ${name}`).toBe(0);
        if (/slide(Master|Layout)/.test(name)) {
          const ids = Array.from(doc.getElementsByTagNameNS("http://schemas.openxmlformats.org/presentationml/2006/main", "cNvPr")).map((e) => e.getAttribute("id"));
          expect(new Set(ids).size, `${v.colorName} ${name}`).toBe(ids.length);
        }
      }
      const theme = strFromU8(files["ppt/theme/theme1.xml"]);
      expect(theme).toContain(`<a:clrScheme name="${v.colorName ?? v.name}">`);
      // The artwork uses scheme colours only, so the variants recolour it.
      for (const part of ["ppt/slideMasters/slideMaster1.xml", "ppt/slideLayouts/slideLayout1.xml", "ppt/slideLayouts/slideLayout3.xml"]) {
        if (d.name !== "Office") expect(strFromU8(files[part]), `${d.name} ${part}`).not.toMatch(/srgbClr/);
        // Only mapped names (tx1, bg2 …) and accents: dk1/lt1/dk2/lt2 are read differently by viewers.
        expect(strFromU8(files[part]), `${d.name} ${part}`).not.toMatch(/schemeClr val="(dk|lt)[12]"/);
      }
    }
  });

  it("keeps text readable on layouts with a background of their own", () => {
    const fails: string[] = [];
    for (const d of DESIGNS) for (const v of variantsOf(d)) for (const l of [v.layouts?.title, v.layouts?.section]) {
      if (!l?.background) continue;
      const names = Array.from(l.background.matchAll(/schemeClr val="(\w+)"/g)).map((m) => m[1]);
      const dark = l.colors ? l.colors === "dark" : !!v.dark;
      const c = v.colors as Record<string, string>;
      const map: Record<string, string> = dark ? { bg1: c.dk1, bg2: c.dk2, tx1: c.lt1, tx2: c.lt2 } : { bg1: c.lt1, bg2: c.lt2, tx1: c.dk1, tx2: c.dk2 };
      const title = map[v.titleColor ?? "tx1"] ?? c[v.titleColor ?? "tx1"];
      for (const n of names) {
        const bg = map[n] ?? c[n];
        if (contrast(map.tx1, bg) < MIN_CONTRAST.text) fails.push(`${v.colorName}: text on ${n} ${contrast(map.tx1, bg).toFixed(2)}`);
        if (contrast(title, bg) < MIN_CONTRAST.text) fails.push(`${v.colorName}: title on ${n} ${contrast(title, bg).toFixed(2)}`);
      }
    }
    expect(fails).toEqual([]);
  });
});
