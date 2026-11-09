// styles.xml of a new document declares the w14 namespace its text effects use (by-hand audit, D-134): a new document
// that took ligatures over into its defaults was called damaged by Word.

import { describe, expect, it } from "vitest";
import { strFromU8, unzipSync } from "fflate";
import { Doc, StyleSheet, makePara } from "@raabe/text-engine";
import { writeDocx } from "../src";

describe("styles.xml namespaces", () => {
  it("declares w14 when the defaults carry ligatures", () => {
    const d = new Doc({ paras: [makePara("Text")] });
    const s = d.styles;
    d.setStyles(new StyleSheet(s.styles, { ...s.defaultsR, textFx: '<w14:ligatures w14:val="standardContextual"/>' }, s.defaultsP, s.theme));
    const xml = strFromU8(unzipSync(writeDocx(d.state(), null))["word/styles.xml"]);
    expect(xml).toContain("<w14:ligatures");
    expect(xml).toMatch(/<w:styles [^>]*xmlns:w14="http:\/\/schemas\.microsoft\.com\/office\/word\/2010\/wordml"/);
  });
});
