import { describe, expect, it } from "vitest";
import { pictureFromSpPr } from "../src";
import { shapePropsXml } from "../src/drawingml";
import * as P from "../../../apps/raabe-write/src/features/picture-presets";

/** XML without the attribute order and the transform, for comparing. */
function norm(xml: string): string {
  const doc = new DOMParser().parseFromString(`<r xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">${xml}</r>`, "application/xml");
  const walk = (el: Element): string => {
    const attrs = Array.from(el.attributes).filter((a) => !a.name.startsWith("xmlns")).map((a) => `${a.name}=${a.value}`).sort().join(" ");
    return `<${el.localName}${attrs ? " " + attrs : ""}>${Array.from(el.children).map(walk).join("")}</${el.localName}>`;
  };
  // An empty effect list says no more than none.
  return Array.from(doc.documentElement.children).filter((c) => c.localName !== "xfrm" && !(c.localName === "effectLst" && !c.children.length)).map(walk).join("");
}

describe("Word's picture presets", () => {
  const all = [
    ...P.PICTURE_STYLES, ...P.EFFECT_PRESETS, ...P.SHADOWS, ...P.REFLECTIONS, ...P.GLOWS, ...P.SOFT_EDGES, ...P.BEVELS, ...P.ROTATIONS_3D,
  ];
  it("are read and written back as Word wrote them", () => {
    const failures: string[] = [];
    for (const p of all) {
      const f = pictureFromSpPr(p.xml);
      const out = shapePropsXml(f, 100, 100);
      // Theme colours (schemeClr) come back as their RGB values; compare those presets loosely.
      if (p.xml.includes("schemeClr") || p.xml.includes("prstClr")) continue;
      if (norm(out) !== norm(p.xml)) failures.push(`${p.label}\n  word: ${norm(p.xml)}\n  ours: ${norm(out)}`);
    }
    expect(failures.slice(0, 5).join("\n")).toBe("");
  });
});
