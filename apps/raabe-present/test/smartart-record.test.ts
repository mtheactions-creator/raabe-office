// SmartArt creation (D-077) and recording (D-078): the parts we write and the engine's arrangement.
import { describe, expect, it } from "vitest";
import { dataModelXml, drawingXml, colorDef, colorKey, schemeOfColors } from "../src/smartart-parts";
import { layout, tree, fitSize } from "../src/smartart-engine";
import { wav } from "../src/record";

const items = [{ text: "Planen", level: 0 }, { text: "Material", level: 1 }, { text: "Bauen", level: 0 }, { text: "Prüfen & Ändern", level: 0 }];

describe("SmartArt data model", () => {
  const { xml, ids } = dataModelXml(items, { layoutId: "urn:microsoft.com/office/officeart/2005/8/layout/default", layoutCat: "list", drawingRel: "rId9" });
  const doc = new DOMParser().parseFromString(xml, "application/xml");
  const NS = "http://schemas.openxmlformats.org/drawingml/2006/diagram";

  it("has a document point, one point per item and a parOf connection for each", () => {
    const pts = Array.from(doc.getElementsByTagNameNS(NS, "pt"));
    expect(pts.filter((p) => p.getAttribute("type") === "doc").length).toBe(1);
    expect(pts.filter((p) => !p.getAttribute("type")).length).toBe(4);
    expect(doc.getElementsByTagNameNS(NS, "cxn").length).toBe(4);
    expect(ids.length).toBe(4);
  });

  it("connects a sub-item to its item and orders siblings", () => {
    const cx = Array.from(doc.getElementsByTagNameNS(NS, "cxn"));
    const sub = cx.find((c) => c.getAttribute("destId") === ids[1])!;
    expect(sub.getAttribute("srcId")).toBe(ids[0]);
    const third = cx.find((c) => c.getAttribute("destId") === ids[3])!;
    expect(third.getAttribute("srcOrd")).toBe("2");
  });

  it("escapes text and names the drawing part", () => {
    expect(xml).toContain("Prüfen &amp; Ändern");
    expect(xml).toContain('relId="rId9"');
  });

  it("keeps the colour set by its PowerPoint id", () => {
    expect(schemeOfColors("urn:microsoft.com/office/officeart/2005/8/colors/colorful1")).toBe("colorful1");
    expect(schemeOfColors(undefined)).toBe("accent1_2");
    // Step 8b's keys still name their sets (D-101).
    expect(colorKey("colorful")).toBe("colorful1");
    expect(colorDef("accent2-outline").name).toBe("Farbige Kontur - Akzent 2");
  });
});

describe("SmartArt engine", () => {
  it("builds the item tree by level", () => {
    const t = tree(items);
    expect(t.map((n) => n.text)).toEqual(["Planen", "Bauen", "Prüfen & Ändern"]);
    expect(t[0].children.map((n) => n.text)).toEqual(["Material"]);
  });

  it("lays a grid out within the frame, one shape per item", () => {
    const W = 8128000, H = 4572000;
    const s = layout({ family: "grid", node: "rect", aspect: 0.6, gap: 0.1 }, items.filter((i) => i.level === 0), W, H);
    expect(s.filter((x) => x.item !== undefined).length).toBe(3);
    for (const x of s) { expect(x.x).toBeGreaterThanOrEqual(-1); expect(x.y).toBeGreaterThanOrEqual(-1); expect(x.x + x.w).toBeLessThanOrEqual(W + 1); expect(x.y + x.h).toBeLessThanOrEqual(H + 1); }
  });

  it("draws the measured arrangement for three items as it is", () => {
    const spec = { family: "row" as const, node: "rect", aspect: 0.4, gap: 0.1, items3: [[0, 0.2, 0.3, 0.3], [0.35, 0.2, 0.3, 0.3], [0.7, 0.2, 0.3, 0.3]] };
    const s = layout(spec, items.filter((i) => i.level === 0), 1000, 1000);
    expect(s.map((x) => [x.x, x.y, x.w, x.h])).toEqual([[0, 200, 300, 300], [350, 200, 300, 300], [700, 200, 300, 300]]);
  });

  it("gives every item one font size that fits", () => {
    const s = layout({ family: "column", node: "rect", aspect: 0.2, gap: 0.1 }, items.filter((i) => i.level === 0), 8128000, 4572000);
    const size = fitSize(s, ["Planen", "Bauen", "Prüfen & Ändern"]);
    expect(size).toBeGreaterThan(5);
    expect(size).toBeLessThanOrEqual(65);
    expect(drawingXml(s, ["{A}", "{B}", "{C}"], ["Planen", "Bauen", "Prüfen & Ändern"])).toContain('modelId="{A}"');
  });
});

describe("Recording", () => {
  it("writes a 16-bit mono WAV", () => {
    const b = wav(Float32Array.from([0, 1, -1, 0.5]), 22050);
    const v = new DataView(b.buffer);
    expect(String.fromCharCode(...b.subarray(0, 4))).toBe("RIFF");
    expect(String.fromCharCode(...b.subarray(8, 12))).toBe("WAVE");
    expect(v.getUint32(24, true)).toBe(22050);
    expect(v.getUint16(34, true)).toBe(16);
    expect(v.getInt16(46, true)).toBe(32767);
    expect(v.getInt16(48, true)).toBe(-32767);
    expect(b.length).toBe(44 + 8);
  });
});
