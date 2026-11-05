import { describe, expect, it } from "vitest";
import type { SheetShape } from "@raabe/sheet-engine";
import { custGeomXml } from "../src/shapes";
import { custGeomFromXml } from "../src/xlsx/custgeom";
import { shapeFromXml, shapeToXml } from "../src/xlsx/shapes";

describe("shapes in a drawing (D-066)", () => {
  it("reads custom geometry back as written", () => {
    const def = "rect l t r b\npath w=95250 h=47625 fill=none: M 0 0 C 10 20 30 40 50 60 L 95250 47625";
    const back = custGeomFromXml(`<xdr:spPr>${custGeomXml(def)}</xdr:spPr>`);
    expect(back).toBe(def);
  });

  it("keeps a freeform, an adjusted preset and the style reference through a round trip", () => {
    const base: SheetShape = { from: { col: 1, row: 1, x: 0, y: 0 }, to: { col: 3, row: 4, x: 0, y: 0 }, geom: "rect" };
    const shapes: SheetShape[] = [
      { ...base, custom: "rect l t r b\npath w=100 h=100: M 0 0 L 100 0 L 50 100 Z", name: "Freihandform: Form 1" },
      { ...base, geom: "roundRect", adjust: { adj: 30000 }, name: "Rechteck: abgerundete Ecken 2" },
      { ...base, geom: "line", connector: true, style: { ln: 2 }, name: "Gerader Verbinder 3" },
    ];
    for (const s of shapes) {
      const back = shapeFromXml(shapeToXml(s, 2, { w: 128, h: 60 }))!;
      const { raw: _r, rawKey: _k, ...model } = back;
      void _r; void _k;
      expect(model.custom).toBe(s.custom);
      expect(model.adjust).toEqual(s.adjust);
      expect(model.geom).toBe(s.geom);
      if (s.connector) expect(model.style).toEqual({ ln: 2 });
      else expect(model.style).toEqual({ ln: 2, shade: 0.15 });
    }
  });

  it("starts a new shape's text at the top left, as Excel 365", () => {
    const xml = shapeToXml({ from: { col: 0, row: 0, x: 0, y: 0 }, to: { col: 2, row: 2, x: 0, y: 0 }, geom: "rect", text: "Hallo" }, 2, { w: 128, h: 40 });
    expect(xml).toContain('anchor="t"');
    expect(xml).toContain('algn="l"');
    expect(xml).toContain('<a:shade val="15000"/>');
  });

  it("keeps a group of shapes with their places in the group", () => {
    const kid = (geom: string, x: number): SheetShape => ({ from: { col: 0, row: 0, x: 0, y: 0 }, to: { col: 0, row: 0, x: 0, y: 0 }, geom, rel: { x, y: 10, w: 50, h: 40 } });
    const g: SheetShape = { from: { col: 1, row: 1, x: 0, y: 0 }, to: { col: 4, row: 5, x: 0, y: 0 }, geom: "group", name: "Gruppieren 3",
      children: [kid("rect", 0), kid("ellipse", 70)], frame: { w: 120, h: 60 } };
    const back = shapeFromXml(shapeToXml(g, 4, { w: 240, h: 120 }))!;
    expect(back.geom).toBe("group");
    expect(back.name).toBe("Gruppieren 3");
    expect(back.children?.map((c) => c.geom)).toEqual(["rect", "ellipse"]);
    // Children come back in the group's extent (here twice the frame).
    expect(back.children?.[1].rel).toEqual({ x: 140, y: 20, w: 100, h: 80 });
  });
});
