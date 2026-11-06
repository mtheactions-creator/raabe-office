// .pptx reading, inheritance and writing (D-070), against presentations PowerPoint built itself
// (tools/feel-lab/make_pptx_fixtures.py: pptx-*.pptx with PowerPoint's read-back in pptx-*.json).

import { describe, expect, it } from "vitest";
import { readFileSync, existsSync } from "node:fs";
import { join } from "node:path";
import { Pptx, newPptx, slideView, LAYOUT_NAMES, type ShapeView } from "../src";

const FIX = join(__dirname, "fixtures");
const load = (name: string) => Pptx.read(new Uint8Array(readFileSync(join(FIX, name))));
const json = (name: string) => JSON.parse(readFileSync(join(FIX, name), "utf-8"));

const flat = (shapes: ShapeView[]): ShapeView[] => shapes.flatMap((s) => [s, ...(s.children ? flat(s.children) : [])]);

describe("new presentation (Office 2013-2022 design)", () => {
  it("has one Titelfolie and eleven layouts named as German PowerPoint", () => {
    const p = newPptx();
    expect(p.slides.length).toBe(1);
    const master = p.masters()[0];
    const names = p.layouts(master).map((l) => p.doc(l).documentElement.getElementsByTagNameNS("http://schemas.openxmlformats.org/presentationml/2006/main", "cSld")[0].getAttribute("name"));
    expect(names).toEqual(LAYOUT_NAMES);
    expect(p.slideSize).toEqual({ cx: 12192000, cy: 6858000 });
  });

  it("places every layout's placeholders where PowerPoint's own blank presentation has them", () => {
    if (!existsSync(join(FIX, "pptx-blank.pptx"))) return;
    const ours = newPptx(), theirs = load("pptx-blank.pptx");
    const lo = ours.layouts(ours.masters()[0]), lt = theirs.layouts(theirs.masters()[0]);
    expect(lo.length).toBe(lt.length);
    const diffs: string[] = [];
    lo.forEach((path, i) => {
      const a = slideView(ours, path).shapes, b = slideView(theirs, lt[i]).shapes;
      if (a.length !== b.length) diffs.push(`${LAYOUT_NAMES[i]}: ${a.length} vs ${b.length} shapes`);
      b.forEach((s, j) => {
        const o = a[j];
        if (!o) return;
        // The 2013-2022 fixture got English names from Office's design file: compare German names with ours only by number.
        const k = (v: ShapeView) => `${v.name.replace(/^.* (\d+)$/, "$1")}|${v.ph?.type}|${v.ph?.idx}|${v.x},${v.y},${v.w},${v.h}`;
        if (k(o) !== k(s)) diffs.push(`${LAYOUT_NAMES[i]}: ${k(o)} vs ${k(s)}`);
      });
    });
    expect(diffs).toEqual([]);
  });

  it("writes a package that reads back the same", () => {
    const p = newPptx();
    const back = Pptx.read(p.write());
    expect(back.slides.length).toBe(1);
    const v = slideView(back, back.slides[0].path);
    expect(v.shapes.map((s) => s.ph?.type)).toEqual(["ctrTitle", "subTitle"]);
    expect(v.shapes[0].text?.prompt).toEqual(["Titel durch Klicken hinzufügen"]);
    // The title inherits the Titelfolie layout's place and 60 pt, Calibri Light.
    expect([v.shapes[0].x, v.shapes[0].y, v.shapes[0].w, v.shapes[0].h]).toEqual([1524000, 1122363, 9144000, 2387600]);
    expect(v.shapes[0].text?.paras[0].end.size).toBe(60);
    expect(v.shapes[0].text?.paras[0].end.font).toBe("Calibri Light");
    expect(v.shapes[0].text?.body.anchor).toBe("b");
  });
});

describe("PowerPoint's fixtures", () => {
  const files = existsSync(FIX) ? ["pptx-layouts", "pptx-bullets", "pptx-shapes", "pptx-textbox", "pptx-group", "pptx-picture", "pptx-table", "pptx-autofit"] : [];
  for (const f of files) {
    it(`${f}: every shape where PowerPoint reports it`, () => {
      if (!existsSync(join(FIX, `${f}.pptx`))) return;
      const p = load(`${f}.pptx`), ref = json(`${f}.json`);
      const diffs: string[] = [];
      ref.slides.forEach((rs: { shapes: { name: string; id: number; emu: number[]; rotation: number; text?: { text?: string } }[] }, i: number) => {
        const v = slideView(p, p.slides[i].path);
        const all = flat(v.shapes);
        for (const s of rs.shapes) {
          // COM reports PowerPoint's internal (English) names; the file has the German ones: match by id.
          const o = all.find((x) => x.id === s.id) ?? all.find((x) => x.name === s.name);
          if (!o) { diffs.push(`${i + 1} ${s.name}: missing`); continue; }
          const ours = [o.x, o.y, o.w, o.h].map(Math.round);
          if (ours.some((v2, k) => Math.abs(v2 - s.emu[k]) > 1)) diffs.push(`${i + 1} ${s.name}: ${ours} vs ${s.emu}`);
          if (Math.abs(o.rot - s.rotation) > 0.01) diffs.push(`${i + 1} ${s.name}: rot ${o.rot} vs ${s.rotation}`);
        }
      });
      expect(diffs).toEqual([]);
    });
  }

  it("writes PowerPoint's files back byte for byte when nothing changed", () => {
    if (!existsSync(join(FIX, "pptx-shapes.pptx"))) return;
    const p = load("pptx-shapes.pptx");
    const back = Pptx.read(p.write());
    for (const [k, v] of Object.entries(p.files)) {
      if (k === "[Content_Types].xml") continue;
      if (!back.files[k]) continue;
      expect(Buffer.from(back.files[k]).equals(Buffer.from(v)), k).toBe(true);
    }
  });
});
