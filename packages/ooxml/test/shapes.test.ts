import { readFileSync, writeFileSync } from "node:fs";
import { join } from "node:path";
import { describe, expect, it } from "vitest";
import { strFromU8, unzipSync } from "fflate";
import { DEFAULT_SHAPE_BODY, Doc, type FloatObject, type Para } from "@raabe/text-engine";
import { readDocx, writeDocx } from "../src";

const fixture = (name: string) => new Uint8Array(readFileSync(join(__dirname, "fixtures", name)));
const floats = (paras: readonly Para[]) => paras.flatMap((p) => p.floats ?? []);

describe("shapes and text boxes (Word's own)", () => {
  it("reads Word's shapes: kinds, geometry, style colours, text", () => {
    const { state, partial } = readDocx(fixture("word-shapes.docx"));
    expect(partial).toEqual([]);
    const fs = floats(state.paras).filter((f) => f.shape);
    expect(fs.length).toBe(9);
    const rect = fs.find((f) => !f.pic?.geometry && f.shape!.kind === "shape")!;
    // Office 2013-2022 theme: accent1 fill, its 15 % shade (in linear light) as the outline.
    expect(rect.shape!.styleFill).toEqual({ kind: "solid", color: "4472C4" });
    expect(rect.shape!.styleLine).toEqual({ color: "172C51", width: 1 });
    expect(rect.shape!.fontColor).toBe("FFFFFF");
    expect(state.stories?.get(rect.shape!.story!)?.[0].text).toBe("Rechteck mit Text");
    const line = fs.find((f) => f.shape!.kind === "connector")!;
    expect(line.pic?.line?.tail?.type).toBe("triangle");
    const box = fs.find((f) => f.shape!.kind === "textbox" && f.wrap === "front")!;
    expect(state.stories?.get(box.shape!.story!)?.map((p) => p.text)).toEqual(["Ein Textfeld", "mit zwei Absätzen."]);
    expect(box.shape!.body.anchor).toBe("t");
    const grad = fs.find((f) => f.pic?.fill?.kind === "gradient");
    expect(grad).toBeTruthy();
  });

  it("writes unchanged shapes back as they were and edited text into the drawing", () => {
    const { state, pkg } = readDocx(fixture("word-shapes.docx"));
    const d = new Doc(state);
    const box = floats(d.main).find((f) => f.shape?.kind === "textbox" && f.wrap === "front")!;
    const story = box.shape!.story!;
    d.setStory(story, d.stories.get(story)!.map((p, i) => (i === 0 ? { ...p, text: "Geändert", runs: [{ len: 8, props: p.runs[0].props }] } : p)));
    const bytes = writeDocx(d.state(), pkg);
    const xml = strFromU8(unzipSync(bytes)["word/document.xml"]);
    expect(xml).toContain("Geändert");
    // The star (no text) keeps its XML with the VML fallback.
    expect(xml).toContain('prst="star5"');
    const again = readDocx(bytes).state;
    const fs = floats(again.paras).filter((f) => f.shape);
    expect(fs.length).toBe(9);
    const box2 = fs.find((f) => f.shape?.kind === "textbox" && f.wrap === "front")!;
    expect(again.stories?.get(box2.shape!.story!)?.[0].text).toBe("Geändert");
    if (process.env.SHAPES_DOCX_OUT) writeFileSync(process.env.SHAPES_DOCX_OUT, bytes);
  });

  it("writes new shapes as Word does (style, text, theme part)", () => {
    const { state } = readDocx(writeDocx(new Doc({}).state(), null));
    const d = new Doc(state);
    const shape: FloatObject = {
      at: 0, data: undefined, width: 120, height: 60, x: { rel: "column", offset: 10 }, y: { rel: "paragraph", offset: 5 },
      wrap: "front", side: "bothSides", dist: [0, 0, 9, 9], pic: { geometry: "roundRect" },
      shape: { kind: "shape", styleFill: { kind: "solid", color: "4472C4" }, body: DEFAULT_SHAPE_BODY, story: "tx:1", adjust: { adj: 30000 } },
    };
    d.setStory("tx:1", [{ text: "Hallo", runs: [{ len: 5, props: d.main[0].mark }], props: { align: "center" }, mark: d.main[0].mark }]);
    d.setStory("main", [{ ...d.main[0], floats: [shape] }]);
    const bytes = writeDocx(d.state(), null);
    const files = unzipSync(bytes);
    expect(Object.keys(files)).toContain("word/theme/theme1.xml");
    const xml = strFromU8(files["word/document.xml"]);
    expect(xml).toContain('<a:prstGeom prst="roundRect"><a:avLst><a:gd name="adj" fmla="val 30000"/></a:avLst></a:prstGeom>');
    expect(xml).toContain('<a:lnRef idx="2"><a:schemeClr val="accent1"><a:shade val="15000"/></a:schemeClr></a:lnRef>');
    expect(xml).toContain("<w:txbxContent>");
    const again = readDocx(bytes).state;
    const f = floats(again.paras)[0];
    expect(f.shape?.styleLine?.color).toBe("172C51");
    expect(again.stories?.get(f.shape!.story!)?.[0].text).toBe("Hallo");
  });
});
