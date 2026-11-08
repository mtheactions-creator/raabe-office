// Pictures with Bildformat in Excel's own files (D-087): xlsx-pictures.xlsx and excel-pictures.json were
// written by German Excel 365 (tools/feel-lab/probe_excel_pictures.py), one picture per setting.
import { readFileSync } from "node:fs";
import { unzipSync, strFromU8 } from "fflate";
import { describe, expect, it } from "vitest";
import type { PictureFormat } from "@raabe/text-engine";
import { PICTURE_STYLES } from "@raabe/text-engine";
import { readXlsx, writeXlsx } from "../src";

const FILE = readFileSync("packages/ooxml/test/fixtures/xlsx-pictures.xlsx");
const REC = JSON.parse(readFileSync("packages/ooxml/test/fixtures/excel-pictures.json", "utf-8")) as {
  cases: { case: string; anchor: string | null; com: Record<string, unknown>; media?: { original: boolean }[] }[];
  styles?: { i: number; name: string; spPr: string }[];
};

const pics = () => readXlsx(FILE).sheets[0].pictures ?? [];
const byCase = (list: ReturnType<typeof pics>, c: string) => list.find((p) => p.name === `Bild ${c}`);
const fmt = (c: string): PictureFormat => (byCase(pics(), c)?.format ?? {}) as PictureFormat;
const drawing = (bytes: Uint8Array) => strFromU8(unzipSync(bytes)["xl/drawings/drawing1.xml"]);

/** XML without attribute order, namespace declarations, the transform's offset and Excel's creation ids, for comparing. */
function norm(xml: string): string {
  const doc = new DOMParser().parseFromString(`<r xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">${xml}</r>`, "application/xml");
  const walk = (el: Element): string => {
    if (el.localName === "off" || (el.localName === "ext" && el.getAttribute("uri") === "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}")) return "";
    if (el.localName === "extLst" && !Array.from(el.children).some((c) => walk(c))) return "";
    const attrs = Array.from(el.attributes).filter((a) => !a.name.startsWith("xmlns") && a.name !== "id").map((a) => `${a.name}=${a.value}`).sort().join(" ");
    return `<${el.localName}${attrs ? " " + attrs : ""}>${Array.from(el.children).map(walk).join("")}</${el.localName}>`;
  };
  return Array.from(doc.documentElement.children).map(walk).join("");
}

describe("Bildformat read from Excel's file", () => {
  it("reads every recorded setting", () => {
    const list = pics();
    // 30 pictures with one setting each and the picture of the group.
    expect(list.length).toBe(31);
    expect(byCase(list, "group")!.group).toBeDefined();
    expect(fmt("plain")).toEqual({});
    const crop = fmt("crop").crop!;
    expect([crop.l, crop.t, crop.r, crop.b]).toEqual([0.05556, 0.06667, 0.11111, 0.03333]);
    expect(fmt("lum").lum).toEqual({ bright: 0.4, contrast: -0.2 });
    expect(fmt("grayscale").recolor).toEqual({ kind: "grayscale" });
    expect(fmt("bw").recolor).toEqual({ kind: "bw", threshold: 0.5 });
    expect(fmt("washout").recolor).toEqual({ kind: "washout" });
    expect(fmt("transparentColor").transparentColor).toBe("FFFFFF");
    expect(fmt("brightnessContrast")).toMatchObject({ brightness: 0.2, contrast: -0.2 });
    expect(fmt("saturation").saturation).toBe(2);
    expect(fmt("temperature").temperature).toBe(8800);
    expect(fmt("sharpen").sharpen).toBe(0.5);
    expect(fmt("soften").sharpen).toBe(-0.5);
    expect(fmt("rotate30").rotation).toBe(30);
    expect(fmt("flipH").flipH).toBe(true);
    expect(fmt("flipV").flipV).toBe(true);
    expect(fmt("line").line).toMatchObject({ color: "C00000", width: 4.5, dash: "dash" });
    expect(fmt("shadowOuter").shadow).toMatchObject({ kind: "outer", color: "000000", blur: 6 });
    expect(fmt("shadowOuter").shadow!.dir).toBeCloseTo(45, 2);
    expect(fmt("shadowOuter").shadow!.transparency).toBeCloseTo(0.6);
    expect(fmt("shadowInner").shadow?.kind).toBe("inner");
    expect(fmt("glow").glow).toMatchObject({ color: "4472C4", radius: 8 });
    expect(fmt("softEdge").softEdge).toBe(6);
    expect(fmt("reflection").reflection?.size).toBe(0.55);
    expect(fmt("bevel").bevel).toMatchObject({ kind: "circle", w: 6, h: 6 });
    expect(fmt("rotation3d").scene3dXml).toContain('lat="1200000"');
    expect(fmt("oval").geometry).toBe("ellipse");
    expect(fmt("roundRect").geometry).toBe("roundRect");
    const alt = byCase(list, "altText")!;
    expect([alt.alt, alt.title]).toEqual(["Ein Testbild mit vier Farben", "Testbild"]);
    expect(byCase(list, "hidden")!.hidden).toBe(true);
    expect(byCase(list, "plain")!.lockAspect).toBeUndefined();
  });

  it("knows which pictures hold their Office 2010 effects in their pixels (Excel draws them as stored)", () => {
    const list = pics();
    for (const c of ["brightnessContrast", "saturation", "temperature", "sharpen", "soften", "artistic"]) {
      expect(byCase(list, c)!.baked, c).toBeDefined();
      expect(byCase(list, c)!.media?.layer, c).toBeTruthy();
    }
    expect(byCase(list, "brightnessContrast")!.baked).toBe('{"brightness":0.2,"contrast":-0.2}');
    // The recording: such pictures are stored changed, the others as inserted.
    for (const r of REC.cases) {
      if (!r.media) continue;
      const baked = ["brightnessContrast", "saturation", "temperature", "sharpen", "soften", "artistic"].includes(r.case);
      expect(r.media[0].original, r.case).toBe(!baked);
    }
  });
});

describe("Bildformat written back", () => {
  it("keeps the drawing as Excel wrote it while nothing changed", () => {
    const wb = readXlsx(FILE);
    expect(drawing(writeXlsx(wb))).toBe(drawing(FILE));
  });

  it("writes every changed picture as Excel does (same DrawingML) and reads it back the same", () => {
    const wb = readXlsx(FILE);
    const before = wb.sheets[0].pictures!.filter((p) => !p.group).map((p) => ({ ...p }));
    // A change that leaves the look as it is: the anchor is rewritten from the model.
    for (const p of wb.sheets[0].pictures!) if (!p.group) p.lookKey = "changed";
    const out = writeXlsx(wb);
    const xml = drawing(out);
    const back = readXlsx(out).sheets[0].pictures!;
    expect(back.length).toBe(before.length + 1);
    const same: string[] = [], differ: string[] = [];
    for (const p of before) {
      const q = back.find((x) => x.name === p.name)!;
      expect(q.format, p.name).toEqual(p.format);
      expect([q.alt, q.title, q.hidden, q.lockAspect, q.baked], p.name).toEqual([p.alt, p.title, p.hidden, p.lockAspect, p.baked]);
      const ours = new RegExp(`<xdr:twoCellAnchor\\b(?:(?!</xdr:twoCellAnchor>)[\\s\\S])*?name="${p.name}"[\\s\\S]*?</xdr:twoCellAnchor>`).exec(xml)?.[0] ?? "";
      (norm(ours) === norm(p.raw!) ? same : differ).push(`${p.name}\n  excel: ${norm(p.raw!)}\n  ours:  ${norm(ours)}`);
    }
    // Excel's XML for each setting, element for element.
    expect(differ.join("\n")).toBe("");
    expect(same.length).toBe(30);
  });

  it("keeps a group holding a picture", () => {
    const wb = readXlsx(FILE);
    wb.sheets[0].pictures![0] = { ...wb.sheets[0].pictures![0], format: { rotation: 10 } };
    const xml = drawing(writeXlsx(wb));
    expect(xml).toContain('name="Bild group"');
    expect(xml).toContain('name="Rechteck group"');
  });

  it("writes a new picture with its format, and Korrekturen with the picture they were worked into", () => {
    const wb = readXlsx(FILE);
    const plain = wb.sheets[0].pictures![0];
    const png = plain.src;
    wb.sheets[0].pictures!.push({ from: { col: 30, row: 2, x: 0, y: 0 }, to: { col: 32, row: 6, x: 0, y: 0 }, size: { w: 128, h: 85 }, src: png, name: "Grafik 99",
      format: { brightness: 0.2, shadow: { kind: "outer", color: "000000", transparency: 0.6, blur: 4, dist: 3, dir: 45 } }, baked: '{"brightness":0.2}', bakedSrc: png.replace(/.{8}$/, "AAAAAAA=") });
    const out = writeXlsx(wb);
    const xml = drawing(out);
    const a = /<xdr:twoCellAnchor\b(?:(?!<\/xdr:twoCellAnchor>)[\s\S])*?name="Grafik 99"[\s\S]*?<\/xdr:twoCellAnchor>/.exec(xml)![0];
    expect(a).toContain('<a14:brightnessContrast bright="20000"/>');
    expect(a).toContain('<a:picLocks noChangeAspect="1"/>');
    expect(a).toContain("<a:outerShdw");
    const back = readXlsx(out).sheets[0].pictures!.find((p) => p.name === "Grafik 99")!;
    expect((back.format as PictureFormat).brightness).toBe(0.2);
    // The stored picture is the one with the effect worked in.
    expect(back.src).toBe(wb.sheets[0].pictures!.at(-1)!.bakedSrc);
  });
});

describe("Excel's picture styles", () => {
  it("write the same DrawingML as Word's 28 recorded ones", () => {
    const styles = REC.styles ?? [];
    expect(styles.length).toBe(28);
    const differ: string[] = [];
    styles.forEach((s, i) => {
      const w = PICTURE_STYLES[i];
      if (w.label !== s.name || norm(w.xml) !== norm(s.spPr)) differ.push(`${i} ${s.name} / ${w.label}\n  excel: ${norm(s.spPr)}\n  word:  ${norm(w.xml)}`);
    });
    expect(differ.join("\n")).toBe("");
  });
});
