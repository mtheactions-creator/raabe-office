// Lasting cases for what the by-hand sweep found (sweep_present.py, D-137): each failure triggered through the real
// controls gets a check here that runs without the UI.

import { describe, expect, it } from "vitest";
import { Pptx } from "@raabe/ooxml";
import { Deck } from "../src/deck";
import { colorEl, colorXml } from "../src/colors";

describe("by-hand sweep findings (D-137)", () => {
  it("takes back several slides made at once with one Strg+Z (Folien aus Gliederung, Folien wiederverwenden)", () => {
    const d = new Deck();
    const lay = d.layouts().find((l) => l.name === "Titel und Inhalt")!.path;
    const n = d.slides.length;
    d.group("Folien aus Gliederung", () => {
      d.addSlide(0, lay);
      d.addSlide(1, lay);
      d.addSlide(2, lay);
    });
    expect(d.slides.length).toBe(n + 3);
    expect(d.undoLabel).toBe("Folien aus Gliederung");
    expect(d.undo()).toBe(true);
    expect(d.slides.length).toBe(n);
    expect(d.canUndo).toBe(false);
    // Wiederholen brings all three back, and the file written then has them.
    expect(d.redo()).toBe(true);
    expect(d.slides.length).toBe(n + 3);
    expect(new Deck(Pptx.read(d.write())).slides.length).toBe(n + 3);
  });

  it("never writes a palette's Automatisch as a colour named auto (PowerPoint refuses such a file)", () => {
    const doc = new DOMParser().parseFromString(`<a:r xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/>`, "application/xml");
    expect(colorEl(doc, "auto").getAttribute("val")).toBe("tx1");
    expect(colorEl(doc, "AUTO").getAttribute("val")).toBe("tx1");
    expect(colorXml("auto")).toBe(`<a:schemeClr val="tx1"/>`);
    expect(colorEl(doc, "ff0000").getAttribute("val")).toBe("FF0000");
  });

  it("makes no undo step for a change that changed nothing (a value written again when its field is left)", () => {
    const d = new Deck();
    const p = d.slides[0].path;
    d.change("Höhe", [p], () => {});
    expect(d.canUndo).toBe(false);
    d.change("Höhe", [p], () => { d.pptx.doc(p).documentElement.setAttribute("show", "0"); });
    d.change("Höhe", [p], () => { d.pptx.doc(p).documentElement.setAttribute("show", "0"); });
    expect(d.undoLabel).toBe("Höhe");
    d.undo();
    expect(d.pptx.doc(p).documentElement.getAttribute("show")).toBe(null);
  });

  it("leaves the undo list alone when nothing changed inside a group", () => {
    const d = new Deck();
    d.group("nichts", () => {});
    expect(d.canUndo).toBe(false);
  });
});
