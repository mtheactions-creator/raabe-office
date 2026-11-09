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

describe("follow-up of the by-hand audit (D-152)", () => {
  it("writes Übergänge > Dauer on a slide without a transition as PowerPoint does: the duration alone", async () => {
    const { setTransitionDuration, readTransition, setTransition } = await import("../src/transitions");
    const d = new Deck();
    const p = d.slides[0].path;
    const app = { deck: d, panel: { selected: new Set([0]) } } as unknown as Parameters<typeof setTransitionDuration>[0];
    expect(readTransition(d.pptx.doc(p).documentElement)).toBe(null);
    setTransitionDuration(app, 3);
    const xml = new XMLSerializer().serializeToString(d.pptx.doc(p));
    // Recorded (probe_powerpoint_followup.py): <p:transition spd="slow" p14:dur="3000"/> in the Choice, no effect.
    expect(xml).toMatch(/<p:transition spd="slow" p14:dur="3000"\/>/);
    expect(xml).not.toMatch(/<p:cut\/>|<p:fade/);
    const t = readTransition(d.pptx.doc(p).documentElement)!;
    expect(t.kind).toBe("none");
    expect(t.dur).toBe(3000);
    // A second change keeps it effect-free; Strg+Z takes both back.
    setTransitionDuration(app, 2.25);
    expect(new XMLSerializer().serializeToString(d.pptx.doc(p))).toMatch(/p14:dur="2250"\/>/);
    d.undo();
    d.undo();
    expect(readTransition(d.pptx.doc(p).documentElement)).toBe(null);
    // An effect picked afterwards writes its element.
    setTransition(app, "fade");
    expect(readTransition(d.pptx.doc(p).documentElement)!.kind).toBe("fade");
    expect(new Deck(Pptx.read(d.write())).slides.length).toBe(d.slides.length);
  });
});
