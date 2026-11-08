// Lasting cases for what the by-hand sweep found (sweep_present.py, D-137): each failure triggered through the real
// controls gets a check here that runs without the UI.

import { describe, expect, it } from "vitest";
import { Pptx } from "@raabe/ooxml";
import { Deck } from "../src/deck";

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

  it("leaves the undo list alone when nothing changed inside a group", () => {
    const d = new Deck();
    d.group("nichts", () => {});
    expect(d.canUndo).toBe(false);
  });
});
