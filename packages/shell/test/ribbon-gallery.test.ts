// Office's inline galleries show one row of entries and scroll by rows (D-180).
import { describe, expect, it } from "vitest";
import { showGalleryRow } from "../src/ribbon";

const items = (n: number) => Array.from({ length: n }, () => document.createElement("button"));
const shown = (list: HTMLElement[]) => list.map((b, i) => (b.style.display === "none" ? -1 : i)).filter((i) => i >= 0);

describe("inline gallery rows", () => {
  it("shows the first row and only enables Zeile nach unten", () => {
    const list = items(10), up = document.createElement("button"), down = document.createElement("button");
    expect(showGalleryRow(list, 4, 0, up, down)).toBe(0);
    expect(shown(list)).toEqual([0, 1, 2, 3]);
    expect(up.disabled).toBe(true);
    expect(down.disabled).toBe(false);
  });

  it("the last row holds the rest; scrolling past it stays there", () => {
    const list = items(10), up = document.createElement("button"), down = document.createElement("button");
    expect(showGalleryRow(list, 4, 5, up, down)).toBe(2);
    expect(shown(list)).toEqual([8, 9]);
    expect(up.disabled).toBe(false);
    expect(down.disabled).toBe(true);
  });

  it("a gallery that fits in one row has no scrolling", () => {
    const list = items(3), up = document.createElement("button"), down = document.createElement("button");
    showGalleryRow(list, 5, 0, up, down);
    expect(shown(list)).toEqual([0, 1, 2]);
    expect(up.disabled && down.disabled).toBe(true);
  });
});
