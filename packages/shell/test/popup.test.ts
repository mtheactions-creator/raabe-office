// How popups close (D-156): Office's rule as recorded in Excel, Word and PowerPoint on 2026-11-11.
import { afterEach, describe, expect, it, vi } from "vitest";
import { closePopup, showPopup } from "../src/popup";

// jsdom has no PointerEvent: a MouseEvent with the pointer type's name carries what the code reads (button, target).
function press(target: Element, type: string, button = 0): MouseEvent {
  const e = new MouseEvent(type, { bubbles: true, cancelable: true, button });
  target.dispatchEvent(e);
  return e;
}

/** A whole press as Chromium sends it: pointerdown, mousedown (unless pointerdown was cancelled), up, click. */
function fullPress(target: Element, button = 0): { down: MouseEvent; reached: string[] } {
  const reached: string[] = [];
  const types = ["pointerdown", "mousedown", "pointerup", "mouseup", "click"];
  for (const t of types) target.addEventListener(t, () => reached.push(t));
  const down = press(target, "pointerdown", button);
  if (!down.defaultPrevented) press(target, "mousedown", button);
  press(target, "pointerup", button);
  press(target, "mouseup", button);
  press(target, "click", button);
  return { down, reached };
}

async function openPopup(): Promise<{ pop: HTMLElement; closed: () => boolean }> {
  const pop = document.createElement("div");
  pop.className = "menu";
  let closed = false;
  showPopup(pop, new DOMRect(10, 10, 20, 20), () => { closed = true; });
  // The outside listeners are added after the opening event has passed.
  await new Promise((r) => setTimeout(r, 0));
  return { pop, closed: () => closed };
}

afterEach(() => {
  closePopup();
  document.body.replaceChildren();
  vi.useRealTimers();
});

describe("a press outside an open popup (D-156)", () => {
  it("in the document: closes the popup and the press does nothing else", async () => {
    const cell = document.body.appendChild(document.createElement("canvas"));
    const { closed } = await openPopup();
    const { reached } = fullPress(cell);
    expect(closed()).toBe(true);
    expect(reached).toEqual([]);
  });

  it("closes even when the surface cancels pointerdown itself (the grid), which stops mousedown", async () => {
    const grid = document.body.appendChild(document.createElement("div"));
    grid.addEventListener("pointerdown", (e) => e.preventDefault());
    const { closed } = await openPopup();
    fullPress(grid);
    expect(closed()).toBe(true);
  });

  it("a right press in the document closes the popup and still acts (the next context menu)", async () => {
    const cell = document.body.appendChild(document.createElement("canvas"));
    const { closed } = await openPopup();
    const { reached } = fullPress(cell, 2);
    expect(closed()).toBe(true);
    expect(reached).toContain("pointerdown");
    expect(reached).toContain("click");
  });

  it("in the app's frame (ribbon, status bar, task pane): closes the popup and the press acts", async () => {
    for (const cls of ["ribbon-card", "tabrow", "statusbar"]) {
      const frame = document.body.appendChild(document.createElement("div"));
      frame.className = cls;
      const button = frame.appendChild(document.createElement("button"));
      const { closed } = await openPopup();
      const { reached } = fullPress(button);
      expect(closed()).toBe(true);
      expect(reached).toEqual(["pointerdown", "mousedown", "pointerup", "mouseup", "click"]);
    }
  });

  it("the next press after a swallowed one is not swallowed", async () => {
    const cell = document.body.appendChild(document.createElement("canvas"));
    await openPopup();
    fullPress(cell);
    const { reached } = fullPress(cell);
    expect(reached).toEqual(["pointerdown", "mousedown", "pointerup", "mouseup", "click"]);
  });

  it("a press inside the popup does not close it", async () => {
    const { pop, closed } = await openPopup();
    const item = pop.appendChild(document.createElement("button"));
    fullPress(item);
    expect(closed()).toBe(false);
  });

  it("Escape and the window's blur close it", async () => {
    let p = await openPopup();
    document.dispatchEvent(new KeyboardEvent("keydown", { key: "Escape", bubbles: true }));
    expect(p.closed()).toBe(true);
    p = await openPopup();
    window.dispatchEvent(new Event("blur"));
    expect(p.closed()).toBe(true);
  });

  it("a popup without a position of its own floats at its anchor", async () => {
    const pop = document.createElement("div");
    showPopup(pop, new DOMRect(40, 50, 20, 10));
    expect(pop.style.position).toBe("fixed");
    expect(pop.classList.contains("popup-card")).toBe(true);
  });
});
