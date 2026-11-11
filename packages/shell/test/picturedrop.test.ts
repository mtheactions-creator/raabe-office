// The shared picture dropdown (D-161): the box shows the current value's picture, the list opens with showPopup,
// and the keyboard works as in Office's pane galleries.
import { afterEach, describe, expect, it } from "vitest";
import { closePopup, isPopupOpen } from "../src/popup";
import { Commands } from "../src/commands";
import { pictureDropdown, type PictureEntry } from "../src/picturedrop";
import { arrowPicture, dashPicture, shadowPicture } from "../src/pictures";

const entries: PictureEntry[] = ["solid", "sysDot", "dash", "lgDash"].map((v) => ({ value: v, label: v, svg: dashPicture(v) }));

function make(columns = 1, value = "solid") {
  let cur = value;
  const picks: string[] = [];
  const box = pictureDropdown({ label: "Strichtyp", entries, columns, size: [104, 20], value: () => cur, set: (v) => { cur = v; picks.push(v); } });
  document.body.append(box);
  return { box, picks };
}
const key = (el: Element, k: string, alt = false) => el.dispatchEvent(new KeyboardEvent("keydown", { key: k, altKey: alt, bubbles: true, cancelable: true }));
const list = () => document.querySelector<HTMLElement>(".picdrop-list");

afterEach(() => { closePopup(); document.body.replaceChildren(); });

describe("the picture dropdown", () => {
  it("shows the current value's picture in the closed box, not its name", () => {
    const { box } = make(1, "dash");
    expect(box.querySelector("svg")).not.toBeNull();
    expect(box.textContent).toBe("");
    expect(box.getAttribute("aria-label")).toBe("Strichtyp");
    expect(box.dataset.value).toBe("dash");
  });

  it("opens a list of pictures with the current one chosen and focused", () => {
    const { box } = make(1, "dash");
    box.click();
    expect(isPopupOpen()).toBe(true);
    const items = [...list()!.querySelectorAll<HTMLElement>("[role=option]")];
    expect(items.map((i) => i.getAttribute("aria-label"))).toEqual(["solid", "sysDot", "dash", "lgDash"]);
    expect(items.every((i) => i.querySelector("svg") && !i.textContent)).toBe(true);
    expect(document.activeElement).toBe(items[2]);
    expect(items[2].getAttribute("aria-selected")).toBe("true");
  });

  it("moves with the arrow keys, picks with Enter and gives the focus back to the box", () => {
    const { box, picks } = make();
    key(box, "Enter");
    key(document.activeElement!, "ArrowDown");
    key(document.activeElement!, "ArrowDown");
    key(document.activeElement!, "Enter");
    expect(picks).toEqual(["dash"]);
    expect(isPopupOpen()).toBe(false);
    expect(document.activeElement).toBe(box);
  });

  it("opens with Alt+Nach-unten and F4, jumps with Pos1 and Ende", () => {
    const { box, picks } = make();
    key(box, "ArrowDown", true);
    expect(isPopupOpen()).toBe(true);
    key(document.activeElement!, "End");
    key(document.activeElement!, " ");
    expect(picks).toEqual(["lgDash"]);
    key(box, "F4");
    expect(isPopupOpen()).toBe(true);
    key(document.activeElement!, "Home");
    expect(document.activeElement?.getAttribute("aria-label")).toBe("solid");
  });

  it("closes on Esc without a pick, and on a second click of the box", async () => {
    const { box, picks } = make();
    box.click();
    // The popup listens for Esc once the opening event has passed.
    await new Promise((r) => setTimeout(r, 0));
    window.dispatchEvent(new KeyboardEvent("keydown", { key: "Escape", bubbles: true, cancelable: true }));
    expect(isPopupOpen()).toBe(false);
    expect(picks).toEqual([]);
    box.click();
    expect(isPopupOpen()).toBe(true);
    box.click();
    expect(isPopupOpen()).toBe(false);
  });

  it("does not set the value it already has", () => {
    const { box, picks } = make(1, "solid");
    box.click();
    (list()!.querySelector("[data-value=solid]") as HTMLElement).click();
    expect(picks).toEqual([]);
  });

  it("previews an entry while the mouse is over it and ends the preview on leaving (Livevorschau, D-166)", () => {
    const seen: (string | null)[] = [];
    const commands = new Commands();
    commands.onHover((t) => { if (t === null) seen.push(null); else if (typeof t === "function") t(); });
    const box = pictureDropdown({ label: "Strichtyp", entries, columns: 1, size: [104, 20], value: () => "solid", set: () => undefined,
      live: { commands, preview: (v) => seen.push(v) } });
    document.body.append(box);
    box.click();
    const item = list()!.querySelector("[data-value=dash]")!;
    item.dispatchEvent(new MouseEvent("mouseenter"));
    item.dispatchEvent(new MouseEvent("mouseleave"));
    expect(seen).toEqual(["dash", null]);
  });

  it("draws our own pictures for arrows and shadows", () => {
    expect(arrowPicture("triangle", "end")).toContain("<polygon");
    expect(arrowPicture("none", "start")).not.toContain("<polygon");
    expect(shadowPicture("outer", 45)).toContain("feGaussianBlur");
  });
});
