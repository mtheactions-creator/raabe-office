// Picture effects (D-072): the adjustments against pixels PowerPoint rendered (probe_powerpoint_picfx.py,
// results/2026-11-07-powerpoint-picfx*.json), and the effect list in PowerPoint's XML.
import { describe, expect, it } from "vitest";
import { applyEffect, effectsOf, temperatureFit, type Pixels } from "../src/picfx";

function px(...rgb: [number, number, number][]): Pixels {
  const data = new Uint8ClampedArray(rgb.length * 4);
  rgb.forEach(([r, g, b], i) => data.set([r, g, b, 255], i * 4));
  return { width: rgb.length, height: 1, data };
}
const rgbAt = (p: Pixels, i: number) => [p.data[i * 4], p.data[i * 4 + 1], p.data[i * 4 + 2]];
const near = (a: number[], b: number[], tol: number) => a.every((v, i) => Math.abs(v - b[i]) <= tol);

describe("picture adjustments as PowerPoint renders them", () => {
  it("Sättigung: 0 % gives HSL lightness", () => {
    const p = applyEffect(px([242, 157, 48], [30, 60, 120], [240, 200, 60]), { name: "saturation", attrs: { sat: "0" } });
    expect(rgbAt(p, 0)).toEqual([145, 145, 145]);
    expect(rgbAt(p, 1)).toEqual([75, 75, 75]);
    expect(rgbAt(p, 2)).toEqual([150, 150, 150]);
  });
  it("Sättigung: 200 %", () => {
    const p = applyEffect(px([242, 157, 48], [240, 200, 60], [30, 60, 120]), { name: "saturation", attrs: { sat: "200000" } });
    expect(near(rgbAt(p, 0), [248, 158, 42], 1)).toBe(true);
    expect(near(rgbAt(p, 1), [247, 204, 53], 1)).toBe(true);
    expect(near(rgbAt(p, 2), [17, 56, 133], 4)).toBe(true);
  });
  it("Temperatur: 4700 K and 11200 K on grey", () => {
    const a = applyEffect(px([191, 191, 191], [63, 63, 63]), { name: "colorTemperature", attrs: { colorTemp: "4700" } });
    expect(near(rgbAt(a, 0), [174, 194, 223], 3)).toBe(true);
    expect(near(rgbAt(a, 1), [55, 64, 79], 3)).toBe(true);
    const b = applyEffect(px([191, 191, 191]), { name: "colorTemperature", attrs: { colorTemp: "11200" } });
    expect(near(rgbAt(b, 0), [211, 189, 160], 3)).toBe(true);
    expect(temperatureFit(6500)).toEqual([1, 0, 1, 0, 1, 0]);
  });
  it("Weichzeichnen softens an edge, Schärfen steepens it", () => {
    const edge = () => ({ width: 1, height: 4, data: new Uint8ClampedArray([79, 79, 79, 255, 79, 79, 79, 255, 30, 60, 120, 255, 30, 60, 120, 255]) });
    const soft = applyEffect(edge(), { name: "sharpenSoften", attrs: { amount: "-50000" } });
    expect(soft.data[4]).toBeLessThan(79);
    expect(soft.data[8]).toBeGreaterThan(30);
    const sharp = applyEffect(edge(), { name: "sharpenSoften", attrs: { amount: "50000" } });
    expect(sharp.data[4]).toBeGreaterThan(79);
    expect(sharp.data[8]).toBeLessThan(30);
  });
});

describe("Hintergrund entfernen", () => {
  it("removes the plain background around an object and keeps the object", () => {
    const w = 20, h = 20, data = new Uint8ClampedArray(w * h * 4);
    for (let y = 0; y < h; y++) for (let x = 0; x < w; x++) {
      const obj = x >= 7 && x < 13 && y >= 7 && y < 13;
      data.set(obj ? [200, 30, 30, 255] : [200, 215, 230 - y], (y * w + x) * 4);
    }
    const p = applyEffect({ width: w, height: h, data }, { name: "backgroundRemoval", attrs: {} });
    expect(p.data[3]).toBe(0);
    expect(p.data[(10 * w + 10) * 4 + 3]).toBe(255);
    expect(p.data[(10 * w + 3) * 4 + 3]).toBe(0);
  });
});

describe("Freistellen marks", () => {
  const picture = () => {
    // A red object left and a blue-grey one right on a light background.
    const w = 30, h = 20, data = new Uint8Array(w * h * 4);
    for (let y = 0; y < h; y++) for (let x = 0; x < w; x++) {
      const left = x >= 5 && x < 12 && y >= 6 && y < 14, right = x >= 18 && x < 25 && y >= 6 && y < 14;
      data.set(left ? [200, 30, 30, 255] : right ? [180, 190, 200, 255] : [200, 215, 230, 255], (y * w + x) * 4);
    }
    return { width: w, height: h, data: new Uint8ClampedArray(data) };
  };
  it("a line to remove takes an object away, a line to keep saves a background-like one", () => {
    const at = (p: { data: Uint8ClampedArray }, x: number, y: number) => p.data[(y * 30 + x) * 4 + 3];
    const plain = applyEffect(picture(), { name: "backgroundRemoval", attrs: {} });
    expect(at(plain, 8, 10)).toBe(255);
    const removed = applyEffect(picture(), { name: "backgroundRemoval", attrs: {}, marks: [{ keep: false, x1: 20000, y1: 50000, x2: 36000, y2: 50000 }] });
    expect(at(removed, 8, 10)).toBe(0);
    const kept = applyEffect(picture(), { name: "backgroundRemoval", attrs: {}, marks: [{ keep: true, x1: 63000, y1: 40000, x2: 80000, y2: 60000 }] });
    expect(at(kept, 21, 10)).toBe(255);
  });
});

describe("the effect list in the blip", () => {
  it("reads PowerPoint's imgProps", () => {
    const xml = `<a:blip xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="rId2"><a:extLst><a:ext uri="{BEBA8EAE-BF5A-486C-A8C5-ECC9F3942E4B}"><a14:imgProps xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main"><a14:imgLayer r:embed="rId3"><a14:imgEffect><a14:artisticBlur/></a14:imgEffect><a14:imgEffect><a14:saturation sat="0"/></a14:imgEffect></a14:imgLayer></a14:imgProps></a:ext></a:extLst></a:blip>`;
    const blip = new DOMParser().parseFromString(xml, "application/xml").documentElement;
    expect(effectsOf(blip)).toEqual([{ name: "artisticBlur", attrs: {} }, { name: "saturation", attrs: { sat: "0" } }]);
  });
});
