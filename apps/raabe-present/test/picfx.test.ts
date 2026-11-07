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

describe("the effect list in the blip", () => {
  it("reads PowerPoint's imgProps", () => {
    const xml = `<a:blip xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="rId2"><a:extLst><a:ext uri="{BEBA8EAE-BF5A-486C-A8C5-ECC9F3942E4B}"><a14:imgProps xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main"><a14:imgLayer r:embed="rId3"><a14:imgEffect><a14:artisticBlur/></a14:imgEffect><a14:imgEffect><a14:saturation sat="0"/></a14:imgEffect></a14:imgLayer></a14:imgProps></a:ext></a:extLst></a:blip>`;
    const blip = new DOMParser().parseFromString(xml, "application/xml").documentElement;
    expect(effectsOf(blip)).toEqual([{ name: "artisticBlur", attrs: {} }, { name: "saturation", attrs: { sat: "0" } }]);
  });
});
