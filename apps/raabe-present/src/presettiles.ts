// Pictures of PowerPoint's effect presets (Schatten, Spiegelung, Leuchten, Weiche Kanten, Abschrägung, 3D-Drehung,
// Voreinstellung) for the galleries of Formeffekte / Texteffekte and the Voreinstellungen of Form formatieren (D-161,
// D-162): each preset applied to a small blue rectangle and drawn by our own shape renderer, as PowerPoint's galleries
// show the effect on a shape. Nothing of Office's pictures is used.

import { DEFAULT_THEME_COLORS, pictureFromSpPr } from "@raabe/ooxml";
import { shapeElement, type PictureFormat } from "@raabe/text-engine";
import type { PicturePreset } from "./effect-presets";

const BLUE = "4472C4";

/** A gallery tile of `w` x `h` px with the preset's effects on a rectangle; null: the plain rectangle ("Kein …"). */
export function presetTile(p: PicturePreset | null, w: number, h: number): HTMLElement {
  const box = document.createElement("span");
  box.style.cssText = `position:relative;display:block;width:${w}px;height:${h}px;overflow:hidden`;
  const look = p ? (pictureFromSpPr(p.xml, DEFAULT_THEME_COLORS) as PictureFormat) : {};
  const f: PictureFormat = { ...look, fill: { kind: "solid", color: BLUE }, geometry: "rect", rotation: undefined, flipH: undefined, flipV: undefined };
  const el = shapeElement(w * 0.5, h * 0.5, f, { kind: "shape", adjust: {}, body: { l: 0, t: 0, r: 0, b: 0, anchor: "t", wrap: false, autofit: "none" } });
  el.style.left = `${w * 0.25}px`;
  el.style.top = `${h * 0.2}px`;
  box.append(el);
  return box;
}
