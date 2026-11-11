// The line submenus of Formkontur, Bildrahmen and Textkontur (Stärke, Striche, Pfeile) as Office shows them: Stärke
// as lines of each weight with their names, Striche and Pfeile as pictures only (their names are the tooltips).
// Shared by the three apps (D-162); the pictures are ours (pictures.ts).

import type { MenuEntry } from "./popup";
import { pictureMenu, weightPicture, type PictureEntry, type PictureLive } from "./picturedrop";
import type { Commands } from "./commands";
import { PIC_SIZE, arrowPicture, dashEntries } from "./pictures";

/** Office's Pfeile menu: Pfeilart 1 to 11 as start and end of the line (a:headEnd / a:tailEnd types). */
export const ARROW_STYLES: [string | null, string | null][] = [
  [null, null], [null, "arrow"], ["arrow", null], ["arrow", "arrow"], [null, "triangle"], ["triangle", null], ["triangle", "triangle"],
  ["oval", "triangle"], ["oval", "oval"], ["diamond", "diamond"], [null, "stealth"],
];

/** A line with an arrowhead at either end (Pfeilart). */
export function arrowPairPicture(head: string | null, tail: string | null, w = 104, h = 20): string {
  // Two half pictures, the start's left half and the end's right half, side by side.
  const left = arrowPicture(head ?? "none", "start", w, h).replace(/^<svg[^>]*>|<\/svg>$/g, "");
  const right = arrowPicture(tail ?? "none", "end", w, h).replace(/^<svg[^>]*>|<\/svg>$/g, "");
  return `<svg xmlns="http://www.w3.org/2000/svg" width="${w}" height="${h}" viewBox="0 0 ${w} ${h}" aria-hidden="true">`
    + `<clipPath id="apl${w}${h}"><rect width="${w / 2}" height="${h}"/></clipPath><clipPath id="apr${w}${h}"><rect x="${w / 2}" width="${w / 2}" height="${h}"/></clipPath>`
    + `<g clip-path="url(#apl${w}${h})">${left}</g><g clip-path="url(#apr${w}${h})">${right}</g></svg>`;
}

/** Stärke: each weight as its line and its name. */
export function weightMenu(weights: [string, number][], pick: (pt: number) => void, commands?: Commands): MenuEntry[] {
  // With the app's commands, each line is previewed while the mouse is over it (Livevorschau, D-166).
  return weights.map(([label, pt]) => ({ label, preview: weightPicture(pt), run: () => pick(pt), ...(commands ? { hover: () => pick(pt) } : {}) }));
}

/** Striche: the dashes as pictures (Uni, Runder Punkt …). */
export function dashMenu(current: string | undefined, pick: (dash: string) => void, commands?: Commands): MenuEntry {
  const live: PictureLive | undefined = commands ? { commands, preview: pick } : undefined;
  return { block: () => pictureMenu(dashEntries(), 1, PIC_SIZE.dash, current, pick, live) };
}

/** Pfeile: Pfeilart 1-11 as pictures. */
export function arrowStyleMenu(pick: (head: string | null, tail: string | null) => void, commands?: Commands): MenuEntry {
  const entries: PictureEntry[] = ARROW_STYLES.map(([hd, tl], i) => ({ value: String(i), label: `Pfeilart ${i + 1}`, svg: arrowPairPicture(hd, tl, ...PIC_SIZE.dash) }));
  const apply = (v: string) => { const [hd, tl] = ARROW_STYLES[Number(v)]; pick(hd, tl); };
  return { block: () => pictureMenu(entries, 1, PIC_SIZE.dash, undefined, apply, commands ? { commands, preview: apply } : undefined) };
}
