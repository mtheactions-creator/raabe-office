// The entries of the Bildformat tab's galleries as Office 365 has them (recorded from Word with
// tools/feel-lab/probe_word_picture.py; Excel's tab has the same ones, D-087): Korrekturen, Farbe,
// Transparenz. Shared by RaabeWrite and RaabeCalc. Also Freistellen / Hintergrund entfernen.

import type { PictureFormat } from "./picture";

export interface GalleryItem {
  label: string;
  /** The settings the entry gives (merged into the picture's). */
  format: PictureFormat;
}

const pct = (x: number) => `${Math.round(x * 100)} %`;
const sign = (x: number) => (x > 0 ? "+" : x < 0 ? "−" : "") + pct(Math.abs(x)).replace(" %", " %");

// ----- Korrekturen -----------------------------------------------------------------------
export const SHARPEN: GalleryItem[] = [-0.5, -0.25, 0, 0.25, 0.5].map((v) => ({
  label: v < 0 ? `Weichzeichnen: ${pct(-v)}` : v > 0 ? `Schärfen: ${pct(v)}` : "Schärfen und Weichzeichnen: 0 %",
  format: { sharpen: v || undefined },
}));
export const BRIGHT_CONTRAST: GalleryItem[] = [-0.4, -0.2, 0, 0.2, 0.4].flatMap((c) => [-0.4, -0.2, 0, 0.2, 0.4].map((b) => ({
  label: `Helligkeit: ${sign(b)}${b === 0 ? " (Normal)" : ""} Kontrast: ${sign(c)}${c === 0 ? " (Normal)" : ""}`,
  format: { brightness: b || undefined, contrast: c || undefined },
})));

// ----- Farbe -----------------------------------------------------------------------------
export const SATURATION: GalleryItem[] = [0, 0.33, 0.66, 1, 2, 3, 4].map((v) => ({
  label: `Sättigung: ${pct(v)}`, format: { saturation: v === 1 ? undefined : v },
}));
export const TEMPERATURE: GalleryItem[] = [4700, 5300, 5900, 6500, 7200, 8800, 11200].map((k) => ({
  label: `Temperatur: ${k.toLocaleString("de-DE")} K`, format: { temperature: k === 6500 ? undefined : k },
}));

// ----- Transparenz -----------------------------------------------------------------------
export const TRANSPARENCY: GalleryItem[] = [0, 0.15, 0.3, 0.5, 0.65, 0.8, 0.95].map((t) => ({
  label: `Transparenz: ${pct(t)}`, format: { transparency: t || undefined },
}));

/** Neu einfärben: Office's entries for the document's theme colours. */
export function recolorGallery(t: Readonly<Record<string, string>>): GalleryItem[] {
  const names = ["Text 2", "Akzent 1", "Akzent 2", "Akzent 3", "Akzent 4", "Akzent 5", "Akzent 6"];
  const keys = ["dk2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6"];
  const lightBg = ["lt2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6"];
  const mix = (hex: string, k: number, toWhite: boolean) => {
    const v = [0, 2, 4].map((i) => parseInt(hex.slice(i, i + 2), 16));
    const out = v.map((x) => Math.round(toWhite ? x + (255 - x) * (1 - k) : x * k));
    return out.map((x) => x.toString(16).padStart(2, "0")).join("").toUpperCase();
  };
  return [
    { label: "Keine Neueinfärbung", format: { recolor: undefined, lum: undefined } },
    { label: "Graustufe", format: { recolor: { kind: "grayscale" } } },
    { label: "Sepia", format: { recolor: { kind: "sepia" } } },
    { label: "Ausgewaschen", format: { recolor: { kind: "washout" } } },
    { label: "Schwarzweiß: 25 %", format: { recolor: { kind: "bw", threshold: 0.25 } } },
    { label: "Schwarzweiß: 50 %", format: { recolor: { kind: "bw", threshold: 0.5 } } },
    { label: "Schwarzweiß: 75 %", format: { recolor: { kind: "bw", threshold: 0.75 } } },
    ...keys.map((k, i) => ({ label: `${names[i]}, dunkle Farbe`, format: { recolor: { kind: "duotone" as const, dark: "000000", light: mix(t[k] ?? "4472C4", 0.45, true) } } })),
    ...lightBg.map((k, i) => ({ label: `${names[i] === "Text 2" ? "Hintergrund 2" : names[i]}, helle Farbe`, format: { recolor: { kind: "duotone" as const, dark: mix(t[k] ?? "4472C4", 0.45, false), light: "FFFFFF" } } })),
  ];
}

/** Freistellen (Word) / Hintergrund entfernen (Excel): pixels connected to the edges and close to the edge colour become transparent. */
export async function removeBackground(bytes: Uint8Array, type: string): Promise<Uint8Array | null> {
  const url = URL.createObjectURL(new Blob([bytes as BlobPart], { type }));
  try {
    const img = await new Promise<HTMLImageElement>((resolve, reject) => { const i = new Image(); i.onload = () => resolve(i); i.onerror = reject; i.src = url; });
    const w = img.naturalWidth, hgt = img.naturalHeight;
    const cv = document.createElement("canvas");
    cv.width = w; cv.height = hgt;
    const g = cv.getContext("2d")!;
    g.drawImage(img, 0, 0);
    const data = g.getImageData(0, 0, w, hgt);
    const px = data.data;
    // The background colour: the average of the border pixels.
    let r = 0, gg = 0, b = 0, n = 0;
    for (let x = 0; x < w; x++) for (const y of [0, hgt - 1]) { const i = (y * w + x) * 4; r += px[i]; gg += px[i + 1]; b += px[i + 2]; n++; }
    for (let y = 0; y < hgt; y++) for (const x of [0, w - 1]) { const i = (y * w + x) * 4; r += px[i]; gg += px[i + 1]; b += px[i + 2]; n++; }
    r /= n; gg /= n; b /= n;
    const near = (i: number) => Math.abs(px[i] - r) + Math.abs(px[i + 1] - gg) + Math.abs(px[i + 2] - b) < 90;
    const seen = new Uint8Array(w * hgt);
    const stack: number[] = [];
    for (let x = 0; x < w; x++) { stack.push(x, (hgt - 1) * w + x); }
    for (let y = 0; y < hgt; y++) { stack.push(y * w, y * w + w - 1); }
    while (stack.length) {
      const p = stack.pop()!;
      if (seen[p]) continue;
      seen[p] = 1;
      if (!near(p * 4)) continue;
      px[p * 4 + 3] = 0;
      const x = p % w, y = (p / w) | 0;
      if (x > 0) stack.push(p - 1);
      if (x < w - 1) stack.push(p + 1);
      if (y > 0) stack.push(p - w);
      if (y < hgt - 1) stack.push(p + w);
    }
    g.putImageData(data, 0, 0);
    const blob = await new Promise<Blob | null>((resolve) => cv.toBlob(resolve, "image/png"));
    return blob ? new Uint8Array(await blob.arrayBuffer()) : null;
  } catch {
    return null;
  } finally {
    URL.revokeObjectURL(url);
  }
}
