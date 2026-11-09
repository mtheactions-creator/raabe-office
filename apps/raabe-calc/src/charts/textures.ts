// The Textur gallery of Bild- oder Texturfüllung (D-155). Excel offers 24 textures under these names; its bitmaps
// are Office's own, so ours are drawn here: a small tile per texture from a few colours and a kind of grain
// (paper fibres, woven threads, marble veins, wood rings, drops …), seeded so the same name always gives the same
// tile. They are tiled as Excel tiles its textures (a:tile, 100 %) and saved into the file as a picture.

/** Excel's texture names in the gallery's order. */
export const TEXTURES = [
  "Papyrus", "Leinen", "Blaue Baumwolle", "Gewebe", "Wassertropfen", "Papiertüte", "Fischfossil", "Sand", "Grüner Marmor", "Weißer Marmor",
  "Brauner Marmor", "Granit", "Zeitungspapier", "Recyclingpapier", "Pergament", "Briefpapier", "Blaues Seidenpapier", "Rosa Seidenpapier",
  "Lila Netz", "Blumenstrauß", "Kork", "Walnuss", "Eiche", "Nussbaum",
] as const;

type Grain = "fibre" | "weave" | "drops" | "speck" | "marble" | "wood" | "net" | "flowers" | "fossil";
const LOOK: Record<string, { base: string; dark: string; grain: Grain; size?: number }> = {
  Papyrus: { base: "#E8DDB5", dark: "#B9A877", grain: "fibre" },
  Leinen: { base: "#E9E2D3", dark: "#C4B9A3", grain: "weave" },
  "Blaue Baumwolle": { base: "#AFC6E3", dark: "#7E9CC4", grain: "weave" },
  Gewebe: { base: "#C9C0AE", dark: "#8E8470", grain: "weave", size: 64 },
  Wassertropfen: { base: "#9DB4C9", dark: "#5E7C99", grain: "drops" },
  Papiertüte: { base: "#C8A97A", dark: "#9C7C4E", grain: "fibre" },
  Fischfossil: { base: "#D9CDB3", dark: "#9C8B6B", grain: "fossil" },
  Sand: { base: "#E2D2A9", dark: "#B49F70", grain: "speck" },
  "Grüner Marmor": { base: "#5E8C6E", dark: "#1F4A33", grain: "marble" },
  "Weißer Marmor": { base: "#EEEEEA", dark: "#A9A9A4", grain: "marble" },
  "Brauner Marmor": { base: "#8E6B4D", dark: "#4B3220", grain: "marble" },
  Granit: { base: "#B4B0AA", dark: "#5C5852", grain: "speck" },
  Zeitungspapier: { base: "#DCD8CC", dark: "#9C978B", grain: "speck" },
  Recyclingpapier: { base: "#D7CCB4", dark: "#A08F70", grain: "fibre" },
  Pergament: { base: "#EFE3C2", dark: "#C9B385", grain: "fibre" },
  Briefpapier: { base: "#F2EEE6", dark: "#CFC6B5", grain: "fibre" },
  "Blaues Seidenpapier": { base: "#C7D6EC", dark: "#94ADD4", grain: "fibre" },
  "Rosa Seidenpapier": { base: "#F1CED8", dark: "#D69BAD", grain: "fibre" },
  "Lila Netz": { base: "#C9B5DA", dark: "#7D5C9C", grain: "net" },
  Blumenstrauß: { base: "#E7D9C8", dark: "#B0606F", grain: "flowers" },
  Kork: { base: "#C49A69", dark: "#7C5531", grain: "speck" },
  Walnuss: { base: "#7A5235", dark: "#4A2E1A", grain: "wood" },
  Eiche: { base: "#C79E68", dark: "#8F6A3C", grain: "wood" },
  Nussbaum: { base: "#9C6B43", dark: "#5E3C22", grain: "wood" },
};

/** A small seeded random generator (mulberry32). */
function rng(seed: number): () => number {
  let a = seed >>> 0;
  return () => {
    a = (a + 0x6d2b79f5) >>> 0;
    let t = a;
    t = Math.imul(t ^ (t >>> 15), t | 1);
    t ^= t + Math.imul(t ^ (t >>> 7), t | 61);
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
}

const seedOf = (name: string) => [...name].reduce((h, c) => Math.imul(h ^ c.charCodeAt(0), 16777619) >>> 0, 2166136261);

function draw(ctx: CanvasRenderingContext2D, n: number, look: (typeof LOOK)[string], rnd: () => number): void {
  const { base, dark, grain } = look;
  ctx.fillStyle = base;
  ctx.fillRect(0, 0, n, n);
  ctx.strokeStyle = dark;
  ctx.fillStyle = dark;
  const line = (x0: number, y0: number, x1: number, y1: number, w: number, a: number) => {
    ctx.globalAlpha = a; ctx.lineWidth = w; ctx.beginPath(); ctx.moveTo(x0, y0); ctx.lineTo(x1, y1); ctx.stroke();
  };
  switch (grain) {
    case "fibre":
      for (let i = 0; i < n * 1.5; i++) { const x = rnd() * n, y = rnd() * n, l = 4 + rnd() * 14, a = rnd() * Math.PI; line(x, y, x + Math.cos(a) * l, y + Math.sin(a) * l, 0.6, 0.12 + rnd() * 0.15); }
      break;
    case "weave":
      for (let y = 0; y < n; y += 2) line(0, y + 0.5, n, y + 0.5, 1, 0.12 + rnd() * 0.12);
      for (let x = 0; x < n; x += 2) line(x + 0.5, 0, x + 0.5, n, 1, 0.1 + rnd() * 0.12);
      break;
    case "drops":
      for (let i = 0; i < 40; i++) {
        const x = rnd() * n, y = rnd() * n, r = 2 + rnd() * 6;
        ctx.globalAlpha = 0.35; ctx.beginPath(); ctx.arc(x, y, r, 0, Math.PI * 2); ctx.fill();
        ctx.globalAlpha = 0.6; ctx.fillStyle = "#FFFFFF"; ctx.beginPath(); ctx.arc(x - r / 3, y - r / 3, r / 4, 0, Math.PI * 2); ctx.fill(); ctx.fillStyle = dark;
      }
      break;
    case "speck":
      for (let i = 0; i < n * n / 6; i++) { ctx.globalAlpha = 0.1 + rnd() * 0.4; ctx.fillRect(rnd() * n, rnd() * n, 1 + rnd() * 1.5, 1 + rnd() * 1.5); }
      break;
    case "marble":
      for (let k = 0; k < 7; k++) {
        let x = rnd() * n, y = 0;
        ctx.globalAlpha = 0.25 + rnd() * 0.3; ctx.lineWidth = 0.6 + rnd() * 1.6; ctx.beginPath(); ctx.moveTo(x, y);
        while (y < n) { y += 4; x += (rnd() - 0.5) * 8; ctx.lineTo(x, y); }
        ctx.stroke();
      }
      break;
    case "wood":
      for (let y = 0; y < n; y += 3 + rnd() * 4) {
        ctx.globalAlpha = 0.2 + rnd() * 0.35; ctx.lineWidth = 0.8 + rnd() * 1.5; ctx.beginPath();
        for (let x = 0; x <= n; x += 4) ctx.lineTo(x, y + Math.sin((x / n) * Math.PI * 2 + y) * 1.5);
        ctx.stroke();
      }
      break;
    case "net":
      for (let i = -n; i < n; i += 8) { line(i, 0, i + n, n, 1, 0.5); line(i + n, 0, i, n, 1, 0.5); }
      break;
    case "flowers":
      for (let i = 0; i < 9; i++) {
        const x = rnd() * n, y = rnd() * n;
        for (let p = 0; p < 5; p++) { const a = (p / 5) * Math.PI * 2; ctx.globalAlpha = 0.55; ctx.beginPath(); ctx.arc(x + Math.cos(a) * 4, y + Math.sin(a) * 4, 3, 0, Math.PI * 2); ctx.fill(); }
      }
      break;
    case "fossil":
      for (let i = 0; i < n * n / 10; i++) { ctx.globalAlpha = 0.08 + rnd() * 0.2; ctx.fillRect(rnd() * n, rnd() * n, 1.5, 1.5); }
      ctx.globalAlpha = 0.45; ctx.lineWidth = 1.2; ctx.beginPath();
      ctx.moveTo(n * 0.2, n * 0.5); ctx.quadraticCurveTo(n * 0.5, n * 0.3, n * 0.8, n * 0.5); ctx.quadraticCurveTo(n * 0.5, n * 0.7, n * 0.2, n * 0.5); ctx.stroke();
      for (let k = 1; k < 6; k++) line(n * (0.2 + k * 0.1), n * 0.42, n * (0.2 + k * 0.1), n * 0.58, 0.8, 0.4);
      break;
  }
  ctx.globalAlpha = 1;
}

const made = new Map<string, string>();

/** The tile of a texture as a PNG data URL, or undefined (no such name, or no canvas). */
export function textureSrc(name: string | undefined): string | undefined {
  if (!name || !LOOK[name]) return undefined;
  const hit = made.get(name);
  if (hit) return hit;
  if (typeof document === "undefined") return undefined;
  const look = LOOK[name];
  const n = look.size ?? 96;
  const c = document.createElement("canvas");
  c.width = c.height = n;
  const ctx = c.getContext("2d");
  if (!ctx) return undefined;
  draw(ctx, n, look, rng(seedOf(name)));
  const src = c.toDataURL("image/png");
  made.set(name, src);
  return src;
}
