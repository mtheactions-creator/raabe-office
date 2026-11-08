// Fonts whose name the browser does not find (D-089). Windows lists some fonts under a family of their own in GDI
// ("Aptos Narrow", "Aptos Display", "Calibri Light", "Segoe UI Semibold") but under a shared typographic family in
// DirectWrite ("Aptos" with the width Narrow, "Calibri" with the weight Light), and WebView2 only knows the latter:
// CSS "Aptos Narrow" fell back to Arial, so Excel 365's standard font, Word's headings and PowerPoint's titles were
// drawn and printed in Arial. For such a name the shared rule below adds @font-face aliases that take the installed
// faces by their full names (local("Aptos Narrow"), local("Aptos Narrow Bold") …), which works for drawing, measuring
// on canvas and printing to PDF. Names the browser finds (and fonts that are not installed) stay untouched.

const checked = new Map<string, boolean>();
const listeners = new Set<() => void>();
let style: HTMLStyleElement | null = null;
let probe: CanvasRenderingContext2D | null = null;

const GENERIC = /^(serif|sans-serif|monospace|cursive|fantasy|system-ui|inherit|initial|unset)$/i;

/** Office's cloud fonts (Aptos …): not Windows fonts, so the native side gives their files (D-089). */
export interface CloudFonts {
  list(): Promise<{ families: string[]; weight: number; italic: boolean; id: string }[]>;
  bytes(id: string): Promise<ArrayBuffer>;
}

let cloud: CloudFonts | null = null;
let cloudList: Promise<{ families: string[]; weight: number; italic: boolean; id: string }[]> | null = null;

/** Gives the font loader Office's cloud fonts (the apps pass their platform's commands). */
export function setCloudFonts(c: CloudFonts): void {
  cloud = c;
  cloudList = c.list().catch(() => []);
}

/** Registers a family from the cloud fonts with FontFace (bytes), when the cache has it. */
async function fromCloud(name: string): Promise<boolean> {
  if (!cloud || !cloudList) return false;
  const l = name.toLowerCase();
  const faces = (await cloudList).filter((f) => f.families.some((x) => x.toLowerCase() === l));
  if (!faces.length) return false;
  await Promise.all(faces.map(async (f) => {
    try {
      const face = new FontFace(name, await cloud!.bytes(f.id), { weight: String(f.weight), style: f.italic ? "italic" : "normal" });
      await face.load();
      document.fonts.add(face);
    } catch { /* a file that does not load: the other faces still count */ }
  }));
  return true;
}

/** Called when an alias has loaded (layouts measured with the fallback should be redone). */
export function onFontAlias(f: () => void): () => void {
  listeners.add(f);
  return () => listeners.delete(f);
}

function widthIn(family: string, fallback: string): number {
  probe ??= document.createElement("canvas").getContext("2d");
  if (!probe) return 0;
  probe.font = `72px ${family}${fallback ? `, ${fallback}` : ""}`;
  return probe.measureText("mmmmmmmmmmlli1WQ@#äö").width;
}

/** Whether the browser draws the family itself (it differs from both generic fallbacks). */
function found(family: string): boolean {
  const q = `"${family.replace(/"/g, "")}"`;
  return widthIn(q, "monospace") !== widthIn("monospace", "") || widthIn(q, "serif") !== widthIn("serif", "");
}

/** Makes sure a family name can be used in CSS: adds the full-name aliases when the browser does not know the name. */
export function ensureFontFamily(family: string | undefined): void {
  if (!family || typeof document === "undefined" || GENERIC.test(family)) return;
  const key = family.toLowerCase();
  if (checked.has(key)) return;
  checked.set(key, true);
  let known = true;
  try { known = found(family); } catch { return; }
  if (known) return;
  const name = family.replace(/["\\]/g, "");
  // Office's own cloud fonts first (Aptos …), else the installed faces by their full names.
  void fromCloud(name).then((ok) => {
    if (ok) for (const f of listeners) f();
    else aliasByFullName(name);
  });
}

function aliasByFullName(name: string): void {
  const faces: [string, string, string][] = [["", "400", "normal"], [" Bold", "700", "normal"], [" Italic", "400", "italic"], [" Bold Italic", "700", "italic"]];
  const rules = faces.map(([suffix, weight, st]) =>
    `@font-face{font-family:"${name}";src:local("${name}${suffix}")${suffix ? "" : `,local("${name} Regular")`};font-weight:${weight};font-style:${st}}`);
  style ??= document.head.appendChild(Object.assign(document.createElement("style"), { id: "raabe-font-alias" }));
  style.append(document.createTextNode(rules.join("\n") + "\n"));
  const loads = faces.map(([, weight, st]) => document.fonts.load(`${st === "italic" ? "italic " : ""}${weight} 16px "${name}"`).catch(() => []));
  void Promise.all(loads).then(() => { for (const f of listeners) f(); });
}
