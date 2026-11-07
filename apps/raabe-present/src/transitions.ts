// Übergänge (D-070): the transition of a slide written as PowerPoint 2010+ writes it (an
// mc:AlternateContent with p14:dur for the duration, and a plain p:transition fallback), read back
// for the show. The gallery's names are PowerPoint's; how each is drawn in the show is our own
// (CSS animations), measured only for duration and order.

import { PptxDom } from "@raabe/ooxml";
import type { PresentApp } from "./app";

const { kid, kids, num, fragment, NS_P } = PptxDom;
const NS_MC = "http://schemas.openxmlformats.org/markup-compatibility/2006";
const NS_P14 = "http://schemas.microsoft.com/office/powerpoint/2010/main";

export const TRANSITION_GROUPS: { title: string; items: [string, string][] }[] = [
  { title: "Dezent", items: [["Ohne", "none"], ["Morphen", "morph"], ["Verblassen", "fade"], ["Schieben", "push"], ["Wischen", "wipe"], ["Teilen", "split"], ["Anzeigen", "reveal"], ["Schnitt", "cut"],
    ["Zufällige Balken", "randomBar"], ["Form", "circle"], ["Aufdecken", "uncover"], ["Bedecken", "cover"], ["Blitz", "flash"]] },
  { title: "Spektakulär", items: [["Umfallen", "fall"], ["Verhängen", "drape"], ["Vorhänge", "curtains"], ["Verwehen", "wind"], ["Wegziehen", "prestige"], ["Zersplittern", "fracture"], ["Zerknüllen", "crush"],
    ["Blatt abziehen", "peelOff"], ["Seite einrollen", "pageCurl"], ["Papierflieger", "airplane"], ["Origami", "origami"], ["Auflösen", "dissolve"], ["Schachbrett", "checker"], ["Jalousie", "blinds"],
    ["Uhr", "wheel"], ["Wellen", "ripple"], ["Bienenwabe", "honeycomb"], ["Glitzern", "glitter"], ["Wirbel", "vortex"], ["Fetzen", "shred"], ["Wechseln", "switch"], ["Spiegeln", "flip"],
    ["Katalog", "gallery"], ["Würfel", "cube"], ["Türen", "doors"], ["Kästchen", "box"], ["Kamm", "comb"], ["Zoom", "zoom"], ["Zufällig", "random"]] },
  { title: "Dynamischer Inhalt", items: [["Schwenken", "pan"], ["Riesenrad", "ferris"], ["Förderband", "conveyor"], ["Drehen", "prism"], ["Fenster", "window"], ["Orbit", "orbit"], ["Hindurchfliegen", "flythrough"]] },
];

/** Elements of the p14 namespace (the newer transitions); the rest are p: elements. */
const P14_KINDS = new Set(["vortex", "switch", "flip", "ripple", "honeycomb", "prism", "doors", "window", "ferris", "gallery", "conveyor", "pan", "glitter", "warp", "flythrough", "flash", "shred", "reveal", "wheelReverse"]);
/** Kinds PowerPoint writes with p15 (prstTrans) and their preset names. */
const PRST = new Map([["fall", "fallOver"], ["drape", "drape"], ["curtains", "curtains"], ["wind", "wind"], ["prestige", "prestige"], ["fracture", "fracture"], ["crush", "crush"], ["peelOff", "peelOff"], ["pageCurl", "pageCurlDouble"], ["airplane", "airplane"], ["origami", "origami"], ["orbit", "orbit"]]);

/** The slide's transition element (inside an AlternateContent's Choice when there is one). */
export function transitionOf(root: Element): { el: Element | null; wrap: Element | null } {
  for (const ac of kids(root, "AlternateContent")) {
    const choice = kid(ac, "Choice");
    const tr = kid(choice, "transition");
    if (tr) return { el: tr, wrap: ac };
  }
  return { el: kid(root, "transition") ?? null, wrap: null };
}

/** The kind of transition (the first child element's local name) and its duration in ms. */
export function readTransition(root: Element): { kind: string; dur: number; advClick: boolean; advTm?: number; dir?: string; thruBlk?: boolean } | null {
  const { el } = transitionOf(root);
  if (!el) return null;
  const child = kids(el).find((c) => c.localName !== "sndAc" && c.localName !== "extLst");
  const dur = num(el, "p14:dur") ?? Number(el.getAttributeNS(NS_P14, "dur") || NaN);
  const spd = el.getAttribute("spd") ?? "fast";
  return {
    kind: child?.localName === "prstTrans" ? child.getAttribute("prst") ?? "fade" : child?.localName ?? "cut",
    dur: Number.isFinite(dur) ? dur : spd === "slow" ? 1000 : spd === "med" ? 750 : 500,
    advClick: el.getAttribute("advClick") !== "0",
    advTm: num(el, "advTm"),
    dir: child?.getAttribute("dir") ?? undefined,
    // Verblassen > Durch Schwarz (PowerPoint writes thruBlk="1").
    thruBlk: child?.getAttribute("thruBlk") === "1" || child?.getAttribute("thruBlk") === "true" || undefined,
  };
}

const DEFAULT_DUR: Record<string, number> = { fade: 700, push: 1000, wipe: 1000, split: 1500, reveal: 1600, cut: 100, randomBar: 1000, circle: 2000, uncover: 1000, cover: 1000, flash: 1000, morph: 2000 };

function transitionXml(kind: string, dur: number, extra: { advClick?: boolean; advTm?: number }): string {
  const spd = dur >= 1000 ? "slow" : dur >= 750 ? "med" : "fast";
  const attrs = `${extra.advClick === false ? ' advClick="0"' : ""}${extra.advTm !== undefined ? ` advTm="${extra.advTm}"` : ""}`;
  if (kind === "none") return "";
  const inner = PRST.has(kind) ? `<p15:prstTrans xmlns:p15="http://schemas.microsoft.com/office/powerpoint/2012/main" prst="${PRST.get(kind)}"/>`
    : P14_KINDS.has(kind) ? `<p14:${kind}/>` : kind === "circle" ? "<p:circle/>" : `<p:${kind}/>`;
  return `<mc:AlternateContent xmlns:mc="${NS_MC}"><mc:Choice xmlns:p14="${NS_P14}" Requires="p14"><p:transition spd="${spd}" p14:dur="${Math.round(dur)}"${attrs}>${inner}</p:transition></mc:Choice>` +
    `<mc:Fallback><p:transition spd="${spd}"${attrs}><p:fade/></p:transition></mc:Fallback></mc:AlternateContent>`;
}

/** Writes `kind` for the selected slides (PowerPoint: the slides selected in the panel). */
export function setTransition(app: PresentApp, kind: string): void {
  const deck = app.deck;
  const paths = [...app.panel.selected].map((i) => deck.slides[i].path);
  deck.change("Übergang", paths, () => {
    for (const p of paths) writeTransition(deck.pptx.doc(p).documentElement, kind, DEFAULT_DUR[kind] ?? 1000);
  });
}

function writeTransition(root: Element, kind: string, dur: number): void {
  const cur = readTransition(root);
  const { el, wrap } = transitionOf(root);
  // The transition's sound stays when the effect changes (PowerPoint).
  const sound = kid(el, "sndAc");
  (wrap ?? el)?.remove();
  if (kind === "none") return;
  const xml = transitionXml(kind, dur, { advClick: cur?.advClick, advTm: cur?.advTm });
  // Schema order: after clrMapOvr, before timing.
  const timing = kid(root, "timing") ?? kid(root, "extLst") ?? null;
  const added = fragment(root.ownerDocument, xml)[0];
  root.insertBefore(added, timing);
  if (sound) for (const tr of [added, ...Array.from(added.getElementsByTagNameNS(NS_P, "transition"))]) if (tr.localName === "transition") tr.append(sound.cloneNode(true));
}

export function setTransitionDuration(app: PresentApp, seconds: number): void {
  const deck = app.deck;
  const paths = [...app.panel.selected].map((i) => deck.slides[i].path);
  deck.change("Dauer", paths, () => {
    for (const p of paths) {
      const root = deck.pptx.doc(p).documentElement;
      const t = readTransition(root);
      if (t) writeTransition(root, t.kind, seconds * 1000);
    }
  });
}

export function setAdvanceAfter(app: PresentApp, seconds: number): void {
  const deck = app.deck;
  const paths = [...app.panel.selected].map((i) => deck.slides[i].path);
  deck.change("Nach", paths, () => {
    for (const p of paths) {
      const root = deck.pptx.doc(p).documentElement;
      if (!transitionOf(root).el) root.insertBefore(fragment(root.ownerDocument, '<p:transition/>')[0], kid(root, "timing") ?? kid(root, "extLst") ?? null);
      for (const tr of [transitionOf(root).el, ...Array.from(root.getElementsByTagNameNS(NS_P, "transition"))]) {
        if (!tr) continue;
        tr.setAttribute("advTm", String(Math.round(seconds * 1000)));
      }
    }
  });
}

/** Effektoptionen of a transition: its directions as PowerPoint names them. */
export function transitionOptions(app: PresentApp): { label: string; run: () => void; checked?: boolean }[] {
  const p = app.slidePath();
  if (!p) return [];
  const root = app.deck.pptx.doc(p).documentElement;
  const t = readTransition(root);
  if (!t) return [];
  const dirs: Record<string, [string, string][]> = {
    push: [["Von unten", "u"], ["Von links", "r"], ["Von rechts", "l"], ["Von oben", "d"]],
    wipe: [["Von rechts", "l"], ["Von oben", "d"], ["Von links", "r"], ["Von unten", "u"]],
    cover: [["Von rechts", "l"], ["Von oben", "d"], ["Von links", "r"], ["Von unten", "u"]],
    uncover: [["Von rechts", "l"], ["Von oben", "d"], ["Von links", "r"], ["Von unten", "u"]],
    split: [["Vertikal nach innen", "in"], ["Vertikal nach außen", "out"]],
  };
  const list = dirs[t.kind];
  if (!list) return [];
  return list.map(([label, dir]) => ({
    label, checked: (t.dir ?? (t.kind === "split" ? "out" : list[0][1])) === dir,
    run: () => {
      const paths = [...app.panel.selected].map((i) => app.deck.slides[i].path);
      app.deck.change("Effektoptionen", paths, () => {
        for (const path of paths) {
          const r = app.deck.pptx.doc(path).documentElement;
          for (const tr of Array.from(r.getElementsByTagNameNS(NS_P, "transition"))) {
            const child = kids(tr).find((c) => c.localName === t.kind);
            if (child) child.setAttribute("dir", dir);
          }
        }
      });
    },
  }));
}

// ----- Sound (D-072) ---------------------------------------------------------------------------------------
// PowerPoint's entries without its own sound files (Microsoft content): [Ohne Sound], [Vorherigen Sound
// anhalten], Anderer Sound... (a WAV file of the user's, embedded) and Bis zum nächsten Sound wiederholen,
// stored as PowerPoint stores them (p:sndAc in every p:transition of the slide).

const RT_AUDIO = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/audio";

function setSoundXml(app: PresentApp, label: string, make: (path: string, tr: Element) => string | null): void {
  const deck = app.deck;
  const paths = [...app.panel.selected].map((i) => deck.slides[i].path);
  deck.change(label, paths, () => {
    for (const p of paths) {
      const root = deck.pptx.doc(p).documentElement;
      if (!transitionOf(root).el) root.insertBefore(fragment(root.ownerDocument, "<p:transition/>")[0], kid(root, "timing") ?? kid(root, "extLst") ?? null);
      for (const tr of [transitionOf(root).el, ...Array.from(root.getElementsByTagNameNS(NS_P, "transition"))]) {
        if (!tr) continue;
        kid(tr, "sndAc")?.remove();
        const xml = make(p, tr);
        if (xml) tr.append(...fragment(tr.ownerDocument, xml));
      }
    }
  });
}

export function soundMenu(app: PresentApp): { label: string; run: () => void; checked?: boolean }[] {
  const root = app.slidePath() ? app.deck.pptx.doc(app.slidePath()!).documentElement : null;
  const snd = root ? kid(transitionOf(root).el, "sndAc") : undefined;
  const st = kid(snd, "stSnd");
  return [
    { label: "[Ohne Sound]", checked: !snd, run: () => setSoundXml(app, "Sound", () => null) },
    { label: "[Vorherigen Sound anhalten]", checked: !!kid(snd, "endSnd"), run: () => setSoundXml(app, "Sound", () => "<p:sndAc><p:endSnd/></p:sndAc>") },
    { label: "Anderer Sound...", run: () => void otherSound(app) },
    { label: "Bis zum nächsten Sound wiederholen", checked: st?.getAttribute("loop") === "1", run: () => {
      if (!st) return;
      const on = st.getAttribute("loop") !== "1";
      const deck = app.deck;
      const paths = [...app.panel.selected].map((i) => deck.slides[i].path);
      deck.change("Sound", paths, () => {
        for (const p of paths) for (const tr of Array.from(deck.pptx.doc(p).getElementsByTagNameNS(NS_P, "transition"))) {
          const s1 = kid(kid(tr, "sndAc"), "stSnd");
          if (s1) { if (on) s1.setAttribute("loop", "1"); else s1.removeAttribute("loop"); }
        }
      });
    } },
  ];
}

async function otherSound(app: PresentApp): Promise<void> {
  const path = await app.platform.openAnyDialog?.("Audio hinzufügen", [{ name: "Audiodateien", extensions: ["wav"] }]);
  if (!path) return;
  const bytes = await app.platform.readFile(path);
  const name = path.split(/[\\/]/).pop() ?? "sound.wav";
  const pptx = app.deck.pptx;
  const media = pptx.freeName("ppt/media/audio#.wav");
  pptx.put(media, bytes, "audio/wav");
  setSoundXml(app, "Sound", (p) => {
    const rid = pptx.link(p, RT_AUDIO, media);
    return `<p:sndAc><p:stSnd><p:snd r:embed="${rid}" name="${PptxDom.esc(name)}"/></p:stSnd></p:sndAc>`;
  });
}

/** The sound a slide's transition plays: the media part, or "stop" for [Vorherigen Sound anhalten]. */
export function transitionSound(app: PresentApp, path: string): { media?: string; stop?: boolean; loop?: boolean } | null {
  const root = app.deck.pptx.doc(path).documentElement;
  const snd = kid(transitionOf(root).el, "sndAc");
  if (!snd) return null;
  if (kid(snd, "endSnd")) return { stop: true };
  const st = kid(snd, "stSnd"), s = kid(st, "snd");
  const rid = s?.getAttributeNS(PptxDom.NS_R, "embed");
  const media = rid ? app.deck.pptx.target(path, rid) : undefined;
  return media ? { media, loop: st?.getAttribute("loop") === "1" } : null;
}

// ----- Anzeigedauer: Bei Mausklick, Nach, Auf alle anwenden (D-072) ----------------------------------------

function selectedPaths(app: PresentApp): string[] {
  return [...app.panel.selected].map((i) => app.deck.slides[i]?.path).filter((p): p is string => !!p);
}

function eachTransition(app: PresentApp, label: string, fn: (tr: Element) => void): void {
  const deck = app.deck, paths = selectedPaths(app);
  deck.change(label, paths, () => {
    for (const p of paths) {
      const root = deck.pptx.doc(p).documentElement;
      if (!transitionOf(root).el) root.insertBefore(fragment(root.ownerDocument, "<p:transition/>")[0], kid(root, "timing") ?? kid(root, "extLst") ?? null);
      for (const tr of Array.from(root.getElementsByTagNameNS(NS_P, "transition"))) fn(tr);
    }
  });
}

export function advanceOnClick(app: PresentApp): boolean {
  const p = app.slidePath();
  return !p || transitionOf(app.deck.pptx.doc(p).documentElement).el?.getAttribute("advClick") !== "0";
}
export function setAdvanceOnClick(app: PresentApp, on: boolean): void {
  eachTransition(app, "Bei Mausklick", (tr) => { if (on) tr.removeAttribute("advClick"); else tr.setAttribute("advClick", "0"); });
}
export function advanceAfter(app: PresentApp): boolean {
  const p = app.slidePath();
  return !!p && !!transitionOf(app.deck.pptx.doc(p).documentElement).el?.hasAttribute("advTm");
}
export function setAdvanceAfterOn(app: PresentApp, on: boolean): void {
  eachTransition(app, "Nach", (tr) => { if (on) { if (!tr.hasAttribute("advTm")) tr.setAttribute("advTm", "0"); } else tr.removeAttribute("advTm"); });
}

/** Auf alle anwenden: the current slide's transition (effect, duration, sound, Anzeigedauer) on every slide. */
export function applyToAll(app: PresentApp): void {
  const deck = app.deck, cur = app.slidePath();
  if (!cur) return;
  const src = transitionOf(deck.pptx.doc(cur).documentElement);
  const node = src.wrap ?? src.el;
  const paths = deck.slides.map((s) => s.path).filter((p) => p !== cur);
  deck.change("Auf alle anwenden", paths, () => {
    for (const p of paths) {
      const root = deck.pptx.doc(p).documentElement;
      const { el, wrap } = transitionOf(root);
      (wrap ?? el)?.remove();
      if (node) root.insertBefore(root.ownerDocument.importNode(node, true), kid(root, "timing") ?? kid(root, "extLst") ?? null);
    }
  });
}
