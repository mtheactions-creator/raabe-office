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
  { title: "Dezent", items: [["Ohne", "none"], ["Morphen", "morph"], ["Verblassen", "fade"], ["Schieben", "push"], ["Wischen", "wipe"], ["Teilen", "split"], ["Enthüllen", "reveal"], ["Schnitt", "cut"], ["Zufällige Balken", "randomBar"], ["Form", "circle"], ["Aufdecken", "uncover"], ["Bedecken", "cover"], ["Blitz", "flash"]] },
  { title: "Spektakulär", items: [["Fallen", "fall"], ["Vorhänge", "curtains"], ["Zerknittern", "crush"], ["Zerreißen", "peelOff"], ["Umblättern", "pageCurl"], ["Flugzeug", "airplane"], ["Origami", "origami"], ["Auflösen", "dissolve"], ["Schachbrett", "checker"], ["Jalousie", "blinds"], ["Uhr", "wheel"], ["Welligkeit", "ripple"], ["Wabe", "honeycomb"], ["Glitzer", "glitter"], ["Strudel", "vortex"], ["Zersplittern", "shred"], ["Wechseln", "switch"], ["Kippen", "flip"], ["Galerie", "gallery"], ["Würfel", "cube"], ["Türen", "doors"], ["Feld", "box"], ["Kamm", "comb"], ["Zoom", "zoom"], ["Zufall", "random"]] },
  { title: "Dynamischer Inhalt", items: [["Verschieben", "pan"], ["Riesenrad", "ferris"], ["Förderband", "conveyor"], ["Drehen", "rotate"], ["Fenster", "window"], ["Umlaufbahn", "orbit"], ["Durchfliegen", "flythrough"]] },
];

/** Elements of the p14 namespace (the newer transitions); the rest are p: elements. */
const P14_KINDS = new Set(["vortex", "switch", "flip", "ripple", "honeycomb", "prism", "doors", "window", "ferris", "gallery", "conveyor", "pan", "glitter", "warp", "flythrough", "flash", "shred", "reveal", "wheelReverse"]);
/** Kinds PowerPoint writes with p15 (prstTrans) and their preset names. */
const PRST = new Map([["fall", "fallOver"], ["curtains", "curtains"], ["crush", "crush"], ["peelOff", "peelOff"], ["pageCurl", "pageCurlDouble"], ["airplane", "airplane"], ["origami", "origami"]]);

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
export function readTransition(root: Element): { kind: string; dur: number; advClick: boolean; advTm?: number; dir?: string } | null {
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
  (wrap ?? el)?.remove();
  if (kind === "none") return;
  const xml = transitionXml(kind, dur, { advClick: cur?.advClick, advTm: cur?.advTm });
  // Schema order: after clrMapOvr, before timing.
  const timing = kid(root, "timing") ?? kid(root, "extLst") ?? null;
  root.insertBefore(fragment(root.ownerDocument, xml)[0], timing);
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
