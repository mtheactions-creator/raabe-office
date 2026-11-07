// Formformat > Formeffekte and Texteffekte > Abschrägung, 3D-Drehung, Transformieren (D-072). The galleries'
// names are PowerPoint's (inventory); the effects are Office's presets as Word writes them (effect-presets.ts,
// recorded from Word, the same Office galleries). A shape's effects go into its spPr (effectLst, scene3d, sp3d in
// that order); a text's 3-D and its Transformieren into the text body's bodyPr, as PowerPoint writes them.

import type { MenuEntry } from "@raabe/shell";
import { PptxDom } from "@raabe/ooxml";
import type { PresentApp } from "./app";
import { BEVELS, EFFECT_PRESETS, GLOWS, REFLECTIONS, ROTATIONS_3D, SHADOWS, SOFT_EDGES, type PicturePreset } from "./effect-presets";
import * as S from "./shapeops";

const { kid, kids, make, fragment } = PptxDom;
const FX_ORDER = ["blur", "fillOverlay", "glow", "innerShdw", "outerShdw", "prstShdw", "reflection", "softEdge"];
const AFTER_FX = ["scene3d", "sp3d", "extLst"];

type Part = "shadow" | "reflection" | "glow" | "softEdge" | "bevel" | "rotation" | "preset";
const FX_TAGS: Record<string, string[]> = { shadow: ["outerShdw", "innerShdw", "prstShdw"], reflection: ["reflection"], glow: ["glow"], softEdge: ["softEdge"] };

/** The preset's elements, parsed (a wrapper holds its spPr children). */
function presetEls(doc: Document, p: PicturePreset): Element[] {
  const wrap = fragment(doc, `<p:spPr xmlns:p="${PptxDom.NS_P}" xmlns:a="${PptxDom.NS_A}">${p.xml}</p:spPr>`)[0];
  return wrap ? Array.from(wrap.children) : [];
}

/** Applies one part of a preset to a properties element (spPr, or bodyPr for text with `text`). */
function applyPart(pr: Element, part: Part, p: PicturePreset | null, text = false): void {
  const doc = pr.ownerDocument;
  const els = p ? presetEls(doc, p) : [];
  const pick = (name: string) => els.find((e) => e.localName === name);
  const insertBeforeAny = (el: Element, names: string[]) => pr.insertBefore(el, kids(pr).find((c) => names.includes(c.localName)) ?? null);
  if (part === "preset") {
    for (const n of ["effectLst", "effectDag", "scene3d", "sp3d"]) kid(pr, n)?.remove();
    for (const n of ["effectLst", "scene3d", "sp3d"]) { const e = pick(n); if (e) insertBeforeAny(e, n === "effectLst" ? AFTER_FX : n === "scene3d" ? ["sp3d", "extLst"] : ["extLst"]); }
    return;
  }
  if (part === "bevel" || part === "rotation") {
    // 3-D: scene3d (camera, light) and sp3d (bevel); a bevel keeps the camera and the other way round.
    const before = text ? ["flatTx", "extLst"] : ["extLst"];
    let scene = kid(pr, "scene3d"), sp3d = kid(pr, "sp3d");
    if (part === "rotation") {
      const cam = pick("scene3d");
      scene?.remove();
      if (cam) { pr.insertBefore(cam, sp3d ?? kids(pr).find((c) => before.includes(c.localName)) ?? null); scene = cam; }
      else if (sp3d) { const flat = fragment(doc, `<a:scene3d xmlns:a="${PptxDom.NS_A}"><a:camera prst="orthographicFront"/><a:lightRig rig="threePt" dir="t"/></a:scene3d>`)[0]; pr.insertBefore(flat, sp3d); }
    } else {
      const bevel = pick("sp3d")?.getElementsByTagNameNS(PptxDom.NS_A, "bevelT")[0];
      kid(sp3d, "bevelT")?.remove();
      if (bevel) {
        if (!scene) { scene = fragment(doc, `<a:scene3d xmlns:a="${PptxDom.NS_A}"><a:camera prst="orthographicFront"/><a:lightRig rig="threePt" dir="t"/></a:scene3d>`)[0]; pr.insertBefore(scene, sp3d ?? kids(pr).find((c) => before.includes(c.localName)) ?? null); }
        if (!sp3d) { sp3d = make(doc, "a:sp3d"); scene.after(sp3d); }
        sp3d.insertBefore(bevel, sp3d.firstChild);
      } else if (sp3d && !sp3d.children.length && !sp3d.attributes.length) sp3d.remove();
    }
    // Nothing 3-D left: no scene either.
    const s3 = kid(pr, "sp3d"), sc = kid(pr, "scene3d");
    if (sc && !s3 && kid(sc, "camera")?.getAttribute("prst") === "orthographicFront" && !kid(kid(sc, "camera"), "rot")) sc.remove();
    return;
  }
  let fx = kid(pr, "effectLst");
  if (!fx) { fx = make(doc, "a:effectLst"); insertBeforeAny(fx, AFTER_FX); }
  for (const t of FX_TAGS[part]) kid(fx, t)?.remove();
  const from = pick("effectLst");
  for (const e of kids(from).filter((x) => FX_TAGS[part].includes(x.localName))) fx.insertBefore(e, kids(fx).find((c) => FX_ORDER.indexOf(c.localName) > FX_ORDER.indexOf(e.localName)) ?? null);
}

function shapeFx(app: PresentApp, part: Part, p: PicturePreset | null, label: string): void {
  const views = app.editor.selectedViews.filter((v) => v.kind === "sp" || v.kind === "pic" || v.kind === "cxnSp");
  app.editor.edit(label, () => {
    for (const v of views) {
      const el = app.editor.shapeEl(v.id);
      const sp = el ? kid(el, "spPr") : undefined;
      if (sp) applyPart(sp, part, p);
    }
  });
}

/** Group headings as PowerPoint's galleries show them. */
function grouped(list: readonly (PicturePreset & { id: string })[], groups: [string, number][], run: (p: PicturePreset | null, label: string) => void, none: string, options: string): MenuEntry[] {
  const out: MenuEntry[] = [{ heading: none }, { label: list[0].label, run: () => run(null, list[0].label) }];
  let i = 1;
  for (const [head, n] of groups) {
    if (head) out.push({ heading: head });
    for (const p of list.slice(i, i + n)) out.push({ label: p.label, run: () => run(p, p.label) });
    i += n;
  }
  out.push("-", { label: options, run: () => {} , disabled: true });
  return out;
}

/** Formformat > Formeffekte, as PowerPoint's menu. */
export function shapeEffectsMenu(app: PresentApp): MenuEntry[] {
  const run = (part: Part) => (p: PicturePreset | null, label: string) => shapeFx(app, part, p, label);
  return [
    { label: "Voreinstellung", submenu: () => grouped(EFFECT_PRESETS, [["Voreinstellungen", 12]], run("preset"), "Keine Voreinstellungen", "3D-Optionen...") },
    { label: "Schatten", submenu: () => grouped(SHADOWS, [["Außen", 9], ["Innen", 9], ["Perspektive", 5]], run("shadow"), "Kein Schatten", "Schattenoptionen...") },
    { label: "Spiegelung", submenu: () => grouped(REFLECTIONS, [["Spiegelungsvarianten", 9]], run("reflection"), "Keine Spiegelung", "Spiegelungsoptionen...") },
    { label: "Leuchteffekt", submenu: () => grouped(GLOWS, [["Leuchteffektvarianten", 24]], run("glow"), "Kein Leuchteffekt", "Leuchteffektoptionen...") },
    { label: "Weiche Kanten", submenu: () => grouped(SOFT_EDGES, [["Varianten weicher Kanten", 6]], run("softEdge"), "Keine weichen Kanten", "Optionen für weiche Kanten...") },
    { label: "Abschrägung", submenu: () => grouped(BEVELS, [["", 12]], run("bevel"), "Keine Abschrägung", "3D-Optionen...") },
    { label: "3D-Drehung", submenu: () => grouped(ROTATIONS_3D, [["Parallel", 10], ["Perspektive", 11], ["Schräg", 4]], run("rotation"), "Keine Drehung", "3D-Drehungsoptionen...") },
  ];
}

// ----- text -----------------------------------------------------------------------------------------------------

function textBodyFx(app: PresentApp, label: string, fn: (bodyPr: Element) => void): void {
  const ids = app.editor.text ? [app.editor.text.id] : app.editor.sel;
  app.editor.edit(label, () => {
    for (const id of ids) {
      const el = app.editor.shapeEl(id);
      const bp = el ? S.bodyPrOf(el) : undefined;
      if (bp) fn(bp);
    }
  });
}

/** Texteffekte > Abschrägung and 3D-Drehung: on the text body (bodyPr), as PowerPoint writes text 3-D. */
export function textThreeD(app: PresentApp, kind: "bevel" | "rotation"): MenuEntry[] {
  const run = (p: PicturePreset | null, label: string) => textBodyFx(app, label, (bp) => applyPart(bp, kind, p, true));
  return kind === "bevel" ? grouped(BEVELS, [["", 12]], run, "Keine Abschrägung", "3D-Optionen...")
    : grouped(ROTATIONS_3D, [["Parallel", 10], ["Perspektive", 11], ["Schräg", 4]], run, "Keine Drehung", "3D-Drehungsoptionen...");
}

/** Texteffekte > Transformieren: PowerPoint's 40 text warps (a:prstTxWarp), with its names. */
const WARPS: [string, [string, string][]][] = [
  ["Folgepfad", [["Arch", "textArchUp"], ["Arch Down", "textArchDown"], ["Kreis", "textCircle"], ["Schaltfläche", "textButton"]]],
  ["Verzerrung", [["Rechteck", "textPlain"], ["Anhalten", "textStop"], ["Dreieck: nach oben", "textTriangle"], ["Dreieck: nach unten", "textTriangleInverted"],
    ["Chevron: nach oben", "textChevron"], ["Chevron: nach unten", "textChevronInverted"], ["Kreis: innen", "textRingInside"], ["Kreis: außen", "textRingOutside"],
    ["Bogen: verzerrt", "textArchUpPour"], ["Bogen: nach unten verzerrt", "textArchDownPour"], ["Verzerrter Kreis", "textCirclePour"], ["Verzerrte Schaltfläche", "textButtonPour"],
    ["Gekrümmt: nach oben", "textCurveUp"], ["Gekrümmt: nach unten", "textCurveDown"], ["Verzerrung nach oben", "textCanUp"], ["Verzerrung nach unten", "textCanDown"],
    ["Welle: nach unten", "textWave1"], ["Welle: nach oben", "textWave2"], ["Doppelte Welle: nach unten und oben", "textDoubleWave1"], ["Doppelte Welle: nach oben und unten", "textWave4"],
    ["Vergrößern", "textInflate"], ["Verkleinern", "textDeflate"], ["Vergrößern: nach unten", "textInflateBottom"], ["Verkleinern: nach unten", "textDeflateBottom"],
    ["Vergrößern: nach oben", "textInflateTop"], ["Verkleinern: nach oben", "textDeflateTop"], ["Verkleinern: nach oben, Vergrößern: nach unten", "textDeflateInflate"],
    ["Vergrößern: zur Mitte, Verkleinern: nach oben und nach unten", "textDeflateInflateDeflate"], ["Abnehmend: nach rechts", "textFadeRight"], ["Abnehmend: nach links", "textFadeLeft"],
    ["Abnehmend: nach oben", "textFadeUp"], ["Abnehmend: nach unten", "textFadeDown"], ["Schräge: nach oben", "textSlantUp"], ["Schräge: nach unten", "textSlantDown"],
    ["Überlappend: nach oben", "textCascadeUp"], ["Überlappend: nach unten", "textCascadeDown"]]],
];

export function transformMenu(app: PresentApp): MenuEntry[] {
  const set = (prst: string | null, label: string) => textBodyFx(app, label, (bp) => {
    kid(bp, "prstTxWarp")?.remove();
    if (prst) bp.insertBefore(fragment(bp.ownerDocument, `<a:prstTxWarp xmlns:a="${PptxDom.NS_A}" prst="${prst}"><a:avLst/></a:prstTxWarp>`)[0], bp.firstChild);
  });
  return [{ heading: "Keine Transformation" }, { label: "Keine", run: () => set(null, "Transformieren") },
    ...WARPS.flatMap(([head, list]) => [{ heading: head } as MenuEntry, ...list.map(([label, prst]) => ({ label, run: () => set(prst, label) }))])];
}

/** Texteffekte > Schatten, Spiegelung, Leuchten: the same Office presets on the text's runs (a:rPr effectLst). */
export function textEffectEntries(app: PresentApp): MenuEntry[] {
  const onRuns = (part: "shadow" | "reflection" | "glow") => (p: PicturePreset | null, label: string) => app.formatText(label, (r) => {
    let fx = kid(r, "effectLst");
    if (!fx) {
      fx = make(r.ownerDocument, "a:effectLst");
      r.insertBefore(fx, kids(r).find((c) => !["ln", "noFill", "solidFill", "gradFill", "pattFill", "blipFill", "grpFill"].includes(c.localName)) ?? null);
    }
    for (const t of FX_TAGS[part]) kid(fx, t)?.remove();
    const from = p ? presetEls(r.ownerDocument, p).find((e) => e.localName === "effectLst") : undefined;
    for (const e of kids(from).filter((x) => FX_TAGS[part].includes(x.localName))) fx.insertBefore(e, kids(fx).find((c) => FX_ORDER.indexOf(c.localName) > FX_ORDER.indexOf(e.localName)) ?? null);
    if (!fx.children.length) fx.remove();
  });
  return [
    { label: "Schatten", submenu: () => grouped(SHADOWS, [["Außen", 9], ["Innen", 9], ["Perspektive", 5]], onRuns("shadow"), "Kein Schatten", "Schattenoptionen...") },
    { label: "Spiegelung", submenu: () => grouped(REFLECTIONS, [["Spiegelungsvarianten", 9]], onRuns("reflection"), "Keine Spiegelung", "Spiegelungsoptionen...") },
    { label: "Leuchten", submenu: () => grouped(GLOWS, [["Leuchteffektvarianten", 24]], onRuns("glow"), "Kein Leuchteffekt", "Leuchteffektoptionen...") },
  ];
}

