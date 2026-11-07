// Form formatieren > Effekte (D-076): Schatten, Spiegelung, Leuchten, Weiche Kanten, 3D-Format and 3D-Drehung
// as sections of the task pane, as PowerPoint's (the "…optionen..." entries of the Formeffekte menus open it).
// Each control reads the first selected shape's spPr and writes every selected shape's: a:effectLst (outerShdw,
// reflection, glow, softEdge in the schema's order), a:scene3d (camera with its rotation, lightRig) and a:sp3d
// (bevelT, bevelB, extrusion, contour, material), as PowerPoint writes them.

import { PptxDom } from "@raabe/ooxml";
import type { PaneControl, PaneSection, PaneTab } from "@raabe/shell";
import { GLOWS, REFLECTIONS, ROTATIONS_3D, SHADOWS, SOFT_EDGES, type PicturePreset } from "./effect-presets";
import { applyPart, type Part } from "./shapefx";

const { kid, kids, make } = PptxDom;
const PT = 12700;
const ANG = 60000;

type Edit = (fn: (spPr: Element) => void, label: string) => void;

/** spPr children in the schema's order (the ones after the outline). */
const SPPR_ORDER = ["xfrm", "custGeom", "prstGeom", "noFill", "solidFill", "gradFill", "blipFill", "pattFill", "grpFill", "ln", "effectLst", "effectDag", "scene3d", "sp3d", "extLst"];
const EFFECT_ORDER = ["blur", "fillOverlay", "glow", "innerShdw", "outerShdw", "prstShdw", "reflection", "softEdge"];
const SP3D_ORDER = ["bevelT", "bevelB", "extrusionClr", "contourClr", "extLst"];

function ensure(parent: Element, name: string, order: string[], attrs?: Record<string, string>): Element {
  let el = kid(parent, name);
  if (el) return el;
  el = make(parent.ownerDocument, `a:${name}`, attrs);
  const i = order.indexOf(name);
  const before = kids(parent).find((c) => order.indexOf(c.localName) > i);
  parent.insertBefore(el, before ?? null);
  return el;
}

const num = (el: Element | undefined, a: string, d = 0) => { const v = el?.getAttribute(a); return v == null ? d : Number(v); };

/** The colour of an effect element (first a:srgbClr/schemeClr child) as #RRGGBB and its transparency. */
function colorOf(el: Element | undefined): { hex: string | null; alpha: number } {
  const c = el ? kids(el).find((x) => x.localName === "srgbClr" || x.localName === "schemeClr" || x.localName === "prstClr") : undefined;
  const a = c ? kid(c, "alpha") : undefined;
  const hex = c?.localName === "srgbClr" ? `#${c.getAttribute("val")}` : c?.localName === "prstClr" && c.getAttribute("val") === "black" ? "#000000" : null;
  return { hex, alpha: a ? 1 - Number(a.getAttribute("val")) / 100000 : 0 };
}

function setColor(el: Element, hex: string | null, alpha: number): void {
  for (const c of kids(el)) if (["srgbClr", "schemeClr", "prstClr", "sysClr", "scrgbClr", "hslClr"].includes(c.localName)) c.remove();
  const c = make(el.ownerDocument, "a:srgbClr", { val: (hex ?? "#000000").replace("#", "").toUpperCase() });
  if (alpha > 0) c.append(make(el.ownerDocument, "a:alpha", { val: String(Math.round((1 - alpha) * 100000)) }));
  el.prepend(c);
}

const MATERIALS: [string, string][] = [["warmMatte", "Warme Mattierung"], ["matte", "Matt"], ["plastic", "Plastik"], ["metal", "Metall"], ["dkEdge", "Dunkle Kante"], ["softEdge", "Weiche Kante"],
  ["flat", "Flach"], ["legacyWireframe", "Drahtmodell"], ["powder", "Puder"], ["translucentPowder", "Transparentes Puder"], ["clear", "Klar"], ["softmetal", "Weiches Metall"]];
const RIGS: [string, string][] = [["threePt", "Drei Punkte"], ["balanced", "Ausgeglichen"], ["soft", "Weich"], ["harsh", "Hart"], ["flood", "Flutlicht"], ["contrasting", "Kontrastreich"],
  ["morning", "Morgen"], ["sunrise", "Sonnenaufgang"], ["sunset", "Sonnenuntergang"], ["chilly", "Kühl"], ["freezing", "Eiskalt"], ["flat", "Flach"], ["twoPt", "Zwei Punkte"], ["glow", "Leuchten"], ["brightRoom", "Heller Raum"]];
const BEVELS: [string, string][] = [["circle", "Kreis"], ["relaxedInset", "Leicht eingesetzt"], ["cross", "Kreuz"], ["coolSlant", "Kühle Neigung"], ["angle", "Winkel"], ["softRound", "Weiche Rundung"],
  ["convex", "Konvex"], ["slope", "Schräge"], ["divot", "Kerbe"], ["riblet", "Rippe"], ["hardEdge", "Harte Kante"], ["artDeco", "Art déco"]];

/** An element's XML without namespace declarations, to compare an effect with a preset's. */
const plain = (el: Element): string => new XMLSerializer().serializeToString(el).replace(/ xmlns(:\w+)?="[^"]*"/g, "");

/** The preset's element of one kind (outerShdw, reflection ...), parsed. */
function presetEl(p: PicturePreset, tag: string): Element | undefined {
  const doc = new DOMParser().parseFromString(`<w xmlns:a="${PptxDom.NS_A}">${p.xml}</w>`, "application/xml");
  return doc.getElementsByTagNameNS(PptxDom.NS_A, tag)[0];
}

/** The Effekte tab for the pane: `spPr` gives the first selected shape's, `edit` changes every selected one. */
export function effectsTab(spPr: () => Element | undefined, edit: Edit): PaneTab {
  const eff = (name: string) => { const s = spPr(); const l = s ? kid(s, "effectLst") : undefined; return l ? kid(l, name) : undefined; };
  const withEffect = (name: string, fn: (e: Element) => void, label: string, attrs?: Record<string, string>) =>
    edit((sp) => fn(ensure(ensure(sp, "effectLst", SPPR_ORDER), name, EFFECT_ORDER, attrs)), label);
  const remove = (name: string, label: string) => edit((sp) => { const l = kid(sp, "effectLst"); kid(l, name)?.remove(); }, label);
  /** Voreinstellungen (PowerPoint's list at the top of each section, D-076): the menu's presets, the first one "none". */
  const presets = (list: readonly (PicturePreset & { id: string })[], tag: string, part: Part, label: string, el: Element | undefined): PaneControl => {
    const cur = el ? plain(el) : "";
    const hit = el ? list.slice(1).find((p) => { const x = presetEl(p, tag); return !!x && plain(x) === cur; }) : undefined;
    const options = list.map((p, i) => ({ value: i ? p.id : "", label: p.label }));
    if (el && !hit) options.push({ value: "custom", label: "Benutzerdefiniert" });
    return { type: "select", label: "Voreinstellungen", options, value: () => (el ? (hit?.id ?? "custom") : ""),
      set: (v) => { if (v === "custom") return; edit((sp) => applyPart(sp, part, list.find((p) => p.id === v && v) ?? null), label); } };
  };

  const shadow: PaneSection = { title: "Schatten", open: false, controls: () => {
    const s = eff("outerShdw");
    const c = colorOf(s);
    const set = (fn: (e: Element) => void) => withEffect("outerShdw", (e) => { if (!kids(e).length) setColor(e, "#000000", 0.6); fn(e); }, "Schatten", { blurRad: "50800", dist: "38100", dir: "2700000", algn: "tl", rotWithShape: "0" });
    return [
      presets(SHADOWS, "outerShdw", "shadow", "Schatten", s ?? eff("innerShdw") ?? eff("prstShdw")),
      { type: "color", label: "Farbe", value: () => c.hex, set: (hex) => set((e) => setColor(e, hex, colorOf(e).alpha)), disabled: () => !s },
      { type: "number", label: "Transparenz", unit: " %", min: 0, max: 100, slider: true, value: () => Math.round(c.alpha * 100), set: (v) => set((e) => setColor(e, colorOf(e).hex, v / 100)), disabled: () => !s },
      { type: "number", label: "Größe", unit: " %", min: 1, max: 200, value: () => Math.round(num(s, "sx", 100000) / 1000), set: (v) => set((e) => { e.setAttribute("sx", String(v * 1000)); e.setAttribute("sy", String(v * 1000)); }), disabled: () => !s },
      { type: "number", label: "Weichzeichnen", unit: " Pt.", min: 0, max: 100, step: 0.5, decimals: 1, value: () => num(s, "blurRad") / PT, set: (v) => set((e) => e.setAttribute("blurRad", String(Math.round(v * PT)))), disabled: () => !s },
      { type: "number", label: "Winkel", unit: "°", min: 0, max: 359.9, value: () => Math.round(num(s, "dir") / ANG), set: (v) => set((e) => e.setAttribute("dir", String(Math.round(v * ANG)))), disabled: () => !s },
      { type: "number", label: "Abstand", unit: " Pt.", min: 0, max: 200, step: 0.5, decimals: 1, value: () => num(s, "dist") / PT, set: (v) => set((e) => e.setAttribute("dist", String(Math.round(v * PT)))), disabled: () => !s },
    ];
  } };

  const reflection: PaneSection = { title: "Spiegelung", open: false, controls: () => {
    const r = eff("reflection");
    const set = (fn: (e: Element) => void) => withEffect("reflection", fn, "Spiegelung", { blurRad: "6350", stA: "50000", endA: "300", endPos: "55000", dir: "5400000", sy: "-100000", algn: "bl", rotWithShape: "0" });
    return [
      presets(REFLECTIONS, "reflection", "reflection", "Spiegelung", r),
      { type: "number", label: "Transparenz", unit: " %", min: 0, max: 100, slider: true, value: () => Math.round(100 - num(r, "stA", 100000) / 1000), set: (v) => set((e) => e.setAttribute("stA", String(Math.round((100 - v) * 1000)))), disabled: () => !r },
      { type: "number", label: "Größe", unit: " %", min: 1, max: 100, slider: true, value: () => Math.round(num(r, "endPos", 100000) / 1000), set: (v) => set((e) => e.setAttribute("endPos", String(v * 1000))), disabled: () => !r },
      { type: "number", label: "Weichzeichnen", unit: " Pt.", min: 0, max: 100, step: 0.5, decimals: 1, value: () => num(r, "blurRad") / PT, set: (v) => set((e) => e.setAttribute("blurRad", String(Math.round(v * PT)))), disabled: () => !r },
      { type: "number", label: "Abstand", unit: " Pt.", min: 0, max: 100, step: 0.5, decimals: 1, value: () => num(r, "dist") / PT, set: (v) => set((e) => e.setAttribute("dist", String(Math.round(v * PT)))), disabled: () => !r },
    ];
  } };

  const glow: PaneSection = { title: "Leuchten", open: false, controls: () => {
    const g = eff("glow");
    const c = colorOf(g);
    const set = (fn: (e: Element) => void) => withEffect("glow", (e) => { if (!kids(e).length) setColor(e, "#4472C4", 0.6); fn(e); }, "Leuchten", { rad: "63500" });
    return [
      presets(GLOWS, "glow", "glow", "Leuchten", g),
      { type: "color", label: "Farbe", value: () => c.hex, set: (hex) => set((e) => setColor(e, hex, colorOf(e).alpha)), disabled: () => !g },
      { type: "number", label: "Größe", unit: " Pt.", min: 0, max: 150, step: 0.5, decimals: 1, value: () => num(g, "rad") / PT, set: (v) => set((e) => e.setAttribute("rad", String(Math.round(v * PT)))), disabled: () => !g },
      { type: "number", label: "Transparenz", unit: " %", min: 0, max: 100, slider: true, value: () => Math.round(c.alpha * 100), set: (v) => set((e) => setColor(e, colorOf(e).hex, v / 100)), disabled: () => !g },
    ];
  } };

  const soft: PaneSection = { title: "Weiche Kanten", open: false, controls: () => {
    const s = eff("softEdge");
    return [presets(SOFT_EDGES, "softEdge", "softEdge", "Weiche Kanten", s),
      { type: "number", label: "Größe", unit: " Pt.", min: 0, max: 100, step: 0.5, decimals: 1, value: () => num(s, "rad") / PT,
      set: (v) => (v > 0 ? withEffect("softEdge", (e) => e.setAttribute("rad", String(Math.round(v * PT))), "Weiche Kanten", { rad: "0" }) : remove("softEdge", "Weiche Kanten")) }];
  } };

  const sp3d = () => { const s = spPr(); return s ? kid(s, "sp3d") : undefined; };
  const scene = () => { const s = spPr(); return s ? kid(s, "scene3d") : undefined; };
  /** sp3d and, as PowerPoint writes it, a scene3d with an orthographic camera and a three-point light. */
  const with3d = (fn: (sp3d: Element) => void, label: string) => edit((sp) => {
    if (!kid(sp, "scene3d")) {
      const sc = ensure(sp, "scene3d", SPPR_ORDER);
      sc.append(make(sp.ownerDocument, "a:camera", { prst: "orthographicFront" }), make(sp.ownerDocument, "a:lightRig", { rig: "threePt", dir: "t" }));
    }
    fn(ensure(sp, "sp3d", SPPR_ORDER));
  }, label);
  const bevel = (name: "bevelT" | "bevelB", title: string): PaneControl[] => {
    const b = sp3d() ? kid(sp3d(), name) : undefined;
    const set = (fn: (e: Element) => void) => with3d((s) => fn(ensure(s, name, SP3D_ORDER, { w: "76200", h: "76200" })), "3D-Format");
    return [
      { type: "select", label: title, options: [{ value: "", label: "Keine" }, ...BEVELS.map(([value, label]) => ({ value, label }))], value: () => (b ? b.getAttribute("prst") ?? "circle" : ""),
        set: (v) => (v ? set((e) => (v === "circle" ? e.removeAttribute("prst") : e.setAttribute("prst", v))) : with3d((s) => kid(s, name)?.remove(), "3D-Format")) },
      { type: "number", label: "Breite", unit: " Pt.", min: 0, max: 1584, step: 0.5, decimals: 1, value: () => num(b, "w", 76200) / PT, set: (v) => set((e) => e.setAttribute("w", String(Math.round(v * PT)))), disabled: () => !b },
      { type: "number", label: "Höhe", unit: " Pt.", min: 0, max: 1584, step: 0.5, decimals: 1, value: () => num(b, "h", 76200) / PT, set: (v) => set((e) => e.setAttribute("h", String(Math.round(v * PT)))), disabled: () => !b },
    ];
  };
  const format3d: PaneSection = { title: "3D-Format", open: false, controls: () => {
    const s = sp3d();
    const lr = scene() ? kid(scene(), "lightRig") : undefined;
    const clr = (name: "extrusionClr" | "contourClr") => colorOf(s ? kid(s, name) : undefined).hex;
    return [
      ...bevel("bevelT", "Abschrägung oben"),
      ...bevel("bevelB", "Abschrägung unten"),
      { type: "heading", label: "Tiefe" },
      { type: "color", label: "Farbe", value: () => clr("extrusionClr"), set: (hex) => with3d((e) => setColor(ensure(e, "extrusionClr", SP3D_ORDER), hex, 0), "3D-Format") },
      { type: "number", label: "Größe", unit: " Pt.", min: 0, max: 1584, step: 0.5, decimals: 1, value: () => num(s, "extrusionH") / PT, set: (v) => with3d((e) => e.setAttribute("extrusionH", String(Math.round(v * PT))), "3D-Format") },
      { type: "heading", label: "Kontur" },
      { type: "color", label: "Farbe", value: () => clr("contourClr"), set: (hex) => with3d((e) => setColor(ensure(e, "contourClr", SP3D_ORDER), hex, 0), "3D-Format") },
      { type: "number", label: "Größe", unit: " Pt.", min: 0, max: 1584, step: 0.5, decimals: 1, value: () => num(s, "contourW") / PT, set: (v) => with3d((e) => e.setAttribute("contourW", String(Math.round(v * PT))), "3D-Format") },
      { type: "heading", label: "Material" },
      { type: "select", label: "Material", options: MATERIALS.map(([value, label]) => ({ value, label })), value: () => s?.getAttribute("prstMaterial") ?? "warmMatte", set: (v) => with3d((e) => e.setAttribute("prstMaterial", v), "3D-Format") },
      { type: "heading", label: "Beleuchtung" },
      { type: "select", label: "Beleuchtung", options: RIGS.map(([value, label]) => ({ value, label })), value: () => lr?.getAttribute("rig") ?? "threePt",
        set: (v) => with3d((e) => { const l = kid(kid(e.parentElement!, "scene3d"), "lightRig"); l?.setAttribute("rig", v); }, "3D-Format") },
      { type: "number", label: "Winkel", unit: "°", min: 0, max: 359.9, value: () => Math.round(num(lr ? kid(lr, "rot") : undefined, "rev") / ANG),
        set: (v) => with3d((e) => { const l = kid(kid(e.parentElement!, "scene3d"), "lightRig"); if (!l) return; let r = kid(l, "rot"); if (!r) { r = make(l.ownerDocument, "a:rot", { lat: "0", lon: "0", rev: "0" }); l.append(r); } r.setAttribute("rev", String(Math.round(v * ANG))); }, "3D-Format") },
      // Zurücksetzen: no 3-D format left (sp3d gone); a camera rotation (3D-Drehung) stays, the light goes back to its default.
      { type: "button", label: "Zurücksetzen", disabled: () => !s, run: () => edit((sp) => {
        kid(sp, "sp3d")?.remove();
        const sc = kid(sp, "scene3d"), cam = sc ? kid(sc, "camera") : undefined;
        if (sc && cam?.getAttribute("prst") === "orthographicFront" && !kid(cam, "rot")) sc.remove();
        else if (sc) { kid(sc, "lightRig")?.remove(); sc.append(make(sp.ownerDocument, "a:lightRig", { rig: "threePt", dir: "t" })); }
      }, "3D-Format") },
    ];
  } };

  const rotation: PaneSection = { title: "3D-Drehung", open: false, controls: () => {
    const cam = scene() ? kid(scene(), "camera") : undefined;
    const rot = cam ? kid(cam, "rot") : undefined;
    const setRot = (attr: "lon" | "lat" | "rev", v: number) => edit((sp) => {
      const sc = ensure(sp, "scene3d", SPPR_ORDER);
      let c = kid(sc, "camera");
      if (!c) { c = make(sp.ownerDocument, "a:camera", { prst: "orthographicFront" }); sc.prepend(c); }
      if (!kid(sc, "lightRig")) sc.append(make(sp.ownerDocument, "a:lightRig", { rig: "threePt", dir: "t" }));
      let r = kid(c, "rot");
      if (!r) { r = make(sp.ownerDocument, "a:rot", { lat: "0", lon: "0", rev: "0" }); c.append(r); }
      r.setAttribute(attr, String(Math.round((((v % 360) + 360) % 360) * ANG)));
    }, "3D-Drehung");
    const persp = !!cam && /perspective/i.test(cam.getAttribute("prst") ?? "");
    // The buttons beside each angle turn it in steps of 10 degrees.
    const nudge = (attr: "lon" | "lat" | "rev", d: number, a: string, b: string): PaneControl => ({ type: "row", items: [
      { type: "button", label: a, run: () => setRot(attr, num(rot, attr) / ANG + d) },
      { type: "button", label: b, run: () => setRot(attr, num(rot, attr) / ANG - d) }] });
    const hasRot = !!rot && ["lon", "lat", "rev"].some((x) => num(rot, x) !== 0);
    return [
      presets(ROTATIONS_3D, "camera", "rotation", "3D-Drehung", cam && (hasRot || cam.getAttribute("prst") !== "orthographicFront") ? cam : undefined),
      { type: "number", label: "X-Drehung", unit: "°", min: 0, max: 359.9, step: 1, value: () => num(rot, "lon") / ANG, set: (v) => setRot("lon", v) },
      nudge("lon", 10, "Links", "Rechts"),
      { type: "number", label: "Y-Drehung", unit: "°", min: 0, max: 359.9, step: 1, value: () => num(rot, "lat") / ANG, set: (v) => setRot("lat", v) },
      nudge("lat", 10, "Nach oben", "Nach unten"),
      { type: "number", label: "Z-Drehung", unit: "°", min: 0, max: 359.9, step: 1, value: () => num(rot, "rev") / ANG, set: (v) => setRot("rev", v) },
      nudge("rev", 10, "Uhrzeigersinn", "Gegen Uhrzeigersinn"),
      { type: "number", label: "Perspektive", unit: "°", min: 0.1, max: 120, step: 1, value: () => (persp ? num(cam, "fov", 2700000) / ANG : null),
        set: (v) => edit((sp) => kid(kid(sp, "scene3d"), "camera")?.setAttribute("fov", String(Math.round(v * ANG))), "3D-Drehung"), disabled: () => !persp },
      { type: "number", label: "Abstand vom Boden", unit: " Pt.", min: -4000, max: 4000, step: 0.5, decimals: 1, value: () => num(sp3d(), "z") / PT,
        set: (v) => with3d((e) => (v ? e.setAttribute("z", String(Math.round(v * PT))) : e.removeAttribute("z")), "3D-Drehung") },
      // Zurücksetzen: no rotation (the camera straight from the front), as the menu's "Keine Drehung".
      { type: "button", label: "Zurücksetzen", disabled: () => !cam, run: () => edit((sp) => applyPart(sp, "rotation", null), "3D-Drehung") },
    ];
  } };

  return { id: "effects", label: "Effekte", icon: "pane-effects", sections: () => [shadow, reflection, glow, soft, format3d, rotation] };
}
