// Animationen (D-070): entrance, emphasis and exit effects written into the slide's p:timing as
// PowerPoint writes them (recorded: Verblassen is presetID 10 with a visibility set and an
// animEffect "fade", Einfliegen presetID 2 with ppt_x/ppt_y animations; results in
// packages/ooxml/test/fixtures/pptx-show.pptx). The show plays the main sequence click by click.

import { PptxDom } from "@raabe/ooxml";
import type { PresentApp } from "./app";

const { kid, kids, num, at, fragment } = PptxDom;

export interface AnimItem { label: string; preset: number; subtype?: number; effect: string }

export const ANIMATION_GROUPS: { title: string; cls: "entr" | "emph" | "exit" | "path"; items: AnimItem[] }[] = [
  { title: "Eingang", cls: "entr", items: [
    { label: "Erscheinen", preset: 1, effect: "appear" }, { label: "Verblassen", preset: 10, effect: "fade" }, { label: "Einfliegen", preset: 2, subtype: 4, effect: "fly" },
    { label: "Hineinschweben", preset: 42, effect: "float" }, { label: "Teilen", preset: 16, subtype: 21, effect: "split" }, { label: "Wischen", preset: 22, subtype: 4, effect: "wipe" },
    { label: "Form", preset: 6, effect: "circle" }, { label: "Rad", preset: 21, subtype: 1, effect: "wheel" }, { label: "Zufällige Balken", preset: 14, subtype: 10, effect: "randombar" },
    { label: "Wachsen und Bewegen", preset: 31, effect: "growturn" }, { label: "Zoom", preset: 53, subtype: 16, effect: "zoom" }, { label: "Drehen", preset: 49, effect: "swivel" }, { label: "Springen", preset: 26, effect: "bounce" }] },
  { title: "Hervorheben", cls: "emph", items: [
    { label: "Impuls", preset: 26, effect: "pulse" }, { label: "Farbimpuls", preset: 27, effect: "colorpulse" }, { label: "Schwanken", preset: 32, effect: "teeter" }, { label: "Rotieren", preset: 8, effect: "spin" },
    { label: "Vergrößern/Verkleinern", preset: 6, effect: "grow" }, { label: "Durchtränken", preset: 19, effect: "desaturate" }, { label: "Verdunkeln", preset: 20, effect: "darken" }, { label: "Erhellen", preset: 21, effect: "lighten" },
    { label: "Transparent", preset: 9, effect: "transparency" }, { label: "Objektfarbe", preset: 54, effect: "objcolor" }, { label: "Komplementärfarbe", preset: 55, effect: "objcolor" }, { label: "Linienfarbe", preset: 7, effect: "objcolor" },
    { label: "Füllfarbe", preset: 1, effect: "objcolor" }, { label: "Pinselfarbe", preset: 16, effect: "objcolor" }, { label: "Schriftfarbe", preset: 3, effect: "objcolor" }, { label: "Unterstreichen", preset: 18, effect: "pulse" },
    { label: "Deutlicher Blitz", preset: 35, effect: "pulse" }, { label: "Fett anzeigen", preset: 15, effect: "pulse" }, { label: "Welle", preset: 34, effect: "teeter" }] },
  { title: "Beenden", cls: "exit", items: [
    { label: "Verschwinden", preset: 1, effect: "appear" }, { label: "Verblassen", preset: 10, effect: "fade" }, { label: "Hinausfliegen", preset: 2, subtype: 4, effect: "fly" },
    { label: "Herausschweben", preset: 42, effect: "float" }, { label: "Teilen", preset: 16, subtype: 21, effect: "split" }, { label: "Wischen", preset: 22, subtype: 4, effect: "wipe" },
    { label: "Form", preset: 6, effect: "circle" }, { label: "Rad", preset: 21, subtype: 1, effect: "wheel" }, { label: "Zufällige Balken", preset: 14, subtype: 10, effect: "randombar" },
    { label: "Verkleinern und Drehen", preset: 31, effect: "growturn" }, { label: "Zoom", preset: 53, subtype: 16, effect: "zoom" }, { label: "Drehen", preset: 49, effect: "swivel" }, { label: "Springen", preset: 26, effect: "bounce" }] },
  { title: "Animationspfade", cls: "path", items: [
    { label: "Linien", preset: 42, effect: "path" }, { label: "Bögen", preset: 37, effect: "path" }, { label: "Wendungen", preset: 39, effect: "path" }, { label: "Formen", preset: 35, effect: "path" },
    { label: "Schleifen", preset: 36, effect: "path" }, { label: "Benutzerdefinierter Pfad", preset: 0, effect: "path" }] },
];

const TIMING_HEAD = '<p:timing><p:tnLst><p:par><p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot"><p:childTnLst><p:seq concurrent="1" nextAc="seek"><p:cTn id="2" dur="indefinite" nodeType="mainSeq"><p:childTnLst>';
const TIMING_TAIL = '</p:childTnLst></p:cTn><p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst><p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:nextCondLst></p:seq></p:childTnLst></p:cTn></p:par></p:tnLst><p:bldLst/></p:timing>';

/** One effect's behaviours, as PowerPoint writes them for the presets we offer. */
/**
 * Animationspfade (D-072): each preset's path in PowerPoint's notation (fractions of the slide, relative to
 * the shape's place); our own curves for Bögen, Wendungen, Formen (a circle) and Schleifen.
 */
export const MOTION_PATHS: Record<number, string> = {
  42: "M 0 0 L 0 0.25 E",
  37: "M 0 0 C 0.069 0 0.125 0.07 0.125 0.157 E",
  39: "M 0 0 L 0 0.12 C 0 0.17 0.03 0.2 0.08 0.2 L 0.25 0.2 E",
  35: "M 0 0 C 0.069 0 0.125 0.056 0.125 0.125 C 0.125 0.194 0.069 0.25 0 0.25 C -0.069 0.25 -0.125 0.194 -0.125 0.125 C -0.125 0.056 -0.069 0 0 0 Z",
  36: "M 0 0 C 0.1 0 0.15 0.08 0.1 0.15 C 0.05 0.22 -0.05 0.15 0 0.08 C 0.04 0.02 0.12 0.05 0.2 0.1 E",
};

/** Points along a motion path (fractions of the slide), for the show to follow. */
export function samplePath(path: string, n = 40): [number, number][] {
  const tok = path.trim().split(/\s+/);
  const segs: ((t: number) => [number, number])[] = [];
  let x = 0, y = 0, sx = 0, sy = 0;
  for (let i = 0; i < tok.length;) {
    const cmd = tok[i++];
    const num = () => Number(tok[i++]);
    if (cmd === "M") { x = num(); y = num(); sx = x; sy = y; }
    else if (cmd === "L") { const x0 = x, y0 = y, x1 = num(), y1 = num(); segs.push((t) => [x0 + (x1 - x0) * t, y0 + (y1 - y0) * t]); x = x1; y = y1; }
    else if (cmd === "C") {
      const x0 = x, y0 = y, x1 = num(), y1 = num(), x2 = num(), y2 = num(), x3 = num(), y3 = num();
      segs.push((t) => { const u = 1 - t; return [u * u * u * x0 + 3 * u * u * t * x1 + 3 * u * t * t * x2 + t * t * t * x3, u * u * u * y0 + 3 * u * u * t * y1 + 3 * u * t * t * y2 + t * t * t * y3]; });
      x = x3; y = y3;
    } else if (cmd === "Z") { const x0 = x, y0 = y; segs.push((t) => [x0 + (sx - x0) * t, y0 + (sy - y0) * t]); x = sx; y = sy; }
  }
  if (!segs.length) return [[0, 0]];
  const out: [number, number][] = [];
  for (let k = 0; k <= n; k++) { const f = (k / n) * segs.length, s = Math.min(segs.length - 1, Math.floor(f)); out.push(segs[s](f - s)); }
  return out;
}

function behaviours(cls: string, it: AnimItem, spid: number, id: () => number, dur: number): string {
  const tgt = `<p:tgtEl><p:spTgt spid="${spid}"/></p:tgtEl>`;
  if (cls === "path") return `<p:animMotion origin="layout" path="${MOTION_PATHS[it.preset] ?? MOTION_PATHS[42]}" pathEditMode="relative" ptsTypes=""><p:cBhvr><p:cTn id="${id()}" dur="${dur}" fill="hold"/>${tgt}<p:attrNameLst><p:attrName>ppt_x</p:attrName><p:attrName>ppt_y</p:attrName></p:attrNameLst></p:cBhvr></p:animMotion>`;
  const vis = (v: string, delay = 0) => `<p:set><p:cBhvr><p:cTn id="${id()}" dur="1" fill="hold"><p:stCondLst><p:cond delay="${delay}"/></p:stCondLst></p:cTn>${tgt}<p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr><p:to><p:strVal val="${v}"/></p:to></p:set>`;
  const fx = (filter: string, dir: string) => `<p:animEffect transition="${dir}" filter="${filter}"><p:cBhvr><p:cTn id="${id()}" dur="${dur}"/>${tgt}</p:cBhvr></p:animEffect>`;
  const anim = (attr: string, from: string, to: string) => `<p:anim calcmode="lin" valueType="num"><p:cBhvr additive="base"><p:cTn id="${id()}" dur="${dur}" fill="hold"/>${tgt}<p:attrNameLst><p:attrName>${attr}</p:attrName></p:attrNameLst></p:cBhvr><p:tavLst><p:tav tm="0"><p:val><p:strVal val="${from}"/></p:val></p:tav><p:tav tm="100000"><p:val><p:strVal val="${to}"/></p:val></p:tav></p:tavLst></p:anim>`;
  if (cls === "emph") {
    if (it.effect === "spin") return `<p:animRot by="21600000"><p:cBhvr><p:cTn id="${id()}" dur="${dur}" fill="hold"/>${tgt}<p:attrNameLst><p:attrName>r</p:attrName></p:attrNameLst></p:cBhvr></p:animRot>`;
    return `<p:animScale><p:cBhvr><p:cTn id="${id()}" dur="${dur}" fill="hold"/>${tgt}</p:cBhvr><p:by x="150000" y="150000"/></p:animScale>`;
  }
  const entr = cls === "entr";
  const filters: Record<string, string> = { fade: "fade", wipe: "wipe(down)", split: "barn(inVertical)", circle: "circle(in)", wheel: "wheel(1)", randombar: "randombar(horizontal)" };
  if (entr) {
    if (it.effect === "appear") return vis("visible");
    if (it.effect === "fly") return vis("visible") + anim("ppt_x", "#ppt_x", "#ppt_x") + anim("ppt_y", "1+#ppt_h/2", "#ppt_y");
    if (it.effect === "float") return vis("visible") + fx("fade", "in") + anim("ppt_x", "#ppt_x", "#ppt_x") + anim("ppt_y", "#ppt_y+.1", "#ppt_y");
    if (it.effect === "zoom" || it.effect === "growturn" || it.effect === "swivel" || it.effect === "bounce") return vis("visible") + anim("ppt_w", "0", "#ppt_w") + anim("ppt_h", "0", "#ppt_h") + fx("fade", "in");
    return vis("visible") + fx(filters[it.effect] ?? "fade", "in");
  }
  if (it.effect === "appear") return vis("hidden");
  if (it.effect === "fly") return anim("ppt_x", "#ppt_x", "#ppt_x") + anim("ppt_y", "#ppt_y", "1+#ppt_h/2") + vis("hidden", dur - 1);
  return fx(filters[it.effect] ?? "fade", "out") + vis("hidden", dur - 1);
}

function maxCtnId(root: Element): number {
  let m = 2;
  for (const c of Array.from(root.getElementsByTagNameNS(PptxDom.NS_P, "cTn"))) m = Math.max(m, num(c, "id") ?? 0);
  return m;
}

/** Animation (replace the selected shapes' effects) or Animation hinzufügen (add another), as a click effect. */
export function addAnimation(app: PresentApp, cls: "entr" | "emph" | "exit" | "path", it: AnimItem, add: boolean): void {
  if (cls === "path" && !MOTION_PATHS[it.preset]) return; // Benutzerdefinierter Pfad: drawing a path is not built yet (greyed).
  const ed = app.editor;
  const ids = ed.text ? [ed.text.id] : ed.sel;
  const path = app.slidePath();
  if (!path || !ids.length) return;
  ed.edit(add ? "Animation hinzufügen" : "Animation", () => {
    const root = app.deck.pptx.doc(path).documentElement;
    let timing = kid(root, "timing");
    if (!timing) {
      timing = fragment(root.ownerDocument, TIMING_HEAD + TIMING_TAIL)[0];
      root.insertBefore(timing, kid(root, "extLst") ?? null);
    }
    const main = Array.from(timing.getElementsByTagNameNS(PptxDom.NS_P, "cTn")).find((c) => c.getAttribute("nodeType") === "mainSeq");
    const list = kid(main, "childTnLst")!;
    if (!add) for (const id of ids) removeEffects(timing, id);
    let next = maxCtnId(root);
    const nid = () => ++next;
    for (const spid of ids) {
      const grp = nextGrp(timing, spid);
      const dur = it.effect === "appear" ? 1 : cls === "emph" || cls === "path" ? 2000 : 500;
      const xml = `<p:par><p:cTn id="${nid()}" fill="hold"><p:stCondLst><p:cond delay="indefinite"/></p:stCondLst><p:childTnLst><p:par><p:cTn id="${nid()}" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>` +
        `<p:par><p:cTn id="${nid()}" presetID="${it.preset}" presetClass="${cls}" presetSubtype="${it.subtype ?? 0}" fill="hold" grpId="${grp}" nodeType="clickEffect"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>` +
        behaviours(cls, it, spid, nid, dur) + `</p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par>`;
      list.append(fragment(root.ownerDocument, xml)[0]);
      let bld = kid(timing, "bldLst");
      if (!bld) { bld = PptxDom.make(root.ownerDocument, "p:bldLst"); timing.append(bld); }
      const isSp = app.editor.shapeView(spid)?.kind === "sp";
      if (isSp && !kids(bld, "bldP").some((b) => b.getAttribute("spid") === String(spid) && b.getAttribute("grpId") === String(grp))) bld.append(PptxDom.make(root.ownerDocument, "p:bldP", { spid, grpId: grp }));
    }
    if (!kids(list).length) timing.remove();
  });
}

function nextGrp(timing: Element, spid: number): number {
  let g = -1;
  for (const c of Array.from(timing.getElementsByTagNameNS(PptxDom.NS_P, "cTn"))) {
    if (!c.hasAttribute("presetClass")) continue;
    const t = c.getElementsByTagNameNS(PptxDom.NS_P, "spTgt")[0];
    if (t?.getAttribute("spid") === String(spid)) g = Math.max(g, num(c, "grpId") ?? 0);
  }
  return g + 1;
}

/** Removes every effect on shape `spid` (the click groups that become empty go too). */
export function removeEffects(timing: Element, spid: number): void {
  for (const c of Array.from(timing.getElementsByTagNameNS(PptxDom.NS_P, "cTn"))) {
    if (!c.hasAttribute("presetClass")) continue;
    const t = c.getElementsByTagNameNS(PptxDom.NS_P, "spTgt")[0];
    if (t?.getAttribute("spid") !== String(spid)) continue;
    let par = c.parentElement;
    c.parentElement?.remove();
    // Climb: a par whose childTnLst is empty goes too.
    while (par && par.localName === "childTnLst" && !par.children.length) { const up = par.parentElement?.parentElement; up?.remove(); par = up?.parentElement ?? null; }
  }
  const bld = kid(timing, "bldLst");
  for (const b of kids(bld, "bldP")) if (b.getAttribute("spid") === String(spid)) b.remove();
}

/** A click step of the main sequence: the effects that start with it, in order (with "Mit/Nach Vorherigen"). */
export interface Step { effects: { spid: number; cls: string; effect: string; dur: number; delay: number; after: boolean; node?: string; path?: string }[] }

/** The main sequence as click steps, and the shapes hidden at the start (those with an entrance first). */
export function mainSequence(root: Element): { steps: Step[]; hidden: Set<number>; auto: Step | null } {
  const steps: Step[] = [];
  let auto: Step | null = null;
  const hidden = new Set<number>();
  const seen = new Set<number>();
  const timing = kid(root, "timing");
  const main = timing ? Array.from(timing.getElementsByTagNameNS(PptxDom.NS_P, "cTn")).find((c) => c.getAttribute("nodeType") === "mainSeq") : undefined;
  for (const click of kids(kid(main, "childTnLst"), "par")) {
    const step: Step = { effects: [] };
    for (const c of Array.from(click.getElementsByTagNameNS(PptxDom.NS_P, "cTn"))) {
      const cls = c.getAttribute("presetClass");
      if (!cls) continue;
      const spid = Number(c.getElementsByTagNameNS(PptxDom.NS_P, "spTgt")[0]?.getAttribute("spid") ?? 0);
      const durs = Array.from(c.getElementsByTagNameNS(PptxDom.NS_P, "cTn")).map((x) => Number(x.getAttribute("dur") ?? 0)).filter((d) => Number.isFinite(d));
      const fx = c.getElementsByTagNameNS(PptxDom.NS_P, "animEffect")[0]?.getAttribute("filter") ?? "";
      const preset = Number(c.getAttribute("presetID") ?? 0);
      const effect = preset === 1 ? "appear" : preset === 2 ? "fly" : preset === 8 ? "spin" : preset === 53 ? "zoom" : fx.split("(")[0] || (cls === "emph" ? "pulse" : "fade");
      const delay = Number(at(c, "stCondLst/cond")?.getAttribute("delay") ?? 0) || 0;
      const motion = c.getElementsByTagNameNS(PptxDom.NS_P, "animMotion")[0]?.getAttribute("path") ?? undefined;
      step.effects.push({ spid, cls, effect: motion ? "path" : effect, dur: Math.max(1, ...durs), delay, after: c.getAttribute("nodeType") === "afterEffect", node: c.getAttribute("nodeType") ?? "clickEffect", path: motion });
      if (cls === "entr" && !seen.has(spid)) hidden.add(spid);
      seen.add(spid);
    }
    if (!step.effects.length) continue;
    // Effects that start "Mit Vorherigen" / "Nach Vorherigen" join the click before them (or play at once).
    const byClick = step.effects.some((x) => !x.after && x.node === "clickEffect");
    if (!byClick && steps.length) {
      const prev = steps[steps.length - 1];
      const end = Math.max(0, ...prev.effects.map((x) => x.delay + x.dur));
      for (const x of step.effects) prev.effects.push({ ...x, delay: x.delay + (x.after ? end : 0) });
    } else if (!byClick) {
      auto ??= { effects: [] };
      auto.effects.push(...step.effects);
    } else steps.push(step);
  }
  return { steps, hidden, auto };
}

export function setDuration(app: PresentApp, seconds: number): void { retime(app, "dur", seconds); }
export function setDelay(app: PresentApp, seconds: number): void { retime(app, "delay", seconds); }

function retime(app: PresentApp, what: "dur" | "delay", seconds: number): void {
  const ed = app.editor;
  const path = app.slidePath();
  if (!path || !ed.sel.length) return;
  ed.edit(what === "dur" ? "Dauer" : "Verzögerung", () => {
    const root = app.deck.pptx.doc(path).documentElement;
    const timing = kid(root, "timing");
    if (!timing) return;
    for (const c of Array.from(timing.getElementsByTagNameNS(PptxDom.NS_P, "cTn"))) {
      if (!c.hasAttribute("presetClass")) continue;
      const spid = Number(c.getElementsByTagNameNS(PptxDom.NS_P, "spTgt")[0]?.getAttribute("spid") ?? 0);
      if (!ed.sel.includes(spid)) continue;
      if (what === "delay") { at(c, "stCondLst/cond")?.setAttribute("delay", String(Math.round(seconds * 1000))); continue; }
      for (const x of Array.from(c.getElementsByTagNameNS(PptxDom.NS_P, "cTn"))) if (x !== c && x.getAttribute("dur") && x.getAttribute("dur") !== "1") x.setAttribute("dur", String(Math.round(seconds * 1000)));
    }
  });
}

/** Effektoptionen of the selected shapes' entrance or exit effect: Einfliegen's directions. */
export function animationOptions(app: PresentApp): { label: string; run: () => void; checked?: boolean }[] {
  const path = app.slidePath();
  const ids = new Set(app.editor.sel);
  if (!path || !ids.size) return [];
  const root = app.deck.pptx.doc(path).documentElement;
  const ctns = Array.from(root.getElementsByTagNameNS(PptxDom.NS_P, "cTn")).filter((c) => c.getAttribute("presetID") === "2" && ids.has(Number(c.getElementsByTagNameNS(PptxDom.NS_P, "spTgt")[0]?.getAttribute("spid"))));
  if (!ctns.length) return [];
  // presetSubtype: 1 von oben, 2 von rechts, 4 von unten, 8 von links (PowerPoint's Einfliegen).
  const dirs: [string, number, string, string][] = [["Von unten", 4, "#ppt_x", "1+#ppt_h/2"], ["Von links", 8, "0-#ppt_w/2", "#ppt_y"], ["Von rechts", 2, "1+#ppt_w/2", "#ppt_y"], ["Von oben", 1, "#ppt_x", "0-#ppt_h/2"]];
  const cur = Number(ctns[0].getAttribute("presetSubtype") ?? 4);
  return dirs.map(([label, sub, x, y]) => ({
    label, checked: cur === sub,
    run: () => app.editor.edit("Effektoptionen", () => {
      for (const c of ctns) {
        c.setAttribute("presetSubtype", String(sub));
        const exit = c.getAttribute("presetClass") === "exit";
        for (const anim of Array.from(c.getElementsByTagNameNS(PptxDom.NS_P, "anim"))) {
          const attr = anim.getElementsByTagNameNS(PptxDom.NS_P, "attrName")[0]?.textContent;
          const vals = Array.from(anim.getElementsByTagNameNS(PptxDom.NS_P, "strVal"));
          if (vals.length < 2) continue;
          const away = attr === "ppt_x" ? x : y, home = attr === "ppt_x" ? "#ppt_x" : "#ppt_y";
          vals[exit ? 1 : 0].setAttribute("val", away);
          vals[exit ? 0 : 1].setAttribute("val", home);
        }
      }
    }),
  }));
}

/**
 * Animation übertragen (D-072): the effects of one shape put on another, each as its own click, as
 * PowerPoint's brush does (the target's own effects are replaced).
 */
export function copyAnimation(app: PresentApp, fromId: number, toId: number): void {
  const path = app.slidePath();
  if (!path || fromId === toId) return;
  app.editor.edit("Animation übertragen", () => {
    const root = app.deck.pptx.doc(path).documentElement;
    const timing = kid(root, "timing");
    if (!timing) return;
    const main = Array.from(timing.getElementsByTagNameNS(PptxDom.NS_P, "cTn")).find((c) => c.getAttribute("nodeType") === "mainSeq");
    const list = kid(main, "childTnLst");
    if (!list) return;
    const sources = Array.from(timing.getElementsByTagNameNS(PptxDom.NS_P, "cTn"))
      .filter((c) => c.hasAttribute("presetClass") && c.getElementsByTagNameNS(PptxDom.NS_P, "spTgt")[0]?.getAttribute("spid") === String(fromId))
      .map((c) => c.parentElement!.cloneNode(true) as Element);
    if (!sources.length) return;
    removeEffects(timing, toId);
    let next = maxCtnId(root);
    for (const effect of sources) {
      for (const t of Array.from(effect.getElementsByTagNameNS(PptxDom.NS_P, "spTgt"))) t.setAttribute("spid", String(toId));
      for (const c of Array.from(effect.getElementsByTagNameNS(PptxDom.NS_P, "cTn"))) c.setAttribute("id", String(++next));
      const top = effect.firstElementChild!;
      const grp = nextGrp(timing, toId);
      top.setAttribute("grpId", String(grp));
      top.setAttribute("nodeType", "clickEffect");
      top.setAttribute("id", String(++next));
      const click = fragment(root.ownerDocument, `<p:par><p:cTn id="${++next}" fill="hold"><p:stCondLst><p:cond delay="indefinite"/></p:stCondLst><p:childTnLst><p:par><p:cTn id="${++next}" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst/></p:cTn></p:par></p:childTnLst></p:cTn></p:par>`)[0];
      click.getElementsByTagNameNS(PptxDom.NS_P, "childTnLst")[1].append(effect);
      list.append(click);
      let bld = kid(timing, "bldLst");
      if (!bld) { bld = PptxDom.make(root.ownerDocument, "p:bldLst"); timing.append(bld); }
      if (app.editor.shapeView(toId)?.kind === "sp") bld.append(PptxDom.make(root.ownerDocument, "p:bldP", { spid: toId, grpId: grp }));
    }
  });
}
