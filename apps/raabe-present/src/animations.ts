// Animationen (D-070): entrance, emphasis and exit effects written into the slide's p:timing as
// PowerPoint writes them (recorded: Verblassen is presetID 10 with a visibility set and an
// animEffect "fade", Einfliegen presetID 2 with ppt_x/ppt_y animations; results in
// packages/ooxml/test/fixtures/pptx-show.pptx). The show plays the main sequence click by click.

import { PptxDom } from "@raabe/ooxml";
import type { PresentApp } from "./app";

const { kid, kids, num, at, fragment } = PptxDom;

export interface AnimItem { label: string; preset: number; subtype?: number; effect: string }

export const ANIMATION_GROUPS: { title: string; cls: "entr" | "emph" | "exit"; items: AnimItem[] }[] = [
  { title: "Eingang", cls: "entr", items: [
    { label: "Erscheinen", preset: 1, effect: "appear" }, { label: "Verblassen", preset: 10, effect: "fade" }, { label: "Einfliegen", preset: 2, subtype: 4, effect: "fly" },
    { label: "Hineinschweben", preset: 42, effect: "float" }, { label: "Teilen", preset: 16, subtype: 21, effect: "split" }, { label: "Wischen", preset: 22, subtype: 4, effect: "wipe" },
    { label: "Form", preset: 6, effect: "circle" }, { label: "Rad", preset: 21, subtype: 1, effect: "wheel" }, { label: "Zufällige Balken", preset: 14, subtype: 10, effect: "randombar" },
    { label: "Vergrößern und drehen", preset: 31, effect: "growturn" }, { label: "Zoom", preset: 53, subtype: 16, effect: "zoom" }, { label: "Drehen", preset: 49, effect: "swivel" }, { label: "Hüpfen", preset: 26, effect: "bounce" }] },
  { title: "Hervorheben", cls: "emph", items: [
    { label: "Impuls", preset: 26, effect: "pulse" }, { label: "Farbpulse", preset: 27, effect: "colorpulse" }, { label: "Wippen", preset: 32, effect: "teeter" }, { label: "Rotieren", preset: 8, effect: "spin" },
    { label: "Vergrößern/Verkleinern", preset: 6, effect: "grow" }, { label: "Entsättigen", preset: 19, effect: "desaturate" }, { label: "Abdunkeln", preset: 20, effect: "darken" }, { label: "Aufhellen", preset: 21, effect: "lighten" },
    { label: "Transparenz", preset: 9, effect: "transparency" }, { label: "Objektfarbe", preset: 54, effect: "objcolor" }] },
  { title: "Beenden", cls: "exit", items: [
    { label: "Verschwinden", preset: 1, effect: "appear" }, { label: "Verblassen", preset: 10, effect: "fade" }, { label: "Hinausfliegen", preset: 2, subtype: 4, effect: "fly" },
    { label: "Hinausschweben", preset: 42, effect: "float" }, { label: "Teilen", preset: 16, subtype: 21, effect: "split" }, { label: "Wischen", preset: 22, subtype: 4, effect: "wipe" },
    { label: "Form", preset: 6, effect: "circle" }, { label: "Rad", preset: 21, subtype: 1, effect: "wheel" }, { label: "Zufällige Balken", preset: 14, subtype: 10, effect: "randombar" },
    { label: "Verkleinern und drehen", preset: 31, effect: "growturn" }, { label: "Zoom", preset: 53, subtype: 16, effect: "zoom" }, { label: "Drehen", preset: 49, effect: "swivel" }, { label: "Hüpfen", preset: 26, effect: "bounce" }] },
];

const TIMING_HEAD = '<p:timing><p:tnLst><p:par><p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot"><p:childTnLst><p:seq concurrent="1" nextAc="seek"><p:cTn id="2" dur="indefinite" nodeType="mainSeq"><p:childTnLst>';
const TIMING_TAIL = '</p:childTnLst></p:cTn><p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst><p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:nextCondLst></p:seq></p:childTnLst></p:cTn></p:par></p:tnLst><p:bldLst/></p:timing>';

/** One effect's behaviours, as PowerPoint writes them for the presets we offer. */
function behaviours(cls: string, it: AnimItem, spid: number, id: () => number, dur: number): string {
  const tgt = `<p:tgtEl><p:spTgt spid="${spid}"/></p:tgtEl>`;
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
export function addAnimation(app: PresentApp, cls: "entr" | "emph" | "exit", it: AnimItem, add: boolean): void {
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
      const dur = it.effect === "appear" ? 1 : cls === "emph" ? 2000 : 500;
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
export interface Step { effects: { spid: number; cls: string; effect: string; dur: number; delay: number; after: boolean; node?: string }[] }

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
      step.effects.push({ spid, cls, effect, dur: Math.max(1, ...durs), delay, after: c.getAttribute("nodeType") === "afterEffect", node: c.getAttribute("nodeType") ?? "clickEffect" });
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
