// Formeln in RaabePresent (D-072): an equation is a text box whose paragraph holds Office Math (a14:m with
// m:oMathPara), inside mc:AlternateContent with a plain-text Fallback for readers before PowerPoint 2010,
// as PowerPoint writes it. Drawn as MathML from the equation tree (packages/ooxml/src/math.ts, shared with
// RaabeWrite). Editing happens in PowerPoint's linear format (UnicodeMath, "Aktuell – Linear"): the
// equation's linear text in a box with the professional form shown live; OK writes it back as OMML.

import { h, showDialog, type Control, type MenuEntry, type Tab } from "@raabe/shell";
import { ommlToSeq, seqToOmml, PptxDom, type ShapeView } from "@raabe/ooxml";
import { buildLinear, parseLatex, seqToLatex, seqToLinear, seqToMathml, type Seq } from "@raabe/text-engine";
import { BUILT_IN } from "./equations-builtin";
import { MATH_GALLERIES, SYMBOL_CATEGORIES } from "./equations-gallery";
import type { PresentApp } from "./app";
import * as S from "./shapeops";

const NS_A14 = "http://schemas.microsoft.com/office/drawing/2010/main";
const NS_MC = "http://schemas.openxmlformats.org/markup-compatibility/2006";
const NS_M = "http://schemas.openxmlformats.org/officeDocument/2006/math";
const MATH_RPR = '<a:rPr lang="de-DE" i="1"><a:latin typeface="Cambria Math" panose="02040503050406030204" pitchFamily="18" charset="0"/></a:rPr>';

/** OMML for a slide: the converter's w:rPr become a:rPr (PowerPoint's run properties inside a14:m). */
export function slideOmml(seq: Seq): string {
  return seqToOmml(seq, true).replace(/<w:rPr>.*?<\/w:rPr>/g, MATH_RPR).replace("<m:oMathPara>", `<m:oMathPara xmlns:m="${NS_M}">`);
}

/** The equation (m:oMathPara) of a shape, if its text holds one. */
export function equationOf(sp: Element | undefined): Element | undefined {
  const m = sp?.getElementsByTagNameNS(NS_A14, "m")[0];
  return m?.firstElementChild ?? undefined;
}

export const hasEquation = (v: ShapeView | undefined): boolean => !!v?.text?.paras.some((p) => p.runs.some((r) => r.kind === "math"));

function equationXml(id: number, x: number, y: number, seq: Seq): string {
  const linear = PptxDom.esc(seqToLinear(seq));
  const box = (body: string) => `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="Textfeld ${id - 1}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>` +
    `<p:spPr><a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="2000000" cy="400000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr>` +
    `<p:txBody><a:bodyPr wrap="none" rtlCol="0"><a:spAutoFit/></a:bodyPr><a:lstStyle/><a:p>${body}<a:endParaRPr lang="de-DE"/></a:p></p:txBody></p:sp>`;
  return `<mc:AlternateContent xmlns:mc="${NS_MC}"><mc:Choice xmlns:a14="${NS_A14}" Requires="a14">${box(`<a14:m>${slideOmml(seq)}</a14:m>`)}</mc:Choice>` +
    `<mc:Fallback>${box(`<a:r><a:rPr lang="de-DE"/><a:t>${linear}</a:t></a:r>`)}</mc:Fallback></mc:AlternateContent>`;
}

/** Einfügen > Formel: a new equation in the middle of the slide (a built-in one, or typed in the linear box). */
export async function insertEquation(app: PresentApp, seq?: Seq): Promise<void> {
  const path = app.slidePath();
  if (!path) return;
  const tree = S.treeOf(app.deck.pptx.doc(path));
  const ed = app.editor;
  if (!seq) {
    const typed = await linearDialog([]);
    if (!typed) return;
    seq = typed;
  }
  const size = app.deck.pptx.slideSize;
  const id = S.maxId(tree) + 1;
  const eq = seq;
  ed.edit("Formel", () => { tree.append(PptxDom.fragment(tree.ownerDocument, equationXml(id, Math.round(size.cx / 2 - 1000000), Math.round(size.cy / 2 - 200000), eq))[0]); });
  ed.updateAutofit(id);
  ed.select([id]);
}

/** Editing an equation: its linear form in the box; OK writes the new tree into the Choice and the Fallback's text. */
export async function editEquation(app: PresentApp, id: number): Promise<void> {
  const ed = app.editor;
  const sp = ed.shapeEl(id);
  const m = equationOf(sp);
  if (!sp || !m) return;
  const seq = await linearDialog(ommlToSeq(m) ?? []);
  if (!seq) return;
  ed.edit("Formel", () => {
    const a14m = m.parentElement!;
    a14m.replaceChildren(PptxDom.fragment(a14m.ownerDocument, slideOmml(seq))[0]);
    // The Fallback's text follows (older readers).
    const ac = sp.parentElement?.parentElement;
    if (ac?.namespaceURI === NS_MC) {
      const t = ac.getElementsByTagNameNS(NS_MC, "Fallback")[0]?.getElementsByTagNameNS(PptxDom.NS_A, "t")[0];
      if (t) t.textContent = seqToLinear(seq);
    }
  });
  ed.updateAutofit(id);
  ed.render();
}

/** The linear box: UnicodeMath (x^2, a/b, √(x), ∑_(k=1)^n …) with the professional form below. */
async function linearDialog(seq: Seq): Promise<Seq | null> {
  const input = h("input", { type: "text", value: seqToLinear(seq), style: "width:420px;font:16px 'Cambria Math', serif" }) as HTMLInputElement;
  const preview = h("div", { class: "eq-preview", style: "min-height:60px;padding:8px;font:24px 'Cambria Math', serif;border:1px solid #c8c8c8;margin-top:8px;background:#fff;color:#000" });
  const show = () => { try { preview.innerHTML = seqToMathml(buildLinear(input.value), true); } catch { preview.textContent = input.value; } };
  input.addEventListener("input", show);
  show();
  setTimeout(() => input.focus(), 30);
  const ok = await showDialog({ title: "Formel", body: h("div", { class: "form" }, h("label", null, "Linear: ", input), preview),
    buttons: [{ label: "OK", value: true, primary: true }, { label: "Abbrechen", value: false }], cancel: false });
  return ok ? buildLinear(input.value) : null;
}

/** Removing a shape: an equation goes with its whole mc:AlternateContent (else the Fallback would show). */
export function removeShapeEl(el: Element): void {
  const choice = el.parentElement;
  if (choice?.namespaceURI === NS_MC && choice.parentElement?.namespaceURI === NS_MC) choice.parentElement.remove();
  else el.remove();
}

/** Writes a new tree into a shape's equation (Choice) and its Fallback text, in one undo step. */
function writeEquation(app: PresentApp, id: number, f: (old: Seq) => Seq | null): void {
  const ed = app.editor;
  const sp = ed.shapeEl(id);
  const m = equationOf(sp);
  if (!sp || !m) return;
  const next = f(ommlToSeq(m) ?? []);
  if (!next) return;
  ed.edit("Formel", () => {
    const a14m = m.parentElement!;
    a14m.replaceChildren(PptxDom.fragment(a14m.ownerDocument, slideOmml(next))[0]);
    const ac = sp.parentElement?.parentElement;
    const t = ac?.namespaceURI === NS_MC ? ac.getElementsByTagNameNS(NS_MC, "Fallback")[0]?.getElementsByTagNameNS(PptxDom.NS_A, "t")[0] : undefined;
    if (t) t.textContent = seqToLinear(next);
  });
  ed.updateAutofit(id);
  ed.render();
}

/** The selected equation's shape id, if one equation is selected. */
export function selectedEquation(app: PresentApp): number | null {
  const vs = app.editor.selectedViews;
  return vs.length === 1 && hasEquation(vs[0]) ? vs[0].id : null;
}

/** A structure or symbol from the tab's galleries goes at the end of the selected equation. */
function appendTo(app: PresentApp, seq: Seq): void {
  const id = selectedEquation(app);
  if (id === null) { void insertEquation(app, seq); return; }
  writeEquation(app, id, (old) => [...old, ...JSON.parse(JSON.stringify(seq))]);
}

async function latexDialog(app: PresentApp): Promise<void> {
  const id = selectedEquation(app);
  if (id === null) return;
  const m = equationOf(app.editor.shapeEl(id));
  const input = h("input", { type: "text", value: seqToLatex(m ? ommlToSeq(m) ?? [] : []), style: "width:420px;font:14px Consolas, monospace" }) as HTMLInputElement;
  setTimeout(() => input.focus(), 30);
  const ok = await showDialog({ title: "Formel", body: h("div", { class: "form" }, h("label", null, "LaTeX: ", input)),
    buttons: [{ label: "OK", value: true, primary: true }, { label: "Abbrechen", value: false }], cancel: false });
  if (ok) writeEquation(app, id, () => parseLatex(input.value));
}

const STRUCTS: [string, string, string][] = [["eq.fraction", "math.fraction", "B"], ["eq.script", "math.script", "T"], ["eq.radical", "math.radical", "W"], ["eq.integral", "math.integral", "I"],
  ["eq.nary", "math.nary", "R"], ["eq.bracket", "math.bracket", "C"], ["eq.function", "math.function", "U"], ["eq.accent", "math.accent", "K"],
  ["eq.limit", "math.limit", "G"], ["eq.operator", "math.operator", "E"], ["eq.matrix", "math.matrix", "M"]];

/** Formel, the contextual tab of a selected equation, with PowerPoint's groups, keytips and entry names. */
export function equationTab(app: PresentApp): Tab {
  const b = (cmd: string, key: string, size: "large" | "small" = "small"): Control => ({ type: "button", cmd, key, size });
  const dd = (cmd: string, key: string, menu: () => MenuEntry[]): Control => ({ type: "dropdown", cmd, key, size: "large", get menu() { return menu(); } } as Control);
  const structMenu = (gal: string) => (): MenuEntry[] => (MATH_GALLERIES[gal] ?? []).flatMap((g): MenuEntry[] => [
    ...(g.heading ? [{ heading: g.heading } as MenuEntry] : []),
    ...g.items.map((it) => ({ label: `Formelstruktur ${it.label}`, run: () => appendTo(app, it.seq) })),
  ]);
  const symbols = (): MenuEntry[] => (SYMBOL_CATEGORIES[0]?.groups ?? []).flatMap((g) => g.items.map(([ch, name]) => ({ label: `Formelsymbol ${name}`, run: () => appendTo(app, [{ k: "t", s: ch }]) })));
  return {
    id: "equation", label: "Formel", key: "JE", contextual: true, shrink: [],
    groups: [
      { id: "tools", label: "Tools", key: "ZT", icon: "equation", controls: [dd("eq.builtin", "A", () => BUILT_IN.map((e) => ({ label: e.name, run: () => void insertEquation(app, e.seq) }))), b("eq.ink", "H1", "large")] },
      { id: "conv", label: "Konvertierungen", key: "ZK", icon: "math-convert", controls: [
        { type: "stack", items: [b("eq.latexToMath", "J"), b("eq.mathToLatex", "H2")] },
        { type: "stack", items: [b("eq.professional", "P"), b("eq.linear", "L"), b("eq.text", "N")] }, b("eq.options", "F")] },
      { id: "symbols", label: "Symbole", key: "ZS", icon: "math-symbols", controls: [dd("eq.symbols", "O", symbols)] },
      { id: "struct", label: "Strukturen", key: "ZU", icon: "math-fraction", controls: STRUCTS.map(([cmd, gal, key]) => dd(cmd, key, structMenu(gal))) },
    ],
  };
}

/** The tab's commands. */
export function implementEquationCommands(app: PresentApp): void {
  const on = () => selectedEquation(app) !== null;
  const impl = (id: string, run: () => void | Promise<void>, enabled = on) => app.commands.implement(id, { run, enabled });
  impl("eq.builtin", () => {}, () => !!app.slidePath());
  impl("eq.latexToMath", () => latexDialog(app));
  impl("eq.mathToLatex", () => latexDialog(app));
  impl("eq.professional", () => {});
  impl("eq.linear", () => { const id = selectedEquation(app); if (id !== null) return editEquation(app, id); });
  impl("eq.symbols", () => {}, () => !!app.slidePath());
  for (const [cmd] of STRUCTS) impl(cmd, () => {}, () => !!app.slidePath());
}
