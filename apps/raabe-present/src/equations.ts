// Formeln in RaabePresent (D-072): an equation is a text box whose paragraph holds Office Math (a14:m with
// m:oMathPara), inside mc:AlternateContent with a plain-text Fallback for readers before PowerPoint 2010,
// as PowerPoint writes it. Drawn as MathML from the equation tree (packages/ooxml/src/math.ts, shared with
// RaabeWrite). Editing happens in PowerPoint's linear format (UnicodeMath, "Aktuell – Linear"): the
// equation's linear text in a box with the professional form shown live; OK writes it back as OMML.

import { h, showDialog } from "@raabe/shell";
import { ommlToSeq, seqToOmml, PptxDom, type ShapeView } from "@raabe/ooxml";
import { buildLinear, seqToLinear, seqToMathml, type Seq } from "@raabe/text-engine";
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
