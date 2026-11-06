// Suchen (Strg+F) and Ersetzen (Strg+H) as PowerPoint's dialogs (D-070): Suchen nach, Ersetzen durch,
// Groß-/Kleinschreibung beachten, Nur ganzes Wort suchen; Weitersuchen goes through the slides' text
// from the caret on and selects the hit in its shape; Ersetzen replaces the selected hit and finds the
// next, Alle ersetzen replaces every hit in one undo step and says how many (our own wording).

import { h, messageBox } from "@raabe/shell";
import { spTreeOf } from "@raabe/ooxml";
import type { PresentApp } from "./app";
import * as S from "./shapeops";
import * as T from "./textops";

interface Hit { slide: number; id: number; p: number; o: number; len: number }

function textShapes(tree: Element): Element[] {
  const out: Element[] = [];
  for (const s of S.shapeEls(tree)) {
    if (s.localName === "grpSp") out.push(...textShapes(s));
    else if (S.txBodyOf(s)) out.push(s);
  }
  return out;
}

export class FindDialog {
  private el: HTMLElement | null = null;
  private what!: HTMLInputElement;
  private with!: HTMLInputElement;
  private matchCase!: HTMLInputElement;
  private whole!: HTMLInputElement;

  constructor(private app: PresentApp) {}

  private regex(): RegExp | null {
    const w = this.what.value;
    if (!w) return null;
    const esc = w.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
    return new RegExp(this.whole.checked ? `(?<![\\p{L}\\p{N}])${esc}(?![\\p{L}\\p{N}])` : esc, `gu${this.matchCase.checked ? "" : "i"}`);
  }

  /** Every hit of the presentation in order: slide, shape (z-order), paragraph, offset. */
  private hits(): Hit[] {
    const re = this.regex();
    if (!re) return [];
    const out: Hit[] = [];
    this.app.deck.slides.forEach((ref, si) => {
      for (const s of textShapes(spTreeOf(this.app.deck.pptx.doc(ref.path))!)) {
        T.paras(S.txBodyOf(s)!).forEach((p, pi) => {
          for (const m of T.paraText(p).matchAll(re)) out.push({ slide: si, id: S.shapeId(s), p: pi, o: m.index!, len: m[0].length });
        });
      }
    });
    return out;
  }

  private findNext(): Hit | null {
    const hits = this.hits();
    if (!hits.length) return null;
    const ed = this.app.editor;
    const cur = this.app.current();
    const t = ed.text;
    const after = (h2: Hit) => {
      if (h2.slide !== cur) return h2.slide > cur;
      if (!t) return true;
      const tree = spTreeOf(this.app.deck.pptx.doc(this.app.deck.slides[cur].path))!;
      const order = textShapes(tree).map((x) => S.shapeId(x));
      const a = order.indexOf(h2.id), b = order.indexOf(t.id);
      if (a !== b) return a > b;
      const end = T.maxPos(t.caret, t.anchor);
      return h2.p > end.p || (h2.p === end.p && h2.o >= end.o);
    };
    return hits.find(after) ?? hits[0];
  }

  private show(hit: Hit): void {
    if (hit.slide !== this.app.current()) this.app.setCurrent(hit.slide);
    const ed = this.app.editor;
    ed.enterText(hit.id, { p: hit.p, o: hit.o + hit.len });
    if (ed.text) { ed.text.anchor = { p: hit.p, o: hit.o }; ed.render(); }
  }

  next(): void {
    const hit = this.findNext();
    if (!hit) { void messageBox("RaabePresent", `„${this.what.value}“ wurde nicht gefunden.`); return; }
    this.show(hit);
  }

  replace(): void {
    const ed = this.app.editor;
    const t = ed.text;
    const re = this.regex();
    if (t && re && T.cmpPos(t.caret, t.anchor) !== 0) {
      const body = ed.textBody(t.id)!;
      const sel = T.rangeText(body, t.anchor, t.caret);
      re.lastIndex = 0;
      const m = re.exec(sel);
      if (m && m[0] === sel) ed.typeText(this.with.value);
    }
    this.next();
  }

  replaceAll(): void {
    const re = this.regex();
    if (!re) return;
    const deck = this.app.deck;
    let n = 0;
    deck.change("Alle ersetzen", deck.slides.map((s) => s.path), () => {
      for (const ref of deck.slides) {
        for (const s of textShapes(spTreeOf(deck.pptx.doc(ref.path))!)) {
          const body = S.txBodyOf(s)!;
          T.paras(body).forEach((p, pi) => {
            // From the end so earlier offsets stay right.
            const ms = [...T.paraText(p).matchAll(re)].reverse();
            for (const m of ms) {
              T.deleteRange(body, { p: pi, o: m.index! }, { p: pi, o: m.index! + m[0].length });
              if (this.with.value) T.insertText(body, { p: pi, o: m.index! }, this.with.value);
              n++;
            }
          });
        }
      }
    });
    void messageBox("RaabePresent", n ? `Es wurden ${n} Ersetzungen vorgenommen.` : `„${this.what.value}“ wurde nicht gefunden.`);
  }

  open(replace: boolean): void {
    this.el?.remove();
    const ed = this.app.editor;
    const t = ed.text;
    const pre = t && T.cmpPos(t.caret, t.anchor) && t.caret.p === t.anchor.p ? T.rangeText(ed.textBody(t.id)!, t.anchor, t.caret) : "";
    this.what = h("input", { type: "text", value: pre || this.what?.value || "", style: "width:260px" }) as HTMLInputElement;
    this.with = h("input", { type: "text", value: this.with?.value ?? "", style: "width:260px" }) as HTMLInputElement;
    this.matchCase = h("input", { type: "checkbox" }) as HTMLInputElement;
    this.whole = h("input", { type: "checkbox" }) as HTMLInputElement;
    const close = () => { this.el?.remove(); this.el = null; this.app.editor.focus(); };
    const btn = (label: string, run: () => void, primary = false) => h("button", { class: primary ? "primary" : "", onclick: run }, label);
    const box = h("div", { class: "dialog find-dialog", role: "dialog", "aria-label": replace ? "Ersetzen" : "Suchen" },
      h("div", { class: "dlg-title" }, h("span", { class: "t" }, replace ? "Ersetzen" : "Suchen"), h("button", { "aria-label": "Schließen", onclick: close }, "✕")),
      h("div", { class: "find-body" },
        h("div", { class: "find-fields" },
          h("label", null, "Suchen nach:"), this.what,
          ...(replace ? [h("label", null, "Ersetzen durch:"), this.with] : []),
          h("label", { class: "dlg-check" }, this.matchCase, " Groß-/Kleinschreibung beachten"),
          h("label", { class: "dlg-check" }, this.whole, " Nur ganzes Wort suchen")),
        h("div", { class: "find-buttons" }, btn("Weitersuchen", () => this.next(), true), btn("Schließen", close),
          replace ? btn("Ersetzen", () => this.replace()) : btn("Ersetzen...", () => this.open(true)),
          ...(replace ? [btn("Alle ersetzen", () => this.replaceAll())] : []))));
    box.addEventListener("keydown", (e) => { if (e.key === "Escape") { e.preventDefault(); close(); } if (e.key === "Enter") { e.preventDefault(); this.next(); } e.stopPropagation(); });
    this.el = h("div", { class: "find-layer" }, box);
    document.body.append(this.el);
    this.what.focus();
    this.what.select();
  }
}
