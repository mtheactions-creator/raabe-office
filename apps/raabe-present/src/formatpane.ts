// "Form formatieren" / "Grafik formatieren" (D-070): the task pane docked at the right, as in
// PowerPoint, with Formoptionen (Füllung und Linie; Größe und Eigenschaften with Größe, Position and
// Textfeld: vertical alignment, Autoanpassen, inner margins, wrapping) and Textoptionen. Changes apply
// at once, each one undo step, to every selected shape.

import { PptxDom, type ShapeView } from "@raabe/ooxml";
import { TaskPane, type PaneControl, type PaneSection, type PaneSpec } from "@raabe/shell";
import type { SlideEditor } from "./editor";
import * as S from "./shapeops";
import * as T from "./textops";
import { effectsTab } from "./effectpane";

const { kid, kids, make, fragment } = PptxDom;
const FILLS = ["noFill", "solidFill", "gradFill", "blipFill", "pattFill", "grpFill"];

/** The runs' look for the pane's Textoptionen (the app formats the selected text, or all of a selected shape's text). */
export interface TextLook {
  fill(): { kind: "auto" | "none" | "solid"; color?: string };
  setFill(hex: string | null | "none"): void;
  outline(): { color: string | null; width: number } | null;
  setOutline(o: { color?: string | null; width?: number }): void;
}

export class FormatPane {
  readonly pane: TaskPane;
  /** Set by the app: Textoptionen (Textfüllung, Textkontur) work on the runs through it. */
  text: TextLook | null = null;
  constructor(private ed: SlideEditor, host: HTMLElement) {
    this.pane = new TaskPane(host);
  }

  get isOpen(): boolean { return this.pane.isOpen; }

  private views(): ShapeView[] { return this.ed.selectedViews; }

  private spPrs(fn: (spPr: Element, el: Element, v: ShapeView) => void, label: string): void {
    const views = this.views();
    this.ed.edit(label, () => {
      for (const v of views) {
        const el = this.ed.shapeEl(v.id);
        const sp = el && kid(el, v.kind === "grpSp" ? "grpSpPr" : "spPr");
        if (sp && el) fn(sp, el, v);
      }
    });
    this.pane.refresh();
  }

  private setFill(kind: "none" | "solid" | "auto", hex?: string, alpha?: number): void {
    this.spPrs((sp, el) => {
      for (const c of kids(sp)) if (FILLS.includes(c.localName)) c.remove();
      if (kind === "auto") return;
      const f = kind === "none" ? make(el.ownerDocument, "a:noFill") : fragment(el.ownerDocument, `<a:solidFill><a:srgbClr val="${hex ?? "4472C4"}">${alpha ? `<a:alpha val="${Math.round((1 - alpha) * 100000)}"/>` : ""}</a:srgbClr></a:solidFill>`)[0];
      const after = kids(sp).find((c) => c.localName === "prstGeom" || c.localName === "custGeom") ?? kid(sp, "xfrm");
      sp.insertBefore(f, after ? after.nextSibling : sp.firstChild);
    }, "Füllung");
  }

  private line(fn: (ln: Element, doc: Document) => void): void {
    this.spPrs((sp, el) => {
      let ln = kid(sp, "ln");
      if (!ln) { ln = make(el.ownerDocument, "a:ln"); sp.insertBefore(ln, kids(sp).find((c) => ["effectLst", "effectDag", "scene3d", "sp3d", "extLst"].includes(c.localName)) ?? null); }
      fn(ln, el.ownerDocument);
    }, "Linie");
  }

  private bodyPr(fn: (bp: Element) => void, label: string): void {
    const views = this.views();
    this.ed.edit(label, () => { for (const v of views) { const el = this.ed.shapeEl(v.id); const bp = el && S.bodyPrOf(el); if (bp) fn(bp); } });
    for (const v of views) this.ed.updateAutofit(v.id);
    this.ed.render();
    this.pane.refresh();
  }

  spec(): PaneSpec {
    const v = () => this.views()[0];
    const pic = v()?.kind === "pic";
    const fillSec: PaneSection = { title: "Füllung", open: true, controls: () => {
      const f = v()?.format.fill;
      const c: PaneControl[] = [
        { type: "radio", label: "Keine Füllung", checked: () => f?.kind === "none", pick: () => this.setFill("none") },
        { type: "radio", label: "Einfarbige Füllung", checked: () => f?.kind === "solid", pick: () => this.setFill("solid", f?.kind === "solid" ? f.color : "4472C4") },
        { type: "radio", label: "Automatisch", checked: () => !f, pick: () => this.setFill("auto") },
      ];
      if (f?.kind === "solid") {
        c.push({ type: "color", label: "Farbe", value: () => `#${f.color}`, set: (hex) => this.setFill("solid", (hex ?? "#4472C4").replace("#", "").toUpperCase(), f.transparency) });
        c.push({ type: "number", label: "Transparenz", unit: " %", min: 0, max: 100, slider: true, value: () => Math.round((f.transparency ?? 0) * 100), set: (t) => this.setFill("solid", f.color, t / 100) });
      }
      return c;
    } };
    const lineSec: PaneSection = { title: "Linie", open: v()?.kind === "cxnSp", controls: () => {
      const l = v()?.format.line;
      const c: PaneControl[] = [
        { type: "radio", label: "Keine Linie", checked: () => l === null, pick: () => this.line((ln, doc) => { for (const x of kids(ln)) if (["noFill", "solidFill", "gradFill", "pattFill"].includes(x.localName)) x.remove(); ln.insertBefore(make(doc, "a:noFill"), ln.firstChild); }) },
        { type: "radio", label: "Einfarbige Linie", checked: () => !!l, pick: () => this.line((ln, doc) => { for (const x of kids(ln)) if (["noFill", "solidFill", "gradFill", "pattFill"].includes(x.localName)) x.remove(); ln.insertBefore(fragment(doc, `<a:solidFill><a:srgbClr val="${l?.color ?? "000000"}"/></a:solidFill>`)[0], ln.firstChild); }) },
      ];
      if (l) {
        c.push({ type: "color", label: "Farbe", value: () => `#${l.color}`, set: (hex) => this.line((ln, doc) => { for (const x of kids(ln)) if (["noFill", "solidFill", "gradFill", "pattFill"].includes(x.localName)) x.remove(); ln.insertBefore(fragment(doc, `<a:solidFill><a:srgbClr val="${(hex ?? "#000000").replace("#", "").toUpperCase()}"/></a:solidFill>`)[0], ln.firstChild); }) });
        c.push({ type: "number", label: "Breite", unit: " Pt.", min: 0, max: 1584, step: 0.25, decimals: 2, value: () => l.width, set: (w) => this.line((ln) => ln.setAttribute("w", String(Math.round(w * 12700)))) });
        c.push({ type: "select", label: "Strichtyp", options: [["solid", "Durchgehend"], ["sysDot", "Rundpunkt"], ["sysDash", "Quadratpunkt"], ["dash", "Strich"], ["dashDot", "Strich-Punkt"], ["lgDash", "Langer Strich"], ["lgDashDot", "Langer Strich-Punkt"], ["lgDashDotDot", "Langer Strich-Punkt-Punkt"]].map(([value, label]) => ({ value, label })),
          value: () => l.dash ?? "solid", set: (d) => this.line((ln, doc) => { kid(ln, "prstDash")?.remove(); const after = kids(ln).find((x) => ["noFill", "solidFill", "gradFill", "pattFill"].includes(x.localName)); ln.insertBefore(make(doc, "a:prstDash", { val: d }), after ? after.nextSibling : ln.firstChild); }) });
      }
      return c;
    } };
    const cm = (emu: number) => Math.round((emu / 360000) * 100) / 100;
    const sizeSec: PaneSection = { title: "Größe", open: true, controls: () => [
      { type: "number", label: "Höhe", unit: " cm", min: 0, max: 142, step: 0.1, decimals: 2, value: () => cm(v()?.h ?? 0), set: (x) => this.place({ h: x * 360000 }) },
      { type: "number", label: "Breite", unit: " cm", min: 0, max: 142, step: 0.1, decimals: 2, value: () => cm(v()?.w ?? 0), set: (x) => this.place({ w: x * 360000 }) },
      { type: "number", label: "Drehung", unit: "°", min: -3600, max: 3600, value: () => Math.round(v()?.rot ?? 0), set: (x) => this.place({ rot: x }) },
    ] };
    const posSec: PaneSection = { title: "Position", open: false, controls: () => [
      { type: "number", label: "Horizontale Position", unit: " cm", min: -142, max: 142, step: 0.1, decimals: 2, value: () => cm(v()?.x ?? 0), set: (x) => this.place({ x: x * 360000 }) },
      { type: "number", label: "Vertikale Position", unit: " cm", min: -142, max: 142, step: 0.1, decimals: 2, value: () => cm(v()?.y ?? 0), set: (x) => this.place({ y: x * 360000 }) },
    ] };
    const textSec: PaneSection = { title: "Textfeld", open: true, controls: () => {
      const b = v()?.text?.body;
      const ins = (name: "lIns" | "rIns" | "tIns" | "bIns", label: string, emu: number): PaneControl =>
        ({ type: "number", label, unit: " cm", min: 0, max: 50, step: 0.05, decimals: 2, value: () => cm(emu), set: (x) => this.bodyPr((bp) => bp.setAttribute(name, String(Math.round(x * 360000))), "Seitenrand") });
      return [
        { type: "select", label: "Vertikale Ausrichtung", options: [{ value: "t", label: "Oben" }, { value: "ctr", label: "Mitte" }, { value: "b", label: "Unten" }], value: () => b?.anchor ?? "t", set: (a) => this.bodyPr((bp) => bp.setAttribute("anchor", a), "Ausrichtung") },
        { type: "select", label: "Textrichtung", options: [{ value: "horz", label: "Horizontal" }, { value: "vert", label: "Alle Texte um 90° drehen" }, { value: "vert270", label: "Alle Texte um 270° drehen" }, { value: "wordArtVert", label: "Gestapelt" }], value: () => b?.vert ?? "horz", set: (d) => this.bodyPr((bp) => T.setAttr(bp, "vert", d === "horz" ? null : d), "Textrichtung") },
        { type: "heading", label: "Autoanpassen" },
        ...(["none", "norm", "shape"] as const).map((k): PaneControl => ({ type: "radio", label: { none: "Nicht automatisch anpassen", norm: "Text bei Überlauf verkleinern", shape: "Größe der Form dem Text anpassen" }[k], checked: () => b?.autofit === k,
          pick: () => this.bodyPr((bp) => { for (const x of kids(bp)) if (["noAutofit", "normAutofit", "spAutoFit"].includes(x.localName)) x.remove(); bp.insertBefore(make(bp.ownerDocument, k === "none" ? "a:noAutofit" : k === "norm" ? "a:normAutofit" : "a:spAutoFit"), kids(bp).find((x) => ["scene3d", "sp3d", "flatTx", "extLst"].includes(x.localName)) ?? null); }, "Autoanpassen") })),
        ins("lIns", "Linker Seitenrand", b?.l ?? 91440), ins("rIns", "Rechter Seitenrand", b?.r ?? 91440), ins("tIns", "Oberer Seitenrand", b?.t ?? 45720), ins("bIns", "Unterer Seitenrand", b?.b ?? 45720),
        { type: "check", label: "Text in Form umbrechen", value: () => b?.wrap ?? true, set: (on) => this.bodyPr((bp) => bp.setAttribute("wrap", on ? "square" : "none"), "Umbrechen") },
      ];
    } };
    const tabs = [
      { id: "fill", label: "Füllung und Linie", icon: "pane-fill-line", sections: () => (v()?.kind === "cxnSp" ? [lineSec] : [fillSec, lineSec]) },
      effectsTab(() => { const x = v(); const el = x && this.ed.shapeEl(x.id); return el ? kid(el, x.kind === "grpSp" ? "grpSpPr" : "spPr") ?? undefined : undefined; }, (fn, label) => this.spPrs((sp) => fn(sp), label)),
      { id: "size", label: "Größe und Eigenschaften", icon: "pane-layout", sections: () => (v()?.text ? [sizeSec, posSec, textSec] : [sizeSec, posSec]) },
    ];
    // Textoptionen (by hand, D-137: Texteffekte formatieren... and the WordArt-Formate launcher had nothing to open):
    // Textfüllung and Textkontur of the text, and the Textfeld settings.
    const t = this.text;
    const textSets = !pic && v()?.text && t ? [{ label: "Textoptionen", tabs: [
      { id: "textFill", label: "Textfüllung und -kontur", icon: "pane-fill-line", sections: () => [
        { title: "Textfüllung", open: true, controls: (): PaneControl[] => {
          const f = t.fill();
          const c: PaneControl[] = [
            { type: "radio", label: "Keine Füllung", checked: () => f.kind === "none", pick: () => { t.setFill("none"); this.pane.refresh(); } },
            { type: "radio", label: "Einfarbige Füllung", checked: () => f.kind === "solid", pick: () => { t.setFill(f.color ?? "000000"); this.pane.refresh(); } },
            { type: "radio", label: "Automatisch", checked: () => f.kind === "auto", pick: () => { t.setFill(null); this.pane.refresh(); } },
          ];
          if (f.kind === "solid") c.push({ type: "color", label: "Farbe", value: () => `#${f.color ?? "000000"}`, set: (hex) => { t.setFill((hex ?? "#000000").replace("#", "").toUpperCase()); this.pane.refresh(); } });
          return c;
        } },
        { title: "Textkontur", open: true, controls: (): PaneControl[] => {
          const o = t.outline();
          const c: PaneControl[] = [
            { type: "radio", label: "Keine Linie", checked: () => !o || o.color === null, pick: () => { t.setOutline({ color: null }); this.pane.refresh(); } },
            { type: "radio", label: "Einfarbige Linie", checked: () => !!o && o.color !== null, pick: () => { t.setOutline({ color: o?.color ?? "000000", width: o?.width || 9525 }); this.pane.refresh(); } },
          ];
          if (o && o.color !== null) {
            c.push({ type: "color", label: "Farbe", value: () => `#${o.color}`, set: (hex) => { t.setOutline({ color: (hex ?? "#000000").replace("#", "").toUpperCase() }); this.pane.refresh(); } });
            c.push({ type: "number", label: "Breite", unit: " Pt.", min: 0, max: 1584, step: 0.25, decimals: 2, value: () => o.width / 12700, set: (w) => { t.setOutline({ width: Math.round(w * 12700) }); this.pane.refresh(); } });
          }
          return c;
        } },
      ] },
      { id: "textBox", label: "Textfeld", icon: "pane-layout", sections: () => [textSec] },
    ] }] : [];
    return { title: pic ? "Grafik formatieren" : "Form formatieren", sets: [{ label: pic ? "Grafik" : "Formoptionen", tabs }, ...textSets] };
  }

  private place(p: Partial<S.Xfrm>): void {
    const views = this.views();
    this.ed.edit("Größe und Position", () => {
      for (const v of views) {
        const el = this.ed.shapeEl(v.id);
        if (el) S.setXfrm(el, { x: v.x, y: v.y, w: v.w, h: v.h, ...p });
      }
    });
    this.pane.refresh();
  }

  open(): void {
    if (!this.views().length) return;
    this.pane.show(this.spec());
  }

  /** Opens the pane at Effekte with one section open (Schattenoptionen..., 3D-Optionen... and so on). */
  openEffects(section: string): void {
    if (!this.views().length) return;
    this.pane.showAt(this.spec(), "effects", section);
  }

  /** Texteffekte formatieren...: the pane at Textoptionen. */
  openText(): void {
    if (!this.views().length) return;
    const spec = this.spec();
    if (spec.sets.length > 1) this.pane.showAt(spec, "textFill"); else this.pane.show(spec);
  }

  /** The selection changed: the open pane shows the new shape (or closes when nothing is selected). */
  sync(): void {
    if (!this.pane.isOpen) return;
    if (!this.views().length) this.pane.close();
    else this.pane.show(this.spec());
  }
}
