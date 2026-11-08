// The chart Format pane's Textoptionen (Textfüllung und Kontur, Texteffekte, Textfeld) and the chart
// area's Größe und Eigenschaften (D-066, D-111).

import type { ChartModel, ChartObjectProps, Effects, Fill, Line, TextBox, TextFormat } from "@raabe/charts";
import type { PaneControl, PaneSection, PaneTab } from "@raabe/shell";
import { effectSections, fillSection, lineSection, type FmtTarget } from "./panefill";
import type { PaneHost } from "./formatpane";

/** Textoptionen: three tabs as in Excel. `kind` decides what the Textfeld offers (axes: no margins). */
export function textSections(t: FmtTarget, kind: string): PaneTab[] {
  const text = () => (t.get() as { text?: TextFormat } | undefined)?.text;
  const setText = (fn: (x: TextFormat) => void) => t.change((f) => {
    const holder = f as { text?: TextFormat };
    const x: TextFormat = { ...holder.text };
    fn(x);
    for (const k of Object.keys(x) as (keyof TextFormat)[]) if (x[k] === undefined) delete x[k];
    if (Object.keys(x).length) holder.text = x; else delete holder.text;
  });
  // The text's fill, outline and effects seen as an element of their own for the shared sections.
  const asShape: FmtTarget = {
    key: t.key + "/text",
    get: () => {
      const x = text();
      if (!x) return undefined;
      const fill: Fill | undefined = x.fill ?? (x.color ? { type: "solid", color: x.color } : undefined);
      return { fill, line: x.outline, effects: x.effects };
    },
    change: (fn) => setText((x) => {
      const shape: { fill?: Fill; line?: Line; effects?: Effects } = { fill: x.fill ?? (x.color ? { type: "solid", color: x.color } : undefined), line: x.outline, effects: x.effects };
      fn(shape);
      // Einfarbige Füllung stays the text's colour (Start > Schriftfarbe), other fills its fill.
      if (shape.fill?.type === "solid" && !shape.fill.transparency) { x.color = shape.fill.color; delete x.fill; }
      else { delete x.color; x.fill = shape.fill; }
      x.outline = shape.line;
      x.effects = shape.effects;
    }),
  };
  const fillTab: PaneSection[] = [
    { ...fillSection(asShape, "#595959"), title: "Textfüllung" },
    { ...lineSection(asShape, { title: "Textkontur", autoColor: "#595959", autoWidth: 0.75 }) },
    { title: "Schrift", controls: () => [
      { type: "number", label: "Schriftgrad", unit: " Pt.", min: 1, max: 409, step: 1, value: () => text()?.size ?? null, set: (v) => setText((x) => { x.size = v; }) },
      { type: "check", label: "Fett", value: () => !!text()?.bold, set: (on) => setText((x) => { x.bold = on || undefined; }) },
      { type: "check", label: "Kursiv", value: () => !!text()?.italic, set: (on) => setText((x) => { x.italic = on || undefined; }) },
      { type: "check", label: "Unterstrichen", value: () => !!text()?.underline, set: (on) => setText((x) => { x.underline = on || undefined; }) },
      { type: "check", label: "Durchgestrichen", value: () => !!text()?.strike, set: (on) => setText((x) => { x.strike = on || undefined; }) },
    ] },
  ];
  const box = () => text()?.box;
  const setBox = (patch: Partial<TextBox>) => setText((x) => {
    const b: TextBox = { ...x.box, ...patch };
    for (const k of Object.keys(b) as (keyof TextBox)[]) if (b[k] === undefined) delete b[k];
    x.box = Object.keys(b).length ? b : undefined;
  });
  const axis = kind === "valAxis" || kind === "catAxis" || kind === "valAxis2";
  const boxControls = (): PaneControl[] => {
    const c: PaneControl[] = [
      { type: "select", label: "Vertikale Ausrichtung", options: [{ label: "Oben", value: "t" }, { label: "Mitte", value: "ctr" }, { label: "Unten", value: "b" }],
        value: () => box()?.vAlign ?? "ctr", set: (v) => setBox({ vAlign: v === "ctr" ? undefined : (v as TextBox["vAlign"]) }) },
      { type: "select", label: "Textrichtung", options: [{ label: "Horizontal", value: "horz" }, { label: "Gesamten Text um 90° drehen", value: "vert" }, { label: "Gesamten Text um 270° drehen", value: "vert270" }, { label: "Gestapelt", value: "wordArtVert" }],
        value: () => box()?.dir ?? "horz", set: (v) => setBox({ dir: v === "horz" ? undefined : (v as TextBox["dir"]) }) },
      { type: "number", label: "Benutzerdefinierter Winkel", unit: "°", min: -90, max: 90, value: () => box()?.rotation ?? 0, set: (v) => setBox({ rotation: Math.round(v) || undefined }), disabled: () => !!box()?.dir && box()?.dir !== "horz" },
    ];
    if (!axis) {
      const mg = box()?.margins ?? { l: 2.88, t: 1.44, r: 2.88, b: 1.44 };
      const side = (label: string, k: keyof typeof mg): PaneControl => ({ type: "number", label, unit: " cm", min: 0, max: 50, step: 0.05, decimals: 2,
        value: () => Math.round((mg[k] / 72) * 2.54 * 100) / 100, set: (v) => setBox({ margins: { ...mg, [k]: Math.round((v / 2.54) * 72 * 100) / 100 } }) });
      c.push(side("Linker Rand", "l"), side("Rechter Rand", "r"), side("Oberer Rand", "t"), side("Unterer Rand", "b"));
      c.push({ type: "check", label: "Text in Form umbrechen", value: () => box()?.wrap !== false, set: (on) => setBox({ wrap: on ? undefined : false }) });
    }
    return c;
  };
  return [
    { id: "text", label: "Textfüllung und Kontur", icon: "font-color", sections: () => fillTab },
    { id: "texteffects", label: "Texteffekte", icon: "text-effects", sections: () => effectSections(asShape, { autoColor: "#595959", soft: false, format3d: false }) },
    { id: "textbox", label: "Textfeld", icon: "pane-layout", sections: () => [{ title: "Textfeld", open: true, controls: boxControls }] },
  ];
}

/** Größe und Eigenschaften of the chart area: Größe, Eigenschaften, Alternativtext. */
export function sizeSections(m: ChartModel, host: PaneHost): PaneSection[] {
  const props = () => m.props;
  const setProps = (patch: Partial<ChartObjectProps>) => host.edit((c) => {
    const p: ChartObjectProps = { ...c.props, ...patch };
    for (const k of Object.keys(p) as (keyof ChartObjectProps)[]) if (p[k] === undefined) delete p[k];
    if (Object.keys(p).length) c.props = p; else delete c.props;
  });
  const placement = () => props()?.placement ?? "twoCell";
  return [
    { title: "Größe", open: true, controls: () => [
      { type: "number", label: "Höhe", unit: " cm", min: 0.1, max: 140, step: 0.1, decimals: 2, value: () => host.size().h,
        set: (v) => { const s = host.size(); host.resize(props()?.lockAspect ? Math.round(s.w * (v / s.h) * 100) / 100 : s.w, v); } },
      { type: "number", label: "Breite", unit: " cm", min: 0.1, max: 140, step: 0.1, decimals: 2, value: () => host.size().w,
        set: (v) => { const s = host.size(); host.resize(v, props()?.lockAspect ? Math.round(s.h * (v / s.w) * 100) / 100 : s.h); } },
      { type: "number", label: "Drehung", unit: "°", value: () => 0, set: () => undefined, disabled: () => true },
      { type: "number", label: "Höhe skalieren", unit: " %", value: () => 100, set: () => undefined, disabled: () => true },
      { type: "number", label: "Breite skalieren", unit: " %", value: () => 100, set: () => undefined, disabled: () => true },
      { type: "check", label: "Seitenverhältnis sperren", value: () => !!props()?.lockAspect, set: (on) => setProps({ lockAspect: on || undefined }) },
    ] },
    { title: "Eigenschaften", controls: () => [
      { type: "radio", label: "Von Zellposition und -größe abhängig", checked: () => placement() === "twoCell", pick: () => setProps({ placement: undefined }) },
      { type: "radio", label: "Von Zellposition abhängig, nicht von Größe", checked: () => placement() === "oneCell", pick: () => setProps({ placement: "oneCell" }) },
      { type: "radio", label: "Von Zellposition und -größe unabhängig", checked: () => placement() === "absolute", pick: () => setProps({ placement: "absolute" }) },
      { type: "check", label: "Objekt drucken", value: () => props()?.print !== false, set: (on) => setProps({ print: on ? undefined : false }) },
      { type: "check", label: "Gesperrt", value: () => props()?.locked !== false, set: (on) => setProps({ locked: on ? undefined : false }) },
    ] },
    { title: "Alternativtext", controls: () => [
      { type: "text", label: "Titel", value: () => props()?.altTitle ?? "", set: (v) => setProps({ altTitle: v || undefined }) },
      { type: "text", label: "Beschreibung", value: () => props()?.altText ?? "", set: (v) => setProps({ altText: v || undefined }) },
    ] },
  ];
}
