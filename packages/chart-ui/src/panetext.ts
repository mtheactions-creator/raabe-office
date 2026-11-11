// The chart Format pane's Textoptionen (Textfüllung und -kontur, Texteffekte, Textfeld), the
// Ausrichtung of text elements and the chart area's Größe und Eigenschaften (D-066, D-111), with
// Excel's labels (results/2026-11-09-excel-chart-pane.json).

import type { ChartModel, ChartObjectProps, Effects, Fill, Line, TextBox, TextFormat } from "@raabe/charts";
import type { PaneControl, PaneSection, PaneTab } from "@raabe/shell";
import { effectSections, fillSection, lineSection, type FmtTarget } from "./panefill";
import type { PaneHost } from "./formatpane";

type TextHolder = { text?: TextFormat };

function textOf(t: FmtTarget): TextFormat | undefined {
  return (t.get() as TextHolder | undefined)?.text;
}

function editText(t: FmtTarget, fn: (x: TextFormat) => void): void {
  t.change((f) => {
    const holder = f as TextHolder;
    const x: TextFormat = { ...holder.text };
    fn(x);
    for (const k of Object.keys(x) as (keyof TextFormat)[]) if (x[k] === undefined) delete x[k];
    if (Object.keys(x).length) holder.text = x; else delete holder.text;
  });
}

const V_ALIGN = [
  { label: "Oben", value: "t" }, { label: "Mitte", value: "ctr" }, { label: "Unten", value: "b" },
  { label: "Oben zentriert", value: "t,c" }, { label: "Mitte zentriert", value: "ctr,c" }, { label: "Unten zentriert", value: "b,c" },
];
const DIRECTIONS = [{ label: "Horizontal", value: "horz" }, { label: "Text um 90° drehen", value: "vert" }, { label: "Text um 270° drehen", value: "vert270" }, { label: "Gestapelt", value: "wordArtVert" }];

/** The Ausrichtung / Textfeld block: Vertikale Ausrichtung, Textrichtung, Benutzerdefinierter Winkel. */
export function boxControls(t: FmtTarget): () => PaneControl[] {
  const box = () => textOf(t)?.box;
  const setBox = (patch: Partial<TextBox>) => editText(t, (x) => {
    const b: TextBox = { ...x.box, ...patch };
    for (const k of Object.keys(b) as (keyof TextBox)[]) if (b[k] === undefined) delete b[k];
    x.box = Object.keys(b).length ? b : undefined;
  });
  return () => [
    { type: "select", label: "Vertikale Ausrichtung", options: V_ALIGN, value: () => `${box()?.vAlign ?? "ctr"}${box()?.anchorCtr === false ? "" : ",c"}`,
      // The "zentriert" variants centre the text block horizontally too (the default); anchorCtr false stands for the others.
      set: (v) => { const [a, c] = v.split(","); setBox({ vAlign: a === "ctr" ? undefined : (a as TextBox["vAlign"]), anchorCtr: c ? undefined : false }); } },
    { type: "select", label: "Textrichtung", options: DIRECTIONS, value: () => box()?.dir ?? "horz", set: (v) => setBox({ dir: v === "horz" ? undefined : (v as TextBox["dir"]) }) },
    { type: "number", label: "Benutzerdefinierter Winkel", unit: "°", min: -90, max: 90, value: () => box()?.rotation ?? 0, set: (v) => setBox({ rotation: Math.round(v) || undefined }), disabled: () => !!box()?.dir && box()?.dir !== "horz" },
  ];
}

/** Textoptionen: three tabs as in Excel. */
export function textSections(t: FmtTarget, _kind: string): PaneTab[] {
  // The text's fill, outline and effects seen as an element of their own for the shared sections.
  const asShape: FmtTarget = {
    key: t.key + "/text",
    get: () => {
      const x = textOf(t);
      if (!x) return undefined;
      const fill: Fill | undefined = x.fill ?? (x.color ? { type: "solid", color: x.color } : undefined);
      return { fill, line: x.outline, effects: x.effects };
    },
    change: (fn) => editText(t, (x) => {
      const shape: { fill?: Fill; line?: Line; effects?: Effects } = { fill: x.fill ?? (x.color ? { type: "solid", color: x.color } : undefined), line: x.outline, effects: x.effects };
      fn(shape);
      // Einfarbige Füllung stays the text's colour (Start > Schriftfarbe), other fills its fill.
      if (shape.fill?.type === "solid" && !shape.fill.transparency) { x.color = shape.fill.color; delete x.fill; }
      else { delete x.color; x.fill = shape.fill; }
      x.outline = shape.line;
      x.effects = shape.effects;
    }),
  };
  const text = () => textOf(t);
  const fillTab: PaneSection[] = [
    { ...fillSection(asShape, "#595959", [], true, false), title: "Textfüllung" },
    { ...lineSection(asShape, { title: "Textkontur", autoColor: "#595959", autoWidth: 0.75, withAuto: false, arrows: "disabled" }) },
    { title: "Schrift", controls: () => [
      { type: "number", label: "Schriftgrad", unit: " Pt.", min: 1, max: 409, step: 1, value: () => text()?.size ?? null, set: (v) => editText(t, (x) => { x.size = v; }) },
      { type: "check", label: "Fett", value: () => !!text()?.bold, set: (on) => editText(t, (x) => { x.bold = on || undefined; }) },
      { type: "check", label: "Kursiv", value: () => !!text()?.italic, set: (on) => editText(t, (x) => { x.italic = on || undefined; }) },
      { type: "check", label: "Unterstrichen", value: () => !!text()?.underline, set: (on) => editText(t, (x) => { x.underline = on || undefined; }) },
      { type: "check", label: "Durchgestrichen", value: () => !!text()?.strike, set: (on) => editText(t, (x) => { x.strike = on || undefined; }) },
    ] },
  ];
  return [
    { id: "text", label: "Textfüllung und -kontur", icon: "font-color", sections: () => fillTab },
    { id: "texteffects", label: "Texteffekte", icon: "text-effects", sections: () => effectSections(asShape, { autoColor: "#4472C4", soft: false, format3d: false, reflection: true }) },
    { id: "textbox", label: "Textfeld", icon: "pane-layout", sections: () => [{ title: "Textfeld", open: true, controls: boxControls(t) }] },
  ];
}

/** Excel's default chart size in cm, the 100 % of Höhe / Breite skalieren. */
const DEFAULT_CM = { w: 12.7, h: 7.62 };

/** Größe und Eigenschaften of the chart area: Größe, Eigenschaften, Alternativtext. */
export function sizeSections(m: ChartModel, host: PaneHost): PaneSection[] {
  const props = () => m.props;
  const setProps = (patch: Partial<ChartObjectProps>) => host.edit((c) => {
    const p: ChartObjectProps = { ...c.props, ...patch };
    for (const k of Object.keys(p) as (keyof ChartObjectProps)[]) if (p[k] === undefined) delete p[k];
    if (Object.keys(p).length) c.props = p; else delete c.props;
  });
  const placement = () => props()?.placement ?? "twoCell";
  const r2 = (v: number) => Math.round(v * 100) / 100;
  const setH = (h: number) => { const s = host.size(); host.resize(props()?.lockAspect ? r2(s.w * (h / s.h)) : s.w, r2(h)); };
  const setW = (w: number) => { const s = host.size(); host.resize(r2(w), props()?.lockAspect ? r2(s.h * (w / s.w)) : s.h); };
  return [
    { title: "Größe", open: true, controls: () => [
      { type: "number", label: "Höhe", unit: " cm", min: 0.1, max: 5963, step: 0.1, decimals: 2, value: () => host.size().h, set: setH },
      { type: "number", label: "Breite", unit: " cm", min: 0.1, max: 5963, step: 0.1, decimals: 2, value: () => host.size().w, set: setW },
      { type: "number", label: "Drehung", unit: "°", value: () => 0, set: () => undefined, disabled: () => true },
      { type: "number", label: "Höhe skalieren", unit: " %", min: 1, max: 10000, value: () => Math.round((host.size().h / DEFAULT_CM.h) * 100), set: (v) => setH((DEFAULT_CM.h * v) / 100) },
      { type: "number", label: "Breite skalieren", unit: " %", min: 1, max: 10000, value: () => Math.round((host.size().w / DEFAULT_CM.w) * 100), set: (v) => setW((DEFAULT_CM.w * v) / 100) },
      { type: "check", label: "Seitenverhältnis sperren", value: () => !!props()?.lockAspect, set: (on) => setProps({ lockAspect: on || undefined }) },
    ] },
    { title: "Eigenschaften", controls: () => [
      { type: "radio", label: "Von Zellposition und -größe abhängig", checked: () => placement() === "twoCell", pick: () => setProps({ placement: undefined }) },
      { type: "radio", label: "Nur von Zellposition abhängig", checked: () => placement() === "oneCell", pick: () => setProps({ placement: "oneCell" }) },
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
