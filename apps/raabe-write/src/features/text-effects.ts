// Start > Texteffekte und Typografie and Formformat > WordArt-Formate (Schnellformatvorlagen,
// Textfüllung, Textkontur, Texteffekte) with the pane Texteffekte formatieren. Every preset,
// Schatten, Spiegelung and Leuchten entry is the XML Word itself writes for it
// (text-effects.gen.ts, recorded with probe_word_texteffects.py); the rest (Kontur colours and
// widths, Leuchten colours, the pane's settings) changes those elements as Word does.

import { closePopup, colorGrid, h, icon, type MenuEntry } from "@raabe/shell";
import { fxColor, fxColorXml, fxNumbers, fxParts, fxSet, outlineXml, textFxCss, type TextFxPart } from "@raabe/text-engine";
import type { AppApi } from "./api";
import type { Dock } from "./dock";
import { GLOWS, REFLECTIONS, SHADOWS, TEXTFX_PRESETS } from "./text-effects.gen";

const WIDTHS = [0.25, 0.5, 0.75, 1, 1.5, 2.25, 3, 4.5, 6];
const DASHES: [string, string][] = [["solid", "Durchgezogen"], ["sysDot", "Runder Punkt"], ["sysDash", "Quadratischer Punkt"], ["dash", "Strich"],
  ["dashDot", "Strich-Punkt"], ["lgDash", "Langer Strich"], ["lgDashDot", "Langer Strich-Punkt"], ["lgDashDotDot", "Langer Strich-Punkt-Punkt"]];
const NUMBER_FORMS: [string, string | null, string | null][] = [["Standard", null, null], ["Proportionale Lineatur", "lining", "proportional"],
  ["Tabularische Lineatur", "lining", "tabular"], ["Proportional Old-Formatvorlage", "oldStyle", "proportional"], ["Tabellarisch Old-Formatvorlage", "oldStyle", "tabular"]];
const LIGATURES: [string, string | null][] = [["Keine", null], ["Nur Standard", "standard"], ["Standard und kontextbezogen", "standardContextual"],
  ["Historisch und frei", "historicalDiscretional"], ["Alle Ligaturen", "all"]];

/** Word's WordArt menus for the Formformat tab (filled in by the feature). */
export const WORDART = {
  styles: (_close: () => void): HTMLElement => h("div"),
  fill: (): MenuEntry[] => [],
  outline: (): MenuEntry[] => [],
  effects: (): MenuEntry[] => [],
};

const setAttr = (xml: string, name: string, value: number | string) =>
  new RegExp(`w14:${name}="[^"]*"`).test(xml) ? xml.replace(new RegExp(`w14:${name}="[^"]*"`), `w14:${name}="${value}"`) : xml.replace(/^(<w14:[A-Za-z0-9]+)/, `$1 w14:${name}="${value}"`);
/** Replaces the colour element (and its modifiers) inside an effect element. */
const setColor = (xml: string, colorXml: string) => xml.replace(/<w14:(srgbClr|schemeClr|prstClr)\b[^>]*?(\/>|>[\s\S]*?<\/w14:\1>)/, colorXml);
/** Alpha (0 … 1 opaque) of the effect's colour. */
const alphaOf = (xml: string | undefined) => { const v = xml ? /<w14:alpha w14:val="(\d+)"/.exec(xml)?.[1] : undefined; return v === undefined ? 1 : Number(v) / 100000; };
/** The colour element inside `xml` (or `xml` itself) with that alpha (1: opaque, no w14:alpha). */
const withAlpha = (xml: string, alpha: number) => xml.replace(/<w14:(srgbClr|schemeClr|prstClr)\b[^>]*?(?:\/>|>[\s\S]*?<\/w14:\1>)/, (el) => {
  const bare = el.replace(/<w14:alpha w14:val="\d+"\/>/g, "");
  if (alpha >= 0.999) return bare.replace(/^<(w14:[A-Za-z]+)\b([^>]*)><\/\1>$/, "<$1$2/>");
  const a = `<w14:alpha w14:val="${Math.round(Math.max(0, alpha) * 100000)}"/>`;
  const self = /^<w14:([A-Za-z]+)\b([^>]*?)\/>$/.exec(bare);
  return self ? `<w14:${self[1]}${self[2]}>${a}</w14:${self[1]}>` : bare.replace(/(<\/w14:[A-Za-z]+>)$/, `${a}$1`);
});

export function textEffectsFeature(api: AppApi, docArea: HTMLElement, dock: Dock) {
  const ed = () => api.editor();
  const c = api.commands;
  const has = () => !!ed();
  const colors = () => ed()?.theme.colors;
  const change = (part: TextFxPart, xml: string | null) => { const e = ed(); if (e) { e.setTextFx((fx) => fxSet(fx, part, xml)); api.refresh(); renderPane(); } };
  const current = (part: TextFxPart) => fxParts(ed()?.textFxAt())[part];

  /** A tile showing an "A" with effects (Word's previews). */
  const tile = (fx: string, label: string, pick: () => void, color?: string) =>
    h("button", { class: "fx-tile", "data-tip": label, "aria-label": label, onclick: pick },
      h("span", { style: `color:${color ?? "#000"};${textFxCss(fx, colors())}` }, "A"));

  // ----- presets (Texteffekte und Typografie, WordArt-Formate) ---------------------------------------
  const applyPreset = (p: (typeof TEXTFX_PRESETS)[number]) => {
    const e = ed();
    if (!e) return;
    // A preset replaces the effects; the typography stays.
    e.setTextFx((fx) => {
      const keep = fxParts(fx);
      const next = fxParts(p.fx);
      for (const k of ["ligatures", "numForm", "numSpacing", "stylisticSets", "cntxtAlts"] as const) next[k] = keep[k];
      return Object.values(next).join("") || undefined;
    }, p.color ? { color: p.color, colorTheme: p.colorTheme } : undefined);
    api.refresh();
    api.focus();
  };
  const presetGrid = (close: () => void) => h("div", { class: "fx-presets" }, ...TEXTFX_PRESETS.map((p) => tile(p.fx, p.name, () => { close(); applyPreset(p); },
    p.colorTheme ? `#${themeHex(p.colorTheme) ?? p.color}` : p.color ? `#${p.color}` : undefined)));
  const themeHex = (ref: string) => {
    const name = /w:themeColor="([^"]*)"/.exec(ref)?.[1];
    const map: Record<string, string> = { text1: "dk1", background1: "lt1", text2: "dk2", background2: "lt2" };
    return name ? colors()?.[map[name] ?? name] : undefined;
  };

  // ----- Kontur ---------------------------------------------------------------------------------------
  const outlineParts = () => {
    const o = current("textOutline");
    const w = o ? Number(/w14:w="(\d+)"/.exec(o)?.[1] ?? "0") / 12700 : 0;
    const colorXml = o && !/<w14:noFill\/>/.test(o) ? /<w14:solidFill>([\s\S]*?)<\/w14:solidFill>/.exec(o)?.[1] ?? null : null;
    const dash = o ? /w14:prstDash w14:val="([^"]*)"/.exec(o)?.[1] ?? "solid" : "solid";
    return { w, colorXml, dash };
  };
  const outlineMenu = (): MenuEntry[] => [
    { block: () => h("div", { class: "fx-palette" }, ...colorGrid((hex, ref) => { const p = outlineParts(); change("textOutline", outlineXml(fxColorXml(hex, ref), p.w || 0.75, p.dash)); }, colors())) },
    { label: "Keine Kontur", run: () => change("textOutline", null) },
    { label: "Weitere Konturfarben...", icon: "colors", run: () => pickColor((hex) => { const p = outlineParts(); change("textOutline", outlineXml(fxColorXml(hex), p.w || 0.75, p.dash)); }) },
    { label: "Stärke", icon: "line-weight", submenu: () => WIDTHS.map((w): MenuEntry => ({
      label: `${w.toLocaleString("de-DE")} Pt.`, get checked() { return Math.abs(outlineParts().w - w) < 0.01; },
      run: () => { const p = outlineParts(); change("textOutline", outlineXml(p.colorXml ?? fxColorXml("000000"), w, p.dash)); },
    })) },
    { label: "Striche", icon: "dash-style", submenu: () => DASHES.map(([v, l]): MenuEntry => ({
      label: l, get checked() { return outlineParts().dash === v; },
      run: () => { const p = outlineParts(); change("textOutline", outlineXml(p.colorXml ?? fxColorXml("000000"), p.w || 0.75, v)); },
    })) },
  ];

  const pickColor = (done: (hex: string) => void) => {
    const input = h("input", { type: "color", style: "position:fixed;opacity:0;width:0;height:0" }) as HTMLInputElement;
    document.body.append(input);
    input.addEventListener("change", () => { done(input.value.slice(1).toUpperCase()); input.remove(); });
    input.addEventListener("blur", () => setTimeout(() => input.remove(), 500));
    input.click();
  };

  // ----- Schatten, Spiegelung, Leuchten ------------------------------------------------------------------
  const gallery = (sections: { heading: string; items: { name: string; xml: string | null }[] }[], part: TextFxPart, cols: number, footer: MenuEntry[]) => () => {
    const box = h("div", { class: "fx-gallery" });
    for (const s of sections) {
      box.append(h("div", { class: "menu-heading" }, s.heading), h("div", { class: "fx-grid", style: `grid-template-columns:repeat(${cols}, 54px)` },
        ...s.items.map((it) => tile(it.xml ?? "", it.name, () => { closeMenus(); change(part, it.xml); }))));
    }
    for (const f of footer) if (typeof f === "object" && "label" in f && "run" in f) box.append(h("button", { class: "fx-cmd", onclick: () => { closeMenus(); f.run(); } }, icon(f.icon), f.label));
    return box;
  };
  const closeMenus = () => closePopup();
  const by = (list: readonly { name: string; xml: string | null }[], prefix: string) => list.filter((x) => x.name.startsWith(prefix));
  const shadowFlyout = gallery([
    { heading: "Kein Schatten", items: by(SHADOWS, "Ohne") }, { heading: "Außen", items: by(SHADOWS, "Offset") },
    { heading: "Innen", items: by(SHADOWS, "Innen") }, { heading: "Perspektive", items: by(SHADOWS, "Perspektive") },
  ], "shadow", 3, [{ label: "Schattenoptionen...", icon: "shadow", run: () => openPane("effects") }]);
  const reflectionFlyout = gallery([
    { heading: "Keine Spiegelung", items: by(REFLECTIONS, "Ohne") }, { heading: "Spiegelungsvarianten", items: REFLECTIONS.filter((x) => x.name !== "Ohne") },
  ], "reflection", 3, [{ label: "Spiegelungsoptionen...", icon: "reflection", run: () => openPane("effects") }]);
  const glowFlyout = gallery([
    { heading: "Kein Leuchteffekt", items: by(GLOWS, "Ohne") }, { heading: "Leuchtvarianten", items: GLOWS.filter((x) => x.name !== "Ohne") },
  ], "glow", 6, [{ label: "Leuchteffektoptionen...", icon: "effects", run: () => openPane("effects") }]);
  const glowColors = (): MenuEntry[] => [{ block: () => h("div", { class: "fx-palette" }, ...colorGrid((hex, ref) => {
    const g = current("glow");
    const rad = g ? fxNumbers(g).rad : 101600;
    change("glow", `<w14:glow w14:rad="${rad}">${withAlpha(fxColorXml(hex, ref), 0.4)}</w14:glow>`);
  }, colors())) }, { label: "Weitere Farben...", run: () => pickColor((hex) => change("glow", `<w14:glow w14:rad="${fxNumbers(current("glow")).rad ?? 101600}">${withAlpha(fxColorXml(hex), 0.4)}</w14:glow>`)) }];

  // ----- Typografie -----------------------------------------------------------------------------------
  const numberMenu = (): MenuEntry[] => NUMBER_FORMS.map(([label, form, spacing]): MenuEntry => ({
    label, get checked() { return (current("numForm") ? /w14:val="([^"]*)"/.exec(current("numForm")!)?.[1] ?? null : null) === form && (current("numSpacing") ? /w14:val="([^"]*)"/.exec(current("numSpacing")!)?.[1] ?? null : null) === spacing; },
    run: () => { const e = ed(); if (!e) return; e.setTextFx((fx) => fxSet(fxSet(fx, "numForm", form ? `<w14:numForm w14:val="${form}"/>` : null), "numSpacing", spacing ? `<w14:numSpacing w14:val="${spacing}"/>` : null)); api.refresh(); },
  }));
  const ligatureMenu = (): MenuEntry[] => LIGATURES.map(([label, v]): MenuEntry => ({
    label, get checked() { return (current("ligatures") ? /w14:val="([^"]*)"/.exec(current("ligatures")!)?.[1] ?? null : null) === v; },
    run: () => change("ligatures", v ? `<w14:ligatures w14:val="${v}"/>` : null),
  }));
  const styleSetMenu = (): MenuEntry[] => [
    { label: "Standard", get checked() { return !current("stylisticSets"); }, run: () => change("stylisticSets", null) },
    "-",
    ...Array.from({ length: 20 }, (_, i): MenuEntry => ({
      label: `Stil-Set ${i + 1}`, get checked() { return new RegExp(`w14:id="${i + 1}"`).test(current("stylisticSets") ?? ""); },
      run: () => change("stylisticSets", `<w14:stylisticSets><w14:styleSet w14:id="${i + 1}"/></w14:stylisticSets>`),
    })),
  ];

  /** Start > Texteffekte und Typografie. */
  const fontEffectsMenu = (): MenuEntry[] => [
    { block: () => presetGrid(closeMenus) },
    "-",
    { label: "Kontur", icon: "text-outline", submenu: outlineMenu },
    { label: "Schatten", icon: "shadow", flyout: shadowFlyout },
    { label: "Spiegelung", icon: "reflection", flyout: reflectionFlyout },
    { label: "Leuchten", icon: "effects", submenu: () => [{ block: glowFlyout }, { label: "Weitere Leuchtfarben", submenu: glowColors }] },
    { label: "Zahlenformatvorlagen", icon: "number-forms", submenu: numberMenu },
    { label: "Ligaturen", icon: "ligatures", submenu: ligatureMenu },
    { label: "Stil-Sets", icon: "stylistic-sets", submenu: styleSetMenu },
  ];

  // ----- Formformat > WordArt-Formate -------------------------------------------------------------------
  const fillParts = () => {
    const f = current("textFill");
    return { none: !!f && /<w14:noFill\/>/.test(f), gradient: !!f && /<w14:gradFill>/.test(f), xml: f };
  };
  const gradientOf = (colorXml: string, light: boolean, ang: number) => {
    const m = /<w14:(srgbClr|schemeClr) w14:val="([^"]*)"/.exec(colorXml);
    const el = m?.[1] ?? "srgbClr", val = m?.[2] ?? "000000";
    const stop = (pos: number, mods: string) => `<w14:gs w14:pos="${pos}"><w14:${el} w14:val="${val}">${mods}</w14:${el}></w14:gs>`;
    const stops = light ? stop(0, "") + stop(100000, '<w14:lumMod w14:val="40000"/><w14:lumOff w14:val="60000"/>')
      : stop(0, '<w14:lumMod w14:val="50000"/>') + stop(100000, "");
    return `<w14:textFill><w14:gradFill><w14:gsLst>${stops}</w14:gsLst><w14:lin w14:ang="${ang * 60000}" w14:scaled="0"/></w14:gradFill></w14:textFill>`;
  };
  const baseColorXml = () => {
    const f = current("textFill");
    const inner = f ? /<w14:solidFill>([\s\S]*?)<\/w14:solidFill>/.exec(f)?.[1] : undefined;
    const e = ed();
    return inner ?? fxColorXml(e?.charState().color && e.charState().color !== "auto" ? e.charState().color! : "000000");
  };
  const fillMenu = (): MenuEntry[] => [
    { block: () => h("div", { class: "fx-palette" }, ...colorGrid((hex, ref) => change("textFill", `<w14:textFill><w14:solidFill>${fxColorXml(hex, ref)}</w14:solidFill></w14:textFill>`), colors())) },
    { label: "Keine Füllung", get checked() { return fillParts().none; }, run: () => change("textFill", "<w14:textFill><w14:noFill/></w14:textFill>") },
    { label: "Weitere Füllfarben...", icon: "colors", run: () => pickColor((hex) => change("textFill", `<w14:textFill><w14:solidFill>${fxColorXml(hex)}</w14:solidFill></w14:textFill>`)) },
    { label: "Farbverlauf", icon: "gradient", flyout: () => {
      const col = baseColorXml();
      const dirs: [number, string][] = [[90, "Linear nach unten"], [270, "Linear nach oben"], [0, "Linear nach rechts"], [180, "Linear nach links"]];
      const box = h("div", { class: "fx-gallery" });
      box.append(h("div", { class: "menu-heading" }, "Keine Farbverläufe"), h("div", { class: "fx-grid", style: "grid-template-columns:repeat(4, 54px)" },
        tile("", "Keine Farbverläufe", () => { closeMenus(); change("textFill", null); })));
      for (const [heading, light] of [["Helle Varianten", true], ["Dunkle Varianten", false]] as const) {
        box.append(h("div", { class: "menu-heading" }, heading), h("div", { class: "fx-grid", style: "grid-template-columns:repeat(4, 54px)" },
          ...dirs.map(([ang, label]) => { const fx = gradientOf(col, light, ang); return tile(fx, label, () => { closeMenus(); change("textFill", fx); }); })));
      }
      box.append(h("button", { class: "fx-cmd", onclick: () => { closeMenus(); openPane("fill"); } }, icon("gradient"), "Weitere Farbverläufe..."));
      return box;
    } },
  ];
  const shapeEffectsMenu = (): MenuEntry[] => [
    { label: "Schatten", icon: "shadow", flyout: shadowFlyout },
    { label: "Spiegelung", icon: "reflection", flyout: reflectionFlyout },
    { label: "Leuchten", icon: "effects", submenu: () => [{ block: glowFlyout }, { label: "Weitere Leuchtfarben", submenu: glowColors }] },
    { label: "Abschrägung", icon: "bevel", submenu: [], disabled: true },
    { label: "3D-Drehung", icon: "rotate-3d", submenu: [], disabled: true },
    { label: "Transformieren", icon: "transform", submenu: [], disabled: true },
  ];
  WORDART.styles = (close) => h("div", { class: "rpanel fx-wordart" }, presetGrid(close));
  WORDART.fill = fillMenu;
  WORDART.outline = outlineMenu;
  WORDART.effects = shapeEffectsMenu;

  // ----- Texteffekte formatieren (pane) -------------------------------------------------------------------
  let paneOpen = false;
  let tab: "fill" | "effects" = "fill";
  const pane = h("div", { class: "fxpane", hidden: true, role: "complementary", "aria-label": "Texteffekte formatieren" });
  docArea.append(pane);
  dock.add(pane, "right", 300);
  const openPane = (t: "fill" | "effects") => { tab = t; paneOpen = true; pane.hidden = false; renderPane(); api.refresh(); };
  const closePane = () => { paneOpen = false; pane.hidden = true; api.refresh(); api.focus(); };

  const slider = (label: string, value: number, min: number, max: number, unit: string, set: (v: number) => void, step = 1) => {
    const range = h("input", { type: "range", min: String(min), max: String(max), step: String(step), value: String(value) }) as HTMLInputElement;
    const box = h("input", { type: "number", min: String(min), max: String(max), step: String(step), value: String(Math.round(value * 10) / 10), style: "width:60px" }) as HTMLInputElement;
    range.onchange = () => { box.value = range.value; set(Number(range.value)); };
    box.onchange = () => { range.value = box.value; set(Number(box.value)); };
    return h("label", { class: "fx-field" }, h("span", null, label), range, h("span", { class: "fx-num" }, box, unit));
  };
  const colorButton = (xml: string | undefined, set: (hex: string, ref?: string) => void) => {
    const b = h("button", { class: "fx-colorbtn", style: `background:${xml ? fxColor(xml, colors()) : "#000"}`, "aria-label": "Farbe" }) as HTMLButtonElement;
    b.onclick = () => {
      const r = b.getBoundingClientRect();
      const box = h("div", { class: "menu palette fx-floatpal", style: `left:${r.left - 120}px;top:${r.bottom + 2}px` }, ...colorGrid((hex, ref) => { box.remove(); set(hex, ref); }, colors()));
      document.body.append(box);
      setTimeout(() => addEventListener("mousedown", function off(ev) { if (!box.contains(ev.target as Node)) { box.remove(); removeEventListener("mousedown", off, true); } }, true));
    };
    return h("label", { class: "fx-field" }, h("span", null, "Farbe"), b);
  };

  const section = (title: string, ...body: HTMLElement[]) => h("details", { class: "fx-sec", open: true }, h("summary", null, title), ...body);

  const renderPane = () => {
    if (!paneOpen) return;
    const e = ed();
    const parts = fxParts(e?.textFxAt());
    const head = h("div", { class: "fx-head" }, h("span", null, "Texteffekte formatieren"),
      h("button", { "aria-label": "Bereich schließen", "data-tip": "Bereich schließen", onclick: closePane }, icon("close")));
    const tabs = h("div", { class: "fx-tabs" },
      h("button", { class: tab === "fill" ? "active" : "", "data-tip": "Textfüllung und Kontur", onclick: () => { tab = "fill"; renderPane(); } }, icon("text-fill")),
      h("button", { class: tab === "effects" ? "active" : "", "data-tip": "Texteffekte", onclick: () => { tab = "effects"; renderPane(); } }, icon("text-effects")));
    const body = h("div", { class: "fx-body" });
    if (tab === "fill") {
      const f = parts.textFill;
      const kind = !f ? "auto" : /<w14:noFill\/>/.test(f) ? "none" : /<w14:gradFill>/.test(f) ? "gradient" : "solid";
      const radio = (k: string, label: string, run: () => void) => h("label", { class: "check" }, h("input", { type: "radio", name: "fxfill", checked: kind === k, onchange: run }), label);
      body.append(section("Textfüllung",
        radio("none", "Keine Füllung", () => change("textFill", "<w14:textFill><w14:noFill/></w14:textFill>")),
        radio("solid", "Einfarbige Füllung", () => change("textFill", `<w14:textFill><w14:solidFill>${baseColorXml()}</w14:solidFill></w14:textFill>`)),
        radio("gradient", "Farbverlauf", () => change("textFill", gradientOf(baseColorXml(), true, 90))),
        ...(kind === "solid" ? [colorButton(f, (hex, ref) => change("textFill", `<w14:textFill><w14:solidFill>${withAlpha(fxColorXml(hex, ref), alphaOf(f))}</w14:solidFill></w14:textFill>`)),
          slider("Transparenz", Math.round((1 - alphaOf(f)) * 100), 0, 100, "%", (v) => change("textFill", withAlpha(f!, 1 - v / 100)))] : [])));
      const o = parts.textOutline;
      const okind = !o || /<w14:noFill\/>/.test(o) ? "none" : "solid";
      const op = outlineParts();
      const oradio = (k: string, label: string, run: () => void) => h("label", { class: "check" }, h("input", { type: "radio", name: "fxline", checked: okind === k, onchange: run }), label);
      const dashSel = h("select", {}, ...DASHES.map(([v, l]) => h("option", { value: v }, l))) as HTMLSelectElement;
      dashSel.value = op.dash;
      dashSel.onchange = () => change("textOutline", outlineXml(op.colorXml ?? fxColorXml("000000"), op.w || 0.75, dashSel.value));
      body.append(section("Textkontur",
        oradio("none", "Keine Linie", () => change("textOutline", null)),
        oradio("solid", "Einfarbige Linie", () => change("textOutline", outlineXml(fxColorXml("000000"), 0.75))),
        ...(okind === "solid" ? [colorButton(o, (hex, ref) => change("textOutline", outlineXml(fxColorXml(hex, ref), op.w || 0.75, op.dash))),
          slider("Breite", op.w, 0, 20, "Pt.", (v) => change("textOutline", outlineXml(op.colorXml ?? fxColorXml("000000"), v, op.dash)), 0.25),
          h("label", { class: "fx-field" }, h("span", null, "Strichtyp"), dashSel)] : [])));
    } else {
      const presetSel = (list: readonly { name: string; xml: string | null }[], part: TextFxPart, cur: string | undefined) => {
        const s = h("select", {}, ...list.map((x, i) => h("option", { value: String(i) }, x.name))) as HTMLSelectElement;
        const hit = list.findIndex((x) => (x.xml ?? "") === (cur ?? ""));
        s.value = String(hit >= 0 ? hit : 0);
        s.onchange = () => change(part, list[Number(s.value)].xml);
        return h("label", { class: "fx-field" }, h("span", null, "Voreinstellungen"), s);
      };
      const sh = parts.shadow, n = fxNumbers(sh);
      const shadowSet = (k: string, v: number) => sh && change("shadow", setAttr(sh, k, Math.round(v)));
      body.append(section("Schatten", presetSel(SHADOWS, "shadow", sh), ...(sh ? [
        colorButton(sh, (hex, ref) => change("shadow", setColor(sh, withAlpha(fxColorXml(hex, ref), alphaOf(sh))))),
        slider("Transparenz", Math.round((1 - alphaOf(sh)) * 100), 0, 100, "%", (v) => change("shadow", withAlpha(sh, 1 - v / 100))),
        slider("Größe", (n.sx ?? 100000) / 1000, 1, 200, "%", (v) => { if (sh) change("shadow", setAttr(setAttr(sh, "sx", Math.round(v * 1000)), "sy", Math.round(v * 1000))); }),
        slider("Weichzeichnen", (n.blurRad ?? 0) / 12700, 0, 100, "Pt.", (v) => shadowSet("blurRad", v * 12700), 0.5),
        slider("Winkel", (n.dir ?? 0) / 60000, 0, 359, "°", (v) => shadowSet("dir", v * 60000)),
        slider("Abstand", (n.dist ?? 0) / 12700, 0, 200, "Pt.", (v) => shadowSet("dist", v * 12700), 0.5)] : [])));
      const rf = parts.reflection, rn = fxNumbers(rf);
      const reflSet = (k: string, v: number) => rf && change("reflection", setAttr(rf, k, Math.round(v)));
      body.append(section("Spiegelung", presetSel(REFLECTIONS, "reflection", rf), ...(rf ? [
        slider("Transparenz", 100 - (rn.stA ?? 50000) / 1000, 0, 100, "%", (v) => reflSet("stA", (100 - v) * 1000)),
        slider("Größe", (rn.endPos ?? 35000) / 1000, 1, 100, "%", (v) => reflSet("endPos", v * 1000)),
        slider("Weichzeichnen", (rn.blurRad ?? 0) / 12700, 0, 100, "Pt.", (v) => reflSet("blurRad", v * 12700), 0.5),
        slider("Abstand", (rn.dist ?? 0) / 12700, 0, 100, "Pt.", (v) => reflSet("dist", v * 12700), 0.5)] : [])));
      const gl = parts.glow, gn = fxNumbers(gl);
      body.append(section("Leuchten", presetSel(GLOWS, "glow", gl), ...(gl ? [
        colorButton(gl, (hex, ref) => change("glow", setColor(gl, withAlpha(fxColorXml(hex, ref), alphaOf(gl))))),
        slider("Größe", (gn.rad ?? 0) / 12700, 0, 150, "Pt.", (v) => change("glow", setAttr(gl, "rad", Math.round(v * 12700))), 0.5),
        slider("Transparenz", Math.round((1 - alphaOf(gl)) * 100), 0, 100, "%", (v) => change("glow", withAlpha(gl, 1 - v / 100)))] : [])));
      body.append(section("3D-Format", h("div", { class: "muted" }, "RaabeWrite zeigt Abschrägungen und 3D-Effekte nicht an; Word sieht sie weiterhin.")));
    }
    pane.replaceChildren(head, tabs, body);
  };

  c.implement("font.effects", { run: () => {}, enabled: has });
  for (const id of ["wordart.styles", "text.fill", "text.outline", "text.effects"]) c.implement(id, { run: () => {}, enabled: () => !!ed()?.selectedShape() });
  c.implement("text.effectsPane", { run: () => (paneOpen ? closePane() : openPane("fill")), enabled: has, checked: () => paneOpen });
  return { get fontEffectsMenu() { return fontEffectsMenu(); }, refresh: renderPane };
}

