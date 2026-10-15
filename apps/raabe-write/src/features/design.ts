// Entwurf > Dokumentformatierung: Designs, Formatvorlagensatz, Farben, Schriftarten,
// Absatzabstand, Effekte and Als Standard festlegen, with Word's entries (what each applies was
// read from Word: tools/feel-lab/probe_word_design.py, design-data.gen.ts). The thumbnails are
// RaabeWrite's own, drawn from each entry's colours and fonts.

import { h, icon, showDialog, type MenuEntry } from "@raabe/shell";
import { readThemeFile } from "@raabe/ooxml";
import { OFFICE_2013_COLORS, StyleSheet, defaultStyleSheet, type CharProps, type Editor, type ParaProps, type Theme } from "@raabe/text-engine";
import type { AppApi } from "./api";
import { EFFECTS, PARAGRAPH_SPACING, STYLE_SETS, THEMES, THEME_COLORS, THEME_FONTS, type StyleSetData } from "./design-data.gen";
import { openModal } from "./modal";

const COLORS_KEY = "raabe.theme.colors", FONTS_KEY = "raabe.theme.fonts", DEFAULTS_KEY = "raabe.write.defaults";
/** Word's order of the colours a scheme shows in the Farben list. */
const ROW = ["dk2", "lt2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6"];
const COLOR_FIELDS: [string, string][] = [["dk1", "Text/Hintergrund - Dunkel 1"], ["lt1", "Text/Hintergrund - Hell 1"], ["dk2", "Text/Hintergrund - Dunkel 2"],
  ["lt2", "Text/Hintergrund - Hell 2"], ["accent1", "Akzent 1"], ["accent2", "Akzent 2"], ["accent3", "Akzent 3"], ["accent4", "Akzent 4"],
  ["accent5", "Akzent 5"], ["accent6", "Akzent 6"], ["hlink", "Link"], ["folHlink", "Besuchter Link"]];

/** RaabeWrite's new documents use Office 2013–2022 (D-011). */
const OFFICE_2013: Theme = { majorFont: "Calibri Light", minorFont: "Calibri", colors: OFFICE_2013_COLORS, name: "Office 2013 – 2022", colorName: "Office 2013 – 2022", fontName: "Office 2013 – 2022" };

type Saved<T> = (T & { name: string })[];
const load = <T>(key: string): Saved<T> => { try { return JSON.parse(localStorage.getItem(key) ?? "[]") as Saved<T>; } catch { return []; } };
const store = (key: string, v: unknown) => { try { localStorage.setItem(key, JSON.stringify(v)); } catch { /* not kept */ } };

/** What "Als Standard festlegen" keeps for new documents. */
interface NewDocDefaults {
  theme: Theme;
  defaultsR: CharProps;
  defaultsP: ParaProps;
  styles: Record<string, { pPr: ParaProps; rPr: CharProps }>;
}

/** The style sheet of a new blank document after Als Standard festlegen, or null (Word's defaults). */
export function newDocumentStyles(): StyleSheet | null {
  let d: NewDocDefaults | null = null;
  try { d = JSON.parse(localStorage.getItem(DEFAULTS_KEY) ?? "null") as NewDocDefaults | null; } catch { return null; }
  if (!d) return null;
  let sheet = defaultStyleSheet().withTheme(d.theme).withDefaults(d.defaultsR, d.defaultsP);
  for (const [name, s] of Object.entries(d.styles)) {
    const id = sheet.idForName(name);
    const cur = id ? sheet.get(id) : undefined;
    if (cur) sheet = sheet.withStyle({ ...cur, pPr: s.pPr, rPr: s.rPr, data: undefined });
  }
  return sheet;
}

export function designFeature(api: AppApi) {
  const ed = () => api.editor();
  const c = api.commands;
  const has = () => !!ed();

  const apply = (fn: (e: Editor) => void) => { const e = ed(); if (e) { fn(e); api.refresh(); api.focus(); } };
  const themeOf = (e: Editor) => e.theme;
  const sameColors = (a: Readonly<Record<string, string>> | undefined, b: Readonly<Record<string, string>> | undefined) =>
    !!a && !!b && Object.keys(OFFICE_2013_COLORS).every((k) => (a[k] ?? "").toUpperCase() === (b[k] ?? "").toUpperCase());

  // ----- Designs ---------------------------------------------------------------------------------
  // Word's gallery names the old Office theme "Office 2013 – 2022-Design" (its colours and fonts without "-Design").
  const tileName = (name: string) => (name === "Office 2013 – 2022" ? `${name}-Design` : name);
  const themeTile = (t: { name: string; colors: Readonly<Record<string, string>>; major: string; minor: string }, active: boolean, pick: () => void) =>
    h("button", { class: `dz-theme${active ? " active" : ""}`, "data-tip": tileName(t.name), "aria-label": tileName(t.name), onclick: pick },
      h("span", { class: "dz-page" },
        h("span", { class: "dz-aa", style: `font-family:"${t.major}";color:#${t.colors.dk2 ?? "000"}` }, "Aa"),
        h("span", { class: "dz-bars" }, ...["accent1", "accent2", "accent3", "accent4", "accent5", "accent6"].map((k) => h("i", { style: `background:#${t.colors[k]}` }))),
        h("span", { class: "dz-lines" }, h("i"), h("i"), h("i"))),
      h("span", { class: "dz-name" }, tileName(t.name)));

  const themesPanel = (close: () => void) => {
    const e = ed();
    const cur = e ? themeOf(e) : OFFICE_2013;
    const grid = h("div", { class: "dz-grid" }, ...THEMES.map((t) => themeTile(t, cur.name === t.name || (cur.name === undefined && t.name.startsWith("Office 2013") && sameColors(cur.colors, t.colors)),
      () => { close(); apply((x) => x.setTheme({ majorFont: t.major, minorFont: t.minor, colors: t.colors, name: t.name, colorName: t.name, fontName: t.name })); })));
    const item = (label: string, ic: string | null, run: () => void, disabled = false, tip?: string) =>
      h("button", { class: "dz-cmd", disabled, onclick: () => { close(); run(); }, ...(tip ? { "data-tip": tip } : {}) }, ic ? icon(ic) : h("span", { class: "dz-noicon" }), label);
    return h("div", { class: "rpanel designpanel" }, h("div", { class: "title" }, "Office"), grid, h("hr"),
      item("Auf das Design aus der Vorlage zurücksetzen", null, () => apply((x) => x.setTheme(OFFICE_2013))),
      item("Nach Designs suchen...", "themes", () => void browseTheme()),
      item("Aktuelles Design speichern...", "save", () => {}, true, "Aktuelles Design speichern\nRaabeWrite speichert keine Designdateien."));
  };

  const browseTheme = async () => {
    const pf = api.platform;
    const path = await pf.openAnyDialog?.("Design oder Designdokument auswählen", [{ name: "Office-Designs und Designdokumente", extensions: ["thmx", "xml", "docx", "dotx"] }]).catch(() => null);
    if (!path) return;
    const bytes = await pf.readFile(path).catch(() => null);
    const t = bytes ? readThemeFile(bytes) : null;
    const e = ed();
    if (!t || !e) { void showDialog({ title: "RaabeWrite", body: "Die Datei enthält kein Design.", buttons: [{ label: "OK", value: undefined, primary: true }], cancel: undefined }); return; }
    const cur = themeOf(e);
    e.setTheme({ ...cur, ...(t.colors ? { colors: t.colors, colorName: t.colorName } : {}), ...(t.major ? { majorFont: t.major } : {}), ...(t.minor ? { minorFont: t.minor } : {}),
      ...(t.fontName ? { fontName: t.fontName } : {}), ...(t.name ? { name: t.name } : {}) });
    api.refresh();
  };

  // ----- Formatvorlagensatz ------------------------------------------------------------------------
  /** A style set applied to a copy of the document's styles, for its thumbnail. */
  const previewSheet = (e: Editor, set: StyleSetData) => {
    let sheet = e.doc.styles.withDefaults({ ...set.defaultsR }, set.defaultsP);
    for (const [name, s] of Object.entries(set.styles)) {
      const id = sheet.idForName(name);
      const cur = id ? sheet.get(id) : defaultStyleSheet().get(defaultStyleSheet().idForName(name) ?? "");
      if (cur) sheet = sheet.withStyle({ ...cur, pPr: s.pPr, rPr: s.rPr, data: undefined });
    }
    return sheet;
  };
  const styleTile = (name: string, sheet: StyleSheet, pick: () => void) => {
    const look = (n: string) => {
      const id = sheet.idForName(n) ?? sheet.defaultParagraphStyle;
      const ch = sheet.resolveChar(id, {});
      const p = sheet.resolvePara({ style: id } as ParaProps);
      const b = p.borders?.bottom;
      return `font-family:"${ch.font}";font-size:${Math.max(6, Math.min(15, ch.size * 0.42))}px;color:${ch.color === "auto" ? "#000" : `#${ch.color}`};` +
        `font-weight:${ch.bold ? 700 : 400};font-style:${ch.italic ? "italic" : "normal"};${ch.caps ? "text-transform:uppercase;" : ""}${ch.smallCaps ? "font-variant:small-caps;" : ""}` +
        `text-align:${p.align === "both" ? "left" : p.align};${b ? `border-bottom:1px solid #${b.color === "auto" ? "000" : b.color};` : ""}${p.shading?.fill && p.shading.fill !== "auto" ? `background:#${p.shading.fill};` : ""}`;
    };
    return h("button", { class: "dz-set", "data-tip": name, "aria-label": name, onclick: pick },
      h("span", { class: "dz-setpage" },
        h("span", { style: look("title") }, "Titel"),
        h("span", { style: look("heading 1") }, "Überschrift 1"),
        h("span", { class: "dz-body", style: look("normal") }, "Auf der Registerkarte „Einfügen“ enthalten die Kataloge Elemente, die mit dem generellen Layout des Dokuments koordiniert werden sollten.")),
      h("span", { class: "dz-name" }, name));
  };
  const styleSetPanel = (close: () => void) => {
    const e = ed();
    if (!e) return h("div", { class: "rpanel" });
    const word2013 = STYLE_SETS.find((s) => s.name === "Word 2013");
    const grid = h("div", { class: "dz-sets" },
      ...STYLE_SETS.map((s) => styleTile(s.name, previewSheet(e, s.set), () => { close(); apply((x) => x.applyStyleSet(s.set)); })));
    return h("div", { class: "rpanel designpanel wide" },
      h("div", { class: "title" }, "Dieses Dokument"), h("div", { class: "dz-sets" }, styleTile("Dokument", e.doc.styles, () => close())),
      h("div", { class: "title" }, "Integriert"), grid, h("hr"),
      h("button", { class: "dz-cmd", onclick: () => { close(); if (word2013) apply((x) => x.applyStyleSet(word2013.set)); } }, h("span", { class: "dz-noicon" }), "Auf den Standardformatvorlagensatz zurücksetzen"),
      h("button", { class: "dz-cmd", disabled: true, "data-tip": "Als neuen Formatvorlagensatz speichern\nRaabeWrite speichert keine Formatvorlagensätze." }, h("span", { class: "dz-noicon" }), "Als neuen Formatvorlagensatz speichern..."));
  };

  // ----- Farben, Schriftarten -----------------------------------------------------------------------
  const colorsPanel = (close: () => void) => {
    const e = ed();
    const cur = e ? themeOf(e) : OFFICE_2013;
    const row = (name: string, colors: Readonly<Record<string, string>>) =>
      h("button", { class: `dz-crow${sameColors(cur.colors, colors) ? " active" : ""}`, onclick: () => { close(); apply((x) => x.setTheme({ ...themeOf(x), colors, colorName: name })); } },
        h("span", { class: "dz-swatches" }, ...ROW.map((k) => h("i", { style: `background:#${colors[k]}` }))), h("span", null, name));
    const custom = load<{ colors: Record<string, string> }>(COLORS_KEY);
    return h("div", { class: "rpanel designpanel list" },
      ...(custom.length ? [h("div", { class: "title" }, "Benutzerdefiniert"), ...custom.map((x) => row(x.name, x.colors))] : []),
      h("div", { class: "title" }, "Office"), ...THEME_COLORS.map((x) => row(x.name, x.colors)), h("hr"),
      h("button", { class: "dz-cmd", onclick: () => { close(); colorsDialog(); } }, h("span", { class: "dz-noicon" }), "Farben anpassen..."));
  };

  const colorsDialog = () => {
    const e = ed();
    if (!e) return;
    const start = { ...OFFICE_2013_COLORS, ...(themeOf(e).colors ?? {}) } as Record<string, string>;
    const vals = { ...start };
    const custom = load<{ colors: Record<string, string> }>(COLORS_KEY);
    const name = h("input", { type: "text", value: `Benutzerdefiniert ${custom.length + 1}`, style: "width:260px" }) as HTMLInputElement;
    const sample = h("div", { class: "dz-sample" });
    const paint = () => {
      sample.replaceChildren(h("div", { class: "dz-sample-text", style: `background:#${vals.lt1};color:#${vals.dk1}` }, h("b", null, "Text"), h("span", { style: `color:#${vals.hlink};text-decoration:underline` }, "Link")),
        h("div", { class: "dz-sample-bars" }, ...["accent1", "accent2", "accent3", "accent4", "accent5", "accent6"].map((k) => h("i", { style: `background:#${vals[k]}` }))));
    };
    const inputs = COLOR_FIELDS.map(([k, label]) => {
      const inp = h("input", { type: "color", value: `#${vals[k]}` }) as HTMLInputElement;
      inp.oninput = () => { vals[k] = inp.value.slice(1).toUpperCase(); paint(); };
      return { k, inp, el: h("label", { class: "dz-colorfield" }, h("span", null, label), inp) };
    });
    paint();
    openModal("Neue Designfarben erstellen", h("div", { class: "dz-colordlg" },
      h("div", { class: "dz-colorgrid" }, h("div", null, h("b", null, "Designfarben"), ...inputs.map((x) => x.el)), h("div", null, h("b", null, "Beispiel"), sample)),
      h("label", { class: "field" }, "Name:", name)), [
      { label: "Zurücksetzen", left: true, run: () => { Object.assign(vals, start); inputs.forEach((x) => { x.inp.value = `#${vals[x.k]}`; }); paint(); } },
      { label: "Speichern", primary: true, run: () => {
        const n = name.value.trim() || `Benutzerdefiniert ${custom.length + 1}`;
        store(COLORS_KEY, [...custom.filter((x) => x.name !== n), { name: n, colors: { ...vals } }]);
        apply((x) => x.setTheme({ ...themeOf(x), colors: { ...vals }, colorName: n }));
        return true;
      } },
      { label: "Abbrechen", run: () => true },
    ], { width: 560 });
  };

  const fontsPanel = (close: () => void) => {
    const e = ed();
    const cur = e ? themeOf(e) : OFFICE_2013;
    const row = (name: string, major: string, minor: string) =>
      h("button", { class: `dz-frow${cur.majorFont === major && cur.minorFont === minor ? " active" : ""}`, "aria-label": `${name}, Überschriftenschriftart: ${major}, Textschriftart: ${minor}`, onclick: () => { close(); apply((x) => x.setTheme({ ...themeOf(x), majorFont: major, minorFont: minor, fontName: name })); } },
        h("span", { class: "dz-faa" }, h("span", { style: `font-family:"${major}"` }, "A"), h("span", { style: `font-family:"${minor}"` }, "a")),
        h("span", { class: "dz-ftext" }, h("b", null, name), h("span", { style: `font-family:"${major}";font-size:15px` }, major), h("span", { style: `font-family:"${minor}"` }, minor)));
    const custom = load<{ major: string; minor: string }>(FONTS_KEY);
    return h("div", { class: "rpanel designpanel list" },
      ...(custom.length ? [h("div", { class: "title" }, "Benutzerdefiniert"), ...custom.map((x) => row(x.name, x.major, x.minor))] : []),
      h("div", { class: "title" }, "Office"), ...THEME_FONTS.map((x) => row(x.name, x.major, x.minor)), h("hr"),
      h("button", { class: "dz-cmd", onclick: () => { close(); fontsDialog(); } }, h("span", { class: "dz-noicon" }), "Schriftarten anpassen..."));
  };

  const fontsDialog = () => {
    const e = ed();
    if (!e) return;
    const t = themeOf(e);
    const fonts = api.fontList();
    const sel = (v: string) => { const s = h("select", { style: "width:220px" }, ...fonts.map((f) => h("option", { value: f }, f))) as HTMLSelectElement; if (!fonts.includes(v)) s.append(h("option", { value: v }, v)); s.value = v; return s; };
    const major = sel(t.majorFont), minor = sel(t.minorFont);
    const custom = load<{ major: string; minor: string }>(FONTS_KEY);
    const name = h("input", { type: "text", value: `Benutzerdefiniert ${custom.length + 1}`, style: "width:260px" }) as HTMLInputElement;
    const sample = h("div", { class: "dz-fsample" });
    const paint = () => sample.replaceChildren(h("div", { style: `font-family:"${major.value}";font-size:20px` }, "Überschrift"),
      h("div", { style: `font-family:"${minor.value}"` }, "Textkörpertext Textkörpertext Textkörpertext."));
    major.onchange = paint;
    minor.onchange = paint;
    paint();
    openModal("Neue Designschriftarten erstellen", h("div", { class: "dz-fontdlg" },
      h("div", { class: "grid2" }, "Überschriftenschriftart:", major, "Textkörperschriftart:", minor), h("b", null, "Beispiel"), sample,
      h("label", { class: "field" }, "Name:", name)), [
      { label: "Speichern", primary: true, run: () => {
        const n = name.value.trim() || `Benutzerdefiniert ${custom.length + 1}`;
        store(FONTS_KEY, [...custom.filter((x) => x.name !== n), { name: n, major: major.value, minor: minor.value }]);
        apply((x) => x.setTheme({ ...themeOf(x), majorFont: major.value, minorFont: minor.value, fontName: n }));
        return true;
      } },
      { label: "Abbrechen", run: () => true },
    ], { width: 480 });
  };

  // ----- Absatzabstand ----------------------------------------------------------------------------
  const SPACING_ICONS = ["spacing-none", "spacing-compact", "spacing-tight", "spacing-open", "spacing-relaxed", "spacing-double"];
  const spacingMenu = (): MenuEntry[] => {
    const e = ed();
    const normal = e ? e.styleLook(e.doc.styles.defaultParagraphStyle).para : undefined;
    const isPreset = (x: (typeof PARAGRAPH_SPACING)[number]) => !!normal && Math.abs(normal.after - x.after) < 0.1 && Math.abs(normal.before - x.before) < 0.1 &&
      normal.lineRule === x.lineRule && Math.abs(normal.line - x.line) < 0.02;
    const presets = PARAGRAPH_SPACING.filter((x) => x.name !== "Standard");
    // "Standard": the spacing of RaabeWrite's default set (Word 2013, D-011).
    const std = STYLE_SETS.find((s) => s.name === "Word 2013")?.set.defaultsP;
    return [
      { heading: "Formatvorlagensatz" },
      { label: "Standard", icon: "paragraph-spacing", run: () => apply((x) => x.setParagraphSpacing({ before: std?.before ?? 0, after: std?.after ?? 8, line: std?.line ?? 1.08, lineRule: std?.lineRule ?? "auto" })) },
      { heading: "Standard" },
      ...presets.map((x, i): MenuEntry => ({ label: x.name, icon: SPACING_ICONS[i], get checked() { return isPreset(x); }, run: () => apply((y) => y.setParagraphSpacing(x)) })),
      "-",
      { label: "Benutzerdefinierter Absatzabstand...", run: () => customSpacing() },
    ];
  };

  const customSpacing = () => {
    const e = ed();
    if (!e) return;
    const p = e.styleLook(e.doc.styles.defaultParagraphStyle).para;
    const num = (v: number) => h("input", { type: "number", min: "0", max: "1584", step: "1", value: String(v), style: "width:80px" }) as HTMLInputElement;
    const before = num(p.before), after = num(p.after);
    const rule = h("select", {}, ...([["1", "Einfach"], ["1.5", "1,5 Zeilen"], ["2", "Doppelt"], ["atLeast", "Mindestens"], ["exact", "Genau"], ["multiple", "Mehrfach"]] as const)
      .map(([v, l]) => h("option", { value: v }, l))) as HTMLSelectElement;
    rule.value = p.lineRule === "exact" ? "exact" : p.lineRule === "atLeast" ? "atLeast" : [1, 1.5, 2].some((v) => Math.abs(p.line - v) < 0.01) ? String(+p.line.toFixed(1)).replace(".0", "") : "multiple";
    const val = h("input", { type: "text", value: p.line.toLocaleString("de-DE"), style: "width:80px" }) as HTMLInputElement;
    openModal("Formatvorlagen verwalten", h("div", { class: "dz-spacing" },
      h("b", null, "Standard festlegen"),
      h("fieldset", null, h("legend", null, "Absatzabstand"), h("div", { class: "grid2" }, "Vor:", before, "Nach:", after, "Zeilenabstand:", rule, "Von:", val))), [
      { label: "OK", primary: true, run: () => {
        const v = Number(val.value.replace(",", ".")) || 1;
        const lr = rule.value === "exact" ? "exact" : rule.value === "atLeast" ? "atLeast" : "auto";
        const line = rule.value === "1" ? 1 : rule.value === "1.5" ? 1.5 : rule.value === "2" ? 2 : v;
        apply((x) => x.setParagraphSpacing({ before: Number(before.value) || 0, after: Number(after.value) || 0, line, lineRule: lr }));
        return true;
      } },
      { label: "Abbrechen", run: () => true },
    ], { width: 420 });
  };

  // ----- Effekte --------------------------------------------------------------------------------------
  const effectsPanel = (close: () => void) => h("div", { class: "rpanel designpanel" }, h("div", { class: "title" }, "Office"),
    h("div", { class: "dz-grid effects" }, ...EFFECTS.map((name) => h("button", {
      class: `dz-effect${name === "Office" ? " active" : ""}`, disabled: name !== "Office", "aria-label": name,
      "data-tip": name === "Office" ? name : `${name}\nRaabeWrite enthält nur die Effekte von Office.`, onclick: () => close(),
    }, h("span", { class: "dz-effpic" }, icon("effects")), h("span", { class: "dz-name" }, name)))));

  // ----- Als Standard festlegen ---------------------------------------------------------------------
  const setDefault = async () => {
    const e = ed();
    if (!e) return;
    const ok = await showDialog({
      title: "RaabeWrite",
      body: "Möchten Sie die aktuellen Formatvorlageneinstellungen, einschließlich Farben, Schriftarten und Absatzabstand, als Standard festlegen? Diese Einstellungen gelten für alle neuen Dokumente.",
      buttons: [{ label: "Ja", value: true, primary: true }, { label: "Nein", value: false }], cancel: false,
    });
    if (!ok) return;
    const sheet = e.doc.styles;
    const styles: NewDocDefaults["styles"] = {};
    for (const n of Object.keys(STYLE_SETS[0]?.set.styles ?? {})) {
      const id = sheet.idForName(n);
      const s = id ? sheet.get(id) : undefined;
      if (s) styles[n] = { pPr: s.pPr, rPr: s.rPr };
    }
    store(DEFAULTS_KEY, { theme: sheet.theme, defaultsR: sheet.defaultsR, defaultsP: sheet.defaultsP, styles } satisfies NewDocDefaults);
    api.focus();
  };

  for (const id of ["design.themes", "design.styleSet", "design.colors", "design.fonts", "design.paragraphSpacing", "design.effects"]) c.implement(id, { run: () => {}, enabled: has });
  c.implement("design.setDefault", { run: () => void setDefault(), enabled: has });
  return { themesPanel, styleSetPanel, colorsPanel, fontsPanel, effectsPanel, get spacingMenu() { return spacingMenu(); } };
}
