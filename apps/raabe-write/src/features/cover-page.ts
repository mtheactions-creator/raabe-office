// Einfügen > Seiten > Deckblatt: Word's gallery (Integriert, three columns) with RaabeWrite's
// own designs for school work (Word's designs are Microsoft's building blocks), and Word's
// entries below it (recorded 2026-10-14): Weitere Deckblätter von Office.com, Aktuelles
// Deckblatt entfernen, Auswahl im Deckblattkatalog speichern... The cover page is stored as
// Word stores its own (a Cover Pages sdt at the start of the document), so Word replaces or
// removes it too. Placeholders are bracketed; the first one is selected after inserting.

import { h, icon } from "@raabe/shell";
import { makePara, type BorderSpec, type CharProps, type Para, type ParaProps } from "@raabe/text-engine";
import type { AppApi } from "./api";

interface Design {
  readonly name: string;
  readonly tip: string;
  readonly build: (accent: string) => Para[];
}

const ACCENT: CharProps = { colorTheme: 'w:themeColor="accent1"' };
const line = (color: string, size: number, space: number, style = "single"): BorderSpec => ({ style, size, color, space });
const P = (text: string, props: ParaProps = {}, run: CharProps = {}): Para => makePara(text, props, run, run);
const title = (text: string, props: ParaProps, run: CharProps = {}) => P(text, { style: "Titel", ...props }, { size: 32, ...run });
const subtitle = (text: string, props: ParaProps, run: CharProps = {}) => P(text, { style: "Untertitel", ...props }, { size: 15, ...run });
const accentRun = (accent: string, run: CharProps = {}): CharProps => ({ ...ACCENT, color: accent, ...run });

export const COVER_DESIGNS: Design[] = [
  { name: "Schlicht", tip: "Titel und Untertitel in der Mitte, Name und Datum unten", build: () => [
    title("[Dokumenttitel]", { align: "center", before: 190 }, { size: 36 }),
    subtitle("[Untertitel des Dokuments]", { align: "center", before: 6 }),
    P("[Name]", { align: "center", before: 270, after: 0 }),
    P("[Datum]", { align: "center", after: 0 }),
  ] },
  { name: "Facharbeit", tip: "Schule, Fach und Thema, darunter Verfasser/in, Kurs, Lehrkraft und Abgabetermin", build: (accent) => {
    const tab: ParaProps = { tabs: [{ pos: 170, align: "left" }], after: 6 };
    return [
      P("[Name der Schule]", { align: "center", after: 0 }, { size: 14, bold: true }),
      P("Schuljahr [Schuljahr]", { align: "center" }),
      P("Facharbeit im Fach [Fach]", { align: "center", before: 120, after: 12 }, accentRun(accent, { size: 16 })),
      title("[Thema der Facharbeit]", { align: "center", after: 0 }, { size: 26 }),
      P("Verfasser/in:\t[Name]", { ...tab, before: 230 }),
      P("Kurs:\t[Kurs]", tab),
      P("Betreuende Lehrkraft:\t[Name der Lehrkraft]", tab),
      P("Abgabetermin:\t[Datum]", tab),
    ];
  } },
  { name: "Referat", tip: "Thema mit einer Linie in der Akzentfarbe, Name, Kurs und Datum unten", build: (accent) => [
    title("[Thema des Referats]", { before: 160, after: 6, borders: { bottom: line(accent, 3, 4) } }),
    subtitle("Referat im Fach [Fach]", { after: 0 }),
    P("[Name]", { before: 300, after: 0 }, { bold: true }),
    P("[Kurs/Klasse]", { after: 0 }),
    P("[Datum]", { after: 0 }),
  ] },
  { name: "Seitenleiste", tip: "Ein Balken in der Akzentfarbe links neben Titel und Untertitel", build: (accent) => {
    const bar: ParaProps = { indLeft: 18, borders: { left: line(accent, 6, 12) } };
    return [
      title("[Dokumenttitel]", { ...bar, before: 150, after: 0 }, { size: 36 }),
      subtitle("[Untertitel des Dokuments]", { ...bar, after: 0 }),
      P("[Name]", { indLeft: 18, before: 290, after: 0 }, accentRun(accent, { bold: true })),
      P("[Kurs/Klasse] · [Datum]", { indLeft: 18, after: 0 }),
    ];
  } },
  { name: "Rahmen", tip: "Titel und Untertitel in einem doppelten Rahmen", build: (accent) => {
    const box: ParaProps = { align: "center", indLeft: 30, indRight: 30, borders: { top: line(accent, 1.5, 12, "double"), bottom: line(accent, 1.5, 12, "double"), left: line(accent, 1.5, 12, "double"), right: line(accent, 1.5, 12, "double") } };
    return [
      title("[Dokumenttitel]", { ...box, before: 170, after: 0 }),
      subtitle("[Untertitel des Dokuments]", { ...box, after: 0 }),
      P("[Name]", { align: "center", before: 280, after: 0 }),
      P("[Datum]", { align: "center", after: 0 }),
    ];
  } },
  { name: "Farbfläche", tip: "Titel und Untertitel weiß auf einer Fläche in der Akzentfarbe", build: (accent) => {
    const band: ParaProps = { align: "center", shading: { fill: accent } };
    const white: CharProps = { color: "FFFFFF" };
    return [
      title("[Dokumenttitel]", { ...band, before: 170, after: 0, line: 1.3, lineRule: "auto" }, { ...white, size: 34 }),
      subtitle("[Untertitel des Dokuments]", { ...band, after: 0, line: 1.6, lineRule: "auto" }, white),
      P("[Name]", { align: "center", before: 280, after: 0 }, { bold: true }),
      P("[Kurs/Klasse]", { align: "center", after: 0 }),
      P("[Datum]", { align: "center", after: 0 }),
    ];
  } },
  { name: "Linie", tip: "Titel zwischen zwei feinen Linien, Name und Datum rechts unten", build: (accent) => [
    title("[Dokumenttitel]", { align: "center", before: 180, after: 0, borders: { top: line(accent, 0.75, 10), bottom: line(accent, 0.75, 10) } }, { size: 30 }),
    P("[UNTERTITEL DES DOKUMENTS]", { align: "center", before: 14 }, accentRun(accent, { size: 12, spacing: 2 })),
    P("[Name]", { align: "right", before: 300, after: 0 }),
    P("[Datum]", { align: "right", after: 0 }),
  ] },
];

/** A thumbnail of a design: the little page drawn from the paragraphs themselves. */
function thumbnail(paras: Para[], page: { width: number; height: number; marginLeft: number; marginRight: number; marginTop: number }, fonts: { major: string; minor: string }): HTMLElement {
  const W = 120;
  const k = W / page.width;
  const px = (pt: number) => `${(pt * k).toFixed(2)}px`;
  const el = h("span", { class: "cp-page", style: `width:${W}px;height:${Math.round(page.height * k)}px;padding:${px(page.marginTop)} ${px(page.marginRight)} 0 ${px(page.marginLeft)}` });
  const edge = (b: BorderSpec | undefined) => (b ? `${Math.max(1, b.size * k).toFixed(2)}px ${b.style === "double" ? "double" : "solid"} #${b.color}` : "none");
  for (const p of paras) {
    const run = p.runs[0]?.props ?? {};
    const heading = p.props.style === "Titel";
    const size = run.size ?? (heading ? 28 : p.props.style === "Untertitel" ? 11 : 11);
    const b = p.props.borders ?? {};
    const style = [
      `margin:${px(p.props.before ?? 0)} ${px(p.props.indRight ?? 0)} ${px(p.props.after ?? 8)} ${px(p.props.indLeft ?? 0)}`,
      `font:${run.bold ? "700 " : ""}${px(size)}/1.25 "${heading ? fonts.major : fonts.minor}",sans-serif`,
      `text-align:${p.props.align === "center" ? "center" : p.props.align === "right" ? "right" : "left"}`,
      `color:${run.color ? `#${run.color}` : heading ? "#000" : "#444"}`,
      `border-top:${edge(b.top)};border-bottom:${edge(b.bottom)};border-left:${edge(b.left)};border-right:${edge(b.right)}`,
      `padding:${px(b.top?.space ?? 0)} ${px(b.right?.space ?? 0)} ${px(b.bottom?.space ?? 0)} ${px(b.left?.space ?? 0)}`,
      p.props.shading ? `background:#${p.props.shading.fill}` : "",
      `tab-size:${px(p.props.tabs?.[0]?.pos ?? 36)}`,
    ].filter(Boolean).join(";");
    el.append(h("span", { class: "cp-para", style }, p.text || " "));
  }
  return el;
}

export function coverPageFeature(api: AppApi) {
  const ed = () => api.editor();
  const c = api.commands;
  c.implement("insert.coverPage", { run: () => {}, enabled: () => !!ed() });

  const accent = () => (ed()?.theme.colors?.accent1 ?? "4472C4").replace(/^#/, "").toUpperCase();

  const panel = (close: () => void) => {
    const e = ed();
    const page = e?.section().page ?? { width: 595.3, height: 841.9, marginLeft: 70.85, marginRight: 70.85, marginTop: 70.85 };
    const fonts = { major: e?.theme.majorFont ?? "Calibri Light", minor: e?.theme.minorFont ?? "Calibri" };
    const el = h("div", { class: "rpanel designpanel cover" }, h("div", { class: "title" }, "Integriert"));
    el.append(h("div", { class: "cp-grid" }, ...COVER_DESIGNS.map((d) => h("button", {
      class: "wm-tile", "data-tip": `${d.name}\n${d.tip}`, "aria-label": d.name,
      onclick: () => {
        close();
        const x = ed();
        if (!x) return;
        x.insertCoverPage(d.build(accent()));
        api.refresh();
        api.focus();
      },
    }, thumbnail(d.build(accent()), page, fonts), h("span", { class: "dz-name" }, d.name)))));
    const cmd = (label: string, ic: string | null, run: () => void, disabled = false, tip?: string) =>
      h("button", { class: "dz-cmd", disabled, ...(tip ? { "data-tip": tip } : {}), onclick: () => { close(); run(); } }, ic ? icon(ic) : h("span", { class: "dz-noicon" }), label);
    el.append(h("hr"),
      cmd("Weitere Deckblätter von Office.com", "globe", () => {}, true, "Weitere Deckblätter von Office.com\nRaabeWrite arbeitet ohne Internet."),
      cmd("Aktuelles Deckblatt entfernen", "close", () => { ed()?.removeCoverPage(); api.refresh(); api.focus(); }),
      cmd("Auswahl im Deckblattkatalog speichern...", "save", () => {}, true));
    return el;
  };
  return { panel };
}
