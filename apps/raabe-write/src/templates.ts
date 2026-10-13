// Raabeschule templates (Datei > Neu): Arbeitsblatt, Facharbeit, Versuchsprotokoll, Handout.
// Our own designs (D-035): ordinary Word formatting (styles, lists, tables, headers, fields),
// so the files work in Word too. Placeholders in brackets are typed over.

import {
  A4, Doc, Edits, OBJ, School, addList, builtinStyle, bulletLevels, defaultStyleSheet, makePara, numberLevels,
  type DocState, type Editor, type Para, type ParaProps, type StyleSheet,
} from "@raabe/text-engine";
import { headingLevels } from "./features/lists";

export type TemplateId = "arbeitsblatt" | "facharbeit" | "versuchsprotokoll" | "handout";

export const TEMPLATES: { id: TemplateId; name: string; desc: string }[] = [
  { id: "arbeitsblatt", name: "Arbeitsblatt", desc: "Kopf mit Name, Klasse und Datum, Aufgaben, Seitenzahlen" },
  { id: "facharbeit", name: "Facharbeit", desc: "Deckblatt, Inhaltsverzeichnis, nummerierte Kapitel, 1,5-zeilig, Korrekturrand" },
  { id: "versuchsprotokoll", name: "Versuchsprotokoll", desc: "Fragestellung, Material, Durchführung, Beobachtung, Auswertung" },
  { id: "handout", name: "Handout", desc: "Handout zum Referat: Gliederung, Begriffe, Kernaussagen, Quellen" },
];

const CM = 28.3465;
const STYLES = ["heading 1", "heading 2", "heading 3", "Title", "Subtitle", "List Paragraph", "Table Grid", "header", "footer",
  "toc 1", "toc 2", "toc 3", "TOC Heading", "No Spacing", "footnote text", "footnote reference", "Hyperlink"];

function sheet(): StyleSheet {
  let s = defaultStyleSheet();
  for (const n of STYLES) { const b = builtinStyle(n); if (b) s = s.withStyle(b); }
  return s;
}

const id = (name: string) => builtinStyle(name)!.id;
const P = (text: string, props: ParaProps = {}, bold = false): Para => makePara(text, props, bold ? { bold: true } : {});
const field = (instr: string, display: string) => Edits.fieldObject(instr, display);

/** A footer "Seite X von Y" in the centre. */
function pageFooter(d: Doc): void {
  const text = `Seite ${OBJ} von ${OBJ}`;
  d.setStory("ftr:1", [{ ...makePara(text, { style: id("footer"), align: "center" }), objects: [field("PAGE", "1"), field("NUMPAGES", "1")] }]);
}

/** Adds a table of `rows` (cell texts) at the end; bold first row when `head`. */
function table(d: Doc, rows: string[][], widths: number[] | null, head = false): void {
  const at = { p: d.paras.length - 1, o: 0 };
  const total = d.settings.section.page.width - d.settings.section.page.marginLeft - d.settings.section.page.marginRight;
  Edits.insertTable(d, at, rows.length, rows[0].length, total, id("Table Grid"), {});
  let k = 0;
  for (let i = 0; i < d.paras.length && k < rows.length * rows[0].length; i++) {
    const c = d.paras[i].cell;
    if (!c || i < at.p) continue;
    const text = rows[c.row][c.col];
    if (text) d.insertText({ p: i, o: 0 }, { p: i, o: 0 }, text, head && c.row === 0 ? { bold: true } : {});
    k++;
  }
  if (widths) {
    const w = widths.map((x) => x * total);
    Edits.updateTable(d, at.p, (t) => ({ ...t, grid: w, rows: t.rows.map((r) => ({ ...r, cells: r.cells.map((c, i) => ({ ...c, width: w[i] })) })) }));
  }
}

function base(margins: { top: number; bottom: number; left: number; right: number }): Doc {
  const d = new Doc({ paras: [makePara("")], styles: sheet() });
  d.setSettings({
    ...d.settings,
    section: { ...d.settings.section, page: { ...A4, marginTop: margins.top * CM, marginBottom: margins.bottom * CM, marginLeft: margins.left * CM, marginRight: margins.right * CM } },
  });
  return d;
}

/** Appends paragraphs before the last (empty) paragraph. */
function add(d: Doc, ...paras: Para[]): void {
  d.splice(d.paras.length - 1, d.paras.length - 2, paras);
}

function list(d: Doc, kind: "bullet" | "number", items: string[]): void {
  const made = addList(d.numbering, kind === "bullet" ? bulletLevels() : numberLevels());
  d.setNumbering(made.numbering);
  add(d, ...items.map((t) => P(t, { style: id("List Paragraph"), numId: made.numId, ilvl: 0 })));
}

function footerAndHeader(d: Doc, header?: string): void {
  pageFooter(d);
  if (header) d.setStory("hdr:1", [makePara(header, { style: id("header") })]);
  const s = d.settings.section;
  d.setSettings({ ...d.settings, section: { ...s, footers: { default: "ftr:1" }, headers: header ? { default: "hdr:1" } : {} } });
}

// ----- the templates ----------------------------------------------------------------------------

function arbeitsblatt(): Doc {
  const d = base({ top: 2, bottom: 2, left: 2, right: 2 });
  table(d, [["Name:", "Klasse:", "Datum:"]], [0.5, 0.2, 0.3]);
  add(d,
    P(""),
    P("Arbeitsblatt: [Thema]", { style: id("Title") }),
    P("[Fach] · [Unterrichtseinheit]", { style: id("Subtitle") }),
    P("Aufgabe 1", { style: id("heading 2") }),
    P("[Aufgabenstellung, z. B.: Lies den Text M1 und fasse die wichtigsten Aussagen in eigenen Worten zusammen.]"),
    P(""),
    P("Aufgabe 2", { style: id("heading 2") }),
    P("[Aufgabenstellung]"),
    P(""),
    P("Aufgabe 3", { style: id("heading 2") }),
    P("[Aufgabenstellung]"),
  );
  footerAndHeader(d, "Raabeschule\t\t[Fach] · Klasse [ ]");
  return d;
}

function versuchsprotokoll(): Doc {
  const d = base({ top: 2.5, bottom: 2, left: 2.5, right: 2.5 });
  add(d, P("Versuchsprotokoll: [Name des Versuchs]", { style: id("Title") }));
  table(d, [["Name:", "Klasse/Kurs:"], ["Fach:", "Datum:"]], [0.6, 0.4]);
  add(d,
    P(""),
    P("Fragestellung", { style: id("heading 1") }),
    P("[Welche Frage soll der Versuch beantworten?]"),
    P("Vermutung", { style: id("heading 1") }),
    P("[Was erwartest du und warum?]"),
    P("Material und Geräte", { style: id("heading 1") }),
  );
  list(d, "bullet", ["[Gerät]", "[Chemikalie / Material]", "[Schutzbrille]"]);
  add(d, P("Durchführung", { style: id("heading 1") }));
  list(d, "number", ["[Erster Arbeitsschritt]", "[Zweiter Arbeitsschritt]", "[Dritter Arbeitsschritt]"]);
  add(d, P("[Versuchsskizze: Einfügen > Bilder]"), P("Beobachtung", { style: id("heading 1") }), P("[Was hast du gesehen, gemessen, gerochen …? Nur beschreiben, nicht deuten.]"));
  table(d, [["Zeit t in s", "Messwert", "Bemerkung"], ["", "", ""], ["", "", ""], ["", "", ""]], null, true);
  add(d,
    P(""),
    P("Auswertung", { style: id("heading 1") }),
    P("[Deute die Beobachtungen. Rechnungen mit Einfügen > Formel (Alt+Umschalt+0), z. B.:]"),
    { ...makePara(OBJ), objects: [{ kind: "math", data: undefined, display: "", latex: "v=\\frac{\\Delta s}{\\Delta t}" }] },
    P("Fehlerbetrachtung", { style: id("heading 1") }),
    P("[Welche Fehlerquellen gab es? Wie könnte man den Versuch verbessern?]"),
    P("Ergebnis", { style: id("heading 1") }),
    P("[Beantworte die Fragestellung. Wurde die Vermutung bestätigt?]"),
  );
  footerAndHeader(d, "Raabeschule\t\tVersuchsprotokoll");
  return d;
}

function handout(): Doc {
  const d = base({ top: 2, bottom: 2, left: 2, right: 2 });
  add(d,
    P("[Thema des Referats]", { style: id("Title") }),
    P("Handout zum Referat von [Name] · [Fach], [Kurs] · [Datum]", { style: id("Subtitle") }),
    P("Gliederung", { style: id("heading 1") }),
  );
  list(d, "number", ["[Einleitung]", "[Hauptpunkt]", "[Hauptpunkt]", "[Fazit]"]);
  add(d, P("Wichtige Begriffe", { style: id("heading 1") }));
  table(d, [["Begriff", "Erklärung"], ["[Begriff]", "[Erklärung]"], ["[Begriff]", "[Erklärung]"]], [0.3, 0.7], true);
  add(d, P(""), P("Kernaussagen", { style: id("heading 1") }));
  list(d, "bullet", ["[Kernaussage]", "[Kernaussage]", "[Kernaussage]"]);
  add(d, P("Quellen", { style: id("heading 1") }), P("[Autor: Titel. Verlag, Ort Jahr.]"), P("[Website: Titel. URL (abgerufen am TT.MM.JJJJ).]"));
  footerAndHeader(d, "Raabeschule\t\t[Fach] · [Kurs]");
  return d;
}

function facharbeit(): Doc {
  const d = base({ top: 2.5, bottom: 2, left: 2.5, right: 4 });
  // Facharbeit: Times New Roman 12 pt, 1,5-zeilig, Blocksatz (common requirements in Niedersachsen).
  let s = d.styles;
  const normal = s.get(id("Normal"))!;
  s = s.withStyle({ ...normal, rPr: { ...normal.rPr, font: "Times New Roman", fontTheme: undefined, size: 12 }, pPr: { ...normal.pPr, line: 1.5, lineRule: "auto", after: 6, align: "both" }, data: undefined });
  for (const k of [1, 2, 3]) {
    const h = s.get(id(`heading ${k}`))!;
    s = s.withStyle({ ...h, rPr: { ...h.rPr, font: "Times New Roman", fontTheme: undefined, color: "000000", colorTheme: undefined, bold: true, size: [16, 14, 12][k - 1] }, pPr: { ...h.pPr, before: [18, 12, 12][k - 1], after: 6 }, data: undefined });
  }
  d.setStyles(s);
  const center = (text: string, size: number, bold = false, after = 0): Para => ({ ...makePara(text, { align: "center", after, line: 1, lineRule: "auto" }, { size, bold }) });
  add(d,
    center("Raabeschule", 14, true, 6),
    center("Facharbeit im Fach [Fach]", 14, false, 120),
    center("[Thema der Facharbeit]", 20, true, 12),
    center("[Untertitel]", 14, false, 160),
    { ...makePara("Verfasserin/Verfasser:\t[Vor- und Nachname]", { tabs: [{ pos: 6.5 * CM, align: "left" }], line: 1.15, lineRule: "auto" }) },
    { ...makePara("Kurs:\t[Kurs]", { tabs: [{ pos: 6.5 * CM, align: "left" }], line: 1.15, lineRule: "auto" }) },
    { ...makePara("Betreuende Lehrkraft:\t[Name]", { tabs: [{ pos: 6.5 * CM, align: "left" }], line: 1.15, lineRule: "auto" }) },
    { ...makePara("Abgabetermin:\t[TT.MM.JJJJ]", { tabs: [{ pos: 6.5 * CM, align: "left" }], line: 1.15, lineRule: "auto" }) },
    { ...makePara("Schuljahr:\t[20.. / 20..]", { tabs: [{ pos: 6.5 * CM, align: "left" }], line: 1.15, lineRule: "auto", pageBreakBefore: false }) },
  );
  // Inhaltsverzeichnis on its own page (filled in after loading, see finishTemplate).
  add(d, { ...makePara("Inhaltsverzeichnis", { style: id("TOC Heading"), pageBreakBefore: true, numId: 0 }) }, P("[Inhaltsverzeichnis]"));
  const chapter = (title: string, text: string, level = 1, pageBreak = false, numbered = true) =>
    [P(title, { style: id(`heading ${level}`), pageBreakBefore: pageBreak || undefined, numId: numbered ? undefined : 0 }), P(text)];
  add(d,
    ...chapter("Einleitung", "[Hinführung zum Thema, Fragestellung und Aufbau der Arbeit.]", 1, true),
    ...chapter("Hauptteil", "[Kapitel des Hauptteils. Überschriften mit Start > Formatvorlagen (Überschrift 1, 2, 3) – sie erscheinen im Inhaltsverzeichnis.]"),
    ...chapter("Unterkapitel", "[Text. Belege mit Fußnoten: Referenzen > Fußnote einfügen (Strg+Alt+F).]", 2),
    ...chapter("Unterkapitel", "[Text]", 2),
    ...chapter("Fazit", "[Zusammenfassung der Ergebnisse und Beantwortung der Fragestellung.]"),
    ...chapter("Literaturverzeichnis", "[Nachname, Vorname: Titel. Verlag, Ort Jahr.]", 1, true, false),
    ...chapter("Anhang", "[Materialien, Tabellen, Abbildungen]", 1, true, false),
    P("Erklärung", { style: id("heading 1"), pageBreakBefore: true, numId: 0 }),
    P("Hiermit erkläre ich, dass ich die vorliegende Facharbeit selbstständig angefertigt, keine anderen als die angegebenen Hilfsmittel benutzt und die Stellen der Facharbeit, die im Wortlaut oder im wesentlichen Inhalt aus anderen Werken entnommen wurden, mit genauer Quellenangabe kenntlich gemacht habe."),
    P(""), P(""),
    { ...makePara("____________________________\t____________________________", { tabs: [{ pos: 8.5 * CM, align: "left" }] }) },
    { ...makePara("Ort, Datum\tUnterschrift", { tabs: [{ pos: 8.5 * CM, align: "left" }] }) },
  );
  // No page number on the title page.
  pageFooter(d);
  const sec = d.settings.section;
  d.setStory("ftr:2", [makePara("", { style: id("footer") })]);
  d.setSettings({ ...d.settings, section: { ...sec, titlePage: true, footers: { default: "ftr:1", first: "ftr:2" } } });
  return d;
}

/** The document a template starts with. */
export function templateState(t: TemplateId): Partial<DocState> {
  const d = { arbeitsblatt, facharbeit, versuchsprotokoll, handout }[t]();
  // Drop the trailing empty paragraph left from building, unless it follows a table.
  const last = d.paras.length - 1;
  if (last > 0 && !d.paras[last].text && !d.paras[last - 1].cell) d.splice(last, last, []);
  return d.state();
}

/** Work that needs the layout: heading numbers and the table of contents (Facharbeit). */
export function finishTemplate(t: TemplateId, e: Editor): void {
  if (t !== "facharbeit") return;
  e.numberHeadings(headingLevels());
  // The Erklärung stays without a number (numId 0 on the paragraph).
  const toc = e.doc.main.findIndex((p) => p.text === "[Inhaltsverzeichnis]");
  if (toc >= 0) {
    e.selectRange(toc, 0, 0);
    e.insertTocAt(toc);
  }
  e.selectRange(0, 0, 0);
  e.clearHistory();
}

export { School };
