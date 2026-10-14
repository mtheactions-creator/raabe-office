// Zitate und Literaturverzeichnis as Word does them: the document's sources (Word's
// bibliography XML, b:Sources in a customXml part), CITATION fields in the text and a
// BIBLIOGRAPHY field whose result is the formatted list. The formatting follows the patterns of
// Word's citation styles (APA 6th, Chicago 15th, Harvard - Anglia, IEEE, ISO 690, MLA 7th …).

export const B_NS = "http://schemas.openxmlformats.org/officeDocument/2006/bibliography";

export type SourceType = "Book" | "BookSection" | "JournalArticle" | "ArticleInAPeriodical" | "ConferenceProceedings" | "Report" | "InternetSite"
  | "DocumentFromInternetSite" | "ElectronicSource" | "Art" | "SoundRecording" | "Performance" | "Film" | "Interview" | "Patent" | "Case" | "Misc";

export interface Person { readonly last: string; readonly first?: string; readonly middle?: string }

/** A contributor role (b:Author/b:Author, b:Author/b:Editor …): people or a company. */
export interface Contributors { readonly people?: readonly Person[]; readonly corporate?: string }

export interface Source {
  readonly tag: string;
  readonly type: SourceType;
  readonly guid?: string;
  readonly lcid?: string;
  /** Roles (Author, BookAuthor, Editor, Translator, Artist, Composer, Performer, Writer, Director, Interviewee, Interviewer, Inventor). */
  readonly contributors?: Readonly<Record<string, Contributors>>;
  /** Plain fields by Word's element name (Title, Year, City, Publisher, Pages, URL …). */
  readonly fields: Readonly<Record<string, string>>;
  /** Neuen Platzhalter hinzufügen: a source without data yet. */
  readonly placeholder?: boolean;
}

export interface Bibliography {
  readonly sources: readonly Source[];
  /** Word's style name ("APA") and its XSL file and version as Word writes them. */
  readonly style: string;
}

/** Word's citation styles: menu label, edition, StyleName, file, version. */
export const CITATION_STYLES: { label: string; edition: string; name: string; file: string; version: string }[] = [
  { label: "APA", edition: "6. Auflage", name: "APA", file: "\\APASixthEditionOfficeOnline.xsl", version: "6" },
  { label: "Chicago", edition: "15. Auflage", name: "Chicago", file: "\\CHICAGO.XSL", version: "15" },
  { label: "GB7714", edition: "2005", name: "GB7714", file: "\\GB.XSL", version: "2005" },
  { label: "Gost - Namenssortierung", edition: "2003", name: "GOST - Name Sort", file: "\\GostName.XSL", version: "2003" },
  { label: "Gost - Titelsortierung", edition: "2003", name: "GOST - Title Sort", file: "\\GostTitle.XSL", version: "2003" },
  { label: "Harvard - Anglia", edition: "2008", name: "Harvard - Anglia", file: "\\HarvardAnglia2008OfficeOnline.xsl", version: "2008" },
  { label: "IEEE", edition: "2006", name: "IEEE", file: "\\IEEE2006OfficeOnline.xsl", version: "2006" },
  { label: "ISO 690 - Erstes Element und Datum", edition: "1987", name: "ISO 690 - First Element and Date", file: "\\ISO690.XSL", version: "1987" },
  { label: "ISO 690 - Numerische Referenz", edition: "1987", name: "ISO 690 - Numerical Reference", file: "\\ISO690Nmerical.XSL", version: "1987" },
  { label: "MLA", edition: "7. Auflage", name: "MLA", file: "\\MLASeventhEditionOfficeOnline.xsl", version: "7" },
  { label: "SIST02", edition: "2003", name: "SIST02", file: "\\SIST02.XSL", version: "2003" },
  { label: "Turabian", edition: "6. Auflage", name: "Turabian", file: "\\TURABIAN.XSL", version: "6" },
];

export const styleByName = (name: string) => CITATION_STYLES.find((s) => s.name === name || s.label === name) ?? CITATION_STYLES[0];

/** Word's source types with the German names and the fields its dialog asks for (recorded). */
export const SOURCE_TYPES: { type: SourceType; label: string; fields: [string, string][] }[] = [
  { type: "Book", label: "Buch", fields: [["Author", "Autor"], ["Title", "Titel"], ["Year", "Jahr"], ["City", "Ort"], ["Publisher", "Verleger"]] },
  { type: "BookSection", label: "Buchabschnitt", fields: [["Author", "Autor"], ["Title", "Titel"], ["BookAuthor", "Buchautor"], ["BookTitle", "Buchtitel"], ["Year", "Jahr"], ["Pages", "Seiten"], ["City", "Ort"], ["Publisher", "Verleger"]] },
  { type: "JournalArticle", label: "Journalartikel", fields: [["Author", "Autor"], ["Title", "Titel"], ["JournalName", "Journalname"], ["Year", "Jahr"], ["Pages", "Seiten"]] },
  { type: "ArticleInAPeriodical", label: "Artikel in einer Zeitschrift", fields: [["Author", "Autor"], ["Title", "Titel"], ["PeriodicalTitle", "Titel der Zeitschrift"], ["Year", "Jahr"], ["Month", "Monat"], ["Day", "Tag"], ["Pages", "Seiten"]] },
  { type: "ConferenceProceedings", label: "Konferenzprotokoll", fields: [["Author", "Autor"], ["Title", "Titel"], ["Pages", "Seiten"], ["Year", "Jahr"], ["ConferenceName", "Name der Konferenzpublikation"], ["City", "Ort"], ["Publisher", "Verleger"]] },
  { type: "Report", label: "Bericht", fields: [["Author", "Autor"], ["Title", "Titel"], ["Year", "Jahr"], ["Publisher", "Verleger"], ["City", "Ort"]] },
  { type: "InternetSite", label: "Website", fields: [["Author", "Autor"], ["Title", "Name der Webseite"], ["InternetSiteTitle", "Name der Website"], ["Year", "Jahr"], ["Month", "Monat"], ["Day", "Tag"], ["URL", "URL"]] },
  { type: "DocumentFromInternetSite", label: "Dokument von Website", fields: [["Author", "Autor"], ["Title", "Name der Webseite"], ["InternetSiteTitle", "Name der Website"], ["Year", "Jahr"], ["Month", "Monat"], ["Day", "Tag"], ["URL", "URL"]] },
  { type: "ElectronicSource", label: "Elektronische Quelle", fields: [["Author", "Autor"], ["Title", "Titel"], ["City", "Ort"], ["StateProvince", "Bundesland/Kanton"], ["CountryRegion", "Land/Region"], ["Year", "Jahr"], ["Month", "Monat"], ["Day", "Tag"]] },
  { type: "Art", label: "Kunst", fields: [["Artist", "Künstler"], ["Title", "Titel"], ["Institution", "Institution"], ["PublicationTitle", "Titel der Publikation"], ["City", "Ort"]] },
  { type: "SoundRecording", label: "Tonaufnahme", fields: [["Composer", "Komponist"], ["Performer", "Interpret"], ["Title", "Titel"], ["Year", "Jahr"], ["City", "Ort"], ["StateProvince", "Bundesland/Kanton"], ["CountryRegion", "Land/Region"]] },
  { type: "Performance", label: "Aufführung", fields: [["Title", "Titel"], ["Writer", "Verfasser"], ["Performer", "Interpret"], ["Theater", "Theater"], ["City", "Ort"], ["StateProvince", "Bundesland/Kanton"], ["CountryRegion", "Land/Region"], ["Year", "Jahr"], ["Month", "Monat"], ["Day", "Tag"]] },
  { type: "Film", label: "Film", fields: [["Title", "Titel"], ["Director", "Regisseur"], ["Year", "Jahr"]] },
  { type: "Interview", label: "Interview", fields: [["Interviewee", "Befragte Person"], ["Title", "Titel"], ["Interviewer", "Interviewer"], ["Year", "Jahr"], ["Month", "Monat"], ["Day", "Tag"]] },
  { type: "Patent", label: "Patent", fields: [["Inventor", "Erfinder"], ["Year", "Jahr"], ["CountryRegion", "Land/Region"], ["PatentNumber", "Patentnummer"]] },
  { type: "Case", label: "Fall", fields: [["Title", "Titel"], ["CaseNumber", "Fallnummer"], ["Court", "Gericht"], ["Year", "Jahr"], ["Month", "Monat"], ["Day", "Tag"]] },
  { type: "Misc", label: "Verschiedenes", fields: [["Author", "Autor"], ["Title", "Titel"], ["PublicationTitle", "Titel der Publikation"], ["Year", "Jahr"], ["Month", "Monat"], ["Day", "Tag"], ["City", "Ort"], ["StateProvince", "Bundesland/Kanton"], ["CountryRegion", "Land/Region"], ["Publisher", "Verleger"]] },
];

/** Further fields "Alle Literaturverzeichnisfelder anzeigen" offers. */
export const MORE_FIELDS: [string, string][] = [
  ["Editor", "Herausgeber"], ["Translator", "Übersetzer"], ["Edition", "Auflage"], ["Volume", "Band"], ["NumberVolumes", "Anzahl Bände"],
  ["ShortTitle", "Kurztitel"], ["StandardNumber", "Standardnummer"], ["Medium", "Medientyp"], ["YearAccessed", "Jahr des Zugriffs"],
  ["MonthAccessed", "Monat des Zugriffs"], ["DayAccessed", "Tag des Zugriffs"], ["DOI", "DOI"], ["Comments", "Kommentare"],
];

export const ROLES = new Set(["Author", "BookAuthor", "Editor", "Translator", "Artist", "Composer", "Performer", "Writer", "Director", "Interviewee", "Interviewer", "Inventor", "Conductor", "Counsel", "Compiler", "ProducerName"]);

// ----- names ----------------------------------------------------------------------------------

/** "Muster, Max; Beispiel, Erika" (Word's Autor field) as people. */
export function parsePeople(text: string): Person[] {
  return text.split(";").map((x) => x.trim()).filter(Boolean).map((x) => {
    if (x.includes(",")) {
      const [last, rest] = x.split(",", 2).map((y) => y.trim());
      const [first, ...middle] = rest.split(/\s+/);
      return { last, first: first || undefined, middle: middle.join(" ") || undefined };
    }
    const parts = x.split(/\s+/);
    const last = parts.pop()!;
    return { last, first: parts[0], middle: parts.slice(1).join(" ") || undefined };
  });
}

export const peopleText = (c: Contributors | undefined) => c?.corporate ?? (c?.people ?? []).map((p) => [p.last, [p.first, p.middle].filter(Boolean).join(" ")].filter(Boolean).join(", ")).join("; ");

const initials = (p: Person) => [p.first, p.middle].filter(Boolean).flatMap((n) => n!.split(/[\s-]+/)).map((n) => `${n[0]}.`).join(" ");

/** A short tag from the first author and the year (Mus23), as Word suggests it. */
export function suggestTag(s: Pick<Source, "contributors" | "fields">, taken: readonly string[]): string {
  const who = firstContributor(s);
  const name = who?.corporate ?? who?.people?.[0]?.last ?? s.fields.Title ?? "Platzhalter";
  const base = `${name.replace(/[^\p{L}\p{N}]/gu, "").slice(0, 3)}${(s.fields.Year ?? "").slice(-2)}` || "Platzhalter";
  let tag = base, n = 1;
  while (taken.includes(tag)) tag = `${base}${++n}`;
  return tag;
}

function firstContributor(s: Pick<Source, "contributors">): Contributors | undefined {
  const c = s.contributors ?? {};
  for (const role of ["Author", "Artist", "Composer", "Writer", "Director", "Interviewee", "Inventor", "Editor", "Performer"]) if (c[role]?.people?.length || c[role]?.corporate) return c[role];
  return undefined;
}

// ----- formatting -------------------------------------------------------------------------------

/** A piece of a formatted entry; titles are italic in most styles. */
export interface Piece { readonly text: string; readonly italic?: boolean }

const join = (parts: (Piece | string | null | undefined | false)[]): Piece[] => parts.filter((x): x is Piece | string => !!x).map((x) => (typeof x === "string" ? { text: x } : x));

function names(c: Contributors | undefined, style: string, max = 7): string {
  if (!c) return "";
  if (c.corporate) return c.corporate;
  const ps = c.people ?? [];
  if (!ps.length) return "";
  const apa = (p: Person) => `${p.last}, ${initials(p)}`.replace(/, $/, "");
  const lastFirst = (p: Person) => [p.last, [p.first, p.middle].filter(Boolean).join(" ")].filter(Boolean).join(", ");
  const firstLast = (p: Person) => [[p.first, p.middle].filter(Boolean).join(" "), p.last].filter(Boolean).join(" ");
  const upper = (p: Person) => [p.last.toUpperCase(), [p.first, p.middle].filter(Boolean).join(" ")].filter(Boolean).join(", ");
  const ieee = (p: Person) => `${initials(p)} ${p.last}`.trim();
  const fmt = style === "APA" ? apa : style.startsWith("ISO") || style.startsWith("Harvard") ? upper : style === "IEEE" || style === "GB7714" ? ieee : lastFirst;
  const list = ps.slice(0, max);
  if (style === "APA") return list.length === 1 ? fmt(list[0]) : `${list.slice(0, -1).map(fmt).join(", ")}, & ${fmt(list[list.length - 1])}`;
  if (style === "MLA" || style === "Chicago" || style === "Turabian") {
    if (list.length === 1) return fmt(list[0]);
    return `${fmt(list[0])}${list.length > 2 ? `, ${list.slice(1, -1).map(firstLast).join(", ")},` : ""} und ${firstLast(list[list.length - 1])}`;
  }
  return list.map(fmt).join(style === "IEEE" ? ", " : "; ");
}

const lastNames = (c: Contributors | undefined) => c?.corporate ?? (c?.people ?? []).map((p) => p.last);

const end = (s: string) => (/[.!?]$/.test(s) ? s : `${s}.`);

/** One entry of the Literaturverzeichnis in a style (`n` for the numeric styles). */
export function formatEntry(s: Source, styleName: string, n: number): Piece[] {
  const st = styleByName(styleName).name;
  const f = s.fields;
  if (s.placeholder) return [{ text: s.tag }];
  const who = names(firstContributor(s), st);
  const title = f.Title ?? "";
  const container = f.BookTitle ?? f.JournalName ?? f.PeriodicalTitle ?? f.ConferenceName ?? f.InternetSiteTitle ?? f.PublicationTitle ?? "";
  const where = [f.City, f.Publisher].filter(Boolean).join(": ");
  const year = f.Year ?? "o. J.";
  const web = f.URL ? `Von ${f.URL} abgerufen` : "";
  const pages = f.Pages ? `S. ${f.Pages}` : "";
  const isPart = !!container;
  switch (st) {
    case "APA":
      return join([who && `${end(who)} `, `(${year}). `, isPart ? `${end(title)} ` : { text: end(title), italic: true }, isPart && " ",
        isPart && { text: container, italic: true }, isPart && (pages ? `, ${pages}` : ""), isPart && ". ", !isPart && " ", where && end(where), web && ` ${web}`]);
    case "MLA":
      return join([who && `${end(who)} `, isPart ? `„${end(title)}“ ` : { text: end(title), italic: true }, isPart && { text: container, italic: true }, isPart && (pages ? `, ${pages}` : ""),
        " ", where && `${where}, `, `${year}.`, f.URL && ` ${f.URL}.`]);
    case "Chicago": case "Turabian":
      return join([who && `${end(who)} `, isPart ? `„${end(title)}“ ` : { text: end(title), italic: true }, isPart && { text: container, italic: true }, " ",
        where && `${where}, `, `${year}.`, f.URL && ` ${f.URL}.`]);
    case "Harvard - Anglia":
      return join([who && `${who}, `, `${year}. `, isPart ? `${title}. ` : { text: `${title}. `, italic: true }, isPart && { text: `${container}. `, italic: true }, where && end(where), f.URL && ` [Online] Verfügbar unter: ${f.URL}`]);
    case "IEEE":
      return join([`[${n}] `, who && `${who}, `, isPart ? `„${title},“ ` : { text: `${title}, `, italic: true }, isPart && { text: `${container}, `, italic: true }, where && `${where}, `, `${year}.`, f.URL && ` [Online]. Available: ${f.URL}.`]);
    case "ISO 690 - First Element and Date":
      return join([who && `${end(who)} `, `${year}. `, { text: end(title), italic: true }, " ", container && `${end(container)} `, [f.City, f.Publisher].filter(Boolean).join(" : "), (f.City || f.Publisher) && ", ", `${year}.`]);
    case "ISO 690 - Numerical Reference":
      return join([`${n}. `, who && `${end(who)} `, { text: end(title), italic: true }, " ", container && `${end(container)} `, [f.City, f.Publisher].filter(Boolean).join(" : "), (f.City || f.Publisher) && ", ", `${year}.`]);
    case "GB7714":
      return join([`[${n}] `, who && `${who}. `, `${title}[M]. `, [f.City, f.Publisher].filter(Boolean).join(": "), `, ${year}.`]);
    case "GOST - Name Sort": case "GOST - Title Sort":
      return join([who && `${who} `, { text: title, italic: false }, where && ` - ${[f.City, f.Publisher].filter(Boolean).join(" : ")}`, `, ${year}.`]);
    case "SIST02":
      return join([`${n}) `, who && `${who}. `, `${title}. `, f.Publisher && `${f.Publisher}, `, `${year}.`]);
    default:
      return [{ text: title }];
  }
}

/** The text of a citation (Zitat) in the text, e.g. "(Muster, 2023)" in APA, "[1]" in IEEE. */
export function formatCitation(s: Source | undefined, tag: string, styleName: string, n: number, pages = ""): string {
  const st = styleByName(styleName).name;
  if (!s || s.placeholder) return `(${tag})`;
  const who = firstContributor(s);
  const lasts = lastNames(who);
  const ln = typeof lasts === "string" ? lasts : lasts.length > 2 ? `${lasts[0]}, et al.` : lasts.join(st === "APA" ? " & " : " und ");
  const name = ln || s.fields.Title || tag;
  const year = s.fields.Year ?? "o. J.";
  const p = pages ? `, S. ${pages}` : "";
  switch (st) {
    case "IEEE": case "GB7714": return `[${n}]`;
    case "ISO 690 - Numerical Reference": return `(${n})`;
    case "SIST02": return `(${n})`;
    case "MLA": return `(${name}${pages ? ` ${pages}` : ""})`;
    case "Chicago": case "Turabian": return `(${name} ${year}${p})`;
    default: return `(${name}, ${year}${p})`;
  }
}

/** The order of the Literaturverzeichnis: by author (or title), numbered styles by first citation. */
export function sortedSources(b: Bibliography, cited: readonly string[]): Source[] {
  const st = styleByName(b.style).name;
  const used = b.sources.filter((s) => cited.includes(s.tag) || !cited.length);
  if (["IEEE", "GB7714", "ISO 690 - Numerical Reference", "SIST02"].includes(st)) return [...used].sort((a, c) => cited.indexOf(a.tag) - cited.indexOf(c.tag));
  const key = (s: Source) => (st === "GOST - Title Sort" ? s.fields.Title ?? "" : (names(firstContributor(s), "MLA") || s.fields.Title || s.tag)).toLowerCase();
  return [...used].sort((a, c) => key(a).localeCompare(key(c), "de"));
}

// ----- Word's XML ----------------------------------------------------------------------------------

const esc = (s: string) => s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");

/** b:Sources as Word writes it (customXml/itemN.xml). */
export function sourcesXml(b: Bibliography): string {
  const st = styleByName(b.style);
  const people = (c: Contributors) => (c.corporate ? `<b:Corporate>${esc(c.corporate)}</b:Corporate>`
    : `<b:NameList>${(c.people ?? []).map((p) => `<b:Person><b:Last>${esc(p.last)}</b:Last>${p.first ? `<b:First>${esc(p.first)}</b:First>` : ""}${p.middle ? `<b:Middle>${esc(p.middle)}</b:Middle>` : ""}</b:Person>`).join("")}</b:NameList>`);
  const src = b.sources.map((s, i) => {
    const roles = Object.entries(s.contributors ?? {}).filter(([, c]) => c.corporate || c.people?.length);
    const author = roles.length ? `<b:Author>${roles.map(([r, c]) => `<b:${r}>${people(c)}</b:${r}>`).join("")}</b:Author>` : "";
    const fields = Object.entries(s.fields).filter(([, v]) => v).map(([k, v]) => `<b:${k}>${esc(v)}</b:${k}>`).join("");
    return `<b:Source><b:Tag>${esc(s.tag)}</b:Tag><b:SourceType>${s.type}</b:SourceType><b:Guid>${esc(s.guid ?? guid(i))}</b:Guid>${s.lcid ? `<b:LCID>${esc(s.lcid)}</b:LCID>` : ""}${author}${fields}<b:RefOrder>${i + 1}</b:RefOrder></b:Source>`;
  }).join("");
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<b:Sources SelectedStyle="${esc(st.file)}" StyleName="${esc(st.name)}" Version="${esc(st.version)}" xmlns:b="${B_NS}" xmlns="${B_NS}">${src}</b:Sources>`;
}

const guid = (i: number) => `{${["5A1B7C3D", "1E2F", "4A5B", "8C9D", (Date.now().toString(16) + String(i).padStart(4, "0")).slice(-12).toUpperCase().padStart(12, "0")].join("-")}}`;

/** b:Sources read from Word's customXml part (null when the XML is not Word's bibliography). */
export function readSources(doc: Document): Bibliography | null {
  const root = doc.documentElement;
  if (root?.localName !== "Sources" || root.namespaceURI !== B_NS) return null;
  const kid = (el: Element, name: string) => Array.from(el.children).find((c) => c.localName === name) ?? null;
  const textOf = (el: Element | null) => el?.textContent ?? "";
  const sources: Source[] = [];
  for (const s of Array.from(root.children).filter((c) => c.localName === "Source")) {
    const contributors: Record<string, Contributors> = {};
    const author = kid(s, "Author");
    for (const role of author ? Array.from(author.children) : []) {
      const corp = kid(role, "Corporate");
      const list = kid(role, "NameList");
      contributors[role.localName] = corp ? { corporate: textOf(corp) } : {
        people: (list ? Array.from(list.children).filter((p) => p.localName === "Person") : []).map((p) => ({
          last: textOf(kid(p, "Last")), first: textOf(kid(p, "First")) || undefined, middle: textOf(kid(p, "Middle")) || undefined,
        })),
      };
    }
    const fields: Record<string, string> = {};
    for (const c of Array.from(s.children)) {
      if (["Tag", "SourceType", "Guid", "LCID", "Author", "RefOrder"].includes(c.localName)) continue;
      fields[c.localName] = c.textContent ?? "";
    }
    sources.push({
      tag: textOf(kid(s, "Tag")), type: (textOf(kid(s, "SourceType")) || "Misc") as SourceType, guid: textOf(kid(s, "Guid")) || undefined,
      lcid: textOf(kid(s, "LCID")) || undefined, contributors, fields,
    });
  }
  const name = root.getAttribute("StyleName") ?? "APA";
  return { sources, style: styleByName(name).name };
}
