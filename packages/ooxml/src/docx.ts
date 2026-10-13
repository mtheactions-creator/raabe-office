// .docx reading and writing for the subset RaabeWrite supports so far: paragraphs of plain
// text with the styles "Standard" and "Überschrift 1". Everything else in a file is
// detected and reported, so the app can refuse to overwrite content it would lose (D-004).

import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";

export type ParagraphStyle = "normal" | "h1";

export interface DocxParagraph {
  style: ParagraphStyle;
  text: string;
}

export interface ReadResult {
  paragraphs: DocxParagraph[];
  /** Human-readable (German) names of content that RaabeWrite cannot keep yet. */
  unsupported: string[];
}

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

// ----- writing ----------------------------------------------------------------------------

function esc(s: string): string {
  return s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}

function runs(text: string): string {
  // Tabs become <w:tab/>; everything else is literal text with spaces preserved.
  return text.split("\t").map((part, i) =>
    (i > 0 ? "<w:r><w:tab/></w:r>" : "") + (part ? `<w:r><w:t xml:space="preserve">${esc(part)}</w:t></w:r>` : "")).join("");
}

const STYLE_ID: Record<ParagraphStyle, string | null> = { normal: null, h1: "berschrift1" };

function documentXml(paras: DocxParagraph[]): string {
  const body = paras.map((p) => {
    const id = STYLE_ID[p.style];
    return `<w:p>${id ? `<w:pPr><w:pStyle w:val="${id}"/></w:pPr>` : ""}${runs(p.text)}</w:p>`;
  }).join("");
  // A4, margins 2.5 / 2.5 / 2.0 / 2.5 cm, header and footer distance 1.25 cm: Word's German defaults.
  const sect = '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>' +
    '<w:pgMar w:top="1417" w:right="1417" w:bottom="1134" w:left="1417" w:header="708" w:footer="708" w:gutter="0"/>' +
    '<w:cols w:space="708"/><w:docGrid w:linePitch="360"/></w:sectPr>';
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n` +
    `<w:document xmlns:w="${W}" xmlns:r="${R}"><w:body>${body}${sect}</w:body></w:document>`;
}

const STYLES_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="${W}">
<w:docDefaults>
<w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:eastAsia="Calibri" w:hAnsi="Calibri" w:cs="Times New Roman"/><w:sz w:val="22"/><w:szCs w:val="22"/><w:lang w:val="de-DE" w:eastAsia="en-US" w:bidi="ar-SA"/></w:rPr></w:rPrDefault>
<w:pPrDefault><w:pPr><w:spacing w:after="160" w:line="259" w:lineRule="auto"/></w:pPr></w:pPrDefault>
</w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="Standard"><w:name w:val="Normal"/><w:qFormat/></w:style>
<w:style w:type="paragraph" w:styleId="berschrift1"><w:name w:val="heading 1"/><w:basedOn w:val="Standard"/><w:next w:val="Standard"/><w:link w:val="berschrift1Zchn"/><w:uiPriority w:val="9"/><w:qFormat/>
<w:pPr><w:keepNext/><w:keepLines/><w:spacing w:before="240" w:after="0"/><w:outlineLvl w:val="0"/></w:pPr>
<w:rPr><w:rFonts w:ascii="Calibri Light" w:eastAsia="Calibri Light" w:hAnsi="Calibri Light" w:cs="Times New Roman"/><w:color w:val="2F5496"/><w:sz w:val="32"/><w:szCs w:val="32"/></w:rPr></w:style>
<w:style w:type="character" w:default="1" w:styleId="Absatz-Standardschriftart"><w:name w:val="Default Paragraph Font"/><w:uiPriority w:val="1"/><w:semiHidden/><w:unhideWhenUsed/></w:style>
<w:style w:type="character" w:customStyle="1" w:styleId="berschrift1Zchn"><w:name w:val="Überschrift 1 Zchn"/><w:basedOn w:val="Absatz-Standardschriftart"/><w:link w:val="berschrift1"/><w:uiPriority w:val="9"/>
<w:rPr><w:rFonts w:ascii="Calibri Light" w:eastAsia="Calibri Light" w:hAnsi="Calibri Light" w:cs="Times New Roman"/><w:color w:val="2F5496"/><w:sz w:val="32"/><w:szCs w:val="32"/></w:rPr></w:style>
</w:styles>`;

const SETTINGS_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="${W}"><w:defaultTabStop w:val="708"/><w:hyphenationZone w:val="425"/><w:characterSpacingControl w:val="doNotCompress"/>
<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>`;

const CONTENT_TYPES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>`;

const ROOT_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>`;

const DOCUMENT_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>
</Relationships>`;

function coreXml(author: string, now: Date): string {
  const t = now.toISOString().replace(/\.\d+Z$/, "Z");
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
<dc:creator>${esc(author)}</dc:creator><cp:lastModifiedBy>${esc(author)}</cp:lastModifiedBy>
<dcterms:created xsi:type="dcterms:W3CDTF">${t}</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">${t}</dcterms:modified>
</cp:coreProperties>`;
}

const APP_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>RaabeWrite</Application></Properties>`;

export function writeDocx(paragraphs: DocxParagraph[], options: { author?: string; now?: Date } = {}): Uint8Array {
  const paras = paragraphs.length ? paragraphs : [{ style: "normal" as const, text: "" }];
  return zipSync({
    "[Content_Types].xml": strToU8(CONTENT_TYPES),
    "_rels/.rels": strToU8(ROOT_RELS),
    "docProps/core.xml": strToU8(coreXml(options.author ?? "", options.now ?? new Date())),
    "docProps/app.xml": strToU8(APP_XML),
    "word/document.xml": strToU8(documentXml(paras)),
    "word/styles.xml": strToU8(STYLES_XML),
    "word/settings.xml": strToU8(SETTINGS_XML),
    "word/_rels/document.xml.rels": strToU8(DOCUMENT_RELS),
  }, { level: 6 });
}

// ----- reading ----------------------------------------------------------------------------

/** Elements inside a run that we understand completely. */
const RUN_OK = new Set(["rPr", "t", "tab", "lastRenderedPageBreak", "softHyphen"]);
/** Run properties that do not change how text looks. */
const RPR_OK = new Set(["lang", "noProof", "rFonts", "kern", "szCs", "bCs", "iCs"]);
/** Paragraph properties we keep or that do not matter for the supported subset. */
const PPR_OK = new Set(["pStyle", "rPr", "spacing", "jc", "keepNext", "keepLines", "widowControl"]);

const REASONS: Record<string, string> = {
  tbl: "Tabellen",
  drawing: "Bilder und Formen",
  pict: "Bilder und Formen",
  object: "eingebettete Objekte",
  numPr: "Aufzählungen und Nummerierungen",
  rPr: "Zeichenformatierung (z. B. fett, kursiv, Farbe)",
  fldChar: "Felder (z. B. Seitenzahlen, Inhaltsverzeichnis)",
  fldSimple: "Felder (z. B. Seitenzahlen, Inhaltsverzeichnis)",
  instrText: "Felder (z. B. Seitenzahlen, Inhaltsverzeichnis)",
  footnoteReference: "Fußnoten",
  endnoteReference: "Endnoten",
  commentRangeStart: "Kommentare",
  commentReference: "Kommentare",
  ins: "nachverfolgte Änderungen",
  del: "nachverfolgte Änderungen",
  hyperlink: "Links",
  sdt: "Inhaltssteuerelemente",
  br: "Zeilen- und Seitenumbrüche",
  sym: "Sonderzeichen",
  header: "Kopf- und Fußzeilen",
  style: "weitere Formatvorlagen",
  pPr: "Absatzformatierung (z. B. Einzüge, Rahmen)",
  sect: "abweichende Seiteneinrichtung oder mehrere Abschnitte",
  other: "weitere Inhalte",
};

function child(el: Element, name: string): Element | null {
  for (const c of Array.from(el.children)) if (c.namespaceURI === W && c.localName === name) return c;
  return null;
}

function styleNames(stylesXml: string | null): Map<string, string> {
  const map = new Map<string, string>();
  if (!stylesXml) return map;
  const doc = new DOMParser().parseFromString(stylesXml, "application/xml");
  for (const s of Array.from(doc.getElementsByTagNameNS(W, "style"))) {
    const id = s.getAttributeNS(W, "styleId") ?? "";
    const name = child(s, "name")?.getAttributeNS(W, "val") ?? id;
    map.set(id, name.toLowerCase());
  }
  return map;
}

export function readDocx(bytes: Uint8Array): ReadResult {
  const files = unzipSync(bytes);
  const text = (path: string) => (files[path] ? strFromU8(files[path]) : null);
  const main = text("word/document.xml");
  if (!main) throw new Error("Die Datei ist kein Word-Dokument (word/document.xml fehlt).");
  const names = styleNames(text("word/styles.xml"));
  const problems = new Set<string>();
  const note = (key: string) => problems.add(REASONS[key] ?? REASONS.other);

  const rels = text("word/_rels/document.xml.rels") ?? "";
  if (/relationships\/(header|footer)"/.test(rels)) note("header");

  const doc = new DOMParser().parseFromString(main, "application/xml");
  const body = doc.getElementsByTagNameNS(W, "body")[0];
  const paragraphs: DocxParagraph[] = [];
  if (!body) return { paragraphs, unsupported: [] };

  for (const el of Array.from(body.children)) {
    if (el.namespaceURI !== W) { note("other"); continue; }
    if (el.localName === "sectPr") continue;
    if (el.localName !== "p") { note(el.localName); continue; }
    let style: ParagraphStyle = "normal";
    let out = "";
    for (const part of Array.from(el.children)) {
      const name = part.localName;
      if (name === "pPr") {
        for (const prop of Array.from(part.children)) {
          if (prop.localName === "pStyle") {
            const styleName = names.get(prop.getAttributeNS(W, "val") ?? "") ?? "";
            if (styleName === "heading 1") style = "h1";
            else if (styleName !== "normal") note("style");
          } else if (prop.localName === "numPr") note("numPr");
          else if (prop.localName === "sectPr") note("sect");
          else if (!PPR_OK.has(prop.localName)) note("pPr");
        }
      } else if (name === "r") {
        for (const item of Array.from(part.children)) {
          if (item.localName === "t") out += item.textContent ?? "";
          else if (item.localName === "tab") out += "\t";
          else if (item.localName === "rPr") {
            for (const prop of Array.from(item.children)) if (!RPR_OK.has(prop.localName)) note("rPr");
          } else if (!RUN_OK.has(item.localName)) note(item.localName);
        }
      } else if (name === "bookmarkStart" || name === "bookmarkEnd" || name === "proofErr") {
        // Harmless markers.
      } else {
        note(name);
      }
    }
    paragraphs.push({ style, text: out });
  }
  return { paragraphs, unsupported: [...problems] };
}
