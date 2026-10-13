// Writing the model back to .docx. With the original package, every part except
// word/document.xml (and added style definitions) is written back byte for byte, and all
// preserved XML goes back in its place.

import { strFromU8, strToU8, zipSync } from "fflate";
import { LINE_BREAK, OBJ, PAGE_BREAK, type DocState, type Para, type StyleDef } from "@raabe/text-engine";
import type { DocxPackage } from "./read";
import { writePPr, writeRPr, type WriteOptions } from "./props";
import { esc, R, W } from "./xml";

const NBHYPHEN = String.fromCharCode(0x2011);
const SOFTHYPHEN = String.fromCharCode(0xad);

function runsXml(p: Para, opts: WriteOptions): string {
  let out = "";
  const anchors = [...(p.anchors ?? [])].sort((a, b) => a.at - b.at);
  let ai = 0;
  let objIdx = 0;
  let pos = 0;
  const anchorsAt = (i: number) => {
    let s = "";
    while (ai < anchors.length && anchors[ai].at <= i) s += String(anchors[ai++].data);
    return s;
  };
  for (const run of p.runs) {
    const rPr = writeRPr(run.props, opts);
    let inner = "";
    let text = "";
    const flushText = () => { if (text) { inner += `<w:t xml:space="preserve">${esc(text)}</w:t>`; text = ""; } };
    const flushRun = () => { flushText(); if (inner) { out += `<w:r>${rPr}${inner}</w:r>`; inner = ""; } };
    for (let i = pos; i < pos + run.len; i++) {
      const a = anchorsAt(i);
      if (a) { flushRun(); out += a; }
      const ch = p.text[i];
      if (ch === OBJ) { flushRun(); out += String(p.objects?.[objIdx++]?.data ?? ""); continue; }
      if (ch === "\t") { flushText(); inner += "<w:tab/>"; }
      else if (ch === LINE_BREAK) { flushText(); inner += "<w:br/>"; }
      else if (ch === PAGE_BREAK) { flushText(); inner += '<w:br w:type="page"/>'; }
      else if (ch === NBHYPHEN) { flushText(); inner += "<w:noBreakHyphen/>"; }
      else if (ch === SOFTHYPHEN) { flushText(); inner += "<w:softHyphen/>"; }
      else text += ch;
    }
    flushRun();
    pos += run.len;
  }
  out += anchorsAt(Number.MAX_SAFE_INTEGER);
  return out;
}

function paraXml(p: Para, opts: WriteOptions): string {
  const leading = (p.leading ?? []).map(String).join("");
  if (p.block) return leading + String(p.block.data);
  return `${leading}<w:p>${writePPr(p.props, p.mark, opts)}${runsXml(p, opts)}</w:p>`;
}

function sectPrXml(state: DocState): string {
  const pg = state.settings.page;
  const t = (pt: number) => Math.round(pt * 20);
  return `<w:sectPr><w:pgSz w:w="${t(pg.width)}" w:h="${t(pg.height)}"${pg.width > pg.height ? ' w:orient="landscape"' : ""}/>` +
    `<w:pgMar w:top="${t(pg.marginTop)}" w:right="${t(pg.marginRight)}" w:bottom="${t(pg.marginBottom)}" w:left="${t(pg.marginLeft)}" w:header="708" w:footer="708" w:gutter="0"/>` +
    '<w:cols w:space="708"/><w:docGrid w:linePitch="360"/></w:sectPr>';
}

const DEFAULT_ROOT = `<w:document xmlns:w="${W}" xmlns:r="${R}">`;

function documentXml(state: DocState, pkg: DocxPackage | null, opts: WriteOptions): string {
  const body = state.paras.map((p) => paraXml(p, opts)).join("");
  const trailing = (pkg?.trailing ?? []).join("");
  const sect = pkg?.sectPr ?? sectPrXml(state);
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n${pkg?.rootTag ?? DEFAULT_ROOT}<w:body>${body}${trailing}${sect}</w:body></w:document>`;
}

// ----- styles -----------------------------------------------------------------------------

export function styleXml(s: StyleDef, opts: WriteOptions): string {
  if (s.data && typeof s.data === "string") return s.data;
  const parts = [`<w:name w:val="${esc(s.name)}"/>`];
  if (s.basedOn) parts.push(`<w:basedOn w:val="${esc(s.basedOn)}"/>`);
  if (s.next) parts.push(`<w:next w:val="${esc(s.next)}"/>`);
  if (s.link) parts.push(`<w:link w:val="${esc(s.link)}"/>`);
  if (s.uiPriority !== undefined) parts.push(`<w:uiPriority w:val="${s.uiPriority}"/>`);
  if (s.semiHidden) parts.push("<w:semiHidden/><w:unhideWhenUsed/>");
  if (s.qFormat) parts.push("<w:qFormat/>");
  if (s.type === "paragraph") parts.push(writePPr(s.pPr, {}, opts));
  parts.push(writeRPr(s.rPr, opts));
  return `<w:style w:type="${s.type}"${s.isDefault ? ' w:default="1"' : ""} w:styleId="${esc(s.id)}">${parts.join("")}</w:style>`;
}

function stylesXml(state: DocState, opts: WriteOptions): string {
  const st = state.styles;
  const defs = [...st.styles.values()].map((s) => styleXml(s, opts)).join("");
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:styles xmlns:w="${W}"><w:docDefaults>` +
    `<w:rPrDefault>${writeRPr(st.defaultsR, opts) || "<w:rPr/>"}</w:rPrDefault>` +
    `<w:pPrDefault>${writePPr(st.defaultsP, {}, opts) || "<w:pPr/>"}</w:pPrDefault></w:docDefaults>${defs}</w:styles>`;
}

/** The original styles.xml plus definitions for styles the document now uses but lacked. */
function mergedStyles(state: DocState, pkg: DocxPackage, opts: WriteOptions): string | null {
  if (!pkg.stylesXml) return null;
  const added = [...state.styles.styles.values()].filter((s) => !pkg.styleIds.has(s.id));
  if (!added.length) return pkg.stylesXml;
  const i = pkg.stylesXml.lastIndexOf("</w:styles>");
  if (i < 0) return pkg.stylesXml;
  return pkg.stylesXml.slice(0, i) + added.map((s) => styleXml(s, opts)).join("") + pkg.stylesXml.slice(i);
}

// ----- the package --------------------------------------------------------------------------

const SETTINGS_XML = (defaultTab: number) => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="${W}"><w:defaultTabStop w:val="${Math.round(defaultTab * 20)}"/><w:hyphenationZone w:val="425"/><w:characterSpacingControl w:val="doNotCompress"/>
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

/** Updates the modification date (and last author) in an existing core.xml. */
function touchCore(xml: string, author: string, now: Date): string {
  const t = now.toISOString().replace(/\.\d+Z$/, "Z");
  let out = xml.replace(/(<dcterms:modified[^>]*>)[^<]*(<\/dcterms:modified>)/, `$1${t}$2`);
  if (author) out = out.replace(/(<cp:lastModifiedBy>)[^<]*(<\/cp:lastModifiedBy>)/, `$1${esc(author)}$2`);
  return out;
}

const APP_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>RaabeWrite</Application></Properties>`;

export function writeDocx(state: DocState, pkg: DocxPackage | null, options: { author?: string; now?: Date } = {}): Uint8Array {
  const now = options.now ?? new Date();
  const author = options.author ?? "";
  if (pkg) {
    const opts: WriteOptions = { hasTheme: pkg.hasTheme, themeFonts: pkg.hasTheme ? undefined : { major: state.styles.theme.majorFont, minor: state.styles.theme.minorFont } };
    const files: Record<string, Uint8Array> = { ...pkg.files };
    files["word/document.xml"] = strToU8(documentXml(state, pkg, opts));
    const styles = mergedStyles(state, pkg, opts);
    if (styles && styles !== pkg.stylesXml) files["word/styles.xml"] = strToU8(styles);
    if (files["docProps/core.xml"]) files["docProps/core.xml"] = strToU8(touchCore(strFromU8(files["docProps/core.xml"]), author, now));
    return zipSync(files, { level: 6 });
  }
  // New documents carry no theme part, so fonts are written by name.
  const opts: WriteOptions = { hasTheme: false, themeFonts: { major: state.styles.theme.majorFont, minor: state.styles.theme.minorFont } };
  return zipSync({
    "[Content_Types].xml": strToU8(CONTENT_TYPES),
    "_rels/.rels": strToU8(ROOT_RELS),
    "docProps/core.xml": strToU8(coreXml(author, now)),
    "docProps/app.xml": strToU8(APP_XML),
    "word/document.xml": strToU8(documentXml(state, null, opts)),
    "word/styles.xml": strToU8(stylesXml(state, opts)),
    "word/settings.xml": strToU8(SETTINGS_XML(state.settings.defaultTab)),
    "word/_rels/document.xml.rels": strToU8(DOCUMENT_RELS),
  }, { level: 6 });
}
