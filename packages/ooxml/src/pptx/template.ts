// A new presentation as PowerPoint 2013-2022 makes it (D-070): 16:9, the Office design with
// Calibri Light / Calibri and its colours, one master with PowerPoint's eleven layouts and one
// Titelfolie. Placeholder positions and the text styles are Office's (recorded from PowerPoint,
// results/2026-11-06-powerpoint-layouts.json); the names are those of German PowerPoint.

import { strToU8 } from "fflate";
import { themeXml } from "../write";
import { OFFICE_2013_THEME } from "../xlsx/colors";
import { NS_A, NS_P, NS_R, XML_DECL, esc } from "./dom";
import { CT_PPTX, Pptx } from "./package";

const NSDECL = `xmlns:a="${NS_A}" xmlns:r="${NS_R}" xmlns:p="${NS_P}"`;
const T = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const RELS = "http://schemas.openxmlformats.org/package/2006/relationships";

export interface DesignSpec {
  /** Name of the design (theme), shown in Entwurf > Designs. */
  name: string;
  colors: Readonly<Record<string, string>>;
  colorName?: string;
  major: string;
  minor: string;
  fontName?: string;
  /** Extra shapes on the master (our own designs: a band, a line), as p:sp XML. */
  masterDeco?: string;
  /** Background of the master (p:bg inner XML); bg1 when missing. */
  background?: string;
  /** Colour of titles (scheme name or hex), tx1 when missing. */
  titleColor?: string;
  /** Light text on the dark colours (colour map bg1 = dk1, tx1 = lt1). */
  dark?: boolean;
}

export const OFFICE_DESIGN: DesignSpec = { name: "Office", colors: OFFICE_2013_THEME, colorName: "Office", major: "Calibri Light", minor: "Calibri", fontName: "Office" };

type Box = [number, number, number, number];

interface Ph {
  name: string;
  type?: string;
  idx?: number;
  orient?: "vert";
  size?: "half" | "quarter";
  box?: Box;
  body?: string;
  lst?: string;
  /** Prompt text in the layout (PowerPoint's own placeholder texts). */
  text?: string[];
}

const xfrm = (b: Box) => `<a:xfrm><a:off x="${b[0]}" y="${b[1]}"/><a:ext cx="${b[2]}" cy="${b[3]}"/></a:xfrm>`;

const MASTER_BOXES = {
  title: [838200, 365125, 10515600, 1325563] as Box,
  body: [838200, 1825625, 10515600, 4351338] as Box,
  dt: [838200, 6356350, 2743200, 365125] as Box,
  ftr: [4038600, 6356350, 4114800, 365125] as Box,
  sldNum: [8610600, 6356350, 2743200, 365125] as Box,
};

const LEVEL_TEXT = ["Mastertextformat bearbeiten", "Zweite Ebene", "Dritte Ebene", "Vierte Ebene", "Fünfte Ebene"];

function phXml(id: number, p: Ph, master: boolean): string {
  const attrs = [p.type ? `type="${p.type}"` : "", p.orient ? `orient="${p.orient}"` : "", p.size ? `sz="${p.size}"` : "", p.idx !== undefined ? `idx="${p.idx}"` : ""].filter(Boolean).join(" ");
  const paras = (p.text ?? [""]).map((t, i) => {
    const lvl = p.text && p.text.length > 1 && i > 0 ? ` lvl="${i}"` : "";
    return `<a:p>${lvl ? `<a:pPr${lvl}/>` : ""}${t ? `<a:r><a:rPr lang="de-DE"/><a:t>${esc(t)}</a:t></a:r>` : '<a:endParaRPr lang="de-DE"/>'}</a:p>`;
  }).join("");
  const field = p.type === "dt" ? `<a:p><a:fld id="{B61BEF0D-F0BB-DE4B-95CE-6DB70DBA9567}" type="datetimeFigureOut"><a:rPr lang="de-DE"/><a:t>06.11.2026</a:t></a:fld><a:endParaRPr lang="de-DE"/></a:p>`
    : p.type === "sldNum" ? `<a:p><a:fld id="{D57F1E4F-1CFF-5643-939E-217C01CDF565}" type="slidenum"><a:rPr lang="de-DE"/><a:t>‹Nr.›</a:t></a:fld><a:endParaRPr lang="de-DE"/></a:p>` : null;
  const txBody = `<p:txBody>${p.body ?? "<a:bodyPr/>"}<a:lstStyle>${p.lst ?? ""}</a:lstStyle>${field ?? paras}</p:txBody>`;
  return `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${esc(p.name)}"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph ${attrs}/></p:nvPr></p:nvSpPr>` +
    `<p:spPr>${p.box ? xfrm(p.box) : ""}${master ? `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>` : ""}</p:spPr>${txBody}</p:sp>`;
}

function spTree(shapes: string): string {
  return `<p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>${shapes}</p:spTree>`;
}

const FOOTERS: Ph[] = [
  { name: "Datumsplatzhalter 3", type: "dt", idx: 10, size: "half" },
  { name: "Fußzeilenplatzhalter 4", type: "ftr", idx: 11, size: "quarter" },
  { name: "Foliennummernplatzhalter 5", type: "sldNum", idx: 12, size: "quarter" },
];

const lvls = (f: (i: number) => string) => Array.from({ length: 9 }, (_, i) => `<a:lvl${i + 1}pPr${f(i)}</a:lvl${i + 1}pPr>`).join("");
const noBullets = (sizes: number[], extra = "") => lvls((i) => ` marL="${i * 457200}" indent="0"${extra}><a:buNone/><a:defRPr sz="${sizes[Math.min(i, sizes.length - 1)]}"/>`);
const tint75 = '<a:solidFill><a:schemeClr val="tx1"><a:lumMod val="75000"/><a:lumOff val="25000"/></a:schemeClr></a:solidFill>';
const halfSizes = [2400, 2000, 1800, 1600];

/** The eleven layouts of Office's master, in PowerPoint's order. */
const LAYOUTS: { name: string; type: string; ph: Ph[]; preserve?: boolean }[] = [
  { name: "Titelfolie", type: "title", ph: [
    { name: "Titel 1", type: "ctrTitle", box: [1524000, 1122363, 9144000, 2387600], body: '<a:bodyPr anchor="b"/>', lst: '<a:lvl1pPr algn="ctr"><a:defRPr sz="6000"/></a:lvl1pPr>', text: ["Mastertitelformat bearbeiten"] },
    { name: "Untertitel 2", type: "subTitle", idx: 1, box: [1524000, 3602038, 9144000, 1655762], lst: lvls((i) => ` marL="${i * 457200}" indent="0" algn="ctr"><a:buNone/><a:defRPr sz="${[2400, 2000, 1800, 1600][Math.min(i, 3)]}"/>`), text: ["Master-Untertitelformat bearbeiten"] },
    ...FOOTERS] },
  { name: "Titel und Inhalt", type: "obj", preserve: true, ph: [
    { name: "Titel 1", type: "title", text: ["Mastertitelformat bearbeiten"] },
    { name: "Inhaltsplatzhalter 2", idx: 1, text: LEVEL_TEXT }, ...FOOTERS] },
  // German PowerPoint names this layout with a line break after the hyphen ("Abschnitts-&#xA;überschrift").
  { name: `Abschnitts-${String.fromCharCode(10)}überschrift`, type: "secHead", preserve: true, ph: [
    { name: "Titel 1", type: "title", box: [831850, 1709738, 10515600, 2852737], body: '<a:bodyPr anchor="b"/>', lst: '<a:lvl1pPr><a:defRPr sz="6000"/></a:lvl1pPr>', text: ["Mastertitelformat bearbeiten"] },
    { name: "Textplatzhalter 2", type: "body", idx: 1, box: [831850, 4589463, 10515600, 1500187], lst: lvls((i) => ` marL="${i * 457200}" indent="0"><a:buNone/><a:defRPr sz="${[2400, 2000, 1800, 1600][Math.min(i, 3)]}">${i === 0 ? tint75 : '<a:solidFill><a:schemeClr val="tx1"><a:tint val="75000"/></a:schemeClr></a:solidFill>'}</a:defRPr>`), text: ["Mastertextformat bearbeiten"] },
    ...FOOTERS] },
  { name: "Zwei Inhalte", type: "twoObj", preserve: true, ph: [
    { name: "Titel 1", type: "title", text: ["Mastertitelformat bearbeiten"] },
    { name: "Inhaltsplatzhalter 2", size: "half", idx: 1, box: [838200, 1825625, 5181600, 4351338], text: LEVEL_TEXT },
    { name: "Inhaltsplatzhalter 3", size: "half", idx: 2, box: [6172200, 1825625, 5181600, 4351338], text: LEVEL_TEXT }, ...FOOTERS] },
  { name: "Vergleich", type: "twoTxTwoObj", preserve: true, ph: [
    { name: "Titel 1", type: "title", box: [839788, 365125, 10515600, 1325563], text: ["Mastertitelformat bearbeiten"] },
    { name: "Textplatzhalter 2", type: "body", idx: 1, box: [839788, 1681163, 5157787, 823912], body: '<a:bodyPr anchor="b"/>', lst: noBullets(halfSizes, "").replace('<a:defRPr sz="2400"/>', '<a:defRPr sz="2400" b="1"/>').replace(/<a:defRPr sz="(\d+)"\/>/g, '<a:defRPr sz="$1" b="1"/>'), text: ["Mastertextformat bearbeiten"] },
    { name: "Inhaltsplatzhalter 3", size: "half", idx: 2, box: [839788, 2505075, 5157787, 3684588], text: LEVEL_TEXT },
    { name: "Textplatzhalter 4", type: "body", size: "quarter", idx: 3, box: [6172200, 1681163, 5183188, 823912], body: '<a:bodyPr anchor="b"/>', lst: noBullets(halfSizes, "").replace(/<a:defRPr sz="(\d+)"\/>/g, '<a:defRPr sz="$1" b="1"/>'), text: ["Mastertextformat bearbeiten"] },
    { name: "Inhaltsplatzhalter 5", size: "quarter", idx: 4, box: [6172200, 2505075, 5183188, 3684588], text: LEVEL_TEXT }, ...FOOTERS.map((f, i) => ({ ...f, name: f.name.replace(/\d+$/, String(6 + i)) }))] },
  { name: "Nur Titel", type: "titleOnly", preserve: true, ph: [{ name: "Titel 1", type: "title", text: ["Mastertitelformat bearbeiten"] }, ...FOOTERS.map((f, i) => ({ ...f, name: f.name.replace(/\d+$/, String(2 + i)) }))] },
  { name: "Leer", type: "blank", preserve: true, ph: FOOTERS.map((f, i) => ({ ...f, name: f.name.replace(/\d+$/, String(1 + i)) })) },
  { name: "Inhalt mit Überschrift", type: "objTx", preserve: true, ph: [
    { name: "Titel 1", type: "title", box: [839788, 457200, 3932237, 1600200], body: '<a:bodyPr anchor="b"/>', lst: '<a:lvl1pPr><a:defRPr sz="3200"/></a:lvl1pPr>', text: ["Mastertitelformat bearbeiten"] },
    { name: "Inhaltsplatzhalter 2", idx: 1, box: [5183188, 987425, 6172200, 4873625], lst: lvls((i) => `><a:defRPr sz="${[3200, 2800, 2400, 2000][Math.min(i, 3)]}"/>`), text: LEVEL_TEXT },
    { name: "Textplatzhalter 3", type: "body", size: "half", idx: 2, box: [839788, 2057400, 3932237, 3811588], lst: noBullets([1600, 1400, 1200, 1000]), text: ["Mastertextformat bearbeiten"] },
    ...FOOTERS.map((f, i) => ({ ...f, name: f.name.replace(/\d+$/, String(4 + i)) }))] },
  { name: "Bild mit Überschrift", type: "picTx", preserve: true, ph: [
    { name: "Titel 1", type: "title", box: [839788, 457200, 3932237, 1600200], body: '<a:bodyPr anchor="b"/>', lst: '<a:lvl1pPr><a:defRPr sz="3200"/></a:lvl1pPr>', text: ["Mastertitelformat bearbeiten"] },
    { name: "Bildplatzhalter 2", type: "pic", idx: 1, box: [5183188, 987425, 6172200, 4873625], lst: lvls((i) => ` marL="${i * 457200}" indent="0"><a:buNone/><a:defRPr sz="${[3200, 2800, 2400, 2000][Math.min(i, 3)]}"/>`), text: [""] },
    { name: "Textplatzhalter 3", type: "body", size: "half", idx: 2, box: [839788, 2057400, 3932237, 3811588], lst: noBullets([1600, 1400, 1200, 1000]), text: ["Mastertextformat bearbeiten"] },
    ...FOOTERS.map((f, i) => ({ ...f, name: f.name.replace(/\d+$/, String(4 + i)) }))] },
  { name: "Titel und vertikaler Text", type: "vertTx", preserve: true, ph: [
    { name: "Titel 1", type: "title", text: ["Mastertitelformat bearbeiten"] },
    { name: "Vertikaler Textplatzhalter 2", type: "body", orient: "vert", idx: 1, body: '<a:bodyPr vert="eaVert"/>', text: LEVEL_TEXT }, ...FOOTERS] },
  { name: "Vertikaler Titel und Text", type: "vertTitleAndTx", preserve: true, ph: [
    { name: "Vertikaler Titel 1", type: "title", orient: "vert", box: [8724900, 365125, 2628900, 5811838], body: '<a:bodyPr vert="eaVert"/>', text: ["Mastertitelformat bearbeiten"] },
    { name: "Vertikaler Textplatzhalter 2", type: "body", orient: "vert", idx: 1, box: [838200, 365125, 7734300, 5811838], body: '<a:bodyPr vert="eaVert"/>', text: LEVEL_TEXT }, ...FOOTERS] },
];

const BODY_LEVELS = [228600, 685800, 1143000, 1600200, 2057400, 2514600, 2971800, 3429000, 3886200];
const BODY_SIZES = [2800, 2400, 2000, 1800, 1800, 1800, 1800, 1800, 1800];
const rpr = (sz: number, font: "mj" | "mn", color = '<a:solidFill><a:schemeClr val="tx1"/></a:solidFill>') =>
  `<a:defRPr sz="${sz}" kern="1200">${color}<a:latin typeface="+${font}-lt"/><a:ea typeface="+${font}-ea"/><a:cs typeface="+${font}-cs"/></a:defRPr>`;

/** The levels of presentation.xml's defaultTextStyle and the master's otherStyle. */
export function otherLevels(): string {
  return lvls((i) => ` marL="${i * 457200}" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">${rpr(1800, "mn")}`);
}

function titleColor(d: DesignSpec): string {
  if (!d.titleColor) return '<a:solidFill><a:schemeClr val="tx1"/></a:solidFill>';
  return /^[0-9A-F]{6}$/i.test(d.titleColor) ? `<a:solidFill><a:srgbClr val="${d.titleColor}"/></a:solidFill>` : `<a:solidFill><a:schemeClr val="${d.titleColor}"/></a:solidFill>`;
}

export function masterXml(d: DesignSpec): string {
  const footerLst = (algn: string) => `<a:lvl1pPr algn="${algn}"><a:defRPr sz="1200"><a:solidFill><a:schemeClr val="tx1"><a:tint val="75000"/></a:schemeClr></a:solidFill></a:defRPr></a:lvl1pPr>`;
  const shapes = [
    phXml(2, { name: "Titelplatzhalter 1", type: "title", box: MASTER_BOXES.title, body: '<a:bodyPr vert="horz" lIns="91440" tIns="45720" rIns="91440" bIns="45720" rtlCol="0" anchor="ctr"><a:normAutofit/></a:bodyPr>', text: ["Mastertitelformat bearbeiten"] }, true),
    phXml(3, { name: "Textplatzhalter 2", type: "body", idx: 1, box: MASTER_BOXES.body, body: '<a:bodyPr vert="horz" lIns="91440" tIns="45720" rIns="91440" bIns="45720" rtlCol="0"><a:normAutofit/></a:bodyPr>', text: LEVEL_TEXT }, true),
    phXml(4, { name: "Datumsplatzhalter 3", type: "dt", idx: 2, size: "half", box: MASTER_BOXES.dt, body: '<a:bodyPr vert="horz" lIns="91440" tIns="45720" rIns="91440" bIns="45720" rtlCol="0" anchor="ctr"/>', lst: footerLst("l") }, true),
    phXml(5, { name: "Fußzeilenplatzhalter 4", type: "ftr", idx: 3, size: "quarter", box: MASTER_BOXES.ftr, body: '<a:bodyPr vert="horz" lIns="91440" tIns="45720" rIns="91440" bIns="45720" rtlCol="0" anchor="ctr"/>', lst: footerLst("ctr"), text: [""] }, true),
    phXml(6, { name: "Foliennummernplatzhalter 5", type: "sldNum", idx: 4, size: "quarter", box: MASTER_BOXES.sldNum, body: '<a:bodyPr vert="horz" lIns="91440" tIns="45720" rIns="91440" bIns="45720" rtlCol="0" anchor="ctr"/>', lst: footerLst("r") }, true),
  ].join("");
  const titleStyle = `<p:titleStyle><a:lvl1pPr algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:lnSpc><a:spcPct val="90000"/></a:lnSpc><a:spcBef><a:spcPct val="0"/></a:spcBef><a:buNone/>${rpr(4400, "mj", titleColor(d))}</a:lvl1pPr></p:titleStyle>`;
  const bodyStyle = "<p:bodyStyle>" + BODY_LEVELS.map((marL, i) =>
    `<a:lvl${i + 1}pPr marL="${marL}" indent="-228600" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:lnSpc><a:spcPct val="90000"/></a:lnSpc><a:spcBef><a:spcPts val="${i === 0 ? 1000 : 500}"/></a:spcBef><a:buFont typeface="Arial" panose="020B0604020202020204" pitchFamily="34" charset="0"/><a:buChar char="•"/>${rpr(BODY_SIZES[i], "mn")}</a:lvl${i + 1}pPr>`).join("") + "</p:bodyStyle>";
  const otherStyle = `<p:otherStyle><a:defPPr><a:defRPr lang="de-DE"/></a:defPPr>${otherLevels()}</p:otherStyle>`;
  const ids = LAYOUTS.map((_, i) => `<p:sldLayoutId id="${2147483649 + i}" r:id="rId${i + 1}"/>`).join("");
  return `${XML_DECL}<p:sldMaster ${NSDECL}><p:cSld><p:bg>${d.background ?? '<p:bgRef idx="1001"><a:schemeClr val="bg1"/></p:bgRef>'}</p:bg>${spTree((d.masterDeco ?? "") + shapes)}</p:cSld>` +
    (d.dark ? `<p:clrMap bg1="dk1" tx1="lt1" bg2="dk2" tx2="lt2" accent1` : `<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1`) + `="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>` +
    `<p:sldLayoutIdLst>${ids}</p:sldLayoutIdLst><p:txStyles>${titleStyle}${bodyStyle}${otherStyle}</p:txStyles></p:sldMaster>`;
}

function layoutXml(l: (typeof LAYOUTS)[number]): string {
  let id = 2;
  // PowerPoint numbers the footer placeholders on after the layout's other placeholders.
  const own = l.ph.filter((p) => !FOOTERS.some((f) => f.type === p.type)).length;
  const shapes = l.ph.map((p, i) => phXml(id++, FOOTERS.some((f) => f.type === p.type) ? { ...p, name: p.name.replace(/\d+$/, String(i + 1 > own ? i + 1 : own + 1)) } : p, false)).join("");
  return `${XML_DECL}<p:sldLayout ${NSDECL} type="${l.type}"${l.preserve ? ' preserve="1"' : ""}><p:cSld name="${esc(l.name).split(String.fromCharCode(10)).join("&#xA;")}">${spTree(shapes)}</p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>`;
}

export const LAYOUT_NAMES = LAYOUTS.map((l) => l.name);

/** A design's theme part. */
export function designTheme(d: DesignSpec): string {
  return themeXml({ majorFont: d.major, minorFont: d.minor, colors: d.colors as Record<string, string>, name: d.name, colorName: d.colorName ?? d.name, fontName: d.fontName ?? d.name });
}

/** The parts of a new presentation (one Titelfolie), as a package. */
export function newPptxFiles(design: DesignSpec = OFFICE_DESIGN, now = new Date()): Record<string, Uint8Array> {
  const t = now.toISOString().replace(/\.\d+Z$/, "Z");
  const layoutCt = LAYOUTS.map((_, i) => `<Override PartName="/ppt/slideLayouts/slideLayout${i + 1}.xml" ContentType="${CT_PPTX.slideLayout}"/>`).join("");
  const files: Record<string, string> = {
    "[Content_Types].xml": `${XML_DECL}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>` +
      `<Default Extension="png" ContentType="image/png"/><Default Extension="jpeg" ContentType="image/jpeg"/>` +
      `<Override PartName="/ppt/presentation.xml" ContentType="${CT_PPTX.presentation}"/><Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="${CT_PPTX.slideMaster}"/>` +
      `<Override PartName="/ppt/slides/slide1.xml" ContentType="${CT_PPTX.slide}"/>` +
      `<Override PartName="/ppt/presProps.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presProps+xml"/><Override PartName="/ppt/viewProps.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.viewProps+xml"/>` +
      `<Override PartName="/ppt/theme/theme1.xml" ContentType="${CT_PPTX.theme}"/><Override PartName="/ppt/tableStyles.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.tableStyles+xml"/>` +
      layoutCt +
      `<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/><Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/></Types>`,
    "_rels/.rels": `${XML_DECL}<Relationships xmlns="${RELS}"><Relationship Id="rId3" Type="${T}/extended-properties" Target="docProps/app.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/><Relationship Id="rId1" Type="${T}/officeDocument" Target="ppt/presentation.xml"/></Relationships>`,
    "docProps/core.xml": `${XML_DECL}<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:title>PowerPoint-Präsentation</dc:title><dc:creator></dc:creator><cp:lastModifiedBy></cp:lastModifiedBy><cp:revision>1</cp:revision><dcterms:created xsi:type="dcterms:W3CDTF">${t}</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">${t}</dcterms:modified></cp:coreProperties>`,
    "docProps/app.xml": `${XML_DECL}<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><Application>RaabePresent</Application><PresentationFormat>Breitbild</PresentationFormat><Slides>1</Slides></Properties>`,
    "ppt/presentation.xml": `${XML_DECL}<p:presentation ${NSDECL} saveSubsetFonts="1"><p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst><p:sldIdLst><p:sldId id="256" r:id="rId2"/></p:sldIdLst>` +
      `<p:sldSz cx="12192000" cy="6858000"/><p:notesSz cx="6858000" cy="9144000"/><p:defaultTextStyle><a:defPPr><a:defRPr lang="de-DE"/></a:defPPr>${otherLevels()}</p:defaultTextStyle></p:presentation>`,
    "ppt/_rels/presentation.xml.rels": `${XML_DECL}<Relationships xmlns="${RELS}"><Relationship Id="rId3" Type="${T}/presProps" Target="presProps.xml"/><Relationship Id="rId2" Type="${T}/slide" Target="slides/slide1.xml"/><Relationship Id="rId1" Type="${T}/slideMaster" Target="slideMasters/slideMaster1.xml"/><Relationship Id="rId6" Type="${T}/tableStyles" Target="tableStyles.xml"/><Relationship Id="rId5" Type="${T}/theme" Target="theme/theme1.xml"/><Relationship Id="rId4" Type="${T}/viewProps" Target="viewProps.xml"/></Relationships>`,
    "ppt/presProps.xml": `${XML_DECL}<p:presentationPr ${NSDECL}/>`,
    "ppt/viewProps.xml": `${XML_DECL}<p:viewPr ${NSDECL}><p:normalViewPr horzBarState="maximized"><p:restoredLeft sz="15987" autoAdjust="0"/><p:restoredTop sz="94660"/></p:normalViewPr><p:slideViewPr><p:cSldViewPr snapToGrid="0"><p:cViewPr varScale="1"><p:scale><a:sx n="114" d="100"/><a:sy n="114" d="100"/></p:scale><p:origin x="414" y="102"/></p:cViewPr><p:guideLst/></p:cSldViewPr></p:slideViewPr><p:gridSpacing cx="72008" cy="72008"/></p:viewPr>`,
    "ppt/tableStyles.xml": `${XML_DECL}<a:tblStyleLst xmlns:a="${NS_A}" def="{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}"/>`,
    "ppt/theme/theme1.xml": designTheme(design),
    "ppt/slideMasters/slideMaster1.xml": masterXml(design),
    "ppt/slideMasters/_rels/slideMaster1.xml.rels": `${XML_DECL}<Relationships xmlns="${RELS}">${LAYOUTS.map((_, i) => `<Relationship Id="rId${i + 1}" Type="${T}/slideLayout" Target="../slideLayouts/slideLayout${i + 1}.xml"/>`).join("")}<Relationship Id="rId${LAYOUTS.length + 1}" Type="${T}/theme" Target="../theme/theme1.xml"/></Relationships>`,
    "ppt/slides/slide1.xml": `${XML_DECL}<p:sld ${NSDECL}><p:cSld>${spTree(
      `<p:sp><p:nvSpPr><p:cNvPr id="2" name="Titel 1"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="ctrTitle"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="de-DE"/></a:p></p:txBody></p:sp>` +
      `<p:sp><p:nvSpPr><p:cNvPr id="3" name="Untertitel 2"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="subTitle" idx="1"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="de-DE"/></a:p></p:txBody></p:sp>`)}</p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>`,
    "ppt/slides/_rels/slide1.xml.rels": `${XML_DECL}<Relationships xmlns="${RELS}"><Relationship Id="rId1" Type="${T}/slideLayout" Target="../slideLayouts/slideLayout1.xml"/></Relationships>`,
  };
  LAYOUTS.forEach((l, i) => {
    files[`ppt/slideLayouts/slideLayout${i + 1}.xml`] = layoutXml(l);
    files[`ppt/slideLayouts/_rels/slideLayout${i + 1}.xml.rels`] = `${XML_DECL}<Relationships xmlns="${RELS}"><Relationship Id="rId1" Type="${T}/slideMaster" Target="../slideMasters/slideMaster1.xml"/></Relationships>`;
  });
  const out: Record<string, Uint8Array> = {};
  for (const [k, v] of Object.entries(files)) out[k] = strToU8(v);
  return out;
}

export function newPptx(design: DesignSpec = OFFICE_DESIGN): Pptx {
  return new Pptx(newPptxFiles(design));
}
