// A new workbook as German Excel 2013-2022 makes it: one sheet "Tabelle1", Calibri 11 in the
// Office 2013-2022 design (D-011, D-061), default row height 15 pt, standard column width
// from baseColWidth 10 (10.71 characters, as German Excel writes it).

import { XML_HEAD } from "./xml";
import { strToU8 } from "fflate";
import type { Workbook } from "@raabe/sheet-engine";
import { themeXml } from "../write";
import { OFFICE_2013_THEME } from "./colors";
import { readPackage } from "./read";
import { NEW_STYLES_XML } from "./styles";

const HEAD = XML_HEAD;
const NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const RELS = "http://schemas.openxmlformats.org/package/2006/relationships";
const T = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

/** The files of a new workbook's package; with `sheet` one empty sheet "Tabelle1". */
export function newPackage(sheet: boolean, now = new Date()): Record<string, Uint8Array> {
  const t = now.toISOString().replace(/\.\d+Z$/, "Z");
  const files: Record<string, string> = {
    "[Content_Types].xml": `${HEAD}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>` +
      `<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>` +
      (sheet ? `<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>` : "") +
      `<Override PartName="/xl/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>` +
      `<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/><Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/></Types>`,
    "_rels/.rels": `${HEAD}<Relationships xmlns="${RELS}"><Relationship Id="rId3" Type="${T}/extended-properties" Target="docProps/app.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/><Relationship Id="rId1" Type="${T}/officeDocument" Target="xl/workbook.xml"/></Relationships>`,
    "docProps/core.xml": `${HEAD}<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:creator></dc:creator><cp:lastModifiedBy></cp:lastModifiedBy><dcterms:created xsi:type="dcterms:W3CDTF">${t}</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">${t}</dcterms:modified></cp:coreProperties>`,
    "docProps/app.xml": `${HEAD}<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><Application>RaabeCalc</Application></Properties>`,
    "xl/workbook.xml": `${HEAD}<workbook xmlns="${NS}" xmlns:r="${NS_R}"><workbookPr/><bookViews><workbookView xWindow="-110" yWindow="-110" windowWidth="19420" windowHeight="10420"/></bookViews>` +
      `<sheets>${sheet ? '<sheet name="Tabelle1" sheetId="1" r:id="rId1"/>' : ""}</sheets><calcPr calcId="191029"/></workbook>`,
    "xl/_rels/workbook.xml.rels": `${HEAD}<Relationships xmlns="${RELS}"><Relationship Id="rId3" Type="${T}/styles" Target="styles.xml"/><Relationship Id="rId2" Type="${T}/theme" Target="theme/theme1.xml"/>` +
      (sheet ? `<Relationship Id="rId1" Type="${T}/worksheet" Target="worksheets/sheet1.xml"/>` : "") + "</Relationships>",
    "xl/styles.xml": NEW_STYLES_XML,
    "xl/theme/theme1.xml": themeXml({ majorFont: "Calibri Light", minorFont: "Calibri", colors: OFFICE_2013_THEME, name: "Office", colorName: "Office", fontName: "Office" }),
  };
  if (sheet) {
    files["xl/worksheets/sheet1.xml"] = `${HEAD}<worksheet xmlns="${NS}" xmlns:r="${NS_R}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac">` +
      `<dimension ref="A1"/><sheetViews><sheetView tabSelected="1" workbookViewId="0"/></sheetViews><sheetFormatPr baseColWidth="10" defaultRowHeight="15" x14ac:dyDescent="0.25"/><sheetData/>` +
      `<pageMargins left="0.7" right="0.7" top="0.78740157499999996" bottom="0.78740157499999996" header="0.3" footer="0.3"/></worksheet>`;
  }
  return Object.fromEntries(Object.entries(files).map(([k, v]) => [k, strToU8(v)]));
}

/** A new workbook (Neue leere Arbeitsmappe). */
export function newWorkbook(): Workbook {
  return readPackage(newPackage(true));
}
