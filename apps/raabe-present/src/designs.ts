// Designs offered in Entwurf > Designs and Datei > Neu (D-070): Office's design as PowerPoint
// 2013-2022 ships it (Calibri Light / Calibri and its colours), plus our own designs, among them a
// Raabeschule design in Raabe red (D-018). PowerPoint's own artwork is not copied: our designs are
// plain DrawingML (bands, lines) on the master, so files from them open in PowerPoint unchanged.

import { OFFICE_DESIGN, type DesignSpec } from "@raabe/ooxml";

const band = (id: number, name: string, x: number, y: number, w: number, h: number, color: string) =>
  `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${name}"/><p:cNvSpPr/><p:nvPr userDrawn="1"/></p:nvSpPr><p:spPr><a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${w}" cy="${h}"/></a:xfrm>` +
  `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill>${/^[0-9A-F]{6}$/.test(color) ? `<a:srgbClr val="${color}"/>` : `<a:schemeClr val="${color}"/>`}</a:solidFill><a:ln><a:noFill/></a:ln></p:spPr>` +
  `<p:txBody><a:bodyPr rtlCol="0" anchor="ctr"/><a:lstStyle/><a:p><a:pPr algn="ctr"/><a:endParaRPr lang="de-DE"/></a:p></p:txBody></p:sp>`;

export const RAABE_DESIGN: DesignSpec = {
  name: "Raabeschule",
  colorName: "Raabeschule",
  fontName: "Raabeschule",
  major: "Calibri Light",
  minor: "Calibri",
  colors: { dk1: "000000", lt1: "FFFFFF", dk2: "3B3B3B", lt2: "F2F2F2", accent1: "B3141B", accent2: "595959", accent3: "D98C1E", accent4: "2E75B6", accent5: "7F7F7F", accent6: "5E9E3A", hlink: "B3141B", folHlink: "7A0E13" },
  // A red band along the top and a thin grey line above the footers.
  masterDeco: band(20, "Raabe-Band", 0, 0, 12192000, 228600, "accent1") + band(21, "Raabe-Linie", 838200, 6264000, 10515600, 12700, "accent2"),
  titleColor: "accent1",
};

export const CLEAR_DESIGN: DesignSpec = {
  name: "Klar",
  colorName: "Klar",
  fontName: "Klar",
  major: "Segoe UI Semibold",
  minor: "Segoe UI",
  colors: { dk1: "000000", lt1: "FFFFFF", dk2: "1F3864", lt2: "EEF2F7", accent1: "2F5597", accent2: "00A3A3", accent3: "F4B183", accent4: "8FAADC", accent5: "C55A11", accent6: "548235", hlink: "2F5597", folHlink: "7030A0" },
  masterDeco: band(20, "Seitenleiste", 0, 0, 228600, 6858000, "accent1"),
};

export const DARK_DESIGN: DesignSpec = {
  name: "Tafel",
  colorName: "Tafel",
  fontName: "Tafel",
  major: "Calibri Light",
  minor: "Calibri",
  colors: { dk1: "1E2B24", lt1: "FFFFFF", dk2: "2C3D33", lt2: "E7E6E6", accent1: "F2C14E", accent2: "6EC6CA", accent3: "F28482", accent4: "A7C957", accent5: "C3B1E1", accent6: "F7F7F7", hlink: "6EC6CA", folHlink: "C3B1E1" },
  dark: true,
  titleColor: "accent1",
};

export const DESIGNS: DesignSpec[] = [OFFICE_DESIGN, RAABE_DESIGN, CLEAR_DESIGN, DARK_DESIGN];
