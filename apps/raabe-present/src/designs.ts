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

// ----- designs with artwork of our own (D-078): vector shapes on the master, in the design's colours --------

const clr = (c: string, alpha?: number) => `<a:${/^[0-9A-F]{6}$/.test(c) ? "srgbClr" : "schemeClr"} val="${c}">${alpha !== undefined ? `<a:alpha val="${Math.round(alpha * 100000)}"/>` : ""}</a:${/^[0-9A-F]{6}$/.test(c) ? "srgbClr" : "schemeClr"}>`;
const deco = (id: number, name: string, x: number, y: number, w: number, h: number, geom: string, fill: string) =>
  `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${name}"/><p:cNvSpPr/><p:nvPr userDrawn="1"/></p:nvSpPr><p:spPr><a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${w}" cy="${h}"/></a:xfrm>` +
  `${geom}<a:solidFill>${fill}</a:solidFill><a:ln><a:noFill/></a:ln></p:spPr><p:txBody><a:bodyPr rtlCol="0" anchor="ctr"/><a:lstStyle/><a:p><a:pPr algn="ctr"/><a:endParaRPr lang="de-DE"/></a:p></p:txBody></p:sp>`;
const prst = (p: string) => `<a:prstGeom prst="${p}"><a:avLst/></a:prstGeom>`;
/** A custom geometry from path commands in a w × h box ("M x y", "C x1 y1 x2 y2 x y", "L x y", "Z"). */
const cust = (w: number, h: number, cmds: string) => {
  const pt = (x: string, y: string) => `<a:pt x="${x}" y="${y}"/>`;
  const parts = cmds.trim().split(/(?=[MLCZ])/).map((c) => {
    const [k, ...n] = c.trim().split(/\s+/);
    if (k === "M") return `<a:moveTo>${pt(n[0], n[1])}</a:moveTo>`;
    if (k === "L") return `<a:lnTo>${pt(n[0], n[1])}</a:lnTo>`;
    if (k === "C") return `<a:cubicBezTo>${pt(n[0], n[1])}${pt(n[2], n[3])}${pt(n[4], n[5])}</a:cubicBezTo>`;
    return "<a:close/>";
  }).join("");
  return `<a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="0" t="0" r="r" b="b"/><a:pathLst><a:path w="${w}" h="${h}">${parts}</a:path></a:pathLst></a:custGeom>`;
};
const W = 12192000, H = 6858000;

export const WAVE_DESIGN: DesignSpec = {
  name: "Wellenband", colorName: "Wellenband", fontName: "Wellenband", major: "Segoe UI Light", minor: "Segoe UI",
  colors: { dk1: "000000", lt1: "FFFFFF", dk2: "16324F", lt2: "EAF2F8", accent1: "1F77B4", accent2: "2CA6A4", accent3: "F2A541", accent4: "5DA9E9", accent5: "E4572E", accent6: "76B041", hlink: "1F77B4", folHlink: "6A4C93" },
  masterDeco: deco(20, "Welle hinten", 0, 5300000, W, 1558000, cust(1000, 1000, "M 0 420 C 250 120 500 820 750 360 C 870 140 950 260 1000 300 L 1000 1000 L 0 1000 Z"), clr("accent2", 0.55)) +
    deco(21, "Welle vorne", 0, 5650000, W, 1208000, cust(1000, 1000, "M 0 500 C 220 900 480 120 720 520 C 850 740 940 560 1000 430 L 1000 1000 L 0 1000 Z"), clr("accent1")),
};

export const ARCS_DESIGN: DesignSpec = {
  name: "Kreisbögen", colorName: "Kreisbögen", fontName: "Kreisbögen", major: "Calibri Light", minor: "Calibri",
  colors: { dk1: "000000", lt1: "FFFFFF", dk2: "2B2D42", lt2: "F1F1F4", accent1: "EF476F", accent2: "FFD166", accent3: "06D6A0", accent4: "118AB2", accent5: "073B4C", accent6: "8D99AE", hlink: "118AB2", folHlink: "EF476F" },
  masterDeco: deco(20, "Kreis groß", 9300000, -2100000, 4600000, 4600000, prst("ellipse"), clr("accent1", 0.85)) +
    deco(21, "Kreis mittel", 10600000, 900000, 2400000, 2400000, prst("ellipse"), clr("accent2", 0.8)) +
    deco(22, "Kreis klein", -700000, 5500000, 1900000, 1900000, prst("ellipse"), clr("accent4", 0.7)),
};

export const WEDGE_DESIGN: DesignSpec = {
  name: "Keil", colorName: "Keil", fontName: "Keil", major: "Segoe UI Semibold", minor: "Segoe UI",
  colors: { dk1: "000000", lt1: "FFFFFF", dk2: "263238", lt2: "ECEFF1", accent1: "00897B", accent2: "26A69A", accent3: "FFB300", accent4: "5C6BC0", accent5: "EF5350", accent6: "8D6E63", hlink: "00897B", folHlink: "5C6BC0" },
  masterDeco: deco(20, "Keil dunkel", 0, 0, 1500000, H, cust(1000, 1000, "M 0 0 L 1000 0 L 420 1000 L 0 1000 Z"), clr("accent1")) +
    deco(21, "Keil hell", 0, 0, 2100000, H, cust(1000, 1000, "M 600 0 L 760 0 L 300 1000 L 140 1000 Z"), clr("accent2", 0.7)),
};

const dots = (() => {
  let s = "";
  let id = 20;
  for (let r = 0; r < 6; r++) for (let c = 0; c < 6 - r; c++) s += deco(id++, `Punkt ${id - 20}`, W - 380000 - c * 300000, 230000 + r * 300000, 120000, 120000, prst("ellipse"), clr(r % 2 ? "accent2" : "accent1"));
  return s;
})();

export const DOTS_DESIGN: DesignSpec = {
  name: "Punkteraster", colorName: "Punkteraster", fontName: "Punkteraster", major: "Calibri Light", minor: "Calibri",
  colors: { dk1: "000000", lt1: "FFFFFF", dk2: "3A3A3A", lt2: "F4F1EC", accent1: "B5651D", accent2: "6B8E23", accent3: "4682B4", accent4: "CD853F", accent5: "8B4513", accent6: "708090", hlink: "4682B4", folHlink: "8B4513" },
  masterDeco: dots + band(60, "Linie", 838200, 6264000, 10515600, 12700, "accent1"),
};

const stripes = (() => {
  let s = "";
  for (let i = 0; i < 5; i++) s += deco(20 + i, `Streifen ${i + 1}`, 7400000 + i * 520000, 0, 900000, 900000, cust(1000, 1000, "M 300 0 L 620 0 L 320 1000 L 0 1000 Z"), clr(i % 2 ? "accent2" : "accent1", i % 2 ? 0.8 : 1));
  return s;
})();

export const STRIPES_DESIGN: DesignSpec = {
  name: "Schräge Streifen", colorName: "Schräge Streifen", fontName: "Schräge Streifen", major: "Segoe UI Light", minor: "Segoe UI",
  colors: { dk1: "000000", lt1: "FFFFFF", dk2: "2E1A47", lt2: "F3EFF8", accent1: "7B2CBF", accent2: "C77DFF", accent3: "FF9E00", accent4: "3A86FF", accent5: "FB5607", accent6: "8AC926", hlink: "3A86FF", folHlink: "7B2CBF" },
  masterDeco: stripes + band(30, "Fußlinie", 0, 6750000, W, 108000, "accent1"),
};

export const EVENING_DESIGN: DesignSpec = {
  name: "Abendblau", colorName: "Abendblau", fontName: "Abendblau", major: "Calibri Light", minor: "Calibri",
  colors: { dk1: "0B1D3A", lt1: "FFFFFF", dk2: "142B52", lt2: "E3ECF7", accent1: "4FC3F7", accent2: "FFD54F", accent3: "81C784", accent4: "F06292", accent5: "BA68C8", accent6: "FF8A65", hlink: "4FC3F7", folHlink: "BA68C8" },
  dark: true,
  titleColor: "accent1",
  background: `<p:bgPr><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:srgbClr val="142B52"/></a:gs><a:gs pos="100000"><a:srgbClr val="0B1D3A"/></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill><a:effectLst/></p:bgPr>`,
  masterDeco: deco(20, "Horizont", 0, 6200000, W, 658000, cust(1000, 1000, "M 0 600 C 300 200 700 200 1000 600 L 1000 1000 L 0 1000 Z"), clr("accent1", 0.25)),
};

export const DESIGNS: DesignSpec[] = [OFFICE_DESIGN, RAABE_DESIGN, CLEAR_DESIGN, DARK_DESIGN, WAVE_DESIGN, ARCS_DESIGN, WEDGE_DESIGN, DOTS_DESIGN, STRIPES_DESIGN, EVENING_DESIGN];
