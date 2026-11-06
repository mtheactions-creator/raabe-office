export { readDocx, readThemeFile, rethemeXml, styleFromXml, themedHex, type DocxPackage, type ReadResult } from "./read";
export { writeDocx, styleXml } from "./write";
export { latexToMathml, latexToOmml, ommlToLatex, ommlToSeq, ommlToWordHtml, parseLatex, seqToOmml, wordHtmlMath } from "./math";
export { DEFAULT_THEME_COLORS, pictureFromSpPr, readColor } from "./drawingml";
export { DEFAULT_LINE_STYLE, DEFAULT_SHAPE_STYLE, shapePreset } from "./shapes";
export { readXlsx } from "./xlsx/read";
export { writeXlsx, type XlsxWriteOptions } from "./xlsx/write";
export { newWorkbook } from "./xlsx/new";
export { widthFromFile, widthToFile } from "./xlsx/units";
export { Pptx, RT as PPTX_RT, CT_PPTX, relTarget, type SlideRef } from "./pptx/package";
export * as PptxDom from "./pptx/dom";
export { newPptx, newPptxFiles, OFFICE_DESIGN, LAYOUT_NAMES, masterXml, designTheme, type DesignSpec } from "./pptx/template";
export {
  SlideResolver, slideView, readThemeInfo, readText, readBodyProps, levelProps, matchPlaceholder, phOf, spTreeOf, pathKind, GroupMap, PROMPTS,
  type SlideView, type ShapeView, type TextInfo, type TextPara, type TextRun, type RunStyle, type ParaStyle, type BodyProps, type Bullet, type ThemeInfo, type PicInfo, type TextContext,
} from "./pptx/resolve";
export { chartFromXml, chartToXml, chartStyleXml, chartColorsXml } from "./xlsx/chart";
export { Rels, relsPath, resolveTarget } from "./package";
