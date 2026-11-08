// A chart part drawn as an SVG picture (Step 7b, D-088): what RaabeWrite shows for a chart in the text, from the values
// the chart part keeps, with the chart engine shared by the three apps. Needs a canvas to measure text (the browser).

import { LOOK, chartScene, sceneMarkup, type ChartModel } from "@raabe/charts";
import { chartFromXml } from "./xlsx/chart";

/** The SVG of a chart (part XML or model) at a size in points, the document theme's accents ("#RRGGBB") given. */
export function chartSvg(chart: string | ChartModel, widthPt: number, heightPt: number, accents?: string[]): string {
  const model = typeof chart === "string" ? chartFromXml(chart) : chart;
  const scene = chartScene(model, widthPt, heightPt, accents ? { accents } : {});
  return `<svg xmlns="http://www.w3.org/2000/svg" width="${(widthPt * 96) / 72}" height="${(heightPt * 96) / 72}" viewBox="0 0 ${scene.width} ${scene.height}" font-family="${LOOK.font}">${sceneMarkup(scene)}</svg>`;
}

/** The same as picture bytes for an image object. */
export function chartImage(chart: string | ChartModel, widthPt: number, heightPt: number, accents?: string[]): { bytes: Uint8Array; type: string } {
  return { bytes: new TextEncoder().encode(chartSvg(chart, widthPt, heightPt, accents)), type: "image/svg+xml" };
}
