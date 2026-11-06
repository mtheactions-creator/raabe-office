// Charts on slides: the chart part read by the .xlsx chart reader and drawn by the chart engine
// shared with RaabeCalc (D-062), from the values PowerPoint keeps in the part's caches.

import { renderChart as drawChart } from "@raabe/charts";
import { chartFromXml, type Pptx } from "@raabe/ooxml";

const cache = new Map<string, { xml: string; model: ReturnType<typeof chartFromXml> }>();

export function renderChart(pptx: Pptx, path: string, w: number, h: number, accents?: string[]): HTMLElement {
  const box = document.createElement("div");
  box.style.cssText = `position:absolute;left:0;top:0;width:${w}px;height:${h}px`;
  const xml = pptx.text(path);
  if (!xml) return box;
  try {
    let hit = cache.get(path);
    if (!hit || hit.xml !== xml) { hit = { xml, model: chartFromXml(xml) }; cache.set(path, hit); }
    box.append(drawChart(hit.model, w, h, accents ? { accents } : {}));
  } catch {
    box.style.background = "#f2f2f2";
  }
  return box;
}
