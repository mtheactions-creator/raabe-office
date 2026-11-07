// Bildformat > Hintergrund entfernen opens Freistellen (D-072), as in PowerPoint: Zu behaltende Bereiche
// markieren and Zu entfernende Bereiche markieren draw lines on the picture (a14:foregroundMark and
// backgroundMark, written as PowerPoint writes them), the picture is drawn again after each line;
// Alle Änderungen verwerfen takes the removal away, Änderungen beibehalten ends Freistellen.

import type { Control, Tab } from "@raabe/shell";
import { PptxDom } from "@raabe/ooxml";
import type { PresentApp } from "./app";
import { effectsOf, setEffect, type BgMark } from "./picfx";

let activeId: number | null = null;

export const freistellenActive = (): number | null => activeId;

function removal(app: PresentApp, id: number): { attrs: Record<string, string>; marks: BgMark[] } | null {
  const blip = PptxDom.at(app.editor.shapeEl(id)!, "blipFill/blip");
  const fx = blip ? effectsOf(blip).find((e) => e.name === "backgroundRemoval") : undefined;
  return fx ? { attrs: fx.attrs, marks: fx.marks ?? [] } : null;
}

/** Hintergrund entfernen: the automatic removal (when not there yet), then Freistellen. */
export async function startFreistellen(app: PresentApp): Promise<void> {
  const v = app.editor.selectedViews.find((x) => x.kind === "pic");
  if (!v) return;
  if (!removal(app, v.id)) await setEffect(app, "backgroundRemoval", { name: "backgroundRemoval", attrs: { t: "10000", b: "90000", l: "10000", r: "90000" } }, "Hintergrund entfernen");
  activeId = v.id;
  app.editor.select([v.id]);
  app.updateChrome();
  app.showTab("freistellen");
}

export function endFreistellen(app: PresentApp, keep: boolean): void {
  const id = activeId;
  activeId = null;
  if (id !== null && !keep) { app.editor.select([id]); void setEffect(app, "backgroundRemoval", null, "Alle Änderungen verwerfen"); }
  app.updateChrome();
}

/** A line drawn on the picture: a mark (to keep or to remove), the picture drawn again. */
export function markTool(app: PresentApp, keep: boolean): void {
  const id = activeId;
  if (id === null) return;
  app.editor.tool = { kind: "points", mode: "scribble", done: (pts) => {
    const v = app.editor.shapeView(id);
    const cur = removal(app, id);
    if (!v || !cur || pts.length < 2) return;
    // From slide EMU to 1/1000 % of the picture (as PowerPoint writes the marks).
    const fx = (x: number) => Math.max(0, Math.min(100000, ((x - v.x) / v.w) * 100000));
    const fy = (y: number) => Math.max(0, Math.min(100000, ((y - v.y) / v.h) * 100000));
    const marks = [...cur.marks];
    for (let i = 1; i < pts.length; i++) marks.push({ keep, x1: fx(pts[i - 1].x), y1: fy(pts[i - 1].y), x2: fx(pts[i].x), y2: fy(pts[i].y) });
    app.editor.select([id]);
    void setEffect(app, "backgroundRemoval", { name: "backgroundRemoval", attrs: cur.attrs, marks }, keep ? "Zu behaltende Bereiche markieren" : "Zu entfernende Bereiche markieren").then(() => app.updateChrome());
  } };
  app.editor.render();
  app.updateChrome();
}

export function freistellenTab(): Tab {
  const b = (cmd: string, key: string): Control => ({ type: "button", cmd, key, size: "large" });
  return {
    id: "freistellen", label: "Freistellen", key: "JF", contextual: true, shrink: [],
    groups: [
      { id: "refine", label: "Verfeinern", key: "ZV", icon: "remove-background", controls: [b("bg.markKeep", "K"), b("bg.markRemove", "R")] },
      { id: "close", label: "Schließen", key: "ZS", icon: "close", controls: [b("bg.discard", "D"), b("bg.keep", "C")] },
    ],
  };
}
