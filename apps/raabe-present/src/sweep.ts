// The by-hand sweep's starting points (tools/feel-lab/sweep_present.py, D-137): a fresh presentation with
// what a student has in front of them before using an option (a shape with text, a text being edited, a
// picture, a table, a chart, an equation, a SmartArt, the master views). Only the start is made here; the
// options themselves are triggered through the real controls by the sweep (CDP mouse and key events).

import type { PresentApp } from "./app";
import { insertChart } from "./charts";
import { insertEquation } from "./equations";
import { createDiagram } from "./smartart-create";
import { SMARTART_LAYOUTS } from "./smartart-catalog.gen";
import { startFreistellen } from "./freistellen";

type Spec = { layout: string; slides?: number; shapes: { kind: string; box?: number[]; text?: string; ph?: number }[]; zoom?: number };
type Api = { setupCase: (s: Spec) => number };

/** A picture of 240 x 150 px (a green block with a darker square), made on a canvas. */
async function pictureBytes(): Promise<Uint8Array> {
  const c = document.createElement("canvas");
  c.width = 240; c.height = 150;
  const g = c.getContext("2d")!;
  g.fillStyle = "#3a7"; g.fillRect(0, 0, 240, 150);
  g.fillStyle = "#246"; g.fillRect(80, 40, 80, 70);
  const b = await new Promise<Blob | null>((r) => c.toBlob(r));
  return new Uint8Array(await b!.arrayBuffer());
}

const frame = () => new Promise((r) => setTimeout(r, 30));

/** Leaves every mode a previous option may have started (show, backstage, master view, panes, tools). */
export function leaveModes(app: PresentApp): void {
  const a = app as unknown as Record<string, any>;
  try { a.show?.end(); } catch { /* already ended */ }
  try { a.backstage?.close(); } catch { /* closed */ }
  try { if (a.masterMode) app.setMasterView(false); } catch { /* not in master view */ }
  try { if (a.view !== "normal") a.setView("normal"); } catch { /* normal */ }
  app.editor.tool = null;
  for (const b of Array.from(document.querySelectorAll<HTMLElement>(".present-main aside:not([hidden]) .tp-close, .present-main aside:not([hidden]) [aria-label=Schließen]"))) b.click();
  for (const d of Array.from(document.querySelectorAll(".dialog-layer, .menu-root, body > .menu, body > .popup, .tooltip"))) d.remove();
}

/** Makes the starting point `kind` and returns what is selected (for the report). */
export async function context(app: PresentApp, api: Api, kind: string): Promise<string> {
  leaveModes(app);
  const ed = app.editor;
  const shapeIds = () => app.deck.view(app.slidePath()!).shapes.filter((s) => !s.ph).map((s) => s.id);
  switch (kind) {
    case "slides": {
      // Three slides: a title slide with text, a Titel und Inhalt slide with bullets, a blank one with a shape.
      api.setupCase({ layout: "Titelfolie", slides: 1, shapes: [{ kind: "ph", ph: 0, text: "Titel der Präsentation" }, { kind: "ph", ph: 1, text: "Untertitel" }], zoom: 50 });
      const lay = app.deck.layouts();
      app.deck.addSlide(1, lay.find((l) => l.name === "Titel und Inhalt")!.path);
      app.deck.addSlide(2, lay.find((l) => l.name === "Leer")!.path);
      app.deck.clearHistory();
      (app as unknown as { active: { savedVersion: number } }).active.savedVersion = app.deck.version;
      app.renderAll();
      return "slide 1 of 3";
    }
    case "shape": case "shape-text": case "textbox": {
      const box = kind === "textbox" ? { kind: "textbox", box: [200, 150, 300, 40], text: "Hallo Welt" } : { kind: "rect", box: [200, 150, 300, 160], text: "Hallo Welt" };
      api.setupCase({ layout: "Leer", slides: 2, shapes: [box], zoom: 50 });
      const id = shapeIds()[0];
      if (kind === "shape-text") ed.enterText(id, "all"); else ed.select([id]);
      return kind === "shape-text" ? "all text of a rectangle" : `a ${kind === "textbox" ? "text box" : "rectangle"} with text`;
    }
    case "two-shapes": {
      api.setupCase({ layout: "Leer", slides: 1, shapes: [{ kind: "rect", box: [100, 100, 200, 120], text: "Eins" }, { kind: "ellipse", box: [400, 220, 160, 160], text: "Zwei" }], zoom: 50 });
      ed.select(shapeIds());
      return "two shapes";
    }
    case "placeholder": {
      api.setupCase({ layout: "Titel und Inhalt", slides: 1, shapes: [{ kind: "ph", ph: 0, text: "Titel" }, { kind: "ph", ph: 1, text: "Erster Punkt\rZweiter Punkt\rDritter Punkt" }], zoom: 50 });
      const ph = app.deck.view(app.slidePath()!).shapes.filter((s) => s.ph)[1];
      ed.enterText(ph.id, "all");
      return "all text of the content placeholder";
    }
    case "picture": {
      api.setupCase({ layout: "Leer", slides: 1, shapes: [], zoom: 50 });
      app.insertPictureBytes(await pictureBytes(), "image/png");
      await frame();
      app.deck.clearHistory();
      return "a picture";
    }
    case "freistellen": {
      await context(app, api, "picture");
      await startFreistellen(app);
      app.deck.clearHistory();
      return "a picture in Freistellen";
    }
    case "table": {
      api.setupCase({ layout: "Leer", slides: 1, shapes: [{ kind: "table", box: [100, 120, 600, 120] }], zoom: 50 });
      ed.enterText(shapeIds()[0], undefined, { r: 0, c: 0 });
      return "the caret in a table's first cell";
    }
    case "chart": {
      api.setupCase({ layout: "Leer", slides: 1, shapes: [], zoom: 50 });
      insertChart(app, "column");
      await frame();
      const c = app.deck.view(app.slidePath()!).shapes.find((s) => s.frame === "chart");
      if (c) ed.select([c.id]);
      app.deck.clearHistory();
      return "a column chart";
    }
    case "equation": {
      api.setupCase({ layout: "Leer", slides: 1, shapes: [], zoom: 50 });
      await insertEquation(app);
      await frame();
      app.deck.clearHistory();
      return "the caret in a new equation";
    }
    case "smartart": {
      api.setupCase({ layout: "Leer", slides: 1, shapes: [], zoom: 50 });
      const size = app.deck.pptx.slideSize;
      const id = createDiagram(app, SMARTART_LAYOUTS[0], [{ text: "Eins", level: 0 }, { text: "Zwei", level: 0 }, { text: "Drei", level: 0 }], { x: 838200, y: 1825625, w: size.cx - 2 * 838200, h: 4351338 });
      ed.select([id]);
      app.deck.clearHistory();
      return "a SmartArt (Einfache Blockliste)";
    }
    case "master": case "notesmaster": case "handoutmaster": {
      await context(app, api, "slides");
      app.setMasterView(true, kind === "master" ? "slides" : kind === "notesmaster" ? "notes" : "handout");
      app.deck.clearHistory();
      return `${kind} view`;
    }
    case "nothing": {
      api.setupCase({ layout: "Leer", slides: 1, shapes: [], zoom: 50 });
      return "an empty slide";
    }
  }
  return "?";
}
