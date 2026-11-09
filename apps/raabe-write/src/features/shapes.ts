// Formen und Textfelder (D-038, package C): Einfügen > Formen with Word's gallery (order and
// names recorded from Word 365, probe_word_shapes_ui.py), drawing on the page, Einfügen >
// Textfeld, and the contextual tab Formformat (shape-tab.ts).

import { h, type Command, type MenuEntry } from "@raabe/shell";
import { SHAPE_GALLERY, shapeIcon, type GalleryEntry } from "@raabe/text-engine";

export { SHAPE_GALLERY, shapeIcon };
import type { AppApi } from "./api";
import { shapeTab, SHAPE_TAB_COMMANDS } from "./shape-tab";

const cmd = (id: string, label: string, icon?: string, extra: Partial<Command> = {}): Command => ({ id, label, icon, ...extra });

export const SHAPE_COMMANDS: Command[] = [
  cmd("insert.textBoxDraw", "Textfeld erstellen", "draw-text-box", { tip: "Textfeld erstellen\nEin Textfeld an beliebiger Stelle im Dokument zeichnen." }),
  // Word saves a shape with its text as a picture; we cannot draw one into a file yet.
  cmd("shape.saveAs", "Als Grafik speichern...", "save-picture", { deferred: true }),
  // Word's drawing canvas (wpc) holds shapes drawn together; not built. Its entry was a button without a command (D-134).
  cmd("insert.drawingCanvas", "Neuer Zeichenbereich", undefined, { deferred: true }),
  ...SHAPE_TAB_COMMANDS,
];

type Entry = GalleryEntry;

export function shapesFeature(api: AppApi) {
  const c = api.commands;
  const ed = () => api.editor();
  const has = () => !!ed();
  const recent: Entry[] = [];

  const draw = (e: Entry) => {
    const editor = ed();
    if (!editor) return;
    const i = recent.findIndex((r) => r.label === e.label);
    if (i >= 0) recent.splice(i, 1);
    recent.unshift(e);
    if (recent.length > 24) recent.pop();
    editor.onDrawEnd = () => { api.refresh(); api.selectTab("shapeFormat"); };
    editor.startDrawing(e.tool);
  };

  /** The gallery (Einfügen > Formen, Formformat > Formen einfügen). */
  const galleryPanel = (close: () => void, sections = SHAPE_GALLERY, withRecent = true) => {
    const tile = (e: Entry) => {
      const b = h("button", { class: "tile shape-tile", "data-tip": e.label, "aria-label": e.label, onclick: () => { close(); draw(e); } });
      b.append(shapeIcon(e));
      return b;
    };
    const parts: HTMLElement[] = [];
    const all = withRecent && recent.length ? [{ title: "Zuletzt verwendete Formen", items: recent }, ...sections] : sections;
    for (const s of all) parts.push(h("div", { class: "title" }, s.title), h("div", { class: "tiles shape-tiles" }, ...s.items.map(tile)));
    return h("div", { class: "rpanel shape-gallery" }, ...parts,
      h("hr"), h("button", { class: "link", disabled: !c.isEnabled("insert.drawingCanvas"), "data-tip": c.tooltip("insert.drawingCanvas", false),
        onclick: () => { close(); void c.run("insert.drawingCanvas"); } }, "Neuer Zeichenbereich"));
  };
  c.implement("insert.shapes", { run: () => {}, enabled: has });

  // Einfügen > Textfeld: the gallery of text box designs and Textfeld erstellen.
  const textBox: Entry = { tool: { geometry: "rect", kind: "textbox" }, label: "Textfeld" };
  c.implement("insert.textBoxDraw", { run: () => draw(textBox), enabled: has });
  const insertDesign = (design: TextBoxDesign) => {
    const e = ed();
    if (!e) return;
    const pg = e.layout.pages[e.currentPage()] ?? e.layout.pages[0];
    if (!pg) return;
    const PT = 96 / 72;
    const w = design.w * PT, hh = design.h * PT;
    const x = design.x === "right" ? (pg.setup.width - pg.setup.marginRight) * PT - w : design.x === "center" ? (pg.setup.width * PT - w) / 2 : pg.setup.marginLeft * PT;
    const y = pg.setup.marginTop * PT + 20;
    e.insertShape("rect", "textbox", pg.index, x, y, w, hh, { pic: design.pic, text: "" });
    api.selectTab("shapeFormat");
  };
  const textBoxMenu: MenuEntry[] = [
    { heading: "Integriert" },
    ...TEXT_BOX_DESIGNS.map((d) => ({ label: d.label, icon: "text-box", run: () => insertDesign(d) })),
    "-",
    { cmd: "insert.textBoxDraw" },
  ];
  c.implement("insert.textBox", { run: () => {}, enabled: has });

  // The in-ribbon gallery (Formformat > Formen einfügen): Word's most used shapes first.
  const all = SHAPE_GALLERY.flatMap((s) => s.items);
  const pick = ["Textfeld", "Linie", "Linienpfeil", "Rechteck", "Ellipse", "Rechteck: abgerundete Ecken", "Gleichschenkliges Dreieck",
    "Verbinder: gewinkelt", "Verbinder: gewinkelt mit Pfeil", "Pfeil: nach rechts", "Pfeil: nach unten", "Freihandform: Form", "Freihand: Skizze",
    "Kurve", "Geschweifte Klammer links", "Geschweifte Klammer rechts", "Stern: 5 Zacken", "Sprechblase: rechteckig", "Raute", "Herz", "Wolke"];
  const ribbon = pick.map((l) => all.find((e) => e.label === l)!).filter(Boolean);
  const { tab, layoutCommands } = shapeTab(api, galleryPanel, {
    count: ribbon.length, icon: (i) => shapeIcon(ribbon[i], 16), label: (i) => ribbon[i].label, draw: (i) => draw(ribbon[i]),
  });
  // Form ändern: Word's gallery without lines and text box; the shape keeps its look.
  const changeSections = SHAPE_GALLERY.filter((s) => s.title !== "Linien").map((s) => ({ ...s, items: s.items.filter((e) => e.tool.kind === "shape" && !e.tool.freeform) }));
  layoutCommands.setChangePanel(() => {
    const tile = (e: Entry) => {
      const b = h("button", {
        class: "tile shape-tile", "data-tip": e.label, "aria-label": e.label,
        onclick: () => {
          import("@raabe/shell").then((m) => m.closePopup());
          const editor = ed();
          editor?.setShapeFormat((f) => ({ ...f, geometry: e.tool.geometry === "rect" ? undefined : e.tool.geometry }));
          editor?.setShape((s) => ({ ...s, adjust: undefined }), false);
          api.refresh();
          api.focus();
        },
      });
      b.append(shapeIcon(e));
      return b;
    };
    return h("div", { class: "rpanel shape-gallery" }, ...changeSections.flatMap((s) => [h("div", { class: "title" }, s.title), h("div", { class: "tiles shape-tiles" }, ...s.items.map(tile))]));
  });
  return { tab, galleryPanel, textBoxMenu, draw, layoutCommands };
}

interface TextBoxDesign { label: string; w: number; h: number; x: "left" | "center" | "right"; pic?: import("@raabe/text-engine").PictureFormat }

/** Text box designs of our own (D-046), each made of ordinary shape settings. */
const TEXT_BOX_DESIGNS: TextBoxDesign[] = [
  { label: "Einfaches Textfeld", w: 220, h: 80, x: "center" },
  { label: "Randleiste rechts", w: 150, h: 560, x: "right", pic: { fill: { kind: "solid", color: "DEEBF7" }, fillXml: '<a:solidFill><a:schemeClr val="accent1"><a:lumMod val="20000"/><a:lumOff val="80000"/></a:schemeClr></a:solidFill>', line: null } },
  { label: "Merkkasten", w: 300, h: 90, x: "center", pic: { fill: { kind: "solid", color: "FFF2CC" }, fillXml: '<a:solidFill><a:schemeClr val="accent4"><a:lumMod val="20000"/><a:lumOff val="80000"/></a:schemeClr></a:solidFill>', line: { color: "FFC000", width: 1.5 } } },
  { label: "Zitat", w: 300, h: 70, x: "center", pic: { line: null, fill: { kind: "none" }, fillXml: "<a:noFill/>" } },
];
