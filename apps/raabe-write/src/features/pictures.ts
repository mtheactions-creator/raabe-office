// Bilder: Einfügen > Bilder > Dieses Gerät, and the contextual tab "Bildformat" with
// Textumbruch and Größe. Datum und Uhrzeit lives here too (Einfügen > Text).

import { closePopup, h, showDialog, showPopup, messageBox, baseName, type Command, type MenuEntry } from "@raabe/shell";
import { imageInfo, naturalSizePt } from "@raabe/text-engine";
import type { AppApi } from "./api";
import { pictureTab, PICTURE_TAB_COMMANDS } from "./picture-tab";

export { PICTURE_TAB_COMMANDS };

/** A small icon of a wrapping style (text lines and a picture), for the layout options. */
/**
 * Word's Textumbruch icons: where the text lines (grey) run around the picture (blue), so each
 * option shows what it does (Layoutoptionen, Textumbruch menu).
 */
export function wrapIcon(wrap: string): HTMLElement {
  const L = (x1: number, x2: number, y: number) => `<path d="M${x1} ${y}H${x2}" stroke="#8a8a8a" stroke-width="2"/>`;
  const pic = (x: number, y: number, w: number, hh: number, o = 1) => `<rect x="${x}" y="${y}" width="${w}" height="${hh}" rx="1" fill="#4f8fd6" fill-opacity="${o}"/>`;
  let body = "";
  switch (wrap) {
    case "inline":
      body = L(2, 30, 5) + L(2, 9, 19) + pic(11, 11, 10, 9) + L(23, 30, 19) + L(2, 30, 27);
      break;
    case "square":
      body = L(2, 30, 4) + [10, 16, 22].map((y) => L(2, 8, y) + L(24, 30, y)).join("") + pic(11, 9, 10, 14) + L(2, 30, 28);
      break;
    case "tight":
      body = L(2, 30, 4) + L(2, 10, 10) + L(22, 30, 10) + L(2, 7, 16) + L(25, 30, 16) + L(2, 10, 22) + L(22, 30, 22) +
        '<circle cx="16" cy="16" r="7" fill="#4f8fd6"/>' + L(2, 30, 28);
      break;
    case "through":
      // Text runs into the hollow of the shape.
      body = L(2, 30, 4) + L(2, 9, 10) + L(23, 30, 10) + L(2, 18, 16) + L(23, 30, 16) + L(2, 9, 22) + L(23, 30, 22) +
        '<path d="M11 8h11v16H11v-4h6v-8h-6z" fill="#4f8fd6"/>' + L(2, 30, 28);
      break;
    case "topAndBottom":
      body = L(2, 30, 4) + L(2, 30, 8) + pic(10, 11, 12, 10) + L(2, 30, 24) + L(2, 30, 28);
      break;
    case "behind":
      body = pic(8, 7, 16, 18, 0.35) + [4, 10, 16, 22, 28].map((y) => L(2, 30, y)).join("");
      break;
    default: // front
      body = [4, 10, 16, 22, 28].map((y) => L(2, 30, y)).join("") + pic(9, 8, 14, 16);
  }
  const box = h("div", { style: "width:32px;height:32px" });
  box.innerHTML = `<svg viewBox="0 0 32 32" width="32" height="32" aria-hidden="true">${body}</svg>`;
  return box;
}

const cmd = (id: string, label: string, icon?: string, extra: Partial<Command> = {}): Command => ({ id, label, icon, ...extra });

export const PICTURE_COMMANDS: Command[] = [
  cmd("picture.fromFile", "Dieses Gerät...", "image-add", { keywords: ["Bild einfügen", "Grafik", "Foto"] }),
  cmd("picture.wrapInline", "Mit Text in Zeile", "wrap-inline"),
  cmd("picture.wrapSquare", "Quadrat", "wrap-square"),
  cmd("picture.wrapTight", "Eng", "wrap-tight"),
  cmd("picture.wrapThrough", "Transparent", "wrap-through"),
  cmd("picture.wrapTopBottom", "Oben und unten", "wrap-top-bottom"),
  cmd("picture.wrapBehind", "Hinter den Text", "wrap-behind"),
  cmd("picture.wrapFront", "Vor den Text", "wrap-front"),
  cmd("picture.height", "Höhe der Form", undefined),
  cmd("picture.width", "Breite der Form", undefined),
  cmd("picture.changeFromFile", "Bild ändern", "change-picture"),
  cmd("picture.changeFromClipboard", "Aus Zwischenablage...", "change-picture"),
  cmd("picture.saveAs", "Als Grafik speichern...", "save-picture"),
  cmd("picture.moreLayout", "Weitere Layoutoptionen...", "size-position"),
  cmd("picture.setDefaultLayout", "Als Standardlayout festlegen"),
  cmd("picture.altText", "Alternativtext", "alt-text"),
  cmd("picture.sizePosition", "Größe und Position...", "size-position"),
  cmd("picture.formatPane", "Grafik formatieren...", "format-picture"),
  cmd("picture.styles", "Bildformatvorlagen", "picture-styles"),
  cmd("picture.crop", "Zuschneiden", "crop"),
];

const MIME: Record<string, string> = { png: "image/png", jpg: "image/jpeg", jpeg: "image/jpeg", gif: "image/gif", bmp: "image/bmp", svg: "image/svg+xml", webp: "image/webp" };

/**
 * Natural size of a picture in points: from the file's resolution as Word does (a 500 dpi
 * PNG comes in small), 96 dpi without one; other formats from the decoded picture.
 */
function naturalSize(bytes: Uint8Array, type: string): Promise<{ w: number; h: number }> {
  const info = imageInfo(bytes);
  if (info) return Promise.resolve(naturalSizePt(info));
  return new Promise((resolve, reject) => {
    const url = URL.createObjectURL(new Blob([bytes as BlobPart], { type }));
    const img = new Image();
    img.onload = () => { resolve({ w: (img.naturalWidth || 300) * 0.75, h: (img.naturalHeight || 200) * 0.75 }); URL.revokeObjectURL(url); };
    img.onerror = () => { URL.revokeObjectURL(url); reject(new Error("Das Bild konnte nicht gelesen werden.")); };
    img.src = url;
  });
}

export function picturesFeature(api: AppApi) {
  const ed = () => api.editor();
  const c = api.commands;
  const has = () => !!ed();
  const sel = () => ed()?.selectedImage() ?? null;
  const picked = () => !!sel();

  const fromFile = async () => {
    const path = await api.platform.openImageDialog();
    if (path) await insertFromPath(path);
  };
  const insertFromPath = async (path: string) => {
    try {
      const bytes = await api.platform.readFile(path);
      const ext = baseName(path).split(".").pop()?.toLowerCase() ?? "";
      // The file's own header decides (a ".png" may well be a JPEG).
      const type = imageInfo(bytes)?.type ?? MIME[ext] ?? "image/png";
      const size = await naturalSize(bytes, type);
      ed()?.insertImage({ bytes, type }, size.w, size.h);
      api.selectTab("pictureFormat");
    } catch (e) {
      await messageBox("RaabeWrite", `Das Bild „${baseName(path)}“ konnte nicht eingefügt werden.\n\n${String(e)}`);
    }
    api.focus();
  };
  c.implement("insert.pictures", { run: () => {}, enabled: has });
  c.implement("picture.fromFile", { run: fromFile, enabled: has });

  const wraps: [string, string][] = [
    ["picture.wrapInline", "inline"], ["picture.wrapSquare", "square"], ["picture.wrapTight", "tight"], ["picture.wrapThrough", "through"],
    ["picture.wrapTopBottom", "topAndBottom"], ["picture.wrapBehind", "behind"], ["picture.wrapFront", "front"],
  ];
  const currentWrap = () => { const s = sel(); return !s ? "" : s.kind === "inline" ? "inline" : s.obj.wrap; };
  for (const [id, wrap] of wraps) {
    c.implement(id, { run: () => { ed()?.setImageWrap(wrap); api.focus(); }, enabled: () => picked() && ed()!.doc.activeStory === "main", checked: () => currentWrap() === wrap });
  }
  c.implement("layout.wrapText", { run: () => {}, enabled: () => picked() && ed()!.doc.activeStory === "main" });

  c.implement("picture.height", { run: () => {}, enabled: picked });
  c.implement("picture.width", { run: () => {}, enabled: picked });

  // The Bildformat tab (picture-tab.ts) and Word's layout options button beside a picture.
  const { tab, arrange } = pictureTab(api, insertFromPath);
  const layoutOptions = (rect: DOMRect) => {
    const s0 = sel();
    if (!s0) return;
    const icon = (id: string, wrap: string) => {
      const cur = currentWrap() === wrap;
      return h("button", { class: `tile${cur ? " on" : ""}`, "data-tip": c.get(id).label, "aria-label": c.get(id).label, style: "width:42px;height:42px", onclick: () => { closePopup(); void c.run(id); } }, wrapIcon(wrap));
    };
    const panel = h("div", { class: "rpanel layout-options", style: "width:220px" },
      h("div", { class: "title" }, "Layoutoptionen"),
      h("div", { class: "subtitle" }, "Mit Text in Zeile"), h("div", { class: "tiles" }, icon("picture.wrapInline", "inline")),
      h("div", { class: "subtitle" }, "Mit Textumbruch"),
      h("div", { class: "tiles", style: "grid-template-columns:repeat(4, auto)" }, ...wraps.slice(1).map(([id, w]) => icon(id, w))),
      h("hr"),
      h("label", { class: "check", style: "display:flex" }, h("input", { type: "radio", name: "lo-move", checked: s0.kind === "float" && s0.obj.moveWithText !== false, disabled: s0.kind !== "float", onchange: () => { ed()?.setPictureMoveWithText(true); } }), "Mit Text verschieben"),
      h("label", { class: "check", style: "display:flex" }, h("input", { type: "radio", name: "lo-move", checked: s0.kind === "float" && s0.obj.moveWithText === false, disabled: s0.kind !== "float", onchange: () => { ed()?.setPictureMoveWithText(false); } }), "Position auf der Seite fixieren"),
      h("button", { class: "link", onclick: () => { closePopup(); void c.run("picture.moreLayout"); } }, "Mehr anzeigen"));
    showPopup(panel, new DOMRect(rect.left, rect.bottom, 0, 0));
  };
  const attach = (e: import("@raabe/text-engine").Editor) => { e.onLayoutOptions = layoutOptions; };

  // ----- Datum und Uhrzeit --------------------------------------------------------------
  const dateTime = async () => {
    const e = ed();
    if (!e) return;
    const now = new Date();
    const pad = (n: number) => String(n).padStart(2, "0");
    const months = ["Januar", "Februar", "März", "April", "Mai", "Juni", "Juli", "August", "September", "Oktober", "November", "Dezember"];
    const days = ["Sonntag", "Montag", "Dienstag", "Mittwoch", "Donnerstag", "Freitag", "Samstag"];
    const d = now.getDate(), m = now.getMonth(), y = now.getFullYear();
    const formats: [string, string][] = [
      ["dd.MM.yyyy", `${pad(d)}.${pad(m + 1)}.${y}`],
      ["dddd, d. MMMM yyyy", `${days[now.getDay()]}, ${d}. ${months[m]} ${y}`],
      ["d. MMMM yyyy", `${d}. ${months[m]} ${y}`],
      ["dd.MM.yy", `${pad(d)}.${pad(m + 1)}.${String(y).slice(2)}`],
      ["yyyy-MM-dd", `${y}-${pad(m + 1)}-${pad(d)}`],
      ["d. MMM. yyyy", `${d}. ${months[m].slice(0, 3)}. ${y}`],
      ["dd.MM.yyyy HH:mm", `${pad(d)}.${pad(m + 1)}.${y} ${pad(now.getHours())}:${pad(now.getMinutes())}`],
      ["HH:mm", `${pad(now.getHours())}:${pad(now.getMinutes())}`],
    ];
    const list = h("select", { size: 8, style: "width:100%" }, ...formats.map(([f, t]) => h("option", { value: f }, t)));
    list.selectedIndex = 0;
    const auto = h("input", { type: "checkbox" });
    const ok = await showDialog({
      title: "Datum und Uhrzeit",
      body: h("div", { style: "width:320px" }, h("div", { style: "margin-bottom:6px" }, "Verfügbare Formate:"), list,
        h("label", { class: "check", style: "margin-top:8px" }, auto, "Automatisch aktualisieren")),
      buttons: [{ label: "OK", value: true, primary: true }, { label: "Abbrechen", value: false }], cancel: false,
    });
    if (!ok) return;
    const [fmt, text] = formats[Math.max(0, list.selectedIndex)];
    if (auto.checked) e.insertField(`${fmt.includes("H") && !fmt.includes("d") ? "TIME" : "DATE"} \\@ "${fmt}"`, text);
    else e.insertText(text);
    api.focus();
  };
  c.implement("insert.dateTime", { run: dateTime, enabled: has });

  // Layout > Anordnen: the same as on Bildformat and Formformat, for the selected object.
  const objectPicked = () => picked() && ed()!.doc.activeStory === "main";
  c.implement("layout.position", { run: () => {}, enabled: objectPicked });
  c.implement("layout.bringForward", { run: () => c.run("picture.bringForward"), enabled: () => sel()?.kind === "float" });
  c.implement("layout.sendBackward", { run: () => c.run("picture.sendBackward"), enabled: () => sel()?.kind === "float" });
  c.implement("layout.align", { run: () => {}, enabled: picked });
  c.implement("layout.rotate", { run: () => {}, enabled: picked });
  return { tab, insertMenu: [{ cmd: "picture.fromFile" }] as MenuEntry[], insertFromPath, attach, arrange };
}
