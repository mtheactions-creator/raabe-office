// Bilder: Einfügen > Bilder > Dieses Gerät, and the contextual tab "Bildformat" with
// Textumbruch and Größe. Datum und Uhrzeit lives here too (Einfügen > Text).

import { closePopup, h, showDialog, showPopup, messageBox, baseName, type Command, type MenuEntry } from "@raabe/shell";
import { imageInfo, naturalSizePt } from "@raabe/text-engine";
import type { AppApi } from "./api";
import { pictureTab, PICTURE_TAB_COMMANDS } from "./picture-tab";

export { PICTURE_TAB_COMMANDS };

/** A small icon of a wrapping style (text lines and a picture), for the layout options. */
function wrapIcon(wrap: string): HTMLElement {
  const box = h("div", { style: "position:relative;width:26px;height:26px" });
  const lines = wrap === "inline" ? [6, 18] : [2, 7, 12, 17, 22];
  for (const y of lines) box.append(h("div", { style: `position:absolute;left:0;right:0;top:${y}px;height:2px;background:#9e9e9e` }));
  const pic = h("div", { style: `position:absolute;left:8px;top:${wrap === "inline" ? 9 : 7}px;width:10px;height:10px;background:${wrap === "behind" ? "rgba(91,155,213,.5)" : "#5b9bd5"}` });
  if (wrap === "front") pic.style.zIndex = "1";
  box.append(pic);
  return box;
}

const cmd = (id: string, label: string, icon?: string, extra: Partial<Command> = {}): Command => ({ id, label, icon, ...extra });

export const PICTURE_COMMANDS: Command[] = [
  cmd("picture.fromFile", "Dieses Gerät...", "image-add", { keywords: ["Bild einfügen", "Grafik", "Foto"] }),
  cmd("picture.wrapInline", "Mit Text in Zeile", "wrap-text"),
  cmd("picture.wrapSquare", "Quadrat", "wrap-text"),
  cmd("picture.wrapTight", "Eng", "wrap-text"),
  cmd("picture.wrapThrough", "Transparent", "wrap-text"),
  cmd("picture.wrapTopBottom", "Oben und unten", "wrap-text"),
  cmd("picture.wrapBehind", "Hinter den Text", "send-backward"),
  cmd("picture.wrapFront", "Vor den Text", "bring-forward"),
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
  const { tab } = pictureTab(api, insertFromPath);
  const layoutOptions = (rect: DOMRect) => {
    const s0 = sel();
    if (!s0) return;
    const icon = (id: string, wrap: string) => {
      const cur = currentWrap() === wrap;
      return h("button", { class: `tile${cur ? " on" : ""}`, "data-tip": c.get(id).label, style: "width:40px;height:40px", onclick: () => { closePopup(); void c.run(id); } }, wrapIcon(wrap));
    };
    const panel = h("div", { class: "rpanel layout-options", style: "width:220px" },
      h("div", { class: "title" }, "Layoutoptionen"),
      h("div", { class: "subtitle" }, "Mit Text in Zeile"), h("div", { class: "tiles" }, icon("picture.wrapInline", "inline")),
      h("div", { class: "subtitle" }, "Mit Textumbruch"),
      h("div", { class: "tiles", style: "grid-template-columns:repeat(4, auto)" }, ...wraps.slice(1).map(([id, w]) => icon(id, w))),
      h("hr"),
      h("label", { class: "check" }, h("input", { type: "radio", name: "lo-move", checked: s0.kind === "float" && s0.obj.moveWithText !== false, disabled: s0.kind !== "float", onchange: () => { ed()?.setPictureMoveWithText(true); } }), "Mit Text verschieben"),
      h("label", { class: "check" }, h("input", { type: "radio", name: "lo-move", checked: s0.kind === "float" && s0.obj.moveWithText === false, disabled: s0.kind !== "float", onchange: () => { ed()?.setPictureMoveWithText(false); } }), "Position auf der Seite fixieren"),
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

  return { tab, insertMenu: [{ cmd: "picture.fromFile" }] as MenuEntry[], insertFromPath, attach };
}
