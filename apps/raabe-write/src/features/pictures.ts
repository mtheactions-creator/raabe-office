// Bilder: Einfügen > Bilder > Dieses Gerät, and the contextual tab "Bildformat" with
// Textumbruch and Größe. Datum und Uhrzeit lives here too (Einfügen > Text).

import { h, showDialog, messageBox, baseName, type Command, type Control, type MenuEntry, type Tab } from "@raabe/shell";
import { imageInfo, naturalSizePt } from "@raabe/text-engine";
import { CM, fmtCm, parseNum, type AppApi } from "./api";

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
    if (!path) return;
    try {
      const bytes = await api.platform.readFile(path);
      const ext = baseName(path).split(".").pop()?.toLowerCase() ?? "";
      const type = MIME[ext] ?? "image/png";
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

  // Größe: changing one side keeps the proportions (Seitenverhältnis sperren, Word's default).
  const size = () => { const s = sel(); return s ? { w: s.obj.width ?? 0, h: s.obj.height ?? 0 } : null; };
  const setSide = (side: "w" | "h", pt: number) => {
    const s = size();
    if (!s || !(pt > 1) || !s.w || !s.h) return;
    const k = side === "w" ? pt / s.w : pt / s.h;
    ed()?.setImageSize(s.w * k, s.h * k);
  };
  c.implement("picture.height", { run: () => {}, enabled: picked });
  c.implement("picture.width", { run: () => {}, enabled: picked });
  const spinner = (id: string, label: string, key: string, side: "w" | "h"): Control => ({
    type: "spinner", cmd: id, label, key,
    value: () => { const s = size(); return s ? fmtCm(side === "w" ? s.w : s.h) : ""; },
    commit: (v) => setSide(side, parseNum(v) * CM),
    step: (d) => { const s = size(); if (s) setSide(side, (side === "w" ? s.w : s.h) + d * 0.1 * CM); },
  });

  const wrapMenu: MenuEntry[] = wraps.map(([id]) => ({ cmd: id }));
  const tab: Tab = {
    id: "pictureFormat", label: "Bildformat", key: "JP", contextual: true, shrink: ["size", "arrange"],
    groups: [
      { id: "arrange", label: "Anordnen", key: "ZA", icon: "wrap-text",
        controls: [{ type: "dropdown", cmd: "layout.wrapText", key: "TW", size: "large", menu: wrapMenu }] },
      { id: "size", label: "Größe", key: "ZG", icon: "image-size",
        controls: [{ type: "stack", items: [spinner("picture.height", "Höhe:", "H", "h"), spinner("picture.width", "Breite:", "W", "w")] }] },
    ],
  };

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

  return { tab, insertMenu: [{ cmd: "picture.fromFile" }] as MenuEntry[] };
}
