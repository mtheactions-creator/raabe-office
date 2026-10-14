// Entwurf > Seitenhintergrund: Wasserzeichen (Word's gallery, recorded 2026-10-14: its
// building blocks are WordArt of fixed proportions, diagonal ones with width + height 659.85 pt,
// horizontal ones 468 pt wide; the dialog Gedrucktes Wasserzeichen) and Seitenfarbe (theme and
// standard colours, Keine Farbe, Weitere Farben). Seitenränder: see borders.ts.

import { colorGrid, h, icon, showFontColors } from "@raabe/shell";
import { imageInfo, naturalSizePt, type Editor, type Watermark } from "@raabe/text-engine";
import type { AppApi } from "./api";
import { openModal } from "./modal";

const GALLERY: { group: string; items: { name: string; text: string; diagonal: boolean; aspect: number }[] }[] = [
  { group: "Eilt", items: [
    { name: "EILT 1", text: "EILT", diagonal: true, aspect: 2 }, { name: "EILT 2", text: "EILT", diagonal: false, aspect: 2 },
    { name: "SFWM 1", text: "SFWM", diagonal: true, aspect: 4 / 3 }, { name: "SFWM 2", text: "SFWM", diagonal: false, aspect: 4 / 3 }] },
  { group: "Haftungsausschlüsse", items: [
    { name: "BEISPIEL 1", text: "BEISPIEL", diagonal: true, aspect: 7 / 3 }, { name: "BEISPIEL 2", text: "BEISPIEL", diagonal: false, aspect: 7 / 3 },
    { name: "ENTWURF 1", text: "ENTWURF", diagonal: true, aspect: 5 / 3 }, { name: "ENTWURF 2", text: "ENTWURF", diagonal: false, aspect: 5 / 3 }] },
  { group: "Vertraulich", items: [
    { name: "NICHT KOPIEREN 1", text: "NICHT KOPIEREN", diagonal: true, aspect: 3 }, { name: "NICHT KOPIEREN 2", text: "NICHT KOPIEREN", diagonal: false, aspect: 3 },
    { name: "VERTRAULICH 1", text: "VERTRAULICH", diagonal: true, aspect: 4 }, { name: "VERTRAULICH 2", text: "VERTRAULICH", diagonal: false, aspect: 4 }] },
];
const TEXTS = ["ASAP", "BEISPIEL", "EILT", "ENTWURF", "KOPIE", "NICHT KOPIEREN", "ORIGINAL", "PERSÖNLICH", "STRENG VERTRAULICH", "VERTRAULICH"];
const SIZES = ["Auto", "36", "40", "44", "48", "54", "60", "66", "72", "80", "90", "96", "105", "120", "144"];
const SCALES: [string, number | undefined][] = [["Auto", undefined], ["500%", 5], ["200%", 2], ["150%", 1.5], ["100%", 1], ["50%", 0.5]];

const r2 = (v: number) => Math.round(v * 100) / 100;

/** Word's building block size for a gallery entry. */
function gallerySize(diagonal: boolean, aspect: number): { width: number; height: number } {
  return diagonal ? { width: r2((659.85 * aspect) / (aspect + 1)), height: r2(659.85 / (aspect + 1)) } : { width: 468, height: r2(468 / aspect) };
}

/** A custom text watermark's size: Auto fits the text area (diagonal: along its diagonal), else the text at that size. */
function textSize(text: string, font: string, size: number | undefined, diagonal: boolean, textWidth: number): { width: number; height: number } {
  const g = document.createElement("canvas").getContext("2d")!;
  g.font = `100px "${font}"`;
  const m = g.measureText(text || " ");
  const w = (m.actualBoundingBoxRight || m.width) + (m.actualBoundingBoxLeft || 0), hgt = (m.actualBoundingBoxAscent || 72) + (m.actualBoundingBoxDescent || 0);
  const aspect = w / Math.max(1, hgt);
  if (size) return { width: r2((w / 100) * size), height: r2((hgt / 100) * size) };
  return diagonal ? { width: r2((textWidth * Math.SQRT2 * aspect) / (aspect + 1)), height: r2((textWidth * Math.SQRT2) / (aspect + 1)) } : { width: r2(textWidth), height: r2(textWidth / aspect) };
}

export function backgroundFeature(api: AppApi) {
  const ed = () => api.editor();
  const c = api.commands;
  const has = () => !!ed();
  const apply = (fn: (e: Editor) => void) => { const e = ed(); if (e) { fn(e); api.refresh(); api.focus(); } };
  const textWidth = (e: Editor) => { const p = e.section().page; return p.width - p.marginLeft - p.marginRight; };

  /** A watermark thumbnail: a little page with the grey text, turned when diagonal. */
  const thumb = (text: string, diagonal: boolean) =>
    h("span", { class: "wm-page" }, h("span", { class: `wm-text${diagonal ? " diag" : ""}` }, text));

  const watermarkPanel = (close: () => void) => {
    const panel = h("div", { class: "rpanel designpanel wm" });
    for (const g of GALLERY) {
      panel.append(h("div", { class: "title" }, g.group), h("div", { class: "wm-grid" }, ...g.items.map((it) => h("button", {
        class: "wm-tile", "data-tip": it.name, "aria-label": it.name,
        onclick: () => {
          close();
          apply((e) => e.setWatermark({ kind: "text", text: it.text, font: "Calibri", color: "C0C0C0", semi: true, diagonal: it.diagonal, ...gallerySize(it.diagonal, it.aspect) }));
        },
      }, thumb(it.text, it.diagonal), h("span", { class: "dz-name" }, it.name)))));
    }
    const cmd = (label: string, ic: string | null, run: () => void, disabled = false, tip?: string) =>
      h("button", { class: "dz-cmd", disabled, ...(tip ? { "data-tip": tip } : {}), onclick: () => { close(); run(); } }, ic ? icon(ic) : h("span", { class: "dz-noicon" }), label);
    panel.append(h("hr"),
      cmd("Weitere Wasserzeichen von Office.com", "globe", () => {}, true, "Weitere Wasserzeichen von Office.com\nRaabeWrite arbeitet ohne Internet."),
      cmd("Benutzerdefiniertes Wasserzeichen...", "watermark", () => watermarkDialog()),
      cmd("Wasserzeichen entfernen", "close", () => apply((e) => e.setWatermark(undefined)), !ed()?.doc.settings.watermark),
      cmd("Auswahl im Wasserzeichenkatalog speichern...", "save", () => {}, true, "Auswahl im Wasserzeichenkatalog speichern\nRaabeWrite hat keinen eigenen Katalog."));
    return panel;
  };

  // ----- Gedrucktes Wasserzeichen ---------------------------------------------------------------
  const watermarkDialog = () => {
    const e = ed();
    if (!e) return;
    const cur = e.doc.settings.watermark;
    const radio = (checked: boolean) => h("input", { type: "radio", name: "wmkind", checked }) as HTMLInputElement;
    const none = radio(!cur), pic = radio(cur?.kind === "picture"), txt = radio(cur?.kind === "text");
    let image = cur?.kind === "picture" ? cur.image : undefined;
    const choose = h("button", {}, "Bild auswählen...") as HTMLButtonElement;
    const picName = h("span", { class: "muted" }, image ? "Bild ausgewählt" : "");
    const scale = h("select", {}, ...SCALES.map(([l]) => h("option", { value: l }, l))) as HTMLSelectElement;
    if (cur?.kind === "picture" && cur.scale) scale.value = SCALES.find(([, v]) => v === cur.scale)?.[0] ?? "Auto";
    const washout = h("input", { type: "checkbox", checked: cur?.kind === "picture" ? cur.washout : true }) as HTMLInputElement;
    const lang = h("select", {}, h("option", null, "Deutsch (Deutschland)"), h("option", null, "Englisch (USA)")) as HTMLSelectElement;
    const text = h("input", { type: "text", list: "wm-texts", value: cur?.kind === "text" ? cur.text : "ENTWURF", style: "width:100%" }) as HTMLInputElement;
    const texts = h("datalist", { id: "wm-texts" }, ...TEXTS.map((t) => h("option", { value: t })));
    const font = h("input", { type: "text", list: "wm-fonts", value: cur?.kind === "text" ? cur.font : "Calibri", style: "width:100%" }) as HTMLInputElement;
    const fonts = h("datalist", { id: "wm-fonts" }, ...api.fontList().map((f) => h("option", { value: f })));
    const size = h("select", {}, ...SIZES.map((s) => h("option", { value: s }, s))) as HTMLSelectElement;
    size.value = cur?.kind === "text" && cur.size ? String(cur.size) : "Auto";
    let color = cur?.kind === "text" ? cur.color : "C0C0C0";
    const colorBtn = h("button", { class: "wm-color", "aria-label": "Farbe" }) as HTMLButtonElement;
    const paint = () => { colorBtn.style.background = `#${color}`; };
    colorBtn.onclick = () => showFontColors(colorBtn.getBoundingClientRect(), (hex) => { color = hex === "auto" ? "000000" : hex; paint(); });
    paint();
    const semi = h("input", { type: "checkbox", checked: cur?.kind === "text" ? cur.semi : true }) as HTMLInputElement;
    const diag = h("input", { type: "radio", name: "wmlayout", checked: cur?.kind === "text" ? cur.diagonal : true }) as HTMLInputElement;
    const horiz = h("input", { type: "radio", name: "wmlayout", checked: cur?.kind === "text" ? !cur.diagonal : false }) as HTMLInputElement;
    const picFields = [choose, scale, washout], txtFields = [lang, text, font, size, colorBtn, semi, diag, horiz];
    const sync = () => {
      for (const f of picFields) f.disabled = !pic.checked;
      for (const f of txtFields) f.disabled = !txt.checked;
    };
    for (const r of [none, pic, txt]) r.onchange = sync;
    sync();
    choose.onclick = async () => {
      const path = await api.platform.openImageDialog().catch(() => null);
      if (!path) return;
      const bytes = await api.platform.readFile(path).catch(() => null);
      const info = bytes ? imageInfo(bytes) : null;
      if (!bytes || !info) return;
      image = { bytes, type: info.type };
      picName.textContent = path.split(/[\\/]/).pop() ?? "Bild ausgewählt";
    };
    const build = (): Watermark | undefined | null => {
      if (none.checked) return undefined;
      if (pic.checked) {
        if (!image) return null;
        const info = imageInfo(image.bytes);
        const nat = info ? naturalSizePt(info) : { w: 200, h: 200 };
        const s = SCALES.find(([l]) => l === scale.value)?.[1];
        // Auto: as wide as the text area (not enlarged beyond 160 %).
        const k = s ?? Math.min(textWidth(e) / Math.max(1, nat.w), 1.6);
        return { kind: "picture", image, ...(s ? { scale: s } : {}), washout: washout.checked, width: r2(nat.w * k), height: r2(nat.h * k) };
      }
      const sz = size.value === "Auto" ? undefined : Number(size.value);
      return { kind: "text", text: text.value || "ENTWURF", font: font.value || "Calibri", ...(sz ? { size: sz } : {}), color, semi: semi.checked, diagonal: diag.checked,
        ...textSize(text.value || "ENTWURF", font.value || "Calibri", sz, diag.checked, textWidth(e)) };
    };
    const commit = () => { const wm = build(); if (wm === null) return false; e.setWatermark(wm); api.refresh(); return true; };
    openModal("Gedrucktes Wasserzeichen", h("div", { class: "wmdlg" },
      h("label", { class: "check" }, none, "Kein Wasserzeichen"),
      h("label", { class: "check" }, pic, "Bildwasserzeichen"),
      h("div", { class: "wm-sub" }, h("div", { class: "row" }, choose, picName), h("div", { class: "row" }, h("span", null, "Skalieren:"), scale, h("label", { class: "check" }, washout, "Auswaschen"))),
      h("label", { class: "check" }, txt, "Textwasserzeichen"),
      h("div", { class: "wm-sub grid2" }, "Sprache:", lang, "Text:", h("div", null, text, texts), "Schriftart:", h("div", null, font, fonts), "Größe:", size,
        "Farbe:", h("div", { class: "row" }, colorBtn, h("label", { class: "check" }, semi, "Halbtransparent")),
        "Layout:", h("div", { class: "row" }, h("label", { class: "check" }, diag, "Diagonal"), h("label", { class: "check" }, horiz, "Horizontal")))), [
      { label: "Übernehmen", run: () => { commit(); } },
      { label: "OK", primary: true, run: () => commit() },
      { label: "Abbrechen", run: () => true },
    ], { width: 520 });
  };

  // ----- Seitenfarbe --------------------------------------------------------------------------------
  const pageColorPanel = (close: () => void) => {
    const e = ed();
    const pick = (hex: string, ref?: string) => { close(); apply((x) => x.setBackground({ color: hex.toUpperCase(), ...(ref ? { colorTheme: ref } : {}) })); };
    const custom = h("input", { type: "color", style: "position:absolute;opacity:0;width:0;height:0" }) as HTMLInputElement;
    custom.addEventListener("change", () => pick(custom.value.slice(1)));
    return h("div", { class: "menu palette" }, ...colorGrid(pick, e?.theme.colors), h("hr"),
      h("button", { onclick: () => { close(); apply((x) => x.setBackground(undefined)); } }, h("span", { class: "auto-swatch none" }), "Keine Farbe"),
      h("button", { onclick: () => custom.click() }, "Weitere Farben..."), custom,
      h("button", { disabled: true, "data-tip": "Fülleffekte\nRaabeWrite füllt Seiten nur mit einer Farbe." }, "Fülleffekte..."));
  };

  c.implement("design.watermark", { run: () => {}, enabled: has });
  c.implement("design.pageColor", { run: () => {}, enabled: has });
  return { watermarkPanel, pageColorPanel, watermarkDialog };
}
