// Datei > Exportieren: Video erstellen and Animiertes GIF erstellen (D-074), with PowerPoint's choices (labels
// as PowerPoint's, the descriptions our own). The work runs with a progress bar and Abbrechen.

import { h, showDialog } from "@raabe/shell";
import type { PresentApp } from "./app";
import { makeGif, makeVideo, videoSupported, type MediaOptions } from "./video";
import { readTransition } from "./transitions";

const num = (v: string) => Number(v.replace(",", "."));
const fmt = (n: number) => n.toFixed(2).replace(".", ",").padStart(5, "0");

/** Whether any slide has a recorded time (Nach), so that "Aufgezeichnete Anzeigedauern" is the default. */
function hasTimings(app: PresentApp): boolean {
  return app.deck.slides.some((s) => readTransition(app.deck.pptx.doc(s.path).documentElement)?.advTm !== undefined);
}

/** Runs an export with a progress bar; returns the bytes or null (stopped, failed). */
async function withProgress(title: string, run: (progress: (d: number) => boolean) => Promise<Uint8Array | null>): Promise<Uint8Array | null> {
  let stop = false;
  const bar = h("div", { class: "export-bar-fill" });
  const label = h("span", null, `${title}: 0 %`);
  const box = h("div", { class: "export-progress", role: "progressbar" }, label, h("div", { class: "export-bar" }, bar),
    h("button", { class: "btn", onclick: () => { stop = true; } }, "Abbrechen"));
  document.body.append(box);
  let last = 0;
  try {
    return await run((d) => {
      const now = performance.now();
      if (now - last > 120) { last = now; bar.style.width = `${Math.round(d * 100)}%`; label.textContent = `${title}: ${Math.round(d * 100)} %`; }
      return !stop;
    });
  } finally { box.remove(); }
}

async function save(app: PresentApp, ext: "mp4" | "gif", bytes: Uint8Array): Promise<void> {
  const filters = ext === "mp4" ? [{ name: "MPEG-4-Video", extensions: ["mp4"] }] : [{ name: "GIF (Graphics Interchange Format)", extensions: ["gif"] }];
  const target = app.platform.saveAnyDialog ? await app.platform.saveAnyDialog("Speichern unter", `${app.docName()}.${ext}`, filters) : null;
  if (target) await app.platform.writeFile(target, bytes);
}

/** The Video erstellen section. */
export function videoSection(app: PresentApp): HTMLElement {
  const quality = h("select", { "aria-label": "Videoqualität" },
    ...[["2160", "Ultra HD (4K)"], ["1080", "Full HD (1080p)"], ["720", "HD (720p)"], ["480", "Standard (480p)"]].map(([v, l]) => h("option", { value: v, selected: v === "1080" }, l)));
  const timed = hasTimings(app);
  const timings = h("select", { "aria-label": "Zeitabläufe und Kommentare" },
    h("option", { value: "0", selected: !timed }, "Aufgezeichnete Zeitabläufe und Kommentare nicht verwenden"),
    h("option", { value: "1", selected: timed, disabled: !timed }, "Aufgezeichnete Zeitabläufe und Kommentare verwenden"));
  const secs = h("input", { type: "text", value: fmt(5), style: "width:64px", "aria-label": "Anzeigedauer der einzelnen Folien in Sekunden" });
  const go = async () => {
    if (!videoSupported()) { await showDialog({ title: "RaabePresent", body: "Videos kann RaabePresent auf diesem Gerät nicht erstellen.", buttons: [{ label: "OK", value: "", primary: true }], cancel: "" }); return; }
    const o: MediaOptions = { height: Number(quality.value), fps: 30, secs: Math.max(0.01, num(secs.value) || 5), useTimings: timings.value === "1" };
    const bytes = await withProgress("Video wird erstellt", (progress) => makeVideo(app, { ...o, progress }));
    if (bytes) await save(app, "mp4", bytes);
  };
  return h("div", { class: "export-section" }, h("h2", null, "Video erstellen"),
    h("p", { class: "muted" }, "Die Präsentation als Videodatei mit Übergängen, Animationen, Videos und Klängen, zum Abspielen ohne RaabePresent."),
    h("div", { class: "export-row" }, quality), h("div", { class: "export-row" }, timings),
    h("div", { class: "export-row" }, h("label", null, "Anzeigedauer der einzelnen Folien in Sekunden: "), secs),
    h("button", { class: "bigbtn", style: "width:160px", onclick: () => void go() }, "Video erstellen"));
}

/** The Animiertes GIF erstellen section. */
export function gifSection(app: PresentApp): HTMLElement {
  const n = app.deck.slides.length;
  const quality = h("select", { "aria-label": "GIF-Qualität" },
    ...[["1080", "Sehr groß"], ["720", "Groß"], ["480", "Mittel"], ["240", "Klein"]].map(([v, l]) => h("option", { value: v, selected: v === "480" }, l)));
  const transparent = h("input", { type: "checkbox", id: "gif-transparent" });
  const secs = h("input", { type: "text", value: fmt(1), style: "width:64px", "aria-label": "Wieviel Sekunden für jede Folie benötigt werden" });
  const from = h("input", { type: "number", min: "1", max: String(n), value: "1", style: "width:56px", "aria-label": "Erste Folie im Exportbereich" });
  const to = h("input", { type: "number", min: "1", max: String(n), value: String(n), style: "width:56px", "aria-label": "Letzte Folie im Exportbereich" });
  const go = async () => {
    const a = Math.max(1, Math.min(n, Number(from.value) || 1)), b = Math.max(a, Math.min(n, Number(to.value) || n));
    // PowerPoint's GIF: about 16 frames per second, and a slide's Nach time is kept (measured).
    const o: MediaOptions = { height: Number(quality.value), fps: 16, secs: Math.max(0.01, num(secs.value) || 1), useTimings: true, from: a, to: b, transparent: transparent.checked, roundWidth: true, gifTiming: true };
    const bytes = await withProgress("GIF wird erstellt", (progress) => makeGif(app, { ...o, progress }));
    if (bytes) await save(app, "gif", bytes);
  };
  return h("div", { class: "export-section" }, h("h2", null, "Animierte GIF-Datei erstellen"),
    h("p", { class: "muted" }, "Die Folien als GIF-Animation, die sich in Webseiten, Chats und andere Dokumente einfügen lässt."),
    h("div", { class: "export-row" }, quality),
    h("div", { class: "export-row" }, transparent, h("label", { for: "gif-transparent" }, " Transparenter Hintergrund")),
    h("div", { class: "export-row" }, h("label", null, "Wieviel Sekunden für jede Folie benötigt werden: "), secs),
    h("div", { class: "export-row" }, h("label", null, "Folien "), from, h("span", null, " bis "), to),
    h("button", { class: "bigbtn", style: "width:160px", onclick: () => void go() }, "GIF erstellen"));
}
