// Ansicht > Ansichten and Zoom: Drucklayout, Weblayout (the text as wide as the window, no
// pages), Lesemodus (Word's reading view, recorded 2026-10-14: no ribbon, Datei / Extras /
// Ansicht in the title bar, the text in screen-sized columns turned with the arrows, "Bildschirm
// 1 von 6"; Esc returns) and Mehrere Seiten (two pages side by side).

import { h, icon, showMenu, type MenuEntry } from "@raabe/shell";
import type { AppApi } from "./api";

export interface ViewsHost {
  titleBar(): HTMLElement;
  readonly docArea: HTMLElement;
  openBackstage(): void;
  setZoom(z: number): void;
}

export function viewsFeature(api: AppApi, host: ViewsHost) {
  const ed = () => api.editor();
  const c = api.commands;
  let zoomBefore = 1;
  let pageColor: "none" | "sepia" | "invert" = "none";

  const menu = (label: string, entries: () => MenuEntry[]) => {
    const b = h("button", { class: "readmenu" }, label);
    b.onclick = () => showMenu(c, entries(), b.getBoundingClientRect());
    return b;
  };
  const readBar = h("div", { class: "readbar" },
    menu("Datei", () => [{ label: "Datei", run: () => { leave(); host.openBackstage(); } }]),
    menu("Extras", () => [
      { label: "Suchen", icon: "search", run: () => c.run("edit.find") },
      { label: "Intelligente Suche", run: () => {}, disabled: true },
    ]),
    menu("Ansicht", () => {
      const e = ed();
      const o = e?.readOptions ?? { column: "standard", paper: false };
      return [
        { label: "Dokument bearbeiten", run: () => leave() },
        "-",
        { cmd: "view.navPane", label: "Navigationsbereich" },
        { cmd: "review.showComments", label: "Kommentare anzeigen" },
        { label: "Spaltenbreite", submenu: (["narrow", "standard", "wide"] as const).map((k) => ({
          label: { narrow: "Schmal", standard: "Standard", wide: "Breit" }[k], get checked() { return o.column === k; },
          disabled: o.paper, run: () => { ed()?.setReadOptions({ column: k }); api.refresh(); },
        })) },
        { label: "Seitenfarbe", submenu: (["none", "sepia", "invert"] as const).map((k) => ({
          label: { none: "Keine", sepia: "Sepia", invert: "Umkehren" }[k], get checked() { return pageColor === k; },
          run: () => setPageColor(k),
        })) },
        { label: "Layout", submenu: [
          { label: "Spaltenlayout", get checked() { return !o.paper; }, run: () => { ed()?.setReadOptions({ paper: false }); api.refresh(); } },
          { label: "Papierlayout", get checked() { return o.paper; }, run: () => { ed()?.setReadOptions({ paper: true }); api.refresh(); } },
        ] },
        { cmd: "review.readAloud" },
      ];
    }));
  const readTitle = h("span", { class: "readtitle" });
  readBar.append(readTitle);
  const prev = h("button", { class: "readarrow prev", "aria-label": "Vorheriger Bildschirm", "data-tip": "Vorheriger Bildschirm", onclick: () => turn(-1) }, icon("chevron-left"));
  const next = h("button", { class: "readarrow next", "aria-label": "Nächster Bildschirm", "data-tip": "Nächster Bildschirm", onclick: () => turn(1) }, icon("chevron-right"));

  const setPageColor = (k: typeof pageColor) => {
    pageColor = k;
    host.docArea.classList.toggle("read-sepia", k === "sepia");
    host.docArea.classList.toggle("read-invert", k === "invert");
  };

  const turn = (dir: 1 | -1) => { ed()?.turnScreen(dir); api.refresh(); };

  const enter = () => {
    const e = ed();
    if (!e || e.viewMode === "read") return;
    zoomBefore = e.zoom;
    document.body.classList.add("readmode");
    const bar = host.titleBar();
    if (!readBar.isConnected) bar.insertBefore(readBar, bar.children[1] ?? null);
    readTitle.textContent = document.title;
    if (!prev.isConnected) host.docArea.append(prev, next);
    // Word reads at a larger size than it edits.
    host.setZoom(Math.max(zoomBefore, 1.2));
    e.setViewMode("read");
    setPageColor(pageColor);
    api.refresh();
    api.focus();
  };

  const leave = (mode: "print" | "web" = "print") => {
    const e = ed();
    document.body.classList.remove("readmode");
    readBar.remove();
    prev.remove();
    next.remove();
    host.docArea.classList.remove("read-sepia", "read-invert");
    if (e?.viewMode === "read") {
      e.setViewMode(mode);
      host.setZoom(zoomBefore);
    } else e?.setViewMode(mode);
    api.refresh();
    api.focus();
  };

  const mode = () => ed()?.viewMode ?? "print";
  c.implement("view.readMode", { run: () => (mode() === "read" ? leave() : enter()), enabled: () => !!ed(), checked: () => mode() === "read" });
  c.implement("view.printLayout", { run: () => leave("print"), enabled: () => !!ed(), checked: () => mode() === "print" });
  c.implement("view.webLayout", { run: () => leave("web"), enabled: () => !!ed(), checked: () => mode() === "web" });
  c.implement("view.multiplePages", {
    run: () => { const e = ed(); if (!e) return; if (e.viewMode !== "print") leave("print"); host.setZoom(e.view.multiplePagesZoom(2)); },
    enabled: () => !!ed(),
  });

  /** The status bar's first cell in the Lesemodus ("Bildschirm 1 von 6"). */
  const statusLabel = (): string | undefined => {
    const e = ed();
    if (!e || e.viewMode !== "read") return undefined;
    const s = e.readScreens();
    const per = Math.max(1, s.perScreen);
    return `Bildschirm ${Math.floor(s.page / per) + 1} von ${Math.max(1, Math.ceil(s.count / per))}`;
  };

  return { leave, statusLabel, isReading: () => mode() === "read" };
}
