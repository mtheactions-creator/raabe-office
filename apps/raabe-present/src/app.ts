// RaabePresent: wires the shell (title bar, ribbon, backstage) to presentations (D-070). Each open
// presentation is a document tab (D-010) shown in Normalansicht: the slide panel at the left, the
// slide in the middle, the notes below. Files are .pptx; everything not edited is kept (D-024).

import { colorXml, specHex } from "./colors";
import { colorWord } from "./quickstyles";
import { Pptx, newPptx, PptxDom, phOf, type ShapeView, type DesignSpec } from "@raabe/ooxml";
import {
  Backstage, Commands, Keytips, Ribbon, TitleBar, baseName, colorGrid, HIGHLIGHTS, h, icon, installTooltips, messageBox, showDialog, showMenu, showPopup, closePopup,
  type BackstageEntry, type Control, type Group, type MenuEntry, type Platform, type Tab,
} from "@raabe/shell";
import { SHAPE_GALLERY, shapeIcon, type GalleryEntry } from "@raabe/text-engine";
import { COMMANDS } from "./command-list";
import { ribbonTabs, type RibbonContext } from "./ribbon-def";
import { Deck, defaultMasterPlaceholder } from "./deck";
import { LINE_DASHES, LINE_WEIGHTS } from "./lines";
import { quickStylesPanel } from "./quickstyles";
import { wordArtInsertPanel } from "./wordart";
import { setShowFlag, showFlag } from "./showsetup";
import { editEquation, equationTab, hasEquation, implementEquationCommands, insertEquation, selectedEquation } from "./equations";
import { BUILT_IN } from "./equations-builtin";
import { SlideEditor, flat, union, type EditorHost } from "./editor";
import { SlidePanel, type PanelHost } from "./panel";
import { renderSlide } from "./render/slide";
import { EMU_PX } from "./render/text";
import { DEFAULT_TABLE_STYLE_ID, DEFAULT_TABLE_STYLE_XML } from "./render/table";
import * as T from "./textops";
import * as S from "./shapeops";
import { DESIGNS } from "./designs";
import { PRESENT_TEMPLATES, templateDeck } from "./templates";
import { TRANSITION_GROUPS } from "./transitions";
import { ANIMATION_GROUPS } from "./animations";
import { COLOR_SETS, FONT_SETS } from "./themesets.gen";
import { SlideShow } from "./show";
import { FormatPane } from "./formatpane";
import { NotesPageView, OutlineView } from "./views";
import { FindDialog } from "./find";
import { AnimationPane, effects } from "./animpane";
import { Comments } from "./comments";
import { BackgroundPane } from "./background";
import * as CT from "./contexttabs";
import { contextHost, implementContext } from "./contextcmds";
import { Sorter } from "./sorter";
import { zoomMenu } from "./zoom";
import { bgStylesMenu } from "./background";
import { shapeEffectsMenu } from "./shapefx";
import { addRecent, loadRecent, loadSettings, removeRecent, saveSettings, setPinned, type Settings } from "./store";
import appIconUrl from "../../../branding/raabe-present-small.svg?url";

const { kid, kids, at, num, make, fragment } = PptxDom;

interface Doc {
  id: string;
  name: string;
  path: string | null;
  deck: Deck;
  current: number;
  savedVersion: number;
  recoveryVersion: number;
  lastSaved: number | null;
}

type View = "normal" | "sorter" | "notesPage" | "reading" | "outline";

/** Interaktive Schaltflächen: PowerPoint's names and the action each gets when it is drawn (D-072). */
export const ACTION_BUTTONS: [string, string, string | null][] = [
  ["actionButtonBackPrevious", "Interaktive Schaltfläche: Zurück oder Vorherige(r)", "previousslide"], ["actionButtonForwardNext", "Interaktive Schaltfläche: Nächste(r) oder Weiter", "nextslide"],
  ["actionButtonBeginning", "Interaktive Schaltfläche: Zum Anfang wechseln", "firstslide"], ["actionButtonEnd", "Interaktive Schaltfläche: Zum Ende wechseln", "lastslide"],
  ["actionButtonHome", "Interaktive Schaltfläche: Zur Startseite wechseln", "firstslide"], ["actionButtonInformation", "Interaktive Schaltfläche: Informationen abrufen", null],
  ["actionButtonReturn", "Interaktive Schaltfläche: Zurückkehren", "lastslideviewed"], ["actionButtonMovie", "Interaktive Schaltfläche: Video", null],
  ["actionButtonDocument", "Interaktive Schaltfläche: Dokument", null], ["actionButtonSound", "Interaktive Schaltfläche: Sound", null],
  ["actionButtonHelp", "Interaktive Schaltfläche: Hilfe", null], ["actionButtonBlank", "Interaktive Schaltfläche: Leer", null],
];
/** Office's Formen gallery and PowerPoint's own last group. */
const PRESENT_SHAPES: { title: string; items: GalleryEntry[] }[] = [...SHAPE_GALLERY, { title: "Interaktive Schaltflächen", items: ACTION_BUTTONS.map(([geometry, label]): GalleryEntry => ({ tool: { geometry, kind: "shape" }, label })) }];

/** PowerPoint's names for three of the highlight colours (Word's are the shell's). */
export const PP_HIGHLIGHT_NAMES: Record<string, string> = { green: "Grelles Grün", darkGray: "Grau-50 %", lightGray: "Grau-25 %" };

const stripExt = (n: string) => n.replace(/\.(pptx|pptm|ppsx|ppsm|potx|potm)$/i, "");

/** The main part's content type per file type (PowerPoint's: with macros, Bildschirmpräsentation, Vorlage). */
export const PRES_TYPES: Record<string, string> = {
  pptx: "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml",
  pptm: "application/vnd.ms-powerpoint.presentation.macroEnabled.main+xml",
  ppsx: "application/vnd.openxmlformats-officedocument.presentationml.slideshow.main+xml",
  ppsm: "application/vnd.ms-powerpoint.slideshow.macroEnabled.main+xml",
  potx: "application/vnd.openxmlformats-officedocument.presentationml.template.main+xml",
  potm: "application/vnd.ms-powerpoint.template.macroEnabled.main+xml",
};
const RT_VBA = "http://schemas.microsoft.com/office/2006/relationships/vbaProject";
const OPEN_FILTERS = [{ name: "Alle PowerPoint-Präsentationen", extensions: ["pptx", "pptm", "ppsx", "ppsm", "potx", "potm"] }];
const SAVE_FILTERS = [{ name: "PowerPoint-Präsentation", extensions: ["pptx"] }, { name: "PowerPoint-Präsentation mit Makros", extensions: ["pptm"] },
  { name: "PowerPoint-Bildschirmpräsentation", extensions: ["ppsx"] }, { name: "PowerPoint-Vorlage", extensions: ["potx"] }];
const samePath = (a: string, b: string) => a.toLowerCase() === b.toLowerCase();
const SIZES = [8, 9, 10, 10.5, 11, 12, 14, 16, 18, 20, 24, 28, 32, 36, 40, 44, 48, 54, 60, 66, 72, 80, 88, 96];

export class PresentApp implements EditorHost, PanelHost {
  readonly commands = new Commands();
  private docs: Doc[] = [];
  private active!: Doc;
  private settings: Settings = loadSettings();
  private untitled = 0;
  private titleBar: TitleBar;
  private ribbon: Ribbon;
  private backstage: Backstage;
  readonly editor: SlideEditor;
  readonly panel: SlidePanel;
  private notes: HTMLTextAreaElement;
  private notesWrap: HTMLElement;
  private status: HTMLElement;
  private statusSlide: HTMLElement;
  private zoomPct: HTMLElement;
  private zoomSlider: HTMLInputElement;
  private main: HTMLElement;
  private center: HTMLElement;
  private sorter: Sorter;
  private outline!: OutlineView;
  private notesPage!: NotesPageView;
  private view: View = "normal";
  private showNotes = true;
  private fontList: string[] = [];
  private imageUrls = new Map<string, { bytes: Uint8Array; url: string }>();
  private clipboard: { xml: string[]; text: string; slide: string; pastes: number } | null = null;
  private slideClipboard: { xml: string; layout: string; text: string; src?: Pptx; path?: string }[] | null = null;
  private painter: { rPr: Element | null; pPr: Element | null; spPr: Element | null; keep: boolean } | null = null;
  private lastFontColor = "FF0000";
  private lastHighlight = "FFFF00";
  private lastFill = "4472C4";
  private lastLine = "000000";
  private show: SlideShow | null = null;
  private formatPane!: FormatPane;
  private animPane!: AnimationPane;
  private comments!: Comments;
  private bgPane!: BackgroundPane;

  constructor(readonly platform: Platform) {
    void platform.listFonts().then((f) => { this.fontList = f; }).catch(() => {});
    this.commands.register(...COMMANDS, ...CT.CONTEXT_COMMANDS);
    this.untitled = 1;
    const first = this.makeDoc(new Deck(), "Präsentation1", null);
    this.docs.push(first);
    this.active = first;
    this.editor = new SlideEditor(this);
    this.panel = new SlidePanel(this);
    this.sorter = new Sorter(this);
    this.implementCommands();
    this.titleBar = new TitleBar(this.commands, platform, {
      appIcon: appIconUrl,
      qat: [{ cmd: "file.save", key: "1" }, { cmd: "edit.undo", key: "2" }, { cmd: "edit.redo", key: "3" }, { cmd: "show.fromStart", key: "4" }],
      onSelectDoc: (id) => this.activate(this.docs.find((d) => d.id === id)!),
      onCloseDoc: (id) => void this.closeDoc(this.docs.find((d) => d.id === id)!),
      onNewDoc: () => this.newDoc(),
      onSearchDone: () => this.focusMain(),
    });
    this.ribbon = new Ribbon(this.commands, this.ribbonLabels(ribbonTabs(this.ribbonContext())), { onFile: () => this.backstage.open("home"), onDone: () => this.focusMain(), tabRowRight: [] });
    this.backstage = new Backstage(this.commands, this.backstageEntries());
    this.backstage.onClose(() => this.focusMain());
    this.notes = document.createElement("textarea");
    this.notes.className = "notes";
    this.notes.placeholder = "Klicken Sie, um Notizen hinzuzufügen";
    this.notes.setAttribute("aria-label", "Notizen");
    this.notes.addEventListener("input", () => {
      const p = this.slidePath();
      if (p) this.deck.setNotes(p, this.notes.value);
    });
    this.notes.addEventListener("contextmenu", (e) => { e.preventDefault(); this.notesMenu(e); });
    this.notesWrap = h("div", { class: "notes-wrap" }, this.notes);
    this.center = h("div", { class: "center" }, this.editor.el, this.notesWrap);
    this.outline = new OutlineView(this);
    this.notesPage = new NotesPageView(this);
    this.main = h("div", { class: "present-main" }, this.panel.el, this.outline.el, h("div", { class: "splitter" }), this.center, this.sorter.el, this.notesPage.el);
    this.formatPane = new FormatPane(this.editor, this.main);
    this.animPane = new AnimationPane(this, this.main);
    this.comments = new Comments(this, this.main, () => this.settings.userName, () => this.settings.initials);
    this.bgPane = new BackgroundPane(this, this.main);
    this.statusSlide = h("span", { class: "cell" });
    this.zoomPct = h("button", { class: "cell pct", "data-tip": "Zoom", onclick: () => void this.commands.run("view.zoom") });
    this.zoomSlider = h("input", { type: "range", min: "0", max: "100", step: "0.5", "aria-label": "Zoom" }) as HTMLInputElement;
    this.zoomSlider.oninput = () => this.setZoom(sliderToZoom(Number(this.zoomSlider.value)));
    const viewBtn = (cmd: string) => { const c = this.commands.get(cmd); const btn = h("button", { class: "cell viewbtn", "aria-label": c.label, "data-tip": this.commands.tooltip(cmd), onclick: () => void this.commands.run(cmd) }, icon(c.icon)); this.commands.onRefresh(() => btn.classList.toggle("checked", this.commands.isChecked(cmd))); return btn; };
    this.status = h("div", { class: "statusbar", role: "status" }, this.statusSlide, h("button", { class: "cell", "data-tip": "Sprache: Deutsch (Deutschland)" }, "Deutsch (Deutschland)"),
      h("div", { class: "right" },
        h("button", { class: "cell", "aria-label": "Notizen", onclick: () => void this.commands.run("view.notes") }, icon("show-notes"), " Notizen"),
        viewBtn("view.normal"), viewBtn("view.sorter"), viewBtn("view.reading"), viewBtn("show.fromCurrent"),
        h("div", { class: "zoom" },
          h("button", { "aria-label": "Verkleinern", onclick: () => this.setZoom(this.editor.zoom - 10) }, "−"), this.zoomSlider,
          h("button", { "aria-label": "Vergrößern", onclick: () => this.setZoom(this.editor.zoom + 10) }, "+")),
        this.zoomPct,
        h("button", { class: "cell", "aria-label": "Folie an aktuelles Fenster anpassen", "data-tip": "Folie an aktuelles Fenster anpassen", onclick: () => this.fitWindow() }, icon("page-width"))));
    document.body.append(this.titleBar.el, this.ribbon.tabRow, this.ribbon.card, this.main, this.status, this.backstage.el);
    document.body.classList.add("present");
    document.body.classList.toggle("touch", this.settings.touchMode);
    this.editor.autoCorrectOn = this.settings.autoCorrect;
    installTooltips();
    new Keytips((el) => (el.dataset.keytipScope === "root" && el.getAttribute("role") === "tab" ? this.ribbon.activeTab : null), () => this.focusMain());
    this.installSplitter();
    addEventListener("keydown", (e) => this.globalKey(e));
    this.editor.el.addEventListener("wheel", (e) => {
      if (!e.ctrlKey) return;
      e.preventDefault();
      // Strg+mouse wheel zooms in steps of 10 % (PowerPoint, to be confirmed by the catalogue).
      this.setZoom(this.editor.zoom + (e.deltaY < 0 ? 10 : -10));
    }, { passive: false });
    platform.window.onCloseRequested(() => this.confirmQuit());
    setInterval(() => void this.autoRecover(), Math.max(10, this.settings.autoRecoverSeconds) * 1000);
    this.activate(first);
    void this.offerRecovery();
    platform.onOpenFiles((paths) => { for (const p of paths) void this.openPath(p); });
  }

  // ----- documents ---------------------------------------------------------------------------------

  get deck(): Deck { return this.active.deck; }

  private makeDoc(deck: Deck, name: string, path: string | null): Doc {
    const doc: Doc = { id: crypto.randomUUID(), name, path, deck, current: 0, savedVersion: 0, recoveryVersion: 0, lastSaved: null };
    deck.onChange = () => { if (doc === this.active) this.afterChange(); else this.updateTabs(); };
    return doc;
  }

  newDoc(design?: DesignSpec): void {
    this.untitled++;
    const doc = this.makeDoc(new Deck(newPptx(design)), `Präsentation${this.untitled}`, null);
    this.docs.push(doc);
    this.activate(doc);
  }

  activate(doc: Doc): void {
    this.editor.endText();
    this.editor.sel = [];
    this.active = doc;
    this.panel.selected = new Set([doc.current]);
    this.panel.syncSelection(doc.current);
    this.editor.fit = true;
    this.renderAll();
    this.editor.fitZoom();
    this.updateTabs();
    this.focusMain();
  }

  private updateTabs(): void {
    this.titleBar.setDocs(this.docs.map((d) => ({ id: d.id, name: d.name, dirty: d.deck.version !== d.savedVersion })), this.active.id);
    document.title = `${this.active.name} - RaabePresent`;
    this.platform.window.setTitle(document.title);
  }

  private afterChange(): void {
    const n = this.deck.slides.length;
    if (this.active.current >= n) this.active.current = Math.max(0, n - 1);
    this.panel.syncSelection(this.active.current);
    this.renderAll();
    this.updateTabs();
  }

  private renderAll(): void {
    this.editor.render();
    this.panel.render();
    if (this.view === "sorter") this.sorter.render();
    this.outline.render();
    this.notesPage.render();
    this.syncNotes();
    this.updateChrome();
  }

  private syncNotes(): void {
    const p = this.slidePath();
    const text = p ? this.deck.notesText(p) : "";
    if (document.activeElement !== this.notes && this.notes.value !== text) this.notes.value = text;
    this.notesWrap.hidden = !this.showNotes;
  }

  updateChrome(): void {
    const n = this.deck.slides.length;
    this.statusSlide.textContent = n ? `Folie ${this.active.current + 1} von ${n}` : "";
    this.zoomPct.textContent = `${this.editor.zoom} %`;
    if (document.activeElement !== this.zoomSlider) this.zoomSlider.value = String(zoomToSlider(this.editor.zoom));
    this.updateContextual();
    this.commands.refresh();
  }

  // ----- EditorHost / PanelHost -----------------------------------------------------------------------

  slidePath(): string | null {
    if (this.masterMode) return this.masterList()[this.masterIndex] ?? null;
    return this.deck.slides[this.active.current]?.path ?? null;
  }

  /** Folienmaster view (Ansicht > Folienmaster): the panel lists the master and its layouts. */
  private masterMode = false;
  private masterIndex = 0;
  private masterKind: "slides" | "notes" | "handout" = "slides";
  /** Handzettelmaster > Folien pro Seite (a view setting, as in PowerPoint not part of the master's shapes). */
  handoutPerPage: 1 | 2 | 3 | 4 | 6 | 9 | 0 = 6;
  private masterList(): string[] {
    if (this.masterKind === "notes") return [this.deck.ensureNotesMasterPath()];
    if (this.masterKind === "handout") return [this.deck.ensureHandoutMasterPath()];
    const m = this.deck.pptx.masters()[0];
    return m ? [m, ...this.deck.pptx.layouts(m)] : [];
  }
  masterPaths(): string[] | null { return this.masterMode ? this.masterList() : null; }

  setMasterView(on: boolean, kind: "slides" | "notes" | "handout" = "slides"): void {
    this.masterKind = kind;
    if (on && kind === "notes") {
      const nm = this.deck.ensureNotesMasterPath();
      if (!this.deck.pptx.has(nm)) this.deck.change("Notizenmaster", [nm], () => this.deck.ensureNotesMaster(nm));
    }
    if (on && kind === "handout") {
      const hm = this.deck.ensureHandoutMasterPath();
      if (!this.deck.pptx.has(hm)) this.deck.change("Handzettelmaster", [hm], () => this.deck.ensureHandoutMaster(hm));
    }
    this.editor.pageDecor = on && kind === "handout" ? (size) => this.handoutDecor(size) : null;
    this.editor.endText();
    this.editor.sel = [];
    this.masterMode = on;
    document.body.classList.toggle("master-view", on);
    if (on) {
      // PowerPoint opens the master view at the current slide's layout.
      const lay = kind === "slides" ? this.deck.pptx.layoutOf(this.deck.slides[this.active.current]?.path ?? "") : undefined;
      const list = this.masterList();
      this.masterIndex = Math.max(0, lay ? list.indexOf(lay) : 0);
      this.panel.selected = new Set([this.masterIndex]);
    } else this.panel.selected = new Set([this.active.current]);
    this.ctxKey = "";
    this.renderAll();
    this.focusMain();
  }
  /** 0 in the master view: layouts show the number field as "‹Nr.›" (PowerPoint). */
  slideNumber(): number { return this.masterMode ? 0 : this.active.current + 1; }
  current(): number { return this.masterMode ? this.masterIndex : this.active.current; }

  setCurrent(i: number): void {
    if (this.masterMode) {
      const n2 = this.masterList().length;
      this.masterIndex = Math.max(0, Math.min(n2 - 1, i));
      this.editor.endText(); this.editor.sel = [];
      this.editor.render(); this.panel.render(); this.updateChrome();
      return;
    }
    const n = this.deck.slides.length;
    i = Math.max(0, Math.min(n - 1, i));
    if (i !== this.active.current) { this.editor.endText(); this.editor.sel = []; }
    this.active.current = i;
    this.editor.render();
    this.panel.render();
    this.syncNotes();
    this.updateChrome();
  }

  image(path: string): string | undefined {
    const bytes = this.deck.pptx.files[path];
    if (!bytes) return undefined;
    const hit = this.imageUrls.get(path);
    if (hit && hit.bytes === bytes) return hit.url;
    const type = mimeOf(path);
    const url = URL.createObjectURL(new Blob([bytes as BlobPart], { type }));
    this.imageUrls.set(path, { bytes, url });
    return url;
  }

  /** Animation übertragen: the shape whose effects the next click takes (once, as PowerPoint's brush). */
  private animBrush: number | null = null;

  selectionChanged(): void {
    if (this.animBrush !== null) {
      const sel = this.editor.sel;
      if (sel.length === 1 && sel[0] !== this.animBrush) {
        const from = this.animBrush;
        this.animBrush = null;
        document.body.classList.remove("painting");
        void import("./animations").then((m) => { m.copyAnimation(this, from, sel[0]); this.animPane?.refresh(); });
      }
    }
    if (this.comments?.pane.isOpen) this.comments.pane.refresh();
    this.formatPane?.sync();
    this.animPane?.refresh();
    this.updateChrome();
  }

  toolDone(): void {
    this.updateChrome();
  }

  /** Markup anzeigen (Überprüfen > Kommentare anzeigen): the comment markers are shown. */
  showMarkup = true;

  commentMarkers(path: string): { x: number; y: number; n: number }[] {
    if (!this.showMarkup) return [];
    return this.comments?.markers(path) ?? [];
  }

  motionPaths(path: string, id: number): string[] {
    const root = this.deck.pptx.doc(path).documentElement;
    const timing = kid(root, "timing");
    if (!timing) return [];
    return Array.from(timing.getElementsByTagNameNS(PptxDom.NS_P, "animMotion"))
      .filter((m) => m.getElementsByTagNameNS(PptxDom.NS_P, "spTgt")[0]?.getAttribute("spid") === String(id))
      .map((m) => m.getAttribute("path") ?? "").filter(Boolean);
  }

  goSlide(key: "PageDown" | "PageUp" | "Home" | "End"): void {
    const i = this.active.current, n = this.deck.slides.length;
    this.setCurrent(key === "PageDown" ? i + 1 : key === "PageUp" ? i - 1 : key === "Home" ? 0 : n - 1);
  }

  equationClick(v: ShapeView, at: { x: number; y: number }): boolean {
    if (!hasEquation(v)) return false;
    this.editor.endText();
    this.editor.select([v.id]);
    void editEquation(this, v.id, false, at);
    return true;
  }

  doubleClick(v: ShapeView): boolean {
    // An equation: its linear box (D-072).
    if (hasEquation(v)) { void editEquation(this, v.id); return true; }
    // Pictures, tables and charts: their contextual tab (PowerPoint); shapes and text enter their text.
    if (v.kind === "pic") { this.editor.select([v.id]); this.ribbon.select("pictureFormat"); return true; }
    if (v.kind === "graphicFrame" && v.frame === "chart") { this.editor.select([v.id]); this.ribbon.select("chartDesign"); return true; }
    if (v.kind === "grpSp") return true;
    return false;
  }

  keys(e: KeyboardEvent): boolean {
    // Keys in the slide panel.
    if (this.globalShortcut(e)) return true;
    if (this.masterMode) return false;
    const ctrl = e.ctrlKey && !e.altKey;
    const sel = [...this.panel.selected].sort((a, b) => a - b);
    if (e.key === "Delete" || e.key === "Backspace") { this.deleteSlides(sel); return true; }
    if (e.key === "Enter") { this.newSlide(); return true; }
    if (ctrl && (e.key === "d" || e.key === "D")) { this.duplicateSlides(sel); return true; }
    if (ctrl && (e.key === "c" || e.key === "C")) { this.copySlides(false); return true; }
    if (ctrl && (e.key === "x" || e.key === "X")) { this.copySlides(true); return true; }
    if (ctrl && (e.key === "v" || e.key === "V")) { this.pasteSlides(); return true; }
    if (ctrl && (e.key === "a" || e.key === "A")) { this.panel.selected = new Set(this.deck.slides.map((_, i) => i)); this.panel.render(); return true; }
    if (ctrl && e.key === "ArrowUp") { this.moveSlidesBy(-1); return true; }
    if (ctrl && e.key === "ArrowDown") { this.moveSlidesBy(1); return true; }
    return false;
  }

  focusSlide(): void { this.editor.focus(); }

  private focusMain(): void {
    if (this.view === "sorter") this.sorter.focus();
    else this.editor.focus();
  }

  // ----- slides ----------------------------------------------------------------------------------------

  newSlide(layout?: string): void {
    const at2 = this.deck.slides.length ? Math.max(...this.panel.selected, this.active.current) : -1;
    this.deck.addSlide(at2, layout);
    this.panel.selected = new Set([at2 + 1]);
    this.setCurrent(at2 + 1);
  }

  private deleteSlides(sel: number[]): void {
    if (!sel.length || !this.deck.slides.length) return;
    const first = Math.min(...sel);
    this.deck.deleteSlides(sel);
    this.active.current = Math.min(first, this.deck.slides.length - 1);
    this.panel.selected = new Set([Math.max(0, this.active.current)]);
    this.afterChange();
  }

  private duplicateSlides(sel: number[]): void {
    if (!sel.length) return;
    const made = this.deck.duplicateSlides(sel);
    const idx = made.map((p) => this.deck.slideIndex(p));
    this.panel.selected = new Set(idx);
    this.setCurrent(idx[0]);
  }

  private moveSlidesBy(d: -1 | 1): void {
    const sel = [...this.panel.selected].sort((a, b) => a - b);
    const to = d < 0 ? Math.max(0, sel[0] - 1) : Math.min(this.deck.slides.length - sel.length, sel[0] + 1);
    if (to === sel[0]) return;
    this.deck.moveSlides(sel, to);
    this.panel.selected = new Set(sel.map((_, k) => to + k));
    this.setCurrent(to + sel.indexOf(this.active.current) >= 0 ? to + Math.max(0, sel.indexOf(this.active.current)) : to);
  }

  private copySlides(cut: boolean): void {
    const sel = [...this.panel.selected].sort((a, b) => a - b);
    this.slideClipboard = sel.map((i) => {
      const p = this.deck.slides[i].path;
      return { xml: PptxDom.serialize(this.deck.pptx.doc(p)), layout: this.deck.pptx.layoutOf(p) ?? "", text: "", src: this.deck.pptx, path: p };
    });
    this.clipboard = null;
    void this.platform.clipboardWrite(sel.map((i) => this.slideText(i)).join("\r\n")).catch(() => {});
    if (cut) this.deleteSlides(sel);
  }

  private pasteSlides(): void {
    if (!this.slideClipboard) return;
    const at2 = this.active.current;
    // Pasted slides come with their pictures, media and charts (slideimport.ts), also from another presentation.
    const clip = this.slideClipboard;
    if (clip.every((s) => s.src && s.path && s.src.has(s.path))) {
      void import("./slideimport").then((m) => { clip.forEach((s, k) => m.importSlide(this.deck, s.src!, s.path!, at2 + k)); this.setCurrent(at2 + 1); });
      return;
    }
    for (const [k, s] of this.slideClipboard.entries()) {
      const layout = this.deck.pptx.has(s.layout) ? s.layout : this.deck.layouts()[1]?.path;
      const path = this.deck.addSlide(at2 + k, layout);
      this.deck.change("Einfügen", [path], () => {
        const doc = PptxDom.parseXml(s.xml);
        // Relationships to parts the target lacks would break: drop shapes that need them.
        for (const el of Array.from(doc.getElementsByTagNameNS(PptxDom.NS_A, "blip"))) el.closest("pic")?.remove();
        this.deck.pptx.put(path, doc);
      });
    }
    this.setCurrent(at2 + 1);
  }

  private slideText(i: number): string {
    const v = this.deck.view(this.deck.slides[i].path);
    return flat(v.shapes).filter((s) => s.text && !s.text.empty).map((s) => s.text!.paras.map((p) => p.runs.map((r) => r.text).join("")).join("\r\n")).join("\r\n");
  }

  // ----- formatting ------------------------------------------------------------------------------------------

  /** Applies a run change to the selection: text being edited (the word at a bare caret), or every selected shape's text. */
  /** formatRuns for the feature modules (Sprache). */
  formatText(label: string, fn: (rPr: Element) => void): void { this.formatRuns(label, fn); }

  private formatRuns(label: string, fn: (rPr: Element) => void): void {
    const ed = this.editor, t = ed.text;
    if (t) {
      const body = ed.textBody(t.id);
      if (!body) return;
      if (T.cmpPos(t.caret, t.anchor) === 0) {
        const p = T.paras(body)[t.caret.p];
        const w = p ? T.wordAt(p, t.caret.o) : { from: 0, to: 0 };
        if (w.from < w.to && w.from < t.caret.o && t.caret.o < w.to) {
          ed.edit(label, () => T.formatRange(body, { p: t.caret.p, o: w.from }, { p: t.caret.p, o: w.to }, fn));
        } else {
          // At the end of a word or in empty text: for the next typed characters (and an empty paragraph's mark).
          const base = (p ? T.formatAt(p, t.caret.o) : null) ?? null;
          const pending = (t.pending ?? (base ? base.cloneNode(true) as Element : make(body.ownerDocument, "a:rPr", { lang: "de-DE" }))) as Element;
          fn(pending);
          t.pending = pending;
          if (p && T.paraLength(p) === 0) ed.edit(label, () => T.formatRange(body, t.caret, t.caret, fn, true));
          else this.updateChrome();
        }
        return;
      }
      ed.edit(label, () => T.formatRange(body, t.anchor, t.caret, fn));
      return;
    }
    const views = ed.selectedViews;
    if (!views.length) return;
    ed.edit(label, () => {
      for (const v of views) for (const m of flat([v])) {
        const el = ed.shapeEl(m.id);
        const body = el ? S.txBodyOf(el) : undefined;
        if (body) T.formatAll(body, fn);
        // Tables: every cell.
        if (el?.localName === "graphicFrame") for (const tx of Array.from(el.getElementsByTagNameNS(PptxDom.NS_A, "txBody"))) T.formatAll(tx, fn);
      }
    });
  }

  /** Applies a paragraph change to the paragraphs of the selection (all paragraphs of selected frames). */
  private formatParas(label: string, fn: (pPr: Element, p: Element) => void): void {
    const ed = this.editor, t = ed.text;
    if (t) {
      const body = ed.textBody(t.id);
      if (!body) return;
      ed.edit(label, () => T.paraProps(body, T.minPos(t.caret, t.anchor).p, T.maxPos(t.caret, t.anchor).p, fn));
      return;
    }
    const views = ed.selectedViews;
    if (!views.length) return;
    ed.edit(label, () => {
      for (const v of views) for (const m of flat([v])) {
        const el = ed.shapeEl(m.id);
        const bodies = el ? [S.txBodyOf(el), ...Array.from(el.getElementsByTagNameNS(PptxDom.NS_A, "txBody"))].filter((b): b is Element => !!b) : [];
        for (const body of bodies) T.paraProps(body, 0, T.paras(body).length - 1, fn);
      }
    });
  }

  /** The run style at the caret (or of the first selected shape's first run). */
  private runStyle(): import("@raabe/ooxml").RunStyle | null {
    const ed = this.editor;
    const id = ed.text?.id ?? ed.sel[0];
    if (id === undefined) return null;
    const v = ed.shapeView(id);
    const text = v?.text;
    if (!text) return null;
    if (ed.text) {
      const pos = T.minPos(ed.text.caret, ed.text.anchor);
      const p = text.paras[pos.p];
      if (!p) return null;
      let off = 0;
      for (const r of p.runs) {
        const len = r.kind === "br" ? 1 : r.text.length;
        if (pos.o > off && pos.o <= off + len) return r;
        if (pos.o === 0 && off === 0 && len) return r;
        off += len;
      }
      return p.runs[p.runs.length - 1] ?? p.end;
    }
    const p = text.paras.find((x) => x.runs.length) ?? text.paras[0];
    return p ? p.runs[0] ?? p.end : null;
  }

  private paraStyle(): import("@raabe/ooxml").TextPara | null {
    const ed = this.editor;
    const id = ed.text?.id ?? ed.sel[0];
    if (id === undefined) return null;
    const text = ed.shapeView(id)?.text;
    if (!text) return null;
    return text.paras[ed.text ? T.minPos(ed.text.caret, ed.text.anchor).p : 0] ?? null;
  }

  private hasText(): boolean {
    const ed = this.editor;
    return !!ed.text || ed.selectedViews.some((v) => !!v.text || (v.kind === "graphicFrame" && v.frame === "table"));
  }

  private toggleRun(attr: "b" | "i", label: string): void {
    const st = this.runStyle();
    const on = !(attr === "b" ? st?.bold : st?.italic);
    this.formatRuns(label, (r) => r.setAttribute(attr, on ? "1" : "0"));
  }

  private setSizeAll(pt: number): void {
    const v = Math.max(1, Math.min(4000, Math.round(pt * 100)));
    this.formatRuns("Schriftgrad", (r) => r.setAttribute("sz", String(v)));
  }

  private stepSize(d: 1 | -1): void {
    const cur = this.runStyle()?.size ?? 18;
    const next = d > 0 ? SIZES.find((s) => s > cur + 0.01) ?? Math.ceil(cur / 10) * 10 + 10 : [...SIZES].reverse().find((s) => s < cur - 0.01) ?? Math.max(1, cur - 1);
    this.setSizeAll(next);
  }

  private align(a: "l" | "ctr" | "r" | "just"): void {
    this.formatParas("Ausrichten", (pPr) => pPr.setAttribute("algn", a));
  }

  /** Aufzählungszeichen / Nummerierung on and off, as PowerPoint writes them. */
  private toggleBullets(kind: "char" | "auto", value = kind === "char" ? "•" : "arabicPeriod"): void {
    const p = this.paraStyle();
    const on = !(p?.bullet && p.bullet.kind === kind && (kind === "char" ? p.bullet.char === value : p.bullet.autoType === value));
    this.formatParas(kind === "char" ? "Aufzählungszeichen" : "Nummerierung", (pPr) => {
      for (const c of kids(pPr)) if (["buNone", "buChar", "buAutoNum", "buBlip", "buFont", "buFontTx"].includes(c.localName)) c.remove();
      const doc = pPr.ownerDocument;
      const order = ["tabLst", "defRPr", "extLst"];
      const before = kids(pPr).find((c) => order.includes(c.localName)) ?? null;
      if (!on) { pPr.insertBefore(make(doc, "a:buNone"), before); pPr.setAttribute("marL", "0"); pPr.setAttribute("indent", "0"); return; }
      const lvl = Number(pPr.getAttribute("lvl") ?? 0);
      const marL = kind === "char" ? 285750 + lvl * 457200 : 342900 + lvl * 457200;
      pPr.setAttribute("marL", String(marL));
      pPr.setAttribute("indent", String(kind === "char" ? -285750 : -342900));
      if (kind === "char") {
        // PowerPoint's fonts for its bullets: Wingdings for the square, star, arrow and tick, Courier New for the hollow circle.
        const face = ["§", "Ø", "ü", "v"].includes(value) ? "Wingdings" : value === "o" ? "Courier New" : "Arial";
        pPr.insertBefore(make(doc, "a:buFont", { typeface: face, panose: face === "Arial" ? "020B0604020202020204" : face === "Courier New" ? "02070309020205020404" : "05000000000000000000", pitchFamily: face === "Wingdings" ? "2" : "34", charset: face === "Wingdings" ? "2" : "0" }), before);
        pPr.insertBefore(make(doc, "a:buChar", { char: value }), before);
      } else pPr.insertBefore(make(doc, "a:buAutoNum", { type: value }), before);
    });
  }

  private lineSpacing(v: number): void {
    this.formatParas("Zeilenabstand", (pPr) => {
      kid(pPr, "lnSpc")?.remove();
      const ln = make(pPr.ownerDocument, "a:lnSpc", {}, make(pPr.ownerDocument, "a:spcPct", { val: String(Math.round(v * 100000)) }));
      pPr.insertBefore(ln, pPr.firstChild);
    });
  }

  /** Format übertragen: picks up the run, paragraph and shape formatting of the selection. */
  private pickFormat(keep: boolean): void {
    const ed = this.editor;
    const id = ed.text?.id ?? ed.sel[0];
    if (id === undefined) return;
    const el = ed.shapeEl(id);
    if (!el) return;
    const body = S.txBodyOf(el);
    let rPr: Element | null = null, pPr: Element | null = null;
    if (body) {
      const pos = ed.text ? T.minPos(ed.text.caret, ed.text.anchor) : { p: 0, o: 1 };
      const p = T.paras(body)[pos.p];
      rPr = p ? T.formatAt(p, Math.max(1, pos.o)) : null;
      pPr = p ? kid(p, "pPr") ?? null : null;
    }
    const spPr = ed.text ? null : kid(el, "spPr") ?? null;
    this.painter = { rPr: rPr?.cloneNode(true) as Element | null, pPr: pPr?.cloneNode(true) as Element | null, spPr: spPr?.cloneNode(true) as Element | null, keep };
    document.body.classList.add("painting");
    this.updateChrome();
  }

  /** Applies picked formatting to the new selection (called after a click). */
  private applyPainter(): void {
    const pt = this.painter;
    if (!pt) return;
    const ed = this.editor;
    const copyAttrs = (from: Element, to: Element) => {
      for (const a of Array.from(to.attributes)) if (a.name !== "lang" && a.name !== "lvl") to.removeAttribute(a.name);
      while (to.firstChild) to.firstChild.remove();
      for (const a of Array.from(from.attributes)) if (a.name !== "lang" && a.name !== "lvl") to.setAttribute(a.name, a.value);
      for (const c of Array.from(from.childNodes)) to.append(c.cloneNode(true));
    };
    if (ed.text && pt.rPr) {
      if (T.cmpPos(ed.text.caret, ed.text.anchor) !== 0) this.formatRuns("Format übertragen", (r) => copyAttrs(pt.rPr!, r));
    } else if (ed.sel.length) {
      ed.edit("Format übertragen", () => {
        for (const v of ed.selectedViews) {
          const el = ed.shapeEl(v.id);
          if (!el) continue;
          if (pt.spPr) {
            const sp = kid(el, "spPr");
            if (sp) {
              const xf = kid(sp, "xfrm");
              const geom = kids(sp).find((c) => c.localName === "prstGeom" || c.localName === "custGeom");
              for (const c of kids(sp)) if (c !== xf && c !== geom) c.remove();
              for (const c of kids(pt.spPr)) if (c.localName !== "xfrm" && c.localName !== "prstGeom" && c.localName !== "custGeom") sp.append(c.cloneNode(true));
            }
          }
          const body = S.txBodyOf(el);
          if (body && pt.rPr) T.formatAll(body, (r) => copyAttrs(pt.rPr!, r));
        }
      });
    } else return;
    if (!pt.keep) { this.painter = null; document.body.classList.remove("painting"); }
  }

  // ----- shapes ---------------------------------------------------------------------------------------------

  private setFill(hex: string | null): void {
    const views = this.editor.selectedViews.filter((v) => v.kind === "sp" || v.kind === "pic" || v.kind === "grpSp");
    if (!views.length) return;
    if (hex) this.lastFill = hex;
    this.editor.edit("Fülleffekt", () => {
      for (const v of views) {
        const el = this.editor.shapeEl(v.id);
        const sp = el && kid(el, v.kind === "grpSp" ? "grpSpPr" : "spPr");
        if (!sp) continue;
        for (const c of kids(sp)) if (["noFill", "solidFill", "gradFill", "blipFill", "pattFill", "grpFill"].includes(c.localName)) c.remove();
        const f = hex ? fragment(el!.ownerDocument, `<a:solidFill>${colorXml(hex)}</a:solidFill>`)[0] : make(el!.ownerDocument, "a:noFill");
        const after = kids(sp).find((c) => c.localName === "prstGeom" || c.localName === "custGeom") ?? kid(sp, "xfrm");
        sp.insertBefore(f, after ? after.nextSibling : sp.firstChild);
      }
    });
  }

  /** The entries under Formkontur's palette. */
  /**
   * Fülleffekt's own entries (PowerPoint's): Keine Füllung, Bild..., Farbverlauf with 13 light and 13 dark
   * variants of the shape's colour (our reading of each direction: lin with PowerPoint's angles, path for the
   * corners and the middle).
   */
  fillExtra(): MenuEntry[] {
    const DIRS: [string, string][] = [
      ["Linear diagonal - oben links nach unten rechts", '<a:lin ang="2700000" scaled="1"/>'], ["Linear unten", '<a:lin ang="5400000" scaled="1"/>'],
      ["Linear diagonal - oben rechts nach unten links", '<a:lin ang="8100000" scaled="1"/>'], ["Von unterer rechter Ecke", '<a:path path="rect"><a:fillToRect l="100000" t="100000"/></a:path>'],
      ["Von unterer linker Ecke", '<a:path path="rect"><a:fillToRect t="100000" r="100000"/></a:path>'], ["Linear rechts", '<a:lin ang="0" scaled="1"/>'],
      ["Aus der Mitte", '<a:path path="circle"><a:fillToRect l="50000" t="50000" r="50000" b="50000"/></a:path>'], ["Linear links", '<a:lin ang="10800000" scaled="1"/>'],
      ["Von oberer rechter Ecke", '<a:path path="rect"><a:fillToRect l="100000" b="100000"/></a:path>'], ["Von oberer linker Ecke", '<a:path path="rect"><a:fillToRect r="100000" b="100000"/></a:path>'],
      ["Linear diagonal - unten links nach oben rechts", '<a:lin ang="18900000" scaled="1"/>'], ["Linear oben", '<a:lin ang="16200000" scaled="1"/>'],
      ["Linear diagonal - unten rechts nach oben links", '<a:lin ang="13500000" scaled="1"/>']];
    const base = () => {
      // The shape's own colour, as written (a design colour stays one), else Akzent 1.
      const v = this.editor.selectedViews[0];
      const sf = v ? kid(kid(this.editor.shapeEl(v.id), "spPr"), "solidFill") : undefined;
      const clr = sf?.firstElementChild;
      return clr ? { tag: clr.localName, val: clr.getAttribute("val") ?? "accent1" } : { tag: "schemeClr", val: "accent1" };
    };
    const stop = (pos: number, mods: string) => { const b = base(); return `<a:gs pos="${pos}"><a:${b.tag} val="${b.val}">${mods}</a:${b.tag}></a:gs>`; };
    const grad = (light: boolean, dir: string) => `<a:gradFill flip="none" rotWithShape="1"><a:gsLst>${light
      ? stop(0, '<a:lumMod val="20000"/><a:lumOff val="80000"/>') + stop(100000, '<a:lumMod val="60000"/><a:lumOff val="40000"/>')
      : stop(0, "") + stop(100000, '<a:lumMod val="50000"/>')}</a:gsLst>${dir}<a:tileRect/></a:gradFill>`;
    const setXml = (label: string, xml: string | null) => {
      const views = this.editor.selectedViews.filter((v) => v.kind === "sp" || v.kind === "cxnSp");
      this.editor.edit(label, () => {
        for (const v of views) {
          const sp = kid(this.editor.shapeEl(v.id), "spPr");
          if (!sp) continue;
          for (const c of kids(sp)) if (["noFill", "solidFill", "gradFill", "blipFill", "pattFill", "grpFill"].includes(c.localName)) c.remove();
          const f = xml ? fragment(sp.ownerDocument, xml)[0] : make(sp.ownerDocument, "a:noFill");
          const after = kids(sp).find((c) => c.localName === "prstGeom" || c.localName === "custGeom") ?? kid(sp, "xfrm");
          sp.insertBefore(f, after ? after.nextSibling : sp.firstChild);
        }
      });
    };
    return [
      { label: "Keine Füllung", run: () => this.setFill(null) },
      { label: "Bild...", run: () => void this.pictureFill() },
      { label: "Farbverlauf", submenu: () => [
        { label: "Kein Farbverlauf", run: () => this.setFill(this.lastFill) },
        { heading: "Helle Varianten" }, ...DIRS.map(([label, dir]) => ({ label, run: () => setXml("Farbverlauf", grad(true, dir)) })),
        { heading: "Dunkle Varianten" }, ...DIRS.map(([label, dir]) => ({ label, run: () => setXml("Farbverlauf", grad(false, dir)) })),
        "-", { label: "Weitere Farbverläufe...", run: () => void this.commands.run("draw.format") }] },
    ];
  }

  /** Fülleffekt > Bild...: a picture file fills the selected shapes (a:blipFill, stretched). */
  private async pictureFill(): Promise<void> {
    const path = await this.platform.openImageDialog();
    const slide = this.slidePath();
    if (!path || !slide) return;
    const bytes = await this.platform.readFile(path);
    const ext = /\.jpe?g$/i.test(path) ? "jpeg" : "png";
    const views = this.editor.selectedViews.filter((v) => v.kind === "sp");
    this.editor.edit("Fülleffekt", () => {
      const media = this.deck.pptx.freeName(`ppt/media/image#.${ext}`);
      this.deck.pptx.put(media, bytes);
      this.deck.pptx.setType(media, ext === "jpeg" ? "image/jpeg" : "image/png", true);
      const rid = this.deck.pptx.link(slide, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image", media);
      for (const v of views) {
        const sp = kid(this.editor.shapeEl(v.id), "spPr");
        if (!sp) continue;
        for (const c of kids(sp)) if (["noFill", "solidFill", "gradFill", "blipFill", "pattFill", "grpFill"].includes(c.localName)) c.remove();
        const f = fragment(sp.ownerDocument, `<a:blipFill dpi="0" rotWithShape="1"><a:blip r:embed="${rid}"/><a:srcRect/><a:stretch><a:fillRect/></a:stretch></a:blipFill>`)[0];
        const after = kids(sp).find((c) => c.localName === "prstGeom" || c.localName === "custGeom") ?? kid(sp, "xfrm");
        sp.insertBefore(f, after ? after.nextSibling : sp.firstChild);
      }
    });
  }

  /** Formkontur > Pfeile: the arrowheads of the selected lines (our reading of PowerPoint's eleven Pfeilarten). */
  private arrowEntries(): MenuEntry[] {
    type End = string | null;
    const ARROWS: [End, End][] = [[null, null], [null, "arrow"], ["arrow", null], ["arrow", "arrow"], [null, "triangle"], ["triangle", null], ["triangle", "triangle"],
      ["oval", "triangle"], ["oval", "oval"], ["diamond", "diamond"], [null, "stealth"]];
    const set = (head: End, tail: End) => {
      const views = this.editor.selectedViews.filter((v) => v.kind === "cxnSp" || v.kind === "sp");
      this.editor.edit("Pfeile", () => {
        for (const v of views) {
          const sp = kid(this.editor.shapeEl(v.id), "spPr");
          if (!sp) continue;
          let ln = kid(sp, "ln");
          if (!ln) { ln = make(sp.ownerDocument, "a:ln"); sp.insertBefore(ln, kids(sp).find((c) => ["effectLst", "effectDag", "scene3d", "sp3d", "extLst"].includes(c.localName)) ?? null); }
          for (const n of ["headEnd", "tailEnd"]) kid(ln, n)?.remove();
          const after = kids(ln).find((c) => c.localName === "extLst") ?? null;
          if (head) ln.insertBefore(make(sp.ownerDocument, "a:headEnd", { type: head }), after);
          if (tail) ln.insertBefore(make(sp.ownerDocument, "a:tailEnd", { type: tail }), after);
        }
      });
    };
    return [...ARROWS.map(([hd, tl], i) => ({ label: `Pfeilart ${i + 1}`, run: () => set(hd, tl) })), "-", { label: "Weitere Pfeile...", run: () => void this.commands.run("draw.format") }];
  }

  outlineExtra(): MenuEntry[] {
    return [{ label: "Keine Kontur", run: () => this.setLine(null) },
      { label: "Stärke", submenu: () => LINE_WEIGHTS.map(([label, w]) => ({ label, run: () => this.setLine(undefined, w) })) },
      { label: "Striche", submenu: () => LINE_DASHES.map(([label, d]) => ({ label, run: () => this.setLine(undefined, undefined, d) })) },
      { label: "Pfeile", submenu: () => this.arrowEntries() }];
  }

  /** Bildformat > Zuschneiden: PowerPoint's menu. */
  cropMenu(): MenuEntry[] {
    const ratios: [string, [number, number][]][] = [["Quadrat", [[1, 1]]], ["Hochformat", [[2, 3], [3, 4], [3, 5], [4, 5]]], ["Querformat", [[3, 2], [4, 3], [5, 3], [5, 4], [16, 9], [16, 10]]]];
    return [
      { label: "Zuschneiden", run: () => this.editor.toggleCrop() },
      { label: "Auf Form zuschneiden", submenu: () => PRESENT_SHAPES.filter((g) => !["Linien", "Interaktive Schaltflächen"].includes(g.title)).flatMap((g) => [{ heading: g.title } as MenuEntry,
        ...g.items.filter((it) => it.tool.kind === "shape" && !it.tool.freeform).map((it) => ({ label: it.label, run: () => this.changeShape(it.tool.geometry, true) }))]) },
      { label: "Seitenverhältnis", submenu: () => ratios.flatMap(([head, list]) => [{ heading: head } as MenuEntry, ...list.map(([a, b]) => ({ label: `${a}:${b}`, run: () => void this.cropToRatio(a / b) }))]) },
      "-",
      { label: "Ausfüllen", run: () => void this.cropFill(true) },
      { label: "Einpassen", run: () => void this.cropFill(false) },
    ];
  }

  /** The picture's own size in pixels (for cropping by proportions). */
  private async pictureSize(id: number): Promise<{ w: number; h: number } | null> {
    const v = this.editor.shapeView(id);
    const url = v?.pic?.path ? this.image(v.pic.path) : undefined;
    if (!url) return null;
    return await new Promise((res) => { const i = new Image(); i.onload = () => res({ w: i.naturalWidth || 1, h: i.naturalHeight || 1 }); i.onerror = () => res(null); i.src = url; });
  }

  /** Seitenverhältnis: the picture cropped about its middle to the proportion, its width kept. */
  private async cropToRatio(ratio: number): Promise<void> {
    const v = this.editor.selectedViews.find((x) => x.kind === "pic");
    const size = v ? await this.pictureSize(v.id) : null;
    if (!v || !size) return;
    const a = size.w / size.h;
    const crop = a > ratio ? { l: (1 - ratio / a) / 2, r: (1 - ratio / a) / 2, t: 0, b: 0 } : { l: 0, r: 0, t: (1 - a / ratio) / 2, b: (1 - a / ratio) / 2 };
    this.setPictureCrop(v.id, crop, { x: v.x, y: v.y + (v.h - v.w / ratio) / 2, w: v.w, h: v.w / ratio }, "Seitenverhältnis");
  }

  /** Ausfüllen: the picture fills its frame (cropped about its middle); Einpassen: the frame takes the picture's proportions. */
  private async cropFill(fill: boolean): Promise<void> {
    const v = this.editor.selectedViews.find((x) => x.kind === "pic");
    const size = v ? await this.pictureSize(v.id) : null;
    if (!v || !size) return;
    const a = size.w / size.h, f = v.w / v.h;
    if (fill) {
      const crop = a > f ? { l: (1 - f / a) / 2, r: (1 - f / a) / 2, t: 0, b: 0 } : { l: 0, r: 0, t: (1 - a / f) / 2, b: (1 - a / f) / 2 };
      this.setPictureCrop(v.id, crop, { x: v.x, y: v.y, w: v.w, h: v.h }, "Ausfüllen");
    } else {
      const w = a > f ? v.w : v.h * a, h = a > f ? v.w / a : v.h;
      this.setPictureCrop(v.id, { l: 0, r: 0, t: 0, b: 0 }, { x: v.x + (v.w - w) / 2, y: v.y + (v.h - h) / 2, w, h }, "Einpassen");
    }
  }

  private setPictureCrop(id: number, c: { l: number; t: number; r: number; b: number }, box: { x: number; y: number; w: number; h: number }, label: string): void {
    this.editor.edit(label, () => {
      const el = this.editor.shapeEl(id);
      const bf = el ? kid(el, "blipFill") : undefined;
      if (!el || !bf) return;
      kid(bf, "srcRect")?.remove();
      const pc = (v: number) => String(Math.round(v * 100000));
      const src = make(el.ownerDocument, "a:srcRect", {});
      for (const [k, v] of Object.entries(c)) if (Math.abs(v) > 0.00001) src.setAttribute(k, pc(v));
      bf.insertBefore(src, kid(bf, "stretch") ?? kid(bf, "tile") ?? null);
      S.setXfrm(el, box);
    });
  }

  /** Textkontur (WordArt-Formate): the runs' outline colour, Keine Kontur, Stärke, Striche. */
  private lastTextOutline = "none";
  textOutlineColor(): string { return this.lastTextOutline === "none" ? "none" : `#${this.lastTextOutline}`; }
  setTextOutline(o: { color?: string | null; width?: number; dash?: string }): void {
    if (o.color) this.lastTextOutline = o.color.replace("#", "").toUpperCase();
    if (o.color === null) this.lastTextOutline = "none";
    const color = o.color ? o.color.replace("#", "").toUpperCase() : o.color;
    this.formatRuns("Textkontur", (r) => T.setRunOutline(r, { ...o, color, width: o.width ?? (o.color && !r.getElementsByTagNameNS(PptxDom.NS_A, "ln")[0]?.getAttribute("w") ? 9525 : o.width) }));
  }
  textOutlineExtra(): MenuEntry[] {
    return [{ label: "Keine Kontur", run: () => this.setTextOutline({ color: null }) },
      { label: "Stärke", submenu: () => LINE_WEIGHTS.map(([label, w]) => ({ label, run: () => this.setTextOutline({ width: Math.round(w * 12700) }) })) },
      { label: "Striche", submenu: () => LINE_DASHES.map(([label, d]) => ({ label, run: () => this.setTextOutline({ dash: d }) })) }];
  }

  private setLine(hex: string | null | undefined, width?: number, dash?: string): void {
    const views = this.editor.selectedViews.filter((v) => v.kind === "sp" || v.kind === "pic" || v.kind === "cxnSp");
    if (!views.length) return;
    if (hex) this.lastLine = hex;
    this.editor.edit("Formkontur", () => {
      for (const v of views) {
        const el = this.editor.shapeEl(v.id);
        const sp = el && kid(el, "spPr");
        if (!sp) continue;
        let ln = kid(sp, "ln");
        if (!ln) {
          ln = make(el!.ownerDocument, "a:ln");
          const before = kids(sp).find((c) => ["effectLst", "effectDag", "scene3d", "sp3d", "extLst"].includes(c.localName)) ?? null;
          sp.insertBefore(ln, before);
        }
        if (hex !== undefined) {
          for (const c of kids(ln)) if (["noFill", "solidFill", "gradFill", "pattFill"].includes(c.localName)) c.remove();
          ln.insertBefore(hex ? fragment(el!.ownerDocument, `<a:solidFill>${colorXml(hex)}</a:solidFill>`)[0] : make(el!.ownerDocument, "a:noFill"), ln.firstChild);
        }
        if (width !== undefined) ln.setAttribute("w", String(Math.round(width * 12700)));
        if (dash) { kid(ln, "prstDash")?.remove(); const after = kids(ln).find((c) => ["noFill", "solidFill", "gradFill", "pattFill"].includes(c.localName)); ln.insertBefore(make(el!.ownerDocument, "a:prstDash", { val: dash }), after ? after.nextSibling : ln.firstChild); }
      }
    });
  }

  /** Anordnen: z-order changes among the slide's top-level shapes. */
  private order(kind: "front" | "back" | "forward" | "backward"): void {
    const ed = this.editor;
    const ids = new Set(ed.sel);
    if (!ids.size) return;
    ed.edit(kind === "front" ? "In den Vordergrund" : kind === "back" ? "In den Hintergrund" : kind === "forward" ? "Ebene nach vorne" : "Ebene nach hinten", () => {
      const tree = S.treeOf(this.deck.pptx.doc(this.slidePath()!));
      const els = S.shapeEls(tree);
      const picked = els.filter((e) => ids.has(S.shapeId(e)));
      if (kind === "front") for (const e of picked) tree.append(e);
      else if (kind === "back") { const first = S.shapeEls(tree)[0]; for (const e of picked) tree.insertBefore(e, first); }
      else if (kind === "forward") for (const e of [...picked].reverse()) { const next = e.nextElementSibling; if (next && S.SHAPE_TAGS.includes(next.localName) && !ids.has(S.shapeId(next))) next.after(e); }
      else for (const e of picked) { const prev = e.previousElementSibling; if (prev && S.SHAPE_TAGS.includes(prev.localName) && !ids.has(S.shapeId(prev))) prev.before(e); }
    });
  }

  /** Ausrichten / Verteilen: to the slide when one shape is selected (or "An Folie ausrichten"), else to the selection. */
  private alignShapes(kind: "l" | "c" | "r" | "t" | "m" | "b" | "dh" | "dv"): void {
    const ed = this.editor;
    const views = ed.selectedViews;
    if (!views.length) return;
    const size = this.deck.pptx.slideSize;
    const toSlide = views.length === 1 || this.alignToSlide;
    const box = toSlide ? { x: 0, y: 0, w: size.cx, h: size.cy } : union(views);
    ed.edit(kind.startsWith("d") ? "Verteilen" : "Ausrichten", () => {
      if (kind === "dh" || kind === "dv") {
        const hz = kind === "dh";
        const sorted = [...views].sort((a, b) => (hz ? a.x - b.x : a.y - b.y));
        const total = sorted.reduce((s, v) => s + (hz ? v.w : v.h), 0);
        const span = hz ? box.w : box.h;
        const gap = sorted.length > 1 ? (span - total) / (sorted.length - 1 + (toSlide && views.length > 2 ? 0 : 0)) : 0;
        let pos = hz ? box.x : box.y;
        for (const v of sorted) { ed.moveShape(v, hz ? pos - v.x : 0, hz ? 0 : pos - v.y); pos += (hz ? v.w : v.h) + gap; }
        return;
      }
      for (const v of views) {
        const dx = kind === "l" ? box.x - v.x : kind === "c" ? box.x + box.w / 2 - (v.x + v.w / 2) : kind === "r" ? box.x + box.w - (v.x + v.w) : 0;
        const dy = kind === "t" ? box.y - v.y : kind === "m" ? box.y + box.h / 2 - (v.y + v.h / 2) : kind === "b" ? box.y + box.h - (v.y + v.h) : 0;
        if (dx || dy) ed.moveShape(v, dx, dy);
      }
    });
  }

  private alignToSlide = false;

  /** Zeilen / Spalten löschen: the row or column of the cell being edited (the whole table when it is the last). */
  private tableDelete(what: "row" | "col"): void {
    const t = this.editor.text;
    if (!t?.cell) return;
    const { r, c } = t.cell;
    const id = t.id;
    this.editor.endText();
    this.editor.edit(what === "row" ? "Zeilen löschen" : "Spalten löschen", () => {
      const el = this.editor.shapeEl(id)!;
      const tbl = el.getElementsByTagNameNS(PptxDom.NS_A, "tbl")[0];
      const rows = kids(tbl, "tr");
      const ext = kid(kid(el, "xfrm"), "ext");
      if (what === "row") {
        if (rows.length <= 1) { el.remove(); return; }
        if (ext) ext.setAttribute("cy", String(Math.max(0, (num(ext, "cy") ?? 0) - (num(rows[r], "h") ?? 0))));
        rows[r].remove();
      } else {
        const cols = kids(kid(tbl, "tblGrid"), "gridCol");
        if (cols.length <= 1) { el.remove(); return; }
        if (ext) ext.setAttribute("cx", String(Math.max(0, (num(ext, "cx") ?? 0) - (num(cols[c], "w") ?? 0))));
        cols[c].remove();
        for (const tr of rows) kids(tr, "tc")[c]?.remove();
      }
    });
    this.editor.select(this.editor.shapeEl(id) ? [id] : []);
  }

  private selectedHasAnimation(): boolean {
    const p = this.slidePath();
    if (!p || !this.editor.sel.length) return false;
    return effects(this.deck.pptx.doc(p).documentElement).list.some((e) => this.editor.sel.includes(e.spid));
  }

  private moveAnimation(d: -1 | 1): void {
    const p = this.slidePath();
    if (!p) return;
    const list = effects(this.deck.pptx.doc(p).documentElement).list;
    const i = list.findIndex((e) => this.editor.sel.includes(e.spid));
    if (i >= 0) this.animPane.move(i, d);
  }

  /** Bild zurücksetzen: the picture's formatting goes (crop, effects, outline); place and size stay. */
  private resetPicture(): void {
    const views = this.editor.selectedViews.filter((v) => v.kind === "pic");
    this.editor.edit("Bild zurücksetzen", () => {
      for (const v of views) {
        const el = this.editor.shapeEl(v.id)!;
        const bf = kid(el, "blipFill");
        kid(bf, "srcRect")?.remove();
        const blip = kid(bf, "blip");
        if (blip) for (const c of kids(blip)) if (c.localName !== "extLst") c.remove();
        const sp = kid(el, "spPr");
        if (sp) for (const c of kids(sp)) if (!["xfrm", "prstGeom"].includes(c.localName)) c.remove();
        const pg = kid(sp, "prstGeom");
        if (pg) pg.setAttribute("prst", "rect");
      }
    });
  }

  /** Grafik ändern: another picture file in the same place, keeping the size of the old one's width. */
  private async changePicture(): Promise<void> {
    const v = this.editor.selectedViews[0];
    const path = await this.platform.openImageDialog();
    if (!path || !v) return;
    const bytes = await this.platform.readFile(path);
    const type = mimeOf(path);
    const ext = type === "image/jpeg" ? "jpeg" : type === "image/gif" ? "gif" : "png";
    const media = this.deck.pptx.freeName(`ppt/media/image#.${ext}`);
    const slide = this.slidePath()!;
    this.editor.edit("Grafik ändern", () => {
      this.deck.pptx.put(media, bytes);
      this.deck.pptx.setType(media, type, true);
      const rid = this.deck.pptx.link(slide, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image", media);
      const blip = at(this.editor.shapeEl(v.id)!, "blipFill/blip");
      blip?.setAttributeNS(PptxDom.NS_R, "r:embed", rid);
    });
  }

  /** Zeilen und Spalten: a row above / below the first / last row, a column left / right (PowerPoint copies the neighbour's formatting). */
  private tableInsert(where: "above" | "below" | "left" | "right"): void {
    const v = this.editor.selectedViews.find((x) => x.frame === "table");
    if (!v) return;
    this.editor.edit(where === "above" || where === "below" ? "Zeile einfügen" : "Spalte einfügen", () => {
      const el = this.editor.shapeEl(v.id)!;
      const tbl = el.getElementsByTagNameNS(PptxDom.NS_A, "tbl")[0];
      const rows = kids(tbl, "tr");
      const emptyCell = (tc: Element) => { const c = tc.cloneNode(true) as Element; const body = kid(c, "txBody"); if (body) { T.replaceAll(body, ""); } return c; };
      if (where === "above" || where === "below") {
        const ref = where === "above" ? rows[0] : rows[rows.length - 1];
        const nr = ref.cloneNode(false) as Element;
        for (const tc of kids(ref, "tc")) nr.append(emptyCell(tc));
        if (where === "above") ref.before(nr); else ref.after(nr);
        const xf = kid(el, "xfrm"), ext = kid(xf, "ext");
        if (ext) ext.setAttribute("cy", String((num(ext, "cy") ?? 0) + (num(ref, "h") ?? 370840)));
      } else {
        const grid = kid(tbl, "tblGrid")!;
        const cols = kids(grid, "gridCol");
        const refCol = where === "left" ? cols[0] : cols[cols.length - 1];
        const nc = refCol.cloneNode(false) as Element;
        if (where === "left") refCol.before(nc); else refCol.after(nc);
        for (const tr of rows) {
          const tcs = kids(tr, "tc");
          const ref = where === "left" ? tcs[0] : tcs[tcs.length - 1];
          const c = emptyCell(ref);
          if (where === "left") ref.before(c); else ref.after(c);
        }
        const xf = kid(el, "xfrm"), ext = kid(xf, "ext");
        if (ext) ext.setAttribute("cx", String((num(ext, "cx") ?? 0) + (num(refCol, "w") ?? 0)));
      }
    });
  }

  private rotateFlip(kind: "r" | "l" | "h" | "v"): void {
    const ed = this.editor;
    const views = ed.selectedViews;
    ed.edit(kind === "h" ? "Horizontal kippen" : kind === "v" ? "Vertikal kippen" : "Drehen", () => {
      for (const v of views) {
        const el = ed.shapeEl(v.id);
        if (!el) continue;
        if (kind === "r" || kind === "l") S.setXfrm(el, { x: v.x, y: v.y, w: v.w, h: v.h, rot: v.rot + (kind === "r" ? 90 : -90) });
        else S.setXfrm(el, { x: v.x, y: v.y, w: v.w, h: v.h, flipH: kind === "h" ? !v.flipH : v.flipH, flipV: kind === "v" ? !v.flipV : v.flipV });
      }
    });
  }

  /** Gruppieren: the selected top-level shapes into a p:grpSp with child coordinates equal to slide ones. */
  private group(): void {
    const ed = this.editor;
    const views = ed.selectedViews.filter((v) => this.deck.view(this.slidePath()!).shapes.includes(v));
    if (views.length < 2) return;
    let gid = 0;
    ed.edit("Gruppieren", () => {
      const doc = this.deck.pptx.doc(this.slidePath()!);
      const tree = S.treeOf(doc);
      const ids = new Set(views.map((v) => v.id));
      const els = S.shapeEls(tree).filter((e) => ids.has(S.shapeId(e)));
      const b = union(views);
      gid = S.maxId(tree) + 1;
      const g = fragment(doc, `<p:grpSp><p:nvGrpSpPr><p:cNvPr id="${gid}" name="Gruppieren ${gid - 1}"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="${Math.round(b.x)}" y="${Math.round(b.y)}"/><a:ext cx="${Math.round(b.w)}" cy="${Math.round(b.h)}"/><a:chOff x="${Math.round(b.x)}" y="${Math.round(b.y)}"/><a:chExt cx="${Math.round(b.w)}" cy="${Math.round(b.h)}"/></a:xfrm></p:grpSpPr></p:grpSp>`)[0];
      els[els.length - 1].after(g);
      for (const e of els) {
        // Placeholders keep their place: they get their own xfrm before leaving the tree.
        const v = views.find((x) => x.id === S.shapeId(e))!;
        if (v.ph) S.setXfrm(e, { x: v.x, y: v.y, w: v.w, h: v.h });
        g.append(e);
      }
    });
    ed.select([gid]);
  }

  /** The members of the group last taken apart (Gruppierung wiederherstellen). */
  private lastUngrouped: number[] = [];

  /** Puts the master's placeholders of these types on a layout, or takes them off (Folienmaster > Titel, Fußzeilen). */
  private toggleLayoutPh(layout: string | undefined, types: string[]): void {
    if (!layout) return;
    const master = this.deck.pptx.masterOf(layout);
    if (!master) return;
    this.deck.change(types.includes("title") ? "Titel" : "Fußzeilen", [layout], () => {
      const doc = this.deck.pptx.doc(layout);
      const tree = S.treeOf(doc);
      const mine = S.shapeEls(tree).filter((sp) => types.includes(phOf(sp)?.type ?? ""));
      if (mine.length) { for (const sp of mine) sp.remove(); return; }
      let next = S.maxId(tree);
      for (const sp of S.shapeEls(S.treeOf(this.deck.pptx.doc(master)))) {
        if (!types.includes(phOf(sp)?.type ?? "")) continue;
        const copy = doc.importNode(sp, true) as Element;
        S.renumber(copy, () => ++next);
        tree.append(copy);
      }
    });
    this.editor.render(); this.panel.render(); this.updateChrome();
  }

  private ungroup(): void {
    const ed = this.editor;
    const groups = ed.selectedViews.filter((v) => v.kind === "grpSp");
    if (!groups.length) return;
    const members: number[] = [];
    ed.edit("Gruppierung aufheben", () => {
      for (const gv of groups) {
        const g = ed.shapeEl(gv.id);
        if (!g) continue;
        for (const c of gv.children ?? []) {
          const el = ed.shapeEl(c.id);
          if (!el) continue;
          S.setXfrm(el, { x: c.x, y: c.y, w: c.w, h: c.h, rot: c.rot });
          g.before(el);
          members.push(c.id);
        }
        g.remove();
      }
    });
    this.lastUngrouped = members;
    ed.select(members);
  }

  // ----- clipboard -------------------------------------------------------------------------------------------

  private async copy(cut: boolean): Promise<void> {
    const ed = this.editor;
    if (ed.text) {
      const body = ed.textBody(ed.text.id);
      if (!body || T.cmpPos(ed.text.caret, ed.text.anchor) === 0) return;
      const text = T.rangeText(body, ed.text.anchor, ed.text.caret);
      this.clipboard = null;
      await this.platform.clipboardWrite(text).catch(() => {});
      if (cut) { const t = ed.text; ed.edit("Ausschneiden", () => { const pos = T.deleteRange(body, t.anchor, t.caret); t.caret = pos; t.anchor = pos; }); }
      return;
    }
    const views = ed.selectedViews;
    if (!views.length) return;
    const xml = views.map((v) => new XMLSerializer().serializeToString(ed.shapeEl(v.id)!));
    const text = views.map((v) => v.text ? v.text.paras.map((p) => p.runs.map((r) => r.text).join("")).join("\r\n") : "").filter(Boolean).join("\r\n");
    this.clipboard = { xml, text, slide: this.slidePath()!, pastes: 0 };
    this.slideClipboard = null;
    await this.platform.clipboardWrite(text || " ").catch(() => {});
    if (cut) { this.clipboard.pastes = -1; ed.deleteSelected(); }
  }

  private async paste(mode: "auto" | "text" = "auto"): Promise<void> {
    const ed = this.editor;
    const sys = await this.platform.clipboardRead().catch(() => "");
    const own = this.clipboard && (sys === (this.clipboard.text || " ") || sys === "");
    if (ed.text) {
      // Into the text being edited: plain text, line breaks as paragraphs.
      const text = own && this.clipboard ? this.clipboard.text : sys;
      if (text) ed.typeText(text.replace(/\r\n?/g, "\n"));
      void mode;
      return;
    }
    if (own && this.clipboard) {
      const cb = this.clipboard;
      // On the slide it came from, each paste moves down and right (PowerPoint's offset).
      cb.pastes++;
      const off = cb.slide === this.slidePath() && cb.pastes > 0 ? cb.pastes * 152400 : 0;
      const ids: number[] = [];
      ed.edit("Einfügen", () => {
        const doc = this.deck.pptx.doc(this.slidePath()!);
        const tree = S.treeOf(doc);
        let next = S.maxId(tree);
        for (const x of cb.xml) {
          const el = fragment(doc, x)[0];
          // Placeholders become plain shapes on another slide only when it has no free one; kept as they are.
          S.renumber(el, () => ++next);
          const xf = el.localName === "graphicFrame" ? kid(el, "xfrm") : at(el, `${el.localName === "grpSp" ? "grpSpPr" : "spPr"}/xfrm`);
          const o = kid(xf, "off");
          if (o && off) { o.setAttribute("x", String((num(o, "x") ?? 0) + off)); o.setAttribute("y", String((num(o, "y") ?? 0) + off)); }
          tree.append(el);
          ids.push(S.shapeId(el));
        }
      });
      ed.select(ids);
      return;
    }
    const img = await this.platform.clipboardReadImage?.().catch(() => null);
    if (img) { this.insertPictureBytes(img, "image/png"); return; }
    if (sys) {
      // Text from elsewhere: a new text box in the middle of the slide.
      const size = this.deck.pptx.slideSize;
      let id = 0;
      ed.edit("Einfügen", () => {
        const doc = this.deck.pptx.doc(this.slidePath()!);
        const tree = S.treeOf(doc);
        id = S.maxId(tree) + 1;
        const el = S.newTextBox(doc, id, size.cx / 2 - 2000000, size.cy / 2 - 200000, 4000000, 369332);
        tree.append(el);
        const body = S.txBodyOf(el)!;
        let pos = { p: 0, o: 0 };
        sys.replace(/\r\n?/g, "\n").split("\n").forEach((line, i) => { if (i) pos = T.splitPara(body, pos); if (line) pos = T.insertText(body, pos, line); });
      });
      ed.updateAutofit(id);
      ed.select([id]);
    }
  }

  // ----- inserting --------------------------------------------------------------------------------------------

  private async insertPicture(): Promise<void> {
    const path = await this.platform.openImageDialog();
    if (!path) return;
    const bytes = await this.platform.readFile(path);
    this.insertPictureBytes(bytes, mimeOf(path));
  }

  insertPictureBytes(bytes: Uint8Array, type: string): void {
    const ext = type === "image/jpeg" ? "jpeg" : type === "image/gif" ? "gif" : type === "image/svg+xml" ? "svg" : "png";
    const media = this.deck.pptx.freeName(`ppt/media/image#.${ext}`);
    const blob = new Blob([bytes as BlobPart], { type });
    const url = URL.createObjectURL(blob);
    const img = new Image();
    img.onload = () => {
      const size = this.deck.pptx.slideSize;
      // PowerPoint inserts at the picture's size (96 dpi) scaled down to fit the slide, centred.
      let w = img.naturalWidth * EMU_PX, hh = img.naturalHeight * EMU_PX;
      const k = Math.min(1, size.cx / w, size.cy / hh);
      w *= k; hh *= k;
      const slide = this.slidePath()!;
      let id = 0;
      this.editor.edit("Grafik einfügen", () => {
        this.deck.pptx.put(media, bytes);
        this.deck.pptx.setType(media, type, true);
        const rid = this.deck.pptx.link(slide, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image", media);
        const doc = this.deck.pptx.doc(slide);
        const tree = S.treeOf(doc);
        id = S.maxId(tree) + 1;
        tree.append(S.newPicture(doc, id, rid, { x: (size.cx - w) / 2, y: (size.cy - hh) / 2, w, h: hh }));
      });
      this.editor.select([id]);
      URL.revokeObjectURL(url);
    };
    img.src = url;
  }

  insertTable(rows: number, cols: number): void {
    const slide = this.slidePath();
    if (!slide) return;
    const size = this.deck.pptx.slideSize;
    let id = 0;
    this.editor.edit("Tabelle einfügen", () => {
      this.ensureTableStyle();
      const doc = this.deck.pptx.doc(slide);
      const tree = S.treeOf(doc);
      id = S.maxId(tree) + 1;
      // PowerPoint puts a new table at the top of the content area, 10 160 000 EMU wide (measured).
      const w = 10160000;
      tree.append(S.newTable(doc, id, rows, cols, { x: (size.cx - w) / 2, y: 1981200, w }, DEFAULT_TABLE_STYLE_ID));
    });
    this.editor.select([id]);
  }

  /** The current slide's design for the colour palettes: its colours, named with PowerPoint's colour words. */
  paletteTheme(): import("@raabe/shell").PaletteTheme | undefined {
    const p = this.slidePath();
    return p ? { colors: this.deck.view(p).theme.colors, word: colorWord } : undefined;
  }

  /** A remembered colour (hex or design reference) as hex, for the colour bars under the buttons. */
  private colorHex(spec: string): string {
    const p = this.slidePath();
    return p ? specHex(spec, this.deck.view(p).theme.colors) : spec;
  }

  /** Entwurf > Farben > Farben anpassen...: the design's twelve colours and a name, as PowerPoint's dialog asks. */
  private async customColorsDialog(): Promise<void> {
    const p = this.slidePath();
    if (!p) return;
    const cur = this.deck.view(p).theme.colors;
    const SLOTS: [string, string][] = [["dk1", "Text/Hintergrund - Dunkel 1"], ["lt1", "Text/Hintergrund - Hell 1"], ["dk2", "Text/Hintergrund - Dunkel 2"], ["lt2", "Text/Hintergrund - Hell 2"],
      ...[1, 2, 3, 4, 5, 6].map((n): [string, string] => [`accent${n}`, `Akzent ${n}`]), ["hlink", "Link"], ["folHlink", "Besuchter Link"]];
    const inputs = SLOTS.map(([k, label]) => { const i = h("input", { type: "color", value: `#${cur[k] ?? "000000"}` }) as HTMLInputElement; return { k, i, el: h("label", { class: "dlg-row" }, h("span", { style: "display:inline-block;width:200px" }, label), i) }; });
    const name = h("input", { type: "text", value: "Benutzerdefiniert 1", style: "width:220px" }) as HTMLInputElement;
    const ok = await showDialog({ title: "Neue Designfarben erstellen", body: h("div", { class: "form" }, h("b", null, "Designfarben"), ...inputs.map((x) => x.el), h("label", null, "Name: ", name)),
      buttons: [{ label: "Speichern", value: true, primary: true }, { label: "Abbrechen", value: false }], cancel: false });
    if (!ok) return;
    const colors: Record<string, string> = { ...cur };
    for (const x of inputs) colors[x.k] = x.i.value.slice(1).toUpperCase();
    this.applyColorSet(name.value.trim() || "Benutzerdefiniert 1", colors);
  }

  /** Start > Ersetzen > Schriftarten ersetzen...: one font for another throughout the presentation (text and design fonts). */
  private async replaceFontsDialog(): Promise<void> {
    const pptx = this.deck.pptx;
    const parts = [...this.deck.slides.map((s) => s.path), ...pptx.masters().flatMap((m) => [m, ...pptx.layouts(m)])].filter((x) => pptx.has(x));
    const themes = [...new Set(pptx.masters().map((m) => pptx.themeOf(m)).filter((t): t is string => !!t))];
    const used = new Set<string>();
    for (const part of [...parts, ...themes]) for (const n of ["latin", "ea", "cs"]) for (const e of Array.from(pptx.doc(part).getElementsByTagNameNS(PptxDom.NS_A, n))) { const t = e.getAttribute("typeface"); if (t && !t.startsWith("+")) used.add(t); }
    if (!used.size) return;
    const from = h("select", null, ...[...used].sort().map((f) => h("option", { value: f }, f))) as HTMLSelectElement;
    const to = h("select", null, ...this.ribbonContext().fontOptions().flatMap((g) => g.items).filter((f, i, a) => a.indexOf(f) === i).map((f) => h("option", { value: f }, f))) as HTMLSelectElement;
    const ok = await showDialog({ title: "Schriftart ersetzen", body: h("div", { class: "form" }, h("label", null, "Ersetzen: ", from), h("label", null, "Durch: ", to)),
      buttons: [{ label: "Ersetzen", value: true, primary: true }, { label: "Schließen", value: false }], cancel: false });
    if (!ok || from.value === to.value) return;
    this.deck.change("Schriftarten ersetzen", [...parts, ...themes], () => {
      for (const part of [...parts, ...themes]) for (const n of ["latin", "ea", "cs"]) for (const e of Array.from(pptx.doc(part).getElementsByTagNameNS(PptxDom.NS_A, n))) if (e.getAttribute("typeface") === from.value) e.setAttribute("typeface", to.value);
    });
  }

  /** Whether the presentation carries a VBA project (an opened .pptm): kept, never run (D-072). */
  hasMacros(deck: Deck): boolean {
    return deck.pptx.rels(deck.pptx.presPath).list.some((r) => r.type === RT_VBA);
  }

  /** Saving with macros as a type without them: the VBA project is left out (unreachable parts are not written). */
  private dropMacros(deck: Deck): void {
    const rels = deck.pptx.rels(deck.pptx.presPath);
    const i = rels.list.findIndex((r) => r.type === RT_VBA);
    if (i >= 0) { rels.list.splice(i, 1); rels.changed = true; }
  }

  /** An Interaktive Schaltfläche: the jump PowerPoint gives it (Weiter: nächste Folie …), as a click action. */
  actionButtonDrawn(el: Element, prst: string): void {
    const jump = ACTION_BUTTONS.find((b) => b[0] === prst)?.[2];
    const nv = el.getElementsByTagNameNS(PptxDom.NS_P, "cNvPr")[0];
    if (!jump || !nv) return;
    const hl = el.ownerDocument.createElementNS(PptxDom.NS_A, "a:hlinkClick");
    hl.setAttributeNS(PptxDom.NS_R, "r:id", "");
    hl.setAttribute("action", `ppaction://hlinkshowjump?jump=${jump}`);
    hl.setAttribute("highlightClick", "1");
    nv.insertBefore(hl, nv.firstChild);
  }

  async openActionDialog(): Promise<void> {
    await (await import("./links")).actionDialog(this);
  }

  /** Tabellenentwurf's pen: Stiftart (null: Kein Rahmen), Stiftstärke in EMU, Stiftfarbe (hex or design colour). */
  tablePen: { dash: string | null; w: number; color: string } = { dash: "solid", w: 12700, color: "000000" };

  /** Tabelle zeichnen: a click on a cell line gives it the pen's line (both cells beside it), as PowerPoint. */
  penCellLine(id: number, r: number, c: number, side: "l" | "r" | "t" | "b"): void {
    const el = this.editor.shapeEl(id);
    const tbl = el?.getElementsByTagNameNS(PptxDom.NS_A, "tbl")[0];
    if (!tbl) return;
    const pen = this.tablePen;
    void import("./contextcmds").then(({ setCellLine }) => this.editor.edit("Tabelle zeichnen", () => {
      const rows = kids(tbl, "tr");
      const tcAt = (rr: number, cc: number) => kids(rows[rr], "tc")[cc];
      const prOf = (tc: Element) => { let pr = kid(tc, "tcPr"); if (!pr) { pr = make(tc.ownerDocument, "a:tcPr"); tc.append(pr); } return pr; };
      const TAG = { l: "lnL", r: "lnR", t: "lnT", b: "lnB" }, OPP = { l: "lnR", r: "lnL", t: "lnB", b: "lnT" };
      const color = pen.dash === null ? null : pen.color;
      const opts = { w: pen.w, dash: pen.dash ?? "solid" };
      const here = tcAt(r, c);
      if (here) setCellLine(prOf(here), TAG[side], color, opts);
      const nr = side === "t" ? r - 1 : side === "b" ? r + 1 : r, nc = side === "l" ? c - 1 : side === "r" ? c + 1 : c;
      const there = rows[nr] ? tcAt(nr, nc) : undefined;
      if (there) setCellLine(prOf(there), OPP[side], color, opts);
    }));
  }

  /** Tabelle zeichnen outside a table: one cell in the drawn box. */
  drawTable(box: { x: number; y: number; w: number; h: number }): void {
    const slide = this.slidePath();
    if (!slide || box.w < 36000 || box.h < 36000) return;
    let id = 0;
    this.editor.edit("Tabelle zeichnen", () => {
      this.ensureTableStyle();
      const doc = this.deck.pptx.doc(slide);
      const tree = S.treeOf(doc);
      id = S.maxId(tree) + 1;
      const el = S.newTable(doc, id, 1, 1, { x: box.x, y: box.y, w: box.w }, DEFAULT_TABLE_STYLE_ID);
      const h = String(Math.round(box.h));
      for (const tr of Array.from(el.getElementsByTagNameNS(PptxDom.NS_A, "tr"))) tr.setAttribute("h", h);
      const ext = el.getElementsByTagNameNS(PptxDom.NS_A, "ext")[0];
      ext?.setAttribute("cy", h);
      tree.append(el);
    });
    this.editor.select([id]);
  }

  ensureTableStyle(): void {
    const p = this.deck.pptx.related(this.deck.pptx.presPath, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/tableStyles");
    if (!p) return;
    const doc = this.deck.pptx.doc(p);
    if (kids(doc.documentElement, "tblStyle").some((s) => s.getAttribute("styleId") === DEFAULT_TABLE_STYLE_ID)) return;
    doc.documentElement.append(...fragment(doc, DEFAULT_TABLE_STYLE_XML));
    this.deck.pptx.touch(p);
  }

  /** Einfügen > Kopf- und Fußzeile: puts the layout's date, footer and number placeholders on the slides. */
  private async headerFooter(): Promise<void> {
    const date = h("input", { type: "checkbox" }) as HTMLInputElement;
    const number = h("input", { type: "checkbox" }) as HTMLInputElement;
    const footer = h("input", { type: "checkbox" }) as HTMLInputElement;
    const footerText = h("input", { type: "text", style: "width:260px" }) as HTMLInputElement;
    const notOnTitle = h("input", { type: "checkbox" }) as HTMLInputElement;
    const body = h("div", { class: "hf-dialog" },
      h("div", { class: "dlg-tabs" }, h("span", { class: "on" }, "Folie"), h("span", null, "Notizen und Handzettel")),
      h("p", null, h("b", null, "Auf Folie einbeziehen")),
      h("label", null, date, " Datum und Uhrzeit"), h("label", null, number, " Foliennummer"), h("label", null, footer, " Fußzeile"), footerText,
      h("label", null, notOnTitle, " Auf Titelfolie nicht anzeigen"));
    const r = await showDialog({ title: "Kopf- und Fußzeile", body, buttons: [{ label: "Übernehmen", value: "one" }, { label: "Für alle übernehmen", value: "all", primary: true }, { label: "Abbrechen", value: "cancel" }], cancel: "cancel" });
    if (r === "cancel") return;
    const indices = r === "all" ? this.deck.slides.map((_, i) => i) : [...this.panel.selected];
    const paths = indices.map((i) => this.deck.slides[i].path);
    this.deck.change("Kopf- und Fußzeile", paths, () => {
      for (const p of paths) {
        const isTitle = this.deck.pptx.doc(this.deck.pptx.layoutOf(p)!).documentElement.getAttribute("type") === "title";
        const want: Record<string, boolean> = { dt: date.checked, sldNum: number.checked, ftr: footer.checked };
        const tree = S.treeOf(this.deck.pptx.doc(p));
        const layoutTree = S.treeOf(this.deck.pptx.doc(this.deck.pptx.layoutOf(p)!));
        for (const type of ["dt", "ftr", "sldNum"]) {
          const existing = S.shapeEls(tree).find((e) => at(e, "nvSpPr/nvPr/ph")?.getAttribute("type") === type);
          const on = want[type] && !(notOnTitle.checked && isTitle);
          if (!on) { existing?.remove(); continue; }
          if (existing) { if (type === "ftr") T.replaceAll(S.txBodyOf(existing)!, footerText.value); continue; }
          const lph = S.shapeEls(layoutTree).find((e) => at(e, "nvSpPr/nvPr/ph")?.getAttribute("type") === type);
          if (!lph) continue;
          const copy = lph.cloneNode(true) as Element;
          const id = S.maxId(tree) + 1;
          S.renumber(copy, () => id);
          const spPr = kid(copy, "spPr");
          if (spPr) while (spPr.firstChild) spPr.firstChild.remove();
          const lst = at(copy, "txBody/lstStyle");
          if (lst) while (lst.firstChild) lst.firstChild.remove();
          if (type === "ftr") T.replaceAll(S.txBodyOf(copy)!, footerText.value);
          tree.append(this.deck.pptx.doc(p).importNode(copy, true));
        }
      }
    });
  }

  // ----- designs ------------------------------------------------------------------------------------------------

  /** Entwurf > Farben: one of Office's colour sets into the theme (a:clrScheme). */
  applyColorSet(name: string, colors: Record<string, string>): void {
    const theme = this.deck.pptx.themeOf(this.deck.pptx.masters()[0] ?? "");
    if (!theme) return;
    this.deck.change("Farben", [theme], () => {
      const scheme = this.deck.pptx.doc(theme).getElementsByTagNameNS(PptxDom.NS_A, "clrScheme")[0];
      if (!scheme) return;
      scheme.setAttribute("name", name);
      for (const c of kids(scheme)) {
        const hex = colors[c.localName];
        if (!hex) continue;
        while (c.firstChild) c.firstChild.remove();
        c.append(c.localName === "dk1" ? make(c.ownerDocument, "a:sysClr", { val: "windowText", lastClr: hex }) : c.localName === "lt1" ? make(c.ownerDocument, "a:sysClr", { val: "window", lastClr: hex }) : make(c.ownerDocument, "a:srgbClr", { val: hex }));
      }
    });
  }

  /** Entwurf > Schriftarten: the theme's heading and body fonts. */
  applyFontSet(name: string, major: string, minor: string): void {
    const theme = this.deck.pptx.themeOf(this.deck.pptx.masters()[0] ?? "");
    if (!theme) return;
    this.deck.change("Schriftarten", [theme], () => {
      const doc = this.deck.pptx.doc(theme);
      const fs = doc.getElementsByTagNameNS(PptxDom.NS_A, "fontScheme")[0];
      fs?.setAttribute("name", name);
      for (const [tag, face] of [["majorFont", major], ["minorFont", minor]] as const) doc.getElementsByTagNameNS(PptxDom.NS_A, tag)[0]?.getElementsByTagNameNS(PptxDom.NS_A, "latin")[0]?.setAttribute("typeface", face);
    });
  }

  /** Entwurf > Designs: the theme's colours and fonts and the master's design shapes of `d`. */
  applyDesign(d: DesignSpec): void {
    const pptx = this.deck.pptx;
    const master = pptx.masters()[0];
    const theme = master ? pptx.themeOf(master) : undefined;
    if (!master || !theme) return;
    import("@raabe/ooxml").then(({ designTheme, masterXml }) => {
      this.deck.change("Design", [theme, master], () => {
        pptx.put(theme, PptxDom.parseXml(designTheme(d)));
        // The master's design shapes (non-placeholders) and background are replaced; placeholders stay.
        const mdoc = pptx.doc(master);
        const fresh = PptxDom.parseXml(masterXml(d));
        const tree = S.treeOf(mdoc);
        for (const e of S.shapeEls(tree)) if (!at(e, "nvSpPr/nvPr/ph")) e.remove();
        const decos = S.shapeEls(S.treeOf(fresh)).filter((e) => !at(e, "nvSpPr/nvPr/ph"));
        const firstPh = S.shapeEls(tree)[0] ?? null;
        for (const e of decos) tree.insertBefore(mdoc.importNode(e, true), firstPh);
        const bg = at(mdoc.documentElement, "cSld/bg"), nbg = at(fresh.documentElement, "cSld/bg");
        if (bg && nbg) bg.replaceWith(mdoc.importNode(nbg, true));
        const cm = kid(mdoc.documentElement, "clrMap"), ncm = kid(fresh.documentElement, "clrMap");
        if (cm && ncm) cm.replaceWith(mdoc.importNode(ncm, true));
        // Title colour of the design.
        const ts = at(mdoc.documentElement, "txStyles/titleStyle"), nts = at(fresh.documentElement, "txStyles/titleStyle");
        if (ts && nts) ts.replaceWith(mdoc.importNode(nts, true));
      });
    });
  }

  // ----- views ------------------------------------------------------------------------------------------------

  setView(v: View): void {
    this.view = v;
    document.body.dataset.view = v;
    this.center.hidden = v === "sorter" || v === "notesPage";
    this.panel.el.hidden = v === "sorter" || v === "reading" || v === "outline" || v === "notesPage";
    this.outline.el.hidden = v !== "outline";
    this.notesPage.el.hidden = v !== "notesPage";
    this.sorter.el.hidden = v !== "sorter";
    if (v === "sorter") this.sorter.render();
    if (v === "reading") { void this.startShow(this.active.current, true); this.view = "normal"; document.body.dataset.view = "normal"; this.center.hidden = false; this.panel.el.hidden = false; }
    this.renderAll();
    this.focusMain();
  }

  private setZoom(z: number): void {
    if (this.view === "sorter") { this.sorter.setZoom(z); return; }
    this.editor.setZoom(z);
    this.updateChrome();
  }

  private fitWindow(): void {
    this.editor.fit = true;
    this.editor.fitZoom();
    this.updateChrome();
  }

  async startShow(from: number, reading = false): Promise<void> {
    if (this.show) return;
    this.editor.endText();
    this.show = new SlideShow(this, from, reading);
    await this.show.done;
    this.active.current = Math.min(this.show.lastIndex, this.deck.slides.length - 1);
    this.show = null;
    this.renderAll();
    this.focusMain();
  }

  // ----- contextual tabs ------------------------------------------------------------------------------------

  private ctxKey = "";
  private updateContextual(): void {
    const views = this.editor.selectedViews;
    const kinds = new Set(views.map((v) => (v.kind === "graphicFrame" ? v.frame : v.kind)));
    const key = [...kinds].sort().join(",") + (this.editor.text ? "+t" : "") + (this.masterMode ? `+m${this.masterKind}` : "") + (selectedEquation(this) !== null ? "+eq" : "");
    if (key === this.ctxKey) return;
    this.ctxKey = key;
    const tabs: Tab[] = [];
    if (kinds.has("sp") || kinds.has("cxnSp") || kinds.has("grpSp") || this.editor.text) tabs.push(this.shapeTab());
    if (kinds.has("pic")) tabs.push(this.pictureTab());
    if (kinds.has("table")) tabs.push(this.tableTab(), this.tableLayoutTab());
    if (kinds.has("chart")) tabs.push(this.chartTab(), this.chartFormatTab());
    if (selectedEquation(this) !== null) tabs.push(this.cachedTab("equation", () => equationTab(this)));
    if (this.masterMode) tabs.unshift(this.masterTab());
    this.ribbon.setContextual(tabs);
    if (this.masterMode && !kinds.size) this.ribbon.select(this.masterKind === "slides" ? "slideMaster" : `${this.masterKind}Master`);
  }

  /**
   * The master views' tabs as PowerPoint's recorded ones (key O, the same keytips): Folienmaster,
   * Notizenmaster, Handzettelmaster.
   */
  masterTab(): Tab {
    const kind = this.masterKind;
    // eslint-disable-next-line @typescript-eslint/no-this-alias
    const app = this;
    return this.cachedTab(`master-${kind}`, () => {
      const b = (cmd: string, key: string, size: "large" | "small" = "large", label?: string): Control => ({ type: "button", cmd, key, size, label });
      const check = (cmd: string, key: string): Control => ({ type: "check", cmd, key });
      const ctx = this.ribbonContext();
      const slideSize: Control = { type: "dropdown", cmd: "design.slideSize", key: "Ö", size: "large", menu: [{ cmd: "design.slide4x3" }, { cmd: "design.slide16x9" }, "-", { cmd: "design.slideCustom" }] };
      const design: Group = { id: "edit2", label: "Design bearbeiten", key: "ZD", icon: "themes", controls: [{ type: "dropdown", cmd: "design.themes", key: "G", size: "large", menu: [], popup: (c) => this.themesGallery(c) }] };
      const background: Group = { id: "bg", label: "Hintergrund", key: "ZH", icon: "page-color", controls: [
        { type: "stack", items: [
          { type: "dropdown", cmd: "design.colors", key: "F", size: "small", get menu() { return ctx.menus.colors(); } } as Control,
          { type: "dropdown", cmd: "design.fonts", key: "6", size: "small", get menu() { return ctx.menus.fonts(); } } as Control,
          b("design.effects", "E", "small")] },
        { type: "stack", items: [{ type: "dropdown", cmd: "design.background", key: "R", size: "small", label: "Hintergrundformate", get menu() { return bgStylesMenu(app); } } as Control, check("master.hideBg", "B")] },
        b("master.formatBackground", "H", "small")] };
      const close: Group = { id: "close", label: "Schließen", key: "ZS", icon: "close", controls: [b("master.close", "Q")] };
      const orient = (cmd: string, key: string): Control => ({ type: "dropdown", cmd, key, size: "large", menu: [{ label: "Hochformat", run: () => this.setNotesOrientation(true) }, { label: "Querformat", run: () => this.setNotesOrientation(false) }] });
      if (kind === "notes") return {
        id: "notesMaster", label: "Notizenmaster", key: "O", contextual: true, shrink: [],
        groups: [
          { id: "page", label: "Seite einrichten", key: "ZE", icon: "page-size", controls: [orient("notesMaster.orientation", "N"), slideSize] },
          { id: "ph", label: "Platzhalter", key: "ZP", icon: "placeholder-source", controls: [{ type: "stack", items: [check("masterPh.hdr", "K"), check("masterPh.sldImg", "O"), check("masterPh.ftr", "U")] }, { type: "stack", items: [check("masterPh.dt", "D"), check("masterPh.body", "T"), check("masterPh.sldNum", "I")] }] },
          design, background, close],
      };
      if (kind === "handout") return {
        id: "handoutMaster", label: "Handzettelmaster", key: "O", contextual: true, shrink: [],
        groups: [
          { id: "page", label: "Seite einrichten", key: "ZE", icon: "page-size", controls: [orient("handoutMaster.orientation", "A"), slideSize,
            { type: "dropdown", cmd: "handoutMaster.perPage", key: "O", size: "large", get menu(): MenuEntry[] { return ([[1, "1 Folie"], [2, "2 Folien"], [3, "3 Folien"], [4, "4 Folien"], [6, "6 Folien"], [9, "9 Folien"], [0, "Foliengliederung"]] as const).map(([n, label]) => ({ label, checked: app.handoutPerPage === n, run: () => { app.handoutPerPage = n; app.editor.render(); } })); } } as Control] },
          { id: "ph", label: "Platzhalter", key: "ZP", icon: "placeholder-source", controls: [{ type: "stack", items: [check("masterPh.hdr", "K"), check("masterPh.ftr", "U")] }, { type: "stack", items: [check("masterPh.dt", "D"), check("masterPh.sldNum", "I")] }] },
          design, background, close],
      };
      return {
        id: "slideMaster", label: "Folienmaster", key: "O", contextual: true, shrink: [],
        groups: [
          { id: "edit", label: "Master bearbeiten", key: "ZM", icon: "slide-layout", controls: [b("master.insertMaster", "N"), b("master.insertLayout", "L"), { type: "stack", items: [b("master.delete", "DL", "small"), b("master.rename", "A", "small"), b("master.preserve", "M", "small")] }] },
          { id: "layout", label: "Masterlayout", key: "ZL", icon: "slide-layout", controls: [b("master.masterLayout", "O"), b("master.insertPlaceholder", "P"), { type: "stack", items: [check("master.title", "T"), check("master.footers", "U")] }] },
          design, background,
          { id: "size", label: "Größe", key: "ZG", icon: "page-size", controls: [slideSize] },
          close],
      };
    });
  }

  /** The Handzettelmaster's slide places for Folien pro Seite (our drawing of PowerPoint's dashed frames). */
  private handoutDecor(size: { cx: number; cy: number }): HTMLElement {
    const el = h("div", { class: "handout-decor" });
    const px = (v: number) => `${v / EMU_PX}px`;
    const n = this.handoutPerPage, s = this.deck.pptx.slideSize, aspect = s.cx / s.cy;
    const ax = size.cx * 0.08, ay = size.cy * 0.09, aw = size.cx * 0.84, ah = size.cy * 0.82;
    const box = (x: number, y: number, w: number, hh: number, cls = "") => { const d = h("div", cls ? { class: cls } : null); d.style.cssText = `left:${px(x)};top:${px(y)};width:${px(w)};height:${px(hh)}`; el.append(d); };
    if (n === 0) { box(ax, ay, aw, ah); return el; }
    const cols = n >= 4 ? (n === 9 ? 3 : 2) : 1, rows = Math.ceil(n / cols);
    const cellW = (n === 3 ? aw * 0.45 : aw) / cols, cellH = ah / rows;
    for (let i = 0; i < n; i++) {
      const c = i % cols, r = Math.floor(i / cols);
      let w = cellW * 0.9, hh = w / aspect;
      if (hh > cellH * 0.9) { hh = cellH * 0.9; w = hh * aspect; }
      const x = ax + c * cellW + (cellW - w) / 2, y = ay + r * cellH + (cellH - hh) / 2;
      box(x, y, w, hh);
      // 3 Folien: lines for notes beside each slide.
      if (n === 3) for (let k = 0; k < 6; k++) box(ax + aw * 0.5, y + (hh / 6) * k, aw * 0.5, hh / 6, "lines");
    }
    return el;
  }

  /** Notizenseitenausrichtung / Handzettelausrichtung: the notes page size (shared by both) turned, the masters' placeholders scaled. */
  private setNotesOrientation(portrait: boolean): void {
    const pres = this.deck.pptx.presentation;
    const sz = PptxDom.at(pres, "notesSz");
    const cx = PptxDom.num(sz, "cx") ?? 6858000, cy = PptxDom.num(sz, "cy") ?? 9144000;
    if (!sz || (cx < cy) === portrait) return;
    const masters = [this.deck.ensureNotesMasterPath(), this.deck.ensureHandoutMasterPath()].filter((p) => this.deck.pptx.has(p));
    this.deck.change(portrait ? "Hochformat" : "Querformat", [this.deck.pptx.presPath, ...masters], () => {
      sz.setAttribute("cx", String(cy)); sz.setAttribute("cy", String(cx));
      const fx = cy / cx, fy = cx / cy;
      for (const m of masters) for (const x of Array.from(this.deck.pptx.doc(m).getElementsByTagNameNS(PptxDom.NS_A, "xfrm"))) {
        const off = PptxDom.kid(x, "off"), ext = PptxDom.kid(x, "ext");
        if (off) { off.setAttribute("x", String(Math.round((PptxDom.num(off, "x") ?? 0) * fx))); off.setAttribute("y", String(Math.round((PptxDom.num(off, "y") ?? 0) * fy))); }
        if (ext) { ext.setAttribute("cx", String(Math.round((PptxDom.num(ext, "cx") ?? 0) * fx))); ext.setAttribute("cy", String(Math.round((PptxDom.num(ext, "cy") ?? 0) * fy))); }
      }
    });
    this.editor.fitZoom();
    this.renderAll();
  }

  /** Platzhalter checkboxes of the notes and handout masters: the placeholder of that type is there or not. */
  private masterPhOn(type: string): boolean {
    const path = this.masterList()[0];
    if (!path || !this.deck.pptx.has(path)) return false;
    return Array.from(this.deck.pptx.doc(path).getElementsByTagNameNS(PptxDom.NS_P, "ph")).some((p) => p.getAttribute("type") === type);
  }

  private toggleMasterPh(type: string): void {
    const path = this.masterList()[0];
    const kind = this.masterKind;
    if (!path || kind === "slides") return;
    const on = this.masterPhOn(type);
    this.deck.change(on ? "Platzhalter entfernen" : "Platzhalter einfügen", [path], () => {
      const doc = this.deck.pptx.doc(path);
      if (on) { for (const p of Array.from(doc.getElementsByTagNameNS(PptxDom.NS_P, "ph"))) if (p.getAttribute("type") === type) p.closest("sp")?.remove(); return; }
      const xml = defaultMasterPlaceholder(kind, type);
      const tree = doc.getElementsByTagNameNS(PptxDom.NS_P, "spTree")[0];
      if (!xml || !tree) return;
      const sp = PptxDom.fragment(doc, xml)[0];
      const used = new Set(Array.from(doc.getElementsByTagNameNS(PptxDom.NS_P, "cNvPr")).map((c) => c.getAttribute("id")));
      const c = sp.getElementsByTagNameNS(PptxDom.NS_P, "cNvPr")[0];
      let id = 2; while (used.has(String(id))) id++;
      c?.setAttribute("id", String(id));
      tree.append(sp);
    });
    this.editor.render();
  }

  private ctxHost: import("./contexttabs").ContextTabHost | null = null;
  private host2(): import("./contexttabs").ContextTabHost { return this.ctxHost ??= contextHost(this); }
  private tabCache = new Map<string, Tab>();
  private cachedTab(id: string, make: () => Tab): Tab {
    let t = this.tabCache.get(id);
    if (!t) { t = this.ribbonLabels([make()])[0]; this.tabCache.set(id, t); }
    return t;
  }
  shapeTab(): Tab { return this.cachedTab("shape", () => CT.shapeFormatTab(this.host2())); }
  pictureTab(): Tab { return this.cachedTab("pic", () => CT.pictureFormatTab(this.host2())); }
  tableTab(): Tab { return this.cachedTab("tableD", () => CT.tableDesignTab(this.host2())); }
  tableLayoutTab(): Tab { return this.cachedTab("tableL", () => CT.tableLayoutTab(this.host2())); }
  chartTab(): Tab { return this.cachedTab("chartD", () => CT.chartDesignTab(this.host2())); }
  chartFormatTab(): Tab { return this.cachedTab("chartF", () => CT.chartFormatTab(this.host2())); }

  /** Linienstärke for the selected shapes (Formkontur > Stärke, Stiftstärke). */
  setLineWidth(pt: number): void { this.setLine(undefined as unknown as string, pt); }

  private effectsMenu(): MenuEntry[] {
    return shapeEffectsMenu(this);
  }











  // ----- galleries -------------------------------------------------------------------------------------------

  /** Formformat > Formen einfügen: the shapes as a gallery in the ribbon (three rows, as PowerPoint's). */
  shapeTiles(): Control {
    const items = PRESENT_SHAPES.flatMap((g) => g.items);
    return {
      type: "tiles", cmd: "insert.shapes", key: "2", count: items.length, visible: 18, size: [22, 22], cls: "shape-tiles-ribbon",
      render: (i) => h("span", { class: "shape-ico" }, shapeIcon(items[i])),
      tip: (i) => items[i].label,
      pick: (i) => { const it = items[i]; if (it.tool.kind === "textbox") this.startTool({ kind: "textbox" }); else this.startTool({ kind: "shape", prst: it.tool.geometry, head: !!it.tool.pic?.line?.head, tail: !!it.tool.pic?.line?.tail, freeform: it.tool.freeform }); },
      redraw: () => "shapes",
      panel: (close) => this.shapeGallery(close),
    };
  }

  /** Form bearbeiten > Form ändern: the selected shapes take another geometry (text and look stay); with `pictures` the pictures (Auf Form zuschneiden). */
  changeShape(geometry: string, pictures = false): void {
    const views = this.editor.selectedViews.filter((v) => v.kind === (pictures ? "pic" : "sp"));
    if (!views.length) return;
    this.editor.edit("Form ändern", () => {
      for (const v of views) {
        const sp = kid(this.editor.shapeEl(v.id), "spPr");
        if (!sp) continue;
        for (const g of kids(sp)) if (g.localName === "prstGeom" || g.localName === "custGeom") g.remove();
        const geom = fragment(sp.ownerDocument, `<a:prstGeom prst="${geometry}"><a:avLst/></a:prstGeom>`)[0];
        const xf = kid(sp, "xfrm");
        sp.insertBefore(geom, xf ? xf.nextSibling : sp.firstChild);
      }
    });
  }

  shapeGallery(close: () => void, pick?: (geometry: string) => void): HTMLElement {
    const wrap = h("div", { class: "shape-gallery-wrap" });
    for (const group of PRESENT_SHAPES) {
      wrap.append(h("div", { class: "gal-head" }, group.title));
      const grid = h("div", { class: "shape-gallery" });
      for (const it of group.items) {
        const btn = h("button", { "data-tip": it.label, "aria-label": it.label, onclick: () => {
          close();
          if (pick) { if (it.tool.kind !== "textbox" && it.tool.geometry && !it.tool.freeform) pick(it.tool.geometry); return; }
          if (it.tool.kind === "textbox") this.startTool({ kind: "textbox" });
          else this.startTool({ kind: "shape", prst: it.tool.geometry, head: !!it.tool.pic?.line?.head, tail: !!it.tool.pic?.line?.tail, freeform: it.tool.freeform });
        } });
        btn.append(shapeIcon(it));
        grid.append(btn);
      }
      wrap.append(grid);
    }
    return wrap;
  }

  private startTool(tool: NonNullable<SlideEditor["tool"]>): void {
    this.editor.endText();
    this.editor.tool = tool;
    this.editor.render();
    this.editor.focus();
  }

  private layoutGallery(close: () => void, pick: (path: string) => void, title: string): HTMLElement {
    const wrap = h("div", { class: "layout-gallery" }, h("div", { class: "gal-head" }, title));
    const grid = h("div", { class: "layout-grid" });
    const size = this.deck.pptx.slideSize;
    const tw = 104, s = tw / (size.cx / EMU_PX);
    for (const l of this.deck.layouts()) {
      const thumb = h("div", { class: "layout-thumb", style: `width:${tw}px;height:${(size.cy / EMU_PX) * s}px` });
      const r = renderSlide(this.deck.view(l.path), size, { editing: false, image: (p) => this.image(p), pptx: this.deck.pptx });
      r.el.style.transform = `scale(${s})`; r.el.style.transformOrigin = "0 0";
      thumb.append(r.el);
      grid.append(h("button", { class: "layout-item", "aria-label": l.name, onclick: () => { close(); pick(l.path); } }, thumb, h("span", null, l.name)));
    }
    wrap.append(grid);
    if (title === "Office-Design" || title.endsWith("Design")) {
      wrap.append(h("div", { class: "menu-sep" }), h("button", { class: "menu-item", onclick: () => { close(); this.duplicateSlides([...this.panel.selected]); } }, "Ausgewählte Folien duplizieren"),
        h("button", { class: "menu-item", disabled: true }, "Folien aus Gliederung..."),
        h("button", { class: "menu-item", onclick: () => { close(); void import("./slideimport").then((m) => m.reuseSlidesDialog(this)); } }, "Folien wiederverwenden..."));
    }
    return wrap;
  }

  private tableGrid(close: () => void): HTMLElement {
    const label = h("div", { class: "gal-head" }, "Tabelle einfügen");
    const grid = h("div", { class: "table-grid" });
    const cells: HTMLElement[] = [];
    for (let r = 0; r < 8; r++) for (let c = 0; c < 10; c++) {
      const cell = h("div", { class: "tg-cell", role: "button", "aria-label": `${c + 1}x${r + 1} Tabelle` });
      cell.onmouseenter = () => { label.textContent = `${c + 1}x${r + 1}-Tabelle`; cells.forEach((x, i) => x.classList.toggle("on", Math.floor(i / 10) <= r && i % 10 <= c)); };
      cell.onclick = () => { close(); this.insertTable(r + 1, c + 1); };
      cells.push(cell);
      grid.append(cell);
    }
    return h("div", { class: "table-popup" }, label, grid, h("div", { class: "menu-sep" }),
      h("button", { class: "menu-item", onclick: () => { close(); void this.tableDialog(); } }, icon("table"), " Tabelle einfügen..."),
      // Tabelle zeichnen: the table pen (Tabellenentwurf's), as in PowerPoint's menu.
      h("button", { class: "menu-item", onclick: () => { close(); void this.commands.run("table.draw"); } }, icon("table-draw"), " Tabelle zeichnen"));
  }

  private async tableDialog(): Promise<void> {
    const cols = h("input", { type: "number", value: "5", min: "1", max: "75", style: "width:60px" }) as HTMLInputElement;
    const rows = h("input", { type: "number", value: "2", min: "1", max: "75", style: "width:60px" }) as HTMLInputElement;
    const r = await showDialog({ title: "Tabelle einfügen", body: h("div", { class: "form" }, h("label", null, "Spaltenanzahl: ", cols), h("label", null, "Zeilenanzahl: ", rows)), buttons: [{ label: "OK", value: true, primary: true }, { label: "Abbrechen", value: false }], cancel: false });
    if (r) this.insertTable(Math.max(1, Math.min(75, Number(rows.value) || 2)), Math.max(1, Math.min(75, Number(cols.value) || 5)));
  }

  /**
   * PowerPoint's ribbon shows no "..." on its buttons ("Zoom", "Kopf- und Fußzeile"), although the commands
   * (and their accessible names) have it (UI screenshot, D-072): the visible labels lose it.
   */
  private ribbonLabels(tabs: Tab[]): Tab[] {
    const fix = (c: Control): void => {
      if (c.type === "stack") { c.items.forEach(fix); return; }
      if (c.type === "rows") { c.rows.flat().forEach((x) => { if (x !== "gap") fix(x); }); return; }
      if ((c.type === "button" || c.type === "dropdown" || c.type === "split") && !c.label && this.commands.has(c.cmd)) {
        const l = this.commands.get(c.cmd).label;
        if (l.endsWith("...")) (c as { label?: string }).label = l.slice(0, -3);
      }
    };
    for (const t of tabs) for (const g of t.groups) g.controls.forEach(fix);
    return tabs;
  }

  /** Entwurf > Designs in the ribbon: each design's title slide, small; the arrow opens the full gallery. */
  private designThumbs = new Map<number, HTMLElement>();
  private designTiles(): Control {
    const thumb = (i: number): HTMLElement => {
      const d = DESIGNS[i];
      const sample = new Deck(newPptx(d));
      const size = sample.pptx.slideSize;
      const tw = 80, s = tw / (size.cx / EMU_PX);
      const box = h("div", { class: "design-tile", style: `width:${tw}px;height:${Math.round((size.cy / EMU_PX) * s)}px` });
      const r = renderSlide(sample.view(sample.slides[0].path), size, { editing: false, image: () => undefined, pptx: sample.pptx });
      r.el.style.transform = `scale(${s})`; r.el.style.transformOrigin = "0 0";
      box.append(r.el, h("div", { class: "theme-label", style: `color:#${d.colors.accent1}` }, "Aa"));
      return box;
    };
    return {
      type: "tiles", cmd: "design.themes", key: "G", count: DESIGNS.length, visible: DESIGNS.length, size: [84, 50], cls: "design-tiles",
      render: (i) => { let t = this.designThumbs.get(i); if (!t) { t = thumb(i); this.designThumbs.set(i, t); } return t.cloneNode(true) as HTMLElement; },
      tip: (i) => (DESIGNS[i].name === "Office" ? "Office-Design" : DESIGNS[i].name),
      pick: (i) => this.applyDesign(DESIGNS[i]),
      current: () => { const n = this.designName(); const i = DESIGNS.findIndex((d) => d.name === n); return i < 0 ? undefined : i; },
      redraw: () => "designs",
      panel: (close) => this.themesGallery(close),
    };
  }

  /** The effects in the ribbon: Übergänge (all named transitions), Animationen (Ohne and the entrance, emphasis and exit effects). */
  private fxTiles(kind: "trans" | "anim"): Control {
    type Fx = { label: string; run: () => void };
    const items: Fx[] = kind === "trans"
      ? TRANSITION_GROUPS.flatMap((g) => g.items.map(([label, k]) => ({ label, run: () => void import("./transitions").then((m) => m.setTransition(this, k)) })))
      : [{ label: "Ohne", run: () => void import("./animations").then((m) => { const p = this.slidePath(); const root = p ? this.deck.pptx.doc(p).documentElement : null; const timing = root && kid(root, "timing"); if (timing) this.editor.edit("Animation entfernen", () => { for (const id of this.editor.sel) m.removeEffects(timing, id); m.tidyTiming(root!); }); }) },
        ...ANIMATION_GROUPS.flatMap((g) => g.items.filter((it) => g.cls !== "path" || it.preset !== 0).map((it) => ({ label: it.label, run: () => void import("./animations").then((m) => m.addAnimation(this, g.cls, it, false)) })))];
    return {
      type: "tiles", cmd: kind === "trans" ? "trans.gallery" : "anim.gallery", key: kind === "trans" ? "Ü" : "N", count: items.length, visible: 6, size: [62, 58], cls: "fx-tiles",
      render: (i) => h("div", { class: "pfx-tile" }, icon(i === 0 ? "close-x" : kind === "trans" ? "transition" : "animation"), h("span", null, items[i].label)),
      tip: (i) => items[i].label,
      pick: (i) => items[i].run(),
      redraw: () => kind,
      panel: (close) => (kind === "trans" ? this.transitionGallery(close) : this.animationGallery(close, false)),
    };
  }

  private themesGallery(close: () => void): HTMLElement {
    const wrap = h("div", { class: "theme-gallery" }, h("div", { class: "gal-head" }, "Office"));
    const grid = h("div", { class: "layout-grid" });
    for (const d of DESIGNS) {
      const sample = new Deck(newPptx(d));
      const size = sample.pptx.slideSize;
      const tw = 120, s = tw / (size.cx / EMU_PX);
      const thumb = h("div", { class: "layout-thumb", style: `width:${tw}px;height:${(size.cy / EMU_PX) * s}px` });
      const r = renderSlide(sample.view(sample.slides[0].path), size, { editing: false, image: () => undefined, pptx: sample.pptx });
      r.el.style.transform = `scale(${s})`; r.el.style.transformOrigin = "0 0";
      const label = h("div", { class: "theme-label", style: `color:#${d.colors.accent1}` }, "Aa");
      thumb.append(r.el, label);
      grid.append(h("button", { class: "layout-item", "data-tip": d.name === "Office" ? "Office-Design" : d.name, "aria-label": d.name, onclick: () => { close(); this.applyDesign(d); } }, thumb, h("span", null, d.name === "Office" ? "Office-Design" : d.name)));
    }
    wrap.append(grid);
    return wrap;
  }

  // ----- context menus --------------------------------------------------------------------------------------

  /** Context menus go through here, so the inventory can collect their entries without showing them. */
  private capture: MenuEntry[][] | null = null;
  private menu(entries: MenuEntry[], at: DOMRect, before?: HTMLElement): void {
    if (this.capture) { this.capture.push(entries); return; }
    showMenu(this.commands, entries, at, before);
  }

  /** PowerPoint's Minileiste above the context menu (recorded): text formatting, or the shape's look. */
  private miniBar(kind: "text" | "shape"): HTMLElement | undefined {
    const ctx = this.ribbonContext();
    const b = (cmd: string) => ({ type: "button" as const, cmd, key: "", size: "small" as const });
    const controls: import("@raabe/shell").Control[] = kind === "text"
      ? [{ type: "rows", rows: [
        [{ type: "combo", cmd: "font.name", width: 104, value: ctx.fontName, options: ctx.fontOptions, commit: ctx.setFont }, { type: "combo", cmd: "font.size", width: 44, value: ctx.fontSize, options: () => [{ items: ["8", "10", "12", "14", "18", "24", "28", "32", "36", "44", "54", "60", "72"] }], commit: ctx.setSize },
          b("font.grow"), b("font.shrink"), b("para.levelDown"), b("para.levelUp")],
        [b("font.bold"), b("font.italic"), b("font.underline"), b("para.left"), b("para.center"), b("para.right"),
          { type: "color", cmd: "font.highlight", palette: "highlight", current: ctx.highlight, pick: ctx.setHighlight, names: PP_HIGHLIGHT_NAMES }, { type: "color", cmd: "font.color", palette: "font", current: ctx.fontColor, pick: ctx.setFontColor, theme: () => this.paletteTheme() }],
      ] }]
      : [{ type: "color", cmd: "draw.fill", palette: "font", current: ctx.fillColor, pick: ctx.setFill, showLabel: true, theme: () => this.paletteTheme(), extra: () => this.fillExtra() }, { type: "color", cmd: "draw.outline", palette: "font", current: ctx.lineColor, pick: ctx.setLine, showLabel: true, extra: () => this.outlineExtra() , theme: () => this.paletteTheme() },
        { type: "button", cmd: "review.newComment", key: "", size: "large" }];
    return this.ribbon.toolbar(controls).el;
  }

  /** Context menus as PowerPoint's (recorded: results/2026-11-06-powerpoint-inventory.json); Übersetzen is left out (cloud). */
  contextMenu(e: MouseEvent, kind: "shape" | "text" | "slide"): void {
    const at2 = new DOMRect(e.clientX, e.clientY, 0, 0);
    const paste = (tip: string): MenuEntry => ({ label: "Einfügeoptionen:", icons: [{ icon: "paste-keep", tip, run: () => void this.paste() }, { icon: "paste-text", tip: "Nur den Text übernehmen (T)", run: () => void this.paste("text") }] });
    const order: MenuEntry[] = [{ label: "In den Vordergrund", submenu: [{ cmd: "arrange.front" }, { cmd: "arrange.forward" }] }, { label: "In den Hintergrund", submenu: [{ cmd: "arrange.back" }, { cmd: "arrange.backward" }] }];
    const group: MenuEntry = { label: "Gruppieren", submenu: [{ cmd: "arrange.group" }, { cmd: "arrange.regroup" }, { cmd: "arrange.ungroup" }] };
    if (kind === "text") {
      const inTable = !!this.editor.text?.cell;
      if (inTable) {
        this.menu([{ cmd: "edit.cut" }, { cmd: "edit.copy" }, paste("Ursprüngliche Formatierung beibehalten (U)"), "-", { cmd: "font.dialog" },
          { label: "Aufzählungszeichen", submenu: this.bulletMenu() }, { label: "Nummerierung", submenu: this.numberMenu() }, "-", { cmd: "shape.lock" }, { cmd: "insert.link", label: "Link" }, { cmd: "review.synonyms" }, "-",
          { label: "Einfügen", submenu: [{ cmd: "table.insertAbove" }, { cmd: "table.insertBelow" }, { cmd: "table.insertLeft" }, { cmd: "table.insertRight" }] },
          { label: "Löschen", submenu: [{ cmd: "table.deleteRows" }, { cmd: "table.deleteCols" }, { cmd: "table.delete" }] },
          { cmd: "table.merge" }, { cmd: "table.split" }, { cmd: "table.selectAll" }, "-", { cmd: "draw.format", label: "Form formatieren..." }, { cmd: "review.newComment" }], at2, this.miniBar("text"));
        return;
      }
      this.menu([{ cmd: "edit.cut" }, { cmd: "edit.copy" }, paste("Ursprüngliche Formatierung beibehalten (U)"), "-",
        { cmd: "text.endEdit" }, "-", { cmd: "font.dialog" }, { cmd: "para.dialog" }, { label: "Aufzählungszeichen", submenu: this.bulletMenu() }, { label: "Nummerierung", submenu: this.numberMenu() },
        { cmd: "para.smartArt" }, "-", { cmd: "shape.lock" }, { cmd: "insert.link", label: "Link" }, { cmd: "review.synonyms" }, "-",
        { cmd: "text.effectsDialog" }, { cmd: "draw.format", label: "Form formatieren..." }, { cmd: "review.newComment" }], at2, this.miniBar("text"));
      return;
    }
    if (kind === "shape") {
      const v = this.editor.selectedViews[0];
      const entries: MenuEntry[] = [{ cmd: "edit.cut" }, { cmd: "edit.copy" }, paste("Ursprüngliche Formatierung beibehalten (U)"), "-"];
      if (v?.kind === "graphicFrame" && v.frame === "chart") {
        this.menu([{ cmd: "edit.delete" }, { cmd: "chart.reset" }, "-", { cmd: "font.dialog" }, { cmd: "chart.changeType" }, { cmd: "chart.editData" }, "-", { cmd: "draw.format", label: "Diagrammbereich formatieren..." }, { cmd: "review.newComment" }], at2);
        return;
      }
      if (v?.kind === "pic") entries.push({ cmd: "pic.change" }, "-", group, ...order, "-", { cmd: "shape.lock" }, { cmd: "shape.link", label: "Link" }, { cmd: "pic.save" }, "-", { cmd: "shape.altText" }, { cmd: "shape.sizePosition" }, { cmd: "pic.format" }, { cmd: "review.newComment" });
      else if (v?.kind === "grpSp") entries.push(group, ...order, "-", { cmd: "shape.lock" }, { cmd: "shape.link", label: "Link" }, { cmd: "pic.save" }, "-", { cmd: "shape.altText" }, { cmd: "shape.sizePosition" }, { cmd: "draw.format", label: "Form formatieren..." }, { cmd: "review.newComment" });
      else entries.push({ cmd: "shape.editText" }, ...(v?.ph ? [] : [{ cmd: "shape.editPoints" } as MenuEntry]), "-", group, ...order, "-", { cmd: "shape.lock" }, { cmd: "shape.link", label: "Link" }, { cmd: "pic.save" }, "-",
        { cmd: "shape.altText" }, ...(v?.ph ? [] : [{ cmd: v?.textBox ? "shape.setDefaultText" : "shape.setDefault" } as MenuEntry]), { cmd: "shape.sizePosition" }, { cmd: "draw.format", label: "Form formatieren..." }, { cmd: "review.newComment" });
      this.menu(entries, at2, v?.kind === "pic" ? undefined : this.miniBar("shape"));
      return;
    }
    this.menu([{ label: "Layout", submenu: () => this.deck.layouts().map((l) => ({ label: l.name.replace("\n", ""), run: () => this.deck.setLayout([this.active.current], l.path) })) },
      { cmd: "slide.reset", label: "Folie zurücksetzen" }, "-", { cmd: "view.gridDialog", label: "Raster und Führungslinien..." }, { cmd: "view.ruler" }, { cmd: "slide.formatBackground" }, { cmd: "review.newComment" }], at2);
  }

  /** The notes pane's context menu (PowerPoint's Notizenbereich; the notes are plain text here). */
  notesMenu(e: MouseEvent): void {
    const ta = this.notes;
    const sel = () => ta.value.slice(ta.selectionStart, ta.selectionEnd);
    const put = (t: string) => { ta.setRangeText(t, ta.selectionStart, ta.selectionEnd, "end"); ta.dispatchEvent(new Event("input")); ta.focus(); };
    const has = !!sel();
    this.menu([
      { label: "Ausschneiden", disabled: !has, run: () => { void this.platform.clipboardWrite(sel()); put(""); } },
      { label: "Kopieren", disabled: !has, run: () => void this.platform.clipboardWrite(sel()) },
      { label: "Einfügeoptionen:", icons: [{ icon: "paste-text", tip: "Nur den Text übernehmen (T)", run: () => void this.platform.clipboardRead().then(put) }] },
      { label: "Einfügen", run: () => void this.platform.clipboardRead().then(put) }, "-",
      { label: "Schriftart...", disabled: true, run: () => {} }, { label: "Absatz...", disabled: true, run: () => {} }, "-",
      { cmd: "review.newComment" }], new DOMRect(e.clientX, e.clientY, 0, 0));
  }

  panelMenu(e: MouseEvent, onSlide: boolean): void {
    const at2 = new DOMRect(e.clientX, e.clientY, 0, 0);
    const paste: MenuEntry = { label: "Einfügeoptionen:", icons: [{ icon: "paste", tip: "Zieldesign verwenden (H)", run: () => this.pasteSlides() }, { icon: "paste-keep", tip: "Ursprüngliche Formatierung beibehalten (U)", run: () => this.pasteSlides() }] };
    const entries: MenuEntry[] = onSlide
      ? [{ cmd: "edit.cut" }, { cmd: "edit.copy" }, paste, "-", { cmd: "slide.new" }, { cmd: "slide.duplicate", label: "Folie duplizieren" }, { cmd: "slide.delete" }, "-", { cmd: "slide.addSection" }, "-",
        { label: "Layout", submenu: () => this.deck.layouts().map((l) => ({ label: l.name.replace("\n", ""), run: () => this.deck.setLayout([...this.panel.selected], l.path) })) },
        { cmd: "slide.reset", label: "Folie zurücksetzen" }, "-", { cmd: "slide.formatBackground" }, { cmd: "insert.photoAlbum" }, { cmd: "show.hideSlide" }, { cmd: "slide.linkTo" }, "-", { cmd: "review.newComment" }]
      : [paste, "-", { cmd: "slide.new" }, { cmd: "slide.addSection" }];
    this.menu(entries, at2);
  }

  private bulletMenu(): MenuEntry[] {
    const chars: [string, string][] = [["Keine", ""], ["Ausgefüllte runde Aufzählungszeichen", "•"], ["Hohle runde Aufzählungszeichen", "o"], ["Ausgefüllte quadratische Aufzählungszeichen", "§"],
      ["Silbentrennung-Aufzählungszeichen", String.fromCharCode(0x2013)], ["Sternförmige Aufzählungszeichen", "v"], ["Pfeilförmige Aufzählungszeichen", "Ø"], ["Häkchenaufzählungszeichen", "ü"]];
    return chars.map(([label, c]) => ({ label, run: () => c ? this.toggleBullets("char", c) : this.formatParas("Aufzählungszeichen", (pPr) => { for (const k of kids(pPr)) if (["buChar", "buAutoNum", "buFont"].includes(k.localName)) k.remove(); if (!kid(pPr, "buNone")) pPr.insertBefore(make(pPr.ownerDocument, "a:buNone"), kids(pPr).find((x) => ["tabLst", "defRPr", "extLst"].includes(x.localName)) ?? null); }) }));
  }

  private numberMenu(): MenuEntry[] {
    const types: [string, string][] = [["1. 2. 3.", "arabicPeriod"], ["1) 2) 3)", "arabicParenR"], ["I. II. III.", "romanUcPeriod"], ["Nummern in Klammern", "arabicParenBoth"], ["a) b) c)", "alphaLcParenR"], ["a. b. c.", "alphaLcPeriod"], ["i. ii. iii.", "romanLcPeriod"]];
    return [{ label: "Keine", run: () => this.toggleBullets("auto", "none") }, ...types.map(([label, t]) => ({ label, run: () => this.toggleBullets("auto", t) })), "-", { label: "Aufzählungszeichen und Nummerierung...", run: () => void this.commands.run("para.dialog") }];
  }

  // ----- commands ------------------------------------------------------------------------------------------

  private implementCommands(): void {
    const c = this.commands;
    const ed = () => this.editor;
    const has = () => this.editor.sel.length > 0 || !!this.editor.text;
    const shapes = () => this.editor.sel.length > 0;
    const text = () => this.hasText();
    const impl = (id: string, run: () => void | Promise<void>, enabled?: () => boolean, checked?: () => boolean) => c.implement(id, { run, enabled, checked });

    impl("file.new", () => this.newDoc());
    impl("file.open", () => this.openDialog());
    impl("file.save", async () => { await this.save(); });
    impl("file.saveAs", async () => { await this.save(this.active, true); });
    impl("file.close", () => this.closeDoc(this.active));
    impl("file.exportPdf", () => this.exportPdf());
    impl("file.print", () => { this.backstage.open("print"); });
    impl("file.options", () => { this.backstage.close(); void this.optionsDialog(); });
    // After Rückgängig the shapes it brings back are selected (PowerPoint: Entf, then Strg+Z).
    const reselect = (step: () => void) => {
      const path = this.slidePath();
      const before = new Set(path ? this.deck.view(path).shapes.map((s) => s.id) : []);
      ed().endText();
      step();
      const now = path && this.slidePath() === path && this.deck.pptx.has(path) ? this.deck.view(path).shapes.map((s) => s.id) : [];
      const back = now.filter((id) => !before.has(id));
      if (back.length) ed().select(back);
    };
    impl("edit.undo", () => reselect(() => this.deck.undo()), () => this.deck.canUndo);
    impl("edit.redo", () => { ed().endText(); this.deck.redo(); }, () => this.deck.canRedo);
    impl("edit.copy", () => this.copy(false), () => has());
    impl("edit.cut", () => this.copy(true), () => has());
    impl("edit.paste", () => this.paste());
    impl("edit.pasteText", () => this.paste("text"));
    impl("edit.pasteKeepSource", () => this.paste());
    impl("edit.pasteDest", () => this.paste());
    impl("edit.duplicate", () => this.duplicateShapes(), shapes);
    impl("edit.formatPainter", () => this.pickFormat(false), () => has(), () => !!this.painter);
    impl("edit.selectAll", () => this.selectAll());
    impl("edit.delete", () => ed().deleteSelected(), shapes);
    impl("edit.selectionPane", () => this.selectionPane());
    impl("arrange.selectionPane", () => this.selectionPane());

    impl("slide.new", () => this.newSlide());
    impl("slide.duplicate", () => this.duplicateSlides([...this.panel.selected].sort((a, b) => a - b)));
    impl("slide.delete", () => this.deleteSlides([...this.panel.selected]));
    impl("slide.reset", () => this.deck.resetSlides([...this.panel.selected]));
    impl("slide.hide", () => this.toggleHidden(), undefined, () => this.isHidden());
    impl("show.hideSlide", () => this.toggleHidden(), undefined, () => this.isHidden());
    impl("slide.addSection", () => this.addSection());
    impl("slide.renameSection", () => this.renameSection(), () => this.deck.sections().length > 0);
    impl("slide.removeSection", () => this.removeSection(false), () => this.deck.sections().length > 0);
    impl("slide.removeAllSections", () => this.removeSection(true), () => this.deck.sections().length > 0);
    impl("slide.collapseAll", () => { for (const s of this.deck.sections()) this.panel.collapsed.add(s.id); this.panel.render(); }, () => this.deck.sections().length > 0);
    impl("slide.expandAll", () => { this.panel.collapsed.clear(); this.panel.render(); }, () => this.deck.sections().length > 0);

    impl("font.bold", () => this.toggleRun("b", "Fett"), text, () => !!this.runStyle()?.bold);
    impl("font.italic", () => this.toggleRun("i", "Kursiv"), text, () => !!this.runStyle()?.italic);
    impl("font.underline", () => { const on = (this.runStyle()?.underline ?? "none") === "none"; this.formatRuns("Unterstrichen", (r) => r.setAttribute("u", on ? "sng" : "none")); }, text, () => (this.runStyle()?.underline ?? "none") !== "none");
    impl("font.strike", () => { const on = (this.runStyle()?.strike ?? "noStrike") === "noStrike"; this.formatRuns("Durchgestrichen", (r) => r.setAttribute("strike", on ? "sngStrike" : "noStrike")); }, text, () => (this.runStyle()?.strike ?? "noStrike") !== "noStrike");
    impl("font.superscript", () => { const on = !(this.runStyle()?.baseline ?? 0 > 0); this.formatRuns("Hochgestellt", (r) => r.setAttribute("baseline", on ? "30000" : "0")); }, text);
    impl("font.subscript", () => { const on = !((this.runStyle()?.baseline ?? 0) < 0); this.formatRuns("Tiefgestellt", (r) => r.setAttribute("baseline", on ? "-25000" : "0")); }, text);
    impl("font.shadow", () => { this.formatRuns("Textschatten", (r) => { const fx = kid(r, "effectLst"); if (fx) fx.remove(); else { const before = kids(r).find((x) => ["highlight", "uLnTx", "uLn", "uFillTx", "uFill", "latin", "ea", "cs", "sym", "hlinkClick", "hlinkMouseOver", "rtl", "extLst"].includes(x.localName)) ?? null; r.insertBefore(fragment(r.ownerDocument, '<a:effectLst><a:outerShdw blurRad="38100" dist="38100" dir="2700000" algn="tl"><a:srgbClr val="000000"><a:alpha val="43137"/></a:srgbClr></a:outerShdw></a:effectLst>')[0], before); } }); }, text);
    impl("font.grow", () => this.stepSize(1), text);
    impl("font.shrink", () => this.stepSize(-1), text);
    impl("font.clear", () => this.formatRuns("Alle Formatierungen löschen", (r) => { for (const a of Array.from(r.attributes)) if (!["lang", "dirty"].includes(a.name)) r.removeAttribute(a.name); while (r.firstChild) r.firstChild.remove(); }), text);
    impl("font.name", () => {}, text);
    impl("font.size", () => {}, text);
    impl("font.color", () => this.formatRuns("Schriftfarbe", (r) => T.setRunColor(r, this.lastFontColor)), text);
    impl("font.highlight", () => this.formatRuns("Texthervorhebungsfarbe", (r) => T.setHighlight(r, this.lastHighlight)), text);
    impl("font.case", () => {}, text);
    impl("font.spacing", () => {}, text);

    impl("para.left", () => this.align("l"), text, () => (this.paraStyle()?.algn ?? "l") === "l");
    impl("para.center", () => this.align("ctr"), text, () => this.paraStyle()?.algn === "ctr");
    impl("para.right", () => this.align("r"), text, () => this.paraStyle()?.algn === "r");
    impl("para.justify", () => this.align("just"), text, () => this.paraStyle()?.algn === "just");
    impl("para.bullets", () => this.toggleBullets("char"), text, () => this.paraStyle()?.bullet?.kind === "char");
    impl("para.numbering", () => this.toggleBullets("auto"), text, () => this.paraStyle()?.bullet?.kind === "auto");
    impl("para.levelUp", () => ed().changeLevel(1), text);
    impl("para.levelDown", () => ed().changeLevel(-1), text);
    impl("para.lineSpacing", () => {}, text);
    impl("para.alignText", () => {}, text);
    impl("para.direction", () => {}, text);
    impl("para.columns", () => {}, text);

    impl("draw.shapes", () => {});
    impl("insert.shapes", () => {});
    impl("insert.chart", () => {}, () => !!this.slidePath());
    impl("insert.photoAlbum", () => import("./photoalbum").then((m) => m.photoAlbumDialog(this)), () => !!this.platform.openImagesDialog);
    impl("view.switchWindow", () => {}, () => this.docs.length > 0);
    impl("view.newWindow", () => this.newWindow(), () => this.docs.length > 0);
    impl("edit.replaceFonts", () => void this.replaceFontsDialog(), () => this.docs.length > 0);
    impl("insert.action", () => import("./links").then((m) => m.actionDialog(this)), () => this.editor.sel.length === 1);
    impl("view.ruler", () => { this.editor.showRuler = !this.editor.showRuler; this.editor.fitZoom(); this.editor.render(); this.updateChrome(); }, () => this.view === "normal", () => this.editor.showRuler);
    impl("draw.arrange", () => {});
    impl("draw.fill", () => this.setFill(this.lastFill), shapes);
    impl("draw.outline", () => this.setLine(this.lastLine), shapes);
    impl("draw.effects", () => {}, shapes);
    impl("draw.quickStyles", () => {}, shapes);
    impl("arrange.front", () => this.order("front"), shapes);
    impl("arrange.back", () => this.order("back"), shapes);
    impl("arrange.forward", () => this.order("forward"), shapes);
    impl("arrange.backward", () => this.order("backward"), shapes);
    impl("arrange.group", () => this.group(), () => this.editor.sel.length > 1);
    impl("arrange.ungroup", () => this.ungroup(), () => this.editor.selectedViews.some((v) => v.kind === "grpSp"));
    // Gruppierung wiederherstellen: the shapes last ungrouped, grouped again (while they are all still there).
    impl("arrange.regroup", () => { this.editor.select(this.lastUngrouped); this.group(); }, () => this.lastUngrouped.length > 1 && this.lastUngrouped.every((id) => !!this.editor.shapeView(id)));
    // Dropdowns whose menus do the work (they were greyed without a command).
    for (const id of ["slide.layout", "slide.section", "edit.select", "design.colors", "design.fonts"]) impl(id, () => {}, () => !!this.slidePath());
    // Folienmaster > Masterlayout: Titel and Fußzeilen on the layout; Master bearbeiten > Beibehalten.
    const layoutPath = () => (this.masterMode && this.masterKind === "slides" && this.masterIndex > 0 ? this.masterList()[this.masterIndex] : undefined);
    const hasPh = (types: string[]) => { const p = layoutPath(); return !!p && Array.from(this.deck.pptx.doc(p).getElementsByTagNameNS(PptxDom.NS_P, "ph")).some((x) => types.includes(x.getAttribute("type") ?? "")); };
    impl("master.title", () => this.toggleLayoutPh(layoutPath(), ["title"]), () => !!layoutPath(), () => hasPh(["title"]));
    impl("master.footers", () => this.toggleLayoutPh(layoutPath(), ["dt", "ftr", "sldNum"]), () => !!layoutPath(), () => hasPh(["dt", "ftr", "sldNum"]));
    const masterRoot = () => { const m = this.deck.pptx.masters()[0]; return m ? this.deck.pptx.doc(m).documentElement : null; };
    impl("master.preserve", () => { const m = this.deck.pptx.masters()[0]; if (!m) return; const on = masterRoot()?.getAttribute("preserve") === "1"; this.deck.change("Beibehalten", [m], () => { const r = this.deck.pptx.doc(m).documentElement; if (on) r.removeAttribute("preserve"); else r.setAttribute("preserve", "1"); }); this.updateChrome(); }, () => this.masterMode, () => masterRoot()?.getAttribute("preserve") === "1");
    impl("arrange.align", () => {}, shapes);
    impl("arrange.rotate", () => {}, shapes);
    for (const [id, k] of [["arrange.alignLeft", "l"], ["arrange.alignCenter", "c"], ["arrange.alignRight", "r"], ["arrange.alignTop", "t"], ["arrange.alignMiddle", "m"], ["arrange.alignBottom", "b"]] as const) impl(id, () => this.alignShapes(k), shapes);
    impl("arrange.distH", () => this.alignShapes("dh"), () => this.editor.sel.length > 2 || (this.editor.sel.length > 0 && this.alignToSlide));
    impl("arrange.distV", () => this.alignShapes("dv"), () => this.editor.sel.length > 2 || (this.editor.sel.length > 0 && this.alignToSlide));
    impl("arrange.toSlide", () => { this.alignToSlide = true; }, undefined, () => this.alignToSlide || this.editor.sel.length === 1);
    impl("arrange.toSelected", () => { this.alignToSlide = false; }, () => this.editor.sel.length > 1, () => !this.alignToSlide && this.editor.sel.length > 1);
    impl("arrange.rotRight", () => this.rotateFlip("r"), shapes);
    impl("arrange.rotLeft", () => this.rotateFlip("l"), shapes);
    impl("arrange.flipH", () => this.rotateFlip("h"), shapes);
    impl("arrange.flipV", () => this.rotateFlip("v"), shapes);

    impl("insert.textBox", () => this.startTool({ kind: "textbox" }));
    impl("insert.textBoxH", () => this.startTool({ kind: "textbox" }));
    impl("insert.textBoxV", () => this.startTool({ kind: "textbox", vertical: true }));
    impl("insert.pictures", () => {});
    impl("insert.pictureFile", () => this.insertPicture());
    impl("insert.table", () => {});
    impl("insert.tableDialog", () => this.tableDialog());
    impl("insert.headerFooter", () => this.headerFooter());
    impl("insert.slideNumber", () => this.headerFooter());
    impl("insert.dateTime", () => this.headerFooter());
    impl("design.themes", () => {});
    impl("design.slideSize", () => {});
    impl("design.slide16x9", () => this.setSlideSize(12192000, 6858000), undefined, () => this.deck.pptx.slideSize.cx === 12192000);
    impl("design.slide4x3", () => this.setSlideSize(9144000, 6858000), undefined, () => this.deck.pptx.slideSize.cx === 9144000);

    impl("show.fromStart", () => this.startShow(0));
    impl("show.fromCurrent", () => this.startShow(this.active.current));
    impl("view.normal", () => { if (this.masterMode) this.setMasterView(false); this.setView("normal"); }, undefined, () => this.view === "normal" && !this.masterMode);
    impl("view.slideMaster", () => this.setMasterView(true), undefined, () => this.masterMode && this.masterKind === "slides");
    impl("view.notesMaster", () => this.setMasterView(true, "notes"), undefined, () => this.masterMode && this.masterKind === "notes");
    impl("view.handoutMaster", () => this.setMasterView(true, "handout"), undefined, () => this.masterMode && this.masterKind === "handout");
    for (const t of ["hdr", "sldImg", "ftr", "dt", "body", "sldNum"]) impl(`masterPh.${t}`, () => this.toggleMasterPh(t), () => this.masterMode && this.masterKind !== "slides", () => this.masterMode && this.masterPhOn(t));
    for (const id of ["notesMaster.orientation", "handoutMaster.orientation", "handoutMaster.perPage"]) impl(id, () => {}, () => this.masterMode);
    impl("master.formatBackground", () => this.bgPane.open(), () => this.masterMode);
    impl("master.hideBg", () => {
      const p = this.masterList()[this.masterIndex];
      if (!p) return;
      const on = this.deck.pptx.doc(p).documentElement.getAttribute("showMasterSp") === "0";
      this.deck.change("Hintergrundgrafiken ausblenden", [p], () => { const r = this.deck.pptx.doc(p).documentElement; if (on) r.removeAttribute("showMasterSp"); else r.setAttribute("showMasterSp", "0"); });
      this.editor.render(); this.panel.render();
    }, () => this.masterMode && this.masterKind === "slides" && this.masterIndex > 0, () => this.masterMode && this.deck.pptx.has(this.masterList()[this.masterIndex] ?? "") && this.deck.pptx.doc(this.masterList()[this.masterIndex]).documentElement.getAttribute("showMasterSp") === "0");
    impl("master.close", () => this.setMasterView(false));
    impl("master.rename", () => this.renameLayout(), () => this.masterMode && this.masterIndex > 0);
    impl("master.insertLayout", () => this.insertLayout(), () => this.masterMode);
    impl("master.delete", () => this.deleteLayout(), () => this.masterMode && this.masterIndex > 0 && !this.layoutInUse());
    impl("view.sorter", () => this.setView("sorter"), undefined, () => this.view === "sorter");
    impl("view.outline", () => this.setView("outline"), undefined, () => this.view === "outline");
    impl("view.notesPage", () => this.setView("notesPage"), undefined, () => this.view === "notesPage");
    impl("view.reading", () => this.setView("reading"), undefined, () => false);
    impl("view.notes", () => { this.showNotes = !this.showNotes; this.syncNotes(); this.editor.fit && this.editor.fitZoom(); }, undefined, () => this.showNotes);
    impl("view.fit", () => this.fitWindow());
    impl("view.zoom", () => this.zoomDialog());
    impl("view.gridlines", () => { ed().showGrid = !ed().showGrid; ed().render(); }, undefined, () => ed().showGrid);
    impl("view.guides", () => { ed().showGuides = !ed().showGuides; ed().render(); }, undefined, () => ed().showGuides);
    impl("view.gridDialog", () => this.gridDialog());
    impl("shape.editText", () => { const v = ed().selectedViews[0]; if (v) ed().enterText(v.id); }, shapes);
    impl("shape.sizePosition", () => this.formatPane.open(), shapes);
    impl("draw.format", () => this.formatPane.open(), shapes);
    impl("pic.format", () => this.formatPane.open(), shapes);
    const dlgHost = () => ({ runStyle: () => this.runStyle(), paraStyle: () => this.paraStyle(), fonts: () => this.fontList, formatRuns: (l: string, f: (r: Element) => void) => this.formatRuns(l, f), formatParas: (l: string, f: (p: Element) => void) => this.formatParas(l, f) });
    const finder = new FindDialog(this);
    impl("edit.find", () => finder.open(false));
    impl("design.slideCustom", () => this.slideSizeDialog());
    implementContext(this);
    impl("insert.symbol", () => import("./symbols").then((m) => m.symbolDialog(this)), () => this.editor.sel.length > 0 || !!this.editor.text);
    impl("insert.equation", () => insertEquation(this), () => !!this.slidePath());
    implementEquationCommands(this);
    impl("insert.symbols", () => {}, () => !!this.slidePath());
    impl("insert.video", () => import("./media").then((m) => m.insertMedia(this, "video")), () => !!this.slidePath());
    impl("insert.audio", () => import("./media").then((m) => m.insertMedia(this, "audio")), () => !!this.slidePath());
    impl("insert.media", () => {}, () => !!this.slidePath());
    const tool = () => this.editor.tool;
    const ink = (color: string, width: number, alpha?: number) => () => { this.editor.endText(); this.editor.tool = { kind: "ink", color, width, alpha }; this.editor.render(); this.updateChrome(); };
    const isInk = (color: string) => () => tool()?.kind === "ink" && (tool() as { color: string }).color === color;
    impl("ink.select", () => { this.editor.tool = null; this.updateChrome(); }, undefined, () => !tool());
    impl("ink.eraser", () => { this.editor.tool = { kind: "eraser" }; this.updateChrome(); }, undefined, () => tool()?.kind === "eraser");
    impl("ink.add", () => {}, () => !!this.slidePath());
    // Lassoauswahl: a loop drawn round shapes selects them (their middle inside the loop).
    impl("ink.lasso", () => {
      this.editor.endText();
      this.editor.tool = { kind: "points", mode: "scribble", done: (pts) => {
        const inside = (x: number, y: number) => { let c = false; for (let i = 0, j = pts.length - 1; i < pts.length; j = i++) { const a = pts[i], b = pts[j]; if ((a.y > y) !== (b.y > y) && x < ((b.x - a.x) * (y - a.y)) / (b.y - a.y) + a.x) c = !c; } return c; };
        const p = this.slidePath();
        const ids = p ? this.deck.view(p).shapes.filter((s) => !s.hidden && inside(s.x + s.w / 2, s.y + s.h / 2)).map((s) => s.id) : [];
        this.editor.select(ids);
        this.updateChrome();
      } };
      this.editor.render();
      this.updateChrome();
    }, () => !!this.slidePath(), () => tool()?.kind === "points");
    impl("ink.penBlack", ink("000000", 1), undefined, isInk("000000"));
    impl("ink.penRed", ink("FF0000", 10), undefined, isInk("FF0000"));
    impl("ink.pencil", ink("808080", 1.4), undefined, isInk("808080"));
    impl("ink.highlighter", ink("FFFF00", 17, 0.5), undefined, isInk("FFFF00"));
    void import("./transitions").then((m) => { this.transMod = m; });
    void import("./animations").then((m) => { this.animMod = m; });
    impl("trans.options", () => {}, () => (this.transMod?.transitionOptions(this).length ?? 0) > 0);
    impl("anim.options", () => {}, () => (this.animMod?.animationOptions(this).length ?? 0) > 0);
    impl("insert.wordArt", () => {}, () => !!this.slidePath());
    impl("slide.formatBackground", () => this.bgPane.open(), () => !!this.slidePath());
    impl("design.formatBackground", () => this.bgPane.open(), () => !!this.slidePath());
    impl("design.background", () => {}, () => !!this.slidePath());
    const cellAt = () => this.editor.text?.cell;
    impl("table.deleteRows", () => this.tableDelete("row"), () => !!cellAt());
    impl("table.deleteCols", () => this.tableDelete("col"), () => !!cellAt());
    impl("table.delete", () => { const t = this.editor.text; if (t) { this.editor.endText(); this.editor.select([t.id]); this.editor.deleteSelected(); } }, () => !!cellAt());
    impl("text.endEdit", () => { this.editor.endText(); this.editor.render(); this.updateChrome(); }, () => !!this.editor.text);
    impl("table.selectAll", () => { if (this.editor.text?.cell) { this.editor.selectAllCells(); return; } const v = this.editor.selectedViews.find((x) => x.frame === "table"); if (v) this.editor.select([v.id]); }, () => this.editor.selectedViews.some((x) => x.frame === "table"));
    const newComment = () => { this.comments.open(); setTimeout(() => (document.querySelector(".comment-input") as HTMLTextAreaElement | null)?.focus(), 50); };
    impl("review.newComment", newComment, () => !!this.slidePath());
    impl("insert.comment", newComment, () => !!this.slidePath());
    impl("review.showComments", () => this.comments.toggle(), undefined, () => this.comments.pane.isOpen);
    impl("view.comments", () => this.comments.toggle(), undefined, () => this.comments.pane.isOpen);
    impl("review.deleteComment", () => { const p = this.slidePath(); const n = p ? this.comments.list(p).length : 0; if (n) this.comments.remove(n - 1); }, () => !!this.slidePath() && this.comments.list(this.slidePath()!).length > 0);
    const jump = (d: 1 | -1) => { const n = this.deck.slides.length; for (let k = 1; k <= n; k++) { const i = (this.active.current + d * k + n * 2) % n; if (this.comments.list(this.deck.slides[i].path).length) { this.setCurrent(i); this.comments.open(); return; } } };
    impl("review.nextComment", () => jump(1));
    impl("review.prevComment", () => jump(-1));
    impl("show.setup", () => import("./showsetup").then((m) => m.showSetupDialog(this.deck)));
    // Einrichten's checkboxes, kept in presProps.xml as PowerPoint keeps them.
    const flag = (f: "narration" | "timings" | "media", id: string) => impl(id, () => { setShowFlag(this.deck, f, !showFlag(this.deck, f)); this.updateChrome(); }, () => !!this.slidePath(), () => showFlag(this.deck, f));
    flag("timings", "show.useTimings"); flag("narration", "show.playNarrations"); flag("media", "show.mediaControls");
    impl("show.custom", () => {}, () => !!this.slidePath());
    impl("show.monitor", () => {}, () => !!this.slidePath());
    impl("anim.pane", () => this.animPane.toggle(), undefined, () => this.animPane.isOpen);
    // Ansicht > Farbe/Graustufe: how slides are shown while editing (the file is unchanged, as in PowerPoint).
    const colorMode = (m: "color" | "gray" | "bw") => () => { document.body.classList.remove("view-gray", "view-bw"); if (m !== "color") document.body.classList.add(`view-${m}`); this.updateChrome(); };
    impl("view.color", colorMode("color"), undefined, () => !document.body.classList.contains("view-gray") && !document.body.classList.contains("view-bw"));
    impl("view.grayscale", colorMode("gray"), undefined, () => document.body.classList.contains("view-gray"));
    impl("view.bw", colorMode("bw"), undefined, () => document.body.classList.contains("view-bw"));
    impl("review.language", () => {}, () => !!this.slidePath());
    impl("review.thesaurus", () => import("./thesaurus").then((m) => m.thesaurusDialog(this)), () => !!this.slidePath());
    impl("review.setLanguage", () => import("./review").then((m) => m.languageDialog(this)), () => !!this.slidePath());
    impl("review.accessibility", () => import("./review").then((m) => m.accessibilityCheck(this)), () => this.deck.slides.length > 0);
    impl("anim.add", () => {}, () => this.editor.sel.length > 0);
    impl("anim.trigger", () => {}, () => this.selectedHasAnimation() || (!!this.slidePath() && this.editor.sel.length > 0 && this.animMod?.triggerOf(this.deck.pptx.doc(this.slidePath()!).documentElement, this.editor.sel[0]) != null));
    impl("anim.painter", () => { this.animBrush = this.animBrush === null ? this.editor.sel[0] ?? null : null; document.body.classList.toggle("painting", this.animBrush !== null); this.updateChrome(); },
      () => this.selectedHasAnimation() || this.animBrush !== null, () => this.animBrush !== null);
    impl("anim.gallery", () => {}, () => this.editor.sel.length > 0);
    impl("trans.gallery", () => {}, () => !!this.slidePath());
    const slide = () => !!this.slidePath();
    impl("trans.sound", () => {}, slide);
    impl("trans.duration", () => {}, slide);
    impl("trans.afterTime", () => {}, slide);
    impl("trans.applyAll", () => this.transMod?.applyToAll(this), slide);
    impl("trans.onClick", () => this.transMod?.setAdvanceOnClick(this, !this.transMod.advanceOnClick(this)), slide, () => this.transMod?.advanceOnClick(this) ?? true);
    impl("trans.after", () => this.transMod?.setAdvanceAfterOn(this, !this.transMod.advanceAfter(this)), slide, () => this.transMod?.advanceAfter(this) ?? false);
    impl("anim.duration", () => {}, () => this.selectedHasAnimation());
    impl("anim.delay", () => {}, () => this.selectedHasAnimation());
    impl("anim.start", () => {}, () => this.selectedHasAnimation());
    impl("anim.earlier", () => this.moveAnimation(-1), () => this.selectedHasAnimation());
    impl("anim.later", () => this.moveAnimation(1), () => this.selectedHasAnimation());
    impl("anim.preview", () => this.startShow(this.active.current), () => !!this.slidePath());
    impl("trans.preview", () => this.startShow(this.active.current), () => !!this.slidePath());
    impl("pic.crop", () => this.editor.toggleCrop(), () => this.editor.selectedViews[0]?.kind === "pic", () => this.editor.crop !== null);
    impl("pic.reset", () => this.resetPicture(), () => this.editor.selectedViews.some((v) => v.kind === "pic"));
    impl("pic.change", () => this.changePicture(), () => this.editor.selectedViews.length === 1 && this.editor.selectedViews[0].kind === "pic");
    impl("insert.link", () => import("./links").then((m) => m.insertLink(this)), () => this.editor.sel.length > 0);
    impl("shape.link", () => import("./links").then((m) => m.insertLink(this)), () => this.editor.sel.length > 0);
    impl("review.spelling", () => import("./spelling").then((m) => m.checkSpelling(this)));
    impl("edit.replace", () => finder.open(true));
    impl("font.dialog", () => import("./dialogs").then((m) => m.fontDialog(dlgHost())), text);
    impl("para.dialog", () => import("./dialogs").then((m) => m.paragraphDialog(dlgHost())), text);
    impl("shape.height", () => {}, shapes);
    impl("shape.width", () => {}, shapes);
    const tableEl = () => { const v = this.editor.selectedViews.find((x) => x.frame === "table"); return v ? this.editor.shapeEl(v.id) : undefined; };
    const tblFlag = (name: string) => () => tableEl()?.getElementsByTagNameNS(PptxDom.NS_A, "tblPr")[0]?.getAttribute(name) === "1";
    for (const [id, attr] of [["table.firstRow", "firstRow"], ["table.lastRow", "lastRow"], ["table.bandRow", "bandRow"], ["table.firstCol", "firstCol"], ["table.lastCol", "lastCol"], ["table.bandCol", "bandCol"]] as const) {
      impl(id, () => {
        const on = !tblFlag(attr)();
        this.editor.edit("Tabellenformatoptionen", () => { const pr = tableEl()?.getElementsByTagNameNS(PptxDom.NS_A, "tblPr")[0]; if (pr) { if (on) pr.setAttribute(attr, "1"); else pr.removeAttribute(attr); } });
      }, () => !!tableEl(), tblFlag(attr));
    }
    for (const [id, where] of [["table.insertAbove", "above"], ["table.insertBelow", "below"], ["table.insertLeft", "left"], ["table.insertRight", "right"]] as const) {
      impl(id, () => this.tableInsert(where), () => !!tableEl());
    }
  }

  private layoutInUse(): boolean {
    const lay = this.masterList()[this.masterIndex];
    return this.deck.slides.some((s) => this.deck.pptx.layoutOf(s.path) === lay);
  }

  private async renameLayout(): Promise<void> {
    const lay = this.masterList()[this.masterIndex];
    const cs = kid(this.deck.pptx.doc(lay).documentElement, "cSld");
    const name = await this.askText("Layout umbenennen", "Layoutname:", cs?.getAttribute("name") ?? "");
    if (name === null || !cs) return;
    this.deck.change("Layout umbenennen", [lay], () => cs.setAttribute("name", name));
  }

  /** Layout einfügen: a new layout (a copy of Nur Titel) after the current one, "Benutzerdefiniertes Layout". */
  private insertLayout(): void {
    const pptx = this.deck.pptx;
    const master = pptx.masters()[0];
    const path = pptx.freeName("ppt/slideLayouts/slideLayout#.xml");
    const list0 = this.masterList();
    const src = list0.find((p) => pptx.doc(p).documentElement.getAttribute("type") === "titleOnly") ?? list0[1];
    this.deck.change("Layout einfügen", [path, master], () => {
      const doc = PptxDom.parseXml(PptxDom.serialize(pptx.doc(src)));
      doc.documentElement.setAttribute("type", "cust");
      doc.documentElement.setAttribute("preserve", "1");
      kid(doc.documentElement, "cSld")?.setAttribute("name", "Benutzerdefiniertes Layout");
      pptx.put(path, doc, "application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml");
      pptx.link(path, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster", master);
      const rid = pptx.link(master, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout", path);
      const list = at(pptx.doc(master).documentElement, "sldLayoutIdLst")!;
      const ids = kids(list, "sldLayoutId").map((x) => num(x, "id") ?? 0);
      const el = make(pptx.doc(master), "p:sldLayoutId", { id: Math.max(...ids, 2147483648) + 1 });
      el.setAttributeNS(PptxDom.NS_R, "r:id", rid);
      const after = kids(list, "sldLayoutId")[this.masterIndex - 1];
      list.insertBefore(el, after ? after.nextSibling : null);
    });
    this.setCurrent(this.masterIndex + 1);
  }

  private deleteLayout(): void {
    const pptx = this.deck.pptx;
    const master = pptx.masters()[0];
    const lay = this.masterList()[this.masterIndex];
    this.deck.change("Layout löschen", [master], () => {
      const rels = pptx.rels(master);
      const r = rels.list.find((x) => pptx.target(master, x.id) === lay);
      if (!r) return;
      for (const x of kids(at(pptx.doc(master).documentElement, "sldLayoutIdLst"), "sldLayoutId")) if (x.getAttributeNS(PptxDom.NS_R, "id") === r.id) x.remove();
      rels.list.splice(rels.list.indexOf(r), 1);
      rels.changed = true;
    });
    this.setCurrent(this.masterIndex - 1);
  }

  private isHidden(): boolean {
    const p = this.slidePath();
    return !!p && this.deck.pptx.doc(p).documentElement.getAttribute("show") === "0";
  }

  private toggleHidden(): void {
    const sel = [...this.panel.selected];
    const hide = !this.isHidden();
    const paths = sel.map((i) => this.deck.slides[i].path);
    this.deck.change("Folie ausblenden", paths, () => {
      for (const p of paths) { const r = this.deck.pptx.doc(p).documentElement; if (hide) r.setAttribute("show", "0"); else r.removeAttribute("show"); }
    });
  }

  private async addSection(): Promise<void> {
    const name = await this.askText("Abschnitt umbenennen", "Abschnittsname:", "Unbenannter Abschnitt");
    if (name === null) return;
    const slides = this.deck.slides;
    const list = this.deck.sections();
    const start = Math.min(...this.panel.selected);
    const id = `{${crypto.randomUUID().toUpperCase()}}`;
    // Slides before the first section get a "Standardabschnitt" (PowerPoint).
    const out: { name: string; id: string; slides: number[] }[] = [];
    if (!list.length) {
      if (start > 0) out.push({ name: "Standardabschnitt", id: `{${crypto.randomUUID().toUpperCase()}}`, slides: slides.slice(0, start).map((s) => s.id) });
      out.push({ name, id, slides: slides.slice(start).map((s) => s.id) });
    } else {
      const startId = slides[start]?.id;
      for (const s of list) {
        const k = s.slides.indexOf(startId);
        if (k < 0) { out.push(s); continue; }
        if (k > 0) out.push({ ...s, slides: s.slides.slice(0, k) });
        out.push({ name, id, slides: s.slides.slice(k) });
        if (k === 0) out.push({ ...s, slides: [] });
      }
    }
    this.deck.setSections(out.filter((s, i) => s.slides.length || i === out.length - 1 || s.id === id), "Abschnitt hinzufügen");
  }

  private sectionAtCurrent(): number {
    const id = this.deck.slides[this.active.current]?.id;
    return this.deck.sections().findIndex((s) => s.slides.includes(id));
  }

  private async renameSection(): Promise<void> {
    const list = this.deck.sections();
    const i = Math.max(0, this.sectionAtCurrent());
    const name = await this.askText("Abschnitt umbenennen", "Abschnittsname:", list[i].name);
    if (name === null) return;
    list[i].name = name;
    this.deck.setSections(list, "Abschnitt umbenennen");
  }

  private removeSection(all: boolean): void {
    const list = this.deck.sections();
    if (all) { this.deck.setSections([], "Alle Abschnitte entfernen"); return; }
    const i = Math.max(0, this.sectionAtCurrent());
    // Its slides join the section before (or the next one).
    if (i > 0) list[i - 1].slides.push(...list[i].slides); else if (list[1]) list[1].slides.unshift(...list[0].slides);
    list.splice(i, 1);
    this.deck.setSections(list, "Abschnitt entfernen");
  }

  private async askText(title: string, label: string, value: string): Promise<string | null> {
    const input = h("input", { type: "text", value, style: "width:260px" }) as HTMLInputElement;
    setTimeout(() => input.select(), 0);
    const r = await showDialog({ title, body: h("div", { class: "form" }, h("label", null, label), input), buttons: [{ label: "Umbenennen", value: true, primary: true }, { label: "Abbrechen", value: false }], cancel: false });
    return r ? input.value : null;
  }

  /** Benutzerdefinierte Foliengröße: PowerPoint's sizes for screens and paper, or one of your own (cm). */
  private async slideSizeDialog(): Promise<void> {
    const sizes: [string, number, number][] = [["Bildschirmpräsentation (4:3)", 25.4, 19.05], ["Bildschirmpräsentation (16:9)", 25.4, 14.288], ["Bildschirmpräsentation (16:10)", 25.4, 15.875],
      ["Breitbild", 33.867, 19.05], ["A4-Papier (210 x 297 mm)", 27.517, 19.05], ["Benutzerdefiniert", 0, 0]];
    const cur = this.deck.pptx.slideSize;
    const cm = (emu: number) => (emu / 360000).toFixed(2).replace(".", ",");
    const w = h("input", { type: "text", value: `${cm(cur.cx)} cm`, style: "width:90px" }) as HTMLInputElement;
    const hh = h("input", { type: "text", value: `${cm(cur.cy)} cm`, style: "width:90px" }) as HTMLInputElement;
    const sel = h("select", { onchange: () => { const s2 = sizes[sel.selectedIndex]; if (s2[1]) { w.value = `${s2[1].toFixed(2).replace(".", ",")} cm`; hh.value = `${s2[2].toFixed(2).replace(".", ",")} cm`; } } }, ...sizes.map(([l]) => h("option", null, l))) as HTMLSelectElement;
    sel.selectedIndex = cur.cx === 12192000 ? 3 : cur.cx === 9144000 && cur.cy === 6858000 ? 0 : 5;
    const ok = await showDialog({ title: "Foliengröße", body: h("div", { class: "form" }, h("label", null, "Foliengröße für:"), sel, h("label", null, "Breite: ", w), h("label", null, "Höhe: ", hh)), buttons: [{ label: "OK", value: true, primary: true }, { label: "Abbrechen", value: false }], cancel: false });
    if (!ok) return;
    const n = (v: string) => Math.round(Number(v.replace(",", ".").replace(/[^\d.]/g, "")) * 360000);
    if (n(w.value) > 0 && n(hh.value) > 0) this.setSlideSize(n(w.value), n(hh.value));
  }

  private setSlideSize(cx: number, cy: number): void {
    const pres = this.deck.pptx.presPath;
    this.deck.change("Foliengröße", [pres], () => {
      const s = at(this.deck.pptx.presentation, "sldSz");
      if (s) { s.setAttribute("cx", String(cx)); s.setAttribute("cy", String(cy)); if (cx === 9144000) s.setAttribute("type", "screen4x3"); else s.removeAttribute("type"); }
    });
    this.editor.fitZoom();
  }

  private duplicateShapes(): void {
    const ed = this.editor;
    const views = ed.selectedViews;
    const ids: number[] = [];
    ed.edit("Duplizieren", () => {
      const tree = S.treeOf(this.deck.pptx.doc(this.slidePath()!));
      let next = S.maxId(tree);
      for (const v of views) {
        const el = ed.shapeEl(v.id);
        if (!el) continue;
        const copy = el.cloneNode(true) as Element;
        S.renumber(copy, () => ++next);
        // PowerPoint's Strg+D offset: 12 pt right and down (behaviour catalogue).
        S.setXfrm(copy, { x: v.x + 152400, y: v.y + 152400, w: v.w, h: v.h });
        tree.append(copy);
        ids.push(S.shapeId(copy));
      }
    });
    ed.select(ids);
  }

  private selectAll(): void {
    const ed = this.editor;
    if (ed.text) return;
    const p = this.slidePath();
    if (!p) return;
    ed.select(this.deck.view(p).shapes.filter((s) => !s.hidden).map((s) => s.id));
  }

  private selectionPane(): void {
    const p = this.slidePath();
    if (!p) return;
    const list = h("div", { class: "sel-pane" });
    const views = [...this.deck.view(p).shapes].reverse();
    for (const v of views) list.append(h("div", { class: `sel-pane-row${this.editor.sel.includes(v.id) ? " on" : ""}`, onclick: () => this.editor.select([v.id]) }, h("span", { class: "sel-pane-name" }, v.name)));
    void showDialog({ title: "Auswahl", body: list, buttons: [{ label: "Schließen", value: 0, primary: true }], cancel: 0 });
  }

  /** Datei > Optionen: the few settings RaabePresent keeps, grouped as PowerPoint's categories. */
  private async optionsDialog(): Promise<void> {
    const s = this.settings;
    const text = (v: string, w = 220) => h("input", { type: "text", value: v, style: `width:${w}px` }) as HTMLInputElement;
    const check = (label: string, on: boolean) => { const c = h("input", { type: "checkbox" }) as HTMLInputElement; c.checked = on; return { c, el: h("label", { class: "dlg-check" }, c, ` ${label}`) }; };
    const name = text(s.userName), ini = text(s.initials, 60);
    const touch = check("Touchmodus (größere Abstände im Menüband)", s.touchMode);
    const ac = check("Während der Eingabe korrigieren (AutoKorrektur)", s.autoCorrect);
    const minutes = h("input", { type: "number", min: "1", max: "120", value: String(Math.max(1, Math.round(s.autoRecoverSeconds / 60))), style: "width:60px" }) as HTMLInputElement;
    const ok = await showDialog({ title: "Optionen", body: h("div", { class: "form" },
      h("b", null, "Allgemein"), h("label", null, "Benutzername: ", name), h("label", null, "Initialen: ", ini), touch.el,
      h("b", null, "Dokumentprüfung"), ac.el,
      h("b", null, "Speichern"), h("label", null, "AutoWiederherstellen-Informationen speichern alle ", minutes, " Minuten")),
      buttons: [{ label: "OK", value: true, primary: true }, { label: "Abbrechen", value: false }], cancel: false });
    if (!ok) return;
    this.settings = { ...s, userName: name.value.trim(), initials: ini.value.trim(), touchMode: touch.c.checked, autoCorrect: ac.c.checked,
      autoRecoverSeconds: Math.max(1, Number(minutes.value) || 1) * 60 };
    saveSettings(this.settings);
    document.body.classList.toggle("touch", this.settings.touchMode);
    this.editor.autoCorrectOn = this.settings.autoCorrect;
  }

  private async zoomDialog(): Promise<void> {
    const v = h("input", { type: "number", value: String(this.editor.zoom), style: "width:70px" }) as HTMLInputElement;
    const r = await showDialog({ title: "Zoom", body: h("div", { class: "form" }, h("label", null, "Prozent: ", v, " %")), buttons: [{ label: "OK", value: true, primary: true }, { label: "Abbrechen", value: false }], cancel: false });
    if (r) this.setZoom(Number(v.value) || 100);
  }

  private async gridDialog(): Promise<void> {
    const ed = this.editor;
    const snap = h("input", { type: "checkbox" }) as HTMLInputElement; snap.checked = ed.snapToGrid;
    const grid = h("input", { type: "checkbox" }) as HTMLInputElement; grid.checked = ed.showGrid;
    const smart = h("input", { type: "checkbox" }) as HTMLInputElement; smart.checked = ed.smartGuides;
    const spacing = h("input", { type: "text", value: `${(ed.gridSpacing / 360000).toFixed(2).replace(".", ",")} cm`, style: "width:80px" }) as HTMLInputElement;
    const r = await showDialog({ title: "Raster und Führungslinien", body: h("div", { class: "form" },
      h("b", null, "Ausrichten"), h("label", null, snap, " Objekte am Raster ausrichten"),
      h("b", null, "Rastereinstellungen"), h("label", null, "Abstand: ", spacing), h("label", null, grid, " Raster auf dem Bildschirm anzeigen"),
      h("b", null, "Einstellungen für Führungslinien"), h("label", null, smart, " Intelligente Führungslinien anzeigen, wenn Formen ausgerichtet sind")),
      buttons: [{ label: "OK", value: true, primary: true }, { label: "Abbrechen", value: false }], cancel: false });
    if (!r) return;
    ed.snapToGrid = snap.checked; ed.showGrid = grid.checked; ed.smartGuides = smart.checked;
    const cm = Number(spacing.value.replace(",", ".").replace(/[^\d.]/g, ""));
    if (cm > 0) ed.gridSpacing = Math.round(cm * 360000);
    ed.render();
  }

  // ----- keys ----------------------------------------------------------------------------------------------

  /** Shortcuts that work everywhere in the window (also from the slide area's textarea). */
  shortcut(e: KeyboardEvent): boolean {
    return this.globalShortcut(e);
  }

  private globalShortcut(e: KeyboardEvent): boolean {
    const ctrl = e.ctrlKey && !e.altKey, shift = e.shiftKey, k = e.key;
    const run = (id: string) => { void this.commands.run(id); return true; };
    if (k === "F5" && !ctrl) return run(shift ? "show.fromCurrent" : "show.fromStart");
    if (e.altKey && shift && !e.ctrlKey && e.code === "Digit0") return run("insert.equation");
    if (k === "F12") return run("file.saveAs");
    if (k === "F7" && !shift) return run("review.spelling");
    if (k === "F7" && shift && !ctrl) return run("review.thesaurus");
    if (this.painter && k === "Escape") { this.painter = null; document.body.classList.remove("painting"); this.updateChrome(); return true; }
    if (!ctrl) return false;
    const key = k.length === 1 ? k.toLowerCase() : k;
    if (shift) {
      switch (key) {
        case "f": return run("font.bold");
        case "k": return run("font.italic");
        case "u": return run("font.underline");
        case "c": this.pickFormat(false); return true;
        case "v": this.applyPainterNow(); return true;
        // German PowerPoint (measured): Strg+Umschalt+G groups, Strg+Umschalt+H ungroups; Strg+G alone does nothing.
        case "g": return run("arrange.group");
        case "h": return run("arrange.ungroup");
        case ">": return run("font.grow");
        case "<": return run("font.shrink");
      }
      if (e.code === "Period" || k === ":") return run("font.grow");
      if (e.code === "Comma" || k === ";") return run("font.shrink");
      return false;
    }
    switch (key) {
      case "s": return run("file.save");
      case "o": return run("file.open");
      case "n": return run("file.new");
      case "w": return run("file.close");
      case "p": return run("file.print");
      case "z": return run("edit.undo");
      case "y": return run("edit.redo");
      case "c": return run("edit.copy");
      case "x": return run("edit.cut");
      case "v": return run("edit.paste");
      case "d": return run("edit.duplicate");
      case "m": return run("slide.new");
      case "g": return true;
      case "a": if (!this.editor.text) { this.selectAll(); return true; } return false;
      case "l": return run("para.left");
      case "e": return run("para.center");
      case "r": return run("para.right");
      case "j": return run("para.justify");
      case "b": return run("font.bold");
      case "i": return run("font.italic");
      case "u": return run("font.underline");
      case "<": return run("font.shrink");
      case ">": return run("font.grow");
      case " ": return run("font.clear");
      case "#": return run("font.superscript");
      case "-": return run("font.subscript");
      case "+": this.setZoom(this.editor.zoom + 10); return true;
      case "k": return run("insert.link");
      case "f": return run("edit.find");
      case "h": return run("edit.replace");
      case "t": return run("font.dialog");
    }
    if (k === "F1") { this.ribbon.toggleMinimized(); return true; }
    return false;
  }

  private applyPainterNow(): void {
    if (!this.painter) return;
    this.applyPainter();
  }

  private globalKey(e: KeyboardEvent): void {
    if (this.backstage.isOpen || this.show) return;
    const t = e.target as HTMLElement;
    if (t === this.editor.input || this.panel.el.contains(t) || this.sorter.el.contains(t)) return;
    if (t.closest?.(".dialog, .ribbon, .titlebar, .menu, input, textarea, select")) return;
    if (this.globalShortcut(e)) e.preventDefault();
  }

  private installSplitter(): void {
    const sp = this.main.querySelector(".splitter") as HTMLElement;
    sp.addEventListener("mousedown", (e) => {
      e.preventDefault();
      const start = e.clientX, w0 = this.panel.el.getBoundingClientRect().width;
      const move = (m: MouseEvent) => { this.panel.el.style.width = `${Math.max(90, Math.min(420, w0 + m.clientX - start))}px`; };
      const up = () => { window.removeEventListener("mousemove", move); window.removeEventListener("mouseup", up); this.panel.render(); if (this.editor.fit) this.editor.fitZoom(); };
      window.addEventListener("mousemove", move);
      window.addEventListener("mouseup", up);
    });
    // Format übertragen: the next click on a shape or text applies the picked formatting.
    this.editor.el.addEventListener("mouseup", () => { if (this.painter) setTimeout(() => this.applyPainter(), 0); });
    this.panel.el.addEventListener("contextmenu", () => {});
  }

  // ----- ribbon context --------------------------------------------------------------------------------------

  ribbonContext(): RibbonContext {
    const fonts = () => {
      const theme = this.slidePath() ? this.deck.view(this.slidePath()!).theme : null;
      return [{ heading: "Designschriftarten", items: theme ? [`${theme.major} (Überschriften)`, `${theme.minor} (Textkörper)`] : [] }, { heading: "Alle Schriftarten", items: this.fontList }];
    };
    const popupMenu = (entries: () => MenuEntry[]) => entries;
    void popupMenu;
    const noSpin = { value: () => "", commit: () => {}, step: () => {} };
    return {
      fontName: () => this.runStyle()?.font ?? "",
      fontSize: () => { const s = this.runStyle()?.size; return s === undefined ? "" : String(Math.round(s * 10) / 10).replace(".", ","); },
      fontOptions: fonts,
      setFont: (name) => {
        const theme = this.slidePath() ? this.deck.view(this.slidePath()!).theme : null;
        const face = name.endsWith(" (Überschriften)") ? "+mj-lt" : name.endsWith(" (Textkörper)") ? "+mn-lt" : name;
        void theme;
        this.formatRuns("Schriftart", (r) => T.setRunFont(r, face));
      },
      setSize: (v) => { const n = Number(v.replace(",", ".")); if (n > 0) this.setSizeAll(n); },
      paletteTheme: () => this.paletteTheme(), highlightNames: PP_HIGHLIGHT_NAMES,
      run: (cmd: string) => void this.commands.run(cmd),
      markup: () => this.showMarkup, toggleMarkup: () => { this.showMarkup = !this.showMarkup; this.editor.render(); this.updateChrome(); },
      fontColor: () => `#${this.colorHex(this.lastFontColor)}`,
      setFontColor: (hex) => { this.lastFontColor = hex.replace("#", "").toUpperCase(); this.formatRuns("Schriftfarbe", (r) => T.setRunColor(r, this.lastFontColor)); },
      highlight: () => `#${this.lastHighlight}`,
      setHighlight: (hex) => { const v = hex.replace("#", "").toUpperCase(); this.lastHighlight = v; this.formatRuns("Texthervorhebungsfarbe", (r) => T.setHighlight(r, v === "" ? null : v)); },
      fillColor: () => `#${this.colorHex(this.lastFill)}`,
      setFill: (hex) => this.setFill(hex.replace("#", "").toUpperCase()),
      lineColor: () => `#${this.colorHex(this.lastLine)}`,
      outlineExtra: () => this.outlineExtra(), fillExtra: () => this.fillExtra(),
      designTiles: () => this.designTiles(),
      fxTiles: (kind) => this.fxTiles(kind),
      setLine: (hex) => this.setLine(hex.replace("#", "").toUpperCase()),
      preview: { font: () => {}, size: () => {} },
      menus: {
        paste: [{ cmd: "edit.pasteDest" }, { cmd: "edit.pasteKeepSource" }, { cmd: "edit.pastePicture" }, { cmd: "edit.pasteText" }, "-", { cmd: "edit.pasteSpecial" }],
        layouts: () => this.deck.layouts().map((l) => ({ label: l.name, run: () => this.deck.setLayout([...this.panel.selected], l.path) })),
        section: [{ cmd: "slide.addSection" }, { cmd: "slide.renameSection" }, { cmd: "slide.removeSection" }, { cmd: "slide.removeAllSections" }, "-", { cmd: "slide.collapseAll" }, { cmd: "slide.expandAll" }],
        spacing: [["Sehr eng", -300], ["Eng", -150], ["Normal", 0], ["Weit", 300], ["Sehr weit", 600]].map(([label, v]) => ({ label: label as string, run: () => this.formatRuns("Zeichenabstand", (r) => { if (v) r.setAttribute("spc", String(v)); else r.removeAttribute("spc"); }) })),
        case: [{ label: "Ersten Buchstaben im Satz großschreiben.", run: () => this.changeCase("sentence") }, { label: "kleinbuchstaben", run: () => this.changeCase("lower") }, { label: "GROSSBUCHSTABEN", run: () => this.changeCase("upper") }, { label: "Ersten Buchstaben Im Wort Großschreiben", run: () => this.changeCase("title") }, { label: "gROSS-/kLEINSCHREIBUNG UMKEHREN", run: () => this.changeCase("toggle") }],
        bullets: this.bulletMenu(),
        numbering: this.numberMenu(),
        lineSpacing: [...[1, 1.5, 2, 2.5, 3].map((v) => ({ label: v.toFixed(1).replace(".", ","), run: () => this.lineSpacing(v) })), "-", { label: "Zeilenabstandsoptionen...", run: () => void this.commands.run("para.dialog") }],
        columns: [1, 2, 3].map((n) => ({ label: ["Eine Spalte", "Zwei Spalten", "Drei Spalten"][n - 1], run: () => this.setColumns(n) })),
        direction: [...[["Horizontal", "horz"], ["Text um 90° drehen", "vert"], ["Text um 270° drehen", "vert270"], ["Gestapelt", "wordArtVert"]].map(([label, v]) => ({ label, run: () => this.setBodyAttr("vert", v === "horz" ? null : v) })), { label: "Weitere Optionen...", run: () => void this.commands.run("draw.format") }],
        alignText: [["Oben", "t"], ["Mitte", "ctr"], ["Unten", "b"]].map(([label, v]) => ({ label, run: () => this.setBodyAttr("anchor", v) })),
        arrange: () => [{ heading: "Objekte sortieren" }, { cmd: "arrange.front" }, { cmd: "arrange.back" }, { cmd: "arrange.forward" }, { cmd: "arrange.backward" }, { heading: "Objekte gruppieren" }, { cmd: "arrange.group" }, { cmd: "arrange.ungroup" }, { cmd: "arrange.regroup" },
          { heading: "Objekte positionieren" }, { label: "Ausrichten", submenu: () => [{ cmd: "arrange.alignLeft" }, { cmd: "arrange.alignCenter" }, { cmd: "arrange.alignRight" }, "-", { cmd: "arrange.alignTop" }, { cmd: "arrange.alignMiddle" }, { cmd: "arrange.alignBottom" }, "-", { cmd: "arrange.distH" }, { cmd: "arrange.distV" }, "-", { cmd: "arrange.toSlide" }, { cmd: "arrange.toSelected" }] },
          { label: "Drehen", submenu: [{ cmd: "arrange.rotRight" }, { cmd: "arrange.rotLeft" }, { cmd: "arrange.flipV" }, { cmd: "arrange.flipH" }] }, "-", { cmd: "arrange.selectionPane" }],
        find: [{ cmd: "edit.find" }],
        select: [{ cmd: "edit.selectAll" }, { label: "Objekte markieren", run: () => this.selectAll() }, { cmd: "edit.selectionPane" }],
        table: [],
        pictures: [{ cmd: "insert.pictureFile" }],
        textBox: [{ cmd: "insert.textBoxH" }, { cmd: "insert.textBoxV" }],
        slideSize: [{ cmd: "design.slide4x3" }, { cmd: "design.slide16x9" }, "-", { cmd: "design.slideCustom" }],
        showCustom: [], record: [],
        transOptions: () => this.transMod?.transitionOptions(this) ?? [], animOptions: () => this.animMod?.animationOptions(this) ?? [], animStart: [{ label: "Beim Klicken", run: () => this.animPane.setStart("clickEffect") }, { label: "Mit Vorherigen", run: () => this.animPane.setStart("withEffect") }, { label: "Nach Vorherigen", run: () => this.animPane.setStart("afterEffect") }],
        colors: () => [...COLOR_SETS.map((c) => ({ label: c.name, run: () => this.applyColorSet(c.name, c.colors) })), "-", { label: "Farben anpassen...", run: () => void this.customColorsDialog() }],
        fonts: () => FONT_SETS.map((f) => ({ label: f.label, run: () => this.applyFontSet(f.name, f.major, f.minor) })),
        underline: [], outlineWeight: [], outlineDash: [], effects: this.effectsMenu(), quickStyles: () => [], zoomFit: [],
        equations: () => BUILT_IN.map((e) => ({ label: e.name, run: () => void insertEquation(this, e.seq) })),
        sounds: () => this.transMod?.soundMenu(this) ?? [],
        zooms: () => zoomMenu(this),
        bgStyles: () => bgStylesMenu(this),
        // Zeichnen > Hinzufügen: PowerPoint's pen kinds; ours take up the pen at once (no own toolbox to add it to).
        inkAdd: () => ([["Füllfederhalter", "1F3864", 1.4, 0], ["Textmarker", "FFFF00", 17, 0.5], ["Bleistift", "808080", 1.4, 0], ["Pinselstift", "00B050", 3, 0], ["Textstift", "000000", 1, 0], ["Stift", "000000", 1, 0]] as [string, string, number, number][])
          .map(([label, color, width, alpha]) => ({ label, run: () => { this.editor.endText(); this.editor.tool = { kind: "ink", color, width, alpha: alpha || undefined }; this.editor.render(); this.updateChrome(); } })),
        // Trigger > Bei Klick auf: the slide's other shapes by name.
        triggers: () => {
          const p = this.slidePath();
          if (!p) return [];
          const sel = this.editor.sel;
          const cur = sel.length ? this.animMod?.triggerOf(this.deck.pptx.doc(p).documentElement, sel[0]) ?? null : null;
          const shapes = this.deck.view(p).shapes.filter((s) => !sel.includes(s.id));
          return [{ label: "Bei Klick auf", submenu: () => [{ label: "Bei Mausklick (kein Trigger)", checked: cur === null, run: () => this.animMod?.setTrigger(this, null) },
            ...shapes.map((s) => ({ label: s.name, checked: cur === s.id, run: () => this.animMod?.setTrigger(this, s.id) }))] }];
        },
        // Fenster wechseln: the open presentations, numbered as PowerPoint lists its windows.
        windows: () => this.docs.map((d, i) => ({ label: `${i + 1} ${d.name}`, checked: d === this.active, run: () => this.activate(d) })),
      },
      panels: {
        newSlide: (close) => this.layoutGallery(close, (p) => this.newSlide(p), this.designName()),
        layout: (close) => this.layoutGallery(close, (p) => this.deck.setLayout([...this.panel.selected], p), this.designName()),
        shapes: (close) => this.shapeGallery(close),
        themes: (close) => this.themesGallery(close),
        transitions: (close) => this.transitions.gallery(close),
        animations: (close) => this.animations.gallery(close, false),
        addAnim: (close) => this.animations.gallery(close, true),
        table: (close) => this.tableGrid(close),
        chart: (close) => this.chartGallery(close),
        wordArt: (close) => this.wordArtGallery(close),
        quickStyles: (close) => quickStylesPanel(this, close),
      },
      spin: {
        transDuration: { value: () => this.transitions.durationText(), commit: (v) => this.transitions.setDuration(v), step: (d) => this.transitions.stepDuration(d) },
        transAfter: { value: () => this.transitions.afterText(), commit: (v) => this.transitions.setAfter(v), step: (d) => this.transitions.stepAfter(d) },
        animDuration: { value: () => this.animations.durationText(), commit: (v) => this.animations.setDuration(v), step: (d) => this.animations.stepDuration(d) },
        animDelay: { value: () => this.animations.delayText(), commit: (v) => this.animations.setDelay(v), step: (d) => this.animations.stepDelay(d) },
        shapeH: noSpin, shapeW: noSpin,
      },
    };
  }

  designName(): string {
    const p = this.slidePath();
    const name = p ? this.deck.view(p).theme.name : "Office";
    return name === "Office" ? "Office-Design" : name;
  }

  private changeCase(mode: "sentence" | "lower" | "upper" | "title" | "toggle"): void {
    const ed = this.editor;
    const conv = (s: string, first: boolean) => mode === "lower" ? s.toLowerCase() : mode === "upper" ? s.toUpperCase() : mode === "toggle" ? [...s].map((ch) => ch === ch.toUpperCase() ? ch.toLowerCase() : ch.toUpperCase()).join("")
      : mode === "title" ? s.toLowerCase().replace(/(^|[\s\-])(\p{L})/gu, (_m, a, b) => a + b.toUpperCase()) : (first ? s.charAt(0).toUpperCase() + s.slice(1).toLowerCase() : s.toLowerCase());
    const apply = (body: Element, from: T.Pos, to: T.Pos) => {
      T.paras(body).forEach((p, pi) => {
        let off = 0, first = true;
        for (const c of kids(p)) {
          if (c.localName !== "r") { off += c.localName === "br" ? 1 : 0; continue; }
          const t = kid(c, "t")!;
          const v = t.textContent ?? "";
          const s = pi === from.p ? from.o : pi > from.p ? 0 : Infinity, e = pi === to.p ? to.o : pi < to.p ? Infinity : -1;
          const a = Math.max(0, s - off), b = Math.min(v.length, e - off);
          if (a < b) t.textContent = v.slice(0, a) + conv(v.slice(a, b), first && a === 0) + v.slice(b);
          if (v) first = false;
          off += v.length;
        }
      });
    };
    if (ed.text) {
      const t = ed.text, body = ed.textBody(t.id);
      if (body) ed.edit("Groß-/Kleinschreibung", () => apply(body, T.minPos(t.caret, t.anchor), T.maxPos(t.caret, t.anchor)));
      return;
    }
    ed.edit("Groß-/Kleinschreibung", () => { for (const v of ed.selectedViews) { const b = ed.textBody(v.id); if (b) apply(b, { p: 0, o: 0 }, T.endPos(b)); } });
  }

  private setBodyAttr(name: string, v: string | null): void {
    const ed = this.editor;
    const ids = ed.text ? [ed.text.id] : ed.sel;
    ed.edit("Text", () => { for (const id of ids) { const el = ed.shapeEl(id); const bp = el && S.bodyPrOf(el); if (bp) T.setAttr(bp, name, v); } });
  }

  private setColumns(n: number): void {
    this.setBodyAttr("numCol", n > 1 ? String(n) : null);
  }

  /** Einfügen > WordArt: a text box "Hier steht Ihr Text" in 54 pt with one of the text styles (our selection of Office's). */
  private wordArtGallery(close: () => void): HTMLElement {
    return wordArtInsertPanel(this, close, (rpr) => this.insertWordArt(rpr));
  }

  private insertWordArt(rpr: string): void {
    const slide = this.slidePath();
    if (!slide) return;
    const size = this.deck.pptx.slideSize;
    let id = 0;
    this.editor.edit("WordArt einfügen", () => {
      const doc = this.deck.pptx.doc(slide);
      const tree = S.treeOf(doc);
      id = S.maxId(tree) + 1;
      const w = 6096000, hh = 923330;
      const el = fragment(doc, `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="Rechteck ${id - 1}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="${(size.cx - w) / 2}" y="${(size.cy - hh) / 2}"/><a:ext cx="${w}" cy="${hh}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr>` +
        `<p:txBody><a:bodyPr wrap="none" lIns="91440" tIns="45720" rIns="91440" bIns="45720"><a:spAutoFit/></a:bodyPr><a:lstStyle/><a:p><a:pPr algn="ctr"/><a:r><a:rPr lang="de-DE" sz="5400" b="1" cap="none" spc="0">${rpr}</a:rPr><a:t>Hier steht Ihr Text</a:t></a:r></a:p></p:txBody></p:sp>`)[0];
      tree.append(el);
    });
    this.editor.updateAutofit(id);
    this.editor.enterText(id, "all");
  }

  private chartGallery(close: () => void): HTMLElement {
    const kinds: [string, string][] = [["Gruppierte Säulen", "column"], ["Linie", "line"], ["Kreis", "pie"], ["Gruppierte Balken", "bar"], ["Fläche", "area"], ["Punkt (X Y)", "scatter"]];
    return h("div", { class: "menu" }, ...kinds.map(([label, k]) => h("button", { class: "menu-item", onclick: () => { close(); this.insertChart(k); } }, label)));
  }

  private insertChart(kind: string): void {
    // A chart with PowerPoint's sample data (Kategorie 1-4, Datenreihe 1-3), its values cached in the part.
    void import("./charts").then(({ insertChart }) => insertChart(this, kind));
  }

  // ----- transitions and animations ---------------------------------------------------------------------------

  readonly transitions = new (class {
    constructor(private app: PresentApp) {}
    gallery(close: () => void): HTMLElement { return this.app.transitionGallery(close); }
    durationText(): string { return this.app.transitionDuration(); }
    setDuration(v: string): void { this.app.setTransitionDuration(v); }
    stepDuration(d: 1 | -1): void { this.app.stepTransitionDuration(d); }
    afterText(): string { return this.app.transitionAfter(); }
    setAfter(v: string): void { this.app.setTransitionAfter(v); }
    stepAfter(d: 1 | -1): void { this.app.stepTransitionAfter(d); }
  })(this);

  readonly animations = new (class {
    constructor(private app: PresentApp) {}
    gallery(close: () => void, add: boolean): HTMLElement { return this.app.animationGallery(close, add); }
    durationText(): string { return this.app.animationDuration(); }
    setDuration(v: string): void { this.app.setAnimationDuration(v); }
    stepDuration(d: 1 | -1): void { this.app.stepAnimationDuration(d); }
    delayText(): string { return this.app.animationDelay(); }
    setDelay(v: string): void { this.app.setAnimationDelay(v); }
    stepDelay(d: 1 | -1): void { this.app.stepAnimationDelay(d); }
  })(this);

  transitionGallery(close: () => void): HTMLElement {
    const grid = h("div", { class: "fx-gallery" });
    for (const g of TRANSITION_GROUPS) {
      grid.append(h("div", { class: "gal-head" }, g.title));
      const row = h("div", { class: "fx-row" });
      for (const [label, kind] of g.items) row.append(h("button", { class: "fx-item", "aria-label": label, onclick: () => { close(); void import("./transitions").then((m) => m.setTransition(this, kind)); } }, icon(kind === "none" ? "close-x" : "transition"), h("span", null, label)));
      grid.append(row);
    }
    return grid;
  }

  transitionDuration(): string { return this.transitionInfo()?.duration ?? "01,00"; }
  transitionAfter(): string { return this.transitionInfo()?.after ?? "00:00,00"; }
  private transitionInfo(): { duration: string; after: string } | null {
    const p = this.slidePath();
    if (!p) return null;
    const tr = kid(this.deck.pptx.doc(p).documentElement, "transition") ?? this.deck.pptx.doc(p).documentElement.getElementsByTagNameNS(PptxDom.NS_P, "transition")[0];
    const dur = tr ? Number(tr.getAttributeNS(PptxDom.NS_P14, "dur") ?? tr.getAttribute("p14:dur") ?? (tr.getAttribute("spd") === "slow" ? 1000 : tr.getAttribute("spd") === "med" ? 750 : 500)) : 1000;
    const adv = tr ? Number(tr.getAttribute("advTm") ?? 0) : 0;
    const fmt = (ms: number) => (ms / 1000).toFixed(2).padStart(5, "0").replace(".", ",");
    return { duration: fmt(dur), after: `${String(Math.floor(adv / 60000)).padStart(2, "0")}:${((adv % 60000) / 1000).toFixed(2).padStart(5, "0").replace(".", ",")}` };
  }
  setTransitionDuration(v: string): void { void import("./transitions").then((m) => m.setTransitionDuration(this, parseSeconds(v))); }
  stepTransitionDuration(d: 1 | -1): void { const cur = parseSeconds(this.transitionDuration()); void import("./transitions").then((m) => m.setTransitionDuration(this, Math.max(0.01, cur + d * 0.25))); }
  setTransitionAfter(v: string): void { void import("./transitions").then((m) => m.setAdvanceAfter(this, parseSeconds(v))); }
  stepTransitionAfter(d: 1 | -1): void { const cur = parseSeconds(this.transitionAfter()); void import("./transitions").then((m) => m.setAdvanceAfter(this, Math.max(0, cur + d))); }

  animationGallery(close: () => void, add: boolean): HTMLElement {
    const grid = h("div", { class: "fx-gallery" });
    if (!add) grid.append(h("div", { class: "fx-row" }, h("button", { class: "fx-item", "aria-label": "Ohne", onclick: () => { close(); void import("./animations").then((m) => { const p = this.slidePath(); const root = p ? this.deck.pptx.doc(p).documentElement : null; const timing = root && kid(root, "timing"); if (timing) this.editor.edit("Animation entfernen", () => { for (const id of this.editor.sel) m.removeEffects(timing, id); m.tidyTiming(root!); }); }); } }, icon("close-x"), h("span", null, "Ohne"))));
    for (const g of ANIMATION_GROUPS) {
      grid.append(h("div", { class: "gal-head" }, g.title));
      const row = h("div", { class: "fx-row" });
      for (const it of g.items) row.append(h("button", { class: `fx-item fx-${g.cls}`, "aria-label": it.label, onclick: () => { close(); void import("./animations").then((m) => m.addAnimation(this, g.cls, it, add)); } }, icon("animation"), h("span", null, it.label)));
      grid.append(row);
    }
    // Weitere …effekte: a list of each kind to pick from (PowerPoint's dialogs, ours with the same effects).
    const more = (label: string, cls: string) => h("button", { class: "menu-item", onclick: () => { close(); void this.moreEffectsDialog(label.replace("...", ""), cls, add); } }, label);
    grid.append(h("div", { class: "menu-sep" }), more("Weitere Eingangseffekte...", "entr"), more("Weitere Hervorhebungseffekte...", "emph"), more("Weitere Ausgangseffekte...", "exit"),
      more("Weitere Animationspfade...", "path"), h("button", { class: "menu-item", disabled: true }, "OLE-Aktionsarten..."));
    return grid;
  }

  private async moreEffectsDialog(title: string, cls: string, add: boolean): Promise<void> {
    const g = ANIMATION_GROUPS.find((x) => x.cls === cls);
    if (!g) return;
    const list = h("select", { size: "12", style: "width:260px" }, ...g.items.map((it, i) => h("option", { value: String(i) }, it.label))) as HTMLSelectElement;
    list.selectedIndex = 0;
    const TITLES: Record<string, string> = { entr: "Eingangseffekt ändern", emph: "Hervorhebungseffekt ändern", exit: "Ausgangseffekt ändern", path: "Animationspfad ändern" };
    void title;
    const ok = await showDialog({ title: TITLES[cls] ?? "Effekt ändern", body: h("div", { class: "form" }, list), buttons: [{ label: "OK", value: true, primary: true }, { label: "Abbrechen", value: false }], cancel: false });
    if (!ok || list.selectedIndex < 0) return;
    const it = g.items[list.selectedIndex];
    void import("./animations").then((m) => m.addAnimation(this, g.cls, it, add));
  }

  /** The selected shape's first effect: its duration and its delay (as PowerPoint shows them). */
  private selectedEffectTiming(): { dur: number; delay: number } | null {
    const p = this.slidePath();
    if (!p) return null;
    const e = effects(this.deck.pptx.doc(p).documentElement).list.find((x) => this.editor.sel.includes(x.spid));
    if (!e) return null;
    const durs = Array.from(e.ctn.getElementsByTagNameNS(PptxDom.NS_P, "cTn")).map((x) => Number(x.getAttribute("dur") ?? 0)).filter((d) => Number.isFinite(d) && d > 1);
    const delay = Number(e.ctn.getElementsByTagNameNS(PptxDom.NS_P, "cond")[0]?.getAttribute("delay") ?? 0) || 0;
    return { dur: durs.length ? Math.max(...durs) : 500, delay };
  }
  animationDuration(): string { return ((this.selectedEffectTiming()?.dur ?? 500) / 1000).toFixed(2).padStart(5, "0").replace(".", ","); }
  setAnimationDuration(v: string): void { void import("./animations").then((m) => m.setDuration(this, parseSeconds(v))); }
  stepAnimationDuration(d: 1 | -1): void { void import("./animations").then((m) => m.setDuration(this, Math.max(0.01, parseSeconds(this.animationDuration()) + d * 0.25))); }
  animationDelay(): string { return ((this.selectedEffectTiming()?.delay ?? 0) / 1000).toFixed(2).padStart(5, "0").replace(".", ","); }
  setAnimationDelay(v: string): void { void import("./animations").then((m) => m.setDelay(this, parseSeconds(v))); }
  stepAnimationDelay(d: 1 | -1): void { void import("./animations").then((m) => m.setDelay(this, Math.max(0, parseSeconds(this.animationDelay()) + d * 0.25))); }

  // ----- files ------------------------------------------------------------------------------------------------

  async openDialog(): Promise<void> {
    const path = this.platform.openAnyDialog ? await this.platform.openAnyDialog("Öffnen", OPEN_FILTERS) : await this.platform.openFileDialog("pptx");
    if (path) await this.openPath(path);
  }

  async openPath(path: string): Promise<void> {
    const existing = this.docs.find((d) => d.path && samePath(d.path, path));
    if (existing) { this.activate(existing); this.backstage.close(); return; }
    let deck: Deck;
    try {
      deck = new Deck(Pptx.read(await this.platform.readFile(path)));
    } catch (e) {
      removeRecent(path);
      await messageBox("RaabePresent", `Die Datei „${baseName(path)}“ kann nicht geöffnet werden.\n\n${e instanceof Error ? e.message : String(e)}`);
      return;
    }
    // A template (.potx) opens as a new presentation based on it, as in PowerPoint.
    const template = /\.pot[xm]$/i.test(path);
    const doc = template ? this.makeDoc(deck, `Präsentation${++this.untitled}`, null) : this.makeDoc(deck, stripExt(baseName(path)), path);
    if (!template) doc.lastSaved = Date.now();
    // An untouched new presentation makes room for the opened file (PowerPoint).
    if (this.docs.length === 1 && this.docs[0].path === null && this.docs[0].deck.version === 0) this.docs = [];
    this.docs.push(doc);
    addRecent(path, baseName(path));
    this.backstage.close();
    this.activate(doc);
  }

  async save(doc = this.active, as = false): Promise<boolean> {
    this.editor.endText();
    let path = doc.path;
    if (as || !path) {
      // With macros (an opened .pptm) PowerPoint offers .pptm first.
      const ext0 = this.hasMacros(doc.deck) ? "pptm" : "pptx";
      const filters = ext0 === "pptm" ? [SAVE_FILTERS[1], SAVE_FILTERS[0], ...SAVE_FILTERS.slice(2)] : SAVE_FILTERS;
      path = this.platform.saveAnyDialog ? await this.platform.saveAnyDialog("Speichern unter", `${doc.name}.${ext0}`, filters) : await this.platform.saveFileDialog(`${doc.name}.pptx`, "pptx");
      if (!path) return false;
      if (!/\.(pptx|pptm|ppsx|potx)$/i.test(path)) path += `.${ext0}`;
    }
    const ext = path.slice(path.lastIndexOf(".") + 1).toLowerCase();
    // Macros cannot go into a file type without them: PowerPoint asks first, and so do we (they are kept, never run).
    if (this.hasMacros(doc.deck) && !ext.endsWith("m")) {
      const ok = await showDialog({ title: "RaabePresent", body: "Diese Präsentation enthält Makros (ein VBA-Projekt), die in diesem Dateityp nicht gespeichert werden können. Ohne Makros speichern?",
        buttons: [{ label: "Ja", value: true, primary: true }, { label: "Nein", value: false }], cancel: false });
      if (!ok) return false;
      this.dropMacros(doc.deck);
    }
    doc.deck.pptx.setType(doc.deck.pptx.presPath, PRES_TYPES[ext] ?? PRES_TYPES.pptx);
    try {
      await this.platform.writeFile(path, doc.deck.write());
    } catch (e) {
      await messageBox("RaabePresent", `Die Datei konnte nicht gespeichert werden.\n\n${e instanceof Error ? e.message : String(e)}`);
      return false;
    }
    doc.path = path;
    doc.name = stripExt(baseName(path));
    doc.savedVersion = doc.deck.version;
    doc.lastSaved = Date.now();
    // Other windows of the same presentation (Neues Fenster) are saved with it.
    for (const t of this.docs) if (t !== doc && t.deck === doc.deck) { t.path = path; t.savedVersion = doc.savedVersion; t.lastSaved = doc.lastSaved; }
    this.nameWindows(doc.deck, doc.name);
    addRecent(path, baseName(path));
    void this.platform.deleteRecovery(doc.id).catch(() => {});
    this.updateTabs();
    return true;
  }

  /**
   * Ansicht > Neues Fenster: another window onto the same presentation (its own slide and view), named
   * "Name:1", "Name:2" as in PowerPoint. RaabePresent shows it as another tab of the window; changes made in
   * one show in the other.
   */
  private newWindow(): void {
    const d = this.active;
    const deck = d.deck;
    const twin: Doc = { ...d, id: crypto.randomUUID() };
    deck.onChange = () => { if (this.active.deck === deck) this.afterChange(); else this.updateTabs(); };
    this.docs.splice(this.docs.indexOf(d) + 1, 0, twin);
    this.nameWindows(deck, d.name);
    this.activate(twin);
  }

  /** The windows of one presentation: "Name:1", "Name:2" … while there are several, else "Name". */
  private nameWindows(deck: Deck, name: string): void {
    const base = name.replace(/:\d+$/, "");
    const wins = this.docs.filter((x) => x.deck === deck);
    wins.forEach((w, i) => { w.name = wins.length > 1 ? `${base}:${i + 1}` : base; });
    this.updateTabs();
  }

  private async closeDoc(doc: Doc): Promise<void> {
    // Another window still shows this presentation: this one closes without asking.
    const twins = this.docs.filter((x) => x !== doc && x.deck === doc.deck);
    if (twins.length) {
      const i = this.docs.indexOf(doc);
      this.docs.splice(i, 1);
      this.nameWindows(doc.deck, doc.name);
      if (doc === this.active) this.activate(twins[0]);
      return;
    }
    if (doc.deck.version !== doc.savedVersion) {
      const r = await showDialog({ title: "RaabePresent", body: `Möchten Sie die Änderungen an „${doc.name}“ speichern?`, buttons: [{ label: "Speichern", value: "save", primary: true }, { label: "Nicht speichern", value: "discard" }, { label: "Abbrechen", value: "cancel" }], cancel: "cancel" });
      if (r === "cancel") return;
      if (r === "save" && !(await this.save(doc))) return;
    }
    void this.platform.deleteRecovery(doc.id).catch(() => {});
    const i = this.docs.indexOf(doc);
    this.docs.splice(i, 1);
    if (!this.docs.length) { await this.platform.window.close(); return; }
    if (doc === this.active) this.activate(this.docs[Math.max(0, i - 1)]);
    else this.updateTabs();
  }

  private async confirmQuit(): Promise<boolean> {
    for (const d of [...this.docs]) {
      if (d.deck.version === d.savedVersion) continue;
      this.activate(d);
      const r = await showDialog({ title: "RaabePresent", body: `Möchten Sie die Änderungen an „${d.name}“ speichern?`, buttons: [{ label: "Speichern", value: "save", primary: true }, { label: "Nicht speichern", value: "discard" }, { label: "Abbrechen", value: "cancel" }], cancel: "cancel" });
      if (r === "cancel") return false;
      if (r === "save" && !(await this.save(d))) return false;
    }
    for (const d of this.docs) void this.platform.deleteRecovery(d.id).catch(() => {});
    return true;
  }

  private async autoRecover(): Promise<void> {
    for (const d of this.docs) {
      if (d.deck.version === d.recoveryVersion || d.deck.version === d.savedVersion) continue;
      try {
        await this.platform.saveRecovery({ id: d.id, title: d.name, path: d.path, savedAt: Date.now() }, d.deck.write());
        d.recoveryVersion = d.deck.version;
      } catch { /* next time */ }
    }
  }

  private async offerRecovery(): Promise<void> {
    const list = await this.platform.listRecovery().catch(() => []);
    for (const e of list) {
      const r = await showDialog({ title: "Dokumentwiederherstellung", body: `RaabePresent wurde unerwartet beendet. „${e.title}“ (gesichert ${new Date(e.savedAt).toLocaleString("de-DE")}) wiederherstellen?`, buttons: [{ label: "Wiederherstellen", value: "yes", primary: true }, { label: "Verwerfen", value: "no" }], cancel: "no" });
      if (r === "yes") {
        try {
          const doc = this.makeDoc(new Deck(Pptx.read(await this.platform.readRecovery(e.id))), e.title, e.path);
          this.docs.push(doc);
          this.activate(doc);
        } catch { /* broken backup */ }
      }
      await this.platform.deleteRecovery(e.id).catch(() => {});
    }
  }

  private transMod: typeof import("./transitions") | null = null;
  private animMod: typeof import("./animations") | null = null;
  private printSettings: import("./print").PrintSettings = { layout: "full", which: "all", range: "", frame: false };
  private printMod: typeof import("./print") | null = null;
  private printPageEl(): HTMLElement {
    if (this.printMod) return this.printMod.printPage(this, this.printSettings);
    const box = h("div", null, h("h1", null, "Drucken"));
    void import("./print").then((m) => { this.printMod = m; box.replaceWith(m.printPage(this, this.printSettings)); });
    return box;
  }

  /** PDF: every slide as a page of the slide's size (Datei > Exportieren). */
  async exportPdf(): Promise<void> {
    const path = await this.platform.saveFileDialog(`${this.active.name}.pdf`, "pdf");
    if (!path) return;
    const { printSlides } = await import("./print");
    await printSlides(this, path);
  }

  // ----- backstage --------------------------------------------------------------------------------------

  private backstageEntries(): BackstageEntry[] {
    const greeting = () => { const hr = new Date().getHours(); return hr < 11 ? "Guten Morgen" : hr < 18 ? "Guten Tag" : "Guten Abend"; };
    const designs = () => h("div", { class: "templates" }, ...DESIGNS.map((d, i) => {
      const sample = new Deck(newPptx(d));
      const size = sample.pptx.slideSize;
      const tw = 176, s = tw / (size.cx / EMU_PX);
      const thumb = h("div", { class: "thumb present-thumb", style: `width:${tw}px;height:${(size.cy / EMU_PX) * s}px;position:relative;overflow:hidden` });
      const r = renderSlide(sample.view(sample.slides[0].path), size, { editing: false, image: () => undefined, pptx: sample.pptx });
      r.el.style.transform = `scale(${s})`; r.el.style.transformOrigin = "0 0";
      thumb.append(r.el);
      return h("button", { class: "template", onclick: () => { this.backstage.close(); this.newDoc(i === 0 ? undefined : d); } }, thumb, h("span", null, i === 0 ? "Leere Präsentation" : d.name));
    }), ...PRESENT_TEMPLATES.map((t) => {
      const sample = templateDeck(t);
      const size = sample.pptx.slideSize;
      const tw = 176, s = tw / (size.cx / EMU_PX);
      const thumb = h("div", { class: "thumb present-thumb", style: `width:${tw}px;height:${(size.cy / EMU_PX) * s}px;position:relative;overflow:hidden` });
      const r = renderSlide(sample.view(sample.slides[0].path), size, { editing: false, image: () => undefined, pptx: sample.pptx });
      r.el.style.transform = `scale(${s})`; r.el.style.transformOrigin = "0 0";
      thumb.append(r.el);
      return h("button", { class: "template", onclick: () => { this.backstage.close(); this.untitled++; const doc = this.makeDoc(templateDeck(t), `${t.name}${this.untitled}`, null); this.docs.push(doc); this.activate(doc); } }, thumb, h("span", null, t.name));
    }));
    return [
      { id: "home", label: "Startseite", icon: "home", page: () => h("div", null, h("h1", null, greeting()), h("h2", null, "Neu"), designs(), h("h2", null, "Zuletzt verwendet"), this.recentTable()) },
      { id: "new", label: "Neu", icon: "new", page: () => h("div", null, h("h1", null, "Neu"), designs()) },
      { id: "open", label: "Öffnen", icon: "open", page: () => h("div", null, h("h1", null, "Öffnen"), h("button", { class: "bigbtn", onclick: () => void this.openDialog() }, icon("open", "l"), "Durchsuchen"), h("h2", null, "Zuletzt verwendet"), this.recentTable()) },
      { id: "info", label: "Informationen", icon: "info", divider: true, page: () => this.infoPage() },
      { id: "save", label: "Speichern", icon: "save", cmd: "file.save" },
      { id: "saveAs", label: "Speichern unter", icon: "save-as", cmd: "file.saveAs" },
      { id: "print", label: "Drucken", icon: "print", page: () => this.printPageEl() },
      { id: "export", label: "Exportieren", icon: "export", page: () => h("div", null, h("h1", null, "Exportieren"), h("button", { class: "bigbtn", style: "width:220px", onclick: () => void this.exportPdf() }, icon("export", "l"), "PDF/XPS-Dokument erstellen"), h("p", { class: "muted" }, "Die Folien als PDF, so wie sie in der Bildschirmpräsentation aussehen.")) },
      { id: "close", label: "Schließen", icon: "close-doc", cmd: "file.close" },
      { id: "options", label: "Optionen", icon: "settings", cmd: "file.options", bottom: true },
    ];
  }

  private infoPage(): HTMLElement {
    const d = this.active;
    const props = h("table", { class: "props" },
      h("tr", null, h("td", null, "Folien"), h("td", null, String(d.deck.slides.length))),
      h("tr", null, h("td", null, "Speicherort"), h("td", null, d.path ?? "Noch nicht gespeichert")),
      h("tr", null, h("td", null, "Zuletzt gespeichert"), h("td", null, d.lastSaved ? new Date(d.lastSaved).toLocaleString("de-DE") : "–")));
    return h("div", null, h("h1", null, "Informationen"), h("h2", { style: "margin-top:0" }, d.name), props);
  }

  private recentTable(): HTMLElement {
    const list = loadRecent();
    if (!list.length) return h("p", { class: "muted" }, "Noch keine zuletzt verwendeten Präsentationen.");
    const sorted = [...list.filter((f) => f.pinned), ...list.filter((f) => !f.pinned)];
    return h("table", { class: "recent" }, ...sorted.map((f) => h("tr", { onclick: () => void this.openPath(f.path) },
      h("td", null, icon("document")), h("td", null, h("div", null, f.name), h("div", { class: "muted" }, f.path)),
      h("td", null, new Date(f.opened).toLocaleDateString("de-DE")),
      h("td", null, h("button", { class: "pinbtn", "aria-label": f.pinned ? "Lösen" : "Anheften", onclick: (e: MouseEvent) => { e.stopPropagation(); setPinned(f.path, !f.pinned); this.backstage.refreshPage(); } }, icon(f.pinned ? "unpin" : "pin"))))));
  }

  // ----- test hooks -------------------------------------------------------------------------------------------

  /** Test hooks for feel-lab (catalogue_present.py): state read back as PowerPoint reports it through COM. */
  testApi() {
    const app = this;
    return {
      app,
      state() {
        const p = app.slidePath();
        const v = p ? app.deck.view(p) : null;
        const ed = app.editor;
        return {
          slide: app.active.current + 1, slides: app.deck.slides.length,
          shapes: v ? v.shapes.map((s) => ({ id: s.id, name: s.name, x: Math.round(s.x), y: Math.round(s.y), w: Math.round(s.w), h: Math.round(s.h), rot: Math.round(s.rot * 100) / 100,
            text: s.text ? s.text.paras.map((q) => q.runs.map((r) => r.text).join("")).join("\r") : null, ph: s.ph?.type ?? null, kind: s.kind })) : [],
          selection: ed.text ? { type: "text", ids: [ed.text.id], start: ed.text.caret, anchor: ed.text.anchor } : ed.sel.length ? { type: "shapes", ids: ed.sel } : { type: "none", ids: [] as number[] },
          zoom: ed.zoom, layout: p ? app.deck.layoutName(p) : "",
          panel: [...app.panel.selected].sort((a, b) => a - b).map((i) => i + 1),
        };
      },
      open: async (path: string) => app.openPath(path),
      saveTo: async (path: string) => { await app.platform.writeFile(path, app.deck.write()); },
      loadBytes: (b64: string) => {
        const bin = atob(b64); const bytes = new Uint8Array(bin.length); for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
        const doc = app.makeDoc(new Deck(Pptx.read(bytes)), "Test", null);
        app.docs.push(doc); app.activate(doc);
      },
      bytes: () => { const b = app.deck.write(); let s = ""; for (let i = 0; i < b.length; i += 0x8000) s += String.fromCharCode(...b.subarray(i, i + 0x8000)); return btoa(s); },
      newDoc: () => { const doc = app.makeDoc(new Deck(), "Test", null); app.docs.push(doc); app.activate(doc); },
      openBytes: (name: string, bytes: Uint8Array) => { const doc = app.makeDoc(new Deck(Pptx.read(bytes)), stripExt(name), null); app.docs.push(doc); app.activate(doc); return doc.deck.slides.length; },
      /** Slide i (1-based) drawn as in the show at `width` CSS px, alone at the page's top left (render comparison). */
      renderAlone: (i: number, width: number) => {
        document.getElementById("render-alone")?.remove();
        const ref = app.deck.slides[i - 1];
        if (!ref) return null;
        const size = app.deck.pptx.slideSize;
        const s = width / (size.cx / EMU_PX);
        const box = h("div", { id: "render-alone", style: `position:fixed;left:0;top:0;z-index:99999;width:${width}px;height:${Math.round((size.cy / EMU_PX) * s)}px;overflow:hidden;background:#fff` });
        const r = renderSlide(app.deck.view(ref.path), size, { editing: false, slideNumber: i, image: (p) => app.image(p), pptx: app.deck.pptx });
        r.el.style.transform = `scale(${s})`; r.el.style.transformOrigin = "0 0";
        box.append(r.el);
        document.body.append(box);
        return { w: width, h: Math.round((size.cy / EMU_PX) * s), images: box.querySelectorAll("img").length };
      },
      clearAlone: () => document.getElementById("render-alone")?.remove(),
      /**
       * Where the lines of shape `id` on slide i (drawn by renderAlone) start and how high they are, as
       * PowerPoint's TextRange.Lines reports them: 1-based offsets over paragraphs joined by one character.
       */
      lineStarts: (id: number) => {
        const shp = document.querySelector(`#render-alone .shp[data-id="${id}"]`);
        if (!shp) return null;
        // Turned text is measured unturned (its lines then run across).
        const turned = Array.from(shp.querySelectorAll(".txt[data-turn]")) as HTMLElement[];
        for (const el of turned) el.style.rotate = "none";
        const out: { start: number; top: number }[] = [];
        let base = 0, lastTop = -1e9;
        const slideTop = (document.getElementById("render-alone") as HTMLElement).getBoundingClientRect().top;
        const scale = (document.getElementById("render-alone") as HTMLElement).getBoundingClientRect().width / (app.deck.pptx.slideSize.cx / 12700);
        for (const para of Array.from(shp.querySelectorAll(".para"))) {
          let paraLen = 0;
          lastTop = -1e9;
          for (const sp of Array.from(para.querySelectorAll(":scope > [data-o]")) as HTMLElement[]) {
            if (sp.classList.contains("end")) {
              if (paraLen === 0) out.push({ start: base + 1, top: (sp.getBoundingClientRect().top - slideTop) / scale });
              continue;
            }
            const o = Number(sp.dataset.o);
            const walker = document.createTreeWalker(sp, NodeFilter.SHOW_TEXT);
            let n: Node | null, acc = 0;
            while ((n = walker.nextNode())) {
              const t = n as Text;
              for (let k = 0; k < t.length; k++) {
                const r = document.createRange();
                r.setStart(t, k); r.setEnd(t, k + 1);
                const rect = r.getClientRects()[0];
                if (!rect) continue;
                if (rect.top > lastTop + rect.height * 0.5) { out.push({ start: base + o + acc + k + 1, top: (rect.top - slideTop) / scale }); lastTop = rect.top; }
              }
              acc += t.length;
            }
            paraLen = Math.max(paraLen, o + acc);
          }
          base += paraLen + 1;
        }
        for (const el of turned) el.style.rotate = `${el.dataset.turn}deg`;
        return out;
      },
      /**
       * A catalogue case's start (catalogue_present.py): a new presentation with `slides` slides of the
       * layout named, the shapes made as PowerPoint's COM calls make them (boxes in points), nothing selected.
       */
      setupCase: (spec: { layout: string; slides?: number; shapes: { kind: string; box?: number[]; text?: string; ph?: number }[]; zoom?: number }) => {
        const deck = new Deck();
        const lay = deck.layouts().find((l) => l.name === spec.layout)?.path ?? deck.layouts()[6].path;
        for (let k = 0; k < (spec.slides ?? 1); k++) deck.addSlide(k, lay);
        // The new presentation's own Titelfolie goes; the case's slides remain.
        deck.deleteSlides([0]);
        const doc = app.makeDoc(deck, "Katalog", null);
        // Earlier case documents go: only one presentation is open, as in PowerPoint's harness.
        app.docs = [doc];
        doc.savedVersion = deck.version;
        app.activate(doc);
        const path = deck.slides[0].path;
        deck.change("Setup", [path], () => {
          const doc = deck.pptx.doc(path);
          const tree = S.treeOf(doc);
          const phs = S.shapeEls(tree);
          for (const sh of spec.shapes) {
            const id = S.maxId(tree) + 1;
            const b = (sh.box ?? [0, 0, 100, 100]).map((v) => v * 12700);
            if (sh.kind === "ph") {
              const el = phs[sh.ph ?? 0];
              const body = el && S.txBodyOf(el);
              if (body && sh.text) { let pos = { p: 0, o: 0 }; sh.text.split("\r").forEach((line, i) => { if (i) pos = T.splitPara(body, pos); if (line) pos = T.insertText(body, pos, line); }); }
              continue;
            }
            let el: Element;
            if (sh.kind === "textbox") el = S.newTextBox(doc, id, b[0], b[1], b[2], b[3]);
            else if (sh.kind === "table") {
              app.ensureTableStyle();
              el = S.newTable(doc, id, 2, 2, { x: b[0], y: b[1], w: b[2] }, DEFAULT_TABLE_STYLE_ID);
            } else el = S.newShape(doc, id, sh.kind === "ellipse" ? "ellipse" : sh.kind === "line" ? "line" : "rect", { x: b[0], y: b[1], w: b[2], h: b[3], rot: 0, flipH: false, flipV: false });
            if (sh.kind === "textbox") S.bodyPrOf(el)!.setAttribute("wrap", "square");
            tree.append(el);
            const body = S.txBodyOf(el);
            if (body && sh.text) { let pos = { p: 0, o: 0 }; sh.text.split("\r").forEach((line, i) => { if (i) pos = T.splitPara(body, pos); if (line) pos = T.insertText(body, pos, line); }); }
          }
        });
        // Text boxes take the height of their text at once, as PowerPoint's AddTextbox does (spAutoFit).
        for (const v of deck.view(path).shapes) if (v.textBox) app.editor.updateAutofit(v.id);
        deck.clearHistory();
        doc.savedVersion = deck.version;
        app.editor.render(); app.panel.render();
        if (spec.zoom) app.setZoom(spec.zoom);
        app.editor.focus();
        return deck.slides.length;
      },
      /** The selection as PowerPoint's COM reports it: type, shape ids, text start (1-based) and length. */
      comSelection: () => {
        const ed = app.editor;
        if (app.panel.el.contains(document.activeElement)) return { type: "Slides", slides: [...app.panel.selected].sort((a, b) => a - b).map((i) => i + 1) };
        if (ed.text) {
          const body = ed.textBody(ed.text.id);
          const lin = (p: { p: number; o: number }) => (body ? T.paras(body).slice(0, p.p).reduce((s, q) => s + T.paraLength(q) + 1, 0) : 0) + p.o;
          const a = lin(ed.text.caret), b = lin(ed.text.anchor);
          return { type: "Text", ids: [ed.text.id], start: Math.min(a, b) + 1, length: Math.abs(a - b) };
        }
        return ed.sel.length ? { type: "Shapes", ids: [...ed.sel] } : { type: "None" };
      },
      slidePointToClient: (xpt: number, ypt: number) => {
        const r = (app.editor.el.querySelector(".stage .slide") as HTMLElement | null)?.getBoundingClientRect();
        if (!r) return null;
        const size = app.deck.pptx.slideSize;
        return { x: r.left + (xpt * 12700 / size.cx) * r.width, y: r.top + (ypt * 12700 / size.cy) * r.height };
      },
      /** The context menus by the names of PowerPoint's inventory (results/*-powerpoint-inventory.json). */
      contextMenus: async () => {
        const ev = new MouseEvent("contextmenu", { clientX: 10, clientY: 10 });
        const names = (menu: MenuEntry[]): { name: string; items?: unknown[] }[] => menu.flatMap((m) => {
          if (m === "-" || "heading" in m || "block" in m) return [];
          if ("cmd" in m) return [{ name: m.label ?? app.commands.get(m.cmd).label }];
          if ("icons" in m) return [{ name: "Einfügeoptionen" }];
          const sub = "submenu" in m ? (typeof m.submenu === "function" ? m.submenu() : m.submenu) : undefined;
          return [{ name: m.label, ...(sub ? { items: names(sub) } : {}) }];
        });
        const grab = (fn: () => void) => { app.capture = []; try { fn(); } finally { /* keep */ } const got = app.capture[0] ?? []; app.capture = null; return names(got); };
        const out: Record<string, unknown> = {};
        const setup = (shapes: { kind: string; box?: number[]; text?: string; ph?: number }[], layout = "Leer") => (window as unknown as { __present: { setupCase: (x: unknown) => void } }).__present.setupCase({ layout, shapes });
        out["Folienminiatur"] = grab(() => app.panelMenu(ev, true));
        out["Leerer Bereich im Folienbereich"] = grab(() => app.panelMenu(ev, false));
        setup([{ kind: "ph", ph: 0, text: "Titel" }], "Titelfolie");
        app.editor.select([2]); out["Platzhalterrahmen"] = grab(() => app.contextMenu(ev, "shape"));
        app.editor.enterText(2); out["Text in Bearbeitung"] = grab(() => app.contextMenu(ev, "text"));
        setup([{ kind: "rect", box: [100, 100, 100, 100] }]); app.editor.select([2]); out["Form"] = grab(() => app.contextMenu(ev, "shape"));
        setup([{ kind: "textbox", box: [100, 100, 200, 40], text: "x" }]); app.editor.select([2]); out["Textfeld"] = grab(() => app.contextMenu(ev, "shape"));
        setup([{ kind: "table", box: [100, 100, 300, 80] }]); app.editor.enterText(2, undefined, { r: 0, c: 0 }); out["Tabelle"] = grab(() => app.contextMenu(ev, "text"));
        setup([]); out["Leerer Folienbereich"] = grab(() => app.contextMenu(ev, "slide"));
        // Picture, group, chart and the notes pane (added with the inventory's later run, D-072).
        setup([{ kind: "rect", box: [100, 100, 100, 100] }, { kind: "ellipse", box: [300, 100, 100, 100] }]); app.editor.select([2, 3]); void app.commands.run("arrange.group");
        const grp = app.deck.view(app.slidePath()!).shapes.find((s) => s.kind === "grpSp"); if (grp) { app.editor.select([grp.id]); out["Gruppe"] = grab(() => app.contextMenu(ev, "shape")); }
        out["Notizenbereich"] = grab(() => app.notesMenu(ev));
        setup([]); (await import("./charts")).insertChart(app, "column");
        const chart = app.deck.view(app.slidePath()!).shapes.find((s) => s.frame === "chart"); if (chart) { app.editor.select([chart.id]); out["Diagramm"] = grab(() => app.contextMenu(ev, "shape")); }
        const fakePic = (): void => {
          const doc = app.deck.pptx.doc(app.slidePath()!); const tree = S.treeOf(doc);
          tree.append(S.newPicture(doc, S.maxId(tree) + 1, "rIdX", { x: 0, y: 0, w: 914400, h: 914400 }));
        };
        setup([]); app.editor.edit("Bild", fakePic); const pic = app.deck.view(app.slidePath()!).shapes.find((s) => s.kind === "pic"); if (pic) { app.editor.select([pic.id]); out["Bild"] = grab(() => app.contextMenu(ev, "shape")); }
        return out;
      },
      /** RaabePresent's ribbon as data for the inventory against PowerPoint (D-070), as RaabeCalc's (D-066). */
      inventory: () => {
        type Item = { name: string; type?: string; deferred?: boolean; enabled?: boolean; items?: Item[]; submenu?: Item[]; panel?: boolean };
        const cmdItem = (id: string, label?: string): Item => {
          const c = app.commands.get(id);
          // The accessible name keeps the "..." the visible label drops (as PowerPoint's).
          if (label && c.label === `${label}...`) label = c.label;
          return { name: label ?? c.label, deferred: !!c.deferred, enabled: app.commands.isImplemented(id) };
        };
        const entries = (menu: MenuEntry[]): Item[] => menu.flatMap((m): Item[] => {
          if (m === "-" || "heading" in m) return [];
          // Own content in a menu (the picture effect galleries): its tiles by their names.
          if ("block" in m) { try { return Array.from(m.block().querySelectorAll("[aria-label]")).map((b) => ({ name: b.getAttribute("aria-label")!.trim() })).filter((x) => x.name); } catch { return []; } }
          if ("cmd" in m) return [cmdItem(m.cmd, m.label)];
          if ("icons" in m) return [{ name: m.label, items: m.icons.map((i) => ({ name: i.tip.replace(/ \(.\)$/, "") })) }];
          const sub = "submenu" in m ? (typeof m.submenu === "function" ? m.submenu() : m.submenu) : undefined;
          return [{ name: m.label, enabled: !("disabled" in m && m.disabled), ...(sub ? { submenu: entries(sub) } : {}) }];
        });
        const ctl = (c: import("@raabe/shell").Control): Item[] => {
          if (c.type === "stack") return c.items.flatMap(ctl);
          if (c.type === "rows") return c.rows.flat().flatMap((x) => (x === "gap" ? [] : ctl(x)));
          if (c.type === "gallery") return c.items.map((i) => ({ ...cmdItem(i.cmd), type: "gallery" }));
          // A spin box: its label as PowerPoint names it (no colon) and its two arrow buttons.
          // A spinner's accessible name is its command's full name (PowerPoint: "Animationsdauer" beside "Dauer:").
          if (c.type === "spinner") return [{ ...cmdItem(c.cmd), type: "spinner" }, { name: "Mehr" }, { name: "Weniger" }];
          const base: Item = "cmd" in c ? cmdItem(c.cmd, "label" in c ? c.label : undefined) : { name: "?" };
          const menu = "menu" in c && c.menu.length ? { items: entries(c.menu) } : {};
          // Own panels (galleries, palettes): their entries as drawn, by their names.
          let panelItems: Item[] = [];
          try {
            const el = "popup" in c && c.popup ? c.popup(() => {}) : c.type === "color" && c.palette === "highlight"
              ? h("div", null, ...HIGHLIGHTS.map(([n, , label]) => h("button", { "aria-label": PP_HIGHLIGHT_NAMES[n] ?? label })), h("button", { "aria-label": "Keine Farbe" }))
              : c.type === "color" ? h("div", null, ...colorGrid(() => {}, app.paletteTheme()?.colors)) : c.type === "tiles" ? c.panel(() => {}) : null;
            if (el) panelItems = Array.from(el.querySelectorAll("button, [role=menuitem], [role=button], [data-tip]")).map((b) => ({ name: (b.getAttribute("aria-label") || (b as HTMLElement).dataset.tip || b.textContent || "").trim() })).filter((x) => x.name);
          } catch { /* a panel that needs a selection */ }
          if (c.type === "color" && c.extra) panelItems.push(...entries(c.extra()));
          const own = ("popup" in c && c.popup) || c.type === "color" || c.type === "tiles" ? { panel: true, ...(panelItems.length ? { items: [...(("items" in menu && menu.items) || []), ...panelItems] } : {}) } : {};
          return [{ ...base, type: (c as { type: string }).type, ...menu, ...own }];
        };
        const kind = app.masterKind;
        const masters = (["slides", "notes", "handout"] as const).map((k) => { app.masterKind = k; return app.masterTab(); });
        app.masterKind = kind;
        const tabs = [...ribbonTabs(app.ribbonContext()), app.shapeTab(), app.pictureTab(), app.tableTab(), app.tableLayoutTab(), app.chartTab(), app.chartFormatTab(), ...masters, equationTab(app)];
        return tabs.map((t) => ({
          name: t.label, contextual: !!t.contextual,
          groups: t.groups.map((g) => ({ name: g.label, launcher: g.launcher ? cmdItem(g.launcher.cmd) : undefined, controls: g.controls.flatMap(ctl) })),
        }));
      },
      handlePoint: (id: number, k: string) => {
        const el = app.editor.el.querySelector(`.overlay .handle[data-id="${id}"].h-${k}`) as HTMLElement | null;
        const r = el?.getBoundingClientRect();
        return r ? { x: r.left + r.width / 2, y: r.top + r.height / 2 } : null;
      },
      shapeRect: (id: number) => { const el = app.editor.el.querySelector(`.shp[data-id="${id}"]`) as HTMLElement | null; const r = el?.getBoundingClientRect(); return r ? { x: r.left, y: r.top, w: r.width, h: r.height } : null; },
      slideRect: () => { const r = (app.editor.el.querySelector(".slide") as HTMLElement | null)?.getBoundingClientRect(); return r ? { x: r.left, y: r.top, w: r.width, h: r.height } : null; },
      thumbRect: (i: number) => { const r = (app.panel.el.querySelector(`.thumb-item[data-i="${i - 1}"]`) as HTMLElement | null)?.getBoundingClientRect(); return r ? { x: r.left, y: r.top, w: r.width, h: r.height } : null; },
      run: (id: string) => app.commands.run(id),
      setZoom: (z: number) => app.setZoom(z),
      fit: () => app.fitWindow(),
      focus: () => app.editor.focus(),
      inShow: () => !!app.show,
      showState: () => app.show?.state() ?? null,
    };
  }
}

function sliderToZoom(v: number): number {
  return v <= 50 ? 10 + (v / 50) * 90 : 100 + ((v - 50) / 50) * 300;
}

function zoomToSlider(z: number): number {
  return z <= 100 ? ((z - 10) / 90) * 50 : 50 + ((z - 100) / 300) * 50;
}

function parseSeconds(v: string): number {
  const s = v.replace(",", ".").trim();
  const m = /^(\d+):(\d+(?:\.\d+)?)$/.exec(s);
  if (m) return Number(m[1]) * 60 + Number(m[2]);
  const n = Number(s.replace(/[^\d.]/g, ""));
  return Number.isFinite(n) ? n : 0;
}

function mimeOf(path: string): string {
  const ext = path.slice(path.lastIndexOf(".") + 1).toLowerCase();
  return ext === "png" ? "image/png" : ext === "jpg" || ext === "jpeg" ? "image/jpeg" : ext === "gif" ? "image/gif" : ext === "svg" ? "image/svg+xml" : ext === "bmp" ? "image/bmp" : ext === "emf" ? "image/x-emf" : ext === "wmf" ? "image/x-wmf" : ext === "webp" ? "image/webp" : "application/octet-stream";
}

export { closePopup, showPopup };
