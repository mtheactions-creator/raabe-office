// RaabePresent: wires the shell (title bar, ribbon, backstage) to presentations (D-070). Each open
// presentation is a document tab (D-010) shown in Normalansicht: the slide panel at the left, the
// slide in the middle, the notes below. Files are .pptx; everything not edited is kept (D-024).

import { Pptx, newPptx, PptxDom, type ShapeView, type DesignSpec } from "@raabe/ooxml";
import {
  Backstage, Commands, Keytips, Ribbon, TitleBar, baseName, h, icon, installTooltips, messageBox, showDialog, showMenu, showPopup, closePopup,
  type BackstageEntry, type MenuEntry, type Platform, type Tab,
} from "@raabe/shell";
import { SHAPE_GALLERY, shapeIcon } from "@raabe/text-engine";
import { COMMANDS } from "./command-list";
import { ribbonTabs, type RibbonContext } from "./ribbon-def";
import { Deck } from "./deck";
import { SlideEditor, flat, union, type EditorHost } from "./editor";
import { SlidePanel, type PanelHost } from "./panel";
import { renderSlide } from "./render/slide";
import { EMU_PX } from "./render/text";
import { DEFAULT_TABLE_STYLE_ID, DEFAULT_TABLE_STYLE_XML } from "./render/table";
import * as T from "./textops";
import * as S from "./shapeops";
import { DESIGNS } from "./designs";
import { SlideShow } from "./show";
import { FormatPane } from "./formatpane";
import { NotesPageView, OutlineView } from "./views";
import { FindDialog } from "./find";
import { Sorter } from "./sorter";
import { addRecent, loadRecent, loadSettings, removeRecent, setPinned, type Settings } from "./store";
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

const stripExt = (n: string) => n.replace(/\.pptx$/i, "");
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
  private slideClipboard: { xml: string; layout: string; text: string } [] | null = null;
  private painter: { rPr: Element | null; pPr: Element | null; spPr: Element | null; keep: boolean } | null = null;
  private lastFontColor = "FF0000";
  private lastHighlight = "FFFF00";
  private lastFill = "4472C4";
  private lastLine = "000000";
  private show: SlideShow | null = null;
  private formatPane!: FormatPane;

  constructor(readonly platform: Platform) {
    void platform.listFonts().then((f) => { this.fontList = f; }).catch(() => {});
    this.commands.register(...COMMANDS);
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
    this.ribbon = new Ribbon(this.commands, ribbonTabs(this.ribbonContext()), { onFile: () => this.backstage.open("home"), onDone: () => this.focusMain(), tabRowRight: [] });
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
    this.notesWrap = h("div", { class: "notes-wrap" }, this.notes);
    this.center = h("div", { class: "center" }, this.editor.el, this.notesWrap);
    this.outline = new OutlineView(this);
    this.notesPage = new NotesPageView(this);
    this.main = h("div", { class: "present-main" }, this.panel.el, this.outline.el, h("div", { class: "splitter" }), this.center, this.sorter.el, this.notesPage.el);
    this.formatPane = new FormatPane(this.editor, this.main);
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
  private masterList(): string[] {
    const m = this.deck.pptx.masters()[0];
    return m ? [m, ...this.deck.pptx.layouts(m)] : [];
  }
  masterPaths(): string[] | null { return this.masterMode ? this.masterList() : null; }

  setMasterView(on: boolean): void {
    this.editor.endText();
    this.editor.sel = [];
    this.masterMode = on;
    document.body.classList.toggle("master-view", on);
    if (on) {
      // PowerPoint opens the master view at the current slide's layout.
      const lay = this.deck.pptx.layoutOf(this.deck.slides[this.active.current]?.path ?? "");
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

  selectionChanged(): void {
    this.formatPane?.sync();
    this.updateChrome();
  }

  toolDone(): void {
    this.updateChrome();
  }

  doubleClick(v: ShapeView): boolean {
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
      return { xml: PptxDom.serialize(this.deck.pptx.doc(p)), layout: this.deck.pptx.layoutOf(p) ?? "", text: "" };
    });
    this.clipboard = null;
    void this.platform.clipboardWrite(sel.map((i) => this.slideText(i)).join("\r\n")).catch(() => {});
    if (cut) this.deleteSlides(sel);
  }

  private pasteSlides(): void {
    if (!this.slideClipboard) return;
    const at2 = this.active.current;
    // Pasted slides keep the copied XML; pictures and charts of another presentation are not carried yet.
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
        pPr.insertBefore(make(doc, "a:buFont", { typeface: value === "•" ? "Arial" : value === "§" || value === "Ø" || value === "ü" ? "Wingdings" : "Arial", panose: "020B0604020202020204", pitchFamily: "34", charset: "0" }), before);
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
        const f = hex ? fragment(el!.ownerDocument, `<a:solidFill><a:srgbClr val="${hex}"/></a:solidFill>`)[0] : make(el!.ownerDocument, "a:noFill");
        const after = kids(sp).find((c) => c.localName === "prstGeom" || c.localName === "custGeom") ?? kid(sp, "xfrm");
        sp.insertBefore(f, after ? after.nextSibling : sp.firstChild);
      }
    });
  }

  private setLine(hex: string | null, width?: number, dash?: string): void {
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
          ln.insertBefore(hex ? fragment(el!.ownerDocument, `<a:solidFill><a:srgbClr val="${hex}"/></a:solidFill>`)[0] : make(el!.ownerDocument, "a:noFill"), ln.firstChild);
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
    const key = [...kinds].sort().join(",") + (this.editor.text ? "+t" : "") + (this.masterMode ? "+m" : "");
    if (key === this.ctxKey) return;
    this.ctxKey = key;
    const tabs: Tab[] = [];
    if (kinds.has("sp") || kinds.has("cxnSp") || kinds.has("grpSp") || this.editor.text) tabs.push(this.shapeTab());
    if (kinds.has("pic")) tabs.push(this.pictureTab());
    if (kinds.has("table")) tabs.push(this.tableTab(), this.tableLayoutTab());
    if (kinds.has("chart")) tabs.push(this.chartTab(), this.chartFormatTab());
    if (this.masterMode) tabs.unshift(this.masterTab());
    this.ribbon.setContextual(tabs);
    if (this.masterMode && !kinds.size) this.ribbon.select("slideMaster");
  }

  private masterTabDef: Tab | null = null;
  masterTab(): Tab {
    const b = (cmd: string, key: string, size: "large" | "small" = "large") => ({ type: "button" as const, cmd, key, size });
    return this.masterTabDef ??= {
      id: "slideMaster", label: "Folienmaster", key: "M", contextual: true, shrink: [],
      groups: [
        { id: "edit", label: "Master bearbeiten", key: "ZM", icon: "slide-layout", controls: [b("master.insertMaster", "K"), b("master.insertLayout", "I"), { type: "stack", items: [b("master.delete", "D", "small"), b("master.rename", "R", "small"), b("master.preserve", "P", "small")] }] },
        { id: "layout", label: "Masterlayout", key: "ZL", icon: "slide-layout", controls: [b("master.masterLayout", "M"), b("master.insertPlaceholder", "A"), { type: "stack", items: [{ type: "check", cmd: "master.title", key: "T" }, { type: "check", cmd: "master.footers", key: "F" }] }] },
        { id: "edit2", label: "Design bearbeiten", key: "ZD", icon: "themes", controls: [{ type: "dropdown", cmd: "design.themes", key: "H", size: "large", menu: [], popup: (c) => this.themesGallery(c) }] },
        { id: "size", label: "Größe", key: "ZG", icon: "page-size", controls: [{ type: "dropdown", cmd: "design.slideSize", key: "S", size: "large", menu: [{ cmd: "design.slide4x3" }, { cmd: "design.slide16x9" }] }] },
        { id: "close", label: "Schließen", key: "ZS", icon: "close", controls: [b("master.close", "C")] },
      ],
    };
  }

  private shapeTabDef: Tab | null = null;
  shapeTab(): Tab {
    const b = (cmd: string, key: string, size: "large" | "small" = "small") => ({ type: "button" as const, cmd, key, size });
    return this.shapeTabDef ??= {
      id: "shapeFormat", label: "Formformat", key: "JD", contextual: true, shrink: [],
      groups: [
        { id: "insert", label: "Formen einfügen", key: "ZE", icon: "shapes", controls: [{ type: "dropdown", cmd: "insert.shapes", key: "SH", size: "large", menu: [], popup: (c) => this.shapeGallery(c) }, { type: "stack", items: [b("insert.textBox", "X")] }] },
        { id: "styles", label: "Formenarten", key: "ZF", icon: "shape-fill", launcher: { cmd: "draw.format" }, controls: [{ type: "stack", items: [
          { type: "color", cmd: "draw.fill", key: "SF", palette: "font", current: () => `#${this.lastFill}`, pick: (c) => this.setFill(c.replace("#", "")), showLabel: true },
          { type: "color", cmd: "draw.outline", key: "SO", palette: "font", current: () => `#${this.lastLine}`, pick: (c) => this.setLine(c.replace("#", "")), showLabel: true },
          { type: "dropdown", cmd: "draw.effects", key: "SE", size: "small", menu: this.effectsMenu() }] }] },
        { id: "arrange", label: "Anordnen", key: "ZA", icon: "bring-forward", controls: [{ type: "stack", items: [
          { type: "split", cmd: "arrange.forward", key: "AF", size: "small", menu: [{ cmd: "arrange.forward" }, { cmd: "arrange.front" }] },
          { type: "split", cmd: "arrange.backward", key: "AE", size: "small", menu: [{ cmd: "arrange.backward" }, { cmd: "arrange.back" }] },
          b("arrange.selectionPane", "AP")] }, { type: "stack", items: [
          { type: "dropdown", cmd: "arrange.align", key: "AA", size: "small", get menu() { return []; } },
          { type: "dropdown", cmd: "arrange.group", key: "AG", size: "small", menu: [{ cmd: "arrange.group" }, { cmd: "arrange.regroup" }, { cmd: "arrange.ungroup" }] },
          { type: "dropdown", cmd: "arrange.rotate", key: "AY", size: "small", menu: [{ cmd: "arrange.rotRight" }, { cmd: "arrange.rotLeft" }, { cmd: "arrange.flipV" }, { cmd: "arrange.flipH" }] }] }] },
        { id: "size", label: "Größe", key: "ZG", icon: "size-position", controls: [{ type: "stack", items: [
          { type: "spinner", cmd: "shape.height", label: "Höhe:", icon: "shape-height", key: "HE", value: () => this.sizeText("h"), commit: (v) => this.setSizeText("h", v), step: (d) => this.stepSizeCm("h", d) },
          { type: "spinner", cmd: "shape.width", label: "Breite:", icon: "shape-width", key: "WI", value: () => this.sizeText("w"), commit: (v) => this.setSizeText("w", v), step: (d) => this.stepSizeCm("w", d) }] }] },
      ],
    };
  }

  private pictureTabDef: Tab | null = null;
  pictureTab(): Tab {
    return this.pictureTabDef ??= {
      id: "pictureFormat", label: "Bildformat", key: "JP", contextual: true, shrink: [],
      groups: [
        { id: "adjust", label: "Anpassen", key: "ZA", icon: "corrections", controls: [{ type: "button", cmd: "pic.change", key: "G", size: "large" }, { type: "button", cmd: "pic.reset", key: "Q", size: "small" }] },
        { id: "styles", label: "Bildformatvorlagen", key: "ZB", icon: "picture-styles", controls: [{ type: "color", cmd: "draw.outline", key: "SO", palette: "font", current: () => `#${this.lastLine}`, pick: (c) => this.setLine(c.replace("#", "")), showLabel: true }] },
        { id: "arrange", label: "Anordnen", key: "ZO", icon: "bring-forward", controls: [{ type: "stack", items: [
          { type: "split", cmd: "arrange.forward", key: "AF", size: "small", menu: [{ cmd: "arrange.forward" }, { cmd: "arrange.front" }] },
          { type: "split", cmd: "arrange.backward", key: "AE", size: "small", menu: [{ cmd: "arrange.backward" }, { cmd: "arrange.back" }] },
          { type: "dropdown", cmd: "arrange.rotate", key: "AY", size: "small", menu: [{ cmd: "arrange.rotRight" }, { cmd: "arrange.rotLeft" }, { cmd: "arrange.flipV" }, { cmd: "arrange.flipH" }] }] }] },
        { id: "size", label: "Größe", key: "ZG", icon: "crop", controls: [{ type: "button", cmd: "pic.crop", key: "C", size: "large" }, { type: "stack", items: [
          { type: "spinner", cmd: "shape.height", label: "Höhe:", icon: "shape-height", key: "HE", value: () => this.sizeText("h"), commit: (v) => this.setSizeText("h", v), step: (d) => this.stepSizeCm("h", d) },
          { type: "spinner", cmd: "shape.width", label: "Breite:", icon: "shape-width", key: "WI", value: () => this.sizeText("w"), commit: (v) => this.setSizeText("w", v), step: (d) => this.stepSizeCm("w", d) }] }] },
      ],
    };
  }

  tableTab(): Tab {
    return { id: "tableDesign", label: "Tabellenentwurf", key: "JT", contextual: true, shrink: [], groups: [
      { id: "options", label: "Tabellenformatoptionen", key: "ZO", icon: "table-style", controls: [{ type: "stack", items: [{ type: "check", cmd: "table.firstRow", key: "A" }, { type: "check", cmd: "table.lastRow", key: "T" }, { type: "check", cmd: "table.bandRow", key: "R" }] }, { type: "stack", items: [{ type: "check", cmd: "table.firstCol", key: "F" }, { type: "check", cmd: "table.lastCol", key: "L" }, { type: "check", cmd: "table.bandCol", key: "C" }] }] },
    ] };
  }

  tableLayoutTab(): Tab {
    const b = (cmd: string, key: string) => ({ type: "button" as const, cmd, key, size: "large" as const });
    return { id: "tableLayout", label: "Layout", key: "JL", contextual: true, shrink: [], groups: [
      { id: "rows", label: "Zeilen und Spalten", key: "ZZ", icon: "table-insert-row", controls: [b("table.insertAbove", "A"), b("table.insertBelow", "BE"), b("table.insertLeft", "L"), b("table.insertRight", "R")] },
    ] };
  }

  chartTab(): Tab {
    return { id: "chartDesign", label: "Diagrammentwurf", key: "JC", contextual: true, shrink: [], groups: [{ id: "data", label: "Daten", key: "ZD", icon: "table", controls: [{ type: "button", cmd: "chart.editData", key: "D", size: "large" }] }] };
  }

  chartFormatTab(): Tab {
    return { id: "chartFormat", label: "Format", key: "JO", contextual: true, shrink: [], groups: [{ id: "current", label: "Aktuelle Auswahl", key: "ZA", icon: "format-picture", controls: [{ type: "button", cmd: "draw.format", key: "M", size: "large" }] }] };
  }

  private effectsMenu(): MenuEntry[] {
    return [{ label: "Schatten", submenu: [{ label: "Kein Schatten", run: () => this.setShadow(null) }, { label: "Außen: Versatz unten rechts", run: () => this.setShadow("outer") }] }];
  }

  private setShadow(kind: "outer" | null): void {
    const views = this.editor.selectedViews.filter((v) => v.kind === "sp" || v.kind === "pic");
    this.editor.edit("Formeffekte", () => {
      for (const v of views) {
        const sp = kid(this.editor.shapeEl(v.id)!, "spPr");
        if (!sp) continue;
        kid(sp, "effectLst")?.remove();
        if (!kind) { sp.append(make(sp.ownerDocument, "a:effectLst")); continue; }
        sp.append(fragment(sp.ownerDocument, '<a:effectLst><a:outerShdw blurRad="50800" dist="38100" dir="2700000" algn="tl" rotWithShape="0"><a:prstClr val="black"><a:alpha val="40000"/></a:prstClr></a:outerShdw></a:effectLst>')[0]);
      }
    });
  }

  private sizeText(k: "w" | "h"): string {
    const v = this.editor.selectedViews[0];
    if (!v) return "";
    return `${((k === "w" ? v.w : v.h) / 360000).toFixed(2).replace(".", ",")} cm`;
  }

  private setSizeText(k: "w" | "h", text: string): void {
    const cm = Number(text.replace(",", ".").replace(/[^\d.]/g, ""));
    if (!Number.isFinite(cm) || cm <= 0) return;
    this.setSizeEmu(k, cm * 360000);
  }

  private stepSizeCm(k: "w" | "h", d: 1 | -1): void {
    const v = this.editor.selectedViews[0];
    if (!v) return;
    this.setSizeEmu(k, Math.max(0, (k === "w" ? v.w : v.h) + d * 36000));
  }

  private setSizeEmu(k: "w" | "h", emu: number): void {
    const views = this.editor.selectedViews;
    this.editor.edit("Größe", () => {
      for (const v of views) {
        const el = this.editor.shapeEl(v.id);
        if (!el) continue;
        const keep = v.kind === "pic";
        const ratio = v.h ? v.w / v.h : 1;
        const w = k === "w" ? emu : keep ? emu * ratio : v.w, hh = k === "h" ? emu : keep ? emu / ratio : v.h;
        S.setXfrm(el, { x: v.x, y: v.y, w, h: hh });
      }
    });
  }

  // ----- galleries -------------------------------------------------------------------------------------------

  shapeGallery(close: () => void): HTMLElement {
    const wrap = h("div", { class: "shape-gallery-wrap" });
    for (const group of SHAPE_GALLERY) {
      wrap.append(h("div", { class: "gal-head" }, group.title));
      const grid = h("div", { class: "shape-gallery" });
      for (const it of group.items) {
        const btn = h("button", { "data-tip": it.label, "aria-label": it.label, onclick: () => {
          close();
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
      wrap.append(h("div", { class: "menu-sep" }), h("button", { class: "menu-item", onclick: () => { close(); this.duplicateSlides([...this.panel.selected]); } }, "Ausgewählte Folien duplizieren"));
    }
    return wrap;
  }

  private tableGrid(close: () => void): HTMLElement {
    const label = h("div", { class: "gal-head" }, "Tabelle einfügen");
    const grid = h("div", { class: "table-grid" });
    const cells: HTMLElement[] = [];
    for (let r = 0; r < 8; r++) for (let c = 0; c < 10; c++) {
      const cell = h("div", { class: "tg-cell" });
      cell.onmouseenter = () => { label.textContent = `${c + 1}x${r + 1}-Tabelle`; cells.forEach((x, i) => x.classList.toggle("on", Math.floor(i / 10) <= r && i % 10 <= c)); };
      cell.onclick = () => { close(); this.insertTable(r + 1, c + 1); };
      cells.push(cell);
      grid.append(cell);
    }
    return h("div", { class: "table-popup" }, label, grid, h("div", { class: "menu-sep" }),
      h("button", { class: "menu-item", onclick: () => { close(); void this.tableDialog(); } }, icon("table"), " Tabelle einfügen..."));
  }

  private async tableDialog(): Promise<void> {
    const cols = h("input", { type: "number", value: "5", min: "1", max: "75", style: "width:60px" }) as HTMLInputElement;
    const rows = h("input", { type: "number", value: "2", min: "1", max: "75", style: "width:60px" }) as HTMLInputElement;
    const r = await showDialog({ title: "Tabelle einfügen", body: h("div", { class: "form" }, h("label", null, "Spaltenanzahl: ", cols), h("label", null, "Zeilenanzahl: ", rows)), buttons: [{ label: "OK", value: true, primary: true }, { label: "Abbrechen", value: false }], cancel: false });
    if (r) this.insertTable(Math.max(1, Math.min(75, Number(rows.value) || 2)), Math.max(1, Math.min(75, Number(cols.value) || 5)));
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

  contextMenu(e: MouseEvent, kind: "shape" | "text" | "slide"): void {
    const at2 = new DOMRect(e.clientX, e.clientY, 0, 0);
    if (kind === "text") {
      showMenu(this.commands, [{ cmd: "edit.cut" }, { cmd: "edit.copy" }, { label: "Einfügeoptionen:", icons: [
        { icon: "paste-keep", tip: "Ursprüngliche Formatierung beibehalten (U)", run: () => void this.paste() },
        { icon: "paste-text", tip: "Nur den Text übernehmen (T)", run: () => void this.paste("text") }] }, "-",
        { cmd: "font.dialog" }, { cmd: "para.dialog" }, { label: "Aufzählungszeichen", submenu: this.bulletMenu() }, { label: "Nummerierung", submenu: this.numberMenu() }, "-",
        { cmd: "insert.link" }, "-", { cmd: "draw.format" }, { cmd: "review.newComment" }], at2);
      return;
    }
    if (kind === "shape") {
      const v = this.editor.selectedViews[0];
      const pic = v?.kind === "pic";
      const entries: MenuEntry[] = [{ cmd: "edit.cut" }, { cmd: "edit.copy" }, { label: "Einfügeoptionen:", icons: [{ icon: "paste-keep", tip: "Ursprüngliche Formatierung beibehalten (U)", run: () => void this.paste() }] }, "-"];
      if (pic) entries.push({ cmd: "pic.change" }, { cmd: "pic.save" }, "-");
      else if (v && (v.kind === "sp")) entries.push({ cmd: "shape.editText" }, ...(v.ph ? [] : [{ cmd: "shape.editPoints" } as MenuEntry]), "-");
      entries.push({ label: "Gruppieren", submenu: [{ cmd: "arrange.group" }, { cmd: "arrange.regroup" }, { cmd: "arrange.ungroup" }] },
        { label: "In den Vordergrund", submenu: [{ cmd: "arrange.front" }, { cmd: "arrange.forward" }] },
        { label: "In den Hintergrund", submenu: [{ cmd: "arrange.back" }, { cmd: "arrange.backward" }] }, "-",
        { cmd: "shape.link" }, "-", { cmd: "shape.altText" }, ...(v?.ph || pic ? [] : [{ cmd: "shape.setDefault" } as MenuEntry]), { cmd: "shape.sizePosition" },
        { cmd: pic ? "pic.format" : "draw.format" }, { cmd: "review.newComment" });
      showMenu(this.commands, entries, at2);
      return;
    }
    showMenu(this.commands, [{ cmd: "edit.cut" }, { cmd: "edit.copy" }, { label: "Einfügeoptionen:", icons: [{ icon: "paste-keep", tip: "Zieldesign verwenden (H)", run: () => void this.paste() }] }, "-",
      { label: "Layout", submenu: () => this.deck.layouts().map((l) => ({ label: l.name, run: () => this.deck.setLayout([this.active.current], l.path) })) },
      { cmd: "slide.reset" }, "-", { cmd: "view.gridDialog" }, { cmd: "slide.formatBackground" }, { cmd: "review.newComment" }], at2);
  }

  panelMenu(e: MouseEvent, onSlide: boolean): void {
    const at2 = new DOMRect(e.clientX, e.clientY, 0, 0);
    const entries: MenuEntry[] = onSlide
      ? [{ cmd: "edit.cut" }, { cmd: "edit.copy" }, { label: "Einfügeoptionen:", icons: [{ icon: "paste", tip: "Zieldesign verwenden (H)", run: () => this.pasteSlides() }] }, "-",
        { cmd: "slide.new" }, { cmd: "slide.duplicate" }, { cmd: "slide.delete" }, "-", { cmd: "slide.addSection" }, "-",
        { label: "Layout", submenu: () => this.deck.layouts().map((l) => ({ label: l.name, run: () => this.deck.setLayout([...this.panel.selected], l.path) })) },
        { cmd: "slide.reset" }, "-", { cmd: "slide.formatBackground" }, { cmd: "show.hideSlide" }, "-", { cmd: "review.newComment" }]
      : [{ label: "Einfügeoptionen:", icons: [{ icon: "paste", tip: "Zieldesign verwenden (H)", run: () => this.pasteSlides() }] }, "-", { cmd: "slide.new" }, { cmd: "slide.addSection" }];
    showMenu(this.commands, entries, at2);
  }

  private bulletMenu(): MenuEntry[] {
    const chars: [string, string][] = [["Ohne", ""], ["Gefüllte runde Aufzählungszeichen", "•"], ["Leere runde Aufzählungszeichen", "o"], ["Gefüllte quadratische Aufzählungszeichen", "§"], ["Pfeil-Aufzählungszeichen", "Ø"], ["Häkchen-Aufzählungszeichen", "ü"], ["Gefüllte Raute", "v"]];
    return chars.map(([label, c]) => ({ label, run: () => c ? this.toggleBullets("char", c) : this.formatParas("Aufzählungszeichen", (pPr) => { for (const k of kids(pPr)) if (["buChar", "buAutoNum", "buFont"].includes(k.localName)) k.remove(); if (!kid(pPr, "buNone")) pPr.insertBefore(make(pPr.ownerDocument, "a:buNone"), kids(pPr).find((x) => ["tabLst", "defRPr", "extLst"].includes(x.localName)) ?? null); }) }));
  }

  private numberMenu(): MenuEntry[] {
    const types: [string, string][] = [["1. 2. 3.", "arabicPeriod"], ["1) 2) 3)", "arabicParenR"], ["I. II. III.", "romanUcPeriod"], ["A. B. C.", "alphaUcPeriod"], ["a) b) c)", "alphaLcParenR"], ["a. b. c.", "alphaLcPeriod"], ["i. ii. iii.", "romanLcPeriod"]];
    return [{ label: "Ohne", run: () => this.toggleBullets("auto", "none") }, ...types.map(([label, t]) => ({ label, run: () => this.toggleBullets("auto", t) }))];
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
    impl("edit.undo", () => { ed().endText(); this.deck.undo(); }, () => this.deck.canUndo);
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
    impl("view.slideMaster", () => this.setMasterView(true), undefined, () => this.masterMode);
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
        // PowerPoint's Strg+D offset: 0,48 cm right and down (to be confirmed by the catalogue).
        S.setXfrm(copy, { x: v.x + 172720, y: v.y + 172720, w: v.w, h: v.h });
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
    if (k === "F12") return run("file.saveAs");
    if (k === "F7" && !shift) return run("review.spelling");
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
        case "g": return run("arrange.ungroup");
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
      case "g": return run("arrange.group");
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
      fontColor: () => `#${this.lastFontColor}`,
      setFontColor: (hex) => { this.lastFontColor = hex.replace("#", "").toUpperCase(); this.formatRuns("Schriftfarbe", (r) => T.setRunColor(r, this.lastFontColor)); },
      highlight: () => `#${this.lastHighlight}`,
      setHighlight: (hex) => { const v = hex.replace("#", "").toUpperCase(); this.lastHighlight = v; this.formatRuns("Texthervorhebungsfarbe", (r) => T.setHighlight(r, v === "" ? null : v)); },
      fillColor: () => `#${this.lastFill}`,
      setFill: (hex) => this.setFill(hex.replace("#", "").toUpperCase()),
      lineColor: () => `#${this.lastLine}`,
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
        direction: [["Horizontal", "horz"], ["Alle Texte um 90° drehen", "vert"], ["Alle Texte um 270° drehen", "vert270"], ["Gestapelt", "wordArtVert"]].map(([label, v]) => ({ label, run: () => this.setBodyAttr("vert", v === "horz" ? null : v) })),
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
        transOptions: () => [], animOptions: () => [], animStart: [{ label: "Beim Klicken", run: () => {} }, { label: "Mit Vorherigen", run: () => {} }, { label: "Nach Vorherigen", run: () => {} }],
        colors: () => DESIGNS.map((d) => ({ label: d.name === "Office" ? "Office" : d.name, run: () => this.applyDesign(d) })),
        fonts: () => DESIGNS.map((d) => ({ label: `${d.name}: ${d.major} / ${d.minor}`, run: () => this.applyDesign(d) })),
        underline: [], outlineWeight: [], outlineDash: [], effects: this.effectsMenu(), quickStyles: () => [], zoomFit: [],
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

  private designName(): string {
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
    const add = (title: string, items: [string, string][]) => {
      grid.append(h("div", { class: "gal-head" }, title));
      const row = h("div", { class: "fx-row" });
      for (const [label, kind] of items) row.append(h("button", { class: "fx-item", "aria-label": label, onclick: () => { close(); void import("./transitions").then((m) => m.setTransition(this, kind)); } }, icon(kind === "none" ? "close-x" : "transition"), h("span", null, label)));
      grid.append(row);
    };
    void import("./transitions").then((m) => {
      for (const g of m.TRANSITION_GROUPS) add(g.title, g.items);
    });
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
    void import("./animations").then((m) => {
      for (const g of m.ANIMATION_GROUPS) {
        grid.append(h("div", { class: "gal-head" }, g.title));
        const row = h("div", { class: "fx-row" });
        for (const it of g.items) row.append(h("button", { class: `fx-item fx-${g.cls}`, "aria-label": it.label, onclick: () => { close(); m.addAnimation(this, g.cls, it, add); } }, icon("animation"), h("span", null, it.label)));
        grid.append(row);
      }
    });
    return grid;
  }
  animationDuration(): string { return "00,50"; }
  setAnimationDuration(v: string): void { void import("./animations").then((m) => m.setDuration(this, parseSeconds(v))); }
  stepAnimationDuration(d: 1 | -1): void { void import("./animations").then((m) => m.setDuration(this, Math.max(0.01, parseSeconds(this.animationDuration()) + d * 0.25))); }
  animationDelay(): string { return "00,00"; }
  setAnimationDelay(v: string): void { void import("./animations").then((m) => m.setDelay(this, parseSeconds(v))); }
  stepAnimationDelay(d: 1 | -1): void { void import("./animations").then((m) => m.setDelay(this, Math.max(0, parseSeconds(this.animationDelay()) + d * 0.25))); }

  // ----- files ------------------------------------------------------------------------------------------------

  async openDialog(): Promise<void> {
    const path = await this.platform.openFileDialog("pptx");
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
    const doc = this.makeDoc(deck, stripExt(baseName(path)), path);
    doc.lastSaved = Date.now();
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
      path = await this.platform.saveFileDialog(`${doc.name}.pptx`, "pptx");
      if (!path) return false;
      if (!/\.pptx$/i.test(path)) path += ".pptx";
    }
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
    addRecent(path, baseName(path));
    void this.platform.deleteRecovery(doc.id).catch(() => {});
    this.updateTabs();
    return true;
  }

  private async closeDoc(doc: Doc): Promise<void> {
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
    }));
    return [
      { id: "home", label: "Startseite", icon: "home", page: () => h("div", null, h("h1", null, greeting()), h("h2", null, "Neu"), designs(), h("h2", null, "Zuletzt verwendet"), this.recentTable()) },
      { id: "new", label: "Neu", icon: "new", page: () => h("div", null, h("h1", null, "Neu"), designs()) },
      { id: "open", label: "Öffnen", icon: "open", page: () => h("div", null, h("h1", null, "Öffnen"), h("button", { class: "bigbtn", onclick: () => void this.openDialog() }, icon("open", "l"), "Durchsuchen"), h("h2", null, "Zuletzt verwendet"), this.recentTable()) },
      { id: "info", label: "Informationen", icon: "info", divider: true, page: () => this.infoPage() },
      { id: "save", label: "Speichern", icon: "save", cmd: "file.save" },
      { id: "saveAs", label: "Speichern unter", icon: "save-as", cmd: "file.saveAs" },
      { id: "print", label: "Drucken", icon: "print", page: () => h("div", null, h("h1", null, "Drucken"), h("button", { class: "bigbtn", onclick: () => void this.exportPdf() }, icon("print", "l"), "Drucken (als PDF)"),
        h("p", { class: "muted" }, "Ganzseitige Folien, eine pro Seite. Handzettel-Layouts folgen.")) },
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
      /** RaabePresent's ribbon as data for the inventory against PowerPoint (D-070), as RaabeCalc's (D-066). */
      inventory: () => {
        type Item = { name: string; type?: string; deferred?: boolean; enabled?: boolean; items?: Item[]; submenu?: Item[]; panel?: boolean };
        const cmdItem = (id: string, label?: string): Item => {
          const c = app.commands.get(id);
          return { name: label ?? c.label, deferred: !!c.deferred, enabled: app.commands.isImplemented(id) };
        };
        const entries = (menu: MenuEntry[]): Item[] => menu.flatMap((m): Item[] => {
          if (m === "-" || "heading" in m || "block" in m) return [];
          if ("cmd" in m) return [cmdItem(m.cmd, m.label)];
          if ("icons" in m) return [{ name: m.label, items: m.icons.map((i) => ({ name: i.tip.replace(/ \(.\)$/, "") })) }];
          const sub = "submenu" in m ? (typeof m.submenu === "function" ? m.submenu() : m.submenu) : undefined;
          return [{ name: m.label, enabled: !("disabled" in m && m.disabled), ...(sub ? { submenu: entries(sub) } : {}) }];
        });
        const ctl = (c: import("@raabe/shell").Control): Item[] => {
          if (c.type === "stack") return c.items.flatMap(ctl);
          if (c.type === "rows") return c.rows.flat().flatMap((x) => (x === "gap" ? [] : ctl(x)));
          if (c.type === "gallery") return c.items.map((i) => ({ ...cmdItem(i.cmd), type: "gallery" }));
          const base: Item = "cmd" in c ? cmdItem(c.cmd, "label" in c ? c.label : undefined) : { name: "?" };
          const menu = "menu" in c && c.menu.length ? { items: entries(c.menu) } : {};
          const own = ("popup" in c && c.popup) || c.type === "color" ? { panel: true } : {};
          return [{ ...base, type: (c as { type: string }).type, ...menu, ...own }];
        };
        const tabs = [...ribbonTabs(app.ribbonContext()), app.shapeTab(), app.pictureTab(), app.tableTab(), app.tableLayoutTab(), app.chartTab(), app.chartFormatTab()];
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
