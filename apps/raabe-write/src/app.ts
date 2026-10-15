// RaabeWrite: wires the shell (title bar, ribbon, backstage, status bar) to documents.
// Each open document is a tab with its own editor. Files are .docx (D-004); content the
// editor cannot change yet is preserved verbatim and written back unchanged (D-024).

import { Doc as TextDoc, Editor, Fonts, PT, seqToLatex, seqToLinear, type CharProps, type CopiedFormat, type DocState, type LineRule, type Para, type ParaProps } from "@raabe/text-engine";
import { ommlToSeq, ommlToWordHtml, readDocx, seqToOmml, wordHtmlMath, writeDocx, type DocxPackage } from "@raabe/ooxml";
import {
  Backstage, Commands, Keytips, Ribbon, StatusBar, TitleBar, baseName, dirName, h, icon, installTooltips, messageBox,
  showDialog, showFontColors, type BackstageEntry, type Platform,
} from "@raabe/shell";
import { COMMANDS } from "./command-list";
import type { AppApi, FontTarget, ParaTarget } from "./features/api";
import { headerFooterFeature, HEADER_COMMANDS } from "./features/headerfooter";
import { findFeature, FIND_COMMANDS } from "./features/find";
import { listsFeature, LIST_COMMANDS } from "./features/lists";
import { pageSetupFeature, PAGE_COMMANDS } from "./features/pagesetup";
import { picturesFeature, PICTURE_COMMANDS, PICTURE_TAB_COMMANDS } from "./features/pictures";
import { mathFeature, MATH_COMMANDS } from "./features/math";
import { symbolsFeature } from "./features/symbols";
import { linksFeature } from "./features/links";
import { captionsFeature } from "./features/captions";
import { citationsFeature } from "./features/citations";
import { notesDialogFeature } from "./features/notes-dialog";
import { sortFeature } from "./features/sort";
import { selectionPaneFeature } from "./features/selection-pane";
import { stylesPaneFeature } from "./features/styles-pane";
import { clipboardPaneFeature } from "./features/clipboard-pane";
import { makeDock } from "./features/dock";
import { viewsFeature } from "./features/views";
import { designFeature, newDocumentStyles } from "./features/design";
import { backgroundFeature } from "./features/background";
import { coverPageFeature } from "./features/cover-page";
import { bordersFeature } from "./features/borders";
import { textEffectsFeature } from "./features/text-effects";
import { formatPaneFeature } from "./features/format-pane";
import { groupsFeature, GROUP_COMMANDS } from "./features/groups";
import { hyphenationFeature, HYPHENATION_COMMANDS } from "./features/hyphenation";
import { rulerFeature } from "./features/ruler";
import { referencesFeature, REFERENCE_COMMANDS } from "./features/references";
import { reviewFeature, REVIEW_COMMANDS } from "./features/review";
import { spellFeature, type Word as SpellWord } from "./features/spell";
import { thesaurusFeature } from "./features/thesaurus";
import { editorPaneFeature, type Issue } from "./features/editor-pane";
import { grammarFeature } from "./features/grammar";
import { contextFeature, CONTEXT_COMMANDS } from "./features/context";
import { shapesFeature, SHAPE_COMMANDS } from "./features/shapes";
import { tablesFeature, TABLE_COMMANDS } from "./features/tables";
import { GALLERY, miniToolbar, ribbonTabs, type RibbonContext } from "./ribbon-def";
import { applyTips } from "./tooltips";
import { TEMPLATES, finishTemplate, templateState, type TemplateId } from "./templates";
import { addRecent, loadRecent, loadSettings, removeRecent, saveSettings, setPinned, type Settings } from "./store";
import appIconUrl from "../../../branding/raabe-write-small.svg?url";
import logoUrl from "../../../branding/raabe-write.svg?url";

/** A sketch of a template's first page for the Neu gallery. */
function templateThumb(t: TemplateId): HTMLElement {
  const bar = (w: number, color = "#c8c8c8", hgt = 3, mt = 5, ml = 0) =>
    h("div", { style: `width:${w}%;height:${hgt}px;background:${color};margin:${mt}px 0 0 ${ml}%` });
  const red = "#c4161c", blue = "#2f5496";
  const box = () => h("div", { style: "flex:1;height:9px;border:1px solid #bdbdbd" });
  const row = (...kids: HTMLElement[]) => h("div", { style: "display:flex;gap:2px;margin-top:5px" }, ...kids);
  const parts: Record<TemplateId, HTMLElement[]> = {
    arbeitsblatt: [row(box(), box(), box()), bar(70, "#404040", 6, 9), bar(45, "#9e9e9e", 3, 4), bar(30, blue, 4, 9), bar(95), bar(80),
      bar(30, blue, 4, 10), bar(95), bar(60), bar(30, blue, 4, 10), bar(90)],
    facharbeit: [bar(40, "#404040", 4, 14, 30), bar(60, "#9e9e9e", 3, 6, 20), bar(80, red, 6, 22, 10), bar(50, "#9e9e9e", 3, 6, 25),
      bar(45, "#c8c8c8", 3, 26), bar(40), bar(50), bar(35)],
    versuchsprotokoll: [bar(80, "#404040", 6, 4), row(box(), box()), row(box(), box()), bar(35, blue, 4, 9), bar(90), bar(35, blue, 4, 8), bar(60),
      bar(35, blue, 4, 8), bar(50, "#c8c8c8", 3, 4, 8), bar(50, "#c8c8c8", 3, 3, 8)],
    handout: [bar(75, "#404040", 6, 4), bar(60, "#9e9e9e", 3, 4), bar(30, blue, 4, 9), bar(60, "#c8c8c8", 3, 4, 8), bar(55, "#c8c8c8", 3, 3, 8),
      bar(40, blue, 4, 8), row(box(), box()), row(box(), box()), bar(35, blue, 4, 8), bar(70, "#c8c8c8", 3, 4, 8)],
  };
  return h("div", { class: "thumb", style: "flex-direction:column;align-items:stretch;justify-content:flex-start;padding:10px 11px;box-sizing:border-box;position:relative" },
    ...parts[t], h("img", { src: logoUrl, width: 14, alt: "", style: "position:absolute;right:6px;bottom:6px;opacity:.85" }));
}

interface Doc {
  id: string;
  name: string;
  path: string | null;
  editor: Editor;
  ws: HTMLElement;
  /** Content at the last save (all stories, styles, settings); dirty when the editor's differs. */
  saved: readonly unknown[];
  recoverySaved: readonly unknown[] | null;
  /** The original file's package (all its parts), for lossless saving. */
  pkg: DocxPackage | null;
  /** Content RaabeWrite shows only partly (German labels); it is saved unchanged. */
  partial: string[];
  lastSaved: number | null;
}

const stripExt = (name: string) => name.replace(/\.docx$/i, "");
const samePath = (a: string, b: string) => a.toLowerCase() === b.toLowerCase();
const fileSafe = (s: string) => s.replace(/[\\/:*?"<>|\t]/g, "").replace(/\s+/g, " ").trim().slice(0, 100);

export class WriteApp {
  readonly commands = new Commands();
  private docs: Doc[] = [];
  private active: Doc | null = null;
  private settings: Settings = loadSettings();
  private untitled = 0;
  private titleBar: TitleBar;
  private ribbon: Ribbon;
  private statusBar: StatusBar;
  private backstage: Backstage;
  private infobarSlot: HTMLElement;
  private printRoot: HTMLElement;
  private wordTimer = 0;
  private refreshTimer = 0;
  private lastTabsKey = "";
  private readonly fonts: Fonts;
  private fontList: string[] = [];
  private recentFonts: string[] = [];
  private lastFontColor = "FF0000";
  private lastHighlight = "yellow";
  /** Strg+Alt+C (Format kopieren), shared by all documents as in Word. */
  private copiedFormat: CopiedFormat | null = null;
  private lastPainterClick = 0;
  private lastShading = "FFFF00";
  private features!: ReturnType<WriteApp["makeFeatures"]>;
  /** Kommentare pane (right of the document). */
  private readonly commentPane = h("div", { class: "cpane", hidden: true });
  private readonly docArea: HTMLElement = h("div", { class: "docarea" }, this.commentPane);
  private readonly dock = makeDock(this.docArea);
  private find!: ReturnType<typeof findFeature>;

  constructor(private platform: Platform, initial?: Partial<DocState>) {
    this.fonts = new Fonts((families) => platform.fontMetrics(families));
    this.fonts.onChange(() => { for (const d of this.docs) d.editor.relayout(); });
    void platform.listFonts().then((f) => { this.fontList = f; }).catch(() => {});
    this.commands.register(...COMMANDS, ...TABLE_COMMANDS, ...PAGE_COMMANDS, ...HEADER_COMMANDS, ...PICTURE_COMMANDS, ...FIND_COMMANDS,
      ...REFERENCE_COMMANDS, ...REVIEW_COMMANDS, ...MATH_COMMANDS, ...CONTEXT_COMMANDS, ...LIST_COMMANDS, ...PICTURE_TAB_COMMANDS, ...SHAPE_COMMANDS, ...GROUP_COMMANDS, ...HYPHENATION_COMMANDS);
    this.implementCommands();
    this.features = this.makeFeatures();
    applyTips(this.commands);
    // Formformat's Anordnen uses Bildformat's panels.
    this.features.shapes.layoutCommands.setPositionPanel(this.features.pictures.arrange.positionPanel);
    this.features.shapes.layoutCommands.setAlignPanel(this.features.pictures.arrange.alignPanel);

    this.titleBar = new TitleBar(this.commands, platform, {
      appIcon: appIconUrl,
      qat: [
        { cmd: "file.save", key: "2" },
        { cmd: "edit.undo", key: "3", menu: [{ cmd: "edit.undo" }] },
        { cmd: "edit.redo", key: "4" },
      ],
      onSelectDoc: (id) => this.activate(this.docs.find((d) => d.id === id)!),
      onCloseDoc: (id) => void this.closeDoc(this.docs.find((d) => d.id === id)!),
      onNewDoc: () => this.newDoc(),
      onSearchDone: () => this.active?.editor.focus(),
    });
    const comments = h("button", { class: "pill", "data-tip": "Kommentare\nKommentare in diesem Dokument anzeigen und bearbeiten.",
      onclick: () => this.features.review.togglePane() }, icon("comment"), "Kommentare");
    this.ribbon = new Ribbon(this.commands, ribbonTabs(this.ribbonContext()), {
      onFile: () => this.openBackstage("home"),
      onDone: () => this.focusDoc(),
      tabRowRight: [comments],
    });
    this.statusBar = new StatusBar(this.commands, (z) => this.setZoom(z), [{ cmd: "view.readMode" }, { cmd: "view.printLayout" }, { cmd: "view.webLayout" }], "view.zoomDialog");
    this.backstage = new Backstage(this.commands, this.backstageEntries());
    this.backstage.onClose(() => this.active?.editor.focus());
    this.find = findFeature(this.api(), this.docArea);
    // Task panes: left Navigation, Zwischenablage; right Formatvorlagen, Auswahl, Kommentare.
    const pane = (sel: string) => this.docArea.querySelector<HTMLElement>(sel);
    for (const [sel, side, size] of [[".navpane", "left", 300], [".clippane", "left", 270], [".stylepane", "right", 260], [".selpane", "right", 240], [".cpane", "right", 340]] as const) {
      const el = pane(sel);
      if (el) this.dock.add(el, side, size);
    }
    this.infobarSlot = h("div");
    this.printRoot = h("div", { id: "print-root" });

    document.body.append(this.titleBar.el, this.ribbon.tabRow, this.ribbon.card, this.infobarSlot, this.docArea, this.statusBar.el,
      this.backstage.el, this.printRoot);
    document.body.classList.toggle("touch", this.settings.touchMode);
    installTooltips();
    new Keytips(
      (el) => el.dataset.keytipScope === "root" && el.getAttribute("role") === "tab" ? this.ribbon.activeTab : null,
      () => this.active?.editor.focus(),
    );
    this.installKeyboard();
    this.installFocusKeeping();

    // Ctrl + mouse wheel zooms, as in Word.
    this.docArea.addEventListener("wheel", (e) => {
      if (!e.ctrlKey || !this.active) return;
      e.preventDefault();
      this.setZoom(Math.round((this.active.editor.zoom + (e.deltaY < 0 ? 0.1 : -0.1)) * 10) / 10);
    }, { passive: false });
    addEventListener("afterprint", () => this.printRoot.replaceChildren());

    platform.window.onCloseRequested(() => this.confirmQuit());
    setInterval(() => void this.autoRecover(), Math.max(10, this.settings.autoRecoverSeconds) * 1000);

    this.newDoc(initial);
    void this.offerRecovery();
    platform.onOpenFiles((paths) => { for (const p of paths) void this.openPath(p); });
  }

  // ----- features (Step 3) --------------------------------------------------------------

  private api(): AppApi {
    return {
      commands: this.commands, platform: this.platform,
      editor: () => this.active?.editor,
      focus: () => this.focusDoc(),
      refresh: () => this.updateChrome(false),
      selectTab: (id) => { this.updateContextualTabs(); this.ribbon?.select(id); },
      clipboard: () => this.readClipboard(false),
      fontDialog: (target) => this.fontDialog(target),
      paragraphDialog: (target) => this.paragraphDialog(target),
      fontList: () => (this.fontList.length ? this.fontList : ["Arial", "Calibri", "Calibri Light", "Cambria", "Times New Roman"]),
      documentPath: () => this.active?.path ?? null,
      createEmptyDocument: async (path, open) => {
        await this.platform.writeFile(path, writeDocx(new TextDoc({}).state(), null));
        if (open) await this.openPath(path);
      },
    };
  }

  /** Text, HTML and a picture from the system clipboard (the platform's text as fallback). */
  private async readClipboard(plain: boolean): Promise<{ text: string; html: string; image: Blob | null }> {
    let html = "";
    let image: Blob | null = null;
    if (!plain) {
      try { if (this.platform.clipboardReadHtml) html = await this.platform.clipboardReadHtml(); } catch { /* the browser's API below */ }
      try {
        for (const item of await navigator.clipboard.read()) {
          if (item.types.includes("text/html") && !html) html = await (await item.getType("text/html")).text();
          const type = item.types.find((t) => t.startsWith("image/"));
          if (type && !image) image = await item.getType(type);
        }
      } catch { /* no HTML access: text only */ }
    }
    const text = await this.platform.clipboardRead().catch(() => "");
    if (!image && !plain && this.platform.clipboardReadImage) {
      const b = await this.platform.clipboardReadImage().catch(() => null);
      if (b) image = new Blob([b as BlobPart]);
    }
    return { text, html, image };
  }

  /** Back to the document: the equation being edited, else the text. */
  private focusDoc(): void {
    if (this.features?.math.editing()) this.features.math.focus();
    else this.active?.editor.focus();
  }

  private makeFeatures() {
    const api = this.api();
    const spell = spellFeature(api);
    const math = mathFeature(api);
    const thesaurus = thesaurusFeature(api, this.docArea, this.dock);
    const grammar = grammarFeature(api, spell);
    const asWord = (i: Issue): SpellWord => ({ a: i.a, b: i.b, text: i.text, lang: i.lang, repeat: i.kind === "repeat" });
    const editorPane = editorPaneFeature(api, this.docArea, this.dock, [{
      prepare: (e) => spell.checkAll(e),
      issues: (e) => spell.issues(e).map(({ p, w }) => ({ p, a: w.a, b: w.b, kind: w.repeat ? "repeat" : "spell", text: w.text, lang: w.lang })),
      suggestions: async (i) => (i.kind === "repeat" ? [""] : spell.suggest(asWord(i))),
      ignoreOnce: (e, i) => spell.ignoreOnce(e, i.p, i.a),
      ignoreAll: (_e, i) => spell.ignoreAll(i.text),
      add: (i) => spell.addToDictionary(i.text),
      apply: (e, i, s) => { if (i.kind !== "repeat" || s) return false; spell.deleteRepeat(e, i.p, asWord(i)); return true; },
    }, grammar.source], () => void this.optionsDialog("Dokumentprüfung"));
    const textEffects = textEffectsFeature(api, this.docArea, this.dock);
    const formatPane = formatPaneFeature(api, this.docArea, this.dock, textEffects);
    // With a shape selected, the text effects' "…optionen" open Form formatieren > Textoptionen (as in Word).
    textEffects.redirect((t) => { if (!api.editor()?.selectedShape()) return false; formatPane.open(t === "fill" ? "textFill" : "textEffects"); return true; });
    return {
      context: contextFeature(api, async (e, pos) => { const s = await spell.spellEntries(e, pos); return s.length ? s : grammar.entries(e, pos); },
        (kind) => this.ribbon?.toolbar(miniToolbar(this.ribbonContext(), kind)).el ?? null,
        (x, y) => math.contextMenu(x, y), thesaurus.synonymEntries),
      thesaurus,
      editorPane,
      grammar,
      lists: listsFeature(api, (start, pick) => void this.features.symbols.dialog({ start, pick })),
      tables: tablesFeature(api),
      pages: pageSetupFeature(api),
      hf: headerFooterFeature(api),
      pictures: picturesFeature(api),
      shapes: shapesFeature(api),
      spell,
      refs: referencesFeature(api),
      math,
      symbols: symbolsFeature(api),
      links: linksFeature(api),
      captions: captionsFeature(api),
      citations: citationsFeature(api),
      notesDialog: notesDialogFeature(api),
      sort: sortFeature(api),
      selection: selectionPaneFeature(api, this.docArea),
      styles: stylesPaneFeature(api, this.docArea),
      review: reviewFeature(api, () => this.settings.userName, this.commentPane, this.docArea, () => void this.optionsDialog(), this.dock),
      clipboard: clipboardPaneFeature(api, this.docArea),
      ruler: rulerFeature(api, this.docArea, this.dock),
      design: designFeature(api),
      background: backgroundFeature(api),
      borders: bordersFeature(api),
      textEffects,
      formatPane,
      groups: groupsFeature(api),
      cover: coverPageFeature(api),
      hyphenation: hyphenationFeature(api),
      views: viewsFeature(api, {
        titleBar: () => this.titleBar.el, docArea: this.docArea,
        openBackstage: () => this.openBackstage("info"), setZoom: (z) => this.setZoom(z),
      }),
    };
  }

  /** Contextual tabs for where the caret is: table, header/footer, picture. */
  private updateContextualTabs(): void {
    const e = this.active?.editor;
    if (!this.ribbon || !e) return;
    const tabs = [];
    if (e.isHeaderFooter()) tabs.push(this.features.hf.tab);
    if (e.inTable()) tabs.push(this.features.tables.designTab, this.features.tables.layoutTab);
    // A shape (also while its text is edited) shows Formformat, a picture Bildformat.
    if (e.selectedShape() || e.selectedGroup()) tabs.push(this.features.shapes.tab);
    else if (e.selectedImage()) tabs.push(this.features.pictures.tab);
    if (this.features.math.editing()) tabs.push(this.features.math.tab);
    const wasHf = this.contextualKey.includes("headerFooter");
    this.ribbon.setContextual(tabs);
    this.contextualKey = tabs.map((t) => t.id).join(",");
    // Opening a header or footer switches to its tab, as in Word; a new picture shows Bildformat.
    if (!wasHf && e.isHeaderFooter()) this.ribbon.select("headerFooter");
  }

  private contextualKey = "";

  // ----- documents ----------------------------------------------------------------------

  private createDoc(name: string, path: string | null, state: Partial<DocState> | undefined, pkg: DocxPackage | null = null, partial: string[] = []): Doc {
    const ws = h("div", { class: "workspace", tabIndex: -1 });
    this.docArea.append(ws);
    const editor = new Editor(ws, state ?? {}, this.fonts);
    // A double-click on a picture or shape opens its tab, as in Word.
    // Rücktaste on cells that are not whole rows or columns: Word's Zellen löschen dialog.
    editor.onDeleteCells = () => void this.commands.run("table.deleteCells");
    // Rahmen übertragen and the Rahmenpipette start and end in the editor too (Esc, a click).
    editor.onBorderPainter = () => this.commands.refresh();
    editor.onObjectOpen = (kind) => { this.updateContextualTabs(); this.ribbon?.select(kind === "picture" ? "pictureFormat" : "shapeFormat"); };
    this.features.spell.attach(editor);
    // Equations go to and come from Word as Word's own HTML (OMML in a conditional comment).
    editor.mathClipboard = {
      toHtml: (o, alone) => {
        if (!o.math) return undefined;
        const omml = ommlToWordHtml(seqToOmml(o.math, alone && !o.mathInline));
        if (!omml) return undefined;
        const lin = seqToLinear(o.math).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
        return `<!--[if gte msEquation 12]>${omml}<![endif]--><![if !msEquation]><span style='font-family:"Cambria Math",serif'>${lin}</span><![endif]>`;
      },
      fromOmml: (fragment) => {
        const el = wordHtmlMath(fragment);
        const math = el ? ommlToSeq(el) : null;
        return el && math ? { kind: "math", data: undefined, display: "", latex: seqToLatex(math), math, mathInline: el.localName === "oMath" || undefined } : null;
      },
    };
    this.features.grammar.attach(editor);
    this.features.context.attach(editor);
    this.features.pictures.attach(editor);
    this.features.tables.attach(editor);
    const pf = this.platform;
    if (pf.clipboardReadHtml && pf.clipboardWriteHtml) {
      editor.clipboardHost = {
        readHtml: () => pf.clipboardReadHtml!(), write: (text, html) => pf.clipboardWriteHtml!(html, text),
        readImage: async () => { const b = await pf.clipboardReadImage?.(); return b ? new Blob([b as BlobPart]) : null; },
      };
    }
    this.features.review.attach(editor);
    editor.onCopied = (text, html, object) => this.features.clipboard.collect(text, html, object);
    this.features.math.attach(editor, !path && !pkg);
    this.features.links.attach(editor);
    const doc: Doc = {
      id: crypto.randomUUID(), name, path, editor, ws, saved: editor.contentStamp(), recoverySaved: null, pkg, partial,
      lastSaved: null,
    };
    editor.onChange((e) => this.onEditorChange(doc, e.content));
    this.docs.push(doc);
    this.activate(doc);
    return doc;
  }

  newDoc(state?: Partial<DocState>): Doc {
    this.untitled++;
    // A blank document takes what Entwurf > Als Standard festlegen kept.
    const styles = state ? undefined : newDocumentStyles();
    return this.createDoc(`Dokument${this.untitled}`, null, styles ? { styles } : state);
  }

  private isDirty(d: Doc): boolean {
    const now = d.editor.contentStamp();
    return now.length !== d.saved.length || now.some((x, i) => x !== d.saved[i]);
  }

  private isPristine(d: Doc): boolean {
    const p = d.editor.paragraphs();
    return d.path === null && !this.isDirty(d) && p.length === 1 && p[0].text === "";
  }

  private activate(doc: Doc): void {
    this.active = doc;
    for (const d of this.docs) d.ws.hidden = d !== doc;
    doc.editor.update();
    doc.editor.focus();
    this.features.review.render();
    this.renderInfobar();
    this.updateStatus(doc, true);
    this.updateChrome(false);
  }

  private onEditorChange(doc: Doc, content: boolean): void {
    if (doc !== this.active) return;
    // Keep the typing path lean: chrome updates are deferred and coalesced.
    clearTimeout(this.refreshTimer);
    this.refreshTimer = window.setTimeout(() => this.updateChrome(content), content ? 80 : 30);
  }

  private updateChrome(countWords: boolean): void {
    const d = this.active;
    if (!d) return;
    const tabsKey = this.docs.map((x) => `${x.id}:${x.name}:${this.isDirty(x)}`).join("|") + `@${d.id}`;
    if (tabsKey !== this.lastTabsKey) {
      this.lastTabsKey = tabsKey;
      this.titleBar.setDocs(this.docs.map((x) => ({ id: x.id, name: x.name, dirty: this.isDirty(x) })), d.id);
      this.platform.window.setTitle(`${d.name} - RaabeWrite`);
    }
    this.updateStatus(d);
    this.updateContextualTabs();
    if (countWords) this.find?.refresh();
    this.features.selection.render();
    this.features.styles.render();
    this.features.ruler.refresh();
    this.features.textEffects.refresh();
    this.features.formatPane.refresh();
    if (countWords && this.wordCache?.paras !== d.editor.paragraphs()) {
      // Counting runs after a short pause in typing, not on every key.
      clearTimeout(this.wordTimer);
      this.wordTimer = window.setTimeout(() => { if (this.active) this.updateStatus(this.active, true); }, 250);
    }
    this.commands.refresh();
  }

  private wordCache: { paras: readonly Para[]; words: number } | null = null;

  private updateStatus(d: Doc, recount = false): void {
    const paras = d.editor.paragraphs();
    if (recount || !this.wordCache) {
      this.wordCache = { paras, words: d.editor.countWords() };
    }
    const s = d.editor.stats();
    this.statusBar.update({ page: s.caretPage + 1, pages: s.pageCount, words: this.wordCache.words, zoom: d.editor.zoom, pageLabel: this.features.views.statusLabel(),
      language: this.features.spell.statusLanguage(d.editor) });
  }

  private renderInfobar(): void {
    this.infobarSlot.replaceChildren();
    const d = this.active;
    if (!d || !d.partial.length) return;
    const close = h("button", { onclick: () => { d.partial = []; this.renderInfobar(); } }, "Schließen");
    this.infobarSlot.append(h("div", { class: "infobar", role: "status" },
      icon("info"),
      h("span", { class: "msg" }, `RaabeWrite zeigt einige Inhalte dieses Dokuments noch nicht vollständig an: ${d.partial.join(", ")}. ` +
        "Sie bleiben beim Speichern unverändert erhalten."),
      close));
  }

  async openPath(path: string): Promise<void> {
    const open = this.docs.find((d) => d.path && samePath(d.path, path));
    if (open) { this.activate(open); return; }
    let bytes: Uint8Array;
    try {
      bytes = await this.platform.readFile(path);
    } catch (e) {
      await messageBox("RaabeWrite", `Die Datei „${baseName(path)}“ konnte nicht geöffnet werden.\n\n${String(e)}`);
      return;
    }
    let result;
    try {
      result = readDocx(bytes);
    } catch {
      await messageBox("RaabeWrite", `„${baseName(path)}“ ist kein Word-Dokument (.docx), das RaabeWrite lesen kann.`);
      return;
    }
    const pristine = this.docs.length === 1 && this.isPristine(this.docs[0]) ? this.docs[0] : null;
    this.createDoc(stripExt(baseName(path)), path, result.state, result.pkg, result.partial);
    if (pristine) this.removeDoc(pristine);
    addRecent(path, baseName(path));
    this.backstage.close();
  }

  private async openDialog(): Promise<void> {
    const path = await this.platform.openFileDialog();
    if (path) await this.openPath(path);
  }

  private suggestedName(d: Doc): string {
    if (d.path) return baseName(d.path);
    const first = d.editor.paragraphs().find((p) => p.text.trim());
    return `${(first && fileSafe(first.text)) || d.name}.docx`;
  }

  async save(d: Doc): Promise<boolean> {
    return d.path ? this.writeTo(d, d.path) : this.saveAs(d);
  }

  async saveAs(d: Doc): Promise<boolean> {
    let path = await this.platform.saveFileDialog(this.suggestedName(d), "docx");
    if (!path) return false;
    if (!/\.docx$/i.test(path)) path += ".docx";
    const other = this.docs.find((x) => x !== d && x.path && samePath(x.path, path!));
    if (other) {
      await messageBox("RaabeWrite", `„${baseName(path)}“ ist bereits in einem anderen Tab geöffnet. Schließen Sie es zuerst.`);
      return false;
    }
    return this.writeTo(d, path);
  }

  private async writeTo(d: Doc, path: string): Promise<boolean> {
    const stamp = d.editor.contentStamp();
    try {
      await this.platform.writeFile(path, writeDocx(d.editor.doc.state(), d.pkg, { author: this.settings.userName }));
    } catch (e) {
      await messageBox("RaabeWrite", `Das Dokument konnte nicht unter „${path}“ gespeichert werden. ` +
        `Vielleicht ist die Datei in einem anderen Programm geöffnet.\n\n${String(e)}`);
      return false;
    }
    d.path = path;
    d.name = stripExt(baseName(path));
    d.saved = stamp;
    d.lastSaved = Date.now();
    addRecent(path, baseName(path));
    void this.platform.deleteRecovery(d.id).catch(() => {});
    d.recoverySaved = null;
    this.renderInfobar();
    this.updateChrome(false);
    return true;
  }

  /** Asks to save a dirty document. Returns false when the user cancels. */
  private async confirmDiscard(d: Doc): Promise<boolean> {
    if (!this.isDirty(d)) return true;
    this.activate(d);
    const choice = await showDialog({
      title: "RaabeWrite",
      body: `Möchten Sie die Änderungen an „${d.name}“ speichern?`,
      buttons: [{ label: "Speichern", value: "save", primary: true }, { label: "Nicht speichern", value: "discard" }, { label: "Abbrechen", value: "cancel" }],
      cancel: "cancel",
    });
    if (choice === "cancel") return false;
    if (choice === "save") return this.save(d);
    return true;
  }

  async closeDoc(d: Doc): Promise<void> {
    if (!(await this.confirmDiscard(d))) return;
    void this.platform.deleteRecovery(d.id).catch(() => {});
    if (this.docs.length === 1) {
      // Closing the last document closes the window, like Word's X.
      this.platform.window.destroy();
      return;
    }
    this.removeDoc(d);
  }

  private removeDoc(d: Doc): void {
    const i = this.docs.indexOf(d);
    d.editor.view.stopBlink();
    d.ws.remove();
    this.docs.splice(i, 1);
    if (this.active === d) this.activate(this.docs[Math.min(i, this.docs.length - 1)]);
    else this.updateChrome(false);
  }

  private async confirmQuit(): Promise<boolean> {
    for (const d of [...this.docs]) if (!(await this.confirmDiscard(d))) return false;
    await Promise.all(this.docs.map((d) => this.platform.deleteRecovery(d.id).catch(() => {})));
    return true;
  }

  // ----- AutoRecover --------------------------------------------------------------------

  private async autoRecover(): Promise<void> {
    for (const d of this.docs) {
      const stamp = d.editor.contentStamp();
      if (!this.isDirty(d) || (d.recoverySaved && stamp.every((x, i) => x === d.recoverySaved![i]))) continue;
      try {
        await this.platform.saveRecovery({ id: d.id, title: d.name, path: d.path, savedAt: Date.now() },
          writeDocx(d.editor.doc.state(), d.pkg, { author: this.settings.userName }));
        d.recoverySaved = stamp;
      } catch {
        // Recovery is best effort; the next interval tries again.
      }
    }
  }

  private async offerRecovery(): Promise<void> {
    let entries;
    try { entries = await this.platform.listRecovery(); } catch { return; }
    if (!entries.length) return;
    const list = h("ul", { style: "margin:8px 0 0;padding-left:20px" },
      ...entries.map((e) => h("li", null, `${e.title} – ${new Date(e.savedAt).toLocaleString("de-DE")}`)));
    const choice = await showDialog({
      title: "Dokumentwiederherstellung",
      body: h("div", null, "RaabeWrite wurde nicht ordnungsgemäß beendet. Diese Dokumente wurden automatisch gesichert:", list),
      buttons: [{ label: "Wiederherstellen", value: "restore", primary: true }, { label: "Verwerfen", value: "discard" }],
      cancel: "later",
    });
    if (choice === "later") return;
    for (const e of entries) {
      if (choice === "restore") {
        try {
          const r = readDocx(await this.platform.readRecovery(e.id));
          const pristine = this.docs.length === 1 && this.isPristine(this.docs[0]) ? this.docs[0] : null;
          const d = this.createDoc(`${e.title} (Wiederhergestellt)`, e.path, r.state, r.pkg, r.partial);
          d.saved = [];
          if (pristine) this.removeDoc(pristine);
          this.updateChrome(true);
        } catch { /* unreadable recovery file: skip */ }
      }
      await this.platform.deleteRecovery(e.id).catch(() => {});
    }
  }

  // ----- printing and PDF ---------------------------------------------------------------

  private async print(pages?: number[]): Promise<void> {
    if (!this.active) return;
    this.backstage.close();
    this.active.editor.renderForPrint(this.printRoot, pages);
    await new Promise(requestAnimationFrame);
    print();
  }

  private async exportPdf(target?: string): Promise<void> {
    const d = this.active;
    if (!d) return;
    const path = target ?? await this.platform.saveFileDialog(this.suggestedName(d).replace(/\.docx$/i, ".pdf"), "pdf");
    if (!path) return;
    this.backstage.close();
    d.editor.renderForPrint(this.printRoot);
    await new Promise(requestAnimationFrame);
    try {
      await this.platform.printToPdf(/\.pdf$/i.test(path) ? path : `${path}.pdf`);
    } catch (e) {
      await messageBox("RaabeWrite", `Die PDF-Datei konnte nicht erstellt werden.\n\n${String(e)}`);
    } finally {
      this.printRoot.replaceChildren();
    }
  }

  // ----- zoom ---------------------------------------------------------------------------

  private setZoom(z: number): void {
    this.active?.editor.setZoom(z);
    this.updateChrome(false);
  }

  private async zoomDialog(): Promise<void> {
    const ed = this.active?.editor;
    if (!ed) return;
    const pct = h("input", { type: "number", min: "10", max: "500", value: String(Math.round(ed.zoom * 100)), style: "width:80px" });
    const preset = (label: string, z: () => number) =>
      h("label", { class: "check" }, h("input", { type: "radio", name: "zoom", onchange: () => { pct.value = String(Math.round(z() * 100)); } }), label);
    const body = h("div", null,
      h("div", { style: "display:grid;grid-template-columns:1fr 1fr;gap:0 24px" },
        preset("200 %", () => 2), preset("Seitenbreite", () => ed.view.pageWidthZoom()),
        preset("100 %", () => 1), preset("Ganze Seite", () => ed.view.onePageZoom()),
        preset("75 %", () => 0.75)),
      h("label", { class: "field" }, "Prozent:", pct));
    const ok = await showDialog({ title: "Zoom", body, buttons: [{ label: "OK", value: true, primary: true }, { label: "Abbrechen", value: false }], cancel: false });
    if (ok) this.setZoom(Math.max(10, Math.min(500, Number(pct.value) || 100)) / 100);
  }

  // ----- dialogs ------------------------------------------------------------------------

  private async wordCountDialog(): Promise<void> {
    const ed = this.active?.editor;
    if (!ed) return;
    const paras = ed.paragraphs();
    const text = paras.map((p) => p.text).join("");
    const rows: [string, number][] = [
      ["Seiten", ed.stats().pageCount],
      ["Wörter", ed.countWords()],
      ["Zeichen (ohne Leerzeichen)", text.replace(/\s/g, "").length],
      ["Zeichen (mit Leerzeichen)", text.length],
      ["Absätze", paras.filter((p) => p.text.trim()).length],
      ["Zeilen", ed.layout.placed.filter((pl) => paras[pl.para].text.trim()).length],
    ];
    const table = h("table", { class: "props", style: "border-spacing:0 4px" },
      h("tr", null, h("td", { style: "font-weight:600;padding-bottom:6px" }, "Statistik:")),
      ...rows.map(([k, v]) => h("tr", null, h("td", { style: "padding-right:40px" }, k), h("td", { style: "text-align:right" }, v.toLocaleString("de-DE")))));
    await showDialog({ title: "Wörter zählen", body: table, buttons: [{ label: "Schließen", value: undefined, primary: true }], cancel: undefined });
  }

  private async optionsDialog(page?: string): Promise<void> {
    const s = { ...this.settings };
    const name = h("input", { type: "text", value: s.userName, style: "width:280px" });
    const touch = h("input", { type: "checkbox", checked: s.touchMode });
    const seconds = h("input", { type: "number", min: "10", max: "3600", value: String(s.autoRecoverSeconds), style: "width:80px" });
    const general = h("div", { class: "panel" },
      h("label", { class: "field" }, "Benutzername (für Autor und Kommentare):", name),
      h("label", { class: "check" }, touch, "Touchmodus: größere Schaltflächen für Tablets und Touchscreens"));
    const saving = h("div", { class: "panel", hidden: true },
      h("label", { class: "field" }, "AutoWiederherstellen-Informationen speichern alle (Sekunden):", seconds),
      h("p", { class: "muted", style: "color:#5b5b5b" }, "Nicht gespeicherte Änderungen werden so oft gesichert und nach einem Absturz angeboten."));
    // Dokumentprüfung and Sprache come from the spelling feature (Word's pages).
    const extra = this.features.spell.optionPages();
    const panels = [general, saving, ...extra.map((x) => x.el)];
    for (const x of extra) x.el.hidden = true;
    const show = (panel: HTMLElement, btn: HTMLElement) => {
      for (const p of panels) p.hidden = p !== panel;
      nav.querySelectorAll("button").forEach((b) => b.classList.toggle("active", b === btn));
    };
    const nav = h("nav", null);
    for (const [label, panel] of [["Allgemein", general], ["Speichern", saving], ...extra.map((x) => [x.label, x.el] as const)] as const) {
      const b = h("button", { onclick: () => show(panel, b) }, label);
      nav.append(b);
    }
    const startAt = [...nav.querySelectorAll("button")].find((b) => b.textContent === page) ?? nav.firstElementChild as HTMLElement;
    show(panels[[...nav.children].indexOf(startAt)] ?? general, startAt);
    const ok = await showDialog({
      title: "RaabeWrite-Optionen", body: h("div", { class: "options" }, nav, ...panels),
      buttons: [{ label: "OK", value: true, primary: true }, { label: "Abbrechen", value: false }], cancel: false,
    });
    if (!ok) return;
    for (const x of extra) x.apply?.();
    this.settings = { userName: name.value.trim(), touchMode: touch.checked, autoRecoverSeconds: Math.max(10, Number(seconds.value) || 30) };
    saveSettings(this.settings);
    document.body.classList.toggle("touch", this.settings.touchMode);
  }

  // ----- backstage ----------------------------------------------------------------------

  private openBackstage(page: string): void {
    this.backstage.open(page);
  }

  private recentTable(): HTMLElement {
    const list = loadRecent().sort((a, b) => Number(b.pinned) - Number(a.pinned) || b.opened - a.opened);
    if (!list.length) return h("p", { class: "muted" }, "Sie haben in letzter Zeit keine Dokumente geöffnet.");
    const table = h("table", { class: "filelist" }, h("tr", null, h("th", { style: "width:40px" }), h("th", null, "Name"), h("th", { style: "width:180px" }, "Geöffnet"), h("th", { style: "width:40px" })));
    for (const r of list) {
      const nameCell = h("td", null, h("div", { class: "fname" }, stripExt(r.name)), h("div", { class: "fpath" }, dirName(r.path)));
      const pin = h("button", {
        class: `pinbtn${r.pinned ? " on" : ""}`, "data-tip": r.pinned ? "Nicht mehr anheften" : "Diese Datei anheften",
        onclick: (e: MouseEvent) => { e.stopPropagation(); setPinned(r.path, !r.pinned); this.backstage.refreshPage(); },
      }, icon(r.pinned ? "unpin" : "pin"));
      const row = h("tr", { class: "row", onclick: () => void this.openRecent(r.path) },
        h("td", null, h("img", { src: appIconUrl, width: 28, height: 28, alt: "" })), nameCell,
        h("td", { class: "muted" }, new Date(r.opened).toLocaleDateString("de-DE", { day: "numeric", month: "long", year: "numeric" })), h("td", null, pin));
      table.append(row);
      void this.platform.fileInfo(r.path).then((info) => { if (!info.exists) nameCell.classList.add("missing"); }).catch(() => {});
    }
    return table;
  }

  private async openRecent(path: string): Promise<void> {
    const info = await this.platform.fileInfo(path).catch(() => ({ exists: true, modified: null }));
    if (!info.exists) {
      const remove = await showDialog({
        title: "RaabeWrite", body: `„${baseName(path)}“ wurde nicht gefunden. Möglicherweise wurde die Datei verschoben, umbenannt oder gelöscht.`,
        buttons: [{ label: "Aus der Liste entfernen", value: true }, { label: "OK", value: false, primary: true }], cancel: false,
      });
      if (remove) { removeRecent(path); this.backstage.refreshPage(); }
      return;
    }
    await this.openPath(path);
  }

  private templates(): HTMLElement {
    const blank = h("button", { class: "template", onclick: () => { this.backstage.close(); this.newDoc(); } },
      h("div", { class: "thumb" }), "Leeres Dokument");
    const school = TEMPLATES.map((t) => h("button", {
      class: "template", "data-tip": `${t.name} (Raabeschule)\n${t.desc}`, onclick: () => this.newFromTemplate(t.id),
    }, templateThumb(t.id), t.name));
    return h("div", { class: "templates" }, blank, ...school);
  }

  /** A new document from a Raabeschule template; it counts as unchanged until edited. */
  newFromTemplate(t: TemplateId): void {
    this.backstage.close();
    const doc = this.newDoc(templateState(t));
    finishTemplate(t, doc.editor);
    doc.saved = doc.editor.contentStamp();
    this.lastTabsKey = "";
    this.updateChrome(false);
  }

  private backstageEntries(): BackstageEntry[] {
    const greeting = () => {
      const hr = new Date().getHours();
      return hr < 11 ? "Guten Morgen" : hr < 18 ? "Guten Tag" : "Guten Abend";
    };
    return [
      { id: "home", label: "Startseite", icon: "home", page: () => h("div", null,
        h("h1", null, greeting()), h("h2", null, "Neu"), this.templates(), h("h2", null, "Zuletzt verwendet"), this.recentTable()) },
      { id: "new", label: "Neu", icon: "new", page: () => h("div", null, h("h1", null, "Neu"), this.templates()) },
      { id: "open", label: "Öffnen", icon: "open", page: () => h("div", null,
        h("h1", null, "Öffnen"),
        h("button", { class: "bigbtn", onclick: () => void this.openDialog() }, icon("open", "l"), "Durchsuchen"),
        h("h2", null, "Zuletzt verwendet"), this.recentTable()) },
      { id: "info", label: "Informationen", icon: "info", divider: true, page: () => this.infoPage() },
      { id: "save", label: "Speichern", icon: "save", cmd: "file.save" },
      { id: "saveAs", label: "Speichern unter", icon: "save-as", cmd: "file.saveAs" },
      { id: "print", label: "Drucken", icon: "print", page: () => this.printPage() },
      { id: "export", label: "Exportieren", icon: "export", page: () => h("div", null,
        h("h1", null, "Exportieren"),
        h("button", { class: "bigbtn", style: "width:200px", onclick: () => void this.exportPdf() }, icon("export", "l"), "PDF-Dokument erstellen"),
        h("p", { class: "muted" }, "Speichert das Dokument als PDF, z. B. zum Abgeben oder Verschicken. Die PDF sieht auf jedem Gerät gleich aus.")) },
      { id: "close", label: "Schließen", icon: "close-doc", cmd: "file.close" },
      { id: "options", label: "Optionen", icon: "settings", cmd: "file.options", bottom: true },
    ];
  }

  private infoPage(): HTMLElement {
    const d = this.active!;
    const ed = d.editor;
    const props = h("table", { class: "props" },
      h("tr", null, h("td", null, "Seiten"), h("td", null, String(ed.stats().pageCount))),
      h("tr", null, h("td", null, "Wörter"), h("td", null, ed.countWords().toLocaleString("de-DE"))),
      h("tr", null, h("td", null, "Speicherort"), h("td", null, d.path ?? "Noch nicht gespeichert")),
      h("tr", null, h("td", null, "Zuletzt gespeichert"), h("td", null, d.lastSaved ? new Date(d.lastSaved).toLocaleString("de-DE") : "–")),
      h("tr", null, h("td", null, "Autor"), h("td", null, this.settings.userName || "– (unter Optionen festlegen)")));
    return h("div", null, h("h1", null, "Informationen"), h("h2", { style: "margin-top:0" }, d.name), props,
      d.partial.length ? h("p", { class: "warn" }, `Noch nicht vollständig angezeigt (bleibt beim Speichern erhalten): ${d.partial.join(", ")}.`) : null);
  }

  /** Datei > Drucken: settings on the left, the page preview with page navigation on the right. */
  private printPage(): HTMLElement {
    const ed = this.active?.editor;
    const count = ed?.stats().pageCount ?? 0;
    let page = Math.min(ed?.stats().caretPage ?? 0, Math.max(0, count - 1));
    const preview = h("div", { class: "print-preview" });
    const pageNo = h("input", { type: "text", value: String(page + 1), style: "width:40px;text-align:center" });
    const showPage = () => {
      preview.replaceChildren();
      if (!ed || !count) return;
      const holder = h("div");
      ed.renderForPrint(holder, [page]);
      const sheet = holder.querySelector<HTMLElement>(".print-page");
      if (!sheet) return;
      const setup = ed.layout.pages[page].setup;
      const w = setup.width * PT, hh = setup.height * PT;
      const scale = Math.min(0.75, (innerHeight - 260) / hh, 560 / w);
      sheet.classList.add("sheet");
      sheet.style.transform = `scale(${scale})`;
      sheet.style.transformOrigin = "top left";
      preview.append(h("div", { style: `width:${w * scale}px;height:${hh * scale}px;overflow:hidden` }, sheet));
      pageNo.value = String(page + 1);
    };
    const nav = (d: number) => { page = Math.max(0, Math.min(count - 1, page + d)); showPage(); };
    pageNo.addEventListener("change", () => { const n = Number(pageNo.value); if (Number.isFinite(n)) { page = Math.max(0, Math.min(count - 1, n - 1)); showPage(); } });
    const which = h("select", { style: "width:100%" },
      h("option", { value: "all" }, "Alle Seiten drucken"), h("option", { value: "current" }, "Aktuelle Seite drucken"), h("option", { value: "custom" }, "Benutzerdefinierter Druck"));
    const range = h("input", { type: "text", placeholder: "z. B. 1-3, 5", style: "width:100%" });
    range.addEventListener("input", () => { which.value = "custom"; });
    const pagesToPrint = (): number[] | undefined => {
      if (which.value === "current") return [page];
      if (which.value !== "custom" || !range.value.trim()) return undefined;
      const out = new Set<number>();
      for (const part of range.value.split(/[,;]/)) {
        const m = /^\s*(\d+)\s*(?:-\s*(\d+))?\s*$/.exec(part);
        if (!m) continue;
        const a = Number(m[1]), b = m[2] ? Number(m[2]) : a;
        for (let i = Math.min(a, b); i <= Math.max(a, b); i++) if (i >= 1 && i <= count) out.add(i - 1);
      }
      return out.size ? [...out].sort((x, y) => x - y) : undefined;
    };
    const setup = ed?.layout.pages[page]?.setup;
    const fmt = setup ? `${(Math.min(setup.width, setup.height) / 28.3465).toLocaleString("de-DE", { maximumFractionDigits: 1 })} x ${(Math.max(setup.width, setup.height) / 28.3465).toLocaleString("de-DE", { maximumFractionDigits: 1 })} cm` : "";
    showPage();
    return h("div", null, h("h1", null, "Drucken"),
      h("div", { class: "print-layout" },
        h("div", { style: "display:flex;flex-direction:column;gap:12px;width:260px" },
          h("button", { class: "bigbtn", onclick: () => void this.print(pagesToPrint()) }, icon("print", "l"), "Drucken"),
          h("div", { class: "muted" }, "Drucker und Exemplare wählen Sie im nächsten Fenster."),
          h("h3", { style: "margin:8px 0 0" }, "Einstellungen"), which, h("label", { class: "field" }, "Seiten:", range),
          h("div", { class: "muted" }, setup ? `${setup.width > setup.height ? "Querformat" : "Hochformat"}, ${fmt}` : ""),
          h("div", { class: "muted" }, `${count} ${count === 1 ? "Seite" : "Seiten"}`)),
        h("div", { style: "flex:1;display:flex;flex-direction:column;gap:8px;align-items:center" }, preview,
          h("div", { class: "print-nav" },
            h("button", { "aria-label": "Vorherige Seite", onclick: () => nav(-1) }, icon("back")), pageNo, h("span", null, `von ${count}`),
            h("button", { "aria-label": "Nächste Seite", onclick: () => nav(1), style: "transform:scaleX(-1)" }, icon("back"))))));
  }

  // ----- commands -----------------------------------------------------------------------

  private implementCommands(): void {
    const c = this.commands;
    const ed = () => this.active?.editor;
    const has = () => !!this.active;
    const run = (id: string, fn: (e: Editor) => void, checked?: (e: Editor) => boolean) =>
      c.implement(id, { run: () => { const e = ed(); if (e) fn(e); }, enabled: has, checked: checked ? () => { const e = ed(); return !!e && checked(e); } : undefined });

    c.implement("file.new", { run: () => { this.backstage.close(); this.newDoc(); } });
    c.implement("file.open", { run: () => this.openBackstage("open") });
    c.implement("file.save", { run: async () => { if (this.active) { this.backstage.close(); await this.save(this.active); } }, enabled: has });
    c.implement("file.saveAs", { run: async () => { if (this.active) { this.backstage.close(); await this.saveAs(this.active); } }, enabled: has });
    c.implement("file.print", { run: () => this.openBackstage("print"), enabled: has });
    c.implement("file.exportPdf", { run: () => this.exportPdf(), enabled: has });
    c.implement("file.close", { run: async () => { if (this.active) { this.backstage.close(); await this.closeDoc(this.active); } }, enabled: has });
    c.implement("file.options", { run: () => this.optionsDialog() });
    c.implement("review.languageOptions", { run: () => this.optionsDialog("Sprache") });
    c.implement("file.info", { run: () => this.openBackstage("info") });

    // Clipboard. The text goes to the system clipboard; the formatted copy stays inside RaabeWrite.
    c.implement("edit.undo", { run: () => ed()?.undo(), enabled: () => !!ed()?.canUndo });
    c.implement("edit.redo", { run: () => ed()?.redo(), enabled: () => !!ed()?.canRedo });
    // Text and HTML (formatting for Word); the system clipboard API first, the platform's text as fallback.
    const write = async (e: Editor, cut: boolean) => {
      const html = e.selectionHtml();
      const text = (cut ? e.cutSelection() : e.copySelection()).replace(/\n/g, "\r\n");
      if (this.platform.clipboardWriteHtml) {
        try { await this.platform.clipboardWriteHtml(html, text); return; } catch { /* the browser's API below */ }
      }
      try {
        await navigator.clipboard.write([new ClipboardItem({ "text/plain": new Blob([text], { type: "text/plain" }), "text/html": new Blob([html], { type: "text/html" }) })]);
      } catch {
        await this.platform.clipboardWrite(text);
      }
    };
    c.implement("edit.copy", { run: async () => { const e = ed(); if (e) await write(e, false); }, enabled: () => !!ed()?.hasSelection });
    c.implement("edit.cut", { run: async () => { const e = ed(); if (e) await write(e, true); }, enabled: () => !!ed()?.hasSelection });
    const paste = (plain: boolean) => async () => {
      const { text: t, html, image } = await this.readClipboard(plain);
      if (plain) { if (t) ed()?.pasteText(t, true); } else ed()?.paste(t, html, image);
    };
    c.implement("edit.paste", { run: paste(false), enabled: has });
    c.implement("edit.pasteText", { run: paste(true), enabled: has });
    c.implement("edit.selectAll", { run: () => ed()?.selectAll(), enabled: has });
    c.implement("edit.selectMenu", { run: () => {}, enabled: has });
    c.implement("edit.selectObjects", { run: () => { const e = ed(); if (e) e.setObjectSelect(!e.objectSelect); this.focusDoc(); }, enabled: has, checked: () => !!ed()?.objectSelect });
    c.implement("edit.selectSimilar", { run: () => {}, enabled: () => false });
    c.implement("edit.formatPainter", {
      run: () => {
        const e = ed();
        if (!e) return;
        const now = Date.now();
        const double = now - this.lastPainterClick < 450;
        this.lastPainterClick = now;
        if (e.painterActive && !double) e.stopPainter();
        else e.startPainter(double);
      },
      enabled: has, checked: () => !!ed()?.painterActive,
    });

    // Character formatting.
    const st = () => ed()?.charState();
    run("font.bold", (e) => e.toggleChar("bold"), () => !!st()?.bold);
    run("font.italic", (e) => e.toggleChar("italic"), () => !!st()?.italic);
    run("font.underline", (e) => e.setUnderline(), () => !!st()?.underline && st()?.underline !== "none");
    for (const kind of ["single", "double", "thick", "dotted", "dash", "wave"]) {
      run(`underline.${kind}`, (e) => e.setUnderline(kind), () => st()?.underline === kind);
    }
    run("font.strike", (e) => e.toggleChar("strike"), () => !!st()?.strike);
    run("font.allCaps", (e) => e.toggleChar("caps"), () => !!st()?.caps);
    run("font.subscript", (e) => e.toggleVertAlign("subscript"), () => !!st()?.subscript);
    run("font.superscript", (e) => e.toggleVertAlign("superscript"), () => !!st()?.superscript);
    run("font.grow", (e) => e.growFont(1));
    run("font.shrink", (e) => e.growFont(-1));
    run("font.clear", (e) => e.clearAllFormatting());
    run("font.color", (e) => { e.setColor(this.lastFontColor); });
    run("font.highlight", (e) => { e.setHighlight(this.lastHighlight); });
    run("font.name", () => this.focusRibbonInput("font.name"));
    run("font.size", () => this.focusRibbonInput("font.size"));
    run("font.dialog", () => void this.fontDialog());
    for (const m of ["sentence", "lower", "upper", "title", "toggle"] as const) run(`case.${m}`, (e) => e.changeCase(m));

    // Paragraph formatting.
    const ps = () => ed()?.paraState();
    run("para.alignLeft", (e) => e.setAlign("left"), () => ps()?.align === "left");
    run("para.alignCenter", (e) => e.setAlign("center"), () => ps()?.align === "center");
    run("para.alignRight", (e) => e.setAlign("right"), () => ps()?.align === "right");
    run("para.justify", (e) => e.setAlign("both"), () => ps()?.align === "both");
    run("para.indentInc", (e) => e.indent(1));
    run("para.indentDec", (e) => e.indent(-1));
    run("para.showMarks", (e) => { const on = !e.showMarks; for (const d of this.docs) d.editor.setShowMarks(on); }, (e) => e.showMarks);
    run("para.shading", (e) => e.paragraphShading(this.lastShading));
    run("para.borders", (e) => (e.inTable() ? void this.commands.run("border.all") : e.paragraphBorders("bottom")));
    run("para.dialog", () => void this.paragraphDialog());
    const line = (rule: LineRule, v: number) => (e: Editor) => e.setLineSpacing(rule, v);
    const isLine = (v: number) => (e: Editor) => { const r = e.paraState(); return r.lineRule === "auto" && Math.abs(r.line - v) < 0.005; };
    run("spacing.1", line("auto", 1), isLine(1));
    run("spacing.1.15", line("auto", 1.15), isLine(1.15));
    run("spacing.1.5", line("auto", 1.5), isLine(1.5));
    run("spacing.2", line("auto", 2), isLine(2));
    run("spacing.2.5", line("auto", 2.5), isLine(2.5));
    run("spacing.3", line("auto", 3), isLine(3));
    // Dropdown buttons whose menu does the work.
    c.implement("font.case", { run: () => {}, enabled: has });
    c.implement("para.lineSpacing", { run: () => {}, enabled: has });
    run("spacing.toggleBefore", (e) => e.setSpacing(e.paraState().before > 0 ? 0 : 12));
    run("spacing.toggleAfter", (e) => e.setSpacing(undefined, e.paraState().after > 0 ? 0 : 12));
    c.onRefresh(() => {
      const r = ps();
      c.get("spacing.toggleBefore").label = r && r.before > 0 ? "Abstand vor Absatz entfernen" : "Abstand vor Absatz hinzufügen";
      c.get("spacing.toggleAfter").label = r && r.after > 0 ? "Abstand nach Absatz entfernen" : "Abstand nach Absatz hinzufügen";
    });

    // Styles: paragraph styles replace the paragraph's style, character styles apply to the selection.
    for (const g of GALLERY) {
      run(g.cmd, (e) => {
        const id = e.styleId(g.name);
        const type = e.doc.styles.get(id)?.type ?? (/(emphasis|strong|reference|book title)/i.test(g.name) ? "character" : "paragraph");
        if (type === "character") e.applyCharStyle(id); else e.applyParaStyle(id);
      }, (e) => {
        const id = e.styleId(g.name);
        return e.doc.styles.get(id)?.type === "character" ? e.charState().charStyle === id : e.currentStyle() === id;
      });
    }
    run("style.h3", (e) => e.applyParaStyle(e.styleId("heading 3")));

    // Layout.
    const layoutOn = (id: string) => c.implement(id, { run: () => this.focusRibbonInput(id), enabled: has });
    for (const id of ["layout.indentLeft", "layout.indentRight", "layout.spaceBefore", "layout.spaceAfter"]) layoutOn(id);
    run("insert.pageBreak", (e) => e.insertBreak("page"));
    run("insert.lineBreak", (e) => e.insertBreak("line"));
    run("insert.blankPage", (e) => { e.insertBreak("page"); e.insertBreak("page"); });
    c.implement("layout.breaks", { run: () => {}, enabled: has });

    // View.
    c.implement("view.zoom100", { run: () => this.setZoom(1), enabled: has });
    c.implement("view.gridlines", { run: () => { const e = ed(); if (e) e.view.setDrawingGrid(!e.view.drawingGrid); }, enabled: has, checked: () => !!ed()?.view.drawingGrid });
    c.implement("view.onePage", { run: () => { const e = ed(); if (e) this.setZoom(e.view.onePageZoom()); }, enabled: has });
    c.implement("view.pageWidth", { run: () => { const e = ed(); if (e) this.setZoom(e.view.pageWidthZoom()); }, enabled: has });
    c.implement("view.zoomDialog", { run: () => this.zoomDialog(), enabled: has });
    c.implement("view.ribbon", { run: () => this.ribbon.toggleMinimized(), checked: () => this.ribbon.isMinimized });
    c.implement("view.touchMode", {
      run: () => { this.settings.touchMode = !this.settings.touchMode; saveSettings(this.settings); document.body.classList.toggle("touch", this.settings.touchMode); },
      checked: () => this.settings.touchMode,
    });
    c.implement("review.wordCount", { run: () => this.wordCountDialog(), enabled: has });
  }

  /** Puts the keyboard focus into a ribbon box (Strg+Umschalt+A for the font, Layout fields …). */
  private focusRibbonInput(cmd: string): void {
    const tab = cmd.startsWith("layout.") ? "layout" : "start";
    if (this.ribbon.activeTab !== tab) this.ribbon.select(tab);
    requestAnimationFrame(() => {
      const el = this.ribbon.card.querySelector<HTMLInputElement>(`[data-tip^="${this.commands.get(cmd).label}"] input`);
      el?.focus();
    });
  }

  // ----- ribbon values --------------------------------------------------------------------

  private ribbonContext(): RibbonContext {
    const ed = () => this.active?.editor;
    // eslint-disable-next-line @typescript-eslint/no-this-alias
    const self = this;
    const de = (n: number, digits = 2) => n.toLocaleString("de-DE", { maximumFractionDigits: digits });
    const parseNum = (v: string) => Number(v.replace(/[^\d,.-]/g, "").replace(",", "."));
    const theme = () => ed()?.doc.styles.theme;
    return {
      citationStyle: {
        value: () => this.features.citations.style.value(),
        options: () => this.features.citations.style.options(),
        commit: (v: string) => this.features.citations.style.commit(v),
      },
      styleHover: (cmd) => {
        const e = ed();
        if (!e) return;
        const g = cmd ? GALLERY.find((x) => x.cmd === cmd) : undefined;
        if (!g) { e.previewStyle(null); return; }
        const id = e.styleId(g.name);
        const type = e.doc.styles.get(id)?.type ?? (/(emphasis|strong|reference|book title)/i.test(g.name) ? "character" : "paragraph");
        e.previewStyle(id, type === "character" ? "character" : "paragraph");
      },
      fontName: () => {
        // In an equation Word shows its font.
        if (this.features?.math.editing()) return "Cambria Math";
        const s = ed()?.charState();
        if (!s?.font) return "";
        const t = theme();
        if (t && s.font === t.minorFont) return `${s.font} (Textkörper)`;
        if (t && s.font === t.majorFont) return `${s.font} (Überschriften)`;
        return s.font;
      },
      fontSize: () => { const s = ed()?.charState().size; return s === undefined ? "" : de(s, 1); },
      fontOptions: () => {
        const t = theme();
        const groups: { heading?: string; items: string[] }[] = [];
        if (t) groups.push({ heading: "Designschriftarten", items: [`${t.majorFont} (Überschriften)`, `${t.minorFont} (Textkörper)`] });
        if (this.recentFonts.length) groups.push({ heading: "Zuletzt verwendete Schriftarten", items: this.recentFonts });
        groups.push({ heading: "Alle Schriftarten", items: this.fontList.length ? this.fontList : ["Arial", "Calibri", "Cambria", "Times New Roman"] });
        return groups;
      },
      setFont: (name) => {
        const clean = name.replace(/ \((Textkörper|Überschriften)\)$/, "");
        ed()?.setFont(clean);
        this.recentFonts = [clean, ...this.recentFonts.filter((f) => f !== clean)].slice(0, 5);
      },
      setSize: (v) => { const n = parseNum(v); if (n > 0) ed()?.setSize(n); },
      fontColor: () => this.lastFontColor,
      setFontColor: (hex) => { if (hex !== "auto") this.lastFontColor = hex; ed()?.setColor(hex); },
      highlight: () => this.lastHighlight,
      setHighlight: (name) => { if (name !== "none") this.lastHighlight = name; ed()?.setHighlight(name); },
      indent: (which) => {
        const r = ed()?.paraState();
        if (!r) return "";
        return `${de((which === "left" ? r.indLeft : r.indRight) / 28.3465, 2)} cm`;
      },
      setIndent: (which, v) => {
        const cm = parseNum(v);
        if (!Number.isFinite(cm)) return;
        ed()?.setParaProps(which === "left" ? { indLeft: cm * 28.3465 } : { indRight: cm * 28.3465 });
      },
      stepIndent: (which, dir) => {
        const r = ed()?.paraState();
        if (!r) return;
        const cur = (which === "left" ? r.indLeft : r.indRight) / 28.3465;
        const next = Math.round((cur + dir * 0.1) * 10) / 10;
        ed()?.setParaProps(which === "left" ? { indLeft: next * 28.3465 } : { indRight: next * 28.3465 });
      },
      spacing: (which) => {
        const r = ed()?.paraState();
        return r ? `${de(which === "before" ? r.before : r.after, 1)} Pt.` : "";
      },
      setSpacing: (which, v) => {
        const n = parseNum(v);
        if (!Number.isFinite(n) || n < 0) return;
        ed()?.setSpacing(which === "before" ? n : undefined, which === "after" ? n : undefined);
      },
      shading: () => this.lastShading,
      setShading: (hex) => { if (hex !== "auto") this.lastShading = hex; ed()?.paragraphShading(hex); },
      panels: {
        bullets: (close) => this.features.lists.bullets(close),
        numbers: (close) => this.features.lists.numbers(close),
        multilevel: (close) => this.features.lists.multilevel(close),
        insertTable: (close) => this.features.tables.insertPanel(close),
        margins: (close) => this.features.pages.marginsPanel(close),
        size: (close) => this.features.pages.sizePanel(close),
        get columns() { return self.features.pages.columnsMenu; },
        get breaks() { return self.features.pages.breaksMenu; },
        get lineNumbers() { return self.features.pages.lineNumbersMenu; },
        get header() { return self.features.hf.headerMenu; },
        get footer() { return self.features.hf.footerMenu; },
        get pageNumber() { return self.features.hf.pageNumberMenu; },
        get pictures() { return self.features.pictures.insertMenu; },
        shapes: (close) => this.features.shapes.galleryPanel(close),
        position: (close) => this.features.pictures.arrange.positionPanel(close),
        align: (close) => this.features.pictures.arrange.alignPanel(close),
        get rotate() { return self.features.pictures.arrange.rotateMenu; },
        forward: [{ cmd: "picture.bringForward" }, { cmd: "picture.bringToFront" }, { cmd: "picture.bringInFrontOfText" }],
        backward: [{ cmd: "picture.sendBackward" }, { cmd: "picture.sendToBack" }, { cmd: "picture.sendBehindText" }],
        get textBox() { return self.features.shapes.textBoxMenu; },
        get toc() { return self.features.refs.tocMenu; },
        equation: (close) => this.features.math.insertPanel(close),
        symbol: (close) => this.features.symbols.menuPanel(close),
        get citation() { return self.features.citations.citationMenu(); },
        bibliography: (close) => this.features.citations.bibliographyPanel(close),
        get addText() { return self.features.refs.addTextMenu; },
        get nextNote() { return self.features.refs.nextNoteMenu; },
        get track() { return self.features.review.trackMenu; },
        get deleteComment() { return self.features.review.deleteMenu; },
        get accept() { return self.features.review.acceptMenu; },
        get reject() { return self.features.review.rejectMenu; },
        get showComments() { return self.features.review.showCommentsMenu; },
        get filterMarkup() { return self.features.review.filterMenu; },
        get reviewingPane() { return self.features.review.reviewingPaneMenu; },
        borders: [...["bottom", "top", "left", "right", "none", "all", "outside", "inside"].map((k) => ({ cmd: `border.${k}` })), "-", { cmd: "border.dialog" }],
        themes: (close) => this.features.design.themesPanel(close),
        styleSet: (close) => this.features.design.styleSetPanel(close),
        colors: (close) => this.features.design.colorsPanel(close),
        fonts: (close) => this.features.design.fontsPanel(close),
        effects: (close) => this.features.design.effectsPanel(close),
        get paragraphSpacing() { return self.features.design.spacingMenu; },
        watermark: (close) => this.features.background.watermarkPanel(close),
        pageColor: (close) => this.features.background.pageColorPanel(close),
        coverPage: (close) => this.features.cover.panel(close),
        get fontEffects() { return self.features.textEffects.fontEffectsMenu; },
        wrap: ["picture.wrapInline", "picture.wrapSquare", "picture.wrapTight", "picture.wrapThrough", "picture.wrapTopBottom", "picture.wrapBehind", "picture.wrapFront"].map((cmd) => ({ cmd })),
      },
      markup: {
        value: () => self.features.review.markup.value(),
        set: (v) => self.features.review.markup.set(v),
        options: () => self.features.review.markup.options(),
      },
      stepSpacing: (which, dir) => {
        const r = ed()?.paraState();
        if (!r) return;
        const cur = which === "before" ? r.before : r.after;
        // Word steps in 6 pt.
        const next = Math.max(0, dir > 0 ? Math.floor(cur / 6 + 1e-6) * 6 + 6 : Math.ceil(cur / 6 - 1e-6) * 6 - 6);
        ed()?.setSpacing(which === "before" ? next : undefined, which === "after" ? next : undefined);
      },
    };
  }

  // ----- Schriftart and Absatz dialogs ---------------------------------------------------

  /** Schriftart (Strg+D); with a target it edits a style's formatting (Formatvorlage ändern > Format > Zeichen). */
  private async fontDialog(target?: FontTarget): Promise<void> {
    const e = this.active?.editor;
    if (!e && !target) return;
    const s = target?.state ?? e!.charState();
    const fonts = this.fontList.length ? this.fontList : ["Arial", "Calibri", "Cambria", "Times New Roman"];
    const fontInput = h("input", { type: "text", value: s.font ?? "", style: "width:100%", list: "rw-fonts" });
    const datalist = h("datalist", { id: "rw-fonts" }, ...fonts.map((f) => h("option", { value: f })));
    const styleSel = h("select", null, ...["Standard", "Kursiv", "Fett", "Fett Kursiv"].map((x) => h("option", null, x)));
    styleSel.selectedIndex = (s.bold ? 2 : 0) + (s.italic ? 1 : 0);
    const sizeInput = h("input", { type: "text", value: s.size === undefined ? "" : String(s.size).replace(".", ","), style: "width:70px" });
    const underlineSel = h("select", null, ...[["none", "(ohne)"], ["single", "Einfach"], ["double", "Doppelt"], ["thick", "Dick"], ["dotted", "Gepunktet"], ["dash", "Gestrichelt"], ["wave", "Wellenlinie"]]
      .map(([v, l]) => h("option", { value: v }, l)));
    underlineSel.value = s.underline ?? "none";
    let color = s.color ?? "auto";
    const colorBtn = h("button", { style: "min-width:120px;height:28px;border:1px solid #c7c7c7;border-radius:4px;display:inline-flex;align-items:center;gap:6px;padding:0 8px" });
    // Like Word: a colour bar, with the word "Automatisch" only for the automatic colour.
    const paintColor = () => colorBtn.replaceChildren(
      h("span", { style: `width:${color === "auto" ? 28 : 84}px;height:12px;background:${color === "auto" ? "#000" : `#${color}`}` }),
      color === "auto" ? "Automatisch" : "", icon("chevron-down"));
    paintColor();
    colorBtn.onclick = () => showFontColors(colorBtn.getBoundingClientRect(), (c) => { color = c; paintColor(); update(); });
    const box = (label: string, on: boolean | undefined) => {
      const cb = h("input", { type: "checkbox", checked: !!on });
      return { cb, el: h("label", { class: "check", style: "margin:4px 0" }, cb, label) };
    };
    const fx = {
      strike: box("Durchgestrichen", s.strike), dstrike: box("Doppelt durchgestrichen", s.dstrike),
      superscript: box("Hochgestellt", s.superscript), subscript: box("Tiefgestellt", s.subscript),
      smallCaps: box("Kapitälchen", s.smallCaps), caps: box("Großbuchstaben", s.caps), hidden: box("Ausgeblendet", s.hidden),
    };
    const preview = h("div", { class: "preview" }, "Beispieltext");
    const update = () => {
      const ital = styleSel.selectedIndex % 2 === 1, bold = styleSel.selectedIndex >= 2;
      preview.style.cssText = `font-family:"${fontInput.value}";font-size:${Math.min(36, Number(sizeInput.value.replace(",", ".")) || 11) * 4 / 3}px;` +
        `font-weight:${bold ? 700 : 400};font-style:${ital ? "italic" : "normal"};color:${color === "auto" ? "#000" : `#${color}`};` +
        `text-decoration:${underlineSel.value !== "none" ? "underline" : ""} ${fx.strike.cb.checked || fx.dstrike.cb.checked ? "line-through" : ""};` +
        `font-variant:${fx.smallCaps.cb.checked ? "small-caps" : "normal"};text-transform:${fx.caps.cb.checked ? "uppercase" : "none"}`;
    };
    for (const el of [fontInput, sizeInput, styleSel, underlineSel, ...Object.values(fx).map((x) => x.cb)]) el.addEventListener("input", update);
    update();
    const body = h("div", { style: "width:520px" },
      h("div", { class: "grid2", style: "grid-template-columns:2fr 1fr 80px" },
        h("label", { class: "field" }, "Schriftart:", fontInput, datalist),
        h("label", { class: "field" }, "Schriftschnitt:", styleSel),
        h("label", { class: "field" }, "Schriftgrad:", sizeInput)),
      h("div", { style: "display:flex;gap:24px;margin:8px 0" },
        h("label", { class: "field" }, "Schriftfarbe:", colorBtn),
        h("label", { class: "field" }, "Unterstreichung:", underlineSel)),
      h("fieldset", null, h("legend", null, "Effekte"),
        h("div", { style: "display:grid;grid-template-columns:1fr 1fr;gap:0 24px" }, ...Object.values(fx).map((x) => x.el))),
      h("fieldset", null, h("legend", null, "Vorschau"), preview));
    const ok = await showDialog({ title: "Schriftart", body, buttons: [{ label: "OK", value: true, primary: true }, { label: "Abbrechen", value: false }], cancel: false });
    if (!ok) return;
    const changed = <T>(a: T | undefined, b: T) => a !== b;
    if (target) {
      const ital = styleSel.selectedIndex % 2 === 1, bold = styleSel.selectedIndex >= 2;
      const size = Number(sizeInput.value.replace(",", "."));
      const out: Record<string, unknown> = {};
      if (fontInput.value && changed(s.font, fontInput.value)) out.font = fontInput.value;
      if (size > 0 && changed(s.size, size)) out.size = size;
      if (changed(!!s.bold, bold)) out.bold = bold;
      if (changed(!!s.italic, ital)) out.italic = ital;
      if (changed(s.underline ?? "none", underlineSel.value)) out.underline = underlineSel.value;
      if (changed(s.color ?? "auto", color)) out.color = color;
      for (const k of ["strike", "dstrike", "smallCaps", "caps", "hidden"] as const) if (changed(!!s[k], fx[k].cb.checked)) out[k] = fx[k].cb.checked;
      const va = fx.superscript.cb.checked ? "superscript" : fx.subscript.cb.checked ? "subscript" : "baseline";
      if (changed(s.superscript ? "superscript" : s.subscript ? "subscript" : "baseline", va)) out.vertAlign = va;
      target.apply(out as CharProps);
      return;
    }
    if (!e) return;
    if (fontInput.value && changed(s.font, fontInput.value)) e.setFont(fontInput.value);
    const size = Number(sizeInput.value.replace(",", "."));
    if (size > 0 && changed(s.size, size)) e.setSize(size);
    const ital = styleSel.selectedIndex % 2 === 1, bold = styleSel.selectedIndex >= 2;
    if (changed(!!s.bold, bold)) e.toggleChar("bold");
    if (changed(!!s.italic, ital)) e.toggleChar("italic");
    if (changed(s.underline ?? "none", underlineSel.value)) e.setUnderline(underlineSel.value);
    if (changed(s.color ?? "auto", color)) e.setColor(color);
    for (const k of ["strike", "dstrike", "smallCaps", "caps", "hidden"] as const) if (changed(!!s[k], fx[k].cb.checked)) e.toggleChar(k);
    if (changed(!!s.superscript, fx.superscript.cb.checked)) e.toggleVertAlign("superscript");
    if (changed(!!s.subscript, fx.subscript.cb.checked)) e.toggleVertAlign("subscript");
  }

  /** Absatz; with a target it edits a style's formatting (Formatvorlage ändern > Format > Absatz). */
  private async paragraphDialog(target?: ParaTarget): Promise<void> {
    const e = this.active?.editor;
    if (!e && !target) return;
    const r = target?.state ?? e!.paraState();
    const cm = (pt: number) => (pt / 28.3465).toLocaleString("de-DE", { maximumFractionDigits: 2 }) + " cm";
    const pt = (v: number) => v.toLocaleString("de-DE", { maximumFractionDigits: 1 }) + " Pt.";
    const num = (v: string) => Number(v.replace(/[^\d,.-]/g, "").replace(",", "."));
    const input = (value: string) => h("input", { type: "text", value, style: "width:90px" });
    const select = (opts: [string, string][], value: string) => {
      const sel = h("select", null, ...opts.map(([v, l]) => h("option", { value: v }, l)));
      sel.value = value;
      return sel;
    };
    const align = select([["left", "Links"], ["center", "Zentriert"], ["right", "Rechts"], ["both", "Blocksatz"]], r.align);
    const left = input(cm(r.indLeft)), right = input(cm(r.indRight));
    const special = select([["none", "(ohne)"], ["first", "Erste Zeile"], ["hanging", "Hängend"]], r.indFirst > 0 ? "first" : r.indFirst < 0 ? "hanging" : "none");
    const by = input(r.indFirst ? cm(Math.abs(r.indFirst)) : "");
    const before = input(pt(r.before)), after = input(pt(r.after));
    const lineKind = r.lineRule === "exact" ? "exact" : r.lineRule === "atLeast" ? "atLeast" :
      Math.abs(r.line - 1) < 0.01 ? "single" : Math.abs(r.line - 1.5) < 0.01 ? "1.5" : Math.abs(r.line - 2) < 0.01 ? "double" : "multiple";
    const lineSel = select([["single", "Einfach"], ["1.5", "1,5 Zeilen"], ["double", "Doppelt"], ["atLeast", "Mindestens"], ["exact", "Genau"], ["multiple", "Mehrfach"]], lineKind);
    const lineVal = input(r.lineRule === "auto" ? r.line.toLocaleString("de-DE", { maximumFractionDigits: 2 }) : pt(r.line));
    const check = (label: string, on: boolean) => { const cb = h("input", { type: "checkbox", checked: on }); return { cb, el: h("label", { class: "check" }, cb, label) }; };
    const contextual = check("Keinen Abstand zwischen Absätzen gleicher Formatvorlage einfügen", r.contextual);
    const widow = check("Absatzkontrolle", r.widow), keepNext = check("Nicht vom nächsten Absatz trennen", r.keepNext);
    const keepLines = check("Diesen Absatz zusammenhalten", r.keepLines), pageBefore = check("Seitenumbruch oberhalb", r.pageBreakBefore);
    const noNumbers = check("Zeilennummern unterdrücken", r.suppressLineNumbers), noHyphens = check("Keine Silbentrennung", r.suppressAutoHyphens);
    const page1 = h("div", null,
      h("fieldset", null, h("legend", null, "Allgemein"), h("label", { class: "field", style: "flex-direction:row;align-items:center;gap:8px" }, "Ausrichtung:", align)),
      h("fieldset", null, h("legend", null, "Einzug"),
        h("div", { class: "grid2" }, "Links:", left, "Sondereinzug:", special, "Rechts:", right, "Um:", by)),
      h("fieldset", null, h("legend", null, "Abstand"),
        h("div", { class: "grid2" }, "Vor:", before, "Zeilenabstand:", lineSel, "Nach:", after, "Maß:", lineVal), contextual.el));
    const page2 = h("div", { hidden: true }, h("fieldset", null, h("legend", null, "Paginierung"), widow.el, keepNext.el, keepLines.el, pageBefore.el),
      h("fieldset", null, h("legend", null, "Formatierungsausnahmen"), noNumbers.el, noHyphens.el));
    const tabs = h("div", { class: "tabs" });
    const tab = (label: string, pg: HTMLElement) => h("button", {
      class: pg === page1 ? "active" : "",
      onclick: (ev: MouseEvent) => { page1.hidden = pg !== page1; page2.hidden = pg !== page2; tabs.querySelectorAll("button").forEach((b) => b.classList.toggle("active", b === ev.currentTarget)); },
    }, label);
    tabs.append(tab("Einzüge und Abstände", page1), tab("Zeilen- und Seitenumbruch", page2));
    const ok = await showDialog({ title: "Absatz", body: h("div", { style: "width:520px" }, tabs, page1, page2), buttons: [{ label: "OK", value: true, primary: true }, { label: "Abbrechen", value: false }], cancel: false });
    if (!ok) return;
    const firstAbs = num(by.value) * 28.3465 || 0;
    const kind = lineSel.value;
    const lineRule: LineRule = kind === "exact" ? "exact" : kind === "atLeast" ? "atLeast" : "auto";
    const lineValue = kind === "single" ? 1 : kind === "1.5" ? 1.5 : kind === "double" ? 2 : num(lineVal.value) || (lineRule === "auto" ? 1 : 12);
    (target ? target.apply : (pp: Partial<ParaProps>) => e!.setParaProps(pp))({
      align: align.value as "left" | "center" | "right" | "both",
      indLeft: num(left.value) * 28.3465 || 0,
      indRight: num(right.value) * 28.3465 || 0,
      indFirst: special.value === "first" ? firstAbs : special.value === "hanging" ? -firstAbs : 0,
      before: num(before.value) || 0,
      after: num(after.value) || 0,
      lineRule, line: lineValue,
      contextual: contextual.cb.checked, widow: widow.cb.checked, keepNext: keepNext.cb.checked,
      keepLines: keepLines.cb.checked, pageBreakBefore: pageBefore.cb.checked,
      suppressLineNumbers: noNumbers.cb.checked || undefined, suppressAutoHyphens: noHyphens.cb.checked || undefined,
    });
  }

  private installKeyboard(): void {
    addEventListener("keydown", (e) => {
      if (e.defaultPrevented || e.isComposing) return;
      // Esc ends the Lesemodus (Word).
      if (e.key === "Escape" && this.features.views.isReading() && !document.querySelector(".dialog-layer, .menu")) { e.preventDefault(); this.features.views.leave(); return; }
      const ctrl = e.ctrlKey, shift = e.shiftKey, alt = e.altKey;
      const k = e.key.length === 1 ? e.key.toLowerCase() : e.key;
      // Word binds some shortcuts to keys, not characters (German layout: # < + - 8).
      const vk = e.keyCode;
      let cmd: string | null = null;
      const ed = this.active?.editor;
      if (ctrl && !alt && !shift) {
        cmd = ({
          n: "file.new", o: "file.open", s: "file.save", p: "file.print", w: "file.close", l: "para.alignLeft", e: "para.alignCenter",
          f: "edit.find", h: "edit.replace", g: "edit.goto", k: "insert.link",
          r: "para.alignRight", b: "para.justify", m: "para.indentInc", d: "font.dialog", q: "", "1": "spacing.1", "2": "spacing.2",
          "5": "spacing.1.5", "0": "view.zoom100", "9": "", "8": "",
        } as Record<string, string>)[k] ?? null;
        if (k === "q") { e.preventDefault(); ed?.resetPara(); return; }
        if (k === "t") { e.preventDefault(); ed?.hanging(1); return; }
        if (k === "9" || k === "8") { e.preventDefault(); ed?.growFont(k === "9" ? 1 : -1, true); return; }
        if (k === " ") { e.preventDefault(); ed?.resetChar(); return; }
        if (vk === 191) cmd = "font.subscript";            // Strg+#
        if (vk === 226) cmd = "font.shrink";               // Strg+<
        if (vk === 187) { e.preventDefault(); this.setZoom(Math.min(5, (ed?.zoom ?? 1) + 0.1)); return; } // Strg++
        if (vk === 189) { e.preventDefault(); this.setZoom(Math.max(0.1, (ed?.zoom ?? 1) - 0.1)); return; } // Strg+-
        if (k === "F4") cmd = "file.close";
        if (k === "F12") { e.preventDefault(); void this.openDialog(); return; }
        if (k === "F1") cmd = "view.ribbon";
      }
      if (!ctrl && !shift && !alt && k === "F9") { e.preventDefault(); this.features.captions.updateFields(); return; }
      if (ctrl && shift && !alt) {
        cmd = ({
          f: "font.bold", k: "font.italic", u: "font.underline", d: "underline.double", g: "font.allCaps", n: "style.normal",
          m: "para.indentDec", v: "edit.pasteText", a: "font.name", p: "font.size",
        } as Record<string, string>)[k] ?? cmd;
        if (k === "e") cmd = "review.trackChanges";
        if (k === "q") { e.preventDefault(); ed?.toggleChar("smallCaps"); return; }
        if (k === "h") { e.preventDefault(); ed?.toggleChar("hidden"); return; }
        if (k === "w") { e.preventDefault(); ed?.setUnderline("words"); return; }
        if (k === "z") { e.preventDefault(); ed?.resetChar(); return; }
        if (k === "t") { e.preventDefault(); ed?.hanging(-1); return; }
        if (k === "l") { e.preventDefault(); ed?.applyListBullet(); return; }
        if (vk === 226) cmd = "font.grow";                 // Strg+> (Umschalt+<)
        if (vk === 189) cmd = "font.subscript";            // Strg+Umschalt+-
        if (vk === 56) cmd = "para.showMarks";             // Strg+* (Umschalt+8)
      }
      if (ctrl && alt && !shift) {
        if (k === "f" && vk === 70) cmd = "ref.footnote";
        if (k === "d" && vk === 68) cmd = "ref.endnote";
        if (k === "k" && vk === 75) cmd = "review.newComment";
        // Word: CopyFormat / PasteFormat (the painter brush is the ribbon button).
        if (k === "c" && vk === 67) { e.preventDefault(); if (ed) this.copiedFormat = ed.copyFormat(); return; }
        if (k === "v" && vk === 86) { e.preventDefault(); if (ed && this.copiedFormat) ed.pasteFormat(this.copiedFormat); return; }
      }
      if (!ctrl && !alt && k === "F12") cmd = "file.saveAs";
      if (alt && shift && !ctrl && vk === 48) cmd = "insert.equation";             // Alt+Umschalt+0
      if (!ctrl && !alt && shift && k === "F7") cmd = "review.thesaurus";
      if (!ctrl && !alt && shift && k === "F3") { e.preventDefault(); ed?.changeCase("cycle"); return; }
      if (ctrl && (k === "Tab" || k === "F6") && this.docs.length > 1 && this.active) {
        e.preventDefault();
        const i = this.docs.indexOf(this.active);
        this.activate(this.docs[(i + (shift ? this.docs.length - 1 : 1)) % this.docs.length]);
        return;
      }
      if (cmd && this.commands.has(cmd)) {
        e.preventDefault();
        void this.commands.run(cmd);
      }
    });
  }

  /** Clicking ribbon or toolbar buttons must not take the focus (and caret) away from the text. */
  private installFocusKeeping(): void {
    addEventListener("mousedown", (e) => {
      const t = e.target as Element;
      if (t.closest("input, select, textarea")) return;
      if (t.closest(".ribbon-card, .qat, .tabrow, .statusbar, .group-popup, .menu, .rpanel")) e.preventDefault();
    }, true);
  }

  /** The active editor, for the feel-lab probe. */
  activeEditor(): Editor {
    return this.active!.editor;
  }

  /** Test hooks for feel-lab runs only (no file dialogs; see tools/feel-lab/smoke_write.py). */
  testApi() {
    return {
      docs: () => this.docs.map((d) => ({ name: d.name, path: d.path, dirty: this.isDirty(d), unsupported: d.partial, active: d === this.active })),
      openPath: (path: string) => this.openPath(path),
      /** Opens a document from bytes (development platform, which has no file system). */
      openBytes: (name: string, bytes: Uint8Array) => {
        const result = readDocx(bytes);
        this.createDoc(stripExt(name), null, result.state, result.pkg, result.partial);
      },
      saveTo: (path: string) => (this.active ? this.writeTo(this.active, path) : Promise.resolve(false)),
      save: () => (this.active ? this.save(this.active) : Promise.resolve(false)),
      exportPdfTo: (path: string) => this.exportPdf(path),
      autoRecoverNow: () => this.autoRecover(),
      run: (id: string) => this.commands.run(id),
      enabled: (id: string) => this.commands.isEnabled(id),
      newDoc: () => { this.newDoc(); },
      newFromTemplate: (t: TemplateId) => this.newFromTemplate(t),
      /** Closes every document except the active one, discarding changes (catalogue runs). */
      closeOthers: () => { for (const d of [...this.docs]) if (d !== this.active) this.removeDoc(d); },
      /** Einfügen > Bilder > Dieses Gerät without the file dialog. */
      insertPicture: (path: string) => this.features.pictures.insertFromPath(path),
      activeRibbonTab: () => this.ribbon.activeTab,
      /** Ribbon commands (all tabs, contextual ones too) that do nothing yet and are not deferred. */
      unimplemented: () => {
        const ids = new Set<string>();
        const walk = (x: unknown): void => {
          if (!x || typeof x !== "object") return;
          if (Array.isArray(x)) { x.forEach(walk); return; }
          const o = x as Record<string, unknown>;
          if (typeof o.cmd === "string") ids.add(o.cmd);
          if (o.launcher) walk(o.launcher);
          for (const k of ["controls", "items", "rows", "menu", "smaller", "groups"]) if (k in o) { try { walk(o[k]); } catch { /* getters needing a document */ } }
        };
        const f = this.features;
        walk([...ribbonTabs(this.ribbonContext()), f.hf.tab, f.tables.designTab, f.tables.layoutTab, f.shapes.tab, f.pictures.tab, f.math.tab]);
        return [...ids].filter((id) => this.commands.has?.(id) !== false && !this.commands.isImplemented(id) && !this.commands.get(id).deferred).sort();
      },
      /** All registered commands (ribbon, menus, context menus) that do nothing yet and are not deferred. */
      unimplementedAll: () => this.commands.all().filter((c) => !c.run && !c.deferred).map((c) => `${c.id} (${c.label})`).sort(),
      /** Every ribbon control per tab with its label and tooltip (tooltip audit, Step 4c F). */
      ribbonTips: () => {
        const out: Record<string, { cmd: string; label: string; tip: string }[]> = {};
        const f = this.features;
        for (const tab of [...ribbonTabs(this.ribbonContext()), f.hf.tab, f.tables.designTab, f.tables.layoutTab, f.shapes.tab, f.pictures.tab, f.math.tab]) {
          const list: { cmd: string; label: string; tip: string }[] = (out[tab.label in out ? `${tab.label} (Tabelle)` : tab.label] = []);
          const walk = (x: unknown, inMenu: boolean): void => {
            if (!x || typeof x !== "object") return;
            if (Array.isArray(x)) { x.forEach((y) => walk(y, inMenu)); return; }
            const o = x as Record<string, unknown>;
            if (typeof o.cmd === "string" && this.commands.has(o.cmd) && !inMenu) {
              list.push({ cmd: o.cmd, label: (o.label as string) ?? this.commands.get(o.cmd).label, tip: this.commands.tooltip(o.cmd, true) });
            }
            if (o.launcher) walk(o.launcher, inMenu);
            for (const k of ["controls", "items", "rows", "groups"]) if (k in o) { try { walk(o[k], inMenu); } catch { /* getters needing a document */ } }
          };
          walk(tab, false);
        }
        return out;
      },
      infobar: () => this.infobarSlot.textContent,
      /** The active editor, for tests that call its commands directly. */
      editor: () => this.active?.editor,
      /** The main text's structure: style, list, cell and section of each paragraph. */
      docInfo: () => {
        const e = this.active?.editor;
        if (!e) return null;
        return {
          paras: e.doc.main.map((p) => ({
            text: p.text, style: e.doc.styles.resolvePara(p.props).style, numId: p.props.numId, ilvl: p.props.ilvl,
            cell: p.cell ? [p.cell.row, p.cell.col] : null, sect: !!p.sect,
          })),
          stories: [...e.doc.stories.keys()],
          pages: e.layout.pageCount,
          story: e.storyKey,
        };
      },
    };
  }
}
