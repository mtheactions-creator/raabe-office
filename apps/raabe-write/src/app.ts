// RaabeWrite: wires the shell (title bar, ribbon, backstage, status bar) to documents.
// Each open document is a tab with its own editor. Files are .docx (D-004); content we
// cannot keep yet is detected on open and the original file is never overwritten.

import { Editor, PAGE, type Para } from "@raabe/text-engine";
import { readDocx, writeDocx } from "@raabe/ooxml";
import {
  Backstage, Commands, Keytips, Ribbon, StatusBar, TitleBar, baseName, dirName, h, icon, installTooltips, messageBox,
  showDialog, type BackstageEntry, type Platform,
} from "@raabe/shell";
import { COMMANDS } from "./command-list";
import { ribbonTabs } from "./ribbon-def";
import { addRecent, loadRecent, loadSettings, removeRecent, saveSettings, setPinned, type Settings } from "./store";
import appIconUrl from "../../../branding/raabe-write-small.svg?url";
import logoUrl from "../../../branding/raabe-write.svg?url";

interface Doc {
  id: string;
  name: string;
  path: string | null;
  editor: Editor;
  ws: HTMLElement;
  /** Paragraph array at the last save; the document is dirty when the editor's differs. */
  saved: readonly Para[];
  recoverySaved: readonly Para[] | null;
  /** Content found in the file that RaabeWrite cannot keep (German labels). */
  unsupported: string[];
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
  private docArea: HTMLElement;
  private infobarSlot: HTMLElement;
  private printRoot: HTMLElement;
  private wordTimer = 0;
  private refreshTimer = 0;
  private lastTabsKey = "";

  constructor(private platform: Platform, initialParas?: Para[]) {
    this.commands.register(...COMMANDS);
    this.implementCommands();

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
    const comments = h("button", { class: "pill", disabled: true, "data-tip": "Kommentare\nNoch nicht verfügbar" }, icon("comment"), "Kommentare");
    this.ribbon = new Ribbon(this.commands, ribbonTabs(() => this.active?.editor ?? null), {
      onFile: () => this.openBackstage("home"),
      tabRowRight: [comments],
    });
    this.statusBar = new StatusBar(this.commands, (z) => this.setZoom(z), [{ cmd: "view.readMode" }, { cmd: "view.printLayout" }], "view.zoomDialog");
    this.backstage = new Backstage(this.commands, this.backstageEntries());
    this.backstage.onClose(() => this.active?.editor.focus());
    this.docArea = h("div", { class: "docarea" });
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

    this.newDoc(initialParas);
    void this.offerRecovery();
    platform.onOpenFiles((paths) => { for (const p of paths) void this.openPath(p); });
  }

  // ----- documents ----------------------------------------------------------------------

  private createDoc(name: string, path: string | null, paras: Para[] | undefined, unsupported: string[] = []): Doc {
    const ws = h("div", { class: "workspace", tabIndex: -1 });
    this.docArea.append(ws);
    const editor = new Editor(ws, paras ?? [{ style: "normal", text: "" }]);
    const doc: Doc = {
      id: crypto.randomUUID(), name, path, editor, ws, saved: editor.paragraphs(), recoverySaved: null, unsupported,
      lastSaved: null,
    };
    editor.onChange((e) => this.onEditorChange(doc, e.content));
    this.docs.push(doc);
    this.activate(doc);
    return doc;
  }

  newDoc(paras?: Para[]): Doc {
    this.untitled++;
    return this.createDoc(`Dokument${this.untitled}`, null, paras);
  }

  private isDirty(d: Doc): boolean {
    return d.editor.paragraphs() !== d.saved;
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
    this.statusBar.update({ page: s.caretPage + 1, pages: s.pageCount, words: this.wordCache.words, zoom: d.editor.zoom });
  }

  private renderInfobar(): void {
    this.infobarSlot.replaceChildren();
    const d = this.active;
    if (!d || !d.unsupported.length) return;
    this.infobarSlot.append(h("div", { class: "infobar", role: "alert" },
      icon("info"),
      h("span", { class: "msg" }, `Dieses Dokument enthält Inhalte, die RaabeWrite noch nicht darstellen kann: ${d.unsupported.join(", ")}. ` +
        "Sie fehlen in der Ansicht. Damit sie im Original erhalten bleiben, speichern Sie Änderungen unter einem neuen Namen."),
      h("button", { onclick: () => void this.saveAs(d) }, "Speichern unter...")));
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
    this.createDoc(stripExt(baseName(path)), path, result.paragraphs.map((p) => ({ style: p.style, text: p.text })), result.unsupported);
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
    if (d.path && !d.unsupported.length) return this.writeTo(d, d.path);
    if (d.path && d.unsupported.length) {
      const choice = await showDialog({
        title: "RaabeWrite",
        body: `„${d.name}“ enthält Inhalte, die RaabeWrite noch nicht unterstützt (${d.unsupported.join(", ")}). ` +
          "Beim Speichern in die Originaldatei gingen sie verloren. Speichern Sie deshalb unter einem neuen Namen.",
        buttons: [{ label: "Speichern unter...", value: "saveAs", primary: true }, { label: "Abbrechen", value: "cancel" }],
        cancel: "cancel",
      });
      if (choice !== "saveAs") return false;
    }
    return this.saveAs(d);
  }

  async saveAs(d: Doc): Promise<boolean> {
    let path = await this.platform.saveFileDialog(this.suggestedName(d), "docx");
    if (!path) return false;
    if (!/\.docx$/i.test(path)) path += ".docx";
    if (d.unsupported.length && d.path && samePath(path, d.path)) {
      await messageBox("RaabeWrite", "Bitte wählen Sie einen anderen Namen, damit die Originaldatei mit allen Inhalten erhalten bleibt.");
      return false;
    }
    const other = this.docs.find((x) => x !== d && x.path && samePath(x.path, path!));
    if (other) {
      await messageBox("RaabeWrite", `„${baseName(path)}“ ist bereits in einem anderen Tab geöffnet. Schließen Sie es zuerst.`);
      return false;
    }
    return this.writeTo(d, path);
  }

  private async writeTo(d: Doc, path: string): Promise<boolean> {
    const paras = d.editor.paragraphs();
    try {
      await this.platform.writeFile(path, writeDocx(paras.map((p) => ({ style: p.style, text: p.text })), { author: this.settings.userName }));
    } catch (e) {
      await messageBox("RaabeWrite", `Das Dokument konnte nicht unter „${path}“ gespeichert werden. ` +
        `Vielleicht ist die Datei in einem anderen Programm geöffnet.\n\n${String(e)}`);
      return false;
    }
    d.path = path;
    d.name = stripExt(baseName(path));
    d.saved = paras;
    d.unsupported = [];
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
      const paras = d.editor.paragraphs();
      if (!this.isDirty(d) || paras === d.recoverySaved) continue;
      try {
        await this.platform.saveRecovery({ id: d.id, title: d.name, path: d.path, savedAt: Date.now() },
          writeDocx(paras.map((p) => ({ style: p.style, text: p.text })), { author: this.settings.userName }));
        d.recoverySaved = paras;
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
          const d = this.createDoc(`${e.title} (Wiederhergestellt)`, e.path, r.paragraphs.map((p) => ({ style: p.style, text: p.text })));
          d.saved = [];
          if (pristine) this.removeDoc(pristine);
          this.updateChrome(true);
        } catch { /* unreadable recovery file: skip */ }
      }
      await this.platform.deleteRecovery(e.id).catch(() => {});
    }
  }

  // ----- printing and PDF ---------------------------------------------------------------

  private async print(): Promise<void> {
    if (!this.active) return;
    this.backstage.close();
    this.active.editor.renderForPrint(this.printRoot);
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

  private async optionsDialog(): Promise<void> {
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
    const navBtn = (label: string, panel: HTMLElement) => h("button", {
      onclick: (e: MouseEvent) => {
        general.hidden = panel !== general; saving.hidden = panel !== saving;
        nav.querySelectorAll("button").forEach((b) => b.classList.toggle("active", b === e.currentTarget));
      },
    }, label);
    const nav = h("nav", null, navBtn("Allgemein", general), navBtn("Speichern", saving));
    nav.firstElementChild!.classList.add("active");
    const ok = await showDialog({
      title: "RaabeWrite-Optionen", body: h("div", { class: "options" }, nav, general, saving),
      buttons: [{ label: "OK", value: true, primary: true }, { label: "Abbrechen", value: false }], cancel: false,
    });
    if (!ok) return;
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
    const school = h("div", { class: "template", style: "opacity:.55", "data-tip": "Raabeschule-Vorlagen\nFolgen in einem späteren Schritt." },
      h("div", { class: "thumb" }, h("img", { src: logoUrl, width: 48, alt: "" })), "Arbeitsblatt (Raabeschule)");
    const facharbeit = h("div", { class: "template", style: "opacity:.55", "data-tip": "Raabeschule-Vorlagen\nFolgen in einem späteren Schritt." },
      h("div", { class: "thumb" }, h("img", { src: logoUrl, width: 48, alt: "" })), "Facharbeit (Raabeschule)");
    return h("div", { class: "templates" }, blank, school, facharbeit);
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
      d.unsupported.length ? h("p", { class: "warn" }, `Nicht unterstützte Inhalte in der Originaldatei: ${d.unsupported.join(", ")}.`) : null);
  }

  private printPage(): HTMLElement {
    const preview = h("div", { class: "print-preview" });
    const holder = h("div");
    this.active?.editor.renderForPrint(holder);
    const first = holder.firstElementChild as HTMLElement | null;
    if (first) {
      const scale = 0.55;
      first.classList.add("sheet");
      first.style.transform = `scale(${scale})`;
      preview.append(h("div", { style: `width:${PAGE.width * scale}px;height:${PAGE.height * scale}px` }, first));
    }
    const pages = this.active?.editor.stats().pageCount ?? 0;
    return h("div", null, h("h1", null, "Drucken"),
      h("div", { class: "print-layout" },
        h("div", { style: "display:flex;flex-direction:column;gap:12px;width:220px" },
          h("button", { class: "bigbtn", onclick: () => void this.print() }, icon("print", "l"), "Drucken"),
          h("p", { class: "muted" }, `${pages} ${pages === 1 ? "Seite" : "Seiten"}, A4. Drucker, Exemplare und Seitenauswahl wählen Sie im nächsten Fenster.`)),
        preview));
  }

  // ----- commands -----------------------------------------------------------------------

  private implementCommands(): void {
    const c = this.commands;
    const ed = () => this.active?.editor;
    const has = () => !!this.active;
    c.implement("file.new", { run: () => { this.backstage.close(); this.newDoc(); } });
    c.implement("file.open", { run: () => this.openBackstage("open") });
    c.implement("file.save", { run: async () => { if (this.active) { this.backstage.close(); await this.save(this.active); } }, enabled: has });
    c.implement("file.saveAs", { run: async () => { if (this.active) { this.backstage.close(); await this.saveAs(this.active); } }, enabled: has });
    c.implement("file.print", { run: () => this.openBackstage("print"), enabled: has });
    c.implement("file.exportPdf", { run: () => this.exportPdf(), enabled: has });
    c.implement("file.close", { run: async () => { if (this.active) { this.backstage.close(); await this.closeDoc(this.active); } }, enabled: has });
    c.implement("file.options", { run: () => this.optionsDialog() });
    c.implement("file.info", { run: () => this.openBackstage("info") });

    c.implement("edit.undo", { run: () => ed()?.undo(), enabled: () => !!ed()?.canUndo });
    c.implement("edit.redo", { run: () => ed()?.redo(), enabled: () => !!ed()?.canRedo });
    c.implement("edit.copy", { run: async () => { const e = ed(); if (e) await this.platform.clipboardWrite(e.selectedText().replace(/\n/g, "\r\n")); }, enabled: () => !!ed()?.hasSelection });
    c.implement("edit.cut", { run: async () => { const e = ed(); if (e) await this.platform.clipboardWrite(e.cutSelection().replace(/\n/g, "\r\n")); }, enabled: () => !!ed()?.hasSelection });
    const paste = async () => { const t = await this.platform.clipboardRead().catch(() => ""); if (t) ed()?.pasteText(t); };
    c.implement("edit.paste", { run: paste, enabled: has });
    c.implement("edit.pasteText", { run: paste, enabled: has });
    c.implement("edit.selectAll", { run: () => ed()?.selectAll(), enabled: has });

    c.implement("style.normal", { run: () => ed()?.setParagraphStyle("normal"), checked: () => ed()?.currentStyle() === "normal", enabled: has });
    c.implement("style.h1", { run: () => ed()?.setParagraphStyle("h1"), checked: () => ed()?.currentStyle() === "h1", enabled: has });
    c.get("style.normal").shortcut = "Strg+Umschalt+N";

    c.implement("view.printLayout", { run: () => {}, checked: () => true });
    c.implement("view.zoom100", { run: () => this.setZoom(1), enabled: has });
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

  private installKeyboard(): void {
    addEventListener("keydown", (e) => {
      if (e.defaultPrevented || e.isComposing) return;
      const ctrl = e.ctrlKey, shift = e.shiftKey, alt = e.altKey;
      const k = e.key.length === 1 ? e.key.toLowerCase() : e.key;
      let cmd: string | null = null;
      if (ctrl && !alt && !shift) {
        cmd = ({ n: "file.new", o: "file.open", s: "file.save", p: "file.print", w: "file.close" } as Record<string, string>)[k] ?? null;
        if (k === "F4") cmd = "file.close";
        if (k === "F12") { e.preventDefault(); void this.openDialog(); return; }
        if (k === "F1") cmd = "view.ribbon";
      }
      if (!ctrl && !alt && k === "F12") cmd = "file.saveAs";
      if (!ctrl && !alt && shift && k === "F7") cmd = "review.thesaurus";
      if (ctrl && alt && !shift && k === "1") cmd = "style.h1";
      if (ctrl && shift && !alt && k === "n") cmd = "style.normal";
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
      if (t.closest(".ribbon-card, .qat, .tabrow, .statusbar, .group-popup, .menu")) e.preventDefault();
    }, true);
  }

  /** The active editor, for the feel-lab probe. */
  activeEditor(): Editor {
    return this.active!.editor;
  }

  /** Test hooks for feel-lab runs only (no file dialogs; see tools/feel-lab/smoke_write.py). */
  testApi() {
    return {
      docs: () => this.docs.map((d) => ({ name: d.name, path: d.path, dirty: this.isDirty(d), unsupported: d.unsupported, active: d === this.active })),
      openPath: (path: string) => this.openPath(path),
      saveTo: (path: string) => (this.active ? this.writeTo(this.active, path) : Promise.resolve(false)),
      save: () => (this.active ? this.save(this.active) : Promise.resolve(false)),
      exportPdfTo: (path: string) => this.exportPdf(path),
      autoRecoverNow: () => this.autoRecover(),
      run: (id: string) => this.commands.run(id),
      enabled: (id: string) => this.commands.isEnabled(id),
      newDoc: () => { this.newDoc(); },
      activeRibbonTab: () => this.ribbon.activeTab,
      infobar: () => this.infobarSlot.textContent,
    };
  }
}
