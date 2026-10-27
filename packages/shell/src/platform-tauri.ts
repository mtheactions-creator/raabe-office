// Platform implementation for Tauri (Windows now, iPadOS later). File access goes through
// our own Rust commands (see apps/*/src-tauri/src/lib.rs) so saving is atomic.

import { invoke } from "@tauri-apps/api/core";
import { listen } from "@tauri-apps/api/event";
import { getCurrentWindow } from "@tauri-apps/api/window";
import { open, save } from "@tauri-apps/plugin-dialog";
import { readText, writeText } from "@tauri-apps/plugin-clipboard-manager";
import type { FileInfo, GrammarIssue, Platform, RecoveryEntry, ThesaurusResult } from "./platform";

const FILTERS = {
  docx: [{ name: "Word-Dokument", extensions: ["docx"] }],
  xlsx: [{ name: "Excel-Arbeitsmappe", extensions: ["xlsx"] }],
  csv: [{ name: "CSV (Trennzeichen-getrennt)", extensions: ["csv"] }],
  xlsxOpen: [{ name: "Alle Excel-Dateien", extensions: ["xlsx", "csv"] }],
  pdf: [{ name: "PDF", extensions: ["pdf"] }],
  image: [{ name: "Alle Bilder", extensions: ["png", "jpg", "jpeg", "gif", "bmp", "svg", "webp"] }],
};

export function tauriPlatform(): Platform {
  const win = getCurrentWindow();
  return {
    kind: "tauri",
    async openAnyDialog(title, filters) {
      const r = await open({ multiple: false, directory: false, filters, title });
      return typeof r === "string" ? r : null;
    },
    async openFileDialog(kind = "docx") {
      const r = await open({ multiple: false, directory: false, filters: kind === "xlsx" ? FILTERS.xlsxOpen : FILTERS.docx, title: "Öffnen" });
      return typeof r === "string" ? r : null;
    },
    async openImageDialog() {
      const r = await open({ multiple: false, directory: false, filters: FILTERS.image, title: "Grafik einfügen" });
      return typeof r === "string" ? r : null;
    },
    async saveAnyDialog(title, defaultName, filters) {
      return (await save({ defaultPath: defaultName, filters, title })) ?? null;
    },
    async saveFileDialog(defaultName, kind) {
      return (await save({ defaultPath: defaultName, filters: FILTERS[kind], title: kind === "pdf" ? "Als PDF exportieren" : "Speichern unter" })) ?? null;
    },
    async readFile(path) {
      return new Uint8Array(await invoke<ArrayBuffer>("read_file", { path }));
    },
    async writeFile(path, bytes) {
      await invoke("write_file", bytes, { headers: { path: encodeURIComponent(path) } });
    },
    fileInfo: (path) => invoke<FileInfo>("file_info", { path }),
    async saveRecovery(entry, bytes) {
      await invoke("save_recovery", bytes, { headers: { meta: encodeURIComponent(JSON.stringify(entry)) } });
    },
    listRecovery: () => invoke<RecoveryEntry[]>("list_recovery"),
    async readRecovery(id) {
      return new Uint8Array(await invoke<ArrayBuffer>("read_recovery", { id }));
    },
    deleteRecovery: (id) => invoke("delete_recovery", { id }),
    clipboardRead: async () => (await readText()) ?? "",
    clipboardWrite: (text) => writeText(text),
    clipboardReadHtml: () => invoke<string>("clipboard_read_html"),
    async clipboardReadImage() {
      const bytes = new Uint8Array(await invoke<ArrayBuffer>("clipboard_read_image"));
      return bytes.length ? bytes : null;
    },
    clipboardWriteHtml: (html, text) => invoke("clipboard_write_html", { html, text }),
    printToPdf: (path, size) => invoke("print_to_pdf", { path, width: size?.width ?? null, height: size?.height ?? null }),
    listFonts: () => invoke<string[]>("list_fonts"),
    fontChars: (family) => invoke("font_chars", { family }),
    listDir: (path) => invoke("list_dir", { path }),
    documentsDir: () => invoke<string>("documents_dir"),
    openExternal: (target) => invoke("open_external", { target }),
    fontMetrics: (families) => invoke("font_metrics", { families }),
    spellCheck: (lang, words) => invoke<boolean[]>("spell_check", { lang, words }),
    spellSuggest: (lang, word) => invoke<string[]>("spell_suggest", { lang, word }),
    thesaurus: (lang, word) => invoke<ThesaurusResult | null>("thesaurus_lookup", { lang, word }),
    grammarCheck: (lang, paragraphs) => invoke<GrammarIssue[][]>("grammar_check", { lang, paragraphs }),
    onOpenFiles(handler) {
      void invoke<string[]>("startup_files").then((paths) => { if (paths.length) handler(paths); });
      void listen<string[]>("open-files", (e) => handler(e.payload));
    },
    window: {
      minimize: () => void win.minimize(),
      toggleMaximize: () => void win.toggleMaximize(),
      close: () => void win.close(),
      isMaximized: () => win.isMaximized(),
      onResized: (h) => void win.onResized(() => h()),
      setTitle: (t) => void win.setTitle(t),
      onCloseRequested(handler) {
        void win.onCloseRequested(async (e) => {
          e.preventDefault();
          if (await handler()) void win.destroy();
        });
      },
      destroy: () => void win.destroy(),
    },
  };
}
