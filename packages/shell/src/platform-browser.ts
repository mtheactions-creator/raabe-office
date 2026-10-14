// Development fallback when the app runs in a normal browser (npm run dev): files live in
// memory, dialogs are prompts on <input type=file>. Not used in the real app.

import type { Platform, RecoveryEntry } from "./platform";

export function browserPlatform(): Platform {
  const files = new Map<string, Uint8Array>();
  const recovery = new Map<string, { entry: RecoveryEntry; bytes: Uint8Array }>();
  return {
    kind: "browser",
    openFileDialog: () => new Promise((resolve) => {
      const input = document.createElement("input");
      input.type = "file";
      input.accept = ".docx";
      input.onchange = async () => {
        const f = input.files?.[0];
        if (!f) return resolve(null);
        files.set(f.name, new Uint8Array(await f.arrayBuffer()));
        resolve(f.name);
      };
      input.click();
    }),
    openAnyDialog: (_title, filters) => new Promise((resolve) => {
      const input = document.createElement("input");
      input.type = "file";
      input.accept = filters.flatMap((f) => f.extensions.map((x) => `.${x}`)).join(",");
      input.onchange = async () => {
        const f = input.files?.[0];
        if (!f) return resolve(null);
        files.set(f.name, new Uint8Array(await f.arrayBuffer()));
        resolve(f.name);
      };
      input.click();
    }),
    openImageDialog: () => new Promise((resolve) => {
      const input = document.createElement("input");
      input.type = "file";
      input.accept = "image/*";
      input.onchange = async () => {
        const f = input.files?.[0];
        if (!f) return resolve(null);
        files.set(f.name, new Uint8Array(await f.arrayBuffer()));
        resolve(f.name);
      };
      input.click();
    }),
    saveFileDialog: async (name) => name,
    readFile: async (path) => {
      const f = files.get(path);
      if (!f) throw new Error(`Datei nicht gefunden: ${path}`);
      return f;
    },
    writeFile: async (path, bytes) => { files.set(path, bytes); },
    fileInfo: async (path) => ({ exists: files.has(path), modified: null }),
    saveRecovery: async (entry, bytes) => { recovery.set(entry.id, { entry, bytes }); },
    listRecovery: async () => [...recovery.values()].map((r) => r.entry),
    readRecovery: async (id) => recovery.get(id)!.bytes,
    deleteRecovery: async (id) => { recovery.delete(id); },
    clipboardRead: () => navigator.clipboard.readText(),
    clipboardWrite: (t) => navigator.clipboard.writeText(t),
    printToPdf: async () => { print(); },
    listFonts: async () => ["Arial", "Calibri", "Calibri Light", "Cambria", "Comic Sans MS", "Courier New", "Georgia", "Segoe UI", "Times New Roman", "Verdana"],
    fontMetrics: async () => ({}),
    // The development platform has no dictionaries: every word counts as correct.
    spellCheck: async (_lang, words) => words.map(() => true),
    spellSuggest: async () => [],
    onOpenFiles: () => {},
    window: {
      minimize() {},
      toggleMaximize() {},
      close() {},
      isMaximized: async () => false,
      onResized: (h) => addEventListener("resize", h),
      setTitle: (t) => { document.title = t; },
      onCloseRequested: () => {},
      destroy() {},
    },
  };
}
