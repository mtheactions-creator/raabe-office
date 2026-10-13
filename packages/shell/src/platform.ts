// Everything that touches the operating system goes through this interface, so the same
// shell runs in Tauri on Windows, later on iPadOS, and in a plain browser for development.

export interface OpenedFile {
  path: string;
  name: string;
  bytes: Uint8Array;
}

export interface RecoveryEntry {
  id: string;
  title: string;
  /** Original file path, if the document had been saved before. */
  path: string | null;
  savedAt: number;
}

export interface FileInfo {
  exists: boolean;
  modified: number | null;
}

export interface Platform {
  readonly kind: "tauri" | "browser";
  openFileDialog(): Promise<string | null>;
  saveFileDialog(defaultName: string, kind: "docx" | "pdf"): Promise<string | null>;
  readFile(path: string): Promise<Uint8Array>;
  writeFile(path: string, bytes: Uint8Array): Promise<void>;
  fileInfo(path: string): Promise<FileInfo>;
  saveRecovery(entry: RecoveryEntry, bytes: Uint8Array): Promise<void>;
  listRecovery(): Promise<RecoveryEntry[]>;
  readRecovery(id: string): Promise<Uint8Array>;
  deleteRecovery(id: string): Promise<void>;
  clipboardRead(): Promise<string>;
  clipboardWrite(text: string): Promise<void>;
  printToPdf(path: string): Promise<void>;
  /** Files passed on the command line (double-click in Explorer) at start and later. */
  onOpenFiles(handler: (paths: string[]) => void): void;
  window: {
    minimize(): void;
    toggleMaximize(): void;
    close(): void;
    isMaximized(): Promise<boolean>;
    onResized(handler: () => void): void;
    setTitle(title: string): void;
    /** The handler decides whether the window may close (e.g. after "Speichern?"). */
    onCloseRequested(handler: () => Promise<boolean>): void;
    destroy(): void;
  };
}

export function baseName(path: string): string {
  return path.split(/[\\/]/).pop() ?? path;
}

export function dirName(path: string): string {
  const i = Math.max(path.lastIndexOf("\\"), path.lastIndexOf("/"));
  return i > 0 ? path.slice(0, i) : "";
}
