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

/** A grammar finding of RaabeWrite's checker (D-051). */
export interface GrammarIssue {
  start: number;
  end: number;
  rule: string;
  category: string;
  message: string;
  short: string;
  suggestions: string[];
}

/** A thesaurus lookup: the word found and its meanings ("Adj." …; antonyms start with "!"). */
export interface ThesaurusResult {
  word: string;
  meanings: { pos: string; terms: string[] }[];
}

export interface Platform {
  readonly kind: "tauri" | "browser";
  /** Öffnen: RaabeWrite opens .docx, RaabeCalc .xlsx and .csv. */
  openFileDialog(kind?: "docx" | "xlsx"): Promise<string | null>;
  /** Chooses a picture file (Einfügen > Bilder > Dieses Gerät). */
  openImageDialog(): Promise<string | null>;
  /** Chooses a file of the given kinds (Entwurf > Designs > Nach Designs suchen …). */
  openAnyDialog?(title: string, filters: { name: string; extensions: string[] }[]): Promise<string | null>;
  saveFileDialog(defaultName: string, kind: "docx" | "pdf" | "xlsx" | "csv"): Promise<string | null>;
  /** Chooses where to save a file of the given kinds (Als Grafik speichern …). */
  saveAnyDialog?(title: string, defaultName: string, filters: { name: string; extensions: string[] }[]): Promise<string | null>;
  readFile(path: string): Promise<Uint8Array>;
  writeFile(path: string, bytes: Uint8Array): Promise<void>;
  fileInfo(path: string): Promise<FileInfo>;
  saveRecovery(entry: RecoveryEntry, bytes: Uint8Array): Promise<void>;
  listRecovery(): Promise<RecoveryEntry[]>;
  readRecovery(id: string): Promise<Uint8Array>;
  deleteRecovery(id: string): Promise<void>;
  clipboardRead(): Promise<string>;
  clipboardWrite(text: string): Promise<void>;
  /** The clipboard's whole HTML document (with Word's list definitions); "" when none. */
  clipboardReadHtml?(): Promise<string>;
  /** A picture on the clipboard (PNG or BMP file bytes); null when there is none. */
  clipboardReadImage?(): Promise<Uint8Array | null>;
  /** Puts HTML and its plain text on the clipboard, the HTML as it is. */
  clipboardWriteHtml?(html: string, text: string): Promise<void>;
  /** The page size in inches (A4 portrait when not given). */
  printToPdf(path: string, size?: { width: number; height: number }): Promise<void>;
  /** The folders and files of a folder (Link einfügen > Aktueller Ordner). */
  listDir?(path: string): Promise<{ name: string; dir: boolean }[]>;
  /** The user's Documents folder. */
  documentsDir?(): Promise<string>;
  /** Opens a web address, mail address or file with Windows' program for it (Strg+Klick on a link). */
  openExternal?(target: string): Promise<void>;
  /** Installed font families (for the Schriftart box). */
  listFonts(): Promise<string[]>;
  /** The characters a font has, as code point ranges; symbol fonts (Symbol, Wingdings) use U+F020-U+F0FF. */
  fontChars?(family: string): Promise<{ found: boolean; symbol: boolean; ranges: [number, number][] }>;
  /** Vertical font metrics from the font files (hhea table), keyed by family name. */
  fontMetrics(families: string[]): Promise<Record<string, { found: boolean; unitsPerEm: number; ascender: number; descender: number; lineGap: number; winAscent?: number; winDescent?: number; math?: boolean }>>;
  /** Spelling: for each word whether it is correct in `lang` ("de-DE", "en-US", "en-GB"). */
  spellCheck(lang: string, words: string[]): Promise<boolean[]>;
  /** Suggestions for a misspelled word. */
  spellSuggest(lang: string, word: string): Promise<string[]>;
  /** Grammar: the findings of each paragraph (UTF-16 offsets); none for languages without rules. */
  grammarCheck(lang: string, paragraphs: string[]): Promise<GrammarIssue[][]>;
  /** Thesaurus: the meanings of a word (or of its base form), null when there are none. */
  thesaurus(lang: string, word: string): Promise<ThesaurusResult | null>;
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
