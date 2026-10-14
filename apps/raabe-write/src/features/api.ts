// What a feature module (lists, tables, page setup …) needs from the app.

import type { CharProps, CharState, Editor, ParaProps, ResolvedPara } from "@raabe/text-engine";
import type { Commands, Platform } from "@raabe/shell";

export interface AppApi {
  readonly commands: Commands;
  readonly platform: Platform;
  /** The editor of the active document. */
  editor(): Editor | undefined;
  /** Gives the keyboard focus back to the document. */
  focus(): void;
  /** Updates ribbon and status bar now. */
  refresh(): void;
  /** Selects a ribbon tab (contextual tabs included). */
  selectTab(id: string): void;
  /** Als Standardlayout festlegen: the wrapping new pictures get. */
  setDefaultPictureWrap?(wrap: string): void;
  /** The active document's file, if it has one. */
  documentPath(): string | null;
  /** Link einfügen > Neues Dokument erstellen: an empty document written to `path` (opened when `open`). */
  createEmptyDocument(path: string, open: boolean): Promise<void>;
  /** What the clipboard holds: text, HTML (formatting) and a picture. */
  clipboard(): Promise<{ text: string; html: string; image: Blob | null }>;
  /** The dialogs Schriftart and Absatz for a style's formatting (Formatvorlage ändern > Format). */
  fontDialog(target: FontTarget): Promise<void>;
  paragraphDialog(target: ParaTarget): Promise<void>;
  /** The fonts installed on this computer. */
  fontList(): readonly string[];
}

/** What Schriftart edits when it is not the selection: shown values and where changes go. */
export interface FontTarget {
  readonly state: CharState;
  apply(changes: CharProps): void;
}

export interface ParaTarget {
  readonly state: ResolvedPara;
  apply(changes: Partial<ParaProps>): void;
}

/** "2,5 cm" ↔ points. */
export const CM = 28.3465;
export const fmtCm = (pt: number) => `${(pt / CM).toLocaleString("de-DE", { maximumFractionDigits: 2 })} cm`;
export const parseNum = (v: string) => Number(v.replace(/[^\d,.-]/g, "").replace(",", "."));
