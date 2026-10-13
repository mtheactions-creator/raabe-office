// What a feature module (lists, tables, page setup …) needs from the app.

import type { Editor } from "@raabe/text-engine";
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
}

/** "2,5 cm" ↔ points. */
export const CM = 28.3465;
export const fmtCm = (pt: number) => `${(pt / CM).toLocaleString("de-DE", { maximumFractionDigits: 2 })} cm`;
export const parseNum = (v: string) => Number(v.replace(/[^\d,.-]/g, "").replace(",", "."));
