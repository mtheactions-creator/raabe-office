// Languages as Word names them (dialog Sprache, status bar) and the proofing languages
// RaabeWrite has: spelling and thesaurus for German, English, French and Spanish (D-049).

import { LANGUAGES } from "./languages.gen";
import type { ProofLang } from "./detect";

export { LANGUAGES };

const NAMES = new Map(LANGUAGES.map(([t, n]) => [t.toLowerCase(), n]));

/** Word's name of a language tag ("de-DE" → "Deutsch (Deutschland)"). */
export const languageName = (tag: string) => NAMES.get(tag.toLowerCase()) ?? tag;

/** The proofing language of a tag (de-AT → de), null when RaabeWrite has no tools for it. */
export function proofLang(tag: string | undefined): ProofLang | null {
  const l = (tag ?? "").slice(0, 2).toLowerCase();
  return l === "de" || l === "en" || l === "fr" || l === "es" ? l : null;
}

/** The tag a detected language gets, near the text's own (en-GB text stays British). */
export function tagFor(lang: ProofLang, near: string): string {
  if (proofLang(near) === lang) return near;
  return { de: "de-DE", en: "en-US", fr: "fr-FR", es: "es-ES" }[lang];
}

/** The languages at the top of the dialog Sprache (Word: the Office editing languages). */
export const PREFERRED = ["de-DE", "en-US", "en-GB", "fr-FR", "es-ES"];
