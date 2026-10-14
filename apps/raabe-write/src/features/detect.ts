// Sprache automatisch erkennen (on by default, as in Word): the language of a sentence among
// RaabeWrite's proofing languages, from its function words and letters. Word decides among
// the editing languages enabled in Office (recorded 2026-10-14: with German and English only,
// French stayed German and Spanish became English); RaabeWrite has German, English, French
// and Spanish. A sentence is only given a language when the evidence is clear; otherwise
// the text keeps its own language. Nothing is written into the document.

export type ProofLang = "de" | "en" | "fr" | "es";

const WORDS: Record<ProofLang, string> = {
  de: "der die das den dem des ein eine einen einem einer eines und oder aber ist sind war waren wird werden hat haben hatte nicht auch auf aus bei mit nach von vor zu zum zur im ins am um über unter für ich du er sie es wir ihr mich mir dich dir sich uns euch sein seine ihre ihr dass wenn weil als wie noch nur schon sehr kann können muss müssen wurde wurden diese dieser dieses jetzt hier dort dann denn doch man kein keine nichts etwas alle immer",
  en: "the a an and or but is are was were be been being has have had do does did not no to of in on at by for with from as that this these those it its he she they we you i me him her them my your our their his which who what when where why how will would can could should there here than then also just only very all some any more most other such into about over after before",
  fr: "le la les un une des du de et ou mais est sont était étaient être avoir a ont avait ne pas plus que qui quoi dont où dans sur sous avec pour par sans chez ce cette ces cet il elle ils elles nous vous je tu me te se mon ma mes ton ta tes son sa ses notre nos votre vos leur leurs au aux en y lui on très bien aussi comme tout tous toute toutes fait faire peut",
  es: "el la los las un una unos unas y o pero es son era eran ser estar está están fue han ha había no que qui quien cual cuando donde como en con por para sin sobre entre del al lo le les se su sus mi mis tu tus nuestro nuestra yo tú él ella ellos ellas nosotros vosotros usted muy más también porque pues este esta estos estas ese esa hay todo todos toda todas mucho hace puede",
};

const SETS = Object.fromEntries(Object.entries(WORDS).map(([k, v]) => [k, new Set(v.split(" "))])) as Record<ProofLang, Set<string>>;
const LANGS: ProofLang[] = ["de", "en", "fr", "es"];

/** Letters that belong to one language (or two). */
const LETTERS: [RegExp, Partial<Record<ProofLang, number>>][] = [
  [/[äöüß]/g, { de: 2 }],
  [/[ñ¿¡]/g, { es: 3 }],
  [/[áíóú]/g, { es: 1.5 }],
  [/[çœàèùâêîôûëï]/g, { fr: 1.5 }],
  [/é/g, { fr: 0.5, es: 0.5 }],
  [/\b(?:l|d|qu|j|n|c|s|m|t)['’]\p{L}/giu, { fr: 2 }],
];

/** The language of a piece of text, or null when it is not clear (fewer than four words …). */
export function detectLanguage(text: string): ProofLang | null {
  const words = text.toLocaleLowerCase().match(/\p{L}+/gu) ?? [];
  if (words.length < 4) return null;
  const score: Record<ProofLang, number> = { de: 0, en: 0, fr: 0, es: 0 };
  for (const w of words) for (const l of LANGS) if (SETS[l].has(w)) score[l] += 1;
  const lower = text.toLocaleLowerCase();
  for (const [re, add] of LETTERS) {
    const n = lower.match(re)?.length ?? 0;
    if (n) for (const [l, v] of Object.entries(add) as [ProofLang, number][]) score[l] += Math.min(n, 4) * v;
  }
  const ranked = [...LANGS].sort((a, b) => score[b] - score[a]);
  const [best, second] = [score[ranked[0]], score[ranked[1]]];
  if (best < 3 || best < second * 2 || best < words.length * 0.2) return null;
  return ranked[0];
}

/** The sentences of a paragraph with their detected language (offsets into the text). */
export function detectSentences(text: string): { a: number; b: number; lang: ProofLang | null }[] {
  const out: { a: number; b: number; lang: ProofLang | null }[] = [];
  const re = /[^.!?…]*(?:[.!?…]+["“”»«')\]]*\s*|$)/gu;
  for (const m of text.matchAll(re)) {
    if (!m[0]) continue;
    out.push({ a: m.index!, b: m.index! + m[0].length, lang: detectLanguage(m[0]) });
    if (m.index! + m[0].length >= text.length) break;
  }
  return out;
}
