// Start > Absatz > Sortieren (Word's dialog Text sortieren): paragraphs or table rows by up to
// three keys, each as text, number or date, ascending or descending. Word's rules: empty and
// key-less entries come first when ascending, text ignores case unless asked, numbers are read
// with a German decimal comma, dates as 14.10.2026, 14. Oktober 2026 or 2026-10-14.

export type SortType = "text" | "number" | "date";

export interface SortKey {
  /** 0: the whole paragraph (Absätze); n ≥ 1: field n (paragraphs split by the separator) or column n. */
  readonly field: number;
  readonly type: SortType;
  readonly desc: boolean;
}

export interface SortSpec {
  readonly keys: readonly SortKey[];
  /** Meine Liste enthält: Überschrift (the first entry stays first). */
  readonly header: boolean;
  /** Trennzeichen for fields: "\t" (Tabstopps), "," (Kommas) or another character. */
  readonly separator: string;
  readonly caseSensitive: boolean;
  /** Sortiersprache (BCP 47), German by default. */
  readonly lang?: string;
}

const MONTHS = ["januar", "februar", "märz", "april", "mai", "juni", "juli", "august", "september", "oktober", "november", "dezember"];
const MONTHS_EN = ["january", "february", "march", "april", "may", "june", "july", "august", "september", "october", "november", "december"];

/** The first number in a text (German: 1.234,5), NaN when there is none. */
export function sortNumber(text: string): number {
  const m = /[-+−]?\d[\d.\s]*(?:,\d+)?/.exec(text);
  if (!m) return NaN;
  const s = m[0].replace(/[\s.]/g, "").replace(",", ".").replace("−", "-");
  return Number(s);
}

/** A date in a text as a sortable number (yyyymmdd), NaN when there is none. */
export function sortDate(text: string): number {
  const t = text.toLowerCase();
  const iso = /(\d{4})-(\d{1,2})-(\d{1,2})/.exec(t);
  if (iso) return Number(iso[1]) * 10000 + Number(iso[2]) * 100 + Number(iso[3]);
  const year = (y: string) => (y.length <= 2 ? 2000 + Number(y) - (Number(y) > 50 ? 100 : 0) : Number(y));
  const dmy = /(\d{1,2})\.\s*(\d{1,2})\.\s*(\d{2,4})/.exec(t);
  if (dmy) return year(dmy[3]) * 10000 + Number(dmy[2]) * 100 + Number(dmy[1]);
  const named = /(\d{1,2})\.?\s+([a-zäöü]+)\.?\s+(\d{2,4})/.exec(t);
  if (named) {
    const i = [MONTHS, MONTHS_EN].map((list) => list.findIndex((m) => m.startsWith(named[2].slice(0, 3)))).find((x) => x >= 0) ?? -1;
    if (i >= 0) return year(named[3]) * 10000 + (i + 1) * 100 + Number(named[1]);
  }
  return NaN;
}

/**
 * The order of `entries` (each a list of field texts; index 0 is the whole entry) by `spec`:
 * a permutation of the indices. Stable, so equal entries keep their order.
 */
export function sortOrder(entries: readonly (readonly string[])[], spec: SortSpec): number[] {
  const coll = new Intl.Collator(spec.lang ?? "de-DE", { sensitivity: spec.caseSensitive ? "variant" : "accent", caseFirst: spec.caseSensitive ? "lower" : "false" });
  const keyOf = (e: readonly string[], k: SortKey): string | number => {
    const text = (e[k.field] ?? "").trim();
    return k.type === "text" ? text : k.type === "number" ? sortNumber(text) : sortDate(text);
  };
  const keys = entries.map((e) => spec.keys.map((k) => keyOf(e, k)));
  const idx = entries.map((_, i) => i);
  const first = spec.header && idx.length ? [idx.shift()!] : [];
  idx.sort((a, b) => {
    for (let k = 0; k < spec.keys.length; k++) {
      const x = keys[a][k], y = keys[b][k];
      let d: number;
      if (typeof x === "number" && typeof y === "number") {
        // Entries without a number (date) come first, as empty text does.
        d = Number.isNaN(x) ? (Number.isNaN(y) ? 0 : -1) : Number.isNaN(y) ? 1 : x - y;
      } else d = coll.compare(String(x), String(y));
      if (d) return spec.keys[k].desc ? -d : d;
    }
    return a - b;
  });
  return [...first, ...idx];
}
