// Kommentare (Excel 365's threaded comments): the threads of a sheet in
// xl/threadedComments/threadedCommentN.xml, their authors in xl/persons/person.xml. Excel keeps a
// legacy note for each thread (author "tc={id}") for older programs; RaabeCalc reads the threads,
// hides those notes, and writes them again with a short fallback text of its own.

import { cellName, parseCell, type CommentThread } from "@raabe/sheet-engine";
import { esc, escText, XML_HEAD } from "./xml";

export const RT_THREADS = "http://schemas.microsoft.com/office/2017/10/relationships/threadedComment";
export const RT_PERSON = "http://schemas.microsoft.com/office/2017/10/relationships/person";
export const CT_THREADS = "application/vnd.ms-excel.threadedcomments+xml";
export const CT_PERSON = "application/vnd.ms-excel.person+xml";
const NS = "http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments";

const unesc = (s: string) => s.replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&#10;/g, "\n").replace(/&amp;/g, "&");
const attr = (tag: string, name: string) => { const m = new RegExp(`\\b${name}="([^"]*)"`).exec(tag); return m ? unesc(m[1]) : undefined; };

/** The persons of the workbook: id → display name. */
export function readPersons(xml: string | null): Map<string, string> {
  const out = new Map<string, string>();
  for (const m of (xml ?? "").matchAll(/<person\b[^>]*\/?>/g)) {
    const id = attr(m[0], "id");
    if (id) out.set(id, attr(m[0], "displayName") ?? "");
  }
  return out;
}

/** The threads of a sheet: first comments with their replies, in the file's order. */
export function readThreads(xml: string, persons: Map<string, string>): CommentThread[] {
  const out: CommentThread[] = [];
  const byId = new Map<string, CommentThread>();
  for (const m of xml.matchAll(/<threadedComment\b([^>]*)>([\s\S]*?)<\/threadedComment>/g)) {
    const tag = m[1];
    const text = unesc(/<text>([\s\S]*?)<\/text>/.exec(m[2])?.[1] ?? "");
    const at = parseCell(attr(tag, "ref") ?? "");
    const id = attr(tag, "id") ?? "";
    const entry = { id, author: persons.get(attr(tag, "personId") ?? "") ?? "", personId: attr(tag, "personId"), date: attr(tag, "dT") ?? "", text };
    const parent = attr(tag, "parentId");
    if (parent && byId.has(parent)) { byId.get(parent)!.replies.push(entry); continue; }
    if (!at) continue;
    const t: CommentThread = { row: at.row, col: at.col, ...entry, replies: [], done: attr(tag, "done") === "1" || undefined };
    byId.set(id, t);
    out.push(t);
  }
  return out;
}

/** A GUID in Excel's form ({XXXXXXXX-XXXX-…}). */
export function guid(): string {
  const b = crypto.getRandomValues(new Uint8Array(16));
  b[6] = (b[6] & 0x0f) | 0x40;
  b[8] = (b[8] & 0x3f) | 0x80;
  const h = [...b].map((x) => x.toString(16).padStart(2, "0").toUpperCase()).join("");
  return `{${h.slice(0, 8)}-${h.slice(8, 12)}-${h.slice(12, 16)}-${h.slice(16, 20)}-${h.slice(20)}}`;
}

/** threadedCommentN.xml for a sheet's threads; persons by display name → id. */
export function threadsXml(threads: CommentThread[], personOf: (name: string) => string): string {
  const one = (ref: string, e: { id: string; author: string; personId?: string; date: string; text: string }, extra: string) =>
    `<threadedComment ref="${ref}" dT="${esc(e.date)}" personId="${esc(e.personId ?? personOf(e.author))}" id="${esc(e.id)}"${extra}><text>${escText(e.text)}</text></threadedComment>`;
  const body = threads.map((t) => {
    const ref = cellName(t.row, t.col);
    return one(ref, t, t.done ? ' done="1"' : "") + t.replies.map((r) => one(ref, r, ` parentId="${esc(t.id)}"`)).join("");
  }).join("");
  return `${XML_HEAD}<ThreadedComments xmlns="${NS}" xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main">${body}</ThreadedComments>`;
}

/** person.xml with the given persons (id → name). */
export function personsXml(persons: Map<string, string>): string {
  const list = [...persons].map(([id, name]) => `<person displayName="${esc(name)}" id="${esc(id)}" userId="${esc(name)}" providerId="None"/>`).join("");
  return `${XML_HEAD}<personList xmlns="${NS}" xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main">${list}</personList>`;
}

/** The legacy note kept for older programs (our own words). */
export function fallbackText(t: CommentThread): string {
  return [`[Kommentar] ${t.author}: ${t.text}`, ...t.replies.map((r) => `${r.author}: ${r.text}`)].join("\n");
}
