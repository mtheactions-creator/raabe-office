// Referenzen > Zitate und Literaturverzeichnis as in Word (recorded with probe_word_dialogs.py
// and probe_word_source_dialog.py): Zitat einfügen (the document's sources, Neue Quelle
// hinzufügen, Neuen Platzhalter hinzufügen), Quelle erstellen/bearbeiten with Word's fields per
// Quellentyp, Quellen verwalten (Masterliste of this user and Aktuelle Liste of the document),
// Formatvorlage (Word's twelve styles) and the Literaturverzeichnis gallery.

import { h, icon, type MenuEntry } from "@raabe/shell";
import {
  CITATION_STYLES, MORE_FIELDS, ROLES, SOURCE_TYPES, formatCitation, formatEntry, parsePeople, peopleText, styleByName, suggestTag,
  type Bibliography, type Contributors, type Source, type SourceType,
} from "@raabe/text-engine";
import type { AppApi } from "./api";
import { openModal } from "./modal";

const LANGUAGES: [string, string][] = [["", "Standard"], ["1031", "Deutsch (Deutschland)"], ["3079", "Deutsch (Österreich)"], ["2055", "Deutsch (Schweiz)"],
  ["1033", "Englisch (Vereinigte Staaten)"], ["2057", "Englisch (Vereinigtes Königreich)"], ["1036", "Französisch (Frankreich)"], ["3082", "Spanisch (Spanien)"],
  ["1040", "Italienisch (Italien)"], ["1043", "Niederländisch (Niederlande)"], ["1045", "Polnisch"], ["1055", "Türkisch"], ["1049", "Russisch"]];

const master = (): Source[] => {
  try { const v = JSON.parse(localStorage.getItem("raabe.sources.master") ?? "[]"); return Array.isArray(v) ? v : []; } catch { return []; }
};
const saveMaster = (list: Source[]) => { try { localStorage.setItem("raabe.sources.master", JSON.stringify(list)); } catch { /* private mode */ } };
const remember = (s: Source) => saveMaster([...master().filter((x) => x.tag !== s.tag), s]);

const firstAuthor = (s: Source) => peopleText(s.contributors?.Author ?? Object.values(s.contributors ?? {})[0]);
const describe = (s: Source) => `${firstAuthor(s) || s.tag}; ${s.fields.Title ?? ""}${s.fields.Year ? ` (${s.fields.Year})` : ""}`;

export function citationsFeature(api: AppApi) {
  const c = api.commands;
  const ed = () => api.editor();
  const bib = (): Bibliography => ed()?.bibliography() ?? { sources: [], style: "APA" };
  const setBib = (b: Bibliography) => { ed()?.setBibliography(b); api.refresh(); };

  // ----- Quelle erstellen / bearbeiten ----------------------------------------------------------
  const sourceDialog = (old?: Source): Promise<Source | null> => new Promise((resolve) => {
    let type: SourceType = old?.type ?? "Book";
    let all = false;
    let tagTouched = !!old;
    const values = new Map<string, string>();
    for (const [k, v] of Object.entries(old?.fields ?? {})) values.set(k, v);
    for (const [k, v] of Object.entries(old?.contributors ?? {})) values.set(k, v.corporate ?? peopleText(v));
    let corporate = !!old?.contributors?.Author?.corporate;
    const typeSel = h("select", {}, ...SOURCE_TYPES.map((t) => h("option", { value: t.type }, t.label))) as HTMLSelectElement;
    typeSel.value = type;
    const lang = h("select", {}, ...LANGUAGES.map(([v, l]) => h("option", { value: v }, l))) as HTMLSelectElement;
    lang.value = old?.lcid ?? "";
    const legend = h("legend", {}, `Literaturverzeichnisfelder für ${styleByName(bib().style).label}`);
    const grid = h("div", { class: "srcfields" });
    const tag = h("input", { type: "text", value: old?.tag ?? "Platzhalter1", style: "width:140px" }) as HTMLInputElement;
    tag.oninput = () => { tagTouched = true; };
    const allBox = h("input", { type: "checkbox" }) as HTMLInputElement;
    const taken = bib().sources.filter((x) => x.tag !== old?.tag).map((x) => x.tag);
    const suggest = () => {
      if (tagTouched) return;
      const s = collect(false);
      if (s) tag.value = suggestTag(s, taken);
    };
    const draw = () => {
      const t = SOURCE_TYPES.find((x) => x.type === type)!;
      const fields = all ? [...t.fields, ...MORE_FIELDS.filter(([k]) => !t.fields.some(([f]) => f === k))] : t.fields;
      grid.replaceChildren();
      for (const [key, label] of fields) {
        const input = h("input", { type: "text", value: values.get(key) ?? "" }) as HTMLInputElement;
        input.oninput = () => { values.set(key, input.value); suggest(); };
        grid.append(h("label", { class: "srclabel" }, label), input);
        if (ROLES.has(key)) {
          grid.append(h("button", { onclick: () => nameDialog(label, values.get(key) ?? "", (v) => { values.set(key, v); input.value = v; suggest(); }) }, "Bearbeiten"));
          if (key === "Author") {
            const corp = h("input", { type: "checkbox", checked: corporate }) as HTMLInputElement;
            const corpName = h("input", { type: "text", disabled: !corporate, value: corporate ? values.get("Author") ?? "" : "" }) as HTMLInputElement;
            corp.onchange = () => { corporate = corp.checked; corpName.disabled = !corporate; input.disabled = corporate; };
            corpName.oninput = () => { values.set("Author", corpName.value); suggest(); };
            input.disabled = corporate;
            grid.append(h("span", {}), h("div", { class: "row" }, h("label", { class: "check" }, corp, "Firma Autor"), corpName), h("span", {}));
          }
        } else grid.append(h("span", {}));
      }
    };
    const collect = (final = true): Source | null => {
      const t = SOURCE_TYPES.find((x) => x.type === type)!;
      const keys = new Set([...t.fields.map(([k]) => k), ...(all ? MORE_FIELDS.map(([k]) => k) : []), ...(old ? Object.keys(old.fields) : [])]);
      const fields: Record<string, string> = {};
      const contributors: Record<string, Contributors> = {};
      for (const k of keys) {
        const v = (values.get(k) ?? "").trim();
        if (!v) continue;
        if (ROLES.has(k)) contributors[k] = k === "Author" && corporate ? { corporate: v } : { people: parsePeople(v) };
        else fields[k] = v;
      }
      const name = tag.value.trim();
      if (final && !name) return null;
      return { tag: name || "Platzhalter1", type, guid: old?.guid, lcid: lang.value || undefined, contributors, fields };
    };
    typeSel.onchange = () => { type = typeSel.value as SourceType; draw(); };
    allBox.onchange = () => { all = allBox.checked; draw(); };
    const body = h("div", { class: "srcdialog" },
      h("div", { class: "row two" }, h("label", { class: "field row" }, "Quellentyp", typeSel), h("label", { class: "field row" }, "Sprache", lang)),
      h("fieldset", {}, legend, grid),
      h("label", { class: "check" }, allBox, "Alle Literaturverzeichnisfelder anzeigen"),
      h("label", { class: "field" }, "Tagname", tag));
    openModal(old ? "Quelle bearbeiten" : "Quelle erstellen", body, [
      { label: "OK", primary: true, run: () => { const s = collect(); if (!s) return false; resolve(s); return true; } },
      { label: "Abbrechen", run: () => { resolve(null); return true; } },
    ], { width: 820, cls: "srcdlg" });
    draw();
  });

  /** Name bearbeiten: Nachname, Vorname, Zweiter Vorname and the list of names. */
  const nameDialog = (role: string, value: string, done: (v: string) => void) => {
    let names = parsePeople(value);
    const last = h("input", { type: "text" }) as HTMLInputElement;
    const first = h("input", { type: "text" }) as HTMLInputElement;
    const middle = h("input", { type: "text" }) as HTMLInputElement;
    const list = h("div", { class: "lb srcnames", tabIndex: 0 });
    let sel = -1;
    const draw = () => list.replaceChildren(...names.map((p, i) => h("div", { class: `lbi${i === sel ? " on" : ""}`, onmousedown: () => { sel = i; draw(); } },
      [p.last, [p.first, p.middle].filter(Boolean).join(" ")].filter(Boolean).join(", "))));
    const add = () => { if (!last.value.trim()) return; names = [...names, { last: last.value.trim(), first: first.value.trim() || undefined, middle: middle.value.trim() || undefined }]; last.value = first.value = middle.value = ""; draw(); };
    const body = h("div", { class: "namedialog" },
      h("div", {}, `${role} hinzufügen`),
      h("div", { class: "grid2" }, "Nachname:", last, "Vorname:", first, "Zweiter Vorname:", middle),
      h("div", { class: "row" }, h("button", { onclick: add }, "Hinzufügen")),
      h("div", {}, "Namen"), list,
      h("div", { class: "row" },
        h("button", { onclick: () => { if (sel > 0) { [names[sel - 1], names[sel]] = [names[sel], names[sel - 1]]; sel--; draw(); } } }, "Nach oben"),
        h("button", { onclick: () => { if (sel >= 0 && sel < names.length - 1) { [names[sel + 1], names[sel]] = [names[sel], names[sel + 1]]; sel++; draw(); } } }, "Nach unten"),
        h("button", { onclick: () => { if (sel >= 0) { const p = names[sel]; last.value = p.last; first.value = p.first ?? ""; middle.value = p.middle ?? ""; names = names.filter((_, i) => i !== sel); sel = -1; draw(); } } }, "Bearbeiten"),
        h("button", { onclick: () => { if (sel >= 0) { names = names.filter((_, i) => i !== sel); sel = -1; draw(); } } }, "Löschen")));
    openModal("Name bearbeiten", body, [
      { label: "OK", primary: true, run: () => { if (last.value.trim()) add(); done(names.map((p) => [p.last, [p.first, p.middle].filter(Boolean).join(" ")].filter(Boolean).join(", ")).join("; ")); return true; } },
      { label: "Abbrechen", run: () => true },
    ], { width: 460 });
    draw();
  };

  // ----- inserting --------------------------------------------------------------------------------
  const addSource = async () => {
    const s = await sourceDialog();
    if (!s) return;
    const b = bib();
    setBib({ ...b, sources: [...b.sources.filter((x) => x.tag !== s.tag), s] });
    remember(s);
    ed()?.insertCitation(s.tag);
    api.refresh();
    api.focus();
  };

  const addPlaceholder = () => {
    const b = bib();
    let n = 1;
    while (b.sources.some((x) => x.tag === `Platzhalter${n}`)) n++;
    const name = h("input", { type: "text", value: `Platzhalter${n}`, style: "width:260px" }) as HTMLInputElement;
    openModal("Platzhaltername", h("label", { class: "field" }, "Geben Sie den Tagnamen der Quelle ein. Sie können später weitere Informationen zu dieser Quelle hinzufügen, indem Sie auf \"Quelle bearbeiten\" klicken.", name), [
      { label: "OK", primary: true, run: () => {
        const tag = name.value.trim().replace(/\s+/g, "");
        if (!tag) return false;
        if (!b.sources.some((x) => x.tag === tag)) setBib({ ...b, sources: [...b.sources, { tag, type: "Book", fields: {}, placeholder: true }] });
        ed()?.insertCitation(tag);
        api.refresh();
        api.focus();
        return true;
      } },
      { label: "Abbrechen", run: () => true },
    ], { width: 420 });
  };

  const citationMenu = (): MenuEntry[] => [
    ...bib().sources.map((s) => ({ label: describe(s), run: () => { ed()?.insertCitation(s.tag); api.refresh(); api.focus(); } })),
    ...(bib().sources.length ? ["-" as const] : []),
    { label: "Neue Quelle hinzufügen...", icon: "citation", run: () => void addSource() },
    { label: "Neuen Platzhalter hinzufügen...", icon: "placeholder-source", run: addPlaceholder },
  ];

  // ----- Quellen-Manager -----------------------------------------------------------------------------
  const managerDialog = () => {
    const e = ed();
    if (!e) return;
    let picked: { list: "master" | "current"; tag: string } | null = null;
    const search = h("input", { type: "text" }) as HTMLInputElement;
    const sort = h("select", {}, ...[["author", "Nach Autor sortieren"], ["tag", "Nach Tag sortieren"], ["title", "Nach Titel sortieren"], ["year", "Nach Jahr sortieren"]]
      .map(([v, l]) => h("option", { value: v }, l))) as HTMLSelectElement;
    const masterBox = h("div", { class: "lb srclist", tabIndex: 0 });
    const currentBox = h("div", { class: "lb srclist", tabIndex: 0 });
    const preview = h("div", { class: "srcpreview" });
    const previewHead = h("div", {}, `Vorschau (${styleByName(bib().style).label}):`);
    const copyBtn = h("button", {}, "Kopieren ->") as HTMLButtonElement;
    const delBtn = h("button", {}, "Löschen") as HTMLButtonElement;
    const editBtn = h("button", {}, "Bearbeiten...") as HTMLButtonElement;
    const newBtn = h("button", {}, "Neu...") as HTMLButtonElement;
    const order = (list: Source[]) => {
      const q = search.value.trim().toLowerCase();
      const key = (s: Source) => (sort.value === "tag" ? s.tag : sort.value === "title" ? s.fields.Title ?? "" : sort.value === "year" ? s.fields.Year ?? "" : firstAuthor(s)).toLowerCase();
      return list.filter((s) => !q || describe(s).toLowerCase().includes(q) || s.tag.toLowerCase().includes(q)).sort((a, b) => key(a).localeCompare(key(b), "de"));
    };
    const draw = () => {
      const cited = e.citedTags();
      const row = (s: Source, list: "master" | "current") => h("div", {
        class: `lbi${picked?.list === list && picked.tag === s.tag ? " on" : ""}`,
        onmousedown: () => { picked = { list, tag: s.tag }; draw(); },
        ondblclick: () => void edit(),
      }, list === "current" ? h("span", { class: "srcmark" }, s.placeholder ? "?" : cited.includes(s.tag) ? "✓" : "") : "", describe(s));
      masterBox.replaceChildren(...order(master()).map((s) => row(s, "master")));
      currentBox.replaceChildren(...order([...bib().sources]).map((s) => row(s, "current")));
      const s = picked && (picked.list === "master" ? master() : bib().sources).find((x) => x.tag === picked!.tag);
      copyBtn.textContent = picked?.list === "current" ? "<- Kopieren" : "Kopieren ->";
      copyBtn.disabled = delBtn.disabled = editBtn.disabled = !s;
      preview.replaceChildren(...(s ? [
        h("div", {}, h("b", {}, "Zitat: "), formatCitation(s, s.tag, bib().style, 1)),
        h("div", {}, h("b", {}, "Literaturverzeichniseintrag:")),
        h("div", { class: "srcentry" }, ...formatEntry(s, bib().style, 1).map((p) => (p.italic ? h("i", {}, p.text) : document.createTextNode(p.text)))),
      ] : []));
    };
    const edit = async () => {
      if (!picked) return;
      const from = picked.list === "master" ? master() : [...bib().sources];
      const s = from.find((x) => x.tag === picked!.tag);
      if (!s) return;
      const next = await sourceDialog(s);
      if (!next) return;
      if (picked.list === "master") saveMaster(master().map((x) => (x.tag === s.tag ? next : x)));
      else { setBib({ ...bib(), sources: bib().sources.map((x) => (x.tag === s.tag ? next : x)) }); remember(next); }
      picked = { list: picked.list, tag: next.tag };
      draw();
    };
    copyBtn.onclick = () => {
      if (!picked) return;
      if (picked.list === "master") {
        const s = master().find((x) => x.tag === picked!.tag);
        if (s && !bib().sources.some((x) => x.tag === s.tag)) setBib({ ...bib(), sources: [...bib().sources, s] });
      } else {
        const s = bib().sources.find((x) => x.tag === picked!.tag);
        if (s && !s.placeholder) remember(s);
      }
      draw();
    };
    delBtn.onclick = () => {
      if (!picked) return;
      if (picked.list === "master") saveMaster(master().filter((x) => x.tag !== picked!.tag));
      else if (!e.citedTags().includes(picked.tag)) setBib({ ...bib(), sources: bib().sources.filter((x) => x.tag !== picked!.tag) });
      picked = null;
      draw();
    };
    editBtn.onclick = () => void edit();
    newBtn.onclick = async () => {
      const s = await sourceDialog();
      if (!s) return;
      setBib({ ...bib(), sources: [...bib().sources.filter((x) => x.tag !== s.tag), s] });
      remember(s);
      picked = { list: "current", tag: s.tag };
      draw();
    };
    search.oninput = draw;
    sort.onchange = draw;
    const body = h("div", { class: "srcmanager" },
      h("div", { class: "row" }, h("label", { class: "field row" }, "Suchen:", search), sort),
      h("div", { class: "srccols" },
        h("div", {}, h("div", { class: "row" }, h("div", {}, h("div", {}, "Quellen verfügbar in:"), h("div", {}, "Masterliste")),
          h("button", { disabled: true, "data-tip": "Durchsuchen...\nRaabeWrite führt eine eigene Masterliste für diesen Benutzer." }, "Durchsuchen...")), masterBox),
        h("div", { class: "srcmid" }, copyBtn, delBtn, editBtn, newBtn),
        h("div", {}, h("div", {}, "Aktuelle Liste"), currentBox,
          h("div", { class: "srclegend" }, h("div", {}, "✓  Zitierte Quelle"), h("div", {}, "?  Platzhalterquelle")))),
      previewHead, preview);
    openModal("Quellen-Manager", body, [{ label: "Schließen", primary: true, run: () => { api.focus(); return true; } }], { width: 900, cls: "srcmgr" });
    draw();
  };

  // ----- Literaturverzeichnis gallery ---------------------------------------------------------------
  const SAMPLE = [
    [{ text: "Muster, M. (2023). " }, { text: "Photosynthese verstehen.", italic: true }, { text: " Hannover: Schulbuchverlag." }],
    [{ text: "Beispiel, E. (2021). " }, { text: "Zellen und Gewebe.", italic: true }, { text: " Berlin: Lehrmittel." }],
    [{ text: "Probe, K. (2019). " }, { text: "Ökologie im Schulgarten.", italic: true }, { text: " Hamburg: Naturverlag." }],
  ];
  const bibliographyPanel = (close: () => void) => {
    const entry = (name: string, heading: string) => h("button", { class: "mbuilt bibbuilt", onclick: () => { close(); ed()?.insertBibliography(heading); api.refresh(); api.focus(); } },
      h("div", { class: "mname" }, name),
      h("div", { class: "bibprev" }, h("div", { class: "bibhead" }, heading), ...SAMPLE.map((pieces) => h("div", { class: "bibline" }, ...pieces.map((p) => (p.italic ? h("i", {}, p.text) : p.text))))));
    return h("div", { class: "rpanel mathpanel bibgallery" },
      h("div", { class: "title" }, "Integriert"),
      entry("Literaturverzeichnis", "Literaturverzeichnis"), entry("Literaturverzeichnis 2", "Literaturverzeichnis"), entry("Verweise", "Verweise"),
      h("hr"),
      h("button", { class: "link", onclick: () => { close(); ed()?.insertBibliography(""); api.refresh(); api.focus(); } }, icon("bibliography"), " Literaturverzeichnis einfügen"),
      h("button", { class: "link", disabled: true, "data-tip": "Auswahl im Literaturverzeichniskatalog speichern...\nIn RaabeWrite nicht verfügbar." }, "Auswahl im Literaturverzeichniskatalog speichern..."));
  };

  c.implement("ref.citation", { run: () => {}, enabled: () => !!ed() });
  c.implement("ref.sources", { run: managerDialog, enabled: () => !!ed() });
  c.implement("ref.citationStyle", { run: () => {}, enabled: () => !!ed() });
  c.implement("ref.bibliography", { run: () => {}, enabled: () => !!ed() });
  return {
    citationMenu, bibliographyPanel, managerDialog,
    style: {
      value: () => styleByName(bib().style).label,
      options: () => [{ items: CITATION_STYLES.map((s) => s.label) }],
      commit: (v: string) => { const s = CITATION_STYLES.find((x) => x.label === v); if (s) setBib({ ...bib(), style: s.name }); },
    },
  };
}
