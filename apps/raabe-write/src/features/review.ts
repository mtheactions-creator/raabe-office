// Überprüfen: Kommentare (cards right of the page, as Word's modern comments, and the
// Kommentare pane), Nachverfolgung (Änderungen nachverfolgen, Markup, Annehmen/Ablehnen), the
// Überarbeitungsbereich (vertical or horizontal) and Optionen zum Nachverfolgen von Änderungen.

import type { Editor, Pos } from "@raabe/text-engine";
import { h, icon, showMenu, type Command, type MenuEntry } from "@raabe/shell";
import type { AppApi } from "./api";
import type { Dock } from "./dock";
import { openModal } from "./modal";

const cmd = (id: string, label: string, icon?: string, extra: Partial<Command> = {}): Command => ({ id, label, icon, ...extra });

export const REVIEW_COMMANDS: Command[] = [
  cmd("review.reviewingPaneVertical", "Überarbeitungsbereich vertikal...", "reviewing-pane"),
  cmd("review.reviewingPaneHorizontal", "Überarbeitungsbereich horizontal...", "reviewing-pane"),
  cmd("review.deleteAllComments", "Alle Kommentare im Dokument löschen", "comment-delete"),
  cmd("review.markupMode", "Markup für Überarbeitung anzeigen", "markup",
    { tip: "Markup für Überarbeitung anzeigen\nWählen Sie aus, wie Änderungen im Dokument angezeigt werden sollen." }),
  cmd("accept.next", "Annehmen und weiter", "accept"),
  cmd("accept.this", "Diese Änderung annehmen", "accept"),
  cmd("accept.all", "Alle Änderungen annehmen", "accept"),
  cmd("accept.allStop", "Alle Änderungen annehmen und Nachverfolgung beenden", "accept"),
  cmd("reject.next", "Ablehnen und weiter", "reject"),
  cmd("reject.this", "Diese Änderung ablehnen", "reject"),
  cmd("reject.all", "Alle Änderungen ablehnen", "reject"),
  cmd("reject.allStop", "Alle Änderungen ablehnen und Nachverfolgung beenden", "reject"),
];

// Word 365's entries (measured 2026-10-14).
export const MARKUP_MODES = [
  ["simple", "Einfaches Markup"], ["all", "Markup: alle"], ["none", "Markup: keine"], ["original", "Original"],
] as const;
type Mode = (typeof MARKUP_MODES)[number][0];

/** Width of the markup area right of the pages (px at 100 %). */
const RAIL = 340;
const AVATAR_COLORS = ["#c239b3", "#0078d4", "#498205", "#ca5010", "#8764b8", "#038387", "#986f0b", "#4f6bed"];

function initialsOf(name: string): string {
  const parts = name.trim().split(/\s+/).filter(Boolean);
  if (!parts.length) return "?";
  return (parts[0][0] + (parts.length > 1 ? parts[parts.length - 1][0] : "")).toUpperCase();
}

function colorOf(author: string): string {
  let n = 0;
  for (const ch of author) n = (n * 31 + ch.charCodeAt(0)) >>> 0;
  return AVATAR_COLORS[n % AVATAR_COLORS.length];
}

/** "Vor 5 Minuten", "Gestern", or the date (as Word's comment cards). */
function when(iso: string): string {
  const t = Date.parse(iso);
  if (!Number.isFinite(t)) return "";
  const min = Math.round((Date.now() - t) / 60000);
  if (min < 1) return "Gerade eben";
  if (min < 60) return min === 1 ? "Vor 1 Minute" : `Vor ${min} Minuten`;
  const hrs = Math.round(min / 60);
  if (hrs < 24) return hrs === 1 ? "Vor 1 Stunde" : `Vor ${hrs} Stunden`;
  return new Date(t).toLocaleDateString("de-DE", { day: "numeric", month: "long", year: "numeric" });
}

interface CardState {
  /** The comment whose text is being written or changed. */
  editing: string | null;
  /** Kommentare pane (list) instead of the cards beside the page. */
  pane: boolean;
  /** Kommentare anzeigen: the cards beside the page are shown. */
  commentsShown: boolean;
}

export function reviewFeature(api: AppApi, userName: () => string, pane: HTMLElement, docArea: HTMLElement, options: () => void, dock: Dock) {
  const ed = () => api.editor();
  const c = api.commands;
  const has = () => !!ed();
  const st: CardState = { editing: null, pane: false, commentsShown: true };
  const author = (e: Editor) => {
    const name = userName().trim() || "Benutzer";
    e.author = name;
    e.authorInitials = initialsOf(name);
    return e;
  };
  const run = (id: string, fn: (e: Editor) => void, enabled = has, checked?: () => boolean) =>
    c.implement(id, { run: () => { const e = ed(); if (e) fn(author(e)); api.focus(); }, enabled, checked });
  const inMain = () => !!ed() && ed()!.doc.activeStory === "main";
  const anyComment = () => !!ed()?.comments().length;
  const anyChange = () => !!ed()?.hasRevisions();

  // ----- comments ----------------------------------------------------------------------------
  const newComment = (e: Editor) => {
    const id = e.addComment("");
    if (!id) return;
    st.editing = id;
    e.focusComment(id);
    render(e);
    focusEditor(id);
  };
  c.implement("review.newComment", { run: () => { const e = ed(); if (e) newComment(author(e)); }, enabled: inMain });
  c.implement("insert.comment", { run: () => { const e = ed(); if (e) newComment(author(e)); }, enabled: inMain });
  const current = (e: Editor) => e.activeComment();
  run("review.deleteComment", (e) => { const id = current(e); if (id) { e.deleteComment(rootOf(e, id)); e.focusComment(null); } }, () => !!ed() && !!current(ed()!));
  run("review.deleteAllComments", (e) => { for (const x of e.comments().filter((x) => !x.parent)) e.deleteComment(x.id); }, anyComment);
  const step = (e: Editor, dir: 1 | -1) => {
    const list = e.comments().filter((x) => !x.parent);
    if (!list.length) return;
    const h0 = e.sel.head;
    const at = (x: (typeof list)[number]) => x.start.p - h0.p || x.start.o - h0.o;
    const cur = current(e);
    const i = cur ? list.findIndex((x) => x.id === cur) : -1;
    const next = i >= 0 ? list[(i + dir + list.length) % list.length]
      : dir > 0 ? list.find((x) => at(x) > 0) ?? list[0] : [...list].reverse().find((x) => at(x) < 0) ?? list[list.length - 1];
    e.selectComment(next.id);
    e.focusComment(next.id);
  };
  run("review.prevComment", (e) => step(e, -1), anyComment);
  run("review.nextComment", (e) => step(e, 1), anyComment);

  // ----- tracked changes --------------------------------------------------------------------
  run("review.trackChanges", (e) => e.setTrackChanges(!e.trackChanges), has, () => !!ed()?.trackChanges);
  c.implement("review.markupMode", { run: () => {}, enabled: has });
  c.implement("review.showMarkup", { run: () => {}, enabled: has });
  run("review.accept", (e) => e.resolveChanges(true, "next"), anyChange);
  run("review.reject", (e) => e.resolveChanges(false, "next"), anyChange);
  run("accept.next", (e) => e.resolveChanges(true, "next"), anyChange);
  run("accept.this", (e) => e.resolveChanges(true), anyChange);
  run("accept.all", (e) => e.resolveChanges(true, "all"), anyChange);
  run("accept.allStop", (e) => { e.resolveChanges(true, "all"); e.setTrackChanges(false); }, has);
  run("reject.next", (e) => e.resolveChanges(false, "next"), anyChange);
  run("reject.this", (e) => e.resolveChanges(false), anyChange);
  run("reject.all", (e) => e.resolveChanges(false, "all"), anyChange);
  run("reject.allStop", (e) => { e.resolveChanges(false, "all"); e.setTrackChanges(false); }, has);
  run("review.prevChange", (e) => e.nextChange(-1), anyChange);
  run("review.nextChange", (e) => e.nextChange(1), anyChange);

  const markupValue = () => MARKUP_MODES.find(([m]) => m === (ed()?.markupMode ?? "all"))?.[1] ?? "";
  const setMarkup = (label: string) => {
    const m = MARKUP_MODES.find(([, l]) => l.toLowerCase() === label.trim().toLowerCase());
    const e = ed();
    if (m && e) { e.setMarkup(m[0] as Mode); render(e); }
    api.focus();
  };

  // ----- the cards -----------------------------------------------------------------------------
  const rootOf = (e: Editor, id: string) => {
    const all = e.comments();
    let x = all.find((y) => y.id === id);
    while (x?.parent) x = all.find((y) => y.id === x!.parent);
    return x?.id ?? id;
  };

  /** Focuses a card's text box after rendering, unless the focus moved on in the meantime. */
  const focusEditor = (id: string) => {
    const before = document.activeElement;
    requestAnimationFrame(() => {
      const t = document.querySelector<HTMLTextAreaElement>(`.ccard [data-edit="${CSS.escape(id)}"]`);
      const now = document.activeElement;
      if (!t || (now !== before && now !== document.body && !now?.closest(".ccard"))) return;
      t.focus();
      t.setSelectionRange(t.value.length, t.value.length);
    });
  };

  /** A text box that posts with Strg+Eingabe or the send button; Esc cancels. */
  const editor = (id: string, value: string, placeholder: string, post: (text: string) => void, cancel: () => void) => {
    const area = h("textarea", { class: "cinput", rows: 1, placeholder, "data-edit": id, spellcheck: true });
    area.value = value;
    const grow = () => { area.style.height = "auto"; area.style.height = `${area.scrollHeight}px`; };
    area.addEventListener("input", grow);
    requestAnimationFrame(grow);
    area.addEventListener("keydown", (ev) => {
      ev.stopPropagation();
      if (ev.key === "Enter" && ev.ctrlKey) { ev.preventDefault(); post(area.value); }
      else if (ev.key === "Escape") { ev.preventDefault(); cancel(); }
    });
    const send = h("button", { class: "csend", "data-tip": "Kommentar posten (Strg+Eingabe)", "aria-label": "Kommentar posten", onclick: () => post(area.value) }, icon("send"));
    const no = h("button", { class: "ccancel", "data-tip": "Abbrechen", "aria-label": "Abbrechen", onclick: () => cancel() }, icon("close"));
    return h("div", { class: "cedit" }, area, h("div", { class: "cbtns" }, send, no));
  };

  const card = (e: Editor, thread: ReturnType<Editor["comments"]>, inPane: boolean) => {
    const root = thread[0];
    const act = e.activeComment();
    const active = act !== null && rootOf(e, act) === root.id;
    const el = h("div", { class: `ccard${active ? " active" : ""}${root.done ? " done" : ""}`, "data-id": root.id });
    for (const x of thread) {
      const more = h("button", { class: "cmore", "aria-label": "Weitere Threadaktionen", "data-tip": "Weitere Threadaktionen" }, icon("more"));
      more.onclick = (ev) => {
        ev.stopPropagation();
        const isRoot = x.id === root.id;
        const entries: MenuEntry[] = [
          { label: "Kommentar bearbeiten", run: () => { st.editing = x.id; e.focusComment(root.id); render(e); focusEditor(x.id); } },
          { label: isRoot ? "Thread löschen" : "Kommentar löschen", run: () => { author(e).deleteComment(x.id); e.focusComment(null); render(e); api.focus(); } },
        ];
        if (isRoot) entries.push({ label: root.done ? "Thread erneut öffnen" : "Thread auflösen", run: () => { e.setCommentDone(root.id, !root.done); render(e); } });
        showMenu(c, entries, more.getBoundingClientRect());
      };
      const head = h("div", { class: "chead" },
        h("span", { class: "avatar", style: `background:${colorOf(x.author)}` }, x.initials || initialsOf(x.author)),
        h("div", { class: "cwho" }, h("div", { class: "cname" }, x.author), h("div", { class: "cdate" }, when(x.date))), more);
      const body = st.editing === x.id
        ? editor(x.id, x.text, x.id === root.id ? "Kommentar hinzufügen …" : "Antworten", (text) => {
          st.editing = null;
          if (!text.trim() && x.id === root.id && !root.text) e.deleteComment(x.id);
          else e.setCommentText(x.id, text.replace(/\s+$/, ""));
          render(e);
          api.focus();
        }, () => {
          st.editing = null;
          if (!x.text) e.deleteComment(x.id);
          render(e);
          api.focus();
        })
        : h("div", { class: "ctext" }, x.text);
      el.append(h("div", { class: "centry" }, head, body));
    }
    if (root.done) el.append(h("div", { class: "cstate" }, icon("accept"), "Gelöst"));
    if ((active || inPane) && !st.editing && !root.done) {
      const replyId = `reply:${root.id}`;
      el.append(editor(replyId, "", "Antworten", (text) => {
        if (!text.trim()) return;
        author(e).replyComment(root.id, text.replace(/\s+$/, ""));
        e.focusComment(root.id);
        render(e);
        focusEditor(replyId);
      }, () => { e.focusComment(null); render(e); api.focus(); }));
    }
    el.addEventListener("mousedown", (ev) => {
      if ((ev.target as Element).closest("textarea, button")) return;
      ev.preventDefault();
      if (e.activeComment() !== root.id) e.focusComment(root.id);
    });
    return el;
  };

  const threads = (e: Editor) => {
    const all = e.comments();
    return all.filter((x) => !x.parent).map((r) => [r, ...all.filter((x) => x.parent === r.id)]);
  };

  let lastKey = "";
  const render = (e: Editor) => {
    author(e);
    const list = threads(e);
    const show = e.markupMode === "all" || e.markupMode === "simple";
    const rail = list.length > 0 && show && !st.pane && st.commentsShown;
    e.setMarkupWidth(rail ? RAIL : 0);
    const tops = rail ? e.commentTops() : new Map<string, number>();
    const key = JSON.stringify([list.map((t) => t.map((x) => [x.id, x.text, x.done, x.date])), [...tops], st, e.activeComment(), e.markupMode, e.zoom]);
    if (key === lastKey && e.markupLayer.childElementCount === (rail ? list.length : 0)) return;
    lastKey = key;
    const layer = e.markupLayer;
    layer.replaceChildren();
    if (rail) {
      const cards = list.map((t) => card(e, t, false));
      layer.append(...cards);
      // Each card at the height of its text, pushed down when the one above is in the way.
      let bottom = 0;
      cards.forEach((el, i) => {
        const top = Math.max(tops.get(list[i][0].id) ?? 0, bottom + 8);
        el.style.top = `${top}px`;
        bottom = top + el.offsetHeight;
      });
    }
    renderPane(e, list);
    renderRev();
  };

  const renderPane = (e: Editor, list: ReturnType<typeof threads>) => {
    pane.hidden = !st.pane;
    if (!st.pane) return;
    const close = h("button", { "aria-label": "Schließen", onclick: () => togglePane(false) }, icon("close"));
    const add = h("button", { class: "cnew", onclick: () => { const x = ed(); if (x) newComment(author(x)); } }, icon("comment"), "Neu");
    const body = h("div", { class: "cpane-list" });
    if (!list.length) body.append(h("div", { class: "cempty" }, "Noch keine Kommentare.", h("br"), "Wählen Sie Text aus und klicken Sie auf „Neuer Kommentar“."));
    for (const t of list) {
      const el = card(e, t, true);
      el.addEventListener("click", () => e.selectComment(t[0].id));
      body.append(el);
    }
    pane.replaceChildren(h("div", { class: "cpane-head" }, h("span", null, "Kommentare"), h("div", null, add, close)), body);
  };

  const togglePane = (on = !st.pane) => {
    st.pane = on;
    lastKey = "";
    const e = ed();
    if (e) render(e);
    else pane.hidden = true;
  };

  // Cards follow the text one frame later, so typing never waits for them.
  let queued = false;
  const attach = (e: Editor) => {
    author(e);
    e.onChange(() => {
      if (queued || ed() !== e) return;
      queued = true;
      requestAnimationFrame(() => { queued = false; const x = ed(); if (x) render(x); });
    });
  };

  // Kommentare anzeigen: on/off; the menu chooses Word's two views (cards beside the text, or the list).
  c.implement("review.showComments", {
    run: () => { st.commentsShown = !st.commentsShown; if (!st.commentsShown) st.pane = false; lastKey = ""; const e = ed(); if (e) render(e); },
    enabled: has, checked: () => st.commentsShown || st.pane,
  });
  const showCommentsMenu: MenuEntry[] = [
    { label: "Kontextbezogen", run: () => { st.commentsShown = true; togglePane(false); }, get checked() { return st.commentsShown && !st.pane; } },
    { label: "Liste", run: () => togglePane(true), get checked() { return st.pane; } },
  ];
  // Markup anzeigen / Sämtliches Markup filtern: what the markup shows.
  let lastMarkup: Mode = "all";
  const setInsDel = (on: boolean) => {
    const e = ed();
    if (!e) return;
    if (!on && e.markupMode !== "none") { lastMarkup = e.markupMode as Mode; e.setMarkup("none"); }
    else if (on && e.markupMode === "none") e.setMarkup(lastMarkup === "none" ? "all" : lastMarkup);
    lastKey = "";
    render(e);
  };
  c.implement("review.filterMarkup", { run: () => {}, enabled: has });
  const filterMenu: MenuEntry[] = [
    { label: "Kommentare", run: () => { st.commentsShown = !st.commentsShown; lastKey = ""; const e = ed(); if (e) render(e); }, get checked() { return st.commentsShown; } },
    { label: "Einfügungen und Löschungen", run: () => setInsDel(ed()?.markupMode === "none"), get checked() { return ed()?.markupMode !== "none"; } },
    { label: "Formatierung", run: () => {}, disabled: true },
  ];

  // ----- Überarbeitungsbereich ----------------------------------------------------------------
  // Word's pane (measured 2026-10-14): "Überarbeitungen", the count with a detailed summary,
  // then each change as "<Autor> Hinzugefügt" / "hat gelöscht" / "hat kommentiert" and its text.
  let revMode: "off" | "vertical" | "horizontal" = "off";
  let revDetails = false;
  const revPane = h("div", { class: "revpane", hidden: true, role: "complementary", "aria-label": "Überarbeitungen" });
  docArea.append(revPane);
  const setRevPane = (mode: typeof revMode) => {
    revMode = mode;
    revPane.hidden = mode === "off";
    revPane.classList.toggle("horizontal", mode === "horizontal");
    dock.add(revPane, mode === "horizontal" ? "bottom" : "left", mode === "horizontal" ? 220 : 270);
    revKey = "";
    renderRev();
    api.refresh();
    if (mode === "off") api.focus();
  };
  let revKey = "";
  const renderRev = () => {
    if (revMode === "off") return;
    const e = ed();
    type Item = { at: Pos; who: string; what: string; text: string; go: () => void };
    const items: Item[] = [];
    const count = { ins: 0, del: 0, comments: 0 };
    if (e) {
      for (const r of e.revisionList()) {
        count[r.kind]++;
        items.push({ at: r.from, who: r.author || "Unbekannter Autor", what: r.kind === "ins" ? "Hinzugefügt" : "hat gelöscht", text: r.text, go: () => e.select(r.from, r.to) });
      }
      for (const t of threads(e)) {
        count.comments++;
        items.push({ at: t[0].start, who: t[0].author, what: "hat kommentiert", text: t[0].text, go: () => e.selectComment(t[0].id) });
      }
      items.sort((a, b) => a.at.p - b.at.p || a.at.o - b.at.o);
    }
    const key = JSON.stringify([revMode, revDetails, items.map((x) => [x.at, x.who, x.what, x.text])]);
    if (key === revKey) return;
    revKey = key;
    const n = items.length;
    const head = h("div", { class: "rev-head" }, h("span", null, "Überarbeitungen"),
      h("div", null,
        h("button", { "aria-label": "Optionen für den Aufgabenbereich", "data-tip": "Optionen für den Aufgabenbereich",
          onclick: (ev: MouseEvent) => showMenu(c, [{ label: "Schließen", run: () => setRevPane("off") }], (ev.currentTarget as HTMLElement).getBoundingClientRect()) }, icon("chevron-down")),
        h("button", { "aria-label": "Bereich schließen", "data-tip": "Bereich schließen", onclick: () => setRevPane("off") }, icon("close"))));
    const summary = h("div", { class: "rev-count" },
      h("button", { class: "rev-toggle", "aria-expanded": String(revDetails), "data-tip": revDetails ? "Detaillierte Zusammenfassung ausblenden" : "Detaillierte Zusammenfassung anzeigen",
        onclick: () => { revDetails = !revDetails; revKey = ""; renderRev(); } }, icon(revDetails ? "chevron-down" : "chevron-up")),
      h("b", null, n === 1 ? "1 Überarbeitung" : `${n} Überarbeitungen`),
      h("button", { class: "rev-refresh", "aria-label": "Aktualisieren", "data-tip": "Aktualisieren", onclick: () => { revKey = ""; renderRev(); } }, icon("update")));
    const details = revDetails ? h("div", { class: "rev-details" },
      ...([["Einfügungen", count.ins], ["Löschungen", count.del], ["Verschiebungen", 0], ["Formatierung", 0], ["Kommentare", count.comments]] as const)
        .map(([k, v]) => h("div", null, `${k}: ${v}`))) : null;
    const list = h("div", { class: "rev-list" });
    for (const x of items) {
      const el = h("button", { class: "rev-item" }, h("div", { class: "rev-who" }, `${x.who} ${x.what}`), h("div", { class: `rev-text${x.what === "hat gelöscht" ? " del" : ""}` }, x.text.replace(/\n$/, "").replace(/[\x00-\x08\x0b-\x1f]/g, " ") || " "));
      el.onclick = () => { x.go(); api.focus(); };
      list.append(el);
    }
    revPane.replaceChildren(head, summary, ...(details ? [details] : []), list);
  };
  c.implement("review.reviewingPane", { run: () => setRevPane(revMode === "off" ? "vertical" : "off"), enabled: has, checked: () => revMode !== "off" });
  c.implement("review.reviewingPaneVertical", { run: () => setRevPane("vertical"), enabled: has, checked: () => revMode === "vertical" });
  c.implement("review.reviewingPaneHorizontal", { run: () => setRevPane("horizontal"), enabled: has, checked: () => revMode === "horizontal" });
  const reviewingPaneMenu: MenuEntry[] = [{ cmd: "review.reviewingPaneVertical" }, { cmd: "review.reviewingPaneHorizontal" }];

  // ----- Optionen zum Nachverfolgen von Änderungen -----------------------------------------------
  // Word's dialog (recorded 2026-10-14). What RaabeWrite does not track (ink, formatting) is shown
  // as Word shows it, greyed where it cannot change anything.
  const trackOptions = () => {
    const e = ed();
    if (!e) return;
    const check = (label: string, checked: boolean, disabled = false, tip?: string) => {
      const box = h("input", { type: "checkbox", checked, disabled }) as HTMLInputElement;
      return { box, el: h("label", { class: `check${disabled ? " disabled" : ""}`, ...(tip ? { "data-tip": tip } : {}) }, box, label) };
    };
    const ink = check("Freihand", true, true, "RaabeWrite hat keine Freihandeingaben.");
    const insDel = check("Einfügungen und Löschungen", e.markupMode !== "none");
    const format = check("Formatierung", true, true, "RaabeWrite verfolgt Formatierungsänderungen nicht nach.");
    const updates = check("Aktualisierungen hervorheben", true, true);
    const others = check("Andere Autoren", true, true);
    const balloons = h("select", {},
      h("option", { value: "revisions", disabled: true }, "Überarbeitungen"),
      h("option", { value: "comments" }, "Kommentare und Formatierungen"),
      h("option", { value: "none" }, "Nichts")) as HTMLSelectElement;
    balloons.value = st.commentsShown || st.pane ? "comments" : "none";
    const paneSel = h("select", {}, h("option", { value: "off" }, "Aus"), h("option", { value: "vertical" }, "Vertikal"), h("option", { value: "horizontal" }, "Horizontal")) as HTMLSelectElement;
    paneSel.value = revMode;
    const body = h("div", { class: "trackopts" },
      h("fieldset", {}, h("legend", {}, "Anzeigen"),
        h("div", { class: "grid2" }, ink.el, updates.el, insDel.el, others.el, format.el)),
      h("div", { class: "grid2 wide" },
        h("span", null, "In Sprechblasen in der Ansicht \"Markup: alle\" anzeigen:"), balloons,
        h("span", null, "Überarbeitungsbereich:"), paneSel),
      h("div", { class: "row" },
        h("button", { disabled: true, "data-tip": "Erweiterte Optionen\nFarben und Markierungen der Änderungen sind in RaabeWrite fest." }, "Erweiterte Optionen..."),
        h("button", { onclick: () => { m.close(); options(); } }, "Benutzernamen ändern...")));
    const m = openModal("Optionen zum Nachverfolgen von Änderungen", body, [
      { label: "OK", primary: true, run: () => {
        if (insDel.box.checked !== (e.markupMode !== "none")) setInsDel(insDel.box.checked);
        const show = balloons.value === "comments";
        if (show !== (st.commentsShown || st.pane)) { st.commentsShown = show; if (!show) st.pane = false; }
        lastKey = "";
        render(e);
        if (paneSel.value !== revMode) setRevPane(paneSel.value as typeof revMode);
        return true;
      } },
      { label: "Abbrechen", run: () => true },
    ], { width: 620 });
  };
  c.implement("review.trackOptions", { run: trackOptions, enabled: has });

  const trackMenu: MenuEntry[] = [{ cmd: "review.trackChanges" }];
  const deleteMenu: MenuEntry[] = [{ cmd: "review.deleteComment" }, { cmd: "review.deleteAllComments" }];
  const acceptMenu: MenuEntry[] = [{ cmd: "accept.next" }, { cmd: "accept.this" }, { cmd: "accept.all" }, { cmd: "accept.allStop" }];
  const rejectMenu: MenuEntry[] = [{ cmd: "reject.next" }, { cmd: "reject.this" }, { cmd: "reject.all" }, { cmd: "reject.allStop" }];
  return {
    attach, render: () => { const e = ed(); if (e) { lastKey = ""; render(e); } }, togglePane,
    trackMenu, deleteMenu, acceptMenu, rejectMenu, showCommentsMenu, filterMenu, reviewingPaneMenu,
    markup: { value: markupValue, set: setMarkup, options: () => [{ items: MARKUP_MODES.map(([, l]) => l) }] },
  };
}
