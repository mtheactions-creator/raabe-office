"""PowerPoint's inventory against RaabePresent's (Step 8): docs/POWERPOINT-INVENTORY.md with, per tab and
group, every control PowerPoint has (with its keytip), the entries of its menus and galleries, the
contextual tabs and when they appear, the context menus, the File backstage, the dialogs with their
control labels, and whether RaabePresent has each control: vorhanden, zurückgestellt (greyed out as
deferred), fehlt, or Cloud (left out on purpose). Labels only (D-058).

    .venv/Scripts/python.exe inventory_report_present.py [powerpoint-inventory.json] [present-inventory.json]

PowerPoint's side comes from probe_powerpoint_inventory.py (the sanitized results file). RaabePresent's
side is read from the newest results/*-present-inventory.json when there is one, in the format of
results/*-raabe-inventory.json (a list of tabs with groups, controls, launchers and menu items, as
inventory_calc.py writes it for RaabeCalc); optionally a dict {"tabs": [...], "context_menus":
{name: [{"name": ...}]}}. Without it every control counts as "fehlt". Rerunnable: the file is
rewritten from the results each time.
"""
from __future__ import annotations

import json
import re
import sys
from pathlib import Path

from inventory_report import entries, flatten_raabe, norm, status

HERE = Path(__file__).parent
ROOT = HERE.parent.parent
OUT = ROOT / "docs" / "POWERPOINT-INVENTORY.md"

ENTRY_TYPES = ("MenuItem", "ListItem", "Button", "SplitButton", "CheckBox", "RadioButton", "ComboBox", "Edit", "Spinner", "Slider",
               "TabItem", "TreeItem", "Hyperlink", "DataItem")
TYPE_DE = {"Button": "Schaltfläche", "SplitButton": "Unterteilte Schaltfläche", "MenuItem": "Menü", "CheckBox": "Kontrollkästchen",
           "ComboBox": "Kombinationsfeld", "Edit": "Eingabefeld", "List": "Katalog", "RadioButton": "Optionsfeld", "Spinner": "Drehfeld"}


def latest(pattern: str) -> Path | None:
    files = sorted((HERE / "results").glob(pattern))
    return files[-1] if files else None


def clean(s: str) -> str:
    """One line, no control or private-use characters, safe in a Markdown table."""
    s = "".join(c for c in s if c >= " " and not 0xE000 <= ord(c) <= 0xF8FF)
    return re.sub(r"\s+", " ", s).strip().replace("|", "\\|")


def menu_text(items: list[dict]) -> str:
    """A menu's entries in order; group headings in italics, submenus in brackets."""
    parts: list[str] = []
    seen: set[tuple] = set()
    for it in items:
        name = clean(it["name"])
        if not name or it["type"] == "Menu":
            continue
        if it["type"] == "Group":
            parts.append(f"_{name}_")
            continue
        if it["type"] not in ENTRY_TYPES:
            continue
        key = (name, it["type"])
        if key in seen:
            continue
        seen.add(key)
        sub = [clean(s["name"]) for s in it.get("submenu", []) or [] if s["type"] in ENTRY_TYPES and clean(s["name"])]
        sub = list(dict.fromkeys(sub))
        parts.append(f"{name} [{' · '.join(sub)}]" if sub else name)
    return " · ".join(parts)


def count_entries(items: list[dict]) -> int:
    return len({(it["name"], it["type"]) for it in items if it["type"] in ENTRY_TYPES and it["name"].strip()}) + \
        sum(len({s["name"] for s in it.get("submenu", []) or [] if s["type"] in ENTRY_TYPES}) for it in items)


def load_present(arg: str | None):
    path = Path(arg) if arg else latest("*-present-inventory.json")
    if path is None or not path.exists():
        return {}, {}, None
    data = json.loads(path.read_text(encoding="utf-8"))
    tabs = data["tabs"] if isinstance(data, dict) else data
    cms = data.get("context_menus", {}) if isinstance(data, dict) else {}
    return flatten_raabe(tabs), {k: {norm(i["name"]) for i in v} for k, v in cms.items()}, path.name


def main() -> None:
    src = Path(sys.argv[1]) if len(sys.argv) > 1 else latest("*-powerpoint-inventory.json")
    if src is None:
        raise SystemExit("no results/*-powerpoint-inventory.json")
    pp = json.loads(src.read_text(encoding="utf-8"))
    raabe, raabe_cm, present_name = load_present(sys.argv[2] if len(sys.argv) > 2 else None)
    counts = {"vorhanden": 0, "zurückgestellt": 0, "fehlt": 0, "Cloud": 0}
    n_controls = n_entries = 0
    body: list[str] = []

    # The tabs and when they appear.
    contexts = pp.get("contexts", {})
    base = [t["name"] if isinstance(t, dict) else t for t in contexts.get("Folie", {}).get("tabs", [])]
    keys = {t["name"]: t.get("key", "") for c in contexts.values() if isinstance(c, dict) for t in c.get("tabs", []) if isinstance(t, dict)}
    body += ["\n## Registerkarten\n", "| Registerkarte | Zugriffstaste | erscheint bei |", "|---|---|---|"]
    for t in base:
        body.append(f"| {clean(t)} | {clean(keys.get(t, ''))} | immer |")
    own: dict[str, list[str]] = {}
    for ctx, c in contexts.items():
        for t in c.get("own", []) if isinstance(c, dict) else []:
            own.setdefault(t, []).append(ctx)
    for t, ctxs in own.items():
        body.append(f"| {clean(t)} | {clean(keys.get(t, ''))} | {', '.join(ctxs)} |")
    failed = {ctx: c["error"] for ctx, c in contexts.items() if isinstance(c, dict) and c.get("error")}
    if failed:
        body.append("\nNicht aufgenommen: " + "; ".join(f"{k} ({clean(v)})" for k, v in failed.items()))

    # Each tab: controls, then the menus' entries.
    walked = {}
    for key, tab in pp.get("tabs", {}).items():
        ctx, name = key.split("/", 1)
        if tab.get("error") and not tab.get("controls"):
            if not any(k.split("/", 1)[1] == name and not v.get("error") for k, v in pp["tabs"].items()):
                walked.setdefault(name, (ctx, tab))
            continue
        walked[name] = (ctx, tab)
    for name, (ctx, tab) in walked.items():
        mine = raabe.get(name, {})
        body.append(f"\n## {clean(name)}" + (f" (bei {ctx})" if ctx != "Folie" else "") + "\n")
        if tab.get("error"):
            body.append(f"Nicht vollständig aufgenommen: {clean(tab['error'])}\n")
        body += ["| Gruppe | Steuerelement | Art | Zugriffstaste | RaabePresent | Einträge |", "|---|---|---|---|---|---|"]
        menus: list[str] = []
        for c in tab.get("controls", []):
            n_controls += 1
            st = "Cloud" if c.get("cloud") else status(mine.get(norm(c["name"])))
            counts[st] += 1
            items = c.get("items") or []
            gallery = list(dict.fromkeys(c.get("gallery") or []))
            n = count_entries(items)
            n_entries += n + len(gallery)
            extra = f"{n}" if items else ("Dialog" if c.get("dialog") else (f"{len(gallery)}" if gallery else ""))
            if items and not c.get("cloud"):
                have = entries(mine.get(norm(c["name"]), {}).get("items"))
                for k in {norm(i["name"]) for i in items if i["type"] in ENTRY_TYPES and i["name"].strip()}:
                    counts["vorhanden" if k in have and not have[k].get("deferred") else "fehlt"] += 1
                menus.append(f"- **{clean(c['name'])}**: {menu_text(items)}")
            elif gallery:
                have = entries(mine.get(norm(c["name"]), {}).get("items"))
                for g in gallery:
                    counts["vorhanden" if norm(g) in have and not have[norm(g)].get("deferred") else "fehlt"] += 1
                menus.append(f"- **{clean(c['name'])}** (Katalog): {' · '.join(clean(g) for g in gallery)}")
            label = TYPE_DE.get(c["type"], c["type"]) + ("" if c.get("enabled", True) else ", grau")
            body.append(f"| {clean(c.get('group', ''))} | {clean(c['name'])} | {label} | {clean(c.get('key', ''))} | {st} | {extra} |")
        if menus:
            body += ["", f"Menüs und Kataloge in {clean(name)}:", "", *menus]

    # Context menus.
    cms = pp.get("context_menus", {})
    body += ["\n## Kontextmenüs\n", "Rechtsklick auf …; die Einträge der Minisymbolleiste stehen mit in der Liste.", ""]
    for name, cm in cms.items():
        items = cm.get("items", [])
        have = raabe_cm.get(name)
        st = "fehlt" if have is None else f"{sum(1 for i in items if norm(i['name']) in have)}/{len(items)} vorhanden"
        if cm.get("error"):
            body.append(f"- **{clean(name)}**: nicht aufgenommen ({clean(cm['error'])})")
            continue
        body.append(f"- **{clean(name)}** (RaabePresent: {st}): {menu_text(items)}")
        if cm.get("mini_toolbar"):
            body.append(f"  - Minisymbolleiste: {menu_text(cm['mini_toolbar'])}")

    # Backstage.
    bs = pp.get("backstage", {})
    if bs and not bs.get("pages"):
        body += ["\n## Datei (Backstage)\n", f"Nicht aufgenommen ({clean(str(bs.get('error', 'keine Seiten gelesen')))})."]
    elif bs:
        body += ["\n## Datei (Backstage)\n", "Navigation: " + " · ".join(clean(i["name"]) for i in bs.get("navigation", [])), ""]
        for page, got in bs.get("pages", {}).items():
            if isinstance(got, list):
                body.append(f"- **{clean(page)}**: {menu_text(got)}")
            elif isinstance(got, dict) and got.get("pages"):
                body.append(f"- **{clean(page)}** ({clean(got.get('title', ''))}):")
                for sub, ctl in got["pages"].items():
                    if isinstance(ctl, list):
                        body.append(f"  - {clean(sub) or '(Hauptseite)'}: {menu_text(ctl)}")
            elif isinstance(got, dict):
                why = "Cloud" if got.get("cloud") else ("Fehler" if got.get("error") else "zeigt eigene Dateien, handelt sofort oder Cloud")
                body.append(f"- **{clean(page)}**: nicht geöffnet ({clean(str(why))})")

    # Dialogs and panes.
    dialogs = pp.get("dialogs", {})
    body += ["\n## Dialoge und Aufgabenbereiche\n", "Geöffnet über Startprogramme, Schaltflächen und Menüeinträge mit „…“, "
             "geschlossen mit Esc oder Abbrechen. Steuerelemente nach Seiten (Registerkarten des Dialogs).", ""]
    for key, d in dialogs.items():
        if d.get("file_dialog"):
            body.append(f"- **{clean(key)}**: Dateidialog (nicht gelesen)")
            continue
        if "pane" in d:
            body.append(f"- **{clean(key)}** (Bereich {clean(d['pane'])}): {menu_text(d.get('controls', []))}")
            continue
        body.append(f"- **{clean(key)}** ({clean(d.get('title', ''))})")
        for page, ctl in d.get("pages", {}).items():
            body.append(f"  - {clean(page) or '(Hauptseite)'}: {menu_text(ctl)}")

    # Keytips.
    kt = pp.get("keytips", {})
    if kt:
        body += ["\n## Zugriffstasten (Alt)\n", "Was nach Alt und den Buchstaben einer Registerkarte angezeigt wurde.", ""]
        for tab, got in kt.items():
            if isinstance(got, dict):
                shown = got.get("shown")
                body.append(f"- **{clean(tab)}** ({clean(got.get('key', ''))}): "
                            + (" · ".join(clean(s) for s in shown) if shown else clean(got.get("skipped", "nichts gelesen"))))
            else:
                body.append(f"- **{clean(tab)}**: " + " · ".join(clean(s) for s in got))

    total = sum(v for k, v in counts.items() if k != "Cloud")
    head = [
        "# PowerPoint's inventory against RaabePresent",
        "",
        "Generated by `tools/feel-lab/inventory_report_present.py` from PowerPoint's own ribbon, menus, context menus,",
        "backstage and dialogs (`probe_powerpoint_inventory.py`, real clicks in a PowerPoint of our own) and RaabePresent's",
        f"ribbon ({present_name or 'not recorded yet: every control counts as fehlt'}). Labels and keytips only, no",
        "descriptions (D-058). Cloud features are listed, marked and left out on purpose; buttons that act at once",
        "(start the show, record, print, save) and dialogs that show the owner's files were not opened.",
        "",
        f"**Recorded:** {len(walked)} tabs with {n_controls} controls and {n_entries} menu and gallery entries; "
        f"{len([c for c in cms.values() if not c.get('error')])} context menus; {len(dialogs)} dialogs and panes"
        + (f"; {len([p for p in bs.get('pages', {}).values() if isinstance(p, (list, dict)) and not (isinstance(p, dict) and (p.get('skipped') or p.get('cloud')))])} backstage pages" if bs.get("pages") else "")
        + ".",
        "",
        f"**Coverage:** {counts['vorhanden']} of {total} controls and menu entries are in RaabePresent "
        f"({100 * counts['vorhanden'] / max(1, total):.1f} %); {counts['zurückgestellt']} are there but greyed out as deferred, "
        f"{counts['fehlt']} are missing; {counts['Cloud']} cloud features are left out on purpose.",
    ]
    if pp.get("blocked"):
        head += ["", f"The recording stopped at a dialog that must not be clicked: {clean(pp['blocked'])}."]
    OUT.write_text("\n".join(head + body) + "\n", encoding="utf-8", newline="\n")
    print(counts, f"tabs {len(walked)} controls {n_controls} entries {n_entries} context menus {len(cms)} dialogs {len(dialogs)} -> {OUT}")


if __name__ == "__main__":
    main()
