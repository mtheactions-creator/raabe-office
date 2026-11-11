"""The sweep "pictures where Office shows pictures" (Step 11, D-163). The owner: RaabeOffice used text where Office uses
symbols (the arrow types of the chart pane). This check compares, for every ribbon dropdown, gallery and submenu that
Office shows, whether its entries are pictures or text in Office and in ours, and fails where Office shows pictures
and we show text (unless the difference is written down in KNOWN with its reason).

Office's side comes from the recorded inventories (our own Office instances, UI Automation; labels and control types
only, D-058): results/2026-11-05-excel-inventory.json and results/2026-11-06-powerpoint-inventory.json. UI Automation
names a gallery's entries ListItem and a menu's entries MenuItem; a dropdown whose entries are mostly ListItems is a
gallery. Galleries whose entries Office draws as text (Zeilenabstand "1,0", the function lists of Mehr Funktionen) are
listed in TEXT_GALLERIES, looked at by eye in Office; everything else in a gallery is a picture (often with its name).
The Format panes are checked by unit tests against their own recordings (packages/chart-ui/test/panepictures.test.ts).

Our side: the apps' own inventories (`__calc.inventory()`, `__present.inventory()`) in a headless Edge of our own on
the development servers (no input, invisible); every entry says how it shows itself (`entryKind`: picture, text or
both).

    .venv/Scripts/python.exe check_picture_controls.py [--calc URL] [--present URL]

Defaults: http://localhost:5175/ (npm run dev:calc) and http://localhost:5176/ (npm run dev:present).
Result: results/2026-11-11-picture-controls.json; exit code 1 when a new difference is found.
"""
from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path

from feellab import present_page

HERE = Path(__file__).parent
OUT = HERE / "results" / "2026-11-11-picture-controls.json"
SKIP = {"Group", "Menu", "Separator", "Text", "ScrollBar", "Thumb", "Image"}

# Galleries whose entries Office shows as text: names (function lists, font names, "1,0") or a small icon beside a
# name that carries the meaning (Hochformat, Fenster fixieren, Textrichtung). Looked at in Excel / PowerPoint 365.
TEXT_GALLERIES = {
    *(f"excel|Formeln|{x}" for x in ("Mehr Funktionen > Statistik", "Mehr Funktionen > Technik", "Mehr Funktionen > Cube", "Mehr Funktionen > Informationen",
                                     "Mehr Funktionen > Kompatibilität", "Mehr Funktionen > Web", "Datum u. Uhrzeit", "Finanzmathematik", "Logisch",
                                     "Mathematik und Trigonometrie", "Nachschlagen und Verweisen", "Text", "Zuletzt verwendet")),
    "excel|Einfügen|Illustrationen > Screenshot", "excel|Ansicht|Fenster fixieren",
    "excel|Seitenlayout|Ausrichtung", "excel|Seitenlayout|Format", "excel|Seitenlayout|Seitenränder", "excel|Seitenlayout|Schriftarten",
    *(f"excel|Diagrammentwurf|Diagrammelement hinzufügen > {x}" for x in ("Achsen", "Achsentitel", "Datenbeschriftungen", "Datentabelle", "Diagrammtitel",
                                                                         "Fehlerindikatoren", "Gitternetzlinien", "Legende", "Linien", "Trendlinie", "Positive/Negative Abweichung")),
    *(f"powerpoint|Diagrammentwurf|Diagrammelement hinzufügen > {x}" for x in ("Achsen", "Achsentitel", "Datenbeschriftungen", "Datentabelle", "Diagrammtitel",
                                                                              "Fehlerindikatoren", "Gitternetzlinien", "Legende", "Linien", "Trendlinie", "Positive/Negative Abweichung")),
    "powerpoint|Start|Zeilenabstand", "powerpoint|Start|Groß-/Kleinschreibung", "powerpoint|Start|Zeichenabstand", "powerpoint|Start|Schnellformatvorlagen > Alle",
    "powerpoint|Start|Spalten", "powerpoint|Start|Text ausrichten", "powerpoint|Start|Textrichtung", "powerpoint|Start|Anordnen > Drehen",
    "powerpoint|Formformat|Drehen", "powerpoint|Bildformat|Drehen", "powerpoint|Tabellenlayout|Textrichtung", "powerpoint|Tabellenlayout|Zellenbegrenzungen",
    "powerpoint|Übergänge|Sound", "powerpoint|Animationen|Animation hinzufügen",
    *(f"powerpoint|{t}|{x}" for t in ("Entwurf", "Folienmaster", "Handzettelmaster", "Notizenmaster") for x in ("Schriftarten", "Foliengröße")),
    "powerpoint|Handzettelmaster|Handzettelausrichtung", "powerpoint|Notizenmaster|Notizenseitenausrichtung",
}

# Differences written down (D-163, D-164): "<app>|<tab>|<path>" -> reason. Emptied by D-164; an entry that has
# become a picture fails the check too, so the list stays true.
KNOWN: dict[str, str] = {}


def norm(s: str) -> str:
    s = (s or "").replace("…", "...").strip()
    s = re.sub(r"\s*\((Strg|Alt|Umschalt)\+[^)]*\)$", "", s)
    return re.sub(r"\.\.\.$", "", s).strip().lower()


def office_kind(items: list[dict]) -> tuple[str | None, int]:
    es = [i for i in items if i.get("type") not in SKIP and (i.get("name") or "").strip() and i.get("name") != "Weitere Optionen"]
    if not es:
        return None, 0
    li = sum(1 for i in es if i.get("type") == "ListItem")
    return ("picture" if li * 2 >= len(es) else "text"), len(es)


def office_side(path: Path, app: str) -> dict[str, dict]:
    """'<app>|<tab>|<control>[ > <submenu>]' -> {kind, n, group}, merged over contexts (Folie/Start and Text/Start)."""
    d = json.loads(path.read_text(encoding="utf-8"))
    out: dict[str, dict] = {}
    for key, t in d["tabs"].items():
        tab = key.split("/", 1)[1]
        for c in t.get("controls", []):
            items = c.get("items") or []
            k, n = office_kind(items)
            if k:
                p = f"{app}|{tab}|{c['name']}"
                if p not in out or k == "picture":
                    out[p] = {"kind": k, "n": n, "group": c.get("group", ""), "context": key.split("/")[0]}
            for i in items:
                k2, n2 = office_kind(i.get("submenu") or [])
                if k2:
                    p = f"{app}|{tab}|{c['name']} > {i['name']}"
                    if p not in out or k2 == "picture":
                        out[p] = {"kind": k2, "n": n2, "group": c.get("group", ""), "context": key.split("/")[0]}
    for p, v in out.items():
        if p in TEXT_GALLERIES and v["kind"] == "picture":
            v["kind"], v["by_eye"] = "text", True
    return out


def ours_kind(items: list[dict]) -> str | None:
    es = [i for i in items if (i.get("name") or "").strip()]
    if not es:
        return None
    pic = sum(1 for i in es if i.get("kind") in ("picture", "both") or i.get("type") == "gallery")
    return "picture" if pic * 2 >= len(es) else "text"


def ours_side(tabs: list[dict], app: str) -> dict[str, dict]:
    """The same paths in ours: '<app>|<tab>|<control>[ > <entry>]' (normalised) -> {kind, n}."""
    out: dict[str, dict] = {}

    def put(p: str, k: str, n: int) -> None:
        # A tab can come twice (Formformat of a shape and of WordArt): a picture in either counts, as on Office's side.
        if p not in out or k == "picture":
            out[p] = {"kind": k, "n": n}

    for t in tabs:
        for g in t["groups"]:
            for c in g["controls"]:
                if c.get("type") == "gallery":
                    put(f"{app}|{norm(t['name'])}|{norm(g['name'])}", "picture", 1)
                    continue
                items = c.get("items") or []
                k = ours_kind(items)
                if k:
                    put(f"{app}|{norm(t['name'])}|{norm(c['name'])}", k, len(items))
                for i in items:
                    k2 = ours_kind(i.get("submenu") or [])
                    if k2:
                        put(f"{app}|{norm(t['name'])}|{norm(c['name'])} > {norm(i['name'])}", k2, len(i.get("submenu") or []))
    return out


def inventory(url: str, ready: str, expr: str, port: int):
    """An app's inventory from its dev server in a headless Edge of our own (PresentPage's way, any app and port)."""
    import subprocess
    import time
    import urllib.request

    import websocket
    profile = present_page.PROFILE.parent / f"raabe-pictures-check-{port}"
    page = present_page.PresentPage.__new__(present_page.PresentPage)
    page.proc = subprocess.Popen([present_page.EDGE, "--headless=new", "--inprivate", f"--remote-debugging-port={port}", f"--user-data-dir={profile}",
                                  "--window-size=1600,1000", "--no-first-run", "--disable-sync", url], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    page.n = 0
    page.ws = None
    try:
        for _ in range(200):
            try:
                with urllib.request.urlopen(f"http://127.0.0.1:{port}/json/list", timeout=2) as r:
                    pages = [p for p in json.load(r) if p.get("type") == "page" and p.get("url", "").startswith(url)]
                if pages:
                    page.ws = websocket.create_connection(pages[0]["webSocketDebuggerUrl"], suppress_origin=True, timeout=180)
                    if page.js(ready):
                        break
            except Exception:  # noqa: BLE001
                pass
            time.sleep(0.25)
        else:
            raise RuntimeError(f"{url} not ready")
        return page.js(expr)
    finally:
        try:
            page.ws and page.ws.close()
        except Exception:  # noqa: BLE001
            pass
        # Only the Edge started with our own profile is ended.
        subprocess.run(["powershell", "-NoProfile", "-Command",
                        "Get-CimInstance Win32_Process -Filter \"Name='msedge.exe'\" | Where-Object { $_.CommandLine -like '*" + profile.name + "*' }"
                        " | ForEach-Object { Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue }"], capture_output=True)


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--calc", default="http://localhost:5175/")
    ap.add_argument("--present", default="http://localhost:5176/")
    a = ap.parse_args()
    office = {**office_side(HERE / "results" / "2026-11-05-excel-inventory.json", "excel"),
              **office_side(HERE / "results" / "2026-11-06-powerpoint-inventory.json", "powerpoint")}
    calc = inventory(a.calc, "typeof window.__calc === 'object'", "__calc.inventory()", 9477)
    pres = inventory(a.present, "typeof window.__present === 'object'", "__present.inventory()", 9478)
    ours = {**ours_side(calc, "excel"), **ours_side(pres, "powerpoint")}
    rows = []
    for p, o in sorted(office.items()):
        app, tab, path = p.split("|", 2)
        key = f"{app}|{norm(tab)}|{norm(path)}"
        mine = ours.get(key)
        if mine is None and " > " not in path:
            # A gallery control is listed under its group in ours.
            mine = ours.get(f"{app}|{norm(tab)}|{norm(o['group'])}")
        status = "fehlt" if mine is None else "gleich" if mine["kind"] == o["kind"] else f"Office {o['kind']}, wir {mine['kind']}"
        rows.append({"office": p, "kind": o["kind"], "n": o["n"], "ours": mine["kind"] if mine else None, "status": status,
                     **({"by_eye": True} if o.get("by_eye") else {}), **({"known": KNOWN[p]} if p in KNOWN else {})})
    pics = [r for r in rows if r["kind"] == "picture"]
    wrong = [r for r in pics if r["ours"] == "text" and "known" not in r]
    # A difference written down that is fixed now: take it off KNOWN.
    wrong += [r for r in pics if "known" in r and r["ours"] != "text"]
    summary = {
        "office picture dropdowns": len(pics), "ours picture": sum(1 for r in pics if r["ours"] == "picture"),
        "ours text": sum(1 for r in pics if r["ours"] == "text"), "not built": sum(1 for r in pics if r["ours"] is None),
        "office text dropdowns": sum(1 for r in rows if r["kind"] == "text"), "known": sum(1 for r in pics if "known" in r),
        "per app": {app: {"office pictures": sum(1 for r in pics if r["office"].startswith(app)), "ours picture": sum(1 for r in pics if r["office"].startswith(app) and r["ours"] == "picture"),
                          "ours text": sum(1 for r in pics if r["office"].startswith(app) and r["ours"] == "text")} for app in ("excel", "powerpoint")},
    }
    data = {"meta": {"recorded": "2026-11-11", "office": "results/2026-11-05-excel-inventory.json, results/2026-11-06-powerpoint-inventory.json (UI Automation types)",
                     "ours": "__calc.inventory(), __present.inventory() in a headless Edge on the dev servers", "rule": __doc__.split("\n\n")[1]},
            "summary": summary, "rows": rows}
    OUT.write_text(json.dumps(data, ensure_ascii=False, indent=1) + "\n", encoding="utf-8", newline="\n")
    print(json.dumps(summary, ensure_ascii=False, indent=1))
    for r in wrong:
        print("TEXT, Office shows pictures:", r["office"])
    sys.exit(1 if wrong else 0)


if __name__ == "__main__":
    main()
