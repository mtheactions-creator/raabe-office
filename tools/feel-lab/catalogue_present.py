"""Behaviour catalogue for RaabePresent (D-070): the same real input in PowerPoint and RaabePresent.

    .venv/Scripts/python.exe catalogue_present.py powerpoint [area ...]   # records PowerPoint's outcomes
    .venv/Scripts/python.exe catalogue_present.py present [area ...]      # runs RaabePresent, compares
    .venv/Scripts/python.exe catalogue_present.py list

Takes over keyboard, mouse and screen: run only with the real-input lock (CLAUDE.md), created here
exclusively and deleted at the end. PowerPoint runs only when no POWERPNT.EXE is running (the
owner's); the recording goes to results/<date>-catalogue-present-powerpoint.json, our run and the
comparison to results/<date>-catalogue-present.json.
"""
from __future__ import annotations

import datetime
import json
import os
import sys
import tempfile
import time
from pathlib import Path

from feellab.present_catalogue import Case, PowerPointSide, PresentSide, differences, normalise

HERE = Path(__file__).parent
LOCK = Path(tempfile.gettempdir()) / "raabe-real-input.lock"

# ----- the cases --------------------------------------------------------------------------------------

RECT = {"kind": "rect", "box": [300, 150, 200, 120], "text": "Form"}
ELLIPSE = {"kind": "ellipse", "box": [300, 150, 200, 120]}
TEXTBOX = {"kind": "textbox", "box": [300, 150, 240, 40], "text": "Hallo Welt"}
LINE = {"kind": "line", "box": [300, 200, 200, 0]}
TABLE = {"kind": "table", "box": [200, 150, 400, 80]}
OTHER = {"kind": "rect", "box": [620, 330, 120, 80]}

OBJECTS = {"rect": RECT, "ellipse": ELLIPSE, "textbox": TEXTBOX}


def cases() -> list[Case]:
    out: list[Case] = []

    def add(name, area, shapes, steps, **kw):
        out.append(Case(name, area, shapes, steps, **kw))

    # Clicks on each object kind: centre, its text, the frame edge, empty slide.
    for nm, obj in OBJECTS.items():
        add(f"click {nm} centre", "click", [obj], [("click", (0, .5, .5))])
        add(f"click {nm} near left", "click", [obj], [("click", (0, .1, .5))])
        add(f"click {nm} top edge", "click", [obj], [("click", (0, .5, 0))])
        add(f"click {nm} then empty", "click", [obj], [("click", (0, .5, .5)), ("click", ("slide", 80, 480))])
        add(f"click {nm} twice", "click", [obj], [("click", (0, .5, .5)), ("wait", .6), ("click", (0, .5, .5))])
        add(f"dblclick {nm}", "dblclick", [obj], [("dblclick", (0, .5, .5))])
        add(f"rclick {nm}", "rclick", [obj], [("rclick", (0, .5, .5))])
        add(f"click {nm} shift other", "multi", [obj, OTHER], [("click", (0, .5, .5)), ("click", (1, .5, .5), "shift")])
        add(f"click {nm} ctrl other", "multi", [obj, OTHER], [("click", (0, .5, .5)), ("click", (1, .5, .5), "ctrl")])
        add(f"click {nm} shift other twice", "multi", [obj, OTHER], [("click", (0, .5, .5)), ("click", (1, .5, .5), "shift"), ("click", (1, .5, .5), "shift")])
        # Dragging the object: plain, Umschalt (one direction), Strg (copy), Alt (no snapping).
        for mods in ("", "shift", "ctrl", "alt"):
            add(f"drag {nm} {mods or 'plain'}", "drag", [obj], [("drag", (0, .5, .5), 60, 25, mods)])
        add(f"drag {nm} edge", "drag", [obj], [("click", (0, .5, .5)), ("keys", "esc"), ("drag", (0, .5, 0), 40, 40)])
        # Handles after a selection by its edge.
        for k in ("se", "e", "s", "nw", "n", "w"):
            for mods in ("", "shift", "ctrl"):
                add(f"handle {nm} {k} {mods or 'plain'}", "handle", [obj], [("click", (0, .5, 0)), ("handle", 0, k, 30, 20, mods)])
        add(f"rotate {nm}", "rotate", [obj], [("click", (0, .5, 0)), ("handle", 0, "rot", 80, 40)])
        add(f"rotate {nm} shift", "rotate", [obj], [("click", (0, .5, 0)), ("handle", 0, "rot", 80, 40, "shift")])
        # Keys with the frame selected.
        sel = [("click", (0, .5, 0))]
        for key in ("left", "right", "up", "down", "ctrl+left", "ctrl+down", "shift+right", "alt+right", "alt+left", "delete", "backspace", "ctrl+d", "tab", "shift+tab", "esc",
                    "enter", "f2", "ctrl+a", "ctrl+c", "ctrl+x", "ctrl+z", "ctrl+shift+f", "ctrl+e", "ctrl+r", "ctrl+l"):
            add(f"frame {nm} {key}", "framekeys", [obj, OTHER], sel + [("keys", key)])
        add(f"frame {nm} type", "framekeys", [obj], sel + [("type", "xy")])
        add(f"frame {nm} ctrl+d twice", "framekeys", [obj], sel + [("keys", "ctrl+d"), ("keys", "ctrl+d")])
        add(f"frame {nm} copy paste", "framekeys", [obj], sel + [("keys", "ctrl+c"), ("keys", "ctrl+v"), ("keys", "ctrl+v")])
        add(f"frame {nm} delete undo", "framekeys", [obj], sel + [("keys", "delete"), ("keys", "ctrl+z")])
        add(f"frame {nm} nudge 3", "framekeys", [obj], sel + [("keys", "right", 3)])

    # Text being edited: keys and clicks in a two-paragraph text box and the title placeholder.
    two = {"kind": "textbox", "box": [200, 150, 360, 80], "text": "Erste Zeile hier\rZweite Zeile"}
    edit = [("click", (0, .5, .5))]
    for key in ("home", "end", "ctrl+home", "ctrl+end", "left", "right", "up", "down", "shift+home", "shift+end", "ctrl+left", "ctrl+right", "ctrl+shift+right",
                "backspace", "delete", "ctrl+backspace", "ctrl+delete", "enter", "shift+enter", "tab", "shift+tab", "esc", "f2", "ctrl+a", "ctrl+shift+f", "ctrl+e",
                "ctrl+z", "alt+shift+right", "alt+shift+left", "ctrl+up", "ctrl+down", "pagedown"):
        add(f"text {key}", "textkeys", [two], edit + [("keys", key)])
        add(f"text sel {key}", "textkeys", [two], edit + [("keys", "shift+right", 3), ("keys", key)])
    add("text type", "textkeys", [two], edit + [("type", "abc")])
    add("text type over selection", "textkeys", [two], edit + [("keys", "ctrl+shift+left"), ("type", "X")])
    add("text dblclick word", "textmouse", [two], [("dblclick", (0, .2, .3))])
    add("text triple", "textmouse", [two], [("click", (0, .3, .3)), ("dblclick", (0, .3, .3))])
    add("text shift click", "textmouse", [two], edit + [("click", (0, .8, .7), "shift")])
    add("text drag select", "textmouse", [two], edit + [("drag", (0, .1, .3), 120, 0)])
    add("text click other shape", "textmouse", [two, OTHER], edit + [("click", (1, .5, .5))])
    add("text click empty", "textmouse", [two], edit + [("click", ("slide", 80, 480))])
    add("text esc esc", "textmouse", [two], edit + [("keys", "esc"), ("keys", "esc")])

    # Placeholders: empty and with text.
    for lay, ph, txt in (("Titelfolie", "ph0", "Mein Titel"), ("Titel und Inhalt", "ph1", "Punkt eins\rPunkt zwei")):
        shapes = [{"kind": "ph", "ph": int(ph[2:]), "text": txt}]
        add(f"{lay} click {ph}", "placeholder", shapes, [("click", (ph, .5, .5))], layout=lay)
        add(f"{lay} click {ph} edge", "placeholder", shapes, [("click", (ph, .5, 0))], layout=lay)
        add(f"{lay} empty click", "placeholder", [], [("click", (ph, .5, .5))], layout=lay)
        add(f"{lay} empty type", "placeholder", [], [("click", (ph, .5, .5)), ("type", "Neu")], layout=lay)
        add(f"{lay} delete text", "placeholder", shapes, [("click", (ph, .5, 0)), ("keys", "delete")], layout=lay)
        add(f"{lay} delete twice", "placeholder", shapes, [("click", (ph, .5, 0)), ("keys", "delete"), ("keys", "delete")], layout=lay)
        add(f"{lay} drag {ph}", "placeholder", shapes, [("drag", (ph, .5, 0), 40, 30)], layout=lay)
        add(f"{lay} tab level", "placeholder", shapes, [("click", (ph, .5, .5)), ("keys", "home"), ("keys", "tab")], layout=lay)
    add("content enter tab", "placeholder", [{"kind": "ph", "ph": 1, "text": "Eins"}], [("click", ("ph1", .5, .5)), ("keys", "ctrl+end"), ("keys", "enter"), ("keys", "tab"), ("type", "Unterpunkt")], layout="Titel und Inhalt")

    # Several selected: keys, Gruppieren.
    two_shapes = [RECT, OTHER]
    msel = [("click", (0, .5, 0)), ("click", (1, .5, 0), "shift")]
    for key in ("ctrl+g", "delete", "right", "ctrl+d", "esc", "tab", "ctrl+a"):
        add(f"multi {key}", "multikeys", two_shapes, msel + [("keys", key)])
    add("multi group ungroup", "multikeys", two_shapes, msel + [("keys", "ctrl+g"), ("keys", "ctrl+shift+g")])
    add("multi drag", "multikeys", two_shapes, msel + [("drag", (0, .5, .5), 30, 30)])
    add("group click member", "group", two_shapes, msel + [("keys", "ctrl+g"), ("click", ("slide", 80, 480)), ("click", (0, .5, .5)), ("click", (0, .5, .5))])
    add("marquee", "marquee", two_shapes, [("drag", ("slide", 250, 100), 550, 350)])
    add("marquee partial", "marquee", two_shapes, [("drag", ("slide", 250, 100), 300, 150)])

    # Lines and tables.
    add("line click", "line", [LINE], [("click", (0, .5, 0))])
    add("line drag", "line", [LINE], [("drag", (0, .5, 0), 40, 30)])
    add("table click", "table", [TABLE], [("click", (0, .25, .25))])
    add("table type", "table", [TABLE], [("click", (0, .25, .25)), ("type", "Zelle")])
    add("table tab", "table", [TABLE], [("click", (0, .25, .25)), ("type", "A"), ("keys", "tab"), ("type", "B")])
    # Slides (Strg+M, slide keys in the slide area).
    add("ctrl+m", "slides", [], [("click", ("slide", 80, 480)), ("keys", "ctrl+m")])
    add("ctrl+m twice", "slides", [], [("click", ("slide", 80, 480)), ("keys", "ctrl+m"), ("keys", "ctrl+m")], layout="Titelfolie")
    add("pgdn slides", "slides", [], [("click", ("slide", 80, 480)), ("keys", "pagedown")], slides=3)
    add("pgdn pgup", "slides", [], [("click", ("slide", 80, 480)), ("keys", "pagedown"), ("keys", "pagedown"), ("keys", "pageup")], slides=3)
    add("end home", "slides", [], [("click", ("slide", 80, 480)), ("keys", "end")], slides=3)
    return out


# ----- running ---------------------------------------------------------------------------------------

def take_lock(name: str) -> bool:
    try:
        if LOCK.exists() and time.time() - LOCK.stat().st_mtime < 1800:
            return False
        LOCK.unlink(missing_ok=True)
        with open(LOCK, "x", encoding="utf-8") as f:
            f.write(f"{name} {datetime.datetime.now().isoformat(timespec='seconds')}")
        return True
    except FileExistsError:
        return False


def run(side_name: str, areas: list[str]) -> None:
    todo = [c for c in cases() if not areas or c.area in areas or c.name in areas]
    date = datetime.date.today().isoformat()
    rec_path = HERE / "results" / f"{date}-catalogue-present-powerpoint.json"
    if side_name == "present":
        recs = sorted((HERE / "results").glob("*-catalogue-present-powerpoint.json"))
        if not recs:
            raise SystemExit("no PowerPoint recording yet")
        rec = json.loads(recs[-1].read_text(encoding="utf-8"))["cases"]
    if not take_lock(f"step8-catalogue-{side_name}"):
        raise SystemExit("real-input lock held by another session: try later")
    out: dict = {}
    side = None
    try:
        side = PowerPointSide() if side_name == "powerpoint" else PresentSide()
        for i, c in enumerate(todo):
            try:
                side.reset(c)
                side.run(c)
                time.sleep(0.3)
                st = normalise(side.state(), c.tol)
                out[c.name] = {"area": c.area, "state": st}
            except Exception as e:  # noqa: BLE001
                import traceback
                where = traceback.extract_tb(e.__traceback__)[-1]
                out[c.name] = {"area": c.area, "error": f"{e!r} at {Path(where.filename).name}:{where.lineno}"[:300]}
            print(f"{i + 1}/{len(todo)} {c.name}: {'ok' if 'state' in out[c.name] else out[c.name]['error']}", flush=True)
            LOCK.touch()
    finally:
        if side:
            side.close()
        LOCK.unlink(missing_ok=True)
    if side_name == "powerpoint":
        old = json.loads(rec_path.read_text(encoding="utf-8"))["cases"] if rec_path.exists() and areas else {}
        old.update(out)
        rec_path.write_text(json.dumps({"date": date, "cases": old}, ensure_ascii=False, indent=1), encoding="utf-8", newline=chr(10))
        print(f"recorded {len(out)} cases")
        return
    res_path = HERE / "results" / f"{date}-catalogue-present.json"
    report = json.loads(res_path.read_text(encoding="utf-8"))["cases"] if res_path.exists() and areas else {}
    for c in todo:
        r, o = rec.get(c.name), out.get(c.name)
        if not r or "state" not in r:
            report[c.name] = {"area": c.area, "result": "no recording"}
            continue
        if "state" not in o:
            report[c.name] = {"area": c.area, "result": "error", "error": o.get("error")}
            continue
        d = differences(r["state"], o["state"], c.tol, c.ignore)
        report[c.name] = {"area": c.area, "result": "same" if not d else "different", "diff": d}
    same = sum(1 for v in report.values() if v["result"] == "same")
    rated = sum(1 for v in report.values() if v["result"] in ("same", "different", "error"))
    res_path.write_text(json.dumps({"date": date, "same": same, "of": rated, "cases": report}, ensure_ascii=False, indent=1), encoding="utf-8", newline=chr(10))
    print(f"same as PowerPoint: {same}/{rated}")


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print(__doc__)
    elif sys.argv[1] == "list":
        cs = cases()
        print(len(cs), "cases")
        for c in cs:
            print(c.area, "|", c.name)
    else:
        run(sys.argv[1], sys.argv[2:])
