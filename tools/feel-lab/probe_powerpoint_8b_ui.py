"""PowerPoint's UI for Step 8b (D-074, D-076), recorded with real input while the owner is away (lock, own
PowerPoint, only when no POWERPNT.EXE runs):

  effects    Form formatieren > Effekte: the sections and their controls' labels, and what X-Drehung shows for a
             shape whose RotationX is 30 through COM
  export     Datei > Exportieren > Video erstellen and Animiertes GIF erstellen: the controls' labels and the
             choices of their lists; then a GIF made through the page (saved in TEMP) and measured
  presenter  the show with Referentenansicht on two screens (the virtual display): the presenter view's controls

    .venv/Scripts/python.exe probe_powerpoint_8b_ui.py [effects] [export] [presenter]

Results: results/<date>-powerpoint-8b-ui.json holds labels (names of buttons, fields, lists, tabs; nothing longer than
60 characters, so no descriptions) and measurements. Full UI dumps and screenshots go to captures/ (not committed).
"""
from __future__ import annotations

import datetime
import json
import sys
import tempfile
import time
from pathlib import Path

from catalogue_present import LOCK, take_lock
from feellab import uia, winapi
from feellab.powerpoint import PowerPointTarget

HERE = Path(__file__).parent
CAP = HERE / "captures" / "ui8b"
LABEL_TYPES = {"Button", "Edit", "ComboBox", "CheckBox", "RadioButton", "TabItem", "ListItem", "MenuItem", "SplitButton", "Spinner", "Slider", "Text", "Group", "Hyperlink"}


def labels(el, max_depth: int = 25, rect: bool = False, ytop: bool = False) -> list[dict]:
    """The labelled controls under an element (type, name, value when it has one), in order."""
    out = []
    def walk(e, d):
        try:
            x = uia.describe(e)
        except Exception:  # noqa: BLE001
            return
        if x["type"] in LABEL_TYPES and x["name"] and len(x["name"]) <= 60:
            item = {"type": x["type"], "name": x["name"], "enabled": x["enabled"]}
            if rect:
                item["x"] = x["rect"][0]
            if ytop:
                item["y"] = x["rect"][1]
            try:
                vp = e.GetCurrentPattern(uia.U.UIA_ValuePatternId)
                if vp:
                    item["value"] = vp.QueryInterface(uia.U.IUIAutomationValuePattern).CurrentValue
            except Exception:  # noqa: BLE001
                pass
            out.append(item)
        if d < max_depth:
            for c in uia.children(e):
                walk(c, d + 1)
    walk(el, 0)
    return out


def page_labels(hwnd: int, root) -> list[dict]:
    """Labels of a window without its title bar (which shows the file name and the signed-in account)."""
    top = winapi.window_rect(hwnd)[1] + 90
    out = []
    for i in labels(root, rect=True, ytop=True):
        x, y = i.pop("x", 0), i.pop("y", 0)
        if y >= top and not i["name"].endswith("gespeichert"):
            out.append(i)
    return out


def dump(name: str, el) -> None:
    CAP.mkdir(parents=True, exist_ok=True)
    (CAP / f"{name}.json").write_text(json.dumps(uia.tree(el, 0, 25), ensure_ascii=False, indent=1), encoding="utf-8")


def shot(name: str, hwnd: int) -> None:
    CAP.mkdir(parents=True, exist_ok=True)
    x0, y0, x1, y1 = winapi.window_rect(hwnd)
    winapi.save_region_png(str(CAP / f"{name}.png"), x0, y0, x1 - x0, y1 - y0)


def by_name(root, name: str, types: set[str] | None = None, depth: int = 25):
    return uia.find(root, lambda x: x["name"] == name and (not types or x["type"] in types), depth)


OWN_PID = 0


class NotOurs(RuntimeError):
    pass


def ours() -> None:
    """Real input only into our own PowerPoint: the window in front must belong to its process."""
    fg = winapi.foreground()
    if not OWN_PID or winapi.window_pid(fg) != OWN_PID:
        raise NotOurs(f"window in front is not our PowerPoint ({winapi.window_title(fg)[:40]!r}): stopped")


def press(key: str) -> None:
    ours()
    winapi.press(key)


def click_el(el) -> None:
    ours()
    r = uia.describe(el)["rect"]
    winapi.click(r[0] + r[2] / 2, r[1] + r[3] / 2)
    time.sleep(0.6)


def effects(t: PowerPointTarget, out: dict, key: str = "effects") -> None:
    pres = t.new_presentation(with_window=True, name="effects")
    sl = pres.Slides.Add(1, 12)
    sh = sl.Shapes.AddShape(1, 200, 150, 300, 200)
    sh.ThreeD.RotationX = 30
    winapi.activate(t.hwnd)
    sh.Select()
    time.sleep(0.5)
    t.app.CommandBars.ExecuteMso("ObjectFormatDialog")
    time.sleep(2)
    root = uia.element_for(t.hwnd)
    # The pane: its tabs (Formoptionen) give its top left; only what lies right of and below them is read.
    top = by_name(root, "Formoptionen", {"TabItem"})
    if top is None:
        out["effects"] = {"error": "pane not found"}
        return
    tr = uia.describe(top)["rect"]
    left, ytop = tr[0] - 30, tr[1] - 5
    in_pane = lambda x: x["rect"][0] >= left and x["rect"][1] >= ytop and x["rect"][2] > 0
    tab = uia.find(root, lambda x: x["name"] == "Effekte" and x["type"] == "ListItem" and in_pane(x), 30)
    if tab is None:
        out["effects"] = {"error": "Effekte not found"}
        return
    try:
        uia.select(tab)
    except Exception:  # noqa: BLE001
        click_el(tab)
    time.sleep(1)
    # Each section's head button (re-found after each click: opening one moves the others down).
    heads = []
    import os
    for sec in (os.environ.get("SECTIONS") or "Schatten,Spiegelung,Leuchten,Weiche Kanten,3D-Format,3D-Drehung").split(","):
        root = uia.element_for(t.hwnd)
        e = uia.find(root, lambda x: x["name"] == sec and x["type"] == "Button" and in_pane(x), 30)
        if e is not None:
            heads.append(sec)
            try:
                uia.invoke(e)
            except Exception:  # noqa: BLE001
                click_el(e)
            time.sleep(0.6)
    time.sleep(1)
    root = uia.element_for(t.hwnd)
    shot("effects-pane", t.hwnd)
    items = []
    for i in labels(root, rect=True, ytop=True):
        x, y = i.pop("x", 0), i.pop("y", 0)
        if x >= left and y >= ytop:
            items.append(i)
    xr = next((i for i in items if i["name"].startswith("X-Drehung") and "value" in i), None)
    out[key] = {"heads": heads, "labels": items, "x_rotation_shown": xr.get("value") if xr else None, "com_rotation_x": float(sh.ThreeD.RotationX)}
    pres.Saved = -1


def export(t: PowerPointTarget, out: dict) -> None:
    deck = HERE / "captures" / "video" / "video-deck.pptx"
    pres = t.open_copy(deck, with_window=True) if deck.exists() else t.new_presentation(with_window=True, name="export")
    time.sleep(1.5)
    t.app.Visible = -1
    t.hwnd = t.main_window()
    t.app.WindowState = 1
    winapi.place(t.hwnd, 40, 40, 1600, 1000)
    winapi.activate(t.hwnd)
    res: dict = {}
    for page in ("Video erstellen", "Animiertes GIF erstellen"):
        press("alt+d")  # Datei
        time.sleep(2)
        root = uia.element_for(t.hwnd)
        exp = by_name(root, "Exportieren", {"ListItem", "TabItem", "Button", "MenuItem", "Hyperlink"})
        if exp is None:
            res[page] = {"error": "Exportieren not found"}
            press("esc")
            continue
        click_el(exp)
        time.sleep(1.5)
        root = uia.element_for(t.hwnd)
        item = by_name(root, page, {"ListItem", "Button", "TabItem", "MenuItem", "Hyperlink", "Text"})
        if item is None:
            res[page] = {"error": f"{page} not found"}
            press("esc")
            continue
        click_el(item)
        time.sleep(1.5)
        root = uia.element_for(t.hwnd)
        dump(f"export-{page}", root)
        shot(f"export-{page}", t.hwnd)
        res[page] = {"labels": page_labels(t.hwnd, root)}
        # The choices of the page's lists (opened one by one, read, closed).
        choices = {}
        for combo in [c for c in labels(root) if c["type"] == "ComboBox"][:4]:
            e = by_name(root, combo["name"], {"ComboBox"})
            if e is None:
                continue
            try:
                click_el(e)
                time.sleep(0.8)
                lst = uia.element_for(t.hwnd)
                choices[combo["name"]] = [i["name"] for i in labels(lst) if i["type"] == "ListItem"][:20]
                press("esc")
                time.sleep(0.5)
            except Exception:  # noqa: BLE001
                press("esc")
        res[page]["choices"] = choices
        if page == "Animiertes GIF erstellen":
            res[page]["gif"] = make_gif(t, root)
        else:
            press("esc")
            time.sleep(1)
    out["export"] = res
    pres.Saved = -1


def make_gif(t: PowerPointTarget, root) -> dict:
    """GIF erstellen with the page's defaults, saved into TEMP, then measured."""
    btn = by_name(root, "GIF erstellen", {"Button"})
    if btn is None:
        return {"error": "GIF erstellen not found"}
    target = Path(tempfile.gettempdir()) / "raabe-ppt-gif-probe.gif"
    target.unlink(missing_ok=True)
    click_el(btn)
    # PowerPoint's save dialog (a window of its own process): the file name field and Speichern through UI Automation.
    dlg = None
    for _ in range(40):
        time.sleep(0.25)
        cands = [h for h in winapi.find_windows(lambda h: winapi.window_pid(h) == OWN_PID) if winapi.window_class(h) == "#32770"]
        if cands:
            dlg = cands[0]
            break
    if not dlg:
        return {"error": "no save dialog"}
    droot = uia.element_for(dlg)
    field = uia.find(droot, lambda x: x["type"] == "Edit" and x["name"].startswith("Dateiname"), 20)
    save = uia.find(droot, lambda x: x["type"] == "Button" and x["name"] == "Speichern", 20)
    if field is None or save is None:
        return {"error": "save dialog fields not found"}
    vp = field.GetCurrentPattern(uia.U.UIA_ValuePatternId).QueryInterface(uia.U.IUIAutomationValuePattern)
    vp.SetValue(str(target))
    time.sleep(0.3)
    uia.invoke(save)
    for _ in range(240):
        time.sleep(0.5)
        if target.exists() and target.stat().st_size > 0:
            size = target.stat().st_size
            time.sleep(2)
            if target.stat().st_size == size:
                break
    if not target.exists():
        return {"error": "no GIF written"}
    from PIL import Image
    im = Image.open(target)
    delays = []
    for i in range(im.n_frames):
        im.seek(i)
        delays.append(im.info.get("duration"))
    return {"size": list(im.size), "frames": im.n_frames, "loop": im.info.get("loop"), "delays_ms": sorted(set(d or 0 for d in delays))[:10], "total_ms": sum(d or 0 for d in delays), "bytes": target.stat().st_size}


def presenter(t: PowerPointTarget, out: dict) -> None:
    from check_present_screens import display, monitors
    before = monitors()
    mons = display("extend", 2)
    try:
        pres = t.new_presentation(with_window=True, name="presenter")
        for i in range(3):
            sl = pres.Slides.Add(i + 1, 2)
            sl.Shapes(1).TextFrame.TextRange.Text = f"Folie {i + 1}"
            sl.NotesPage.Shapes.Placeholders(2).TextFrame.TextRange.Text = f"Notiz {i + 1}"
        ss = pres.SlideShowSettings
        ss.ShowPresenterView = -1
        ss.Run()
        time.sleep(4)
        wins = winapi.find_windows(lambda h: winapi.window_pid(h) == t.pid)
        info = []
        for h in wins:
            title = winapi.window_title(h)
            rect = winapi.window_rect(h)
            info.append({"title": title[:60], "class": winapi.window_class(h), "rect": rect})
        out["presenter_windows"] = info
        # The presenter view: the show window on the main screen (the other is the audience's).
        main = [h for h in wins if winapi.window_class(h) == "PodiumParent"]
        if main:
            # In front first: the screenshot is taken only when the presenter view itself is the window in front.
            winapi.activate(main[0])
            time.sleep(2.5)
            root = uia.element_for(main[0])
            dump("presenter-view", root)
            fg = winapi.foreground()
            if winapi.window_pid(fg) == OWN_PID and winapi.window_class(fg) == "PodiumParent":
                shot("presenter-view", main[0])
            else:
                out["presenter_shot"] = "skipped: presenter view not in front"
            out["presenter"] = {"labels": page_labels(main[0], root)}
        else:
            out["presenter"] = {"error": "presenter window not found", "monitors": len(mons), "before": len(before)}
        try:
            pres.SlideShowWindow.View.Exit()
        except Exception:  # noqa: BLE001
            press("esc")
        time.sleep(1)
        pres.Saved = -1
    finally:
        display("internal", 1)


def main() -> None:
    parts = sys.argv[1:] or ["effects", "export", "presenter"]
    if not take_lock("step8b-ui"):
        raise SystemExit("real-input lock held: try later")
    out: dict = {"date": datetime.date.today().isoformat()}
    t = None
    try:
        t = PowerPointTarget(visible=True)
        global OWN_PID
        OWN_PID = t.pid
        for p in parts:
            try:
                {"effects": effects, "effects3d": lambda t, o: effects(t, o, "effects3d"), "export": export, "presenter": presenter}[p](t, out)
            except NotOurs as e:
                out[f"{p}_error"] = str(e)[:300]
                break
            except Exception as e:  # noqa: BLE001
                out[f"{p}_error"] = str(e)[:300]
            LOCK.touch()
    finally:
        if t:
            out["close"] = t.close(kill_if_unreachable=True)
        LOCK.unlink(missing_ok=True)
    p = HERE / "results" / f"{out['date']}-powerpoint-8b-ui.json"
    old = json.loads(p.read_text(encoding="utf-8")) if p.exists() else {}
    old.update(out)
    p.write_text(json.dumps(old, ensure_ascii=False, indent=1), encoding="utf-8", newline="\n")
    print(json.dumps({k: (v if not isinstance(v, dict) else {kk: (vv if not isinstance(vv, list) else len(vv)) for kk, vv in v.items()}) for k, v in out.items()}, ensure_ascii=False)[:2500])


if __name__ == "__main__":
    main()
