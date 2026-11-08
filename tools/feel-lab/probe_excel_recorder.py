"""Records Excel's macro recorder (Makro aufzeichnen, D-119) by hand: for every case of recorder_cases.py our own visible
Excel records a macro while the case's keys and clicks are sent with real input, exactly as a student would do them.
Also recorded: the Makros menu (entries and keytips, before and while recording), the Makro aufzeichnen dialog
(labels, defaults, tab order), the status bar's record button and what Excel's own recorded macros do when run again.

Takes over keyboard and mouse: only under the real-input lock (CLAUDE.md), and keys are sent only while our own Excel
is in front. The recorder writes into "Diese Arbeitsmappe" (never the Persönliche Makroarbeitsmappe). Relative
Verweise verwenden is put back as it was. The VBA project is read from the saved .xlsm (feellab/vbasource.py), without
access to Excel's VBA object model.

    .venv/Scripts/python.exe probe_excel_recorder.py [case ...]

Result: results/2026-11-09-excel-recorder.json (code per case, the end state, the UI).
"""
from __future__ import annotations

import json
import re
import traceback
import sys
import tempfile
import time
from pathlib import Path

from feellab import uia, vbasource, winapi
from feellab.calc_catalogue import Case, ExcelSide, run_step, settle
from feellab.inputlock import real_input, touch
from recorder_cases import CASES

HERE = Path(__file__).parent
RESULT = HERE / "results" / "2026-11-09-excel-recorder.json"
TMP = Path(tempfile.gettempdir()) / "raabe-recorder"
TMP.mkdir(exist_ok=True)


def descendants(root) -> list[tuple]:
    out = []
    try:
        f = root.FindAll(uia.U.TreeScope_Descendants, uia._uia.CreateTrueCondition())
    except Exception:  # noqa: BLE001
        return out
    for i in range(f.Length):
        e = f.GetElement(i)
        try:
            out.append((e, uia.describe(e)))
        except Exception:  # noqa: BLE001
            continue
    return out


def value_of(el) -> str:
    try:
        return el.GetCurrentPattern(uia.U.UIA_ValuePatternId).QueryInterface(uia.U.IUIAutomationValuePattern).CurrentValue
    except Exception:  # noqa: BLE001
        return ""


class Recorder:
    def __init__(self, side: ExcelSide):
        self.side, self.x, self.pid = side, side.x, side.x.pid
        self.root = side.root

    def front(self, hwnd: int | None = None) -> None:
        """Our Excel (or its dialog) in front; other sessions' windows may take the foreground now and then."""
        for _ in range(6):
            try:
                winapi.activate(hwnd or self.x.hwnd, 0.3)
            except RuntimeError:
                pass
            if winapi.window_pid(winapi.foreground()) == self.pid:
                return
            time.sleep(0.5)
        raise RuntimeError("our Excel is not in front")

    def cleanup(self) -> None:
        """After a failure: our own dialogs closed (Abbrechen), so that Excel can be closed."""
        for h in winapi.find_windows(lambda h: winapi.window_pid(h) == self.pid and h != self.x.hwnd and winapi.window_class(h) in ("bosa_sdm_XL9", "#32770")):
            winapi.user32.PostMessageW(h, 0x0010, 0, 0)
        time.sleep(0.8)

    def all_elements(self) -> list[tuple]:
        res = descendants(self.root)
        for h in winapi.find_windows(lambda h: winapi.window_pid(h) == self.pid and h != self.x.hwnd):
            res.extend(descendants(uia.element_for(h)))
        return res

    def tab(self, name: str) -> None:
        t = uia.find(self.root, lambda d: d["type"] == "TabItem" and d["name"] == name, 25)
        uia.select(t)
        time.sleep(0.8)

    def macros_menu(self) -> list[dict]:
        """Opens Ansicht > Makros' menu; its entries (name, keytip, enabled)."""
        self.tab("Ansicht")
        split = next(e for e, d in descendants(self.root) if d["type"] == "SplitButton" and d["name"] == "Makros")
        split.GetCurrentPattern(uia.U.UIA_ExpandCollapsePatternId).QueryInterface(uia.U.IUIAutomationExpandCollapsePattern).Expand()
        time.sleep(1.0)
        self.split = split
        items, seen = [], set()
        for e, d in self.all_elements():
            if d["type"] in ("MenuItem", "Button", "CheckBox") and d["name"] and ("Makro" in d["name"] or "Aufzeichnung" in d["name"] or "Verweise" in d["name"]):
                if (d["name"], d["type"]) in seen:
                    continue
                seen.add((d["name"], d["type"]))
                toggled = None
                try:
                    toggled = bool(e.GetCurrentPattern(uia.U.UIA_TogglePatternId).QueryInterface(uia.U.IUIAutomationTogglePattern).CurrentToggleState)
                except Exception:  # noqa: BLE001
                    pass
                items.append({"name": d["name"], "type": d["type"], "key": d["key"], "enabled": d["enabled"], "toggled": toggled, "el": e})
        return items

    def close_menu(self) -> None:
        try:
            self.split.GetCurrentPattern(uia.U.UIA_ExpandCollapsePatternId).QueryInterface(uia.U.IUIAutomationExpandCollapsePattern).Collapse()
        except Exception:  # noqa: BLE001
            pass
        time.sleep(0.4)

    def menu_entry(self, starts: str) -> None:
        items = self.macros_menu()
        hit = next((i for i in items if i["name"].startswith(starts) and i["type"] == "MenuItem"), None) or next((i for i in items if i["name"].startswith(starts)), None)
        if hit is None:
            self.close_menu()
            raise RuntimeError(f"no menu entry {starts}: {[i['name'] for i in items]}")
        uia.invoke(hit["el"])
        time.sleep(1.0)

    def dialog(self) -> int | None:
        for _ in range(30):
            hs = winapi.find_windows(lambda h: winapi.window_pid(h) == self.pid and h != self.x.hwnd and winapi.window_title(h) == "Makro aufzeichnen")
            if hs:
                return hs[0]
            time.sleep(0.2)
        return None

    def dialog_open(self) -> bool:
        return bool(winapi.find_windows(lambda h: winapi.window_pid(h) == self.pid and h != self.x.hwnd and winapi.window_title(h) == "Makro aufzeichnen"))

    def keytips(self, *letters: str) -> None:
        """Alt, then the letters, with real keys while our Excel is in front."""
        self.front()
        winapi.press("alt")
        time.sleep(0.4)
        for ch in letters:
            winapi.press(ch)
            time.sleep(0.35)
        time.sleep(0.6)

    def relative_item(self) -> bool | None:
        """Opens Ansicht > Makros by its keytips (Alt, F, M) and reads Relative Verweise verwenden's state; the menu
        stays open (the caller presses V or Esc)."""
        for _ in range(4):
            self.keytips("f", "m")
            time.sleep(0.5)
            for e, d in self.all_elements():
                if d["type"] == "MenuItem" and d["name"] == "Relative Verweise verwenden":
                    try:
                        return bool(e.GetCurrentPattern(uia.U.UIA_TogglePatternId).QueryInterface(uia.U.IUIAutomationTogglePattern).CurrentToggleState)
                    except Exception:  # noqa: BLE001
                        return None
            winapi.press("esc")
            winapi.press("esc")
            time.sleep(0.5)
        raise RuntimeError("the Makros menu does not open")

    def relative(self) -> bool:
        on = self.relative_item()
        winapi.press("esc")
        winapi.press("esc")
        time.sleep(0.3)
        return bool(on)

    def set_relative(self, want: bool) -> None:
        for _ in range(3):
            on = self.relative_item()
            if bool(on) == want:
                winapi.press("esc")
                winapi.press("esc")
                time.sleep(0.3)
                return
            winapi.press("v")
            time.sleep(0.8)
        raise RuntimeError("Relative Verweise verwenden does not change")

    def status_buttons(self) -> list[dict]:
        bar = uia.find(self.root, lambda d: d["type"] == "StatusBar", max_depth=6)
        return [{k: d[k] for k in ("type", "name")} for _, d in descendants(bar) if d["type"] == "Button" and d["name"]] if bar else []


GUARD = {"pid": 0, "hwnd": 0}
_send = winapi._send


def _guarded(*inputs) -> None:
    """Every key and mouse event goes only to our own Excel: other sessions' windows take the foreground now and
    then (UI Automation calls of invisible probes); the event waits until ours is in front again."""
    for _ in range(40):
        if winapi.window_pid(winapi.foreground()) == GUARD["pid"]:
            return _send(*inputs)
        try:
            popup = winapi.user32.GetLastActivePopup(GUARD["hwnd"])
            winapi.activate(popup or GUARD["hwnd"], 0.15)
        except Exception:  # noqa: BLE001
            time.sleep(0.15)
    raise RuntimeError("our Excel does not stay in front")


winapi._send = _guarded


def force_close(pid: int) -> None:
    """Our own Excel when COM no longer answers: its window closed, its own workbook not saved."""
    for h in winapi.find_windows(lambda h: winapi.window_pid(h) == pid and winapi.window_class(h) == "XLMAIN"):
        winapi.user32.PostMessageW(h, 0x0010, 0, 0)
    time.sleep(2)
    for h in winapi.find_windows(lambda h: winapi.window_pid(h) == pid and winapi.window_class(h) == "NUIDialog" and winapi.user32.IsWindowVisible(h)):
        b = next((e for e, d in descendants(uia.element_for(h)) if d["type"] == "Button" and d["name"].startswith("Nicht speichern")), None)
        if b is not None:
            uia.invoke(b)


def main() -> None:
    sys.stdout.reconfigure(encoding="utf-8")
    want = set(sys.argv[1:])
    cases = [c for c in CASES if not want or c["name"] in want]
    out: dict = {"cases": {}}
    with real_input("calc-misc probe_excel_recorder"):
        side = ExcelSide()
        rec = Recorder(side)
        GUARD.update(pid=side.x.pid, hwnd=side.x.hwnd)
        relative0 = None
        try:
            rec.front()
            relative0 = False
            menu = rec.macros_menu()
            out["menu"] = [{k: v for k, v in i.items() if k != "el"} for i in menu]
            rec.close_menu()
            out["statusBar"] = rec.status_buttons()
            print("menu", out["menu"], out["statusBar"], flush=True)
            for i, c in enumerate(cases, 1):
                try:
                    touch()
                    case = Case(c["name"], c["setup"], c["steps"], "recorder", start=c["start"])
                    side.reset(case)
                    side.prepare(case)
                    rec.front()
                    if c.get("relative"):
                        rec.keytips("f", "m", "v")
                        winapi.press("esc")
                    rec.front()
                    name = f"Makro_{c['name']}"
                    # Ansicht > Makros > Makro aufzeichnen... by its keytips (Alt, F, M, R as recorded above).
                    h = None
                    for _ in range(3):
                        rec.keytips("f", "m", "r")
                        h = rec.dialog()
                        if h:
                            break
                        winapi.press("esc")
                        winapi.press("esc")
                    if not h:
                        raise RuntimeError("no Makro aufzeichnen dialog")
                    if i == 1:
                        el = uia.element_for(h)
                        out["dialog"] = {"title": winapi.window_title(h), "class": winapi.window_class(h),
                                         "controls": [{"type": d["type"], "name": d["name"], "key": d["key"], "value": value_of(e)}
                                                      for e, d in descendants(el) if d["name"] or d["type"] in ("Edit", "ComboBox")]}
                        out["dialogRect"] = list(winapi.window_rect(h))
                    rec.front(h)
                    # By hand: the name is selected in Makroname; Tab goes to the shortcut, the place, the description.
                    winapi.press("ctrl+a")
                    winapi.type_text(name, gap=0.02)
                    winapi.press("tab")
                    if c.get("shortcut"):
                        winapi.type_text(c["shortcut"], gap=0.02)
                    winapi.press("tab")
                    winapi.press("tab")
                    if c.get("description"):
                        winapi.type_text(c["description"], gap=0.02)
                    time.sleep(0.2)
                    if i == 1:
                        out["dialogFilled"] = [{"type": d["type"], "name": d["name"], "value": value_of(e)} for e, d in descendants(uia.element_for(h)) if d["type"] in ("Edit", "ComboBox")]
                    # Enter would start a new line in Beschreibung: OK as a click on it.
                    ok = next(e for e, d in descendants(uia.element_for(h)) if d["type"] == "Button" and d["name"] == "OK")
                    r = uia.describe(ok)["rect"]
                    winapi.click(r[0] + r[2] / 2, r[1] + r[3] / 2)
                    time.sleep(0.8)
                    if rec.dialog_open():
                        raise RuntimeError("Makro aufzeichnen did not close")
                    rec.front()
                    if i == 1:
                        out["statusBarRecording"] = rec.status_buttons()
                        try:
                            menu = rec.macros_menu()
                            out["menuRecording"] = [{k: v for k, v in it.items() if k != "el"} for it in menu]
                        except Exception as e:  # noqa: BLE001
                            out["menuRecording"] = str(e)
                        rec.close_menu()
                        rec.front()
                        winapi.press("esc")
                        print("recording", out["menuRecording"], out["statusBarRecording"], flush=True)
                        rec.front()
                        rec.tab("Start")
                    rec.front()
                    for st in c["steps"]:
                        if winapi.window_pid(winapi.foreground()) != rec.pid:
                            rec.front()
                        if st[0] == "keytips":
                            winapi.press("alt")
                            time.sleep(0.35)
                            for tip in st[1].split():
                                for ch in tip:
                                    winapi.press(ch.lower())
                                    time.sleep(0.25)
                            time.sleep(0.5)
                        elif st[0] == "ribbon":
                            # The left part of the Füllfarbe split button (it fills with the last colour, yellow at first).
                            rec.tab("Start")
                            btn = next(e for e, d in descendants(rec.root) if d["type"] in ("SplitButton", "Button") and d["name"].startswith("Füllfarbe"))
                            r = uia.describe(btn)["rect"]
                            winapi.click(r[0] + min(r[2] * 0.3, 20), r[1] + r[3] / 2)
                            time.sleep(0.6)
                            rec.front()
                        else:
                            if st[0] == "fill":
                                # The fill handle: 3 px inside the corner (1 px landed in the cell below at 150 %).
                                last = st[1].split(":")[-1]
                                x1, y1 = side.point(last, "br")
                                x2, y2 = side.point(st[2])
                                winapi.drag(x1 - 3, y1 - 3, x2, y2)
                            else:
                                run_step(side, st)
                    time.sleep(0.6)
                    # Aufzeichnung beenden through the status bar's button (Excel shows it while recording).
                    bar = uia.find(rec.root, lambda d: d["type"] == "StatusBar", max_depth=6)
                    stop = next((e for e, d in descendants(bar) if d["type"] == "Button" and "aufzeichnung" in d["name"].lower()), None) if bar else None
                    if stop is not None:
                        uia.invoke(stop)
                    else:
                        rec.menu_entry("Aufzeichnung beenden")
                    time.sleep(0.8)
                    st = side.state(case)
                    ws = side.x.app.ActiveSheet
                    extra = {}
                    for a in st["cells"]:
                        r = ws.Range(a)
                        extra[a] = {"fill": None if r.Interior.ColorIndex == -4142 else int(r.Interior.Color), "size": r.Font.Size, "h": r.HorizontalAlignment}
                    st["extra"] = extra
                    st["rowHidden"] = [k for k in range(1, 12) if ws.Rows(k).Hidden]
                    out["cases"][c["name"]] = {"macro": name, "state": st}
                    if c.get("relative"):
                        rec.keytips("f", "m", "v")
                        winapi.press("esc")
                    print(c["name"], "recorded", flush=True)
                    settle()
                except Exception as e:  # noqa: BLE001
                    print(c["name"], "FAILED", repr(e), traceback.format_exc(), flush=True)
                    out["cases"][c["name"]] = {"failed": repr(e)[:300]}
                    rec.cleanup()
                    try:
                        bar = uia.find(rec.root, lambda d: d["type"] == "StatusBar", max_depth=6)
                        stop = next((e2 for e2, d in descendants(bar) if d["type"] == "Button" and d["name"].endswith("Wird aufgezeichnet")), None)
                        if stop is not None:
                            uia.invoke(stop)
                        winapi.press("esc")
                    except Exception:  # noqa: BLE001
                        pass
            # Put Relative Verweise verwenden back.
            rec.front()
            try:
                relative0 = None
            except Exception as e:  # noqa: BLE001
                print("Relative Verweise not put back", e, flush=True)
            path = TMP / "excel-recorder.xlsm"
            if path.exists():
                path.unlink()
            side.x.app.DisplayAlerts = False
            side.x.wb.SaveAs(str(path), 52)
        finally:
            rec.cleanup()
            try:
                pass
            except Exception:  # noqa: BLE001
                pass
            try:
                side.close()
            except Exception:  # noqa: BLE001
                force_close(rec.pid)
    mods = vbasource.modules(TMP / "excel-recorder.xlsm")
    code = chr(10).join(v.replace(chr(13), "") for v in mods.values())
    # Each case's recorded Sub, Attribute lines included (they carry the shortcut and the description).
    for name, rec_ in out["cases"].items():
        m = re.search(r"^Sub Makro_" + name + r"\(\)$.*?^End Sub$", code, re.S | re.M)
        if m and "macro" in rec_:
            rec_["code"] = m.group(0)
    if want and RESULT.exists():
        old = json.loads(RESULT.read_text(encoding="utf-8"))
        old["cases"].update(out["cases"])
        out = {**old, "cases": old["cases"]}
    out.pop("modules", None)
    RESULT.write_text(json.dumps(out, ensure_ascii=False, indent=1), encoding="utf-8", newline="\n")
    for k, v in out["cases"].items():
        print("=====", k)
        print(v.get("code") or v.get("failed"))


if __name__ == "__main__":
    main()
