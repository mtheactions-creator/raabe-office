"""Probe: how PowerPoint writes Einfügen > Zoom > Folienzoom (Step 8, D-072). COM has no method for zooms,
so this probe uses PowerPoint's keytips (Alt, I, Y, F) in a visible PowerPoint of our own on a presentation it
made in TEMP: real input, run only with the real-input lock (CLAUDE.md) and when the owner is not at the PC.
The dialog's slide check box is ticked and Einfügen pressed through UI Automation (no mouse). Any other dialog
(sign-in, consent, licence, Neuerungen) is not touched: the probe ends its own PowerPoint and stops.

    python probe_powerpoint_zoom.py

Result: results/<date>-powerpoint-zoom.json (the zoom's XML, relationships and content types); the .pptx in
captures/zoom/.
"""
from __future__ import annotations

import datetime
import json
import re
import shutil
import sys
import tempfile
import time
import zipfile
from pathlib import Path

from feellab import uia, winapi
from feellab.powerpoint import OwnerPowerPointRunning, PowerPointTarget

HERE = Path(__file__).parent
LOCK = Path(tempfile.gettempdir()) / "raabe-real-input.lock"


def take_lock() -> bool:
    try:
        if LOCK.exists() and time.time() - LOCK.stat().st_mtime < 1800:
            return False
        LOCK.unlink(missing_ok=True)
        with open(LOCK, "x", encoding="utf-8") as f:
            f.write(f"step8-probe-zoom {datetime.datetime.now().isoformat(timespec='seconds')}")
        return True
    except FileExistsError:
        return False


def toggle(el) -> None:
    p = el.GetCurrentPattern(uia.U.UIA_TogglePatternId)
    p.QueryInterface(uia.U.IUIAutomationTogglePattern).Toggle()


def main() -> int:
    if not take_lock():
        print("real-input lock held: try later")
        return 3
    try:
        try:
            ppt = PowerPointTarget(visible=True)
        except OwnerPowerPointRunning as e:
            print("skipped:", e)
            return 2
        out: dict = {"date": datetime.date.today().isoformat(), "what": "PowerPoint's XML for Einfügen > Zoom > Folienzoom"}
        try:
            pres = ppt.new_presentation(with_window=True, name="zoom")
            for i in range(3):
                s = ppt.add_slide(pres, layout=2 if i else 1)
                s.Shapes(1).TextFrame.TextRange.Text = f"Folie {i + 1}"
            pres.Save()
            pres.Windows(1).View.GotoSlide(1)
            main_hwnd = ppt.hwnd or ppt.main_window()
            before = set(ppt.windows())
            winapi.activate(main_hwnd, settle=0.5)
            for k in ["alt", "i", "y", "f"]:
                winapi.press(k)
                time.sleep(0.6)
            # The dialog: a new window of our process.
            dlg = 0
            for _ in range(30):
                new = [w for w in ppt.windows() if w not in before and w != main_hwnd]
                if new:
                    dlg = new[0]
                    break
                time.sleep(0.3)
            if not dlg:
                out["error"] = "no dialog"
                return 1
            title = winapi.window_title(dlg)
            out["dialog_title"] = title
            if "zoom" not in title.lower():
                out["error"] = "an unexpected dialog appeared; not touched"
                return 1
            root = uia.element_for(dlg)
            boxes = []
            stack = [root]
            while stack:
                e = stack.pop()
                try:
                    d = uia.describe(e)
                except Exception:  # noqa: BLE001
                    continue
                if d["type"] == "CheckBox":
                    boxes.append((d["name"], e))
                stack.extend(uia.children(e))
            out["checkboxes"] = [n for n, _ in boxes]
            target = next((e for n, e in boxes if "3" in n), boxes[-1][1] if boxes else None)
            if target is None:
                out["error"] = "no slide check boxes"
                return 1
            toggle(target)
            time.sleep(0.4)
            ok = uia.find(root, lambda d: d["type"] == "Button" and d["name"] == "Einfügen")
            if ok is None:
                out["error"] = "no Einfügen button"
                return 1
            uia.invoke(ok)
            time.sleep(2.0)
            slide = pres.Slides(1)
            out["shapes"] = [{"name": str(slide.Shapes(k).Name), "type": int(slide.Shapes(k).Type), "box": [float(slide.Shapes(k).Left), float(slide.Shapes(k).Top), float(slide.Shapes(k).Width), float(slide.Shapes(k).Height)]} for k in range(1, slide.Shapes.Count + 1)]
            copy = ppt.tmp / "zoom-copy.pptx"
            ppt.save_copy(pres, copy)
            cap = HERE / "captures" / "zoom"
            cap.mkdir(parents=True, exist_ok=True)
            shutil.copy(copy, cap / "zoom.pptx")
            z = zipfile.ZipFile(copy)
            xml = z.read("ppt/slides/slide1.xml").decode("utf-8")
            m = re.search(r"<mc:AlternateContent.*</mc:AlternateContent>", xml, re.S)
            out["slide_root"] = xml[:600]
            out["zoom_xml"] = m.group(0) if m else None
            out["rels"] = z.read("ppt/slides/_rels/slide1.xml.rels").decode("utf-8")
            out["content_types"] = z.read("[Content_Types].xml").decode("utf-8")
            out["presentation_ext"] = re.findall(r"<p:extLst>.*?</p:extLst>", z.read("ppt/presentation.xml").decode("utf-8"), re.S)
            out["media"] = [n for n in z.namelist() if "/media/" in n]
        finally:
            print(ppt.close(kill_if_unreachable=True))
            res = HERE / "results" / f"{out['date']}-powerpoint-zoom.json"
            res.write_text(json.dumps(out, indent=1, ensure_ascii=False), encoding="utf-8", newline="\n")
            print("wrote", res.name, out.get("error", "ok"))
    finally:
        LOCK.unlink(missing_ok=True)
    return 0


if __name__ == "__main__":
    sys.exit(main())
