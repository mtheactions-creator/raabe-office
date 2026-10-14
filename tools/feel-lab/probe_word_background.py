"""Probe: how Word writes Entwurf > Seitenhintergrund. In our own visible Word instance a
watermark (ENTWURF 1 and 2, real clicks in the gallery), a page colour and page borders are
applied and the document saved; the XML Word wrote is printed (header, w:background,
w:pgBorders), and the dialog Gedrucktes Wasserzeichen is captured. Takes over mouse and screen.

    .venv/Scripts/python.exe probe_word_background.py <scratch folder>
"""
from __future__ import annotations

import re
import sys
import time
import zipfile
from pathlib import Path

from feellab import uia, winapi
from probe_word_dialogs import CAP, Word, new_windows


def gallery_pick(w: Word, button: str, item: str) -> bool:
    ctl = w.control(button)
    if ctl is None:
        print("no", button)
        return False
    r = uia.describe(ctl)["rect"]
    winapi.click(r[0] + r[2] / 2, r[1] + r[3] - 10 if r[3] > 60 else r[1] + r[3] / 2)
    time.sleep(1.4)
    el = uia.find(w.root, lambda d: d["name"] == item and d["type"] in ("ListItem", "MenuItem"), max_depth=22)
    if el is None:
        print("no item", item)
        winapi.press("esc")
        return False
    try:
        uia.invoke(el)
    except Exception:  # noqa: BLE001
        rr = uia.describe(el)["rect"]
        winapi.click(rr[0] + rr[2] / 2, rr[1] + rr[3] / 2)
    time.sleep(1.5)
    return True


def parts(path: Path) -> dict[str, str]:
    z = zipfile.ZipFile(path)
    return {n: z.read(n).decode("utf-8", "replace") for n in z.namelist() if n.endswith(".xml")}


def main() -> None:
    sys.stdout.reconfigure(encoding="utf-8")
    out = Path(sys.argv[1])
    w = Word()
    d = w.d
    try:
        w.tab("Entwurf")
        for n, item in [(1, "ENTWURF 1"), (2, "ENTWURF 2")]:
            if gallery_pick(w, "Wasserzeichen", item):
                p = out / f"wm{n}.docx"
                d.SaveAs2(str(p), 16)
                for name, xml in parts(p).items():
                    if "PowerPlusWaterMark" in xml:
                        m = re.search(r"<w:sdt>.*?</w:sdt>", xml, re.S)
                        print(f"--- {item} in {name}:\n{(m.group(0) if m else xml)[:3000]}\n", flush=True)
        # Benutzerdefiniertes Wasserzeichen...
        ctl = w.control("Wasserzeichen")
        if ctl is not None:
            r = uia.describe(ctl)["rect"]
            winapi.click(r[0] + r[2] / 2, r[1] + r[3] - 10)
            time.sleep(1.4)
            el = uia.find(w.root, lambda x: x["name"] == "Benutzerdefiniertes Wasserzeichen..." and x["type"] == "MenuItem", max_depth=22)
            if el is not None:
                before = set(winapi.find_windows(lambda h: True))
                rr = uia.describe(el)["rect"]
                winapi.click(rr[0] + rr[2] / 2, rr[1] + rr[3] / 2)
                time.sleep(1.8)
                dlg = next((h for h in new_windows(before) if winapi.window_class(h) in ("bosa_sdm_msword", "NUIDialog", "#32770")), None)
                if dlg:
                    l, t, rgt, b = winapi.window_rect(dlg)
                    winapi.save_region_png(str(CAP / "watermark dialog.png"), l, t, rgt - l, b - t)
                    cancel = uia.find(uia.element_for(dlg), lambda x: x["type"] == "Button" and x["name"] == "Abbrechen", max_depth=10)
                    if cancel is not None:
                        rr = uia.describe(cancel)["rect"]
                        winapi.click(rr[0] + rr[2] / 2, rr[1] + rr[3] / 2)
                    time.sleep(1.0)
                    print("watermark dialog captured", flush=True)
            else:
                winapi.press("esc")
        # Seitenfarbe and Seitenränder.
        if gallery_pick(w, "Seitenfarbe", "Orange, Akzent 2, heller 80%"):
            pass
        b = d.Sections(1).Borders
        for side in (-1, -2, -3, -4):
            b(side).LineStyle = 1
            b(side).LineWidth = 12
            b(side).Color = 255
        b.DistanceFrom = 1
        b.DistanceFromTop = 24
        p = out / "bg.docx"
        d.SaveAs2(str(p), 16)
        x = parts(p)
        doc = x["word/document.xml"]
        print("--- background:", re.findall(r"<w:background[^>]*?(?:/>|>.*?</w:background>)", doc, re.S)[:2])
        print("--- pgBorders:", re.findall(r"<w:pgBorders.*?</w:pgBorders>", doc, re.S)[:1])
        print("--- settings:", re.findall(r"<w:displayBackgroundShape[^>]*/>", x["word/settings.xml"]))
    finally:
        w.close()


if __name__ == "__main__":
    main()
