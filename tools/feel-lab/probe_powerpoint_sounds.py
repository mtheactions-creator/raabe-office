"""Übergänge > Sound as PowerPoint writes it (Step 8b, D-078): an invisible PowerPoint of our own (only when no
POWERPNT.EXE runs) gives one slide each of its built-in transition sounds (SoundEffect.Name by the German name),
saves the presentation in TEMP and records the slide's p:sndAc and the embedded file's name, part path and size
(the sound files themselves are Microsoft's and are not kept).

    .venv/Scripts/python.exe probe_powerpoint_sounds.py

Result: results/<date>-powerpoint-sounds.json.
"""
from __future__ import annotations

import datetime
import json
import re
import zipfile
from pathlib import Path

from feellab.powerpoint import PowerPointTarget

HERE = Path(__file__).parent
NAMES = ["Applaus", "Bombe", "Brise", "Explosion", "Glocken", "Hammer", "Kamera", "Kasse", "Klick", "Laser", "Münze", "Pfeil",
         "Schreibmaschine", "Sog", "Spannung", "Stoß", "Trommelwirbel", "Wind", "Zischen"]


def main() -> None:
    out: dict = {"date": datetime.date.today().isoformat(), "sounds": []}
    t = PowerPointTarget(visible=False)
    try:
        pres = t.new_presentation(with_window=False, name="sounds")
        for i, n in enumerate(NAMES, start=1):
            sl = pres.Slides.Add(i, 12)
            sl.SlideShowTransition.EntryEffect = 1793
            try:
                sl.SlideShowTransition.SoundEffect.Name = n
            except Exception as e:  # noqa: BLE001
                out["sounds"].append({"name": n, "error": str(e)[:120]})
        path = t.tmp / "sounds.pptx"
        t.save_copy(pres, path)
        with zipfile.ZipFile(path) as z:
            names = z.namelist()
            for i, n in enumerate(NAMES, start=1):
                xml = z.read(f"ppt/slides/slide{i}.xml").decode("utf-8")
                rels = z.read(f"ppt/slides/_rels/slide{i}.xml.rels").decode("utf-8")
                snd = re.search(r"<p:sndAc>.*?</p:sndAc>", xml, re.S)
                rid = re.search(r'r:embed="(rId\d+)"', snd.group(0)) if snd else None
                target = re.search(rf'Id="{rid.group(1)}"[^>]*Target="([^"]+)"', rels) if rid else None
                if not target and rid:
                    target = re.search(rf'Target="([^"]+)"[^>]*Id="{rid.group(1)}"', rels)
                part = "ppt/" + target.group(1).replace("../", "") if target else None
                out["sounds"].append({"name": n, "sndAc": snd.group(0) if snd else None, "part": part,
                                      "bytes": z.getinfo(part).file_size if part and part in names else None})
    finally:
        out["close"] = t.close(kill_if_unreachable=True)
    p = HERE / "results" / f"{out['date']}-powerpoint-sounds.json"
    p.write_text(json.dumps(out, ensure_ascii=False, indent=1), encoding="utf-8", newline="\n")
    print(json.dumps(out, ensure_ascii=False)[:1500])


if __name__ == "__main__":
    main()
