"""Probe: what PowerPoint's object model tells about its defaults (Step 8), in a PowerPoint of our own
without a window (feellab/powerpoint.py: refuses to run while the owner's PowerPoint is open).

- the slide layouts of the default master, for PowerPoint's own default design (Office, Aptos) and
  for the Office 2013-2022 design: every placeholder with its type, position and size in pt and EMU,
  its master text (as far as the layout holds one) and the text sizes per level;
- the slide master's placeholders and its five text levels;
- the designs' theme colours (the 12 slots) and fonts (major/minor, Latin/East Asian/complex);
- the AutoShape types (MsoAutoShapeType) with the name PowerPoint gives a new shape of that type and
  its default adjustment values;
- the slide sizes PowerPoint offers (ppSlideSize*) with their size in pt.

Keyboard shortcuts are not readable through COM. Names and numbers only (D-058).

    .venv/Scripts/python.exe probe_powerpoint_keys.py

Result: results/<date>-powerpoint-layouts.json.
"""
from __future__ import annotations

import datetime
import json
import sys
from pathlib import Path

from feellab.powerpoint import PP_PLACEHOLDER, THEME_2013, OwnerPowerPointRunning, PowerPointTarget, emu, shape_info
from feellab.sanitize import clean

HERE = Path(__file__).parent
OFFICE_LIB = "{2DF8D04C-5BFA-101B-BDE5-00AA0044DE52}"  # Microsoft Office 16.0 Object Library
THEME_COLORS = ["Dark1", "Light1", "Dark2", "Light2", "Accent1", "Accent2", "Accent3", "Accent4", "Accent5", "Accent6",
                "Hyperlink", "FollowedHyperlink"]


def hexrgb(v: int) -> str:
    """PowerPoint's BGR value as #RRGGBB."""
    return f"#{v & 0xFF:02X}{(v >> 8) & 0xFF:02X}{(v >> 16) & 0xFF:02X}"


def placeholder(sh) -> dict:
    info = shape_info(sh)
    out = {"name": info["name"], "type": info.get("placeholder"), "pt": info["pt"], "emu": info["emu"]}
    txt = info.get("text") or {}
    if txt.get("paragraphs"):
        out["text"] = [{"text": p["text"], "level": p["level"], "size": p["size"], "font": p["font"],
                        "bullet": (p.get("bullet") or {}).get("char") if (p.get("bullet") or {}).get("visible") else None}
                       for p in txt["paragraphs"]]
    if txt:
        out["anchor"] = txt.get("anchor")
        out["margins_pt"] = txt.get("margins_pt")
        out["autosize"] = txt.get("autosize")
    return out


def design(pres) -> dict:
    master = pres.SlideMaster
    theme = master.Theme
    cs = theme.ThemeColorScheme
    fs = theme.ThemeFontScheme
    fonts = {}
    for part, getter in (("major", fs.MajorFont), ("minor", fs.MinorFont)):
        fonts[part] = {script: getter(i).Name for i, script in ((1, "latin"), (2, "east_asian"), (3, "complex"))}
    layouts = []
    for i in range(1, master.CustomLayouts.Count + 1):
        lay = master.CustomLayouts(i)
        layouts.append({
            "index": i, "name": lay.Name.replace(chr(13), " ").replace(chr(10), " ").replace(chr(11), " "),
            "placeholders": [placeholder(lay.Shapes.Placeholders(k)) for k in range(1, lay.Shapes.Placeholders.Count + 1)],
            "other_shapes": [lay.Shapes(k).Name for k in range(1, lay.Shapes.Count + 1) if int(lay.Shapes(k).Type) != 14],
        })
    return {
        "design": master.Design.Name,
        "colors": {name: hexrgb(int(cs.Colors(i + 1).RGB)) for i, name in enumerate(THEME_COLORS)},
        "fonts": fonts,
        "master": {"name": master.Name,
                   "placeholders": [placeholder(master.Shapes.Placeholders(k)) for k in range(1, master.Shapes.Placeholders.Count + 1)],
                   "background": hexrgb(int(master.Background.Fill.ForeColor.RGB))},
        "layouts": layouts,
    }


def autoshapes(t: PowerPointTarget, pres, constants: dict) -> list[dict]:
    slide = t.add_slide(pres, 7)
    out = []
    for name, value in sorted(constants.items(), key=lambda kv: kv[1]):
        if value <= 0:
            continue
        entry = {"type": value, "constant": name}
        try:
            sh = slide.Shapes.AddShape(value, 100, 100, 200, 120)
            entry["shape_name"] = str(sh.Name).rsplit(" ", 1)[0]
            adj = sh.Adjustments
            if adj.Count:
                entry["adjustments"] = [round(float(adj.Item(i)), 5) for i in range(1, adj.Count + 1)]
            sh.Delete()
        except Exception as e:  # noqa: BLE001
            entry["error"] = True  # PowerPoint refuses this type (e.g. msoShapeNotPrimitive)
        out.append(entry)
    return out


def main() -> None:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    from win32com.client import gencache
    office = gencache.EnsureModule(OFFICE_LIB, 0, 2, 8)
    try:
        t = PowerPointTarget(visible=False)
    except OwnerPowerPointRunning as e:
        raise SystemExit(f"skipped: {e}")
    try:
        # PowerPoint's own type library: the package the typed Application object was generated into.
        ppt = sys.modules[type(t.app).__module__.rsplit(".", 1)[0]]
        shapes_c = {k: v for k, v in vars(office.constants).items() if k.startswith("msoShape") and not k.startswith("msoShapeStyle") and isinstance(v, int)}
        sizes_c = {k: v for k, v in vars(ppt.constants).items() if k.startswith("ppSlideSize") and isinstance(v, int)}
        data: dict = {"date": datetime.date.today().isoformat(), "built_by": f"PowerPoint {t.app.Version} build {t.app.Build}"}
        p = t.new_presentation(with_window=False, name="keys")
        data["default"] = design(p)
        ps = p.PageSetup
        data["default_slide_size"] = {"type": int(ps.SlideSize), "pt": [float(ps.SlideWidth), float(ps.SlideHeight)],
                                      "emu": [emu(ps.SlideWidth), emu(ps.SlideHeight)]}
        sizes = []
        for name, v in sorted(sizes_c.items(), key=lambda kv: kv[1]):
            try:
                ps.SlideSize = v
                sizes.append({"constant": name, "type": v, "pt": [float(ps.SlideWidth), float(ps.SlideHeight)]})
            except Exception as e:  # noqa: BLE001
                sizes.append({"constant": name, "type": v, "error": True})
        data["slide_sizes"] = sizes
        p2 = t.new_presentation(with_window=False, theme=THEME_2013, name="keys2013")
        data["office_2013_2022"] = design(p2)
        data["autoshapes"] = autoshapes(t, p2, shapes_c)
        data["placeholder_types"] = PP_PLACEHOLDER
    finally:
        print(t.close())
    out = HERE / "results" / f"{datetime.date.today().isoformat()}-powerpoint-layouts.json"
    out.write_text(json.dumps(clean(data), ensure_ascii=False, indent=1), encoding="utf-8", newline="\n")
    print(len(data["default"]["layouts"]), "layouts,", len(data["autoshapes"]), "autoshapes ->", out.name)


if __name__ == "__main__":
    main()
