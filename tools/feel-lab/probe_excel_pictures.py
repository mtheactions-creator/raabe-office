"""Records how German Excel 365 writes and draws a picture's Bildformat settings (D-087), in an Excel of
our own, without real input.

    .venv/Scripts/python.exe probe_excel_pictures.py [formats] [ui]

formats (invisible Excel, COM only): a test picture built here in TEMP is inserted once per case
(Shapes.AddPicture) and given one setting through COM (PictureFormat, Fill.PictureEffects, Line,
Shadow, Glow, SoftEdge, Reflection, ThreeD, AutoShapeType, Rotation, Flip, AlternativeText, ...).
The workbook is saved as the fixture packages/ooxml/test/fixtures/xlsx-pictures.xlsx; each case's
anchor XML, the settings Excel reports back through COM and whether the stored image is still the
original go to packages/ooxml/test/fixtures/excel-pictures.json; Excel's rendering of each case
(Range.CopyPicture, uses the clipboard for a moment) goes to TEMP/raabe-pictures/excel-<case>.png.

ui (our Excel made visible off screen; selecting a tab or invoking a gallery item through UI
Automation brings it to the front for a moment, the focus is given back at once): the Bildformat tab's
controls with their keytips and states, the entries of its menus, and the 28 picture styles of the
Bildformatvorlagen gallery applied one by one (gallery item invoked through UI Automation) and read
back from the file Excel saves. Labels and keytips only (D-058).

Result: results/2026-11-08-excel-pictures.json (labels, keytips, case list, numbers).
"""
from __future__ import annotations

import json
import re
import shutil
import sys
import tempfile
import time
import zipfile
from pathlib import Path

from feellab.excel import ExcelTarget

HERE = Path(__file__).parent
ROOT = HERE.parents[1]
TMP = Path(tempfile.gettempdir())
SHOTS = TMP / "raabe-pictures"
FIXTURE = ROOT / "packages" / "ooxml" / "test" / "fixtures" / "xlsx-pictures.xlsx"
FIXJSON = ROOT / "packages" / "ooxml" / "test" / "fixtures" / "excel-pictures.json"
RESULT = HERE / "results" / "2026-11-08-excel-pictures.json"
PNG = TMP / "raabe-picture-test.png"
STYLES_FILE = TMP / "raabe-picture-styles.xlsx"

# Block of cells per case (columns x rows) and cases per band.
BW, BH, PER_ROW = 4, 10, 5
PIC_W, PIC_H = 96.0, 64.0


def make_png(path: Path) -> None:
    """A 240 x 160 test picture: four coloured quadrants, a grey ramp and a white field (for
    Transparente Farbe), so that colour settings and crops are visible."""
    from PIL import Image, ImageDraw
    img = Image.new("RGB", (240, 160), (255, 255, 255))
    d = ImageDraw.Draw(img)
    d.rectangle([0, 0, 119, 79], fill=(200, 40, 40))
    d.rectangle([120, 0, 239, 79], fill=(40, 150, 60))
    d.rectangle([0, 80, 119, 139], fill=(40, 80, 200))
    d.rectangle([120, 80, 179, 139], fill=(240, 200, 40))
    for x in range(240):
        v = int(x * 255 / 239)
        d.line([x, 140, x, 159], fill=(v, v, v))
    img.save(path)


def rgb(r: int, g: int, b: int) -> int:
    return r | (g << 8) | (b << 16)


def _transparency(s) -> None:
    # Transparenz of a picture (Excel 365): the fill's transparency is the picture's own.
    s.Fill.Transparency = 0.5


def _effect(kind: int, values: list[float] | None = None):
    def run(s) -> None:
        fx = s.Fill.PictureEffects.Insert(kind, -1)
        for i, v in enumerate(values or [], start=1):
            fx.EffectParameters(i).Value = v
    return run


def _shadow(style: int):
    def run(s) -> None:
        sh = s.Shadow
        sh.Visible = -1
        sh.Style = style
        sh.ForeColor.RGB = 0
        sh.OffsetX = 4
        sh.OffsetY = 4
        sh.Blur = 6
        sh.Transparency = 0.6
    return run


def _line(s) -> None:
    s.Line.Visible = -1
    s.Line.ForeColor.RGB = rgb(192, 0, 0)
    s.Line.Weight = 4.5
    s.Line.DashStyle = 4


def _glow(s) -> None:
    s.Glow.Radius = 8
    s.Glow.Color.RGB = rgb(68, 114, 196)
    s.Glow.Transparency = 0.6


def _bevel(s) -> None:
    s.ThreeD.BevelTopType = 3
    s.ThreeD.BevelTopInset = 6
    s.ThreeD.BevelTopDepth = 6


def _rot3d(s) -> None:
    s.ThreeD.RotationX = 30
    s.ThreeD.RotationY = 20


def _crop(s) -> None:
    pf = s.PictureFormat
    pf.CropLeft = 10
    pf.CropTop = 8
    pf.CropRight = 20
    pf.CropBottom = 4


def _style_set(s) -> None:
    pf = s.PictureFormat
    pf.Brightness = 0.7
    pf.Contrast = 0.4


def _tcolor(s) -> None:
    s.PictureFormat.TransparentBackground = -1
    s.PictureFormat.TransparencyColor = rgb(255, 255, 255)


def _alt(s) -> None:
    s.AlternativeText = "Ein Testbild mit vier Farben"
    s.Title = "Testbild"


CASES: list[tuple[str, object]] = [
    ("plain", None),
    ("crop", _crop),
    ("lum", _style_set),
    ("grayscale", lambda s: setattr(s.PictureFormat, "ColorType", 2)),
    ("bw", lambda s: setattr(s.PictureFormat, "ColorType", 3)),
    ("washout", lambda s: setattr(s.PictureFormat, "ColorType", 4)),
    ("transparentColor", _tcolor),
    ("transparency", _transparency),
    ("brightnessContrast", _effect(3, [0.2, -0.2])),
    ("saturation", _effect(24, [2.0])),
    ("temperature", _effect(7, [8800])),
    ("sharpen", _effect(25, [0.5])),
    ("soften", _effect(25, [-0.5])),
    ("artistic", _effect(21)),
    ("rotate30", lambda s: setattr(s, "Rotation", 30)),
    ("flipH", lambda s: s.Flip(0)),
    ("flipV", lambda s: s.Flip(1)),
    ("line", _line),
    ("shadowOuter", _shadow(2)),
    ("shadowInner", _shadow(1)),
    ("glow", _glow),
    ("softEdge", lambda s: setattr(s.SoftEdge, "Radius", 6)),
    ("reflection", lambda s: setattr(s.Reflection, "Type", 2)),
    ("bevel", _bevel),
    ("rotation3d", _rot3d),
    ("oval", lambda s: setattr(s, "AutoShapeType", 9)),
    ("roundRect", lambda s: setattr(s, "AutoShapeType", 5)),
    ("altText", _alt),
    ("noAspect", lambda s: setattr(s, "LockAspectRatio", 0)),
    ("hidden", lambda s: setattr(s, "Visible", 0)),
]


def strip_path(path: Path) -> None:
    """Excel records the folder it saved to (x15ac:absPath in workbook.xml): left out of the fixture."""
    with zipfile.ZipFile(path) as z:
        items = [(info, z.read(info)) for info in z.infolist()]
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        for info, data in items:
            if info.filename == "xl/workbook.xml":
                data = re.sub(rb"<mc:AlternateContent\b(?:(?!</mc:AlternateContent>)[\s\S])*?x15ac:absPath[\s\S]*?</mc:AlternateContent>", b"", data)
            z.writestr(info, data)


def office_anonymize(path: Path) -> None:
    """The Office user's name in the saved file (comments, people …) replaced by the harness's
    (anonymize_office.py); Excel's own user name is never changed (it is shared by all Office apps)."""
    import subprocess
    subprocess.run([sys.executable, str(HERE / "anonymize_office.py"), str(path)], check=True, capture_output=True)


def set_author(wb, name: str) -> None:
    """Author fields of the workbook (anonymize() rewrites core.xml afterwards as well)."""
    for prop in ("Author", "Last author"):
        try:
            wb.BuiltinDocumentProperties.Item(prop).Value = name
        except Exception:  # noqa: BLE001
            pass


def block(i: int) -> tuple[int, int]:
    """Top-left cell (row, column, 1-based) of case i's block."""
    return 1 + (i // PER_ROW) * BH, 1 + (i % PER_ROW) * BW


def block_range(ws, i: int):
    r, c = block(i)
    return ws.Range(ws.Cells(r, c), ws.Cells(r + BH - 1, c + BW - 1))


def readback(s) -> dict:
    """What Excel reports for a picture through COM (the round trip checks the same properties)."""
    out: dict = {}

    def get(name: str, fn) -> None:
        try:
            v = fn()
            out[name] = round(v, 3) if isinstance(v, float) else v
        except Exception:  # noqa: BLE001
            out[name] = None
    pf = s.PictureFormat
    get("left", lambda: s.Left)
    get("top", lambda: s.Top)
    get("width", lambda: s.Width)
    get("height", lambda: s.Height)
    get("rotation", lambda: s.Rotation)
    get("flipH", lambda: s.HorizontalFlip)
    get("flipV", lambda: s.VerticalFlip)
    get("cropLeft", lambda: pf.CropLeft)
    get("cropTop", lambda: pf.CropTop)
    get("cropRight", lambda: pf.CropRight)
    get("cropBottom", lambda: pf.CropBottom)
    get("brightness", lambda: pf.Brightness)
    get("contrast", lambda: pf.Contrast)
    get("colorType", lambda: pf.ColorType)
    get("transparentBackground", lambda: pf.TransparentBackground)
    get("fillTransparency", lambda: s.Fill.Transparency)
    get("effects", lambda: [s.Fill.PictureEffects(i).Type for i in range(1, s.Fill.PictureEffects.Count + 1)])
    get("lineVisible", lambda: s.Line.Visible)
    get("lineWeight", lambda: s.Line.Weight)
    get("lineColor", lambda: s.Line.ForeColor.RGB)
    get("lineDash", lambda: s.Line.DashStyle)
    get("shadowVisible", lambda: s.Shadow.Visible)
    get("shadowStyle", lambda: s.Shadow.Style)
    get("shadowBlur", lambda: s.Shadow.Blur)
    get("glowRadius", lambda: s.Glow.Radius)
    get("softEdge", lambda: s.SoftEdge.Radius)
    get("reflection", lambda: s.Reflection.Type)
    get("bevel", lambda: s.ThreeD.BevelTopType)
    get("rotationX", lambda: s.ThreeD.RotationX)
    get("autoShapeType", lambda: s.AutoShapeType)
    get("alt", lambda: s.AlternativeText)
    get("title", lambda: s.Title)
    get("lockAspect", lambda: s.LockAspectRatio)
    get("visible", lambda: s.Visible)
    get("name", lambda: s.Name)
    get("zorder", lambda: s.ZOrderPosition)
    return out


def anchors(path: Path) -> tuple[list[str], dict]:
    """The drawing's anchors in order and the drawing's media (name -> bytes)."""
    with zipfile.ZipFile(path) as z:
        xml = z.read("xl/drawings/drawing1.xml").decode("utf-8")
        rels = z.read("xl/drawings/_rels/drawing1.xml.rels").decode("utf-8")
        media = {n: z.read(n) for n in z.namelist() if n.startswith("xl/media/")}
    out = re.findall(r"<xdr:(?:twoCellAnchor|oneCellAnchor|absoluteAnchor)\b[\s\S]*?</xdr:(?:twoCellAnchor|oneCellAnchor|absoluteAnchor)>", xml)
    targets = dict(re.findall(r'Id="([^"]+)"[^>]*Target="([^"]+)"', rels))
    targets.update({k: v for v, k in re.findall(r'Target="([^"]+)"[^>]*Id="([^"]+)"', rels)})
    return out, {"media": media, "targets": targets, "head": xml[: xml.index("<xdr:", xml.index("<xdr:wsDr") + 5)]}


def formats() -> dict:
    from make_xlsx_fixtures import NEUTRAL, anonymize, check_private
    make_png(PNG)
    SHOTS.mkdir(exist_ok=True)
    x = ExcelTarget(visible=False)
    user = ""
    cases = []
    try:
        user = str(x.app.UserName)
        ws = x.wb.Worksheets(1)
        ws.Name = "Bilder"
        for i, (name, fn) in enumerate(CASES):
            r, c = block(i)
            cell = ws.Cells(r + 1, c)
            s = ws.Shapes.AddPicture(str(PNG), 0, -1, cell.Left + 8, cell.Top + 4, PIC_W, PIC_H)
            s.Name = f"Bild {name}"
            err = None
            try:
                if fn:
                    fn(s)
            except Exception as e:  # noqa: BLE001
                err = str(e)[:160]
            ws.Cells(r, c).Value = name
            cases.append({"case": name, "error": err, "com": readback(s)})
            print(name, "error:" if err else "ok", err or "", flush=True)
        # Gruppieren: a picture and a rectangle as one group (the last block).
        i = len(CASES)
        r, c = block(i)
        cell = ws.Cells(r + 1, c)
        p = ws.Shapes.AddPicture(str(PNG), 0, -1, cell.Left + 8, cell.Top + 4, PIC_W, PIC_H)
        p.Name = "Bild group"
        rect = ws.Shapes.AddShape(1, cell.Left + 8 + PIC_W + 6, cell.Top + 4, 40, 30)
        rect.Name = "Rechteck group"
        g = ws.Shapes.Range([p.Name, rect.Name]).Group()
        g.Name = "Gruppe group"
        ws.Cells(r, c).Value = "group"
        cases.append({"case": "group", "error": None, "com": readback(g)})
        # Excel's rendering of each block.
        from PIL import Image, ImageGrab
        for i, cs in enumerate(cases):
            img = None
            for _ in range(8):
                try:
                    block_range(ws, i).CopyPicture(1, 2)
                    time.sleep(0.3)
                    img = ImageGrab.grabclipboard()
                    if isinstance(img, Image.Image):
                        break
                except Exception:  # noqa: BLE001
                    time.sleep(0.4)
            if isinstance(img, Image.Image):
                img.convert("RGB").save(SHOTS / f"excel-{cs['case']}.png")
        set_author(x.wb, NEUTRAL)
        out = TMP / "raabe-pictures-excel.xlsx"
        if out.exists():
            out.unlink()
        x.wb.SaveAs(str(out), 51)
    finally:
        x.close()
    strip_path(out)
    anonymize(out)
    office_anonymize(out)
    check_private(out, user)
    shutil.copyfile(out, FIXTURE)
    xs, info = anchors(out)
    orig = PNG.read_bytes()
    by_name = {}
    for a in xs:
        m = re.search(r'<xdr:cNvPr\b[^>]*name="([^"]*)"', a)
        if m:
            by_name[m.group(1)] = a
    for cs in cases:
        key = "Gruppe group" if cs["case"] == "group" else f"Bild {cs['case']}"
        a = by_name.get(key)
        cs["anchor"] = a
        if a:
            embeds = re.findall(r'r:embed="([^"]+)"', a)
            cs["media"] = []
            for rid in embeds:
                t = info["targets"].get(rid, "")
                part = "xl/" + t.replace("../", "")
                data = info["media"].get(part, b"")
                cs["media"].append({"rid": rid, "ext": part.rsplit(".", 1)[-1], "size": len(data), "original": data == orig})
    FIXJSON.write_text(json.dumps({"picture": [240, 160], "size": [PIC_W, PIC_H], "cases": cases}, ensure_ascii=False, indent=1) + "\n",
                       encoding="utf-8", newline="\n")
    print("fixture written:", len(cases), "cases")
    return {"cases": [c["case"] for c in cases], "errors": {c["case"]: c["error"] for c in cases if c["error"]}}


# ----- UI Automation helpers (our Excel only) ---------------------------------------------------------

_fg0: list[int] = []


def give_back() -> None:
    """Selecting or invoking through UI Automation brings our Excel to the front: the focus goes back at once."""
    from feellab import winapi
    if not _fg0:
        return
    if _fg0[0] and winapi.foreground() != _fg0[0]:
        try:
            winapi.activate(_fg0[0], 0)
        except Exception:  # noqa: BLE001
            pass


def descendants(root) -> list[tuple]:
    from feellab import uia
    out = []
    f = root.FindAll(uia.U.TreeScope_Descendants, uia._uia.CreateTrueCondition())
    for i in range(f.Length):
        e = f.GetElement(i)
        try:
            out.append((e, uia.describe(e)))
        except Exception:  # noqa: BLE001
            continue
    return out


def _popups(pid: int) -> list[tuple]:
    """Every element of our Excel's top-level windows other than the main one (menus, gallery popups)."""
    from feellab import uia
    out = []
    root = uia._uia.GetRootElement()
    cond = uia._uia.CreatePropertyCondition(uia.U.UIA_ProcessIdPropertyId, pid)
    f = root.FindAll(uia.U.TreeScope_Children, cond)
    for i in range(f.Length):
        out.extend(descendants(f.GetElement(i)))
    return out


def expand(el) -> None:
    from feellab import uia
    p = el.GetCurrentPattern(uia.U.UIA_ExpandCollapsePatternId)
    p.QueryInterface(uia.U.IUIAutomationExpandCollapsePattern).Expand()
    give_back()


def collapse(el) -> None:
    from feellab import uia
    try:
        p = el.GetCurrentPattern(uia.U.UIA_ExpandCollapsePatternId)
        p.QueryInterface(uia.U.IUIAutomationExpandCollapsePattern).Collapse()
    except Exception:  # noqa: BLE001
        pass
    give_back()


def menu_items(root, pid: int) -> list[dict]:
    """The entries of the open menus (below the ribbon), each once."""
    out, seen = [], set()
    for _, d in descendants(root) + _popups(pid):
        if d["rect"][1] <= 268 or not d["name"] or d["type"] not in ("MenuItem", "Button", "CheckBox", "RadioButton", "ListItem", "Group"):
            continue
        k = (d["name"], tuple(d["rect"]))
        if k in seen:
            continue
        seen.add(k)
        out.append({k2: d[k2] for k2 in ("type", "name", "key", "enabled")})
    return out


# ----- the Bildformat tab through UI Automation ------------------------------------------------------

def ui() -> dict:
    from feellab import uia, winapi
    from make_xlsx_fixtures import NEUTRAL, anonymize, check_private
    make_png(PNG)
    result: dict = {}
    x = ExcelTarget(visible=False)
    user = ""
    try:
        user = str(x.app.UserName)
        ws = x.wb.Worksheets(1)
        pics = []
        for i in range(28):
            r, c = block(i)
            cell = ws.Cells(r + 1, c)
            s = ws.Shapes.AddPicture(str(PNG), 0, -1, cell.Left + 8, cell.Top + 4, PIC_W, PIC_H)
            s.Name = f"Stil {i}"
            pics.append(s)
        winapi.place_quietly(x.hwnd, -4000, 0, 2600, 1300)
        x.app.Visible = True
        time.sleep(1.0)
        give_back()
        # The contextual tab appears only for a selection made while the window is shown.
        pics[0].Select()
        time.sleep(1.5)
        root = uia.element_for(x.hwnd)
        tab = next(e for e, d in descendants(root) if d["type"] == "TabItem" and d["name"] == "Bildformat")
        result["tab"] = {"name": "Bildformat", "key": uia.describe(tab)["key"]}
        uia.select(tab)
        give_back()
        time.sleep(1.2)
        ribbon = [d for _, d in descendants(root) if 74 < d["rect"][1] < 268 and d["name"]]
        result["controls"] = [{k: d[k] for k in ("type", "name", "key", "enabled")} for d in ribbon if d["type"] != "ListItem"]
        result["groups"] = [d["name"] for d in ribbon if d["type"] in ("Group", "ToolBar") and d["name"]]
        print("controls", len(result["controls"]), flush=True)
        # Menus: every menu button and split button half of the tab, one level.
        result["menus"] = {}
        els = [(e, d) for e, d in descendants(root) if 74 < d["rect"][1] < 268 and d["name"]]
        anchors_ = [(e, d) for e, d in els if d["type"] in ("MenuItem", "SplitButton") or (d["type"] == "Button" and d["name"] == "Weitere Optionen")]
        seen_names = set()
        for e, d in anchors_:
            label = d["name"]
            if d["type"] == "Button":
                # A split button's arrow: name it after the split button it is in.
                ax, ay = d["rect"][0], d["rect"][1]
                owner = next((dd["name"] for _, dd in els if dd["type"] == "SplitButton" and dd["rect"][0] <= ax < dd["rect"][0] + dd["rect"][2]
                              and dd["rect"][1] <= ay < dd["rect"][1] + dd["rect"][3]), None)
                if not owner:
                    continue
                label = owner
            elif d["type"] == "SplitButton":
                continue
            if label in seen_names:
                continue
            seen_names.add(label)
            try:
                expand(e)
                time.sleep(0.8)
                result["menus"][label] = menu_items(root, x.pid)
                print("menu", label, len(result["menus"][label]), flush=True)
            except Exception as ex:  # noqa: BLE001
                result["menus"][label] = {"error": str(ex)[:100]}
            finally:
                collapse(e)
                time.sleep(0.4)
        # Bildformatvorlagen: apply style i to picture i.
        names: list[str] = []
        for i in range(28):
            pics[i].Select()
            time.sleep(0.5)
            more = next((e for e, d in descendants(root) if d["type"] == "Button" and d["name"] in ("Weitere", "Bildformatvorlagen", "Schnellformatvorlagen")
                         and 74 < d["rect"][1] < 268), None)
            items = []
            if more is not None:
                uia.invoke(more)
                give_back()
                time.sleep(0.9)
                items = [(e, d) for e, d in _popups(x.pid) + descendants(root) if d["type"] == "ListItem" and d["name"]]
            uniq = []
            for e, d in items:
                if d["name"] not in [u[1]["name"] for u in uniq]:
                    uniq.append((e, d))
            if not names:
                names = [d["name"] for _, d in uniq]
                print(len(names), "styles", names[:3], flush=True)
            if i >= len(names):
                break
            target = next((e for e, d in uniq if d["name"] == names[i]), None)
            if target is None:
                continue
            try:
                uia.invoke(target)
            except Exception:  # noqa: BLE001
                uia.select(target)
            give_back()
            time.sleep(0.5)
        result["styles"] = names
        x.app.Visible = False
        set_author(x.wb, NEUTRAL)
        if STYLES_FILE.exists():
            STYLES_FILE.unlink()
        x.wb.SaveAs(str(STYLES_FILE), 51)
    finally:
        give_back()
        x.close()
        give_back()
    strip_path(STYLES_FILE)
    anonymize(STYLES_FILE)
    office_anonymize(STYLES_FILE)
    check_private(STYLES_FILE, user)
    xs, _ = anchors(STYLES_FILE)
    spprs = []
    for a in xs:
        m = re.search(r'name="Stil (\d+)"', a)
        sp = re.search(r"<xdr:spPr\b[^>]*>([\s\S]*?)</xdr:spPr>", a)
        if m and sp:
            spprs.append({"i": int(m.group(1)), "spPr": re.sub(r"<a:xfrm\b[\s\S]*?</a:xfrm>", "", sp.group(1)),
                          "blipFill": re.search(r"<xdr:blipFill\b[\s\S]*?</xdr:blipFill>", a).group(0)})
    spprs.sort(key=lambda s: s["i"])
    fixture = json.loads(FIXJSON.read_text(encoding="utf-8")) if FIXJSON.exists() else {}
    fixture["styles"] = [{"name": names[s["i"]] if s["i"] < len(names) else None, **s} for s in spprs]
    FIXJSON.write_text(json.dumps(fixture, ensure_ascii=False, indent=1) + "\n", encoding="utf-8", newline="\n")
    return result


MENU_ANCHORS = ["Zuschneiden", "Bildrahmen", "Ebene nach vorne", "Ebene nach hinten", "Drehen", "Ausrichten", "Bildeffekte", "Bild zurücksetzen", "Bild ändern"]
SUBMENUS = {"Zuschneiden": ["Auf Form zuschneiden", "Seitenverhältnis"], "Bildrahmen": ["Stärke", "Striche"],
            "Bildeffekte": ["Voreinstellung", "Schatten", "Spiegelung", "Leuchteffekt", "Weiche Kanten", "Abschrägung", "3D-Drehung"]}


MAIN_NAMES: set[str] = set()


def menus() -> dict:
    """The menus of the tab, read from the popup windows only (the first pass also saw the sheet)."""
    from feellab import uia, winapi
    make_png(PNG)
    out: dict = {}
    x = ExcelTarget(visible=False)
    try:
        ws = x.wb.Worksheets(1)
        s = ws.Shapes.AddPicture(str(PNG), 0, -1, 40, 40, PIC_W, PIC_H)
        winapi.place_quietly(x.hwnd, -4000, 0, 2600, 1300)
        x.app.Visible = True
        time.sleep(1.0)
        give_back()
        s.Select()
        time.sleep(1.5)
        root = uia.element_for(x.hwnd)
        tab = next(e for e, d in descendants(root) if d["type"] == "TabItem" and d["name"] == "Bildformat")
        uia.select(tab)
        give_back()
        time.sleep(1.2)
        MAIN_NAMES.update(d["name"] for _, d in _popups(x.pid))

        def popup_items() -> list[dict]:
            seen, items = set(), []
            for _, d in _popups(x.pid):
                if not d["name"] or d["type"] not in ("MenuItem", "Button", "CheckBox", "RadioButton", "ListItem"):
                    continue
                k = (d["name"], tuple(d["rect"]))
                if k in seen:
                    continue
                seen.add(k)
                items.append({k2: d[k2] for k2 in ("type", "name", "key", "enabled")})
            return items

        for anchor in MENU_ANCHORS:
            els = [(e, d) for e, d in descendants(root) if 74 < d["rect"][1] < 268]
            a = next(((e, d) for e, d in els if d["name"] == anchor and d["type"] in ("SplitButton", "MenuItem")), None)
            if a is None:
                continue
            el, ad = a
            if ad["type"] == "SplitButton":
                ax, ay, aw, ah = ad["rect"]
                el = next((e for e, d in els if d["name"] == "Weitere Optionen" and ax <= d["rect"][0] < ax + aw and ay <= d["rect"][1] < ay + ah), None)
                if el is None:
                    continue
            try:
                # The menu stays open only while our Excel has the focus: read it, then give the focus back.
                p_ = el.GetCurrentPattern(uia.U.UIA_ExpandCollapsePatternId)
                p_.QueryInterface(uia.U.IUIAutomationExpandCollapsePattern).Expand()
                time.sleep(0.7)
                base = {(d["name"], tuple(d["rect"])) for _, d in descendants(root)}
                out[anchor] = [i for i in popup_items()]
                out[anchor] = [i for i in out[anchor] if i["name"] not in MAIN_NAMES]
                for sub in SUBMENUS.get(anchor, []):
                    se = next((e for e, d in _popups(x.pid) if d["name"] == sub and d["type"] == "MenuItem"), None)
                    if se is None:
                        continue
                    before = {(i["name"]) for i in out[anchor]}
                    try:
                        expand(se)
                        time.sleep(0.8)
                        out[f"{anchor} > {sub}"] = [i for i in popup_items() if i["name"] not in before]
                    finally:
                        collapse(se)
                        time.sleep(0.3)
                print("menu", anchor, len(out[anchor]), flush=True)
            except Exception as ex:  # noqa: BLE001
                out[anchor] = {"error": str(ex)[:100]}
            finally:
                collapse(el)
                time.sleep(0.4)
        x.app.Visible = False
    finally:
        give_back()
        x.close()
        give_back()
    return out


def main() -> None:
    from feellab import winapi
    sys.stdout.reconfigure(encoding="utf-8")
    _fg0.append(winapi.foreground())
    what = sys.argv[1:] or ["formats", "ui", "menus"]
    res = json.loads(RESULT.read_text(encoding="utf-8")) if RESULT.exists() else {}
    if "formats" in what:
        res["formats"] = formats()
    if "ui" in what:
        res["ui"] = ui()
    if "menus" in what:
        res["menus"] = menus()
    from feellab.sanitize import clean
    RESULT.write_text(json.dumps(clean(res), ensure_ascii=False, indent=1) + "\n", encoding="utf-8", newline="\n")
    print("written", RESULT.name)


if __name__ == "__main__":
    main()
