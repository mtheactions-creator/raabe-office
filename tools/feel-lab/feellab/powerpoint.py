"""Microsoft PowerPoint as the reference for RaabePresent (Step 8), driven through COM.

PowerPoint runs as ONE process per user: `DispatchEx` does not start a second one but reaches
the PowerPoint that is already running, with the owner's presentations in it. So the harness
follows the owner's rules for PowerPoint (Step 8, CLAUDE.md harness safety):

- Before anything, `owner_powerpoint_running()` looks for POWERPNT.EXE. If one runs, it is the
  owner's: it is not attached to, read or closed; `PowerPointTarget()` refuses to start
  (`OwnerPowerPointRunning`) and the probe does other work.
- Otherwise the target starts PowerPoint itself and works only on presentations it created; they
  are saved into a TEMP folder of its own at once. `close()` closes exactly those, and quits
  PowerPoint only if no other presentation is open then (the owner may have started PowerPoint in
  the meantime: their files land in our process). If PowerPoint lingers after Quit with nothing
  open, only the process we started (its PID) is ended.
- Options a probe changes (the reduced ribbon, AutoKorrektur settings ...) are put back.
- Windowless presentations (`with_window=False`) keep everything invisible; real-input runs
  open a window and place it at WINDOW_RECT, as `excel.py` does.

Read-back: shapes with name, type, position and size in points and EMU (12700 per pt),
rotation, flips, text with paragraphs and line positions, placeholder type, z-order; the
selection of the active window; slides with layout, notes, section, transition and animations.
"""

from __future__ import annotations

import subprocess
import tempfile
import time
from pathlib import Path

from . import winapi

WINDOW_RECT = (0, 0, 2000, 2050)
EMU_PER_PT = 12700
THEMES = Path(r"C:\Program Files\Microsoft Office\root\Document Themes 16")
THEME_2013 = THEMES / "Office 2013 - 2022 Theme.thmx"
THEME_NEW = THEMES / "Office Theme.thmx"

# Enumerations (Office VBA reference), for readable read-backs.
MSO_SHAPE_TYPE = {1: "AutoShape", 2: "Callout", 3: "Chart", 4: "Comment", 5: "Freeform", 6: "Group", 7: "EmbeddedOLEObject",
                  8: "FormControl", 9: "Line", 10: "LinkedOLEObject", 11: "LinkedPicture", 12: "OLEControlObject", 13: "Picture",
                  14: "Placeholder", 15: "TextEffect", 16: "Media", 17: "TextBox", 18: "ScriptAnchor", 19: "Table", 20: "Canvas",
                  21: "Diagram", 22: "Ink", 23: "InkComment", 24: "SmartArt", 25: "Slicer", 26: "WebVideo", 27: "ContentApp",
                  28: "Graphic", 29: "LinkedGraphic", 30: "3DModel", 31: "Linked3DModel"}
PP_PLACEHOLDER = {1: "Title", 2: "Body", 3: "CenterTitle", 4: "Subtitle", 5: "VerticalTitle", 6: "VerticalBody", 7: "Object",
                  8: "Chart", 9: "Bitmap", 10: "MediaClip", 11: "OrgChart", 12: "Table", 13: "SlideNumber", 14: "Header",
                  15: "Footer", 16: "Date", 17: "VerticalObject", 18: "Picture", 19: "Cameo", 0: "Mixed"}
PP_SELECTION = {0: "None", 1: "Slides", 2: "Shapes", 3: "Text"}
PP_LAYOUT = {1: "Title", 2: "Text", 3: "TwoColumnText", 4: "Table", 5: "TextAndChart", 6: "ChartAndText", 7: "OrgChart",
             8: "Chart", 9: "TextAndClipArt", 10: "ClipArtAndText", 11: "TitleOnly", 12: "Blank", 13: "TextAndObject",
             14: "ObjectAndText", 15: "LargeObject", 16: "Object", 17: "TextAndMediaClip", 18: "MediaClipAndText",
             19: "ObjectOverText", 20: "TextOverObject", 21: "TextAndTwoObjects", 22: "TwoObjectsAndText",
             23: "TwoObjectsOverText", 24: "FourObjects", 25: "VerticalText", 26: "ClipArtAndVerticalText",
             27: "VerticalTitleAndText", 28: "VerticalTitleAndTextOverChart", 29: "TwoObjects", 30: "ObjectAndTwoObjects",
             31: "TwoObjectsAndObject", 32: "Custom", 33: "SectionHeader", 34: "Comparison", 35: "ContentWithCaption",
             36: "PictureWithCaption"}
AUTOSIZE = {0: "None", 1: "ShapeToFitText", 2: "Mixed", 3: "TextToFitShape"}  # MsoAutoSize (TextFrame2)
PP_SAVE_PPTX, PP_SAVE_PDF = 24, 32


class OwnerPowerPointRunning(RuntimeError):
    """A POWERPNT.EXE runs that this harness did not start: it is the owner's, hands off."""


def powerpoint_pids() -> list[int]:
    out = subprocess.run(["tasklist", "/FI", "IMAGENAME eq POWERPNT.EXE", "/FO", "CSV", "/NH"],
                         capture_output=True, text=True, errors="replace").stdout
    pids = []
    for line in out.splitlines():
        parts = [p.strip('"') for p in line.split('","')]
        if len(parts) > 1 and parts[0].upper() == "POWERPNT.EXE" and parts[1].isdigit():
            pids.append(int(parts[1]))
    return pids


def owner_powerpoint_running(own: tuple[int, ...] = ()) -> bool:
    """True when a POWERPNT.EXE runs that is not one of `own` (PIDs this harness started)."""
    return any(p not in own for p in powerpoint_pids())


def emu(pt: float) -> int:
    return int(round(pt * EMU_PER_PT))


def _try(fn, default=None):
    try:
        return fn()
    except Exception:  # noqa: BLE001
        return default


class PowerPointTarget:
    name = "powerpoint"

    def __init__(self, visible: bool = False, place: bool = True):
        if owner_powerpoint_running():
            raise OwnerPowerPointRunning("POWERPNT.EXE is running (the owner's): no PowerPoint probe now")
        import win32com.client
        self.app = win32com.client.gencache.EnsureDispatch(win32com.client.DispatchEx("PowerPoint.Application"))
        pids = powerpoint_pids()
        self.pid = pids[0] if len(pids) == 1 else 0
        self.visible = visible
        self.place = place
        self.own: list = []
        self.tmp = Path(tempfile.mkdtemp(prefix="raabe-ppt-"))
        self._options: dict[str, tuple] = {}
        self._ribbon_minimized = None
        self.hwnd = 0
        if self.app.Presentations.Count:
            # Someone else's presentation is in this process (cannot happen after the guard,
            # unless the owner started PowerPoint at the same moment): touch nothing.
            self.app = None
            raise OwnerPowerPointRunning("PowerPoint already holds presentations that are not ours")

    # ----- presentations ------------------------------------------------------------------

    def new_presentation(self, with_window: bool | None = None, theme: Path | str | None = None, name: str = "probe"):
        """A new presentation from PowerPoint's default (16:9), saved at once into our TEMP folder."""
        win = self.visible if with_window is None else with_window
        pres = self.app.Presentations.Add(-1 if win else 0)
        self.own.append(pres)
        if theme:
            pres.ApplyTheme(str(theme))
        pres.SaveAs(str(self.tmp / f"{name}-{len(self.own)}.pptx"), PP_SAVE_PPTX)
        if win:
            self.app.Visible = -1
            self.hwnd = self.main_window()
            if self.place and self.hwnd:
                self.app.WindowState = 1  # ppWindowNormal
                winapi.place(self.hwnd, *WINDOW_RECT)
        return pres

    def windows(self) -> list[int]:
        """Visible top-level windows of our PowerPoint process."""
        return winapi.find_windows(lambda h: winapi.window_pid(h) == self.pid) if self.pid else []

    def main_window(self, timeout: float = 10.0) -> int:
        """Our PowerPoint's frame window (class PPTFrameClass); COM's Application.HWND is not in the type library."""
        end = time.time() + timeout
        while time.time() < end:
            hits = [h for h in self.windows() if winapi.window_class(h) == "PPTFrameClass"]
            if hits:
                return hits[0]
            time.sleep(0.2)
        return 0

    def open_copy(self, path: Path | str, with_window: bool | None = None):
        """Opens a copy of a file in our TEMP folder (never the file itself)."""
        import shutil
        win = self.visible if with_window is None else with_window
        dst = self.tmp / f"copy-{len(self.own)}-{Path(path).name}"
        shutil.copyfile(path, dst)
        pres = self.app.Presentations.Open(str(dst), 0, 0, -1 if win else 0)
        self.own.append(pres)
        return pres

    @staticmethod
    def layout(pres, name_or_index):
        """A custom layout of the first master by its name or 1-based index."""
        lays = pres.SlideMaster.CustomLayouts
        if isinstance(name_or_index, int):
            return lays(name_or_index)
        for i in range(1, lays.Count + 1):
            if lays(i).Name == name_or_index:
                return lays(i)
        raise KeyError(name_or_index)

    def add_slide(self, pres, layout=1, index: int | None = None):
        idx = index or pres.Slides.Count + 1
        return pres.Slides.AddSlide(idx, self.layout(pres, layout))

    @staticmethod
    def set_properties(pres, author: str = "Raabe Office") -> None:
        for prop in ("Author", "Last author"):
            _try(lambda p=prop: setattr(pres.BuiltInDocumentProperties(p), "Value", author))

    def save_copy(self, pres, path: Path | str) -> None:
        """Writes the presentation as .pptx elsewhere; the working copy stays in TEMP."""
        path = Path(path)
        if path.exists():
            path.unlink()
        pres.SaveCopyAs(str(path), PP_SAVE_PPTX)

    def export_pdf(self, pres, path: Path | str) -> None:
        path = Path(path)
        if path.exists():
            path.unlink()
        try:
            pres.SaveCopyAs(str(path), PP_SAVE_PDF)
        except Exception:  # noqa: BLE001
            pres.ExportAsFixedFormat(str(path), 2)  # ppFixedFormatTypePDF

    @staticmethod
    def export_png(pres, index: int, path: Path | str, width: int = 1280) -> None:
        w, h = pres.PageSetup.SlideWidth, pres.PageSetup.SlideHeight
        pres.Slides(index).Export(str(path), "PNG", width, int(round(width * h / w)))

    # ----- options ------------------------------------------------------------------------

    def set_option(self, obj_path: str, value) -> None:
        """Sets an application option such as "AutoCorrect.DisplayAutoCorrectOptions"; close() restores it."""
        obj = self.app
        *parts, last = obj_path.split(".")
        for p in parts:
            obj = getattr(obj, p)
        if obj_path not in self._options:
            self._options[obj_path] = (obj, last, getattr(obj, last))
        setattr(obj, last, value)

    def remember_ribbon(self) -> None:
        """The reduced ribbon outlives our instance: note it before a probe may change it."""
        if self._ribbon_minimized is None:
            self._ribbon_minimized = _try(lambda: bool(self.app.CommandBars.GetPressedMso("MinimizeRibbon")))

    def restore_ribbon(self) -> None:
        if self._ribbon_minimized is None:
            return
        now = _try(lambda: bool(self.app.CommandBars.GetPressedMso("MinimizeRibbon")))
        if now is not None and now != self._ribbon_minimized:
            _try(lambda: self.app.CommandBars.ExecuteMso("MinimizeRibbon"))

    # ----- read-back ----------------------------------------------------------------------

    def selection(self) -> dict:
        win = _try(lambda: self.app.ActiveWindow)
        if win is None:
            return {"type": "None"}
        sel = win.Selection
        t = _try(lambda: int(sel.Type), 0)
        out: dict = {"type": PP_SELECTION.get(t, str(t))}
        if t in (2, 3):
            sr = _try(lambda: sel.ShapeRange)
            if sr is not None:
                out["shapes"] = [sr(i).Name for i in range(1, sr.Count + 1)]
            if _try(lambda: bool(sel.HasChildShapeRange), False):
                cr = sel.ChildShapeRange
                out["children"] = [cr(i).Name for i in range(1, cr.Count + 1)]
        if t == 3:
            tr = sel.TextRange2 if _try(lambda: sel.TextRange2) is not None else sel.TextRange
            out["text_start"] = int(tr.Start)
            out["text_length"] = int(tr.Length)
        if t == 1:
            out["slides"] = [int(s.SlideIndex) for s in sel.SlideRange]
        out["slide"] = _try(lambda: int(win.View.Slide.SlideIndex))
        return out

    def slide_info(self, slide) -> dict:
        out = {
            "index": int(slide.SlideIndex),
            "id": int(slide.SlideID),
            "layout": PP_LAYOUT.get(int(slide.Layout), int(slide.Layout)),
            "custom_layout": _try(lambda: slide.CustomLayout.Name),
            "section": _try(lambda: int(slide.sectionIndex)),
            "shapes": [shape_info(slide.Shapes(i)) for i in range(1, slide.Shapes.Count + 1)],
        }
        notes = _try(lambda: notes_text(slide))
        if notes:
            out["notes"] = notes
        tr = slide.SlideShowTransition
        eff = _try(lambda: int(tr.EntryEffect), 0)
        if eff:
            out["transition"] = {"effect": eff, "duration": _try(lambda: float(tr.Duration))}
        seq = slide.TimeLine.MainSequence
        if seq.Count:
            out["animations"] = [{"shape": _try(lambda e=seq(i): e.Shape.Name), "effect": _try(lambda e=seq(i): int(e.EffectType)),
                                  "trigger": _try(lambda e=seq(i): int(e.Timing.TriggerType)),
                                  "duration": _try(lambda e=seq(i): float(e.Timing.Duration))} for i in range(1, seq.Count + 1)]
        return out

    def presentation_info(self, pres) -> dict:
        ps = pres.PageSetup
        master = pres.SlideMaster
        fonts = _try(lambda: master.Theme.ThemeFontScheme)
        out = {
            "slide_size": {"pt": [float(ps.SlideWidth), float(ps.SlideHeight)], "emu": [emu(ps.SlideWidth), emu(ps.SlideHeight)],
                           "size_type": int(ps.SlideSize)},
            "design": _try(lambda: master.Design.Name),
            "fonts": {"major": _try(lambda: fonts.MajorFont(1).Name), "minor": _try(lambda: fonts.MinorFont(1).Name)} if fonts else None,
            "layouts": [master.CustomLayouts(i).Name for i in range(1, master.CustomLayouts.Count + 1)],
            "slides": [self.slide_info(pres.Slides(i)) for i in range(1, pres.Slides.Count + 1)],
        }
        sp = pres.SectionProperties
        if sp.Count:
            out["sections"] = [{"name": sp.Name(i), "first": int(sp.FirstSlide(i)), "count": int(sp.SlidesCount(i))}
                               for i in range(1, sp.Count + 1)]
        return out

    # ----- the end ------------------------------------------------------------------------

    def close(self, kill_if_unreachable: bool = False) -> dict:
        """Closes our presentations, puts options back, quits only if nothing else is open.
        kill_if_unreachable: when COM does not answer (a modal dialog we must not click, such as a
        sign-in), end the process we started instead. Returns what was done (for the probe's log)."""
        done = {"closed": 0, "quit": False, "left_open": 0}
        if self.app is None:
            return done
        self.restore_ribbon()
        for obj, last, value in self._options.values():
            _try(lambda o=obj, a=last, v=value: setattr(o, a, v))
        for pres in self.own:
            try:
                pres.Saved = -1
                pres.Close()
                done["closed"] += 1
            except Exception:  # noqa: BLE001
                pass
        self.own = []
        left = _try(lambda: int(self.app.Presentations.Count), -1)
        done["left_open"] = left
        if left == 0:
            _try(self.app.Quit)
            done["quit"] = True
            self.app = None
            # PowerPoint lingers after Quit while COM references live; end only our own process.
            if self.pid:
                for _ in range(30):
                    time.sleep(0.2)
                    if self.pid not in powerpoint_pids():
                        break
                else:
                    subprocess.run(["taskkill", "/PID", str(self.pid), "/F"], capture_output=True)
                    for _ in range(25):
                        time.sleep(0.2)
                        if self.pid not in powerpoint_pids():
                            break
        elif left == -1 and kill_if_unreachable and self.pid:
            self.app = None
            subprocess.run(["taskkill", "/PID", str(self.pid), "/F"], capture_output=True)
            done["killed"] = True
        else:
            # Someone (the owner) opened a presentation in our process meanwhile: leave PowerPoint running.
            self.app = None
        import shutil
        shutil.rmtree(self.tmp, ignore_errors=True)
        return done


def notes_text(slide) -> str:
    page = slide.NotesPage
    for i in range(1, page.Shapes.Placeholders.Count + 1):
        ph = page.Shapes.Placeholders(i)
        if int(ph.PlaceholderFormat.Type) == 2:
            return str(ph.TextFrame.TextRange.Text)
    return ""


def text_info(sh) -> dict | None:
    if not _try(lambda: bool(sh.HasTextFrame), False):
        return None
    tf2 = sh.TextFrame2
    tf = sh.TextFrame
    out = {
        "autosize": AUTOSIZE.get(_try(lambda: int(tf2.AutoSize), -1), None),
        "wordwrap": _try(lambda: bool(tf2.WordWrap)),
        "anchor": _try(lambda: int(tf2.VerticalAnchor)),
        "margins_pt": _try(lambda: [float(tf2.MarginLeft), float(tf2.MarginTop), float(tf2.MarginRight), float(tf2.MarginBottom)]),
        "text": "",
    }
    if not _try(lambda: bool(tf.HasText), False):
        return out
    tr = tf.TextRange
    out["text"] = str(tr.Text)
    paras = []
    for i in range(1, tr.Paragraphs().Count + 1):
        p = tr.Paragraphs(i)
        f = p.Font
        paras.append({
            "text": str(p.Text).rstrip("\r"),
            "level": _try(lambda: int(p.IndentLevel)),
            "font": _try(lambda: str(f.Name)),
            "size": _try(lambda: float(f.Size)),
            "bold": _try(lambda: bool(f.Bold)),
            "bullet": _try(lambda: {"visible": bool(p.ParagraphFormat.Bullet.Visible), "type": int(p.ParagraphFormat.Bullet.Type),
                                    "char": int(p.ParagraphFormat.Bullet.Character)}),
            "align": _try(lambda: int(p.ParagraphFormat.Alignment)),
        })
    out["paragraphs"] = paras
    # Line positions as PowerPoint lays the text out (points on the slide).
    lines = []
    for i in range(1, 400):
        ln = _try(lambda n=i: tr.Lines(n, 1))
        if ln is None or _try(lambda: int(ln.Length), 0) == 0:
            break
        lines.append({"start": int(ln.Start), "length": int(ln.Length),
                      "left": _try(lambda: round(float(ln.BoundLeft), 2)), "top": _try(lambda: round(float(ln.BoundTop), 2)),
                      "width": _try(lambda: round(float(ln.BoundWidth), 2)), "height": _try(lambda: round(float(ln.BoundHeight), 2))})
    out["lines"] = lines
    return out


def shape_info(sh) -> dict:
    left, top, width, height = (float(sh.Left), float(sh.Top), float(sh.Width), float(sh.Height))
    t = int(sh.Type)
    out: dict = {
        "name": str(sh.Name),
        "id": _try(lambda: int(sh.Id)),
        "type": MSO_SHAPE_TYPE.get(t, t),
        "pt": [round(left, 3), round(top, 3), round(width, 3), round(height, 3)],
        "emu": [emu(left), emu(top), emu(width), emu(height)],
        "rotation": _try(lambda: float(sh.Rotation)),
        "hflip": _try(lambda: bool(sh.HorizontalFlip)),
        "vflip": _try(lambda: bool(sh.VerticalFlip)),
        "z": _try(lambda: int(sh.ZOrderPosition)),
    }
    if t == 14:
        out["placeholder"] = PP_PLACEHOLDER.get(_try(lambda: int(sh.PlaceholderFormat.Type), -1))
        out["contains"] = MSO_SHAPE_TYPE.get(_try(lambda: int(sh.PlaceholderFormat.ContainedType), -1))
    ast = _try(lambda: int(sh.AutoShapeType))
    if ast and ast > 0 and t in (1, 14, 17):
        out["autoshape"] = ast
        adj = _try(lambda: sh.Adjustments)
        if adj is not None and _try(lambda: adj.Count, 0):
            out["adjustments"] = [round(float(adj.Item(i)), 5) for i in range(1, adj.Count + 1)]
    if t in (1, 5, 14, 17):
        out["fill"] = _try(lambda: {"visible": bool(sh.Fill.Visible), "rgb": int(sh.Fill.ForeColor.RGB)})
        out["line"] = _try(lambda: {"visible": bool(sh.Line.Visible), "rgb": int(sh.Line.ForeColor.RGB), "weight": (float(sh.Line.Weight) if float(sh.Line.Weight) >= 0 else None)})
    txt = text_info(sh)
    if txt is not None and (txt["text"] or t in (14, 17)):
        out["text"] = txt
    if t == 6:
        gi = sh.GroupItems
        out["items"] = [shape_info(gi(i)) for i in range(1, gi.Count + 1)]
    if _try(lambda: bool(sh.HasTable), False):
        tb = sh.Table
        out["table"] = {"rows": int(tb.Rows.Count), "cols": int(tb.Columns.Count),
                        "style": _try(lambda: str(tb.Style.Name)),
                        "col_widths": [float(tb.Columns(c).Width) for c in range(1, tb.Columns.Count + 1)],
                        "row_heights": [float(tb.Rows(r).Height) for r in range(1, tb.Rows.Count + 1)],
                        "cells": [[str(tb.Cell(r, c).Shape.TextFrame.TextRange.Text) for c in range(1, tb.Columns.Count + 1)]
                                  for r in range(1, tb.Rows.Count + 1)]}
    if _try(lambda: bool(sh.HasChart), False):
        out["chart"] = {"type": _try(lambda: int(sh.Chart.ChartType)), "style": _try(lambda: int(sh.Chart.ChartStyle))}
    if t in (13, 11) or (t == 14 and _try(lambda: int(sh.PlaceholderFormat.ContainedType)) == 13):
        pf = sh.PictureFormat
        out["crop_pt"] = _try(lambda: [float(pf.CropLeft), float(pf.CropTop), float(pf.CropRight), float(pf.CropBottom)])
    return out
