"""Video erstellen and Animiertes GIF erstellen as PowerPoint makes them (Step 8b, D-074). An invisible
PowerPoint of our own (only when no POWERPNT.EXE runs) builds a test presentation in which every slide and
shape has a colour of its own, so a frame tells which slide and which animation state it shows:

  1 red background, nothing animated
  2 green background, two shapes with a click entrance (Erscheinen)
  3 blue background, transition Verblassen 1 s, one shape with a click entrance (Verblassen 0.5 s)
  4 yellow background, one shape entering after the previous (Erscheinen)
  5 magenta background, Nach 3 s (advance after time)

It saves the deck and asks PowerPoint for videos (CreateVideo: without and with recorded timings, 2 s per slide,
480p, 30 fps) and an animated GIF. Files go to %TEMP%/raabe-video-probe (not committed); the result
(durations, sizes) to results/<date>-powerpoint-video.json. check_present_video.py compares their frames
with ours.

    .venv/Scripts/python.exe probe_powerpoint_video.py
"""
from __future__ import annotations

import datetime
import json
import shutil
import struct
import tempfile
import time
from pathlib import Path

from feellab.powerpoint import PowerPointTarget

HERE = Path(__file__).parent
OUT = Path(tempfile.gettempdir()) / "raabe-video-probe"

COLORS = [(230, 30, 30), (30, 170, 60), (40, 70, 220), (240, 210, 30), (220, 40, 200)]
SHAPE = [(0, 0, 0), (255, 255, 255)]


def rgb(c: tuple[int, int, int]) -> int:
    return c[0] | (c[1] << 8) | (c[2] << 16)


def mp4_duration(path: Path) -> float | None:
    """The movie's duration from mvhd (seconds)."""
    b = path.read_bytes()
    i = b.find(b"mvhd")
    if i < 0:
        return None
    ver = b[i + 4]
    if ver == 1:
        scale, dur = struct.unpack(">IQ", b[i + 24:i + 36])
    else:
        scale, dur = struct.unpack(">II", b[i + 16:i + 24])
    return dur / scale if scale else None


def wait_video(pres, timeout: float = 300) -> int:
    """PowerPoint encodes in the background: CreateVideoStatus 1 in progress, 2 queued, 3 done, 4 failed."""
    end = time.time() + timeout
    st = 0
    while time.time() < end:
        st = int(pres.CreateVideoStatus)
        if st in (3, 4):
            return st
        time.sleep(0.5)
    return st


def main() -> None:
    OUT.mkdir(exist_ok=True)
    out: dict = {"date": datetime.date.today().isoformat(), "slides": [], "videos": {}}
    t = PowerPointTarget(visible=False)
    try:
        pres = t.new_presentation(with_window=False, name="video")
        blank = None
        for i in range(1, pres.SlideMaster.CustomLayouts.Count + 1):
            if pres.SlideMaster.CustomLayouts(i).Name in ("Leer", "Blank"):
                blank = pres.SlideMaster.CustomLayouts(i)
        for n, col in enumerate(COLORS, start=1):
            sl = pres.Slides.AddSlide(n, blank or pres.SlideMaster.CustomLayouts(7))
            sl.FollowMasterBackground = 0
            sl.Background.Fill.Solid()
            sl.Background.Fill.ForeColor.RGB = rgb(col)
            info = {"slide": n, "color": col}
            if n == 2:
                for k in range(2):
                    s = sl.Shapes.AddShape(1, 100 + k * 400, 150, 250, 200)
                    s.Fill.ForeColor.RGB = rgb(SHAPE[k]); s.Line.Visible = 0
                    sl.TimeLine.MainSequence.AddEffect(s, 1, 0, 1)  # msoAnimEffectAppear, on click
                info["clicks"] = 2
            if n == 3:
                sl.SlideShowTransition.EntryEffect = 1793  # ppEffectFade
                sl.SlideShowTransition.Duration = 1.0
                s = sl.Shapes.AddShape(1, 300, 150, 300, 200)
                s.Fill.ForeColor.RGB = rgb(SHAPE[0]); s.Line.Visible = 0
                e = sl.TimeLine.MainSequence.AddEffect(s, 10, 0, 1)  # msoAnimEffectFade, on click
                e.Timing.Duration = 0.5
                info["transition"] = "fade 1 s"; info["clicks"] = 1
            if n == 4:
                s = sl.Shapes.AddShape(1, 300, 150, 300, 200)
                s.Fill.ForeColor.RGB = rgb(SHAPE[1]); s.Line.Visible = 0
                sl.TimeLine.MainSequence.AddEffect(s, 1, 0, 3)  # after previous
                info["afterPrevious"] = 1
            if n == 5:
                sl.SlideShowTransition.AdvanceOnTime = -1
                sl.SlideShowTransition.AdvanceTime = 3
                info["advanceAfter"] = 3
            out["slides"].append(info)
        deck = OUT / "video-deck.pptx"
        t.save_copy(pres, deck)
        for name, timed in (("untimed", False), ("timed", True)):
            path = OUT / f"powerpoint-{name}.mp4"
            path.unlink(missing_ok=True)
            started = time.time()
            pres.CreateVideo(str(path), timed, 2, 480, 30, 85)
            st = wait_video(pres)
            time.sleep(1)
            out["videos"][name] = {"status": st, "seconds_to_make": round(time.time() - started, 1),
                                  "bytes": path.stat().st_size if path.exists() else 0,
                                  "duration": mp4_duration(path) if path.exists() else None}
        # Animiertes GIF: PowerPoint's CreateVideo takes a .gif name too.
        gif = OUT / "powerpoint.gif"
        gif.unlink(missing_ok=True)
        try:
            pres.CreateVideo(str(gif), False, 1, 480, 30, 85)
            st = wait_video(pres)
            time.sleep(1)
            head = gif.read_bytes()[:12] if gif.exists() else b""
            # Measured: CreateVideo writes an MP4 whatever the name; PowerPoint's GIF needs its UI.
            out["videos"]["gif"] = {"status": st, "bytes": gif.stat().st_size if gif.exists() else 0, "is_gif": head.startswith(b"GIF8"), "is_mp4": head[4:8] == b"ftyp"}
        except Exception as e:  # noqa: BLE001
            out["videos"]["gif"] = {"error": str(e)[:200]}
    finally:
        out["close"] = t.close(kill_if_unreachable=True)
    p = HERE / "results" / f"{out['date']}-powerpoint-video.json"
    p.write_text(json.dumps(out, ensure_ascii=False, indent=1), encoding="utf-8", newline="\n")
    print(json.dumps(out))


if __name__ == "__main__":
    main()
