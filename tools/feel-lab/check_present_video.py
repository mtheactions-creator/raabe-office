"""Video erstellen and Animiertes GIF erstellen compared with PowerPoint (Step 8b, D-074). Uses the deck and the
videos of probe_powerpoint_video.py (%TEMP%/raabe-video-probe, copied to captures/video, not committed): each
video is read frame by frame in a headless Edge (dev server) and every frame named by the colours at four points
(background; the shapes of slides 2-4), so the moments when slides change and animations play can be compared.
Our GIF of the same deck is checked with Pillow (frames, delays, size, loop).

    .venv/Scripts/python.exe check_present_video.py      (npm run dev:present must be running)

Result: results/<date>-present-video.json.
"""
from __future__ import annotations

import base64
import datetime
import io
import json
import shutil
import tempfile
from pathlib import Path

from PIL import Image

from feellab.present_page import PresentPage

HERE = Path(__file__).parent
SRC = Path(tempfile.gettempdir()) / "raabe-video-probe"
CAP = HERE / "captures" / "video"

ANALYSE = r"""
async (url) => {
  const v = document.createElement("video"); v.muted = true; v.src = url;
  await new Promise(r => { v.onloadeddata = r; v.onerror = r; setTimeout(r, 8000); });
  const W = v.videoWidth, H = v.videoHeight, c = document.createElement("canvas"); c.width = W; c.height = H;
  const g = c.getContext("2d", { willReadFrequently: true });
  const px = (x, y) => { const d = g.getImageData(Math.round(x * W), Math.round(y * H), 1, 1).data; return [d[0], d[1], d[2]]; };
  const name = ([r, gg, b]) => r > 200 && gg < 80 && b < 80 ? "R" : r < 80 && gg > 130 && b < 100 ? "G" : r < 90 && b > 180 ? "B" : r > 200 && gg > 180 && b < 90 ? "Y" : r > 190 && gg < 90 && b > 160 ? "M" : r < 40 && gg < 40 && b < 40 ? "k" : r > 230 && gg > 230 && b > 230 ? "w" : "?";
  const rows = []; let last = "";
  for (let f = 0; f < v.duration * 30; f++) {
    await new Promise(r => { v.onseeked = r; v.currentTime = (f + 0.5) / 30; setTimeout(r, 2000); });
    g.drawImage(v, 0, 0);
    const s = name(px(0.02, 0.03)) + name(px(225 / 960, 250 / 540)) + name(px(625 / 960, 250 / 540)) + name(px(450 / 960, 250 / 540));
    if (s !== last) { rows.push([+(f / 30).toFixed(3), s]); last = s; }
  }
  return { width: W, height: H, duration: v.duration, rows };
}"""


def compare(ppt: list, ours: list, tol: float = 0.2) -> dict:
    """For each state PowerPoint shows with clear colours (no '?'), when ours first shows the same state."""
    clear = [(t, s) for t, s in ppt if "?" not in s]
    found = []
    for t, s in clear:
        hits = [u for u, x in ours if x == s]
        near = min(hits, key=lambda u: abs(u - t)) if hits else None
        found.append({"state": s, "powerpoint": t, "ours": near, "diff": None if near is None else round(near - t, 3)})
    ok = sum(1 for f in found if f["diff"] is not None and abs(f["diff"]) <= tol)
    return {"matched": f"{ok}/{len(found)}", "states": found}


def main() -> None:
    CAP.mkdir(parents=True, exist_ok=True)
    for f in SRC.glob("*"):
        shutil.copy(f, CAP / f.name)
    out: dict = {"date": datetime.date.today().isoformat()}
    page = PresentPage()
    try:
        page.open_file(CAP / "video-deck.pptx")
        base = "/@fs/" + CAP.resolve().as_posix() + "/"
        for name, timed in (("untimed", False), ("timed", True)):
            ppt = page.js(f"({ANALYSE})({json.dumps(base + f'powerpoint-{name}.mp4')})")
            ours = page.js(f"""(async () => {{ const v = await import('/src/video.ts');
                const b = await v.makeVideo(window.__present.app, {{ height: 480, fps: 30, secs: 2, useTimings: {str(timed).lower()} }});
                return await ({ANALYSE})(URL.createObjectURL(new Blob([b], {{ type: 'video/mp4' }}))); }})()""")
            out[name] = {"powerpoint": {k: ppt[k] for k in ("width", "height", "duration")}, "ours": {k: ours[k] for k in ("width", "height", "duration")},
                         "compare": compare(ppt["rows"], ours["rows"]), "powerpoint_rows": ppt["rows"], "our_rows": ours["rows"]}
        b64 = page.js("""(async () => { const v = await import('/src/video.ts');
            const b = await v.makeGif(window.__present.app, { height: 240, fps: 15, secs: 1, useTimings: false });
            let s = ''; for (let i = 0; i < b.length; i += 0x8000) s += String.fromCharCode(...b.subarray(i, i + 0x8000)); return btoa(s); })()""")
        gif = base64.b64decode(b64)
        (CAP / "ours.gif").write_bytes(gif)
        im = Image.open(io.BytesIO(gif))
        delays = []
        for i in range(im.n_frames):
            im.seek(i)
            delays.append(im.info.get("duration"))
        out["gif"] = {"size": list(im.size), "frames": im.n_frames, "loop": im.info.get("loop"), "total_ms": sum(d or 0 for d in delays), "bytes": len(gif)}
    finally:
        page.close()
    p = HERE / "results" / f"{out['date']}-present-video.json"
    p.write_text(json.dumps(out, ensure_ascii=False, indent=1), encoding="utf-8", newline="\n")
    print(json.dumps({k: (v["compare"]["matched"] if isinstance(v, dict) and "compare" in v else v) for k, v in out.items()}))


if __name__ == "__main__":
    main()
