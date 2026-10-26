"""The things we measure: Microsoft Word (driven through COM) and the prototypes
(Tauri/WebView2, read through the Chrome DevTools Protocol). Both expose the same
interface so every test runs unchanged against each of them.

Offsets always index the document text with paragraphs joined by "\\n" and no final
paragraph mark. Screen coordinates are physical pixels."""

from __future__ import annotations

import json
import os
import subprocess
import time
import urllib.request
from pathlib import Path

import websocket

from . import winapi

ROOT = Path(__file__).resolve().parents[3]
FIXTURES = json.loads((ROOT / "prototypes/src/shared/fixtures.json").read_text(encoding="utf-8"))
PROTOTYPE_EXE = ROOT / "prototypes/src-tauri/target/release/raabe-feel.exe"

# Same screen rectangle for every target so measurements are comparable.
WINDOW_RECT = (0, 0, 2000, 2050)


def fixture(name: str) -> list[dict]:
    base = FIXTURES["base"]
    if name != "long":
        return [dict(p) for p in base]
    out: list[dict] = []
    i = 1
    while len(out) < FIXTURES["longMinParagraphs"]:
        for p in base:
            out.append({"style": "h1", "text": f"{i}. {p['text']}"} if p["style"] == "h1" else dict(p))
        i += 1
    return out


class Target:
    name: str
    hwnd: int

    def activate(self) -> None:
        winapi.activate(self.hwnd)

    def char_center(self, offset: int) -> tuple[float, float]:
        """Screen point in the middle of the character at `offset`."""
        x0, y0, h = self.point_at(offset)
        x1, y1, _ = self.point_at(offset + 1)
        if abs(y1 - y0) > 1:  # next character is on another line
            x1 = x0 + 6
        return (x0 + x1) / 2, y0 + h / 2

    def caret_point(self, offset: int) -> tuple[float, float]:
        x, y, h = self.point_at(offset)
        return x, y + h / 2

    def find(self, needle: str, start: int = 0) -> int:
        i = self.text().find(needle, start)
        if i < 0:
            raise ValueError(f"{needle!r} not in document")
        return i

    # Implemented by subclasses: text, selection, set_selection, point_at, lines, pages,
    # load, ensure_visible, close.


# ----- Microsoft Word -------------------------------------------------------------------

WD_STYLE_NORMAL, WD_STYLE_HEADING1 = -1, -2


class WordTarget(Target):
    name = "word"

    def __init__(self, fixture_name: str = "short", visible: bool = True):
        """`visible=False` for runs without real input (Word stays invisible)."""
        import win32com.client
        # A separate Word process, so the harness never touches the user's own documents.
        self.app = win32com.client.gencache.EnsureDispatch(win32com.client.DispatchEx("Word.Application"))
        try:
            self.app.Visible = visible
            self.doc = self.app.Documents.Add()
            self.doc.ShowSpellingErrors = False
            self.doc.ShowGrammaticalErrors = False
            self._setup_styles()
            win = self.doc.Windows(1)
            win.View.Type = 3  # print layout
            win.View.Zoom.Percentage = 100
            win.View.ShowAll = False
            self.hwnd = win.Hwnd
            if visible:
                winapi.place(self.hwnd, *WINDOW_RECT)
            self.load(fixture_name)
        except Exception:
            self.app.Quit(0)
            raise

    def _setup_styles(self) -> None:
        d = self.doc
        ps = d.PageSetup
        ps.PageWidth, ps.PageHeight = 595.3, 841.9
        ps.TopMargin, ps.BottomMargin, ps.LeftMargin, ps.RightMargin = 70.85, 56.7, 70.85, 70.85
        multiple = 12 * 1.08  # "multiple 1.08" is stored as points of a 12 pt line
        for sid, font, size, color, before, after, keep in (
            (WD_STYLE_NORMAL, "Calibri", 11, -16777216, 0, 8, False),  # automatic colour
            (WD_STYLE_HEADING1, "Calibri Light", 16, 0x96542F, 12, 0, True),
        ):
            s = d.Styles(sid)
            s.Font.Name, s.Font.Size, s.Font.Color, s.Font.Bold = font, size, color, False
            s.Font.Kerning, s.Font.Ligatures = 0, 0
            pf = s.ParagraphFormat
            pf.SpaceBefore, pf.SpaceAfter = before, after
            pf.LineSpacingRule, pf.LineSpacing = 5, multiple
            pf.WidowControl, pf.KeepWithNext = True, keep

    def load(self, name: str) -> None:
        paras = fixture(name)
        self.doc.Content.Text = "\r".join(p["text"] for p in paras)
        # Formatting from a previous case must not survive the text replacement.
        self.doc.Content.Font.Reset()
        self.doc.Content.ParagraphFormat.Reset()
        self.doc.Content.Style = WD_STYLE_NORMAL
        for i, p in enumerate(paras, start=1):
            if p["style"] == "h1":
                self.doc.Paragraphs(i).Style = WD_STYLE_HEADING1
        self.doc.UndoClear()
        self.set_selection(0, 0)
        self.doc.Windows(1).ScrollIntoView(self.doc.Range(0, 0), True)
        self.doc.Windows(1).VerticalPercentScrolled = 0

    def text(self) -> str:
        r = self.doc.Content
        r.TextRetrievalMode.IncludeHiddenText = True  # hidden text is part of the document
        t = r.Text
        if t.endswith("\r"):
            t = t[:-1]
        return t.replace("\r", "\n")

    def wb(self, command: str) -> None:
        """Runs a built-in Word command by name (the one a shortcut is bound to)."""
        import pythoncom
        o = self.app.WordBasic._oleobj_
        o.Invoke(o.GetIDsOfNames(command), 0, pythoncom.DISPATCH_METHOD, False)

    def selection(self) -> tuple[int, int]:
        s = self.app.Selection
        n = len(self.text())
        start, end = min(s.Start, n), min(s.End, n)
        return (end, start) if s.StartIsActive else (start, end)

    def set_selection(self, anchor: int, head: int) -> None:
        s = self.app.Selection
        s.SetRange(min(anchor, head), max(anchor, head))
        s.StartIsActive = head < anchor

    def point_at(self, offset: int) -> tuple[float, float, float]:
        rng = self.doc.Range(offset, offset)
        left, top, width, height = self.doc.Windows(1).GetPoint(0, 0, 0, 0, rng)
        return float(left), float(top), float(height)

    def lines(self) -> list[int]:
        s = self.app.Selection
        saved = (s.Start, s.End)
        starts: list[int] = []
        s.HomeKey(6)  # wdStory
        while True:
            s.HomeKey(5)  # wdLine
            starts.append(s.Start)
            if s.MoveDown(5, 1) == 0:
                break
        s.SetRange(*saved)
        return starts

    def pages(self) -> list[int]:
        out: list[int] = []
        last = 0
        for start in self.lines():
            page = self.doc.Range(start, start).Information(3)  # wdActiveEndPageNumber
            if page != last:
                out.append(start)
                last = page
        return out

    def ensure_visible(self, offset: int) -> None:
        self.doc.Windows(1).ScrollIntoView(self.doc.Range(offset, offset), True)

    def close(self) -> None:
        self.doc.Close(0)
        self.app.Quit(0)


# ----- prototypes and apps -----------------------------------------------------------------

# RAABE_WRITE_EXE overrides it: a build in another target folder, while the owner has the app
# open and the release exe cannot be replaced.
RAABE_WRITE_EXE = Path(os.environ.get("RAABE_WRITE_EXE") or ROOT / "apps/raabe-write/src-tauri/target/release/raabe-write.exe")


def _pid(hwnd: int) -> int:
    import ctypes
    pid = ctypes.c_ulong()
    winapi.user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
    return pid.value


def clear_feel_recovery() -> None:
    """Test runs end the app abruptly; their AutoRecover backups must not reach the next run."""
    import tempfile
    d = Path(tempfile.gettempdir()) / "raabe-write-feel-recovery"
    for f in d.glob("*") if d.exists() else []:
        f.unlink(missing_ok=True)


def raabe_write(fixture_name: str = "short", port: int = 9370, extra: list[str] | None = None,
                keep_recovery: bool = False, background: bool = False, browser_args: str = "") -> "WebViewTarget":
    """`background`: off screen and never focused, for runs that need no real input."""
    if not keep_recovery:
        clear_feel_recovery()
    return WebViewTarget("raabe-write", fixture_name, port, exe=RAABE_WRITE_EXE,
                         args=["--feel", fixture_name, *(extra or [])], title_suffix=" - RaabeWrite", background=background,
                         browser_args=browser_args)

class WebViewTarget(Target):
    def __init__(self, engine: str, fixture_name: str = "short", port: int = 9333,
                 exe: Path | None = None, args: list[str] | None = None, title_suffix: str | None = None,
                 background: bool = False, browser_args: str = ""):
        self.name = f"proto-{engine}" if exe is None else engine
        self.engine = engine
        env = dict(os.environ, RAABE_FEEL_DEBUG_PORT=str(port))
        if background:
            env["RAABE_FEEL_BACKGROUND"] = "1"
        if browser_args:
            env["RAABE_FEEL_BROWSER_ARGS"] = browser_args
        self.started = time.perf_counter()
        self.proc = subprocess.Popen([str(exe or PROTOTYPE_EXE), *(args if args is not None else [engine, fixture_name])], env=env)
        if title_suffix is None:
            title = f"Feel-Prototyp {engine.upper()}"
            match = lambda h: winapi.window_title(h) == title
        else:
            match = lambda h: winapi.window_title(h).endswith(title_suffix) and _pid(h) == self.proc.pid
        self.hwnd = winapi.wait_for_window(match)
        self.window_shown = time.perf_counter()
        if not background:
            winapi.place(self.hwnd, *WINDOW_RECT)
        self.ws = self._connect(port)
        self._id = 0
        while not self.js("typeof window.__probe === 'object'"):
            time.sleep(0.01)
        self.ready = time.perf_counter()

    def _connect(self, port: int):
        deadline = time.perf_counter() + 20
        while True:
            try:
                with urllib.request.urlopen(f"http://127.0.0.1:{port}/json/list", timeout=1) as r:
                    pages = [p for p in json.load(r) if p.get("type") == "page"]
                if pages:
                    return websocket.create_connection(pages[0]["webSocketDebuggerUrl"], suppress_origin=True)
            except OSError:
                pass
            if time.perf_counter() > deadline:
                raise TimeoutError("DevTools endpoint not reachable")
            time.sleep(0.05)

    def cdp(self, method: str, params: dict | None = None):
        self._id += 1
        self.ws.send(json.dumps({"id": self._id, "method": method, "params": params or {}}))
        while True:
            msg = json.loads(self.ws.recv())
            if msg.get("id") == self._id:
                if "error" in msg:
                    raise RuntimeError(msg["error"])
                return msg.get("result")

    # Input through the DevTools protocol goes straight into the page: it never touches the
    # real keyboard or mouse, so these helpers can run while someone uses the PC.
    KEYS = {"Enter": 13, "Backspace": 8, "Tab": 9, "Escape": 27, "Delete": 46, "Home": 36, "End": 35,
            "ArrowLeft": 37, "ArrowUp": 38, "ArrowRight": 39, "ArrowDown": 40, "F1": 112, "F10": 121, "F12": 123, "Alt": 18}

    def send_text(self, text: str) -> None:
        for ch in text:
            self.cdp("Input.dispatchKeyEvent", {"type": "keyDown", "text": ch, "key": ch, "unmodifiedText": ch})
            self.cdp("Input.dispatchKeyEvent", {"type": "keyUp", "key": ch})

    def send_key(self, key: str, ctrl: bool = False, shift: bool = False, alt: bool = False, vk: int | None = None) -> None:
        """`vk` overrides the virtual-key code (keys such as < or # on a German keyboard)."""
        mods = (1 if alt else 0) | (2 if ctrl else 0) | (8 if shift else 0)
        vk = vk if vk is not None else self.KEYS.get(key, ord(key.upper()) if len(key) == 1 else 0)
        code = key if len(key) > 1 else f"Key{key.upper()}" if key.isalpha() else f"Digit{key}"
        base = {"key": key, "code": code, "windowsVirtualKeyCode": vk, "nativeVirtualKeyCode": vk, "modifiers": mods}
        self.cdp("Input.dispatchKeyEvent", {"type": "rawKeyDown", **base})
        self.cdp("Input.dispatchKeyEvent", {"type": "keyUp", **base})

    def js(self, expression: str):
        self._id += 1
        self.ws.send(json.dumps({"id": self._id, "method": "Runtime.evaluate",
                                 "params": {"expression": expression, "returnByValue": True, "awaitPromise": True}}))
        while True:
            msg = json.loads(self.ws.recv())
            if msg.get("id") == self._id:
                result = msg["result"]
                if "exceptionDetails" in result:
                    raise RuntimeError(result["exceptionDetails"])
                return result["result"].get("value")

    def _viewport_origin(self) -> tuple[int, int]:
        for child in winapi.child_windows(self.hwnd):
            if winapi.window_class(child) == "Chrome_RenderWidgetHostHWND":
                return winapi.client_origin(child)
        return winapi.client_origin(self.hwnd)

    def text(self) -> str:
        return self.js("__probe.text()")

    def selection(self) -> tuple[int, int]:
        s = self.js("__probe.selection()")
        return s["anchor"], s["head"]

    def set_selection(self, anchor: int, head: int) -> None:
        self.js(f"__probe.setSelection({anchor}, {head})")

    def point_at(self, offset: int) -> tuple[float, float, float]:
        p = self.js(f"__probe.pointAt({offset})")
        dpr = self.js("devicePixelRatio")
        ox, oy = self._viewport_origin()
        return ox + p["x"] * dpr, oy + p["y"] * dpr, p["h"] * dpr

    def lines(self) -> list[int]:
        return self.js("__probe.lines()")

    def pages(self) -> list[int]:
        return self.js("__probe.pages()")

    def load(self, name: str) -> None:
        self.js(f"__probe.load({json.dumps(name)})")

    def ensure_visible(self, offset: int) -> None:
        self.js(f"__probe.reveal({offset})")

    def close(self) -> None:
        try:
            self.ws.close()
        finally:
            self.proc.terminate()
            self.proc.wait(5)
