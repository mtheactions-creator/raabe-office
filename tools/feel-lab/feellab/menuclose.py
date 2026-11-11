"""Step 11 (D-156): what the menu-closing checks share. Trusted mouse and key input through the DevTools protocol
(Input.dispatchMouseEvent: the page sees real, trusted events, but the OS mouse and keyboard stay untouched), whether a
popup is open, where a control is. `side` is any of our app sides with a DevTools websocket (`ws`, `_id`, `js`):
CalcSide, PresentSide, WebViewTarget."""
from __future__ import annotations

import json
import time

# Is a popup open: the shell marks it on <html> (showPopup), the apps' own popups are found by their classes, and
# anything else floating above the app counts too.
JS_POPUPS = """(() => {
  const vis = (e) => e && e.isConnected && e.offsetWidth > 40 && e.offsetHeight > 30 && getComputedStyle(e).display !== 'none' && getComputedStyle(e).visibility !== 'hidden';
  const fixed = [...document.body.children].filter((e) => !e.matches('.tooltip, script, style') && getComputedStyle(e).position === 'fixed' && vis(e));
  const own = [...document.querySelectorAll('.quick-panel, .filter-panel, .pick-list, .dv-list, .chart-flyout, .cb-flyout')].filter(vis);
  const shell = document.documentElement.hasAttribute('data-popup-open') ? [...document.body.children].filter((e) => e.matches('.menu-root, .tile-panel, .menu, .palette, .rpanel, .group-popup, .popup-card')) : [];
  return [...new Set([...fixed, ...own, ...shell])].filter(vis).map((e) => { const r = e.getBoundingClientRect(); return [r.left, r.top, r.right, r.bottom]; });
})()"""

JS_RECT = """((sel, text) => {
  const els = [...document.querySelectorAll(sel)].filter((e) => !text || e.textContent.trim() === text || e.getAttribute('aria-label') === text);
  const e = els.find((x) => x.offsetWidth > 0);
  if (!e) return null;
  const r = e.getBoundingClientRect();
  return { x: r.left + r.width / 2, y: r.top + r.height / 2 };
})"""

JS_RIBBON_TAB = "document.querySelector('.tabrow [role=tab][aria-selected=true]')?.textContent?.trim()"


def _send(side, method: str, params: dict) -> None:
    side._id += 1
    side.ws.send(json.dumps({"id": side._id, "method": method, "params": params}))
    while json.loads(side.ws.recv()).get("id") != side._id:
        pass


def mouse(side, x: float, y: float, button: str = "left") -> None:
    """A whole press at (x, y) in CSS pixels of the page: move, press, release."""
    for t in ("mouseMoved", "mousePressed", "mouseReleased"):
        _send(side, "Input.dispatchMouseEvent", {
            "type": t, "x": x, "y": y, "button": "none" if t == "mouseMoved" else button,
            "buttons": (1 if button == "left" else 2) if t == "mousePressed" else 0,
            "clickCount": 0 if t == "mouseMoved" else 1})
        time.sleep(0.03)


def key(side, name: str = "Escape", code: int = 27) -> None:
    for t in ("rawKeyDown", "keyUp"):
        _send(side, "Input.dispatchKeyEvent", {"type": t, "key": name, "code": name, "windowsVirtualKeyCode": code})


def rect(side, sel: str, text: str = "") -> tuple[float, float]:
    p = side.js(f"{JS_RECT}({json.dumps(sel)}, {json.dumps(text)})")
    if not p:
        raise RuntimeError(f"not found: {sel} {text}")
    return p["x"], p["y"]


def color_arrow(side, label: str) -> tuple[float, float]:
    """The arrow of a ribbon colour split button (Füllfarbe, Schriftfarbe …)."""
    p = side.js(f"""(() => {{ const s = [...document.querySelectorAll('.rsplit')].find((x) => x.querySelector('.color-main')?.getAttribute('aria-label') === {json.dumps(label)} && x.offsetWidth > 0);
      const r = s.querySelector('.arrow').getBoundingClientRect(); return {{ x: r.left + r.width / 2, y: r.top + r.height / 2 }}; }})()""")
    return p["x"], p["y"]


def popups(side) -> list:
    return side.js(JS_POPUPS)


def close_popups(side) -> None:
    for _ in range(3):
        if popups(side):
            key(side)
            time.sleep(0.15)


def wait_popup(side, tries: int = 10) -> bool:
    for _ in range(tries):
        time.sleep(0.1)
        if popups(side):
            return True
    return False


def pick_outside(targets: dict, pops: list, margin: float = 8):
    """The first target outside every open popup: a press meant for "outside" must never land on a menu entry."""
    for name, (x, y) in targets.items():
        if all(not (r[0] - margin <= x <= r[2] + margin and r[1] - margin <= y <= r[3] + margin) for r in pops):
            return name, (x, y)
    return None, None
