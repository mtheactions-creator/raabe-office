"""By-hand sweep of RaabeCalc's options (D-141): starts the release exe off screen (an instance of our
own, RAABE_FEEL_BACKGROUND), and for every control a user can reach (ribbon tabs, every button, split
arrow, dropdown, menu entry, submenu entry, gallery entry, colour, combo list entry, launcher, the cell,
row, column and sheet-tab context menus, the status bar, the Datei backstage) it sets up the same small
workbook, triggers the option with trusted DevTools mouse clicks at the control's position (as a user
does; the real mouse is not touched) and records:

- whether the control was there, enabled and not covered;
- what it opened (a menu with its entries, a dialog with its controls, a pane);
- exceptions and console errors;
- the effect on the workbook, the view and the drawn grid;
- for a change of the workbook: whether Strg+Z, pressed as a key right after the click (where the focus is
  then), gives back the workbook as it was, and whether Strg+Y redoes it.
Dialogs are run twice: cancelled with Esc and confirmed with their default button as they open.

    .venv/Scripts/python.exe sweep_calc.py [--tabs Start,Daten] [--only <id substring>] [--quick]

Results: results/2026-11-09-calc-sweep.json (every option with its outcome) and
results/2026-11-09-calc-options.json (the inventory: every option with its kind and where it is).
"""
from __future__ import annotations

import argparse
import ctypes
import datetime
import json
import os
import sys
import tempfile
import time
from pathlib import Path

from feellab import winapi
from feellab.cdp_app import AppExe

HERE = Path(__file__).parent
JS = (HERE / "feellab" / "sweep_calc.js").read_text(encoding="utf-8")
TMP = Path(tempfile.gettempdir()) / "raabe-calc-sweep"

#: A small workbook with what most options need: a header row, text, numbers, dates, a formula, a
#: bold header and a number format (a mixed selection shows options that wrongly copy the active cell's format).
SETUP = {
    "A1": "Name", "B1": "Punkte", "C1": "Datum", "D1": "Note",
    "A2": "Anna", "B2": "12", "C2": "01.03.2026", "D2": "2",
    "A3": "Ben", "B3": "7,5", "C3": "15.03.2026", "D3": "4",
    "A4": "Cem", "B4": "15", "C4": "02.04.2026", "D4": "1",
    "A5": "Dora", "B5": "9", "C5": "20.04.2026", "D5": "3",
    "A6": "Summe", "B6": "=SUMME(B2:B5)",
}
#: The selection most options act on (mixed: A1 is bold after the setup, B2:B5 have a number format).
SELECTION = "A1:B4"

#: Options that open a native window of the system (file dialogs, printing) or end the app: not clicked.
SKIP = {
    "Dieses Gerät...", "Aus Text/CSV", "Öffnen", "Speichern unter", "Durchsuchen", "Schließen#backstage", "Drucken#print",
    "Aus Datei", "Aus Arbeitsmappe", "Aus Text/CSV-Datei", "Beenden",
}

WM_CLOSE = 0x0010


def native_windows(pid: int, main: int) -> list[int]:
    return [h for h in winapi.find_windows(lambda h: winapi.window_pid(h) == pid and h != main
                                           and winapi.window_class(h) in ("#32770",))]


class Sweep:
    def __init__(self, port: int = 9441, dev: str | None = None):
        """`dev` ("9447:http://localhost:5195/"): a dev page already open in a headless browser instead of the exe."""
        if dev:
            from feellab.cdp_app import DevPage
            dport, url = dev.split(":", 1)
            self.exe = DevPage(int(dport), url)
            self.exe.hwnd = 0
        else:
            self.exe = AppExe("calc", port=port)
        self.p = self.exe.page
        if dev:
            self.p.cdp("Page.reload")
            time.sleep(2)
        while not self.p.js("typeof window.__calc === 'object'"):
            time.sleep(0.1)
        self.p.js(JS)
        self.results: list[dict] = []
        self.inventory: dict = {}

    # ----- setup ---------------------------------------------------------------------------------
    def reset(self, selection: str = SELECTION) -> None:
        p = self.p
        for _ in range(4):
            # Popups, dialogs, the backstage: Esc as a user would; what stays is closed by its button.
            st = p.js("({dlg: !!document.querySelector('.dialog-layer'), pop: [...document.body.children].some(e => e.matches('.menu-root, .menu, .rpanel, .popup, .combo-list') ), bs: !document.querySelector('.backstage')?.hidden})")
            if not (st["dlg"] or st["pop"] or st["bs"]):
                break
            p.key("Escape")
            p.pump(0.1)
        p.js("document.querySelectorAll('.dialog-layer').forEach(e => e.remove()); document.querySelectorAll('.menu-root, .combo-list, .fc-pop').forEach(e => e.remove())")
        for h in (native_windows(self.exe.proc.pid, self.exe.hwnd) if self.exe.proc else []):
            ctypes.windll.user32.PostMessageW(h, WM_CLOSE, 0, 0)
        p.js(f"""(() => {{
          const app = __calc.app;
          document.querySelectorAll('.taskpane:not([hidden]) .tp-close').forEach(b => b.click());
          if (app.painter) app.painter = null;
          __calc.reset({json.dumps(SETUP)}, "A1", 2, []);
          // A bold header and a number format in B2:B5: a mixed selection.
          const P = __calc.parse;
          app.change("Test", () => {{}});
          return true;
        }})()""")
        p.js("""(() => { const app = __calc.app, A = __calc.parse; const g = app.grid;
          g.select(A("A1:D1")); app.commands.run("font.bold");
          g.select(A("B2:B5")); app.commands.run("number.thousands");
          app.history.clear ? app.history.clear() : 0; return true; })()""")
        p.js(f"""(() => {{ const app = __calc.app, r = __calc.parse({json.dumps(selection)}); app.grid.select(r, {{ row: r.r0, col: r.c0 }}); app.grid.focus?.(); __sweep.mark(); return true; }})()""")
        p.pump(0.05)
        p.errors.clear()

    # ----- path replay ---------------------------------------------------------------------------
    def step(self, kind: str, ident: str, button: str = "left") -> dict:
        el = self.p.js(f"__sweep.find({json.dumps(kind)}, {json.dumps(ident)})")
        if el is None:
            return {"ok": False, "why": "not found"}
        if el.get("disabled"):
            return {"ok": False, "why": "disabled"}
        if not el.get("reachable", True):
            return {"ok": False, "why": "covered"}
        self.p.click(el["x"], el["y"], button=button)
        self.p.pump(0.25)
        return {"ok": True}

    def replay(self, path: list[tuple[str, str]]) -> dict:
        for kind, ident in path:
            if kind == "ctx":
                r = self.context_menu(ident)
            else:
                r = self.step(kind, ident)
            if not r["ok"]:
                return {**r, "at": ident}
        return {"ok": True}

    def context_menu(self, where: str) -> dict:
        """A right click on a cell, a row or column header, a sheet tab, the status bar."""
        pt = self.p.js(f"""(() => {{
          const w = {json.dumps(where)};
          if (w === "cell") return __calc.point("B3", "center");
          if (w === "row") return __calc.header("row", "3");
          if (w === "col") return __calc.header("col", "B");
          if (w === "sheettab") {{ const t = [...document.querySelectorAll('.sheettab')][0]; const r = t.getBoundingClientRect(); return {{x: r.left + r.width / 2, y: r.top + r.height / 2}}; }}
          if (w === "statusbar") {{ const t = document.querySelector('.statusbar .stats') || document.querySelector('.statusbar'); const r = t.getBoundingClientRect(); return {{x: r.left + 20, y: r.top + r.height / 2}}; }}
          return null; }})()""")
        if not pt:
            return {"ok": False, "why": "not found"}
        self.p.click(pt["x"], pt["y"], button="right")
        self.p.pump(0.3)
        return {"ok": True}

    # ----- observing -----------------------------------------------------------------------------
    def snap(self) -> dict:
        return self.p.js("__sweep.snap()")

    def dialog_close(self, how: str) -> list[str]:
        """Closes the open dialog: "esc" or "ok" (a click on its default button). Returns the titles of the
        dialogs met on the way (a message box after OK)."""
        seen = []
        for _ in range(4):
            t = self.p.js("__sweep.dialogTitle()")
            if not t:
                break
            seen.append(t)
            if how == "ok" and len(seen) == 1:
                b = self.p.js("""(() => { const d = [...document.querySelectorAll('.dialog-layer .dialog')].pop();
                  const b = d && (d.querySelector('.dlg-buttons button.primary') || d.querySelector('.dlg-buttons button') || d.querySelector('button.primary'));
                  if (!b) return null; const r = b.getBoundingClientRect(); return {x: r.left + r.width / 2, y: r.top + r.height / 2, label: b.textContent}; })()""")
                if b:
                    self.p.click(b["x"], b["y"])
                else:
                    self.p.key("Escape")
            else:
                self.p.key("Escape")
            self.p.pump(0.3)
        return seen

    def leaf(self, path: list[tuple[str, str]], opened_by: dict | None = None, variant: str = "") -> dict:
        """Runs one option: replays the path, looks at what happened, undoes and redoes."""
        self.reset()
        before = self.snap()
        res: dict = {"path": [i for _, i in path], "variant": variant}
        r = self.replay(path)
        if not r["ok"]:
            res.update(outcome=r["why"], at=r.get("at"))
            res["errors"] = list(self.p.errors)
            return res
        opened = self.p.js("__sweep.opened()")
        natives = (native_windows(self.exe.proc.pid, self.exe.hwnd) if self.exe.proc else [])
        if natives:
            res["native"] = [winapi.window_title(h) for h in natives]
            for h in natives:
                ctypes.windll.user32.PostMessageW(h, WM_CLOSE, 0, 0)
        if opened and opened["kind"] == "dialog":
            res["dialog"] = opened["title"]
            res["dialogControls"] = opened["controls"]
            res["dialogs"] = self.dialog_close(variant or "esc")
        elif opened and opened["kind"] == "popup" and not variant:
            res["popup"] = [{"id": i["id"], "disabled": i["disabled"], "sub": i.get("sub"), "tag": i["tag"], "level": i.get("level", 0)} for i in opened["items"]]
        self.p.pump(0.15)
        # A popup left open would swallow the keys: Esc closes it first (as a user would) unless it is the result.
        after = self.snap()
        res["model"] = after["model"] != before["model"]
        res["ui"] = after["ui"] != before["ui"]
        res["canvas"] = after["canvas"] != before["canvas"]
        if res["ui"]:
            a, b = json.loads(before["ui"]), json.loads(after["ui"])
            res["uiDiff"] = {k: [a.get(k), b.get(k)] for k in set(a) | set(b) if a.get(k) != b.get(k) and k not in ("status",)}
        hist = self.p.js("__sweep.history()")
        res["history"] = hist
        if res["model"]:
            focus = self.p.js("(() => { const a = document.activeElement; return a ? (a.className || a.tagName) + '|' + (a.getAttribute('aria-label') || '') : null; })()")
            res["focus"] = focus
            self.p.key("z", "ctrl")
            self.p.pump(0.25)
            undone = self.snap()
            res["undo"] = undone["model"] == before["model"]
            if not res["undo"]:
                # Was it the key that went nowhere, or the undo itself? Run Rückgängig directly to tell.
                res["undoKeyLost"] = undone["model"] == after["model"]
                if res["undoKeyLost"]:
                    self.p.js("__calc.app.commands.run('edit.undo')")
                    self.p.pump(0.2)
                    res["undoCmd"] = self.snap()["model"] == before["model"]
            else:
                self.p.key("y", "ctrl")
                self.p.pump(0.25)
                res["redo"] = self.snap()["model"] == after["model"]
        res["errors"] = list(self.p.errors)
        return res

    # ----- walking -------------------------------------------------------------------------------
    def walk(self, path: list[tuple[str, str]], depth: int, label: str, kind: str, limit: int = 60, sub: bool = False) -> None:
        """One option and, when it opens a menu, every entry in it (depth-first)."""
        t0 = time.perf_counter()
        try:
            res = self.leaf(path)
        except Exception as e:  # noqa: BLE001 - one option's trouble must not end the sweep
            res = {"path": [i for _, i in path], "outcome": "harness error", "why": str(e)[:300], "errors": list(self.p.errors)}
        res["kind"] = kind
        res["ms"] = round((time.perf_counter() - t0) * 1000)
        self.results.append(res)
        self.log(res)
        if res.get("dialog") and not res.get("errors"):
            try:
                ok = self.leaf(path, variant="ok")
            except Exception as e:  # noqa: BLE001
                ok = {"path": [i for _, i in path], "variant": "ok", "outcome": "harness error", "why": str(e)[:300], "errors": list(self.p.errors)}
            ok["kind"] = kind
            self.results.append(ok)
            self.log(ok)
        items = res.get("popup") or []
        if sub:
            # A submenu entry opened its submenu: its entries are the ones to walk.
            items = [i for i in items if i.get("level", 0) >= 1]
        if depth <= 0 or not items:
            return
        picks = [i for i in items if not i["disabled"] and i["tag"] == "button"]
        if len(picks) > 40:
            limit = min(limit, 8)
        if len(picks) > limit:
            # Long lists (fonts, colours): the first, some in between and the last.
            step = len(picks) / limit
            picks = [picks[int(k * step)] for k in range(limit)]
        for it in picks:
            if it["id"] in SKIP:
                self.results.append({"path": [i for _, i in path] + [it["id"]], "outcome": "skipped (native window)"})
                continue
            self.walk(path + [("item", it["id"])], depth - 1 if not it.get("sub") else depth, it["id"], "entry", limit=limit if depth > 1 else 16, sub=bool(it.get("sub")))

    def log(self, r: dict) -> None:
        flags = []
        if r.get("errors"):
            flags.append("ERR")
        if r.get("model") and r.get("undo") is False:
            flags.append("UNDO-FAIL" + ("(key lost)" if r.get("undoKeyLost") else ""))
        if r.get("outcome"):
            flags.append(r["outcome"])
        what = "dialog " + r["dialog"] if r.get("dialog") else f"menu({len(r['popup'])})" if r.get("popup") else \
            "model" if r.get("model") else "ui" if r.get("ui") else "canvas" if r.get("canvas") else "-"
        print(" > ".join(r["path"]), r.get("variant", ""), "|", what, " ".join(flags), r.get("ms", ""), flush=True)

    def tabs(self) -> list[str]:
        return [t["label"] for t in self.p.js("__sweep.tabs()") if t["label"] != "Datei"]

    def sweep_tab(self, tab: str, only: str | None) -> None:
        self.reset()
        self.step("tab", tab)
        ctrls = self.p.js("__sweep.controls()")
        self.inventory[tab] = ctrls
        for c in ctrls:
            if only and only not in c["id"]:
                continue
            if c["kind"] in ("combo",):
                # The text box itself: the list opens from its arrow (Öffnen) inside.
                self.walk([("tab", tab), ("ctrl", c["id"])], 1, c["id"], c["kind"])
                continue
            self.walk([("tab", tab), ("ctrl", c["id"])], 2, c["id"], c["kind"])

    def close(self) -> None:
        self.exe.close()


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--tabs", default="")
    ap.add_argument("--only", default=None)
    ap.add_argument("--context", action="store_true", help="context menus only")
    ap.add_argument("--out", default="calc-sweep")
    ap.add_argument("--dev", default=None, help="port:url of a dev page in a headless browser")
    a = ap.parse_args()
    s = Sweep(dev=a.dev)
    t0 = time.perf_counter()
    try:
        tabs = [t for t in s.tabs() if not a.tabs or t in a.tabs.split(",")]
        if not a.context:
            for t in tabs:
                s.sweep_tab(t, a.only)
        if a.context or not a.tabs:
            for where in ("cell", "row", "col", "sheettab", "statusbar"):
                s.walk([("ctx", where)], 2, where, "context")
    finally:
        s.close()
    out = HERE / "results" / f"2026-11-09-{a.out}.json"
    summary = summarise(s.results)
    summary["seconds"] = round(time.perf_counter() - t0)
    out.write_text(json.dumps({"summary": summary, "results": s.results, "inventory": s.inventory}, ensure_ascii=False, indent=1), encoding="utf-8", newline="\n")
    print(json.dumps(summary, ensure_ascii=False, indent=1))


def summarise(results: list[dict]) -> dict:
    n = len(results)
    errs = [r for r in results if r.get("errors")]
    undo = [r for r in results if r.get("model") and r.get("undo") is False]
    none = [r for r in results if not r.get("outcome") and not (r.get("model") or r.get("ui") or r.get("canvas") or r.get("dialog") or r.get("popup") or r.get("native"))]
    return {"options": n, "errors": len(errs), "undoFailures": len(undo), "noEffect": len(none),
            "disabled": sum(1 for r in results if r.get("outcome") == "disabled"),
            "notFound": sum(1 for r in results if r.get("outcome") in ("not found", "covered"))}


if __name__ == "__main__":
    main()
