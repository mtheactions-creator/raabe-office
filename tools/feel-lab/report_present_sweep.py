"""The by-hand sweep's report (D-137): merges sweep_present.py's result files into the option list
(results/<date>-present-options.json: every option reached through the UI, with how it fared, and the ribbon's defined
entries no walk reached), and opens the presentations the sweep saved in PowerPoint (an instance of our own, only when
no POWERPNT.EXE runs, windowless) to check that Office reads every one of them.

    .venv/Scripts/python.exe report_present_sweep.py merge OUT_NAME RESULT.json [RESULT.json ...]
    .venv/Scripts/python.exe report_present_sweep.py office SAVED_DIR [SAVED_DIR ...] [--max N]
"""
from __future__ import annotations

import argparse
import datetime
import hashlib
import json
import time
from pathlib import Path

HERE = Path(__file__).parent


def names_defined(defined: list) -> set[str]:
    out: set[str] = set()

    def walk(items):
        for it in items or []:
            if it.get("name"):
                out.add(it["name"].strip())
            walk(it.get("items"))
            walk(it.get("submenu"))
    for t in defined or []:
        for g in t["groups"]:
            walk(g["controls"])
    return out


def commands() -> list[dict]:
    """RaabePresent's command list, read from the dev page (npm run dev:present on port 5186, as the sweep's --dev)."""
    from sweep_present import start
    page = start(True, "http://localhost:5186/", 9509)
    try:
        return page.js("window.__present.app.commands.all().map((c) => ({ id: c.id, label: c.label, shortcut: c.shortcut || null, hidden: !!c.hidden, built: !!c.run, deferred: !!c.deferred }))")
    finally:
        page.close()


def merge(out_name: str, files: list[str]) -> None:
    results, defined = [], None
    for f in files:
        d = json.loads(Path(f).read_text(encoding="utf-8"))
        results += d["results"]
        defined = d.get("defined") or defined
    # One entry per option id (a later file wins: a rerun after fixes).
    by_id: dict = {}
    for r in results:
        by_id[r.get("id", "?")] = r
    opts = list(by_id.values())
    counts: dict = {}
    for r in opts:
        counts[r["status"]] = counts.get(r["status"], 0) + 1
    reached = {(r.get("name") or "").strip() for r in opts} | {i.strip() for r in opts for i in (r.get("items") or [])}
    reached |= {(r.get("name") or "").replace(" (Pfeil)", "").strip() for r in opts}
    unreached = sorted(n for n in names_defined(defined) if n not in reached and not n.endswith("...") or (n.endswith("...") and n not in reached and n[:-3] not in reached))
    unreached = [n for n in unreached if n not in reached]
    # The command list: commands no walk reached by their label (shortcut-only, menu-only under another name, or none).
    cmds = commands()
    cmd_unreached = sorted({f"{c['id']} ({c['label']})" for c in cmds if not c.get("hidden") and c["label"].strip() not in reached
                            and c["label"].replace("...", "").strip() not in reached and f"{c['label']} ({c.get('shortcut')})" not in reached})
    out = {"date": datetime.date.today().isoformat(), "options": len(opts), "counts": counts,
           "by_tab": {}, "defined_not_reached": unreached, "commands": len(cmds), "commands_not_reached_by_label": cmd_unreached,
           "options_list": [{k: r.get(k) for k in ("id", "tab", "group", "name", "kind", "ctx", "status", "why")} for r in opts]}
    for r in opts:
        t = out["by_tab"].setdefault(r.get("tab") or "?", {})
        t[r["status"]] = t.get(r["status"], 0) + 1
    p = HERE / "results" / f"{out['date']}-{out_name}.json"
    p.write_text(json.dumps(out, ensure_ascii=False, indent=1), encoding="utf-8", newline=chr(10))
    print(json.dumps({"options": len(opts), "counts": counts, "defined_not_reached": len(unreached)}, ensure_ascii=False), "->", p.name)


def office(dirs: list[str], limit: int) -> None:
    from feellab.powerpoint import OwnerPowerPointRunning, PowerPointTarget
    files, seen = [], set()
    for d in dirs:
        for f in sorted(Path(d).glob("*.pptx")):
            h = hashlib.md5(f.read_bytes()).hexdigest()
            if h not in seen:
                seen.add(h)
                files.append(f)
    if limit and len(files) > limit:
        step = len(files) / limit
        files = [files[int(i * step)] for i in range(limit)]
    try:
        t = PowerPointTarget(visible=False)
    except OwnerPowerPointRunning as e:
        print("PowerPoint:", e)
        return
    bad = []
    try:
        for i, f in enumerate(files):
            try:
                pres = t.open_copy(f, with_window=False)
                n = int(pres.Slides.Count)
                pres.Close()
                t.own.remove(pres)
            except Exception as e:  # noqa: BLE001
                bad.append({"file": f.name, "dir": f.parent.name, "error": str(e)[:200]})
                print("FAIL", f.parent.name, f.name, str(e)[:120], flush=True)
            if i % 50 == 0:
                print(i, "/", len(files), flush=True)
            time.sleep(0.05)
    finally:
        print(t.close(kill_if_unreachable=True))
    out = {"date": datetime.date.today().isoformat(), "opened": len(files) - len(bad), "total": len(files), "failed": bad}
    p = HERE / "results" / f"{out['date']}-present-sweep-office-open.json"
    p.write_text(json.dumps(out, ensure_ascii=False, indent=1), encoding="utf-8", newline=chr(10))
    print(out["opened"], "/", out["total"], "->", p.name)


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("what", choices=["merge", "office"])
    ap.add_argument("args", nargs="+")
    ap.add_argument("--max", type=int, default=400)
    a = ap.parse_args()
    if a.what == "merge":
        merge(a.args[0], a.args[1:])
    else:
        office(a.args, a.max)


if __name__ == "__main__":
    main()
