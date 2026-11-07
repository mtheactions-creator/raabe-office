"""RaabePresent's ribbon and context menus as data (`__present.inventory()`, `__present.contextMenus()`) for
inventory_report_present.py (D-070). Runs the dev build in a headless Edge (npm run dev:present); no
input is sent.

    .venv/Scripts/python.exe inventory_present.py

Result: results/<date>-present-inventory.json.
"""
from __future__ import annotations

import datetime
import json
from pathlib import Path

from feellab.present_page import PresentPage

HERE = Path(__file__).parent


def main() -> None:
    page = PresentPage(1600, 1000, 1.0)
    try:
        # PowerPoint was recorded in its default design's colours (Office: Dunkelblaugrün, Orange …); ours starts in
        # Office 2013-2022. The same colour set is applied first, so colour entries are named alike (D-072).
        page.js("(async () => { const m = await import('/src/themesets.gen.ts'); const s = m.COLOR_SETS.find((c) => c.name === 'Office'); __present.app.applyColorSet(s.name, s.colors); })()")
        tabs = page.js("__present.inventory()")
        menus = page.js("__present.contextMenus()")
    finally:
        page.close()
    out = HERE / "results" / f"{datetime.date.today().isoformat()}-present-inventory.json"
    out.write_text(json.dumps({"tabs": tabs, "context_menus": menus}, ensure_ascii=False, indent=1), encoding="utf-8", newline=chr(10))
    print(len(tabs), "tabs,", len(menus), "context menus ->", out.name)


if __name__ == "__main__":
    main()
