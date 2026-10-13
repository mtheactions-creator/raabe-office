"""Feel-lab entry point.

    python run.py fidelity            # F4, no input needed
    python run.py docx [file.docx …]  # F4b, Word vs RaabeWrite on real files, no input needed
    python run.py startup             # F5
    python run.py latency             # F1, takes over keyboard and screen
    python run.py behaviour           # F3, takes over keyboard, mouse and screen
    python run.py formatting          # F3b, formatting commands vs Word, no input needed

Results are written to results/<date>-<test>.json.
"""

from __future__ import annotations

import datetime as dt
import json
import sys
from pathlib import Path

from feellab import targets

RESULTS = Path(__file__).parent / "results"


def save(test: str, data: dict) -> Path:
    RESULTS.mkdir(exist_ok=True)
    path = RESULTS / f"{dt.date.today().isoformat()}-{test}.json"
    meta = {"date": dt.datetime.now().isoformat(timespec="seconds"), "test": test}
    path.write_text(json.dumps({"meta": meta, "results": data}, ensure_ascii=False, indent=2), encoding="utf-8")
    return path


def with_targets(fn, fixture="short", quiet=False):
    """Word as the reference plus the real RaabeWrite app (the Step 0 prototypes are retired).
    `quiet`: Word invisible and RaabeWrite off screen, for tests without real input."""
    word = targets.WordTarget(fixture, visible=not quiet)
    protos = []
    try:
        protos.append(targets.raabe_write(fixture, port=9333, background=quiet))
        return fn(word, protos)
    finally:
        for p in protos:
            p.close()
        word.close()


def main(argv: list[str]) -> None:
    test = argv[1] if len(argv) > 1 else "fidelity"
    if test == "fidelity":
        from feellab import fidelity
        data = with_targets(lambda w, ps: fidelity.run(w, ps, ["short", "long"]), quiet=True)
    elif test == "startup":
        from feellab import startup
        data = startup.run()
    elif test == "latency":
        from feellab import latency
        data = latency.run()
    elif test == "docx":
        from feellab import docx_fidelity
        fixtures = targets.ROOT / "packages/ooxml/test/fixtures"
        files = [Path(a).resolve() for a in argv[2:]] or [fixtures / "word-layout.docx"]
        raabe = targets.raabe_write("short", port=9333, background=True)
        try:
            data = docx_fidelity.run(raabe, files)
        finally:
            raabe.close()
        for name, r in data.items():
            print(name, json.dumps({k: v for k, v in r.items() if k != "lines"}, ensure_ascii=False, indent=1)[:5000])
        print("saved to", save(test, data))
        return
    elif test == "formatting":
        from feellab import formatting
        word = targets.WordTarget("short", visible=False)
        ours = targets.raabe_write("short", port=9333, background=True)
        try:
            data = formatting.run(word, ours)
        finally:
            ours.close()
            word.close()
        print("saved to", save(test, data))
        return
    elif test == "behaviour":
        from feellab import behaviour
        data = with_targets(lambda w, ps: behaviour.run([w, *ps]))
    else:
        raise SystemExit(__doc__)
    print(json.dumps(data, ensure_ascii=False, indent=2)[:6000])
    print("saved to", save(test, data))


if __name__ == "__main__":
    main(sys.argv)
