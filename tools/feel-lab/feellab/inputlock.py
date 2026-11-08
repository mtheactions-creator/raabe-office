"""The real-input lock (D-070, D-080): only one session at a time may move the mouse and press keys. A run that uses
real input takes %TEMP%/raabe-real-input.lock exclusively (holder and start time inside) and deletes it when done;
a lock older than 30 minutes counts as left over from a run that died.

    with real_input("drag test"):
        ...
"""
from __future__ import annotations

import contextlib
import os
import tempfile
import time
from pathlib import Path

LOCK = Path(tempfile.gettempdir()) / "raabe-real-input.lock"
STALE = 30 * 60


def holder() -> str | None:
    try:
        return LOCK.read_text(encoding="utf-8")
    except OSError:
        return None


@contextlib.contextmanager
def real_input(who: str, wait: float = 3600, poll: float = 5):
    """Takes the lock (waiting up to `wait` seconds for another holder); releases it afterwards."""
    end = time.time() + wait
    while True:
        try:
            fd = os.open(LOCK, os.O_CREAT | os.O_EXCL | os.O_WRONLY)
            with os.fdopen(fd, "w", encoding="utf-8") as f:
                f.write(f"{who} {time.strftime('%Y-%m-%d %H:%M:%S')}")
            break
        except FileExistsError:
            try:
                if time.time() - LOCK.stat().st_mtime > STALE:
                    LOCK.unlink(missing_ok=True)
                    continue
            except OSError:
                continue
            if time.time() > end:
                raise TimeoutError(f"real-input lock held by: {holder()}")
            time.sleep(poll)
    try:
        yield
    finally:
        LOCK.unlink(missing_ok=True)


def touch() -> None:
    """Keeps a long run's lock fresh (call now and then, so it is not taken for left over)."""
    with contextlib.suppress(OSError):
        os.utime(LOCK)
