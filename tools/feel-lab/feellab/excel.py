"""Microsoft Excel as the reference for RaabeCalc (D-061), driven through COM.

The harness always starts an Excel process of its own (`DispatchEx`) and closes only the
workbooks it created, as for Word (CLAUDE.md, harness safety). Cells are addressed as Excel
addresses ("B3"); values come back as Excel shows them (`Text`) and as stored (`Value2`,
`Formula`, `FormulaLocal`)."""

from __future__ import annotations

import time

from . import winapi

WINDOW_RECT = (0, 0, 2000, 2050)


class ExcelTarget:
    name = "excel"

    def __init__(self, visible: bool = True, place: bool = True):
        import win32com.client
        self.app = win32com.client.gencache.EnsureDispatch(win32com.client.DispatchEx("Excel.Application"))
        try:
            self.app.Visible = visible
            self.app.DisplayAlerts = False
            self.wb = self.app.Workbooks.Add()
            self.hwnd = int(self.app.Hwnd)
            self.pid = _pid(self.hwnd)
            if visible and place:
                winapi.place(self.hwnd, *WINDOW_RECT)
        except Exception:
            self.app.Quit()
            raise
        self._options = {}

    @property
    def sheet(self):
        return self.wb.ActiveSheet

    def remember_option(self, name: str) -> None:
        """Excel keeps application options for the user: note one before a probe changes it."""
        if name not in self._options:
            self._options[name] = getattr(self.app, name)

    def set_option(self, name: str, value) -> None:
        self.remember_option(name)
        setattr(self.app, name, value)

    def cell(self, addr: str) -> dict:
        r = self.sheet.Range(addr)
        return {"text": r.Text, "value": r.Value2, "formula": r.Formula, "local": r.FormulaLocal,
                "format": r.NumberFormatLocal}

    def set(self, addr: str, value) -> None:
        self.sheet.Range(addr).FormulaLocal = value

    def selection(self) -> tuple[str, str]:
        """(selected range, active cell) as Excel addresses without $."""
        sel = self.app.Selection
        return sel.GetAddress(False, False), self.app.ActiveCell.GetAddress(False, False)

    def clear(self) -> None:
        self.sheet.Cells.Clear()
        self.sheet.Range("A1").Select()

    def close(self) -> None:
        for name, value in self._options.items():
            try:
                setattr(self.app, name, value)
            except Exception:  # noqa: BLE001
                pass
        try:
            self.wb.Close(False)
        finally:
            self.app.Quit()
            self.app = self.wb = None
            # Excel lingers after Quit while COM references live; end only our own process.
            import subprocess
            for _ in range(30):
                time.sleep(0.2)
                if not _alive(self.pid):
                    return
            subprocess.run(["taskkill", "/PID", str(self.pid), "/F"], capture_output=True)


def _pid(hwnd: int) -> int:
    import ctypes
    pid = ctypes.c_ulong()
    winapi.user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
    return pid.value


def _alive(pid: int) -> bool:
    import subprocess
    out = subprocess.run(["tasklist", "/FI", f"PID eq {pid}", "/NH"], capture_output=True, text=True, errors="replace").stdout
    return str(pid) in out
