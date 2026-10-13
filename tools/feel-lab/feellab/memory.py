"""Memory use: RaabeWrite (its process and the WebView2 processes it starts) and Word (its
process), each with the same document open. Working set and private bytes, summed over the
process tree, read with the Windows API (no extra packages)."""

from __future__ import annotations

import ctypes
import time
from ctypes import wintypes

from . import targets
from .docx_fidelity import WordDocument, open_in_raabe

DOC = targets.ROOT / "packages/ooxml/test/fixtures/word-tables.docx"

kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
psapi = ctypes.WinDLL("psapi", use_last_error=True)


class PROCESSENTRY32W(ctypes.Structure):
    _fields_ = [("dwSize", wintypes.DWORD), ("cntUsage", wintypes.DWORD), ("th32ProcessID", wintypes.DWORD),
                ("th32DefaultHeapID", ctypes.c_size_t), ("th32ModuleID", wintypes.DWORD), ("cntThreads", wintypes.DWORD),
                ("th32ParentProcessID", wintypes.DWORD), ("pcPriClassBase", ctypes.c_long), ("dwFlags", wintypes.DWORD),
                ("szExeFile", ctypes.c_wchar * 260)]


class PROCESS_MEMORY_COUNTERS_EX(ctypes.Structure):
    _fields_ = [("cb", wintypes.DWORD), ("PageFaultCount", wintypes.DWORD), ("PeakWorkingSetSize", ctypes.c_size_t),
                ("WorkingSetSize", ctypes.c_size_t), ("QuotaPeakPagedPoolUsage", ctypes.c_size_t),
                ("QuotaPagedPoolUsage", ctypes.c_size_t), ("QuotaPeakNonPagedPoolUsage", ctypes.c_size_t),
                ("QuotaNonPagedPoolUsage", ctypes.c_size_t), ("PagefileUsage", ctypes.c_size_t),
                ("PeakPagefileUsage", ctypes.c_size_t), ("PrivateUsage", ctypes.c_size_t)]


def _processes() -> list[tuple[int, int, str]]:
    snap = kernel32.CreateToolhelp32Snapshot(0x2, 0)
    out = []
    e = PROCESSENTRY32W()
    e.dwSize = ctypes.sizeof(e)
    ok = kernel32.Process32FirstW(snap, ctypes.byref(e))
    while ok:
        out.append((e.th32ProcessID, e.th32ParentProcessID, e.szExeFile))
        ok = kernel32.Process32NextW(snap, ctypes.byref(e))
    kernel32.CloseHandle(snap)
    return out


def _memory(pid: int) -> tuple[int, int]:
    h = kernel32.OpenProcess(0x0400 | 0x0010, False, pid)
    if not h:
        return 0, 0
    try:
        c = PROCESS_MEMORY_COUNTERS_EX()
        c.cb = ctypes.sizeof(c)
        if not psapi.GetProcessMemoryInfo(h, ctypes.byref(c), c.cb):
            return 0, 0
        return c.WorkingSetSize, c.PrivateUsage
    finally:
        kernel32.CloseHandle(h)


def tree(pid: int) -> dict:
    procs = _processes()
    ids = {pid}
    changed = True
    while changed:
        changed = False
        for p, parent, _ in procs:
            if parent in ids and p not in ids:
                ids.add(p)
                changed = True
    ws = priv = 0
    for p in ids:
        a, b = _memory(p)
        ws += a
        priv += b
    return {"processes": len(ids), "working_set_mb": round(ws / 2**20, 1), "private_mb": round(priv / 2**20, 1)}


def run() -> dict:
    out: dict = {"document": DOC.name}
    raabe = targets.raabe_write("short", port=9333, background=True)
    try:
        time.sleep(3)
        out["raabe_start"] = tree(raabe.proc.pid)
        open_in_raabe(raabe, DOC)
        time.sleep(3)
        out["raabe_document"] = tree(raabe.proc.pid)
    finally:
        raabe.close()
    word = WordDocument(DOC)
    try:
        time.sleep(3)
        pid = wintypes.DWORD()
        ctypes.windll.user32.GetWindowThreadProcessId(word.win.Hwnd, ctypes.byref(pid))
        out["word_document"] = tree(pid.value)
    finally:
        word.close()
    return out
