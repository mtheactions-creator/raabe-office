"""Reads Word's user interface through Windows UI Automation: element names, types,
positions, sizes and access keys (Alt shortcuts). Selecting a ribbon tab uses the
automation API, not simulated input, so this never takes over keyboard or mouse."""

from __future__ import annotations

import comtypes.client

comtypes.client.GetModule("UIAutomationCore.dll")
from comtypes.gen import UIAutomationClient as U  # noqa: E402

_uia = comtypes.client.CreateObject(U.CUIAutomation, interface=U.IUIAutomation)

CONTROL_TYPES = {v: k[4:-len("ControlTypeId")] for k, v in vars(U).items() if k.startswith("UIA_") and k.endswith("ControlTypeId")}


def element_for(hwnd: int):
    return _uia.ElementFromHandle(hwnd)


def describe(el) -> dict:
    r = el.CurrentBoundingRectangle
    return {
        "type": CONTROL_TYPES.get(el.CurrentControlType, str(el.CurrentControlType)),
        "name": el.CurrentName or "",
        "id": el.CurrentAutomationId or "",
        "class": el.CurrentClassName or "",
        "key": el.CurrentAccessKey or "",
        "enabled": bool(el.CurrentIsEnabled),
        "rect": [r.left, r.top, r.right - r.left, r.bottom - r.top],
    }


def children(el) -> list:
    walker = _uia.ControlViewWalker
    out = []
    c = walker.GetFirstChildElement(el)
    while c:
        out.append(c)
        c = walker.GetNextSiblingElement(c)
    return out


def tree(el, depth: int = 0, max_depth: int = 12) -> dict:
    node = describe(el)
    if depth < max_depth:
        kids = [tree(c, depth + 1, max_depth) for c in children(el)]
        if kids:
            node["children"] = kids
    return node


def find(el, predicate, max_depth: int = 12):
    """Depth-first search for the first element matching predicate(describe(el))."""
    stack = [(el, 0)]
    while stack:
        e, d = stack.pop()
        if predicate(describe(e)):
            return e
        if d < max_depth:
            stack.extend((c, d + 1) for c in reversed(children(e)))
    return None


def select(el) -> None:
    pattern = el.GetCurrentPattern(U.UIA_SelectionItemPatternId)
    pattern.QueryInterface(U.IUIAutomationSelectionItemPattern).Select()


def invoke(el) -> None:
    pattern = el.GetCurrentPattern(U.UIA_InvokePatternId)
    pattern.QueryInterface(U.IUIAutomationInvokePattern).Invoke()
