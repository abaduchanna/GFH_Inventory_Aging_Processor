"""Native Windows title-bar theming for partner-branded desktop apps."""
from __future__ import annotations

import ctypes
import os


def _colourref(value: str) -> int:
    value = value.lstrip("#")
    red, green, blue = (int(value[i:i + 2], 16) for i in (0, 2, 4))
    return red | (green << 8) | (blue << 16)


def _apply(window, dark: bool) -> None:
    if os.name != "nt":
        return
    try:
        window.update_idletasks()
        hwnd = int(window.winfo_id())
        user32 = ctypes.windll.user32
        parent = user32.GetParent(hwnd)
        while parent:
            hwnd = parent
            parent = user32.GetParent(hwnd)

        dwm = ctypes.windll.dwmapi
        enabled = ctypes.c_int(1 if dark else 0)
        result = dwm.DwmSetWindowAttribute(
            hwnd, 20, ctypes.byref(enabled), ctypes.sizeof(enabled))
        if result != 0:
            dwm.DwmSetWindowAttribute(
                hwnd, 19, ctypes.byref(enabled), ctypes.sizeof(enabled))

        caption = ctypes.c_uint(_colourref("#090d26" if dark else "#eef1f7"))
        text = ctypes.c_uint(_colourref("#ffffff" if dark else "#10131f"))
        border = ctypes.c_uint(_colourref("#090d26" if dark else "#d5d9e5"))
        for attribute, colour in ((35, caption), (36, text), (34, border)):
            dwm.DwmSetWindowAttribute(
                hwnd, attribute, ctypes.byref(colour), ctypes.sizeof(colour))
        user32.SetWindowPos(
            hwnd, 0, 0, 0, 0, 0,
            0x0001 | 0x0002 | 0x0004 | 0x0020 | 0x0040)
    except Exception:
        pass


def sync_titlebar(window, dark: bool) -> None:
    """Match native Windows chrome to the app's current light/dark theme."""
    _apply(window, dark)
    try:
        pending = getattr(window, "_partner_titlebar_after", None)
        if pending:
            window.after_cancel(pending)
        window._partner_titlebar_after = window.after(
            120, lambda: _apply(window, dark))
    except Exception:
        pass
