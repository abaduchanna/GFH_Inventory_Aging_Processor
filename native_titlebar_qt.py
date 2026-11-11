"""Native Windows title-bar theming for PySide6 windows.

Same DWM calls as native_titlebar.py (Tkinter version) - the Windows API
doesn't care which GUI toolkit owns the HWND, so this is the same ctypes
logic, just getting the handle via Qt's winId() instead of Tk's
winfo_id().
"""
from __future__ import annotations

import ctypes
import os


GA_ROOT = 2
SWP_FLAGS = 0x0001 | 0x0002 | 0x0004 | 0x0020 | 0x0040


def _colourref(value: str) -> int:
    value = value.lstrip("#")
    red, green, blue = (int(value[i:i + 2], 16) for i in (0, 2, 4))
    return red | (green << 8) | (blue << 16)


def _apply(hwnd: int, dark: bool) -> None:
    if os.name != "nt":
        return
    try:
        user32 = ctypes.windll.user32
        root_hwnd = user32.GetAncestor(hwnd, GA_ROOT)
        if root_hwnd:
            hwnd = root_hwnd

        dwm = ctypes.windll.dwmapi
        enabled = ctypes.c_int(1 if dark else 0)
        result = dwm.DwmSetWindowAttribute(
            hwnd, 20, ctypes.byref(enabled), ctypes.sizeof(enabled))
        if result != 0:
            dwm.DwmSetWindowAttribute(
                hwnd, 19, ctypes.byref(enabled), ctypes.sizeof(enabled))

        try:
            ctypes.windll.uxtheme.SetWindowTheme(
                hwnd, "DarkMode_Explorer" if dark else "Explorer", None)
        except Exception:
            pass

        caption = ctypes.c_uint(_colourref("#090d26" if dark else "#eef1f7"))
        text = ctypes.c_uint(_colourref("#ffffff" if dark else "#10131f"))
        border = ctypes.c_uint(_colourref("#090d26" if dark else "#d5d9e5"))
        for attribute, colour in ((35, caption), (36, text), (34, border)):
            dwm.DwmSetWindowAttribute(
                hwnd, attribute, ctypes.byref(colour), ctypes.sizeof(colour))
        user32.SetWindowPos(hwnd, 0, 0, 0, 0, 0, SWP_FLAGS)
    except Exception:
        pass


def sync_titlebar_qt(window, dark: bool) -> None:
    """Match native Windows chrome to the app's current light/dark theme.

    `window` is any QWidget with a native window handle (QMainWindow,
    QDialog, ...). Call again whenever the theme toggles.
    """
    if os.name != "nt":
        return
    try:
        hwnd = int(window.winId())
        _apply(hwnd, dark)
    except Exception:
        pass
