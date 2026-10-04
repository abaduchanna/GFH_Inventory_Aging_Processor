"""Native Windows title-bar theming for partner-branded desktop apps."""
from __future__ import annotations

import ctypes
import os


GA_ROOT = 2
SWP_FLAGS = 0x0001 | 0x0002 | 0x0004 | 0x0020 | 0x0040


def _colourref(value: str) -> int:
    value = value.lstrip("#")
    red, green, blue = (int(value[i:i + 2], 16) for i in (0, 2, 4))
    return red | (green << 8) | (blue << 16)


def _apply(window, dark: bool) -> None:
    if os.name != "nt":
        return
    try:
        window.update_idletasks()
        user32 = ctypes.windll.user32
        hwnd = int(window.winfo_id())
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

        # Ask Windows common controls to use matching non-client chrome too.
        # This is important on Tk/PyInstaller windows whose wrapper HWND is
        # created after the first theme application.
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
        user32.SetWindowPos(
            hwnd, 0, 0, 0, 0, 0, SWP_FLAGS)
    except Exception:
        pass


def sync_titlebar(window, dark: bool) -> None:
    """Match native Windows chrome to the app's current light/dark theme."""
    window._partner_titlebar_dark = bool(dark)

    def refresh(_event=None):
        current = bool(getattr(window, "_partner_titlebar_dark", dark))
        _apply(window, current)

    try:
        if not getattr(window, "_partner_titlebar_bound", False):
            window.bind("<Map>", refresh, add="+")
            window.bind("<Visibility>", refresh, add="+")
            window._partner_titlebar_bound = True

        for pending in getattr(window, "_partner_titlebar_afters", ()):
            try:
                window.after_cancel(pending)
            except Exception:
                pass
        # Tk creates/remaps its real top-level wrapper asynchronously. Apply
        # now and again after mapping so the theme cannot fall back to light.
        window._partner_titlebar_afters = [
            window.after(delay, refresh) for delay in (0, 40, 160, 500, 1200)
        ]
    except Exception:
        refresh()
