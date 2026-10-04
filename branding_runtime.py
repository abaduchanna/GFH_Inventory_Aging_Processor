"""Shared 3S Verse native window branding and animated background.

The animation intentionally uses the same artwork and timing as the
3SVerse WiFi Transfer / 3sverse.com hero: 120-second spiral rotation,
140-second orb rotation, and a 12 px / 13-second orb float.
"""
from __future__ import annotations

import ctypes
import math
import os
import sys
from pathlib import Path


BRAND_DARK = "#07060b"
BRAND_LIGHT = "#f6f7fb"
BRAND_TEXT_DARK = "#10131f"
BRAND_TEXT_LIGHT = "#ffffff"


def _asset_path(name: str) -> str:
    base = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
    return str(base / name)


def _colourref(hex_colour: str) -> int:
    value = hex_colour.lstrip("#")
    red, green, blue = (int(value[i:i + 2], 16) for i in (0, 2, 4))
    return red | (green << 8) | (blue << 16)


def _top_level_hwnd(hwnd: int) -> int:
    if os.name != "nt":
        return hwnd
    user32 = ctypes.windll.user32
    parent = user32.GetParent(hwnd)
    while parent:
        hwnd = parent
        parent = user32.GetParent(hwnd)
    return hwnd


def apply_native_window_branding(window, dark: bool = True) -> None:
    """Set the exact website favicon and native Windows title-bar colours."""
    try:
        icon = _asset_path("3sverse_website_favicon.ico")
        if os.path.exists(icon):
            window.iconbitmap(default=icon)
    except Exception:
        pass

    if os.name != "nt":
        return
    try:
        window.update_idletasks()
        hwnd = _top_level_hwnd(int(window.winfo_id()))
        dwm = ctypes.windll.dwmapi
        enabled = ctypes.c_int(1 if dark else 0)
        # DWMWA_USE_IMMERSIVE_DARK_MODE: 20 on current Windows, 19 on
        # older Windows 10 releases.
        result = dwm.DwmSetWindowAttribute(
            hwnd, 20, ctypes.byref(enabled), ctypes.sizeof(enabled))
        if result != 0:
            dwm.DwmSetWindowAttribute(
                hwnd, 19, ctypes.byref(enabled), ctypes.sizeof(enabled))

        caption = ctypes.c_uint(_colourref("#090d26" if dark else "#eef1f7"))
        text = ctypes.c_uint(_colourref(BRAND_TEXT_LIGHT if dark else BRAND_TEXT_DARK))
        border = ctypes.c_uint(_colourref("#090d26" if dark else "#d5d9e5"))
        for attribute, value in ((35, caption), (36, text), (34, border)):
            dwm.DwmSetWindowAttribute(
                hwnd, attribute, ctypes.byref(value), ctypes.sizeof(value))
        ctypes.windll.user32.SetWindowPos(
            hwnd, 0, 0, 0, 0, 0,
            0x0001 | 0x0002 | 0x0004 | 0x0020 | 0x0040)
    except Exception:
        pass


class BrandAnimation:
    """Exact WiFi Transfer ring/orb animation for a Tk canvas."""

    RING_FRAMES = 24
    RING_MS = 5000                 # 24 x 5s = 120s / revolution
    ORB_TICK_MS = 40
    ORB_AMPLITUDE = 12
    ORB_PERIOD = 13.0
    ORB_SPIN_TICKS = 146           # ~140s / revolution
    RING_FLOAT_AMPLITUDE = 16
    RING_FLOAT_PERIOD = 12.0

    def __init__(self, owner, canvas, title: str | None = None):
        self.owner = owner
        self.canvas = canvas
        self.title = title
        self._images = []
        self._ring_frames = []
        self._orb_frames = []
        self._ring_index = 0
        self._orb_index = 0
        self._orb_ticks = 0
        self._orb_phase = 0.0
        self._ring_phase = 0.0
        self._ring_x = 0.0
        self._ring_y = 0.0
        self._ring_item = None
        self._orb_item = None
        self._orb_x = 0.0
        self._orb_y = 0.0
        self._load()
        canvas.bind("<Configure>", self._redraw, add="+")
        owner.after_idle(self._redraw)
        owner.after(self.RING_MS, self._animate_ring)
        owner.after(self.ORB_TICK_MS, self._animate_orb)

    def _load(self):
        try:
            from PIL import Image, ImageTk
            resampling = getattr(Image, "Resampling", Image)
            for filename, maximum, target in (
                    ("brand_bg_ring.png", 600, self._ring_frames),
                    ("brand_bg_orb.png", 460, self._orb_frames)):
                image = Image.open(_asset_path(filename)).convert("RGBA")
                if max(image.size) > maximum:
                    scale = maximum / float(max(image.size))
                    image = image.resize(
                        (max(1, int(image.width * scale)),
                         max(1, int(image.height * scale))),
                        resampling.LANCZOS)
                for index in range(self.RING_FRAMES):
                    angle = round(index * 360.0 / self.RING_FRAMES)
                    frame = image.rotate(
                        angle, resample=resampling.BICUBIC, expand=False)
                    target.append(ImageTk.PhotoImage(frame))
        except Exception:
            # Static Tk images are still preferable to losing branding when
            # Pillow is unavailable in a development environment.
            try:
                import tkinter as tk
                self._ring_frames = [tk.PhotoImage(file=_asset_path("brand_bg_ring.png"))]
                self._orb_frames = [tk.PhotoImage(file=_asset_path("brand_bg_orb.png"))]
            except Exception:
                self._ring_frames = []
                self._orb_frames = []

    def _redraw(self, _event=None):
        try:
            canvas = self.canvas
            canvas.delete("brand_art")
            width = max(2, int(canvas.winfo_width()))
            height = max(2, int(canvas.winfo_height()))
            if self._ring_frames:
                ring = self._ring_frames[self._ring_index % len(self._ring_frames)]
                self._ring_x = width + ring.width() * 0.16
                self._ring_y = height * 0.44
                self._ring_item = canvas.create_image(
                    self._ring_x, self._ring_y,
                    image=ring, anchor="center", tags=("brand_art",))
            if self._orb_frames:
                orb = self._orb_frames[self._orb_index % len(self._orb_frames)]
                self._orb_x = -orb.width() * 0.22
                self._orb_y = height - orb.height() * 0.42
                self._orb_item = canvas.create_image(
                    self._orb_x, self._orb_y, image=orb, anchor="center",
                    tags=("brand_art",))
            if self.title:
                canvas.create_text(
                    width / 2, height / 2, text=self.title,
                    font=("Segoe UI", 16, "bold"), fill="#ffffff",
                    tags=("brand_art", "brand_title"))
        except Exception:
            pass

    def _active(self):
        try:
            return bool(self.owner.winfo_exists()) and self.owner.state() != "iconic"
        except Exception:
            return False

    def _animate_ring(self):
        try:
            if self._active() and len(self._ring_frames) > 1 and self._ring_item:
                self._ring_index = (self._ring_index + 1) % len(self._ring_frames)
                self.canvas.itemconfigure(
                    self._ring_item, image=self._ring_frames[self._ring_index])
        except Exception:
            pass
        try:
            self.owner.after(self.RING_MS, self._animate_ring)
        except Exception:
            pass

    def _animate_orb(self):
        try:
            if self._active() and self._orb_item:
                step = (2 * math.pi * (self.ORB_TICK_MS / 1000.0)
                        / self.ORB_PERIOD)
                self._orb_phase = (self._orb_phase + step) % (2 * math.pi)
                offset = math.sin(self._orb_phase) * self.ORB_AMPLITUDE
                self.canvas.coords(
                    self._orb_item, self._orb_x, self._orb_y + offset)
                if self._ring_item:
                    ring_step = (2 * math.pi * (self.ORB_TICK_MS / 1000.0)
                                 / self.RING_FLOAT_PERIOD)
                    self._ring_phase = (self._ring_phase + ring_step) % (2 * math.pi)
                    ring_offset = math.sin(self._ring_phase) * self.RING_FLOAT_AMPLITUDE
                    self.canvas.coords(
                        self._ring_item, self._ring_x, self._ring_y + ring_offset)
                self._orb_ticks += 1
                if len(self._orb_frames) > 1 and self._orb_ticks >= self.ORB_SPIN_TICKS:
                    self._orb_ticks = 0
                    self._orb_index = (self._orb_index + 1) % len(self._orb_frames)
                    self.canvas.itemconfigure(
                        self._orb_item, image=self._orb_frames[self._orb_index])
        except Exception:
            pass
        try:
            self.owner.after(self.ORB_TICK_MS, self._animate_orb)
        except Exception:
            pass


def install_branding(window, dark: bool = True, background: str | None = None):
    """Install/update native chrome and one shared root background."""
    apply_native_window_branding(window, dark)
    # Tk can recreate its wrapper HWND while mapping. Re-apply once after
    # realization so the title bar never remains light in dark mode.
    try:
        pending = getattr(window, "_3sv_native_after", None)
        if pending:
            window.after_cancel(pending)
        window._3sv_native_after = window.after(
            120, lambda: apply_native_window_branding(window, dark))
    except Exception:
        pass

    if getattr(window, "_3sv_brand_animation", None) is not None:
        return
    try:
        import tkinter as tk
        canvas = tk.Canvas(
            window, bg=background or (BRAND_DARK if dark else BRAND_LIGHT),
            highlightthickness=0, borderwidth=0, bd=0)
        canvas.place(x=0, y=0, relwidth=1, relheight=1)
        window.tk.call("lower", canvas._w)
        window._3sv_brand_canvas = canvas
        window._3sv_brand_animation = BrandAnimation(window, canvas)
    except Exception:
        pass


def install_header_animation(header_manager):
    """Put the same animation visibly inside a FixedHeaderManager band."""
    try:
        import tkinter as tk
        old = getattr(header_manager, "texture_canvas", None)
        if old is not None:
            old.destroy()
        canvas = tk.Canvas(
            header_manager.header_frame, bg=header_manager.BRAND_NAVY,
            highlightthickness=0, borderwidth=0, bd=0)
        canvas.place(x=0, y=0, relwidth=1, relheight=1)
        header_manager.texture_canvas = canvas
        header_manager._brand_animation = BrandAnimation(
            header_manager.parent, canvas, header_manager.title)
        for edge_name in ("left_frame", "divider_frame", "right_frame"):
            edge = getattr(header_manager, edge_name, None)
            if edge is not None:
                header_manager.header_frame.tk.call("raise", edge._w)
    except Exception:
        pass
