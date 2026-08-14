"""
Logo Handler - Standardized for GFH/VidaPay Ecosystem
Developed by Abad Umair Channa | Copyright © {year} | All rights reserved.
"""
import os

# ── Lazy tkinter import ──
# tkinter is imported inside methods, not at module level.


class LogoHandler:
    """Manages logo display with theme compatibility."""

    def __init__(self, parent):
        self.parent = parent
        self.logo_widget = None
        self.photo_image = None

    def load_logo_from_file(self, logo_path, width=120, height=45):
        """Load and display logo from file (PNG/JPG)."""
        if not logo_path or not os.path.exists(logo_path):
            return False
        try:
            from PIL import Image, ImageTk
            import tkinter as tk

            img = Image.open(logo_path)
            img.thumbnail((width, height), Image.Resampling.LANCZOS)
            self.photo_image = ImageTk.PhotoImage(img)
            self.logo_widget = tk.Label(
                self.parent,
                image=self.photo_image,
                bg="#090d26",
                borderwidth=0,
            )
            return True
        except Exception:
            return False

    def create_text_placeholder(self, text="App", color="#090d26", size=12):
        """Create text placeholder when image unavailable."""
        import tkinter as tk

        self.logo_widget = tk.Label(
            self.parent,
            text=text,
            font=("Segoe UI", size, "bold"),
            fg=color,
            bg="#090d26",
            borderwidth=0,
        )

    def pack(self, side="left", padx=0, pady=0):
        """Pack the logo widget."""
        if self.logo_widget:
            self.logo_widget.pack(side=side, padx=padx, pady=pady)

    def apply_theme(self, colors):
        """Logo doesn't change colors - theme changes only affect other widgets."""
        pass


def add_logo_to_header(header_frame, logo_path=None, text_fallback="Logo"):
    """
    Easy function to add logo to header.

    Args:
        header_frame: tk.Frame to add logo to
        logo_path: Optional path to logo image file
        text_fallback: Text to show if no image (app-specific)

    Returns:
        LogoHandler instance
    """
    handler = LogoHandler(header_frame)
    if logo_path and os.path.exists(logo_path):
        if handler.load_logo_from_file(logo_path):
            handler.pack(side="left", padx=12, pady=8)
        else:
            handler.create_text_placeholder(text_fallback)
            handler.pack(side="left", padx=12, pady=8)
    else:
        handler.create_text_placeholder(text_fallback)
        handler.pack(side="left", padx=12, pady=8)
    return handler
