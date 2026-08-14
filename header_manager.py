"""
Header Manager - Standardized for GFH/VidaPay Ecosystem
Developed by Abad Umair Channa | Copyright © {year} | All rights reserved.
"""
import os

# ── Lazy tkinter import ──
# tkinter is imported inside methods, not at module level.


class FixedHeaderManager:
    """Manages header with centered title, logo, and theme toggle."""

    BRAND_NAVY = "#090d26"
    BRAND_RED = "#f0541c"
    BRAND_WHITE = "#ffffff"

    def __init__(self, parent, title="App", height=108):
        import tkinter as tk

        self.parent = parent
        self.height = height
        self.header_frame = tk.Frame(parent, height=height, bg=self.BRAND_NAVY)
        self.header_frame.pack(side=tk.TOP, fill=tk.X)
        self.header_frame.pack_propagate(False)

        # LEFT: Logo (larger for visibility)
        self.logo_frame = tk.Frame(self.header_frame, bg=self.BRAND_NAVY, width=200)
        self.logo_frame.pack(side=tk.LEFT, fill=tk.Y, padx=(12, 0))
        self.logo_frame.pack_propagate(False)

        self.logo_label = tk.Label(
            self.logo_frame,
            bg=self.BRAND_NAVY,
            fg=self.BRAND_RED,
            font=("Segoe UI", 10, "bold"),
            borderwidth=0,
        )
        self.logo_label.pack(expand=True, anchor="w")

        # CENTER: Title
        self.title_label = tk.Label(
            self.header_frame,
            text=title,
            bg=self.BRAND_NAVY,
            fg=self.BRAND_WHITE,
            font=("Segoe UI", 16, "bold"),
            borderwidth=0,
        )
        self.title_label.pack(side=tk.LEFT, expand=True, padx=(0, 0))

        # RIGHT: Theme toggle + Copyright
        self.right_frame = tk.Frame(self.header_frame, bg=self.BRAND_NAVY)
        self.right_frame.pack(side=tk.RIGHT, fill=tk.Y, padx=(0, 12))

    def set_logo(self, logo_path=None, text="Logo"):
        """Set the logo in the header. Visible size: 120x45 pixels."""
        if logo_path and os.path.exists(logo_path):
            try:
                from PIL import Image, ImageTk
                img = Image.open(logo_path)
                # Larger thumbnail for header visibility
                img.thumbnail((120, 45), Image.Resampling.LANCZOS)
                self.photo = ImageTk.PhotoImage(img)
                self.logo_label.configure(image=self.photo, text="")
            except Exception:
                self.logo_label.configure(text=text, fg=self.BRAND_RED)
        else:
            self.logo_label.configure(text=text, fg=self.BRAND_RED)

    def set_title(self, title):
        """Update the header title."""
        self.title_label.configure(text=title)

    def add_copyright(self, text=None):
        """Add copyright text to the right side of the header."""
        import tkinter as tk

        if text is None:
            from datetime import datetime
            text = f"© {datetime.now().year} Developed by Abad Umair Channa"

        lbl = tk.Label(
            self.right_frame,
            text=text,
            bg=self.BRAND_NAVY,
            fg="#9d9db8",
            font=("Segoe UI", 7),
            borderwidth=0,
        )
        lbl.pack(side=tk.BOTTOM, pady=(0, 4))
        return lbl

    def add_theme_toggle(self, theme_manager, callback=None):
        """Add theme toggle button to header."""
        import tkinter as tk

        btn = theme_manager.create_theme_toggle_button(self.right_frame, callback)
        btn.pack(side=tk.TOP, pady=(8, 0))
        return btn

    def get_frame(self):
        """Return the header frame for packing additional widgets."""
        return self.header_frame
