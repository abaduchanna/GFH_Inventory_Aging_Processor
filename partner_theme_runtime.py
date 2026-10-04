"""Runtime enforcement for partner footers and button hover colours."""
from __future__ import annotations


def enforce_brand_interactions(window, footer_bg: str, hover_bg: str) -> None:
    """Keep the footer dark and give every enabled Tk button a brand hover."""
    import tkinter as tk

    def paint_footer(widget) -> None:
        try:
            widget.configure(bg=footer_bg)
        except (tk.TclError, TypeError):
            pass
        try:
            widget.configure(fg="#ffffff")
        except (tk.TclError, TypeError):
            pass
        for child in widget.winfo_children():
            paint_footer(child)

    def bind_hover(button) -> None:
        try:
            button._partner_normal_bg = button.cget("bg")
            button._partner_normal_fg = button.cget("fg")
            button.configure(activebackground=hover_bg, activeforeground="#ffffff")
        except tk.TclError:
            return

        if getattr(button, "_partner_hover_bound", False):
            return

        def enter(_event=None):
            try:
                if str(button.cget("state")) != "disabled":
                    button.configure(bg=hover_bg, fg="#ffffff")
            except tk.TclError:
                pass

        def leave(_event=None):
            try:
                button.configure(
                    bg=button._partner_normal_bg,
                    fg=button._partner_normal_fg,
                )
            except tk.TclError:
                pass

        button.bind("<Enter>", enter, add="+")
        button.bind("<Leave>", leave, add="+")
        button._partner_hover_bound = True

    def walk(widget) -> None:
        for child in widget.winfo_children():
            tag = getattr(child, "_tag", None)
            if tag == "footer":
                paint_footer(child)
                continue
            if child.winfo_class() == "Button" and tag not in {
                    "header", "header_label", "brand", "logo", "footer"}:
                bind_hover(child)
            walk(child)

    walk(window)
