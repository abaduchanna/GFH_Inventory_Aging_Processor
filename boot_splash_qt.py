"""Branded boot splash with an animated loading bar.

Replaces the old PyInstaller text splash ("Starting..." / "Loading
interface..."): the specs no longer configure any splash text, and this
module takes over the moment Python is up - it closes the PyInstaller
splash and re-shows the SAME branded splash.png with a brand-colored
indeterminate progress bar, so the swap is seamless and the boot shows a
loading bar instead of text.

Usage in main() / __main__:

    from boot_splash_qt import show_boot_splash
    _boot = show_boot_splash("#e83030")   # this brand's bar color
    ... license check / UI build ...
    _boot.finish(window)                  # once the main window is up

Developed by www.3SVerse.com
"""
import os
import sys

from PySide6.QtGui import QPixmap
from PySide6.QtWidgets import QApplication, QProgressBar, QSplashScreen


def _splash_image_path():
    """splash.png next to this module (dev) or in the PyInstaller bundle."""
    cands = []
    if getattr(sys, "frozen", False):
        base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(sys.executable)))
        cands.append(os.path.join(base, "splash.png"))
    cands.append(os.path.join(os.path.dirname(os.path.abspath(__file__)), "splash.png"))
    for c in cands:
        if os.path.exists(c):
            return c
    return None


class BootSplash(QSplashScreen):
    """splash.png + brand-colored indeterminate loading bar."""

    def __init__(self, bar_color):
        path = _splash_image_path()
        pix = QPixmap(path) if path else QPixmap(640, 400)
        if pix.isNull():
            pix = QPixmap(640, 400)
        super().__init__(pix)
        # Deliberately NOT WindowStaysOnTopHint: any early dialog (license
        # activation, missing-package warning) must be able to sit above
        # the splash instead of being covered by it.
        bar_h = 10
        self._bar = QProgressBar(self)
        self._bar.setRange(0, 0)          # indeterminate "loading" shuttle
        self._bar.setTextVisible(False)
        self._bar.setGeometry(70, self.height() - 56, self.width() - 140, bar_h)
        r = max(2, bar_h // 2)
        self._bar.setStyleSheet(
            "QProgressBar { border: none; border-radius: %dpx;"
            " background: rgba(255, 255, 255, 0.14); }"
            "QProgressBar::chunk { border-radius: %dpx; background: %s; }"
            % (r, r, bar_color))

    def pulse(self):
        """Let queued paint events through so the shuttle keeps moving
        during long blocking init phases (license check, UI build)."""
        app = QApplication.instance()
        if app is not None:
            app.processEvents()


def show_boot_splash(bar_color):
    """Show the branded bar splash. Closes the static PyInstaller splash
    first so the two never stack (same image = seamless handover)."""
    try:
        import pyi_splash  # type: ignore
        pyi_splash.close()
    except Exception:
        pass
    splash = BootSplash(bar_color)
    splash.show()
    splash.pulse()
    return splash
