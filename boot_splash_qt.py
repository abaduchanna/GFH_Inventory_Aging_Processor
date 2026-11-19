"""Branded boot splash with a LIVE animated loading bar.

v2 - "live loading" rewrite (owner: "bar live nahi chalti, frozen dikhti hai"):

The v1 splash used a Qt indeterminate QProgressBar shuttle. That shuttle
only advances while the Qt event loop runs, so during the blocking startup
phases (license verify network call, MainWindow construction) the bar sat
frozen. This version paints the bar itself and keeps it moving:

- The bar is manually painted (trough + brand chunk) with a light band
  that sweeps across it on a wall-clock basis. Every paint pass advances
  the band, and the pump helper (pulse/set_progress) repaints continuously,
  so the bar visibly "chalta hai" through the whole boot.
- set_progress() SLIDES the fill smoothly from the old percent to the new
  one while pumping events, driven by the real startup phases:
  license_core._boot_status() forwards "Checking license..." /
  "Verifying license..." / "Loading <app>..." with percent bumps.
- license_core._bounded_call() pumps the splash while it waits on the
  bounded network call (boot_pulse), so the shimmer never stalls there.
- boot_close() lets license_core dismiss the splash right before any Tk
  activation dialog (same contract as _close_boot_splash()).

Usage in main() / __main__:

    from boot_splash_qt import show_boot_splash
    _boot = show_boot_splash("#f0541c")   # this brand's bar color
    ... license check / UI build ...      # bar animates via license_core pump
    _boot.finish(window)                  # once the main window is up

Developed by www.3SVerse.com
"""
import os
import sys
import time

from PySide6.QtCore import QRect, Qt
from PySide6.QtGui import (QColor, QFont, QLinearGradient, QPainter,
                           QPainterPath, QPixmap)
from PySide6.QtWidgets import QApplication, QSplashScreen

# Active splash (set by show_boot_splash, cleared by boot_close). The
# license_core hooks read this through boot_progress/boot_pulse/boot_close.
_SPLASH = None


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
    """splash.png + brand-colored LIVE loading bar + status text."""

    BAR_H = 10
    MARGIN_X = 70

    def __init__(self, bar_color):
        path = _splash_image_path()
        pix = QPixmap(path) if path else QPixmap(640, 400)
        if pix.isNull():
            pix = QPixmap(640, 400)
        super().__init__(pix)
        # Deliberately NOT WindowStaysOnTopHint: any early dialog (license
        # activation, missing-package warning) must be able to sit above
        # the splash instead of being covered by it.
        self._color = QColor(bar_color)
        self._pct = 0.0
        self._text = "Starting..."
        self._t0 = time.monotonic()
        self._radius = max(2, self.BAR_H // 2)

    # ------------------------------------------------------------------
    # geometry
    # ------------------------------------------------------------------
    def _trough(self):
        w = max(120, self.width() - 2 * self.MARGIN_X)
        return QRect(self.MARGIN_X, self.height() - 56, w, self.BAR_H)

    # ------------------------------------------------------------------
    # painting - the whole bar is hand-drawn so it never depends on the
    # widget style's internal shuttle timers (those freeze when the event
    # loop is busy; ours advances on wall-clock time instead).
    # ------------------------------------------------------------------
    def paintEvent(self, event):
        p = QPainter(self)
        try:
            pm = self.pixmap()
            if pm is not None and not pm.isNull():
                p.drawPixmap(self.rect(), pm)

            p.setRenderHint(QPainter.Antialiasing, True)
            p.setPen(Qt.NoPen)

            r = self._trough()
            p.setBrush(QColor(255, 255, 255, 38))
            p.drawRoundedRect(r, self._radius, self._radius)

            w = int(r.width() * (self._pct / 100.0))
            w = max(w, 14) if self._pct > 0 else 0
            if w > 0:
                w = min(w, r.width())
                chunk = QRect(r.x(), r.y(), w, r.height())
                path = QPainterPath()
                path.addRoundedRect(chunk, self._radius, self._radius)
                p.save()
                p.setClipPath(path)
                p.setBrush(self._color)
                p.drawRoundedRect(chunk, self._radius, self._radius)
                # Sweeping light band - advances on wall-clock time, so
                # every repaint makes it visibly move.
                cycle = 1.35
                t = (time.monotonic() - self._t0) % cycle / cycle
                band_w = max(30, chunk.width() // 3)
                bx = chunk.x() - band_w + t * (chunk.width() + 2 * band_w)
                grad = QLinearGradient(bx, 0, bx + band_w, 0)
                grad.setColorAt(0.0, QColor(255, 255, 255, 0))
                grad.setColorAt(0.5, QColor(255, 255, 255, 95))
                grad.setColorAt(1.0, QColor(255, 255, 255, 0))
                p.fillRect(QRect(int(bx), r.y(), int(band_w), r.height()), grad)
                p.restore()

            if self._text:
                f = QFont("Segoe UI", 10)
                p.setFont(f)
                tr = QRect(0, r.bottom() + 10, self.width(), 24)
                p.setPen(QColor(0, 0, 0, 150))
                p.drawText(tr.translated(0, 1),
                           Qt.AlignHCenter | Qt.AlignVCenter, self._text)
                p.setPen(QColor(255, 255, 255, 240))
                p.drawText(tr, Qt.AlignHCenter | Qt.AlignVCenter, self._text)
        finally:
            p.end()

    # ------------------------------------------------------------------
    # live-update helpers (all called from the MAIN thread)
    # ------------------------------------------------------------------
    def _pump(self):
        app = QApplication.instance()
        if app is not None:
            app.processEvents()

    def pulse(self):
        """One paint pass + event pump: the shimmer band moves each call.
        license_core._bounded_call() calls this every ~30 ms while it waits
        on the bounded network call, so the bar keeps dancing."""
        self.update()
        self._pump()

    def set_progress(self, pct, text=None):
        """Slide the fill smoothly to `pct` (0-100) while pumping events.
        Called from the real startup phases - the bar visibly travels
        instead of jumping or freezing."""
        try:
            pct = max(0, min(100, int(pct)))
        except Exception:
            return
        if text is not None:
            self._text = str(text)
        start = self._pct
        if pct <= start:
            self._pct = float(pct)
            self.update()
            self._pump()
            return
        dur = min(0.5, 0.05 + 0.008 * (pct - start))
        t0 = time.monotonic()
        while True:
            t = (time.monotonic() - t0) / dur
            if t >= 1.0:
                break
            e = t * t * (3.0 - 2.0 * t)  # smoothstep ease
            self._pct = start + (pct - start) * e
            self.update()
            self._pump()
            time.sleep(0.016)
        self._pct = float(pct)
        self.update()
        self._pump()

    def finish(self, window):
        self.set_progress(100, "Ready")
        super().finish(window)


# ----------------------------------------------------------------------
# module-level hooks used by license_core (safe no-ops without a splash)
# ----------------------------------------------------------------------
def boot_progress(pct, text=None):
    """Set bar percent and/or status text on the active splash."""
    s = _SPLASH
    if s is None:
        return
    try:
        if pct is None:
            if text is not None:
                s._text = str(text)
            s.update()
            s._pump()
        else:
            s.set_progress(int(pct), text)
    except Exception:
        pass


def boot_pulse():
    """Keep the splash animating while the main thread is blocked."""
    s = _SPLASH
    if s is None:
        return
    try:
        s.pulse()
    except Exception:
        pass


def boot_close():
    """Dismiss the active splash (license_core does this right before any
    Tk activation dialog so the dialog is never covered)."""
    global _SPLASH
    s = _SPLASH
    _SPLASH = None
    if s is None:
        return
    try:
        s.close()
    except Exception:
        pass


def show_boot_splash(bar_color):
    """Show the branded LIVE bar splash. Closes the static PyInstaller
    splash first so the two never stack (same image = seamless handover)."""
    global _SPLASH
    try:
        import pyi_splash  # type: ignore
        pyi_splash.close()
    except Exception:
        pass
    splash = BootSplash(bar_color)
    _SPLASH = splash
    splash.show()
    splash.set_progress(6, "Starting...")
    return splash
