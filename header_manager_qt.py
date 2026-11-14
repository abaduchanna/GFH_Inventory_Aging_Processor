"""
Fixed Header Manager (Qt) - logo left + red divider + centered title.
Header always stays navy - never changes on theme toggle.
Developed by www.3SVerse.com

PySide6 port of header_manager.py / FixedHeaderManager. Same layout and
colors (logo left, red vertical divider, centered title that spans the
full header width, theme-toggle sun/moon glyph on the right, pinned
navy copyright footer) as every GFH/VDS/VidaPay app.
"""
import math
import os

from PySide6.QtCore import Qt, QRectF, QPointF
from PySide6.QtGui import QPainter, QPixmap, QColor, QPen, QFont, QPolygonF, QFontMetrics
from PySide6.QtWidgets import (
    QWidget, QLabel, QFrame, QHBoxLayout, QVBoxLayout, QSizePolicy
)

_HEADER_MANAGER_QT_VERSION = "1.1.0"

# Title headline size chosen in the VidaPay_Rebate_Filing UI-polish round.
# The label auto-shrinks (never below TITLE_MIN_PT) when the text does not
# fit, so long product names (e.g. "VDS Desk Solutions - Legacy Excel
# Converter" in an 1100px window) clip-proof while short titles keep the
# full 32pt headline look.
TITLE_BASE_PT = 32
TITLE_MIN_PT = 15

BRAND_NAVY = "#090d26"
BRAND_RED = "#e83030"
BRAND_WHITE = "#ffffff"


class ThemeToggleGlyph(QWidget):
    """Vector sun (dark theme active -> click for light) / crescent moon
    (light theme active -> click for dark), same glyph as
    theme_manager.draw_theme_glyph but painted with QPainter instead of a
    Tk canvas."""

    def __init__(self, theme_manager, on_click, parent=None, size=36):
        super().__init__(parent)
        self.theme_manager = theme_manager
        self.on_click = on_click
        self.setFixedSize(size, size)
        self.setCursor(Qt.PointingHandCursor)
        self.setObjectName("gfhHeader")

    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton and self.on_click:
            self.on_click()
        self.update()

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing)
        cx, cy = self.width() / 2.0, self.height() / 2.0
        color = QColor(BRAND_WHITE)
        is_dark = getattr(self.theme_manager, "current_theme", "dark") == "dark"

        if not is_dark:
            # Crescent moon (light theme active).
            R = 8.0
            d = 5.5
            half = math.acos(d / (2 * R))
            pts = []
            steps = 36
            sweep = 2 * math.pi - 2 * half
            for i in range(steps + 1):
                a = -half - sweep * i / steps
                pts.append(QPointF(cx + R * math.cos(a), cy + R * math.sin(a)))
            a_end = math.atan2(R * math.sin(half), R * math.cos(half) - d)
            a_start = math.atan2(R * math.sin(-half), R * math.cos(-half) - d)
            for i in range(steps + 1):
                a = a_end + (math.pi - a_end) * i / steps
                pts.append(QPointF(cx + d + R * math.cos(a), cy + R * math.sin(a)))
            for i in range(steps + 1):
                a = -math.pi + (a_start + math.pi) * i / steps
                pts.append(QPointF(cx + d + R * math.cos(a), cy + R * math.sin(a)))
            painter.setBrush(color)
            painter.setPen(Qt.NoPen)
            painter.drawPolygon(QPolygonF(pts))
        else:
            # Sun: ring + 8 rays.
            r = 6.0
            pen = QPen(color, 2)
            painter.setPen(pen)
            painter.setBrush(Qt.NoBrush)
            painter.drawEllipse(QRectF(cx - r, cy - r, r * 2, r * 2))
            for i in range(8):
                a = math.radians(i * 45)
                painter.drawLine(
                    int(cx + math.cos(a) * (r + 2.5)), int(cy + math.sin(a) * (r + 2.5)),
                    int(cx + math.cos(a) * (r + 6.5)), int(cy + math.sin(a) * (r + 6.5)),
                )
        painter.end()


class _TitleLabel(QLabel):
    """Centered headline title that keeps its font-size in its own
    stylesheet (the app-wide QSS cascade `QWidget { font-size: 13px; }`
    silently wins over setFont() - see the title block below) and
    auto-shrinks on resize so the text never clips its label."""

    def __init__(self, text):
        super().__init__(text)
        # _pt starts at 0 so the first _apply_fit() ALWAYS writes the
        # stylesheet - if the very first width already fit 32pt, skipping the
        # write would leave the app-wide QSS `QWidget { font-size: 13px; }`
        # cascade in charge of this label's font again.
        self._pt = 0
        self._apply_fit()

    def _apply_fit(self):
        pt = TITLE_BASE_PT
        avail = max(self.width() - 8, 60)
        while pt > TITLE_MIN_PT:
            fm = QFontMetrics(QFont("Segoe UI", pt, QFont.Bold))
            if fm.horizontalAdvance(self.text()) <= avail:
                break
            pt -= 1
        if pt != self._pt:
            self._pt = pt
            self.setStyleSheet(self._stylesheet(pt))

    def _stylesheet(self, pt):
        return (
            f"background: {BRAND_NAVY}; color: {BRAND_WHITE}; "
            f"font-size: {pt}pt; font-weight: bold; font-family: 'Segoe UI';"
        )

    def resizeEvent(self, event):
        self._apply_fit()
        super().resizeEvent(event)


class FixedHeaderManagerQt:
    """Manages a navy header bar with centered title, left logo, red
    divider and a theme-toggle glyph. Attach to any QMainWindow/QWidget
    whose central layout you control."""

    BRAND_NAVY = BRAND_NAVY
    BRAND_RED = BRAND_RED

    def __init__(self, parent_layout, title="App", height=90):
        self.title = title
        self.height = height
        self.theme_manager = None
        self.theme_toggle_glyph = None
        self.copyright_label = None
        self.footer_frame = None

        self.header_frame = QFrame()
        self.header_frame.setObjectName("gfhHeader")
        self.header_frame.setFixedHeight(height)
        self.header_frame.setStyleSheet(f"background: {self.BRAND_NAVY};")

        outer = QHBoxLayout(self.header_frame)
        outer.setContentsMargins(18, 9, 18, 9)
        outer.setSpacing(0)

        # LEFT: logo + divider, grouped in their own frame so its total width
        # can be balanced against the right-side frame's width (see
        # _rebalance_flanks) - otherwise the title's AlignCenter only centers
        # it within whatever space is left between two DIFFERENTLY-sized
        # flanks (a wide logo vs. a narrow theme-toggle glyph), so the title
        # visibly sits off-center from the window's actual midpoint.
        self.left_frame = QFrame()
        self.left_frame.setObjectName("gfhHeader")
        self.left_frame.setStyleSheet(f"background: {self.BRAND_NAVY};")
        left_layout = QHBoxLayout(self.left_frame)
        left_layout.setContentsMargins(0, 0, 0, 0)
        left_layout.setSpacing(0)

        self.logo_label = QLabel()
        self.logo_label.setObjectName("gfhHeader")
        self.logo_label.setStyleSheet(f"background: {self.BRAND_NAVY}; color: {self.BRAND_RED};")
        self.logo_label.setFont(QFont("Segoe UI", 9, QFont.Bold))
        left_layout.addWidget(self.logo_label, 0, Qt.AlignVCenter)

        # Red vertical divider
        self.divider_frame = QFrame()
        self.divider_frame.setObjectName("gfhHeader")
        self.divider_frame.setFixedWidth(3)
        self.divider_frame.setStyleSheet(f"background: {self.BRAND_RED}; margin: 12px 0;")
        left_layout.addSpacing(14)
        left_layout.addWidget(self.divider_frame)
        # Any extra width _rebalance_flanks later gives this frame goes HERE
        # (between the divider and the title), so logo+divider stay pinned
        # to the header's left edge instead of drifting toward the middle.
        left_layout.addStretch(1)
        outer.addWidget(self.left_frame, 0, Qt.AlignVCenter)

        # CENTER: title - stretches to fill remaining space, text centered.
        # font-size must be set in THIS widget's own stylesheet, not just via
        # setFont(): theme_manager_qt.py's global QSS has `QWidget { font-size:
        # 13px; }`, and Qt's stylesheet cascade always wins over setFont() for
        # any property the app-wide stylesheet also sets on this widget - that
        # 13px rule was silently overriding every setFont() size tried here.
        # _TitleLabel additionally shrink-fits the point size to the label's
        # width on every resize, so the headline never clips on narrower
        # windows (base 32pt, floor TITLE_MIN_PT).
        self.title_label = _TitleLabel(title)
        self.title_label.setObjectName("gfhHeader")
        self.title_label.setAlignment(Qt.AlignCenter)
        outer.addWidget(self.title_label, 1)

        # RIGHT: theme toggle placeholder (added via add_theme_toggle)
        self.right_frame = QFrame()
        self.right_frame.setObjectName("gfhHeader")
        self.right_frame.setStyleSheet(f"background: {self.BRAND_NAVY};")
        self.right_layout = QHBoxLayout(self.right_frame)
        self.right_layout.setContentsMargins(0, 0, 0, 0)
        outer.addWidget(self.right_frame, 0, Qt.AlignVCenter)

        parent_layout.addWidget(self.header_frame)

    def set_logo(self, logo_path=None, text="Logo", max_w=190, max_h=72):
        if logo_path and os.path.exists(logo_path):
            pix = QPixmap(logo_path)
            if not pix.isNull():
                pix = pix.scaled(max_w, max_h, Qt.KeepAspectRatio, Qt.SmoothTransformation)
                self.logo_label.setPixmap(pix)
                self.logo_label.setText("")
                return
        self.logo_label.setText(text)

    def add_theme_toggle(self, theme_manager, callback=None):
        self.theme_manager = theme_manager

        def toggle_and_callback():
            theme_manager.toggle()
            if self.theme_toggle_glyph:
                self.theme_toggle_glyph.update()
            self.update_button_text()
            if callback:
                callback()

        # Stretch BEFORE the glyph so it stays pinned to the header's right
        # edge if _rebalance_flanks() below widens this frame - mirrors the
        # stretch-AFTER-content on left_frame, which pins logo+divider left.
        self.right_layout.addStretch(1)
        self.theme_toggle_glyph = ThemeToggleGlyph(theme_manager, toggle_and_callback, self.right_frame)
        self.right_layout.addWidget(self.theme_toggle_glyph, 0, Qt.AlignCenter)
        self._rebalance_flanks()

    def _rebalance_flanks(self):
        """Force left_frame (logo+divider) and right_frame (theme toggle) to
        the SAME width, whichever is naturally wider. Without this, the
        title's AlignCenter only centers it within whatever space is left
        between two differently-sized flanks (e.g. a wide logo vs. a narrow
        toggle glyph), so the title visibly sits off the window's true
        midpoint rather than looking centered."""
        self.left_frame.setMinimumWidth(0)
        self.left_frame.setMaximumWidth(16777215)
        self.right_frame.setMinimumWidth(0)
        self.right_frame.setMaximumWidth(16777215)
        w = max(self.left_frame.sizeHint().width(), self.right_frame.sizeHint().width())
        self.left_frame.setFixedWidth(w)
        self.right_frame.setFixedWidth(w)

    def add_copyright(self, theme_manager, parent_layout):
        """Pinned navy footer bar with centered copyright text. Caller
        passes the outer (window-level) layout to append it to, same as
        the Tk version pinning to the window, not the header frame."""
        copyright_text = theme_manager.get_copyright_text()

        self.footer_frame = QFrame()
        self.footer_frame.setObjectName("gfhFooter")
        self.footer_frame.setFixedHeight(24)
        self.footer_frame.setStyleSheet(f"background: {self.BRAND_NAVY};")
        footer_layout = QHBoxLayout(self.footer_frame)
        footer_layout.setContentsMargins(0, 0, 0, 0)

        self.copyright_label = QLabel(copyright_text)
        self.copyright_label.setObjectName("gfhFooter")
        self.copyright_label.setAlignment(Qt.AlignCenter)
        self.copyright_label.setStyleSheet(f"background: {self.BRAND_NAVY}; color: #c7cbe0; font-size: 8pt;")
        footer_layout.addWidget(self.copyright_label)

        parent_layout.addWidget(self.footer_frame)

    def update_button_text(self):
        if self.copyright_label and self.theme_manager:
            self.copyright_label.setText(self.theme_manager.get_copyright_text())

    def update_for_theme(self, colors=None):
        """Called when theme changes - header never recolors, only the
        toggle glyph redraws and the copyright year text refreshes."""
        self.update_button_text()
