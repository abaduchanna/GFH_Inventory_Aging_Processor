"""
Theme Manager (Qt) - Standardized for GFH/VidaPay Ecosystem
Developed by www.3SVerse.com | Copyright (c) {year} | All rights reserved.

PySide6 port of theme_manager.py. Same color tokens, same on-disk config
location/format (so a user's saved light/dark choice survives migrating an
app from the Tkinter build to the Qt build), same app_name-scoped config
dir. Styling is applied via one QSS stylesheet instead of walking the
widget tree - Qt's cascade handles child widgets automatically, so there
is no "_walk" step and no risk of a widget being missed/stuck on the old
theme.
"""
import json
import os
from datetime import datetime


_THEME_MANAGER_QT_VERSION = "1.0.0"


class ThemeManagerQt:
    """Manages light/dark themes for a PySide6 app via QSS."""

    BRAND_NAVY = "#090d26"
    BRAND_RED = "#e83030"
    BRAND_WHITE = "#ffffff"

    THEMES = {
        "light": {
            "bg": "#f6f7fb",
            "panel": "#ffffff",
            "panel_alt": "#eef0f6",
            "text": "#16213a",
            "text_dim": "#5b6478",
            "input": "#ffffff",
            "border": "#d5d9e5",
            "navy": "#090d26",
            "red": "#e83030",
            "log_bg": "#ffffff",
            "log_fg": "#16213a",
        },
        "dark": {
            "bg": "#0b1020",
            "panel": "#141b38",
            "panel_alt": "#1c2447",
            "text": "#e8ecf7",
            "text_dim": "#9aa4c0",
            "input": "#1c2447",
            "border": "#2b3561",
            "navy": "#090d26",
            "red": "#e83030",
            "log_bg": "#05070f",
            "log_fg": "#cbd5e1",
        },
    }

    def __init__(self, default="dark", app_name="GFH"):
        self.app_name = app_name
        # Same config path as the Tkinter ThemeManager - a saved choice
        # carries over whichever build (Tk or Qt) the user has installed.
        self.CONFIG_DIR = os.path.join(os.path.expanduser("~"), ".config", "gfh-telecom", app_name)
        self.CONFIG_FILE = os.path.join(self.CONFIG_DIR, "theme.json")
        self.current_theme = self._load_theme() or default
        os.makedirs(self.CONFIG_DIR, exist_ok=True)

    def _load_theme(self):
        if os.path.exists(self.CONFIG_FILE):
            try:
                with open(self.CONFIG_FILE, "r") as f:
                    return json.load(f).get("theme", "dark")
            except Exception:
                pass
        return None

    def save_theme(self, theme_name):
        self.current_theme = theme_name
        try:
            with open(self.CONFIG_FILE, "w") as f:
                json.dump({"theme": theme_name}, f)
        except Exception:
            pass

    def get_colors(self):
        return self.THEMES.get(self.current_theme, self.THEMES["dark"]).copy()

    def toggle(self):
        self.current_theme = "light" if self.current_theme == "dark" else "dark"
        self.save_theme(self.current_theme)
        return self.current_theme

    def build_qss(self):
        """One stylesheet for the whole app. Objects tagged objectName
        'gfhHeader'/'gfhFooter' (see header_manager_qt.py) are explicitly
        excluded so they always stay navy regardless of theme - same rule
        as the Tkinter _PROTECTED_TAGS set."""
        c = self.get_colors()
        return f"""
            QWidget {{
                background: {c['bg']};
                color: {c['text']};
                font-family: "Segoe UI";
                font-size: 13px;
            }}
            QWidget#gfhHeader, QWidget#gfhHeader QWidget,
            QWidget#gfhFooter, QWidget#gfhFooter QWidget {{
                background: {self.BRAND_NAVY};
                color: {self.BRAND_WHITE};
            }}
            QMainWindow, QDialog {{
                background: {c['bg']};
            }}
            QFrame#gfhPanel {{
                background: {c['panel']};
                border: 1px solid {c['border']};
                border-radius: 6px;
            }}
            QLabel {{
                background: transparent;
            }}
            QPushButton {{
                background: {c['panel_alt']};
                color: {c['text']};
                border: 1px solid {c['border']};
                border-radius: 4px;
                padding: 6px 14px;
            }}
            QPushButton:hover {{
                background: {self.BRAND_RED};
                color: {self.BRAND_WHITE};
                border-color: {self.BRAND_RED};
            }}
            QPushButton:disabled {{
                color: {c['text_dim']};
                background: {c['panel_alt']};
            }}
            QLineEdit, QTextEdit, QPlainTextEdit, QSpinBox, QComboBox {{
                background: {c['input']};
                color: {c['text']};
                border: 1px solid {c['border']};
                border-radius: 4px;
                padding: 4px 6px;
                selection-background-color: {c['red']};
                selection-color: {self.BRAND_WHITE};
            }}
            QComboBox QAbstractItemView {{
                background: {c['panel']};
                color: {c['text']};
                selection-background-color: {c['red']};
                selection-color: {self.BRAND_WHITE};
            }}
            QCheckBox, QRadioButton {{
                background: transparent;
                spacing: 6px;
            }}
            QTabWidget::pane {{
                border: 1px solid {c['border']};
                background: {c['bg']};
            }}
            QTabBar::tab {{
                background: {c['panel_alt']};
                color: {c['text']};
                padding: 8px 16px;
                border: 1px solid {c['border']};
                border-bottom: none;
            }}
            QTabBar::tab:selected {{
                background: {c['panel']};
                color: {c['text']};
            }}
            QTableWidget, QTreeWidget, QListWidget {{
                background: {c['panel']};
                color: {c['text']};
                gridline-color: {c['border']};
                border: 1px solid {c['border']};
            }}
            QHeaderView::section {{
                background: {self.BRAND_NAVY};
                color: {self.BRAND_WHITE};
                padding: 4px 8px;
                border: none;
                font-weight: bold;
            }}
            QScrollBar:vertical {{
                background: {c['bg']};
                width: 12px;
            }}
            QScrollBar::handle:vertical {{
                background: {c['panel_alt']};
                border-radius: 5px;
                min-height: 24px;
            }}
            QScrollBar::handle:vertical:hover {{
                background: {c['red']};
            }}
            QProgressBar {{
                background: {c['panel_alt']};
                border: none;
                border-radius: 4px;
                text-align: center;
            }}
            QProgressBar::chunk {{
                background: {c['red']};
                border-radius: 4px;
            }}
            QGroupBox {{
                border: 1px solid {c['border']};
                border-radius: 6px;
                margin-top: 10px;
                font-weight: bold;
            }}
            QGroupBox::title {{
                subcontrol-origin: margin;
                left: 10px;
                padding: 0 4px;
            }}
        """

    def apply_theme_to_app(self, app):
        """Apply the current theme's stylesheet app-wide (QApplication)."""
        app.setStyleSheet(self.build_qss())
        try:
            from native_titlebar_qt import sync_titlebar_qt
            for w in app.topLevelWidgets():
                sync_titlebar_qt(w, dark=(self.current_theme == "dark"))
        except Exception:
            pass

    @staticmethod
    def get_copyright_year():
        return datetime.now().year

    @staticmethod
    def get_copyright_text():
        return f"Developed by www.3SVerse.com | Copyright (c) {datetime.now().year} | All rights reserved."


def apply_theme_to_app(app, theme_manager=None):
    """Convenience function (mirrors theme_manager.apply_theme_to_window)."""
    if theme_manager is None:
        theme_manager = ThemeManagerQt()
    theme_manager.apply_theme_to_app(app)
    return theme_manager


def get_copyright_year():
    return ThemeManagerQt.get_copyright_year()
