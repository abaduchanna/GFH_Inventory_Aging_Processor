import datetime as _doc_dt
from datetime import date
_DOC_YEAR = _doc_dt.date.today().year

f"""
GFH Inventory Aging Processor
==============================
Pick an Inventory Aging Excel export, then:

  PROCESS & EMAIL (local):
    1. Remove SIM/SIM-card rows.
    2. Remove rows with blank Serial 1 (IMEI) or blocked IMEIs.
    3. Keep only devices aged 20+ days.
    4. Merge Arizona D1/D2 into a single Arizona district.
    5. Save a styled .xlsx per district (frozen header, auto-filter, alternating fills).
    6. Email each district file via Outlook Desktop with inline GFH logo signature.

  UPLOAD TO GOOGLE SHEETS:
    1-4 same cleaning as above, but keep devices aged 14+ days.
    5. Upload to Google Sheets ("GFH Inventory Aging") with ONE TAB PER DISTRICT.
    6. Build an "Executive Dashboard" tab highlighting key KPIs and risks.
    7. Every run reconciles the Google Sheet to exactly today's data.

EXCLUSIONS PANEL:
    Stores / IMEIs listed there are removed from the Google Sheets upload
    (district tabs + Executive Dashboard) and from the district files that
    get emailed via Outlook. "Always-blocked IMEIs" are removed from BOTH
    pipelines at cleaning time. Everything is editable in the app and saved
    to gfh_aging_exclusions.json next to the app.

SECOND TAB - "WhatsApp - Aged Devices" (independent of the above):
    Pick a raw aging export, tick which districts to send and type each
    one's WhatsApp Desktop group name by hand, set an Aged (days) or more
    cutoff, optionally exclude stores (separate list from the panel above -
    this tab only). Save Mapping / Send to WhatsApp Desktop then renders one
    GFH-branded table screenshot per ticked district (DISTRICT, STORE,
    PRODUCT, IMEI, AGE IN COMPANY, PO RECEIVING DATE) and pastes + sends it
    into that WhatsApp Desktop group. Everything persists to
    gfh_aging_whatsapp_config.json next to the app. Requires pyautogui.

Install once:
    pip install pandas openpyxl gspread gspread-formatting pywin32 tkinterdnd2 pillow pyautogui

Developed by www.3SVerse.com | Copyright © {date.today().year} | All rights reserved.
"""

import os, sys, re, io, json, base64, tempfile, threading, traceback, queue, argparse, webbrowser, time
from pathlib import Path
from datetime import datetime

try:
    import tkinter as tk
    from theme_manager import ThemeManager, apply_theme_to_window, get_copyright_year
    from tkinter import ttk, filedialog, messagebox, scrolledtext
except ImportError:
    print("tkinter not available."); sys.exit(1)

from header_manager import FixedHeaderManager
from logo_handler import LogoHandler

try:
    from tkinterdnd2 import TkinterDnD, DND_FILES
    HAS_DND = True
except ImportError:
    HAS_DND = False

_MISSING = []
_HAS_GSPREAD = True
_HAS_GSPREAD_FMT = True
_HAS_OUTLOOK = True
_HAS_PIL = True

try:
    import pandas as pd
except ImportError:
    _MISSING.append("pandas")
try:
    import openpyxl
except ImportError:
    _MISSING.append("openpyxl")
try:
    import gspread
    from gspread.exceptions import WorksheetNotFound, SpreadsheetNotFound, APIError
except ImportError:
    _HAS_GSPREAD = False
try:
    import gspread_formatting
    from gspread_formatting import *
except ImportError:
    _HAS_GSPREAD_FMT = False
try:
    import win32com.client as _win32, pythoncom as _pythoncom
except ImportError:
    _HAS_OUTLOOK = False
try:
    from PIL import Image as _PI, ImageTk as _PIT
except ImportError:
    _HAS_PIL = False
try:
    # Presence check only - no import. The actual WhatsApp Desktop automation
    # (ported from GFH_Audit_Automation.py / GFH_Inventory_Audit.py, both
    # already shipping and proven) is pure ctypes + win32clipboard (already
    # a required dependency here for Outlook) and needs no extra package at
    # all; pyautogui is kept only for the same fail-fast install check the
    # proven apps use, imported lazily inside the send function itself.
    import importlib.util as _ilu
    _HAS_WA_AUTOMATION_DEPS = _ilu.find_spec("pyautogui") is not None
except Exception:
    _HAS_WA_AUTOMATION_DEPS = False

# Import openpyxl Border/Side with aliases to avoid conflict with gspread_formatting's Border
from openpyxl.styles import Font, PatternFill, Alignment, Border as XlBorder, Side as XlSide
from openpyxl.utils import get_column_letter
import queue
import threading
import re
import base64
import tempfile
import json

if not _HAS_GSPREAD or not _HAS_GSPREAD_FMT:
    _MISSING.append("gspread + gspread-formatting")
if not _HAS_OUTLOOK:
    _MISSING.append("pywin32")
if not _HAS_PIL:
    _MISSING.append("pillow")

# Show error only for REQUIRED packages (pandas, openpyxl)
_required_missing = [p for p in ("pandas", "openpyxl") if p in _MISSING]
if _required_missing:
    _r = tk.Tk(); _r.withdraw()
    messagebox.showerror(
        "GFH Inventory Aging Processor - missing packages",
        "Required packages not installed:\n\n"
        + "\n".join(f"  - {p}" for p in _required_missing)
        + "\n\nInstall and reopen:\n\npip install " + " ".join(_required_missing),
    )
    sys.exit(1)

# Warn about optional packages
_optional_missing = [p for p in _MISSING if p not in ("pandas", "openpyxl")]
if _optional_missing:
    _r = tk.Tk(); _r.withdraw()
    messagebox.showwarning(
        "GFH Inventory Aging Processor",
        "Optional packages not installed (some features disabled):\n\n"
        + "\n".join(f"  - {p}" for p in _optional_missing)
        + "\n\nInstall with: pip install pandas openpyxl gspread gspread-formatting pywin32 tkinterdnd2 pillow",
    )

# ==========================================================
# CONFIGURATION
# ==========================================================
APP_TITLE = "GFH Inventory Aging Processor"
ICON_ICO_NAME = "GFH_icon.ico"      # used for taskbar + titlebar (Windows .ico)
LOGO_PNG_NAME = "GFH_Telecom_Logo.png"        # used in the header (resized at runtime via PIL)

NAVY       = "#090d26"
EMBEDDED_LOGO_B64 = open(os.path.join(os.path.dirname(os.path.abspath(__file__)), "assets", "embedded_logo_b64.txt"), "r").read().strip() if not getattr(sys, "frozen", False) else open(os.path.join(getattr(sys, "_MEIPASS", "."), "assets", "embedded_logo_b64.txt"), "r").read().strip()
EMBEDDED_ICON_B64 = open(os.path.join(os.path.dirname(os.path.abspath(__file__)), "assets", "embedded_icon_b64.txt"), "r").read().strip() if not getattr(sys, "frozen", False) else open(os.path.join(getattr(sys, "_MEIPASS", "."), "assets", "embedded_icon_b64.txt"), "r").read().strip()

NAVY_DARK  = "#05060f"
RED        = "#e83030"
RED_DARK   = "#c82020"
LIGHT      = "#f6f7fb"
WHITE      = "#ffffff"
COLOR_BG   = "#f6f7fb"
COLOR_TEXT  = "#1c1c1c"
COLOR_LOG_BG = "#0d0d1f"
COLOR_LOG_TEXT = "#d7e3f0"

COPYRIGHT_TEXT = f"Developed by www.3SVerse.com | Copyright © {date.today().year} | All rights reserved."

# ── Local Excel + Email config ──
MIN_AGE       = 20
KEEP_COLUMNS  = ["District","Store","Product Desc Full","Serial 1","Age in Company"]
# Seed for the Exclusions panel's "Always-blocked IMEIs" box on the very
# first run only. After that gfh_aging_exclusions.json is the single source
# of truth (editable in the Exclusions dialog). Blank-entry guards ("", " ")
# are unnecessary now that blank serials are removed separately.
BLOCKED_IMEIS = {"358975210745726", "350776860110726", "358975210799012",
                 "358975210797339", "354709280259373", "358975210792793"}

CONTACTS_CONFIG_FILE = "gfh_aging_contacts.json"
EXCLUSIONS_CONFIG_FILE = "gfh_aging_exclusions.json"

# Used only to seed gfh_aging_contacts.json the very first time it's created.
# Everything here is editable (and district rows are addable/removable) from
# the app's Settings dialog afterward — this is just the starting point.
_DEFAULT_SENDER = {
    "sender_name": "Abad Umair Channa",
    "sender_email": "sender@example.com",
    "sender_title": "Senior Inventory & Supply Chain Analyst",
    "sender_mobile": "346-385-9107",
    "cc_name": "",
    "cc_email": "",
}
_DEFAULT_DISTRICTS = {
    "Arizona":       {"name": "Mohammad Farhan Mohiuddin", "email": "arizona@example.com"},
    "Houston":       {"name": "Muhammad Hamza",            "email": "houston@example.com"},
    "Louisiana":     {"name": "Asif Khan",                 "email": "louisiana@example.com"},
    "Colorado West": {"name": "Raiyan Baig",               "email": "colorado-west@example.com"},
    "Colorado East": {"name": "Shehriyar Ali",             "email": "colorado-east@example.com"},
    "Tennessee":     {"name": "Ahmed Siraj",                "email": "tennessee@example.com"},
}


def _contacts_config_path() -> str:
    return os.path.join(get_app_dir(), CONTACTS_CONFIG_FILE)


def load_contacts_config() -> dict:
    """Load sender/CC/per-district contacts from the JSON config file.
    On first run (no config file yet) this is seeded with the existing
    defaults so nothing looks stripped. After that, the JSON file is the
    single source of truth — districts can be freely added, renamed, or
    removed from the Settings dialog and will persist here, not just the
    original fixed six."""
    path = _contacts_config_path()
    if not os.path.exists(path):
        cfg = dict(_DEFAULT_SENDER)
        cfg["districts"] = {d: dict(v) for d, v in _DEFAULT_DISTRICTS.items()}
        save_contacts_config(cfg)
        return cfg

    cfg = {
        "sender_name": "", "sender_email": "", "sender_title": "", "sender_mobile": "",
        "cc_name": "", "cc_email": "", "districts": {},
    }
    try:
        with open(path, "r", encoding="utf-8") as f:
            saved = json.load(f)
        for k in ("sender_name", "sender_email", "sender_title", "sender_mobile", "cc_name", "cc_email"):
            cfg[k] = saved.get(k, "")
        # Districts are whatever is actually saved — fully dynamic, not
        # restricted to any fixed list, so newly-added districts persist.
        saved_districts = saved.get("districts", {})
        for d, entry in saved_districts.items():
            cfg["districts"][d] = {
                "name": entry.get("name", ""),
                "email": entry.get("email", ""),
            }
    except Exception:
        pass
    return cfg


def save_contacts_config(cfg: dict) -> None:
    path = _contacts_config_path()
    with open(path, "w", encoding="utf-8") as f:
        json.dump(cfg, f, indent=2, ensure_ascii=False)


def contacts_config_to_legacy(cfg: dict):
    """Convert the config dict into the SENDER_EMAIL / CC_EMAIL / DISTRICT_EMAILS
    shapes the rest of the pipeline already expects."""
    sender_email = cfg.get("sender_email", "") or ""
    cc_name = cfg.get("cc_name", "")
    cc_email = cfg.get("cc_email", "")
    cc_full = f"{cc_name} <{cc_email}>" if cc_name and cc_email else (cc_email or "")
    district_emails = {}
    for d, entry in cfg.get("districts", {}).items():
        name = entry.get("name", "")
        email = entry.get("email", "")
        if not email:
            continue
        district_emails[d] = f"{name} <{email}>" if name else email
    return sender_email, cc_full, district_emails

# ── Google Sheets config ──
GOOGLE_SHEET_NAME   = "GFH Inventory Aging"
DASHBOARD_TAB_NAME  = "Executive Dashboard"
CREDENTIALS_FILE    = "credentials.json"
AUTHORIZED_USER_FILE = "authorized_user.json"
AGE_THRESHOLD_DAYS  = 14

HEADER_BG = {"red": 0x16 / 255, "green": 0x16 / 255, "blue": 0x32 / 255}
HEADER_FG = {"red": 1, "green": 1, "blue": 1}

COLUMN_CANDIDATES = {
    "District": ["District"],
    "Store": ["Store"],
    "Product Description": ["Product Desc Full", "Product Description", "Product Desc Short", "Product Desc"],
    "Serial 1": ["Serial 1", "Serial1", "Serial"],
    "Age in Company": ["Age in Company", "Age In Company"],
    "PO Date": ["PO Received Date", "PO Date", "PO Receipt Date"],
    "Retail Price": ["Retail Price", "Retail"],
}

# ── Embedded images (base64) ──
GFH_LOGO_B64 = open(os.path.join(os.path.dirname(os.path.abspath(__file__)), "assets", "gfh_logo_b64.txt"), "r").read().strip() if not getattr(sys, "frozen", False) else open(os.path.join(getattr(sys, "_MEIPASS", "."), "assets", "gfh_logo_b64.txt"), "r").read().strip()
ICON_ICO_B64 = open(os.path.join(os.path.dirname(os.path.abspath(__file__)), "assets", "icon_ico_b64.txt"), "r").read().strip() if not getattr(sys, "frozen", False) else open(os.path.join(getattr(sys, "_MEIPASS", "."), "assets", "icon_ico_b64.txt"), "r").read().strip()
HEADER_LOGO_B64 = open(os.path.join(os.path.dirname(os.path.abspath(__file__)), "assets", "header_logo_b64.txt"), "r").read().strip() if not getattr(sys, "frozen", False) else open(os.path.join(getattr(sys, "_MEIPASS", "."), "assets", "header_logo_b64.txt"), "r").read().strip()


def get_script_dir() -> str:
    return os.path.dirname(os.path.abspath(__file__))


def get_app_dir() -> str:
    """Directory where user-facing files live (credentials.json, authorized_user.json,
    gfh_aging_contacts.json).

    When frozen into a PyInstaller onefile EXE, ``__file__`` points at the temporary
    ``_MEIPASS`` extraction folder, which is wiped on each run and is never where the
    user placed their files. In that case resolve relative to the EXE itself so files
    sitting next to the app are actually found.
    """
    if getattr(sys, "frozen", False) and getattr(sys, "executable", None):
        return os.path.dirname(os.path.abspath(sys.executable))
    return os.path.dirname(os.path.abspath(__file__))


# ==========================================================
# EXCLUSIONS (stores / IMEIs hidden from outputs)
# ==========================================================

def normalize_imei(value):
    """Digits-only form of an IMEI entry (spaces, dashes, dots ignored).
    Used both to clean user input from the Exclusions panel and to compare
    Serial 1 values against it."""
    return re.sub(r"\D", "", "" if value is None else str(value))


def store_matches_pattern(store_value, pattern):
    """Case-insensitive store match: equal OR either side contains the other,
    so typing 'houston #12' matches 'Houston #12' exactly and typing
    'houston' matches every store whose name contains it (and pasting a
    longer full name still matches a shorter store cell)."""
    s = ("" if store_value is None else str(store_value)).strip().casefold()
    p = ("" if pattern is None else str(pattern)).strip().casefold()
    if not s or not p:
        return False
    return p in s or s in p


def load_exclusions() -> dict:
    """Load the exclusions config: stores / IMEIs hidden from the Google
    Sheets upload + emailed district files, and always-blocked IMEIs removed
    from BOTH pipelines at cleaning time. On first run the file is seeded,
    migrating the old hardcoded BLOCKED_IMEIS list so behaviour stays
    identical until the user edits it in the Exclusions panel."""
    path = os.path.join(get_app_dir(), EXCLUSIONS_CONFIG_FILE)
    exc = {"stores": [], "imeis": [], "blocked_imeis": sorted(BLOCKED_IMEIS)}
    if os.path.exists(path):
        try:
            with open(path, "r", encoding="utf-8") as f:
                saved = json.load(f)
            for k in ("stores", "imeis", "blocked_imeis"):
                vals = saved.get(k, [])
                if isinstance(vals, list):
                    exc[k] = [str(v).strip() for v in vals if str(v).strip()]
        except Exception:
            pass
    else:
        save_exclusions(exc)
    return exc


def save_exclusions(exc: dict) -> None:
    """Persist the exclusions config, trimming blanks, normalizing IMEI
    entries to digits and de-duplicating (case-insensitive for stores)."""
    path = os.path.join(get_app_dir(), EXCLUSIONS_CONFIG_FILE)
    clean = {}
    for k in ("stores", "imeis", "blocked_imeis"):
        seen, out = set(), []
        for v in exc.get(k, []):
            v = ("" if v is None else str(v)).strip()
            if not v:
                continue
            if k != "stores":
                v = normalize_imei(v) or v
            key = v.casefold()
            if key in seen:
                continue
            seen.add(key)
            out.append(v)
        clean[k] = out
    with open(path, "w", encoding="utf-8") as f:
        json.dump(clean, f, indent=2, ensure_ascii=False)


def apply_output_exclusions(df, col_serial, col_store, log=print):
    """Remove excluded stores/IMEIs from an OUTBOUND dataset (the Google
    Sheets district tabs + Executive Dashboard, or the district .xlsx files
    that get emailed). IMEIs compare digits-only; stores match per
    store_matches_pattern. Logs exactly what was removed. Returns the
    filtered DataFrame (input untouched)."""
    exc = load_exclusions()
    store_pats = [s for s in exc.get("stores", []) if s.strip()]
    imei_set = {normalize_imei(x) for x in exc.get("imeis", [])}
    imei_set.discard("")
    if not store_pats and not imei_set:
        return df

    before = len(df)
    if store_pats:
        stores = df[col_store].fillna("").astype(str).str.strip()
        bad = {u for u in stores.unique()
               if any(store_matches_pattern(u, p) for p in store_pats)}
        if bad:
            df = df[~stores.isin(list(bad))]
            shown = ", ".join(sorted(bad)[:8]) + (" ..." if len(bad) > 8 else "")
            log(f"  Exclusions: removed {before - len(df)} row(s) from excluded store(s): {shown}")
        before = len(df)
    if imei_set:
        serials = df[col_serial].astype(str).str.strip().map(normalize_imei)
        df = df[~serials.isin(list(imei_set))]
        log(f"  Exclusions: removed {before - len(df)} row(s) from excluded IMEI(s).")
    return df


def log_exclusions_summary(log=print):
    """One-line Activity Log summary of what the Exclusions panel will do."""
    exc = load_exclusions()
    n_s = len([s for s in exc.get("stores", []) if s.strip()])
    n_i = len([x for x in exc.get("imeis", []) if normalize_imei(x)])
    n_b = len([x for x in exc.get("blocked_imeis", []) if normalize_imei(x)])
    log(f"Exclusions: {n_s} store(s) + {n_i} IMEI(s) hidden from Sheet/emails; "
        f"{n_b} always-blocked IMEI(s) removed at cleaning.")


# ==========================================================
# WHATSAPP — AGED DEVICES (second tab)
# ==========================================================
# Independent of the Email/Google Sheets pipeline above: lets the user pick
# an aging export, map each District to a WhatsApp Desktop group (typed in
# by hand — no fixed list), tick which districts to send, and push one
# GFH-branded table image per ticked district straight into that group's
# chat via WhatsApp Desktop (not WhatsApp Web).
WHATSAPP_CONFIG_FILE = "gfh_aging_whatsapp_config.json"

WA_SRC_COLUMNS = {
    "district": ["District"],
    "store": ["Store"],
    "product": ["Product Desc Short", "Product Desc Full"],
    "imei": ["Serial 1"],
    "age_in_company": ["Age in Company"],
    "po_date": ["PO Received Date"],
}
WA_TABLE_COLUMNS = [
    ("DISTRICT", "district", 150),
    ("STORE", "store", 230),
    ("PRODUCT", "product", 300),
    ("IMEI", "imei", 190),
    ("AGE IN COMPANY", "age_in_company", 150),
    ("PO RECEIVING DATE", "po_date", 170),
]

WA_DEFAULT_MESSAGE = (
    "Team, the attached devices have been sitting in inventory for {days}+ "
    "days. Please prioritize them for sale."
)



def _wa_norm_header(h):
    return re.sub(r"\s+", " ", str(h or "")).strip().casefold()


def wa_read_aging_rows(path):
    """Read a raw Inventory Aging export for the WhatsApp tab. Independent of
    the pandas-based Email/Sheets pipeline above — this tab works on the file
    exactly as exported, no cleaning/merging applied, so the user sees (and
    sends) literally what is in the source file."""
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    ws = wb.worksheets[0]
    rows_iter = ws.iter_rows(values_only=True)
    header = next(rows_iter, None)
    if header is None:
        wb.close()
        raise ValueError("File appears to be empty.")
    hmap = {_wa_norm_header(h): i for i, h in enumerate(header)}
    col_idx = {}
    for key, candidates in WA_SRC_COLUMNS.items():
        idx = None
        for c in candidates:
            if _wa_norm_header(c) in hmap:
                idx = hmap[_wa_norm_header(c)]
                break
        col_idx[key] = idx
    missing = [k for k, v in col_idx.items() if v is None]
    if missing:
        raise ValueError(
            "This file is missing expected column(s): " + ", ".join(missing)
            + "\n\nExpected (any of): " + "; ".join(
                f"{k} = {' / '.join(v)}" for k, v in WA_SRC_COLUMNS.items())
        )

    out = []
    for row in rows_iter:
        if row is None or all(v is None for v in row):
            continue

        def get(key):
            i = col_idx[key]
            return row[i] if i is not None and i < len(row) else None

        district = get("district")
        if district is None or str(district).strip() == "":
            continue
        age = get("age_in_company")
        try:
            age = int(age) if age is not None and str(age).strip() != "" else None
        except Exception:
            age = None
        po = get("po_date")
        if isinstance(po, (datetime, date)):
            po_str = po.strftime("%m/%d/%Y")
        else:
            po_str = str(po).strip() if po not in (None, "") else ""
        out.append({
            "district": str(district).strip(),
            "store": str(get("store") or "").strip(),
            "product": str(get("product") or "").strip(),
            "imei": str(get("imei") or "").strip(),
            "age_in_company": age,
            "po_date": po_str,
        })
    wb.close()
    return out


def wa_unique_districts(rows):
    seen, out = set(), []
    for r in rows:
        d = r["district"]
        if d not in seen:
            seen.add(d)
            out.append(d)
    return sorted(out, key=str.casefold)


def wa_filter_rows(rows, district, aged_days, excluded_stores=None):
    """Devices in `district` aged `aged_days` or more (blank/unknown age
    excluded), minus any store matching an entry in `excluded_stores`
    (same case-insensitive contains-match as the main Exclusions panel,
    via store_matches_pattern)."""
    pats = [s for s in (excluded_stores or []) if s and s.strip()]
    out = []
    for r in rows:
        if r["district"] != district:
            continue
        if r["age_in_company"] is None or r["age_in_company"] < aged_days:
            continue
        if pats and any(store_matches_pattern(r["store"], p) for p in pats):
            continue
        out.append(r)
    return out


def wa_render_table_image(district, rows, logo_path=None, group_label=None):
    """GFH-branded screenshot: navy top bar + logo, dark-navy column header
    row (matches every other GFH/VDS app's header styling), alternating rows."""
    from PIL import Image, ImageDraw, ImageFont

    navy = (9, 13, 38)
    red = (232, 48, 48)
    white = (255, 255, 255)
    row_alt = (238, 242, 248)
    text_c = (22, 33, 58)
    border_c = (200, 210, 224)

    def font(size, bold=False):
        for c in (
            "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf" if bold
            else "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
            "C:/Windows/Fonts/segoeuib.ttf" if bold else "C:/Windows/Fonts/segoeui.ttf",
            "C:/Windows/Fonts/calibrib.ttf" if bold else "C:/Windows/Fonts/calibri.ttf",
        ):
            try:
                if os.path.exists(c):
                    return ImageFont.truetype(c, size)
            except Exception:
                pass
        return ImageFont.load_default()

    pad = 14
    row_h = 30
    brand_bar_h = 90   # matches FixedHeaderManager's standard GFH header height
    col_header_h = 36
    width = max(
        sum(c[2] for c in WA_TABLE_COLUMNS) + pad * 2,
        560,  # keeps the centered title legible even for a 1-2 column table
    )
    height = brand_bar_h + col_header_h + max(1, len(rows)) * row_h + pad

    img = Image.new("RGB", (width, height), white)
    d = ImageDraw.Draw(img)

    # Standard GFH header bar: logo + red divider on the left, title
    # CENTERED across the full bar \u2014 identical layout to every other
    # GFH/VDS app window (FixedHeaderManager: logo | divider | centered title).
    d.rectangle([0, 0, width, brand_bar_h], fill=navy)
    logo_right_edge = pad
    if logo_path and os.path.exists(logo_path):
        try:
            logo = Image.open(logo_path).convert("RGBA")
            logo.thumbnail((150, 56))
            img.paste(logo, (pad, (brand_bar_h - logo.height) // 2), logo)
            logo_right_edge = pad + logo.width
        except Exception:
            pass
    divider_x = logo_right_edge + 14
    d.rectangle([divider_x, 14, divider_x + 3, brand_bar_h - 14], fill=red)

    title = f"Aged Devices \u2014 {district}"
    if group_label:
        title += f"  ({group_label})"
    d.text((width / 2, brand_bar_h / 2), title, font=font(16, True),
           fill=white, anchor="mm")

    y = brand_bar_h
    d.rectangle([0, y, width, y + col_header_h], fill=navy)
    x = pad
    f_head = font(11, True)
    for label, _key, w in WA_TABLE_COLUMNS:
        d.text((x + 6, y + col_header_h // 2), label, font=f_head, fill=white, anchor="lm")
        x += w

    y += col_header_h
    f_row = font(11)
    if not rows:
        d.text((pad + 6, y + row_h // 2), "No devices at/above this age threshold.",
                font=f_row, fill=text_c, anchor="lm")
    for i, r in enumerate(rows):
        if i % 2 == 1:
            d.rectangle([0, y, width, y + row_h], fill=row_alt)
        x = pad
        for _label, key, w in WA_TABLE_COLUMNS:
            val = r.get(key, "")
            val = "" if val is None else str(val)
            d.text((x + 6, y + row_h // 2), val, font=f_row, fill=text_c, anchor="lm")
            x += w
        d.line([0, y + row_h, width, y + row_h], fill=border_c, width=1)
        y += row_h

    return img


def load_whatsapp_config():
    path = os.path.join(get_app_dir(), WHATSAPP_CONFIG_FILE)
    cfg = {
        "aged_days": 30, "last_file": "", "districts": {}, "excluded_stores": [],
        "message_template": WA_DEFAULT_MESSAGE,
    }
    if os.path.exists(path):
        try:
            with open(path, "r", encoding="utf-8") as f:
                saved = json.load(f)
            cfg["aged_days"] = int(saved.get("aged_days", 30))
            cfg["last_file"] = saved.get("last_file", "") or ""
            d = saved.get("districts", {})
            if isinstance(d, dict):
                for k, v in d.items():
                    if not isinstance(v, dict):
                        continue
                    cfg["districts"][str(k)] = {
                        "checked": bool(v.get("checked", False)),
                        "group": str(v.get("group", "") or ""),
                    }
            ex = saved.get("excluded_stores", [])
            if isinstance(ex, list):
                cfg["excluded_stores"] = [str(s).strip() for s in ex if str(s).strip()]
            msg = saved.get("message_template", "")
            if isinstance(msg, str) and msg.strip():
                cfg["message_template"] = msg
        except Exception:
            pass
    return cfg


def save_whatsapp_config(cfg):
    path = os.path.join(get_app_dir(), WHATSAPP_CONFIG_FILE)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(cfg, f, indent=2, ensure_ascii=False)


def wa_copy_image_to_clipboard(img):
    """Put a PIL Image on the Windows clipboard as a CF_DIB bitmap, exactly
    like Ctrl+C on an image — WhatsApp Desktop's message box accepts Ctrl+V
    of this the same way it accepts a copied screenshot."""
    import io as _io
    import win32clipboard  # part of pywin32, already a required dependency

    buf = _io.BytesIO()
    img.convert("RGB").save(buf, "BMP")
    data = buf.getvalue()[14:]  # strip the 14-byte BITMAPFILEHEADER; CF_DIB wants the rest
    win32clipboard.OpenClipboard()
    try:
        win32clipboard.EmptyClipboard()
        win32clipboard.SetClipboardData(win32clipboard.CF_DIB, data)
    finally:
        win32clipboard.CloseClipboard()


# ===========================================================================
# ROBUST WHATSAPP DESKTOP LAYER - ported verbatim (logic unchanged) from
# GFH_Audit_Automation.py / GFH_Inventory_Audit.py, both already shipping and
# proven in production. Pure ctypes + win32clipboard - no pywinauto, no
# comtypes, nothing that adds meaningful bundle size or does COM codegen on
# launch. (An earlier pywinauto/comtypes-based version of this function
# caused the app to hang on startup in a frozen EXE - comtypes' UI-Automation
# codegen has nowhere persistent to cache in a onefile build, so it redid
# that slow codegen on every single launch before the window could appear.
# This layer avoids that family of problem entirely.)
# ===========================================================================
import ctypes as _wa_ct
from ctypes import wintypes as _wa_wt

_WA_BROWSER_PROCS = ("msedge", "chrome", "firefox", "brave", "opera",
                     "iexplore", "vivaldi")
_WA_SELF_PROCS = ("gfh", "python")
_WA_EXT_VKS = (0x21, 0x22, 0x23, 0x24, 0x25, 0x26, 0x27, 0x28, 0x2D, 0x2E)
_WA_VK_MAP = {
    "ctrl": 0x11, "control": 0x11, "alt": 0x12, "menu": 0x12,
    "shift": 0x10, "win": 0x5B, "enter": 0x0D, "return": 0x0D,
    "tab": 0x09, "esc": 0x1B, "escape": 0x1B, "backspace": 0x08,
    "delete": 0x2E, "del": 0x2E, "home": 0x24, "end": 0x23,
    "pageup": 0x21, "pagedown": 0x22, "up": 0x26, "down": 0x28,
    "left": 0x25, "right": 0x27, "space": 0x20, "a": 0x41, "c": 0x43,
    "f": 0x46, "v": 0x56, "x": 0x58,
}
_WA_LOGGER = None


def _wa_log(msg):
    try:
        (_WA_LOGGER or print)(msg)
    except Exception:
        pass


def _wa_user32():
    u = _wa_ct.windll.user32
    u.GetForegroundWindow.restype = _wa_wt.HWND
    u.GetAncestor.restype = _wa_wt.HWND
    u.GetAncestor.argtypes = [_wa_wt.HWND, _wa_wt.UINT]
    return u


def _wa_proc_name(pid):
    """Lower-case exe name of a process id ('' when unknown)."""
    try:
        k32 = _wa_ct.windll.kernel32
        k32.OpenProcess.restype = _wa_ct.c_void_p
        k32.OpenProcess.argtypes = [_wa_wt.DWORD, _wa_wt.BOOL, _wa_wt.DWORD]
        k32.QueryFullProcessImageNameW.argtypes = [
            _wa_ct.c_void_p, _wa_wt.DWORD, _wa_wt.LPWSTR,
            _wa_ct.POINTER(_wa_wt.DWORD)]
        k32.CloseHandle.argtypes = [_wa_ct.c_void_p]
        h = k32.OpenProcess(0x1000, False, pid)  # QUERY_LIMITED_INFORMATION
        if not h:
            return ""
        try:
            buf = _wa_ct.create_unicode_buffer(520)
            n = _wa_wt.DWORD(520)
            if k32.QueryFullProcessImageNameW(h, 0, buf, _wa_ct.byref(n)):
                return os.path.basename(buf.value).lower()
        finally:
            k32.CloseHandle(h)
    except Exception:
        pass
    return ""


def _wa_window_pid(hwnd):
    pid = _wa_wt.DWORD(0)
    _wa_user32().GetWindowThreadProcessId(hwnd, _wa_ct.byref(pid))
    return pid.value


def _wa_find_hwnd():
    """HWND of the real WhatsApp Desktop window (never a browser tab or this
    app's own window), or None. Minimised windows are included."""
    if sys.platform != "win32":
        return None
    try:
        u = _wa_user32()
        own = os.getpid()
        found = []
        enum_proc = _wa_ct.WINFUNCTYPE(_wa_wt.BOOL, _wa_wt.HWND, _wa_wt.LPARAM)

        def _cb(hwnd, _lp):
            try:
                if not u.IsWindowVisible(hwnd):
                    return True
                n = u.GetWindowTextLengthW(hwnd)
                if n <= 0:
                    return True
                buf = _wa_ct.create_unicode_buffer(n + 1)
                u.GetWindowTextW(hwnd, buf, n + 1)
                title = buf.value
                if "whatsapp" not in title.lower():
                    return True
                pid = _wa_window_pid(hwnd)
                if pid == own:
                    return True
                pn = _wa_proc_name(pid)
                if any(b in pn for b in _WA_BROWSER_PROCS):
                    return True
                if any(pn.startswith(s) for s in _WA_SELF_PROCS):
                    return True
                if "whatsapp" in pn:
                    score = 0
                elif pn.startswith("applicationframehost"):
                    score = 1
                elif title.strip().lower() == "whatsapp":
                    score = 2
                else:
                    return True
                found.append((score, hwnd))
            except Exception:
                pass
            return True

        u.EnumWindows(enum_proc(_cb), 0)
        found.sort(key=lambda t: t[0])
        return found[0][1] if found else None
    except Exception:
        return None


def _wa_is_foreground(hwnd):
    """True when WhatsApp (or its child/core window) owns the foreground."""
    try:
        u = _wa_user32()
        fg = u.GetForegroundWindow()
        if not fg:
            return False
        if fg == hwnd or u.GetAncestor(fg, 2) == hwnd:  # GA_ROOT
            return True
        return "whatsapp" in _wa_proc_name(_wa_window_pid(fg))
    except Exception:
        return False


def _wa_focus_hwnd(hwnd, log=None):
    """Bring WhatsApp to the foreground and VERIFY it. Returns True/False.
    Escalates: ALT-tap unlock + AttachThreadInput -> SwitchToThisWindow ->
    minimise/restore. Only un-minimises when actually minimised, so a snapped
    or maximised window keeps its layout."""
    try:
        u = _wa_user32()
        k32 = _wa_ct.windll.kernel32
        try:  # allow this process to take the foreground
            u.SystemParametersInfoW(0x2001, 0, 0, 0x0002)
        except Exception:
            pass
        if u.IsIconic(hwnd):
            u.ShowWindow(hwnd, 9)  # SW_RESTORE (from minimised only)
            time.sleep(0.5)
        for attempt in range(1, 5):
            _wa_win32_key_down(0x12)  # ALT tap unlocks SetForegroundWindow
            _wa_win32_key_up(0x12)
            fg = u.GetForegroundWindow()
            fg_thread = u.GetWindowThreadProcessId(fg, None) if fg else 0
            our_thread = k32.GetCurrentThreadId()
            wa_thread = u.GetWindowThreadProcessId(hwnd, None)
            pairs = [(fg_thread, our_thread), (fg_thread, wa_thread)]
            for a, b in pairs:
                if a and b and a != b:
                    u.AttachThreadInput(a, b, True)
            u.BringWindowToTop(hwnd)
            u.ShowWindow(hwnd, 5)  # SW_SHOW (keeps maximised state)
            u.SetForegroundWindow(hwnd)
            for a, b in pairs:
                if a and b and a != b:
                    u.AttachThreadInput(a, b, False)
            if attempt >= 2:
                u.SwitchToThisWindow(hwnd, True)
            if attempt >= 3:
                u.ShowWindow(hwnd, 6)  # SW_MINIMIZE
                time.sleep(0.3)
                u.ShowWindow(hwnd, 9)  # restores previous state
            time.sleep(0.4)
            if _wa_is_foreground(hwnd):
                return True
    except Exception as exc:
        if log:
            log(f"WhatsApp focus error: {exc}")
    return False


def _wa_launch():
    """Open WhatsApp Desktop (several methods)."""
    import subprocess as _sp
    for cmd in (["explorer.exe", "whatsapp:"],
                "start whatsapp:",
                r'explorer.exe "shell:AppsFolder\5319275A.WhatsAppDesktop_cv1g1gvanyjgm!App"'):
        try:
            _sp.Popen(cmd, shell=isinstance(cmd, str))
            return
        except Exception:
            continue


def _wa_ensure_focused(log=None):
    log = log or _wa_log
    """Find (launching if needed) WhatsApp Desktop, focus it, verify it.
    Returns hwnd on success, None on failure."""
    hwnd = _wa_find_hwnd()
    if hwnd is None:
        log("WhatsApp Desktop not open — launching...")
        _wa_launch()
        for _ in range(40):  # wait up to 20 s for the window
            time.sleep(0.5)
            hwnd = _wa_find_hwnd()
            if hwnd:
                time.sleep(2.0)  # let the chat list finish loading
                break
    if hwnd is None:
        log("WhatsApp Desktop window not found — is it installed and signed in?")
        return None
    for _ in range(3):
        if _wa_focus_hwnd(hwnd, log):
            time.sleep(0.4)
            log("WhatsApp Desktop focused.")
            return hwnd
        hwnd = _wa_find_hwnd() or hwnd
        time.sleep(1.0)
    hint = ""
    try:
        if _wa_ct.windll.shell32.IsUserAnAdmin():
            hint = (" This app is running as Administrator; Windows blocks it "
                    "from controlling a normal-user WhatsApp - run both at "
                    "the same level.")
    except Exception:
        pass
    log("Could not bring WhatsApp Desktop to the front." + hint)
    return None


def _wa_guard(hwnd, log=None):
    log = log or _wa_log
    """Re-verify (and re-take) focus right before typing. Never type blind."""
    if _wa_is_foreground(hwnd):
        return True
    log("WhatsApp lost focus — re-focusing...")
    if _wa_focus_hwnd(hwnd, log):
        time.sleep(0.3)
        return True
    log("WhatsApp is not in front — aborting so nothing is typed elsewhere.")
    return False


# ---- keyboard: scan-code aware key events + Unicode SendInput text --------
class _WA_KEYBDINPUT(_wa_ct.Structure):
    _fields_ = [("wVk", _wa_wt.WORD), ("wScan", _wa_wt.WORD),
                ("dwFlags", _wa_wt.DWORD), ("time", _wa_wt.DWORD),
                ("dwExtraInfo", _wa_ct.c_size_t)]


class _WA_MOUSEINPUT(_wa_ct.Structure):
    _fields_ = [("dx", _wa_ct.c_long), ("dy", _wa_ct.c_long),
                ("mouseData", _wa_wt.DWORD), ("dwFlags", _wa_wt.DWORD),
                ("time", _wa_wt.DWORD), ("dwExtraInfo", _wa_ct.c_size_t)]


class _WA_INPUT(_wa_ct.Structure):
    class _U(_wa_ct.Union):
        _fields_ = [("ki", _WA_KEYBDINPUT), ("mi", _WA_MOUSEINPUT)]
    _anonymous_ = ("u",)
    _fields_ = [("type", _wa_wt.DWORD), ("u", _U)]


def _wa_send_key(vk, scan, flags):
    inp = _WA_INPUT(type=1)
    inp.ki = _WA_KEYBDINPUT(vk, scan, flags, 0, 0)
    _wa_ct.windll.user32.SendInput(1, _wa_ct.byref(inp), _wa_ct.sizeof(inp))


def _wa_save_rect(hwnd):
    """Window rect before focus, so snap-layout can be restored after
    (identical to WhatsAppSender._save_whatsapp_rect in the proven apps)."""
    try:
        rect = _wa_wt.RECT()
        _wa_ct.windll.user32.GetWindowRect(hwnd, _wa_ct.byref(rect))
        return (rect.left, rect.top, rect.right - rect.left, rect.bottom - rect.top)
    except Exception:
        return None


def _wa_click(x, y):
    """Left-click at an absolute screen coordinate WITHOUT moving the real
    OS cursor: finds the actual child window under that point and posts it
    WM_MOUSEMOVE/WM_LBUTTONDOWN/WM_LBUTTONUP directly (SendMessage, not
    SendInput+SetCursorPos) - invisible on screen, the person's own mouse
    is never touched while this runs in the background."""
    u = _wa_ct.windll.user32
    pt = _wa_wt.POINT(int(x), int(y))
    target = u.WindowFromPoint(pt)
    if not target:
        return
    cpt = _wa_wt.POINT(int(x), int(y))
    u.ScreenToClient(target, _wa_ct.byref(cpt))
    lparam = (cpt.y << 16) | (cpt.x & 0xFFFF)
    WM_MOUSEMOVE, WM_LBUTTONDOWN, WM_LBUTTONUP, MK_LBUTTON = 0x0200, 0x0201, 0x0202, 0x0001
    u.SendMessageW(target, WM_MOUSEMOVE, 0, lparam)
    u.SendMessageW(target, WM_LBUTTONDOWN, MK_LBUTTON, lparam)
    time.sleep(0.03)
    u.SendMessageW(target, WM_LBUTTONUP, 0, lparam)


def _wa_win32_key_down(vk_code):
    scan = _wa_ct.windll.user32.MapVirtualKeyW(vk_code, 0)
    _wa_send_key(vk_code, scan, 0x1 if vk_code in _WA_EXT_VKS else 0)


def _wa_win32_key_up(vk_code):
    scan = _wa_ct.windll.user32.MapVirtualKeyW(vk_code, 0)
    _wa_send_key(vk_code, scan, 0x2 | (0x1 if vk_code in _WA_EXT_VKS else 0))


def _wa_win32_hotkey(*keys):
    codes = []
    for k in keys:
        kl = k.lower()
        if kl in _WA_VK_MAP:
            codes.append(_WA_VK_MAP[kl])
        elif len(k) == 1:
            codes.append(ord(k.upper()))
        else:
            return
    for c in codes:
        _wa_win32_key_down(c)
        time.sleep(0.04)
    time.sleep(0.04)
    for c in reversed(codes):
        _wa_win32_key_up(c)
        time.sleep(0.04)


def _wa_win32_press(key):
    kl = key.lower()
    if kl in _WA_VK_MAP:
        code = _WA_VK_MAP[kl]
    elif len(key) == 1:
        code = ord(key.upper())
    else:
        return
    _wa_win32_key_down(code)
    time.sleep(0.04)
    _wa_win32_key_up(code)
    time.sleep(0.04)


def _wa_win32_write(text, interval=0.0):
    """Type text as Unicode characters (layout independent, any language)."""
    for ch in str(text):
        for up in (0, 2):
            _wa_send_key(0, ord(ch), 0x4 | up)  # KEYEVENTF_UNICODE
        time.sleep(max(0.01, interval))


def _wa_set_clipboard_text(text):
    """Put plain text on the clipboard via ctypes (no pyperclip needed)."""
    u = _wa_ct.windll.user32
    k = _wa_ct.windll.kernel32
    k.GlobalAlloc.argtypes = [_wa_wt.UINT, _wa_ct.c_size_t]
    k.GlobalAlloc.restype = _wa_ct.c_void_p
    k.GlobalLock.argtypes = [_wa_ct.c_void_p]
    k.GlobalLock.restype = _wa_ct.c_void_p
    k.GlobalUnlock.argtypes = [_wa_ct.c_void_p]
    u.OpenClipboard.argtypes = [_wa_ct.c_void_p]
    u.SetClipboardData.argtypes = [_wa_wt.UINT, _wa_ct.c_void_p]
    u.SetClipboardData.restype = _wa_ct.c_void_p
    data = (str(text) + "\x00").encode("utf-16-le")
    for _ in range(10):  # clipboard may be briefly busy
        if u.OpenClipboard(None):
            break
        time.sleep(0.05)
    else:
        return False
    try:
        u.EmptyClipboard()
        h = k.GlobalAlloc(0x0002, len(data))
        if not h:
            return False
        p = k.GlobalLock(h)
        _wa_ct.memmove(p, data, len(data))
        k.GlobalUnlock(h)
        return bool(u.SetClipboardData(13, h))  # CF_UNICODETEXT
    finally:
        u.CloseClipboard()


def _wa_paste_text(text):
    """Enter text into the focused field: clipboard + Ctrl+V, falling back to
    Unicode typing. NOT guarded (matches _wa_paste_text in the confirmed-
    working VidaPay_Incentive_Extractor.py/Device_Ordering.py exactly) - an
    earlier version of this file guarded every single hotkey/paste/press
    call, which meant the aggressive refocus escalation in _wa_guard (ALT-
    tap, AttachThreadInput, SwitchToThisWindow, even a minimize/restore
    cycle) could fire BETWEEN every keystroke of a single logical step, and
    very plausibly collapsed whatever WhatsApp had just opened (e.g. the
    chat-list search) back to its default view before the next keystroke
    landed. Guarding only belongs at specific checkpoints between logical
    steps - see _wa_type_group_search and send_image_to_whatsapp_group,
    which now guard exactly where _wa_open_chat/_wa_send_in_chat do and
    nowhere else."""
    try:
        if _wa_set_clipboard_text(text):
            time.sleep(0.2)
            _wa_win32_hotkey("ctrl", "v")
            time.sleep(0.3)
            return
    except Exception:
        pass
    _wa_win32_write(text, 0.02)


# Unguarded aliases - same as the confirmed-working apps' own
# "_wa_hotkey = _wa_win32_hotkey" / "_wa_press = _wa_win32_press" routing.
_wa_hotkey = _wa_win32_hotkey
_wa_press = _wa_win32_press
_wa_write = _wa_win32_write


def _wa_type_group_search(hwnd, group_name, log=None):
    """Open the chat-list search and jump to `group_name` by clicking it -
    no keyboard shortcut involved anywhere in this. Ctrl+F and Ctrl+K were
    both tried and both reproducibly failed on this WhatsApp Desktop build:
    live testing showed the group name landing in whatever chat was
    already open instead, proving focus never actually reached a search
    box either way - so neither is used, by design, not as a fallback.

    Clicks the "Search or start a new chat" box by its position
    (proportional to the actual window rect, not a fixed pixel guess) using
    posted window messages (_wa_click - no SetCursorPos/SendInput, so the
    person's real mouse never visibly moves), types the group name
    straight in (no Ctrl+A - the box should be empty right after opening),
    then clicks the first filtered result INSTEAD OF pressing Enter - a
    missed click then just types into nothing rather than risking a send
    into the wrong chat. Returns True when the chat should now be open,
    False if focus was lost and nothing more should be typed. Guarded only
    at specific checkpoints (never before every single keystroke - an
    earlier version did that, and the aggressive refocus it could trigger,
    up to a minimize/restore cycle, was itself capable of collapsing
    whatever had just opened)."""
    log = log or _wa_log
    rect = _wa_save_rect(hwnd)
    if rect is None:
        log("Could not read the WhatsApp window's position for a click.")
        return False
    left, top, w, h = rect
    search_x, search_y = left + max(70, int(w * 0.12)), top + int(h * 0.065)
    log(f"  click method: window=({left},{top},{w}x{h}) -> search click at ({search_x},{search_y})")
    _wa_click(search_x, search_y)
    time.sleep(0.4)
    if not _wa_guard(hwnd, log):
        return False
    _wa_paste_text(group_name)
    time.sleep(1.2)
    result_y = search_y + 64
    log(f"  click method: first-result click at ({search_x},{result_y})")
    _wa_click(search_x, result_y)  # click the result, never Enter here
    time.sleep(1.3)
    return _wa_guard(hwnd, log)
# ==================== END ROBUST WHATSAPP DESKTOP LAYER ====================


def send_image_to_whatsapp_group(group_name, img, log=print, timeout=20, caption=""):
    """WhatsApp Desktop (Win32 app, not WhatsApp Web) automation.

    Focus WhatsApp (via _wa_ensure_focused - launching it only if it isn't
    already running, never force-relaunching an already-open instance),
    find the group by clicking the search box (_wa_type_group_search - no
    keyboard shortcut: Ctrl+F and Ctrl+K were both tried and both
    reproducibly failed on this WhatsApp Desktop build), paste the image +
    caption and send. Every step is logged and wrapped so a failure
    anywhere still leaves the image sitting in the compose box (and on the
    clipboard) for the user to finish by hand with one Enter — it never
    silently does nothing.

    Returns (ok: bool, message: str).
    """
    global _WA_LOGGER
    _WA_LOGGER = log
    if not _HAS_WA_AUTOMATION_DEPS:
        return False, ("pyautogui not installed — run: pip install pyautogui"
                        " (image copied to clipboard; open WhatsApp Desktop,"
                        " the chat, and press Ctrl+V yourself).")
    try:
        import pyautogui  # noqa: F401  (fail-fast check only, same as the proven apps)
    except Exception as e:
        return False, f"pyautogui import failed: {e}"

    try:
        wa_copy_image_to_clipboard(img)
    except Exception as e:
        return False, f"Could not copy image to clipboard: {e}"

    log(f"Opening WhatsApp Desktop for {group_name}...")
    try:
        hwnd = _wa_ensure_focused(log)
    except Exception as e:
        hwnd = None
        log(f"WhatsApp focus error: {e}")
    if hwnd is None:
        return False, ("Could not bring WhatsApp Desktop to the front. Is it"
                        " installed and signed in? (Image is on the"
                        " clipboard — you can paste it yourself with Ctrl+V.)")

    log(f"Searching group: {group_name}")
    try:
        opened = _wa_type_group_search(hwnd, group_name, log=log)
    except Exception as e:
        opened = False
        log(f"Search error: {e}")
    if not opened:
        return False, (f"Could not open group '{group_name}' — WhatsApp lost focus or the "
                        "search didn't land (nothing further was typed, so nothing should "
                        "have been sent). Image is on the clipboard if you want to finish by hand.")

    # From here on, same checkpoint-guard style as _wa_send_in_chat: guard
    # right before anything that could land in the wrong place, never guard
    # mid-keystroke.
    if not _wa_guard(hwnd, log):
        return False, "WhatsApp lost focus before the image could be pasted — nothing was sent."
    try:
        _wa_hotkey("ctrl", "v")  # opens WhatsApp's image-preview/send dialog
    except Exception as e:
        return False, f"Could not paste the image: {e}"
    time.sleep(3.0)  # let the preview dialog fully load

    if caption:
        if not _wa_guard(hwnd, log):
            return False, ("WhatsApp lost focus before the caption could be typed — the image"
                            " may be sitting unsent in the preview dialog.")
        try:
            # WhatsApp Desktop focuses the preview's caption field
            # automatically when the image-attach dialog opens - type
            # straight into it.
            _wa_paste_text(caption)
        except Exception as e:
            return False, f"Could not type the caption: {e}"
        time.sleep(0.8)

    if not _wa_guard(hwnd, log):
        return False, ("WhatsApp lost focus before sending — the image/caption may be sitting"
                        " unsent in the preview dialog; check WhatsApp and press Enter yourself.")
    try:
        _wa_press("enter")  # Enter confirms the image-preview dialog
    except Exception as e:
        return False, f"Could not confirm the send: {e}"
    time.sleep(2.0)
    return True, f"Sent to '{group_name}'" + (" with caption." if caption else ".")


# ==========================================================
# DATA CLEANING / PROCESSING
# ==========================================================
class PipelineError(Exception):
    pass

def resolve_column(df, logical_name):
    for candidate in COLUMN_CANDIDATES[logical_name]:
        if candidate in df.columns:
            return candidate
    raise PipelineError(
        f'Could not find a "{logical_name}" column. Looked for: '
        + ", ".join(COLUMN_CANDIDATES[logical_name])
    )



def format_date_value(value):
    if pd.isna(value):
        return ""
    if isinstance(value, (pd.Timestamp, datetime)):
        return value.strftime("%m/%d/%Y")
    return str(value).strip()



def normalize_district(value):
    text = "" if value is None else str(value).strip()
    if re.match(r"(?i)^arizona\b", text):
        return "Arizona"
    return text




def load_and_clean(excel_path, log):
    """Load Excel, resolve columns, remove SIM/blank/blocked rows, normalize districts."""
    log(f"Reading: {excel_path}")
    df = pd.read_excel(excel_path, engine="openpyxl", dtype=str)
    df = df.dropna(how="all")
    log(f"  Loaded {len(df)} rows.")

    col_district = resolve_column(df, "District")
    col_store = resolve_column(df, "Store")
    col_desc = resolve_column(df, "Product Description")
    col_serial = resolve_column(df, "Serial 1")
    col_age = resolve_column(df, "Age in Company")
    col_po = resolve_column(df, "PO Date")
    col_retail = resolve_column(df, "Retail Price")

    # Remove SIM rows
    is_sim = df[col_desc].astype(str).str.contains(r"\bSIM\b|SIM CARD|SIM KIT", case=False, na=False, regex=True)
    sim_count = int(is_sim.sum())
    df = df[~is_sim].copy()
    log(f"  Removed {sim_count} SIM row(s).")

    # Remove blank/NaN Serial 1
    before = len(df)
    df[col_serial] = df[col_serial].astype(str).str.strip()
    df = df[(df[col_serial] != "") & (df[col_serial].str.lower() != "nan")]
    log(f"  Removed {before - len(df)} row(s) with blank Serial 1.")

    # Remove always-blocked IMEIs (editable in the Exclusions panel;
    # compared digits-only so formatted serials still match)
    _blocked = {normalize_imei(x) for x in load_exclusions().get("blocked_imeis", [])}
    _blocked.discard("")
    before2 = len(df)
    if _blocked:
        df = df[~df[col_serial].astype(str).str.strip().map(normalize_imei).isin(list(_blocked))]
    log(f"  Removed {before2 - len(df)} blocked IMEI row(s).")

    # Convert numeric columns
    df[col_age] = pd.to_numeric(df[col_age], errors="coerce")
    df[col_retail] = pd.to_numeric(df[col_retail], errors="coerce")
    df[col_po] = pd.to_datetime(df[col_po], errors="coerce")

    # Normalize district names (merge Arizona sub-districts)
    df["_District"] = df[col_district].apply(normalize_district)
    merged_count = (df[col_district].astype(str).str.strip() != df["_District"]).sum()
    if merged_count:
        log(f"  Merged {merged_count} row(s) from Arizona sub-districts into 'Arizona'.")

    log(f"  {len(df)} rows remain after cleaning.")
    return df, col_store, col_desc, col_serial, col_age, col_po, col_retail

def build_aged_subset(df, col_age, log):
    aged = df[df[col_age] > AGE_THRESHOLD_DAYS].copy()
    log(f"  {len(aged)} of {len(df)} devices are aged over {AGE_THRESHOLD_DAYS} days.")
    return aged



def build_district_tabs(aged, col_store, col_desc, col_serial, col_age, col_po, log,
                        ensure_districts=None):
    """One tab per district. ensure_districts forces an (empty) tab for every
    listed district that has zero rows left after exclusions, so those
    districts still get an empty Google Sheet tab instead of their old tab
    being silently removed as leftover."""
    tabs = {}
    for district in sorted(aged["_District"].dropna().unique()):
        if not district:
            continue
        group = aged[aged["_District"] == district]
        tab_df = pd.DataFrame(
            {
                "District": district,
                "Store": group[col_store].fillna("").astype(str).str.strip(),
                "Product Description": group[col_desc].fillna("").astype(str).str.strip(),
                "Serial 1": group[col_serial],
                "Age in Company": group[col_age],
                "PO Date": group[col_po].apply(format_date_value),
            }
        ).sort_values(["Store", "Age in Company"], ascending=[True, False]).reset_index(drop=True)
        tabs[district] = tab_df
        log(f"  {district}: {len(tab_df)} aged device row(s)")

    if ensure_districts:
        for district in sorted(set(ensure_districts) - set(tabs)):
            tabs[district] = pd.DataFrame({
                "District": pd.Series(dtype=str),
                "Store": pd.Series(dtype=str),
                "Product Description": pd.Series(dtype=str),
                "Serial 1": pd.Series(dtype=str),
                "Age in Company": pd.Series(dtype=float),
                "PO Date": pd.Series(dtype=str),
            })
            log(f"  {district}: 0 aged device row(s) (all devices excluded)")
    return tabs


# ==========================================================
# GOOGLE SHEETS UPLOAD
# ==========================================================


# ==========================================================
# LOCAL EXCEL + EMAIL
# ==========================================================

def save_district_xlsx(df_d, district, out_dir):
    safe     = re.sub(r'[\\/:*?"<>|]',"_",district).strip()
    out_path = Path(out_dir) / f"Inventory_Aging_District_{safe}.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = district[:31]
    headers = list(df_d.columns)
    hf = Font(name="Calibri",bold=True,color="FFFFFF",size=11)
    hfill = PatternFill("solid",fgColor="1F3864")
    ha = Alignment(horizontal="center",vertical="center",wrap_text=True)
    thin = XlSide(style="thin",color="FFFFFF")
    hb = XlBorder(left=thin,right=thin,top=thin,bottom=thin)
    for ci,h in enumerate(headers,1):
        c=ws.cell(row=1,column=ci,value=h); c.font=hf; c.fill=hfill; c.alignment=ha; c.border=hb
    ws.row_dimensions[1].height=28
    df_ = Font(name="Calibri",size=10)
    da = Alignment(vertical="center")
    db = XlBorder(left=XlSide(style="thin",color="CCCCCC"),right=XlSide(style="thin",color="CCCCCC"),
                top=XlSide(style="thin",color="CCCCCC"),bottom=XlSide(style="thin",color="CCCCCC"))
    af = PatternFill("solid",fgColor="EAF2FF")
    col_w={"District":18,"Store":22,"Product Desc Full":45,"Serial 1":20,"Age in Company":16}
    for ri,(_,row) in enumerate(df_d.iterrows(),2):
        fill=af if ri%2==0 else None
        for ci,h in enumerate(headers,1):
            v=row[h]
            if h=="Serial 1":
                cell=ws.cell(row=ri,column=ci,value=str(v) if pd.notna(v) else ""); cell.number_format="@"
            elif h=="Age in Company":
                cell=ws.cell(row=ri,column=ci,value=int(v) if pd.notna(v) else "")
            else:
                cell=ws.cell(row=ri,column=ci,value=v if pd.notna(v) else "")
            cell.font=df_; cell.alignment=da; cell.border=db
            if fill: cell.fill=fill
    for ci,h in enumerate(headers,1):
        ws.column_dimensions[get_column_letter(ci)].width=col_w.get(h,18)
    ws.freeze_panes="A2"; ws.auto_filter.ref=ws.dimensions
    wb.save(out_path); return out_path

# ─────────────────────────────────────────────────────────────────────────────
# EMAIL
# ─────────────────────────────────────────────────────────────────────────────

def build_html_email(district, to_address="", sender_cfg=None):
    """sender_cfg: dict with sender_name/sender_email/sender_title/sender_mobile,
    all editable in Settings. No identity is hardcoded here any more."""
    sender_cfg = sender_cfg or {}
    sender_name   = sender_cfg.get("sender_name", "") or "GFH Telecom"
    sender_email  = sender_cfg.get("sender_email", "")
    sender_title  = sender_cfg.get("sender_title", "")
    sender_mobile = sender_cfg.get("sender_mobile", "")

    today = datetime.now().strftime("%B %d, %Y")
    # Extract first name from "Full Name <email>" or "Full Name"
    dm_name = "Team"
    if to_address:
        raw = to_address.split("<")[0].strip()   # "Full Name"
        if raw:
            dm_name = raw.split()[0]

    email_row = (
        f'<tr><td style="color:#e83030;font-weight:700;padding-right:6px;font-family:Calibri;">E:</td>'
        f'<td style="font-family:Calibri;"><a href="mailto:{sender_email}" '
        f'style="color:#090d26;text-decoration:none;">{sender_email}</a></td></tr>'
    ) if sender_email else ""
    mobile_row = (
        f'<tr><td style="color:#e83030;font-weight:700;padding-right:6px;font-family:Calibri;">M:</td>'
        f'<td style="color:#090d26;font-family:Calibri;">{sender_mobile}</td></tr>'
    ) if sender_mobile else ""

    return f"""
<html><body style="font-family:Calibri,Arial,sans-serif;font-size:14px;color:#1a1a1a;">
<p>Dear {dm_name},</p>
<p>Please find attached the <strong>Inventory Aging Report — {district} District</strong> as of {today}.</p>
<p>This report includes all devices with <strong>20+ days</strong> age in company. Kindly review and take necessary action on aging inventory.</p>
<p>Thank you,</p>
<br>
<table style="border-top:2px solid #e83030;padding-top:12px;">
  <tr>
    <td style="padding-right:16px;vertical-align:top;">
      <img src="cid:gfhlogo" width="170" alt="GFH Telecom">
    </td>
    <td style="border-left:2px solid #e83030;padding-left:16px;vertical-align:top;font-family:Calibri,Arial,sans-serif;">
      <div style="font-size:16px;font-weight:700;color:#090d26;">{sender_name}</div>
      <div style="font-size:11px;color:#666;margin-top:2px;">{sender_title}</div>
      <table style="margin-top:8px;font-size:12px;border-collapse:collapse;">
        {email_row}
        {mobile_row}
      </table>
    </td>
  </tr>
</table>
</body></html>"""


def send_via_outlook(to, subject, html_body, attachment_path, log, sender_email="", cc_email=""):
    try:
        import win32com.client as win32, pythoncom
        pythoncom.CoInitialize()
    except ImportError:
        log("ERROR: pywin32 not installed. Run: pip install pywin32"); return False
    logo_tmp = None
    try:
        # Decode logo to temp file
        data = base64.b64decode(GFH_LOGO_B64.replace("\n","").replace(" ","").strip())
        logo_tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".png")
        logo_tmp.write(data); logo_tmp.close(); logo_tmp = logo_tmp.name

        outlook = win32.Dispatch("Outlook.Application")
        mail    = outlook.CreateItem(0)

        # Set sender account, if configured and matching a signed-in Outlook account
        if sender_email:
            for acc in outlook.Session.Accounts:
                if sender_email.lower() in str(acc.SmtpAddress).lower():
                    mail._oleobj_.Invoke(*(64209,0,8,0,acc)); break

        mail.To      = to
        if cc_email:
            mail.CC  = cc_email
        mail.Subject = subject

        # Inline logo attachment (CID)
        att = mail.Attachments.Add(logo_tmp)
        att.PropertyAccessor.SetProperty(
            "http://schemas.microsoft.com/mapi/proptag/0x3712001F","gfhlogo")
        att.PropertyAccessor.SetProperty(
            "http://schemas.microsoft.com/mapi/proptag/0x7FFF000B",True)

        mail.HTMLBody = html_body

        # District Excel attachment
        if attachment_path and Path(attachment_path).exists():
            mail.Attachments.Add(str(Path(attachment_path).resolve()))

        mail.Send()
        return True
    except Exception as exc:
        log(f"Outlook error: {exc}"); return False
    finally:
        if logo_tmp and os.path.exists(logo_tmp):
            try: os.unlink(logo_tmp)
            except: pass

# ─────────────────────────────────────────────────────────────────────────────
# GUI
# ─────────────────────────────────────────────────────────────────────────────


# ==========================================================
# GOOGLE SHEETS UPLOAD
# ==========================================================

def authenticate_google(log):
    creds_path = os.path.join(get_app_dir(), CREDENTIALS_FILE)
    auth_user_path = os.path.join(get_app_dir(), AUTHORIZED_USER_FILE)

    if not os.path.exists(creds_path):
        raise PipelineError(
            f"{CREDENTIALS_FILE} not found beside this script.\n"
            f"Expected here:\n{creds_path}\n\n"
            "Place your Google OAuth Desktop credentials.json in the same "
            "folder as this app (same file used by GFH Rebate Google Drive Sync)."
        )

    log("Connecting to Google (a browser window may open to authorize)...")
    return gspread.oauth(
        credentials_filename=creds_path,
        authorized_user_filename=auth_user_path,
    )



def open_or_create_spreadsheet(gc, log):
    try:
        return gc.open(GOOGLE_SHEET_NAME)
    except SpreadsheetNotFound:
        log(f"Google Sheet '{GOOGLE_SHEET_NAME}' not found. Creating it...")
        return gc.create(GOOGLE_SHEET_NAME)



def get_or_create_worksheet(spreadsheet, title, rows, cols):
    try:
        worksheet = spreadsheet.worksheet(title)
    except WorksheetNotFound:
        worksheet = spreadsheet.add_worksheet(title=title, rows=max(rows, 2), cols=max(cols, 1))
    worksheet.resize(rows=max(rows, 2), cols=max(cols, 1))
    return worksheet



def dataframe_to_values(df):
    values = [list(df.columns)]
    for row in df.itertuples(index=False, name=None):
        clean_row = []
        for v in row:
            if pd.isna(v):
                clean_row.append("")
            elif isinstance(v, bool):
                clean_row.append(v)
            elif isinstance(v, (int, float)):
                clean_row.append(v.item() if hasattr(v, "item") else v)
            else:
                clean_row.append(str(v))
        values.append(clean_row)
    return values



def clear_and_upload_values(worksheet, values, value_input_option="RAW"):
    worksheet.clear()
    if not values:
        return
    worksheet.update(values=values, range_name="A1", value_input_option=value_input_option)



def apply_header_formatting(spreadsheet, worksheet, row_count, col_count):
    sheet_id = worksheet.id
    requests = [
        {
            "updateSheetProperties": {
                "properties": {"sheetId": sheet_id, "gridProperties": {"frozenRowCount": 1}},
                "fields": "gridProperties.frozenRowCount",
            }
        },
        {
            "repeatCell": {
                "range": {
                    "sheetId": sheet_id,
                    "startRowIndex": 0,
                    "endRowIndex": 1,
                    "startColumnIndex": 0,
                    "endColumnIndex": col_count,
                },
                "cell": {
                    "userEnteredFormat": {
                        "backgroundColor": HEADER_BG,
                        "horizontalAlignment": "CENTER",
                        "textFormat": {"foregroundColor": HEADER_FG, "bold": True},
                    }
                },
                "fields": "userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)",
            }
        },
        {
            "autoResizeDimensions": {
                "dimensions": {"sheetId": sheet_id, "dimension": "COLUMNS", "startIndex": 0, "endIndex": col_count}
            }
        },
    ]
    if row_count > 1 and col_count > 0:
        requests.append(
            {
                "setBasicFilter": {
                    "filter": {
                        "range": {
                            "sheetId": sheet_id,
                            "startRowIndex": 0,
                            "endRowIndex": row_count,
                            "startColumnIndex": 0,
                            "endColumnIndex": col_count,
                        }
                    }
                }
            }
        )
    spreadsheet.batch_update({"requests": requests})



def sanitize_tab_name(name):
    cleaned = re.sub(r"[\[\]\*\?/\\:]", "-", str(name)).strip()
    return (cleaned or "Unknown")[:100]



def build_executive_dashboard(sh, df_aged):
    sheet_name = DASHBOARD_TAB_NAME
    
    try:
        ws = sh.worksheet(sheet_name)
        ws.clear()
        try:
            ws.clear_basic_filter()
        except:
            pass
    except gspread.exceptions.WorksheetNotFound:
        ws = sh.add_worksheet(title=sheet_name, rows="100", cols="15")

    sh.reorder_worksheets([ws] + [s for s in sh.worksheets() if s.title != sheet_name])

    col_retail = resolve_column(df_aged, "Retail Price")
    
    total_value = float(pd.to_numeric(df_aged[col_retail], errors="coerce").sum())
    total_devices = len(df_aged)
    
    district_risk = df_aged.groupby('_District')[col_retail].sum().reset_index()
    district_risk = district_risk.sort_values(col_retail, ascending=False)
    top_district = district_risk.iloc[0]['_District'] if not district_risk.empty else "N/A"
    top_district_val = float(district_risk.iloc[0][col_retail]) if not district_risk.empty else 0

    col_store = resolve_column(df_aged, "Store")
    store_risk = df_aged.groupby(col_store)[col_retail].sum().reset_index()
    store_risk = store_risk.sort_values(col_retail, ascending=False)
    top_store = store_risk.iloc[0][col_store] if not store_risk.empty else "N/A"
    top_store_val = float(store_risk.iloc[0][col_retail]) if not store_risk.empty else 0

    col_serial = resolve_column(df_aged, "Serial 1")

    dashboard_data = [
        ["GFH TELECOM - AGED INVENTORY DASHBOARD (>14 Days)", "", "", "", "", "", ""],
        ["", "", "", "", "", "", ""],
        ["TOTAL VALUE AT RISK", "", "TOTAL AGED DEVICES", "", "HIGHEST RISK DISTRICT", "", "HIGHEST RISK STORE"],
        [f"${total_value:,.2f}", "", f"{total_devices}", "", f"{top_district} (${top_district_val:,.2f})", "", f"{top_store} (${top_store_val:,.2f})"],
        ["", "", "", "", "", "", ""],
        ["", "", "", "", "", "", ""],
        ["DISTRICT BREAKDOWN", "", "", "", "TOP 15 STORES AT RISK", "", ""],
        ["District", "Aged Devices", "Total Value ($)", "", "Store Name", "District", "Total Value ($)"]
    ]

    district_summary = df_aged.groupby('_District').agg(
        Devices=(col_serial, 'count'),
        Total_Value=(col_retail, 'sum')
    ).reset_index().sort_values('Total_Value', ascending=False)

    top_15_stores = df_aged.groupby([col_store, '_District']).agg(
        Total_Value=(col_retail, 'sum')
    ).reset_index().sort_values('Total_Value', ascending=False).head(15)

    max_rows = max(len(district_summary), len(top_15_stores))

    for i in range(max_rows):
        row = []
        if i < len(district_summary):
            row.extend([
                district_summary.iloc[i]['_District'],
                int(district_summary.iloc[i]['Devices']),
                f"${float(district_summary.iloc[i]['Total_Value']):,.2f}"
            ])
        else:
            row.extend(["", "", ""])
        
        row.append("")

        if i < len(top_15_stores):
            row.extend([
                top_15_stores.iloc[i][col_store],
                top_15_stores.iloc[i]['_District'],
                f"${float(top_15_stores.iloc[i]['Total_Value']):,.2f}"
            ])
        else:
            row.extend(["", "", ""])

        dashboard_data.append(row)

    ws.update(range_name='A1', values=dashboard_data)

    header_fmt = cellFormat(
        backgroundColor=color(0.08, 0.08, 0.2),
        textFormat=textFormat(bold=True, foregroundColor=color(1, 1, 1), fontSize=14),
        horizontalAlignment='CENTER'
    )
    kpi_title_fmt = cellFormat(
        backgroundColor=color(0.9, 0.9, 0.9),
        textFormat=textFormat(bold=True, fontSize=10, foregroundColor=color(0.3, 0.3, 0.3)),
        horizontalAlignment='CENTER'
    )
    kpi_value_fmt = cellFormat(
        textFormat=textFormat(bold=True, fontSize=16, foregroundColor=color(0.8, 0, 0)),
        horizontalAlignment='CENTER'
    )
    table_header_fmt = cellFormat(
        backgroundColor=color(0.2, 0.2, 0.2),
        textFormat=textFormat(bold=True, foregroundColor=color(1, 1, 1)),
        horizontalAlignment='CENTER'
    )

    formats = [
        ('A1:G1', header_fmt),
        ('A3:B3', kpi_title_fmt), ('C3:D3', kpi_title_fmt), 
        ('E3:F3', kpi_title_fmt), ('G3:H3', kpi_title_fmt),
        ('A4:B4', kpi_value_fmt), ('C4:D4', kpi_value_fmt), 
        ('E4:F4', kpi_value_fmt), ('G4:H4', kpi_value_fmt),
        ('A7:C7', header_fmt), ('E7:G7', header_fmt),
        ('A8:C8', table_header_fmt), ('E8:G8', table_header_fmt)
    ]
    
    format_cell_ranges(ws, formats)

    ws.merge_cells('A1:H1') 
    ws.merge_cells('A3:B3') 
    ws.merge_cells('A4:B4') 
    ws.merge_cells('C3:D3') 
    ws.merge_cells('C4:D4') 
    ws.merge_cells('E3:F3') 
    ws.merge_cells('E4:F4') 
    ws.merge_cells('G3:H3') 
    ws.merge_cells('G4:H4') 

    ws.merge_cells('A7:C7')
    ws.merge_cells('E7:G7')

    set_column_width(ws, 'A', 180)
    set_column_width(ws, 'B', 100)
    set_column_width(ws, 'C', 120)
    set_column_width(ws, 'D', 40)
    set_column_width(ws, 'E', 250)
    set_column_width(ws, 'F', 180)
    set_column_width(ws, 'G', 120)

    return True



def upload_to_google_sheets(district_tabs, aged_df, log):
    gc = authenticate_google(log)
    spreadsheet = open_or_create_spreadsheet(gc, log)
    log(f"Google Sheet ready: {spreadsheet.url}")

    target_tab_names = {sanitize_tab_name(d) for d in district_tabs.keys()}
    target_tab_names.add(DASHBOARD_TAB_NAME)

    log("Building Executive Dashboard...")
    try:
        build_executive_dashboard(spreadsheet, aged_df)
    except Exception as e:
        log(f"Warning: Failed to build dashboard: {e}")

    district_worksheets = []
    for district, tab_df in district_tabs.items():
        tab_name = sanitize_tab_name(district)
        log(f"Uploading '{tab_name}' tab ({len(tab_df)} rows)...")
        values = dataframe_to_values(tab_df)
        ws = get_or_create_worksheet(spreadsheet, tab_name, rows=len(values), cols=len(values[0]))
        clear_and_upload_values(ws, values)
        apply_header_formatting(spreadsheet, ws, len(values), len(values[0]))
        district_worksheets.append(ws)

    try:
        dashboard_ws = spreadsheet.worksheet(DASHBOARD_TAB_NAME)
        ordered = [dashboard_ws] + sorted(district_worksheets, key=lambda w: w.title)
        spreadsheet.reorder_worksheets(ordered)
    except Exception:
        pass

    removed = 0
    for ws in spreadsheet.worksheets():
        if ws.title not in target_tab_names:
            spreadsheet.del_worksheet(ws)
            removed += 1
    if removed:
        log(f"Removed {removed} leftover tab(s) from previous runs.")

    log("Upload complete.")
    return spreadsheet.url




# ==========================================================
# PIPELINE ORCHESTRATORS
# ==========================================================

def run_local_pipeline(excel_path, log, send_email=True):
    """Process Excel locally: save district .xlsx files, optionally email via Outlook."""
    log("=" * 60)
    log("GFH INVENTORY AGING: PROCESS & EMAIL")
    log("=" * 60)

    df, col_store, col_desc, col_serial, col_age, col_po, col_retail = load_and_clean(excel_path, log)

    aged = df[df[col_age] >= MIN_AGE].copy()
    log(f"  {len(aged)} of {len(df)} devices are aged {MIN_AGE}+ days.")

    # District list BEFORE exclusions, so a district whose devices were all
    # excluded still gets its (empty) file and email this run.
    districts = sorted(aged["_District"].dropna().unique())
    log(f"  Districts: {', '.join(districts)}")
    log_exclusions_summary(log)
    aged = apply_output_exclusions(aged, col_serial, col_store, log)
    log("")

    saved = []
    for i, dist in enumerate(districts, 1):
        df_d = aged[aged["_District"] == dist][KEEP_COLUMNS].copy()
        df_d = df_d.sort_values(by="Store", key=lambda s: s.str.lower(),
                                kind="mergesort", ignore_index=True)
        log(f"[{i}/{len(districts)}] Saving: {dist}  ({len(df_d)} rows)")
        path = save_district_xlsx(df_d, dist, Path(excel_path).parent)
        log(f"  Saved: {path.name}")
        saved.append((dist, path))

    if send_email and _HAS_OUTLOOK:
        cfg = load_contacts_config()
        sender_email, cc_full, district_emails = contacts_config_to_legacy(cfg)
        log("")
        log("Sending emails via Outlook Desktop...")
        for dist, fp in saved:
            to = district_emails.get(dist, "")
            if not to:
                log(f"  SKIP {dist} - no email configured (set it in Settings)"); continue
            log(f"  -> {dist}  >  {to}")
            subj = f"GFH Telecom - Inventory Aging Report: {dist} District"
            body = build_html_email(dist, to, sender_cfg=cfg)
            ok = send_via_outlook(to, subj, body, fp, log, sender_email=sender_email, cc_email=cc_full)
            log(f"    {'Sent' if ok else 'FAILED'}")

    log("")
    log(f"COMPLETE - {len(saved)} district file(s) processed")
    return saved


def run_google_pipeline(excel_path, log):
    """Process Excel and upload to Google Sheets."""
    if not _HAS_GSPREAD or not _HAS_GSPREAD_FMT:
        raise PipelineError(
            "Google Sheets upload requires gspread + gspread-formatting.\n"
            "Install with: pip install gspread gspread-formatting")

    log("=" * 60)
    log("GFH INVENTORY AGING: UPLOAD TO GOOGLE SHEETS")
    log("=" * 60)

    df, col_store, col_desc, col_serial, col_age, col_po, col_retail = load_and_clean(excel_path, log)

    log("")
    log(f"Filtering to devices aged over {AGE_THRESHOLD_DAYS} days...")
    aged = build_aged_subset(df, col_age, log)
    pre_excl_districts = sorted(aged["_District"].dropna().unique())

    log("")
    log_exclusions_summary(log)
    aged = apply_output_exclusions(aged, col_serial, col_store, log)

    log("")
    log("Building one tab per district...")
    district_tabs = build_district_tabs(aged, col_store, col_desc, col_serial, col_age, col_po, log,
                                        ensure_districts=pre_excl_districts)

    total_count = len(aged)
    total_loss = float(pd.to_numeric(aged[col_retail], errors="coerce").sum())

    log("")
    sheet_url = upload_to_google_sheets(district_tabs, aged, log)

    log("")
    log("=" * 60)
    log("SUCCESS")
    log(f"Districts uploaded: {len(district_tabs)}")
    log(f"Devices aged >{AGE_THRESHOLD_DAYS} days: {total_count}")
    log(f"Total potential loss: ${total_loss:,.2f}")
    log(f"Google Sheet: {sheet_url}")
    log("=" * 60)

    return sheet_url


# ==========================================================
# GUI
# ==========================================================



def _extract_embedded_icon(b64, filename):
    """Decode an embedded base64 icon to a temp file; return path or None."""
    try:
        if not b64:
            return None
        import base64 as _b64, tempfile, os
        target = os.path.join(tempfile.gettempdir(), filename)
        with open(target, "wb") as fh:
            fh.write(_b64.b64decode(b64))
        return target if os.path.isfile(target) else None
    except Exception:
        return None

def _set_window_icon(root):
    """Set taskbar + titlebar icon from embedded base64 ICO."""
    import base64, tempfile, atexit, os, sys

    # 1. Try sys._MEIPASS (PyInstaller onefile extraction dir)
    meipass = getattr(sys, "_MEIPASS", None)
    if meipass:
        ico_path = os.path.join(meipass, "GFH_icon.ico")
        if os.path.exists(ico_path):
            try:
                root.iconbitmap(default=ico_path)
                root.after(200, lambda p=ico_path: root.iconbitmap(default=p))
                return
            except Exception:
                pass

    # 2. Try next to the exe/script
    if getattr(sys, "frozen", False):
        base_dir = os.path.dirname(sys.executable)
    else:
        base_dir = os.path.dirname(os.path.abspath(__file__))
    ico_path = os.path.join(base_dir, "GFH_icon.ico")
    if os.path.exists(ico_path):
        try:
            root.iconbitmap(default=ico_path)
            root.after(200, lambda p=ico_path: root.iconbitmap(default=p))
            return
        except Exception:
            pass

    # 3. Decode EMBEDDED_ICON_B64 to %TEMP% (no spaces, always writable)
    try:
        data = base64.b64decode(EMBEDDED_ICON_B64.strip())
        tmp_dir = os.environ.get("TEMP", tempfile.gettempdir())
        ico_path = os.path.join(tmp_dir, "gfh_app_icon.ico")
        with open(ico_path, "wb") as f:
            f.write(data)
        root.iconbitmap(default=ico_path)
        root.after(200, lambda p=ico_path: root.iconbitmap(default=p))
        return
    except Exception:
        pass


class SettingsDialog(tk.Toplevel):
    """Lets the user enter the sender identity, CC, and a name+email per
    district. Districts are fully dynamic here — add, rename, or remove
    rows freely; whatever's in the dialog on Save becomes the new district
    list, persisted to CONTACTS_CONFIG_FILE (gfh_aging_contacts.json)."""

    def __init__(self, parent):
        super().__init__(parent)
        self.title("Settings — Sender & District Contacts")
        self.configure(bg=LIGHT)
        self.resizable(False, True)
        self.transient(parent)
        self.grab_set()

        cfg = load_contacts_config()
        self._vars = {}
        self._dist_rows = []  # list of (row_frame, district_name_var, name_var, email_var)

        pad = {"padx": 8, "pady": 4}

        sender_frame = tk.LabelFrame(self, text="Sender identity (used as email signature)",
                                      bg=LIGHT, fg=NAVY, font=("Calibri", 10, "bold"))
        sender_frame.pack(fill="x", padx=12, pady=(12, 6))

        def add_row(frame, row, label, key, default=""):
            tk.Label(frame, text=label, bg=LIGHT, fg=NAVY,
                     font=("Calibri", 9)).grid(row=row, column=0, sticky="w", **pad)
            v = tk.StringVar(value=cfg.get(key, default))
            tk.Entry(frame, textvariable=v, width=40, font=("Calibri", 9)
                     ).grid(row=row, column=1, sticky="ew", **pad)
            self._vars[key] = v

        add_row(sender_frame, 0, "Your name:", "sender_name")
        add_row(sender_frame, 1, "Your email:", "sender_email")
        add_row(sender_frame, 2, "Your title:", "sender_title")
        add_row(sender_frame, 3, "Your mobile:", "sender_mobile")
        add_row(sender_frame, 4, "CC name:", "cc_name")
        add_row(sender_frame, 5, "CC email:", "cc_email")
        sender_frame.columnconfigure(1, weight=1)

        dist_outer = tk.LabelFrame(self, text="District contacts (add/remove districts freely)",
                                    bg=LIGHT, fg=NAVY, font=("Calibri", 10, "bold"))
        dist_outer.pack(fill="both", expand=True, padx=12, pady=(6, 6))

        hdr = tk.Frame(dist_outer, bg=LIGHT); hdr.pack(fill="x", padx=8, pady=(6, 0))
        tk.Label(hdr, text="District", bg=LIGHT, fg="#666", font=("Calibri", 8), width=16, anchor="w").pack(side="left")
        tk.Label(hdr, text="Contact name", bg=LIGHT, fg="#666", font=("Calibri", 8), width=20, anchor="w").pack(side="left")
        tk.Label(hdr, text="Email", bg=LIGHT, fg="#666", font=("Calibri", 8), width=28, anchor="w").pack(side="left")

        self._dist_container = tk.Frame(dist_outer, bg=LIGHT)
        self._dist_container.pack(fill="both", expand=True, padx=8, pady=(2, 6))

        for d, entry in cfg["districts"].items():
            self._add_district_row(d, entry.get("name", ""), entry.get("email", ""))

        ttk.Button(dist_outer, text="+ Add District", style="Browse.TButton",
                   command=lambda: self._add_district_row("", "", "")
                   ).pack(anchor="w", padx=8, pady=(0, 8))

        btn_row = tk.Frame(self, bg=LIGHT)
        btn_row.pack(fill="x", padx=12, pady=(6, 12))
        ttk.Button(btn_row, text="Save", style="Run.TButton",
                   command=self._save).pack(side="right", padx=(6, 0))
        ttk.Button(btn_row, text="Cancel", style="Browse.TButton",
                   command=self.destroy).pack(side="right")

    def _add_district_row(self, district_name, contact_name, contact_email):
        row = tk.Frame(self._dist_container, bg=LIGHT)
        row.pack(fill="x", pady=2)

        dv = tk.StringVar(value=district_name)
        nv = tk.StringVar(value=contact_name)
        ev = tk.StringVar(value=contact_email)

        tk.Entry(row, textvariable=dv, width=16, font=("Calibri", 9)).pack(side="left", padx=(0, 4))
        tk.Entry(row, textvariable=nv, width=20, font=("Calibri", 9)).pack(side="left", padx=(0, 4))
        tk.Entry(row, textvariable=ev, width=28, font=("Calibri", 9)).pack(side="left", padx=(0, 4))

        entry_tuple = (row, dv, nv, ev)

        def remove():
            row.destroy()
            self._dist_rows.remove(entry_tuple)

        tk.Button(row, text="✕", command=remove, bg=LIGHT, fg="#e83030",
                  font=("Calibri", 9, "bold"), bd=0, cursor="hand2").pack(side="left")

        self._dist_rows.append(entry_tuple)

    def _save(self):
        cfg = {k: v.get().strip() for k, v in self._vars.items()}
        districts = {}
        for _row, dv, nv, ev in self._dist_rows:
            d = dv.get().strip()
            if not d:
                continue  # skip rows where the district name was left blank
            districts[d] = {"name": nv.get().strip(), "email": ev.get().strip()}
        cfg["districts"] = districts
        save_contacts_config(cfg)
        messagebox.showinfo("Saved", "Contacts saved to gfh_aging_contacts.json", parent=self)
        self.destroy()


class ExclusionsDialog(tk.Toplevel):
    """Paste-box panel controlling what leaves this app:

      - Stores / IMEIs boxes: hidden from the Google Sheets upload (district
        tabs + Executive Dashboard) and from the district .xlsx files that
        get emailed via Outlook.
      - Always-blocked IMEIs box: removed from BOTH pipelines at cleaning
        time (this box took over the old hardcoded BLOCKED_IMEIS list).

    One entry per line; everything persists to gfh_aging_exclusions.json
    next to the app and survives restarts."""

    _KEYS = ("stores", "imeis", "blocked_imeis")

    def __init__(self, parent):
        super().__init__(parent)
        self.title("Exclusions — Stores & IMEIs")
        self.configure(bg=LIGHT)
        self.resizable(False, True)
        self.transient(parent)
        self.grab_set()

        exc = load_exclusions()
        self._boxes = {}
        self._counts = {}

        spec = [
            ("stores",
             "Stores to exclude (Google Sheet + emailed files)",
             "One store per line. A store is excluded when its name equals the entry "
             "or either side contains the other (case ignored).",
             9),
            ("imeis",
             "IMEIs to exclude (Google Sheet + emailed files)",
             "One IMEI per line. Spaces and dashes are ignored when matching.",
             9),
            ("blocked_imeis",
             "Always-blocked IMEIs (removed from EVERY pipeline at cleaning)",
             "These never appear anywhere — migrated from the old hardcoded block list.",
             7),
        ]

        for key, title, hint, height in spec:
            frame = tk.LabelFrame(self, text=title, bg=LIGHT, fg=NAVY,
                                  font=("Calibri", 10, "bold"))
            frame.pack(fill="both", expand=True, padx=12, pady=(10, 2))
            tk.Label(frame, text=hint, bg=LIGHT, fg="#666",
                     font=("Calibri", 8), wraplength=540, justify="left"
                     ).pack(anchor="w", padx=8, pady=(4, 2))
            box = scrolledtext.ScrolledText(
                frame, height=height, font=("Consolas", 9),
                bg=WHITE, fg=COLOR_TEXT, relief="flat", wrap="word",
                highlightbackground="#b0c4de", highlightthickness=1)
            box.pack(fill="both", expand=True, padx=8, pady=(0, 2))
            box.insert("1.0", "\n".join(exc.get(key, [])))
            cnt = tk.Label(frame, text="", bg=LIGHT, fg=NAVY,
                           font=("Calibri", 8, "bold"))
            cnt.pack(anchor="e", padx=8, pady=(0, 6))
            box.bind("<KeyRelease>",
                     lambda e, k=key, b=box, c=cnt: self._update_count(k, b, c))
            self._boxes[key] = box
            self._counts[key] = cnt
            self._update_count(key, box, cnt)

        btn_row = tk.Frame(self, bg=LIGHT)
        btn_row.pack(fill="x", padx=12, pady=(8, 12))
        ttk.Button(btn_row, text="Save", style="Run.TButton",
                   command=self._save).pack(side="right", padx=(6, 0))
        ttk.Button(btn_row, text="Cancel", style="Browse.TButton",
                   command=self.destroy).pack(side="right")
        ttk.Button(btn_row, text="Clear All", style="Browse.TButton",
                   command=self._clear).pack(side="left")

    @staticmethod
    def _parse(box):
        """Non-empty, de-duplicated (case-insensitive) lines of a paste box."""
        seen, out = set(), []
        for line in box.get("1.0", "end").splitlines():
            v = line.strip()
            if not v:
                continue
            k = v.casefold()
            if k in seen:
                continue
            seen.add(k)
            out.append(v)
        return out

    def _update_count(self, key, box, cnt):
        entries = self._parse(box)
        label = {"stores": "store(s)", "imeis": "IMEI(s)",
                 "blocked_imeis": "blocked IMEI(s)"}[key]
        bad = [e for e in entries
               if key != "stores" and normalize_imei(e)
               and len(normalize_imei(e)) != 15]
        text = f"{len(entries)} {label}"
        if bad:
            text += f"   |   {len(bad)} not 15 digits"
            cnt.configure(fg=RED_DARK)
        else:
            cnt.configure(fg=NAVY)
        cnt.configure(text=text)

    def _clear(self):
        for key, box in self._boxes.items():
            box.delete("1.0", "end")
            self._update_count(key, box, self._counts[key])

    def _save(self):
        exc = {k: self._parse(self._boxes[k]) for k in self._KEYS}
        warnings = []
        for key in ("imeis", "blocked_imeis"):
            bad = [e for e in exc[key] if len(normalize_imei(e)) != 15]
            if bad:
                shown = ", ".join(bad[:6]) + ("..." if len(bad) > 6 else "")
                warnings.append(f"{key.replace('_', ' ')}: {shown}")
        if warnings:
            if not messagebox.askyesno(
                    "Non-15-digit IMEIs",
                    "These entries do not look like 15-digit IMEIs:\n\n"
                    + "\n".join(warnings)
                    + "\n\nThey are still saved and compared digits-only.\n\nSave anyway?",
                    parent=self):
                return
        save_exclusions(exc)
        messagebox.showinfo("Saved", "Exclusions saved to gfh_aging_exclusions.json",
                            parent=self)
        self.destroy()


def _force_maximized(window):
    """Always open the window MAXIMIZED (user request) and KEEP it maximized
    while the window maps. Tk silently drops the zoomed state when it is set
    while the window is still being mapped — the zoomed state is therefore
    re-asserted several times across the first second."""
    def _zoom():
        try:
            window.state("zoomed")
        except Exception:
            try:
                window.attributes("-zoomed", True)
            except Exception:
                pass
    for _delay in (10, 80, 250, 600):
        try:
            window.after(_delay, _zoom)
        except Exception:
            pass


class App:
    def __init__(self, root):
        self.root = root
        self.file_path = tk.StringVar()
        self._q = queue.Queue()
        self._running = False
        self.sheet_url_cache = None
        self._wordmark_img = None

        root.title(APP_TITLE)
        # Dynamic screen resolution support: size to 90% of the screen and
        # center it (DPI-aware), then stay a normal resizable top-level so
        # Windows Snap (50% left/right, corners, Win+arrow) keeps working.
        self._apply_dynamic_geometry()
        _force_maximized(self.root)  # immediate zoom + re-asserts while Tk maps
        root.configure(bg=LIGHT)
        _set_window_icon(root)
        root.protocol("WM_DELETE_WINDOW", root.destroy)
        self.theme_manager = ThemeManager("GFH Inventory Aging Processor", app_name="gfh-inventory-aging-processor")
        self._styles(); self._header(); self._tabs(); self._copyright_bar()
        apply_theme_to_window(self.root, self.theme_manager)
        self._poll()

    def _tabs(self):
        """Email/Google-Sheets pipeline (tab 1, unchanged) + WhatsApp — Aged
        Devices (tab 2, new): two independent tools in one app."""
        nb = ttk.Style()
        nb.configure("GFH.TNotebook", background=LIGHT, borderwidth=0)
        nb.configure("GFH.TNotebook.Tab", font=("Calibri", 10, "bold"),
                     padding=(16, 9), background="#dfe6f2", foreground=NAVY)
        nb.map("GFH.TNotebook.Tab",
               background=[("selected", NAVY)], foreground=[("selected", WHITE)])

        self.notebook = ttk.Notebook(self.root, style="GFH.TNotebook")
        self.notebook.pack(fill="both", expand=True)

        tab1 = tk.Frame(self.notebook, bg=LIGHT)
        tab2 = tk.Frame(self.notebook, bg=LIGHT)
        self.notebook.add(tab1, text="Email / Google Sheets")
        self.notebook.add(tab2, text="WhatsApp — Aged Devices")

        self._body(tab1)
        self._whatsapp_tab(tab2)

    def _apply_dynamic_geometry(self) -> None:
        """Fill the entire work area (screen minus taskbar) BEFORE the window
        is ever mapped, so the very first paint already looks maximized — the
        old 90%-centered geometry plus a delayed state('zoomed') made the exe
        visibly open centered and then jump to full screen.

        Windows: the real work area is read via SystemParametersInfoW so the
        taskbar is respected on any monitor/DPI. Other platforms fall back to
        the full screen size. The window stays resizable so Windows Snap keeps
        working, and state('zoomed') is still applied for a true maximize."""
        try:
            root = self.root
            root.update_idletasks()
            sw, sh = root.winfo_screenwidth(), root.winfo_screenheight()
            w, h, x, y = sw, sh, 0, 0
            try:
                import ctypes

                class _WORKAREA_RECT(ctypes.Structure):
                    _fields_ = [("left", ctypes.c_long), ("top", ctypes.c_long),
                                ("right", ctypes.c_long), ("bottom", ctypes.c_long)]

                _rect = _WORKAREA_RECT()
                # 0x0030 = SPI_GETWORKAREA
                if ctypes.windll.user32.SystemParametersInfoW(0x0030, 0, ctypes.byref(_rect), 0):
                    w = _rect.right - _rect.left
                    h = _rect.bottom - _rect.top
                    x = _rect.left
                    y = _rect.top
            except Exception:
                h = max(0, sh - 48)  # taskbar fallback if the Win32 call fails
            root.geometry(f"{max(640, min(w, sw))}x{max(480, min(h, sh))}+{max(0, x)}+{max(0, y)}")
            root.minsize(min(1220, max(560, sw // 2)),
                         min(740, max(420, sh // 2)))
            root.resizable(True, True)
        except Exception:
            pass

    def _styles(self):
        s = ttk.Style(); s.theme_use("clam")
        s.configure("Run.TButton", background=RED, foreground=WHITE,
                    font=("Calibri", 11, "bold"), padding=(16, 9), borderwidth=0)
        s.map("Run.TButton", background=[("active", RED_DARK), ("disabled", "#aaa")])
        s.configure("Browse.TButton", background=NAVY, foreground=WHITE,
                    font=("Calibri", 10), padding=(10, 6), borderwidth=0)
        s.map("Browse.TButton", background=[("active", "#1a2550")])
        s.configure("Link.TButton", background=NAVY, foreground=WHITE,
                    font=("Calibri", 10), padding=(10, 6), borderwidth=0)
        s.map("Link.TButton", background=[("active", "#1a2550")])
        s.configure("Accent.Horizontal.TProgressbar",
                    troughcolor="#dde6f0", background=RED, borderwidth=0)


    def _extract_embedded(self, b64, filename):
        """Decode an embedded base64 asset into a temp file; return path or None."""
        try:
            if not b64:
                return None
            import base64 as _b64, tempfile, os
            target = os.path.join(tempfile.gettempdir(), filename)
            with open(target, "wb") as fh:
                fh.write(_b64.b64decode(b64))
            return target if os.path.isfile(target) else None
        except Exception:
            return None


    def _lock_header_colors(self, widget, navy):
        """Recursively bind <Enter>/<Leave> on all header widgets to force navy."""
        try:
            widget.bind("<Enter>", lambda e, w=widget, c=navy: w.configure(bg=c) if not isinstance(w, type(None)) else None)
            widget.bind("<Leave>", lambda e, w=widget, c=navy: w.configure(bg=c) if not isinstance(w, type(None)) else None)
        except Exception:
            pass
        try:
            for child in widget.winfo_children():
                self._lock_header_colors(child, navy)
        except Exception:
            pass
    def _header(self):
        self.header_mgr = FixedHeaderManager(self.root, title="GFH Inventory Aging Processor")
        self.header_mgr.add_theme_toggle(self.theme_manager, callback=self._apply_theme)
        # header_manager already tags all header widgets at init — no override needed
        try:
            _lp = _resource_path(LOGO_PNG_NAME) if "_resource_path" in dir() else os.path.join(os.path.dirname(os.path.abspath(__file__)), LOGO_PNG_NAME)
            if os.path.exists(_lp):
                self.header_mgr.set_logo(logo_path=_lp, text="GFH")
        except Exception:
            pass


    def _apply_theme(self, colors=None):
        """Apply theme colors to all widgets."""
        if colors is None:
            colors = self.theme_manager.get_colors()
        apply_theme_to_window(self.root, self.theme_manager)
        try:
            self.root.configure(bg=colors.get("bg", "#f6f7fb"))
        except Exception:
            pass
        self._wa_theme_fixups(colors)

    def _wa_theme_fixups(self, colors=None):
        """Re-color the WhatsApp tab's "brand"-tagged widgets (the ones the
        shared theme walker is told to skip) by hand, since they need a
        readonlybackground the walker doesn't know to touch. Safe to call
        before the tab exists yet."""
        entry = getattr(self, "wa_file_entry", None)
        if entry is None:
            return
        if colors is None:
            colors = self.theme_manager.get_colors()
        bg = colors.get("input", "#e8eff8")
        fg = colors.get("text", NAVY)
        try:
            entry.configure(bg=bg, fg=fg, readonlybackground=bg, insertbackground=fg)
        except Exception:
            pass
    def _body(self, parent):
        body = tk.Frame(parent, bg=LIGHT)
        body.pack(fill="both", expand=True, padx=24, pady=18)

        # Drop zone
        outer = tk.Frame(body, bg="#b8cce4", bd=0)
        outer.pack(fill="x", pady=(0, 14))
        self.drop_frame = tk.Frame(outer, bg=WHITE); self.drop_frame.pack(fill="x", padx=2, pady=2)
        dnd_note = "" if HAS_DND else "\n(install tkinterdnd2 for drag-and-drop)"
        self._dlbl = tk.Label(self.drop_frame,
            text=f"Drag & Drop  Inventory_Aging.xlsx  here\nor click Browse to select{dnd_note}",
            font=("Calibri", 11), fg="#4a6080", bg=WHITE, pady=28, cursor="hand2")
        self._dlbl.pack(fill="x")
        self._dlbl.bind("<Button-1>", lambda e: self._browse())
        if HAS_DND:
            for w in (self.drop_frame, self._dlbl):
                w.drop_target_register(DND_FILES)
                w.dnd_bind("<<Drop>>", self._on_drop)

        # File row
        row = tk.Frame(body, bg=LIGHT); row.pack(fill="x", pady=(0, 14))
        row.columnconfigure(0, weight=1)
        tk.Entry(row, textvariable=self.file_path, state="readonly",
                 font=("Calibri", 9), relief="flat", bg="#e8eff8", fg=NAVY,
                 readonlybackground="#e8eff8",
                 highlightbackground="#b0c4de", highlightthickness=1
                 ).grid(row=0, column=0, sticky="ew", ipady=5, padx=(0, 8))
        ttk.Button(row, text="Browse", style="Browse.TButton",
                   command=self._browse).grid(row=0, column=1)

        # Options
        opt = tk.Frame(body, bg=LIGHT); opt.pack(fill="x", pady=(0, 14))
        self.send_var = tk.BooleanVar(value=True)
        tk.Checkbutton(opt, text="Send emails via Outlook Desktop after processing",
                       variable=self.send_var, bg=LIGHT, fg=NAVY,
                       activebackground=LIGHT, selectcolor=WHITE,
                       font=("Calibri", 10)).pack(side="left")

        # Action buttons
        btn_row = tk.Frame(body, bg=LIGHT); btn_row.pack(fill="x", pady=(0, 12))
        self.run_btn = ttk.Button(btn_row, text="Process & Email Reports",
                                  style="Run.TButton", command=self._start_local)
        self.run_btn.pack(side="left", fill="x", expand=True, padx=(0, 4))

        gs_state = "!disabled" if (_HAS_GSPREAD and _HAS_GSPREAD_FMT) else "disabled"
        self.upload_btn = ttk.Button(btn_row, text="Upload to Google Sheets",
                                     style="Run.TButton", command=self._start_google,
                                     state=gs_state)
        self.upload_btn.pack(side="left", fill="x", expand=True, padx=(4, 4))

        self.link_btn = ttk.Button(btn_row, text="Open Google Sheet",
                                   style="Link.TButton", command=self._open_sheet,
                                   state="disabled")
        self.link_btn.pack(side="left", fill="x", expand=True, padx=(4, 0))

        self.settings_btn = ttk.Button(btn_row, text="Settings (Contacts)",
                                        style="Browse.TButton", command=self._open_settings)
        self.settings_btn.pack(side="left", fill="x", expand=True, padx=(4, 0))

        self.excl_btn = ttk.Button(btn_row, text="Exclusions",
                                    style="Browse.TButton", command=self._open_exclusions)
        self.excl_btn.pack(side="left", fill="x", expand=True, padx=(4, 0))

        # Progress
        self.progress = ttk.Progressbar(body, mode="indeterminate",
                                      style="Accent.Horizontal.TProgressbar")
        self.progress.pack(fill="x", pady=(0, 10))

        # Log
        tk.Label(body, text="Activity Log", font=("Calibri", 9, "bold"),
                 fg=NAVY, bg=LIGHT).pack(anchor="w")
        self.log_box = scrolledtext.ScrolledText(
            body, height=11, font=("Consolas", 8),
            bg="#10182e", fg="#a8d8ff", relief="flat", state="disabled", wrap="word")
        self.log_box.pack(fill="both", expand=True)

    def _open_settings(self):
        SettingsDialog(self.root)

    def _open_exclusions(self):
        ExclusionsDialog(self.root)

    # ------------------------------------------------------------------
    # TAB 2: WhatsApp — Aged Devices
    # ------------------------------------------------------------------
    def _whatsapp_tab(self, parent):
        """Pick a raw aging export, map District -> WhatsApp group by hand,
        tick which districts to send, type the aged-days cutoff, and push
        one GFH-branded table image per ticked district into WhatsApp
        Desktop. Everything (mapping, checks, aged-days, excluded stores)
        persists to gfh_aging_whatsapp_config.json for next time."""
        self.wa_file_path = tk.StringVar()
        self.wa_cfg = load_whatsapp_config()
        self.wa_rows = []
        self.wa_districts = []
        self.wa_row_widgets = {}  # district -> (BooleanVar checked, StringVar group)
        self._wa_running = False
        self._wa_q = queue.Queue()

        body = tk.Frame(parent, bg=LIGHT)
        body.pack(fill="both", expand=True, padx=24, pady=18)

        # File row
        frow = tk.Frame(body, bg=LIGHT); frow.pack(fill="x", pady=(0, 10))
        frow.columnconfigure(0, weight=1)
        self.wa_file_entry = tk.Entry(
            frow, textvariable=self.wa_file_path, state="readonly",
            font=("Calibri", 9), relief="flat",
            highlightbackground="#b0c4de", highlightthickness=1)
        self.wa_file_entry.grid(row=0, column=0, sticky="ew", ipady=5, padx=(0, 8))
        # The shared theme walker recolors Entry widgets' fg to the current
        # theme's text color but only knows to recolor plain bg, not
        # readonlybackground (what a readonly Entry actually displays) - a
        # fixed/hardcoded readonlybackground would go invisible against a
        # fg color that changes with the theme. Tag it "brand" (an existing
        # protected tag) so the walker leaves it alone entirely, and we set
        # BOTH colors ourselves instead - see _wa_theme_fixups(), called
        # here and again on every theme toggle via _apply_theme().
        self.wa_file_entry._tag = "brand"
        self._wa_theme_fixups()
        ttk.Button(frow, text="Browse", style="Browse.TButton",
                   command=self._wa_browse).grid(row=0, column=1, padx=(0, 8))
        self.wa_load_btn = ttk.Button(frow, text="Load Districts", style="Run.TButton",
                                       command=self._wa_load)
        self.wa_load_btn.grid(row=0, column=2)

        # Aged-days + store-exclusions row
        opt = tk.Frame(body, bg=LIGHT); opt.pack(fill="x", pady=(0, 10))
        tk.Label(opt, text="Aged (days) or more:", bg=LIGHT, fg=NAVY,
                 font=("Calibri", 10, "bold")).pack(side="left")
        self.wa_aged_var = tk.IntVar(value=int(self.wa_cfg.get("aged_days", 30)))
        tk.Spinbox(opt, from_=0, to=9999, width=6, textvariable=self.wa_aged_var,
                   font=("Calibri", 10), justify="center"
                   ).pack(side="left", padx=(8, 20))
        ttk.Button(opt, text="Store Exclusions", style="Browse.TButton",
                   command=self._wa_open_store_exclusions).pack(side="left")
        tk.Label(opt, text="  Tick a district, type its WhatsApp group name, then Save.",
                 bg=LIGHT, fg="#666", font=("Calibri", 9)).pack(side="left")

        msg_row = tk.Frame(body, bg=LIGHT); msg_row.pack(fill="x", pady=(0, 10))
        tk.Label(msg_row, text="Message sent with each image ({days} = the Aged-days box above):",
                 bg=LIGHT, fg=NAVY, font=("Calibri", 9, "bold")).pack(anchor="w")
        self.wa_message_var = tk.StringVar(value=self.wa_cfg.get(
            "message_template", WA_DEFAULT_MESSAGE))
        tk.Entry(msg_row, textvariable=self.wa_message_var, font=("Calibri", 10),
                 relief="flat", highlightbackground="#b0c4de", highlightthickness=1
                 ).pack(fill="x", ipady=4, pady=(2, 0))

        # Scrollable districts panel
        tk.Label(body, text="Districts", font=("Calibri", 9, "bold"),
                 fg=NAVY, bg=LIGHT).pack(anchor="w")
        panel_outer = tk.Frame(body, bg=WHITE, highlightbackground="#b0c4de",
                                highlightthickness=1)
        panel_outer.pack(fill="both", expand=True, pady=(2, 10))
        canvas = tk.Canvas(panel_outer, bg=WHITE, highlightthickness=0)
        vbar = ttk.Scrollbar(panel_outer, orient="vertical", command=canvas.yview)
        self.wa_districts_frame = tk.Frame(canvas, bg=WHITE)
        self.wa_districts_frame.bind(
            "<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=self.wa_districts_frame, anchor="nw")
        canvas.configure(yscrollcommand=vbar.set)
        canvas.pack(side="left", fill="both", expand=True)
        vbar.pack(side="right", fill="y")
        self.wa_canvas = canvas

        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        # bind only while hovering this panel so it doesn't steal scrolling
        # from the rest of the app when this tab isn't visible
        panel_outer.bind("<Enter>", lambda e: canvas.bind_all("<MouseWheel>", _on_mousewheel))
        panel_outer.bind("<Leave>", lambda e: canvas.unbind_all("<MouseWheel>"))

        self.wa_empty_label = tk.Label(
            self.wa_districts_frame,
            text="Select a file above and click \"Load Districts\".",
            bg=WHITE, fg="#888", font=("Calibri", 10), pady=16)
        self.wa_empty_label.pack(fill="x")
        # These widgets are created AFTER the one apply_theme_to_window() call
        # in __init__, so they'd otherwise be stuck at their WHITE
        # creation-time color forever in dark mode. Re-run it now (and again
        # whenever the district rows are rebuilt) so they pick up whichever
        # theme is actually active, exactly like every other widget does.
        self._apply_theme()

        # Buttons
        btn_row = tk.Frame(body, bg=LIGHT); btn_row.pack(fill="x", pady=(0, 10))
        ttk.Button(btn_row, text="Save Mapping", style="Browse.TButton",
                   command=self._wa_save_mapping).pack(side="left", padx=(0, 8))
        self.wa_send_btn = ttk.Button(btn_row, text="Send to WhatsApp Desktop",
                                       style="Run.TButton", command=self._wa_send,
                                       state="disabled")
        self.wa_send_btn.pack(side="left", fill="x", expand=True)
        tk.Label(body, text="Tip: tick only ONE district and send it first to confirm it "
                            "opens the right group before sending to everyone.",
                 bg=LIGHT, fg="#b45309", font=("Calibri", 8, "italic")
                 ).pack(anchor="w", pady=(2, 0))

        self.wa_progress = ttk.Progressbar(body, mode="indeterminate",
                                            style="Accent.Horizontal.TProgressbar")
        self.wa_progress.pack(fill="x", pady=(0, 10))

        tk.Label(body, text="Activity Log", font=("Calibri", 9, "bold"),
                 fg=NAVY, bg=LIGHT).pack(anchor="w")
        self.wa_log_box = scrolledtext.ScrolledText(
            body, height=9, font=("Consolas", 8),
            bg="#10182e", fg="#a8d8ff", relief="flat", state="disabled", wrap="word")
        self.wa_log_box.pack(fill="both", expand=True)

        if not _HAS_WA_AUTOMATION_DEPS:
            self._wa_log("NOTE: pyautogui not installed — install it to enable "
                         "one-click sending (pip install pyautogui). The image "
                         "will still be generated and copied to the clipboard.")
        n = len(self.wa_cfg.get("excluded_stores", []))
        if n:
            self._wa_log(f"{n} store(s) excluded from sends — see Store Exclusions.")

        # Widgets above were built before any theme re-walk could reach them
        # (the only one so far ran right after the file-path entry, earlier
        # in this method) - run it once more now that the whole tab exists,
        # so it looks correct immediately rather than only after Load.
        self._apply_theme()

    def _wa_log(self, msg):
        """Thread-safe: always goes through the queue _poll() drains, even
        when called from the send worker thread."""
        self._wa_q.put(msg)

    def _wa_browse(self):
        p = filedialog.askopenfilename(title="Select Inventory Aging Excel file",
            filetypes=[("Excel files", "*.xlsx *.xlsm *.xls"), ("All files", "*.*")])
        if p:
            self.wa_file_path.set(p)
            self._wa_log(f"Selected: {Path(p).name}")

    def _wa_load(self):
        p = self.wa_file_path.get().strip()
        if not p or not Path(p).exists():
            self._wa_log("ERROR: Please select a valid Excel file first.")
            return
        try:
            self.wa_rows = wa_read_aging_rows(p)
        except Exception as e:
            self._wa_log(f"ERROR: {e}")
            return
        self.wa_districts = wa_unique_districts(self.wa_rows)
        if not self.wa_districts:
            self._wa_log("No District column values found in this file.")
            return
        self.wa_cfg["last_file"] = p
        self._wa_build_district_rows()
        self.wa_send_btn.config(state="normal")
        self._wa_log(f"Loaded {len(self.wa_rows)} rows across {len(self.wa_districts)} "
                     f"district(s): {', '.join(self.wa_districts)}")

    def _wa_build_district_rows(self):
        for child in self.wa_districts_frame.winfo_children():
            child.destroy()
        self.wa_row_widgets = {}
        saved = self.wa_cfg.get("districts", {})
        # Seed with the CURRENT theme's colors (not hardcoded light ones) so
        # a fresh Load looks right immediately. Deliberately NOT tagged
        # "brand": tagging a Frame also stops the shared theme walker from
        # ever recursing into its children (that was the actual bug - the
        # WhatsApp-group Entry fields stayed stuck light because the row
        # Frame around them was tagged and silently blocked the walker from
        # reaching them at all). Leaving everything untagged means toggling
        # the theme later re-colors these rows correctly too, same as any
        # other widget in the app - the one trade-off is the alternating
        # row stripe flattens to one color after a toggle, which is a minor
        # cosmetic cost worth paying for colors that are always correct.
        colors = self.theme_manager.get_colors()
        panel = colors.get("panel", WHITE)
        panel_alt = colors.get("panel_alt", "#eef2f8")
        text_fg = colors.get("text", NAVY)
        for i, d in enumerate(self.wa_districts):
            entry = saved.get(d, {})
            stripe = panel if i % 2 == 0 else panel_alt
            row = tk.Frame(self.wa_districts_frame, bg=stripe)
            row.pack(fill="x")
            chk_var = tk.BooleanVar(value=bool(entry.get("checked", False)))
            tk.Checkbutton(row, variable=chk_var, bg=stripe, activebackground=stripe
                           ).pack(side="left", padx=(10, 6), pady=6)
            tk.Label(row, text=d, bg=stripe, fg=text_fg, font=("Calibri", 10, "bold"),
                     width=18, anchor="w").pack(side="left", padx=(0, 10))
            tk.Label(row, text="WhatsApp group:", bg=stripe, fg=text_fg,
                     font=("Calibri", 9)).pack(side="left", padx=(0, 6))
            grp_var = tk.StringVar(value=entry.get("group", ""))
            tk.Entry(row, textvariable=grp_var, font=("Calibri", 10),
                     relief="flat", bg=colors.get("input", "#e8eff8"), fg=text_fg,
                     insertbackground=text_fg,
                     highlightbackground="#b0c4de", highlightthickness=1
                     ).pack(side="left", fill="x", expand=True, padx=(0, 12), ipady=3)
            self.wa_row_widgets[d] = (chk_var, grp_var)
        self._apply_theme()

    def _wa_open_store_exclusions(self):
        colors = self.theme_manager.get_colors()
        bg = colors.get("bg", LIGHT)
        text_fg = colors.get("text", COLOR_TEXT)
        input_bg = colors.get("input", WHITE)

        win = tk.Toplevel(self.root)
        win.title("Store Exclusions (WhatsApp tab)")
        win.configure(bg=bg)
        win.geometry("520x520")
        win.minsize(420, 360)
        win.transient(self.root)
        tk.Label(win, text="Stores listed here are left out of every WhatsApp send "
                           "(this tab only — the Email/Sheets Exclusions panel is separate).",
                 bg=bg, fg=text_fg, font=("Calibri", 9), wraplength=480,
                 justify="left").pack(anchor="w", padx=10, pady=(10, 4))
        tk.Label(win, text="One store name per line. Matches case-insensitively; "
                           "partial names match too (e.g. \"Houston\" matches every "
                           "store containing that word).",
                 bg=bg, fg="#888", font=("Calibri", 8), wraplength=480,
                 justify="left").pack(anchor="w", padx=10, pady=(0, 6))

        # Button row packed FIRST at the bottom, so it always keeps its space
        # regardless of the ScrolledText's size - previously it was packed
        # after an expand=True ScrolledText with no height limit, which could
        # claim the whole window and push Save/Cancel out of view.
        btns = tk.Frame(win, bg=bg)
        btns.pack(side="bottom", fill="x", padx=10, pady=10)

        box = scrolledtext.ScrolledText(win, font=("Consolas", 9), bg=input_bg,
                                        fg=text_fg, insertbackground=text_fg,
                                        relief="flat", wrap="word", height=14,
                                        highlightbackground="#b0c4de", highlightthickness=1)
        box.pack(fill="both", expand=True, padx=10, pady=(0, 10))
        box.insert("1.0", "\n".join(self.wa_cfg.get("excluded_stores", [])))

        def _save():
            lines = [l.strip() for l in box.get("1.0", "end").splitlines()]
            seen, out = set(), []
            for l in lines:
                if not l:
                    continue
                key = l.casefold()
                if key in seen:
                    continue
                seen.add(key)
                out.append(l)
            self.wa_cfg["excluded_stores"] = out
            save_whatsapp_config(self._wa_collect_config())
            self._wa_log(f"Store Exclusions saved: {len(out)} store(s) excluded.")
            win.destroy()

        ttk.Button(btns, text="Save", style="Run.TButton", command=_save
                   ).pack(side="right", padx=(8, 0))
        ttk.Button(btns, text="Cancel", style="Browse.TButton", command=win.destroy
                   ).pack(side="right")

    def _wa_collect_config(self):
        cfg = {
            "aged_days": int(self.wa_aged_var.get() or 0),
            "last_file": self.wa_cfg.get("last_file", ""),
            "districts": {},
            "excluded_stores": list(self.wa_cfg.get("excluded_stores", [])),
            "message_template": self.wa_message_var.get().strip() or WA_DEFAULT_MESSAGE,
        }
        for d, (chk_var, grp_var) in self.wa_row_widgets.items():
            cfg["districts"][d] = {"checked": bool(chk_var.get()),
                                   "group": grp_var.get().strip()}
        return cfg

    def _wa_save_mapping(self):
        cfg = self._wa_collect_config()
        save_whatsapp_config(cfg)
        self.wa_cfg = cfg
        self._wa_log("Saved group mapping + aged-days to gfh_aging_whatsapp_config.json.")

    def _wa_send(self):
        if self._wa_running:
            return
        if not self.wa_rows:
            self._wa_log("ERROR: Load a file first.")
            return
        self._wa_save_mapping()
        targets = [(d, gv.get().strip()) for d, (cv, gv) in self.wa_row_widgets.items()
                   if cv.get() and gv.get().strip()]
        if not targets:
            self._wa_log("Nothing to send: tick at least one district and type its group name.")
            return
        self._wa_running = True
        self.wa_send_btn.config(state="disabled")
        self.wa_progress.start(12)
        aged_days = int(self.wa_aged_var.get() or 0)
        excluded_stores = list(self.wa_cfg.get("excluded_stores", []))
        template = self.wa_message_var.get().strip() or WA_DEFAULT_MESSAGE
        try:
            caption = template.format(days=aged_days)
        except Exception:
            caption = template  # bad {placeholder} in a hand-edited template - send as-is
        self._wa_log("-" * 55)
        self._wa_log(f"Sending {len(targets)} district(s), aged >= {aged_days} days"
                     f"{f' (excluding {len(excluded_stores)} store pattern(s))' if excluded_stores else ''}...")
        threading.Thread(target=self._wa_send_worker,
                         args=(targets, aged_days, excluded_stores, caption),
                         daemon=True).start()

    def _wa_send_worker(self, targets, aged_days, excluded_stores, caption):
        logo_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), LOGO_PNG_NAME)
        for district, group in targets:
            rows = wa_filter_rows(self.wa_rows, district, aged_days, excluded_stores)
            if not rows:
                self._wa_log(f"{district}: 0 device(s) aged >= {aged_days} days after "
                             f"exclusions — skipped.")
                continue
            try:
                img = wa_render_table_image(district, rows, logo_path=logo_path, group_label=group)
            except Exception as e:
                self._wa_log(f"{district}: ERROR building image: {e}")
                continue
            self._wa_log(f"{district}: {len(rows)} device(s) -> '{group}' — sending...")
            try:
                ok, msg = send_image_to_whatsapp_group(
                    group, img, log=self._wa_log, caption=caption)
            except Exception as e:
                ok, msg = False, str(e)
            self._wa_log(f"{district}: {'OK' if ok else 'FAILED'} — {msg}")
        self.root.after(0, self._wa_send_done)

    def _wa_send_done(self):
        self.wa_progress.stop()
        self.wa_send_btn.config(state="normal")
        self._wa_running = False
        self._wa_log("-" * 55)

    def _copyright_bar(self):
        bar = tk.Frame(self.root, bg=NAVY, height=26)
        bar.pack(fill="x", side="bottom"); bar.pack_propagate(False)
        tk.Label(bar, text=COPYRIGHT_TEXT, bg=NAVY, fg="#9d9db8",
                 font=("Calibri", 8)).pack(pady=4)

    def _on_drop(self, event):
        p = event.data.strip().strip("{}")
        if p.lower().endswith((".xlsx", ".xlsm", ".xls")):
            self.file_path.set(p); self._dlbl.config(fg="#1a6630")
            self._log(f"File dropped: {Path(p).name}")
        else:
            self._log("Please drop an Excel (.xlsx) file.")

    def _browse(self):
        p = filedialog.askopenfilename(title="Select Inventory Aging Excel file",
            filetypes=[("Excel files", "*.xlsx *.xlsm *.xls"), ("All files", "*.*")])
        if p: self.file_path.set(p); self._log(f"Selected: {Path(p).name}")

    def _log(self, msg):
        self._q.put(msg)

    def _poll(self):
        try:
            while True:
                msg = self._q.get_nowait()
                self.log_box.config(state="normal")
                self.log_box.insert("end", msg + "\n"); self.log_box.see("end")
                self.log_box.config(state="disabled")
        except queue.Empty:
            pass
        # WhatsApp tab: separate queue so its background send-thread never
        # touches the Email/Sheets tab's widgets (and vice versa).
        try:
            while True:
                msg = self._wa_q.get_nowait()
                self.wa_log_box.config(state="normal")
                self.wa_log_box.insert("end", msg + "\n"); self.wa_log_box.see("end")
                self.wa_log_box.config(state="disabled")
        except queue.Empty:
            pass
        self.root.after(80, self._poll)

    def _start_local(self):
        if self._running: return
        p = self.file_path.get().strip()
        if not p or not Path(p).exists():
            self._log("ERROR: Please select a valid Excel file first."); return
        self._running = True
        self.run_btn.config(state="disabled"); self.upload_btn.config(state="disabled")
        self.progress.start(12)
        self._log("-" * 55)
        _cfg = load_contacts_config()
        _sender_email, _cc_full, _ = contacts_config_to_legacy(_cfg)
        self._log(f"Input : {Path(p).name}")
        self._log(f"Sender: {_sender_email or '(not set — see Settings)'}  |  CC: {_cc_full or '(none)'}")
        self._log(f"Emails: {'YES' if self.send_var.get() else 'NO (disabled)'}")
        self._log("-" * 55)
        threading.Thread(target=self._worker_local,
                         args=(Path(p), self.send_var.get()), daemon=True).start()

    def _worker_local(self, input_path, send_email):
        try:
            run_local_pipeline(input_path, self._log, send_email=send_email)
            self.root.after(0, self._done)
        except Exception as exc:
            self._log(f"ERROR: {exc}"); self._log(traceback.format_exc())
            self.root.after(0, self._done)

    def _start_google(self):
        if self._running: return
        p = self.file_path.get().strip()
        if not p or not Path(p).exists():
            self._log("ERROR: Please select a valid Excel file first."); return
        self._running = True
        self.run_btn.config(state="disabled"); self.upload_btn.config(state="disabled")
        self.progress.start(12)
        self._log("-" * 55)
        self._log(f"Input: {Path(p).name}")
        self._log("Uploading to Google Sheets...")
        self._log("-" * 55)
        threading.Thread(target=self._worker_google, args=(p,), daemon=True).start()

    def _worker_google(self, excel_path):
        try:
            url = run_google_pipeline(excel_path, self._log)
            self.sheet_url_cache = url
            self.root.after(0, self._done_google, url)
        except PipelineError as e:
            self._log(f"ERROR: {e}")
            self.root.after(0, self._done)
        except Exception as e:
            self._log(f"ERROR: {e}"); self._log(traceback.format_exc())
            self.root.after(0, self._done)

    def _done(self):
        self.progress.stop()
        self.run_btn.config(state="normal")
        gs_state = "normal" if (_HAS_GSPREAD and _HAS_GSPREAD_FMT) else "disabled"
        self.upload_btn.config(state=gs_state)
        self._running = False

    def _done_google(self, url):
        self._done()
        self.link_btn.config(state="normal")
        messagebox.showinfo(APP_TITLE, f"Upload complete!\n\nGoogle Sheet:\n{url}")

    def _open_sheet(self):
        if self.sheet_url_cache:
            webbrowser.open(self.sheet_url_cache)


# ==========================================================
# ENTRY POINT
# ==========================================================
def _enable_dpi_awareness() -> None:
    """Make Windows report physical pixels so winfo_screen* is accurate on
    high-DPI displays (1080p, 1440p, 2K, 4K, DPI-scaled laptops)."""
    if sys.platform != "win32":
        return
    try:
        import ctypes
        try:
            ctypes.windll.shcore.SetProcessDpiAwareness(1)  # system DPI aware
        except Exception:
            ctypes.windll.user32.SetProcessDPIAware()
    except Exception:
        pass


def main():
    _enable_dpi_awareness()
    # Must be before the first tk.Tk() or Windows ignores it
    try:
        import ctypes
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("GFHTelecom.InventoryAgingProcessor")
    except Exception:
        pass
    parser = argparse.ArgumentParser()
    parser.add_argument("--input", default=None)
    parser.add_argument("--no-email", action="store_true")
    parser.add_argument("--no-upload", action="store_true")
    args = parser.parse_args()

    if args.input:
        input_path = Path(args.input)

        saved = run_local_pipeline(input_path, print, send_email=not args.no_email)
        for dist, fp in saved:
            print(f"Saved: {fp.name}")

        if not args.no_upload and _HAS_GSPREAD and _HAS_GSPREAD_FMT:
            url = run_google_pipeline(str(input_path), print)
            print(f"Google Sheet: {url}")
        return

    root = (TkinterDnD.Tk() if HAS_DND else tk.Tk())
    _vp_hide(root)
    App(root); _vp_reveal(root); _close_pyi_splash(); root.mainloop()


def _vp_hide(win):
    """ANTI-FLASH: keep the window fully transparent while the UI is built, so a
    half-built window never flashes behind the boot splash."""
    try:
        win.attributes("-alpha", 0.0)
    except Exception:
        pass
    try:
        import pyi_splash  # type: ignore
        pyi_splash.update_text("Loading interface...")
    except Exception:
        pass


def _vp_reveal(win):
    """Show the finished window in one step (called just before the splash closes)."""
    try:
        win.update_idletasks()
    except Exception:
        pass
    try:
        win.attributes("-alpha", 1.0)
        win.lift()
    except Exception:
        pass


def _close_pyi_splash():
    """Close the PyInstaller splash screen (frozen onefile builds only).

    Called once the main window is fully built and shown; harmless no-op
    in dev runs (pyi_splash only exists inside PyInstaller builds)."""
    try:
        import pyi_splash  # type: ignore
        pyi_splash.close()
    except Exception:
        pass


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        pass
    except Exception:
        traceback.print_exc()
        try:
            _r = tk.Tk(); _r.withdraw()
            messagebox.showerror("Fatal Error", traceback.format_exc())
        except:
            pass
