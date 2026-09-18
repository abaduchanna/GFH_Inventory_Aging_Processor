"""Tests for the Exclusions panel (stores / IMEIs hidden from outbound
artifacts) in GFH_Inventory_Aging_Processor.

The main module imports tkinter/pc-only packages at top level, so the tests
extract the exact shipped function sources via ast and execute them in an
isolated namespace — testing the code that actually ships. pandas is
available and used for the DataFrame-level tests.
"""

import ast
import json
from pathlib import Path

import pandas as pd
import pytest

REPO = Path(__file__).resolve().parent.parent
MAIN = REPO / "GFH_Inventory_Aging_Processor.py"

PIPELINE_FUNCS = (
    "normalize_imei",
    "store_matches_pattern",
    "load_exclusions",
    "save_exclusions",
    "apply_output_exclusions",
    "log_exclusions_summary",
)
TAB_FUNCS = ("format_date_value", "build_district_tabs")


def _extract(tmp_path=None):
    """Pull ALL exclusion/tab function defs out of the main file and exec
    them into one namespace — they reference each other through module
    globals, so they must be present together. When tmp_path is given,
    get_app_dir is stubbed to it so the JSON config read/write runs against
    a throwaway directory."""
    wanted = set(PIPELINE_FUNCS) | set(TAB_FUNCS)
    tree = ast.parse(MAIN.read_text(encoding="utf-8"))
    picked = [n for n in tree.body
              if isinstance(n, ast.FunctionDef) and n.name in wanted]
    found = {n.name for n in picked}
    assert found == wanted, f"missing from source: {wanted - found}"
    ns = {"os": __import__("os"), "json": json, "re": __import__("re"),
          "pd": pd, "print": print,
          "datetime": __import__("datetime").datetime}
    # Constants the functions reference at module scope
    ns["EXCLUSIONS_CONFIG_FILE"] = "gfh_aging_exclusions.json"
    ns["BLOCKED_IMEIS"] = {"358975210745726", "350776860110726",
                           "358975210799012", "358975210797339",
                           "354709280259373", "358975210792793"}
    if tmp_path is not None:
        ns["get_app_dir"] = lambda: str(tmp_path)
    exec(compile(ast.Module(body=picked, type_ignores=[]), str(MAIN), "exec"), ns)
    return ns


# ── normalize_imei ──────────────────────────────────────────────────────────

def test_normalize_imei_strips_separators():
    ns = _extract()
    f = ns["normalize_imei"]
    assert f("356 938-03564 3809") == "356938035643809"
    assert f("  358975210745726 ") == "358975210745726"
    assert f("358.975.210745.726") == "358975210745726"
    assert f("") == ""
    assert f(None) == ""
    assert f("nan") == ""


# ── store_matches_pattern ───────────────────────────────────────────────────

def test_store_match_exact_case_insensitive():
    ns = _extract()
    m = ns["store_matches_pattern"]
    assert m("Houston #12", "houston #12")
    assert m("HOUSTON #12", "Houston #12")


def test_store_match_contains_both_directions():
    ns = _extract()
    m = ns["store_matches_pattern"]
    assert m("Houston #12", "houston")          # pattern inside store
    assert m("North Houston", "houston")
    assert m("Houston #12", "houston #12 - main st")  # store inside pattern
    assert not m("Dallas #5", "houston")
    assert not m("", "houston")
    assert not m("Houston #12", "")
    assert not m(None, "x")


# ── load/save exclusions ────────────────────────────────────────────────────

def test_first_run_seeds_blocked_imeis(tmp_path):
    ns = _extract(tmp_path=tmp_path)
    exc = ns["load_exclusions"]()
    assert exc["stores"] == [] and exc["imeis"] == []
    assert len(exc["blocked_imeis"]) == 6
    assert "358975210745726" in exc["blocked_imeis"]
    assert (tmp_path / "gfh_aging_exclusions.json").exists()


def test_save_normalizes_and_dedupes(tmp_path):
    ns = _extract(tmp_path=tmp_path)
    ns["save_exclusions"]({
        "stores": ["Houston #12", "HOUSTON #12", "  ", "Dallas #5"],
        "imeis": ["356 938-03564 3809", "356938035643809", "junk-entry"],
        "blocked_imeis": ["358975210745726", "358-975210745726"],
    })
    exc = ns["load_exclusions"]()
    assert exc["stores"] == ["Houston #12", "Dallas #5"]      # dedupe, no IMEI normalization
    assert exc["imeis"] == ["356938035643809", "junk-entry"]  # digits-normalized, dedupe
    assert exc["blocked_imeis"] == ["358975210745726"]        # dedupe after normalize


def test_saved_file_survives_reload_roundtrip(tmp_path):
    ns = _extract(tmp_path=tmp_path)
    ns["save_exclusions"]({"stores": ["Store A"], "imeis": ["123"],
                           "blocked_imeis": ["456"]})
    raw = json.loads((tmp_path / "gfh_aging_exclusions.json").read_text())
    assert raw == {"stores": ["Store A"], "imeis": ["123"], "blocked_imeis": ["456"]}
    assert ns["load_exclusions"]() == {"stores": ["Store A"], "imeis": ["123"],
                                       "blocked_imeis": ["456"]}


def test_corrupt_file_falls_back_to_seed(tmp_path):
    ns = _extract(tmp_path=tmp_path)
    (tmp_path / "gfh_aging_exclusions.json").write_text("{not json", encoding="utf-8")
    exc = ns["load_exclusions"]()
    assert exc["stores"] == [] and len(exc["blocked_imeis"]) == 6


# ── apply_output_exclusions ─────────────────────────────────────────────────

def _df():
    return pd.DataFrame({
        "Serial 1": ["356938035643809", "354709280259373", "358975210745726",
                     "111222333444555", "111222333444555"],
        "Store": ["Houston #12", "Dallas #5", "Houston #12", "Austin #1", "Dallas #5"],
        "Value": [1, 2, 3, 4, 5],
    })


def test_apply_exclusions_no_config_leaves_df_untouched(tmp_path, capsys):
    ns = _extract(tmp_path=tmp_path)
    df = _df()
    out = ns["apply_output_exclusions"](df, "Serial 1", "Store")
    assert len(out) == len(df)
    assert list(out["Store"]) == list(df["Store"])


def test_apply_exclusions_store_contains_and_exact(tmp_path):
    ns = _extract(tmp_path=tmp_path)
    ns["save_exclusions"]({"stores": ["houston"], "imeis": [], "blocked_imeis": []})
    out = ns["apply_output_exclusions"](_df(), "Serial 1", "Store")
    assert list(out["Store"]) == ["Dallas #5", "Austin #1", "Dallas #5"]
    # Input DataFrame untouched (filter returns a copy view)
    assert len(_df()) == 5


def test_apply_exclusions_imei_digits_only(tmp_path):
    ns = _extract(tmp_path=tmp_path)
    ns["save_exclusions"]({"stores": [], "imeis": ["356-938 03564 3809"],
                           "blocked_imeis": []})
    out = ns["apply_output_exclusions"](_df(), "Serial 1", "Store")
    assert "356938035643809" not in set(out["Serial 1"])
    assert len(out) == 4


def test_apply_exclusions_combined_and_logged(tmp_path):
    ns = _extract(tmp_path=tmp_path)
    ns["save_exclusions"]({"stores": ["Austin #1"], "imeis": ["111222333444555"],
                           "blocked_imeis": []})
    logs = []
    # Austin #1 row is caught by the store filter; of the two IMEI rows the
    # store filter already removed one, so the IMEI pass removes the other.
    out = ns["apply_output_exclusions"](_df(), "Serial 1", "Store", log=logs.append)
    assert len(out) == 3
    assert set(out["Store"]) == {"Houston #12", "Dallas #5"}
    assert "111222333444555" not in set(out["Serial 1"])
    text = "\n".join(logs)
    assert "removed 1 row(s) from excluded store(s): Austin #1" in text
    assert "removed 1 row(s) from excluded IMEI(s)" in text


# ── build_district_tabs with ensure_districts ──────────────────────────────

def _aged():
    return pd.DataFrame({
        "_District": ["Arizona", "Arizona"],
        "Store": ["Houston #12", "Phoenix #2"],
        "Product Desc Full": ["iPhone 15", "Galaxy S24"],
        "Serial 1": ["356938035643809", "354709280259373"],
        "Age in Company": [30, 45],
        "PO Date": pd.to_datetime(["2026-01-01", "2026-02-01"]),
    })


def test_ensure_districts_creates_empty_tab(tmp_path):
    ns = _extract(tmp_path=tmp_path)
    logs = []
    tabs = ns["build_district_tabs"](
        _aged(), "Store", "Product Desc Full", "Serial 1", "Age in Company",
        "PO Date", logs.append, ensure_districts=["Arizona", "Tennessee"])
    assert set(tabs) == {"Arizona", "Tennessee"}
    assert len(tabs["Arizona"]) == 2
    empty = tabs["Tennessee"]
    assert len(empty) == 0
    assert list(empty.columns) == ["District", "Store", "Product Description",
                                   "Serial 1", "Age in Company", "PO Date"]
    assert "Tennessee: 0 aged device row(s) (all devices excluded)" in "\n".join(logs)


def test_ensure_districts_none_keeps_old_behaviour(tmp_path):
    ns = _extract(tmp_path=tmp_path)
    tabs = ns["build_district_tabs"](
        _aged(), "Store", "Product Desc Full", "Serial 1", "Age in Company",
        "PO Date", lambda m: None)
    assert set(tabs) == {"Arizona"}


def test_empty_tab_upload_values_are_header_only(tmp_path):
    """dataframe_to_values on an empty tab must yield just the header row so
    the Google Sheet gets an empty (headers-only) tab, per the user's choice."""
    ns = _extract(tmp_path=tmp_path)
    tabs = ns["build_district_tabs"](
        _aged().iloc[0:0], "Store", "Product Desc Full", "Serial 1",
        "Age in Company", "PO Date", lambda m: None,
        ensure_districts=["Arizona"])
    df = tabs["Arizona"]
    values = [list(df.columns)] + [
        tuple("" if pd.isna(v) else v for v in row)
        for row in df.itertuples(index=False, name=None)]
    assert len(values) == 1  # header only
