"""Focused regression tests for the matcher embedded in views/negative_stock.py.

Run: pip install pandas pytest streamlit openpyxl altair && pytest -q
The matcher section is executed in isolation so importing the Streamlit page does
not run its UI/authentication code.
"""
import math
import re
from pathlib import Path

import pandas as pd


class _StreamlitStub:
    @staticmethod
    def cache_data(*args, **kwargs):
        def decorate(fn):
            return fn
        return decorate


def _load_matcher():
    source = (Path(__file__).parents[1] / "views" / "negative_stock.py").read_text(encoding="utf-8")
    start = source.index("# ==================== matching v2 ====================")
    end = source.index("# ==================== sheet writer ====================", start)
    namespace = {
        "pd": pd, "re": re, "math": math, "st": _StreamlitStub(),
        "COST_DP": 4,
        "cost_drift": lambda derived, code, pairs: None,
    }
    exec(compile(source[start:end], "matching_section", "exec"), namespace)
    return namespace


def _master(rows):
    defaults = {
        "Item No": "", "Stock": 0, "Cost": 0, "WAC": 0, "Net MRP": 0,
        "Is Active": "Checked", "Class": "GROCERY", "Category": "SUPERMARKET",
        "Group": "DRINKS", "Brand": "TEST", "Pack Qty": 0, "Packing": "",
        "IsBundleItem": "",
    }
    return pd.DataFrame([{**defaults, **r} for r in rows]).rename(columns={
        "Stock": "stock", "Cost": "last_cost", "WAC": "wac", "Net MRP": "mrp"
    }).assign(cost=lambda df: df[["last_cost", "wac"]].max(axis=1))


def _negative(barcode, name, qty=-6, **extra):
    return {"bc": barcode, "Item Name": name, "qty": qty, "val": -12.0,
            "category": "GROCERY", **extra}


def test_pack_parent_is_matched_and_conversion_is_explicit():
    ns = _load_matcher()
    master = _master([
        {"Item Barcode": "100", "Item Name": "TEST COLA 330ML 6 PACK",
         "stock": 10, "Cost": 12, "WAC": 12, "Net MRP": 15, "Pack Qty": 6},
        {"Item Barcode": "200", "Item Name": "TEST COLA 330ML",
         "stock": 0, "Cost": 2, "WAC": 2, "Net MRP": 2.5},
    ])
    breaks, combos = ns["find_breaks"](master, pd.DataFrame([_negative("200", "TEST COLA 330ML")]), 0.35)
    assert combos.empty
    assert len(breaks) == 1
    assert breaks.iloc[0]["par_bc"] == "100"
    assert breaks.iloc[0]["conv"] == 6
    assert breaks.iloc[0]["outers_needed"] == 1
    assert bool(breaks.iloc[0]["covered"])


def test_size_conflict_does_not_match_different_product_size():
    ns = _load_matcher()
    master = _master([
        {"Item Barcode": "100", "Item Name": "TEST COLA 330ML 6 PACK",
         "stock": 10, "Cost": 12, "WAC": 12, "Net MRP": 15, "Pack Qty": 6},
        {"Item Barcode": "200", "Item Name": "TEST COLA 500ML",
         "stock": 0, "Cost": 2, "WAC": 2, "Net MRP": 2.5},
    ])
    breaks, _ = ns["find_breaks"](master, pd.DataFrame([_negative("200", "TEST COLA 500ML")]), 0.35)
    assert breaks.empty


def test_mixed_size_combo_is_flagged_for_manual_review():
    ns = _load_matcher()
    master = _master([
        {"Item Barcode": "100", "Item Name": "TEST SHAMPOO 100ML + 50ML",
         "stock": 10, "Cost": 12, "WAC": 12, "Net MRP": 15, "Pack Qty": 2},
    ])
    breaks, combos = ns["find_breaks"](
        master, pd.DataFrame([_negative("200", "TEST SHAMPOO 100ML + 50ML")]), 0.35
    )
    assert breaks.empty
    assert len(combos) == 1
    assert "verify manually" in combos.iloc[0]["why"]


def test_wrong_sale_candidate_uses_stock_and_size_guard():
    ns = _load_matcher()
    master = _master([
        {"Item Barcode": "100", "Item Name": "TEST COLA 330ML",
         "stock": 50, "Cost": 2, "WAC": 2, "Net MRP": 2.5},
        {"Item Barcode": "101", "Item Name": "TEST COLA 500ML",
         "stock": 50, "Cost": 3, "WAC": 3, "Net MRP": 3.5},
        {"Item Barcode": "200", "Item Name": "OTHER ITEM 330ML",
         "stock": 0, "Cost": 2, "WAC": 2, "Net MRP": 2.5},
    ])
    swaps = ns["find_swaps"](master, pd.DataFrame([_negative("200", "TEST COLA 330ML", qty=-2)]), 0.1, 0.8, 0.25)
    assert len(swaps) == 1
    assert swaps.iloc[0]["par_bc"] == "100"
