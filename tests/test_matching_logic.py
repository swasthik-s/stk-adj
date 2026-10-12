"""Regression tests for the standalone negative-stock matcher."""
import pandas as pd
from matching_engine import find_breaks, find_swaps, _features

def master(rows):
    defaults = {
        "Item Barcode":"", "Item Name":"", "Stock":0, "Cost":0, "WAC":0,
        "Net MRP":0, "LP Price":0, "Is Active":"Checked", "Class":"GROCERY",
        "Category":"SUPERMARKET", "Group":"DRINKS", "Brand":"TEST", "Pack Qty":1,
        "Packing":"", "IsBundleItem":"Unchecked",
    }
    return pd.DataFrame([{**defaults, **row} for row in rows])

def neg(bc, name, qty=-6, category="GROCERY", sp=2.5):
    return {"bc":bc, "Item Name":name, "qty":qty, "val":-12.0, "category":category, "sp":sp}

def test_explicit_pack_match_and_conversion():
    m=master([
        {"Item Barcode":"100","Item Name":"TEST COLA 330ML 6 PACK","Stock":10,"Cost":12,"WAC":12,"LP Price":15,"Pack Qty":6},
        {"Item Barcode":"200","Item Name":"TEST COLA 330ML","Stock":0,"Cost":2,"WAC":2,"LP Price":2.5,"Pack Qty":1},
    ])
    out, combos=find_breaks(m,pd.DataFrame([neg("200","TEST COLA 330ML")]),0.35)
    assert combos.empty and len(out)==1
    assert out.iloc[0]["par_bc"]=="100" and out.iloc[0]["conv"]==6 and out.iloc[0]["outers_needed"]==1

def test_wrong_size_is_rejected():
    m=master([
        {"Item Barcode":"100","Item Name":"TEST COLA 330ML 6 PACK","Stock":10,"Cost":12,"WAC":12,"LP Price":15,"Pack Qty":6},
        {"Item Barcode":"200","Item Name":"TEST COLA 500ML","Stock":0,"Cost":2,"WAC":2,"LP Price":2.5,"Pack Qty":1},
    ])
    out,_=find_breaks(m,pd.DataFrame([neg("200","TEST COLA 500ML")]),0.35)
    assert out.empty

def test_combo_and_product_set_are_not_broken():
    m=master([
        {"Item Barcode":"100","Item Name":"TEST SHAMPOO 100ML + 50ML","Stock":10,"Cost":12,"WAC":12,"Pack Qty":2},
        {"Item Barcode":"101","Item Name":"STONE BANGLES 2PCS SET 2005","Stock":10,"Cost":12,"WAC":12,"Pack Qty":2},
    ])
    out, combos=find_breaks(m,pd.DataFrame([
        neg("200","TEST SHAMPOO 100ML + 50ML"),
        neg("201","STONE BANGLES 2PCS SET 2005"),
    ]),0.35)
    assert out.empty and len(combos)==2

def test_unrelated_shared_word_and_wrong_brand_are_rejected():
    m=master([
        {"Item Barcode":"100","Item Name":"AL AREEN PLASTICE BOWL BIG","Stock":8,"Cost":4.5,"WAC":4.5,"LP Price":4.5,"Pack Qty":20,"Class":"ROASTERY","Brand":"AL AREEN"},
        {"Item Barcode":"200","Item Name":"AJWA BIG AL RAIQA","Stock":-3,"Cost":52,"WAC":52,"LP Price":65,"Pack Qty":1,"Class":"ROASTERY","Brand":"AL RAIQA"},
    ])
    out,_=find_breaks(m,pd.DataFrame([neg("200","AJWA BIG AL RAIQA",qty=-3,category="ROASTERY",sp=65)]),0.35)
    assert out.empty

def test_metadata_pack_qty_needs_exact_name_and_unit_cost_evidence():
    good=master([
        {"Item Barcode":"100","Item Name":"TEST COLA 330ML","Stock":10,"Cost":12,"WAC":12,"Pack Qty":6},
        {"Item Barcode":"200","Item Name":"TEST COLA 330ML","Stock":0,"Cost":2,"WAC":2,"Pack Qty":1},
    ])
    out,_=find_breaks(good,pd.DataFrame([neg("200","TEST COLA 330ML")]),0.35)
    assert len(out)==1 and out.iloc[0]["conv"]==6
    bad=master([
        {"Item Barcode":"100","Item Name":"TEST COLA 330ML","Stock":10,"Cost":60,"WAC":60,"Pack Qty":6},
        {"Item Barcode":"200","Item Name":"TEST COLA 330ML","Stock":0,"Cost":2,"WAC":2,"Pack Qty":1},
    ])
    out,_=find_breaks(bad,pd.DataFrame([neg("200","TEST COLA 330ML")]),0.35)
    assert out.empty

def test_conflicting_name_and_master_pack_quantity_are_rejected():
    m=master([
        {"Item Barcode":"100","Item Name":"TEST COLA 330ML X 2","Stock":10,"Cost":12,"WAC":12,"Pack Qty":6},
        {"Item Barcode":"200","Item Name":"TEST COLA 330ML","Stock":0,"Cost":2,"WAC":2,"Pack Qty":1},
    ])
    out,_=find_breaks(m,pd.DataFrame([neg("200","TEST COLA 330ML")]),0.35)
    assert out.empty

def test_wrong_sale_matches_same_product_but_not_variant_or_model():
    m=master([
        {"Item Barcode":"100","Item Name":"TEST COLA 330ML","Stock":50,"Cost":2,"WAC":2,"LP Price":2.5,"Pack Qty":1},
        {"Item Barcode":"101","Item Name":"TEST COLA 500ML","Stock":50,"Cost":3,"WAC":3,"LP Price":2.5,"Pack Qty":1},
        {"Item Barcode":"102","Item Name":"MARMUM ORANGE JUICE 500ML","Stock":50,"Cost":2,"WAC":2,"LP Price":2.5,"Pack Qty":1,"Brand":"MARMUM"},
        {"Item Barcode":"103","Item Name":"RAJ FLARE LADDLE 3","Stock":50,"Cost":2,"WAC":2,"LP Price":2.5,"Pack Qty":1,"Brand":"RAJ"},
        {"Item Barcode":"200","Item Name":"TEST COLA 330ML","Stock":0,"Cost":2,"WAC":2,"LP Price":2.5,"Pack Qty":1},
        {"Item Barcode":"201","Item Name":"MARMUM APPLE JUICE 500ML","Stock":0,"Cost":2,"WAC":2,"LP Price":2.5,"Pack Qty":1,"Brand":"MARMUM"},
        {"Item Barcode":"202","Item Name":"RAJ FLARE LADDLE 2","Stock":0,"Cost":2,"WAC":2,"LP Price":2.5,"Pack Qty":1,"Brand":"RAJ"},
    ])
    swaps=find_swaps(m,pd.DataFrame([
        neg("200","TEST COLA 330ML",qty=-2),
        neg("201","MARMUM APPLE JUICE 500ML",qty=-2),
        neg("202","RAJ FLARE LADDLE 2",qty=-2),
    ]),0.1,0.8,0.25)
    assert len(swaps)==1 and swaps.iloc[0]["neg_bc"]=="200" and swaps.iloc[0]["par_bc"]=="100"

def test_common_pack_formats_and_product_features():
    assert _features("TEST COLA 12X330ML")["explicit_pack"]==12
    assert _features("TEST COLA PACK OF 6")["explicit_pack"]==6
    assert _features("TEST COLA 12 PCS")["explicit_pack"]==12
    assert _features("PRESIDENT BACK PACK 22 inch")["explicit_pack"]==0
    assert _features("STONE BANGLES 2PCS SET 2005")["is_set"]
    assert _features("COCA-COLA 1 LTR-SPRITE+FANTA 3 X 1LTR")["is_combo"]
