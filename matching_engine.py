"""Conservative, explainable matcher for iTrade negative-stock adjustments.

This module only proposes candidates. It never posts stock adjustments.
Pack Qty is metadata, not proof that a SKU is an outer pack.
"""
from __future__ import annotations
import math
import re
from collections import defaultdict
from typing import Any
import pandas as pd

SIZE_RE = re.compile(r"(?:(?<![A-Z0-9])|(?<=[Xx]))(\d+(?:\.\d+)?)\s*(KG|KGS|G|GM|GMS|GRAMS?|ML|MLTR|LTR|L)(?![A-Z])", re.I)
PACK_PATTERNS = (
    re.compile(r"(?<![A-Z0-9])(\d{1,3})\s*[xX]\s*(?=\d)", re.I),
    re.compile(r"\b[xX]\s*(\d{1,3})\b(?!\s*(?:KG|KGS|GM|GMS|G|ML|MLTR|LTR|L)\b)", re.I),
    re.compile(r"\b(?:PACK|PACKS|CARTON|CTN|CASE|BOX|BUNDLE|MULTI)\s+OF\s+(\d{1,3})\b", re.I),
    re.compile(r"\b(\d{1,3})\s*(?:PACK|PACKS|PCS|PC|PIECES|UNITS|BOTTLES|CANS|TINS|ROLLS|BARS|SACHETS|CTN|CARTON|BOXES|CASES)\b", re.I),
    re.compile(r"\b(?:TWIN|TWN)\s*[- ]?PACK\b", re.I),
)
CLOTHING_SIZE_RE = re.compile(r"(?<![A-Z0-9'’])(?:XXS|XS|S|M|L|XL|XXL|XXXL|[2-9]XL)(?![A-Z0-9])", re.I)
NOISE = {
    "THE","AND","WITH","FOR","OF","IN","NO","NBR","ITEM","PACK","PACKS","PACKET",
    "BOX","BOXES","CARTON","CTN","CASE","CASES","BUNDLE","MULTI","OFFER","OFR",
    "SPECIAL","SPCL","SET","PCS","PC","PIECES","UNITS","BOTTLES","CANS","TINS",
    "ROLLS","BARS","BAR","SACHETS","KG","KGS","G","GM","GMS","GRAM","GRAMS","ML","MLTR",
    "LTR","L","EA","EACH","SIZE","VALUE","PROMO","PROMOTION","X","S"
}
ALIASES = {
    "GMS":"GM","GRAM":"GM","GRAMS":"GM","G":"GM","KGS":"KG","KILO":"KG",
    "KILOGRAM":"KG","KILOGRAMS":"KG","MILLILITER":"ML","MILLILITRE":"ML",
    "MLTR":"ML","LITRE":"LTR","LITER":"LTR","LITRES":"LTR","LITERS":"LTR",
    "EVPRTD":"EVAPORATED","EVP":"EVAPORATED","CLS":"CLASSIC","BISCTS":"BISCUIT",
    "PC":"PCS","PIECE":"PCS","PIECES":"PCS","PCE":"PCS","EA":"PCS","EACH":"PCS",
    "UNIT":"PCS","CTN":"CARTON","PKT":"PACKET","PCH":"POUCH","CHKN":"CHICKEN",
    "CHOC":"CHOCOLATE","CRM":"CREAM","ASSTD":"ASSORTED","MIXD":"MIXED",
    "VEGETABLES":"VEGETABLE","WALLNUT":"WALNUT","BLK":"BLACK","BLU":"BLUE",
    "GRY":"GRAY","GREY":"GRAY","WHT":"WHITE","NVY":"NAVY","BRN":"BROWN",
    "GRN":"GREEN","YLW":"YELLOW","ORG":"ORANGE","HNY":"HONEY","LADIES":"WOMENS",
    "WMNS":"WOMENS","MNS":"MENS","INFNTY":"INFINITY","OXGNTD":"OXYGENATED",
    "MOISTURISING":"MOISTURIZING","CHILLY":"CHILLI","SHMP":"SHAMPOO"
}
VARIANT_TERMS = {
    "EXTRA","STRONG","SOFT","MILD","ORIGINAL","REGULAR","NEW","CLASSIC","BROWN","WHITE","BLACK",
    "RED","BLUE","GREEN","YELLOW","PINK","ORANGE","APPLE","MANGO","STRAWBERRY","VANILLA",
    "CHOCOLATE","ALMOND","CASHEW","WALNUT","SALTED","UNSALTED","SWEET","SUGAR","FREE",
    "LIGHT","DARK","LEMON","LIME","MINT","GARLIC","ONION","CHILLI","CHILI","HOT","SPICY",
    "FAT","LOW","FULL","LF","FF","SKIMMED","FRESH","DRY","LARGE","SMALL","BIG","BABY",
    "MEN","MENS","WOMEN","WOMENS","GIRLS","BOYS","LADIES","GENTS","HERBAL","BASMATI",
    "JASMINE","LONG","SHORT","DEEP","CLEAN","DETOX","PUDDING","FRUIT","MULTIGRAIN",
    "WHOLEWHEAT","WHOLEMEAL","WHEAT","RYE","HONEY","NAVY","GRAY","LAVENDER","CITRUS",
    "CEDAR","AGAR","ALOE","VERA","TOMATO","SAUCE","GOLD","MORNING","MOOD","SPIRIT",
    "ARGAN","TURMERIC","PEPPER","GINGER","OXYGENATED","PIMPLES","GENTLE","MOISTURIZING",
    "KASHMIRI","COCOA","COCONUT","COCKTAIL","NECTAR","SHOWER","VITA","FRANKS","CONE",
    "SANDALWOOD","ROSE","STATUS","ABSOLUTE","BOUND","INTENSE","PANGUSH","CHICKEN","GIZZARD",
    "LIVER","MUTTON","BEEF","VEAL","TUNA","SARDINE","MYSOREPAK","LADDU","FROCK","CARGO",
    "WINTER","SUMMER","SLICE","SLICED"
}
def _text(v: Any) -> str:
    return "" if v is None or str(v).strip().lower() in {"nan","none"} else str(v).strip()
def _num(v: Any) -> float:
    try:
        x = pd.to_numeric(_text(v).replace(",",""), errors="coerce")
        return float(x) if pd.notna(x) else 0.0
    except (TypeError, ValueError):
        return 0.0
def _canon_size(amount, unit):
    q, u = float(amount), ALIASES.get(str(unit).upper(), str(unit).upper())
    if u == "KG": return (round(q*1000, 5), "MASS")
    if u in {"GM","G","GMS"}: return (round(q, 5), "MASS")
    if u in {"LTR","L"}: return (round(q*1000, 5), "VOLUME")
    if u in {"ML","MLTR"}: return (round(q, 5), "VOLUME")
    return (round(q,5),u)
def _features(desc, pack_qty=None, packing=None):
    raw = _text(desc).upper()
    size_matches = list(SIZE_RE.finditer(raw))
    sizes = {_canon_size(m.group(1),m.group(2)) for m in size_matches}
    clothing_sizes = {m.group(0).upper() for m in CLOTHING_SIZE_RE.finditer(raw)}
    explicit_pack = 0
    for pat in PACK_PATTERNS:
        match = pat.search(raw)
        if match:
            explicit_pack = max(explicit_pack, int(match.group(1)) if match.lastindex else 2)
    is_set = bool(re.search(r"\bSET\b", raw))
    packing_text = _text(packing).upper()
    if not explicit_pack and re.search(r"\b(?:TWN|TWIN)\s*PK\b", packing_text):
        explicit_pack = 2
    metadata_pack = _num(pack_qty)
    is_combo = ("+" in raw) or bool(re.search(r"\b(?:COMBO\s*PACK|VARIETY\s*PACK)\b", raw))
    weighted = bool(re.search(r"(?<!\d)\bKG\b", raw) and not re.search(r"\d+(?:\.\d+)?\s*KG\b", raw))
    cleaned = raw.replace("&", " AND ")
    for match in reversed(size_matches):
        cleaned = cleaned[:match.start()] + " " + cleaned[match.end():]
    for pat in PACK_PATTERNS:
        cleaned = pat.sub(" ", cleaned)
    cleaned = CLOTHING_SIZE_RE.sub(" ", cleaned)
    cleaned = re.sub(r"\b\d+\s*['’]S\b", " ", cleaned)
    tokens = set()
    for word in re.findall(r"[A-Z0-9]+(?:\.[0-9]+)?", cleaned):
        word = ALIASES.get(word, word)
        if word in NOISE or re.fullmatch(r"\d+[Xx]|[Xx]\d+", word):
            continue
        tokens.add(word)
    ordered_tokens = []
    for word in re.findall(r"[A-Z0-9]+(?:\.[0-9]+)?", cleaned):
        word = ALIASES.get(word, word)
        if word in NOISE or re.fullmatch(r"\d+[Xx]|[Xx]\d+", word):
            continue
        if word not in ordered_tokens:
            ordered_tokens.append(word)
    numeric_codes = {t for t in tokens if re.fullmatch(r"\d+", t)}
    variants = tokens & VARIANT_TERMS
    return {
        "tokens":tokens,"first_token":ordered_tokens[0] if ordered_tokens else "",
        "sizes":sizes,"clothing_sizes":clothing_sizes,"numeric_codes":numeric_codes,
        "variants":variants,"explicit_pack":explicit_pack,
        "metadata_pack":int(round(metadata_pack)) if metadata_pack>=1 else 0,
        "is_combo":is_combo,"is_set":is_set,"weighted":weighted,
    }
def _similarity(a,b):
    if not a or not b: return 0.0,0,0.0,0.0
    overlap=len(a&b); jaccard=overlap/max(1,len(a|b)); containment=overlap/max(1,min(len(a),len(b)))
    return 0.60*containment+0.40*jaccard,overlap,containment,jaccard
def _active(v):
    return _text(v).upper() not in {"UNCHECKED","FALSE","NO","0","INACTIVE"}
def _normalize_master(master):
    m=master.copy()
    aliases={"stock":"Stock","last_cost":"Cost","wac":"WAC","mrp":"Net MRP"}
    for dst,src in aliases.items():
        if dst not in m.columns and src in m.columns: m[dst]=m[src]
    defaults={"Item Barcode":"","Item Name":"","Class":"","Category":"","Group":"","Brand":"",
              "Is Active":"Checked","IsBundleItem":"","Packing":"","Pack Qty":0,
              "stock":0,"last_cost":0,"wac":0,"mrp":0,"cost":0}
    for c,v in defaults.items():
        if c not in m: m[c]=v
    if ("Net MRP" not in master.columns or not pd.to_numeric(master.get("Net MRP",pd.Series(dtype=float)),errors="coerce").fillna(0).gt(0).any()) and "LP Price" in master.columns:
        m["mrp"]=pd.to_numeric(m["LP Price"].astype(str).str.replace(",","",regex=False),errors="coerce").fillna(0)
    for c in ("stock","last_cost","wac","mrp","cost","Pack Qty"):
        m[c]=pd.to_numeric(m[c].astype(str).str.replace(",","",regex=False),errors="coerce").fillna(0)
    m["cost"]=m[["last_cost","wac","cost"]].max(axis=1)
    m["Item Barcode"]=m["Item Barcode"].astype(str).str.strip()
    m["_active"]=m["Is Active"].map(_active)
    m["_features"]=[_features(r.get("Item Name"),r.get("Pack Qty"),r.get("Packing")) for r in m.to_dict("records")]
    m["_class"]=m["Class"].fillna("").astype(str).str.strip().str.upper()
    m["_brand"]=m["Brand"].fillna("").astype(str).str.strip().str.upper()
    m["_group"]=m["Group"].fillna("").astype(str).str.strip().str.upper()
    return m.drop_duplicates("Item Barcode").reset_index(drop=True)
def _core_gate(nf,pf,nclass,pclass,nbrand="",pbrand="",min_containment=0.62,min_jaccard=0.55):
    if nclass and pclass and nclass!=pclass: return False,"department/class mismatch"
    if nbrand and pbrand and nbrand not in {"NO BRAND","N/A","NONE"} and pbrand not in {"NO BRAND","N/A","NONE"} and nbrand!=pbrand:
        return False,"brand mismatch"
    if nf.get("first_token") and pf.get("first_token") and nf["first_token"]!=pf["first_token"] and not(nbrand and pbrand and nbrand==pbrand):
        return False,"leading product token mismatch"
    if nf["sizes"] and pf["sizes"] and nf["sizes"].isdisjoint(pf["sizes"]): return False,"explicit size mismatch"
    if nf.get("clothing_sizes")!=pf.get("clothing_sizes") and (nf.get("clothing_sizes") or pf.get("clothing_sizes")):
        return False,"clothing size mismatch"
    if nf.get("numeric_codes")!=pf.get("numeric_codes") and (nf.get("numeric_codes") or pf.get("numeric_codes")):
        return False,"model/style code mismatch"
    if nf.get("variants")!=pf.get("variants") and (nf.get("variants") or pf.get("variants")):
        return False,"product variant mismatch"
    sim,overlap,contain,jaccard=_similarity(nf["tokens"],pf["tokens"])
    if overlap<2: return False,"insufficient distinctive token overlap"
    if contain<min_containment: return False,"description containment below threshold"
    if jaccard<min_jaccard: return False,"too many unmatched product descriptors"
    return True,""
def _score(nf,pf,nclass,pclass,nbrand="",pbrand="",ngroup="",pgroup=""):
    sim,overlap,contain,jaccard=_similarity(nf["tokens"],pf["tokens"])
    score=sim; reasons=[f"description overlap {sim:.2f} ({overlap} shared tokens)"]
    if nf["sizes"] and pf["sizes"] and nf["sizes"]&pf["sizes"]:
        score+=0.12; reasons.append("size agrees")
    if nclass and pclass and nclass==pclass:
        score+=0.10; reasons.append("same department")
    if nbrand and pbrand and nbrand==pbrand:
        score+=0.08; reasons.append("same brand")
    if ngroup and pgroup and ngroup==pgroup:
        score+=0.05; reasons.append("same group")
    return round(max(0,min(1,score)),3),"; ".join(reasons),overlap,contain
def _cost_ratio(candidate_cost,conv,own_costs):
    unit_cost=candidate_cost/conv if conv>0 else 0
    costs=[x for x in own_costs if x>0]
    if unit_cost<=0 or not costs: return None
    return min(abs(unit_cost-c)/max(c,unit_cost) for c in costs)
def _make_index(records):
    idx=defaultdict(list)
    for i,r in enumerate(records):
        for token in r["_features"]["tokens"]: idx[token].append(i)
    return idx
def _candidate_indices(index,tokens):
    hits=set()
    for token in tokens: hits.update(index.get(token,()))
    return hits
def _own_cost_map(m):
    return {str(r["Item Barcode"]):(float(r["last_cost"]),float(r["wac"]),float(r["cost"])) for r in m.to_dict("records")}
def _cost_drift(derived,own_pair):
    if not derived: return None
    vals=[c for c in own_pair if c and c>0]
    if not vals: return None
    return round(min(((derived-c)/c*100 for c in vals),key=abs),1)
def find_breaks(master,negatives,threshold,skip_combo=True):
    m=_normalize_master(master); records=m.to_dict("records"); token_idx=_make_index(records)
    own=_own_cost_map(m); bybc={str(r["Item Barcode"]):r for r in records}
    out,combos=[],[]
    for r in negatives.to_dict("records"):
        bc=_text(r.get("bc")); name=_text(r.get("Item Name")); nf=_features(name,r.get("Pack Qty"),r.get("Packing"))
        nclass=_text(r.get("category")).upper(); ownrow=bybc.get(bc,{})
        nbrand=_text(ownrow.get("_brand")); ngroup=_text(ownrow.get("_group"))
        if not nf["tokens"]: continue
        if skip_combo and (nf["is_combo"] or nf["is_set"]):
            combos.append({"neg_bc":bc,"neg_desc":name,"neg_qty":r.get("qty",0),"neg_val":r.get("val",0),
                           "par_bc":"","par_desc":"","why":"combo/set SKU; manual review required"})
            continue
        if nf["explicit_pack"]>1:
            combos.append({"neg_bc":bc,"neg_desc":name,"neg_qty":r.get("qty",0),"neg_val":r.get("val",0),
                           "par_bc":"","par_desc":"","why":"negative SKU is already a multipack; verify manually"})
            continue
        if nf["weighted"]: continue
        for i in _candidate_indices(token_idx,nf["tokens"]):
            p=records[i]; pf=p["_features"]
            if p["Item Barcode"]==bc or not p["_active"] or float(p["stock"])<=0: continue
            if p["_class"] and nclass and p["_class"]!=nclass: continue
            if pf["is_set"] or pf["is_combo"]: continue
            conv=pf["explicit_pack"]; explicit=conv>1
            if explicit and pf["metadata_pack"]>1 and pf["metadata_pack"]!=conv: continue
            if conv<=1:
                conv=pf["metadata_pack"]
                if conv<=1: continue
            okay,_=_core_gate(nf,pf,nclass,p["_class"],nbrand,p["_brand"],
                              min_containment=0.68 if not explicit else 0.60,
                              min_jaccard=0.65 if explicit else 0.75)
            if not okay: continue
            own_costs=[x for x in own.get(bc,(0,0,0)) if x>0]
            ratio=_cost_ratio(float(p["cost"]),conv,own_costs)
            if not explicit:
                if ratio is None or ratio>0.35: continue
                if nf["tokens"]!=pf["tokens"] or _similarity(nf["tokens"],pf["tokens"])[3]<0.98: continue
            elif ratio is not None and ratio>0.65: continue
            qty=abs(_num(r.get("qty"))); conv=int(conv); need=math.ceil(qty/conv)
            available=float(p["stock"])
            if need>available: continue
            # Clear the negative first, then use up to three additional outer
            # packs when they are physically available. Keep the total in
            # outers, because this is the quantity staff must break.
            spare=max(0, math.floor(available)-need)
            extra=min(3, spare)
            break_count=need+extra
            score,why_score,_,_=_score(nf,pf,nclass,p["_class"],nbrand,p["_brand"],ngroup,p["_group"])
            if score<max(0.62,float(threshold)): continue
            own_pair=own.get(bc,(0,0,0)); derived=float(p["cost"])/conv
            out.append({"neg_bc":bc,"neg_desc":name,"category":r.get("category",""),
                "neg_qty":r.get("qty",0),"neg_val":r.get("val",0),"par_bc":p["Item Barcode"],
                "par_desc":p["Item Name"],"par_stock":p["stock"],"par_cost":p["cost"],
                "conv":conv,"outers_needed":break_count,"required_outers":need,
                "extra_outers":extra,"covered":True,"score":score,"kind":"Bundle break",
                "own_cost":own_pair[2],"derived_cost":round(derived,7),
                "cost_drift_pct":_cost_drift(derived,own_pair[:2]),
                "why_match":why_score+("; explicit pack quantity" if explicit else "; metadata pack qty plus unit-cost corroboration")})
    if out:
        out.sort(key=lambda x:(x["score"],-abs(x.get("cost_drift_pct") or 0)),reverse=True)
        seen=set(); filtered=[]
        for row in out:
            if row["neg_bc"] in seen: continue
            seen.add(row["neg_bc"]); filtered.append(row)
        out=filtered
    return pd.DataFrame(out),pd.DataFrame(combos)
def find_swaps(master,negatives,lo,hi,price_tol):
    m=_normalize_master(master); records=m.to_dict("records"); token_idx=_make_index(records)
    own=_own_cost_map(m); bybc={str(r["Item Barcode"]):r for r in records}; out=[]
    for r in negatives.to_dict("records"):
        bc=_text(r.get("bc")); name=_text(r.get("Item Name")); nf=_features(name,r.get("Pack Qty"),r.get("Packing"))
        nclass=_text(r.get("category")).upper(); ownrow=bybc.get(bc,{})
        nbrand=_text(ownrow.get("_brand")); ngroup=_text(ownrow.get("_group"))
        if len(nf["tokens"])<2 or nf["is_combo"] or nf["is_set"] or nf["explicit_pack"]>1: continue
        candidates=[]
        for i in _candidate_indices(token_idx,nf["tokens"]):
            p=records[i]; pf=p["_features"]
            if p["Item Barcode"]==bc or not p["_active"] or float(p["stock"])<=0: continue
            if p["_class"] and nclass and p["_class"]!=nclass: continue
            if pf["explicit_pack"]>1 or pf["is_set"] or pf["is_combo"]: continue
            okay,_=_core_gate(nf,pf,nclass,p["_class"],nbrand,p["_brand"],min_containment=0.90,min_jaccard=0.85)
            if not okay: continue
            score,why_score,_,_=_score(nf,pf,nclass,p["_class"],nbrand,p["_brand"],ngroup,p["_group"])
            if score<max(0.64,float(lo)): continue
            qty=abs(_num(r.get("qty")))
            if float(p["stock"])<qty: continue
            sp=_num(r.get("sp")); mrp=float(p["mrp"])
            if sp>0 and mrp>0 and abs(sp-mrp)/max(sp,mrp)>float(price_tol): continue
            own_pair=own.get(bc,(0,0,0)); derived=float(p["cost"])
            candidates.append({"neg_bc":bc,"neg_desc":name,"category":r.get("category",""),
                "neg_qty":r.get("qty",0),"neg_val":r.get("val",0),"par_bc":p["Item Barcode"],
                "par_desc":p["Item Name"],"par_stock":p["stock"],"par_cost":derived,
                "conv":1,"outers_needed":qty,"covered":True,"score":score,"kind":"Wrong sale",
                "own_cost":own_pair[2],"derived_cost":round(derived,7),
                "cost_drift_pct":_cost_drift(derived,own_pair[:2]),"why_match":why_score})
        if candidates:
            candidates.sort(key=lambda x:x["score"],reverse=True); out.append(candidates[0])
    return pd.DataFrame(out)
