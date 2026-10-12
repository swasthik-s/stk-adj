"""Higher-recall, explainable matching for iTrade negative-stock adjustments.

This module proposes candidates only. It never posts or applies an adjustment.
Candidates still require business-rule validation and human review.
"""
from __future__ import annotations
import math
import re
from collections import defaultdict
from typing import Any
import pandas as pd

ALIASES = {
    "GMS":"GM","GRAM":"GM","GRAMS":"GM","G":"GM","KGS":"KG","KILO":"KG",
    "KILOGRAM":"KG","KILOGRAMS":"KG","MILLILITER":"ML","MILLILITRE":"ML",
    "MLTR":"ML","LITRE":"LTR","LITER":"LTR","LITRES":"LTR","LITERS":"LTR",
    "PC":"PCS","PIECE":"PCS","PIECES":"PCS","PCE":"PCS","EA":"PCS","EACH":"PCS",
    "UNIT":"PCS","CTN":"CARTON","PKT":"PACKET","PCH":"POUCH","CHKN":"CHICKEN",
    "CHOC":"CHOCOLATE","CRM":"CREAM","ASSTD":"ASSORTED","MIXD":"MIXED",
    "VEGETABLES":"VEGETABLE","WALLNUT":"WALNUT","BLK":"BLACK","BLU":"BLUE",
}
SIZE_RE = re.compile(r"(?<![A-Z0-9])(\d+(?:\.\d+)?)\s*(KG|KGS|GM|GMS|G|ML|MLTR|LTR|L)(?![A-Z])", re.I)
PACK_RE = re.compile(r"(?<![A-Z0-9])(?:X\s*(\d{1,3})|(\d{1,3})\s*X|(\d{1,3})\s*(?:PCS|PC|PCE|EA))(?![A-Z])", re.I)
STOP = {"THE","AND","WITH","FOR","OF","NO","NBR","ITEM","PACK","PACKET","BOX","CARTON","OFFER","SPECIAL","SET","PCS","KG","GM","ML","LTR"}
def txt(v: Any) -> str:
    return "" if v is None or str(v).strip().lower() in {"nan","none"} else str(v).strip()
def num(v: Any) -> float:
    x = pd.to_numeric(txt(v).replace(",",""), errors="coerce")
    return float(x) if pd.notna(x) else 0.0
def features(value: Any):
    s = txt(value).upper().replace("&"," AND ")
    sizes = set()
    for amount, unit in SIZE_RE.findall(s):
        unit = ALIASES.get(unit.upper(), unit.upper())
        q = str(float(amount)).rstrip("0").rstrip(".")
        sizes.add((q,unit))
    packs = [int(next(g for g in match if g)) for match in PACK_RE.findall(s) if any(match)]
    tokens = set()
    for w in re.findall(r"[A-Z0-9]+(?:\.[0-9]+)?", s):
        w = ALIASES.get(w,w)
        if w not in STOP and len(w) > 1:
            tokens.add(w)
    return tokens, sizes, max(packs) if packs else 0
def canonical_size(pair):
    q,u = float(pair[0]), ALIASES.get(pair[1].upper(),pair[1].upper())
    if u == "KG": return (round(q*1000,5),"MASS")
    if u == "GM": return (round(q,5),"MASS")
    if u == "LTR": return (round(q*1000,5),"VOLUME")
    if u == "ML": return (round(q,5),"VOLUME")
    return (round(q,5),u)
def size_conflict(a,b):
    return bool(a and b and not ({canonical_size(x) for x in a} & {canonical_size(x) for x in b}))
def similarity(a,b):
    if not a or not b: return 0.0
    overlap = len(a & b)
    containment = overlap / max(1,min(len(a),len(b)))
    jaccard = overlap / max(1,len(a | b))
    return 0.65*containment + 0.35*jaccard
def normalize_master(master):
    m = master.copy()
    for dst,src in (("stock","Stock"),("last_cost","Cost"),("wac","WAC"),("mrp","Net MRP")):
        if dst not in m and src in m: m[dst] = m[src]
    defaults = {"Item Barcode":"","Item Name":"","Item No":"","Class":"","Category":"","Group":"","Brand":"","Is Active":"Checked","IsBundleItem":"","Packing":""}
    for c,v in defaults.items():
        if c not in m: m[c] = v
    for c in ("stock","last_cost","wac","mrp","Pack Qty"):
        if c not in m: m[c] = 1 if c == "Pack Qty" else 0
        m[c] = pd.to_numeric(m[c].astype(str).str.replace(",","",regex=False),errors="coerce").fillna(1 if c=="Pack Qty" else 0)
    if "cost" not in m: m["cost"] = m[["last_cost","wac"]].max(axis=1)
    m["cost"] = pd.to_numeric(m["cost"],errors="coerce").fillna(0)
    m["Item Barcode"] = m["Item Barcode"].astype(str).str.strip()
    return m.drop_duplicates("Item Barcode").reset_index(drop=True)
def truthy(v):
    return txt(v).strip().lower() in {"checked","true","yes","1","y","bundle"}
def build_index(master):
    m = normalize_master(master)
    records=[]; token_idx=defaultdict(list); class_idx=defaultdict(list); group_idx=defaultdict(list)
    for row in m.to_dict("records"):
        tokens,sizes,explicit_pack = features(row.get("Item Name"))
        brand = txt(row.get("Brand")).upper()
        bt = features(brand)[0] if brand and brand not in {"NO BRAND","N/A"} else set()
        cls = txt(row.get("Class") or row.get("category")).upper()
        group = txt(row.get("Group") or row.get("subgroup")).upper()
        active = txt(row.get("Is Active")).lower() not in {"unchecked","false","no","0","inactive"}
        packqty = num(row.get("Pack Qty"))
        # Only use an explicit numeric pack field or a parsed multiplier; never assume "box" means x2.
        if packqty <= 0: packqty = 1
        bundle = truthy(row.get("IsBundleItem"))
        rec = {"bc":txt(row.get("Item Barcode")),"name":txt(row.get("Item Name")),"tokens":tokens,"sizes":sizes,
               "pack":explicit_pack,"brand_tokens":bt,"class":cls,"group":group,"active":active,
               "bundle":bundle,"packqty":packqty,"stock":num(row.get("stock")),"mrp":num(row.get("mrp")),
               "cost":num(row.get("cost")),"last_cost":num(row.get("last_cost")),"wac":num(row.get("wac"))}
        idx=len(records); records.append(rec)
        for token in tokens | bt:
            token_idx[token].append(idx)
        if cls: class_idx[cls].append(idx)
        if group: group_idx[group].append(idx)
    return {"records":records,"tokens":token_idx,"classes":class_idx,"groups":group_idx}
def retrieve(row,index,limit=2200):
    tokens,_,_=features(row.get("Item Name"))
    hits=set()
    # Rare tokens first; avoid common-token flooding and keep the candidate pool bounded.
    for token in sorted(tokens,key=lambda t:(len(index["tokens"].get(t,())), -len(t))):
        hits.update(index["tokens"].get(token,()))
        if len(hits)>=limit: break
    cat=txt(row.get("category")).upper(); group=txt(row.get("subgroup")).upper()
    for key,idx in ((cat,index["classes"]),(group,index["groups"])):
        for i in idx.get(key,()):
            hits.add(i)
            if len(hits)>=limit: break
    return list(hits)
def rank(row,index,mode,top_k=80,min_score=0.0):
    name=txt(row.get("Item Name")); nt,ns,_=features(name)
    bc=txt(row.get("bc")); cls=txt(row.get("category")).upper(); group=txt(row.get("subgroup")).upper()
    price=num(row.get("sp")); ranked=[]
    for i in retrieve(row,index):
        p=index["records"][i]
        if not p["bc"] or p["bc"]==bc or not p["active"] or p["stock"]<=0: continue
        if mode=="break" and p["packqty"]<=1 and p["pack"]<=1: continue
        if mode=="swap" and p["bundle"]: continue
        pt=p["tokens"] | p["brand_tokens"]
        base=similarity(nt,pt); score=base; reasons=[f"name/token similarity {base:.2f}"]
        conflict=size_conflict(ns,p["sizes"])
        if conflict: score-=0.22; reasons.append("explicit size mismatch (-0.22)")
        if p["brand_tokens"]:
            if p["brand_tokens"] & nt: score+=0.10; reasons.append("brand evidence (+0.10)")
            else: score-=0.04
        if cls and p["class"]==cls: score+=0.08; reasons.append("same class (+0.08)")
        elif cls and p["class"]: score-=0.04
        if group and p["group"]==group: score+=0.06; reasons.append("same group (+0.06)")
        if mode=="swap" and price>0 and p["mrp"]>0:
            ratio=abs(price-p["mrp"])/max(price,p["mrp"])
            score+=max(-0.12,0.08-ratio*0.5); reasons.append(f"price distance {ratio:.0%}")
        if len(nt)>=2 and not (nt & pt): score-=0.25; reasons.append("no informative token overlap (-0.25)")
        score=max(0,min(1,score))
        if score<min_score: continue
        ranked.append({"neg_bc":bc,"neg_desc":name,"par_bc":p["bc"],"par_desc":p["name"],"par_stock":p["stock"],
                       "par_cost":p["cost"],"par_mrp":p["mrp"],"par_pack_qty":p["packqty"],
                       "kind":"Bundle break" if mode=="break" else "Wrong sale","score":round(score,3),
                       "why_match":"; ".join(reasons),"size_conflict":conflict,"same_class":p["class"]==cls,
                       "same_group":bool(group and p["group"]==group),"brand_match":bool(p["brand_tokens"] & nt)})
    ranked.sort(key=lambda x:x["score"],reverse=True)
    return ranked[:top_k]
def cost_drift(derived, own_values):
    vals=[num(x) for x in own_values if num(x)>0]
    if not vals or derived<=0: return None
    return round(min(((derived-v)/v*100 for v in vals),key=abs),1)
def find_breaks(master,negatives,threshold,skip_combo=True):
    idx=build_index(master); m=normalize_master(master)
    own={txt(r["Item Barcode"]):(r["cost"],r["last_cost"],r["wac"]) for r in m.to_dict("records")}
    out=[]; combos=[]
    for r in negatives.to_dict("records"):
        candidates=rank(r,idx,"break",top_k=100,min_score=max(0,threshold-0.20))
        for c in candidates:
            if c["score"] < threshold-0.05: break
            p=next((x for x in idx["records"] if x["bc"]==c["par_bc"]),None)
            if not p: continue
            # Two distinct pack sizes indicate a mixed combo; do not auto-suggest a single-target break.
            if skip_combo and len(p["sizes"])>=2:
                combos.append({"neg_bc":r.get("bc"),"neg_desc":r.get("Item Name"),"neg_qty":r.get("qty"),
                               "neg_val":r.get("val"),"par_bc":p["bc"],"par_desc":p["name"],
                               "why":"combo pack — multiple distinct sizes; verify manually"})
                continue
            conv=int(round(p["packqty"])) if p["packqty"]>1 else int(p["pack"])
            if conv<=1: continue
            need=math.ceil(abs(num(r.get("qty")))/conv)
            ownvals=own.get(txt(r.get("bc")),(0,0,0))
            drift=cost_drift(p["cost"]/conv,ownvals[1:])
            out.append({"neg_bc":r.get("bc"),"neg_desc":r.get("Item Name"),"category":r.get("category",""),
                        "neg_qty":r.get("qty"),"neg_val":r.get("val"),"par_bc":p["bc"],"par_desc":p["name"],
                        "par_stock":p["stock"],"par_cost":p["cost"],"conv":conv,"outers_needed":need,
                        "covered":need<=p["stock"],"score":c["score"],"kind":"Bundle break",
                        "own_cost":ownvals[0],"derived_cost":round(p["cost"]/conv,7),
                        "cost_drift_pct":drift,"why_match":c["why_match"]})
            break
    return pd.DataFrame(out),pd.DataFrame(combos)
def find_swaps(master,negatives,lo,hi,price_tol):
    idx=build_index(master); m=normalize_master(master)
    own={txt(r["Item Barcode"]):(r["cost"],r["last_cost"],r["wac"]) for r in m.to_dict("records")}
    out=[]
    for r in negatives.to_dict("records"):
        candidates=rank(r,idx,"swap",top_k=120,min_score=max(0,lo))
        for c in candidates:
            p=next((x for x in idx["records"] if x["bc"]==c["par_bc"]),None)
            if not p or p["stock"]<abs(num(r.get("qty"))): continue
            sp=num(r.get("sp"))
            if sp>0 and p["mrp"]>0 and abs(sp-p["mrp"])/max(sp,p["mrp"])>price_tol: continue
            vals=own.get(txt(r.get("bc")),(0,0,0))
            out.append({"neg_bc":r.get("bc"),"neg_desc":r.get("Item Name"),"category":r.get("category",""),
                        "neg_qty":r.get("qty"),"neg_val":r.get("val"),"par_bc":p["bc"],"par_desc":p["name"],
                        "par_stock":p["stock"],"par_cost":p["cost"],"conv":1,"outers_needed":abs(num(r.get("qty"))),
                        "covered":True,"score":c["score"],"kind":"Wrong sale","own_cost":vals[0],
                        "derived_cost":round(p["cost"],7),"cost_drift_pct":cost_drift(p["cost"],vals[1:]),
                        "why_match":c["why_match"]})
            break
    return pd.DataFrame(out)
