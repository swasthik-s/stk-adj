"""
mongo_store.py — keep every run, and every generated sheet, in MongoDB Atlas.

Secrets (.streamlit/secrets.toml or the Secrets box on Streamlit Cloud):

    [mongo]
    uri = "mongodb+srv://user:password@cluster0.xxxxx.mongodb.net/?retryWrites=true&w=majority"
    db  = "stockadj"

Atlas free tier (M0) gives 512 MB. A generated sheet is about 7 KB and an
import file under 1 KB, so storage is not the constraint — thousands of
batches fit. Documents are capped at 16 MB by MongoDB, which a zip of ~40
sheets stays well under, but save_batch checks anyway.

Streamlit Cloud has no fixed outbound IP, so Atlas Network Access must allow
0.0.0.0/0. The password is then the only barrier — make it long, and give the
database user access to this one database only.

Collections
    runs     one document per matching run: settings, counts, candidate list
    batches  one document per generated set: metadata + the files as binary
"""

from datetime import datetime, timezone

from bson.binary import Binary
from pymongo import MongoClient, DESCENDING
from pymongo.errors import PyMongoError

STORE_VERSION = 2          # bump when methods are added
MAX_DOC = 15 * 1024 * 1024          # leave headroom under Mongo's 16 MB limit


class MongoStore:
    def __init__(self, uri, dbname="stockadj", timeout_ms=6000):
        self.client = MongoClient(uri, serverSelectionTimeoutMS=timeout_ms,
                                  connectTimeoutMS=timeout_ms)
        self.db = self.client[dbname]
        self.runs = self.db["runs"]
        self.batches = self.db["batches"]

    # ---------- health ----------
    def check(self):
        try:
            self.client.admin.command("ping")
        except PyMongoError as e:
            return False, f"Cannot reach MongoDB: {type(e).__name__}. " \
                          f"Check the URI, the password, and that Atlas Network " \
                          f"Access allows 0.0.0.0/0."
        try:
            stats = self.db.command("dbstats")
            mb = stats.get("dataSize", 0) / 1e6
            return True, (f"Connected to '{self.db.name}' — "
                          f"{self.runs.estimated_document_count()} runs, "
                          f"{self.batches.estimated_document_count()} batches, "
                          f"{mb:.1f} MB used of 512 MB.")
        except Exception:
            return True, f"Connected to '{self.db.name}'."

    def ensure_indexes(self):
        try:
            self.runs.create_index([("at", DESCENDING)])
            self.batches.create_index([("at", DESCENDING)])
            self.batches.create_index("run_id")
            self.db["snapshots"].create_index([("at", DESCENDING)])
            self.db["sessions"].create_index([("at", DESCENDING)])
        except PyMongoError:
            pass

    # ---------- writes ----------
    def save_snapshot(self, *, totals, by_category, source=None, note=""):
        """Called once per upload. A few hundred bytes — the point is the
        curve over time, so you can show whether the negative is falling."""
        doc = {
            "at": datetime.now(timezone.utc),
            "kind": "snapshot",
            "totals": totals,
            "by_category": by_category,
            "source": source or {},
            "note": note,
        }
        try:
            return True, str(self.db["snapshots"].insert_one(doc).inserted_id)
        except PyMongoError as e:
            return False, f"{type(e).__name__}"

    # ---------- sessions ----------------------------------------------------
    # A session is one working set: the negative report as loaded, the
    # candidates, and the settings. The masterlist is NOT stored — it is 100k
    # rows of reference data you already hold, and every figure a session
    # needs to rebuild its sheets (costs, conversions) lives on the candidates.

    NEG_COLS = ["bc", "Item Name", "category", "group", "qty", "val"]
    CAND_COLS = ["neg_bc", "neg_desc", "category", "neg_qty", "neg_val",
                 "par_bc", "par_desc", "par_stock", "par_cost", "conv",
                 "outers_needed", "kind", "cost_drift_pct", "problems", "use"]

    @staticmethod
    def _records(df, cols):
        if df is None or not len(df):
            return []
        keep = [c for c in cols if c in df.columns]
        out = df[keep].copy()
        for c in out.columns:          # Mongo cannot store numpy scalars/NaN
            if out[c].dtype.kind in "fc":
                out[c] = out[c].astype(float).where(out[c].notna(), None)
        return out.to_dict("records")

    def save_session(self, *, neg_df, source, settings, cand_df=None,
                     note="auto"):
        doc = {
            "at": datetime.now(timezone.utc),
            "updated": datetime.now(timezone.utc),
            "source": source,
            "settings": settings,
            "note": note,
            "neg_lines": int(len(neg_df)),
            "neg_value": round(float(neg_df["val"].sum()), 2),
            "negatives": self._records(neg_df, self.NEG_COLS),
            "candidates": self._records(cand_df, self.CAND_COLS),
            "n_candidates": 0 if cand_df is None else int(len(cand_df)),
            "cleared": (0.0 if cand_df is None or not len(cand_df)
                        else round(float(cand_df.loc[
                            cand_df.get("use", True) == True,
                            "neg_val"].abs().sum()), 2)),
        }
        try:
            return True, str(self.db["sessions"].insert_one(doc).inserted_id)
        except PyMongoError as e:
            return False, type(e).__name__

    def update_session(self, _id, *, cand_df=None, settings=None, note=None):
        from bson import ObjectId
        upd = {"updated": datetime.now(timezone.utc)}
        if cand_df is not None:
            upd["candidates"] = self._records(cand_df, self.CAND_COLS)
            upd["n_candidates"] = int(len(cand_df))
            use = cand_df["use"] if "use" in cand_df.columns else True
            upd["cleared"] = round(float(
                cand_df.loc[use == True, "neg_val"].abs().sum()), 2) \
                if len(cand_df) else 0.0
        if settings is not None:
            upd["settings"] = settings
        if note is not None:
            upd["note"] = note
        try:
            self.db["sessions"].update_one({"_id": ObjectId(_id)},
                                           {"$set": upd})
            return True, _id
        except PyMongoError as e:
            return False, type(e).__name__

    def list_sessions(self, limit=100):
        try:
            return list(self.db["sessions"].find(
                {}, {"negatives": 0, "candidates": 0})
                .sort("at", DESCENDING).limit(limit))
        except PyMongoError:
            return []

    def get_session(self, _id):
        from bson import ObjectId
        try:
            return self.db["sessions"].find_one({"_id": ObjectId(_id)})
        except PyMongoError:
            return None

    def delete_session(self, _id):
        from bson import ObjectId
        try:
            self.db["sessions"].delete_one({"_id": ObjectId(_id)})
            return True
        except PyMongoError:
            return False

    def list_snapshots(self, limit=200):
        try:
            return list(self.db["snapshots"].find({}, {"by_category": 0})
                        .sort("at", DESCENDING).limit(limit))
        except PyMongoError:
            return []

    def get_snapshot(self, _id):
        from bson import ObjectId
        try:
            return self.db["snapshots"].find_one({"_id": ObjectId(_id)})
        except PyMongoError:
            return None

    def save_run(self, *, categories, mode, settings, candidates_df,
                 combos_df=None, note=""):
        """Called automatically after every matching run."""
        def slim(df):
            if df is None or not len(df):
                return []
            keep = [c for c in ("neg_bc", "neg_desc", "category", "neg_qty",
                                "neg_val", "par_bc", "par_desc", "par_stock",
                                "par_cost", "conv", "outers_needed", "score",
                                "kind", "cost_drift_pct", "problems",
                                "blocked", "use") if c in df.columns]
            return df[keep].to_dict("records")

        doc = {
            "at": datetime.now(timezone.utc),
            "categories": list(categories),
            "mode": mode,
            "settings": settings,
            "n_candidates": int(len(candidates_df)) if candidates_df is not None else 0,
            "value": float(candidates_df["neg_val"].sum())
                     if candidates_df is not None and len(candidates_df) else 0.0,
            "candidates": slim(candidates_df),
            "combos_skipped": slim(combos_df),
            "note": note,
        }
        try:
            return True, str(self.runs.insert_one(doc).inserted_id)
        except PyMongoError as e:
            return False, f"{type(e).__name__}"

    def save_batch(self, *, label, files: dict, meta=None, run_id=None):
        """files: {filename: bytes}. Stored as binary inside one document."""
        total = sum(len(v) for v in files.values())
        if total > MAX_DOC:
            return False, (f"{total/1e6:.1f} MB is too large for one document. "
                           f"Generate fewer sheets per batch.")
        doc = {
            "at": datetime.now(timezone.utc),
            "label": label,
            "run_id": run_id,
            "meta": meta or {},
            "n_files": len(files),
            "bytes": total,
            "files": [{"name": k, "size": len(v), "data": Binary(v)}
                      for k, v in files.items()],
        }
        try:
            return True, str(self.batches.insert_one(doc).inserted_id)
        except PyMongoError as e:
            return False, f"{type(e).__name__}"

    # ---------- reads ----------
    def list_runs(self, limit=40):
        try:
            return list(self.runs.find(
                {}, {"candidates": 0, "combos_skipped": 0}
            ).sort("at", DESCENDING).limit(limit))
        except PyMongoError:
            return []

    def get_run(self, _id):
        from bson import ObjectId
        try:
            return self.runs.find_one({"_id": ObjectId(_id)})
        except PyMongoError:
            return None

    def list_batches(self, limit=60):
        try:
            return list(self.batches.find(
                {}, {"files.data": 0}
            ).sort("at", DESCENDING).limit(limit))
        except PyMongoError:
            return []

    def get_file(self, batch_id, name):
        from bson import ObjectId
        try:
            doc = self.batches.find_one({"_id": ObjectId(batch_id)},
                                        {"files": 1})
        except PyMongoError:
            return None
        for f in (doc or {}).get("files", []):
            if f["name"] == name:
                return bytes(f["data"])
        return None

    def delete_batch(self, batch_id):
        from bson import ObjectId
        try:
            self.batches.delete_one({"_id": ObjectId(batch_id)})
            return True
        except PyMongoError:
            return False
