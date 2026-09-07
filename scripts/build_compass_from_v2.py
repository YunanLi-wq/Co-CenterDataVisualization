#!/usr/bin/env python3
"""
Build Compass tree JSON from flat data-item rows.

Reads:
  exports/compass_dataitems_v2_final.json

Writes (same schema the app expects):
  [
    {
      "Quadrant": "...",
      "children": [
        {"Segment": "...", "Factor": ["...", "..."]}
      ]
    },
    ...
  ]

Targets:
  - compass.json                               (project root; app fallback)
  - exports/local_20260622_161257/Compass.json (docker import folder)
  - exports/local_20260622_161257/rolNLDraft.json  (table data; full flat rows)
  - exports/local_20260622_161257/_manifest.json

Optional:
  --mongo   also replace Compass (+ optionally rolNLDraft) in MongoDB
"""

from __future__ import annotations

import argparse
import json
import os
import shutil
import sys
from collections import OrderedDict
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
DEFAULT_INPUT = ROOT / "exports" / "compass_dataitems_v2_final.json"
DEFAULT_EXPORT_DIR = ROOT / "exports" / "local_20260622_161257"
DEFAULT_EXPORT_OUT = DEFAULT_EXPORT_DIR / "Compass.json"
DEFAULT_ROL_OUT = DEFAULT_EXPORT_DIR / "rolNLDraft.json"
DEFAULT_MANIFEST = DEFAULT_EXPORT_DIR / "_manifest.json"
DEFAULT_ROOT_OUT = ROOT / "compass.json"
DEFAULT_LOCAL_COMPASS = ROOT / "local.Compass.json"


def load_json(path: Path):
    with path.open("r", encoding="utf-8") as f:
        return json.load(f)


def save_json(path: Path, data) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
        f.write("\n")


def backup(path: Path) -> Path | None:
    if not path.exists():
        return None
    bak = path.with_suffix(path.suffix + ".bak")
    shutil.copy2(path, bak)
    return bak


def build_tree(rows: list) -> list:
    """Flat rows -> Compass docs with Quadrant/children/Segment/Factor."""
    tree: OrderedDict = OrderedDict()
    skipped = 0

    for row in rows:
        if not isinstance(row, dict):
            skipped += 1
            continue
        q = str(row.get("Quadrant") or "").strip()
        s = str(row.get("Segment") or "").strip()
        f = str(row.get("Factor") or "").strip()
        if not q or not s or not f:
            skipped += 1
            continue
        if q not in tree:
            tree[q] = OrderedDict()
        if s not in tree[q]:
            tree[q][s] = OrderedDict()
        tree[q][s][f] = True

    docs = []
    for q, segs in tree.items():
        children = [
            {"Segment": s, "Factor": list(factors.keys())}
            for s, factors in segs.items()
        ]
        docs.append({"Quadrant": q, "children": children})

    return docs, skipped


def attach_ids_from_existing(new_docs: list, existing: list) -> list:
    """Reuse _id when quadrant names match (normalized)."""

    def norm(name: str) -> str:
        return (
            (name or "")
            .replace("&", "and")
            .replace("’", "'")
            .replace("'", "'")
            .lower()
            .strip()
        )

    id_by_q = {}
    for doc in existing or []:
        if isinstance(doc, dict) and doc.get("Quadrant") and "_id" in doc:
            id_by_q[norm(doc["Quadrant"])] = doc["_id"]

    out = []
    for doc in new_docs:
        oid = id_by_q.get(norm(doc["Quadrant"]))
        if oid is not None:
            out.append({"_id": oid, **doc})
        else:
            out.append(doc)
    return out


def validate(docs: list) -> None:
    if not docs:
        raise SystemExit("ERROR: produced 0 Compass documents")
    for i, doc in enumerate(docs):
        if "Quadrant" not in doc or "children" not in doc:
            raise SystemExit(f"ERROR: doc[{i}] missing Quadrant/children")
        if not isinstance(doc["children"], list):
            raise SystemExit(f"ERROR: doc[{i}].children must be a list")
        for j, child in enumerate(doc["children"]):
            if set(child.keys()) != {"Segment", "Factor"}:
                raise SystemExit(
                    f"ERROR: doc[{i}].children[{j}] keys must be Segment+Factor, got {set(child.keys())}"
                )
            if not isinstance(child["Factor"], list) or not child["Factor"]:
                raise SystemExit(f"ERROR: doc[{i}].children[{j}].Factor must be non-empty list")


def push_mongo(compass_docs: list, rol_docs: list | None, mongo_uri: str) -> None:
    try:
        from pymongo import MongoClient
    except ImportError:
        raise SystemExit("ERROR: pymongo not installed. pip install pymongo")

    client = MongoClient(mongo_uri, serverSelectionTimeoutMS=5000)
    client.admin.command("ping")
    db = client["local"]

    compass_clean = [{k: v for k, v in d.items() if k != "_id"} for d in compass_docs]
    deleted = db["Compass"].delete_many({})
    if compass_clean:
        db["Compass"].insert_many(compass_clean)
    print(f"MongoDB: deleted {deleted.deleted_count}, inserted {len(compass_clean)} into local.Compass")

    if rol_docs is not None:
        rol_clean = [{k: v for k, v in d.items() if k != "_id"} for d in rol_docs]
        deleted_r = db["rolNLDraft"].delete_many({})
        if rol_clean:
            db["rolNLDraft"].insert_many(rol_clean)
        print(
            f"MongoDB: deleted {deleted_r.deleted_count}, inserted {len(rol_clean)} into local.rolNLDraft"
        )

    client.close()


def build_flat_rows(rows: list) -> list:
    flat = []
    for row in rows:
        if not isinstance(row, dict):
            continue
        doc = dict(row)
        for key in [
            "Quadrant",
            "Segment",
            "Factor",
            "Location",
            "Indicator",
            "Title",
            "Url",
            "Datatype",
        ]:
            doc.setdefault(key, "")
        flat.append(doc)
    return flat


def main() -> None:
    parser = argparse.ArgumentParser(description="Build Compass.json from v2 flat data items")
    parser.add_argument(
        "-i",
        "--input",
        type=Path,
        default=DEFAULT_INPUT,
        help="Flat data-items JSON (default: exports/compass_dataitems_v2_final.json)",
    )
    parser.add_argument(
        "--export-out",
        type=Path,
        default=DEFAULT_EXPORT_OUT,
        help="Docker-import Compass.json path",
    )
    parser.add_argument(
        "--root-out",
        type=Path,
        default=DEFAULT_ROOT_OUT,
        help="Project-root compass.json path",
    )
    parser.add_argument(
        "--mongo",
        action="store_true",
        help="Also replace MongoDB local.Compass collection",
    )
    parser.add_argument(
        "--mongo-uri",
        default=os.environ.get("MONGO_URI", "mongodb://127.0.0.1:27017/"),
        help="MongoDB URI (default: env MONGO_URI or mongodb://127.0.0.1:27017/)",
    )
    parser.add_argument(
        "--no-backup",
        action="store_true",
        help="Do not write .bak files before overwrite",
    )
    args = parser.parse_args()

    if not args.input.exists():
        raise SystemExit(f"ERROR: input not found: {args.input}")

    rows = load_json(args.input)
    if not isinstance(rows, list):
        raise SystemExit("ERROR: input must be a JSON array of objects")

    docs, skipped = build_tree(rows)

    # Prefer ids from existing export Compass if present
    existing = []
    if args.export_out.exists():
        try:
            existing = load_json(args.export_out)
        except Exception:
            existing = []
    docs_with_ids = attach_ids_from_existing(docs, existing)
    validate(docs_with_ids)

    # Root compass.json usually has no _id
    docs_no_id = [{k: v for k, v in d.items() if k != "_id"} for d in docs_with_ids]

    print(f"Input rows: {len(rows)}")
    print(f"Skipped rows (missing Quadrant/Segment/Factor): {skipped}")
    print(f"Compass documents: {len(docs_with_ids)}")
    for d in docs_with_ids:
        n_seg = len(d["children"])
        n_fac = sum(len(c["Factor"]) for c in d["children"])
        print(f"  - {d['Quadrant']}: {n_seg} segments, {n_fac} factors")

    targets = [
        (args.export_out, docs_with_ids),
        (args.root_out, docs_no_id),
    ]
    for path, payload in targets:
        if not args.no_backup:
            bak = backup(path)
            if bak:
                print(f"Backup: {bak}")
        save_json(path, payload)
        print(f"Wrote:  {path}")

    if args.mongo:
        push_mongo(docs_no_id, args.mongo_uri)

    print("Done. Schema matches app expectation: Quadrant + children[].Segment + children[].Factor")


if __name__ == "__main__":
    main()
