#!/usr/bin/env python3
"""
Export and import MongoDB datasets for the Co-Centre interface.

Examples:
  python dataset_io.py export
  python dataset_io.py export --folder
  python dataset_io.py export --folder -o exports/my_backup
  python dataset_io.py export --all --output exports/full_backup.json
  python dataset_io.py export --collection Researcher
  python dataset_io.py import local.rolNLDraft.json
  python dataset_io.py import exports/my_backup --folder --replace
  python dataset_io.py import exports/full_backup.json --all --replace
"""

import argparse
import json
import sys
from datetime import datetime
from pathlib import Path
from typing import Optional

from bson import json_util
from pymongo import MongoClient

MONGO_URI = "mongodb://127.0.0.1:27017/"
DB_NAME = "local"

COLLECTIONS = [
    "rolNLDraft",
    "Researcher",
    "Compass",
    "Manager",
    "HigherManager",
    "PendingItems",
    "description",
]

DEFAULT_COLLECTION = "rolNLDraft"
EXPORTS_DIR = Path(__file__).resolve().parent / "exports"


def connect():
    client = MongoClient(MONGO_URI)
    return client, client[DB_NAME]


def timestamp() -> str:
    return datetime.now().strftime("%Y%m%d_%H%M%S")


def default_export_path(collection: Optional[str], all_collections: bool) -> Path:
    EXPORTS_DIR.mkdir(parents=True, exist_ok=True)
    if all_collections:
        return EXPORTS_DIR / f"{DB_NAME}_all_{timestamp()}.json"
    name = collection or DEFAULT_COLLECTION
    return EXPORTS_DIR / f"{name}_{timestamp()}.json"


def export_collection(db, collection_name: str) -> list:
    return list(db[collection_name].find())


def default_export_folder() -> Path:
    EXPORTS_DIR.mkdir(parents=True, exist_ok=True)
    return EXPORTS_DIR / f"{DB_NAME}_{timestamp()}"


def write_json(path: Path, data) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8") as f:
        json.dump(data, f, default=json_util.default, indent=2, ensure_ascii=False)


def collection_names(db, selected: Optional[str], all_collections: bool) -> list:
    if selected:
        return [selected]
    if all_collections:
        return COLLECTIONS
    return [DEFAULT_COLLECTION]


def export_collections_to_folder(db, output_dir: Path, names: list) -> int:
    output_dir.mkdir(parents=True, exist_ok=True)
    total = 0
    manifest = {
        "database": DB_NAME,
        "exported_at": datetime.now().isoformat(timespec="seconds"),
        "collections": {},
    }

    for name in names:
        if name not in db.list_collection_names():
            print(f"  skip {name}: collection not found")
            continue
        docs = export_collection(db, name)
        write_json(output_dir / f"{name}.json", docs)
        manifest["collections"][name] = len(docs)
        total += len(docs)
        print(f"  {name}: {len(docs)} documents -> {name}.json")

    write_json(output_dir / "_manifest.json", manifest)
    return total


def cmd_export(args):
    client, db = connect()
    try:
        names = collection_names(db, args.collection, args.all or args.folder)

        if args.folder:
            output_dir = Path(args.output) if args.output else default_export_folder()
            total = export_collections_to_folder(db, output_dir, names)
            print(f"✅ Exported {total} documents to folder {output_dir}")
            return

        output_path = Path(args.output) if args.output else default_export_path(
            args.collection, args.all
        )

        if args.all:
            payload = {}
            for name in names:
                if name not in db.list_collection_names():
                    print(f"  skip {name}: collection not found")
                    continue
                docs = export_collection(db, name)
                payload[name] = docs
                print(f"  {name}: {len(docs)} documents")

            write_json(output_path, payload)
        else:
            collection_name = names[0]
            docs = export_collection(db, collection_name)
            write_json(output_path, docs)
            print(f"  {collection_name}: {len(docs)} documents")

        print(f"✅ Exported to {output_path}")
    except Exception as e:
        print(f"❌ Export failed: {e}")
        sys.exit(1)
    finally:
        client.close()


def load_json(path: Path):
    with path.open("r", encoding="utf-8") as f:
        return json.load(f, object_hook=json_util.object_hook)


def import_documents(db, collection_name: str, documents: list, replace: bool) -> int:
    if not isinstance(documents, list):
        raise ValueError(f"Expected a list of documents for collection '{collection_name}'")

    collection = db[collection_name]
    if replace:
        collection.delete_many({})
        if not documents:
            return 0
        result = collection.insert_many(documents)
        return len(result.inserted_ids)

    inserted = 0
    updated = 0
    for doc in documents:
        doc_id = doc.get("_id")
        if doc_id is None:
            collection.insert_one(doc)
            inserted += 1
        else:
            result = collection.replace_one({"_id": doc_id}, doc, upsert=True)
            if result.matched_count:
                updated += 1
            else:
                inserted += 1

    print(f"    inserted/upserted: {inserted}, updated: {updated}")
    return inserted + updated


def import_from_folder(db, input_dir: Path, replace: bool) -> int:
    json_files = sorted(
        path for path in input_dir.glob("*.json")
        if path.name != "_manifest.json"
    )
    if not json_files:
        print(f"❌ No JSON files found in {input_dir}")
        sys.exit(1)

    total = 0
    for path in json_files:
        collection_name = path.stem
        documents = load_json(path)
        print(f"  {collection_name}:")
        total += import_documents(db, collection_name, documents, replace)

    return total


def cmd_import(args):
    input_path = Path(args.input)
    if not input_path.exists():
        print(f"❌ Path not found: {input_path}")
        sys.exit(1)

    client, db = connect()
    try:
        if args.folder or input_path.is_dir():
            if not input_path.is_dir():
                print(f"❌ Expected a folder: {input_path}")
                sys.exit(1)
            total = import_from_folder(db, input_path, args.replace)
            print(f"✅ Imported {total} documents from folder {input_path}")
            return

        data = load_json(input_path)

        if args.all or isinstance(data, dict) and not args.collection:
            if not isinstance(data, dict):
                print("❌ Full-database import expects a JSON object keyed by collection name.")
                sys.exit(1)

            total = 0
            for collection_name, documents in data.items():
                print(f"  {collection_name}:")
                total += import_documents(db, collection_name, documents, args.replace)
            print(f"✅ Imported {total} documents from {input_path}")
            return

        collection_name = args.collection or DEFAULT_COLLECTION
        if isinstance(data, dict):
            if collection_name in data and isinstance(data[collection_name], list):
                documents = data[collection_name]
            else:
                print(
                    f"❌ JSON object does not contain collection '{collection_name}'. "
                    "Use --all or specify --collection."
                )
                sys.exit(1)
        else:
            documents = data

        count = import_documents(db, collection_name, documents, args.replace)
        print(f"✅ Imported {count} documents into '{collection_name}' from {input_path}")
    except Exception as e:
        print(f"❌ Import failed: {e}")
        sys.exit(1)
    finally:
        client.close()


def build_parser():
    parser = argparse.ArgumentParser(
        description="Export and import MongoDB datasets for the Co-Centre interface."
    )
    subparsers = parser.add_subparsers(dest="command", required=True)

    export_parser = subparsers.add_parser("export", help="Export data from MongoDB to JSON")
    export_parser.add_argument(
        "-o", "--output",
        help="Output JSON file path (default: exports/<collection>_<timestamp>.json)",
    )
    export_parser.add_argument(
        "-c", "--collection",
        help=f"Collection name (default: {DEFAULT_COLLECTION})",
    )
    export_parser.add_argument(
        "--all",
        action="store_true",
        help="Export all known collections into one JSON file",
    )
    export_parser.add_argument(
        "--folder",
        action="store_true",
        help="Export each collection as a separate JSON file in a folder",
    )

    import_parser = subparsers.add_parser("import", help="Import data from JSON into MongoDB")
    import_parser.add_argument("input", help="Input JSON file or folder path")
    import_parser.add_argument(
        "-c", "--collection",
        help=f"Target collection when importing a document list (default: {DEFAULT_COLLECTION})",
    )
    import_parser.add_argument(
        "--all",
        action="store_true",
        help="Import a full-database JSON object keyed by collection name",
    )
    import_parser.add_argument(
        "--folder",
        action="store_true",
        help="Import all JSON files from a folder (one file per collection)",
    )
    import_parser.add_argument(
        "--replace",
        action="store_true",
        help="Delete existing documents in the target collection(s) before import",
    )

    return parser


def main():
    parser = build_parser()
    args = parser.parse_args()

    if args.command == "export":
        cmd_export(args)
    elif args.command == "import":
        cmd_import(args)


if __name__ == "__main__":
    main()
