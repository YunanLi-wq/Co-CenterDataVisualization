#!/bin/sh
set -e

MONGO_URI="${MONGO_URI:-mongodb://mongo:27017/}"
DATASET_IMPORT_PATH="${DATASET_IMPORT_PATH:-}"
DATASET_IMPORT_REPLACE="${DATASET_IMPORT_REPLACE:-true}"
INIT_ADMIN="${INIT_ADMIN:-false}"

echo "Waiting for MongoDB at ${MONGO_URI}..."
until python -c "
from pymongo import MongoClient
import sys
try:
    MongoClient('${MONGO_URI}', serverSelectionTimeoutMS=2000).admin.command('ping')
except Exception:
    sys.exit(1)
"; do
  sleep 2
done
echo "MongoDB is ready."

if [ -n "$DATASET_IMPORT_PATH" ] && [ -d "$DATASET_IMPORT_PATH" ]; then
  echo "Importing dataset from ${DATASET_IMPORT_PATH}..."
  if [ "$DATASET_IMPORT_REPLACE" = "true" ]; then
    python dataset_io.py import "$DATASET_IMPORT_PATH" --folder --replace
  else
    python dataset_io.py import "$DATASET_IMPORT_PATH" --folder
  fi
fi

if [ "$INIT_ADMIN" = "true" ]; then
  echo "Ensuring default admin user exists..."
  python init_admin.py || true
fi

echo "Starting Flask application..."
exec python app.py
