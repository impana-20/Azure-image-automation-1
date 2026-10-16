#!/usr/bin/env python3
"""
Upload images under data_collection/ to the Azure Blob Storage container `nivora`.

Default mode is a dry run. Use --upload to actually upload.
Existing blobs are skipped by default. Add --overwrite only if you intentionally
want matching blobs in Azure to be replaced.

Expected destination:
  nivora/Nivora Products/categories/<Category>/<relative subfolder>/<image>

Credentials are read from AZURE_STORAGE_CONNECTION_STRING in the project-root .env.
Never commit .env or real credentials.
"""
import argparse
import os
import sys
from pathlib import Path

from dotenv import load_dotenv
from azure.core.exceptions import ResourceExistsError
from azure.storage.blob import BlobServiceClient, ContentSettings

PROJECT_DIR = Path(__file__).resolve().parent.parent
DATA_DIR = PROJECT_DIR / "data_collection"
ENV_PATH = PROJECT_DIR / ".env"
CONTAINER_NAME = "nivora"
AZURE_BASE_FOLDER = "Nivora Products/categories"

# Source folder -> category path in Azure. Keep target names consistent with
# the existing Nivora category naming.
CATEGORY_MAP = {
    "BATH & KITCHEN ESSENTIALS": "BATH & KITCHEN ESSENTIALS",
    "ELECTRICAL": "ELECTRICAL",
    "HARDWARE": "HARDWARE",
    "INTERIOR DESIGNING": "INTERIOR DESIGNING",
    "PLUMBING": "PLUMBING",
    "SANITARY": "SANITARY",
}

CONTENT_TYPES = {
    ".jpg": "image/jpeg",
    ".jpeg": "image/jpeg",
    ".png": "image/png",
    ".webp": "image/webp",
    ".gif": "image/gif",
    ".bmp": "image/bmp",
    ".tif": "image/tiff",
    ".tiff": "image/tiff",
}

def main():
    parser = argparse.ArgumentParser(description="Upload data_collection images to Azure Blob Storage.")
    parser.add_argument("--upload", action="store_true", help="Perform uploads; without this flag, only preview.")
    parser.add_argument("--overwrite", action="store_true", help="Replace blobs that already exist (requires --upload).")
    args = parser.parse_args()

    if args.overwrite and not args.upload:
        parser.error("--overwrite can only be used together with --upload")

    load_dotenv(ENV_PATH)
    connection_string = os.getenv("AZURE_STORAGE_CONNECTION_STRING", "").strip()
    if not connection_string:
        sys.exit(f"Missing AZURE_STORAGE_CONNECTION_STRING. Add it to {ENV_PATH} (do not share or commit it).")

    if not DATA_DIR.is_dir():
        sys.exit(f"Data folder not found: {DATA_DIR}")

    service = BlobServiceClient.from_connection_string(connection_string)
    container = service.get_container_client(CONTAINER_NAME)

    planned = []
    missing = []
    for source_name, target_category in CATEGORY_MAP.items():
        source_dir = DATA_DIR / source_name
        if not source_dir.is_dir():
            missing.append(source_name)
            continue
        for file_path in source_dir.rglob("*"):
            if not file_path.is_file():
                continue
            content_type = CONTENT_TYPES.get(file_path.suffix.lower())
            if not content_type:
                continue
            relative_path = file_path.relative_to(source_dir).as_posix()
            blob_name = f"{AZURE_BASE_FOLDER}/{target_category}/{relative_path}"
            planned.append((file_path, blob_name, content_type))

    print(f"Container: {CONTAINER_NAME}")
    print(f"Destination prefix: {AZURE_BASE_FOLDER}/")
    print(f"Image files discovered: {len(planned)}")
    if missing:
        print("Source folders not found (skipped): " + ", ".join(missing))
    print("Mode: " + ("UPLOAD" if args.upload else "DRY RUN — no files will be uploaded"))
    if args.overwrite:
        print("Overwrite: enabled")
    print()

    uploaded = 0
    skipped_existing = 0
    failed = 0

    for index, (file_path, blob_name, content_type) in enumerate(planned, start=1):
        print(f"[{index}/{len(planned)}] {file_path.relative_to(PROJECT_DIR)} -> {CONTAINER_NAME}/{blob_name}")
        if not args.upload:
            continue
        try:
            blob_client = container.get_blob_client(blob_name)
            if not args.overwrite and blob_client.exists():
                skipped_existing += 1
                print("    SKIPPED (already exists)")
                continue
            with file_path.open("rb") as data:
                blob_client.upload_blob(
                    data,
                    overwrite=args.overwrite,
                    content_settings=ContentSettings(content_type=content_type),
                )
            uploaded += 1
        except Exception as exc:
            failed += 1
            print(f"    ERROR: {exc}", file=sys.stderr)

    print("\nSummary")
    print(f"Images discovered: {len(planned)}")
    if args.upload:
        print(f"Uploaded: {uploaded}")
        print(f"Skipped existing: {skipped_existing}")
        print(f"Failed: {failed}")
        if failed:
            sys.exit(1)
    else:
        print("No uploads performed. Review the planned destination paths, then rerun with --upload.")

if __name__ == "__main__":
    main()
