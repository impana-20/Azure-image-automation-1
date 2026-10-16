"""
Unified Azure Blob URL mapper for the MERGED Alpha_product_list.xlsx workbook.

Place this script in project-root/scripts/map_merged_excel_azure_urls.py.
Run from project root:
    python scripts/map_merged_excel_azure_urls.py

It reads the single-sheet merged workbook Alpha_product_list.xlsx, inventories actual
image blobs in Azure container `nivora`, and writes:
    output/Alpha_product_list_with_azure_image_urls.xlsx
    output/azure_image_url_mapping_report.csv

Every assigned hyperlink displays exactly "Open Image". All URLs generated in one run
share one SAS expiry timestamp. Existing URLs in the Image URL column are replaced with
fresh links when a confident real blob match is found. The original workbook is preserved.
Rows with ambiguous/unmatched images are recorded in the report; this script never invents
a blob name or URL. Complete coverage requires a real image blob for every product row.
"""
import csv
import os
import re
import sys
from collections import defaultdict
from datetime import datetime, timedelta, timezone
from difflib import SequenceMatcher
from pathlib import Path
from urllib.parse import unquote, urlparse

from dotenv import load_dotenv
from openpyxl import load_workbook
from openpyxl.styles import Font
from azure.storage.blob import BlobServiceClient, BlobSasPermissions, generate_blob_sas

ROOT = Path(__file__).resolve().parent.parent
ENV_PATH = ROOT / ".env"
INPUT_XLSX = ROOT / "Alpha_product_list.xlsx"
OUTPUT_DIR = ROOT / "output"
OUTPUT_XLSX = OUTPUT_DIR / "Alpha_product_list_with_azure_image_urls.xlsx"
REPORT_CSV = OUTPUT_DIR / "azure_image_url_mapping_report.csv"

CONTAINER = "nivora"
AZURE_ROOT = "Nivora Products/categories"
SAS_DAYS = 365
IMAGE_EXTS = {".jpg", ".jpeg", ".png", ".webp", ".gif", ".bmp", ".tif", ".tiff"}
DISPLAY_TEXT = "Open Image"

# Excel category value -> actual Azure category folder.
CATEGORY_MAP = {
    "BATH & KITCHEN ESSENTIALS": "BATH & KITCHEN ESSENTIALS",
    "ELECTRICAL": "ELECTRICAL",
    "HARDWARE": "HARDWARE",
    "INTERIOR DESIGNING": "INTERIOR DESIGNING",
    "PLUMBING": "PLUMBING",
    "SANITARY": "SANITARY",
}


def norm(value):
    if value is None:
        return ""
    s = str(value).strip().lower().replace("&", " and ")
    s = re.sub(r"\be\.?\s*w\.?\s*c\.?\b", " ewc ", s)
    s = re.sub(r"[^a-z0-9]+", " ", s)
    return re.sub(r"\s+", " ", s).strip()

def compact(value):
    return re.sub(r"[^a-z0-9]", "", norm(value))

def tokens(value):
    ignore = {"the", "and", "of", "for", "with", "product", "products", "item"}
    return {t for t in norm(value).split() if len(t) > 1 and t not in ignore}

def header_indices(ws):
    return {norm(ws.cell(1, c).value): c for c in range(1, ws.max_column + 1)
            if ws.cell(1, c).value is not None}

def find_col(headers, *labels):
    for label in labels:
        if norm(label) in headers:
            return headers[norm(label)]
    return None

def cell_text(ws, row, headers, *labels):
    col = find_col(headers, *labels)
    return ws.cell(row, col).value if col else None

def product_description(ws, row, headers):
    labels = (
        ("Sub category", "Subcategory"),
        ("Item type",),
        ("Item Subtype 1",),
        ("Item Subtype 2",),
        ("Item Name",),
        ("Item Description",),
        ("Item Color",),
    )
    values = []
    for aliases in labels:
        value = cell_text(ws, row, headers, *aliases)
        if value not in (None, ""):
            values.append(str(value))
    return " ".join(values)

def parse_connection(connection):
    parts = {}
    for piece in connection.split(";"):
        if "=" in piece:
            k, v = piece.split("=", 1)
            parts[k.strip().lower()] = v.strip()
    return parts.get("accountname", ""), parts.get("accountkey", ""), parts.get("sharedaccesssignature", "")

def make_sas_url(blob_client, account_name, account_key, common_expiry, fallback_sas):
    if account_key:
        sas = generate_blob_sas(
            account_name=account_name,
            container_name=CONTAINER,
            blob_name=blob_client.blob_name,
            account_key=account_key,
            permission=BlobSasPermissions(read=True),
            expiry=common_expiry,
            protocol="https",
        )
        return f"{blob_client.url}?{sas}"
    if fallback_sas:
        return f"{blob_client.url}?{fallback_sas.lstrip('?')}"
    raise RuntimeError(
        "Cannot create usable read URLs: connection string has no AccountKey and no SAS token. "
        "Use your team's approved SAS or Entra ID workflow."
    )

def score_candidate(product_text, blob_name, category_prefix):
    relative = blob_name[len(category_prefix):] if blob_name.startswith(category_prefix) else blob_name
    stem = Path(relative).stem
    parent = Path(relative).parent.name
    product_norm = norm(product_text)
    stem_norm = norm(stem)
    parent_norm = norm(parent)

    # Strong exact phrase matches first.
    if stem_norm and re.search(r"(?<![a-z0-9])" + re.escape(stem_norm) + r"(?![a-z0-9])", product_norm):
        return 1.0
    if parent_norm and len(parent_norm) >= 4 and re.search(r"(?<![a-z0-9])" + re.escape(parent_norm) + r"(?![a-z0-9])", product_norm):
        return 0.98

    # Compare identifying tokens; penalize conflicting product variants such as Indian/Western.
    p_tokens, b_tokens = tokens(product_text), tokens(stem)
    if not p_tokens or not b_tokens:
        return 0.0
    overlap = len(p_tokens & b_tokens)
    precision = overlap / len(b_tokens)
    recall = overlap / len(p_tokens)
    dice = 2 * precision * recall / (precision + recall) if precision + recall else 0.0
    p_norm, b_norm = compact(product_text), compact(stem)
    sequence = SequenceMatcher(None, p_norm, b_norm).ratio()
    score = 0.78 * dice + 0.22 * sequence

    variant_groups = [
        {"indian", "western"},
        {"single", "double", "triple"},
        {"left", "right"},
        {"hot", "cold"},
        {"white", "black", "ivory", "cream"},
    ]
    for group in variant_groups:
        if (p_tokens & group) and (b_tokens & group) and not ((p_tokens & group) & (b_tokens & group)):
            score -= 0.25
    return max(0.0, min(1.0, score))

def set_open_image_hyperlink(cell, url):
    cell.value = DISPLAY_TEXT
    cell.hyperlink = url
    cell.style = "Hyperlink"
    cell.font = Font(name=cell.font.name or "Calibri", size=cell.font.sz or 11,
                     bold=cell.font.bold, italic=cell.font.italic,
                     color="0563C1", underline="single")

def main():
    if not ENV_PATH.exists():
        raise FileNotFoundError(f"Missing project-root .env: {ENV_PATH}")
    if not INPUT_XLSX.exists():
        raise FileNotFoundError(
            f"Base workbook not found: {INPUT_XLSX}. Place Alpha_product_list.xlsx beside .env."
        )

    load_dotenv(ENV_PATH, override=True)
    connection = os.getenv("AZURE_STORAGE_CONNECTION_STRING", "").strip()
    extra_sas = os.getenv("AZURE_BLOB_SAS_TOKEN", "").strip()
    if not connection:
        raise RuntimeError("AZURE_STORAGE_CONNECTION_STRING is missing from project-root .env.")

    service = BlobServiceClient.from_connection_string(connection)
    container_client = service.get_container_client(CONTAINER)
    account_name, account_key, connection_sas = parse_connection(connection)
    fallback_sas = extra_sas or connection_sas
    # One expiry timestamp for every generated URL in this workbook run.
    common_expiry = datetime.now(timezone.utc) + timedelta(days=SAS_DAYS)

    inventory = defaultdict(list)
    print("Scanning actual Azure image blobs...")
    for category in sorted(set(CATEGORY_MAP.values())):
        prefix = f"{AZURE_ROOT}/{category}/"
        for blob in container_client.list_blobs(name_starts_with=prefix):
            if Path(blob.name).suffix.lower() in IMAGE_EXTS:
                inventory[category].append(blob.name)
        print(f"  {category}: {len(inventory[category])} image blobs")

    wb = load_workbook(INPUT_XLSX)
    ws = wb.active
    headers = header_indices(ws)
    url_col = find_col(headers, "Image URL", "Image Url", "Image URLs", "Image Link")
    category_col = find_col(headers, "Category")
    code_col = find_col(headers, "Item Code")
    name_col = find_col(headers, "Item Name")
    if not url_col or not category_col:
        raise RuntimeError("Workbook must have 'Category' and 'Image URL' columns in row 1.")

    counts = defaultdict(int)
    report = []
    rows_total = max(0, ws.max_row - 1)

    for row in range(2, ws.max_row + 1):
        raw_category = str(ws.cell(row, category_col).value or "").strip()
        category = CATEGORY_MAP.get(raw_category.upper())
        code = ws.cell(row, code_col).value if code_col else ""
        name = ws.cell(row, name_col).value if name_col else ""
        current_cell = ws.cell(row, url_col)
        previous_url = current_cell.hyperlink.target if current_cell.hyperlink else (
            str(current_cell.value).strip() if current_cell.value and str(current_cell.value).strip().lower().startswith(("http://", "https://")) else ""
        )

        status, matched_blob, score, note = "", "", "", ""
        candidates = []

        if not category:
            status = "CATEGORY_UNKNOWN"
            note = f"Unrecognized category value: {raw_category}"
        elif not inventory[category]:
            status = "NO_AZURE_IMAGES_FOR_CATEGORY"
            note = f"No actual image blobs found under {AZURE_ROOT}/{category}/"
        else:
            product_text = product_description(ws, row, headers)
            prefix = f"{AZURE_ROOT}/{category}/"
            ranked = sorted(
                ((score_candidate(product_text, blob, prefix), blob) for blob in inventory[category]),
                key=lambda item: item[0], reverse=True
            )
            best_score, best_blob = ranked[0]
            second_score = ranked[1][0] if len(ranked) > 1 else 0.0

            # Conservative auto-match threshold. Review report for all unresolved rows.
            if best_score >= 0.58 and (best_score - second_score >= 0.10 or best_score >= 0.98):
                blob_client = container_client.get_blob_client(best_blob)
                fresh_url = make_sas_url(blob_client, account_name, account_key, common_expiry, fallback_sas)
                set_open_image_hyperlink(current_cell, fresh_url)
                status, matched_blob, score = "MATCHED", best_blob, round(best_score, 3)
                note = "Assigned actual Azure blob URL as an 'Open Image' hyperlink."
            else:
                # Replace an old link only when a confident replacement exists. Do not leave
                # an expired/stale link that looks valid; clear it and report for manual review.
                current_cell.value = None
                current_cell.hyperlink = None
                status = "AMBIGUOUS" if best_score >= 0.42 else "UNMATCHED"
                score = round(best_score, 3)
                note = "No unique confident match; Image URL cell left blank for review."
                candidates = [f"{s:.3f}:{b}" for s, b in ranked[:5]]
        counts[status] += 1
        report.append({
            "excel_row": row,
            "item_code": code or "",
            "category": raw_category,
            "item_name": name or "",
            "status": status,
            "matched_blob": matched_blob,
            "score": score,
            "candidate_blobs": " || ".join(candidates),
            "previous_url_present": "yes" if previous_url else "no",
            "notes": note,
        })

    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    wb.save(OUTPUT_XLSX)
    with REPORT_CSV.open("w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=[
            "excel_row", "item_code", "category", "item_name", "status", "matched_blob",
            "score", "candidate_blobs", "previous_url_present", "notes"
        ])
        writer.writeheader()
        writer.writerows(report)

    matched = counts["MATCHED"]
    unresolved = rows_total - matched
    print("\n========== SUMMARY ==========")
    print(f"Product rows: {rows_total}")
    print(f"Matched with fresh URL hyperlink: {matched}")
    print(f"Unresolved (review report): {unresolved}")
    for key in ("CATEGORY_UNKNOWN", "NO_AZURE_IMAGES_FOR_CATEGORY", "AMBIGUOUS", "UNMATCHED"):
        if counts[key]:
            print(f"{key}: {counts[key]}")
    print(f"Common SAS expiry (UTC): {common_expiry.isoformat()}")
    print(f"Output workbook: {OUTPUT_XLSX}")
    print(f"Audit report: {REPORT_CSV}")
    print("Original workbook was not overwritten.")
    print("No URL is fabricated; unmatched rows are listed in the report.")

if __name__ == "__main__":
    try:
        main()
    except Exception as exc:
        print(f"ERROR: {type(exc).__name__}: {exc}", file=sys.stderr)
        sys.exit(1)
