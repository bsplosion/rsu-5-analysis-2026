"""
Fetch RSU 5 adopted budget documents from rsu5.org for FY22–FY27.

The district publishes budget PDFs through a Finalsite resource manager.
This script downloads the key documents for each fiscal year:

  - Citizens' Adopted Budget  (voter-facing summary)
  - Budget Handbook           (detailed line-item budget)
  - Budget Articles           (warrant articles sent to referendum)
  - Board-Adopted Budget      (articles as adopted by the Board of Directors)

Usage:
    python fetch_budgets.py              # download all FY22–FY27 documents
    python fetch_budgets.py --dry-run    # list documents without downloading
    python fetch_budgets.py --fy 26      # download only FY26 documents
"""

import argparse
import sys
import time
from datetime import datetime
from pathlib import Path

import requests

BASE_URL = "https://www.rsu5.org"
RESOURCE_URL = BASE_URL + "/fs/resource-manager/view/{}"
BUDGET_PAGE_URL = BASE_URL + "/budget/{}"

DATA_DIR = Path(__file__).parent
PDF_DIR = DATA_DIR / "pdf"

# Each entry: (fiscal_year, slug, label, resource_manager_uuid)
# slug becomes the filename: pdf/FY{yy}-{slug}.pdf
# For each FY we capture the final/adopted versions of each document type.
BUDGET_DOCUMENTS = [
    # ── FY22 (2021-2022) ────────────────────────────────────────────────
    (22, "citizens-adopted-budget",
     "FY22 Citizens' Adopted Budget",
     "e9ec20a7-2e8d-44a8-840f-28b79f07dbf8"),
    (22, "superintendent-handbook",
     "2021-2022 Superintendent's Recommended Budget Handbook (Revised Apr 14)",
     "4e366c9c-8e73-4e47-bcb9-e8dc8de17042"),
    (22, "budget-worksheet",
     "2021-2022 Superintendent's Recommended Budget Worksheet (Revised Apr 14)",
     "0c7fb758-fbaf-488b-a99c-9d60cec1a07e"),
    (22, "budget-overview",
     "2021-2022 Superintendent's Recommended Budget Overview (Revised Apr 14)",
     "9001d364-14b2-421e-bef4-100c68b51468"),

    # ── FY23 (2022-2023) ────────────────────────────────────────────────
    (23, "citizens-adopted-budget",
     "FY23 Citizens' Adopted Budget",
     "a83623b7-cabf-4b3b-8061-ab83fac4322f"),
    (23, "board-adopted-budget",
     "FY23 Board of Directors Adopted Budget",
     "07cd8056-2424-48db-8056-8c643c73f8d2"),
    (23, "board-adopted-summary",
     "FY23 Budget Summary Documents for Board of Directors Adopted Budget",
     "9b78848e-5a06-473b-b908-cb45cb9b0aaf"),
    (23, "superintendent-handbook",
     "2022-2023 Superintendent's Recommended Budget Handbook (Revised Mar 23)",
     "7daca974-929e-4454-952a-b7385b7a71d6"),
    (23, "budget-articles",
     "2022-2023 Superintendent's Recommended Budget Articles (Mar 23)",
     "dd48b04b-f4c5-4ed9-9e95-4e276b83ccb1"),

    # ── FY24 (2023-2024) ────────────────────────────────────────────────
    (24, "citizens-adopted-budget",
     "FY24 Citizens' Adopted Budget",
     "9fedddf4-3173-4e93-8fd4-3057f6592662"),
    (24, "board-adopted-budget",
     "FY24 Board of Directors Adopted Budget",
     "e83ed7f2-3c2d-4dca-80c4-778c57c29b99"),
    (24, "board-adopted-handbook",
     "FY24 Board of Directors Adopted Budget Handbook",
     "36544929-6683-42fe-aa26-b6a0212ff34e"),
    (24, "superintendent-handbook",
     "2023-2024 Superintendent's Recommended Budget Handbook (Revised Mar 22)",
     "539660a7-7280-453d-b5b4-1052ed29f4c9"),
    (24, "budget-articles",
     "2023-2024 Superintendent's Recommended Budget Articles (Mar 22)",
     "653eeb67-20db-46e0-be3d-1e41dd535e15"),

    # ── FY25 (2024-2025) ────────────────────────────────────────────────
    (25, "citizens-adopted-budget",
     "FY25 Citizens' Adopted Budget",
     "805aa39f-7658-436d-b81f-e49bf698d591"),
    (25, "board-adopted-budget",
     "FY25 Board of Directors Adopted Budget",
     "56dfda0c-b059-4e64-a9b1-cb83a11ed539"),
    (25, "board-adopted-handbook",
     "FY25 Board of Directors Adopted Budget Handbook",
     "55d08dc3-9994-4b26-b91e-b8a23bbecc4f"),
    (25, "superintendent-handbook",
     "2024-2025 Superintendent's Recommended Budget Handbook (Revised Mar 27)",
     "5df6821b-6911-49c4-85e9-470863197ef2"),
    (25, "budget-articles",
     "2024-2025 Superintendent's Recommended Budget Articles (Revised Mar 13)",
     "e6bee6f8-1b48-4f60-9124-65e1773ccd6a"),

    # ── FY26 (2025-2026) ────────────────────────────────────────────────
    (26, "citizens-adopted-budget",
     "FY26 Citizens' Adopted Budget",
     "10d5594e-5852-48a9-b360-ca0ea23e71e8"),
    (26, "board-adopted-budget",
     "FY26 Board of Directors Adopted Budget",
     "cf517220-85f6-434e-9070-d30333798b25"),
    (26, "board-adopted-handbook",
     "FY26 Board of Directors Adopted Budget Handbook",
     "05ebdd9a-c0b4-41cf-82c1-30be30922cf4"),
    (26, "superintendent-handbook",
     "2025-2026 Superintendent's Recommended Budget Handbook (Updated 3/19/25)",
     "85a85f5b-822a-4ba1-ad28-35929fce4237"),
    (26, "budget-articles",
     "2025-2026 Superintendent's Recommended Budget Articles",
     "3e6d92e0-ad8f-4d70-8d0a-d8c4a87d1ddb"),

    # ── FY27 (2026-2027) — budget cycle in progress ────────────────────
    (27, "superintendent-handbook",
     "2026-2027 Superintendent's Recommended Budget Handbook (Revised Mar 11)",
     "a3843da0-775f-4d2a-bee4-2f23c831c784"),
    (27, "budget-articles",
     "2026-2027 Superintendent's Recommended Budget Articles (Revised Feb 11)",
     "578b94ee-2505-47e4-bf76-1c9a0845515f"),
    (27, "projected-budgets-fy27-29",
     "Projected FY27, FY28, and FY29 Budgets (Feb 25)",
     "c31d8676-131c-4729-9439-ecb2b7ef96d0"),
    (27, "year-to-year-staffing",
     "Year to Year Staffing Comparison (Mar 11)",
     "4414c97f-782f-4ae9-8bd6-0d6efa2e99d5"),
]


def resolve_pdf_url(viewer_url: str) -> str:
    """Follow the /fs/resource-manager/view/ redirect to the actual file URL."""
    resp = requests.head(viewer_url, allow_redirects=False, timeout=15)
    if resp.status_code in (301, 302, 303, 307, 308):
        return resp.headers["Location"]
    return viewer_url


def download_pdf(url: str, dest: Path) -> bool:
    """Download a PDF to dest. Returns True if newly downloaded."""
    if dest.exists():
        return False
    resp = requests.get(url, timeout=60)
    resp.raise_for_status()
    dest.parent.mkdir(parents=True, exist_ok=True)
    dest.write_bytes(resp.content)
    return True


def main():
    parser = argparse.ArgumentParser(
        description="Fetch RSU 5 budget documents (FY22–FY27)")
    parser.add_argument("--dry-run", action="store_true",
                        help="List documents without downloading")
    parser.add_argument("--fy", type=int, default=None,
                        help="Download only a specific fiscal year (e.g. 26)")
    parser.add_argument("--force", action="store_true",
                        help="Re-download even if file already exists")
    args = parser.parse_args()

    PDF_DIR.mkdir(parents=True, exist_ok=True)

    docs = BUDGET_DOCUMENTS
    if args.fy is not None:
        docs = [d for d in docs if d[0] == args.fy]
        if not docs:
            print(f"No documents found for FY{args.fy}", file=sys.stderr)
            sys.exit(1)

    if args.dry_run:
        current_fy = None
        for fy, slug, label, uuid in docs:
            if fy != current_fy:
                print(f"\n  FY{fy} ({2000 + fy - 1}-{2000 + fy}):")
                current_fy = fy
            viewer = RESOURCE_URL.format(uuid)
            print(f"    {slug:30s}  {label}")
            print(f"      {viewer}")
        return

    downloaded = 0
    skipped = 0
    errors = []

    current_fy = None
    for i, (fy, slug, label, uuid) in enumerate(docs):
        if fy != current_fy:
            print(f"\nFY{fy} ({2000 + fy - 1}-{2000 + fy}):")
            current_fy = fy

        tag = f"[{i + 1}/{len(docs)}]"
        filename = f"FY{fy}-{slug}.pdf"
        dest = PDF_DIR / filename
        viewer_url = RESOURCE_URL.format(uuid)

        if dest.exists() and not args.force:
            print(f"  {tag} {filename} — already exists, skipping")
            skipped += 1
            continue

        if args.force and dest.exists():
            dest.unlink()

        try:
            pdf_url = resolve_pdf_url(viewer_url)
            print(f"  {tag} Downloading {label}...")
            print(f"       -> {filename}")
            download_pdf(pdf_url, dest)
            size_kb = dest.stat().st_size / 1024
            print(f"       {size_kb:.0f} KB")
            downloaded += 1
            time.sleep(0.5)
        except Exception as e:
            print(f"  {tag} ERROR: {e}", file=sys.stderr)
            errors.append((filename, str(e)))

    print(f"\nDone. {downloaded} downloaded, {skipped} already existed, "
          f"{len(errors)} errors.")
    if errors:
        print("Errors:")
        for name, err in errors:
            print(f"  {name}: {err}")
    print(f"Output: {PDF_DIR}")


if __name__ == "__main__":
    main()
