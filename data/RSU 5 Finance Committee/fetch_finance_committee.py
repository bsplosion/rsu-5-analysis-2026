"""
Fetch RSU 5 Finance Committee documents from rsu5.org:
  - Cost sharing analyses and presentations
  - Finance Committee meeting agendas and minutes
  - Capital plan presentations
  - Other financial documents

Usage:
    python fetch_finance_committee.py                   # download + convert all
    python fetch_finance_committee.py --category cost-sharing  # just cost sharing docs
    python fetch_finance_committee.py --dry-run          # list what would be fetched
    python fetch_finance_committee.py --reocr             # re-process existing PDFs
    python fetch_finance_committee.py --force             # re-download everything
"""

import argparse
import io
import os
import re
import sys
import time
from datetime import datetime, date
from pathlib import Path
from urllib.parse import urljoin

import pdfplumber
import pymupdf
import pytesseract
import requests
from PIL import Image

TESSERACT_CMD = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
FC_URL = "https://www.rsu5.org/board-of-directors-and-policies/board-committees/standing-committees/finance-committee"
BASE_URL = "https://www.rsu5.org"

DATA_DIR = Path(__file__).parent
PDF_DIRS = {
    "cost-sharing": DATA_DIR / "pdf" / "cost-sharing",
    "capital-plans": DATA_DIR / "pdf" / "capital-plans",
    "minutes": DATA_DIR / "pdf" / "minutes",
    "other": DATA_DIR / "pdf" / "other",
}
MD_DIR = DATA_DIR / "md"


COST_SHARING_DOCS = [
    ("2024-01-10 Cost Sharing Presentation",
     "https://www.rsu5.org/fs/resource-manager/view/3ad01542-1cd2-4253-a024-1c43a0b33e65",
     "cost-sharing"),
    ("2023-02-08 Cost Sharing Analysis - Suzan Beaudoin",
     "https://www.rsu5.org/fs/resource-manager/view/9b729fe8-4714-4151-9089-b2685143d291",
     "cost-sharing"),
    ("2020-02-12 Cost Sharing Update from November",
     "https://www.rsu5.org/fs/resource-manager/view/a3d51528-d90d-4e6b-a1f5-7b89934a8043",
     "cost-sharing"),
    ("2020-02-12 Cost Sharing Options 1 2 3",
     "https://www.rsu5.org/fs/resource-manager/view/5d81442a-a346-4818-bea7-1ec6ef1879de",
     "cost-sharing"),
    ("2019-10-23 Cost Sharing Presentation and Handouts",
     "https://www.rsu5.org/fs/resource-manager/view/e796d782-3cd9-40a5-a469-170f4012f863",
     "cost-sharing"),
    ("2019-11-06 Cost Sharing Additional Information",
     "https://www.rsu5.org/fs/resource-manager/view/a0dc7eb1-8193-45c6-8653-bdf87eacb943",
     "cost-sharing"),
]

CAPITAL_PLAN_DOCS = [
    ("2022-10-12 Capital Plan Presentation",
     "https://www.rsu5.org/fs/resource-manager/view/66ff4553-3ef7-4860-aa0b-e9cfe2345ff9",
     "capital-plans"),
    ("2021-10-13 Capital Plan Presentation",
     "https://www.rsu5.org/fs/resource-manager/view/81387fbe-67ca-4542-a7cf-655365e5c5df",
     "capital-plans"),
    ("2020-10-28 Capital Plan Presentation",
     "https://www.rsu5.org/fs/resource-manager/view/0194dc21-c95c-43cf-869e-c5b0e5fad0d0",
     "capital-plans"),
    ("2019-10-23 5-Year Capital Plan",
     "https://www.rsu5.org/fs/resource-manager/view/31d29c73-9881-4bcf-bc4f-0b5348318438",
     "capital-plans"),
    ("2018-10-24 5-Year Capital Plan",
     "https://www.rsu5.org/fs/resource-manager/view/c42593d9-68a8-40c6-9208-01ddec20cb25",
     "capital-plans"),
    ("2017-10-25 5-Year Capital Plan",
     "https://www.rsu5.org/fs/resource-manager/view/46a74027-d811-4a5e-9082-d7598a166584",
     "capital-plans"),
    ("2016-11-30 5-Year Capital Plan",
     "https://www.rsu5.org/fs/resource-manager/view/b76b020d-6291-4bf6-8d66-3d9b37278cb8",
     "capital-plans"),
    ("2015-10-28 5-Year Capital Plan",
     "https://www.rsu5.org/fs/resource-manager/view/d09b8ad9-49be-4967-a6e0-723a186e54b9",
     "capital-plans"),
]

OTHER_DOCS = [
    ("2012-01-25 3-5 Year Financial Plan",
     "https://www.rsu5.org/fs/resource-manager/view/4788dfa5-b9a6-477b-a02f-9183dcd629c8",
     "other"),
    ("2010-04-28 3-5 Year Financial Analysis",
     "https://www.rsu5.org/fs/resource-manager/view/4469816b-0f6f-40fe-aa76-6a135ca78214",
     "other"),
    ("2010-01-27 Transportation Analysis",
     "https://www.rsu5.org/fs/resource-manager/view/c5bbd5c1-9037-44d9-8a6c-4f41bf1736c4",
     "other"),
]


def resolve_pdf_url(viewer_url: str) -> str:
    resp = requests.head(viewer_url, allow_redirects=False, timeout=15)
    if resp.status_code in (301, 302, 303, 307, 308):
        return resp.headers["Location"]
    return viewer_url


def download_pdf(pdf_url: str, dest: Path, force: bool = False) -> bool:
    if dest.exists() and not force:
        return False
    resp = requests.get(pdf_url, timeout=60)
    resp.raise_for_status()
    dest.parent.mkdir(parents=True, exist_ok=True)
    dest.write_bytes(resp.content)
    return True


def extract_text_native(pdf_path: Path) -> str | None:
    """Try native text extraction with pdfplumber. Returns None if mostly images."""
    try:
        with pdfplumber.open(str(pdf_path)) as pdf:
            texts = []
            for page in pdf.pages:
                t = page.extract_text() or ""
                texts.append(t.strip())
            combined = "\n\n".join(t for t in texts if t)
            if len(combined) > 200:
                return combined
    except Exception:
        pass
    return None


def ocr_pdf(pdf_path: Path) -> str:
    """OCR a PDF using PyMuPDF + Tesseract."""
    pytesseract.pytesseract.tesseract_cmd = TESSERACT_CMD
    doc = pymupdf.open(str(pdf_path))
    page_texts = []
    for page in doc:
        pix = page.get_pixmap(dpi=300)
        img = Image.open(io.BytesIO(pix.tobytes("png")))
        text = pytesseract.image_to_string(img, lang="eng", config="--psm 6")
        page_texts.append(text.strip())
    doc.close()
    return "\n\n".join(page_texts)


def pdf_to_markdown(pdf_path: Path, md_path: Path, label: str, viewer_url: str) -> str:
    """Extract text from a PDF (native first, OCR fallback) and write Markdown."""
    text = extract_text_native(pdf_path)
    method = "native text extraction"
    if text is None:
        text = ocr_pdf(pdf_path)
        method = "OCR (Tesseract)"

    text = re.sub(r"\n{3,}", "\n\n", text).strip()

    header = (
        f"# {label}\n\n"
        f"**Source:** [{label}]({viewer_url})  \n"
        f"**Extraction method:** {method}  \n"
        f"**Processed:** {datetime.now().strftime('%Y-%m-%d %H:%M')}  \n"
        f"\n---\n\n"
    )

    md_content = header + text
    md_path.parent.mkdir(parents=True, exist_ok=True)
    md_path.write_text(md_content, encoding="utf-8")
    return md_content


def scrape_minutes_links() -> list[tuple[str, str, str]]:
    """Parse the Finance Committee page for agenda/minutes links."""
    resp = requests.get(FC_URL, timeout=30)
    resp.raise_for_status()

    link_re = re.compile(
        r'<a[^>]+href="([^"]+(?:resource-manager|finalsite)[^"]+)"[^>]*>\s*'
        r'([^<]*)',
        re.IGNORECASE,
    )

    date_re = re.compile(
        r"(January|February|March|April|May|June|July|August|"
        r"September|October|November|December)\s+(\d{1,2}),?\s+(\d{4})",
        re.IGNORECASE,
    )
    month_map = {
        "january": 1, "february": 2, "march": 3, "april": 4,
        "may": 5, "june": 6, "july": 7, "august": 8,
        "september": 9, "october": 10, "november": 11, "december": 12,
    }

    results = []
    seen_urls = set()
    for href, raw_label in link_re.findall(resp.text):
        label = raw_label.strip()
        label_lower = label.lower()
        if "agenda" not in label_lower and "minutes" not in label_lower:
            continue

        url_full = urljoin(BASE_URL, href)
        if url_full in seen_urls:
            continue
        seen_urls.add(url_full)

        m = date_re.search(label)
        if not m:
            continue

        month = month_map[m.group(1).lower()]
        day = int(m.group(2))
        year = int(m.group(3))
        try:
            d = date(year, month, day)
        except ValueError:
            continue

        doc_type = "agenda" if "agenda" in label_lower else "minutes"
        filename = f"FC-{d.isoformat()}-{doc_type}"
        results.append((filename, url_full, "minutes"))

    results.sort(key=lambda x: x[0], reverse=True)
    return results


def process_doc(label: str, viewer_url: str, category: str,
                force: bool = False, reocr: bool = False) -> bool:
    """Download and convert a single document. Returns True if processed."""
    safe_name = re.sub(r'[^\w\s\-]', '', label).strip().replace(' ', '_')
    pdf_dir = PDF_DIRS[category]
    pdf_path = pdf_dir / f"{safe_name}.pdf"
    md_path = MD_DIR / f"{safe_name}.md"

    if md_path.exists() and not reocr and not force:
        return False

    if not pdf_path.exists() or force:
        try:
            pdf_url = resolve_pdf_url(viewer_url)
            download_pdf(pdf_url, pdf_path, force=force)
            time.sleep(0.3)
        except Exception as e:
            print(f"  ERROR downloading {label}: {e}", file=sys.stderr)
            return False

    try:
        pdf_to_markdown(pdf_path, md_path, label, viewer_url)
        return True
    except Exception as e:
        print(f"  ERROR converting {label}: {e}", file=sys.stderr)
        return False


def main():
    parser = argparse.ArgumentParser(description="Fetch RSU 5 Finance Committee documents")
    parser.add_argument("--category", choices=["cost-sharing", "minutes", "capital-plans", "other", "all"],
                        default="all", help="Category to fetch (default: all)")
    parser.add_argument("--dry-run", action="store_true", help="List what would be fetched")
    parser.add_argument("--force", action="store_true", help="Re-download existing files")
    parser.add_argument("--reocr", action="store_true", help="Re-process existing PDFs")
    args = parser.parse_args()

    for d in PDF_DIRS.values():
        d.mkdir(parents=True, exist_ok=True)
    MD_DIR.mkdir(parents=True, exist_ok=True)

    docs_to_process = []

    if args.category in ("cost-sharing", "all"):
        docs_to_process.extend(COST_SHARING_DOCS)

    if args.category in ("minutes", "all"):
        print("Scraping Finance Committee page for minutes/agendas...")
        minutes = scrape_minutes_links()
        print(f"  Found {len(minutes)} agenda/minutes links")
        docs_to_process.extend(minutes)

    if args.category in ("capital-plans", "all"):
        docs_to_process.extend(CAPITAL_PLAN_DOCS)

    if args.category in ("other", "all"):
        docs_to_process.extend(OTHER_DOCS)

    print(f"\n{len(docs_to_process)} documents to process\n")

    if args.dry_run:
        for label, url, cat in docs_to_process:
            print(f"  [{cat}] {label}")
        return

    processed = 0
    skipped = 0
    errors = 0

    for i, (label, url, category) in enumerate(docs_to_process):
        tag = f"[{i+1}/{len(docs_to_process)}]"
        safe = re.sub(r'[^\w\s\-]', '', label).strip().replace(' ', '_')
        md_path = MD_DIR / f"{safe}.md"

        if md_path.exists() and not args.reocr and not args.force:
            skipped += 1
            continue

        print(f"  {tag} {label}...")
        ok = process_doc(label, url, category, force=args.force, reocr=args.reocr)
        if ok:
            processed += 1
        else:
            errors += 1

    print(f"\nDone. Processed: {processed}, Skipped: {skipped}, Errors: {errors}")
    print(f"Markdown files: {len(list(MD_DIR.glob('*.md')))}")


if __name__ == "__main__":
    main()
