"""
Fetch RSU 5 Board of Directors meeting minutes from rsu5.org,
download the PDFs, and OCR them to Markdown.

The minutes are published as scanned image PDFs (no embedded text),
so we render each page with PyMuPDF and OCR with Tesseract.

Usage:
    python fetch_minutes.py                    # last 2 school years
    python fetch_minutes.py --years 3          # last 3 school years
    python fetch_minutes.py --after 2024-01-01 # everything after a date
    python fetch_minutes.py --single <url>     # process a single PDF URL
    python fetch_minutes.py --reocr            # re-OCR already-downloaded PDFs
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

import pymupdf
import pytesseract
import requests
from PIL import Image

TESSERACT_CMD = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
MINUTES_URL = "https://www.rsu5.org/board-of-directors-and-policies/board-agendas-and-minutes/minutes"
ARCHIVED_URL = "https://www.rsu5.org/fs/pages/1279"
BASE_URL = "https://www.rsu5.org"

DATA_DIR = Path(__file__).parent
PDF_DIR = DATA_DIR / "pdf"
MD_DIR = DATA_DIR / "md"

MONTH_NAMES = {
    "january": 1, "february": 2, "march": 3, "april": 4,
    "may": 5, "june": 6, "july": 7, "august": 8,
    "september": 9, "october": 10, "november": 11, "december": 12,
}

DATE_PATTERN = re.compile(
    r"(January|February|March|April|May|June|July|August|"
    r"September|October|November|December)\s+(\d{1,2}),?\s+(\d{4})",
    re.IGNORECASE,
)


def parse_meeting_date(text: str) -> date | None:
    m = DATE_PATTERN.search(text)
    if not m:
        return None
    month = MONTH_NAMES[m.group(1).lower()]
    day = int(m.group(2))
    year = int(m.group(3))
    try:
        return date(year, month, day)
    except ValueError:
        return None


def _scrape_page(url: str) -> list[tuple[str, str, date]]:
    """Scrape a single page for minutes links."""
    resp = requests.get(url, timeout=30)
    resp.raise_for_status()

    link_pattern = re.compile(
        r'<a[^>]+href="([^"]+)"[^>]*>'
        r"([^<]*(?:January|February|March|April|May|June|July|August|"
        r"September|October|November|December)[^<]*)</a>",
        re.IGNORECASE,
    )

    results = []
    for href, label in link_pattern.findall(resp.text):
        label = label.strip()
        d = parse_meeting_date(label)
        if d is None:
            continue
        url_full = urljoin(BASE_URL, href)
        results.append((label, url_full, d))
    return results


def scrape_minutes_links(include_archived: bool = True) -> list[tuple[str, str, date]]:
    """Return list of (label, url, date) from the minutes page(s)."""
    results = _scrape_page(MINUTES_URL)
    if include_archived:
        results.extend(_scrape_page(ARCHIVED_URL))

    seen: set[date] = set()
    deduped = []
    for label, url, d in results:
        if d not in seen:
            seen.add(d)
            deduped.append((label, url, d))

    deduped.sort(key=lambda x: x[2], reverse=True)
    return deduped


def resolve_pdf_url(viewer_url: str) -> str:
    """Follow the /fs/resource-manager/view/ redirect to get the actual PDF URL."""
    resp = requests.head(viewer_url, allow_redirects=False, timeout=15)
    if resp.status_code in (301, 302, 303, 307, 308):
        return resp.headers["Location"]
    return viewer_url


def download_pdf(pdf_url: str, dest: Path) -> bool:
    if dest.exists():
        return False
    resp = requests.get(pdf_url, timeout=60)
    resp.raise_for_status()
    dest.parent.mkdir(parents=True, exist_ok=True)
    dest.write_bytes(resp.content)
    return True


def ocr_pdf_to_markdown(
    pdf_path: Path,
    md_path: Path,
    meeting_label: str,
    meeting_date: date,
    viewer_url: str = "",
    pdf_url: str = "",
    retrieved_date: str = "",
) -> str:
    """OCR a scanned PDF and write structured Markdown. Returns the text."""
    pytesseract.pytesseract.tesseract_cmd = TESSERACT_CMD

    doc = pymupdf.open(str(pdf_path))
    page_texts = []

    for page in doc:
        pix = page.get_pixmap(dpi=300)
        img = Image.open(io.BytesIO(pix.tobytes("png")))
        text = pytesseract.image_to_string(img, lang="eng", config="--psm 6")
        page_texts.append(text.strip())

    raw = "\n\n".join(page_texts)
    cleaned = postprocess_ocr(raw)

    if not retrieved_date:
        retrieved_date = datetime.now().strftime("%Y-%m-%d")

    header = (
        f"# RSU 5 Board of Directors — Meeting Minutes\n"
        f"**Date:** {meeting_date.strftime('%B %d, %Y')}  \n"
        f"**Source:** [{meeting_label}]({viewer_url})  \n"
        f"**PDF:** [{pdf_url.rsplit('/', 1)[-1] if pdf_url else 'N/A'}]({pdf_url})  \n"
        f"**Retrieved:** {retrieved_date}  \n"
        f"**OCR processed:** {datetime.now().strftime('%Y-%m-%d %H:%M')}  \n"
        f"\n---\n\n"
    )

    md_content = header + cleaned
    md_path.parent.mkdir(parents=True, exist_ok=True)
    md_path.write_text(md_content, encoding="utf-8")
    return md_content


def postprocess_ocr(text: str) -> str:
    """Fix common OCR artifacts from scanned board minutes."""
    text = text.replace("\u00e2\u0080\u0099", "'")
    text = text.replace("\u2019", "'")
    text = text.replace("\ufffd", "'")

    # Fix superscript ordinals: "2™ Read" -> "2nd Read", "1* Read" -> "1st Read"
    text = re.sub(r"(\d)\s*[°™�]\s*(Read|read)", r"\1nd \2", text)
    text = re.sub(r"(\d)\s*\*\s*(Read|read)", r"\1st \2", text)

    # Fix "Superintendent" variants
    text = re.sub(r"\bSuperi[ñnjr]?[stx]?i?te[hn]?dent\b", "Superintendent", text, flags=re.IGNORECASE)
    text = re.sub(r"\bSuperi\s*x\s*endent\b", "Superintendent", text, flags=re.IGNORECASE)

    # Fix bullet points: lone "e " at start of line (OCR of ● character) -> "- "
    text = re.sub(r"(?m)^(\s*)e\s+((?:[A-Z]|Report|Anna|shared|The ))", r"\1- \2", text)

    # Fix garbled signature block at the end
    text = re.sub(
        r"\n[q=\s+—\-\w]*\n?\s*Tom\s+Gray[,]?\s+Superintendent\s+of\s+Schools\s*$",
        "\n\nTom Gray, Superintendent of Schools\n",
        text, flags=re.DOTALL,
    )
    # Strip garbled lines near the signature area (OCR of handwritten signatures)
    text = re.sub(r"\n[=\-\s\w]{3,20}=\s*\n", "\n", text)

    # Fallback for heavily garbled signatures
    text = re.sub(
        r"\n[^\n]*Tom\s+\w+[,.]?\s+Super.*$",
        "\n\nTom Gray, Superintendent of Schools\n",
        text, flags=re.DOTALL,
    )

    # Fix comma/period confusion on item numbers: "9," -> "9."
    text = re.sub(r"(?m)^(\d{1,2}),(\s+[A-Z])", r"\1.\2", text)

    # Normalize multiple blank lines
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()


def date_filename(d: date) -> str:
    return d.strftime("%Y-%m-%d")


def school_year_start(d: date) -> int:
    """Return the calendar year in which the school year started (July cutoff)."""
    return d.year if d.month >= 7 else d.year - 1


def main():
    parser = argparse.ArgumentParser(description="Fetch and OCR RSU 5 board minutes")
    parser.add_argument("--years", type=int, default=2,
                        help="Number of school years to fetch (default: 2)")
    parser.add_argument("--after", type=str, default=None,
                        help="Only fetch minutes after this date (YYYY-MM-DD)")
    parser.add_argument("--single", type=str, default=None,
                        help="Process a single viewer URL")
    parser.add_argument("--reocr", action="store_true",
                        help="Re-OCR already-downloaded PDFs (skip download)")
    parser.add_argument("--dry-run", action="store_true",
                        help="List what would be fetched without downloading")
    args = parser.parse_args()

    PDF_DIR.mkdir(parents=True, exist_ok=True)
    MD_DIR.mkdir(parents=True, exist_ok=True)

    if args.single:
        label = "Single PDF"
        viewer_url = args.single
        pdf_url = resolve_pdf_url(viewer_url)
        fname = pdf_url.rsplit("/", 1)[-1]
        pdf_path = PDF_DIR / fname
        md_name = fname.replace(".pdf", ".md")
        md_path = MD_DIR / md_name

        print(f"Downloading {pdf_url}")
        download_pdf(pdf_url, pdf_path)
        retrieved = datetime.now().strftime("%Y-%m-%d")
        print(f"OCR -> {md_path}")
        ocr_pdf_to_markdown(
            pdf_path, md_path, label, date.today(),
            viewer_url=viewer_url, pdf_url=pdf_url, retrieved_date=retrieved,
        )
        print("Done.")
        return

    print("Scraping minutes page...")
    links = scrape_minutes_links()
    print(f"  Found {len(links)} meeting minutes links")

    if args.after:
        cutoff = date.fromisoformat(args.after)
        links = [(label, url, d) for label, url, d in links if d >= cutoff]
    else:
        today = date.today()
        current_sy = school_year_start(today)
        earliest_sy = current_sy - args.years + 1
        links = [
            (label, url, d) for label, url, d in links
            if school_year_start(d) >= earliest_sy
        ]

    print(f"  {len(links)} meetings in scope\n")

    if args.dry_run:
        for label, url, d in links:
            print(f"  {d.isoformat()}  {label}")
        return

    for i, (label, viewer_url, d) in enumerate(links):
        prefix = date_filename(d)

        pdf_path = PDF_DIR / f"{prefix}.pdf"
        md_path = MD_DIR / f"{prefix}.md"

        tag = f"[{i+1}/{len(links)}]"

        if md_path.exists() and not args.reocr:
            print(f"  {tag} {label} — already processed, skipping")
            continue

        pdf_url = ""
        retrieved = datetime.now().strftime("%Y-%m-%d")

        if not pdf_path.exists():
            try:
                pdf_url = resolve_pdf_url(viewer_url)
                print(f"  {tag} Downloading {label}...")
                download_pdf(pdf_url, pdf_path)
                time.sleep(0.5)
            except Exception as e:
                print(f"  {tag} ERROR downloading {label}: {e}", file=sys.stderr)
                continue
        else:
            print(f"  {tag} PDF exists: {pdf_path.name}")
            if not pdf_url:
                try:
                    pdf_url = resolve_pdf_url(viewer_url)
                except Exception:
                    pdf_url = viewer_url

        try:
            print(f"       OCR -> {md_path.name}")
            ocr_pdf_to_markdown(
                pdf_path, md_path, label, d,
                viewer_url=viewer_url, pdf_url=pdf_url, retrieved_date=retrieved,
            )
        except Exception as e:
            print(f"  {tag} ERROR during OCR of {label}: {e}", file=sys.stderr)
            continue

    print(f"\nDone. {len(list(MD_DIR.glob('*.md')))} markdown files in {MD_DIR}")


if __name__ == "__main__":
    main()
