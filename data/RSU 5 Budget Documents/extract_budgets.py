"""
Extract content from RSU 5 budget PDFs into markdown and CSV files.

  md/   -- full text content with inline pipe-tables, citation headers,
           page anchors, and printer noise stripped
  csv/  -- tabular data extracted via column-position analysis

The budget PDFs are native text (not scanned images), so no OCR is needed.
Tables are detected by clustering the right-edge (x1) positions of numeric
words on each page; right-aligned dollar amounts in the same column share
an x1 position, giving clean column detection even without grid lines.

Markdown files are designed for LLM consumption and citation:
  - Tabular pages render as markdown pipe-tables (unambiguous columns)
  - Text pages render as cleaned prose
  - Each page has an HTML anchor (<!-- page:N -->) for precise citation
  - Header includes direct source URL to the original document

Usage:
    python extract_budgets.py              # extract all PDFs
    python extract_budgets.py --fy 26      # extract only FY26
    python extract_budgets.py --force      # re-extract existing files
    python extract_budgets.py --dry-run    # list PDFs without extracting
"""

import argparse
import csv
import re
import sys
from datetime import datetime
from pathlib import Path

import pdfplumber

DATA_DIR = Path(__file__).parent
PDF_DIR = DATA_DIR / "pdf"
CSV_DIR = DATA_DIR / "csv"
MD_DIR = DATA_DIR / "md"

RESOURCE_VIEW_URL = "https://www.rsu5.org/fs/resource-manager/view/{}"

DOC_LABELS = {
    "citizens-adopted-budget": "Citizens' Adopted Budget",
    "board-adopted-budget": "Board of Directors Adopted Budget",
    "board-adopted-handbook": "Board of Directors Adopted Budget Handbook",
    "board-adopted-summary": "Budget Summary (Board Adopted)",
    "superintendent-handbook": "Superintendent's Recommended Budget Handbook",
    "budget-articles": "Budget Articles",
    "budget-overview": "Budget Overview",
    "budget-worksheet": "Budget Worksheet",
    "projected-budgets-fy27-29": "Projected FY27, FY28, FY29 Budgets",
    "year-to-year-staffing": "Year to Year Staffing Comparison",
}

BUDGET_PAGE_URLS = {
    22: "https://www.rsu5.org/budget/budget-2021-2022",
    23: "https://www.rsu5.org/budget/budget-2022-2023",
    24: "https://www.rsu5.org/budget/FY24",
    25: "https://www.rsu5.org/budget/FY25",
    26: "https://www.rsu5.org/budget/FY26",
    27: "https://www.rsu5.org/budget/FY27",
}

# Mapping from (fy, slug) -> resource manager UUID, imported from
# the fetch_budgets.py manifest so we can build direct source links.
DOCUMENT_UUIDS = {
    (22, "citizens-adopted-budget"): "e9ec20a7-2e8d-44a8-840f-28b79f07dbf8",
    (22, "superintendent-handbook"): "4e366c9c-8e73-4e47-bcb9-e8dc8de17042",
    (22, "budget-worksheet"): "0c7fb758-fbaf-488b-a99c-9d60cec1a07e",
    (22, "budget-overview"): "9001d364-14b2-421e-bef4-100c68b51468",
    (23, "citizens-adopted-budget"): "a83623b7-cabf-4b3b-8061-ab83fac4322f",
    (23, "board-adopted-budget"): "07cd8056-2424-48db-8056-8c643c73f8d2",
    (23, "board-adopted-summary"): "9b78848e-5a06-473b-b908-cb45cb9b0aaf",
    (23, "superintendent-handbook"): "7daca974-929e-4454-952a-b7385b7a71d6",
    (23, "budget-articles"): "dd48b04b-f4c5-4ed9-9e95-4e276b83ccb1",
    (24, "citizens-adopted-budget"): "9fedddf4-3173-4e93-8fd4-3057f6592662",
    (24, "board-adopted-budget"): "e83ed7f2-3c2d-4dca-80c4-778c57c29b99",
    (24, "board-adopted-handbook"): "36544929-6683-42fe-aa26-b6a0212ff34e",
    (24, "superintendent-handbook"): "539660a7-7280-453d-b5b4-1052ed29f4c9",
    (24, "budget-articles"): "653eeb67-20db-46e0-be3d-1e41dd535e15",
    (25, "citizens-adopted-budget"): "805aa39f-7658-436d-b81f-e49bf698d591",
    (25, "board-adopted-budget"): "56dfda0c-b059-4e64-a9b1-cb83a11ed539",
    (25, "board-adopted-handbook"): "55d08dc3-9994-4b26-b91e-b8a23bbecc4f",
    (25, "superintendent-handbook"): "5df6821b-6911-49c4-85e9-470863197ef2",
    (25, "budget-articles"): "e6bee6f8-1b48-4f60-9124-65e1773ccd6a",
    (26, "citizens-adopted-budget"): "10d5594e-5852-48a9-b360-ca0ea23e71e8",
    (26, "board-adopted-budget"): "cf517220-85f6-434e-9070-d30333798b25",
    (26, "board-adopted-handbook"): "05ebdd9a-c0b4-41cf-82c1-30be30922cf4",
    (26, "superintendent-handbook"): "85a85f5b-822a-4ba1-ad28-35929fce4237",
    (26, "budget-articles"): "3e6d92e0-ad8f-4d70-8d0a-d8c4a87d1ddb",
    (27, "superintendent-handbook"): "a3843da0-775f-4d2a-bee4-2f23c831c784",
    (27, "budget-articles"): "578b94ee-2505-47e4-bf76-1c9a0845515f",
    (27, "projected-budgets-fy27-29"): "c31d8676-131c-4729-9439-ecb2b7ef96d0",
    (27, "year-to-year-staffing"): "4414c97f-782f-4ae9-8bd6-0d6efa2e99d5",
}

_NUM_RE = re.compile(r"[\$,%\(\)\-\s]")

# Patterns for printer/report noise to strip from page text
_NOISE_PATTERNS = [
    re.compile(r"^Printed:\s+\d.+$", re.MULTILINE),
    re.compile(r"^rptGL\w+$", re.MULTILINE),
    re.compile(r"^Page:?\s*\d+\s*$", re.MULTILINE),
    re.compile(r"^\d{1,3}\s*$", re.MULTILINE),  # bare page numbers
    re.compile(
        r"^Print accounts with zero balance.*$", re.MULTILINE
    ),
    re.compile(r"^Exclude inactive accounts.*$", re.MULTILINE),
    re.compile(
        r"^(Include pre encumbrance|Filter Encumbrance).*$", re.MULTILINE
    ),
]


def _is_numeric(text: str) -> bool:
    t = text.strip()
    if not t:
        return False
    cleaned = _NUM_RE.sub("", t)
    if not cleaned:
        return t in ("$", "%", "-")
    digits = sum(c.isdigit() or c == "." for c in cleaned)
    return digits > 0 and digits >= len(cleaned) * 0.5


def _cluster(values, tol=10, min_count=3):
    if not values:
        return []
    vs = sorted(values)
    clusters, cur = [], [vs[0]]
    for v in vs[1:]:
        if v - cur[-1] <= tol:
            cur.append(v)
        else:
            if len(cur) >= min_count:
                clusters.append(sorted(cur)[len(cur) // 2])
            cur = [v]
    if len(cur) >= min_count:
        clusters.append(sorted(cur)[len(cur) // 2])
    return clusters


# -- table extraction via x1-clustering -----------------------------------


def _detect_columns(words, page_width):
    num_x1 = [round(w["x1"], 0) for w in words if _is_numeric(w["text"])]
    txt_x0 = [round(w["x0"], 0) for w in words if not _is_numeric(w["text"])]

    num_cols = _cluster(num_x1, tol=10, min_count=3)
    txt_cols = _cluster(txt_x0, tol=10, min_count=3)

    if len(num_cols) < 2:
        return None, None

    all_cols = sorted(set(num_cols + txt_cols))

    merged = [all_cols[0]]
    for c in all_cols[1:]:
        if c - merged[-1] < 20:
            merged[-1] = (merged[-1] + c) / 2
        else:
            merged.append(c)

    if len(merged) < 3:
        return None, None

    boundaries = [0.0]
    for i in range(len(merged) - 1):
        boundaries.append((merged[i] + merged[i + 1]) / 2)
    boundaries.append(page_width)

    return merged, boundaries


def _group_rows(words, y_tol=3):
    if not words:
        return []
    ws = sorted(words, key=lambda w: (w["top"], w["x0"]))
    rows, cur, cy = [], [ws[0]], ws[0]["top"]
    for w in ws[1:]:
        if abs(w["top"] - cy) <= y_tol:
            cur.append(w)
        else:
            rows.append(sorted(cur, key=lambda w: w["x0"]))
            cur, cy = [w], w["top"]
    if cur:
        rows.append(sorted(cur, key=lambda w: w["x0"]))
    return rows


def _word_to_col(w, boundaries):
    center = (w["x0"] + w["x1"]) / 2
    for i in range(len(boundaries) - 1):
        if boundaries[i] <= center < boundaries[i + 1]:
            return i
    return len(boundaries) - 2


def extract_table(page):
    words = page.extract_words(x_tolerance=5, y_tolerance=3,
                               keep_blank_chars=False)
    if len(words) < 10:
        return None

    num_frac = sum(_is_numeric(w["text"]) for w in words) / len(words)
    if num_frac < 0.10:
        return None

    col_positions, boundaries = _detect_columns(words, page.width)
    if col_positions is None:
        return None

    n_cols = len(col_positions)
    rows = _group_rows(words, y_tol=3)
    if len(rows) < 3:
        return None

    table = []
    for row_words in rows:
        cells = [""] * n_cols
        for w in row_words:
            ci = _word_to_col(w, boundaries)
            if 0 <= ci < n_cols:
                cells[ci] = (cells[ci] + " " + w["text"]).strip()
        if any(c.strip() for c in cells):
            table.append(cells)

    if len(table) < 3 or n_cols < 4:
        return None

    return _filter_noise_rows(table)


# -- noise stripping -------------------------------------------------------

_NOISE_ROW_RE = [
    re.compile(r"Printed:\s*\d", re.IGNORECASE),
    re.compile(r"^rptGL\w+$"),
    re.compile(r"Print accounts with zero balance", re.IGNORECASE),
    re.compile(r"Exclude inactive accounts", re.IGNORECASE),
    re.compile(r"Include pre.?encumbrance", re.IGNORECASE),
    re.compile(r"Filter Encumbrance", re.IGNORECASE),
    re.compile(r"Round to whole dollars", re.IGNORECASE),
    re.compile(r"Account on new page", re.IGNORECASE),
]


def _filter_noise_rows(table):
    """Remove printer footers, report metadata, and bare page-number rows."""
    cleaned = []
    for row in table:
        text = " ".join(c.strip() for c in row if c.strip())
        if not text:
            continue
        if any(pat.search(text) for pat in _NOISE_ROW_RE):
            continue
        # Bare page number in last cell (e.g., "Page 1" or just "47")
        filled = [c.strip() for c in row if c.strip()]
        if len(filled) == 1 and re.match(r"^(Page\s*)?\d{1,3}$", filled[0]):
            continue
        cleaned.append(row)
    return cleaned if len(cleaned) >= 3 else None


def _clean_text(text: str) -> str:
    """Remove printer footers, report metadata, and bare page numbers."""
    for pat in _NOISE_PATTERNS:
        text = pat.sub("", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()


# -- markdown pipe-table rendering -----------------------------------------


def _escape_pipe(s: str) -> str:
    return s.replace("|", "\\|")


def _render_md_table(table: list[list[str]]) -> str:
    """Render a list-of-lists as a markdown pipe-table.

    The first row with any non-numeric, non-empty content is treated as the
    header.  If no clear header exists, generic column names are used.
    """
    if not table or not table[0]:
        return ""

    n_cols = max(len(r) for r in table)

    # Pad short rows
    padded = [list(r) + [""] * (n_cols - len(r)) for r in table]

    # Detect header: first row where majority of filled cells are non-numeric
    header_idx = 0
    for i, row in enumerate(padded[:5]):
        filled = [c for c in row if c.strip()]
        if filled and sum(1 for c in filled if not _is_numeric(c)) > len(filled) * 0.4:
            header_idx = i
            break

    header = padded[header_idx]
    data_rows = padded[header_idx + 1:]

    # If no useful header was found, skip rows before the first data
    if all(not c.strip() for c in header):
        header = [f"Col {i+1}" for i in range(n_cols)]
        data_rows = padded

    lines = []
    lines.append("| " + " | ".join(_escape_pipe(h.strip()) for h in header) + " |")
    lines.append("|" + "|".join(" --- " for _ in header) + "|")

    for row in data_rows:
        cells = [_escape_pipe(c.strip()) for c in row[:n_cols]]
        if any(cells):
            lines.append("| " + " | ".join(cells) + " |")

    return "\n".join(lines)


# -- markdown document generation -----------------------------------------


def _parse_stem(stem):
    m = re.match(r"FY(\d+)-(.+)", stem)
    return (int(m.group(1)), m.group(2)) if m else (None, stem)


def build_markdown(pdf_path, pages_content, n_pages):
    """Build a citation-ready markdown document.

    pages_content: list of (page_1based, cleaned_text, table_or_None)
    """
    stem = pdf_path.stem
    fy, slug = _parse_stem(stem)
    label = DOC_LABELS.get(slug, slug.replace("-", " ").title())
    fy_str = f"FY{fy} ({2000 + fy - 1}-{2000 + fy})" if fy else stem
    budget_page = BUDGET_PAGE_URLS.get(fy, "")
    uuid = DOCUMENT_UUIDS.get((fy, slug), "") if fy else ""
    doc_url = RESOURCE_VIEW_URL.format(uuid) if uuid else ""

    lines = [f"# RSU 5 Budget -- {label}"]
    lines.append(f"**Fiscal Year:** {fy_str}  ")
    if doc_url:
        lines.append(f"**Document:** [{label}]({doc_url})  ")
    if budget_page:
        lines.append(f"**Budget Page:** [{fy_str}]({budget_page})  ")
    lines.append(f"**PDF:** `{pdf_path.name}` ({n_pages} pages)  ")
    lines.append(f"**Extracted:** {datetime.now().strftime('%Y-%m-%d %H:%M')}  ")
    lines.append("")
    lines.append("---")

    for page_num, text, table in pages_content:
        if not text.strip() and table is None:
            continue

        lines.append("")
        lines.append(f"<!-- page:{page_num} -->")
        lines.append(f"### Page {page_num}")
        lines.append("")

        if table is not None:
            # For table pages, render any preamble text that precedes the
            # tabular data (like article headers, section titles), then
            # the table itself as a pipe-table.
            preamble = _extract_preamble(text, table)
            if preamble:
                lines.append(preamble)
                lines.append("")
            lines.append(_render_md_table(table))
        else:
            lines.append(text)

    lines.append("")
    return "\n".join(lines) + "\n"


def _extract_preamble(full_text: str, table: list[list[str]]) -> str:
    """Extract non-tabular text that precedes a table on the same page.

    We find lines in the full text that don't appear to be part of the
    extracted table data (e.g., article headers, fiscal year labels).
    """
    if not full_text or not table:
        return ""

    # Collect all non-empty cell values from the first few table rows
    # as a quick lookup to identify which text lines are "in" the table
    table_tokens = set()
    for row in table[:5]:
        for cell in row:
            for token in cell.strip().split():
                if len(token) > 2:
                    table_tokens.add(token)

    preamble_lines = []
    text_lines = full_text.strip().split("\n")

    for line in text_lines:
        stripped = line.strip()
        if not stripped:
            continue
        # Check how many tokens in this line match table content
        words = stripped.split()
        if not words:
            continue
        match_count = sum(1 for w in words if w in table_tokens)
        match_frac = match_count / len(words) if words else 0

        # If < 40% of words match table content, it's likely preamble
        if match_frac < 0.4 and len(stripped) > 3:
            preamble_lines.append(stripped)
        else:
            break  # once we hit table-like content, stop

    return "\n".join(preamble_lines)


# -- CSV output ------------------------------------------------------------


def write_csv(table, path):
    path.parent.mkdir(parents=True, exist_ok=True)
    with open(path, "w", newline="", encoding="utf-8") as f:
        csv.writer(f).writerows(table)


def _is_repeated_header(row, header_row):
    if not header_row:
        return False
    matches = sum(1 for a, b in zip(row, header_row)
                  if a.strip() == b.strip() and a.strip())
    return matches >= max(2, len(header_row) * 0.4)


# -- main ------------------------------------------------------------------


def main():
    parser = argparse.ArgumentParser(
        description="Extract text and tables from RSU 5 budget PDFs")
    parser.add_argument("--dry-run", action="store_true")
    parser.add_argument("--fy", type=int, default=None)
    parser.add_argument("--force", action="store_true")
    parser.add_argument("--single", type=str, default=None)
    args = parser.parse_args()

    MD_DIR.mkdir(parents=True, exist_ok=True)
    CSV_DIR.mkdir(parents=True, exist_ok=True)

    pdfs = sorted(PDF_DIR.glob("*.pdf"))
    if not pdfs:
        print("No PDFs found in pdf/. Run fetch_budgets.py first.",
              file=sys.stderr)
        sys.exit(1)

    if args.single:
        pdfs = [p for p in pdfs if args.single in (p.name, p.stem)]
    elif args.fy is not None:
        pdfs = [p for p in pdfs if p.name.startswith(f"FY{args.fy}-")]

    if not pdfs:
        print("No matching PDFs found.", file=sys.stderr)
        sys.exit(1)

    if args.dry_run:
        for p in pdfs:
            pdf = pdfplumber.open(str(p))
            print(f"  {p.name}: {len(pdf.pages)} pages, "
                  f"{p.stat().st_size // 1024} KB")
            pdf.close()
        return

    total_md = 0
    total_csv = 0

    for idx, pdf_path in enumerate(pdfs):
        tag = f"[{idx + 1}/{len(pdfs)}]"
        stem = pdf_path.stem
        md_path = MD_DIR / f"{stem}.md"

        if md_path.exists() and not args.force:
            print(f"  {tag} {pdf_path.name} -- already extracted, skipping")
            continue

        try:
            pdf = pdfplumber.open(str(pdf_path))
            n_pages = len(pdf.pages)
            print(f"  {tag} {pdf_path.name} ({n_pages} pages)")

            pages_content = []   # (page_num, cleaned_text, table_or_None)
            page_tables = []     # (page_num, table) for CSV output

            for i, page in enumerate(pdf.pages):
                raw_text = page.extract_text() or ""
                cleaned = _clean_text(raw_text)
                tbl = extract_table(page)

                pages_content.append((i + 1, cleaned, tbl))
                if tbl:
                    page_tables.append((i + 1, tbl))

            # -- write markdown --
            md_content = build_markdown(pdf_path, pages_content, n_pages)
            md_path.write_text(md_content, encoding="utf-8")
            total_md += 1
            tbl_pages = sum(1 for _, _, t in pages_content if t is not None)
            txt_pages = sum(1 for _, t, tbl in pages_content
                           if t.strip() and tbl is None)
            print(f"       -> md/{stem}.md "
                  f"({tbl_pages} table pages, {txt_pages} text pages)")

            # -- write CSVs --
            if page_tables:
                groups = [[page_tables[0]]]
                for pt in page_tables[1:]:
                    prev_page, prev_tbl = groups[-1][-1]
                    cur_page, cur_tbl = pt
                    same_cols = (len(cur_tbl[0]) == len(prev_tbl[0]))
                    contiguous = (cur_page == prev_page + 1)
                    if same_cols and contiguous:
                        groups[-1].append(pt)
                    else:
                        groups.append([pt])

                for group in groups:
                    combined = []
                    header_row = None
                    for page_num, tbl in group:
                        if not combined:
                            combined.extend(tbl)
                            header_row = tbl[0] if tbl else None
                        else:
                            start = 0
                            if tbl and _is_repeated_header(tbl[0], header_row):
                                start = 1
                            combined.extend(tbl[start:])

                    csv_name = (f"{stem}.csv" if len(groups) == 1
                                else f"{stem}-p{group[0][0]:02d}.csv")
                    csv_path = CSV_DIR / csv_name
                    write_csv(combined, csv_path)
                    total_csv += 1

                    n_rows = len(combined)
                    n_cols = len(combined[0]) if combined else 0
                    span = (f"p{group[0][0]}-{group[-1][0]}"
                            if len(group) > 1 else f"p{group[0][0]}")
                    print(f"       -> csv/{csv_name} "
                          f"({n_rows} rows x {n_cols} cols, {span})")

            pdf.close()

        except Exception as e:
            print(f"  {tag} ERROR: {e}", file=sys.stderr)
            import traceback
            traceback.print_exc()

    print(f"\nDone. {total_md} markdown files, {total_csv} CSV files.")
    print(f"Output: {MD_DIR}, {CSV_DIR}")


if __name__ == "__main__":
    main()
