"""Parse extracted budget CSVs into BudgetLineItem and SummaryRow objects.

The extracted CSVs are messy: descriptions are split across cells, dollar
amounts use ``$X,XXX.XX`` formatting, page headers repeat mid-file, and
program/function/cost-center summary rows are interleaved with line items.

Two main formats exist:
  * **Board-adopted / Citizens-adopted / Budget articles** -- dot-separated
    account codes, ``$``-formatted amounts.
  * **Budget worksheet** -- hyphen-separated codes, line-numbered rows,
    plain numeric amounts.

This module handles both transparently.
"""

from __future__ import annotations

import csv
import re
from pathlib import Path

from rsu5.model import BudgetLineItem, SummaryRow

# ── Account code patterns ──────────────────────────────────────
_DOT_RE = re.compile(r"(\d{4})\.(\d{4})\.(\d{4})\.(\d{5})\.(\d{3})")
_HYP_RE = re.compile(r"(\d{4})-(\d{4})-(\d{4})-(\d{5})-(\d{3})")

# ── Summary-row patterns ──────────────────────────────────────
_PROGRAM_RE = re.compile(
    r"Program:?\s*(.+)", re.IGNORECASE
)
_FUNCTION_RE = re.compile(
    r"Function:?\s*(.+)", re.IGNORECASE
)
_COST_CENTER_RE = re.compile(
    r"Cost\s+Center:?\s*(.+)", re.IGNORECASE
)

# ── Article header detection ──────────────────────────────────
_ARTICLE_HEADER_RE = re.compile(
    r"(?:ARTICLE\s+)?(\d{1,2})\s*[-–]\s*(.+)", re.IGNORECASE
)
_ARTICLE_LINE_RE = re.compile(
    r"(?:FY\d{2}\s+)?Article\s+(\d{1,2})\b", re.IGNORECASE
)

# ── Noise patterns (rows to skip entirely) ─────────────────────
_NOISE_PATTERNS = [
    re.compile(r"RSU\s+No\.?\s*5", re.IGNORECASE),
    re.compile(r"From\s+Date:", re.IGNORECASE),
    re.compile(r"Page\s+\d+\s+of\s+\d+", re.IGNORECASE),
    re.compile(r"\d+\s+of\s+\d+\s*$"),
    re.compile(r"Printed:", re.IGNORECASE),
    re.compile(r"Definition:\s*FY", re.IGNORECASE),
    re.compile(r"^,*\s*$"),  # blank rows
    re.compile(r"^,*Account,", re.IGNORECASE),  # column header
    re.compile(r"^,*LINE,", re.IGNORECASE),  # worksheet header
]

# ── Dollar parsing ─────────────────────────────────────────────

def _parse_dollar(raw: str) -> float | None:
    """Parse a dollar string into a float, handling ``$``, commas, parens.

    Returns None if the string is not a recognizable dollar amount.
    """
    if not raw:
        return None
    s = raw.strip()
    if not s or s == "-":
        return None

    negative = False
    if s.startswith("(") and s.endswith(")"):
        negative = True
        s = s[1:-1]
    elif s.startswith("- ") or s.startswith("-$"):
        negative = True
        s = s.lstrip("- ")

    s = s.replace("$", "").replace(",", "").strip()
    if not s:
        return None

    # Handle parenthetical negatives that may appear after stripping
    if s.startswith("(") and s.endswith(")"):
        negative = True
        s = s[1:-1]

    try:
        val = float(s)
    except ValueError:
        return None

    if negative:
        val = -val
    return val


def _is_noise(cells: list[str]) -> bool:
    """True if this row is page-header noise that should be skipped."""
    joined = ",".join(cells)
    for pat in _NOISE_PATTERNS:
        if pat.search(joined):
            return True
    # Pure column-header rows (contain "Actual", "Adopted", etc. but no numbers)
    if any(
        h in joined
        for h in ["FY1", "FY2", "FY3", "Total Budget", "Dollar", "Percent"]
    ) and not _DOT_RE.search(joined) and not _HYP_RE.search(joined):
        return True
    return False


def _find_account_code(cells: list[str]) -> tuple[re.Match | None, int]:
    """Search cells for an account code, returning (match, cell_index)."""
    for i, cell in enumerate(cells):
        m = _DOT_RE.search(cell)
        if m:
            return m, i
        m = _HYP_RE.search(cell)
        if m:
            return m, i
    return None, -1


def _extract_description(cells: list[str], acct_idx: int) -> str:
    """Reconstruct the description from cells following the account code.

    The RSU PDFs split descriptions across 2-4 CSV cells.
    """
    desc_parts = []
    for i in range(acct_idx + 1, len(cells)):
        cell = cells[i].strip()
        if not cell:
            continue
        # Stop when we hit a dollar amount
        if _parse_dollar(cell) is not None:
            break
        # Stop at noise fragments
        if cell in ("-", "%"):
            break
        desc_parts.append(cell)
    return " ".join(desc_parts).strip().rstrip(",").strip()


def _extract_amounts(cells: list[str], acct_idx: int) -> list[float]:
    """Extract all dollar amounts from cells after the account code.

    Empty cells between values are PDF extraction artifacts, not missing
    data -- we skip them and collect only actual numbers.
    """
    amounts: list[float] = []
    for i in range(acct_idx + 1, len(cells)):
        cell = cells[i].strip()
        if not cell or cell == "-":
            continue
        # Skip description fragments (non-numeric text)
        pct = cell.rstrip("%")
        val = _parse_dollar(pct)
        if val is not None:
            amounts.append(val)
    return amounts


_SUMMARY_KEYWORDS = [
    ("program", re.compile(r"Program:?", re.IGNORECASE)),
    ("function", re.compile(r"Function:?", re.IGNORECASE)),
    ("cost_center", re.compile(r"Cost\s*Center:?", re.IGNORECASE)),
]


def _detect_summary(cells: list[str]) -> tuple[str, str, list[float]] | None:
    """Detect a summary row (Program:, Function:, Cost Center:).

    Returns (level, label, amounts) or None.
    """
    # Find the keyword cell
    level = ""
    keyword_idx = -1
    for i, cell in enumerate(cells):
        for lv, pat in _SUMMARY_KEYWORDS:
            if pat.search(cell):
                level = lv
                keyword_idx = i
                break
        if keyword_idx >= 0:
            break

    if keyword_idx < 0:
        # Also check multi-cell "Cost" + "Center:" pattern
        for i, cell in enumerate(cells):
            if cell.strip().upper() == "COST" and i + 1 < len(cells):
                if re.match(r"Center:?", cells[i + 1].strip(), re.IGNORECASE):
                    level = "cost_center"
                    keyword_idx = i + 1
                    break

    if keyword_idx < 0:
        return None

    # Gather label fragments (non-dollar cells after keyword, before amounts)
    label_parts: list[str] = []
    amounts_start = len(cells)
    for i in range(keyword_idx + 1, len(cells)):
        cell = cells[i].strip()
        if not cell:
            continue
        # Once we see a dollar-formatted value (with $ or a number > 1000
        # that isn't a code), we've reached the amounts region.
        if "$" in cell:
            amounts_start = i
            break
        # Check if it looks like a code fragment (e.g. "- 1100", "010")
        # These are label parts, not amounts.
        code_like = re.match(r"^-?\s*\d{3,5}\s*$", cell)
        if code_like:
            label_parts.append(cell)
            continue
        # Check for a standalone dollar amount without $
        val = _parse_dollar(cell)
        if val is not None and abs(val) > 100:
            amounts_start = i
            break
        label_parts.append(cell)

    label = " ".join(label_parts).strip()

    # Extract amounts from remaining cells
    amounts: list[float] = []
    for cell in cells[amounts_start:]:
        val = _parse_dollar(cell)
        if val is not None:
            amounts.append(val)

    if not amounts:
        return None

    return level, label, amounts


def _extract_summary_code(label: str) -> str:
    """Pull the numeric code from a summary label like ``ELEMENTARY PROGRAMS - 1100``."""
    m = re.search(r"-\s*(\d{3,5})\s*$", label)
    if m:
        return m.group(1)
    m = re.search(r"^(\d{3,5})\s*-", label)
    if m:
        return m.group(1)
    return label.strip()


def _detect_article_from_row(cells: list[str], current_article: int) -> int:
    """Check if a row is an article header and update the current article number."""
    joined = " ".join(c.strip() for c in cells if c.strip())

    m = _ARTICLE_LINE_RE.search(joined)
    if m:
        return int(m.group(1))

    if "ARTICLE" in joined.upper():
        m2 = _ARTICLE_HEADER_RE.search(joined)
        if m2:
            return int(m2.group(1))

    if "EXPENDITURES" in joined.upper() or "REVENUES" in joined.upper():
        return current_article

    return current_article


def parse_csv_file(
    csv_path: Path,
    fy: int,
    doc_type: str,
    column_names: list[str],
    initial_article: int = 0,
) -> tuple[list[BudgetLineItem], list[SummaryRow]]:
    """Parse a single extracted budget CSV file.

    Args:
        csv_path: Path to the CSV file.
        fy: Fiscal year (2-digit, e.g. 27).
        doc_type: Document type (``"adopted"``, ``"proposed"``, etc.).
        column_names: Ordered list of column semantics from config.
        initial_article: Starting article number (for multi-page CSVs
            where the article was established on a prior page).

    Returns:
        Tuple of (line_items, summary_rows).
    """
    line_items: list[BudgetLineItem] = []
    summary_rows: list[SummaryRow] = []
    current_article = initial_article

    # Infer source page from filename (e.g. "FY27-budget-articles-p03.csv")
    page_match = re.search(r"-p(\d+)\.csv$", csv_path.name)
    source_page = int(page_match.group(1)) if page_match else 0

    with open(csv_path, encoding="utf-8", errors="replace") as f:
        reader = csv.reader(f)
        for cells in reader:
            if not cells or _is_noise(cells):
                # But first check for article headers in noise
                current_article = _detect_article_from_row(
                    cells, current_article
                )
                continue

            # Check for article header
            current_article = _detect_article_from_row(cells, current_article)

            # Check for summary row
            summary = _detect_summary(cells)
            if summary:
                level, label, amounts = summary
                code = _extract_summary_code(label)

                amt_dict: dict[str, float] = {}
                for i, val in enumerate(amounts):
                    if val is not None and i < len(column_names):
                        amt_dict[column_names[i]] = val

                summary_rows.append(
                    SummaryRow(
                        fy=fy,
                        doc_type=doc_type,
                        level=level,
                        code=code,
                        label=label,
                        amounts=amt_dict,
                        article=current_article,
                        source_file=csv_path.name,
                    )
                )
                continue

            # Check for account code (line item)
            acct_match, acct_idx = _find_account_code(cells)
            if acct_match is None:
                continue

            fund, program, function, obj, cc = acct_match.groups()
            desc = _extract_description(cells, acct_idx)
            raw_amounts = _extract_amounts(cells, acct_idx)

            amt_dict = {}
            for i, val in enumerate(raw_amounts):
                if val is not None and i < len(column_names):
                    amt_dict[column_names[i]] = val

            line_items.append(
                BudgetLineItem(
                    fy=fy,
                    doc_type=doc_type,
                    article=current_article,
                    fund=fund,
                    program=program,
                    function=function,
                    object_code=obj,
                    cost_center=cc,
                    description=desc,
                    amounts=amt_dict,
                    source_file=csv_path.name,
                    source_page=source_page,
                )
            )

    return line_items, summary_rows


def parse_document_csvs(
    csv_dir: Path,
    prefix: str,
    fy: int,
    doc_type: str,
    column_names: list[str],
) -> tuple[list[BudgetLineItem], list[SummaryRow]]:
    """Parse all CSV pages for a single budget document.

    CSVs are named ``{prefix}-p{NN}.csv`` and are processed in page order
    so article tracking carries across pages.

    Returns combined (line_items, summary_rows).
    """
    pattern = f"{prefix}-p*.csv"
    csv_files = sorted(csv_dir.glob(pattern))

    # Also check for single-page documents (no page suffix)
    single = csv_dir / f"{prefix}.csv"
    if single.exists() and single not in csv_files:
        csv_files.insert(0, single)

    if not csv_files:
        return [], []

    all_items: list[BudgetLineItem] = []
    all_summaries: list[SummaryRow] = []
    current_article = 0

    for csv_path in csv_files:
        items, summaries = parse_csv_file(
            csv_path, fy, doc_type, column_names,
            initial_article=current_article,
        )
        all_items.extend(items)
        all_summaries.extend(summaries)

        # Carry article context to next page
        if items:
            current_article = items[-1].article
        elif summaries:
            current_article = summaries[-1].article

    return all_items, all_summaries
