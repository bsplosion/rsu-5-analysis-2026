"""Parse superintendent/board-adopted handbook CSVs into structured data.

Handbook CSVs vary significantly across fiscal years (FY22-FY27) in layout,
column count, and formatting. This module uses content-based detection to
identify what type of data each CSV page contains, then applies the
appropriate extraction logic.

Data types extracted:
  - Budget history:  10-year adopted budget totals with year-over-year changes
  - Article totals:  Per-article budget breakdown (adopted vs proposed)
  - Enrollment:      School-by-school enrollment with projections
  - Reductions:      Staffing/spending reductions by tier (T1/T2/etc.)
"""

from __future__ import annotations

import csv
import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any


# ---------------------------------------------------------------------------
# Data structures
# ---------------------------------------------------------------------------

@dataclass
class BudgetHistoryEntry:
    fy: int
    adopted: float
    difference: float
    pct_increase: float
    source_file: str = ""


@dataclass
class ArticleTotal:
    article: int
    name: str
    adopted: float          # prior year adopted
    proposed: float         # current year proposed/adopted
    difference: float
    pct_change: float
    source_file: str = ""


@dataclass
class EnrollmentEntry:
    school: str
    years: dict[int, int]   # calendar year -> count
    source_file: str = ""


@dataclass
class ReductionItem:
    tier: str               # M, N, R, T1, T2, R/T1, R/T1/T2
    description: str
    location: str           # school/dept abbreviation
    initial_request: float
    reduction_amount: float
    proposed_amount: float
    source_file: str = ""


@dataclass
class HandbookData:
    """All structured data extracted from handbook CSVs for a single FY."""
    fy: int
    doc_type: str           # "superintendent" or "board-adopted"
    budget_history: list[BudgetHistoryEntry] = field(default_factory=list)
    article_totals: list[ArticleTotal] = field(default_factory=list)
    enrollment: list[EnrollmentEntry] = field(default_factory=list)
    reductions: list[ReductionItem] = field(default_factory=list)
    grand_total_adopted: float | None = None
    grand_total_proposed: float | None = None
    source_files: list[str] = field(default_factory=list)


# ---------------------------------------------------------------------------
# Dollar / percentage parsing
# ---------------------------------------------------------------------------

_DOLLAR_RE = re.compile(r"[\$,\s]")

def _parse_dollar(raw: str) -> float | None:
    """Parse a dollar amount, handling parens for negatives."""
    s = raw.strip()
    if not s or s == "-":
        return None
    s = s.replace("\u2013", "-").replace("\u2014", "-")
    neg = False
    if s.startswith("(") and s.endswith(")"):
        neg = True
        s = s[1:-1]
    elif s.startswith("$(") and s.endswith(")"):
        neg = True
        s = s[2:-1]
    s = _DOLLAR_RE.sub("", s)
    if not s:
        return None
    # Fix OCR space-in-number errors like "7 20,164"
    s = re.sub(r"(\d) (\d)", r"\1\2", s)
    # Fix period-as-comma typo like "$100.000" that should be "$100,000"
    if re.match(r"^\d+\.\d{3}$", s):
        s = s.replace(".", "")
    try:
        val = float(s)
        return -val if neg else val
    except ValueError:
        return None


def _parse_pct(raw: str) -> float | None:
    """Parse a percentage like '6.83%' or '(2.13%)'."""
    s = raw.strip()
    if not s or s == "-":
        return None
    neg = False
    if s.startswith("(") and s.endswith(")"):
        neg = True
        s = s[1:-1]
    s = s.rstrip("%").strip()
    s = s.replace(",", "")
    try:
        val = float(s)
        return -val if neg else val
    except ValueError:
        return None


# ---------------------------------------------------------------------------
# Content-type detection
# ---------------------------------------------------------------------------

def _join_cells(cells: list[str]) -> str:
    """Join cells into a single string for pattern matching."""
    return " ".join(c.strip() for c in cells if c.strip()).lower()


def _detect_page_type(rows: list[list[str]]) -> str | None:
    """Detect what data a handbook CSV page contains."""
    if len(rows) < 3:
        return None

    full_text = " ".join(_join_cells(r) for r in rows[:10])

    if "adopted budget" in full_text and "history" in full_text:
        return "budget_history"
    if "10 year" in full_text and "adopted" in full_text:
        return "budget_history"
    if "adopted budget" in full_text and "expenditure" in full_text and "difference" in full_text:
        return "budget_history"

    if "article" in full_text and ("proposed" in full_text or "adopted" in full_text):
        if any("article" in _join_cells(r) and any(
            c.strip().startswith("$") for c in r
        ) for r in rows[3:10]):
            return "article_totals"
        header_text = full_text
        if "description" in header_text and ("difference" in header_text or "budget" in header_text):
            return "article_totals"

    if "enrollment" in full_text or "projected" in full_text:
        if "october" in full_text or "oct" in full_text:
            return "enrollment"

    if "reductions" in full_text or "expenditure/reductions" in full_text:
        return "reductions"
    if "rating" in full_text and ("description" in full_text) and ("initial" in full_text or "request" in full_text):
        return "reductions"
    # Check for tier markers
    tier_count = sum(1 for r in rows if r and r[0].strip() in ("T1", "T2", "M", "N", "R"))
    if tier_count >= 3:
        return "reductions"

    if "reserve" in full_text and "account" in full_text:
        return "reserve_accounts"

    return None


# ---------------------------------------------------------------------------
# Budget history parser
# ---------------------------------------------------------------------------

def _parse_budget_history(rows: list[list[str]], source_file: str) -> list[BudgetHistoryEntry]:
    """Parse 10-year adopted budget history table."""
    entries = []
    for row in rows:
        joined = _join_cells(row)
        if not joined:
            continue

        fy_match = re.search(r"\bfy\s*(\d{2})\b", joined, re.IGNORECASE)
        if not fy_match:
            continue

        fy = int(fy_match.group(1))
        dollars = []
        pct = None
        for cell in row:
            s = cell.strip()
            if not s:
                continue
            if s.endswith("%") or (s.startswith("(") and s.endswith("%)")):
                pct = _parse_pct(s)
            elif "$" in s or (re.search(r"\d{2,3},\d{3}", s)):
                val = _parse_dollar(s)
                if val is not None:
                    dollars.append(val)

        if len(dollars) >= 2:
            entries.append(BudgetHistoryEntry(
                fy=fy,
                adopted=dollars[0],
                difference=dollars[1],
                pct_increase=pct if pct is not None else 0.0,
                source_file=source_file,
            ))
        elif len(dollars) == 1:
            entries.append(BudgetHistoryEntry(
                fy=fy,
                adopted=dollars[0],
                difference=0.0,
                pct_increase=pct if pct is not None else 0.0,
                source_file=source_file,
            ))

    return entries


# ---------------------------------------------------------------------------
# Article totals parser
# ---------------------------------------------------------------------------

def _parse_article_totals(
    rows: list[list[str]], source_file: str
) -> tuple[list[ArticleTotal], float | None, float | None]:
    """Parse article-by-article budget table.

    Returns (article_totals, grand_total_adopted, grand_total_proposed).
    """
    totals: list[ArticleTotal] = []
    grand_adopted = None
    grand_proposed = None

    for row in rows:
        joined = _join_cells(row)
        if not joined:
            continue

        art_match = re.search(r"article\s+(\d{1,2})", joined, re.IGNORECASE)
        is_total_row = "total" in joined.lower() and "article" not in joined.lower()
        is_operating = "operating" in joined.lower() and "total" in joined.lower()

        if not art_match and not is_total_row and not is_operating:
            if "total" in joined.lower() and "article" in joined.lower():
                is_total_row = True
            else:
                continue

        dollars = []
        pct = None
        for cell in row:
            s = cell.strip()
            if not s:
                continue
            if "%" in s:
                pct = _parse_pct(s)
            elif "$" in s or re.search(r"\d{2,3},\d{3}", s):
                val = _parse_dollar(s)
                if val is not None:
                    dollars.append(val)

        if art_match and len(dollars) >= 2:
            art_num = int(art_match.group(1))
            name_parts = []
            for cell in row:
                s = cell.strip()
                if s and not s.startswith("$") and not s.endswith("%") and \
                   not re.match(r"^[\d,\.\$\(\)\-%]+$", s) and \
                   not re.match(r"^article\s*$", s, re.IGNORECASE):
                    cleaned = re.sub(r"^\d+\s*", "", s).strip()
                    if cleaned and cleaned.lower() not in ("article", "#"):
                        name_parts.append(cleaned)
            name = " ".join(name_parts).strip()
            name = re.sub(r"^(Special|Career|Other|Student|System|School|Transportation|Facilities|Debt|All)\s", r"\1 ", name)

            adopted = dollars[0]
            proposed = dollars[1] if len(dollars) >= 2 else dollars[0]
            diff = dollars[2] if len(dollars) >= 3 else proposed - adopted

            totals.append(ArticleTotal(
                article=art_num,
                name=name,
                adopted=adopted,
                proposed=proposed,
                difference=diff,
                pct_change=pct if pct is not None else 0.0,
                source_file=source_file,
            ))

        elif (is_total_row or is_operating) and len(dollars) >= 2:
            if is_operating or "operating" in joined.lower():
                grand_adopted = dollars[0]
                grand_proposed = dollars[1]
            elif grand_adopted is None:
                grand_adopted = dollars[0]
                grand_proposed = dollars[1]

    return totals, grand_adopted, grand_proposed


# ---------------------------------------------------------------------------
# Enrollment parser
# ---------------------------------------------------------------------------

_SCHOOL_NAMES = {
    "morse": "Morse Street",
    "morsestreet": "Morse Street",
    "morse street": "Morse Street",
    "mast": "Mast Landing",
    "mastlanding": "Mast Landing",
    "mast landing": "Mast Landing",
    "pownal": "Pownal Elementary",
    "pownalelemntary": "Pownal Elementary",
    "pownal elementary": "Pownal Elementary",
    "pownalelem": "Pownal Elementary",
    "durham": "Durham Community",
    "durhamcommunity": "Durham Community",
    "durham community": "Durham Community",
    "freeport middle": "Freeport Middle",
    "freeportmiddle": "Freeport Middle",
    "fms": "Freeport Middle",
    "freeport high": "Freeport High",
    "freeporthigh": "Freeport High",
    "fhs": "Freeport High",
}


def _normalize_school(text: str) -> str | None:
    """Match school name from potentially mangled text."""
    t = text.lower().strip()
    t = re.sub(r"\s+", " ", t)
    for key, name in _SCHOOL_NAMES.items():
        if key in t:
            return name
    return None


def _parse_enrollment(rows: list[list[str]], source_file: str) -> list[EnrollmentEntry]:
    """Parse enrollment table with multi-year columns."""
    entries: list[EnrollmentEntry] = []

    years: list[int] = []
    for row in rows[:6]:
        for cell in row:
            for m in re.finditer(r"\b(20[12]\d)\b", cell):
                y = int(m.group(1))
                if y not in years:
                    years.append(y)
    years.sort()

    if not years:
        return entries

    for row in rows:
        joined = " ".join(c.strip() for c in row if c.strip())
        school = _normalize_school(joined)
        if not school:
            continue
        if "grand" in joined.lower() and "total" in joined.lower():
            school = "Grand Total"

        numbers = []
        for cell in row:
            s = cell.strip().replace(",", "")
            if re.match(r"^\d{2,4}$", s):
                val = int(s)
                if 10 <= val <= 3000:
                    numbers.append(val)

        if not numbers:
            continue

        year_map: dict[int, int] = {}
        for i, y in enumerate(years):
            if i < len(numbers):
                year_map[y] = numbers[i]

        if year_map:
            existing = next((e for e in entries if e.school == school), None)
            if existing:
                existing.years.update(year_map)
            else:
                entries.append(EnrollmentEntry(
                    school=school,
                    years=year_map,
                    source_file=source_file,
                ))

    return entries


# ---------------------------------------------------------------------------
# Reductions parser
# ---------------------------------------------------------------------------

_TIER_RE = re.compile(r"^(M|N|R|T1|T2|R/T1|R/T1/T2|OR)\b", re.IGNORECASE)
_LOCATION_RE = re.compile(
    r"\b(DCS|MSS|MLS|PES|FMS|FHS|DW|CTE|FY\d+)\b", re.IGNORECASE
)


def _parse_reductions(rows: list[list[str]], source_file: str) -> list[ReductionItem]:
    """Parse expenditure/reductions summary."""
    items: list[ReductionItem] = []
    current_tier = ""
    pending_desc_parts: list[str] = []

    for row in rows:
        if not any(c.strip() for c in row):
            continue
        joined = _join_cells(row)
        if "rating" in joined and "description" in joined:
            continue
        if "total" in joined and ("year over" in joined or "increase" in joined):
            continue
        if joined.startswith("m=") or joined.startswith("n=") or \
           joined.startswith("r=") or joined.startswith("t1=") or \
           joined.startswith("t2=") or joined.startswith("updated"):
            continue

        first = row[0].strip() if row else ""
        tier_match = _TIER_RE.match(first)

        dollars = []
        for cell in row:
            val = _parse_dollar(cell.strip())
            if val is not None:
                dollars.append(val)

        desc_parts = []
        for cell in row:
            s = cell.strip()
            if not s:
                continue
            if _TIER_RE.match(s) and s == first:
                continue
            if _parse_dollar(s) is not None:
                continue
            if s.endswith("%"):
                continue
            desc_parts.append(s)
        desc_text = " ".join(desc_parts)

        if tier_match:
            current_tier = tier_match.group(1).upper()

            if dollars:
                loc_match = _LOCATION_RE.search(desc_text)
                location = loc_match.group(1).upper() if loc_match else ""
                init_req = dollars[0] if len(dollars) >= 1 else 0.0
                reduction = dollars[1] if len(dollars) >= 2 else 0.0
                proposed = dollars[-1] if len(dollars) >= 2 else init_req

                items.append(ReductionItem(
                    tier=current_tier,
                    description=desc_text.strip(),
                    location=location,
                    initial_request=init_req,
                    reduction_amount=reduction,
                    proposed_amount=proposed,
                    source_file=source_file,
                ))
                pending_desc_parts = []
            else:
                pending_desc_parts = [desc_text]

        elif current_tier and dollars and pending_desc_parts:
            full_desc = " ".join(pending_desc_parts + [desc_text]).strip()
            loc_match = _LOCATION_RE.search(full_desc)
            location = loc_match.group(1).upper() if loc_match else ""
            init_req = dollars[0] if len(dollars) >= 1 else 0.0
            reduction = dollars[1] if len(dollars) >= 2 else 0.0
            proposed = dollars[-1] if len(dollars) >= 2 else init_req

            items.append(ReductionItem(
                tier=current_tier,
                description=full_desc,
                location=location,
                initial_request=init_req,
                reduction_amount=reduction,
                proposed_amount=proposed,
                source_file=source_file,
            ))
            pending_desc_parts = []

        elif not dollars and desc_text:
            pending_desc_parts.append(desc_text)

    return items


# ---------------------------------------------------------------------------
# Top-level parsing
# ---------------------------------------------------------------------------

def _read_csv(path: Path) -> list[list[str]]:
    """Read a CSV file, returning rows as lists of strings."""
    rows = []
    with open(path, newline="", encoding="utf-8") as f:
        for row in csv.reader(f):
            rows.append(row)
    return rows


def parse_handbook_csv(csv_path: Path, fy: int, doc_type: str) -> dict[str, Any]:
    """Parse a single handbook CSV page, returning extracted data by type.

    Returns a dict with keys like 'budget_history', 'article_totals', etc.
    Only populated keys are present.
    """
    rows = _read_csv(csv_path)
    page_type = _detect_page_type(rows)
    result: dict[str, Any] = {"page_type": page_type, "source": str(csv_path)}

    if page_type == "budget_history":
        result["budget_history"] = _parse_budget_history(rows, str(csv_path.name))

    elif page_type == "article_totals":
        arts, grand_a, grand_p = _parse_article_totals(rows, str(csv_path.name))
        result["article_totals"] = arts
        result["grand_total_adopted"] = grand_a
        result["grand_total_proposed"] = grand_p

    elif page_type == "enrollment":
        result["enrollment"] = _parse_enrollment(rows, str(csv_path.name))

    elif page_type == "reductions":
        result["reductions"] = _parse_reductions(rows, str(csv_path.name))

    return result


def parse_handbook_for_fy(
    csv_dir: Path, fy: int, doc_type: str = "superintendent"
) -> HandbookData:
    """Parse all handbook CSV pages for a given FY and doc type.

    Args:
        csv_dir: Directory containing extracted CSV files.
        fy: Fiscal year (e.g., 27).
        doc_type: 'superintendent' or 'board-adopted'.

    Returns:
        HandbookData with all extractable data merged.
    """
    slug = f"FY{fy}-{doc_type}-handbook" if doc_type == "board-adopted" else \
           f"FY{fy}-superintendent-handbook"
    csvs = sorted(csv_dir.glob(f"{slug}*.csv"))

    data = HandbookData(fy=fy, doc_type=doc_type)

    for csv_path in csvs:
        result = parse_handbook_csv(csv_path, fy, doc_type)
        data.source_files.append(csv_path.name)

        if "budget_history" in result and result["budget_history"]:
            existing_fys = {e.fy for e in data.budget_history}
            for entry in result["budget_history"]:
                if entry.fy not in existing_fys:
                    data.budget_history.append(entry)
                    existing_fys.add(entry.fy)

        if "article_totals" in result and result["article_totals"]:
            existing_arts = {a.article for a in data.article_totals}
            for art in result["article_totals"]:
                if art.article not in existing_arts:
                    data.article_totals.append(art)
                    existing_arts.add(art.article)

        if result.get("grand_total_adopted") is not None:
            data.grand_total_adopted = result["grand_total_adopted"]
        if result.get("grand_total_proposed") is not None:
            data.grand_total_proposed = result["grand_total_proposed"]

        if "enrollment" in result and result["enrollment"]:
            for entry in result["enrollment"]:
                existing = next(
                    (e for e in data.enrollment if e.school == entry.school), None
                )
                if existing:
                    existing.years.update(entry.years)
                else:
                    data.enrollment.append(entry)

        if "reductions" in result and result["reductions"]:
            data.reductions.extend(result["reductions"])

    data.budget_history.sort(key=lambda e: e.fy)
    data.article_totals.sort(key=lambda a: a.article)

    return data


def load_all_handbooks(csv_dir: Path | None = None) -> dict[int, HandbookData]:
    """Load handbook data for all available FYs.

    Returns dict mapping FY number to HandbookData.
    Prefers superintendent handbook; falls back to board-adopted.
    """
    if csv_dir is None:
        csv_dir = Path(__file__).parent.parent.parent / "data" / "RSU 5 Budget Documents" / "csv"

    all_data: dict[int, HandbookData] = {}

    for fy in range(22, 28):
        supt = parse_handbook_for_fy(csv_dir, fy, "superintendent")
        board = parse_handbook_for_fy(csv_dir, fy, "board-adopted")

        if supt.source_files and board.source_files:
            merged = supt
            if not merged.article_totals and board.article_totals:
                merged.article_totals = board.article_totals
            if not merged.budget_history and board.budget_history:
                merged.budget_history = board.budget_history
            if not merged.enrollment and board.enrollment:
                merged.enrollment = board.enrollment
            all_data[fy] = merged
        elif supt.source_files:
            all_data[fy] = supt
        elif board.source_files:
            all_data[fy] = board

    return all_data
