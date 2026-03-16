"""Parse Maine DOE staffing data from the RSU 5 Staff by FTE Excel file.

The DOE provides historical staffing data (Dec 1 snapshot each year) with
columns: Year Code, SAU OrgId, SAU Name, School OrgId, School Name,
Position TitleId, Position Title, FTE Total.

Data spans 2016-2026 across all RSU 5 schools plus district-wide positions.
"""

from __future__ import annotations

from pathlib import Path

import openpyxl

from rsu5.model import StaffingRecord

_DEFAULT_PATH = (
    Path(__file__).parent.parent.parent
    / "data" / "DOE" / "MDohle RSU 5 Staff by FTE.xlsx"
)

_SCHOOL_NORMALIZE = {
    "Durham Community School": "DCS",
    "Freeport High School": "FHS",
    "Freeport Middle School": "FMS",
    "Mast Landing School": "MLS",
    "Morse Street School": "MSS",
    "Pownal Elementary School": "PES",
}


def parse_doe_staffing(xlsx_path: Path | None = None) -> list[StaffingRecord]:
    """Parse the DOE staffing XLSX into StaffingRecord objects.

    Args:
        xlsx_path: Path to the XLSX file. Defaults to the project's
            data/DOE/MDohle RSU 5 Staff by FTE.xlsx.

    Returns:
        List of StaffingRecord objects, one per (year, school, position) row.
    """
    if xlsx_path is None:
        xlsx_path = _DEFAULT_PATH

    wb = openpyxl.load_workbook(str(xlsx_path), data_only=True, read_only=True)
    ws = wb["Data Table"]

    records: list[StaffingRecord] = []
    header_seen = False

    for row in ws.iter_rows(values_only=True):
        if not header_seen:
            if row and str(row[0]).strip() == "Year Code":
                header_seen = True
            continue

        year_raw = row[0] if row else None
        if not year_raw:
            continue
        try:
            year = int(year_raw)
        except (ValueError, TypeError):
            continue

        school_raw = str(row[4]).strip() if row[4] else ""
        school = _SCHOOL_NORMALIZE.get(school_raw, "District")

        category = str(row[6]).strip() if row[6] else ""
        if not category:
            continue

        fte_raw = row[7] if len(row) > 7 else None
        if fte_raw is None or str(fte_raw).strip() == "":
            fte = 0.0
        else:
            try:
                fte = float(fte_raw)
            except (ValueError, TypeError):
                fte = 0.0

        records.append(StaffingRecord(
            year=year,
            school=school,
            category=category,
            fte=fte,
        ))

    wb.close()
    return records


def staffing_by_year(records: list[StaffingRecord]) -> dict[int, list[StaffingRecord]]:
    """Group records by year."""
    by_year: dict[int, list[StaffingRecord]] = {}
    for r in records:
        by_year.setdefault(r.year, []).append(r)
    return by_year


def staffing_summary(records: list[StaffingRecord]) -> dict[int, dict[str, float]]:
    """Summarize total FTE by year and school.

    Returns {year: {school: total_fte}}.
    """
    summary: dict[int, dict[str, float]] = {}
    for r in records:
        summary.setdefault(r.year, {})
        summary[r.year][r.school] = summary[r.year].get(r.school, 0.0) + r.fte
    return summary
