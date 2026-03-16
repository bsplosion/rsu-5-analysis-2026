"""Orchestrate data ingestion across all fiscal years and data sources.

Loads budget line-item CSVs, DOE staffing data, and superintendent handbook
data into a unified query interface.

Usage::

    from rsu5.ingest.data_loader import BudgetData

    data = BudgetData.load()           # ingest all configured FYs
    data = BudgetData.load(fys=[27])   # ingest FY27 only

    # Budget queries
    items = data.items_by_article(27, 1)
    total = data.article_total(27, 1, "FY27 Proposed")

    # Handbook queries
    hb = data.handbook(27)
    enrollment = data.enrollment(27)
    history = data.budget_history(27)

    # DOE staffing queries
    staffing = data.staffing_for_year(2024)
    fte = data.school_fte(2024, "DCS")
"""

from __future__ import annotations

from collections import defaultdict
from pathlib import Path

from rsu5.config import cfg
from rsu5.ingest.budget_csv_parser import parse_document_csvs
from rsu5.ingest.doe_staffing_parser import parse_doe_staffing, staffing_summary
from rsu5.ingest.handbook_parser import (
    HandbookData,
    load_all_handbooks,
    EnrollmentEntry,
    BudgetHistoryEntry,
    ArticleTotal,
    ReductionItem,
)
from rsu5.model import BudgetLineItem, StaffingRecord, SummaryRow


_CSV_DIR = (
    Path(__file__).resolve().parent.parent.parent
    / "data"
    / "RSU 5 Budget Documents"
    / "csv"
)


class BudgetData:
    """Holds all parsed budget data with query methods."""

    def __init__(self) -> None:
        self.line_items: list[BudgetLineItem] = []
        self.summary_rows: list[SummaryRow] = []

        # DOE staffing
        self._staffing: list[StaffingRecord] = []
        self._staffing_by_year: dict[int, list[StaffingRecord]] = {}
        self._staffing_fte: dict[int, dict[str, float]] = {}

        # Handbook data per FY
        self._handbooks: dict[int, HandbookData] = {}

        # Indexes built after loading
        self._by_fy: dict[int, list[BudgetLineItem]] = defaultdict(list)
        self._by_fy_article: dict[tuple[int, int], list[BudgetLineItem]] = (
            defaultdict(list)
        )
        self._summaries_by_fy: dict[int, list[SummaryRow]] = defaultdict(list)

    def _index(self) -> None:
        self._by_fy.clear()
        self._by_fy_article.clear()
        self._summaries_by_fy.clear()
        for li in self.line_items:
            self._by_fy[li.fy].append(li)
            self._by_fy_article[(li.fy, li.article)].append(li)
        for sr in self.summary_rows:
            self._summaries_by_fy[sr.fy].append(sr)

    # ── Query interface ────────────────────────────────────────

    def fiscal_years(self) -> list[int]:
        return sorted(self._by_fy.keys())

    def items_for_fy(self, fy: int) -> list[BudgetLineItem]:
        return self._by_fy.get(fy, [])

    def items_by_article(self, fy: int, article: int) -> list[BudgetLineItem]:
        return self._by_fy_article.get((fy, article), [])

    def items_by_cost_center(
        self, fy: int, cost_center: str, article: int | None = None
    ) -> list[BudgetLineItem]:
        source = (
            self.items_by_article(fy, article)
            if article is not None
            else self.items_for_fy(fy)
        )
        return [li for li in source if li.cost_center == cost_center]

    def summaries_for_fy(self, fy: int) -> list[SummaryRow]:
        return self._summaries_by_fy.get(fy, [])

    def article_total(self, fy: int, article: int, column: str) -> float:
        return sum(
            li.amounts.get(column, 0.0)
            for li in self.items_by_article(fy, article)
        )

    def cost_center_total(
        self,
        fy: int,
        cost_center: str,
        column: str,
        article: int | None = None,
    ) -> float:
        items = self.items_by_cost_center(fy, cost_center, article)
        return sum(li.amounts.get(column, 0.0) for li in items)

    def all_columns(self, fy: int) -> list[str]:
        """Return all distinct column names present in data for *fy*."""
        cols: set[str] = set()
        for li in self.items_for_fy(fy):
            cols.update(li.amounts.keys())
        try:
            pref = cfg.preferred_doc(fy)
            doc_type_slug = pref.split("-", 1)[1]
            ordered = cfg.column_layout(doc_type_slug, fy)
            return [c for c in ordered if c in cols]
        except (KeyError, IndexError):
            return sorted(cols)

    # ── Handbook queries ───────────────────────────────────────

    def handbook(self, fy: int) -> HandbookData | None:
        """Return handbook data for a fiscal year, or None."""
        return self._handbooks.get(fy)

    def enrollment(self, fy: int) -> list[EnrollmentEntry]:
        """Return enrollment entries for a fiscal year."""
        hb = self._handbooks.get(fy)
        return hb.enrollment if hb else []

    def budget_history(self, fy: int) -> list[BudgetHistoryEntry]:
        """Return 10-year budget history from a specific FY's handbook."""
        hb = self._handbooks.get(fy)
        return hb.budget_history if hb else []

    def article_totals_from_handbook(self, fy: int) -> list[ArticleTotal]:
        """Return article totals from handbook for a fiscal year."""
        hb = self._handbooks.get(fy)
        return hb.article_totals if hb else []

    def reductions(self, fy: int) -> list[ReductionItem]:
        """Return reduction items for a fiscal year."""
        hb = self._handbooks.get(fy)
        return hb.reductions if hb else []

    def adopted_total(self, fy: int) -> float | None:
        """Return the RSU's stated adopted total for a fiscal year.

        First checks handbook data, then falls back to budget_config.yaml.
        """
        hb = self._handbooks.get(fy)
        if hb and hb.grand_total_adopted:
            return hb.grand_total_adopted
        totals = cfg.raw.get("budget_totals", {})
        fy_key = f"FY{fy}"
        if fy_key in totals:
            return totals[fy_key].get("adopted")
        return None

    def proposed_total(self, fy: int) -> float | None:
        """Return the RSU's stated proposed total for a fiscal year."""
        hb = self._handbooks.get(fy)
        if hb and hb.grand_total_proposed:
            return hb.grand_total_proposed
        totals = cfg.raw.get("budget_totals", {})
        fy_key = f"FY{fy}"
        if fy_key in totals:
            return totals[fy_key].get("proposed")
        return None

    # ── DOE staffing queries ───────────────────────────────────

    def staffing_for_year(self, year: int) -> list[StaffingRecord]:
        """Return all staffing records for a calendar year."""
        return self._staffing_by_year.get(year, [])

    def school_fte(self, year: int, school: str) -> float:
        """Return total FTE for a school in a given calendar year."""
        return self._staffing_fte.get(year, {}).get(school, 0.0)

    def total_fte(self, year: int) -> float:
        """Return total FTE across all schools for a calendar year."""
        return sum(self._staffing_fte.get(year, {}).values())

    def staffing_years(self) -> list[int]:
        """Return sorted list of calendar years with staffing data."""
        return sorted(self._staffing_by_year.keys())

    # ── Loading ────────────────────────────────────────────────

    @classmethod
    def load(
        cls,
        csv_dir: Path | None = None,
        fys: list[int] | None = None,
        preferred_only: bool = True,
    ) -> "BudgetData":
        """Ingest all data sources for the specified fiscal years.

        Loads budget CSVs, DOE staffing XLSX, and handbook CSVs.

        Args:
            csv_dir: Directory containing extracted CSVs.
            fys: Fiscal years to load (2-digit).  Defaults to all configured.
            preferred_only: If True, only load the preferred document per FY.
        """
        csv_dir = csv_dir or _CSV_DIR
        data = cls()

        if fys is None:
            fys = [
                int(k.replace("FY", ""))
                for k in cfg.raw.get("line_item_documents", {})
            ]

        # Budget line-item CSVs
        for fy in sorted(fys):
            if preferred_only:
                prefixes = [cfg.preferred_doc(fy)]
            else:
                prefixes = cfg.line_item_docs(fy)

            for prefix in prefixes:
                doc_type = cfg.doc_type_from_prefix(prefix)

                slug = prefix.split("-", 1)[1] if "-" in prefix else prefix
                try:
                    col_names = cfg.column_layout(slug, fy)
                except KeyError:
                    print(
                        f"  WARN: No column layout for {prefix} "
                        f"(slug={slug!r}, fy={fy}); skipping."
                    )
                    continue

                items, summaries = parse_document_csvs(
                    csv_dir, prefix, fy, doc_type, col_names
                )

                data.line_items.extend(items)
                data.summary_rows.extend(summaries)

                print(
                    f"  {prefix}: {len(items)} line items, "
                    f"{len(summaries)} summary rows"
                )

        data._index()

        # DOE staffing
        try:
            data._staffing = parse_doe_staffing()
            from rsu5.ingest.doe_staffing_parser import staffing_by_year
            data._staffing_by_year = staffing_by_year(data._staffing)
            data._staffing_fte = staffing_summary(data._staffing)
            print(f"  DOE staffing: {len(data._staffing)} records, "
                  f"{len(data._staffing_by_year)} years")
        except Exception as e:
            print(f"  WARN: Could not load DOE staffing: {e}")

        # Handbook data
        try:
            data._handbooks = load_all_handbooks(csv_dir)
            loaded = [fy for fy in data._handbooks if data._handbooks[fy].source_files]
            print(f"  Handbooks: loaded for FY{', FY'.join(str(f) for f in sorted(loaded))}")
        except Exception as e:
            print(f"  WARN: Could not load handbook data: {e}")

        return data
