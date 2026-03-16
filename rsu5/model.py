"""Data model for the RSU 5 budget analysis pipeline.

Every number that flows through the system is represented by one of these
dataclasses.  The ``VerifiedBaseline`` is the single gate between raw
ingestion and downstream analysis -- nothing reaches the Excel output
layer without passing through reconciliation first.
"""

from __future__ import annotations

from dataclasses import dataclass, field


@dataclass
class BudgetLineItem:
    """A single account-code row from a budget document.

    Attributes:
        fy: Two-digit fiscal year (22-27).
        doc_type: Source document kind -- ``"adopted"``, ``"proposed"``,
            ``"worksheet"``, or ``"citizens"``.
        article: Budget article number (1-11).
        fund: Fund segment of the account code (e.g. ``"1000"``).
        program: Program segment (e.g. ``"1100"``).
        function: Function segment (e.g. ``"1000"``).
        object_code: Object segment (e.g. ``"51010"``).
        cost_center: Cost-center segment (e.g. ``"010"``).
        description: Human-readable line-item label.
        amounts: Mapping of column header to dollar value, e.g.
            ``{"FY23 Actual": 1965610.00, "FY27 Proposed": 2712435.00}``.
        source_file: CSV filename this row was parsed from.
        source_page: Page number within the source PDF, if known.
    """

    fy: int
    doc_type: str
    article: int
    fund: str
    program: str
    function: str
    object_code: str
    cost_center: str
    description: str
    amounts: dict[str, float] = field(default_factory=dict)
    source_file: str = ""
    source_page: int = 0

    @property
    def account_code(self) -> str:
        """Canonical dot-separated account code."""
        return (
            f"{self.fund}.{self.program}.{self.function}"
            f".{self.object_code}.{self.cost_center}"
        )


@dataclass
class SummaryRow:
    """An RSU-stated subtotal row (Program:, Function:, Cost Center:, or article total).

    These are captured during parsing so we can compare our independently
    computed sums against the district's own totals.
    """

    fy: int
    doc_type: str
    level: str  # "program", "function", "cost_center", or "article"
    code: str  # e.g. "1100", "1000", "010"
    label: str  # e.g. "ELEMENTARY PROGRAMS - 1100"
    amounts: dict[str, float] = field(default_factory=dict)
    article: int = 0
    source_file: str = ""


@dataclass
class ReconciliationResult:
    """Outcome of comparing our computed total against the RSU's stated total
    for one hierarchy level (program, function, or cost center).
    """

    level: str
    code: str
    label: str
    column: str
    computed: float
    stated: float
    difference: float
    is_match: bool
    contributing_accounts: list[str] = field(default_factory=list)


@dataclass
class VerifiedBaseline:
    """Reconciled dataset confirmed to match the district's numbers.

    This is the single source of truth that all analysis builds on.
    If ``is_clean`` is False, downstream outputs carry a visible
    "unreconciled" warning.
    """

    fy: int
    doc_type: str
    line_items: list[BudgetLineItem] = field(default_factory=list)
    reconciliation_results: list[ReconciliationResult] = field(
        default_factory=list
    )
    is_clean: bool = True
    notes: list[str] = field(default_factory=list)

    @property
    def mismatches(self) -> list[ReconciliationResult]:
        return [r for r in self.reconciliation_results if not r.is_match]

    def article_total(self, article: int, column: str) -> float:
        """Sum all line items for *article* in the given *column*."""
        return sum(
            li.amounts.get(column, 0.0)
            for li in self.line_items
            if li.article == article
        )

    def cost_center_total(
        self, cost_center: str, column: str, article: int | None = None
    ) -> float:
        """Sum line items for a cost center, optionally filtered by article."""
        return sum(
            li.amounts.get(column, 0.0)
            for li in self.line_items
            if li.cost_center == cost_center
            and (article is None or li.article == article)
        )


@dataclass
class StaffingRecord:
    """A single row from the Maine DOE staffing data.

    Attributes:
        year: Calendar year of the Dec 1 snapshot.
        school: School name or code (e.g. ``"PES"``).
        category: Position category (e.g. ``"Classroom Teacher"``).
        fte: Full-time equivalent count.
    """

    year: int
    school: str
    category: str
    fte: float
