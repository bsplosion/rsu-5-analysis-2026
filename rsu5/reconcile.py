"""Independent verification engine.

Computes totals from parsed line items and compares against the RSU's
stated summary rows.  Produces a ``VerifiedBaseline`` -- the single gate
between raw data and downstream analysis.

Usage::

    from rsu5.ingest.data_loader import BudgetData
    from rsu5.reconcile import reconcile

    data = BudgetData.load()
    baselines = reconcile(data)      # dict[int, VerifiedBaseline]
    for fy, bl in baselines.items():
        print(f"FY{fy}: {'CLEAN' if bl.is_clean else 'MISMATCHES'}")
"""

from __future__ import annotations

from collections import defaultdict

from rsu5.config import cfg
from rsu5.ingest.data_loader import BudgetData
from rsu5.model import (
    BudgetLineItem,
    ReconciliationResult,
    SummaryRow,
    VerifiedBaseline,
)


# Tolerance for floating-point comparison (1 cent)
_TOLERANCE = 0.01


def _group_items(
    items: list[BudgetLineItem], key: str
) -> dict[str, list[BudgetLineItem]]:
    """Group line items by a field name (e.g. ``"cost_center"``)."""
    groups: dict[str, list[BudgetLineItem]] = defaultdict(list)
    for li in items:
        groups[getattr(li, key)].append(li)
    return groups


def _sum_amounts(items: list[BudgetLineItem], column: str) -> float:
    return sum(li.amounts.get(column, 0.0) for li in items)


def _reconcile_level(
    items: list[BudgetLineItem],
    summaries: list[SummaryRow],
    level: str,
    item_key: str,
    columns: list[str],
) -> list[ReconciliationResult]:
    """Compare computed totals against stated totals for one hierarchy level."""
    results: list[ReconciliationResult] = []

    # Group stated summaries by code
    stated_by_code: dict[str, SummaryRow] = {}
    for sr in summaries:
        if sr.level == level:
            code = sr.code
            # Prefer the last occurrence (some docs restate totals)
            stated_by_code[code] = sr

    # Group items and compute
    grouped = _group_items(items, item_key)

    for code, group_items in sorted(grouped.items()):
        stated = stated_by_code.get(code)
        if stated is None:
            continue

        for col in columns:
            if col in ("Dollar Difference", "Percent Difference"):
                continue
            if col not in stated.amounts:
                continue

            computed = _sum_amounts(group_items, col)
            stated_val = stated.amounts[col]
            diff = computed - stated_val
            is_match = abs(diff) <= _TOLERANCE

            results.append(
                ReconciliationResult(
                    level=level,
                    code=code,
                    label=stated.label if stated else code,
                    column=col,
                    computed=computed,
                    stated=stated_val,
                    difference=diff,
                    is_match=is_match,
                    contributing_accounts=(
                        []
                        if is_match
                        else [li.account_code for li in group_items]
                    ),
                )
            )

    return results


def reconcile_fy(
    data: BudgetData,
    fy: int,
) -> VerifiedBaseline:
    """Reconcile one fiscal year's data and produce a VerifiedBaseline."""
    items = data.items_for_fy(fy)
    summaries = data.summaries_for_fy(fy)
    columns = data.all_columns(fy)

    doc_type = ""
    if items:
        doc_type = items[0].doc_type

    results: list[ReconciliationResult] = []
    notes: list[str] = []

    if not summaries:
        notes.append(
            f"FY{fy}: No summary rows available for reconciliation. "
            f"Data is UNVERIFIED."
        )
    else:
        # Reconcile by article (group items, compare against
        # cost_center-level summaries which are the most reliable)
        for article_num in sorted(cfg.articles):
            art_items = data.items_by_article(fy, article_num)
            art_summaries = [
                s for s in summaries if s.article == article_num
            ]

            if not art_items:
                continue

            # Cost-center level
            cc_results = _reconcile_level(
                art_items, art_summaries,
                "cost_center", "cost_center", columns,
            )
            results.extend(cc_results)

            # Program level
            prog_results = _reconcile_level(
                art_items, art_summaries,
                "program", "program", columns,
            )
            results.extend(prog_results)

            # Function level
            func_results = _reconcile_level(
                art_items, art_summaries,
                "function", "function", columns,
            )
            results.extend(func_results)

    is_clean = all(r.is_match for r in results) if results else False
    mismatches = [r for r in results if not r.is_match]

    if mismatches:
        notes.append(
            f"FY{fy}: {len(mismatches)} reconciliation mismatches "
            f"out of {len(results)} checks."
        )
    elif results:
        notes.append(
            f"FY{fy}: All {len(results)} checks passed. Baseline verified."
        )

    return VerifiedBaseline(
        fy=fy,
        doc_type=doc_type,
        line_items=list(items),
        reconciliation_results=results,
        is_clean=is_clean,
        notes=notes,
    )


def reconcile(
    data: BudgetData,
    fys: list[int] | None = None,
) -> dict[int, VerifiedBaseline]:
    """Reconcile all fiscal years and return verified baselines.

    Args:
        data: Loaded budget data.
        fys: Fiscal years to reconcile.  Defaults to all loaded.
    """
    if fys is None:
        fys = data.fiscal_years()

    baselines: dict[int, VerifiedBaseline] = {}

    for fy in sorted(fys):
        bl = reconcile_fy(data, fy)
        baselines[fy] = bl

        # Print summary
        status = "VERIFIED" if bl.is_clean else "UNVERIFIED"
        n_checks = len(bl.reconciliation_results)
        n_mismatches = len(bl.mismatches)

        print(f"  FY{fy} [{status}]: {n_checks} checks, {n_mismatches} mismatches")
        if bl.mismatches:
            for m in bl.mismatches[:5]:
                print(
                    f"    {m.level} {m.code} / {m.column}: "
                    f"computed={m.computed:,.2f} stated={m.stated:,.2f} "
                    f"diff={m.difference:+,.2f}"
                )
            if n_mismatches > 5:
                print(f"    ... and {n_mismatches - 5} more")
        for note in bl.notes:
            print(f"    {note}")

    return baselines
