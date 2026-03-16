"""Summary sheet: single-FY overview with verification status.

The Summary tab is the first thing a stakeholder sees.  For each per-FY
workbook it shows:
  1. Verification tie-out status against RSU stated totals
  2. Article breakdown for this FY
  3. Key context notes (budget history, reductions, enrollment)
  4. Reconciliation details if mismatches exist
"""

from __future__ import annotations

from openpyxl.workbook.workbook import Workbook

from rsu5.config import cfg
from rsu5.excel.helpers import col_widths, dat, hdr, note, put, sec, source_block, ttl
from rsu5.excel.styles import (
    BOLD,
    CALC_FILL,
    HEADER_FILL,
    MISMATCH_FILL,
    NOTE,
    RESULT_FILL,
    RESULT_FONT,
    TAB_SUMMARY,
    THICK_BOTTOM,
    USD,
    PCT2,
    VERIFIED_FILL,
    WARN_FONT,
)
from rsu5.ingest.data_loader import BudgetData
from rsu5.model import VerifiedBaseline


def build_summary_sheet(
    wb: Workbook,
    fy: int,
    data: BudgetData,
    baseline: VerifiedBaseline | None = None,
) -> None:
    """Create the Summary tab for a single FY workbook."""
    ws = wb.create_sheet("Summary", 0)
    ws.sheet_properties.tabColor = TAB_SUMMARY
    col_widths(ws, [38, 18, 18, 18, 14])

    r = ttl(ws, 1, f"RSU 5 Budget Analysis -- FY{fy}")
    r = note(ws, r + 1, f"Fiscal Year {2000 + fy - 1}-{2000 + fy}")

    adopted = data.adopted_total(fy)
    proposed = data.proposed_total(fy)
    if adopted:
        r = note(ws, r, f"Prior year adopted: ${adopted:,.0f}")
    if proposed:
        r = note(ws, r, f"This year proposed/adopted: ${proposed:,.0f}")
        if adopted:
            pct = (proposed - adopted) / adopted * 100
            r = note(ws, r, f"Year-over-year change: {pct:+.2f}%")
    r += 1

    # Verification status
    targets = cfg.raw.get("verification_targets", {}).get(f"FY{fy}", {})
    art_targets = targets.get("articles", {})
    hb_arts = data.article_totals_from_handbook(fy)
    if not art_targets and hb_arts:
        for ha in hb_arts:
            art_targets[ha.article] = {"proposed": ha.proposed}

    r = sec(ws, r, "Verification Status")
    if art_targets:
        columns = data.all_columns(fy)
        proposed_col = None
        for c in columns:
            if "proposed" in c.lower() or f"fy{fy}" in c.lower():
                proposed_col = c
                break
        if not proposed_col and columns:
            proposed_col = columns[-1]

        total_parsed = sum(
            data.article_total(fy, a, proposed_col or "") for a in range(1, 12)
        )
        total_stated = sum(
            art_targets[a].get("proposed", art_targets[a].get("adopted", 0))
            for a in art_targets
        )
        diff = total_parsed - total_stated
        is_clean = abs(diff) < 1.0

        if is_clean:
            put(ws, r, 1, "VERIFIED", fill=VERIFIED_FILL, font=BOLD)
            put(ws, r, 2, "All articles match RSU stated totals", fill=VERIFIED_FILL)
        else:
            pct = diff / total_stated * 100 if total_stated else 0
            put(ws, r, 1, "UNVERIFIED", fill=MISMATCH_FILL, font=WARN_FONT)
            put(ws, r, 2, f"Variance: ${diff:+,.0f} ({pct:+.3f}%)", fill=MISMATCH_FILL)
        r += 1
        put(ws, r, 1, "Parsed total", fill=CALC_FILL)
        put(ws, r, 2, total_parsed, fmt=USD, fill=CALC_FILL)
        r += 1
        put(ws, r, 1, "RSU stated total", fill=CALC_FILL)
        put(ws, r, 2, total_stated, fmt=USD, fill=CALC_FILL)
    else:
        put(ws, r, 1, "No verification targets available for this FY", fill=HEADER_FILL)
    r += 2

    # Article breakdown
    r = sec(ws, r, "Article Breakdown")
    for i, h in enumerate(["Article", "Name", "Total", "% of Budget"], 1):
        ws.cell(r, i, h)
    hdr(ws, r, 4)
    r += 1

    columns = data.all_columns(fy)
    proposed_col = None
    for c in columns:
        if "proposed" in c.lower() or f"fy{fy}" in c.lower():
            proposed_col = c
            break
    if not proposed_col and columns:
        proposed_col = columns[-1]

    grand = sum(data.article_total(fy, a, proposed_col or "") for a in range(1, 12))

    for art_num in range(1, 12):
        art_cfg = cfg.articles.get(art_num)
        total = data.article_total(fy, art_num, proposed_col or "")
        put(ws, r, 1, f"Art {art_num}", fill=None)
        put(ws, r, 2, art_cfg.name if art_cfg else "", fill=None)
        put(ws, r, 3, total, fmt=USD, fill=None)
        if grand > 0:
            put(ws, r, 4, total / grand, fmt=PCT2, fill=None)
        r += 1

    put(ws, r, 1, "TOTAL", fill=RESULT_FILL, font=BOLD)
    put(ws, r, 2, "", fill=RESULT_FILL)
    put(ws, r, 3, grand, fmt=USD, fill=RESULT_FILL, font=RESULT_FONT)
    put(ws, r, 4, 1.0, fmt=PCT2, fill=RESULT_FILL)
    r += 2

    # Key context
    enrollment = data.enrollment(fy)
    if enrollment:
        r = sec(ws, r, "Enrollment")
        for e in enrollment:
            if e.school == "Grand Total":
                continue
            latest_year = max(e.years.keys()) if e.years else 0
            latest_count = e.years.get(latest_year, 0)
            put(ws, r, 1, e.school, fill=None)
            put(ws, r, 2, latest_count, fill=None)
            r += 1
        total_entry = next((e for e in enrollment if e.school == "Grand Total"), None)
        if total_entry and total_entry.years:
            latest = total_entry.years[max(total_entry.years.keys())]
            put(ws, r, 1, "Total Enrollment", fill=RESULT_FILL, font=BOLD)
            put(ws, r, 2, latest, fill=RESULT_FILL, font=BOLD)
            r += 1
        r += 1

    reductions = data.reductions(fy)
    if reductions:
        r = sec(ws, r, f"FY{fy} Reductions Summary")
        t1 = [rd for rd in reductions if rd.tier == "T1"]
        t2 = [rd for rd in reductions if rd.tier == "T2"]
        t1_total = sum(rd.proposed_amount for rd in t1)
        t2_total = sum(rd.proposed_amount for rd in t2)
        put(ws, r, 1, f"Tier 1: {len(t1)} positions", fill=None)
        put(ws, r, 2, t1_total, fmt=USD, fill=None)
        r += 1
        put(ws, r, 1, f"Tier 2: {len(t2)} positions", fill=None)
        put(ws, r, 2, t2_total, fmt=USD, fill=None)
        r += 1
        total_red = sum(rd.proposed_amount for rd in reductions)
        put(ws, r, 1, "Total reductions", fill=RESULT_FILL, font=BOLD)
        put(ws, r, 2, total_red, fmt=USD, fill=RESULT_FILL, font=BOLD)
        r += 2

    r = source_block(ws, r, [
        "Data parsed from RSU 5 budget documents (rsu5.org/budget).",
        "All totals independently computed from line-item data.",
        "See C-Verification sheet for detailed tie-out.",
    ])
