"""C-Verification sheet: parsed article totals vs RSU stated totals.

Shows per-article match/mismatch, overall accuracy, and data sources.
This is the core tie-out that establishes the verified baseline.
"""

from __future__ import annotations

from openpyxl.workbook import Workbook

from rsu5.config import cfg
from rsu5.excel.helpers import col_widths, dat, hdr, note, put, sec, source_block, ttl
from rsu5.excel.styles import (
    BOLD,
    CALC_FILL,
    HEADER_FILL,
    MISMATCH_FILL,
    PCT2,
    RESULT_FILL,
    RESULT_FONT,
    TAB_CALC,
    USD,
    VERIFIED_FILL,
    WARN_FONT,
)
from rsu5.ingest.data_loader import BudgetData
from rsu5.model import VerifiedBaseline


def build_verification(wb: Workbook, fy: int, data: BudgetData,
                       baseline: VerifiedBaseline | None = None) -> None:
    """Build the C-Verification sheet."""
    ws = wb.create_sheet("C-Verification")
    ws.sheet_properties.tabColor = TAB_CALC
    col_widths(ws, [32, 18, 18, 18, 10, 30])

    r = ttl(ws, 1, f"VERIFICATION: FY{fy} Budget Tie-Out")
    r = note(ws, r + 1, f"Compares parsed line-item totals against RSU stated figures.")
    r = note(ws, r, "A clean tie-out means our data matches the district's numbers exactly.")
    r += 1

    targets = cfg.raw.get("verification_targets", {}).get(f"FY{fy}", {})
    art_targets = targets.get("articles", {})
    stated_total = targets.get("proposed_total") or targets.get("adopted_total")

    hb_arts = data.article_totals_from_handbook(fy)
    if not art_targets and hb_arts:
        for ha in hb_arts:
            art_targets[ha.article] = {
                "adopted": ha.adopted,
                "proposed": ha.proposed,
            }
        if not stated_total:
            stated_total = data.proposed_total(fy)

    columns = data.all_columns(fy)
    proposed_col = None
    for c in columns:
        if "proposed" in c.lower() or "fy" in c.lower():
            proposed_col = c
            break
    if not proposed_col and columns:
        proposed_col = columns[-1]

    r = sec(ws, r, "Per-Article Comparison")
    headers = ["Article", "Parsed Total", "RSU Stated", "Difference", "Match?", "Notes"]
    for i, h in enumerate(headers, 1):
        ws.cell(r, i, h)
    hdr(ws, r, len(headers))
    r += 1

    total_parsed = 0.0
    total_stated = 0.0
    all_match = True

    for art_num in range(1, 12):
        art_cfg = cfg.articles.get(art_num)
        label = f"Art {art_num}" + (f" - {art_cfg.name}" if art_cfg else "")

        parsed = data.article_total(fy, art_num, proposed_col) if proposed_col else 0.0
        total_parsed += parsed

        stated = None
        if art_num in art_targets:
            stated = art_targets[art_num].get("proposed") or art_targets[art_num].get("adopted")

        if stated is not None:
            total_stated += stated
            diff = parsed - stated
            is_match = abs(diff) < 1.0
            if not is_match:
                all_match = False

            put(ws, r, 1, label, fill=CALC_FILL)
            put(ws, r, 2, parsed, fmt=USD, fill=CALC_FILL)
            put(ws, r, 3, stated, fmt=USD, fill=CALC_FILL)
            put(ws, r, 4, diff, fmt=USD, fill=VERIFIED_FILL if is_match else MISMATCH_FILL)
            put(ws, r, 5, "YES" if is_match else "NO",
                fill=VERIFIED_FILL if is_match else MISMATCH_FILL,
                font=BOLD if not is_match else None)
            if not is_match and abs(diff) > 0:
                pct = diff / stated * 100 if stated else 0
                put(ws, r, 6, f"{pct:+.2f}% ({diff:+,.0f})", fill=MISMATCH_FILL)
        else:
            put(ws, r, 1, label, fill=CALC_FILL)
            put(ws, r, 2, parsed, fmt=USD, fill=CALC_FILL)
            put(ws, r, 3, "N/A", fill=CALC_FILL)
            put(ws, r, 4, "", fill=CALC_FILL)
            put(ws, r, 5, "N/A", fill=CALC_FILL)
            put(ws, r, 6, "No stated total available", fill=CALC_FILL)
        r += 1

    r += 1
    put(ws, r, 1, "TOTAL", fill=RESULT_FILL, font=BOLD)
    put(ws, r, 2, total_parsed, fmt=USD, fill=RESULT_FILL, font=BOLD)
    if total_stated:
        diff = total_parsed - total_stated
        is_match = abs(diff) < 1.0
        put(ws, r, 3, total_stated, fmt=USD, fill=RESULT_FILL, font=BOLD)
        put(ws, r, 4, diff, fmt=USD,
            fill=VERIFIED_FILL if is_match else MISMATCH_FILL, font=BOLD)
        put(ws, r, 5, "YES" if is_match else "NO",
            fill=VERIFIED_FILL if is_match else MISMATCH_FILL,
            font=RESULT_FONT)
        if not is_match:
            pct = diff / total_stated * 100
            put(ws, r, 6, f"{pct:+.3f}% variance", fill=MISMATCH_FILL)
    r += 2

    r = sec(ws, r, "Verification Status")
    if all_match and total_stated:
        put(ws, r, 1, "VERIFIED -- All articles match RSU stated totals",
            fill=VERIFIED_FILL, font=BOLD)
    elif total_stated:
        mismatch_count = sum(
            1 for a in range(1, 12)
            if a in art_targets and abs(
                data.article_total(fy, a, proposed_col or "") -
                (art_targets[a].get("proposed") or art_targets[a].get("adopted") or 0)
            ) >= 1.0
        )
        put(ws, r, 1,
            f"UNVERIFIED -- {mismatch_count} article(s) have discrepancies",
            fill=MISMATCH_FILL, font=WARN_FONT)
    else:
        put(ws, r, 1, "NO VERIFICATION TARGETS -- Cannot verify this FY",
            fill=HEADER_FILL)
    r += 2

    sources = [
        f"Parsed from: {proposed_col or 'N/A'} column in budget CSV data",
        f"RSU stated totals from: budget_config.yaml verification_targets.FY{fy}",
    ]
    hb = data.handbook(fy)
    if hb and hb.source_files:
        sources.append(f"Handbook sources: {', '.join(hb.source_files[:3])}")
    r = source_block(ws, r, sources)
