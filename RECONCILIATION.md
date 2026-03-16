# Reconciliation Policy

## What reconciliation means in this project

Reconciliation is the process of independently computing budget totals from
parsed line-item data and comparing them against the RSU's own stated totals.
It produces a **verified baseline** -- a dataset that is confirmed to match the
district's published numbers -- which all downstream analysis builds on.

## The two purposes

1. **Independent verification**: Catching errors in the district's published
   documents, in our parsing, or in our calculations.  When we present analysis
   to the board, we can demonstrate that we started from the exact same numbers
   they did, and any differences come from our analysis choices, not from data
   errors.

2. **Establishing the exact shared baseline**: When we develop budget proposals,
   scenarios, or projections, we need to start from the same starting point as
   the district.  The verified baseline ensures our numbers are
   apples-to-apples.  No one can dismiss the analysis by saying "you're using
   different numbers" because we can show that our starting point matches
   theirs.

## How to run reconciliation

```bash
# Run with reconciliation output and skip Excel generation
python build_workbook.py --dry-run

# Run for a specific fiscal year
python build_workbook.py --dry-run --fy 27

# Full build (includes reconciliation report in console + Summary tab)
python build_workbook.py
```

The console output shows reconciliation status for each fiscal year.  The
generated Excel workbook includes a **Summary** tab with reconciliation status,
check counts, and a detailed list of any mismatches.

## How to interpret results

Each reconciliation check compares a computed total (sum of parsed line items)
against a stated total (the RSU's own subtotal row from the budget document).
Checks are performed at three hierarchy levels:

- **Cost center**: Sum of all line items for a cost center within an article
- **Program**: Sum of all line items for a program within an article
- **Function**: Sum of all line items for a function within an article

A **VERIFIED** fiscal year means all checks pass (computed equals stated within
$0.01 tolerance).  An **UNVERIFIED** fiscal year has at least one mismatch.

## The standing rule

**No new data enters the model without passing reconciliation.**

This applies to:

- **New fiscal years**: When FY28 data becomes available, ingesting it requires
  running reconciliation before any analysis is built on top of it.
- **New data sources**: Revenue data, enrollment projections, or staffing figures
  must be verified against the district's published versions before use.
- **Alternative proposals**: Any budget scenario or projection must trace back to
  a `VerifiedBaseline`.  The output should make this lineage visible.
- **Manual overrides**: When `budget_config.yaml` values are updated,
  reconciliation must re-run to confirm downstream totals still hold.

## How to handle discrepancies

When reconciliation finds mismatches:

1. **Investigate the source**: Check the original CSV file and the budget PDF.
   Is the mismatch from a parsing error (fixable in the parser) or a genuine
   inconsistency in the district's documents?

2. **Document the finding**: If it's a parser gap, add it to the parser's known
   issues.  If it's a data inconsistency, note it in the `VerifiedBaseline`
   notes and in analysis output.

3. **Never silently absorb differences**: Every mismatch must be acknowledged.
   The Summary tab in the workbook shows them.  Console output prints them.
   If analysis proceeds on unverified data, it carries a visible warning.

4. **Fix or explain**: Either fix the root cause (parser improvement, data
   correction) or document why the difference exists and confirm it doesn't
   affect downstream analysis conclusions.

## Current reconciliation status

As of the initial implementation:

| FY | Status | Checks | Mismatches | Notes |
|----|--------|--------|------------|-------|
| 22 | UNVERIFIED | 0 | 0 | No summary rows in FY22 format |
| 23 | UNVERIFIED | 25 | 11 | Small rounding diffs (~$1-5) |
| 24 | UNVERIFIED | 25 | 7 | Function 2600 column alignment |
| 25 | UNVERIFIED | 570 | 187 | Summary row column alignment |
| 26 | UNVERIFIED | 570 | 219 | Summary row column alignment |
| 27 | UNVERIFIED | 565 | 188 | Two article-level gaps: Art 1 ($13K), Art 4 ($29.5K) |

FY27 article-level validation against `create_excel.py` hardcoded totals:
9 of 11 articles match exactly.  Total gap: $42,565 of $47.3M (0.09%).

Most FY25-27 mismatches stem from summary-row column alignment in the parser
(the stated totals are being mapped to wrong columns).  The line-item data
itself parses accurately; it's the reconciliation reference points that need
tuning.

## Files involved

- `rsu5/reconcile.py` -- Reconciliation engine
- `rsu5/ingest/budget_csv_parser.py` -- CSV parser (where parsing fixes go)
- `rsu5/model.py` -- `VerifiedBaseline` and `ReconciliationResult` data models
- `build_workbook.py` -- Entry point that runs reconciliation
- `budget_config.yaml` -- Configuration (column layouts, cost centers, etc.)
- `.cursor/rules/reconciliation.mdc` -- AI-facing rule for future dev sessions
