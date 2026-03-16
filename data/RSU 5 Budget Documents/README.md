# RSU 5 Budget Documents Archive

## Overview

This directory contains official budget documents for Regional School Unit 5
(Freeport, Durham, Pownal, Maine) downloaded from the district's budget pages:

- <https://www.rsu5.org/budget/budget-2021-2022> (FY22)
- <https://www.rsu5.org/budget/budget-2022-2023> (FY23)
- <https://www.rsu5.org/budget/FY24> (FY24)
- <https://www.rsu5.org/budget/FY25> (FY25)
- <https://www.rsu5.org/budget/FY26> (FY26)
- <https://www.rsu5.org/budget/FY27> (FY27, in progress)

## Contents

| Path | Description |
|---|---|
| `fetch_budgets.py` | Download script with hardcoded document manifest |
| `extract_budgets.py` | Text + table extraction pipeline |
| `pdf/` | Downloaded PDFs organized as `FY{yy}-{slug}.pdf` (28 files, 39 MB) |
| `md/` | Full text content extracted to Markdown (28 files, 2.5 MB) |
| `csv/` | Tabular data extracted to CSV via column-position analysis (724 files, 4.4 MB) |

## Document Types

For each fiscal year, we capture the final/adopted versions of these documents:

| Slug | Description |
|---|---|
| `citizens-adopted-budget` | Voter-facing summary mailed to households before the BVR |
| `board-adopted-budget` | Budget articles as adopted by the Board of Directors |
| `board-adopted-handbook` | Full handbook version of the board-adopted budget |
| `superintendent-handbook` | Superintendent's Recommended Budget Handbook (detailed line items) |
| `budget-articles` | Warrant articles broken out by cost center |
| `budget-overview` | High-level budget summary (FY22 only) |
| `budget-worksheet` | Detailed worksheet (FY22 only) |

### What's in the Budget Handbook?

The Superintendent's Recommended Budget Handbook is the most comprehensive
document. It typically includes:

- Line-item budget by cost center (instruction, support, operations, etc.)
- Year-over-year comparisons (prior year actual, current year budget, proposed)
- Revenue sources (state subsidy, local appropriation, tuition, grants)
- Tax impact analysis by town (Freeport, Durham, Pownal)
- Enrollment projections
- Staffing summaries
- Warrant article detail

## Coverage

| FY | School Year | Status | PDFs | CSVs |
|----|-------------|--------|------|------|
| 22 | 2021-2022 | Adopted | 4 | ~60 |
| 23 | 2022-2023 | Adopted | 5 | ~120 |
| 24 | 2023-2024 | Adopted | 5 | ~120 |
| 25 | 2024-2025 | Adopted | 5 | ~120 |
| 26 | 2025-2026 | Adopted | 5 | ~160 |
| 27 | 2026-2027 | In progress | 4 | ~140 |

## Script Usage

### Downloading PDFs

```bash
python fetch_budgets.py              # download all FY22-FY27 documents
python fetch_budgets.py --dry-run    # list documents without downloading
python fetch_budgets.py --fy 26      # download only FY26 documents
python fetch_budgets.py --force      # re-download even if files exist
```

### Extracting content

```bash
python extract_budgets.py              # extract all PDFs to md/ and csv/
python extract_budgets.py --fy 26      # extract only FY26 documents
python extract_budgets.py --force      # re-extract existing files
python extract_budgets.py --single FY26-superintendent-handbook.pdf
python extract_budgets.py --dry-run    # list PDFs without extracting
```

## Extraction Method

The budget PDFs are native text (not scanned images), so no OCR is needed.

**Markdown extraction** uses `pdfplumber.extract_text()` to capture the full
content of each page, with a metadata header (fiscal year, source URL, PDF
filename, extraction timestamp).

**Table extraction** uses a column-detection algorithm based on word positions:

1. Extract all words with their bounding-box coordinates
2. Classify each word as numeric (contains digits, `$`, `%`, etc.) or text
3. Cluster the **right-edge (x1)** positions of numeric words -- right-aligned
   dollar amounts in the same column share an x1 position
4. Cluster the **left-edge (x0)** positions of text words
5. Merge column positions and compute column boundaries
6. Assign words to columns; output as CSV

Pages where <10% of words are numeric are treated as text-only (no CSV).
Contiguous pages with the same column count are merged into a single CSV,
with repeated header rows deduplicated.

### CSV naming

- `FY26-superintendent-handbook.csv` -- single table spanning the document
- `FY26-superintendent-handbook-p03.csv` -- table starting at page 3
  (when a document has multiple tables with different column structures)

## Dependencies

- Python 3.10+
- `requests` (for fetch_budgets.py)
- `pdfplumber` (for extract_budgets.py)

## Notes

- The FY22 budget page used a different structure (Overview + Worksheet
  instead of Board-Adopted Budget + Handbook). Later years standardized
  on the Handbook + Articles format.
- FY27 documents will be updated as the budget cycle progresses through
  board adoption, annual budget meeting, and budget validation referendum.
- All documents are hosted on the RSU 5 Finalsite resource manager and
  may be updated by the district at any time.
- Column counts may vary across tables within the same PDF because different
  budget articles have different numbers of prior-year comparison columns.
  The extraction preserves whatever structure is on each page.
