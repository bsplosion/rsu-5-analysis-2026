# RSU 5 Board of Directors — Meeting Minutes Archive

## Overview

This directory contains OCR-processed meeting minutes for Regional School Unit 5
(Freeport, Durham, Pownal, Maine) scraped from the district's official minutes page:

<https://www.rsu5.org/board-of-directors-and-policies/board-agendas-and-minutes/minutes>

The minutes are published as scanned image PDFs (no embedded text). Each PDF is
downloaded, rendered at 300 DPI via PyMuPDF, and OCR'd with Tesseract 5.5 using
`--psm 6` (uniform text block) mode. Post-processing corrects common OCR artifacts
(curly quotes, superscript ordinals, garbled signature blocks).

## Contents

| Path | Description |
|---|---|
| `fetch_minutes.py` | Scraper + OCR pipeline script |
| `pdf/` | Original downloaded PDFs (37 MB, 239 files) |
| `md/` | OCR'd Markdown output (1.3 MB, 239 files) |

## Coverage

- **Date range:** January 2009 — February 25, 2026
- **Meetings:** ~389 unique board meetings (full RSU 5 history)
- **Initially retrieved:** March 12, 2026 (239 meetings from 2015-2026)
- **Expanded:** March 16, 2026 (added ~150 meetings from 2009-2015 archived page)
- **Source pages:** Main minutes page (241 links, 2 duplicates) + archived page (rsu5.org/fs/pages/1279)
- **Known gaps:** 3 archived PDFs returned 404 (Feb 26, 2014; Aug 28, 2013; Aug 14, 2013 — old-format URLs)

## Markdown Format

Each `.md` file includes a citation header:

```
# RSU 5 Board of Directors — Meeting Minutes
**Date:** February 25, 2026
**Source:** [February 25, 2026](https://www.rsu5.org/fs/resource-manager/view/...)
**PDF:** [2-25-26Minutes.pdf](https://resources.finalsite.net/images/...)
**Retrieved:** 2026-03-12
**OCR processed:** 2026-03-12 15:11
```

Followed by the full OCR text of the minutes.

## Known Limitations

- **Summary minutes, not transcripts.** Votes, movers/seconders, and agenda items are
  captured. Discussion substance, public comment content, and budget presentation
  details are not recorded in the original documents.
- **OCR artifacts.** Occasional minor errors remain (e.g., "Schoo!" for "School",
  stray characters near page headers/footers, garbled handwritten signatures on
  older documents where the superintendent was not Tom Gray).
- **No content from presentations.** Budget reviews, facilities reports, and other
  materials presented at meetings are referenced by title only.

## Script Usage

```bash
python fetch_minutes.py                    # last 2 school years
python fetch_minutes.py --after 2015-08-26 # full archive
python fetch_minutes.py --years 5          # last 5 school years
python fetch_minutes.py --reocr            # re-OCR existing PDFs (e.g., after fixing post-processing)
python fetch_minutes.py --dry-run          # list meetings without downloading
```

## Dependencies

- Python 3.10+
- `pymupdf`, `pytesseract`, `Pillow`, `requests`
- [Tesseract OCR](https://github.com/tesseract-ocr/tesseract) system install
  (`winget install tesseract-ocr.tesseract` on Windows)
