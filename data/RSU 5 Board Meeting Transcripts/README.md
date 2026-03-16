# RSU 5 Board of Directors — Meeting Transcripts

## Overview

This directory contains auto-generated YouTube caption transcripts for RSU 5
board meeting recordings from the
[RSU5 Livestream](https://www.youtube.com/channel/UC97VXXLhRFRjSPv1wfo1ACA/)
YouTube channel.

Transcripts are extracted via the `youtube-transcript-api` Python library,
then post-processed with regex corrections for known proper noun errors
in the auto-captions.

## Contents

| Path | Description |
|---|---|
| `fetch_transcripts.py` | Channel scraper + transcript extraction pipeline |
| `md/` | Timestamped Markdown transcripts |

## Coverage

- **Date range:** September 13, 2017 — March 11, 2026
- **Meetings:** ~177 unique board meeting recordings
- **Total runtime:** ~442 hours of video
- **Source:** YouTube auto-generated captions (English)

## Markdown Format

Each `.md` file includes a citation header and timestamped paragraphs:

```
# RSU 5 Board of Directors — Meeting Transcript
**Date:** February 25, 2026
**Video:** [2026-02-25 RSU 5 Board Meeting](https://www.youtube.com/watch?v=...)
**Duration:** 3:54:51
**Words:** 37,285
**Retrieved:** 2026-03-12
**Source:** YouTube auto-generated captions (regex-corrected)

---

[00:00] Good evening. Welcome to the RSU 5 board of directors meeting...

[00:54] Okay. So, do I have a motion for consideration?...

[20:10] Rich Brereton, Pownal, Libby Road. I have a first grader
in Pownal Elementary School now...
```

Speaker changes (detected via `>>` markers in YouTube captions) start new
paragraphs. Timestamps reference the original video for verification.

## Regex Corrections Applied

| Auto-caption error | Corrected to |
|---|---|
| Panel, Pel, Powell | Pownal |
| pees, Pees | PES |
| rc5, RC5, RSU5 | RSU 5 |

## Known Limitations

- **No speaker identification.** YouTube auto-captions don't label speakers.
  Speaker changes are sometimes marked with `>>` but not always. Speakers
  often self-identify ("I'm Matt Alieri") or are addressed by the chair
  ("Thank you Rich"), providing context cues.
- **Word-level errors.** Auto-captions occasionally mishear words, especially
  proper nouns not covered by the regex corrections, names of community
  members, and technical/financial terms.
- **No punctuation in some passages.** YouTube's auto-captioning quality
  varies; some segments have good punctuation, others have none.
- **Very recent recordings** (within ~24 hours of upload) may not have
  auto-captions available yet.

## Complementary Data

These transcripts pair with the OCR'd summary minutes in
`../RSU 5 Board of Directors Meeting minutes/`. The minutes capture formal
votes and agenda structure; the transcripts capture discussion substance,
public comment content, and budget presentation details.

## Script Usage

```bash
python fetch_transcripts.py                    # all available transcripts
python fetch_transcripts.py --after 2024-01-01 # only after a date
python fetch_transcripts.py --dry-run          # list available videos
python fetch_transcripts.py --refetch          # re-download existing
```

Note: YouTube may rate-limit requests if too many are made in rapid
succession. The script includes exponential backoff and retry logic.
If rate-limited, wait 1-2 hours and re-run; already-fetched transcripts
will be skipped automatically.

## Dependencies

- Python 3.10+
- `youtube-transcript-api`, `scrapetube`, `yt-dlp`
