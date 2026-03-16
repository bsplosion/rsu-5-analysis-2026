"""
Fetch YouTube auto-generated transcripts for RSU 5 board meeting recordings
from the RSU5 Livestream channel, apply regex corrections for known proper
nouns, and output structured Markdown with timestamps.

Usage:
    python fetch_transcripts.py                    # all available transcripts
    python fetch_transcripts.py --after 2024-01-01 # only after a date
    python fetch_transcripts.py --dry-run          # list available videos
    python fetch_transcripts.py --refetch          # re-download existing transcripts
"""

import argparse
import random
import re
import sys
import time
from datetime import datetime, date
from pathlib import Path

import scrapetube
from youtube_transcript_api import YouTubeTranscriptApi

CHANNEL_ID = "UC97VXXLhRFRjSPv1wfo1ACA"
CHANNEL_URL = f"https://www.youtube.com/channel/{CHANNEL_ID}"
VIDEO_URL = "https://www.youtube.com/watch?v={}"

DATA_DIR = Path(__file__).parent
MD_DIR = DATA_DIR / "md"

DATE_YMD = re.compile(r"(\d{4})\s*[-/]\s*(\d{1,2})\s*[-/]\s*(\d{1,2})")
DATE_MDY = re.compile(r"(\d{1,2})\s*[-/:]\s*(\d{1,2})\s*[-/:]\s*(\d{2,4})")

PROPER_NOUN_FIXES = {
    r"\bPanel\b": "Pownal",
    r"\bpanel\b": "Pownal",
    r"\bPel\b": "Pownal",
    r"\bPowell\b(?!\s+Street)": "Pownal",
    r"\bpowell\b(?!\s+street)": "Pownal",
    r"\bPES\b": "PES",
    r"\b[Pp]ees\b": "PES",
    r"\brc5\b": "RSU 5",
    r"\bRC5\b": "RSU 5",
    r"\bRSU5\b": "RSU 5",
    r"\brsu5\b": "RSU 5",
    r"\bMSS\b": "MSS",
    r"\bDCS\b": "DCS",
    r"\bFMS\b": "FMS",
    r"\bFHS\b": "FHS",
    r"\bECSE\b": "ECSE",
    r"\bCDS\b": "CDS",
}


def parse_video_date(title: str) -> date | None:
    # Try YYYY-MM-DD first (used in recent streams)
    m = DATE_YMD.search(title)
    if m:
        try:
            return date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
        except ValueError:
            pass

    # Try M/D/YYYY or M:D:YYYY (used in older videos)
    m = DATE_MDY.search(title)
    if m:
        month, day, year = int(m.group(1)), int(m.group(2)), int(m.group(3))
        if year < 100:
            year += 2000
        try:
            return date(year, month, day)
        except ValueError:
            pass

    return None


def format_timestamp(seconds: float) -> str:
    h = int(seconds // 3600)
    m = int((seconds % 3600) // 60)
    s = int(seconds % 60)
    if h > 0:
        return f"{h}:{m:02d}:{s:02d}"
    return f"{m:02d}:{s:02d}"


def format_duration(seconds: float) -> str:
    h = int(seconds // 3600)
    m = int((seconds % 3600) // 60)
    s = int(seconds % 60)
    return f"{h}:{m:02d}:{s:02d}"


def parse_duration_str(s: str) -> int:
    parts = s.split(":")
    if len(parts) == 3:
        return int(parts[0]) * 3600 + int(parts[1]) * 60 + int(parts[2])
    if len(parts) == 2:
        return int(parts[0]) * 60 + int(parts[1])
    return 0


def fix_proper_nouns(text: str) -> str:
    for pattern, replacement in PROPER_NOUN_FIXES.items():
        text = re.sub(pattern, replacement, text)
    return text


def format_transcript_markdown(
    snippets: list,
    video_id: str,
    title: str,
    meeting_date: date,
    duration_str: str,
) -> str:
    """Build a Markdown document from transcript snippets."""

    total_duration = max(s.start + s.duration for s in snippets) if snippets else 0
    total_words = sum(len(s.text.split()) for s in snippets)

    header = (
        f"# RSU 5 Board of Directors — Meeting Transcript\n"
        f"**Date:** {meeting_date.strftime('%B %d, %Y')}  \n"
        f"**Video:** [{title}]({VIDEO_URL.format(video_id)})  \n"
        f"**Duration:** {duration_str}  \n"
        f"**Words:** {total_words:,}  \n"
        f"**Retrieved:** {datetime.now().strftime('%Y-%m-%d')}  \n"
        f"**Source:** YouTube auto-generated captions (regex-corrected)  \n"
        f"\n---\n\n"
    )

    # Build timestamped paragraphs grouped by ~60 second intervals
    paragraphs = []
    current_para = []
    para_start_time = 0

    for snippet in snippets:
        if not current_para:
            para_start_time = snippet.start

        text = snippet.text.strip()
        if not text:
            continue

        # Detect speaker change markers
        if ">>" in text:
            # Flush current paragraph before speaker change
            if current_para:
                ts = format_timestamp(para_start_time)
                content = fix_proper_nouns(" ".join(current_para))
                paragraphs.append(f"[{ts}] {content}")
                current_para = []
                para_start_time = snippet.start

            # Split on >> and process each part
            parts = [p.strip() for p in text.split(">>") if p.strip()]
            for j, part in enumerate(parts):
                if j > 0 and current_para:
                    ts = format_timestamp(para_start_time)
                    content = fix_proper_nouns(" ".join(current_para))
                    paragraphs.append(f"[{ts}] {content}")
                    current_para = []
                    para_start_time = snippet.start
                current_para.append(part)
        else:
            current_para.append(text)

        # Break paragraphs every ~60 seconds
        if snippet.start - para_start_time >= 60 and current_para:
            ts = format_timestamp(para_start_time)
            content = fix_proper_nouns(" ".join(current_para))
            paragraphs.append(f"[{ts}] {content}")
            current_para = []

    if current_para:
        ts = format_timestamp(para_start_time)
        content = fix_proper_nouns(" ".join(current_para))
        paragraphs.append(f"[{ts}] {content}")

    body = "\n\n".join(paragraphs)
    return header + body + "\n"


def discover_meeting_videos() -> list[dict]:
    """Scrape the RSU5 Livestream channel for all board meeting videos."""
    all_items = []

    for content_type in ("streams", "videos"):
        items = list(scrapetube.get_channel(CHANNEL_ID, content_type=content_type))
        for v in items:
            vid = v["videoId"]
            title = ""
            if "title" in v and "runs" in v["title"]:
                title = v["title"]["runs"][0]["text"]
            elif "title" in v:
                title = v["title"].get("simpleText", "")

            duration = v.get("lengthText", {}).get("simpleText", "0:00")

            is_meeting = any(
                kw in title.lower()
                for kw in ("board", "meeting", "budget")
            )
            if not is_meeting:
                continue

            meeting_date = parse_video_date(title)
            if meeting_date is None:
                continue

            all_items.append({
                "video_id": vid,
                "title": title,
                "date": meeting_date,
                "duration": duration,
                "duration_secs": parse_duration_str(duration),
            })

    # Deduplicate by date (keep longer recording if same date)
    by_date: dict[date, dict] = {}
    for item in all_items:
        d = item["date"]
        if d not in by_date or item["duration_secs"] > by_date[d]["duration_secs"]:
            by_date[d] = item
    deduped = sorted(by_date.values(), key=lambda x: x["date"], reverse=True)
    return deduped


def main():
    parser = argparse.ArgumentParser(description="Fetch RSU 5 board meeting transcripts from YouTube")
    parser.add_argument("--after", type=str, default=None,
                        help="Only fetch transcripts after this date (YYYY-MM-DD)")
    parser.add_argument("--refetch", action="store_true",
                        help="Re-download existing transcripts")
    parser.add_argument("--dry-run", action="store_true",
                        help="List available videos without downloading")
    parser.add_argument("--delay", type=float, default=8.0,
                        help="Base seconds between requests (default: 8.0, jitter adds 0-100%%)")
    args = parser.parse_args()

    MD_DIR.mkdir(parents=True, exist_ok=True)

    print("Discovering meeting videos on RSU5 Livestream channel...")
    meetings = discover_meeting_videos()
    print(f"  Found {len(meetings)} meeting recordings")

    if args.after:
        cutoff = date.fromisoformat(args.after)
        meetings = [m for m in meetings if m["date"] >= cutoff]
        print(f"  {len(meetings)} after {args.after}")

    print()

    if args.dry_run:
        for m in meetings:
            print(f"  {m['date'].isoformat()}  {m['duration']:>10}  {m['title']}")
        return

    def jittered_sleep(base: float) -> float:
        """Sleep for base + random upward jitter (0-100% of base)."""
        actual = base + random.uniform(0, base)
        time.sleep(actual)
        return actual

    ytt_api = YouTubeTranscriptApi()
    success = 0
    skipped = 0
    failed = []
    base_delay = args.delay
    backoff = base_delay

    for i, m in enumerate(meetings):
        tag = f"[{i+1}/{len(meetings)}]"
        md_path = MD_DIR / f"{m['date'].isoformat()}.md"

        if md_path.exists() and not args.refetch:
            skipped += 1
            continue

        try:
            transcript = ytt_api.fetch(m["video_id"])
            snippets = transcript.snippets
            words = sum(len(s.text.split()) for s in snippets)
            print(f"  {tag} {m['date']}  {m['duration']:>10}  {words:>6} words  {m['title']}")

            md_content = format_transcript_markdown(
                snippets, m["video_id"], m["title"], m["date"], m["duration"],
            )
            md_path.write_text(md_content, encoding="utf-8")
            success += 1
            backoff = base_delay
            jittered_sleep(base_delay)

        except Exception as e:
            err_type = type(e).__name__
            err_str = str(e)
            if "IpBlocked" in err_type or "RequestBlocked" in err_type or "IpBlocked" in err_str:
                backoff = min(backoff * 2, 300)
                wait = backoff + random.uniform(0, backoff * 0.5)
                print(f"  {tag} {m['date']}  Rate limited — backing off {wait:.0f}s", file=sys.stderr)
                time.sleep(wait)
                # Retry once after backoff
                try:
                    transcript = ytt_api.fetch(m["video_id"])
                    snippets = transcript.snippets
                    words = sum(len(s.text.split()) for s in snippets)
                    print(f"  {tag} {m['date']}  {m['duration']:>10}  {words:>6} words  {m['title']}  (retry)")
                    md_content = format_transcript_markdown(
                        snippets, m["video_id"], m["title"], m["date"], m["duration"],
                    )
                    md_path.write_text(md_content, encoding="utf-8")
                    success += 1
                    backoff = base_delay
                    jittered_sleep(base_delay)
                    continue
                except Exception:
                    pass
                failed.append(m)
            elif "NoTranscript" in err_type or "TranscriptsDisabled" in err_type or "not retrievable" in err_str.lower():
                print(f"  {tag} {m['date']}  No transcript available", file=sys.stderr)
                failed.append(m)
                jittered_sleep(base_delay)
            else:
                err_msg = err_str[:120]
                print(f"  {tag} {m['date']}  FAILED: {err_msg}", file=sys.stderr)
                failed.append(m)
                jittered_sleep(base_delay)

    if skipped:
        print(f"  ({skipped} already fetched, skipped)")
    print(f"\nDone. {success} transcripts saved, {len(failed)} failed.")
    if failed:
        print("Failed videos:")
        for m in failed:
            print(f"  {m['date']}  {m['video_id']}  {m['title']}")
    print(f"Output: {MD_DIR}")


if __name__ == "__main__":
    main()
