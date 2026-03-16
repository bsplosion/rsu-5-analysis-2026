"""
Export 'RSU 5 Planning and PES reconciliation 2026.md' to a professionally
formatted PDF using markdown-it-py (markdown -> HTML) and xhtml2pdf (HTML -> PDF).

xhtml2pdf natively repeats <thead> rows across page breaks, supports @page
CSS rules for headers/footers/page numbers, and is pure Python (ReportLab).

Usage:
    python export_pdf.py                          # defaults
    python export_pdf.py -i other.md -o out.pdf   # custom paths

Requirements:  pip install -r requirements.txt
"""

import argparse
import re
from pathlib import Path

from markdown_it import MarkdownIt
from xhtml2pdf import pisa

DEFAULT_INPUT = "RSU 5 Planning and PES reconciliation 2026.md"
DEFAULT_OUTPUT = "RSU 5 Planning and PES Reconciliation 2026.pdf"

CSS = """
@page {
    size: letter;
    margin: 1in 1in 1.15in 1in;

    @frame footer {
        -pdf-frame-content: page-footer;
        bottom: 0.4in;
        margin-left: 1in;
        margin-right: 1in;
        height: 0.4in;
    }
}

#page-footer {
    font-family: Helvetica, Arial, sans-serif;
    font-size: 8pt;
    color: #888888;
    text-align: center;
    border-top: 0.5pt solid #cccccc;
    padding-top: 4pt;
}

/* ── Base typography ──────────────────────────────────────── */

body {
    font-family: Georgia, 'Times New Roman', serif;
    font-size: 10.5pt;
    line-height: 1.5;
    color: #1a1a1a;
}

/* ── Cover page title ─────────────────────────────────────── */

.cover-title {
    font-family: Helvetica, Arial, sans-serif;
    font-size: 28pt;
    font-weight: bold;
    color: #1a3c5e;
    margin-top: 2.2in;
    margin-bottom: 6pt;
    padding-bottom: 6pt;
    border-bottom: 3pt solid #1a3c5e;
    line-height: 1.15;
}

.cover-subtitle {
    font-family: Helvetica, Arial, sans-serif;
    font-size: 12pt;
    color: #4a4a4a;
    margin-bottom: 0.5in;
}

/* ── Headings ─────────────────────────────────────────────── */

h1 {
    font-family: Helvetica, Arial, sans-serif;
    font-size: 20pt;
    font-weight: bold;
    color: #1a3c5e;
    border-bottom: 2pt solid #1a3c5e;
    padding-bottom: 3pt;
    margin-top: 18pt;
    margin-bottom: 10pt;
}

h2 {
    font-family: Helvetica, Arial, sans-serif;
    font-size: 14pt;
    font-weight: bold;
    color: #27548a;
    border-bottom: 0.75pt solid #b0c4de;
    padding-bottom: 2pt;
    margin-top: 14pt;
    margin-bottom: 8pt;
}

h3 {
    font-family: Helvetica, Arial, sans-serif;
    font-size: 12pt;
    font-weight: bold;
    color: #2d6496;
    margin-top: 12pt;
    margin-bottom: 6pt;
}

h4 {
    font-family: Helvetica, Arial, sans-serif;
    font-size: 10.5pt;
    font-weight: bold;
    color: #3a3a3a;
    margin-top: 10pt;
    margin-bottom: 4pt;
}

/* ── Paragraphs and lists ─────────────────────────────────── */

p {
    margin-top: 4pt;
    margin-bottom: 4pt;
    text-align: justify;
}

ul, ol {
    margin-top: 4pt;
    margin-bottom: 6pt;
}

li {
    margin-bottom: 2pt;
}

/* ── Tables ────────────────────────────────────────────────── */

table {
    width: 100%;
    border-collapse: collapse;
    margin-top: 6pt;
    margin-bottom: 10pt;
    font-family: Helvetica, Arial, sans-serif;
    font-size: 9pt;
    line-height: 1.3;
}

thead {
    display: table-header-group;
}

thead th {
    background-color: #1a3c5e;
    color: #ffffff;
    font-weight: bold;
    text-align: left;
    padding: 5pt 6pt;
    border-bottom: 2pt solid #1a3c5e;
}

tbody td {
    padding: 4pt 6pt;
    border-bottom: 0.5pt solid #d0d0d0;
    vertical-align: top;
}

tbody tr:nth-child(even) td {
    background-color: #f0f4f8;
}

/* ── Horizontal rules ─────────────────────────────────────── */

hr {
    border: none;
    border-top: 1pt solid #cccccc;
    margin-top: 12pt;
    margin-bottom: 12pt;
}

/* ── Emphasis and inline ──────────────────────────────────── */

strong {
    font-weight: bold;
}

em {
    font-style: italic;
    color: #333333;
}

code {
    font-family: Courier, monospace;
    font-size: 9pt;
    background-color: #f0f0f0;
    padding: 1pt 3pt;
}

/* ── Footnote references ──────────────────────────────────── */

sup {
    font-size: 7.5pt;
    color: #27548a;
}

/* ── Block quotes ─────────────────────────────────────────── */

blockquote {
    margin: 6pt 0;
    padding: 4pt 8pt 4pt 10pt;
    border-left: 3pt solid #27548a;
    background-color: #f4f7fb;
    color: #2a2a2a;
    font-size: 10pt;
}

blockquote p {
    margin: 2pt 0;
}
"""


def preprocess(text: str) -> str:
    """Massage markdown for better PDF rendering."""
    text = re.sub(r'\[\^(\d+)\]', r'<sup>[\1]</sup>', text)
    # Replace empty table cells (| |) with a non-breaking space so
    # xhtml2pdf's column-width calculator doesn't produce zero-width columns.
    text = re.sub(r'\| \|', '| &nbsp; |', text)
    return text


def md_to_html(md_text: str) -> str:
    """Convert markdown to HTML body using markdown-it-py."""
    md = MarkdownIt('commonmark').enable('table')
    return md.render(md_text)


def build_cover(title: str, subtitle: str) -> str:
    """Build cover page HTML."""
    return f"""
    <div class="cover-title">{title}</div>
    <div class="cover-subtitle">{subtitle}</div>
    <div style="page-break-after: always;"></div>
    """


def extract_and_replace_cover(html: str):
    """Extract the first h1 + following paragraph as cover material,
    replace them with styled cover elements, and return modified HTML.
    """
    title_match = re.search(r'<h1>(.*?)</h1>', html, re.DOTALL)
    if not title_match:
        return html

    title = title_match.group(1).strip()
    after_title = html[title_match.end():]

    subtitle_match = re.match(r'\s*<p>(.*?)</p>', after_title, re.DOTALL)
    subtitle = subtitle_match.group(1).strip() if subtitle_match else ""
    rest_start = title_match.end() + (subtitle_match.end() if subtitle_match else 0)

    # Skip the first <hr> after subtitle (it's the cover separator)
    remaining = html[rest_start:]
    hr_match = re.match(r'\s*<hr\s*/?>', remaining)
    if hr_match:
        rest_start += hr_match.end()

    cover = build_cover(title, subtitle)
    return cover + html[rest_start:]


def insert_page_breaks(html: str) -> str:
    """Insert page breaks before Part headings (h1 tags containing 'Part')."""
    return re.sub(
        r'<h1>(Part \d+)',
        r'<div style="page-break-before: always;"></div><h1>\1',
        html,
    )


def build_full_html(md_text: str) -> str:
    """Build the complete HTML document."""
    md_text = preprocess(md_text)
    body_html = md_to_html(md_text)
    body_html = extract_and_replace_cover(body_html)
    body_html = insert_page_breaks(body_html)

    return f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<style>
{CSS}
</style>
</head>
<body>
{body_html}

<div id="page-footer">
    <span class="page-num"></span>
</div>
</body>
</html>"""


def export_pdf(input_path: str, output_path: str) -> None:
    md_text = Path(input_path).read_text(encoding='utf-8')
    html_string = build_full_html(md_text)

    with open(output_path, 'wb') as f:
        status = pisa.CreatePDF(html_string, dest=f)

    if status.err:
        print(f"Errors during PDF creation: {status.err}")
    else:
        print(f"PDF written to: {output_path}")


def main():
    parser = argparse.ArgumentParser(
        description="Export RSU 5 markdown report to professionally formatted PDF.",
    )
    parser.add_argument(
        '-i', '--input',
        default=DEFAULT_INPUT,
        help=f"Input markdown file (default: {DEFAULT_INPUT})",
    )
    parser.add_argument(
        '-o', '--output',
        default=DEFAULT_OUTPUT,
        help=f"Output PDF path (default: {DEFAULT_OUTPUT})",
    )
    args = parser.parse_args()
    export_pdf(args.input, args.output)


if __name__ == '__main__':
    main()
