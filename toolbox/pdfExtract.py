"""
pdf_to_markdown_txt.py

Extract readable text from a PDF and save it to a Markdown-formatted .txt file.
Designed to produce clean, structured text suitable for LLM context.

What it does:
- Parses PDF text with layout info using PyMuPDF (pymupdf).
- Heuristically detects headings by larger font sizes and renders them as #, ##, ...
- Converts bullet and numbered lists to Markdown lists with indentation.
- Merges wrapped lines into paragraphs, fixing hyphenation across line breaks (e.g., "inter-\nnational" -> "international").
- Removes repeated page headers/footers (detected across pages; on by default).
- Optionally includes "## Page N" markers.
- Escapes Markdown special characters to produce stable output.

Usage:
  python pdf_to_markdown_txt.py input.pdf [-o output.txt]
Options:
  --no-page-headings     Do not add "## Page N" headings.
  --keep-headers         Keep repeated page headers/footers (do not attempt removal).
  --max-pages N          Limit number of pages processed (0 = no limit).
  --encoding ENC         Output encoding (default: utf-8).

Requirements:
  pip install pymupdf
"""

import argparse
import os
import re
import sys
from collections import Counter, defaultdict
from typing import List, Dict, Any, Tuple, Optional

try:
    import fitz  # PyMuPDF
except ImportError as e:
    print("Missing dependency: pymupdf. Install it with:\n  pip install pymupdf", file=sys.stderr)
    sys.exit(1)


# ------------------------------ Markdown helpers ------------------------------


def md_escape(text: str) -> str:
    """
    Escape Markdown special characters in inline text.
    We do not escape '-' to preserve hyphenated words naturally.
    """
    if not text:
        return ""
    text = text.replace("\\", "\\\\")
    chars = r"`*_{}[]()#+.!>"
    return re.sub(r"([%s])" % re.escape(chars), r"\\\1", text)


def strip_trailing_hyphen_md(s: str) -> str:
    """
    Remove a trailing hyphen from a Markdown string even if it's just before
    closing emphasis markers. Example: '**com-**' -> '**com**'
    """
    if not s:
        return s
    i = len(s) - 1
    # skip spaces and emphasis/backticks at end
    while i >= 0 and s[i] in " *_`":
        i -= 1
    if i >= 0 and s[i] == "-":
        return s[:i] + s[i + 1 :]
    return s


# ------------------------------ PDF text processing ------------------------------


def font_is_bold(font_name: str) -> bool:
    if not font_name:
        return False
    f = font_name.lower()
    return any(k in f for k in ("bold", "black", "heavy", "demi", "semibold", "extrabold", "ultrabold"))


def font_is_italic(font_name: str) -> bool:
    if not font_name:
        return False
    f = font_name.lower()
    return any(k in f for k in ("italic", "oblique"))


def render_span_to_md(text: str, bold: bool, italic: bool) -> str:
    if not text:
        return ""
    s = md_escape(text)
    if bold and italic:
        return f"***{s}***"
    elif bold:
        return f"**{s}**"
    elif italic:
        return f"*{s}*"
    return s


def render_spans_from_offset(spans: List[Dict[str, Any]], skip_chars: int = 0) -> str:
    """
    Render a list of spans to Markdown starting after `skip_chars` characters of plain text.
    """
    out_parts: List[str] = []
    remaining_skip = skip_chars
    for sp in spans:
        t: str = sp.get("text", "") or ""
        if not t:
            continue
        if remaining_skip > 0:
            if len(t) <= remaining_skip:
                remaining_skip -= len(t)
                continue
            else:
                t = t[remaining_skip:]
                remaining_skip = 0
        bold = sp.get("_bold", False)
        italic = sp.get("_italic", False)
        out_parts.append(render_span_to_md(t, bold, italic))
    return "".join(out_parts)


def unify_size(size: float) -> float:
    """Round font size to nearest 0.5pt to stabilize histograms."""
    return round(size * 2) / 2.0


def normalize_for_repeat(s: str) -> str:
    """
    Normalize a line for header/footer repetition detection:
    - lowercase
    - collapse whitespace
    - replace digit sequences with '#'
    """
    s = (s or "").strip().lower()
    s = re.sub(r"\s+", " ", s)
    s = re.sub(r"\d+", "#", s)
    return s


def detect_bullet(plain: str) -> Tuple[bool, bool, int]:
    """
    Detect list bullets / ordered markers at the start of a line.
    Returns (is_list, is_ordered, prefix_char_count_to_strip)
    """
    if not plain:
        return (False, False, 0)
    s = plain.lstrip()
    lead_ws = len(plain) - len(s)

    # Unordered bullets (common glyphs)
    m = re.match(r"^([•‣◦▪\-\–\—·●♦\*])\s+", s)
    if m:
        return (True, False, lead_ws + m.end())

    # Ordered: 1.  1)  (1)  a.  a)  (a)  i.  i)  (i)
    m = re.match(r"^(\(?(\d+|[a-z]{1,3}|[ivxlcdm]{1,6})[\.\)]\)?|\(\d+\)|\([a-z]{1,3}\)|\([ivxlcdm]{1,6}\))\s+",
                 s, flags=re.I)
    if m:
        return (True, True, lead_ws + m.end())

    return (False, False, 0)


def get_page_lines(page) -> Tuple[List[Dict[str, Any]], float, float]:
    """
    Extract text lines with spans and geometry from a page.
    Returns (lines, page_width, page_height)
    Each line dict includes:
      - x0, y0, y1
      - spans: [{text, size, font, _bold, _italic}]
      - plain: concatenated plain text
      - size_avg: average (char-weighted) font size
      - block_id: incremental block index on the page
    """
    page_dict = page.get_text("dict")
    pw = float(page.rect.width)
    ph = float(page.rect.height)

    lines_out: List[Dict[str, Any]] = []
    block_id = -1

    for blk in page_dict.get("blocks", []):
        if blk.get("type", 0) != 0:
            continue  # not a text block
        block_id += 1
        for ln in blk.get("lines", []):
            line_bbox = ln.get("bbox", blk.get("bbox", (0, 0, 0, 0)))
            x0, y0, x1, y1 = [float(v) for v in line_bbox]
            spans_in = ln.get("spans", [])
            spans: List[Dict[str, Any]] = []
            total_chars = 0
            size_weighted = 0.0
            bold_chars = 0

            for sp in spans_in:
                t = sp.get("text", "") or ""
                if not t:
                    continue
                size = float(sp.get("size", 0.0) or 0.0)
                font = str(sp.get("font", "") or "")
                b = font_is_bold(font)
                i = font_is_italic(font)
                spans.append({"text": t, "size": size, "font": font, "_bold": b, "_italic": i})
                n = len(t)
                total_chars += n
                size_weighted += size * n
                if b:
                    bold_chars += n

            if not spans:
                continue

            size_avg = size_weighted / max(total_chars, 1)
            plain = "".join(sp["text"] for sp in spans)
            # collapse internal whitespace, but keep spaces (for bullets detection we'll use the raw)
            plain = plain.replace("\r", "").replace("\n", "")

            lines_out.append({
                "x0": x0, "y0": y0, "y1": y1,
                "spans": spans,
                "plain": plain,
                "size_avg": size_avg,
                "block_id": block_id,
            })

    return lines_out, pw, ph


def compute_body_font_size(all_lines: List[Dict[str, Any]]) -> float:
    """
    Compute a representative body font size as the mode (by char-weight) of rounded sizes.
    Fallback to median of sizes if needed.
    """
    size_counter: Counter = Counter()
    sizes_all: List[float] = []
    for ln in all_lines:
        s = unify_size(float(ln.get("size_avg", 0.0) or 0.0))
        if s <= 0:
            continue
        sizes_all.append(s)
        # approximate char-weight by length of text
        w = max(1, len(ln.get("plain", "") or ""))
        size_counter[s] += w

    if not size_counter:
        return 12.0

    body_size, _ = size_counter.most_common(1)[0]
    return float(body_size)


def heading_level_for_size(size: float, body: float) -> Optional[int]:
    """
    Map a font size to a Markdown heading level relative to the body size.
    Returns 1..6 for headings, or None for normal text.
    """
    if body <= 0:
        return None
    ratio = size / body
    # Conservative thresholds; adjust as needed
    if ratio >= 2.1:
        return 1
    if ratio >= 1.8:
        return 2
    if ratio >= 1.6:
        return 3
    if ratio >= 1.45:
        return 4
    if ratio >= 1.3:
        return 5
    if ratio >= 1.18:
        return 6
    return None


# ------------------------------ Main conversion logic ------------------------------


def pdf_to_markdown(input_path: str,
                    include_page_headings: bool = True,
                    remove_headers_footers: bool = True,
                    max_pages: int = 0) -> str:
    """
    Convert a PDF to Markdown string.
    """
    doc = fitz.open(input_path)
    try:
        n_pages_total = doc.page_count
        n_pages = n_pages_total if not max_pages or max_pages <= 0 else min(max_pages, n_pages_total)

        all_pages_lines: List[List[Dict[str, Any]]] = []
        page_dims: List[Tuple[float, float]] = []
        all_lines_flat: List[Dict[str, Any]] = []

        # First pass: collect lines and candidates for header/footer
        top_counts: Counter = Counter()
        bot_counts: Counter = Counter()
        left_by_page: List[float] = []

        for pi in range(n_pages):
            page = doc.load_page(pi)
            lines, pw, ph = get_page_lines(page)
            all_pages_lines.append(lines)
            page_dims.append((pw, ph))
            all_lines_flat.extend(lines)

            # Left margin baseline for indent calculation
            left_min = min((ln["x0"] for ln in lines), default=0.0)
            left_by_page.append(left_min)

            if remove_headers_footers:
                top_cut = ph * 0.12
                bot_cut = ph * 0.88
                for ln in lines:
                    norm = normalize_for_repeat(ln["plain"])
                    if not norm or len(norm) < 3:
                        continue
                    if ln["y0"] <= top_cut:
                        top_counts[norm] += 1
                    elif ln["y1"] >= bot_cut:
                        bot_counts[norm] += 1

        # Determine body font size
        body_size = compute_body_font_size(all_lines_flat)

        # Determine repeated headers/footers
        reps_top, reps_bot = set(), set()
        if remove_headers_footers and n_pages >= 3:
            thr = max(2, int(round(0.5 * n_pages)))  # seen on >= ~half pages
            reps_top = {k for k, c in top_counts.items() if c >= thr}
            reps_bot = {k for k, c in bot_counts.items() if c >= thr}

        out_lines: List[str] = []
        out_lines.append("")

        # Second pass: render Markdown
        for pi in range(n_pages):
            pw, ph = page_dims[pi]
            page_lines = all_pages_lines[pi]
            left_base = left_by_page[pi]

            if include_page_headings:
                out_lines.append(f"## Page {pi + 1}")
                out_lines.append("")

            # Iterate in source order, grouping paragraphs within blocks
            current_block_id = None
            para_parts: List[str] = []
            prev_plain: Optional[str] = None

            def flush_paragraph():
                nonlocal para_parts
                if para_parts:
                    # Join parts already contain leading spaces when needed
                    out_lines.append("".join(para_parts).strip())
                    out_lines.append("")
                    para_parts = []

            for ln in page_lines:
                # Header/footer removal
                if remove_headers_footers:
                    top_cut = ph * 0.12
                    bot_cut = ph * 0.88
                    norm = normalize_for_repeat(ln["plain"])
                    if ln["y0"] <= top_cut and norm in reps_top:
                        continue
                    if ln["y1"] >= bot_cut and norm in reps_bot:
                        continue

                # New block -> flush paragraph
                if current_block_id is None:
                    current_block_id = ln["block_id"]
                elif ln["block_id"] != current_block_id:
                    flush_paragraph()
                    current_block_id = ln["block_id"]
                    prev_plain = None

                plain = ln["plain"].strip()
                if not plain:
                    # Empty line -> paragraph break
                    flush_paragraph()
                    prev_plain = None
                    continue

                # List detection
                is_list, is_ordered, skip_chars = detect_bullet(ln["plain"])
                if is_list:
                    flush_paragraph()
                    # Indentation level based on left indent
                    indent_level = int(round(max(0.0, (ln["x0"] - left_base) / 18.0)))
                    indent_level = max(0, min(8, indent_level))
                    indent = "  " * indent_level
                    marker = "1. " if is_ordered else "- "
                    content_md = render_spans_from_offset(ln["spans"], skip_chars=skip_chars).strip()
                    if not content_md:
                        continue
                    out_lines.append(f"{indent}{marker}{content_md}")
                    prev_plain = None
                    continue

                # Heading detection
                lvl = heading_level_for_size(float(ln["size_avg"]), body_size)
                if lvl is not None and len(plain) <= 160:
                    flush_paragraph()
                    text_md = render_spans_from_offset(ln["spans"], skip_chars=0).strip()
                    out_lines.append(f"{'#' * lvl} {text_md}")
                    out_lines.append("")
                    prev_plain = None
                    continue

                # Normal paragraph line: merge with previous
                rendered = render_spans_from_offset(ln["spans"], skip_chars=0).strip()
                if not rendered:
                    continue

                if not para_parts:
                    para_parts.append(rendered)
                else:
                    # Hyphenation join: previous ended with '-' and this starts with lowercase letter
                    prev_plain_eff = (prev_plain or "").rstrip()
                    this_plain = ln["plain"].lstrip()
                    if prev_plain_eff.endswith("-") and re.match(r"^[a-z]", this_plain):
                        # Remove trailing hyphen in previous MD and concatenate with no space
                        para_parts[-1] = strip_trailing_hyphen_md(para_parts[-1])
                        para_parts.append(rendered.lstrip())
                    else:
                        # Ordinary space between wrapped lines
                        para_parts.append(" " + rendered.lstrip())

                prev_plain = ln["plain"]

            # Flush any remainder at page end
            flush_paragraph()

        md = "\n".join(out_lines).rstrip() + "\n"
        return md
    finally:
        doc.close()


# ------------------------------ CLI ------------------------------


def main(inputPath):
    ap = argparse.ArgumentParser(description="Extract text from a PDF into Markdown-formatted .txt for LLM context.")
    ap.add_argument("--no-page-headings", action="store_true", help="Do not include '## Page N' headings.")
    ap.add_argument("--keep-headers", action="store_true", help="Keep repeated page headers/footers (do not attempt removal).")
    ap.add_argument("--max-pages", type=int, default=0, help="Limit number of pages to process (0 = no limit).")
    ap.add_argument("--encoding", default="utf-8", help="Output file encoding (default: utf-8).")
    args = ap.parse_args()

    input_path = inputPath
    if not os.path.isfile(input_path):
        print(f"Error: File not found: {input_path}", file=sys.stderr)
        sys.exit(1)

    _, ext = os.path.splitext(input_path)
    if ext.lower() != ".pdf":
        print("Error: Only .pdf files are supported.", file=sys.stderr)
        sys.exit(1)

    try:
        md = pdf_to_markdown(
            input_path=input_path,
            include_page_headings=(not args.no_page_headings),
            remove_headers_footers=(not args.keep_headers),
            max_pages=args.max_pages or 0,
        )
    except Exception as e:
        print(f"Conversion failed: {e}", file=sys.stderr)
        sys.exit(1)

    return md

