#!/usr/bin/env python3
"""
docx_to_markdown_txt.py

A Python program that extracts text content from a Word (.docx) document
and saves it to a .txt file in Markdown format.

It will:
- Prefer using Pandoc via pypandoc if available (best fidelity).
- Fallback to a pure-Python converter using python-docx that handles:
  - Headings (Heading 1..6 -> #..######)
  - Paragraphs
  - Bold/Italic runs
  - Hyperlinks
  - Lists (bullet and numbered; numbered items output as "1.")
  - Tables (simple pipe Markdown)
  - Line and tab breaks

Usage:
  python docx_to_markdown_txt.py input.docx [-o output.txt]
"""

import argparse
import os
import re
import sys

def try_convert_with_pandoc(input_path: str) -> str | None:
    """
    Try converting the docx to Markdown using Pandoc via pypandoc.
    Returns Markdown string on success, or None if pypandoc/pandoc is unavailable or fails.
    """
    try:
        import pypandoc  # type: ignore
    except Exception:
        return None
    try:
        # 'gfm' (GitHub Flavored Markdown) yields robust Markdown output.
        # You can change to 'markdown' if desired.
        md = pypandoc.convert_file(input_path, to="gfm", format="docx")
        return md
    except Exception:
        return None


def convert_with_python_docx(input_path: str) -> str:
    """
    Convert a DOCX file to Markdown using python-docx and basic formatting rules.
    """
    try:
        from docx import Document
        from docx.document import Document as _Document
        from docx.table import Table, _Cell
        from docx.text.paragraph import Paragraph
        from docx.oxml.ns import qn, nsmap
    except ImportError as e:
        raise SystemExit(
            "Missing dependency: python-docx. Install it with:\n  pip install python-docx"
        ) from e

    doc = Document(input_path)

    def md_escape(text: str, escape_pipes: bool = False) -> str:
        """
        Escape Markdown special characters in text.
        """
        if not text:
            return ""
        # Escape backslash first to avoid double-escaping.
        text = text.replace("\\", "\\\\")
        # Characters to escape in MD
        chars = r"`*_{}[]()#+-.!>"
        if escape_pipes:
            chars += r"|"
        return re.sub(r"([%s])" % re.escape(chars), r"\\\1", text)

    def is_true_on(node) -> bool:
        """
        A Word boolean element like w:b or w:i might have w:val='0' or 'false' for off.
        Presence with no w:val or w:val not in ('0','false') means true.
        """
        if node is None:
            return False
        val = node.get(qn("w:val"))
        if val is None:
            return True
        return str(val).lower() not in ("0", "false", "off")

    def extract_text_with_breaks(run_like_el) -> str:
        """
        Extract text from a w:r or w:hyperlink subtree, converting:
          - w:t -> text
          - w:tab -> 4 spaces
          - w:br/w:cr -> newline
        """
        parts: list[str] = []
        # Use .iter() to catch all nested w:t nodes
        for node in run_like_el.iter():
            tag = node.tag
            if tag == qn("w:t"):
                parts.append(node.text or "")
            elif tag == qn("w:tab"):
                parts.append("    ")
            elif tag in (qn("w:br"), qn("w:cr")):
                parts.append("\n")
        return "".join(parts)

    def get_hyperlink_url(hlink_el, paragraph) -> str | None:
        """
        Resolve hyperlink URL (external or anchor) from a w:hyperlink element.
        """
        rid = hlink_el.get(qn("r:id"))
        anchor = hlink_el.get(qn("w:anchor"))
        if rid:
            # python-docx relationship mapping
            try:
                rel = paragraph.part.rels[rid]
                url = getattr(rel, "target_ref", None) or getattr(rel, "_target", None)
                return url
            except Exception:
                return None
        if anchor:
            return f"#{anchor}"
        return None

    def run_element_to_markdown(run_el) -> str:
        """
        Convert a single w:r element to Markdown with bold/italic.
        """
        rPr = run_el.find(qn("w:rPr"))
        bold = is_true_on(rPr.find(qn("w:b"))) if rPr is not None else False
        italic = is_true_on(rPr.find(qn("w:i"))) if rPr is not None else False

        # Optional: detect a "code" character style (if used in the doc)
        code = False
        if rPr is not None:
            rstyle = rPr.find(qn("w:rStyle"))
            if rstyle is not None:
                style_val = (rstyle.get(qn("w:val")) or "").lower()
                if "code" in style_val:
                    code = True

        text = extract_text_with_breaks(run_el)
        if not text:
            return ""

        # Escape inline text for Markdown
        esc = md_escape(text)

        # Apply inline formatting
        if code:
            # Use backticks for inline code; if text contains backticks, use double backticks
            if "`" in esc:
                esc = "``" + esc.replace("``", "\\`\\`") + "``"
            else:
                esc = "`" + esc + "`"
        else:
            if bold and italic:
                esc = "***" + esc + "***"
            elif bold:
                esc = "**" + esc + "**"
            elif italic:
                esc = "*" + esc + "*"

        return esc

    def paragraph_inline_to_markdown(paragraph) -> str:
        """
        Convert inline content of a paragraph (runs and hyperlinks) into Markdown string.
        """
        p = paragraph._p  # CT_P element
        parts: list[str] = []
        for child in p.iterchildren():
            tag = child.tag
            if tag == qn("w:hyperlink"):
                text = extract_text_with_breaks(child)
                url = get_hyperlink_url(child, paragraph)
                esc_text = md_escape(text)
                if url:
                    parts.append(f"[{esc_text}]({url})")
                else:
                    parts.append(esc_text)
            elif tag == qn("w:r"):
                parts.append(run_element_to_markdown(child))
            elif tag == qn("w:fldSimple"):
                # Simple fields: render inner text
                parts.append(extract_text_with_breaks(child))
            else:
                # Other elements (bookmarks, smart tags, etc.) -> extract any visible text
                parts.append(extract_text_with_breaks(child))
        # Normalize spaces but preserve newlines
        joined = "".join(parts)
        # Replace multiple spaces with single, but avoid touching newlines
        joined = re.sub(r"[ \t]+", " ", joined)
        # Collapse spaces around newlines
        joined = re.sub(r" *\n *", "\n", joined)
        return joined.strip()

    def get_heading_level(paragraph) -> int | None:
        """
        Detect heading level from paragraph style name (Heading 1..6).
        """
        try:
            style_name = (paragraph.style.name or "").strip()
        except Exception:
            style_name = ""
        m = re.match(r"Heading\s+([1-6])", style_name, flags=re.I)
        if m:
            return int(m.group(1))
        # Optional: treat "Title" as H1
        if style_name.lower() == "title":
            return 1
        return None

    def get_list_info(paragraph):
        """
        Determine if a paragraph is part of a list and whether it is ordered or unordered.
        Returns tuple (is_list, is_ordered, level).
        """
        p = paragraph._p
        pPr = p.pPr
        if pPr is None or pPr.numPr is None:
            return (False, False, 0)

        ilvl = 0
        if pPr.numPr.ilvl is not None:
            try:
                ilvl = int(pPr.numPr.ilvl.val)
            except Exception:
                ilvl = 0

        is_ordered = None
        try:
            numId = pPr.numPr.numId.val
            numbering_el = paragraph.part.numbering_part.element  # CT_Numbering
            # Find abstractNumId for numId
            nums = numbering_el.xpath(f'./w:num[@w:numId="{numId}"]/w:abstractNumId', namespaces=nsmap)
            if nums:
                abstract_id = nums[0].get(qn("w:val"))
                if abstract_id is None:
                    abstract_id = nums[0].get("{%s}val" % nsmap["w"])
                # Find numFmt for this level
                fmt_nodes = numbering_el.xpath(
                    f'./w:abstractNum[@w:abstractNumId="{abstract_id}"]/w:lvl[@w:ilvl="{ilvl}"]/w:numFmt',
                    namespaces=nsmap,
                )
                if fmt_nodes:
                    fmt_val = fmt_nodes[0].get(qn("w:val")) or fmt_nodes[0].get("{%s}val" % nsmap["w"])
                    # Treat 'bullet' and 'none' as unordered; anything else as ordered
                    is_ordered = fmt_val not in ("bullet", "none")
        except Exception:
            pass

        if is_ordered is None:
            is_ordered = False

        return (True, is_ordered, ilvl)

    def iter_block_items(parent):
        """
        Iterate over document contents in order, yielding Paragraph and Table objects.
        Supports nesting in table cells.
        """
        from docx.document import Document as _Document
        from docx.table import _Cell, Table
        from docx.text.paragraph import Paragraph

        if isinstance(parent, _Document):
            parent_elm = parent.element.body
        elif isinstance(parent, _Cell):
            parent_elm = parent._tc
        else:
            raise ValueError("Unknown parent type")

        for child in parent_elm.iterchildren():
            if child.tag == qn("w:p"):
                yield Paragraph(child, parent)
            elif child.tag == qn("w:tbl"):
                yield Table(child, parent)

    def table_to_markdown(table: "Table") -> str:
        """
        Convert a table to a simple Markdown pipe table.
        First row treated as header.
        """
        # Collect rows -> list of list of cell texts
        rows: list[list[str]] = []
        for row in table.rows:
            cells_text: list[str] = []
            # Use unique cells to avoid duplicates from merged cells referencing same object
            seen_tc_addrs = set()
            for cell in row.cells:
                tc_addr = id(cell._tc)
                if tc_addr in seen_tc_addrs:
                    # Skip repeated reference to the same merged cell
                    continue
                seen_tc_addrs.add(tc_addr)
                # Combine cell paragraphs inline; replace newlines with <br> to keep table valid MD
                txt_parts = []
                for p in cell.paragraphs:
                    inline = paragraph_inline_to_markdown(p)
                    if inline:
                        txt_parts.append(inline)
                cell_text = " ".join(txt_parts).strip()
                cell_text = cell_text.replace("\n", "<br>")
                # Escape pipes to avoid breaking table
                cell_text = md_escape(cell_text, escape_pipes=True)
                cells_text.append(cell_text)
            rows.append(cells_text)

        if not rows:
            return ""

        # Normalize column count to the max width across rows
        max_cols = max((len(r) for r in rows), default=0)
        for r in rows:
            while len(r) < max_cols:
                r.append("")

        # Build Markdown lines
        lines: list[str] = []
        header = rows[0]
        sep = ["---"] * max_cols
        lines.append("| " + " | ".join(header) + " |")
        lines.append("| " + " | ".join(sep) + " |")
        for r in rows[1:]:
            lines.append("| " + " | ".join(r) + " |")

        return "\n".join(lines)

    out_lines: list[str] = []
    prev_was_list = False

    for block in iter_block_items(doc):
        if block.__class__.__name__ == "Table":
            # Ensure blank line before table (except at start)
            if out_lines and out_lines[-1] != "":
                out_lines.append("")
            out_lines.append(table_to_markdown(block))
            out_lines.append("")  # blank line after table
            prev_was_list = False
            continue

        # Paragraph
        para = block
        heading_level = get_heading_level(para)
        inline = paragraph_inline_to_markdown(para)
        if not inline and not heading_level:
            # preserve blank line between blocks
            if out_lines and out_lines[-1] != "":
                out_lines.append("")
            prev_was_list = False
            continue

        if heading_level:
            # Ensure blank line before heading (except at start)
            if out_lines and out_lines[-1] != "":
                out_lines.append("")
            out_lines.append("#" * heading_level + " " + inline.strip())
            out_lines.append("")  # blank after heading
            prev_was_list = False
            continue

        is_list, is_ordered, level = get_list_info(para)
        if is_list:
            indent = "  " * level
            bullet = "1. " if is_ordered else "- "
            # For ordered lists, '1.' for every item lets Markdown auto-number them.
            out_lines.append(f"{indent}{bullet}{inline.strip()}")
            prev_was_list = True
        else:
            # Non-list paragraph
            if prev_was_list and out_lines and out_lines[-1] != "":
                # Blank line to terminate a preceding list
                out_lines.append("")
            out_lines.append(inline.strip())
            out_lines.append("")  # Blank line after paragraph
            prev_was_list = False

    # Clean trailing whitespace/newlines
    md = "\n".join(out_lines).rstrip() + "\n"
    return md


def convert_docx_to_markdown(input_path: str) -> str:
    """
    Convert DOCX to Markdown, preferring Pandoc if available for best fidelity.
    """
    md = try_convert_with_pandoc(input_path)
    if md is not None and md.strip():
        return md
    return convert_with_python_docx(input_path)


def main(inputPath):
    parser = argparse.ArgumentParser(description="Extract text from a Word (.docx) document and save to a .txt file in Markdown format.")
    parser.add_argument("--encoding", default="utf-8", help="Text encoding for output file (default: utf-8)")
    args = parser.parse_args()

    input_path = inputPath
    if not os.path.isfile(input_path):
        print(f"Error: File not found: {input_path}", file=sys.stderr)
        sys.exit(1)

    _, ext = os.path.splitext(input_path)
    if ext.lower() != ".docx":
        print("Error: Only .docx files are supported.", file=sys.stderr)
        sys.exit(1)


    try:
        md = convert_docx_to_markdown(input_path)
    except Exception as e:
        print(f"Conversion failed: {e}", file=sys.stderr)
        sys.exit(1)

    return md