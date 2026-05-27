#!/usr/bin/env python3
"""extract_references.py - Extract references section from .docx

Detects the references/bibliography section in a thesis .docx and extracts
individual citation paragraphs for downstream processing by citation_repair.py.
"""

from __future__ import annotations
import re
from pathlib import Path
from typing import NamedTuple


REFERENCES_KEYWORDS = [
    "参考文献",
    "References",
    "Works Cited",
    "引用文献",
    "引用书目",
    "Bibliography",
    "文献目录",
    "参考书目",
]

# Stop words: section headers that indicate end of references section
STOP_HEADING_PREFIXES = (
    "Heading",
    "TOC",
    "目录",
    "附录",
    "Appendix",
    "致谢",
    "Acknowledgments",
    "作者简介",
)


class ReferencesResult(NamedTuple):
    found: bool
    section_start: int | None
    section_end: int | None
    raw_citations: list[str]
    citation_count: int


def detect_references_section(docx_path: str | Path) -> ReferencesResult:
    """
    Scan a .docx file for a references/bibliography section.

    Returns a ReferencesResult with:
      - found: bool indicating whether a references section was detected
      - section_start: paragraph index where references begin (None if not found)
      - section_end: paragraph index where references end (None if not found)
      - raw_citations: list of raw citation paragraph texts
      - citation_count: number of citations found
    """
    try:
        from docx import Document
    except ImportError:
        raise RuntimeError(
            "python-docx is required. Install with: pip install python-docx"
        )

    doc = Document(docx_path)
    paragraphs = doc.paragraphs

    in_references = False
    section_start: int | None = None
    section_end: int | None = None
    raw_citations: list[str] = []

    for idx, para in enumerate(paragraphs):
        text = para.text.strip()

        # Detect section header
        if not in_references:
            if _is_references_header(para):
                in_references = True
                section_start = idx
            continue

        # End conditions for references section
        if _is_stop_heading(para):
            section_end = idx
            break

        # Empty paragraphs between citations — skip
        if not text:
            continue

        # Skip numbered list items that are just numbers/bullets
        if _is_list_item_stub(text):
            continue

        raw_citations.append(text)

    found = in_references and len(raw_citations) > 0
    return ReferencesResult(
        found=found,
        section_start=section_start,
        section_end=section_end,
        raw_citations=raw_citations,
        citation_count=len(raw_citations),
    )


def _has_cjk(text: str) -> bool:
    """Return True if text contains CJK characters (U+4E00-U+9FFF)."""
    return any('一' <= c <= '鿿' for c in text)


def _is_references_header(para) -> bool:
    """Return True if paragraph is a references/bibliography section heading.

    Requires either a heading style OR a short standalone keyword (≤20 chars),
    to avoid matching keywords that appear mid-sentence in body text.
    """
    text = para.text.strip() if hasattr(para, "text") else str(para)
    if not text:
        return False
    style_name = (para.style.name if para.style else "") if hasattr(para, "style") else ""
    text_lower = text.lower()
    kw_match = any(kw.lower() in text_lower for kw in REFERENCES_KEYWORDS)

    if not kw_match:
        return False

    # Heading-style paragraph — any keyword match is valid
    if style_name.startswith(("Heading", "heading", "标题")):
        return True

    # Non-heading: only accept if the paragraph is essentially just the keyword
    # (≤20 chars covers "参考文献", "References", "Works Cited", etc.)
    if len(text) <= 20:
        return True

    return False


def _is_stop_heading(para) -> bool:
    """Return True if paragraph style indicates end of references section."""
    style_name = para.style.name if para.style else ""
    if style_name.startswith("Heading"):
        return True
    text = para.text.strip()
    # Use Unicode range check to handle encoding issues
    if _has_cjk(text):
        for kw in ["附录", "致谢", "作者简介"]:
            if kw in text:
                return True
    return False


def _is_list_item_stub(text: str) -> bool:
    """Return True if text is just a bullet/number with no real content."""
    return bool(re.match(r"^[\[\]•\d\.\)\-]+$", text))


def extract_citations(docx_path: str | Path) -> list[str]:
    """
    Convenience function — returns just the raw citation list.
    Use detect_references_section() for full diagnostic info.
    """
    result = detect_references_section(docx_path)
    return result.raw_citations


if __name__ == "__main__":
    import sys

    if len(sys.argv) < 2:
        print("Usage: python extract_references.py <thesis.docx>")
        sys.exit(1)

    result = detect_references_section(sys.argv[1])
    print(f"References section found: {result.found}")
    print(f"Citation count: {result.citation_count}")
    if result.raw_citations:
        print("\nCitations:")
        for i, c in enumerate(result.raw_citations, 1):
            print(f"  [{i}] {c[:80]}{'...' if len(c) > 80 else ''}")