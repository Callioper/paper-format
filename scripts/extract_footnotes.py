#!/usr/bin/env python3
"""extract_footnotes.py - Extract and classify footnotes from a .docx file."""

from __future__ import annotations
import re
import zipfile
from dataclasses import dataclass, field
from pathlib import Path

_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
_NS = {"w": _W}

# Patterns that introduce a citation after prose
_RE_SEE = re.compile(r"(?:参见|详见)\s*(.+)|[Ss]ee\s+(.+)")

# Subsequent citation heuristic: author + title + page, no publication data
_RE_SUBSEQ_ZH = re.compile(
    r"^.+[：:]《.+》[，,]\s*第\s*\d+[—\-]\d+\s*页|^.+[：:]《.+》[，,]\s*第\s*\d+\s*页"
)


@dataclass
class FootnoteItem:
    footnote_id: int
    raw_text: str       # full text as extracted (whitespace-normalized)
    cite_text: str      # citation portion only (strips prose prefix)
    is_subsequent: bool = False  # re-citation in shortened form
    is_prose_only: bool = False  # no citation content at all


def _join_runs(fn_element) -> str:
    """Join all w:t text nodes, then normalize whitespace."""
    parts = [t.text for t in fn_element.findall(".//w:t", _NS) if t.text]
    raw = "".join(parts)
    # Collapse runs of whitespace to single space
    raw = re.sub(r"\s+", " ", raw).strip()
    # Fix spaces leaking inside curly/straight quote pairs around a title
    # e.g.  " Against Theory , "  →  "Against Theory,"
    raw = re.sub(r'["“]\s+', "“", raw)   # open quote
    raw = re.sub(r'\s+["”]', "”", raw)   # close quote
    # Normalise space after p. / pp.  (p.728 → p. 728)
    raw = re.sub(r"\b(pp?)\.\s*(\d)", r"\1. \2", raw)
    return raw


def _extract_cite_text(text: str) -> str:
    """Strip prose prefix (quoted block + 参见/详见/see) and return citation part."""
    m = _RE_SEE.search(text)
    if m:
        return (m.group(1) or m.group(2) or "").strip()
    return text


def _is_subsequent(text: str) -> bool:
    """
    Detect shortened re-citation form:
      Chinese: Author：《Title》，第X页。  (no publisher/place/year)
      English: Author, "Title," p. X.    (no journal/publisher/year)
    """
    # Must have a title marker
    has_title = bool(re.search(r"《.+》", text) or
                     re.search(r'["“].+["”]', text) or
                     re.search(r"\*.+\*", text))
    if not has_title:
        return False

    # Must have a page reference
    has_pages = bool(re.search(r"第\s*\d+|pp?\.\s*\d+", text, re.IGNORECASE))
    if not has_pages:
        return False

    # Must NOT have year or publisher indicators
    has_year = bool(re.search(r"(?<!\d)((?:19|20)\d{2})(?!\d)", text))
    has_pub = bool(re.search(r"出版社|出版|Press|Publisher", text, re.IGNORECASE))
    if has_year or has_pub:
        return False

    return True


def _is_prose_only(text: str) -> bool:
    """True when footnote has no recognisable citation content."""
    has_title = bool(re.search(r"《.+》", text) or
                     re.search(r'["“].+["”]', text) or
                     re.search(r"\*.+\*", text))
    has_author_sep = bool(re.search(r"[：:,，]", text))
    return not has_title and not has_author_sep


def extract_footnotes(docx_path: str | Path) -> list[FootnoteItem]:
    """
    Extract all non-separator footnotes from a .docx file.

    Returns a list of FootnoteItem, ordered by footnote_id.
    """
    docx_path = Path(docx_path)
    items: list[FootnoteItem] = []

    with zipfile.ZipFile(docx_path) as z:
        if "word/footnotes.xml" not in z.namelist():
            return []
        xml = z.read("word/footnotes.xml")

    from lxml import etree
    root = etree.fromstring(xml)

    for fn in root.findall("w:footnote", _NS):
        fid = int(fn.get(f"{{{_W}}}id", -999))
        fn_type = fn.get(f"{{{_W}}}type", "")
        if fn_type in ("separator", "continuationSeparator") or fid < 1:
            continue

        raw = _join_runs(fn)
        if not raw:
            continue

        cite = _extract_cite_text(raw)
        subsequent = _is_subsequent(cite)
        prose_only = _is_prose_only(cite)

        items.append(FootnoteItem(
            footnote_id=fid,
            raw_text=raw,
            cite_text=cite,
            is_subsequent=subsequent,
            is_prose_only=prose_only,
        ))

    items.sort(key=lambda x: x.footnote_id)
    return items
