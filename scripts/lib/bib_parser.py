"""bib_parser.py - Better BibTeX .bib file parsing and lookup.

Shared by citation_repair.py (metadata enrichment) and verify_local.py
(local database verification). Single source of truth for .bib parsing.

Supports:
  - Standard BibTeX/BibLaTeX entries
  - Better BibTeX file field (PDF path extraction)
  - Title-based fuzzy matching with prefix index
"""

from __future__ import annotations

import re
from difflib import SequenceMatcher
from pathlib import Path
from typing import Any


def parse_bib_file(bib_path: str | Path) -> list[dict[str, Any]]:
    """Load and parse a .bib file into a list of entry dicts.

    Each entry contains:
      key, type, title, author, year, publisher, address, pages, doi,
      journal, volume, number, booktitle, file,
      title_norm, authors_list, year_int, pdf_path
    """
    bib_path = Path(bib_path)
    if not bib_path.exists():
        return []

    text = bib_path.read_text(encoding="utf-8", errors="replace")
    entries: list[dict[str, Any]] = []

    # Match @type{key, fields...} — handles nested braces via non-greedy + newline stop
    entry_re = re.compile(r'@(\w+)\s*\{([^,]+),\s*(.*?)\n\}', re.DOTALL)
    field_re = re.compile(r'(\w+)\s*=\s*[{\"](.+?)[}\"]', re.DOTALL)

    for match in entry_re.finditer(text):
        entry_type = match.group(1).lower()
        entry_key = match.group(2).strip()
        fields_text = match.group(3)

        entry: dict[str, Any] = {"key": entry_key, "type": entry_type}

        for fm in field_re.finditer(fields_text):
            fname = fm.group(1).lower()
            fval = fm.group(2).strip()
            fval = re.sub(r'\s+', ' ', fval).strip('{}')
            entry[fname] = fval

        # Normalize for matching
        title = entry.get("title", "")
        entry["title_norm"] = _normalize_title(title)

        author_str = entry.get("author", "")
        entry["authors_list"] = _parse_authors(author_str)

        year_str = entry.get("year", "")
        entry["year_int"] = _parse_year(year_str)

        # Parse Better BibTeX file field for PDF path
        entry["pdf_path"] = extract_pdf_path(entry.get("file", ""))

        if entry["title_norm"]:
            entries.append(entry)

    return entries


def extract_pdf_path(file_field: str) -> str:
    """Extract PDF path from Better BibTeX file field.

    Format: :/absolute/path.pdf:application/pdf  or  :C:/path/file.pdf:application/pdf
    Multiple files separated by ';'.
    """
    if not file_field:
        return ""
    for segment in file_field.split(";"):
        segment = segment.strip()
        parts = segment.split(":")
        # Better BibTeX format: :path:type (leading colon means path starts at index 1)
        if len(parts) >= 3 and parts[0] == "":
            path = parts[1].strip()
            mime = parts[2].strip() if len(parts) > 2 else ""
            if "pdf" in mime.lower() or path.lower().endswith(".pdf"):
                return path
        elif len(parts) >= 2:
            path = parts[0].strip()
            if path.lower().endswith(".pdf"):
                return path
    return ""


def build_title_index(entries: list[dict]) -> dict[str, list[dict]]:
    """Build a title-prefix index for fast lookup.

    Keys are the first 20 characters of normalized titles.
    """
    index: dict[str, list[dict]] = {}
    for entry in entries:
        norm = entry.get("title_norm", "")
        if norm:
            prefix = norm[:20]
            index.setdefault(prefix, []).append(entry)
    return index


def search_by_title(
    query_title: str,
    entries: list[dict],
    title_index: dict[str, list[dict]],
    threshold: float = 0.80,
) -> list[dict]:
    """Search .bib entries by title similarity.

    Uses prefix index for fast candidate selection, then fuzzy matches.
    Returns entries sorted by similarity (best first).
    """
    title_norm = _normalize_title(query_title)
    if not title_norm:
        return []

    # Try prefix index first
    prefix = title_norm[:20]
    candidates = title_index.get(prefix, [])
    if not candidates and len(entries) < 500:
        candidates = entries  # fallback to full scan for small libraries

    matches: list[tuple[float, dict]] = []
    for entry in candidates:
        entry_title = entry.get("title_norm", "")
        score = SequenceMatcher(None, title_norm, entry_title).ratio()
        if score >= threshold:
            matches.append((score, entry))

    matches.sort(key=lambda x: x[0], reverse=True)
    return [m[1] for m in matches]


def merge_bib_into_parsed(parsed: dict, bib_entry: dict) -> dict:
    """Merge .bib metadata into a parsed citation dict.

    Only fills in fields that are empty in the parsed version.
    Always carries over pdf_path if available.
    """
    merged = dict(parsed)
    field_map = {
        "author": "author",
        "title": "title",
        "year": "year",
        "publisher": "publisher",
        "place": "address",
        "pages": "pages",
        "doi": "doi",
        "journal": "journal",
        "volume": "volume",
        "number": "number",
        "booktitle": "booktitle",
    }
    for parsed_key, bib_key in field_map.items():
        bib_val = bib_entry.get(bib_key, "")
        if bib_val and not merged.get(parsed_key):
            merged[parsed_key] = bib_val

    if bib_entry.get("pdf_path"):
        merged["pdf_path"] = bib_entry["pdf_path"]

    merged["bib_merge"] = True
    return merged


# -------------------------------------------------------------------
# Shared helpers
# -------------------------------------------------------------------

def _normalize_title(title: str) -> str:
    """Normalize a title for fuzzy matching."""
    if not title:
        return ""
    title = re.sub(r'[{}\\]', '', title)
    title = re.sub(r'[^\w\s]', '', title, flags=re.UNICODE)
    title = re.sub(r'\s+', ' ', title).strip().lower()
    return title


def _parse_authors(author_str: str) -> list[str]:
    """Parse BibTeX author string (Last, First and Last, First)."""
    if not author_str:
        return []
    authors = []
    for part in author_str.split(" and "):
        part = part.strip()
        if "," in part:
            last, first = part.split(",", 1)
            authors.append(f"{first.strip()} {last.strip()}")
        else:
            authors.append(part)
    return authors


def _parse_year(year_str: str) -> int:
    """Extract 4-digit year from a string."""
    m = re.search(r'((?:19|20)\d{2})', year_str)
    return int(m.group(1)) if m else 0
