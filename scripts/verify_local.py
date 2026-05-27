#!/usr/bin/env python3
"""verify_local.py - Local database citation verification.

Step 1 of the two-step citation verification pipeline.
Validates citations against a local bibliography file (.bib, .ris, .xml).

Supported formats:
  - .bib (BibLaTeX/BibTeX) — Zotero, EndNote, Mendeley export
  - .ris (Research Information Systems) — EndNote, NoteExpress export
  - .xml (Zotero RDF/XML) — Zotero export

Usage:
    python verify_local.py thesis.docx --bib refs.bib --output local_result.json
    python verify_local.py thesis.docx --ris refs.ris --output local_result.json
"""

from __future__ import annotations

import argparse
import json
import re
import sys
from difflib import SequenceMatcher
from pathlib import Path
from typing import Any

_THIS = Path(__file__).resolve()
_REPO_ROOT = _THIS.parents[1]
if str(_REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(_REPO_ROOT))

from scripts.extract_references import detect_references_section
from scripts.citation_repair import parse_raw_citation
from scripts.lib.citation_models import CitationRecord, ValidationResult, citations_from_raw


# -------------------------------------------------------------------
# BibLaTeX / BibTeX parser
# -------------------------------------------------------------------

def parse_bib(bib_path: str | Path) -> list[dict[str, Any]]:
    """Parse a .bib file and return list of citation records."""
    bib_path = Path(bib_path)
    if not bib_path.exists():
        return []

    text = bib_path.read_text(encoding="utf-8", errors="replace")
    entries: list[dict[str, Any]] = []

    # Match @type{key, fields...}
    entry_re = re.compile(
        r'@(\w+)\s*\{([^,]+),\s*(.*?)\n\}',
        re.DOTALL
    )

    for match in entry_re.finditer(text):
        entry_type = match.group(1).lower()
        entry_key = match.group(2).strip()
        fields_text = match.group(3)

        entry: dict[str, Any] = {
            "key": entry_key,
            "type": entry_type,
        }

        # Parse fields: field = {value} or field = value
        field_re = re.compile(r'(\w+)\s*=\s*[{"](.+?)[}"]', re.DOTALL)
        for fm in field_re.finditer(fields_text):
            fname = fm.group(1).lower()
            fval = fm.group(2).strip()
            # Clean up BibTeX formatting
            fval = re.sub(r'\s+', ' ', fval)
            fval = fval.strip('{}')
            entry[fname] = fval

        # Normalize
        entry["title_norm"] = _normalize_title(entry.get("title", ""))
        entry["authors_list"] = _parse_bib_authors(entry.get("author", ""))
        entry["year_int"] = _parse_year(entry.get("year", ""))

        if entry["title_norm"]:
            entries.append(entry)

    return entries


def _parse_bib_authors(author_str: str) -> list[str]:
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
    """Extract 4-digit year."""
    m = re.search(r'((?:19|20)\d{2})', year_str)
    return int(m.group(1)) if m else 0


# -------------------------------------------------------------------
# RIS parser
# -------------------------------------------------------------------

def parse_ris(ris_path: str | Path) -> list[dict[str, Any]]:
    """Parse a .ris file and return list of citation records."""
    ris_path = Path(ris_path)
    if not ris_path.exists():
        return []

    text = ris_path.read_text(encoding="utf-8", errors="replace")
    entries: list[dict[str, Any]] = []
    current: dict[str, Any] = {}

    for line in text.splitlines():
        line = line.rstrip()
        if not line:
            continue

        tag_match = re.match(r'^([A-Z][A-Z0-9])\s*-\s*(.*)', line)
        if not tag_match:
            continue

        tag = tag_match.group(1)
        value = tag_match.group(2).strip()

        if tag == "TY":
            current = {"type": value, "authors_list": []}
        elif tag == "ER":
            if current:
                current["title_norm"] = _normalize_title(current.get("title", ""))
                current["year_int"] = _parse_year(current.get("year", ""))
                if current["title_norm"]:
                    entries.append(current)
            current = {}
        elif tag == "TI" or tag == "T1":
            current["title"] = value
        elif tag == "AU":
            current.setdefault("authors_list", []).append(value)
        elif tag == "PY" or tag == "Y1":
            current["year"] = value
        elif tag == "JO" or tag == "JA" or tag == "T2":
            current["venue"] = value
        elif tag == "DO":
            current["doi"] = value
        elif tag == "SN":
            current["isbn"] = value
        elif tag == "UR":
            current["url"] = value

    return entries


# -------------------------------------------------------------------
# XML (Zotero RDF) parser (simplified)
# -------------------------------------------------------------------

def parse_xml(xml_path: str | Path) -> list[dict[str, Any]]:
    """Parse a Zotero XML export and return list of citation records."""
    try:
        from lxml import etree
    except ImportError:
        import xml.etree.ElementTree as etree

    xml_path = Path(xml_path)
    if not xml_path.exists():
        return []

    try:
        tree = etree.parse(str(xml_path))
        root = tree.getroot()
    except Exception:
        return []

    entries: list[dict[str, Any]] = []

    # Zotero RDF uses various namespaces — try common patterns
    for elem in root.iter():
        tag = etree.QName(elem.tag).localname if hasattr(elem.tag, 'startswith') else ""
        if tag not in ("item", "entry", "record"):
            continue

        entry: dict[str, Any] = {"authors_list": []}
        for child in elem:
            ctag = etree.QName(child.tag).localname if hasattr(child.tag, 'startswith') else ""
            ctext = (child.text or "").strip()
            if ctag in ("title",):
                entry["title"] = ctext
            elif ctag in ("creator", "author"):
                if ctext:
                    entry["authors_list"].append(ctext)
            elif ctag in ("date", "year"):
                entry["year"] = ctext
            elif ctag in ("publisher", "publicationTitle", "journal"):
                entry["venue"] = ctext
            elif ctag in ("DOI", "doi"):
                entry["doi"] = ctext

        entry["title_norm"] = _normalize_title(entry.get("title", ""))
        entry["year_int"] = _parse_year(entry.get("year", ""))
        if entry["title_norm"]:
            entries.append(entry)

    return entries


# -------------------------------------------------------------------
# Matching logic
# -------------------------------------------------------------------

def _normalize_title(title: str) -> str:
    """Normalize a title for fuzzy matching."""
    if not title:
        return ""
    # Remove common LaTeX/BibTeX formatting
    title = re.sub(r'[{}\\]', '', title)
    # Remove punctuation and normalize whitespace
    title = re.sub(r'[^\w\s]', '', title, flags=re.UNICODE)
    title = re.sub(r'\s+', ' ', title).strip().lower()
    return title


def _title_similarity(t1: str, t2: str) -> float:
    """Compute similarity between two normalized titles."""
    if not t1 or not t2:
        return 0.0
    return SequenceMatcher(None, t1, t2).ratio()


def verify_against_local(
    citations: list[CitationRecord],
    local_db: list[dict[str, Any]],
    db_source: str = "bib",
    similarity_threshold: float = 0.80,
) -> list[ValidationResult]:
    """Verify citations against a local database.

    Args:
        citations: List of CitationRecord from thesis
        local_db: List of parsed entries from .bib/.ris/.xml
        db_source: Source format label
        similarity_threshold: Minimum title similarity for match

    Returns:
        List of ValidationResult
    """
    # Build title index for fast lookup
    title_index: dict[str, list[dict]] = {}
    for entry in local_db:
        norm = entry.get("title_norm", "")
        if norm:
            title_index.setdefault(norm[:20], []).append(entry)

    results: list[ValidationResult] = []

    for cit in citations:
        vr = ValidationResult(
            citation_index=cit.index,
            raw_text=cit.raw_text,
        )

        cit_title_norm = _normalize_title(cit.title)
        if not cit_title_norm:
            vr.issues.append("无法从引用中提取标题，跳过本地验证")
            results.append(vr)
            continue

        # Search for match
        best_match = None
        best_score = 0.0

        # Exact title match first
        prefix = cit_title_norm[:20]
        candidates = title_index.get(prefix, [])
        # Also check all entries (title might differ in prefix)
        all_entries = local_db if len(local_db) < 500 else candidates

        for entry in all_entries:
            score = _title_similarity(cit_title_norm, entry.get("title_norm", ""))
            if score > best_score:
                best_score = score
                best_match = entry

        if best_match and best_score >= similarity_threshold:
            vr.local_match = True
            vr.local_source = db_source
            vr.local_match_title = best_match.get("title", "")

            # Verify fields
            # Year
            if cit.year and best_match.get("year_int"):
                if cit.year == best_match["year_int"]:
                    vr.verified_fields["year"] = "match"
                else:
                    vr.verified_fields["year"] = f"mismatch: doc={cit.year}, db={best_match['year_int']}"
                    vr.issues.append(f"年份不一致：论文={cit.year}，数据库={best_match['year_int']}")

            # Authors (first author check)
            if cit.authors and best_match.get("authors_list"):
                cit_first = cit.authors[0].lower().strip()
                db_first = best_match["authors_list"][0].lower().strip()
                if cit_first in db_first or db_first in cit_first:
                    vr.verified_fields["author"] = "match"
                else:
                    vr.verified_fields["author"] = f"possible_mismatch: doc='{cit.authors[0]}', db='{best_match['authors_list'][0]}'"
                    vr.issues.append(f"作者可能不一致：论文={cit.authors[0]}，数据库={best_match['authors_list'][0]}")

            # DOI
            if cit.doi and best_match.get("doi"):
                if cit.doi.lower() == best_match["doi"].lower():
                    vr.verified_fields["doi"] = "match"
                else:
                    vr.verified_fields["doi"] = "mismatch"
        else:
            vr.local_match = False
            vr.issues.append(f"本地数据库未找到匹配 (best score={best_score:.2f})")

        results.append(vr)

    return results


# -------------------------------------------------------------------
# Main
# -------------------------------------------------------------------

def main():
    from scripts.lib import setup_utf8_output
    setup_utf8_output()
    parser = argparse.ArgumentParser(description="Local citation verification")
    parser.add_argument("thesis", help="Path to thesis .docx")
    parser.add_argument("--bib", help="Path to .bib file")
    parser.add_argument("--ris", help="Path to .ris file")
    parser.add_argument("--xml", help="Path to Zotero XML file")
    parser.add_argument("--output", "-o", help="Output JSON path")
    parser.add_argument("--threshold", type=float, default=0.80, help="Similarity threshold (0-1)")
    args = parser.parse_args()

    # Load local database
    local_db: list[dict[str, Any]] = []
    db_source = ""

    if args.bib:
        local_db = parse_bib(args.bib)
        db_source = "bib"
        print(f"[verify_local] Loaded {len(local_db)} entries from {args.bib}")
    elif args.ris:
        local_db = parse_ris(args.ris)
        db_source = "ris"
        print(f"[verify_local] Loaded {len(local_db)} entries from {args.ris}")
    elif args.xml:
        local_db = parse_xml(args.xml)
        db_source = "xml"
        print(f"[verify_local] Loaded {len(local_db)} entries from {args.xml}")
    else:
        print("Error: 请提供 --bib, --ris, 或 --xml 参数")
        sys.exit(1)

    if not local_db:
        print("Warning: 本地数据库为空")
        sys.exit(0)

    # Extract citations from thesis
    ref_result = detect_references_section(args.thesis)
    if not ref_result.found:
        print("No references section found in thesis.")
        sys.exit(0)

    citations = citations_from_raw(ref_result.raw_citations)
    print(f"[verify_local] Found {len(citations)} citations in thesis")

    # Verify
    results = verify_against_local(citations, local_db, db_source, args.threshold)

    matched = sum(1 for r in results if r.local_match)
    print(f"[verify_local] Matched: {matched}/{len(results)}")

    output = {
        "source_file": args.thesis,
        "db_file": getattr(args, db_source, ""),
        "db_source": db_source,
        "db_entries": len(local_db),
        "citations_total": len(citations),
        "matched": matched,
        "unmatched": len(results) - matched,
        "results": [r.to_dict() for r in results],
    }

    if args.output:
        Path(args.output).write_text(json.dumps(output, ensure_ascii=False, indent=2), encoding="utf-8")
        print(f"Results written to {args.output}")
    else:
        print(json.dumps(output, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
