"""Citation data models for paper-check verification pipeline."""

from __future__ import annotations

from dataclasses import dataclass, field
import json


@dataclass
class CitationRecord:
    """A single citation extracted from a thesis document."""
    index: int = 0
    raw_text: str = ""
    authors: list[str] = field(default_factory=list)
    title: str = ""
    year: int = 0
    venue: str = ""
    doi: str = ""
    doc_type: str = ""
    language: str = "zh"
    pages: str = ""

    def to_dict(self) -> dict:
        return {
            "index": self.index,
            "raw_text": self.raw_text[:120],
            "authors": self.authors,
            "title": self.title,
            "year": self.year,
            "venue": self.venue,
            "doi": self.doi,
            "doc_type": self.doc_type,
            "language": self.language,
            "pages": self.pages,
        }


@dataclass
class ValidationResult:
    """Verification result for a single citation."""
    citation_index: int = 0
    raw_text: str = ""
    local_match: bool = False
    local_source: str = ""
    local_match_title: str = ""
    external_match: bool = False
    external_source: str = ""
    external_match_title: str = ""
    verified_fields: dict = field(default_factory=dict)
    issues: list[str] = field(default_factory=list)

    def to_dict(self) -> dict:
        return {
            "citation_index": self.citation_index,
            "raw_text": self.raw_text[:120],
            "local_match": self.local_match,
            "local_source": self.local_source,
            "external_match": self.external_match,
            "external_source": self.external_source,
            "verified_fields": self.verified_fields,
            "issues": self.issues,
        }


@dataclass
class CrossCheckResult:
    """Cross-check between in-text citations and reference list."""
    orphan_citations: list[dict] = field(default_factory=list)
    unused_references: list[dict] = field(default_factory=list)
    duplicates: list[dict] = field(default_factory=list)
    incomplete: list[dict] = field(default_factory=list)

    def to_dict(self) -> dict:
        return {
            "orphan_citations": self.orphan_citations,
            "unused_references": self.unused_references,
            "duplicates": self.duplicates,
            "incomplete": self.incomplete,
        }


def citations_from_raw(raw_citations: list[str]) -> list[CitationRecord]:
    """Convert raw citation texts to CitationRecord objects using parse_raw_citation."""
    import re
    from scripts.citation_repair import parse_raw_citation

    records: list[CitationRecord] = []
    for i, raw in enumerate(raw_citations, 1):
        parsed = parse_raw_citation(raw)
        rec = CitationRecord(
            index=i,
            raw_text=raw,
            authors=[parsed.get("author", "")] if parsed.get("author") else [],
            title=parsed.get("title", ""),
            year=int(parsed["year"]) if parsed.get("year") and parsed["year"].isdigit() else 0,
            venue=parsed.get("journal", parsed.get("publisher", "")),
            doi=parsed.get("doi", ""),
            doc_type=parsed.get("type", ""),
            language=parsed.get("language", "zh"),
            pages=parsed.get("pages", ""),
        )
        records.append(rec)
    return records
