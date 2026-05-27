#!/usr/bin/env python3
"""verify_external.py - External API citation verification.

Step 2 of the two-step citation verification pipeline.
For citations not matched locally, queries external academic databases.

Supported sources:
  Free (no configuration):  OpenAlex, CrossRef, Semantic Scholar, CORE
  Experimental (config):    CNKI (需Cookie), Wanfang (公开搜索)

Usage:
    python verify_external.py thesis.docx --sources openalex,crossref --output ext_result.json
    python verify_external.py thesis.docx --sources all --output ext_result.json
"""

from __future__ import annotations

import argparse
import json
import sys
from concurrent.futures import ThreadPoolExecutor, as_completed
from difflib import SequenceMatcher
from pathlib import Path
from typing import Any

_THIS = Path(__file__).resolve()
_REPO_ROOT = _THIS.parents[1]
if str(_REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(_REPO_ROOT))

from scripts.extract_references import detect_references_section
from scripts.lib.citation_models import CitationRecord, ValidationResult, citations_from_raw


# -------------------------------------------------------------------
# Source registry
# -------------------------------------------------------------------

def _load_sources() -> dict[str, Any]:
    """Load available source modules."""
    from scripts.sources import source_openalex, source_crossref, source_semantic_scholar, source_core
    from scripts.sources import source_cnki, source_wanfang

    return {
        "openalex": source_openalex,
        "crossref": source_crossref,
        "semantic_scholar": source_semantic_scholar,
        "core": source_core,
        "cnki": source_cnki,
        "wanfang": source_wanfang,
    }


def list_sources() -> dict[str, dict]:
    """List all sources and their availability."""
    sources = _load_sources()
    info = {}
    for name, mod in sources.items():
        info[name] = {
            "available": mod.is_available(),
            "description": (mod.__doc__ or "").strip().split("\n")[0],
        }
    return info


# -------------------------------------------------------------------
# Matching
# -------------------------------------------------------------------

def _normalize_title(title: str) -> str:
    """Normalize a title for fuzzy matching."""
    import re
    if not title:
        return ""
    title = re.sub(r'[{}\\]', '', title)
    title = re.sub(r'[^\w\s]', '', title, flags=re.UNICODE)
    title = re.sub(r'\s+', ' ', title).strip().lower()
    return title


def _title_similarity(t1: str, t2: str) -> float:
    """Compute similarity between two normalized titles."""
    if not t1 or not t2:
        return 0.0
    return SequenceMatcher(None, t1, t2).ratio()


def _build_search_query(cit: CitationRecord) -> str:
    """Build a search query from citation fields."""
    parts = []
    if cit.authors:
        parts.append(cit.authors[0])
    if cit.title:
        # Use first 60 chars of title for search
        parts.append(cit.title[:60])
    return " ".join(parts) if parts else cit.raw_text[:80]


def _verify_citation_against_source(
    cit: CitationRecord,
    source_name: str,
    source_mod: Any,
    similarity_threshold: float = 0.75,
) -> dict[str, Any]:
    """Query a single source for a single citation."""
    query = _build_search_query(cit)
    results = source_mod.search(query, limit=5)

    if not results:
        return {"match": False, "source": source_name}

    if isinstance(results, list) and results and "error" in results[0]:
        return {"match": False, "source": source_name, "error": results[0]["error"]}

    cit_title_norm = _normalize_title(cit.title)
    best_match = None
    best_score = 0.0

    for r in results:
        r_title = r.get("title", "")
        score = _title_similarity(cit_title_norm, _normalize_title(r_title))
        if score > best_score:
            best_score = score
            best_match = r

    if best_match and best_score >= similarity_threshold:
        # Verify additional fields
        verified: dict[str, str] = {}
        issues: list[str] = []

        # Year check
        r_year = best_match.get("year", 0)
        if cit.year and r_year:
            if cit.year == r_year:
                verified["year"] = "match"
            elif abs(cit.year - r_year) <= 1:
                verified["year"] = f"close: doc={cit.year}, api={r_year}"
            else:
                verified["year"] = f"mismatch: doc={cit.year}, api={r_year}"
                issues.append(f"年份不一致：论文={cit.year}，API={r_year}")

        # Author check
        if cit.authors and best_match.get("authors"):
            cit_first = cit.authors[0].lower()
            api_first = best_match["authors"][0].lower() if best_match["authors"] else ""
            if cit_first in api_first or api_first in cit_first:
                verified["author"] = "match"
            else:
                verified["author"] = f"possible_mismatch"

        return {
            "match": True,
            "source": source_name,
            "score": round(best_score, 3),
            "matched_title": best_match.get("title", ""),
            "matched_doi": best_match.get("doi", ""),
            "verified_fields": verified,
            "issues": issues,
        }

    return {
        "match": False,
        "source": source_name,
        "best_score": round(best_score, 3) if best_match else 0,
    }


# -------------------------------------------------------------------
# Main verification
# -------------------------------------------------------------------

def verify_citations_external(
    citations: list[CitationRecord],
    sources: list[str] | None = None,
    threshold: float = 0.75,
    max_workers: int = 4,
) -> list[ValidationResult]:
    """Verify citations against external academic databases.

    Args:
        citations: List of CitationRecord from thesis
        sources: Source names to query (None = all available free sources)
        threshold: Title similarity threshold
        max_workers: Max parallel workers per source

    Returns:
        List of ValidationResult
    """
    all_sources = _load_sources()

    if sources:
        active = {k: v for k, v in all_sources.items() if k in sources and v.is_available()}
    else:
        # Default: free sources only (no config needed)
        active = {k: v for k, v in all_sources.items()
                  if k in ("openalex", "crossref", "semantic_scholar") and v.is_available()}

    if not active:
        return [ValidationResult(
            citation_index=c.index,
            raw_text=c.raw_text,
            issues=["没有可用的外部数据源"]
        ) for c in citations]

    source_names = list(active.keys())
    print(f"[verify_external] Using sources: {', '.join(source_names)}")

    results: list[ValidationResult] = []

    for cit in citations:
        vr = ValidationResult(
            citation_index=cit.index,
            raw_text=cit.raw_text,
        )

        if not cit.title:
            vr.issues.append("无法从引用中提取标题，跳过外部验证")
            results.append(vr)
            continue

        # Query each source (sequentially per citation, across sources in parallel)
        with ThreadPoolExecutor(max_workers=len(active)) as pool:
            futures = {
                pool.submit(_verify_citation_against_source, cit, name, mod, threshold): name
                for name, mod in active.items()
            }

            for future in as_completed(futures):
                source_name = futures[future]
                try:
                    result = future.result()
                    if result.get("match"):
                        vr.external_match = True
                        vr.external_source = source_name
                        vr.external_match_title = result.get("matched_title", "")
                        if result.get("matched_doi"):
                            vr.verified_fields["doi_from_api"] = result["matched_doi"]
                        if result.get("verified_fields"):
                            vr.verified_fields.update(result["verified_fields"])
                        if result.get("issues"):
                            vr.issues.extend(result["issues"])
                        break  # First match wins
                except Exception as e:
                    vr.issues.append(f"[{source_name}] 查询出错: {e}")

        if not vr.external_match:
            vr.issues.append(f"外部数据源均未找到匹配")

        results.append(vr)

    return results


# -------------------------------------------------------------------
# CLI
# -------------------------------------------------------------------

def main():
    from scripts.lib import setup_utf8_output
    setup_utf8_output()
    parser = argparse.ArgumentParser(description="External citation verification")
    parser.add_argument("thesis", nargs="?", default="", help="Path to thesis .docx")
    parser.add_argument("--sources", help="Comma-separated source names (default: openalex,crossref,semantic_scholar)")
    parser.add_argument("--output", "-o", help="Output JSON path")
    parser.add_argument("--threshold", type=float, default=0.75, help="Similarity threshold")
    parser.add_argument("--list-sources", action="store_true", help="List available sources")
    args = parser.parse_args()

    if args.list_sources:
        info = list_sources()
        print("\nAvailable sources:")
        for name, detail in info.items():
            status = "OK" if detail["available"] else "N/A"
            print(f"  [{status}] {name}: {detail['description']}")
        return

    # Extract citations
    ref_result = detect_references_section(args.thesis)
    if not ref_result.found:
        print("No references section found in thesis.")
        return

    citations = citations_from_raw(ref_result.raw_citations)
    print(f"[verify_external] Found {len(citations)} citations")

    sources_list = [s.strip() for s in args.sources.split(",")] if args.sources else None

    results = verify_citations_external(citations, sources_list, args.threshold)

    matched = sum(1 for r in results if r.external_match)
    print(f"[verify_external] Matched: {matched}/{len(results)}")

    output = {
        "source_file": args.thesis,
        "sources_used": [s.strip() for s in args.sources.split(",")] if args.sources else ["openalex", "crossref", "semantic_scholar"],
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
