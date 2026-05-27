"""Semantic Scholar data source - AI-powered academic search.

Free public API. Optional S2_API_KEY for higher rate limits.
https://api.semanticscholar.org/
"""

from __future__ import annotations

import json
import os
import urllib.request
import urllib.parse
from typing import Any

API_BASE = "https://api.semanticscholar.org/graph/v1/paper/search"


def is_available() -> bool:
    """Semantic Scholar is always available (no key required)."""
    return True


def search(query: str, limit: int = 10) -> list[dict[str, Any]]:
    """Search Semantic Scholar for matching papers.

    Returns list of normalized citation dicts.
    """
    params = urllib.parse.urlencode({
        "query": query,
        "limit": min(limit, 25),
        "fields": "title,authors,year,venue,externalIds,publicationTypes",
    })
    url = f"{API_BASE}?{params}"

    headers = {"Accept": "application/json"}
    api_key = os.environ.get("S2_API_KEY", "")
    if api_key:
        headers["x-api-key"] = api_key

    try:
        req = urllib.request.Request(url, headers=headers)
        with urllib.request.urlopen(req, timeout=15) as resp:
            data = json.loads(resp.read())
    except Exception as e:
        return [{"error": str(e), "source": "semantic_scholar"}]

    results: list[dict[str, Any]] = []
    for paper in data.get("data", []):
        authors = [a.get("name", "") for a in paper.get("authors", []) if a.get("name")]

        ext_ids = paper.get("externalIds", {}) or {}
        doi = ext_ids.get("DOI", "")

        results.append({
            "title": (paper.get("title") or "").strip(),
            "authors": authors[:5],
            "year": paper.get("year", 0) or 0,
            "venue": paper.get("venue", "") or "",
            "doi": doi,
            "source": "semantic_scholar",
            "url": f"https://doi.org/{doi}" if doi else "",
        })

    return results
