"""CORE data source - 300M+ open access papers.

Free API with optional CORE_API_KEY for higher limits.
https://core.ac.uk/docs/api
"""

from __future__ import annotations

import json
import os
import urllib.request
import urllib.parse
from typing import Any

API_BASE = "https://api.core.ac.uk/v3/search/works"


def is_available() -> bool:
    """CORE is always available (no key required for basic access)."""
    return True


def search(query: str, limit: int = 10) -> list[dict[str, Any]]:
    """Search CORE for matching works.

    Returns list of normalized citation dicts.
    """
    params = urllib.parse.urlencode({
        "q": query,
        "limit": min(limit, 25),
    })
    url = f"{API_BASE}?{params}"

    headers = {"Accept": "application/json"}
    api_key = os.environ.get("CORE_API_KEY", "")
    if api_key:
        headers["Authorization"] = f"Bearer {api_key}"

    try:
        req = urllib.request.Request(url, headers=headers)
        with urllib.request.urlopen(req, timeout=15) as resp:
            data = json.loads(resp.read())
    except Exception as e:
        return [{"error": str(e), "source": "core"}]

    results: list[dict[str, Any]] = []
    for work in data.get("results", []):
        authors = [a.get("name", "") for a in (work.get("authors") or []) if a.get("name")]

        results.append({
            "title": (work.get("title") or "").strip(),
            "authors": authors[:5],
            "year": work.get("yearPublished", 0) or 0,
            "venue": (work.get("publisher") or "").strip(),
            "doi": work.get("doi", "") or "",
            "source": "core",
            "url": work.get("downloadUrl", "") or "",
        })

    return results
