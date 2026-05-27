"""OpenAlex data source - free academic search API (250M+ papers).

No API key required. Uses polite pool with email for better rate limits.
https://docs.openalex.org/
"""

from __future__ import annotations

import json
import os
import urllib.request
import urllib.parse
from typing import Any

API_BASE = "https://api.openalex.org/works"
CONTACT_EMAIL = os.environ.get("CONTACT_EMAIL", "paper-check@example.com")


def is_available() -> bool:
    """OpenAlex is always available (no key required)."""
    return True


def search(query: str, limit: int = 10) -> list[dict[str, Any]]:
    """Search OpenAlex for matching works.

    Returns list of normalized citation dicts.
    """
    params = urllib.parse.urlencode({
        "search": query,
        "per_page": min(limit, 25),
        "mailto": CONTACT_EMAIL,
    })
    url = f"{API_BASE}?{params}"

    try:
        req = urllib.request.Request(url, headers={"Accept": "application/json"})
        with urllib.request.urlopen(req, timeout=15) as resp:
            data = json.loads(resp.read())
    except Exception as e:
        return [{"error": str(e), "source": "openalex"}]

    results: list[dict[str, Any]] = []
    for work in data.get("results", []):
        # Extract authors
        authors = []
        for auth in work.get("authorships", []):
            name = auth.get("author", {}).get("display_name", "")
            if name:
                authors.append(name)

        # Extract publication year
        year = work.get("publication_year", 0)

        # Extract venue
        venue = ""
        loc = work.get("primary_location", {})
        if loc and loc.get("source"):
            venue = loc["source"].get("display_name", "")

        results.append({
            "title": (work.get("title") or "").strip(),
            "authors": authors[:5],
            "year": year,
            "venue": venue,
            "doi": (work.get("doi") or "").replace("https://doi.org/", ""),
            "source": "openalex",
            "url": work.get("id", ""),
        })

    return results
