"""CrossRef data source - DOI metadata and citation lookup.

Free REST API. Polite pool requires email in query.
https://api.crossref.org/
"""

from __future__ import annotations

import json
import os
import urllib.request
import urllib.parse
from typing import Any

API_BASE = "https://api.crossref.org/works"
CONTACT_EMAIL = os.environ.get("CONTACT_EMAIL", "paper-check@example.com")


def is_available() -> bool:
    """CrossRef is always available (no key required)."""
    return True


def search(query: str, limit: int = 10) -> list[dict[str, Any]]:
    """Search CrossRef for matching works.

    Returns list of normalized citation dicts.
    """
    params = urllib.parse.urlencode({
        "query": query,
        "rows": min(limit, 25),
        "mailto": CONTACT_EMAIL,
    })
    url = f"{API_BASE}?{params}"

    try:
        req = urllib.request.Request(url, headers={"Accept": "application/json"})
        with urllib.request.urlopen(req, timeout=15) as resp:
            data = json.loads(resp.read())
    except Exception as e:
        return [{"error": str(e), "source": "crossref"}]

    results: list[dict[str, Any]] = []
    for item in data.get("message", {}).get("items", []):
        # Authors
        authors = []
        for auth in item.get("author", []):
            parts = []
            if auth.get("given"):
                parts.append(auth["given"])
            if auth.get("family"):
                parts.append(auth["family"])
            if parts:
                authors.append(" ".join(parts))

        # Year from published-print or published-online
        year = 0
        for date_field in ("published-print", "published-online", "issued"):
            dp = item.get(date_field, {}).get("date-parts", [[]])
            if dp and dp[0] and dp[0][0]:
                year = dp[0][0]
                break

        # Venue
        venue = ""
        if item.get("container-title"):
            venue = item["container-title"][0] if isinstance(item["container-title"], list) else item["container-title"]

        # DOI
        doi = item.get("DOI", "")

        # Title
        title = ""
        if item.get("title"):
            title = item["title"][0] if isinstance(item["title"], list) else item["title"]

        results.append({
            "title": title.strip(),
            "authors": authors[:5],
            "year": year,
            "venue": venue,
            "doi": doi,
            "source": "crossref",
            "url": f"https://doi.org/{doi}" if doi else "",
        })

    return results
