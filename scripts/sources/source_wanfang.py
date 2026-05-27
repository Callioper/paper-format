"""万方数据 (Wanfang) 数据源 - 公开搜索页抓取（实验性）

万方提供了部分公开的搜索接口，不需要登录即可获取基本搜索结果。
注意：万方反爬机制可能变化，此模块为实验性功能。
"""

from __future__ import annotations

import json
import re
import urllib.request
import urllib.parse
from typing import Any

SEARCH_URL = "https://s.wanfangdata.com.cn/paper"

_HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                  "(KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Language": "zh-CN,zh;q=0.9",
    "Referer": "https://www.wanfangdata.com.cn/",
}


def is_available() -> bool:
    """万方公开搜索不需要额外配置"""
    return True


def search(query: str, limit: int = 10) -> list[dict[str, Any]]:
    """Search Wanfang data.

    Returns list of normalized citation dicts.
    """
    params = urllib.parse.urlencode({
        "q": query,
        "style": "detail",
        "page": "1",
        "size": str(min(limit, 20)),
    })
    url = f"{SEARCH_URL}?{params}"

    try:
        req = urllib.request.Request(url, headers=_HEADERS)
        with urllib.request.urlopen(req, timeout=15) as resp:
            html = resp.read().decode("utf-8", errors="replace")
    except Exception as e:
        return [{"error": str(e), "source": "wanfang"}]

    return _parse_html(html)


def _parse_html(html: str) -> list[dict[str, Any]]:
    """Parse Wanfang search results HTML."""
    results: list[dict[str, Any]] = []

    # Try multiple selectors for result items
    items = re.findall(
        r'<div[^>]*class="[^"]*(?:normal-list|result-item)[^"]*"[^>]*>(.*?)</div>\s*(?=<div[^>]*class="[^"]*(?:normal-list|result-item))',
        html, re.DOTALL
    )

    if not items:
        # Fallback: look for title links
        title_pattern = re.compile(r'<a[^>]*class="[^"]*title[^"]*"[^>]*href="([^"]*)"[^>]*>(.*?)</a>', re.DOTALL)
        for match in title_pattern.finditer(html):
            url = match.group(1)
            title = re.sub(r'<[^>]+>', '', match.group(2)).strip()
            if not title:
                continue
            if url and not url.startswith("http"):
                url = "https://s.wanfangdata.com.cn" + url
            results.append({
                "title": title,
                "authors": [],
                "year": 0,
                "venue": "",
                "doi": "",
                "source": "wanfang",
                "url": url,
            })
        return results

    for item in items:
        # Title
        title_m = re.search(r'<a[^>]*class="[^"]*title[^"]*"[^>]*>(.*?)</a>', item, re.DOTALL)
        if not title_m:
            title_m = re.search(r'<h3[^>]*>.*?<a[^>]*>(.*?)</a>', item, re.DOTALL)
        if not title_m:
            continue
        title = re.sub(r'<[^>]+>', '', title_m.group(1)).strip()
        if not title:
            continue

        # URL
        url_m = re.search(r'<a[^>]*class="[^"]*title[^"]*"[^>]*href="([^"]*)"', item)
        url = ""
        if url_m:
            url = url_m.group(1)
            if url and not url.startswith("http"):
                url = "https://s.wanfangdata.com.cn" + url

        # Authors
        authors = []
        author_m = re.search(r'<span[^>]*class="[^"]*author[^"]*"[^>]*>(.*?)</span>', item, re.DOTALL)
        if author_m:
            author_tags = re.findall(r'<a[^>]*>(.*?)</a>', author_m.group(1))
            authors = [a.strip() for a in author_tags if a.strip()]

        # Year
        year = 0
        date_m = re.search(r'<span[^>]*class="[^"]*(?:year|date)[^"]*"[^>]*>(.*?)</span>', item, re.DOTALL)
        if date_m:
            year_text = re.sub(r'<[^>]+>', '', date_m.group(1)).strip()
            year_match = re.search(r'((?:19|20)\d{2})', year_text)
            if year_match:
                year = int(year_match.group(1))

        # Venue
        venue = ""
        source_m = re.search(r'<span[^>]*class="[^"]*source[^"]*"[^>]*>.*?<a[^>]*>(.*?)</a>', item, re.DOTALL)
        if source_m:
            venue = re.sub(r'<[^>]+>', '', source_m.group(1)).strip()

        results.append({
            "title": title,
            "authors": authors[:5],
            "year": year,
            "venue": venue,
            "doi": "",
            "source": "wanfang",
            "url": url,
        })

    return results
