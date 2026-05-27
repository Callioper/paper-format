"""知网 CNKI 数据源 - 需要 Cookie 配置（实验性）

使用方式：
1. 在浏览器中登录知网 (https://kns.cnki.net)
2. 从浏览器开发者工具中复制 Cookie
3. 设置环境变量 CNKI_COOKIE 或写入 scripts/.env 文件

注意：知网反爬机制较强，此模块为实验性功能。
"""

from __future__ import annotations

import json
import os
import re
import sys
import urllib.request
import urllib.parse
from typing import Any

SEARCH_URL = "https://kns.cnki.net/kns8s/brief/grid"

_HEADERS_BASE = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                  "(KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Language": "zh-CN,zh;q=0.9",
    "Referer": "https://kns.cnki.net/",
}


def _env_get(key: str) -> str:
    """Get env var, checking .env file as fallback."""
    val = os.environ.get(key, "")
    if val:
        return val
    env_path = Path(__file__).resolve().parent / ".env"
    if env_path.exists():
        for line in env_path.read_text(encoding="utf-8").splitlines():
            line = line.strip()
            if line.startswith(f"{key}="):
                return line[len(key) + 1:].strip().strip('"').strip("'")
    return ""


from pathlib import Path


def is_available() -> bool:
    """Check if CNKI Cookie is configured."""
    return bool(_env_get("CNKI_COOKIE"))


def search(query: str, limit: int = 10) -> list[dict[str, Any]]:
    """Search CNKI for matching papers.

    Requires CNKI_COOKIE environment variable.
    Returns list of normalized citation dicts.
    """
    cookie = _env_get("CNKI_COOKIE")
    if not cookie:
        return [{"error": "未配置 CNKI_COOKIE，请在 .env 文件中设置知网 Cookie", "source": "cnki"}]

    headers = {**_HEADERS_BASE, "Cookie": cookie}
    params = {
        "QueryJson": json.dumps({
            "Platform": "",
            "DBCode": "CFLS",
            "KuaKuCode": "CJFQ,CDMD,CIPD,CCND,CISD,SNAD,BDZK,CCJD,CCVD,CJFN",
            "QNode": {"QGroup": [{"Key": "Subject", "Title": "", "Logic": 1,
                                   "Items": [{"Title": "主题", "Name": "SU", "Value": query,
                                               "Operate": "=", "BlurType": ""}],
                                   "ChildItems": []}]},
        }, ensure_ascii=False),
        "SearchSql": query,
        "PageName": "DefaultResult",
        "DBCode": "CFLS",
        "KuaKuCodes": "CJFQ,CDMD,CIPD,CCND,CISD,SNAD,BDZK,CCJD,CCVD,CJFN",
        "CurPage": "1",
        "RecordsCntPerPage": str(min(limit, 20)),
    }

    try:
        encoded = urllib.parse.urlencode(params)
        req = urllib.request.Request(
            f"{SEARCH_URL}?{encoded}", headers=headers, method="GET"
        )
        with urllib.request.urlopen(req, timeout=15) as resp:
            html = resp.read().decode("utf-8", errors="replace")
    except Exception as e:
        return [{"error": str(e), "source": "cnki"}]

    return _parse_html(html)


def _parse_html(html: str) -> list[dict[str, Any]]:
    """Parse CNKI search results HTML."""
    results: list[dict[str, Any]] = []

    # Try to extract from table rows
    # Pattern: <td class="name"><a ...>Title</a></td>
    rows = re.findall(r'<tr[^>]*>(.*?)</tr>', html, re.DOTALL)
    for row in rows:
        # Title
        title_m = re.search(r'<td[^>]*class="name"[^>]*>.*?<a[^>]*>(.*?)</a>', row, re.DOTALL)
        if not title_m:
            continue
        title = re.sub(r'<[^>]+>', '', title_m.group(1)).strip()
        if not title:
            continue

        # Author
        author_m = re.search(r'<td[^>]*class="author"[^>]*>(.*?)</td>', row, re.DOTALL)
        authors = []
        if author_m:
            author_tags = re.findall(r'<a[^>]*>(.*?)</a>', author_m.group(1))
            authors = [a.strip() for a in author_tags if a.strip()]

        # Source/Venue
        venue = ""
        source_m = re.search(r'<td[^>]*class="source"[^>]*>.*?<a[^>]*>(.*?)</a>', row, re.DOTALL)
        if source_m:
            venue = re.sub(r'<[^>]+>', '', source_m.group(1)).strip()

        # Year
        year = 0
        date_m = re.search(r'<td[^>]*class="date"[^>]*>(.*?)</td>', row, re.DOTALL)
        if date_m:
            year_text = re.sub(r'<[^>]+>', '', date_m.group(1)).strip()
            year_match = re.search(r'((?:19|20)\d{2})', year_text)
            if year_match:
                year = int(year_match.group(1))

        results.append({
            "title": title,
            "authors": authors[:5],
            "year": year,
            "venue": venue,
            "doi": "",
            "source": "cnki",
            "url": "",
        })

    return results
