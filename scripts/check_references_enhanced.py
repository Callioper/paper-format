#!/usr/bin/env python3
"""check_references_enhanced.py — 增强型引用完整性检查。

在现有引文验证基础上新增：
  - DOI 格式校验（正则匹配）
  - BibTeX 必填字段检查
  - URL 可达性验证（HEAD 请求）
  - 缺失 DOI 候选推荐

Usage:
    python scripts/check_references_enhanced.py "论文.docx" --bib refs.bib --output ref_enhanced.json
    python scripts/check_references_enhanced.py "论文.docx" --output ref_enhanced.json
"""

from __future__ import annotations

import argparse
import json
import re
import sys
import urllib.error
import urllib.request
from pathlib import Path
from typing import Any

try:
    from docx import Document
except ImportError:
    print("需要 python-docx: pip install python-docx", file=sys.stderr)
    sys.exit(1)


# ── DOI 正则 ─────────────────────────────────────────────

_DOI_PATTERN = re.compile(
    r"(?:doi[:\s]*|https?://doi\.org/|https?://dx\.doi\.org/)"
    r"(10\.\d{4,9}/[^\s,;\"'<>]+)",
    re.IGNORECASE,
)
_DOI_STANDALONE = re.compile(r"\b(10\.\d{4,9}/[^\s,;\"'<>]+)\b")

# ── BibTeX 必填字段 ──────────────────────────────────────

_REQUIRED_FIELDS = {
    "article": ["author", "title", "journal", "year"],
    "book": ["author", "title", "publisher", "year"],
    "inproceedings": ["author", "title", "booktitle", "year"],
    "incollection": ["author", "title", "booktitle", "year"],
    "phdthesis": ["author", "title", "school", "year"],
    "mastersthesis": ["author", "title", "school", "year"],
    "techreport": ["author", "title", "institution", "year"],
    "misc": ["author", "title"],
    "online": ["author", "title", "url"],
    "inbook": ["author", "title", "publisher", "year"],
}

_RECOMMENDED_FIELDS = {
    "article": ["volume", "pages", "doi"],
    "book": ["address", "isbn"],
    "inproceedings": ["pages", "address"],
    "phdthesis": ["address"],
}


# ── .bib 解析（轻量级） ──────────────────────────────────

def _parse_bib_entries(bib_path: Path) -> list[dict[str, Any]]:
    """解析 .bib 文件，提取条目类型、key 和字段。"""
    if not bib_path.exists():
        return []
    text = bib_path.read_text(encoding="utf-8", errors="ignore")
    entries = []
    # 匹配 @type{key, ... }
    entry_pattern = re.compile(
        r"@(\w+)\s*\{([^,]+),\s*(.*?)\n\}",
        re.DOTALL | re.IGNORECASE,
    )
    for m in entry_pattern.finditer(text):
        entry_type = m.group(1).lower()
        entry_key = m.group(2).strip()
        body = m.group(3)

        fields = {}
        # 提取字段 key = {value} 或 key = value
        field_pattern = re.compile(
            r"(\w+)\s*=\s*[\{](.*?)[\}]",
            re.DOTALL,
        )
        for fm in field_pattern.finditer(body):
            fields[fm.group(1).lower()] = fm.group(2).strip()

        entries.append({
            "key": entry_key,
            "type": entry_type,
            "fields": fields,
        })
    return entries


# ── DOI 校验 ─────────────────────────────────────────────

def _validate_doi(doi: str) -> dict[str, Any]:
    """校验 DOI 格式。"""
    doi = doi.strip().lower()
    # 去掉 URL 前缀
    doi_clean = re.sub(r"^https?://(dx\.)?doi\.org/", "", doi)

    if not re.match(r"^10\.\d{4,9}/", doi_clean):
        return {"valid": False, "doi": doi, "reason": "DOI 格式不符合 10.xxxx/... 规范"}

    # 检查长度
    if len(doi_clean) > 256:
        return {"valid": False, "doi": doi, "reason": "DOI 过长，可能包含多余内容"}

    return {"valid": True, "doi": doi_clean}


# ── URL 可达性 ────────────────────────────────────────────

def _check_url_reachable(url: str, timeout: int = 5) -> dict[str, Any]:
    """HEAD 请求检查 URL 是否可达。"""
    if not url.startswith(("http://", "https://")):
        url = "https://" + url
    try:
        req = urllib.request.Request(url, method="HEAD")
        req.add_header("User-Agent", "Mozilla/5.0 (compatible; paper-format/1.0)")
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            return {"url": url, "reachable": True, "status": resp.status}
    except urllib.error.HTTPError as e:
        return {"url": url, "reachable": False, "status": e.code, "reason": str(e)}
    except Exception as e:
        return {"url": url, "reachable": False, "reason": str(e)[:100]}


# ── 字段检查 ──────────────────────────────────────────────

def _check_entry_fields(entry: dict[str, Any]) -> dict[str, Any]:
    """检查条目的必填字段和推荐字段。"""
    entry_type = entry["type"]
    fields = entry["fields"]

    required = _REQUIRED_FIELDS.get(entry_type, [])
    recommended = _RECOMMENDED_FIELDS.get(entry_type, [])

    missing_required = [f for f in required if f not in fields]
    missing_recommended = [f for f in recommended if f not in fields]

    return {
        "key": entry["key"],
        "type": entry_type,
        "missing_required": missing_required,
        "missing_recommended": missing_recommended,
        "has_doi": "doi" in fields,
        "has_url": "url" in fields,
    }


# ── 主检查流程 ──────────────────────────────────────────

def check_references_enhanced(
    docx_path: str | Path,
    bib_path: Path | None = None,
    check_urls: bool = False,
) -> dict[str, Any]:
    """执行增强型引用检查。"""
    report: dict[str, Any] = {
        "file": str(docx_path),
        "bib_file": str(bib_path) if bib_path else None,
        "doi_validation": [],
        "field_completeness": [],
        "url_reachability": [],
        "missing_doi_candidates": [],
        "summary": {},
    }

    # 1. 如果有 .bib 文件，检查字段完整性
    entries = []
    if bib_path and bib_path.exists():
        entries = _parse_bib_entries(bib_path)
        report["summary"]["total_entries"] = len(entries)

        for entry in entries:
            # 字段完整性
            field_check = _check_entry_fields(entry)
            report["field_completeness"].append(field_check)

            # DOI 校验
            doi_val = entry["fields"].get("doi", "")
            if doi_val:
                doi_result = _validate_doi(doi_val)
                doi_result["entry_key"] = entry["key"]
                report["doi_validation"].append(doi_result)
            elif entry["type"] in ("article", "inproceedings", "book"):
                # 推荐有 DOI 但缺失
                report["missing_doi_candidates"].append({
                    "key": entry["key"],
                    "type": entry["type"],
                    "title": entry["fields"].get("title", "")[:80],
                    "reason": f"{entry['type']} 类条目建议补充 DOI",
                })

            # URL 可达性（可选，较慢）
            if check_urls:
                url = entry["fields"].get("url", "")
                if url:
                    url_result = _check_url_reachable(url)
                    url_result["entry_key"] = entry["key"]
                    report["url_reachability"].append(url_result)

    # 2. 扫描文档中的引用
    doc = Document(str(docx_path))
    all_text = "\n".join(p.text for p in doc.paragraphs)

    # 提取正文中的 DOI
    text_dois = []
    for m in _DOI_STANDALONE.finditer(all_text):
        doi = m.group(1)
        if doi not in text_dois:
            text_dois.append(doi)

    report["text_dois"] = text_dois

    # 提取正文中的 URL
    url_pattern = re.compile(r"https?://[^\s,;\"'<>]+")
    text_urls = list(set(url_pattern.findall(all_text)))
    report["text_urls"] = text_urls[:20]  # 最多报告 20 个

    # 汇总统计
    missing_required_count = sum(
        1 for f in report["field_completeness"] if f["missing_required"]
    )
    missing_recommended_count = sum(
        1 for f in report["field_completeness"] if f["missing_recommended"]
    )
    invalid_doi_count = sum(
        1 for d in report["doi_validation"] if not d["valid"]
    )
    unreachable_url_count = sum(
        1 for u in report["url_reachability"] if not u["reachable"]
    )

    report["summary"].update({
        "entries_with_missing_required_fields": missing_required_count,
        "entries_with_missing_recommended_fields": missing_recommended_count,
        "invalid_dois": invalid_doi_count,
        "missing_doi_candidates": len(report["missing_doi_candidates"]),
        "unreachable_urls": unreachable_url_count,
        "dois_in_text": len(text_dois),
        "urls_in_text": len(text_urls),
    })

    total_issues = (
        missing_required_count
        + invalid_doi_count
        + unreachable_url_count
    )
    report["total_issues"] = total_issues

    return report


# ── CLI ──────────────────────────────────────────────────

def main() -> int:
    parser = argparse.ArgumentParser(description="增强型引用完整性检查")
    parser.add_argument("docx", help="待检查的 .docx 文件")
    parser.add_argument("--bib", help=".bib 参考文献文件路径")
    parser.add_argument("--check-urls", action="store_true",
                        help="检查 URL 可达性（较慢）")
    parser.add_argument("--output", "-o", help="输出 JSON 路径")
    args = parser.parse_args()

    docx_path = Path(args.docx)
    if not docx_path.exists():
        print(f"文件不存在: {docx_path}", file=sys.stderr)
        return 1

    bib_path = Path(args.bib) if args.bib else None
    if bib_path and not bib_path.exists():
        print(f"bib 文件不存在: {bib_path}", file=sys.stderr)
        return 1

    report = check_references_enhanced(docx_path, bib_path, args.check_urls)

    if args.output:
        Path(args.output).write_text(
            json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8"
        )
        print(f"检查完成: {report['total_issues']} 个问题 → {args.output}")
    else:
        print(json.dumps(report, ensure_ascii=False, indent=2))

    # 打印摘要
    s = report["summary"]
    print(f"\n=== 引用增强检查摘要 ===")
    if "total_entries" in s:
        print(f"  bib 条目总数: {s['total_entries']}")
    print(f"  缺失必填字段的条目: {s.get('entries_with_missing_required_fields', 0)}")
    print(f"  缺失推荐字段的条目: {s.get('entries_with_missing_recommended_fields', 0)}")
    print(f"  无效 DOI: {s.get('invalid_dois', 0)}")
    print(f"  缺失 DOI 候选: {s.get('missing_doi_candidates', 0)}")
    if args.check_urls:
        print(f"  不可达 URL: {s.get('unreachable_urls', 0)}")
    print(f"  正文中的 DOI: {s.get('dois_in_text', 0)}")
    print(f"  正文中的 URL: {s.get('urls_in_text', 0)}")

    return 0


if __name__ == "__main__":
    sys.exit(main())
