#!/usr/bin/env python3
"""cross_check.py - Cross-reference validation between in-text citations and reference list.

Checks:
  1. Forward: in-text [1][2][3] have corresponding reference list entries
  2. Backward: reference list entries are all cited in the text
  3. Duplicate: multiple citations pointing to the same literature
  4. Completeness: required fields (author/title/year) are present

Usage:
    python cross_check.py thesis.docx --output cross_check_result.json
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
from scripts.lib.citation_models import CrossCheckResult, citations_from_raw


def _normalize_title(title: str) -> str:
    """Normalize a title for fuzzy matching."""
    if not title:
        return ""
    title = re.sub(r'[{}\\]', '', title)
    title = re.sub(r'[^\w\s]', '', title, flags=re.UNICODE)
    title = re.sub(r'\s+', ' ', title).strip().lower()
    return title


def extract_in_text_citations(docx_path: str | Path) -> dict[int, str]:
    """Extract in-text citation references like [1], [2], [3] from body paragraphs.

    Returns dict mapping citation_index -> surrounding text context.
    """
    from docx import Document

    doc = Document(docx_path)
    citations: dict[int, str] = {}

    for para in doc.paragraphs:
        text = para.text.strip()
        if not text:
            continue
        # Skip reference list section
        if re.match(r'^参考文献|References|Bibliography', text):
            break

        # Find [N] patterns
        for m in re.finditer(r'\[(\d+)\]', text):
            idx = int(m.group(1))
            # Get surrounding context
            start = max(0, m.start() - 30)
            end = min(len(text), m.end() + 30)
            context = text[start:end]
            if idx not in citations:
                citations[idx] = context

    return citations


def cross_check(docx_path: str | Path) -> CrossCheckResult:
    """Perform cross-reference validation.

    Returns CrossCheckResult with orphan citations, unused references, etc.
    """
    result = CrossCheckResult()

    # 1. Extract in-text citations
    in_text = extract_in_text_citations(docx_path)
    if not in_text:
        result.orphan_citations.append({
            "issue": "正文中未检测到引用标记 [N]",
        })
        return result

    # 2. Extract reference list
    ref_result = detect_references_section(docx_path)
    if not ref_result.found:
        result.unused_references.append({
            "issue": "未检测到参考文献章节",
        })
        # All in-text citations are orphans
        for idx, ctx in in_text.items():
            result.orphan_citations.append({
                "index": idx,
                "context": ctx[:80],
                "issue": f"正文引用[{idx}]但参考文献中无此条目",
            })
        return result

    # Build reference list index
    ref_indices: set[int] = set()
    ref_texts: dict[int, str] = {}
    for raw in ref_result.raw_citations:
        m = re.match(r'^\[?(\d+)\]?\s*', raw)
        if m:
            idx = int(m.group(1))
            ref_indices.add(idx)
            ref_texts[idx] = raw

    # 3. Forward check: in-text citations have matching reference entries
    for idx, ctx in in_text.items():
        if idx not in ref_indices:
            result.orphan_citations.append({
                "index": idx,
                "context": ctx[:80],
                "issue": f"正文引用[{idx}]但参考文献中无此条目",
            })

    # 4. Backward check: reference entries are all cited
    cited_indices = set(in_text.keys())
    for idx in ref_indices:
        if idx not in cited_indices:
            ref_text = ref_texts.get(idx, "")[:80]
            result.unused_references.append({
                "index": idx,
                "text": ref_text,
                "issue": f"参考文献[{idx}]未被正文引用",
            })

    # 5. Duplicate check: similar titles in reference list
    citations = citations_from_raw(ref_result.raw_citations)
    seen_titles: dict[str, list[int]] = {}
    for cit in citations:
        norm = _normalize_title(cit.title)
        if not norm or len(norm) < 10:
            continue
        prefix = norm[:30]
        if prefix in seen_titles:
            seen_titles[prefix].append(cit.index)
        else:
            seen_titles[prefix] = [cit.index]

    for prefix, indices in seen_titles.items():
        if len(indices) > 1:
            # Verify they're really duplicates (full title comparison)
            titles = []
            for cit in citations:
                if cit.index in indices:
                    titles.append((cit.index, _normalize_title(cit.title)))

            if len(titles) >= 2:
                sim = SequenceMatcher(None, titles[0][1], titles[1][1]).ratio()
                if sim > 0.85:
                    result.duplicates.append({
                        "indices": [t[0] for t in titles],
                        "similarity": round(sim, 3),
                        "text": titles[0][1][:60],
                        "issue": f"参考文献{[t[0] for t in titles]}可能指向同一文献 (相似度={sim:.1%})",
                    })

    # 6. Completeness check
    for cit in citations:
        missing: list[str] = []
        if not cit.authors or (len(cit.authors) == 1 and not cit.authors[0]):
            missing.append("author")
        if not cit.title:
            missing.append("title")
        if not cit.year:
            missing.append("year")
        if missing:
            result.incomplete.append({
                "index": cit.index,
                "missing": missing,
                "text": cit.raw_text[:80],
                "issue": f"参考文献[{cit.index}]缺少必填字段: {', '.join(missing)}",
            })

    return result


# -------------------------------------------------------------------
# CLI
# -------------------------------------------------------------------

def main():
    from scripts.lib import setup_utf8_output
    setup_utf8_output()
    parser = argparse.ArgumentParser(description="Cross-reference validation")
    parser.add_argument("thesis", help="Path to thesis .docx")
    parser.add_argument("--output", "-o", help="Output JSON path")
    args = parser.parse_args()

    result = cross_check(args.thesis)
    output = result.to_dict()

    # Summary
    total_issues = (
        len(result.orphan_citations) +
        len(result.unused_references) +
        len(result.duplicates) +
        len(result.incomplete)
    )
    print(f"[cross_check] Issues found: {total_issues}")
    print(f"  - 孤立引用: {len(result.orphan_citations)}")
    print(f"  - 未使用文献: {len(result.unused_references)}")
    print(f"  - 重复引用: {len(result.duplicates)}")
    print(f"  - 不完整条目: {len(result.incomplete)}")

    if args.output:
        Path(args.output).write_text(json.dumps(output, ensure_ascii=False, indent=2), encoding="utf-8")
        print(f"Results written to {args.output}")
    else:
        print(json.dumps(output, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
