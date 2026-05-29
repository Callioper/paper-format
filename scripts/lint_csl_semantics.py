#!/usr/bin/env python3
"""Lint a CSL file for common semantic mistakes seen in Chinese journal styles."""

from __future__ import annotations

import argparse
import sys
import xml.etree.ElementTree as ET
from pathlib import Path


CSL_NAMESPACE = "http://purl.org/net/xbiblio/csl"
NS = {"csl": CSL_NAMESPACE}
STYLE_TAG = f"{{{CSL_NAMESPACE}}}style"


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Lint CSL semantics.")
    parser.add_argument("csl_path", help="Path to the .csl file")
    return parser.parse_args()


def text_content(element: ET.Element | None) -> str:
    if element is None:
        return ""
    return "".join(element.itertext())


def has_descendant_text(element: ET.Element | None, needles: list[str]) -> bool:
    haystack = text_content(element)
    return any(needle in haystack for needle in needles)


def find_macro(root: ET.Element, name: str) -> ET.Element | None:
    return root.find(f"./csl:macro[@name='{name}']", NS)


def children_with_type(macro: ET.Element | None, item_type: str) -> list[ET.Element]:
    if macro is None:
        return []
    matches: list[ET.Element] = []
    for element in macro.iter():
        if element.tag.endswith("if") or element.tag.endswith("else-if"):
            if element.get("type", "") and item_type in element.get("type", "").split():
                matches.append(element)
    return matches


def main() -> int:
    args = parse_args()
    csl_path = Path(args.csl_path).expanduser().resolve()

    if not csl_path.is_file():
        print(f"Missing CSL file: {csl_path}", file=sys.stderr)
        return 1

    try:
        tree = ET.parse(csl_path)
    except ET.ParseError as exc:
        print(f"XML parse error: {exc}", file=sys.stderr)
        return 1

    root = tree.getroot()
    if root.tag != STYLE_TAG:
        print(f"Invalid root element: {root.tag}", file=sys.stderr)
        return 1

    issues: list[str] = []
    warnings: list[str] = []

    default_locale = root.get("default-locale", "")
    if default_locale.startswith("zh"):
        english_entry = find_macro(root, "entry-en-note")
        if english_entry is not None and english_entry.get("locale") != "en":
            citation_else = root.find("./csl:citation/csl:layout/csl:choose/csl:else/csl:text[@macro='entry-en-note']", NS)
            bibliography_else = root.find("./csl:bibliography/csl:layout/csl:choose/csl:else/csl:text[@macro='entry-en-note']", NS)
            if citation_else is None or citation_else.get("locale") != "en":
                issues.append("default-locale 是中文，但 citation 中英文分支没有显式 locale='en'。")
            if bibliography_else is None or bibliography_else.get("locale") != "en":
                issues.append("default-locale 是中文，但 bibliography 中英文分支没有显式 locale='en'。")

    access_zh = find_macro(root, "access-zh")
    access_en = find_macro(root, "access-en")
    if has_descendant_text(access_zh, ["URL"]) and not children_with_type(access_zh, "webpage"):
        warnings.append("access-zh 包含 URL 输出，请确认只对电子文献类型生效。")
    if has_descendant_text(access_en, ["URL"]) and not children_with_type(access_en, "webpage"):
        warnings.append("access-en 包含 URL 输出，请确认只对电子文献类型生效。")

    zh_journal = find_macro(root, "entry-zh-note")
    for journal_branch in children_with_type(zh_journal, "article-journal"):
        if has_descendant_text(journal_branch, ["vol.", "no.", "pp.", "p."]):
            issues.append("中文期刊分支中出现英文卷期或页码标签。")
        if has_descendant_text(journal_branch, ["URL", "accessed", "访问"]):
            issues.append("中文期刊分支中出现链接或访问日期输出。")

    zh_newspaper = find_macro(root, "entry-zh-note")
    for news_branch in children_with_type(zh_newspaper, "article-newspaper"):
        if has_descendant_text(news_branch, ["URL", "accessed", "访问"]):
            issues.append("中文报纸分支中出现链接或访问日期输出。")

    en_entry = find_macro(root, "entry-en-note")
    if has_descendant_text(en_entry, ["年", "月", "日", "页", "第"]):
        issues.append("英文分支中混入了中文日期或页码用语。")

    en_periodical = find_macro(root, "container-periodical-en")
    if has_descendant_text(en_periodical, ["年", "月", "日", "页", "第"]):
        issues.append("英文期刊容器分支中混入了中文本地化文本。")

    locators_en = find_macro(root, "locators-en")
    if has_descendant_text(locators_en, ["页", "第"]):
        issues.append("英文页码分支中混入中文页码标签。")

    locator_only_zh = find_macro(root, "locator-only-zh")
    if locator_only_zh is None:
        warnings.append("没有 locator-only-zh 这类仅输出引文页码的宏，请确认中文期刊脚注不会误带条目总页码。")

    if issues:
        print("CSL semantic lint failed:")
        for issue in issues:
            print(f"- {issue}")
        if warnings:
            print("Warnings:")
            for warning in warnings:
                print(f"- {warning}")
        return 1

    print(f"CSL semantic lint passed: {csl_path}")
    if warnings:
      print("Warnings:")
      for warning in warnings:
          print(f"- {warning}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
